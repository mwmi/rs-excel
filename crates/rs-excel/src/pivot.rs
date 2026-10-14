//! 数据透视表与切片器：创建（`worksheet.add_pivot`）、布局、值字段、日期/数值分组、
//! 按值排序、Top-N、页筛选与项筛选、读取结果、切片器批量点选。
//!
//! ```no_run
//! # let sheet = unreachable!();
//! use rs_excel::{AggFunc, DatePart};
//! // 源数据 A1:F100 → 透视表放 H3（可跨表 "Sheet2!H3"）
//! let pt = sheet.add_pivot("A1:F100", "H3", None)?;
//! pt.set_fields(&["地区"], &["产品"], &["日期"])?; // 行/列/页字段
//! pt.add_data("销售额", AggFunc::Sum)?;            // 值字段
//! pt.group_dates("日期", &[DatePart::Years, DatePart::Months])?; // 日期按年+月分组
//! pt.sort_by_value("地区", "销售额", false)?;      // 行字段按值升序
//! pt.show_top("地区", "销售额", 5, true)?;         // 只看前 5
//! pt.filter_items("产品", &["苹果", "香蕉"])?;      // 项筛选
//! println!("{}", pt.data_range()?.read_tsv()?);    // 读取结果
//! # Ok::<(), anyhow::Error>(())
//! ```
//!
//! 字段名均为数据源列名（`fields()` 可查）；筛选值以 `items(field)` 为准。
//! 个别成员报错时改走 `pivot.raw()` 直调 COM，或降级用 `filter_items()` / `set_page()`
//! 达成同样筛选效果。

use anyhow::{bail, Result};

use crate::com::dispatch::IDispatchWrapper;
use crate::com::safearray::list_to_variant;
use crate::com::variant::CellValue;
use crate::com::variant::ComValue;
use crate::error::ExcelError;
use crate::range::Range;

// ── 枚举（屏蔽数值常量，AI 无需查表） ──

/// 聚合函数（透视表值字段与区域分类汇总共用，XlConsolidationFunction）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AggFunc {
    Sum,
    Average,
    Count,
    CountNums,
    Max,
    Min,
    Product,
    StdDev,
    Var,
}

impl AggFunc {
    pub fn code(self) -> i32 {
        match self {
            Self::Sum => -4105,
            Self::Average => -4106,
            Self::Count => -4112,
            Self::CountNums => -4113,
            Self::Max => -4136,
            Self::Min => -4139,
            Self::Product => -4149,
            Self::StdDev => -4155,
            Self::Var => -4164,
        }
    }

    /// 值字段的默认名前缀（与 Excel/WPS 自动命名一致，如「求和项:销售额」）
    fn prefix(self) -> &'static str {
        match self {
            Self::Sum => "求和项",
            Self::Average => "平均值项",
            Self::Count => "计数项",
            Self::CountNums => "数值计数项",
            Self::Max => "最大值项",
            Self::Min => "最小值项",
            Self::Product => "乘积项",
            Self::StdDev => "标准差项",
            Self::Var => "方差项",
        }
    }

    /// 由 XlConsolidationFunction 码还原（读回 Function 实际值核对时用）
    fn from_code(code: i32) -> Option<Self> {
        [
            Self::Sum,
            Self::Average,
            Self::Count,
            Self::CountNums,
            Self::Max,
            Self::Min,
            Self::Product,
            Self::StdDev,
            Self::Var,
        ]
        .into_iter()
        .find(|func| func.code() == code)
    }
}

/// 日期分组周期（`group_dates` 的 Periods 数组语义化下标）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DatePart {
    Seconds = 0,
    Minutes = 1,
    Hours = 2,
    Days = 3,
    Months = 4,
    Quarters = 5,
    Years = 6,
}

impl DatePart {
    /// 由所选周期生成 Group 的 7 元素布尔数组 [秒,分,时,日,月,季,年]。
    fn periods(parts: &[DatePart]) -> [bool; 7] {
        let mut p = [false; 7];
        for part in parts {
            p[*part as usize] = true;
        }
        p
    }
}

// ── 快照信息 ──

/// 透视表字段快照。
#[derive(Debug, Clone)]
pub struct PivotFieldInfo {
    pub name: String,
    /// 数据源唯一名（AutoSort/AutoShow 底层用它，本库已自动解析，AI 无需关心）
    pub source_name: String,
    /// hidden / row / col / page / data
    pub orientation: String,
}

/// 透视字段项快照。
#[derive(Debug, Clone)]
pub struct PivotItemInfo {
    pub name: String,
    pub visible: bool,
}

/// 切片器项快照。
#[derive(Debug, Clone)]
pub struct SlicerItemInfo {
    pub name: String,
    pub selected: bool,
}

/// 解析单元格规格：`"H3"` 相对 ws，`"Sheet2!H3"` 跨表经 wb。返回 Range COM 包装。
pub(crate) fn resolve_cell(
    ws: &IDispatchWrapper,
    wb: &IDispatchWrapper,
    spec: &str,
) -> Result<IDispatchWrapper> {
    let spec = spec.trim();
    match spec.rsplit_once('!') {
        Some((sheet, cell)) => {
            let sheet = sheet.trim_matches('\'').trim_matches('"').trim();
            let cell = cell.trim().trim_start_matches('$');
            wb.call("Worksheets", vec![sheet.into()])?
                .idispatch()
                .map_err(|e| {
                    ExcelError::hint(e, format!("跨表规格 \"{spec}\" 中的工作表不存在：先用 sheets() 列出"))
                })?
                .call("Range", vec![cell.into()])?
                .idispatch_child(spec)
                .map_err(|e| ExcelError::hint(e, "单元格地址格式如 H3"))
        }
        None => ws
            .call("Range", vec![spec.into()])?
            .idispatch_child(spec)
            .map_err(|e| ExcelError::hint(e, "单元格地址格式如 H3，或跨表 Sheet2!H3")),
    }
}

/// 数据透视表（PivotTable COM 对象）：布局、分组、排序、筛选一站式入口。
pub struct PivotTable {
    raw: IDispatchWrapper,
}

impl PivotTable {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    pub fn name(&self) -> Result<String> {
        self.raw.get("Name")?.string()
    }

    /// 刷新（源数据变更后调用）。
    pub fn refresh(&self) -> Result<()> {
        self.raw.call("RefreshTable", vec![]).map(|_| ())
    }

    /// 清空全部字段布局（透视表壳保留）。
    pub fn clear(&self) -> Result<()> {
        self.raw.call("ClearTable", vec![]).map(|_| ())
    }

    /// 绑定字段（PivotFields(name)，名称为数据源列名）。
    fn field(&self, name: &str) -> Result<IDispatchWrapper> {
        self.raw
            .call("PivotFields", vec![name.into()])?
            .idispatch_child(name)
            .map_err(|e| ExcelError::hint(e, format!("字段 \"{name}\" 不存在：用 fields() 查看可用字段名")))
    }

    /// 数据字段唯一名：AutoSort/AutoShow 需 SourceName 而非显示名（如 "求和项:销售额" 的
    /// SourceName 为 "销售额"），自动解析失败时回退原名。
    fn source_name_of(&self, data_field: &str) -> String {
        for name in [data_field.to_string(), format!("求和项:{data_field}")] {
            if let Some(s) = self
                .field(&name)
                .ok()
                .and_then(|pf| pf.get("SourceName").ok())
                .and_then(|v| v.string().ok())
                .filter(|s| !s.is_empty())
            {
                return s;
            }
        }
        data_field.to_string()
    }

    /// 列出全部字段（名称/源名/所在区域）。字段不存在或拼错时先查这里。
    pub fn fields(&self) -> Result<Vec<PivotFieldInfo>> {
        let coll = self.raw.get_obj("PivotFields", "PivotFields")?;
        let count = coll.get("Count")?.int()?;
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let pf = coll.call("Item", vec![i.into()])?.idispatch()?;
            let name = pf.get("Name")?.string()?;
            // 源名与区域只是参考（AutoSort 用源名、AI 看区域）：WPS 对个别字段会直接抛
            // 0x80020009（实报：清单整张失败），读不到就退回默认值，别让一行带崩全表
            let source_name = pf
                .get("SourceName")
                .ok()
                .and_then(|v| v.string().ok())
                .filter(|text| !text.is_empty())
                .unwrap_or_else(|| name.clone());
            let orientation = match pf.get("Orientation").ok().and_then(|v| v.int().ok()) {
                // xlHidden=0 / xlRowField=1 / xlColumnField=2 / xlPageField=3 / xlDataField=4
                Some(1) => "row",
                Some(2) => "col",
                Some(3) => "page",
                Some(4) => "data",
                _ => "hidden",
            }
            .to_string();
            out.push(PivotFieldInfo { name, source_name, orientation });
        }
        Ok(out)
    }

    /// 列出字段的项（名称/可见性）。筛选值以此为准；分组后项为组名。
    pub fn items(&self, field: &str) -> Result<Vec<PivotItemInfo>> {
        let coll = self
            .field(field)?
            .get_obj("PivotItems", "PivotItems")
            .map_err(|e| ExcelError::hint(e, "该字段无项列表：须先放进行/列/页区域"))?;
        let count = coll.get("Count")?.int()?;
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let it = coll.call("Item", vec![i.into()])?.idispatch()?;
            let name = it.get("Name")?.string()?;
            let visible = matches!(it.get("Visible")?.value(), CellValue::Bool(true));
            out.push(PivotItemInfo { name, visible });
        }
        Ok(out)
    }

    /// 一次设置行/列/页字段布局（空切片表示该区域不设置）。
    pub fn set_fields(&self, rows: &[&str], cols: &[&str], pages: &[&str]) -> Result<()> {
        if rows.is_empty() && cols.is_empty() && pages.is_empty() {
            bail!("set_fields：至少指定一个字段");
        }
        let mut named: Vec<(&str, ComValue)> = Vec::new();
        for (list, key) in [(rows, "RowFields"), (cols, "ColumnFields"), (pages, "PageFields")] {
            if list.is_empty() {
                continue;
            }
            let arr = list_to_variant(list.iter().map(|s| (*s).into()).collect())?;
            named.push((key, ComValue::Raw(arr)));
        }
        self.raw
            .call_named("AddFields", vec![], &named)
            .map_err(|e| ExcelError::hint(e, "字段名须为数据源列名：用 fields() 查看"))
            .map(|_| ())
    }

    /// 添加值字段，如 `add_data("销售额", AggFunc::Sum)`。
    ///
    /// **只走 `Orientation = xlDataField` 一条路**：两个引擎都认，值字段按引擎默认口径建出来
    /// （数值列 = 求和、文本列 = 计数）。不用 `AddDataField`、也不依赖 `PivotField.Function` ——
    /// WPS 上这两个都必抛 0x80020009（实报）；好在这条路**天然幂等**（同一字段不会两次进值区域，
    /// 重复调用不会多出一列）。
    ///
    /// `Ok(s)` 里的 `s` 是**口径说明**（空串 = 就是 `func` 要的口径）：引擎不让改口径时值字段
    /// 仍然建好了，只是口径是引擎默认的 —— 这时**不能报错**（报错会让调用方重发整条建表指令、
    /// 撞上「at 被占/重名」更难懂），把这句话原样带回给模型，让它改用公式或在回答里说明。
    pub fn add_data(&self, field: &str, func: AggFunc) -> Result<String> {
        let pf = self.field(field)?;
        pf.put("Orientation", vec![(4i32).into()]) // xlDataField
            .map_err(|e| ExcelError::hint(e, "进不了值区域：值字段必须是数据源里的列"))?;
        // 挪动后旧对象可能作废：按源名重取，取不到再试引擎给的默认显示名（求和项/计数项）
        let data_field = self
            .field(field)
            .or_else(|_| self.field(&format!("求和项:{field}")))
            .or_else(|_| self.field(&format!("计数项:{field}")))?;
        // 求和是数值列的默认口径，不必碰 Function（WPS 那条 put 必抛）；
        // 要别的口径：改一次，改不动就刷新一次再改（值字段刚建出来可能还没生效）
        if func != AggFunc::Sum {
            if data_field.put("Function", vec![func.code().into()]).is_err() {
                let _ = self.refresh();
                let _ = data_field.put("Function", vec![func.code().into()]);
            }
        }
        // 核对实际口径：读得出来就以它为准；读不出来（getter 也抛、或给出认不出的码）时按引擎
        // 默认口径理解 —— 求和本来就是数值列的默认，不该为读不到就吭声
        let known = data_field.get("Function").ok().and_then(|value| value.int().ok()).and_then(AggFunc::from_code);
        if known == Some(func) || (known.is_none() && func == AggFunc::Sum) {
            return Ok(String::new());
        }
        let got = known.map_or("默认口径（数值列求和、文本列计数）".to_string(), |f| format!("「{}」", f.prefix()));
        Ok(format!(
            "值字段「{field}」实际口径是{got}（该引擎不让改，只给默认）——要的是「{}」：\
             别再重发 data（值字段已经在值区域里），改用 write 把公式写进单元格\
             （文本列先转成数值），或在回答里说明实际口径",
            func.prefix()
        ))
    }

    /// 值区域（数据主体），读取透视结果用。无值字段时报错。
    pub fn data_range(&self) -> Result<Range> {
        let res = self.raw.get("DataBodyRange")?;
        if !res.is_object() {
            bail!("透视表暂无值区域：先 add_data() 添加值字段");
        }
        Ok(Range::from_raw(res.idispatch()?))
    }

    /// 整表区域（含行列标题，TableRange2）。
    pub fn table_range(&self) -> Result<Range> {
        Ok(Range::from_raw(self.raw.get_obj("TableRange2", "TableRange2")?))
    }

    /// 页字段（报表筛选器）设值。
    pub fn set_page(&self, field: &str, value: &str) -> Result<()> {
        self.field(field)?
            .put("CurrentPage", vec![value.into()])
            .map_err(|e| {
                ExcelError::hint(
                    e,
                    format!("值须为字段 \"{field}\" 的现有项：用 items(\"{field}\") 查看；且字段须在页区域（set_fields 第三参）"),
                )
            })
    }

    /// 项筛选：只保留 `keep` 中的项（其余隐藏）。项名须与 items() 完全一致。
    pub fn filter_items(&self, field: &str, keep: &[&str]) -> Result<()> {
        if keep.is_empty() {
            bail!("filter_items：keep 不能为空（清筛选用 clear_field_filter）");
        }
        let coll = self
            .field(field)?
            .get_obj("PivotItems", "PivotItems")
            .map_err(|e| ExcelError::hint(e, "该字段无项：须先放进行/列/页区域"))?;
        let count = coll.get("Count")?.int()?;
        for i in 1..=count {
            let it = coll.call("Item", vec![i.into()])?.idispatch()?;
            let name = it.get("Name")?.string()?;
            let want = keep.iter().any(|k| *k == name);
            it.put("Visible", vec![want.into()])
                .map_err(|e| ExcelError::hint(e, "至少须保留一个可见项：keep 须含 items() 中的现有项名"))?;
        }
        Ok(())
    }

    /// 清除字段全部筛选（标签/手动项）。
    pub fn clear_field_filter(&self, field: &str) -> Result<()> {
        self.field(field)?.call("ClearAllFilters", vec![]).map(|_| ())
    }

    /// 行/列字段按值字段排序（`descending=false` 升序）。
    /// 值字段唯一名自动解析，AI 直接传数据源列名即可。
    pub fn sort_by_value(&self, field: &str, data_field: &str, descending: bool) -> Result<()> {
        let src = self.source_name_of(data_field);
        let order: i32 = if descending { 2 } else { 1 }; // xlDescending / xlAscending
        self.field(field)?
            .call("AutoSort", vec![order.into(), src.into()])
            .map_err(|e| ExcelError::hint(e, "字段须已放在行/列区域（set_fields），data_field 须为值字段（add_data）"))
            .map(|_| ())
    }

    /// 关闭字段自动排序。
    pub fn sort_off(&self, field: &str) -> Result<()> {
        self.field(field)?
            .call("AutoSort", vec![(-4105i32).into(), "".into()]) // xlManual
            .map(|_| ())
    }

    /// 行/列字段 Top-N / Bottom-N 展示（`top=false` 取后 N）。
    pub fn show_top(&self, field: &str, data_field: &str, n: i32, top: bool) -> Result<()> {
        let src = self.source_name_of(data_field);
        let rule: i32 = if top { 1 } else { 2 }; // xlTop / xlBottom
        self.field(field)?
            .call(
                "AutoShow",
                vec![(1i32).into(), rule.into(), n.into(), src.into()], // xlAutomatic
            )
            .map_err(|e| ExcelError::hint(e, "字段须在行/列区域，data_field 须为值字段；n ≥ 1"))
            .map(|_| ())
    }

    /// 关闭 Top-N 展示。
    pub fn show_off(&self, field: &str) -> Result<()> {
        self.field(field)?
            .call("AutoShow", vec![(2i32).into(), (1i32).into(), (1i32).into(), "".into()]) // xlManual
            .map(|_| ())
    }

    /// 分组定位格：字段的 DataRange 左上角（Group/Ungroup 必须打到单格，传整区域会静默失败）。
    fn group_cell(&self, field: &str) -> Result<IDispatchWrapper> {
        self.field(field)?
            .get_obj("DataRange", "DataRange")
            .map_err(|e| ExcelError::hint(e, format!("字段 \"{field}\" 无 DataRange：须先放进行/列区域（set_fields）")))? 
            .call("Cells", vec![(1i32).into(), (1i32).into()])?
            .idispatch()
    }

    /// 日期字段分组，如按年+月：`group_dates("日期", &[DatePart::Years, DatePart::Months])`。
    pub fn group_dates(&self, field: &str, parts: &[DatePart]) -> Result<()> {
        if parts.is_empty() {
            bail!("group_dates：parts 不能为空，如 &[DatePart::Years, DatePart::Months]");
        }
        let arr = list_to_variant(DatePart::periods(parts).into_iter().map(ComValue::Bool).collect())?;
        self.group_cell(field)?
            .call_named("Group", vec![], &[("Periods", ComValue::Raw(arr))])
            .map_err(|e| {
                ExcelError::hint(e, "日期列须为真日期值而非文本；已分组的字段先 ungroup()；分组结果为多个新字段")
            })
            .map(|_| ())
    }

    /// 数值字段按步长分组，如 `group_numbers("数量", 1000, None, None)` 每千一档。
    pub fn group_numbers(&self, field: &str, step: f64, start: Option<f64>, end: Option<f64>) -> Result<()> {
        if step <= 0.0 {
            bail!("group_numbers：step 须 > 0");
        }
        let mut named: Vec<(&str, ComValue)> = vec![("By", step.into())];
        if let Some(s) = start {
            named.push(("Start", s.into()));
        }
        if let Some(e) = end {
            named.push(("End", e.into()));
        }
        self.group_cell(field)?
            .call_named("Group", vec![], &named)
            .map_err(|e| ExcelError::hint(e, "字段须为行/列区域的数值字段；已分组的字段先 ungroup()"))
            .map(|_| ())
    }

    /// 取消字段分组。
    pub fn ungroup(&self, field: &str) -> Result<()> {
        self.group_cell(field)?.call("Ungroup", vec![]).map(|_| ())
    }

    /// 在 `anchor`（如 "K3" 或 "Sheet2!K3"）处为字段创建切片器。
    ///
    /// 切片器是**界面层对象**（浮动控件，同图表）：先 `app.set_fast(false)` 把屏幕刷新打开再调
    /// （f-manager 侧的 excel op 就是 `repaint()` 守卫）。
    ///
    /// 缓存用 **`Add2` + SourceField 传「字段名」** —— 这是 WPS 录制器与实测能跑通的形态
    /// （实报：同样调用 `Add2(pt, "班级")` 的宏在 WPS 上正常，而 `Add(pt, <PivotField>, 名)` 会
    /// 「不报错、集合里也有条目，但缓存没绑字段，界面上永远不出现」）。要判成败只看调用抛不抛，
    /// 别拿读回属性判（`Slicer.Width` 等读不出来是常事，会把真建出来的判成幽灵）。
    pub fn add_slicer(&self, field: &str, anchor: &str) -> Result<Slicer> {
        self.field(field)?; // 先校验字段存在：名字错时这里就报「字段不存在」，不用等引擎抛
        let ws = self.raw.get_obj("Parent", "Parent")?; // 透视表所在工作表
        let wb = ws.get_obj("Parent", "Parent")?; // 工作簿
        let a = resolve_cell(&ws, &wb, anchor)?;
        let (left, top) = (a.get("Left")?.float()?, a.get("Top")?.float()?);
        // 切片器名就用字段名（与录制器一致），缓存名用「切片器_字段」这个自有口径
        let cache_name = format!("切片器_{field}");
        drop_stale_slicers(&wb, &cache_name, field);
        let caches = wb.get_obj("SlicerCaches", "SlicerCaches")?;
        // 带缓存名是为了重跑时认出自己、清干净；引擎不收这个参数就退回录制器的两参形态
        let cache = caches
            .call(
                "Add2",
                vec![
                    ComValue::Obj(self.raw.disp.clone()),
                    field.into(),
                    cache_name.clone().into(),
                ],
            )
            .or_else(|_| caches.call("Add2", vec![ComValue::Obj(self.raw.disp.clone()), field.into()]))
            .map_err(|e| ExcelError::hint(e, slicer_hint()))?;
        // Slicers.Add(SlicerDestination, Level, Name, Caption, Top, Left, Width, Height)
        // Level 仅 OLAP 用，非 OLAP**必须真的省略**（用 Missing，不是 Empty —— 传 Empty 会被
        // 当成「Level = 空」去解析而抛 0x80020009，Excel 上实报）；宽高 144/180 磅兜底
        let slicer = cache
            .idispatch()?
            .get_obj("Slicers", "Slicers")?
            .call(
                "Add",
                vec![
                    ComValue::Obj(ws.disp.clone()),
                    ComValue::Missing,
                    field.into(),
                    field.into(),
                    top.into(),
                    left.into(),
                    (144f64).into(),
                    (180f64).into(),
                ],
            )
            .map_err(|e| ExcelError::hint(e, slicer_hint()))?
            .idispatch_child(field)?;
        Ok(Slicer::from_raw(slicer))
    }
}

/// 清掉同一个字段的切片器残留：`cache_name` 与 `slicer_name` 两份名字在工作簿里都必须唯一，
/// 留着就必然抛 0x80020009（实报：连试三次全失败）。非 OLAP 缓存一个字段一个，所以命中
/// 名字时**删整个缓存**（`SlicerCache.Delete` 连它名下的切片器一起带走）最省事，也顺带清掉
/// 「缓存已建但切片器没建成」的半成品。
fn drop_stale_slicers(wb: &IDispatchWrapper, cache_name: &str, slicer_name: &str) {
    let Ok(caches) = wb.get_obj("SlicerCaches", "SlicerCaches") else { return };
    let count = caches.get("Count").ok().and_then(|value| value.int().ok()).unwrap_or(0);
    for index in (1..=count).rev() {
        let Ok(cache) = caches.call("Item", vec![index.into()]).and_then(|value| value.idispatch()) else { continue };
        if cache.get("Name").ok().and_then(|value| value.string().ok()).as_deref() == Some(cache_name) {
            let _ = cache.call("Delete", vec![]);
            continue;
        }
        let Ok(slicers) = cache.get_obj("Slicers", "Slicers") else { continue };
        let total = slicers.get("Count").ok().and_then(|value| value.int().ok()).unwrap_or(0);
        for at in 1..=total {
            let Ok(slicer) = slicers.call("Item", vec![at.into()]).and_then(|value| value.idispatch()) else { continue };
            if slicer.get("Name").ok().and_then(|value| value.string().ok()).as_deref() == Some(slicer_name) {
                let _ = cache.call("Delete", vec![]);
                break;
            }
        }
    }
}

/// 切片器建不出来时的说明（SlicerCaches 与 Slicers 两处共用）
fn slicer_hint() -> &'static str {
    "字段名须与 fields() 一致、锚点单元格所在表须能放浮动控件，且切片器名在工作簿里唯一；\
     仍建不出来就改用项筛选（filter_items）达成同样筛选效果"
}

/// 切片器（Slicer COM 对象）：批量点选、全选、查看项。
pub struct Slicer {
    raw: IDispatchWrapper,
}

impl Slicer {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 标题。
    pub fn name(&self) -> Result<String> {
        match self.raw.get("Caption") {
            Ok(v) => v.string(),
            Err(_) => self.raw.get("Name")?.string(),
        }
    }

    /// 批量点选：只保留 `visible` 中的项。项名须与切片器显示项完全一致。
    pub fn set_items(&self, visible: &[&str]) -> Result<()> {
        if visible.is_empty() {
            bail!("set_items：visible 不能为空（全选用 select_all）");
        }
        let arr = list_to_variant(visible.iter().map(|s| (*s).into()).collect())?;
        self.raw
            .get_obj("SlicerCache", "SlicerCache")?
            .put_raw("VisibleSlicerItemsList", arr)
            .map_err(|e| {
                ExcelError::hint(e, "项名须与切片器显示项完全一致；WPS 未实现时改用透视表 filter_items()")
            })
    }

    /// 清除筛选（全选）。
    pub fn select_all(&self) -> Result<()> {
        self.raw
            .get_obj("SlicerCache", "SlicerCache")?
            .call("ClearManualFilter", vec![])
            .map(|_| ())
    }

    /// 切片器当前项（名称 + 是否选中）。
    pub fn items(&self) -> Result<Vec<SlicerItemInfo>> {
        let coll = self
            .raw
            .get_obj("SlicerCache", "SlicerCache")?
            .get_obj("SlicerCacheItems", "SlicerCacheItems")?;
        let count = coll.get("Count")?.int()?;
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let it = coll.call("Item", vec![i.into()])?.idispatch()?;
            let name = it.get("Name")?.string()?;
            let selected = matches!(it.get("Selected")?.value(), CellValue::Bool(true));
            out.push(SlicerItemInfo { name, selected });
        }
        Ok(out)
    }
}
