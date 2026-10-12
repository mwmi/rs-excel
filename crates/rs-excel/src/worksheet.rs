//! 工作表：信息、选择、增删、Range 工厂、区域级 TSV 批量读写、图表创建。

use anyhow::{bail, Result};

use crate::chart::Chart;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::{CellValue, ComValue};
use crate::error::{ole_to_hex, ExcelError};
use crate::pagesetup::PageSetup;
use crate::range::Range;
use crate::shape::Shapes;

/// 工作表（Worksheet COM 对象）。
pub struct Worksheet {
    raw: IDispatchWrapper,
}

impl Worksheet {
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

    /// 重命名。
    pub fn rename(&self, name: &str) -> Result<()> {
        self.raw
            .put("Name", vec![name.into()])
            .map_err(|e| ExcelError::hint(e, "名称不能与现有工作表重复，且不能含 : \\ / ? * [ ]"))
    }

    /// 在工作簿中的序号（1 基）。
    pub fn index(&self) -> Result<i32> {
        self.raw.get("Index")?.int()
    }

    /// 标签页颜色 `#RRGGBB`，未设置为空串。
    pub fn tab_color(&self) -> Result<String> {
        let v = self.raw.get_obj("Tab", "Tab")?.get("Color")?.value();
        Ok(match v {
            CellValue::Number(c) => ole_to_hex(c as u32),
            _ => String::new(),
        })
    }

    /// 设置标签页颜色（OLE 格式，用 `rgb()` 助手）。
    pub fn set_tab_color(&self, ole_color: u32) -> Result<()> {
        self.raw.get_obj("Tab", "Tab")?.put("Color", vec![ole_color.into()])
    }

    pub fn visible(&self) -> Result<bool> {
        Ok(matches!(self.raw.get("Visible")?.value(), CellValue::Number(v) if v == -1.0))
    }

    pub fn set_visible(&self, on: bool) -> Result<()> {
        self.raw.put("Visible", vec![(if on { -1i32 } else { 0 }).into()])
    }

    /// 激活工作表。
    pub fn activate(&self) -> Result<()> {
        self.raw.call("Activate", vec![]).map(|_| ())
    }

    /// 选中工作表。
    pub fn select(&self) -> Result<()> {
        self.raw.call("Select", vec![]).map(|_| ())
    }

    /// 删除工作表。
    pub fn delete(&self) -> Result<()> {
        self.raw.call("Delete", vec![]).map(|_| ())
    }

    /// 绑定区域，如 `range("A1")` / `range("A1:C3")` / 命名区域。
    pub fn range(&self, addr: &str) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw
                .call("Range", vec![addr.into()])?
                .idispatch_child(addr)
                .map_err(|e| ExcelError::hint(e, "地址格式如 A1、A1:C3 或命名区域"))?,
        ))
    }

    /// 按行列号绑定单元格（1 基）。
    pub fn cells(&self, row: i32, col: i32) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw
                .call("Cells", vec![row.into(), col.into()])?
                .idispatch_child(&format!("R{row}C{col}"))?,
        ))
    }

    /// 已使用区域（UsedRange）。
    pub fn used_range(&self) -> Result<Range> {
        let r = self.raw.get("UsedRange")?.idispatch()?;
        let addr = r.get("Address")?.string().unwrap_or_default();
        Ok(Range::from_raw(r.deeper(&addr)))
    }

    /// 已使用区域地址（如 `A1:F23`），判断数据范围的首选（一次 COM 调用）。
    pub fn used_addr(&self) -> Result<String> {
        self.used_range()?.address()
    }

    /// 已使用区域的行数。
    pub fn rows(&self) -> Result<i32> {
        self.used_range()?.rows()
    }

    /// 已使用区域的列数。
    pub fn columns(&self) -> Result<i32> {
        self.used_range()?.columns()
    }

    /// 在第 `at` 行前插入 `count` 行。
    pub fn insert_rows(&self, at: i32, count: i32) -> Result<()> {
        self.range(&row_block(at, count)?)?.insert(true)
    }

    /// 删除第 `at` 行起的 `count` 行。
    pub fn delete_rows(&self, at: i32, count: i32) -> Result<()> {
        self.range(&row_block(at, count)?)?.delete(true)
    }

    /// 在第 `at` 列前插入 `count` 列。
    pub fn insert_cols(&self, at: i32, count: i32) -> Result<()> {
        self.range(&col_block(at, count)?)?.insert(false)
    }

    /// 删除第 `at` 列起的 `count` 列。
    pub fn delete_cols(&self, at: i32, count: i32) -> Result<()> {
        self.range(&col_block(at, count)?)?.delete(false)
    }

    /// 保护工作表（内容不可编辑）。
    pub fn protect(&self, password: &str) -> Result<()> {
        self.raw
            .call_named(
                "Protect",
                vec![],
                &[("Password", password.into()), ("Contents", true.into())],
            )
            .map(|_| ())
    }

    /// 取消保护。
    pub fn unprotect(&self, password: &str) -> Result<()> {
        self.raw.call("Unprotect", vec![password.into()]).map(|_| ())
    }

    /// 全表查找（在已使用区域内包含匹配）。未找到返回 `Ok(None)`，不报错。
    pub fn find(&self, what: &str) -> Result<Option<Range>> {
        self.used_range()?.find(what, false)
    }

    /// 复制本表为副本：`after=None` 放到末尾，`Some(锚表名)` 放到该表之后。返回新表。
    pub fn copy_sheet(&self, after: Option<&str>) -> Result<Worksheet> {
        let wb = self.raw.get("Parent")?.idispatch()?;
        let anchor = match after {
            Some(name) => wb
                .call("Worksheets", vec![name.into()])?
                .idispatch()
                .map_err(|e| ExcelError::hint(e, "锚点工作表名不存在：先用 sheets() 列出"))?,
            None => {
                let sheets = wb.get_obj("Worksheets", "Worksheets")?;
                let count = sheets.get("Count")?.int()?;
                sheets.call("Item", vec![count.into()])?.idispatch()?
            }
        };
        self.raw
            .call_named("Copy", vec![], &[("After", ComValue::Obj(anchor.disp.clone()))])?;
        // Copy 不返回新表，新表复制后成为活动表
        let new_ws = self.raw.get("Application")?.idispatch()?.get_obj("ActiveSheet", "ActiveSheet")?;
        let name = new_ws.get("Name")?.string().unwrap_or_default();
        Ok(Worksheet::from_raw(new_ws.deeper(&name)))
    }

    /// 重算本表。
    pub fn calculate(&self) -> Result<()> {
        self.raw.call("Calculate", vec![]).map(|_| ())
    }

    /// 页面设置（页边距/纸张/方向/打印区域/缩放/页眉页脚/打印标题）。
    pub fn page_setup(&self) -> Result<PageSetup> {
        Ok(PageSetup::from_raw(
            self.raw.get_obj("PageSetup", "PageSetup")?,
        ))
    }

    /// 区域批量读为 TSV（Tab 分列、换行分行，AI/用户直接可读）。
    pub fn read_tsv(&self, addr: &str) -> Result<String> {
        self.range(addr)?.read_tsv()
    }

    /// 区域批量写：`tsv` 的形状自动适配，从 `anchor`（如 "A1"）开始向右下扩展。
    /// 数字/布尔自动转类型，`=` 开头写公式，前导 0 的串保持文本。
    pub fn write_tsv(&self, anchor: &str, tsv: &str) -> Result<()> {
        let grid = parse_tsv(tsv);
        if grid.is_empty() {
            bail!("write_tsv：没有可写入的数据");
        }
        let (c0, r0) = parse_addr(anchor)
            .ok_or_else(|| ExcelError::new("write_tsv", anchor, "锚点地址无法解析").with_hint("锚点格式如 A1"))?;
        let rows = grid.len() as i32;
        let cols = grid[0].len() as i32;
        let end = format!("{}{}", num_to_col(c0 + cols - 1), r0 + rows - 1);
        let rng = self.range(&format!("{anchor}:{end}"))?;
        rng.set_values(&grid)
    }

    /// 形状集合（图形/图片/文本框/连接线/图表等的增删改查入口）。
    pub fn shapes(&self) -> Result<Shapes> {
        Ok(Shapes::from_raw(self.raw.get_obj("Shapes", "Shapes")?))
    }

    /// 创建图表。`data_range` 数据区域，`anchor` 左上角锚点单元格。
    pub fn add_chart(
        &self,
        kind: crate::chart::ChartKind,
        data_range: &str,
        anchor: &str,
    ) -> Result<Chart> {
        let shapes = self.raw.get_obj("Shapes", "Shapes")?;
        let a = self.range(anchor)?;
        let left = a.raw().get("Left")?.float()?;
        let top = a.raw().get("Top")?.float()?;
        let width = a.raw().get("Width")?.float()? * 3.0;
        let height = a.raw().get("Height")?.float()? * 12.0;
        // AddChart2(Style=-1, XlChartType, Left, Top, Width, Height)；旧引擎回退 AddChart
        let shape = match shapes.call(
            "AddChart2",
            vec![
                (-1i32).into(),
                (kind as i32).into(),
                left.into(),
                top.into(),
                width.into(),
                height.into(),
            ],
        ) {
            Ok(s) => s,
            Err(_) => shapes.call(
                "AddChart",
                vec![
                    (kind as i32).into(),
                    left.into(),
                    top.into(),
                    width.into(),
                    height.into(),
                ],
            )?,
        };
        let chart = shape.idispatch_child("chart")?.get_obj("Chart", "Chart")?;
        let src = self.range(data_range)?;
        let src_obj = ComValue::Obj(src.raw().disp.clone());
        // SetSourceData：优先命名参数 Source:=，失败回退位置参数
        if chart
            .call_named("SetSourceData", vec![], &[("Source", src_obj.clone())])
            .is_err()
        {
            chart.call("SetSourceData", vec![src_obj]).map_err(|e| {
                ExcelError::hint(e, "数据区域无效，或工作表当前无可用数据源")
            })?;
        }
        Ok(crate::chart::Chart::from_raw(chart))
    }
}

/// 行区间地址，如 (5, 3) → "5:7"。
fn row_block(at: i32, count: i32) -> Result<String> {
    if at < 1 || count < 1 {
        return Err(ExcelError::new("row_block", "", "行列号须 ≥ 1")
            .with_hint("行号从 1 开始，count 至少为 1")
            .into());
    }
    Ok(if count == 1 {
        format!("{at}:{at}")
    } else {
        format!("{at}:{}", at + count - 1)
    })
}

/// 列区间地址，如 (2, 3) → "B:D"。
fn col_block(at: i32, count: i32) -> Result<String> {
    if at < 1 || count < 1 {
        return Err(ExcelError::new("col_block", "", "行列号须 ≥ 1")
            .with_hint("列号从 1（A 列）开始，count 至少为 1")
            .into());
    }
    Ok(format!("{}:{}", num_to_col(at), num_to_col(at + count - 1)))
}

/// 解析单元格地址（如 "A1"、"$B$12"）→ (列号, 行号)，1 基。
pub(crate) fn parse_addr(addr: &str) -> Option<(i32, i32)> {
    let a = addr.trim().trim_start_matches('$');
    let mut col = 0i32;
    let mut idx = 0usize;
    for ch in a.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
            idx += ch.len_utf8();
        } else {
            break;
        }
    }
    if col == 0 || idx == 0 {
        return None;
    }
    let row: i32 = a[idx..].trim_start_matches('$').parse().ok()?;
    if row < 1 {
        return None;
    }
    Some((col, row))
}

/// 列号 → 列字母（1 → "A"）。
pub(crate) fn num_to_col(mut n: i32) -> String {
    let mut s = String::new();
    while n > 0 {
        let r = ((n - 1) % 26) as u8;
        s.insert(0, (b'A' + r) as char);
        n = (n - 1) / 26;
    }
    s
}

/// TSV 文本 → 网格。规则："" → Empty；"true"/"false" → Bool；
/// 可解析为数字且无前导 0 → Number；其余 → Text（"=" 开头经 Value 写入即公式）。
/// 末尾全空行自动去除。
pub(crate) fn parse_tsv(text: &str) -> Vec<Vec<CellValue>> {
    let mut rows: Vec<Vec<CellValue>> = text
        .lines()
        .map(|line| {
            line.trim_end_matches('\r')
                .split('\t')
                .map(|c| {
                    let c = c.trim();
                    if c.is_empty() {
                        CellValue::Empty
                    } else if c == "true" {
                        CellValue::Bool(true)
                    } else if c == "false" {
                        CellValue::Bool(false)
                    } else if let Ok(f) = c.parse::<f64>() {
                        let has_leading_zero =
                            c.starts_with('0') && c.len() > 1 && !c.starts_with("0.");
                        // "inf"/"nan" 等也能被 parse 为 f64，按文本保留
                        let not_number = has_leading_zero
                            || c.eq_ignore_ascii_case("inf")
                            || c.eq_ignore_ascii_case("infinity")
                            || c.eq_ignore_ascii_case("nan");
                        if not_number {
                            CellValue::Text(c.to_string())
                        } else {
                            CellValue::Number(f)
                        }
                    } else {
                        CellValue::Text(c.to_string())
                    }
                })
                .collect()
        })
        .collect();
    while let Some(last) = rows.last() {
        if last.iter().all(|c| matches!(c, CellValue::Empty)) {
            rows.pop();
        } else {
            break;
        }
    }
    rows
}
