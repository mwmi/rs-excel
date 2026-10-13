//! 条件格式（FormatConditions）：单元格值/公式/文本/重复唯一值/前后 N 项/平均偏差/空值/错误值/
//! 日期周期规则，以及色阶/数据条/图标集，规则优先级与删除。
//!
//! 经 `range.format_conditions()` 获取集合管理入口；快速添加用 `Range` 上的
//! `cond_xxx` 一次性方法（添加规则 + 应用命中格式一步完成）：
//!
//! ```no_run
//! # let sheet = unreachable!();
//! use rs_excel::{rgb, CondFormat, CondOp};
//! let rng = sheet.range("B2:B20")?;
//! // 大于 100 标红加粗
//! rng.cond_value(CondOp::Greater, "100", None, &CondFormat::new().fill(rgb(255, 0, 0)).bold())?;
//! // 三色阶
//! rng.cond_color_scale(rgb(248, 105, 107), rgb(255, 235, 156), Some(rgb(99, 190, 123)))?;
//! // 查看已有规则（单行）/全部清除
//! println!("{}", rng.format_conditions()?.list()?);
//! rng.clear_cond()?;
//! # Ok::<(), anyhow::Error>(())
//! ```
//!
//! 注意：公式类条件基于区域**左上角**相对引用；未覆盖的冷门场景走
//! `range.raw().get("FormatConditions")` 直调 COM。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::{ole_to_hex, ExcelError};
use crate::range::Range;

// ── 枚举（屏蔽数值常量，AI 无需查表） ──

/// 单元格值比较运算符（xlFormatConditionOperator 常用子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CondOp {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    Greater,
    Less,
    GreaterEqual,
    LessEqual,
}

impl CondOp {
    pub fn code(self) -> i32 {
        match self {
            Self::Between => 1,
            Self::NotBetween => 2,
            Self::Equal => 3,
            Self::NotEqual => 4,
            Self::Greater => 5,
            Self::Less => 6,
            Self::GreaterEqual => 7,
            Self::LessEqual => 8,
        }
    }

    /// 由 COM 读回的 `Operator` 常量反推（读不出按 Equal）。
    pub fn from_code(code: Option<i32>) -> Self {
        match code {
            Some(1) => Self::Between,
            Some(2) => Self::NotBetween,
            Some(3) => Self::Equal,
            Some(4) => Self::NotEqual,
            Some(5) => Self::Greater,
            Some(6) => Self::Less,
            Some(7) => Self::GreaterEqual,
            Some(8) => Self::LessEqual,
            _ => Self::Equal,
        }
    }
}

/// 文本匹配方式（xlTextOperator）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CondText {
    Contains,
    NotContains,
    BeginsWith,
    EndsWith,
}

impl CondText {
    pub fn code(self) -> i32 {
        match self {
            Self::Contains => 0,
            Self::NotContains => 1,
            Self::BeginsWith => 2,
            Self::EndsWith => 3,
        }
    }

    /// 由 COM 读回的 `TextOperator` 常量反推（读不出按 Contains）。
    pub fn from_code(code: Option<i32>) -> Self {
        match code {
            Some(1) => Self::NotContains,
            Some(2) => Self::BeginsWith,
            Some(3) => Self::EndsWith,
            _ => Self::Contains,
        }
    }
}

/// 日期周期（xlDateOperator，"今天的日期"类规则）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DateOp {
    Today,
    Yesterday,
    Tomorrow,
    Last7Days,
    ThisMonth,
    LastMonth,
    NextMonth,
    ThisWeek,
    LastWeek,
    NextWeek,
}

impl DateOp {
    pub fn code(self) -> i32 {
        match self {
            Self::Today => 0,
            Self::Yesterday => 1,
            Self::Tomorrow => 2,
            Self::Last7Days => 3,
            Self::ThisMonth => 4,
            Self::LastMonth => 5,
            Self::NextMonth => 6,
            Self::ThisWeek => 7,
            Self::LastWeek => 8,
            Self::NextWeek => 9,
        }
    }
}

/// 图标集（xlIconSet 常用子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum IconSet {
    Arrows3,
    ArrowsGray3,
    Flags3,
    Rating4,
    Rating5,
    TrafficLights3,
    Arrows4,
    Arrows5,
    Quarters5,
    /// 其余图标集代码直通（完整表见 Excel 文档 XlIconSet）
    Code(i32),
}

impl IconSet {
    pub fn code(self) -> i32 {
        match self {
            Self::Arrows3 => 5,
            Self::ArrowsGray3 => 6,
            Self::Flags3 => 7,
            Self::Rating4 => 3,
            Self::Rating5 => 8,
            Self::TrafficLights3 => 4,
            Self::Arrows4 => 1,
            Self::Arrows5 => 12,
            Self::Quarters5 => 19,
            Self::Code(c) => c,
        }
    }
}

// ── 命中格式 ──

/// 条件命中时的格式（链式构造，未设置项不改动）。
///
/// ```ignore
/// CondFormat::new().fill(rgb(255, 0, 0)).font(rgb(0, 0, 255)).bold()
/// ```
#[derive(Default)]
pub struct CondFormat {
    fill: Option<u32>,
    font: Option<u32>,
    bold: Option<bool>,
}

impl CondFormat {
    pub fn new() -> Self {
        Self::default()
    }

    /// 命中填充色（用 [`rgb`] 构造）。
    pub fn fill(mut self, ole: u32) -> Self {
        self.fill = Some(ole);
        self
    }

    /// 命中字色（用 [`rgb`] 构造）。
    pub fn font(mut self, ole: u32) -> Self {
        self.font = Some(ole);
        self
    }

    /// 命中加粗。
    pub fn bold(mut self) -> Self {
        self.bold = Some(true);
        self
    }
}

/// 给规则对象应用命中格式（一次性方法的公共尾部）。
fn apply_format(cond: &IDispatchWrapper, fmt: &CondFormat) -> Result<()> {
    if let Some(c) = fmt.fill {
        cond.get_obj("Interior", "Interior")?.put("Color", vec![c.into()])?;
    }
    if fmt.font.is_some() || fmt.bold.is_some() {
        let f = cond.get_obj("Font", "Font")?;
        if let Some(c) = fmt.font {
            f.put("Color", vec![c.into()])?;
        }
        if let Some(b) = fmt.bold {
            f.put("Bold", vec![b.into()])?;
        }
    }
    Ok(())
}

// ── 集合与单条规则 ──

/// 一条条件格式规则的快照（读不出的项留空，重建时按缺省处理）。
///
/// 为什么要它：`FormatConditions` 没有克隆接口，想撤销「AI 加的条件格式」只能把每条
/// 规则的判定与命中格式逐项读出来存着，撤销时先 `delete_all` 再逐条 `Add` 回来。
#[derive(Clone, Debug, Default)]
pub struct CondSnap {
    /// xlFormatConditionType；`<= 0` = 这条没读出来，重建时跳过
    pub type_code: i32,
    pub operator: Option<i32>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub text_op: Option<i32>,
    pub string: Option<String>,
    pub dupe_unique: Option<i32>,
    pub rank: Option<i32>,
    pub percent: Option<bool>,
    pub top_bottom: Option<i32>,
    pub above_below: Option<i32>,
    pub stop_if_true: Option<bool>,
    pub fill: Option<u32>,
    pub font_color: Option<u32>,
    pub bold: Option<bool>,
    /// 色阶的 2~3 个颜色（低→中→高）
    pub scale_colors: Vec<u32>,
    /// 数据条颜色
    pub bar_color: Option<u32>,
    /// 图标集代码
    pub icon_set: Option<i32>,
}

/// 条件格式集合（Range.FormatConditions）。快速添加见 [`Range`] 的 `cond_xxx` 方法。
pub struct FormatConditions {
    raw: IDispatchWrapper,
}

impl FormatConditions {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 规则数量。
    pub fn count(&self) -> Result<i32> {
        self.raw.get("Count")?.int()
    }

    /// 按索引取规则（1 基）。
    pub fn item(&self, index: i32) -> Result<FormatCondition> {
        Ok(FormatCondition {
            raw: self
                .raw
                .call("Item", vec![index.into()])?
                .idispatch_child(&format!("规则{index}"))?,
        })
    }

    /// 单行列出全部规则（AI 读一次即知全貌），如
    /// `#1 单元格值 100 填充#FF0000 | #2 色阶`；无规则返回固定文案。
    pub fn list(&self) -> Result<String> {
        let n = self.count()?;
        let mut parts = Vec::with_capacity(n.max(0) as usize);
        for i in 1..=n {
            parts.push(self.item(i).map(|c| c.info()).unwrap_or_else(|_| format!("#{i} <读取失败>")));
        }
        Ok(if parts.is_empty() {
            "无条件格式规则".into()
        } else {
            parts.join(" | ")
        })
    }

    /// 删除本区域全部条件格式。
    pub fn delete_all(&self) -> Result<()> {
        self.raw.call("Delete", vec![]).map(|_| ())
    }

    /// 快照全部规则（顺序即优先级顺序，供 [`Range::apply_conds`] 原样重建）。
    pub fn snapshot(&self) -> Result<Vec<CondSnap>> {
        let n = self.count()?.max(0);
        let mut out = Vec::with_capacity(n as usize);
        for i in 1..=n {
            out.push(self.item(i)?.snapshot());
        }
        Ok(out)
    }
}

/// 单条条件格式规则（FormatCondition）。
pub struct FormatCondition {
    raw: IDispatchWrapper,
}

impl FormatCondition {
    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 命中填充色（未设置返回 None）。
    pub fn fill_color(&self) -> Option<u32> {
        let c = self
            .raw
            .get("Interior")
            .ok()?
            .idispatch()
            .ok()?
            .get("Color")
            .ok()?
            .int()
            .ok()?;
        (c > 0 && c < 0x0100_0000).then_some(c as u32)
    }

    /// 命中字色（未设置返回 None）。
    pub fn font_color(&self) -> Option<u32> {
        let c = self
            .raw
            .get("Font")
            .ok()?
            .idispatch()
            .ok()?
            .get("Color")
            .ok()?
            .int()
            .ok()?;
        (c > 0 && c < 0x0100_0000).then_some(c as u32)
    }

    /// 命中加粗（未设置返回 None）。
    pub fn bold(&self) -> Option<bool> {
        self.raw.get("Font").ok()?.idispatch().ok()?.get("Bold").and_then(|v| v.bool()).ok()
    }

    /// 快照本条规则（判定 + 命中格式），供 [`Range::apply_conds`] 重建。
    pub fn snapshot(&self) -> CondSnap {
        let int_of = |name: &str| self.raw.get(name).and_then(|v| v.int()).ok();
        let txt_of = |name: &str| {
            self.raw
                .get(name)
                .ok()
                .and_then(|v| v.string().ok())
                .filter(|s| !s.is_empty())
        };
        let mut snap = CondSnap {
            type_code: int_of("Type").unwrap_or(0),
            operator: int_of("Operator"),
            formula1: txt_of("Formula1"),
            formula2: txt_of("Formula2"),
            text_op: int_of("TextOperator"),
            string: txt_of("String"),
            dupe_unique: int_of("DupeUnique"),
            rank: int_of("Rank"),
            percent: self.raw.get("Percent").and_then(|v| v.bool()).ok(),
            top_bottom: int_of("TopBottom"),
            above_below: int_of("AboveBelow"),
            stop_if_true: self.raw.get("StopIfTrue").and_then(|v| v.bool()).ok(),
            fill: self.fill_color(),
            font_color: self.font_color(),
            bold: self.bold(),
            ..Default::default()
        };
        match snap.type_code {
            // 色阶：逐个 ColorScaleCriteria 读 FormatColor（2 色阶梯没有第 3 个，读失败即止）
            3 => {
                for i in 1..=3 {
                    let Ok(c) = self
                        .raw
                        .call("ColorScaleCriteria", vec![i.into()])
                        .and_then(|v| v.idispatch())
                        .and_then(|v| v.get_obj("FormatColor", "FormatColor"))
                        .and_then(|v| v.get("Color"))
                        .and_then(|v| v.int())
                    else {
                        break;
                    };
                    if c > 0 && c < 0x0100_0000 {
                        snap.scale_colors.push(c as u32);
                    }
                }
            }
            // 数据条
            4 => {
                snap.bar_color = self
                    .raw
                    .get_obj("BarColor", "BarColor")
                    .ok()
                    .and_then(|v| v.get("Color").and_then(|v| v.int()).ok())
                    .filter(|c| *c > 0 && *c < 0x0100_0000)
                    .map(|c| c as u32);
            }
            // 图标集：常量或 IconSet 对象都试
            6 => snap.icon_set = self.raw.get("IconSet").and_then(|v| v.int()).ok(),
            _ => {}
        }
        snap
    }

    /// 单行描述：优先级、类型、条件公式、填充色。
    pub fn info(&self) -> String {
        let int_of = |name: &str| self.raw.get(name).and_then(|v| v.int()).ok();
        let txt_of = |name: &str| {
            self.raw
                .get(name)
                .ok()
                .and_then(|v| v.string().ok())
                .filter(|s| !s.is_empty())
        };
        let mut s = format!(
            "#{} {}",
            int_of("Priority").unwrap_or(0),
            match int_of("Type").unwrap_or(-1) {
                1 => "单元格值",
                2 => "公式",
                3 => "色阶",
                4 => "数据条",
                5 => "前后N项",
                6 => "图标集",
                7 => "重复值",
                8 => "唯一值",
                9 => "文本",
                10 => "空值",
                11 => "日期",
                12 => "平均值",
                _ => "未知类型",
            }
        );
        if let Some(f) = txt_of("Formula1") {
            s.push_str(&format!(" {f}"));
        }
        if let Some(f) = txt_of("Formula2") {
            s.push_str(&format!(" ~ {f}"));
        }
        if let Some(c) = self.fill_color() {
            s.push_str(&format!(" 填充{}", ole_to_hex(c)));
        }
        s
    }

    /// 修改命中格式（未设置项不改动）。
    pub fn set_format(&self, fmt: &CondFormat) -> Result<()> {
        apply_format(&self.raw, fmt)
    }

    /// 调整优先级（数字越小越先应用；或用 [`Self::set_first_priority`]）。
    pub fn set_priority(&self, priority: i32) -> Result<()> {
        self.raw.put("Priority", vec![priority.into()])
    }

    /// 设为最高优先级。
    pub fn set_first_priority(&self) -> Result<()> {
        self.raw.call("SetFirstPriority", vec![]).map(|_| ())
    }

    /// 命中后是否停止应用后续规则。
    pub fn set_stop_if_true(&self, stop: bool) -> Result<()> {
        self.raw.put("StopIfTrue", vec![stop.into()])
    }

    /// 删除本条规则。
    pub fn delete(&self) -> Result<()> {
        self.raw.call("Delete", vec![]).map(|_| ())
    }
}

// ── Range 上的条件格式 API ──

impl Range {
    /// 条件格式集合（管理入口：查数量/逐条读/删全部；快速添加用 `cond_xxx` 系列）。
    pub fn format_conditions(&self) -> Result<FormatConditions> {
        Ok(FormatConditions::from_raw(
            self.raw().get_obj("FormatConditions", "条件格式")?,
        ))
    }

    /// 删除本区域全部条件格式。
    pub fn clear_cond(&self) -> Result<()> {
        self.format_conditions()?.delete_all()
    }

    /// 写回条件格式快照（先清空本区域，再按快照顺序逐条重建 = 撤销/恢复条件格式用）。
    ///
    /// `type_code <= 0` 的条目跳过（当时就没读出来）；单条重建失败不影响其余条目 ——
    /// 与格式快照还原同口径：外观类还原不该让整条撤销失败卡死栈。
    pub fn apply_conds(&self, snaps: &[CondSnap]) -> Result<()> {
        self.clear_cond()?;
        for s in snaps {
            if s.type_code <= 0 {
                continue;
            }
            let fmt = CondFormat {
                fill: s.fill,
                font: s.font_color,
                bold: s.bold,
            };
            let _ = match s.type_code {
                1 => self.cond_value(
                    CondOp::from_code(s.operator),
                    s.formula1.as_deref().unwrap_or_default(),
                    s.formula2.as_deref(),
                    &fmt,
                ),
                2 => self.cond_formula(s.formula1.as_deref().unwrap_or_default(), &fmt),
                9 => self.cond_text(CondText::from_code(s.text_op), s.string.as_deref().unwrap_or_default(), &fmt),
                // xlDuplicate = 7 / xlUnique = 8
                7 | 8 => self.cond_dupe(s.type_code == 8, &fmt),
                5 => self.cond_top(
                    s.rank.unwrap_or(10),
                    s.percent.unwrap_or(false),
                    s.top_bottom == Some(2),
                    &fmt,
                ),
                12 => self.cond_average(s.above_below == Some(1), &fmt),
                3 => self.cond_color_scale(
                    s.scale_colors.first().copied().unwrap_or(0),
                    s.scale_colors.get(1).copied().unwrap_or(0),
                    s.scale_colors.get(2).copied(),
                ),
                4 => self.cond_data_bar(s.bar_color.unwrap_or(0x63BE7B)),
                6 => self.cond_icon(IconSet::Code(s.icon_set.unwrap_or(3))),
                _ => Ok(()),
            };
        }
        Ok(())
    }

    /// 添加规则并应用命中格式（一次性方法的公共尾部）。
    fn cond_add(
        &self,
        args: Vec<ComValue>,
        named: &[(&str, ComValue)],
        hint: &str,
        fmt: &CondFormat,
    ) -> Result<()> {
        let cond = self
            .raw()
            .get_obj("FormatConditions", "条件格式")?
            .call_named("Add", args, named)
            .map_err(|e| ExcelError::hint(e, hint))?
            .idispatch_child("规则")?;
        apply_format(&cond, fmt)
    }

    /// 单元格值规则：如 `cond_value(CondOp::Greater, "100", None, &fmt)` 标红大于 100。
    /// 值可直接写 `"100"`，也可引用 `"=$B$1"`；`v2` 仅 Between/NotBetween 需要。
    pub fn cond_value(&self, op: CondOp, v1: &str, v2: Option<&str>, fmt: &CondFormat) -> Result<()> {
        let mut named = vec![
            ("Operator", op.code().into()),
            ("Formula1", v1.into()),
        ];
        if let Some(v2) = v2 {
            named.push(("Formula2", v2.into()));
        }
        self.cond_add(
            vec![(1i32).into()],
            &named,
            "值可写 \"100\" 或引用 \"=$B$1\"；v2 仅 Between/NotBetween 需要",
            fmt,
        )
    }

    /// 公式规则，`formula` 不以 = 开头会自动补。
    /// 公式基于区域**左上角**相对引用，如 `cond_formula("MOD(ROW(),2)=0", &fmt)` 偶数行填充。
    pub fn cond_formula(&self, formula: &str, fmt: &CondFormat) -> Result<()> {
        let f = if formula.starts_with('=') {
            formula.to_string()
        } else {
            format!("={formula}")
        };
        self.cond_add(
            vec![(2i32).into()],
            &[("Formula1", f.into())],
            "公式基于区域左上角相对引用，如 MOD(ROW(),2)=0",
            fmt,
        )
    }

    /// 文本规则：包含/不包含/开头/结尾匹配。
    pub fn cond_text(&self, kind: CondText, text: &str, fmt: &CondFormat) -> Result<()> {
        self.cond_add(
            vec![(9i32).into()],
            &[
                ("TextOperator", kind.code().into()),
                ("String", text.into()),
            ],
            "kind 为 Contains/NotContains/BeginsWith/EndsWith",
            fmt,
        )
    }

    /// 重复值/唯一值规则：`unique=false` 标重复，`true` 标唯一。
    pub fn cond_dupe(&self, unique: bool, fmt: &CondFormat) -> Result<()> {
        let cond = self
            .raw()
            .get_obj("FormatConditions", "条件格式")?
            .call("AddUniqueValues", vec![])?
            .idispatch_child("规则")?;
        // xlDuplicate = 1, xlUnique = 0
        cond.put("DupeUnique", vec![(if unique { 0i32 } else { 1 }).into()])?;
        apply_format(&cond, fmt)
    }

    /// 前/后 N 项规则：`percent=true` 时 n 按百分比；`bottom=true` 取最小 N 项。
    pub fn cond_top(&self, n: i32, percent: bool, bottom: bool, fmt: &CondFormat) -> Result<()> {
        let cond = self
            .raw()
            .get_obj("FormatConditions", "条件格式")?
            .call("AddTop10", vec![])?
            .idispatch_child("规则")?;
        cond.put("Rank", vec![n.into()])?;
        cond.put("Percent", vec![percent.into()])?;
        // xlTop10Top = 1, xlTop10Bottom = 2
        cond.put("TopBottom", vec![(if bottom { 2i32 } else { 1 }).into()])?;
        apply_format(&cond, fmt)
    }

    /// 高于/低于平均值规则：`below=true` 取低于平均。
    pub fn cond_average(&self, below: bool, fmt: &CondFormat) -> Result<()> {
        let cond = self
            .raw()
            .get_obj("FormatConditions", "条件格式")?
            .call("AddAboveAverage", vec![])?
            .idispatch_child("规则")?;
        // xlAboveAverage = 0, xlBelowAverage = 1
        cond.put("AboveBelow", vec![(if below { 1i32 } else { 0 }).into()])?;
        apply_format(&cond, fmt)
    }

    /// 空单元格规则（公式法实现，全引擎可靠）。
    pub fn cond_blanks(&self, fmt: &CondFormat) -> Result<()> {
        let f = format!("=ISBLANK({})", self.top_left_addr()?);
        self.cond_add(vec![(2i32).into()], &[("Formula1", f.into())], "", fmt)
    }

    /// 错误值规则（公式法实现，全引擎可靠）。
    pub fn cond_errors(&self, fmt: &CondFormat) -> Result<()> {
        let f = format!("=ISERROR({})", self.top_left_addr()?);
        self.cond_add(vec![(2i32).into()], &[("Formula1", f.into())], "", fmt)
    }

    /// 色阶：2 色传 `c3=None`，3 色传全部（低→中→高，颜色用 [`rgb`] 构造）。
    pub fn cond_color_scale(&self, c1: u32, c2: u32, c3: Option<u32>) -> Result<()> {
        // xl2ColorScale = 2, xl3ColorScale = 3
        let cs = self
            .raw()
            .get_obj("FormatConditions", "条件格式")?
            .call("AddColorScale", vec![(if c3.is_some() { 3i32 } else { 2 }).into()])?
            .idispatch_child("色阶")?;
        let set_color = |i: i32, c: u32| -> Result<()> {
            cs.call("ColorScaleCriteria", vec![i.into()])?
                .idispatch()?
                .get_obj("FormatColor", "FormatColor")?
                .put("Color", vec![c.into()])
        };
        set_color(1, c1)?;
        set_color(2, c2)?;
        if let Some(c) = c3 {
            set_color(3, c)?;
        }
        Ok(())
    }

    /// 数据条（颜色用 [`rgb`] 构造）。
    pub fn cond_data_bar(&self, color: u32) -> Result<()> {
        self.raw()
            .get_obj("FormatConditions", "条件格式")?
            .call("AddDatabar", vec![])?
            .idispatch_child("数据条")?
            .get_obj("BarColor", "BarColor")?
            .put("Color", vec![color.into()])
    }

    /// 图标集规则。
    pub fn cond_icon(&self, set: IconSet) -> Result<()> {
        let ics = self
            .raw()
            .get_obj("FormatConditions", "条件格式")?
            .call("AddIconSetCondition", vec![])?
            .idispatch_child("图标集")?;
        match ics.put("IconSet", vec![set.code().into()]) {
            Ok(()) => Ok(()),
            // 个别引擎要求 IconSet 对象而非常量：经 ActiveWorkbook.IconSets 取
            Err(_) => {
                let sets = self
                    .raw()
                    .get("Worksheet")?
                    .idispatch()?
                    .get("Parent")?
                    .idispatch()?
                    .get("IconSets")?
                    .idispatch()?;
                let obj = sets.call("Item", vec![(set.code()).into()])?.idispatch()?;
                ics.put("IconSet", vec![ComValue::Obj(obj.disp.clone())])
            }
        }
    }

    /// 日期周期规则（今天/昨天/本周等，按单元格日期值判定）。
    /// 若引擎不支持，错误提示会建议改用 [`Range::cond_formula`] 公式实现。
    pub fn cond_date(&self, op: DateOp, fmt: &CondFormat) -> Result<()> {
        // xlTimePeriod = 11，Operator 位携带日期算子
        self.cond_add(
            vec![(11i32).into()],
            &[("Operator", op.code().into())],
            "引擎不支持日期规则时，改用 cond_formula(\"=TODAY()\", ..) 等公式实现",
            fmt,
        )
    }

    /// 区域左上角单元格地址（无 $，如 `A1`），公式条件引用基准。
    fn top_left_addr(&self) -> Result<String> {
        let a = self.address()?;
        Ok(a.split([':', ',']).next().unwrap_or("A1").replace('$', ""))
    }
}
