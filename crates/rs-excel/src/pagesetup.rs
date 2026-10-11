//! 页面设置（PageSetup）：页边距、纸张大小、纸张方向、打印区域、打印缩放、页眉页脚、打印标题。
//!
//! 经 `worksheet.page_setup()` 获取；边距/缩放单位为磅（point，1/72 英寸，
//! Excel 默认上下边距 0.75 英寸 ≈ 54 磅）。改动需 `workbook.save()` 后才落盘。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::ExcelError;

/// 纸张方向（xlOrientation）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Orientation {
    Portrait,
    Landscape,
}

impl Orientation {
    fn code(self) -> i32 {
        match self {
            Self::Portrait => 1,
            Self::Landscape => 2,
        }
    }
}

impl std::fmt::Display for Orientation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::Portrait => "纵向",
            Self::Landscape => "横向",
        })
    }
}

/// 纸张大小（常用 xlPaperSize 常量；完整表见 Excel 文档）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PaperSize {
    Letter,
    Legal,
    Executive,
    A3,
    A4,
    A5,
    B4,
    B5,
    A2,
    A1,
    /// 其他纸张代码直通
    Code(i32),
}

impl PaperSize {
    pub fn code(self) -> i32 {
        match self {
            Self::Letter => 1,
            Self::Legal => 5,
            Self::Executive => 7,
            Self::A3 => 8,
            Self::A4 => 9,
            Self::A5 => 11,
            Self::B4 => 12,
            Self::B5 => 13,
            Self::A2 => 66,
            Self::A1 => 67,
            Self::Code(c) => c,
        }
    }

    /// 按名解析（不区分大小写，如 "A4"/"letter"；纯数字视为纸张代码）。
    pub fn parse(s: &str) -> Option<Self> {
        let s = s.trim();
        if let Ok(c) = s.parse::<i32>() {
            return Some(Self::Code(c));
        }
        Some(match s.to_ascii_lowercase().as_str() {
            "letter" => Self::Letter,
            "legal" => Self::Legal,
            "executive" => Self::Executive,
            "a3" => Self::A3,
            "a4" => Self::A4,
            "a5" => Self::A5,
            "b4" => Self::B4,
            "b5" => Self::B5,
            "a2" => Self::A2,
            "a1" => Self::A1,
            _ => return None,
        })
    }

    /// 纸张代码 → 可读名（未知代码返回 "代码{n}"）。
    pub fn name_of(code: i32) -> String {
        match code {
            1 => "Letter".into(),
            5 => "Legal".into(),
            7 => "Executive".into(),
            8 => "A3".into(),
            9 => "A4".into(),
            11 => "A5".into(),
            12 => "B4".into(),
            13 => "B5".into(),
            66 => "A2".into(),
            67 => "A1".into(),
            n => format!("代码{n}"),
        }
    }
}

/// 页边距位（单位：磅）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Margin {
    Top,
    Bottom,
    Left,
    Right,
    Header,
    Footer,
}

impl Margin {
    fn prop(self) -> &'static str {
        match self {
            Self::Top => "TopMargin",
            Self::Bottom => "BottomMargin",
            Self::Left => "LeftMargin",
            Self::Right => "RightMargin",
            Self::Header => "HeaderMargin",
            Self::Footer => "FooterMargin",
        }
    }

    /// 按名解析（不区分大小写，如 "top"/"左"）。
    pub fn parse(s: &str) -> Option<Self> {
        Some(match s.to_ascii_lowercase().as_str() {
            "top" | "上" => Self::Top,
            "bottom" | "下" => Self::Bottom,
            "left" | "左" => Self::Left,
            "right" | "右" => Self::Right,
            "header" | "页眉" => Self::Header,
            "footer" | "页脚" => Self::Footer,
            _ => return None,
        })
    }
}

/// 页眉/页脚三等分位。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Section {
    Left,
    Center,
    Right,
}

impl Section {
    fn prop(self, kind: &str) -> String {
        format!(
            "{}{kind}",
            match self {
                Self::Left => "Left",
                Self::Center => "Center",
                Self::Right => "Right",
            }
        )
    }

    /// 按名解析（不区分大小写，如 "center"/"中"）。
    pub fn parse(s: &str) -> Option<Self> {
        Some(match s.to_ascii_lowercase().as_str() {
            "left" | "左" => Self::Left,
            "center" | "中" => Self::Center,
            "right" | "右" => Self::Right,
            _ => return None,
        })
    }
}

impl std::fmt::Display for Section {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
        })
    }
}

/// 页面设置（Worksheet 级 PageSetup COM 对象）。
pub struct PageSetup {
    raw: IDispatchWrapper,
}

impl PageSetup {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心：Orientation/PaperSize 等未封装成员可直调）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    fn put(&self, prop: &str, v: ComValue) -> Result<()> {
        self.raw.put(prop, vec![v]).map(|_| ())
    }

    // ── 页边距（磅） ──

    pub fn set_margin(&self, m: Margin, points: f64) -> Result<()> {
        self.put(m.prop(), points.into())
    }

    /// 一次设置上下左右边距（磅）。
    pub fn set_margins(&self, top: f64, bottom: f64, left: f64, right: f64) -> Result<()> {
        for (m, v) in [
            (Margin::Top, top),
            (Margin::Bottom, bottom),
            (Margin::Left, left),
            (Margin::Right, right),
        ] {
            self.set_margin(m, v)?;
        }
        Ok(())
    }

    pub fn margin(&self, m: Margin) -> Result<f64> {
        self.raw.get(m.prop())?.float()
    }

    // ── 纸张与方向 ──

    pub fn set_paper(&self, size: PaperSize) -> Result<()> {
        self.put("PaperSize", size.code().into())
    }

    /// 纸张代码（可用 [`PaperSize::name_of`] 转可读名）。
    pub fn paper(&self) -> Result<i32> {
        self.raw.get("PaperSize")?.int()
    }

    pub fn set_orientation(&self, o: Orientation) -> Result<()> {
        self.put("Orientation", o.code().into())
    }

    pub fn orientation(&self) -> Result<Orientation> {
        Ok(match self.raw.get("Orientation")?.int()? {
            2 => Orientation::Landscape,
            _ => Orientation::Portrait,
        })
    }

    // ── 打印区域 ──

    /// 设置打印区域（如 "A1:H40" 或多区域 "A1:B10,D1:F10"）。
    pub fn set_print_area(&self, addr: &str) -> Result<()> {
        self.put("PrintArea", addr.into())
            .map_err(|e| ExcelError::hint(e, "区域格式如 A1:H40，多区域用逗号分隔"))
    }

    /// 清除打印区域（恢复打印全部内容）。
    pub fn clear_print_area(&self) -> Result<()> {
        self.put("PrintArea", "".into())
    }

    /// 打印区域（未设置为空串/false）。
    pub fn print_area(&self) -> Result<String> {
        self.raw.get("PrintArea")?.string()
    }

    // ── 打印缩放 ──

    /// 按百分比缩放打印（10–400）。
    pub fn set_zoom(&self, percent: i32) -> Result<()> {
        self.put("Zoom", percent.into())
            .map_err(|e| ExcelError::hint(e, "缩放百分比有效范围 10–400"))
    }

    /// 恢复 100% 正常缩放（同时退出"适应页数"模式）。
    pub fn zoom_reset(&self) -> Result<()> {
        self.set_zoom(100)
    }

    /// 适应页宽/页高（0 = 该方向不限页数）。内部先关闭固定百分比缩放。
    pub fn fit_to(&self, pages_wide: i32, pages_tall: i32) -> Result<()> {
        self.put("Zoom", false.into())?; // Zoom=false 后 FitTo* 才生效
        self.put("FitToPagesWide", pages_wide.into())?;
        self.put("FitToPagesTall", pages_tall.into())
    }

    /// 缩放百分比；为 0 表示处于"适应页数"模式。
    pub fn zoom(&self) -> Result<i32> {
        self.raw.get("Zoom")?.int()
    }

    /// 适应页宽/页高设置（0 = 不限）。
    pub fn fit_pages(&self) -> Result<(i32, i32)> {
        let w = self.raw.get("FitToPagesWide")?.int()?;
        let h = self.raw.get("FitToPagesTall")?.int()?;
        Ok((w, h))
    }

    // ── 页眉页脚（支持 Excel 格式代码：&P 页码 &N 总页数 &D 日期 &T 时间 &F 文件名 &A 表名） ──

    pub fn set_header(&self, s: Section, text: &str) -> Result<()> {
        self.put(&s.prop("Header"), text.into())
    }

    pub fn set_footer(&self, s: Section, text: &str) -> Result<()> {
        self.put(&s.prop("Footer"), text.into())
    }

    pub fn header(&self, s: Section) -> Result<String> {
        self.raw.get(&s.prop("Header"))?.string()
    }

    pub fn footer(&self, s: Section) -> Result<String> {
        self.raw.get(&s.prop("Footer"))?.string()
    }

    // ── 打印标题（每页重复） ──

    /// 顶端标题行，如 "$1:$2"。
    pub fn set_title_rows(&self, addr: &str) -> Result<()> {
        self.put("PrintTitleRows", addr.into())
            .map_err(|e| ExcelError::hint(e, "行区域格式如 $1:$2"))
    }

    /// 左端标题列，如 "$A:$B"。
    pub fn set_title_cols(&self, addr: &str) -> Result<()> {
        self.put("PrintTitleColumns", addr.into())
            .map_err(|e| ExcelError::hint(e, "列区域格式如 $A:$B"))
    }

    /// 清除顶端标题行。
    pub fn clear_title_rows(&self) -> Result<()> {
        self.put("PrintTitleRows", "".into())
    }

    /// 清除左端标题列。
    pub fn clear_title_cols(&self) -> Result<()> {
        self.put("PrintTitleColumns", "".into())
    }

    /// 清除打印标题行/列。
    pub fn clear_titles(&self) -> Result<()> {
        self.clear_title_rows()?;
        self.clear_title_cols()
    }

    /// 顶端标题行（未设置为空串/false）。
    pub fn title_rows(&self) -> Result<String> {
        self.raw.get("PrintTitleRows")?.string()
    }

    /// 左端标题列（未设置为空串/false）。
    pub fn title_cols(&self) -> Result<String> {
        self.raw.get("PrintTitleColumns")?.string()
    }
}
