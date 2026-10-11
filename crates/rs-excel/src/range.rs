//! 区域（Range）：单元格/区域读写、公式、合并、格式、对齐、行高列宽、批量 TSV、样式入口。

use anyhow::{bail, Result};

use crate::com::dispatch::IDispatchWrapper;
use crate::com::safearray::{rows_to_variant, variant_to_rows};
use crate::com::variant::{CellValue, ComValue};
use crate::error::ExcelError;
use crate::style::{Borders, Font, Interior};
use crate::worksheet::parse_tsv;

/// 水平对齐（xlHAlign 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AlignH {
    Left = -4131,
    Center = -4108,
    Right = -4152,
}

/// 垂直对齐（xlVAlign 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AlignV {
    Top = -4160,
    Center = -4108,
    Bottom = -4107,
}

/// 区域（Range COM 对象）：读写均为区域级批量。
pub struct Range {
    raw: IDispatchWrapper,
}

impl Range {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 绝对地址，如 `$A$1:$C$3`。
    pub fn address(&self) -> Result<String> {
        self.raw.get("Address")?.string()
    }

    /// 行数。
    pub fn rows(&self) -> Result<i32> {
        self.raw.get_obj("Rows", "Rows")?.get("Count")?.int()
    }

    /// 列数。
    pub fn columns(&self) -> Result<i32> {
        self.raw.get_obj("Columns", "Columns")?.get("Count")?.int()
    }

    // ── 读 ──

    /// 单格值（多格区域返回左上角值；底层 SAFEARRAY 标量/数组统一处理）。
    pub fn value(&self) -> Result<CellValue> {
        let res = self.raw.get("Value")?;
        let rows = variant_to_rows(&res.v)?;
        Ok(rows.into_iter().flatten().next().unwrap_or(CellValue::Empty))
    }

    /// 显示文本（应用格式化后的样子）。
    pub fn text(&self) -> Result<String> {
        self.raw.get("Text")?.string()
    }

    /// 批量读整个区域为网格。
    pub fn values(&self) -> Result<Vec<Vec<CellValue>>> {
        variant_to_rows(&self.raw.get("Value")?.v)
    }

    /// 批量读为 TSV 文本（Tab 分列、换行分行）。
    pub fn read_tsv(&self) -> Result<String> {
        let grid = self.values()?;
        Ok(grid
            .iter()
            .map(|row| {
                row.iter()
                    .map(CellValue::to_tsv)
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n"))
    }

    /// 公式（无公式返回空串）。
    pub fn formula(&self) -> Result<String> {
        Ok(match self.raw.get("Formula")?.value() {
            CellValue::Text(s) => s,
            _ => String::new(),
        })
    }

    // ── 写 ──

    /// 写单值（区域会被整体填充同值）。
    pub fn set_value(&self, v: impl Into<ComValue>) -> Result<()> {
        self.raw.put("Value", vec![v.into()])
    }

    /// 批量写网格。形状必须与区域一致；SAFEARRAY 批量失败自动回退逐格。
    pub fn set_values(&self, grid: &[Vec<CellValue>]) -> Result<()> {
        if grid.is_empty() || grid[0].is_empty() {
            bail!("set_values：数据为空");
        }
        match rows_to_variant(grid) {
            Ok(v) => self.raw.put_raw("Value", v),
            Err(_) => {
                // 回退：逐格写（相对区域左上角）
                for (r, row) in grid.iter().enumerate() {
                    for (c, val) in row.iter().enumerate() {
                        self.raw
                            .call("Cells", vec![(r as i32 + 1).into(), (c as i32 + 1).into()])?
                            .idispatch()?
                            .put("Value", vec![ComValue::from(val.clone())])?;
                    }
                }
                Ok(())
            }
        }
    }

    /// 批量写 TSV 文本（形状必须与区域一致）。
    pub fn write_tsv(&self, tsv: &str) -> Result<()> {
        let grid = parse_tsv(tsv);
        if grid.is_empty() {
            bail!("write_tsv：没有可写入的数据");
        }
        self.set_values(&grid)
    }

    /// 写公式，如 `set_formula("=SUM(A1:A9)")`。
    pub fn set_formula(&self, f: &str) -> Result<()> {
        self.raw
            .put("Formula", vec![f.into()])
            .map_err(|e| ExcelError::hint(e, "公式需以 = 开头，如 =SUM(A1:A9)"))
    }

    /// 数字格式，如 `number_format("0.00")` / `"yyyy-mm-dd"` / `"0.00%"`。
    pub fn number_format(&self, fmt: &str) -> Result<()> {
        self.raw.put("NumberFormat", vec![fmt.into()])
    }

    // ── 合并 / 对齐 / 尺寸 ──

    pub fn merge(&self) -> Result<()> {
        self.raw.call("Merge", vec![]).map(|_| ())
    }

    pub fn unmerge(&self) -> Result<()> {
        self.raw.call("UnMerge", vec![]).map(|_| ())
    }

    /// 是否合并区域（区域内混合状态返回 false）。
    pub fn is_merged(&self) -> Result<bool> {
        Ok(matches!(self.raw.get("MergeCells")?.value(), CellValue::Bool(true)))
    }

    /// 设置对齐与自动换行（None 表示不改该项）。
    pub fn set_alignment(&self, h: Option<AlignH>, v: Option<AlignV>, wrap: Option<bool>) -> Result<()> {
        if let Some(h) = h {
            self.raw.put("HorizontalAlignment", vec![(h as i32).into()])?;
        }
        if let Some(v) = v {
            self.raw.put("VerticalAlignment", vec![(v as i32).into()])?;
        }
        if let Some(w) = wrap {
            self.raw.put("WrapText", vec![w.into()])?;
        }
        Ok(())
    }

    pub fn set_col_width(&self, w: f64) -> Result<()> {
        self.raw.put("ColumnWidth", vec![w.into()])
    }

    pub fn col_width(&self) -> Result<f64> {
        self.raw.get("ColumnWidth")?.float()
    }

    pub fn set_row_height(&self, h: f64) -> Result<()> {
        self.raw.put("RowHeight", vec![h.into()])
    }

    pub fn row_height(&self) -> Result<f64> {
        self.raw.get("RowHeight")?.float()
    }

    /// 自动调整列宽与行高。
    pub fn auto_fit(&self) -> Result<()> {
        self.raw.get("Columns")?.idispatch()?.call("AutoFit", vec![])?;
        self.raw.get("Rows")?.idispatch()?.call("AutoFit", vec![]).map(|_| ())
    }

    // ── 样式（联动入口：拿到 Font/Interior/Borders 后可继续设置各自属性） ──

    /// 字体对象（随后 `set_bold` / `set_size` / `set_color` 等）。
    pub fn font(&self) -> Result<Font> {
        Ok(Font::from_raw(self.raw.get_obj("Font", "Font")?))
    }

    /// 填充对象。
    pub fn interior(&self) -> Result<Interior> {
        Ok(Interior::from_raw(self.raw.get_obj("Interior", "Interior")?))
    }

    /// 边框对象。
    pub fn borders(&self) -> Result<Borders> {
        Ok(Borders::from_raw(self.raw.get_obj("Borders", "Borders")?))
    }

    /// 一次性设置字体（None 表示不改该项）。颜色为 OLE 格式（用 `rgb()` 助手）。
    pub fn set_font(
        &self,
        bold: Option<bool>,
        italic: Option<bool>,
        size: Option<f64>,
        color: Option<u32>,
        name: Option<&str>,
    ) -> Result<()> {
        let f = self.font()?;
        if let Some(v) = bold {
            f.set_bold(v)?;
        }
        if let Some(v) = italic {
            f.set_italic(v)?;
        }
        if let Some(v) = size {
            f.set_size(v)?;
        }
        if let Some(v) = color {
            f.set_color(v)?;
        }
        if let Some(v) = name {
            f.set_name(v)?;
        }
        Ok(())
    }

    // ── 清理 / 复制 ──

    /// 清除全部（内容+格式）。
    pub fn clear(&self) -> Result<()> {
        self.raw.call("Clear", vec![]).map(|_| ())
    }

    /// 只清内容。
    pub fn clear_contents(&self) -> Result<()> {
        self.raw.call("ClearContents", vec![]).map(|_| ())
    }

    /// 只清格式。
    pub fn clear_formats(&self) -> Result<()> {
        self.raw.call("ClearFormats", vec![]).map(|_| ())
    }

    /// 复制到目标区域（尺寸不一致时按源区域尺寸粘贴）。
    pub fn copy_to(&self, dst: &Range) -> Result<()> {
        self.raw
            .call("Copy", vec![ComValue::Obj(dst.raw().disp.clone())])
            .map(|_| ())
    }
}
