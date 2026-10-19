//! 读区域：预算内的二维网格 + 总量/截断信息。
//!
//! 与旧 `read_tsv` 的区别：
//! - 返回矩阵元信息（`total_rows` / `total_cols` / `truncated`），调用方不会悄悄丢数据；
//! - `Values` 走 `Value` 兼容通道（日期为 `2026-01-05` 文本、货币为数字、错误值为 `#DIV/0!`）；
//! - `Formulas` 走 `Formula2`（引擎不支持自动回退 `Formula`）。

use crate::ai::error::{AiResult, ToolError};
use crate::ai::target::{clean_addr, Target};
use crate::com::variant::CellValue;
use crate::range::Range;

/// 读什么。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Projection {
    /// 值（日期转文本、公式格给计算结果）。
    Values,
    /// 显示文本：单格走 `Text`，多格退回值与 [`Projection::Values`] 一致。
    Text,
    /// 公式：常量格返回原值，公式格返回 `=...` 文本。
    Formulas,
}

/// 单次读取窗口。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReadWindow {
    pub max_rows: i32,
    pub max_cols: i32,
}

impl Default for ReadWindow {
    fn default() -> Self {
        Self {
            max_rows: 200,
            max_cols: 30,
        }
    }
}

impl ReadWindow {
    pub fn new(max_rows: i32, max_cols: i32) -> Self {
        Self {
            max_rows: max_rows.clamp(1, 2000),
            max_cols: max_cols.clamp(1, 200),
        }
    }
}

/// 读回的矩阵与元信息。
#[derive(Debug, Clone)]
pub struct Matrix {
    /// 实际返回的区域地址（无 `$`）。
    pub addr: String,
    pub rows: i32,
    pub cols: i32,
    /// 目标区域（未截断前）的总行数。
    pub total_rows: i32,
    /// 目标区域（未截断前）的总列数。
    pub total_cols: i32,
    pub truncated: bool,
    pub cells: Vec<Vec<CellValue>>,
}

impl Matrix {
    pub fn read(range: &Range, window: ReadWindow) -> AiResult<Self> {
        Self::read_with(range, window, Projection::Values)
    }

    pub fn read_with(range: &Range, window: ReadWindow, projection: Projection) -> AiResult<Self> {
        // 先用 Address 推导行列（1 次 COM），推不出来（整列/整行）才退回 Rows/Columns.Count。
        // 旧实现固定打 4 次 Count + 1 次 Address，这里常见路径只剩 1 次。
        let addr = clean_addr(&range.address().unwrap_or_default());
        let known = Target {
            book: None,
            sheet: None,
            range: Some(addr.clone()),
        }
        .parse()
        .ok()
        .and_then(|parsed| Some((parsed.rows?, parsed.cols?)));
        let (total_rows, total_cols) = match known {
            Some((rows, cols)) => (rows, cols),
            None => (
                range.rows().map_err(ToolError::from)?,
                range.columns().map_err(ToolError::from)?,
            ),
        };
        if total_rows <= 0 || total_cols <= 0 {
            return Ok(Self {
                addr,
                rows: 0,
                cols: 0,
                total_rows: total_rows.max(0),
                total_cols: total_cols.max(0),
                truncated: false,
                cells: Vec::new(),
            });
        }
        let take_rows = total_rows.min(window.max_rows);
        let take_cols = total_cols.min(window.max_cols);
        let resized;
        let target = if take_rows == total_rows && take_cols == total_cols {
            range
        } else {
            resized = range.resize(take_rows, take_cols).map_err(ToolError::from)?;
            &resized
        };
        let cells = match projection {
            Projection::Formulas => target.formulas().map_err(ToolError::from)?,
            // `Range.Text` 只对单格有意义（多格在引擎上返回 Null），多格退回值通道
            Projection::Text if take_rows * take_cols == 1 => {
                vec![vec![CellValue::Text(
                    target.text().map_err(ToolError::from)?,
                )]]
            }
            _ => target.values_compat().map_err(ToolError::from)?,
        };
        let rows = cells.len() as i32;
        let cols = cells.first().map_or(0, Vec::len) as i32;
        Ok(Self {
            addr: clean_addr(&target.address().unwrap_or(addr)),
            rows,
            cols,
            total_rows,
            total_cols,
            truncated: take_rows < total_rows || take_cols < total_cols,
            cells,
        })
    }

    /// 空矩阵（区域没有内容时用）。
    pub fn is_empty(&self) -> bool {
        self.rows == 0 || self.cols == 0
    }
}
