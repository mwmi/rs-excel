//! 写区域：自然类型网格 → 一次 `Value2` 批量写；日期走一次 `NumberFormat` 数组补齐。
//!
//! 旧实现（调用方自己补日期）对每个日期格都要 offset + resize + 读格式 + 写值 + 写格式，
//! 200 行日期 ≈ 800+ 次 COM。这里改成：
//! 1. 有日期格时读一次整块 `NumberFormat`（1 次 COM）；
//! 2. 日期格按格式决定写序列值还是文本；
//! 3. 整块 `Value2` 写一次（1 次 COM）；
//! 4. 需要改格式时整块 `NumberFormat` 写一次（1 次 COM）。

use crate::ai::cell::{serial_from_iso, validate_grid, CellInput};
use crate::ai::error::{AiResult, ToolError};
use crate::ai::target::clean_addr;
use crate::com::safearray::{rows_to_variant, variant_to_rows};
use crate::com::variant::CellValue;
use crate::range::Range;

/// 实际写入路径。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WritePath {
    /// 一次 SAFEARRAY 批量写。
    SafeArray,
    /// SAFEARRAY 构造失败后逐格写（慢路径）。
    Cell,
}

impl WritePath {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::SafeArray => "safearray",
            Self::Cell => "cell",
        }
    }
}

/// 写入报告（`Range::set_values` 返回，`write_grid` 补充地址与日期信息）。
#[derive(Debug, Clone)]
pub struct WriteReport {
    pub addr: String,
    pub rows: i32,
    pub cols: i32,
    pub cells: i64,
    pub path: WritePath,
    pub fallback: bool,
    /// 按日期写入的格数。
    pub dates: i32,
    /// 非致命问题（如日期格式没补上），回执里原样转告模型。
    pub warning: Option<String>,
}

impl WriteReport {
    pub fn new(rows: i32, cols: i32, path: WritePath) -> Self {
        Self {
            addr: String::new(),
            rows,
            cols,
            cells: rows as i64 * cols as i64,
            path,
            fallback: path == WritePath::Cell,
            dates: 0,
            warning: None,
        }
    }
}

/// 按自然类型写一整块区域。
pub fn write_grid(range: &Range, grid: &[Vec<CellInput>]) -> AiResult<WriteReport> {
    let (rows, cols) = validate_grid(grid)?;
    let has_date = grid.iter().flatten().any(CellInput::is_date);
    let mut dates = 0;
    let mut warning = None;

    let (values, formats) = if has_date {
        match number_formats(range, rows, cols) {
            Ok(current) => {
                let mut values = Vec::with_capacity(rows);
                let mut formats_out = Vec::with_capacity(rows);
                let mut changed = false;
                for (r, row) in grid.iter().enumerate() {
                    let mut value_row = Vec::with_capacity(cols);
                    let mut format_row = Vec::with_capacity(cols);
                    for (c, cell) in row.iter().enumerate() {
                        let format = &current[r][c];
                        match cell {
                            CellInput::Date(text) => {
                                dates += 1;
                                if format.contains('@') {
                                    // 文本格式的格写文本，避免日期序列号裸奔
                                    value_row.push(CellValue::Text(text.clone()));
                                    format_row.push(CellValue::Text(format.clone()));
                                } else {
                                    let serial = serial_from_iso(text).ok_or_else(|| {
                                        ToolError::validation(
                                            "BAD_DATE",
                                            format!("日期不合法：「{text}」"),
                                            "写 2026-01-05 或 2026-01-05 09:30:00",
                                        )
                                    })?;
                                    value_row.push(CellValue::Number(serial));
                                    if is_date_format(format) {
                                        format_row.push(CellValue::Text(format.clone()));
                                    } else {
                                        format_row.push(CellValue::Text("yyyy-mm-dd".into()));
                                        changed = true;
                                    }
                                }
                            }
                            other => {
                                value_row.push(other.to_value());
                                format_row.push(CellValue::Text(format.clone()));
                            }
                        }
                    }
                    values.push(value_row);
                    formats_out.push(format_row);
                }
                (values, changed.then_some(formats_out))
            }
            Err(_) => {
                // 读不到区域格式：日期按文本写（宁可显示正确，也不让序列号裸奔）
                warning = Some("读不到区域数字格式，日期已按文本写入；需要真日期请先 format 设 yyyy-mm-dd".into());
                let values = grid
                    .iter()
                    .map(|row| {
                        row.iter()
                            .map(|cell| match cell {
                                CellInput::Date(text) => {
                                    dates += 1;
                                    CellValue::Text(text.clone())
                                }
                                other => other.to_value(),
                            })
                            .collect()
                    })
                    .collect();
                (values, None)
            }
        }
    } else {
        (
            grid.iter()
                .map(|row| row.iter().map(CellInput::to_value).collect())
                .collect(),
            None,
        )
    };

    let mut report = range.set_values(&values).map_err(ToolError::from)?;
    report.addr = clean_addr(&range.address().unwrap_or_default());
    report.dates = dates;
    report.warning = warning;

    if let Some(formats) = formats {
        if let Err(e) = write_number_formats(range, &formats) {
            report.warning = Some(format!(
                "值已写入，但日期格式没补上：{e}。可再发一条 format num=\"yyyy-mm-dd\" 修这一列"
            ));
        }
    }
    Ok(report)
}

/// 读整块数字格式（单格返回会扩展成 rows × cols）。
fn number_formats(range: &Range, rows: usize, cols: usize) -> AiResult<Vec<Vec<String>>> {
    let result = range.raw().get("NumberFormat").map_err(ToolError::from)?;
    let grid = variant_to_rows(&result.v).map_err(ToolError::from)?;
    let mut out: Vec<Vec<String>> = grid
        .iter()
        .map(|row| row.iter().map(|cell| cell.to_string()).collect())
        .collect();
    if out.len() == 1 && out[0].len() == 1 && (rows > 1 || cols > 1) {
        out = vec![vec![out[0][0].clone(); cols]; rows];
    }
    if out.len() != rows || out.first().map_or(0, Vec::len) != cols {
        return Err(ToolError::validation(
            "FORMAT_READ_FAILED",
            format!(
                "区域数字格式是 {}×{}，与数据 {rows}×{cols} 不一致",
                out.len(),
                out.first().map_or(0, Vec::len)
            ),
            "把写入区域收缩成与数据相同的形状",
        ));
    }
    Ok(out)
}

/// 一次写整块数字格式（`NumberFormat` 接受二维数组）。
fn write_number_formats(range: &Range, formats: &[Vec<CellValue>]) -> AiResult<()> {
    let variant = rows_to_variant(formats).map_err(ToolError::from)?;
    range
        .raw()
        .put_raw("NumberFormat", variant)
        .map_err(ToolError::from)
}

/// 是否已经带日期/时间格式（含 `y` / `d` / `h` 的非文本格式）。
fn is_date_format(format: &str) -> bool {
    let mut cleaned = String::with_capacity(format.len());
    let mut in_quote = false;
    let mut escaped = false;
    for ch in format.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match ch {
            '\\' => escaped = true,
            '"' => in_quote = !in_quote,
            _ if in_quote => {}
            _ => cleaned.push(ch.to_ascii_lowercase()),
        }
    }
    cleaned.contains(['y', 'd', 'h'])
}
