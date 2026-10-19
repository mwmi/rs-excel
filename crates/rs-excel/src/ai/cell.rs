//! 单元格输入：二维数组的「自然类型」契约与日期换算。
//!
//! | JSON 写法 | 本类型 | 写入语义 |
//! |---|---|---|
//! | `null` | `Empty` | 清空该格 |
//! | `"001"` | `Text` | 原样文本（不转数字、不丢前导零） |
//! | `12` / `1.5` | `Number` | 数值 |
//! | `true` / `false` | `Bool` | 布尔 |
//! | `"=SUM(A1:A3)"` | `Formula` | 公式（`=` 开头） |
//! | `{"f":"=..."}` | `Formula` | 强制公式（普通文本以 `=` 开头时用 `{"t":...}`） |
//! | `{"t":"..."}` | `Text` | 强制文本 |
//! | `"2026-01-05"` | `Date` | 真日期（序列值 + 日期格式） |
//!
//! 日期判定与序列换算都在这里完成，调用方不需要自己做任何类型推断。

use crate::com::variant::CellValue;

/// 写入单元格的自然类型。
#[derive(Debug, Clone, PartialEq)]
pub enum CellInput {
    Empty,
    Text(String),
    Number(f64),
    Bool(bool),
    Formula(String),
    /// ISO 日期 / 日期时间（`2026-01-05`、`2026-01-05 09:30:00`）。
    Date(String),
}

impl CellInput {
    /// 转成 `Value2` 通道的值（日期在 `write_grid` 里另行处理）。
    pub fn to_value(&self) -> CellValue {
        match self {
            Self::Empty => CellValue::Empty,
            Self::Text(text) | Self::Date(text) => CellValue::Text(text.clone()),
            Self::Number(number) => CellValue::Number(*number),
            Self::Bool(flag) => CellValue::Bool(*flag),
            // `Value2` 对 `=` 开头的字符串按公式解释（与引擎手动输入一致）
            Self::Formula(formula) => CellValue::Text(formula.clone()),
        }
    }

    pub fn is_date(&self) -> bool {
        matches!(self, Self::Date(_))
    }
}

/// 单次写入的单元格上限：超过说明该换一种做法（分块 / 只写汇总）。
pub const MAX_WRITE_CELLS: usize = 200_000;

/// 校验网格：非空、矩形、行宽一致、数值有限、总量不超限。
/// 返回 `(行数, 列数)`。任何一条不满足都在触碰 COM 之前失败。
pub fn validate_grid(grid: &[Vec<CellInput>]) -> Result<(usize, usize), crate::ai::error::ToolError> {
    use crate::ai::error::ToolError;

    if grid.is_empty() || grid[0].is_empty() {
        return Err(ToolError::validation(
            "EMPTY_GRID",
            "写入数据为空",
            "data 至少给一行一列，如 [[\"名称\",\"数量\"]]",
        ));
    }
    let cols = grid[0].len();
    if let Some((index, row)) = grid.iter().enumerate().find(|(_, row)| row.len() != cols) {
        return Err(ToolError::validation(
            "RAGGED_GRID",
            format!(
                "data 第 {} 行有 {} 列，第 1 行有 {cols} 列",
                index + 1,
                row.len()
            ),
            "所有行的列数必须一致；缺的格写 null",
        ));
    }
    let cells = grid.len() * cols;
    if cells > MAX_WRITE_CELLS {
        return Err(ToolError::validation(
            "LIMIT_EXCEEDED",
            format!("一次写入 {cells} 格，超过上限 {MAX_WRITE_CELLS}"),
            "拆成多次写入，或只写需要的区域",
        ));
    }
    for (r, row) in grid.iter().enumerate() {
        for (c, cell) in row.iter().enumerate() {
            match cell {
                CellInput::Number(number) if !number.is_finite() => {
                    return Err(ToolError::validation(
                        "BAD_NUMBER",
                        format!("data[{}][{}] 不是有限数值", r + 1, c + 1),
                        "写普通数字，或用 {\"t\":\"...\"} 强制文本",
                    ));
                }
                CellInput::Date(text) if serial_from_iso(text).is_none() => {
                    return Err(ToolError::validation(
                        "BAD_DATE",
                        format!("data[{}][{}] 日期不合法：「{text}」", r + 1, c + 1),
                        "日期写 2026-01-05 或 2026-01-05 09:30:00；要原样文本用 {\"t\":\"...\"}",
                    ));
                }
                _ => {}
            }
        }
    }
    Ok((grid.len(), cols))
}

/// 是否 `YYYY-MM-DD`（可带 ` HH:MM[:SS]`）形态。
pub fn is_iso_date(text: &str) -> bool {
    let text = text.trim();
    let bytes = text.as_bytes();
    bytes.len() >= 10
        && bytes[..4].iter().all(u8::is_ascii_digit)
        && bytes[4] == b'-'
        && bytes[5..7].iter().all(u8::is_ascii_digit)
        && bytes[7] == b'-'
        && bytes[8..10].iter().all(u8::is_ascii_digit)
        && serial_from_iso(text).is_some()
}

/// ISO 日期 → Excel 序列值（1899-12-30 纪元；日期时间带小数部分）。
pub fn serial_from_iso(text: &str) -> Option<f64> {
    let text = text.trim();
    let split = text.find(|c: char| c == 'T' || c == ' ').unwrap_or(text.len());
    let (date, time) = text.split_at(split);
    let mut parts = date.split('-');
    let year: i64 = parts.next()?.parse().ok()?;
    let month: u32 = parts.next()?.parse().ok()?;
    let day: u32 = parts.next()?.parse().ok()?;
    if parts.next().is_some() || !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }
    let mut serial = (days_from_civil(year, month, day) + 25_569) as f64;
    let time = time.trim().trim_start_matches(['T', ' ']).trim();
    if !time.is_empty() {
        let mut pieces = time.split(':');
        let hour: f64 = pieces.next()?.trim().parse().ok()?;
        let minute: f64 = pieces.next().unwrap_or("0").trim().parse().ok()?;
        let second: f64 = pieces.next().unwrap_or("0").trim().parse().ok()?;
        if !(0.0..24.0).contains(&hour) || !(0.0..60.0).contains(&minute) || !(0.0..60.0).contains(&second) {
            return None;
        }
        serial += (hour * 3600.0 + minute * 60.0 + second) / 86_400.0;
    }
    Some(serial)
}

/// 公历 → 1970-01-01 起的天数（Howard Hinnant 的 `days_from_civil`）。
fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = if month <= 2 { year - 1 } else { year };
    let era = if year >= 0 { year } else { year - 399 } / 400;
    let yoe = year - era * 400;
    let mp = if month > 2 { month - 3 } else { month + 9 } as i64;
    let doy = (153 * mp + 2) / 5 + day as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}
