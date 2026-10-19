//! 目标解析：`book` / `sheet` / `range` 三个显式字段。
//!
//! 规则（纯字符串解析，不碰 COM）：
//! - `range` 仍接受旧的 `"表名!区域"` 写法；显式给了 `sheet` 时不再按 `!` 拆；
//! - 去 `$`、去引号；裸列 `E` → `E:E`，裸行 `3` → `3:3`（否则 WPS 抛 0x80020009）；
//! - 工作表一律按**名称**选择，序号走独立的 [`Target::resolve_index`]；
//! - 解析结果带锚点与可推导的尺寸，供写入前校验矩阵形状。

use crate::ai::error::{AiResult, Category, ToolError};
use crate::app::App;
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;

/// 目标：工作簿可选（缺省 = 界面选中 / 活动工作簿），工作表与区域可选。
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Target {
    pub book: Option<String>,
    pub sheet: Option<String>,
    pub range: Option<String>,
}

/// `Target` 的纯解析结果。
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParsedTarget {
    pub book: Option<String>,
    pub sheet: Option<String>,
    /// 归一化后的区域地址；`None` = 由操作的默认语义决定（UsedRange / 选区）。
    pub addr: Option<String>,
    /// 区域左上角（写入锚点）。
    pub anchor: Option<String>,
    /// 可推导的总行数（整列 / 命名区域为 None）。
    pub rows: Option<i32>,
    /// 可推导的总列数（整行 / 命名区域为 None）。
    pub cols: Option<i32>,
}

/// 解析 + 绑定后的目标。
pub struct Resolved {
    pub ws: Worksheet,
    pub sheet: String,
    pub addr: Option<String>,
    pub anchor: Option<String>,
    pub rows: Option<i32>,
    pub cols: Option<i32>,
}

impl Target {
    pub fn new(book: Option<String>, sheet: Option<String>, range: Option<String>) -> Self {
        Self { book, sheet, range }
    }

    /// 只解析、不连接 COM。任一处非法立刻返回 `RANGE_INVALID`。
    pub fn parse(&self) -> AiResult<ParsedTarget> {
        let explicit = self
            .sheet
            .as_deref()
            .map(str::trim)
            .filter(|s| !s.is_empty())
            .map(clean_name);
        let raw = self.range.as_deref().unwrap_or("").trim();
        // 即使显式给了 sheet，也把 range 里的 `表名!` 拆出来核对：
        // 两处写成不同的表时直接报错，绝不按其中一个静默落点
        let (prefixed, rest) = split_sheet(raw);
        let sheet = match (explicit, prefixed) {
            (Some(a), Some(b)) if a != b => {
                return Err(ToolError::validation(
                    "RANGE_INVALID",
                    format!("sheet 写的是「{a}」，range 前缀是「{b}」"),
                    "两处只留一个，或改成同一张表",
                ))
            }
            (Some(a), _) => Some(a),
            (None, b) => b,
        };
        let (addr, anchor, rows, cols) = match rest.is_empty() {
            true => (None, None, None, None),
            false => {
                let (addr, anchor, rows, cols) = normalize_addr(&rest)?;
                (Some(addr), Some(anchor), rows, cols)
            }
        };
        Ok(ParsedTarget {
            book: self
                .book
                .as_deref()
                .map(str::trim)
                .filter(|s| !s.is_empty())
                .map(str::to_string),
            sheet,
            addr,
            anchor,
            rows,
            cols,
        })
    }

    /// 在指定工作簿上解析工作表（名称缺失时用 `default_sheet`，再退回活动表）。
    pub fn resolve(&self, book: &Workbook, default_sheet: Option<&str>) -> AiResult<Resolved> {
        let parsed = self.parse()?;
        let ws = match parsed.sheet.as_deref() {
            Some(name) => book.sheet(name).map_err(|e| sheet_error(e, book, name))?,
            None => {
                let mut found = None;
                if let Some(name) = default_sheet.map(str::trim).filter(|s| !s.is_empty()) {
                    found = book.sheet(name).ok();
                }
                match found {
                    Some(ws) => ws,
                    None => book.active_sheet().map_err(ToolError::from)?,
                }
            }
        };
        let sheet = ws.name().unwrap_or_default();
        Ok(Resolved {
            ws,
            sheet,
            addr: parsed.addr,
            anchor: parsed.anchor,
            rows: parsed.rows,
            cols: parsed.cols,
        })
    }

    /// 按序号选择工作表（只给 `sheet_switch` 之类明确使用序号的场景）。
    pub fn resolve_index(&self, book: &Workbook, index: i32) -> AiResult<Worksheet> {
        book.sheet_at(index).map_err(|e| {
            ToolError::from_anyhow(&e)
                .with_code("SHEET_NOT_FOUND")
                .with_category(Category::TargetNotFound)
                .with_target(index.to_string())
        })
    }
}

/// 选择工作簿：给了名字/完整路径就精确匹配，否则用活动工作簿。
pub fn pick_workbook(app: &App, name: Option<&str>) -> AiResult<Workbook> {
    match name.map(str::trim).filter(|s| !s.is_empty()) {
        None => app
            .active_workbook()
            .map_err(|e| ToolError::from_anyhow(&e).with_code("BOOK_NOT_FOUND")),
        Some(name) => {
            let list = app.workbooks().map_err(ToolError::from)?;
            match list
                .iter()
                .find(|info| info.name == name || info.full_name == name)
            {
                Some(info) => app.workbook(&info.name).map_err(ToolError::from),
                None => {
                    let names = list
                        .iter()
                        .map(|info| info.name.as_str())
                        .collect::<Vec<_>>()
                        .join("、");
                    Err(ToolError::new(
                        "BOOK_NOT_FOUND",
                        Category::TargetNotFound,
                        format!("工作簿「{name}」没有打开"),
                    )
                    .with_hint(if names.is_empty() {
                        "当前没有打开的工作簿，先让用户打开一个".to_string()
                    } else {
                        format!("已打开的工作簿：{names}")
                    }))
                }
            }
        }
    }
}

fn sheet_error(e: anyhow::Error, book: &Workbook, name: &str) -> ToolError {
    let names = book
        .sheets()
        .map(|list| {
            list.into_iter()
                .map(|info| info.name)
                .collect::<Vec<_>>()
                .join("、")
        })
        .unwrap_or_default();
    ToolError::from_anyhow(&e)
        .with_code("SHEET_NOT_FOUND")
        .with_category(Category::TargetNotFound)
        .with_target(name)
        .with_hint(if names.is_empty() {
            format!("工作表「{name}」不存在；先 state 看当前表名")
        } else {
            format!("工作表「{name}」不存在；现有：{names}")
        })
}

/// 去 `$` 与首尾空白（COM 的 Address 带 `$`，回执里不需要）。
pub fn clean_addr(addr: &str) -> String {
    addr.replace('$', "").trim().to_string()
}

fn clean_name(name: &str) -> String {
    name.trim().trim_matches('\'').replace("''", "'")
}

/// `"表名!A1"` → (`Some("表名")`, `"A1"`)；支持 `'含 ! 的表名'!A1`。
fn split_sheet(raw: &str) -> (Option<String>, String) {
    let raw = raw.trim();
    if let Some(body) = raw.strip_prefix('\'') {
        if let Some(end) = body.find('\'') {
            let name = clean_name(&body[..end]);
            let rest = body[end + 1..].trim_start();
            if let Some(rest) = rest.strip_prefix('!') {
                return (Some(name), rest.trim().to_string());
            }
        }
    }
    match raw.find('!') {
        Some(pos) => {
            let name = clean_name(&raw[..pos]);
            let name = (!name.is_empty()).then_some(name);
            (name, raw[pos + 1..].trim().to_string())
        }
        None => (None, raw.to_string()),
    }
}

/// 归一化地址 → `(addr, anchor, rows, cols)`。
fn normalize_addr(raw: &str) -> AiResult<(String, String, Option<i32>, Option<i32>)> {
    let addr = raw.replace('$', "");
    let addr = addr.trim();
    if addr.is_empty() {
        return Err(ToolError::validation(
            "RANGE_INVALID",
            "区域地址为空",
            "写 \"A1\"、「A1:C10」或省略（读整张已用区域）",
        ));
    }
    if addr.contains(['[', ']']) {
        return Err(ToolError::validation(
            "RANGE_INVALID",
            format!("不支持跨工作簿引用：「{addr}」"),
            "只写本工作簿内的地址，如 A1:C10",
        ));
    }
    if addr.contains([',', ';']) {
        return Err(ToolError::validation(
            "RANGE_INVALID",
            format!("不支持多区域引用：「{addr}」"),
            "一条 op 只作用一个连续区域；多块请拆成多条 op",
        ));
    }
    if addr.matches(':').count() > 1 {
        return Err(invalid_range(addr));
    }
    // 裸列 / 裸行
    if is_column(addr) {
        return Ok((format!("{addr}:{addr}"), format!("{addr}1"), None, Some(1)));
    }
    if is_row(addr) {
        return Ok((format!("{addr}:{addr}"), format!("A{addr}"), Some(1), None));
    }
    if let Some((left, right)) = addr.split_once(':') {
        let left = left.trim();
        let right = right.trim();
        if left.is_empty() || right.is_empty() {
            return Err(invalid_range(addr));
        }
        // 整列 A:C / 整行 3:5
        if is_column(left) && is_column(right) {
            let (l, r) = (col_index(left), col_index(right));
            let cols = match (l, r) {
                (Some(l), Some(r)) if r >= l => Some(r - l + 1),
                _ => return Err(invalid_range(addr)),
            };
            return Ok((format!("{left}:{right}"), format!("{left}1"), None, cols));
        }
        if is_row(left) && is_row(right) {
            let (l, r) = (left.parse::<i32>(), right.parse::<i32>());
            let rows = match (l, r) {
                (Ok(l), Ok(r)) if l >= 1 && r >= l => Some(r - l + 1),
                _ => return Err(invalid_range(addr)),
            };
            return Ok((format!("{left}:{right}"), format!("A{left}"), rows, None));
        }
        // 单元格矩形
        if let (Some((r1, c1)), Some((r2, c2))) = (parse_cell(left), parse_cell(right)) {
            if r2 < r1 || c2 < c1 {
                return Err(invalid_range(addr));
            }
            return Ok((
                addr.to_string(),
                format!("{}{}", col_letters(c1), r1),
                Some(r2 - r1 + 1),
                Some(c2 - c1 + 1),
            ));
        }
        // 混合边界（A1:C）或命名区域：尺寸未知，按原样交给 COM
        return Ok((addr.to_string(), anchor_of(addr).unwrap_or_else(|| "A1".into()), None, None));
    }
    // 单格 / 命名区域
    let anchor = anchor_of(addr).unwrap_or_else(|| addr.to_string());
    let dims = parse_cell(addr).map(|_| (Some(1), Some(1))).unwrap_or((None, None));
    Ok((addr.to_string(), anchor, dims.0, dims.1))
}

fn invalid_range(addr: &str) -> ToolError {
    ToolError::validation(
        "RANGE_INVALID",
        format!("区域地址不合法：「{addr}」"),
        "合法写法：A1、A1:C10、A:C、3:3、命名区域",
    )
}

fn anchor_of(addr: &str) -> Option<String> {
    let left = addr.split(':').next()?.trim();
    if let Some((row, col)) = parse_cell(left) {
        return Some(format!("{}{}", col_letters(col), row));
    }
    if is_column(left) {
        return Some(format!("{left}1"));
    }
    if is_row(left) {
        return Some(format!("A{left}"));
    }
    None
}

fn is_column(text: &str) -> bool {
    let bytes = text.as_bytes();
    (1..=3).contains(&bytes.len()) && bytes.iter().all(u8::is_ascii_alphabetic)
}

fn is_row(text: &str) -> bool {
    !text.is_empty() && text.len() <= 7 && text.bytes().all(|b| b.is_ascii_digit())
}

/// `B12` / `$B$12` → (12, 2)。
fn parse_cell(text: &str) -> Option<(i32, i32)> {
    let text = text.replace('$', "");
    let split = text.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = text.split_at(split);
    if letters.is_empty() || digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row = digits.parse::<i32>().ok()?;
    let col = col_index(letters)?;
    (row >= 1 && col >= 1).then_some((row, col))
}

/// 列字母 → 1 基列号（`A` → 1，`AB` → 28）。
pub fn col_index(letters: &str) -> Option<i32> {
    let mut value = 0i32;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        value = value
            .checked_mul(26)?
            .checked_add(ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1)?;
    }
    (value > 0).then_some(value)
}

/// 1 基列号 → 列字母（1 → `A`）。
pub fn col_letters(mut index: i32) -> String {
    if index < 1 {
        return "A".into();
    }
    let mut out = Vec::new();
    while index > 0 {
        let rem = (index - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        index = (index - 1) / 26;
    }
    out.iter().rev().collect()
}
