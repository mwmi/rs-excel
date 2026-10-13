//! 操作历史记录与撤销（Undo）。
//!
//! **为什么需要本模块**：Excel/WPS 的撤销栈只记录用户界面操作，经 COM 自动化做的
//! 修改不入栈，`Application.Undo` 无法撤销程序化改动。因此撤销由库自行记录快照：
//! 每次修改**前**对目标区域做快照，`undo` 按后进先出写回原地址。
//!
//! ```no_run
//! use rs_excel::prelude::*;
//!
//! let mut hist = History::new();
//! let sheet = wb.active_sheet()?;
//! hist.record(&sheet, "A1:C5", "tset A1")?;   // 修改前记录
//! sheet.write_tsv("A1", "姓名\t数量\n苹果\t12")?;
//! let undone = hist.undo(&wb)?;                // 返回 "已撤销: ..." 单行描述
//! ```
//!
//! ## 覆盖范围（AI 注意）
//!
//! - 只能撤销**经 record 系列记录**的操作（内容快照恢复值+公式，格式快照恢复格式）；
//! - 形状/图表/页面设置/删除工作表不在撤销范围；
//! - 类数字文本（如 "001"）写回时会被引擎重定型为数字；
//! - 行列插入/删除经整表内容快照恢复值/公式，被移动区域的格式不还原；
//! - `clear formats` 撤销后边框不还原（边框不入格式快照）。

use anyhow::{bail, Result};
use std::cell::RefCell;

use crate::com::variant::CellValue;
use crate::error::ExcelError;
use crate::range::Range;
use crate::worksheet::{num_to_col, parse_addr, parse_tsv, Worksheet};
use crate::workbook::Workbook;

/// 内容快照：某区域修改前的值+公式网格。
struct CellsSnapshot {
    sheet: String,
    addr: String,
    grid: Vec<Vec<CellValue>>,
    merged: bool,
}

/// 格式快照：`None` = 混合/未知（COM 返回 Null），撤销时跳过不还原。
struct FormatSnapshot {
    sheet: String,
    addr: String,
    number_format: Option<String>,
    h_align: Option<i32>,
    v_align: Option<i32>,
    wrap: Option<bool>,
    indent: Option<i32>,
    orientation: Option<i32>,
    shrink: Option<bool>,
    col_width: Option<f64>,
    row_height: Option<f64>,
    font_bold: Option<bool>,
    font_italic: Option<bool>,
    font_underline: Option<bool>,
    font_size: Option<f64>,
    font_color: Option<u32>,
    font_name: Option<String>,
    interior_color: Option<u32>,
}

/// 单条撤销记录的逆向操作。
enum UndoOp {
    Cells(CellsSnapshot),
    Formats(FormatSnapshot),
    Both(CellsSnapshot, FormatSnapshot),
    /// 撤销 = 删除该工作表
    SheetAdded { name: String },
    /// 撤销 = 改回旧名
    SheetRenamed { old: String, new: String },
}

/// 一条历史记录（含操作描述）。
pub struct UndoEntry {
    desc: String,
    op: UndoOp,
}

impl std::fmt::Display for UndoEntry {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.desc)?;
        match &self.op {
            UndoOp::Cells(c) => write!(f, " @ {}!{}", c.sheet, c.addr),
            UndoOp::Formats(x) => write!(f, " @ {}!{}", x.sheet, x.addr),
            UndoOp::Both(c, _) => write!(f, " @ {}!{}", c.sheet, c.addr),
            UndoOp::SheetAdded { name } | UndoOp::SheetRenamed { new: name, .. } => {
                write!(f, "（工作表 {name}）")
            }
        }
    }
}

/// 操作历史栈（后进先出撤销）。默认上限 50 条，超出丢弃最旧。
/// 内部 `RefCell`：record/undo 均只需 `&self`，调用方无需 mut 绑定。
pub struct History {
    entries: RefCell<Vec<UndoEntry>>,
    limit: usize,
}

impl Default for History {
    fn default() -> Self {
        Self::new()
    }
}

impl History {
    pub fn new() -> Self {
        Self::with_limit(50)
    }

    pub fn with_limit(limit: usize) -> Self {
        Self {
            entries: RefCell::new(Vec::new()),
            limit: limit.max(1),
        }
    }

    pub fn len(&self) -> usize {
        self.entries.borrow().len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.borrow().is_empty()
    }

    /// 全部记录摘要（旧→新），如 `#1 tset A1:C5 @ Sheet1!A1:C5`。
    pub fn summary(&self) -> Vec<String> {
        self.entries
            .borrow()
            .iter()
            .enumerate()
            .map(|(i, e)| format!("#{} {}", i + 1, e))
            .collect()
    }

    /// 清空历史。
    pub fn clear(&self) {
        self.entries.borrow_mut().clear();
    }

    // ── 记录（均在修改操作执行前调用） ──

    /// 记录内容快照（值+公式+合并状态）。
    pub fn record(&self, ws: &Worksheet, addr: &str, desc: impl Into<String>) -> Result<()> {
        let snap = snap_cells(ws, addr)?;
        self.push(desc.into(), UndoOp::Cells(snap));
        Ok(())
    }

    /// 记录格式快照（数字格式/对齐/字体/填充/行高列宽）。
    pub fn record_fmt(&self, ws: &Worksheet, addr: &str, desc: impl Into<String>) -> Result<()> {
        let snap = snap_formats(ws, addr)?;
        self.push(desc.into(), UndoOp::Formats(snap));
        Ok(())
    }

    /// 记录内容+格式快照（clear all / 合并类操作用）。
    pub fn record_all(&self, ws: &Worksheet, addr: &str, desc: impl Into<String>) -> Result<()> {
        let cells = snap_cells(ws, addr)?;
        let fmt = snap_formats(ws, addr)?;
        self.push(desc.into(), UndoOp::Both(cells, fmt));
        Ok(())
    }

    /// 记录已使用区域的内容快照（整表替换/行列增删类操作用）。
    pub fn record_used(&self, ws: &Worksheet, desc: impl Into<String>) -> Result<()> {
        let addr = ws.used_addr()?;
        let snap = snap_cells(ws, &addr)?;
        self.push(desc.into(), UndoOp::Cells(snap));
        Ok(())
    }

    /// 按 TSV 形状记录目标块内容快照（`write_tsv(anchor, tsv)` 前调用）。
    pub fn record_tsv(
        &self,
        ws: &Worksheet,
        anchor: &str,
        tsv: &str,
        desc: impl Into<String>,
    ) -> Result<()> {
        let grid = parse_tsv(tsv);
        if grid.is_empty() {
            bail!("record_tsv：没有可写入的数据");
        }
        let (c0, r0) = parse_addr(anchor).ok_or_else(|| {
            ExcelError::new("record_tsv", anchor, "锚点地址无法解析").with_hint("锚点格式如 A1")
        })?;
        let end = format!(
            "{}{}",
            num_to_col(c0 + grid[0].len() as i32 - 1),
            r0 + grid.len() as i32 - 1
        );
        self.record(ws, &format!("{anchor}:{end}"), desc)
    }

    /// 记录"新增工作表"（撤销 = 删除该表）。
    pub fn record_sheet_added(&self, name: impl Into<String>, desc: impl Into<String>) {
        self.push(desc.into(), UndoOp::SheetAdded { name: name.into() });
    }

    /// 记录"重命名工作表"（撤销 = 改回旧名）。
    pub fn record_sheet_renamed(
        &self,
        old: impl Into<String>,
        new: impl Into<String>,
        desc: impl Into<String>,
    ) {
        self.push(
            desc.into(),
            UndoOp::SheetRenamed {
                old: old.into(),
                new: new.into(),
            },
        );
    }

    // ── 撤销 ──

    /// 撤销最近一条记录的操作，返回单行描述（如 `tset A1:C5 @ Sheet1!A1:C5`）。
    pub fn undo(&self, wb: &Workbook) -> Result<String> {
        let Some(entry) = self.entries.borrow_mut().pop() else {
            return Err(ExcelError::new("undo", "history", "历史记录为空，没有可撤销的操作")
                .with_hint("仅记录过的操作可撤销（修改类操作需先 record）；用 summary() 查看历史")
                .into());
        };
        let desc = entry.to_string();
        restore(wb, entry.op)?;
        Ok(desc)
    }

    fn push(&self, desc: String, op: UndoOp) {
        let mut entries = self.entries.borrow_mut();
        while entries.len() >= self.limit {
            entries.remove(0);
        }
        entries.push(UndoEntry { desc, op });
    }
}

/// 内容快照：读公式网格（常量格得到原类型值，公式格得到 `=...`）+ 合并状态。
fn snap_cells(ws: &Worksheet, addr: &str) -> Result<CellsSnapshot> {
    let rng = ws.range(addr)?;
    Ok(CellsSnapshot {
        sheet: ws.name()?,
        addr: addr.to_string(),
        grid: rng.formulas()?,
        merged: rng.is_merged()?,
    })
}

/// 格式快照：逐项 `.ok()` 读取，混合态（Null）自动记 None。
fn snap_formats(ws: &Worksheet, addr: &str) -> Result<FormatSnapshot> {
    let rng = ws.range(addr)?;
    let g = |m: &str| rng.raw().get(m).ok().map(|v| v.value());
    let i = |m: &str| rng.raw().get(m).ok().and_then(|v| v.int().ok());
    let b = |m: &str| rng.raw().get(m).ok().and_then(|v| v.bool().ok());
    let f = rng.font().ok();
    let text = |v: Option<CellValue>| match v {
        Some(CellValue::Text(s)) => Some(s),
        _ => None,
    };
    Ok(FormatSnapshot {
        sheet: ws.name()?,
        addr: addr.to_string(),
        number_format: text(g("NumberFormat")),
        h_align: i("HorizontalAlignment"),
        v_align: i("VerticalAlignment"),
        wrap: b("WrapText"),
        indent: i("IndentLevel"),
        orientation: i("Orientation"),
        shrink: b("ShrinkToFit"),
        col_width: rng.raw().get("ColumnWidth").ok().and_then(|v| v.float().ok()),
        row_height: rng.raw().get("RowHeight").ok().and_then(|v| v.float().ok()),
        font_bold: f.as_ref().and_then(|f| f.bold().ok()),
        font_italic: f.as_ref().and_then(|f| f.italic().ok()),
        font_underline: f.as_ref().and_then(|f| f.underline().ok()),
        font_size: f.as_ref().and_then(|f| f.size().ok()),
        font_color: f.as_ref().and_then(|f| f.color().ok()),
        font_name: f.as_ref().and_then(|f| f.name().ok()),
        interior_color: rng.interior().ok().and_then(|x| x.color().ok()),
    })
}

/// 按类型执行逆向恢复。
fn restore(wb: &Workbook, op: UndoOp) -> Result<()> {
    match op {
        UndoOp::Cells(s) => {
            let ws = target_sheet(wb, &s.sheet)?;
            let rng = ws.range(&s.addr)?;
            if rng.is_merged()? && !s.merged {
                rng.unmerge()?;
            }
            write_back(&ws, &rng, &s)?;
            if s.merged && !rng.is_merged()? {
                rng.merge()?;
            }
            Ok(())
        }
        UndoOp::Formats(s) => {
            let rng = target_range(wb, &s.sheet, &s.addr)?;
            apply_formats(&rng, &s)
        }
        UndoOp::Both(c, f) => {
            let ws = target_sheet(wb, &c.sheet)?;
            let rng = ws.range(&c.addr)?;
            if rng.is_merged()? && !c.merged {
                rng.unmerge()?;
            }
            write_back(&ws, &rng, &c)?;
            if c.merged && !rng.is_merged()? {
                rng.merge()?;
            }
            apply_formats(&rng, &f)
        }
        UndoOp::SheetAdded { name } => {
            let ws = target_sheet(wb, &name)?;
            // 删表可能弹确认框：临时关告警，结束恢复
            let app = wb.raw().get("Application")?.idispatch()?;
            let prev = app.get("DisplayAlerts").ok().and_then(|v| v.bool().ok());
            let _ = app.put("DisplayAlerts", vec![false.into()]);
            let res = ws.delete();
            if let Some(p) = prev {
                let _ = app.put("DisplayAlerts", vec![p.into()]);
            }
            res
        }
        UndoOp::SheetRenamed { old, new } => target_sheet(wb, &new)?.rename(&old),
    }
}

/// 快照网格写回：批量写值，随后对 `=` 开头的文本格逐格强制为公式
/// （修正批量写可能被引擎当文本的边缘情况）。
fn write_back(ws: &Worksheet, rng: &Range, s: &CellsSnapshot) -> Result<()> {
    rng.set_values(&s.grid)?;
    let (c0, r0) = match parse_addr(s.addr.split(':').next().unwrap_or(&s.addr)) {
        Some(p) => p,
        None => return Ok(()), // 命名区域等无法定位基点：公式已随批量写入
    };
    for (r, row) in s.grid.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            if let CellValue::Text(f) = v {
                if f.starts_with('=') {
                    let addr = format!("{}{}", num_to_col(c0 + c as i32), r0 + r as i32);
                    ws.range(&addr)?.set_formula(f)?;
                }
            }
        }
    }
    Ok(())
}

/// 格式快照逐项还原（None 跳过）。
fn apply_formats(rng: &Range, s: &FormatSnapshot) -> Result<()> {
    let raw = rng.raw();
    if let Some(v) = &s.number_format {
        raw.put("NumberFormat", vec![v.clone().into()])?;
    }
    if let Some(v) = s.h_align {
        raw.put("HorizontalAlignment", vec![v.into()])?;
    }
    if let Some(v) = s.v_align {
        raw.put("VerticalAlignment", vec![v.into()])?;
    }
    if let Some(v) = s.wrap {
        raw.put("WrapText", vec![v.into()])?;
    }
    if let Some(v) = s.indent {
        raw.put("IndentLevel", vec![v.into()])?;
    }
    if let Some(v) = s.orientation {
        raw.put("Orientation", vec![v.into()])?;
    }
    if let Some(v) = s.shrink {
        raw.put("ShrinkToFit", vec![v.into()])?;
    }
    if let Some(v) = s.col_width {
        raw.put("ColumnWidth", vec![v.into()])?;
    }
    if let Some(v) = s.row_height {
        raw.put("RowHeight", vec![v.into()])?;
    }
    let f = rng.font()?;
    if let Some(v) = s.font_bold {
        f.set_bold(v)?;
    }
    if let Some(v) = s.font_italic {
        f.set_italic(v)?;
    }
    if let Some(v) = s.font_underline {
        f.set_underline(v)?;
    }
    if let Some(v) = s.font_size {
        f.set_size(v)?;
    }
    if let Some(v) = s.font_color {
        f.set_color(v)?;
    }
    if let Some(v) = &s.font_name {
        f.set_name(v)?;
    }
    if let Some(v) = s.interior_color {
        rng.interior()?.set_color(v)?;
    }
    Ok(())
}

fn target_sheet(wb: &Workbook, name: &str) -> Result<Worksheet> {
    wb.sheet(name).map_err(|e| {
        ExcelError::hint(
            e,
            format!("工作表 '{name}' 不存在：可能已被删除或改名，先用 sheets() 核对"),
        )
    })
}

fn target_range(wb: &Workbook, sheet: &str, addr: &str) -> Result<Range> {
    target_sheet(wb, sheet)?.range(addr)
}
