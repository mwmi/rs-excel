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
//! - 格式快照里，区域填充混杂时会逐格细读（小区域才做，见 [`FILL_CELL_LIMIT`]）；
//!   边框混杂时按「跳过」处理（不猜一个值去还原）。

use anyhow::{bail, Result};
use std::cell::RefCell;

use crate::com::variant::CellValue;
use crate::condfmt::CondSnap;
use crate::error::ExcelError;
use crate::range::Range;
use crate::style::EdgeStyle;
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
    /// 填充（底纹），见 [`FillSnap`]。
    fill: FillSnap,
    /// 六条边各自的边框状态，见 [`crate::style::Borders::snapshot`]；空 = 读不到边框对象。
    borders: Vec<Option<EdgeStyle>>,
}

/// 条件格式快照：某区域修改前的全部规则（空 = 当时没有规则，撤销即清空）。
struct CondSnapshot {
    sheet: String,
    addr: String,
    rules: Vec<CondSnap>,
}

/// 一次填充快照允许细分的格子上限：超过它（整列那种上百万格）不细读，只在整块一致时还原。
/// 最坏情况（每一行内部都混杂）要逐格读，按 WPS 每次 COM 往返约 1ms 估，400 格 ≈ 1.5s 封顶。
const FILL_CELL_LIMIT: i64 = 400;

/// 填充（底纹）快照。逐项语义同 [`crate::style::Interior::fill`] 的三态。
///
/// 为什么要三态：整块读一次只在填充一致时给得出值（`Pattern`），混杂时是 Null。
/// 而 AI 改底纹的区域常常正是混杂的（表头有底纹、表体没有），只记「混杂」就等于
/// 撤销时什么都不做 —— 新刷的底纹撤不掉（用户实报）。
enum FillSnap {
    /// 整块同一状态：`Some(色)` = 同色，`None` = 整块无填充。
    Uniform(Option<u32>),
    /// 整块混杂：逐格记（行优先）。每格同样是 `Some(色)` / `None`（无填充），
    /// 单格读不出来才是 `None` 外层（跳过该格）。`cols` 用于把下标还原成行列。
    Cells { cols: i32, states: Vec<Option<Option<u32>>> },
    /// 混杂且格子超过 [`FILL_CELL_LIMIT`]：不还原。
    Skip,
}

/// 单条撤销记录的逆向操作（= 应用它就把改动反做一遍）。
/// 每个变体都是可逆的：撤销时应用它，重做时应用它的反向（见 [`snapshot_like`]）。
enum UndoOp {
    Cells(CellsSnapshot),
    Formats(FormatSnapshot),
    Both(CellsSnapshot, FormatSnapshot),
    /// 条件格式规则集（撤销 = 写回这一组；恢复 ⇄ 撤销互换时重拍当前规则）。
    Conds(CondSnapshot),
    /// 应用到表上 = 删除该工作表（「新增工作表」的逆向）。
    SheetDelete { name: String },
    /// 应用到表上 = 新建一张空表（删表被撤销后再恢复用；内容由后续 Cells 记录补回）。
    SheetCreate { name: String },
    /// 应用到表上 = 把 `from` 改名为 `to`。
    SheetRename { from: String, to: String },
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
            UndoOp::Conds(c) => write!(f, " @ {}!{}", c.sheet, c.addr),
            UndoOp::SheetDelete { name } | UndoOp::SheetCreate { name } => write!(f, "（工作表 {name}）"),
            UndoOp::SheetRename { from, to } => write!(f, "（工作表 {from} → {to}）"),
        }
    }
}

/// 操作历史栈（后进先出撤销 + 恢复）。默认上限 50 条，超出丢弃最旧。
/// 内部 `RefCell`：record/undo/redo 均只需 `&self`，调用方无需 mut 绑定。
///
/// **撤销与恢复是一对快照的互换**：撤销时先把目标区域的「当前状态」拍成一条记录压入
/// [`redo`](Self::redo)，再写回旧快照；恢复则反过来。所以恢复不需要另存一份「操作日志」，
/// 两个栈里的条目是同一种东西，只是方向相反。
pub struct History {
    /// 撤销栈：**末尾** = 最近一次操作 = 下一个被撤销的
    entries: RefCell<Vec<UndoEntry>>,
    /// 恢复栈：按被撤销的先后压入（`entries` 弹出的顺序），**末尾** = 最近一次被撤销的
    redo: RefCell<Vec<UndoEntry>>,
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
            redo: RefCell::new(Vec::new()),
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
        summarize(&self.entries)
    }

    /// 可恢复记录的数量。
    pub fn redo_len(&self) -> usize {
        self.redo.borrow().len()
    }

    /// 可恢复记录摘要（旧→新，与 [`summary`](Self::summary) 同口径）。
    pub fn redo_summary(&self) -> Vec<String> {
        summarize(&self.redo)
    }

    /// 清空历史（撤销栈与恢复栈一起清）。
    pub fn clear(&self) {
        self.entries.borrow_mut().clear();
        self.redo.borrow_mut().clear();
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

    /// 记录条件格式快照（`cond_xxx` / `clear_cond` 前调用）。
    pub fn record_cond(&self, ws: &Worksheet, addr: &str, desc: impl Into<String>) -> Result<()> {
        let snap = snap_conds(ws, addr)?;
        self.push(desc.into(), UndoOp::Conds(snap));
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

    /// 记录"新增工作表"（撤销 = 删除该表，重做 = 重新建一张空表）。
    pub fn record_sheet_added(&self, name: impl Into<String>, desc: impl Into<String>) {
        self.push(desc.into(), UndoOp::SheetDelete { name: name.into() });
    }

    /// 记录"重命名工作表"（撤销 = 改回旧名，重做 = 再改成新名）。
    pub fn record_sheet_renamed(
        &self,
        old: impl Into<String>,
        new: impl Into<String>,
        desc: impl Into<String>,
    ) {
        // 记的是「从新名改回旧名」这一步：撤销直接应用它，重做应用它的反向
        self.push(
            desc.into(),
            UndoOp::SheetRename {
                from: new.into(),
                to: old.into(),
            },
        );
    }

    // ── 撤销 / 恢复 ──

    /// 撤销最近一条记录的操作，返回单行描述（如 `tset A1:C5 @ Sheet1!A1:C5`）。
    pub fn undo(&self, wb: &Workbook) -> Result<String> {
        self.swap(wb, &self.entries, &self.redo)?.ok_or_else(|| {
            ExcelError::new("undo", "history", "历史记录为空，没有可撤销的操作")
                .with_hint("仅记录过的操作可撤销（修改类操作需先 record）；用 summary() 查看历史")
                .into()
        })
    }

    /// 恢复最近一次被撤销的操作（重做），返回单行描述。与 [`undo`](Self::undo) 完全对称。
    pub fn redo(&self, wb: &Workbook) -> Result<String> {
        self.swap(wb, &self.redo, &self.entries)?.ok_or_else(|| {
            ExcelError::new("redo", "history", "没有可恢复的操作")
                .with_hint("只有撤销过的操作能恢复；撤销后又做了新的修改，恢复即作废")
                .into()
        })
    }

    /// 撤销 / 恢复的公共过程（两栈对称）：从 `from` 弹一条，按其逆向 `restore`；
    /// 成功再把目标区域「当前状态」的反向记录压入 `to`，失败则把条目放回 `from`。
    ///
    /// **失败要把条目放回去**：否则调用方一重试就会动到下一条（撤错东西），用户看到的
    /// 也是"这条已经撤过了"的假象。WPS 在连续批量操作时会偶发 0x80020009（"发生意外"，
    /// 通常是服务端正忙），正是靠回退重试来救的。返回 `Ok(None)` = 栈已空。
    fn swap(
        &self,
        wb: &Workbook,
        from: &RefCell<Vec<UndoEntry>>,
        to: &RefCell<Vec<UndoEntry>>,
    ) -> Result<Option<String>> {
        let Some(entry) = from.borrow_mut().pop() else {
            return Ok(None);
        };
        let desc = entry.to_string();
        // 先把目标区域的「当前状态」拍下来，它就是要恢复回去的那一版
        let forward = match snapshot_like(wb, &entry.op) {
            Ok(op) => op,
            Err(e) => {
                self.append(from, entry);
                return Err(e);
            }
        };
        if let Err(e) = restore(wb, &entry.op) {
            self.append(from, entry);
            return Err(e);
        }
        self.append(to, UndoEntry { desc: desc.clone(), op: forward });
        Ok(Some(desc))
    }

    /// 记录一次新操作：**旧的恢复点作废**（撤销后又改，就回不去了 —— 与 Word/Excel 一致）
    fn push(&self, desc: String, op: UndoOp) {
        self.redo.borrow_mut().clear();
        self.append(&self.entries, UndoEntry { desc, op });
    }

    /// 压栈并裁掉最旧的（上限）。
    fn append(&self, stack: &RefCell<Vec<UndoEntry>>, entry: UndoEntry) {
        let mut stack = stack.borrow_mut();
        while stack.len() >= self.limit {
            stack.remove(0);
        }
        stack.push(entry);
    }
}

/// 把一个栈格式化成「`#序号 描述`」清单（旧→新）。
fn summarize(stack: &RefCell<Vec<UndoEntry>>) -> Vec<String> {
    stack
        .borrow()
        .iter()
        .enumerate()
        .map(|(i, e)| format!("#{} {}", i + 1, e))
        .collect()
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
        // 0 不是合法 xlHAlign/xlVAlign：WPS 对「混合/常规」区域会返回 0（而非 Null），
        // 这种值写回必抛 0x80020009，记 None 跳过（其余格式项照常还原）
        h_align: i("HorizontalAlignment").filter(|v| *v != 0),
        v_align: i("VerticalAlignment").filter(|v| *v != 0),
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
        fill: snap_fill(&rng),
        borders: rng.borders().map(|b| b.snapshot()).unwrap_or_default(),
    })
}

/// 条件格式快照：读该区域全部规则（读不出来时 rules 为空，撤销表现为清空该区域规则）。
fn snap_conds(ws: &Worksheet, addr: &str) -> Result<CondSnapshot> {
    let rng = ws.range(addr)?;
    Ok(CondSnapshot {
        sheet: ws.name()?,
        addr: addr.to_string(),
        rules: rng.format_conditions()?.snapshot().unwrap_or_default(),
    })
}

/// 按类型执行逆向恢复（借 op：失败时调用方还要把它装回栈里）。
fn restore(wb: &Workbook, op: &UndoOp) -> Result<()> {
    match op {
        UndoOp::Cells(s) => {
            let ws = target_sheet(wb, &s.sheet)?;
            let rng = ws.range(&s.addr)?;
            if rng.is_merged()? && !s.merged {
                rng.unmerge()?;
            }
            write_back(&ws, &rng, s)?;
            if s.merged && !rng.is_merged()? {
                rng.merge()?;
            }
            Ok(())
        }
        UndoOp::Formats(s) => {
            let rng = target_range(wb, &s.sheet, &s.addr)?;
            apply_formats(&rng, s)
        }
        UndoOp::Both(c, f) => {
            let ws = target_sheet(wb, &c.sheet)?;
            let rng = ws.range(&c.addr)?;
            if rng.is_merged()? && !c.merged {
                rng.unmerge()?;
            }
            write_back(&ws, &rng, c)?;
            if c.merged && !rng.is_merged()? {
                rng.merge()?;
            }
            apply_formats(&rng, f)
        }
        UndoOp::SheetDelete { name } => {
            let ws = target_sheet(wb, name)?;
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
        // 建空表：内容由恢复栈里排在它后面的 Cells 记录补回（撤销时删表是最后一步，
        // 恢复时建表就是第一步 —— 两栈后进先出正好对上）
        UndoOp::SheetCreate { name } => wb.add_sheet(name, None).map(|_| ()),
        UndoOp::Conds(c) => target_range(wb, &c.sheet, &c.addr)?.apply_conds(&c.rules),
        UndoOp::SheetRename { from, to } => target_sheet(wb, from)?.rename(to),
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
///
/// **单项写不进去不能中断整条撤销**：有些值引擎读得出来却写不回去（见 `snap_formats`
/// 里 WPS 返回 0 的对齐），一 put 就抛 0x80020009。旧实现遇错即整个 restore 失败，
/// 而 `History::undo` 失败会把条目放回栈顶 —— 于是整个回退栈被一条永远撤不掉的记录
/// 卡死（用户实报：无论点哪个回退按钮都报同一个错）。格式只是外观，写不进去的那一项
/// 跳过即可，别拖垮其余项与整个栈。
fn apply_formats(rng: &Range, s: &FormatSnapshot) -> Result<()> {
    let raw = rng.raw();
    if let Some(v) = &s.number_format {
        let _ = raw.put("NumberFormat", vec![v.clone().into()]);
    }
    if let Some(v) = s.h_align {
        let _ = raw.put("HorizontalAlignment", vec![v.into()]);
    }
    if let Some(v) = s.v_align {
        let _ = raw.put("VerticalAlignment", vec![v.into()]);
    }
    if let Some(v) = s.wrap {
        let _ = raw.put("WrapText", vec![v.into()]);
    }
    if let Some(v) = s.indent {
        let _ = raw.put("IndentLevel", vec![v.into()]);
    }
    if let Some(v) = s.orientation {
        let _ = raw.put("Orientation", vec![v.into()]);
    }
    if let Some(v) = s.shrink {
        let _ = raw.put("ShrinkToFit", vec![v.into()]);
    }
    if let Some(v) = s.col_width {
        let _ = raw.put("ColumnWidth", vec![v.into()]);
    }
    if let Some(v) = s.row_height {
        let _ = raw.put("RowHeight", vec![v.into()]);
    }
    if let Ok(f) = rng.font() {
        if let Some(v) = s.font_bold {
            let _ = f.set_bold(v);
        }
        if let Some(v) = s.font_italic {
            let _ = f.set_italic(v);
        }
        if let Some(v) = s.font_underline {
            let _ = f.set_underline(v);
        }
        if let Some(v) = s.font_size {
            let _ = f.set_size(v);
        }
        if let Some(v) = s.font_color {
            let _ = f.set_color(v);
        }
        if let Some(v) = &s.font_name {
            let _ = f.set_name(v);
        }
    }
    restore_fill(&rng, &s.fill);
    // 边框逐边还原（快照为空 = 当时读不到边框对象，跳过）
    if let Ok(borders) = rng.borders() {
        borders.restore(&s.borders);
    }
    Ok(())
}

/// 填充快照（见 [`FillSnap`]）：整块一致一次读搞定；混杂才逐格细读，且只在小区域上做。
fn snap_fill(rng: &Range) -> FillSnap {
    if let Some(state) = rng.interior().ok().and_then(|x| x.fill()) {
        return FillSnap::Uniform(state);
    }
    let (rows, cols) = (rng.rows().unwrap_or(0), rng.columns().unwrap_or(0));
    if rows <= 0 || cols <= 0 || rows as i64 * cols as i64 > FILL_CELL_LIMIT {
        return FillSnap::Skip;
    }
    let mut states = Vec::with_capacity((rows * cols) as usize);
    for r in 0..rows {
        // 先按行读一次：表头/表体那种成块布局，绝大多数行一次就读出统一值，
        // 省下逐格的往返（逐格读一格要 3~4 次 COM 调用，小区域也得上百 ms）
        match row(rng, r, cols).and_then(|x| x.interior().ok()).and_then(|x| x.fill()) {
            Some(state) => states.resize(states.len() + cols as usize, Some(state)),
            None => {
                for c in 0..cols {
                    states.push(cell(rng, r, c).and_then(|x| x.interior().ok()).and_then(|x| x.fill()));
                }
            }
        }
    }
    FillSnap::Cells { cols, states }
}

/// 还原填充：整块直接写整块，逐格的写回各自那一格（`None` 跳过 = 当时读不出来）。
fn restore_fill(rng: &Range, snap: &FillSnap) {
    match snap {
        FillSnap::Uniform(state) => put_fill(rng, Some(*state)),
        FillSnap::Cells { cols, states } => {
            for (index, state) in states.iter().enumerate() {
                let (r, c) = (index as i32 / cols, index as i32 % cols);
                if let Some(cell) = cell(rng, r, c) {
                    put_fill(&cell, *state);
                }
            }
        }
        FillSnap::Skip => {}
    }
}

/// 写一格（或一块）的填充：无填充要清掉（否则撤销后新刷的底纹还留在表上），读不出的跳过。
fn put_fill(rng: &Range, state: Option<Option<u32>>) {
    let Ok(interior) = rng.interior() else { return };
    match state {
        Some(None) => {
            let _ = interior.clear();
        }
        Some(Some(v)) => {
            let _ = interior.set_color(v);
        }
        None => {}
    }
}

/// 区域左上角偏移 `r` 行 `c` 列的那一格。
fn cell(rng: &Range, r: i32, c: i32) -> Option<Range> {
    rng.offset(r, c).ok()?.resize(1, 1).ok()
}

/// 区域的第 `r` 行（`cols` 列）。
fn row(rng: &Range, r: i32, cols: i32) -> Option<Range> {
    rng.offset(r, 0).ok()?.resize(1, cols).ok()
}

/// 按某条记录的**目标区域**取「当前状态」快照，做成一条可反向应用的条目
/// （撤销 ⇄ 恢复互换用）。
///
/// 区域类（Cells/Formats/Both）拍当前状态；表结构类没有可读的区域快照，
/// 直接把动作反过来即可 —— 这样撤销「新增工作表」后那条会进恢复栈，重做时能建回空表，
/// 排在它后面的内容记录也就能照常恢复（否则恢复栈里那些记录会指向一张已删掉的表，
/// 一点恢复就报「工作表不存在」，且条目反复留在栈里永远撤不掉）。
fn snapshot_like(wb: &Workbook, op: &UndoOp) -> Result<UndoOp> {
    Ok(match op {
        UndoOp::Cells(c) => UndoOp::Cells(snap_cells(&target_sheet(wb, &c.sheet)?, &c.addr)?),
        UndoOp::Formats(f) => UndoOp::Formats(snap_formats(&target_sheet(wb, &f.sheet)?, &f.addr)?),
        UndoOp::Both(c, f) => UndoOp::Both(
            snap_cells(&target_sheet(wb, &c.sheet)?, &c.addr)?,
            snap_formats(&target_sheet(wb, &f.sheet)?, &f.addr)?,
        ),
        UndoOp::SheetDelete { name } => UndoOp::SheetCreate { name: name.clone() },
        UndoOp::SheetCreate { name } => UndoOp::SheetDelete { name: name.clone() },
        UndoOp::Conds(c) => UndoOp::Conds(snap_conds(&target_sheet(wb, &c.sheet)?, &c.addr)?),
        UndoOp::SheetRename { from, to } => UndoOp::SheetRename { from: to.clone(), to: from.clone() },
    })
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
