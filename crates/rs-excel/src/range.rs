//! 区域（Range）：单元格/区域读写、公式、合并、格式、对齐、行高列宽、批量 TSV、样式入口。

use anyhow::{bail, Result};

use crate::com::dispatch::{member_notfound, IDispatchWrapper};
use crate::com::safearray::{list_to_variant, rows_to_variant, variant_to_rows, variant_to_rows_typed};
use crate::com::variant::{CellValue, ComValue};
use crate::error::ExcelError;
use crate::pivot::AggFunc;
use crate::style::{Borders, Font, Interior};
use crate::worksheet::parse_tsv;

/// 水平对齐（xlHAlign 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AlignH {
    Left = -4131,
    Center = -4108,
    Right = -4152,
    Justify = -4130,
    Distributed = -4117,
}

/// 垂直对齐（xlVAlign 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AlignV {
    Top = -4160,
    Center = -4108,
    Bottom = -4107,
    Justify = -4130,
    Distributed = -4117,
}

/// `End` 方向（等价 Ctrl+方向键，xlUp/xlDown/xlToLeft/xlToRight）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EndDir {
    Up = -4162,
    Down = -4121,
    Left = -4159,
    Right = -4158,
}

/// 特殊单元格类型（xlCellType 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SpecialCells {
    Blanks = 4,
    Constants = 2,
    Formulas = -4123,
    Visible = 12,
    LastCell = 11,
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
    /// 默认走 `Value2`（最快）；日期返回序列值 f64，日期文本用 [`Range::value_compat`]。
    pub fn value(&self) -> Result<CellValue> {
        let res = self.raw.get("Value2")?;
        let rows = variant_to_rows(&res.v)?;
        Ok(rows.into_iter().flatten().next().unwrap_or(CellValue::Empty))
    }

    /// 显示文本（应用格式化后的样子）。
    pub fn text(&self) -> Result<String> {
        self.raw.get("Text")?.string()
    }

    /// 批量读整个区域为网格（走 `Value2`，日期为序列值）。
    pub fn values(&self) -> Result<Vec<Vec<CellValue>>> {
        variant_to_rows(&self.raw.get("Value2")?.v)
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

    /// 公式（无公式返回空串）。优先 `Formula2`（动态数组），引擎不支持自动回退 `Formula`。
    pub fn formula(&self) -> Result<String> {
        self.formula_by("Formula2")
    }

    // ── 兼容通道（VBA 原生 Value/Formula，速度略慢；日期转为可读文本） ──

    /// 单格值，走 `Value`：日期/货币由引擎转换为类型，日期返回 `yyyy-mm-dd` 文本。
    pub fn value_compat(&self) -> Result<CellValue> {
        let res = self.raw.get("Value")?;
        let rows = variant_to_rows_typed(&res.v)?;
        Ok(rows.into_iter().flatten().next().unwrap_or(CellValue::Empty))
    }

    /// 批量读网格，走 `Value`（日期为可读文本）。
    pub fn values_compat(&self) -> Result<Vec<Vec<CellValue>>> {
        variant_to_rows_typed(&self.raw.get("Value")?.v)
    }

    /// 写单值，走 `Value`（文本会按单元格格式解析，如 "2026-01-01" 识别为日期）。
    pub fn set_value_compat(&self, v: impl Into<ComValue>) -> Result<()> {
        self.raw.put("Value", vec![v.into()])
    }

    /// 公式，走 `Formula`（旧语义，无动态数组标记）。
    pub fn formula_compat(&self) -> Result<String> {
        self.formula_by("Formula")
    }

    /// 写公式，走 `Formula`。
    pub fn set_formula_compat(&self, f: &str) -> Result<()> {
        self.set_formula_by("Formula", f)
    }

    /// 区域内是否含公式。
    pub fn has_formula(&self) -> Result<bool> {
        Ok(matches!(self.raw.get("HasFormula")?.value(), CellValue::Bool(true)))
    }

    // ── 写 ──

    /// 写单值（区域会被整体填充同值）。走 `Value2`（最快）。
    pub fn set_value(&self, v: impl Into<ComValue>) -> Result<()> {
        self.raw.put("Value2", vec![v.into()])
    }

    /// 批量写网格（走 `Value2`）。形状必须与区域一致；SAFEARRAY 批量失败自动回退逐格。
    pub fn set_values(&self, grid: &[Vec<CellValue>]) -> Result<()> {
        if grid.is_empty() || grid[0].is_empty() {
            bail!("set_values：数据为空");
        }
        match rows_to_variant(grid) {
            Ok(v) => self.raw.put_raw("Value2", v),
            Err(_) => {
                // 回退：逐格写（相对区域左上角）
                for (r, row) in grid.iter().enumerate() {
                    for (c, val) in row.iter().enumerate() {
                        self.raw
                            .call("Cells", vec![(r as i32 + 1).into(), (c as i32 + 1).into()])?
                            .idispatch()?
                            .put("Value2", vec![ComValue::from(val.clone())])?;
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

    /// 写公式，如 `set_formula("=SUM(A1:A9)")`。优先 `Formula2`（动态数组），
    /// 引擎不支持自动回退 `Formula`。
    pub fn set_formula(&self, f: &str) -> Result<()> {
        self.set_formula_by("Formula2", f)
    }

    /// 数字格式，如 `number_format("0.00")` / `"yyyy-mm-dd"` / `"0.00%"`。
    pub fn number_format(&self, fmt: &str) -> Result<()> {
        self.raw.put("NumberFormat", vec![fmt.into()])
    }

    /// 批量读公式网格（走 `Formula2`，引擎不支持自动回退 `Formula`）。
    /// 常量格返回其原类型值（Number/Text/Bool），公式格返回 `=...` 公式文本。
    pub fn formulas(&self) -> Result<Vec<Vec<CellValue>>> {
        let res = match self.raw.get("Formula2") {
            Ok(r) => r,
            Err(e) if member_notfound(&e) => self.raw.get("Formula")?,
            Err(e) => return Err(e),
        };
        variant_to_rows(&res.v)
    }

    /// 按成员名读公式；Formula2 不存在（MEMBERNOTFOUND）时回退 Formula。
    fn formula_by(&self, member: &str) -> Result<String> {
        let res = match self.raw.get(member) {
            Ok(r) => r,
            Err(e) if member_notfound(&e) => self.raw.get("Formula")?,
            Err(e) => return Err(e),
        };
        Ok(match res.value() {
            CellValue::Text(s) => s,
            _ => String::new(),
        })
    }

    /// 按成员名写公式；Formula2 不存在时回退 Formula。
    fn set_formula_by(&self, member: &str, f: &str) -> Result<()> {
        match self.raw.put(member, vec![f.into()]) {
            Ok(()) => Ok(()),
            Err(e) if member_notfound(&e) => self.raw.put("Formula", vec![f.into()]),
            Err(e) => Err(ExcelError::hint(
                e,
                "公式需以 = 开头，如 =SUM(A1:A9)",
            )),
        }
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

    /// 水平对齐原始值（xlHAlign，1 = 常规）。
    pub fn h_align(&self) -> Result<i32> {
        self.raw.get("HorizontalAlignment")?.int()
    }

    /// 垂直对齐原始值（xlVAlign）。
    pub fn v_align(&self) -> Result<i32> {
        self.raw.get("VerticalAlignment")?.int()
    }

    /// 是否自动换行（多段显示必须开）。
    pub fn wrap_text(&self) -> Result<bool> {
        self.raw.get("WrapText")?.bool()
    }

    /// 缩进级别（0~15）。
    pub fn indent(&self) -> Result<i32> {
        self.raw.get("IndentLevel")?.int()
    }

    pub fn set_indent(&self, level: i32) -> Result<()> {
        self.raw.put("IndentLevel", vec![level.into()])
    }

    /// 文字方向（角度 -90~90，竖排为 -4166）。
    pub fn orientation(&self) -> Result<i32> {
        self.raw.get("Orientation")?.int()
    }

    pub fn set_orientation(&self, deg: i32) -> Result<()> {
        self.raw.put("Orientation", vec![deg.into()])
    }

    /// 是否缩小字体填充。
    pub fn shrink(&self) -> Result<bool> {
        self.raw.get("ShrinkToFit")?.bool()
    }

    pub fn set_shrink(&self, on: bool) -> Result<()> {
        self.raw.put("ShrinkToFit", vec![on.into()])
    }

    // ── 部分字符（单元格内局部文字格式，1 基） ──

    /// 引用第 `start` 个字符起 `len` 个字符的字体（按段落局部格式化的核心入口）。
    pub fn char_font(&self, start: i32, len: i32) -> Result<Font> {
        Ok(Font::from_raw(
            self.characters(start, len)?
                .get_obj("Font", "Chars.Font")?,
        ))
    }

    /// 读取部分文本。
    pub fn char_text(&self, start: i32, len: i32) -> Result<String> {
        self.characters(start, len)?.get("Text")?.string()
    }

    /// 替换部分文本（新文本长度可不同于原文，即插入/删除）。
    pub fn set_char_text(&self, start: i32, len: i32, text: &str) -> Result<()> {
        self.characters(start, len)?
            .put("Text", vec![text.into()])
            .map_err(|e| ExcelError::hint(e, "替换长度可不同（插入/删除），但 start+len 须不超过原文本长度"))
    }

    /// Characters(start, len) 对象。start/len 均 1 基且 ≥1。
    fn characters(&self, start: i32, len: i32) -> Result<IDispatchWrapper> {
        if start < 1 || len < 1 {
            bail!("Characters 参数须 ≥1（得到 start={start}, len={len}）");
        }
        self.raw
            .call("Characters", vec![start.into(), len.into()])?
            .idispatch_child(&format!("Chars({start},{len})"))
            .map_err(|e| ExcelError::hint(e, "start+len 不得超过单元格文本长度（可先 get 读取确认）"))
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

    /// 复制到剪贴板（随后在目标区域 `paste_values` / `paste_all`）。
    pub fn copy(&self) -> Result<()> {
        self.raw.call("Copy", vec![]).map(|_| ())
    }

    /// 剪切到目标区域。
    pub fn cut_to(&self, dst: &Range) -> Result<()> {
        self.raw
            .call("Cut", vec![ComValue::Obj(dst.raw().disp.clone())])
            .map(|_| ())
    }

    /// 只粘贴值（调用前先在源区域 `copy()`），粘贴后清剪贴板状态。
    pub fn paste_values(&self) -> Result<()> {
        // xlPasteValues = -4163
        self.raw.call("PasteSpecial", vec![(-4163i32).into()]).map(|_| ())?;
        self.release_clipboard();
        Ok(())
    }

    /// 全量粘贴（内容+格式，调用前先 `copy()`），粘贴后清剪贴板状态。
    pub fn paste_all(&self) -> Result<()> {
        // xlPasteAll = -4104
        self.raw.call("PasteSpecial", vec![(-4104i32).into()]).map(|_| ())?;
        self.release_clipboard();
        Ok(())
    }

    /// 清除剪贴板复制状态（Application.CutCopyMode = False）。失败不影响主流程。
    fn release_clipboard(&self) {
        let _ = self
            .raw
            .get("Application")
            .and_then(|a| a.idispatch())
            .and_then(|a| a.put("CutCopyMode", vec![false.into()]));
    }

    // ── 定位导航 ──

    /// 偏移：`offset(1, 2)` 把 A1 变为 C2。
    pub fn offset(&self, rows: i32, cols: i32) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw
                .call("Offset", vec![rows.into(), cols.into()])?
                .idispatch()?,
        ))
    }

    /// 调整区域尺寸（左上角不变）：`resize(10, 3)` 把 A1 变为 A1:C10。
    pub fn resize(&self, rows: i32, cols: i32) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw
                .call("Resize", vec![rows.into(), cols.into()])?
                .idispatch()?,
        ))
    }

    /// 边界定位（等价 Ctrl+方向键）：如 `range("A1").end(EndDir::Down)` 得到 A 列最后非空格。
    pub fn end(&self, dir: EndDir) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw.call("End", vec![(dir as i32).into()])?.idispatch()?,
        ))
    }

    /// 当前连续数据区域（等价 Ctrl+A）。
    pub fn current_region(&self) -> Result<Range> {
        Ok(Range::from_raw(self.raw.get_obj("CurrentRegion", "CurrentRegion")?))
    }

    /// 整行区域。
    pub fn entire_row(&self) -> Result<Range> {
        Ok(Range::from_raw(self.raw.get_obj("EntireRow", "EntireRow")?))
    }

    /// 整列区域。
    pub fn entire_column(&self) -> Result<Range> {
        Ok(Range::from_raw(self.raw.get_obj("EntireColumn", "EntireColumn")?))
    }

    /// 定位特殊单元格（空值/常量/公式/可见/最后单元格）。无匹配返回 `Ok(None)`，不报错。
    pub fn special_cells(&self, kind: SpecialCells) -> Result<Option<Range>> {
        match self.raw.call("SpecialCells", vec![(kind as i32).into()]) {
            Ok(v) if v.is_object() => Ok(Some(Range::from_raw(v.idispatch()?))),
            // 0x800A03EC = "未找到单元格"，属正常业务结果
            Err(e) if e.to_string().contains("0x800A03EC") => Ok(None),
            Err(e) => Err(e),
            Ok(_) => Ok(None),
        }
    }

    // ── 编辑 ──

    /// 插入单元格，`shift_down=true` 下方单元格下移，否则右侧右移。
    pub fn insert(&self, shift_down: bool) -> Result<()> {
        // xlShiftDown = -4121, xlShiftToRight = -4159
        let shift: i32 = if shift_down { -4121 } else { -4159 };
        self.raw.call("Insert", vec![shift.into()]).map(|_| ())
    }

    /// 删除单元格，`shift_up=true` 下方单元格上移，否则左侧左移。
    pub fn delete(&self, shift_up: bool) -> Result<()> {
        // xlShiftUp = -4162, xlShiftToLeft = -4158
        let shift: i32 = if shift_up { -4162 } else { -4158 };
        self.raw.call("Delete", vec![shift.into()]).map(|_| ())
    }

    /// 向下填充（首行内容填充到区域其余行）。
    pub fn fill_down(&self) -> Result<()> {
        self.raw.call("FillDown", vec![]).map(|_| ())
    }

    /// 向右填充（首列内容填充到区域其余列）。
    pub fn fill_right(&self) -> Result<()> {
        self.raw.call("FillRight", vec![]).map(|_| ())
    }

    // ── 查找 / 替换 / 排序 / 筛选 ──

    /// 查找第一个匹配。未找到返回 `Ok(None)`（不报错，AI 无需处理错误分支）。
    pub fn find(&self, what: &str, whole: bool) -> Result<Option<Range>> {
        // LookAt：xlWhole = 1（全字匹配），xlPart = 2（包含匹配）
        let look_at: i32 = if whole { 1 } else { 2 };
        match self
            .raw
            .call_named("Find", vec![what.into()], &[("LookAt", look_at.into())])
        {
            Ok(v) if v.is_object() => Ok(Some(Range::from_raw(v.idispatch()?))),
            Ok(_) => Ok(None),
            Err(e) if e.to_string().contains("0x800A03EC") => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// 区域内替换文本。
    pub fn replace(&self, what: &str, with: &str) -> Result<()> {
        self.raw
            .call("Replace", vec![what.into(), with.into()])
            .map_err(|e| ExcelError::hint(e, "参数顺序为 (旧文本, 新文本)"))
            .map(|_| ())
    }

    /// 多键排序（区域内 1 基列号，最多 3 个键）。`keys` 如 `sort_by_cols(&[(3, true), (9, false)], true)`。
    ///
    /// 走**宏录制器那条** `Range.Sort(Key1, Order1, Key2, Order2, Key3, Order3, Header)`：
    /// Key 取该列的**首行数据单元格**——有表头就是第 2 行。官方示例即此写法
    /// （`Columns("A:C").Sort key1:=Range("C2"), order1:=xlAscending, header:=xlYes`）。
    ///
    /// **不要用 SortFields + Apply 那条现代路径**：WPS 上 `Apply` 必抛 0x80020009「发生意外」
    /// （实测：同一区域两条路径全挂，报错完全一样），而它还得额外对齐 MatchCase / Orientation /
    /// SortMethod，失败面更大。录制器能跑通的那条才是 WPS 真正实现的（同 SlicerCaches.Add2 那条经验）。
    pub fn sort_by_cols(&self, keys: &[(i32, bool)], has_header: bool) -> Result<()> {
        if keys.is_empty() {
            bail!("sort_by_cols：keys 不能为空");
        }
        if keys.len() > 3 {
            bail!("sort_by_cols：最多 3 个键（Range.Sort 只有 Key1~Key3）");
        }
        // Key 必须落在数据区（不含表头行），否则 WPS 直接抛 0x80020009
        let row = if has_header { 2 } else { 1 };
        let mut named: Vec<(&str, ComValue)> = Vec::with_capacity(7);
        for (index, (col, asc)) in keys.iter().enumerate() {
            let (key_name, order_name) = [("Key1", "Order1"), ("Key2", "Order2"), ("Key3", "Order3")][index];
            let key = self.raw.call("Cells", vec![row.into(), (*col).into()])?.idispatch()?;
            named.push((key_name, ComValue::Obj(key.disp)));
            named.push((order_name, (if *asc { 1i32 } else { 2i32 }).into())); // xlAscending / xlDescending
        }
        named.push(("Header", (if has_header { 1i32 } else { 2i32 }).into())); // xlYes / xlNo
        self.raw
            .call_named("Sort", vec![], &named)
            .map_err(|e| ExcelError::hint(e, "Key 须是本区域内该列的数据单元格（不含表头行），且排序范围不得含透视表"))
            .map(|_| ())
    }

    /// 自动筛选：`field` 为区域内第几列（1 基），`criteria` 如 ">100"、"张三"。
    pub fn auto_filter(&self, field: i32, criteria: &str) -> Result<()> {
        self.raw
            .call_named(
                "AutoFilter",
                vec![],
                &[("Field", field.into()), ("Criteria1", criteria.into())],
            )
            .map(|_| ())
    }

    /// 取消筛选（对已筛选区域再次无参调用 AutoFilter）。
    pub fn filter_off(&self) -> Result<()> {
        self.raw.call("AutoFilter", vec![]).map(|_| ())
    }

    /// 双条件 AND 筛选，如 `filter_and(1, ">100", "<1000")`（数值区间）。
    pub fn filter_and(&self, field: i32, c1: &str, c2: &str) -> Result<()> {
        self.filter2(field, 1, c1, c2) // xlAnd
    }

    /// 双条件 OR 筛选，如 `filter_or(2, "北京", "上海")`。
    pub fn filter_or(&self, field: i32, c1: &str, c2: &str) -> Result<()> {
        self.filter2(field, 2, c1, c2) // xlOr
    }

    fn filter2(&self, field: i32, op: i32, c1: &str, c2: &str) -> Result<()> {
        self.raw
            .call_named(
                "AutoFilter",
                vec![],
                &[
                    ("Field", field.into()),
                    ("Criteria1", c1.into()),
                    ("Operator", op.into()),
                    ("Criteria2", c2.into()),
                ],
            )
            .map_err(|e| ExcelError::hint(e, "field 为区域内 1 基列号；条件支持 \">100\"、\"<>x\" 与通配符 * ?"))
            .map(|_| ())
    }

    /// 多选筛选（等价筛选项勾选），如 `filter_values(2, &["苹果", "香蕉"])`。
    /// 值须与单元格文本完全一致。
    pub fn filter_values(&self, field: i32, values: &[&str]) -> Result<()> {
        if values.is_empty() {
            bail!("filter_values：values 不能为空（清筛选用 filter_off）");
        }
        let arr = list_to_variant(values.iter().map(|s| (*s).into()).collect())?;
        self.raw
            .call_named(
                "AutoFilter",
                vec![],
                &[
                    ("Field", field.into()),
                    ("Criteria1", ComValue::Raw(arr)),
                    ("Operator", (7i32).into()), // xlFilterValues
                ],
            )
            .map_err(|e| ExcelError::hint(e, "值须与单元格文本完全一致（区分大小写与空格）"))
            .map(|_| ())
    }

    /// 前/后 N 项筛选，如 `filter_top(3, 10, false, true)` 取数值前 10；
    /// `percent=true` 时按百分比（1~100）。
    pub fn filter_top(&self, field: i32, n: i32, percent: bool, from_top: bool) -> Result<()> {
        // xlTop10Items=3 / xlBottom10Items=4 / xlTop10Percent=5 / xlBottom10Percent=6
        let op = match (percent, from_top) {
            (false, true) => 3,
            (false, false) => 4,
            (true, true) => 5,
            (true, false) => 6,
        };
        self.raw
            .call_named(
                "AutoFilter",
                vec![],
                &[("Field", field.into()), ("Criteria1", n.into()), ("Operator", op.into())],
            )
            .map_err(|e| ExcelError::hint(e, "field 为区域内 1 基列号，n 为 1~500 的项数或 1~100 的百分比"))
            .map(|_| ())
    }

    /// 高级筛选：`criteria` 条件区域（None 免条件；同行 AND、跨行 OR），
    /// `to` 目标区左上角（None 原位筛选，不满足条件的行被隐藏），`unique` 只取唯一行。
    pub fn filter_advanced(&self, criteria: Option<&Range>, to: Option<&Range>, unique: bool) -> Result<()> {
        let mut named: Vec<(&str, ComValue)> = vec![
            // xlFilterInPlace = 1 / xlFilterCopy = 2
            ("Action", (if to.is_some() { 2i32 } else { 1 }).into()),
            ("Unique", unique.into()),
        ];
        if let Some(c) = criteria {
            named.push(("CriteriaRange", ComValue::Obj(c.raw().disp.clone())));
        }
        if let Some(t) = to {
            named.push(("CopyToRange", ComValue::Obj(t.raw().disp.clone())));
        }
        self.raw
            .call_named("AdvancedFilter", vec![], &named)
            .map_err(|e| {
                ExcelError::hint(e, "条件区域首行须为与数据区一致的列标题（同行 AND、跨行 OR）；复制模式须给 to")
            })
            .map(|_| ())
    }

    /// 提取唯一行到目标区域（不动原数据）。
    pub fn extract_unique(&self, to: &Range) -> Result<()> {
        self.filter_advanced(None, Some(to), true)
    }

    /// 分类汇总。`group_col` 分组列（区域内 1 基列号），`total_cols` 参与汇总的列号，
    /// 如 `subtotal(1, AggFunc::Sum, &[2, 3])`。区域须含表头。
    pub fn subtotal(&self, group_col: i32, func: AggFunc, total_cols: &[i32]) -> Result<()> {
        if total_cols.is_empty() {
            bail!("subtotal：total_cols 不能为空");
        }
        let arr = list_to_variant(total_cols.iter().map(|c| (*c).into()).collect())?;
        self.raw
            .call_named(
                "Subtotal",
                vec![],
                &[
                    ("GroupBy", group_col.into()),
                    ("Function", func.code().into()),
                    ("TotalList", ComValue::Raw(arr)),
                    ("Replace", true.into()),
                ],
            )
            .map_err(|e| ExcelError::hint(e, "区域须含表头行；group_col/total_cols 为区域内 1 基列号"))
            .map(|_| ())
    }

    /// 移除分类汇总。
    pub fn remove_subtotal(&self) -> Result<()> {
        self.raw.call("RemoveSubtotals", vec![]).map(|_| ())
    }

    /// 行分组（行号左侧可折叠）。
    pub fn group_rows(&self) -> Result<()> {
        self.raw.get_obj("Rows", "Rows")?.call("Group", vec![]).map(|_| ())
    }

    /// 取消行分组。
    pub fn ungroup_rows(&self) -> Result<()> {
        self.raw.get_obj("Rows", "Rows")?.call("Ungroup", vec![]).map(|_| ())
    }

    /// 列分组（列号上方可折叠）。
    pub fn group_cols(&self) -> Result<()> {
        self.raw.get_obj("Columns", "Columns")?.call("Group", vec![]).map(|_| ())
    }

    /// 取消列分组。
    pub fn ungroup_cols(&self) -> Result<()> {
        self.raw.get_obj("Columns", "Columns")?.call("Ungroup", vec![]).map(|_| ())
    }

    // ── 批注 / 超链接 / 数据验证 ──

    /// 写批注（已有批注先清除）。
    pub fn set_comment(&self, text: &str) -> Result<()> {
        let _ = self.raw.call("ClearComments", vec![]); // 无批注时报错，属正常前置清理
        self.raw.call("AddComment", vec![text.into()]).map(|_| ())
    }

    /// 添加超链接。`text` 为显示文本（None 时保持原内容）。
    pub fn add_hyperlink(&self, url: &str, text: Option<&str>) -> Result<()> {
        let anchor = ComValue::Obj(self.raw.disp.clone());
        let named: Vec<(&str, ComValue)> = match text {
            Some(t) => vec![("TextToDisplay", t.into())],
            None => vec![],
        };
        self.raw
            .get_obj("Hyperlinks", "Hyperlinks")?
            .call_named("Add", vec![anchor, url.into()], &named)
            .map(|_| ())
    }

    /// 设置下拉列表验证，如 `set_validation_list("是,否")`。
    pub fn set_validation_list(&self, items: &str) -> Result<()> {
        let v = self.raw.get_obj("Validation", "Validation")?;
        let _ = v.call("Delete", vec![]); // 无验证时报错，属正常前置清理
        // xlValidateList = 3
        v.call_named("Add", vec![(3i32).into()], &[("Formula1", items.into())])
            .map_err(|e| {
                ExcelError::hint(e, "items 为逗号分隔列表，如 \"是,否\"；区域须未受保护")
            })
            .map(|_| ())
    }
}
