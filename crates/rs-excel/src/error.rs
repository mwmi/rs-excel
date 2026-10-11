//! 单行结构化错误 + 颜色助手。
//!
//! 错误输出格式（单行，AI 一次读取即获得全部定位要素）：
//! `[操作] @ 对象路径 | 底层原因 | 提示: 修复建议`
//! 例：`[put('Value')] @ WPS > Book1.xlsx > Sheet1 > A1:C3 | 0x800A03EC ... | 提示: 区域形状需与数据一致`

use std::fmt;

/// 结构化错误：操作 / 对象路径 / 底层原因 / 修复建议 四要素。
#[derive(Debug)]
pub struct ExcelError {
    /// 什么操作，如 `put('Value')`
    pub op: String,
    /// 对象路径，如 `WPS > Book1.xlsx > Sheet1 > A1:C3`
    pub obj: String,
    /// 底层 COM 原因，如 `0x800A03EC 参数类型错误`
    pub source: String,
    /// 修复建议
    pub hint: String,
}

impl ExcelError {
    pub fn new(op: impl Into<String>, obj: impl Into<String>, source: impl Into<String>) -> Self {
        Self {
            op: op.into(),
            obj: obj.into(),
            source: source.into(),
            hint: String::new(),
        }
    }

    pub fn with_hint(mut self, hint: impl Into<String>) -> Self {
        self.hint = hint.into();
        self
    }

    /// 给已有错误补提示（尽量保持 ExcelError 类型不丢失上下文）。
    pub fn hint(e: anyhow::Error, hint: impl Into<String>) -> anyhow::Error {
        match e.downcast::<ExcelError>() {
            Ok(mut err) => {
                err.hint = hint.into();
                err.into()
            }
            Err(e) => ExcelError::new("error", "", e.to_string())
                .with_hint(hint)
                .into(),
        }
    }
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "[{}] @ {} | {}", self.op, self.obj, self.source)?;
        if !self.hint.is_empty() {
            write!(f, " | 提示: {}", self.hint)?;
        }
        Ok(())
    }
}

impl std::error::Error for ExcelError {}

/// RGB → OLE 颜色。COM 内部颜色格式为 `0x00BBGGRR`（与网页 RGB 相反）。
///
/// ```ignore
/// use rs_excel::rgb;
/// font.set_color(rgb(255, 0, 0))?; // 红色
/// ```
pub fn rgb(r: u8, g: u8, b: u8) -> u32 {
    (b as u32) << 16 | (g as u32) << 8 | r as u32
}

/// OLE 颜色 → `#RRGGBB` 文本。
pub(crate) fn ole_to_hex(c: u32) -> String {
    format!(
        "#{:02X}{:02X}{:02X}",
        c & 0xFF,
        (c >> 8) & 0xFF,
        (c >> 16) & 0xFF
    )
}
