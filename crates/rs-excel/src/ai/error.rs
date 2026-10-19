//! 面向调用方（AI 工具层）的结构化错误。
//!
//! 库内部仍用 `anyhow` 传递错误；对外统一转成 [`ToolError`]：
//! 稳定的 `code` + 分类 + 可读原因 + 下一步提示 + HRESULT + 是否可重试 + 副作用状态。
//!
//! 这样调用方不需要再从 `"[put('Value')] @ WPS > Book1 > A1 | 0x800A03EC ..."` 里
//! 猜哪一段是错误码、能不能重发。

use std::fmt;

use crate::error::ExcelError;

/// `ai` 模块统一的返回类型。
pub type AiResult<T> = Result<T, ToolError>;

/// 错误分类：调用方按它决定「改参数 / 换目标 / 重试 / 停手」。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Category {
    /// 参数不合法（地址、形状、枚举值），未触碰 COM。
    Validation,
    /// 工作簿 / 工作表 / 命名区域不存在或已失效。
    TargetNotFound,
    /// 权限不足（覆盖文件、保护状态、只读）。
    Permission,
    /// 当前引擎/版本不支持该能力。
    Unsupported,
    /// COM 瞬时失败（WPS 常见的 0x80020009），可安全重试一次。
    ComTransient,
    /// 其他 COM / 应用层失败。
    ComFailure,
    /// 执行后恢复应用状态失败。
    Recovery,
}

impl Category {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Validation => "validation",
            Self::TargetNotFound => "target_not_found",
            Self::Permission => "permission_denied",
            Self::Unsupported => "unsupported",
            Self::ComTransient => "com_transient",
            Self::ComFailure => "com_failure",
            Self::Recovery => "recovery_failure",
        }
    }

    /// 分类默认提示：只在底层没给 hint 时使用。
    pub fn default_hint(self) -> &'static str {
        match self {
            Self::Validation => "按 message 指出的字段/形状改一次再发",
            Self::TargetNotFound => "先看回执里列出的可用名称，不要猜",
            Self::Permission => "换一个路径，或先关闭占用该文件的程序",
            Self::Unsupported => "当前引擎不支持该能力，改用等价操作",
            Self::ComTransient => "可原样重试一次；重复失败说明参数或状态有问题，先读目标区域",
            Self::ComFailure => "按 message 里的对象路径与原因修正后重试，不要盲目重发",
            Self::Recovery => "应用状态可能没完全恢复，先看当前设置再继续",
        }
    }

    pub fn retryable(self) -> bool {
        matches!(self, Self::ComTransient)
    }
}

/// 失败时操作产生的副作用状态。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    /// 可证明在产生副作用之前失败。
    NotApplied,
    /// 已生效。
    Applied,
    /// 无法确认是否已生效：先读目标再决定是否重试。
    Unknown,
}

impl Outcome {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::NotApplied => "not_applied",
            Self::Applied => "applied",
            Self::Unknown => "unknown",
        }
    }
}

/// 结构化错误（可序列化为工具回执里的 `error` 对象）。
#[derive(Debug, Clone)]
pub struct ToolError {
    /// 稳定错误码，如 `SHEET_NOT_FOUND` / `COM_TRANSIENT`。
    pub code: &'static str,
    pub category: Category,
    /// 出错的操作名（如 `write`）。
    pub operation: String,
    /// 出错的目标（工作表名 / 区域地址）。
    pub target: String,
    pub message: String,
    pub hint: String,
    /// 底层 HRESULT（如 `0x80020009`）。
    pub hresult: Option<String>,
    pub retryable: bool,
    pub outcome: Outcome,
}

impl ToolError {
    pub fn new(code: &'static str, category: Category, message: impl Into<String>) -> Self {
        Self {
            code,
            category,
            operation: String::new(),
            target: String::new(),
            message: message.into(),
            hint: category.default_hint().to_string(),
            hresult: None,
            retryable: category.retryable(),
            outcome: Outcome::Unknown,
        }
    }

    /// 参数校验失败：确定没产生副作用。
    pub fn validation(
        code: &'static str,
        message: impl Into<String>,
        hint: impl Into<String>,
    ) -> Self {
        Self {
            code,
            category: Category::Validation,
            operation: String::new(),
            target: String::new(),
            message: message.into(),
            hint: hint.into(),
            hresult: None,
            retryable: false,
            outcome: Outcome::NotApplied,
        }
    }

    pub fn with_code(mut self, code: &'static str) -> Self {
        self.code = code;
        self
    }

    pub fn with_category(mut self, category: Category) -> Self {
        self.category = category;
        self.retryable = category.retryable();
        if self.hint.is_empty() {
            self.hint = category.default_hint().to_string();
        }
        self
    }

    pub fn with_operation(mut self, operation: impl Into<String>) -> Self {
        self.operation = operation.into();
        self
    }

    pub fn with_target(mut self, target: impl Into<String>) -> Self {
        self.target = target.into();
        self
    }

    pub fn with_hint(mut self, hint: impl Into<String>) -> Self {
        self.hint = hint.into();
        self
    }

    pub fn with_hresult(mut self, hresult: impl Into<String>) -> Self {
        self.hresult = Some(hresult.into());
        self
    }

    pub fn with_outcome(mut self, outcome: Outcome) -> Self {
        self.outcome = outcome;
        self
    }

    pub fn with_retryable(mut self, retryable: bool) -> Self {
        self.retryable = retryable;
        self
    }

    /// 从 `anyhow` 错误转结构化错误（保留 `ExcelError` 的对象路径与 hint）。
    pub fn from_anyhow(e: &anyhow::Error) -> Self {
        // 已经是结构化错误（校验、保存策略等）就不再包一层。
        if let Some(err) = e.downcast_ref::<ToolError>() {
            return err.clone();
        }
        let message = e.to_string();
        let (operation, target, hint, source) = match e.downcast_ref::<ExcelError>() {
            Some(err) => (
                err.op.clone(),
                err.obj.clone(),
                err.hint.clone(),
                err.source.clone(),
            ),
            None => (String::new(), String::new(), String::new(), message.clone()),
        };
        let hresult = extract_hresult(&source).or_else(|| extract_hresult(&message));
        let (code, category) = match hresult.as_deref() {
            Some("0x80020009") => ("COM_TRANSIENT", Category::ComTransient),
            Some("0x80020003") => ("UNSUPPORTED", Category::Unsupported),
            Some("0x80020005") => ("BAD_TYPE", Category::Validation),
            Some("0x800A03EC") => ("COM_FAILURE", Category::ComFailure),
            Some(_) => ("COM_FAILURE", Category::ComFailure),
            None => ("FAILED", Category::ComFailure),
        };
        Self {
            code,
            category,
            operation,
            target,
            message,
            hint: if hint.is_empty() {
                category.default_hint().to_string()
            } else {
                hint
            },
            hresult,
            retryable: category.retryable(),
            outcome: Outcome::Unknown,
        }
    }

    /// 消息里是否带 WPS 的瞬时错误码（调用方重试判据）。
    pub fn is_transient(message: &str) -> bool {
        message.contains("0x80020009")
    }
}

impl fmt::Display for ToolError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "[{}] {}", self.code, self.message)?;
        if !self.hint.is_empty() {
            write!(f, " | 提示: {}", self.hint)?;
        }
        Ok(())
    }
}

impl std::error::Error for ToolError {}

impl From<anyhow::Error> for ToolError {
    fn from(e: anyhow::Error) -> Self {
        Self::from_anyhow(&e)
    }
}

/// 从错误文本里取第一个 `0xXXXXXXXX`（统一大写）。
fn extract_hresult(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let mut i = 0;
    while i + 10 <= bytes.len() {
        if bytes[i] == b'0'
            && (bytes[i + 1] == b'x' || bytes[i + 1] == b'X')
            && bytes[i + 2..i + 10].iter().all(|b| b.is_ascii_hexdigit())
        {
            return Some(format!("0x{}", &text[i + 2..i + 10]).to_uppercase());
        }
        i += 1;
    }
    None
}
