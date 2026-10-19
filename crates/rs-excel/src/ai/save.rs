//! 另存为：显式覆盖策略。
//!
//! 旧约定要求调用方「先删除目标文件」来绕开 WPS 的覆盖弹窗——路径写错就是数据丢失。
//! 这里改成：
//! - 默认 `FailIfExists`：目标存在就报 `FILE_EXISTS`，不删不覆盖；
//! - `Replace`：先把原文件改名成 `*.bak-时间戳`，再 SaveAs；失败把备份改回去；
//! - 保存后核实文件确实存在（引擎偶发「调用成功但没落盘」）。

use std::path::Path;
use std::time::{SystemTime, UNIX_EPOCH};

use crate::ai::error::{AiResult, Category, Outcome, ToolError};
use crate::workbook::{FileFormat, Workbook};

/// 覆盖策略。
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum OverwritePolicy {
    /// 目标存在就失败（默认）。
    #[default]
    FailIfExists,
    /// 先备份原文件再覆盖。
    Replace,
}

/// 保存结果。
#[derive(Debug, Clone)]
pub struct SaveReport {
    pub path: String,
    /// 是否覆盖了已有文件。
    pub replaced: bool,
    /// 备份文件路径（仅 `Replace` 且确实覆盖时有值）。
    pub backup: Option<String>,
    /// 落盘后的文件字节数（核实不到为 `None`）。
    pub bytes: Option<u64>,
}

/// 按策略另存为。
pub fn save_as(
    book: &Workbook,
    path: &str,
    format: FileFormat,
    policy: OverwritePolicy,
) -> AiResult<SaveReport> {
    let path = path.trim();
    if path.is_empty() {
        return Err(ToolError::validation(
            "BAD_PATH",
            "保存路径为空",
            "写绝对路径，如 D:\\报表\\销售.xlsx",
        ));
    }
    if !Path::new(path).is_absolute() {
        return Err(ToolError::validation(
            "BAD_PATH",
            format!("保存路径必须是绝对路径：「{path}」"),
            "如 D:\\报表\\销售.xlsx",
        ));
    }
    let current = book.full_name().unwrap_or_default();
    if !current.is_empty() && same_path(&current, path) {
        return Err(ToolError::validation(
            "SAME_FILE",
            "目标就是当前工作簿本身",
            "写回原路径用 save；save_as 用来另存到新文件",
        ));
    }
    let existed = Path::new(path).exists();
    let mut backup = None;
    if existed {
        match policy {
            OverwritePolicy::FailIfExists => {
                return Err(ToolError::new(
                    "FILE_EXISTS",
                    Category::Permission,
                    format!("目标文件已存在：{path}"),
                )
                .with_hint("换一个文件名；确实要覆盖就加 overwrite:\"replace\"（会先备份原文件）"));
            }
            OverwritePolicy::Replace => {
                let bak = backup_path(path);
                std::fs::rename(path, &bak).map_err(|e| {
                    ToolError::new(
                        "BACKUP_FAILED",
                        Category::Permission,
                        format!("覆盖前备份失败：{e}"),
                    )
                    .with_hint("关闭占用该文件的程序，或改用其他路径")
                })?;
                backup = Some(bak);
            }
        }
    }
    if let Err(e) = book.save_as_raw(path, format) {
        // 保存失败：把备份改回去，别让用户的原文件只剩一个 .bak
        let mut error = ToolError::from_anyhow(&e)
            .with_operation("save_as")
            .with_target(path)
            .with_hint("路径需为绝对路径；同名文件若已打开或只读会失败");
        if let Some(bak) = &backup {
            if std::fs::rename(bak, path).is_err() {
                error = error.with_hint(format!(
                    "保存失败，且原文件备份未能改回：{bak}（原文件内容还在这个 .bak 里，先别删）"
                ));
            }
        }
        return Err(error);
    }
    let bytes = std::fs::metadata(path).ok().map(|meta| meta.len());
    if bytes.is_none() {
        return Err(ToolError::new(
            "SAVE_UNVERIFIED",
            Category::ComFailure,
            format!("保存调用已返回，但目标文件不存在：{path}"),
        )
        .with_outcome(Outcome::Unknown)
        .with_hint("先确认文件是否真的落盘，再决定是否重试"));
    }
    Ok(SaveReport {
        path: path.to_string(),
        replaced: existed,
        backup,
        bytes,
    })
}

/// 两个路径是否指向同一个文件（先按字面，再按规范化路径）。
fn same_path(left: &str, right: &str) -> bool {
    if left.eq_ignore_ascii_case(right) {
        return true;
    }
    match (
        std::fs::canonicalize(left),
        std::fs::canonicalize(right),
    ) {
        (Ok(a), Ok(b)) => a == b,
        _ => false,
    }
}

fn backup_path(path: &str) -> String {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0);
    let mut candidate = format!("{path}.bak-{stamp}");
    let mut index = 1;
    while Path::new(&candidate).exists() {
        candidate = format!("{path}.bak-{stamp}-{index}");
        index += 1;
    }
    candidate
}
