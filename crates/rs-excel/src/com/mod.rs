//! COM 底层：IDispatch 调用、值转换、SAFEARRAY 批量读写。

pub mod dispatch;
pub mod safearray;
pub mod variant;

/// 窗口「序号或标题」选择器：`1..=count` 内的纯数字解析为序号，否则 None（按标题处理）。
pub(crate) fn sel_index(sel: &str, count: i32) -> Option<i32> {
    sel.parse::<i32>()
        .ok()
        .filter(|n| (1..=count).contains(n))
}
