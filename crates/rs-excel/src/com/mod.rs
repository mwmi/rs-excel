//! COM 底层：IDispatch 调用、值转换、SAFEARRAY 批量读写。

pub mod dispatch;
pub mod safearray;
pub mod variant;

use crate::com::variant::ComValue;

/// 「索引或名称」选择器：`1..=count` 内的纯数字解析为索引，否则 None（按名称/标题处理）。
///
/// 纯数字既可能是索引也可能是名称（如 `addsheet 2333` 后按名切换）：
/// 范围内优先按索引（与补全菜单插入的索引语义一致），范围外按名称查找。
pub(crate) fn sel_index(sel: &str, count: i32) -> Option<i32> {
    sel.parse::<i32>()
        .ok()
        .filter(|n| (1..=count).contains(n))
}

/// 同 [sel_index]，直接产出 COM 参数（索引或原字符串名称）。
pub(crate) fn sel_arg(sel: &str, count: i32) -> ComValue {
    match sel_index(sel, count) {
        Some(n) => n.into(),
        None => sel.to_string().into(),
    }
}
