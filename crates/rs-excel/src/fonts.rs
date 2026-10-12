//! 系统字体名枚举（GDI）：引擎无关（WPS 不支持 `Application.FontNames`），
//! 供字体名补全等场景使用。

use windows::Win32::Foundation::LPARAM;
use windows::Win32::Graphics::Gdi::{
    EnumFontFamiliesExW, GetDC, ReleaseDC, LOGFONTW, TEXTMETRICW, DEFAULT_CHARSET,
};

/// 枚举系统已安装字体名（去重、排序）。纯 Win32，不依赖 COM 连接。
pub fn system_font_names() -> Vec<String> {
    unsafe {
        let hdc = GetDC(None);
        let mut names: Vec<String> = Vec::new();
        let lf = LOGFONTW {
            lfCharSet: DEFAULT_CHARSET,
            ..Default::default()
        };
        EnumFontFamiliesExW(
            hdc,
            &lf,
            Some(font_proc),
            LPARAM(&mut names as *mut Vec<String> as isize),
            0,
        );
        let _ = ReleaseDC(None, hdc);
        names.retain(|s| !s.starts_with('@')); // 去掉竖排变体（GDI 对 CJK 字体的重复登记）
        names.sort();
        names.dedup();
        names
    }
}

/// 回调：收集字体面名（按 charset 会重复枚举同一字体，调用方去重）。
unsafe extern "system" fn font_proc(
    lpelfe: *const LOGFONTW,
    _lpntme: *const TEXTMETRICW,
    _font_type: u32,
    lparam: LPARAM,
) -> i32 {
    let names = &mut *(lparam.0 as *mut Vec<String>);
    let face = &(*lpelfe).lfFaceName;
    let len = face.iter().position(|&c| c == 0).unwrap_or(face.len());
    let s = String::from_utf16_lossy(&face[..len]);
    if !s.is_empty() {
        names.push(s);
    }
    1 // 继续枚举
}
