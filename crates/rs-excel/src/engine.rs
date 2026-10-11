//! 引擎选择（WPS/Excel）、连接模式与 COM 套间守卫。

use anyhow::Result;
use windows::Win32::System::Com::{CoInitialize, CoUninitialize};

/// 驱动引擎：默认 WPS。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Engine {
    Wps,
    Excel,
}

impl Engine {
    /// COM ProgID。WPS 表格 = `KET.Application`，Excel = `Excel.Application`。
    pub fn prog_id(self) -> &'static str {
        match self {
            Self::Wps => "KET.Application",
            Self::Excel => "Excel.Application",
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::Wps => "WPS",
            Self::Excel => "Excel",
        }
    }

    /// 应用可执行文件名（进程查看/结束时匹配用）。
    pub fn exe_name(self) -> &'static str {
        match self {
            Self::Wps => "et.exe",
            Self::Excel => "excel.exe",
        }
    }
}

/// 连接模式：新建应用实例 / 接管已运行实例（ROT）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ConnectMode {
    New,
    Attach,
}

/// COM STA 套间守卫：`App` 内部持有，Drop 时 `CoUninitialize`，
/// 保证任何出错路径都不会泄漏 COM 套间。
pub(crate) struct ComApartment;

impl ComApartment {
    pub(crate) fn init() -> Result<Self> {
        unsafe { CoInitialize(None).ok()? };
        Ok(Self)
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}
