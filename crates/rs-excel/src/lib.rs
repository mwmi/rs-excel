//! # rs-excel
//!
//! Rust + COM 驱动 WPS/Excel（Windows 专用），基于 `IDispatch` 后期绑定。
//!
//! ## 快速上手
//!
//! ```no_run
//! use rs_excel::prelude::*;
//!
//! let app = App::connect(Engine::Wps, ConnectMode::Attach)?; // 接管已打开的 WPS
//! let wb = app.workbook("Book1")?;                          // 按序号或名称绑定
//! let sheet = wb.active_sheet()?;
//! sheet.write_tsv("A1", "姓名\t数量\n苹果\t12\n香蕉\t34")?;   // 一段 TSV 写一块区域
//! let v = sheet.range("B2")?.value()?;                      // 读单元格
//! sheet.range("A1:B1")?.set_font(Some(true), None, None, None, None)?; // 加粗
//! ```
//!
//! ## 设计要点
//!
//! - **AI 友好**：所有错误为单行结构化格式
//!   `[操作] @ 对象路径 | 底层原因 | 提示: 修复建议`，一次读取即可定位。
//! - **批量读写**：`write_tsv` / `read_tsv` 一段文本搞定整块区域，最省 token。
//! - **扩展性**：库未封装的 COM 成员用 `.raw().get/call/call_named/put` 一行直调：
//!   ```no_run
//!   sheet.raw().call("PageSetup", vec![])?.idispatch()?.put("Zoom", vec![false.into()])?;
//!   ```
//! - **联动链路**：`range` → `font()` / `interior()` / `borders()` → 各自属性设置；
//!   `worksheet.page_setup()` → 页边距/纸张/方向/打印区域/缩放/页眉页脚/打印标题。
//! - **VBA 双通道**：简单操作走内置 API；复杂操作走 `Vba::run_macro`（无需设置）
//!   或 `Vba::execute`（注入临时模块，需开启信任访问）。
//! - **颜色格式**：COM 内部为 OLE 格式 `0x00BBGGRR`，统一用 [`rgb`] 助手构造。

pub mod com;
pub mod engine;
pub mod error;
pub mod pagesetup;
pub mod process;

mod app;
mod chart;
mod range;
mod style;
mod vba;
mod workbook;
mod worksheet;

pub mod prelude {
    pub use crate::app::{App, WindowInfo, WindowState};
    pub use crate::chart::{Chart, ChartKind};
    pub use crate::com::dispatch::IDispatchWrapper;
    pub use crate::com::variant::{CellValue, ComValue};
    pub use crate::engine::{ConnectMode, Engine};
    pub use crate::error::{rgb, ExcelError};
    pub use crate::pagesetup::{Margin, Orientation, PageSetup, PaperSize, Section};
    pub use crate::process::{kill, kill_all, processes, ProcessInfo};
    pub use crate::range::{AlignH, AlignV, Range};
    pub use crate::style::{Borders, Font, Interior};
    pub use crate::vba::{Vba, VbaStatus};
    pub use crate::workbook::{FileFormat, SheetInfo, Workbook, WorkbookInfo};
    pub use crate::worksheet::Worksheet;
    /// 供 `ComValue::Obj(...)` 传 COM 对象参数使用
    pub use windows::Win32::System::Com::IDispatch;
}

pub use app::{App, WindowInfo, WindowState};
pub use chart::{Chart, ChartKind};
pub use com::variant::{CellValue, ComValue};
pub use engine::{ConnectMode, Engine};
pub use error::{rgb, ExcelError};
pub use pagesetup::{Margin, Orientation, PageSetup, PaperSize, Section};
pub use process::{kill, kill_all, processes, ProcessInfo};
pub use range::{AlignH, AlignV, Range};
pub use style::{Borders, Font, Interior};
pub use vba::{Vba, VbaStatus};
pub use workbook::{FileFormat, SheetInfo, Workbook, WorkbookInfo};
pub use worksheet::Worksheet;
