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
//! - **最快路径默认**：值/公式读写默认走 `Value2` / `Formula2`（日期返回序列值）；
//!   旧 `Value` / `Formula` 语义（日期转可读文本）用 `_compat` 后缀方法访问。
//! - **批量提速**：批量操作前 `app.set_fast(true)`（四合一批量模式），结束 `set_fast(false)` 恢复。
//! - **扩展性**：库未封装的 COM 成员用 `.raw().get/call/call_named/put` 一行直调：
//!   ```no_run
//!   sheet.raw().call("PageSetup", vec![])?.idispatch()?.put("Zoom", vec![false.into()])?;
//!   ```
//! - **联动链路**：`range` → `font()` / `interior()` / `borders()` → 各自属性设置；
//!   `range.format_conditions()` 或 `cond_value` / `cond_formula` / `cond_text` /
//!   `cond_color_scale` / `cond_data_bar` / `cond_icon` 等一次性方法 → 条件格式规则增删查；
//!   `worksheet.page_setup()` → 页边距/纸张/方向/打印区域/缩放/页眉页脚/打印标题；
//!   `worksheet.shapes()` → 形状/图片/文本框的增删改查、外观、文本与批量操作。
//! - **VBA 双通道**：简单操作走内置 API；复杂操作走 `Vba::run_macro`（无需设置）
//!   或 `Vba::execute`（注入临时模块，需开启信任访问）。
//! - **操作撤销**：经 COM 做的修改不进应用撤销栈（`Application.Undo` 无效），
//!   用 [`History`] 在修改前记录快照、事后撤销（默认 50 条，后进先出）：
//!   ```no_run
//!   # let hist = rs_excel::History::new();
//!   # let (sheet, wb) = unreachable!();
//!   # let addr = "A1";
//!   hist.record(&sheet, addr, "写入前")?; // 修改前记录
//!   // ... 执行修改 ...
//!   let desc = hist.undo(&wb)?;           // 撤销最近一次，返回单行描述
//!   ```
//! - **事件订阅**：[`EventMonitor`] 经 COM 连接点订阅 Application 级事件
//!   （工作簿开关/激活、工作表激活、选区变化、单元格修改等），原生回调推送
//!   零轮询零延迟；[`App::state`] 按需读取单行状态快照：
//!   ```no_run
//!   # let mut app = unreachable!();
//!   let mon = EventMonitor::start(&app)?;
//!   for ev in mon.poll() { println!("{ev}"); } // 如 "选区变化: Book1.xlsx > Sheet3!B2:C5"
//!   println!("{}", app.state()?);              // "Book1.xlsx | Sheet3 | 选区 B2:C5 | ..."
//!   ```
//! - **颜色格式**：COM 内部为 OLE 格式 `0x00BBGGRR`，统一用 [`rgb`] 助手构造。

pub mod com;
pub mod condfmt;
pub mod engine;
pub mod error;
pub mod events;
pub mod fonts;
pub mod history;
pub mod pagesetup;
pub mod process;
pub mod shape;

mod app;
mod chart;
mod range;
mod style;
mod vba;
mod workbook;
mod worksheet;

pub mod prelude {
    pub use crate::app::{App, CalcMode, WindowInfo, WindowState};
    pub use crate::chart::{Chart, ChartKind};
    pub use crate::condfmt::{CondFormat, CondOp, CondText, DateOp, FormatCondition, FormatConditions, IconSet};
    pub use crate::com::dispatch::IDispatchWrapper;
    pub use crate::com::variant::{CellValue, ComValue};
    pub use crate::engine::{ConnectMode, Engine};
    pub use crate::error::{rgb, ExcelError};
    pub use crate::events::{AppState, EventMonitor, ExcelEvent};
    pub use crate::fonts::system_font_names;
    pub use crate::history::{History, UndoEntry};
    pub use crate::pagesetup::{Margin, Orientation, PageSetup, PaperSize, Section};
    pub use crate::process::{bring_to_front, flash_window, kill, kill_all, processes, restore_if_minimized, ProcessInfo};
    pub use crate::range::{AlignH, AlignV, EndDir, Range, SpecialCells};
    pub use crate::shape::{
        type_name, AlignKind, AutoShapeKind, ConnectorKind, DashStyle, DistributeKind, FlipDir,
        Placement, Shape, ShapeInfo, ShapeRange, Shapes, ZOrderCmd,
    };
    pub use crate::style::{Borders, Font, Interior};
    pub use crate::vba::{Vba, VbaStatus};
    pub use crate::workbook::{FileFormat, SheetInfo, Workbook, WorkbookInfo};
    pub use crate::worksheet::Worksheet;
    /// 供 `ComValue::Obj(...)` 传 COM 对象参数使用
    pub use windows::Win32::System::Com::IDispatch;
}

pub use app::{App, CalcMode, WindowInfo, WindowState};
pub use chart::{Chart, ChartKind};
pub use com::variant::{CellValue, ComValue};
pub use condfmt::{CondFormat, CondOp, CondText, DateOp, FormatCondition, FormatConditions, IconSet};
pub use engine::{ConnectMode, Engine};
pub use error::{rgb, ExcelError};
pub use events::{AppState, EventMonitor, ExcelEvent};
pub use fonts::system_font_names;
pub use history::{History, UndoEntry};
pub use pagesetup::{Margin, Orientation, PageSetup, PaperSize, Section};
pub use process::{bring_to_front, flash_window, kill, kill_all, processes, restore_if_minimized, ProcessInfo};
pub use range::{AlignH, AlignV, EndDir, Range, SpecialCells};
pub use shape::{
    type_name, AlignKind, AutoShapeKind, ConnectorKind, DashStyle, DistributeKind, FlipDir,
    Placement, Shape, ShapeInfo, ShapeRange, Shapes, ZOrderCmd,
};
pub use style::{Borders, Font, Interior};
pub use vba::{Vba, VbaStatus};
pub use workbook::{FileFormat, SheetInfo, Workbook, WorkbookInfo};
pub use worksheet::Worksheet;
