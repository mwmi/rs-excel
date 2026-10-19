//! 面向 AI / 工具调用方的类型化原语。
//!
//! 这一层**只做接口与返回格式**，内部仍调用既有 COM 能力（Range/Workbook/Pivot/History）：
//! - [`Target`]：`book` / `sheet` / `range` 三字段解析（含旧 `表名!区域` 兼容）；
//! - [`CellInput`]：二维数组的自然类型（文本/数字/布尔/空/公式/日期）；
//! - [`Matrix`]：带预算与截断信息的读取结果；
//! - [`write_grid`]：一次批量写 + 一次数字格式补齐；
//! - [`FastGuard`]：快照并恢复应用设置；
//! - [`ToolError`]：结构化错误（code/category/hint/retryable/outcome）；
//! - [`save_as`]：显式覆盖策略与备份。
//!
//! 调用方负责 JSON 编解码；本模块不引入 serde，保持 COM 库轻量。

pub mod cell;
pub mod error;
pub mod guard;
pub mod read;
pub mod save;
pub mod target;
pub mod write;

pub use cell::{is_iso_date, serial_from_iso, validate_grid, CellInput, MAX_WRITE_CELLS};
pub use error::{AiResult, Category, Outcome, ToolError};
pub use guard::{fast, recovery_error, FastGuard};
pub use read::{Matrix, Projection, ReadWindow};
pub use save::{save_as, OverwritePolicy, SaveReport};
pub use target::{clean_addr, col_index, col_letters, pick_workbook, ParsedTarget, Resolved, Target};
pub use write::{write_grid, WritePath, WriteReport};
