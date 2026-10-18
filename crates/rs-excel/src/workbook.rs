//! 工作簿：信息、工作表管理、保存/另存/导出/关闭。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::sel_arg;
use crate::com::variant::{CellValue, ComValue};
use crate::error::{ole_to_hex, ExcelError};
use crate::worksheet::Worksheet;

/// 工作簿快照信息。
#[derive(Debug, Clone)]
pub struct WorkbookInfo {
    pub index: i32,
    pub name: String,
    pub full_name: String,
    pub path: String,
    pub saved: bool,
    pub sheet_count: i32,
    pub is_active: bool,
}

/// 工作表快照信息。
#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub index: i32,
    pub name: String,
    /// 标签页颜色 `#RRGGBB`，未设置时为空串
    pub tab_color_hex: String,
    /// visible / hidden / very_hidden
    pub visible: String,
}

/// 保存文件格式（xlFileFormat 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FileFormat {
    Xlsx = 51,
    Xlsm = 52,
    Xls = 56,
    Csv = 6,
}

/// 工作簿（Workbook COM 对象）。
pub struct Workbook {
    raw: IDispatchWrapper,
}

impl Workbook {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    pub fn name(&self) -> Result<String> {
        self.raw.get("Name")?.string()
    }

    pub fn full_name(&self) -> Result<String> {
        self.raw.get("FullName")?.string()
    }

    pub fn path(&self) -> Result<String> {
        self.raw.get("Path")?.string()
    }

    /// 是否无未保存改动。
    pub fn is_saved(&self) -> Result<bool> {
        self.raw.get("Saved")?.bool()
    }

    /// 是否只读打开。
    pub fn read_only(&self) -> Result<bool> {
        self.raw.get("ReadOnly")?.bool()
    }

    /// 文件格式原始码（51=xlsx，52=xlsm，56=xls，6=csv，见 [FileFormat]）。
    pub fn file_format(&self) -> Result<i32> {
        self.raw.get("FileFormat")?.int()
    }

    pub fn sheet_count(&self) -> Result<i32> {
        self.raw.get_obj("Worksheets", "Worksheets")?.get("Count")?.int()
    }

    /// 激活工作簿。
    pub fn activate(&self) -> Result<()> {
        self.raw.call("Activate", vec![]).map(|_| ())
    }

    /// 列出所有工作表（序号/名称/标签颜色/可见性）。
    pub fn sheets(&self) -> Result<Vec<SheetInfo>> {
        let sheets = self.raw.get_obj("Worksheets", "Worksheets")?;
        let count = sheets.get("Count")?.int()?;
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let sh = sheets
                .call("Item", vec![i.into()])?
                .idispatch_child(&format!("工作表[{i}]"))?;
            let name = sh.get("Name")?.string()?;
            let tab_color_hex = match sh.get("Tab").ok().and_then(|t| t.idispatch().ok()) {
                Some(tab) => match tab.get("Color").map(|v| v.value()) {
                    Ok(CellValue::Number(c)) => ole_to_hex(c as u32),
                    _ => String::new(),
                },
                None => String::new(),
            };
            let visible = match sh.get("Visible")?.value() {
                // xlSheetVisible=-1, xlSheetHidden=0, xlSheetVeryHidden=2
                CellValue::Number(0.0) => "hidden",
                CellValue::Number(2.0) => "very_hidden",
                _ => "visible",
            }
            .to_string();
            out.push(SheetInfo {
                index: i,
                name,
                tab_color_hex,
                visible,
            });
        }
        Ok(out)
    }

    /// 按序号（1 基，须在范围内，否则按名称）绑定工作表。
    pub fn sheet(&self, idx_or_name: &str) -> Result<Worksheet> {
        let sheets = self.raw.get_obj("Worksheets", "Worksheets")?;
        let count = sheets.get("Count")?.int()?;
        let raw = sheets
            .call("Item", vec![sel_arg(idx_or_name, count)])
            .and_then(|v| v.idispatch_child(idx_or_name))
            .map_err(|e| ExcelError::hint(e, "工作表序号或名称不存在：先用 sheets() 列出"))?;
        Ok(Worksheet::from_raw(raw))
    }

    /// 当前活动工作表。
    pub fn active_sheet(&self) -> Result<Worksheet> {
        Ok(Worksheet::from_raw(
            self.raw.get_obj("ActiveSheet", "ActiveSheet")?,
        ))
    }

    /// 新增工作表。`before` 为锚点工作表名：有值则插到它前面，None 则加到末尾。
    pub fn add_sheet(&self, name: &str, before: Option<&str>) -> Result<Worksheet> {
        let sheets = self.raw.get_obj("Worksheets", "Worksheets")?;
        let raw = match before {
            Some(anchor) => {
                let anchor_ws = self.sheet(anchor)?;
                sheets
                    .call("Add", vec![ComValue::Obj(anchor_ws.raw().disp.clone())])?
                    .idispatch_child(name)?
            }
            None => {
                // **不要用 `Add(Empty, After)` 这种省略参数占位**：WPS 的 COM 对 VT_EMPTY
                // 占位参数处理与 Excel 不一致，会直接报 0x800A03EC（应用层错误）。
                // 实报 2026-10-18：用户说「创建数据透视表」，sheet_add 连试两个名字都失败，
                // 模型只好改用「在原表右侧建透视」。零参数 Add 是唯一双方都稳的形态，
                // 代价是新表落在活动表之前（不影响功能，需要末尾位置由调用方 Move）
                match sheets.call("Add", vec![]) {
                    Ok(added) => added.idispatch_child(name)?,
                    Err(err) => {
                        // 另一种常见原因：工作簿结构受保护（ProtectStructure），
                        // 此时单元格能改、却加不了表 —— 现象与上面一模一样，必须分开告知
                        let protected = self
                            .raw
                            .get("ProtectStructure")
                            .map(|result| match result.value() {
                                CellValue::Bool(flag) => flag,
                                CellValue::Number(number) => number != 0.0,
                                _ => false,
                            })
                            .unwrap_or(false);
                        return Err(if protected {
                            ExcelError::hint(
                                err,
                                "工作簿结构受保护（ProtectStructure），加不了工作表：先取消保护，或直接建在已有表上",
                            )
                        } else {
                            ExcelError::hint(err, "新增工作表失败：确认工作簿未处于保护/共享状态")
                        }
                        .into());
                    }
                }
            }
        };
        let ws = Worksheet::from_raw(raw);
        ws.rename(name)?; // Add() 默认名改为指定名
        Ok(ws)
    }

    pub fn save(&self) -> Result<()> {
        self.raw.call("Save", vec![]).map(|_| ())
    }

    /// 另存为（绝对路径 + 格式）。
    ///
    /// 注意：目标文件已存在时会弹覆盖确认框（WPS 该弹窗不响应 DisplayAlerts）；
    /// 需免弹窗覆盖的场景由调用方先行删除目标文件。
    pub fn save_as(&self, path: &str, format: FileFormat) -> Result<()> {
        self.raw
            .call("SaveAs", vec![path.into(), (format as i32).into()])
            .map(|_| ())
            .map_err(|e| {
                ExcelError::hint(e, "路径需为绝对路径；同名文件若已打开或只读会失败")
            })
    }

    /// 导出为 PDF（绝对路径）。
    pub fn export_pdf(&self, path: &str) -> Result<()> {
        self.raw
            .call("ExportAsFixedFormat", vec![0i32.into(), path.into()]) // 0 = xlTypePDF
            .map(|_| ())
            .map_err(|e| ExcelError::hint(e, "路径需为绝对路径；确认工作簿可打印区域有效"))
    }

    /// 关闭工作簿。`save` 是否保存改动。
    pub fn close(&self, save: bool) -> Result<()> {
        self.raw.call("Close", vec![save.into()]).map(|_| ())
    }

    /// 保护工作簿结构（禁止增删/移动工作表）。
    pub fn protect(&self, password: &str) -> Result<()> {
        self.raw
            .call_named(
                "Protect",
                vec![],
                &[("Password", password.into()), ("Structure", true.into())],
            )
            .map(|_| ())
    }

    /// 取消结构保护。
    pub fn unprotect(&self, password: &str) -> Result<()> {
        self.raw.call("Unprotect", vec![password.into()]).map(|_| ())
    }

    /// 保存副本（备份首选，不影响当前文件与保存状态）。路径需为绝对路径。
    pub fn save_copy_as(&self, path: &str) -> Result<()> {
        self.raw
            .call("SaveCopyAs", vec![path.into()])
            .map(|_| ())
            .map_err(|e| ExcelError::hint(e, "路径需为绝对路径；确认目录存在且文件未被独占锁定"))
    }
}
