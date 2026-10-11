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
                let count = sheets.get("Count")?.int()?;
                let last = sheets
                    .call("Item", vec![count.into()])?
                    .idispatch()?
                    .disp
                    .clone();
                // Add(Before=省略, After=最后一张)
                sheets
                    .call("Add", vec![ComValue::Empty, ComValue::Obj(last)])?
                    .idispatch_child(name)?
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
}
