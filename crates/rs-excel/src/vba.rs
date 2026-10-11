//! VBA：能力检测、运行已有宏、注入临时模块执行代码。
//!
//! 双通道策略（AI 自动化基础保障）：
//! - `run_macro`：`Application.Run` 运行**已存在于打开工作簿**的宏，无需任何设置；
//! - `execute`：注入临时标准模块执行任意 VBA 代码，需开启
//!   「信任中心 → 宏设置 → 信任对 VBA 工程对象模型的访问」。

use anyhow::{bail, Result};
use windows::Win32::System::Variant::VARIANT;

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::ExcelError;

/// VBA 能力状态。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct VbaStatus {
    /// VBA 环境是否安装（WPS 需 VBA 支持库）
    pub env_installed: bool,
    /// 是否开启「信任对 VBA 工程对象模型的访问」
    pub trust_access: bool,
}

/// VBA 操作入口。
pub struct Vba {
    app: IDispatchWrapper,
}

impl Vba {
    pub fn new(app: &App) -> Self {
        Self {
            app: app.raw().clone(),
        }
    }

    /// 检测 VBA 能力。
    pub fn status(&self) -> VbaStatus {
        let vbe = self.app.get("VBE").ok().and_then(|v| v.idispatch().ok());
        let env_installed = vbe.is_some();
        let trust_access = vbe
            .map(|v| v.get("ActiveVBProject").is_ok())
            .unwrap_or(false);
        VbaStatus {
            env_installed,
            trust_access,
        }
    }

    /// 运行已存在于打开工作簿中的宏，如 `run_macro("模块1.Main", &[123.into()])`。
    pub fn run_macro(&self, name: &str, args: &[ComValue]) -> Result<()> {
        let mut a = vec![name.into()];
        a.extend(args.iter().cloned());
        self.app
            .call("Run", a)
            .map(|_| ())
            .map_err(|e| {
                ExcelError::hint(
                    e,
                    "宏必须已存在于打开的工作簿中；带模块时用 '模块名.过程名'",
                )
            })
    }

    /// 注入临时标准模块执行 VBA 代码，执行后自动清理。
    ///
    /// - `code`：完整 VBA 代码（须含公共 Sub）
    /// - `entry`：入口过程名（如 "Main"，代码里须有 `Public Sub Main()`）
    pub fn execute(&self, code: &str, entry: &str) -> Result<()> {
        let st = self.status();
        if !st.env_installed {
            bail!(ExcelError::new("vba.execute", "VBA", "VBA 环境未安装")
                .with_hint("WPS 需先安装 VBA 支持库，或使用 MS Excel"));
        }
        if !st.trust_access {
            bail!(ExcelError::new(
                "vba.execute",
                "VBA",
                "未开启「信任对 VBA 工程对象模型的访问」"
            )
            .with_hint("在 WPS/Excel 信任中心 → 宏设置 中勾选「信任对 VBA 工程对象模型的访问」后重试"));
        }
        let vbe = self.app.get("VBE")?.idispatch()?;
        let project = vbe.get("ActiveVBProject")?.idispatch()?;
        let comps = project.get_obj("VBComponents", "VBComponents")?;
        // 1 = vbext_ct_StdModule（标准模块）
        let comp = comps
            .call("Add", vec![1i32.into()])?
            .idispatch_child("临时模块")?;
        let mod_name = format!(
            "RustTmp{}",
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.subsec_nanos())
                .unwrap_or(0)
        );
        comp.put("Name", vec![mod_name.into()])?;
        comp.get_obj("CodeModule", "CodeModule")?
            .call("AddFromString", vec![code.into()])
            .map_err(|e| {
                ExcelError::hint(e, "VBA 代码语法错误或包含不允许的语句")
            })?;
        // Run 后无论成败都移除临时模块
        let run_res = self.app.call("Run", vec![entry.into()]).map(|_| ());
        let _ = comps.call("Remove", vec![ComValue::Obj(comp.disp.clone())]);
        run_res.map_err(|e| {
            ExcelError::hint(
                e,
                format!("宏运行失败：确认入口过程名 '{entry}' 与代码中的公共 Sub 名一致"),
            )
        })
    }

    /// 直接调用 `Application.Run`（返回 VARIANT，供需要返回值的高级用法）。
    pub fn run_raw(&self, name: &str, args: &[ComValue]) -> Result<VARIANT> {
        let mut a = vec![name.into()];
        a.extend(args.iter().cloned());
        Ok(self.app.call("Run", a)?.v)
    }
}
