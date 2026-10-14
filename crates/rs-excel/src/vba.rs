//! VBA：能力检测、运行已有宏、注入临时模块执行代码。
//!
//! 双通道策略（AI 自动化基础保障）：
//! - `run_macro`：`Application.Run` 运行**已存在于打开工作簿**的宏，无需任何设置；
//! - `execute`：注入临时标准模块执行任意 VBA 代码，需开启
//!   「信任中心 → 宏设置 → 信任对 VBA 工程对象模型的访问」。
//!
//! **运行时错误不弹窗**：`execute` 把入口包进一个 `On Error GoTo` 的包装 Function，
//! 未处理的运行时错误会被抓住、写成返回值带出来（否则 WPS/Excel 会弹模态框，
//! 整条 COM 链路卡到用户手动关窗为止，调用方什么错误信息都拿不到）。

use anyhow::{bail, Result};
use windows::Win32::System::Com::IDispatch;
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

/// 注入 / 执行期间藏起 VBE 主窗口（退出作用域恢复原状）。
/// 编译错误时 VBE 会自行弹模态错误框、还把编辑器窗口打开 —— 自动化链路里这就是
/// 「莫名弹出的报错窗口」。拿不到 VBE（没装支持库）时返回 None，无可藏也无从弹
struct VbeHidden {
    window: IDispatchWrapper,
    shown: bool,
}

impl Drop for VbeHidden {
    fn drop(&mut self) {
        let _ = self.window.put("Visible", vec![self.shown.into()]);
    }
}

fn hide_vbe(app: &IDispatchWrapper) -> Option<VbeHidden> {
    // Application.VBE.MainWindow：无条件藏（编译错误会把藏着的 VBE 也弹出来），原状记下来供恢复
    let window = app.get("VBE").ok()?.idispatch().ok()?.get("MainWindow").ok()?.idispatch().ok()?;
    let shown = window.get("Visible").ok().and_then(|v| v.bool().ok()).unwrap_or(false);
    window.put("Visible", vec![false.into()]).ok()?;
    Some(VbeHidden { window, shown })
}

/// 清掉上次残留的临时模块（只删 `RustTmp` 前缀的 —— 那是我们自己的，用户模块一概不碰）。
/// 注入失败 / 弹窗卡死被强杀时模块会留在用户工程里，可能带着坏代码；而
/// `Application.Run` 会编译**整个工程**，之后每次调用都弹一次「编译错误」
/// （实报：点了确定下次还弹、再点又能跑通 —— 坏模块在旁边，我们的代码是好的）
fn drop_stale_modules(comps: &IDispatchWrapper) {
    let Ok(count) = comps.get("Count").and_then(|v| v.int()) else { return };
    // 倒着删：移除会让后面的序号前移，从尾部开始不影响还没查的
    for index in (1..=count).rev() {
        let Ok(comp) = comps.call("Item", vec![ComValue::Int(index)]).and_then(|v| v.idispatch()) else { continue };
        let name = comp.get("Name").ok().and_then(|v| v.string().ok()).unwrap_or_default();
        if name.starts_with("RustTmp") {
            let _ = comps.call("Remove", vec![ComValue::Obj(comp.disp.clone())]);
        }
    }
}

/// 临时模块的 RAII 守卫：退出作用域（含 AddFromString 失败、提前 return）一律移除，
/// 不在用户工程里留模块 —— 残留的坏模块会让之后每次运行都弹编译错误
struct TempModule<'a> {
    comps: &'a IDispatchWrapper,
    disp: IDispatch,
}

impl Drop for TempModule<'_> {
    fn drop(&mut self) {
        let _ = self.comps.call("Remove", vec![ComValue::Obj(self.disp.clone())]);
    }
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
    ///
    /// **运行时错误不弹窗**：入口外面包一层 `On Error GoTo` 的包装 Function，
    /// 错误被抓住后写成返回值带出来（`Application.Run` 调 Function 能拿到返回值），
    /// 由调用方转成文字错误交给上层（AI 据此改码重发）。用户代码内部的错误处理不受影响。
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
        // 先清掉上次残留的临时模块：注入失败 / 弹窗卡死被强杀时模块会留在用户工程里，
        // 且可能带着坏代码 —— 而 Application.Run 会编译整个工程，之后每次调用都弹一次
        // 「编译错误」（实报：点了确定下次还弹、再点又能跑通）。只删 RustTmp 前缀的
        drop_stale_modules(&comps);
        // 注入与执行期间把 VBE 窗口藏起来：编译错误（入口名写错、代码里有未定义的引用）时
        // VBE 会自己弹模态「编译错误」并把编辑器窗口打开（实报截图），藏起来挡掉这类窗口，
        // 退出作用域按原状恢复（用户本来开着 VBE 的话不动它）
        let _vbe_hidden = hide_vbe(&self.app);
        // 1 = vbext_ct_StdModule（标准模块）
        let comp = comps
            .call("Add", vec![1i32.into()])?
            .idispatch_child("临时模块")?;
        // 守卫：任何退出路径（含 AddFromString 失败）都带走模块，不在用户工程里留垃圾
        let _temp = TempModule { comps: &comps, disp: comp.disp.clone() };
        // 同一时刻只可能有一次注入：模块名与包装过程共用同一个尾号
        let tag = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.subsec_nanos())
            .unwrap_or(0);
        let mod_name = format!("RustTmp{tag}");
        comp.put("Name", vec![mod_name.into()])?;
        // 包装 Function：抓住入口过程未处理的运行时错误，写成返回值带出来
        let wrapper = format!("RustTmpRun{tag}");
        let trapped = format!(
            "{code}\n\nPublic Function {wrapper}() As String\n\
             \x20   On Error GoTo VbaErr\n\
             \x20   {entry}\n\
             \x20   Exit Function\n\
             VbaErr:\n\
             \x20   {wrapper} = \"VBA 运行时错误 \" & Err.Number & \"：\" & Err.Description\n\
             End Function"
        );
        comp.get_obj("CodeModule", "CodeModule")?
            .call("AddFromString", vec![trapped.into()])
            .map_err(|e| {
                ExcelError::hint(e, "VBA 代码语法错误或包含不允许的语句")
            })?;
        let run_res = self.app.call("Run", vec![wrapper.into()]);
        match run_res {
            // 包装 Function 的返回值：空 = 正常跑完；非空 = 抓到的运行时错误原文
            Ok(result) => {
                let message = result.string().unwrap_or_default();
                if message.trim().is_empty() {
                    Ok(())
                } else {
                    Err(ExcelError::new("vba.execute", "VBA", message)
                        .with_hint("按错误号与描述修正 code 后重发；已成功的操作不会重复执行")
                        .into())
                }
            }
            Err(e) => Err(ExcelError::hint(
                e,
                format!("宏运行失败：确认入口过程名 '{entry}' 与代码中的公共 Sub 名一致"),
            )
            .into()),
        }
    }

    /// 直接调用 `Application.Run`（返回 VARIANT，供需要返回值的高级用法）。
    pub fn run_raw(&self, name: &str, args: &[ComValue]) -> Result<VARIANT> {
        let mut a = vec![name.into()];
        a.extend(args.iter().cloned());
        Ok(self.app.call("Run", a)?.v)
    }
}
