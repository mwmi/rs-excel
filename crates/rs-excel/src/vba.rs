//! VBA：能力检测、运行已有宏、注入临时模块执行代码。
//!
//! 双通道策略（AI 自动化基础保障）：
//! - `run_macro`：`Application.Run` 运行**已存在于打开工作簿**的宏，无需任何设置；
//! - `execute`：注入临时标准模块执行任意 VBA 代码，需开启
//!   「信任中心 → 宏设置 → 信任对 VBA 工程对象模型的访问」。
//!
//! **错误不弹窗、不挂死、不假成功**（三层，全是实测出来的，2026-10-14 WPS）：
//!   ① 运行时错误：`execute` 把入口包进 `On Error GoTo` 的包装 Function，未处理的
//!      运行时错误被抓住、写成返回值带出来（否则 WPS/Excel 弹模态框，整条 COM 链路
//!      卡到用户手动关窗为止）；
//!   ② 编译错误：`Application.Run` 卡在模态「编译错误」框上**永远不返回**（无人点击
//!      实测 60s+ 不返回），且会把 VBE 编辑器窗口强行显示出来（`MainWindow.Visible=false`
//!      藏不住，属性还读回 false）—— 由 [VbaGuard] 看门狗在 Run 期间按**窗口类名**
//!      处理：模态框露面等 400ms 再点掉（并把框里的错误原文带回来），VBE 窗口露面**最多
//!      藏两次**（循环藏是在跟 VBA 抢窗口，实测把 et 进程整个搞挂，见常量注释）；
//!   ③ 编译错误被点掉后 Run **自己不当回事**（不报错、只回个空值）—— 当成成功就是
//!      「假执行成功」，模型以为改好了其实一个格子都没动（用户实报）。两个判据都要用：
//!      Run 期间**出现过 VBA 错误框**（看门狗点掉它、带回框里原文），或返回值**不是文本**
//!      —— 包装 Function 声明的是 `As String`，真跑起来必然是 VT_BSTR（空串也算），
//!      回 VT_EMPTY 就是它压根没执行（2026-10-15 探针实测：WPS 有时既不弹框也不抛错）；
//!   ④ 断点救援：编译错误的确定框被点掉后 WPS **并不返回 Run**，而是直接进**中断态**
//!     （VBE 标题带「[中断]」，窗口被藏、工程停在断点上，Run 永远阻塞 —— 一切 Application
//!      级 COM 调用都被拒，整个应用点不动，用户实报）。Run 期间由旁路救援线程盯 VBE 标题，
//!      发现中断态就从**第二条 COM 连接**执行 VBE「运行 > 重置」，阻塞中的 Run 才能返回，
//!      见 [spawn_break_rescuer]（主线程堵在 Run 里，本线程的任何收尾都执行不到）。

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::thread::sleep;
use std::time::{Duration, Instant};

use anyhow::{bail, Result};
use windows::Win32::Foundation::{HWND, LPARAM};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, EnumWindows, GetClassNameW, GetWindowTextW, IsWindowVisible, SendMessageW,
    ShowWindow, BM_CLICK, SW_HIDE, SW_SHOW,
};

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::{CellValue, ComValue};
use crate::engine::{ConnectMode, Engine};
use crate::error::ExcelError;

/// VBA 能力状态。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct VbaStatus {
    /// VBA 环境是否安装（WPS 需 VBA 支持库）
    pub env_installed: bool,
    /// 是否开启「信任对 VBA 工程对象模型的访问」
    pub trust_access: bool,
}

// ===== Run 期间的外观守卫 =====

/// VBE 主窗口的窗口类名（WPS/Excel 通用）。**只能按类名认窗口** —— 那个模态错误框的
/// 标题也含 "Visual Basic"，按标题认会把模态框一起藏了，WPS 卡死得更彻底（探针踩过：
/// 藏了的模态框还在跑模态循环，整个软件的 COM 调用全部被拒）
const VBE_CLASS: &str = "wndclass_desked_gsk";
/// VBA 模态错误框：标准对话框类 + 标题含 Visual Basic（MsgBox 标题不是它，不会误伤）
const DIALOG_CLASS: &str = "#32770";
/// 看门狗轮询间隔
const WATCH_TICK: Duration = Duration::from_millis(80);
/// 对话框露面多久才去点：露面瞬间就点 WPS 接不住（点击后进程挂死，实测），
/// 等它初始化完再点，点完还要验一次真关了、没关就重点
const CLICK_DELAY: Duration = Duration::from_millis(400);
/// 每个对话框最多点几次（前两次点「确定」按钮，最后一次对框发 WM_CLOSE 兜底）
const MAX_CLICKS: u8 = 3;
/// 同一个 VBE 窗口最多藏几次：露面藏一次 + 对话框点掉后补一次，收尾再补一刀。
/// **绝不能循环藏** —— 2026-10-15 实测：每轮都 SW_HIDE 露头的 VBE 是在跟 VBA 抢窗口，
/// et 进程整个挂死（Run 卡 51 秒后 RPC 失败、之后 COM 全不可用），比断点状态糟得多
const MAX_HIDES_PER_WINDOW: usize = 2;

/// 扫描到的窗口（hwnd, 类名, 标题）：只收 VBE 与模态框两类，别的不进锁
static FOUND: Mutex<Vec<(isize, String, String)>> = Mutex::new(Vec::new());
/// 模态框子窗口里的静态文本（错误原文）
static STATIC_TEXT: Mutex<Vec<String>> = Mutex::new(Vec::new());

unsafe extern "system" fn scan_top(hwnd: HWND, _: LPARAM) -> windows::core::BOOL {
    let mut cls = [0u16; 64];
    let m = GetClassNameW(hwnd, &mut cls);
    let class = String::from_utf16_lossy(&cls[..m as usize]);
    if class == VBE_CLASS || class == DIALOG_CLASS {
        let mut title = [0u16; 256];
        let n = GetWindowTextW(hwnd, &mut title);
        let text = String::from_utf16_lossy(&title[..n as usize]);
        FOUND.lock().unwrap().push((hwnd.0 as isize, class, text));
    }
    windows::core::BOOL(1)
}

unsafe extern "system" fn scan_static(hwnd: HWND, _: LPARAM) -> windows::core::BOOL {
    let mut cls = [0u16; 64];
    let m = GetClassNameW(hwnd, &mut cls);
    let class = String::from_utf16_lossy(&cls[..m as usize]);
    if class == "Static" {
        let mut title = [0u16; 512];
        let n = GetWindowTextW(hwnd, &mut title);
        let text = String::from_utf16_lossy(&title[..n as usize]).replace(['\r', '\n'], " ");
        if !text.trim().is_empty() {
            STATIC_TEXT.lock().unwrap().push(text);
        }
    }
    windows::core::BOOL(1)
}

unsafe extern "system" fn scan_button(hwnd: HWND, _: LPARAM) -> windows::core::BOOL {
    let mut cls = [0u16; 64];
    let m = GetClassNameW(hwnd, &mut cls);
    let class = String::from_utf16_lossy(&cls[..m as usize]);
    if class == "Button" {
        let mut title = [0u16; 64];
        let n = GetWindowTextW(hwnd, &mut title);
        let text = String::from_utf16_lossy(&title[..n as usize]);
        BUTTONS.lock().unwrap().push((hwnd.0 as isize, text));
    }
    windows::core::BOOL(1)
}

static BUTTONS: Mutex<Vec<(isize, String)>> = Mutex::new(Vec::new());

fn windows_now() -> Vec<(isize, String, String)> {
    FOUND.lock().unwrap().clear();
    unsafe { let _ = EnumWindows(Some(scan_top), LPARAM(0)); }
    FOUND.lock().unwrap().drain(..).collect()
}

fn children(dialog: isize, cb: unsafe extern "system" fn(HWND, LPARAM) -> windows::core::BOOL) {
    unsafe { let _ = EnumChildWindows(Some(HWND(dialog as *mut core::ffi::c_void)), Some(cb), LPARAM(0)); }
}

fn visible(hwnd: isize) -> bool {
    unsafe { IsWindowVisible(HWND(hwnd as *mut core::ffi::c_void)) }.as_bool()
}

fn hide(hwnd: isize) {
    unsafe { let _ = ShowWindow(HWND(hwnd as *mut core::ffi::c_void), SW_HIDE); }
}

fn show(hwnd: isize) {
    unsafe { let _ = ShowWindow(HWND(hwnd as *mut core::ffi::c_void), SW_SHOW); }
}

/// 读模态框里的错误原文（静态文本，如「编译错误： 方法或数据成员未找到」）
fn dialog_text(dialog: isize) -> String {
    STATIC_TEXT.lock().unwrap().clear();
    children(dialog, scan_static);
    STATIC_TEXT.lock().unwrap().drain(..).collect::<Vec<_>>().join(" ")
}

/// 点掉模态框：非最后一次点「确定/OK/是」按钮，最后一次对框发 WM_CLOSE 兜底。
/// 不点它，`Application.Run` 永远不返回
fn dismiss_dialog(dialog: isize, last_try: bool) {
    if !last_try {
        BUTTONS.lock().unwrap().clear();
        children(dialog, scan_button);
        let buttons: Vec<(isize, String)> = BUTTONS.lock().unwrap().drain(..).collect();
        for (hwnd, text) in buttons {
            let label = text.trim();
            if label == "确定" || label == "OK" || label == "是" {
                unsafe { SendMessageW(HWND(hwnd as *mut core::ffi::c_void), BM_CLICK, None, None) };
                return;
            }
        }
    }
    unsafe { SendMessageW(HWND(dialog as *mut core::ffi::c_void), 0x0010, None, None) }; // WM_CLOSE
}

/// Run 期间的外观守卫（RAII）：VBA 模态错误框露面 → 等 [CLICK_DELAY] → 点掉 → 验证；
/// VBE 窗口露面 → 最多藏 [MAX_HIDES_PER_WINDOW] 次（**不跟 VBA 抢窗口**，见常量注释）；
/// `finish` 时收尾补藏一次、恢复用户**原本**开着的 VBE 窗口，交回捕获到的错误原文。
///
/// 为什么必须看门狗：编译错误时这些窗口是在 `Application.Run` **执行期间**弹出来的，
/// 主线程堵在 COM 调用里出不来，只能从旁线程处理。
pub struct VbaGuard {
    stop: Arc<AtomicBool>,
    caught: Arc<Mutex<String>>,
    handle: Option<std::thread::JoinHandle<Vec<isize>>>,
}

impl VbaGuard {
    pub fn start() -> Self {
        let stop = Arc::new(AtomicBool::new(false));
        let caught = Arc::new(Mutex::new(String::new()));
        let handle = {
            let stop = stop.clone();
            let caught = caught.clone();
            std::thread::spawn(move || {
                // 第一轮就可见的 VBE 窗口 = 用户自己开着的，finish 时还给人家；
                // 之后才出现的（编译错误定位到行弹出的）不属于用户，藏到底
                let mut keep: Vec<isize> = Vec::new();
                // (hwnd, 首次露面时刻, 已点次数)；关掉了就从表里去掉
                let mut dialogs: Vec<(isize, Instant, u8)> = Vec::new();
                // 每个 VBE 窗口已藏过几次（上限 MAX_HIDES_PER_WINDOW）
                let mut hides: Vec<isize> = Vec::new();
                let mut first = true;
                while !stop.load(Ordering::SeqCst) {
                    let now = windows_now();
                    let present: Vec<isize> = now.iter().map(|(hwnd, _, _)| *hwnd).collect();
                    for (hwnd, class, title) in &now {
                        if *class == VBE_CLASS {
                            if first && visible(*hwnd) {
                                keep.push(*hwnd);
                            }
                        } else if *class == DIALOG_CLASS
                            && title.contains("Visual Basic")
                            && !dialogs.iter().any(|(seen, _, _)| seen == hwnd)
                        {
                            dialogs.push((*hwnd, Instant::now(), 0));
                        }
                    }
                    first = false;
                    // 对话框：露面满 CLICK_DELAY 才点；点前把原文读走；点完等 500ms 验一次
                    for (hwnd, seen, tries) in dialogs.iter_mut() {
                        if *tries >= MAX_CLICKS || seen.elapsed() < CLICK_DELAY {
                            continue;
                        }
                        if !present.contains(hwnd) {
                            continue; // 已经关了
                        }
                        // 点之前把原文读走；读不到也留个占位 —— 「有框」本身就是编译错误的信号
                        let text = dialog_text(*hwnd);
                        let note = if text.is_empty() {
                            "（VBA 错误对话框，未读到原文）".to_string()
                        } else {
                            text
                        };
                        if let Ok(mut slot) = caught.lock() {
                            if slot.is_empty() {
                                *slot = note;
                            }
                        }
                        dismiss_dialog(*hwnd, *tries + 1 >= MAX_CLICKS);
                        *tries += 1;
                        sleep(Duration::from_millis(500)); // 等它关，顺带别重复点
                    }
                    dialogs.retain(|(hwnd, _, tries)| *tries < MAX_CLICKS && present.contains(hwnd));
                    // VBE 露头就藏，但有次数上限 —— 循环藏会跟 VBA 抢窗口把进程搞挂
                    for (hwnd, class, _) in &now {
                        if *class == VBE_CLASS && visible(*hwnd) {
                            let times = hides.iter().filter(|hidden| *hidden == hwnd).count();
                            if times < MAX_HIDES_PER_WINDOW {
                                hides.push(*hwnd);
                                hide(*hwnd);
                            }
                        }
                    }
                    sleep(WATCH_TICK);
                }
                keep
            })
        };
        Self { stop, caught, handle: Some(handle) }
    }

    /// 停看门狗；Run 已返回，此时补藏一次 VBE（没人跟它抢了）、恢复用户原本开着的
    /// VBE 窗口，返回捕获到的错误框原文（没有则空串）
    pub fn finish(mut self) -> String {
        self.stop.store(true, Ordering::SeqCst);
        let keep = self.handle.take().and_then(|handle| handle.join().ok()).unwrap_or_default();
        for (hwnd, class, _) in windows_now() {
            if class == VBE_CLASS && !keep.contains(&hwnd) && visible(hwnd) {
                hide(hwnd);
            }
        }
        for hwnd in keep {
            show(hwnd);
        }
        self.caught.lock().map(|text| text.clone()).unwrap_or_default()
    }
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

/// 把 VBE 从「中断态」拉回停止态（等价于人手点 VBE 工具栏的「重置」方块按钮）。
///
/// 中断态下一切 Application 级 COM 调用都被拒，但 VBE 仍可交互（人能在 VBE 里点恢复），
/// 所以走 VBE 对象模型：CommandBars 菜单里按标题找「重置」（各语言标题不同 ——
/// Reset / 重置，模糊匹配；FindControl 的控件 ID 各引擎版本对不上，不赌它）并执行。
/// 工程本就停止时执行重置是无害的空操作。**谁调它**：见 [spawn_break_rescuer] ——
/// 主线程堵在 Run 里时只有旁路线程能走到这里
fn exit_break_mode(raw: &IDispatchWrapper) {
    let Ok(vbe) = raw.get("VBE").and_then(|v| v.idispatch()) else { return };
    let Ok(bars) = vbe.get_obj("CommandBars", "CommandBars") else { return };
    // **不能按名称取**（Item("Menu Bar") 在 WPS 的 VBE 上必抛 0x80020009，实测）：
    // 按序号枚举全部命令栏，用 Name 属性认「Menu Bar」
    let Ok(count) = bars.get("Count").and_then(|v| v.int()) else { return };
    for at in 1..=count {
        let Ok(bar) = bars.call("Item", vec![(at as i32).into()]).and_then(|v| v.idispatch()) else { continue };
        let name = bar.get("Name").ok().and_then(|v| v.string().ok()).unwrap_or_default();
        // 菜单栏的名字是**本地化**的（WPS 中文版叫「菜单条」，英文版 Menu Bar —— 实测）
        if !(name.eq_ignore_ascii_case("menu bar") || name.contains("菜单条") || name.contains("菜单栏")) {
            continue;
        }
        let Ok(tops) = bar.get_obj("Controls", "Controls") else { continue };
        let Ok(count) = tops.get("Count").and_then(|v| v.int()) else { continue };
        for at in 1..=count {
            let Ok(top) = tops.call("Item", vec![(at as i32).into()]).and_then(|v| v.idispatch()) else { continue };
            // 顶层是弹出菜单（msoControlPopup）；普通按钮没有 Controls 子集合，读失败正好跳过
            let Ok(items) = top.get_obj("Controls", "Controls") else { continue };
            let Ok(n) = items.get("Count").and_then(|v| v.int()) else { continue };
            for index in 1..=n {
                let Ok(item) = items.call("Item", vec![(index as i32).into()]).and_then(|v| v.idispatch()) else { continue };
                let caption = item.get("Caption").ok().and_then(|v| v.string().ok()).unwrap_or_default();
                // VBE 的「重置」命令 ID = 228（跨语言稳定，实测 WPS 中文版菜单项叫「重新设置」——
                // 不含「重置」子串，按标题匹配会漏）；标题匹配保留兜底
                let id = item.get("Id").ok().and_then(|v| v.int().ok());
                if id == Some(228)
                    || caption.to_ascii_lowercase().contains("reset")
                    || caption.contains("重置")
                {
                    let _ = item.call("Execute", vec![]);
                    return;
                }
            }
        }
    }
}

/// 断点救援线程：Run 期间旁路轮询 VBE 主窗口标题，发现「[中断]/[break]」就从
/// **第二条 COM 连接**执行 VBE「运行 > 重置」。
///
/// 为什么要它：编译错误的确定框被点掉后 WPS 并不返回 Run，而是直接进中断态
/// （VBE 标题「…[中断]…」，2026-10-15 实测）—— 主线程从此堵死在 Run 里，工程停在断点上，
/// 之后一切 Application 级 COM 调用都被拒（用户实报：出错一次整个应用点不动）。
/// 中断态下 VBE 仍可交互，所以旁路连接（COM 服务器在断点态照常处理新进调用）能替人执行重置；
/// 工程回到停止态后 Run 自然返回。主线程 Run 返回时置 `stop` 让线程退出
fn spawn_break_rescuer(engine: Engine, stop: Arc<AtomicBool>) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        // 兜底上限：60s 内救不回来（重试也无效）就放弃 —— 别陪着一条死连接转圈
        for _ in 0..120 {
            if stop.load(Ordering::SeqCst) {
                return;
            }
            sleep(Duration::from_millis(500));
            // 只认 VBE 主窗口标题里的中断标记：正常编译过/没在跑的工程没有这个词
            let broken = windows_now().iter().any(|(_, class, title)| {
                class == VBE_CLASS
                    && (title.contains("[中断]") || title.to_ascii_lowercase().contains("[break]"))
            });
            if !broken {
                continue;
            }
            // 主连接的线程堵在 Run 里：重置必须走第二条 COM 连接（App::connect 自管套间）
            if let Ok(app) = App::connect(engine, ConnectMode::Attach) {
                exit_break_mode(app.raw());
            }
        }
    })
}

/// VBA 操作入口。
pub struct Vba {
    app: IDispatchWrapper,
    /// 所属引擎（断点救援的旁路连接要按它 Attach 同一个实例）
    engine: Engine,
}

impl Vba {
    pub fn new(app: &App) -> Self {
        Self {
            app: app.raw().clone(),
            engine: app.engine(),
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
    /// 用户的工程本身有编译错误时同样会弹 VBE / 卡模态框，一并交给 [VbaGuard]
    pub fn run_macro(&self, name: &str, args: &[ComValue]) -> Result<()> {
        let mut a = vec![name.into()];
        a.extend(args.iter().cloned());
        let guard = VbaGuard::start();
        let rescue_stop = Arc::new(AtomicBool::new(false));
        let rescuer = spawn_break_rescuer(self.engine, Arc::clone(&rescue_stop));
        let outcome = self.app.call("Run", a);
        let dialog = guard.finish();
        rescue_stop.store(true, Ordering::SeqCst);
        let _ = rescuer.join();
        // Run 期间弹过 VBA 错误框（看门狗已点掉）= 宏没跑起来：工程编译不过，或宏自己
        // 抛了未处理的错误。Sub 正常结束也回 VT_EMPTY，失败只能靠这个信号识破
        if !dialog.is_empty() && outcome.is_ok() {
            return Err(ExcelError::new(
                "vba.run_macro",
                "VBA",
                format!("宏 {name} 没有执行：{dialog}"),
            )
            .with_hint("工程存在编译错误，或宏抛出了未处理的错误：在 VBE 里打开该工程修正后重试")
            .into());
        }
        let outcome = outcome
            .map(|_| ())
            .map_err(|e| {
                ExcelError::hint(
                    e,
                    "宏必须已存在于打开的工作簿中；带模块时用 '模块名.过程名'",
                )
            })
            .map_err(Into::into);
        outcome
    }

    /// 注入临时标准模块执行 VBA 代码，执行后自动清理。
    ///
    /// - `code`：完整 VBA 代码（须含公共 Sub）
    /// - `entry`：入口过程名（如 "Main"，代码里须有 `Public Sub Main()`）
    ///
    /// **运行时错误不弹窗**：入口外面包一层 `On Error GoTo` 的包装 Function，
    /// 错误被抓住后写成返回值带出来（`Application.Run` 调 Function 能拿到返回值），
    /// 由调用方转成文字错误交给上层（AI 据此改码重发）。用户代码内部的错误处理不受影响。
    /// **编译错误不弹窗不挂死不假成功**：见模块文档第 ②③ 层。
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
        // 包装 Function：抓住入口过程未处理的运行时错误，写成返回值带出来。
        // **用户代码原样注入，不加行号前缀** —— 曾为了 Erl 回报出错行给每行加数字行标签，
        // 实测（2026-10-15）两头都不成立：① 行标签加不得（过程声明行 / End 行 / 注释行 /
        // 空行 / Case 子句上都非法），整个模块编译不过，WPS 只是让 Run 抛 0x80020009 或干脆
        // 静默不跑，注入器却把它报成「入口过程名不一致」，模型于是反复改 Sub 名、整段重写
        //（用户实报：代码看着没问题却一直失败）；② 报错出在被调用的入口过程里，包装函数的
        // Erl 永远是 0，行号本来也带不出来（探针实测）
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
        // Run 期间开着外观守卫：编译错误也不弹 VBE、不卡模态框（模块文档第 ② 层）；
        // 同时开着断点救援：编译错误点掉确定后 WPS 直接进中断态、Run 永不返回（模块文档第 ④ 层）
        let guard = VbaGuard::start();
        let rescue_stop = Arc::new(AtomicBool::new(false));
        let rescuer = spawn_break_rescuer(self.engine, Arc::clone(&rescue_stop));
        let run_res = self.app.call("Run", vec![wrapper.into()]);
        let dialog = guard.finish();
        rescue_stop.store(true, Ordering::SeqCst);
        let _ = rescuer.join();
        let outcome: Result<()> = match run_res {
            // 编译错误：Run 期间弹过 VBA 错误框（看门狗已点掉，原文带回）。此刻包装
            // Function 根本没跑起来，而 Run 自己不当回事（不报错、只回个空值）——
            // 不认这个「有框」信号就会回一句「执行完成」，模型以为改好了，
            // 其实一个格子都没动（用户实报的假成功）
            Ok(_) if !dialog.is_empty() => Err(ExcelError::new(
                "vba.execute",
                "VBA",
                format!("VBA 编译错误，未执行任何操作（{dialog}）"),
            )
            .with_hint(
                "常见原因：调用了不存在的成员或方法（透视表缓存是 ActiveWorkbook.PivotCaches.Create，\
                 不是 ws.PivotCache）、常量未定义、If/For/With 块没闭合、跨行没加续行符 ` _`。逐行检查后修正再发",
            )
            .into()),
            // 包装 Function 的返回值：空串 = 正常跑完；非空 = 抓到的运行时错误原文。
            // **拿不到字符串就不是跑完了**：包装声明的是 As String，真跑起来必然是文本
            //（空串也算）—— 回空值说明模块没编译过、函数压根没执行，而 WPS 这时可能既不弹框
            // 也不报错（2026-10-15 探针实测：Run 静默回 VT_EMPTY、目标格一个没动），
            // 当成成功就是「假执行成功」
            Ok(result) => {
                let CellValue::Text(message) = result.value() else {
                    return Err(ExcelError::new(
                        "vba.execute",
                        "VBA",
                        "VBA 编译错误，未执行任何操作（没拿到包装函数的返回值）",
                    )
                    .with_hint(
                        "常见原因：调用了不存在的成员或方法、常量未定义、If/For/With 块没闭合、\
                         跨行没加续行符 ` _`。逐行检查后修正再发",
                    )
                    .into());
                };
                if message.trim().is_empty() {
                    Ok(())
                } else {
                    Err(ExcelError::new("vba.execute", "VBA", message)
                        // 「别整段重写」由 ops.rs 的就地修正指引统一给，这别再重复一句
                        .with_hint("已成功的操作不会重复执行，但它们留下的状态还在（如上轮已建好的表会让重名再失败）")
                        .into())
                }
            }
            Err(e) => {
                // 别把入口名当唯一嫌疑：模块没编译过同样走这条（WPS 不弹编译框，直接抛这个错）
                let mut hint = "宏运行失败：注入的代码没通过 VBA 编译。常见原因：代码里有 VBA 不接受的写法\
                                （跨行没加续行符 ` _`、块/括号没闭合、调用了不存在的成员），\
                                或入口过程名与代码里的公共 Sub 名一致"
                    .to_string();
                // 有对话框原文就一并带上（个别引擎下编译错误走的是抛错而非静默）
                if !dialog.is_empty() {
                    hint.push_str(&format!("；{dialog}"));
                }
                Err(ExcelError::hint(e, hint).into())
            }
        };
        outcome
    }

    /// 把 VBE 从「中断态」拉回停止态（公开入口，机制见 [exit_break_mode]）。
    pub fn reset_break_mode(&self) {
        exit_break_mode(&self.app);
    }

    /// 直接调用 `Application.Run`（返回 VARIANT，供需要返回值的高级用法）。
    pub fn run_raw(&self, name: &str, args: &[ComValue]) -> Result<VARIANT> {
        let mut a = vec![name.into()];
        a.extend(args.iter().cloned());
        Ok(self.app.call("Run", a)?.v)
    }
}
