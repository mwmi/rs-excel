//! 事件订阅（Application 级）：COM 连接点 + IDispatch 事件接收器，零轮询零延迟。
//!
//! 对应 VBA 的 `Workbook_*` / `Sheet*`（即 Application 层 `App_*`）事件族：
//! 工作簿开关/激活、工作表激活、选区变化、单元格修改等，由 WPS/Excel 原生回调推送。
//!
//! ```no_run
//! # let app = unreachable!();
//! use rs_excel::EventMonitor;
//!
//! let mut mon = EventMonitor::start(&app)?;      // 订阅全部支持的事件
//! // ... 用户在 WPS 里操作 ...
//! for ev in mon.poll() { println!("{ev}"); }     // 单行事件流，AI 直接读
//! mon.stop();
//! ```
//!
//! 注意：`app.enable_events(false)`（批量场景用 `app.fast_guard()`）期间应用暂停派发事件，
//! 属于 Excel 语义，订阅本身不受影响。

use std::fmt;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{self, Receiver, Sender};
use std::sync::Arc;
use std::time::Duration;

use anyhow::Result;
use windows::core::{GUID, HSTRING, Interface, PCWSTR};
use windows::Win32::Foundation::{S_OK, WAIT_FAILED};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS,
    DISPPARAMS, EXCEPINFO, IConnectionPoint, IConnectionPointContainer, IDispatch, IDispatch_Impl,
    ITypeInfo,
};
use windows::Win32::System::Variant::{VARIANT, VT_DISPATCH};
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, MSG, PM_REMOVE, PeekMessageW, QS_ALLINPUT, TranslateMessage, WM_QUIT,
    MsgWaitForMultipleObjects,
};

use crate::app::{attach_engine, App};
use crate::com::dispatch::IDispatchWrapper;
use crate::engine::Engine;
use crate::error::ExcelError;

/// Excel AppEvents 事件源 DIID（WPS 兼容 Excel typelib）。
const DIID_APP_EVENTS: GUID = GUID::from_u128(0x0002_4413_0000_0000_C000_0000_0000_0046);

// ---- DISPID 事件表（真机探针结论）----
//
// WPS 的 AppEvents 为自有定义（源自 et 类型库），与 Excel 官方 DISPID 不同；
// Excel 用官方 AppEvents 权威表。两引擎均按「参数对象分类」提取详情，
// 天然兼容 WPS（全命名参数、正序）与 Excel（位置参数、逆序）的传递差异。

/// 事件 DISPID → 事件名（按引擎查表，未注册返回 None）。
fn event_kind(dispid: i32, engine: Engine) -> Option<&'static str> {
    Some(match engine {
        Engine::Wps => match dispid {
            1554 => "窗口缩放",
            1556 => "窗口激活",
            1557 => "窗口失活",
            1558 => "选区变化",
            1559 => "双击",
            1560 => "右键",
            1561 => "工作表激活",
            1562 => "工作表失活",
            1563 => "重算",
            1564 => "单元格修改",
            1565 => "工作簿新建",
            1567 => "工作簿打开",
            1568 => "工作簿激活",
            1569 => "工作簿失活",
            1570 => "工作簿关闭前",
            1571 => "保存前",
            1572 => "打印前",
            1573 => "新增工作表",
            1854 => "超链接跳转",
            2612 => "计算完成",
            2911 => "保存后",
            2912 => "新增图表",
            _ => return None,
        },
        Engine::Excel => match dispid {
            1545 => "工作簿新建",
            1547 => "选区变化",
            1548 => "双击",
            1549 => "右键",
            1550 => "工作表激活",
            1551 => "工作表失活",
            1552 => "重算",
            1553 => "单元格修改",
            1554 => "工作簿打开",
            1555 => "工作簿激活",
            1556 => "工作簿失活",
            1557 => "工作簿关闭前",
            1558 => "保存前",
            1559 => "打印前",
            1560 => "新增工作表",
            1563 => "窗口缩放",
            1564 => "窗口激活",
            1565 => "窗口失活",
            2612 => "计算完成",
            2911 => "保存后",
            2912 => "新增图表",
            _ => return None,
        },
    })
}

/// 一条事件：`Display` 输出单行 `事件名: 详情`，AI 直接读。
#[derive(Debug, Clone)]
pub struct ExcelEvent {
    /// 事件名，如 `"工作表激活"` / `"选区变化"`
    pub kind: &'static str,
    /// 详情，如 `"Book1.xlsx > Sheet3!B2:C5"`
    pub detail: String,
}

impl fmt::Display for ExcelEvent {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.detail.is_empty() {
            f.write_str(self.kind)
        } else {
            write!(f, "{}: {}", self.kind, self.detail)
        }
    }
}

/// 应用状态快照（不订阅事件时按需一次读取全貌）。
///
/// `Display` 单行输出：
/// `Book1.xlsx | Sheet3 | 选区 B2:C5 | 活动格 B2 | 打开 2: Book1.xlsx, Book2.xlsx`
#[derive(Debug, Clone, Default)]
pub struct AppState {
    /// 活动工作簿 FullName（未保存即为名称；WPS 首页态为空）
    pub workbook: String,
    /// 活动工作表名
    pub sheet: String,
    /// 选区地址（已去 `$`，如 `B2:C5`；选中图表等非区域对象时为空）
    pub selection: String,
    /// 活动单元格（已去 `$`，如 `B2`）
    pub cell: String,
    /// 打开的工作簿名列表
    pub open: Vec<String>,
}

impl fmt::Display for AppState {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{} | {}",
            if self.workbook.is_empty() { "无活动工作簿" } else { &self.workbook },
            if self.sheet.is_empty() { "无活动工作表" } else { &self.sheet }
        )?;
        if !self.selection.is_empty() {
            write!(f, " | 选区 {}", self.selection)?;
        }
        if !self.cell.is_empty() {
            write!(f, " | 活动格 {}", self.cell)?;
        }
        if self.open.is_empty() {
            write!(f, " | 打开 0")
        } else {
            write!(f, " | 打开 {}: {}", self.open.len(), self.open.join(", "))
        }
    }
}

impl App {
    /// 当前状态快照（活动工作簿/工作表/选区/活动单元格/打开列表）。
    ///
    /// 任一项读取失败（如 WPS 首页无文档、选区为形状）时该项留空，不报错。
    pub fn state(&self) -> Result<AppState> {
        let raw = self.raw();
        let workbook = raw
            .get("ActiveWorkbook").ok().and_then(|v| v.idispatch().ok())
            .and_then(|w| w.get("FullName").ok().and_then(|v| v.string().ok()))
            .unwrap_or_default();
        let sheet = raw
            .get("ActiveSheet").ok().and_then(|v| v.idispatch().ok())
            .and_then(|s| s.get("Name").ok().and_then(|v| v.string().ok()))
            .unwrap_or_default();
        let selection = raw
            .get("Selection").ok().and_then(|v| v.idispatch().ok())
            .and_then(|r| addr_of_wrap(&r))
            .unwrap_or_default();
        let cell = raw
            .get("ActiveCell").ok().and_then(|v| v.idispatch().ok())
            .and_then(|r| addr_of_wrap(&r))
            .unwrap_or_default();
        let mut open = Vec::new();
        if let Ok(wbs) = raw.get_obj("Workbooks", "Workbooks") {
            if let Ok(count) = wbs.get("Count").and_then(|v| v.int()) {
                for i in 1..=count {
                    if let Ok(name) = wbs
                        .call("Item", vec![i.into()])
                        .and_then(|v| v.idispatch())
                        .and_then(|w| w.get("Name").and_then(|v| v.string()))
                    {
                        open.push(name);
                    }
                }
            }
        }
        Ok(AppState { workbook, sheet, selection, cell, open })
    }
}

/// 事件订阅器：经 COM 连接点订阅 Application 级事件，原生回调推送（无轮询）。
///
/// 内部一个专用 STA 监听线程持有接收器并泵消息（WPS/Excel 的事件回调在该线程
/// 投递），与 `App` 的 STA 互不干扰；事件在回调线程解析为纯文本后入队，
/// `poll()` 一次取走全部。
pub struct EventMonitor {
    rx: Receiver<ExcelEvent>,
    stop: Arc<AtomicBool>,
    handle: Option<std::thread::JoinHandle<()>>,
}

impl EventMonitor {
    /// 订阅 `app` 所属引擎的 Application 级事件（工作簿开关/激活、工作表激活、
    /// 选区变化、单元格修改、双击/右键、重算、保存/打印/关闭前、新增工作表）。
    ///
    /// 失败原因（引擎未运行 / 不支持 AppEvents 连接点）以单行结构化错误返回。
    pub fn start(app: &App) -> Result<Self> {
        start_monitor(app.engine())
    }

    /// 取走自上次调用以来的全部事件（不丢不重，取后清空）。
    pub fn poll(&self) -> Vec<ExcelEvent> {
        let mut out = Vec::new();
        while let Ok(ev) = self.rx.try_recv() {
            out.push(ev);
        }
        out
    }

    /// 停止订阅并回收监听线程（`Drop` 自动调用，通常无需显式调）。
    pub fn stop(&mut self) {
        self.stop.store(true, Ordering::Relaxed);
        if let Some(h) = self.handle.take() {
            let _ = h.join();
        }
    }
}

impl Drop for EventMonitor {
    fn drop(&mut self) {
        self.stop();
    }
}

fn start_monitor(engine: Engine) -> Result<EventMonitor> {
    let (ev_tx, ev_rx) = mpsc::channel::<ExcelEvent>();
    let (ready_tx, ready_rx) = mpsc::channel::<Result<()>>();
    let stop = Arc::new(AtomicBool::new(false));
    let stop2 = stop.clone();
    let handle = std::thread::Builder::new()
        .name("rs-excel-events".into())
        .spawn(move || {
            let init = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
            if let Err(e) = init.ok() {
                let _ = ready_tx.send(Err(e.into()));
                return;
            }
            match subscribe(engine, &ev_tx) {
                Ok((cp, cookie)) => {
                    let _ = ready_tx.send(Ok(()));
                    unsafe { pump_until_stop(&stop2) }; // 泵消息期间事件回调送达
                    unsafe { let _ = cp.Unadvise(cookie); }
                }
                Err(e) => {
                    let _ = ready_tx.send(Err(e));
                }
            }
            unsafe { CoUninitialize() };
        })?;
    match ready_rx.recv_timeout(Duration::from_secs(10)) {
        Ok(Ok(())) => Ok(EventMonitor {
            rx: ev_rx,
            stop,
            handle: Some(handle),
        }),
        Ok(Err(e)) => Err(e),
        Err(_) => {
            let _ = handle.join();
            Err(ExcelError::new(
                "events.start",
                engine.label(),
                "订阅超时（10s 内未完成 Advise）",
            )
            .with_hint("确认 WPS/Excel 可正常响应 COM 调用（未被模态弹窗阻塞）")
            .into())
        }
    }
}

/// STA 消息泵：期间 COM 事件回调经本线程投递；stop 置位或收到 WM_QUIT 时退出。
/// 100ms 一轮 `MsgWaitForMultipleObjects`（内核等待，零 CPU 占用）。
unsafe fn pump_until_stop(stop: &AtomicBool) {
    let mut msg = MSG::default();
    loop {
        // 无句柄：仅等待消息，兼作可中断的睡眠
        let r = MsgWaitForMultipleObjects(None, false, 100, QS_ALLINPUT);
        if stop.load(Ordering::Relaxed) || r == WAIT_FAILED {
            return;
        }
        while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
            if msg.message == WM_QUIT {
                return;
            }
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

/// 附加引擎 → 找 AppEvents 连接点 → Advise 事件接收器。
fn subscribe(engine: Engine, tx: &Sender<ExcelEvent>) -> Result<(IConnectionPoint, u32)> {
    unsafe {
        let clsid = CLSIDFromProgID(PCWSTR::from_raw(HSTRING::from(engine.prog_id()).as_ptr()))
            .map_err(|e| {
                ExcelError::new(
                    "events.start",
                    engine.label(),
                    format!("0x{:08X} {e}", e.code().0 as u32),
                )
                .with_hint(format!("ProgID '{}' 未注册：确认 WPS/Excel 已安装", engine.prog_id()))
            })?;
        let app = attach_engine(&clsid, engine)?;
        let container: IConnectionPointContainer = app.cast().map_err(|e| {
            ExcelError::new(
                "events.start",
                engine.label(),
                format!("0x{:08X} {e}", e.code().0 as u32),
            )
            .with_hint("Application 不支持连接点容器，无法订阅事件")
        })?;
        let cp = match container.FindConnectionPoint(&DIID_APP_EVENTS) {
            Ok(cp) => cp,
            Err(_) => return Err(no_app_events(&container, engine)),
        };
        let sink: IDispatch = EventSink { tx: tx.clone(), engine }.into();
        let cookie = cp.Advise(&sink).map_err(|e| {
            ExcelError::new(
                "events.start",
                engine.label(),
                format!("0x{:08X} {e}", e.code().0 as u32),
            )
            .with_hint("事件接收器 Advise 失败")
        })?;
        Ok((cp, cookie))
    }
}

/// FindConnectionPoint 失败时的诊断：枚举引擎实际支持的连接点 DIID 一并报告。
fn no_app_events(container: &IConnectionPointContainer, engine: Engine) -> anyhow::Error {
    let mut found: Vec<String> = Vec::new();
    unsafe {
        if let Ok(ecp) = container.EnumConnectionPoints() {
            let mut arr: [Option<IConnectionPoint>; 8] = std::array::from_fn(|_| None);
            let mut fetched = 0u32;
            while ecp.Next(&mut arr, &mut fetched) == S_OK && fetched > 0 {
                for cp in arr.iter_mut().take(fetched as usize).flatten() {
                    if let Ok(g) = cp.GetConnectionInterface() {
                        found.push(format!("{g:?}"));
                    }
                }
                if (fetched as usize) < arr.len() {
                    break;
                }
            }
        }
    }
    ExcelError::new(
        "events.start",
        engine.label(),
        format!("未找到 AppEvents 连接点 {DIID_APP_EVENTS:?}"),
    )
    .with_hint(if found.is_empty() {
        "引擎未暴露任何事件连接点".to_string()
    } else {
        format!("引擎实际支持的连接点 DIID: {}", found.join(", "))
    })
    .into()
}

/// IDispatch 事件接收器：`Invoke` 按引擎查 DISPID 表分发，解析为单行文本事件入队。
#[windows::core::implement(IDispatch)]
struct EventSink {
    tx: Sender<ExcelEvent>,
    engine: Engine,
}

impl IDispatch_Impl for EventSink_Impl {
    fn GetTypeInfoCount(&self) -> windows::core::Result<u32> {
        Ok(0)
    }

    fn GetTypeInfo(&self, _itinfo: u32, _lcid: u32) -> windows::core::Result<ITypeInfo> {
        Err(e_notimpl())
    }

    fn GetIDsOfNames(
        &self,
        _riid: *const GUID,
        _rgsznames: *const PCWSTR,
        _cnames: u32,
        _lcid: u32,
        _rgdispid: *mut i32,
    ) -> windows::core::Result<()> {
        Err(e_notimpl())
    }

    fn Invoke(
        &self,
        dispidmember: i32,
        _riid: *const GUID,
        _lcid: u32,
        _wflags: DISPATCH_FLAGS,
        pdispparams: *const DISPPARAMS,
        _pvarresult: *mut VARIANT,
        _pexcepinfo: *mut EXCEPINFO,
        _puargerr: *mut u32,
    ) -> windows::core::Result<()> {
        if !pdispparams.is_null() {
            // 任意解析失败仅丢弃该条事件，绝不影响订阅
            if let Some(ev) = self.on_event(dispidmember, unsafe { &*pdispparams }) {
                let _ = self.tx.send(ev);
            }
        }
        Ok(())
    }
}

impl EventSink_Impl {
    /// DISPID → 事件。详情按「参数对象分类」提取（选区→"工作簿 > 表!地址"、
    /// 工作表→"工作簿 > 表"、工作簿→FullName、窗口→标题），
    /// 不依赖参数顺序，同时兼容 WPS（命名参数正序）与 Excel（位置参数逆序）。
    fn on_event(&self, dispid: i32, p: &DISPPARAMS) -> Option<ExcelEvent> {
        let Some(kind) = event_kind(dispid, self.engine) else {
            // 未注册 DISPID 不静默丢弃：上抛 dispid + 参数分类，便于发现引擎差异
            let n = p.cArgs as usize;
            let mut parts = Vec::with_capacity(n);
            unsafe {
                for i in 1..=n {
                    match arg_disp(p, i) {
                        Some(d) => {
                            let (k, txt) = classify(&d);
                            parts.push(format!("#{i} {k}:{txt}"));
                        }
                        None => parts.push(format!("#{i} 非对象")),
                    }
                }
            }
            return Some(ExcelEvent {
                kind: "未知事件",
                detail: format!("dispid={dispid}，参数: {}", parts.join(" | ")),
            });
        };
        let (mut range, mut sheet, mut wb, mut win) = (None, None, None, None);
        unsafe {
            for i in 1..=p.cArgs as usize {
                if let Some(d) = arg_disp(p, i) {
                    match arg_class(&d) {
                        ArgClass::Range => range = range.or(Some(d)),
                        ArgClass::Worksheet => sheet = sheet.or(Some(d)),
                        ArgClass::Workbook => wb = wb.or(Some(d)),
                        ArgClass::Window => win = win.or(Some(d)),
                        ArgClass::Other => {}
                    }
                }
            }
        }
        let detail = if let Some(r) = range {
            target_path(&r)
        } else if let Some(s) = sheet {
            sheet_path(&s)
        } else if let Some(w) = wb {
            full_name(&w)
        } else if let Some(w) = win {
            prop_str(&w, "Caption")
        } else {
            None
        }
        .unwrap_or_default();
        Some(ExcelEvent { kind, detail })
    }
}

fn e_notimpl() -> windows::core::Error {
    windows::core::Error::from_hresult(windows::core::HRESULT(0x8000_4001u32 as i32))
}

// ---- 事件参数解析助手（全部容错，失败返回 None）----

/// 位置参数取 COM 对象：`i` 为 1 基从左数（COM `rgvarg` 逆序存放）。
unsafe fn arg_disp(p: &DISPPARAMS, i: usize) -> Option<IDispatch> {
    let n = p.cArgs as usize;
    if i == 0 || i > n {
        return None;
    }
    let v = &*p.rgvarg.add(n - i);
    if v.Anonymous.Anonymous.vt != VT_DISPATCH {
        return None;
    }
    (*v.Anonymous.Anonymous.Anonymous.pdispVal).clone()
}

fn wrap(d: &IDispatch) -> IDispatchWrapper {
    IDispatchWrapper::root(d.clone(), "event")
}

fn prop_str(d: &IDispatch, prop: &str) -> Option<String> {
    wrap(d).get(prop).ok()?.string().ok()
}

/// 工作簿标识：FullName 优先（未保存时即名称）。
fn full_name(wb: &IDispatch) -> Option<String> {
    prop_str(wb, "FullName").or_else(|| prop_str(wb, "Name"))
}

/// `"工作簿 > 工作表"`。
fn sheet_path(sh: &IDispatch) -> Option<String> {
    let wb = wrap(sh).get("Parent").ok()?.idispatch().ok()?.disp;
    Some(format!("{} > {}", full_name(&wb)?, prop_str(sh, "Name")?))
}

/// `"工作簿 > 工作表!地址"`（地址已去 `$`）。
fn target_path(target: &IDispatch) -> Option<String> {
    let sh = wrap(target).get("Parent").ok()?.idispatch().ok()?.disp;
    Some(format!("{}!{}", sheet_path(&sh)?, addr_of(target)?))
}

fn addr_of(r: &IDispatch) -> Option<String> {
    prop_str(r, "Address").map(|a| a.replace('$', ""))
}

/// 参数对象分类（判定依据互斥：Range 有 Address，Window 有 Hwnd，
/// Workbook 有 FullName，Worksheet 的 Parent 是 Workbook）。
enum ArgClass {
    Range,
    Worksheet,
    Workbook,
    Window,
    Other,
}

fn arg_class(d: &IDispatch) -> ArgClass {
    if addr_of(d).is_some() {
        return ArgClass::Range;
    }
    if wrap(d).get("Hwnd").is_ok() {
        return ArgClass::Window;
    }
    if prop_str(d, "FullName").is_some() {
        return ArgClass::Workbook;
    }
    if prop_str(d, "Name").is_some() {
        if let Ok(parent) = wrap(d).get("Parent").and_then(|p| p.idispatch()) {
            if prop_str(&parent.disp, "FullName").is_some() {
                return ArgClass::Worksheet;
            }
        }
    }
    ArgClass::Other
}

/// 未知事件的参数概要文本（诊断用）。
fn classify(d: &IDispatch) -> (&'static str, String) {
    match arg_class(d) {
        ArgClass::Range => ("Range", addr_of(d).unwrap_or_default()),
        ArgClass::Window => ("Window", prop_str(d, "Caption").unwrap_or_default()),
        ArgClass::Workbook => ("Workbook", full_name(d).unwrap_or_default()),
        ArgClass::Worksheet => ("Worksheet", prop_str(d, "Name").unwrap_or_default()),
        ArgClass::Other => ("Other", prop_str(d, "Name").unwrap_or_default()),
    }
}

fn addr_of_wrap(w: &IDispatchWrapper) -> Option<String> {
    w.get("Address").ok()?.string().ok().map(|a| a.replace('$', ""))
}
