//! 应用层（Application）：连接、可见性、窗口管理、工作簿列表/打开/新建/退出。
//!
//! `App` 内部持有 COM 套间守卫，调用方无需理解 COM 初始化。
//! 注意：COM 单线程（STA）使用；`New` 模式拉起的应用进程在调用 `quit()` 前保持运行。

use anyhow::Result;
use windows::core::{GUID, HSTRING, IUnknown, Interface, PCWSTR};
use windows::Win32::Foundation::S_OK;
use windows::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, CLSIDFromProgID, CoCreateInstance, GetRunningObjectTable, IDispatch,
    IMoniker,
};
use windows::Win32::System::Ole::GetActiveObject;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::sel_index;
use crate::com::variant::{CellValue, ComValue};
use crate::engine::{ComApartment, ConnectMode, Engine};
use crate::error::ExcelError;
use crate::process::{exe_name_of, pid_of_hwnd, show_window};
use crate::range::Range;
use crate::worksheet::Worksheet;
use crate::workbook::{Workbook, WorkbookInfo};

/// 应用窗口状态。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WindowState {
    Maximized,
    Minimized,
    Normal,
}

/// 计算模式（xlCalculation 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CalcMode {
    Automatic = -4105,
    /// Excel 的「自动（除数据表外）」；WPS 可能不认，恢复时原样写回。
    SemiAutomatic = 2,
    Manual = -4135,
}

impl WindowState {
    fn code(self) -> i32 {
        // xlMaximized / xlMinimized / xlNormal
        match self {
            Self::Maximized => -4137,
            Self::Minimized => -4140,
            Self::Normal => -4139,
        }
    }
}

impl std::fmt::Display for WindowState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::Maximized => "最大化",
            Self::Minimized => "最小化",
            Self::Normal => "普通",
        })
    }
}

/// 窗口快照信息。
#[derive(Debug, Clone)]
pub struct WindowInfo {
    pub index: i32,
    pub caption: String,
    pub is_active: bool,
}

/// WPS/Excel 应用（Application COM 对象）。
pub struct App {
    raw: IDispatchWrapper,
    engine: Engine,
    _apartment: ComApartment,
}

impl App {
    /// 统一连接入口。
    ///
    /// - `Engine::Wps`（默认）/ `Engine::Excel`
    /// - `ConnectMode::Attach` 接管已打开实例；`ConnectMode::New` 新建实例
    pub fn connect(engine: Engine, mode: ConnectMode) -> Result<Self> {
        let apartment = ComApartment::init()?;
        let prog_id = engine.prog_id();
        let clsid = unsafe {
            CLSIDFromProgID(PCWSTR::from_raw(HSTRING::from(prog_id).as_ptr()))
        }
        .map_err(|e| {
            ExcelError::new(
                "connect",
                engine.label(),
                format!("0x{:08X} {e}", e.code().0 as u32),
            )
            .with_hint(format!("ProgID '{prog_id}' 未注册：确认 WPS/Excel 已安装"))
        })?;
        let disp: IDispatch = match mode {
            ConnectMode::New => {
                let d: IDispatch = unsafe { CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)? };
                verify_engine(&d, engine)?;
                d
            }
            ConnectMode::Attach => attach_engine(&clsid, engine)?,
        };
        Ok(Self {
            raw: IDispatchWrapper::root(disp, engine.label()),
            engine,
            _apartment: apartment,
        })
    }

    pub fn engine(&self) -> Engine {
        self.engine
    }

    /// 底层 COM 包装（扩展性核心：任意未封装成员可直调）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    pub fn caption(&self) -> Result<String> {
        self.raw.get("Caption")?.string()
    }

    pub fn set_visible(&self, on: bool) -> Result<()> {
        self.raw.put("Visible", vec![on.into()])
    }

    pub fn visible(&self) -> Result<bool> {
        self.raw.get("Visible")?.bool()
    }

    /// 关闭弹窗告警（SaveAs 遇同名文件等不再卡模态框）。
    pub fn set_display_alerts(&self, on: bool) -> Result<()> {
        self.raw.put("DisplayAlerts", vec![on.into()])
    }

    /// 是否弹窗告警（WPS 可能以布尔或整型返回，`bool()` 均可解析）。
    pub fn display_alerts(&self) -> Result<bool> {
        self.raw.get("DisplayAlerts")?.bool()
    }

    pub fn set_screen_updating(&self, on: bool) -> Result<()> {
        self.raw.put("ScreenUpdating", vec![on.into()])
    }

    pub fn screen_updating(&self) -> Result<bool> {
        self.raw.get("ScreenUpdating")?.bool()
    }

    /// 开关事件响应（批量写入前关闭可提速，并防止事件递归）。
    pub fn enable_events(&self, on: bool) -> Result<()> {
        self.raw.put("EnableEvents", vec![on.into()])
    }

    /// 事件响应当前是否开启（批量守卫恢复原值用）。
    pub fn events_enabled(&self) -> Result<bool> {
        self.raw.get("EnableEvents")?.bool()
    }

    /// 当前计算模式。
    pub fn calculation(&self) -> Result<CalcMode> {
        Ok(match self.raw.get("Calculation")?.int()? {
            -4135 => CalcMode::Manual,
            2 => CalcMode::SemiAutomatic,
            _ => CalcMode::Automatic,
        })
    }

    /// 设置计算模式。批量写公式前可切 Manual，结束切回 Automatic。
    pub fn set_calculation(&self, mode: CalcMode) -> Result<()> {
        self.raw.put("Calculation", vec![(mode as i32).into()])
    }

    /// 状态栏提示文字。
    pub fn set_status_bar(&self, text: &str) -> Result<()> {
        self.raw.put("StatusBar", vec![text.into()])
    }

    /// 恢复状态栏默认显示。
    pub fn clear_status_bar(&self) -> Result<()> {
        self.raw.put("StatusBar", vec![false.into()])
    }

    /// ★ 四合一批量模式开关。
    ///
    /// - `on=true`：关屏幕刷新 + 关事件 + 关弹窗 + 手动计算（批量操作前提速）；
    /// - `on=false`：全部恢复默认（刷新/事件/弹窗开启，自动计算）。
    ///
    /// **注意**：`on=false` 恢复的是默认值（刷新/事件/弹窗开、自动计算），不是调用前的状态。
    /// 需要「用完恢复原样」时用 [`App::fast_guard`]（RAII，嵌套安全）。
    pub fn set_fast(&self, on: bool) -> Result<()> {
        self.raw.put("ScreenUpdating", vec![(!on).into()])?;
        self.raw.put("EnableEvents", vec![(!on).into()])?;
        self.raw.put("DisplayAlerts", vec![(!on).into()])?;
        self.raw.put(
            "Calculation",
            vec![(if on { CalcMode::Manual } else { CalcMode::Automatic } as i32).into()],
        )
    }

    /// 批量提速守卫：快照当前刷新/事件/弹窗/计算设置，关闭后返回 RAII 守卫，
    /// 退出（含提前 return / 报错）时恢复**原值**。嵌套调用只有最外层真正改设置。
    pub fn fast_guard(&self) -> Result<crate::ai::FastGuard<'_>> {
        Ok(crate::ai::FastGuard::new(self)?)
    }

    /// 清除剪贴板复制状态（复制/粘贴后调用）。
    pub fn cut_copy_mode_off(&self) -> Result<()> {
        self.raw.put("CutCopyMode", vec![false.into()])
    }

    /// 重算所有打开的工作簿。
    pub fn calculate(&self) -> Result<()> {
        self.raw.call("Calculate", vec![]).map(|_| ())
    }

    /// ★ 计算工作表表达式，一个入口覆盖全部工作表函数：
    /// `evaluate("SUM(A1:A10)")` / `evaluate("VLOOKUP(...)")` / `evaluate("MATCH(...)")`。
    pub fn evaluate(&self, expr: &str) -> Result<CellValue> {
        Ok(self.raw.call("Evaluate", vec![expr.into()])?.value())
    }

    /// 合并两个区域为多区域（联合）。
    pub fn union(&self, a: &Range, b: &Range) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw
                .call(
                    "Union",
                    vec![
                        ComValue::Obj(a.raw().disp.clone()),
                        ComValue::Obj(b.raw().disp.clone()),
                    ],
                )?
                .idispatch()?,
        ))
    }

    /// 两个区域的交集。无交集返回 `Ok(None)`，不报错。
    pub fn intersect(&self, a: &Range, b: &Range) -> Result<Option<Range>> {
        let v = self.raw.call(
            "Intersect",
            vec![
                ComValue::Obj(a.raw().disp.clone()),
                ComValue::Obj(b.raw().disp.clone()),
            ],
        )?;
        if v.is_object() {
            Ok(Some(Range::from_raw(v.idispatch()?)))
        } else {
            Ok(None)
        }
    }

    /// 阻塞等待 `seconds` 秒（期间应用无响应，慎用长等待）。
    pub fn wait(&self, seconds: f64) -> Result<()> {
        let now = self.raw.get("Now")?.float()?;
        // Wait 接收时间序列值，1 天 = 1.0
        self.raw.call("Wait", vec![(now + seconds / 86400.0).into()])
            .map(|_| ())
    }

    /// 应用版本号（如 "16.0"）。
    pub fn version(&self) -> Result<String> {
        self.raw.get("Version")?.string()
    }

    /// 当前用户名。
    pub fn user_name(&self) -> Result<String> {
        self.raw.get("UserName")?.string()
    }

    pub fn set_window_state(&self, st: WindowState) -> Result<()> {
        // WPS 拒绝 COM 写 xlNormal/xlMinimized（0x80020009，xlMaximized 可用），失败时回退 Win32
        match self.raw.put("WindowState", vec![st.code().into()]) {
            Ok(()) => Ok(()),
            Err(e) => {
                let hwnd = self.raw.get("Hwnd").and_then(|v| v.int())?;
                if show_window(hwnd, st) {
                    Ok(())
                } else {
                    Err(e)
                }
            }
        }
    }

    /// 当前窗口状态（xlMaximized / xlMinimized / xlNormal）。
    pub fn window_state(&self) -> Result<WindowState> {
        let v = self.raw.get("WindowState")?.int()?;
        Ok(match v {
            -4137 => WindowState::Maximized,
            -4140 => WindowState::Minimized,
            _ => WindowState::Normal,
        })
    }

    /// 列出应用窗口。
    pub fn windows(&self) -> Result<Vec<WindowInfo>> {
        let wins = self.raw.get_obj("Windows", "Windows")?;
        let count = wins.get("Count")?.int()?;
        let active = self
            .raw
            .get("ActiveWindow")
            .ok()
            .and_then(|w| w.idispatch().ok())
            .and_then(|w| w.get("Caption").ok())
            .and_then(|v| v.string().ok());
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let w = wins.call("Item", vec![i.into()])?.idispatch()?;
            let caption = w.get("Caption")?.string()?;
            let is_active = active.as_deref() == Some(caption.as_str());
            out.push(WindowInfo {
                index: i,
                caption,
                is_active,
            });
        }
        Ok(out)
    }

    /// 激活窗口（1 基序号须在范围内，否则按标题匹配；标题来自 `windows()`）。
    pub fn activate_window(&self, idx_or_caption: &str) -> Result<()> {
        let wins = self.raw.get_obj("Windows", "Windows")?;
        let count = wins.get("Count")?.int()?;
        if let Some(n) = sel_index(idx_or_caption, count) {
            return wins
                .call("Item", vec![n.into()])
                .and_then(|v| v.idispatch())
                .and_then(|v| v.call("Activate", vec![]).map(|_| ()))
                .map_err(|e| ExcelError::hint(e, "窗口序号不存在：先用 windows() 查看"));
        }
        let list = self.windows()?;
        let Some(w) = list.iter().find(|w| w.caption == idx_or_caption) else {
            return Err(ExcelError::new(
                "activate_window",
                &self.raw.ctx,
                format!("窗口 '{idx_or_caption}' 不存在"),
            )
            .with_hint("先用 windows() 查看窗口列表")
            .into());
        };
        self.activate_window(&w.index.to_string())
    }

    /// 关闭第 `index` 个窗口（未保存会弹确认框，可先 `set_display_alerts(false)`）。
    pub fn close_window(&self, index: i32) -> Result<()> {
        self.raw
            .get_obj("Windows", "Windows")?
            .call("Item", vec![index.into()])?
            .idispatch()?
            .call("Close", vec![])
            .map(|_| ())
    }

    /// 列出打开的工作簿（名称/路径/保存状态/工作表数/是否活动）。
    pub fn workbooks(&self) -> Result<Vec<WorkbookInfo>> {
        let wbs = self.raw.get_obj("Workbooks", "Workbooks")?;
        let count = wbs.get("Count")?.int()?;
        let active_full = self
            .active_workbook()
            .ok()
            .and_then(|wb| wb.full_name().ok());
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let wb = Workbook::from_raw(
                wbs.call("Item", vec![i.into()])?
                    .idispatch_child(&format!("工作簿[{i}]"))?,
            );
            let name = wb.name()?;
            let full_name = wb.full_name()?;
            out.push(WorkbookInfo {
                index: i,
                name: name.clone(),
                full_name: full_name.clone(),
                path: wb.path()?,
                saved: wb.is_saved()?,
                sheet_count: wb.sheet_count()?,
                is_active: active_full.as_deref() == Some(full_name.as_str()),
            });
        }
        Ok(out)
    }

    /// 按**名称**绑定已打开的工作簿（数字名称也按名称解释）。
    /// 序号请用 [`App::workbook_at`] —— 与工作表一样，不再靠「字符串像不像数字」猜。
    pub fn workbook(&self, name: &str) -> Result<Workbook> {
        let wbs = self.raw.get_obj("Workbooks", "Workbooks")?;
        let raw = wbs
            .call("Item", vec![name.into()])
            .and_then(|v| v.idispatch_child(name))
            .map_err(|e| ExcelError::hint(e, "工作簿未打开或名称错误：先用 workbooks() 列出"))?;
        Ok(Workbook::from_raw(raw))
    }

    /// 按序号（1 基）绑定已打开的工作簿。
    pub fn workbook_at(&self, index: i32) -> Result<Workbook> {
        if index < 1 {
            return Err(ExcelError::new(
                "workbook_at",
                "Application",
                format!("工作簿序号从 1 开始（得到 {index}）"),
            )
            .into());
        }
        let wbs = self.raw.get_obj("Workbooks", "Workbooks")?;
        let count = wbs.get("Count")?.int()?;
        if index > count {
            return Err(ExcelError::new(
                "workbook_at",
                "Application",
                format!("工作簿序号 {index} 超出范围（共 {count} 本）"),
            )
            .with_hint("先用 workbooks() 看序号")
            .into());
        }
        let raw = wbs
            .call("Item", vec![index.into()])
            .and_then(|v| v.idispatch_child(&format!("工作簿[{index}]")))
            .map_err(|e| ExcelError::hint(e, "工作簿序号不存在：先用 workbooks() 列出"))?;
        Ok(Workbook::from_raw(raw))
    }

    /// 当前活动工作簿。
    pub fn active_workbook(&self) -> Result<Workbook> {
        Ok(Workbook::from_raw(
            self.raw.get_obj("ActiveWorkbook", "ActiveWorkbook")?,
        ))
    }

    /// 代码所在工作簿（COM 服务进程视角，等价 VBA ThisWorkbook）。
    pub fn this_workbook(&self) -> Result<Workbook> {
        Ok(Workbook::from_raw(
            self.raw.get_obj("ThisWorkbook", "ThisWorkbook")?,
        ))
    }

    /// 当前活动工作表。
    pub fn active_sheet(&self) -> Result<Worksheet> {
        Ok(Worksheet::from_raw(
            self.raw.get_obj("ActiveSheet", "ActiveSheet")?,
        ))
    }

    /// 当前活动单元格。
    pub fn active_cell(&self) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw.get_obj("ActiveCell", "ActiveCell")?,
        ))
    }

    /// 当前选中区域（可能是图表等非区域对象，此时后续操作会报错）。
    pub fn selection(&self) -> Result<Range> {
        Ok(Range::from_raw(
            self.raw.get_obj("Selection", "Selection")?,
        ))
    }

    /// 新建工作簿。
    pub fn new_workbook(&self) -> Result<Workbook> {
        let wbs = self.raw.get_obj("Workbooks", "Workbooks")?;
        Ok(Workbook::from_raw(
            wbs.call("Add", vec![])?
                .idispatch_child("新建工作簿")?,
        ))
    }

    /// 打开工作簿文件。相对路径按调用方当前目录自动转为绝对路径
    /// （COM 服务进程的工作目录与调用方不一致，直接传相对路径会"找不到文件"）。
    pub fn open_workbook(&self, path: &str) -> Result<Workbook> {
        let full = abs_path(path)?;
        let wbs = self.raw.get_obj("Workbooks", "Workbooks")?;
        Ok(Workbook::from_raw(
            wbs.call("Open", vec![full.clone().into()])?
                .idispatch_child(&full)
                .map_err(|e| {
                    ExcelError::hint(e, "确认文件存在且未被独占锁定：可用 ps 查看应用进程状态")
                })?,
        ))
    }

    /// 退出应用（进程结束；未保存工作簿是否提示取决于 DisplayAlerts）。
    pub fn quit(&self) -> Result<()> {
        self.raw.call("Quit", vec![]).map(|_| ())
    }
}

/// COM 对象归属引擎的探测结果。
#[derive(Debug)]
enum EngineProbe {
    /// 对象无 `Hwnd` 属性或进程不可查，无法判定
    Unknown,
    /// 进程可执行名与目标引擎一致
    Match,
    /// 归属其他进程（可执行名 + PID）
    Other { exe: String, pid: u32 },
}

/// 经 `Application.Hwnd` → PID → 可执行文件名探测 COM 对象归属的引擎。
fn engine_of(disp: &IDispatch, engine: Engine) -> EngineProbe {
    let Ok(hwnd) = IDispatchWrapper::root(disp.clone(), engine.label())
        .get("Hwnd")
        .and_then(|v| v.int())
    else {
        return EngineProbe::Unknown;
    };
    let Ok(pid) = pid_of_hwnd(hwnd) else {
        return EngineProbe::Unknown;
    };
    match exe_name_of(pid) {
        Ok(name) if name == engine.exe_name() => EngineProbe::Match,
        Ok(name) if !name.is_empty() => EngineProbe::Other { exe: name, pid },
        _ => EngineProbe::Unknown,
    }
}

/// New 模式校验：ProgID 可能被另一引擎接管（如 WPS 的 Office 兼容模式），
/// CoCreateInstance 拿到的对象必须真正属于目标引擎。
fn verify_engine(disp: &IDispatch, engine: Engine) -> Result<()> {
    if let EngineProbe::Other { exe, pid } = engine_of(disp, engine) {
        return Err(ExcelError::new(
            "connect",
            engine.label(),
            format!("ProgID '{}' 实际连接到 {exe}（PID {pid}）", engine.prog_id()),
        )
        .with_hint(
            "该 ProgID 已被另一引擎接管（如 WPS 的 Office 兼容模式）；确认目标应用已安装并在运行，或调整 COM 注册/兼容设置",
        )
        .into());
    }
    Ok(())
}

/// Attach：经 ROT 接管已运行的目标引擎实例。
///
/// WPS 兼容模式会把自身实例注册到 `Excel.Application` 等兼容 CLSID 的 ROT 项，
/// `GetActiveObject` 可能连到被接管的实例（而注册表未劫持时 `connect new` 却能
/// 创建正确的引擎）。因此快路径命中后按进程可执行名校验，不匹配则枚举 ROT
/// 逐项校验，找到真正属于目标引擎的实例。
pub(crate) fn attach_engine(clsid: &GUID, engine: Engine) -> Result<IDispatch> {
    let mut unk: Option<IUnknown> = None;
    unsafe { GetActiveObject(clsid, None, &mut unk).ok() };
    if let Some(disp) = unk.and_then(|u| u.cast::<IDispatch>().ok()) {
        if matches!(engine_of(&disp, engine), EngineProbe::Match) {
            return Ok(disp);
        }
    }
    // 慢路径：枚举 ROT，按进程可执行文件名过滤出目标引擎实例
    unsafe {
        let rot = GetRunningObjectTable(0)?;
        let enum_moniker = rot.EnumRunning()?;
        let mut mons = [None::<IMoniker>];
        let mut fetched = 0u32;
        while enum_moniker.Next(&mut mons, Some(&mut fetched)) == S_OK && fetched == 1 {
            let Some(m) = mons[0].take() else { break };
            let Ok(obj) = rot.GetObject(&m) else { continue };
            let Ok(disp) = obj.cast::<IDispatch>() else { continue };
            if matches!(engine_of(&disp, engine), EngineProbe::Match) {
                return Ok(disp);
            }
        }
    }
    Err(ExcelError::new(
        "connect.attach",
        engine.label(),
        format!("ROT 中未找到运行中的 {} 实例", engine.exe_name()),
    )
    .with_hint(format!(
        "ProgID '{}' 可能被另一引擎接管（如 WPS 的 Office 兼容模式），或目标引擎未运行；确认 {} 已打开后重试",
        engine.prog_id(),
        engine.exe_name()
    ))
    .into())
}

/// 相对路径 → 绝对路径。不要求文件已存在（存在时经 canonicalize 归一化，
/// 并去掉 Windows 扩展前缀 `\\?\`）。
fn abs_path(path: &str) -> Result<String> {
    let p = std::path::Path::new(path);
    let abs = if p.is_absolute() {
        p.to_path_buf()
    } else {
        std::env::current_dir()?.join(p)
    };
    let s = std::fs::canonicalize(&abs)
        .map(|c| c.to_string_lossy().into_owned())
        .unwrap_or_else(|_| abs.to_string_lossy().into_owned());
    Ok(s.strip_prefix(r"\\?\").map(str::to_string).unwrap_or(s))
}
