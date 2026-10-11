//! 进程控制：查看 / 强制结束目标引擎进程（WPS = et.exe，Excel = excel.exe）。
//!
//! COM 连接失败或应用卡死时，先 [`processes`] 查看状态，再 [`kill`] 结束指定
//! 进程或 [`kill_all`] 结束全部引擎进程。纯 Win32 实现，不依赖 COM 连接可用。

use anyhow::{anyhow, bail, Result};
use windows::core::PWSTR;
use windows::Win32::Foundation::{CloseHandle, HWND};
use windows::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W, TH32CS_SNAPPROCESS,
};
use windows::Win32::System::Threading::{
    OpenProcess, QueryFullProcessImageNameW, TerminateProcess, PROCESS_NAME_WIN32,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

use crate::engine::Engine;

/// 进程快照信息。
#[derive(Debug, Clone)]
pub struct ProcessInfo {
    pub pid: u32,
    /// 可执行文件名（小写），如 "et.exe"
    pub name: String,
}

impl std::fmt::Display for ProcessInfo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "PID {}  {}", self.pid, self.name)
    }
}

/// 列出引擎对应的进程（按可执行文件名精确匹配）。
pub fn processes(engine: Engine) -> Result<Vec<ProcessInfo>> {
    let target = engine.exe_name();
    let snap = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) }
        .map_err(|e| anyhow!("创建进程快照失败: {e}"))?;
    let mut entry = PROCESSENTRY32W {
        dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
        ..Default::default()
    };
    let mut out = Vec::new();
    if unsafe { Process32FirstW(snap, &mut entry) }.is_ok() {
        loop {
            let name = String::from_utf16_lossy(&entry.szExeFile);
            let name = name.trim_end_matches('\0').to_lowercase();
            if name == target {
                out.push(ProcessInfo {
                    pid: entry.th32ProcessID,
                    name,
                });
            }
            if unsafe { Process32NextW(snap, &mut entry) }.is_err() {
                break;
            }
        }
    }
    unsafe {
        let _ = CloseHandle(snap);
    }
    Ok(out)
}

/// 强制结束进程（TerminateProcess，退出码 1）。
pub fn kill(pid: u32) -> Result<()> {
    unsafe {
        let h = OpenProcess(PROCESS_TERMINATE, false, pid).map_err(|e| {
            anyhow!("打开进程 {pid} 失败（权限不足或进程已退出）: {e}")
        })?;
        let r = TerminateProcess(h, 1).map_err(|e| anyhow!("结束进程 {pid} 失败: {e}"));
        let _ = CloseHandle(h);
        r
    }
}

/// 窗口句柄 → 所属进程 PID（COM `Application.Hwnd` 归属校验用）。
pub fn pid_of_hwnd(hwnd: i32) -> Result<u32> {
    let mut pid = 0u32;
    unsafe {
        GetWindowThreadProcessId(HWND(hwnd as u32 as usize as *mut _), Some(&mut pid));
    }
    if pid == 0 {
        bail!("窗口句柄 {hwnd:#x} 无效或已失效");
    }
    Ok(pid)
}

/// 进程的可执行文件名（小写，如 "et.exe"）。
pub fn exe_name_of(pid: u32) -> Result<String> {
    unsafe {
        let h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid)
            .map_err(|e| anyhow!("打开进程 {pid} 失败（权限不足或已退出）: {e}"))?;
        let mut buf = [0u16; 1024];
        let mut len = buf.len() as u32;
        let r = QueryFullProcessImageNameW(
            h,
            PROCESS_NAME_WIN32,
            PWSTR(buf.as_mut_ptr()),
            &mut len,
        )
        .map_err(|e| anyhow!("查询进程 {pid} 路径失败: {e}"));
        let _ = CloseHandle(h);
        r?;
        let full = String::from_utf16_lossy(&buf[..len as usize]);
        Ok(full.rsplit(['\\', '/']).next().unwrap_or("").to_lowercase())
    }
}

/// 强制结束引擎的全部进程，返回成功结束的数量。
pub fn kill_all(engine: Engine) -> Result<usize> {
    let mut n = 0;
    for p in processes(engine)? {
        if kill(p.pid).is_ok() {
            n += 1;
        }
    }
    Ok(n)
}
