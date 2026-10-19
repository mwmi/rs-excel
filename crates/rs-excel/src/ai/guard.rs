//! 批量提速守卫：快照应用设置 → 关闭刷新/事件/弹窗 + 手动计算 → 退出时恢复**原值**。
//!
//! 与旧 `set_fast(false)` 的区别：旧实现恢复的是「默认值」（刷新开、事件开、弹窗开、自动计算），
//! 用户原本关掉刷新或设了手动计算时会被悄悄改掉；这里恢复的是调用前的真实值。
//!
//! 嵌套：只有最外层真正改设置，内层只加减深度，避免提前恢复外层。

use std::cell::Cell;

use crate::ai::error::{AiResult, ToolError};
use crate::app::{App, CalcMode};

thread_local! {
    /// 同一 STA 线程上的嵌套深度（每条 Tauri 命令一个线程，跨线程不共享）。
    static DEPTH: Cell<usize> = const { Cell::new(0) };
}

struct Previous {
    screen: Option<bool>,
    events: Option<bool>,
    alerts: Option<bool>,
    calculation: Option<CalcMode>,
}

/// RAII 守卫：`Drop` 兜底恢复；需要把恢复失败转告模型时调 [`FastGuard::finish`]。
pub struct FastGuard<'a> {
    app: &'a App,
    prev: Option<Previous>,
    done: bool,
    notes: Vec<String>,
}

impl<'a> FastGuard<'a> {
    pub fn new(app: &'a App) -> AiResult<Self> {
        let outer = DEPTH.with(|depth| {
            let now = depth.get();
            depth.set(now + 1);
            now == 0
        });
        if !outer {
            return Ok(Self {
                app,
                prev: None,
                done: false,
                notes: Vec::new(),
            });
        }
        // 先快照再改：读不到的那一项记为 None，退出时不碰它
        let prev = Previous {
            screen: app.screen_updating().ok(),
            events: app.events_enabled().ok(),
            alerts: app.display_alerts().ok(),
            calculation: app.calculation().ok(),
        };
        let mut notes = Vec::new();
        if let Err(e) = app.set_screen_updating(false) {
            notes.push(format!("关闭刷新失败：{e}"));
        }
        if let Err(e) = app.enable_events(false) {
            notes.push(format!("关闭事件失败：{e}"));
        }
        if let Err(e) = app.set_display_alerts(false) {
            notes.push(format!("关闭弹窗失败：{e}"));
        }
        if let Err(e) = app.set_calculation(CalcMode::Manual) {
            notes.push(format!("切手动计算失败：{e}"));
        }
        Ok(Self {
            app,
            prev: Some(prev),
            done: false,
            notes,
        })
    }

    /// 显式收尾：恢复并返回需要转告模型的问题（无问题返回 `None`）。
    pub fn finish(mut self) -> Option<String> {
        self.leave();
        (!self.notes.is_empty()).then(|| self.notes.join("；"))
    }

    fn leave(&mut self) {
        if self.done {
            return;
        }
        self.done = true;
        DEPTH.with(|depth| depth.set(depth.get().saturating_sub(1)));
        let Some(prev) = self.prev.take() else {
            return;
        };
        if let Some(on) = prev.screen {
            if let Err(e) = self.app.set_screen_updating(on) {
                self.notes.push(format!("恢复刷新设置失败：{e}"));
            }
        }
        if let Some(on) = prev.events {
            if let Err(e) = self.app.enable_events(on) {
                self.notes.push(format!("恢复事件设置失败：{e}"));
            }
        }
        if let Some(on) = prev.alerts {
            if let Err(e) = self.app.set_display_alerts(on) {
                self.notes.push(format!("恢复弹窗设置失败：{e}"));
            }
        }
        if let Some(mode) = prev.calculation {
            if let Err(e) = self.app.set_calculation(mode) {
                self.notes.push(format!("恢复计算模式失败：{e}"));
            }
        }
    }
}

impl Drop for FastGuard<'_> {
    fn drop(&mut self) {
        self.leave();
    }
}

/// 便捷包装：调用方只关心「有没有开启失败」。
pub fn fast(app: &App) -> Option<FastGuard<'_>> {
    FastGuard::new(app).ok()
}

/// 把守卫的问题转成结构化错误（恢复失败时用）。
pub fn recovery_error(note: &str) -> ToolError {
    ToolError::new("RECOVERY_FAILED", crate::ai::error::Category::Recovery, note)
        .with_hint("应用状态可能没完全恢复，先看当前设置再继续")
}
