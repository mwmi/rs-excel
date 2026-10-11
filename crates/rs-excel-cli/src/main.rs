//! rs-excel-cli —— WPS/Excel COM 交互式 REPL 测试工具。

mod commands;
mod completion;

use std::cell::RefCell;
use std::rc::Rc;

use anyhow::Result;
use reedline::Signal;

use commands::State;
use completion::{DynPrompt, editor};

fn main() -> Result<()> {
    println!("rs-excel-cli —— WPS/Excel COM 交互测试工具");
    println!("直接 list 可自动接管已打开的 WPS/Excel 并列出工作簿；connect attach [窗口序号] 可绑定 wins 列出的窗口；connect new 新建实例。");
    println!("输入 help 查看全部命令。Tab 在下方打开候选菜单（纵向列表、反色高亮），方向键导航，回车确认（自动补空格）。");
    let state = Rc::new(RefCell::new(State::new()));
    let mut line_editor = editor(state.clone());
    loop {
        let prompt = DynPrompt {
            text: state.borrow().prompt(),
        };
        match line_editor.read_line(&prompt)? {
            Signal::Success(line) => {
                let line = line.trim();
                if line.is_empty() {
                    continue;
                }
                if !commands::execute(&state, line, &mut line_editor) {
                    break;
                }
            }
            Signal::CtrlC | Signal::CtrlD => break,
            _ => {}
        }
    }
    println!("已退出（连接的应用保持运行；如需关闭应用进程，重新进入后执行 appquit，或直接关闭 WPS/Excel）。");
    Ok(())
}
