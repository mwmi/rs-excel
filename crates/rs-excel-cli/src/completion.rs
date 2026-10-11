//! 表格驱动的补全 + 候选菜单 + help 生成：[COMMANDS] 表是唯一事实来源。
//!
//! **扩展方式**（新命令两步完成，帮助与补全自动生效）：
//! ① 在 [COMMANDS] 表加一条 `CommandSpec`；
//! ② 在 commands.rs 的 `run()` 加一个 match 分支。
//!
//! 动态候选：实现 `fn(&State) -> Vec<String>`，以 `ArgSpec::Dynamic(fn, 显示标签)` 挂到对应参数位。
//! 索引类候选（工作簿/工作表/窗口）用「序号. 名称」单行格式：菜单显示整行，
//! 回车只插入序号（避免同一项的序号与名称拆成多行显示）。
//! 子命令：用 `ArgSpec::Sub(&[(分支名, 说明, 后续参数规格)])`，分支参数从位置 1 开始；
//! 分支候选在菜单中带说明显示，回车只插入分支名。
//! 固定候选：`ArgSpec::Fixed/OptFixed(&[(值, 说明)])`，同样「菜单显示说明、插入只取值」。
//!
//! 补全交互（ohmyzsh / PSReadLine 风格，基于 reedline ColumnarMenu）：
//! 输入行**下方纵向列出候选**（单列，每行一个），**选中项反色高亮**；
//! `Tab` 打开菜单并循环下一候选，`Shift+Tab` 反向，**方向键导航**（菜单激活时），
//! 回车接受候选，并在候选文本后**自动补一个空格表示确认**（目录路径除外，便于续接子路径）。
//! 命令候选**上下文感知**：按会话级别（无工作簿 → 有工作簿 → 有工作表）
//! 隐藏无法执行的命令，并把当前上下文相关的命令组置顶。

use std::borrow::Cow;
use std::cell::RefCell;
use std::rc::Rc;

use reedline::{
    ColumnarMenu, CompletionResult, Completer, Emacs, KeyCode, KeyModifiers, MenuBuilder, Prompt,
    PromptEditMode, PromptHistorySearch, Reedline, ReedlineEvent, ReedlineMenu, Signal, Span,
    Suggestion, default_emacs_keybindings,
};

use rs_excel::processes;

use crate::commands::State;

// ── 候选常量 ──

/// 固定候选（值, 说明）：菜单显示「值 说明」，插入只取值。
const ENGINES: &[(&str, &str)] = &[("wps", "WPS 表格"), ("excel", "Microsoft Excel")];
const YES_NO: &[(&str, &str)] = &[("y", "是"), ("n", "否")];
const ON_OFF: &[(&str, &str)] = &[("on", "开"), ("off", "关")];
const CHART_TYPES: &[(&str, &str)] = &[
    ("column", "柱形图"),
    ("bar", "条形图"),
    ("line", "折线图"),
    ("pie", "饼图"),
    ("area", "面积图"),
    ("scatter", "散点图"),
    ("doughnut", "圆环图"),
];
const FONT_PROPS: &[(&str, &str)] = &[
    ("bold", "加粗"),
    ("italic", "斜体"),
    ("underline", "下划线"),
    ("size", "字号"),
    ("color", "颜色(#RRGGBB/r,g,b)"),
    ("name", "字体名"),
];
const RAW_KINDS: &[(&str, &str)] = &[("get", "读属性"), ("call", "调方法"), ("put", "写属性")];
const PAPERS: &[(&str, &str)] = &[
    ("A4", "纸张 A4"),
    ("A3", "纸张 A3"),
    ("A5", "纸张 A5"),
    ("B4", "纸张 B4"),
    ("B5", "纸张 B5"),
    ("Letter", "信纸"),
    ("Legal", "法律用纸"),
    ("Executive", "行政用纸"),
    ("A2", "纸张 A2"),
    ("A1", "纸张 A1"),
];
const ORIENTS: &[(&str, &str)] = &[("portrait", "纵向"), ("landscape", "横向")];
const HF_POS: &[(&str, &str)] = &[("left", "左"), ("center", "中"), ("right", "右")];
const TITLE_KINDS: &[(&str, &str)] = &[("rows", "标题行"), ("cols", "标题列")];
const WINDOW_STATES: &[(&str, &str)] = &[("max", "最大化"), ("min", "最小化"), ("normal", "普通")];

// ── 命令上下文分组（Tab 候选的可见性与置顶） ──

/// 未选定工作簿时置顶的工作簿管理命令（save/saveas/pdf/close 因需要工作簿而被隐藏）。
const WB_MGMT_CMDS: &[&str] = &["list", "use", "new", "open", "save", "saveas", "pdf", "close"];
/// 已选定工作簿时置顶的工作表命令。
const SHEET_CMDS: &[&str] = &["sheets", "sheet", "addsheet"];
/// 已选定工作表时置顶的单元格/页面/VBA 命令。
const SHEET_CTX_CMDS: &[&str] = &[
    "info", "get", "set", "tget", "tset", "formula", "fmt", "font", "merge", "unmerge", "chart",
    "page", "vba",
];

/// 命令的最低上下文要求：2 = 需选定工作表，1 = 需选定工作簿，0 = 无要求。
fn required_ctx(name: &str) -> u8 {
    if SHEET_CTX_CMDS.contains(&name) {
        2
    } else if SHEET_CMDS.contains(&name) {
        1
    } else {
        0
    }
}

/// 帮助分组（避免散落字面量）。
const G_HELP: &str = "帮助";
const G_CONN: &str = "连接与引擎";
const G_WB: &str = "工作簿";
const G_WS: &str = "工作表";
const G_CELL: &str = "单元格";
const G_VBA: &str = "VBA";
const G_WIN: &str = "应用与窗口";
const G_EXT: &str = "扩展";

// ── 命令表（单一事实来源） ──

/// 参数位补全规格（下标 = 参数位置）。
pub enum ArgSpec {
    /// 固定候选（值, 说明）
    Fixed(&'static [(&'static str, &'static str)]),
    /// 可选固定候选（帮助中显示为 [...]）
    OptFixed(&'static [(&'static str, &'static str)]),
    /// 动态候选（可变借用会话状态，候选函数可自动接管应用）+ 帮助显示标签
    Dynamic(fn(&mut State) -> Vec<String>, &'static str),
    /// 文件路径（扫描文件系统）
    Path,
    /// 自由文本：不补全，帮助中显示 <占位符>
    Any(&'static str),
    /// 可选自由文本：帮助中显示 [占位符]
    OptAny(&'static str),
    /// 子命令分支：分支名占位置 0，分支参数从位置 1 开始；(分支名, 说明, 分支参数)
    Sub(&'static [(&'static str, &'static str, &'static [ArgSpec])]),
}

/// 命令规格。
pub struct CommandSpec {
    /// 帮助分组（空串 = 别名，不进帮助）
    pub group: &'static str,
    pub name: &'static str,
    pub desc: &'static str,
    pub args: &'static [ArgSpec],
}

/// 全部命令。分组连续排列，帮助按分组输出。
pub static COMMANDS: &[CommandSpec] = &[
    CommandSpec { group: G_HELP, name: "help", desc: "本帮助", args: &[] },
    // 连接与引擎
    CommandSpec {
        group: G_CONN,
        name: "connect",
        desc: "接管已打开的应用（可带 wins 窗口序号|标题绑定）/ 新建应用",
        args: &[ArgSpec::Sub(&[
            ("attach", "接管已打开实例，可带窗口序号|标题绑定", &[ArgSpec::Dynamic(window_candidates, "<窗口序号|标题>")]),
            ("new", "新建应用实例", &[]),
        ])],
    },
    CommandSpec { group: G_CONN, name: "engine", desc: "查看/切换引擎（切换后需重新 connect）", args: &[ArgSpec::OptFixed(ENGINES)] },
    CommandSpec { group: G_CONN, name: "appquit", desc: "关闭应用进程并断开", args: &[] },
    CommandSpec { group: G_CONN, name: "ps", desc: "查看当前引擎进程（et.exe/excel.exe，不依赖连接）", args: &[] },
    CommandSpec { group: G_CONN, name: "kill", desc: "强制结束指定 PID 进程（Tab 列出引擎进程）", args: &[ArgSpec::Dynamic(process_candidates, "<PID>")] },
    CommandSpec { group: G_CONN, name: "killall", desc: "强制结束当前引擎全部进程", args: &[] },
    CommandSpec { group: G_CONN, name: "quit", desc: "退出本工具（不影响应用）", args: &[] },
    // 工作簿
    CommandSpec { group: G_WB, name: "list", desc: "列出打开的工作簿（未连接自动接管应用）", args: &[] },
    CommandSpec { group: G_WB, name: "use", desc: "选定当前工作簿", args: &[ArgSpec::Dynamic(workbook_candidates, "<序号|名称>")] },
    CommandSpec { group: G_WB, name: "new", desc: "新建工作簿并选定", args: &[] },
    CommandSpec { group: G_WB, name: "open", desc: "打开工作簿并选定", args: &[ArgSpec::Path] },
    CommandSpec { group: G_WB, name: "save", desc: "保存当前工作簿", args: &[] },
    CommandSpec { group: G_WB, name: "saveas", desc: "另存为（按扩展名自动选格式）", args: &[ArgSpec::Path] },
    CommandSpec { group: G_WB, name: "pdf", desc: "导出 PDF", args: &[ArgSpec::Path] },
    CommandSpec { group: G_WB, name: "close", desc: "关闭当前工作簿", args: &[ArgSpec::OptFixed(YES_NO)] },
    // 工作表
    CommandSpec { group: G_WS, name: "sheets", desc: "列出工作表（序号/名称/颜色/可见性）", args: &[] },
    CommandSpec { group: G_WS, name: "sheet", desc: "选定工作表", args: &[ArgSpec::Dynamic(sheet_candidates, "<序号|名称>")] },
    CommandSpec { group: G_WS, name: "addsheet", desc: "末尾新增工作表", args: &[ArgSpec::Any("名称")] },
    CommandSpec { group: G_WS, name: "info", desc: "当前上下文摘要", args: &[] },
    // 单元格
    CommandSpec { group: G_CELL, name: "get", desc: "读单元格/区域（多格输出表格）", args: &[ArgSpec::Any("地址")] },
    CommandSpec { group: G_CELL, name: "set", desc: "写单元格", args: &[ArgSpec::Any("地址"), ArgSpec::Any("文本")] },
    CommandSpec { group: G_CELL, name: "tget", desc: "区域批量读（TSV 输出）", args: &[ArgSpec::Any("区域")] },
    CommandSpec { group: G_CELL, name: "tset", desc: "区域批量写（逐行输入 TSV，单独 . 结束）", args: &[ArgSpec::Any("锚点")] },
    CommandSpec { group: G_CELL, name: "formula", desc: "写公式", args: &[ArgSpec::Any("地址"), ArgSpec::Any("公式")] },
    CommandSpec { group: G_CELL, name: "fmt", desc: "数字格式（如 0.00%）", args: &[ArgSpec::Any("地址"), ArgSpec::Any("数字格式")] },
    CommandSpec { group: G_CELL, name: "font", desc: "字体属性读取/设置", args: &[ArgSpec::Any("地址"), ArgSpec::Fixed(FONT_PROPS), ArgSpec::OptAny("值")] },
    CommandSpec { group: G_CELL, name: "merge", desc: "合并区域", args: &[ArgSpec::Any("区域")] },
    CommandSpec { group: G_CELL, name: "unmerge", desc: "取消合并", args: &[ArgSpec::Any("区域")] },
    CommandSpec { group: G_CELL, name: "chart", desc: "创建图表", args: &[ArgSpec::Fixed(CHART_TYPES), ArgSpec::Any("数据区"), ArgSpec::Any("锚点")] },
    CommandSpec { group: G_WS, name: "page", desc: "页面设置（无参数查看概览）", args: &[ArgSpec::Sub(&[
        ("margins", "页边距（磅，上 下 左 右 [页眉 页脚]）", &[ArgSpec::Any("上磅"), ArgSpec::Any("下磅"), ArgSpec::Any("左磅"), ArgSpec::Any("右磅"), ArgSpec::OptAny("页眉磅"), ArgSpec::OptAny("页脚磅")]),
        ("paper", "纸张大小", &[ArgSpec::Fixed(PAPERS)]),
        ("orient", "纸张方向", &[ArgSpec::Fixed(ORIENTS)]),
        ("area", "打印区域", &[ArgSpec::Any("区域|off")]),
        ("zoom", "打印缩放百分比", &[ArgSpec::Any("百分比|off")]),
        ("fit", "适应页宽/页高", &[ArgSpec::Any("宽页数"), ArgSpec::Any("高页数")]),
        ("header", "页眉", &[ArgSpec::Fixed(HF_POS), ArgSpec::Any("文本")]),
        ("footer", "页脚", &[ArgSpec::Fixed(HF_POS), ArgSpec::Any("文本")]),
        ("titles", "打印标题行/列", &[ArgSpec::Fixed(TITLE_KINDS), ArgSpec::Any("区域|off")]),
    ])] },
    // VBA
    CommandSpec {
        group: G_VBA,
        name: "vba",
        desc: "VBA 能力检测 / 运行宏 / 注入执行",
        args: &[ArgSpec::Sub(&[
            ("run", "运行已存在的宏", &[ArgSpec::Any("宏名"), ArgSpec::OptAny("参数")]),
            ("exec", "注入 bas 文件执行", &[ArgSpec::Path, ArgSpec::Any("入口过程名")]),
        ])],
    },
    // 应用与窗口
    CommandSpec { group: G_WIN, name: "wins", desc: "列出窗口", args: &[] },
    CommandSpec { group: G_WIN, name: "win", desc: "激活窗口", args: &[ArgSpec::Dynamic(window_candidates, "<窗口序号|标题>")] },
    CommandSpec { group: G_WIN, name: "visible", desc: "显示/隐藏应用窗口", args: &[ArgSpec::Fixed(ON_OFF)] },
    CommandSpec { group: G_WIN, name: "alerts", desc: "开关弹窗告警", args: &[ArgSpec::Fixed(ON_OFF)] },
    CommandSpec { group: G_WIN, name: "state", desc: "查看/设置窗口状态（无参数=查看）", args: &[ArgSpec::OptFixed(WINDOW_STATES)] },
    // 扩展
    CommandSpec { group: G_EXT, name: "raw", desc: "直接调用 COM 成员（测试未封装 API）", args: &[ArgSpec::Fixed(RAW_KINDS), ArgSpec::Any("成员"), ArgSpec::OptAny("参数")] },
    // 别名（group 为空，不进帮助）
    CommandSpec { group: "", name: "h", desc: "同 help", args: &[] },
    CommandSpec { group: "", name: "?", desc: "同 help", args: &[] },
    CommandSpec { group: "", name: "exit", desc: "同 quit", args: &[] },
];

/// 按名查命令。
pub fn find_command(name: &str) -> Option<&'static CommandSpec> {
    COMMANDS.iter().find(|c| c.name == name)
}

// ── help 生成（与补全同源，永不失同步） ──

/// 由命令表生成帮助文本。
pub fn help_text() -> String {
    let mut out = String::from(
        "命令列表（提示符 = 引擎|工作簿|工作表；Tab 打开候选菜单，方向键导航，回车确认）：\n",
    );
    let mut group = "";
    for c in COMMANDS {
        if c.group.is_empty() {
            continue;
        }
        if c.group != group {
            out.push_str(&format!("\n── {} ──\n", c.group));
            group = c.group;
        }
        let mut usage = c.name.to_string();
        for a in c.args {
            usage.push(' ');
            usage.push_str(&a.usage());
        }
        // CJK 按显示宽度 2 列计算，保证描述列对齐
        let pad = 30usize.saturating_sub(display_width(&usage));
        out.push_str(&format!("  {}{} {}\n", usage, " ".repeat(pad), c.desc));
    }
    out
}

/// 估算终端显示宽度：ASCII 1 列，其余（CJK 等）2 列。
fn display_width(s: &str) -> usize {
    s.chars().map(|c| if c.is_ascii() { 1 } else { 2 }).sum()
}

impl ArgSpec {
    /// 帮助中的用法片段。
    fn usage(&self) -> String {
        match self {
            Self::Fixed(f) => f.iter().map(|(v, _)| *v).collect::<Vec<_>>().join("|"),
            Self::OptFixed(f) => {
                let v = f.iter().map(|(v, _)| *v).collect::<Vec<_>>().join("|");
                format!("[{v}]")
            }
            Self::Dynamic(_, label) => label.to_string(),
            Self::Path => "<路径>".to_string(),
            Self::Any(p) => format!("<{p}>"),
            Self::OptAny(p) => format!("[{p}]"),
            Self::Sub(branches) => branches
                .iter()
                .map(|(n, _, args)| {
                    let mut u = n.to_string();
                    for a in *args {
                        u.push(' ');
                        u.push_str(&a.usage());
                    }
                    u
                })
                .collect::<Vec<_>>()
                .join(" | "),
        }
    }
}

// ── 动态候选 ──

/// 已打开工作簿候选：每簿一行「序号. 名称」（显示合并、插入只取索引）。
/// 未连接时经 `st.app()` 自动接管应用，无需先执行 list/wins。
fn workbook_candidates(st: &mut State) -> Vec<String> {
    let Some(app) = st.app().ok() else { return vec![] };
    match app.workbooks() {
        Ok(wbs) => wbs.iter().map(|w| format!("{}. {}", w.index, w.name)).collect(),
        Err(_) => vec![],
    }
}

/// 当前工作簿的工作表候选：每表一行「序号. 名称」（须先选定工作簿）。
fn sheet_candidates(st: &mut State) -> Vec<String> {
    let Some(wb) = &st.wb else { return vec![] };
    match wb.sheets() {
        Ok(sheets) => sheets.iter().map(|s| format!("{}. {}", s.index, s.name)).collect(),
        Err(_) => vec![],
    }
}

/// 应用窗口候选：每窗一行「序号. 标题」。
/// 未连接时经 `st.app()` 自动接管应用，无需先执行 wins。
fn window_candidates(st: &mut State) -> Vec<String> {
    let Some(app) = st.app().ok() else { return vec![] };
    match app.windows() {
        Ok(wins) => wins.iter().map(|w| format!("{}. {}", w.index, w.caption)).collect(),
        Err(_) => vec![],
    }
}

/// 引擎进程候选：每进程一行「PID. 名称」（纯 Win32 查询，不依赖连接）。
fn process_candidates(st: &mut State) -> Vec<String> {
    match processes(st.engine) {
        Ok(ps) => ps.iter().map(|p| format!("{}. {}", p.pid, p.name)).collect(),
        Err(_) => vec![],
    }
}

/// 文件路径候选：扫描当前 token 所在目录，目录候选自动带 `\` 便于续接。
/// 显示只给文件名，插入值是完整路径；含空格路径无法作为单 token 输入（REPL 以空白分参）。
fn path_suggestions(partial: &str, span: Span) -> Vec<Suggestion> {
    let (dir, prefix, base) = match partial.rfind(['\\', '/']) {
        Some(i) => (&partial[..=i], &partial[i + 1..], partial[..=i].to_string()),
        None => (".", partial, String::new()),
    };
    let Ok(entries) = std::fs::read_dir(dir) else {
        return vec![];
    };
    let lower = prefix.to_lowercase();
    entries
        .flatten()
        .filter(|e| e.file_name().to_string_lossy().to_lowercase().starts_with(&lower))
        .map(|e| {
            let name = e.file_name().to_string_lossy().to_string();
            let is_dir = e.file_type().map(|t| t.is_dir()).unwrap_or(false);
            Suggestion {
                // 目录带 `\` 续接子路径，不加确认空格；文件加空格表示确认
                value: if is_dir {
                    format!("{base}{name}\\")
                } else {
                    format!("{base}{name} ")
                },
                display_override: Some(name),
                span,
                ..Default::default()
            }
        })
        .collect()
}

// ── 补全引擎 ──

/// 当前输入位的补全来源。
enum Source {
    /// 命令名列表
    Commands,
    /// 文件路径
    Path,
    /// 自由文本（不补全）
    Free,
    /// 候选列表
    List(Vec<String>),
}

/// 候选是否匹配已输入前缀：整体前缀，或任一空白分段前缀
///（如 `1. 工作簿1` 支持按 `1` 或 `工作` 过滤）。
fn matches_candidate(c: &str, lower: &str) -> bool {
    c.to_lowercase().starts_with(lower)
        || c.split_whitespace().any(|t| t.to_lowercase().starts_with(lower))
}

/// 插入值：取首段并去掉尾部 `.`（`1. 工作簿1` → `1`），固定候选（`max`）原样返回。
fn insert_value(c: &str) -> &str {
    c.split_whitespace()
        .next()
        .unwrap_or(c)
        .trim_end_matches('.')
}

/// 补全器：把当前输入位置解析为 [Source]，交给 reedline 菜单展示。
pub struct MenuCompleter {
    state: Rc<RefCell<State>>,
}

// SAFETY：State 内含 COM 接口指针（本非 Send），但本工具是 COM STA 单线程模型：
// reedline 的事件循环与补全调用始终发生在创建行编辑器的同一线程上
//（补全器使用默认 poll_completion == Idle，不参与后台补全），指针永远不会跨线程移动。
unsafe impl Send for MenuCompleter {}

impl MenuCompleter {
    pub fn new(state: Rc<RefCell<State>>) -> Self {
        Self { state }
    }

    /// 解析当前输入位置对应的补全来源。返回 (当前 token 起始位置, 来源)；命令未知返回 None。
    fn source_at(&self, line: &str, pos: usize) -> Option<(usize, Source)> {
        let before = &line[..pos];
        let token_start = before.rfind(char::is_whitespace).map_or(0, |i| i + 1);
        let tokens: Vec<&str> = before[..token_start].split_whitespace().collect();

        // 位置 0：命令名
        let Some(cmd_name) = tokens.first() else {
            return Some((token_start, Source::Commands));
        };
        let spec = find_command(cmd_name)?;
        let arg_index = tokens.len() - 1;

        // 首参数位是子命令分支（connect/vba/page）：位置 0 = 分支名，位置 ≥1 = 分支内参数
        if let Some(ArgSpec::Sub(branches)) = spec.args.first() {
            if arg_index == 0 {
                // 候选带说明（菜单显示整行，插入只取分支名）
                let cands = branches
                    .iter()
                    .map(|(n, d, _)| format!("{n:<10}{d}"))
                    .collect();
                return Some((token_start, Source::List(cands)));
            }
            let branch = branches.iter().find(|(n, _, _)| *n == tokens[1]);
            let Some(arg) = branch.and_then(|(_, _, branch_args)| branch_args.get(arg_index - 1)) else {
                return Some((token_start, Source::Free));
            };
            return Some((token_start, self.classify(arg)));
        }

        // 普通命令：按参数位取规格
        let Some(arg) = spec.args.get(arg_index) else {
            return Some((token_start, Source::Free));
        };
        Some((token_start, self.classify(arg)))
    }

    /// ArgSpec → 补全来源。
    fn classify(&self, arg: &ArgSpec) -> Source {
        match arg {
            // 候选带说明（菜单显示整行，插入只取值）
            ArgSpec::Fixed(f) | ArgSpec::OptFixed(f) => {
                Source::List(f.iter().map(|(v, d)| format!("{v:<10}{d}")).collect())
            }
            ArgSpec::Dynamic(f, _) => match self.state.try_borrow_mut() {
                // 会话状态被占用时（如 tset 多行输入中）安静退出，避免 panic
                Ok(mut st) => Source::List(f(&mut st)),
                Err(_) => Source::Free,
            },
            ArgSpec::Path => Source::Path,
            ArgSpec::Any(_) | ArgSpec::OptAny(_) | ArgSpec::Sub(_) => Source::Free,
        }
    }
}

impl Completer for MenuCompleter {
    fn complete(&mut self, line: &str, pos: usize) -> CompletionResult {
        // 无候选时的占位项（插入值 = 原文，不改变输入），
        // 覆盖 reedline 硬编码的英文 "NO RECORDS FOUND"
        let fallback = |token_start: usize| {
            CompletionResult::fresh(vec![Suggestion {
                value: line[token_start..pos].to_string(),
                display_override: Some("（无可用候选）".to_string()),
                span: Span::new(token_start, pos),
                ..Default::default()
            }])
        };
        let Some((token_start, src)) = self.source_at(line, pos) else {
            // 未知命令等无规格场景
            return fallback(line[..pos].rfind(char::is_whitespace).map_or(0, |i| i + 1));
        };
        let partial = &line[token_start..pos];
        let lower = partial.to_lowercase();
        let span = Span::new(token_start, pos);
        let suggestions: Vec<Suggestion> = match src {
            Source::Free => vec![],
            Source::Path => path_suggestions(partial, span),
            Source::Commands => {
                // 上下文级别：0 无工作簿 → 1 已选工作簿 → 2 已选工作表
                let (ctx, preferred) = match self.state.try_borrow() {
                    Ok(s) if s.sheet.is_some() => (2u8, SHEET_CTX_CMDS),
                    Ok(s) if s.wb.is_some() => (1, SHEET_CMDS),
                    _ => (0, WB_MGMT_CMDS),
                };
                let mut specs: Vec<&CommandSpec> = COMMANDS
                    .iter()
                    // 别名含非字母字符的不参与命令补全；上下文不足的命令隐藏
                    .filter(|c| c.name.bytes().all(|b| b.is_ascii_alphanumeric()))
                    .filter(|c| required_ctx(c.name) <= ctx)
                    .collect();
                // 上下文相关命令置顶（稳定排序，其余保持命令表原顺序）
                specs.sort_by_key(|c| !preferred.contains(&c.name));
                specs.iter()
                    .filter(|c| c.name.starts_with(&lower))
                    .map(|c| Suggestion {
                        // 接受候选后补一个空格表示确认（目录路径除外）
                        value: format!("{} ", c.name),
                        display_override: Some(format!("{:<10} {}", c.name, c.desc)),
                        span,
                        ..Default::default()
                    })
                    .collect()
            }
            Source::List(cands) => cands
                .into_iter()
                .filter(|c| matches_candidate(c, &lower))
                .map(|c| Suggestion {
                    // 插入索引（带确认空格），整行「序号. 名称」仅作显示，避免同一项拆成多行
                    value: format!("{} ", insert_value(&c)),
                    display_override: Some(c),
                    span,
                    ..Default::default()
                })
                .collect(),
        };
        if suggestions.is_empty() {
            return fallback(token_start);
        }
        CompletionResult::fresh(suggestions)
    }
}

// ── 行编辑器组装 ──

/// REPL 行编辑器类型。
pub type ReplEditor = Reedline;

/// 动态提示符（引擎|工作簿|工作表，随会话状态变化）。
pub struct DynPrompt {
    pub text: String,
}

impl Prompt for DynPrompt {
    fn render_prompt_left(&self) -> Cow<'_, str> {
        Cow::Borrowed(&self.text)
    }

    fn render_prompt_right(&self) -> Cow<'_, str> {
        Cow::Borrowed("")
    }

    fn render_prompt_indicator(&self, _prompt_mode: PromptEditMode) -> Cow<'_, str> {
        Cow::Borrowed("")
    }

    fn render_prompt_multiline_indicator(&self) -> Cow<'_, str> {
        Cow::Borrowed(":: ")
    }

    fn render_prompt_history_search_indicator(
        &self,
        _search: PromptHistorySearch,
    ) -> Cow<'_, str> {
        Cow::Borrowed("历史搜索: ")
    }
}

/// 构建带候选菜单的行编辑器：
/// 候选**纵向单列**显示在输入行下方，选中项反色高亮，`Tab` 打开/循环、
/// `Shift+Tab` 反向、方向键导航、回车接受（均为 reedline 默认键位语义）。
pub fn editor(state: Rc<RefCell<State>>) -> ReplEditor {
    let mut keybindings = default_emacs_keybindings();
    // Tab：首次打开菜单，菜单激活时循环下一候选
    keybindings.add_binding(
        KeyModifiers::NONE,
        KeyCode::Tab,
        ReedlineEvent::UntilFound(vec![
            ReedlineEvent::Menu("completion_menu".to_string()),
            ReedlineEvent::MenuNext,
        ]),
    );
    let menu = ColumnarMenu::default()
        .with_name("completion_menu")
        .with_columns(1); // 单列纵向，每行一个候选
    Reedline::create()
        .with_completer(Box::new(MenuCompleter::new(state)))
        .with_menu(ReedlineMenu::EngineCompleter(Box::new(menu)))
        .with_edit_mode(Box::new(Emacs::new(keybindings)))
}

/// 续行输入（tset 多行等）：无补全菜单的简单行编辑；中断视为输入结束。
pub fn read_continuation(rl: &mut Reedline, prompt: &str) -> anyhow::Result<String> {
    match rl.read_line(&DynPrompt {
        text: prompt.to_string(),
    })? {
        Signal::Success(buffer) => Ok(buffer),
        _ => Ok(".".to_string()),
    }
}
