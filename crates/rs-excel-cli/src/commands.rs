//! REPL 命令解析与执行。
//!
//! 命令的补全规格与帮助文本统一由 completion.rs 的 [COMMANDS] 表定义；
//! 新增命令时在 completion.rs 加一条表项，再在此处加一个 match 分支。

use std::cell::RefCell;

use anyhow::{bail, Result};

use rs_excel::prelude::*;

use crate::completion::{ReplEditor, help_text, read_continuation};

/// REPL 会话状态。
pub struct State {
    pub engine: Engine,
    pub app: Option<App>,
    pub wb: Option<Workbook>,
    pub sheet: Option<Worksheet>,
    wb_label: String,
    sheet_label: String,
}

impl State {
    pub fn new() -> Self {
        Self {
            engine: Engine::Wps,
            app: None,
            wb: None,
            sheet: None,
            wb_label: String::new(),
            sheet_label: String::new(),
        }
    }

    pub fn prompt(&self) -> String {
        // 始终显示当前引擎（未连接也显示，engine 切换即时生效）
        match (&self.wb, &self.sheet) {
            (None, _) => format!("{}> ", self.engine.label()),
            (Some(_), None) => format!("{}|{}> ", self.engine.label(), self.wb_label),
            (Some(_), Some(_)) => {
                format!("{}|{}|{}> ", self.engine.label(), self.wb_label, self.sheet_label)
            }
        }
    }

    /// 获取应用连接；未连接时自动接管（attach）已打开实例。
    ///
    /// list/use 等"先发现后绑定"的流程无需先 connect：
    /// attach 是只读接管（ROT），不干扰用户正在使用的表格；出错时返回带定位提示的错误。
    pub(crate) fn app(&mut self) -> Result<&App> {
        if self.app.is_none() {
            self.app = Some(App::connect(self.engine, ConnectMode::Attach)?);
        }
        Ok(self.app.as_ref().unwrap())
    }

    fn wb(&self) -> Result<&Workbook> {
        self.wb
            .as_ref()
            .ok_or_else(|| anyhow::anyhow!("未选定工作簿：先执行 use <序号|名称>"))
    }

    fn sheet(&self) -> Result<&Worksheet> {
        self.sheet
            .as_ref()
            .ok_or_else(|| anyhow::anyhow!("未选定工作表：先执行 sheet <序号|名称>"))
    }

    fn clear_wb(&mut self) {
        self.sheet = None;
        self.sheet_label.clear();
        self.wb = None;
        self.wb_label.clear();
    }

    fn set_wb(&mut self, wb: Workbook) -> Result<()> {
        self.wb_label = wb.name()?;
        self.clear_sheet();
        self.wb = Some(wb);
        Ok(())
    }

    fn clear_sheet(&mut self) {
        self.sheet = None;
        self.sheet_label.clear();
    }

    fn set_sheet(&mut self, sh: Worksheet) -> Result<()> {
        self.sheet_label = sh.name()?;
        self.sheet = Some(sh);
        Ok(())
    }
}

/// 执行一条命令。返回 false 表示退出 REPL（告别语由 main 统一输出）。
pub fn execute(st: &RefCell<State>, line: &str, rl: &mut ReplEditor) -> bool {
    let cmd = first(line);
    if cmd == "quit" || cmd == "exit" {
        return false;
    }
    let args = rest(line);
    let mut state = st.borrow_mut();
    if let Err(e) = run(&mut state, cmd, args, rl) {
        println!("错误: {e:#}");
    }
    true
}

fn run(st: &mut State, cmd: &str, args: &str, rl: &mut ReplEditor) -> Result<()> {
    match cmd {
        "help" | "h" | "?" => print!("{}", help_text()),
        "engine" => {
            if args.is_empty() {
                println!("当前引擎: {}（connect 后生效）", st.engine.label());
            } else {
                st.engine = match args {
                    "excel" => Engine::Excel,
                    "wps" => Engine::Wps,
                    other => bail!("未知引擎 '{other}'，可选 wps|excel"),
                };
                st.clear_wb();
                st.app = None;
                println!("已切换为 {}，重新 connect 生效", st.engine.label());
            }
        }
        "connect" => {
            let mut parts = args.split_whitespace();
            let mode_str = parts.next().unwrap_or("");
            let had = st.app.is_some();
            let mut extra = String::new();
            let app = match mode_str {
                "new" => App::connect(st.engine, ConnectMode::New)?,
                "attach" => {
                    let app = App::connect(st.engine, ConnectMode::Attach)?;
                    // 可选绑定 wins 列表中的指定窗口（序号或标题）
                    if let Some(w) = parts.next() {
                        app.activate_window(w)?;
                        extra = format!("，已绑定窗口 {w}");
                    }
                    app
                }
                other => bail!("用法: connect attach [窗口序号|标题] / connect new（实际输入 '{other}'）"),
            };
            let cap = app.caption().unwrap_or_default();
            st.clear_wb(); // 重连必清工作簿/工作表上下文（断开旧连接）
            if had {
                println!("已断开上一个连接");
            }
            st.app = Some(app);
            println!("已连接: {cap}（{mode_str} 模式{extra}）");
        }
        "list" => {
            let app = st.app()?;
            println!(
                "已自动接管: {}（{} 引擎；如需新建实例用 connect new）",
                app.caption().unwrap_or_default(),
                app.engine().label()
            );
            let wbs = app.workbooks()?;
            if wbs.is_empty() {
                println!("当前没有打开的工作簿");
            }
            for wb in wbs {
                println!(
                    "  {}{}. {}  路径: {}  已保存: {}  工作表数: {}",
                    if wb.is_active { "[活动]" } else { "" },
                    wb.index,
                    wb.name,
                    if wb.full_name.is_empty() { "-" } else { &wb.full_name },
                    if wb.saved { "是" } else { "否" },
                    wb.sheet_count
                );
            }
        }
        "use" => {
            if args.is_empty() {
                bail!("用法: use <序号|名称>");
            }
            let wb = st.app()?.workbook(args)?;
            st.set_wb(wb)?;
            println!("已选定工作簿: {}", st.wb_label);
        }
        "new" => {
            let wb = st.app()?.new_workbook()?;
            st.set_wb(wb)?;
            println!("已新建工作簿: {}", st.wb_label);
        }
        "open" => {
            if args.is_empty() {
                bail!("用法: open <路径>（相对/绝对均可，自动转为绝对路径后打开）");
            }
            let wb = st.app()?.open_workbook(args)?;
            st.set_wb(wb)?;
            println!("已打开工作簿: {}", st.wb_label);
        }
        "save" => {
            st.wb()?.save()?;
            println!("已保存");
        }
        "saveas" => {
            if args.is_empty() {
                bail!("用法: saveas <绝对路径>（按扩展名自动选格式）");
            }
            st.wb()?.save_as(args, infer_format(args))?;
            if let Some(wb) = &st.wb {
                if let Ok(name) = wb.name() {
                    st.wb_label = name;
                }
            }
            println!("已另存为: {args}");
        }
        "pdf" => {
            if args.is_empty() {
                bail!("用法: pdf <绝对路径>");
            }
            st.wb()?.export_pdf(args)?;
            println!("已导出 PDF: {args}");
        }
        "close" => {
            let save = args != "n";
            st.wb()?.close(save)?;
            st.clear_wb();
            println!("已关闭工作簿（保存: {}）", if save { "是" } else { "否" });
        }
        "appquit" => {
            st.app()?.quit()?;
            st.clear_wb();
            st.app = None;
            println!("应用已退出，连接已断开");
        }
        "sheets" => {
            let sheets = st.wb()?.sheets()?;
            if sheets.is_empty() {
                println!("没有工作表");
            }
            for s in sheets {
                println!(
                    "  {}. {}  颜色: {}  可见: {}",
                    s.index,
                    s.name,
                    if s.tab_color_hex.is_empty() { "-" } else { &s.tab_color_hex },
                    s.visible
                );
            }
        }
        "sheet" => {
            if args.is_empty() {
                bail!("用法: sheet <序号|名称>");
            }
            let sh = st.wb()?.sheet(args)?;
            st.set_sheet(sh)?;
            println!("已选定工作表: {}", st.sheet_label);
        }
        "addsheet" => {
            if args.is_empty() {
                bail!("用法: addsheet <名称>");
            }
            let sh = st.wb()?.add_sheet(args, None)?;
            println!("已新增工作表: {}", sh.name()?);
        }
        "info" => {
            let app = st.app()?;
            println!(
                "引擎: {}  应用: {}",
                app.engine().label(),
                app.caption().unwrap_or_default()
            );
            let wb = st.wb()?;
            println!(
                "工作簿: {}  完整路径: {}  已保存: {}",
                st.wb_label,
                wb.full_name().unwrap_or_default(),
                if wb.is_saved().unwrap_or(false) { "是" } else { "否" }
            );
            let sh = st.sheet()?;
            println!(
                "工作表: {}（第 {} 张）  标签颜色: {}",
                st.sheet_label,
                sh.index()?,
                sh.tab_color().unwrap_or_default()
            );
        }
        "get" => {
            let r = st.sheet()?.range(args)?;
            let (nr, nc) = (r.rows()?, r.columns()?);
            if nr * nc > 1 {
                // 多格区域：表格化输出（Tab 分列、换行分行）
                println!("{args} =\n{}", r.read_tsv()?);
            } else {
                println!("{args} = {}", r.value()?);
            }
        }
        "set" => {
            let (addr, text) = args
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: set <地址> <文本>"))?;
            if text.is_empty() {
                bail!("用法: set <地址> <文本>");
            }
            st.sheet()?.range(addr.trim())?.set_value(text)?;
            println!("已写入 {}（{}）", addr.trim(), text);
        }
        "tget" => {
            if args.is_empty() {
                bail!("用法: tget <区域>（如 A1:C3）");
            }
            println!("{}", st.sheet()?.read_tsv(args)?);
        }
        "tset" => {
            if args.is_empty() {
                bail!("用法: tset <锚点地址>（如 A1），随后逐行输入 TSV，单独一行 . 结束");
            }
            println!("逐行输入 TSV 内容（Tab 分列），单独输入 . 结束：");
            let mut buf = String::new();
            loop {
                let l = read_continuation(rl, "tset> ")?;
                if l.trim() == "." {
                    break;
                }
                buf.push_str(&l);
                buf.push('\n');
            }
            st.sheet()?.write_tsv(args, &buf)?;
            println!("已写入 {}", st.sheet_label);
        }
        "formula" => {
            let (addr, f) = args
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: formula <地址> <公式>"))?;
            if f.is_empty() {
                bail!("用法: formula <地址> <公式>");
            }
            st.sheet()?.range(addr.trim())?.set_formula(f)?;
            println!("已写入公式 {}（{f}）", addr.trim());
        }
        "fmt" => {
            let (addr, fmt) = args
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: fmt <地址> <数字格式>"))?;
            if fmt.is_empty() {
                bail!("用法: fmt <地址> <数字格式>");
            }
            st.sheet()?.range(addr.trim())?.number_format(fmt)?;
            println!("已设置 {} 数字格式: {fmt}", addr.trim());
        }
        "font" => font_cmd(st, args)?,
        "merge" => {
            st.sheet()?.range(args)?.merge()?;
            println!("已合并 {args}");
        }
        "unmerge" => {
            st.sheet()?.range(args)?.unmerge()?;
            println!("已取消合并 {args}");
        }
        "chart" => {
            let p: Vec<&str> = args.split_whitespace().collect();
            if p.len() != 3 {
                bail!("用法: chart <类型> <数据区> <锚点>，类型: column|bar|line|pie|area|scatter|doughnut");
            }
            let kind = ChartKind::parse(p[0])
                .ok_or_else(|| anyhow::anyhow!("未知图表类型 '{}'，可选 column|bar|line|pie|area|scatter|doughnut", p[0]))?;
            st.sheet()?.add_chart(kind, p[1], p[2])?;
            println!("已创建 {:?} 图表（数据 {}，锚点 {}）", kind, p[1], p[2]);
        }
        "vba" => vba_cmd(st, args)?,
        "wins" => {
            let wins = st.app()?.windows()?;
            for w in &wins {
                println!(
                    "  {}. {}{}",
                    w.index,
                    w.caption,
                    if w.is_active { "  [活动]" } else { "" }
                );
            }
            if !wins.is_empty() {
                println!("（可用 connect attach <序号|标题> 直接绑定窗口）");
            }
        }
        "win" => {
            if args.is_empty() {
                bail!("用法: win <序号|标题>");
            }
            st.app()?.activate_window(args)?;
            println!("已激活窗口 {args}");
        }
        "visible" => match args {
            "on" => {
                st.app()?.set_visible(true)?;
                println!("已显示应用窗口");
            }
            "off" => {
                st.app()?.set_visible(false)?;
                println!("已隐藏应用窗口");
            }
            other => bail!("用法: visible on|off（得到 '{other}'）"),
        },
        "alerts" => match args {
            "on" => {
                st.app()?.set_display_alerts(true)?;
                println!("已开启弹窗告警");
            }
            "off" => {
                st.app()?.set_display_alerts(false)?;
                println!("已关闭弹窗告警");
            }
            other => bail!("用法: alerts on|off（得到 '{other}'）"),
        },
        "state" => {
            if args.is_empty() {
                println!("当前窗口状态: {}", st.app()?.window_state()?);
            } else {
                let s = match args {
                    "max" => WindowState::Maximized,
                    "min" => WindowState::Minimized,
                    "normal" => WindowState::Normal,
                    other => bail!("用法: state [max|min|normal]，无参数查看当前状态（得到 '{other}'）"),
                };
                st.app()?.set_window_state(s)?;
                println!("已设置窗口状态: {s}");
            }
        }
        "page" => page_cmd(st, args)?,
        "ps" => {
            // 纯 Win32 查询，不依赖 COM 连接（连接卡死时也能用）
            let ps = processes(st.engine)?;
            if ps.is_empty() {
                println!("当前没有 {} 进程", st.engine.exe_name());
            }
            for p in &ps {
                println!("  {p}");
            }
        }
        "kill" => {
            let pid: u32 = args
                .parse()
                .map_err(|_| anyhow::anyhow!("用法: kill <PID>（先用 ps 查看进程）"))?;
            kill(pid)?;
            println!("已结束进程 {pid}");
        }
        "killall" => {
            let n = kill_all(st.engine)?;
            println!("已结束 {} 个 {} 进程", n, st.engine.exe_name());
        }
        "raw" => raw_cmd(st, args)?,
        other => bail!("未知命令 '{other}'，输入 help 查看命令列表"),
    }
    Ok(())
}

fn font_cmd(st: &State, args: &str) -> Result<()> {
    let p: Vec<&str> = args.split_whitespace().collect();
    if p.len() < 2 {
        bail!("用法: font <地址> <bold|italic|underline|size|color|name> [值]（无值=读取）");
    }
    let r = st.sheet()?.range(p[0])?;
    let f = r.font()?;
    let prop = p[1];
    let val = p.get(2).copied();
    match (prop, val) {
        ("bold", Some(v)) => f.set_bold(parse_bool(v)?)?,
        ("bold", None) => println!("bold = {}", f.bold()?),
        ("italic", Some(v)) => f.set_italic(parse_bool(v)?)?,
        ("italic", None) => println!("italic = {}", f.italic()?),
        ("underline", Some(v)) => f.set_underline(parse_bool(v)?)?,
        ("underline", None) => println!("underline = {}", f.underline()?),
        ("size", Some(v)) => f.set_size(v.parse()?)?,
        ("size", None) => println!("size = {}", f.size()?),
        ("color", Some(v)) => f.set_color(parse_color(v)?)?,
        ("color", None) => println!("color = #{:06X}", f.color()?),
        ("name", Some(v)) => f.set_name(v)?,
        ("name", None) => println!("name = {}", f.name()?),
        (other, _) => bail!("未知字体属性 '{other}'，可选 bold|italic|underline|size|color|name"),
    }
    if val.is_some() {
        println!("已设置 {} 字体 {prop}", p[0]);
    }
    Ok(())
}

/// 页面设置命令：无参数查看概览，子命令修改（margins/paper/orient/area/zoom/fit/header/footer/titles）。
fn page_cmd(st: &State, args: &str) -> Result<()> {
    let ps = st.sheet()?.page_setup()?;
    if args.is_empty() {
        let disp = |v: String| {
            if v.is_empty() || v == "false" {
                "-".to_string()
            } else {
                v
            }
        };
        let (zw, zh) = ps.fit_pages().unwrap_or((0, 0));
        println!(
            "方向: {}  纸张: {}  缩放: {}  打印区域: {}",
            ps.orientation()?,
            PaperSize::name_of(ps.paper()?),
            match ps.zoom()? {
                0 => format!("适应页宽{zw}×页高{zh}"),
                z => format!("{z}%"),
            },
            disp(ps.print_area()?)
        );
        println!(
            "边距(磅): 上{} 下{} 左{} 右{} 页眉{} 页脚{}",
            ps.margin(Margin::Top)?,
            ps.margin(Margin::Bottom)?,
            ps.margin(Margin::Left)?,
            ps.margin(Margin::Right)?,
            ps.margin(Margin::Header)?,
            ps.margin(Margin::Footer)?
        );
        println!(
            "标题行: {}  标题列: {}",
            disp(ps.title_rows()?),
            disp(ps.title_cols()?)
        );
        println!(
            "页眉: 左[{}] 中[{}] 右[{}]",
            ps.header(Section::Left)?,
            ps.header(Section::Center)?,
            ps.header(Section::Right)?
        );
        println!(
            "页脚: 左[{}] 中[{}] 右[{}]",
            ps.footer(Section::Left)?,
            ps.footer(Section::Center)?,
            ps.footer(Section::Right)?
        );
        return Ok(());
    }
    let (sub, r1) = args
        .split_once(char::is_whitespace)
        .ok_or_else(|| anyhow::anyhow!("用法: page（无参数查看）/ page <margins|paper|orient|area|zoom|fit|header|footer|titles> ..."))?;
    let parse_num = |s: &str| -> Result<f64> {
        s.parse()
            .map_err(|_| anyhow::anyhow!("'{s}' 不是数字（边距/缩放单位为磅）"))
    };
    match sub {
        "margins" => {
            let p: Vec<&str> = r1.split_whitespace().collect();
            if p.len() < 4 {
                bail!("用法: page margins <上> <下> <左> <右> [页眉] [页脚]（单位：磅）");
            }
            ps.set_margins(
                parse_num(p[0])?,
                parse_num(p[1])?,
                parse_num(p[2])?,
                parse_num(p[3])?,
            )?;
            if p.len() >= 6 {
                ps.set_margin(Margin::Header, parse_num(p[4])?)?;
                ps.set_margin(Margin::Footer, parse_num(p[5])?)?;
            }
            println!("已设置页边距（磅）");
        }
        "paper" => {
            let size = PaperSize::parse(r1)
                .ok_or_else(|| anyhow::anyhow!("未知纸张 '{r1}'，可选 A4|A3|A5|B4|B5|Letter|Legal|Executive|A2|A1 或数字代码"))?;
            ps.set_paper(size)?;
            println!("已设置纸张: {}", PaperSize::name_of(size.code()));
        }
        "orient" => {
            let o = match r1 {
                "portrait" | "纵向" => Orientation::Portrait,
                "landscape" | "横向" => Orientation::Landscape,
                other => bail!("方向应为 portrait|landscape（得到 '{other}'）"),
            };
            ps.set_orientation(o)?;
            println!("已设置方向: {o}");
        }
        "area" => {
            if r1 == "off" {
                ps.clear_print_area()?;
                println!("已清除打印区域");
            } else {
                ps.set_print_area(r1)?;
                println!("已设置打印区域: {r1}");
            }
        }
        "zoom" => {
            if r1 == "off" {
                ps.zoom_reset()?;
                println!("已恢复 100% 缩放");
            } else {
                let z = parse_num(r1)? as i32;
                ps.set_zoom(z)?;
                println!("已设置打印缩放: {z}%");
            }
        }
        "fit" => {
            let p: Vec<&str> = r1.split_whitespace().collect();
            if p.len() != 2 {
                bail!("用法: page fit <宽页数> <高页数>（0 = 该方向不限）");
            }
            ps.fit_to(parse_num(p[0])? as i32, parse_num(p[1])? as i32)?;
            println!("已设置适应页宽{}×页高{}", p[0], p[1]);
        }
        "header" | "footer" => {
            let (sec, text) = r1
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: page {sub} <left|center|right> <文本>（支持 &P 页码 &D 日期 &F 文件名）"))?;
            let s = Section::parse(sec)
                .ok_or_else(|| anyhow::anyhow!("位置应为 left|center|right（得到 '{sec}'）"))?;
            if sub == "header" {
                ps.set_header(s, text)?;
            } else {
                ps.set_footer(s, text)?;
            }
            println!(
                "已设置页{} {s}: {text}",
                if sub == "header" { "眉" } else { "脚" }
            );
        }
        "titles" => {
            let (kind, addr) = r1
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: page titles <rows|cols> <区域|off>（行如 $1:$2，列如 $A:$B）"))?;
            let off = addr == "off";
            match (kind, off) {
                ("rows", false) => ps.set_title_rows(addr)?,
                ("rows", true) => ps.clear_title_rows()?,
                ("cols", false) => ps.set_title_cols(addr)?,
                ("cols", true) => ps.clear_title_cols()?,
                (other, _) => bail!("标题类型应为 rows|cols（得到 '{other}'）"),
            }
            println!(
                "已{}打印标题{}",
                if off { "清除" } else { "设置" },
                if kind == "rows" { "行" } else { "列" }
            );
        }
        other => bail!("未知子命令 '{other}'，可选 margins|paper|orient|area|zoom|fit|header|footer|titles"),
    }
    Ok(())
}

fn vba_cmd(st: &mut State, args: &str) -> Result<()> {
    if args.is_empty() {
        let s = Vba::new(st.app()?).status();
        println!(
            "VBA 环境: {}  信任访问: {}",
            if s.env_installed { "已安装" } else { "未安装（WPS 需 VBA 支持库）" },
            if s.trust_access { "已开启" } else { "未开启（信任中心 → 宏设置）" }
        );
        println!("用法: vba run <宏名> [参数...]  |  vba exec <bas文件> <入口过程名>");
        return Ok(());
    }
    let sub = first(args);
    let rest_args = rest(args);
    match sub {
        "run" => {
            let (name, rest2) = rest_args
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: vba run <宏名> [参数...]"))?;
            let vals: Vec<ComValue> = rest2.split_whitespace().map(parse_arg).collect();
            Vba::new(st.app()?).run_macro(name, &vals)?;
            println!("宏 {name} 执行完成");
        }
        "exec" => {
            let (file, entry) = rest_args
                .split_once(char::is_whitespace)
                .ok_or_else(|| anyhow::anyhow!("用法: vba exec <bas文件> <入口过程名>"))?;
            let code = std::fs::read_to_string(file)
                .map_err(|e| anyhow::anyhow!("读取 {file} 失败: {e}"))?;
            Vba::new(st.app()?).execute(&code, entry)?;
            println!("VBA 代码执行完成（入口 {entry}），临时模块已清理");
        }
        other => bail!("未知子命令 '{other}'，可选 run|exec"),
    }
    Ok(())
}

fn raw_cmd(st: &mut State, args: &str) -> Result<()> {
    let (kind, r1) = args
        .split_once(char::is_whitespace)
        .ok_or_else(|| anyhow::anyhow!("用法: raw <get|call|put> <成员> [参数...]"))?;
    let member = first(r1);
    let rest_args = rest(r1);
    let target: &IDispatchWrapper = if let Some(s) = &st.sheet {
        s.raw()
    } else if let Some(w) = &st.wb {
        w.raw()
    } else {
        st.app()?.raw()
    };
    let vals: Vec<ComValue> = rest_args.split_whitespace().map(parse_arg).collect();
    match kind {
        "get" => {
            let res = target.get(member)?;
            if res.is_object() {
                let obj = res.idispatch()?;
                println!("(COM 对象, 路径: {})", obj.ctx);
            } else {
                println!("{}", res.value());
            }
        }
        "call" => {
            let res = target.call(member, vals)?;
            if res.is_object() {
                let obj = res.idispatch()?;
                println!("(COM 对象, 路径: {})", obj.ctx);
            } else {
                println!("{}", res.value());
            }
        }
        "put" => {
            target.put(member, vals)?;
            println!("已写入 {member}");
        }
        other => bail!("用法: raw <get|call|put> <成员> [参数...]（得到 '{other}'）"),
    }
    Ok(())
}

fn first(s: &str) -> &str {
    s.split_whitespace().next().unwrap_or("")
}

fn rest(s: &str) -> &str {
    s.trim_start()
        .split_once(char::is_whitespace)
        .map(|(_, r)| r.trim())
        .unwrap_or("")
}

fn parse_arg(s: &str) -> ComValue {
    if s == "true" {
        true.into()
    } else if s == "false" {
        false.into()
    } else if let Ok(i) = s.parse::<i32>() {
        i.into()
    } else if let Ok(f) = s.parse::<f64>() {
        f.into()
    } else {
        s.to_string().into()
    }
}

fn parse_bool(s: &str) -> Result<bool> {
    match s {
        "true" | "1" | "on" | "是" => Ok(true),
        "false" | "0" | "off" | "否" => Ok(false),
        other => bail!("布尔值应为 true|false（得到 '{other}'）"),
    }
}

/// 颜色参数：`#RRGGBB`、`RRGGBB` 或 `r,g,b`
fn parse_color(s: &str) -> Result<u32> {
    let s = s.trim();
    if let Some(hex) = s.strip_prefix('#').or_else(|| {
        if s.len() == 6 && s.chars().all(|c| c.is_ascii_hexdigit()) {
            Some(s)
        } else {
            None
        }
    }) {
        let v = u32::from_str_radix(hex, 16).map_err(|_| anyhow::anyhow!("非法颜色 '{s}'"))?;
        return Ok(rgb((v >> 16) as u8, (v >> 8) as u8, v as u8));
    }
    let p: Vec<&str> = s.split(',').collect();
    if p.len() == 3 {
        return Ok(rgb(p[0].trim().parse()?, p[1].trim().parse()?, p[2].trim().parse()?));
    }
    bail!("颜色格式应为 #RRGGBB 或 r,g,b（得到 '{s}'）")
}

/// 按扩展名推断保存格式。
fn infer_format(path: &str) -> FileFormat {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".csv") {
        FileFormat::Csv
    } else if lower.ends_with(".xlsm") {
        FileFormat::Xlsm
    } else if lower.ends_with(".xls") {
        FileFormat::Xls
    } else {
        FileFormat::Xlsx
    }
}
