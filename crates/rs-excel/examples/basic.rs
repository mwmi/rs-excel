//! AI 集成参考示例：新建 → 批量写 → 样式 → 公式 → 读回 → 图表 → 保存 → VBA 检测 → 关闭。
//!
//! 运行：`cargo run -p rs-excel --example basic`（加 `--excel` 驱动 MS Excel）。

use anyhow::Result;
use rs_excel::prelude::*;

fn main() -> Result<()> {
    let engine = if std::env::args().any(|a| a == "--excel") {
        Engine::Excel
    } else {
        Engine::Wps
    };

    // 1. 连接（New 会拉起新进程；Attach 接管已打开实例）
    let app = App::connect(engine, ConnectMode::New)?;
    app.set_visible(true)?;
    app.set_display_alerts(false)?; // 关弹窗，避免 SaveAs 卡模态框

    // 2. 新建工作簿，一段 TSV 写入一块区域
    let wb = app.new_workbook()?;
    let sheet = wb.active_sheet()?;
    sheet.write_tsv("A1", "姓名\t数量\n苹果\t12\n香蕉\t34\n樱桃\t8")?;

    // 3. 样式：标题加粗 + 表体边框 + 填充
    sheet
        .range("A1:B1")?
        .set_font(Some(true), None, None, None, None)?;
    sheet.range("A1:B4")?.borders()?.set_all(1, 2, rgb(0, 0, 0))?;
    sheet.range("A1:B1")?.interior()?.set_color(rgb(220, 235, 255))?;

    // 4. 公式 + 读回
    sheet.range("B5")?.set_formula("=SUM(B2:B4)")?;
    println!("B5 公式: {}", sheet.range("B5")?.formula()?);
    println!("B5 数值: {}", sheet.range("B5")?.value()?);

    // 5. 区域批量读（TSV 只给 CLI / 人工；AI 用 ai::Matrix）
    println!("A1:B4 内容:\n{}", sheet.read_tsv("A1:B4")?);

    // 5.5 批量模式（RAII：退出作用域恢复调用前的刷新/事件/弹窗/计算）+ evaluate
    {
        let _guard = app.fast_guard()?;
        println!("已用区域: {}", sheet.used_addr()?);
        let total = app.evaluate("SUM(B2:B4)")?;
        println!("evaluate SUM(B2:B4) = {total}");
    }

    // 5.6 AI 门面：自然类型批量写（字符串 "001" 保持文本，日期写成真日期）
    let grid = vec![
        vec![CellInput::Text("001".into()), CellInput::Number(12.0)],
        vec![CellInput::Date("2026-01-05".into()), CellInput::Empty],
    ];
    let block = sheet.range("H1")?.resize(2, 2)?;
    let report = rs_excel::ai::write_grid(&block, &grid)?;
    println!("写入 {}（{} 格，日期 {} 格）", report.addr, report.cells, report.dates);

    // 6. 图表
    let chart = sheet.add_chart(ChartKind::Column, "A1:B4", "D2")?;
    chart.set_title("水果数量")?;

    // 7. VBA 能力检测（简单操作走内置 API，复杂操作走 VBA）
    let vba = Vba::new(&app);
    let st = vba.status();
    println!(
        "VBA 环境: {}，信任访问: {}",
        if st.env_installed { "已安装" } else { "未安装" },
        if st.trust_access { "已开启" } else { "未开启" }
    );

    // 8. 保存并关闭（覆盖已有文件由库按策略备份处理，不需要调用方自己删文件）
    let out = std::env::current_dir()?.join("rs_excel_demo.xlsx");
    let report = wb.save_as(
        &out.to_string_lossy(),
        FileFormat::Xlsx,
        OverwritePolicy::Replace,
    )?;
    println!("已保存: {}", report.path);
    wb.close(false)?;
    app.quit()?;
    Ok(())
}
