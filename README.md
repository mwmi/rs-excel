# rs-excel

Rust + COM 驱动 WPS/Excel 的表格自动化库（**Windows 专用**，基于 `IDispatch` 后期绑定）。
默认驱动 WPS（`KET.Application`），可切换 MS Excel（`Excel.Application`），API 完全一致。

```
crates/
├── rs-excel/        # 库本体
│   └── examples/basic.rs   # 端到端参考示例
└── rs-excel-cli/    # 交互式 REPL 测试工具
```

## 快速上手

```rust
use rs_excel::prelude::*;

let app = App::connect(Engine::Wps, ConnectMode::Attach)?;   // 接管已打开的 WPS
let wb = app.workbook("Book1")?;                             // 按序号("1")或名称绑定
let sheet = wb.active_sheet()?;

sheet.write_tsv("A1", "姓名\t数量\n苹果\t12\n香蕉\t34")?;     // 一段 TSV 写一块区域
println!("{}", sheet.read_tsv("A1:B3")?);                    // 一段 TSV 读回

let rng = sheet.range("A1:B1")?;                             // 联动链路
rng.set_font(Some(true), None, None, None, None)?;           // 加粗
rng.borders()?.set_all(1, 2, rgb(0, 0, 0))?;                 // 边框（实线/细线/黑色）
rng.set_alignment(Some(AlignH::Center), None, None)?;        // 居中

let chart = sheet.add_chart(ChartKind::Column, "A1:B3", "D2")?; // 图表
chart.set_title("数量统计")?;

wb.save_as(r"D:\out.xlsx", FileFormat::Xlsx)?;
wb.export_pdf(r"D:\out.pdf")?;
```

完整流程见 `cargo run -p rs-excel --example basic`（加 `--excel` 驱动 Excel）。

## API 地图

| 对象          | 获取方式                         | 常用能力                                                                                                                                                                                                                                                                                                      |
| ----------- | ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `App`       | `App::connect(engine, mode)` | `visible` / `display_alerts` / `window_state` / `windows()` / `workbooks()` / `workbook(n或名)` / `active_workbook()` / `new_workbook()` / `open_workbook(path)` / `quit()`                                                                                                                                 |
| `Workbook`  | `app.workbook(...)` 等        | `name/full_name/path/is_saved` / `sheets()` / `sheet(n或名)` / `active_sheet()` / `add_sheet(name, before)` / `save` / `save_as` / `export_pdf` / `close(save)`                                                                                                                                             |
| `Worksheet` | `wb.sheet(...)` 等            | `name/rename/index/tab_color/visible` / `range(addr)` / `cells(r,c)` / `used_range()` / `rows/columns` / `read_tsv/write_tsv` / `add_chart(...)` / `activate/select/delete`                                                                                                                               |
| `Range`     | `sheet.range("A1:C3")`       | 读：`value/values/text/read_tsv/formula/address/rows/columns`；写：`set_value/set_values/write_tsv/set_formula/number_format`；格式：`merge/unmerge/is_merged/set_alignment/col_width/row_height/auto_fit`；样式入口：`font()/interior()/borders()/set_font(...)`；清理：`clear/clear_contents/clear_formats`；`copy_to(dst)` |
| `Font`      | `range.font()`               | `bold/italic/underline/size/color/name/set_strikethrough`（get+set）                                                                                                                                                                                                                                        |
| `Interior`  | `range.interior()`           | `color/set_color/clear`                                                                                                                                                                                                                                                                                   |
| `Borders`   | `range.borders()`            | `set_all(线型,粗细,色)` / `set_outline(...)` / `clear`                                                                                                                                                                                                                                                         |
| `Chart`     | `sheet.add_chart(...)`       | `set_title/set_type/export_image/delete`                                                                                                                                                                                                                                                                  |
| `Vba`       | `Vba::new(&app)`             | `status()` / `run_macro(name, args)` / `execute(code, entry)`                                                                                                                                                                                                                                             |
| 底层          | 任意对象 `.raw()`                | `get/call/call_named/put/put_raw` —— 直调任意 COM 成员                                                                                                                                                                                                                                                          |

## AI 调用要点（省 token 设计）

1. **批量读写用 TSV**：`write_tsv(anchor, "a\tb\nc\td")` / `read_tsv(addr)` —— 一次调用处理整块区域，避免逐格循环。
2. **错误即定位**：所有错误单行输出四要素，AI 一次读取即可修复：
   ```
   [put('Value')] @ WPS > Book1.xlsx > Sheet1 > A1:C3 | 0x800A03EC ... | 提示: 区域形状需与数据一致
   ```
3. **类型自动推断**：`write_tsv` 中数字/布尔自动转类型，`=` 开头写公式，前导 0 的串保持文本。
4. **颜色**：COM 内部为 OLE 格式 `0x00BBGGRR`，**统一用** **`rgb(r, g, b)`** **构造**，不要手写整数。
5. **日期**：单元格日期以 Excel 序列值（f64）返回，配合 `number_format("yyyy-mm-dd")` 展示。
6. **连接模型**：`ConnectMode::Attach` 接管已打开实例（不干扰用户）；`ConnectMode::New` 新建实例（用完调 `app.quit()`）。COM 单线程使用；`App` 内部已托管 COM 初始化，无需关心。

## 扩展指南（添加缺失的 COM API）

WPS/Excel 的 COM 对象模型有数百个成员，库只封装了高频部分。补齐缺失功能有三档方式：

### 第 0 档：立即直调（零改动）

任何对象的 `.raw()` 暴露 `get / call / call_named / put / put_raw`：

```rust
// 页面缩放（PageSetup 是方法 → idispatch 拿子对象）
sheet.raw().call("PageSetup", vec![])?.idispatch()?.put("Zoom", vec![false.into()])?;
// 读任意属性
let v = sheet.raw().get("StandardWidth")?.float()?;
```

### 第 1 档：加一行薄封装（推荐给高频 API）

在对应模块加一个单行方法即可，风格与现有代码一致：

```rust
// range.rs 内
pub fn set_outline_font(&self) -> Result<()> {
    self.raw.get_obj("Font", "Font")?.put("Underline", vec![2i32.into()])
}
```

### 第 2 档：新对象类型（模板）

```rust
pub struct PageSetup { raw: IDispatchWrapper }           // 1. 包装 struct

impl PageSetup {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self { Self { raw } }
    pub fn raw(&self) -> &IDispatchWrapper { &self.raw } // 2. 保留直调口
    pub fn zoom(&self, on: bool) -> Result<()> {         // 3. 具体 API
        self.raw.put("Zoom", vec![on.into()])
    }
}

// worksheet.rs 内加工厂：
pub fn page_setup(&self) -> Result<PageSetup> {
    Ok(PageSetup::from_raw(self.raw.get_obj("PageSetup", "PageSetup")?))
}
```

命名约定：setter 一律 `set_xxx`，集合取信息返回 `XxxInfo` 快照结构体，错误用 `ExcelError::hint` 补提示。修改后跑 `cargo clippy --workspace --all-targets` 验证。

## REPL 测试工具

```text
cargo run -p rs-excel-cli
```

**无需预连接**：`list` 会自动接管（attach）已打开的 WPS/Excel 并列出工作簿（只读接管，不干扰用户），随后 `use` 支持 Tab 补全工作簿名/序号实现快速绑定；`connect attach [窗口序号|标题]` 可绑定 `wins` 列出的窗口，重连时自动断开上一个连接；`connect new` 新建实例。提示符始终显示当前引擎（`engine` 切换即时生效）。
**Tab 补全菜单**：输入时下方实时显示候选列表，`Tab`/`Shift+Tab` 循环选择（当前项 `[]` 高亮→菜单内反色高亮），方向键导航，回车确认并**自动补一个空格表示确认**（目录路径除外，便于续接子路径）；`help` 输出与补全同源于一张命令表，永不失同步。典型走查：

```text
wins                     # 列出窗口
connect attach 1         # 接管应用并绑定窗口 1（Tab 可补全窗口序号|标题）
list                     # 列出工作簿
use 1                    # 绑定第 1 个（Tab 补全工作簿名/序号）
sheets                   # 查看工作表
sheet 1                  # 选定
set A1 标题              # 写单元格
tset A1                  # 多行 TSV 批量写（. 结束输入）
tget A1:B4               # 批量读
formula B10 =SUM(B2:B9)
font A1 bold true
merge A1:C1
chart column A1:B9 D2
vba                      # VBA 能力检测
saveas D:\test.xlsx
pdf D:\test.pdf
raw get StandardWidth    # 直调未封装 COM 成员
quit
```

**给 REPL 加新命令**（两步，补全与帮助自动生效）：在 `crates/rs-excel-cli/src/completion.rs` 的 `COMMANDS` 表加一条 `CommandSpec`（用 `ArgSpec::Fixed/Dynamic/Path/Any/Sub` 描述各参数位的补全规则），再到 `commands.rs` 的 `run()` 加一个 match 分支。

## VBA 前置条件

- **`run_macro`**（运行已存在的宏）：无任何设置要求，宏须在打开的工作簿中。
- **`execute`**（注入临时模块执行任意代码）：需在 WPS/Excel 的
  **信任中心 → 宏设置 → 勾选「信任对 VBA 工程对象模型的访问」**。
- WPS 个人版可能未内置 VBA 环境（`Vba::status().env_installed` 为 false 时需安装 VBA 支持库）。
- 策略：简单操作（字体、边框、合并、读写）走内置 API；复杂逻辑（透视、批量重算、事件）走 VBA。

## 依赖

- `windows 0.62`（COM / IDispatch / SafeArray）
- `anyhow`
- CLI：`rustyline`（行编辑/历史）
