//! 形状：Shapes 集合 / Shape / ShapeRange —— 图形、文本框、图片、连接线的增删改查与批量操作。
//!
//! 单行结构化错误格式见 [`crate::error::ExcelError`]。常用链路：
//!
//! ```no_run
//! use rs_excel::prelude::*;
//!
//! # let sheet: Worksheet = todo!();
//! let shapes = sheet.shapes()?;
//! let s = shapes.add_shape(AutoShapeKind::RoundedRectangle, 100.0, 100.0, 200.0, 80.0)?;
//! s.set_text("点击我")?;
//! s.set_fill_color(rgb(68, 114, 196))?; // OLE 色，rgb() 助手构造
//! for info in shapes.list()? {
//!     println!("{}", info.to_line()); // 单行 Tab 分隔，AI 一次读全景
//! }
//! # Ok::<(), anyhow::Error>(())
//! ```
//!
//! WPS 实测差异（真机探针验证）：
//! - 连接线 `Type` 报 1（autoshape）而非 Excel 的 9（line）；
//! - `AddPicture` 必须 7 参数全给，宽高传 -1 用图片原始尺寸；
//! - `ScaleWidth/ScaleHeight` 返回 E_FAIL，缩放改用宽高乘法实现；
//! - `ZOrderPosition` 未实现，故不封装（`ZOrder` 调整层级可用）；
//! - `Shapes.Range` 数字数组行为异常，实现时先把 key 解析为名称再建范围。

use anyhow::{bail, Result};
use windows::core::BSTR;
use windows::Win32::System::Com::SAFEARRAYBOUND;
use windows::Win32::System::Ole::{SafeArrayCreate, SafeArrayPutElement};
use windows::Win32::System::Variant::{VARIANT, VT_ARRAY, VT_VARIANT};

use crate::com::dispatch::{member_notfound, DispatchResult, IDispatchWrapper};
use crate::com::variant::ComValue;
use crate::error::{ole_to_hex, ExcelError};

/// 枚举样板：生成 `code()`（COM 常量）与 `parse()`（英文名 + 数字代码双解析）。
macro_rules! shape_enum {
    ($(#[$meta:meta])* $name:ident { $($variant:ident = $code:expr => $label:pat,)+ }) => {
        $(#[$meta])*
        #[derive(Clone, Copy, Debug, PartialEq, Eq)]
        pub enum $name { $($variant,)+ }
        impl $name {
            /// COM 常量值。
            pub fn code(self) -> i32 {
                match self { $(Self::$variant => $code,)+ }
            }
            /// 按英文名或数字代码解析（AI 传枚举原始值也能用）。
            pub fn parse(s: &str) -> Option<Self> {
                let t = s.trim();
                if let Ok(n) = t.parse::<i32>() {
                    return match n { $($code => Some(Self::$variant),)+ _ => None };
                }
                Some(match t { $($label => Self::$variant,)+ _ => return None })
            }
        }
    };
}

shape_enum! {
    /// 连接线类型（MsoConnectorType）。
    ConnectorKind {
        Straight = 1 => "straight",
        Elbow = 2 => "elbow",
        Curve = 3 => "curve",
    }
}

shape_enum! {
    /// 层级调整（MsoZOrderCmd）。
    ZOrderCmd {
        BringToFront = 0 => "bring_front",
        SendToBack = 1 => "send_back",
        BringForward = 2 => "bring_forward",
        SendBackward = 3 => "send_backward",
    }
}

shape_enum! {
    /// 对齐方式（MsoAlignCmd）。
    AlignKind {
        Lefts = 0 => "lefts",
        Centers = 1 => "centers",
        Rights = 2 => "rights",
        Tops = 3 => "tops",
        Middles = 4 => "middles",
        Bottoms = 5 => "bottoms",
    }
}

shape_enum! {
    /// 分布方式（MsoDistributeCmd）。
    DistributeKind {
        Horizontally = 0 => "horizontal",
        Vertically = 1 => "vertical",
    }
}

shape_enum! {
    /// 线型（MsoLineDashStyle）。
    DashStyle {
        Solid = 1 => "solid",
        SquareDot = 2 => "square_dot",
        Dot = 3 => "dot",
        Dash = 4 => "dash",
        DashDot = 5 => "dash_dot",
        DashDotDot = 6 => "dash_dot_dot",
        LongDash = 7 => "long_dash",
        LongDashDot = 8 => "long_dash_dot",
    }
}

shape_enum! {
    /// 随单元格变化方式（XlPlacement；WPS 实测仅 1/2/3 合法）。
    Placement {
        MoveSize = 1 => "move_size",
        MoveOnly = 2 => "move_only",
        Free = 3 => "free",
    }
}

shape_enum! {
    /// 翻转方向（MsoFlipCmd：msoFlipHorizontal=0 / msoFlipVertical=1）。
    FlipDir {
        Horizontal = 0 => "horizontal",
        Vertical = 1 => "vertical",
    }
}

/// 自选图形类型（MsoAutoShapeType）。`parse` 接受英文名或数字代码，
/// 未收录的代码经 [`AutoShapeKind::Other`] 直传，全类型可用。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AutoShapeKind {
    Rectangle,
    Diamond,
    RoundedRectangle,
    IsoscelesTriangle,
    RightTriangle,
    Oval,
    Hexagon,
    Cube,
    SmileyFace,
    Donut,
    Heart,
    LightningBolt,
    Sun,
    Moon,
    Cloud,
    RightArrow,
    LeftArrow,
    UpArrow,
    DownArrow,
    Chevron,
    Star5,
    /// 未收录类型直传 MsoAutoShapeType 数字代码
    Other(i32),
}

impl AutoShapeKind {
    /// COM 常量值。
    pub fn code(self) -> i32 {
        match self {
            Self::Rectangle => 1,
            Self::Diamond => 4,
            Self::RoundedRectangle => 5,
            Self::IsoscelesTriangle => 7,
            Self::RightTriangle => 8,
            Self::Oval => 9,
            Self::Hexagon => 10,
            Self::Cube => 14,
            Self::SmileyFace => 17,
            Self::Donut => 18,
            Self::Heart => 21,
            Self::LightningBolt => 22,
            Self::Sun => 23,
            Self::Moon => 25,
            Self::RightArrow => 33,
            Self::LeftArrow => 34,
            Self::UpArrow => 35,
            Self::DownArrow => 36,
            Self::Chevron => 52,
            Self::Star5 => 92,
            Self::Cloud => 179,
            Self::Other(n) => n,
        }
    }

    /// 按英文名或数字代码解析（数字无需收录即可直传）。
    pub fn parse(s: &str) -> Option<Self> {
        let t = s.trim();
        if let Ok(n) = t.parse::<i32>() {
            return Some(Self::Other(n));
        }
        Some(match t {
            "rectangle" => Self::Rectangle,
            "diamond" => Self::Diamond,
            "rounded_rectangle" => Self::RoundedRectangle,
            "isosceles_triangle" | "triangle" => Self::IsoscelesTriangle,
            "right_triangle" => Self::RightTriangle,
            "oval" | "ellipse" => Self::Oval,
            "hexagon" => Self::Hexagon,
            "cube" => Self::Cube,
            "smiley" | "smiley_face" => Self::SmileyFace,
            "donut" | "ring" => Self::Donut,
            "heart" => Self::Heart,
            "lightning" | "lightning_bolt" => Self::LightningBolt,
            "sun" => Self::Sun,
            "moon" => Self::Moon,
            "cloud" => Self::Cloud,
            "right_arrow" | "arrow" => Self::RightArrow,
            "left_arrow" => Self::LeftArrow,
            "up_arrow" => Self::UpArrow,
            "down_arrow" => Self::DownArrow,
            "chevron" => Self::Chevron,
            "star5" | "star_5" | "5point_star" => Self::Star5,
            _ => return None,
        })
    }
}

/// MsoShapeType 数字 → 可读名（未知类型返回 `Type(n)`）。
pub fn type_name(t: i32) -> String {
    let s = match t {
        1 => "autoshape", // WPS 连接线也报 1
        2 => "callout",
        3 => "chart",
        4 => "comment",
        5 => "freeform",
        6 => "group",
        7 => "ole_object",
        8 => "form_control",
        9 => "line",
        10 => "linked_ole_object",
        11 => "linked_picture",
        12 => "activex",
        13 => "picture",
        15 => "placeholder",
        17 => "textbox",
        19 => "table",
        20 => "canvas",
        21 => "diagram",
        22 => "ink",
        24 => "smartart",
        _ => return format!("Type({t})"),
    };
    s.to_string()
}

/// 形状快照：一次 COM 遍历读出的关键字段（[`Shapes::list`] 批量返回）。
#[derive(Debug, Clone)]
pub struct ShapeInfo {
    pub name: String,
    /// 可读类型名（见 [`type_name`]）
    pub kind: String,
    pub visible: bool,
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
    pub rotation: f64,
    /// TopLeftCell 地址（读取失败为空串）
    pub anchor: String,
    /// 形状文本（截断到 20 字符）
    pub text: String,
}

impl ShapeInfo {
    /// 表头，与 [`ShapeInfo::to_line`] 的列一一对应。
    pub fn header() -> &'static str {
        "名称\t类型\t可见\t左\t上\t宽\t高\t旋转\t锚点\t文本"
    }

    /// 单行 Tab 分隔输出（AI 直接可读）。
    pub fn to_line(&self) -> String {
        let text: String = self.text.chars().take(20).collect();
        format!(
            "{}\t{}\t{}\t{:.0}\t{:.0}\t{:.0}\t{:.0}\t{:.0}\t{}\t{}",
            self.name, self.kind, self.visible, self.left, self.top, self.width, self.height,
            self.rotation, self.anchor, text
        )
    }
}

/// 工作表形状集合。
pub struct Shapes {
    raw: IDispatchWrapper,
}

impl Shapes {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 形状数量。
    pub fn count(&self) -> Result<i32> {
        self.raw.get("Count")?.int()
    }

    /// 一次遍历读出全部形状关键字段（AI 用一次调用拿到全景，最省 token）。
    pub fn list(&self) -> Result<Vec<ShapeInfo>> {
        let n = self.count()?;
        (1..=n)
            .map(|i| {
                let s = self.raw.call("Item", vec![i.into()])?.idispatch_child(&i.to_string())?;
                Shape::from_raw(s).info()
            })
            .collect()
    }

    /// 名称或 1 基序号取形状（纯数字按序号，其余按名称）。
    pub fn get(&self, key: &str) -> Result<Shape> {
        let arg: ComValue = match key.trim().parse::<i32>() {
            Ok(i) if i >= 1 => i.into(),
            _ => key.into(),
        };
        self.raw
            .call("Item", vec![arg])?
            .idispatch_child(key)
            .map(Shape::from_raw)
            .map_err(|e| ExcelError::hint(e, "形状不存在：先用 shapes().list() 查看现有形状名称/序号"))
    }

    /// 添加自选图形。
    pub fn add_shape(&self, kind: AutoShapeKind, left: f64, top: f64, width: f64, height: f64) -> Result<Shape> {
        self.wrap_shape(self.raw.call(
            "AddShape",
            vec![kind.code().into(), left.into(), top.into(), width.into(), height.into()],
        ))
    }

    /// 添加文本框（水平方向）。
    pub fn add_textbox(&self, left: f64, top: f64, width: f64, height: f64) -> Result<Shape> {
        self.wrap_shape(self.raw.call(
            "AddTextbox",
            vec![1i32.into(), left.into(), top.into(), width.into(), height.into()],
        ))
    }

    /// 添加直线。
    pub fn add_line(&self, x1: f64, y1: f64, x2: f64, y2: f64) -> Result<Shape> {
        self.wrap_shape(self.raw.call("AddLine", vec![x1.into(), y1.into(), x2.into(), y2.into()]))
    }

    /// 添加连接线。
    pub fn add_connector(&self, kind: ConnectorKind, x1: f64, y1: f64, x2: f64, y2: f64) -> Result<Shape> {
        self.wrap_shape(self.raw.call(
            "AddConnector",
            vec![kind.code().into(), x1.into(), y1.into(), x2.into(), y2.into()],
        ))
    }

    /// 添加图片。`width`/`height` 传 `None` 用图片原始尺寸（-1）。
    pub fn add_picture(&self, path: &str, left: f64, top: f64, width: Option<f64>, height: Option<f64>) -> Result<Shape> {
        // WPS 要求 7 参数全给：AddPicture(FileName, LinkToFile=msoFalse, SaveWithDocument=msoTrue, Left, Top, Width, Height)
        self.raw
            .call(
                "AddPicture",
                vec![
                    path.into(),
                    false.into(),
                    true.into(),
                    left.into(),
                    top.into(),
                    width.unwrap_or(-1.0).into(),
                    height.unwrap_or(-1.0).into(),
                ],
            )
            .and_then(|v| v.idispatch_child("Picture").map(Shape::from_raw))
            .map_err(|e| ExcelError::hint(e, "路径需为绝对路径且文件存在"))
    }

    /// 名称/序号混选多个形状为批量范围。
    pub fn range(&self, keys: &[&str]) -> Result<ShapeRange> {
        if keys.is_empty() {
            bail!("range：至少需要 1 个形状名称或序号");
        }
        // 先统一解析为名称再建范围（WPS 的 Range 数字数组行为异常）
        let names: Vec<String> = keys.iter().map(|k| self.get(k)?.name()).collect::<Result<_>>()?;
        self.raw
            .call("Range", vec![names_variant(&names)?])?
            .idispatch_child("ShapeRange")
            .map(ShapeRange::from_raw)
    }

    /// 删除全部形状（倒序循环，避免索引错乱）。
    pub fn delete_all(&self) -> Result<()> {
        for i in (1..=self.count()?).rev() {
            self.raw.call("Item", vec![i.into()])?.idispatch()?.call("Delete", vec![])?;
        }
        Ok(())
    }

    /// 全选形状。
    pub fn select_all(&self) -> Result<()> {
        self.raw.call("SelectAll", vec![]).map(|_| ())
    }

    fn wrap_shape(&self, r: Result<DispatchResult>) -> Result<Shape> {
        r.and_then(|v| v.idispatch_child("Shape").map(Shape::from_raw))
    }
}

/// 字符串数组 → 1 维 `VT_ARRAY|VT_VARIANT`（Shapes.Range 参数）。
fn names_variant(names: &[String]) -> Result<ComValue> {
    let bounds = [SAFEARRAYBOUND { cElements: names.len() as u32, lLbound: 0 }];
    let psa = unsafe { SafeArrayCreate(VT_VARIANT, 1, bounds.as_ptr()) };
    if psa.is_null() {
        bail!("SafeArrayCreate 失败");
    }
    for (i, n) in names.iter().enumerate() {
        let v = VARIANT::from(BSTR::from(n.as_str()));
        let idx = [i as i32];
        unsafe { SafeArrayPutElement(psa, idx.as_ptr(), &v as *const VARIANT as *mut _)? };
        drop(v); // PutElement 内部已深拷贝，清掉本地副本
    }
    let mut var = VARIANT::default();
    unsafe {
        (*var.Anonymous.Anonymous).vt = VT_ARRAY | VT_VARIANT;
        (*var.Anonymous.Anonymous).Anonymous.parray = psa;
    }
    Ok(ComValue::Raw(var))
}

/// 单个形状。
pub struct Shape {
    raw: IDispatchWrapper,
}

impl Shape {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    // ── 标识 ──

    pub fn name(&self) -> Result<String> {
        self.raw.get("Name")?.string()
    }

    pub fn set_name(&self, name: &str) -> Result<()> {
        self.raw
            .put("Name", vec![name.into()])
            .map_err(|e| ExcelError::hint(e, "名称不能与同表其他形状重复"))
    }

    /// 类型名（如 autoshape/picture/textbox，见 [`type_name`]）。
    pub fn type_name(&self) -> Result<String> {
        Ok(type_name(self.raw.get("Type")?.int()?))
    }

    /// 关键字段快照（个别字段读取失败填默认值，尽量返回可用数据）。
    pub fn info(&self) -> Result<ShapeInfo> {
        let read_f = |name: &str| self.raw.get(name).and_then(|v| v.float()).unwrap_or(0.0);
        Ok(ShapeInfo {
            name: self.name().unwrap_or_default(),
            kind: self.type_name().unwrap_or_default(),
            visible: self.visible().unwrap_or(true),
            left: read_f("Left"),
            top: read_f("Top"),
            width: read_f("Width"),
            height: read_f("Height"),
            rotation: read_f("Rotation"),
            anchor: self.top_left_cell().unwrap_or_default(),
            text: self.text().unwrap_or_default(),
        })
    }

    // ── 位置大小 ──

    /// 一次设置多项几何，`None` 项跳过。
    pub fn set_geometry(&self, left: Option<f64>, top: Option<f64>, width: Option<f64>, height: Option<f64>) -> Result<()> {
        if let Some(v) = left {
            self.raw.put("Left", vec![v.into()])?;
        }
        if let Some(v) = top {
            self.raw.put("Top", vec![v.into()])?;
        }
        if let Some(v) = width {
            self.raw.put("Width", vec![v.into()])?;
        }
        if let Some(v) = height {
            self.raw.put("Height", vec![v.into()])?;
        }
        Ok(())
    }

    pub fn rotation(&self) -> Result<f64> {
        self.raw.get("Rotation")?.float()
    }

    pub fn set_rotation(&self, deg: f64) -> Result<()> {
        self.raw.put("Rotation", vec![deg.into()]).map(|_| ())
    }

    /// 相对当前位置移动（磅）。
    pub fn move_by(&self, dx: f64, dy: f64) -> Result<()> {
        self.raw.call("IncrementLeft", vec![dx.into()])?;
        self.raw.call("IncrementTop", vec![dy.into()]).map(|_| ())
    }

    /// 增量旋转（度）。
    pub fn rotate_by(&self, deg: f64) -> Result<()> {
        self.raw.call("IncrementRotation", vec![deg.into()]).map(|_| ())
    }

    /// 按比例缩放（1.0 = 不变；WPS 未实现 ScaleWidth，用宽高乘法替代）。
    pub fn scale(&self, w: f64, h: f64) -> Result<()> {
        let width = self.raw.get("Width")?.float()?;
        let height = self.raw.get("Height")?.float()?;
        self.raw.put("Width", vec![(width * w).into()])?;
        self.raw.put("Height", vec![(height * h).into()]).map(|_| ())
    }

    /// 翻转。
    pub fn flip(&self, dir: FlipDir) -> Result<()> {
        self.raw.call("Flip", vec![dir.code().into()]).map(|_| ())
    }

    // ── 可见与锚定 ──

    pub fn visible(&self) -> Result<bool> {
        self.raw.get("Visible")?.bool()
    }

    pub fn set_visible(&self, on: bool) -> Result<()> {
        put_visible(&self.raw, on)
    }

    /// 设置随单元格变化方式。
    pub fn set_placement(&self, placement: Placement) -> Result<()> {
        self.raw.put("Placement", vec![placement.code().into()]).map(|_| ())
    }

    /// 左上角所在单元格地址。
    pub fn top_left_cell(&self) -> Result<String> {
        self.raw.get("TopLeftCell")?.idispatch()?.get("Address")?.string()
    }

    /// 右下角所在单元格地址。
    pub fn bottom_right_cell(&self) -> Result<String> {
        self.raw.get("BottomRightCell")?.idispatch()?.get("Address")?.string()
    }

    // ── 层级 ──

    /// 调整层级（置于顶层/底层/上移一层/下移一层）。
    pub fn z_order(&self, cmd: ZOrderCmd) -> Result<()> {
        self.raw.call("ZOrder", vec![cmd.code().into()]).map(|_| ())
    }

    // ── 填充与线条 ──

    /// 填充色 `#RRGGBB`。
    pub fn fill_color(&self) -> Result<String> {
        let c = self.raw.get_obj("Fill", "Fill")?.get_obj("ForeColor", "ForeColor")?.get("RGB")?.float()?;
        Ok(ole_to_hex(c as u32))
    }

    /// 设置填充色（同时确保填充可见）。
    pub fn set_fill_color(&self, ole: u32) -> Result<()> {
        set_fill(&self.raw, ole)
    }

    /// 取消填充。
    pub fn no_fill(&self) -> Result<()> {
        self.raw.get_obj("Fill", "Fill")?.put("Visible", vec![(0i32).into()]).map(|_| ())
    }

    /// 线条色 `#RRGGBB`。
    pub fn line_color(&self) -> Result<String> {
        let c = self.raw.get_obj("Line", "Line")?.get_obj("ForeColor", "ForeColor")?.get("RGB")?.float()?;
        Ok(ole_to_hex(c as u32))
    }

    /// 设置线条色（同时确保线条可见）。
    pub fn set_line_color(&self, ole: u32) -> Result<()> {
        set_line(&self.raw, ole)
    }

    /// 线宽（磅）。
    pub fn set_line_weight(&self, pt: f64) -> Result<()> {
        self.raw.get_obj("Line", "Line")?.put("Weight", vec![pt.into()]).map(|_| ())
    }

    /// 线型。
    pub fn set_line_dash(&self, style: DashStyle) -> Result<()> {
        self.raw.get_obj("Line", "Line")?.put("DashStyle", vec![style.code().into()]).map(|_| ())
    }

    /// 取消线条。
    pub fn no_line(&self) -> Result<()> {
        self.raw.get_obj("Line", "Line")?.put("Visible", vec![(0i32).into()]).map(|_| ())
    }

    // ── 文本（TextFrame2 优先，WPS 回退 TextFrame）──

    /// 形状文本（无文本框的形状返回空串）。
    pub fn text(&self) -> Result<String> {
        if let Ok(t) = self
            .raw
            .get_obj("TextFrame2", "TextFrame2")
            .and_then(|f| f.get_obj("TextRange", "TextRange"))
            .and_then(|r| r.get("Text"))
            .and_then(|v| v.string())
        {
            return Ok(t);
        }
        match self
            .raw
            .get_obj("TextFrame", "TextFrame")?
            .call("Characters", vec![])?
            .idispatch()?
            .get("Text")?
            .string()
        {
            Ok(t) => Ok(t),
            Err(e) if member_notfound(&e) => Ok(String::new()),
            Err(e) => Err(e),
        }
    }

    /// 写形状文本。
    pub fn set_text(&self, text: &str) -> Result<()> {
        if let Ok(f) = self.raw.get_obj("TextFrame2", "TextFrame2") {
            if let Ok(r) = f.get_obj("TextRange", "TextRange") {
                if r.put("Text", vec![text.into()]).is_ok() {
                    return Ok(());
                }
            }
        }
        self.raw
            .get_obj("TextFrame", "TextFrame")?
            .call("Characters", vec![])?
            .idispatch()?
            .put("Text", vec![text.into()])
            .map_err(|e| ExcelError::hint(e, "该形状可能不支持文本（仅文本框/自选图形/标注等可含文本）"))
    }

    /// 一次设置文本字体多项，`None` 项跳过。
    pub fn set_text_style(&self, size: Option<f64>, bold: Option<bool>, color_ole: Option<u32>, name: Option<&str>) -> Result<()> {
        match self
            .raw
            .get_obj("TextFrame2", "TextFrame2")
            .and_then(|f| f.get_obj("TextRange", "TextRange"))
            .and_then(|r| r.get_obj("Font", "Font"))
        {
            Ok(font) => {
                put_font(&font, size, bold, name)?;
                if let Some(v) = color_ole {
                    font.get_obj("Fill", "Fill")?.get_obj("ForeColor", "ForeColor")?.put("RGB", vec![v.into()])?;
                }
                Ok(())
            }
            Err(_) => {
                let font = self
                    .raw
                    .get_obj("TextFrame", "TextFrame")?
                    .call("Characters", vec![])?
                    .idispatch()?
                    .get_obj("Font", "Font")?;
                put_font(&font, size, bold, name)?;
                if let Some(v) = color_ole {
                    font.put("Color", vec![v.into()])?;
                }
                Ok(())
            }
        }
    }

    // ── 生命周期 ──

    pub fn delete(&self) -> Result<()> {
        self.raw.call("Delete", vec![]).map(|_| ())
    }

    /// 创建副本并返回。
    pub fn duplicate(&self) -> Result<Shape> {
        self.raw
            .call("Duplicate", vec![])?
            .idispatch_child("Duplicate")
            .map(Shape::from_raw)
    }

    /// 复制到剪贴板。
    pub fn copy(&self) -> Result<()> {
        self.raw.call("Copy", vec![]).map(|_| ())
    }

    /// 剪切到剪贴板。
    pub fn cut(&self) -> Result<()> {
        self.raw.call("Cut", vec![]).map(|_| ())
    }

    /// 取消组合。
    pub fn ungroup(&self) -> Result<()> {
        self.raw.call("Ungroup", vec![]).map(|_| ())
    }

    /// 组合形状的子形状快照。
    pub fn group_items(&self) -> Result<Vec<ShapeInfo>> {
        let gi = self.raw.get_obj("GroupItems", "GroupItems")?;
        let n = gi.get("Count")?.int()?;
        (1..=n)
            .map(|i| {
                let s = gi.call("Item", vec![i.into()])?.idispatch_child(&i.to_string())?;
                Shape::from_raw(s).info()
            })
            .collect()
    }

    // ── 交互 ──

    /// 点击时运行的宏名（WPS 读回带工作簿限定，如 `工作簿1!MyMacro`）。
    pub fn on_action(&self) -> Result<String> {
        self.raw.get("OnAction")?.string()
    }

    /// 设置点击时运行的宏。
    pub fn set_on_action(&self, macro_name: &str) -> Result<()> {
        self.raw
            .put("OnAction", vec![macro_name.into()])
            .map_err(|e| ExcelError::hint(e, "宏必须在标准模块中，且不能带参数"))
    }

    /// 设置超链接（已有链接改 Address，否则经 Hyperlinks.Add 挂载）。
    pub fn set_hyperlink(&self, addr: &str) -> Result<()> {
        match self.raw.get("Hyperlink") {
            Ok(h) => h
                .idispatch()?
                .put("Address", vec![addr.into()])
                .map_err(|e| ExcelError::hint(e, "地址需为合法 URL 或文件路径")),
            Err(_) => self
                .raw
                .get_obj("Parent", "Parent")?
                .get("Hyperlinks")?
                .idispatch()?
                .call_named(
                    "Add",
                    vec![],
                    &[("Anchor", ComValue::Obj(self.raw.disp.clone())), ("Address", addr.into())],
                )
                .map(|_| ())
                .map_err(|e| ExcelError::hint(e, "地址需为合法 URL 或文件路径")),
        }
    }

    /// 替代文本（无障碍说明）。
    pub fn alt_text(&self) -> Result<String> {
        self.raw.get("AlternativeText")?.string()
    }

    pub fn set_alt_text(&self, text: &str) -> Result<()> {
        self.raw.put("AlternativeText", vec![text.into()]).map(|_| ())
    }

    // ── 微调 ──

    /// 调整值个数（随形状类型而定，如箭头有 2 个）。
    pub fn adjustment_count(&self) -> Result<i32> {
        self.raw.get_obj("Adjustments", "Adjustments")?.get("Count")?.int()
    }

    /// 设置第 `idx` 个调整值（1 基，取值约 0~1，随形状而定）。
    pub fn set_adjustment(&self, idx: i32, value: f64) -> Result<()> {
        if idx < 1 {
            bail!("调整值序号从 1 开始（得到 {idx}）");
        }
        self.raw
            .get_obj("Adjustments", "Adjustments")?
            .put("Item", vec![idx.into(), value.into()])
            .map_err(|e| ExcelError::hint(e, "序号须 ≤ adjustment_count()，且该形状需有调整值"))
    }

    /// 阴影开关（WPS 读回可能不生效，以界面为准）。
    pub fn set_shadow(&self, on: bool) -> Result<()> {
        self.raw
            .get_obj("Shadow", "Shadow")?
            .put("Visible", vec![(if on { -1i32 } else { 0 }).into()])
            .map(|_| ())
    }
}

/// 多个形状的批量范围（`Shapes::range` 创建）。
pub struct ShapeRange {
    raw: IDispatchWrapper,
}

impl ShapeRange {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 成员数量。
    pub fn count(&self) -> Result<i32> {
        self.raw.get("Count")?.int()
    }

    /// 对齐。
    pub fn align(&self, kind: AlignKind, relative_to: bool) -> Result<()> {
        self.raw.call("Align", vec![kind.code().into(), relative_to.into()]).map(|_| ())
    }

    /// 分布。
    pub fn distribute(&self, kind: DistributeKind) -> Result<()> {
        self.raw.call("Distribute", vec![kind.code().into()]).map(|_| ())
    }

    /// 组合为一个形状（至少 2 个成员）。
    pub fn group(&self) -> Result<Shape> {
        self.raw
            .call("Group", vec![])?
            .idispatch_child("Group")
            .map(Shape::from_raw)
            .map_err(|e| ExcelError::hint(e, "组合至少需要 2 个形状"))
    }

    /// 取消组合。
    pub fn ungroup(&self) -> Result<()> {
        self.raw.call("Ungroup", vec![]).map(|_| ())
    }

    /// 批量删除。
    pub fn delete(&self) -> Result<()> {
        self.raw.call("Delete", vec![]).map(|_| ())
    }

    /// 批量显示/隐藏。
    pub fn set_visible(&self, on: bool) -> Result<()> {
        put_visible(&self.raw, on)
    }

    /// 批量设置填充色（同时确保填充可见）。
    pub fn set_fill_color(&self, ole: u32) -> Result<()> {
        set_fill(&self.raw, ole)
    }

    /// 批量设置线条色（同时确保线条可见）。
    pub fn set_line_color(&self, ole: u32) -> Result<()> {
        set_line(&self.raw, ole)
    }
}

// ── Shape / ShapeRange 共用小助手 ──

fn put_visible(obj: &IDispatchWrapper, on: bool) -> Result<()> {
    obj.put("Visible", vec![(if on { -1i32 } else { 0 }).into()]).map(|_| ())
}

fn set_fill(obj: &IDispatchWrapper, ole: u32) -> Result<()> {
    let f = obj.get_obj("Fill", "Fill")?;
    f.put("Visible", vec![(-1i32).into()])?;
    f.get_obj("ForeColor", "ForeColor")?.put("RGB", vec![ole.into()]).map(|_| ())
}

fn set_line(obj: &IDispatchWrapper, ole: u32) -> Result<()> {
    let l = obj.get_obj("Line", "Line")?;
    l.put("Visible", vec![(-1i32).into()])?;
    l.get_obj("ForeColor", "ForeColor")?.put("RGB", vec![ole.into()]).map(|_| ())
}

fn put_font(font: &IDispatchWrapper, size: Option<f64>, bold: Option<bool>, name: Option<&str>) -> Result<()> {
    if let Some(v) = size {
        font.put("Size", vec![v.into()])?;
    }
    if let Some(v) = bold {
        font.put("Bold", vec![(if v { -1i32 } else { 0 }).into()])?;
    }
    if let Some(v) = name {
        font.put("Name", vec![v.into()])?;
    }
    Ok(())
}
