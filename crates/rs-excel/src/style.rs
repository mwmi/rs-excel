//! 样式对象：Font / Interior / Borders。均为 Range 的联动子对象，setter 可链式调用。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;

/// 字体（Range.Font）。
pub struct Font {
    raw: IDispatchWrapper,
}

impl Font {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    pub fn bold(&self) -> Result<bool> {
        self.raw.get("Bold")?.bool()
    }

    pub fn set_bold(&self, v: bool) -> Result<()> {
        self.raw.put("Bold", vec![v.into()])
    }

    pub fn italic(&self) -> Result<bool> {
        self.raw.get("Italic")?.bool()
    }

    pub fn set_italic(&self, v: bool) -> Result<()> {
        self.raw.put("Italic", vec![v.into()])
    }

    /// 是否下划线。
    pub fn underline(&self) -> Result<bool> {
        // xlUnderlineStyleNone = -4142
        Ok(self.raw.get("Underline")?.int()? != -4142)
    }

    pub fn set_underline(&self, v: bool) -> Result<()> {
        // xlUnderlineStyleSingle = 2, xlUnderlineStyleNone = -4142
        self.raw.put("Underline", vec![(if v { 2i32 } else { -4142 }).into()])
    }

    pub fn size(&self) -> Result<f64> {
        self.raw.get("Size")?.float()
    }

    pub fn set_size(&self, v: f64) -> Result<()> {
        self.raw.put("Size", vec![v.into()])
    }

    /// 颜色（OLE 格式 0x00BBGGRR）。
    pub fn color(&self) -> Result<u32> {
        Ok(self.raw.get("Color")?.int()? as u32)
    }

    /// 设置颜色（用 `rs_excel::rgb(r, g, b)` 构造）。
    pub fn set_color(&self, ole: u32) -> Result<()> {
        self.raw.put("Color", vec![ole.into()])
    }

    pub fn name(&self) -> Result<String> {
        self.raw.get("Name")?.string()
    }

    pub fn set_name(&self, v: &str) -> Result<()> {
        self.raw.put("Name", vec![v.into()])
    }

    /// 删除线。
    pub fn set_strikethrough(&self, v: bool) -> Result<()> {
        self.raw.put("Strikethrough", vec![v.into()])
    }
}

/// 填充（Range.Interior）。
pub struct Interior {
    raw: IDispatchWrapper,
}

impl Interior {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 填充色（OLE 格式）。
    pub fn color(&self) -> Result<u32> {
        Ok(self.raw.get("Color")?.int()? as u32)
    }

    /// 整块填充状态的三态：`Some(None)` = 整块无填充，`Some(Some(色))` = 整块同一填充色，
    /// `None` = 混合/读不出。
    ///
    /// **判整块必须看 `Pattern`，不能只看 `Color`**：区域里填充不一致时
    /// MS Excel 的 `Interior.Color` 返回 **0**（会被当成黑色填充），WPS 则返回 Null。
    /// `Pattern` 在两家引擎下混合时都是 Null，整块一致时才有值（xlNone = -4142 / 1~18）。
    pub fn fill(&self) -> Option<Option<u32>> {
        match self.raw.get("Pattern").ok().and_then(|v| v.int().ok()) {
            Some(-4142) => Some(None), // xlNone：整块无填充
            Some(p) if p != 0 => self.color().ok().map(Some),
            _ => None, // Null（混合）或 0：取不到统一值，调用方跳过即可
        }
    }

    /// 设置填充色（用 `rs_excel::rgb(r, g, b)` 构造）。
    pub fn set_color(&self, ole: u32) -> Result<()> {
        self.raw.put("Color", vec![ole.into()])
    }

    /// 取消填充。
    pub fn clear(&self) -> Result<()> {
        // xlNone = -4142
        self.raw.put("Pattern", vec![(-4142i32).into()])
    }
}

/// 边框集合（Range.Borders）。
pub struct Borders {
    raw: IDispatchWrapper,
}

/// xlEdge 常量：左/上/下/右/内部竖线/内部横线
const EDGES: [i32; 6] = [7, 8, 9, 10, 11, 12];
/// xlEdge 常量：仅外框
const OUTLINE: [i32; 4] = [7, 8, 9, 10];

/// 一条边的边框状态：`(xlLineStyle, xlBorderWeight, OLE 色)`。
/// 线型 = `xlNone`(-4142) 表示这条边没有边框。
pub type EdgeStyle = (i32, i32, u32);

impl Borders {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 写一条边；`line_style` 为 `xlNone` 时只清线型（粗细/色对无边框没有意义）。
    fn put_edge(&self, edge: i32, line_style: i32, weight: i32, ole_color: u32) -> Result<()> {
        let b = self.raw.call("Item", vec![edge.into()])?.idispatch()?;
        b.put("LineStyle", vec![line_style.into()])?;
        if line_style != -4142 {
            b.put("Weight", vec![weight.into()])?;
            b.put("Color", vec![ole_color.into()])?;
        }
        Ok(())
    }

    fn apply(&self, edges: &[i32], line_style: i32, weight: i32, ole_color: u32) -> Result<()> {
        for edge in edges {
            self.put_edge(*edge, line_style, weight, ole_color)?;
        }
        Ok(())
    }

    /// 六条边各自的边框状态，顺序同 `EDGES`；某条边读不出统一值（区域里不一致）时为 `None`。
    ///
    /// **判据取 `LineStyle` 读不读得出，不能取 `Color`**：混合时 WPS 的 `LineStyle` 返回 Null、
    /// 而 `Color` 给出 0，拿 0 当颜色记下来就会把整块还原成黑框（与 `Interior::fill` 同一坑）。
    pub fn snapshot(&self) -> Vec<Option<EdgeStyle>> {
        EDGES
            .iter()
            .map(|edge| {
                let b = self.raw.call("Item", vec![(*edge).into()]).ok()?.idispatch().ok()?;
                Some((
                    b.get("LineStyle").ok()?.int().ok()?,
                    b.get("Weight").ok()?.int().ok()?,
                    b.get("Color").ok()?.int().ok()? as u32,
                ))
            })
            .collect()
    }

    /// 按 [`snapshot`](Self::snapshot) 的结果逐边还原（`None` 的边跳过）。
    /// 写不进去的边跳过 —— 边框只是外观，不该拖垮整条撤销。
    pub fn restore(&self, states: &[Option<EdgeStyle>]) {
        for (edge, state) in EDGES.iter().zip(states) {
            if let Some((line_style, weight, color)) = state {
                let _ = self.put_edge(*edge, *line_style, *weight, *color);
            }
        }
    }

    /// 全部加边框（四边+内部网格线）。默认参数：`set_all(1, 2, rgb(0,0,0))`
    /// （1 = 连续实线 xlContinuous，2 = 细线 xlThin）。
    pub fn set_all(&self, line_style: i32, weight: i32, ole_color: u32) -> Result<()> {
        self.apply(&EDGES, line_style, weight, ole_color)
    }

    /// 只加外框。
    pub fn set_outline(&self, line_style: i32, weight: i32, ole_color: u32) -> Result<()> {
        self.apply(&OUTLINE, line_style, weight, ole_color)
    }

    /// 清除全部边框。
    pub fn clear(&self) -> Result<()> {
        // xlLineStyleNone = -4142
        self.apply(&EDGES, -4142, 2, 0)
    }
}
