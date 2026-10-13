//! 样式对象：Font / Interior / Borders。均为 Range 的联动子对象，setter 可链式调用。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::error::ExcelError;

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

impl Borders {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    fn apply(&self, edges: &[i32], line_style: i32, weight: i32, ole_color: u32) -> Result<()> {
        for edge in edges {
            let b = self.raw.call("Item", vec![(*edge).into()])?.idispatch()?;
            b.put("LineStyle", vec![line_style.into()])?;
            b.put("Weight", vec![weight.into()])?;
            b.put("Color", vec![ole_color.into()])?;
        }
        Ok(())
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
        for edge in EDGES {
            let b = self.raw.call("Item", vec![edge.into()])?.idispatch()?;
            // xlLineStyleNone = -4142
            b.put("LineStyle", vec![(-4142i32).into()])
                .map_err(|e| ExcelError::hint(e, "清除边框失败"))?;
        }
        Ok(())
    }
}
