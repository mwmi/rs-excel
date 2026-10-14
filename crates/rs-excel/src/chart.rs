//! 图表：创建、类型、标题、导出图片、删除。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::ExcelError;
use crate::range::Range;

/// 图表类型（XlChartType 常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChartKind {
    Area = 1,
    Column = 51,
    Bar = 57,
    Line = 4,
    Pie = 5,
    Scatter = -4169,
    Doughnut = -4120,
}

impl ChartKind {
    /// 按名称解析（REPL 用）：column/bar/line/pie/area/scatter/doughnut。
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_ascii_lowercase().as_str() {
            "column" => Some(Self::Column),
            "bar" => Some(Self::Bar),
            "line" => Some(Self::Line),
            "pie" => Some(Self::Pie),
            "area" => Some(Self::Area),
            "scatter" => Some(Self::Scatter),
            "doughnut" => Some(Self::Doughnut),
            _ => None,
        }
    }
}

/// 图表（Chart COM 对象）。
pub struct Chart {
    raw: IDispatchWrapper,
}

impl Chart {
    pub(crate) fn from_raw(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 设置标题。
    pub fn set_title(&self, title: &str) -> Result<()> {
        self.raw.put("HasTitle", vec![true.into()])?;
        self.raw
            .get_obj("ChartTitle", "ChartTitle")?
            .put("Text", vec![title.into()])
    }

    /// 更改图表类型。
    pub fn set_type(&self, kind: ChartKind) -> Result<()> {
        self.raw.put("ChartType", vec![(kind as i32).into()])
    }

    /// 导出为图片（按扩展名决定格式：.png/.jpg）。
    pub fn export_image(&self, path: &str) -> Result<()> {
        self.raw
            .call("Export", vec![path.into()])
            .map(|_| ())
            .map_err(|e| ExcelError::hint(e, "路径需为绝对路径；按扩展名决定图片格式"))
    }

    /// 删除图表。
    pub fn delete(&self) -> Result<()> {
        self.raw.get_obj("Parent", "Parent")?.call("Delete", vec![]).map(|_| ())
    }

    /// 把每个序列的**分类轴（X 轴标签）**指向 `cat`，`cat` 只给数据行那些格（如 `B2:B11`）。
    ///
    /// 为什么需要它：`SetSourceData` 只收**一整块矩形**，而「姓名 + 总分」这种分类列与数值列
    /// **不相邻**的图它表达不了 —— 两块拼成一个 union 传进去会被当成**两个序列**（姓名那列
    /// 自己也变成一条折线），列再多就是一串序列。只能建完图后逐条设 `XValues`。
    pub fn set_categories(&self, cat: &Range) -> Result<()> {
        // 先把取向定死成「每列一条序列」：单列数据默认按「每行一条」解，一列 10 个数会变成
        // 10 条各含 1 个点的序列（实测：SERIES 里 values 是 $I$2、$I$3… 一格一条，图是一片平线）
        self.raw.put("PlotBy", vec![(2i32).into()])?; // xlColumns
        let series = self.raw.get_obj("SeriesCollection", "SeriesCollection")?;
        let count = series.get("Count")?.int()?;
        for index in 1..=count {
            series
                .call("Item", vec![index.into()])?
                .idispatch()?
                .put("XValues", vec![ComValue::Obj(cat.raw().disp.clone())])?;
        }
        Ok(())
    }
}
