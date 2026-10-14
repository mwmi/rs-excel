//! 2D SAFEARRAY ↔ 网格数据互转（Range 批量读写的底层）。
//!
//! Excel/WPS 的 `Range.Value`：
//! - 多格区域返回 `VT_ARRAY|VT_VARIANT` 的 2D 数组（1 基，维度1=行、维度2=列）；
//! - 单格区域返回标量 VARIANT。
//!
//! SafeArrayGetElement/PutElement 的索引向量与维度编号一一对应：
//! `rgindices[0]` = 第 1 维（行）、`rgindices[1]` = 第 2 维（列），即 `[行, 列]`。
//! （实测 WPS：`[1,1]` 成功而 `[1,2]` 在 10×1 数组上越界，行在前才自洽。）

use anyhow::{anyhow, Result};
use windows::Win32::System::Com::{SAFEARRAY, SAFEARRAYBOUND};
use windows::Win32::System::Ole::{
    SafeArrayCreate, SafeArrayGetDim, SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound,
    SafeArrayPutElement,
};
use windows::Win32::System::Variant::{VARIANT, VT_ARRAY, VT_VARIANT};

use crate::com::variant::{variant_to_cellvalue, variant_to_cellvalue_typed, CellValue, ComValue};

/// 一维 `ComValue` 列表 → `VT_ARRAY|VT_VARIANT` 的 1 维 SAFEARRAY VARIANT（0 基）。
/// 供 AutoFilter 多选 Criteria1、透视表 Group 的 Periods、Subtotal 的 TotalList、
/// AddFields 字段名数组、VisibleSlicerItemsList 等数组参数复用。
pub fn list_to_variant(items: Vec<ComValue>) -> Result<VARIANT> {
    if items.is_empty() {
        return Err(anyhow!("数组参数为空"));
    }
    let n = items.len() as u32;
    let bounds = [SAFEARRAYBOUND {
        cElements: n,
        lLbound: 0,
    }];
    let psa = unsafe { SafeArrayCreate(VT_VARIANT, 1, bounds.as_ptr()) };
    if psa.is_null() {
        return Err(anyhow!("SafeArrayCreate 失败"));
    }
    for (i, item) in items.into_iter().enumerate() {
        let v = VARIANT::from(item);
        let idx = [i as i32];
        unsafe {
            SafeArrayPutElement(psa, idx.as_ptr(), &v as *const VARIANT as *mut _)?;
        }
        drop(v); // PutElement 内部已深拷贝，清掉本地副本
    }
    let mut var = VARIANT::default();
    unsafe {
        (*var.Anonymous.Anonymous).vt = VT_ARRAY | VT_VARIANT;
        (*var.Anonymous.Anonymous).Anonymous.parray = psa;
    }
    Ok(var)
}

/// 读取 `Range.Value` 返回的 VARIANT：多格 → 2D 网格；单格 → 1x1 网格。
pub fn variant_to_rows(v: &VARIANT) -> Result<Vec<Vec<CellValue>>> {
    variant_to_rows_with(v, variant_to_cellvalue)
}

/// 同 [variant_to_rows]，但日期元素转为可读文本（`Value` 兼容通道用）。
pub fn variant_to_rows_typed(v: &VARIANT) -> Result<Vec<Vec<CellValue>>> {
    variant_to_rows_with(v, variant_to_cellvalue_typed)
}

fn variant_to_rows_with(v: &VARIANT, conv: fn(&VARIANT) -> CellValue) -> Result<Vec<Vec<CellValue>>> {
    let vt: u16 = unsafe { v.Anonymous.Anonymous.vt }.0;
    if vt & VT_ARRAY.0 == 0 {
        return Ok(vec![vec![conv(v)]]);
    }
    let psa: *mut SAFEARRAY = unsafe { v.Anonymous.Anonymous.Anonymous.parray };
    if psa.is_null() {
        return Err(anyhow!("SAFEARRAY 指针为空"));
    }
    unsafe {
        let dims = SafeArrayGetDim(psa);
        if dims != 2 {
            return Err(anyhow!("期望 2 维数组，实际 {dims} 维"));
        }
        let r0 = SafeArrayGetLBound(psa, 1)?;
        let r1 = SafeArrayGetUBound(psa, 1)?;
        let c0 = SafeArrayGetLBound(psa, 2)?;
        let c1 = SafeArrayGetUBound(psa, 2)?;
        let mut rows = Vec::with_capacity((r1 - r0 + 1).max(0) as usize);
        for r in r0..=r1 {
            let mut row = Vec::with_capacity((c1 - c0 + 1).max(0) as usize);
            for c in c0..=c1 {
                let idx = [r, c]; // 索引向量与维度编号对应：[行, 列]
                let mut elem = VARIANT::default();
                SafeArrayGetElement(psa, idx.as_ptr(), &mut elem as *mut VARIANT as *mut _)?;
                row.push(conv(&elem));
                // elem 离开作用域自动 VariantClear，无泄漏
            }
            rows.push(row);
        }
        Ok(rows)
    }
}

/// 网格 → `VT_ARRAY|VT_VARIANT` 的 2D SAFEARRAY VARIANT（0 基）。
/// 返回的 VARIANT 拥有数组，Drop 时自动 SafeArrayDestroy。
pub fn rows_to_variant(grid: &[Vec<CellValue>]) -> Result<VARIANT> {
    let nrows = grid.len() as u32;
    let ncols = grid.first().map_or(0, |r| r.len()) as u32;
    if nrows == 0 || ncols == 0 {
        return Err(anyhow!("数据为空"));
    }
    if grid.iter().any(|r| r.len() as u32 != ncols) {
        return Err(anyhow!("各行长度不一致"));
    }
    // rgsabound[0] 是最右维（列），rgsabound[1] 是行
    let bounds = [
        SAFEARRAYBOUND {
            cElements: ncols,
            lLbound: 0,
        },
        SAFEARRAYBOUND {
            cElements: nrows,
            lLbound: 0,
        },
    ];
    let psa = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
    if psa.is_null() {
        return Err(anyhow!("SafeArrayCreate 失败"));
    }
    for (r, row) in grid.iter().enumerate() {
        for (c, val) in row.iter().enumerate() {
            let v = VARIANT::from(ComValue::from(val.clone()));
            let idx = [r as i32, c as i32]; // 索引向量与维度编号对应：[行, 列]
            unsafe {
                SafeArrayPutElement(psa, idx.as_ptr(), &v as *const VARIANT as *mut _)?;
            }
            drop(v); // PutElement 内部已深拷贝，清掉本地副本
        }
    }
    let mut var = VARIANT::default();
    unsafe {
        (*var.Anonymous.Anonymous).vt = VT_ARRAY | VT_VARIANT;
        (*var.Anonymous.Anonymous).Anonymous.parray = psa;
    }
    Ok(var)
}
