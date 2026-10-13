//! IDispatch 后期绑定薄封装 —— 整个库的 COM 底座。
//!
//! 设计要点：
//! - 所有对象包装（App/Workbook/Worksheet/Range/...）内部都持有一个 `IDispatchWrapper`；
//! - `get/call/put` 自动在 PROPERTYGET 与 METHOD 之间回退，兼容 WPS 把属性暴露成方法等怪癖；
//! - 每个包装自带对象路径 `ctx`，错误信息自动带上完整定位（见 [crate::error::ExcelError]）；
//! - **扩展性核心**：库未封装的 COM 成员，用 `.raw().get/call/call_named/put` 一行直调。

use anyhow::Result;
use windows::core::{BSTR, GUID, PCWSTR};
use windows::Win32::Globalization::LOCALE_USER_DEFAULT;
use windows::Win32::System::Com::{
    DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS,
    IDispatch,
};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARIANT, VT_DISPATCH};

use crate::com::variant::{variant_to_cellvalue, CellValue, ComValue};
use crate::error::ExcelError;

const DISP_E_MEMBERNOTFOUND: u32 = 0x8002_0003;

/// 常见 HRESULT → 修复提示。
fn common_hint(hr: u32) -> &'static str {
    match hr {
        0x8002_0003 => "成员不存在：检查成员名拼写（可在 REPL 用 raw 命令探测）",
        0x8002_0005 => "类型不匹配：检查参数类型",
        0x8002_0009 => "COM 调用抛出异常：常见于操作被应用弹窗取消（如另存覆盖确认点了取消）、序号越界或参数不被接受",
        0x800A_03EC => "应用层错误：常见于非法单元格地址、区域越界、对象不存在",
        0x800A_01A8 => "对象下标越界：检查集合索引或名称",
        _ => "",
    }
}

/// COM 错误 → 带上下文的 ExcelError。
pub(crate) fn com_err(
    kind: &str,
    member: &str,
    ctx: &str,
    e: windows::core::Error,
) -> ExcelError {
    let hr = e.code().0 as u32;
    ExcelError::new(
        format!("{kind}('{member}')"),
        ctx,
        format!("0x{hr:08X} {e}"),
    )
    .with_hint(common_hint(hr))
}

/// 错误是否为 DISP_E_MEMBERNOTFOUND（成员不存在，用于 Formula2→Formula 等回退）。
pub(crate) fn member_notfound(e: &anyhow::Error) -> bool {
    e.to_string().contains("0x80020003")
}

/// 包住一个 `IDispatch` 指针 + 对象路径。
#[derive(Clone)]
pub struct IDispatchWrapper {
    pub disp: IDispatch,
    /// 对象路径，如 `WPS > Book1.xlsx > Sheet1 > A1:C3`，错误定位用。
    pub ctx: String,
}

impl IDispatchWrapper {
    /// 根对象（Application）。`engine_label` 如 "WPS"。
    pub fn root(disp: IDispatch, engine_label: &str) -> Self {
        Self {
            disp,
            ctx: engine_label.to_string(),
        }
    }

    /// 子对象：继承父路径并追加自身标签。
    pub fn child(&self, disp: IDispatch, label: &str) -> Self {
        Self {
            disp,
            ctx: format!("{} > {}", self.ctx, label),
        }
    }

    /// 同一对象、更深一层的上下文（如 Range → Font）。
    pub fn deeper(&self, label: &str) -> Self {
        Self {
            disp: self.disp.clone(),
            ctx: format!("{} > {}", self.ctx, label),
        }
    }

    /// get 属性并转为子对象（ctx 自动追加 label）。
    pub fn get_obj(&self, name: &str, label: &str) -> Result<IDispatchWrapper> {
        self.get(name)?.idispatch_child(label)
    }

    /// call 方法并转为子对象（ctx 自动追加 label）。
    pub fn call_obj(
        &self,
        name: &str,
        label: &str,
        args: Vec<ComValue>,
    ) -> Result<IDispatchWrapper> {
        self.call(name, args)?.idispatch_child(label)
    }

    fn get_dispid(&self, name: &str) -> windows::core::Result<i32> {
        let bstr = BSTR::from(name);
        let names = [PCWSTR(bstr.as_ptr())];
        let mut dispid: i32 = 0;
        unsafe {
            self.disp.GetIDsOfNames(
                &GUID::default(),
                names.as_ptr(),
                names.len() as u32,
                LOCALE_USER_DEFAULT,
                &mut dispid,
            )?;
        }
        Ok(dispid)
    }

    unsafe fn invoke_raw(
        &self,
        dispid: i32,
        flags: DISPATCH_FLAGS,
        params: &DISPPARAMS,
    ) -> windows::core::Result<VARIANT> {
        let mut result = VARIANT::default();
        self.disp.Invoke(
            dispid,
            &GUID::default(),
            LOCALE_USER_DEFAULT,
            flags,
            params,
            Some(&mut result),
            Some(std::ptr::null_mut()),
            Some(std::ptr::null_mut()),
        )?;
        Ok(result)
    }

    /// 读取属性，如 `get("Caption")` / `get("Workbooks")` / `get("ActiveSheet")`。
    /// WPS 个别成员是"方法"形式：属性读取报 MEMBERNOTFOUND 时自动退化为 METHOD 再试。
    pub fn get(&self, name: &str) -> Result<DispatchResult> {
        let dispid = self
            .get_dispid(name)
            .map_err(|e| com_err("dispid", name, &self.ctx, e))?;
        let params = DISPPARAMS {
            rgvarg: std::ptr::null_mut(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 0,
            cNamedArgs: 0,
        };
        match unsafe { self.invoke_raw(dispid, DISPATCH_PROPERTYGET, &params) } {
            Ok(v) => Ok(DispatchResult::new(v, &self.ctx)),
            Err(e) if e.code().0 as u32 == DISP_E_MEMBERNOTFOUND => unsafe {
                self.invoke_raw(dispid, DISPATCH_METHOD, &params)
                    .map(|v| DispatchResult::new(v, &self.ctx))
                    .map_err(|e| com_err("get", name, &self.ctx, e).into())
            },
            Err(e) => Err(com_err("get", name, &self.ctx, e).into()),
        }
    }

    /// 调用方法，如 `call("Add", vec![])` / `call("Range", vec!["A1".into()])`。
    /// `Range("A1")`/`Cells(1,1)` 这类参数化属性用 METHOD 报 MEMBERNOTFOUND 时
    /// 自动退化为 PROPERTYGET 再试，一套 `call` 同时覆盖两种形态。
    pub fn call(&self, name: &str, args: Vec<ComValue>) -> Result<DispatchResult> {
        let dispid = self
            .get_dispid(name)
            .map_err(|e| com_err("dispid", name, &self.ctx, e))?;
        let mut variants: Vec<VARIANT> = args.into_iter().map(VARIANT::from).collect();
        variants.reverse();
        let params = DISPPARAMS {
            rgvarg: variants.as_mut_ptr(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: variants.len() as u32,
            cNamedArgs: 0,
        };
        match unsafe { self.invoke_raw(dispid, DISPATCH_METHOD, &params) } {
            Ok(v) => Ok(DispatchResult::new(v, &self.ctx)),
            Err(e) if e.code().0 as u32 == DISP_E_MEMBERNOTFOUND => unsafe {
                self.invoke_raw(dispid, DISPATCH_PROPERTYGET, &params)
                    .map(|v| DispatchResult::new(v, &self.ctx))
                    .map_err(|e| com_err("call", name, &self.ctx, e).into())
            },
            Err(e) => Err(com_err("call", name, &self.ctx, e).into()),
        }
    }

    /// 带命名参数的方法调用，如 `call_named("SetSourceData", vec![], &[("Source", v)])`。
    ///
    /// **顺序很重要**：COM 的 DISPPARAMS 规定 `rgvarg` 里**具名参数在前**（顺序与
    /// `rgdispidNamedArgs` 一一对应），位置参数排在后面且**逆序**。
    /// 这里曾经写成「位置参数逆序放前、具名参数追加在后」—— 运行时按 `cNamedArgs` 从头取具名参数，
    /// 取到的其实是位置参数，于是：同类型的参数被静默对调（`Find` 的 What/LookAt、`Copy` 的 After…），
    /// 类型不同的直接报 `0x80020005 类型不匹配`（条件格式的 `Add`，用户实报）。
    pub fn call_named(
        &self,
        name: &str,
        args: Vec<ComValue>,
        named: &[(&str, ComValue)],
    ) -> Result<DispatchResult> {
        let mut names: Vec<BSTR> = Vec::with_capacity(1 + named.len());
        names.push(BSTR::from(name));
        names.extend(named.iter().map(|(n, _)| BSTR::from(*n)));
        let ptrs: Vec<PCWSTR> = names.iter().map(|b| PCWSTR(b.as_ptr())).collect();
        let mut dispids: Vec<i32> = vec![0; names.len()];
        unsafe {
            self.disp.GetIDsOfNames(
                &GUID::default(),
                ptrs.as_ptr(),
                ptrs.len() as u32,
                LOCALE_USER_DEFAULT,
                dispids.as_mut_ptr(),
            )
        }
        .map_err(|e| com_err("dispid", name, &self.ctx, e))?;

        let mut variants: Vec<VARIANT> = Vec::with_capacity(args.len() + named.len());
        for (_, v) in named {
            variants.push(VARIANT::from(v.clone()));
        }
        let mut positional: Vec<VARIANT> = args.into_iter().map(VARIANT::from).collect();
        positional.reverse();
        variants.extend(positional);
        let mut named_ids: Vec<i32> = dispids[1..].to_vec();
        let params = DISPPARAMS {
            rgvarg: variants.as_mut_ptr(),
            rgdispidNamedArgs: named_ids.as_mut_ptr(),
            cArgs: variants.len() as u32,
            cNamedArgs: named_ids.len() as u32,
        };
        unsafe { self.invoke_raw(dispids[0], DISPATCH_METHOD, &params) }
            .map(|v| DispatchResult::new(v, &self.ctx))
            .map_err(|e| com_err("call", name, &self.ctx, e).into())
    }

    /// 写属性，如 `put("Value", vec!["Hello".into()])`。
    pub fn put(&self, name: &str, args: Vec<ComValue>) -> Result<()> {
        let variants: Vec<VARIANT> = args.into_iter().map(VARIANT::from).collect();
        self.put_variants(name, variants)
    }

    /// put 一个已构造好的 VARIANT（供 SAFEARRAY 批量写入使用）。
    pub fn put_raw(&self, name: &str, value: VARIANT) -> Result<()> {
        self.put_variants(name, vec![value])
    }

    fn put_variants(&self, name: &str, mut variants: Vec<VARIANT>) -> Result<()> {
        let dispid = self
            .get_dispid(name)
            .map_err(|e| com_err("dispid", name, &self.ctx, e))?;
        variants.reverse();
        let mut dispid_named: i32 = DISPID_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: variants.as_mut_ptr(),
            rgdispidNamedArgs: &mut dispid_named,
            cArgs: variants.len() as u32,
            cNamedArgs: 1,
        };
        unsafe { self.invoke_raw(dispid, DISPATCH_PROPERTYPUT, &params) }
            .map_err(|e| com_err("put", name, &self.ctx, e).into())
            .map(|_| ())
    }
}

/// 方法调用 / 属性读取的返回值。携带调用处的对象路径，便于错误定位。
pub struct DispatchResult {
    pub v: VARIANT,
    ctx: String,
}

impl DispatchResult {
    pub(crate) fn new(v: VARIANT, ctx: &str) -> Self {
        Self {
            v,
            ctx: ctx.to_string(),
        }
    }

    /// 返回值是否为 COM 对象（VT_DISPATCH）。
    pub fn is_object(&self) -> bool {
        unsafe { self.v.Anonymous.Anonymous.vt }.0 == VT_DISPATCH.0
    }

    /// 按 vt 转成 [CellValue]：空单元格 / 数字 / 文本 / 布尔 / 错误值 都安全。
    pub fn value(self) -> CellValue {
        variant_to_cellvalue(&self.v)
    }

    /// 转文本（数字/布尔/空/错误值都给出可读形式）。
    pub fn string(self) -> Result<String> {
        Ok(self.value().to_string())
    }

    pub fn int(self) -> Result<i32> {
        let Self { v, ctx } = self;
        match variant_to_cellvalue(&v) {
            CellValue::Number(f) => Ok(f as i32),
            CellValue::Bool(b) => Ok(b as i32),
            other => Err(ExcelError::new("to_int", ctx, format!("期望整数，实际为 {other:?}")).into()),
        }
    }

    pub fn float(self) -> Result<f64> {
        let Self { v, ctx } = self;
        match variant_to_cellvalue(&v) {
            CellValue::Number(f) => Ok(f),
            other => Err(ExcelError::new("to_float", ctx, format!("期望数字，实际为 {other:?}")).into()),
        }
    }

    pub fn bool(self) -> Result<bool> {
        let Self { v, ctx } = self;
        match variant_to_cellvalue(&v) {
            CellValue::Bool(b) => Ok(b),
            CellValue::Number(f) => Ok(f != 0.0),
            other => Err(ExcelError::new("to_bool", ctx, format!("期望布尔值，实际为 {other:?}")).into()),
        }
    }

    /// 转为 COM 对象包装（ctx 继承调用处的对象路径）。
    pub fn idispatch(self) -> Result<IDispatchWrapper> {
        let ctx = self.ctx;
        let disp: Option<IDispatch> =
            unsafe { (*self.v.Anonymous.Anonymous.Anonymous.pdispVal).clone() };
        match disp {
            Some(d) => Ok(IDispatchWrapper { disp: d, ctx }),
            None => Err(ExcelError::new(
                "to_object",
                ctx,
                "返回值不是 COM 对象（可能为空/标量/数组）",
            )
            .into()),
        }
    }

    /// 转为 COM 对象包装，ctx 追加子对象标签。
    pub fn idispatch_child(self, label: &str) -> Result<IDispatchWrapper> {
        let ctx = format!("{} > {}", self.ctx, label);
        let disp: Option<IDispatch> =
            unsafe { (*self.v.Anonymous.Anonymous.Anonymous.pdispVal).clone() };
        match disp {
            Some(d) => Ok(IDispatchWrapper { disp: d, ctx }),
            None => Err(ExcelError::new(
                "to_object",
                ctx,
                "返回值不是 COM 对象（可能为空/标量/数组）",
            )
            .into()),
        }
    }
}
