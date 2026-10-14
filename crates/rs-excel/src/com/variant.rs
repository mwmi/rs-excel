//! COM 值类型：`ComValue`（传给 COM）与 `CellValue`（从 COM 读出，AI 直接可读）。

use std::mem::ManuallyDrop;

use windows::core::BSTR;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::{
    VARIANT, VT_BOOL, VT_BSTR, VT_DATE, VT_DISPATCH, VT_EMPTY, VT_ERROR, VT_I2, VT_I4, VT_I8,
    VT_NULL, VT_R4, VT_R8,
};

/// 传给 COM 的参数值。实现了常用类型的 `From`，调用处直接 `"A1".into()` / `123.into()`。
#[derive(Clone)]
pub enum ComValue {
    Str(String),
    Bool(bool),
    Int(i32),
    Float(f64),
    /// VT_EMPTY：**空值**（多数成员的可选参数收它就行，如 `Worksheets.Add` 的 Before）
    Empty,
    /// 真的「省略这个可选参数」：VT_ERROR + DISP_E_PARAMNOTFOUND（VBA 里 `, ,` 那个空档）。
    /// VT_EMPTY 是空**值**不是省略 —— 引擎会把它当参数值去解析，解析不了就直接抛异常
    /// （实报：Excel 上 `Slicers.Add` 的 Level「非 OLAP 不支持此参数」，传 Empty 必抛 0x80020009）
    Missing,
    /// COM 对象参数（如把 Range 传给 SetSourceData）
    Obj(IDispatch),
    /// 预构造的 VARIANT（如 SAFEARRAY 数组参数），原样传递
    Raw(VARIANT),
}

impl std::fmt::Debug for ComValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Str(s) => write!(f, "Str({s:?})"),
            Self::Bool(b) => write!(f, "Bool({b})"),
            Self::Int(i) => write!(f, "Int({i})"),
            Self::Float(v) => write!(f, "Float({v})"),
            Self::Empty => write!(f, "Empty"),
            Self::Missing => write!(f, "Missing"),
            Self::Obj(_) => write!(f, "Obj(<IDispatch>)"),
            Self::Raw(v) => {
                let vt = unsafe { v.Anonymous.Anonymous.vt }.0;
                write!(f, "Raw(<VARIANT vt=0x{vt:04X}>)")
            }
        }
    }
}

impl From<&str> for ComValue {
    fn from(s: &str) -> Self {
        Self::Str(s.to_string())
    }
}
impl From<String> for ComValue {
    fn from(s: String) -> Self {
        Self::Str(s)
    }
}
impl From<bool> for ComValue {
    fn from(b: bool) -> Self {
        Self::Bool(b)
    }
}
impl From<i32> for ComValue {
    fn from(i: i32) -> Self {
        Self::Int(i)
    }
}
impl From<u32> for ComValue {
    fn from(i: u32) -> Self {
        Self::Int(i as i32)
    }
}
impl From<f64> for ComValue {
    fn from(f: f64) -> Self {
        Self::Float(f)
    }
}
impl From<IDispatch> for ComValue {
    fn from(d: IDispatch) -> Self {
        Self::Obj(d)
    }
}
impl From<CellValue> for ComValue {
    fn from(v: CellValue) -> Self {
        match v {
            CellValue::Empty => Self::Empty,
            CellValue::Text(s) => Self::Str(s),
            CellValue::Number(f) => Self::Float(f),
            CellValue::Bool(b) => Self::Bool(b),
            CellValue::Error(e) => Self::Str(e),
        }
    }
}

impl From<ComValue> for VARIANT {
    fn from(v: ComValue) -> VARIANT {
        match v {
            ComValue::Str(s) => VARIANT::from(BSTR::from(s.as_str())),
            ComValue::Bool(b) => VARIANT::from(b),
            ComValue::Int(i) => VARIANT::from(i),
            ComValue::Float(f) => VARIANT::from(f),
            ComValue::Empty => VARIANT::default(),
            ComValue::Missing => {
                // DISP_E_PARAMNOTFOUND：可选参数「省略」的官方编码（VBA 的 `, ,` 就是它）
                let mut var = VARIANT::default();
                unsafe {
                    (*var.Anonymous.Anonymous).vt = VT_ERROR;
                    (*var.Anonymous.Anonymous).Anonymous.scode = 0x8002_0004u32 as i32;
                }
                var
            }
            ComValue::Raw(v) => v,
            ComValue::Obj(d) => {
                let mut var = VARIANT::default();
                unsafe {
                    (*var.Anonymous.Anonymous).vt = VT_DISPATCH;
                    (*var.Anonymous.Anonymous).Anonymous.pdispVal = ManuallyDrop::new(Some(d));
                }
                var
            }
        }
    }
}

/// 单元格读出的值。`Display` 输出 AI/用户直接可读的文本。
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    Bool(bool),
    /// 单元格错误值（如 #REF!），携带错误码
    Error(String),
}

impl CellValue {
    /// TSV 输出：制表符/换行替换为空格，保证表格结构不被破坏。
    pub fn to_tsv(&self) -> String {
        match self {
            CellValue::Text(s) => s.replace(['\t', '\n', '\r'], " "),
            other => other.to_string(),
        }
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Empty => Ok(()),
            Self::Text(s) => f.write_str(s),
            Self::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    write!(f, "{}", *n as i64)
                } else {
                    write!(f, "{n}")
                }
            }
            Self::Bool(b) => write!(f, "{b}"),
            Self::Error(e) => write!(f, "#{e}"),
        }
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        Self::Text(s.to_string())
    }
}
impl From<String> for CellValue {
    fn from(s: String) -> Self {
        Self::Text(s)
    }
}
impl From<f64> for CellValue {
    fn from(f: f64) -> Self {
        Self::Number(f)
    }
}
impl From<i32> for CellValue {
    fn from(i: i32) -> Self {
        Self::Number(i as f64)
    }
}
impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        Self::Bool(b)
    }
}

/// Excel 序列日期（1899-12-30 纪元）→ `"yyyy-mm-dd"` / `"yyyy-mm-dd hh:mm:ss"`。
fn serial_to_date_text(serial: f64) -> String {
    let mut days = serial.floor() as i64;
    let rem = (serial - serial.floor()) * 86400.0;
    let mut secs = rem.round() as i64;
    if secs >= 86400 {
        // 四舍五入进位到次日
        days += 1;
        secs = 0;
    }
    // 天数自 1970-01-01 起算，套用 civil_from_days（Hinnant 算法）
    let z = days - 25569 + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let mut y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    if m <= 2 {
        y += 1;
    }
    let date = format!("{y:04}-{m:02}-{d:02}");
    if secs == 0 {
        date
    } else {
        format!(
            "{date} {:02}:{:02}:{:02}",
            secs / 3600,
            secs % 3600 / 60,
            secs % 60
        )
    }
    .trim_end()
    .to_string()
}

/// VARIANT → CellValue（严格按 vt 分派，杜绝直接读 union 拿到垃圾值）。
pub(crate) fn variant_to_cellvalue(v: &VARIANT) -> CellValue {
    let vt: u16 = unsafe { v.Anonymous.Anonymous.vt }.0;
    unsafe {
        let u = &v.Anonymous.Anonymous.Anonymous;
        if vt == VT_EMPTY.0 || vt == VT_NULL.0 {
            CellValue::Empty
        } else if vt == VT_BSTR.0 {
            let b: &BSTR = &u.bstrVal;
            CellValue::Text(b.to_string())
        } else if vt == VT_I4.0 {
            CellValue::Number(u.lVal as f64)
        } else if vt == VT_I2.0 {
            CellValue::Number(u.iVal as f64)
        } else if vt == VT_I8.0 {
            CellValue::Number(u.llVal as f64)
        } else if vt == VT_R4.0 {
            CellValue::Number(u.fltVal as f64)
        } else if vt == VT_R8.0 {
            CellValue::Number(u.dblVal)
        } else if vt == VT_DATE.0 {
            // 日期以 Excel 序列值（f64）返回，README 有说明
            CellValue::Number(u.dblVal)
        } else if vt == VT_BOOL.0 {
            CellValue::Bool(bool::from(u.boolVal))
        } else if vt == VT_ERROR.0 {
            CellValue::Error(format!("0x{:08X}", u.scode as u32))
        } else {
            CellValue::Text(format!("<未支持的 VARIANT 类型 0x{vt:04X}>"))
        }
    }
}

/// VARIANT → CellValue（类型化通道）：唯一差异是 `VT_DATE` → 可读日期文本，
/// 供 `Value`（非 Value2）兼容读取使用。
pub(crate) fn variant_to_cellvalue_typed(v: &VARIANT) -> CellValue {
    if unsafe { v.Anonymous.Anonymous.vt }.0 == VT_DATE.0 {
        let serial = unsafe { v.Anonymous.Anonymous.Anonymous.dblVal };
        return CellValue::Text(serial_to_date_text(serial));
    }
    variant_to_cellvalue(v)
}
