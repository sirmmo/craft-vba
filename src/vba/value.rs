//! VBA values and their conversions ([MS-VBAL] §5.5: let-coercion; §5.6: operators).

use std::cell::RefCell;
use std::rc::Rc;

use crate::excel::XlObj;

/// A run-time error: VBA's number and description.
#[derive(Clone, Debug, PartialEq)]
pub struct VbaError {
    pub number: i32,
    pub description: String,
}

impl VbaError {
    pub fn new(number: i32) -> VbaError {
        VbaError { number, description: describe(number).to_string() }
    }
    pub fn with(number: i32, description: impl Into<String>) -> VbaError {
        VbaError { number, description: description.into() }
    }
}

/// VBA's messages for its error numbers.
pub fn describe(n: i32) -> &'static str {
    match n {
        3 => "Return without GoSub",
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        7 => "Out of memory",
        9 => "Subscript out of range",
        10 => "This array is fixed or temporarily locked",
        11 => "Division by zero",
        13 => "Type mismatch",
        17 => "Can't perform requested operation",
        20 => "Resume without error",
        28 => "Out of stack space",
        35 => "Sub or Function not defined",
        91 => "Object variable or With block variable not set",
        92 => "For loop not initialized",
        94 => "Invalid use of Null",
        424 => "Object required",
        438 => "Object doesn't support this property or method",
        440 => "Automation error",
        449 => "Argument not optional",
        450 => "Wrong number of arguments or invalid property assignment",
        1004 => "Application-defined or object-defined error",
        _ => "Application-defined or object-defined error",
    }
}

pub type R<T> = Result<T, VbaError>;

/// An array: its bounds per dimension and its elements (row-major on the first index).
#[derive(Clone, Debug, PartialEq)]
pub struct Array {
    pub bounds: Vec<(i64, i64)>,
    pub data: Vec<Value>,
}

/// Most elements one array may hold.
pub const MAX_ELEMENTS: usize = 10_000_000;

impl Array {
    pub fn new(bounds: Vec<(i64, i64)>) -> R<Array> {
        let mut n: usize = 1;
        for (lo, hi) in &bounds {
            if hi < lo {
                if *hi == lo - 1 {
                    n = 0;
                    continue;
                }
                return Err(VbaError::new(9));
            }
            let len = usize::try_from(hi - lo + 1).map_err(|_| VbaError::new(7))?;
            n = n.checked_mul(len).ok_or(VbaError::new(7))?;
        }
        if n > MAX_ELEMENTS {
            return Err(VbaError::new(7));
        }
        Ok(Array { bounds, data: vec![Value::Empty; n] })
    }

    pub fn index(&self, idx: &[i64]) -> R<usize> {
        if idx.len() != self.bounds.len() {
            return Err(VbaError::new(9));
        }
        let mut at: usize = 0;
        for (i, (lo, hi)) in idx.iter().zip(&self.bounds) {
            if i < lo || i > hi {
                return Err(VbaError::new(9));
            }
            let len = usize::try_from(hi - lo + 1).map_err(|_| VbaError::new(9))?;
            at = at * len + usize::try_from(i - lo).map_err(|_| VbaError::new(9))?;
        }
        Ok(at)
    }

    pub fn get(&self, idx: &[i64]) -> R<Value> {
        let i = self.index(idx)?;
        self.data.get(i).cloned().ok_or(VbaError::new(9))
    }

    pub fn set(&mut self, idx: &[i64], v: Value) -> R<()> {
        let i = self.index(idx)?;
        match self.data.get_mut(i) {
            Some(slot) => {
                *slot = v;
                Ok(())
            }
            None => Err(VbaError::new(9)),
        }
    }

    /// A 1-D array of `items`, from `lo`.
    pub fn from_vec(items: Vec<Value>, lo: i64) -> Array {
        let hi = lo + items.len() as i64 - 1;
        Array { bounds: vec![(lo, hi)], data: items }
    }
}

/// An object reference.
#[derive(Clone, Debug, PartialEq)]
pub enum Obj {
    Xl(XlObj),
    /// The `Err` object.
    Err,
    /// The `Debug` object.
    Debug,
    /// A `Collection`.
    Collection(Rc<RefCell<Vec<(Option<String>, Value)>>>),
}

#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Empty,
    Null,
    Bool(bool),
    Int(i16),
    Long(i32),
    Double(f64),
    Date(f64),
    Str(Rc<str>),
    Array(Rc<RefCell<Array>>),
    Object(Obj),
    Nothing,
    /// A left-out optional argument.
    Missing,
}

impl Value {
    pub fn str(s: impl Into<Rc<str>>) -> Value {
        Value::Str(s.into())
    }

    pub fn type_name(&self) -> &'static str {
        match self {
            Value::Empty => "Empty",
            Value::Null => "Null",
            Value::Bool(_) => "Boolean",
            Value::Int(_) => "Integer",
            Value::Long(_) => "Long",
            Value::Double(_) => "Double",
            Value::Date(_) => "Date",
            Value::Str(_) => "String",
            Value::Array(_) => "Variant()",
            Value::Object(Obj::Xl(x)) => x.type_name(),
            Value::Object(Obj::Err) => "ErrObject",
            Value::Object(Obj::Debug) => "Debug",
            Value::Object(Obj::Collection(_)) => "Collection",
            Value::Nothing => "Nothing",
            Value::Missing => "Error",
        }
    }

    pub fn is_numeric(&self) -> bool {
        matches!(self, Value::Int(_) | Value::Long(_) | Value::Double(_) | Value::Date(_) | Value::Bool(_) | Value::Empty)
    }

    /// The number a value is, by let-coercion (strings that read as numbers too).
    pub fn to_f64(&self) -> R<f64> {
        match self {
            Value::Empty => Ok(0.0),
            Value::Bool(b) => Ok(if *b { -1.0 } else { 0.0 }),
            Value::Int(i) => Ok(f64::from(*i)),
            Value::Long(i) => Ok(f64::from(*i)),
            Value::Double(d) | Value::Date(d) => Ok(*d),
            Value::Str(s) => parse_number(s).ok_or(VbaError::new(13)),
            Value::Null => Err(VbaError::new(94)),
            _ => Err(VbaError::new(13)),
        }
    }

    /// Banker's rounding to a Long (`CLng`).
    pub fn to_i32(&self) -> R<i32> {
        let f = round_half_even(self.to_f64()?);
        if f < f64::from(i32::MIN) - 0.5 || f > f64::from(i32::MAX) + 0.5 || !f.is_finite() {
            return Err(VbaError::new(6));
        }
        Ok(f as i32)
    }

    pub fn to_i64(&self) -> R<i64> {
        Ok(i64::from(self.to_i32()?))
    }

    pub fn to_i16(&self) -> R<i16> {
        let v = self.to_i32()?;
        i16::try_from(v).map_err(|_| VbaError::new(6))
    }

    pub fn to_bool(&self) -> R<bool> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Str(s) => match s.trim().to_ascii_lowercase().as_str() {
                "true" => Ok(true),
                "false" => Ok(false),
                _ => parse_number(s).map(|n| n != 0.0).ok_or(VbaError::new(13)),
            },
            other => Ok(other.to_f64()? != 0.0),
        }
    }

    /// The string a value is (`CStr`).
    pub fn to_str(&self) -> R<String> {
        Ok(match self {
            Value::Empty | Value::Missing => String::new(),
            Value::Bool(b) => if *b { "True" } else { "False" }.to_string(),
            Value::Int(i) => i.to_string(),
            Value::Long(i) => i.to_string(),
            Value::Double(d) => format_number(*d),
            Value::Date(d) => format_number(*d),
            Value::Str(s) => s.to_string(),
            Value::Null => return Err(VbaError::new(94)),
            _ => return Err(VbaError::new(13)),
        })
    }
}

/// Reads a number the way `Val` and let-coercion do (spaces around, `&H`, a decimal point).
pub fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    if let Some(h) = t.strip_prefix("&H").or_else(|| t.strip_prefix("&h")) {
        return i64::from_str_radix(h, 16).ok().map(|v| v as f64);
    }
    let t = t.replace(',', "");
    t.parse::<f64>().ok().filter(|f| f.is_finite())
}

pub fn round_half_even(f: f64) -> f64 {
    let r = f.round();
    if (f - f.trunc()).abs() == 0.5 { 2.0 * (f / 2.0).round() } else { r }
}

/// A Double as VBA writes it: up to 15 significant digits, no trailing zeros, E notation for
/// large and small magnitudes.
pub fn format_number(d: f64) -> String {
    if d == 0.0 {
        return "0".into();
    }
    if !d.is_finite() {
        return if d.is_nan() {
            "NaN".into()
        } else if d > 0.0 {
            "Infinity".into()
        } else {
            "-Infinity".into()
        };
    }
    let e = format!("{d:.14e}");
    let (mantissa, exp) = e.split_once('e').unwrap_or((&e, "0"));
    let exp: i32 = exp.parse().unwrap_or(0);
    let negative = mantissa.starts_with('-');
    let digits: String = mantissa.chars().filter(|c| c.is_ascii_digit()).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let sign = if negative { "-" } else { "" };
    if !(-5..15).contains(&exp) {
        let (first, rest) = digits.split_at(1.min(digits.len()));
        let frac = if rest.is_empty() { String::new() } else { format!(".{rest}") };
        let es = if exp < 0 { "-" } else { "+" };
        return format!("{sign}{first}{frac}E{es}{:02}", exp.abs());
    }
    let n = digits.len() as i32;
    let s = if exp >= 0 {
        if n <= exp + 1 {
            format!("{digits}{}", "0".repeat((exp + 1 - n) as usize))
        } else {
            let (int, frac) = digits.split_at((exp + 1) as usize);
            format!("{int}.{frac}")
        }
    } else {
        format!("0.{}{digits}", "0".repeat((-exp - 1) as usize))
    };
    format!("{sign}{s}")
}

/// The narrowest numeric value for an integer result: Integer, Long, else Double.
pub fn int_value(v: f64) -> Value {
    if v.fract() == 0.0 && v >= f64::from(i16::MIN) && v <= f64::from(i16::MAX) {
        Value::Int(v as i16)
    } else if v.fract() == 0.0 && v >= f64::from(i32::MIN) && v <= f64::from(i32::MAX) {
        Value::Long(v as i32)
    } else {
        Value::Double(v)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numbers_print_like_vba() {
        assert_eq!(format_number(0.1 + 0.2), "0.3");
        assert_eq!(format_number(1.0 / 3.0), "0.333333333333333");
        assert_eq!(format_number(1234.5), "1234.5");
        assert_eq!(format_number(-2.0), "-2");
        assert_eq!(format_number(1e15), "1E+15");
        assert_eq!(format_number(1.5e-7), "1.5E-07");
        assert_eq!(format_number(0.00012), "0.00012");
        assert_eq!(round_half_even(2.5), 2.0);
        assert_eq!(round_half_even(3.5), 4.0);
        assert_eq!(round_half_even(-2.5), -2.0);
        assert_eq!(Value::str("2.5").to_i32().unwrap(), 2);
        assert_eq!(Value::Double(40000.0).to_i16().unwrap_err().number, 6);
        assert_eq!(Value::str("abc").to_f64().unwrap_err().number, 13);
        assert_eq!(Value::Bool(true).to_str().unwrap(), "True");
    }

    #[test]
    fn arrays_check_their_bounds() {
        let mut a = Array::new(vec![(1, 3), (0, 1)]).unwrap();
        a.set(&[3, 1], Value::Int(7)).unwrap();
        assert_eq!(a.get(&[3, 1]).unwrap(), Value::Int(7));
        assert_eq!(a.get(&[4, 0]).unwrap_err().number, 9);
        assert_eq!(a.get(&[1]).unwrap_err().number, 9);
        assert!(Array::new(vec![(0, 1_000_000), (0, 1_000_000)]).is_err());
        assert_eq!(Array::new(vec![(0, -1)]).unwrap().data.len(), 0);
    }
}
