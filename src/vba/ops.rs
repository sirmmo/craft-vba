//! VBA's operators on values ([MS-VBAL] §5.6.9): arithmetic with Integer/Long overflow,
//! comparison across numbers and strings, `Like`, and the logical/bitwise operators.

use std::cmp::Ordering;

use super::ast::{BinOp, UnOp};
use super::value::{R, Value, VbaError, int_value, round_half_even};

/// How wide a numeric operand is: Integer (and Boolean, Empty) < Long < Double; Date is Double.
fn rank(v: &Value) -> u8 {
    match v {
        Value::Empty | Value::Bool(_) | Value::Int(_) => 1,
        Value::Long(_) => 2,
        Value::Date(_) => 4,
        _ => 3,
    }
}

/// The result typed as the wider operand: Integer and Long overflow is error 6, as in VBA.
fn typed(v: f64, rank: u8) -> R<Value> {
    match rank {
        1 => {
            if v.is_finite() && v >= f64::from(i16::MIN) && v <= f64::from(i16::MAX) {
                Ok(Value::Int(v as i16))
            } else {
                Err(VbaError::new(6))
            }
        }
        2 => {
            if v.is_finite() && v >= f64::from(i32::MIN) && v <= f64::from(i32::MAX) {
                Ok(Value::Long(v as i32))
            } else {
                Err(VbaError::new(6))
            }
        }
        4 => Ok(Value::Date(v)),
        _ => {
            if v.is_finite() {
                Ok(Value::Double(v))
            } else {
                Err(VbaError::new(6))
            }
        }
    }
}

/// A string operand of arithmetic: its number (Double), or a type mismatch.
fn num_operand(v: &Value) -> R<(f64, u8)> {
    match v {
        Value::Str(_) => Ok((v.to_f64()?, 3)),
        Value::Null => Err(VbaError::new(94)),
        Value::Array(_) | Value::Object(_) | Value::Nothing | Value::Missing => Err(VbaError::new(13)),
        other => Ok((other.to_f64()?, rank(other))),
    }
}

pub fn binary(op: BinOp, a: &Value, b: &Value, compare_text: bool) -> R<Value> {
    use BinOp::*;
    // Null propagates through arithmetic and comparison.
    let null = matches!(a, Value::Null) || matches!(b, Value::Null);
    match op {
        Concat => {
            let s = |v: &Value| if matches!(v, Value::Null) { Ok(String::new()) } else { v.to_str() };
            if matches!(a, Value::Null) && matches!(b, Value::Null) {
                return Ok(Value::Null);
            }
            Ok(Value::str(format!("{}{}", s(a)?, s(b)?)))
        }
        Add | Sub | Mul | Div | IntDiv | Mod | Pow if null => Ok(Value::Null),
        Add if matches!(a, Value::Str(_)) && matches!(b, Value::Str(_) | Value::Empty) || matches!(a, Value::Empty) && matches!(b, Value::Str(_)) => {
            Ok(Value::str(format!("{}{}", a.to_str()?, b.to_str()?)))
        }
        Add | Sub | Mul => {
            let ((x, ra), (y, rb)) = (num_operand(a)?, num_operand(b)?);
            let r = ra.max(rb);
            let v = match op {
                Add => x + y,
                Sub => x - y,
                _ => x * y,
            };
            // A Date minus a Date is a number of days.
            let r = if (op == Sub && ra == 4 && rb == 4) || (r == 4 && op == Mul) { 3 } else { r };
            typed(v, r)
        }
        Div => {
            let ((x, _), (y, _)) = (num_operand(a)?, num_operand(b)?);
            if y == 0.0 {
                return Err(VbaError::new(if x == 0.0 { 6 } else { 11 }));
            }
            typed(x / y, 3)
        }
        IntDiv | Mod => {
            let ((x, ra), (y, rb)) = (num_operand(a)?, num_operand(b)?);
            let (x, y) = (round_half_even(x), round_half_even(y));
            if y == 0.0 {
                return Err(VbaError::new(11));
            }
            let v = if op == IntDiv { (x / y).trunc() } else { x % y };
            // Integer when both are Integer-sized, else Long (Doubles are rounded to Long first).
            typed(v, if ra == 1 && rb == 1 { 1 } else { 2 })
        }
        Pow => {
            let ((x, _), (y, _)) = (num_operand(a)?, num_operand(b)?);
            if x < 0.0 && y.fract() != 0.0 {
                return Err(VbaError::new(5));
            }
            typed(x.powf(y), 3)
        }
        Eq | Ne | Lt | Gt | Le | Ge => {
            if null {
                return Ok(Value::Null);
            }
            let o = compare(a, b, compare_text)?;
            Ok(Value::Bool(match op {
                Eq => o == Ordering::Equal,
                Ne => o != Ordering::Equal,
                Lt => o == Ordering::Less,
                Gt => o == Ordering::Greater,
                Le => o != Ordering::Greater,
                _ => o != Ordering::Less,
            }))
        }
        Like => {
            if null {
                return Ok(Value::Null);
            }
            Ok(Value::Bool(like(&a.to_str()?, &b.to_str()?, compare_text)))
        }
        Is => Ok(Value::Bool(match (a, b) {
            (Value::Nothing, Value::Nothing) => true,
            (Value::Object(x), Value::Object(y)) => x == y,
            (Value::Object(_), Value::Nothing) | (Value::Nothing, Value::Object(_)) => false,
            _ => return Err(VbaError::new(424)),
        })),
        And | Or | Xor | Eqv | Imp => {
            if let (Value::Bool(x), Value::Bool(y)) = (a, b) {
                return Ok(Value::Bool(match op {
                    And => *x && *y,
                    Or => *x || *y,
                    Xor => x != y,
                    Eqv => x == y,
                    _ => !*x || *y,
                }));
            }
            if null {
                return Ok(Value::Null);
            }
            let (x, y) = (a.to_i32()?, b.to_i32()?);
            let v = match op {
                And => x & y,
                Or => x | y,
                Xor => x ^ y,
                Eqv => !(x ^ y),
                _ => !x | y,
            };
            Ok(if rank(a).max(rank(b)) == 1 { int_value(f64::from(v)) } else { Value::Long(v) })
        }
    }
}

pub fn unary(op: UnOp, v: &Value) -> R<Value> {
    match (op, v) {
        (_, Value::Null) => Ok(Value::Null),
        (UnOp::Not, Value::Bool(b)) => Ok(Value::Bool(!b)),
        (UnOp::Not, other) => {
            let x = other.to_i32()?;
            Ok(if rank(other) == 1 { int_value(f64::from(!x)) } else { Value::Long(!x) })
        }
        (UnOp::Neg, Value::Empty) => Ok(Value::Int(0)),
        (UnOp::Neg, Value::Int(i)) => Ok(i.checked_neg().map(Value::Int).unwrap_or(Value::Long(-i32::from(*i)))),
        (UnOp::Neg, Value::Long(i)) => i.checked_neg().map(Value::Long).ok_or(VbaError::new(6)),
        (UnOp::Neg, Value::Bool(b)) => Ok(Value::Int(if *b { 1 } else { 0 })),
        (UnOp::Neg, Value::Date(d)) => Ok(Value::Date(-d)),
        (UnOp::Neg, other) => {
            let (x, _) = num_operand(other)?;
            Ok(Value::Double(-x))
        }
    }
}

/// Compares two values: numbers numerically, strings by code (or ignoring case with
/// `Option Compare Text`); a number is less than a string that isn't one.
pub fn compare(a: &Value, b: &Value, text: bool) -> R<Ordering> {
    let num = |v: &Value| matches!(v, Value::Int(_) | Value::Long(_) | Value::Double(_) | Value::Date(_) | Value::Bool(_));
    match (a, b) {
        (Value::Str(x), Value::Str(y)) => Ok(str_cmp(x, y, text)),
        (Value::Empty, Value::Str(y)) => Ok(str_cmp("", y, text)),
        (Value::Str(x), Value::Empty) => Ok(str_cmp(x, "", text)),
        (Value::Str(s), n) if num(n) || matches!(n, Value::Empty) => Ok(match super::value::parse_number(s) {
            Some(x) => x.partial_cmp(&n.to_f64()?).unwrap_or(Ordering::Equal),
            None => Ordering::Greater,
        }),
        (n, Value::Str(s)) if num(n) || matches!(n, Value::Empty) => Ok(match super::value::parse_number(s) {
            Some(y) => n.to_f64()?.partial_cmp(&y).unwrap_or(Ordering::Equal),
            None => Ordering::Less,
        }),
        (x, y) if (num(x) || matches!(x, Value::Empty)) && (num(y) || matches!(y, Value::Empty)) => {
            Ok(x.to_f64()?.partial_cmp(&y.to_f64()?).unwrap_or(Ordering::Equal))
        }
        _ => Err(VbaError::new(13)),
    }
}

pub fn str_cmp(a: &str, b: &str, text: bool) -> Ordering {
    if text { a.to_lowercase().cmp(&b.to_lowercase()) } else { a.cmp(b) }
}

/// `Like` patterns: `?` one character, `*` any run, `#` a digit, `[a-z]` and `[!a-z]` lists.
pub fn like(s: &str, pattern: &str, text: bool) -> bool {
    let (s, p): (Vec<char>, Vec<char>) =
        if text { (s.to_lowercase().chars().collect(), pattern.to_lowercase().chars().collect()) } else { (s.chars().collect(), pattern.chars().collect()) };
    // Tokens of the pattern.
    enum Pat {
        Any,
        One,
        Digit,
        Lit(char),
        Set(bool, Vec<(char, char)>),
    }
    let mut toks = Vec::new();
    let mut i = 0;
    while let Some(&c) = p.get(i) {
        i += 1;
        toks.push(match c {
            '*' => Pat::Any,
            '?' => Pat::One,
            '#' => Pat::Digit,
            '[' => {
                let neg = p.get(i) == Some(&'!');
                if neg {
                    i += 1;
                }
                let mut ranges = Vec::new();
                while let Some(&x) = p.get(i) {
                    i += 1;
                    if x == ']' {
                        break;
                    }
                    if p.get(i) == Some(&'-') && p.get(i + 1).is_some_and(|y| *y != ']') {
                        let y = p.get(i + 1).copied().unwrap_or(x);
                        ranges.push((x, y));
                        i += 2;
                    } else {
                        ranges.push((x, x));
                    }
                }
                Pat::Set(neg, ranges)
            }
            other => Pat::Lit(other),
        });
    }
    // dp[j]: the first j pattern tokens match the text so far.
    let at = |v: &[bool], j: usize| v.get(j).copied().unwrap_or(false);
    let put = |v: &mut Vec<bool>, j: usize, x: bool| {
        if let Some(slot) = v.get_mut(j) {
            *slot = x;
        }
    };
    let mut dp = vec![false; toks.len() + 1];
    put(&mut dp, 0, true);
    for (j, t) in toks.iter().enumerate() {
        if matches!(t, Pat::Any) {
            let v = at(&dp, j);
            put(&mut dp, j + 1, v);
        }
    }
    for &c in &s {
        let mut next = vec![false; toks.len() + 1];
        for (j, t) in toks.iter().enumerate() {
            let here = at(&dp, j);
            let v = match t {
                // `*` takes this character (dp[j + 1]) or none of it (next[j]).
                Pat::Any => at(&dp, j + 1) || at(&next, j),
                Pat::One => here,
                Pat::Digit => here && c.is_ascii_digit(),
                Pat::Lit(l) => here && *l == c,
                Pat::Set(neg, ranges) => here && (ranges.iter().any(|(a, b)| *a <= c && c <= *b) != *neg),
            };
            put(&mut next, j + 1, v);
        }
        dp = next;
    }
    at(&dp, toks.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn b(op: BinOp, x: Value, y: Value) -> R<Value> {
        binary(op, &x, &y, false)
    }

    #[test]
    fn arithmetic_types_and_overflow() {
        assert_eq!(b(BinOp::Add, Value::Int(2), Value::Int(3)).unwrap(), Value::Int(5));
        assert_eq!(b(BinOp::Add, Value::Int(30000), Value::Int(30000)).unwrap_err().number, 6);
        assert_eq!(b(BinOp::Add, Value::Long(30000), Value::Int(30000)).unwrap(), Value::Long(60000));
        assert_eq!(b(BinOp::Div, Value::Int(7), Value::Int(2)).unwrap(), Value::Double(3.5));
        assert_eq!(b(BinOp::IntDiv, Value::Int(7), Value::Int(2)).unwrap(), Value::Int(3));
        assert_eq!(b(BinOp::Mod, Value::Int(-7), Value::Int(3)).unwrap(), Value::Int(-1));
        assert_eq!(b(BinOp::Div, Value::Int(1), Value::Int(0)).unwrap_err().number, 11);
        assert_eq!(b(BinOp::Add, Value::str("2"), Value::Int(3)).unwrap(), Value::Double(5.0));
        assert_eq!(b(BinOp::Add, Value::str("a"), Value::str("b")).unwrap(), Value::str("ab"));
        assert_eq!(b(BinOp::Add, Value::str("a"), Value::Int(1)).unwrap_err().number, 13);
        assert_eq!(b(BinOp::Concat, Value::Int(1), Value::Double(2.5)).unwrap(), Value::str("12.5"));
        assert_eq!(b(BinOp::Pow, Value::Int(2), Value::Int(10)).unwrap(), Value::Double(1024.0));
        assert_eq!(unary(UnOp::Neg, &Value::Int(i16::MIN)).unwrap(), Value::Long(32768));
    }

    #[test]
    fn comparisons_and_logic() {
        assert_eq!(b(BinOp::Lt, Value::Int(2), Value::str("10")).unwrap(), Value::Bool(true));
        assert_eq!(b(BinOp::Lt, Value::str("2"), Value::str("10")).unwrap(), Value::Bool(false), "strings compare as text");
        assert_eq!(b(BinOp::Lt, Value::Int(5), Value::str("abc")).unwrap(), Value::Bool(true));
        assert_eq!(binary(BinOp::Eq, &Value::str("ABC"), &Value::str("abc"), true).unwrap(), Value::Bool(true));
        assert_eq!(b(BinOp::Eq, Value::Empty, Value::Int(0)).unwrap(), Value::Bool(true));
        assert_eq!(b(BinOp::Eq, Value::Empty, Value::str("")).unwrap(), Value::Bool(true));
        assert_eq!(b(BinOp::And, Value::Bool(true), Value::Bool(false)).unwrap(), Value::Bool(false));
        assert_eq!(b(BinOp::And, Value::Int(12), Value::Int(10)).unwrap(), Value::Int(8));
        assert_eq!(unary(UnOp::Not, &Value::Int(0)).unwrap(), Value::Int(-1));
        assert_eq!(b(BinOp::Eq, Value::Null, Value::Int(1)).unwrap(), Value::Null);
    }

    #[test]
    fn like_patterns() {
        assert!(like("Hello", "H*o", false));
        assert!(like("Hello", "H?llo", false));
        assert!(!like("Hello", "h*", false));
        assert!(like("Hello", "h*", true));
        assert!(like("A1", "[A-C]#", false));
        assert!(!like("D1", "[A-C]#", false));
        assert!(like("D1", "[!A-C]#", false));
        assert!(like("", "*", false));
        assert!(like("abc", "*b*", false));
        assert!(!like("abc", "*d*", false));
    }
}
