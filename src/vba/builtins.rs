//! VBA's built-in functions that don't touch the host ([MS-VBAL] §6), and its constants.

use std::cell::RefCell;
use std::rc::Rc;

use super::ops;
use super::value::{Array, R, Value, VbaError, format_number, int_value, parse_number, round_half_even};

/// A built-in constant (`vbCrLf`, `xlUp`).
pub fn constant(name: &str) -> Option<Value> {
    let s = |x: &str| Some(Value::str(x));
    let l = |x: i32| Some(Value::Long(x));
    let i = |x: i16| Some(Value::Int(x));
    match name.to_ascii_lowercase().as_str() {
        "vbcrlf" | "vbnewline" => s("\r\n"),
        "vbcr" => s("\r"),
        "vblf" => s("\n"),
        "vbtab" => s("\t"),
        "vbnullstring" => s(""),
        "vbnullchar" => s("\0"),
        "vbokonly" => i(0),
        "vbokcancel" => i(1),
        "vbabortretryignore" => i(2),
        "vbyesnocancel" => i(3),
        "vbyesno" => i(4),
        "vbretrycancel" => i(5),
        "vbcritical" => i(16),
        "vbquestion" => i(32),
        "vbexclamation" => i(48),
        "vbinformation" => i(64),
        "vbok" => i(1),
        "vbcancel" => i(2),
        "vbabort" => i(3),
        "vbretry" => i(4),
        "vbignore" => i(5),
        "vbyes" => i(6),
        "vbno" => i(7),
        "vbblack" => l(0),
        "vbred" => l(255),
        "vbgreen" => l(65_280),
        "vbyellow" => l(65_535),
        "vbblue" => l(16_711_680),
        "vbmagenta" => l(16_711_935),
        "vbcyan" => l(16_776_960),
        "vbwhite" => l(16_777_215),
        "vbbinarycompare" => i(0),
        "vbtextcompare" => i(1),
        "vbobjecterror" => l(-2_147_221_504),
        "vbempty" => i(0),
        "vbnull" => i(1),
        "vbinteger" => i(2),
        "vblong" => i(3),
        "vbsingle" => i(4),
        "vbdouble" => i(5),
        "vbcurrency" => i(6),
        "vbdate" => i(7),
        "vbstring" => i(8),
        "vbobject" => i(9),
        "vberror" => i(10),
        "vbboolean" => i(11),
        "vbvariant" => i(12),
        "vbarray" => i(8192),
        "xlup" => l(-4162),
        "xldown" => l(-4121),
        "xltoleft" => l(-4159),
        "xltoright" => l(-4161),
        "xlnone" | "xlunderlinestylenone" => l(-4142),
        "xlunderlinestylesingle" => l(2),
        "xlunderlinestyledouble" => l(-4119),
        "xlcalculationautomatic" => l(-4105),
        "xlcalculationmanual" => l(-4135),
        "xlvalues" => l(-4163),
        "xlformulas" => l(-4123),
        "xlwhole" => l(1),
        "xlpart" => l(2),
        "xlautomatic" => l(-4105),
        _ => None,
    }
}

fn arg(args: &[Value], i: usize) -> Option<&Value> {
    args.get(i).filter(|v| !matches!(v, Value::Missing))
}

fn need(args: &[Value], i: usize) -> R<&Value> {
    arg(args, i).ok_or(VbaError::new(449))
}

fn s(args: &[Value], i: usize) -> R<String> {
    need(args, i)?.to_str()
}

fn n(args: &[Value], i: usize) -> R<f64> {
    need(args, i)?.to_f64()
}

fn count(args: &[Value], i: usize) -> R<usize> {
    let v = need(args, i)?.to_i64()?;
    usize::try_from(v).map_err(|_| VbaError::new(5))
}

/// Characters of `s` from `start` (0-based), at most `len`.
fn chars_of(s: &str, start: usize, len: usize) -> String {
    s.chars().skip(start).take(len).collect()
}

/// Days since 1899-12-30 for a civil date.
pub fn date_serial(y: i64, m: i64, d: i64) -> f64 {
    // Months beyond 12 or below 1 roll the year, as DateSerial does.
    let total = y * 12 + (m - 1);
    let (y, m) = (total.div_euclid(12), total.rem_euclid(12) + 1);
    let y2 = if m <= 2 { y - 1 } else { y };
    let era = y2.div_euclid(400);
    let yoe = y2 - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days = era * 146_097 + doe - 719_468 + (d - 1) + 1;
    // 1970-01-01 is serial 25569.
    (days + 25_568) as f64
}

/// (year, month, day) of a serial date.
pub fn civil(serial: f64) -> (i64, i64, i64) {
    let z = serial.floor() as i64 - 25_569 + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = yoe + era * 400 + if m <= 2 { 1 } else { 0 };
    (y, m, d)
}

fn now() -> f64 {
    let secs = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_secs_f64()).unwrap_or(0.0);
    25_569.0 + secs / 86_400.0
}

/// `Format(number, "0.00")`-style patterns: decimals, thousands separators, percent.
fn format_value(v: &Value, fmt: &str) -> R<String> {
    let named = match fmt.to_ascii_lowercase().as_str() {
        "general number" => return v.to_str(),
        "fixed" => "0.00",
        "standard" => "#,##0.00",
        "percent" => "0.00%",
        "currency" => "$#,##0.00",
        "yes/no" => return Ok(if v.to_bool()? { "Yes" } else { "No" }.into()),
        "true/false" => return Ok(if v.to_bool()? { "True" } else { "False" }.into()),
        "on/off" => return Ok(if v.to_bool()? { "On" } else { "Off" }.into()),
        _ => fmt,
    };
    if named.is_empty() {
        return v.to_str();
    }
    let x = match v {
        Value::Str(st) => match parse_number(st) {
            Some(x) => x,
            None => return Ok(st.to_string()),
        },
        other => other.to_f64()?,
    };
    if named.contains(['y', 'd', 'h', 's']) && !named.contains('0') && !named.contains('#') {
        // Date patterns: yyyy-mm-dd and friends.
        let (y, m, d) = civil(x);
        let frac = x - x.floor();
        let secs = (frac * 86_400.0).round() as i64;
        let out = named
            .replace("yyyy", &format!("{y:04}"))
            .replace("mm", &format!("{m:02}"))
            .replace("dd", &format!("{d:02}"))
            .replace("hh", &format!("{:02}", secs / 3600))
            .replace("nn", &format!("{:02}", secs / 60 % 60))
            .replace("ss", &format!("{:02}", secs % 60));
        return Ok(out);
    }
    let percent = named.contains('%');
    let x = if percent { x * 100.0 } else { x };
    let num_part: String = named.chars().filter(|c| matches!(c, '0' | '#' | '.' | ',')).collect();
    let decimals = num_part.split_once('.').map(|(_, f)| f.chars().filter(|c| *c == '0' || *c == '#').count()).unwrap_or(0);
    let min_int = num_part.split('.').next().unwrap_or("").chars().filter(|c| *c == '0').count();
    let grouped = num_part.contains(',');
    let factor = 10f64.powi(decimals as i32);
    let rounded = (x.abs() * factor).round() / factor;
    let text = format!("{rounded:.decimals$}");
    let (int, frac) = text.split_once('.').unwrap_or((&text, ""));
    let mut int = int.trim_start_matches('0').to_string();
    while int.len() < min_int {
        int.insert(0, '0');
    }
    if grouped {
        let digits: Vec<char> = int.chars().collect();
        let mut g = String::new();
        for (k, c) in digits.iter().enumerate() {
            if k > 0 && (digits.len() - k).is_multiple_of(3) {
                g.push(',');
            }
            g.push(*c);
        }
        int = g;
    }
    let mut body = if frac.is_empty() { int } else { format!("{int}.{frac}") };
    if body.is_empty() {
        body = "0".into();
    }
    let prefix: String = named.chars().take_while(|c| !matches!(c, '0' | '#' | '.' | ',')).filter(|c| *c != '\\' && *c != '"').collect();
    let negative = x < 0.0 && rounded != 0.0;
    let sign = if negative { "-" } else { "" };
    Ok(format!("{sign}{prefix}{body}{}", if percent { "%" } else { "" }))
}

fn array_of(items: Vec<Value>, lo: i64) -> Value {
    Value::Array(Rc::new(RefCell::new(Array::from_vec(items, lo))))
}

/// A pure built-in function, or `None` when there is none by that name. `base` is the module's
/// `Option Base`, `text` its `Option Compare Text`.
pub fn call(name: &str, args: &[Value], base: i64, text: bool, rnd: &mut u64) -> Option<R<Value>> {
    let r = (|| -> R<Option<Value>> {
        Ok(Some(match name.to_ascii_lowercase().as_str() {
            "len" | "lenb" => match need(args, 0)? {
                Value::Null => Value::Null,
                v => Value::Long(v.to_str()?.chars().count() as i32),
            },
            "left" | "left$" => Value::str(chars_of(&s(args, 0)?, 0, count(args, 1)?)),
            "right" | "right$" => {
                let st = s(args, 0)?;
                let len = st.chars().count();
                let k = count(args, 1)?.min(len);
                Value::str(chars_of(&st, len - k, k))
            }
            "mid" | "mid$" => {
                let st = s(args, 0)?;
                let start = need(args, 1)?.to_i64()?;
                if start < 1 {
                    return Err(VbaError::new(5));
                }
                let len = match arg(args, 2) {
                    Some(v) => usize::try_from(v.to_i64()?).map_err(|_| VbaError::new(5))?,
                    None => usize::MAX,
                };
                Value::str(chars_of(&st, (start - 1) as usize, len))
            }
            "instr" => {
                // InStr([start,] s1, s2[, compare]).
                let (start, a, b, cmp) = if args.len() >= 3 && need(args, 0)?.is_numeric() && !matches!(need(args, 0)?, Value::Str(_)) {
                    (need(args, 0)?.to_i64()?, s(args, 1)?, s(args, 2)?, arg(args, 3))
                } else {
                    (1, s(args, 0)?, s(args, 1)?, None)
                };
                if start < 1 {
                    return Err(VbaError::new(5));
                }
                let insensitive = match cmp {
                    Some(c) => c.to_i32()? == 1,
                    None => text,
                };
                let (hay, needle) = if insensitive { (a.to_lowercase(), b.to_lowercase()) } else { (a, b) };
                let hay_chars: Vec<char> = hay.chars().collect();
                let skip = (start - 1) as usize;
                if skip > hay_chars.len() {
                    return Ok(Some(Value::Long(0)));
                }
                let rest: String = hay_chars.iter().skip(skip).collect();
                Value::Long(match rest.find(&needle) {
                    Some(byte) => (rest.get(..byte).unwrap_or("").chars().count() + skip + 1) as i32,
                    None => 0,
                })
            }
            "instrrev" => {
                let (a, b) = (s(args, 0)?, s(args, 1)?);
                Value::Long(match a.rfind(&b) {
                    Some(byte) => (a.get(..byte).unwrap_or("").chars().count() + 1) as i32,
                    None => 0,
                })
            }
            "replace" => {
                let (src, find, with) = (s(args, 0)?, s(args, 1)?, s(args, 2)?);
                let start = arg(args, 3).map(Value::to_i64).transpose()?.unwrap_or(1).max(1) as usize;
                let limit = arg(args, 4).map(Value::to_i64).transpose()?.unwrap_or(-1);
                let tail: String = src.chars().skip(start - 1).collect();
                if find.is_empty() {
                    Value::str(tail)
                } else if limit < 0 {
                    Value::str(tail.replace(&find, &with))
                } else {
                    Value::str(tail.replacen(&find, &with, limit as usize))
                }
            }
            "trim" | "trim$" => Value::str(s(args, 0)?.trim_matches(' ')),
            "ltrim" | "ltrim$" => Value::str(s(args, 0)?.trim_start_matches(' ')),
            "rtrim" | "rtrim$" => Value::str(s(args, 0)?.trim_end_matches(' ')),
            "ucase" | "ucase$" => Value::str(s(args, 0)?.to_uppercase()),
            "lcase" | "lcase$" => Value::str(s(args, 0)?.to_lowercase()),
            "strreverse" => Value::str(s(args, 0)?.chars().rev().collect::<String>()),
            "space" | "space$" => Value::str(" ".repeat(count(args, 0)?.min(1 << 24))),
            "string" | "string$" => {
                let k = count(args, 0)?.min(1 << 24);
                let c = match need(args, 1)? {
                    Value::Str(x) => x.chars().next().ok_or(VbaError::new(5))?,
                    v => char::from_u32(v.to_i64()? as u32).ok_or(VbaError::new(5))?,
                };
                Value::str(std::iter::repeat_n(c, k).collect::<String>())
            }
            "split" => {
                let st = s(args, 0)?;
                let delim = arg(args, 1).map(Value::to_str).transpose()?.unwrap_or_else(|| " ".into());
                let limit = arg(args, 2).map(Value::to_i64).transpose()?.unwrap_or(-1);
                let parts: Vec<Value> = if st.is_empty() {
                    Vec::new()
                } else if delim.is_empty() {
                    vec![Value::str(st)]
                } else if limit > 0 {
                    st.splitn(limit as usize, &delim).map(Value::str).collect()
                } else {
                    st.split(&delim).map(Value::str).collect()
                };
                array_of(parts, 0)
            }
            "join" => {
                let Value::Array(a) = need(args, 0)? else { return Err(VbaError::new(13)) };
                let delim = arg(args, 1).map(Value::to_str).transpose()?.unwrap_or_else(|| " ".into());
                let parts = a.borrow().data.iter().map(Value::to_str).collect::<R<Vec<_>>>()?;
                Value::str(parts.join(&delim))
            }
            "format" | "format$" => {
                let v = need(args, 0)?;
                let f = arg(args, 1).map(Value::to_str).transpose()?.unwrap_or_default();
                Value::str(format_value(v, &f)?)
            }
            "cstr" => Value::str(need(args, 0)?.to_str()?),
            "cint" => Value::Int(need(args, 0)?.to_i16()?),
            "clng" => Value::Long(need(args, 0)?.to_i32()?),
            "cbyte" => {
                let v = need(args, 0)?.to_i32()?;
                if !(0..=255).contains(&v) {
                    return Err(VbaError::new(6));
                }
                Value::Int(v as i16)
            }
            "cdbl" | "csng" | "ccur" | "cdec" => Value::Double(n(args, 0)?),
            "cbool" => Value::Bool(need(args, 0)?.to_bool()?),
            "cvar" => need(args, 0)?.clone(),
            "cdate" | "datevalue" => match need(args, 0)? {
                Value::Str(st) => Value::Date(parse_date(st).ok_or(VbaError::new(13))?),
                v => Value::Date(v.to_f64()?),
            },
            "val" => {
                let st = s(args, 0)?;
                let t: String = st.chars().filter(|c| !c.is_whitespace()).collect();
                let mut end = 0;
                let mut seen_dot = false;
                for (k, c) in t.char_indices() {
                    if c.is_ascii_digit() || (k == 0 && (c == '-' || c == '+')) || (c == '.' && !seen_dot) {
                        seen_dot |= c == '.';
                        end = k + c.len_utf8();
                    } else {
                        break;
                    }
                }
                Value::Double(t.get(..end).and_then(|x| x.parse::<f64>().ok()).or_else(|| parse_number(&t)).unwrap_or(0.0))
            }
            "isnumeric" => Value::Bool(match need(args, 0)? {
                Value::Str(st) => parse_number(st).is_some(),
                v => v.is_numeric(),
            }),
            "isempty" => Value::Bool(matches!(need(args, 0)?, Value::Empty)),
            "isnull" => Value::Bool(matches!(need(args, 0)?, Value::Null)),
            "isarray" => Value::Bool(matches!(need(args, 0)?, Value::Array(_))),
            "isobject" => Value::Bool(matches!(need(args, 0)?, Value::Object(_) | Value::Nothing)),
            "ismissing" => Value::Bool(matches!(args.first(), Some(Value::Missing) | None)),
            "isdate" => Value::Bool(match need(args, 0)? {
                Value::Date(_) => true,
                Value::Str(st) => parse_date(st).is_some(),
                _ => false,
            }),
            "iserror" => Value::Bool(false),
            "str" | "str$" => {
                let x = n(args, 0)?;
                let t = format_number(x);
                Value::str(if x >= 0.0 { format!(" {t}") } else { t })
            }
            "chr" | "chr$" | "chrw" | "chrw$" => {
                let c = need(args, 0)?.to_i64()?;
                let c = if name.to_ascii_lowercase().starts_with("chrw") {
                    char::from_u32(c as u32)
                } else {
                    (0..=255).contains(&c).then(|| super::super::ovba::codepage::decode(&[c as u8], 1252).chars().next()).flatten()
                };
                Value::str(c.ok_or(VbaError::new(5))?.to_string())
            }
            "asc" | "ascw" => {
                let st = s(args, 0)?;
                let c = st.chars().next().ok_or(VbaError::new(5))?;
                let code = if name.eq_ignore_ascii_case("asc") {
                    super::super::ovba::codepage::encode(&c.to_string(), 1252).first().copied().map(i32::from).unwrap_or(63)
                } else {
                    c as i32
                };
                Value::Int(code as i16)
            }
            "abs" => match need(args, 0)? {
                Value::Int(i) => i.checked_abs().map(Value::Int).unwrap_or(Value::Long(32768)),
                Value::Long(i) => Value::Long(i.checked_abs().ok_or(VbaError::new(6))?),
                v => Value::Double(v.to_f64()?.abs()),
            },
            "int" => match need(args, 0)? {
                v @ (Value::Int(_) | Value::Long(_)) => v.clone(),
                v => Value::Double(v.to_f64()?.floor()),
            },
            "fix" => match need(args, 0)? {
                v @ (Value::Int(_) | Value::Long(_)) => v.clone(),
                v => Value::Double(v.to_f64()?.trunc()),
            },
            "round" => {
                let x = n(args, 0)?;
                let d = arg(args, 1).map(Value::to_i64).transpose()?.unwrap_or(0);
                if !(0..=22).contains(&d) {
                    return Err(VbaError::new(5));
                }
                let f = 10f64.powi(d as i32);
                Value::Double(round_half_even(x * f) / f)
            }
            "sqr" => {
                let x = n(args, 0)?;
                if x < 0.0 {
                    return Err(VbaError::new(5));
                }
                Value::Double(x.sqrt())
            }
            "sgn" => Value::Int(match n(args, 0)? {
                x if x > 0.0 => 1,
                x if x < 0.0 => -1,
                _ => 0,
            }),
            "exp" => Value::Double(n(args, 0)?.exp()),
            "log" => {
                let x = n(args, 0)?;
                if x <= 0.0 {
                    return Err(VbaError::new(5));
                }
                Value::Double(x.ln())
            }
            "sin" => Value::Double(n(args, 0)?.sin()),
            "cos" => Value::Double(n(args, 0)?.cos()),
            "tan" => Value::Double(n(args, 0)?.tan()),
            "atn" => Value::Double(n(args, 0)?.atan()),
            "rnd" => {
                *rnd ^= *rnd << 13;
                *rnd ^= *rnd >> 7;
                *rnd ^= *rnd << 17;
                Value::Double((*rnd >> 11) as f64 / (1u64 << 53) as f64)
            }
            "hex" | "hex$" => Value::str(format!("{:X}", need(args, 0)?.to_i32()?)),
            "oct" | "oct$" => Value::str(format!("{:o}", need(args, 0)?.to_i32()?)),
            "now" => Value::Date(now()),
            "date" | "date$" => Value::Date(now().floor()),
            "time" | "time$" => Value::Date(now().fract()),
            "timer" => Value::Double(now().fract() * 86_400.0),
            "year" => Value::Int(civil(n(args, 0)?).0 as i16),
            "month" => Value::Int(civil(n(args, 0)?).1 as i16),
            "day" => Value::Int(civil(n(args, 0)?).2 as i16),
            "hour" => Value::Int(((n(args, 0)?.fract() * 24.0 + 1e-9).floor()) as i16),
            "minute" => Value::Int(((n(args, 0)?.fract() * 1440.0 + 1e-9).floor() as i64 % 60) as i16),
            "second" => Value::Int(((n(args, 0)?.fract() * 86_400.0).round() as i64 % 60) as i16),
            "weekday" => Value::Int(((n(args, 0)?.floor() as i64 - 1).rem_euclid(7) + 1) as i16),
            "dateserial" => Value::Date(date_serial(need(args, 0)?.to_i64()?, need(args, 1)?.to_i64()?, need(args, 2)?.to_i64()?)),
            "timeserial" => Value::Date((n(args, 0)? * 3600.0 + n(args, 1)? * 60.0 + n(args, 2)?) / 86_400.0),
            "ubound" | "lbound" => {
                let Value::Array(a) = need(args, 0)? else { return Err(VbaError::new(13)) };
                let dim = arg(args, 1).map(Value::to_i64).transpose()?.unwrap_or(1);
                let a = a.borrow();
                let (lo, hi) = *a.bounds.get(usize::try_from(dim - 1).map_err(|_| VbaError::new(9))?).ok_or(VbaError::new(9))?;
                Value::Long(if name.eq_ignore_ascii_case("ubound") { hi } else { lo } as i32)
            }
            "array" => array_of(args.to_vec(), base),
            "iif" => {
                if need(args, 0)?.to_bool()? {
                    need(args, 1)?.clone()
                } else {
                    need(args, 2)?.clone()
                }
            }
            "choose" => {
                let k = need(args, 0)?.to_i64()?;
                usize::try_from(k).ok().and_then(|k| args.get(k)).cloned().unwrap_or(Value::Null)
            }
            "typename" => Value::str(need(args, 0)?.type_name()),
            "vartype" => Value::Int(match need(args, 0)? {
                Value::Empty | Value::Missing => 0,
                Value::Null => 1,
                Value::Int(_) => 2,
                Value::Long(_) => 3,
                Value::Double(_) => 5,
                Value::Date(_) => 7,
                Value::Str(_) => 8,
                Value::Object(_) | Value::Nothing => 9,
                Value::Bool(_) => 11,
                Value::Array(_) => 8204,
            }),
            "strcomp" => {
                let insensitive = match arg(args, 2) {
                    Some(c) => c.to_i32()? == 1,
                    None => text,
                };
                Value::Int(match ops::str_cmp(&s(args, 0)?, &s(args, 1)?, insensitive) {
                    std::cmp::Ordering::Less => -1,
                    std::cmp::Ordering::Equal => 0,
                    std::cmp::Ordering::Greater => 1,
                })
            }
            "rgb" => {
                let c = |i: usize| -> R<i32> { Ok(need(args, i)?.to_i32()?.clamp(0, 255)) };
                Value::Long(c(0)? + c(1)? * 256 + c(2)? * 65_536)
            }
            "doevents" => Value::Int(0),
            "environ" | "environ$" => Value::str(""),
            "nz" => need(args, 0)?.clone(),
            "createobject" | "getobject" => return Err(VbaError::with(429, "ActiveX component can't create object (COM isn't available to macros)")),
            "shell" => return Err(VbaError::with(5, "Shell isn't available to macros")),
            "inputbox" => return Err(VbaError::with(5, "InputBox isn't supported yet")),
            _ => return Ok(None),
        }))
    })();
    r.transpose()
}

/// `2026-10-14`, `10/14/2026`, with an optional `hh:mm[:ss]`.
fn parse_date(s: &str) -> Option<f64> {
    let s = s.trim();
    let (d, t) = s.split_once(' ').map(|(a, b)| (a, Some(b))).unwrap_or((s, None));
    let nums: Vec<i64> = d.split(['-', '/', '.']).map(|x| x.parse().ok()).collect::<Option<_>>()?;
    let (y, m, dd) = match nums.as_slice() {
        [a, b, c] if *a > 31 => (*a, *b, *c),
        [a, b, c] => (*c, *a, *b),
        _ => return None,
    };
    if !(1..=12).contains(&m) || !(1..=31).contains(&dd) {
        return None;
    }
    let mut v = date_serial(y, m, dd);
    if let Some(t) = t {
        let parts: Vec<f64> = t.split(':').map(|x| x.trim().parse().ok()).collect::<Option<_>>()?;
        let secs = parts.first().copied().unwrap_or(0.0) * 3600.0 + parts.get(1).copied().unwrap_or(0.0) * 60.0 + parts.get(2).copied().unwrap_or(0.0);
        v += secs / 86_400.0;
    }
    Some(v)
}

/// The value `n` as an integer-typed value when it is one (for `UBound` and friends' callers).
pub fn int(n: i64) -> Value {
    int_value(n as f64)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn c(name: &str, args: &[Value]) -> Value {
        let mut seed = 1;
        call(name, args, 0, false, &mut seed).unwrap().unwrap()
    }

    #[test]
    fn strings() {
        let st = |x: &str| Value::str(x);
        assert_eq!(c("Left", &[st("Hello"), Value::Int(2)]), st("He"));
        assert_eq!(c("Right", &[st("Hello"), Value::Int(9)]), st("Hello"));
        assert_eq!(c("Mid", &[st("Hello"), Value::Int(2), Value::Int(3)]), st("ell"));
        assert_eq!(c("InStr", &[st("Hello"), st("l")]), Value::Long(3));
        assert_eq!(c("InStr", &[Value::Int(4), st("Hello"), st("l")]), Value::Long(4));
        assert_eq!(c("InStr", &[st("Hello"), st("z")]), Value::Long(0));
        assert_eq!(c("Replace", &[st("a-b-c"), st("-"), st("+")]), st("a+b+c"));
        assert_eq!(c("Trim", &[st("  x  ")]), st("x"));
        assert_eq!(c("Join", &[c("Split", &[st("a,b,c"), st(",")]), st(";")]), st("a;b;c"));
        assert_eq!(c("Format", &[Value::Double(1234.567), st("#,##0.00")]), st("1,234.57"));
        assert_eq!(c("Format", &[Value::Double(0.256), st("0.0%")]), st("25.6%"));
        assert_eq!(c("Format", &[Value::Double(5.0), st("000")]), st("005"));
        assert_eq!(c("Val", &[st(" 12 abc")]), Value::Double(12.0));
        assert_eq!(c("Str", &[Value::Int(5)]), st(" 5"));
        assert_eq!(c("Chr", &[Value::Int(65)]), st("A"));
        assert_eq!(c("Asc", &[st("€")]), Value::Int(128));
        assert_eq!(c("UCase", &[st("straße")]), st("STRASSE"));
    }

    #[test]
    fn numbers_and_dates() {
        assert_eq!(c("Round", &[Value::Double(2.5)]), Value::Double(2.0));
        assert_eq!(c("Round", &[Value::Double(1.2345), Value::Int(2)]), Value::Double(1.23));
        assert_eq!(c("Int", &[Value::Double(-2.5)]), Value::Double(-3.0));
        assert_eq!(c("Fix", &[Value::Double(-2.5)]), Value::Double(-2.0));
        assert_eq!(c("RGB", &[Value::Int(255), Value::Int(0), Value::Int(0)]), Value::Long(255));
        assert_eq!(date_serial(2026, 10, 14), 46309.0);
        assert_eq!(date_serial(1900, 1, 1), 2.0);
        assert_eq!(civil(46309.0), (2026, 10, 14));
        assert_eq!(c("Year", &[Value::Date(46309.0)]), Value::Int(2026));
        assert_eq!(c("CDate", &[Value::str("2026-10-14")]), Value::Date(46309.0));
        assert_eq!(c("IsNumeric", &[Value::str("3.5")]), Value::Bool(true));
        assert_eq!(c("IsNumeric", &[Value::str("x")]), Value::Bool(false));
        let mut seed = 1;
        assert!(call("Sqr", &[Value::Int(-1)], 0, false, &mut seed).unwrap().is_err());
        assert!(call("NoSuchFunction", &[], 0, false, &mut seed).is_none());
        assert_eq!(constant("vbCrLf"), Some(Value::str("\r\n")));
        assert_eq!(constant("xlUp"), Some(Value::Long(-4162)));
    }
}
