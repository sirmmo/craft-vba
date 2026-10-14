//! Excel's object model for macros (`Application`, `Workbook`, `Worksheets`, `Worksheet`,
//! `Range`, `Font`, `Interior`), over the host's `host/sync`: each object is a path from the
//! workbook, as the Office add-ins' bridge reads them (see GridCraft's `office.sync`).
//!
//! Reads are cached by cell until a write could change them (writes don't change other plain
//! values, only formulas), and the first read of a sheet loads its used range in one request.
//! Writes are queued and sent together before the next read, at most 500 at a time, and when the
//! macro ends.

use std::collections::HashMap;
use std::rc::Rc;

use serde_json::{Value as Json, json};

use crate::host::{Host, HostError};
use crate::vba::value::{Array, R, Value, VbaError};

pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u32 = 16_384;
/// Most cells one value read or write may cover.
const MAX_CELLS: u64 = 1_000_000;
const MAX_PENDING: usize = 500;

/// A rectangle of cells on a sheet (0-based, inclusive). `sheet` is the [`Excel::sheets`] id.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct XlRange {
    pub sheet: usize,
    pub r1: u32,
    pub c1: u32,
    pub r2: u32,
    pub c2: u32,
}

impl XlRange {
    pub fn rows(&self) -> u32 {
        self.r2 - self.r1 + 1
    }
    pub fn cols(&self) -> u32 {
        self.c2 - self.c1 + 1
    }
    pub fn count(&self) -> u64 {
        u64::from(self.rows()) * u64::from(self.cols())
    }
    fn a1(&self) -> String {
        if self.r1 == self.r2 && self.c1 == self.c2 {
            cell_a1(self.r1, self.c1)
        } else {
            format!("{}:{}", cell_a1(self.r1, self.c1), cell_a1(self.r2, self.c2))
        }
    }
    /// `$A$1` or `$A$1:$B$2`, as `Address` writes it.
    fn absolute(&self) -> String {
        let one = |r: u32, c: u32| format!("${}${}", col_name(c), r + 1);
        if self.r1 == self.r2 && self.c1 == self.c2 { one(self.r1, self.c1) } else { format!("{}:{}", one(self.r1, self.c1), one(self.r2, self.c2)) }
    }
    fn cell(&self, r: u32, c: u32) -> XlRange {
        XlRange { sheet: self.sheet, r1: r, c1: c, r2: r, c2: c }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum XlObj {
    Application,
    Workbook,
    Worksheets,
    Worksheet(usize),
    Range(XlRange),
    Font(XlRange),
    Interior(XlRange),
    Rows(XlRange),
    Columns(XlRange),
    WorksheetFunction,
    /// The ribbon button a macro was called from (`Sub Go(control As IRibbonControl)`).
    RibbonControl {
        id: Rc<str>,
        tag: Rc<str>,
    },
}

impl XlObj {
    pub fn type_name(&self) -> &'static str {
        match self {
            XlObj::Application => "Application",
            XlObj::Workbook => "Workbook",
            XlObj::Worksheets => "Sheets",
            XlObj::Worksheet(_) => "Worksheet",
            XlObj::Range(_) | XlObj::Rows(_) | XlObj::Columns(_) => "Range",
            XlObj::Font(_) => "Font",
            XlObj::Interior(_) => "Interior",
            XlObj::WorksheetFunction => "WorksheetFunction",
            XlObj::RibbonControl { .. } => "IRibbonControl",
        }
    }
}

pub fn col_name(c: u32) -> String {
    let mut n = c + 1;
    let mut s = Vec::new();
    while n > 0 {
        let r = (n - 1) % 26;
        s.push(b'A' + r as u8);
        n = (n - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap_or_default()
}

fn cell_a1(r: u32, c: u32) -> String {
    format!("{}{}", col_name(c), r + 1)
}

/// `B2`, `$B$2` → (row, col).
fn parse_cell(s: &str) -> Option<(u32, u32)> {
    let s = s.replace('$', "");
    let letters: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let digits = s.get(letters.len()..)?;
    if letters.is_empty() || digits.is_empty() || letters.len() > 3 {
        return None;
    }
    let col = letters.to_ascii_uppercase().bytes().fold(0u32, |a, b| a * 26 + u32::from(b - b'A' + 1)) - 1;
    let row: u32 = digits.parse().ok()?;
    ((1..=MAX_ROWS).contains(&row) && col < MAX_COLS).then(|| (row - 1, col))
}

/// `A1`, `A1:B2`, `A:C`, `2:5` (no sheet) → the rectangle.
pub fn parse_a1(s: &str) -> Option<(u32, u32, u32, u32)> {
    let s = s.trim();
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    if let (Some((r1, c1)), Some((r2, c2))) = (parse_cell(a), parse_cell(b)) {
        return Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)));
    }
    // Whole columns or rows.
    let col = |t: &str| -> Option<u32> {
        let t = t.replace('$', "");
        (!t.is_empty() && t.len() <= 3 && t.chars().all(|c| c.is_ascii_alphabetic()))
            .then(|| t.to_ascii_uppercase().bytes().fold(0u32, |a, b| a * 26 + u32::from(b - b'A' + 1)) - 1)
    };
    let row = |t: &str| -> Option<u32> { t.replace('$', "").parse::<u32>().ok().filter(|r| *r >= 1 && *r <= MAX_ROWS).map(|r| r - 1) };
    if let (Some(c1), Some(c2)) = (col(a), col(b)) {
        return Some((0, c1.min(c2), MAX_ROWS - 1, c1.max(c2)));
    }
    if let (Some(r1), Some(r2)) = (row(a), row(b)) {
        return Some((r1.min(r2), 0, r1.max(r2), MAX_COLS - 1));
    }
    None
}

/// `Sheet1!A1:B2`, `'My Sheet'!A1` → (sheet, rectangle).
fn parse_address(s: &str) -> Option<(String, (u32, u32, u32, u32))> {
    let (sheet, a1) = s.rsplit_once('!')?;
    let sheet = sheet.strip_prefix('\'').and_then(|x| x.strip_suffix('\'')).map(|x| x.replace("''", "'")).unwrap_or_else(|| sheet.to_string());
    Some((sheet, parse_a1(a1)?))
}

/// `#RRGGBB` → VBA's colour number (`R + G*256 + B*65536`).
pub(crate) fn color_from_hex(h: &str) -> Option<i32> {
    let h = h.strip_prefix('#')?;
    let v = u32::from_str_radix(h, 16).ok()?;
    let (r, g, b) = ((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
    i32::try_from(r | (g << 8) | (b << 16)).ok()
}

pub(crate) fn color_to_hex(n: i32) -> String {
    let n = n.max(0) as u32;
    format!("#{:02X}{:02X}{:02X}", n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF)
}

/// A cell's value in the bridge's JSON.
pub fn to_value(j: &Json) -> Value {
    match j {
        Json::Null => Value::Empty,
        Json::Bool(b) => Value::Bool(*b),
        Json::Number(n) => Value::Double(n.as_f64().unwrap_or(0.0)),
        Json::String(s) => Value::str(s.as_str()),
        Json::Object(o) if o.contains_key("error") => Value::str(o.get("error").and_then(Json::as_str).unwrap_or("#VALUE!")),
        other => Value::str(other.to_string()),
    }
}

pub fn to_json(v: &Value) -> R<Json> {
    Ok(match v {
        Value::Empty | Value::Missing | Value::Null => Json::Null,
        Value::Bool(b) => json!(b),
        Value::Int(i) => json!(i),
        Value::Long(i) => json!(i),
        Value::Double(d) | Value::Date(d) => {
            if d.is_finite() {
                json!(d)
            } else {
                return Err(VbaError::new(6));
            }
        }
        Value::Str(s) => json!(s.as_ref()),
        _ => return Err(VbaError::new(13)),
    })
}

pub(crate) fn host_error(e: HostError) -> VbaError {
    let code = e.data.get("code").and_then(Json::as_str).unwrap_or("");
    let number = match code {
        "ItemNotFound" => 9,
        "NotImplemented" => 438,
        _ => 1004,
    };
    VbaError::with(number, e.message)
}

#[derive(Clone, Copy, Debug)]
struct Cached {
    formula: bool,
}

/// The macro's view of the workbook.
#[derive(Default)]
pub struct Excel {
    /// Sheets the macro has seen, by id (a renamed sheet keeps its id).
    pub sheets: Vec<String>,
    pending: Vec<Json>,
    cache: HashMap<(usize, u32, u32), (Json, Cached)>,
    /// Sheets whose used range was loaded (cells outside it are empty), until something is written.
    used: HashMap<usize, Option<(u32, u32, u32, u32)>>,
    /// Requests sent (for tests and tuning).
    pub requests: usize,
    /// Something was written: reads no longer load whole used ranges.
    wrote: bool,
    /// The active sheet, until something may change it.
    active: Option<usize>,
}

type Args<'a> = &'a [Value];

impl Excel {
    pub fn sheet_id(&mut self, name: &str) -> usize {
        match self.sheets.iter().position(|n| n.eq_ignore_ascii_case(name)) {
            Some(i) => i,
            None => {
                self.sheets.push(name.to_string());
                self.sheets.len() - 1
            }
        }
    }

    fn sheet_path(&self, id: usize) -> Json {
        let name = self.sheets.get(id).cloned().unwrap_or_default();
        json!([{"p": "workbook"}, {"p": "worksheets"}, {"m": "getItem", "a": [name]}])
    }

    fn range_path(&self, r: &XlRange) -> Json {
        let mut p = self.sheet_path(r.sheet);
        if let Some(a) = p.as_array_mut() {
            a.push(json!({"m": "getRange", "a": [r.a1()]}));
        }
        p
    }

    fn sub_path(&self, r: &XlRange, steps: &[&str]) -> Json {
        let mut p = self.range_path(r);
        if let Some(a) = p.as_array_mut() {
            for s in steps {
                a.push(json!({"p": s}));
            }
        }
        p
    }

    /// Sends the queued writes and `actions`; the results of `actions`.
    fn sync(&mut self, host: &mut dyn Host, actions: Vec<Json>) -> R<Vec<Json>> {
        let queued = self.pending.len();
        let mut all = std::mem::take(&mut self.pending);
        all.extend(actions);
        if all.is_empty() {
            return Ok(Vec::new());
        }
        self.requests += 1;
        let out = host.call("host/sync", json!({"actions": all})).map_err(host_error)?;
        let results = out.get("results").and_then(Json::as_array).cloned().unwrap_or_default();
        Ok(results.into_iter().skip(queued).collect())
    }

    /// Sends the queued writes (the macro ended, or something outside the cache is next).
    pub fn flush(&mut self, host: &mut dyn Host) -> R<()> {
        if !self.pending.is_empty() {
            self.sync(host, Vec::new())?;
        }
        Ok(())
    }

    fn write(&mut self, host: &mut dyn Host, action: Json) -> R<()> {
        self.pending.push(action);
        self.wrote = true;
        // Formula cells may change with any write; the used range may grow.
        self.cache.retain(|_, (_, c)| !c.formula);
        self.used.clear();
        if self.pending.len() >= MAX_PENDING {
            self.flush(host)?;
        }
        Ok(())
    }

    fn load(&mut self, host: &mut dyn Host, path: Json, props: &[&str]) -> R<Json> {
        let mut r = self.sync(host, vec![json!({"op": "load", "path": path, "props": props})])?;
        Ok(r.pop().unwrap_or(Json::Null))
    }

    fn call_host(&mut self, host: &mut dyn Host, path: Json, method: &str, args: Vec<Json>) -> R<Json> {
        let mut r = self.sync(host, vec![json!({"op": "call", "path": path, "method": method, "args": args})])?;
        Ok(r.pop().unwrap_or(Json::Null))
    }

    /// Something changed the sheets (added, deleted, renamed, activated): nothing cached holds.
    fn forget(&mut self) {
        self.cache.clear();
        self.used.clear();
        self.active = None;
    }

    fn cache_rows(&mut self, sheet: usize, r1: u32, c1: u32, values: &Json, formulas: &Json) {
        for (i, row) in values.as_array().into_iter().flatten().enumerate() {
            for (j, v) in row.as_array().into_iter().flatten().enumerate() {
                let f = formulas.get(i).and_then(|r| r.get(j)).and_then(Json::as_str).is_some_and(|s| s.starts_with('='));
                self.cache.insert((sheet, r1 + i as u32, c1 + j as u32), (v.clone(), Cached { formula: f }));
            }
        }
    }

    fn cached(&self, sheet: usize, r: u32, c: u32) -> Option<Json> {
        if let Some((v, _)) = self.cache.get(&(sheet, r, c)) {
            return Some(v.clone());
        }
        match self.used.get(&sheet) {
            Some(Some((r1, c1, r2, c2))) if r < *r1 || r > *r2 || c < *c1 || c > *c2 => Some(Json::Null),
            Some(None) => Some(Json::Null),
            _ => None,
        }
    }

    /// Loads a sheet's used range into the cache (one request for a macro's reads).
    fn ensure_used(&mut self, host: &mut dyn Host, sheet: usize) -> R<Option<(u32, u32, u32, u32)>> {
        if let Some(u) = self.used.get(&sheet) {
            return Ok(*u);
        }
        let mut path = self.sheet_path(sheet);
        if let Some(a) = path.as_array_mut() {
            a.push(json!({"m": "getUsedRange"}));
        }
        let out = self.load(host, path, &["address", "values", "formulas"])?;
        let rect = out.get("address").and_then(Json::as_str).and_then(parse_address).map(|(_, r)| r);
        let values = out.get("values").cloned().unwrap_or(Json::Null);
        let empty = values.as_array().is_none_or(|rows| rows.iter().all(|r| r.as_array().is_none_or(|c| c.iter().all(Json::is_null))));
        let used = if empty { None } else { rect };
        if let Some((r1, c1, _, _)) = used {
            let formulas = out.get("formulas").cloned().unwrap_or(Json::Null);
            self.cache_rows(sheet, r1, c1, &values, &formulas);
        }
        self.used.insert(sheet, used);
        Ok(used)
    }

    fn values(&mut self, host: &mut dyn Host, r: &XlRange) -> R<Vec<Vec<Json>>> {
        if r.count() > MAX_CELLS {
            return Err(VbaError::with(7, "Out of memory (more than a million cells)"));
        }
        let grid = |me: &Excel| -> Option<Vec<Vec<Json>>> {
            (r.r1..=r.r2).map(|row| (r.c1..=r.c2).map(|c| me.cached(r.sheet, row, c)).collect::<Option<Vec<_>>>()).collect()
        };
        if let Some(g) = grid(self) {
            return Ok(g);
        }
        // A sheet's first read loads its used range; it's small enough and saves a request per cell.
        if !self.wrote && !self.used.contains_key(&r.sheet) {
            self.ensure_used(host, r.sheet)?;
            if let Some(g) = grid(self) {
                return Ok(g);
            }
        }
        let out = self.load(host, self.range_path(r), &["values", "formulas"])?;
        let values = out.get("values").cloned().unwrap_or(Json::Null);
        let formulas = out.get("formulas").cloned().unwrap_or(Json::Null);
        self.cache_rows(r.sheet, r.r1, r.c1, &values, &formulas);
        Ok(values.as_array().cloned().unwrap_or_default().into_iter().map(|row| row.as_array().cloned().unwrap_or_default()).collect())
    }

    /// Rows of `prop` (`text`, `formulas`, `numberFormat`) for a range.
    fn rows_of(&mut self, host: &mut dyn Host, r: &XlRange, prop: &str) -> R<Vec<Vec<Json>>> {
        if r.count() > MAX_CELLS {
            return Err(VbaError::with(7, "Out of memory (more than a million cells)"));
        }
        let out = self.load(host, self.range_path(r), &[prop])?;
        Ok(out.get(prop).and_then(Json::as_array).cloned().unwrap_or_default().into_iter().map(|row| row.as_array().cloned().unwrap_or_default()).collect())
    }

    /// One value for one cell, a 1-based 2-D array for several.
    fn grid_value(rows: Vec<Vec<Json>>, f: impl Fn(&Json) -> Value) -> R<Value> {
        let h = rows.len();
        let w = rows.first().map(Vec::len).unwrap_or(0);
        if h == 1 && w == 1 {
            return Ok(rows.first().and_then(|r| r.first()).map(&f).unwrap_or(Value::Empty));
        }
        let mut a = Array::new(vec![(1, h as i64), (1, w as i64)])?;
        for (i, row) in rows.iter().enumerate() {
            for (j, v) in row.iter().enumerate() {
                a.set(&[i as i64 + 1, j as i64 + 1], f(v))?;
            }
        }
        Ok(Value::Array(std::rc::Rc::new(std::cell::RefCell::new(a))))
    }

    /// The same value for every cell, else Null (Excel's `NumberFormat`, `Text` on several cells).
    fn uniform(rows: Vec<Vec<Json>>, f: impl Fn(&Json) -> Value) -> Value {
        let mut it = rows.iter().flatten();
        let Some(first) = it.next() else { return Value::Null };
        if it.all(|v| v == first) { f(first) } else { Value::Null }
    }

    /// `value` as rows the size of `r` (Excel fills cells beyond a smaller array with #N/A).
    fn rows_for(r: &XlRange, value: &Value) -> R<Json> {
        match value {
            Value::Array(a) => {
                let a = a.borrow();
                let get = |i: u32, j: u32| -> R<Json> {
                    let v = match a.bounds.as_slice() {
                        [(lo, hi)] => {
                            // A 1-D array fills a row, and repeats down the rows.
                            let k = lo + i64::from(j);
                            if k <= *hi { a.get(&[k])? } else { Value::str("#N/A") }
                        }
                        [(lo1, hi1), (lo2, hi2)] => {
                            let (x, y) = (lo1 + i64::from(i), lo2 + i64::from(j));
                            if x <= *hi1 && y <= *hi2 { a.get(&[x, y])? } else { Value::str("#N/A") }
                        }
                        _ => return Err(VbaError::new(13)),
                    };
                    to_json(&v)
                };
                let mut rows = Vec::new();
                for i in 0..r.rows() {
                    let mut row = Vec::new();
                    for j in 0..r.cols() {
                        row.push(get(i, j)?);
                    }
                    rows.push(Json::Array(row));
                }
                Ok(Json::Array(rows))
            }
            other => to_json(other),
        }
    }

    fn set_values(&mut self, host: &mut dyn Host, r: &XlRange, prop: &str, value: &Value) -> R<()> {
        if r.count() > MAX_CELLS {
            return Err(VbaError::with(7, "Out of memory (more than a million cells)"));
        }
        let rows = Excel::rows_for(r, value)?;
        let action = json!({"op": "set", "path": self.range_path(r), "prop": prop, "value": rows});
        self.write(host, action)?;
        // What was written is known, unless it's text the app may read as a number or a formula.
        for row in r.r1..=r.r2 {
            for c in r.c1..=r.c2 {
                let v = match &rows {
                    Json::Array(rs) => rs.get((row - r.r1) as usize).and_then(|x| x.get((c - r.c1) as usize)).cloned().unwrap_or(Json::Null),
                    v => v.clone(),
                };
                if prop == "values" && !v.is_string() {
                    self.cache.insert((r.sheet, row, c), (v, Cached { formula: false }));
                } else {
                    self.cache.remove(&(r.sheet, row, c));
                }
            }
        }
        Ok(())
    }

    /// The range an object stands for (a cell's or a sheet's).
    fn range_of(&self, o: &XlObj) -> Option<XlRange> {
        match o {
            XlObj::Range(r) | XlObj::Rows(r) | XlObj::Columns(r) => Some(*r),
            _ => None,
        }
    }

    fn active_sheet(&mut self, host: &mut dyn Host) -> R<usize> {
        if let Some(a) = self.active {
            return Ok(a);
        }
        let out = self.load(host, json!([{"p": "workbook"}, {"p": "worksheets"}, {"m": "getActiveWorksheet"}]), &["name"])?;
        let name = out.get("name").and_then(Json::as_str).ok_or(VbaError::new(91))?.to_string();
        let id = self.sheet_id(&name);
        self.active = Some(id);
        Ok(id)
    }

    fn range_at(&mut self, host: &mut dyn Host, path: Json) -> R<XlRange> {
        let out = self.load(host, path, &["address"])?;
        let (sheet, (r1, c1, r2, c2)) = out.get("address").and_then(Json::as_str).and_then(parse_address).ok_or(VbaError::new(1004))?;
        Ok(XlRange { sheet: self.sheet_id(&sheet), r1, c1, r2, c2 })
    }

    fn sheet_names(&mut self, host: &mut dyn Host) -> R<Vec<String>> {
        let out = self.load(host, json!([{"p": "workbook"}, {"p": "worksheets"}]), &["items"])?;
        Ok(out.get("items").and_then(Json::as_array).into_iter().flatten().filter_map(|i| i.get("name").and_then(Json::as_str)).map(str::to_string).collect())
    }

    /// `Range("A1")`, `Range("A1", "B2")`, `Range(Cells(1, 1), Cells(2, 2))` on `sheet`.
    fn range_arg(&mut self, host: &mut dyn Host, sheet: usize, args: Args) -> R<XlRange> {
        let one = |me: &mut Excel, v: &Value| -> R<XlRange> {
            match v {
                Value::Object(crate::vba::value::Obj::Xl(o)) => me.range_of(o).ok_or(VbaError::new(13)),
                v => {
                    let s = v.to_str()?;
                    if let Some((name, (r1, c1, r2, c2))) = parse_address(&s) {
                        return Ok(XlRange { sheet: me.sheet_id(&name), r1, c1, r2, c2 });
                    }
                    let (r1, c1, r2, c2) = parse_a1(&s).ok_or_else(|| VbaError::with(1004, format!("Method 'Range' of object '_Global' failed (`{s}`)")))?;
                    Ok(XlRange { sheet, r1, c1, r2, c2 })
                }
            }
        };
        let _ = host;
        match args {
            [a] => one(self, a),
            [a, b] => {
                let (x, y) = (one(self, a)?, one(self, b)?);
                Ok(XlRange { sheet: x.sheet, r1: x.r1.min(y.r1), c1: x.c1.min(y.c1), r2: x.r2.max(y.r2), c2: x.c2.max(y.c2) })
            }
            _ => Err(VbaError::new(450)),
        }
    }

    /// `Cells(r, c)` relative to `base` (1-based); `Cells` alone is all of it.
    fn cells(base: XlRange, args: Args) -> R<XlObj> {
        match args {
            [] => Ok(XlObj::Range(base)),
            [i] => {
                // Cells(n): the nth cell, row by row.
                let n = i.to_i64()?;
                if n < 1 {
                    return Err(VbaError::new(1004));
                }
                let w = i64::from(base.cols());
                let (r, c) = ((n - 1) / w, (n - 1) % w);
                Excel::offset(base.cell(base.r1, base.c1), r, c).map(XlObj::Range)
            }
            [r, c] => {
                let col = match c {
                    Value::Str(s) => i64::from(parse_a1(&format!("{s}1")).map(|x| x.1).ok_or(VbaError::new(13))?) + 1,
                    v => v.to_i64()?,
                };
                Excel::offset(base.cell(base.r1, base.c1), r.to_i64()? - 1, col - 1).map(XlObj::Range)
            }
            _ => Err(VbaError::new(450)),
        }
    }

    fn offset(r: XlRange, dr: i64, dc: i64) -> R<XlRange> {
        let shift = |v: u32, d: i64, max: u32| -> R<u32> {
            let n = i64::from(v) + d;
            if n < 0 || n >= i64::from(max) { Err(VbaError::new(1004)) } else { Ok(n as u32) }
        };
        Ok(XlRange {
            sheet: r.sheet,
            r1: shift(r.r1, dr, MAX_ROWS)?,
            c1: shift(r.c1, dc, MAX_COLS)?,
            r2: shift(r.r2, dr, MAX_ROWS)?,
            c2: shift(r.c2, dc, MAX_COLS)?,
        })
    }

    fn whole_sheet(sheet: usize) -> XlRange {
        XlRange { sheet, r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }
    }

    /// `Range.End(direction)`: the edge of the region, as Ctrl+arrow moves.
    fn end(&mut self, host: &mut dyn Host, r: XlRange, dir: i64) -> R<XlRange> {
        self.flush(host)?;
        self.used.remove(&r.sheet);
        let used = self.ensure_used(host, r.sheet)?;
        let filled = |me: &Excel, row: u32, col: u32| -> bool { me.cached(r.sheet, row, col).is_some_and(|v| !v.is_null() && v != json!("")) };
        let (dr, dc): (i64, i64) = match dir {
            -4162 => (-1, 0),
            -4121 => (1, 0),
            -4159 => (0, -1),
            -4161 => (0, 1),
            _ => return Err(VbaError::new(5)),
        };
        let (mut row, mut col) = (i64::from(r.r1), i64::from(r.c1));
        let inside = |row: i64, col: i64| row >= 0 && col >= 0 && row < i64::from(MAX_ROWS) && col < i64::from(MAX_COLS);
        // Beyond the used range everything is empty: jump straight to its edge.
        let (ur1, uc1, ur2, uc2) = used.unwrap_or((0, 0, 0, 0));
        let next_filled = |me: &Excel, row: i64, col: i64| inside(row, col) && filled(me, row as u32, col as u32);
        let here = next_filled(self, row, col);
        let ahead = next_filled(self, row + dr, col + dc);
        let mut steps = 0u64;
        if here && ahead {
            while next_filled(self, row + dr, col + dc) && steps < u64::from(MAX_ROWS) {
                row += dr;
                col += dc;
                steps += 1;
            }
        } else {
            // Skip the empty cells to the next filled one, or to the sheet's edge.
            loop {
                let (nr, nc) = (row + dr, col + dc);
                if !inside(nr, nc) {
                    break;
                }
                row = nr;
                col = nc;
                if filled(self, row as u32, col as u32) {
                    break;
                }
                // Outside the used range nothing is filled further along: go to the edge.
                let outside = used.is_none()
                    || row < i64::from(ur1) && dr < 0
                    || row > i64::from(ur2) && dr > 0
                    || col < i64::from(uc1) && dc < 0
                    || col > i64::from(uc2) && dc > 0;
                if outside {
                    row = if dr < 0 {
                        0
                    } else if dr > 0 {
                        i64::from(MAX_ROWS) - 1
                    } else {
                        row
                    };
                    col = if dc < 0 {
                        0
                    } else if dc > 0 {
                        i64::from(MAX_COLS) - 1
                    } else {
                        col
                    };
                    break;
                }
                steps += 1;
                if steps > u64::from(MAX_ROWS) {
                    break;
                }
            }
        }
        Ok(r.cell(row as u32, col as u32))
    }

    fn numbers(&mut self, host: &mut dyn Host, args: Args, out: &mut Vec<f64>, count_text: &mut usize) -> R<()> {
        for a in args {
            match a {
                Value::Object(crate::vba::value::Obj::Xl(o)) => {
                    let r = self.range_of(o).ok_or(VbaError::new(13))?;
                    for v in self.values(host, &r)?.into_iter().flatten() {
                        match v {
                            Json::Number(n) => out.push(n.as_f64().unwrap_or(0.0)),
                            Json::Null => {}
                            _ => *count_text += 1,
                        }
                    }
                }
                Value::Array(arr) => {
                    for v in arr.borrow().data.iter() {
                        if v.is_numeric() && !matches!(v, Value::Empty) {
                            out.push(v.to_f64()?);
                        }
                    }
                }
                Value::Missing | Value::Empty => {}
                v => out.push(v.to_f64()?),
            }
        }
        Ok(())
    }

    fn sheet_count_name(&mut self, host: &mut dyn Host) -> R<String> {
        let names = self.sheet_names(host)?;
        let mut n = names.len() + 1;
        loop {
            let candidate = format!("Sheet{n}");
            if !names.iter().any(|x| x.eq_ignore_ascii_case(&candidate)) {
                return Ok(candidate);
            }
            n += 1;
        }
    }

    /// `object.name` or `object.name(args)`: a property's value or a method's result.
    pub fn member(&mut self, host: &mut dyn Host, o: &XlObj, name: &str, args: Args, named: &[(String, Value)]) -> R<Value> {
        let n = name.to_ascii_lowercase();
        let obj = |x: XlObj| Ok(Value::Object(crate::vba::value::Obj::Xl(x)));
        match (o, n.as_str()) {
            // Application and the globals.
            (XlObj::Application | XlObj::Workbook, "activesheet") => {
                let s = self.active_sheet(host)?;
                obj(XlObj::Worksheet(s))
            }
            (XlObj::Application, "activeworkbook" | "thisworkbook") => obj(XlObj::Workbook),
            (XlObj::Application, "activecell") => {
                let r = self.range_at(host, json!([{"p": "workbook"}, {"m": "getActiveCell"}]))?;
                obj(XlObj::Range(r))
            }
            (XlObj::Application, "selection") => {
                let r = self.range_at(host, json!([{"p": "workbook"}, {"m": "getSelectedRange"}]))?;
                obj(XlObj::Range(r))
            }
            (XlObj::Application | XlObj::Workbook, "worksheets" | "sheets") => {
                if args.is_empty() {
                    obj(XlObj::Worksheets)
                } else {
                    self.member(host, &XlObj::Worksheets, "item", args, named)
                }
            }
            (XlObj::Application, "range" | "cells" | "rows" | "columns") => {
                let s = self.active_sheet(host)?;
                self.member(host, &XlObj::Worksheet(s), name, args, named)
            }
            (XlObj::Application, "worksheetfunction") => obj(XlObj::WorksheetFunction),
            (XlObj::RibbonControl { id, .. }, "id") => Ok(Value::Str(id.clone())),
            (XlObj::RibbonControl { tag, .. }, "tag") => Ok(Value::Str(tag.clone())),
            (XlObj::RibbonControl { .. }, "context") => obj(XlObj::Application),
            (XlObj::Application, "screenupdating" | "displayalerts" | "enableevents" | "interactive") => Ok(Value::Bool(true)),
            (XlObj::Application, "calculation") => Ok(Value::Long(-4105)),
            (XlObj::Application, "statusbar") => Ok(Value::Bool(false)),
            (XlObj::Application, "name") => Ok(Value::str("GridCraft")),
            (XlObj::Application, "version") => Ok(Value::str("16.0")),
            (XlObj::Application, "application") | (_, "application") => obj(XlObj::Application),
            (XlObj::Workbook, "name") => {
                let out = self.load(host, json!([{"p": "workbook"}]), &["name"])?;
                Ok(to_value(out.get("name").unwrap_or(&Json::Null)))
            }
            (XlObj::Workbook, "parent") => obj(XlObj::Application),

            // Worksheets.
            (XlObj::Worksheets, "count") => Ok(Value::Long(self.sheet_names(host)?.len() as i32)),
            (XlObj::Worksheets, "item") => {
                let names = self.sheet_names(host)?;
                let name = match args.first() {
                    Some(Value::Str(s)) => names.iter().find(|n| n.eq_ignore_ascii_case(s)).cloned().ok_or(VbaError::new(9))?,
                    Some(v) => {
                        let i = v.to_i64()?;
                        names.get(usize::try_from(i - 1).map_err(|_| VbaError::new(9))?).cloned().ok_or(VbaError::new(9))?
                    }
                    None => return Err(VbaError::new(449)),
                };
                obj(XlObj::Worksheet(self.sheet_id(&name)))
            }
            (XlObj::Worksheets, "add") => {
                let name = self.sheet_count_name(host)?;
                self.flush(host)?;
                let out = self.call_host(host, json!([{"p": "workbook"}, {"p": "worksheets"}]), "add", vec![json!(name)])?;
                let name = out.get("name").and_then(Json::as_str).map(str::to_string).unwrap_or(name);
                self.forget();
                let id = self.sheet_id(&name);
                // A new sheet becomes the active one, as in Excel.
                self.call_host(host, self.sheet_path(id), "activate", Vec::new())?;
                obj(XlObj::Worksheet(id))
            }

            // A worksheet.
            (XlObj::Worksheet(s), "name") => Ok(Value::str(self.sheets.get(*s).cloned().unwrap_or_default())),
            (XlObj::Worksheet(s), "index") => {
                let out = self.load(host, self.sheet_path(*s), &["position"])?;
                Ok(Value::Long(out.get("position").and_then(Json::as_i64).unwrap_or(0) as i32 + 1))
            }
            (XlObj::Worksheet(s), "range") => {
                let r = self.range_arg(host, *s, args)?;
                obj(XlObj::Range(r))
            }
            (XlObj::Worksheet(s), "cells") => obj(Excel::cells(Excel::whole_sheet(*s), args)?),
            (XlObj::Worksheet(s), "rows") => self
                .member(host, &XlObj::Rows(Excel::whole_sheet(*s)), "item", args, named)
                .or_else(|e| if args.is_empty() { obj(XlObj::Rows(Excel::whole_sheet(*s))) } else { Err(e) }),
            (XlObj::Worksheet(s), "columns") => self
                .member(host, &XlObj::Columns(Excel::whole_sheet(*s)), "item", args, named)
                .or_else(|e| if args.is_empty() { obj(XlObj::Columns(Excel::whole_sheet(*s))) } else { Err(e) }),
            (XlObj::Worksheet(s), "usedrange") => {
                let mut path = self.sheet_path(*s);
                if let Some(a) = path.as_array_mut() {
                    a.push(json!({"m": "getUsedRange"}));
                }
                let r = self.range_at(host, path)?;
                obj(XlObj::Range(r))
            }
            (XlObj::Worksheet(s), "activate" | "select") => {
                self.flush(host)?;
                self.call_host(host, self.sheet_path(*s), "activate", Vec::new())?;
                self.forget();
                Ok(Value::Empty)
            }
            (XlObj::Worksheet(s), "delete") => {
                self.flush(host)?;
                self.call_host(host, self.sheet_path(*s), "delete", Vec::new())?;
                self.forget();
                Ok(Value::Empty)
            }
            (XlObj::Worksheet(_), "parent") => obj(XlObj::Workbook),

            // Rows and Columns.
            (XlObj::Rows(r), "count") => Ok(Value::Long(r.rows() as i32)),
            (XlObj::Columns(r), "count") => Ok(Value::Long(r.cols() as i32)),
            (XlObj::Rows(r), "item") => {
                let [i] = args else { return obj(XlObj::Rows(*r)) };
                let i = i.to_i64()? - 1;
                let row = Excel::offset(XlRange { r2: r.r1, ..*r }, i, 0)?;
                obj(XlObj::Range(row))
            }
            (XlObj::Columns(r), "item") => {
                let [i] = args else { return obj(XlObj::Columns(*r)) };
                let c = match i {
                    Value::Str(s) => i64::from(parse_a1(&format!("{s}1")).map(|x| x.1).ok_or(VbaError::new(13))?) - i64::from(r.c1),
                    v => v.to_i64()? - 1,
                };
                let col = Excel::offset(XlRange { c2: r.c1, ..*r }, 0, c)?;
                obj(XlObj::Range(col))
            }

            // A range.
            (XlObj::Range(r) | XlObj::Rows(r) | XlObj::Columns(r), p) => self.range_member(host, *r, p, args, named),

            // Font and Interior.
            (XlObj::Font(r), p) => {
                let key = match p {
                    "bold" | "italic" | "size" | "name" | "color" | "underline" | "strikethrough" => p,
                    _ => return Err(VbaError::new(438)),
                };
                let host_key = if key == "strikethrough" { "strikethrough" } else { key };
                let out = self.load(host, self.sub_path(r, &["format", "font"]), &[host_key])?;
                let v = out.get(host_key).cloned().unwrap_or(Json::Null);
                Ok(match key {
                    "color" => Value::Long(v.as_str().and_then(color_from_hex).unwrap_or(0)),
                    "underline" => Value::Long(match v.as_str().unwrap_or("None") {
                        "None" => -4142,
                        "Double" => -4119,
                        _ => 2,
                    }),
                    "size" => Value::Double(v.as_f64().unwrap_or(11.0)),
                    _ => to_value(&v),
                })
            }
            (XlObj::Interior(r), "color") => {
                let out = self.load(host, self.sub_path(r, &["format", "fill"]), &["color"])?;
                Ok(Value::Long(out.get("color").and_then(Json::as_str).and_then(color_from_hex).unwrap_or(16_777_215)))
            }
            (XlObj::Interior(_), "colorindex") => Err(VbaError::with(438, "Interior.ColorIndex isn't supported yet (use Interior.Color)")),

            // WorksheetFunction.
            (XlObj::WorksheetFunction, f) => {
                let mut nums = Vec::new();
                let mut text = 0usize;
                self.numbers(host, args, &mut nums, &mut text)?;
                Ok(match f {
                    "sum" => Value::Double(nums.iter().sum()),
                    "max" => Value::Double(nums.iter().copied().fold(None, |m: Option<f64>, x| Some(m.map_or(x, |m| m.max(x)))).unwrap_or(0.0)),
                    "min" => Value::Double(nums.iter().copied().fold(None, |m: Option<f64>, x| Some(m.map_or(x, |m| m.min(x)))).unwrap_or(0.0)),
                    "average" => {
                        if nums.is_empty() {
                            return Err(VbaError::with(1004, "Unable to get the Average property of the WorksheetFunction class"));
                        }
                        Value::Double(nums.iter().sum::<f64>() / nums.len() as f64)
                    }
                    "count" => Value::Double(nums.len() as f64),
                    "counta" => Value::Double((nums.len() + text) as f64),
                    _ => return Err(VbaError::with(438, format!("WorksheetFunction.{name} isn't supported yet"))),
                })
            }
            _ => Err(VbaError::new(438)),
        }
    }

    fn range_member(&mut self, host: &mut dyn Host, r: XlRange, p: &str, args: Args, named: &[(String, Value)]) -> R<Value> {
        let obj = |x: XlObj| Ok(Value::Object(crate::vba::value::Obj::Xl(x)));
        let _ = named;
        match p {
            "value" | "value2" => {
                let rows = self.values(host, &r)?;
                Excel::grid_value(rows, to_value)
            }
            "formula" | "formulalocal" => {
                let rows = self.rows_of(host, &r, "formulas")?;
                Excel::grid_value(rows, |v| Value::str(to_value(v).to_str().unwrap_or_default()))
            }
            "text" => {
                let rows = self.rows_of(host, &r, "text")?;
                Ok(Excel::uniform(rows, to_value))
            }
            "numberformat" | "numberformatlocal" => {
                let rows = self.rows_of(host, &r, "numberFormat")?;
                Ok(Excel::uniform(rows, to_value))
            }
            "address" => Ok(Value::str(r.absolute())),
            "row" => Ok(Value::Long(r.r1 as i32 + 1)),
            "column" => Ok(Value::Long(r.c1 as i32 + 1)),
            "count" | "countlarge" => Ok(Value::Double(r.count() as f64)),
            "rows" => {
                if args.is_empty() {
                    obj(XlObj::Rows(r))
                } else {
                    self.member(host, &XlObj::Rows(r), "item", args, named)
                }
            }
            "columns" => {
                if args.is_empty() {
                    obj(XlObj::Columns(r))
                } else {
                    self.member(host, &XlObj::Columns(r), "item", args, named)
                }
            }
            "cells" | "item" => obj(Excel::cells(r, args)?),
            "offset" => {
                let dr = args.first().filter(|v| !matches!(v, Value::Missing)).map(Value::to_i64).transpose()?.unwrap_or(0);
                let dc = args.get(1).filter(|v| !matches!(v, Value::Missing)).map(Value::to_i64).transpose()?.unwrap_or(0);
                obj(XlObj::Range(Excel::offset(r, dr, dc)?))
            }
            "resize" => {
                let h = args.first().filter(|v| !matches!(v, Value::Missing)).map(Value::to_i64).transpose()?.unwrap_or(i64::from(r.rows()));
                let w = args.get(1).filter(|v| !matches!(v, Value::Missing)).map(Value::to_i64).transpose()?.unwrap_or(i64::from(r.cols()));
                if h < 1 || w < 1 {
                    return Err(VbaError::new(1004));
                }
                let end = Excel::offset(r.cell(r.r1, r.c1), h - 1, w - 1)?;
                obj(XlObj::Range(XlRange { r2: end.r1, c2: end.c1, ..r }))
            }
            "entirerow" => obj(XlObj::Range(XlRange { c1: 0, c2: MAX_COLS - 1, ..r })),
            "entirecolumn" => obj(XlObj::Range(XlRange { r1: 0, r2: MAX_ROWS - 1, ..r })),
            "end" => {
                let dir = args.first().ok_or(VbaError::new(449))?.to_i64()?;
                obj(XlObj::Range(self.end(host, r, dir)?))
            }
            "font" => obj(XlObj::Font(r)),
            "interior" => obj(XlObj::Interior(r)),
            "worksheet" | "parent" => obj(XlObj::Worksheet(r.sheet)),
            "range" => {
                // A range relative to this one: Range("A1") is its top-left cell.
                let inner = self.range_arg(host, r.sheet, args)?;
                obj(XlObj::Range(Excel::offset(inner, i64::from(r.r1), i64::from(r.c1))?))
            }
            "clearcontents" | "clearformats" | "clear" => {
                let apply = match p {
                    "clearcontents" => "Contents",
                    "clearformats" => "Formats",
                    _ => "All",
                };
                let action = json!({"op": "call", "path": self.range_path(&r), "method": "clear", "args": [apply]});
                self.write(host, action)?;
                if apply != "Formats" && r.count() <= MAX_CELLS {
                    for row in r.r1..=r.r2 {
                        for c in r.c1..=r.c2 {
                            self.cache.insert((r.sheet, row, c), (Json::Null, Cached { formula: false }));
                        }
                    }
                }
                Ok(Value::Empty)
            }
            "select" | "activate" => {
                self.flush(host)?;
                self.call_host(host, self.range_path(&r), "select", Vec::new())?;
                self.active = Some(r.sheet);
                Ok(Value::Empty)
            }
            _ => Err(VbaError::with(438, format!("Range.{p} isn't supported yet"))),
        }
    }

    /// `object.name = value` (and `object.name(args) = value`).
    pub fn assign(&mut self, host: &mut dyn Host, o: &XlObj, name: &str, args: Args, value: &Value) -> R<()> {
        let n = name.to_ascii_lowercase();
        match (o, n.as_str()) {
            (XlObj::Application, "screenupdating" | "displayalerts" | "enableevents" | "calculation" | "statusbar" | "interactive" | "cutcopymode") => Ok(()),
            (XlObj::Application, "range" | "cells") => {
                let s = self.active_sheet(host)?;
                self.assign(host, &XlObj::Worksheet(s), name, args, value)
            }
            (XlObj::Worksheet(s), "name") => {
                let new = value.to_str()?;
                self.flush(host)?;
                let action = json!({"op": "set", "path": self.sheet_path(*s), "prop": "name", "value": new});
                self.sync(host, vec![action])?;
                if let Some(slot) = self.sheets.get_mut(*s) {
                    *slot = new;
                }
                self.forget();
                Ok(())
            }
            (XlObj::Worksheet(s), "range" | "cells") => {
                let target = match n.as_str() {
                    "range" => XlObj::Range(self.range_arg(host, *s, args)?),
                    _ => Excel::cells(Excel::whole_sheet(*s), args)?,
                };
                self.assign(host, &target, "value", &[], value)
            }
            (XlObj::Range(r), "value" | "value2") => self.set_values(host, r, "values", value),
            (XlObj::Range(r), "formula" | "formulalocal") => self.set_values(host, r, "formulas", value),
            (XlObj::Range(r), "numberformat" | "numberformatlocal") => {
                let action = json!({"op": "set", "path": self.range_path(r), "prop": "numberFormat", "value": value.to_str()?});
                self.write(host, action)
            }
            (XlObj::Range(r), "cells" | "item") => {
                let target = Excel::cells(*r, args)?;
                self.assign(host, &target, "value", &[], value)
            }
            (XlObj::Font(r), p) if p != "colorindex" => {
                let v = match p {
                    "bold" | "italic" | "strikethrough" => json!(value.to_bool()?),
                    "size" => json!(value.to_f64()?),
                    "name" => json!(value.to_str()?),
                    "color" => json!(color_to_hex(value.to_i32()?)),
                    "underline" => json!(match value {
                        Value::Bool(true) => "Single",
                        Value::Bool(false) => "None",
                        v => match v.to_i32()? {
                            -4142 => "None",
                            -4119 => "Double",
                            _ => "Single",
                        },
                    }),
                    _ => return Err(VbaError::new(438)),
                };
                let action = json!({"op": "set", "path": self.sub_path(r, &["format", "font"]), "prop": p, "value": v});
                self.write(host, action)
            }
            (XlObj::Interior(r), "color") => {
                let action = json!({"op": "set", "path": self.sub_path(r, &["format", "fill"]), "prop": "color", "value": color_to_hex(value.to_i32()?)});
                self.write(host, action)
            }
            // No fill and the automatic font colour (what recordings write); other palette indexes aren't supported.
            (XlObj::Interior(r), "colorindex" | "pattern") if value.to_i32()? == -4142 => {
                let action = json!({"op": "set", "path": self.sub_path(r, &["format", "fill"]), "prop": "color", "value": "none"});
                self.write(host, action)
            }
            (XlObj::Font(r), "colorindex") if value.to_i32()? == -4105 => {
                let action = json!({"op": "set", "path": self.sub_path(r, &["format", "font"]), "prop": "color", "value": "auto"});
                self.write(host, action)
            }
            (XlObj::Interior(_) | XlObj::Font(_), "colorindex" | "pattern") => {
                Err(VbaError::with(438, "ColorIndex and Pattern aren't supported yet, except xlNone and xlAutomatic (use .Color)"))
            }
            _ => Err(VbaError::new(438)),
        }
    }

    /// `object(args)`: the default member (`Range(1, 1)` is `Range.Item(1, 1)`).
    pub fn default_call(&mut self, host: &mut dyn Host, o: &XlObj, args: Args) -> R<Value> {
        match o {
            XlObj::Range(_) | XlObj::Rows(_) | XlObj::Columns(_) if !args.is_empty() => self.member(host, o, "item", args, &[]),
            XlObj::Worksheets => self.member(host, o, "item", args, &[]),
            XlObj::Range(_) => self.member(host, o, "value", args, &[]),
            _ => Err(VbaError::new(438)),
        }
    }

    /// The objects `For Each` walks: a range's cells (row by row), the worksheets.
    pub fn items(&mut self, host: &mut dyn Host, o: &XlObj) -> R<Vec<Value>> {
        let obj = |x: XlObj| Value::Object(crate::vba::value::Obj::Xl(x));
        match o {
            XlObj::Worksheets => {
                let names = self.sheet_names(host)?;
                Ok(names.iter().map(|n| obj(XlObj::Worksheet(self.sheet_id(n)))).collect())
            }
            XlObj::Range(r) => {
                if r.count() > MAX_CELLS {
                    return Err(VbaError::with(7, "Out of memory (more than a million cells)"));
                }
                Ok((r.r1..=r.r2).flat_map(|row| (r.c1..=r.c2).map(move |c| (row, c))).map(|(row, c)| obj(XlObj::Range(r.cell(row, c)))).collect())
            }
            XlObj::Rows(r) => Ok((r.r1..=r.r2).take(MAX_CELLS as usize).map(|row| obj(XlObj::Range(XlRange { r1: row, r2: row, ..*r }))).collect()),
            XlObj::Columns(r) => Ok((r.c1..=r.c2).map(|c| obj(XlObj::Range(XlRange { c1: c, c2: c, ..*r }))).collect()),
            _ => Err(VbaError::new(438)),
        }
    }

    /// The value an object stands for where a value is wanted (a range's `Value`).
    pub fn default_value(&mut self, host: &mut dyn Host, o: &XlObj) -> R<Value> {
        match o {
            XlObj::Range(_) => self.member(host, o, "value", &[], &[]),
            XlObj::Worksheet(s) => Ok(Value::str(self.sheets.get(*s).cloned().unwrap_or_default())),
            XlObj::Application => Ok(Value::str("GridCraft")),
            _ => Err(VbaError::new(438)),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn addresses() {
        assert_eq!(col_name(0), "A");
        assert_eq!(col_name(25), "Z");
        assert_eq!(col_name(26), "AA");
        assert_eq!(col_name(16_383), "XFD");
        assert_eq!(parse_a1("B2"), Some((1, 1, 1, 1)));
        assert_eq!(parse_a1("$C$3:A1"), Some((0, 0, 2, 2)));
        assert_eq!(parse_a1("B:C"), Some((0, 1, MAX_ROWS - 1, 2)));
        assert_eq!(parse_a1("2:4"), Some((1, 0, 3, MAX_COLS - 1)));
        assert_eq!(parse_a1("ZZZZ1"), None);
        assert_eq!(parse_a1("A0"), None);
        assert_eq!(parse_address("'It''s'!A1:B2"), Some(("It's".into(), (0, 0, 1, 1))));
        let r = XlRange { sheet: 0, r1: 0, c1: 0, r2: 1, c2: 2 };
        assert_eq!(r.absolute(), "$A$1:$C$2");
        assert_eq!(color_from_hex("#FF0000"), Some(255));
        assert_eq!(color_from_hex("#0000FF"), Some(16_711_680));
        assert_eq!(color_to_hex(65_280), "#00FF00");
    }
}
