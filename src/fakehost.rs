//! A workbook in memory that answers the host requests as GridCraft does (its `office.sync`
//! paths, props and result shapes, `host/vbaProject/*`, `host/message`, `host/log`). Tests only;
//! it depends on nothing but serde_json, so the end-to-end test includes it too.
#![allow(clippy::unwrap_used, clippy::indexing_slicing, clippy::panic, dead_code)]

use std::collections::BTreeMap;

use serde_json::{Value as Json, json};

#[derive(Clone, Debug, Default)]
pub struct Cell {
    pub value: Json,
    pub formula: Option<String>,
    pub numfmt: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub size: f64,
    pub font: String,
    pub color: String,
    pub fill: String,
    pub underline: String,
}

#[derive(Clone, Debug, Default)]
pub struct Sheet {
    pub name: String,
    pub cells: BTreeMap<(u32, u32), Cell>,
}

#[derive(Clone, Debug)]
pub struct Book {
    pub sheets: Vec<Sheet>,
    pub active: usize,
    /// (sheet, r1, c1, r2, c2) and the active cell.
    pub selection: (usize, u32, u32, u32, u32),
    pub active_cell: (u32, u32),
    pub vba: Option<String>,
    pub messages: Vec<(String, Option<String>)>,
    pub logs: Vec<String>,
    pub syncs: usize,
    pub actions: usize,
}

impl Default for Book {
    fn default() -> Book {
        Book {
            sheets: vec![Sheet { name: "Sheet1".into(), cells: BTreeMap::new() }],
            active: 0,
            selection: (0, 0, 0, 0, 0),
            active_cell: (0, 0),
            vba: None,
            messages: Vec::new(),
            logs: Vec::new(),
            syncs: 0,
            actions: 0,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
enum Obj {
    Workbook,
    Worksheets,
    Worksheet(usize),
    Range(usize, u32, u32, u32, u32),
    Format(usize, u32, u32, u32, u32),
    Font(usize, u32, u32, u32, u32),
    Fill(usize, u32, u32, u32, u32),
}

type E = (&'static str, String);

pub fn col_name(c: u32) -> String {
    let mut n = c + 1;
    let mut s = Vec::new();
    while n > 0 {
        s.push(b'A' + ((n - 1) % 26) as u8);
        n = (n - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

fn parse_cell(s: &str) -> Option<(u32, u32)> {
    let letters: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let digits = &s[letters.len()..];
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let col = letters.to_ascii_uppercase().bytes().fold(0u32, |a, b| a * 26 + u32::from(b - b'A' + 1)) - 1;
    Some((digits.parse::<u32>().ok()? - 1, col))
}

pub fn parse_range(s: &str) -> Option<(u32, u32, u32, u32)> {
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    if let (Some((r1, c1)), Some((r2, c2))) = (parse_cell(a), parse_cell(b)) {
        return Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)));
    }
    let colnum =
        |t: &str| t.chars().all(|c| c.is_ascii_alphabetic()).then(|| t.to_ascii_uppercase().bytes().fold(0u32, |a, b| a * 26 + u32::from(b - b'A' + 1)) - 1);
    if let (Some(c1), Some(c2)) = (colnum(a), colnum(b)) {
        return Some((0, c1, 1_048_575, c2));
    }
    if let (Ok(r1), Ok(r2)) = (a.parse::<u32>(), b.parse::<u32>()) {
        return Some((r1 - 1, 0, r2 - 1, 16_383));
    }
    None
}

fn a1(r1: u32, c1: u32, r2: u32, c2: u32) -> String {
    if r1 == r2 && c1 == c2 { format!("{}{}", col_name(c1), r1 + 1) } else { format!("{}{}:{}{}", col_name(c1), r1 + 1, col_name(c2), r2 + 1) }
}

impl Book {
    pub fn with_sheets(names: &[&str]) -> Book {
        Book { sheets: names.iter().map(|n| Sheet { name: n.to_string(), cells: BTreeMap::new() }).collect(), ..Default::default() }
    }

    pub fn get(&self, sheet: &str, a: &str) -> Json {
        let (r, c) = parse_cell(a).unwrap();
        let s = self.sheets.iter().find(|s| s.name == sheet).unwrap();
        s.cells.get(&(r, c)).map(|c| c.value.clone()).unwrap_or(Json::Null)
    }

    pub fn cell(&self, sheet: &str, a: &str) -> Cell {
        let (r, c) = parse_cell(a).unwrap();
        let s = self.sheets.iter().find(|s| s.name == sheet).unwrap();
        s.cells.get(&(r, c)).cloned().unwrap_or_default()
    }

    pub fn put(&mut self, sheet: &str, a: &str, v: Json) {
        let (r, c) = parse_cell(a).unwrap();
        let s = self.sheets.iter_mut().find(|s| s.name == sheet).unwrap();
        s.cells.entry((r, c)).or_default().value = v;
    }

    /// One host request; errors are (JSON-RPC code, message, data).
    pub fn handle(&mut self, method: &str, params: &Json) -> Result<Json, (i64, String, Json)> {
        match method {
            "host/sync" => {
                self.syncs += 1;
                let before = self.clone();
                let actions = params["actions"].as_array().cloned().unwrap_or_default();
                let mut results = Vec::new();
                for (i, a) in actions.iter().enumerate() {
                    self.actions += 1;
                    match self.action(a) {
                        Ok(v) => results.push(v),
                        Err((code, message)) => {
                            let (syncs, actions_n) = (self.syncs, self.actions);
                            *self = before;
                            self.syncs = syncs;
                            self.actions = actions_n;
                            return Err((-32000, message.clone(), json!({"code": code, "message": message, "action": i})));
                        }
                    }
                }
                Ok(json!({"results": results}))
            }
            "host/vbaProject/get" => Ok(json!({"base64": self.vba})),
            "host/vbaProject/set" => {
                self.vba = params["base64"].as_str().map(str::to_string);
                Ok(json!({}))
            }
            "host/message" => {
                self.messages.push((params["text"].as_str().unwrap_or("").to_string(), params["title"].as_str().map(str::to_string)));
                Ok(json!({"result": 1}))
            }
            "host/log" => {
                self.logs.push(params["text"].as_str().unwrap_or("").to_string());
                Ok(json!({}))
            }
            other => Err((-32601, format!("unknown method {other}"), Json::Null)),
        }
    }

    fn sheet_index(&self, name: &str) -> Result<usize, E> {
        self.sheets.iter().position(|s| s.name.eq_ignore_ascii_case(name)).ok_or(("ItemNotFound", format!("no worksheet `{name}`")))
    }

    fn used(&self, i: usize) -> (u32, u32, u32, u32) {
        let s = &self.sheets[i];
        let filled: Vec<&(u32, u32)> = s.cells.iter().filter(|(_, c)| !c.value.is_null() || c.formula.is_some()).map(|(k, _)| k).collect();
        if filled.is_empty() {
            return (0, 0, 0, 0);
        }
        let r1 = filled.iter().map(|k| k.0).min().unwrap();
        let c1 = filled.iter().map(|k| k.1).min().unwrap();
        let r2 = filled.iter().map(|k| k.0).max().unwrap();
        let c2 = filled.iter().map(|k| k.1).max().unwrap();
        (r1, c1, r2, c2)
    }

    fn resolve(&self, path: &Json) -> Result<Obj, E> {
        let mut o: Option<Obj> = None;
        for step in path.as_array().cloned().unwrap_or_default() {
            let p = step.get("p").and_then(Json::as_str);
            let m = step.get("m").and_then(Json::as_str);
            let a = step.get("a").cloned().unwrap_or(json!([]));
            o = Some(match (&o, p, m) {
                (None, Some("workbook"), _) => Obj::Workbook,
                (Some(Obj::Workbook), Some("worksheets"), _) => Obj::Worksheets,
                (Some(Obj::Workbook), _, Some("getSelectedRange")) => {
                    let (s, r1, c1, r2, c2) = self.selection;
                    Obj::Range(s, r1, c1, r2, c2)
                }
                (Some(Obj::Workbook), _, Some("getActiveCell")) => {
                    Obj::Range(self.active, self.active_cell.0, self.active_cell.1, self.active_cell.0, self.active_cell.1)
                }
                (Some(Obj::Worksheets), _, Some("getActiveWorksheet")) => Obj::Worksheet(self.active),
                (Some(Obj::Worksheets), _, Some("getItem")) => Obj::Worksheet(self.sheet_index(a[0].as_str().unwrap_or(""))?),
                (Some(Obj::Worksheet(i)), _, Some("getRange")) => {
                    let (r1, c1, r2, c2) = parse_range(a[0].as_str().unwrap_or("")).ok_or(("InvalidArgument", "bad address".to_string()))?;
                    Obj::Range(*i, r1, c1, r2, c2)
                }
                (Some(Obj::Worksheet(i)), _, Some("getUsedRange")) => {
                    let (r1, c1, r2, c2) = self.used(*i);
                    Obj::Range(*i, r1, c1, r2, c2)
                }
                (Some(Obj::Range(i, r1, c1, _, _)), _, Some("getCell")) => {
                    let (r, c) = (r1 + a[0].as_u64().unwrap() as u32, c1 + a[1].as_u64().unwrap() as u32);
                    Obj::Range(*i, r, c, r, c)
                }
                (Some(Obj::Range(i, r1, c1, r2, c2)), Some("format"), _) => Obj::Format(*i, *r1, *c1, *r2, *c2),
                (Some(Obj::Format(i, r1, c1, r2, c2)), Some("font"), _) => Obj::Font(*i, *r1, *c1, *r2, *c2),
                (Some(Obj::Format(i, r1, c1, r2, c2)), Some("fill"), _) => Obj::Fill(*i, *r1, *c1, *r2, *c2),
                (o, p, m) => return Err(("NotImplemented", format!("{o:?} has no {p:?}/{m:?}"))),
            });
        }
        o.ok_or(("InvalidArgument", "empty path".into()))
    }

    fn cells(&self, i: usize, r1: u32, c1: u32, r2: u32, c2: u32, f: impl Fn(&Cell) -> Json) -> Json {
        let s = &self.sheets[i];
        let empty = Cell::default();
        Json::Array((r1..=r2).map(|r| Json::Array((c1..=c2).map(|c| f(s.cells.get(&(r, c)).unwrap_or(&empty))).collect())).collect())
    }

    fn action(&mut self, a: &Json) -> Result<Json, E> {
        let obj = self.resolve(&a["path"])?;
        match a["op"].as_str() {
            Some("load") => {
                let props: Vec<String> = a["props"].as_array().cloned().unwrap_or_default().iter().filter_map(|p| p.as_str().map(str::to_string)).collect();
                let mut out = serde_json::Map::new();
                for p in props {
                    let v = match (&obj, p.as_str()) {
                        (Obj::Range(i, r1, c1, r2, c2), "values") => self.cells(*i, *r1, *c1, *r2, *c2, |c| c.value.clone()),
                        (Obj::Range(i, r1, c1, r2, c2), "formulas") => {
                            self.cells(*i, *r1, *c1, *r2, *c2, |c| c.formula.clone().map(Json::String).unwrap_or(c.value.clone()))
                        }
                        (Obj::Range(i, r1, c1, r2, c2), "text") => self.cells(*i, *r1, *c1, *r2, *c2, |c| {
                            json!(match (&c.value, c.numfmt.as_deref()) {
                                (Json::Number(n), Some("0.00")) => format!("{:.2}", n.as_f64().unwrap()),
                                (Json::Null, _) => String::new(),
                                (Json::String(s), _) => s.clone(),
                                (v, _) => v.to_string(),
                            })
                        }),
                        (Obj::Range(i, r1, c1, r2, c2), "numberFormat") => {
                            self.cells(*i, *r1, *c1, *r2, *c2, |c| json!(c.numfmt.clone().unwrap_or("General".into())))
                        }
                        (Obj::Range(i, r1, c1, r2, c2), "address") => json!(format!("{}!{}", self.sheets[*i].name, a1(*r1, *c1, *r2, *c2))),
                        (Obj::Range(_, r1, _, r2, _), "rowCount") => json!(r2 - r1 + 1),
                        (Obj::Range(_, _, c1, _, c2), "columnCount") => json!(c2 - c1 + 1),
                        (Obj::Font(i, r1, c1, _, _), prop) => {
                            let c = self.sheets[*i].cells.get(&(*r1, *c1)).cloned().unwrap_or_default();
                            match prop {
                                "bold" => json!(c.bold),
                                "italic" => json!(c.italic),
                                "size" => json!(if c.size == 0.0 { 11.0 } else { c.size }),
                                "name" => json!(if c.font.is_empty() { "Calibri".to_string() } else { c.font }),
                                "color" => json!(if c.color.is_empty() { "#000000".to_string() } else { c.color }),
                                "underline" => json!(if c.underline.is_empty() { "None".to_string() } else { c.underline }),
                                _ => return Err(("NotImplemented", prop.to_string())),
                            }
                        }
                        (Obj::Fill(i, r1, c1, _, _), "color") => {
                            let c = self.sheets[*i].cells.get(&(*r1, *c1)).cloned().unwrap_or_default();
                            json!(if c.fill.is_empty() { "#FFFFFF".to_string() } else { c.fill })
                        }
                        (Obj::Worksheet(i), "name") => json!(self.sheets[*i].name),
                        (Obj::Worksheet(i), "position") => json!(i),
                        (Obj::Worksheets, "items") => {
                            json!(self.sheets.iter().enumerate().map(|(i, s)| json!({"name": s.name, "position": i})).collect::<Vec<_>>())
                        }
                        (Obj::Workbook, "name") => json!("Book1"),
                        (o, p) => return Err(("NotImplemented", format!("{o:?}.{p}"))),
                    };
                    out.insert(p, v);
                }
                Ok(Json::Object(out))
            }
            Some("set") => {
                let value = &a["value"];
                match (&obj, a["prop"].as_str().unwrap_or("")) {
                    (Obj::Range(i, r1, c1, r2, c2), prop @ ("values" | "formulas" | "numberFormat")) => {
                        let (h, w) = ((r2 - r1 + 1) as usize, (c2 - c1 + 1) as usize);
                        if let Json::Array(rows) = value
                            && (rows.len() != h || rows.iter().any(|r| r.as_array().map(Vec::len) != Some(w)))
                        {
                            return Err(("InvalidArgument", "size mismatch".into()));
                        }
                        for r in *r1..=*r2 {
                            for c in *c1..=*c2 {
                                let v = match value {
                                    Json::Array(rows) => rows[(r - r1) as usize][(c - c1) as usize].clone(),
                                    v => v.clone(),
                                };
                                let cell = self.sheets[*i].cells.entry((r, c)).or_default();
                                match prop {
                                    "numberFormat" => cell.numfmt = v.as_str().map(str::to_string),
                                    _ => {
                                        if let Some(f) = v.as_str().filter(|s| s.starts_with('=')) {
                                            cell.formula = Some(f.to_string());
                                            cell.value = Json::Null;
                                        } else {
                                            cell.formula = None;
                                            cell.value = v;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    (Obj::Font(i, r1, c1, r2, c2), prop) => {
                        for r in *r1..=*r2 {
                            for c in *c1..=*c2 {
                                let cell = self.sheets[*i].cells.entry((r, c)).or_default();
                                match prop {
                                    "bold" => cell.bold = value.as_bool().unwrap_or(false),
                                    "italic" => cell.italic = value.as_bool().unwrap_or(false),
                                    "size" => cell.size = value.as_f64().unwrap_or(11.0),
                                    "name" => cell.font = value.as_str().unwrap_or("").into(),
                                    "color" => cell.color = value.as_str().unwrap_or("").into(),
                                    "underline" => cell.underline = value.as_str().unwrap_or("").into(),
                                    _ => return Err(("NotImplemented", prop.to_string())),
                                }
                            }
                        }
                    }
                    (Obj::Fill(i, r1, c1, r2, c2), "color") => {
                        for r in *r1..=*r2 {
                            for c in *c1..=*c2 {
                                self.sheets[*i].cells.entry((r, c)).or_default().fill = value.as_str().unwrap_or("").into();
                            }
                        }
                    }
                    (Obj::Worksheet(i), "name") => self.sheets[*i].name = value.as_str().unwrap_or("").into(),
                    (o, p) => return Err(("NotImplemented", format!("set {o:?}.{p}"))),
                }
                Ok(Json::Null)
            }
            Some("call") => {
                let args = a["args"].as_array().cloned().unwrap_or_default();
                match (&obj, a["method"].as_str().unwrap_or("")) {
                    (Obj::Range(i, r1, c1, r2, c2), "clear") => {
                        let apply = args.first().and_then(Json::as_str).unwrap_or("All").to_string();
                        for r in *r1..=*r2 {
                            for c in *c1..=*c2 {
                                if let Some(cell) = self.sheets[*i].cells.get_mut(&(r, c)) {
                                    if apply != "Formats" {
                                        cell.value = Json::Null;
                                        cell.formula = None;
                                    }
                                    if apply != "Contents" {
                                        *cell = Cell { value: cell.value.clone(), formula: cell.formula.clone(), ..Default::default() };
                                    }
                                }
                            }
                        }
                        Ok(Json::Null)
                    }
                    (Obj::Range(i, r1, c1, r2, c2), "select") => {
                        self.active = *i;
                        self.selection = (*i, *r1, *c1, *r2, *c2);
                        self.active_cell = (*r1, *c1);
                        Ok(Json::Null)
                    }
                    (Obj::Worksheet(i), "activate") => {
                        self.active = *i;
                        Ok(Json::Null)
                    }
                    (Obj::Worksheet(i), "delete") => {
                        if self.sheets.len() == 1 {
                            return Err(("InvalidArgument", "a workbook keeps one sheet".into()));
                        }
                        self.sheets.remove(*i);
                        self.active = 0;
                        Ok(Json::Null)
                    }
                    (Obj::Worksheets, "add") => {
                        let name = args.first().and_then(Json::as_str).unwrap_or("SheetX").to_string();
                        if self.sheets.iter().any(|s| s.name.eq_ignore_ascii_case(&name)) {
                            return Err(("InvalidArgument", "that name is taken".into()));
                        }
                        self.sheets.push(Sheet { name: name.clone(), cells: BTreeMap::new() });
                        Ok(json!({"name": name}))
                    }
                    (o, m) => Err(("NotImplemented", format!("{o:?}.{m}()"))),
                }
            }
            _ => Err(("InvalidArgument", "bad op".into())),
        }
    }
}
