//! A document in memory that answers the host requests as WordCraft's Office bridge does (the
//! Word JavaScript API's paths, props and result shapes per the bridge contract, `host/vbaProject/*`,
//! `host/message`, `host/log`). Tests only; it depends on nothing but serde_json, so the
//! end-to-end test includes it too.
//!
//! Positions are (paragraph, character); ids are the contract's: `p:<i>` for a paragraph,
//! `r:<i>:<byte>-<j>:<byte>` for a range. A paragraph's whole range takes its paragraph mark
//! (it ends at the start of the next paragraph), except the last one's.
#![allow(clippy::unwrap_used, clippy::indexing_slicing, clippy::panic, dead_code)]

use serde_json::{Value as Json, json};

#[derive(Clone, Debug, PartialEq)]
pub struct Fmt {
    pub bold: bool,
    pub italic: bool,
    /// `None`, `Single`, `Double`, … (the Word API's names).
    pub underline: String,
    pub size: f64,
    pub name: String,
    pub color: String,
    /// `None` is no highlight.
    pub highlight: Option<String>,
}

impl Default for Fmt {
    fn default() -> Fmt {
        Fmt { bold: false, italic: false, underline: "None".into(), size: 11.0, name: "Calibri".into(), color: "#000000".into(), highlight: None }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Para {
    pub chars: Vec<(char, Fmt)>,
    /// The formatting of the paragraph mark (what typing into an empty paragraph gets).
    pub mark: Fmt,
    pub style: String,
    /// `Left`, `Centered`, `Right`, `Justified`.
    pub alignment: String,
    pub first_line_indent: f64,
    pub left_indent: f64,
    pub right_indent: f64,
    pub line_spacing: f64,
    pub space_after: f64,
    pub space_before: f64,
}

impl Para {
    pub fn new(text: &str, style: &str) -> Para {
        Para {
            chars: text.chars().map(|c| (c, Fmt::default())).collect(),
            mark: Fmt::default(),
            style: style.into(),
            alignment: "Left".into(),
            first_line_indent: 0.0,
            left_indent: 0.0,
            right_indent: 0.0,
            line_spacing: 12.0,
            space_after: 8.0,
            space_before: 0.0,
        }
    }

    pub fn text(&self) -> String {
        self.chars.iter().map(|(c, _)| *c).collect()
    }

    fn empty_like(&self) -> Para {
        Para { chars: Vec::new(), mark: self.mark.clone(), ..self.clone() }
    }

    fn byte(&self, ci: usize) -> usize {
        self.chars.iter().take(ci).map(|(c, _)| c.len_utf8()).sum()
    }

    fn char_at_byte(&self, b: usize) -> Option<usize> {
        let mut at = 0;
        for (i, (c, _)) in self.chars.iter().enumerate() {
            if at == b {
                return Some(i);
            }
            at += c.len_utf8();
        }
        (at == b).then_some(self.chars.len())
    }
}

type Pos = (usize, usize);

#[derive(Clone, Debug)]
pub struct Doc {
    pub paras: Vec<Para>,
    /// The selection: anchor and end, in order.
    pub sel: (Pos, Pos),
    pub title: String,
    pub subject: String,
    pub author: String,
    pub keywords: String,
    pub comments: String,
    /// The document's name; `None` answers `name` as an app without it would.
    pub name: Option<String>,
    pub saves: usize,
    pub vba: Option<String>,
    pub messages: Vec<(String, Option<String>)>,
    pub logs: Vec<String>,
    pub syncs: usize,
    pub actions: usize,
}

impl Default for Doc {
    /// Three paragraphs: a centred title ("Hello " bold, "world" not), and two plain ones; the
    /// caret at the end of the last.
    fn default() -> Doc {
        let mut title = Para::new("Hello world", "Title");
        title.alignment = "Centered".into();
        for (_, f) in title.chars.iter_mut().take(6) {
            f.bold = true;
        }
        Doc {
            paras: vec![title, Para::new("Second paragraph.", "Normal"), Para::new("Third.", "Normal")],
            sel: ((2, 6), (2, 6)),
            title: "Report".into(),
            subject: String::new(),
            author: "Ada".into(),
            keywords: String::new(),
            comments: String::new(),
            name: Some("Report.docx".into()),
            saves: 0,
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
    Document,
    Body,
    Properties,
    /// Paragraphs `first..=last` (`None`: the document's).
    Paras(Option<(usize, usize)>),
    Para(usize),
    Range(Pos, Pos),
    Font(Pos, Pos),
}

type E = (&'static str, String);

fn bad(m: impl Into<String>) -> E {
    ("InvalidArgument", m.into())
}

fn not_found(m: impl Into<String>) -> E {
    ("ItemNotFound", m.into())
}

fn color(j: &Json) -> Result<String, E> {
    let s = j.as_str().ok_or(bad("a colour is #RRGGBB"))?;
    if s.len() == 7 && s.starts_with('#') && u32::from_str_radix(&s[1..], 16).is_ok() {
        Ok(s.to_ascii_uppercase())
    } else {
        Err(bad(format!("`{s}` isn't a colour")))
    }
}

const HIGHLIGHT_NAMES: &[&str] = &[
    "Black",
    "Blue",
    "Turquoise",
    "BrightGreen",
    "Pink",
    "Red",
    "Yellow",
    "White",
    "DarkBlue",
    "Teal",
    "Green",
    "Violet",
    "DarkRed",
    "DarkYellow",
    "Gray50",
    "Gray25",
];

const UNDERLINES: &[&str] = &["None", "Single", "Word", "Double", "Dotted", "Thick", "DashLine", "DotDashLine", "TwoDotDashLine", "Wave"];

impl Doc {
    /// A document of plain Normal paragraphs.
    pub fn with_text(paras: &[&str]) -> Doc {
        let mut d = Doc { paras: paras.iter().map(|t| Para::new(t, "Normal")).collect(), ..Doc::default() };
        let last = d.paras.len() - 1;
        let end = d.paras[last].chars.len();
        d.sel = ((last, end), (last, end));
        d
    }

    pub fn texts(&self) -> Vec<String> {
        self.paras.iter().map(Para::text).collect()
    }

    /// The formatting of character `ci` of paragraph `pi`.
    pub fn fmt(&self, pi: usize, ci: usize) -> &Fmt {
        &self.paras[pi].chars[ci].1
    }

    /// One host request; errors are (JSON-RPC code, message, data).
    pub fn handle(&mut self, method: &str, params: &Json) -> Result<Json, (i64, String, Json)> {
        match method {
            "host/sync" => {
                self.syncs += 1;
                let before = self.clone();
                let actions = params["actions"].as_array().cloned().unwrap_or_default();
                let mut results: Vec<Json> = Vec::new();
                for (i, a) in actions.iter().enumerate() {
                    self.actions += 1;
                    match self.resolve_refs(a, &results).and_then(|a| self.action(&a)) {
                        Ok(v) => results.push(v),
                        Err((code, message)) => {
                            let (syncs, actions_n) = (self.syncs, self.actions);
                            *self = before;
                            self.syncs = syncs;
                            self.actions = actions_n;
                            return Err((-32000, message, json!({"code": code, "action": i})));
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

    /// `{"result": i}` in a step's arguments is the id of what action `i` made.
    fn resolve_refs(&self, a: &Json, results: &[Json]) -> Result<Json, E> {
        let mut a = a.clone();
        if let Some(steps) = a.get_mut("path").and_then(Json::as_array_mut) {
            for st in steps {
                if let Some(args) = st.get_mut("a").and_then(Json::as_array_mut) {
                    for x in args {
                        if let Some(i) = x.get("result").and_then(Json::as_u64) {
                            let id = results.get(i as usize).and_then(|r| r.get("id")).cloned().ok_or(bad(format!("action {i} made nothing with an id")))?;
                            *x = id;
                        }
                    }
                }
            }
        }
        Ok(a)
    }

    fn last(&self) -> usize {
        self.paras.len() - 1
    }

    fn end(&self) -> Pos {
        (self.last(), self.paras[self.last()].chars.len())
    }

    fn para_range(&self, i: usize) -> (Pos, Pos) {
        if i < self.last() { ((i, 0), (i + 1, 0)) } else { ((i, 0), (i, self.paras[i].chars.len())) }
    }

    fn para_id(i: usize) -> Json {
        json!(format!("p:{i}"))
    }

    fn range_id(&self, a: Pos, b: Pos) -> Json {
        json!(format!("r:{}:{}-{}:{}", a.0, self.paras[a.0].byte(a.1), b.0, self.paras[b.0].byte(b.1)))
    }

    fn parse_para_id(&self, j: &Json) -> Result<usize, E> {
        let s = j.as_str().ok_or(bad("a paragraph id"))?;
        let i: usize = s.strip_prefix("p:").and_then(|n| n.parse().ok()).ok_or(not_found(format!("no paragraph {s}")))?;
        if i < self.paras.len() { Ok(i) } else { Err(not_found(format!("no paragraph {s}"))) }
    }

    fn parse_range_id(&self, j: &Json) -> Result<(Pos, Pos), E> {
        let s = j.as_str().ok_or(bad("a range id"))?;
        let gone = || not_found(format!("no range {s}"));
        let body = s.strip_prefix("r:").ok_or_else(gone)?;
        let (from, to) = body.split_once('-').ok_or_else(gone)?;
        let pos = |x: &str| -> Result<Pos, E> {
            let (p, b) = x.rsplit_once(':').ok_or_else(gone)?;
            let p: usize = p.parse().map_err(|_| gone())?;
            let b: usize = b.parse().map_err(|_| gone())?;
            let para = self.paras.get(p).ok_or_else(gone)?;
            Ok((p, para.char_at_byte(b).ok_or_else(gone)?))
        };
        let (a, b) = (pos(from)?, pos(to)?);
        if a > b { Err(gone()) } else { Ok((a, b)) }
    }

    /// The paragraphs a range covers (one ending at the start of a paragraph leaves it out).
    fn scope(a: Pos, b: Pos) -> (usize, usize) {
        let last = if b.1 == 0 && b.0 > a.0 { b.0 - 1 } else { b.0 };
        (a.0, last)
    }

    fn scope_ids(&self, scope: Option<(usize, usize)>) -> Vec<usize> {
        match scope {
            None => (0..self.paras.len()).collect(),
            Some((a, b)) => (a..=b.min(self.last())).collect(),
        }
    }

    fn location(a: &Json) -> String {
        a.as_str().unwrap_or("").to_ascii_lowercase()
    }

    fn resolve(&self, path: &Json) -> Result<Obj, E> {
        let mut o: Option<Obj> = None;
        for step in path.as_array().cloned().unwrap_or_default() {
            let p = step.get("p").and_then(Json::as_str);
            let m = step.get("m").and_then(Json::as_str);
            let a = step.get("a").cloned().unwrap_or(json!([]));
            let a0 = a.get(0).cloned().unwrap_or(Json::Null);
            o = Some(match (&o, p, m) {
                (None, Some("document"), _) => Obj::Document,
                (Some(Obj::Document), Some("body"), _) => Obj::Body,
                (Some(Obj::Document), Some("properties"), _) => Obj::Properties,
                (Some(Obj::Document | Obj::Body), Some("paragraphs"), _) => Obj::Paras(None),
                (Some(Obj::Document), _, Some("getSelection")) => Obj::Range(self.sel.0, self.sel.1),
                (Some(Obj::Document), _, Some("getRangeById")) => {
                    let (x, y) = self.parse_range_id(&a0)?;
                    Obj::Range(x, y)
                }
                (Some(Obj::Paras(scope)), _, Some(f @ ("getById" | "getItemAt" | "getFirst" | "getLast"))) => {
                    let ids = self.scope_ids(*scope);
                    let i = match f {
                        "getById" => Some(self.parse_para_id(&a0)?).filter(|i| ids.contains(i)),
                        "getItemAt" => a0.as_u64().and_then(|n| ids.get(n as usize).copied()),
                        "getFirst" => ids.first().copied(),
                        _ => ids.last().copied(),
                    };
                    Obj::Para(i.ok_or(not_found(format!("no paragraph {a0}")))?)
                }
                (Some(Obj::Para(i)), _, Some("getNext")) => {
                    if *i >= self.last() {
                        return Err(not_found("the last paragraph has no next"));
                    }
                    Obj::Para(i + 1)
                }
                (Some(Obj::Para(i)), _, Some("getPrevious")) => Obj::Para(i.checked_sub(1).ok_or(not_found("the first paragraph has no previous"))?),
                (Some(Obj::Para(i)), _, Some("getRange")) => {
                    let (x, y) = self.para_range(*i);
                    match Self::location(&a0).as_str() {
                        "whole" => Obj::Range(x, y),
                        "content" => Obj::Range(x, (*i, self.paras[*i].chars.len())),
                        "start" => Obj::Range(x, x),
                        "end" => {
                            let e = (*i, self.paras[*i].chars.len());
                            Obj::Range(e, e)
                        }
                        l => return Err(bad(format!("Paragraph.getRange({l})"))),
                    }
                }
                (Some(Obj::Para(i)), Some("font"), _) => {
                    let (x, y) = self.para_range(*i);
                    Obj::Font(x, y)
                }
                (Some(Obj::Body), _, Some("getRange")) => match Self::location(&a0).as_str() {
                    "whole" | "content" => Obj::Range((0, 0), self.end()),
                    "start" => Obj::Range((0, 0), (0, 0)),
                    "end" => Obj::Range(self.end(), self.end()),
                    l => return Err(bad(format!("Body.getRange({l})"))),
                },
                (Some(Obj::Body), Some("font"), _) => Obj::Font((0, 0), self.end()),
                (Some(Obj::Range(x, y)), Some("font"), _) => Obj::Font(*x, *y),
                (Some(Obj::Range(x, y)), Some("paragraphs"), _) => Obj::Paras(Some(Self::scope(*x, *y))),
                (o, p, m) => return Err(bad(format!("{o:?} has no {p:?}/{m:?}"))),
            });
        }
        o.ok_or(bad("empty path"))
    }

    fn text_of(&self, a: Pos, b: Pos) -> String {
        let mut out = String::new();
        for pi in a.0..=b.0 {
            let p = &self.paras[pi];
            let from = if pi == a.0 { a.1 } else { 0 };
            let to = if pi == b.0 { b.1 } else { p.chars.len() };
            out.extend(p.chars[from..to].iter().map(|(c, _)| *c));
            if pi < b.0 {
                out.push('\n');
            }
        }
        out
    }

    fn fmts_of(&self, a: Pos, b: Pos) -> Vec<Fmt> {
        let mut v = Vec::new();
        for pi in a.0..=b.0 {
            let p = &self.paras[pi];
            let from = if pi == a.0 { a.1 } else { 0 };
            let to = if pi == b.0 { b.1 } else { p.chars.len() };
            v.extend(p.chars[from..to].iter().map(|(_, f)| f.clone()));
        }
        v
    }

    /// The formatting text typed at `pos` gets.
    fn fmt_at(&self, pos: Pos) -> Fmt {
        let p = &self.paras[pos.0];
        if pos.1 > 0 {
            p.chars[pos.1 - 1].1.clone()
        } else if let Some((_, f)) = p.chars.first() {
            f.clone()
        } else {
            p.mark.clone()
        }
    }

    fn clamp(&mut self) {
        let fix = |d: &Doc, p: Pos| -> Pos {
            let pi = p.0.min(d.last());
            (pi, p.1.min(d.paras[pi].chars.len()))
        };
        self.sel = (fix(self, self.sel.0), fix(self, self.sel.1));
    }

    fn insert_at(&mut self, pos: Pos, text: &str) -> (Pos, Pos) {
        let fmt = self.fmt_at(pos);
        let (pi, ci) = pos;
        let tail = self.paras[pi].chars.split_off(ci);
        let mut parts = text.split('\n');
        self.paras[pi].chars.extend(parts.next().unwrap_or("").chars().map(|c| (c, fmt.clone())));
        let mut cur = pi;
        for part in parts {
            let mut np = self.paras[cur].empty_like();
            np.chars = part.chars().map(|c| (c, fmt.clone())).collect();
            self.paras.insert(cur + 1, np);
            cur += 1;
        }
        let end = (cur, self.paras[cur].chars.len());
        self.paras[cur].chars.extend(tail);
        (pos, end)
    }

    fn delete(&mut self, a: Pos, b: Pos) {
        if a.1 == 0 && b.1 == 0 && b.0 > a.0 {
            // Whole paragraphs, marks included.
            self.paras.drain(a.0..b.0);
        } else if a.0 == b.0 {
            self.paras[a.0].chars.drain(a.1..b.1);
        } else {
            let tail: Vec<_> = self.paras[b.0].chars[b.1..].to_vec();
            self.paras[a.0].chars.truncate(a.1);
            self.paras[a.0].chars.extend(tail);
            self.paras.drain(a.0 + 1..=b.0);
        }
    }

    fn new_paragraph(&mut self, at: usize, like: usize, text: &str) -> Json {
        let mut np = self.paras[like].empty_like();
        np.chars = text.chars().map(|c| (c, np.mark.clone())).collect();
        self.paras.insert(at, np);
        json!({"id": Self::para_id(at)})
    }

    fn load(&self, obj: &Obj, props: &[String]) -> Result<Json, E> {
        let mut out = serde_json::Map::new();
        for p in props {
            let v = match (obj, p.as_str()) {
                (Obj::Document, "name") => json!(self.name.clone().ok_or(bad("Document.name"))?),
                (Obj::Properties, k) => json!(match k {
                    "title" => &self.title,
                    "subject" => &self.subject,
                    "author" => &self.author,
                    "keywords" => &self.keywords,
                    "comments" => &self.comments,
                    _ => return Err(bad(format!("DocumentProperties.{k}"))),
                }),
                (Obj::Paras(s), "count") => json!(self.scope_ids(*s).len()),
                (Obj::Paras(s), "items") => json!(self.scope_ids(*s).into_iter().map(|i| json!({"id": Self::para_id(i)})).collect::<Vec<_>>()),
                (Obj::Para(i), k) => {
                    let para = &self.paras[*i];
                    match k {
                        "id" => Self::para_id(*i),
                        "text" => json!(para.text()),
                        "style" => json!(para.style),
                        "styleBuiltIn" => json!(para.style.replace(' ', "")),
                        "alignment" => json!(para.alignment),
                        "firstLineIndent" => json!(para.first_line_indent),
                        "leftIndent" => json!(para.left_indent),
                        "rightIndent" => json!(para.right_indent),
                        "lineSpacing" => json!(para.line_spacing),
                        "spaceAfter" => json!(para.space_after),
                        "spaceBefore" => json!(para.space_before),
                        "isListItem" => json!(false),
                        _ => return Err(bad(format!("Paragraph.{k}"))),
                    }
                }
                (Obj::Body, "text") => json!(self.text_of((0, 0), self.end())),
                (Obj::Range(a, b), k) => match k {
                    "id" => self.range_id(*a, *b),
                    "text" => json!(self.text_of(*a, *b)),
                    "isEmpty" => json!(a == b),
                    "style" => {
                        let (x, y) = Self::scope(*a, *b);
                        let first = &self.paras[x].style;
                        if self.paras[x..=y].iter().all(|p| &p.style == first) { json!(first) } else { Json::Null }
                    }
                    _ => return Err(bad(format!("Range.{k}"))),
                },
                (Obj::Font(a, b), k) => {
                    let mut fmts = self.fmts_of(*a, *b);
                    if fmts.is_empty() {
                        fmts.push(self.fmt_at(*a));
                    }
                    let get = |f: &Fmt| -> Json {
                        match k {
                            "bold" => json!(f.bold),
                            "italic" => json!(f.italic),
                            "underline" => json!(f.underline),
                            "size" => json!(f.size),
                            "name" => json!(f.name),
                            "color" => json!(f.color),
                            "highlightColor" => json!(f.highlight),
                            _ => Json::String("?".into()),
                        }
                    };
                    let first = get(&fmts[0]);
                    if first == json!("?") {
                        return Err(bad(format!("Font.{k}")));
                    }
                    if fmts.iter().all(|f| get(f) == first) { first } else { Json::Null }
                }
                (o, p) => return Err(bad(format!("{o:?}.{p}"))),
            };
            out.insert(p.clone(), v);
        }
        Ok(Json::Object(out))
    }

    fn set_style(&mut self, paras: (usize, usize), value: &Json) -> Result<(), E> {
        let s = value.as_str().filter(|s| !s.trim().is_empty()).ok_or(bad("a style name"))?.to_string();
        for p in &mut self.paras[paras.0..=paras.1] {
            p.style = s.clone();
        }
        Ok(())
    }

    fn set(&mut self, obj: &Obj, prop: &str, value: &Json) -> Result<(), E> {
        let num = || value.as_f64().ok_or(bad(format!("{prop} is a number")));
        let text = || value.as_str().map(str::to_string).ok_or(bad(format!("{prop} is text")));
        match obj {
            Obj::Properties => {
                let v = text()?;
                match prop {
                    "title" => self.title = v,
                    "subject" => self.subject = v,
                    "author" => self.author = v,
                    "keywords" => self.keywords = v,
                    "comments" => self.comments = v,
                    _ => return Err(bad(format!("set DocumentProperties.{prop}"))),
                }
            }
            Obj::Para(i) => match prop {
                "style" => self.set_style((*i, *i), value)?,
                "styleBuiltIn" => {
                    let s = text()?;
                    let name = match s.as_str() {
                        h if h.starts_with("Heading") => format!("Heading {}", &h[7..]),
                        "IntenseQuote" => "Intense Quote".into(),
                        "ListParagraph" => "List Paragraph".into(),
                        o => o.to_string(),
                    };
                    self.paras[*i].style = name;
                }
                "alignment" => {
                    let a = text()?;
                    if !["Left", "Centered", "Right", "Justified"].contains(&a.as_str()) {
                        return Err(bad(format!("alignment {a}")));
                    }
                    self.paras[*i].alignment = a;
                }
                "firstLineIndent" => self.paras[*i].first_line_indent = num()?,
                "leftIndent" => self.paras[*i].left_indent = num()?,
                "rightIndent" => self.paras[*i].right_indent = num()?,
                "lineSpacing" => self.paras[*i].line_spacing = num()?,
                "spaceAfter" => self.paras[*i].space_after = num()?,
                "spaceBefore" => self.paras[*i].space_before = num()?,
                _ => return Err(bad(format!("set Paragraph.{prop}"))),
            },
            Obj::Range(a, b) if prop == "style" => self.set_style(Self::scope(*a, *b), value)?,
            Obj::Body if prop == "style" => self.set_style((0, self.last()), value)?,
            Obj::Font(a, b) => {
                let apply: Box<dyn Fn(&mut Fmt)> = match prop {
                    "bold" => {
                        let v = value.as_bool().ok_or(bad("bold is true or false"))?;
                        Box::new(move |f| f.bold = v)
                    }
                    "italic" => {
                        let v = value.as_bool().ok_or(bad("italic is true or false"))?;
                        Box::new(move |f| f.italic = v)
                    }
                    "underline" => {
                        let v = text()?;
                        if !UNDERLINES.contains(&v.as_str()) {
                            return Err(bad(format!("underline {v}")));
                        }
                        Box::new(move |f| f.underline = v.clone())
                    }
                    "size" => {
                        let v = num()?;
                        Box::new(move |f| f.size = v)
                    }
                    "name" => {
                        let v = text()?;
                        Box::new(move |f| f.name = v.clone())
                    }
                    "color" => {
                        let v = color(value)?;
                        Box::new(move |f| f.color = v.clone())
                    }
                    "highlightColor" => {
                        let v = match value {
                            Json::Null => None,
                            Json::String(s) if HIGHLIGHT_NAMES.iter().any(|n| n.eq_ignore_ascii_case(s)) => Some(s.clone()),
                            other => Some(color(other)?),
                        };
                        Box::new(move |f| f.highlight = v.clone())
                    }
                    _ => return Err(bad(format!("set Font.{prop}"))),
                };
                let (a, b) = (*a, *b);
                if a == b {
                    apply(&mut self.paras[a.0].mark);
                }
                for pi in a.0..=b.0 {
                    let len = self.paras[pi].chars.len();
                    let from = if pi == a.0 { a.1 } else { 0 };
                    let to = if pi == b.0 { b.1 } else { len };
                    for (_, f) in &mut self.paras[pi].chars[from..to] {
                        apply(f);
                    }
                }
            }
            o => return Err(bad(format!("set {o:?}.{prop}"))),
        }
        Ok(())
    }

    fn action(&mut self, a: &Json) -> Result<Json, E> {
        let obj = self.resolve(&a["path"])?;
        let r = match a["op"].as_str() {
            Some("load") => {
                let props: Vec<String> = a["props"].as_array().cloned().unwrap_or_default().iter().filter_map(|p| p.as_str().map(str::to_string)).collect();
                return self.load(&obj, &props);
            }
            Some("set") => {
                self.set(&obj, a["prop"].as_str().unwrap_or(""), &a["value"])?;
                Json::Null
            }
            Some("call") => {
                let args = a["args"].as_array().cloned().unwrap_or_default();
                let text = args.first().and_then(Json::as_str).unwrap_or("").to_string();
                let loc = args.get(1).map(Self::location).unwrap_or_default();
                match (&obj, a["method"].as_str().unwrap_or("")) {
                    (Obj::Document, "save") => {
                        self.saves += 1;
                        Json::Null
                    }
                    (Obj::Body, "clear") => {
                        let like = self.paras[0].empty_like();
                        self.paras = vec![like];
                        Json::Null
                    }
                    (Obj::Body, "insertText") => self.insert_text((0, 0), self.end(), &text, &loc)?,
                    (Obj::Body, "insertParagraph") => match loc.as_str() {
                        "start" => self.new_paragraph(0, 0, &text),
                        "end" => self.new_paragraph(self.paras.len(), self.last(), &text),
                        l => return Err(bad(format!("Body.insertParagraph({l})"))),
                    },
                    (Obj::Para(i), "insertText") => {
                        let (x, y) = self.para_range(*i);
                        self.insert_text(x, y, &text, &loc)?
                    }
                    (Obj::Para(i), "insertParagraph") => match loc.as_str() {
                        "before" => self.new_paragraph(*i, *i, &text),
                        "after" => self.new_paragraph(i + 1, *i, &text),
                        l => return Err(bad(format!("Paragraph.insertParagraph({l})"))),
                    },
                    (Obj::Para(i), "delete") => {
                        if self.paras.len() > 1 {
                            self.paras.remove(*i);
                        } else {
                            self.paras[0].chars.clear();
                        }
                        Json::Null
                    }
                    (Obj::Para(i), "select") => {
                        self.sel = self.para_range(*i);
                        Json::Null
                    }
                    (Obj::Range(x, y), "insertText") => self.insert_text(*x, *y, &text, &loc)?,
                    (Obj::Range(x, y), "insertParagraph") => {
                        let (f, l) = Self::scope(*x, *y);
                        match loc.as_str() {
                            "before" => self.new_paragraph(f, f, &text),
                            "after" => self.new_paragraph(l + 1, l, &text),
                            l => return Err(bad(format!("Range.insertParagraph({l})"))),
                        }
                    }
                    (Obj::Range(x, y), "delete") => {
                        self.delete(*x, *y);
                        Json::Null
                    }
                    (Obj::Range(x, y), "select") => {
                        self.sel = match args.first().map(Self::location).as_deref() {
                            Some("start") => (*x, *x),
                            Some("end") => (*y, *y),
                            None | Some("select" | "") => (*x, *y),
                            Some(m) => return Err(bad(format!("select({m})"))),
                        };
                        Json::Null
                    }
                    (o, m) => return Err(bad(format!("{o:?}.{m}()"))),
                }
            }
            _ => return Err(bad("bad op")),
        };
        self.clamp();
        Ok(r)
    }

    /// `insertText(text, location)` on the range `a..b`: the inserted text's range id.
    fn insert_text(&mut self, a: Pos, b: Pos, text: &str, loc: &str) -> Result<Json, E> {
        // Text "after" a range that ends with a paragraph mark goes before the mark.
        let end = if b.1 == 0 && b.0 > a.0 { (b.0 - 1, self.paras[b.0 - 1].chars.len()) } else { b };
        let at = match loc {
            "start" | "before" => a,
            "end" | "after" => end,
            "replace" => {
                self.delete(a, end);
                a
            }
            l => return Err(bad(format!("insertText({l})"))),
        };
        let (x, y) = self.insert_at(at, text);
        Ok(json!({"id": self.range_id(x, y)}))
    }
}
