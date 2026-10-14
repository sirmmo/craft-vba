//! Word's object model for macros run in WordCraft (`Application`, `Documents`, `Document`,
//! `Paragraphs`, `Paragraph`, `Range`, `Font`, `ParagraphFormat`, `Selection`, `Styles`,
//! `BuiltInDocumentProperties`), over the host's `host/sync`: each object is a path from the
//! document, as the Word JavaScript API names them (see WordCraft's `office.sync`).
//!
//! Paragraphs are known by the ids the host hands out (positions in the document: valid until
//! paragraphs are added or removed before them). A range is a path: the document's content, a
//! paragraph's whole range, the selection, or one the host named by id (`Selection.Range`); an
//! id range moves to the new text when its `Text` is set, and stops being usable when text is
//! inserted before it or it is deleted. Properties are read in groups (a font's together, a
//! paragraph's format together) and cached until anything is written; paragraph lists are cached
//! until paragraphs may have been added or removed. Writes are queued and sent together before
//! the next read, at most 500 at a time, and when the macro ends.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use serde_json::{Map, Value as Json, json};

use crate::excel::{color_from_hex, color_to_hex, host_error};
use crate::host::Host;
use crate::vba::value::{Obj, R, Value, VbaError};

const MAX_PENDING: usize = 500;
/// Most paragraphs one collection read takes.
const MAX_ITEMS: usize = 1_000_000;

/// What Word answers for a property that differs across the text (`wdUndefined`).
pub const WD_UNDEFINED: i32 = 9_999_999;
/// `wdToggle`: a font flag set to its opposite.
pub const WD_TOGGLE: i32 = 9_999_998;
/// `wdColorAutomatic`.
pub const WD_COLOR_AUTOMATIC: i32 = -16_777_216;
const WD_STORY: i32 = 6;
const WD_EXTEND: i32 = 1;

/// A range the host named by id. Every copy of the VBA object shares it: setting the range's
/// text moves it to the new text; inserting before it or deleting it leaves it empty (stale).
#[derive(Clone, Debug, PartialEq)]
pub struct Handle(Rc<RefCell<Option<Rc<str>>>>);

impl Handle {
    fn new(id: Rc<str>) -> Handle {
        Handle(Rc::new(RefCell::new(Some(id))))
    }

    fn id(&self) -> R<Rc<str>> {
        self.0.borrow().clone().ok_or_else(|| {
            VbaError::with(5, "This range changed (text was inserted before it, or it was deleted): reach it again (Selection.Range, Paragraphs(i).Range)")
        })
    }

    fn set(&self, id: Option<Rc<str>>) {
        *self.0.borrow_mut() = id;
    }
}

/// The text a range, a font or a paragraph format is of.
#[derive(Clone, Debug, PartialEq)]
pub enum Target {
    /// The whole document (`ActiveDocument.Content`).
    Content,
    /// The start of the document (`ActiveDocument.Range(0, 0)`).
    Start,
    /// A paragraph's whole range.
    Para(Rc<str>),
    /// A range the host named (`Selection.Range`).
    Id(Handle),
    /// The selection, as it is when the range is used.
    Selection,
}

#[derive(Clone, Debug, PartialEq)]
pub enum WdObj {
    Application,
    Documents,
    Document,
    /// The document's paragraphs (`None`), or a range's.
    Paragraphs(Option<Target>),
    Paragraph(Rc<str>),
    Range(Target),
    Font(Target),
    ParagraphFormat(Target),
    Selection,
    /// `BuiltInDocumentProperties`.
    Properties,
    /// One of them, by the bridge's property name (`title`).
    Property(&'static str),
    Styles,
    Style(Rc<str>),
    /// The ribbon button a macro was called from (`Sub Go(control As IRibbonControl)`).
    RibbonControl {
        id: Rc<str>,
        tag: Rc<str>,
    },
}

impl WdObj {
    pub fn type_name(&self) -> &'static str {
        match self {
            WdObj::Application => "Application",
            WdObj::Documents => "Documents",
            WdObj::Document => "Document",
            WdObj::Paragraphs(_) => "Paragraphs",
            WdObj::Paragraph(_) => "Paragraph",
            WdObj::Range(_) => "Range",
            WdObj::Font(_) => "Font",
            WdObj::ParagraphFormat(_) => "ParagraphFormat",
            WdObj::Selection => "Selection",
            WdObj::Properties => "DocumentProperties",
            WdObj::Property(_) => "DocumentProperty",
            WdObj::Styles => "Styles",
            WdObj::Style(_) => "Style",
            WdObj::RibbonControl { .. } => "IRibbonControl",
        }
    }
}

/// `wdColor…` constants (BGR numbers) and their names.
const WD_COLORS: &[(&str, i32)] = &[
    ("black", 0),
    ("white", 16_777_215),
    ("red", 255),
    ("green", 32_768),
    ("brightgreen", 65_280),
    ("blue", 16_711_680),
    ("yellow", 65_535),
    ("turquoise", 16_776_960),
    ("pink", 16_711_935),
    ("darkred", 128),
    ("darkblue", 8_388_608),
    ("teal", 8_421_376),
    ("violet", 8_388_736),
    ("darkyellow", 32_896),
    ("gray50", 8_421_504),
    ("gray25", 12_632_256),
    ("orange", 26_367),
    ("gold", 52_479),
    ("automatic", WD_COLOR_AUTOMATIC),
];

/// The highlight colour indexes (`wdYellow` 7…): the constant's name, the name the bridge
/// takes, and the colour the bridge may answer instead.
const HIGHLIGHTS: &[(i32, &str, &str, &str)] = &[
    (1, "black", "Black", "#000000"),
    (2, "blue", "Blue", "#0000FF"),
    (3, "turquoise", "Turquoise", "#00FFFF"),
    (4, "brightgreen", "BrightGreen", "#00FF00"),
    (5, "pink", "Pink", "#FF00FF"),
    (6, "red", "Red", "#FF0000"),
    (7, "yellow", "Yellow", "#FFFF00"),
    (8, "white", "White", "#FFFFFF"),
    (9, "darkblue", "DarkBlue", "#000080"),
    (10, "teal", "Teal", "#008080"),
    (11, "green", "Green", "#008000"),
    (12, "violet", "Violet", "#800080"),
    (13, "darkred", "DarkRed", "#800000"),
    (14, "darkyellow", "DarkYellow", "#808000"),
    (15, "gray50", "Gray50", "#808080"),
    (16, "gray25", "Gray25", "#C0C0C0"),
];

/// `wdUnderline…`: the constant's name, its value and the bridge's name (the Word JavaScript
/// API's `UnderlineType`).
const UNDERLINES: &[(&str, i32, &str)] = &[
    ("none", 0, "None"),
    ("single", 1, "Single"),
    ("words", 2, "Word"),
    ("double", 3, "Double"),
    ("dotted", 4, "Dotted"),
    ("thick", 6, "Thick"),
    ("dash", 7, "DashLine"),
    ("dotdash", 9, "DotDashLine"),
    ("dotdotdash", 10, "TwoDotDashLine"),
    ("wavy", 11, "Wave"),
];

/// Other spellings of underline styles an app may answer (WordCraft's own command names).
fn underline_number(name: &str) -> Option<i32> {
    let n = name.to_ascii_lowercase();
    if let Some((_, v, _)) = UNDERLINES.iter().find(|(k, _, api)| api.eq_ignore_ascii_case(&n) || *k == n) {
        return Some(*v);
    }
    Some(match n.as_str() {
        "word" => 2,
        "dotline" => 4,
        "dashline" | "dashed" => 7,
        "dotdashline" => 9,
        "twodotdashline" => 10,
        "wave" => 11,
        "mixed" => return None,
        _ => 1,
    })
}

/// `wdStyle…` built-in styles: (constant, the style's name, the bridge's `styleBuiltIn`).
const BUILT_IN_STYLES: &[(i32, &str, &str)] = &[
    (-1, "Normal", "Normal"),
    (-2, "Heading 1", "Heading1"),
    (-3, "Heading 2", "Heading2"),
    (-4, "Heading 3", "Heading3"),
    (-5, "Heading 4", "Heading4"),
    (-6, "Heading 5", "Heading5"),
    (-7, "Heading 6", "Heading6"),
    (-8, "Heading 7", "Heading7"),
    (-9, "Heading 8", "Heading8"),
    (-10, "Heading 9", "Heading9"),
    (-63, "Title", "Title"),
    (-75, "Subtitle", "Subtitle"),
    (-180, "List Paragraph", "ListParagraph"),
    (-181, "Quote", "Quote"),
    (-182, "Intense Quote", "IntenseQuote"),
];

/// The built-in document properties: (VBA name, `wdProperty…` number, the bridge's name).
const PROPERTIES: &[(&str, i32, &str)] =
    &[("Title", 1, "title"), ("Subject", 2, "subject"), ("Author", 3, "author"), ("Keywords", 4, "keywords"), ("Comments", 5, "comments")];

/// Word's constants (`wdAlignParagraphCenter`, `wdColorRed`, `wdYellow`, `wdUndefined`).
pub fn constant(name: &str) -> Option<Value> {
    let l = |x: i32| Some(Value::Long(x));
    let n = name.to_ascii_lowercase();
    if let Some(c) = n.strip_prefix("wdcolor")
        && let Some((_, v)) = WD_COLORS.iter().find(|(k, _)| *k == c)
    {
        return l(*v);
    }
    if let Some(u) = n.strip_prefix("wdunderline")
        && let Some((_, v, _)) = UNDERLINES.iter().find(|(k, _, _)| *k == u)
    {
        return l(*v);
    }
    if let Some(s) = n.strip_prefix("wdstyle")
        && let Some((v, _, _)) = BUILT_IN_STYLES.iter().find(|(_, _, api)| api.eq_ignore_ascii_case(s))
    {
        return l(*v);
    }
    if let Some(p) = n.strip_prefix("wdproperty")
        && let Some((_, v, _)) = PROPERTIES.iter().find(|(k, _, _)| k.eq_ignore_ascii_case(p))
    {
        return l(*v);
    }
    if let Some(h) = n.strip_prefix("wd")
        && let Some((v, _, _, _)) = HIGHLIGHTS.iter().find(|(_, k, _, _)| *k == h)
    {
        return l(*v);
    }
    match n.as_str() {
        "wdalignparagraphleft" => l(0),
        "wdalignparagraphcenter" => l(1),
        "wdalignparagraphright" => l(2),
        "wdalignparagraphjustify" => l(3),
        "wdalignparagraphdistribute" => l(4),
        "wdnohighlight" | "wdauto" => l(0),
        "wdundefined" => l(WD_UNDEFINED),
        "wdtoggle" => l(WD_TOGGLE),
        "wdcharacter" => l(1),
        "wdword" => l(2),
        "wdsentence" => l(3),
        "wdparagraph" => l(4),
        "wdline" => l(5),
        "wdstory" => l(WD_STORY),
        "wdmove" => l(0),
        "wdextend" => l(WD_EXTEND),
        _ => None,
    }
}

fn p(name: &str) -> Json {
    json!({ "p": name })
}

fn m(name: &str, args: Json) -> Json {
    json!({ "m": name, "a": args })
}

fn doc_path() -> Vec<Json> {
    vec![p("document")]
}

fn body_path() -> Vec<Json> {
    vec![p("document"), p("body")]
}

fn paragraphs_path() -> Vec<Json> {
    vec![p("document"), p("paragraphs")]
}

fn para_path(id: &str) -> Vec<Json> {
    let mut v = paragraphs_path();
    v.push(m("getById", json!([id])));
    v
}

fn selection_path() -> Vec<Json> {
    vec![p("document"), m("getSelection", json!([]))]
}

fn body_range(location: &str) -> Vec<Json> {
    let mut v = body_path();
    v.push(m("getRange", json!([location])));
    v
}

fn target_path(t: &Target) -> R<Vec<Json>> {
    Ok(match t {
        Target::Content => body_range("Whole"),
        Target::Start => body_range("Start"),
        Target::Para(id) => {
            let mut v = para_path(id);
            v.push(m("getRange", json!(["Whole"])));
            v
        }
        Target::Id(h) => vec![p("document"), m("getRangeById", json!([h.id()?.as_ref()]))],
        Target::Selection => selection_path(),
    })
}

fn with(mut path: Vec<Json>, step: Json) -> Vec<Json> {
    path.push(step);
    path
}

fn key_of(path: &[Json], props: &[&str]) -> String {
    format!("{}|{}", Json::Array(path.to_vec()), props.join(","))
}

/// An id as the bridge sends it (a string, or a number).
fn id_str(j: &Json) -> Option<Rc<str>> {
    match j {
        Json::String(s) => Some(s.as_str().into()),
        Json::Number(n) => Some(n.to_string().into()),
        _ => None,
    }
}

fn ids_of(out: &Map<String, Json>) -> Vec<Rc<str>> {
    out.get("items")
        .and_then(Json::as_array)
        .into_iter()
        .flatten()
        .filter_map(Json::as_object)
        .take(MAX_ITEMS)
        .filter_map(|i| i.get("id").and_then(id_str))
        .collect()
}

fn obj(x: WdObj) -> R<Value> {
    Ok(Value::Object(Obj::Wd(x)))
}

fn finite(v: &Value) -> R<f64> {
    let f = v.to_f64()?;
    if f.is_finite() { Ok(f) } else { Err(VbaError::new(6)) }
}

fn in_range(v: &Value, lo: f64, hi: f64, what: &str) -> R<f64> {
    let f = finite(v)?;
    if (lo..=hi).contains(&f) { Ok(f) } else { Err(VbaError::with(5, format!("{what} runs from {lo} to {hi} points"))) }
}

/// The `n`th (1-based) of `len` things.
fn position(v: &Value, len: usize) -> R<usize> {
    let i = v.to_i64()?;
    usize::try_from(i).ok().filter(|i| (1..=len).contains(i)).map(|i| i - 1).ok_or(VbaError::new(9))
}

/// Argument `i` or the one named `key`, unless left out.
fn arg<'v>(args: &'v [Value], named: &'v [(String, Value)], i: usize, key: &str) -> Option<&'v Value> {
    args.get(i).filter(|v| !matches!(v, Value::Missing)).or_else(|| named.iter().find(|(k, _)| k.eq_ignore_ascii_case(key)).map(|(_, v)| v))
}

fn need<'v>(args: &'v [Value], named: &'v [(String, Value)], i: usize, key: &str) -> R<&'v Value> {
    arg(args, named, i, key).ok_or(VbaError::new(449))
}

/// Text as Word's macros see it: paragraphs end with a carriage return (`vbCr`).
fn text_out(s: &str) -> String {
    s.replace("\r\n", "\n").replace('\n', "\r")
}

/// Text as the host takes it: paragraphs separated by `\n`.
fn text_in(s: &str) -> String {
    s.replace("\r\n", "\n").replace('\r', "\n")
}

/// A paragraph's (or the document's) text with its final paragraph mark, as Word reads it.
fn with_mark(s: &str) -> String {
    let mut t = text_out(s.strip_suffix('\n').unwrap_or(s));
    t.push('\r');
    t
}

/// A flag as Word reads it: `True` (-1), `False` (0), `wdUndefined` when mixed.
fn flag(j: &Json) -> Value {
    Value::Long(match j.as_bool() {
        Some(true) => -1,
        Some(false) => 0,
        None => WD_UNDEFINED,
    })
}

fn number(j: &Json) -> Value {
    j.as_f64().map(Value::Double).unwrap_or(Value::Long(WD_UNDEFINED))
}

fn alignment_number(j: &Json) -> i32 {
    match j.as_str().map(str::to_ascii_lowercase).as_deref() {
        Some("left") => 0,
        Some("centered" | "center") => 1,
        Some("right") => 2,
        Some("justified" | "justify") => 3,
        _ => WD_UNDEFINED,
    }
}

fn alignment_name(v: &Value) -> R<&'static str> {
    match v.to_i32()? {
        0 => Ok("Left"),
        1 => Ok("Centered"),
        2 => Ok("Right"),
        3 => Ok("Justified"),
        4 => Err(VbaError::with(5, "wdAlignParagraphDistribute isn't supported yet")),
        _ => Err(VbaError::new(5)),
    }
}

fn color_value(j: &Json) -> Value {
    match j.as_str() {
        Some(s) if s.eq_ignore_ascii_case("auto") || s.eq_ignore_ascii_case("automatic") => Value::Long(WD_COLOR_AUTOMATIC),
        Some(s) => Value::Long(color_from_hex(s).unwrap_or(WD_UNDEFINED)),
        None => Value::Long(WD_UNDEFINED),
    }
}

fn color_arg(v: &Value) -> R<String> {
    let n = v.to_i32()?;
    if n == WD_COLOR_AUTOMATIC {
        return Err(VbaError::with(5, "wdColorAutomatic isn't supported yet (give the colour: wdColorBlack, RGB(…))"));
    }
    if !(0..=0xFF_FFFF).contains(&n) {
        return Err(VbaError::with(5, "Colours run from 0 to 16777215 (RGB(255, 255, 255))"));
    }
    Ok(color_to_hex(n))
}

fn highlight_value(j: &Json) -> Value {
    match j.as_str() {
        None => Value::Long(0),
        Some(s) => {
            let k = s.replace([' ', '-', '%'], "");
            Value::Long(
                HIGHLIGHTS
                    .iter()
                    .find(|(_, name, api, hex)| k.eq_ignore_ascii_case(name) || k.eq_ignore_ascii_case(api) || s.eq_ignore_ascii_case(hex))
                    .map(|(v, ..)| *v)
                    .unwrap_or(if k.eq_ignore_ascii_case("none") || k.is_empty() { 0 } else { WD_UNDEFINED }),
            )
        }
    }
}

fn highlight_arg(v: &Value) -> R<Json> {
    let n = v.to_i32()?;
    if n == 0 {
        return Ok(Json::Null);
    }
    HIGHLIGHTS
        .iter()
        .find(|(x, ..)| *x == n)
        .map(|(_, _, api, _)| json!(api))
        .ok_or_else(|| VbaError::with(5, "Highlight colour indexes run from wdNoHighlight (0) to wdGray25 (16)"))
}

/// A style given to `Style =`: a name, or a `wdStyle…` constant. `(prop, value)` to set.
fn style_arg(v: &Value, built_in_prop: bool) -> R<(&'static str, String)> {
    match v {
        Value::Str(s) => Ok(("style", s.to_string())),
        Value::Object(Obj::Wd(WdObj::Style(s))) => Ok(("style", s.to_string())),
        other => {
            let n = other.to_i32()?;
            let (_, name, api) =
                BUILT_IN_STYLES.iter().find(|(x, _, _)| *x == n).ok_or_else(|| VbaError::with(5, format!("Built-in style {n} isn't supported yet")))?;
            Ok(if built_in_prop { ("styleBuiltIn", api.to_string()) } else { ("style", name.to_string()) })
        }
    }
}

type Args<'a> = &'a [Value];

/// The macro's view of the document.
#[derive(Default)]
pub struct Word {
    pending: Vec<Json>,
    /// Requests sent (for tests and tuning).
    pub requests: usize,
    /// The document's paragraphs, until paragraphs may have been added or removed.
    paras: Option<Vec<Rc<str>>>,
    /// A range's paragraphs, until anything is written.
    range_paras: HashMap<String, Vec<Rc<str>>>,
    props: HashMap<String, Map<String, Json>>,
}

const FONT: &[&str] = &["bold", "italic", "underline", "size", "name", "color", "highlightColor"];
const PARA: &[&str] = &["style", "alignment", "firstLineIndent", "leftIndent", "rightIndent", "spaceAfter", "spaceBefore", "lineSpacing"];
const RANGE: &[&str] = &["text", "style"];
const DOC_PROPS: &[&str] = &["title", "subject", "author", "keywords", "comments"];

impl Word {
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

    fn load(&mut self, host: &mut dyn Host, path: Vec<Json>, props: &[&str]) -> R<Map<String, Json>> {
        let key = key_of(&path, props);
        if let Some(v) = self.props.get(&key) {
            return Ok(v.clone());
        }
        let mut r = self.sync(host, vec![json!({"op": "load", "path": path, "props": props})])?;
        let out = match r.pop() {
            Some(Json::Object(o)) => o,
            _ => Map::new(),
        };
        self.props.insert(key, out.clone());
        Ok(out)
    }

    fn prop(&mut self, host: &mut dyn Host, path: Vec<Json>, group: &[&str], prop: &str) -> R<Json> {
        Ok(self.load(host, path, group)?.get(prop).cloned().unwrap_or(Json::Null))
    }

    /// Queues a write; what was read no longer holds (and the paragraph list neither, when
    /// paragraphs may come or go).
    fn write(&mut self, host: &mut dyn Host, action: Json, structural: bool) -> R<()> {
        self.pending.push(action);
        self.changed(structural);
        if self.pending.len() >= MAX_PENDING {
            self.flush(host)?;
        }
        Ok(())
    }

    fn changed(&mut self, structural: bool) {
        self.props.clear();
        self.range_paras.clear();
        if structural {
            self.paras = None;
        }
    }

    fn set(&mut self, host: &mut dyn Host, path: Vec<Json>, prop: &str, value: Json) -> R<()> {
        self.write(host, json!({"op": "set", "path": path, "prop": prop, "value": value}), false)
    }

    fn call(&mut self, host: &mut dyn Host, path: Vec<Json>, method: &str, args: Json, structural: bool) -> R<()> {
        self.write(host, json!({"op": "call", "path": path, "method": method, "args": args}), structural)
    }

    /// A call whose result is needed, sent now.
    fn now(&mut self, host: &mut dyn Host, path: Vec<Json>, method: &str, args: Json) -> R<Json> {
        let r = self.sync(host, vec![json!({"op": "call", "path": path, "method": method, "args": args})]);
        self.changed(true);
        Ok(r?.pop().unwrap_or(Json::Null))
    }

    fn created_id(out: &Json, what: &str) -> R<Rc<str>> {
        out.get("id").and_then(id_str).ok_or_else(|| VbaError::with(1004, format!("{what}: the app didn't say what it created")))
    }

    /// The document's paragraphs (`None`) or a range's.
    fn paragraph_ids(&mut self, host: &mut dyn Host, scope: &Option<Target>) -> R<Vec<Rc<str>>> {
        match scope {
            None => {
                if let Some(p) = &self.paras {
                    return Ok(p.clone());
                }
                let out = self.load(host, paragraphs_path(), &["items"])?;
                let ids = ids_of(&out);
                self.paras = Some(ids.clone());
                Ok(ids)
            }
            Some(Target::Para(id)) => Ok(vec![id.clone()]),
            Some(t) => {
                let path = with(target_path(t)?, p("paragraphs"));
                let key = key_of(&path, &["items"]);
                if let Some(p) = self.range_paras.get(&key) {
                    return Ok(p.clone());
                }
                let out = self.load(host, path, &["items"])?;
                let ids = ids_of(&out);
                self.range_paras.insert(key, ids.clone());
                Ok(ids)
            }
        }
    }

    /// The paragraph `step` (+1 next, -1 previous) from `id`, or `Nothing`.
    fn neighbour(&mut self, host: &mut dyn Host, id: &Rc<str>, step: i64) -> R<Value> {
        let all = self.paragraph_ids(host, &None)?;
        if let Some(i) = all.iter().position(|x| x == id) {
            let j = i64::try_from(i).unwrap_or(i64::MAX).saturating_add(step);
            return match usize::try_from(j).ok().and_then(|j| all.get(j)) {
                Some(n) => obj(WdObj::Paragraph(n.clone())),
                None => Ok(Value::Nothing),
            };
        }
        // Not in the list (a paragraph inside something the list leaves out): ask the app,
        // alone, so that its "no such paragraph" costs no queued write.
        self.flush(host)?;
        let method = if step > 0 { "getNext" } else { "getPrevious" };
        match self.load(host, with(para_path(id), m(method, json!([]))), &["id"]) {
            Ok(out) => match out.get("id").and_then(id_str) {
                Some(n) => obj(WdObj::Paragraph(n)),
                None => Ok(Value::Nothing),
            },
            Err(e) if e.number == 9 => Ok(Value::Nothing),
            Err(e) => Err(e),
        }
    }

    fn doc_name(&mut self, host: &mut dyn Host) -> R<Value> {
        // Alone, so that an app without the property costs no queued write.
        self.flush(host)?;
        let name = match self.load(host, doc_path(), &["name"]) {
            Ok(out) => out.get("name").and_then(Json::as_str).filter(|n| !n.is_empty()).map(str::to_string),
            Err(_) => None,
        };
        Ok(Value::str(name.unwrap_or_else(|| "Document1".into())))
    }

    fn property(key: &Value) -> R<&'static str> {
        let found = match key {
            Value::Str(s) => PROPERTIES.iter().find(|(n, _, _)| n.eq_ignore_ascii_case(s.trim())),
            v => {
                let n = v.to_i32()?;
                PROPERTIES.iter().find(|(_, x, _)| *x == n)
            }
        };
        found
            .map(|(_, _, api)| *api)
            .ok_or_else(|| VbaError::with(5, "Only the Title, Subject, Author, Keywords and Comments document properties are supported yet"))
    }

    fn style_of(key: &Value) -> R<Rc<str>> {
        match key {
            Value::Str(s) => Ok(s.clone()),
            v => {
                let n = v.to_i32()?;
                BUILT_IN_STYLES
                    .iter()
                    .find(|(x, _, _)| *x == n)
                    .map(|(_, name, _)| Rc::from(*name))
                    .ok_or_else(|| VbaError::with(5, format!("Built-in style {n} isn't supported yet")))
            }
        }
    }

    fn style_value(j: &Json) -> Value {
        match j.as_str() {
            Some(s) => Value::Object(Obj::Wd(WdObj::Style(s.into()))),
            None => Value::Null,
        }
    }

    /// A range's text as Word reads it.
    fn range_text(&mut self, host: &mut dyn Host, t: &Target) -> R<String> {
        let path = target_path(t)?;
        let text = self.prop(host, path, RANGE, "text")?;
        let text = text.as_str().unwrap_or("");
        Ok(match t {
            Target::Content | Target::Para(_) => with_mark(text),
            _ => text_out(text),
        })
    }

    /// `range.Text = text`.
    fn set_text(&mut self, host: &mut dyn Host, t: &Target, text: &str) -> R<()> {
        let mut text = text_in(text);
        // The paragraph's (or the document's last) mark stays.
        if matches!(t, Target::Para(_) | Target::Content) && text.ends_with('\n') {
            text.pop();
        }
        let structural = text.contains('\n') || matches!(t, Target::Content);
        match t {
            Target::Id(h) => {
                let out = self.now(host, target_path(t)?, "insertText", json!([text, "Replace"]))?;
                h.set(out.get("id").and_then(id_str));
                Ok(())
            }
            Target::Selection => self.replace_selection(host, &text, "Select"),
            _ => self.call(host, target_path(t)?, "insertText", json!([text, "Replace"]), structural),
        }
    }

    /// Replaces the selection with `text` and selects it (`Select`) or puts the caret after it
    /// (`End`), in one sync with the queued writes.
    fn replace_selection(&mut self, host: &mut dyn Host, text: &str, mode: &str) -> R<()> {
        let i = self.pending.len();
        self.pending.push(json!({"op": "call", "path": selection_path(), "method": "insertText", "args": [text, "Replace"]}));
        let inserted = vec![p("document"), m("getRangeById", json!([{"result": i}]))];
        self.pending.push(json!({"op": "call", "path": inserted, "method": "select", "args": [mode]}));
        self.changed(text.contains('\n'));
        if self.pending.len() >= MAX_PENDING {
            self.flush(host)?;
        }
        Ok(())
    }

    fn insert(&mut self, host: &mut dyn Host, t: &Target, text: &str, after: bool) -> R<()> {
        let text = text_in(text);
        let structural = text.contains('\n');
        self.call(host, target_path(t)?, "insertText", json!([text, if after { "End" } else { "Start" }]), structural)?;
        if let (Target::Id(h), false) = (t, after) {
            h.set(None);
        }
        Ok(())
    }

    fn insert_paragraph(&mut self, host: &mut dyn Host, t: &Target, after: bool) -> R<()> {
        let path = match t {
            Target::Para(id) => para_path(id),
            other => target_path(other)?,
        };
        self.call(host, path, "insertParagraph", json!(["", if after { "After" } else { "Before" }]), true)?;
        if let (Target::Id(h), false) = (t, after) {
            h.set(None);
        }
        Ok(())
    }

    fn font_prop(&mut self, host: &mut dyn Host, t: &Target, prop: &str) -> R<Json> {
        let path = with(target_path(t)?, p("font"));
        self.prop(host, path, FONT, prop)
    }

    fn set_font(&mut self, host: &mut dyn Host, t: &Target, prop: &str, value: Json) -> R<()> {
        let path = with(target_path(t)?, p("font"));
        self.set(host, path, prop, value)
    }

    /// A paragraph format property of the range's paragraphs: their common value, or
    /// `wdUndefined`.
    fn format_prop(&mut self, host: &mut dyn Host, t: &Target, prop: &str) -> R<Option<Json>> {
        let ids = self.paragraph_ids(host, &Some(t.clone()))?;
        // The paragraphs not read yet, in one request (a few thousand at a time).
        let missing: Vec<Vec<Json>> = ids.iter().map(|id| para_path(id)).filter(|path| !self.props.contains_key(&key_of(path, PARA))).collect();
        for chunk in missing.chunks(MAX_PENDING * 4) {
            let actions = chunk.iter().map(|path| json!({"op": "load", "path": path, "props": PARA})).collect();
            let out = self.sync(host, actions)?;
            for (path, r) in chunk.iter().zip(out) {
                if let Json::Object(o) = r {
                    self.props.insert(key_of(path, PARA), o);
                }
            }
        }
        let mut common: Option<Json> = None;
        for id in &ids {
            let v = self.prop(host, para_path(id), PARA, prop)?;
            match &common {
                None => common = Some(v),
                Some(c) if *c == v => {}
                Some(_) => return Ok(None),
            }
        }
        Ok(common)
    }

    fn set_format(&mut self, host: &mut dyn Host, t: &Target, prop: &str, value: Json) -> R<()> {
        let ids = self.paragraph_ids(host, &Some(t.clone()))?;
        for id in ids {
            self.set(host, para_path(&id), prop, value.clone())?;
        }
        Ok(())
    }

    fn paragraph_format(&mut self, host: &mut dyn Host, t: &Target, n: &str) -> R<Value> {
        let prop = match n {
            "alignment" => "alignment",
            "spaceafter" => "spaceAfter",
            "spacebefore" => "spaceBefore",
            "leftindent" => "leftIndent",
            "rightindent" => "rightIndent",
            "firstlineindent" => "firstLineIndent",
            "linespacing" => "lineSpacing",
            _ => return Err(VbaError::with(438, format!("ParagraphFormat.{n} isn't supported yet"))),
        };
        let v = self.format_prop(host, t, prop)?;
        Ok(match (prop, v) {
            (_, None) => Value::Long(WD_UNDEFINED),
            ("alignment", Some(a)) => Value::Long(alignment_number(&a)),
            (_, Some(x)) => number(&x),
        })
    }

    fn set_paragraph_format(&mut self, host: &mut dyn Host, t: &Target, n: &str, value: &Value) -> R<()> {
        let (prop, v) = match n {
            "alignment" => ("alignment", json!(alignment_name(value)?)),
            "spaceafter" => ("spaceAfter", json!(in_range(value, 0.0, 1584.0, "Spacing")?)),
            "spacebefore" => ("spaceBefore", json!(in_range(value, 0.0, 1584.0, "Spacing")?)),
            "linespacing" => ("lineSpacing", json!(in_range(value, 0.0, 1584.0, "Line spacing")?)),
            "leftindent" => ("leftIndent", json!(in_range(value, -1584.0, 1584.0, "An indent")?)),
            "rightindent" => ("rightIndent", json!(in_range(value, -1584.0, 1584.0, "An indent")?)),
            "firstlineindent" => ("firstLineIndent", json!(in_range(value, -1584.0, 1584.0, "An indent")?)),
            _ => return Err(VbaError::with(438, format!("ParagraphFormat.{n} can't be set (or isn't supported yet)"))),
        };
        self.set_format(host, t, prop, v)
    }

    /// A font flag written: `True`/`False`, or `wdToggle`.
    fn flag_arg(&mut self, host: &mut dyn Host, t: &Target, prop: &str, value: &Value) -> R<bool> {
        match value.to_i32() {
            Ok(WD_TOGGLE) => Ok(self.font_prop(host, t, prop)?.as_bool() != Some(true)),
            Ok(WD_UNDEFINED) => Err(VbaError::new(5)),
            _ => value.to_bool(),
        }
    }

    /// `HomeKey`/`EndKey`: only to the start or end of the document (`Unit:=wdStory`).
    fn story_key(&mut self, host: &mut dyn Host, args: Args, named: &[(String, Value)], end: bool, name: &str) -> R<Value> {
        let unit = arg(args, named, 0, "unit").map(Value::to_i32).transpose()?.unwrap_or(5);
        let extend = arg(args, named, 1, "extend").map(Value::to_i32).transpose()?.unwrap_or(0);
        if unit != WD_STORY || extend == WD_EXTEND {
            return Err(VbaError::with(5, format!("Selection.{name} supports only Unit:=wdStory (moving to the start or end of the document) yet")));
        }
        self.call(host, body_range(if end { "End" } else { "Start" }), "select", json!([]), false)?;
        Ok(Value::Long(0))
    }

    /// `object.name` or `object.name(args)`: a property's value or a method's result.
    pub fn member(&mut self, host: &mut dyn Host, o: &WdObj, name: &str, args: Args, named: &[(String, Value)]) -> R<Value> {
        let n = name.to_ascii_lowercase();
        let item = |me: &mut Self, host: &mut dyn Host, x: WdObj| if args.is_empty() { obj(x) } else { me.member(host, &x, "item", args, named) };
        match (o, n.as_str()) {
            // Application.
            (WdObj::Application, "activedocument" | "thisdocument") => obj(WdObj::Document),
            (WdObj::Application, "documents") => item(self, host, WdObj::Documents),
            (WdObj::Application, "selection") => obj(WdObj::Selection),
            (WdObj::Application, "name") => Ok(Value::str("WordCraft")),
            (WdObj::Application, "version") => Ok(Value::str("16.0")),
            (WdObj::Application, "screenupdating") => Ok(Value::Bool(true)),
            (WdObj::Application, "displayalerts") => Ok(Value::Long(-1)),
            (_, "application") => obj(WdObj::Application),
            (WdObj::RibbonControl { id, .. }, "id") => Ok(Value::Str(id.clone())),
            (WdObj::RibbonControl { tag, .. }, "tag") => Ok(Value::Str(tag.clone())),
            (WdObj::RibbonControl { .. }, "context") => obj(WdObj::Document),

            (WdObj::Documents, "count") => Ok(Value::Long(1)),
            (WdObj::Documents, "item") => {
                let key = need(args, named, 0, "index")?;
                if let Value::Str(s) = key {
                    let name = self.doc_name(host)?.to_str()?;
                    if !s.trim().eq_ignore_ascii_case(&name) {
                        return Err(VbaError::with(9, format!("Documents.Item: no document named '{s}'")));
                    }
                } else {
                    position(key, 1)?;
                }
                obj(WdObj::Document)
            }
            (WdObj::Documents | WdObj::Document, "parent") => obj(WdObj::Application),

            // The document.
            (WdObj::Document, "content") => obj(WdObj::Range(Target::Content)),
            (WdObj::Document, "paragraphs") => item(self, host, WdObj::Paragraphs(None)),
            (WdObj::Document, "range") => {
                let start = arg(args, named, 0, "start");
                let end = arg(args, named, 1, "end");
                match (start, end) {
                    (None, None) => obj(WdObj::Range(Target::Content)),
                    (Some(s), Some(e)) if s.to_i64()? == 0 && e.to_i64()? == 0 => obj(WdObj::Range(Target::Start)),
                    _ => Err(VbaError::with(
                        5,
                        "Document.Range(Start, End) with character positions isn't supported yet: use Content, Range(0, 0), Paragraphs(i).Range or Selection.Range",
                    )),
                }
            }
            (WdObj::Document, "name" | "fullname") => self.doc_name(host),
            (WdObj::Document, "builtindocumentproperties") => item(self, host, WdObj::Properties),
            (WdObj::Document, "styles") => item(self, host, WdObj::Styles),
            (WdObj::Document, "save") => {
                self.call(host, doc_path(), "save", json!([]), false)?;
                Ok(Value::Empty)
            }
            (WdObj::Document, "activate") => Ok(Value::Empty),
            (WdObj::Document, "select") => self.member(host, &WdObj::Selection, "wholestory", &[], &[]),

            (WdObj::Properties, "item") => obj(WdObj::Property(Self::property(need(args, named, 0, "index")?)?)),
            (WdObj::Properties, "count") => Ok(Value::Long(PROPERTIES.len() as i32)),
            (WdObj::Property(api), "value") => {
                let v = self.prop(host, with(doc_path(), p("properties")), DOC_PROPS, api)?;
                Ok(Value::str(v.as_str().unwrap_or("")))
            }
            (WdObj::Property(api), "name") => Ok(Value::str(PROPERTIES.iter().find(|(_, _, a)| a == api).map(|(n, _, _)| *n).unwrap_or(""))),
            (WdObj::Properties | WdObj::Styles | WdObj::Paragraphs(_), "parent") => obj(WdObj::Document),

            (WdObj::Styles, "item") => obj(WdObj::Style(Self::style_of(need(args, named, 0, "index")?)?)),
            (WdObj::Style(s), "name" | "namelocal") => Ok(Value::Str(s.clone())),

            // Paragraphs.
            (WdObj::Paragraphs(scope), "count") => Ok(Value::Long(i32::try_from(self.paragraph_ids(host, scope)?.len()).unwrap_or(i32::MAX))),
            (WdObj::Paragraphs(scope), "item") => {
                let ids = self.paragraph_ids(host, scope)?;
                let i = position(need(args, named, 0, "index")?, ids.len())?;
                obj(WdObj::Paragraph(ids.get(i).cloned().ok_or(VbaError::new(9))?))
            }
            (WdObj::Paragraphs(scope), "first" | "last") => {
                let ids = self.paragraph_ids(host, scope)?;
                let id = if n == "first" { ids.first() } else { ids.last() };
                obj(WdObj::Paragraph(id.cloned().ok_or_else(|| VbaError::with(91, "There are no paragraphs here"))?))
            }
            (WdObj::Paragraphs(_), "add") => {
                let out = self.now(host, body_path(), "insertParagraph", json!(["", "End"]))?;
                obj(WdObj::Paragraph(Self::created_id(&out, "Paragraphs.Add")?))
            }

            // A paragraph.
            (WdObj::Paragraph(id), "range") => obj(WdObj::Range(Target::Para(id.clone()))),
            (WdObj::Paragraph(id), "format" | "paragraphformat") => obj(WdObj::ParagraphFormat(Target::Para(id.clone()))),
            (WdObj::Paragraph(id), "style") => Ok(Self::style_value(&self.prop(host, para_path(id), PARA, "style")?)),
            (WdObj::Paragraph(id), "next") => self.neighbour(host, id, 1),
            (WdObj::Paragraph(id), "previous") => self.neighbour(host, id, -1),
            (WdObj::Paragraph(_), "parent") => obj(WdObj::Document),
            (WdObj::Paragraph(id), p) => self.paragraph_format(host, &Target::Para(id.clone()), p),
            (WdObj::ParagraphFormat(t), p) => self.paragraph_format(host, t, p),

            // Ranges.
            (WdObj::Range(t), "text") => Ok(Value::str(self.range_text(host, t)?)),
            (WdObj::Range(t), "insertafter" | "insertbefore") => {
                let text = need(args, named, 0, "text")?.to_str()?;
                self.insert(host, t, &text, n == "insertafter")?;
                Ok(Value::Empty)
            }
            (WdObj::Range(t), "insertparagraphafter" | "insertparagraphbefore" | "insertparagraph") => {
                self.insert_paragraph(host, t, n != "insertparagraphbefore")?;
                Ok(Value::Empty)
            }
            (WdObj::Range(t), "font") => obj(WdObj::Font(t.clone())),
            (WdObj::Range(t), "paragraphs") => item(self, host, WdObj::Paragraphs(Some(t.clone()))),
            (WdObj::Range(t), "paragraphformat") => obj(WdObj::ParagraphFormat(t.clone())),
            (WdObj::Range(t), "style") => {
                let path = target_path(t)?;
                Ok(Self::style_value(&self.prop(host, path, RANGE, "style")?))
            }
            (WdObj::Range(t), "highlightcolorindex") => Ok(highlight_value(&self.font_prop(host, t, "highlightColor")?)),
            (WdObj::Range(t), "bold" | "italic") => Ok(flag(&self.font_prop(host, t, &n)?)),
            (WdObj::Range(t), "select") => {
                self.call(host, target_path(t)?, "select", json!([]), false)?;
                Ok(Value::Empty)
            }
            (WdObj::Range(t), "delete") => {
                self.call(host, target_path(t)?, "delete", json!([]), true)?;
                if let Target::Id(h) = t {
                    h.set(None);
                }
                Ok(Value::Empty)
            }
            (WdObj::Range(_), "parent" | "document") => obj(WdObj::Document),
            (WdObj::Range(_) | WdObj::Selection, "start" | "end") => {
                Err(VbaError::with(438, format!("{}.{name}: character positions aren't supported yet", o.type_name())))
            }

            // Fonts.
            (WdObj::Font(t), p @ ("bold" | "italic")) => Ok(flag(&self.font_prop(host, t, p)?)),
            (WdObj::Font(t), "size") => Ok(number(&self.font_prop(host, t, "size")?)),
            (WdObj::Font(t), "name") => Ok(Value::str(self.font_prop(host, t, "name")?.as_str().unwrap_or(""))),
            (WdObj::Font(t), "color") => Ok(color_value(&self.font_prop(host, t, "color")?)),
            (WdObj::Font(t), "underline") => Ok(Value::Long(self.font_prop(host, t, "underline")?.as_str().and_then(underline_number).unwrap_or(WD_UNDEFINED))),
            (WdObj::Font(_), "parent") => obj(WdObj::Document),

            // The selection.
            (WdObj::Selection, "text") => Ok(Value::str(self.range_text(host, &Target::Selection)?)),
            (WdObj::Selection, "typetext") => {
                let text = need(args, named, 0, "text")?.to_str()?;
                self.replace_selection(host, &text_in(&text), "End")?;
                Ok(Value::Empty)
            }
            (WdObj::Selection, "typeparagraph") => {
                self.replace_selection(host, "\n", "End")?;
                Ok(Value::Empty)
            }
            (WdObj::Selection, "range") => {
                let path = selection_path();
                self.flush(host)?;
                let out = self.load(host, path, &["id"])?;
                let id = out.get("id").and_then(id_str).ok_or_else(|| VbaError::with(1004, "Selection.Range: the app didn't name the selected range"))?;
                obj(WdObj::Range(Target::Id(Handle::new(id))))
            }
            (WdObj::Selection, "homekey") => self.story_key(host, args, named, false, name),
            (WdObj::Selection, "endkey") => self.story_key(host, args, named, true, name),
            (WdObj::Selection, "wholestory") => {
                self.call(host, body_range("Whole"), "select", json!([]), false)?;
                Ok(Value::Empty)
            }
            (WdObj::Selection, "document") => obj(WdObj::Document),
            (WdObj::Selection, "parent") => obj(WdObj::Application),
            (WdObj::Selection, _) => self.member(host, &WdObj::Range(Target::Selection), name, args, named),

            (o, _) => Err(VbaError::with(438, format!("{}.{name} isn't supported yet", o.type_name()))),
        }
    }

    /// `object.name = value` (`object.name(args) = value` for `BuiltInDocumentProperties(…)`).
    pub fn assign(&mut self, host: &mut dyn Host, o: &WdObj, name: &str, args: Args, value: &Value) -> R<()> {
        let n = name.to_ascii_lowercase();
        if !args.is_empty() {
            return match (o, n.as_str()) {
                (WdObj::Document, "builtindocumentproperties") | (WdObj::Properties, "item") => {
                    let api = Self::property(args.first().ok_or(VbaError::new(449))?)?;
                    self.assign(host, &WdObj::Property(api), "value", &[], value)
                }
                _ => Err(VbaError::new(450)),
            };
        }
        match (o, n.as_str()) {
            (WdObj::Application, "screenupdating" | "displayalerts") => Ok(()),
            (WdObj::Property(api), "value") => self.set(host, with(doc_path(), p("properties")), api, json!(value.to_str()?)),
            (WdObj::Paragraph(id), "style") => {
                let (prop, v) = style_arg(value, true)?;
                self.set(host, para_path(id), prop, json!(v))
            }
            (WdObj::Paragraph(id), p) => self.set_paragraph_format(host, &Target::Para(id.clone()), p, value),
            (WdObj::ParagraphFormat(t), p) => self.set_paragraph_format(host, t, p, value),
            (WdObj::Range(t), "text") => self.set_text(host, t, &value.to_str()?),
            (WdObj::Range(t), "style") => {
                let (prop, v) = style_arg(value, false)?;
                self.set(host, target_path(t)?, prop, json!(v))
            }
            (WdObj::Range(t), "highlightcolorindex") => {
                let v = highlight_arg(value)?;
                self.set_font(host, t, "highlightColor", v)
            }
            (WdObj::Range(t), "bold" | "italic") => self.assign(host, &WdObj::Font(t.clone()), name, &[], value),
            (WdObj::Font(t), p @ ("bold" | "italic")) => {
                let on = self.flag_arg(host, t, p, value)?;
                self.set_font(host, t, p, json!(on))
            }
            (WdObj::Font(t), "size") => {
                let v = finite(value)?;
                if !(1.0..=1638.0).contains(&v) {
                    return Err(VbaError::with(5, "Font sizes run from 1 to 1638 points"));
                }
                self.set_font(host, t, "size", json!(v))
            }
            (WdObj::Font(t), "name") => {
                let name = value.to_str()?;
                if name.trim().is_empty() {
                    return Err(VbaError::with(5, "A font needs a name"));
                }
                self.set_font(host, t, "name", json!(name))
            }
            (WdObj::Font(t), "color") => {
                let hex = color_arg(value)?;
                self.set_font(host, t, "color", json!(hex))
            }
            (WdObj::Font(t), "underline") => {
                let v = value.to_i32()?;
                let v = if v == -1 { 1 } else { v };
                let api = UNDERLINES
                    .iter()
                    .find(|(_, x, _)| *x == v)
                    .map(|(_, _, api)| *api)
                    .ok_or_else(|| VbaError::with(5, format!("Underline style {v} isn't supported yet")))?;
                self.set_font(host, t, "underline", json!(api))
            }
            (WdObj::Selection, "text") => self.set_text(host, &Target::Selection, &value.to_str()?),
            (WdObj::Selection, _) => self.assign(host, &WdObj::Range(Target::Selection), name, &[], value),
            (o, _) => Err(VbaError::with(438, format!("{}.{name} can't be set (or isn't supported yet)", o.type_name()))),
        }
    }

    /// `object = value` where the object has a default property (`Range.Text`).
    pub fn assign_default(&mut self, host: &mut dyn Host, o: &WdObj, value: &Value) -> R<()> {
        match o {
            WdObj::Range(_) | WdObj::Selection => self.assign(host, o, "text", &[], value),
            WdObj::Property(_) => self.assign(host, o, "value", &[], value),
            _ => Err(VbaError::new(438)),
        }
    }

    /// `object(args)`: the default member (`Paragraphs(1)` is `Paragraphs.Item(1)`).
    pub fn default_call(&mut self, host: &mut dyn Host, o: &WdObj, args: Args) -> R<Value> {
        match o {
            WdObj::Paragraphs(_) | WdObj::Documents | WdObj::Properties | WdObj::Styles => self.member(host, o, "item", args, &[]),
            _ => Err(VbaError::new(438)),
        }
    }

    /// The objects `For Each` walks: paragraphs.
    pub fn items(&mut self, host: &mut dyn Host, o: &WdObj) -> R<Vec<Value>> {
        Ok(match o {
            WdObj::Paragraphs(scope) => self.paragraph_ids(host, scope)?.into_iter().map(|id| Value::Object(Obj::Wd(WdObj::Paragraph(id)))).collect(),
            WdObj::Documents => vec![Value::Object(Obj::Wd(WdObj::Document))],
            _ => return Err(VbaError::new(438)),
        })
    }

    /// The value an object stands for where a value is wanted (a range's `Text`).
    pub fn default_value(&mut self, host: &mut dyn Host, o: &WdObj) -> R<Value> {
        match o {
            WdObj::Range(_) | WdObj::Selection => self.member(host, o, "text", &[], &[]),
            WdObj::Property(_) => self.member(host, o, "value", &[], &[]),
            WdObj::Style(s) => Ok(Value::Str(s.clone())),
            WdObj::Document => self.doc_name(host),
            WdObj::Application => Ok(Value::str("WordCraft")),
            _ => Err(VbaError::new(438)),
        }
    }

    /// The names a macro reaches without an object (`ActiveDocument`, `Selection`).
    pub fn global(&mut self, host: &mut dyn Host, name: &str, args: Args) -> R<Option<Value>> {
        let n = name.to_ascii_lowercase();
        Ok(Some(match n.as_str() {
            "application" => Value::Object(Obj::Wd(WdObj::Application)),
            "activedocument" | "thisdocument" | "documents" | "selection" => self.member(host, &WdObj::Application, name, args, &[])?,
            _ => return Ok(None),
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn constants_and_conversions() {
        assert_eq!(constant("wdAlignParagraphCenter"), Some(Value::Long(1)));
        assert_eq!(constant("wdAlignParagraphJustify"), Some(Value::Long(3)));
        assert_eq!(constant("wdColorRed"), Some(Value::Long(255)));
        assert_eq!(constant("wdColorBlue"), Some(Value::Long(16_711_680)));
        assert_eq!(constant("wdColorAutomatic"), Some(Value::Long(-16_777_216)));
        assert_eq!(constant("wdYellow"), Some(Value::Long(7)));
        assert_eq!(constant("wdBrightGreen"), Some(Value::Long(4)));
        assert_eq!(constant("wdNoHighlight"), Some(Value::Long(0)));
        assert_eq!(constant("wdUnderlineSingle"), Some(Value::Long(1)));
        assert_eq!(constant("wdUnderlineDouble"), Some(Value::Long(3)));
        assert_eq!(constant("wdUndefined"), Some(Value::Long(9_999_999)));
        assert_eq!(constant("wdStyleHeading1"), Some(Value::Long(-2)));
        assert_eq!(constant("wdPropertyAuthor"), Some(Value::Long(3)));
        assert_eq!(constant("wdStory"), Some(Value::Long(6)));
        assert_eq!(constant("xlUp"), None);
        assert_eq!(constant("wdFrobnicate"), None);
        assert_eq!(text_out("a\nb"), "a\rb");
        assert_eq!(text_in("a\r\nb\rc"), "a\nb\nc");
        assert_eq!(with_mark("a\nb"), "a\rb\r");
        assert_eq!(with_mark("a\n"), "a\r");
        assert_eq!(color_arg(&Value::Long(255)).unwrap(), "#FF0000");
        assert_eq!(color_arg(&Value::Long(16_711_680)).unwrap(), "#0000FF");
        assert_eq!(color_arg(&Value::Long(-1)).unwrap_err().number, 5);
        assert_eq!(color_arg(&Value::Long(WD_COLOR_AUTOMATIC)).unwrap_err().number, 5);
        assert_eq!(color_value(&json!("#00FF00")), Value::Long(65_280));
        assert_eq!(color_value(&Json::Null), Value::Long(WD_UNDEFINED));
        assert_eq!(highlight_value(&json!("yellow")), Value::Long(7));
        assert_eq!(highlight_value(&json!("#FFFF00")), Value::Long(7));
        assert_eq!(highlight_value(&json!("Bright Green")), Value::Long(4));
        assert_eq!(highlight_value(&Json::Null), Value::Long(0));
        assert_eq!(highlight_arg(&Value::Long(7)).unwrap(), json!("Yellow"));
        assert_eq!(highlight_arg(&Value::Long(0)).unwrap(), Json::Null);
        assert_eq!(highlight_arg(&Value::Long(99)).unwrap_err().number, 5);
        assert_eq!(underline_number("single"), Some(1));
        assert_eq!(underline_number("None"), Some(0));
        assert_eq!(underline_number("dotDash"), Some(9));
        assert_eq!(underline_number("Mixed"), None);
        assert_eq!(alignment_number(&json!("Centered")), 1);
        assert_eq!(alignment_number(&Json::Null), WD_UNDEFINED);
        assert_eq!(alignment_name(&Value::Long(4)).unwrap_err().number, 5);
    }
}
