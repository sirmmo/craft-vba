//! PowerPoint's object model for macros run in DeckCraft (`Application`, `Presentation`,
//! `Slides`, `Slide`, `Shapes`, `Shape`, `TextFrame`, `TextRange`, `Font`, `ColorFormat`,
//! `FillFormat`, `LineFormat`, `ActiveWindow.Selection`), over the host's `host/sync`: each
//! object is a path from the presentation, as the PowerPoint JavaScript API names them (see
//! DeckCraft's `office.sync`).
//!
//! Slides and shapes are known by their ids (a slide keeps its id when others move), VBA's 1-based
//! indexes become 0-based positions. A shape's properties are read in groups (its position and size
//! together, its font together) and cached until something writes to that shape; slide and shape
//! lists are cached until slides or shapes are added or removed. Writes are queued and sent
//! together before the next read, at most 500 at a time, and when the macro ends.

use std::collections::HashMap;
use std::rc::Rc;

use serde_json::{Map, Value as Json, json};

use crate::excel::{color_from_hex, color_to_hex, host_error};
use crate::host::Host;
use crate::vba::value::{Obj, R, Value, VbaError};

const MAX_PENDING: usize = 500;
/// Most slides or shapes one collection read takes.
const MAX_ITEMS: usize = 100_000;

pub const MSO_TRUE: i32 = -1;
pub const MSO_FALSE: i32 = 0;
pub const MSO_MIXED: i32 = -2;
const MSO_TOGGLE: i32 = -3;

/// A shape: its slide (`None` when it was reached through the selection and its slide isn't
/// known) and its id.
#[derive(Clone, Debug, PartialEq)]
pub struct ShapeRef {
    pub slide: Option<Rc<str>>,
    pub id: Rc<str>,
}

/// Whose colour a `ColorFormat` is.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ColorOf {
    Font,
    Fill,
    Line,
}

#[derive(Clone, Debug, PartialEq)]
pub enum PpObj {
    Application,
    Presentations,
    Presentation,
    Slides,
    Slide(Rc<str>),
    /// A slide's layout (`Slide.CustomLayout`).
    Layout(Rc<str>),
    Shapes(Rc<str>),
    Shape(ShapeRef),
    TextFrame(ShapeRef),
    TextRange(ShapeRef),
    Font(ShapeRef),
    Fill(ShapeRef),
    Line(ShapeRef),
    Color(ShapeRef, ColorOf),
    Window,
    View,
    Selection,
    /// The shapes selected when `Selection.ShapeRange` was read.
    ShapeRange(Rc<[ShapeRef]>),
    /// The slides selected when `Selection.SlideRange` was read.
    SlideRange(Rc<[Rc<str>]>),
    /// The ribbon button a macro was called from (`Sub Go(control As IRibbonControl)`).
    RibbonControl {
        id: Rc<str>,
        tag: Rc<str>,
    },
}

impl PpObj {
    pub fn type_name(&self) -> &'static str {
        match self {
            PpObj::Application => "Application",
            PpObj::Presentations => "Presentations",
            PpObj::Presentation => "Presentation",
            PpObj::Slides => "Slides",
            PpObj::Slide(_) => "Slide",
            PpObj::Layout(_) => "CustomLayout",
            PpObj::Shapes(_) => "Shapes",
            PpObj::Shape(_) => "Shape",
            PpObj::TextFrame(_) => "TextFrame",
            PpObj::TextRange(_) => "TextRange",
            PpObj::Font(_) => "Font",
            PpObj::Fill(_) => "FillFormat",
            PpObj::Line(_) => "LineFormat",
            PpObj::Color(..) => "ColorFormat",
            PpObj::Window => "DocumentWindow",
            PpObj::View => "View",
            PpObj::Selection => "Selection",
            PpObj::ShapeRange(_) => "ShapeRange",
            PpObj::SlideRange(_) => "SlideRange",
            PpObj::RibbonControl { .. } => "IRibbonControl",
        }
    }
}

/// PowerPoint's constants (`ppLayoutText`, `msoShapeRectangle`, `msoTrue`).
pub fn constant(name: &str) -> Option<Value> {
    let l = |x: i32| Some(Value::Long(x));
    let n = name.to_ascii_lowercase();
    if let Some(shape) = n.strip_prefix("msoshape")
        && let Some((v, _)) = AUTO_SHAPES.iter().chain(AUTO_SHAPE_ALIASES).find(|(_, k)| k.eq_ignore_ascii_case(shape))
    {
        return l(*v);
    }
    if let Some(layout) = n.strip_prefix("pplayout")
        && let Some((v, _)) = PP_LAYOUTS.iter().find(|(_, k)| k.eq_ignore_ascii_case(layout))
    {
        return l(*v);
    }
    match n.as_str() {
        "msotrue" => l(MSO_TRUE),
        "msofalse" => l(MSO_FALSE),
        "msoctrue" => l(1),
        "msotristatemixed" => l(MSO_MIXED),
        "msotristatetoggle" => l(MSO_TOGGLE),
        "msotextorientationhorizontal" => l(1),
        "msotextorientationupward" => l(2),
        "msotextorientationdownward" => l(3),
        "msotextorientationverticalfareast" => l(4),
        "msotextorientationvertical" => l(5),
        "msoautoshape" => l(1),
        "msogroup" => l(6),
        "msoline" => l(9),
        "msopicture" => l(13),
        "msoplaceholder" => l(14),
        "msotextbox" => l(17),
        "msotable" => l(19),
        "msoshapetypemixed" => l(-2),
        "msocolortypergb" => l(1),
        "msofillsolid" => l(1),
        "msofillbackground" => l(5),
        "ppselectionnone" => l(0),
        "ppselectionslides" => l(1),
        "ppselectionshapes" => l(2),
        "ppselectiontext" => l(3),
        "ppautosizenone" => l(0),
        "ppautosizeshapetofittext" => l(1),
        "ppautosizemixed" => l(-2),
        "msoautosizenone" => l(0),
        "msoautosizeshapetofittext" => l(1),
        "msoautosizetexttofitshape" => l(2),
        "ppalertsnone" => l(1),
        "ppalertsall" => l(2),
        _ => None,
    }
}

/// `MsoAutoShapeType` values and the PowerPoint JavaScript API's `GeometricShapeType` names.
const AUTO_SHAPES: &[(i32, &str)] = &[
    (1, "rectangle"),
    (2, "parallelogram"),
    (3, "trapezoid"),
    (4, "diamond"),
    (5, "roundRectangle"),
    (6, "octagon"),
    (7, "triangle"),
    (8, "rightTriangle"),
    (9, "ellipse"),
    (10, "hexagon"),
    (11, "plus"),
    (12, "pentagon"),
    (13, "can"),
    (14, "cube"),
    (15, "bevel"),
    (16, "foldedCorner"),
    (17, "smileyFace"),
    (18, "donut"),
    (19, "noSmoking"),
    (20, "blockArc"),
    (21, "heart"),
    (22, "lightningBolt"),
    (23, "sun"),
    (24, "moon"),
    (33, "rightArrow"),
    (34, "leftArrow"),
    (35, "upArrow"),
    (36, "downArrow"),
    (37, "leftRightArrow"),
    (38, "upDownArrow"),
    (51, "homePlate"),
    (52, "chevron"),
    (61, "flowChartProcess"),
    (91, "star4"),
    (92, "star5"),
    (93, "star8"),
    (94, "star16"),
    (95, "star24"),
    (96, "star32"),
];

/// The `msoShape…` constants' own names, where they differ from the JavaScript ones.
const AUTO_SHAPE_ALIASES: &[(i32, &str)] = &[
    (5, "roundedRectangle"),
    (7, "isoscelesTriangle"),
    (9, "oval"),
    (11, "cross"),
    (12, "regularPentagon"),
    (19, "noSymbol"),
    (51, "pentagon"),
    (91, "4pointStar"),
    (92, "5pointStar"),
    (93, "8pointStar"),
    (94, "16pointStar"),
    (95, "24pointStar"),
    (96, "32pointStar"),
];

/// The `ppLayout…` constants (the name after `ppLayout`).
const PP_LAYOUTS: &[(i32, &str)] = &[
    (1, "title"),
    (2, "text"),
    (3, "twoColumnText"),
    (4, "table"),
    (11, "titleOnly"),
    (12, "blank"),
    (16, "object"),
    (32, "custom"),
    (33, "sectionHeader"),
    (34, "comparison"),
    (35, "contentWithCaption"),
    (36, "pictureWithCaption"),
];

/// The layout names a `ppLayout…` value stands for (the standard layouts' names).
fn layout_names(pp: i32) -> &'static [&'static str] {
    match pp {
        1 => &["Title Slide"],
        2 | 16 => &["Title and Content", "Title and Text"],
        3 => &["Two Content", "Two Column Text"],
        11 => &["Title Only"],
        12 => &["Blank"],
        33 => &["Section Header"],
        34 => &["Comparison"],
        35 => &["Content with Caption"],
        36 => &["Picture with Caption"],
        _ => &[],
    }
}

fn layout_number(name: &str) -> i32 {
    [1, 2, 3, 11, 12, 33, 34, 35, 36].into_iter().find(|pp| layout_names(*pp).iter().any(|n| n.eq_ignore_ascii_case(name.trim()))).unwrap_or(32)
}

/// The `GeometricShapeType` an `msoShape…` value draws.
fn geometric(v: i32) -> Option<&'static str> {
    AUTO_SHAPES.iter().find(|(n, _)| *n == v).map(|(_, k)| *k)
}

/// A group of properties read together.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum Part {
    Slide,
    Layout,
    Shape,
    TextFrame,
    TextRange,
    Font,
    Fill,
    Line,
}

impl Part {
    fn props(self) -> &'static [&'static str] {
        match self {
            Part::Slide => &["id", "index"],
            Part::Layout => &["id", "name"],
            Part::Shape => &["id", "name", "type", "left", "top", "width", "height"],
            Part::TextFrame => &["hasText", "wordWrap", "autoSizeSetting"],
            Part::TextRange => &["text", "length", "start"],
            Part::Font => &["bold", "italic", "underline", "size", "name", "color"],
            Part::Fill => &["type", "foregroundColor", "transparency"],
            Part::Line => &["color", "weight", "visible"],
        }
    }
}

fn p(name: &str) -> Json {
    json!({ "p": name })
}

fn m(name: &str, args: Json) -> Json {
    json!({ "m": name, "a": args })
}

fn presentation_path() -> Vec<Json> {
    vec![p("presentation")]
}

fn slides_path() -> Vec<Json> {
    vec![p("presentation"), p("slides")]
}

fn slide_path(id: &str) -> Vec<Json> {
    let mut v = slides_path();
    v.push(m("getItem", json!([id])));
    v
}

fn shapes_path(slide: &str) -> Vec<Json> {
    let mut v = slide_path(slide);
    v.push(p("shapes"));
    v
}

fn shape_path(r: &ShapeRef) -> Vec<Json> {
    let mut v = match &r.slide {
        Some(s) => shapes_path(s),
        None => vec![p("presentation"), m("getSelectedShapes", json!([]))],
    };
    v.push(m("getItem", json!([r.id.as_ref()])));
    v
}

fn part_path(r: &ShapeRef, part: Part) -> Vec<Json> {
    let mut v = shape_path(r);
    match part {
        Part::TextFrame => v.push(p("textFrame")),
        Part::TextRange => v.extend([p("textFrame"), p("textRange")]),
        Part::Font => v.extend([p("textFrame"), p("textRange"), p("font")]),
        Part::Fill => v.push(p("fill")),
        Part::Line => v.push(p("lineFormat")),
        Part::Shape | Part::Slide | Part::Layout => {}
    }
    v
}

/// An id as the bridge sends it (a string, or a number).
fn id_str(j: &Json) -> Option<Rc<str>> {
    match j {
        Json::String(s) => Some(s.as_str().into()),
        Json::Number(n) => Some(n.to_string().into()),
        _ => None,
    }
}

fn items(out: &Map<String, Json>) -> Vec<&Map<String, Json>> {
    out.get("items").and_then(Json::as_array).into_iter().flatten().filter_map(Json::as_object).take(MAX_ITEMS).collect()
}

fn ids_of(out: &Map<String, Json>) -> Vec<Rc<str>> {
    items(out).into_iter().filter_map(|i| i.get("id").and_then(id_str)).collect()
}

/// An id as VBA's `Long` (`SlideID`, `Shape.Id`), else as text.
fn id_value(id: &str) -> Value {
    id.parse::<i32>().map(Value::Long).unwrap_or_else(|_| Value::str(id))
}

fn tri(j: &Json) -> Value {
    Value::Long(match j.as_bool() {
        Some(true) => MSO_TRUE,
        Some(false) => MSO_FALSE,
        None => MSO_MIXED,
    })
}

fn number(j: &Json) -> Value {
    j.as_f64().map(Value::Double).unwrap_or(Value::Null)
}

fn finite(v: &Value) -> R<f64> {
    let f = v.to_f64()?;
    if f.is_finite() { Ok(f) } else { Err(VbaError::new(6)) }
}

fn color_arg(v: &Value) -> R<String> {
    let n = v.to_i32()?;
    if !(0..=0xFF_FFFF).contains(&n) {
        return Err(VbaError::with(5, "Colours run from 0 to 16777215 (RGB(255, 255, 255))"));
    }
    Ok(color_to_hex(n))
}

/// Text as PowerPoint's macros see it: paragraphs end with a carriage return (`vbCr`).
fn text_out(s: &str) -> String {
    s.replace("\r\n", "\n").replace('\n', "\r")
}

fn text_in(s: &str) -> String {
    s.replace("\r\n", "\n").replace('\r', "\n")
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

fn obj(x: PpObj) -> R<Value> {
    Ok(Value::Object(Obj::Pp(x)))
}

type Args<'a> = &'a [Value];
type Key = (Option<Rc<str>>, Option<Rc<str>>, Part);

/// The macro's view of the presentation.
#[derive(Default)]
pub struct PowerPoint {
    pending: Vec<Json>,
    /// Requests sent (for tests and tuning).
    pub requests: usize,
    slides: Option<Vec<Rc<str>>>,
    shapes: HashMap<Rc<str>, Vec<(Rc<str>, String)>>,
    props: HashMap<Key, Map<String, Json>>,
    /// The layouts' (id, name), once looked up; empty when the app doesn't list them.
    layouts: Option<Vec<(Rc<str>, String)>>,
}

impl PowerPoint {
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
        let mut r = self.sync(host, vec![json!({"op": "load", "path": path, "props": props})])?;
        Ok(match r.pop() {
            Some(Json::Object(o)) => o,
            _ => Map::new(),
        })
    }

    /// Queues a write to shape `id`; what was read of it no longer holds.
    fn write(&mut self, host: &mut dyn Host, action: Json, id: &str) -> R<()> {
        self.pending.push(action);
        self.props.retain(|(_, sh, _), _| sh.as_deref() != Some(id));
        if self.pending.len() >= MAX_PENDING {
            self.flush(host)?;
        }
        Ok(())
    }

    fn set(&mut self, host: &mut dyn Host, r: &ShapeRef, part: Part, prop: &str, value: Json) -> R<()> {
        let action = json!({"op": "set", "path": part_path(r, part), "prop": prop, "value": value});
        self.write(host, action, &r.id)
    }

    fn call(&mut self, host: &mut dyn Host, r: &ShapeRef, part: Part, method: &str, args: Json) -> R<()> {
        let action = json!({"op": "call", "path": part_path(r, part), "method": method, "args": args});
        self.write(host, action, &r.id)
    }

    /// A call that changes the slides, shapes or selection, sent now: nothing cached holds.
    fn now(&mut self, host: &mut dyn Host, path: Vec<Json>, method: &str, args: Json) -> R<Json> {
        let r = self.sync(host, vec![json!({"op": "call", "path": path, "method": method, "args": args})]);
        self.forget();
        Ok(r?.pop().unwrap_or(Json::Null))
    }

    fn forget(&mut self) {
        self.slides = None;
        self.shapes.clear();
        self.props.clear();
    }

    /// One property of a shape's `part` (the whole group is read once).
    fn shape_prop(&mut self, host: &mut dyn Host, r: &ShapeRef, part: Part, prop: &str) -> R<Json> {
        let key = (r.slide.clone(), Some(r.id.clone()), part);
        if let Some(v) = self.props.get(&key) {
            return Ok(v.get(prop).cloned().unwrap_or(Json::Null));
        }
        let out = self.load(host, part_path(r, part), part.props())?;
        let v = out.get(prop).cloned().unwrap_or(Json::Null);
        self.props.insert(key, out);
        Ok(v)
    }

    fn slide_prop(&mut self, host: &mut dyn Host, slide: &Rc<str>, part: Part, prop: &str) -> R<Json> {
        let key = (Some(slide.clone()), None, part);
        if let Some(v) = self.props.get(&key) {
            return Ok(v.get(prop).cloned().unwrap_or(Json::Null));
        }
        let mut path = slide_path(slide);
        if part == Part::Layout {
            path.push(p("layout"));
        }
        let out = self.load(host, path, part.props())?;
        let v = out.get(prop).cloned().unwrap_or(Json::Null);
        self.props.insert(key, out);
        Ok(v)
    }

    fn slide_ids(&mut self, host: &mut dyn Host) -> R<Vec<Rc<str>>> {
        if let Some(s) = &self.slides {
            return Ok(s.clone());
        }
        let out = self.load(host, slides_path(), &["items"])?;
        let ids = ids_of(&out);
        self.slides = Some(ids.clone());
        Ok(ids)
    }

    fn shape_list(&mut self, host: &mut dyn Host, slide: &Rc<str>) -> R<Vec<(Rc<str>, String)>> {
        if let Some(s) = self.shapes.get(slide) {
            return Ok(s.clone());
        }
        let out = self.load(host, shapes_path(slide), &["items", "name"])?;
        let list: Vec<(Rc<str>, String)> = items(&out)
            .into_iter()
            .filter_map(|i| Some((i.get("id").and_then(id_str)?, i.get("name").and_then(Json::as_str).unwrap_or("").to_string())))
            .collect();
        self.shapes.insert(slide.clone(), list.clone());
        Ok(list)
    }

    /// The selected slides and shapes (ids), in one request.
    fn selection(&mut self, host: &mut dyn Host) -> R<(Vec<Rc<str>>, Vec<Rc<str>>)> {
        let mut slides = presentation_path();
        slides.push(m("getSelectedSlides", json!([])));
        let mut shapes = presentation_path();
        shapes.push(m("getSelectedShapes", json!([])));
        let out =
            self.sync(host, vec![json!({"op": "load", "path": slides, "props": ["items"]}), json!({"op": "load", "path": shapes, "props": ["items"]})])?;
        let ids = |j: Option<&Json>| j.and_then(Json::as_object).map(ids_of).unwrap_or_default();
        Ok((ids(out.first()), ids(out.get(1))))
    }

    /// The selected shapes, with their slide when it's the one selected slide.
    fn selected_shapes(&mut self, host: &mut dyn Host) -> R<Vec<ShapeRef>> {
        let (slides, shapes) = self.selection(host)?;
        let slide = match slides.as_slice() {
            [one] if !shapes.is_empty() => Some(one.clone()),
            _ => None,
        };
        let on_slide = match &slide {
            Some(s) => self.shape_list(host, s)?,
            None => Vec::new(),
        };
        Ok(shapes
            .into_iter()
            .map(|id| {
                let known = on_slide.iter().any(|(x, _)| *x == id);
                ShapeRef { slide: if known { slide.clone() } else { None }, id }
            })
            .collect())
    }

    /// The layouts the presentation's first master offers (PowerPoint API 1.3's
    /// `slideMasters`/`layouts`), or none when the app doesn't list them.
    fn layouts(&mut self, host: &mut dyn Host) -> R<Vec<(Rc<str>, String)>> {
        if let Some(l) = &self.layouts {
            return Ok(l.clone());
        }
        // Alone, so that a refusal costs no queued write.
        self.flush(host)?;
        let path = vec![p("presentation"), p("slideMasters"), m("getItemAt", json!([0])), p("layouts")];
        let found = match self.load(host, path, &["items", "name"]) {
            Ok(out) => items(&out)
                .into_iter()
                .filter_map(|i| Some((i.get("id").and_then(id_str)?, i.get("name").and_then(Json::as_str).unwrap_or("").to_string())))
                .collect(),
            Err(_) => Vec::new(),
        };
        self.layouts = Some(found.clone());
        Ok(found)
    }

    fn shape_type(&mut self, host: &mut dyn Host, r: &ShapeRef) -> R<String> {
        Ok(self.shape_prop(host, r, Part::Shape, "type")?.as_str().unwrap_or("unsupported").to_string())
    }

    fn has_text_frame(&mut self, host: &mut dyn Host, r: &ShapeRef) -> R<bool> {
        Ok(matches!(self.shape_type(host, r)?.as_str(), "geometricShape" | "textBox"))
    }

    fn shape_named(&mut self, host: &mut dyn Host, slide: &Rc<str>, key: &Value) -> R<ShapeRef> {
        let list = self.shape_list(host, slide)?;
        let id = match key {
            Value::Str(name) => list
                .iter()
                .find(|(_, n)| n == name.as_ref())
                .or_else(|| list.iter().find(|(_, n)| n.eq_ignore_ascii_case(name)))
                .map(|(id, _)| id.clone())
                .ok_or_else(|| VbaError::with(9, format!("Shapes.Item: no shape named '{name}' on this slide")))?,
            v => list.get(position(v, list.len())?).map(|(id, _)| id.clone()).ok_or(VbaError::new(9))?,
        };
        Ok(ShapeRef { slide: Some(slide.clone()), id })
    }

    fn slide_at(&mut self, host: &mut dyn Host, key: &Value) -> R<Rc<str>> {
        let ids = self.slide_ids(host)?;
        if let Value::Str(s) = key
            && s.trim().parse::<f64>().is_err()
        {
            return Err(VbaError::with(9, format!("Slides.Item: no slide named '{s}'")));
        }
        ids.get(position(key, ids.len())?).cloned().ok_or(VbaError::new(9))
    }

    fn add_shape(&mut self, host: &mut dyn Host, slide: &Rc<str>, method: &str, first: Json, (l, t, w, h): (f64, f64, f64, f64)) -> R<Value> {
        let opts = json!({"left": l, "top": t, "width": w, "height": h});
        let out = self.now(host, shapes_path(slide), method, json!([first, opts]))?;
        let id = out.get("id").and_then(id_str).ok_or_else(|| VbaError::with(1004, format!("Shapes.{method}: the app didn't say which shape it added")))?;
        obj(PpObj::Shape(ShapeRef { slide: Some(slide.clone()), id }))
    }

    /// `object.name` or `object.name(args)`: a property's value or a method's result.
    pub fn member(&mut self, host: &mut dyn Host, o: &PpObj, name: &str, args: Args, named: &[(String, Value)]) -> R<Value> {
        let n = name.to_ascii_lowercase();
        match (o, n.as_str()) {
            // Application.
            (PpObj::Application, "activepresentation") => obj(PpObj::Presentation),
            (PpObj::Application, "activewindow") => obj(PpObj::Window),
            (PpObj::Application, "presentations") => {
                if args.is_empty() {
                    obj(PpObj::Presentations)
                } else {
                    self.member(host, &PpObj::Presentations, "item", args, named)
                }
            }
            (PpObj::Application, "name") => Ok(Value::str("DeckCraft")),
            (PpObj::Application, "version") => Ok(Value::str("16.0")),
            (PpObj::Application, "displayalerts") => Ok(Value::Long(2)),
            (_, "application") => obj(PpObj::Application),
            (PpObj::RibbonControl { id, .. }, "id") => Ok(Value::Str(id.clone())),
            (PpObj::RibbonControl { tag, .. }, "tag") => Ok(Value::Str(tag.clone())),
            (PpObj::RibbonControl { .. }, "context") => obj(PpObj::Window),

            (PpObj::Presentations, "count") => Ok(Value::Long(1)),
            (PpObj::Presentations, "item") => {
                position(need(args, named, 0, "index")?, 1)?;
                obj(PpObj::Presentation)
            }
            (PpObj::Presentations | PpObj::Presentation, "parent") => obj(PpObj::Application),

            // The presentation.
            (PpObj::Presentation, "slides") => {
                if args.is_empty() {
                    obj(PpObj::Slides)
                } else {
                    self.member(host, &PpObj::Slides, "item", args, named)
                }
            }
            (PpObj::Presentation, "name" | "fullname") => {
                let out = self.load(host, presentation_path(), &["title"])?;
                Ok(Value::str(out.get("title").and_then(Json::as_str).unwrap_or("")))
            }
            (PpObj::Presentation, "windows") => obj(PpObj::Window),

            // Slides.
            (PpObj::Slides, "count") => Ok(Value::Long(self.slide_ids(host)?.len() as i32)),
            (PpObj::Slides, "item") => {
                let id = self.slide_at(host, need(args, named, 0, "index")?)?;
                obj(PpObj::Slide(id))
            }
            (PpObj::Slides, "findbyslideid") => {
                let want = need(args, named, 0, "slideid")?.to_str()?;
                let ids = self.slide_ids(host)?;
                let id = ids.into_iter().find(|i| **i == *want).ok_or_else(|| VbaError::with(9, format!("Slides.FindBySlideID: no slide {want}")))?;
                obj(PpObj::Slide(id))
            }
            (PpObj::Slides, "add") => {
                let index = need(args, named, 0, "index")?.to_i64()?;
                let layout = need(args, named, 1, "layout")?.to_i32()?;
                let count = self.slide_ids(host)?.len();
                if index < 1 || index > count as i64 + 1 {
                    return Err(VbaError::with(9, format!("Slides.Add: Index out of range (1 to {})", count + 1)));
                }
                let mut opts = Map::new();
                opts.insert("index".into(), json!(index - 1));
                let names = layout_names(layout);
                if !names.is_empty() {
                    let layouts = self.layouts(host)?;
                    if let Some((id, _)) = names.iter().find_map(|n| layouts.iter().find(|(_, x)| x.eq_ignore_ascii_case(n))) {
                        opts.insert("layoutId".into(), json!(id.as_ref()));
                    }
                }
                let out = self.now(host, slides_path(), "add", json!([opts]))?;
                let id = out.get("id").and_then(id_str).ok_or_else(|| VbaError::with(1004, "Slides.Add: the app didn't say which slide it added"))?;
                obj(PpObj::Slide(id))
            }
            (PpObj::Slides, "parent") => obj(PpObj::Presentation),

            // A slide.
            (PpObj::Slide(s), "slideindex" | "slidenumber") => {
                let i = self.slide_prop(host, s, Part::Slide, "index")?.as_i64().ok_or(VbaError::new(1004))?;
                Ok(Value::Long(i32::try_from(i + 1).map_err(|_| VbaError::new(6))?))
            }
            (PpObj::Slide(s), "slideid") => Ok(id_value(s)),
            (PpObj::Slide(s), "shapes") => {
                if args.is_empty() {
                    obj(PpObj::Shapes(s.clone()))
                } else {
                    self.member(host, &PpObj::Shapes(s.clone()), "item", args, named)
                }
            }
            (PpObj::Slide(s), "delete") => {
                self.now(host, slide_path(s), "delete", json!([]))?;
                Ok(Value::Empty)
            }
            (PpObj::Slide(s), "select") => {
                self.now(host, presentation_path(), "setSelectedSlides", json!([[s.as_ref()]]))?;
                Ok(Value::Empty)
            }
            (PpObj::Slide(s), "layout") => {
                let name = self.slide_prop(host, s, Part::Layout, "name")?;
                Ok(Value::Long(layout_number(name.as_str().unwrap_or(""))))
            }
            (PpObj::Slide(s), "customlayout") => obj(PpObj::Layout(s.clone())),
            (PpObj::Slide(_), "parent") => obj(PpObj::Presentation),
            (PpObj::Layout(s), "name") => Ok(Value::str(self.slide_prop(host, s, Part::Layout, "name")?.as_str().unwrap_or(""))),
            (PpObj::Layout(s), "parent") => obj(PpObj::Slide(s.clone())),

            // Shapes.
            (PpObj::Shapes(s), "count") => Ok(Value::Long(self.shape_list(host, s)?.len() as i32)),
            (PpObj::Shapes(s), "item") => {
                let r = self.shape_named(host, s, need(args, named, 0, "index")?)?;
                obj(PpObj::Shape(r))
            }
            (PpObj::Shapes(s), "addtextbox") => {
                let orientation = need(args, named, 0, "orientation")?.to_i32()?;
                if orientation != 1 {
                    return Err(VbaError::with(5, "AddTextbox: only msoTextOrientationHorizontal is supported yet"));
                }
                let box_ = (
                    finite(need(args, named, 1, "left")?)?,
                    finite(need(args, named, 2, "top")?)?,
                    finite(need(args, named, 3, "width")?)?,
                    finite(need(args, named, 4, "height")?)?,
                );
                self.add_shape(host, s, "addTextBox", json!(""), box_)
            }
            (PpObj::Shapes(s), "addshape") => {
                let t = need(args, named, 0, "type")?.to_i32()?;
                let kind = geometric(t).ok_or_else(|| VbaError::with(5, format!("AddShape: AutoShape type {t} isn't supported yet")))?;
                let box_ = (
                    finite(need(args, named, 1, "left")?)?,
                    finite(need(args, named, 2, "top")?)?,
                    finite(need(args, named, 3, "width")?)?,
                    finite(need(args, named, 4, "height")?)?,
                );
                self.add_shape(host, s, "addGeometricShape", json!(kind), box_)
            }
            (PpObj::Shapes(s), "addline") => {
                let (x1, y1) = (finite(need(args, named, 0, "beginx")?)?, finite(need(args, named, 1, "beginy")?)?);
                let (x2, y2) = (finite(need(args, named, 2, "endx")?)?, finite(need(args, named, 3, "endy")?)?);
                self.add_shape(host, s, "addLine", json!("straight"), (x1.min(x2), y1.min(y2), (x2 - x1).abs(), (y2 - y1).abs()))
            }
            (PpObj::Shapes(s), "findbyid") => {
                let want = need(args, named, 0, "id")?.to_str()?;
                let id = self
                    .shape_list(host, s)?
                    .into_iter()
                    .map(|(id, _)| id)
                    .find(|i| **i == *want.trim())
                    .ok_or_else(|| VbaError::with(9, format!("Shapes.FindById: no shape {want} on this slide")))?;
                obj(PpObj::Shape(ShapeRef { slide: Some(s.clone()), id }))
            }
            (PpObj::Shapes(s), "parent") => obj(PpObj::Slide(s.clone())),

            // A shape.
            (PpObj::Shape(r), "name") => Ok(Value::str(self.shape_prop(host, r, Part::Shape, "name")?.as_str().unwrap_or(""))),
            (PpObj::Shape(r), p @ ("left" | "top" | "width" | "height")) => Ok(number(&self.shape_prop(host, r, Part::Shape, p)?)),
            (PpObj::Shape(r), "id") => Ok(id_value(&r.id)),
            (PpObj::Shape(r), "type") => Ok(Value::Long(match self.shape_type(host, r)?.as_str() {
                "geometricShape" => 1,
                "group" => 6,
                "line" => 9,
                "image" => 13,
                "textBox" => 17,
                "table" => 19,
                _ => -2,
            })),
            (PpObj::Shape(r), "hastextframe") => Ok(Value::Long(if self.has_text_frame(host, r)? { MSO_TRUE } else { MSO_FALSE })),
            (PpObj::Shape(r), "textframe" | "textframe2") => {
                if !self.has_text_frame(host, r)? {
                    return Err(VbaError::with(5, "This type of shape has no text frame (check HasTextFrame)"));
                }
                obj(PpObj::TextFrame(r.clone()))
            }
            (PpObj::Shape(r), "fill") => obj(PpObj::Fill(r.clone())),
            (PpObj::Shape(r), "line") => obj(PpObj::Line(r.clone())),
            (PpObj::Shape(r), "delete") => {
                self.now(host, shape_path(r), "delete", json!([]))?;
                Ok(Value::Empty)
            }
            (PpObj::Shape(r), "select") => {
                let Some(s) = &r.slide else { return Ok(Value::Empty) };
                let replace = arg(args, named, 0, "replace").map(Value::to_bool).transpose()?.unwrap_or(true);
                let mut ids: Vec<Rc<str>> = if replace { Vec::new() } else { self.selection(host)?.1 };
                if !ids.contains(&r.id) {
                    ids.push(r.id.clone());
                }
                let ids: Vec<&str> = ids.iter().map(|i| i.as_ref()).collect();
                self.now(host, slide_path(s), "setSelectedShapes", json!([ids]))?;
                Ok(Value::Empty)
            }
            (PpObj::Shape(r), p @ ("incrementleft" | "incrementtop")) => {
                let by = finite(need(args, named, 0, "increment")?)?;
                let prop = if p == "incrementleft" { "left" } else { "top" };
                let v = self.shape_prop(host, r, Part::Shape, prop)?.as_f64().unwrap_or(0.0) + by;
                if !v.is_finite() {
                    return Err(VbaError::new(6));
                }
                self.set(host, r, Part::Shape, prop, json!(v))?;
                Ok(Value::Empty)
            }
            (PpObj::Shape(r), "parent") => match &r.slide {
                Some(s) => obj(PpObj::Slide(s.clone())),
                None => Err(VbaError::with(438, "This shape's slide isn't known (it was reached through the selection)")),
            },

            // Text.
            (PpObj::TextFrame(r), "textrange") => obj(PpObj::TextRange(r.clone())),
            (PpObj::TextFrame(r), "hastext") => Ok(tri(&self.shape_prop(host, r, Part::TextFrame, "hasText")?)),
            (PpObj::TextFrame(r), "wordwrap") => Ok(tri(&self.shape_prop(host, r, Part::TextFrame, "wordWrap")?)),
            (PpObj::TextFrame(r), "autosize") => Ok(Value::Long(match self.shape_prop(host, r, Part::TextFrame, "autoSizeSetting")?.as_str() {
                Some("autoSizeShapeToFitText") => 1,
                Some("autoSizeTextToFitShape") => 2,
                Some(_) => 0,
                None => -2,
            })),
            (PpObj::TextFrame(r), "parent") => obj(PpObj::Shape(r.clone())),
            (PpObj::TextRange(r), "text") => Ok(Value::str(text_out(self.shape_prop(host, r, Part::TextRange, "text")?.as_str().unwrap_or("")))),
            (PpObj::TextRange(r), "length") => {
                Ok(Value::Long(self.shape_prop(host, r, Part::TextRange, "length")?.as_i64().unwrap_or(0).clamp(0, i64::from(i32::MAX)) as i32))
            }
            (PpObj::TextRange(r), "start") => {
                Ok(Value::Long(self.shape_prop(host, r, Part::TextRange, "start")?.as_i64().unwrap_or(0).clamp(0, i64::from(i32::MAX) - 1) as i32 + 1))
            }
            (PpObj::TextRange(r), "font") => obj(PpObj::Font(r.clone())),
            (PpObj::TextRange(r), "insertafter" | "insertbefore") => {
                let more = text_in(&arg(args, named, 0, "newtext").map(Value::to_str).transpose()?.unwrap_or_default());
                let now = self.shape_prop(host, r, Part::TextRange, "text")?.as_str().unwrap_or("").to_string();
                let text = if n == "insertafter" { format!("{now}{more}") } else { format!("{more}{now}") };
                self.set(host, r, Part::TextRange, "text", json!(text))?;
                obj(PpObj::TextRange(r.clone()))
            }
            (PpObj::TextRange(r), "parent") => obj(PpObj::TextFrame(r.clone())),

            // Font.
            (PpObj::Font(r), p @ ("bold" | "italic")) => Ok(tri(&self.shape_prop(host, r, Part::Font, p)?)),
            (PpObj::Font(r), "underline") => Ok(Value::Long(match self.shape_prop(host, r, Part::Font, "underline")?.as_str() {
                Some("none") => MSO_FALSE,
                Some(_) => MSO_TRUE,
                None => MSO_MIXED,
            })),
            (PpObj::Font(r), "size") => Ok(number(&self.shape_prop(host, r, Part::Font, "size")?)),
            (PpObj::Font(r), "name") => Ok(Value::str(self.shape_prop(host, r, Part::Font, "name")?.as_str().unwrap_or(""))),
            (PpObj::Font(r), "color") => obj(PpObj::Color(r.clone(), ColorOf::Font)),

            // Colours.
            (PpObj::Color(r, of), "rgb") => {
                let (part, prop) = match of {
                    ColorOf::Font => (Part::Font, "color"),
                    ColorOf::Fill => (Part::Fill, "foregroundColor"),
                    ColorOf::Line => (Part::Line, "color"),
                };
                let v = self.shape_prop(host, r, part, prop)?;
                Ok(match v.as_str().and_then(color_from_hex) {
                    Some(c) => Value::Long(c),
                    None if *of == ColorOf::Font => Value::Null,
                    None => Value::Long(0),
                })
            }
            (PpObj::Color(..), "type") => Ok(Value::Long(1)),

            // Fill and line.
            (PpObj::Fill(r), "forecolor") => obj(PpObj::Color(r.clone(), ColorOf::Fill)),
            (PpObj::Fill(r), "visible") => {
                let t = self.shape_prop(host, r, Part::Fill, "type")?;
                Ok(Value::Long(if t.as_str() == Some("noFill") { MSO_FALSE } else { MSO_TRUE }))
            }
            (PpObj::Fill(r), "transparency") => Ok(number(&self.shape_prop(host, r, Part::Fill, "transparency")?)),
            (PpObj::Fill(r), "solid") => {
                let c = self.shape_prop(host, r, Part::Fill, "foregroundColor")?;
                let c = c.as_str().filter(|c| color_from_hex(c).is_some()).unwrap_or("#FFFFFF").to_string();
                self.call(host, r, Part::Fill, "setSolidColor", json!([c]))?;
                Ok(Value::Empty)
            }
            (PpObj::Line(r), "forecolor") => obj(PpObj::Color(r.clone(), ColorOf::Line)),
            (PpObj::Line(r), "weight") => Ok(number(&self.shape_prop(host, r, Part::Line, "weight")?)),
            (PpObj::Line(r), "visible") => Ok(tri(&self.shape_prop(host, r, Part::Line, "visible")?)),
            (PpObj::Fill(r) | PpObj::Line(r), "parent") => obj(PpObj::Shape(r.clone())),

            // The window and its selection.
            (PpObj::Window, "selection") => obj(PpObj::Selection),
            (PpObj::Window, "view") => obj(PpObj::View),
            (PpObj::Window, "presentation") => obj(PpObj::Presentation),
            (PpObj::Window | PpObj::Selection | PpObj::View, "parent") => obj(PpObj::Window),
            (PpObj::View, "gotoslide") => {
                let id = self.slide_at(host, need(args, named, 0, "index")?)?;
                self.now(host, presentation_path(), "setSelectedSlides", json!([[id.as_ref()]]))?;
                Ok(Value::Empty)
            }
            (PpObj::View, "slide") => {
                let (slides, _) = self.selection(host)?;
                let id = slides.into_iter().next().ok_or_else(|| VbaError::with(91, "No slide is showing"))?;
                obj(PpObj::Slide(id))
            }
            (PpObj::Selection, "shaperange") => {
                let shapes = self.selected_shapes(host)?;
                let range = PpObj::ShapeRange(shapes.into());
                if args.is_empty() { obj(range) } else { self.member(host, &range, "item", args, named) }
            }
            (PpObj::Selection, "sliderange") => {
                let (slides, _) = self.selection(host)?;
                let range = PpObj::SlideRange(slides.into());
                if args.is_empty() { obj(range) } else { self.member(host, &range, "item", args, named) }
            }
            (PpObj::Selection, "textrange") => match self.selected_shapes(host)?.as_slice() {
                [one] => {
                    if !self.has_text_frame(host, one)? {
                        return Err(VbaError::with(5, "The selected shape has no text frame"));
                    }
                    obj(PpObj::TextRange(one.clone()))
                }
                [] => Err(VbaError::with(91, "Nothing appropriate is currently selected (no shapes are selected)")),
                _ => Err(VbaError::with(438, "Selection.TextRange with several shapes selected isn't supported yet (use ShapeRange(i))")),
            },
            (PpObj::Selection, "unselect") => {
                let (slides, shapes) = self.selection(host)?;
                if let (Some(s), false) = (slides.first(), shapes.is_empty()) {
                    self.now(host, slide_path(s), "setSelectedShapes", json!([Vec::<&str>::new()]))?;
                }
                Ok(Value::Empty)
            }
            (PpObj::Selection, "type") => {
                let (slides, shapes) = self.selection(host)?;
                Ok(Value::Long(if !shapes.is_empty() {
                    2
                } else if !slides.is_empty() {
                    1
                } else {
                    0
                }))
            }

            // Ranges: a collection, and the one thing in them when there's one.
            (PpObj::ShapeRange(shapes), "count") => Ok(Value::Long(shapes.len() as i32)),
            (PpObj::ShapeRange(shapes), "item") => {
                let key = need(args, named, 0, "index")?;
                let r = match key {
                    Value::Str(name) => {
                        let mut found = None;
                        for r in shapes.iter() {
                            if self.shape_prop(host, r, Part::Shape, "name")?.as_str().is_some_and(|n| n.eq_ignore_ascii_case(name)) {
                                found = Some(r.clone());
                                break;
                            }
                        }
                        found.ok_or_else(|| VbaError::with(9, format!("ShapeRange.Item: no shape named '{name}' is selected")))?
                    }
                    v => shapes.get(position(v, shapes.len())?).cloned().ok_or(VbaError::new(9))?,
                };
                obj(PpObj::Shape(r))
            }
            (PpObj::ShapeRange(shapes), "delete") => {
                for r in shapes.iter() {
                    self.now(host, shape_path(r), "delete", json!([]))?;
                }
                Ok(Value::Empty)
            }
            (PpObj::ShapeRange(shapes), "incrementleft" | "incrementtop") if shapes.len() > 1 => {
                for r in shapes.iter() {
                    self.member(host, &PpObj::Shape(r.clone()), name, args, named)?;
                }
                Ok(Value::Empty)
            }
            (PpObj::ShapeRange(shapes), _) => match shapes.as_ref() {
                [one] => self.member(host, &PpObj::Shape(one.clone()), name, args, named),
                [] => Err(VbaError::with(91, "Nothing appropriate is currently selected (no shapes are selected)")),
                _ => Err(VbaError::with(438, format!("ShapeRange.{name} with several shapes selected isn't supported yet (use ShapeRange(i))"))),
            },
            (PpObj::SlideRange(slides), "count") => Ok(Value::Long(slides.len() as i32)),
            (PpObj::SlideRange(slides), "item") => {
                let i = position(need(args, named, 0, "index")?, slides.len())?;
                obj(PpObj::Slide(slides.get(i).cloned().ok_or(VbaError::new(9))?))
            }
            (PpObj::SlideRange(slides), "delete") if slides.len() > 1 => {
                for s in slides.iter() {
                    self.now(host, slide_path(s), "delete", json!([]))?;
                }
                Ok(Value::Empty)
            }
            (PpObj::SlideRange(slides), _) => match slides.as_ref() {
                [one] => self.member(host, &PpObj::Slide(one.clone()), name, args, named),
                [] => Err(VbaError::with(91, "Nothing appropriate is currently selected (no slides are selected)")),
                _ => Err(VbaError::with(438, format!("SlideRange.{name} with several slides selected isn't supported yet (use SlideRange(i))"))),
            },
            (o, _) => Err(VbaError::with(438, format!("{}.{name} isn't supported yet", o.type_name()))),
        }
    }

    /// A tri-state written: `msoTrue`/`True`, `msoFalse`/`False`, or `msoTriStateToggle`.
    fn tri_arg(&mut self, host: &mut dyn Host, value: &Value, current: impl FnOnce(&mut Self, &mut dyn Host) -> R<Json>) -> R<bool> {
        match value.to_i32() {
            Ok(MSO_TOGGLE) => Ok(current(self, host)?.as_bool() != Some(true)),
            Ok(MSO_MIXED) => Err(VbaError::new(5)),
            _ => value.to_bool(),
        }
    }

    /// `object.name = value` (PowerPoint's settable properties take no arguments).
    pub fn assign(&mut self, host: &mut dyn Host, o: &PpObj, name: &str, value: &Value) -> R<()> {
        let n = name.to_ascii_lowercase();
        match (o, n.as_str()) {
            (PpObj::Application, "displayalerts") => Ok(()),
            (PpObj::Shape(r), "name") => {
                let new = value.to_str()?;
                self.set(host, r, Part::Shape, "name", json!(new))?;
                self.shapes.clear();
                Ok(())
            }
            (PpObj::Shape(r), p @ ("left" | "top" | "width" | "height")) => {
                let v = finite(value)?;
                if matches!(p, "width" | "height") && v < 0.0 {
                    return Err(VbaError::with(5, format!("Shape.{name} can't be negative")));
                }
                self.set(host, r, Part::Shape, p, json!(v))
            }
            (PpObj::TextFrame(r), "wordwrap") => {
                let on = self.tri_arg(host, value, |me, h| me.shape_prop(h, r, Part::TextFrame, "wordWrap"))?;
                self.set(host, r, Part::TextFrame, "wordWrap", json!(on))
            }
            (PpObj::TextFrame(r), "autosize") => {
                let s = match value.to_i32()? {
                    0 => "autoSizeNone",
                    1 => "autoSizeShapeToFitText",
                    2 => "autoSizeTextToFitShape",
                    _ => return Err(VbaError::new(5)),
                };
                self.set(host, r, Part::TextFrame, "autoSizeSetting", json!(s))
            }
            (PpObj::TextRange(r), "text") => self.set(host, r, Part::TextRange, "text", json!(text_in(&value.to_str()?))),
            (PpObj::Font(r), p @ ("bold" | "italic")) => {
                let on = self.tri_arg(host, value, |me, h| me.shape_prop(h, r, Part::Font, p))?;
                self.set(host, r, Part::Font, p, json!(on))
            }
            (PpObj::Font(r), "underline") => {
                let on = self.tri_arg(host, value, |me, h| Ok(json!(me.shape_prop(h, r, Part::Font, "underline")?.as_str().is_some_and(|u| u != "none"))))?;
                self.set(host, r, Part::Font, "underline", json!(if on { "single" } else { "none" }))
            }
            (PpObj::Font(r), "size") => {
                let v = finite(value)?;
                if !(1.0..=4000.0).contains(&v) {
                    return Err(VbaError::with(5, "Font sizes run from 1 to 4000 points"));
                }
                self.set(host, r, Part::Font, "size", json!(v))
            }
            (PpObj::Font(r), "name") => self.set(host, r, Part::Font, "name", json!(value.to_str()?)),
            (PpObj::Font(r), "color") => self.assign(host, &PpObj::Color(r.clone(), ColorOf::Font), "rgb", value),
            (PpObj::Fill(r), "forecolor") => self.assign(host, &PpObj::Color(r.clone(), ColorOf::Fill), "rgb", value),
            (PpObj::Line(r), "forecolor") => self.assign(host, &PpObj::Color(r.clone(), ColorOf::Line), "rgb", value),
            (PpObj::Color(r, of), "rgb") => {
                let hex = color_arg(value)?;
                match of {
                    ColorOf::Font => self.set(host, r, Part::Font, "color", json!(hex)),
                    // A fill given a colour is a solid fill of it, as in PowerPoint.
                    ColorOf::Fill => self.call(host, r, Part::Fill, "setSolidColor", json!([hex])),
                    ColorOf::Line => self.set(host, r, Part::Line, "color", json!(hex)),
                }
            }
            (PpObj::Fill(r), "visible") => {
                let on = self.tri_arg(host, value, |me, h| Ok(json!(me.shape_prop(h, r, Part::Fill, "type")?.as_str() != Some("noFill"))))?;
                if on {
                    if self.shape_prop(host, r, Part::Fill, "type")?.as_str() == Some("noFill") {
                        let c = self.shape_prop(host, r, Part::Fill, "foregroundColor")?;
                        let c = c.as_str().filter(|c| color_from_hex(c).is_some()).unwrap_or("#FFFFFF").to_string();
                        self.call(host, r, Part::Fill, "setSolidColor", json!([c]))?;
                    }
                    Ok(())
                } else {
                    self.call(host, r, Part::Fill, "clear", json!([]))
                }
            }
            (PpObj::Fill(r), "transparency") => {
                let v = finite(value)?;
                if !(0.0..=1.0).contains(&v) {
                    return Err(VbaError::with(5, "Transparency runs from 0 to 1"));
                }
                self.set(host, r, Part::Fill, "transparency", json!(v))
            }
            (PpObj::Line(r), "weight") => {
                let v = finite(value)?;
                if !(0.0..=1584.0).contains(&v) {
                    return Err(VbaError::with(5, "Line weights run from 0 to 1584 points"));
                }
                self.set(host, r, Part::Line, "weight", json!(v))
            }
            (PpObj::Line(r), "visible") => {
                let on = self.tri_arg(host, value, |me, h| me.shape_prop(h, r, Part::Line, "visible"))?;
                self.set(host, r, Part::Line, "visible", json!(on))
            }
            (PpObj::ShapeRange(shapes), p) => match shapes.as_ref() {
                [one] => self.assign(host, &PpObj::Shape(one.clone()), name, value),
                [] => Err(VbaError::with(91, "Nothing appropriate is currently selected (no shapes are selected)")),
                many if matches!(p, "left" | "top" | "width" | "height") => {
                    for r in many {
                        self.assign(host, &PpObj::Shape(r.clone()), name, value)?;
                    }
                    Ok(())
                }
                _ => Err(VbaError::with(438, format!("ShapeRange.{name} with several shapes selected isn't supported yet (use ShapeRange(i))"))),
            },
            (PpObj::SlideRange(slides), _) => match slides.as_ref() {
                [one] => self.assign(host, &PpObj::Slide(one.clone()), name, value),
                _ => Err(VbaError::new(438)),
            },
            (o, _) => Err(VbaError::with(438, format!("{}.{name} can't be set (or isn't supported yet)", o.type_name()))),
        }
    }

    /// `object = value` where the object has a default property (`TextRange.Text`, `ColorFormat.RGB`).
    pub fn assign_default(&mut self, host: &mut dyn Host, o: &PpObj, value: &Value) -> R<()> {
        match o {
            PpObj::TextRange(_) => self.assign(host, o, "text", value),
            PpObj::Color(..) => self.assign(host, o, "rgb", value),
            _ => Err(VbaError::new(438)),
        }
    }

    /// `object(args)`: the default member (`Slides(1)` is `Slides.Item(1)`).
    pub fn default_call(&mut self, host: &mut dyn Host, o: &PpObj, args: Args) -> R<Value> {
        match o {
            PpObj::Slides | PpObj::Shapes(_) | PpObj::ShapeRange(_) | PpObj::SlideRange(_) | PpObj::Presentations => self.member(host, o, "item", args, &[]),
            _ => Err(VbaError::new(438)),
        }
    }

    /// The objects `For Each` walks: slides, shapes, the selected ones.
    pub fn items(&mut self, host: &mut dyn Host, o: &PpObj) -> R<Vec<Value>> {
        let wrap = |x: PpObj| Value::Object(Obj::Pp(x));
        Ok(match o {
            PpObj::Slides => self.slide_ids(host)?.into_iter().map(|id| wrap(PpObj::Slide(id))).collect(),
            PpObj::Shapes(s) => self.shape_list(host, s)?.into_iter().map(|(id, _)| wrap(PpObj::Shape(ShapeRef { slide: Some(s.clone()), id }))).collect(),
            PpObj::ShapeRange(shapes) => shapes.iter().map(|r| wrap(PpObj::Shape(r.clone()))).collect(),
            PpObj::SlideRange(slides) => slides.iter().map(|s| wrap(PpObj::Slide(s.clone()))).collect(),
            PpObj::Presentations => vec![wrap(PpObj::Presentation)],
            _ => return Err(VbaError::new(438)),
        })
    }

    /// The value an object stands for where a value is wanted (a text range's `Text`).
    pub fn default_value(&mut self, host: &mut dyn Host, o: &PpObj) -> R<Value> {
        match o {
            PpObj::TextRange(_) => self.member(host, o, "text", &[], &[]),
            PpObj::Color(..) => self.member(host, o, "rgb", &[], &[]),
            PpObj::Presentation => self.member(host, o, "name", &[], &[]),
            PpObj::Application => Ok(Value::str("DeckCraft")),
            _ => Err(VbaError::new(438)),
        }
    }

    /// The names a macro reaches without an object (`ActivePresentation`, `ActiveWindow`).
    pub fn global(&mut self, host: &mut dyn Host, name: &str, args: Args) -> R<Option<Value>> {
        let n = name.to_ascii_lowercase();
        Ok(Some(match n.as_str() {
            "application" => Value::Object(Obj::Pp(PpObj::Application)),
            "activepresentation" | "activewindow" | "presentations" => self.member(host, &PpObj::Application, name, args, &[])?,
            _ => return Ok(None),
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn constants_and_shapes() {
        assert_eq!(constant("msoShapeRectangle"), Some(Value::Long(1)));
        assert_eq!(constant("msoShapeRoundedRectangle"), Some(Value::Long(5)));
        assert_eq!(constant("msoShapeOval"), Some(Value::Long(9)));
        assert_eq!(constant("msoShapeIsoscelesTriangle"), Some(Value::Long(7)));
        assert_eq!(constant("msoShapeRightArrow"), Some(Value::Long(33)));
        assert_eq!(constant("msoShape5pointStar"), Some(Value::Long(92)));
        assert_eq!(constant("ppLayoutText"), Some(Value::Long(2)));
        assert_eq!(constant("ppLayoutTitle"), Some(Value::Long(1)));
        assert_eq!(constant("ppLayoutBlank"), Some(Value::Long(12)));
        assert_eq!(constant("msoTrue"), Some(Value::Long(-1)));
        assert_eq!(constant("xlUp"), None);
        assert_eq!(geometric(9), Some("ellipse"));
        assert_eq!(geometric(92), Some("star5"));
        assert_eq!(geometric(999), None);
        assert_eq!(layout_number("Blank"), 12);
        assert_eq!(layout_number("title and content"), 2);
        assert_eq!(layout_number("My layout"), 32);
        assert_eq!(text_out("a\nb"), "a\rb");
        assert_eq!(text_in("a\r\nb\rc"), "a\nb\nc");
        assert_eq!(color_arg(&Value::Long(255)).unwrap(), "#FF0000");
        assert_eq!(color_arg(&Value::Long(-1)).unwrap_err().number, 5);
    }
}
