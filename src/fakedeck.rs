//! A presentation in memory that answers the host requests as DeckCraft's Office bridge does
//! (the PowerPoint JavaScript API's paths, props and result shapes, `host/vbaProject/*`,
//! `host/message`, `host/log`). Tests only; it depends on nothing but serde_json, so the
//! end-to-end test includes it too.
#![allow(clippy::unwrap_used, clippy::indexing_slicing, clippy::panic, dead_code)]

use serde_json::{Value as Json, json};

#[derive(Clone, Debug)]
pub struct Shape {
    pub id: u32,
    pub name: String,
    /// `geometricShape`, `textBox`, `line`, `image`.
    pub kind: String,
    /// The geometric shape type (`rectangle`, `ellipse`, …) or the line kind.
    pub geometry: String,
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub size: f64,
    pub font: String,
    pub color: String,
    /// `None` is no fill.
    pub fill: Option<String>,
    pub transparency: f64,
    pub line_color: String,
    pub line_weight: f64,
    pub line_visible: bool,
    pub word_wrap: bool,
    pub auto_size: String,
}

impl Shape {
    pub fn new(id: u32, name: &str, kind: &str, geometry: &str) -> Shape {
        Shape {
            id,
            name: name.into(),
            kind: kind.into(),
            geometry: geometry.into(),
            left: 0.0,
            top: 0.0,
            width: 100.0,
            height: 50.0,
            text: String::new(),
            bold: false,
            italic: false,
            underline: false,
            size: 18.0,
            font: "Calibri".into(),
            color: "#000000".into(),
            fill: if kind == "geometricShape" { Some("#4472C4".into()) } else { None },
            transparency: 0.0,
            line_color: "#2F528F".into(),
            line_weight: 1.0,
            line_visible: kind != "textBox",
            word_wrap: true,
            auto_size: if kind == "textBox" { "autoSizeShapeToFitText".into() } else { "autoSizeNone".into() },
        }
    }

    fn has_text(&self) -> bool {
        matches!(self.kind.as_str(), "geometricShape" | "textBox")
    }
}

#[derive(Clone, Debug)]
pub struct Slide {
    pub id: u32,
    pub layout: u32,
    pub shapes: Vec<Shape>,
}

#[derive(Clone, Debug)]
pub struct Deck {
    pub slides: Vec<Slide>,
    /// Layouts of the one master: (id, name).
    pub layouts: Vec<(u32, String)>,
    /// Answer `slideMasters`/`layouts` (PowerPoint API 1.3); the contract doesn't list them.
    pub layouts_api: bool,
    pub next_id: u32,
    pub title: String,
    pub selected_slides: Vec<u32>,
    pub selected_shapes: Vec<u32>,
    pub vba: Option<String>,
    pub messages: Vec<(String, Option<String>)>,
    pub logs: Vec<String>,
    pub syncs: usize,
    pub actions: usize,
}

impl Default for Deck {
    /// Two slides: a title slide (title and subtitle placeholders) and a blank one.
    fn default() -> Deck {
        let mut title = Shape::new(260, "Title 1", "textBox", "");
        title.text = "Hello".into();
        title.size = 44.0;
        let mut sub = Shape::new(261, "Subtitle 2", "textBox", "");
        sub.text = "First line\nSecond line".into();
        sub.top = 120.0;
        Deck {
            slides: vec![Slide { id: 256, layout: 2147483649, shapes: vec![title, sub] }, Slide { id: 257, layout: 2147483655, shapes: Vec::new() }],
            layouts: vec![(2147483649, "Title Slide".into()), (2147483650, "Title and Content".into()), (2147483655, "Blank".into())],
            layouts_api: true,
            next_id: 300,
            title: "Deck1".into(),
            selected_slides: vec![256],
            selected_shapes: Vec::new(),
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
    Presentation,
    Slides,
    SelectedSlides,
    Slide(u32),
    Layout(u32),
    Masters,
    Master,
    Layouts,
    Shapes(u32),
    SelectedShapes,
    Shape(u32),
    TextFrame(u32),
    TextRange(u32),
    Font(u32),
    Fill(u32),
    Line(u32),
}

type E = (&'static str, String);

fn num_id(j: &Json) -> Option<u32> {
    match j {
        Json::String(s) => s.parse().ok(),
        Json::Number(n) => n.as_u64().map(|n| n as u32),
        _ => None,
    }
}

fn color(j: &Json) -> Result<String, E> {
    let s = j.as_str().ok_or(("InvalidArgument", "a colour is #RRGGBB".to_string()))?;
    if s.len() == 7 && s.starts_with('#') && u32::from_str_radix(&s[1..], 16).is_ok() {
        Ok(s.to_ascii_uppercase())
    } else {
        Err(("InvalidArgument", format!("`{s}` isn't a colour")))
    }
}

impl Deck {
    pub fn slide(&self, i: usize) -> &Slide {
        &self.slides[i]
    }

    /// The shape named `name` on slide `i` (0-based).
    pub fn shape(&self, i: usize, name: &str) -> &Shape {
        self.slides[i].shapes.iter().find(|s| s.name == name).unwrap_or_else(|| panic!("no shape {name}"))
    }

    fn find_shape(&self, id: u32) -> Option<(usize, usize)> {
        self.slides.iter().enumerate().find_map(|(i, s)| s.shapes.iter().position(|x| x.id == id).map(|j| (i, j)))
    }

    fn shape_mut(&mut self, id: u32) -> &mut Shape {
        let (i, j) = self.find_shape(id).unwrap();
        &mut self.slides[i].shapes[j]
    }

    fn shape_ref(&self, id: u32) -> &Shape {
        let (i, j) = self.find_shape(id).unwrap();
        &self.slides[i].shapes[j]
    }

    fn slide_index(&self, id: u32) -> Result<usize, E> {
        self.slides.iter().position(|s| s.id == id).ok_or(("ItemNotFound", format!("no slide {id}")))
    }

    /// The slide the shape selection is on (the first selected slide).
    fn current_slide(&self) -> Option<u32> {
        self.selected_slides.first().copied()
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
                            return Err((-32000, message.clone(), json!({"code": code, "action": i})));
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

    fn resolve(&self, path: &Json) -> Result<Obj, E> {
        let mut o: Option<Obj> = None;
        for step in path.as_array().cloned().unwrap_or_default() {
            let p = step.get("p").and_then(Json::as_str);
            let m = step.get("m").and_then(Json::as_str);
            let a = step.get("a").cloned().unwrap_or(json!([]));
            let not_found = |what: &str| ("ItemNotFound", format!("no {what} {a}"));
            o = Some(match (&o, p, m) {
                (None, Some("presentation"), _) => Obj::Presentation,
                (Some(Obj::Presentation), Some("slides"), _) => Obj::Slides,
                (Some(Obj::Presentation), Some("slideMasters"), _) if self.layouts_api => Obj::Masters,
                (Some(Obj::Masters), _, Some("getItemAt")) if a[0] == json!(0) => Obj::Master,
                (Some(Obj::Master), Some("layouts"), _) => Obj::Layouts,
                (Some(Obj::Presentation), _, Some("getSelectedSlides")) => Obj::SelectedSlides,
                (Some(Obj::Presentation), _, Some("getSelectedShapes")) => Obj::SelectedShapes,
                (Some(Obj::Slides), _, Some("getItem")) => {
                    let id = num_id(&a[0]).filter(|id| self.slides.iter().any(|s| s.id == *id)).ok_or_else(|| not_found("slide"))?;
                    Obj::Slide(id)
                }
                (Some(Obj::Slides), _, Some("getItemAt")) => {
                    let i = a[0].as_u64().ok_or_else(|| not_found("slide"))? as usize;
                    Obj::Slide(self.slides.get(i).ok_or_else(|| not_found("slide"))?.id)
                }
                (Some(Obj::SelectedSlides), _, Some("getItem")) => {
                    let id = num_id(&a[0]).filter(|id| self.selected_slides.contains(id)).ok_or_else(|| not_found("selected slide"))?;
                    Obj::Slide(id)
                }
                (Some(Obj::Slide(s)), Some("shapes"), _) => Obj::Shapes(*s),
                (Some(Obj::Slide(s)), Some("layout"), _) => Obj::Layout(*s),
                (Some(Obj::Shapes(s)), _, Some("getItem")) => {
                    let slide = &self.slides[self.slide_index(*s)?];
                    let id = num_id(&a[0]).filter(|id| slide.shapes.iter().any(|x| x.id == *id)).ok_or_else(|| not_found("shape"))?;
                    Obj::Shape(id)
                }
                (Some(Obj::Shapes(s)), _, Some("getItemAt")) => {
                    let slide = &self.slides[self.slide_index(*s)?];
                    let i = a[0].as_u64().ok_or_else(|| not_found("shape"))? as usize;
                    Obj::Shape(slide.shapes.get(i).ok_or_else(|| not_found("shape"))?.id)
                }
                (Some(Obj::SelectedShapes), _, Some("getItem")) => {
                    let id = num_id(&a[0]).filter(|id| self.selected_shapes.contains(id)).ok_or_else(|| not_found("selected shape"))?;
                    Obj::Shape(id)
                }
                (Some(Obj::Shape(id)), Some("textFrame"), _) => {
                    if !self.shape_ref(*id).has_text() {
                        return Err(("InvalidArgument", "this shape has no text frame".into()));
                    }
                    Obj::TextFrame(*id)
                }
                (Some(Obj::TextFrame(id)), Some("textRange"), _) => Obj::TextRange(*id),
                (Some(Obj::TextRange(id)), Some("font"), _) => Obj::Font(*id),
                (Some(Obj::Shape(id)), Some("fill"), _) => Obj::Fill(*id),
                (Some(Obj::Shape(id)), Some("lineFormat"), _) => Obj::Line(*id),
                (o, p, m) => return Err(("InvalidArgument", format!("{o:?} has no {p:?}/{m:?}"))),
            });
        }
        o.ok_or(("InvalidArgument", "empty path".into()))
    }

    fn shape_ids(&self, o: &Obj) -> Vec<u32> {
        match o {
            Obj::Shapes(s) => self.slides.iter().find(|x| x.id == *s).map(|x| x.shapes.iter().map(|s| s.id).collect()).unwrap_or_default(),
            _ => self.selected_shapes.clone(),
        }
    }

    fn load(&self, obj: &Obj, props: &[String]) -> Result<Json, E> {
        let mut out = serde_json::Map::new();
        let others: Vec<&String> = props.iter().filter(|p| *p != "items").collect();
        for p in props {
            let v = match (obj, p.as_str()) {
                (Obj::Presentation, "title") => json!(self.title),
                (Obj::Slides, "count") => json!(self.slides.len()),
                (Obj::Slides, "items") => json!(self.slides.iter().map(|s| json!({"id": s.id.to_string()})).collect::<Vec<_>>()),
                (Obj::SelectedSlides, "count") => json!(self.selected_slides.len()),
                (Obj::SelectedSlides, "items") => json!(self.selected_slides.iter().map(|s| json!({"id": s.to_string()})).collect::<Vec<_>>()),
                (Obj::Layouts, "items") => json!(
                    self.layouts
                        .iter()
                        .map(|(id, n)| {
                            let mut item = json!({"id": id.to_string()});
                            if others.iter().any(|o| *o == "name") {
                                item["name"] = json!(n);
                            }
                            item
                        })
                        .collect::<Vec<_>>()
                ),
                (Obj::Slide(s), "id") => json!(s.to_string()),
                (Obj::Slide(s), "index") => json!(self.slide_index(*s)?),
                (Obj::Layout(s), "id") => json!(self.slides[self.slide_index(*s)?].layout.to_string()),
                (Obj::Layout(s), "name") => {
                    let l = self.slides[self.slide_index(*s)?].layout;
                    json!(self.layouts.iter().find(|(id, _)| *id == l).map(|(_, n)| n.clone()).unwrap_or_default())
                }
                (Obj::Shapes(_) | Obj::SelectedShapes, "count") => json!(self.shape_ids(obj).len()),
                (Obj::Shapes(_) | Obj::SelectedShapes, "items") => json!(
                    self.shape_ids(obj)
                        .iter()
                        .map(|id| {
                            let s = self.shape_ref(*id);
                            let mut item = json!({"id": id.to_string()});
                            if others.iter().any(|o| *o == "name") {
                                item["name"] = json!(s.name);
                            }
                            item
                        })
                        .collect::<Vec<_>>()
                ),
                (Obj::Shapes(_) | Obj::SelectedShapes | Obj::Slides | Obj::Layouts, "name") => continue,
                (Obj::Shape(id), prop) => {
                    let s = self.shape_ref(*id);
                    match prop {
                        "id" => json!(id.to_string()),
                        "name" => json!(s.name),
                        "type" => json!(s.kind),
                        "left" => json!(s.left),
                        "top" => json!(s.top),
                        "width" => json!(s.width),
                        "height" => json!(s.height),
                        _ => return Err(("InvalidArgument", format!("Shape.{prop}"))),
                    }
                }
                (Obj::TextFrame(id), prop) => {
                    let s = self.shape_ref(*id);
                    match prop {
                        "hasText" => json!(!s.text.is_empty()),
                        "wordWrap" => json!(s.word_wrap),
                        "autoSizeSetting" => json!(s.auto_size),
                        _ => return Err(("InvalidArgument", format!("TextFrame.{prop}"))),
                    }
                }
                (Obj::TextRange(id), prop) => {
                    let s = self.shape_ref(*id);
                    match prop {
                        "text" => json!(s.text),
                        "length" => json!(s.text.chars().count()),
                        "start" => json!(0),
                        _ => return Err(("InvalidArgument", format!("TextRange.{prop}"))),
                    }
                }
                (Obj::Font(id), prop) => {
                    let s = self.shape_ref(*id);
                    match prop {
                        "bold" => json!(s.bold),
                        "italic" => json!(s.italic),
                        "underline" => json!(if s.underline { "single" } else { "none" }),
                        "size" => json!(s.size),
                        "name" => json!(s.font),
                        "color" => json!(s.color),
                        _ => return Err(("InvalidArgument", format!("Font.{prop}"))),
                    }
                }
                (Obj::Fill(id), prop) => {
                    let s = self.shape_ref(*id);
                    match prop {
                        "type" => json!(if s.fill.is_some() { "solid" } else { "noFill" }),
                        "foregroundColor" => json!(s.fill.clone().unwrap_or("#FFFFFF".into())),
                        "transparency" => json!(s.transparency),
                        _ => return Err(("InvalidArgument", format!("Fill.{prop}"))),
                    }
                }
                (Obj::Line(id), prop) => {
                    let s = self.shape_ref(*id);
                    match prop {
                        "color" => json!(s.line_color),
                        "weight" => json!(s.line_weight),
                        "visible" => json!(s.line_visible),
                        _ => return Err(("InvalidArgument", format!("Line.{prop}"))),
                    }
                }
                (o, p) => return Err(("InvalidArgument", format!("{o:?}.{p}"))),
            };
            out.insert(p.clone(), v);
        }
        Ok(Json::Object(out))
    }

    fn set(&mut self, obj: &Obj, prop: &str, value: &Json) -> Result<(), E> {
        let num = || value.as_f64().ok_or(("InvalidArgument", format!("{prop} is a number")));
        let flag = || value.as_bool().ok_or(("InvalidArgument", format!("{prop} is true or false")));
        match obj {
            Obj::Shape(id) => {
                let v = match prop {
                    "name" => {
                        self.shape_mut(*id).name = value.as_str().unwrap_or("").into();
                        return Ok(());
                    }
                    "left" | "top" | "width" | "height" => num()?,
                    _ => return Err(("InvalidArgument", format!("set Shape.{prop}"))),
                };
                let s = self.shape_mut(*id);
                match prop {
                    "left" => s.left = v,
                    "top" => s.top = v,
                    "width" => s.width = v,
                    _ => s.height = v,
                }
            }
            Obj::TextFrame(id) => match prop {
                "wordWrap" => self.shape_mut(*id).word_wrap = flag()?,
                "autoSizeSetting" => self.shape_mut(*id).auto_size = value.as_str().unwrap_or("").into(),
                _ => return Err(("InvalidArgument", format!("set TextFrame.{prop}"))),
            },
            Obj::TextRange(id) if prop == "text" => self.shape_mut(*id).text = value.as_str().unwrap_or("").into(),
            Obj::Font(id) => match prop {
                "bold" => self.shape_mut(*id).bold = flag()?,
                "italic" => self.shape_mut(*id).italic = flag()?,
                "underline" => self.shape_mut(*id).underline = value.as_str() == Some("single"),
                "size" => self.shape_mut(*id).size = num()?,
                "name" => self.shape_mut(*id).font = value.as_str().unwrap_or("").into(),
                "color" => self.shape_mut(*id).color = color(value)?,
                _ => return Err(("InvalidArgument", format!("set Font.{prop}"))),
            },
            Obj::Fill(id) => match prop {
                "foregroundColor" => self.shape_mut(*id).fill = Some(color(value)?),
                "transparency" => self.shape_mut(*id).transparency = num()?,
                _ => return Err(("InvalidArgument", format!("set Fill.{prop}"))),
            },
            Obj::Line(id) => match prop {
                "color" => self.shape_mut(*id).line_color = color(value)?,
                "weight" => self.shape_mut(*id).line_weight = num()?,
                "visible" => self.shape_mut(*id).line_visible = flag()?,
                _ => return Err(("InvalidArgument", format!("set Line.{prop}"))),
            },
            o => return Err(("InvalidArgument", format!("set {o:?}.{prop}"))),
        }
        Ok(())
    }

    fn add(&mut self, slide: u32, kind: &str, geometry: &str, text: &str, opts: &Json) -> Result<Json, E> {
        let i = self.slide_index(slide)?;
        let id = self.next_id;
        self.next_id += 1;
        let label = match (kind, geometry) {
            ("textBox", _) => "TextBox".to_string(),
            ("line", _) => "Straight Connector".to_string(),
            (_, "ellipse") => "Oval".to_string(),
            (_, g) => {
                let mut c = g.chars();
                c.next().map(|f| f.to_ascii_uppercase().to_string() + c.as_str()).unwrap_or_default()
            }
        };
        let mut s = Shape::new(id, &format!("{label} {}", self.slides[i].shapes.len() + 1), kind, geometry);
        s.text = text.into();
        for (k, slot) in [("left", &mut s.left), ("top", &mut s.top), ("width", &mut s.width), ("height", &mut s.height)] {
            if let Some(v) = opts.get(k).and_then(Json::as_f64) {
                *slot = v;
            }
        }
        self.slides[i].shapes.push(s);
        Ok(json!({"id": id.to_string()}))
    }

    fn action(&mut self, a: &Json) -> Result<Json, E> {
        let obj = self.resolve(&a["path"])?;
        match a["op"].as_str() {
            Some("load") => {
                let props: Vec<String> = a["props"].as_array().cloned().unwrap_or_default().iter().filter_map(|p| p.as_str().map(str::to_string)).collect();
                self.load(&obj, &props)
            }
            Some("set") => {
                self.set(&obj, a["prop"].as_str().unwrap_or(""), &a["value"])?;
                Ok(Json::Null)
            }
            Some("call") => {
                let args = a["args"].as_array().cloned().unwrap_or_default();
                let arg = |i: usize| args.get(i).cloned().unwrap_or(Json::Null);
                match (&obj, a["method"].as_str().unwrap_or("")) {
                    (Obj::Slides, "add") => {
                        let opts = arg(0);
                        let index = opts.get("index").and_then(Json::as_u64).map(|i| i as usize).unwrap_or(self.slides.len());
                        if index > self.slides.len() {
                            return Err(("InvalidArgument", "index out of range".into()));
                        }
                        let layout = match opts.get("layoutId") {
                            Some(l) => {
                                num_id(l).filter(|l| self.layouts.iter().any(|(id, _)| id == l)).ok_or(("ItemNotFound", "no such layout".to_string()))?
                            }
                            None => self.layouts[1].0,
                        };
                        let id = self.next_id;
                        self.next_id += 1;
                        self.slides.insert(index, Slide { id, layout, shapes: Vec::new() });
                        Ok(json!({"id": id.to_string()}))
                    }
                    (Obj::Slide(s), "delete") => {
                        let i = self.slide_index(*s)?;
                        self.slides.remove(i);
                        self.selected_slides.retain(|x| x != s);
                        Ok(Json::Null)
                    }
                    (Obj::Slide(s), "setSelectedShapes") => {
                        let i = self.slide_index(*s)?;
                        let ids: Vec<u32> = arg(0).as_array().cloned().unwrap_or_default().iter().filter_map(num_id).collect();
                        if ids.iter().any(|id| !self.slides[i].shapes.iter().any(|x| x.id == *id)) {
                            return Err(("ItemNotFound", "no such shape on this slide".into()));
                        }
                        self.selected_slides = vec![*s];
                        self.selected_shapes = ids;
                        Ok(Json::Null)
                    }
                    (Obj::Presentation, "setSelectedSlides") => {
                        let ids: Vec<u32> = arg(0).as_array().cloned().unwrap_or_default().iter().filter_map(num_id).collect();
                        for id in &ids {
                            self.slide_index(*id)?;
                        }
                        self.selected_slides = ids;
                        self.selected_shapes.clear();
                        Ok(Json::Null)
                    }
                    (Obj::Shapes(s), "addTextBox") => self.add(*s, "textBox", "", arg(0).as_str().unwrap_or(""), &arg(1)),
                    (Obj::Shapes(s), "addGeometricShape") => {
                        let g = arg(0).as_str().unwrap_or("").to_string();
                        if g.is_empty() {
                            return Err(("InvalidArgument", "a geometric shape type".into()));
                        }
                        self.add(*s, "geometricShape", &g, "", &arg(1))
                    }
                    (Obj::Shapes(s), "addLine") => self.add(*s, "line", arg(0).as_str().unwrap_or("straight"), "", &arg(1)),
                    (Obj::Shape(id), "delete") => {
                        let (i, j) = self.find_shape(*id).unwrap();
                        self.slides[i].shapes.remove(j);
                        self.selected_shapes.retain(|x| x != id);
                        Ok(Json::Null)
                    }
                    (Obj::Fill(id), "setSolidColor") => {
                        self.shape_mut(*id).fill = Some(color(&arg(0))?);
                        Ok(Json::Null)
                    }
                    (Obj::Fill(id), "clear") => {
                        self.shape_mut(*id).fill = None;
                        Ok(Json::Null)
                    }
                    (o, m) => Err(("InvalidArgument", format!("{o:?}.{m}()"))),
                }
            }
            _ => Err(("InvalidArgument", "bad op".into())),
        }
    }
}

impl Deck {
    /// The id the shape selection sits on, for tests.
    pub fn current(&self) -> Option<u32> {
        self.current_slide()
    }
}
