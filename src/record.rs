//! `macros/record`: GridCraft's, WordCraft's or DeckCraft's commands, as the app recorded them,
//! written as a VBA Sub (Excel's, Word's or PowerPoint's objects).

use serde_json::Value as Json;

/// Most lines one recording writes.
const MAX_LINES: usize = 10_000;

fn quote(s: &str) -> String {
    format!("\"{}\"", s.replace('"', "\"\""))
}

/// A literal for a cell value: numbers as numbers, the rest quoted.
fn literal(v: &Json) -> String {
    match v {
        Json::Null => "Empty".into(),
        Json::Bool(b) => if *b { "True" } else { "False" }.into(),
        Json::Number(n) => n.to_string(),
        Json::String(s) => quote(s),
        other => quote(&other.to_string()),
    }
}

/// `Range("B2")`, or the selection when the command had no range.
fn target(p: &Json, keys: &[&str]) -> String {
    let addr = keys.iter().find_map(|k| p.get(*k).and_then(Json::as_str));
    let sheet = p.get("sheet").and_then(Json::as_str);
    match (addr, sheet) {
        (Some(a), Some(s)) => format!("Worksheets({}).Range({})", quote(s), quote(a)),
        (Some(a), None) => match a.rsplit_once('!') {
            Some((s, r)) => format!("Worksheets({}).Range({})", quote(s.trim_matches('\'')), quote(r)),
            None => format!("Range({})", quote(a)),
        },
        _ => "Selection".into(),
    }
}

fn rgb(hex: &str) -> Option<String> {
    let h = hex.strip_prefix('#')?;
    let v = u32::from_str_radix(h, 16).ok()?;
    Some(format!("RGB({}, {}, {})", (v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF))
}

fn number_format(f: &str) -> String {
    match f {
        "Number" => "0.00".into(),
        "Currency" | "Accounting" => "$#,##0.00".into(),
        "Percentage" => "0.00%".into(),
        "Short Date" => "m/d/yyyy".into(),
        "Long Date" => "dddd, mmmm d, yyyy".into(),
        "Time" => "h:mm:ss AM/PM".into(),
        "Scientific" => "0.00E+00".into(),
        "Fraction" => "# ?/?".into(),
        "Text" => "@".into(),
        other => other.into(),
    }
}

fn on_off(p: &Json, what: &str, t: &str) -> String {
    match p.get("on").and_then(Json::as_bool) {
        Some(b) => format!("{t}.Font.{what} = {}", if b { "True" } else { "False" }),
        None => format!("{t}.Font.{what} = Not {t}.Font.{what}"),
    }
}

fn line(command: &str, p: &Json, out: &mut Vec<String>) {
    let t = target(p, &["range", "cell"]);
    let s = |k: &str| p.get(k).and_then(Json::as_str).unwrap_or("");
    match command {
        "cell.set" => {
            let t = if p.get("cell").is_some() || p.get("range").is_some() { t } else { "ActiveCell".into() };
            let input = s("input");
            if input.starts_with('=') {
                out.push(format!("{t}.Formula = {}", quote(input)));
            } else if let Ok(n) = input.trim().parse::<f64>() {
                out.push(format!("{t}.Value = {}", if n.fract() == 0.0 && n.abs() < 1e15 { format!("{}", n as i64) } else { format!("{n}") }));
            } else {
                out.push(format!("{t}.Value = {}", quote(input)));
            }
        }
        "range.setValues" => {
            let start = s("range");
            let start = if start.is_empty() { "A1" } else { start };
            let rows = p.get("values").and_then(Json::as_array).cloned().unwrap_or_default();
            let (r0, c0) = crate::excel::parse_a1(start.rsplit('!').next().unwrap_or(start)).map(|r| (r.0, r.1)).unwrap_or((0, 0));
            for (i, row) in rows.iter().enumerate() {
                for (j, v) in row.as_array().into_iter().flatten().enumerate() {
                    let cell = format!("{}{}", crate::excel::col_name(c0 + j as u32), r0 + i as u32 + 1);
                    let prop = if v.as_str().is_some_and(|x| x.starts_with('=')) { "Formula" } else { "Value" };
                    out.push(format!("Range({}).{prop} = {}", quote(&cell), literal(v)));
                }
            }
        }
        "selection.set" => out.push(format!("{t}.Select")),
        "home.bold" => out.push(on_off(p, "Bold", &t)),
        "home.italic" => out.push(on_off(p, "Italic", &t)),
        "home.fontColor" => match rgb(s("color")) {
            Some(c) => out.push(format!("{t}.Font.Color = {c}")),
            None => out.push(format!("{t}.Font.ColorIndex = xlAutomatic")),
        },
        "home.fillColor" => match rgb(s("color")) {
            Some(c) => out.push(format!("{t}.Interior.Color = {c}")),
            None => out.push(format!("{t}.Interior.ColorIndex = xlNone")),
        },
        "home.fontSize" => out.push(format!("{t}.Font.Size = {}", p.get("size").map(|v| v.to_string()).unwrap_or_else(|| "11".into()))),
        "home.fontName" => out.push(format!("{t}.Font.Name = {}", quote(s("name")))),
        "home.numberFormat" => out.push(format!("{t}.NumberFormat = {}", quote(&number_format(s("format"))))),
        "edit.clearContents" => out.push(format!("{t}.ClearContents")),
        "edit.clearFormats" => out.push(format!("{t}.ClearFormats")),
        "edit.clearAll" => out.push(format!("{t}.Clear")),
        "sheet.activate" => match p.get("sheet") {
            Some(Json::Number(n)) => out.push(format!("Worksheets({}).Activate", n.as_i64().unwrap_or(0) + 1)),
            Some(Json::String(n)) => out.push(format!("Worksheets({}).Activate", quote(n))),
            _ => {}
        },
        "home.insertSheet" => match p.get("name").and_then(Json::as_str) {
            Some(n) => out.push(format!("Worksheets.Add.Name = {}", quote(n))),
            None => out.push("Worksheets.Add".into()),
        },
        "sheet.rename" => out.push(format!("ActiveSheet.Name = {}", quote(s("name")))),
        other => out.push(format!("' Not recorded: {other} {}", p).replace(['\r', '\n'], " ")),
    }
}

/// The Sub for `steps` (`[{command, params}]`), called `name`.
pub fn record(name: &str, steps: &[Json]) -> String {
    let mut lines = Vec::new();
    for st in steps {
        if lines.len() >= MAX_LINES {
            lines.push("' (the recording was cut short)".to_string());
            break;
        }
        let command = st.get("command").and_then(Json::as_str).unwrap_or("");
        let params = st.get("params").cloned().unwrap_or(Json::Null);
        line(command, &params, &mut lines);
    }
    let mut src = format!("Sub {name}()\r\n");
    for l in lines {
        src.push_str("    ");
        src.push_str(&l);
        src.push_str("\r\n");
    }
    src.push_str("End Sub\r\n");
    src
}

/// Text as a VBA expression: quoted, with tabs and other control characters as `vbTab`/`Chr(n)`.
fn text_literal(s: &str) -> String {
    let mut parts: Vec<String> = Vec::new();
    let mut run = String::new();
    for c in s.chars() {
        if c.is_control() {
            if !run.is_empty() {
                parts.push(quote(&std::mem::take(&mut run)));
            }
            parts.push(if c == '\t' { "vbTab".to_string() } else { format!("ChrW({})", u32::from(c)) });
        } else {
            run.push(c);
        }
    }
    if !run.is_empty() || parts.is_empty() {
        parts.push(quote(&run));
    }
    parts.join(" & ")
}

fn not_recorded(command: &str, out: &mut Vec<String>) {
    let name: String = command.chars().filter(|c| !c.is_control()).take(200).collect();
    out.push(format!("' (step not recorded: {name})"));
}

/// `True`/`False` for a flag command's `value`, `wdToggle` without one (as Word records a
/// toggle button).
fn word_flag(p: &Json, what: &str, out: &mut Vec<String>) {
    let v = match p.get("value").and_then(Json::as_bool) {
        Some(true) => "True",
        Some(false) => "False",
        None => "wdToggle",
    };
    out.push(format!("Selection.Font.{what} = {v}"));
}

/// The `wdColor…` constant for `RRGGBB`, else `RGB(r, g, b)`.
fn word_color(hex: &str) -> Option<String> {
    let h = hex.trim().trim_start_matches('#');
    if h.len() != 6 {
        return None;
    }
    let v = u32::from_str_radix(h, 16).ok()?;
    let (r, g, b) = ((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
    let named = match (r, g, b) {
        (0, 0, 0) => Some("wdColorBlack"),
        (255, 255, 255) => Some("wdColorWhite"),
        (255, 0, 0) => Some("wdColorRed"),
        (0, 0, 255) => Some("wdColorBlue"),
        (0, 128, 0) => Some("wdColorGreen"),
        (255, 255, 0) => Some("wdColorYellow"),
        _ => None,
    };
    Some(named.map(str::to_string).unwrap_or_else(|| format!("RGB({r}, {g}, {b})")))
}

fn word_highlight(name: &str) -> Option<&'static str> {
    let k = name.replace([' ', '-', '%'], "").to_ascii_lowercase();
    Some(match k.as_str() {
        "none" => "wdNoHighlight",
        "yellow" => "wdYellow",
        "brightgreen" => "wdBrightGreen",
        "turquoise" => "wdTurquoise",
        "pink" => "wdPink",
        "blue" => "wdBlue",
        "red" => "wdRed",
        "darkblue" => "wdDarkBlue",
        "teal" => "wdTeal",
        "green" => "wdGreen",
        "violet" => "wdViolet",
        "darkred" => "wdDarkRed",
        "darkyellow" => "wdDarkYellow",
        "gray50" => "wdGray50",
        "gray25" => "wdGray25",
        "black" => "wdBlack",
        "white" => "wdWhite",
        _ => return None,
    })
}

fn word_underline(style: &str) -> Option<&'static str> {
    Some(match style.to_ascii_lowercase().as_str() {
        "none" => "wdUnderlineNone",
        "single" => "wdUnderlineSingle",
        "words" | "word" => "wdUnderlineWords",
        "double" => "wdUnderlineDouble",
        "dotted" => "wdUnderlineDotted",
        "thick" => "wdUnderlineThick",
        "dash" => "wdUnderlineDash",
        "dotdash" => "wdUnderlineDotDash",
        "dotdotdash" => "wdUnderlineDotDotDash",
        "wave" | "wavy" => "wdUnderlineWavy",
        _ => return None,
    })
}

fn align(value: &str) -> Option<&'static str> {
    Some(match value {
        "left" => "wdAlignParagraphLeft",
        "center" => "wdAlignParagraphCenter",
        "right" => "wdAlignParagraphRight",
        "justify" | "both" => "wdAlignParagraphJustify",
        _ => return None,
    })
}

fn word_line(command: &str, p: &Json, out: &mut Vec<String>) {
    let s = |k: &str| p.get(k).and_then(Json::as_str);
    let extend = p.get("extend").and_then(Json::as_bool) == Some(true);
    match command {
        "text.insert" => {
            let text = s("text").unwrap_or("").replace("\r\n", "\n").replace('\r', "\n");
            for (i, part) in text.split('\n').enumerate() {
                if i > 0 {
                    out.push("Selection.TypeParagraph".into());
                }
                if !part.is_empty() {
                    out.push(format!("Selection.TypeText Text:={}", text_literal(part)));
                }
            }
        }
        "text.newParagraph" => out.push("Selection.TypeParagraph".into()),
        "format.bold" => word_flag(p, "Bold", out),
        "format.italic" => word_flag(p, "Italic", out),
        "format.underline" => {
            let on = p.get("value").and_then(Json::as_bool);
            match (on, s("style")) {
                (Some(false), _) => out.push("Selection.Font.Underline = wdUnderlineNone".into()),
                (_, Some(style)) => match word_underline(style) {
                    Some(c) => out.push(format!("Selection.Font.Underline = {c}")),
                    None => not_recorded(command, out),
                },
                (Some(true), None) => out.push("Selection.Font.Underline = wdUnderlineSingle".into()),
                (None, None) => out.push(
                    "If Selection.Font.Underline = wdUnderlineNone Then Selection.Font.Underline = wdUnderlineSingle Else Selection.Font.Underline = wdUnderlineNone"
                        .into(),
                ),
            }
        }
        "format.doubleUnderline" => out.push("Selection.Font.Underline = wdUnderlineDouble".into()),
        "format.size" => match p.get("size").and_then(Json::as_f64).filter(|v| v.is_finite()) {
            Some(v) => out.push(format!("Selection.Font.Size = {v}")),
            None => not_recorded(command, out),
        },
        "format.font" => match s("name").filter(|n| !n.trim().is_empty()) {
            Some(n) => out.push(format!("Selection.Font.Name = {}", quote(n))),
            None => not_recorded(command, out),
        },
        "format.color" => match s("color").and_then(word_color) {
            Some(c) => out.push(format!("Selection.Font.Color = {c}")),
            None => not_recorded(&format!("{command} {}", s("color").unwrap_or("")), out),
        },
        "format.highlight" => match word_highlight(s("color").unwrap_or("yellow")) {
            Some(c) => out.push(format!("Selection.Range.HighlightColorIndex = {c}")),
            None => not_recorded(command, out),
        },
        "para.style" => match s("style").filter(|n| !n.trim().is_empty()) {
            Some(n) => out.push(format!("Selection.Style = ActiveDocument.Styles({})", quote(n))),
            None => not_recorded(command, out),
        },
        "para.alignLeft" | "para.alignCenter" | "para.alignRight" | "para.justify" => {
            let v = match command {
                "para.alignLeft" => "wdAlignParagraphLeft",
                "para.alignCenter" => "wdAlignParagraphCenter",
                "para.alignRight" => "wdAlignParagraphRight",
                _ => "wdAlignParagraphJustify",
            };
            out.push(format!("Selection.ParagraphFormat.Alignment = {v}"));
        }
        "para.align" => match s("value").and_then(align) {
            Some(v) => out.push(format!("Selection.ParagraphFormat.Alignment = {v}")),
            None => not_recorded(command, out),
        },
        "caret.docStart" if !extend => out.push("Selection.HomeKey Unit:=wdStory".into()),
        "caret.docEnd" if !extend => out.push("Selection.EndKey Unit:=wdStory".into()),
        "select.all" => out.push("Selection.WholeStory".into()),
        other => not_recorded(other, out),
    }
}

/// The Sub for WordCraft's `steps` (`[{command, params}]`), called `name`.
pub fn record_word(name: &str, steps: &[Json]) -> String {
    let mut lines = Vec::new();
    for st in steps {
        if lines.len() >= MAX_LINES {
            lines.push("' (the recording was cut short)".to_string());
            break;
        }
        let command = st.get("command").and_then(Json::as_str).unwrap_or("");
        let params = st.get("params").cloned().unwrap_or(Json::Null);
        word_line(command, &params, &mut lines);
    }
    lines.truncate(MAX_LINES + 1);
    let mut src = format!("Sub {name}()\r\n");
    for l in lines {
        src.push_str("    ");
        src.push_str(&l);
        src.push_str("\r\n");
    }
    src.push_str("End Sub\r\n");
    src
}

/// A number of points (or a font size) as a VBA literal: finite, within reason, 4 decimals.
fn pp_f(v: f64) -> Option<String> {
    if !v.is_finite() || v.abs() > 1e9 {
        return None;
    }
    let r = (v * 10_000.0).round() / 10_000.0;
    Some(format!("{}", if r == 0.0 { 0.0 } else { r }))
}

fn pp_num(v: Option<&Json>) -> Option<String> {
    pp_f(v?.as_f64()?)
}

/// [`pp_num`] for a value that must lie in `lo..=hi`.
fn pp_in(v: Option<&Json>, lo: f64, hi: f64) -> Option<String> {
    pp_f(v?.as_f64().filter(|v| (lo..=hi).contains(v))?)
}

/// `RGB(r, g, b)` for `#RRGGBB` (the colours a craft app journals); theme colours aren't.
fn pp_rgb(hex: Option<&str>) -> Option<String> {
    let h = hex?.trim().trim_start_matches('#');
    if h.len() != 6 || !h.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }
    let v = u32::from_str_radix(h, 16).ok()?;
    Some(format!("RGB({}, {}, {})", (v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF))
}

/// Characters per text literal, so a recorded line stays well under VBA's 1023.
const PP_TEXT_CHUNK: usize = 200;

/// Text as PowerPoint's VBA writes it: paragraphs end with `vbCr`, line breaks are `Chr(11)`.
fn pp_text_literal(s: &str) -> String {
    let mut parts: Vec<String> = Vec::new();
    let mut run = String::new();
    for c in s.chars() {
        if c.is_control() {
            if !run.is_empty() {
                parts.push(quote(&std::mem::take(&mut run)));
            }
            parts.push(match c {
                '\n' => "vbCr".to_string(),
                '\t' => "vbTab".to_string(),
                c => format!("ChrW({})", u32::from(c)),
            });
        } else {
            run.push(c);
        }
    }
    if !run.is_empty() || parts.is_empty() {
        parts.push(quote(&run));
    }
    parts.join(" & ")
}

/// `range.Text = …` (or `range.InsertAfter …` when `replace` is false), in chunks.
fn pp_text(range: &str, text: &str, replace: bool, out: &mut Vec<String>) {
    let text = text.replace("\r\n", "\n").replace('\r', "\n");
    let chars: Vec<char> = text.chars().collect();
    if chars.is_empty() {
        if replace {
            out.push(format!("{range}.Text = \"\""));
        }
        return;
    }
    for (i, chunk) in chars.chunks(PP_TEXT_CHUNK).enumerate() {
        if out.len() > MAX_LINES {
            return;
        }
        let lit = pp_text_literal(&chunk.iter().collect::<String>());
        if i == 0 && replace {
            out.push(format!("{range}.Text = {lit}"));
        } else {
            out.push(format!("{range}.InsertAfter {lit}"));
        }
    }
}

/// A shape on the slide showing, by its id.
fn pp_shape_by_id(id: u64) -> String {
    format!("ActiveWindow.View.Slide.Shapes.FindById({id})")
}

/// The shapes a command acts on: those in `ids` (or `id`), else the selected ones. `None` when
/// `ids` was given but names no shape.
fn pp_targets(p: &Json) -> Option<Vec<String>> {
    let ids: Option<Vec<u64>> = match (p.get("ids"), p.get("id")) {
        (Some(Json::Array(a)), _) => Some(a.iter().filter_map(Json::as_u64).take(MAX_LINES).collect()),
        (Some(Json::Null) | None, Some(v)) => v.as_u64().map(|v| vec![v]).or(Some(Vec::new())),
        (Some(_), _) => Some(Vec::new()),
        (None, None) => None,
    };
    match ids {
        None => Some(vec!["ActiveWindow.Selection.ShapeRange".to_string()]),
        Some(ids) if ids.is_empty() => None,
        Some(ids) => Some(ids.into_iter().map(pp_shape_by_id).collect()),
    }
}

/// `msoTrue`/`msoFalse` for a flag command's `on`, `msoTriStateToggle` without one.
fn pp_flag(p: &Json, what: &str, out: &mut Vec<String>) {
    let v = match p.get("on").and_then(Json::as_bool) {
        Some(true) => "msoTrue",
        Some(false) => "msoFalse",
        None => "msoTriStateToggle",
    };
    out.push(format!("ActiveWindow.Selection.TextRange.Font.{what} = {v}"));
}

/// The `ppLayout…` constant for a DeckCraft layout name or kind.
fn pp_layout(key: &str) -> Option<&'static str> {
    let k: String = key.chars().filter(|c| c.is_alphanumeric()).collect::<String>().to_ascii_lowercase();
    Some(match k.as_str() {
        "title" | "titleslide" => "ppLayoutTitle",
        "titleandcontent" | "content" | "obj" | "titleandtext" => "ppLayoutText",
        "sectionheader" | "section" => "ppLayoutSectionHeader",
        "twocontent" | "two" | "twocolumntext" => "ppLayoutTwoColumnText",
        "comparison" => "ppLayoutComparison",
        "titleonly" => "ppLayoutTitleOnly",
        "blank" => "ppLayoutBlank",
        "contentwithcaption" => "ppLayoutContentWithCaption",
        "picturewithcaption" | "picture" => "ppLayoutPictureWithCaption",
        _ => return None,
    })
}

/// The `msoShape…` constant for a DeckCraft preset (the ones craft-vba draws back).
fn pp_auto_shape(preset: &str) -> Option<&'static str> {
    Some(match preset {
        "rect" => "msoShapeRectangle",
        "roundRect" => "msoShapeRoundedRectangle",
        "ellipse" => "msoShapeOval",
        "triangle" => "msoShapeIsoscelesTriangle",
        "rtTriangle" => "msoShapeRightTriangle",
        "diamond" => "msoShapeDiamond",
        "parallelogram" => "msoShapeParallelogram",
        "trapezoid" => "msoShapeTrapezoid",
        "octagon" => "msoShapeOctagon",
        "hexagon" => "msoShapeHexagon",
        "pentagon" => "msoShapeRegularPentagon",
        "plus" => "msoShapeCross",
        "can" => "msoShapeCan",
        "cube" => "msoShapeCube",
        "bevel" => "msoShapeBevel",
        "foldedCorner" => "msoShapeFoldedCorner",
        "smileyFace" => "msoShapeSmileyFace",
        "donut" => "msoShapeDonut",
        "noSmoking" => "msoShapeNoSymbol",
        "blockArc" => "msoShapeBlockArc",
        "heart" => "msoShapeHeart",
        "lightningBolt" => "msoShapeLightningBolt",
        "sun" => "msoShapeSun",
        "moon" => "msoShapeMoon",
        "rightArrow" => "msoShapeRightArrow",
        "leftArrow" => "msoShapeLeftArrow",
        "upArrow" => "msoShapeUpArrow",
        "downArrow" => "msoShapeDownArrow",
        "leftRightArrow" => "msoShapeLeftRightArrow",
        "upDownArrow" => "msoShapeUpDownArrow",
        "homePlate" => "msoShapePentagon",
        "chevron" => "msoShapeChevron",
        "flowChartProcess" => "msoShapeFlowchartProcess",
        "star4" => "msoShape4pointStar",
        "star5" => "msoShape5pointStar",
        "star8" => "msoShape8pointStar",
        "star16" => "msoShape16pointStar",
        "star24" => "msoShape24pointStar",
        "star32" => "msoShape32pointStar",
        _ => return None,
    })
}

/// `[x, y, w, h]` in points, as literals; DeckCraft's default box (centred on a 960×540 slide)
/// when the step has none.
fn pp_rect(p: &Json, default: (f64, f64)) -> Option<[String; 4]> {
    let [x, y, w, h] = match p.get("rect") {
        None | Some(Json::Null) => [(960.0 - default.0) / 2.0, (540.0 - default.1) / 2.0, default.0, default.1],
        Some(r) => {
            let f = |i: usize| r.get(i).and_then(Json::as_f64);
            [f(0)?, f(1)?, f(2)?.abs(), f(3)?.abs()]
        }
    };
    Some([pp_f(x)?, pp_f(y)?, pp_f(w)?, pp_f(h)?])
}

/// The new shape's text, typed into it once it's selected.
fn pp_new_text(p: &Json, out: &mut Vec<String>) {
    if let Some(t) = p.get("text").and_then(Json::as_str).filter(|t| !t.is_empty()) {
        pp_text("ActiveWindow.Selection.TextRange", t, true, out);
    }
}

/// A 1-based slide number for a 0-based `index`.
fn pp_index(p: &Json, key: &str) -> Option<u64> {
    p.get(key)?.as_u64()?.checked_add(1)
}

fn pp_line(command: &str, p: &Json, out: &mut Vec<String>) {
    let s = |k: &str| p.get(k).and_then(Json::as_str);
    let before = out.len();
    match command {
        "slide.new" => {
            let at = match p.get("at") {
                None | Some(Json::Null) => Some("ActiveWindow.View.Slide.SlideIndex + 1".to_string()),
                Some(v) => v.as_u64().and_then(|v| v.checked_add(1)).map(|v| v.to_string()),
            };
            let layout = match p.get("layout") {
                None | Some(Json::Null) => {
                    // As the app does: Title and Content after a title slide, else the same layout.
                    Some("IIf(ActiveWindow.View.Slide.Layout = ppLayoutTitle, ppLayoutText, ActiveWindow.View.Slide.Layout)".to_string())
                }
                Some(v) => v.as_str().and_then(pp_layout).map(str::to_string),
            };
            let (Some(at), Some(layout)) = (at, layout) else { return not_recorded(command, out) };
            out.push(format!("ActiveWindow.View.GotoSlide ActivePresentation.Slides.Add({at}, {layout}).SlideIndex"));
            if let Some(t) = s("title") {
                pp_text("ActiveWindow.View.Slide.Shapes(1).TextFrame.TextRange", t, true, out);
            }
            if let Some(b) = s("body") {
                pp_text("ActiveWindow.View.Slide.Shapes(2).TextFrame.TextRange", b, true, out);
            }
        }
        "slide.delete" => match (p.get("ids").and_then(Json::as_array), pp_index(p, "index")) {
            (Some(ids), _) => {
                for id in ids.iter().filter_map(Json::as_u64).take(MAX_LINES) {
                    out.push(format!("ActivePresentation.Slides.FindBySlideID({id}).Delete"));
                }
            }
            (None, Some(i)) => out.push(format!("ActivePresentation.Slides({i}).Delete")),
            (None, None) => out.push("ActiveWindow.Selection.SlideRange.Delete".into()),
        },
        "slide.go" => match (pp_index(p, "index"), p.get("id").and_then(Json::as_u64)) {
            (Some(i), _) => out.push(format!("ActiveWindow.View.GotoSlide {i}")),
            (None, Some(id)) => out.push(format!("ActiveWindow.View.GotoSlide ActivePresentation.Slides.FindBySlideID({id}).SlideIndex")),
            _ => {}
        },
        "slide.next" => out.push(
            "If ActiveWindow.View.Slide.SlideIndex < ActivePresentation.Slides.Count Then ActiveWindow.View.GotoSlide ActiveWindow.View.Slide.SlideIndex + 1"
                .into(),
        ),
        "slide.previous" => {
            out.push("If ActiveWindow.View.Slide.SlideIndex > 1 Then ActiveWindow.View.GotoSlide ActiveWindow.View.Slide.SlideIndex - 1".into())
        }
        "slide.first" => out.push("ActiveWindow.View.GotoSlide 1".into()),
        "slide.last" => out.push("ActiveWindow.View.GotoSlide ActivePresentation.Slides.Count".into()),
        "slide.selectSlides" if p.get("add").and_then(Json::as_bool) != Some(true) => {
            // One slide picked (several can't be selected from VBA here).
            let one = |k: &str| match p.get(k).and_then(Json::as_array).map(Vec::as_slice) {
                Some([v]) => v.as_u64(),
                _ => None,
            };
            match (p.get("ids"), one("ids"), one("indices").and_then(|i| i.checked_add(1))) {
                (Some(_), Some(id), None) if p.get("indices").is_none() => out.push(format!("ActivePresentation.Slides.FindBySlideID({id}).Select")),
                (None, None, Some(i)) => out.push(format!("ActivePresentation.Slides({i}).Select")),
                _ => {}
            }
        }
        "insert.textBox" => {
            // A click-to-create text box (no width) starts 100 points wide.
            let narrow = p.get("rect").and_then(|r| r.get(2)).and_then(Json::as_f64).is_some_and(|w| w.abs() < 1.0);
            let Some([x, y, w, h]) = pp_rect(p, (216.0, 36.0)) else { return not_recorded(command, out) };
            let w = if narrow { "100".to_string() } else { w };
            out.push(format!("ActiveWindow.View.Slide.Shapes.AddTextbox(msoTextOrientationHorizontal, {x}, {y}, {w}, {h}).Select"));
            pp_new_text(p, out);
        }
        "shape.insert" => {
            let preset = s("preset").unwrap_or("rect");
            if preset == "line" {
                let f = |i: usize| p.get("rect").and_then(|r| r.get(i)).and_then(Json::as_f64);
                let ends = (|| Some((pp_f(f(0)?)?, pp_f(f(1)?)?, pp_f(f(0)? + f(2)?.abs())?, pp_f(f(1)? + f(3)?.abs())?)))();
                let Some((x1, y1, x2, y2)) = ends else { return not_recorded(command, out) };
                out.push(format!("ActiveWindow.View.Slide.Shapes.AddLine({x1}, {y1}, {x2}, {y2}).Select"));
                return;
            }
            let (Some(kind), Some([x, y, w, h])) = (pp_auto_shape(preset), pp_rect(p, (72.0, 72.0))) else {
                return not_recorded(&format!("{command} {preset}"), out);
            };
            out.push(format!("ActiveWindow.View.Slide.Shapes.AddShape({kind}, {x}, {y}, {w}, {h}).Select"));
            pp_new_text(p, out);
        }
        "text.insert" => pp_text("ActiveWindow.Selection.TextRange", s("text").unwrap_or(""), false, out),
        "text.set" if p.get("cell").is_none() => {
            let range = match p.get("id").and_then(Json::as_u64) {
                Some(id) => format!("{}.TextFrame.TextRange", pp_shape_by_id(id)),
                None => "ActiveWindow.Selection.TextRange".to_string(),
            };
            pp_text(&range, s("text").unwrap_or(""), true, out);
        }
        // Editing a shape's text selects it; the object model needs nothing more.
        "text.edit" if p.get("cell").is_none() && p.get("notes").is_none() => {
            if let Some(id) = p.get("id").and_then(Json::as_u64) {
                out.push(format!("{}.Select", pp_shape_by_id(id)));
            }
            return;
        }
        "text.exit" => return,
        "format.bold" => pp_flag(p, "Bold", out),
        "format.italic" => pp_flag(p, "Italic", out),
        "format.underline" => {
            let v = match (p.get("on").and_then(Json::as_bool), s("style")) {
                (Some(false), _) | (_, Some("none")) => "msoFalse",
                (Some(true), _) | (None, Some(_)) => "msoTrue",
                (None, None) => "msoTriStateToggle",
            };
            out.push(format!("ActiveWindow.Selection.TextRange.Font.Underline = {v}"));
        }
        "format.size" => {
            if let Some(v) = pp_in(p.get("size"), 1.0, 4000.0) {
                out.push(format!("ActiveWindow.Selection.TextRange.Font.Size = {v}"));
            }
        }
        "format.font" => {
            if let Some(n) = s("family").filter(|n| !n.trim().is_empty() && !n.chars().any(char::is_control)) {
                out.push(format!("ActiveWindow.Selection.TextRange.Font.Name = {}", quote(n)));
            }
        }
        "format.color" => {
            if let Some(c) = pp_rgb(s("color")) {
                out.push(format!("ActiveWindow.Selection.TextRange.Font.Color.RGB = {c}"));
            }
        }
        "shape.fill" => {
            let Some(targets) = pp_targets(p) else { return not_recorded(command, out) };
            for t in targets {
                if p.get("none").and_then(Json::as_bool) == Some(true) {
                    out.push(format!("{t}.Fill.Visible = msoFalse"));
                } else if let Some(c) = pp_rgb(s("color")) {
                    out.push(format!("{t}.Fill.ForeColor.RGB = {c}"));
                }
                if let Some(v) = pp_in(p.get("transparency"), 0.0, 1.0) {
                    out.push(format!("{t}.Fill.Transparency = {v}"));
                }
            }
        }
        "shape.line" => {
            let Some(targets) = pp_targets(p) else { return not_recorded(command, out) };
            for t in targets {
                if p.get("none").and_then(Json::as_bool) == Some(true) {
                    out.push(format!("{t}.Line.Visible = msoFalse"));
                } else if let Some(c) = pp_rgb(s("color")) {
                    out.push(format!("{t}.Line.ForeColor.RGB = {c}"));
                }
                if let Some(v) = pp_in(p.get("width"), 0.0, 1584.0) {
                    out.push(format!("{t}.Line.Weight = {v}"));
                }
            }
        }
        "shape.move" | "arrange.nudge" => {
            let Some(targets) = pp_targets(p) else { return not_recorded(command, out) };
            for t in targets {
                for (key, line) in [("dx", "IncrementLeft "), ("dy", "IncrementTop "), ("x", "Left = "), ("y", "Top = ")] {
                    if let Some(v) = pp_num(p.get(key)) {
                        out.push(format!("{t}.{line}{v}"));
                    }
                }
            }
        }
        "shape.resize" => {
            let Some(targets) = pp_targets(p) else { return not_recorded(command, out) };
            let lock = p.get("lockAspect").and_then(Json::as_bool) == Some(true);
            let (w, h) = (pp_in(p.get("w"), 0.0, 1e9), pp_in(p.get("h"), 0.0, 1e9));
            for t in targets {
                match (w.as_deref(), h.as_deref(), lock) {
                    (Some(w), _, true) => {
                        out.push(format!("{t}.Height = {t}.Height * {w} / {t}.Width"));
                        out.push(format!("{t}.Width = {w}"));
                    }
                    (None, Some(h), true) => {
                        out.push(format!("{t}.Width = {t}.Width * {h} / {t}.Height"));
                        out.push(format!("{t}.Height = {h}"));
                    }
                    _ => {
                        if let Some(w) = &w {
                            out.push(format!("{t}.Width = {w}"));
                        }
                        if let Some(h) = &h {
                            out.push(format!("{t}.Height = {h}"));
                        }
                    }
                }
            }
        }
        "shape.setBounds" => {
            let Some(targets) = pp_targets(p) else { return not_recorded(command, out) };
            let vals =
                [("Left", pp_num(p.get("x"))), ("Top", pp_num(p.get("y"))), ("Width", pp_in(p.get("w"), 0.0, 1e9)), ("Height", pp_in(p.get("h"), 0.0, 1e9))];
            for t in targets {
                for (prop, v) in &vals {
                    if let Some(v) = v {
                        out.push(format!("{t}.{prop} = {v}"));
                    }
                }
            }
        }
        "shape.rename" => {
            if let Some(n) = s("name").filter(|n| !n.chars().any(char::is_control)) {
                let t = match p.get("id").and_then(Json::as_u64) {
                    Some(id) => pp_shape_by_id(id),
                    None => "ActiveWindow.Selection.ShapeRange".to_string(),
                };
                out.push(format!("{t}.Name = {}", quote(n)));
            }
        }
        "edit.select" | "selection.select" if p.get("toggle").and_then(Json::as_bool) != Some(true) => {
            let add = p.get("add").and_then(Json::as_bool) == Some(true);
            match pp_targets(p) {
                Some(targets) if p.get("ids").is_some() || p.get("id").is_some() => {
                    for (i, t) in targets.into_iter().enumerate() {
                        out.push(if i == 0 && !add { format!("{t}.Select") } else { format!("{t}.Select msoFalse") });
                    }
                }
                // No shapes: nothing selected any more.
                None if !add && p.get("ids").and_then(Json::as_array).is_some_and(Vec::is_empty) => out.push("ActiveWindow.Selection.Unselect".into()),
                _ => {}
            }
        }
        "edit.deselect" => out.push("ActiveWindow.Selection.Unselect".into()),
        "edit.delete" => {
            if p.get("ids").is_none() && s("scope") == Some("slides") {
                out.push("ActiveWindow.Selection.SlideRange.Delete".into());
            } else if let Some(targets) = pp_targets(p) {
                for t in targets {
                    out.push(format!("{t}.Delete"));
                }
            }
        }
        other => return not_recorded(other, out),
    }
    if out.len() == before {
        not_recorded(command, out);
    }
}

/// The Sub for DeckCraft's `steps` (`[{command, params}]`), called `name`, as PowerPoint VBA.
/// PowerPoint has no macro recorder (it went in 2007); this one is a craft extra. Typing that
/// arrives as several `text.insert` steps in a row is written as one.
pub fn record_powerpoint(name: &str, steps: &[Json]) -> String {
    let mut lines = Vec::new();
    let mut typed = String::new();
    for st in steps {
        if lines.len() >= MAX_LINES {
            lines.push("' (the recording was cut short)".to_string());
            typed.clear();
            break;
        }
        let command = st.get("command").and_then(Json::as_str).unwrap_or("");
        let params = st.get("params").cloned().unwrap_or(Json::Null);
        if command == "text.insert" {
            typed.push_str(params.get("text").and_then(Json::as_str).unwrap_or(""));
            continue;
        }
        if !typed.is_empty() {
            pp_text("ActiveWindow.Selection.TextRange", &std::mem::take(&mut typed), false, &mut lines);
        }
        pp_line(command, &params, &mut lines);
    }
    if !typed.is_empty() {
        pp_text("ActiveWindow.Selection.TextRange", &typed, false, &mut lines);
    }
    lines.truncate(MAX_LINES + 1);
    let mut src = format!("Sub {name}()\r\n");
    for l in lines {
        src.push_str("    ");
        src.push_str(&l);
        src.push_str("\r\n");
    }
    src.push_str("End Sub\r\n");
    src
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    #[test]
    fn commands_become_vba() {
        let src = super::record(
            "Macro1",
            &[
                json!({"command": "cell.set", "params": {"cell": "B2", "input": "5"}}),
                json!({"command": "cell.set", "params": {"cell": "B3", "input": "=B2*2"}}),
                json!({"command": "cell.set", "params": {"input": "say \"hi\""}}),
                json!({"command": "range.setValues", "params": {"range": "C1", "values": [[1, "x"]]}}),
                json!({"command": "selection.set", "params": {"range": "A1:B2"}}),
                json!({"command": "home.bold", "params": {"on": true}}),
                json!({"command": "home.fillColor", "params": {"range": "A1", "color": "#FF0000"}}),
                json!({"command": "home.numberFormat", "params": {"range": "A1", "format": "Percentage"}}),
                json!({"command": "sheet.activate", "params": {"sheet": 1}}),
                json!({"command": "home.insertSheet", "params": {"name": "Report"}}),
                json!({"command": "chart.insert", "params": {}}),
            ],
        );
        let expected = "Sub Macro1()\r\n    Range(\"B2\").Value = 5\r\n    Range(\"B3\").Formula = \"=B2*2\"\r\n    ActiveCell.Value = \"say \"\"hi\"\"\"\r\n    \
            Range(\"C1\").Value = 1\r\n    Range(\"D1\").Value = \"x\"\r\n    Range(\"A1:B2\").Select\r\n    Selection.Font.Bold = True\r\n    \
            Range(\"A1\").Interior.Color = RGB(255, 0, 0)\r\n    Range(\"A1\").NumberFormat = \"0.00%\"\r\n    Worksheets(2).Activate\r\n    \
            Worksheets.Add.Name = \"Report\"\r\n    ' Not recorded: chart.insert {}\r\nEnd Sub\r\n";
        assert_eq!(src, expected);
        // What it records parses.
        assert!(crate::vba::parser::parse_module("M", &src).is_ok());
    }

    #[test]
    fn wordcraft_steps_become_word_vba() {
        let src = super::record_word(
            "Macro1",
            &[
                json!({"command": "text.insert", "params": {"text": "Say \"hi\"\tnow"}}),
                json!({"command": "text.newParagraph", "params": {}}),
                json!({"command": "text.insert", "params": {"text": "a\rb"}}),
                json!({"command": "format.bold", "params": {"value": true}}),
                json!({"command": "format.italic", "params": {"value": false}}),
                json!({"command": "format.bold", "params": {}}),
                json!({"command": "format.underline", "params": {"value": true}}),
                json!({"command": "format.underline", "params": {"style": "double"}}),
                json!({"command": "format.underline", "params": {}}),
                json!({"command": "format.size", "params": {"size": 14}}),
                json!({"command": "format.size", "params": {"size": 10.5}}),
                json!({"command": "format.font", "params": {"name": "Georgia"}}),
                json!({"command": "format.color", "params": {"color": "FF0000"}}),
                json!({"command": "format.color", "params": {"color": "1A2B3C"}}),
                json!({"command": "format.color", "params": {"color": "auto"}}),
                json!({"command": "format.highlight", "params": {"color": "brightGreen"}}),
                json!({"command": "para.style", "params": {"style": "Heading 1"}}),
                json!({"command": "para.alignCenter"}),
                json!({"command": "para.justify", "params": {}}),
                json!({"command": "para.align", "params": {"value": "right"}}),
                json!({"command": "caret.left", "params": {}}),
                json!({"command": "caret.docStart", "params": {}}),
                json!({"command": "caret.docEnd", "params": {"extend": true}}),
                json!({"command": "select.all", "params": {}}),
                json!({"command": "table.insert\r\nEvil", "params": {"rows": 2}}),
                json!({"params": {}}),
            ],
        );
        let expected = [
            "Sub Macro1()",
            "    Selection.TypeText Text:=\"Say \"\"hi\"\"\" & vbTab & \"now\"",
            "    Selection.TypeParagraph",
            "    Selection.TypeText Text:=\"a\"",
            "    Selection.TypeParagraph",
            "    Selection.TypeText Text:=\"b\"",
            "    Selection.Font.Bold = True",
            "    Selection.Font.Italic = False",
            "    Selection.Font.Bold = wdToggle",
            "    Selection.Font.Underline = wdUnderlineSingle",
            "    Selection.Font.Underline = wdUnderlineDouble",
            "    If Selection.Font.Underline = wdUnderlineNone Then Selection.Font.Underline = wdUnderlineSingle Else Selection.Font.Underline = wdUnderlineNone",
            "    Selection.Font.Size = 14",
            "    Selection.Font.Size = 10.5",
            "    Selection.Font.Name = \"Georgia\"",
            "    Selection.Font.Color = wdColorRed",
            "    Selection.Font.Color = RGB(26, 43, 60)",
            "    ' (step not recorded: format.color auto)",
            "    Selection.Range.HighlightColorIndex = wdBrightGreen",
            "    Selection.Style = ActiveDocument.Styles(\"Heading 1\")",
            "    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter",
            "    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify",
            "    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight",
            "    ' (step not recorded: caret.left)",
            "    Selection.HomeKey Unit:=wdStory",
            "    ' (step not recorded: caret.docEnd)",
            "    Selection.WholeStory",
            "    ' (step not recorded: table.insertEvil)",
            "    ' (step not recorded: )",
            "End Sub",
            "",
        ]
        .join("\r\n");
        assert_eq!(src, expected);
        assert!(crate::vba::parser::parse_module("M", &src).is_ok());
        assert_eq!(super::text_literal("\u{b}"), "ChrW(11)");
        assert_eq!(super::text_literal(""), "\"\"");
        let many: Vec<_> = (0..20_000).map(|_| json!({"command": "text.newParagraph"})).collect();
        assert!(super::record_word("M", &many).lines().count() < 10_010);
    }

    #[test]
    fn deckcraft_steps_become_powerpoint_vba() {
        let src = super::record_powerpoint(
            "Macro1",
            &[
                json!({"command": "slide.new", "params": {}}),
                json!({"command": "slide.new", "params": {"layout": "Title Only", "at": 2, "title": "Q3 \"plan\"", "body": "a\n\tb"}}),
                json!({"command": "slide.new", "params": {"layout": "My custom"}}),
                json!({"command": "slide.delete", "params": {"index": 2}}),
                json!({"command": "slide.delete", "params": {"ids": [257, "x", 258]}}),
                json!({"command": "slide.delete", "params": {}}),
                json!({"command": "slide.go", "params": {"index": 3}}),
                json!({"command": "slide.go", "params": {"id": 300}}),
                json!({"command": "slide.go", "params": {}}),
                json!({"command": "slide.previous"}),
                json!({"command": "slide.selectSlides", "params": {"indices": [1]}}),
                json!({"command": "slide.selectSlides", "params": {"ids": [256]}}),
                json!({"command": "slide.selectSlides", "params": {"ids": [256, 257]}}),
                json!({"command": "insert.textBox", "params": {}}),
                json!({"command": "insert.textBox", "params": {"rect": [10.25, 20, 0, 30]}}),
                json!({"command": "shape.insert", "params": {"preset": "star5", "rect": [1, 2, 3, 4]}}),
                json!({"command": "shape.insert", "params": {"preset": "line", "rect": [10, 20, 100, -50]}}),
                json!({"command": "shape.insert", "params": {"preset": "cloud", "rect": [1, 2, 3, 4]}}),
                json!({"command": "shape.insert", "params": {"preset": "rect", "rect": [1e300, 2, 3, 4]}}),
                json!({"command": "text.insert", "params": {"text": "a"}}),
                json!({"command": "text.insert", "params": {"text": "b\u{b}c"}}),
                json!({"command": "format.bold", "params": {}}),
                json!({"command": "format.underline", "params": {"style": "none"}}),
                json!({"command": "format.size", "params": {"size": 1e9}}),
                json!({"command": "format.color", "params": {"color": "accent1"}}),
                json!({"command": "text.set", "params": {"id": 261, "text": ""}}),
                json!({"command": "shape.fill", "params": {"gradient": {}}}),
                json!({"command": "shape.fill", "params": {"color": "#123456", "ids": [5, 6]}}),
                json!({"command": "shape.line", "params": {"none": true}}),
                json!({"command": "shape.move", "params": {"x": 72, "y": 36.5, "ids": []}}),
                json!({"command": "shape.resize", "params": {"h": 10, "lockAspect": true}}),
                json!({"command": "shape.resize", "params": {"w": 10, "h": 20}}),
                json!({"command": "edit.select", "params": {"ids": [5, 6], "add": true}}),
                json!({"command": "edit.select", "params": {"ids": []}}),
                json!({"command": "selection.select", "params": {"id": 7, "toggle": true}}),
                json!({"command": "edit.delete", "params": {"scope": "slides"}}),
                json!({"command": "edit.delete", "params": {}}),
                json!({"command": "shape.rename", "params": {"id": 5, "name": "A \"b\""}}),
                json!({"command": "evil\r\nEnd Sub", "params": {}}),
            ],
        );
        let expected = [
            "Sub Macro1()",
            "    ActiveWindow.View.GotoSlide ActivePresentation.Slides.Add(ActiveWindow.View.Slide.SlideIndex + 1, IIf(ActiveWindow.View.Slide.Layout = ppLayoutTitle, ppLayoutText, ActiveWindow.View.Slide.Layout)).SlideIndex",
            "    ActiveWindow.View.GotoSlide ActivePresentation.Slides.Add(3, ppLayoutTitleOnly).SlideIndex",
            "    ActiveWindow.View.Slide.Shapes(1).TextFrame.TextRange.Text = \"Q3 \"\"plan\"\"\"",
            "    ActiveWindow.View.Slide.Shapes(2).TextFrame.TextRange.Text = \"a\" & vbCr & vbTab & \"b\"",
            "    ' (step not recorded: slide.new)",
            "    ActivePresentation.Slides(3).Delete",
            "    ActivePresentation.Slides.FindBySlideID(257).Delete",
            "    ActivePresentation.Slides.FindBySlideID(258).Delete",
            "    ActiveWindow.Selection.SlideRange.Delete",
            "    ActiveWindow.View.GotoSlide 4",
            "    ActiveWindow.View.GotoSlide ActivePresentation.Slides.FindBySlideID(300).SlideIndex",
            "    ' (step not recorded: slide.go)",
            "    If ActiveWindow.View.Slide.SlideIndex > 1 Then ActiveWindow.View.GotoSlide ActiveWindow.View.Slide.SlideIndex - 1",
            "    ActivePresentation.Slides(2).Select",
            "    ActivePresentation.Slides.FindBySlideID(256).Select",
            "    ' (step not recorded: slide.selectSlides)",
            "    ActiveWindow.View.Slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 372, 252, 216, 36).Select",
            "    ActiveWindow.View.Slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10.25, 20, 100, 30).Select",
            "    ActiveWindow.View.Slide.Shapes.AddShape(msoShape5pointStar, 1, 2, 3, 4).Select",
            "    ActiveWindow.View.Slide.Shapes.AddLine(10, 20, 110, 70).Select",
            "    ' (step not recorded: shape.insert cloud)",
            "    ' (step not recorded: shape.insert rect)",
            "    ActiveWindow.Selection.TextRange.InsertAfter \"ab\" & ChrW(11) & \"c\"",
            "    ActiveWindow.Selection.TextRange.Font.Bold = msoTriStateToggle",
            "    ActiveWindow.Selection.TextRange.Font.Underline = msoFalse",
            "    ' (step not recorded: format.size)",
            "    ' (step not recorded: format.color)",
            "    ActiveWindow.View.Slide.Shapes.FindById(261).TextFrame.TextRange.Text = \"\"",
            "    ' (step not recorded: shape.fill)",
            "    ActiveWindow.View.Slide.Shapes.FindById(5).Fill.ForeColor.RGB = RGB(18, 52, 86)",
            "    ActiveWindow.View.Slide.Shapes.FindById(6).Fill.ForeColor.RGB = RGB(18, 52, 86)",
            "    ActiveWindow.Selection.ShapeRange.Line.Visible = msoFalse",
            "    ' (step not recorded: shape.move)",
            "    ActiveWindow.Selection.ShapeRange.Width = ActiveWindow.Selection.ShapeRange.Width * 10 / ActiveWindow.Selection.ShapeRange.Height",
            "    ActiveWindow.Selection.ShapeRange.Height = 10",
            "    ActiveWindow.Selection.ShapeRange.Width = 10",
            "    ActiveWindow.Selection.ShapeRange.Height = 20",
            "    ActiveWindow.View.Slide.Shapes.FindById(5).Select msoFalse",
            "    ActiveWindow.View.Slide.Shapes.FindById(6).Select msoFalse",
            "    ActiveWindow.Selection.Unselect",
            "    ' (step not recorded: selection.select)",
            "    ActiveWindow.Selection.SlideRange.Delete",
            "    ActiveWindow.Selection.ShapeRange.Delete",
            "    ActiveWindow.View.Slide.Shapes.FindById(5).Name = \"A \"\"b\"\"\"",
            "    ' (step not recorded: evilEnd Sub)",
            "End Sub",
            "",
        ]
        .join("\r\n");
        assert_eq!(src, expected);
        assert!(crate::vba::parser::parse_module("M", &src).is_ok());
        assert_eq!(super::pp_text_literal("\n"), "vbCr");
        assert_eq!(super::pp_text_literal(""), "\"\"");
        // Long typing is cut into literals VBA's editor takes; the output stays bounded.
        let long = super::record_powerpoint("M", &[json!({"command": "text.insert", "params": {"text": "x".repeat(450)}})]);
        assert_eq!(long.matches(".InsertAfter").count(), 3);
        let many: Vec<_> = (0..20_000).map(|_| json!({"command": "slide.next"})).collect();
        assert!(super::record_powerpoint("M", &many).lines().count() < 10_010);
        let wide = json!({"command": "shape.setBounds", "params": {"x": 1, "y": 1, "w": 1, "h": 1, "ids": (0..50_000).collect::<Vec<_>>()}});
        assert!(super::record_powerpoint("M", &[wide.clone(), wide]).lines().count() < 10_010);
    }
}
