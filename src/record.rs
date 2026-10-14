//! `macros/record`: GridCraft's (or WordCraft's) commands, as the app recorded them, written as a
//! VBA Sub.

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
}
