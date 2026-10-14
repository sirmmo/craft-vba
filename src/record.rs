//! `macros/record`: GridCraft commands, as the app recorded them, written as a VBA Sub.

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
}
