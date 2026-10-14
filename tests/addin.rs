//! The add-in end to end: the `craft-vba` binary, driven over its standard input and output by a
//! test playing the app (which answers the add-in's `host/…` requests from a workbook in memory).

#[path = "../src/fakedeck.rs"]
mod fakedeck;
#[path = "../src/fakehost.rs"]
mod fakehost;
#[path = "../src/fakeword.rs"]
mod fakeword;

use std::io::{BufRead, BufReader, Write};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};

use fakedeck::Deck;
use fakehost::Book;
use fakeword::Doc;
use serde_json::{Value as Json, json};

struct App {
    child: Child,
    stdin: ChildStdin,
    stdout: BufReader<ChildStdout>,
    book: Book,
    /// Answers the host requests instead of the workbook when the app is DeckCraft.
    deck: Option<Deck>,
    /// Answers the host requests instead of the workbook when the app is WordCraft.
    doc: Option<Doc>,
    next: u64,
}

impl App {
    fn start() -> App {
        let mut child = Command::new(env!("CARGO_BIN_EXE_craft-vba")).stdin(Stdio::piped()).stdout(Stdio::piped()).spawn().unwrap();
        let stdin = child.stdin.take().unwrap();
        let stdout = BufReader::new(child.stdout.take().unwrap());
        App { child, stdin, stdout, book: Book::default(), deck: None, doc: None, next: 1 }
    }

    /// Sends a request and answers the add-in's host requests until its reply comes.
    fn request(&mut self, method: &str, params: Json) -> Json {
        let id = self.next;
        self.next += 1;
        writeln!(self.stdin, "{}", json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params})).unwrap();
        loop {
            let mut line = String::new();
            assert!(self.stdout.read_line(&mut line).unwrap() > 0, "the add-in exited");
            let msg: Json = serde_json::from_str(&line).unwrap();
            if let Some(m) = msg.get("method").and_then(Json::as_str) {
                let answer = match (&mut self.deck, &mut self.doc) {
                    (Some(d), _) => d.handle(m, &msg["params"]),
                    (None, Some(d)) => d.handle(m, &msg["params"]),
                    (None, None) => self.book.handle(m, &msg["params"]),
                };
                let reply = match answer {
                    Ok(result) => json!({"jsonrpc": "2.0", "id": msg["id"], "result": result}),
                    Err((code, message, data)) => json!({"jsonrpc": "2.0", "id": msg["id"], "error": {"code": code, "message": message, "data": data}}),
                };
                writeln!(self.stdin, "{reply}").unwrap();
                continue;
            }
            assert_eq!(msg["id"], json!(id));
            return msg;
        }
    }
}

#[test]
fn macros_are_written_listed_run_recorded_and_removed() {
    let manifest: Json = serde_json::from_slice(&Command::new(env!("CARGO_BIN_EXE_craft-vba")).arg("--manifest").output().unwrap().stdout).unwrap();
    assert_eq!(manifest["id"], "org.craft.vba");
    assert_eq!(manifest["macros"]["language"], "VBA");

    let mut app = App::start();
    assert_eq!(app.request("initialize", json!({"protocolVersion": 1, "host": {"id": "gridcraft"}}))["result"]["protocolVersion"], 1);
    assert_eq!(app.request("macros/list", json!({}))["result"]["macros"], json!([]), "no project yet");

    let source = "Sub Hello()\r\n    Range(\"A1\").Value = \"hi\"\r\n    Range(\"A2\").Value = Twice(21)\r\n    MsgBox \"done\"\r\nEnd Sub\r\n\
                  Function Twice(x)\r\n    Twice = x * 2\r\nEnd Function\r\nSub Fail()\r\n    Dim x\r\n    x = 1 / 0\r\nEnd Sub\r\nPrivate Sub Hidden()\r\nEnd Sub\r\n";
    let saved = app.request("macros/saveModule", json!({"name": "Module1", "source": source}));
    assert_eq!(saved["result"], json!({}), "{saved}");
    assert!(app.book.vba.is_some(), "the project was written to the workbook");

    let list = app.request("macros/list", json!({}));
    let names: Vec<&str> = list["result"]["macros"].as_array().unwrap().iter().map(|m| m["name"].as_str().unwrap()).collect();
    assert_eq!(names, ["Module1.Hello", "Module1.Fail"]);

    assert_eq!(app.request("macros/run", json!({"name": "Hello"}))["result"], json!({}));
    assert_eq!(app.book.get("Sheet1", "A1"), json!("hi"));
    assert_eq!(app.book.get("Sheet1", "A2"), json!(42));
    assert_eq!(app.book.messages, vec![("done".to_string(), None)]);

    let failed = app.request("macros/run", json!({"name": "Module1.Fail"}));
    assert_eq!(failed["error"]["message"], "Run-time error '11': Division by zero");
    // Lines count the stored source, its `Attribute VB_Name` line included (as macros/modules returns it).
    assert_eq!(failed["error"]["data"]["line"], 12);
    assert_eq!(failed["error"]["data"]["module"], "Module1");
    assert!(app.request("macros/run", json!({"name": "Nope"}))["error"]["message"].as_str().unwrap().contains("Cannot run the macro"));

    // A ribbon button's macro gets the button; one without parameters runs as it is.
    let callback = "Sub Pressed(control As IRibbonControl)\r\n    Range(\"C1\").Value = control.Id & \"/\" & control.Tag\r\nEnd Sub\r\n";
    assert_eq!(app.request("macros/saveModule", json!({"name": "Ribbon", "source": callback}))["result"], json!({}));
    let control = json!({"id": "craftButton1", "tag": "north"});
    assert_eq!(app.request("macros/run", json!({"name": "Ribbon.Pressed", "control": control}))["result"], json!({}));
    assert_eq!(app.book.get("Sheet1", "C1"), json!("craftButton1/north"));
    assert_eq!(app.request("macros/run", json!({"name": "Hello", "control": control}))["result"], json!({}));
    let names: Vec<String> =
        app.request("macros/list", json!({}))["result"]["macros"].as_array().unwrap().iter().map(|m| m["name"].as_str().unwrap().to_string()).collect();
    assert!(!names.contains(&"Ribbon.Pressed".to_string()), "a callback isn't a macro to run from the Macros list");
    assert_eq!(app.request("macros/deleteModule", json!({"name": "Ribbon"}))["result"], json!({}));

    let modules = app.request("macros/modules", json!({}));
    let m = &modules["result"]["modules"][0];
    assert_eq!((m["name"].as_str(), m["kind"].as_str()), (Some("Module1"), Some("standard")));
    assert!(m["source"].as_str().unwrap().starts_with("Attribute VB_Name = \"Module1\"\r\nSub Hello()"));

    // A source with a syntax error is saved and reported.
    let bad = app.request("macros/saveModule", json!({"name": "Draft", "source": "Sub A(\r\nEnd Sub"}));
    assert_eq!(bad["result"]["compileError"]["line"], 1);
    let recorded = app.request("macros/record", json!({"steps": [{"command": "cell.set", "params": {"cell": "B2", "input": "5"}}]}));
    assert_eq!(recorded["result"]["source"], "Sub Macro1()\r\n    Range(\"B2\").Value = 5\r\nEnd Sub\r\n");
    assert_eq!(app.request("macros/deleteModule", json!({"name": "Draft"}))["result"], json!({}));
    assert!(app.request("macros/deleteModule", json!({"name": "Draft"}))["error"].is_object());
    assert_eq!(app.request("macros/frobnicate", json!({}))["error"]["code"], -32601);

    writeln!(app.stdin, "{}", json!({"jsonrpc": "2.0", "method": "shutdown"})).unwrap();
    let status = app.child.wait().unwrap();
    assert!(status.success());
}

#[test]
fn powerpoint_macros_run_and_record_in_deckcraft() {
    let manifest: Json = serde_json::from_slice(&Command::new(env!("CARGO_BIN_EXE_craft-vba")).arg("--manifest").output().unwrap().stdout).unwrap();
    assert_eq!(manifest["apps"], json!(["gridcraft", "deckcraft", "wordcraft"]));

    let mut app = App::start();
    app.deck = Some(Deck::default());
    assert_eq!(app.request("initialize", json!({"protocolVersion": 1, "host": {"id": "deckcraft"}}))["result"]["protocolVersion"], 1);
    let source = "Sub Build()\r\n\
                  \x20   Dim s As Slide, sh As Shape\r\n\
                  \x20   Set s = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)\r\n\
                  \x20   Set sh = s.Shapes.AddShape(msoShapeOval, 10, 20, 100, 50)\r\n\
                  \x20   sh.Fill.ForeColor.RGB = RGB(255, 0, 0)\r\n\
                  \x20   sh.TextFrame.TextRange.Text = \"Slide \" & s.SlideIndex\r\n\
                  \x20   MsgBox ActivePresentation.Slides.Count & \" slides\"\r\n\
                  End Sub\r\n\
                  Sub Bad()\r\n    ActivePresentation.Slides(99).Delete\r\nEnd Sub\r\n\
                  Sub Pressed(control As IRibbonControl)\r\n    ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text = control.Id & \"/\" & control.Tag\r\nEnd Sub\r\n";
    assert_eq!(app.request("macros/saveModule", json!({"name": "Module1", "source": source}))["result"], json!({}));
    let names: Vec<String> =
        app.request("macros/list", json!({}))["result"]["macros"].as_array().unwrap().iter().map(|m| m["name"].as_str().unwrap().to_string()).collect();
    assert_eq!(names, ["Module1.Build", "Module1.Bad"]);

    assert_eq!(app.request("macros/run", json!({"name": "Build"}))["result"], json!({}));
    let deck = app.deck.as_ref().unwrap();
    assert_eq!(deck.slides.len(), 3);
    let oval = &deck.slides[2].shapes[0];
    assert_eq!((oval.geometry.as_str(), oval.fill.as_deref(), oval.text.as_str()), ("ellipse", Some("#FF0000"), "Slide 3"));
    assert_eq!((oval.left, oval.top, oval.width, oval.height), (10.0, 20.0, 100.0, 50.0));
    assert_eq!(deck.messages, vec![("3 slides".to_string(), None)]);

    let failed = app.request("macros/run", json!({"name": "Bad"}));
    assert_eq!(failed["error"]["message"], "Run-time error '9': Subscript out of range");
    assert_eq!(failed["error"]["data"]["number"], 9);

    let control = json!({"id": "deckButton", "tag": "east"});
    assert_eq!(app.request("macros/run", json!({"name": "Module1.Pressed", "control": control}))["result"], json!({}));
    assert_eq!(app.deck.as_ref().unwrap().shape(0, "Title 1").text, "deckButton/east");

    // The recorder writes DeckCraft's steps as PowerPoint VBA (a craft extra), and what it writes runs.
    let steps = json!([
        {"command": "slide.go", "params": {"index": 0}},
        {"command": "shape.insert", "params": {"preset": "roundRect", "rect": [5, 6, 70, 80], "text": "Go"}},
        {"command": "format.bold", "params": {"on": true}},
        {"command": "shape.fill", "params": {"color": "#0000FF"}},
        {"command": "transition.apply", "params": {}}
    ]);
    let recorded = app.request("macros/record", json!({"steps": steps, "name": "Drawn"}));
    let src = recorded["result"]["source"].as_str().unwrap().to_string();
    assert_eq!(
        src,
        "Sub Drawn()\r\n    ActiveWindow.View.GotoSlide 1\r\n    \
         ActiveWindow.View.Slide.Shapes.AddShape(msoShapeRoundedRectangle, 5, 6, 70, 80).Select\r\n    \
         ActiveWindow.Selection.TextRange.Text = \"Go\"\r\n    ActiveWindow.Selection.TextRange.Font.Bold = msoTrue\r\n    \
         ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB = RGB(0, 0, 255)\r\n    ' (step not recorded: transition.apply)\r\nEnd Sub\r\n"
    );
    assert_eq!(app.request("macros/saveModule", json!({"name": "Recorded", "source": src}))["result"], json!({}));
    assert_eq!(app.request("macros/run", json!({"name": "Drawn"}))["result"], json!({}));
    let drawn = app.deck.as_ref().unwrap().slides[0].shapes.last().cloned().unwrap();
    assert_eq!((drawn.geometry.as_str(), drawn.text.as_str(), drawn.bold, drawn.fill.as_deref()), ("roundRectangle", "Go", true, Some("#0000FF")));

    writeln!(app.stdin, "{}", json!({"jsonrpc": "2.0", "method": "shutdown"})).unwrap();
    assert!(app.child.wait().unwrap().success());
}

#[test]
fn word_macros_run_and_record_in_wordcraft() {
    let mut app = App::start();
    app.doc = Some(Doc::default());
    assert_eq!(app.request("initialize", json!({"protocolVersion": 1, "host": {"id": "wordcraft"}}))["result"]["protocolVersion"], 1);
    assert_eq!(app.request("macros/list", json!({}))["result"]["macros"], json!([]), "no project yet");
    let run = app.request("macros/run", json!({"name": "Build"}));
    assert_eq!(run["error"]["message"], "Cannot run the macro 'Build'. This document has no macros.");
    let source = "Sub Build()\r\n\
                  \x20   Dim p As Paragraph\r\n\
                  \x20   For Each p In ActiveDocument.Paragraphs\r\n\
                  \x20       p.SpaceAfter = 6\r\n\
                  \x20   Next\r\n\
                  \x20   With ActiveDocument.Paragraphs.Add\r\n\
                  \x20       .Range.Text = \"Added\"\r\n\
                  \x20       .Style = wdStyleHeading2\r\n\
                  \x20       .Alignment = wdAlignParagraphCenter\r\n\
                  \x20       .Range.Font.Color = RGB(255, 0, 0)\r\n\
                  \x20   End With\r\n\
                  \x20   ActiveDocument.BuiltInDocumentProperties(\"Title\") = \"Built\"\r\n\
                  \x20   MsgBox ActiveDocument.Paragraphs.Count & \" paragraphs\"\r\n\
                  End Sub\r\n\
                  Sub Bad()\r\n    ActiveDocument.Paragraphs(99).Range.Delete\r\nEnd Sub\r\n\
                  Sub Pressed(control As IRibbonControl)\r\n    ActiveDocument.Paragraphs(1).Range.Text = control.Id & \"/\" & control.Tag\r\nEnd Sub\r\n";
    assert_eq!(app.request("macros/saveModule", json!({"name": "Module1", "source": source}))["result"], json!({}));
    let names: Vec<String> =
        app.request("macros/list", json!({}))["result"]["macros"].as_array().unwrap().iter().map(|m| m["name"].as_str().unwrap().to_string()).collect();
    assert_eq!(names, ["Module1.Build", "Module1.Bad"]);

    assert_eq!(app.request("macros/run", json!({"name": "Build"}))["result"], json!({}));
    let doc = app.doc.as_ref().unwrap();
    assert_eq!(doc.texts(), ["Hello world", "Second paragraph.", "Third.", "Added"]);
    let added = &doc.paras[3];
    assert_eq!((added.style.as_str(), added.alignment.as_str(), added.chars[0].1.color.as_str()), ("Heading 2", "Centered", "#FF0000"));
    assert!(doc.paras[..3].iter().all(|p| p.space_after == 6.0));
    assert_eq!(doc.title, "Built");
    assert_eq!(doc.messages, vec![("4 paragraphs".to_string(), None)]);

    let failed = app.request("macros/run", json!({"name": "Bad"}));
    assert_eq!(failed["error"]["message"], "Run-time error '9': Subscript out of range");
    assert_eq!(failed["error"]["data"]["number"], 9);

    let control = json!({"id": "wordButton", "tag": "west"});
    assert_eq!(app.request("macros/run", json!({"name": "Module1.Pressed", "control": control}))["result"], json!({}));
    assert_eq!(app.doc.as_ref().unwrap().texts()[0], "wordButton/west");
    assert_eq!(app.request("macros/run", json!({"name": "Build", "control": control}))["result"], json!({}), "a macro without parameters runs as it is");

    // The recorder writes WordCraft's steps as Word VBA, and what it writes runs.
    let steps = json!([
        {"command": "caret.docEnd", "params": {}},
        {"command": "text.newParagraph", "params": {}},
        {"command": "format.bold", "params": {"value": true}},
        {"command": "text.insert", "params": {"text": "Recorded"}},
        {"command": "para.alignRight", "params": {}},
        {"command": "para.style", "params": {"style": "Quote"}},
        {"command": "caret.left", "params": {}}
    ]);
    let recorded = app.request("macros/record", json!({"steps": steps, "name": "Typed"}));
    let src = recorded["result"]["source"].as_str().unwrap().to_string();
    assert_eq!(
        src,
        "Sub Typed()\r\n    Selection.EndKey Unit:=wdStory\r\n    Selection.TypeParagraph\r\n    Selection.Font.Bold = True\r\n    \
         Selection.TypeText Text:=\"Recorded\"\r\n    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight\r\n    \
         Selection.Style = ActiveDocument.Styles(\"Quote\")\r\n    ' (step not recorded: caret.left)\r\nEnd Sub\r\n"
    );
    assert_eq!(app.request("macros/saveModule", json!({"name": "Recorded", "source": src}))["result"], json!({}));
    assert_eq!(app.request("macros/run", json!({"name": "Typed"}))["result"], json!({}));
    let doc = app.doc.as_ref().unwrap();
    let last = doc.paras.last().unwrap();
    assert_eq!((last.text().as_str(), last.alignment.as_str(), last.style.as_str()), ("Recorded", "Right", "Quote"));
    assert!(last.chars.iter().all(|(_, f)| f.bold), "typed after Bold was set at the caret");

    writeln!(app.stdin, "{}", json!({"jsonrpc": "2.0", "method": "shutdown"})).unwrap();
    assert!(app.child.wait().unwrap().success());
}
