//! The add-in end to end: the `craft-vba` binary, driven over its standard input and output by a
//! test playing the app (which answers the add-in's `host/…` requests from a workbook in memory).

#[path = "../src/fakehost.rs"]
mod fakehost;

use std::io::{BufRead, BufReader, Write};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};

use fakehost::Book;
use serde_json::{Value as Json, json};

struct App {
    child: Child,
    stdin: ChildStdin,
    stdout: BufReader<ChildStdout>,
    book: Book,
    next: u64,
}

impl App {
    fn start() -> App {
        let mut child = Command::new(env!("CARGO_BIN_EXE_craft-vba")).stdin(Stdio::piped()).stdout(Stdio::piped()).spawn().unwrap();
        let stdin = child.stdin.take().unwrap();
        let stdout = BufReader::new(child.stdout.take().unwrap());
        App { child, stdin, stdout, book: Book::default(), next: 1 }
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
                let reply = match self.book.handle(m, &msg["params"]) {
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
