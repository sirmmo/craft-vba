//! The add-in: JSON-RPC 2.0 over standard input and output, one message per line (the craft
//! plugin protocol 1, with 1.1's `macros/*` requests and `host/*` calls).
//!
//! While the app waits for the answer to one of its requests, the add-in sends it `host/…`
//! requests of its own and reads their answers before going on: one thread, one stream, in order.

use std::io::{BufRead, Write};

use serde_json::{Value as Json, json};

use crate::host::{Host, HostError};
use crate::model::App;
use crate::ovba::{ModuleKind, Project};
use crate::vba::ast::ProcKind;
use crate::vba::{Interp, Program};

/// Longest line read (a large project in base64 fits).
const MAX_LINE: usize = 256 << 20;

/// The manifest `--manifest` prints.
pub fn manifest() -> Json {
    json!({
        "id": "org.craft.vba",
        "name": "VBA",
        "version": env!("CARGO_PKG_VERSION"),
        "description": "Runs the document's VBA macros (Excel's objects in GridCraft, PowerPoint's in DeckCraft, Word's in WordCraft), and edits and records them (the Developer tab).",
        "author": "Marco Montanari and the craft-vba contributors",
        "apps": ["gridcraft", "deckcraft", "wordcraft"],
        "run": {"command": "./craft-vba"},
        "macros": {"language": "VBA"},
        "commands": []
    })
}

/// The standard input and output, as the [`Host`] macros call.
pub struct Stdio<R: BufRead, W: Write> {
    input: R,
    output: W,
    next: u64,
    /// The app closed the stream.
    pub closed: bool,
}

impl<R: BufRead, W: Write> Stdio<R, W> {
    pub fn new(input: R, output: W) -> Self {
        Stdio { input, output, next: 1, closed: false }
    }

    fn send(&mut self, msg: &Json) -> bool {
        writeln!(self.output, "{msg}").and_then(|_| self.output.flush()).is_ok()
    }

    /// The next message, `None` at the end of the stream.
    fn read(&mut self) -> Option<Json> {
        loop {
            let mut line = Vec::new();
            match std::io::Read::take(&mut self.input, MAX_LINE as u64).read_until(b'\n', &mut line) {
                Ok(0) | Err(_) => {
                    self.closed = true;
                    return None;
                }
                Ok(_) => {
                    if let Ok(v) = serde_json::from_slice::<Json>(&line) {
                        return Some(v);
                    }
                    let text = String::from_utf8_lossy(&line);
                    if !text.trim().is_empty() {
                        eprintln!("craft-vba: ignored a line that isn't JSON");
                    }
                }
            }
        }
    }
}

impl<R: BufRead, W: Write> Host for Stdio<R, W> {
    fn call(&mut self, method: &str, params: Json) -> Result<Json, HostError> {
        let id = format!("h{}", self.next);
        self.next += 1;
        let gone = || HostError { code: -32000, message: "the app closed the connection".into(), data: Json::Null };
        if !self.send(&json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params})) {
            return Err(gone());
        }
        loop {
            let msg = self.read().ok_or_else(gone)?;
            if msg.get("method").is_some() {
                // The app asked something while waiting for us: one thing at a time.
                if let Some(other) = msg.get("id").cloned() {
                    let busy = json!({"jsonrpc": "2.0", "id": other, "error": {"code": -32603, "message": "busy running a macro"}});
                    if !self.send(&busy) {
                        return Err(gone());
                    }
                } else if msg.get("method").and_then(Json::as_str) == Some("shutdown") {
                    self.closed = true;
                    return Err(gone());
                }
                continue;
            }
            if msg.get("id").and_then(Json::as_str) != Some(id.as_str()) {
                continue;
            }
            if let Some(e) = msg.get("error") {
                return Err(HostError {
                    code: e.get("code").and_then(Json::as_i64).unwrap_or(-32000),
                    message: e.get("message").and_then(Json::as_str).unwrap_or("the app refused").to_string(),
                    data: e.get("data").cloned().unwrap_or(Json::Null),
                });
            }
            return Ok(msg.get("result").cloned().unwrap_or(Json::Null));
        }
    }
}

/// A request's error: (code, message, data).
type Fail = (i64, String, Json);

fn fail(message: impl Into<String>) -> Fail {
    (1, message.into(), Json::Null)
}

fn project(app: App, host: &mut dyn Host) -> Result<Option<Project>, Fail> {
    let out = host.call("host/vbaProject/get", json!({})).map_err(|e| (e.code, e.message, e.data))?;
    match out.get("base64").and_then(Json::as_str) {
        Some(b) if !b.is_empty() => {
            let bytes = crate::base64::decode(b).ok_or_else(|| fail(format!("the {}'s VBA project isn't valid base64", app.document())))?;
            Project::read(&bytes).map(Some).map_err(fail)
        }
        _ => Ok(None),
    }
}

fn save(host: &mut dyn Host, p: &Project) -> Result<(), Fail> {
    let bytes = p.write().map_err(fail)?;
    host.call("host/vbaProject/set", json!({"base64": crate::base64::encode(&bytes)})).map(|_| ()).map_err(|e| (e.code, e.message, e.data))
}

fn program(p: &Project) -> Program {
    let sources: Vec<(String, String)> = p.modules.iter().map(|m| (m.name.clone(), m.source.clone())).collect();
    Program::new(&sources)
}

fn str_param<'a>(params: &'a Json, key: &str) -> Result<&'a str, Fail> {
    params.get(key).and_then(Json::as_str).filter(|s| !s.trim().is_empty()).ok_or_else(|| (-32602, format!("`{key}` (text) is required"), Json::Null))
}

/// Answers one request from the app. `initialize` sets `app` from its `host.id`: DeckCraft's
/// macros get PowerPoint's objects, WordCraft's Word's, every other app Excel's.
pub fn handle(app: &mut App, host: &mut dyn Host, method: &str, params: &Json) -> Result<Json, Fail> {
    let doc = app.document();
    match method {
        "initialize" => {
            *app = App::for_host(params.get("host").and_then(|h| h.get("id")).and_then(Json::as_str).unwrap_or(""));
            Ok(json!({"protocolVersion": 1, "capabilities": {"macros": {"language": "VBA"}}}))
        }
        "commands/run" => Err(fail("craft-vba has no commands of its own; run macros with macros/run")),
        "macros/list" => {
            let Some(p) = project(*app, host)? else { return Ok(json!({"macros": []})) };
            let prog = program(&p);
            let mut macros = Vec::new();
            for m in &prog.modules {
                let kind = p.module(&m.name).map(|x| x.kind).unwrap_or(ModuleKind::Standard);
                if kind == ModuleKind::Class {
                    continue;
                }
                for proc in &m.procs {
                    if proc.kind == ProcKind::Sub && proc.public && proc.params.iter().all(|q| q.optional) {
                        macros.push(json!({"name": format!("{}.{}", m.name, proc.name), "module": m.name, "procedure": proc.name}));
                    }
                }
            }
            Ok(json!({"macros": macros}))
        }
        "macros/modules" => {
            let Some(p) = project(*app, host)? else { return Ok(json!({"modules": []})) };
            Ok(json!({"modules": p.modules.iter().map(|m| json!({"name": m.name, "kind": m.kind.as_str(), "source": m.source})).collect::<Vec<_>>()}))
        }
        "macros/run" => {
            let name = str_param(params, "name")?;
            let p = project(*app, host)?.ok_or_else(|| fail(format!("Cannot run the macro '{name}'. This {doc} has no macros.")))?;
            let prog = program(&p);
            let mut it = Interp::for_app(&prog, host, *app);
            let (m, pr) = it.find(name).map_err(fail)?;
            // From a ribbon button: Excel passes the button (`Sub Go(control As IRibbonControl)`);
            // a macro without parameters runs as it is.
            let mut args = Vec::new();
            if let Some(c) = params.get("control").filter(|c| c.is_object())
                && it.param_count(m, pr) > 0
            {
                let text = |k: &str| -> std::rc::Rc<str> { c.get(k).and_then(Json::as_str).unwrap_or("").into() };
                args.push(it.model.ribbon_control(text("id"), text("tag")));
            }
            match it.run_with(m, pr, args) {
                Ok(()) => Ok(json!({})),
                Err(f) => Err((1, f.text(), json!({"module": f.module, "line": f.line, "number": f.number, "compile": f.compile}))),
            }
        }
        "macros/saveModule" => {
            let name = str_param(params, "name")?;
            let source = params.get("source").and_then(Json::as_str).ok_or_else(|| (-32602, "`source` (text) is required".to_string(), Json::Null))?;
            let mut p = project(*app, host)?.unwrap_or_else(Project::new);
            if p.module(name).is_some_and(|m| m.kind != ModuleKind::Standard) {
                // Document and class modules keep their kind; only their source changes.
                if let Some(m) = p.modules.iter_mut().find(|m| m.name.eq_ignore_ascii_case(name)) {
                    m.source = source.replace("\r\n", "\n").replace('\n', "\r\n");
                }
            } else {
                p.set_module(name, source).map_err(fail)?;
            }
            save(host, &p)?;
            // Saved either way, as the editor does; a compile error is reported alongside.
            let check = crate::vba::parser::parse_module(name, source);
            Ok(match check {
                Ok(_) => json!({}),
                Err((line, message)) => json!({"compileError": {"line": line, "message": format!("Compile error: {message}")}}),
            })
        }
        "macros/deleteModule" => {
            let name = str_param(params, "name")?;
            let mut p = project(*app, host)?.ok_or_else(|| fail(format!("This {doc} has no macros.")))?;
            if p.module(name).is_some_and(|m| m.kind == ModuleKind::Document) {
                return Err(fail(format!("{name} belongs to the {doc} and can't be removed")));
            }
            if !p.remove_module(name) {
                return Err(fail(format!("There's no module named {name}")));
            }
            save(host, &p)?;
            Ok(json!({}))
        }
        "macros/record" => {
            let steps = params.get("steps").and_then(Json::as_array).cloned().unwrap_or_default();
            let name = params
                .get("name")
                .and_then(Json::as_str)
                .filter(|n| n.chars().next().is_some_and(char::is_alphabetic) && n.chars().all(|c| c.is_alphanumeric() || c == '_'))
                .unwrap_or("Macro1");
            let source = match *app {
                App::Word => crate::record::record_word(name, &steps),
                App::PowerPoint => crate::record::record_powerpoint(name, &steps),
                App::Excel => crate::record::record(name, &steps),
            };
            Ok(json!({"source": source}))
        }
        other => Err((-32601, format!("unknown method {other}"), Json::Null)),
    }
}

/// Serves the app until it closes the stream or says `shutdown`.
pub fn serve<R: BufRead, W: Write>(input: R, output: W) {
    let mut io = Stdio::new(input, output);
    let mut app = App::default();
    while let Some(msg) = io.read() {
        let Some(method) = msg.get("method").and_then(Json::as_str).map(str::to_string) else { continue };
        let Some(id) = msg.get("id").cloned() else {
            if method == "shutdown" {
                return;
            }
            continue;
        };
        let params = msg.get("params").cloned().unwrap_or(Json::Null);
        // A bug here must cost the request, never the add-in.
        let answer = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| handle(&mut app, &mut io, &method, &params)))
            .unwrap_or_else(|_| Err(fail("craft-vba hit an internal error")));
        let reply = match answer {
            Ok(result) => json!({"jsonrpc": "2.0", "id": id, "result": result}),
            Err((code, message, data)) => {
                let mut e = json!({"code": code, "message": message});
                if let (false, Some(o)) = (data.is_null(), e.as_object_mut()) {
                    o.insert("data".into(), data);
                }
                json!({"jsonrpc": "2.0", "id": id, "error": e})
            }
        };
        if io.closed || !io.send(&reply) {
            return;
        }
    }
}
