//! The host application's objects, as macros see them: Excel's in GridCraft, PowerPoint's in
//! DeckCraft. The interpreter knows only this seam: it hands every host object ([`Obj::Xl`],
//! [`Obj::Pp`]) and every name it doesn't know itself to the [`Model`] chosen at `initialize`
//! (from the app's `host.id`).

use std::rc::Rc;

use crate::excel::{Excel, XlObj};
use crate::host::Host;
use crate::powerpoint::{PowerPoint, PpObj};
use crate::vba::value::{Obj, R, Value, VbaError};

/// Which object model an app gets.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum App {
    /// Excel's (GridCraft, and any app that doesn't say otherwise).
    #[default]
    Excel,
    /// PowerPoint's (DeckCraft).
    PowerPoint,
}

impl App {
    /// The model for `initialize`'s `host.id`.
    pub fn for_host(id: &str) -> App {
        if id.trim().eq_ignore_ascii_case("deckcraft") { App::PowerPoint } else { App::Excel }
    }

    /// What the app calls the document macros live in.
    pub fn document(self) -> &'static str {
        match self {
            App::Excel => "workbook",
            App::PowerPoint => "presentation",
        }
    }
}

pub enum Model {
    Excel(Excel),
    PowerPoint(PowerPoint),
}

impl Default for Model {
    fn default() -> Model {
        Model::Excel(Excel::default())
    }
}

type Args<'a> = &'a [Value];

/// An object of the other app's model (an Excel range in DeckCraft): it has no members here.
fn foreign() -> VbaError {
    VbaError::with(438, "Object doesn't support this property or method (it belongs to another application)")
}

impl Model {
    pub fn new(app: App) -> Model {
        match app {
            App::Excel => Model::Excel(Excel::default()),
            App::PowerPoint => Model::PowerPoint(PowerPoint::default()),
        }
    }

    pub fn app(&self) -> App {
        match self {
            Model::Excel(_) => App::Excel,
            Model::PowerPoint(_) => App::PowerPoint,
        }
    }

    /// Requests sent to the host (for tests and tuning).
    pub fn requests(&self) -> usize {
        match self {
            Model::Excel(x) => x.requests,
            Model::PowerPoint(p) => p.requests,
        }
    }

    /// Sends the queued writes.
    pub fn flush(&mut self, host: &mut dyn Host) -> R<()> {
        match self {
            Model::Excel(x) => x.flush(host),
            Model::PowerPoint(p) => p.flush(host),
        }
    }

    /// The app's constants (`ppLayoutText`, `msoTrue`); the language's own come first.
    pub fn constant(&self, name: &str) -> Option<Value> {
        match self {
            Model::Excel(_) => None,
            Model::PowerPoint(_) => crate::powerpoint::constant(name),
        }
    }

    /// The names a macro reaches without an object (`Application`, `ActiveSheet`, `Range`,
    /// `ActivePresentation`).
    pub fn global(&mut self, host: &mut dyn Host, name: &str, args: Args) -> R<Option<Value>> {
        let n = name.to_ascii_lowercase();
        match self {
            Model::Excel(x) => {
                let o = |o: XlObj| Ok(Some(Value::Object(Obj::Xl(o))));
                match n.as_str() {
                    "application" => o(XlObj::Application),
                    "worksheetfunction" => o(XlObj::WorksheetFunction),
                    "thisworkbook" | "activeworkbook" => o(XlObj::Workbook),
                    "activesheet" | "activecell" | "selection" | "worksheets" | "sheets" | "range" | "cells" | "rows" | "columns" => {
                        x.member(host, &XlObj::Application, name, args, &[]).map(Some)
                    }
                    _ => Ok(None),
                }
            }
            Model::PowerPoint(p) => p.global(host, name, args),
        }
    }

    pub fn member(&mut self, host: &mut dyn Host, o: &Obj, name: &str, args: Args, named: &[(String, Value)]) -> R<Value> {
        match (self, o) {
            (Model::Excel(x), Obj::Xl(o)) => x.member(host, o, name, args, named),
            (Model::PowerPoint(p), Obj::Pp(o)) => p.member(host, o, name, args, named),
            _ => Err(foreign()),
        }
    }

    pub fn assign(&mut self, host: &mut dyn Host, o: &Obj, name: &str, args: Args, value: &Value) -> R<()> {
        match (self, o) {
            (Model::Excel(x), Obj::Xl(o)) => x.assign(host, o, name, args, value),
            (Model::PowerPoint(p), Obj::Pp(o)) if args.is_empty() => p.assign(host, o, name, value),
            (Model::PowerPoint(_), Obj::Pp(_)) => Err(VbaError::new(450)),
            _ => Err(foreign()),
        }
    }

    /// `object = value` through the object's default property (`rng = 5`).
    pub fn assign_default(&mut self, host: &mut dyn Host, o: &Obj, value: &Value) -> R<()> {
        match (self, o) {
            (Model::Excel(x), Obj::Xl(o)) => x.assign(host, o, "value", &[], value),
            (Model::PowerPoint(p), Obj::Pp(o)) => p.assign_default(host, o, value),
            _ => Err(foreign()),
        }
    }

    /// `Name(args) = value` with a global name (`Range("A1") = 5`, `Cells(1, 1) = 5`).
    pub fn assign_global(&mut self, host: &mut dyn Host, name: &str, args: Args, value: &Value) -> R<()> {
        match self {
            Model::Excel(x) => x.assign(host, &XlObj::Application, name, args, value),
            Model::PowerPoint(_) => Err(VbaError::with(438, format!("{name}(…) can't be assigned to"))),
        }
    }

    pub fn default_call(&mut self, host: &mut dyn Host, o: &Obj, args: Args) -> R<Value> {
        match (self, o) {
            (Model::Excel(x), Obj::Xl(o)) => x.default_call(host, o, args),
            (Model::PowerPoint(p), Obj::Pp(o)) => p.default_call(host, o, args),
            _ => Err(foreign()),
        }
    }

    pub fn default_value(&mut self, host: &mut dyn Host, o: &Obj) -> R<Value> {
        match (self, o) {
            (Model::Excel(x), Obj::Xl(o)) => x.default_value(host, o),
            (Model::PowerPoint(p), Obj::Pp(o)) => p.default_value(host, o),
            _ => Err(foreign()),
        }
    }

    /// What `For Each` walks.
    pub fn items(&mut self, host: &mut dyn Host, o: &Obj) -> R<Vec<Value>> {
        match (self, o) {
            (Model::Excel(x), Obj::Xl(o)) => x.items(host, o),
            (Model::PowerPoint(p), Obj::Pp(o)) => p.items(host, o),
            _ => Err(foreign()),
        }
    }

    /// An object passed where a value is wanted becomes its value (a range, a text range).
    pub fn is_value(o: &Obj) -> bool {
        matches!(o, Obj::Xl(XlObj::Range(_)) | Obj::Pp(PpObj::TextRange(_)))
    }

    /// The `Application` object (`Application.Run`).
    pub fn is_application(o: &Obj) -> bool {
        matches!(o, Obj::Xl(XlObj::Application) | Obj::Pp(PpObj::Application))
    }

    /// The ribbon button a macro was called from, as an `IRibbonControl`.
    pub fn ribbon_control(&self, id: Rc<str>, tag: Rc<str>) -> Value {
        Value::Object(match self {
            Model::Excel(_) => Obj::Xl(XlObj::RibbonControl { id, tag }),
            Model::PowerPoint(_) => Obj::Pp(PpObj::RibbonControl { id, tag }),
        })
    }
}
