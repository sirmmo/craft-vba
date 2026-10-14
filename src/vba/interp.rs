//! Runs VBA: procedures, variables (ByRef/ByVal), typed declarations, arrays, control flow, `With`,
//! `On Error`, the `Err` and `Debug` objects, `Collection`, built-in functions, and the host
//! application's objects (through its [`Model`] and the [`Host`]).
//!
//! Every statement costs a step: a macro stops after [`MAX_STEPS`] (an endless loop is an error,
//! never a hang), and calls nest at most [`MAX_DEPTH`] deep.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use serde_json::json;

use super::ast::*;
use super::builtins;
use super::ops;
use super::value::{Array, Obj, R, Value, VbaError};
use crate::host::Host;
use crate::model::{App, Model};

pub const MAX_STEPS: u64 = 20_000_000;
pub const MAX_DEPTH: usize = 200;

/// A parsed project: its modules (and the ones that didn't compile, with why).
pub struct Program {
    pub modules: Vec<ModuleAst>,
    pub broken: Vec<(String, u32, String)>,
}

/// Where a macro stopped with an error.
#[derive(Clone, Debug, PartialEq)]
pub struct Failure {
    pub number: i32,
    pub message: String,
    pub module: String,
    pub line: u32,
    /// A compile error (or the step limit): not a run-time error a handler can catch.
    pub compile: bool,
}

impl Failure {
    pub fn text(&self) -> String {
        if self.compile {
            format!("Compile error: {}", self.message)
        } else if self.number == 0 {
            self.message.clone()
        } else {
            format!("Run-time error '{}': {}", self.number, self.message)
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Ty {
    Variant,
    Integer,
    Long,
    Double,
    String,
    Boolean,
    Date,
    Object,
}

fn ty_of(name: Option<&str>) -> Ty {
    match name.map(str::to_ascii_lowercase).as_deref() {
        None | Some("variant") => Ty::Variant,
        Some("integer" | "byte") => Ty::Integer,
        Some("long" | "longlong" | "longptr") => Ty::Long,
        Some("double" | "single" | "currency" | "decimal") => Ty::Double,
        Some("string") => Ty::String,
        Some("boolean") => Ty::Boolean,
        Some("date") => Ty::Date,
        Some(_) => Ty::Object,
    }
}

fn default_of(t: Ty) -> Value {
    match t {
        Ty::Variant => Value::Empty,
        Ty::Integer => Value::Int(0),
        Ty::Long => Value::Long(0),
        Ty::Double => Value::Double(0.0),
        Ty::String => Value::str(""),
        Ty::Boolean => Value::Bool(false),
        Ty::Date => Value::Date(0.0),
        Ty::Object => Value::Nothing,
    }
}

fn coerce(t: Ty, v: Value) -> R<Value> {
    Ok(match (t, v) {
        (Ty::Variant, Value::Array(a)) => Value::Array(Rc::new(RefCell::new(a.borrow().clone()))),
        (Ty::Variant, v) => v,
        (Ty::Object, v @ (Value::Object(_) | Value::Nothing)) => v,
        (Ty::Object, _) => return Err(VbaError::new(13)),
        (_, Value::Array(_) | Value::Object(_) | Value::Nothing) => return Err(VbaError::new(13)),
        (Ty::Integer, v) => Value::Int(v.to_i16()?),
        (Ty::Long, v) => Value::Long(v.to_i32()?),
        (Ty::Double, v) => Value::Double(v.to_f64()?),
        (Ty::String, v) => Value::str(v.to_str()?),
        (Ty::Boolean, v) => Value::Bool(v.to_bool()?),
        (Ty::Date, v) => Value::Date(v.to_f64()?),
    })
}

/// A variable: its value and declared type (an array's element type).
struct Var {
    value: Value,
    ty: Ty,
    /// Declared as an array (`Dim a()`, `Dim a(1 To 3)`): it takes arrays whatever `ty` is.
    array: bool,
}

type Slot = Rc<RefCell<Var>>;

fn slot(value: Value, ty: Ty) -> Slot {
    Rc::new(RefCell::new(Var { value, ty, array: false }))
}

fn array_slot(value: Value, ty: Ty) -> Slot {
    Rc::new(RefCell::new(Var { value, ty, array: true }))
}

#[derive(Clone, Debug, PartialEq)]
enum OnErr {
    Off,
    ResumeNext,
    GoTo(String),
}

struct Frame {
    module: usize,
    locals: HashMap<String, Slot>,
    consts: HashMap<String, Value>,
    with: Vec<Value>,
    on_error: OnErr,
    /// Inside the error handler (a second error isn't caught by it).
    handling: bool,
    /// The top-level statement where the handled error happened (Resume Next continues after it).
    resume_at: usize,
    line: u32,
}

/// What a statement does to the flow.
enum Flow {
    Next,
    Exit(ExitKind),
    GoTo(String),
    ResumeNext,
    End,
}

/// An error on its way out: the error, and whether a handler may catch it.
struct Fail {
    err: VbaError,
    line: u32,
    module: usize,
    fatal: bool,
}

type X<T> = Result<T, Fail>;

/// An argument as passed: a variable (ByRef) or a value.
#[derive(Clone)]
enum ArgVal {
    Ref(Slot),
    Val(Value),
}

impl ArgVal {
    fn value(&self) -> Value {
        match self {
            ArgVal::Ref(s) => s.borrow().value.clone(),
            ArgVal::Val(v) => v.clone(),
        }
    }
}

pub struct Interp<'a> {
    prog: &'a Program,
    host: &'a mut dyn Host,
    /// The host application's objects (Excel's, PowerPoint's).
    pub model: Model,
    globals: Vec<HashMap<String, Slot>>,
    frames: Vec<Frame>,
    steps: u64,
    err: (i32, String, String),
    rnd: u64,
    const_depth: u32,
}

const COMPILE: i32 = -1;
/// `End`: the macro stops, without an error.
const END: i32 = -2;

fn compile_err(msg: impl Into<String>) -> VbaError {
    VbaError::with(COMPILE, msg)
}

impl<'a> Interp<'a> {
    /// An interpreter with Excel's objects.
    pub fn new(prog: &'a Program, host: &'a mut dyn Host) -> Interp<'a> {
        Interp::for_app(prog, host, App::Excel)
    }

    /// An interpreter with `app`'s objects.
    pub fn for_app(prog: &'a Program, host: &'a mut dyn Host, app: App) -> Interp<'a> {
        let mut globals = Vec::new();
        for m in &prog.modules {
            let mut vars = HashMap::new();
            for (d, _) in &m.vars {
                let ty = ty_of(d.ty.as_deref());
                let value = match &d.dims {
                    Some(_) => Value::Empty,
                    None => default_of(ty),
                };
                vars.insert(d.name.to_ascii_lowercase(), if d.dims.is_some() { array_slot(value, ty) } else { slot(value, ty) });
            }
            globals.push(vars);
        }
        let mut it = Interp {
            prog,
            host,
            model: Model::new(app),
            globals,
            frames: Vec::new(),
            steps: 0,
            err: (0, String::new(), String::new()),
            rnd: 0x2545_F491_4F6C_DD1D,
            const_depth: 0,
        };
        // Module arrays with bounds (Dim a(1 To 10) at module level).
        for (mi, m) in prog.modules.iter().enumerate() {
            for (d, _) in &m.vars {
                if let Some(dims) = &d.dims
                    && !dims.is_empty()
                    && let Ok(bounds) = it.const_bounds(mi, dims)
                    && let Ok(a) = Array::new(bounds)
                {
                    let ty = ty_of(d.ty.as_deref());
                    let mut a = a;
                    a.data.iter_mut().for_each(|x| *x = default_of(ty));
                    if let Some(s) = it.globals.get(mi).and_then(|g| g.get(&d.name.to_ascii_lowercase())) {
                        s.borrow_mut().value = Value::Array(Rc::new(RefCell::new(a)));
                    }
                }
            }
        }
        it
    }

    /// Finds a public parameterless Sub by `Module.Proc` or `Proc`.
    pub fn find(&self, name: &str) -> Result<(usize, usize), String> {
        let (module, proc) = match name.split_once('.') {
            Some((m, p)) => (Some(m), p),
            None => (None, name),
        };
        if let Some(m) = module
            && let Some((_, line, why)) = self.prog.broken.iter().find(|(n, _, _)| n.eq_ignore_ascii_case(m))
        {
            return Err(format!("Compile error in {m}, line {line}: {why}"));
        }
        let mut found = Vec::new();
        for (mi, m) in self.prog.modules.iter().enumerate() {
            if module.is_some_and(|x| !x.eq_ignore_ascii_case(&m.name)) {
                continue;
            }
            for (pi, p) in m.procs.iter().enumerate() {
                if p.name.eq_ignore_ascii_case(proc) && p.kind == ProcKind::Sub {
                    found.push((mi, pi));
                }
            }
        }
        match found.as_slice() {
            [one] => Ok(*one),
            [] => Err(format!("Cannot run the macro '{name}'. The macro may not be available in this workbook.")),
            _ => Err(format!("Ambiguous name detected: {proc}")),
        }
    }

    /// Runs a macro to its end. The queued writes are sent at the end, also after an error.
    pub fn run(&mut self, module: usize, proc: usize) -> Result<(), Failure> {
        self.run_with(module, proc, Vec::new())
    }

    /// How many parameters the procedure declares.
    pub fn param_count(&self, module: usize, proc: usize) -> usize {
        self.prog.modules.get(module).and_then(|m| m.procs.get(proc)).map_or(0, |p| p.params.len())
    }

    /// Runs a procedure with arguments (a ribbon button's macro gets the button).
    pub fn run_with(&mut self, module: usize, proc: usize, args: Vec<Value>) -> Result<(), Failure> {
        let r = self.call_proc(module, proc, args.into_iter().map(ArgVal::Val).collect(), Vec::new());
        let flushed = self.model.flush(self.host);
        let failure = |me: &Interp, f: Fail| Failure {
            number: if f.err.number == COMPILE { 0 } else { f.err.number },
            message: f.err.description,
            module: me.prog.modules.get(f.module).map(|m| m.name.clone()).unwrap_or_default(),
            line: f.line,
            compile: f.err.number == COMPILE,
        };
        match r {
            Ok(_) => flushed.map_err(|e| failure(self, Fail { err: e, line: 0, module, fatal: true })),
            Err(f) if f.err.number == END => flushed.map_err(|e| failure(self, Fail { err: e, line: 0, module, fatal: true })),
            Err(f) => Err(failure(self, f)),
        }
    }

    fn frame(&self) -> X<&Frame> {
        self.frames.last().ok_or_else(|| self.fail(VbaError::new(5)))
    }

    fn frame_mut(&mut self) -> Option<&mut Frame> {
        self.frames.last_mut()
    }

    fn fail(&self, err: VbaError) -> Fail {
        let (line, module) = self.frames.last().map(|f| (f.line, f.module)).unwrap_or((0, 0));
        let fatal = err.number == COMPILE;
        Fail { err, line, module, fatal }
    }

    fn x<T>(&self, r: R<T>) -> X<T> {
        r.map_err(|e| self.fail(e))
    }

    fn module(&self) -> Option<&'a ModuleAst> {
        let prog: &'a Program = self.prog;
        self.frames.last().and_then(|f| prog.modules.get(f.module))
    }

    fn compare_text(&self) -> bool {
        self.module().is_some_and(|m| m.compare_text)
    }

    // ------------------------------------------------------------ names

    fn lookup(&self, name: &str) -> Option<Slot> {
        let key = name.to_ascii_lowercase();
        let f = self.frames.last()?;
        if let Some(s) = f.locals.get(&key) {
            return Some(s.clone());
        }
        if let Some(s) = self.globals.get(f.module).and_then(|g| g.get(&key)) {
            return Some(s.clone());
        }
        // Public variables of the other modules.
        for (mi, m) in self.prog.modules.iter().enumerate() {
            if mi != f.module && m.vars.iter().any(|(d, public)| *public && d.name.eq_ignore_ascii_case(name)) {
                return self.globals.get(mi).and_then(|g| g.get(&key)).cloned();
            }
        }
        None
    }

    fn constant(&mut self, name: &str) -> X<Option<Value>> {
        let key = name.to_ascii_lowercase();
        if let Some(v) = self.frames.last().and_then(|f| f.consts.get(&key)) {
            return Ok(Some(v.clone()));
        }
        let here = self.frames.last().map(|f| f.module).unwrap_or(0);
        let prog: &'a Program = self.prog;
        let found =
            prog.modules.get(here).and_then(|m| m.consts.iter().find(|(n, _, _)| n.eq_ignore_ascii_case(name)).map(|(_, e, _)| (here, e))).or_else(|| {
                prog.modules
                    .iter()
                    .enumerate()
                    .find_map(|(mi, m)| m.consts.iter().find(|(n, _, public)| *public && n.eq_ignore_ascii_case(name)).map(|(_, e, _)| (mi, e)))
            });
        if let Some((mi, e)) = found {
            return self.const_value(mi, e).map(Some);
        }
        Ok(builtins::constant(name).or_else(|| self.model.constant(name)))
    }

    /// A module constant's value (constants are literals, other constants and operators).
    fn const_value(&mut self, module: usize, e: &Expr) -> X<Value> {
        self.const_depth += 1;
        if self.const_depth > 64 {
            self.const_depth = 0;
            return Err(self.fail(compile_err("Circular constant definition")));
        }
        let r = (|| -> X<Value> {
            Ok(match e {
                Expr::Num(n, int) => num_literal(*n, *int),
                Expr::Str(s) => Value::str(s.as_str()),
                Expr::Bool(b) => Value::Bool(*b),
                Expr::Nothing => Value::Nothing,
                Expr::Empty => Value::Empty,
                Expr::Null => Value::Null,
                Expr::Name(n) => {
                    // Constants of the constant's own module come first.
                    let prog: &'a Program = self.prog;
                    let local = prog.modules.get(module).and_then(|m| m.consts.iter().find(|(c, _, _)| c.eq_ignore_ascii_case(n)).map(|(_, e, _)| e));
                    match local {
                        Some(e) => self.const_value(module, e)?,
                        None => self.constant(n)?.ok_or_else(|| self.fail(compile_err(format!("Constant expression required ({n})"))))?,
                    }
                }
                Expr::Unary(op, a) => {
                    let a = self.const_value(module, a)?;
                    self.x(ops::unary(*op, &a))?
                }
                Expr::Binary(op, a, b) => {
                    let (a, b) = (self.const_value(module, a)?, self.const_value(module, b)?);
                    self.x(ops::binary(*op, &a, &b, false))?
                }
                Expr::Call(f, args) => {
                    // A few functions are fine in constants in practice (Chr, RGB).
                    let Expr::Name(fname) = &**f else { return Err(self.fail(compile_err("Constant expression required"))) };
                    let mut vals = Vec::new();
                    for a in args {
                        match &a.value {
                            Some(e) => vals.push(self.const_value(module, e)?),
                            None => vals.push(Value::Missing),
                        }
                    }
                    let r = builtins::call(fname, &vals, 0, false, &mut self.rnd).ok_or_else(|| self.fail(compile_err("Constant expression required")))?;
                    self.x(r)?
                }
                _ => return Err(self.fail(compile_err("Constant expression required"))),
            })
        })();
        self.const_depth = self.const_depth.saturating_sub(1);
        r
    }

    /// The procedure called `name`: in this module, else a public one elsewhere.
    fn find_proc(&self, name: &str) -> Option<(usize, usize)> {
        let here = self.frames.last().map(|f| f.module)?;
        let in_module = |mi: usize, public_only: bool| {
            self.prog.modules.get(mi).and_then(|m| {
                m.procs.iter().position(|p| {
                    p.name.eq_ignore_ascii_case(name) && (!public_only || p.public) && p.kind != ProcKind::PropertyLet && p.kind != ProcKind::PropertySet
                })
            })
        };
        if let Some(p) = in_module(here, false) {
            return Some((here, p));
        }
        // `Module1.Proc` is reached through the member path; plain names look in the others.
        (0..self.prog.modules.len()).filter(|m| *m != here).find_map(|m| in_module(m, true).map(|p| (m, p)))
    }

    fn module_index(&self, name: &str) -> Option<usize> {
        self.prog.modules.iter().position(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// A global object: the language's own (`Err`, `Debug`), or one the host's model answers
    /// for (`Application`, `ActiveSheet`, `Range`, `ActivePresentation`).
    fn global(&mut self, name: &str, args: &[Value]) -> X<Option<Value>> {
        match name.to_ascii_lowercase().as_str() {
            "err" if args.is_empty() => Ok(Some(Value::Object(Obj::Err))),
            "debug" => Ok(Some(Value::Object(Obj::Debug))),
            _ => {
                let v = self.model.global(self.host, name, args);
                self.x(v)
            }
        }
    }

    // ------------------------------------------------------------ evaluation

    fn tick(&mut self) -> X<()> {
        self.steps += 1;
        if self.steps > MAX_STEPS {
            return Err(Fail {
                err: VbaError::with(0, "The macro ran too long (more than 20,000,000 steps) and was stopped"),
                line: self.frames.last().map(|f| f.line).unwrap_or(0),
                module: self.frames.last().map(|f| f.module).unwrap_or(0),
                fatal: true,
            });
        }
        Ok(())
    }

    fn eval(&mut self, e: &Expr) -> X<Value> {
        match e {
            Expr::Num(n, int) => Ok(num_literal(*n, *int)),
            Expr::Str(s) => Ok(Value::str(s.as_str())),
            Expr::Bool(b) => Ok(Value::Bool(*b)),
            Expr::Nothing => Ok(Value::Nothing),
            Expr::Empty => Ok(Value::Empty),
            Expr::Null => Ok(Value::Null),
            Expr::New(t) => match t.to_ascii_lowercase().as_str() {
                "collection" => Ok(Value::Object(Obj::Collection(Rc::new(RefCell::new(Vec::new()))))),
                _ => Err(self.fail(VbaError::with(429, format!("Can't create a {t} (class modules and libraries aren't supported yet)")))),
            },
            Expr::Name(n) => self.name_value(n),
            Expr::Member(obj, m) => {
                // `Module1.Proc` and `Module1.Variable`.
                if let Expr::Name(modname) = &**obj
                    && self.lookup(modname).is_none()
                    && let Some(mi) = self.module_index(modname)
                {
                    return self.module_member(mi, m, Vec::new());
                }
                let o = self.eval(obj)?;
                self.member(o, m, Vec::new(), Vec::new())
            }
            Expr::WithMember(m) => {
                let o = self.with_object()?;
                self.member(o, m, Vec::new(), Vec::new())
            }
            Expr::Call(callee, args) => self.call_expr(callee, args),
            Expr::Unary(op, a) => {
                let a = self.eval_value(a)?;
                self.x(ops::unary(*op, &a))
            }
            Expr::Binary(op, a, b) => {
                if *op == BinOp::Is {
                    let (a, b) = (self.eval(a)?, self.eval(b)?);
                    return self.x(ops::binary(*op, &a, &b, false));
                }
                let a = self.eval_value(a)?;
                let b = self.eval_value(b)?;
                let text = self.compare_text();
                self.x(ops::binary(*op, &a, &b, text))
            }
        }
    }

    /// An expression's value with objects read through their default member (`Range` → `Value`).
    fn eval_value(&mut self, e: &Expr) -> X<Value> {
        let v = self.eval(e)?;
        self.deref(v)
    }

    fn deref(&mut self, v: Value) -> X<Value> {
        match v {
            Value::Object(o) if o.is_host() => {
                let r = self.model.default_value(self.host, &o);
                self.x(r)
            }
            Value::Object(Obj::Err) => Ok(Value::Long(self.err.0)),
            Value::Object(_) => Err(self.fail(VbaError::new(438))),
            Value::Nothing => Err(self.fail(VbaError::new(91))),
            other => Ok(other),
        }
    }

    fn with_object(&self) -> X<Value> {
        self.frames.last().and_then(|f| f.with.last().cloned()).ok_or_else(|| self.fail(compile_err("Invalid or unqualified reference")))
    }

    fn name_value(&mut self, n: &str) -> X<Value> {
        if let Some(s) = self.lookup(n) {
            return Ok(s.borrow().value.clone());
        }
        if let Some(v) = self.constant(n)? {
            return Ok(v);
        }
        if let Some((m, p)) = self.find_proc(n) {
            return self.call_proc(m, p, Vec::new(), Vec::new());
        }
        if let Some(v) = self.global(n, &[])? {
            return Ok(v);
        }
        if let Some(r) = builtins::call(n, &[], 0, false, &mut self.rnd) {
            return self.x(r);
        }
        if self.module().is_some_and(|m| m.option_explicit) {
            return Err(self.fail(compile_err(format!("Variable not defined ({n})"))));
        }
        Ok(Value::Empty)
    }

    fn module_member(&mut self, mi: usize, m: &str, args: Vec<ArgVal>) -> X<Value> {
        let prog: &'a Program = self.prog;
        let module = prog.modules.get(mi).ok_or_else(|| self.fail(VbaError::new(438)))?;
        if let Some(p) = module.procs.iter().position(|p| p.name.eq_ignore_ascii_case(m)) {
            return self.call_proc(mi, p, args, Vec::new());
        }
        if let Some(s) = self.globals.get(mi).and_then(|g| g.get(&m.to_ascii_lowercase())) {
            return Ok(s.borrow().value.clone());
        }
        // A document module is also the document (`ThisDocument.Content`, `ThisWorkbook.Sheets`).
        if let Some(o) = self.global(&module.name, &[])? {
            let vals = args.iter().map(ArgVal::value).collect();
            return self.member(o, m, vals, Vec::new());
        }
        Err(self.fail(compile_err(format!("Method or data member not found ({m})"))))
    }

    /// Arguments as values (left-out ones as `Missing`), and the named ones.
    fn eval_args(&mut self, args: &[Arg]) -> X<(Vec<Value>, Vec<(String, Value)>)> {
        let mut pos = Vec::new();
        let mut named = Vec::new();
        for a in args {
            let v = match &a.value {
                Some(e) => self.eval(e)?,
                None => Value::Missing,
            };
            match &a.name {
                Some(n) => named.push((n.clone(), v)),
                None => pos.push(v),
            }
        }
        Ok((pos, named))
    }

    /// Arguments for a procedure: variables by reference.
    fn proc_args(&mut self, args: &[Arg]) -> X<(Vec<ArgVal>, Vec<(String, ArgVal)>)> {
        let mut pos = Vec::new();
        let mut named = Vec::new();
        for a in args {
            let v = match &a.value {
                Some(Expr::Name(n)) => match self.lookup(n) {
                    Some(s) => ArgVal::Ref(s),
                    None => ArgVal::Val(self.eval(&Expr::Name(n.clone()))?),
                },
                Some(e) => ArgVal::Val(self.eval(e)?),
                None => ArgVal::Val(Value::Missing),
            };
            match &a.name {
                Some(n) => named.push((n.clone(), v)),
                None => pos.push(v),
            }
        }
        Ok((pos, named))
    }

    fn index_array(&self, a: &Rc<RefCell<Array>>, args: &[Value]) -> X<Value> {
        let idx = args.iter().map(Value::to_i64).collect::<R<Vec<_>>>();
        let idx = self.x(idx)?;
        let r = a.borrow().get(&idx);
        self.x(r)
    }

    fn call_expr(&mut self, callee: &Expr, args: &[Arg]) -> X<Value> {
        match callee {
            Expr::Name(n) => {
                if let Some(s) = self.lookup(n) {
                    let v = s.borrow().value.clone();
                    let (vals, _) = self.eval_args(args)?;
                    return match v {
                        Value::Array(a) => self.index_array(&a, &vals),
                        Value::Object(o) => self.default_call(o, vals),
                        _ => Err(self.fail(VbaError::new(13))),
                    };
                }
                if let Some((m, p)) = self.find_proc(n) {
                    let (pos, named) = self.proc_args(args)?;
                    return self.call_proc(m, p, pos, named);
                }
                let (vals, named) = self.eval_args(args)?;
                if let Some(v) = self.special(n, &vals, &named)? {
                    return Ok(v);
                }
                if let Some(v) = self.global(n, &vals)? {
                    return Ok(v);
                }
                let derefd = vals.into_iter().map(|v| self.deref_arg(v)).collect::<X<Vec<_>>>()?;
                let base = self.module().map(|m| m.option_base).unwrap_or(0);
                let text = self.compare_text();
                match builtins::call(n, &derefd, base, text, &mut self.rnd) {
                    Some(r) => self.x(r),
                    None => Err(self.fail(compile_err(format!("Sub or Function not defined ({n})")))),
                }
            }
            Expr::Member(obj, m) => {
                if let Expr::Name(modname) = &**obj
                    && self.lookup(modname).is_none()
                    && let Some(mi) = self.module_index(modname)
                {
                    let (pos, _) = self.proc_args(args)?;
                    return self.module_member(mi, m, pos);
                }
                let o = self.eval(obj)?;
                let (vals, named) = self.eval_args(args)?;
                self.member(o, m, vals, named)
            }
            Expr::WithMember(m) => {
                let o = self.with_object()?;
                let (vals, named) = self.eval_args(args)?;
                self.member(o, m, vals, named)
            }
            other => {
                let v = self.eval(other)?;
                let (vals, _) = self.eval_args(args)?;
                match v {
                    Value::Array(a) => self.index_array(&a, &vals),
                    Value::Object(o) => self.default_call(o, vals),
                    _ => Err(self.fail(VbaError::new(13))),
                }
            }
        }
    }

    /// Arguments to built-in functions are values: a range becomes its value, but arrays stay.
    fn deref_arg(&mut self, v: Value) -> X<Value> {
        match v {
            Value::Object(ref o) if Model::is_value(o) => self.deref(v),
            Value::Object(Obj::Err) => self.deref(v),
            other => Ok(other),
        }
    }

    /// Built-ins that need the host or the program: MsgBox, Application.Run-style calls.
    fn special(&mut self, n: &str, vals: &[Value], named: &[(String, Value)]) -> X<Option<Value>> {
        match n.to_ascii_lowercase().as_str() {
            "msgbox" => {
                let get = |i: usize, key: &str| {
                    vals.get(i)
                        .filter(|v| !matches!(v, Value::Missing))
                        .cloned()
                        .or_else(|| named.iter().find(|(k, _)| k.eq_ignore_ascii_case(key)).map(|(_, v)| v.clone()))
                };
                let prompt = match get(0, "prompt") {
                    Some(v) => {
                        let v = self.deref(v)?;
                        self.x(v.to_str())?
                    }
                    None => return Err(self.fail(VbaError::new(449))),
                };
                let mut params = json!({"text": prompt});
                if let Some(t) = get(2, "title") {
                    let t = self.deref(t)?;
                    let title = self.x(t.to_str())?;
                    if let Some(o) = params.as_object_mut() {
                        o.insert("title".into(), json!(title));
                    }
                }
                let flushed = self.model.flush(self.host);
                self.x(flushed)?;
                let out = self.host.call("host/message", params).map_err(|e| self.fail(VbaError::with(1004, e.message)))?;
                let result = out.get("result").and_then(serde_json::Value::as_i64).unwrap_or(1);
                Ok(Some(Value::Int(result as i16)))
            }
            _ => Ok(None),
        }
    }

    fn default_call(&mut self, o: Obj, vals: Vec<Value>) -> X<Value> {
        match o {
            x if x.is_host() => {
                let r = self.model.default_call(self.host, &x, &vals);
                self.x(r)
            }
            Obj::Collection(c) => self.collection(&c, "item", vals),
            _ => Err(self.fail(VbaError::new(438))),
        }
    }

    fn member(&mut self, o: Value, m: &str, vals: Vec<Value>, named: Vec<(String, Value)>) -> X<Value> {
        let lower = m.to_ascii_lowercase();
        match o {
            Value::Object(ref app) if lower == "run" && Model::is_application(app) => {
                let name = vals.first().cloned().ok_or_else(|| self.fail(VbaError::new(449)))?;
                let name = self.x(name.to_str())?;
                let target = match name.rsplit_once('!') {
                    Some((_, n)) => n.to_string(),
                    None => name,
                };
                let (mi, pi) = match target.split_once('.') {
                    Some((mo, p)) => {
                        let mi = self.module_index(mo).ok_or_else(|| self.fail(VbaError::with(1004, format!("Cannot run the macro '{target}'"))))?;
                        let pi = self
                            .prog
                            .modules
                            .get(mi)
                            .and_then(|x| x.procs.iter().position(|q| q.name.eq_ignore_ascii_case(p)))
                            .ok_or_else(|| self.fail(VbaError::with(1004, format!("Cannot run the macro '{target}'"))))?;
                        (mi, pi)
                    }
                    None => self.find_proc(&target).ok_or_else(|| self.fail(VbaError::with(1004, format!("Cannot run the macro '{target}'"))))?,
                };
                let args = vals.into_iter().skip(1).map(ArgVal::Val).collect();
                self.call_proc(mi, pi, args, Vec::new())
            }
            Value::Object(x) if x.is_host() => {
                let vals = vals.into_iter().map(|v| self.deref_arg(v)).collect::<X<Vec<_>>>()?;
                let r = self.model.member(self.host, &x, m, &vals, &named);
                self.x(r)
            }
            Value::Object(Obj::Err) => match lower.as_str() {
                "number" => Ok(Value::Long(self.err.0)),
                "description" => Ok(Value::str(self.err.1.as_str())),
                "source" => Ok(Value::str(self.err.2.as_str())),
                "clear" => {
                    self.err = (0, String::new(), String::new());
                    Ok(Value::Empty)
                }
                "raise" => {
                    let get = |i: usize, key: &str| {
                        vals.get(i)
                            .filter(|v| !matches!(v, Value::Missing))
                            .cloned()
                            .or_else(|| named.iter().find(|(k, _)| k.eq_ignore_ascii_case(key)).map(|(_, v)| v.clone()))
                    };
                    let number = get(0, "number").ok_or_else(|| self.fail(VbaError::new(449)))?;
                    let number = self.x(number.to_i32())?;
                    let description = match get(2, "description") {
                        Some(d) => self.x(d.to_str())?,
                        None => super::value::describe(number).to_string(),
                    };
                    if let Some(src) = get(1, "source") {
                        self.err.2 = self.x(src.to_str())?;
                    }
                    Err(self.fail(VbaError::with(number, description)))
                }
                _ => Err(self.fail(VbaError::new(438))),
            },
            Value::Object(Obj::Debug) => match lower.as_str() {
                "print" => {
                    let parts = vals.iter().map(Value::to_str).collect::<R<Vec<_>>>();
                    let text = self.x(parts)?.join(" ");
                    self.log(&text)?;
                    Ok(Value::Empty)
                }
                "assert" => Ok(Value::Empty),
                _ => Err(self.fail(VbaError::new(438))),
            },
            Value::Object(Obj::Collection(c)) => self.collection(&c, &lower, vals),
            Value::Nothing => Err(self.fail(VbaError::new(91))),
            _ => Err(self.fail(VbaError::new(424))),
        }
    }

    fn collection(&mut self, c: &Rc<RefCell<Vec<(Option<String>, Value)>>>, m: &str, vals: Vec<Value>) -> X<Value> {
        let find = |c: &Vec<(Option<String>, Value)>, key: &Value| -> R<usize> {
            match key {
                Value::Str(k) => c.iter().position(|(n, _)| n.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(k))).ok_or(VbaError::new(5)),
                v => {
                    let i = v.to_i64()?;
                    usize::try_from(i - 1).ok().filter(|i| *i < c.len()).ok_or(VbaError::new(9))
                }
            }
        };
        match m {
            "add" => {
                let item = vals.first().cloned().ok_or_else(|| self.fail(VbaError::new(449)))?;
                let key = match vals.get(1) {
                    Some(Value::Missing) | None => None,
                    Some(k) => Some(self.x(k.to_str())?),
                };
                if let Some(k) = &key
                    && c.borrow().iter().any(|(n, _)| n.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(k)))
                {
                    return Err(self.fail(VbaError::with(457, "This key is already associated with an element of this collection")));
                }
                if c.borrow().len() >= super::value::MAX_ELEMENTS {
                    return Err(self.fail(VbaError::new(7)));
                }
                c.borrow_mut().push((key, item));
                Ok(Value::Empty)
            }
            "count" => Ok(Value::Long(c.borrow().len() as i32)),
            "item" => {
                let key = vals.first().ok_or_else(|| self.fail(VbaError::new(449)))?;
                let i = self.x(find(&c.borrow(), key))?;
                Ok(c.borrow().get(i).map(|(_, v)| v.clone()).unwrap_or(Value::Empty))
            }
            "remove" => {
                let key = vals.first().ok_or_else(|| self.fail(VbaError::new(449)))?;
                let i = self.x(find(&c.borrow(), key))?;
                c.borrow_mut().remove(i);
                Ok(Value::Empty)
            }
            _ => Err(self.fail(VbaError::new(438))),
        }
    }

    fn log(&mut self, text: &str) -> X<()> {
        self.host.call("host/log", json!({"text": text})).map(|_| ()).map_err(|e| self.fail(VbaError::with(1004, e.message)))
    }

    // ------------------------------------------------------------ assignment

    fn store(&self, s: &Slot, v: Value, set: bool) -> X<()> {
        let (ty, array) = (s.borrow().ty, s.borrow().array);
        if set && !matches!(v, Value::Object(_) | Value::Nothing) {
            return Err(self.fail(VbaError::new(424)));
        }
        let v = if array {
            match v {
                Value::Array(a) => Value::Array(Rc::new(RefCell::new(a.borrow().clone()))),
                _ => return Err(self.fail(VbaError::with(13, "Can't assign to array"))),
            }
        } else {
            self.x(coerce(ty, v))?
        };
        s.borrow_mut().value = v;
        Ok(())
    }

    /// The variable `name`, declared on first use where Option Explicit allows it.
    fn var(&mut self, name: &str) -> X<Slot> {
        if let Some(s) = self.lookup(name) {
            return Ok(s);
        }
        if self.module().is_some_and(|m| m.option_explicit) {
            return Err(self.fail(compile_err(format!("Variable not defined ({name})"))));
        }
        let s = slot(Value::Empty, Ty::Variant);
        if let Some(f) = self.frame_mut() {
            f.locals.insert(name.to_ascii_lowercase(), s.clone());
        }
        Ok(s)
    }

    fn assign(&mut self, target: &Expr, value: &Expr, set: bool) -> X<()> {
        let v = if set { self.eval(value)? } else { self.eval_value(value)? };
        self.assign_value(target, v, set)
    }

    fn assign_value(&mut self, target: &Expr, v: Value, set: bool) -> X<()> {
        match target {
            Expr::Name(n) => {
                let s = self.var(n)?;
                // `rng = 5` with rng a Range sets its value (the default member).
                let current = s.borrow().value.clone();
                if !set
                    && let Value::Object(o) = current
                    && o.is_host()
                {
                    let r = self.model.assign_default(self.host, &o, &v);
                    return self.x(r);
                }
                self.store(&s, v, set)
            }
            Expr::Call(callee, args) => {
                let (vals, _) = self.eval_args(args)?;
                match &**callee {
                    Expr::Name(n) => {
                        if let Some(s) = self.lookup(n) {
                            let current = s.borrow().value.clone();
                            return match current {
                                Value::Array(a) => {
                                    let ty = s.borrow().ty;
                                    let v = self.x(coerce(ty, v))?;
                                    let idx = vals.iter().map(Value::to_i64).collect::<R<Vec<_>>>();
                                    let idx = self.x(idx)?;
                                    let r = a.borrow_mut().set(&idx, v);
                                    self.x(r)
                                }
                                Value::Object(o) if o.is_host() => {
                                    let r = self.model.assign(self.host, &o, "item", &vals, &v);
                                    self.x(r)
                                }
                                _ => Err(self.fail(VbaError::new(13))),
                            };
                        }
                        // `Range("A1") = 5`, `Cells(1, 1) = 5`.
                        let vals = vals.into_iter().map(|v| self.deref_arg(v)).collect::<X<Vec<_>>>()?;
                        let r = self.model.assign_global(self.host, n, &vals, &v);
                        self.x(r)
                    }
                    Expr::Member(obj, m) => {
                        let o = self.eval(obj)?;
                        self.member_assign(o, m, vals, v)
                    }
                    Expr::WithMember(m) => {
                        let o = self.with_object()?;
                        self.member_assign(o, m, vals, v)
                    }
                    other => {
                        let o = self.eval(other)?;
                        self.member_assign(o, "item", vals, v)
                    }
                }
            }
            Expr::Member(obj, m) => {
                if let Expr::Name(modname) = &**obj
                    && self.lookup(modname).is_none()
                    && let Some(mi) = self.module_index(modname)
                    && let Some(s) = self.globals.get(mi).and_then(|g| g.get(&m.to_ascii_lowercase())).cloned()
                {
                    return self.store(&s, v, set);
                }
                let o = self.eval(obj)?;
                self.member_assign(o, m, Vec::new(), v)
            }
            Expr::WithMember(m) => {
                let o = self.with_object()?;
                self.member_assign(o, m, Vec::new(), v)
            }
            _ => Err(self.fail(compile_err("Invalid assignment target"))),
        }
    }

    fn member_assign(&mut self, o: Value, m: &str, args: Vec<Value>, v: Value) -> X<()> {
        match o {
            Value::Object(x) if x.is_host() => {
                let args = args.into_iter().map(|a| self.deref_arg(a)).collect::<X<Vec<_>>>()?;
                // A value assigned to a range property is the other range's value (`a.Value = b`).
                let v = match v {
                    Value::Object(ref o) if Model::is_value(o) => self.deref(v)?,
                    other => other,
                };
                let r = self.model.assign(self.host, &x, m, &args, &v);
                self.x(r)
            }
            Value::Object(Obj::Err) => match m.to_ascii_lowercase().as_str() {
                "number" => {
                    self.err.0 = self.x(v.to_i32())?;
                    Ok(())
                }
                "description" => {
                    self.err.1 = self.x(v.to_str())?;
                    Ok(())
                }
                "source" => {
                    self.err.2 = self.x(v.to_str())?;
                    Ok(())
                }
                _ => Err(self.fail(VbaError::new(438))),
            },
            Value::Nothing => Err(self.fail(VbaError::new(91))),
            _ => Err(self.fail(VbaError::new(424))),
        }
    }

    // ------------------------------------------------------------ procedures

    fn call_proc(&mut self, module: usize, proc: usize, pos: Vec<ArgVal>, named: Vec<(String, ArgVal)>) -> X<Value> {
        let prog: &'a Program = self.prog;
        let p = prog.modules.get(module).and_then(|m| m.procs.get(proc)).ok_or_else(|| self.fail(VbaError::new(35)))?;
        if self.frames.len() >= MAX_DEPTH {
            return Err(self.fail(VbaError::new(28)));
        }
        let regular = p.params.iter().filter(|q| !q.param_array).count();
        if pos.len() > regular && !p.params.iter().any(|q| q.param_array) {
            return Err(self.fail(VbaError::new(450)));
        }
        let mut frame = Frame {
            module,
            locals: HashMap::new(),
            consts: HashMap::new(),
            with: Vec::new(),
            on_error: OnErr::Off,
            handling: false,
            resume_at: 0,
            line: p.line,
        };
        // Bind the parameters.
        let mut pos = pos.into_iter();
        let mut pending_defaults = Vec::new();
        for q in &p.params {
            let key = q.name.to_ascii_lowercase();
            let ty = ty_of(q.ty.as_deref());
            if q.param_array {
                let rest: Vec<Value> = pos.by_ref().map(|a| a.value()).collect();
                frame.locals.insert(key, slot(Value::Array(Rc::new(RefCell::new(Array::from_vec(rest, 0)))), Ty::Variant));
                continue;
            }
            let given = pos.next().or_else(|| named.iter().find(|(n, _)| n.eq_ignore_ascii_case(&q.name)).map(|(_, a)| a.clone()));
            let given = given.filter(|a| !matches!(a, ArgVal::Val(Value::Missing)));
            let s = match given {
                Some(ArgVal::Ref(s)) if !q.by_val => s,
                Some(a) => {
                    let v = a.value();
                    let v = coerce(ty, v).map_err(|e| self.fail(VbaError::with(e.number, format!("{} (argument {})", e.description, q.name))))?;
                    slot(v, ty)
                }
                None if q.optional => {
                    if let Some(d) = &q.default {
                        pending_defaults.push((key.clone(), d, ty));
                    }
                    slot(if ty == Ty::Variant && q.default.is_none() { Value::Missing } else { default_of(ty) }, ty)
                }
                None => return Err(self.fail(VbaError::with(449, format!("Argument not optional ({})", q.name)))),
            };
            frame.locals.insert(key, s);
        }
        if matches!(p.kind, ProcKind::Function | ProcKind::PropertyGet) {
            let ty = ty_of(p.ret_ty.as_deref());
            frame.locals.insert(p.name.to_ascii_lowercase(), slot(default_of(ty), ty));
        }
        self.frames.push(frame);
        let r = (|| -> X<Value> {
            for (key, e, ty) in pending_defaults {
                let v = self.const_value(module, e)?;
                let v = self.x(coerce(ty, v))?;
                if let Some(s) = self.frames.last().and_then(|f| f.locals.get(&key)) {
                    s.borrow_mut().value = v;
                }
            }
            self.run_body(&p.body)?;
            let ret = self.frames.last().and_then(|f| f.locals.get(&p.name.to_ascii_lowercase())).map(|s| s.borrow().value.clone());
            Ok(if matches!(p.kind, ProcKind::Function | ProcKind::PropertyGet) { ret.unwrap_or(Value::Empty) } else { Value::Empty })
        })();
        self.frames.pop();
        r
    }

    /// A procedure's body, with its error handler: a run-time error goes to the `On Error GoTo`
    /// label; `Resume Next` continues after the statement that failed.
    fn run_body(&mut self, body: &[Line]) -> X<()> {
        let mut start = 0usize;
        loop {
            match self.block_from(body, start, true) {
                Ok(Flow::Next | Flow::Exit(ExitKind::Sub | ExitKind::Function | ExitKind::Property)) => return Ok(()),
                Ok(Flow::End) => return Err(Fail { err: VbaError::with(END, "End"), line: 0, module: 0, fatal: true }),
                Ok(Flow::Exit(_)) => return Ok(()),
                Ok(Flow::GoTo(label)) => {
                    start = self.label_index(body, &label)?;
                }
                Ok(Flow::ResumeNext) => {
                    let f = self.frames.last_mut();
                    let at = f.as_ref().map(|f| f.resume_at).unwrap_or(0);
                    if let Some(f) = f {
                        f.handling = false;
                    }
                    self.err = (0, String::new(), String::new());
                    start = at + 1;
                }
                Err(fail) => {
                    let (handler, handling) = match self.frames.last() {
                        Some(f) => (f.on_error.clone(), f.handling),
                        None => (OnErr::Off, false),
                    };
                    match handler {
                        OnErr::GoTo(label) if !fail.fatal && !handling => {
                            self.err = (fail.err.number, fail.err.description.clone(), self.module().map(|m| m.name.clone()).unwrap_or_default());
                            if let Some(f) = self.frames.last_mut() {
                                f.handling = true;
                            }
                            start = self.label_index(body, &label)?;
                        }
                        _ => return Err(fail),
                    }
                }
            }
        }
    }

    fn label_index(&self, body: &[Line], label: &str) -> X<usize> {
        body.iter()
            .position(|l| matches!(&l.stmt, Stmt::Label(x) if x.eq_ignore_ascii_case(label)))
            .ok_or_else(|| self.fail(compile_err(format!("Label not defined ({label})"))))
    }

    /// Runs `lines` from `start`. `top`: the procedure's own statements (where errors are
    /// remembered for `Resume Next`).
    fn block_from(&mut self, lines: &[Line], start: usize, top: bool) -> X<Flow> {
        let mut i = start;
        while let Some(line) = lines.get(i) {
            if let Some(f) = self.frame_mut() {
                f.line = line.line;
                if top {
                    f.resume_at = i;
                }
            }
            self.tick()?;
            match self.exec(line) {
                Ok(Flow::Next) => i += 1,
                Ok(Flow::GoTo(label)) => match lines.iter().position(|l| matches!(&l.stmt, Stmt::Label(x) if x.eq_ignore_ascii_case(&label))) {
                    Some(j) => i = j + 1,
                    None => return Ok(Flow::GoTo(label)),
                },
                Ok(other) => return Ok(other),
                Err(fail) => {
                    let resume = matches!(self.frames.last().map(|f| &f.on_error), Some(OnErr::ResumeNext));
                    if resume && !fail.fatal {
                        self.err = (fail.err.number, fail.err.description, self.module().map(|m| m.name.clone()).unwrap_or_default());
                        i += 1;
                    } else {
                        return Err(fail);
                    }
                }
            }
        }
        Ok(Flow::Next)
    }

    fn block(&mut self, lines: &[Line]) -> X<Flow> {
        self.block_from(lines, 0, false)
    }

    fn cond(&mut self, e: &Expr) -> X<bool> {
        let v = self.eval_value(e)?;
        if matches!(v, Value::Null) {
            return Ok(false);
        }
        self.x(v.to_bool())
    }

    fn bounds(&mut self, dims: &[(Option<Expr>, Expr)]) -> X<Vec<(i64, i64)>> {
        let base = self.module().map(|m| m.option_base).unwrap_or(0);
        let mut out = Vec::new();
        for (lo, hi) in dims {
            let lo = match lo {
                Some(e) => {
                    let v = self.eval_value(e)?;
                    self.x(v.to_i64())?
                }
                None => base,
            };
            let hi = self.eval_value(hi)?;
            out.push((lo, self.x(hi.to_i64())?));
        }
        Ok(out)
    }

    fn const_bounds(&mut self, module: usize, dims: &[(Option<Expr>, Expr)]) -> X<Vec<(i64, i64)>> {
        let base = self.prog.modules.get(module).map(|m| m.option_base).unwrap_or(0);
        let mut out = Vec::new();
        for (lo, hi) in dims {
            let lo = match lo {
                Some(e) => {
                    let v = self.const_value(module, e)?;
                    self.x(v.to_i64())?
                }
                None => base,
            };
            let hi = self.const_value(module, hi)?;
            out.push((lo, self.x(hi.to_i64())?));
        }
        Ok(out)
    }

    fn new_array(&self, bounds: Vec<(i64, i64)>, ty: Ty) -> X<Value> {
        let mut a = self.x(Array::new(bounds))?;
        let d = default_of(ty);
        a.data.iter_mut().for_each(|x| *x = d.clone());
        Ok(Value::Array(Rc::new(RefCell::new(a))))
    }

    fn exec(&mut self, line: &Line) -> X<Flow> {
        match &line.stmt {
            Stmt::Dim(vars) => {
                for d in vars {
                    let key = d.name.to_ascii_lowercase();
                    // Declarations aren't executable: a Dim inside a loop doesn't reset the variable.
                    if self.frames.last().is_some_and(|f| f.locals.contains_key(&key)) {
                        continue;
                    }
                    let ty = ty_of(d.ty.as_deref());
                    let value = match &d.dims {
                        Some(dims) if !dims.is_empty() => {
                            let b = self.bounds(dims)?;
                            self.new_array(b, ty)?
                        }
                        Some(_) => Value::Empty,
                        None if d.new_object && d.ty.as_deref().is_some_and(|t| t.eq_ignore_ascii_case("collection")) => {
                            Value::Object(Obj::Collection(Rc::new(RefCell::new(Vec::new()))))
                        }
                        None => default_of(ty),
                    };
                    let s = if d.dims.is_some() { array_slot(value, ty) } else { slot(value, ty) };
                    if let Some(f) = self.frame_mut() {
                        f.locals.insert(key, s);
                    }
                }
                Ok(Flow::Next)
            }
            Stmt::ReDim { preserve, vars } => {
                for (name, dims) in vars {
                    let s = self.var(name)?;
                    let b = self.bounds(dims)?;
                    let ty = s.borrow().ty;
                    let fresh = self.new_array(b, ty)?;
                    if *preserve && let (Value::Array(old), Value::Array(new)) = (s.borrow().value.clone(), &fresh) {
                        let old = old.borrow();
                        let mut new = new.borrow_mut();
                        // Every index the old and new bounds share keeps its element.
                        if old.bounds.len() == new.bounds.len() {
                            let total = old.data.len();
                            for flat in 0..total {
                                let mut rest = flat;
                                let mut idx = vec![0i64; old.bounds.len()];
                                for (k, (lo, hi)) in old.bounds.iter().enumerate().rev() {
                                    let len = (hi - lo + 1).max(1) as usize;
                                    if let Some(slot) = idx.get_mut(k) {
                                        *slot = lo + (rest % len) as i64;
                                    }
                                    rest /= len;
                                }
                                if let (Ok(v), true) = (old.get(&idx), new.index(&idx).is_ok()) {
                                    let _ = new.set(&idx, v);
                                }
                            }
                        }
                    }
                    let mut v = s.borrow_mut();
                    v.value = fresh;
                    v.array = true;
                }
                Ok(Flow::Next)
            }
            Stmt::Const(list) => {
                let module = self.frame()?.module;
                for (name, e) in list {
                    let v = self.const_value(module, e)?;
                    if let Some(f) = self.frame_mut() {
                        f.consts.insert(name.to_ascii_lowercase(), v);
                    }
                }
                Ok(Flow::Next)
            }
            Stmt::Assign { target, value, set } => {
                self.assign(target, value, *set)?;
                Ok(Flow::Next)
            }
            Stmt::Call(callee, args) => {
                match callee {
                    Expr::Name(n) => {
                        if let Some((m, p)) = self.find_proc(n) {
                            let (pos, named) = self.proc_args(args)?;
                            self.call_proc(m, p, pos, named)?;
                        } else if args.is_empty() && self.lookup(n).is_some() {
                            return Err(self.fail(compile_err(format!("Expected Sub, Function, or Property ({n})"))));
                        } else {
                            self.call_expr(callee, args)?;
                        }
                    }
                    other => {
                        if args.is_empty() {
                            self.eval(other)?;
                        } else {
                            self.call_expr(other, args)?;
                        }
                    }
                }
                Ok(Flow::Next)
            }
            Stmt::If { branches, otherwise } => {
                for (c, body) in branches {
                    if self.cond(c)? {
                        return self.block(body);
                    }
                }
                match otherwise {
                    Some(body) => self.block(body),
                    None => Ok(Flow::Next),
                }
            }
            Stmt::Select { expr, cases, otherwise } => {
                let v = self.eval_value(expr)?;
                let text = self.compare_text();
                for (tests, body) in cases {
                    let mut hit = false;
                    for t in tests {
                        hit = match t {
                            CaseTest::Value(e) => {
                                let x = self.eval_value(e)?;
                                self.x(ops::compare(&v, &x, text))? == std::cmp::Ordering::Equal
                            }
                            CaseTest::Range(a, b) => {
                                let (a, b) = (self.eval_value(a)?, self.eval_value(b)?);
                                self.x(ops::compare(&v, &a, text))? != std::cmp::Ordering::Less
                                    && self.x(ops::compare(&v, &b, text))? != std::cmp::Ordering::Greater
                            }
                            CaseTest::Is(op, e) => {
                                let x = self.eval_value(e)?;
                                matches!(self.x(ops::binary(*op, &v, &x, text))?, Value::Bool(true))
                            }
                        };
                        if hit {
                            break;
                        }
                    }
                    if hit {
                        return self.block(body);
                    }
                }
                match otherwise {
                    Some(body) => self.block(body),
                    None => Ok(Flow::Next),
                }
            }
            Stmt::For { var, from, to, step, body } => {
                let start = self.eval_value(from)?;
                let end = self.eval_value(to)?;
                let step = match step {
                    Some(s) => self.eval_value(s)?,
                    None => Value::Int(1),
                };
                let (endf, stepf) = (self.x(end.to_f64())?, self.x(step.to_f64())?);
                self.assign_value(var, start, false)?;
                loop {
                    let cur = self.eval_value(var)?;
                    let c = self.x(cur.to_f64())?;
                    if (stepf >= 0.0 && c > endf) || (stepf < 0.0 && c < endf) {
                        break;
                    }
                    match self.block(body)? {
                        Flow::Next => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                    self.tick()?;
                    let cur = self.eval_value(var)?;
                    let next = self.x(ops::binary(BinOp::Add, &cur, &step, false))?;
                    self.assign_value(var, next, false)?;
                }
                Ok(Flow::Next)
            }
            Stmt::ForEach { var, collection, body } => {
                let c = self.eval(collection)?;
                let items: Vec<Value> = match c {
                    Value::Array(a) => a.borrow().data.clone(),
                    Value::Object(o) if o.is_host() => {
                        let r = self.model.items(self.host, &o);
                        self.x(r)?
                    }
                    Value::Object(Obj::Collection(c)) => c.borrow().iter().map(|(_, v)| v.clone()).collect(),
                    Value::Nothing => return Err(self.fail(VbaError::new(91))),
                    _ => return Err(self.fail(VbaError::new(424))),
                };
                for item in items {
                    let set = matches!(item, Value::Object(_));
                    self.assign_value(var, item, set)?;
                    match self.block(body)? {
                        Flow::Next => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                    self.tick()?;
                }
                Ok(Flow::Next)
            }
            Stmt::Do { pre, post, body } => {
                loop {
                    if let Some((until, c)) = pre
                        && self.cond(c)? == *until
                    {
                        break;
                    }
                    match self.block(body)? {
                        Flow::Next => {}
                        Flow::Exit(ExitKind::Do) => break,
                        other => return Ok(other),
                    }
                    if let Some((until, c)) = post
                        && self.cond(c)? == *until
                    {
                        break;
                    }
                    self.tick()?;
                }
                Ok(Flow::Next)
            }
            Stmt::While { cond, body } => {
                while self.cond(cond)? {
                    match self.block(body)? {
                        Flow::Next => {}
                        other => return Ok(other),
                    }
                    self.tick()?;
                }
                Ok(Flow::Next)
            }
            Stmt::Exit(k) => {
                if matches!(k, ExitKind::Sub | ExitKind::Function | ExitKind::Property) {
                    self.err = (0, String::new(), String::new());
                }
                Ok(Flow::Exit(*k))
            }
            Stmt::With { object, body } => {
                let o = self.eval(object)?;
                if let Some(f) = self.frame_mut() {
                    f.with.push(o);
                }
                let r = self.block(body);
                if let Some(f) = self.frame_mut() {
                    f.with.pop();
                }
                r
            }
            Stmt::OnError(on) => {
                self.err = (0, String::new(), String::new());
                if let Some(f) = self.frame_mut() {
                    f.on_error = match on {
                        OnError::ResumeNext => OnErr::ResumeNext,
                        OnError::GoTo(l) => OnErr::GoTo(l.clone()),
                        OnError::GoToZero => OnErr::Off,
                    };
                    if matches!(on, OnError::GoToZero) {
                        f.handling = false;
                    }
                }
                Ok(Flow::Next)
            }
            Stmt::GoTo(l) => Ok(Flow::GoTo(l.clone())),
            Stmt::Label(_) => Ok(Flow::Next),
            Stmt::ResumeNext => {
                if !self.frames.last().is_some_and(|f| f.handling) {
                    return Err(self.fail(VbaError::new(20)));
                }
                Ok(Flow::ResumeNext)
            }
            Stmt::ResumeLabel(l) => {
                if !self.frames.last().is_some_and(|f| f.handling) {
                    return Err(self.fail(VbaError::new(20)));
                }
                self.err = (0, String::new(), String::new());
                if let Some(f) = self.frame_mut() {
                    f.handling = false;
                }
                Ok(Flow::GoTo(l.clone()))
            }
            Stmt::Erase(names) => {
                for n in names {
                    let s = self.var(n)?;
                    let ty = s.borrow().ty;
                    let v = s.borrow().value.clone();
                    if let Value::Array(a) = v {
                        a.borrow_mut().data.iter_mut().for_each(|x| *x = default_of(ty));
                    }
                }
                Ok(Flow::Next)
            }
            Stmt::Print(items) => {
                let mut parts = Vec::new();
                for e in items {
                    let v = self.eval_value(e)?;
                    parts.push(self.x(v.to_str())?);
                }
                self.log(&parts.join(" "))?;
                Ok(Flow::Next)
            }
            Stmt::End => Ok(Flow::End),
        }
    }
}

fn num_literal(n: f64, int: bool) -> Value {
    if int { super::value::int_value(n) } else { Value::Double(n) }
}

impl Program {
    /// Parses the modules; one that doesn't compile is left out and remembered.
    pub fn new(modules: &[(String, String)]) -> Program {
        let mut p = Program { modules: Vec::new(), broken: Vec::new() };
        for (name, src) in modules {
            match super::parser::parse_module(name, src) {
                Ok(m) => p.modules.push(m),
                Err((line, why)) => p.broken.push((name.clone(), line, why)),
            }
        }
        p
    }
}
