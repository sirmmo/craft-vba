//! The syntax tree of a VBA module.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    IntDiv,
    Mod,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    Like,
    Is,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnOp {
    Neg,
    Not,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    /// A number; `true` when written as an integer.
    Num(f64, bool),
    Str(String),
    Bool(bool),
    Nothing,
    Empty,
    Null,
    Name(String),
    /// `object.member`.
    Member(Box<Expr>, String),
    /// `.member` inside `With`.
    WithMember(String),
    /// `f(args)`: a call, an array index or a default member.
    Call(Box<Expr>, Vec<Arg>),
    Unary(UnOp, Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    /// `New Collection`.
    New(String),
}

#[derive(Clone, Debug, PartialEq)]
pub struct Arg {
    /// `name:=value`.
    pub name: Option<String>,
    /// `None`: left out (`f(a, , c)`).
    pub value: Option<Expr>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VarDecl {
    pub name: String,
    /// `None`: not an array; `Some(vec![])`: a dynamic array `a()`; bounds otherwise.
    pub dims: Option<Vec<(Option<Expr>, Expr)>>,
    pub ty: Option<String>,
    pub new_object: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub enum CaseTest {
    Value(Expr),
    Range(Expr, Expr),
    Is(BinOp, Expr),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    Property,
    For,
    Do,
}

#[derive(Clone, Debug, PartialEq)]
pub enum OnError {
    ResumeNext,
    GoTo(String),
    GoToZero,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Stmt {
    Dim(Vec<VarDecl>),
    ReDim {
        preserve: bool,
        vars: Vec<(String, Vec<(Option<Expr>, Expr)>)>,
    },
    Const(Vec<(String, Expr)>),
    /// `target = value`, or `Set target = value`.
    Assign {
        target: Expr,
        value: Expr,
        set: bool,
    },
    /// A call statement: `Foo a, b`, `Call Foo(a, b)`, `obj.Method a`.
    Call(Expr, Vec<Arg>),
    If {
        branches: Vec<(Expr, Vec<Line>)>,
        otherwise: Option<Vec<Line>>,
    },
    Select {
        expr: Expr,
        cases: Vec<(Vec<CaseTest>, Vec<Line>)>,
        otherwise: Option<Vec<Line>>,
    },
    For {
        var: Expr,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Line>,
    },
    ForEach {
        var: Expr,
        collection: Expr,
        body: Vec<Line>,
    },
    /// `Do [While|Until c] … Loop [While|Until c]`: (until, condition).
    Do {
        pre: Option<(bool, Expr)>,
        post: Option<(bool, Expr)>,
        body: Vec<Line>,
    },
    While {
        cond: Expr,
        body: Vec<Line>,
    },
    Exit(ExitKind),
    With {
        object: Expr,
        body: Vec<Line>,
    },
    OnError(OnError),
    GoTo(String),
    Label(String),
    /// `Resume Next`.
    ResumeNext,
    /// `Resume label`: leaves the error handler for the label.
    ResumeLabel(String),
    /// `Erase a`.
    Erase(Vec<String>),
    /// `Debug.Print a; b`.
    Print(Vec<Expr>),
    /// `End`: stops the macro.
    End,
}

/// A statement and the line it starts on.
#[derive(Clone, Debug, PartialEq)]
pub struct Line {
    pub line: u32,
    pub stmt: Stmt,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Param {
    pub name: String,
    pub by_val: bool,
    pub optional: bool,
    pub default: Option<Expr>,
    pub param_array: bool,
    pub ty: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Proc {
    pub name: String,
    pub kind: ProcKind,
    pub params: Vec<Param>,
    pub body: Vec<Line>,
    pub public: bool,
    pub ret_ty: Option<String>,
    pub line: u32,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ModuleAst {
    pub name: String,
    pub option_explicit: bool,
    pub option_base: i64,
    pub compare_text: bool,
    /// Module-level variables: (declaration, public).
    pub vars: Vec<(VarDecl, bool)>,
    pub consts: Vec<(String, Expr, bool)>,
    pub procs: Vec<Proc>,
}
