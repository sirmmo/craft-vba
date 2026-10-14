//! Parses a module's tokens into a [`ModuleAst`] ([MS-VBAL] §5). Errors are VBA's compile errors:
//! (line, message).

use super::ast::*;
use super::lexer::{Tok, Token, lex};

pub type PResult<T> = Result<T, (u32, String)>;

/// Deepest nesting of blocks and expressions.
const MAX_DEPTH: u32 = 100;

struct Parser {
    toks: Vec<Token>,
    pos: usize,
    depth: u32,
}

/// Words that end a block (checked at the start of a statement).
type Terminator = &'static [&'static str];

const END_IF: Terminator = &["end", "if"];
const ELSE: Terminator = &["else"];
const ELSEIF: Terminator = &["elseif"];
const NEXT: Terminator = &["next"];
const LOOP: Terminator = &["loop"];
const WEND: Terminator = &["wend"];
const CASE: Terminator = &["case"];
const END_SELECT: Terminator = &["end", "select"];
const END_WITH: Terminator = &["end", "with"];
const END_SUB: Terminator = &["end", "sub"];
const END_FUNCTION: Terminator = &["end", "function"];
const END_PROPERTY: Terminator = &["end", "property"];

pub fn parse_module(name: &str, src: &str) -> PResult<ModuleAst> {
    let toks = lex(src)?;
    let mut p = Parser { toks, pos: 0, depth: 0 };
    p.module(name)
}

impl Parser {
    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.pos).map(|t| &t.tok)
    }
    fn peek_at(&self, n: usize) -> Option<&Tok> {
        self.toks.get(self.pos + n).map(|t| &t.tok)
    }
    fn line(&self) -> u32 {
        self.toks.get(self.pos).or_else(|| self.toks.last()).map(|t| t.line).unwrap_or(1)
    }
    fn bump(&mut self) -> Option<Tok> {
        let t = self.toks.get(self.pos).map(|t| t.tok.clone());
        self.pos += 1;
        t
    }
    fn err<T>(&self, msg: impl Into<String>) -> PResult<T> {
        Err((self.line(), msg.into()))
    }
    fn is_kw_at(&self, n: usize, kw: &str) -> bool {
        matches!(self.peek_at(n), Some(Tok::Ident(s)) if s.eq_ignore_ascii_case(kw))
    }
    fn is_kw(&self, kw: &str) -> bool {
        self.is_kw_at(0, kw)
    }
    fn eat_kw(&mut self, kw: &str) -> bool {
        if self.is_kw(kw) {
            self.pos += 1;
            true
        } else {
            false
        }
    }
    fn expect_kw(&mut self, kw: &str) -> PResult<()> {
        if self.eat_kw(kw) { Ok(()) } else { self.err(format!("Expected: {kw}")) }
    }
    fn eat(&mut self, t: &Tok) -> bool {
        if self.peek() == Some(t) {
            self.pos += 1;
            true
        } else {
            false
        }
    }
    fn expect(&mut self, t: &Tok, what: &str) -> PResult<()> {
        if self.eat(t) { Ok(()) } else { self.err(format!("Expected: {what}")) }
    }
    fn ident(&mut self) -> PResult<String> {
        match self.peek() {
            Some(Tok::Ident(s)) => {
                let s = s.clone();
                self.pos += 1;
                Ok(s)
            }
            _ => self.err("Expected: identifier"),
        }
    }
    fn at_stmt_end(&self) -> bool {
        matches!(self.peek(), None | Some(Tok::Newline | Tok::Colon)) || self.is_kw("else")
    }
    fn skip_line(&mut self) {
        while !matches!(self.peek(), None | Some(Tok::Newline)) {
            self.pos += 1;
        }
    }
    fn skip_seps(&mut self) {
        while matches!(self.peek(), Some(Tok::Newline | Tok::Colon)) {
            self.pos += 1;
        }
    }
    fn end_stmt(&mut self) -> PResult<()> {
        if self.at_stmt_end() { Ok(()) } else { self.err("Expected: end of statement") }
    }
    fn at_terminator(&self, t: Terminator) -> bool {
        t.iter().enumerate().all(|(i, kw)| self.is_kw_at(i, kw))
    }
    fn enter(&mut self) -> PResult<()> {
        self.depth += 1;
        if self.depth > MAX_DEPTH { self.err("Out of memory (nested too deeply)") } else { Ok(()) }
    }

    // ------------------------------------------------------------ module level

    fn module(&mut self, name: &str) -> PResult<ModuleAst> {
        let mut m = ModuleAst { name: name.to_string(), ..Default::default() };
        loop {
            self.skip_seps();
            if self.peek().is_none() {
                break;
            }
            if self.eat_kw("attribute") || self.is_kw("implements") || self.is_kw("event") || self.is_kw("declare") || self.is_kw_at(1, "declare") {
                // Attribute lines, interfaces, events and Windows API declarations are skipped.
                self.skip_line();
                continue;
            }
            if self.eat_kw("option") {
                if self.eat_kw("explicit") {
                    m.option_explicit = true;
                } else if self.eat_kw("base") {
                    m.option_base = match self.bump() {
                        Some(Tok::Num(n, _)) if n == 0.0 || n == 1.0 => n as i64,
                        _ => return self.err("Expected: 0 or 1"),
                    };
                } else if self.eat_kw("compare") {
                    m.compare_text = self.eat_kw("text");
                    let _ = self.eat_kw("binary") || self.eat_kw("database");
                } else if self.eat_kw("private") {
                    self.expect_kw("module")?;
                }
                self.end_stmt()?;
                continue;
            }
            let mut public = true;
            let mut scoped = false;
            if self.eat_kw("public") || self.eat_kw("global") {
                scoped = true;
            } else if self.eat_kw("private") {
                public = false;
                scoped = true;
            } else if self.eat_kw("friend") {
                scoped = true;
            }
            let _ = self.eat_kw("static");
            if self.is_kw("sub") || self.is_kw("function") || self.is_kw("property") {
                let proc = self.procedure(public)?;
                m.procs.push(proc);
                continue;
            }
            if self.eat_kw("enum") {
                self.enumeration(&mut m, public)?;
                continue;
            }
            if self.eat_kw("type") {
                // User-defined types aren't supported yet: their blocks are skipped.
                while self.peek().is_some() && !self.at_terminator(&["end", "type"]) {
                    self.skip_line();
                    self.skip_seps();
                }
                self.pos += 2;
                continue;
            }
            if self.eat_kw("const") {
                for (n, e) in self.const_list()? {
                    m.consts.push((n, e, public && scoped));
                }
                self.end_stmt()?;
                continue;
            }
            if self.eat_kw("dim") || scoped {
                // Module variables: Dim and Private are private, Public and Global public.
                let module_public = scoped && public;
                for d in self.var_list()? {
                    m.vars.push((d, module_public));
                }
                self.end_stmt()?;
                continue;
            }
            return self.err("Invalid outside procedure");
        }
        Ok(m)
    }

    fn enumeration(&mut self, m: &mut ModuleAst, public: bool) -> PResult<()> {
        let _enum_name = self.ident()?;
        let mut prev: Option<String> = None;
        loop {
            self.skip_seps();
            if self.peek().is_none() {
                return self.err("Expected: End Enum");
            }
            if self.at_terminator(&["end", "enum"]) {
                self.pos += 2;
                return Ok(());
            }
            let name = self.ident()?;
            let value = if self.eat(&Tok::Op("=")) {
                self.expr()?
            } else {
                match &prev {
                    Some(p) => Expr::Binary(BinOp::Add, Box::new(Expr::Name(p.clone())), Box::new(Expr::Num(1.0, true))),
                    None => Expr::Num(0.0, true),
                }
            };
            m.consts.push((name.clone(), value, public));
            prev = Some(name);
            self.end_stmt()?;
        }
    }

    fn type_name(&mut self) -> PResult<Option<String>> {
        if !self.eat_kw("as") {
            return Ok(None);
        }
        let _ = self.eat_kw("new");
        let mut t = self.ident()?;
        // `As Excel.Range`, `As String * 10`.
        while self.eat(&Tok::Dot) {
            t = self.ident()?;
        }
        if self.eat(&Tok::Op("*")) {
            let _ = self.bump();
        }
        Ok(Some(t))
    }

    fn bounds(&mut self) -> PResult<Vec<(Option<Expr>, Expr)>> {
        // After `(`.
        let mut dims = Vec::new();
        if self.eat(&Tok::RParen) {
            return Ok(dims);
        }
        loop {
            let a = self.expr()?;
            if self.eat_kw("to") {
                let b = self.expr()?;
                dims.push((Some(a), b));
            } else {
                dims.push((None, a));
            }
            if self.eat(&Tok::RParen) {
                break;
            }
            self.expect(&Tok::Comma, ", or )")?;
        }
        if dims.len() > 60 {
            return self.err("Too many dimensions");
        }
        Ok(dims)
    }

    fn var_list(&mut self) -> PResult<Vec<VarDecl>> {
        let mut out = Vec::new();
        loop {
            let _ = self.eat_kw("withevents");
            let name = self.ident()?;
            let dims = if self.eat(&Tok::LParen) { Some(self.bounds()?) } else { None };
            let new_object = self.is_kw("as") && self.is_kw_at(1, "new");
            let ty = self.type_name()?;
            out.push(VarDecl { name, dims, ty, new_object });
            if !self.eat(&Tok::Comma) {
                break;
            }
        }
        Ok(out)
    }

    fn const_list(&mut self) -> PResult<Vec<(String, Expr)>> {
        let mut out = Vec::new();
        loop {
            let name = self.ident()?;
            let _ = self.type_name()?;
            self.expect(&Tok::Op("="), "=")?;
            out.push((name, self.expr()?));
            if !self.eat(&Tok::Comma) {
                break;
            }
        }
        Ok(out)
    }

    fn procedure(&mut self, public: bool) -> PResult<Proc> {
        let line = self.line();
        let (kind, end) = if self.eat_kw("sub") {
            (ProcKind::Sub, END_SUB)
        } else if self.eat_kw("function") {
            (ProcKind::Function, END_FUNCTION)
        } else {
            self.expect_kw("property")?;
            let k = if self.eat_kw("get") {
                ProcKind::PropertyGet
            } else if self.eat_kw("let") {
                ProcKind::PropertyLet
            } else {
                self.expect_kw("set")?;
                ProcKind::PropertySet
            };
            (k, END_PROPERTY)
        };
        let name = self.ident()?;
        let mut params = Vec::new();
        if self.eat(&Tok::LParen) && !self.eat(&Tok::RParen) {
            loop {
                let optional = self.eat_kw("optional");
                let by_val = if self.eat_kw("byval") {
                    true
                } else {
                    let _ = self.eat_kw("byref");
                    false
                };
                let param_array = self.eat_kw("paramarray");
                let pname = self.ident()?;
                if self.eat(&Tok::LParen) {
                    self.expect(&Tok::RParen, ")")?;
                }
                let ty = self.type_name()?;
                let default = if self.eat(&Tok::Op("=")) { Some(self.expr()?) } else { None };
                params.push(Param { name: pname, by_val, optional: optional || param_array, default, param_array, ty });
                if self.eat(&Tok::RParen) {
                    break;
                }
                self.expect(&Tok::Comma, ", or )")?;
            }
        }
        let ret_ty = self.type_name()?;
        self.end_stmt()?;
        let (body, _) = self.block(&[end])?;
        self.pos += end.len();
        Ok(Proc { name, kind, params, body, public, ret_ty, line })
    }

    // ------------------------------------------------------------ statements

    /// Statements up to one of `ends` (not consumed); which one was reached.
    fn block(&mut self, ends: &[Terminator]) -> PResult<(Vec<Line>, Terminator)> {
        self.enter()?;
        let mut out = Vec::new();
        loop {
            self.skip_seps();
            if self.peek().is_none() {
                let want = ends.first().map(|t| t.join(" ")).unwrap_or_default();
                return self.err(format!("Expected: {}", title(&want)));
            }
            if let Some(t) = ends.iter().find(|t| self.at_terminator(t)) {
                self.depth -= 1;
                return Ok((out, t));
            }
            // A label at the start of a line: `name:` or a line number.
            let at_line_start = self.pos == 0 || matches!(self.toks.get(self.pos.wrapping_sub(1)).map(|t| &t.tok), Some(Tok::Newline));
            if at_line_start {
                match (self.peek(), self.peek_at(1)) {
                    (Some(Tok::Ident(n)), Some(Tok::Colon)) if !is_reserved(n) => {
                        let line = self.line();
                        let n = n.clone();
                        self.pos += 2;
                        out.push(Line { line, stmt: Stmt::Label(n) });
                        continue;
                    }
                    (Some(Tok::Num(n, true)), _) => {
                        let line = self.line();
                        let n = format!("{n}");
                        self.pos += 1;
                        let _ = self.eat(&Tok::Colon);
                        out.push(Line { line, stmt: Stmt::Label(n) });
                        continue;
                    }
                    _ => {}
                }
            }
            let line = self.line();
            if let Some(stmt) = self.statement()? {
                out.push(Line { line, stmt });
            }
            if !self.at_stmt_end() {
                return self.err("Expected: end of statement");
            }
        }
    }

    /// One statement on the rest of the line (single-line `If`): up to `Else` or the line's end.
    fn inline_block(&mut self) -> PResult<Vec<Line>> {
        let mut out = Vec::new();
        loop {
            while self.eat(&Tok::Colon) {}
            if matches!(self.peek(), None | Some(Tok::Newline)) || self.is_kw("else") {
                return Ok(out);
            }
            let line = self.line();
            if let Some(stmt) = self.statement()? {
                out.push(Line { line, stmt });
            }
            if !self.at_stmt_end() {
                return self.err("Expected: end of statement");
            }
        }
    }

    fn statement(&mut self) -> PResult<Option<Stmt>> {
        let Some(Tok::Ident(word)) = self.peek().cloned() else {
            if self.peek() == Some(&Tok::Dot) {
                return self.expr_statement().map(Some);
            }
            return self.err("Expected: statement");
        };
        let w = word.to_ascii_lowercase();
        Ok(Some(match w.as_str() {
            "dim" | "static" => {
                self.pos += 1;
                Stmt::Dim(self.var_list()?)
            }
            "redim" => {
                self.pos += 1;
                let preserve = self.eat_kw("preserve");
                let mut vars = Vec::new();
                loop {
                    let name = self.ident()?;
                    self.expect(&Tok::LParen, "(")?;
                    let dims = self.bounds()?;
                    let _ = self.type_name()?;
                    vars.push((name, dims));
                    if !self.eat(&Tok::Comma) {
                        break;
                    }
                }
                Stmt::ReDim { preserve, vars }
            }
            "const" => {
                self.pos += 1;
                Stmt::Const(self.const_list()?)
            }
            "set" => {
                self.pos += 1;
                let target = self.postfix()?;
                self.expect(&Tok::Op("="), "=")?;
                Stmt::Assign { target, value: self.expr()?, set: true }
            }
            "let" => {
                self.pos += 1;
                let target = self.postfix()?;
                self.expect(&Tok::Op("="), "=")?;
                Stmt::Assign { target, value: self.expr()?, set: false }
            }
            "call" => {
                self.pos += 1;
                let target = self.postfix()?;
                match target {
                    Expr::Call(f, args) => Stmt::Call(*f, args),
                    other => Stmt::Call(other, Vec::new()),
                }
            }
            "if" => return self.if_statement().map(Some),
            "select" => return self.select_statement().map(Some),
            "for" => return self.for_statement().map(Some),
            "do" => {
                self.pos += 1;
                let pre = self.loop_condition()?;
                let (body, _) = self.block(&[LOOP])?;
                self.pos += 1;
                let post = self.loop_condition()?;
                if pre.is_some() && post.is_some() {
                    return self.err("Expected: end of statement (a Do loop has one condition)");
                }
                Stmt::Do { pre, post, body }
            }
            "while" => {
                self.pos += 1;
                let cond = self.expr()?;
                let (body, _) = self.block(&[WEND])?;
                self.pos += 1;
                Stmt::While { cond, body }
            }
            "exit" => {
                self.pos += 1;
                let k = if self.eat_kw("sub") {
                    ExitKind::Sub
                } else if self.eat_kw("function") {
                    ExitKind::Function
                } else if self.eat_kw("property") {
                    ExitKind::Property
                } else if self.eat_kw("for") {
                    ExitKind::For
                } else if self.eat_kw("do") {
                    ExitKind::Do
                } else {
                    return self.err("Expected: Do or For or Sub or Function or Property");
                };
                Stmt::Exit(k)
            }
            "with" => {
                self.pos += 1;
                let object = self.expr()?;
                let (body, _) = self.block(&[END_WITH])?;
                self.pos += 2;
                Stmt::With { object, body }
            }
            "on" if self.is_kw_at(1, "error") => {
                self.pos += 2;
                if self.eat_kw("resume") {
                    self.expect_kw("next")?;
                    Stmt::OnError(OnError::ResumeNext)
                } else {
                    self.expect_kw("goto")?;
                    match self.bump() {
                        Some(Tok::Num(0.0, _)) => Stmt::OnError(OnError::GoToZero),
                        Some(Tok::Op("-")) => {
                            let _ = self.bump();
                            Stmt::OnError(OnError::GoToZero)
                        }
                        Some(Tok::Num(n, _)) => Stmt::OnError(OnError::GoTo(format!("{n}"))),
                        Some(Tok::Ident(l)) => Stmt::OnError(OnError::GoTo(l)),
                        _ => return self.err("Expected: label"),
                    }
                }
            }
            "goto" => {
                self.pos += 1;
                match self.bump() {
                    Some(Tok::Ident(l)) => Stmt::GoTo(l),
                    Some(Tok::Num(n, _)) => Stmt::GoTo(format!("{n}")),
                    _ => return self.err("Expected: label"),
                }
            }
            "resume" => {
                self.pos += 1;
                if self.eat_kw("next") {
                    Stmt::ResumeNext
                } else {
                    match self.peek().cloned() {
                        Some(Tok::Ident(l)) => {
                            self.pos += 1;
                            Stmt::ResumeLabel(l)
                        }
                        Some(Tok::Num(n, _)) if n != 0.0 => {
                            self.pos += 1;
                            Stmt::ResumeLabel(format!("{n}"))
                        }
                        _ => return self.err("Resume without a label isn't supported yet (use Resume Next or Resume label)"),
                    }
                }
            }
            "erase" => {
                self.pos += 1;
                let mut names = vec![self.ident()?];
                while self.eat(&Tok::Comma) {
                    names.push(self.ident()?);
                }
                Stmt::Erase(names)
            }
            "debug" if self.peek_at(1) == Some(&Tok::Dot) && self.is_kw_at(2, "print") => {
                self.pos += 3;
                let mut items = Vec::new();
                while !self.at_stmt_end() {
                    if self.eat(&Tok::Semi) || self.eat(&Tok::Comma) {
                        continue;
                    }
                    items.push(self.expr()?);
                }
                Stmt::Print(items)
            }
            "end" if matches!(self.peek_at(1), None | Some(Tok::Newline | Tok::Colon)) => {
                self.pos += 1;
                Stmt::End
            }
            "stop" => {
                self.pos += 1;
                Stmt::End
            }
            "option" | "attribute" => {
                self.skip_line();
                return Ok(None);
            }
            "else" | "elseif" | "wend" | "loop" | "next" | "case" => return self.err(format!("{} without a block", title(&w))),
            "end" => return self.err(format!("End {} without a block", self.peek_at(1).map(tok_text).unwrap_or_default())),
            _ => self.expr_statement()?,
        }))
    }

    fn loop_condition(&mut self) -> PResult<Option<(bool, Expr)>> {
        if self.eat_kw("while") {
            Ok(Some((false, self.expr()?)))
        } else if self.eat_kw("until") {
            Ok(Some((true, self.expr()?)))
        } else {
            Ok(None)
        }
    }

    fn if_statement(&mut self) -> PResult<Stmt> {
        self.pos += 1;
        let cond = self.expr()?;
        self.expect_kw("then")?;
        if !matches!(self.peek(), None | Some(Tok::Newline)) {
            // Single-line If.
            let then = self.inline_block()?;
            let otherwise = if self.eat_kw("else") { Some(self.inline_block()?) } else { None };
            return Ok(Stmt::If { branches: vec![(cond, then)], otherwise });
        }
        let mut branches = Vec::new();
        let mut otherwise = None;
        let mut cond = cond;
        loop {
            let (body, end) = self.block(&[END_IF, ELSEIF, ELSE])?;
            branches.push((cond, body));
            if end == ELSEIF {
                self.pos += 1;
                cond = self.expr()?;
                self.expect_kw("then")?;
                continue;
            }
            if end == ELSE {
                self.pos += 1;
                let (body, _) = self.block(&[END_IF])?;
                otherwise = Some(body);
            }
            self.pos += 2;
            return Ok(Stmt::If { branches, otherwise });
        }
    }

    fn select_statement(&mut self) -> PResult<Stmt> {
        self.pos += 1;
        self.expect_kw("case")?;
        let expr = self.expr()?;
        self.end_stmt()?;
        let mut cases = Vec::new();
        let mut otherwise = None;
        self.skip_seps();
        loop {
            if self.at_terminator(END_SELECT) {
                self.pos += 2;
                break;
            }
            self.expect_kw("case")?;
            if self.eat_kw("else") {
                let (body, _) = self.block(&[END_SELECT])?;
                otherwise = Some(body);
                continue;
            }
            let mut tests = Vec::new();
            loop {
                if self.eat_kw("is") {
                    let op = match self.bump() {
                        Some(Tok::Op(o)) => cmp_op(o).ok_or((self.line(), "Expected: comparison operator".to_string()))?,
                        _ => return self.err("Expected: comparison operator"),
                    };
                    tests.push(CaseTest::Is(op, self.expr()?));
                } else {
                    let a = self.expr()?;
                    if self.eat_kw("to") {
                        tests.push(CaseTest::Range(a, self.expr()?));
                    } else {
                        tests.push(CaseTest::Value(a));
                    }
                }
                if !self.eat(&Tok::Comma) {
                    break;
                }
            }
            let (body, _) = self.block(&[CASE, END_SELECT])?;
            cases.push((tests, body));
        }
        Ok(Stmt::Select { expr, cases, otherwise })
    }

    fn for_statement(&mut self) -> PResult<Stmt> {
        self.pos += 1;
        if self.eat_kw("each") {
            let var = self.postfix()?;
            self.expect_kw("in")?;
            let collection = self.expr()?;
            let body = self.for_body()?;
            return Ok(Stmt::ForEach { var, collection, body });
        }
        let var = self.postfix()?;
        self.expect(&Tok::Op("="), "=")?;
        let from = self.expr()?;
        self.expect_kw("to")?;
        let to = self.expr()?;
        let step = if self.eat_kw("step") { Some(self.expr()?) } else { None };
        let body = self.for_body()?;
        Ok(Stmt::For { var, from, to, step, body })
    }

    fn for_body(&mut self) -> PResult<Vec<Line>> {
        let (body, _) = self.block(&[NEXT])?;
        self.pos += 1;
        // `Next i` (the variable is optional and not checked).
        if matches!(self.peek(), Some(Tok::Ident(_))) && !self.is_kw("else") {
            self.pos += 1;
        }
        Ok(body)
    }

    fn expr_statement(&mut self) -> PResult<Stmt> {
        let target = self.postfix()?;
        if self.eat(&Tok::Op("=")) {
            return Ok(Stmt::Assign { target, value: self.expr()?, set: false });
        }
        if self.at_stmt_end() {
            return Ok(match target {
                Expr::Call(f, args) => Stmt::Call(*f, args),
                other => Stmt::Call(other, Vec::new()),
            });
        }
        // `Foo a, b` (and `Foo (a), b`: the first argument was parsed with the name).
        let (callee, mut args) = match target {
            Expr::Call(f, args) if self.peek() == Some(&Tok::Comma) => (*f, args),
            other => (other, Vec::new()),
        };
        if self.peek() == Some(&Tok::Comma) {
            self.pos += 1;
        }
        loop {
            if self.at_stmt_end() {
                break;
            }
            if self.peek() == Some(&Tok::Comma) {
                self.pos += 1;
                args.push(Arg { name: None, value: None });
                continue;
            }
            args.push(self.arg()?);
            if !self.eat(&Tok::Comma) {
                break;
            }
            if self.at_stmt_end() {
                args.push(Arg { name: None, value: None });
            }
        }
        Ok(Stmt::Call(callee, args))
    }

    fn arg(&mut self) -> PResult<Arg> {
        if let (Some(Tok::Ident(n)), Some(Tok::Assign)) = (self.peek(), self.peek_at(1)) {
            let name = n.clone();
            self.pos += 2;
            return Ok(Arg { name: Some(name), value: Some(self.expr()?) });
        }
        let _ = self.eat_kw("byval");
        Ok(Arg { name: None, value: Some(self.expr()?) })
    }

    /// After `(`: the arguments up to `)`.
    fn paren_args(&mut self) -> PResult<Vec<Arg>> {
        let mut args = Vec::new();
        if self.eat(&Tok::RParen) {
            return Ok(args);
        }
        loop {
            if matches!(self.peek(), Some(Tok::Comma | Tok::RParen)) {
                args.push(Arg { name: None, value: None });
            } else {
                args.push(self.arg()?);
            }
            if self.eat(&Tok::RParen) {
                break;
            }
            self.expect(&Tok::Comma, ", or )")?;
        }
        if args.len() > 255 {
            return self.err("Too many arguments");
        }
        Ok(args)
    }

    // ------------------------------------------------------------ expressions

    pub fn expr(&mut self) -> PResult<Expr> {
        self.enter()?;
        let e = self.binary(0);
        self.depth -= 1;
        e
    }

    /// Precedence climbing over [MS-VBAL] §5.6.9's operator levels (lowest first).
    fn binary(&mut self, level: usize) -> PResult<Expr> {
        const LEVELS: usize = 10;
        if level >= LEVELS {
            return self.unary();
        }
        if level == 5 && self.eat_kw("not") {
            // `Not` binds looser than comparisons: Not a = b is Not (a = b).
            self.enter()?;
            let e = self.binary(level)?;
            self.depth -= 1;
            return Ok(Expr::Unary(UnOp::Not, Box::new(e)));
        }
        let mut lhs = self.binary(level + 1)?;
        loop {
            let op = match (level, self.peek()) {
                (0, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("imp") => BinOp::Imp,
                (1, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("eqv") => BinOp::Eqv,
                (2, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("xor") => BinOp::Xor,
                (3, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("or") => BinOp::Or,
                (4, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("and") => BinOp::And,
                (6, Some(Tok::Op(o))) if cmp_op(o).is_some() => cmp_op(o).unwrap_or(BinOp::Eq),
                (6, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("like") => BinOp::Like,
                (6, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("is") => BinOp::Is,
                (7, Some(Tok::Op("&"))) => BinOp::Concat,
                (8, Some(Tok::Op("+"))) => BinOp::Add,
                (8, Some(Tok::Op("-"))) => BinOp::Sub,
                (9, Some(Tok::Ident(s))) if s.eq_ignore_ascii_case("mod") => BinOp::Mod,
                _ => break,
            };
            self.pos += 1;
            let rhs = self.binary(level + 1)?;
            lhs = Expr::Binary(op, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    /// `\`, then `* /`, then unary minus, then `^`.
    fn unary(&mut self) -> PResult<Expr> {
        let mut lhs = self.mul()?;
        while self.eat(&Tok::Op("\\")) {
            let rhs = self.mul()?;
            lhs = Expr::Binary(BinOp::IntDiv, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn mul(&mut self) -> PResult<Expr> {
        let mut lhs = self.neg()?;
        loop {
            let op = match self.peek() {
                Some(Tok::Op("*")) => BinOp::Mul,
                Some(Tok::Op("/")) => BinOp::Div,
                _ => break,
            };
            self.pos += 1;
            let rhs = self.neg()?;
            lhs = Expr::Binary(op, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn neg(&mut self) -> PResult<Expr> {
        if self.eat(&Tok::Op("-")) {
            self.enter()?;
            let e = self.neg()?;
            self.depth -= 1;
            return Ok(Expr::Unary(UnOp::Neg, Box::new(e)));
        }
        if self.eat(&Tok::Op("+")) {
            return self.neg();
        }
        self.pow()
    }

    fn pow(&mut self) -> PResult<Expr> {
        let mut lhs = self.postfix()?;
        while self.eat(&Tok::Op("^")) {
            // The exponent may be negated: 2 ^ -1.
            let rhs = if self.eat(&Tok::Op("-")) { Expr::Unary(UnOp::Neg, Box::new(self.postfix()?)) } else { self.postfix()? };
            lhs = Expr::Binary(BinOp::Pow, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn postfix(&mut self) -> PResult<Expr> {
        let mut e = self.primary()?;
        loop {
            match self.peek() {
                Some(Tok::Dot) => {
                    self.pos += 1;
                    let m = self.ident()?;
                    e = Expr::Member(Box::new(e), m);
                }
                Some(Tok::Bang) => {
                    // `coll!key` is coll("key").
                    self.pos += 1;
                    let key = self.ident()?;
                    e = Expr::Call(Box::new(e), vec![Arg { name: None, value: Some(Expr::Str(key)) }]);
                }
                Some(Tok::LParen) => {
                    self.pos += 1;
                    let args = self.paren_args()?;
                    e = Expr::Call(Box::new(e), args);
                }
                _ => return Ok(e),
            }
        }
    }

    fn primary(&mut self) -> PResult<Expr> {
        match self.bump() {
            Some(Tok::Num(n, int)) => Ok(Expr::Num(n, int)),
            Some(Tok::Str(s)) => Ok(Expr::Str(s)),
            Some(Tok::LParen) => {
                let e = self.expr()?;
                self.expect(&Tok::RParen, ")")?;
                Ok(e)
            }
            Some(Tok::Dot) => Ok(Expr::WithMember(self.ident()?)),
            Some(Tok::Ident(s)) => Ok(match s.to_ascii_lowercase().as_str() {
                "true" => Expr::Bool(true),
                "false" => Expr::Bool(false),
                "nothing" => Expr::Nothing,
                "empty" => Expr::Empty,
                "null" => Expr::Null,
                "new" => {
                    let mut t = self.ident()?;
                    while self.eat(&Tok::Dot) {
                        t = self.ident()?;
                    }
                    Expr::New(t)
                }
                w if is_reserved(w) => {
                    self.pos -= 1;
                    return self.err(format!("Expected: expression (found `{s}`)"));
                }
                _ => Expr::Name(s),
            }),
            Some(t) => {
                self.pos -= 1;
                self.err(format!("Expected: expression (found `{}`)", tok_text(&t)))
            }
            None => self.err("Expected: expression"),
        }
    }
}

fn cmp_op(o: &str) -> Option<BinOp> {
    Some(match o {
        "=" => BinOp::Eq,
        "<>" => BinOp::Ne,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::Le,
        ">=" => BinOp::Ge,
        _ => return None,
    })
}

/// Words that can't be a name in an expression.
fn is_reserved(w: &str) -> bool {
    matches!(
        w.to_ascii_lowercase().as_str(),
        "then"
            | "else"
            | "elseif"
            | "end"
            | "to"
            | "step"
            | "next"
            | "loop"
            | "wend"
            | "and"
            | "or"
            | "not"
            | "xor"
            | "eqv"
            | "imp"
            | "mod"
            | "is"
            | "like"
            | "dim"
            | "if"
            | "for"
            | "do"
            | "while"
            | "sub"
            | "function"
            | "case"
            | "select"
            | "with"
            | "goto"
            | "exit"
            | "set"
            | "let"
            | "call"
    )
}

fn tok_text(t: &Tok) -> String {
    match t {
        Tok::Ident(s) => s.clone(),
        Tok::Num(n, _) => format!("{n}"),
        Tok::Str(s) => format!("\"{s}\""),
        Tok::Op(o) => (*o).to_string(),
        Tok::LParen => "(".into(),
        Tok::RParen => ")".into(),
        Tok::Comma => ",".into(),
        Tok::Dot => ".".into(),
        Tok::Bang => "!".into(),
        Tok::Semi => ";".into(),
        Tok::Assign => ":=".into(),
        Tok::Colon => ":".into(),
        Tok::Newline => "end of line".into(),
    }
}

/// `end if` → `End If`.
fn title(s: &str) -> String {
    s.split(' ')
        .map(|w| {
            let mut c = w.chars();
            match c.next() {
                Some(f) => f.to_uppercase().chain(c).collect::<String>(),
                None => String::new(),
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn module(src: &str) -> ModuleAst {
        parse_module("Module1", src).unwrap_or_else(|e| panic!("{e:?}"))
    }

    #[test]
    fn procedures_and_statements_parse() {
        let m = module(
            "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\nPrivate count As Long\r\nPublic Const Rate = 0.2\r\n\r\n\
             Sub Hello()\r\n  Dim i As Integer, s$\r\n  For i = 1 To 10 Step 2\r\n    If i > 5 Then Exit For Else s = s & i\r\n  Next i\r\n  \
             With Range(\"A1\")\r\n    .Value = \"x\"\r\n    .Font.Bold = True\r\n  End With\r\n  MsgBox \"done\", vbInformation, Title:=\"T\"\r\nEnd Sub\r\n\
             Function Twice(ByVal x As Double, Optional y = 1) As Double\r\n  Select Case x\r\n    Case 1, 2: Twice = 0\r\n    Case 3 To 5\r\n      Twice = x\r\n    \
             Case Is > 10\r\n      Twice = -x ^ 2\r\n    Case Else\r\n      Twice = x * 2\r\n  End Select\r\nEnd Function\r\n",
        );
        assert!(m.option_explicit);
        assert_eq!(m.vars.len(), 1);
        assert_eq!(m.consts.len(), 1);
        assert_eq!(m.procs.len(), 2);
        let hello = &m.procs[0];
        assert_eq!(hello.name, "Hello");
        assert_eq!(hello.body.len(), 4);
        assert!(matches!(&hello.body[3].stmt, Stmt::Call(Expr::Name(n), args) if n == "MsgBox" && args.len() == 3 && args[2].name.as_deref() == Some("Title")));
        let twice = &m.procs[1];
        assert_eq!(twice.params.len(), 2);
        assert!(twice.params[0].by_val && twice.params[1].optional);
        let Stmt::Select { cases, otherwise, .. } = &twice.body[0].stmt else { panic!() };
        assert_eq!(cases.len(), 3);
        assert!(otherwise.is_some());
        // -x ^ 2 is -(x ^ 2).
        let Stmt::Assign { value, .. } = &cases[2].1[0].stmt else { panic!() };
        assert!(matches!(value, Expr::Unary(UnOp::Neg, inner) if matches!(**inner, Expr::Binary(BinOp::Pow, ..))));
    }

    #[test]
    fn loops_errors_and_labels() {
        let m = module(
            "Sub A()\r\n  On Error GoTo Fail\r\n  Do While x < 3\r\n    x = x + 1\r\n  Loop\r\n  Do\r\n  Loop Until x > 9\r\n  While False\r\n  Wend\r\n  \
             Exit Sub\r\nFail:\r\n  Debug.Print \"error\"; Err.Number\r\n  Resume Next\r\nEnd Sub",
        );
        let body = &m.procs[0].body;
        assert!(matches!(body[0].stmt, Stmt::OnError(OnError::GoTo(_))));
        assert!(matches!(&body[1].stmt, Stmt::Do { pre: Some((false, _)), post: None, .. }));
        assert!(matches!(&body[2].stmt, Stmt::Do { pre: None, post: Some((true, _)), .. }));
        assert!(matches!(&body[5].stmt, Stmt::Label(l) if l == "Fail"));
        assert!(matches!(&body[6].stmt, Stmt::Print(items) if items.len() == 2));
        assert_eq!(body[6].line, 12);
    }

    #[test]
    fn compile_errors_name_the_line() {
        for (src, line) in
            [("Sub A()\r\n  If x Then\r\n    y = 1\r\nEnd Sub", 4), ("Sub A()\r\n  x = (1 + \r\nEnd Sub", 2), ("Sub A()\r\n  Next\r\nEnd Sub", 2), ("x = 1", 1)]
        {
            let e = parse_module("M", src).unwrap_err();
            assert_eq!(e.0, line, "{src:?}: {e:?}");
        }
        let deep = format!("Sub A()\r\n x = {}1{}\r\nEnd Sub", "(".repeat(500), ")".repeat(500));
        let r = std::thread::Builder::new().stack_size(crate::STACK).spawn(move || parse_module("M", &deep).is_err()).unwrap().join().unwrap();
        assert!(r, "nesting is bounded");
    }

    #[test]
    fn enums_and_declares() {
        let m = module(
            "Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\r\nEnum Color\r\n  Red\r\n  Green = 5\r\n  Blue\r\nEnd Enum\r\nSub A()\r\nEnd Sub",
        );
        assert_eq!(m.consts.len(), 3);
        assert_eq!(m.procs.len(), 1);
    }
}

#[cfg(test)]
mod hostile {
    /// Junk built from VBA's own pieces: every result is a module or a compile error.
    #[test]
    fn junk_never_panics() {
        let pieces = [
            "Sub", "End", "If", "Then", "Else", "For", "Next", "Do", "Loop", "(", ")", ",", ".", ":", "=", "\"x\"", "\r\n", "1", "a", "_", "'", "With",
            "Select", "Case", "Is", "To", "Dim", "As", "Long", "&H", "#", "!", ":=", "Function", "Property", "Get", "Enum", "Type",
        ];
        let mut seed = 0x1234_5678_u64;
        for _ in 0..3000 {
            let mut src = String::new();
            for _ in 0..(seed % 40) {
                seed ^= seed << 13;
                seed ^= seed >> 7;
                seed ^= seed << 17;
                src.push_str(pieces[(seed % pieces.len() as u64) as usize]);
                src.push(' ');
            }
            let _ = super::parse_module("M", &src);
        }
    }
}
