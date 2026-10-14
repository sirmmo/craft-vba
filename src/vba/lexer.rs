//! Tokens of VBA source ([MS-VBAL] §3): identifiers (keywords are identifiers the parser knows),
//! numbers, strings, operators and statement separators. Line continuations (` _` at a line's
//! end) join lines; comments (`'…`, `Rem …`) are dropped.

#[derive(Clone, Debug, PartialEq)]
pub enum Tok {
    Ident(String),
    /// A number and whether it was written as an integer (no `.`, exponent or `#`/`!` suffix).
    Num(f64, bool),
    Str(String),
    /// `+ - * / \ ^ & = < > <= >= <>`
    Op(&'static str),
    LParen,
    RParen,
    Comma,
    Dot,
    Bang,
    Semi,
    /// `:=` (named argument).
    Assign,
    /// `:` (statement separator).
    Colon,
    Newline,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Token {
    pub tok: Tok,
    /// The physical line (1-based) the token starts on.
    pub line: u32,
}

/// Longest source accepted, in bytes.
pub const MAX_SOURCE: usize = 16 << 20;

pub fn lex(src: &str) -> Result<Vec<Token>, (u32, String)> {
    if src.len() > MAX_SOURCE {
        return Err((1, "the module is too long".into()));
    }
    let chars: Vec<char> = src.chars().collect();
    let mut out: Vec<Token> = Vec::new();
    let mut i = 0usize;
    let mut line = 1u32;
    let at = |i: usize| chars.get(i).copied();
    let push = |out: &mut Vec<Token>, tok: Tok, line: u32| out.push(Token { tok, line });
    while let Some(c) = at(i) {
        // Statement start: `Rem` comments.
        match c {
            '\r' => {
                i += 1;
            }
            '\n' => {
                push(&mut out, Tok::Newline, line);
                line += 1;
                i += 1;
            }
            ' ' | '\t' | '\u{a0}' => i += 1,
            '\'' => {
                while at(i).is_some_and(|c| c != '\n') {
                    i += 1;
                }
            }
            '_' if at(i + 1).is_none_or(|n| n == ' ' || n == '\t' || n == '\r' || n == '\n') && i > 0 && at(i - 1).is_some_and(|p| p == ' ' || p == '\t') => {
                // A line continuation: skip to the next line.
                let mut j = i + 1;
                while at(j).is_some_and(|c| c == ' ' || c == '\t' || c == '\r') {
                    j += 1;
                }
                if at(j) == Some('\n') {
                    line += 1;
                    i = j + 1;
                } else {
                    return Err((line, "Expected: line continuation at the end of the line".into()));
                }
            }
            '"' => {
                let start_line = line;
                let mut s = String::new();
                i += 1;
                loop {
                    match at(i) {
                        Some('"') if at(i + 1) == Some('"') => {
                            s.push('"');
                            i += 2;
                        }
                        Some('"') => {
                            i += 1;
                            break;
                        }
                        Some('\n') | None => return Err((start_line, "Expected: \" (a string isn't closed)".into())),
                        Some(c) => {
                            s.push(c);
                            i += 1;
                        }
                    }
                }
                push(&mut out, Tok::Str(s), start_line);
            }
            '&' if at(i + 1).is_some_and(|n| matches!(n, 'h' | 'H' | 'o' | 'O')) && at(i + 2).is_some_and(|n| n.is_ascii_alphanumeric()) => {
                let radix = if matches!(at(i + 1), Some('h' | 'H')) { 16 } else { 8 };
                let mut j = i + 2;
                let mut digits = String::new();
                while let Some(d) = at(j).filter(|d| d.is_digit(radix)) {
                    digits.push(d);
                    j += 1;
                }
                // A type suffix (`&H10&`).
                if at(j).is_some_and(|s| matches!(s, '&' | '%' | '^')) {
                    j += 1;
                }
                let v = u64::from_str_radix(&digits, radix).map_err(|_| (line, "Expected: a number".to_string()))?;
                // &HFFFF is Integer -1; &HFFFFFFFF is Long -1 (two's complement), as in VBA.
                let v = if digits.len() <= 4 && radix == 16 {
                    f64::from(v as u16 as i16)
                } else if v <= 0xFFFF_FFFF {
                    f64::from(v as u32 as i32)
                } else {
                    v as f64
                };
                push(&mut out, Tok::Num(v, true), line);
                i = j;
            }
            c if c.is_ascii_digit() || (c == '.' && at(i + 1).is_some_and(|d| d.is_ascii_digit())) => {
                let mut j = i;
                let mut text = String::new();
                let mut integer = true;
                while let Some(d) = at(j).filter(|d| d.is_ascii_digit()) {
                    text.push(d);
                    j += 1;
                }
                if at(j) == Some('.') && at(j + 1).is_some_and(|d| d.is_ascii_digit()) || (at(j) == Some('.') && !text.is_empty()) {
                    integer = false;
                    text.push('.');
                    j += 1;
                    while let Some(d) = at(j).filter(|d| d.is_ascii_digit()) {
                        text.push(d);
                        j += 1;
                    }
                }
                if at(j).is_some_and(|e| e == 'e' || e == 'E') && at(j + 1).is_some_and(|d| d.is_ascii_digit() || d == '+' || d == '-') {
                    integer = false;
                    text.push('e');
                    j += 1;
                    if let Some(sign) = at(j).filter(|s| *s == '+' || *s == '-') {
                        text.push(sign);
                        j += 1;
                    }
                    while let Some(d) = at(j).filter(|d| d.is_ascii_digit()) {
                        text.push(d);
                        j += 1;
                    }
                }
                match at(j) {
                    Some('#' | '!' | '@') => {
                        integer = false;
                        j += 1;
                    }
                    Some('%' | '&' | '^') => j += 1,
                    _ => {}
                }
                let v: f64 = text.parse().map_err(|_| (line, "Expected: a number".to_string()))?;
                push(&mut out, Tok::Num(v, integer), line);
                i = j;
            }
            c if c.is_alphabetic() || c == '[' => {
                let mut name = String::new();
                if c == '[' {
                    // A bracketed name: [Some Name].
                    let mut j = i + 1;
                    while let Some(d) = at(j).filter(|d| *d != ']' && *d != '\n') {
                        name.push(d);
                        j += 1;
                    }
                    i = j + 1;
                } else {
                    while let Some(d) = at(i).filter(|d| d.is_alphanumeric() || *d == '_') {
                        name.push(d);
                        i += 1;
                    }
                    // A type character: Dim s$, n%. (`&` is always the concatenation operator here.)
                    if at(i).is_some_and(|s| matches!(s, '$' | '%' | '!' | '#' | '@')) && !at(i + 1).is_some_and(|n| n.is_alphanumeric() || n == '_') {
                        i += 1;
                    }
                }
                let at_statement_start = matches!(out.last().map(|t| &t.tok), None | Some(Tok::Newline | Tok::Colon));
                if at_statement_start && name.eq_ignore_ascii_case("rem") {
                    while at(i).is_some_and(|c| c != '\n') {
                        i += 1;
                    }
                    continue;
                }
                push(&mut out, Tok::Ident(name), line);
            }
            _ => {
                let two: String = [c, at(i + 1).unwrap_or(' ')].iter().collect();
                let (tok, len) = match two.as_str() {
                    "<=" => (Tok::Op("<="), 2),
                    ">=" => (Tok::Op(">="), 2),
                    "<>" => (Tok::Op("<>"), 2),
                    "=<" => (Tok::Op("<="), 2),
                    "=>" => (Tok::Op(">="), 2),
                    ":=" => (Tok::Assign, 2),
                    _ => (
                        match c {
                            '+' => Tok::Op("+"),
                            '-' => Tok::Op("-"),
                            '*' => Tok::Op("*"),
                            '/' => Tok::Op("/"),
                            '\\' => Tok::Op("\\"),
                            '^' => Tok::Op("^"),
                            '&' => Tok::Op("&"),
                            '=' => Tok::Op("="),
                            '<' => Tok::Op("<"),
                            '>' => Tok::Op(">"),
                            '(' => Tok::LParen,
                            ')' => Tok::RParen,
                            ',' => Tok::Comma,
                            '.' => Tok::Dot,
                            '!' => Tok::Bang,
                            ';' => Tok::Semi,
                            ':' => Tok::Colon,
                            '#' => return Err((line, "Date literals (#…#) and file numbers aren't supported".into())),
                            other => return Err((line, format!("Invalid character `{other}`"))),
                        },
                        1,
                    ),
                };
                push(&mut out, tok, line);
                i += len;
            }
        }
    }
    push(&mut out, Tok::Newline, line);
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn toks(s: &str) -> Vec<Tok> {
        lex(s).unwrap().into_iter().map(|t| t.tok).collect()
    }

    #[test]
    fn tokens() {
        assert_eq!(
            toks("x = \"a\"\"b\" & 3.5 ' note\r\nRem all of this\r\nIf a <> &HFF Then"),
            vec![
                Tok::Ident("x".into()),
                Tok::Op("="),
                Tok::Str("a\"b".into()),
                Tok::Op("&"),
                Tok::Num(3.5, false),
                Tok::Newline,
                Tok::Newline,
                Tok::Ident("If".into()),
                Tok::Ident("a".into()),
                Tok::Op("<>"),
                Tok::Num(255.0, true),
                Tok::Ident("Then".into()),
                Tok::Newline,
            ]
        );
        assert_eq!(
            toks("a = 1 + _\r\n    2"),
            vec![Tok::Ident("a".into()), Tok::Op("="), Tok::Num(1.0, true), Tok::Op("+"), Tok::Num(2.0, true), Tok::Newline]
        );
        assert_eq!(toks("s$ = n%"), vec![Tok::Ident("s".into()), Tok::Op("="), Tok::Ident("n".into()), Tok::Newline]);
        assert_eq!(toks("&HFFFF"), vec![Tok::Num(-1.0, true), Tok::Newline]);
        assert_eq!(toks("1E3 2#"), vec![Tok::Num(1000.0, false), Tok::Num(2.0, false), Tok::Newline]);
        assert!(lex("x = \"open").is_err());
        let lines: Vec<u32> = lex("a\r\nb _\r\n c\r\nd").unwrap().into_iter().map(|t| t.line).collect();
        assert_eq!(lines, vec![1, 1, 2, 3, 3, 4, 4]);
    }
}
