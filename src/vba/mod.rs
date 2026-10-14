//! VBA: the language ([MS-VBAL]).

pub mod ast;
pub mod builtins;
pub mod interp;
pub mod lexer;
pub mod ops;
pub mod parser;
pub mod value;

pub use interp::{Failure, Interp, Program};
