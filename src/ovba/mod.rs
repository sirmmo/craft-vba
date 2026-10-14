//! [MS-OVBA]: the VBA project format.

pub mod codepage;
pub mod compress;
pub mod project;

pub use project::{Module, ModuleKind, Project};
