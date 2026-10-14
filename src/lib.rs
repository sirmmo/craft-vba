//! craft-vba: VBA macros for the craft apps, as an out-of-process add-in.
//!
//! - [`ovba`]: the workbook's VBA project (`vbaProject.bin`, [MS-OVBA]): modules in and out.
//! - [`vba`]: the language ([MS-VBAL]): parser and interpreter.
//! - [`excel`]: Excel's objects over the host's `host/sync` (the craft plugin protocol 1.1).
//! - [`plugin`]: the JSON-RPC add-in (`macros/*`).
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable, clippy::indexing_slicing)]
#![forbid(unsafe_code)]

/// The stack macros run with (calls nest up to 200 deep; parsing nests up to 100).
pub const STACK: usize = 256 << 20;

pub mod base64;
pub mod excel;
pub mod host;
pub mod ovba;
pub mod plugin;
pub mod record;
pub mod vba;

#[cfg(test)]
mod fakehost;
#[cfg(test)]
mod tests;
