#![forbid(unsafe_code)]

extern crate strum;
#[macro_use]
extern crate strum_macros;

pub mod error;
pub mod package;
pub mod wml;

pub extern crate msoffice_shared as shared;