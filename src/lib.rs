#![forbid(unsafe_code)]

extern crate strum;
#[macro_use]
extern crate strum_macros;

pub mod package;
pub mod resolvedstyle;
pub mod wml;

pub extern crate msoffice_shared as shared;
