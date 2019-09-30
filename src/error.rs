use std::{
    error::Error,
    fmt::{Display, Formatter},
};

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct NoSuchStyleError;

impl Display for NoSuchStyleError {
    fn fmt(&self, f: &mut Formatter) -> std::fmt::Result {
        write!(f, "Style not found")
    }
}

impl Error for NoSuchStyleError {}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct NoSuchFileError {}

impl Display for NoSuchFileError {
    fn fmt(&self, f: &mut Formatter) -> std::fmt::Result {
        write!(f, "File not found")
    }
}

impl Error for NoSuchFileError {}
