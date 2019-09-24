use std::{
    error::Error,
    fmt::{Display, Formatter},
};

#[derive(Debug, Clone, PartialEq, Default)]
pub struct NoSuchStyleError;

impl Display for NoSuchStyleError {
    fn fmt(&self, f: &mut Formatter) -> std::fmt::Result {
        write!(f, "Style not found")
    }
}

impl Error for NoSuchStyleError {}
