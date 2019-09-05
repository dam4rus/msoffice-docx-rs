use msoffice_shared::error::{ParseEnumError, ParseHexColorRGBError};
use std::{
    error::Error,
    fmt::{Display, Formatter},
};

#[derive(Debug, Clone, PartialEq)]
pub enum ParseHexColorError {
    Enum(ParseEnumError),
    HexColorRGB(ParseHexColorRGBError),
}

impl Display for ParseHexColorError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match *self {
            ParseHexColorError::Enum(ref e) => e.fmt(f),
            ParseHexColorError::HexColorRGB(ref e) => e.fmt(f),
        }
    }
}

impl Error for ParseHexColorError {}

impl From<ParseEnumError> for ParseHexColorError {
    fn from(v: ParseEnumError) -> Self {
        ParseHexColorError::Enum(v)
    }
}

impl From<ParseHexColorRGBError> for ParseHexColorError {
    fn from(v: ParseHexColorRGBError) -> Self {
        ParseHexColorError::HexColorRGB(v)
    }
}
