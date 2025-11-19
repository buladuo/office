use thiserror::Error;
use zip::result::ZipError;
use quick_xml::Error as XmlError;
use quick_xml::events::attributes::AttrError;
use quick_xml::encoding::EncodingError;
use std::string::FromUtf8Error;
use std::io::Error as IoError;

#[derive(Debug, Error)]
pub enum OfficeError {
    #[error("I/O error: {0}")]
    Io(#[from] IoError),

    #[error("Zip error: {0}")]
    Zip(#[from] ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] XmlError),

    #[error("XML attribute error: {0}")]
    Attribute(#[from] AttrError),

    #[error("XML encoding error: {0}")]
    Encoding(#[from] EncodingError),

    #[error("UTF-8 conversion error: {0}")]
    Parse(#[from] FromUtf8Error),

    #[error("File not found in archive: {0}")]
    FileNotFoundInArchive(String),

    #[error("Invalid file format: {0}")]
    InvalidFormat(String),

    #[error("Unsupported format or feature: {0}")]
    Unsupported(String),
}

pub type Result<T> = std::result::Result<T, OfficeError>;