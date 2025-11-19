pub mod error;
pub mod common;

#[cfg(feature = "docx")]
pub mod docx;

#[cfg(feature = "xlsx")]
pub mod xlsx;

#[cfg(feature = "pptx")]
pub mod pptx;