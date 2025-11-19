use crate::common::package::open_package;
use crate::error::Result;
use std::path::Path;

pub mod document;
pub mod numbering;
pub mod properties;
pub mod styles;

use document::Document;
use numbering::Numbering;
use styles::Styles;

#[derive(Debug, Default)]
pub struct Docx {
    pub document: Document,
    pub styles: Styles,
    pub numbering: Numbering,
}

impl Docx {
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let mut package = open_package(path)?;

        let styles_content =
            package.read_file_by_path("word/styles.xml")?;
        let styles = Styles::from_xml(&styles_content)?;

        let numbering_content = package
            .read_file_by_path("word/numbering.xml")
            .unwrap_or_default();
        let numbering = Numbering::from_xml(&numbering_content)?;

        let document_content = package.read_file_by_path("word/document.xml")?;
        let document = Document::from_xml(&document_content)?;

        Ok(Docx {
            document,
            styles,
            numbering,
        })
    }
}
