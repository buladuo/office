//! DOCX numbering (lists) format definitions.

use serde::Deserialize;

#[derive(Debug, Default, Deserialize)]
#[serde(rename = "numbering")]
pub struct Numbering {
    // We will add fields here to represent abstractNum, num, etc.
}

impl Numbering {
    pub fn from_xml(xml_content: &str) -> crate::error::Result<Self> {
        if xml_content.is_empty() {
            return Ok(Numbering::default());
        }
        let numbering: Numbering = quick_xml::de::from_str(xml_content)?;
        Ok(numbering)
    }
}
