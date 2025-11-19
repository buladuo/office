//! Defines structs for Core Document Properties.

use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename = "cp:coreProperties")]
pub struct CoreProps {
    #[serde(
        rename = "dc:title",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub title: Option<String>,
    #[serde(
        rename = "dc:creator",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub creator: Option<String>,
    #[serde(
        rename = "dc:description",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub description: Option<String>,
    #[serde(
        rename = "cp:lastModifiedBy",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub last_modified_by: Option<String>,
    #[serde(
        rename = "cp:revision",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub revision: Option<String>,
    #[serde(
        rename = "dcterms:created",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub created: Option<String>, // Could be parsed into a DateTime
    #[serde(
        rename = "dcterms:modified",
        with = "quick_xml::serde_helpers::text_content",
        skip_serializing_if = "Option::is_none"
    )]
    pub modified: Option<String>, // Could be parsed into a DateTime
}

impl CoreProps {
    pub fn from_xml(xml_content: &str) -> crate::error::Result<Self> {
        if xml_content.is_empty() {
            // Create a default if the content is empty
            let props: CoreProps = CoreProps {
                title: None,
                creator: None,
                description: None,
                created: None,
                modified: None,
                last_modified_by: None,
                revision: None,
            };
            return Ok(props);
        }
        let props: CoreProps = quick_xml::de::from_str(xml_content)?;
        Ok(props)
    }
}
