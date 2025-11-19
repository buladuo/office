//! Defines structs for Application-Specific Document Properties.

use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename = "Properties")]
pub struct AppProps {
    #[serde(rename = "Application", skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,
    #[serde(rename = "AppVersion", skip_serializing_if = "Option::is_none")]
    pub app_version: Option<String>,
    // Other properties like Pages, Words, Characters can be added here.
}

impl AppProps {
    pub fn from_xml(xml_content: &str) -> crate::error::Result<Self> {
        if xml_content.is_empty() {
            // Create a default if the content is empty
            let props: AppProps = AppProps {
                application: None,
                app_version: None,
            };
            return Ok(props);
        }
        let props: AppProps = quick_xml::de::from_str(xml_content)?;
        Ok(props)
    }
}
