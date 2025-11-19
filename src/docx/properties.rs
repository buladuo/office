//! DOCX格式属性的定义
//! Defines structs for properties in DOCX format.

use serde::Deserialize;

/// 段落对齐方式
/// Paragraph alignment
#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum JustificationVal {
    #[default]
    Left,
    Center,
    Right,
    Both,
    Distribute,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename = "jc")]
pub struct Justification {
    #[serde(rename = "@val")]
    pub val: JustificationVal,
}

/// 段落样式
/// Paragraph style
#[derive(Debug, Default, Deserialize)]
pub struct ParagraphStyle {
    #[serde(rename = "@val")]
    pub val: String,
}

/// 段落属性
/// Paragraph properties
#[derive(Debug, Default, Deserialize)]
#[serde(rename = "pPr")]
pub struct ParagraphProperties {
    #[serde(rename = "pStyle")]
    pub style: Option<ParagraphStyle>,
    #[serde(rename = "jc")]
    pub justification: Option<Justification>,
    // Future properties can be added here, e.g., spacing, indentation
}

/// 运行属性 (文字属性)
/// Run properties (text properties)
#[derive(Debug, Default, Deserialize)]
#[serde(rename = "rPr")]
pub struct RunProperties {
    #[serde(rename = "b")]
    pub bold: Option<serde::de::IgnoredAny>, // For toggle properties, we just care if the tag exists
    #[serde(rename = "i")]
    pub italic: Option<serde::de::IgnoredAny>,
    #[serde(rename = "u")]
    pub underline: Option<serde::de::IgnoredAny>,
    #[serde(rename = "rStyle")]
    pub style: Option<ParagraphStyle>, // Re-using ParagraphStyle for run style
                                       // Future properties can be added here, e.g., color, size
}
