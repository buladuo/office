//! DOCX格式属性的定义
//! Defines structs for properties in DOCX format.

use serde::{Deserialize, Serialize};

/// A generic struct for elements that only have a `w:val` attribute.
#[derive(Debug, Deserialize, Serialize)]
pub struct Val<T> {
    #[serde(rename = "@w:val")]
    pub val: T,
}

/// 段落对齐方式
/// Paragraph alignment
#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum JustificationVal {
    #[default]
    Left,
    Center,
    Right,
    Both,
    Distribute,
}

#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename = "w:jc")]
pub struct Justification {
    #[serde(rename = "@w:val")]
    pub val: JustificationVal,
}

/// 段落样式
/// Paragraph style
#[derive(Debug, Default, Deserialize, Serialize)]
pub struct ParagraphStyle {
    #[serde(rename = "@w:val")]
    pub val: String,
}

/// 列表的缩进级别
/// Indentation level for a list item.
#[derive(Debug, Deserialize, Serialize)]
#[serde(rename = "w:ilvl")]
pub struct NumLvl {
    #[serde(rename = "@w:val")]
    pub val: i32,
}

/// 列表属性，关联一个段落到一个列表
/// Numbering properties, associating a paragraph with a list.
#[derive(Debug, Deserialize, Serialize)]
#[serde(rename = "w:numPr")]
pub struct NumPr {
    #[serde(rename = "w:ilvl")]
    pub level: NumLvl,
    #[serde(rename = "w:numId")]
    pub num_id: Val<i32>,
}

/// 段落属性
/// Paragraph properties
#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename = "w:pPr")]
pub struct ParagraphProperties {
    #[serde(rename = "w:pStyle", skip_serializing_if = "Option::is_none")]
    pub style: Option<ParagraphStyle>,
    #[serde(rename = "w:jc", skip_serializing_if = "Option::is_none")]
    pub justification: Option<Justification>,
    #[serde(rename = "w:numPr", skip_serializing_if = "Option::is_none")]
    pub num_pr: Option<NumPr>,
}

/// 运行属性 (文字属性)
/// Run properties (text properties)
#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename = "w:rPr")]
pub struct RunProperties {
    #[serde(rename = "w:b", skip_serializing_if = "Option::is_none")]
    pub bold: Option<()>, // For toggle properties, we just care if the tag exists
    #[serde(rename = "w:i", skip_serializing_if = "Option::is_none")]
    pub italic: Option<()>,
    #[serde(rename = "w:u", skip_serializing_if = "Option::is_none")]
    pub underline: Option<()>,
    #[serde(rename = "w:rStyle", skip_serializing_if = "Option::is_none")]
    pub style: Option<ParagraphStyle>, // Re-using ParagraphStyle for run style
                                       // Future properties can be added here, e.g., color, size
}
