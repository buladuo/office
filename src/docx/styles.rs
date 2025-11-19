use crate::error::Result;
use crate::docx::properties::{ParagraphProperties, RunProperties};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// 样式类型枚举
#[derive(Debug, Deserialize, Serialize, PartialEq, Eq, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum StyleType {
    /// 段落样式
    Paragraph,
    /// 字符样式
    Character,
    /// 表格样式
    Table,
    /// 编号样式
    Numbering,
}

/// 样式结构体，表示文档中的一个样式定义
#[derive(Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Style {
    /// 样式类型
    #[serde(rename = "@w:type")]
    pub style_type: StyleType,
    /// 样式ID
    #[serde(rename = "@w:styleId")]
    pub style_id: String,
    /// 段落属性
    #[serde(rename = "w:pPr", skip_serializing_if = "Option::is_none")]
    pub paragraph_properties: Option<ParagraphProperties>,
    /// 文本运行属性
    #[serde(rename = "w:rPr", skip_serializing_if = "Option::is_none")]
    pub run_properties: Option<RunProperties>,
    // 其他字段如名称、基于等可以在这里添加
}

/// 样式集合结构体，包含所有样式定义
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:styles")]
pub struct Styles {
    /// 样式列表
    #[serde(rename = "w:style")]
    styles: Vec<Style>,
}

/// 样式根结构体，用于反序列化
#[derive(Debug, Deserialize)]
struct StylesRoot {
    /// 样式列表
    #[serde(rename = "style", default)]
    styles: Vec<Style>,
}

impl Styles {
    /// 从XML内容解析样式
    /// 
    /// # 参数
    /// * `xml_content` - XML格式的样式内容
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        // 如果XML内容为空，返回默认样式集合
        if xml_content.is_empty() {
            return Ok(Styles::default());
        }

        // 反序列化XML内容
        let root: StylesRoot = quick_xml::de::from_str(xml_content)?;
        let styles = root
            .styles
            .into_iter()
            // 可以在这里映射为HashMap等其他数据结构
            .collect();

        Ok(Styles { styles })
    }

    /// 根据样式ID查找样式
    /// 
    /// # 参数
    /// * `style_id` - 要查找的样式ID
    pub fn find_style(&self, style_id: &str) -> Option<&Style> {
        self.styles.iter().find(|s| s.style_id == style_id)
    }
}