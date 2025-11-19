use quick_xml::de::from_str;
use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;

use crate::common::relations::Relationships;
use crate::common::xml_utils::read_element_xml;
use crate::docx::properties::ParagraphProperties;
use crate::docx::styles::{Style, Styles};
use crate::error::{OfficeError, Result};

use super::hyperlink::Hyperlink;
use super::run::Run;

/// 段落内容枚举，表示段落中可能包含的内容类型
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ParagraphContent {
    /// 文本运行
    #[serde(rename = "w:r")]
    Run(Run),
    /// 超链接
    #[serde(rename = "w:hyperlink")]
    Hyperlink(Hyperlink),
}

impl From<Run> for ParagraphContent {
    fn from(r: Run) -> Self {
        ParagraphContent::Run(r)
    }
}

impl From<Hyperlink> for ParagraphContent {
    fn from(h: Hyperlink) -> Self {
        ParagraphContent::Hyperlink(h)
    }
}

/// 段落结构体，表示文档中的段落元素
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:p")]
pub struct Paragraph {
    /// 段落属性
    #[serde(rename = "w:pPr", skip_serializing_if = "Option::is_none")]
    pub properties: Option<ParagraphProperties>,
    /// 段落内容列表
    #[serde(rename = "$value")]
    pub content: Vec<ParagraphContent>,
}

impl Paragraph {
    /// 获取段落的样式
    /// 
    /// # 参数
    /// * `styles` - 样式集合
    pub fn get_style<'a>(&'a self, styles: &'a Styles) -> Option<&'a Style> {
        self.properties
            .as_ref()
            .and_then(|p| p.style.as_ref())
            .and_then(|s| styles.find_style(&s.val))
    }

    /// 从XML读取器中解析段落
    /// 
    /// # 参数
    /// * `reader` - XML读取器
    /// * `tag_name` - 标签名称
    /// * `rels` - 文档关系信息
    pub fn from_xml_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
        rels: Option<&Relationships>,
    ) -> Result<Self> {
        let mut paragraph = Paragraph::default();
        let mut buf = Vec::new();

        // 循环读取段落中的内容
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理开始标签
                Event::Start(e) => match e.name().as_ref() {
                    // 段落属性标签
                    b"w:pPr" => {
                        let p_pr_xml = read_element_xml(reader, &e)?;
                        paragraph.properties = from_str(&p_pr_xml).ok();
                    }
                    // 文本运行标签
                    b"w:r" => {
                        let run = Run::from_xml_reader(reader, e.name())?;
                        paragraph.content.push(ParagraphContent::Run(run));
                    }
                    // 超链接标签
                    b"w:hyperlink" => {
                        let hyperlink = Hyperlink::from_xml_reader(reader, &e, rels)?;
                        paragraph
                            .content
                            .push(ParagraphContent::Hyperlink(hyperlink));
                    }
                    // 其他标签直接跳过
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                // 处理段落结束标签
                Event::End(e) if e.name() == tag_name => break,
                // 处理意外的文件结束
                Event::Eof => {
                    return Err(OfficeError::InvalidFormat(
                        "Unexpected EOF in paragraph".to_string(),
                    ))
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(paragraph)
    }
}