use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;

use crate::common::relations::Relationships;
use crate::error::Result;

use super::elements::{BodyContent, Paragraph, Table};

/// DOCX文档结构体，表示整个文档
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:document")]
pub struct Document {
    /// 文档主体内容
    #[serde(rename = "w:body")]
    pub body: Body,
    /// WordML命名空间
    #[serde(rename = "@xmlns:w")]
    pub xmlns_w: String,
    /// 关系命名空间
    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,
}

/// 文档主体结构体，包含文档的主要内容
#[derive(Debug, Default, Serialize)]
pub struct Body {
    /// 主体内容，可以是段落或表格等
    #[serde(rename = "$value")]
    pub content: Vec<BodyContent>,
}

impl Document {
    /// 创建默认的Document实例
    pub fn default() -> Self {
        Document {
            body: Body::default(),
            xmlns_w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
        }
    }
    
    /// 从XML内容解析Document
    /// 
    /// # 参数
    /// * `xml_content` - XML格式的文档内容
    /// * `rels` - 文档关系信息
    pub fn from_xml(
        xml_content: &str,
        rels: Option<&Relationships>,
    ) -> Result<Self> {
        // 创建XML读取器
        let mut reader = Reader::from_str(xml_content);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let mut doc = Document::default();

        // 读取XML事件，寻找body标签
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.name().as_ref() == b"w:body" => {
                    // 解析body内容
                    doc.body = Body::from_body_reader(&mut reader, e.name(), rels)?;
                    break;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(doc)
    }
}

impl Body {
    /// 从XML读取器中解析body内容
    /// 
    /// # 参数
    /// * `reader` - XML读取器
    /// * `tag_name` - 当前标签名称
    /// * `rels` - 文档关系信息
    fn from_body_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
        rels: Option<&Relationships>,
    ) -> Result<Self> {
        let mut body = Body::default();
        let mut buf = Vec::new();

        // 循环读取body中的内容
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理开始标签
                Event::Start(e) => match e.name().as_ref() {
                    // 段落标签
                    b"w:p" => {
                        let paragraph = Paragraph::from_xml_reader(reader, e.name(), rels)?;
                        body.content.push(BodyContent::Paragraph(paragraph));
                    }
                    // 表格标签
                    b"w:tbl" => {
                        let table = Table::from_xml_reader(reader, e.name(), rels)?;
                        body.content.push(BodyContent::Table(table));
                    }
                    // 其他标签直接跳过
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                // 处理空的段落标签
                Event::Empty(e) if e.name().as_ref() == b"w:p" => {
                    body.content
                        .push(BodyContent::Paragraph(Paragraph::default()));
                }
                // 结束标签，结束解析
                Event::End(e) if e.name() == tag_name => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(body)
    }
}