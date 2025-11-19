use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;

use crate::common::relations::Relationships;
use crate::error::Result;

use super::elements::{BodyContent, Paragraph, Table};

#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:document")]
pub struct Document {
    #[serde(rename = "w:body")]
    pub body: Body,
    #[serde(rename = "@xmlns:w")]
    pub xmlns_w: String,
    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,
}

#[derive(Debug, Default, Serialize)]
pub struct Body {
    #[serde(rename = "$value")]
    pub content: Vec<BodyContent>,
}

impl Document {
    pub fn default() -> Self {
        Document {
            body: Body::default(),
            xmlns_w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
        }
    }
    pub fn from_xml(
        xml_content: &str,
        rels: Option<&Relationships>,
    ) -> Result<Self> {
        let mut reader = Reader::from_str(xml_content);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let mut doc = Document::default();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.name().as_ref() == b"w:body" => {
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
    fn from_body_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
        rels: Option<&Relationships>,
    ) -> Result<Self> {
        let mut body = Body::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => match e.name().as_ref() {
                    b"w:p" => {
                        let paragraph = Paragraph::from_xml_reader(reader, e.name(), rels)?;
                        body.content.push(BodyContent::Paragraph(paragraph));
                    }
                    b"w:tbl" => {
                        let table = Table::from_xml_reader(reader, e.name(), rels)?;
                        body.content.push(BodyContent::Table(table));
                    }
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                Event::Empty(e) if e.name().as_ref() == b"w:p" => {
                    body.content
                        .push(BodyContent::Paragraph(Paragraph::default()));
                }
                Event::End(e) if e.name() == tag_name => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(body)
    }
}
