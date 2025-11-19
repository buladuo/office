use crate::common::xml_utils::read_element_xml;
use crate::error::{OfficeError, Result};
use quick_xml::de::from_str;
use quick_xml::events::Event;
use quick_xml::Reader;

use super::properties::{ParagraphProperties, RunProperties};

#[derive(Debug, Default)]
pub struct Document {
    pub paragraphs: Vec<Paragraph>,
}

#[derive(Debug, Default)]
pub struct Paragraph {
    pub properties: Option<ParagraphProperties>,
    pub runs: Vec<Run>,
}

#[derive(Debug, Default)]
pub struct Run {
    pub properties: Option<RunProperties>,
    pub text: Option<String>,
}

impl Document {
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml_content);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        let mut doc = Document::default();

        // Find the <w:body> tag
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.name().as_ref() == b"w:body" => {
                    doc = Document::from_body_reader(&mut reader, e.name())?;
                    break; // Found and parsed the body
                }
                Event::Eof => break, // End of file
                _ => {}
            }
            buf.clear();
        }

        Ok(doc)
    }

    fn from_body_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
    ) -> Result<Self> {
        let mut doc = Document::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) if e.name().as_ref() == b"w:p" => {
                    doc.paragraphs
                        .push(Paragraph::from_xml_reader(reader, e.name())?);
                }
                Event::Empty(e) if e.name().as_ref() == b"w:p" => {
                    doc.paragraphs.push(Paragraph::default());
                }
                Event::End(e) if e.name() == tag_name => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(doc)
    }
}

impl Paragraph {
    fn from_xml_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
    ) -> Result<Self> {
        let mut paragraph = Paragraph::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => match e.name().as_ref() {
                    b"w:pPr" => {
                        let p_pr_xml = read_element_xml(reader, &e)?;
                        paragraph.properties = from_str(&p_pr_xml).ok();
                    }
                    b"w:r" => {
                        paragraph
                            .runs
                            .push(Run::from_xml_reader(reader, e.name())?);
                    }
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                Event::End(e) if e.name() == tag_name => break,
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

impl Run {
    fn from_xml_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
    ) -> Result<Self> {
        let mut run = Run::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => match e.name().as_ref() {
                    b"w:rPr" => {
                        let r_pr_xml = read_element_xml(reader, &e)?;
                        run.properties = from_str(&r_pr_xml).ok();
                    }
                    b"w:t" => {
                        let mut text_val = String::new();
                        let mut text_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut text_buf)? {
                                Event::Text(t) => {
                                    text_val.push_str(t.decode()?.as_ref());
                                }
                                Event::End(end) if end.name() == e.name() => break,
                                Event::Eof => {
                                    return Err(OfficeError::InvalidFormat(
                                        "Unexpected EOF in <w:t>".to_string(),
                                    ));
                                }
                                _ => {}
                            }
                            text_buf.clear();
                        }
                        run.text = Some(text_val);
                    }
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                Event::Empty(e) => {
                    if e.name().as_ref() == b"w:t" {
                        run.text = Some("".to_string());
                    }
                }
                Event::End(e) if e.name() == tag_name => break,
                Event::Eof => {
                    return Err(OfficeError::InvalidFormat(
                        "Unexpected EOF in run".to_string(),
                    ))
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(run)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::properties::JustificationVal;

    const MOCK_DOCUMENT_XML: &str = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="Title"/>
                    </w:pPr>
                    <w:r>
                        <w:rPr>
                            <w:b/>
                        </w:rPr>
                        <w:t>Hello, </w:t>
                    </w:r>
                    <w:r>
                        <w:rPr>
                            <w:i/>
                        </w:rPr>
                        <w:t>World!</w:t>
                    </w:r>
                </w:p>
                <w:p>
                    <w:pPr>
                        <w:jc w:val="center"/>
                    </w:pPr>
                    <w:r>
                        <w:t>This is the second paragraph.</w:t>
                    </w:r>
                </w:p>
                <w:p/>
            </w:body>
        </w:document>
    "#;

    #[test]
    fn test_parse_document_with_properties() {
        let doc = Document::from_xml(MOCK_DOCUMENT_XML).unwrap();

        assert_eq!(doc.paragraphs.len(), 3);

        // First paragraph
        let p1 = &doc.paragraphs[0];
        assert_eq!(p1.runs.len(), 2);
        assert!(p1.properties.is_some());
        assert_eq!(
            p1.properties
                .as_ref()
                .unwrap()
                .style
                .as_ref()
                .unwrap()
                .val,
            "Title"
        );

        // First run in first paragraph
        let r1 = &p1.runs[0];
        assert_eq!(r1.text.as_deref(), Some("Hello, "));
        assert!(r1.properties.is_some());
        assert!(r1.properties.as_ref().unwrap().bold.is_some());
        assert!(r1.properties.as_ref().unwrap().italic.is_none());

        // Second run in first paragraph
        let r2 = &p1.runs[1];
        assert_eq!(r2.text.as_deref(), Some("World!"));
        assert!(r2.properties.is_some());
        assert!(r2.properties.as_ref().unwrap().italic.is_some());
        assert!(r2.properties.as_ref().unwrap().bold.is_none());

        // Second paragraph
        let p2 = &doc.paragraphs[1];
        assert_eq!(p2.runs.len(), 1);
        assert!(p2.properties.is_some());
        let p2_props = p2.properties.as_ref().unwrap();
        assert!(p2_props.justification.is_some());
        matches!(
            p2_props.justification.as_ref().unwrap().val,
            JustificationVal::Center
        );
        assert_eq!(
            p2.runs[0].text.as_deref(),
            Some("This is the second paragraph.")
        );

        // Third paragraph (empty)
        let p3 = &doc.paragraphs[2];
        assert_eq!(p3.runs.len(), 0);
        assert!(p3.properties.is_none());
    }
}
