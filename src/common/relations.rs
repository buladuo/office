use crate::error::Result;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::Cursor;

#[derive(Debug, Default)]
pub struct Relationships {
    map: HashMap<String, String>,
}

impl Relationships {
    pub fn new(map: HashMap<String, String>) -> Self {
        Relationships { map }
    }

    pub fn from_xml(xml_content: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml_content);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut rels = Relationships::default();
        let decoder = reader.decoder();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                    let mut id = None;
                    let mut target = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"Id" => id = Some(attr.decode_and_unescape_value(decoder)?.into_owned()),
                            b"Target" => {
                                target = Some(attr.decode_and_unescape_value(decoder)?.into_owned())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        rels.map.insert(id, target);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
    }

    pub fn get_target(&self, id: &str) -> Option<&String> {
        self.map.get(id)
    }

    pub fn to_xml(&self) -> Result<String> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        let mut root = BytesStart::new("Relationships");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(root))?;

        for (id, target) in &self.map {
            let mut element = BytesStart::new("Relationship");
            element.push_attribute(("Id", id.as_str()));
            element.push_attribute(("Target", target.as_str()));
            writer.write_event(Event::Empty(element))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;
        let result = writer.into_inner().into_inner();
        Ok(String::from_utf8(result)?)
    }

    #[cfg(test)]
    pub fn from_map(map: HashMap<String, String>) -> Self {
        Relationships { map }
    }
}
