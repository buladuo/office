use crate::error::{OfficeError, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::writer::Writer;
use quick_xml::Reader;
use std::io::Cursor;

/// Reads an entire XML element, from its start tag to its corresponding end tag,
/// and returns it as a string. This is useful for deserializing a whole element.
/// This function is called after the initial start event has been read.
pub fn read_element_xml<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start_tag: &BytesStart, // The start tag that was just read
) -> Result<String> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    writer.write_event(Event::Start(start_tag.clone()))?;

    let mut depth = 1;

    loop {
        let mut buf = Vec::new();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                depth += 1;
                writer.write_event(Event::Start(e))?;
            }
            Event::End(e) => {
                depth -= 1;
                writer.write_event(Event::End(e))?;
                if depth == 0 {
                    break;
                }
            }
            Event::Empty(e) => {
                writer.write_event(Event::Empty(e))?;
            }
            Event::Eof => {
                return Err(OfficeError::InvalidFormat(format!(
                    "Unexpected EOF while reading element '{}'",
                    String::from_utf8_lossy(start_tag.name().as_ref())
                )));
            }
            event => {
                writer.write_event(event)?;
            }
        }
    }

    let result = writer.into_inner().into_inner();
    Ok(String::from_utf8(result)?)
}