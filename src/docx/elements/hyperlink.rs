use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use serde::Serialize;

use crate::common::relations::Relationships;
use crate::error::{OfficeError, Result};

use super::run::Run;

/// 超链接结构体，表示文档中的超链接元素
#[derive(Debug, Default, Serialize)]
pub struct Hyperlink {
    /// 关系ID，用于查找超链接的目标URL
    #[serde(rename = "@r:id")]
    pub r_id: String,
    /// 超链接中的文本运行列表
    #[serde(rename = "$value")]
    pub runs: Vec<Run>,
}

impl Hyperlink {
    /// 从XML读取器中解析超链接
    /// 
    /// # 参数
    /// * `reader` - XML读取器
    /// * `start_tag` - 起始标签
    /// * `rels` - 文档关系信息，用于解析超链接目标
    pub fn from_xml_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        rels: Option<&Relationships>,
    ) -> Result<Self> {
        let mut hyperlink = Hyperlink::default();
        let mut buf = Vec::new();

        // 获取解码器，用于解码属性值
        let decoder = reader.decoder();
        // 如果有关系信息，则解析关系ID
        if rels.is_some() {
            for attr in start_tag.attributes() {
                let attr = attr?;
                // 查找关系ID属性
                if attr.key.as_ref() == b"r:id" {
                    let r_id = attr.decode_and_unescape_value(decoder)?.to_string();
                    hyperlink.r_id = r_id;
                    break;
                }
            }
        }

        // 循环读取超链接中的内容
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理文本运行开始标签
                Event::Start(e) if e.name().as_ref() == b"w:r" => {
                    hyperlink
                        .runs
                        .push(Run::from_xml_reader(reader, e.name())?);
                }
                // 处理超链接结束标签
                Event::End(e) if e.name() == start_tag.name() => break,
                // 处理意外的文件结束
                Event::Eof => {
                    return Err(OfficeError::InvalidFormat(
                        "Unexpected EOF in hyperlink".to_string(),
                    ))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(hyperlink)
    }
}