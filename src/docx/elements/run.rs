use quick_xml::de::from_str;
use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;

use crate::common::xml_utils::read_element_xml;
use crate::docx::properties::RunProperties;
use crate::error::{OfficeError, Result};

/// 文本运行结构体，表示文档中具有相同属性的一段文本
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:r")]
pub struct Run {
    /// 文本运行属性
    #[serde(rename = "w:rPr", skip_serializing_if = "Option::is_none")]
    pub properties: Option<RunProperties>,
    /// 文本运行内容列表
    #[serde(rename = "$value")]
    pub content: Vec<RunContent>,
}

/// 文本运行内容枚举，表示文本运行中可能包含的内容类型
#[derive(Debug, Serialize)]
pub enum RunContent {
    /// 文本内容
    #[serde(rename = "w:t")]
    Text(String),
    /// 换行符
    #[serde(rename = "w:br")]
    Break,
    /// 制表符
    #[serde(rename = "w:tab")]
    Tab,
}

impl Run {
    /// 从XML读取器中解析文本运行
    /// 
    /// # 参数
    /// * `reader` - XML读取器
    /// * `tag_name` - 标签名称
    pub fn from_xml_reader<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        tag_name: quick_xml::name::QName,
    ) -> Result<Self> {
        let mut run = Run::default();
        let mut buf = Vec::new();

        // 循环读取文本运行中的内容
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理开始标签
                Event::Start(e) => match e.name().as_ref() {
                    // 文本运行属性标签
                    b"w:rPr" => {
                        let r_pr_xml = read_element_xml(reader, &e)?;
                        run.properties = from_str(&r_pr_xml).ok();
                    }
                    // 文本标签
                    b"w:t" => {
                        run.content.push(RunContent::Text(read_text_node(
                            reader,
                            e.name(),
                        )?));
                    }
                    // 其他标签直接跳过
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                // 处理空标签
                Event::Empty(e) => match e.name().as_ref() {
                    // 空文本标签
                    b"w:t" => run.content.push(RunContent::Text(String::new())),
                    // 换行标签
                    b"w:br" => run.content.push(RunContent::Break),
                    // 制表符标签
                    b"w:tab" => run.content.push(RunContent::Tab),
                    _ => {}
                },
                // 处理文本运行结束标签
                Event::End(e) if e.name() == tag_name => break,
                // 处理意外的文件结束
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

/// 读取文本节点内容
/// 
/// # 参数
/// * `reader` - XML读取器
/// * `tag_name` - 标签名称
pub(crate) fn read_text_node<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    tag_name: quick_xml::name::QName,
) -> Result<String> {
    let mut text_val = String::new();
    let mut text_buf = Vec::new();
    // 循环读取文本内容
    loop {
        match reader.read_event_into(&mut text_buf)? {
            // 处理文本内容
            Event::Text(t) => {
                text_val.push_str(t.decode()?.as_ref());
            }
            // 处理文本标签结束
            Event::End(end) if end.name() == tag_name => break,
            // 处理意外的文件结束
            Event::Eof => {
                return Err(OfficeError::InvalidFormat(
                    "Unexpected EOF in <w:t>".to_string(),
                ));
            }
            _ => {}
        }
        text_buf.clear();
    }
    Ok(text_val)
}