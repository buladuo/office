use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;

use crate::common::relations::Relationships;
use crate::error::{OfficeError, Result};

use super::{BodyContent, Paragraph};

/// 表格结构体，表示文档中的表格元素
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:tbl")]
pub struct Table {
    /// 表格行列表
    #[serde(rename = "w:tr")]
    pub rows: Vec<TableRow>,
}

/// 表格行结构体，表示表格中的一行
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:tr")]
pub struct TableRow {
    /// 表格单元格列表
    #[serde(rename = "w:tc")]
    pub cells: Vec<TableCell>,
}

/// 表格单元格结构体，表示表格中的一个单元格
#[derive(Debug, Default, Serialize)]
#[serde(rename = "w:tc")]
pub struct TableCell {
    /// 单元格内容，可以是段落或其他元素
    #[serde(rename = "$value")]
    pub content: Vec<BodyContent>,
}

impl Table {
    /// 从XML读取器中解析表格
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
        let mut table = Table::default();
        let mut buf = Vec::new();

        // 循环读取表格中的行
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理行开始标签
                Event::Start(e) if e.name().as_ref() == b"w:tr" => {
                    table
                        .rows
                        .push(TableRow::from_xml_reader(reader, e.name(), rels)?);
                }
                // 处理表格结束标签
                Event::End(e) if e.name() == tag_name => break,
                // 处理意外的文件结束
                Event::Eof => {
                    return Err(OfficeError::InvalidFormat(
                        "Unexpected EOF in table".to_string(),
                    ))
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(table)
    }
}

impl TableRow {
    /// 从XML读取器中解析表格行
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
        let mut row = TableRow::default();
        let mut buf = Vec::new();

        // 循环读取行中的单元格
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理单元格开始标签
                Event::Start(e) if e.name().as_ref() == b"w:tc" => {
                    row.cells
                        .push(TableCell::from_xml_reader(reader, e.name(), rels)?);
                }
                // 处理行结束标签
                Event::End(e) if e.name() == tag_name => break,
                // 处理意外的文件结束
                Event::Eof => {
                    return Err(OfficeError::InvalidFormat(
                        "Unexpected EOF in table row".to_string(),
                    ))
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(row)
    }
}

impl TableCell {
    /// 从XML读取器中解析表格单元格
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
        let mut cell = TableCell::default();
        let mut buf = Vec::new();

        // 循环读取单元格中的内容
        loop {
            match reader.read_event_into(&mut buf)? {
                // 处理开始标签
                Event::Start(e) => match e.name().as_ref() {
                    // 段落标签
                    b"w:p" => {
                        let p = Paragraph::from_xml_reader(reader, e.name(), rels)?;
                        cell.content.push(BodyContent::Paragraph(p));
                    }
                    // 嵌套表格标签
                    b"w:tbl" => {
                        let t = Table::from_xml_reader(reader, e.name(), rels)?;
                        cell.content.push(BodyContent::Table(t));
                    }
                    // 其他标签直接跳过
                    _ => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                },
                // 处理单元格结束标签
                Event::End(e) if e.name() == tag_name => break,
                // 处理意外的文件结束
                Event::Eof => {
                    return Err(OfficeError::InvalidFormat(
                        "Unexpected EOF in table cell".to_string(),
                    ))
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(cell)
    }
}