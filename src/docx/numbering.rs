//! DOCX 编号（列表）格式定义

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

use crate::docx::properties::Val;

// --- 用于反序列化的原始结构体 ---

/// 编号格式值枚举
#[derive(Debug, Deserialize, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
enum NumFmtVal {
    /// 项目符号
    Bullet,
    /// 十进制数字
    Decimal,
    /// 小写字母
    LowerLetter,
    /// 大写字母
    UpperLetter,
    /// 小写罗马数字
    LowerRoman,
    /// 大写罗马数字
    UpperRoman,
    /// 无格式
    None,
}

/// 原始编号格式结构体
#[derive(Debug, Deserialize, Serialize)]
struct RawNumFmt {
    /// 格式值
    #[serde(rename = "@w:val")]
    val: NumFmtVal,
}

/// 原始级别结构体
#[derive(Debug, Deserialize, Serialize)]
struct RawLevel {
    /// 级别索引
    #[serde(rename = "@w:ilvl")]
    level: i32,
    /// 起始编号
    #[serde(rename = "w:start", skip_serializing_if = "Option::is_none")]
    start: Option<Val<i32>>,
    /// 编号格式
    #[serde(rename = "w:numFmt", skip_serializing_if = "Option::is_none")]
    format: Option<RawNumFmt>,
    /// 级别文本
    #[serde(rename = "w:lvlText", skip_serializing_if = "Option::is_none")]
    level_text: Option<Val<String>>,
}

/// 原始抽象编号结构体
#[derive(Debug, Deserialize, Serialize)]
struct RawAbstractNum {
    /// 抽象编号ID
    #[serde(rename = "@w:abstractNumId")]
    id: i32,
    /// 级别列表
    #[serde(rename = "w:lvl", default, skip_serializing_if = "Vec::is_empty")]
    levels: Vec<RawLevel>,
}

/// 原始编号结构体
#[derive(Debug, Deserialize, Serialize)]
struct RawNum {
    /// 编号ID
    #[serde(rename = "@w:numId")]
    id: i32,
    /// 抽象编号ID
    #[serde(rename = "w:abstractNumId")]
    abstract_num_id: Val<i32>,
}

/// 编号结构体，表示文档中的编号定义
#[derive(Debug, Default, Deserialize, Serialize)]
#[serde(rename = "w:numbering")]
pub struct Numbering {
    /// 抽象编号列表
    #[serde(rename = "w:abstractNum", default, skip_serializing_if = "Vec::is_empty")]
    abstract_nums: Vec<RawAbstractNum>,
    /// 编号列表
    #[serde(rename = "w:num", default, skip_serializing_if = "Vec::is_empty")]
    nums: Vec<RawNum>,
}

// --- 面向公众的结构体和实现 ---

impl Numbering {
    /// 从XML内容解析编号定义
    /// 
    /// # 参数
    /// * `xml_content` - XML格式的编号定义内容
    pub fn from_xml(xml_content: &str) -> crate::error::Result<Self> {
        // 如果XML内容为空，返回默认编号定义
        if xml_content.is_empty() {
            return Ok(Numbering::default());
        }

        // 反序列化XML内容
        quick_xml::de::from_str(xml_content).map_err(Into::into)
    }

    /// 获取指定编号和级别的文本
    /// 
    /// # 参数
    /// * `num_id` - 编号ID
    /// * `level` - 级别
    /// * `count` - 计数
    pub fn get_level_text(&self, num_id: i32, level: i32, count: i32) -> Option<String> {
        // 查找编号定义
        let num = self.nums.iter().find(|n| n.id == num_id)?;
        // 查找抽象编号定义
        let abstract_num = self
            .abstract_nums
            .iter()
            .find(|an| an.id == num.abstract_num_id.val)?;

        // 查找级别定义
        let level_def = abstract_num.levels.iter().find(|l| l.level == level)?;

        // 获取格式
        let format = level_def.format.as_ref()?.val;
        // 根据格式生成文本
        match format {
            // 项目符号
            NumFmtVal::Bullet => level_def.level_text.as_ref().map(|lt| lt.val.clone()),
            // 十进制数字
            NumFmtVal::Decimal => {
                let text_format = level_def.level_text.as_ref()?.val.clone();
                Some(text_format.replace(&format!("%{}", level + 1), &count.to_string()))
            }
            // 其他格式暂未实现
            _ => None,
        }
    }
}