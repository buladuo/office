use serde::Serialize;

/// 超链接模块
pub mod hyperlink;
/// 段落模块
pub mod paragraph;
/// 文本运行模块
pub mod run;
/// 表格模块
pub mod table;

pub use hyperlink::Hyperlink;
pub use paragraph::{Paragraph, ParagraphContent};
pub use run::{Run, RunContent};
pub use table::{Table, TableCell, TableRow};

/// 文档主体内容枚举，表示文档主体中可能包含的元素类型
#[derive(Debug, Serialize)]
pub enum BodyContent {
    /// 段落
    Paragraph(Paragraph),
    /// 表格
    Table(Table),
}

impl From<Paragraph> for BodyContent {
    fn from(p: Paragraph) -> Self {
        BodyContent::Paragraph(p)
    }
}

impl From<Table> for BodyContent {
    fn from(t: Table) -> Self {
        BodyContent::Table(t)
    }
}