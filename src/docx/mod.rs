use crate::common::package::open_package;
use crate::common::relations::Relationships;
use crate::docprops::{AppProps, CoreProps};
use crate::error::Result;
use quick_xml::se::to_string;
use std::io::{Cursor, Write};
use std::path::Path;
use zip::write::{FileOptions, ZipWriter};

/// 文档模块
pub mod document;
/// 元素模块
pub mod elements;
/// 编号模块
pub mod numbering;
/// 属性模块
pub mod properties;
/// 样式模块
pub mod styles;

use document::Document;
pub use elements::{BodyContent, Paragraph, ParagraphContent, Run, RunContent};
use numbering::Numbering;
use styles::Styles;

/// DOCX文档结构体，表示整个DOCX文件
#[derive(Debug, Default)]
pub struct Docx {
    /// 文档主体内容
    pub document: Document,
    /// 样式定义
    pub styles: Styles,
    /// 编号定义
    pub numbering: Numbering,
    /// 文档关系
    pub relationships: Option<Relationships>,
    /// 应用程序属性
    pub app_props: Option<AppProps>,
    /// 核心属性
    pub core_props: Option<CoreProps>,
}

impl Docx {
    /// 打开并解析DOCX文件
    /// 
    /// # 参数
    /// * `path` - DOCX文件路径
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        // 打开DOCX包
        let mut package = open_package(path)?;

        // 读取应用程序属性
        let app_props = package
            .read_file_by_path("docProps/app.xml")
            .ok()
            .and_then(|content| AppProps::from_xml(&content).ok());

        // 读取核心属性
        let core_props = package
            .read_file_by_path("docProps/core.xml")
            .ok()
            .and_then(|content| CoreProps::from_xml(&content).ok());

        // 读取文档关系
        let rels_content = package
            .read_file_by_path("word/_rels/document.xml.rels")
            .unwrap_or_default();
        let relationships = Some(Relationships::from_xml(&rels_content)?);

        // 读取样式定义
        let styles_content = package.read_file_by_path("word/styles.xml")?;
        let styles = Styles::from_xml(&styles_content)?;

        // 读取编号定义
        let numbering_content = package
            .read_file_by_path("word/numbering.xml")
            .unwrap_or_default();
        let numbering = Numbering::from_xml(&numbering_content)?;

        // 读取主文档内容
        let document_content = package.read_file_by_path("word/document.xml")?;
        let document = Document::from_xml(&document_content, relationships.as_ref())?;

        // 构造并返回DOCX对象
        Ok(Docx {
            document,
            styles,
            numbering,
            relationships,
            app_props,
            core_props,
        })
    }

    /// 保存DOCX文件
    /// 
    /// # 参数
    /// * `path` - 保存路径
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        // 创建文件
        let file = std::fs::File::create(path)?;
        // 创建ZIP写入器
        let mut zip = ZipWriter::new(file);
        // 设置文件选项
        let options: FileOptions<'static, ()> =
            FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        // 写入[Content_Types].xml文件
        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.themeManager+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"#)?;

        // 写入_rels/.rels文件
        zip.add_directory("_rels", options)?;
        zip.start_file("_rels/.rels", options)?;
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="word/theme/theme1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="word/fontTable.xml"/>
    <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="word/settings.xml"/>
</Relationships>"#)?;

        // 写入docProps/app.xml文件
        if let Some(app_props) = &self.app_props {
            zip.add_directory("docProps", options)?;
            zip.start_file("docProps/app.xml", options)?;
            let app_props_xml = to_string(app_props)?;
            zip.write_all(app_props_xml.as_bytes())?;
        }

        // 写入docProps/core.xml文件
        if let Some(core_props) = &self.core_props {
            zip.start_file("docProps/core.xml", options)?;
            let core_props_xml = to_string(core_props)?;
            zip.write_all(core_props_xml.as_bytes())?;
        }

        // 写入word/document.xml文件
        zip.add_directory("word", options)?;
        zip.start_file("word/document.xml", options)?;
        let document_xml = to_string(&self.document)?;
        zip.write_all(document_xml.as_bytes())?;

        // 写入word/styles.xml文件
        zip.start_file("word/styles.xml", options)?;
        let styles_xml = to_string(&self.styles)?;
        zip.write_all(styles_xml.as_bytes())?;

        // 写入word/numbering.xml文件
        zip.start_file("word/numbering.xml", options)?;
        let numbering_xml = to_string(&self.numbering)?;
        zip.write_all(numbering_xml.as_bytes())?;

        // 写入word/theme/theme1.xml
        zip.add_directory("word/theme", options)?;
        zip.start_file("word/theme/theme1.xml", options)?;
        zip.write_all(DEFAULT_THEME_XML)?;

        // 写入word/fontTable.xml
        zip.start_file("word/fontTable.xml", options)?;
        zip.write_all(DEFAULT_FONT_TABLE_XML)?;

        // 写入word/settings.xml
        zip.start_file("word/settings.xml", options)?;
        zip.write_all(DEFAULT_SETTINGS_XML)?;

        // 写入word/_rels/document.xml.rels文件
        if let Some(rels) = &self.relationships {
            zip.add_directory("word/_rels", options)?;
            zip.start_file("word/_rels/document.xml.rels", options)?;
            let rels_xml = rels.to_xml()?;
            zip.write_all(rels_xml.as_bytes())?;
        }

        // 完成ZIP文件写入
        zip.finish()?;
        Ok(())
    }
}

const DEFAULT_THEME_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:noFill/>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:prstShdw prst="shd5">
              <a:clr>
                <a:srgbClr val="000000"/>
                <a:alpha val="63000"/>
              </a:clr>
            </a:prstShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:noFill/>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="95000"/>
                <a:satMod val="170000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="50000">
              <a:schemeClr val="phClr">
                <a:tint val="93000"/>
                <a:satMod val="150000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="90000"/>
                <a:satMod val="250000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
  <a:extLst>
    <a:ext uri="{05BEAD80-6487-4522-90FC-B00981E7B00E}">
      <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93CD-4838-8197-D165EEAE06A0}" vid="{4FBDCFC4-1987-4678-8A7A-B1A9299EC6DA}"/>
    </a:ext>
  </a:extLst>
</a:theme>"#;

const DEFAULT_FONT_TABLE_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E00002FF" w:usb1="4000ACFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
  </w:font>
</w:fonts>"#;

const DEFAULT_SETTINGS_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:settings>"#;
}