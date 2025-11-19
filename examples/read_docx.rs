use office::docx::elements::{BodyContent, ParagraphContent, RunContent};
use office::docx::Docx;
use std::path::Path;

fn main() {
    let docx_path = "test.docx";
    if let Err(e) = create_dummy_docx(docx_path) {
        eprintln!("Failed to create dummy docx file: {}", e);
        return;
    }

    println!("-- Reading docx file: {} --", docx_path);

    match Docx::open(Path::new(docx_path)) {
        Ok(docx) => {
            for content in docx.document.body.content.iter() {
                match content {
                    BodyContent::Paragraph(para) => {
                        let para_text = get_paragraph_text(&para, &docx);
                        println!("{}", para_text);
                    }
                    BodyContent::Table(table) => {
                        println!("-- Table --");
                        for row in &table.rows {
                            for cell in &row.cells {
                                let cell_text: String = cell
                                    .content
                                    .iter()
                                    .map(|c| match c {
                                        BodyContent::Paragraph(p) => get_paragraph_text(p, &docx),
                                        BodyContent::Table(_) => "[Nested Table]".to_string(),
                                    })
                                    .collect::<Vec<_>>()
                                    .join("\n");
                                print!("[{}]\t", cell_text);
                            }
                            println!();
                        }
                        println!("-- End Table --");
                    }
                }
            }
        }
        Err(e) => eprintln!("Error reading docx file: {}", e),
    }

    let _ = std::fs::remove_file(docx_path);
    println!("\n-- Dummy file removed. --");
}

fn get_paragraph_text(para: &office::docx::elements::Paragraph, docx: &Docx) -> String {
    let mut text = String::new();
    for item in &para.content {
        match item {
            ParagraphContent::Run(run) => {
                for rc in &run.content {
                    match rc {
                        RunContent::Text(t) => text.push_str(t),
                        RunContent::Break => text.push('\n'),
                        RunContent::Tab => text.push('\t'),
                    }
                }
            }
            ParagraphContent::Hyperlink(hyperlink) => {
                let link_text: String = hyperlink
                    .runs
                    .iter()
                    .flat_map(|r| &r.content)
                    .map(|rc| match rc {
                        RunContent::Text(t) => t.as_str(),
                        _ => "",
                    })
                    .collect();

                let url = docx
                    .relationships
                    .as_ref()
                    .and_then(|rels| rels.get_target(&hyperlink.r_id))
                    .map(|s| s.as_str())
                    .unwrap_or("");

                text.push_str(&format!("[{}]({})", link_text, url));
            }
        }
    }
    text
}

fn create_dummy_docx(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    use std::fs::File;
    use std::io::Write;
    use zip::write::{FileOptions, ZipWriter};

    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options: FileOptions<'static, ()> = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(br###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"###)?;

    zip.add_directory("_rels", options)?;
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(br###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"###)?;

    zip.add_directory("word", options)?;
    zip.add_directory("word/_rels", options)?;
    zip.start_file("word/_rels/document.xml.rels", options)?;
    zip.write_all(br###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com" TargetMode="External"/>
</Relationships>"###)?;

    zip.start_file("word/document.xml", options)?;
    zip.write_all(br###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r><w:t>Here is a </w:t></w:r>
            <w:hyperlink r:id="rId1">
                <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>hyperlink</w:t></w:r>
            </w:hyperlink>
            <w:r><w:t>.</w:t></w:r>
        </w:p>
        <w:tbl>
            <w:tr><w:tc><w:p><w:r><w:t>Cell A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell B1</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
    </w:body>
</w:document>"###)?;

    zip.start_file("word/styles.xml", options)?;
    zip.write_all(br###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:styles>"###)?;

    zip.finish()?;
    Ok(())
}