use office::docx::Docx;
use std::path::Path;

fn main() {
    // This example creates a dummy docx file in memory, then parses it.
    let docx_path = "test.docx";
    if let Err(e) = create_dummy_docx(docx_path) {
        eprintln!("Failed to create dummy docx file: {}", e);
        return;
    }

    println!("--- Reading docx file: {} ---", docx_path);

    match Docx::open(Path::new(docx_path)) {
        Ok(docx) => {
            for para in docx.document.paragraphs {
                let text: String = para
                    .runs
                    .iter()
                    .map(|r| r.text.as_deref().unwrap_or(""))
                    .collect();
                println!("{}", text);
            }
        }
        Err(e) => eprintln!("Error reading docx file: {}", e),
    }

    // Clean up the dummy file
    let _ = std::fs::remove_file(docx_path);
    println!("\n--- Dummy file removed. ---");
}

/// Creates a minimal, valid docx file for testing purposes.
fn create_dummy_docx(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    use std::fs::File;
    use std::io::Write;
    use zip::write::{FileOptions, ZipWriter};

    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);

    let options: FileOptions<'static, ()> = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    // [Content_Types].xml - Defines the content types of the parts in the package.
    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#)?;

    // _rels/.rels - The package-level relationships.
    zip.add_directory("_rels", options)?;
    zip.start_file("_rels/.rels", options)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#)?;

    // word/document.xml - The main document content.
    zip.add_directory("word", options)?;
    zip.start_file("word/document.xml", options)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello, World!</w:t></w:r></w:p>
        <w:p><w:r><w:t>This is a test document.</w:t></w:r></w:p>
        <w:p/>
        <w:p><w:r><w:t>The third paragraph.</w:t></w:r></w:p>
    </w:body>
</w:document>"#)?;

    // word/styles.xml - A minimal styles part.
    zip.start_file("word/styles.xml", options)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:styles>"#)?;

    zip.finish()?;
    Ok(())
}
