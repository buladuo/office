# Office Documents in Rust

This library provides a simple way to read and write Microsoft Office documents in Rust.

## Supported Formats

*   `.docx` (Word documents)
*   `.xlsx` (Excel spreadsheets)
*   `.pptx` (PowerPoint presentations)

## Usage

### Reading a `.docx` file

```rust
use office::docx::Docx;
use std::path::Path;

fn main() {
    let path = Path::new("test.docx");
    match Docx::open(path) {
        Ok(docx) => {
            // ... process the docx file
        }
        Err(e) => eprintln!("Error reading docx file: {}", e),
    }
}
```

### Writing a `.docx` file

```rust
use office::docx::elements::{Paragraph, Run, RunContent};
use office::docx::Docx;
use std::path::Path;

fn main() {
    let mut docx = Docx::default();
    let para = Paragraph {
        content: vec![Run {
            content: vec![RunContent::Text("Hello, world!".to_string())],
            ..Default::default()
        }
        .into()],
        ..Default::default()
    };
    docx.document.body.content.push(para.into());

    let path = Path::new("test_write.docx");
    if let Err(e) = docx.save(path) {
        eprintln!("Failed to save docx file: {}", e);
    } else {
        println!("Docx file saved to: {:?}", path);
    }
}
```
