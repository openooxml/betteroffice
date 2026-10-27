use std::sync::atomic::{AtomicUsize, Ordering};

use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::json;

static CALLS: AtomicUsize = AtomicUsize::new(0);

fn count() {
    CALLS.fetch_add(1, Ordering::Relaxed);
}

fn package(body: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#).into_bytes()),
    ]).unwrap()
}

#[test]
fn long_parse_and_save_report_progress() {
    let source = package(&"<w:p><w:r><w:t>Line</w:t></w:r></w:p>".repeat(3000));
    ooxml_opc::progress::set_hook(Some(count));

    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let blocks = 3000 / ooxml_opc::progress::EVERY as usize;
    assert!(CALLS.swap(0, Ordering::Relaxed) >= blocks, "parse progress");

    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {
            "content": parsed.document.package.document.content,
            "finalSectionProperties": parsed.document.package.document.final_section_properties,
        },
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    write_docx_s13(request, &source).unwrap();
    assert!(CALLS.load(Ordering::Relaxed) >= blocks, "save progress");
}
