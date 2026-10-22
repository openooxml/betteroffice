use std::sync::atomic::{AtomicUsize, Ordering};

use docx_edit::{EditingDoc, seed_from_docx};

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
fn long_open_and_identity_scan_report_progress() {
    let source = package(&"<w:p><w:r><w:t>Line</w:t></w:r></w:p>".repeat(3000));
    ooxml_opc::progress::set_hook(Some(count));
    let blocks = 3000 / ooxml_opc::progress::EVERY as usize;

    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &source).unwrap();
    assert!(
        CALLS.swap(0, Ordering::Relaxed) >= 3 * blocks,
        "open progress"
    );

    assert!(doc.paragraph_identities().paragraphs.len() >= 3000);
    assert!(CALLS.load(Ordering::Relaxed) >= blocks, "identity progress");
}
