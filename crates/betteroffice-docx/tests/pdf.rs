use betteroffice_docx::Document;

#[test]
fn exports_a_docx_as_a_paginated_pdf_with_embedded_fonts() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/corpus/fixtures/betteroffice-demo.docx");
    let bytes = std::fs::read(path).expect("demo docx");
    let document = Document::open(&bytes).expect("open");
    let pdf = document.export_pdf().expect("export pdf");
    let text = String::from_utf8_lossy(&pdf);

    assert!(text.starts_with("%PDF-1."));
    assert!(
        text.contains("/Count 2"),
        "demo document paginates to two pages"
    );
    assert!(text.contains("/MediaBox [0 0 612 792]"));
    assert!(text.contains("/FontFile2"), "outlines must be embedded");
    assert!(text.contains("/Identity-H"));
    assert!(text.contains("/ToUnicode"));
    assert!(
        !text.contains("/BaseFont /Helvetica"),
        "glyph runs need no Base-14 fallback"
    );
}
