use docx_edit::{EditingDoc, seed_from_docx};

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

fn build_docx(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
        ("_rels/.rels".to_owned(), ROOT_RELS.into()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .expect("synthetic package zips")
}

fn seed_error(body: &str) -> String {
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &build_docx(body)).expect_err("over-budget opaque seed must refuse")
}

#[test]
fn small_opaque_payloads_seed_normally() {
    let body =
        r#"<w:p><w:r><w:object><o:OLEObject Type="Embed" ProgID="Eq"/></w:object></w:r></w:p>"#;
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &build_docx(body)).expect("small opaque seed must succeed");
}

#[test]
fn one_huge_opaque_payload_refuses_to_seed() {
    let filler = "x".repeat(docx_edit::OPAQUE_SEED_BUDGET_BYTES as usize + 1024);
    let body = format!(
        r#"<w:p><w:r><w:object><o:OLEObject Type="Embed" ProgID="Eq"/><w:filler>{filler}</w:filler></w:object></w:r></w:p>"#
    );
    let error = seed_error(&body);
    assert!(
        error.contains("opaque drawing seed budget exceeded"),
        "unexpected seed error: {error}"
    );
}

#[test]
fn many_small_opaque_payloads_refuse_to_seed_in_aggregate() {
    let filler = "y".repeat(120 * 1024);
    let count = docx_edit::OPAQUE_SEED_BUDGET_BYTES as usize / filler.len() + 1;
    let paragraph = format!(r#"<w:p><w:r><w:object>{filler}</w:object></w:r></w:p>"#);
    let error = seed_error(&paragraph.repeat(count));
    assert!(
        error.contains("opaque drawing seed budget exceeded"),
        "unexpected seed error: {error}"
    );
}
