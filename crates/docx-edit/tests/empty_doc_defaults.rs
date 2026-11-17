use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EngineSession, seed_from_docx};
use serde_json::Value;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// A package whose styles part is `styles` and whose body is one unstyled
/// paragraph and one table cell, neither naming a style.
fn package(styles: &str) -> Vec<u8> {
    let body = r#"<w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>"#;
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
        ("word/styles.xml", styles.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// The spacing of the body paragraph and of the table cell's paragraph.
fn spacing(bytes: &[u8]) -> [Value; 2] {
    let engine = EngineSession::new(75406);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let blocks = serde_json::to_value(
        yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default()).unwrap(),
    )
    .unwrap();
    let blocks = blocks.as_array().unwrap();
    let body = blocks
        .iter()
        .find(|block| block["kind"] == "paragraph")
        .unwrap();
    let table = blocks
        .iter()
        .find(|block| block["kind"] == "table")
        .unwrap();
    [
        body["attrs"]["spacing"].clone(),
        table["rows"][0]["cells"][0]["blocks"][0]["attrs"]["spacing"].clone(),
    ]
}

#[test]
fn an_empty_doc_defaults_element_leaves_unstyled_paragraphs_without_spacing() {
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style></w:styles>"#
    );
    for spacing in spacing(&package(&styles)) {
        assert!(
            spacing.get("after").is_none_or(|after| after == 0.0),
            "{spacing}"
        );
        assert!(
            spacing.get("line").is_none_or(|line| line == 1.0),
            "{spacing}"
        );
    }
}

#[test]
fn a_styles_part_without_doc_defaults_or_normal_keeps_words_normal_fallback() {
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>"#
    );
    for spacing in spacing(&package(&styles)) {
        let after = spacing["after"].as_f64().unwrap();
        assert!((after - 160.0 / 15.0).abs() < 1e-6, "{spacing}");
        let line = spacing["line"].as_f64().unwrap();
        assert!((line - 259.0 / 240.0).abs() < 1e-9, "{spacing}");
    }
}
