use docx_parse::{ParseError, S9ParseOptions, parse_docx_s5_projection, parse_docx_s9_wire};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn package(paragraph: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(r#"<w:document xmlns:w="{W}"><w:body><w:p>{paragraph}</w:p><w:sectPr/></w:body></w:document>"#)
                .into_bytes(),
        ),
    ])
    .unwrap()
}

fn runs(count: usize) -> String {
    "<w:r><w:t>a</w:t></w:r>".repeat(count)
}

fn heavy_scope(inner: &str) -> String {
    let attribute = "A".repeat(64 * 1024);
    format!(r#"<w:dir w:val="rtl" w:x="{attribute}">{inner}</w:dir>"#)
}

fn refused(error: &Option<ParseError>) -> bool {
    matches!(
        error,
        Some(ParseError::ResourceLimit {
            kind: "directionScopeBytes",
            ..
        })
    )
}

#[test]
fn direction_data_repeated_on_children_past_the_budget_fails_the_open() {
    let bytes = package(&heavy_scope(&runs(1_100)));
    let error = parse_docx_s9_wire(&bytes, S9ParseOptions::default()).err();
    assert!(refused(&error), "{error:?}");
}

#[test]
fn direction_data_repeated_on_hyperlink_children_past_the_budget_fails_the_projection() {
    let link = format!(
        r#"<w:hyperlink w:anchor="target">{}</w:hyperlink>"#,
        heavy_scope(&runs(1_100))
    );
    let error = parse_docx_s5_projection(&package(&link)).err();
    assert!(refused(&error), "{error:?}");
}

#[test]
fn real_shaped_nested_scopes_over_many_runs_open_as_before() {
    let paragraph = format!(
        r#"<w:dir w:val="rtl"><w:bdo w:val="ltr">{}</w:bdo></w:dir>"#,
        runs(2_000)
    );
    let bytes = package(&paragraph);
    let wire = parse_docx_s9_wire(&bytes, S9ParseOptions::default()).unwrap();
    let json = serde_json::to_string(&wire).unwrap();
    assert_eq!(json.matches(r#""kind":"bdo""#).count(), 2_000);
    assert!(parse_docx_s5_projection(&bytes).is_ok());
}
