use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

fn start(id: u32) -> String {
    format!(r#"<w:bookmarkStart w:id="{id}" w:name="Mark{id}"/>"#)
}

fn end(id: u32) -> String {
    format!(r#"<w:bookmarkEnd w:id="{id}"/>"#)
}

fn cell(text: &str) -> String {
    format!(r#"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"#)
}

/// Bookmarks before the table properties, between rows, after the last row,
/// before the first cell, between cells and after the last cell.
fn table() -> String {
    format!(
        r#"<w:tbl>{}<w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>{}<w:tr>{}{}{}{}</w:tr>{}<w:tr>{}{}{}</w:tr>{}</w:tbl>"#,
        start(1),
        start(2),
        start(3),
        cell("A"),
        end(3),
        cell("B"),
        end(2),
        cell("C"),
        start(4),
        cell("D"),
        end(4) + &end(1),
    )
}

fn fixture(body: &str, header: Option<&str>) -> Vec<u8> {
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
    ];
    let mut section = String::new();
    if let Some(header) = header {
        section = r#"<w:headerReference w:type="default" r:id="header"/>"#.to_owned();
        parts.push(("word/_rels/document.xml.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned().into_bytes()));
        parts.push((
            "word/header1.xml".to_owned(),
            format!("<w:hdr {NS}>{header}<w:p/></w:hdr>").into_bytes(),
        ));
    }
    parts.push((
        "word/document.xml".to_owned(),
        format!("<w:document {NS}><w:body>{body}<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr>{section}</w:sectPr></w:body></w:document>").into_bytes(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn package(source: &[u8]) -> Value {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    package
}

fn save(source: &[u8], package: &Value) -> Result<Vec<u8>, docx_parse::ParseError> {
    let mut request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    });
    for key in ["headerEntries", "footerEntries", "footnotes", "endnotes"] {
        if let Some(value) = package.get(key) {
            request[key] = value.clone();
        }
    }
    let request: S13SaveRequest = serde_json::from_value(request).unwrap();
    write_docx_s13(request, source)
}

fn part(saved: &[u8], name: &str) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(saved)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == name)
            .unwrap()
            .1,
    )
    .unwrap()
}

fn assert_order(xml: &str, fragments: &[String]) {
    let mut offset = 0;
    for fragment in fragments {
        if fragment.starts_with("<w:bookmark") {
            assert_eq!(
                xml.matches(fragment.as_str()).count(),
                1,
                "{fragment}: {xml}"
            );
        }
        let at = xml[offset..]
            .find(fragment.as_str())
            .unwrap_or_else(|| panic!("{fragment} out of order: {xml}"));
        offset += at + fragment.len();
    }
}

fn expected_order() -> Vec<String> {
    [
        start(1),
        start(2),
        "<w:tr>".to_owned(),
        start(3),
        ">A<".to_owned(),
        end(3),
        ">B<".to_owned(),
        "</w:tr>".to_owned(),
        end(2),
        ">C<".to_owned(),
        start(4),
        ">D<".to_owned(),
        "</w:tr>".to_owned(),
        end(4),
        end(1),
        "</w:tbl>".to_owned(),
    ]
    .into()
}

#[test]
fn table_and_row_bookmarks_survive_a_save_in_place() {
    for (name, body) in [
        ("body", table()),
        (
            "control",
            format!(
                r#"<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
                table()
            ),
        ),
    ] {
        let source = fixture(&body, None);
        let xml = part(
            &save(&source, &package(&source)).unwrap(),
            "word/document.xml",
        );
        assert_order(&xml, &expected_order());
        assert!(
            xml.find(&start(1)).unwrap() > xml.find("</w:tblGrid>").unwrap(),
            "{name}: {xml}"
        );
    }
}

#[test]
fn header_table_bookmarks_survive_a_save() {
    let source = fixture("", Some(&table()));
    let xml = part(
        &save(&source, &package(&source)).unwrap(),
        "word/header1.xml",
    );
    assert_order(&xml, &expected_order());
}

#[test]
fn table_bookmark_fields_reject_other_xml() {
    let source = fixture(&table(), None);
    for (path, xml) in [
        ("/document/content/0/rows/0/leadingBookmarks/0", "<w:p/>"),
        (
            "/document/content/0/rows/0/cells/0/leadingBookmarks/0",
            "<w:bookmarkStart w:id=\"3\"",
        ),
        (
            "/document/content/0/trailingBookmarks/0",
            "<x:bookmarkEnd xmlns:x=\"urn:x\"/>",
        ),
    ] {
        let mut model = package(&source);
        *model.pointer_mut(path).unwrap() = json!(xml);
        assert!(save(&source, &model).is_err(), "{path}: {xml}");
    }
}

#[test]
fn table_bookmarks_round_trip_through_the_model() {
    let source = fixture(&table(), None);
    let model = package(&source);
    let table = &model["document"]["content"][0];
    assert_eq!(
        table["rows"][0]["leadingBookmarks"],
        json!([start(1), start(2)])
    );
    assert_eq!(
        table["rows"][0]["cells"][0]["leadingBookmarks"],
        json!([start(3)])
    );
    assert_eq!(
        table["rows"][0]["cells"][1]["leadingBookmarks"],
        json!([end(3)])
    );
    assert_eq!(table["rows"][1]["leadingBookmarks"], json!([end(2)]));
    assert_eq!(
        table["rows"][1]["cells"][1]["leadingBookmarks"],
        json!([start(4)])
    );
    assert_eq!(table["trailingBookmarks"], json!([end(4), end(1)]));
    assert!(table["rows"][0].get("trailingBookmarks").is_none());
}
