use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

const STYLES: &str = concat!(
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:style w:type="table" w:styleId="HeaderGrid"><w:name w:val="Header Grid"/>"#,
    r#"<w:trPr><w:tblHeader/></w:trPr><w:tblStylePr w:type="firstRow"><w:trPr><w:tblHeader/></w:trPr></w:tblStylePr>"#,
    r#"</w:style></w:styles>"#
);

fn document(body: &str) -> String {
    format!("<w:document {NS}><w:body>{body}</w:body></w:document>")
}

fn table(properties: &str, rows: &[&str]) -> String {
    let rows: String = rows
        .iter()
        .enumerate()
        .map(|(index, row)| {
            let properties = if row.is_empty() {
                String::new()
            } else {
                format!("<w:trPr>{row}</w:trPr>")
            };
            format!(
                "<w:tr>{properties}<w:tc><w:p><w:r><w:t>Row {index}</w:t></w:r></w:p></w:tc></w:tr>"
            )
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr>{properties}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>{rows}</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>"#
    )
}

fn fixture(document: &str, styles: bool) -> Vec<u8> {
    let style_override = if styles {
        r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#
    } else {
        ""
    };
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{style_override}</Types>"#
            )
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.as_bytes().to_vec()),
    ];
    if styles {
        parts.push((
            "word/_rels/document.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_vec(),
        ));
        parts.push(("word/styles.xml".to_owned(), STYLES.as_bytes().to_vec()));
    }
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

fn save(source: &[u8], package: &Value) -> Vec<u8> {
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    write_docx_s13(request, source).unwrap()
}

fn ordinals(value: &mut Value, next: &mut u32) {
    match value {
        Value::Object(object) => {
            if object.get("type").and_then(Value::as_str) == Some("paragraph") {
                object.insert("sourceOrdinal".to_owned(), json!(*next));
                *next += 1;
            }
            for child in object.values_mut() {
                ordinals(child, next);
            }
        }
        Value::Array(values) => {
            for child in values {
                ordinals(child, next);
            }
        }
        _ => {}
    }
}

fn save_unchanged(source: &[u8], xml: &str) -> String {
    let mut content = package(source)["document"]["content"].clone();
    let mut next = 0;
    ordinals(&mut content, &mut next);
    let paragraphs: Vec<u32> = (0..next).collect();
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": content},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/document.xml",
            "sha256": format!("{:x}", Sha256::digest(xml.as_bytes())),
            "paragraphs": paragraphs,
            "changed": [],
            "inserted": [],
            "removed": [],
        }]},
    }))
    .unwrap();
    let saved = write_docx_s13(request, source).unwrap();
    part(&saved, "word/document.xml")
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

fn flag(model: &Value, key: &str) -> Value {
    model["document"]["content"][0]["rows"][0]["formatting"][key].clone()
}

fn written(xml: &str, element: &str) -> Value {
    if xml.contains(&format!("<{element}/>")) {
        return json!(true);
    }
    for (value, state) in [
        ("0", false),
        ("false", false),
        ("off", false),
        ("1", true),
        ("true", true),
        ("on", true),
    ] {
        if xml.contains(&format!(r#"<{element} w:val="{value}"/>"#)) {
            return json!(state);
        }
    }
    Value::Null
}

#[test]
fn row_header_and_cant_split_keep_on_off_and_absent_through_a_save() {
    for (element, key) in [("w:tblHeader", "header"), ("w:cantSplit", "cantSplit")] {
        for (markup, expected) in [
            (format!("<{element}/>"), json!(true)),
            (format!(r#"<{element} w:val="0"/>"#), json!(false)),
            (String::new(), Value::Null),
        ] {
            let source = fixture(&document(&table("", &[&markup])), false);
            let model = package(&source);
            assert_eq!(flag(&model, key), expected, "{markup}");
            let saved = save(&source, &model);
            let xml = part(&saved, "word/document.xml");
            assert_eq!(written(&xml, element), expected, "{xml}");
            assert_eq!(flag(&package(&saved), key), expected, "{xml}");
        }
    }
}

#[test]
fn an_explicit_header_off_under_a_header_row_table_style_survives_a_save() {
    let body = table(
        r#"<w:tblStyle w:val="HeaderGrid"/><w:tblLook w:firstRow="1"/>"#,
        &[r#"<w:tblHeader w:val="0"/>"#, ""],
    );
    let source = fixture(&document(&body), true);
    let model = package(&source);
    assert_eq!(flag(&model, "header"), json!(false));
    let saved = save(&source, &model);
    let xml = part(&saved, "word/document.xml");
    assert_eq!(written(&xml, "w:tblHeader"), json!(false), "{xml}");
    assert_eq!(flag(&package(&saved), "header"), json!(false), "{xml}");
}

#[test]
fn a_save_without_row_edits_keeps_the_document_bytes() {
    let body = table(
        "",
        &[
            "<w:cantSplit/><w:tblHeader/>",
            r#"<w:cantSplit w:val="0"/><w:tblHeader w:val="0"/>"#,
            r#"<w:cantSplit w:val="false"/><w:tblHeader w:val="off"/>"#,
            r#"<w:cantSplit w:val="on"/><w:tblHeader w:val="1"/>"#,
            "",
        ],
    );
    let xml = document(&body);
    assert_eq!(save_unchanged(&fixture(&xml, false), &xml), xml);
}
