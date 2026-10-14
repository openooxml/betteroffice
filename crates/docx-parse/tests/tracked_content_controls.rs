use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REVISIONS: [(&str, &str); 4] = [
    ("ins", "insertion"),
    ("del", "deletion"),
    ("moveFrom", "moveFrom"),
    ("moveTo", "moveTo"),
];

fn run(text: &str, deleted: bool) -> String {
    let tag = if deleted { "delText" } else { "t" };
    format!("<w:r><w:{tag}>{text}</w:{tag}></w:r>")
}

fn control(content: &str) -> String {
    format!(
        r#"<w:sdt><w:sdtPr><w:tag w:val="tracked"/><w:alias w:val="Tracked control"/><w:id w:val="42"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"#
    )
}

fn tracked(tag: &str, content: &str) -> String {
    format!(r#"<w:{tag} w:id="5" w:author="A" w:date="2024-01-01T00:00:00Z">{content}</w:{tag}>"#)
}

fn package(content: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="{W}"><w:body><w:p>{content}</w:p></w:body></w:document>"#).into_bytes()),
    ]).unwrap()
}

fn parse(bytes: &[u8]) -> Value {
    let document = parse_docx_s9_wire(bytes, S9ParseOptions::default()).unwrap();
    serde_json::to_value(document.document.package.document.content).unwrap()
}

fn assert_content(model: &Value, kind: &str, outer_control: bool, mixed: bool) {
    let outer = &model[0]["content"][0];
    let (sdt, revision) = if outer_control {
        (outer, &outer["content"][if mixed { 1 } else { 0 }])
    } else {
        (&outer["content"][0], outer)
    };
    assert_eq!(sdt["type"], "inlineSdt");
    assert_eq!(sdt["properties"]["tag"], "tracked");
    assert_eq!(sdt["properties"]["alias"], "Tracked control");
    assert_eq!(sdt["properties"]["id"].as_f64(), Some(42.0));
    assert_eq!(revision["type"], kind);
    assert_eq!(
        revision["info"],
        json!({"id": 5.0, "author": "A", "date": "2024-01-01T00:00:00Z"})
    );
    let run = if outer_control {
        &revision["content"][0]
    } else {
        &sdt["content"][0]
    };
    assert_eq!(run["content"][0]["text"], "Inserted");
    if mixed {
        assert_eq!(sdt["content"][0]["content"][0]["text"], "A");
        assert_eq!(sdt["content"][2]["content"][0]["text"], "C");
    }
}

fn round_trip(tag: &str, kind: &str, outer_control: bool, mixed: bool) {
    let text = run("Inserted", matches!(tag, "del" | "moveFrom"));
    let content = if outer_control {
        control(&format!(
            "{}{}{}",
            if mixed {
                run("A", false)
            } else {
                String::new()
            },
            tracked(tag, &text),
            if mixed {
                run("C", false)
            } else {
                String::new()
            }
        ))
    } else {
        tracked(tag, &control(&text))
    };
    let original = package(&content);
    let model = parse(&original);
    assert_content(&model, kind, outer_control, mixed);
    let saved = repack(&original, model);
    assert_content(&parse(&saved), kind, outer_control, mixed);
    if matches!(tag, "del" | "moveFrom") {
        let xml = ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1;
        assert!(
            String::from_utf8(xml)
                .unwrap()
                .contains("<w:delText>Inserted</w:delText>")
        );
    }
}

fn repack(original: &[u8], model: Value) -> Vec<u8> {
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": model},
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    write_docx_s13(request, original).unwrap()
}

#[test]
fn revisions_inside_inline_controls_parse_and_repack() {
    for (tag, kind) in REVISIONS {
        round_trip(tag, kind, true, false);
    }
}

#[test]
fn inline_controls_inside_revisions_parse_and_repack() {
    for (tag, kind) in REVISIONS {
        round_trip(tag, kind, false, false);
    }
}

#[test]
fn mixed_control_content_preserves_order_and_revision() {
    round_trip("ins", "insertion", true, true);
}

#[test]
fn tracked_control_children_preserve_fields_and_math() {
    let content = format!(
        "{}{}{}",
        run("Control", false),
        r#"<w:fldSimple w:instr="PAGE"><w:r><w:t>Field</w:t></w:r></w:fldSimple>"#,
        r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>Equation</m:t></m:r></m:oMath>"#
    );
    let original = package(&tracked("ins", &control(&content)));
    let model = parse(&original);
    let reopened = parse(&repack(&original, model.clone()));
    for parsed in [&model, &reopened] {
        let children = &parsed[0]["content"][0]["content"];
        assert_eq!(children[0]["type"], "inlineSdt");
        let content = &children[0]["content"];
        assert_eq!(content[0]["content"][0]["text"], "Control");
        assert_eq!(content[1]["type"], "simpleField");
        assert_eq!(content[1]["content"][0]["content"][0]["text"], "Field");
        assert_eq!(content[2]["type"], "mathEquation");
        assert_eq!(content[2]["plainText"], "Equation");
    }
}

#[test]
fn ordinary_tracked_children_keep_legacy_output() {
    let content = format!(
        "{}{}{}",
        run("Deleted", true),
        r#"<w:hyperlink w:anchor="target"><w:r><w:t>Linked</w:t></w:r></w:hyperlink>"#,
        r#"<w:fldSimple w:instr="PAGE"><w:r><w:t>Field</w:t></w:r></w:fldSimple>"#,
    );
    let original = package(&tracked("del", &content));
    let model = parse(&original);
    let children = model[0]["content"][0]["content"].as_array().unwrap();
    assert_eq!(children.len(), 2);
    assert_eq!(children[0]["type"], "run");
    assert_eq!(children[1]["type"], "hyperlink");
    let saved = repack(&original, model);
    let xml = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap()
        .1;
    let xml = String::from_utf8(xml).unwrap();
    assert!(xml.contains("<w:delText>Deleted</w:delText>"));
    assert!(
        xml.contains(
            r#"<w:hyperlink w:anchor="target"><w:r><w:t>Linked</w:t></w:r></w:hyperlink>"#
        )
    );
    assert!(!xml.contains("<w:fldSimple"));
}
