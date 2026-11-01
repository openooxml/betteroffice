use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

fn fixture(body: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_document(source: &[u8]) -> String {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, source).unwrap();
    String::from_utf8(
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

fn assert_order(xml: &str, fragments: &[&str]) {
    let mut offset = 0;
    for fragment in fragments {
        let at = xml[offset..]
            .find(fragment)
            .unwrap_or_else(|| panic!("{fragment} missing or out of order: {xml}"));
        offset += at + fragment.len();
    }
}

const BEGIN: &str = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#;
const SEPARATE: &str = r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#;
const END: &str = r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
const START: &str = r#"<w:bookmarkStart w:id="0" w:name="Text1"/>"#;
const STOP: &str = r#"<w:bookmarkEnd w:id="0"/>"#;

fn instruction(text: &str) -> String {
    format!(r#"<w:r><w:instrText xml:space="preserve"> {text} </w:instrText></w:r>"#)
}

#[test]
fn a_form_field_bookmark_in_the_field_code_survives_a_save() {
    let body = format!(
        r#"<w:p>{BEGIN}{START}{}{SEPARATE}<w:r><w:t>Value</w:t></w:r>{END}{STOP}</w:p>"#,
        instruction("FORMTEXT")
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "w:fldCharType=\"begin\"",
            START,
            "FORMTEXT",
            "w:fldCharType=\"separate\"",
            ">Value<",
            "w:fldCharType=\"end\"",
            STOP,
        ],
    );
}

#[test]
fn bookmarks_on_both_sides_of_the_field_separator_survive_a_save() {
    let body = format!(
        r#"<w:p>{BEGIN}{}{START}{SEPARATE}<w:r><w:t>Value</w:t></w:r>{STOP}{END}</w:p>"#,
        instruction("REF Target")
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "REF Target",
            START,
            "w:fldCharType=\"separate\"",
            ">Value<",
            STOP,
            "w:fldCharType=\"end\"",
        ],
    );
}

#[test]
fn a_nested_field_in_the_field_code_survives_a_save() {
    let nested = format!(
        r#"{BEGIN}{}{SEPARATE}<w:r><w:t>inner</w:t></w:r>{END}"#,
        instruction("MERGEFIELD Name")
    );
    let body = format!(
        r#"<w:p>{BEGIN}{}{nested}{}{SEPARATE}<w:r><w:t>outer</w:t></w:r>{END}</w:p>"#,
        instruction("IF"),
        instruction("= \"x\" \"y\" \"z\""),
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "IF",
            "MERGEFIELD Name",
            ">inner<",
            "&quot;y&quot;",
            ">outer<",
        ],
    );
    assert_eq!(xml.matches("w:fldCharType=\"begin\"").count(), 2, "{xml}");
}

#[test]
fn a_simple_field_in_the_field_code_is_written_once() {
    let body = format!(
        r#"<w:p>{BEGIN}{}<w:fldSimple w:instr="AUTHOR"><w:r><w:t>Ada</w:t></w:r></w:fldSimple>{START}{SEPARATE}<w:r><w:t>Ada</w:t></w:r>{END}{STOP}</w:p>"#,
        instruction("QUOTE")
    );
    let xml = saved_document(&fixture(&body));
    assert_eq!(xml.matches(">Ada<").count(), 2, "{xml}");
    assert_order(
        &xml,
        &[
            "QUOTE",
            "w:instr=\"AUTHOR\"",
            ">Ada<",
            START,
            "\"separate\"",
            ">Ada<",
            STOP,
        ],
    );
}

#[test]
fn an_edited_field_code_drops_its_old_nested_fields_and_keeps_its_bookmarks() {
    let nested = format!(
        r#"{BEGIN}{}{SEPARATE}<w:r><w:t>2</w:t></w:r>{END}"#,
        instruction("PAGE")
    );
    let body = format!(
        r#"<w:p>{BEGIN}{START}{}{nested}{}{SEPARATE}<w:r><w:t>2</w:t></w:r>{END}{STOP}</w:p>"#,
        instruction("="),
        instruction("/2+1"),
    );
    let source = fixture(&body);
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let content = package["document"]["content"][0]["content"]
        .as_array_mut()
        .unwrap();
    let field = content
        .iter_mut()
        .find(|node| node["type"] == "complexField")
        .unwrap();
    field["instruction"] = json!("NUMPAGES");
    field["fieldCode"] = json!([]);
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, &source).unwrap();
    let xml = String::from_utf8(
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(!xml.contains("PAGE<") && !xml.contains(" PAGE "), "{xml}");
    assert_order(
        &xml,
        &[
            "\"begin\"",
            START,
            "NUMPAGES",
            "\"separate\"",
            "\"end\"",
            STOP,
        ],
    );
}

#[test]
fn a_nested_field_after_an_instruction_in_the_begin_run_survives_a_save() {
    let nested = format!(
        r#"{BEGIN}{}{SEPARATE}<w:r><w:t>4</w:t></w:r>{END}"#,
        instruction("PAGE")
    );
    let body = format!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> = </w:instrText></w:r>{nested}{SEPARATE}<w:r><w:t>4</w:t></w:r>{END}</w:p>"#
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "\"begin\"",
            " = ",
            "\"begin\"",
            " PAGE ",
            "\"separate\"",
            "\"end\"",
            "\"separate\"",
            "\"end\"",
        ],
    );
    assert_eq!(xml.matches("w:fldCharType=\"begin\"").count(), 2, "{xml}");
}
