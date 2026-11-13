use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

const TEXT_DATA: &str = r#"<w:ffData><w:name w:val="Text1"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="Your name"/><w:maxLength w:val="20"/></w:textInput></w:ffData>"#;
const CHECK_DATA: &str = r#"<w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox></w:ffData>"#;
const DROP_DATA: &str = r#"<w:ffData><w:name w:val="Drop1"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:result w:val="1"/><w:listEntry w:val="Alpha"/><w:listEntry w:val="Beta"/><w:listEntry w:val="Gamma"/></w:ddList></w:ffData>"#;

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

fn text(value: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{value}</w:t></w:r>"#)
}

fn form_field(data: &str, instruction: &str, result: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin">{data}</w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

fn form_paragraph() -> String {
    format!(
        "<w:p>{}{}{}{}{}{}{}</w:p>",
        text("Name: "),
        form_field(TEXT_DATA, "FORMTEXT", &text("Ada")),
        text(" Agree "),
        form_field(CHECK_DATA, "FORMCHECKBOX", ""),
        text(" Pick "),
        form_field(DROP_DATA, "FORMDROPDOWN", &text("Beta")),
        text(" end."),
    )
}

#[test]
fn keeps_form_field_data_when_a_paragraph_is_written_from_the_model() {
    let xml = saved_document(&fixture(&form_paragraph()));
    let begins: Vec<String> = [TEXT_DATA, CHECK_DATA, DROP_DATA]
        .iter()
        .map(|data| format!(r#"<w:fldChar w:fldCharType="begin">{data}</w:fldChar>"#))
        .collect();
    for begin in &begins {
        assert_eq!(xml.matches(begin.as_str()).count(), 1, "{begin}\n{xml}");
    }
    assert_eq!(xml.matches("<w:ffData>").count(), 3, "{xml}");
    let mut offset = 0;
    for fragment in [
        "Name: ",
        begins[0].as_str(),
        " FORMTEXT ",
        " Agree ",
        begins[1].as_str(),
        " FORMCHECKBOX ",
        " Pick ",
        begins[2].as_str(),
        " FORMDROPDOWN ",
        " end.",
    ] {
        let at = xml[offset..]
            .find(fragment)
            .unwrap_or_else(|| panic!("{fragment} missing or out of order: {xml}"));
        offset += at + fragment.len();
    }
}

#[test]
fn keeps_a_dirty_field_marked_dirty_when_a_paragraph_is_written_from_the_model() {
    let body = format!(
        r#"<w:p>{}<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{}<w:r><w:fldChar w:fldCharType="end"/></w:r>{}</w:p>"#,
        text("Page "),
        text("1"),
        text(" of many"),
    );
    let xml = saved_document(&fixture(&body));
    assert_eq!(
        xml.matches(r#"<w:fldChar w:fldCharType="begin" w:dirty="true"/>"#).count(),
        1,
        "{xml}"
    );
    assert_eq!(xml.matches("w:dirty=").count(), 1, "{xml}");
}
