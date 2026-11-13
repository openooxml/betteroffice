use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

const TEXT_DATA: &str = r#"<w:ffData><w:name w:val="Text1"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="Your name"/><w:maxLength w:val="20"/></w:textInput></w:ffData>"#;
const CHECK_DATA: &str = r#"<w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox></w:ffData>"#;
const DROP_DATA: &str = r#"<w:ffData><w:name w:val="Drop1"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:result w:val="1"/><w:listEntry w:val="Alpha"/><w:listEntry w:val="Beta"/><w:listEntry w:val="Gamma"/></w:ddList></w:ffData>"#;

fn fixture(body: &str) -> Vec<u8> {
    fixture_in(NS, body)
}

fn fixture_in(namespaces: &str, body: &str) -> Vec<u8> {
    package(format!(
        "<w:document {namespaces}><w:body>{body}<w:sectPr/></w:body></w:document>"
    ))
}

fn package(document: String) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_document(source: &[u8]) -> String {
    save(source).unwrap()
}

fn save(source: &[u8]) -> Result<String, String> {
    save_with(source, Value::Null)
}

fn save_with(source: &[u8], selective: Value) -> Result<String, String> {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let mut request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    });
    if !selective.is_null() {
        request["selective"] = selective;
    }
    let request: S13SaveRequest = serde_json::from_value(request).unwrap();
    let saved = write_docx_s13(request, source).map_err(|error| error.to_string())?;
    Ok(String::from_utf8(
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap())
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

const ALIASED: &str = r#"xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const ALIASED_DATA: &str = r#"<ns0:ffData><ns0:name ns0:val="Ext1"/><ns0:enabled/><ns0:textInput/></ns0:ffData>"#;

fn aliased_document() -> (Vec<u8>, String) {
    let field = format!(
        r#"<w:p w14:paraId="73000003">{}{}</w:p>"#,
        text("Label "),
        form_field(ALIASED_DATA, "FORMTEXT", &text("Ada")),
    );
    let body = format!(r#"{field}<w:p w14:paraId="73000004">{}</w:p>"#, text("Other"));
    (fixture_in(&format!("{NS} {ALIASED}"), &body), field)
}

#[test]
fn refuses_to_rewrite_form_field_data_it_cannot_write_naming_the_field_and_the_reason() {
    let (source, _) = aliased_document();
    let error = save(&source).unwrap_err();
    assert!(error.contains("\"Ext1\""), "{error}");
    assert!(error.contains("found ns0:ffData"), "{error}");
}

#[test]
fn keeps_the_source_bytes_of_an_untouched_paragraph_whose_form_field_data_it_cannot_write() {
    let (source, field) = aliased_document();
    let xml = save_with(&source, json!({"changedParaIds": ["73000004"]})).unwrap();
    assert_eq!(xml.matches(field.as_str()).count(), 1, "{xml}");
}

fn binds_ext_foo(xml: &str) -> bool {
    let mut reader = NsReader::from_str(xml);
    let mut found = false;
    loop {
        let (resolved, event) = reader.read_resolved_event().unwrap();
        let namespace = match resolved {
            ResolveResult::Unknown(prefix) => {
                panic!("unbound prefix {}: {xml}", String::from_utf8_lossy(&prefix))
            }
            ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
            ResolveResult::Unbound => None,
        };
        match event {
            Event::Start(element) | Event::Empty(element) => {
                found |= element.local_name().as_ref() == b"foo"
                    && namespace.as_deref() == Some(b"urn:x".as_slice());
                for attribute in element.attributes() {
                    let key = attribute.unwrap().key;
                    if key.as_ref().starts_with(b"xmlns") || key.as_ref().starts_with(b"xml:") {
                        continue;
                    }
                    if let (ResolveResult::Unknown(prefix), _) =
                        reader.resolver().resolve_attribute(key)
                    {
                        panic!(
                            "unbound attribute prefix {}: {xml}",
                            String::from_utf8_lossy(&prefix)
                        );
                    }
                }
            }
            Event::Eof => return found,
            _ => {}
        }
    }
}

#[test]
fn keeps_form_field_data_bound_to_a_prefix_declared_on_the_run_the_begin_or_the_root() {
    let data = r#"<w:ffData><w:name w:val="T"/><ext:foo/></w:ffData>"#;
    let ext = r#" xmlns:ext="urn:x""#;
    for (root, run, begin) in [("", ext, ""), ("", "", ext), (ext, "", "")] {
        let body = format!(
            r#"<w:p>{}<w:r{run}><w:fldChar w:fldCharType="begin"{begin}>{data}</w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
            text("Label "),
            text("Ada"),
        );
        let xml = saved_document(&fixture_in(&format!("{NS}{root}"), &body));
        assert!(binds_ext_foo(&xml), "{xml}");
        assert_eq!(xml.matches("<w:ffData").count(), 1, "{xml}");
        let reopened = package(xml.clone());
        parse_docx_s9_wire(&reopened, S9ParseOptions::default()).unwrap();
        assert!(binds_ext_foo(&save(&reopened).unwrap()), "{xml}");
    }
}
