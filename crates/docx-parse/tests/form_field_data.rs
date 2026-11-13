use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

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
    save_edited(source, selective, |_| {})
}

fn save_edited(
    source: &[u8],
    selective: Value,
    edit: impl FnOnce(&mut Value),
) -> Result<String, String> {
    save_request(source, |request| {
        if !selective.is_null() {
            request["selective"] = selective;
        }
        edit(&mut request["document"]);
    })
}

fn save_request(source: &[u8], edit: impl FnOnce(&mut Value)) -> Result<String, String> {
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
    edit(&mut request);
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
        xml.matches(r#"<w:fldChar w:fldCharType="begin" w:dirty="true"/>"#)
            .count(),
        1,
        "{xml}"
    );
    assert_eq!(xml.matches("w:dirty=").count(), 1, "{xml}");
}

const ALIASED: &str = r#"xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const ALIASED_DATA: &str =
    r#"<ns0:ffData><ns0:name ns0:val="Ext1"/><ns0:enabled/><ns0:textInput/></ns0:ffData>"#;

const NORMALIZED_DATA: &str =
    r#"<w:ffData><w:name w:val="Ext1"/><w:enabled/><w:textInput/></w:ffData>"#;
const FOREIGN: &str =
    r#"xmlns:x="urn:x" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const FOREIGN_DATA: &str = r#"<x:ffData><x:name x:val="Ext1"/><x:enabled/></x:ffData>"#;
const W_URI: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn two_paragraphs(namespaces: &str, data: &str) -> (Vec<u8>, String, String) {
    let field = format!(
        r#"<w:p w14:paraId="73000003">{}{}</w:p>"#,
        text("Label "),
        form_field(data, "FORMTEXT", &text("Ada")),
    );
    let body = format!(
        r#"{field}<w:p w14:paraId="73000004">{}</w:p>"#,
        text("Other")
    );
    let document =
        format!("<w:document {NS} {namespaces}><w:body>{body}<w:sectPr/></w:body></w:document>");
    (package(document.clone()), document, field)
}

fn aliased_document() -> (Vec<u8>, String, String) {
    two_paragraphs(ALIASED, ALIASED_DATA)
}

fn lossy(bytes: &[u8]) -> String {
    String::from_utf8_lossy(bytes).into_owned()
}

/// Each element and attribute inside `w:ffData`, by namespace URI and local name, in order;
/// panics on any prefix of the part that no declaration binds.
fn form_data_events(xml: &str) -> Vec<String> {
    let mut reader = NsReader::from_str(xml);
    let mut events = Vec::new();
    let mut depth = 0usize;
    loop {
        let (resolved, event) = reader.read_resolved_event().unwrap();
        let namespace = match resolved {
            ResolveResult::Unknown(prefix) => panic!("unbound prefix {}: {xml}", lossy(&prefix)),
            ResolveResult::Bound(Namespace(uri)) => uri.to_vec(),
            ResolveResult::Unbound => Vec::new(),
        };
        let (element, empty) = match &event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) if depth > 0 => {
                events.push("/".to_owned());
                depth -= 1;
                continue;
            }
            Event::Eof => return events,
            _ => continue,
        };
        let local = lossy(element.local_name().as_ref());
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.unwrap();
            if attribute.key.as_ref().starts_with(b"xmlns")
                || attribute.key.as_ref().starts_with(b"xml:")
            {
                continue;
            }
            let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
            let uri = match resolved {
                ResolveResult::Unknown(prefix) => {
                    panic!("unbound attribute prefix {}: {xml}", lossy(&prefix))
                }
                ResolveResult::Bound(Namespace(uri)) => lossy(uri),
                ResolveResult::Unbound => String::new(),
            };
            attributes.push(format!(
                "{uri} {}={}",
                lossy(name.as_ref()),
                lossy(&attribute.value)
            ));
        }
        if depth == 0 && !(local == "ffData" && namespace == W_URI) {
            continue;
        }
        events.push(format!("{} {local} {attributes:?}", lossy(&namespace)));
        if !empty {
            depth += 1;
        }
    }
}

fn assert_same_form_data(source: &str, xml: &str) {
    let events = form_data_events(source);
    assert!(events.len() > 1, "{source}");
    assert_eq!(form_data_events(xml), events, "{xml}");
}

#[test]
fn keeps_form_field_data_bound_through_a_root_alias_on_a_full_write() {
    let (source, document, _) = aliased_document();
    let xml = save(&source).unwrap();
    assert_same_form_data(&document, &xml);
    assert_eq!(xml.matches(NORMALIZED_DATA).count(), 1, "{xml}");
    let reopened = package(xml.clone());
    parse_docx_s9_wire(&reopened, S9ParseOptions::default()).unwrap();
    assert_same_form_data(&document, &save(&reopened).unwrap());
}

#[test]
fn keeps_form_field_data_bound_through_a_root_alias_when_its_own_paragraph_is_written() {
    let (source, document, _) = aliased_document();
    let xml = save_with(&source, json!({"changedParaIds": ["73000003"]})).unwrap();
    assert_same_form_data(&document, &xml);
    assert_eq!(xml.matches(NORMALIZED_DATA).count(), 1, "{xml}");
}

#[test]
fn keeps_the_source_bytes_of_a_paragraph_with_aliased_form_field_data_when_another_is_written() {
    let (source, document, field) = aliased_document();
    let xml = save_with(&source, json!({"changedParaIds": ["73000004"]})).unwrap();
    assert_eq!(xml.matches(field.as_str()).count(), 1, "{xml}");
    assert_same_form_data(&document, &xml);
}

#[test]
fn saves_an_ff_data_element_of_another_namespace_as_no_form_field_data() {
    let (source, _, _) = two_paragraphs(FOREIGN, FOREIGN_DATA);
    for selective in [Value::Null, json!({"changedParaIds": ["73000003"]})] {
        let xml = save_with(&source, selective).unwrap();
        assert_eq!(xml.matches("ffData").count(), 0, "{xml}");
        let begin = r#"<w:fldChar w:fldCharType="begin"/>"#;
        assert_eq!(xml.matches(begin).count(), 1, "{xml}");
        parse_docx_s9_wire(&package(xml), S9ParseOptions::default()).unwrap();
    }
}

fn replace_raw_form_data(value: &mut Value, raw: &str) -> usize {
    match value {
        Value::Object(object) => {
            let mut own = 0;
            if let Some(kept) = object.get_mut("rawFormDataXml") {
                *kept = Value::String(raw.to_owned());
                own = 1;
            }
            own + object
                .values_mut()
                .map(|child| replace_raw_form_data(child, raw))
                .sum::<usize>()
        }
        Value::Array(items) => items
            .iter_mut()
            .map(|child| replace_raw_form_data(child, raw))
            .sum(),
        _ => 0,
    }
}

#[test]
fn refuses_to_write_kept_form_field_data_that_is_not_w_ff_data_naming_the_field_and_the_reason() {
    let (source, _, field) = aliased_document();
    let unwritable = |document: &mut Value| {
        assert!(replace_raw_form_data(document, FOREIGN_DATA) > 0);
    };
    for selective in [Value::Null, json!({"changedParaIds": ["73000003"]})] {
        let error = save_edited(&source, selective, unwritable).unwrap_err();
        assert!(error.contains("\"Ext1\""), "{error}");
        assert!(error.contains("found x:ffData"), "{error}");
    }
    let xml = save_edited(&source, json!({"changedParaIds": ["73000004"]}), unwritable).unwrap();
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

#[test]
fn refuses_model_form_field_data_with_an_unbound_prefix_naming_the_field_and_the_prefix() {
    let source = fixture(&form_paragraph());
    let unbound = r#"<w:ffData><w:name w:val="T"/><ext:foo/></w:ffData>"#;
    let error = save_edited(&source, Value::Null, |document| {
        assert!(replace_raw_form_data(document, unbound) > 0);
    })
    .unwrap_err();
    assert!(error.contains("\"T\""), "{error}");
    assert!(error.contains("unbound prefix ext"), "{error}");
    let declared = r#"<w:ffData xmlns:ext="urn:x"><w:name w:val="T"/><ext:foo/></w:ffData>"#;
    let xml = save_edited(&source, Value::Null, |document| {
        assert_eq!(replace_raw_form_data(document, declared), 3);
    })
    .unwrap();
    assert_eq!(xml.matches(declared).count(), 3, "{xml}");
}

const XSI: &str = "http://www.w3.org/2001/XMLSchema-instance";

#[test]
fn keeps_the_declaration_of_a_w_alias_an_attribute_value_of_the_form_field_data_names() {
    let w = lossy(W_URI);
    let own = format!(
        r#"<w:ffData xmlns:q="{w}" xmlns:xsi="{XSI}" xsi:type="q:CT_FFData"><q:name q:val="T"/></w:ffData>"#
    );
    let carried = format!(
        r#"<w:ffData xmlns:xsi="{XSI}" xsi:type="q:CT_FFData"><q:name q:val="T"/></w:ffData>"#
    );
    let on_run = format!(r#" xmlns:q="{w}""#);
    for (run, data) in [("", own.as_str()), (on_run.as_str(), carried.as_str())] {
        let body = format!(
            r#"<w:p>{}<w:r{run}><w:fldChar w:fldCharType="begin">{data}</w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
            text("Label "),
            text("Ada"),
        );
        let xml = saved_document(&fixture(&body));
        let data = &xml[xml.find("<w:ffData").expect(&xml)..];
        let tag = &data[..data.find('>').unwrap()];
        assert!(tag.contains(&format!(r#"xmlns:q="{w}""#)), "{xml}");
        assert!(tag.contains(r#"xsi:type="q:CT_FFData""#), "{xml}");
        assert!(
            data.starts_with(&format!(r#"{tag}><w:name w:val="T"/></w:ffData>"#)),
            "{xml}"
        );
    }
}

const COMMENTED_DATA: &str = r#"<w:ffData><w:name w:val="Check1"/><!--bo-ffdata 0--></w:ffData>"#;

fn text_and_checkbox_fields() -> (String, String) {
    let first = format!(
        r#"<w:p w14:paraId="73000005">{}{}</w:p>"#,
        text("Label "),
        form_field(TEXT_DATA, "FORMTEXT", &text("Ada")),
    );
    let second = format!(
        r#"<w:p w14:paraId="73000006">{}{}</w:p>"#,
        text("Agree "),
        form_field(CHECK_DATA, "FORMCHECKBOX", ""),
    );
    let w14 = r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
    let document =
        format!("<w:document {NS} {w14}><w:body>{first}{second}<w:sectPr/></w:body></w:document>");
    (document, first)
}

fn foreign_then_commented(request: &mut Value) {
    let paragraphs = request["document"]["content"].as_array_mut().unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(replace_raw_form_data(&mut paragraphs[0], FOREIGN_DATA), 1);
    assert_eq!(replace_raw_form_data(&mut paragraphs[1], COMMENTED_DATA), 1);
}

#[test]
fn saves_model_form_field_data_holding_the_placeholder_text_of_a_field_kept_as_source() {
    let (document, first) = text_and_checkbox_fields();
    let xml = save_request(&package(document), |request| {
        foreign_then_commented(request);
        request["selective"] = json!({"changedParaIds": ["73000006"]});
    })
    .unwrap();
    assert_eq!(xml.matches(first.as_str()).count(), 1, "{xml}");
    let begin = format!(r#"<w:fldChar w:fldCharType="begin">{COMMENTED_DATA}</w:fldChar>"#);
    assert_eq!(xml.matches(begin.as_str()).count(), 1, "{xml}");
}

#[test]
fn a_changed_field_paragraph_sends_splices_to_the_whole_part_write_validating_every_form_field() {
    let (document, _) = text_and_checkbox_fields();
    let sha256 = format!("{:x}", Sha256::digest(document.as_bytes()));
    let error = save_request(&package(document), |request| {
        foreign_then_commented(request);
        let paragraphs = request["document"]["content"].as_array_mut().unwrap();
        for (ordinal, paragraph) in paragraphs.iter_mut().enumerate() {
            paragraph["sourceOrdinal"] = json!(ordinal);
        }
        request["paragraphIds"] = json!({"splicedParts": [{
            "part": "word/document.xml",
            "sha256": sha256,
            "paragraphs": [0, 1],
            "changed": [1],
        }]});
    })
    .unwrap_err();
    assert!(error.contains("\"Ext1\""), "{error}");
}

fn replace_raw_xml(value: &mut Value, xml: &str) -> usize {
    match value {
        Value::Object(object) if object.get("type").and_then(Value::as_str) == Some("rawXml") => {
            object.insert("xml".to_owned(), json!(xml));
            1
        }
        Value::Object(object) => object
            .values_mut()
            .map(|child| replace_raw_xml(child, xml))
            .sum(),
        Value::Array(items) => items
            .iter_mut()
            .map(|child| replace_raw_xml(child, xml))
            .sum(),
        _ => 0,
    }
}

#[test]
fn writes_model_text_shaped_like_a_form_field_data_token_beside_a_field_kept_as_source() {
    let note = r#"<x:note xmlns:x="urn:x"><![CDATA[<!--bo-ffdata 0-->]]></x:note>"#;
    let kept = r#"<w:ffData><w:name w:val="Check1"/><!--bo-ffdata kept--></w:ffData>"#;
    let first = format!(
        "<w:p>{}{}</w:p>",
        text("Label "),
        form_field(TEXT_DATA, "FORMTEXT", &text("Ada")),
    );
    let second = format!(
        "<w:p>{}{}</w:p>",
        text("Agree "),
        form_field(CHECK_DATA, "FORMCHECKBOX", ""),
    );
    let third = format!(r#"<w:p>{}<x:note xmlns:x="urn:x">n</x:note></w:p>"#, text("Note "));
    let part = format!("<w:document {NS}><w:body>{first}{second}{third}</w:body></w:document>");
    let sha256 = format!("{:x}", <sha2::Sha256 as sha2::Digest>::digest(part.as_bytes()));
    let source = package(part);
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut content = serde_json::to_value(&parsed.document.package.document.content).unwrap();
    let paragraphs = content.as_array_mut().unwrap();
    assert_eq!(paragraphs.len(), 3);
    for (ordinal, paragraph) in paragraphs.iter_mut().enumerate() {
        paragraph["sourceOrdinal"] = json!(ordinal);
    }
    assert_eq!(replace_raw_form_data(&mut paragraphs[0], FOREIGN_DATA), 1);
    assert_eq!(replace_raw_form_data(&mut paragraphs[1], kept), 1);
    assert_eq!(replace_raw_xml(&mut paragraphs[2], note), 1);
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": content},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/document.xml",
            "sha256": sha256,
            "paragraphs": [0, 1, 2],
            "changed": [2],
        }]},
    }))
    .unwrap();
    let saved = write_docx_s13(request, &source).map_err(|error| error.to_string());
    assert!(saved.is_ok(), "{saved:?}");
    let (_, xml) = ooxml_opc::unzip_parts(&saved.unwrap())
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let xml = String::from_utf8(xml).unwrap();
    assert_eq!(xml.matches(first.as_str()).count(), 1, "{xml}");
    assert_eq!(xml.matches(second.as_str()).count(), 1, "{xml}");
    assert_eq!(xml.matches(note).count(), 1, "{xml}");
    assert_eq!(xml.matches("bo-ffdata").count(), 1, "{xml}");
}
