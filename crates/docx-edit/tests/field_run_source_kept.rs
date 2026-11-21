use std::sync::Arc;

use docx_edit::{EditingDoc, EngineSession, SegmentContent, seed_from_docx, seed_from_parsed_docx};
use docx_parse::{BlockContent, ComplexField, InlineNode, ParagraphContent, S9WireEnvelope};
use serde_json::{Value, json};
use yrs::{Any, Map, ReadTxn, Transact};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const ONE_FIELD_RUN: &str = concat!(
    r#"<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText>"#,
    r#"<w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
);
const THEMED_ONE_FIELD_RUN: &str = concat!(
    r#"<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr><w:fldChar w:fldCharType="begin"/>"#,
    r#"<w:instrText xml:space="preserve"> PAGE </w:instrText><w:fldChar w:fldCharType="separate"/>"#,
    r#"<w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
);
const THEME_TYPE: &str = concat!(
    r#"<Override PartName="/word/theme/theme1.xml" "#,
    r#"ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#
);
const THEME_RELS: &str = concat!(
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
    "</Relationships>"
);
const THEME: &str = concat!(
    r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>"#,
    r#"<a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>"#,
    r#"<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
);

fn field(code: &str) -> String {
    format!(
        r#"<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> {code} </w:instrText><w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/>"#
    )
}

fn text(value: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{value}</w:t></w:r>"#)
}

/// Two pages; the first holds one run of a `PAGE` and a `NUMPAGES` field with " of " between.
fn document() -> Vec<u8> {
    document_between(" of ")
}

fn document_between(middle: &str) -> Vec<u8> {
    let run = format!(
        r#"<w:r>{}<w:t xml:space="preserve">{middle}</w:t>{}</w:r>"#,
        field("PAGE"),
        field("NUMPAGES")
    );
    document_holding(&run)
}

/// Two pages; the first holds `run` between "AB" and "CD".
fn document_holding(run: &str) -> Vec<u8> {
    let body = format!(
        r#"<w:p>{}</w:p><w:p>{}{run}{}</w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr>{}</w:p>"#,
        text("Prev"),
        text("AB"),
        text("CD"),
        text("Two")
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn display_list(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(75310);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap()
}

/// Each painted field's category and text, in paint order.
fn painted_fields(value: &Value, out: &mut Vec<(String, String)>) {
    match value {
        Value::Object(map) if map.contains_key("field") && map.contains_key("width") => {
            let category = map["field"]["category"].as_str().unwrap_or_default();
            let text = map["text"].as_str().unwrap_or_default();
            out.push((category.to_owned(), text.to_owned()));
        }
        Value::Object(map) => map.values().for_each(|value| painted_fields(value, out)),
        Value::Array(items) => items.iter().for_each(|value| painted_fields(value, out)),
        _ => {}
    }
}

#[test]
fn a_run_holding_two_fields_paints_the_first_fields_label_as_main_does() {
    docx_layout::clear_measure_fonts();
    let display = display_list(&document());
    let pages = display["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 2, "{display}");
    let mut first = Vec::new();
    painted_fields(&pages[0], &mut first);
    assert_eq!(first, [("PAGE".to_owned(), "1".to_owned())], "{display}");
    let mut second = Vec::new();
    painted_fields(&pages[1], &mut second);
    assert!(second.is_empty(), "{display}");
}

#[test]
fn a_run_holding_one_field_paints_its_label_as_main_does() {
    docx_layout::clear_measure_fonts();
    let display = display_list(&document_holding(ONE_FIELD_RUN));
    let pages = display["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 2, "{display}");
    let mut first = Vec::new();
    painted_fields(&pages[0], &mut first);
    assert_eq!(first, [("PAGE".to_owned(), "1".to_owned())], "{display}");
    let mut second = Vec::new();
    painted_fields(&pages[1], &mut second);
    assert!(second.is_empty(), "{display}");
}

#[test]
fn seq_text_between_the_fields_seeds_no_sequence_metadata_as_on_main() {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &document_between("SEQ")).unwrap();
    let txn = doc.yrs_doc().transact();
    let session = txn.get_map("session").unwrap();
    assert_eq!(session.get(&txn, "opaqueSequences"), None);
}

/// `bytes` with a theme part whose minor font is Calibri.
fn themed(bytes: &[u8]) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(bytes).unwrap();
    for (path, xml) in &mut parts {
        if path == "[Content_Types].xml" {
            let types = String::from_utf8(std::mem::take(xml)).unwrap();
            *xml = types.replace("</Types>", &format!("{THEME_TYPE}</Types>")).into_bytes();
        }
    }
    parts.extend([
        (
            "word/_rels/document.xml.rels".to_owned(),
            THEME_RELS.as_bytes().to_vec(),
        ),
        (
            "word/theme/theme1.xml".to_owned(),
            THEME.as_bytes().to_vec(),
        ),
    ]);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The stored data of the first field seeded into the body.
fn seeded_field(doc: &EditingDoc) -> Value {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::OtherEmbed { kind, payload } if kind == "field" => {
                match payload.get("fieldData") {
                    Some(Any::String(data)) => Some(serde_json::from_str(data).unwrap()),
                    _ => None,
                }
            }
            _ => None,
        })
        .unwrap()
}

/// The document part of `bytes` written from its parsed model with `field` in place of the field
/// of its second paragraph.
fn written_with(bytes: &[u8], field: Value) -> String {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut document = serde_json::to_value(&package.document).unwrap();
    document.as_object_mut().unwrap().remove("sections");
    let content = document["content"][1]["content"].as_array_mut().unwrap();
    let at = content
        .iter()
        .position(|node| node["type"] == "complexField")
        .unwrap();
    content[at] = field;
    let request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": document,
        "options": {"updateModifiedDate": false}
    });
    let request: docx_parse::S13SaveRequest = serde_json::from_value(request).unwrap();
    let saved = docx_parse::write_docx_s13(request, bytes).unwrap();
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    String::from_utf8(xml).unwrap()
}

fn without_source(value: &mut Value) {
    if let Some(object) = value.as_object_mut() {
        object.remove("sourceRunXml");
        object.remove("sourceRunDigest");
    }
}

fn field_in(envelope: &mut S9WireEnvelope) -> &mut ComplexField {
    let BlockContent::Paragraph(paragraph) = &mut envelope.document.package.document.content[1]
    else {
        panic!("the second block is a paragraph");
    };
    Arc::make_mut(paragraph)
        .content
        .iter_mut()
        .find_map(|content| match content {
            ParagraphContent::Inline(InlineNode::ComplexField(field)) => Some(&mut **field),
            _ => None,
        })
        .unwrap()
}

#[test]
fn a_seeded_field_run_with_a_theme_font_is_written_as_read() {
    let two_fields = format!(
        r#"<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr>{}<w:t xml:space="preserve"> of </w:t>{}</w:r>"#,
        field("PAGE"),
        field("NUMPAGES")
    );
    for run in [THEMED_ONE_FIELD_RUN, two_fields.as_str()] {
        let bytes = themed(&document_holding(run));
        let doc = EditingDoc::new(1);
        seed_from_docx(&doc, &bytes).unwrap();
        let field = seeded_field(&doc);
        assert_eq!(field["sourceRunXml"], run, "{field}");
        let xml = written_with(&bytes, field);
        assert!(xml.contains(run), "{xml}");
    }
}

#[test]
fn a_field_edited_before_seeding_is_written_from_the_model() {
    let bytes = themed(&document_holding(THEMED_ONE_FIELD_RUN));
    let (mut envelope, parts) = docx_parse::parse_docx_s9_wire_parts_with_limits(
        &bytes,
        Default::default(),
        &Default::default(),
    )
    .unwrap();
    field_in(&mut envelope).instruction = "DATE".to_owned();
    let doc = EditingDoc::new(1);
    seed_from_parsed_docx(&doc, envelope, parts, &bytes).unwrap();
    let field = seeded_field(&doc);
    let mut main = field.clone();
    without_source(&mut main);
    let saved = written_with(&bytes, field);
    assert!(!saved.contains(THEMED_ONE_FIELD_RUN), "{saved}");
    assert!(saved.contains("DATE"), "{saved}");
    assert_eq!(saved, written_with(&bytes, main));
}
