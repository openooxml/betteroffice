use docx_edit::{EditingDoc, EngineSession, seed_from_docx};
use serde_json::{Value, json};
use yrs::{Map, ReadTxn, Transact};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const ONE_FIELD_RUN: &str = concat!(
    r#"<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText>"#,
    r#"<w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
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
