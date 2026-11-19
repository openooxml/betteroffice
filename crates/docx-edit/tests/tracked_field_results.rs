use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

fn char_run(kind: &str) -> String {
    format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#)
}

/// A HYPERLINK field's begin, code and separate runs.
fn opening(deleted: bool) -> String {
    let code = if deleted { "delInstrText" } else { "instrText" };
    format!(
        r#"{}<w:r><w:{code} xml:space="preserve"> HYPERLINK "https://example.test" </w:{code}></w:r>{}"#,
        char_run("begin"),
        char_run("separate")
    )
}

fn text(value: &str, deleted: bool) -> String {
    let element = if deleted { "delText" } else { "t" };
    format!("<w:r><w:{element}>{value}</w:{element}></w:r>")
}

fn tracked(kind: &str, id: u32, content: &str) -> String {
    format!(
        r#"<w:{kind} w:id="{id}" w:author="A" w:date="2026-01-01T00:00:00Z">{content}</w:{kind}>"#
    )
}

fn document(paragraphs: &[String]) -> Vec<u8> {
    let body = paragraphs
        .iter()
        .map(|content| format!("<w:p>{content}</w:p>"))
        .collect::<String>();
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
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
    let engine = EngineSession::new(75301);
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

/// The painted field results of paragraph `index`, with their revision kinds.
fn field_results(display: &Value, index: usize) -> Vec<(String, Option<String>)> {
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| {
            primitive["blockKey"] == format!("body:p{index}")
                && primitive["kind"] == "text"
                && primitive["field"]["type"] == "HYPERLINK"
        })
        .map(|primitive| {
            (
                primitive["text"].as_str().unwrap().to_owned(),
                primitive["revision"]["kind"].as_str().map(str::to_owned),
            )
        })
        .collect()
}

fn decorations(display: &Value, index: usize) -> Vec<String> {
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| {
            primitive["blockKey"] == format!("body:p{index}") && primitive["kind"] == "decoration"
        })
        .map(|primitive| primitive["deco"].as_str().unwrap().to_owned())
        .collect()
}

#[test]
fn field_results_inside_revisions_paint_with_their_revision() {
    docx_layout::clear_measure_fonts();
    let whole = |kind: &str, id: u32, deleted: bool| {
        tracked(
            kind,
            id,
            &format!(
                "{}{}{}",
                opening(deleted),
                text("Result", deleted),
                char_run("end")
            ),
        )
    };
    let paragraphs = [
        format!(
            "{}{}{}",
            text("Before ", false),
            whole("ins", 1, false),
            text(" after", false)
        ),
        whole("del", 2, true),
        whole("moveTo", 3, false),
        whole("moveFrom", 4, true),
        format!(
            "{}{}",
            tracked(
                "ins",
                5,
                &format!("{}{}", opening(false), text("Result", false))
            ),
            tracked("ins", 6, &char_run("end"))
        ),
        format!(
            "{}{}",
            tracked(
                "ins",
                7,
                &format!("{}{}", opening(false), text("Result", false))
            ),
            char_run("end")
        ),
    ];
    let display = display_list(&document(&paragraphs));
    let result = |kind: &str| vec![("Result".to_owned(), Some(kind.to_owned()))];
    for (index, kind) in [
        (0, "ins"),
        (1, "del"),
        (2, "ins"),
        (3, "del"),
        (4, "ins"),
        (5, "ins"),
    ] {
        assert_eq!(
            field_results(&display, index),
            result(kind),
            "paragraph {index}"
        );
    }
    assert!(decorations(&display, 0).contains(&"underline".to_owned()));
    assert!(decorations(&display, 1).contains(&"strike".to_owned()));
}
