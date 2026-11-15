use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NUMBERING: &str = r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;

fn revision(kind: &str, id: u32, author: &str, content: &str) -> String {
    format!(
        r#"<w:{kind} w:id="{id}" w:author="{author}" w:date="2026-01-01T00:00:00Z">{content}</w:{kind}>"#
    )
}

fn text(value: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{value}</w:t></w:r>"#)
}

fn deleted(value: &str) -> String {
    format!(r#"<w:r><w:delText xml:space="preserve">{value}</w:delText></w:r>"#)
}

fn list_item(content: &str, mark: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:rPr>{mark}</w:rPr></w:pPr>{content}</w:p>"#
    )
}

fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="num" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_owned()),
        ("word/numbering.xml", format!(r#"<w:numbering xmlns:w="{W}">{NUMBERING}</w:numbering>"#)),
        ("word/document.xml", format!(r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn primitives(bytes: &[u8]) -> Vec<Value> {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(75302);
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
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .clone()
}

fn in_paragraph(primitives: &[Value], index: usize) -> Vec<&Value> {
    primitives
        .iter()
        .filter(|primitive| primitive["blockKey"] == format!("body:p{index}"))
        .collect()
}

/// Paragraph `index`'s painted text, merged into runs of one color.
fn painted(primitives: &[Value], index: usize) -> Vec<(String, String)> {
    let mut runs: Vec<(String, String)> = Vec::new();
    for primitive in in_paragraph(primitives, index) {
        if primitive["kind"] != "text" || primitive["listMarker"] == true {
            continue;
        }
        let (text, color) = (
            primitive["text"].as_str().unwrap(),
            primitive["color"].as_str().unwrap(),
        );
        match runs.last_mut() {
            Some((run, last)) if last == color => run.push_str(text),
            _ => runs.push((text.to_owned(), color.to_owned())),
        }
    }
    runs
}

fn text_color(primitives: &[Value], index: usize, value: &str) -> String {
    painted(primitives, index)
        .into_iter()
        .find(|(text, _)| text == value)
        .unwrap_or_else(|| panic!("no {value:?} in paragraph {index}"))
        .1
}

/// Paragraph `index`'s distinct revision rules, in paint order.
fn rules(primitives: &[Value], index: usize) -> Vec<(String, String, Option<String>)> {
    let mut rules = Vec::new();
    for primitive in in_paragraph(primitives, index) {
        if primitive["kind"] != "decoration" || primitive["listMarker"] == true {
            continue;
        }
        let rule = (
            primitive["deco"].as_str().unwrap().to_owned(),
            primitive["color"].as_str().unwrap().to_owned(),
            primitive["style"].as_str().map(str::to_owned),
        );
        if !rules.contains(&rule) {
            rules.push(rule);
        }
    }
    rules
}

fn rule(deco: &str, color: &str, style: Option<&str>) -> (String, String, Option<String>) {
    (deco.to_owned(), color.to_owned(), style.map(str::to_owned))
}

#[test]
fn authors_take_word_colors_in_document_order() {
    docx_layout::clear_measure_fonts();
    let body = [
        format!(
            r#"<w:p><w:pPr><w:rPr>{}</w:rPr></w:pPr>{}</w:p>"#,
            revision("ins", 1, "Mark", ""),
            text("plain")
        ),
        format!("<w:p>{}</w:p>", revision("ins", 2, "Ann", &text("ann"))),
        format!("<w:p>{}</w:p>", revision("del", 3, "Bob", &deleted("bob"))),
        format!(
            "<w:p>{}{}</w:p>",
            revision(
                "ins",
                4,
                "Ann",
                &revision("del", 5, "Bob", &deleted("both"))
            ),
            revision("moveTo", 6, "Ann", &text("moved"))
        ),
    ]
    .concat();
    let primitives = primitives(&document(&body));
    assert_eq!(text_color(&primitives, 1, "ann"), "#0078D4");
    assert_eq!(rules(&primitives, 1), [rule("underline", "#0078D4", None)]);
    assert_eq!(text_color(&primitives, 2, "bob"), "#5C2E91");
    assert_eq!(rules(&primitives, 2), [rule("strike", "#5C2E91", None)]);
    assert_eq!(text_color(&primitives, 3, "both"), "#5C2E91");
    assert_eq!(text_color(&primitives, 3, "moved"), "#2C6234");
    assert_eq!(
        rules(&primitives, 3),
        [
            rule("underline", "#0078D4", None),
            rule("strike", "#5C2E91", None),
            rule("underline", "#2C6234", Some("double")),
        ]
    );
    let bars: Vec<&Value> = primitives
        .iter()
        .filter(|primitive| primitive["kind"] == "rect" && primitive["fill"] == "#000000")
        .collect();
    assert!(!bars.is_empty());
    assert!(bars.iter().all(|bar| bar["x"].as_f64() == Some(48.0)));
}

#[test]
fn renumbered_list_items_show_original_and_final_numbers() {
    docx_layout::clear_measure_fonts();
    let deleted_mark = revision("del", 9, "Ann", "");
    let body = [
        list_item(&text("one"), ""),
        list_item(&revision("del", 10, "Ann", &deleted("two")), &deleted_mark),
        list_item(&text("three"), ""),
    ]
    .concat();
    let primitives = primitives(&document(&body));
    let marker = |index: usize| {
        in_paragraph(&primitives, index)
            .into_iter()
            .filter(|primitive| primitive["listMarker"] == true && primitive["kind"] == "text")
            .map(|primitive| primitive["text"].as_str().unwrap().to_owned())
            .collect::<String>()
    };
    assert_eq!(marker(0), "1.");
    assert_eq!(marker(1), "2.");
    assert_eq!(marker(2), "3.2.");
    let renumbering: Vec<(String, String)> = in_paragraph(&primitives, 2)
        .into_iter()
        .filter(|primitive| primitive["listMarker"] == true && primitive["kind"] == "decoration")
        .map(|primitive| {
            (
                primitive["deco"].as_str().unwrap().to_owned(),
                primitive["color"].as_str().unwrap().to_owned(),
            )
        })
        .collect();
    assert_eq!(
        renumbering,
        [
            ("strike".to_owned(), "#0078D4".to_owned()),
            ("underline".to_owned(), "#0078D4".to_owned()),
        ]
    );
}
