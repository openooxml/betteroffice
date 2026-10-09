//! Word keeps a paragraph's space-before when the paragraph breaks the page
//! itself — by a hard `w:br w:type="page"` run or by `w:pageBreakBefore` —
//! and drops it when the page turned by itself. What it keeps is the
//! collapsed gap less the previous paragraph's space-after. Measured against
//! Word 16.113 (oxi-ja-policies-01 page 45: +12.00pt for `w:before="240"`,
//! page 46: +6.00pt for `w:before="120"`, and hand-authored probes).

use docx_edit::{EngineSession, bridge::RenderEnv, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const BREAK: &str = r#"<w:r><w:br w:type="page"/></w:r>"#;

fn document(body: &str) -> Vec<u8> {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#;
    let section = r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#;
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}{section}</w:body></w:document>"#)),
        ("word/styles.xml", styles.to_owned()),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn session(bytes: &[u8]) -> EngineSession {
    let engine = EngineSession::new(76400);
    seed_from_docx(engine.doc(), bytes).unwrap();
    engine
}

fn render(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = session(bytes);
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                    "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

/// `y` of the fragment whose text is `TARGET`, and its page number.
fn target(output: &Value) -> (u64, f64) {
    for page in output["layout"]["pages"].as_array().unwrap() {
        for fragment in page["fragments"].as_array().unwrap() {
            let text: String = fragment["resolvedLines"]
                .as_array()
                .map(|lines| {
                    lines
                        .iter()
                        .flat_map(|line| line["segments"].as_array().cloned().unwrap_or_default())
                        .filter_map(|segment| segment["text"].as_str().map(str::to_owned))
                        .collect()
                })
                .unwrap_or_default();
            if text == "TARGET" {
                return (
                    page["number"].as_u64().unwrap(),
                    fragment["y"].as_f64().unwrap(),
                );
            }
        }
    }
    panic!("no TARGET fragment");
}

const CONTENT_TOP: f64 = 96.0;
const SPACE_BEFORE: f64 = 32.0;

#[test]
fn a_leading_hard_page_break_run_keeps_the_paragraph_space_before() {
    for fillers in [1, 10, 40] {
        let body = format!(
            "{}<w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr>{BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p>".repeat(fillers)
        );
        let (page, y) = target(&render(&document(&body)));
        assert_eq!(page, 2, "{fillers} fillers");
        assert!(
            (y - CONTENT_TOP - SPACE_BEFORE).abs() < 0.01,
            "{fillers} fillers: {y}"
        );
    }
}

#[test]
fn a_standalone_break_paragraph_still_drops_the_next_space_before() {
    let body = format!(
        "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p>{BREAK}</w:p><w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr><w:r><w:t>TARGET</w:t></w:r></w:p>"
    );
    let (page, y) = target(&render(&document(&body)));
    assert_eq!(page, 2);
    assert!((y - CONTENT_TOP).abs() < 0.01, "{y}");
}

#[test]
fn a_column_break_run_keeps_the_paragraph_space_before() {
    let body = r#"<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="480"/></w:pPr><w:r><w:br w:type="column"/></w:r><w:r><w:t>TARGET</w:t></w:r></w:p>"#;
    let (page, y) = target(&render(&document(body)));
    assert_eq!(page, 2);
    assert!((y - CONTENT_TOP - SPACE_BEFORE).abs() < 0.01, "{y}");
}

#[test]
fn a_page_break_before_property_keeps_the_paragraph_space_before() {
    for fillers in [1, 10, 40] {
        let body = format!(
            "{}<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before=\"480\"/></w:pPr><w:r><w:t>TARGET</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p>".repeat(fillers)
        );
        let (page, y) = target(&render(&document(&body)));
        assert_eq!(page, 2, "{fillers} fillers");
        assert!(
            (y - CONTENT_TOP - SPACE_BEFORE).abs() < 0.01,
            "{fillers} fillers: {y}"
        );
    }
}

/// Word lays the collapsed gap out from the bottom up, so the break carries
/// only `max(0, before - after)`. Measured with 0/12/24/36pt space-after
/// against 24pt space-before.
#[test]
fn an_authored_break_spends_the_previous_space_after() {
    for (after, kept) in [(0, 32.0), (240, 16.0), (480, 0.0), (720, 0.0)] {
        for break_markup in [("<w:pageBreakBefore/>", ""), ("", BREAK)] {
            let (property, run) = break_markup;
            let body = format!(
                "<w:p><w:pPr><w:spacing w:after=\"{after}\"/></w:pPr><w:r><w:t>FILLER</w:t></w:r></w:p>\
                 <w:p><w:pPr>{property}<w:spacing w:before=\"480\"/></w:pPr>{run}<w:r><w:t>TARGET</w:t></w:r></w:p>"
            );
            let (page, y) = target(&render(&document(&body)));
            assert_eq!(page, 2, "after {after}");
            assert!((y - CONTENT_TOP - kept).abs() < 0.01, "after {after}: {y}");
        }
    }
}

/// The break is a run, not `w:pPr`, so the save projection must not learn a
/// `w:pageBreakBefore` the file never had.
#[test]
fn a_leading_hard_break_run_does_not_claim_page_break_before() {
    let body = format!(
        "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr>{BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>"
    );
    let engine = session(&document(&body));
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &RenderEnv::default())
            .unwrap(),
    )
    .unwrap();
    let attrs = &blocks.as_array().unwrap()[1]["attrs"];
    assert_eq!(attrs["pageBreakBeforeRun"], json!(true));
    assert_eq!(attrs["pageBreakBefore"], Value::Null);
}
