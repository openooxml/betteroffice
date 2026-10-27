//! Table of contents page fill: a TOC entry is one line in Word when its
//! heading text, dot leader tab and page number fit the column, even when
//! the entry's right tab stop is authored with the ST_TabJc spelling
//! `right` (ECMA-376 17.18.87) rather than the internal `end`, and even
//! when that stop sits past the right edge of the column. Measured
//! against the corpus TOC pages, where BO gave single Word lines two
//! lines (heights 41.2px against Word's single 20.6px step in the EPA
//! documents, 48.8px against 24.4px in the Canada documents) and so
//! filled TOC pages with fewer entries than Word.
//!
//! The measure-level tests feed the authored spellings straight to the
//! paragraph measurer, the one path that receives tab stops without the
//! editing bridge's `right` to `end` conversion. The engine-level test
//! is the control: a whole synthetic document whose TOC already laid
//! out on one line per entry before this change, and must still.

use docx_edit::{EngineSession, seed_from_docx};
use ooxml_text::{FontStore, measure_paragraph_json};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
/// One `0` glyph at 12pt in the fixture font: 1139/128 px.
const W0: f64 = 1139.0 / 128.0;

fn store() -> FontStore {
    let mut store = FontStore::new();
    store.register(FONT.to_vec()).expect("fixture registers");
    store
}

fn measure(block: Value, max_width: f64) -> Value {
    let input = json!({
        "block": block,
        "maxWidth": max_width,
        "fontChains": { "liberation sans|0|0": [0] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" }
    });
    let out = measure_paragraph_json(&store(), &input.to_string()).expect("measure succeeds");
    serde_json::from_str(&out).expect("output is valid JSON")
}

fn approx(actual: f64, expected: f64, label: &str) {
    assert!(
        (actual - expected).abs() < 0.01,
        "{label}: expected {expected}, got {actual}"
    );
}

/// A TOC entry shaped like the corpus ones: heading text, a leader tab,
/// then the PAGEREF result as one field run. The tab stop is at 100px on
/// a 105px column, so the entry fits on one line only if the tab
/// reserves the page number's width against the stop, as a right tab
/// does in Word. Treating the authored `right` as a start tab leaves
/// the number no room: the tab and number wrap together to a second
/// line and the entry takes two lines where Word takes one.
#[test]
fn a_right_authored_tab_reserves_room_for_the_page_number() {
    for val in ["right", "end"] {
        let measured = measure(
            json!({
                "kind": "paragraph",
                "runs": [
                    { "kind": "text", "text": "0".repeat(10) },
                    { "kind": "tab" },
                    { "kind": "field", "fallback": "9" }
                ],
                "attrs": { "tabs": [{ "val": val, "pos": 1500.0, "leader": "dot" }] }
            }),
            105.0,
        );
        let lines = measured["lines"].as_array().expect("lines");
        assert_eq!(lines.len(), 1, "val {val}: entry must stay on one line");
        approx(
            lines[0]["width"].as_f64().expect("width"),
            100.0,
            "page number right edge on the stop",
        );
    }
}

/// The authored `left` spelling is a start tab: it reserves nothing for
/// what follows, so the line runs past the stop by the text width.
#[test]
fn a_left_authored_tab_reserves_nothing() {
    for val in ["left", "start"] {
        let measured = measure(
            json!({
                "kind": "paragraph",
                "runs": [{ "kind": "tab" }, { "kind": "text", "text": "00" }],
                "attrs": { "tabs": [{ "val": val, "pos": 1500.0 }] }
            }),
            300.0,
        );
        let lines = measured["lines"].as_array().expect("lines");
        assert_eq!(lines.len(), 1, "val {val}");
        approx(
            lines[0]["width"].as_f64().expect("width"),
            100.0 + 2.0 * W0,
            "start tab plus following text",
        );
    }
}

/// Control: a right tab stop past the right edge of the column (the
/// Welsh corpus TOC stops at 9629 twips on a 9638 twip column, and
/// stops past the edge occur in the wild) clamps the page number to
/// the line edge instead of wrapping the entry. This already held and
/// must keep holding.
#[test]
fn a_right_tab_past_the_column_edge_clamps_to_the_edge() {
    for val in ["right", "end"] {
        let measured = measure(
            json!({
                "kind": "paragraph",
                "runs": [
                    { "kind": "text", "text": "00" },
                    { "kind": "tab" },
                    { "kind": "field", "fallback": "9" }
                ],
                "attrs": { "tabs": [{ "val": val, "pos": 3000.0, "leader": "dot" }] }
            }),
            100.0,
        );
        let lines = measured["lines"].as_array().expect("lines");
        assert_eq!(lines.len(), 1, "val {val}: entry must stay on one line");
        approx(
            lines[0]["width"].as_f64().expect("width"),
            100.0,
            "page number clamped to the line edge",
        );
    }
}

fn document(body: &str) -> Vec<u8> {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120"/><w:ind w:left="360" w:hanging="360"/></w:pPr></w:style></w:styles>"#;
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

fn render(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76400);
    seed_from_docx(engine.doc(), bytes).unwrap();
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

/// Control at document level: a TOC whose entries carry an authored
/// right dot leader tab in the paragraph properties, the shape the
/// corpus documents use, lays every entry out on one line on the
/// first page. This already held through the editing bridge, which
/// converts the authored spelling, and must keep holding.
#[test]
fn toc_entries_with_a_right_leader_tab_stay_on_one_line() {
    let entry = |title: &str, page: &str| {
        format!(
            r#"<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">{title}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>{page}</w:t></w:r></w:p>"#
        )
    };
    let body = format!(
        "<w:p><w:r><w:t>Table of Contents</w:t></w:r></w:p>{}{}{}",
        entry("1. Introduction", "2"),
        entry("2. Overall Objective and Goals of the Monitoring Plan", "2"),
        entry("3. Air Quality Monitoring Plan Development Process", "3"),
    );
    let output = render(&document(&body));
    let pages = output["layout"]["pages"].as_array().expect("pages");
    assert_eq!(pages.len(), 1, "the whole TOC fits the first page");
    let mut entries = 0;
    for fragment in pages[0]["fragments"].as_array().expect("fragments") {
        let lines = fragment["resolvedLines"].as_array().cloned().unwrap_or_default();
        let text: String = lines
            .iter()
            .flat_map(|line| line["segments"].as_array().cloned().unwrap_or_default())
            .filter_map(|segment| segment["text"].as_str().map(str::to_owned))
            .collect();
        if text.contains("Introduction")
            || text.contains("Overall Objective")
            || text.contains("Development Process")
        {
            entries += 1;
            assert_eq!(lines.len(), 1, "entry {text:?} must stay on one line");
        }
    }
    assert_eq!(entries, 3, "all three entries are laid out");
}
