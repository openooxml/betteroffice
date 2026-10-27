//! Space before at the top of a page, measured against Word: at
//! compatibilityMode 15 Word spends the space before of a paragraph that
//! breaks its own page (pageBreakBefore or a leading page break run);
//! below 15 it keeps it unless suppressSpBfAfterPgBrk spends the run's.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const BREAK: &str = r#"<w:r><w:br w:type="page"/></w:r>"#;

fn document(body: &str, section: &str, settings: Option<&str>) -> Vec<u8> {
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#;
    let mut types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#.to_owned();
    let mut rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#.to_owned();
    let mut parts = vec![
        (
            "_rels/.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}{section}</w:body></w:document>"#
            ),
        ),
        ("word/styles.xml".to_owned(), styles.to_owned()),
    ];
    if let Some(settings) = settings {
        types.push_str(
            r#"<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>"#,
        );
        rels.push_str(
            r#"<Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>"#,
        );
        parts.push((
            "word/settings.xml".to_owned(),
            format!(
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{settings}</w:settings>"#
            ),
        ));
    }
    types.push_str("</Types>");
    rels.push_str("</Relationships>");
    parts.insert(0, ("[Content_Types].xml".to_owned(), types));
    parts.insert(2, ("word/_rels/document.xml.rels".to_owned(), rels));
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn compat(mode: u8, extra: &str) -> String {
    format!(
        r#"<w:compat>{extra}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat>"#
    )
}

const LETTER: &str = r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#;
/// A page with 100 px of body height: six 16 px lines fill it.
fn render(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76401);
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

const TOP: f64 = 96.0;
const BEFORE: f64 = 32.0;

#[test]
fn compat15_spends_space_before_after_page_break_before() {
    let body = "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before=\"480\"/></w:pPr><w:r><w:t>TARGET</w:t></w:r></w:p>";
    let output = render(&document(body, LETTER, Some(&compat(15, ""))));
    let (page, y) = target(&output);
    assert_eq!(page, 2);
    assert!((y - TOP).abs() < 0.01, "{y}");
}

#[test]
fn compat15_spends_space_before_after_a_leading_break_run() {
    let body = format!(
        "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr>{BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>"
    );
    let output = render(&document(&body, LETTER, Some(&compat(15, ""))));
    let (page, y) = target(&output);
    assert_eq!(page, 2);
    assert!((y - TOP).abs() < 0.01, "{y}");
}

#[test]
fn compat14_keeps_space_before_after_an_authored_break() {
    for target_para in [
        "<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before=\"480\"/></w:pPr><w:r><w:t>TARGET</w:t></w:r></w:p>".to_owned(),
        format!("<w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr>{BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>"),
    ] {
        let body = format!(
            "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p>{target_para}"
        );
        let output = render(&document(&body, LETTER, Some(&compat(14, ""))));
        let (page, y) = target(&output);
        assert_eq!(page, 2);
        assert!((y - TOP - BEFORE).abs() < 0.01, "{y}");
    }
}

#[test]
fn suppress_flag_at_compat14_spends_only_the_break_run() {
    let settings = compat(14, "<w:suppressSpBfAfterPgBrk/>");
    let hard = format!(
        "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr>{BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>"
    );
    let output = render(&document(&hard, LETTER, Some(&settings)));
    assert!((target(&output).1 - TOP).abs() < 0.01);
    let property = "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before=\"480\"/></w:pPr><w:r><w:t>TARGET</w:t></w:r></w:p>";
    let output = render(&document(property, LETTER, Some(&settings)));
    assert!((target(&output).1 - TOP - BEFORE).abs() < 0.01);
}

#[test]
fn compat15_keeps_space_before_after_a_section_break() {
    let body = concat!(
        "<w:p><w:r><w:t>FILLER</w:t></w:r></w:p>",
        "<w:p><w:pPr><w:sectPr><w:type w:val=\"nextPage\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:pPr><w:r><w:t>ENDSEC</w:t></w:r></w:p>",
        "<w:p><w:pPr><w:spacing w:before=\"480\"/></w:pPr><w:r><w:t>TARGET</w:t></w:r></w:p>"
    );
    let output = render(&document(body, LETTER, Some(&compat(15, ""))));
    let (page, y) = target(&output);
    assert_eq!(page, 2);
    assert!((y - TOP - BEFORE).abs() < 0.01, "{y}");
}
