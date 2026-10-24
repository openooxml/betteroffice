//! Word lays a paragraph holding only a VML horizontal-rule picture
//! (`v:shape o:hr="t"` with `v:imagedata`) at the rule's own height: the
//! paragraph mark's font adds no line height, so a footer that starts with
//! such a rule stays inside the authored bottom margin.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const RULE: &str = r#"o:hrpct="0" o:hralign="center" o:hr="t""#;

fn rule_paragraph(rule_attrs: &str, spacing: &str, rules: usize) -> String {
    let rule = format!(
        r#"<w:r><w:rPr><w:sz w:val="14"/></w:rPr><w:pict><v:shape id="rule" style="width:487.05pt;height:2.85pt" {rule_attrs}><v:imagedata r:id="rIdRule" o:title="rule"/></v:shape></w:pict></w:r>"#
    );
    format!(
        r#"<w:p><w:pPr>{spacing}<w:jc w:val="right"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>{}</w:p>"#,
        rule.repeat(rules)
    )
}

fn document(rule_attrs: &str) -> Vec<u8> {
    document_with(&rule_paragraph(rule_attrs, "", 1))
}

fn document_with(paragraph: &str) -> Vec<u8> {
    let footer = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W}" xmlns:r="{R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{paragraph}<w:p/></w:ftr>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#
    );
    let body = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:t>BODY</w:t></w:r></w:p><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>"#
    );
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_owned(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>"#.into(),
        ),
        (
            "_rels/.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.into(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"#.into(),
        ),
        (
            "word/_rels/footer1.xml.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRule" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/rule.png"/></Relationships>"#.into(),
        ),
        ("word/document.xml".to_owned(), body.into_bytes()),
        ("word/styles.xml".to_owned(), styles.into_bytes()),
        ("word/footer1.xml".to_owned(), footer.into_bytes()),
        ("word/media/rule.png".to_owned(), PNG.to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn layout(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76601);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
            })
            .to_string(),
        )
        .unwrap();
    serde_json::from_str(&output).unwrap()
}

/// The footer's first paragraph height and the first page's body bottom margin, in px.
fn rule_height_and_body_bottom(rule_attrs: &str) -> (f64, f64) {
    footer_heights(&document(rule_attrs))
}

fn footer_heights(bytes: &[u8]) -> (f64, f64) {
    let output = layout(bytes);
    let footer = output["headersFooters"]["variants"]
        .as_array()
        .unwrap()
        .iter()
        .find(|variant| variant["kind"] == "footer")
        .unwrap();
    let height = footer["measured"][0]["measure"]["totalHeight"]
        .as_f64()
        .unwrap();
    let page = &output["layout"]["pages"][0];
    let margins = if page["bodyMargins"].is_object() {
        &page["bodyMargins"]
    } else {
        &page["margins"]
    };
    (height, margins["bottom"].as_f64().unwrap())
}

#[test]
fn a_footer_led_by_a_vml_horizontal_rule_keeps_the_authored_bottom_margin() {
    let (height, bottom) = rule_height_and_body_bottom(RULE);
    eprintln!("hr paragraph height {height}px, body bottom margin {bottom}px");
    assert!((height - 3.8).abs() < 0.5, "rule paragraph height {height}");
    assert!(bottom <= 48.5, "body bottom margin {bottom}");
}

#[test]
fn a_footer_led_by_a_plain_vml_picture_keeps_the_marks_line_height() {
    let (height, bottom) = rule_height_and_body_bottom("");
    eprintln!("picture paragraph height {height}px, body bottom margin {bottom}px");
    assert!(height > 15.0, "picture paragraph height {height}");
    assert!(bottom > 48.5, "body bottom margin {bottom}");
}

#[test]
fn a_line_of_two_vml_horizontal_rules_takes_the_rule_height() {
    let paragraph = rule_paragraph(RULE, "", 2).replace("487.05pt", "200pt");
    let (height, _) = footer_heights(&document_with(&paragraph));
    assert!(
        (height - 3.8).abs() < 0.5,
        "two-rule paragraph height {height}"
    );
}

#[test]
fn a_one_and_a_half_spaced_vml_horizontal_rule_keeps_the_marks_line_height() {
    let spacing = r#"<w:spacing w:line="360" w:lineRule="auto"/>"#;
    let (height, _) = footer_heights(&document_with(&rule_paragraph(RULE, spacing, 1)));
    assert!(height > 15.0, "1.5-spaced rule paragraph height {height}");
}
