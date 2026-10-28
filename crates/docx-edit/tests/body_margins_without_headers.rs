use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

#[allow(dead_code)]
#[path = "support/structured_fixture.rs"]
mod fixture;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn document(header: Option<&str>, footer: Option<&str>, title: bool, even: bool) -> Vec<u8> {
    let references: String = [("header", header), ("footer", footer)]
        .into_iter()
        .filter_map(|(kind, variant)| {
            variant.map(|variant| {
                format!(r#"<w:{kind}Reference w:type="{variant}" r:id="{kind}"/>"#)
            })
        })
        .collect();
    let title = if title { "<w:titlePg/>" } else { "" };
    let section = format!(
        r#"<w:sectPr>{references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="450" w:right="1440" w:bottom="450" w:left="1440" w:header="708" w:footer="708"/>{title}</w:sectPr>"#
    );
    let paragraph = |text: &str, page_break: bool| {
        let page_break = if page_break {
            "<w:pageBreakBefore/>"
        } else {
            ""
        };
        format!(
            r#"<w:p><w:pPr>{page_break}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let body = format!(
        "{}{}{}{section}",
        paragraph("First", false),
        paragraph("Second", true),
        paragraph("Third", true),
    );
    let band = paragraph("Band", false);
    let even = if even { "<w:evenAndOddHeaders/>" } else { "" };
    fixture::Package::new(&body)
        .part(
            "header.xml",
            "header",
            "header",
            "header",
            &format!(r#"<w:hdr {}>{band}</w:hdr>"#, fixture::NS),
        )
        .part(
            "footer.xml",
            "footer",
            "footer",
            "footer",
            &format!(r#"<w:ftr {}>{band}</w:ftr>"#, fixture::NS),
        )
        .part(
            "settings.xml",
            "settings",
            "settings",
            "settings",
            &format!(r#"<w:settings {}>{even}</w:settings>"#, fixture::NS),
        )
        .bytes()
}

fn layout(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76610);
    seed_from_docx(engine.doc(), bytes).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {
                        "sections": [{"properties": package.document.final_section_properties}],
                        "settings": package.settings
                    },
                    "measurement": {
                        "fontChains": {"calibri|0|0": [font]},
                        "defaults": {"fontFamily": "Calibri", "fontSize": 12}
                    }
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn assert_body_margins(output: &Value, expected: [(f64, f64); 3]) {
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), expected.len());
    for (page, (top, bottom)) in pages.iter().zip(expected) {
        let margins = if page["bodyMargins"].is_object() {
            &page["bodyMargins"]
        } else {
            &page["margins"]
        };
        assert!((margins["top"].as_f64().unwrap() - top).abs() < 1e-6);
        assert!((margins["bottom"].as_f64().unwrap() - bottom).abs() < 1e-6);
        assert!((page["fragments"][0]["y"].as_f64().unwrap() - top).abs() < 1e-6);
        assert_eq!(margins["header"].as_f64(), Some(47.2));
        assert_eq!(margins["footer"].as_f64(), Some(47.2));
    }
}

#[test]
fn no_header_reference_keeps_the_top_margin_with_a_default_footer() {
    let output = layout(&document(None, Some("default"), false, false));
    assert_body_margins(&output, [(30.0, 63.2); 3]);
}

#[test]
fn default_header_and_footer_keep_the_existing_extension() {
    let output = layout(&document(Some("default"), Some("default"), false, false));
    assert_body_margins(&output, [(63.2, 63.2); 3]);
}

#[test]
fn no_footer_reference_keeps_the_bottom_margin_with_a_default_header() {
    let output = layout(&document(Some("default"), None, false, false));
    assert_body_margins(&output, [(63.2, 30.0); 3]);
}

#[test]
fn unreferenced_parts_leave_both_body_margins_unchanged() {
    let output = layout(&document(None, None, false, false));
    assert_body_margins(&output, [(30.0, 30.0); 3]);
}

#[test]
fn title_pages_reserve_only_the_referenced_kind() {
    let output = layout(&document(Some("default"), Some("default"), true, false));
    assert_body_margins(&output, [(30.0, 30.0), (63.2, 63.2), (63.2, 63.2)]);
    let output = layout(&document(Some("first"), Some("first"), true, false));
    assert_body_margins(&output, [(63.2, 63.2), (30.0, 30.0), (30.0, 30.0)]);
}

#[test]
fn even_and_odd_pages_reserve_only_the_referenced_kind() {
    let output = layout(&document(Some("default"), Some("default"), false, true));
    assert_body_margins(&output, [(63.2, 63.2), (30.0, 30.0), (63.2, 63.2)]);
    let output = layout(&document(Some("even"), Some("even"), false, true));
    assert_body_margins(&output, [(30.0, 30.0), (63.2, 63.2), (30.0, 30.0)]);
    let output = layout(&document(Some("even"), Some("even"), false, false));
    assert_body_margins(&output, [(30.0, 30.0); 3]);
}
