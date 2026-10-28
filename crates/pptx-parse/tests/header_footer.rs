#[path = "support/header_footer.rs"]
mod support;

use pptx_parse::parse_pptx;
use serde_json::{Value, json};
use support::{LAYOUT, MASTER, deck, field, group, placeholder, text};

#[test]
fn parse_header_footer_settings_defaults_and_boolean_spellings() {
    for (xml, expected) in [
        (
            r#"<p:hf hdr="0" ftr="0" dt="0"/>"#,
            json!({"slideNumber": true, "footer": false, "date": false, "header": false}),
        ),
        (
            r#"<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>"#,
            json!({"slideNumber": false, "footer": false, "date": false, "header": false}),
        ),
        (
            "<p:hf/>",
            json!({"slideNumber": true, "footer": true, "date": true, "header": true}),
        ),
        ("", Value::Null),
        (
            r#"<p:hf sldNum="false" hdr="true" ftr="false" dt="true"/>"#,
            json!({"slideNumber": false, "footer": false, "date": true, "header": true}),
        ),
    ] {
        let shape = group(&placeholder(2, "ftr", &text("Footer"), false))
            + &placeholder(
                3,
                "dt",
                &format!("<a:p>{}</a:p>", field("datetime1", "Date")),
                false,
            );
        let bytes = deck(&[(&shape, false)], xml, xml, "", None);
        let package = parse_pptx(&bytes).unwrap();
        let master = package
            .masters
            .iter()
            .find(|master| master.part_path == MASTER)
            .unwrap();
        let layout = package
            .layouts
            .iter()
            .find(|layout| layout.part_path == LAYOUT)
            .unwrap();
        assert_eq!(
            serde_json::to_value(&master.header_footer).unwrap(),
            expected
        );
        assert_eq!(
            serde_json::to_value(&layout.header_footer).unwrap(),
            expected
        );
    }
}

#[test]
fn parse_layout_does_not_copy_master_header_footer() {
    let bytes = deck(&[("", false)], "", "<p:hf/>", "", None);
    let package = parse_pptx(&bytes).unwrap();
    assert!(package.masters[0].header_footer.is_some());
    assert_eq!(package.layouts[0].header_footer, None);
}
