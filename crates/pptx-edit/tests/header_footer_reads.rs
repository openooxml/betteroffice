#[path = "../../pptx-parse/tests/support/header_footer.rs"]
mod support;

use pptx_edit::{DeckSession, ReadRequest};
use serde_json::{Value, json};
use support::{LAYOUT, MASTER, deck, field, group, placeholder, text};

fn all_placeholders() -> String {
    [
        placeholder(2, "ftr", &text("Footer"), false),
        placeholder(
            3,
            "sldNum",
            &format!("<a:p>{}</a:p>", field("slidenum", "9")),
            false,
        ),
        placeholder(
            4,
            "dt",
            &format!("<a:p>{}</a:p>", field("datetime1", "Stored date")),
            false,
        ),
    ]
    .join("")
}

#[test]
fn read_header_footer_text_ids_and_json() {
    let bytes = deck(&[(&all_placeholders(), false)], "", "", "", None);
    let session = DeckSession::open(&bytes, 910).unwrap();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    let slide = &read.slides[0];
    let facts = slide.header_footer.as_ref().unwrap();
    assert_eq!(facts.footer.as_ref().unwrap().shape_id, slide.shapes[0].id);
    assert_eq!(
        facts.slide_number.as_ref().unwrap().shape_id,
        slide.shapes[1].id
    );
    assert_eq!(facts.date.as_ref().unwrap().shape_id, slide.shapes[2].id);
    assert_eq!(
        serde_json::to_value(slide).unwrap()["headerFooter"],
        json!({
            "footer": {"shapeId": slide.shapes[0].id, "text": "Footer"},
            "slideNumber": {"shapeId": slide.shapes[1].id, "text": "1"},
            "date": {"shapeId": slide.shapes[2].id, "text": "Stored date", "field": "datetime1"}
        })
    );
}

#[test]
fn read_fixed_date_and_empty_footer() {
    let shapes =
        placeholder(2, "ftr", "<a:p/>", false) + &placeholder(3, "dt", &text("Fixed date"), false);
    let session = DeckSession::open(&deck(&[(&shapes, false)], "", "", "", None), 911).unwrap();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    let facts = read.slides[0].header_footer.as_ref().unwrap();
    assert_eq!(facts.footer.as_ref().unwrap().text, "");
    let date = facts.date.as_ref().unwrap();
    assert_eq!(date.text, "Fixed date");
    assert_eq!(date.field, None);
}

#[test]
fn read_date_reports_first_datetime_field_and_stored_text() {
    let paragraphs = format!(
        "<a:p>{}{}{}{}<a:r><a:t> literal</a:t></a:r></a:p>",
        field("custom", "Prefix"),
        field("datetime2", "First date"),
        field("datetime1", "Second date"),
        field("datetime99", "Other date")
    );
    let shape = placeholder(2, "dt", &paragraphs, false);
    let session = DeckSession::open(&deck(&[(&shape, false)], "", "", "", None), 919).unwrap();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    let date = read.slides[0]
        .header_footer
        .as_ref()
        .unwrap()
        .date
        .as_ref()
        .unwrap();
    assert_eq!(date.field.as_deref(), Some("datetime2"));
    assert_eq!(date.text, "PrefixFirst dateSecond dateOther date literal");
}

#[test]
fn read_slide_placeholders_ignore_stored_settings_and_parent_text() {
    let off = r#"<p:hf sldNum="0" ftr="0" dt="0"/>"#;
    for own in [true, false] {
        let shapes = all_placeholders();
        let bytes = deck(
            &[(if own { &shapes } else { "" }, false)],
            off,
            off,
            &shapes,
            None,
        );
        let session = DeckSession::open(&bytes, 912).unwrap();
        let read = session
            .read_content(&ReadRequest::default())
            .unwrap()
            .unwrap();
        assert!(read.slides[0].header_footer.is_some());
        let facts = read.slides[0].header_footer.as_ref().unwrap();
        assert_eq!(facts.footer.is_some(), own);
        assert_eq!(facts.slide_number.is_some(), own);
        assert_eq!(facts.date.is_some(), own);
    }
}

#[test]
fn read_hidden_placeholders_are_absent() {
    let shapes = ["ftr", "sldNum", "dt"]
        .iter()
        .enumerate()
        .map(|(index, kind)| placeholder(index as u32 + 2, kind, &text("Hidden"), true))
        .collect::<String>();
    let session = DeckSession::open(&deck(&[(&shapes, false)], "", "", "", None), 913).unwrap();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    assert_eq!(
        serde_json::to_value(&read.slides[0]).unwrap()["headerFooter"],
        json!({
            "footer": null, "slideNumber": null, "date": null
        })
    );
}

#[test]
fn read_first_placeholder_in_group_order_and_multiline_text() {
    let first = placeholder(2, "ftr", &(text("First") + &text("Second")), false);
    let shapes = group(&first) + &placeholder(3, "ftr", &text("Later"), false);
    let session = DeckSession::open(&deck(&[(&shapes, false)], "", "", "", None), 914).unwrap();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    let slide = &read.slides[0];
    let footer = slide
        .header_footer
        .as_ref()
        .unwrap()
        .footer
        .as_ref()
        .unwrap();
    assert_eq!(footer.text, "First\nSecond");
    assert_eq!(footer.shape_id, slide.shapes[0].children[0].id);
}

#[test]
fn read_slide_numbers_count_hidden_slides_and_preserve_literals() {
    let paragraphs = format!(
        "<a:p><a:r><a:t>Seite </a:t></a:r>{}</a:p>",
        field("slidenum", "‹#›")
    );
    let shape = placeholder(2, "sldNum", &paragraphs, false);
    let bytes = deck(
        &[(&shape, false), (&shape, true), (&shape, false)],
        "",
        "",
        "",
        Some(5),
    );
    let session = DeckSession::open(&bytes, 915).unwrap();
    let before = session.encode_state_as_update_v1();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    assert_eq!(session.encode_state_as_update_v1(), before);
    let texts = read
        .slides
        .iter()
        .map(|slide| {
            slide
                .header_footer
                .as_ref()
                .unwrap()
                .slide_number
                .as_ref()
                .unwrap()
                .text
                .as_str()
        })
        .collect::<Vec<_>>();
    assert_eq!(texts, ["Seite 5", "Seite 6", "Seite 7"]);
    let request = ReadRequest {
        slide_ids: Some(vec![read.slides[1].id.clone()]),
        ..ReadRequest::default()
    };
    let filtered = session.read_content(&request).unwrap().unwrap();
    assert_eq!(filtered.slides.len(), 1);
    assert_eq!(
        filtered.slides[0]
            .header_footer
            .as_ref()
            .unwrap()
            .slide_number
            .as_ref()
            .unwrap()
            .text,
        "Seite 6"
    );
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn read_slide_number_fixed_text_and_stale_fields() {
    for (paragraphs, expected) in [
        (text("2"), "2"),
        (format!("<a:p>{}</a:p>", field("slidenum", "9")), "5"),
        (
            format!(
                "<a:p>{}{}</a:p>",
                field("slidenum", "9"),
                field("slidenum", "0")
            ),
            "55",
        ),
    ] {
        let shape = placeholder(2, "sldNum", &paragraphs, false);
        let session =
            DeckSession::open(&deck(&[(&shape, false)], "", "", "", Some(5)), 916).unwrap();
        let read = session
            .read_content(&ReadRequest::default())
            .unwrap()
            .unwrap();
        assert!(read.slides[0].header_footer.is_some());
        assert_eq!(
            read.slides[0]
                .header_footer
                .as_ref()
                .unwrap()
                .slide_number
                .as_ref()
                .unwrap()
                .text,
            expected
        );
    }
}

#[test]
fn read_layout_header_footer_inheritance_defaults_and_override() {
    let master = r#"<p:hf hdr="0" ftr="0" dt="0"/>"#;
    for (layout_hf, master_hf, expected) in [
        (
            "",
            master,
            json!({"slideNumber": true, "footer": false, "date": false, "header": false}),
        ),
        (
            r#"<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>"#,
            master,
            json!({"slideNumber": false, "footer": false, "date": false, "header": false}),
        ),
        (
            "<p:hf/>",
            master,
            json!({"slideNumber": true, "footer": true, "date": true, "header": true}),
        ),
        ("", "", Value::Null),
        (
            r#"<p:hf sldNum="false" hdr="true" ftr="false" dt="true"/>"#,
            master,
            json!({"slideNumber": false, "footer": false, "date": true, "header": true}),
        ),
    ] {
        let session =
            DeckSession::open(&deck(&[("", false)], layout_hf, master_hf, "", None), 917).unwrap();
        let read = session
            .read_content(&ReadRequest::default())
            .unwrap()
            .unwrap();
        assert!(read.slides[0].header_footer.is_some());
        assert_eq!(
            serde_json::to_value(read.slides[0].layout.as_ref().unwrap()).unwrap()["headerFooter"],
            expected
        );
    }
}

#[test]
fn save_preserves_header_footer_source_parts() {
    let bytes = deck(
        &[(&all_placeholders(), false)],
        "<p:hf/>",
        r#"<p:hf dt="false" ftr="0"/>"#,
        "",
        None,
    );
    let session = DeckSession::open(&bytes, 918).unwrap();
    let saved = session.save().unwrap();
    let source = ooxml_opc::unzip_parts(&bytes).unwrap();
    let output = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in [MASTER, LAYOUT] {
        assert_eq!(
            source.iter().find(|(name, _)| name == path).unwrap().1,
            output.iter().find(|(name, _)| name == path).unwrap().1
        );
    }
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.slides[0].header_footer.is_some());
    assert!(
        read.slides[0]
            .layout
            .as_ref()
            .unwrap()
            .header_footer
            .is_some()
    );
}
