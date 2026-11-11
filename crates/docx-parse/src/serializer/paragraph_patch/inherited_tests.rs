use serde_json::{Value, json};

use super::Failure;
use super::properties::run_properties;
use crate::paragraph_patch::{AttrDeltaV1, Attrs, SlotV1};

fn attrs(value: Value) -> Attrs {
    serde_json::from_value(value).unwrap()
}

fn slot(value: Option<Value>) -> SlotV1 {
    value.map_or(SlotV1::Missing, |value| SlotV1::Present { value })
}

fn delta(path: &[&str], before: Option<Value>, after: Option<Value>) -> AttrDeltaV1 {
    AttrDeltaV1 {
        path: path.iter().map(|key| (*key).to_owned()).collect(),
        before: slot(before),
        after: slot(after),
    }
}

fn written(source: &str, deltas: &[AttrDeltaV1], inherit: Value, target: Value) -> String {
    written_in(source, deltas, inherit, target, true)
}

fn written_in(
    source: &str,
    deltas: &[AttrDeltaV1],
    inherit: Value,
    target: Value,
    current: bool,
) -> String {
    let result: Result<String, Failure> =
        run_properties(source, deltas, &attrs(inherit), &attrs(target), current);
    result.unwrap_or_else(|error| panic!("{}", error.detail))
}

#[test]
fn without_current_inheritance_a_changed_value_stays_explicit() {
    let on = written_in(
        r#"<w:rPr><w:b w:val="0"/></w:rPr>"#,
        &[delta(&["bold"], Some(json!(false)), Some(json!(true)))],
        heading(),
        heading(),
        false,
    );
    assert_eq!(on, "<w:rPr><w:b/></w:rPr>");
    for inherit in [heading(), json!({})] {
        let off = written_in(
            "<w:rPr><w:b/></w:rPr>",
            &[delta(&["bold"], Some(json!(true)), Some(json!(false)))],
            inherit,
            json!({"bold": false}),
            false,
        );
        assert_eq!(off, r#"<w:rPr><w:b w:val="0"/></w:rPr>"#);
    }
    let sized = written_in(
        r#"<w:rPr><w:sz w:val="40"/></w:rPr>"#,
        &[delta(
            &["fontSize", "size"],
            Some(json!(40)),
            Some(json!(32)),
        )],
        heading(),
        heading(),
        false,
    );
    assert_eq!(sized, r#"<w:rPr><w:sz w:val="32"/></w:rPr>"#);
}

fn heading() -> Value {
    json!({"bold": true, "fontSize": {"size": 32}})
}

#[test]
fn a_value_changed_to_what_the_style_gives_removes_the_element() {
    let cleared = written(
        r#"<w:rPr><w:b w:val="0"/><w:i/><w:color w:val="C00000"/></w:rPr>"#,
        &[
            delta(&["bold"], Some(json!(false)), Some(json!(true))),
            delta(&["italic"], Some(json!(true)), None),
            delta(&["textColor"], Some(json!({"rgb": "C00000"})), None),
        ],
        heading(),
        heading(),
    );
    assert_eq!(cleared, "");
    let sized = written(
        r#"<w:rPr><w:i/><w:sz w:val="40"/></w:rPr>"#,
        &[delta(
            &["fontSize", "size"],
            Some(json!(40)),
            Some(json!(32)),
        )],
        heading(),
        json!({"bold": true, "italic": true, "fontSize": {"size": 32}}),
    );
    assert_eq!(sized, "<w:rPr><w:i/></w:rPr>");
    let colored = written(
        "",
        &[delta(&["textColor", "rgb"], None, Some(json!("1F3864")))],
        json!({"textColor": {"rgb": "1F3864"}}),
        json!({"textColor": {"rgb": "1F3864"}}),
    );
    assert_eq!(colored, "");
}

#[test]
fn a_value_changed_away_from_the_style_is_written() {
    let off = written(
        "",
        &[delta(&["bold"], Some(json!(true)), Some(json!(false)))],
        heading(),
        json!({"bold": false, "fontSize": {"size": 32}}),
    );
    assert_eq!(off, r#"<w:rPr><w:b w:val="0"/></w:rPr>"#);
    let on = written(
        "",
        &[delta(&["bold"], None, Some(json!(true)))],
        json!({}),
        json!({"bold": true}),
    );
    assert_eq!(on, "<w:rPr><w:b/></w:rPr>");
    let sized = written(
        r#"<w:rPr><w:sz w:val="40"/></w:rPr>"#,
        &[delta(
            &["fontSize", "size"],
            Some(json!(40)),
            Some(json!(22)),
        )],
        heading(),
        json!({"bold": true, "fontSize": {"size": 22}}),
    );
    assert_eq!(sized, r#"<w:rPr><w:sz w:val="22"/></w:rPr>"#);
}

#[test]
fn removing_the_last_property_drops_the_container() {
    let body = written(
        "<w:rPr><w:b/></w:rPr>",
        &[delta(&["bold"], Some(json!(true)), None)],
        json!({}),
        json!({}),
    );
    assert_eq!(body, "");
    let underline = written(
        r#"<w:rPr><w:u w:val="single" w:color="FF0000"/><w:vertAlign w:val="superscript"/></w:rPr>"#,
        &[
            delta(
                &["underline"],
                Some(json!({"style": "single", "color": "FF0000"})),
                None,
            ),
            delta(&["superscript"], Some(json!(true)), None),
        ],
        json!({}),
        json!({}),
    );
    assert_eq!(underline, "");
}

#[test]
fn keys_without_a_change_keep_their_source_bytes() {
    let source = "<w:rPr><w:b/><w:i/></w:rPr>";
    assert_eq!(written(source, &[], heading(), heading()), source);
    let italic_cleared = written(
        source,
        &[delta(&["italic"], Some(json!(true)), None)],
        heading(),
        heading(),
    );
    assert_eq!(italic_cleared, "<w:rPr><w:b/></w:rPr>");
}
