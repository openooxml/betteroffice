//! Raw ops addressing a story that opens with an astral scalar must match the
//! absolute-index path, which never rounds a position to a code-point boundary.

use docx_edit::{EditCtx, EditingDoc, RawOp};
use yrs::Any;
use yrs::types::Attrs;

const LAPTOP: &str = "\u{1F4BB}";

fn story(text: &str) -> (EditingDoc, EditCtx) {
    let doc = EditingDoc::new(7);
    doc.create_story("body", text, "Normal", "left").unwrap();
    (doc, EditCtx::local(String::new(), String::new()))
}

fn text(doc: &EditingDoc) -> String {
    doc.paragraphs("body")
        .unwrap()
        .iter()
        .map(|paragraph| paragraph.text.clone())
        .collect::<Vec<_>>()
        .join("|")
}

fn bold() -> Attrs {
    let mut attrs = Attrs::new();
    attrs.insert("bold".into(), Any::Bool(true));
    attrs
}

fn insert(index: u32, text: &str) -> RawOp {
    RawOp::Insert {
        index,
        text: text.into(),
        attrs: Attrs::new(),
    }
}

#[test]
fn a_second_insert_lands_ahead_of_the_first() {
    let (doc, ctx) = story(&format!("{LAPTOP}AB"));
    doc.apply_raw_ops("body", vec![insert(1, "Q"), insert(4, "R")], &ctx)
        .unwrap();
    assert_eq!(text(&doc), format!("{LAPTOP}QARB"));
}

#[test]
fn an_insert_after_a_format_keeps_its_position() {
    let (doc, ctx) = story(&format!("{LAPTOP}AB"));
    doc.apply_raw_ops(
        "body",
        vec![
            RawOp::Format {
                index: 0,
                len: 1,
                attrs: bold(),
            },
            insert(3, "R"),
        ],
        &ctx,
    )
    .unwrap();
    assert_eq!(text(&doc), format!("{LAPTOP}ARB"));
}

#[test]
fn an_insert_at_the_last_position_stays_in_the_paragraph() {
    let (doc, ctx) = story(&format!("{LAPTOP}AB"));
    doc.apply_raw_ops("body", vec![insert(1, "Q"), insert(5, "S")], &ctx)
        .unwrap();
    assert_eq!(text(&doc), format!("{LAPTOP}QABS"));
}

#[test]
fn a_delete_reaching_past_a_formatted_astral_scalar_succeeds() {
    let (doc, ctx) = story(&format!("{LAPTOP}AB"));
    doc.apply_raw_ops(
        "body",
        vec![
            RawOp::Format {
                index: 0,
                len: 1,
                attrs: bold(),
            },
            RawOp::Delete { index: 2, len: 3 },
        ],
        &ctx,
    )
    .unwrap();
    assert_eq!(text(&doc), "");
}

#[test]
fn a_delete_after_inserting_an_astral_scalar_succeeds() {
    let (doc, ctx) = story("AB");
    doc.apply_raw_ops(
        "body",
        vec![
            insert(0, LAPTOP),
            insert(1, "X"),
            RawOp::Delete { index: 3, len: 3 },
        ],
        &ctx,
    )
    .unwrap();
    assert_eq!(text(&doc), "");
}
