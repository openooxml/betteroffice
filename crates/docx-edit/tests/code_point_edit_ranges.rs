use std::collections::BTreeMap;

use docx_edit::{
    ChangeTarget, EditCtx, EditingDoc, FormatPolicy, InlineFormatDelta, Patch, Position, Receipt,
    RichRun, SegmentContent, SimpleFormat, StoryRange,
};
use yrs::Any;

fn fixture(first: &str) -> (EditingDoc, EditCtx) {
    let doc = EditingDoc::new(7);
    let ctx = EditCtx::local("Ada", "2026-10-03T00:00:00Z");
    doc.create_story("body", &format!("{first}b"), "Normal", "left")
        .unwrap();
    doc.split_paragraph(
        &ctx,
        Position::new("body", first.encode_utf16().count() as u32),
        None,
    )
    .unwrap();
    (doc, ctx)
}

fn story() -> (EditingDoc, EditCtx) {
    fixture("a😀")
}

fn range(start: u32, end: u32) -> StoryRange {
    StoryRange::new("body", start, end)
}

fn texts(doc: &EditingDoc) -> Vec<String> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

#[track_caller]
fn assert_texts(doc: &EditingDoc, expected: &[&str]) {
    assert_eq!(texts(doc), expected);
}

#[track_caller]
fn assert_round_trips(doc: &EditingDoc) {
    let peer = EditingDoc::new(8);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(texts(&peer), texts(doc));
}

#[track_caller]
fn assert_receipt(doc: &EditingDoc, receipt: &Receipt, start: u32, end: u32) {
    let reported = receipt.range.as_ref().unwrap();
    assert_eq!(doc.locate_range(reported).unwrap(), range(start, end));
}

#[track_caller]
fn assert_same_story(doc: &EditingDoc, expected: &EditingDoc) {
    assert_eq!(
        doc.paragraphs("body").unwrap(),
        expected.paragraphs("body").unwrap()
    );
    assert_eq!(
        doc.story_segments("body").unwrap(),
        expected.story_segments("body").unwrap()
    );
    assert_round_trips(expected);
}

fn bold_run() -> [RichRun; 1] {
    [RichRun {
        text: "x".into(),
        attrs: BTreeMap::from([("bold".into(), Any::Bool(true))]),
    }]
}

fn image_payload() -> Vec<(String, Any)> {
    vec![("rId".into(), Any::from("rIdImage"))]
}

fn leading_embed_story() -> (EditingDoc, EditCtx) {
    let (doc, ctx) = fixture("😀");
    doc.insert_embed(&ctx, Position::new("body", 0), "image", image_payload())
        .unwrap();
    assert_texts(&doc, &["😀", "b"]);
    assert_eq!(doc.story_len("body").unwrap(), 6);
    (doc, ctx)
}

#[test]
fn delete_start_inside_pair_keeps_the_paragraph_mark() {
    let (doc, ctx) = story();
    let receipt = doc.delete_range(&ctx, range(2, 3)).unwrap();
    assert_texts(&doc, &["a", "b"]);
    assert_receipt(&doc, &receipt, 1, 1);
    assert_round_trips(&doc);
}

#[test]
fn delete_double_click_range_removes_the_whole_emoji() {
    let (doc, ctx) = story();
    let receipt = doc.delete_range(&ctx, range(1, 2)).unwrap();
    assert_texts(&doc, &["a", "b"]);
    assert_receipt(&doc, &receipt, 1, 1);
    assert_round_trips(&doc);
}

#[test]
fn delete_three_units_from_inside_pair_keeps_the_final_mark() {
    let (doc, ctx) = story();
    let receipt = doc.delete_range(&ctx, range(2, 5)).unwrap();
    assert_texts(&doc, &["a"]);
    assert_receipt(&doc, &receipt, 1, 1);
    assert_round_trips(&doc);
}

#[test]
fn delete_end_inside_pair_keeps_an_empty_first_paragraph() {
    let (doc, ctx) = story();
    let receipt = doc.delete_range(&ctx, range(0, 2)).unwrap();
    assert_texts(&doc, &["", "b"]);
    assert_receipt(&doc, &receipt, 0, 0);
    assert_round_trips(&doc);
}

#[test]
fn insert_text_inside_pair_lands_before_the_emoji() {
    let (doc, ctx) = story();
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    let receipt = doc
        .insert_text(&ctx, Position::new("body", 2), "x", FormatPolicy::Inherit)
        .unwrap();
    assert_texts(&doc, &["ax😀", "b"]);
    let reported = receipt.range.as_ref().unwrap();
    assert_eq!(reported.start.para, para_id);
    assert_eq!(reported.start.offset, 1);
    assert_eq!(reported.end.offset, 2);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn replace_double_click_range_replaces_the_whole_emoji() {
    let (doc, ctx) = story();
    let receipt = doc.replace_range(&ctx, range(1, 2), "x").unwrap();
    assert_texts(&doc, &["ax", "b"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn collapsed_replace_inside_pair_lands_before_the_emoji() {
    let (doc, ctx) = story();
    let receipt = doc.replace_range(&ctx, range(2, 2), "x").unwrap();
    assert_texts(&doc, &["ax😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn replace_start_inside_pair_and_mark_merges_paragraphs() {
    let (doc, ctx) = story();
    let receipt = doc.replace_range(&ctx, range(2, 4), "x").unwrap();
    assert_texts(&doc, &["axb"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn rich_replace_double_click_range_replaces_the_whole_emoji() {
    let (doc, ctx) = story();
    let receipt = doc
        .replace_range_rich(&ctx, range(1, 2), &bold_run())
        .unwrap();
    assert_texts(&doc, &["ax", "b"]);
    assert_receipt(&doc, &receipt, 1, 2);
    let segments = doc.story_segments("body").unwrap();
    let inserted = segments
        .iter()
        .find(|segment| segment.content == SegmentContent::Text("x".into()))
        .unwrap();
    assert_eq!(inserted.attributes.get("bold"), Some(&Any::Bool(true)));
    assert_round_trips(&doc);
}

#[test]
fn collapsed_rich_replace_inside_pair_lands_before_the_emoji() {
    let (doc, ctx) = story();
    let receipt = doc
        .replace_range_rich(&ctx, range(2, 2), &bold_run())
        .unwrap();
    assert_texts(&doc, &["ax😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn rich_replace_start_inside_pair_and_mark_merges_paragraphs() {
    let (doc, ctx) = story();
    let receipt = doc
        .replace_range_rich(&ctx, range(2, 4), &bold_run())
        .unwrap();
    assert_texts(&doc, &["axb"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn suggested_delete_double_click_range_accepts_the_whole_emoji() {
    let (doc, ctx) = story();
    let receipt = doc
        .delete_range(&ctx.clone().suggesting(), range(1, 2))
        .unwrap();
    assert_texts(&doc, &["a😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 3);
    assert_round_trips(&doc);
    doc.accept_change(
        &ctx,
        &ChangeTarget::Range(range(0, doc.story_len("body").unwrap())),
    )
    .unwrap();
    assert_texts(&doc, &["a", "b"]);
    assert_round_trips(&doc);
}

#[test]
fn suggested_replace_double_click_range_accepts_the_whole_emoji() {
    let (doc, ctx) = story();
    let receipt = doc
        .replace_range(&ctx.clone().suggesting(), range(1, 2), "x")
        .unwrap();
    assert_texts(&doc, &["a😀x", "b"]);
    assert_receipt(&doc, &receipt, 3, 4);
    assert_round_trips(&doc);
    doc.accept_change(
        &ctx,
        &ChangeTarget::Range(range(0, doc.story_len("body").unwrap())),
    )
    .unwrap();
    assert_texts(&doc, &["ax", "b"]);
    assert_round_trips(&doc);
}

#[test]
fn accept_range_start_inside_pair_removes_the_whole_suggested_deletion() {
    let (doc, ctx) = story();
    doc.delete_range(&ctx.clone().suggesting(), range(1, 3))
        .unwrap();
    let receipt = doc
        .accept_change(&ctx, &ChangeTarget::Range(range(2, 3)))
        .unwrap();
    assert_texts(&doc, &["a", "b"]);
    assert_receipt(&doc, &receipt, 1, 1);
    assert!(doc.list_changes("body").unwrap().is_empty());
    assert_round_trips(&doc);
}

#[test]
fn reject_range_end_inside_pair_removes_the_whole_suggested_insertion() {
    let (doc, ctx) = fixture("a");
    doc.insert_text(
        &ctx.clone().suggesting(),
        Position::new("body", 1),
        "😀",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let receipt = doc
        .reject_change(&ctx, &ChangeTarget::Range(range(1, 2)))
        .unwrap();
    assert_texts(&doc, &["a", "b"]);
    assert_receipt(&doc, &receipt, 1, 1);
    assert!(doc.list_changes("body").unwrap().is_empty());
    assert_round_trips(&doc);
}

#[test]
fn accept_range_end_inside_pair_clears_the_whole_insertion_stamp() {
    let (doc, ctx) = fixture("a");
    doc.insert_text(
        &ctx.clone().suggesting(),
        Position::new("body", 1),
        "😀",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let receipt = doc
        .accept_change(&ctx, &ChangeTarget::Range(range(1, 2)))
        .unwrap();
    assert_texts(&doc, &["a😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 3);
    assert!(doc.list_changes("body").unwrap().is_empty());
    assert_round_trips(&doc);
}

#[test]
fn reject_range_start_inside_pair_clears_the_whole_deletion_stamp() {
    let (doc, ctx) = story();
    doc.delete_range(&ctx.clone().suggesting(), range(1, 3))
        .unwrap();
    let receipt = doc
        .reject_change(&ctx, &ChangeTarget::Range(range(2, 3)))
        .unwrap();
    assert_texts(&doc, &["a😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 3);
    assert!(doc.list_changes("body").unwrap().is_empty());
    assert_round_trips(&doc);
}

#[test]
fn hard_break_inside_pair_matches_insertion_before_the_emoji() {
    let (doc, ctx) = story();
    let (expected, _) = story();
    let receipt = doc
        .insert_hard_break(&ctx, Position::new("body", 2))
        .unwrap();
    expected
        .insert_hard_break(&ctx, Position::new("body", 1))
        .unwrap();
    assert_receipt(&doc, &receipt, 1, 2);
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn tab_inside_pair_lands_before_the_emoji() {
    let (doc, ctx) = story();
    let receipt = doc.insert_tab(&ctx, Position::new("body", 2)).unwrap();
    assert_texts(&doc, &["a\t😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_round_trips(&doc);
}

#[test]
fn split_inside_pair_moves_the_whole_emoji_to_the_next_paragraph() {
    let (doc, ctx) = story();
    doc.split_paragraph(&ctx, Position::new("body", 2), None)
        .unwrap();
    assert_texts(&doc, &["a", "😀", "b"]);
    assert_round_trips(&doc);
}

#[test]
fn inline_embed_inside_pair_matches_insertion_before_the_emoji() {
    let (doc, ctx) = story();
    let (expected, _) = story();
    let receipt = doc
        .insert_embed(&ctx, Position::new("body", 2), "image", image_payload())
        .unwrap();
    expected
        .insert_embed(&ctx, Position::new("body", 1), "image", image_payload())
        .unwrap();
    assert_receipt(&doc, &receipt, 1, 2);
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn table_inside_pair_matches_insertion_before_the_emoji() {
    let (doc, ctx) = story();
    let (expected, _) = story();
    let receipt = doc
        .insert_table(&ctx, Position::new("body", 2), 1, 1)
        .unwrap();
    let expected_receipt = expected
        .insert_table(&ctx, Position::new("body", 1), 1, 1)
        .unwrap();
    assert_eq!(receipt, expected_receipt);
    for story_id in &receipt.created_story_ids {
        assert_eq!(
            doc.paragraphs(story_id).unwrap(),
            expected.paragraphs(story_id).unwrap()
        );
    }
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn toggle_bold_double_click_range_formats_the_whole_emoji() {
    let (doc, ctx) = story();
    let (expected, _) = story();
    let receipt = doc
        .toggle_format(&ctx, range(1, 2), SimpleFormat::Bold)
        .unwrap();
    expected
        .toggle_format(&ctx, range(1, 3), SimpleFormat::Bold)
        .unwrap();
    assert_receipt(&doc, &receipt, 1, 3);
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn format_range_start_inside_pair_formats_the_whole_emoji() {
    let (doc, ctx) = story();
    let (expected, _) = story();
    let delta = InlineFormatDelta {
        bold: Patch::Set(true),
        ..Default::default()
    };
    let receipt = doc.format_range(&ctx, range(2, 3), &delta).unwrap();
    expected.format_range(&ctx, range(1, 3), &delta).unwrap();
    assert_receipt(&doc, &receipt, 1, 3);
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn hyperlink_double_click_range_links_the_whole_emoji() {
    let (doc, ctx) = story();
    let (expected, _) = story();
    let hyperlink = Any::from("https://example.com");
    let receipt = doc
        .set_hyperlink(&ctx, range(1, 2), Some(hyperlink.clone()))
        .unwrap();
    expected
        .set_hyperlink(&ctx, range(1, 3), Some(hyperlink))
        .unwrap();
    assert_receipt(&doc, &receipt, 1, 3);
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn comment_partial_pair_ranges_anchor_the_whole_emoji() {
    for partial in [range(1, 2), range(2, 3)] {
        let (doc, _) = story();
        let (expected, _) = story();
        let comment = doc
            .add_comment(&[partial], "Ada", "", Any::from("note"))
            .unwrap();
        let full = expected
            .add_comment(&[range(1, 3)], "Ada", "", Any::from("note"))
            .unwrap();
        let anchors = doc.resolve_comment(&comment).unwrap();
        assert_eq!(anchors, expected.resolve_comment(&full).unwrap());
        assert_eq!((anchors[0].start, anchors[0].end), (1, 3));
        assert_round_trips(&expected);
        assert_round_trips(&doc);
    }
}

#[test]
fn triple_click_text_length_after_embed_deletes_the_embed_and_whole_emoji() {
    let (doc, ctx) = leading_embed_story();
    let end = doc.paragraphs("body").unwrap()[0]
        .text
        .encode_utf16()
        .count() as u32;
    assert_eq!(end, 2);
    let receipt = doc.delete_range(&ctx, range(0, end)).unwrap();
    assert_texts(&doc, &["", "b"]);
    assert_eq!(
        doc.embed_kind(&Position::new("body", 0))
            .unwrap()
            .as_deref(),
        Some("pilcrow")
    );
    assert_receipt(&doc, &receipt, 0, 0);
    assert_round_trips(&doc);
}

#[test]
fn right_arrow_offset_after_embed_inserts_before_the_emoji() {
    let (doc, ctx) = leading_embed_story();
    let (expected, _) = leading_embed_story();
    let receipt = doc
        .insert_text(&ctx, Position::new("body", 2), "x", FormatPolicy::Inherit)
        .unwrap();
    expected
        .insert_text(&ctx, Position::new("body", 1), "x", FormatPolicy::Inherit)
        .unwrap();
    assert_texts(&doc, &["x😀", "b"]);
    assert_receipt(&doc, &receipt, 1, 2);
    assert_same_story(&doc, &expected);
    assert_round_trips(&doc);
}

#[test]
fn backspace_fallback_after_embed_deletes_the_emoji_and_keeps_the_embed() {
    let (doc, ctx) = leading_embed_story();
    let receipt = doc.delete_range(&ctx, range(2, 3)).unwrap();
    assert_texts(&doc, &["", "b"]);
    assert_eq!(
        doc.embed_kind(&Position::new("body", 0))
            .unwrap()
            .as_deref(),
        Some("image")
    );
    assert_receipt(&doc, &receipt, 1, 1);
    assert_round_trips(&doc);
}

#[test]
fn code_point_boundary_edits_keep_their_existing_offsets() {
    for (start, end, expected) in [
        (0, 1, vec!["😀", "b"]),
        (1, 3, vec!["a", "b"]),
        (3, 4, vec!["a😀b"]),
    ] {
        let (doc, ctx) = story();
        let receipt = doc.delete_range(&ctx, range(start, end)).unwrap();
        assert_texts(&doc, &expected);
        assert_receipt(&doc, &receipt, start, start);
        assert_round_trips(&doc);
    }
    for (offset, expected) in [(1, "ax😀"), (3, "a😀x")] {
        let (doc, ctx) = story();
        let receipt = doc
            .insert_text(
                &ctx,
                Position::new("body", offset),
                "x",
                FormatPolicy::Inherit,
            )
            .unwrap();
        assert_texts(&doc, &[expected, "b"]);
        assert_receipt(&doc, &receipt, offset, offset + 1);
        assert_round_trips(&doc);
    }
    for (start, end, expected) in [(1, 3, "ax"), (3, 3, "a😀x")] {
        let (doc, ctx) = story();
        let receipt = doc.replace_range(&ctx, range(start, end), "x").unwrap();
        assert_texts(&doc, &[expected, "b"]);
        assert_receipt(&doc, &receipt, start, start + 1);
        assert_round_trips(&doc);
    }
}
