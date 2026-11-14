use std::sync::Arc;

use pptx_edit::{DeckSession, EditCtx, EditError, ShapeDraft, ShapeRect, TextStyle};
use yrs::{Any, Array, ArrayRef, Map, MapRef, Out, ReadTxn, Transact};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const GROUP_OPEN: &str = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"900\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";
const INNER_GROUP_OPEN: &str = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"901\" name=\"Inner\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";

fn grouped_session(depth: usize, client_id: u64) -> DeckSession {
    let mut parts = ooxml_opc::unzip_parts(DEMO).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert!(xml.contains("<p:sp>"));
    let opens = [GROUP_OPEN, INNER_GROUP_OPEN][..depth].concat();
    *bytes = xml
        .replacen("<p:sp>", &format!("{opens}<p:sp>"), 1)
        .replacen(
            "</p:spTree>",
            &format!("{}</p:spTree>", "</p:grpSp>".repeat(depth)),
            1,
        )
        .into_bytes();
    DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), client_id).unwrap()
}

struct Grouped {
    slide: String,
    path: Vec<String>,
    story: String,
}

fn grouped_story(session: &DeckSession) -> Grouped {
    let slide = session.snapshot().unwrap().slides.remove(0);
    let mut pending: Vec<(Vec<String>, &pptx_edit::ShapeSnapshot)> = slide
        .shapes
        .iter()
        .map(|shape| (vec![shape.id.clone()], shape))
        .collect();
    while let Some((path, shape)) = pending.pop() {
        if path.len() > 1
            && let Some(text) = shape
                .text_stories
                .iter()
                .find(|text| !text.plain_text().is_empty())
        {
            return Grouped {
                slide: slide.id.clone(),
                path,
                story: text.id.clone(),
            };
        }
        for child in &shape.children {
            let mut next = path.clone();
            next.push(child.id.clone());
            pending.push((next, child));
        }
    }
    panic!("no grouped text");
}

fn text_box(session: &DeckSession, slide: &str, text: &str) -> String {
    let added = session
        .add_text_box(
            &EditCtx::local("test"),
            slide,
            &ShapeDraft {
                name: "Box".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                text: text.to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    format!("story:{}:0", added.shape_id)
}

fn sync(from: &DeckSession, to: &DeckSession) {
    let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&diff).unwrap();
}

fn slide_order<T: ReadTxn>(txn: &T, slide: &str) -> ArrayRef {
    txn.get_map("pptx:slides")
        .unwrap()
        .get(txn, slide)
        .unwrap()
        .cast::<MapRef>()
        .unwrap()
        .get(txn, "shapes")
        .unwrap()
        .cast::<ArrayRef>()
        .unwrap()
}

fn shape_entry<T: ReadTxn>(txn: &T, shape: &str) -> MapRef {
    txn.get_map("pptx:shapes")
        .unwrap()
        .get(txn, shape)
        .unwrap()
        .cast::<MapRef>()
        .unwrap()
}

fn children<T: ReadTxn>(txn: &T, shape: &MapRef) -> Vec<String> {
    match shape.get(txn, "children") {
        Some(Out::Any(Any::Array(values))) => values
            .iter()
            .filter_map(|value| match value {
                Any::String(value) => Some(value.to_string()),
                _ => None,
            })
            .collect(),
        _ => Vec::new(),
    }
}

fn string_list(ids: &[String]) -> Any {
    Any::Array(Arc::from(
        ids.iter()
            .map(|id| Any::from(id.as_str()))
            .collect::<Vec<_>>(),
    ))
}

fn order_index<T: ReadTxn>(txn: &T, order: &ArrayRef, id: &str) -> u32 {
    order
        .iter(txn)
        .position(|value| matches!(value, Out::Any(Any::String(entry)) if &*entry == id))
        .unwrap() as u32
}

// No grouping API exists, so ungroup and move_out_of_group edit the document directly.
fn ungroup(session: &DeckSession, slide: &str, group: &str) {
    let mut txn = session.yrs_doc().transact_mut();
    let order = slide_order(&txn, slide);
    let entry = shape_entry(&txn, group);
    let members = children(&txn, &entry);
    let index = order_index(&txn, &order, group);
    order.remove(&mut txn, index);
    for (offset, member) in members.iter().enumerate() {
        order.insert(&mut txn, index + offset as u32, member.as_str());
    }
    entry.insert(&mut txn, "children", string_list(&[]));
    entry.remove(&mut txn, "id");
}

fn move_out_of_group(session: &DeckSession, slide: &str, group: &str, member: &str) {
    let mut txn = session.yrs_doc().transact_mut();
    let order = slide_order(&txn, slide);
    let entry = shape_entry(&txn, group);
    let rest: Vec<String> = children(&txn, &entry)
        .into_iter()
        .filter(|id| id != member)
        .collect();
    entry.insert(&mut txn, "children", string_list(&rest));
    order.push_back(&mut txn, member);
}

#[test]
fn a_cached_owner_does_not_outlive_a_remote_slide_delete() {
    let local = DeckSession::open(DEMO, 84301).unwrap();
    let peer = DeckSession::open(DEMO, 84302).unwrap();
    let slide = local.snapshot().unwrap().slides[0].id.clone();
    let story = text_box(&local, &slide, "ABC");
    sync(&local, &peer);
    let anchor = peer.anchor_caret(&story, 2).unwrap();
    assert_eq!(peer.resolve_caret_anchor(&anchor), Some(2));
    assert_eq!(peer.resolve_caret_anchor(&anchor), Some(2));
    local
        .delete_slide(&EditCtx::local("test"), &slide)
        .unwrap();
    sync(&local, &peer);
    assert_eq!(peer.resolve_caret_anchor(&anchor), None);
}

#[test]
fn a_cached_owner_moves_to_the_duplicate_once_the_original_slide_is_deleted() {
    let session = DeckSession::open(DEMO, 84303).unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let story = text_box(&session, &slide, "ABC");
    let context = EditCtx::local("test");
    let original = session.anchor_caret(&story, 2).unwrap();
    assert_eq!(session.resolve_caret_anchor(&original), Some(2));
    let duplicate = session.duplicate_slide(&context, &slide, 1).unwrap();
    let copy = session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find(|candidate| candidate.id == duplicate.slide_id)
        .unwrap()
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    let copied = session.anchor_caret(&copy, 1).unwrap();
    assert_eq!(session.resolve_caret_anchor(&copied), Some(1));
    session.delete_slide(&context, &slide).unwrap();
    assert_eq!(session.resolve_caret_anchor(&original), None);
    assert_eq!(session.resolve_caret_anchor(&copied), Some(1));
}

#[test]
fn a_cached_owner_follows_a_slide_delete_and_its_undo() {
    let session = DeckSession::open(DEMO, 84304).unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let story = text_box(&session, &slide, "ABC");
    session.add_undo_barrier();
    let anchor = session.anchor_caret(&story, 2).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
    session
        .delete_slide(&EditCtx::local("test"), &slide)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
    assert!(session.undo());
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
}

#[test]
fn two_sessions_on_one_document_keep_their_own_owners() {
    let first = DeckSession::open(DEMO, 84305).unwrap();
    let second = DeckSession::open(DEMO, 84306).unwrap();
    let slide = first.snapshot().unwrap().slides[0].id.clone();
    let story = text_box(&first, &slide, "ABC");
    sync(&first, &second);
    let anchor = first.anchor_caret(&story, 2).unwrap();
    assert_eq!(first.resolve_caret_anchor(&anchor), Some(2));
    second
        .delete_slide(&EditCtx::local("test"), &slide)
        .unwrap();
    assert_eq!(second.resolve_caret_anchor(&anchor), None);
    assert_eq!(first.resolve_caret_anchor(&anchor), Some(2));
    sync(&second, &first);
    assert_eq!(first.resolve_caret_anchor(&anchor), None);
    let context = EditCtx::local("test");
    assert!(
        first
            .insert_text(&context, &story, 0, "X", &TextStyle::default())
            .is_err()
    );
}

#[test]
fn a_grouped_story_resolves_until_its_group_is_removed() {
    let session = grouped_session(1, 84307);
    let grouped = grouped_story(&session);
    let anchor = session.anchor_caret(&grouped.story, 1).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    session
        .insert_text(
            &EditCtx::local("test"),
            &grouped.story,
            0,
            "X",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
    session
        .remove_shape(&EditCtx::local("test"), &grouped.slide, &grouped.path[0])
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
    assert!(
        session
            .insert_text(
                &EditCtx::local("test"),
                &grouped.story,
                0,
                "X",
                &TextStyle::default(),
            )
            .is_err()
    );
}

#[test]
fn an_ungrouped_story_resolves_at_the_top_level_until_its_shape_is_removed() {
    let session = grouped_session(1, 84308);
    let grouped = grouped_story(&session);
    let member = grouped.path[1].clone();
    let anchor = session.anchor_caret(&grouped.story, 1).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    ungroup(&session, &grouped.slide, &grouped.path[0]);
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    session
        .remove_shape(&EditCtx::local("test"), &grouped.slide, &member)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn a_story_moved_out_of_its_group_outlives_the_group_until_it_is_removed() {
    let session = grouped_session(1, 84309);
    let grouped = grouped_story(&session);
    let member = grouped.path[1].clone();
    let anchor = session.anchor_caret(&grouped.story, 1).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    move_out_of_group(&session, &grouped.slide, &grouped.path[0], &member);
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    let context = EditCtx::local("test");
    session
        .remove_shape(&context, &grouped.slide, &grouped.path[0])
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    session
        .remove_shape(&context, &grouped.slide, &member)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn a_story_two_groups_deep_resolves_until_its_outer_group_is_removed() {
    let session = grouped_session(2, 84310);
    let grouped = grouped_story(&session);
    assert!(grouped.path.len() >= 3);
    let anchor = session.anchor_caret(&grouped.story, 1).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    let (outer, inner) = (&grouped.path[0], &grouped.path[1]);
    move_out_of_group(&session, &grouped.slide, outer, inner);
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
    session
        .remove_shape(&EditCtx::local("test"), &grouped.slide, inner)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn a_resolved_story_outlives_unrelated_damage_that_a_fresh_session_reports() {
    let warm = DeckSession::open(DEMO, 84311).unwrap();
    let fresh = DeckSession::open(DEMO, 84312).unwrap();
    let slides = warm.snapshot().unwrap().slides;
    let story = text_box(&warm, &slides[1].id, "ABC");
    sync(&warm, &fresh);
    let anchor = warm.anchor_caret(&story, 2).unwrap();
    assert_eq!(warm.resolve_caret_anchor(&anchor), Some(2));
    let damaged = slides[0].shapes[0].id.clone();
    for session in [&warm, &fresh] {
        let mut txn = session.yrs_doc().transact_mut();
        let entry = shape_entry(&txn, &damaged);
        entry.insert(&mut txn, "children", string_list(&["missing-shape".to_owned()]));
    }
    assert_eq!(warm.resolve_caret_anchor(&anchor), Some(2));
    assert_eq!(fresh.resolve_caret_anchor(&anchor), None);
    let written = fresh.insert_text(
        &EditCtx::local("test"),
        &story,
        0,
        "X",
        &TextStyle::default(),
    );
    assert!(
        matches!(written, Err(EditError::InvalidState(_))),
        "{written:?}"
    );
}
