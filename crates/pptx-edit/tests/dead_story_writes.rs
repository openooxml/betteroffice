use pptx_edit::{
    DeckSession, DeckSnapshot, EditCtx, EditError, EditResult, ShapeDraft, ShapeRect,
    TableInsertPosition, TextParagraphDraft, TextRunDraft, TextStyle, TextStylePatch,
};
use pptx_parse::GraphicFrameData;
use yrs::{ReadTxn, StateVector, Transact};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const TABLE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

enum Removal {
    Slide,
    Shape,
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
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|candidate| &candidate.shapes)
        .find(|shape| shape.id == added.shape_id)
        .unwrap()
        .text_stories[0]
        .id
        .clone()
}

fn session_with_text(text: &str) -> (DeckSession, String) {
    let session = DeckSession::open(DEMO, 84201).unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let story = text_box(&session, &slide, text);
    session.add_undo_barrier();
    (session, story)
}

fn story_home(session: &DeckSession, story: &str) -> (String, String) {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| shape.text_stories.iter().any(|text| text.id == story))
                .map(|shape| (slide.id.clone(), shape.id.clone()))
        })
        .unwrap()
}

fn live_text(session: &DeckSession, story: &str) -> Option<String> {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(|shape| &shape.text_stories)
        .find(|text| text.id == story)
        .map(|text| text.plain_text())
}

fn unmerged_table_session(client_id: u64) -> DeckSession {
    let mut parts = ooxml_opc::unzip_parts(TABLE).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    for attribute in [
        "gridSpan=\"2\"",
        "rowSpan=\"2\"",
        "hMerge=\"1\"",
        "vMerge=\"1\"",
    ] {
        xml = xml.replace(attribute, "");
    }
    *bytes = xml.into_bytes();
    DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), client_id).unwrap()
}

fn table_cell(session: &DeckSession, row: usize, column: usize) -> String {
    let snapshot = session.snapshot().unwrap();
    let table = &snapshot.slides[0].shapes[0];
    let base = format!("story:{}:table:{row}:{column}", table.id);
    table
        .text_stories
        .iter()
        .find(|story| story.id == base || story.id.starts_with(&format!("{base}:")))
        .unwrap()
        .id
        .clone()
}

fn drafts(texts: &[&str]) -> Vec<TextParagraphDraft> {
    texts
        .iter()
        .map(|text| TextParagraphDraft {
            alignment: None,
            bullet: None,
            runs: vec![TextRunDraft {
                text: (*text).to_owned(),
                style: TextStyle::default(),
            }],
        })
        .collect()
}

fn sync(from: &DeckSession, to: &DeckSession) -> DeckSnapshot {
    let state = to.encode_state_vector_v1();
    let diff = from.encode_diff_v1(&state).unwrap();
    to.apply_update_v1(&diff).unwrap()
}

fn remove(session: &DeckSession, slide: &str, shape: &str, removal: Removal) {
    let ctx = EditCtx::local("test");
    match removal {
        Removal::Slide => {
            session.delete_slide(&ctx, slide).unwrap();
        }
        Removal::Shape => {
            session.remove_shape(&ctx, slide, shape).unwrap();
        }
    }
}

fn document_state(session: &DeckSession) -> Vec<u8> {
    session
        .yrs_doc()
        .transact()
        .encode_state_as_update_v1(&StateVector::default())
}

fn paragraph_count(session: &DeckSession, story: &str) -> usize {
    session.story(story).unwrap().paragraphs.len()
}

fn slide_is_live(deck: &DeckSnapshot, slide: &str) -> bool {
    deck.slides.iter().any(|candidate| candidate.id == slide)
}

fn slide_placing(deck: &DeckSnapshot, shape: &str) -> Option<String> {
    deck.slides
        .iter()
        .find(|slide| slide.shapes.iter().any(|placed| placed.id == shape))
        .map(|slide| slide.id.clone())
}

fn table_row_count(deck: &DeckSnapshot, shape: &str) -> usize {
    let placed = deck
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|placed| placed.id == shape)
        .unwrap();
    match &placed.graphic {
        Some(GraphicFrameData::Table(table)) => table.rows.len(),
        _ => panic!("{shape} is not a table"),
    }
}

fn table_of(story: &str) -> Option<&str> {
    story
        .strip_prefix("story:")
        .and_then(|id| id.split_once(":table:"))
        .map(|(shape, _)| shape)
}

fn dead_story(removal: Removal) -> (DeckSession, String) {
    let (session, story) = session_with_text("ABC");
    let ctx = EditCtx::local("test");
    session.insert_paragraph_break(&ctx, &story, 1).unwrap();
    assert_eq!(paragraph_count(&session, &story), 2);
    let (slide, shape) = story_home(&session, &story);
    let slide_removed = matches!(removal, Removal::Slide);
    remove(&session, &slide, &shape, removal);
    let deck = session.snapshot().unwrap();
    assert_eq!(slide_is_live(&deck, &slide), !slide_removed);
    assert_eq!(slide_placing(&deck, &shape), None);
    (session, story)
}

fn ownerless_cell(client_id: u64, retire: bool, removal: Option<Removal>) -> (DeckSession, String) {
    let session = unmerged_table_session(client_id);
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let rows = table_row_count(&snapshot, &shape);
    let cell = table_cell(&session, 1, 1);
    session
        .insert_paragraph_break(&EditCtx::local("test"), &cell, 0)
        .unwrap();
    assert_eq!(paragraph_count(&session, &cell), 2);
    if retire {
        session
            .insert_table_row(
                &EditCtx::local("test"),
                &slide,
                &shape,
                0,
                TableInsertPosition::After,
                &["XY".to_owned()],
            )
            .unwrap();
        assert_eq!(live_text(&session, &cell), None);
    }
    let slide_removed = matches!(removal, Some(Removal::Slide));
    let table_removed = removal.is_some();
    if let Some(removal) = removal {
        remove(&session, &slide, &shape, removal);
    }
    let deck = session.snapshot().unwrap();
    assert_eq!(slide_is_live(&deck, &slide), !slide_removed);
    if table_removed {
        assert_eq!(slide_placing(&deck, &shape), None);
        assert!(
            matches!(session.story(&cell), Err(EditError::ShapeNotFound(id)) if id == shape),
            "{:?}",
            session.story(&cell)
        );
    } else {
        assert_eq!(slide_placing(&deck, &shape), Some(slide));
        assert_eq!(table_row_count(&deck, &shape), rows + 1);
        assert!(
            matches!(session.story(&cell), Err(EditError::StoryNotFound(id)) if id == cell),
            "{:?}",
            session.story(&cell)
        );
    }
    (session, cell)
}

enum Refusal {
    MissingStory,
    MissingShape,
}

fn refused<T: std::fmt::Debug>(
    session: &DeckSession,
    method: &str,
    expected: &Refusal,
    story: &str,
    write: impl FnOnce() -> EditResult<T>,
) {
    let state = document_state(session);
    let result = write();
    let named = match (expected, &result) {
        (Refusal::MissingStory, Err(EditError::StoryNotFound(id))) => id == story,
        (Refusal::MissingShape, Err(EditError::ShapeNotFound(id))) => {
            Some(id.as_str()) == table_of(story)
        }
        _ => false,
    };
    assert!(named, "{method}: {result:?}");
    assert!(
        document_state(session) == state,
        "{method} changed the document"
    );
}

fn assert_every_write_refused(session: &DeckSession, story: &str, expected: Refusal) {
    let deck = session.snapshot().unwrap();
    let text = session.story(story).ok();
    let ctx = EditCtx::local("test");
    let style = TextStyle::default();
    let emphasis = TextStylePatch {
        bold: Some(true),
        ..TextStylePatch::default()
    };
    let center = Some("ctr");
    let grown = drafts(&["A", "BC", "D"]);
    let shrunk = drafts(&["ABC"]);
    if let Some(text) = &text {
        assert!(text.paragraphs.len() > shrunk.len(), "{text:?}");
    }
    let insert_text = || session.insert_text(&ctx, story, 1, "X", &style);
    let delete_text = || session.delete_text(&ctx, story, 0, 1);
    let join = || session.delete_text_joining_paragraphs(&ctx, story, 1, 2);
    let format_text = || session.format_text(&ctx, story, 0, 1, &emphasis);
    let set_alignment = || session.set_paragraph_alignment(&ctx, story, 0, 1, center);
    let insert_break = || session.insert_paragraph_break(&ctx, story, 1);
    let delete_break = || session.delete_paragraph_break(&ctx, story, 1);
    let grow = || session.set_story_paragraphs(&ctx, story, &grown);
    let shrink = || session.set_story_paragraphs(&ctx, story, &shrunk);
    refused(session, "insert_text", &expected, story, insert_text);
    refused(session, "delete_text", &expected, story, delete_text);
    refused(
        session,
        "delete_text_joining_paragraphs",
        &expected,
        story,
        join,
    );
    refused(session, "format_text", &expected, story, format_text);
    refused(
        session,
        "set_paragraph_alignment",
        &expected,
        story,
        set_alignment,
    );
    refused(
        session,
        "insert_paragraph_break",
        &expected,
        story,
        insert_break,
    );
    refused(
        session,
        "delete_paragraph_break",
        &expected,
        story,
        delete_break,
    );
    refused(
        session,
        "set_story_paragraphs adding a paragraph",
        &expected,
        story,
        grow,
    );
    refused(
        session,
        "set_story_paragraphs removing a paragraph",
        &expected,
        story,
        shrink,
    );
    assert_eq!(session.snapshot().unwrap(), deck);
    assert_eq!(session.story(story).ok(), text);
}

#[test]
fn every_story_write_is_refused_once_its_slide_is_deleted() {
    let (session, story) = dead_story(Removal::Slide);
    assert_every_write_refused(&session, &story, Refusal::MissingStory);
}

#[test]
fn every_story_write_is_refused_once_its_shape_is_removed() {
    let (session, story) = dead_story(Removal::Shape);
    assert_every_write_refused(&session, &story, Refusal::MissingStory);
}

#[test]
fn every_cell_write_is_refused_as_a_missing_shape_once_its_table_is_removed() {
    let (session, cell) = ownerless_cell(84208, false, Some(Removal::Shape));
    assert_every_write_refused(&session, &cell, Refusal::MissingShape);
}

#[test]
fn every_cell_write_is_refused_as_a_missing_shape_once_its_slide_is_deleted() {
    let (session, cell) = ownerless_cell(84209, false, Some(Removal::Slide));
    assert_every_write_refused(&session, &cell, Refusal::MissingShape);
}

#[test]
fn every_write_to_a_cell_id_retired_by_a_row_insert_is_refused_as_a_missing_story() {
    let (session, cell) = ownerless_cell(84210, true, None);
    assert_every_write_refused(&session, &cell, Refusal::MissingStory);
}

#[test]
fn every_write_to_a_retired_cell_id_is_refused_as_a_missing_shape_once_its_table_is_removed() {
    let (session, cell) = ownerless_cell(84211, true, Some(Removal::Shape));
    assert_every_write_refused(&session, &cell, Refusal::MissingShape);
}

#[test]
fn every_write_to_a_retired_cell_id_is_refused_as_a_missing_shape_once_its_slide_is_deleted() {
    let (session, cell) = ownerless_cell(84212, true, Some(Removal::Slide));
    assert_every_write_refused(&session, &cell, Refusal::MissingShape);
}

#[test]
fn every_write_to_a_cell_of_a_live_table_on_no_slide_is_refused_as_a_missing_story() {
    let local = unmerged_table_session(84213);
    let peer = unmerged_table_session(84214);
    let snapshot = local.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let cell = table_cell(&local, 1, 1);
    local
        .insert_paragraph_break(&EditCtx::local("test"), &cell, 0)
        .unwrap();
    assert_eq!(paragraph_count(&local, &cell), 2);
    local.add_undo_barrier();
    remove(&local, &slide, &shape, Removal::Shape);
    local.add_undo_barrier();
    assert!(matches!(
        local.story(&cell),
        Err(EditError::ShapeNotFound(id)) if id == shape
    ));
    sync(&local, &peer);
    peer.delete_slide(&EditCtx::local("test"), &slide).unwrap();
    sync(&peer, &local);
    assert!(local.undo());
    let deck = local.snapshot().unwrap();
    assert!(!deck.slides.iter().any(|next| next.id == slide));
    assert!(
        !deck
            .slides
            .iter()
            .flat_map(|next| &next.shapes)
            .any(|placed| placed.id == shape)
    );
    assert!(local.story(&cell).is_ok());
    assert_eq!(paragraph_count(&local, &cell), 2);
    assert_eq!(live_text(&local, &cell), None);
    assert_every_write_refused(&local, &cell, Refusal::MissingStory);
}

#[test]
fn the_encoded_document_state_changes_after_a_successful_story_write() {
    let (session, story) = session_with_text("ABC");
    let ctx = EditCtx::local("test");
    let before = document_state(&session);
    assert_eq!(document_state(&session), before);
    session
        .insert_text(&ctx, &story, 1, "X", &TextStyle::default())
        .unwrap();
    assert_ne!(document_state(&session), before);
}

#[test]
fn insert_text_lands_once_its_slide_deletion_is_undone() {
    let (session, story) = session_with_text("ABC");
    let (slide, _) = story_home(&session, &story);
    let ctx = EditCtx::local("test");
    session.delete_slide(&ctx, &slide).unwrap();
    assert!(session.undo());
    session
        .insert_text(&ctx, &story, 2, "X", &TextStyle::default())
        .unwrap();
    assert_eq!(live_text(&session, &story).as_deref(), Some("ABXC"));
}

#[test]
fn a_new_text_box_takes_writes_after_another_slide_is_deleted() {
    let session = DeckSession::open(DEMO, 84202).unwrap();
    let slides = session.snapshot().unwrap().slides;
    let ctx = EditCtx::local("test");
    session.delete_slide(&ctx, &slides[1].id).unwrap();
    let story = text_box(&session, &slides[0].id, "ABC");
    session
        .insert_text(&ctx, &story, 3, "D", &TextStyle::default())
        .unwrap();
    assert_eq!(live_text(&session, &story).as_deref(), Some("ABCD"));
}

#[test]
fn a_duplicated_slide_takes_writes_after_another_slide_is_deleted() {
    let (session, story) = session_with_text("ABC");
    let (slide, _) = story_home(&session, &story);
    let other = session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find(|candidate| candidate.id != slide)
        .unwrap()
        .id
        .clone();
    let ctx = EditCtx::local("test");
    session.delete_slide(&ctx, &other).unwrap();
    let copy = session.duplicate_slide(&ctx, &slide, 1).unwrap();
    let copied = session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find(|candidate| candidate.id == copy.slide_id)
        .unwrap()
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    assert_ne!(copied, story);
    session
        .insert_text(&ctx, &copied, 0, "X", &TextStyle::default())
        .unwrap();
    assert_eq!(live_text(&session, &copied).as_deref(), Some("XABC"));
    assert_eq!(live_text(&session, &story).as_deref(), Some("ABC"));
}

#[test]
fn a_new_table_row_takes_writes_after_another_slide_is_deleted() {
    let session = unmerged_table_session(84203);
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let ctx = EditCtx::local("test");
    let copy = session.duplicate_slide(&ctx, &slide, 1).unwrap();
    session.add_undo_barrier();
    session.delete_slide(&ctx, &copy.slide_id).unwrap();
    session
        .insert_table_row(
            &ctx,
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &["XY".to_owned()],
        )
        .unwrap();
    let cell = table_cell(&session, 1, 0);
    let before = live_text(&session, &cell).unwrap();
    session
        .insert_text(&ctx, &cell, 0, "Z", &TextStyle::default())
        .unwrap();
    assert_eq!(live_text(&session, &cell), Some(format!("Z{before}")));
}

#[test]
fn a_remote_slide_deletion_applies_on_a_peer_holding_an_anchor_in_it() {
    let local = DeckSession::open(DEMO, 84204).unwrap();
    let peer = DeckSession::open(DEMO, 84205).unwrap();
    let slide = local.snapshot().unwrap().slides[0].id.clone();
    let story = text_box(&local, &slide, "ABC");
    sync(&local, &peer);
    let anchor = peer.anchor_caret(&story, 2).unwrap();
    let ctx = EditCtx::local("test");
    local.delete_slide(&ctx, &slide).unwrap();
    let snapshot = sync(&local, &peer);
    assert!(!snapshot.slides.iter().any(|next| next.id == slide));
    assert_eq!(peer.resolve_caret_anchor(&anchor), None);
}

#[test]
fn a_text_box_added_to_a_slide_a_peer_deleted_concurrently_takes_no_writes() {
    let adder = DeckSession::open(DEMO, 84206).unwrap();
    let deleter = DeckSession::open(DEMO, 84207).unwrap();
    let slide = adder.snapshot().unwrap().slides[0].id.clone();
    let ctx = EditCtx::local("test");
    let story = text_box(&adder, &slide, "ABC");
    adder
        .insert_text(&ctx, &story, 3, "D", &TextStyle::default())
        .unwrap();
    let anchor = adder.anchor_caret(&story, 2).unwrap();
    assert_eq!(adder.resolve_caret_anchor(&anchor), Some(2));
    deleter.delete_slide(&ctx, &slide).unwrap();
    sync(&adder, &deleter);
    sync(&deleter, &adder);
    for peer in [&adder, &deleter] {
        let snapshot = peer.snapshot().unwrap();
        assert!(!snapshot.slides.iter().any(|next| next.id == slide));
        assert_eq!(peer.resolve_caret_anchor(&anchor), None);
        let insert_text = || peer.insert_text(&ctx, &story, 1, "X", &TextStyle::default());
        refused(
            peer,
            "insert_text",
            &Refusal::MissingStory,
            &story,
            insert_text,
        );
    }
}
