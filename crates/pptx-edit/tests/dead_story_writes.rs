use pptx_edit::{
    DeckSession, DeckSnapshot, EditCtx, EditError, EditResult, ShapeDraft, ShapeRect,
    TableInsertPosition, TextParagraphDraft, TextRunDraft, TextStyle, TextStylePatch,
};

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

fn dead_story(removal: Removal) -> (DeckSession, String) {
    let (session, story) = session_with_text("ABC");
    let ctx = EditCtx::local("test");
    session.insert_paragraph_break(&ctx, &story, 1).unwrap();
    let (slide, shape) = story_home(&session, &story);
    remove(&session, &slide, &shape, removal);
    (session, story)
}

fn ownerless_cell(client_id: u64, retire: bool, removal: Option<Removal>) -> (DeckSession, String) {
    let session = unmerged_table_session(client_id);
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let cell = table_cell(&session, 1, 1);
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
    if let Some(removal) = removal {
        remove(&session, &slide, &shape, removal);
    }
    (session, cell)
}

enum Refusal {
    MissingStory,
    MissingShape,
}

fn refused<T: std::fmt::Debug>(method: &str, expected: &Refusal, result: EditResult<T>) {
    let named = match expected {
        Refusal::MissingStory => matches!(result, Err(EditError::StoryNotFound(_))),
        Refusal::MissingShape => matches!(result, Err(EditError::ShapeNotFound(_))),
    };
    assert!(named, "{method}: {result:?}");
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
    refused(
        "insert_text",
        &expected,
        session.insert_text(&ctx, story, 1, "X", &style),
    );
    refused(
        "delete_text",
        &expected,
        session.delete_text(&ctx, story, 0, 1),
    );
    refused(
        "delete_text_joining_paragraphs",
        &expected,
        session.delete_text_joining_paragraphs(&ctx, story, 1, 2),
    );
    refused(
        "format_text",
        &expected,
        session.format_text(&ctx, story, 0, 1, &emphasis),
    );
    refused(
        "set_paragraph_alignment",
        &expected,
        session.set_paragraph_alignment(&ctx, story, 0, 1, center),
    );
    refused(
        "insert_paragraph_break",
        &expected,
        session.insert_paragraph_break(&ctx, story, 1),
    );
    refused(
        "delete_paragraph_break",
        &expected,
        session.delete_paragraph_break(&ctx, story, 1),
    );
    refused(
        "set_story_paragraphs adding a paragraph",
        &expected,
        session.set_story_paragraphs(&ctx, story, &grown),
    );
    refused(
        "set_story_paragraphs removing a paragraph",
        &expected,
        session.set_story_paragraphs(&ctx, story, &shrunk),
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
    local.add_undo_barrier();
    remove(&local, &slide, &shape, Removal::Shape);
    local.add_undo_barrier();
    assert!(matches!(
        local.story(&cell),
        Err(EditError::ShapeNotFound(_))
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
    assert_eq!(live_text(&local, &cell), None);
    assert_every_write_refused(&local, &cell, Refusal::MissingStory);
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
        refused(
            "insert_text",
            &Refusal::MissingStory,
            peer.insert_text(&ctx, &story, 1, "X", &TextStyle::default()),
        );
    }
}
