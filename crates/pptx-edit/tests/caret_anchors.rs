use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};

fn session_with_text(text: &str) -> (DeckSession, String) {
    let session = DeckSession::open(
        include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx"),
        84092,
    )
    .unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let added = session
        .add_text_box(
            &EditCtx::local("test"),
            &slide,
            &ShapeDraft {
                name: "Anchors".to_owned(),
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
    let story = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == added.shape_id)
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session.add_undo_barrier();
    (session, story)
}

#[test]
fn caret_anchor_follows_undo_and_redo_of_matching_text() {
    let (session, story) = session_with_text("ABC");
    let context = EditCtx::local("test");
    session
        .insert_text(&context, &story, 0, "A", &TextStyle::default())
        .unwrap();
    let after_typed = session.anchor_caret(&story, 1).unwrap();

    assert!(session.undo());
    assert_eq!(session.resolve_caret_anchor(&after_typed), Some(0));
    assert!(session.redo());
    assert_eq!(session.resolve_caret_anchor(&after_typed), Some(1));
}

#[test]
fn caret_anchor_at_story_start_stays_before_later_insertions() {
    let (session, story) = session_with_text("ABC");
    let start = session.anchor_caret(&story, 0).unwrap();
    let end = session.anchor_caret(&story, 3).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "xy",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&start), Some(0));
    assert_eq!(session.resolve_caret_anchor(&end), Some(5));
    assert!(session.anchor_caret(&story, 6).is_err());
}
