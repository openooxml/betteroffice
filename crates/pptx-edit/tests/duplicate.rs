use pptx_edit::{DeckSession, EditCtx, TextParagraphDraft, TextRunDraft, TextStyle};

const SOURCE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

#[test]
fn duplicate_preserves_shapes_and_notes_and_can_outlive_its_source() {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let ctx = EditCtx::local("test");
    let before = deck.snapshot().unwrap();
    let source = &before.slides[0];
    deck.set_slide_notes(&ctx, &source.id, "speaker notes")
        .unwrap();
    let clone = deck.duplicate_slide(&ctx, &source.id, 1).unwrap();
    let snapshot = deck.snapshot().unwrap();
    let copied = &snapshot.slides[1];
    assert_eq!(copied.notes, "speaker notes");
    assert_eq!(copied.shapes.len(), source.shapes.len());
    for (left, right) in source.shapes.iter().zip(&copied.shapes) {
        assert_ne!(left.id, right.id);
        assert_eq!(left.name, right.name);
        assert_eq!(left.source_id, right.source_id);
        assert_eq!(left.kind, right.kind);
    }
    deck.delete_slide(&ctx, &source.id).unwrap();
    let saved = deck.save().unwrap();
    let restored = DeckSession::open(&saved, 8).unwrap().snapshot().unwrap();
    assert_eq!(restored.slides.len(), before.slides.len());
    assert_eq!(restored.slides[0].notes, "speaker notes");
    assert_eq!(restored.slides[0].shapes.len(), copied.shapes.len());
    assert_eq!(deck.snapshot().unwrap().slides[0].id, clone.slide_id);
}

#[test]
fn duplicate_is_undoable_and_refuses_invalid_targets_without_updates() {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let ctx = EditCtx::local("test");
    let baseline = deck.snapshot().unwrap();
    let source = &baseline.slides[0];
    let version = deck.version();
    assert!(deck.duplicate_slide(&ctx, "missing", 0).is_err());
    assert!(deck.duplicate_slide(&ctx, &source.id, 99999).is_err());
    assert_eq!(deck.version(), version);
    deck.duplicate_slide(&ctx, &source.id, 1).unwrap();
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), baseline);
}

#[test]
fn set_story_paragraphs_replaces_multiple_paragraphs_and_preserves_story_identity() {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let ctx = EditCtx::local("test");
    let snapshot = deck.snapshot().unwrap();
    let story = snapshot
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .flat_map(|s| &s.text_stories)
        .next()
        .unwrap();
    let paragraphs = vec![
        TextParagraphDraft {
            alignment: Some("ctr".into()),
            runs: vec![TextRunDraft {
                text: "New title 😀".into(),
                style: TextStyle {
                    bold: Some(true),
                    font_size_pt: Some(24.0),
                    ..TextStyle::default()
                },
            }],
        },
        TextParagraphDraft {
            alignment: None,
            runs: vec![TextRunDraft {
                text: "Second paragraph".into(),
                style: TextStyle::default(),
            }],
        },
    ];
    let receipt = deck
        .set_story_paragraphs(&ctx, &story.id, &paragraphs)
        .unwrap();
    assert_eq!(receipt.id, story.id);
    assert_eq!(receipt.paragraphs.len(), 2);
    assert_eq!(receipt.plain_text(), "New title 😀\nSecond paragraph");
    let restored = DeckSession::open(&deck.save().unwrap(), 9).unwrap();
    let current = restored.snapshot().unwrap();
    let text = current
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .flat_map(|s| &s.text_stories)
        .find(|s| s.plain_text().starts_with("New title"))
        .unwrap();
    assert_eq!(text.paragraphs[0].alignment.as_deref(), Some("ctr"));
    assert_eq!(text.paragraphs[0].runs[0].style.bold, Some(true));
    let version = deck.version();
    let invalid = vec![TextParagraphDraft {
        alignment: None,
        runs: vec![TextRunDraft {
            text: "invalid\nrun".into(),
            style: TextStyle::default(),
        }],
    }];
    assert!(
        deck.set_story_paragraphs(&ctx, &story.id, &invalid)
            .is_err()
    );
    assert_eq!(deck.version(), version);
}
