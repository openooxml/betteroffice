use pptx_edit::{DeckSession, EditCtx, EditError};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn context() -> EditCtx {
    EditCtx::local("test")
}

#[test]
fn slide_scope_matches_the_full_snapshot_position() {
    let session = DeckSession::open(FIXTURE, 7).unwrap();
    let deck = session.snapshot().unwrap();
    for (i, slide) in deck.slides.iter().enumerate() {
        let scope = session.slide_scope(i).unwrap();
        assert_eq!(scope.index, i);
        assert_eq!(&scope.slide, slide);
        assert_eq!(scope.width_emu, deck.width_emu);
        assert_eq!(scope.height_emu, deck.height_emu);
    }
    assert!(matches!(
        session.slide_scope(deck.slides.len()),
        Err(EditError::OutOfBounds { .. })
    ));
}

#[test]
fn slide_scope_tracks_insert_and_reorder() {
    let session = DeckSession::open(FIXTURE, 7).unwrap();
    let inserted = session.insert_slide(&context(), 0, None).unwrap();
    let deck = session.snapshot().unwrap();
    let scope = session.slide_scope(0).unwrap();
    assert_eq!(scope.slide, deck.slides[0]);
    assert_eq!(scope.slide.id, inserted.slide_id);

    let moved_id = deck.slides[1].id.clone();
    session.move_slide(&context(), &moved_id, 0).unwrap();
    let deck = session.snapshot().unwrap();
    assert_eq!(deck.slides[0].id, moved_id);
    let scope = session.slide_scope(0).unwrap();
    assert_eq!(scope.slide, deck.slides[0]);
    assert_eq!(scope.slide.id, moved_id);
}
