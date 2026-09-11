use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/run-spacing.pptx");
const STORY: &str = "story:slide:0:256:shape:0:0";

#[test]
fn tracking_formatting_round_trips_and_rejects_invalid_values() {
    let session = DeckSession::open(DECK, 32506).unwrap();
    let context = EditCtx::local("test");
    let end = session.story(STORY).unwrap().length - 1;
    for spacing in [3.0, -1.0, 0.0] {
        session
            .format_text(
                &context,
                STORY,
                0,
                end,
                &TextStylePatch {
                    spacing_pt: Some(spacing),
                    ..Default::default()
                },
            )
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 32507).unwrap();
        assert_eq!(
            reopened.story(STORY).unwrap(),
            session.story(STORY).unwrap()
        );
        assert_eq!(
            reopened.story(STORY).unwrap().paragraphs[0].runs[0]
                .style
                .spacing_pt,
            Some(spacing)
        );
    }
    let before = session.encode_state_as_update_v1();
    for spacing in [
        f64::NAN,
        f64::INFINITY,
        f64::NEG_INFINITY,
        -4000.01,
        4000.01,
    ] {
        assert!(
            session
                .format_text(
                    &context,
                    STORY,
                    0,
                    1,
                    &TextStylePatch {
                        spacing_pt: Some(spacing),
                        ..Default::default()
                    }
                )
                .is_err()
        );
        assert!(
            session
                .insert_text(
                    &context,
                    STORY,
                    0,
                    "X",
                    &TextStyle {
                        spacing_pt: Some(spacing),
                        ..Default::default()
                    }
                )
                .is_err()
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}
