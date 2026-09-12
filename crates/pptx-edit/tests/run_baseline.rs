use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/text-baseline-script.pptx");
const STORY: &str = "story:slide:0:256:shape:3:0";

#[test]
fn baseline_formatting_round_trips_and_rejects_invalid_values() {
    let session = DeckSession::open(DECK, 33105).unwrap();
    let context = EditCtx::local("test");
    for baseline in [150.0, -150.0, 0.0] {
        session
            .format_text(
                &context,
                STORY,
                6,
                7,
                &TextStylePatch {
                    baseline_pct: Some(baseline),
                    ..Default::default()
                },
            )
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 33106).unwrap();
        assert_eq!(
            reopened.story(STORY).unwrap(),
            session.story(STORY).unwrap()
        );
        assert_eq!(
            reopened.story(STORY).unwrap().paragraphs[0].runs[1]
                .style
                .baseline_pct,
            Some(baseline)
        );
    }
    let before = session.encode_state_as_update_v1();
    for baseline in [
        f64::NAN,
        f64::INFINITY,
        f64::NEG_INFINITY,
        2147483.648,
        -2147483.649,
    ] {
        assert!(
            session
                .format_text(
                    &context,
                    STORY,
                    6,
                    7,
                    &TextStylePatch {
                        baseline_pct: Some(baseline),
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
                        baseline_pct: Some(baseline),
                        ..Default::default()
                    }
                )
                .is_err()
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn absent_baseline_preserves_legacy_style_json() {
    let json = r#"{"bold":null,"italic":null,"fontSizePt":null,"color":null,"fontFamily":null,"underline":null}"#;
    assert_eq!(
        serde_json::to_string(&serde_json::from_str::<TextStyle>(json).unwrap()).unwrap(),
        json
    );
    assert_eq!(
        serde_json::to_string(&serde_json::from_str::<TextStylePatch>(json).unwrap()).unwrap(),
        json
    );
}
