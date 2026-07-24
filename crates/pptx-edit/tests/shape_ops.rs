use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, PresetShapeDraft, ShapeRect, ShapeStroke};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

#[test]
fn shape_operations_round_trip_through_history_and_updates() {
    let session = DeckSession::open(FIXTURE, 701).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let before_add = session.snapshot().unwrap();
    let receipt = session
        .add_shape(
            &context,
            &slide_id,
            &PresetShapeDraft {
                name: "Rounded rectangle".to_owned(),
                geometry: "roundRect".to_owned(),
                rect: ShapeRect {
                    x: 900_000,
                    y: 1_100_000,
                    width: 3_200_000,
                    height: 1_400_000,
                },
                fill: Some("#D9EAF7".to_owned()),
            },
        )
        .unwrap();
    assert_history_round_trip(&session, before_add);

    let before_fill = session.snapshot().unwrap();
    let fill = session
        .set_shape_fill(&context, &slide_id, &receipt.shape_id, Some("#3367D6"))
        .unwrap();
    assert_eq!(fill.after.as_deref(), Some("#3367D6"));
    assert_history_round_trip(&session, before_fill);

    let before_stroke = session.snapshot().unwrap();
    let stroke = session
        .set_shape_stroke(
            &context,
            &slide_id,
            &receipt.shape_id,
            &ShapeStroke {
                color: Some("#EA4335".to_owned()),
                width_pt: Some(3.0),
            },
        )
        .unwrap();
    assert_eq!(stroke.after.unwrap().width_pt, Some(3.0));
    assert_history_round_trip(&session, before_stroke);

    let before_adjust = session.snapshot().unwrap();
    let adjust = session
        .set_shape_adjust(
            &context,
            &slide_id,
            &receipt.shape_id,
            &BTreeMap::from([("adj".to_owned(), 0.8)]),
        )
        .unwrap();
    assert_eq!(adjust.after.get("adj"), Some(&0.5));
    assert_history_round_trip(&session, before_adjust);

    let before_no_fill = session.snapshot().unwrap();
    session
        .set_shape_fill(&context, &slide_id, &receipt.shape_id, None)
        .unwrap();
    assert_history_round_trip(&session, before_no_fill);

    let before_no_line = session.snapshot().unwrap();
    session
        .set_shape_stroke(
            &context,
            &slide_id,
            &receipt.shape_id,
            &ShapeStroke::default(),
        )
        .unwrap();
    assert_history_round_trip(&session, before_no_line);

    let update = session.encode_state_as_update_v1();
    let replica = DeckSession::open_from_update(&update, 702).unwrap();
    assert_eq!(replica.snapshot().unwrap(), session.snapshot().unwrap());
}

#[test]
fn add_shape_rejects_unknown_geometry_and_invalid_colors() {
    let session = DeckSession::open(FIXTURE, 703).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let draft = PresetShapeDraft {
        name: "Unknown".to_owned(),
        geometry: "not-a-preset".to_owned(),
        rect: ShapeRect {
            x: 0,
            y: 0,
            width: 1_000_000,
            height: 1_000_000,
        },
        fill: None,
    };
    assert!(session.add_shape(&context, &slide_id, &draft).is_err());

    let mut invalid_color = draft;
    invalid_color.geometry = "rect".to_owned();
    invalid_color.fill = Some("#xyz".to_owned());
    assert!(
        session
            .add_shape(&context, &slide_id, &invalid_color)
            .is_err()
    );
}

fn assert_history_round_trip(session: &DeckSession, before: pptx_edit::DeckSnapshot) {
    let after = session.snapshot().unwrap();
    assert_ne!(after, before);
    session.add_undo_barrier();
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.redo());
    assert_eq!(session.snapshot().unwrap(), after);
    session.add_undo_barrier();
}
