use pptx_edit::{DeckSession, EditCtx, ShapeStroke};
use yrs::{Any, Map, Out, ReadTxn, Transact};

const FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/gradient-outline.pptx");
const V10: &[u8] = include_bytes!("fixtures/gradient-outline-main-v10.update.bin");
const V13: &[u8] = include_bytes!("fixtures/gradient-outline-main-v13.update.bin");

#[test]
fn legacy_gradient_outlines_recover_from_source_without_changing_zip_parts() {
    for old in [V10, V13] {
        legacy_gradient_outlines_recover(old);
    }
}

fn legacy_gradient_outlines_recover(old: &[u8]) {
    let pending = DeckSession::open_from_update(old, 32201).unwrap();
    assert!(
        pending.snapshot().unwrap().slides[0].shapes[0]
            .outline
            .as_ref()
            .unwrap()
            .gradient
            .is_none()
    );
    let session = DeckSession::open_from_update_with_source(
        &pending.encode_state_as_update_v1(),
        FIXTURE,
        32202,
    )
    .unwrap();
    assert_eq!(
        session.snapshot().unwrap(),
        DeckSession::open(FIXTURE, 32203)
            .unwrap()
            .snapshot()
            .unwrap()
    );
    assert_eq!(
        ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap(),
        ooxml_opc::unzip_parts(FIXTURE).unwrap()
    );
    let txn = session.yrs_doc().transact();
    assert_eq!(
        txn.get_map("pptx:meta").unwrap().get(&txn, "schemaVersion"),
        Some(Out::Any(Any::Number(17.0)))
    );
    drop(txn);
    let reopened = DeckSession::open_from_update_with_source(
        &session.encode_state_as_update_v1(),
        FIXTURE,
        32204,
    )
    .unwrap();
    assert_eq!(reopened.snapshot().unwrap(), session.snapshot().unwrap());
    assert_eq!(
        reopened.encode_state_vector_v1(),
        session.encode_state_vector_v1()
    );
    let detached =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 32209).unwrap();
    assert_eq!(detached.package().slides, session.package().slides);
    assert_eq!(detached.snapshot().unwrap(), session.snapshot().unwrap());
}

#[test]
fn legacy_gradient_recovery_preserves_colour_removal_and_geometry_edits() {
    for old in [V10, V13] {
        legacy_gradient_recovery_preserves_edits(old);
    }
}

fn legacy_gradient_recovery_preserves_edits(old: &[u8]) {
    for stroke in [
        None,
        Some(ShapeStroke {
            color: Some("#123456".into()),
            width_pt: None,
        }),
        Some(ShapeStroke::default()),
    ] {
        let old = DeckSession::open_from_update(old, 32205).unwrap();
        let original = old.snapshot().unwrap();
        let slide = &original.slides[0];
        let shape = &slide.shapes[0];
        old.move_shape(
            &EditCtx::local("test"),
            &slide.id,
            &shape.id,
            100_000,
            200_000,
        )
        .unwrap();
        if let Some(stroke) = &stroke {
            old.set_shape_stroke(&EditCtx::local("test"), &slide.id, &shape.id, stroke)
                .unwrap();
        }
        let session = DeckSession::open_from_update_with_source(
            &old.encode_state_as_update_v1(),
            FIXTURE,
            32206,
        )
        .unwrap();
        let snapshot = session.snapshot().unwrap();
        let outline = snapshot.slides[0].shapes[0].outline.as_ref().unwrap();
        assert_eq!(outline.gradient.is_some(), stroke.is_none());
        if stroke.is_some() {
            assert_eq!(
                snapshot.slides[0].shapes[0].outline,
                old.snapshot().unwrap().slides[0].shapes[0].outline
            );
        }
        assert_eq!(
            (
                snapshot.slides[0].shapes[0].x,
                snapshot.slides[0].shapes[0].y
            ),
            (100_000, 200_000)
        );
        let saved = DeckSession::open(&session.save().unwrap(), 32207)
            .unwrap()
            .snapshot()
            .unwrap();
        assert_eq!(
            saved.slides[0].shapes[0]
                .outline
                .as_ref()
                .and_then(|outline| outline.gradient.as_ref()),
            outline.gradient.as_ref()
        );
        assert_eq!(
            saved.slides[0].shapes[0].resolved_outline_color,
            snapshot.slides[0].shapes[0].resolved_outline_color
        );
    }
}
