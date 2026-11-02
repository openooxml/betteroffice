use betteroffice_pptx::{
    EditCtx, EditError, Error, Presentation, ProposalEdit, ProposalRequest, ShapeRect,
    UndoCaptureMode,
};
use pptx_edit::engine::{Document, Host};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const CLIENT_ID: u64 = 101;

fn presentations() -> (Presentation, Presentation, Document) {
    let original = Presentation::open_collaborative(FIXTURE, CLIENT_ID).unwrap();
    let alias = Presentation::open_collaborative(FIXTURE, CLIENT_ID).unwrap();
    let engine = Document::open_collaborative(
        FIXTURE,
        CLIENT_ID as f64,
        Host {
            entropy: || 1,
            timer_ms: || 0.0,
        },
    )
    .unwrap();
    (original, alias, engine)
}

fn move_shape(original: &Presentation, alias: &Presentation, engine: &Document, x: i64) {
    let snapshot = original.snapshot().unwrap();
    let slide = &snapshot.slides[0].id;
    let shape = &snapshot.slides[0].shapes[0].id;
    for presentation in [original, alias] {
        presentation
            .move_shape(&EditCtx::local("wasm"), slide, shape, x, x)
            .unwrap();
    }
    engine
        .move_shape_json(&format!(
            r#"{{"slideId":"{slide}","shapeId":"{shape}","x":{x},"y":{x}}}"#,
        ))
        .unwrap();
}

#[test]
fn encoding_aliases_match_original_bytes_and_the_ts_engine() {
    let (original, alias, engine) = presentations();
    let remote_state_vector = original.encode_state_vector_v1();
    move_shape(&original, &alias, &engine, 111);
    assert_eq!(alias.snapshot().unwrap(), original.snapshot().unwrap());
    assert_eq!(
        alias.encode_state_vector(),
        original.encode_state_vector_v1()
    );
    assert_eq!(alias.encode_state_vector(), engine.encode_state_vector());
    assert_eq!(
        alias.encode_state_as_update(None).unwrap(),
        original.encode_state_as_update_v1()
    );
    assert_eq!(
        alias.encode_state_as_update(None).unwrap(),
        engine.encode_state_as_update()
    );
    let expected = original.encode_diff_v1(&remote_state_vector).unwrap();
    assert_eq!(alias.encode_diff(&remote_state_vector).unwrap(), expected);
    assert_eq!(
        alias
            .encode_state_as_update(Some(&remote_state_vector))
            .unwrap(),
        expected
    );
    assert_eq!(engine.encode_diff(&remote_state_vector).unwrap(), expected);
    let current = original.encode_state_vector_v1();
    assert_eq!(
        alias.encode_diff(&current).unwrap(),
        original.encode_diff_v1(&current).unwrap()
    );
    assert_eq!(
        alias.encode_state_as_update(Some(&current)).unwrap(),
        original.encode_diff_v1(&current).unwrap()
    );
    let before = alias.snapshot().unwrap();
    for invalid in [b"".as_slice(), b"\xff".as_slice()] {
        let expected = original.encode_diff_v1(invalid).unwrap_err();
        for result in [
            alias.encode_diff(invalid),
            alias.encode_state_as_update(Some(invalid)),
        ] {
            let error = result.unwrap_err();
            assert!(matches!(
                error,
                Error::Edit(EditError::InvalidStateVector(_))
            ));
            assert_eq!(error.to_string(), expected.to_string());
        }
    }
    assert_eq!(alias.snapshot().unwrap(), before);
}

#[test]
fn apply_update_alias_matches_original_and_the_ts_engine() {
    let (original, alias, engine) = presentations();
    let source = Presentation::open_collaborative(FIXTURE, 202).unwrap();
    let snapshot = source.snapshot().unwrap();
    source
        .move_shape(
            &EditCtx::local("wasm"),
            &snapshot.slides[0].id,
            &snapshot.slides[0].shapes[0].id,
            111,
            222,
        )
        .unwrap();
    let update = source
        .encode_diff_v1(&original.encode_state_vector_v1())
        .unwrap();
    let expected = original.apply_update_v1(&update).unwrap();
    assert_eq!(alias.apply_update(&update).unwrap(), expected);
    engine.apply_update_json(&update).unwrap();
    assert_eq!(alias.snapshot().unwrap(), expected);
    assert_eq!(engine.session().snapshot().unwrap(), expected);
    assert_eq!(source.snapshot().unwrap(), expected);
    assert_eq!(alias.apply_update(&update).unwrap(), expected);
    let error = alias.apply_update(b"\xff").unwrap_err();
    assert!(matches!(error, Error::Edit(EditError::InvalidUpdate(_))));
    assert_eq!(
        error.to_string(),
        original.apply_update_v1(b"\xff").unwrap_err().to_string()
    );
    assert_eq!(alias.snapshot().unwrap(), expected);
}

#[test]
fn undo_boundary_alias_matches_original_grouping_and_the_ts_engine() {
    let (original, alias, engine) = presentations();
    let before = original.snapshot().unwrap();
    original.set_undo_capture_mode(UndoCaptureMode::Manual);
    alias.set_undo_capture_mode(UndoCaptureMode::Manual);
    engine.set_undo_capture_mode("manual").unwrap();
    move_shape(&original, &alias, &engine, 111);
    move_shape(&original, &alias, &engine, 222);
    original.add_undo_barrier();
    alias.add_undo_boundary();
    engine.add_undo_boundary();
    move_shape(&original, &alias, &engine, 333);
    for x in [222, before.slides[0].shapes[0].x] {
        assert!(original.undo());
        assert!(alias.undo());
        engine.undo_json().unwrap();
        let actual = alias.snapshot().unwrap();
        assert_eq!(actual.slides[0].shapes[0].x, x);
        assert_eq!(actual, original.snapshot().unwrap());
        assert_eq!(actual, engine.session().snapshot().unwrap());
    }
    assert_eq!(alias.snapshot().unwrap(), before);
    assert!(!alias.can_undo());
    for _ in 0..2 {
        assert!(original.redo());
        assert!(alias.redo());
        engine.redo_json().unwrap();
        assert_eq!(alias.snapshot().unwrap(), original.snapshot().unwrap());
        assert_eq!(
            alias.snapshot().unwrap(),
            engine.session().snapshot().unwrap()
        );
    }
    assert!(!alias.can_redo());
}

#[test]
fn proposal_list_alias_matches_original_and_the_ts_engine() {
    let (original, alias, engine) = presentations();
    assert!(original.proposals().unwrap().is_empty());
    assert_eq!(
        alias.list_proposals().unwrap(),
        original.proposals().unwrap()
    );
    let snapshot = original.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    let request = ProposalRequest {
        agent_id: "alias-test".into(),
        note: Some("Move shape".into()),
        edits: vec![ProposalEdit::SetShapeRect {
            slide_id: slide.id.clone(),
            shape_id: shape.id.clone(),
            rect: ShapeRect {
                x: 111,
                y: 222,
                width: shape.width,
                height: shape.height,
            },
        }],
    };
    let proposal = original.propose(request.clone()).unwrap();
    assert_eq!(alias.propose(request).unwrap(), proposal);
    engine
        .propose_json(&format!(
            r#"{{"agentId":"alias-test","note":"Move shape","edits":[{{"type":"setShapeRect","slideId":"{}","shapeId":"{}","rect":{{"x":111,"y":222,"width":{},"height":{}}}}}]}}"#,
            slide.id, shape.id, shape.width, shape.height,
        ))
        .unwrap();
    assert_eq!(alias.list_proposals().unwrap(), vec![proposal.clone()]);
    assert_eq!(
        alias.list_proposals().unwrap(),
        original.proposals().unwrap()
    );
    assert_eq!(
        alias.list_proposals().unwrap(),
        engine.session().proposals().unwrap()
    );
    move_shape(&original, &alias, &engine, 333);
    let listed = alias.list_proposals().unwrap();
    assert_eq!(listed[0].stale_targets, vec![shape.id.clone()]);
    assert_eq!(listed, original.proposals().unwrap());
    assert_eq!(listed, engine.session().proposals().unwrap());
    assert!(original.reject_proposal(&proposal.id));
    assert!(alias.reject_proposal(&proposal.id));
    assert!(alias.list_proposals().unwrap().is_empty());
    assert_eq!(
        alias.list_proposals().unwrap(),
        original.proposals().unwrap()
    );
}
