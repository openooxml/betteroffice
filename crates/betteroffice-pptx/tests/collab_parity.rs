use std::cell::RefCell;
use std::rc::Rc;

use betteroffice_pptx::{
    EditCtx, EditError, Error, Presentation, TableEdit, TableInsertPosition, UndoCaptureMode,
    UpdateOrigin,
};
use pptx_edit::engine::{Document, Host};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn host() -> Host {
    Host {
        entropy: || 1,
        timer_ms: || 0.0,
    }
}

fn target(presentation: &Presentation) -> (String, String) {
    let snapshot = presentation.snapshot().unwrap();
    (
        snapshot.slides[0].id.clone(),
        snapshot.slides[0].shapes[0].id.clone(),
    )
}

fn table_source() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    let cell = "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>";
    let row = format!(r#"<a:tr h="400000">{cell}{cell}</a:tr>"#);
    let table = format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10000" name="Collaboration table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="800000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000000"/><a:gridCol w="1000000"/></a:tblGrid>{row}{row}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
    );
    xml.insert_str(xml.find("</p:spTree>").unwrap(), &table);
    *bytes = xml.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn peers_exchange_updates_and_match_the_ts_engine() {
    let left = Presentation::open_collaborative(FIXTURE, 101).unwrap();
    let right = Presentation::open_collaborative(FIXTURE, 202).unwrap();
    let engine = Document::open_collaborative(FIXTURE, 101.0, host()).unwrap();
    let (slide, shape) = target(&left);
    let context = EditCtx::local("wasm");
    left.move_shape(&context, &slide, &shape, 111, 222).unwrap();
    engine
        .move_shape_json(&format!(
            r#"{{"slideId":"{slide}","shapeId":"{shape}","x":111,"y":222}}"#,
        ))
        .unwrap();
    let merged = right
        .apply_update_v1(
            &left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap(),
        )
        .unwrap();
    assert_eq!(merged, left.snapshot().unwrap());
    assert_eq!(merged, engine.session().snapshot().unwrap());
    right
        .move_shape(&context, &slide, &shape, 333, 444)
        .unwrap();
    left.apply_update_v1(
        &right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
}

#[test]
fn joining_an_update_attaches_only_its_matching_source() {
    let source = Presentation::open_collaborative(FIXTURE, 101).unwrap();
    let (slide, shape) = target(&source);
    source
        .move_shape(&EditCtx::local("wasm"), &slide, &shape, 111, 222)
        .unwrap();
    let update = source.encode_state_as_update_v1();
    let different = source.save().unwrap();
    for bytes in [
        Some(FIXTURE),
        Some(different.as_slice()),
        Some(b"foreign".as_slice()),
        None,
    ] {
        let joined = Presentation::open_from_update(&update, 202, bytes).unwrap();
        let engine = Document::open_collaborative_from_update(
            &update,
            202.0,
            bytes.map(<[u8]>::to_vec),
            host(),
        )
        .unwrap();
        assert_eq!(joined.client_id(), 202);
        assert_eq!(
            joined.snapshot().unwrap(),
            engine.session().snapshot().unwrap()
        );
        let actual = joined.save();
        let expected = engine.save_bytes();
        assert_eq!(actual.is_ok(), bytes == Some(FIXTURE));
        assert_eq!(actual.is_ok(), expected.is_ok());
        if let (Ok(actual), Ok(expected)) = (actual, expected) {
            assert_eq!(actual, expected);
            let reopened = Presentation::open(&actual).unwrap();
            let moved = &reopened.snapshot().unwrap().slides[0].shapes[0];
            assert_eq!((moved.x, moved.y), (111, 222));
        }
    }
    assert!(matches!(
        Presentation::open_from_update(&update, 0, None),
        Err(Error::Edit(EditError::InvalidClientId(0)))
    ));
    assert!(matches!(
        Presentation::open_from_update(b"\xff", 202, None),
        Err(Error::Edit(EditError::InvalidUpdate(_)))
    ));
}

#[test]
fn shared_table_structure_refusals_match_the_ts_engine() {
    let bytes = table_source();
    let presentation = Presentation::open_collaborative(&bytes, 101).unwrap();
    let engine = Document::open_collaborative(&bytes, 101.0, host()).unwrap();
    let snapshot = presentation.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Collaboration table")
        .unwrap();
    let context = EditCtx::local("wasm");
    presentation
        .insert_table_row(
            &context,
            &slide.id,
            &shape.id,
            0,
            TableInsertPosition::After,
            &[],
        )
        .unwrap();
    engine
        .session()
        .insert_table_row(
            &context,
            &slide.id,
            &shape.id,
            0,
            TableInsertPosition::After,
            &[],
        )
        .unwrap();
    assert!(presentation.can_undo());
    let edited = presentation.snapshot().unwrap();
    assert!(!presentation.is_collaborative());
    presentation.set_collaborative().unwrap();
    engine.set_collaborative().unwrap();
    assert!(presentation.is_collaborative());
    assert!(engine.session().is_collaborative());
    assert!(!presentation.can_undo());
    assert!(!engine.session().can_undo());
    for edit in [
        TableEdit::InsertRow,
        TableEdit::DeleteRow,
        TableEdit::InsertColumn,
        TableEdit::DeleteColumn,
    ] {
        let actual = presentation.check_table_edit(
            edit,
            &slide.id,
            &shape.id,
            0,
            TableInsertPosition::Before,
            &[],
        );
        let expected = engine.session().check_table_edit(
            edit,
            &slide.id,
            &shape.id,
            0,
            TableInsertPosition::Before,
            &[],
        );
        assert!(matches!(
            actual,
            Err(Error::Edit(EditError::CollaborativeStructureOperation))
        ));
        assert!(matches!(
            expected,
            Err(EditError::CollaborativeStructureOperation)
        ));
    }
    assert!(matches!(
        presentation.insert_table_row(
            &context,
            &slide.id,
            &shape.id,
            0,
            TableInsertPosition::After,
            &[],
        ),
        Err(Error::Edit(EditError::CollaborativeStructureOperation))
    ));
    assert_eq!(presentation.snapshot().unwrap(), edited);
    assert_eq!(
        presentation.snapshot().unwrap(),
        engine.session().snapshot().unwrap()
    );
}

#[test]
fn manual_capture_and_barriers_group_edits_like_the_ts_engine() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let engine = Document::open_collaborative(FIXTURE, 1.0, host()).unwrap();
    assert_eq!(presentation.undo_capture_mode(), UndoCaptureMode::Auto);
    assert_eq!(engine.undo_capture_mode(), "auto");
    presentation.set_undo_capture_mode(UndoCaptureMode::Manual);
    engine.set_undo_capture_mode("manual").unwrap();
    assert_eq!(presentation.undo_capture_mode(), UndoCaptureMode::Manual);
    let before = presentation.snapshot().unwrap();
    let (slide, shape) = target(&presentation);
    let context = EditCtx::local("wasm");
    for x in [111, 222, 333] {
        if x == 333 {
            presentation.add_undo_barrier();
            engine.add_undo_boundary();
        }
        presentation
            .move_shape(&context, &slide, &shape, x, x)
            .unwrap();
        engine
            .move_shape_json(&format!(
                r#"{{"slideId":"{slide}","shapeId":"{shape}","x":{x},"y":{x}}}"#,
            ))
            .unwrap();
    }
    for x in [222, before.slides[0].shapes[0].x] {
        assert!(presentation.undo());
        engine.undo_json().unwrap();
        let snapshot = presentation.snapshot().unwrap();
        assert_eq!(snapshot.slides[0].shapes[0].x, x);
        assert_eq!(snapshot, engine.session().snapshot().unwrap());
    }
    assert!(!presentation.can_undo());
    assert_eq!(presentation.snapshot().unwrap(), before);
    assert!(presentation.redo());
    engine.redo_json().unwrap();
    assert_eq!(
        presentation.snapshot().unwrap(),
        engine.session().snapshot().unwrap()
    );
    presentation.set_undo_capture_mode(UndoCaptureMode::Auto);
    assert_eq!(presentation.undo_capture_mode(), UndoCaptureMode::Auto);
}

#[test]
fn update_observers_report_origins_and_stop_on_drop() {
    let left = Presentation::open_collaborative(FIXTURE, 101).unwrap();
    let right = Presentation::open_collaborative(FIXTURE, 202).unwrap();
    let local = Rc::new(RefCell::new(Vec::new()));
    let remote = Rc::new(RefCell::new(Vec::new()));
    let events = Rc::clone(&local);
    let guard = left
        .observe_update_v1(move |event| events.borrow_mut().push(event))
        .unwrap();
    let events = Rc::clone(&remote);
    let remote_guard = right
        .observe_update_v1(move |event| events.borrow_mut().push(event))
        .unwrap();
    let (slide, shape) = target(&left);
    left.move_shape(&EditCtx::local("wasm"), &slide, &shape, 111, 222)
        .unwrap();
    assert_eq!(local.borrow().len(), 1);
    assert_eq!(local.borrow()[0].origin, UpdateOrigin::Local);
    right.apply_update_v1(&local.borrow()[0].update).unwrap();
    assert_eq!(remote.borrow().len(), 1);
    assert_eq!(remote.borrow()[0].origin, UpdateOrigin::Remote);
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    drop(guard);
    drop(remote_guard);
    left.move_shape(&EditCtx::local("wasm"), &slide, &shape, 333, 444)
        .unwrap();
    right
        .apply_update_v1(
            &left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap(),
        )
        .unwrap();
    assert_eq!(local.borrow().len(), 1);
    assert_eq!(remote.borrow().len(), 1);
}
