use betteroffice_pptx::{EditCtx, Error, Presentation, ShapeNode};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn opens_edits_renders_saves_and_reopens() {
    let mut presentation = Presentation::open(FIXTURE).unwrap();
    assert_eq!(presentation.slides().len(), 3);
    assert_eq!(presentation.layouts().len(), 1);
    assert_eq!(presentation.masters().len(), 1);
    assert!(presentation.slides().iter().any(|slide| {
        slide.shapes.iter().any(|shape| {
            matches!(shape, ShapeNode::Shape(shape) if shape.text.as_ref().is_some_and(|body| {
                body.paragraphs.iter().flat_map(|paragraph| &paragraph.runs).any(|run| run.text.contains("Rust"))
            }))
        })
    }));

    let before = presentation.snapshot().unwrap();
    let slide_id = before.slides[0].id.clone();
    let shape_id = before.slides[0].shapes[0].id.clone();
    let receipt = presentation
        .move_shape(
            &EditCtx::local("facade-test"),
            &slide_id,
            &shape_id,
            1_111_111,
            2_222_222,
        )
        .unwrap();
    assert_eq!((receipt.after.x, receipt.after.y), (1_111_111, 2_222_222));
    let after = presentation.snapshot().unwrap();
    assert_eq!(
        (after.slides[0].shapes[0].x, after.slides[0].shapes[0].y),
        (1_111_111, 2_222_222)
    );

    for bold in [false, true] {
        presentation
            .register_font("Inter", bold, false, FONT)
            .unwrap();
    }
    let rendered = presentation.render_slide(0).unwrap();
    assert_eq!(
        (rendered.display_list.width, rendered.display_list.height),
        (1280.0, 720.0)
    );
    assert!(!rendered.display_list.primitives.is_empty());

    let saved = presentation.save().unwrap();
    let reopened = Presentation::open(&saved).unwrap();
    assert_eq!(reopened.slides().len(), 3);
    assert_eq!(reopened.package().presentation.slides.len(), 3);
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(
        (
            snapshot.slides[0].shapes[0].x,
            snapshot.slides[0].shapes[0].y
        ),
        (1_111_111, 2_222_222)
    );
}

#[test]
fn collaborative_facade_exchanges_typed_updates() {
    let left = Presentation::open_collaborative(FIXTURE, 101).unwrap();
    let right = Presentation::open_collaborative(FIXTURE, 202).unwrap();
    let before = left.snapshot().unwrap();
    let slide_id = before.slides[0].id.clone();
    let shape_id = before.slides[0].shapes[0].id.clone();

    left.move_shape(
        &EditCtx::local("left"),
        &slide_id,
        &shape_id,
        3_333_333,
        4_444_444,
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let snapshot = right.apply_update_v1(&update).unwrap();

    assert_eq!(
        (
            snapshot.slides[0].shapes[0].x,
            snapshot.slides[0].shapes[0].y
        ),
        (3_333_333, 4_444_444)
    );
}

#[test]
fn facade_deletes_a_paragraph_break() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let story_id = presentation
        .snapshot()
        .unwrap()
        .slides
        .into_iter()
        .flat_map(|slide| slide.shapes)
        .flat_map(|shape| shape.text_stories)
        .map(|story| story.id)
        .next()
        .unwrap();
    let before = presentation.story(&story_id).unwrap();
    let context = EditCtx::local("facade-test");
    presentation
        .insert_paragraph_break(&context, &story_id, 1)
        .unwrap();
    presentation
        .delete_paragraph_break(&context, &story_id, 1)
        .unwrap();
    assert_eq!(presentation.story(&story_id).unwrap(), before);
}

#[test]
fn a_facade_seeded_update_accepts_the_source_file() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let update = presentation.encode_state_as_update_v1();
    let replica = pptx_edit::DeckSession::open_from_update_with_source(&update, FIXTURE, 303)
        .expect("the facade fingerprints the file bytes");
    assert!(replica.save().is_ok());
}

#[test]
fn reports_native_parse_errors() {
    assert!(matches!(
        Presentation::open(b"not a presentation"),
        Err(Error::Parse(_))
    ));
}
