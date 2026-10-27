use betteroffice_pptx::{
    ColorValue, EditCtx, Error, GradientFill, ModelGradientStop, Presentation, ShapeFill, ShapeNode,
};

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
    let split = presentation.story(&story_id).unwrap();
    presentation
        .delete_paragraph_break(&context, &story_id, 1)
        .unwrap();
    let joined = presentation.story(&story_id).unwrap();
    assert_eq!(joined.plain_text(), before.plain_text());
    assert_eq!(joined.length, before.length);
    assert_eq!(joined.paragraphs.len(), before.paragraphs.len());
    assert_eq!(joined.paragraphs[0].id, split.paragraphs[0].id);
    assert_eq!(
        joined.paragraphs[0].alignment,
        split.paragraphs[0].alignment
    );
    assert_eq!(joined.paragraphs[0].runs, before.paragraphs[0].runs);
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
fn edit_batches_run_through_the_facade() {
    use betteroffice_pptx::{
        EditFailureCode, EditHistory, EditRequest, EditSource, EditStep, ReadRequest, TargetEdge,
        TextRange, TextTarget,
    };

    let presentation = Presentation::open(FIXTURE).unwrap();
    let read = presentation
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert_eq!(read.version, presentation.version());
    let story = &read.stories[0];
    let request = EditRequest {
        expect_version: read.version.clone(),
        source: EditSource::Agent,
        history: EditHistory::Separate,
        steps: vec![EditStep::InsertText {
            target: TextTarget::Range(TextRange {
                slide_id: story.slide_id.clone(),
                shape_id: story.shape_id.clone(),
                story_id: story.story_id.clone(),
                start: 0,
                end: 0,
            }),
            at: TargetEdge::Start,
            text: "Draft: ".into(),
            expect: None,
        }],
    };
    assert!(
        presentation
            .validate_edits(&request)
            .unwrap()
            .unwrap()
            .would_apply
    );
    let applied = presentation.apply_edits(&request).unwrap().unwrap();
    assert!(applied.applied);
    assert!(presentation.can_undo());
    let stale = presentation.apply_edits(&request).unwrap().unwrap_err();
    assert_eq!(stale.failure.code, EditFailureCode::StaleVersion);
    let reopened = Presentation::open(&presentation.save().unwrap()).unwrap();
    assert!(
        reopened
            .story(&story.story_id)
            .unwrap()
            .plain_text()
            .starts_with("Draft: ")
    );
}

#[test]
fn structured_exports_run_through_the_facade() {
    use betteroffice_pptx::{
        AnchorScope, ExportError, ExportFailureCode, PptxExportOptions, PptxMarkdownOptions,
        export_pptx_markdown, export_pptx_structured, render_pptx_markdown,
    };

    let presentation = Presentation::open(FIXTURE).unwrap();
    let options = PptxExportOptions::default();
    let read = presentation.export_structured(&options).unwrap().unwrap();
    assert_eq!(read.version, presentation.version());
    assert_eq!(read.content.anchor_scope, AnchorScope::Session);
    assert_eq!(read.content.slides.len(), 3);
    let markdown = presentation.export_markdown(&options).unwrap().unwrap();
    assert_eq!(markdown.version, read.version);

    let snapshot = export_pptx_structured(FIXTURE, &options).unwrap();
    assert_eq!(snapshot.anchor_scope, AnchorScope::Snapshot);
    let rendered = render_pptx_markdown(&snapshot, &PptxMarkdownOptions::default()).unwrap();
    assert_eq!(rendered, export_pptx_markdown(FIXTURE, &options).unwrap());
    assert_eq!(rendered.markdown, markdown.content.markdown);

    let tiny = PptxExportOptions {
        max_bytes: Some(1),
        ..PptxExportOptions::default()
    };
    let refusal = presentation.export_structured(&tiny).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, ExportFailureCode::InvalidOptions);
    assert!(matches!(
        export_pptx_structured(FIXTURE, &tiny),
        Err(Error::Export(ExportError::Refused(failure))) if failure.code == ExportFailureCode::InvalidOptions
    ));
    assert!(matches!(
        export_pptx_structured(b"not a deck", &options),
        Err(Error::Export(ExportError::Parse(_)))
    ));
}

#[test]
fn sets_and_clears_a_slide_background() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let slide_id = presentation.snapshot().unwrap().slides[0].id.clone();
    let fill = ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(ColorValue {
            rgb: Some("0F1720".to_owned()),
            ..ColorValue::default()
        }),
        gradient: None,
    };

    let receipt = presentation
        .set_slide_background(&EditCtx::local("facade-test"), &slide_id, Some(&fill))
        .unwrap();
    assert_eq!(receipt.after, Some(fill.clone()));

    let saved = presentation.save().unwrap();
    let reopened = Presentation::open(&saved).unwrap();
    assert_eq!(
        reopened.snapshot().unwrap().slides[0].background,
        Some(fill)
    );

    let receipt = reopened
        .set_slide_background(&EditCtx::local("facade-test"), &slide_id, None)
        .unwrap();
    assert_eq!(receipt.after, None);
    let saved = reopened.save().unwrap();
    assert_eq!(
        Presentation::open(&saved)
            .unwrap()
            .snapshot()
            .unwrap()
            .slides[0]
            .background,
        None
    );
}

#[test]
fn sets_and_clears_a_gradient_slide_background() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let slide_id = presentation.snapshot().unwrap().slides[0].id.clone();
    let fill = ShapeFill {
        fill_type: "gradient".to_owned(),
        color: None,
        gradient: Some(GradientFill {
            gradient_type: "linear".to_owned(),
            angle: Some(45.0),
            stops: vec![
                ModelGradientStop {
                    position: 0.0,
                    color: ColorValue {
                        rgb: Some("0F1720".to_owned()),
                        ..ColorValue::default()
                    },
                },
                ModelGradientStop {
                    position: 100_000.0,
                    color: ColorValue {
                        rgb: Some("0F766E".to_owned()),
                        ..ColorValue::default()
                    },
                },
            ],
        }),
    };

    presentation
        .set_slide_background(&EditCtx::local("facade-test"), &slide_id, Some(&fill))
        .unwrap();
    let saved = presentation.save().unwrap();
    assert_eq!(
        Presentation::open(&saved)
            .unwrap()
            .snapshot()
            .unwrap()
            .slides[0]
            .background,
        Some(fill)
    );
}

#[test]
fn reports_native_parse_errors() {
    assert!(matches!(
        Presentation::open(b"not a presentation"),
        Err(Error::Parse(_))
    ));
}
