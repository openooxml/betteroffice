use std::collections::BTreeMap;

use pptx_edit::{
    DeckSession, DeckSnapshot, EditCtx, PresetShapeDraft, ShapeDraft, ShapeKind, ShapeRect,
    ShapeSnapshot, ShapeStroke, TextStyle, TextStylePatch,
};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn open_fixture() -> DeckSession {
    DeckSession::open(FIXTURE, 7).unwrap()
}

fn reopen(session: &DeckSession) -> DeckSession {
    DeckSession::open(&session.save().unwrap(), 8).unwrap()
}

fn context() -> EditCtx {
    EditCtx::local("test")
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn first_shape_of_kind(snapshot: &DeckSnapshot, kind: ShapeKind) -> (String, ShapeSnapshot) {
    snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| shape.kind == kind && shape.source_id != 0)
                .map(|shape| (slide.id.clone(), shape.clone()))
        })
        .expect("fixture has a shape of the requested kind")
}

fn first_story(snapshot: &DeckSnapshot) -> (String, String, String) {
    snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide.shapes.iter().find_map(|shape| {
                shape
                    .text_stories
                    .first()
                    .map(|story| (slide.id.clone(), shape.id.clone(), story.id.clone()))
            })
        })
        .expect("fixture has a text story")
}

fn story_text(snapshot: &DeckSnapshot, story_id: &str) -> String {
    snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(shape_stories)
        .find(|story| story.id == story_id)
        .map(|story| story.plain_text())
        .expect("story survives the round trip")
}

fn shape_stories(shape: &ShapeSnapshot) -> Vec<&pptx_edit::StorySnapshot> {
    let mut stories: Vec<_> = shape.text_stories.iter().collect();
    for child in &shape.children {
        stories.extend(shape_stories(child));
    }
    stories
}

#[test]
fn an_unedited_deck_saves_part_identical() {
    let session = open_fixture();
    assert_eq!(parts(&session.save().unwrap()), parts(FIXTURE));
}

#[test]
fn a_moved_shape_survives_reopen_and_other_parts_stay_byte_identical() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = first_shape_of_kind(&snapshot, ShapeKind::Shape);
    session
        .move_shape(&context(), &slide_id, &shape.id, 1_111_111, 2_222_222)
        .unwrap();

    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 9).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    let moved = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|candidate| candidate.source_id == shape.source_id)
        .unwrap();
    assert_eq!((moved.x, moved.y), (1_111_111, 2_222_222));
    assert_eq!((moved.width, moved.height), (shape.width, shape.height));

    let source_slide_part = snapshot
        .slides
        .iter()
        .find(|slide| slide.shapes.iter().any(|s| s.source_id == shape.source_id))
        .and_then(|slide| slide.source_part_path.clone())
        .unwrap();
    let before = parts(FIXTURE);
    let after = parts(&saved);
    assert_eq!(before.len(), after.len());
    for (path, bytes) in &before {
        if path != &source_slide_part {
            assert_eq!(after.get(path), Some(bytes), "{path} should be untouched");
        }
    }
    assert_ne!(
        after.get(&source_slide_part),
        Some(&before[&source_slide_part])
    );
}

#[test]
fn inserted_text_survives_reopen() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    let before = session.story(&story_id).unwrap().plain_text();
    session
        .insert_text(&context(), &story_id, 0, "Hello ", &TextStyle::default())
        .unwrap();

    let reopened = reopen(&session);
    let text = story_text(&reopened.snapshot().unwrap(), &story_id);
    assert_eq!(text, format!("Hello {before}"));
}

#[test]
fn a_formatted_range_survives_reopen() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    session
        .format_text(
            &context(),
            &story_id,
            0,
            3,
            &TextStylePatch {
                bold: Some(true),
                color: Some("#FF00AA".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();

    let reopened = reopen(&session);
    let snapshot = reopened.snapshot().unwrap();
    let story = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(shape_stories)
        .find(|story| story.id == story_id)
        .unwrap();
    let first_run = &story.paragraphs[0].runs[0];
    assert_eq!(first_run.style.bold, Some(true));
    assert_eq!(first_run.style.color.as_deref(), Some("#FF00AA"));
}

#[test]
fn fill_and_stroke_survive_reopen() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = first_shape_of_kind(&snapshot, ShapeKind::Shape);
    session
        .set_shape_fill(&context(), &slide_id, &shape.id, Some("#FF0000"))
        .unwrap();
    session
        .set_shape_stroke(
            &context(),
            &slide_id,
            &shape.id,
            &ShapeStroke {
                color: Some("#00FF00".to_owned()),
                width_pt: Some(2.0),
            },
        )
        .unwrap();

    let reopened = reopen(&session);
    let snapshot = reopened.snapshot().unwrap();
    let restyled = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|candidate| candidate.source_id == shape.source_id)
        .unwrap();
    assert_eq!(restyled.resolved_fill_color.as_deref(), Some("#FF0000"));
    assert_eq!(restyled.resolved_outline_color.as_deref(), Some("#00FF00"));
    assert_eq!(restyled.outline.as_ref().unwrap().width, Some(25_400.0));
}

#[test]
fn added_shapes_survive_reopen() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[0].id.clone();
    let shape_count = snapshot.slides[0].shapes.len();
    session
        .add_text_box(
            &context(),
            &slide_id,
            &ShapeDraft {
                name: "Note".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 200_000,
                    width: 3_000_000,
                    height: 500_000,
                },
                text: "Saved note".to_owned(),
                style: TextStyle {
                    bold: Some(true),
                    font_size_pt: Some(20.0),
                    color: Some("#112233".to_owned()),
                    ..TextStyle::default()
                },
            },
        )
        .unwrap();
    let mut adjustments = BTreeMap::new();
    adjustments.insert("adj".to_owned(), 0.3);
    let ellipse = session
        .add_shape(
            &context(),
            &slide_id,
            &PresetShapeDraft {
                name: "Badge".to_owned(),
                geometry: "roundRect".to_owned(),
                rect: ShapeRect {
                    x: 400_000,
                    y: 400_000,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: Some("#00AAFF".to_owned()),
            },
        )
        .unwrap();
    session
        .set_shape_adjust(&context(), &slide_id, &ellipse.shape_id, &adjustments)
        .unwrap();

    let reopened = reopen(&session);
    let snapshot = reopened.snapshot().unwrap();
    let shapes = &snapshot.slides[0].shapes;
    assert_eq!(shapes.len(), shape_count + 2);
    let note = &shapes[shape_count];
    assert_eq!(note.name, "Note");
    assert_eq!((note.x, note.y), (100_000, 200_000));
    assert_eq!(note.text_stories[0].plain_text(), "Saved note");
    let run = &note.text_stories[0].paragraphs[0].runs[0];
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.font_size_pt, Some(20.0));
    assert_eq!(run.style.color.as_deref(), Some("#112233"));
    let badge = &shapes[shape_count + 1];
    assert_eq!(badge.geometry, "roundRect");
    assert_eq!(badge.resolved_fill_color.as_deref(), Some("#00AAFF"));
    assert_eq!(badge.adjust_values.get("adj"), Some(&0.3));
}

#[test]
fn a_removed_shape_stays_removed_after_reopen() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = first_shape_of_kind(&snapshot, ShapeKind::Shape);
    let slide_part = snapshot
        .slides
        .iter()
        .find(|slide| slide.id == slide_id)
        .and_then(|slide| slide.source_part_path.clone())
        .unwrap();
    let shape_count = snapshot
        .slides
        .iter()
        .find(|slide| slide.id == slide_id)
        .unwrap()
        .shapes
        .len();
    session
        .remove_shape(&context(), &slide_id, &shape.id)
        .unwrap();

    let reopened = reopen(&session);
    let snapshot = reopened.snapshot().unwrap();
    let slide = snapshot
        .slides
        .iter()
        .find(|slide| slide.source_part_path.as_deref() == Some(slide_part.as_str()))
        .unwrap();
    assert_eq!(slide.shapes.len(), shape_count - 1);
    assert!(
        slide
            .shapes
            .iter()
            .all(|candidate| candidate.source_id != shape.source_id)
    );
}

#[test]
fn slide_insert_delete_and_move_survive_reopen() {
    let session = open_fixture();
    let baseline: Vec<String> = session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .map(|slide| slide.source_part_path.clone().unwrap())
        .collect();
    assert_eq!(baseline.len(), 3);

    let inserted = session.insert_slide(&context(), 1, None).unwrap();
    session
        .add_text_box(
            &context(),
            &inserted.slide_id,
            &ShapeDraft {
                name: "Fresh".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 2_000_000,
                    height: 1_000_000,
                },
                text: "Inserted slide".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let reopened = reopen(&session);
    let package = reopened.package();
    assert_eq!(package.presentation.slides.len(), 4);
    assert_eq!(
        package.slides[1].layout_part_path.as_deref(),
        package
            .layouts
            .first()
            .map(|layout| layout.part_path.as_str())
    );
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides.len(), 4);
    assert_eq!(
        snapshot.slides[1].shapes[0].text_stories[0].plain_text(),
        "Inserted slide"
    );

    let session = open_fixture();
    let first = session.snapshot().unwrap().slides[0].id.clone();
    session.delete_slide(&context(), &first).unwrap();
    let reopened = reopen(&session);
    let remaining: Vec<String> = reopened
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .map(|slide| slide.source_part_path.clone().unwrap())
        .collect();
    assert_eq!(remaining, baseline[1..]);
    assert!(!parts(&session.save().unwrap()).contains_key(&baseline[0]));

    let session = open_fixture();
    let first = session.snapshot().unwrap().slides[0].id.clone();
    session.move_slide(&context(), &first, 2).unwrap();
    let reopened = reopen(&session);
    let order: Vec<String> = reopened
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .map(|slide| slide.source_part_path.clone().unwrap())
        .collect();
    assert_eq!(
        order,
        [
            baseline[1].clone(),
            baseline[2].clone(),
            baseline[0].clone()
        ]
    );
}

#[test]
fn a_table_cell_edit_survives_reopen() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(shape_stories)
        .map(|story| story.id.clone())
        .find(|id| id.contains(":table:"))
        .expect("fixture has a table story");
    let before = session.story(&story_id).unwrap().plain_text();
    session
        .insert_text(&context(), &story_id, 0, "Cell: ", &TextStyle::default())
        .unwrap();

    let reopened = reopen(&session);
    let text = story_text(&reopened.snapshot().unwrap(), &story_id);
    assert_eq!(text, format!("Cell: {before}"));
}

#[test]
fn a_paragraph_break_survives_reopen() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    let before = session.story(&story_id).unwrap();
    session
        .insert_paragraph_break(&context(), &story_id, 1)
        .unwrap();

    let reopened = reopen(&session);
    let snapshot = reopened.snapshot().unwrap();
    let story = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(shape_stories)
        .find(|story| story.id == story_id)
        .unwrap();
    assert_eq!(story.paragraphs.len(), before.paragraphs.len() + 1);
    assert_eq!(
        story.plain_text().replace('\n', ""),
        before.plain_text().replace('\n', "")
    );
}

#[test]
fn an_edited_save_reaches_a_part_fixed_point() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = first_shape_of_kind(&snapshot, ShapeKind::Shape);
    session
        .move_shape(&context(), &slide_id, &shape.id, 999_999, 888_888)
        .unwrap();
    let (_, _, story_id) = first_story(&snapshot);
    session
        .insert_text(&context(), &story_id, 0, "Twice ", &TextStyle::default())
        .unwrap();
    session.insert_slide(&context(), 3, None).unwrap();

    let first = session.save().unwrap();
    let second = DeckSession::open(&first, 9).unwrap().save().unwrap();
    assert_eq!(parts(&first), parts(&second));
}

#[test]
fn undo_restores_a_part_identical_save() {
    let session = open_fixture();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = first_shape_of_kind(&snapshot, ShapeKind::Shape);
    session
        .move_shape(&context(), &slide_id, &shape.id, 555_555, 666_666)
        .unwrap();
    assert!(session.undo());
    assert_eq!(parts(&session.save().unwrap()), parts(FIXTURE));
}
