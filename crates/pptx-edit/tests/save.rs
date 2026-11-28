use std::collections::BTreeMap;

use pptx_edit::{
    CommentFlavor, DeckSession, DeckSnapshot, EditCtx, EditError, PresetShapeDraft, ShapeDraft,
    ShapeKind, ShapeRect, ShapeSnapshot, ShapeStroke, TextStyle, TextStylePatch,
};
use yrs::{Map, MapRef, ReadTxn, Transact};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const GRADIENT_OUTLINE_FIXTURE: &[u8] =
    include_bytes!("../../pptx-parse/tests/fixtures/gradient-outline.pptx");
const PROPOSAL_SLIDE_PROPS_FIXTURE: &[u8] = include_bytes!("fixtures/proposal-slide-props.pptx");
const NUMBERED_FIXTURE: &[u8] =
    include_bytes!("../../pptx-parse/tests/fixtures/slide-number-fields.pptx");

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
fn restored_connector_with_omitted_guide_saves_part_identical() {
    let mut entries = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let (_, slide) = entries
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    *slide = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="2" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>
        <a:prstGeom prst="bentConnector4"><a:avLst><a:gd name="adj1" fmla="*/ 50000 1 2"/>
        </a:avLst></a:prstGeom></p:spPr></p:cxnSp></p:spTree></p:cSld></p:sld>"#
        .to_vec();
    let source = ooxml_opc::rezip_parts(&entries).unwrap();
    let session = DeckSession::open(&source, 7).unwrap();
    let snapshot = session.snapshot().unwrap();
    let connector = &snapshot.slides[0].shapes[0];
    assert_eq!(connector.adjust_values.get("adj1"), Some(&0.25));
    {
        let mut txn = session.yrs_doc().transact_mut();
        let shape = txn
            .get_map("pptx:shapes")
            .unwrap()
            .get(&txn, &connector.id)
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        shape.insert(&mut txn, "adjustValuesJson", r#"{"adj1":0.25}"#);
    }
    let restored =
        DeckSession::open_from_update_with_source(&session.encode_state_as_update_v1(), &source, 8)
            .unwrap();
    assert_eq!(parts(&restored.save().unwrap()), parts(&source));
}

#[test]
fn an_unedited_deck_numbered_from_ten_saves_part_identical() {
    let session = DeckSession::open(NUMBERED_FIXTURE, 7).unwrap();
    assert_eq!(session.package().presentation.first_slide_num, 10);
    assert_eq!(parts(&session.save().unwrap()), parts(NUMBERED_FIXTURE));
    let restored = DeckSession::open_from_update_with_source(
        &session.encode_state_as_update_v1(),
        NUMBERED_FIXTURE,
        8,
    )
    .unwrap();
    assert_eq!(parts(&restored.save().unwrap()), parts(NUMBERED_FIXTURE));
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
fn a_paragraph_alignment_survives_reopen() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    session
        .set_paragraph_alignment(&context(), &story_id, 0, 0, Some("ctr"))
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
    assert_eq!(story.paragraphs[0].alignment.as_deref(), Some("ctr"));
}

#[test]
fn an_alignment_with_the_caret_at_the_story_end_reaches_the_last_paragraph() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    let length = session.story(&story_id).unwrap().length;
    session
        .set_paragraph_alignment(&context(), &story_id, length, length, Some("r"))
        .unwrap();

    let story = session.story(&story_id).unwrap();
    assert_eq!(
        story.paragraphs.last().unwrap().alignment.as_deref(),
        Some("r")
    );
}

#[test]
fn an_unknown_paragraph_alignment_is_rejected() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    let error = session
        .set_paragraph_alignment(&context(), &story_id, 0, 0, Some("middle"))
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidText(message) if message.contains("middle")));
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
fn a_width_only_stroke_edit_keeps_a_gradient_outline() {
    let session = DeckSession::open(GRADIENT_OUTLINE_FIXTURE, 21).unwrap();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| {
                    shape
                        .outline
                        .as_ref()
                        .is_some_and(|outline| outline.gradient.is_some())
                })
                .map(|shape| (slide.id.clone(), shape.clone()))
        })
        .expect("fixture has a gradient-stroked shape");
    session
        .set_shape_stroke(
            &context(),
            &slide_id,
            &shape.id,
            &ShapeStroke {
                color: None,
                width_pt: Some(3.0),
            },
        )
        .unwrap();

    let reopened = reopen(&session);
    let restyled = reopened
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|candidate| candidate.source_id == shape.source_id)
        .cloned()
        .unwrap();
    let outline = restyled.outline.as_ref().unwrap();

    assert_eq!(outline.width, Some(38_100.0));
    assert_eq!(outline.gradient, shape.outline.as_ref().unwrap().gradient);
    assert!(outline.color.is_none());
}

#[test]
fn a_colour_stroke_edit_replaces_a_gradient_outline() {
    let session = DeckSession::open(GRADIENT_OUTLINE_FIXTURE, 22).unwrap();
    let snapshot = session.snapshot().unwrap();
    let (slide_id, shape) = snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| {
                    shape
                        .outline
                        .as_ref()
                        .is_some_and(|outline| outline.gradient.is_some())
                })
                .map(|shape| (slide.id.clone(), shape.clone()))
        })
        .expect("fixture has a gradient-stroked shape");
    session
        .set_shape_stroke(
            &context(),
            &slide_id,
            &shape.id,
            &ShapeStroke {
                color: Some("#00FF00".to_owned()),
                width_pt: None,
            },
        )
        .unwrap();

    assert!(
        session.snapshot().unwrap().slides[0].shapes[0]
            .outline
            .as_ref()
            .unwrap()
            .gradient
            .is_none()
    );

    let reopened = reopen(&session);
    let restyled = reopened
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find(|candidate| candidate.source_id == shape.source_id)
        .cloned()
        .unwrap();

    assert_eq!(restyled.resolved_outline_color.as_deref(), Some("#00FF00"));
    assert!(restyled.outline.as_ref().unwrap().gradient.is_none());
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
fn deleting_a_paragraph_break_merges_and_survives_reopen() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    let before = session.story(&story_id).unwrap();
    session
        .insert_paragraph_break(&context(), &story_id, 1)
        .unwrap();
    let split = session.story(&story_id).unwrap();
    session.add_undo_barrier();
    let receipt = session
        .delete_paragraph_break(&context(), &story_id, 1)
        .unwrap();
    assert_eq!(
        (receipt.start, receipt.end, receipt.text.as_str()),
        (1, 2, "\n")
    );
    let joined = session.story(&story_id).unwrap();
    assert_eq!(joined.plain_text(), before.plain_text());
    assert_eq!(joined.length, before.length);
    assert_eq!(joined.paragraphs[0].id, split.paragraphs[0].id);
    assert_eq!(
        joined.paragraphs[0].alignment,
        split.paragraphs[0].alignment
    );
    assert_eq!(joined.paragraphs[0].level, split.paragraphs[0].level);
    assert_eq!(
        joined.paragraphs[0].bullet_json,
        split.paragraphs[0].bullet_json
    );
    assert_eq!(joined.paragraphs[0].runs, before.paragraphs[0].runs);
    assert!(session.undo());
    assert_eq!(session.story(&story_id).unwrap(), split);
    assert!(session.redo());
    assert_eq!(session.story(&story_id).unwrap(), joined);

    let reopened = reopen(&session);
    let restored = reopened.story(&story_id).unwrap();
    assert_eq!(restored.plain_text(), joined.plain_text());
    assert_eq!(
        restored.paragraphs[0].alignment,
        joined.paragraphs[0].alignment
    );
    assert_eq!(restored.paragraphs[0].level, joined.paragraphs[0].level);
    assert_eq!(
        restored.paragraphs[0].bullet_json,
        joined.paragraphs[0].bullet_json
    );
    assert_eq!(restored.paragraphs[0].runs, joined.paragraphs[0].runs);
}

#[test]
fn deleting_a_paragraph_break_rejects_text_and_the_final_pilcrow() {
    let session = open_fixture();
    let (_, _, story_id) = first_story(&session.snapshot().unwrap());
    let story = session.story(&story_id).unwrap();
    let error = session
        .delete_paragraph_break(&context(), &story_id, 0)
        .unwrap_err();
    assert!(matches!(
        error,
        EditError::ParagraphBoundary { start: 0, end: 1 }
    ));
    let final_pilcrow = story.length - 1;
    let error = session
        .delete_paragraph_break(&context(), &story_id, final_pilcrow)
        .unwrap_err();
    assert!(matches!(
        error,
        EditError::OutOfBounds { index, length }
            if index == final_pilcrow && length == final_pilcrow
    ));
    assert_eq!(session.story(&story_id).unwrap(), story);
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

#[test]
fn a_comment_added_to_the_demo_deck_survives_a_reopen() {
    let session = open_fixture();
    let slide = session.snapshot().unwrap().slides[1].id.clone();
    session
        .add_comment(
            &context(),
            &slide,
            "Ada Lovelace",
            "AL",
            "Tighten this claim.",
            "2026-09-01T10:00:00.000",
            1_828_800,
            914_400,
        )
        .unwrap();

    let saved = session.save().unwrap();
    let written = parts(&saved);
    assert!(written.contains_key("ppt/comments/comment2.xml"));
    assert!(written.contains_key("ppt/commentAuthors.xml"));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.comments.len(), 1);
    assert_eq!(snapshot.comments[0].text, "Tighten this claim.");
    assert_eq!(snapshot.comments[0].author, "Ada Lovelace");
    assert_eq!(snapshot.comments[0].slide_id, snapshot.slides[1].id);
    assert_eq!(snapshot.comments[0].x_emu, 1_828_800);
    assert_eq!(snapshot.comments[0].y_emu, 914_400);
}

#[test]
fn a_deck_with_no_comment_edits_keeps_every_part_untouched() {
    let source = include_bytes!("fixtures/modern-comments.pptx");
    let session = DeckSession::open(source, 719).unwrap();
    assert_eq!(parts(&session.save().unwrap()), parts(source));
}

#[test]
fn two_authors_on_one_slide_each_get_their_own_index_run() {
    let session = open_fixture();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    for (author, initials, text, created) in [
        (
            "Ada Lovelace",
            "AL",
            "First from Ada.",
            "2026-09-01T10:00:00.000",
        ),
        (
            "Grace Hopper",
            "GH",
            "First from Grace.",
            "2026-09-01T10:01:00.000",
        ),
        (
            "Ada Lovelace",
            "AL",
            "Second from Ada.",
            "2026-09-01T10:02:00.000",
        ),
    ] {
        session
            .add_comment(&context(), &slide, author, initials, text, created, 0, 0)
            .unwrap();
    }

    let saved = parts(&session.save().unwrap());
    let authors = String::from_utf8(saved["ppt/commentAuthors.xml"].clone()).unwrap();
    assert!(authors.contains(r#"id="0" initials="AL" lastIdx="2" name="Ada Lovelace""#));
    assert!(authors.contains(r#"id="1" initials="GH" lastIdx="1" name="Grace Hopper""#));

    let comments = String::from_utf8(saved["ppt/comments/comment1.xml"].clone()).unwrap();
    assert_eq!(comments.matches(r#"authorId="0""#).count(), 2);
    assert_eq!(comments.matches(r#"authorId="1""#).count(), 1);
    assert!(
        comments.contains(r#"idx="2""#),
        "Ada's second comment takes index 2"
    );
}

#[test]
fn an_empty_comment_is_rejected_at_the_edit() {
    let session = open_fixture();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    assert!(matches!(
        session.add_comment(
            &context(),
            &slide,
            "Ada",
            "AL",
            "",
            "2026-09-01T10:00:00.000",
            0,
            0
        ),
        Err(EditError::InvalidComment(_))
    ));
}

#[test]
fn a_comment_on_an_unknown_slide_is_rejected() {
    let session = open_fixture();
    assert!(matches!(
        session.add_comment(
            &context(),
            "slide:404:404",
            "Ada",
            "AL",
            "Nowhere.",
            "2026-09-01T10:00:00.000",
            0,
            0
        ),
        Err(EditError::SlideNotFound(_))
    ));
}

#[test]
fn removing_a_thread_root_removes_its_replies() {
    let session = open_fixture();
    session
        .set_comment_flavor(&context(), CommentFlavor::Modern)
        .unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let root = session
        .add_comment(
            &context(),
            &slide,
            "Ada",
            "AL",
            "Root.",
            "2026-09-01T10:00:00.000",
            0,
            0,
        )
        .unwrap();
    session
        .reply_to_comment(
            &context(),
            &root.comment_id,
            "Grace",
            "GH",
            "Reply.",
            "2026-09-01T10:01:00.000",
        )
        .unwrap();
    assert_eq!(session.comments().unwrap().len(), 2);

    session
        .remove_comment(&context(), &root.comment_id)
        .unwrap();
    assert!(session.comments().unwrap().is_empty());
}

#[test]
fn an_unknown_comment_is_reported_not_ignored() {
    let session = open_fixture();
    assert!(matches!(
        session.remove_comment(&context(), "comment:0:0"),
        Err(EditError::CommentNotFound(_))
    ));
}

#[test]
fn a_comment_on_a_newly_inserted_slide_survives_the_save() {
    let session = open_fixture();
    let slide = session.insert_slide(&context(), 1, None).unwrap();
    session
        .add_comment(
            &context(),
            &slide.slide_id,
            "Ada Lovelace",
            "AL",
            "Fresh slide, fresh comment.",
            "2026-09-02T10:00:00.000",
            914_400,
            457_200,
        )
        .unwrap();

    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 13).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides.len(), 4);
    assert_eq!(
        snapshot.comments.len(),
        1,
        "a comment anchored to a minted slide must reach the saved deck"
    );
    assert_eq!(snapshot.comments[0].text, "Fresh slide, fresh comment.");
    assert_eq!(snapshot.comments[0].slide_id, snapshot.slides[1].id);
}

#[test]
fn width_edits_preserve_authored_gradient_geometry() {
    let mut parts = ooxml_opc::unzip_parts(GRADIENT_OUTLINE_FIXTURE).unwrap();
    let (_, slide) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    *slide = String::from_utf8(slide.clone())
        .unwrap()
        .replace(
            "<a:gradFill>",
            "<a:gradFill flip=\"xy\" rotWithShape=\"0\">",
        )
        .replace("scaled=\"1\"", "scaled=\"0\"")
        .replace(
            "</a:gradFill>",
            "<a:tileRect l=\"10000\" t=\"20000\" r=\"30000\" b=\"40000\"/></a:gradFill>",
        )
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let session = DeckSession::open(&source, 32208).unwrap();
    let snapshot = session.snapshot().unwrap();
    session
        .set_shape_stroke(
            &EditCtx::local("test"),
            &snapshot.slides[0].id,
            &snapshot.slides[0].shapes[0].id,
            &ShapeStroke {
                color: None,
                width_pt: Some(3.0),
            },
        )
        .unwrap();
    let saved = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    let xml = String::from_utf8(
        saved
            .iter()
            .find(|(path, _)| path == "ppt/slides/slide1.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert_eq!(xml.matches("flip=\"xy\"").count(), 2);
    assert_eq!(xml.matches("scaled=\"0\"").count(), 2);
    assert_eq!(xml.matches("<a:tileRect").count(), 2);
    assert_eq!(xml.matches("rotWithShape=\"0\"").count(), 2);
}

#[test]
fn instantiated_rotated_flipped_placeholder_round_trips_on_save() {
    let mut archive_parts = ooxml_opc::unzip_parts(PROPOSAL_SLIDE_PROPS_FIXTURE).unwrap();
    let (_, layout_xml) = archive_parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slideLayouts/slideLayout1.xml")
        .unwrap();
    let patched = String::from_utf8(layout_xml.clone()).unwrap().replace(
        "<p:spPr/>",
        "<p:spPr><a:xfrm rot=\"1800000\" flipH=\"1\" flipV=\"1\"><a:off x=\"457200\" y=\"274638\"/><a:ext cx=\"8229600\" cy=\"1143000\"/></a:xfrm></p:spPr>",
    );
    assert!(patched.contains("rot=\"1800000\""));
    *layout_xml = patched.into_bytes();
    let source = ooxml_opc::rezip_parts(&archive_parts).unwrap();

    let session = DeckSession::open(&source, 32209).unwrap();
    session
        .insert_slide_with_placeholders(&context(), 1, Some("ppt/slideLayouts/slideLayout1.xml"))
        .unwrap();
    let live = session.snapshot().unwrap().slides[1].shapes[0].clone();
    assert!((live.rotation_deg - 30.0).abs() < 1e-6);
    assert!(live.flip_h);
    assert!(live.flip_v);

    let reopened = reopen(&session);
    let saved_shape = &reopened.snapshot().unwrap().slides[1].shapes[0];
    let (rotation, flip_h, flip_v) = saved_shape.inherited.map_or(
        (
            saved_shape.rotation_deg,
            saved_shape.flip_h,
            saved_shape.flip_v,
        ),
        |transform| (transform.rotation_deg, transform.flip_h, transform.flip_v),
    );
    assert!(
        (rotation - 30.0).abs() < 0.01,
        "expected rotation 30.0 after reopen, got {rotation}"
    );
    assert!(flip_h, "expected flip_h preserved after reopen");
    assert!(flip_v, "expected flip_v preserved after reopen");
    assert_eq!(
        saved_shape
            .placeholder
            .as_ref()
            .and_then(|ph| ph.placeholder_type.as_deref()),
        Some("title")
    );
}
