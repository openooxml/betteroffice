use std::sync::{Arc, Mutex};

use pptx_edit::{
    DeckSession, EditCtx, EditError, ShapeDraft, ShapeRect, TextStyle, TextStylePatch, UpdateOrigin,
};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn first_text_story(session: &DeckSession) -> String {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find_map(|shape| shape.text_stories.first())
        .unwrap()
        .id
        .clone()
}

fn first_multi_paragraph_story(session: &DeckSession) -> String {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(|shape| &shape.text_stories)
        .find(|story| {
            story
                .paragraphs
                .iter()
                .filter(|paragraph| !paragraph.runs.is_empty())
                .count()
                >= 2
        })
        .unwrap()
        .id
        .clone()
}

#[test]
fn shared_update_opens_byte_identical_state_for_distinct_clients() {
    let source = DeckSession::open(FIXTURE, 101).unwrap();
    let seed = source.encode_state_as_update_v1();
    let left = DeckSession::open_from_update(&seed, 202).unwrap();
    let right = DeckSession::open_from_update(&seed, 303).unwrap();

    assert_eq!(left.client_id(), 202);
    assert_eq!(right.client_id(), 303);
    assert_eq!(left.snapshot().unwrap(), source.snapshot().unwrap());
    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(left.encode_state_as_update_v1(), seed);
    assert_eq!(right.encode_state_as_update_v1(), seed);
    assert_eq!(left.package().media, source.package().media);
}

#[test]
fn two_sessions_converge_after_text_slide_and_shape_edits() {
    let left = DeckSession::open(FIXTURE, 101).unwrap();
    let right = DeckSession::open(FIXTURE, 202).unwrap();
    let baseline = left.encode_state_vector_v1();
    assert_eq!(baseline, right.encode_state_vector_v1());
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());

    let initial = left.snapshot().unwrap();
    let first_slide = initial.slides[0].id.clone();
    let second_slide = initial.slides[1].id.clone();
    let third_slide = initial.slides[2].id.clone();
    let first_shape = initial.slides[0].shapes[0].id.clone();
    let removed_shape = initial.slides[1].shapes[1].id.clone();
    let story_id = first_text_story(&left);
    let insertion_index = left.story(&story_id).unwrap().length - 1;
    let context = EditCtx::local("local");

    left.insert_text(
        &context,
        &story_id,
        insertion_index,
        " LEFT",
        &TextStyle::default(),
    )
    .unwrap();
    left.move_slide(&context, &third_slide, 0).unwrap();
    left.move_shape(&context, &first_slide, &first_shape, 1_111_111, 2_222_222)
        .unwrap();

    right
        .insert_text(
            &context,
            &story_id,
            insertion_index,
            " RIGHT",
            &TextStyle::default(),
        )
        .unwrap();
    right.move_slide(&context, &first_slide, 2).unwrap();
    right
        .resize_shape(&context, &first_slide, &first_shape, 7_777_777, 888_888)
        .unwrap();
    right
        .add_text_box(
            &context,
            &second_slide,
            &ShapeDraft {
                name: "Collaborative note".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 200_000,
                    width: 3_000_000,
                    height: 800_000,
                },
                text: "Created on the right".to_owned(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
            },
        )
        .unwrap();
    right
        .remove_shape(&context, &second_slide, &removed_shape)
        .unwrap();

    let left_update = left.encode_diff_v1(&baseline).unwrap();
    let right_update = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&right_update).unwrap();
    right.apply_update_v1(&left_update).unwrap();

    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let converged = left.snapshot().unwrap();
    let edited_shape = converged
        .slides
        .iter()
        .find(|slide| slide.id == first_slide)
        .unwrap()
        .shapes
        .iter()
        .find(|shape| shape.id == first_shape)
        .unwrap();
    assert_eq!((edited_shape.x, edited_shape.y), (1_111_111, 2_222_222));
    assert_eq!(
        (edited_shape.width, edited_shape.height),
        (7_777_777, 888_888)
    );
    let text = left.story(&story_id).unwrap().plain_text();
    assert!(text.contains("LEFT"));
    assert!(text.contains("RIGHT"));
    let second = converged
        .slides
        .iter()
        .find(|slide| slide.id == second_slide)
        .unwrap();
    assert!(
        second
            .shapes
            .iter()
            .any(|shape| shape.name == "Collaborative note")
    );
    assert!(second.shapes.iter().all(|shape| shape.id != removed_shape));
}

#[test]
fn undo_reverts_only_local_origin_and_preserves_remote_text() {
    let left = DeckSession::open(FIXTURE, 303).unwrap();
    let right = DeckSession::open(FIXTURE, 404).unwrap();
    let baseline = left.encode_state_vector_v1();
    let story_id = first_text_story(&left);
    let index = left.story(&story_id).unwrap().length - 1;
    right
        .insert_text(
            &EditCtx::local("right"),
            &story_id,
            index,
            " REMOTE",
            &TextStyle::default(),
        )
        .unwrap();
    left.apply_update_v1(&right.encode_diff_v1(&baseline).unwrap())
        .unwrap();
    left.insert_text(
        &EditCtx::local("left"),
        &story_id,
        index,
        " LOCAL",
        &TextStyle::default(),
    )
    .unwrap();
    left.add_undo_barrier();
    assert!(left.can_undo());
    assert!(left.undo());
    let text = left.story(&story_id).unwrap().plain_text();
    assert!(text.contains("REMOTE"));
    assert!(!text.contains("LOCAL"));
    assert!(left.can_redo());
    assert!(left.redo());
    assert!(
        left.story(&story_id)
            .unwrap()
            .plain_text()
            .contains("LOCAL")
    );
}

#[test]
fn cross_paragraph_format_is_one_undoable_operation() {
    let session = DeckSession::open(FIXTURE, 405).unwrap();
    let story_id = first_multi_paragraph_story(&session);
    let before = session.story(&story_id).unwrap();

    let receipt = session
        .format_text(
            &EditCtx::local("local"),
            &story_id,
            0,
            before.length,
            &TextStylePatch {
                bold: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();

    assert_eq!((receipt.start, receipt.end), (0, before.length));
    let formatted = session.story(&story_id).unwrap();
    let formatted_paragraphs = formatted
        .paragraphs
        .iter()
        .filter(|paragraph| !paragraph.runs.is_empty())
        .collect::<Vec<_>>();
    assert!(formatted_paragraphs.len() >= 2);
    assert!(formatted_paragraphs.iter().all(|paragraph| {
        paragraph
            .runs
            .iter()
            .all(|run| run.style.bold == Some(true))
    }));

    session.add_undo_barrier();
    assert!(session.undo());
    assert_eq!(session.story(&story_id).unwrap(), before);
    assert!(!session.can_undo());
}

#[test]
fn cross_paragraph_format_rejects_out_of_bounds_ranges() {
    let session = DeckSession::open(FIXTURE, 406).unwrap();
    let story_id = first_multi_paragraph_story(&session);
    let length = session.story(&story_id).unwrap().length;

    assert!(matches!(
        session.format_text(
            &EditCtx::local("local"),
            &story_id,
            0,
            length + 1,
            &TextStylePatch {
                italic: Some(true),
                ..TextStylePatch::default()
            },
        ),
        Err(EditError::OutOfBounds {
            index,
            length: story_length
        }) if index == length + 1 && story_length == length
    ));
}

#[test]
fn update_events_tag_local_and_remote_origins() {
    let left = DeckSession::open(FIXTURE, 505).unwrap();
    let right = DeckSession::open(FIXTURE, 606).unwrap();
    let baseline = left.encode_state_vector_v1();
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = left
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.origin))
        .unwrap();
    let story_id = first_text_story(&left);
    let index = left.story(&story_id).unwrap().length - 1;
    left.insert_text(
        &EditCtx::local("left"),
        &story_id,
        index,
        " local",
        &TextStyle::default(),
    )
    .unwrap();
    right
        .insert_text(
            &EditCtx::local("right"),
            &story_id,
            index,
            " remote",
            &TextStyle::default(),
        )
        .unwrap();
    left.apply_update_v1(&right.encode_diff_v1(&baseline).unwrap())
        .unwrap();
    assert_eq!(
        events.lock().unwrap().as_slice(),
        [UpdateOrigin::Local, UpdateOrigin::Remote]
    );
}

#[test]
fn malformed_updates_and_state_vectors_leave_state_unchanged() {
    let session = DeckSession::open(FIXTURE, 707).unwrap();
    let state = session.encode_state_as_update_v1();
    assert!(session.apply_update_v1(&[0xff]).is_err());
    assert!(session.encode_diff_v1(&[0xff]).is_err());
    assert_eq!(session.encode_state_as_update_v1(), state);
}

#[test]
fn explicit_text_style_converges_after_update_exchange() {
    let left = DeckSession::open(FIXTURE, 808).unwrap();
    let right = DeckSession::open(FIXTURE, 909).unwrap();
    let baseline = right.encode_state_vector_v1();
    let story_id = first_text_story(&left);
    let index = left.story(&story_id).unwrap().length - 1;
    let first = right.snapshot().unwrap().slides[0].clone();
    left.insert_text(
        &EditCtx::local("left"),
        &story_id,
        index,
        " styled",
        &TextStyle {
            bold: Some(true),
            ..TextStyle::default()
        },
    )
    .unwrap();
    right
        .insert_slide(
            &EditCtx::local("right"),
            1,
            first.layout_part_path.as_deref(),
        )
        .unwrap();
    left.apply_update_v1(&right.encode_diff_v1(&baseline).unwrap())
        .unwrap();
    right
        .apply_update_v1(&left.encode_diff_v1(&baseline).unwrap())
        .unwrap();
    assert_eq!(
        left.story(&story_id).unwrap(),
        right.story(&story_id).unwrap()
    );
}
