use std::collections::HashSet;
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
fn cross_paragraph_alignment_is_one_undoable_operation() {
    let session = DeckSession::open(FIXTURE, 407).unwrap();
    let story_id = first_multi_paragraph_story(&session);
    let before = session.story(&story_id).unwrap();

    session
        .set_paragraph_alignment(
            &EditCtx::local("local"),
            &story_id,
            0,
            before.length,
            Some("r"),
        )
        .unwrap();

    let aligned = session.story(&story_id).unwrap();
    assert!(aligned.paragraphs.len() >= 2);
    assert!(
        aligned
            .paragraphs
            .iter()
            .all(|paragraph| paragraph.alignment.as_deref() == Some("r"))
    );

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

#[test]
fn two_sessions_converge_after_comment_edits() {
    let left = DeckSession::open(FIXTURE, 404).unwrap();
    let right = DeckSession::open(FIXTURE, 505).unwrap();
    let context = EditCtx::local("local");
    let initial = left.snapshot().unwrap();
    let first_slide = initial.slides[0].id.clone();
    let second_slide = initial.slides[1].id.clone();

    let mine = left
        .add_comment(
            &context,
            &first_slide,
            "Ada Lovelace",
            "AL",
            "Check this figure.",
            "2026-09-01T10:00:00.000",
            914_400,
            457_200,
        )
        .unwrap();
    right
        .add_comment(
            &context,
            &second_slide,
            "Grace Hopper",
            "GH",
            "Reword the heading.",
            "2026-09-01T10:01:00.000",
            0,
            0,
        )
        .unwrap();

    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let converged = left.snapshot().unwrap();
    assert_eq!(converged.comments.len(), 2);
    assert_eq!(
        converged
            .comments
            .iter()
            .map(|comment| comment.text.as_str())
            .collect::<Vec<_>>(),
        ["Check this figure.", "Reword the heading."],
        "comments order by (created, id) on every peer"
    );

    left.remove_comment(&context, &mine.comment_id).unwrap();
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(right.snapshot().unwrap().comments.len(), 1);
}

#[test]
fn comment_ids_are_deterministic_per_client_not_random() {
    let left = DeckSession::open(FIXTURE, 606).unwrap();
    let right = DeckSession::open(FIXTURE, 606).unwrap();
    let context = EditCtx::local("local");
    let slide = left.snapshot().unwrap().slides[0].id.clone();
    let arguments = ("Ada", "AL", "Same edit.", "2026-09-01T10:00:00.000");

    let mine = left
        .add_comment(
            &context,
            &slide,
            arguments.0,
            arguments.1,
            arguments.2,
            arguments.3,
            0,
            0,
        )
        .unwrap();
    let theirs = right
        .add_comment(
            &context,
            &slide,
            arguments.0,
            arguments.1,
            arguments.2,
            arguments.3,
            0,
            0,
        )
        .unwrap();

    assert_eq!(mine.comment_id, theirs.comment_id);
    assert_eq!(mine.comment_id, "comment:606:0");
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let second = left
        .add_comment(&context, &slide, "Ada", "AL", "Second", arguments.3, 0, 0)
        .unwrap();
    assert_eq!(second.comment_id, "comment:606:1");
    assert_ne!(mine.comment_id, second.comment_id);
}

#[test]
fn manual_capture_groups_text_and_comment_edits_until_a_boundary() {
    use pptx_edit::UndoCaptureMode;
    let session = DeckSession::open(FIXTURE, 9904).unwrap();
    let before = session.snapshot().unwrap();
    let story = first_text_story(&session);
    let context = EditCtx::local("host");
    assert_eq!(session.undo_capture_mode(), UndoCaptureMode::Auto);
    session.set_undo_capture_mode(UndoCaptureMode::Manual);
    session
        .insert_text(&context, &story, 0, "First ", &TextStyle::default())
        .unwrap();
    std::thread::sleep(std::time::Duration::from_millis(510));
    session.set_undo_capture_mode(UndoCaptureMode::Manual);
    session
        .add_comment(
            &context,
            &before.slides[0].id,
            "Host",
            "H",
            "Grouped",
            "2026-09-23T00:00:00Z",
            0,
            0,
        )
        .unwrap();
    session.add_undo_barrier();
    let grouped = session.snapshot().unwrap();
    session
        .insert_text(&context, &story, 0, "Second ", &TextStyle::default())
        .unwrap();
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), grouped);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.redo());
    assert_eq!(session.snapshot().unwrap(), grouped);
    session.set_undo_capture_mode(UndoCaptureMode::Auto);
    assert_eq!(session.undo_capture_mode(), UndoCaptureMode::Auto);
    assert!(session.can_redo());
}

fn sync(from: &DeckSession, to: &DeckSession) {
    to.apply_update_v1(&from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap())
        .unwrap();
}

fn exchange(peers: &[&DeckSession]) {
    for from in peers {
        for to in peers {
            if !std::ptr::eq(*from, *to) {
                sync(from, to);
            }
        }
    }
}

fn assert_converged(peers: &[&DeckSession]) -> Vec<String> {
    let first = peers[0];
    let snapshot = first.snapshot().unwrap();
    let ids = first.slide_ids().unwrap();
    assert_eq!(
        snapshot
            .slides
            .iter()
            .map(|slide| slide.id.clone())
            .collect::<Vec<_>>(),
        ids
    );
    assert_eq!(ids.iter().collect::<HashSet<_>>().len(), ids.len());
    for (index, slide) in snapshot.slides.iter().enumerate() {
        assert_eq!(&first.slide_scope(index).unwrap().slide, slide);
    }
    assert!(matches!(
        first.slide_scope(ids.len()),
        Err(EditError::OutOfBounds { .. })
    ));
    let saved = DeckSession::open(&first.save().unwrap(), 9).unwrap();
    assert_eq!(saved.slide_ids().unwrap().len(), ids.len());
    for peer in &peers[1..] {
        assert_eq!(
            peer.encode_state_vector_v1(),
            first.encode_state_vector_v1()
        );
        assert_eq!(peer.slide_ids().unwrap(), ids);
        assert_eq!(peer.snapshot().unwrap(), snapshot);
    }
    ids
}

fn slide_ids(session: &DeckSession) -> [String; 3] {
    session.slide_ids().unwrap().try_into().unwrap()
}

#[test]
fn concurrent_move_and_delete_of_one_slide_converge_in_either_delivery_order() {
    let context = EditCtx::local("local");
    for move_first in [true, false] {
        let mover = DeckSession::open(FIXTURE, 601).unwrap();
        let deleter = DeckSession::open(FIXTURE, 602).unwrap();
        let observer = DeckSession::open(FIXTURE, 603).unwrap();
        let [first, second, third] = slide_ids(&mover);
        mover.move_slide(&context, &first, 1).unwrap();
        deleter.delete_slide(&context, &first).unwrap();
        let updates = [
            mover.encode_state_as_update_v1(),
            deleter.encode_state_as_update_v1(),
        ];
        let order = if move_first { [0, 1] } else { [1, 0] };
        for index in order {
            observer.apply_update_v1(&updates[index]).unwrap();
        }
        deleter.apply_update_v1(&updates[0]).unwrap();
        mover.apply_update_v1(&updates[1]).unwrap();
        assert_eq!(
            assert_converged(&[&mover, &deleter, &observer]),
            [second.clone(), third.clone()]
        );
    }
}

#[test]
fn three_peers_converge_after_move_delete_and_move_of_one_slide_in_any_order() {
    let context = EditCtx::local("local");
    let peers = [601, 602, 603].map(|client| DeckSession::open(FIXTURE, client).unwrap());
    let [first, second, third] = slide_ids(&peers[0]);
    peers[0].move_slide(&context, &second, 0).unwrap();
    peers[1].delete_slide(&context, &second).unwrap();
    peers[2].move_slide(&context, &second, 2).unwrap();
    let updates = peers
        .each_ref()
        .map(|peer| peer.encode_state_as_update_v1());
    let orders = [
        [0, 1, 2],
        [0, 2, 1],
        [1, 0, 2],
        [1, 2, 0],
        [2, 0, 1],
        [2, 1, 0],
    ];
    let observers: Vec<DeckSession> = orders
        .iter()
        .enumerate()
        .map(|(client, order)| {
            let observer = DeckSession::open(FIXTURE, 700 + client as u64).unwrap();
            for index in order {
                observer.apply_update_v1(&updates[*index]).unwrap();
            }
            observer
        })
        .collect();
    exchange(&peers.each_ref());
    let mut all: Vec<&DeckSession> = peers.iter().collect();
    all.extend(observers.iter());
    assert_eq!(assert_converged(&all), [first, third]);
}

#[test]
fn slide_indices_follow_the_visible_order_after_concurrent_moves_of_one_slide() {
    let context = EditCtx::local("local");
    let left = DeckSession::open(FIXTURE, 611).unwrap();
    let right = DeckSession::open(FIXTURE, 612).unwrap();
    let [first, second, third] = slide_ids(&left);
    left.move_slide(&context, &third, 0).unwrap();
    right.move_slide(&context, &third, 1).unwrap();
    exchange(&[&left, &right]);
    assert_eq!(
        assert_converged(&[&left, &right]),
        [third.clone(), first.clone(), second.clone()]
    );

    let moved = left.move_slide(&context, &first, 2).unwrap();
    assert_eq!((moved.from_index, moved.to_index), (Some(1), Some(2)));
    assert_eq!(
        left.slide_ids().unwrap(),
        [third.clone(), second.clone(), first.clone()]
    );
    assert_eq!(left.slide_scope(2).unwrap().slide.id, first);

    let moved = right.move_slide(&context, &third, 2).unwrap();
    assert_eq!((moved.from_index, moved.to_index), (Some(0), Some(2)));
    assert_eq!(
        right.slide_ids().unwrap(),
        [first.clone(), second.clone(), third.clone()]
    );
    assert!(matches!(
        right.move_slide(&context, &second, 3),
        Err(EditError::OutOfBounds {
            index: 3,
            length: 3
        })
    ));
    assert!(matches!(
        right.insert_slide(&context, 4, None),
        Err(EditError::OutOfBounds {
            index: 4,
            length: 3
        })
    ));
    let inserted = right.insert_slide(&context, 3, None).unwrap();
    assert_eq!(inserted.to_index, Some(3));
    assert_eq!(right.slide_ids().unwrap()[3], inserted.slide_id);
    let duplicated = right.duplicate_slide(&context, &first, 1).unwrap();
    assert_eq!(right.slide_ids().unwrap()[1], duplicated.slide_id);
    let deleted = right.delete_slide(&context, &third).unwrap();
    assert_eq!(deleted.from_index, Some(3));

    exchange(&[&left, &right]);
    let ids = assert_converged(&[&left, &right]);
    assert_eq!(ids.len(), 4);
    assert!(!ids.contains(&third));
}

#[test]
fn deleting_a_slide_converges_with_concurrent_edits_inserts_and_moves_around_it() {
    let context = EditCtx::local("local");
    let peers = [621, 622, 623].map(|client| DeckSession::open(FIXTURE, client).unwrap());
    let initial = peers[0].snapshot().unwrap();
    let (deleted_index, story_id) = initial
        .slides
        .iter()
        .enumerate()
        .find_map(|(index, slide)| {
            slide
                .shapes
                .iter()
                .find_map(|shape| shape.text_stories.first())
                .map(|story| (index, story.id.clone()))
        })
        .unwrap();
    let deleted = initial.slides[deleted_index].id.clone();
    let other = initial
        .slides
        .iter()
        .rev()
        .find(|slide| slide.id != deleted)
        .unwrap()
        .id
        .clone();

    peers[0].delete_slide(&context, &deleted).unwrap();

    let editor = &peers[1];
    editor
        .add_text_box(
            &context,
            &deleted,
            &ShapeDraft {
                name: "Concurrent note".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 200_000,
                    width: 3_000_000,
                    height: 800_000,
                },
                text: "added while deleted".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let length = editor.story(&story_id).unwrap().length;
    editor
        .insert_text(
            &context,
            &story_id,
            length - 1,
            " edited",
            &TextStyle::default(),
        )
        .unwrap();
    editor
        .set_slide_notes(&context, &deleted, "notes while deleted")
        .unwrap();
    editor
        .add_comment(
            &context,
            &deleted,
            "Ada Lovelace",
            "AL",
            "Still needed?",
            "2026-09-01T10:00:00.000",
            0,
            0,
        )
        .unwrap();

    let arranger = &peers[2];
    let inserted = arranger
        .insert_slide(&context, deleted_index as u32 + 1, None)
        .unwrap();
    arranger.move_slide(&context, &other, 0).unwrap();
    let mut expected = arranger.slide_ids().unwrap();
    expected.retain(|id| id != &deleted);

    exchange(&peers.each_ref());
    assert_eq!(assert_converged(&peers.each_ref()), expected);
    assert!(expected.contains(&inserted.slide_id));
    let snapshot = peers[0].snapshot().unwrap();
    assert!(
        snapshot
            .comments
            .iter()
            .all(|comment| comment.slide_id != deleted)
    );
}

#[test]
fn undo_after_a_merged_move_and_delete_converges_on_every_peer() {
    let context = EditCtx::local("local");
    for mover_undoes_first in [true, false] {
        let mover = DeckSession::open(FIXTURE, 631).unwrap();
        let deleter = DeckSession::open(FIXTURE, 632).unwrap();
        let initial = mover.snapshot().unwrap();
        let [first, second, third] = slide_ids(&mover);
        mover.move_slide(&context, &first, 1).unwrap();
        mover.add_undo_barrier();
        deleter.delete_slide(&context, &first).unwrap();
        deleter.add_undo_barrier();
        exchange(&[&mover, &deleter]);
        assert_eq!(
            assert_converged(&[&mover, &deleter]),
            [second.clone(), third.clone()]
        );

        let undoers = if mover_undoes_first {
            [&mover, &deleter]
        } else {
            [&deleter, &mover]
        };
        for undoer in undoers {
            assert!(undoer.undo());
            exchange(&[&mover, &deleter]);
            assert_converged(&[&mover, &deleter]);
        }
        assert_eq!(
            assert_converged(&[&mover, &deleter]),
            [first.clone(), second.clone(), third.clone()]
        );
        assert_eq!(mover.snapshot().unwrap().slides[0], initial.slides[0]);
    }
}

#[test]
fn undo_after_concurrent_deletes_of_one_slide_restores_it_everywhere() {
    let context = EditCtx::local("local");
    let left = DeckSession::open(FIXTURE, 641).unwrap();
    let right = DeckSession::open(FIXTURE, 642).unwrap();
    let initial = left.snapshot().unwrap();
    let [first, second, third] = slide_ids(&left);
    left.delete_slide(&context, &first).unwrap();
    left.add_undo_barrier();
    right.delete_slide(&context, &first).unwrap();
    exchange(&[&left, &right]);
    assert_eq!(
        assert_converged(&[&left, &right]),
        [second.clone(), third.clone()]
    );
    assert!(left.undo());
    exchange(&[&left, &right]);
    assert_eq!(assert_converged(&[&left, &right]), [first, second, third]);
    assert_eq!(right.snapshot().unwrap().slides[0], initial.slides[0]);
}

fn diff_since(session: &DeckSession, state_vector: &mut Vec<u8>) -> Vec<u8> {
    let update = session.encode_diff_v1(state_vector).unwrap();
    *state_vector = session.encode_state_vector_v1();
    update
}

fn undo_and_redo_over_merged_concurrent_moves(delete: bool) {
    let context = EditCtx::local("local");
    for reverse in [false, true] {
        let peers = [651, 652, 653].map(|client| DeckSession::open(FIXTURE, client).unwrap());
        let [left, right, merger] = &peers;
        let [first, second, third] = slide_ids(left);
        left.move_slide(&context, &first, 1).unwrap();
        right.move_slide(&context, &first, 2).unwrap();
        let moves = [
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1(),
        ];
        sync(left, merger);
        sync(right, merger);
        let before = merger.snapshot().unwrap();
        assert_eq!(
            merger.slide_ids().unwrap(),
            [second.clone(), first.clone(), third.clone()]
        );

        let mut state_vector = merger.encode_state_vector_v1();
        let expected = if delete {
            merger.delete_slide(&context, &first).unwrap();
            vec![second.clone(), third.clone()]
        } else {
            merger.move_slide(&context, &first, 0).unwrap();
            vec![first.clone(), second.clone(), third.clone()]
        };
        let after = merger.snapshot().unwrap();
        assert_eq!(merger.slide_ids().unwrap(), expected);
        let mut updates = vec![diff_since(merger, &mut state_vector)];
        assert!(merger.undo());
        assert_eq!(merger.snapshot().unwrap(), before);
        updates.push(diff_since(merger, &mut state_vector));
        assert!(merger.redo());
        assert_eq!(merger.snapshot().unwrap(), after);
        updates.push(diff_since(merger, &mut state_vector));

        for (peer, other) in [(left, &moves[1]), (right, &moves[0])] {
            let mut delivery: Vec<&Vec<u8>> = std::iter::once(other).chain(&updates).collect();
            if reverse {
                delivery.reverse();
            }
            for update in delivery {
                peer.apply_update_v1(update).unwrap();
            }
        }
        assert_eq!(assert_converged(&peers.each_ref()), expected);
    }
}

#[test]
fn undo_and_redo_of_a_move_over_merged_concurrent_moves_restore_the_visible_order() {
    undo_and_redo_over_merged_concurrent_moves(false);
}

#[test]
fn undo_and_redo_of_a_delete_over_merged_concurrent_moves_restore_the_visible_order() {
    undo_and_redo_over_merged_concurrent_moves(true);
}

#[test]
fn a_peer_can_apply_what_the_loser_of_a_concurrent_move_publishes_before_the_winning_move() {
    let context = EditCtx::local("local");
    let peer = DeckSession::open(FIXTURE, 201).unwrap();
    let loser = DeckSession::open(FIXTURE, 202).unwrap();
    let winner = DeckSession::open(FIXTURE, 203).unwrap();
    let baseline = loser.encode_state_vector_v1();
    let initial = loser.snapshot().unwrap();
    let (slide, shape) = (&initial.slides[0].id, &initial.slides[0].shapes[0].id);
    loser
        .move_shape(&context, slide, shape, 1_111_111, 2_222_222)
        .unwrap();
    winner
        .move_shape(&context, slide, shape, 3_333_333, 4_444_444)
        .unwrap();
    let lost = loser.encode_diff_v1(&baseline).unwrap();
    let won = winner.encode_diff_v1(&baseline).unwrap();
    let seen = loser.snapshot().unwrap();

    let published = Arc::new(Mutex::new(Vec::new()));
    let sink = Arc::clone(&published);
    let _subscription = loser
        .observe_update_v1(move |event| {
            if event.origin == UpdateOrigin::Local {
                sink.lock().unwrap().push(event.update);
            }
        })
        .unwrap();
    loser.apply_update_v1(&won).unwrap();

    peer.apply_update_v1(&lost).unwrap();
    for update in published.lock().unwrap().iter() {
        peer.apply_update_v1(update).unwrap();
    }
    assert_eq!(peer.snapshot().unwrap(), seen);

    peer.apply_update_v1(&won).unwrap();
    winner.apply_update_v1(&lost).unwrap();
    assert_converged(&[&loser, &winner, &peer]);
}
