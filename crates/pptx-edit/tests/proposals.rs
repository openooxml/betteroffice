use pptx_edit::{DeckSession, EditCtx, ProposalEdit, ProposalError, ProposalRequest, ShapeRect};
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

const FILE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn edits(session: &DeckSession) -> Vec<ProposalEdit> {
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = slide
        .shapes
        .iter()
        .find(|shape| !shape.text_stories.is_empty())
        .unwrap();
    let story = &shape.text_stories[0];
    let end = story.paragraphs[0]
        .runs
        .iter()
        .map(|run| run.text.encode_utf16().count() as u32)
        .sum();
    vec![
        ProposalEdit::ReplaceText {
            story_id: story.id.clone(),
            start: 0,
            end,
            text: "A reviewed title".into(),
            style: None,
        },
        ProposalEdit::SetShapeRect {
            slide_id: slide.id.clone(),
            shape_id: shape.id.clone(),
            rect: ShapeRect {
                x: shape.x + 100_000,
                y: shape.y + 200_000,
                width: shape.width,
                height: shape.height,
            },
        },
    ]
}

fn request(edits: Vec<ProposalEdit>) -> ProposalRequest {
    ProposalRequest {
        agent_id: "editor-agent".into(),
        note: Some("Tighten the title and align it".into()),
        edits,
    }
}

#[test]
fn formatting_geometry_and_notes_preview_and_survive_save() {
    use pptx_edit::{PresetShapeDraft, ShapeStroke, TextStylePatch};
    use std::collections::BTreeMap;
    let session = DeckSession::open(FILE, 506).unwrap();
    let first = session.snapshot().unwrap().slides[0].clone();
    let rect = ShapeRect {
        x: 100_000,
        y: 100_000,
        width: 1_000_000,
        height: 1_000_000,
    };
    let added = session
        .add_shape(
            &EditCtx::local("human"),
            &first.id,
            &PresetShapeDraft {
                name: "Metric card".into(),
                geometry: "roundRect".into(),
                rect,
                fill: None,
            },
        )
        .unwrap();
    let mut proposed = edits(&session);
    let story_id = first
        .shapes
        .iter()
        .find_map(|shape| shape.text_stories.first())
        .unwrap()
        .id
        .clone();
    proposed.extend([
        ProposalEdit::FormatText {
            story_id: story_id.clone(),
            start: 0,
            end: 5,
            patch: TextStylePatch {
                bold: Some(true),
                color: Some("#2255AA".into()),
                ..Default::default()
            },
        },
        ProposalEdit::SetParagraphAlignment {
            story_id,
            start: 0,
            end: 5,
            alignment: Some("ctr".into()),
        },
        ProposalEdit::SetShapeFill {
            slide_id: first.id.clone(),
            shape_id: added.shape_id.clone(),
            color: Some("#CCDDFF".into()),
        },
        ProposalEdit::SetShapeStroke {
            slide_id: first.id.clone(),
            shape_id: added.shape_id.clone(),
            stroke: ShapeStroke {
                color: Some("#2255AA".into()),
                width_pt: Some(2.0),
            },
        },
        ProposalEdit::SetShapeAdjust {
            slide_id: first.id.clone(),
            shape_id: added.shape_id.clone(),
            adjustments: BTreeMap::from([("adj".into(), 0.2)]),
        },
        ProposalEdit::SetSlideNotes {
            slide_id: first.id.clone(),
            text: "Review the metric card".into(),
        },
    ]);
    let proposal = session.propose(request(proposed)).unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    let reopened = DeckSession::open(&session.save().unwrap(), 507)
        .unwrap()
        .snapshot()
        .unwrap();
    let slide = &reopened.slides[0];
    assert_eq!(slide.notes, "Review the metric card");
    let shape = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Metric card")
        .unwrap();
    assert_eq!(shape.resolved_fill_color.as_deref(), Some("#CCDDFF"));
    assert_eq!(shape.resolved_outline_color.as_deref(), Some("#2255AA"));
    assert_eq!(shape.adjust_values.get("adj"), Some(&0.2));
}

#[test]
fn noop_acceptance_does_not_create_history_and_invalid_wire_fields_are_rejected() {
    let session = DeckSession::open(FILE, 508).unwrap();
    let slide = &session.snapshot().unwrap().slides[0];
    let proposal = session
        .propose(request(vec![ProposalEdit::SetSlideNotes {
            slide_id: slide.id.clone(),
            text: slide.notes.clone(),
        }]))
        .unwrap();
    assert!(
        !session
            .accept_proposal(&proposal.id, false)
            .unwrap()
            .applied
    );
    assert!(!session.can_undo());
    assert!(
        serde_json::from_str::<ProposalEdit>(
            r#"{"type":"setSlideNotes","slideId":"s1","text":"value","typo":true}"#
        )
        .is_err()
    );
}

#[test]
fn staging_preview_acceptance_and_history_are_atomic() {
    let session = DeckSession::open(FILE, 501).unwrap();
    let before = session.snapshot().unwrap();
    let state = session.encode_state_as_update_v1();
    let events = Arc::new(AtomicUsize::new(0));
    let counter = events.clone();
    let _observer = session
        .observe_update_v1(move |_| {
            counter.fetch_add(1, Ordering::Relaxed);
        })
        .unwrap();
    let proposal = session.propose(request(edits(&session))).unwrap();
    assert_eq!(proposal.changes.len(), 1);
    assert!(proposal.changes[0].new_text.contains("A reviewed title"));
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_ne!(preview.snapshot, before);
    assert_eq!(session.snapshot().unwrap(), before);
    assert_eq!(session.encode_state_as_update_v1(), state);
    assert_eq!(events.load(Ordering::Relaxed), 0);
    assert!(!session.can_undo());
    let accepted = session.accept_proposal(&proposal.id, false).unwrap();
    assert!(accepted.applied);
    assert_eq!(accepted.snapshot, preview.snapshot);
    assert_eq!(events.load(Ordering::Relaxed), 1);
    assert!(session.proposals().unwrap().is_empty());
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.redo());
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::NotFound(_))
    ));
}

#[test]
fn rejection_and_failed_batches_leave_the_document_and_history_untouched() {
    let session = DeckSession::open(FILE, 502).unwrap();
    let before = session.encode_state_as_update_v1();
    let mut invalid = edits(&session);
    invalid.push(ProposalEdit::SetSlideNotes {
        slide_id: "missing".into(),
        text: "bad target".into(),
    });
    assert!(session.propose(request(invalid)).is_err());
    assert!(session.proposals().unwrap().is_empty());
    let first = session.propose(request(edits(&session))).unwrap();
    assert!(session.reject_proposal(&first.id));
    assert!(!session.reject_proposal(&first.id));
    let second = session.propose(request(edits(&session))).unwrap();
    assert_ne!(first.id, second.id);
    assert_eq!(session.encode_state_as_update_v1(), before);
    assert!(!session.can_undo());
}

#[test]
fn stale_targets_require_force_and_removed_targets_remain_invalid() {
    let session = DeckSession::open(FILE, 503).unwrap();
    let proposal = session.propose(request(edits(&session))).unwrap();
    let change = &proposal.changes[0];
    let shape = change.before.as_ref().unwrap();
    session
        .move_shape(
            &EditCtx::local("human"),
            &change.slide_id,
            &shape.id,
            shape.x + 500_000,
            shape.y,
        )
        .unwrap();
    assert_eq!(
        session.proposals().unwrap()[0].stale_targets,
        vec![shape.id.clone()]
    );
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(
        preview.proposal.changes[0].before.as_ref().unwrap().x,
        shape.x + 500_000
    );
    session.accept_proposal(&proposal.id, true).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    assert!(session.undo());
    let proposal = session
        .propose(request(vec![edits(&session)[0].clone()]))
        .unwrap();
    session
        .remove_shape(&EditCtx::local("human"), &change.slide_id, &shape.id)
        .unwrap();
    assert!(session.accept_proposal(&proposal.id, true).is_err());
    assert_eq!(session.proposals().unwrap().len(), 1);
}

#[test]
fn accepted_proposals_sync_and_undo_preserves_unrelated_peer_edits() {
    let local = DeckSession::open(FILE, 504).unwrap();
    let peer = DeckSession::open(FILE, 505).unwrap();
    let original = local.snapshot().unwrap();
    let proposal = local.propose(request(edits(&local))).unwrap();
    peer.apply_update_v1(
        &local
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert!(peer.proposals().unwrap().is_empty());
    assert_eq!(peer.snapshot().unwrap(), original);
    let other_slide = &original.slides[1].id;
    peer.set_slide_notes(
        &EditCtx::local("peer"),
        other_slide,
        "Keep the peer's notes",
    )
    .unwrap();
    local
        .apply_update_v1(
            &peer
                .encode_diff_v1(&local.encode_state_vector_v1())
                .unwrap(),
        )
        .unwrap();
    assert!(local.proposals().unwrap()[0].stale_targets.is_empty());
    local.accept_proposal(&proposal.id, false).unwrap();
    peer.apply_update_v1(
        &local
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(local.snapshot().unwrap(), peer.snapshot().unwrap());
    assert!(local.undo());
    peer.apply_update_v1(
        &local
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(local.snapshot().unwrap(), peer.snapshot().unwrap());
    assert_eq!(peer.snapshot().unwrap().slides[0], original.slides[0]);
    assert_eq!(
        peer.snapshot().unwrap().slides[1].notes,
        "Keep the peer's notes"
    );
}

#[test]
fn inline_diff_preserves_context_utf16_offsets_and_never_enters_the_document() {
    use pptx_edit::{ProposalTextChangeKind, ShapeDraft, TextStyle, TextStylePatch};
    let session = DeckSession::open(FILE, 509).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let text = "Q1 revenue 😀 and costs fell";
    let added = session
        .add_text_box(
            &EditCtx::local("human"),
            &slide_id,
            &ShapeDraft {
                name: "Review target".into(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 100_000,
                    width: 2_000_000,
                    height: 1_000_000,
                },
                text: text.into(),
                style: TextStyle {
                    font_size_pt: Some(24.0),
                    ..Default::default()
                },
            },
        )
        .unwrap();
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == added.shape_id)
        .unwrap();
    let story = &shape.text_stories[0];
    let end = text.encode_utf16().count() as u32;
    let proposal = session
        .propose(request(vec![
            ProposalEdit::ReplaceText {
                story_id: story.id.clone(),
                start: 0,
                end: 2,
                text: "Q2".into(),
                style: None,
            },
            ProposalEdit::ReplaceText {
                story_id: story.id.clone(),
                start: end - 4,
                end,
                text: "rose".into(),
                style: None,
            },
            ProposalEdit::FormatText {
                story_id: story.id.clone(),
                start: 3,
                end: 10,
                patch: TextStylePatch {
                    bold: Some(true),
                    ..Default::default()
                },
            },
        ]))
        .unwrap();
    let state = session.encode_state_as_update_v1();
    let diff = session.preview_proposal_diff(&proposal.id).unwrap();
    assert_eq!(session.encode_state_as_update_v1(), state);
    let mixed = &diff.snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == added.shape_id)
        .unwrap()
        .text_stories[0];
    assert!(mixed.plain_text().contains("😀 and costs"));
    assert_eq!(mixed.plain_text().matches("😀").count(), 1);
    assert_eq!(
        mixed.length,
        mixed.plain_text().encode_utf16().count() as u32 + 1
    );
    let without = |kind| {
        let mut offset = 0;
        mixed
            .plain_text()
            .chars()
            .filter(|ch| {
                let at = offset;
                offset += ch.len_utf16() as u32;
                !diff
                    .text_changes
                    .iter()
                    .any(|change| change.kind == kind && change.start <= at && at < change.end)
            })
            .collect::<String>()
    };
    assert_eq!(without(ProposalTextChangeKind::Insertion), text);
    assert_eq!(
        without(ProposalTextChangeKind::Deletion),
        "Q2 revenue 😀 and costs rose"
    );
    assert!(
        mixed.paragraphs[0]
            .runs
            .iter()
            .any(|run| run.text == "revenue"
                && run.style.bold == Some(true)
                && run.style.color.as_deref() == Some("#166534"))
    );
    let accepted = session.preview_proposal(&proposal.id).unwrap().snapshot;
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), accepted);
    assert_eq!(
        session.story(&story.id).unwrap().plain_text(),
        "Q2 revenue 😀 and costs rose"
    );
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), snapshot);
}

#[test]
fn inline_diff_uses_current_targets_and_handles_long_replacements() {
    let session = DeckSession::open(FILE, 510).unwrap();
    let proposal = session
        .propose(request(vec![edits(&session)[0].clone()]))
        .unwrap();
    let story_id = match &proposal.edits[0] {
        ProposalEdit::ReplaceText { story_id, .. } => story_id.clone(),
        _ => unreachable!(),
    };
    session
        .insert_text(
            &EditCtx::local("human"),
            &story_id,
            0,
            "Human ",
            &Default::default(),
        )
        .unwrap();
    let state = session.encode_state_as_update_v1();
    let diff = session.preview_proposal_diff(&proposal.id).unwrap();
    assert!(diff.proposal.changes[0].old_text.starts_with("Human "));
    assert!(!diff.proposal.stale_targets.is_empty());
    assert_eq!(session.encode_state_as_update_v1(), state);
    session.reject_proposal(&proposal.id);
    let old = "old ".repeat(600);
    let new = "new ".repeat(600);
    let end = session.story(&story_id).unwrap().paragraphs[0]
        .runs
        .iter()
        .map(|run| run.text.encode_utf16().count() as u32)
        .sum();
    session
        .delete_text(&EditCtx::local("human"), &story_id, 0, end)
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("human"),
            &story_id,
            0,
            &old,
            &Default::default(),
        )
        .unwrap();
    let proposal = session
        .propose(request(vec![ProposalEdit::ReplaceText {
            story_id,
            start: 0,
            end: old.len() as u32,
            text: new,
            style: None,
        }]))
        .unwrap();
    let diff = session.preview_proposal_diff(&proposal.id).unwrap();
    assert!(!diff.text_changes.is_empty());
    assert_eq!(session.proposals().unwrap().len(), 1);
}

#[test]
fn scoped_diff_preview_matches_the_deck_wide_one() {
    let session = DeckSession::open(FILE, 509).unwrap();
    let proposal = session.propose(request(edits(&session))).unwrap();
    let deck = session.preview_proposal_diff(&proposal.id).unwrap();
    let scoped = session
        .preview_proposal_diff_slide(&proposal.id, 0)
        .unwrap();
    assert_eq!(scoped.proposal, deck.proposal);
    assert_eq!(scoped.text_changes, deck.text_changes);
    assert_eq!(scoped.scope.slide, deck.snapshot.slides[0]);
    assert_eq!(scoped.snapshot.slides.len(), 1);
    assert_eq!(scoped.snapshot.slides[0], scoped.scope.slide);
}

#[test]
fn scoped_diff_preview_reports_stale_targets_like_the_deck_wide_one() {
    let session = DeckSession::open(FILE, 510).unwrap();
    let proposal = session.propose(request(edits(&session))).unwrap();
    let story_id = match &proposal.edits[0] {
        ProposalEdit::ReplaceText { story_id, .. } => story_id.clone(),
        _ => unreachable!(),
    };
    session
        .insert_text(
            &EditCtx::local("human"),
            &story_id,
            0,
            "Human ",
            &Default::default(),
        )
        .unwrap();
    let deck = session.preview_proposal_diff(&proposal.id).unwrap();
    let scoped = session
        .preview_proposal_diff_slide(&proposal.id, 0)
        .unwrap();
    assert_eq!(scoped.proposal.stale_targets, deck.proposal.stale_targets);
    assert!(!scoped.proposal.stale_targets.is_empty());
    assert_eq!(scoped.text_changes, deck.text_changes);
}

#[test]
fn scoped_diff_preview_rejects_unknown_ids_and_out_of_range_slides() {
    let session = DeckSession::open(FILE, 511).unwrap();
    let proposal = session.propose(request(edits(&session))).unwrap();
    let slide_count = session.snapshot().unwrap().slides.len();
    assert!(session.preview_proposal_diff_slide("nope", 0).is_err());
    assert!(
        session
            .preview_proposal_diff_slide(&proposal.id, slide_count)
            .is_err()
    );
}
