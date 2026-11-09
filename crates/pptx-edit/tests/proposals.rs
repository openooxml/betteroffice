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

#[test]
fn slide_structure_previews_accepts_rejects_and_undoes_atomically() {
    for kind in ["insertSlide", "duplicateSlide", "deleteSlide", "moveSlide"] {
        let session = DeckSession::open(FILE, 520).unwrap();
        let before = session.snapshot().unwrap();
        let first = &before.slides[0].id;
        let edit = match kind {
            "insertSlide" => serde_json::json!({"type": kind, "ref": "new", "afterSlideId": first}),
            "duplicateSlide" => serde_json::json!({"type": kind, "ref": "new", "slideId": first}),
            "deleteSlide" => serde_json::json!({"type": kind, "slideId": first}),
            _ => {
                serde_json::json!({"type": kind, "slideId": first, "afterSlideId": before.slides.last().unwrap().id})
            }
        };
        let edit: ProposalEdit = serde_json::from_value(edit).unwrap();
        let rejected = session.propose(request(vec![edit.clone()])).unwrap();
        assert!(session.reject_proposal(&rejected.id));
        assert_eq!(session.snapshot().unwrap(), before);
        let proposal = session.propose(request(vec![edit])).unwrap();
        let state = session.encode_state_as_update_v1();
        let preview = session.preview_proposal(&proposal.id).unwrap();
        let strip = session.proposal_slide_strip(&proposal.id).unwrap();
        assert!(
            strip
                .iter()
                .any(|entry| entry.status != pptx_edit::ProposalStripStatus::Unchanged)
        );
        let target = proposal
            .changes
            .iter()
            .find(|change| change.kind.is_some())
            .unwrap();
        let diff = session
            .preview_proposal_diff_slide_by_id(&proposal.id, &target.slide_id)
            .unwrap();
        assert_eq!(diff.scope.slide.id, target.slide_id);
        assert_eq!(session.encode_state_as_update_v1(), state);
        session.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
        assert!(session.redo());
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    }
}

#[test]
fn slide_refs_chain_and_invalid_refs_are_atomic() {
    let session = DeckSession::open(FILE, 521).unwrap();
    let before = session.snapshot().unwrap();
    let edits = serde_json::from_value(serde_json::json!([
        {"type": "insertSlide", "ref": "s1", "afterSlideId": null},
        {"type": "setSlideNotes", "slideId": "@s1", "text": "Draft notes"},
        {"type": "duplicateSlide", "ref": "s2", "slideId": "@s1", "afterSlideId": "@s1"},
        {"type": "moveSlide", "slideId": "@s2", "afterSlideId": null}
    ]))
    .unwrap();
    let proposal = session.propose(request(edits)).unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(preview.snapshot.slides[0].id, proposal.refs["s2"]);
    assert_eq!(preview.snapshot.slides[1].id, proposal.refs["s1"]);
    assert_eq!(preview.snapshot.slides[0].notes, "Draft notes");
    session.accept_proposal(&proposal.id, false).unwrap();
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    for edits in [
        serde_json::json!([{"type": "setSlideNotes", "slideId": "@missing", "text": "x"}]),
        serde_json::json!([
            {"type": "insertSlide", "ref": "same", "afterSlideId": null},
            {"type": "insertSlide", "ref": "same", "afterSlideId": null}
        ]),
    ] {
        assert!(matches!(
            session.propose(request(serde_json::from_value(edits).unwrap())),
            Err(ProposalError::Invalid(_))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
    }
}

#[test]
fn structural_stale_rules_and_forced_missing_targets() {
    for kind in ["deleteSlide", "duplicateSlide"] {
        let session = DeckSession::open(FILE, 522).unwrap();
        let before = session.snapshot().unwrap();
        let source = &before.slides[0].id;
        let proposal = session
            .propose(request(vec![
                serde_json::from_value(serde_json::json!({"type": kind, "slideId": source}))
                    .unwrap(),
            ]))
            .unwrap();
        session
            .set_slide_notes(&EditCtx::local("human"), source, "New work")
            .unwrap();
        assert_eq!(
            session.proposals().unwrap()[0].stale_targets,
            vec![source.clone()]
        );
        assert!(matches!(
            session.accept_proposal(&proposal.id, false),
            Err(ProposalError::Stale(_))
        ));
        session
            .delete_slide(&EditCtx::local("human"), source)
            .unwrap();
        assert!(session.accept_proposal(&proposal.id, true).is_err());
    }
    for kind in ["insertSlide", "moveSlide", "duplicateSlide"] {
        let session = DeckSession::open(FILE, 523).unwrap();
        let before = session.snapshot().unwrap();
        let source = &before.slides[0].id;
        let anchor = &before.slides[1].id;
        let mut edit = serde_json::json!({"type": kind, "afterSlideId": anchor});
        if kind != "insertSlide" {
            edit["slideId"] = source.clone().into();
        }
        let proposal = session
            .propose(request(vec![serde_json::from_value(edit).unwrap()]))
            .unwrap();
        session
            .move_slide(&EditCtx::local("human"), anchor, 0)
            .unwrap();
        assert!(session.proposals().unwrap()[0].stale_targets.is_empty());
        session
            .set_slide_notes(&EditCtx::local("human"), anchor, "Anchor content")
            .unwrap();
        assert!(session.proposals().unwrap()[0].stale_targets.is_empty());
        session
            .delete_slide(&EditCtx::local("human"), anchor)
            .unwrap();
        assert!(
            session.proposals().unwrap()[0]
                .stale_targets
                .contains(anchor)
        );
        assert!(matches!(
            session.accept_proposal(&proposal.id, false),
            Err(ProposalError::Stale(_))
        ));
        let preview = session.preview_proposal(&proposal.id).unwrap();
        session.accept_proposal(&proposal.id, true).unwrap();
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        let last = session
            .snapshot()
            .unwrap()
            .slides
            .last()
            .unwrap()
            .id
            .clone();
        if kind == "moveSlide" {
            assert_eq!(&last, source);
        } else {
            assert!(!before.slides.iter().any(|slide| slide.id == last));
        }
    }
}

#[test]
fn structural_proposals_keep_edit_and_pending_limits() {
    let session = DeckSession::open(FILE, 524).unwrap();
    let edit: ProposalEdit =
        serde_json::from_value(serde_json::json!({"type": "insertSlide", "afterSlideId": null}))
            .unwrap();
    assert!(session.propose(request(vec![])).is_err());
    assert!(session.propose(request(vec![edit.clone(); 257])).is_err());
    let maximum = session.propose(request(vec![edit.clone(); 256])).unwrap();
    session.reject_proposal(&maximum.id);
    for _ in 0..64 {
        session.propose(request(vec![edit.clone()])).unwrap();
    }
    assert!(session.propose(request(vec![edit])).is_err());
}

#[test]
fn duplicate_destination_distinguishes_omitted_and_null() {
    let session = DeckSession::open(FILE, 525).unwrap();
    let source = session.snapshot().unwrap().slides[0].id.clone();
    for (destination, expected) in [(None, 1), (Some(serde_json::Value::Null), 0)] {
        let mut value =
            serde_json::json!({"type": "duplicateSlide", "ref": "copy", "slideId": source});
        if let Some(destination) = destination {
            value["afterSlideId"] = destination;
        }
        let edit: ProposalEdit = serde_json::from_value(value.clone()).unwrap();
        let serialized = serde_json::to_value(&edit).unwrap();
        assert_eq!(serialized.get("afterSlideId"), value.get("afterSlideId"));
        let proposal = session.propose(request(vec![edit])).unwrap();
        assert_eq!(
            session
                .preview_proposal(&proposal.id)
                .unwrap()
                .snapshot
                .slides[expected]
                .id,
            proposal.refs["copy"]
        );
        session.reject_proposal(&proposal.id);
    }
}

#[test]
fn missing_moved_slide_errors_even_when_forced_and_ref_ids_refresh() {
    let session = DeckSession::open(FILE, 526).unwrap();
    let source = session.snapshot().unwrap().slides[0].id.clone();
    let proposal = session
        .propose(request(vec![ProposalEdit::MoveSlide {
            slide_id: source.clone(),
            after_slide_id: None,
        }]))
        .unwrap();
    session
        .delete_slide(&EditCtx::local("human"), &source)
        .unwrap();
    assert!(matches!(
        session.accept_proposal(&proposal.id, true),
        Err(ProposalError::Edit(pptx_edit::EditError::SlideNotFound(_)))
    ));
    session.reject_proposal(&proposal.id);
    let proposal = session
        .propose(request(
            serde_json::from_value(serde_json::json!([
                {"type": "insertSlide", "ref": "draft", "afterSlideId": null},
                {"type": "setSlideNotes", "slideId": "@draft", "text": "Draft"}
            ]))
            .unwrap(),
        ))
        .unwrap();
    session
        .insert_slide(&EditCtx::local("human"), 0, None)
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_ne!(preview.proposal.refs["draft"], proposal.refs["draft"]);
    assert_eq!(session.proposals().unwrap()[0].refs, preview.proposal.refs);
    assert_eq!(
        session.proposals().unwrap()[0].changes,
        preview.proposal.changes
    );
    assert!(preview.proposal.stale_targets.is_empty());
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
}

#[test]
fn accepted_slide_structure_syncs_as_an_ordinary_update() {
    let session = DeckSession::open(FILE, 527).unwrap();
    let peer = DeckSession::open(FILE, 528).unwrap();
    let before = session.snapshot().unwrap();
    let proposal = session
        .propose(request(
            serde_json::from_value(serde_json::json!([
                {"type": "duplicateSlide", "ref": "copy", "slideId": before.slides[0].id},
                {"type": "setSlideNotes", "slideId": "@copy", "text": "Copied notes"},
                {"type": "moveSlide", "slideId": "@copy", "afterSlideId": null}
            ]))
            .unwrap(),
        ))
        .unwrap();
    session.accept_proposal(&proposal.id, false).unwrap();
    peer.apply_update_v1(
        &session
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(peer.snapshot().unwrap(), session.snapshot().unwrap());
    assert!(peer.proposals().unwrap().is_empty());
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    peer.apply_update_v1(
        &session
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(peer.snapshot().unwrap(), before);
}

#[test]
fn structural_accept_is_isolated_in_manual_undo_capture() {
    let session = DeckSession::open(FILE, 529).unwrap();
    session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
    let first = session.snapshot().unwrap().slides[0].id.clone();
    session
        .set_slide_notes(&EditCtx::local("human"), &first, "Before acceptance")
        .unwrap();
    let before = session.snapshot().unwrap();
    let proposal = session
        .propose(request(
            serde_json::from_value(serde_json::json!([
                {"type": "insertSlide", "ref": "new", "afterSlideId": null},
                {"type": "setSlideNotes", "slideId": "@new", "text": "Proposal notes"}
            ]))
            .unwrap(),
        ))
        .unwrap();
    session.accept_proposal(&proposal.id, false).unwrap();
    let accepted = session.snapshot().unwrap();
    session
        .set_slide_notes(&EditCtx::local("human"), &first, "After acceptance")
        .unwrap();
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), accepted);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.can_undo());
}

const SHAPE_DECK: &[u8] = include_bytes!("fixtures/proposal-shapes.pptx");
const PIXEL: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

fn shape_edit(kind: &str, slide: &str, shape: &str) -> ProposalEdit {
    let rect = serde_json::json!({"x": 100000, "y": 100000, "width": 1000000, "height": 500000});
    let draft = match kind {
        "addTextBox" => {
            serde_json::json!({"name": "Draft box", "rect": rect, "text": "", "style": {}})
        }
        "addShape" => {
            serde_json::json!({"name": "Draft shape", "rect": rect, "geometry": "rect", "fill": "#123456"})
        }
        _ => {
            serde_json::json!({"name": "Draft picture", "rect": rect, "contentType": "image/png", "mediaBase64": PIXEL})
        }
    };
    serde_json::from_value(if kind == "removeShape" {
        serde_json::json!({"type": kind, "slideId": slide, "shapeId": shape})
    } else {
        serde_json::json!({"type": kind, "ref": "created", "slideId": slide, "draft": draft})
    })
    .unwrap()
}

#[test]
fn shape_variants_preview_accept_reject_and_undo() {
    for kind in ["addTextBox", "addShape", "addPicture", "removeShape"] {
        let session = DeckSession::open(SHAPE_DECK, 530).unwrap();
        session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let edit = shape_edit(kind, &slide.id, &slide.shapes[0].id);
        let rejected = session.propose(request(vec![edit.clone()])).unwrap();
        assert!(session.reject_proposal(&rejected.id));
        assert_eq!(session.snapshot().unwrap(), before);
        let proposal = session.propose(request(vec![edit])).unwrap();
        let preview = session.preview_proposal(&proposal.id).unwrap();
        let change = &proposal.changes[0];
        assert_eq!(
            change.kind,
            Some(if kind == "removeShape" {
                pptx_edit::ProposalChangeKind::Deleted
            } else {
                pptx_edit::ProposalChangeKind::Inserted
            })
        );
        assert_eq!(change.before.is_none(), kind != "removeShape");
        assert_eq!(change.after.is_none(), kind == "removeShape");
        let diff = session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert!(
            diff.scope
                .slide
                .shapes
                .iter()
                .any(|shape| Some(&shape.id) == change.shape_id.as_ref())
        );
        assert_eq!(
            diff.shape_changes[0].shape_id,
            *change.shape_id.as_ref().unwrap()
        );
        assert_eq!(session.snapshot().unwrap(), before);
        session.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.redo());
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    }
}

#[test]
fn slide_shape_and_story_refs_chain_and_refresh() {
    let session = DeckSession::open(SHAPE_DECK, 531).unwrap();
    let mut box_edit = shape_edit("addTextBox", "@slide", "");
    if let ProposalEdit::AddTextBox { reference, .. } = &mut box_edit {
        *reference = Some("box".into());
    }
    let proposal = session
        .propose(request(vec![
            ProposalEdit::InsertSlide {
                reference: Some("slide".into()),
                after_slide_id: None,
                layout_part_path: None,
            },
            box_edit,
            ProposalEdit::ReplaceText {
                story_id: "@box".into(),
                start: 0,
                end: 0,
                text: "Filled".into(),
                style: None,
            },
            ProposalEdit::FormatText {
                story_id: "@box".into(),
                start: 0,
                end: 6,
                patch: pptx_edit::TextStylePatch {
                    bold: Some(true),
                    ..Default::default()
                },
            },
            ProposalEdit::SetParagraphAlignment {
                story_id: "@box".into(),
                start: 0,
                end: 6,
                alignment: Some("ctr".into()),
            },
            ProposalEdit::SetShapeFill {
                slide_id: "@slide".into(),
                shape_id: "@box".into(),
                color: Some("#123456".into()),
            },
        ]))
        .unwrap();
    session
        .insert_slide(&EditCtx::local("human"), 0, None)
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_ne!(preview.proposal.refs["box"], proposal.refs["box"]);
    assert!(preview.proposal.stale_targets.is_empty());
    let shape = preview.snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == preview.proposal.refs["box"])
        .unwrap();
    assert_eq!(shape.text_stories[0].plain_text(), "Filled");
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
}

#[test]
fn shape_staleness_preserves_user_edits_and_missing_targets_error_under_force() {
    let session = DeckSession::open(SHAPE_DECK, 532).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    let shape = &slide.shapes[0];
    let removal = session
        .propose(request(vec![shape_edit(
            "removeShape",
            &slide.id,
            &shape.id,
        )]))
        .unwrap();
    let addition = session
        .propose(request(vec![shape_edit("addTextBox", &slide.id, "")]))
        .unwrap();
    session
        .set_shape_fill(
            &EditCtx::local("human"),
            &slide.id,
            &shape.id,
            Some("#112233"),
        )
        .unwrap();
    assert!(matches!(
        session.accept_proposal(&removal.id, false),
        Err(ProposalError::Stale(_))
    ));
    assert!(
        session
            .proposals()
            .unwrap()
            .iter()
            .find(|p| p.id == addition.id)
            .unwrap()
            .stale_targets
            .is_empty()
    );
    session
        .remove_shape(&EditCtx::local("human"), &slide.id, &shape.id)
        .unwrap();
    assert!(matches!(
        session.accept_proposal(&removal.id, true),
        Err(ProposalError::Edit(pptx_edit::EditError::ShapeNotFound(_)))
    ));
    session
        .delete_slide(&EditCtx::local("human"), &slide.id)
        .unwrap();
    assert!(matches!(
        session.accept_proposal(&addition.id, false),
        Err(ProposalError::Stale(_))
    ));
    assert!(matches!(
        session.accept_proposal(&addition.id, true),
        Err(ProposalError::Edit(pptx_edit::EditError::SlideNotFound(_)))
    ));
}

#[test]
fn proposal_picture_cap_counts_decoded_bytes_and_the_total() {
    let session = DeckSession::open(SHAPE_DECK, 533).unwrap();
    let before = session.snapshot().unwrap();
    let picture = |size| ProposalEdit::AddPicture {
        reference: None,
        slide_id: before.slides[0].id.clone(),
        draft: pptx_edit::PictureDraft {
            name: "Picture".into(),
            rect: ShapeRect {
                x: 0,
                y: 0,
                width: 100000,
                height: 100000,
            },
            content_type: "image/png".into(),
            media_bytes: vec![0; size],
        },
    };
    for edits in [
        vec![picture(8 * 1024 * 1024 + 1)],
        vec![picture(4 * 1024 * 1024), picture(4 * 1024 * 1024 + 1)],
    ] {
        assert!(matches!(
            session.propose(request(edits)),
            Err(ProposalError::PictureMediaTooLarge { .. })
        ));
        assert!(session.proposals().unwrap().is_empty());
        assert_eq!(session.snapshot().unwrap(), before);
    }
    let wire = serde_json::to_string(&picture(8 * 1024 * 1024)).unwrap();
    let edit = serde_json::from_str(&wire).unwrap();
    assert!(session.propose(request(vec![edit])).is_ok());
}

#[test]
fn shape_diff_reports_geometry_appearance_and_deleted_shapes() {
    use pptx_edit::ProposalShapeChangeKind::{Changed, Deleted, Moved, Resized};
    for (kind, expected) in [
        ("move", Moved),
        ("resize", Resized),
        ("fill", Changed),
        ("remove", Deleted),
    ] {
        let session = DeckSession::open(SHAPE_DECK, 534).unwrap();
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let shape = &slide.shapes[0];
        let edit = match kind {
            "fill" => ProposalEdit::SetShapeFill {
                slide_id: slide.id.clone(),
                shape_id: shape.id.clone(),
                color: Some("#123456".into()),
            },
            "remove" => shape_edit("removeShape", &slide.id, &shape.id),
            _ => ProposalEdit::SetShapeRect {
                slide_id: slide.id.clone(),
                shape_id: shape.id.clone(),
                rect: ShapeRect {
                    x: shape.x + if kind == "move" { 100000 } else { 0 },
                    y: shape.y,
                    width: shape.width + if kind == "resize" { 100000 } else { 0 },
                    height: shape.height,
                },
            },
        };
        let proposal = session.propose(request(vec![edit])).unwrap();
        let diff = session
            .preview_proposal_diff_slide_by_id(&proposal.id, &slide.id)
            .unwrap();
        assert_eq!(diff.shape_changes[0].kind, expected);
        if kind == "remove" {
            assert_eq!(&diff.scope.slide.shapes[0], shape);
        }
    }
}

#[test]
fn created_shape_refs_support_geometry_stroke_adjustment_and_removal() {
    let session = DeckSession::open(SHAPE_DECK, 535).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0].id;
    let proposal = session
        .propose(request(vec![
            shape_edit("addShape", slide, ""),
            ProposalEdit::SetShapeRect {
                slide_id: slide.clone(),
                shape_id: "@created".into(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 2000000,
                    height: 1000000,
                },
            },
            ProposalEdit::SetShapeStroke {
                slide_id: slide.clone(),
                shape_id: "@created".into(),
                stroke: pptx_edit::ShapeStroke {
                    color: Some("#123456".into()),
                    width_pt: Some(2.0),
                },
            },
            ProposalEdit::SetShapeAdjust {
                slide_id: slide.clone(),
                shape_id: "@created".into(),
                adjustments: Default::default(),
            },
            ProposalEdit::RemoveShape {
                slide_id: slide.clone(),
                shape_id: "@created".into(),
            },
        ]))
        .unwrap();
    assert_eq!(
        session.preview_proposal(&proposal.id).unwrap().snapshot,
        before
    );
    assert!(
        !session
            .accept_proposal(&proposal.id, false)
            .unwrap()
            .applied
    );
    for edits in [
        vec![
            shape_edit("addShape", slide, ""),
            shape_edit("addTextBox", slide, ""),
        ],
        vec![ProposalEdit::RemoveShape {
            slide_id: slide.clone(),
            shape_id: "@unknown".into(),
        }],
    ] {
        assert!(matches!(
            session.propose(request(edits)),
            Err(ProposalError::Invalid(_))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
    }
}

#[test]
fn created_shape_refs_do_not_capture_unrelated_live_shapes() {
    let session = DeckSession::open(SHAPE_DECK, 536).unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let proposal = session
        .propose(request(vec![
            shape_edit("addTextBox", &slide, ""),
            ProposalEdit::ReplaceText {
                story_id: "@created".into(),
                start: 0,
                end: 0,
                text: "Draft".into(),
                style: None,
            },
            ProposalEdit::SetShapeFill {
                slide_id: slide.clone(),
                shape_id: "@created".into(),
                color: Some("#123456".into()),
            },
        ]))
        .unwrap();
    session
        .add_shape(
            &EditCtx::local("human"),
            &slide,
            &pptx_edit::PresetShapeDraft {
                name: "Human shape".into(),
                geometry: "rect".into(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 100000,
                    height: 100000,
                },
                fill: None,
            },
        )
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert!(preview.proposal.stale_targets.is_empty());
    assert_ne!(preview.proposal.refs["created"], proposal.refs["created"]);
    session.accept_proposal(&proposal.id, false).unwrap();
    let shapes = &session.snapshot().unwrap().slides[0].shapes;
    assert!(shapes.iter().any(|shape| shape.name == "Human shape"));
    assert!(shapes.iter().any(|shape| {
        shape
            .text_stories
            .iter()
            .any(|story| story.plain_text() == "Draft")
    }));
}

fn paragraph_edit(story_id: &str, paragraphs: serde_json::Value) -> ProposalEdit {
    serde_json::from_value(serde_json::json!({
        "type": "setStoryParagraphs", "storyId": story_id, "paragraphs": paragraphs,
    }))
    .unwrap()
}

fn paragraph_drafts() -> serde_json::Value {
    serde_json::json!([
        {"alignment": "ctr", "bullet": true, "runs": [
            {"text": "First ", "style": {"bold": true, "color": "#123456"}},
            {"text": "point", "style": {"italic": true, "fontSizePt": 20.0}},
        ]},
        {"alignment": "r", "bullet": false, "runs": [{"text": "Second 😀", "style": {}}]},
    ])
}

#[test]
fn paragraph_proposal_previews_accepts_rejects_and_undoes_one_step() {
    let session = DeckSession::open(SHAPE_DECK, 540).unwrap();
    session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
    let initial = session.snapshot().unwrap();
    let slide = &initial.slides[0];
    let shape = &slide.shapes[0];
    let story_id = &shape.text_stories[0].id;
    session
        .insert_text(
            &EditCtx::local("human"),
            story_id,
            0,
            "Human ",
            &Default::default(),
        )
        .unwrap();
    let before = session.snapshot().unwrap();
    let state = session.encode_state_as_update_v1();
    let edit = paragraph_edit(story_id, paragraph_drafts());
    let proposal = session.propose(request(vec![edit.clone()])).unwrap();
    assert_eq!(proposal.edits, vec![edit.clone()]);
    let wire = serde_json::to_value(&edit).unwrap();
    assert_eq!(wire["type"], "setStoryParagraphs");
    assert_eq!(wire["storyId"], story_id.as_str());
    let change = &proposal.changes[0];
    assert_eq!(change.before.as_ref(), Some(&before.slides[0].shapes[0]));
    assert_eq!(change.old_text, "Human alpha");
    assert_eq!(change.new_text, "First point\nSecond 😀");
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(
        change.after.as_ref(),
        Some(&preview.snapshot.slides[0].shapes[0])
    );
    let diff = session
        .preview_proposal_diff_slide(&proposal.id, 0)
        .unwrap();
    assert!(
        diff.text_changes
            .iter()
            .any(|change| { change.kind == pptx_edit::ProposalTextChangeKind::Insertion })
    );
    assert!(
        diff.text_changes
            .iter()
            .any(|change| { change.kind == pptx_edit::ProposalTextChangeKind::Deletion })
    );
    assert_eq!(
        diff.shape_changes[0].kind,
        pptx_edit::ProposalShapeChangeKind::Changed
    );
    assert_eq!(
        diff.scope.slide.shapes[0].text_stories[0].plain_text(),
        "Human alphaFirst point\nSecond 😀"
    );
    assert_eq!(
        diff.text_changes,
        vec![
            pptx_edit::ProposalTextChange {
                story_id: story_id.clone(),
                start: 0,
                end: 11,
                kind: pptx_edit::ProposalTextChangeKind::Deletion
            },
            pptx_edit::ProposalTextChange {
                story_id: story_id.clone(),
                start: 11,
                end: 22,
                kind: pptx_edit::ProposalTextChangeKind::Insertion
            },
            pptx_edit::ProposalTextChange {
                story_id: story_id.clone(),
                start: 23,
                end: 32,
                kind: pptx_edit::ProposalTextChangeKind::Insertion
            },
        ]
    );
    assert_eq!(session.snapshot().unwrap(), before);
    assert_eq!(session.encode_state_as_update_v1(), state);
    let accepted = session.accept_proposal(&proposal.id, false).unwrap();
    assert!(accepted.applied);
    assert_eq!(accepted.snapshot, preview.snapshot);
    let story = session.story(story_id).unwrap();
    assert_eq!(story.plain_text(), "First point\nSecond 😀");
    assert_eq!(story.paragraphs[0].alignment.as_deref(), Some("ctr"));
    assert_eq!(story.paragraphs[1].alignment.as_deref(), Some("r"));
    assert!(
        story.paragraphs[0]
            .bullet_json
            .as_ref()
            .unwrap()
            .contains("character")
    );
    assert!(
        story.paragraphs[1]
            .bullet_json
            .as_ref()
            .unwrap()
            .contains("none")
    );
    assert_eq!(story.paragraphs[0].runs[0].style.bold, Some(true));
    assert_eq!(story.paragraphs[0].runs[1].style.italic, Some(true));
    let reopened = DeckSession::open(&session.save().unwrap(), 541).unwrap();
    let saved = reopened.snapshot().unwrap().slides[0].shapes[0].text_stories[0].clone();
    assert_eq!(saved.plain_text(), story.plain_text());
    for (saved, expected) in saved.paragraphs.iter().zip(&story.paragraphs) {
        assert_eq!(saved.alignment, expected.alignment);
        assert_eq!(saved.bullet_json, expected.bullet_json);
        assert_eq!(saved.runs, expected.runs);
    }
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), initial);
    let rejected = session.propose(request(vec![edit])).unwrap();
    assert!(session.reject_proposal(&rejected.id));
    assert!(session.proposals().unwrap().is_empty());
    assert_eq!(session.snapshot().unwrap(), initial);
}

#[test]
fn paragraph_rewrite_diff_keeps_deleted_paragraph_text() {
    let session = DeckSession::open(SHAPE_DECK, 542).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    let paragraphs = serde_json::json!([
        {"runs": [{"text": "Keep", "style": {}}]},
        {"runs": [{"text": "Remove me", "style": {}}]},
        {"runs": [{"text": "Tail", "style": {}}]},
    ]);
    session
        .set_story_paragraphs(
            &EditCtx::local("human"),
            &story_id,
            &serde_json::from_value::<Vec<pptx_edit::TextParagraphDraft>>(paragraphs).unwrap(),
        )
        .unwrap();
    let proposal = session
        .propose(request(vec![paragraph_edit(
            &story_id,
            serde_json::json!([
                {"runs": [{"text": "Keep", "style": {}}]},
                {"runs": [{"text": "Tail", "style": {}}]},
            ]),
        )]))
        .unwrap();
    assert_eq!(proposal.changes[0].old_text, "Keep\nRemove me\nTail");
    assert_eq!(proposal.changes[0].new_text, "Keep\nTail");
    for diff in [
        session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap(),
        session
            .preview_proposal_diff_slide_by_id(&proposal.id, &proposal.changes[0].slide_id)
            .unwrap(),
    ] {
        let story = &diff.scope.slide.shapes[0].text_stories[0];
        let text = story.plain_text();
        assert_eq!(text, "Keep\nRemove me\nTail");
        let removed = text.find("Remove me").unwrap() as u32;
        assert!(diff.text_changes.iter().any(|change| {
            change.kind == pptx_edit::ProposalTextChangeKind::Deletion
                && change.start == removed
                && change.end == removed + 9
        }));
        assert!(
            !diff
                .text_changes
                .iter()
                .any(|change| { change.kind == pptx_edit::ProposalTextChangeKind::Insertion })
        );
    }
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.story(&story_id).unwrap().plain_text(), "Keep\nTail");
}

#[test]
fn paragraph_rewrite_diff_aligns_trailing_deletions() {
    assert_paragraph_rewrite_diff(
        &["Keep", "Remove me"],
        &["Keep"],
        "Keep\nRemove me",
        &[(pptx_edit::ProposalTextChangeKind::Deletion, 5, 14)],
    );
    assert_paragraph_rewrite_diff(
        &["Keep", "Remove me", "Remove too"],
        &["Keep"],
        "Keep\nRemove me\nRemove too",
        &[
            (pptx_edit::ProposalTextChangeKind::Deletion, 5, 14),
            (pptx_edit::ProposalTextChangeKind::Deletion, 15, 25),
        ],
    );
}

#[test]
fn paragraph_rewrite_diff_aligns_trailing_insertions() {
    assert_paragraph_rewrite_diff(
        &["Alpha"],
        &["Alpha", "Beta"],
        "Alpha\nBeta",
        &[(pptx_edit::ProposalTextChangeKind::Insertion, 6, 10)],
    );
    assert_paragraph_rewrite_diff(
        &["Alpha"],
        &["Alpha", "Beta", "Gamma"],
        "Alpha\nBeta\nGamma",
        &[
            (pptx_edit::ProposalTextChangeKind::Insertion, 6, 10),
            (pptx_edit::ProposalTextChangeKind::Insertion, 11, 16),
        ],
    );
}

#[test]
fn paragraph_rewrite_diff_aligns_middle_insertions() {
    assert_paragraph_rewrite_diff(
        &["Keep", "Tail"],
        &["Keep", "Middle", "Tail"],
        "Keep\nMiddle\nTail",
        &[(pptx_edit::ProposalTextChangeKind::Insertion, 5, 11)],
    );
    assert_paragraph_rewrite_diff(
        &["Keep 😀", "Tail"],
        &["Keep 😀", "Middle 😀", "Tail"],
        "Keep 😀\nMiddle 😀\nTail",
        &[(pptx_edit::ProposalTextChangeKind::Insertion, 8, 17)],
    );
}

#[test]
fn paragraph_rewrite_diff_edits_before_deleted_paragraphs() {
    use pptx_edit::ProposalTextChangeKind::{Deletion, Insertion};
    for tail in [false, true] {
        let mut old = serde_json::json!([
            {"alignment": "l", "bullet": false, "runs": [{"text": "Hello old", "style": {}}]},
            {"alignment": "r", "bullet": true, "runs": [{"text": "Remove me", "style": {}}]},
        ]);
        let mut new = serde_json::json!([
            {"alignment": "ctr", "bullet": true, "runs": [{"text": "Hello new", "style": {}}]},
        ]);
        if tail {
            let paragraph = serde_json::json!({"alignment": "l", "bullet": false, "runs": [{"text": "Tail", "style": {}}]});
            old.as_array_mut().unwrap().push(paragraph.clone());
            new.as_array_mut().unwrap().push(paragraph);
        }
        let story = assert_paragraph_draft_diff(
            old,
            new,
            if tail {
                "Hello oldnew\nRemove me\nTail"
            } else {
                "Hello oldnew\nRemove me"
            },
            &[(Deletion, 6, 9), (Insertion, 9, 12), (Deletion, 13, 22)],
        );
        assert_eq!(story.paragraphs[0].alignment.as_deref(), Some("ctr"));
        assert_eq!(story.paragraphs[1].alignment.as_deref(), Some("r"));
        for paragraph in &story.paragraphs[..2] {
            assert!(
                paragraph
                    .bullet_json
                    .as_ref()
                    .unwrap()
                    .contains("character")
            );
        }
        if tail {
            assert_eq!(story.paragraphs[2].alignment.as_deref(), Some("l"));
            assert!(
                story.paragraphs[2]
                    .bullet_json
                    .as_ref()
                    .unwrap()
                    .contains("none")
            );
        }
    }
}

#[test]
fn paragraph_rewrite_diff_inserts_before_rewritten_tail() {
    use pptx_edit::ProposalTextChangeKind::{Deletion, Insertion};
    assert_paragraph_rewrite_diff(
        &["Keep", "OldTail"],
        &["Keep", "Middle", "NewTail"],
        "Keep\nMiddle\nOldTailNewTail",
        &[(Insertion, 5, 11), (Deletion, 12, 19), (Insertion, 19, 26)],
    );
}

#[test]
fn paragraph_rewrite_diff_deletes_middle_while_editing_next() {
    use pptx_edit::ProposalTextChangeKind::{Deletion, Insertion};
    assert_paragraph_rewrite_diff(
        &["Keep", "Remove me", "Tail old"],
        &["Keep", "Tail new"],
        "Keep\nRemove me\nTail oldnew",
        &[(Deletion, 5, 14), (Deletion, 20, 23), (Insertion, 23, 26)],
    );
}

#[test]
fn paragraph_rewrite_diff_inserts_before_restyled_tail() {
    use pptx_edit::ProposalTextChangeKind::{Deletion, Insertion};
    let story = assert_paragraph_draft_diff(
        serde_json::json!([
            {"runs": [{"text": "Keep", "style": {}}]},
            {"runs": [{"text": "Tail", "style": {}}]},
        ]),
        serde_json::json!([
            {"runs": [{"text": "Keep", "style": {}}]},
            {"alignment": "ctr", "bullet": true, "runs": [{"text": "Middle", "style": {}}]},
            {"alignment": "r", "bullet": false, "runs": [{"text": "Tail", "style": {"bold": true}}]},
        ]),
        "Keep\nMiddle\nTailTail",
        &[(Insertion, 5, 11), (Deletion, 12, 16), (Insertion, 16, 20)],
    );
    assert_eq!(story.paragraphs[1].alignment.as_deref(), Some("ctr"));
    assert!(
        story.paragraphs[1]
            .bullet_json
            .as_ref()
            .unwrap()
            .contains("character")
    );
    assert_eq!(story.paragraphs[2].alignment.as_deref(), Some("r"));
    assert!(
        story.paragraphs[2]
            .bullet_json
            .as_ref()
            .unwrap()
            .contains("none")
    );
    assert_eq!(story.paragraphs[2].runs[1].style.bold, Some(true));
}

#[test]
fn paragraph_rewrite_diff_inserts_while_restyling_a_word() {
    use pptx_edit::ProposalTextChangeKind::{Deletion, Insertion};
    let story = assert_paragraph_draft_diff(
        serde_json::json!([
            {"runs": [{"text": "Keep", "style": {}}]},
            {"runs": [{"text": "Tail word", "style": {}}]},
        ]),
        serde_json::json!([
            {"runs": [{"text": "Keep", "style": {}}]},
            {"runs": [{"text": "Middle", "style": {}}]},
            {"runs": [{"text": "Tail ", "style": {}}, {"text": "word", "style": {"bold": true}}]},
        ]),
        "Keep\nMiddle\nTail wordword",
        &[(Insertion, 5, 11), (Deletion, 17, 21), (Insertion, 21, 25)],
    );
    assert_eq!(story.paragraphs[2].runs[2].style.bold, Some(true));
}

#[test]
fn paragraph_rewrite_diff_pairs_every_rewritten_paragraph_in_order() {
    use pptx_edit::ProposalTextChangeKind::{Deletion, Insertion};
    assert_paragraph_rewrite_diff(
        &["Old first", "Old second", "Old third"],
        &["New first", "New second", "New third"],
        "OldNew first\nOldNew second\nOldNew third",
        &[
            (Deletion, 0, 3),
            (Insertion, 3, 6),
            (Deletion, 13, 16),
            (Insertion, 16, 19),
            (Deletion, 27, 30),
            (Insertion, 30, 33),
        ],
    );
    assert_paragraph_rewrite_diff(
        &["Alpha", "Beta", "Gamma"],
        &["One", "Two", "Three"],
        "AlphaOne\nBetaTwo\nGammaThree",
        &[
            (Deletion, 0, 5),
            (Insertion, 5, 8),
            (Deletion, 9, 13),
            (Insertion, 13, 16),
            (Deletion, 17, 22),
            (Insertion, 22, 27),
        ],
    );
}

fn assert_paragraph_rewrite_diff(
    old: &[&str],
    new: &[&str],
    expected_text: &str,
    expected_changes: &[(pptx_edit::ProposalTextChangeKind, u32, u32)],
) {
    let drafts = |texts: &[&str]| {
        serde_json::Value::Array(
            texts
                .iter()
                .map(|text| serde_json::json!({"runs": [{"text": text, "style": {}}]}))
                .collect(),
        )
    };
    assert_paragraph_draft_diff(drafts(old), drafts(new), expected_text, expected_changes);
}

fn assert_paragraph_draft_diff(
    old: serde_json::Value,
    new: serde_json::Value,
    expected_text: &str,
    expected_changes: &[(pptx_edit::ProposalTextChangeKind, u32, u32)],
) -> pptx_edit::StorySnapshot {
    let session = DeckSession::open(SHAPE_DECK, 547).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    let new_drafts =
        serde_json::from_value::<Vec<pptx_edit::TextParagraphDraft>>(new.clone()).unwrap();
    let new_text = new_drafts
        .iter()
        .map(|p| {
            p.runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n");
    session
        .set_story_paragraphs(
            &EditCtx::local("human"),
            &story_id,
            &serde_json::from_value::<Vec<pptx_edit::TextParagraphDraft>>(old).unwrap(),
        )
        .unwrap();
    let before = session.snapshot().unwrap();
    let proposal = session
        .propose(request(vec![paragraph_edit(&story_id, new)]))
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(
        preview.snapshot.slides[0].shapes[0].text_stories[0].plain_text(),
        new_text
    );
    let deck_diff = session.preview_proposal_diff(&proposal.id).unwrap();
    for diff in [
        session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap(),
        session
            .preview_proposal_diff_slide_by_id(&proposal.id, &before.slides[0].id)
            .unwrap(),
    ] {
        let story = &diff.scope.slide.shapes[0].text_stories[0];
        assert_eq!(story.plain_text(), expected_text);
        assert_eq!(
            story
                .paragraphs
                .iter()
                .map(|p| p
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>())
                .collect::<Vec<_>>(),
            expected_text.split('\n').collect::<Vec<_>>()
        );
        assert_eq!(
            story.length,
            expected_text.encode_utf16().count() as u32 + 1
        );
        assert_eq!(diff.snapshot.slides[0], diff.scope.slide);
        assert_eq!(diff.snapshot.slides[0], deck_diff.snapshot.slides[0]);
        assert_eq!(diff.text_changes, deck_diff.text_changes);
        assert_eq!(
            diff.text_changes,
            expected_changes
                .iter()
                .map(|&(kind, start, end)| pptx_edit::ProposalTextChange {
                    story_id: story_id.clone(),
                    start,
                    end,
                    kind,
                })
                .collect::<Vec<_>>()
        );
        let mut offset = 0;
        for paragraph in &story.paragraphs {
            for run in &paragraph.runs {
                let end = offset + run.text.encode_utf16().count() as u32;
                let marked = expected_changes
                    .iter()
                    .find(|&&(_, start, stop)| start <= offset && end <= stop);
                match marked.map(|&(kind, _, _)| kind) {
                    Some(pptx_edit::ProposalTextChangeKind::Deletion) => {
                        assert_eq!(run.style.color.as_deref(), Some("#b91c1c"));
                    }
                    Some(pptx_edit::ProposalTextChangeKind::Insertion) => {
                        assert_eq!(run.style.color.as_deref(), Some("#166534"));
                    }
                    None => {
                        assert_ne!(run.style.color.as_deref(), Some("#b91c1c"));
                        assert_ne!(run.style.color.as_deref(), Some("#166534"));
                    }
                }
                offset = end;
            }
            offset += 1;
        }
    }
    assert_eq!(session.snapshot().unwrap(), before);
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    deck_diff.snapshot.slides[0].shapes[0].text_stories[0].clone()
}

#[test]
fn paragraph_properties_only_still_mark_the_shape_changed() {
    for property in [
        serde_json::json!({"bullet": true}),
        serde_json::json!({"alignment": "ctr"}),
    ] {
        let session = DeckSession::open(SHAPE_DECK, 543).unwrap();
        let snapshot = session.snapshot().unwrap();
        let shape = &snapshot.slides[0].shapes[0];
        let mut draft = property;
        draft["runs"] = serde_json::json!([{"text": "alpha", "style": {}}]);
        let proposal = session
            .propose(request(vec![paragraph_edit(
                &shape.text_stories[0].id,
                serde_json::json!([draft]),
            )]))
            .unwrap();
        assert_eq!(proposal.changes[0].old_text, proposal.changes[0].new_text);
        assert_ne!(proposal.changes[0].before, proposal.changes[0].after);
        let diff = session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert!(diff.text_changes.is_empty());
        assert_eq!(diff.shape_changes[0].shape_id, shape.id);
        assert_eq!(
            diff.shape_changes[0].kind,
            pptx_edit::ProposalShapeChangeKind::Changed
        );
        assert_eq!(
            session.proposal_slide_strip(&proposal.id).unwrap()[0].status,
            pptx_edit::ProposalStripStatus::Changed
        );
    }
}

#[test]
fn paragraph_proposal_resolves_created_shape_story_and_accepts_atomically() {
    let session = DeckSession::open(SHAPE_DECK, 544).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    let proposal = session
        .propose(request(vec![
            shape_edit("addTextBox", &slide.id, ""),
            paragraph_edit("@created", paragraph_drafts()),
        ]))
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    let shape = preview.snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == proposal.refs["created"])
        .unwrap();
    assert_eq!(shape.text_stories[0].plain_text(), "First point\nSecond 😀");
    assert_eq!(proposal.changes[0].old_text, "");
    assert_eq!(proposal.changes[0].new_text, "First point\nSecond 😀");
    assert!(proposal.changes[0].before.is_none());
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
}

#[test]
fn paragraph_proposal_stales_on_shape_edits_and_force_requires_a_live_story() {
    let session = DeckSession::open(SHAPE_DECK, 545).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    let story_id = &shape.text_stories[0].id;
    let proposal = session
        .propose(request(vec![paragraph_edit(story_id, paragraph_drafts())]))
        .unwrap();
    session
        .set_shape_fill(
            &EditCtx::local("human"),
            &slide.id,
            &shape.id,
            Some("#112233"),
        )
        .unwrap();
    assert_eq!(
        session.proposals().unwrap()[0].stale_targets.as_slice(),
        std::slice::from_ref(&shape.id)
    );
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
    session.preview_proposal(&proposal.id).unwrap();
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
    session.accept_proposal(&proposal.id, true).unwrap();
    let proposal = session
        .propose(request(vec![paragraph_edit(story_id, paragraph_drafts())]))
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("human"),
            story_id,
            0,
            "User ",
            &Default::default(),
        )
        .unwrap();
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
    session
        .remove_shape(&EditCtx::local("human"), &slide.id, &shape.id)
        .unwrap();
    assert!(matches!(
        session.accept_proposal(&proposal.id, true),
        Err(ProposalError::Edit(pptx_edit::EditError::StoryNotFound(_)))
    ));
    assert!(matches!(
        session.propose(request(vec![paragraph_edit("missing", paragraph_drafts())])),
        Err(ProposalError::Edit(pptx_edit::EditError::StoryNotFound(_)))
    ));
}

#[test]
fn invalid_paragraph_drafts_refuse_the_entire_proposal_without_history() {
    let session = DeckSession::open(SHAPE_DECK, 546).unwrap();
    let before = session.snapshot().unwrap();
    let story_id = &before.slides[0].shapes[0].text_stories[0].id;
    let draft = serde_json::json!({"runs": [{"text": "text", "style": {}}]});
    for paragraphs in [
        serde_json::json!([]),
        serde_json::json!(vec![draft.clone(); 101]),
        serde_json::json!([{"runs": vec![serde_json::json!({"text": "x", "style": {}}); 65]}]),
        serde_json::json!([{"alignment": "bad", "runs": []}]),
        serde_json::json!([{"runs": [{"text": "bad\u{0}", "style": {}}]}]),
        serde_json::json!([{"runs": [{"text": "bad\r", "style": {}}]}]),
        serde_json::json!([{"runs": [{"text": "x", "style": {"color": "bad"}}]}]),
        serde_json::json!([{"runs": [{"text": "x", "style": {"fontSizePt": 0}}]}]),
        serde_json::json!([{"runs": [{"text": "😀".repeat(8001), "style": {}}]}]),
    ] {
        let state = session.encode_state_as_update_v1();
        assert!(matches!(
            session.propose(request(vec![
                shape_edit("addTextBox", &before.slides[0].id, ""),
                paragraph_edit(story_id, paragraphs),
            ])),
            Err(ProposalError::Edit(pptx_edit::EditError::InvalidText(_)))
        ));
        assert_eq!(session.encode_state_as_update_v1(), state);
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.proposals().unwrap().is_empty());
        assert!(!session.can_undo());
    }
    assert!(
        serde_json::from_value::<ProposalEdit>(serde_json::json!({
            "type": "setStoryParagraphs", "storyId": story_id,
            "paragraphs": [{"runs": [], "indent": 1}],
        }))
        .is_err()
    );
    let proposal = session
        .propose(request(vec![paragraph_edit(
            story_id,
            serde_json::json!([{"runs": []}]),
        )]))
        .unwrap();
    assert_eq!(proposal.id, "p1");
}

const SLIDE_PROPS_DECK: &[u8] = include_bytes!("fixtures/proposal-slide-props.pptx");
const PROPOSED_LAYOUT: &str = "ppt/slideLayouts/slideLayout2.xml";

fn slide_property_edit(kind: &str, slide: &str, shape: &str) -> ProposalEdit {
    serde_json::from_value(match kind {
        "setSlideLayout" => serde_json::json!({
            "type": kind, "slideId": slide, "layoutPartPath": PROPOSED_LAYOUT,
            "mapping": [[shape, 8]],
        }),
        _ => serde_json::json!({
            "type": kind, "slideId": slide,
            "background": {"type": "gradient", "gradient": {"type": "linear", "angle": 90,
                "stops": [{"position": 0, "color": {"rgb": "123456"}}, {"position": 100000, "color": {"rgb": "ABCDEF"}}]}},
        }),
    })
    .unwrap()
}

#[test]
fn slide_property_variants_preview_accept_reject_and_undo() {
    for kind in ["setSlideLayout", "setSlideBackground"] {
        let session = DeckSession::open(SLIDE_PROPS_DECK, 550).unwrap();
        session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let edit = slide_property_edit(kind, &slide.id, &slide.shapes[0].id);
        let proposal = session.propose(request(vec![edit.clone()])).unwrap();
        assert_eq!(proposal.changes.len(), 1);
        assert_eq!(proposal.changes[0].shape_id, None);
        assert_eq!(
            proposal.changes[0].kind,
            Some(pptx_edit::ProposalChangeKind::Content)
        );
        assert_eq!(
            session.proposal_slide_strip(&proposal.id).unwrap()[0].status,
            pptx_edit::ProposalStripStatus::Changed
        );
        let preview = session.preview_proposal(&proposal.id).unwrap();
        assert_ne!(preview.snapshot, before);
        let by_id = session
            .preview_proposal_diff_slide_by_id(&proposal.id, &slide.id)
            .unwrap();
        let by_index = session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert_eq!(by_id.scope, by_index.scope);
        assert_eq!(by_id.scope.slide, preview.snapshot.slides[0]);
        if kind == "setSlideLayout" {
            assert_eq!(
                by_id.scope.slide.layout_part_path.as_deref(),
                Some(PROPOSED_LAYOUT)
            );
            assert_eq!(
                by_id.scope.slide.shapes[0]
                    .placeholder
                    .as_ref()
                    .unwrap()
                    .index,
                Some(4)
            );
        } else {
            assert_eq!(
                by_id.scope.slide.background.as_ref().unwrap().fill_type,
                "gradient"
            );
        }
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
        assert!(session.reject_proposal(&proposal.id));
        assert_eq!(session.snapshot().unwrap(), before);
        let proposal = session.propose(request(vec![edit])).unwrap();
        session.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }
}

#[test]
fn slide_property_refs_and_background_clear_are_atomic() {
    let session = DeckSession::open(SLIDE_PROPS_DECK, 551).unwrap();
    let before = session.snapshot().unwrap();
    let background =
        serde_json::from_value(serde_json::json!({"type": "solid", "color": {"rgb": "123456"}}))
            .unwrap();
    let proposal = session
        .propose(request(vec![
            ProposalEdit::InsertSlide {
                reference: Some("new".into()),
                after_slide_id: None,
                layout_part_path: Some("ppt/slideLayouts/slideLayout1.xml".into()),
            },
            ProposalEdit::SetSlideLayout {
                slide_id: "@new".into(),
                layout_part_path: PROPOSED_LAYOUT.into(),
                mapping: vec![],
            },
            ProposalEdit::SetSlideBackground {
                slide_id: "@new".into(),
                background: Some(background),
            },
        ]))
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(preview.snapshot.slides[0].id, proposal.refs["new"]);
    assert_eq!(
        preview.snapshot.slides[0]
            .background
            .as_ref()
            .unwrap()
            .fill_type,
        "solid"
    );
    assert!(session.proposals().unwrap()[0].stale_targets.is_empty());
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    let slide = &before.slides[0];
    session
        .set_slide_background(
            &EditCtx::local("human"),
            &slide.id,
            preview.snapshot.slides[0].background.as_ref(),
        )
        .unwrap();
    let proposal = session
        .propose(request(vec![ProposalEdit::SetSlideBackground {
            slide_id: slide.id.clone(),
            background: None,
        }]))
        .unwrap();
    let diff = session
        .preview_proposal_diff_slide(&proposal.id, 0)
        .unwrap();
    assert!(diff.scope.slide.background.is_none());
    session.accept_proposal(&proposal.id, false).unwrap();
    assert!(session.snapshot().unwrap().slides[0].background.is_none());
}

#[test]
fn slide_layout_fingerprint_and_effective_background_have_distinct_stale_rules() {
    for kind in ["setSlideLayout", "setSlideBackground"] {
        let session = DeckSession::open(SLIDE_PROPS_DECK, 552).unwrap();
        let slide = session.snapshot().unwrap().slides[0].clone();
        let proposal = session
            .propose(request(vec![slide_property_edit(
                kind,
                &slide.id,
                &slide.shapes[0].id,
            )]))
            .unwrap();
        session
            .set_slide_notes(&EditCtx::local("human"), &slide.id, "Human notes")
            .unwrap();
        assert_eq!(
            session.proposals().unwrap()[0].stale_targets.is_empty(),
            kind == "setSlideBackground"
        );
        let diff = session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert_eq!(
            diff.proposal.stale_targets.is_empty(),
            kind == "setSlideBackground"
        );
        if kind == "setSlideLayout" {
            assert!(matches!(
                session.accept_proposal(&proposal.id, false),
                Err(ProposalError::Stale(_))
            ));
        }
        session.accept_proposal(&proposal.id, true).unwrap();
    }
    for (layout, stale) in [
        ("ppt/slideLayouts/slideLayout3.xml", false),
        (PROPOSED_LAYOUT, true),
    ] {
        let session = DeckSession::open(SLIDE_PROPS_DECK, 553).unwrap();
        let slide = session.snapshot().unwrap().slides[0].clone();
        let proposal = session
            .propose(request(vec![slide_property_edit(
                "setSlideBackground",
                &slide.id,
                "",
            )]))
            .unwrap();
        session
            .set_slide_layout(&EditCtx::local("human"), &slide.id, layout, &[])
            .unwrap();
        assert_eq!(
            !session.proposals().unwrap()[0].stale_targets.is_empty(),
            stale
        );
        assert_eq!(
            !session
                .preview_proposal_diff_slide(&proposal.id, 0)
                .unwrap()
                .proposal
                .stale_targets
                .is_empty(),
            stale
        );
        if stale {
            assert!(matches!(
                session.accept_proposal(&proposal.id, false),
                Err(ProposalError::Stale(_))
            ));
        }
        session.accept_proposal(&proposal.id, true).unwrap();
    }
    let session = DeckSession::open(SLIDE_PROPS_DECK, 554).unwrap();
    let slide = session.snapshot().unwrap().slides[0].clone();
    let proposal = session
        .propose(request(vec![slide_property_edit(
            "setSlideBackground",
            &slide.id,
            "",
        )]))
        .unwrap();
    let fill =
        serde_json::from_value(serde_json::json!({"type": "solid", "color": {"rgb": "FFFFFF"}}))
            .unwrap();
    session
        .set_slide_background(&EditCtx::local("human"), &slide.id, Some(&fill))
        .unwrap();
    assert_eq!(
        session.proposals().unwrap()[0].stale_targets,
        vec![slide.id]
    );
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
    session.accept_proposal(&proposal.id, true).unwrap();
}

#[test]
fn invalid_slide_properties_and_forced_missing_targets_are_typed_and_leave_state_intact() {
    let session = DeckSession::open(SLIDE_PROPS_DECK, 555).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    for background in [
        serde_json::json!({"type": "solid", "color": {"rgb": "bad"}}),
        serde_json::json!({"type": "gradient", "gradient": {"type": "linear", "stops": []}}),
        serde_json::json!({"type": "picture"}),
    ] {
        let edit = serde_json::from_value(serde_json::json!({"type": "setSlideBackground", "slideId": slide.id, "background": background})).unwrap();
        assert!(matches!(
            session.propose(request(vec![edit])),
            Err(ProposalError::Edit(pptx_edit::EditError::InvalidFill(_)))
        ));
    }
    assert!(matches!(
        session.propose(request(vec![ProposalEdit::SetSlideLayout {
            slide_id: slide.id.clone(),
            layout_part_path: "unknown.xml".into(),
            mapping: vec![],
        }])),
        Err(ProposalError::Edit(pptx_edit::EditError::InvalidState(_)))
    ));
    for kind in ["setSlideLayout", "setSlideBackground"] {
        assert!(matches!(
            session.propose(request(vec![slide_property_edit(kind, "missing", "")])),
            Err(ProposalError::Edit(pptx_edit::EditError::SlideNotFound(_)))
        ));
    }
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.proposals().unwrap().is_empty());
    assert!(!session.can_undo());
    for kind in ["setSlideLayout", "setSlideBackground"] {
        let session = DeckSession::open(SLIDE_PROPS_DECK, 556).unwrap();
        let slide = session.snapshot().unwrap().slides[0].clone();
        let proposal = session
            .propose(request(vec![slide_property_edit(
                kind,
                &slide.id,
                &slide.shapes[0].id,
            )]))
            .unwrap();
        session
            .delete_slide(&EditCtx::local("human"), &slide.id)
            .unwrap();
        let before = session.snapshot().unwrap();
        assert!(matches!(
            session.accept_proposal(&proposal.id, true),
            Err(ProposalError::Edit(pptx_edit::EditError::SlideNotFound(_)))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
        assert_eq!(session.proposals().unwrap().len(), 1);
    }
}

#[test]
fn slide_layout_stales_on_shape_work_but_background_does_not() {
    for kind in ["setSlideLayout", "setSlideBackground"] {
        let session = DeckSession::open(SLIDE_PROPS_DECK, 557).unwrap();
        let slide = session.snapshot().unwrap().slides[0].clone();
        let proposal = session
            .propose(request(vec![slide_property_edit(
                kind,
                &slide.id,
                &slide.shapes[0].id,
            )]))
            .unwrap();
        let story = &slide.shapes[0].text_stories[0].id;
        session
            .insert_text(
                &EditCtx::local("human"),
                story,
                0,
                "Human work ",
                &pptx_edit::TextStyle::default(),
            )
            .unwrap();
        let stale = session.proposals().unwrap()[0].stale_targets.clone();
        if kind == "setSlideLayout" {
            assert_eq!(stale, vec![slide.id]);
            assert!(matches!(
                session.accept_proposal(&proposal.id, false),
                Err(ProposalError::Stale(_))
            ));
        } else {
            assert!(stale.is_empty());
            session.accept_proposal(&proposal.id, false).unwrap();
            assert!(
                session
                    .story(story)
                    .unwrap()
                    .plain_text()
                    .starts_with("Human work ")
            );
        }
    }
}

#[test]
fn slide_property_wire_requires_background_and_valid_placeholder_mapping() {
    assert!(
        serde_json::from_value::<ProposalEdit>(serde_json::json!({
            "type": "setSlideBackground", "slideId": "slide",
        }))
        .is_err()
    );
    let session = DeckSession::open(SLIDE_PROPS_DECK, 558).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    for mapping in [
        vec![(slide.shapes[0].id.clone(), 999)],
        vec![(slide.shapes[1].id.clone(), 8)],
        vec![
            (slide.shapes[0].id.clone(), 8),
            (slide.shapes[0].id.clone(), 8),
        ],
    ] {
        assert!(matches!(
            session.propose(request(vec![ProposalEdit::SetSlideLayout {
                slide_id: slide.id.clone(),
                layout_part_path: PROPOSED_LAYOUT.into(),
                mapping,
            }])),
            Err(ProposalError::Edit(_))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.proposals().unwrap().is_empty());
        assert!(!session.can_undo());
    }
}

const TABLE_DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

fn table_proposal_source() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(TABLE_DECK).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    for attribute in [
        "gridSpan=\"2\"",
        "rowSpan=\"2\"",
        "hMerge=\"1\"",
        "vMerge=\"1\"",
    ] {
        xml = xml.replace(attribute, "");
    }
    *bytes = xml.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn table_proposal_edit(kind: &str, slide: &str, shape: &str) -> ProposalEdit {
    let mut value = serde_json::json!({
        "type": kind, "slideId": slide, "shapeId": shape, "index": 0,
    });
    if kind.starts_with("insert") {
        value["position"] = "after".into();
        value["cells"] = serde_json::json!(["New cell"]);
    }
    serde_json::from_value(value).unwrap()
}

#[test]
fn table_variants_preview_accept_reject_and_undo() {
    use pptx_edit::{ProposalChangeKind, ProposalShapeChangeKind, ProposalStripStatus};
    for kind in [
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
    ] {
        let session = DeckSession::open(&table_proposal_source(), 560).unwrap();
        session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
        let original = session.snapshot().unwrap();
        let slide_id = original.slides[0].id.clone();
        session
            .set_slide_notes(&EditCtx::local("human"), &slide_id, "Earlier work")
            .unwrap();
        let before = session.snapshot().unwrap();
        let shape_id = &before.slides[0].shapes[0].id;
        let edit = table_proposal_edit(kind, &slide_id, shape_id);
        let proposal = session.propose(request(vec![edit.clone()])).unwrap();
        let preview = session.preview_proposal(&proposal.id).unwrap();
        let change = &proposal.changes[0];
        assert_eq!(change.kind, Some(ProposalChangeKind::Content));
        assert_eq!(change.shape_id.as_ref(), Some(shape_id));
        let shape = &preview.snapshot.slides[0].shapes[0];
        let Some(pptx_parse::GraphicFrameData::Table(table)) = &shape.graphic else {
            panic!("expected table")
        };
        let rows = if kind == "insertTableRow" {
            4
        } else if kind == "deleteTableRow" {
            2
        } else {
            3
        };
        let columns = if kind == "insertTableColumn" {
            4
        } else if kind == "deleteTableColumn" {
            2
        } else {
            3
        };
        assert_eq!((table.rows.len(), table.grid.len()), (rows, columns));
        assert_eq!(shape.text_stories.len(), rows * columns);
        let indexed = session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        let by_id = session
            .preview_proposal_diff_slide_by_id(&proposal.id, &slide_id)
            .unwrap();
        assert_eq!(indexed, by_id);
        assert_eq!(by_id.scope.slide.shapes[0].graphic, shape.graphic);
        assert_eq!(
            by_id.shape_changes[0].kind,
            ProposalShapeChangeKind::Changed
        );
        assert_eq!(
            session.proposal_slide_strip(&proposal.id).unwrap()[0].status,
            ProposalStripStatus::Changed
        );
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.reject_proposal(&proposal.id));
        assert_eq!(session.snapshot().unwrap(), before);
        let accepted = session.propose(request(vec![edit])).unwrap();
        session.accept_proposal(&accepted.id, false).unwrap();
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        let reopened = DeckSession::open(&session.save().unwrap(), 561).unwrap();
        let saved = reopened.snapshot().unwrap();
        let Some(pptx_parse::GraphicFrameData::Table(saved_table)) =
            &saved.slides[0].shapes[0].graphic
        else {
            panic!("expected saved table")
        };
        assert_eq!(
            (saved_table.rows.len(), saved_table.grid.len()),
            (rows, columns)
        );
        if kind.starts_with("insert") {
            assert!(
                saved.slides[0].shapes[0]
                    .text_stories
                    .iter()
                    .any(|story| story.plain_text() == "New cell")
            );
        }
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), original);
    }
}

#[test]
fn table_insert_then_replace_text_uses_the_new_cell_story() {
    let session = DeckSession::open(&table_proposal_source(), 562).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    let insert = table_proposal_edit("insertTableRow", &slide.id, &slide.shapes[0].id);
    let draft = session.propose(request(vec![insert.clone()])).unwrap();
    let preview = session.preview_proposal(&draft.id).unwrap();
    let story_id = preview.snapshot.slides[0].shapes[0].text_stories[3]
        .id
        .clone();
    assert!(session.reject_proposal(&draft.id));
    let proposal = session
        .propose(request(vec![
            insert,
            ProposalEdit::ReplaceText {
                story_id: story_id.clone(),
                start: 0,
                end: 8,
                text: "Reviewed cell".into(),
                style: None,
            },
        ]))
        .unwrap();
    assert_eq!(proposal.changes.len(), 1);
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(
        preview.snapshot.slides[0].shapes[0].text_stories[3].plain_text(),
        "Reviewed cell"
    );
    session
        .preview_proposal_diff_slide(&proposal.id, 0)
        .unwrap();
    assert_eq!(session.snapshot().unwrap(), before);
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(
        session.story(&story_id).unwrap().plain_text(),
        "Reviewed cell"
    );
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(!session.can_undo());
}

#[test]
fn table_proposals_are_stale_after_cell_or_structure_changes() {
    use pptx_edit::{TableInsertPosition, TextStyle};
    for kind in [
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
    ] {
        for structure in [false, true] {
            let session = DeckSession::open(&table_proposal_source(), 563).unwrap();
            let slide = session.snapshot().unwrap().slides[0].clone();
            let shape = &slide.shapes[0];
            let proposal = session
                .propose(request(vec![table_proposal_edit(
                    kind, &slide.id, &shape.id,
                )]))
                .unwrap();
            if structure {
                session
                    .insert_table_row(
                        &EditCtx::local("human"),
                        &slide.id,
                        &shape.id,
                        0,
                        TableInsertPosition::Before,
                        &[],
                    )
                    .unwrap();
            } else {
                session
                    .insert_text(
                        &EditCtx::local("human"),
                        &shape.text_stories[0].id,
                        0,
                        "Human ",
                        &TextStyle::default(),
                    )
                    .unwrap();
            }
            let before = session.snapshot().unwrap();
            assert_eq!(
                session.proposals().unwrap()[0].stale_targets,
                vec![shape.id.clone()]
            );
            assert!(matches!(
                session.accept_proposal(&proposal.id, false),
                Err(ProposalError::Stale(targets)) if targets == vec![shape.id.clone()]
            ));
            assert_eq!(session.snapshot().unwrap(), before);
        }
    }
}

#[test]
fn table_proposal_refusals_are_typed_and_leave_nothing_staged() {
    use pptx_edit::EditError;
    for kind in [
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
    ] {
        for (source, invalid_index) in [
            (table_proposal_source(), true),
            (TABLE_DECK.to_vec(), false),
            (SHAPE_DECK.to_vec(), false),
        ] {
            let session = DeckSession::open(&source, 564).unwrap();
            let before = session.snapshot().unwrap();
            let slide = &before.slides[0];
            let mut edit =
                serde_json::to_value(table_proposal_edit(kind, &slide.id, &slide.shapes[0].id))
                    .unwrap();
            if invalid_index {
                edit["index"] = 99.into();
            }
            let result = session.propose(request(vec![
                ProposalEdit::SetSlideNotes {
                    slide_id: slide.id.clone(),
                    text: "Must not stage".into(),
                },
                serde_json::from_value(edit).unwrap(),
            ]));
            assert!(matches!(
                result,
                Err(ProposalError::Edit(EditError::InvalidTable(_)))
            ));
            assert_eq!(session.snapshot().unwrap(), before);
            assert!(session.proposals().unwrap().is_empty());
            assert!(!session.can_undo());
        }
    }
    for kind in ["insertTableRow", "insertTableColumn"] {
        let session = DeckSession::open(&table_proposal_source(), 565).unwrap();
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        for cells in [vec!["cell"; 4], vec!["line\nbreak"], vec!["\u{0}"]] {
            let mut edit =
                serde_json::to_value(table_proposal_edit(kind, &slide.id, &slide.shapes[0].id))
                    .unwrap();
            edit["cells"] = serde_json::json!(cells);
            assert!(matches!(
                session.propose(request(vec![serde_json::from_value(edit).unwrap()])),
                Err(ProposalError::Edit(
                    EditError::InvalidTable(_) | EditError::InvalidText(_)
                ))
            ));
            assert_eq!(session.snapshot().unwrap(), before);
            assert!(session.proposals().unwrap().is_empty());
            assert!(!session.can_undo());
        }
    }
}

#[test]
fn table_proposal_wire_defaults_positions_widths_and_refs() {
    use pptx_edit::{TableColumnWidth, TableInsertPosition};
    for position in ["before", "after"] {
        for width in [None, Some("split"), Some("copy")] {
            let session = DeckSession::open(&table_proposal_source(), 566).unwrap();
            let before = session.snapshot().unwrap();
            let slide = &before.slides[0];
            let mut wire = serde_json::json!({
                "type": "insertTableColumn", "slideId": slide.id, "shapeId": slide.shapes[0].id,
                "index": 0, "position": position,
            });
            if let Some(width) = width {
                wire["width"] = width.into();
            }
            let edit: ProposalEdit = serde_json::from_value(wire).unwrap();
            let proposal = session.propose(request(vec![edit])).unwrap();
            let preview = session.preview_proposal(&proposal.id).unwrap();
            let direct = DeckSession::open(&table_proposal_source(), 566).unwrap();
            direct
                .insert_table_column_with_width(
                    &EditCtx::local("test"),
                    &slide.id,
                    &slide.shapes[0].id,
                    0,
                    if position == "before" {
                        TableInsertPosition::Before
                    } else {
                        TableInsertPosition::After
                    },
                    &[],
                    if width == Some("copy") {
                        TableColumnWidth::Copy
                    } else {
                        TableColumnWidth::Split
                    },
                )
                .unwrap();
            assert_eq!(preview.snapshot, direct.snapshot().unwrap());
        }
    }
    let session = DeckSession::open(SHAPE_DECK, 567).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    for kind in [
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
    ] {
        let edit = table_proposal_edit(kind, "@new", "@shape");
        assert!(matches!(session.propose(request(vec![
            ProposalEdit::DuplicateSlide { reference: Some("new".into()), slide_id: slide.id.clone(), after_slide_id: None },
            shape_edit("addShape", "@new", "unused"),
            edit,
        ])), Err(ProposalError::Invalid(message)) if message.contains("unknown ref")));
        let mut created = shape_edit("addShape", "@new", "unused");
        if let ProposalEdit::AddShape { reference, .. } = &mut created {
            *reference = Some("shape".into());
        }
        assert!(matches!(
            session.propose(request(vec![
                ProposalEdit::DuplicateSlide {
                    reference: Some("new".into()),
                    slide_id: slide.id.clone(),
                    after_slide_id: None
                },
                created,
                table_proposal_edit(kind, "@new", "@shape"),
            ])),
            Err(ProposalError::Edit(pptx_edit::EditError::InvalidTable(_)))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.proposals().unwrap().is_empty());
    }
}

#[test]
fn table_proposals_preserve_collaborative_refusals_and_undo_tracking() {
    for kind in [
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
    ] {
        let session = DeckSession::open(&table_proposal_source(), 568).unwrap();
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let edit = table_proposal_edit(kind, &slide.id, &slide.shapes[0].id);
        let proposal = session.propose(request(vec![edit.clone()])).unwrap();
        session.accept_proposal(&proposal.id, false).unwrap();
        assert!(session.can_undo());
        session.set_collaborative().unwrap();
        assert!(!session.can_undo());
        let before = session.snapshot().unwrap();
        assert!(matches!(
            session.propose(request(vec![edit])),
            Err(ProposalError::Edit(
                pptx_edit::EditError::CollaborativeStructureOperation
            ))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.proposals().unwrap().is_empty());
    }
    let session = DeckSession::open(&table_proposal_source(), 569).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    let proposal = session
        .propose(request(vec![table_proposal_edit(
            "insertTableRow",
            &slide.id,
            &slide.shapes[0].id,
        )]))
        .unwrap();
    session.preview_proposal(&proposal.id).unwrap();
    session.proposal_slide_strip(&proposal.id).unwrap();
    let story_id = slide.shapes[0].text_stories[0].id.clone();
    let text_proposal = session
        .propose(request(vec![ProposalEdit::ReplaceText {
            story_id: story_id.clone(),
            start: 0,
            end: 0,
            text: "Reviewed cell ".into(),
            style: None,
        }]))
        .unwrap();
    let text_preview = session.preview_proposal(&text_proposal.id).unwrap();
    session.set_collaborative().unwrap();
    assert!(matches!(
        session.preview_proposal(&proposal.id),
        Err(ProposalError::Edit(
            pptx_edit::EditError::CollaborativeStructureOperation
        ))
    ));
    assert!(matches!(
        session.proposal_slide_strip(&proposal.id),
        Err(ProposalError::Edit(
            pptx_edit::EditError::CollaborativeStructureOperation
        ))
    ));
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Edit(
            pptx_edit::EditError::CollaborativeStructureOperation
        ))
    ));
    assert_eq!(session.snapshot().unwrap(), before);
    assert_eq!(
        session.preview_proposal(&text_proposal.id).unwrap(),
        text_preview
    );
    session.proposal_slide_strip(&text_proposal.id).unwrap();
    assert!(
        session
            .accept_proposal(&text_proposal.id, false)
            .unwrap()
            .applied
    );
    assert_eq!(session.snapshot().unwrap(), text_preview.snapshot);
    assert!(
        session
            .story(&story_id)
            .unwrap()
            .plain_text()
            .starts_with("Reviewed cell ")
    );
}

#[test]
fn table_proposal_wire_refuses_malformed_edits() {
    let base = serde_json::json!({
        "type": "insertTableColumn", "slideId": "slide", "shapeId": "shape",
        "index": 0, "position": "before",
    });
    for (key, value) in [
        ("index", serde_json::json!(-1)),
        ("index", serde_json::json!(0.5)),
        ("index", serde_json::json!(4294967296u64)),
        ("position", serde_json::json!("middle")),
        ("cells", serde_json::json!([1])),
        ("width", serde_json::json!("stretch")),
        ("extra", serde_json::json!(true)),
    ] {
        let mut wire = base.clone();
        wire[key] = value;
        assert!(serde_json::from_value::<ProposalEdit>(wire).is_err());
    }
    let mut wire = base;
    wire.as_object_mut().unwrap().remove("position");
    assert!(serde_json::from_value::<ProposalEdit>(wire).is_err());
}

#[test]
fn accepted_table_proposals_keep_remote_conflict_and_history_tracking() {
    let bytes = table_proposal_source();
    let first = DeckSession::open(&bytes, 570).unwrap();
    let second = DeckSession::open(&bytes, 571).unwrap();
    let before = first.snapshot().unwrap();
    let slide = &before.slides[0];
    let shape_id = &slide.shapes[0].id;
    for (session, kind) in [(&first, "insertTableRow"), (&second, "insertTableColumn")] {
        let proposal = session
            .propose(request(vec![table_proposal_edit(
                kind, &slide.id, shape_id,
            )]))
            .unwrap();
        session.accept_proposal(&proposal.id, false).unwrap();
    }
    let a = first.encode_state_as_update_v1();
    let b = second.encode_state_as_update_v1();
    first.apply_update_v1(&b).unwrap();
    second.apply_update_v1(&a).unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    assert!(!first.can_undo());
    assert!(!second.can_undo());
    let discarded = [
        first.take_discarded_table_edits(),
        second.take_discarded_table_edits(),
    ];
    assert_eq!(
        discarded
            .iter()
            .filter(|ids| ids.as_slice() == std::slice::from_ref(shape_id))
            .count(),
        1
    );
}

fn reorder_edit(slide: &str, shape: &str, to: &str) -> ProposalEdit {
    serde_json::from_value(serde_json::json!({
        "type": "reorderShape", "slideId": slide, "shapeId": shape, "to": to,
    }))
    .unwrap()
}

#[test]
fn reorder_variants_preview_accept_reject_and_undo() {
    use pptx_edit::{ProposalChangeKind, ProposalShapeChangeKind, ProposalStripStatus};
    for to in ["front", "back", "forward", "backward"] {
        let session = DeckSession::open(SHAPE_DECK, 580).unwrap();
        session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
        let original = session.snapshot().unwrap();
        let slide_id = &original.slides[0].id;
        session
            .set_slide_notes(&EditCtx::local("human"), slide_id, "Earlier work")
            .unwrap();
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let from = usize::from(matches!(to, "back" | "backward"));
        let shape_id = &slide.shapes[from].id;
        let edit = reorder_edit(slide_id, shape_id, to);
        let proposal = session.propose(request(vec![edit.clone()])).unwrap();
        let preview = session.preview_proposal(&proposal.id).unwrap();
        let expected: Vec<_> = slide.shapes.iter().rev().map(|shape| &shape.id).collect();
        assert_eq!(
            preview.snapshot.slides[0]
                .shapes
                .iter()
                .map(|shape| &shape.id)
                .collect::<Vec<_>>(),
            expected
        );
        assert_eq!(proposal.changes[0].kind, Some(ProposalChangeKind::Content));
        assert_eq!(proposal.changes[0].shape_id.as_ref(), Some(shape_id));
        assert_eq!(proposal.changes[0].before_index, Some(Some(from)));
        assert_eq!(proposal.changes[0].after_index, Some(Some(1 - from)));
        let indexed = session
            .preview_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        let by_id = session
            .preview_proposal_diff_slide_by_id(&proposal.id, slide_id)
            .unwrap();
        assert_eq!(indexed, by_id);
        assert_eq!(by_id.scope.slide.shapes, preview.snapshot.slides[0].shapes);
        assert_eq!(by_id.shape_changes[0].shape_id, *shape_id);
        assert_eq!(
            by_id.shape_changes[0].kind,
            ProposalShapeChangeKind::Reordered
        );
        assert_eq!(
            session.proposal_slide_strip(&proposal.id).unwrap()[0].status,
            ProposalStripStatus::Changed
        );
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.reject_proposal(&proposal.id));
        assert_eq!(session.snapshot().unwrap(), before);
        let proposal = session.propose(request(vec![edit])).unwrap();
        assert!(
            session
                .accept_proposal(&proposal.id, false)
                .unwrap()
                .applied
        );
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        let reopened = DeckSession::open(&session.save().unwrap(), 581).unwrap();
        assert_eq!(
            reopened.snapshot().unwrap().slides[0]
                .shapes
                .iter()
                .map(|shape| shape.source_id)
                .collect::<Vec<_>>(),
            preview.snapshot.slides[0]
                .shapes
                .iter()
                .map(|shape| shape.source_id)
                .collect::<Vec<_>>()
        );
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), original);
    }
}

#[test]
fn reorder_proposals_resolve_created_slide_and_shape_refs() {
    let session = DeckSession::open(SHAPE_DECK, 582).unwrap();
    let before = session.snapshot().unwrap();
    let proposal = session
        .propose(request(vec![
            ProposalEdit::InsertSlide {
                reference: Some("slide".into()),
                after_slide_id: None,
                layout_part_path: None,
            },
            shape_edit("addShape", "@slide", ""),
            ProposalEdit::AddShape {
                reference: None,
                slide_id: "@slide".into(),
                draft: serde_json::from_value(serde_json::json!({
                    "name": "Other", "geometry": "rect",
                    "rect": {"x": 0, "y": 0, "width": 100000, "height": 100000},
                }))
                .unwrap(),
            },
            reorder_edit("@slide", "@created", "front"),
        ]))
        .unwrap();
    let preview = session.preview_proposal(&proposal.id).unwrap();
    assert_eq!(
        preview.snapshot.slides[0].shapes[1].id,
        proposal.refs["created"]
    );
    assert_eq!(session.snapshot().unwrap(), before);
    session.accept_proposal(&proposal.id, false).unwrap();
    assert_eq!(session.snapshot().unwrap(), preview.snapshot);
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
}

#[test]
fn reorder_proposals_stale_on_shape_capture_or_any_top_level_order_change() {
    for change_shape in [false, true] {
        let session = DeckSession::open(SHAPE_DECK, 583).unwrap();
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let target = &slide.shapes[0].id;
        let proposal = session
            .propose(request(vec![reorder_edit(&slide.id, target, "front")]))
            .unwrap();
        if change_shape {
            session
                .set_shape_fill(&EditCtx::local("human"), &slide.id, target, Some("#123456"))
                .unwrap();
        } else {
            session
                .send_to_back(&EditCtx::local("human"), &slide.id, &slide.shapes[1].id)
                .unwrap();
        }
        let changed = session.snapshot().unwrap();
        assert_eq!(
            session.proposals().unwrap()[0].stale_targets,
            std::slice::from_ref(target)
        );
        assert_eq!(
            session
                .preview_proposal_diff_slide(&proposal.id, 0)
                .unwrap()
                .proposal
                .stale_targets,
            std::slice::from_ref(target)
        );
        assert_eq!(
            session
                .preview_proposal_diff_slide_by_id(&proposal.id, &slide.id)
                .unwrap()
                .proposal
                .stale_targets,
            std::slice::from_ref(target)
        );
        assert!(matches!(
            session.accept_proposal(&proposal.id, false),
            Err(ProposalError::Stale(_))
        ));
        assert_eq!(session.snapshot().unwrap(), changed);
        assert_eq!(session.proposals().unwrap().len(), 1);
    }
}

#[test]
fn reorder_boundary_proposals_are_accepted_without_changes_or_undo() {
    for to in ["front", "back", "forward", "backward"] {
        let session = DeckSession::open(SHAPE_DECK, 584).unwrap();
        let before = session.snapshot().unwrap();
        let slide = &before.slides[0];
        let index = usize::from(matches!(to, "front" | "forward"));
        let proposal = session
            .propose(request(vec![reorder_edit(
                &slide.id,
                &slide.shapes[index].id,
                to,
            )]))
            .unwrap();
        assert_eq!(
            session.preview_proposal(&proposal.id).unwrap().snapshot,
            before
        );
        assert!(
            session
                .preview_proposal_diff_slide(&proposal.id, 0)
                .unwrap()
                .shape_changes
                .is_empty()
        );
        assert!(
            !session
                .accept_proposal(&proposal.id, false)
                .unwrap()
                .applied
        );
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }
}

#[test]
fn reorder_refusals_are_typed_and_atomic_including_grouped_children() {
    let session = DeckSession::open(SHAPE_DECK, 585).unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    for (slide_id, shape_id, missing_slide) in [
        ("missing", slide.shapes[0].id.as_str(), true),
        (slide.id.as_str(), "missing", false),
    ] {
        let error = session
            .propose(request(vec![
                ProposalEdit::SetSlideNotes {
                    slide_id: slide.id.clone(),
                    text: "Staged only".into(),
                },
                reorder_edit(slide_id, shape_id, "front"),
            ]))
            .unwrap_err();
        assert!(matches!(
            (&error, missing_slide),
            (
                ProposalError::Edit(pptx_edit::EditError::SlideNotFound(_)),
                true
            ) | (
                ProposalError::Edit(pptx_edit::EditError::ShapeNotFound(_)),
                false
            )
        ));
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(session.proposals().unwrap().is_empty());
        assert!(!session.can_undo());
    }
    let mut parts = ooxml_opc::unzip_parts(SHAPE_DECK).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    *bytes = xml.replacen("<p:sp>", "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"50\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp>", 1)
        .replacen("</p:spTree>", "</p:grpSp></p:spTree>", 1).into_bytes();
    let grouped = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 586).unwrap();
    let before = grouped.snapshot().unwrap();
    let slide = &before.slides[0];
    let child = &slide.shapes[0].children[0].id;
    assert!(matches!(
        grouped.propose(request(vec![reorder_edit(&slide.id, child, "front")])),
        Err(ProposalError::Edit(pptx_edit::EditError::ShapeNotFound(_)))
    ));
    assert_eq!(grouped.snapshot().unwrap(), before);
    assert!(grouped.proposals().unwrap().is_empty());
    assert!(!grouped.can_undo());
    for to in [
        serde_json::json!("up"),
        serde_json::Value::Null,
        serde_json::json!(42),
    ] {
        assert!(
            serde_json::from_value::<ProposalEdit>(serde_json::json!({
                "type": "reorderShape", "slideId": slide.id, "shapeId": child, "to": to,
            }))
            .is_err()
        );
    }
}

#[test]
fn reorder_directions_distinguish_extremes_from_single_steps() {
    for (to, from, target) in [
        ("front", 0, 2),
        ("forward", 0, 1),
        ("back", 2, 0),
        ("backward", 2, 1),
    ] {
        let session = DeckSession::open(SHAPE_DECK, 587).unwrap();
        let slide_id = session.snapshot().unwrap().slides[0].id.clone();
        let ProposalEdit::AddShape { draft, .. } = shape_edit("addShape", &slide_id, "") else {
            unreachable!()
        };
        session
            .add_shape(&EditCtx::local("human"), &slide_id, &draft)
            .unwrap();
        let before = session.snapshot().unwrap();
        let shape = &before.slides[0].shapes[from].id;
        let mut expected: Vec<_> = before.slides[0]
            .shapes
            .iter()
            .map(|shape| shape.id.clone())
            .collect();
        expected.remove(from);
        expected.insert(target, shape.clone());
        let proposal = session
            .propose(request(vec![reorder_edit(&slide_id, shape, to)]))
            .unwrap();
        let preview = session.preview_proposal(&proposal.id).unwrap();
        assert_eq!(
            preview.snapshot.slides[0]
                .shapes
                .iter()
                .map(|shape| shape.id.clone())
                .collect::<Vec<_>>(),
            expected
        );
        assert_eq!(session.snapshot().unwrap(), before);
        session.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(session.snapshot().unwrap(), preview.snapshot);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
    }
}

#[test]
fn reorder_stales_when_other_shapes_swap_without_moving_the_target() {
    let session = DeckSession::open(SHAPE_DECK, 588).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let ProposalEdit::AddShape { draft, .. } = shape_edit("addShape", &slide_id, "") else {
        unreachable!()
    };
    session
        .add_shape(&EditCtx::local("human"), &slide_id, &draft)
        .unwrap();
    let before = session.snapshot().unwrap();
    let slide = &before.slides[0];
    let target = &slide.shapes[0].id;
    let proposal = session
        .propose(request(vec![reorder_edit(&slide_id, target, "front")]))
        .unwrap();
    session
        .bring_to_front(&EditCtx::local("human"), &slide_id, &slide.shapes[1].id)
        .unwrap();
    assert_eq!(
        session.snapshot().unwrap().slides[0].shapes[0],
        slide.shapes[0]
    );
    assert_eq!(
        session.proposals().unwrap()[0].stale_targets,
        std::slice::from_ref(target)
    );
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
}
