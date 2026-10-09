use std::collections::{BTreeMap, BTreeSet, HashMap};

use serde::{Deserialize, Serialize};
use yrs::{ReadTxn, Transact};

use crate::staging::Adoption;
use crate::{
    DeckSession, DeckSnapshot, EditCtx, EditError, ShapeRect, ShapeSnapshot, ShapeStroke,
    TextStyle, TextStylePatch,
};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum ProposalEdit {
    ReplaceText {
        story_id: String,
        start: u32,
        end: u32,
        text: String,
        #[serde(default)]
        style: Option<TextStyle>,
    },
    FormatText {
        story_id: String,
        start: u32,
        end: u32,
        patch: TextStylePatch,
    },
    SetParagraphAlignment {
        story_id: String,
        start: u32,
        end: u32,
        alignment: Option<String>,
    },
    SetShapeRect {
        slide_id: String,
        shape_id: String,
        rect: ShapeRect,
    },
    SetShapeFill {
        slide_id: String,
        shape_id: String,
        color: Option<String>,
    },
    SetShapeStroke {
        slide_id: String,
        shape_id: String,
        stroke: ShapeStroke,
    },
    SetShapeAdjust {
        slide_id: String,
        shape_id: String,
        adjustments: BTreeMap<String, f64>,
    },
    SetSlideNotes {
        slide_id: String,
        text: String,
    },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ProposalRequest {
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposalEdit>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalChange {
    pub slide_id: String,
    pub shape_id: Option<String>,
    pub before: Option<ShapeSnapshot>,
    pub after: Option<ShapeSnapshot>,
    pub old_text: String,
    pub new_text: String,
}

impl ProposalChange {
    pub(crate) fn key(&self) -> String {
        self.shape_id
            .clone()
            .unwrap_or_else(|| self.slide_id.clone())
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Proposal {
    pub id: String,
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposalEdit>,
    pub changes: Vec<ProposalChange>,
    pub stale_targets: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalAcceptance {
    pub proposal_id: String,
    pub applied: bool,
    pub snapshot: DeckSnapshot,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalPreview {
    pub proposal: Proposal,
    pub snapshot: DeckSnapshot,
}

#[derive(Debug, thiserror::Error)]
pub enum ProposalError {
    #[error("no proposal {0}")]
    NotFound(String),
    #[error("stale proposal targets: {0:?}")]
    Stale(Vec<String>),
    #[error("invalid proposal: {0}")]
    Invalid(String),
    #[error(transparent)]
    Edit(#[from] EditError),
}

pub type ProposalResult<T> = Result<T, ProposalError>;

#[derive(Default)]
pub(crate) struct ProposalStore {
    next_id: u64,
    pending: Vec<Proposal>,
    /// Memoized previews of pending proposals, keyed on proposal id and the
    /// doc's epoch; a preview is a pure function of the proposal's edits and
    /// the doc state.
    pub(crate) previews: HashMap<String, (u64, ProposalPreview)>,
}

impl ProposalStore {
    pub(crate) fn counters(&self) -> (u64, usize) {
        (self.next_id, self.pending.len())
    }

    pub(crate) fn adopt_counter(&mut self, counter: u64) {
        self.next_id = counter;
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CompactProposalAcceptance {
    pub proposal_id: String,
    pub applied: bool,
    pub changed_targets: Vec<String>,
}

impl DeckSession {
    pub fn propose(&self, request: ProposalRequest) -> ProposalResult<Proposal> {
        if request.agent_id.trim().is_empty()
            || request.edits.is_empty()
            || request.edits.len() > 256
        {
            return Err(ProposalError::Invalid(
                "an agent and 1 to 256 edits are required".into(),
            ));
        }
        if self.proposals.borrow().pending.len() >= 64 {
            return Err(ProposalError::Invalid(
                "64 proposals are already pending".into(),
            ));
        }
        let before = self.snapshot()?;
        let (_, snapshot) = self.preview_edits(&before, &request.edits)?;
        let changes = changes_for(&before, &snapshot, &request.edits)?;
        let mut store = self.proposals.borrow_mut();
        store.next_id += 1;
        let proposal = Proposal {
            id: format!("p{}", store.next_id),
            agent_id: request.agent_id,
            note: request.note,
            edits: request.edits,
            changes,
            stale_targets: Vec::new(),
        };
        store.pending.push(proposal.clone());
        let epoch = self.epoch();
        store.previews.insert(
            proposal.id.clone(),
            (
                epoch,
                ProposalPreview {
                    proposal: proposal.clone(),
                    snapshot,
                },
            ),
        );
        Ok(proposal)
    }

    pub fn proposals(&self) -> ProposalResult<Vec<Proposal>> {
        let snapshot = self.snapshot()?;
        Ok(self
            .proposals
            .borrow()
            .pending
            .iter()
            .map(|proposal| {
                let mut proposal = proposal.clone();
                proposal.stale_targets = stale_targets(&snapshot, &proposal);
                proposal
            })
            .collect())
    }

    pub fn preview_proposal(&self, id: &str) -> ProposalResult<ProposalPreview> {
        let mut proposal = self.pending_proposal(id)?;
        let epoch = self.epoch();
        if let Some((cached_epoch, preview)) = self.proposals.borrow().previews.get(id)
            && *cached_epoch == epoch
        {
            return Ok(preview.clone());
        }
        self.proposals.borrow_mut().previews.remove(id);
        let before = self.snapshot()?;
        let (_, snapshot) = self.preview_edits(&before, &proposal.edits)?;
        proposal.stale_targets = stale_targets(&before, &proposal);
        proposal.changes = changes_for(&before, &snapshot, &proposal.edits)?;
        let preview = ProposalPreview { proposal, snapshot };
        self.proposals
            .borrow_mut()
            .previews
            .insert(id.to_owned(), (epoch, preview.clone()));
        Ok(preview)
    }

    pub fn proposal_preview_session(&self, id: &str) -> ProposalResult<DeckSession> {
        let proposal = self.pending_proposal(id)?;
        let before = self.snapshot()?;
        Ok(self.preview_edits(&before, &proposal.edits)?.0)
    }

    pub fn accept_proposal(&self, id: &str, force: bool) -> ProposalResult<ProposalAcceptance> {
        let (outcome, snapshot) = self.accept_proposal_inner(id, force)?;
        Ok(ProposalAcceptance {
            proposal_id: outcome.proposal_id,
            applied: outcome.applied,
            snapshot,
        })
    }

    pub(crate) fn accept_proposal_compact(
        &self,
        id: &str,
        force: bool,
    ) -> ProposalResult<CompactProposalAcceptance> {
        self.accept_proposal_inner(id, force)
            .map(|(outcome, _)| outcome)
    }

    fn accept_proposal_inner(
        &self,
        id: &str,
        force: bool,
    ) -> ProposalResult<(CompactProposalAcceptance, DeckSnapshot)> {
        let proposal = self.pending_proposal(id)?;
        let before = self.snapshot()?;
        let stale = stale_targets(&before, &proposal);
        if !force && !stale.is_empty() {
            return Err(ProposalError::Stale(stale));
        }
        let (preview, snapshot) = self.preview_edits(&before, &proposal.edits)?;
        let applied = before != snapshot;
        if applied {
            let (_, update) = preview.staged_update(&self.doc.transact().state_vector())?;
            self.adopt(&preview, update, Adoption::Proposal)?;
        }
        self.reject_proposal(id);
        let changed_targets = proposal
            .changes
            .iter()
            .filter(|change| {
                capture(&snapshot, &change.slide_id, change.shape_id.as_deref()).is_ok_and(
                    |after| {
                        capture(&before, &change.slide_id, change.shape_id.as_deref())
                            .is_ok_and(|before| after != before)
                    },
                )
            })
            .map(ProposalChange::key)
            .collect::<BTreeSet<_>>()
            .into_iter()
            .collect();
        Ok((
            CompactProposalAcceptance {
                proposal_id: id.to_owned(),
                applied,
                changed_targets,
            },
            snapshot,
        ))
    }

    pub fn reject_proposal(&self, id: &str) -> bool {
        let mut store = self.proposals.borrow_mut();
        store.previews.remove(id);
        let before = store.pending.len();
        store.pending.retain(|proposal| proposal.id != id);
        before != store.pending.len()
    }

    pub(crate) fn pending_proposal(&self, id: &str) -> ProposalResult<Proposal> {
        self.proposals
            .borrow()
            .pending
            .iter()
            .find(|proposal| proposal.id == id)
            .cloned()
            .ok_or_else(|| ProposalError::NotFound(id.to_owned()))
    }

    /// Applies `edits` to a hydrated clone of the live doc — never the live
    /// doc itself, so a preview produces no observable update events and
    /// leaves `encode_state_as_update_v1` byte-identical — validates it, and
    /// returns the scratch session together with the validated snapshot.
    fn preview_edits(
        &self,
        before: &DeckSnapshot,
        edits: &[ProposalEdit],
    ) -> ProposalResult<(DeckSession, DeckSnapshot)> {
        for edit in edits {
            let (slide_id, shape_id) = target(before, edit)?;
            check_target(before, &slide_id, shape_id.as_deref())?;
        }
        let preview = self.stage()?;
        for edit in edits {
            apply_edit(&preview, edit)?;
        }
        let snapshot = preview.validated_snapshot()?;
        Ok((preview, snapshot))
    }
}

pub(crate) fn apply_edit(session: &DeckSession, edit: &ProposalEdit) -> Result<(), EditError> {
    let context = EditCtx::local("proposal");
    match edit {
        ProposalEdit::ReplaceText {
            story_id,
            start,
            end,
            text,
            style,
        } => {
            let story = session.story(story_id)?;
            let inherited = inherited_style(&story, *start);
            session.delete_text(&context, story_id, *start, *end)?;
            session.insert_text(
                &context,
                story_id,
                *start,
                text,
                style.as_ref().unwrap_or(&inherited),
            )?;
        }
        ProposalEdit::FormatText {
            story_id,
            start,
            end,
            patch,
        } => {
            session.format_text(&context, story_id, *start, *end, patch)?;
        }
        ProposalEdit::SetParagraphAlignment {
            story_id,
            start,
            end,
            alignment,
        } => {
            session.set_paragraph_alignment(
                &context,
                story_id,
                *start,
                *end,
                alignment.as_deref(),
            )?;
        }
        ProposalEdit::SetShapeRect {
            slide_id,
            shape_id,
            rect,
        } => {
            session.set_shape_rect(&context, slide_id, shape_id, *rect)?;
        }
        ProposalEdit::SetShapeFill {
            slide_id,
            shape_id,
            color,
        } => {
            session.set_shape_fill(&context, slide_id, shape_id, color.as_deref())?;
        }
        ProposalEdit::SetShapeStroke {
            slide_id,
            shape_id,
            stroke,
        } => {
            session.set_shape_stroke(&context, slide_id, shape_id, stroke)?;
        }
        ProposalEdit::SetShapeAdjust {
            slide_id,
            shape_id,
            adjustments,
        } => {
            session.set_shape_adjust(&context, slide_id, shape_id, adjustments)?;
        }
        ProposalEdit::SetSlideNotes { slide_id, text } => {
            session.set_slide_notes(&context, slide_id, text)?
        }
    }
    Ok(())
}

pub(crate) fn inherited_style(story: &crate::StorySnapshot, at: u32) -> TextStyle {
    let mut offset = 0;
    let mut previous = TextStyle::default();
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            offset += run.text.encode_utf16().count() as u32;
            if at < offset {
                return run.style.clone();
            }
            previous = run.style.clone();
        }
        if at == offset {
            return previous;
        }
        offset += 1;
    }
    previous
}

fn find_shape<'a>(shapes: &'a [ShapeSnapshot], id: &str) -> Option<&'a ShapeSnapshot> {
    shapes.iter().find_map(|shape| {
        if shape.id == id {
            Some(shape)
        } else {
            find_shape(&shape.children, id)
        }
    })
}

fn find_story_shape<'a>(shapes: &'a [ShapeSnapshot], id: &str) -> Option<&'a ShapeSnapshot> {
    shapes.iter().find_map(|shape| {
        if shape.text_stories.iter().any(|story| story.id == id) {
            Some(shape)
        } else {
            find_story_shape(&shape.children, id)
        }
    })
}

pub(crate) fn shape_text(shape: &ShapeSnapshot) -> String {
    shape
        .text_stories
        .iter()
        .map(|story| story.plain_text())
        .collect::<Vec<_>>()
        .join("\n")
}

fn target(
    snapshot: &DeckSnapshot,
    edit: &ProposalEdit,
) -> Result<(String, Option<String>), EditError> {
    match edit {
        ProposalEdit::ReplaceText { story_id, .. }
        | ProposalEdit::FormatText { story_id, .. }
        | ProposalEdit::SetParagraphAlignment { story_id, .. } => snapshot
            .slides
            .iter()
            .find_map(|slide| {
                find_story_shape(&slide.shapes, story_id)
                    .map(|shape| (slide.id.clone(), Some(shape.id.clone())))
            })
            .ok_or_else(|| EditError::StoryNotFound(story_id.clone())),
        ProposalEdit::SetShapeRect {
            slide_id, shape_id, ..
        }
        | ProposalEdit::SetShapeFill {
            slide_id, shape_id, ..
        }
        | ProposalEdit::SetShapeStroke {
            slide_id, shape_id, ..
        }
        | ProposalEdit::SetShapeAdjust {
            slide_id, shape_id, ..
        } => Ok((slide_id.clone(), Some(shape_id.clone()))),
        ProposalEdit::SetSlideNotes { slide_id, .. } => Ok((slide_id.clone(), None)),
    }
}

fn check_target(
    snapshot: &DeckSnapshot,
    slide_id: &str,
    shape_id: Option<&str>,
) -> Result<(), EditError> {
    let slide = snapshot
        .slides
        .iter()
        .find(|slide| slide.id == slide_id)
        .ok_or_else(|| EditError::SlideNotFound(slide_id.to_owned()))?;
    if let Some(id) = shape_id {
        find_shape(&slide.shapes, id).ok_or_else(|| EditError::ShapeNotFound(id.to_owned()))?;
    }
    Ok(())
}

fn capture(
    snapshot: &DeckSnapshot,
    slide_id: &str,
    shape_id: Option<&str>,
) -> Result<(Option<ShapeSnapshot>, String), EditError> {
    let slide = snapshot
        .slides
        .iter()
        .find(|slide| slide.id == slide_id)
        .ok_or_else(|| EditError::SlideNotFound(slide_id.to_owned()))?;
    match shape_id {
        Some(id) => {
            let shape = find_shape(&slide.shapes, id)
                .ok_or_else(|| EditError::ShapeNotFound(id.to_owned()))?;
            Ok((Some(shape.clone()), shape_text(shape)))
        }
        None => Ok((None, slide.notes.clone())),
    }
}

fn changes_for(
    before: &DeckSnapshot,
    after: &DeckSnapshot,
    edits: &[ProposalEdit],
) -> Result<Vec<ProposalChange>, EditError> {
    let targets = edits
        .iter()
        .map(|edit| target(before, edit))
        .collect::<Result<BTreeSet<_>, _>>()?;
    targets
        .into_iter()
        .map(|(slide_id, shape_id)| {
            let (before, old_text) = capture(before, &slide_id, shape_id.as_deref())?;
            let (after, new_text) = capture(after, &slide_id, shape_id.as_deref())?;
            Ok(ProposalChange {
                slide_id,
                shape_id,
                before,
                after,
                old_text,
                new_text,
            })
        })
        .collect()
}

fn stale_targets(snapshot: &DeckSnapshot, proposal: &Proposal) -> Vec<String> {
    proposal
        .changes
        .iter()
        .filter(|change| {
            capture(snapshot, &change.slide_id, change.shape_id.as_deref())
                .map_or(true, |(shape, text)| {
                    shape != change.before || text != change.old_text
                })
        })
        .map(ProposalChange::key)
        .collect()
}

#[cfg(test)]
mod tests {
    use crate::{DeckSession, EditCtx, ProposalEdit, ProposalError, ProposalRequest};

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

    fn request(edit: ProposalEdit) -> ProposalRequest {
        ProposalRequest {
            agent_id: "agent".into(),
            note: None,
            edits: vec![edit],
        }
    }

    fn notes(slide_id: &str, text: &str) -> ProposalRequest {
        request(ProposalEdit::SetSlideNotes {
            slide_id: slide_id.into(),
            text: text.into(),
        })
    }

    #[test]
    fn previews_leave_the_cache_with_their_proposal() {
        let session = DeckSession::open(DECK, 97).unwrap();
        let slide = session.snapshot().unwrap().slides[0].clone();
        let cached = || session.proposals.borrow().previews.len();
        let missing = |id: &str| {
            matches!(
                session.preview_proposal(id),
                Err(ProposalError::NotFound(_))
            )
        };
        let accept = |id: &str| session.accept_proposal(id, false).unwrap().applied;

        let rejected = session.propose(notes(&slide.id, "Rejected")).unwrap();
        session.preview_proposal(&rejected.id).unwrap();
        assert!(session.reject_proposal(&rejected.id));
        assert!(missing(&rejected.id));
        let unchanged = session.propose(notes(&slide.id, &slide.notes)).unwrap();
        assert!(!accept(&unchanged.id));
        assert!(missing(&unchanged.id));
        let accepted = session.propose(notes(&slide.id, "Accepted")).unwrap();
        assert!(accept(&accepted.id));
        assert!(missing(&accepted.id));
        assert_eq!(cached(), 0);

        for cycle in 0..50 {
            let proposal = session
                .propose(notes(&slide.id, &format!("Cycle {cycle}")))
                .unwrap();
            session.preview_proposal(&proposal.id).unwrap();
            assert!(session.reject_proposal(&proposal.id));
        }
        assert_eq!(cached(), 0);

        let shape = &slide.shapes[0];
        let removed = session
            .propose(request(ProposalEdit::SetShapeFill {
                slide_id: slide.id.clone(),
                shape_id: shape.id.clone(),
                color: Some("#FF0000".into()),
            }))
            .unwrap();
        assert_eq!(cached(), 1);
        session
            .remove_shape(&EditCtx::local("human"), &slide.id, &shape.id)
            .unwrap();
        assert!(session.preview_proposal(&removed.id).is_err());
        assert_eq!(session.proposals().unwrap().len(), 1);
        assert_eq!(cached(), 0);
    }
}
