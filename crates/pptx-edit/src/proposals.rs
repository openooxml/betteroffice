use std::collections::{BTreeMap, BTreeSet, HashMap};

use serde::{Deserialize, Serialize};
use yrs::{ReadTxn, Transact};

use crate::staging::Adoption;
use crate::{
    DeckSession, DeckSnapshot, EditCtx, EditError, PictureDraft, PresetShapeDraft, ShapeDraft,
    ShapeRect, ShapeSnapshot, ShapeStroke, TextStyle, TextStylePatch,
};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum ProposalEdit {
    AddTextBox {
        #[serde(default, rename = "ref", skip_serializing_if = "Option::is_none")]
        reference: Option<String>,
        slide_id: String,
        draft: ShapeDraft,
    },
    AddShape {
        #[serde(default, rename = "ref", skip_serializing_if = "Option::is_none")]
        reference: Option<String>,
        slide_id: String,
        draft: PresetShapeDraft,
    },
    AddPicture {
        #[serde(default, rename = "ref", skip_serializing_if = "Option::is_none")]
        reference: Option<String>,
        slide_id: String,
        draft: PictureDraft,
    },
    RemoveShape {
        slide_id: String,
        shape_id: String,
    },
    InsertSlide {
        #[serde(default, rename = "ref", skip_serializing_if = "Option::is_none")]
        reference: Option<String>,
        #[serde(deserialize_with = "required_anchor")]
        after_slide_id: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        layout_part_path: Option<String>,
    },
    DuplicateSlide {
        #[serde(default, rename = "ref", skip_serializing_if = "Option::is_none")]
        reference: Option<String>,
        slide_id: String,
        #[serde(
            default,
            skip_serializing_if = "Option::is_none",
            deserialize_with = "present_anchor"
        )]
        after_slide_id: Option<Option<String>>,
    },
    DeleteSlide {
        slide_id: String,
    },
    MoveSlide {
        slide_id: String,
        #[serde(deserialize_with = "required_anchor")]
        after_slide_id: Option<String>,
    },
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

fn required_anchor<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Option<String>, D::Error> {
    Option::<String>::deserialize(deserializer)
}

fn present_anchor<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Option<Option<String>>, D::Error> {
    Option::<String>::deserialize(deserializer).map(Some)
}

fn present_index<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> Result<Option<Option<usize>>, D::Error> {
    Option::<usize>::deserialize(deserializer).map(Some)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ProposalChangeKind {
    Content,
    Inserted,
    Deleted,
    Moved,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ProposalStripStatus {
    Unchanged,
    Changed,
    Inserted,
    Deleted,
    Moved,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalStripEntry {
    pub slide_id: String,
    pub status: ProposalStripStatus,
    pub live_index: Option<usize>,
    pub preview_index: Option<usize>,
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
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<ProposalChangeKind>,
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        deserialize_with = "present_index"
    )]
    pub before_index: Option<Option<usize>>,
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        deserialize_with = "present_index"
    )]
    pub after_index: Option<Option<usize>>,
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
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub refs: BTreeMap<String, String>,
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
    #[error(
        "proposalPictureMediaTooLarge: picture media exceeds {limit_bytes} decoded bytes ({total_bytes} bytes)"
    )]
    PictureMediaTooLarge {
        total_bytes: usize,
        limit_bytes: usize,
    },
    #[error(transparent)]
    Edit(#[from] EditError),
}

pub type ProposalResult<T> = Result<T, ProposalError>;

#[derive(Default)]
pub(crate) struct ProposalStore {
    next_id: u64,
    pending: Vec<Proposal>,
    baselines: HashMap<String, ProposalBaseline>,
    /// Memoized previews of pending proposals, keyed on proposal id and the
    /// doc's epoch; a preview is a pure function of the proposal's edits and
    /// the doc state.
    pub(crate) previews: HashMap<String, (u64, ProposalPreview)>,
}

pub(crate) struct ProposalStage {
    pub session: DeckSession,
    pub snapshot: DeckSnapshot,
    pub refs: BTreeMap<String, String>,
    pub resolved: Vec<ProposalEdit>,
}

struct ProposalBaseline {
    snapshot: DeckSnapshot,
    edits: Vec<ProposalEdit>,
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
        let ProposalStage {
            snapshot,
            refs,
            resolved,
            ..
        } = self.preview_edits(&request.edits, false)?;
        let changes = changes_for(&before, &snapshot, &resolved)?;
        let mut store = self.proposals.borrow_mut();
        store.next_id += 1;
        let proposal = Proposal {
            id: format!("p{}", store.next_id),
            agent_id: request.agent_id,
            note: request.note,
            edits: request.edits,
            changes,
            stale_targets: Vec::new(),
            refs,
        };
        store.baselines.insert(
            proposal.id.clone(),
            ProposalBaseline {
                snapshot: before,
                edits: resolved,
            },
        );
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
                proposal.stale_targets = self.proposal_stale_targets(&snapshot, &proposal);
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
        let ProposalStage {
            snapshot,
            refs,
            resolved,
            ..
        } = self.preview_edits(&proposal.edits, true)?;
        proposal.stale_targets = self.proposal_stale_targets(&before, &proposal);
        proposal.changes = changes_for(&before, &snapshot, &resolved)?;
        proposal.refs = refs;
        let preview = ProposalPreview { proposal, snapshot };
        if let Some(pending) = self
            .proposals
            .borrow_mut()
            .pending
            .iter_mut()
            .find(|proposal| proposal.id == id)
        {
            pending.refs = preview.proposal.refs.clone();
            if pending.edits.iter().any(|edit| {
                matches!(
                    edit,
                    ProposalEdit::InsertSlide { .. }
                        | ProposalEdit::DuplicateSlide { .. }
                        | ProposalEdit::DeleteSlide { .. }
                        | ProposalEdit::MoveSlide { .. }
                        | ProposalEdit::AddTextBox { .. }
                        | ProposalEdit::AddShape { .. }
                        | ProposalEdit::AddPicture { .. }
                        | ProposalEdit::RemoveShape { .. }
                )
            }) {
                pending.changes = preview.proposal.changes.clone();
            }
        }
        self.proposals
            .borrow_mut()
            .previews
            .insert(id.to_owned(), (epoch, preview.clone()));
        Ok(preview)
    }

    pub fn proposal_preview_session(&self, id: &str) -> ProposalResult<DeckSession> {
        let proposal = self.preview_proposal(id)?.proposal;
        Ok(self.preview_edits(&proposal.edits, true)?.session)
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
        let stale = self.proposal_stale_targets(&before, &proposal);
        if !force && !stale.is_empty() {
            return Err(ProposalError::Stale(stale));
        }
        let ProposalStage {
            session: preview,
            snapshot,
            resolved,
            ..
        } = self.preview_edits(&proposal.edits, force)?;
        let applied = before != snapshot;
        if applied {
            let (_, update) = preview.staged_update(&self.doc.transact().state_vector())?;
            let structural = proposal.edits.iter().any(|edit| {
                matches!(
                    edit,
                    ProposalEdit::InsertSlide { .. }
                        | ProposalEdit::DuplicateSlide { .. }
                        | ProposalEdit::DeleteSlide { .. }
                        | ProposalEdit::MoveSlide { .. }
                        | ProposalEdit::AddTextBox { .. }
                        | ProposalEdit::AddShape { .. }
                        | ProposalEdit::AddPicture { .. }
                        | ProposalEdit::RemoveShape { .. }
                )
            });
            if structural {
                self.add_undo_barrier();
            }
            let adopted = self.adopt(&preview, update, Adoption::Proposal);
            if structural {
                self.add_undo_barrier();
            }
            adopted?;
        }
        self.reject_proposal(id);
        let changed_targets = changes_for(&before, &snapshot, &resolved)?
            .iter()
            .filter(|change| {
                change.kind.is_some()
                    || change.before != change.after
                    || change.old_text != change.new_text
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
        store.baselines.remove(id);
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

    fn proposal_stale_targets(&self, snapshot: &DeckSnapshot, proposal: &Proposal) -> Vec<String> {
        let store = self.proposals.borrow();
        let Some(baseline) = store.baselines.get(&proposal.id) else {
            return Vec::new();
        };
        stale_targets(snapshot, baseline)
    }

    pub fn proposal_slide_strip(&self, id: &str) -> ProposalResult<Vec<ProposalStripEntry>> {
        let live = self.snapshot()?;
        let preview = self.preview_proposal(id)?;
        let mut entries: Vec<_> = preview
            .snapshot
            .slides
            .iter()
            .enumerate()
            .map(|(index, slide)| {
                let live_index = live.slides.iter().position(|old| old.id == slide.id);
                let status = if live_index.is_none() {
                    ProposalStripStatus::Inserted
                } else if preview.proposal.changes.iter().any(|change| {
                    change.slide_id == slide.id && change.kind == Some(ProposalChangeKind::Moved)
                }) {
                    ProposalStripStatus::Moved
                } else if preview.proposal.changes.iter().any(|change| {
                    change.slide_id == slide.id
                        && (change.before != change.after || change.old_text != change.new_text)
                }) {
                    ProposalStripStatus::Changed
                } else {
                    ProposalStripStatus::Unchanged
                };
                ProposalStripEntry {
                    slide_id: slide.id.clone(),
                    status,
                    live_index,
                    preview_index: Some(index),
                }
            })
            .collect();
        for (index, slide) in live.slides.iter().enumerate() {
            if !preview
                .snapshot
                .slides
                .iter()
                .any(|next| next.id == slide.id)
            {
                entries.insert(
                    index.min(entries.len()),
                    ProposalStripEntry {
                        slide_id: slide.id.clone(),
                        status: ProposalStripStatus::Deleted,
                        live_index: Some(index),
                        preview_index: None,
                    },
                );
            }
        }
        Ok(entries)
    }

    pub(crate) fn preview_edits(
        &self,
        edits: &[ProposalEdit],
        force: bool,
    ) -> ProposalResult<ProposalStage> {
        let total_bytes = edits.iter().fold(0usize, |total, edit| {
            total.saturating_add(match edit {
                ProposalEdit::AddPicture { draft, .. } => draft.media_bytes.len(),
                _ => 0,
            })
        });
        let limit_bytes = 8 * 1024 * 1024;
        if total_bytes > limit_bytes {
            return Err(ProposalError::PictureMediaTooLarge {
                total_bytes,
                limit_bytes,
            });
        }
        let preview = self.stage()?;
        let mut refs = BTreeMap::new();
        let mut resolved = Vec::with_capacity(edits.len());
        for edit in edits {
            let edit = resolve_edit(edit, &refs, &preview)?;
            let reference = match &edit {
                ProposalEdit::InsertSlide { reference, .. }
                | ProposalEdit::DuplicateSlide { reference, .. }
                | ProposalEdit::AddTextBox { reference, .. }
                | ProposalEdit::AddShape { reference, .. }
                | ProposalEdit::AddPicture { reference, .. } => reference.as_ref(),
                _ => None,
            };
            if let Some(reference) = reference
                && (reference.is_empty()
                    || reference.starts_with('@')
                    || refs.contains_key(reference))
            {
                return Err(ProposalError::Invalid(format!(
                    "invalid or duplicate ref {reference:?}"
                )));
            }
            let created = apply_structural_edit(&preview, &edit, force)?;
            if let Some(reference) = reference {
                refs.insert(
                    reference.clone(),
                    created.expect("creating edit returns an id"),
                );
            }
            resolved.push(edit);
        }
        let snapshot = preview.validated_snapshot()?;
        Ok(ProposalStage {
            session: preview,
            snapshot,
            refs,
            resolved,
        })
    }
}

fn resolve_edit(
    edit: &ProposalEdit,
    refs: &BTreeMap<String, String>,
    session: &DeckSession,
) -> ProposalResult<ProposalEdit> {
    let resolve = |id: &mut String| -> ProposalResult<()> {
        if let Some(reference) = id.strip_prefix('@') {
            *id = refs
                .get(reference)
                .cloned()
                .ok_or_else(|| ProposalError::Invalid(format!("unknown ref {reference:?}")))?;
        }
        Ok(())
    };
    let mut edit = edit.clone();
    match &mut edit {
        ProposalEdit::InsertSlide {
            after_slide_id: Some(id),
            ..
        } => resolve(id)?,
        ProposalEdit::InsertSlide {
            after_slide_id: None,
            ..
        } => {}
        ProposalEdit::DuplicateSlide {
            slide_id,
            after_slide_id,
            ..
        } => {
            resolve(slide_id)?;
            if let Some(Some(id)) = after_slide_id {
                resolve(id)?;
            }
        }
        ProposalEdit::MoveSlide {
            slide_id,
            after_slide_id,
        } => {
            resolve(slide_id)?;
            if let Some(id) = after_slide_id {
                resolve(id)?;
            }
        }
        ProposalEdit::DeleteSlide { slide_id }
        | ProposalEdit::SetSlideNotes { slide_id, .. }
        | ProposalEdit::AddTextBox { slide_id, .. }
        | ProposalEdit::AddShape { slide_id, .. }
        | ProposalEdit::AddPicture { slide_id, .. } => resolve(slide_id)?,
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
        }
        | ProposalEdit::RemoveShape { slide_id, shape_id } => {
            resolve(slide_id)?;
            resolve(shape_id)?;
        }
        ProposalEdit::ReplaceText { story_id, .. }
        | ProposalEdit::FormatText { story_id, .. }
        | ProposalEdit::SetParagraphAlignment { story_id, .. } => {
            if story_id.starts_with('@') {
                resolve(story_id)?;
                let snapshot = session.snapshot()?;
                let shape = snapshot
                    .slides
                    .iter()
                    .find_map(|slide| find_shape(&slide.shapes, story_id))
                    .ok_or_else(|| EditError::ShapeNotFound(story_id.clone()))?;
                *story_id = shape
                    .text_stories
                    .first()
                    .ok_or_else(|| EditError::StoryNotFound(story_id.clone()))?
                    .id
                    .clone();
            }
        }
    }
    Ok(edit)
}

fn apply_structural_edit(
    session: &DeckSession,
    edit: &ProposalEdit,
    force: bool,
) -> ProposalResult<Option<String>> {
    let context = EditCtx::local("proposal");
    let index_after = |anchor: Option<&str>, moving: Option<&str>| -> ProposalResult<u32> {
        let snapshot = session.snapshot()?;
        if let Some(id) = moving {
            check_target(&snapshot, id, None)?;
        }
        let slides: Vec<_> = snapshot
            .slides
            .iter()
            .filter(|slide| Some(slide.id.as_str()) != moving)
            .collect();
        match anchor {
            None => Ok(0),
            Some(id) if Some(id) == moving => snapshot
                .slides
                .iter()
                .position(|slide| slide.id == id)
                .map(|index| index as u32)
                .ok_or_else(|| EditError::SlideNotFound(id.into()).into()),
            Some(id) => slides
                .iter()
                .position(|slide| slide.id == id)
                .map(|index| index as u32 + 1)
                .or_else(|| force.then_some(slides.len() as u32))
                .ok_or_else(|| EditError::SlideNotFound(id.into()).into()),
        }
    };
    match edit {
        ProposalEdit::AddTextBox {
            slide_id, draft, ..
        } => Ok(Some(
            session.add_text_box(&context, slide_id, draft)?.shape_id,
        )),
        ProposalEdit::AddShape {
            slide_id, draft, ..
        } => Ok(Some(session.add_shape(&context, slide_id, draft)?.shape_id)),
        ProposalEdit::AddPicture {
            slide_id, draft, ..
        } => Ok(Some(
            session.add_picture(&context, slide_id, draft)?.shape_id,
        )),
        ProposalEdit::RemoveShape { slide_id, shape_id } => {
            session.remove_shape(&context, slide_id, shape_id)?;
            Ok(None)
        }
        ProposalEdit::InsertSlide {
            after_slide_id,
            layout_part_path,
            ..
        } => {
            let index = index_after(after_slide_id.as_deref(), None)?;
            Ok(Some(
                session
                    .insert_slide(&context, index, layout_part_path.as_deref())?
                    .slide_id,
            ))
        }
        ProposalEdit::DuplicateSlide {
            slide_id,
            after_slide_id,
            ..
        } => {
            check_target(&session.snapshot()?, slide_id, None)?;
            let anchor = after_slide_id
                .as_ref()
                .map_or(Some(slide_id.as_str()), |id| id.as_deref());
            let index = index_after(anchor, None)?;
            Ok(Some(
                session.duplicate_slide(&context, slide_id, index)?.slide_id,
            ))
        }
        ProposalEdit::DeleteSlide { slide_id } => {
            session.delete_slide(&context, slide_id)?;
            Ok(None)
        }
        ProposalEdit::MoveSlide {
            slide_id,
            after_slide_id,
        } => {
            let index = index_after(after_slide_id.as_deref(), Some(slide_id))?;
            session.move_slide(&context, slide_id, index)?;
            Ok(None)
        }
        _ => {
            let snapshot = session.snapshot()?;
            let (slide_id, shape_id) = target(&snapshot, edit)?;
            check_target(&snapshot, &slide_id, shape_id.as_deref())?;
            apply_edit(session, edit)?;
            Ok(None)
        }
    }
}

pub(crate) fn apply_edit(session: &DeckSession, edit: &ProposalEdit) -> Result<(), EditError> {
    let context = EditCtx::local("proposal");
    match edit {
        ProposalEdit::AddTextBox { .. }
        | ProposalEdit::AddShape { .. }
        | ProposalEdit::AddPicture { .. }
        | ProposalEdit::RemoveShape { .. }
        | ProposalEdit::InsertSlide { .. }
        | ProposalEdit::DuplicateSlide { .. }
        | ProposalEdit::DeleteSlide { .. }
        | ProposalEdit::MoveSlide { .. } => {
            return Err(EditError::InvalidState(
                "structural edit requires proposal resolution".into(),
            ));
        }
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
        ProposalEdit::InsertSlide { .. }
        | ProposalEdit::DuplicateSlide { .. }
        | ProposalEdit::DeleteSlide { .. }
        | ProposalEdit::MoveSlide { .. } => Err(EditError::InvalidState(
            "structural edit requires proposal resolution".into(),
        )),
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
        ProposalEdit::RemoveShape { slide_id, shape_id }
        | ProposalEdit::SetShapeRect {
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
        ProposalEdit::AddTextBox { slide_id, .. }
        | ProposalEdit::AddShape { slide_id, .. }
        | ProposalEdit::AddPicture { slide_id, .. }
        | ProposalEdit::SetSlideNotes { slide_id, .. } => Ok((slide_id.clone(), None)),
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
    let mut changes = BTreeMap::new();
    for edit in edits {
        if matches!(
            edit,
            ProposalEdit::AddTextBox { .. }
                | ProposalEdit::AddShape { .. }
                | ProposalEdit::AddPicture { .. }
                | ProposalEdit::InsertSlide { .. }
                | ProposalEdit::DuplicateSlide { .. }
                | ProposalEdit::DeleteSlide { .. }
                | ProposalEdit::MoveSlide { .. }
        ) {
            continue;
        }
        let Ok((slide_id, shape_id)) = target(after, edit).or_else(|_| target(before, edit)) else {
            continue;
        };
        let (old, old_text) = capture(before, &slide_id, shape_id.as_deref()).unwrap_or_default();
        let (new, new_text) = capture(after, &slide_id, shape_id.as_deref()).unwrap_or_default();
        if shape_id.is_some() && old.is_none() && new.is_none() {
            continue;
        }
        changes.insert(
            (slide_id.clone(), shape_id.clone()),
            ProposalChange {
                kind: match (&old, &new, &shape_id) {
                    (None, Some(_), Some(_)) => Some(ProposalChangeKind::Inserted),
                    (Some(_), None, Some(_)) => Some(ProposalChangeKind::Deleted),
                    _ => None,
                },
                slide_id,
                shape_id,
                before: old,
                after: new,
                old_text,
                new_text,
                before_index: None,
                after_index: None,
            },
        );
    }
    let added_slides: BTreeSet<_> = edits
        .iter()
        .filter_map(|edit| match edit {
            ProposalEdit::AddTextBox { slide_id, .. }
            | ProposalEdit::AddShape { slide_id, .. }
            | ProposalEdit::AddPicture { slide_id, .. } => Some(slide_id),
            _ => None,
        })
        .collect();
    for slide_id in added_slides {
        let old_slide = before.slides.iter().find(|slide| &slide.id == slide_id);
        if let Some(slide) = after.slides.iter().find(|slide| &slide.id == slide_id) {
            for shape in &slide.shapes {
                if old_slide.is_some_and(|old| find_shape(&old.shapes, &shape.id).is_some()) {
                    continue;
                }
                changes.insert(
                    (slide_id.clone(), Some(shape.id.clone())),
                    ProposalChange {
                        slide_id: slide_id.clone(),
                        shape_id: Some(shape.id.clone()),
                        before: None,
                        after: Some(shape.clone()),
                        old_text: String::new(),
                        new_text: shape_text(shape),
                        kind: Some(ProposalChangeKind::Inserted),
                        before_index: None,
                        after_index: None,
                    },
                );
            }
        }
    }
    let ids: BTreeSet<_> = before
        .slides
        .iter()
        .chain(&after.slides)
        .map(|slide| slide.id.clone())
        .collect();
    for slide_id in ids {
        let before_index = before.slides.iter().position(|slide| slide.id == slide_id);
        let after_index = after.slides.iter().position(|slide| slide.id == slide_id);
        let kind = match (before_index, after_index) {
            (None, Some(_)) => Some(ProposalChangeKind::Inserted),
            (Some(_), None) => Some(ProposalChangeKind::Deleted),
            (Some(a), Some(b)) if a != b && edits.iter().any(|edit| matches!(edit, ProposalEdit::MoveSlide { slide_id: id, .. } if id == &slide_id)) => Some(ProposalChangeKind::Moved),
            _ => None,
        };
        if let Some(kind) = kind {
            let key = (slide_id.clone(), None);
            let change = changes.entry(key).or_insert_with(|| ProposalChange {
                slide_id: slide_id.clone(),
                shape_id: None,
                before: None,
                after: None,
                old_text: capture(before, &slide_id, None).unwrap_or_default().1,
                new_text: capture(after, &slide_id, None).unwrap_or_default().1,
                kind: None,
                before_index: None,
                after_index: None,
            });
            change.kind = Some(kind);
            change.before_index = Some(before_index);
            change.after_index = Some(after_index);
        }
    }
    Ok(changes.into_values().collect())
}

fn stale_targets(snapshot: &DeckSnapshot, baseline: &ProposalBaseline) -> Vec<String> {
    let original = |id: &str| baseline.snapshot.slides.iter().find(|slide| slide.id == id);
    let live = |id: &str| snapshot.slides.iter().find(|slide| slide.id == id);
    let mut stale = BTreeSet::new();
    for edit in &baseline.edits {
        if matches!(
            edit,
            ProposalEdit::AddTextBox { .. }
                | ProposalEdit::AddShape { .. }
                | ProposalEdit::AddPicture { .. }
        ) {
            continue;
        }
        if let Ok((slide_id, shape_id)) = target(&baseline.snapshot, edit)
            && capture(&baseline.snapshot, &slide_id, shape_id.as_deref()).is_ok()
            && capture(snapshot, &slide_id, shape_id.as_deref()).ok()
                != capture(&baseline.snapshot, &slide_id, shape_id.as_deref()).ok()
        {
            stale.insert(shape_id.unwrap_or(slide_id));
        }
    }
    for edit in &baseline.edits {
        let (source, anchor, fingerprint) = match edit {
            ProposalEdit::InsertSlide { after_slide_id, .. } => {
                (None, after_slide_id.as_deref(), false)
            }
            ProposalEdit::DuplicateSlide {
                slide_id,
                after_slide_id,
                ..
            } => (
                Some(slide_id.as_str()),
                after_slide_id
                    .as_ref()
                    .map_or(Some(slide_id.as_str()), |id| id.as_deref()),
                true,
            ),
            ProposalEdit::DeleteSlide { slide_id } => (Some(slide_id.as_str()), None, true),
            ProposalEdit::AddTextBox { slide_id, .. }
            | ProposalEdit::AddShape { slide_id, .. }
            | ProposalEdit::AddPicture { slide_id, .. } => (Some(slide_id.as_str()), None, false),
            ProposalEdit::MoveSlide {
                slide_id,
                after_slide_id,
            } => (Some(slide_id.as_str()), after_slide_id.as_deref(), false),
            _ => continue,
        };
        if let Some(id) = source
            && let Some(before) = original(id)
            && (live(id).is_none() || (fingerprint && live(id) != Some(before)))
        {
            stale.insert(id.to_owned());
        }
        if let Some(id) = anchor
            && original(id).is_some()
            && live(id).is_none()
        {
            stale.insert(id.to_owned());
        }
    }
    stale.into_iter().collect()
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
