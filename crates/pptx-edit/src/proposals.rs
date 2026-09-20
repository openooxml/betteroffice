use std::collections::{BTreeMap, BTreeSet};
use std::sync::atomic::Ordering;

use serde::{Deserialize, Serialize};
use yrs::Transact;

use crate::{
    DeckSession, DeckSnapshot, DeckUndoManager, EditCtx, EditError, ShapeRect, ShapeSnapshot,
    ShapeStroke, TextStyle, TextStylePatch, decode_update_v1, doc_with_client_id, hydrate_doc,
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
    fn key(&self) -> String {
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
        let preview = self.preview_edits(&request.edits)?;
        let changes = changes_for(&before, &preview.snapshot()?, &request.edits)?;
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
        let before = self.snapshot()?;
        let snapshot = self.preview_edits(&proposal.edits)?.snapshot()?;
        proposal.stale_targets = stale_targets(&before, &proposal);
        proposal.changes = changes_for(&before, &snapshot, &proposal.edits)?;
        Ok(ProposalPreview { proposal, snapshot })
    }

    pub fn proposal_preview_session(&self, id: &str) -> ProposalResult<DeckSession> {
        self.preview_edits(&self.pending_proposal(id)?.edits)
    }

    pub fn accept_proposal(&self, id: &str, force: bool) -> ProposalResult<ProposalAcceptance> {
        let proposal = self.pending_proposal(id)?;
        let before = self.snapshot()?;
        let stale = stale_targets(&before, &proposal);
        if !force && !stale.is_empty() {
            return Err(ProposalError::Stale(stale));
        }
        let preview = self.preview_edits(&proposal.edits)?;
        let snapshot = preview.snapshot()?;
        let applied = before != snapshot;
        if applied {
            let update = preview.encode_diff_v1(&self.encode_state_vector_v1())?;
            let update = decode_update_v1(&update).map_err(EditError::InvalidUpdate)?;
            self.add_undo_barrier();
            self.doc
                .transact_mut_with(self.client_id)
                .apply_update(update)
                .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
            self.id_counter.store(
                preview.id_counter.load(Ordering::Relaxed),
                Ordering::Relaxed,
            );
            self.add_undo_barrier();
        }
        self.reject_proposal(id);
        Ok(ProposalAcceptance {
            proposal_id: id.to_owned(),
            applied,
            snapshot,
        })
    }

    pub fn reject_proposal(&self, id: &str) -> bool {
        let mut store = self.proposals.borrow_mut();
        let before = store.pending.len();
        store.pending.retain(|proposal| proposal.id != id);
        before != store.pending.len()
    }

    fn pending_proposal(&self, id: &str) -> ProposalResult<Proposal> {
        self.proposals
            .borrow()
            .pending
            .iter()
            .find(|proposal| proposal.id == id)
            .cloned()
            .ok_or_else(|| ProposalError::NotFound(id.to_owned()))
    }

    fn preview_edits(&self, edits: &[ProposalEdit]) -> ProposalResult<DeckSession> {
        let snapshot = self.snapshot()?;
        for edit in edits {
            let (slide_id, shape_id) = target(&snapshot, edit)?;
            capture(&snapshot, &slide_id, shape_id.as_deref())?;
        }
        let doc = doc_with_client_id(self.client_id);
        hydrate_doc(&doc, &self.encode_state_as_update_v1())?;
        let undo = DeckUndoManager::new(&doc, self.client_id)?;
        let preview = DeckSession {
            doc,
            staged: std::cell::RefCell::new(None),
            package_json: std::cell::RefCell::new(self.package_json.borrow().clone()),
            client_id: self.client_id,
            id_counter: self.id_counter.load(Ordering::Relaxed).into(),
            package: self.package.clone(),
            undo: std::cell::RefCell::new(undo),
            proposals: Default::default(),
        };
        for edit in edits {
            apply_edit(&preview, edit)?;
        }
        crate::deck::validate_doc(&preview.doc)?;
        Ok(preview)
    }
}

fn apply_edit(session: &DeckSession, edit: &ProposalEdit) -> Result<(), EditError> {
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

fn inherited_style(story: &crate::StorySnapshot, at: u32) -> TextStyle {
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

fn shape_text(shape: &ShapeSnapshot) -> String {
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
