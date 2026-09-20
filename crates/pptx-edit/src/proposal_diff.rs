use std::collections::{BTreeSet, HashSet};
use std::sync::atomic::Ordering;

use pptx_parse::PptxPackage;
use serde::{Deserialize, Serialize};
use yrs::{ArrayRef, Map, MapRef, ReadTxn, Transact};

use crate::comments::{snapshot_comments, snapshot_flavor};
use crate::deck::{
    live_shape_order, map_string, map_string_array, required_map, required_order, slide_notes,
    slide_ref, slide_shape_order, snapshot_shape, string_array_ref,
};
use crate::proposals::{apply_edit, shape_text};
use crate::{
    DeckSession, DeckSnapshot, DeckUndoManager, EditError, EditResult, Proposal, ProposalChange,
    ProposalEdit, ProposalResult, SHAPES, SLIDES, ShapeSnapshot, SlideScope, StorySnapshot,
    TextRunSnapshot, TextStyle, doc_with_client_id, hydrate_doc,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ProposalTextChangeKind {
    Insertion,
    Deletion,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalTextChange {
    pub story_id: String,
    pub start: u32,
    pub end: u32,
    pub kind: ProposalTextChangeKind,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalDiffPreview {
    pub proposal: Proposal,
    pub snapshot: DeckSnapshot,
    pub text_changes: Vec<ProposalTextChange>,
}

/// Slide-scoped diff preview: `scope` is the annotated render input and
/// `snapshot` carries only that slide — the payload callers serialize.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalSlideDiff {
    pub proposal: Proposal,
    pub scope: SlideScope,
    pub snapshot: DeckSnapshot,
    pub text_changes: Vec<ProposalTextChange>,
}

impl DeckSession {
    /// Slide-scoped `preview_proposal_diff`: resolves targets, staleness and
    /// changes against the proposal's touched shapes only, and materializes
    /// just the rendered slide instead of the whole deck.
    pub fn preview_proposal_diff_slide(
        &self,
        id: &str,
        slide_index: usize,
    ) -> ProposalResult<ProposalSlideDiff> {
        let mut proposal = self.pending_proposal(id)?;
        let (targets, before, stale_targets) = {
            let txn = self.doc.transact();
            let mut targets = BTreeSet::new();
            for edit in &proposal.edits {
                let (slide_id, shape_id) = scoped_target(&txn, edit)?;
                scoped_capture(&txn, &self.package, &slide_id, shape_id.as_deref())?;
                targets.insert((slide_id, shape_id));
            }
            let stale_targets = scoped_stale_targets(&txn, &self.package, &proposal);
            let mut before = Vec::with_capacity(targets.len());
            for (slide_id, shape_id) in &targets {
                before.push(scoped_capture(
                    &txn,
                    &self.package,
                    slide_id,
                    shape_id.as_deref(),
                )?);
            }
            (targets, before, stale_targets)
        };
        proposal.stale_targets = stale_targets;
        let preview = self.preview_doc_with_edits(&proposal.edits)?;
        let (changes, comment_flavor, comments) = {
            let txn = preview.doc.transact();
            let mut changes = Vec::with_capacity(targets.len());
            for ((slide_id, shape_id), (before, old_text)) in targets.iter().zip(before) {
                let (after, new_text) =
                    scoped_capture(&txn, &preview.package, slide_id, shape_id.as_deref())?;
                changes.push(ProposalChange {
                    slide_id: slide_id.clone(),
                    shape_id: shape_id.clone(),
                    before,
                    after,
                    old_text,
                    new_text,
                });
            }
            (changes, snapshot_flavor(&txn)?, snapshot_comments(&txn)?)
        };
        proposal.changes = changes;
        let mut scope = preview.slide_scope(slide_index)?;
        let mut text_changes = Vec::new();
        for change in &proposal.changes {
            let Some(before) = &change.before else {
                continue;
            };
            if change.slide_id == scope.slide.id {
                let Some(after) = find_shape_mut(&mut scope.slide.shapes, &before.id) else {
                    continue;
                };
                for story in &mut after.text_stories {
                    if let Some(original) =
                        before.text_stories.iter().find(|old| old.id == story.id)
                    {
                        diff_story(original, story, &mut text_changes);
                    }
                }
            } else if let Some(after) = &change.after {
                let mut after = after.clone();
                for story in &mut after.text_stories {
                    if let Some(original) =
                        before.text_stories.iter().find(|old| old.id == story.id)
                    {
                        diff_story(original, story, &mut text_changes);
                    }
                }
            }
        }
        Ok(ProposalSlideDiff {
            proposal,
            snapshot: DeckSnapshot {
                width_emu: scope.width_emu,
                height_emu: scope.height_emu,
                slides: vec![scope.slide.clone()],
                comment_flavor,
                comments,
            },
            scope,
            text_changes,
        })
    }

    /// Hydrates a scratch doc from the live state and applies `edits` — the
    /// `preview_edits` tail, with target validation left to the caller.
    fn preview_doc_with_edits(&self, edits: &[ProposalEdit]) -> ProposalResult<DeckSession> {
        let doc = doc_with_client_id(self.client_id);
        hydrate_doc(&doc, &self.encode_state_as_update_v1())?;
        let undo = DeckUndoManager::new(&doc, self.client_id)?;
        let preview = DeckSession {
            doc,
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

    /// Builds a render-only snapshot whose offsets are not editable.
    pub fn preview_proposal_diff(&self, id: &str) -> ProposalResult<ProposalDiffPreview> {
        let preview = self.preview_proposal(id)?;
        let mut snapshot = preview.snapshot;
        let mut text_changes = Vec::new();
        for change in &preview.proposal.changes {
            let Some(before) = &change.before else {
                continue;
            };
            let Some(slide) = snapshot
                .slides
                .iter_mut()
                .find(|slide| slide.id == change.slide_id)
            else {
                continue;
            };
            let Some(after) = find_shape_mut(&mut slide.shapes, &before.id) else {
                continue;
            };
            for story in &mut after.text_stories {
                if let Some(original) = before.text_stories.iter().find(|old| old.id == story.id) {
                    diff_story(original, story, &mut text_changes);
                }
            }
        }
        Ok(ProposalDiffPreview {
            proposal: preview.proposal,
            snapshot,
            text_changes,
        })
    }
}

fn find_shape_mut<'a>(shapes: &'a mut [ShapeSnapshot], id: &str) -> Option<&'a mut ShapeSnapshot> {
    for shape in shapes {
        if shape.id == id {
            return Some(shape);
        }
        if let Some(child) = find_shape_mut(&mut shape.children, id) {
            return Some(child);
        }
    }
    None
}

/// `proposals::target` against the doc maps: only story edits need a lookup —
/// the first slide+shape owning the story in document order.
fn scoped_target<T: ReadTxn>(txn: &T, edit: &ProposalEdit) -> EditResult<(String, Option<String>)> {
    match edit {
        ProposalEdit::ReplaceText { story_id, .. }
        | ProposalEdit::FormatText { story_id, .. }
        | ProposalEdit::SetParagraphAlignment { story_id, .. } => story_owner(txn, story_id)?
            .map(|(slide_id, shape_id)| (slide_id, Some(shape_id)))
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

/// `proposals::capture` against the doc maps: materializes only the targeted
/// shape subtree instead of a whole deck snapshot.
fn scoped_capture<T: ReadTxn>(
    txn: &T,
    package: &PptxPackage,
    slide_id: &str,
    shape_id: Option<&str>,
) -> EditResult<(Option<ShapeSnapshot>, String)> {
    let slide = slide_ref(txn, slide_id)?;
    match shape_id {
        Some(shape_id) => {
            if !shape_in_tree(txn, &slide_shape_order(&slide, txn)?, shape_id)? {
                return Err(EditError::ShapeNotFound(shape_id.to_owned()));
            }
            let theme = pptx_parse::slide_theme(
                package,
                map_string(&slide, txn, "sourcePartPath").as_deref(),
                map_string(&slide, txn, "layoutPartPath").as_deref(),
            );
            let shape = snapshot_shape(
                &required_map(txn, SHAPES)?,
                &required_map(txn, crate::STORIES)?,
                txn,
                shape_id,
                &mut HashSet::new(),
                Some(&theme),
            )?;
            let text = shape_text(&shape);
            Ok((Some(shape), text))
        }
        None => Ok((None, slide_notes(&slide, txn, package))),
    }
}

fn scoped_stale_targets<T: ReadTxn>(
    txn: &T,
    package: &PptxPackage,
    proposal: &Proposal,
) -> Vec<String> {
    proposal
        .changes
        .iter()
        .filter(|change| {
            scoped_capture(txn, package, &change.slide_id, change.shape_id.as_deref())
                .map_or(true, |(shape, text)| {
                    shape != change.before || text != change.old_text
                })
        })
        .map(ProposalChange::key)
        .collect()
}

/// The first (slide_id, shape_id) owning `story_id`, walking each slide's
/// shape tree in document order like `find_story_shape` over a snapshot.
fn story_owner<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<Option<(String, String)>> {
    let slides = required_map(txn, SLIDES)?;
    let shapes = required_map(txn, SHAPES)?;
    let mut seen_slides = HashSet::new();
    for slide_id in string_array_ref(&required_order(txn)?, txn) {
        if !seen_slides.insert(slide_id.clone()) {
            continue;
        }
        let slide = slides
            .get(txn, &slide_id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .ok_or_else(|| EditError::InvalidState(format!("missing slide {slide_id}")))?;
        let mut pending = live_shape_order(&slide_shape_order(&slide, txn)?, txn)?;
        pending.reverse();
        let mut seen_shapes = HashSet::new();
        while let Some(shape_id) = pending.pop() {
            if !seen_shapes.insert(shape_id.clone()) {
                continue;
            }
            let shape = shapes
                .get(txn, &shape_id)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| EditError::InvalidState(format!("missing shape {shape_id}")))?;
            if map_string_array(&shape, txn, "textStories")?
                .iter()
                .any(|id| id.as_str() == story_id)
            {
                return Ok(Some((slide_id, shape_id)));
            }
            let mut children = map_string_array(&shape, txn, "children")?;
            children.reverse();
            pending.extend(children);
        }
    }
    Ok(None)
}

/// Whether `shape_id` is reachable from the slide's shape order — the scoped
/// equivalent of `find_shape` over a materialized slide.
fn shape_in_tree<T: ReadTxn>(txn: &T, order: &ArrayRef, shape_id: &str) -> EditResult<bool> {
    let shapes = required_map(txn, SHAPES)?;
    let mut pending = live_shape_order(order, txn)?;
    pending.reverse();
    let mut seen = HashSet::new();
    while let Some(id) = pending.pop() {
        if id == shape_id {
            return Ok(true);
        }
        if !seen.insert(id.clone()) {
            continue;
        }
        let shape = shapes
            .get(txn, &id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .ok_or_else(|| EditError::InvalidState(format!("missing shape {id}")))?;
        let mut children = map_string_array(&shape, txn, "children")?;
        children.reverse();
        pending.extend(children);
    }
    Ok(false)
}

#[derive(Clone, Debug, PartialEq)]
struct Token {
    text: String,
    style: TextStyle,
}

fn tokens(runs: &[TextRunSnapshot]) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    for run in runs {
        for ch in run.text.chars() {
            if let Some(last) = result.last_mut()
                && last.style == run.style
                && last.text.ends_with(char::is_whitespace) == ch.is_whitespace()
            {
                last.text.push(ch);
            } else {
                result.push(Token {
                    text: ch.to_string(),
                    style: run.style.clone(),
                });
            }
        }
    }
    result
}

fn diff_story(
    before: &StorySnapshot,
    after: &mut StorySnapshot,
    changes: &mut Vec<ProposalTextChange>,
) {
    let mut offset = 0;
    for paragraph in &mut after.paragraphs {
        let old = before.paragraphs.iter().find(|old| old.id == paragraph.id);
        let old = tokens(old.map_or(&[], |paragraph| &paragraph.runs));
        let new = tokens(&paragraph.runs);
        let mut runs: Vec<TextRunSnapshot> = Vec::new();
        for (token, kind) in diff_tokens(&old, &new) {
            let start = offset;
            offset += token.text.encode_utf16().count() as u32;
            let mut style = token.style.clone();
            if let Some(kind) = kind {
                style.color = Some(
                    match kind {
                        ProposalTextChangeKind::Insertion => "#166534",
                        ProposalTextChangeKind::Deletion => "#b91c1c",
                    }
                    .into(),
                );
                if let Some(last) = changes.last_mut()
                    && last.story_id == after.id
                    && last.kind == kind
                    && last.end == start
                {
                    last.end = offset;
                } else {
                    changes.push(ProposalTextChange {
                        story_id: after.id.clone(),
                        start,
                        end: offset,
                        kind,
                    });
                }
            }
            if let Some(last) = runs.last_mut()
                && last.style == style
            {
                last.text.push_str(&token.text);
            } else {
                runs.push(TextRunSnapshot {
                    text: token.text.clone(),
                    style,
                });
            }
        }
        paragraph.runs = runs;
        offset += 1;
    }
    after.length = offset;
}

fn diff_tokens<'a>(
    old: &'a [Token],
    new: &'a [Token],
) -> Vec<(&'a Token, Option<ProposalTextChangeKind>)> {
    use ProposalTextChangeKind::{Deletion, Insertion};
    let prefix = old.iter().zip(new).take_while(|(a, b)| a == b).count();
    let suffix = old[prefix..]
        .iter()
        .rev()
        .zip(new[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let mut result: Vec<_> = new[..prefix].iter().map(|token| (token, None)).collect();
    let a = &old[prefix..old.len() - suffix];
    let b = &new[prefix..new.len() - suffix];
    if a.len()
        .saturating_add(1)
        .saturating_mul(b.len().saturating_add(1))
        > 250_000
    {
        result.extend(a.iter().map(|token| (token, Some(Deletion))));
        result.extend(b.iter().map(|token| (token, Some(Insertion))));
    } else {
        let width = b.len() + 1;
        let mut lengths = vec![0u32; (a.len() + 1) * width];
        for i in (0..a.len()).rev() {
            for j in (0..b.len()).rev() {
                lengths[i * width + j] = if a[i] == b[j] {
                    lengths[(i + 1) * width + j + 1] + 1
                } else {
                    lengths[(i + 1) * width + j].max(lengths[i * width + j + 1])
                };
            }
        }
        let (mut i, mut j) = (0, 0);
        while i < a.len() || j < b.len() {
            if i < a.len() && j < b.len() && a[i] == b[j] {
                result.push((&b[j], None));
                i += 1;
                j += 1;
            } else if i < a.len()
                && (j == b.len() || lengths[(i + 1) * width + j] >= lengths[i * width + j + 1])
            {
                result.push((&a[i], Some(Deletion)));
                i += 1;
            } else {
                result.push((&b[j], Some(Insertion)));
                j += 1;
            }
        }
    }
    result.extend(new[new.len() - suffix..].iter().map(|token| (token, None)));
    result
}
