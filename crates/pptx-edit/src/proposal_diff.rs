use std::collections::{BTreeSet, HashSet};

use ooxml_diff::{DiffKind, DiffLimits, LimitFallback};
use pptx_parse::PptxPackage;
use serde::{Deserialize, Serialize};
use yrs::{ArrayRef, Map, MapRef, ReadTxn, Transact};

use crate::comments::{snapshot_comments, snapshot_flavor};
use crate::deck::{
    live_shape_order, live_slide_ids, map_string, map_string_array, required_map, slide_notes,
    slide_ref, slide_shape_order, snapshot_shape,
};
use crate::inherit::{SlideContext, record_inherited};
use crate::proposals::shape_text;
use crate::{
    DeckSession, DeckSnapshot, EditError, EditResult, ParagraphSnapshot, Proposal, ProposalChange,
    ProposalEdit, ProposalResult, SHAPES, SLIDES, ShapeSnapshot, SlideScope, StorySnapshot,
    TextRunSnapshot, TextStyle,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ProposalTextChangeKind {
    Insertion,
    Deletion,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ProposalShapeChangeKind {
    Inserted,
    Deleted,
    Moved,
    Resized,
    Changed,
    Reordered,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalShapeChange {
    pub shape_id: String,
    pub kind: ProposalShapeChangeKind,
}

fn proposal_shape_changes(proposal: &Proposal, slide_id: &str) -> Vec<ProposalShapeChange> {
    if !proposal.edits.iter().any(|edit| {
        matches!(
            edit,
            ProposalEdit::AddTextBox { .. }
                | ProposalEdit::AddShape { .. }
                | ProposalEdit::AddPicture { .. }
                | ProposalEdit::RemoveShape { .. }
                | ProposalEdit::SetShapeRect { .. }
                | ProposalEdit::SetShapeFill { .. }
                | ProposalEdit::SetShapeStroke { .. }
                | ProposalEdit::SetShapeAdjust { .. }
                | ProposalEdit::SetStoryParagraphs { .. }
        )
    }) {
        return Vec::new();
    }
    proposal
        .changes
        .iter()
        .filter_map(|change| {
            if change.slide_id != slide_id {
                return None;
            }
            let shape_id = change.shape_id.clone()?;
            let kind = match (&change.before, &change.after) {
                (None, Some(_)) => ProposalShapeChangeKind::Inserted,
                (Some(_), None) => ProposalShapeChangeKind::Deleted,
                (Some(before), Some(after))
                    if before.width != after.width || before.height != after.height =>
                {
                    ProposalShapeChangeKind::Resized
                }
                (Some(before), Some(after)) if before.x != after.x || before.y != after.y => {
                    ProposalShapeChangeKind::Moved
                }
                (Some(before), Some(after)) if before != after => ProposalShapeChangeKind::Changed,
                _ => return None,
            };
            Some(ProposalShapeChange { shape_id, kind })
        })
        .collect()
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

/// Slide-scoped diff preview: `scope` is the render input, `snapshot` carries that slide.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalSlideDiff {
    pub proposal: Proposal,
    pub scope: SlideScope,
    #[serde(default)]
    pub shape_changes: Vec<ProposalShapeChange>,
    pub snapshot: DeckSnapshot,
    pub text_changes: Vec<ProposalTextChange>,
}

impl DeckSession {
    /// Slide-scoped `preview_proposal_diff`.
    pub fn preview_proposal_diff_slide(
        &self,
        id: &str,
        slide_index: usize,
    ) -> ProposalResult<ProposalSlideDiff> {
        let mut proposal = self.pending_proposal(id)?;
        if proposal.edits.iter().any(|edit| {
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
            let slide_id =
                self.slide_ids()?
                    .get(slide_index)
                    .cloned()
                    .ok_or(EditError::OutOfBounds {
                        index: slide_index as u32,
                        length: self.slide_ids()?.len() as u32,
                    })?;
            return self.preview_proposal_diff_slide_by_id(id, &slide_id);
        }
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
                    kind: None,
                    before_index: None,
                    after_index: None,
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
            shape_changes: proposal_shape_changes(&proposal, &scope.slide.id),
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

    pub fn preview_proposal_diff_slide_by_id(
        &self,
        id: &str,
        slide_id: &str,
    ) -> ProposalResult<ProposalSlideDiff> {
        let diff = self.preview_proposal_diff(id)?;
        let preview = self.proposal_preview_session(id)?;
        let ids = preview.slide_ids()?;
        let mut scope = if let Some(index) = ids.iter().position(|id| id == slide_id) {
            preview.slide_scope(index)?
        } else {
            let live_ids = self.slide_ids()?;
            let index = live_ids
                .iter()
                .position(|id| id == slide_id)
                .ok_or_else(|| EditError::SlideNotFound(slide_id.to_owned()))?;
            self.slide_scope(index)?
        };
        if let Some(slide) = diff
            .snapshot
            .slides
            .iter()
            .find(|slide| slide.id == slide_id)
        {
            scope.slide = slide.clone();
        }
        Ok(ProposalSlideDiff {
            shape_changes: proposal_shape_changes(&diff.proposal, &scope.slide.id),
            proposal: diff.proposal,
            snapshot: DeckSnapshot {
                width_emu: scope.width_emu,
                height_emu: scope.height_emu,
                slides: vec![scope.slide.clone()],
                comment_flavor: diff.snapshot.comment_flavor,
                comments: diff.snapshot.comments,
            },
            scope,
            text_changes: diff.text_changes,
        })
    }

    fn preview_doc_with_edits(&self, edits: &[ProposalEdit]) -> ProposalResult<DeckSession> {
        Ok(self.preview_edits(edits, true)?.session)
    }

    /// Builds a render-only snapshot whose offsets are not editable.
    pub fn preview_proposal_diff(&self, id: &str) -> ProposalResult<ProposalDiffPreview> {
        let preview = self.preview_proposal(id)?;
        let mut snapshot = preview.snapshot;
        let before = self.snapshot()?;
        for slide in &mut snapshot.slides {
            if let Some(original) = before.slides.iter().find(|old| old.id == slide.id) {
                for (index, shape) in original.shapes.iter().enumerate() {
                    if preview.proposal.changes.iter().any(|change| {
                        change.slide_id == slide.id
                            && change.shape_id.as_deref() == Some(shape.id.as_str())
                            && change.before.is_some()
                            && change.after.is_none()
                    }) {
                        slide
                            .shapes
                            .insert(index.min(slide.shapes.len()), shape.clone());
                    }
                }
            }
        }
        let mut text_changes = Vec::new();
        for change in &preview.proposal.changes {
            if change.after.is_none() {
                continue;
            }
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
        | ProposalEdit::SetStoryParagraphs { story_id, .. }
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
        _ => Err(EditError::InvalidState(
            "structural edit requires proposal resolution".into(),
        )),
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
            let source_part_path = map_string(&slide, txn, "sourcePartPath");
            let layout_part_path = map_string(&slide, txn, "layoutPartPath");
            let theme = pptx_parse::slide_theme(
                package,
                source_part_path.as_deref(),
                layout_part_path.as_deref(),
            );
            let mut shape = snapshot_shape(
                &required_map(txn, SHAPES)?,
                &required_map(txn, crate::STORIES)?,
                txn,
                shape_id,
                &mut HashSet::new(),
                Some(&theme),
            )?;
            record_inherited(
                std::slice::from_mut(&mut shape),
                &SlideContext::new(
                    package,
                    source_part_path.as_deref(),
                    layout_part_path.as_deref(),
                ),
            );
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
    for slide_id in live_slide_ids(txn)? {
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

#[derive(PartialEq)]
struct ParagraphKey {
    text: String,
    alignment: Option<String>,
    level: u32,
    bullet_json: Option<String>,
    bigrams: BTreeSet<(char, char)>,
}

impl ParagraphKey {
    fn new(paragraph: &ParagraphSnapshot) -> Self {
        let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
        let bigrams = text
            .chars()
            .zip(text.chars().skip(1))
            .filter(|(a, b)| !a.is_whitespace() || !b.is_whitespace())
            .collect();
        Self {
            text,
            alignment: paragraph.alignment.clone(),
            level: paragraph.level,
            bullet_json: paragraph.bullet_json.clone(),
            bigrams,
        }
    }

    fn similarity(&self, other: &Self) -> u32 {
        if self.text == other.text {
            return 1000;
        }
        let total = self.bigrams.len() + other.bigrams.len();
        if total == 0 {
            return 0;
        }
        (2000 * self.bigrams.intersection(&other.bigrams).count() / total) as u32
    }
}

type ParagraphPair = (Option<usize>, Option<usize>);

fn align_changed_paragraphs(
    old: &[ParagraphKey],
    new: &[ParagraphKey],
    old_start: usize,
    new_start: usize,
    pairs: &mut Vec<ParagraphPair>,
) {
    let width = new.len().saturating_add(1);
    let cells = old.len().saturating_add(1).saturating_mul(width);
    if old.is_empty() || new.is_empty() || cells > PREVIEW_LCS_CELLS {
        pairs.extend((0..old.len()).map(|i| (Some(old_start + i), None)));
        pairs.extend((0..new.len()).map(|j| (None, Some(new_start + j))));
        return;
    }
    let mut scores = vec![(0u32, 0u32); cells];
    for i in (0..old.len()).rev() {
        for j in (0..new.len()).rev() {
            let next = scores[(i + 1) * width + j + 1];
            let paired = (next.0 + 1, next.1 + old[i].similarity(&new[j]));
            scores[i * width + j] = paired
                .max(scores[(i + 1) * width + j])
                .max(scores[i * width + j + 1]);
        }
    }
    let (mut i, mut j) = (0, 0);
    while i < old.len() || j < new.len() {
        if i < old.len() && j < new.len() {
            let next = scores[(i + 1) * width + j + 1];
            if scores[i * width + j] == (next.0 + 1, next.1 + old[i].similarity(&new[j])) {
                pairs.push((Some(old_start + i), Some(new_start + j)));
                i += 1;
                j += 1;
                continue;
            }
        }
        if i < old.len()
            && (j == new.len() || scores[(i + 1) * width + j] >= scores[i * width + j + 1])
        {
            pairs.push((Some(old_start + i), None));
            i += 1;
        } else {
            pairs.push((None, Some(new_start + j)));
            j += 1;
        }
    }
}

fn align_paragraphs(before: &StorySnapshot, after: &StorySnapshot) -> Vec<ParagraphPair> {
    let old: Vec<_> = before.paragraphs.iter().map(ParagraphKey::new).collect();
    let new: Vec<_> = after.paragraphs.iter().map(ParagraphKey::new).collect();
    let mut pairs = Vec::new();
    let (mut old_start, mut new_start) = (0, 0);
    for hunk in bounded_diff(&old, &new).hunks {
        if hunk.kind != DiffKind::Equal {
            continue;
        }
        align_changed_paragraphs(
            &old[old_start..hunk.old.start],
            &new[new_start..hunk.new.start],
            old_start,
            new_start,
            &mut pairs,
        );
        old_start = hunk.old.end;
        new_start = hunk.new.end;
        pairs.extend(hunk.old.zip(hunk.new).map(|(i, j)| (Some(i), Some(j))));
    }
    align_changed_paragraphs(
        &old[old_start..],
        &new[new_start..],
        old_start,
        new_start,
        &mut pairs,
    );
    pairs
}

fn diff_story(
    before: &StorySnapshot,
    after: &mut StorySnapshot,
    changes: &mut Vec<ProposalTextChange>,
) {
    let mut paragraphs = Vec::new();
    let mut offset = 0;
    for (old_index, new_index) in align_paragraphs(before, after) {
        let old = old_index.map(|i| &before.paragraphs[i]);
        let new = new_index.map(|j| &after.paragraphs[j]);
        let old_tokens = old.map(|p| tokens(&p.runs)).unwrap_or_default();
        let new_tokens = new.map(|p| tokens(&p.runs)).unwrap_or_default();
        let mut paragraph = new.or(old).expect("each pair has a paragraph").clone();
        paragraph.runs.clear();
        for (token, kind) in diff_tokens(&old_tokens, &new_tokens) {
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
            if let Some(last) = paragraph.runs.last_mut()
                && last.style == style
            {
                last.text.push_str(&token.text);
            } else {
                paragraph.runs.push(TextRunSnapshot {
                    text: token.text.clone(),
                    style,
                });
            }
        }
        paragraphs.push(paragraph);
        offset += 1;
    }
    after.paragraphs = paragraphs;
    after.length = offset;
}

/// Proposal previews coarsen a middle whose LCS table would exceed this many cells.
const PREVIEW_LCS_CELLS: usize = 250_000;

fn bounded_diff<T: PartialEq>(old: &[T], new: &[T]) -> ooxml_diff::TokenDiff {
    let limits = DiffLimits::new(usize::MAX, PREVIEW_LCS_CELLS, LimitFallback::ReplaceMiddle);
    ooxml_diff::diff_tokens(old, new, limits).expect("the replace-middle fallback never refuses")
}

fn diff_tokens<'a, T: PartialEq>(
    old: &'a [T],
    new: &'a [T],
) -> Vec<(&'a T, Option<ProposalTextChangeKind>)> {
    use ProposalTextChangeKind::{Deletion, Insertion};
    let diff = bounded_diff(old, new);
    let mut result = Vec::with_capacity(old.len().max(new.len()));
    for hunk in diff.hunks {
        match hunk.kind {
            DiffKind::Equal => result.extend(new[hunk.new].iter().map(|token| (token, None))),
            DiffKind::Delete => {
                result.extend(old[hunk.old].iter().map(|token| (token, Some(Deletion))))
            }
            DiffKind::Insert => {
                result.extend(new[hunk.new].iter().map(|token| (token, Some(Insertion))))
            }
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The proposal diff as it was before the LCS moved into `ooxml-diff`.
    fn reference<'a>(
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

    fn runs(text: &str, bold: bool) -> Vec<TextRunSnapshot> {
        vec![TextRunSnapshot {
            text: text.to_owned(),
            style: TextStyle {
                bold: Some(bold),
                ..TextStyle::default()
            },
        }]
    }

    fn assert_parity(old: &[TextRunSnapshot], new: &[TextRunSnapshot]) {
        let (old, new) = (tokens(old), tokens(new));
        assert_eq!(diff_tokens(&old, &new), reference(&old, &new));
    }

    #[test]
    fn preview_diff_matches_the_previous_lcs() {
        let cases = [
            ("", ""),
            ("same text", "same text"),
            ("the quick brown fox", "the slow brown fox jumps"),
            ("ab", "ba"),
            ("a b a b", "b a b a"),
            ("x y z", ""),
            ("", "x y z"),
            ("one two three", "three two one"),
            ("tied tied tied", "tied tied"),
        ];
        for (old, new) in cases {
            assert_parity(&runs(old, false), &runs(new, false));
        }
        assert_parity(&runs("style only", false), &runs("style only", true));
    }

    #[test]
    fn paragraph_alignment_keys_include_properties_but_ignore_ids_and_styles() {
        let first = ParagraphSnapshot {
            id: "old-first".into(),
            runs: runs("Same", false),
            ..ParagraphSnapshot::default()
        };
        for second in [
            ParagraphSnapshot {
                alignment: Some("ctr".into()),
                ..first.clone()
            },
            ParagraphSnapshot {
                level: 2,
                ..first.clone()
            },
            ParagraphSnapshot {
                bullet_json: Some("{\"type\":\"none\"}".into()),
                ..first.clone()
            },
        ] {
            let before = StorySnapshot {
                id: "story".into(),
                length: 10,
                paragraphs: vec![first.clone(), second.clone()],
            };
            let mut after = StorySnapshot {
                id: before.id.clone(),
                length: 5,
                paragraphs: vec![ParagraphSnapshot {
                    id: "new-id".into(),
                    runs: runs("Same", true),
                    ..second
                }],
            };
            assert_eq!(
                align_paragraphs(&before, &after),
                [(Some(0), None), (Some(1), Some(0))]
            );
            let expected = after.paragraphs[0].clone();
            let mut changes = Vec::new();
            diff_story(&before, &mut after, &mut changes);
            assert_eq!(after.plain_text(), "Same\nSameSame");
            assert_eq!(after.length, 14);
            assert_eq!(after.paragraphs[0].alignment, first.alignment);
            assert_eq!(after.paragraphs[0].level, first.level);
            assert_eq!(after.paragraphs[0].bullet_json, first.bullet_json);
            assert_eq!(after.paragraphs[1].alignment, expected.alignment);
            assert_eq!(after.paragraphs[1].level, expected.level);
            assert_eq!(after.paragraphs[1].bullet_json, expected.bullet_json);
            assert_eq!(
                changes
                    .iter()
                    .map(|c| (c.kind, c.start, c.end))
                    .collect::<Vec<_>>(),
                [
                    (ProposalTextChangeKind::Deletion, 0, 4),
                    (ProposalTextChangeKind::Deletion, 5, 9),
                    (ProposalTextChangeKind::Insertion, 9, 13),
                ]
            );
        }
    }

    #[test]
    fn oversized_paragraph_alignment_still_replaces_the_middle() {
        let story = |prefix: &str| StorySnapshot {
            id: "story".into(),
            length: 0,
            paragraphs: (0..500)
                .map(|index| ParagraphSnapshot {
                    runs: runs(&format!("{prefix}{index}"), false),
                    ..ParagraphSnapshot::default()
                })
                .collect(),
        };
        let old = story("old");
        let new = story("new");
        let pairs = align_paragraphs(&old, &new);
        assert_eq!(pairs.len(), 1000);
        assert!(
            pairs[..500]
                .iter()
                .enumerate()
                .all(|(i, pair)| *pair == (Some(i), None))
        );
        assert!(
            pairs[500..]
                .iter()
                .enumerate()
                .all(|(j, pair)| *pair == (None, Some(j)))
        );
    }

    #[test]
    fn oversized_preview_diffs_still_replace_the_middle() {
        let words = |prefix: &str| {
            (0..600)
                .map(|index| format!("{prefix}{index}"))
                .collect::<Vec<_>>()
                .join(" ")
        };
        let (old, new) = (
            format!("start {} end", words("a")),
            format!("start {} end", words("b")),
        );
        assert_parity(&runs(&old, false), &runs(&new, false));
        let (old_tokens, new_tokens) = (tokens(&runs(&old, false)), tokens(&runs(&new, false)));
        let diff = diff_tokens(&old_tokens, &new_tokens);
        let first_insertion = diff
            .iter()
            .position(|(_, kind)| *kind == Some(ProposalTextChangeKind::Insertion))
            .unwrap();
        assert!(
            diff[..first_insertion]
                .iter()
                .skip(2)
                .all(|(_, kind)| *kind == Some(ProposalTextChangeKind::Deletion))
        );
    }
}
