//! Text editing operations.

use std::collections::BTreeMap;
use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapPrelim, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::format::{FormatPolicy, HYPERLINK, PROTECTED_ATTRS};
use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::{
    Chunk, ChunkKind, adjacent_paragraph_change_revision_id, adjacent_revision_id, adopt_pilcrow,
    block_embed_at, capture_pilcrow, last_pilcrow, paragraph_content_before, snapshot_range,
    utf16_len,
};
use crate::{
    BREAK_KIND, DEL, EditCtx, EditingDoc, INS, KIND_KEY, Position, StoryRange, check_position,
    check_range, revision_value, story_ref,
};

const FORBIDDEN_TEXT_CHARS: [char; 5] = ['\n', '\r', '\u{000B}', '\u{2028}', '\u{2029}'];

pub(crate) fn validate_text(text: &str) -> OpResult<()> {
    if text.contains(FORBIDDEN_TEXT_CHARS) {
        return Err(OpError::TextContainsBreak);
    }
    Ok(())
}

/// One explicitly formatted run for [`EditingDoc::replace_range_rich`].
#[derive(Clone, Debug, Default, PartialEq)]
pub struct RichRun {
    pub text: String,
    /// Formatting attributes in the story vocabulary (`bold`, `textColor`, ...). Tracked-change
    /// keys are ignored — the op stamps its own.
    pub attrs: BTreeMap<String, Any>,
}

/// Resolves inherited, explicit, or plain insertion attributes.
fn policy_attrs(chunks: &[Chunk], at: u32, policy: &FormatPolicy) -> Vec<(String, Any)> {
    match policy {
        FormatPolicy::Plain => Vec::new(),
        FormatPolicy::Explicit(map) => map
            .iter()
            .filter(|(key, _)| !matches!(key.as_str(), INS | DEL))
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect(),
        FormatPolicy::Inherit => {
            let unit_at = |index: u32| -> Option<&Chunk> {
                chunks
                    .iter()
                    .find(|chunk| chunk.start <= index && index < chunk.end())
            };
            let formatting_of = |chunk: &Chunk| -> Vec<(String, Any)> {
                chunk
                    .attrs
                    .iter()
                    .filter(|(key, value)| {
                        !matches!(key.as_str(), INS | DEL) && **value != Any::Null
                    })
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect()
            };
            let left = at
                .checked_sub(1)
                .and_then(unit_at)
                .filter(|chunk| !matches!(chunk.kind, ChunkKind::Pilcrow(_)));
            let right = unit_at(at).filter(|chunk| !matches!(chunk.kind, ChunkKind::Pilcrow(_)));
            let (source, other) = match (left, right) {
                (Some(left), right) => (left, right),
                (None, Some(right)) => (right, None),
                (None, None) => return Vec::new(),
            };
            let mut attrs = formatting_of(source);
            if let Some(index) = attrs.iter().position(|(key, _)| key == HYPERLINK) {
                let same_on_other =
                    other.is_some_and(|other| other.attrs.get(HYPERLINK) == Some(&attrs[index].1));
                if !same_on_other {
                    attrs.remove(index);
                }
            }
            attrs
        }
    }
}

/// [`policy_attrs`], except that inheriting into a paragraph without content takes its mark's run
/// formatting, as Word types into an empty paragraph.
fn insertion_attrs<T: ReadTxn>(
    chunks: &[Chunk],
    at: u32,
    policy: &FormatPolicy,
    txn: &T,
) -> Vec<(String, Any)> {
    if matches!(policy, FormatPolicy::Inherit) {
        let unit_at = |index: u32| {
            chunks
                .iter()
                .find(|chunk| chunk.start <= index && index < chunk.end())
        };
        let opens_paragraph = at
            .checked_sub(1)
            .and_then(unit_at)
            .is_none_or(|chunk| match &chunk.kind {
                ChunkKind::Pilcrow(_) => true,
                ChunkKind::Embed(Some(map)) => crate::map_string(map, txn, KIND_KEY)
                    .is_some_and(|kind| crate::segments::is_block_embed(&kind)),
                _ => false,
            });
        if opens_paragraph
            && let Some(ChunkKind::Pilcrow(mark)) = unit_at(at).map(|chunk| &chunk.kind)
            && let Some(Out::Any(defaults)) = mark.get(txn, "defaultTextFormatting")
        {
            return crate::seed::mark_run_attrs(&defaults);
        }
    }
    policy_attrs(chunks, at, policy)
}

fn stamped_attrs(formatting: Vec<(String, Any)>, ins: Option<Any>) -> Attrs {
    let mut attrs: Attrs = formatting
        .into_iter()
        .map(|(key, value)| (Arc::from(key.as_str()), value))
        .collect();
    attrs.insert(Arc::from(INS), ins.unwrap_or(Any::Null));
    attrs.insert(Arc::from(DEL), Any::Null);
    attrs
}

/// Outcome of the shared delete engine.
pub(crate) struct DeleteOutcome {
    /// Units physically removed from the story (all of them in plain mode; only own pending
    /// insertions in suggesting mode).
    pub removed: u32,
}

/// Chunks covering the insertion/deletion boundary.
fn boundary_chunks<T: yrs::ReadTxn>(story: &TextRef, txn: &T, index: u32) -> Vec<Chunk> {
    snapshot_range(story, txn, index.saturating_sub(1), index.saturating_add(1))
}

/// Where a suggested replacement of `range` inserts, before its deletion runs:
/// after the text it strikes out, as Word places it, and at most before the
/// story's final unit. None in plain mode, and where that unit is a block embed,
/// which only a paragraph boundary may precede: the text then goes at the start.
/// `chunks` cover `range` and the unit after it.
fn after_struck_text(
    story: &TextRef,
    txn: &TransactionMut<'_>,
    ctx: &EditCtx,
    range: &StoryRange,
    len: u32,
    chunks: &[Chunk],
) -> Option<u32> {
    if !ctx.is_suggesting() || len == 0 {
        return None;
    }
    let at = range.end.min(story.len(txn) - 1);
    let block = chunks.iter().any(|chunk| {
        chunk.start == at
            && matches!(&chunk.kind, ChunkKind::Embed(Some(map))
                if crate::map_string(map, txn, KIND_KEY)
                    .is_some_and(|kind| crate::segments::is_block_embed(&kind)))
    });
    (!block).then_some(at)
}

/// Stamps retained content, removes owned insertions, and protects the final pilcrow.
pub(crate) fn suggest_delete(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    ctx: &EditCtx,
    revision: &Any,
    start: u32,
    end: u32,
    chunks: &[Chunk],
) -> DeleteOutcome {
    let final_pilcrow = chunks
        .iter()
        .any(|chunk| {
            matches!(chunk.kind, ChunkKind::Pilcrow(_))
                && chunk.end().min(end) > chunk.start.max(start)
        })
        .then(|| last_pilcrow(story, txn).map(|(index, _)| index))
        .flatten();
    let mut removed = 0;
    for chunk in chunks.iter().rev() {
        let overlap_start = chunk.start.max(start);
        let overlap_end = chunk.end().min(end);
        if overlap_end <= overlap_start {
            continue;
        }
        let overlap = overlap_end - overlap_start;
        match &chunk.kind {
            ChunkKind::Pilcrow(map) => {
                if Some(chunk.start) == final_pilcrow || chunk.attr_active(DEL) {
                    continue;
                }
                story.format(
                    txn,
                    chunk.start,
                    1,
                    Attrs::from([(Arc::from(DEL), revision.clone())]),
                );
                map.insert(txn, crate::PPR_DEL, revision.clone());
            }
            ChunkKind::Text(_) | ChunkKind::Embed(_) => {
                if chunk.attr_active(INS)
                    && chunk.revision_author(INS).as_deref() == Some(ctx.author.as_str())
                {
                    story.remove_range(txn, overlap_start, overlap);
                    removed += overlap;
                } else if !chunk.attr_active(DEL) {
                    story.format(
                        txn,
                        overlap_start,
                        overlap,
                        Attrs::from([(Arc::from(DEL), revision.clone())]),
                    );
                }
            }
        }
    }
    DeleteOutcome { removed }
}

/// Plain-mode delete with the R6 survival rule: when the range removes pilcrows, the surviving
/// paragraph adopts the FIRST affected paragraph's pPr + paraId. The story's final pilcrow is
/// never removed (Word keeps the last paragraph mark); content around it still is.
pub(crate) fn plain_delete(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    start: u32,
    end: u32,
    chunks: &[Chunk],
    replacement_has_content: bool,
) -> DeleteOutcome {
    let pilcrows_in_range: Vec<(u32, yrs::MapRef)> = chunks
        .iter()
        .filter_map(|chunk| match &chunk.kind {
            ChunkKind::Pilcrow(map) if chunk.start >= start && chunk.start < end => {
                Some((chunk.start, map.clone()))
            }
            _ => None,
        })
        .collect();
    let final_pilcrow = if pilcrows_in_range.is_empty() {
        None
    } else {
        last_pilcrow(story, txn)
    };
    let donor = pilcrows_in_range
        .first()
        .map(|(_, map)| capture_pilcrow(map, txn));

    let protected_pilcrow = final_pilcrow
        .as_ref()
        .filter(|(index, _)| *index >= start && *index < end)
        .cloned()
        .or_else(|| {
            pilcrows_in_range
                .last()
                .filter(|_| {
                    block_embed_at(story, txn, end)
                        && (replacement_has_content || paragraph_content_before(story, txn, start))
                        && snapshot_range(story, txn, end, end.saturating_add(1))
                            .first()
                            .and_then(|chunk| chunk.block_revisions(txn))
                            .is_none_or(|revisions| revisions.iter().all(Option::is_none))
                })
                .cloned()
        });
    let (removed, survivor) = if let Some((index, map)) = protected_pilcrow {
        if end > index + 1 {
            story.remove_range(txn, index + 1, end - index - 1);
        }
        if index > start {
            story.remove_range(txn, start, index - start);
        }
        ((end - start) - 1, Some(map))
    } else {
        story.remove_range(txn, start, end - start);
        let survivor = if pilcrows_in_range.is_empty() {
            None
        } else {
            crate::next_pilcrow(story, txn, start).map(|(_, map)| map)
        };
        (end - start, survivor)
    };

    if let (Some((donor_id, donor_props)), Some(survivor)) = (donor, survivor) {
        let survivor_id = crate::map_string(&survivor, txn, crate::PARA_ID);
        if survivor_id.as_deref() != Some(donor_id.as_str()) {
            adopt_pilcrow(txn, &survivor, &donor_id, &donor_props);
        }
    }
    DeleteOutcome { removed }
}

impl EditingDoc {
    /// Inserts break-free text with explicit stamps and formatting policy.
    pub fn insert_text(
        &self,
        ctx: &EditCtx,
        at: Position,
        text: &str,
        policy: FormatPolicy,
    ) -> OpResult<Receipt> {
        self.insert_text_observed(ctx, at, text, policy, |_| {})
    }

    pub(crate) fn insert_text_observed(
        &self,
        ctx: &EditCtx,
        at: Position,
        text: &str,
        policy: FormatPolicy,
        observe: impl FnOnce(&Attrs),
    ) -> OpResult<Receipt> {
        validate_text(text)?;
        let mut txn = self.transact_for(ctx);
        self.insert_text_observed_in(&mut txn, ctx, at, text, policy, observe)
    }

    pub(crate) fn insert_text_observed_in(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        at: Position,
        text: &str,
        policy: FormatPolicy,
        observe: impl FnOnce(&Attrs),
    ) -> OpResult<Receipt> {
        validate_text(text)?;
        let story = story_ref(txn, &at.story)?;
        check_position(&story, txn, at.index)?;
        if text.is_empty() {
            let range = loc_range_in_txn(&at.story, &story, txn, at.index, at.index)?;
            return Ok(Receipt {
                range: Some(range),
                ..Receipt::default()
            });
        }
        crate::identity::promote_at(self, txn, &at.story, &story, at.index);
        let chunks = boundary_chunks(&story, txn, at.index);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, INS, &ctx.author)
                .or_else(|| {
                    adjacent_paragraph_change_revision_id(&chunks, at.index, txn, &ctx.author)
                })
                .unwrap_or_else(|| self.next_id())
        });
        let formatting = insertion_attrs(&chunks, at.index, &policy, txn);
        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let attrs = stamped_attrs(formatting, ins);
        observe(&attrs);
        story.insert_with_attributes(txn, at.index, text, attrs);
        let end = at.index + utf16_len(text);
        let range = loc_range_in_txn(&at.story, &story, txn, at.index, end)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(range),
        })
    }

    /// Deletes a range while preserving surviving paragraph properties.
    pub fn delete_range(&self, ctx: &EditCtx, range: StoryRange) -> OpResult<Receipt> {
        let len = crate::format::range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        self.delete_range_in(&mut txn, ctx, range)
    }

    pub(crate) fn delete_range_in(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        range: StoryRange,
    ) -> OpResult<Receipt> {
        let len = crate::format::range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let story = story_ref(txn, &range.story)?;
        check_range(&story, txn, range.start, len)?;
        let chunks = snapshot_range(
            &story,
            txn,
            range.start.saturating_sub(1),
            range.end.saturating_add(1),
        );
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, range.start, DEL, &ctx.author)
                .or_else(|| adjacent_revision_id(&chunks, range.end, DEL, &ctx.author))
                .unwrap_or_else(|| self.next_id())
        });
        let result_end = if let Some(id) = revision_id.as_ref() {
            let revision = revision_value(id, &ctx.revision_author());
            let outcome =
                suggest_delete(txn, &story, ctx, &revision, range.start, range.end, &chunks);
            range.end - outcome.removed
        } else {
            plain_delete(txn, &story, range.start, range.end, &chunks, false);
            range.start
        };
        let loc_range = loc_range_in_txn(&range.story, &story, txn, range.start, result_end)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(loc_range),
        })
    }

    /// Replaces a range with plain text in ONE transaction — the primitive behind
    /// type-over-selection, paste, find-replace, and agent proposeChange. In suggesting mode the
    /// delete and insert stamps share one revision ID; insertion follows the struck-out text,
    /// as Word places it. The inserted text adopts the formatting of the first replaced text
    /// unit (type-over keeps formatting); for a collapsed range it
    /// inherits like typing.
    pub fn replace_range(&self, ctx: &EditCtx, range: StoryRange, text: &str) -> OpResult<Receipt> {
        self.replace_range_placed(ctx, range, text)
            .map(|(receipt, _)| receipt)
    }

    /// [`EditingDoc::replace_range`], with the story index the text was inserted at.
    pub(crate) fn replace_range_placed(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        text: &str,
    ) -> OpResult<(Receipt, u32)> {
        validate_text(text)?;
        let len = crate::format::range_len(&range)?;
        if len == 0 && text.is_empty() {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        self.replace_range_placed_in(&mut txn, ctx, range, text)
    }

    pub(crate) fn replace_range_placed_in(
        &self,
        txn: &mut TransactionMut<'_>,
        ctx: &EditCtx,
        range: StoryRange,
        text: &str,
    ) -> OpResult<(Receipt, u32)> {
        validate_text(text)?;
        let len = crate::format::range_len(&range)?;
        if len == 0 && text.is_empty() {
            return Err(OpError::EmptyRange);
        }
        let story = story_ref(txn, &range.story)?;
        check_range(&story, txn, range.start, len)?;
        let chunks = snapshot_range(
            &story,
            txn,
            range.start.saturating_sub(1),
            range.end.saturating_add(1),
        );
        let after = after_struck_text(&story, txn, ctx, &range, len, &chunks);
        let mut at = after.unwrap_or(range.start);
        if !text.is_empty() {
            crate::identity::promote_at(self, txn, &range.story, &story, at);
        }
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, range.start, INS, &ctx.author)
                .or_else(|| adjacent_revision_id(&chunks, range.start, DEL, &ctx.author))
                .or_else(|| adjacent_revision_id(&chunks, range.end, INS, &ctx.author))
                .or_else(|| adjacent_revision_id(&chunks, range.end, DEL, &ctx.author))
                .unwrap_or_else(|| self.next_id())
        });
        let formatting = chunks
            .iter()
            .find(|chunk| {
                matches!(chunk.kind, ChunkKind::Text(_))
                    && chunk.end() > range.start
                    && chunk.start < range.end
            })
            .map(|chunk| {
                chunk
                    .attrs
                    .iter()
                    .filter(|(key, value)| {
                        !matches!(key.as_str(), INS | DEL) && **value != Any::Null
                    })
                    .map(|(key, value)| (key.clone(), value.clone()))
                    .collect()
            })
            .unwrap_or_else(|| insertion_attrs(&chunks, range.start, &FormatPolicy::Inherit, txn));

        let revision = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        if len > 0 {
            if let Some(revision) = revision.as_ref() {
                let outcome =
                    suggest_delete(txn, &story, ctx, revision, range.start, range.end, &chunks);
                if after.is_some() {
                    at -= outcome.removed;
                }
            } else {
                plain_delete(
                    txn,
                    &story,
                    range.start,
                    range.end,
                    &chunks,
                    !text.is_empty(),
                );
            }
        }
        if !text.is_empty() {
            story.insert_with_attributes(txn, at, text, stamped_attrs(formatting, revision));
        }
        let end = at + utf16_len(text);
        let loc_range = loc_range_in_txn(&range.story, &story, txn, at, end)?;
        Ok((
            Receipt {
                new_para_ids: Vec::new(),
                revision_ids: revision_id.into_iter().collect(),
                range: Some(loc_range),
            },
            at,
        ))
    }

    /// [`EditingDoc::replace_range`] with explicitly formatted runs (rich paste / proposeChange
    /// with formatting). All runs and the delete share one revision ID in suggesting mode;
    /// insertion follows the struck-out text, as Word places it.
    pub fn replace_range_rich(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        runs: &[RichRun],
    ) -> OpResult<Receipt> {
        self.replace_range_rich_placed(ctx, range, runs)
            .map(|(receipt, _)| receipt)
    }

    /// [`EditingDoc::replace_range_rich`], with the story index the runs were inserted at.
    pub(crate) fn replace_range_rich_placed(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        runs: &[RichRun],
    ) -> OpResult<(Receipt, u32)> {
        for run in runs {
            validate_text(&run.text)?;
        }
        let len = crate::format::range_len(&range)?;
        let total: u32 = runs.iter().map(|run| utf16_len(&run.text)).sum();
        if len == 0 && total == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        check_range(&story, &txn, range.start, len)?;
        let chunks = snapshot_range(
            &story,
            &txn,
            range.start.saturating_sub(1),
            range.end.saturating_add(1),
        );
        let after = after_struck_text(&story, &txn, ctx, &range, len, &chunks);
        let mut at = after.unwrap_or(range.start);
        if total > 0 {
            crate::identity::promote_at(self, &mut txn, &range.story, &story, at);
        }
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, range.start, INS, &ctx.author)
                .or_else(|| adjacent_revision_id(&chunks, range.start, DEL, &ctx.author))
                .or_else(|| adjacent_revision_id(&chunks, range.end, INS, &ctx.author))
                .or_else(|| adjacent_revision_id(&chunks, range.end, DEL, &ctx.author))
                .unwrap_or_else(|| self.next_id())
        });
        let revision = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        if len > 0 {
            if let Some(revision) = revision.as_ref() {
                let outcome = suggest_delete(
                    &mut txn,
                    &story,
                    ctx,
                    revision,
                    range.start,
                    range.end,
                    &chunks,
                );
                if after.is_some() {
                    at -= outcome.removed;
                }
            } else {
                plain_delete(
                    &mut txn,
                    &story,
                    range.start,
                    range.end,
                    &chunks,
                    runs.iter().any(|run| !run.text.is_empty()),
                );
            }
        }
        let mut cursor = at;
        for run in runs {
            if run.text.is_empty() {
                continue;
            }
            let formatting: Vec<(String, Any)> = run
                .attrs
                .iter()
                .filter(|(key, _)| !PROTECTED_ATTRS.contains(&key.as_str()))
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect();
            story.insert_with_attributes(
                &mut txn,
                cursor,
                &run.text,
                stamped_attrs(formatting, revision.clone()),
            );
            cursor += utf16_len(&run.text);
        }
        let loc_range = loc_range_in_txn(&range.story, &story, &txn, at, cursor)?;
        Ok((
            Receipt {
                new_para_ids: Vec::new(),
                revision_ids: revision_id.into_iter().collect(),
                range: Some(loc_range),
            },
            at,
        ))
    }

    /// Inserts a one-unit hard-break embed (`_kind: "break"`), inheriting run formatting like
    /// typing.
    pub fn insert_hard_break(&self, ctx: &EditCtx, at: Position) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        crate::identity::promote_at(self, &mut txn, &at.story, &story, at.index);
        let chunks = boundary_chunks(&story, &txn, at.index);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, INS, &ctx.author)
                .unwrap_or_else(|| self.next_id())
        });
        let formatting = policy_attrs(&chunks, at.index, &FormatPolicy::Inherit);
        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let embed = story.insert_embed_with_attributes(
            &mut txn,
            at.index,
            MapPrelim::default(),
            stamped_attrs(formatting, ins),
        );
        embed.insert(&mut txn, KIND_KEY, BREAK_KIND);
        let range = loc_range_in_txn(&at.story, &story, &txn, at.index, at.index + 1)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(range),
        })
    }

    /// Inserts a one-unit tab. Tabs are the `\t` character in the story vocabulary (the render
    /// bridge splits text runs at `\t`), inheriting run formatting like typing.
    pub fn insert_tab(&self, ctx: &EditCtx, at: Position) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        crate::identity::promote_at(self, &mut txn, &at.story, &story, at.index);
        let chunks = boundary_chunks(&story, &txn, at.index);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, INS, &ctx.author)
                .unwrap_or_else(|| self.next_id())
        });
        let formatting = policy_attrs(&chunks, at.index, &FormatPolicy::Inherit);
        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        story.insert_with_attributes(&mut txn, at.index, "\t", stamped_attrs(formatting, ins));
        let range = loc_range_in_txn(&at.story, &story, &txn, at.index, at.index + 1)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(range),
        })
    }
}
