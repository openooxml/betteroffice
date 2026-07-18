//! Text ops: `insert_text`, `delete_range`, `replace_range`, `replace_range_rich`,
//! `insert_hard_break`, `insert_tab` (op-contract §1 "Text").
//!
//! Story vocabulary (render-bridge spec): a tab is the `\t` character (one UTF-16 unit); a hard
//! break is a one-unit embed carrying `_kind: "break"`. Inserted text may never contain a
//! paragraph or line break — paragraph structure changes only through `split_paragraph` /
//! `merge_paragraphs`, line breaks only through `insert_hard_break`.

use std::collections::BTreeMap;
use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapPrelim, Text, TextRef, TransactionMut};

use crate::format::{FormatPolicy, HYPERLINK, PROTECTED_ATTRS};
use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::{
    Chunk, ChunkKind, adjacent_paragraph_change_revision_id, adjacent_revision_id, adopt_pilcrow,
    capture_pilcrow, snapshot, utf16_len,
};
use crate::{
    BREAK_KIND, DEL, EditCtx, EditingDoc, INS, KIND_KEY, Position, StoryRange, check_position,
    check_range, revision_value, story_ref,
};

const FORBIDDEN_TEXT_CHARS: [char; 5] = ['\n', '\r', '\u{000B}', '\u{2028}', '\u{2029}'];

fn validate_text(text: &str) -> OpResult<()> {
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

/// Resolves the formatting attributes a plain insertion should carry (PM typing semantics).
///
/// `Inherit` copies the formatting of the unit before `at` (or after it at a paragraph start).
/// Tracked-change stamps are never inherited. A hyperlink is inherited only when the units on
/// BOTH sides carry the same hyperlink value — the `inclusive: false` PM boundary rule, so typing
/// at a link edge does not extend the link.
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

/// Suggesting-mode delete: the op-contract 3-way split.
///
/// - plain retained content → `del` revision on the retained range;
/// - the author's OWN pending insertions → real CRDT removal (they were never "in" the document);
/// - pilcrows → retained with a `del` stamp AND a `pPrDel` paragraph-mark revision (the story's
///   final pilcrow is never marked — Word cannot delete the final paragraph mark);
/// - content already carrying a `del` stamp is left untouched (no double-stamping).
pub(crate) fn suggest_delete(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    ctx: &EditCtx,
    revision: &Any,
    start: u32,
    end: u32,
) -> DeleteOutcome {
    let chunks = snapshot(story, txn);
    let final_pilcrow = chunks.iter().rev().find_map(|chunk| match chunk.kind {
        ChunkKind::Pilcrow(_) => Some(chunk.start),
        _ => None,
    });
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
) -> DeleteOutcome {
    let chunks = snapshot(story, txn);
    let pilcrows_in_range: Vec<(u32, yrs::MapRef)> = chunks
        .iter()
        .filter_map(|chunk| match &chunk.kind {
            ChunkKind::Pilcrow(map) if chunk.start >= start && chunk.start < end => {
                Some((chunk.start, map.clone()))
            }
            _ => None,
        })
        .collect();
    let final_pilcrow = chunks.iter().rev().find_map(|chunk| match &chunk.kind {
        ChunkKind::Pilcrow(map) => Some((chunk.start, map.clone())),
        _ => None,
    });
    let donor = pilcrows_in_range
        .first()
        .map(|(_, map)| capture_pilcrow(map, txn));

    let final_in_range = final_pilcrow
        .as_ref()
        .filter(|(index, _)| *index >= start && *index < end)
        .cloned();
    let (removed, survivor) = if let Some((final_index, final_map)) = final_in_range {
        // Keep the final pilcrow alive: remove around it.
        if end > final_index + 1 {
            story.remove_range(txn, final_index + 1, end - final_index - 1);
        }
        if final_index > start {
            story.remove_range(txn, start, final_index - start);
        }
        ((end - start) - 1, Some(final_map))
    } else {
        story.remove_range(txn, start, end - start);
        let survivor = if pilcrows_in_range.is_empty() {
            None
        } else {
            snapshot(story, txn)
                .into_iter()
                .find_map(|chunk| match chunk.kind {
                    ChunkKind::Pilcrow(map) if chunk.start >= start => Some(map),
                    _ => None,
                })
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
    /// Inserts break-free text with explicit tracked-change stamps and a [`FormatPolicy`]
    /// (op-contract R2: never a bare insert).
    pub fn insert_text(
        &self,
        ctx: &EditCtx,
        at: Position,
        text: &str,
        policy: FormatPolicy,
    ) -> OpResult<Receipt> {
        validate_text(text)?;
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        if text.is_empty() {
            let range = loc_range_in_txn(&at.story, &story, &txn, at.index, at.index)?;
            return Ok(Receipt {
                range: Some(range),
                ..Receipt::default()
            });
        }
        let chunks = snapshot(&story, &txn);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, INS, &ctx.author)
                .or_else(|| {
                    adjacent_paragraph_change_revision_id(&chunks, at.index, &txn, &ctx.author)
                })
                .unwrap_or_else(|| self.next_id())
        });
        let formatting = policy_attrs(&chunks, at.index, &policy);
        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        story.insert_with_attributes(&mut txn, at.index, text, stamped_attrs(formatting, ins));
        let end = at.index + utf16_len(text);
        let range = loc_range_in_txn(&at.story, &story, &txn, at.index, end)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(range),
        })
    }

    /// Deletes a range (op-contract §1). Plain mode removes it with the R6 pPr-survival rule;
    /// suggesting mode performs the 3-way split documented on [`suggest_delete`].
    pub fn delete_range(&self, ctx: &EditCtx, range: StoryRange) -> OpResult<Receipt> {
        let len = crate::format::range_len(&range)?;
        if len == 0 {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        check_range(&story, &txn, range.start, len)?;
        let chunks = snapshot(&story, &txn);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, range.start, DEL, &ctx.author)
                .or_else(|| adjacent_revision_id(&chunks, range.end, DEL, &ctx.author))
                .unwrap_or_else(|| self.next_id())
        });
        let result_end = if let Some(id) = revision_id.as_ref() {
            let revision = revision_value(id, &ctx.revision_author());
            let outcome = suggest_delete(&mut txn, &story, ctx, &revision, range.start, range.end);
            range.end - outcome.removed
        } else {
            plain_delete(&mut txn, &story, range.start, range.end);
            range.start
        };
        let loc_range = loc_range_in_txn(&range.story, &story, &txn, range.start, result_end)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(loc_range),
        })
    }

    /// Replaces a range with plain text in ONE transaction — the primitive behind
    /// type-over-selection, paste, find-replace, and agent proposeChange. In suggesting mode the
    /// delete and insert stamps share one revision ID. The inserted text adopts the formatting of
    /// the first replaced text unit (type-over keeps formatting); for a collapsed range it
    /// inherits like typing.
    pub fn replace_range(&self, ctx: &EditCtx, range: StoryRange, text: &str) -> OpResult<Receipt> {
        validate_text(text)?;
        let len = crate::format::range_len(&range)?;
        if len == 0 && text.is_empty() {
            return Err(OpError::EmptyRange);
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &range.story)?;
        check_range(&story, &txn, range.start, len)?;

        let chunks = snapshot(&story, &txn);
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
            .unwrap_or_else(|| policy_attrs(&chunks, range.start, &FormatPolicy::Inherit));

        let revision = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        if len > 0 {
            if let Some(revision) = revision.as_ref() {
                suggest_delete(&mut txn, &story, ctx, revision, range.start, range.end);
            } else {
                plain_delete(&mut txn, &story, range.start, range.end);
            }
        }
        if !text.is_empty() {
            story.insert_with_attributes(
                &mut txn,
                range.start,
                text,
                stamped_attrs(formatting, revision),
            );
        }
        let end = range.start + utf16_len(text);
        let loc_range = loc_range_in_txn(&range.story, &story, &txn, range.start, end)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(loc_range),
        })
    }

    /// [`EditingDoc::replace_range`] with explicitly formatted runs (rich paste / proposeChange
    /// with formatting). All runs and the delete share one revision ID in suggesting mode.
    pub fn replace_range_rich(
        &self,
        ctx: &EditCtx,
        range: StoryRange,
        runs: &[RichRun],
    ) -> OpResult<Receipt> {
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
        let chunks = snapshot(&story, &txn);
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
                suggest_delete(&mut txn, &story, ctx, revision, range.start, range.end);
            } else {
                plain_delete(&mut txn, &story, range.start, range.end);
            }
        }
        let mut cursor = range.start;
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
        let loc_range = loc_range_in_txn(&range.story, &story, &txn, range.start, cursor)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(loc_range),
        })
    }

    /// Inserts a one-unit hard-break embed (`_kind: "break"`), inheriting run formatting like
    /// typing.
    pub fn insert_hard_break(&self, ctx: &EditCtx, at: Position) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        let chunks = snapshot(&story, &txn);
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
        let chunks = snapshot(&story, &txn);
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
