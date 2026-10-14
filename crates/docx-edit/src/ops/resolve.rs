//! Tracked-change resolution: `accept_change` and `reject_change`.
//!
//! Each stamp class resolves two ways:
//!
//! | stamp    | accept                           | reject                           |
//! |----------|----------------------------------|----------------------------------|
//! | `ins`    | drop the stamp, text stays       | remove the text                  |
//! | `del`    | remove the text                  | drop the stamp, text stays       |
//! | `pPrIns` | clear the marker, split stays    | remove the pilcrow, paragraphs join |
//! | `pPrDel` | remove the pilcrow, paragraphs join | clear the marker, split stays |
//!
//! A unit carrying BOTH `ins` and `del` — one author suggesting over another's
//! suggestion — is removed either way: the remove class wins over the keep
//! class on the same content.
//!
//! Removing a boundary pilcrow joins two paragraphs, and it is the FOLLOWING
//! paragraph's mark that survives, so the merged paragraph keeps the SECOND
//! paragraph's properties and paraId. That is the OOXML rule — the surviving
//! `w:p` owns the properties — and it deliberately differs from a plain
//! delete, which models a user removing a paragraph mark rather than a
//! revision being applied. A story's FINAL pilcrow is never removed, because a
//! document always keeps its last paragraph mark; a join that would remove it
//! clears the markers instead. A mark before a surviving block embed stays when
//! paragraph content would otherwise interrupt the block, inheriting any pending
//! revision that could remove that block.
//!
//! Resolving authors no new revision id; the context's suggesting mode is ignored.
//!
//! Structural table-row revisions (`trIns`/`trDel`) live in each row's `trPr`
//! bag and resolve in the same transaction as story-unit revisions; removing a
//! row also removes the cell stories it made unreachable. Paragraph-property
//! revisions (`pPrChange`) resolve alongside them: accepting drops the record,
//! rejecting restores the properties it captured.

use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::table::resolve_table_row_revisions;
use crate::ops::{Chunk, ChunkKind, inherit_block_revisions, snapshot};
use crate::queries::revision_parts;
use crate::{
    DEL, EditCtx, EditingDoc, INS, PPR_CHANGE, PPR_DEL, PPR_INS, RevisionId, StoryRange,
    check_range, story_ref,
};

/// What a resolve op targets.
#[derive(Clone, Debug, PartialEq)]
pub enum ChangeTarget {
    /// Resolve every tracked change overlapping the range (no id filtering).
    Range(StoryRange),
    /// Resolve every unit stamped with this revision id, in any story.
    Revision(RevisionId),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ResolveMode {
    Accept,
    Reject,
}

/// Returns the stamp value when it is active (non-null) and — under a revision-id
/// filter — carries that id. Stamps without a parseable revision id never match a filter
/// (they remain resolvable by range).
fn active_stamp(value: Option<Any>, filter: Option<&str>) -> Option<Any> {
    let value = value?;
    if value == Any::Null {
        return None;
    }
    match filter {
        None => Some(value),
        Some(id) => match revision_parts(&value) {
            Some((stamp_id, ..)) if stamp_id == id => Some(value),
            _ => None,
        },
    }
}

fn map_stamp<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<Any> {
    match map.get(txn, key) {
        Some(Out::Any(value)) => Some(value),
        _ => None,
    }
}

/// Records the revision id carried by a resolved stamp (deduplicated, resolution order).
fn record(resolved: &mut Vec<String>, stamp: Option<&Any>) {
    let Some(stamp) = stamp else {
        return;
    };
    if let Some((id, ..)) = revision_parts(stamp)
        && !resolved.contains(&id)
    {
        resolved.push(id);
    }
}

fn clear_attr(txn: &mut TransactionMut<'_>, story: &TextRef, start: u32, len: u32, key: &str) {
    story.format(txn, start, len, Attrs::from([(Arc::from(key), Any::Null)]));
}

fn resolve_inline_content(
    content: &Any,
    mode: ResolveMode,
    filter: Option<&str>,
    resolved: &mut Vec<String>,
) -> Any {
    let Any::Array(children) = content else {
        return content.clone();
    };
    let mut remaining = Vec::new();
    for child in children.iter() {
        let Any::Map(child) = child else {
            remaining.push(child.clone());
            continue;
        };
        let mut child = child.as_ref().clone();
        let attrs = match child.get("attrs") {
            Some(Any::Map(attrs)) => attrs.as_ref().clone(),
            _ => Default::default(),
        };
        let ins = active_stamp(attrs.get(INS).cloned(), filter);
        let del = active_stamp(attrs.get(DEL).cloned(), filter);
        let remove = match mode {
            ResolveMode::Accept => del.as_ref(),
            ResolveMode::Reject => ins.as_ref(),
        };
        if remove.is_some() {
            record(resolved, remove);
            continue;
        }
        let (key, keep) = match mode {
            ResolveMode::Accept => (INS, ins.as_ref()),
            ResolveMode::Reject => (DEL, del.as_ref()),
        };
        if keep.is_some() {
            record(resolved, keep);
            let mut attrs = attrs;
            attrs.remove(key);
            child.insert("attrs".to_owned(), Any::Map(Arc::new(attrs)));
        }
        if let Some(Any::Map(payload)) = child.get("payload")
            && let Some(content) = payload.get("content")
        {
            let content = resolve_inline_content(content, mode, filter, resolved);
            let mut payload = payload.as_ref().clone();
            payload.insert("content".to_owned(), content);
            child.insert("payload".to_owned(), Any::Map(Arc::new(payload)));
        }
        remaining.push(Any::Map(Arc::new(child)));
    }
    Any::Array(Arc::from(remaining))
}

fn property_map<'a>(
    change: &'a Any,
    key: &str,
) -> Option<&'a std::collections::HashMap<String, Any>> {
    let Any::Map(change) = change else {
        return None;
    };
    match change.get(key) {
        Some(Any::Map(value)) => Some(value.as_ref()),
        _ => None,
    }
}

/// Rewinds a pilcrow to the `previousFormatting` a `pPrChange` captured:
/// properties the change introduced are removed, then the previous values are
/// written back. Introducing `numPr` also clears the derived list attributes.
/// Schema-managed keys are never touched.
fn restore_paragraph_properties(txn: &mut TransactionMut<'_>, map: &MapRef, change: &Any) {
    let previous = property_map(change, "previousFormatting");
    let current = property_map(change, "currentFormatting");
    if let Some(current) = current {
        for key in current.keys() {
            if previous.is_none_or(|prior| !prior.contains_key(key))
                && !crate::is_identity_key(key)
                && !matches!(key.as_str(), PPR_INS | PPR_DEL | PPR_CHANGE)
            {
                map.remove(txn, key);
            }
        }
        if current.contains_key("numPr")
            && previous.is_none_or(|prior| !prior.contains_key("numPr"))
        {
            for key in [
                "numPr",
                "listIsBullet",
                "listNumFmt",
                "listMarker",
                "listLevel",
                "listStart",
            ] {
                map.remove(txn, key);
            }
        }
    }
    if let Some(previous) = previous {
        for (key, value) in previous {
            if !crate::is_identity_key(key)
                && !matches!(key.as_str(), PPR_INS | PPR_DEL | PPR_CHANGE)
            {
                map.insert(txn, key.clone(), value.clone());
            }
        }
    }
}

/// Resolves the `pPrChange` records on one pilcrow. Matching records are
/// consumed — rejecting restores their captured properties, accepting only
/// drops them — and non-matching records stay for a later resolve.
fn resolve_paragraph_property_changes(
    txn: &mut TransactionMut<'_>,
    map: &MapRef,
    mode: ResolveMode,
    filter: Option<&str>,
    resolved: &mut Vec<String>,
) {
    let changes = match map.get(txn, PPR_CHANGE) {
        Some(Out::Any(Any::Array(changes))) => changes.to_vec(),
        _ => return,
    };
    let total = changes.len();
    let mut remaining = Vec::new();
    for change in changes {
        if active_stamp(Some(change.clone()), filter).is_some() {
            record(resolved, Some(&change));
            if mode == ResolveMode::Reject {
                restore_paragraph_properties(txn, map, &change);
            }
        } else {
            remaining.push(change);
        }
    }
    if remaining.is_empty() {
        map.remove(txn, PPR_CHANGE);
    } else if remaining.len() < total {
        map.insert(txn, PPR_CHANGE, Any::Array(Arc::from(remaining)));
    }
}

fn removes_chunk<T: ReadTxn>(
    chunk: &Chunk,
    txn: &T,
    mode: ResolveMode,
    span: (u32, u32),
    filter: Option<&str>,
) -> bool {
    if chunk.start < span.0 || chunk.end() > span.1 {
        return false;
    }
    let (attr_key, ppr_key) = match mode {
        ResolveMode::Accept => (DEL, PPR_DEL),
        ResolveMode::Reject => (INS, PPR_INS),
    };
    active_stamp(chunk.attrs.get(attr_key).cloned(), filter).is_some()
        || matches!(&chunk.kind, ChunkKind::Pilcrow(map)
            if active_stamp(map_stamp(map, txn, ppr_key), filter).is_some())
}

/// Resolves one story's tracked changes in place. `span` limits the walk to a story range
/// (`None` = the whole story, the by-id path); `filter` limits it to one revision id.
/// Returns the number of units physically removed inside `span`.
fn resolve_story(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    mode: ResolveMode,
    span: Option<(u32, u32)>,
    filter: Option<&str>,
    resolved: &mut Vec<String>,
) -> u32 {
    let (span_start, span_end) = span.unwrap_or((0, u32::MAX));
    let chunks = snapshot(story, txn);
    let final_pilcrow = chunks.iter().rev().find_map(|chunk| match chunk.kind {
        ChunkKind::Pilcrow(_) => Some(chunk.start),
        _ => None,
    });
    let mut content = false;
    let content_before: Vec<bool> = chunks
        .iter()
        .map(|chunk| {
            let before = content;
            if Some(chunk.start) == final_pilcrow
                || !removes_chunk(chunk, txn, mode, (span_start, span_end), filter)
            {
                content =
                    !matches!(chunk.kind, ChunkKind::Pilcrow(_)) && !chunk.is_block_embed(txn);
            }
            before
        })
        .collect();
    let mut removed = 0;
    let mut next_block_revisions = None;
    // Reverse walk so physical removals never shift the indices still to be visited.
    for (slot, chunk) in chunks.iter().enumerate().rev() {
        let overlap_start = chunk.start.max(span_start);
        let overlap_end = chunk.end().min(span_end);
        if overlap_end <= overlap_start {
            next_block_revisions = chunk.block_revisions(txn);
            continue;
        }
        match &chunk.kind {
            ChunkKind::Pilcrow(map) => {
                resolve_paragraph_property_changes(txn, map, mode, filter, resolved);
                // A suggested split/merge stamps BOTH the pilcrow unit's text attr and the
                // pPr marker; either signal (matching the filter) selects the mark.
                let ppr_ins = active_stamp(map_stamp(map, txn, PPR_INS), filter);
                let ppr_del = active_stamp(map_stamp(map, txn, PPR_DEL), filter);
                let attr_ins = active_stamp(chunk.attrs.get(INS).cloned(), filter);
                let attr_del = active_stamp(chunk.attrs.get(DEL).cloned(), filter);
                let ins_hit = ppr_ins.is_some() || attr_ins.is_some();
                let del_hit = ppr_del.is_some() || attr_del.is_some();
                let join = match mode {
                    ResolveMode::Accept => del_hit,
                    ResolveMode::Reject => ins_hit,
                };
                let boundary_revisions = next_block_revisions
                    .as_ref()
                    .filter(|_| content_before[slot]);
                if join && Some(chunk.start) != final_pilcrow && boundary_revisions.is_none() {
                    match mode {
                        ResolveMode::Accept => {
                            record(resolved, ppr_del.as_ref());
                            record(resolved, attr_del.as_ref());
                        }
                        ResolveMode::Reject => {
                            record(resolved, ppr_ins.as_ref());
                            record(resolved, attr_ins.as_ref());
                        }
                    }
                    story.remove_range(txn, chunk.start, 1);
                    removed += 1;
                    continue;
                }
                for (ppr_key, attr_key, ppr_stamp, attr_stamp) in [
                    (PPR_INS, INS, ppr_ins, attr_ins),
                    (PPR_DEL, DEL, ppr_del, attr_del),
                ] {
                    record(resolved, ppr_stamp.as_ref());
                    record(resolved, attr_stamp.as_ref());
                    if ppr_stamp.is_some() {
                        map.remove(txn, ppr_key);
                    }
                    if attr_stamp.is_some() {
                        clear_attr(txn, story, chunk.start, 1, attr_key);
                    }
                }
                if join
                    && Some(chunk.start) != final_pilcrow
                    && let Some(revisions) = boundary_revisions
                {
                    inherit_block_revisions(txn, story, chunk.start, map, revisions);
                }
                next_block_revisions = None;
            }
            ChunkKind::Text(_) | ChunkKind::Embed(_) => {
                let ins = active_stamp(chunk.attrs.get(INS).cloned(), filter);
                let del = active_stamp(chunk.attrs.get(DEL).cloned(), filter);
                // A unit carrying BOTH stamps is removed in either mode: the
                // remove class wins over the keep class on the same content.
                let remove = match mode {
                    ResolveMode::Accept => del.is_some(),
                    ResolveMode::Reject => ins.is_some(),
                };
                if remove {
                    match mode {
                        ResolveMode::Accept => record(resolved, del.as_ref()),
                        ResolveMode::Reject => record(resolved, ins.as_ref()),
                    }
                    story.remove_range(txn, overlap_start, overlap_end - overlap_start);
                    removed += overlap_end - overlap_start;
                    if overlap_start > chunk.start || overlap_end < chunk.end() {
                        next_block_revisions = None;
                    }
                } else {
                    if let ChunkKind::Embed(Some(map)) = &chunk.kind
                        && crate::map_string(map, txn, crate::KIND_KEY).as_deref() == Some("sdt")
                        && let Some(Out::Any(content)) = map.get(txn, "content")
                    {
                        let mut tracked = false;
                        crate::inline_content::visit(&content, &mut |child| {
                            if let Some(Any::Map(attrs)) = child.get("attrs") {
                                tracked |= attrs.contains_key(INS) || attrs.contains_key(DEL);
                            }
                        });
                        if tracked {
                            let updated = resolve_inline_content(&content, mode, filter, resolved);
                            if updated != content {
                                map.insert(txn, "content", updated);
                            }
                        }
                    }
                    match mode {
                        ResolveMode::Accept if ins.is_some() => {
                            record(resolved, ins.as_ref());
                            clear_attr(txn, story, overlap_start, overlap_end - overlap_start, INS);
                        }
                        ResolveMode::Reject if del.is_some() => {
                            record(resolved, del.as_ref());
                            clear_attr(txn, story, overlap_start, overlap_end - overlap_start, DEL);
                        }
                        _ => {}
                    }
                    next_block_revisions = chunk.block_revisions(txn).map(|mut revisions| {
                        for stamp in &mut revisions {
                            if active_stamp(stamp.clone(), filter).is_some() {
                                *stamp = None;
                            }
                        }
                        revisions
                    });
                }
            }
        }
    }
    removed
}

impl EditingDoc {
    /// Accepts the targeted changes: pending insertions become plain content
    /// and pending deletions are carried out. See the module docs for the full
    /// matrix and the join rule.
    ///
    /// The receipt's `revision_ids` lists the ids resolved, deduplicated and
    /// in resolution order; a range target also echoes the surviving range.
    /// A [`ChangeTarget::Range`] errors when empty, a
    /// [`ChangeTarget::Revision`] when the id matches nothing.
    pub fn accept_change(&self, ctx: &EditCtx, target: &ChangeTarget) -> OpResult<Receipt> {
        self.resolve_change(ctx, target, ResolveMode::Accept)
    }

    /// Rejects the targeted changes — the inverse of
    /// [`EditingDoc::accept_change`]: pending insertions are rolled back and
    /// pending deletions restored to plain content. Same receipt and errors.
    pub fn reject_change(&self, ctx: &EditCtx, target: &ChangeTarget) -> OpResult<Receipt> {
        self.resolve_change(ctx, target, ResolveMode::Reject)
    }

    fn resolve_change(
        &self,
        ctx: &EditCtx,
        target: &ChangeTarget,
        mode: ResolveMode,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let mut resolved: Vec<String> = Vec::new();
        match target {
            ChangeTarget::Range(range) => {
                let len = crate::format::range_len(range)?;
                if len == 0 {
                    return Err(OpError::EmptyRange);
                }
                let story = story_ref(&txn, &range.story)?;
                check_range(&story, &txn, range.start, len)?;
                let (start, end) =
                    crate::ops::code_point_range(&story, &txn, range.start, range.end);
                let range = &StoryRange::new(range.story.clone(), start, end);
                let removed_tables = resolve_table_row_revisions(
                    &mut txn,
                    &story,
                    &range.story,
                    mode == ResolveMode::Accept,
                    Some((range.start, range.end)),
                    None,
                    &mut resolved,
                )?;
                let removed = resolve_story(
                    &mut txn,
                    &story,
                    mode,
                    Some((range.start, range.end - removed_tables)),
                    None,
                    &mut resolved,
                );
                let loc_range = loc_range_in_txn(
                    &range.story,
                    &story,
                    &txn,
                    range.start,
                    range.end - removed_tables - removed,
                )?;
                Ok(Receipt {
                    new_para_ids: Vec::new(),
                    revision_ids: resolved,
                    range: Some(loc_range),
                })
            }
            ChangeTarget::Revision(revision_id) => {
                let stories: Vec<(String, TextRef)> = {
                    let Some(stories) = txn.get_map(crate::STORIES) else {
                        return Err(OpError::UnknownChange(revision_id.clone()));
                    };
                    let mut ids: Vec<String> =
                        stories.keys(&txn).map(|key| key.to_string()).collect();
                    ids.sort();
                    ids.into_iter()
                        .filter_map(|story_id| match stories.get(&txn, &story_id) {
                            Some(Out::YText(story)) => Some((story_id, story)),
                            _ => None,
                        })
                        .collect()
                };
                for (story_id, story) in &stories {
                    resolve_table_row_revisions(
                        &mut txn,
                        story,
                        story_id,
                        mode == ResolveMode::Accept,
                        None,
                        Some(revision_id.as_str()),
                        &mut resolved,
                    )?;
                    resolve_story(
                        &mut txn,
                        story,
                        mode,
                        None,
                        Some(revision_id.as_str()),
                        &mut resolved,
                    );
                }
                if resolved.is_empty() {
                    return Err(OpError::UnknownChange(revision_id.clone()));
                }
                Ok(Receipt {
                    new_para_ids: Vec::new(),
                    revision_ids: resolved,
                    range: None,
                })
            }
        }
    }
}
