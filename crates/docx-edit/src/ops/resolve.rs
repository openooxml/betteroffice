//! Tracked-change RESOLVE ops: `accept_change`, `reject_change` (S4b).
//!
//! The op-contract 3-way matrix, mirroring the PM resolver
//! (`packages/core/src/prosemirror/commands/comments.ts` `resolveById`):
//!
//! - **accept:** `ins` text → drop the stamp (text stays); `del` text → physical removal;
//!   `pPrIns` → clear the marker (the split stays); `pPrDel` → remove the pilcrow (join).
//! - **reject:** `ins` text → physical removal; `del` text → drop the stamp (text stays);
//!   `pPrIns` → remove the pilcrow (join back); `pPrDel` → clear the marker (the split stays).
//!
//! Removing a boundary pilcrow joins two paragraphs; the FOLLOWING paragraph's pilcrow
//! survives, so the merged paragraph keeps the SECOND paragraph's pPr + paraId — exactly the
//! PM resolver's `inheritFromSecond` join and the OOXML rule (the surviving `w:p` owns the
//! properties). This deliberately differs from the plain-delete R6 donor rule, which models
//! a USER deletion, not a revision resolution. A story's FINAL pilcrow is never removed
//! (Word keeps the last paragraph mark): a join that would remove it clears the markers
//! instead — the PM resolver's last-paragraph edge case.
//!
//! Resolving is APPLYING a revision, not authoring one: no new revision is ever stamped and
//! the context's suggesting mode is ignored (the PM twin sets `SUGGESTION_BYPASS_META`).
//!
//! Structural table-row revisions (`trIns`/`trDel`) live in each structural
//! row's `trPr` bag. They are resolved in the same transaction as story-unit
//! revisions; physical row removal also removes unreachable cell stories.
//! `pPrChange`/`rPrChange` property-revision payloads remain a separate path.

use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::table::resolve_table_row_revisions;
use crate::ops::{ChunkKind, snapshot};
use crate::queries::revision_parts;
use crate::{
    DEL, EditCtx, EditingDoc, INS, KIND_KEY, PARA_ID, PPR_CHANGE, PPR_DEL, PPR_INS, RevisionId,
    StoryRange, check_range, story_ref,
};

/// What a resolve op targets: an explicit story range (the PM range commands) or one
/// coalesced revision id across every story (the PM by-id commands).
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

fn restore_paragraph_properties(txn: &mut TransactionMut<'_>, map: &MapRef, change: &Any) {
    let previous = property_map(change, "previousFormatting");
    let current = property_map(change, "currentFormatting");
    if let Some(current) = current {
        for key in current.keys() {
            if previous.is_none_or(|prior| !prior.contains_key(key))
                && !matches!(
                    key.as_str(),
                    KIND_KEY | PARA_ID | PPR_INS | PPR_DEL | PPR_CHANGE
                )
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
            if !matches!(
                key.as_str(),
                KIND_KEY | PARA_ID | PPR_INS | PPR_DEL | PPR_CHANGE
            ) {
                map.insert(txn, key.clone(), value.clone());
            }
        }
    }
}

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
    } else {
        map.insert(txn, PPR_CHANGE, Any::Array(Arc::from(remaining)));
    }
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
    let chunks = snapshot(story, txn);
    let final_pilcrow = chunks.iter().rev().find_map(|chunk| match chunk.kind {
        ChunkKind::Pilcrow(_) => Some(chunk.start),
        _ => None,
    });
    let (span_start, span_end) = span.unwrap_or((0, u32::MAX));
    let mut removed = 0;
    // Reverse walk so physical removals never shift the indices still to be visited.
    for chunk in chunks.iter().rev() {
        let overlap_start = chunk.start.max(span_start);
        let overlap_end = chunk.end().min(span_end);
        if overlap_end <= overlap_start {
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
                if join {
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
                    if Some(chunk.start) == final_pilcrow {
                        // The final paragraph mark can never be removed — clear instead.
                        let (ppr_key, attr_key) = match mode {
                            ResolveMode::Accept => (PPR_DEL, DEL),
                            ResolveMode::Reject => (PPR_INS, INS),
                        };
                        map.remove(txn, ppr_key);
                        clear_attr(txn, story, chunk.start, 1, attr_key);
                    } else {
                        story.remove_range(txn, chunk.start, 1);
                        removed += 1;
                    }
                } else {
                    match mode {
                        ResolveMode::Accept if ins_hit => {
                            record(resolved, ppr_ins.as_ref());
                            record(resolved, attr_ins.as_ref());
                            map.remove(txn, PPR_INS);
                            clear_attr(txn, story, chunk.start, 1, INS);
                        }
                        ResolveMode::Reject if del_hit => {
                            record(resolved, ppr_del.as_ref());
                            record(resolved, attr_del.as_ref());
                            map.remove(txn, PPR_DEL);
                            clear_attr(txn, story, chunk.start, 1, DEL);
                        }
                        _ => {}
                    }
                }
            }
            ChunkKind::Text(_) | ChunkKind::Embed(_) => {
                let ins = active_stamp(chunk.attrs.get(INS).cloned(), filter);
                let del = active_stamp(chunk.attrs.get(DEL).cloned(), filter);
                // A unit carrying BOTH stamps (concurrent suggest-over-suggest, case E) is
                // physically removed in either mode — matching the PM range resolver, where
                // the "remove" mark class wins over the "keep" class on the same text.
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
                } else {
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
                }
            }
        }
    }
    removed
}

impl EditingDoc {
    /// Accepts tracked changes (op-contract §1 "Resolve", S4b): pending insertions become
    /// plain content, pending deletions are carried out. See the module docs for the full
    /// 3-way matrix and join semantics. The receipt lists the resolved revision ids; a
    /// range target also echoes the surviving range.
    pub fn accept_change(&self, ctx: &EditCtx, target: &ChangeTarget) -> OpResult<Receipt> {
        self.resolve_change(ctx, target, ResolveMode::Accept)
    }

    /// Rejects tracked changes — the inverse of [`EditingDoc::accept_change`]: pending
    /// insertions are rolled back, pending deletions are restored to plain content.
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
                resolve_table_row_revisions(
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
                    Some((range.start, range.end)),
                    None,
                    &mut resolved,
                );
                let loc_range =
                    loc_range_in_txn(&range.story, &story, &txn, range.start, range.end - removed)?;
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
