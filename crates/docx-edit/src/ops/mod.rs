//! Mutating editing operations.

pub(crate) mod content_control;
pub mod embed;
pub mod paragraph;
pub mod resolve;
pub mod table;
pub mod text;

use std::collections::BTreeMap;

use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::{
    DEL, INS, KIND_KEY, PARA_ID, PPR_CHANGE, PPR_DEL, PPR_INS, StoryRange, is_pilcrow, map_string,
    out_len,
};

/// One formatting-run chunk of a story, snapshotted for index-stable reverse walks.
pub(crate) struct Chunk {
    pub start: u32,
    pub len: u32,
    pub kind: ChunkKind,
    pub attrs: BTreeMap<String, Any>,
}

pub(crate) enum ChunkKind {
    Text(String),
    Pilcrow(MapRef),
    Embed(Option<MapRef>),
}

impl Chunk {
    pub fn end(&self) -> u32 {
        self.start + self.len
    }

    pub fn attr_active(&self, key: &str) -> bool {
        matches!(self.attrs.get(key), Some(value) if *value != Any::Null)
    }

    pub fn is_block_embed<T: ReadTxn>(&self, txn: &T) -> bool {
        matches!(&self.kind, ChunkKind::Embed(Some(map)) if map_string(map, txn, KIND_KEY)
            .is_some_and(|kind| crate::segments::is_block_embed(&kind)))
    }

    pub fn block_revisions<T: ReadTxn>(&self, txn: &T) -> Option<[Option<Any>; 2]> {
        if !self.is_block_embed(txn) {
            return None;
        }
        let table = match &self.kind {
            ChunkKind::Embed(Some(map))
                if map_string(map, txn, KIND_KEY).as_deref() == Some("table") =>
            {
                table::table_revisions(map, txn)
            }
            _ => [None, None],
        };
        Some([INS, DEL].map(|key| {
            self.attrs
                .get(key)
                .filter(|stamp| **stamp != Any::Null)
                .cloned()
                .or_else(|| table[usize::from(key == DEL)].clone())
        }))
    }

    /// The `author` of an `ins`/`del` revision value on this chunk, if any.
    pub fn revision_author(&self, key: &str) -> Option<String> {
        let Some(Any::Map(revision)) = self.attrs.get(key) else {
            return None;
        };
        match revision.get("author") {
            Some(Any::String(author)) => Some(author.to_string()),
            _ => None,
        }
    }
}

pub(crate) fn revision_id_for_author(value: &Any, author: &str) -> Option<String> {
    let Any::Map(revision) = value else {
        return None;
    };
    if !matches!(revision.get("author"), Some(Any::String(value)) if value.as_ref() == author) {
        return None;
    }
    match revision.get("id").or_else(|| revision.get("revisionId")) {
        Some(Any::String(id)) => Some(id.to_string()),
        Some(Any::Number(id)) if id.is_finite() => Some(id.to_string()),
        Some(Any::BigInt(id)) => Some(id.to_string()),
        _ => None,
    }
}

/// Reuse a same-author revision touching an edit boundary. This models Word's
/// continuous suggestion run: separately dispatched keystrokes and paragraph
/// breaks stay one revision while their stamped units remain adjacent.
pub(crate) fn adjacent_revision_id(
    chunks: &[Chunk],
    index: u32,
    key: &str,
    author: &str,
) -> Option<String> {
    chunks
        .iter()
        .rev()
        .filter(|chunk| chunk.end() == index)
        .chain(chunks.iter().filter(|chunk| chunk.start == index))
        .find_map(|chunk| {
            chunk
                .attrs
                .get(key)
                .and_then(|value| revision_id_for_author(value, author))
        })
}

/// Reuse a same-author paragraph-property revision at an edit boundary.
///
/// A list command commonly authors `pPrChange` on an empty paragraph before
/// the first character is typed. Treat that pending property change as the
/// start of the same continuous suggestion run so the list formatting, text,
/// and subsequent paragraph breaks resolve from one sidebar card.
pub(crate) fn adjacent_paragraph_change_revision_id<T: ReadTxn>(
    chunks: &[Chunk],
    index: u32,
    txn: &T,
    author: &str,
) -> Option<String> {
    chunks
        .iter()
        .rev()
        .filter(|chunk| chunk.end() == index)
        .chain(chunks.iter().filter(|chunk| chunk.start == index))
        .find_map(|chunk| {
            let ChunkKind::Pilcrow(map) = &chunk.kind else {
                return None;
            };
            let Some(Out::Any(Any::Array(changes))) = map.get(txn, PPR_CHANGE) else {
                return None;
            };
            changes.iter().rev().find_map(|change| {
                let Any::Map(change) = change else {
                    return None;
                };
                change
                    .get("info")
                    .and_then(|info| revision_id_for_author(info, author))
            })
        })
}

pub(crate) fn revision_id_in_range(
    chunks: &[Chunk],
    start: u32,
    end: u32,
    key: &str,
    author: &str,
) -> Option<String> {
    chunks
        .iter()
        .filter(|chunk| chunk.end() > start && chunk.start < end)
        .find_map(|chunk| {
            chunk
                .attrs
                .get(key)
                .and_then(|value| revision_id_for_author(value, author))
        })
}

fn chunk_at<T: ReadTxn>(insert: &Out, attrs: Option<&Attrs>, offset: u32, txn: &T) -> Chunk {
    let kind = match insert {
        Out::Any(Any::String(value)) => ChunkKind::Text(value.to_string()),
        Out::YMap(map) if is_pilcrow(map, txn) => ChunkKind::Pilcrow(map.clone()),
        Out::YMap(map) => ChunkKind::Embed(Some(map.clone())),
        _ => ChunkKind::Embed(None),
    };
    let attrs = attrs
        .into_iter()
        .flat_map(|attrs| attrs.iter())
        .map(|(key, value)| (key.to_string(), value.clone()))
        .collect();
    Chunk {
        start: offset,
        len: out_len(insert),
        kind,
        attrs,
    }
}

pub(crate) fn snapshot<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<Chunk> {
    let mut offset = 0;
    story
        .diff(txn, YChange::identity)
        .into_iter()
        .map(|diff| {
            let chunk = chunk_at(&diff.insert, diff.attributes.as_deref(), offset, txn);
            offset += chunk.len;
            chunk
        })
        .collect()
}

/// `snapshot` over the chunks overlapping `[lo, hi)`.
pub(crate) fn snapshot_range<T: ReadTxn>(story: &TextRef, txn: &T, lo: u32, hi: u32) -> Vec<Chunk> {
    let mut offset = 0;
    let mut chunks = Vec::new();
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if offset >= hi {
            break;
        }
        if offset + len > lo {
            chunks.push(chunk_at(
                &diff.insert,
                diff.attributes.as_deref(),
                offset,
                txn,
            ));
        }
        offset += len;
    }
    chunks
}

pub(crate) fn block_embed_at<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> bool {
    snapshot_range(story, txn, index, index.saturating_add(1))
        .first()
        .is_some_and(|chunk| chunk.is_block_embed(txn))
}

pub(crate) fn inherit_block_revisions(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    index: u32,
    map: &MapRef,
    revisions: &[Option<Any>; 2],
) {
    for ((attr_key, ppr_key), stamp) in [(INS, PPR_INS), (DEL, PPR_DEL)].into_iter().zip(revisions)
    {
        if let Some(stamp) = stamp {
            map.insert(txn, ppr_key, stamp.clone());
            story.format(
                txn,
                index,
                1,
                Attrs::from([(attr_key.into(), stamp.clone())]),
            );
        }
    }
}

pub(crate) fn paragraph_content_before<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> bool {
    index > 0
        && snapshot_range(story, txn, index - 1, index)
            .last()
            .is_some_and(|chunk| {
                !matches!(chunk.kind, ChunkKind::Pilcrow(_)) && !chunk.is_block_embed(txn)
            })
}

/// The story's last pilcrow as `(index, map)`.
pub(crate) fn last_pilcrow<T: ReadTxn>(story: &TextRef, txn: &T) -> Option<(u32, MapRef)> {
    let mut offset = 0;
    let mut last = None;
    for diff in story.diff(txn, YChange::identity) {
        if let Out::YMap(map) = &diff.insert
            && is_pilcrow(map, txn)
        {
            last = Some((offset, map.clone()));
        }
        offset += out_len(&diff.insert);
    }
    last
}

/// Captures every pilcrow property except the schema discriminator, plus the paraId.
pub(crate) fn capture_pilcrow<T: ReadTxn>(map: &MapRef, txn: &T) -> (String, Vec<(String, Any)>) {
    let para_id = map_string(map, txn, PARA_ID).unwrap_or_default();
    let props = map
        .iter(txn)
        .filter_map(|(key, value)| {
            if matches!(key, KIND_KEY | PARA_ID) {
                return None;
            }
            let Out::Any(value) = value else {
                return None;
            };
            Some((key.to_string(), value))
        })
        .collect();
    (para_id, props)
}

/// Replaces a pilcrow's ID and properties with the donor's.
pub(crate) fn adopt_pilcrow(
    txn: &mut TransactionMut<'_>,
    survivor: &MapRef,
    donor_para_id: &str,
    donor_props: &[(String, Any)],
) {
    let existing: Vec<String> = survivor
        .iter(txn)
        .map(|(key, _)| key.to_string())
        .filter(|key| key != KIND_KEY)
        .collect();
    for key in existing {
        survivor.remove(txn, &key);
    }
    survivor.insert(txn, PARA_ID, donor_para_id);
    for (key, value) in donor_props {
        survivor.insert(txn, key.clone(), value.clone());
    }
}

pub(crate) fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}

/// Whether UTF-16 `offset` falls between the two units of a surrogate pair in `text`.
fn inside_pair(text: &str, offset: u32) -> bool {
    let mut at = 0;
    for ch in text.chars() {
        if at >= offset {
            return false;
        }
        at += ch.len_utf16() as u32;
        if at > offset {
            return true;
        }
    }
    false
}

/// `index` moved off the middle of a surrogate pair in `chunks`: back to the pair's start, or
/// past its end when `forward`. yrs splits such an index after the pair while the new item's
/// id assumes the index itself, which corrupts the story.
fn code_point_index(chunks: &[Chunk], index: u32, forward: bool) -> u32 {
    let inside = chunks.iter().any(|chunk| {
        matches!(&chunk.kind, ChunkKind::Text(text)
            if chunk.start < index && inside_pair(text, index - chunk.start))
    });
    match (inside, forward) {
        (false, _) => index,
        (true, false) => index - 1,
        (true, true) => index + 1,
    }
}

/// Chunks covering the units on both sides of `index`, after moving `index` off the middle of
/// a surrogate pair to the pair's start.
pub(crate) fn position_chunks<T: ReadTxn>(story: &TextRef, txn: &T, index: &mut u32) -> Vec<Chunk> {
    let take =
        |index: u32| snapshot_range(story, txn, index.saturating_sub(1), index.saturating_add(1));
    let chunks = take(*index);
    let snapped = code_point_index(&chunks, *index, false);
    if snapped == *index {
        return chunks;
    }
    *index = snapped;
    take(snapped)
}

/// Chunks covering `range` and the unit on each side, after widening `range` to whole code
/// points (a collapsed range moves to the pair's start).
pub(crate) fn range_chunks<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    range: &mut StoryRange,
) -> Vec<Chunk> {
    let take = |range: &StoryRange| {
        snapshot_range(
            story,
            txn,
            range.start.saturating_sub(1),
            range.end.saturating_add(1),
        )
    };
    let chunks = take(range);
    let start = code_point_index(&chunks, range.start, false);
    let end = if range.start == range.end {
        start
    } else {
        code_point_index(&chunks, range.end, true)
    };
    if (start, end) == (range.start, range.end) {
        return chunks;
    }
    range.start = start;
    range.end = end;
    take(range)
}

/// `[start, end)` widened to whole code points, for ops that take no chunk snapshot; a
/// collapsed range moves to the pair's start.
pub(crate) fn code_point_range<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    start: u32,
    end: u32,
) -> (u32, u32) {
    let mut range = (start, end);
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        if offset >= end {
            break;
        }
        if let Out::Any(Any::String(text)) = &diff.insert {
            if offset < start && inside_pair(text, start - offset) {
                range.0 = start - 1;
            }
            if start < end && inside_pair(text, end - offset) {
                range.1 = end + 1;
            }
        }
        offset += out_len(&diff.insert);
    }
    if start == end {
        range.1 = range.0;
    }
    range
}
