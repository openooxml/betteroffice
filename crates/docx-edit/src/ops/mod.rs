//! The S1 mutating op surface (op-contract §1), split across `text` and `paragraph`.

pub mod embed;
pub mod paragraph;
pub mod resolve;
pub mod table;
pub mod text;

use std::collections::BTreeMap;

use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::{KIND_KEY, PARA_ID, PPR_CHANGE, is_pilcrow, map_string, out_len};

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

fn revision_id_for_author(value: &Any, author: &str) -> Option<String> {
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

pub(crate) fn snapshot<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<Chunk> {
    let mut offset = 0;
    story
        .diff(txn, YChange::identity)
        .into_iter()
        .map(|diff| {
            let len = out_len(&diff.insert);
            let kind = match &diff.insert {
                Out::Any(Any::String(value)) => ChunkKind::Text(value.to_string()),
                Out::YMap(map) if is_pilcrow(map, txn) => ChunkKind::Pilcrow(map.clone()),
                Out::YMap(map) => ChunkKind::Embed(Some(map.clone())),
                _ => ChunkKind::Embed(None),
            };
            let attrs = diff
                .attributes
                .as_deref()
                .into_iter()
                .flat_map(|attrs| attrs.iter())
                .map(|(key, value)| (key.to_string(), value.clone()))
                .collect();
            let chunk = Chunk {
                start: offset,
                len,
                kind,
                attrs,
            };
            offset += len;
            chunk
        })
        .collect()
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

/// Replaces a pilcrow's paraId and full property set with the donor's (op-contract R6: the
/// surviving paragraph adopts the FIRST affected paragraph's pPr + paraId).
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
