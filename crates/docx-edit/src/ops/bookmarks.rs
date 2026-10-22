//! Bookmark offsets on paragraph marks, carried through the edits that move the text around
//! them.
//!
//! A mark's `bookmarks` entries hold offsets into its paragraph's inline content. An edit
//! captures the entries of the paragraphs it touches as story positions, maps them through the
//! units it inserted and removed, and writes each entry back to the paragraph that now holds it.
//! As in Word, text typed at a range's start or inside it extends it, text typed at its end
//! stays outside, a collapsed bookmark keeps the text after it, removed text collapses what it
//! covered, and a paragraph split at a range's start moves the start into the new paragraph.

use std::collections::HashMap;
use std::sync::Arc;

use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::ops::{IndexedStory, past_leading_blocks_indexed};
use crate::{KIND_KEY, is_pilcrow, map_string, out_len};

pub(crate) const BOOKMARKS: &str = "bookmarks";

/// One unit-level change, in story positions at the time it applies.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Change {
    Insert {
        at: u32,
        len: u32,
    },
    /// A paragraph mark inserted at `at`.
    Split {
        at: u32,
    },
    Remove {
        from: u32,
        to: u32,
    },
}

#[derive(Clone, Copy, PartialEq)]
enum Kind {
    Start,
    End,
    /// An entry without a kind spans its whole paragraph and follows the mark.
    Paragraph,
}

struct Entry {
    at: u32,
    kind: Kind,
    id: Option<Any>,
    value: Arc<HashMap<String, Any>>,
}

struct Paragraph {
    pilcrow: Option<u32>,
    lead: u32,
    map: MapRef,
    added: bool,
}

/// The bookmark entries of the paragraphs one edit touches.
pub(crate) struct Carry {
    base: u32,
    paragraphs: Vec<Paragraph>,
    entries: Vec<Entry>,
}

fn entries<T: ReadTxn>(map: &MapRef, txn: &T) -> Option<Arc<[Any]>> {
    match map.get(txn, BOOKMARKS) {
        Some(Out::Any(Any::Array(entries))) if !entries.is_empty() => Some(entries),
        _ => None,
    }
}

fn content_start(base: u32, paragraphs: &[Paragraph], index: usize) -> u32 {
    let start = match index {
        0 => base,
        _ => paragraphs[index - 1]
            .pilcrow
            .map_or(base, |pilcrow| pilcrow + 1),
    };
    start + paragraphs[index].lead
}

fn kind(value: &HashMap<String, Any>) -> Kind {
    match value.get("kind") {
        Some(Any::String(kind)) if kind.as_ref() == "start" => Kind::Start,
        Some(Any::String(kind)) if kind.as_ref() == "end" => Kind::End,
        _ => Kind::Paragraph,
    }
}

fn block_embed<T: ReadTxn>(map: &MapRef, txn: &T) -> bool {
    map_string(map, txn, KIND_KEY).is_some_and(|kind| crate::segments::is_block_embed(&kind))
}

impl Carry {
    /// Captures the bookmarks of the paragraphs holding `[lo, hi]`, or `None` when they hold
    /// none.
    pub(crate) fn capture<T: ReadTxn>(
        story: &TextRef,
        txn: &T,
        lo: u32,
        hi: u32,
        indexed: Option<&IndexedStory>,
    ) -> Option<Self> {
        let (base, paragraphs) = match indexed.and_then(|index| index.pilcrow_at(txn, lo)) {
            Some((bounds, map)) if hi <= bounds.pilcrow => {
                entries(&map, txn)?;
                let content = past_leading_blocks_indexed(story, txn, bounds.start, indexed);
                let paragraph = Paragraph {
                    pilcrow: Some(bounds.pilcrow),
                    lead: content - bounds.start,
                    map,
                    added: false,
                };
                (bounds.start, vec![paragraph])
            }
            _ => Self::scan(story, txn, lo, hi)?,
        };
        let mut carry = Self {
            base,
            paragraphs,
            entries: Vec::new(),
        };
        for (index, paragraph) in carry.paragraphs.iter().enumerate() {
            let Some(values) = entries(&paragraph.map, txn) else {
                continue;
            };
            let pilcrow = paragraph.pilcrow.unwrap_or_default();
            let content = content_start(carry.base, &carry.paragraphs, index).min(pilcrow);
            for value in values.iter() {
                let Any::Map(value) = value else { continue };
                let offset = value
                    .get("offset")
                    .and_then(|offset| u32::try_from(offset.clone()).ok())
                    .unwrap_or_default();
                let kind = kind(value);
                carry.entries.push(Entry {
                    at: match kind {
                        Kind::Paragraph => pilcrow,
                        _ => content.saturating_add(offset).min(pilcrow),
                    },
                    kind,
                    id: value.get("id").cloned(),
                    value: value.clone(),
                });
            }
        }
        (!carry.entries.is_empty()).then_some(carry)
    }

    fn scan<T: ReadTxn>(
        story: &TextRef,
        txn: &T,
        lo: u32,
        hi: u32,
    ) -> Option<(u32, Vec<Paragraph>)> {
        let mut offset = 0;
        let mut start = 0;
        let mut lead = 0;
        let mut leading = true;
        let mut base = None;
        let mut paragraphs = Vec::new();
        for diff in story.diff(txn, YChange::identity) {
            match &diff.insert {
                Out::YMap(map) if is_pilcrow(map, txn) => {
                    if offset >= lo {
                        base.get_or_insert(start);
                        paragraphs.push(Paragraph {
                            pilcrow: Some(offset),
                            lead,
                            map: map.clone(),
                            added: false,
                        });
                        if offset >= hi {
                            break;
                        }
                    }
                    start = offset + 1;
                    lead = 0;
                    leading = true;
                }
                Out::YMap(map) if leading && block_embed(map, txn) => lead += 1,
                _ => leading = false,
            }
            offset += out_len(&diff.insert);
        }
        paragraphs
            .iter()
            .any(|paragraph| entries(&paragraph.map, txn).is_some())
            .then(|| (base.unwrap_or_default(), paragraphs))
    }

    /// [`Carry::apply`] for `len` units inserted at `at`.
    pub(crate) fn inserted(self, txn: &mut TransactionMut<'_>, at: u32, len: u32) {
        self.apply(txn, &[Change::Insert { at, len }], &[]);
    }

    /// [`Carry::apply`] for the units `[from, to)` removed.
    pub(crate) fn removed(self, txn: &mut TransactionMut<'_>, from: u32, to: u32) {
        self.apply(txn, &[Change::Remove { from, to }], &[]);
    }

    /// Maps the captured entries through `changes` and writes them to the paragraphs that hold
    /// them now; `added` names marks the edit inserted, at their final positions.
    pub(crate) fn apply(
        mut self,
        txn: &mut TransactionMut<'_>,
        changes: &[Change],
        added: &[(u32, MapRef)],
    ) {
        for change in changes {
            self.map_through(*change);
        }
        let captured: Vec<(Option<Any>, Kind)> = self
            .entries
            .iter()
            .map(|entry| (entry.id.clone(), entry.kind))
            .collect();
        self.paragraphs
            .retain(|paragraph| paragraph.pilcrow.is_some());
        self.paragraphs
            .extend(added.iter().map(|(at, map)| Paragraph {
                pilcrow: Some(*at),
                lead: 0,
                map: map.clone(),
                added: true,
            }));
        self.paragraphs.sort_by_key(|paragraph| paragraph.pilcrow);
        let Some(last) = self.paragraphs.len().checked_sub(1) else {
            return;
        };
        let mut owned: Vec<Vec<Any>> = vec![Vec::new(); self.paragraphs.len()];
        for entry in &self.entries {
            let owner = self
                .paragraphs
                .iter()
                .position(|paragraph| paragraph.pilcrow.is_some_and(|at| at >= entry.at))
                .unwrap_or(last);
            let pilcrow = self.paragraphs[owner].pilcrow.unwrap_or_default();
            let content = content_start(self.base, &self.paragraphs, owner).min(pilcrow);
            let value = if entry.kind == Kind::Paragraph {
                Any::Map(entry.value.clone())
            } else {
                let mut value = (*entry.value).clone();
                value.insert(
                    "offset".to_owned(),
                    Any::from(entry.at.clamp(content, pilcrow) - content),
                );
                Any::Map(Arc::new(value))
            };
            owned[owner].push(value);
        }
        for (paragraph, mut values) in self.paragraphs.iter().zip(owned) {
            let current = entries(&paragraph.map, txn);
            if paragraph.added
                && let Some(current) = current.as_ref()
            {
                let foreign = current.iter().filter(|value| {
                    let Any::Map(value) = value else { return true };
                    !captured.contains(&(value.get("id").cloned(), kind(value)))
                });
                values.splice(0..0, foreign.cloned());
            }
            if current.as_deref() == Some(values.as_slice()) {
                continue;
            }
            if values.is_empty() {
                paragraph.map.remove(txn, BOOKMARKS);
            } else {
                paragraph
                    .map
                    .insert(txn, BOOKMARKS, Any::Array(Arc::from(values)));
            }
        }
    }

    fn map_through(&mut self, change: Change) {
        match change {
            Change::Insert { at, len } => self.insert(at, len, false),
            Change::Split { at } => self.insert(at, 1, true),
            Change::Remove { from, to } => self.remove(from, to),
        }
    }

    fn insert(&mut self, at: u32, len: u32, split: bool) {
        let shifted: Vec<bool> = self
            .entries
            .iter()
            .map(|entry| {
                entry.at > at
                    || entry.at == at
                        && match entry.kind {
                            Kind::Start => {
                                split
                                    && !self.entries.iter().any(|end| {
                                        end.kind == Kind::End && end.id == entry.id && end.at == at
                                    })
                            }
                            Kind::End => false,
                            Kind::Paragraph => true,
                        }
            })
            .collect();
        for (entry, shifted) in self.entries.iter_mut().zip(shifted) {
            if shifted {
                entry.at += len;
            }
        }
        for paragraph in &mut self.paragraphs {
            if let Some(pilcrow) = paragraph.pilcrow.as_mut()
                && *pilcrow >= at
            {
                *pilcrow += len;
            }
        }
    }

    fn remove(&mut self, from: u32, to: u32) {
        let map = |at: u32| match at {
            at if at <= from => at,
            at if at >= to => at - (to - from),
            _ => from,
        };
        for entry in &mut self.entries {
            entry.at = map(entry.at);
        }
        for paragraph in &mut self.paragraphs {
            paragraph.pilcrow = paragraph
                .pilcrow
                .filter(|pilcrow| !(from..to).contains(pilcrow))
                .map(map);
        }
    }
}
