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
    story: TextRef,
    base: u32,
    paragraphs: Vec<Paragraph>,
    entries: Vec<Entry>,
    /// A change moved paragraph boundaries, so `lead` is recounted from the story.
    stale: bool,
}

fn entries<T: ReadTxn>(map: &MapRef, txn: &T) -> Option<Arc<[Any]>> {
    match map.get(txn, BOOKMARKS) {
        Some(Out::Any(Any::Array(entries))) if !entries.is_empty() => Some(entries),
        _ => None,
    }
}

fn paragraph_start(base: u32, paragraphs: &[Paragraph], index: usize) -> u32 {
    match index {
        0 => base,
        _ => paragraphs[index - 1]
            .pilcrow
            .map_or(base, |pilcrow| pilcrow + 1),
    }
}

fn content_start(base: u32, paragraphs: &[Paragraph], index: usize) -> u32 {
    paragraph_start(base, paragraphs, index) + paragraphs[index].lead
}

/// The block embeds that open the paragraphs starting at `starts`, ascending.
fn leading_blocks<T: ReadTxn>(story: &TextRef, txn: &T, starts: &[u32]) -> Vec<u32> {
    let mut leads = vec![0; starts.len()];
    let mut next = 0;
    let mut counting = None;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let block = match &diff.insert {
            Out::YMap(map) => !is_pilcrow(map, txn) && block_embed(map, txn),
            _ => false,
        };
        match counting {
            Some(index) if block => leads[index] += 1,
            _ => counting = None,
        }
        while next < starts.len() && starts[next] < offset {
            next += 1;
        }
        if counting.is_none() && starts.get(next) == Some(&offset) {
            if block {
                leads[next] = 1;
                counting = Some(next);
            }
            next += 1;
        }
        if counting.is_none() && next == starts.len() {
            break;
        }
        offset += out_len(&diff.insert);
    }
    leads
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
            story: story.clone(),
            base,
            paragraphs,
            entries: Vec::new(),
            stale: false,
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
        if self.stale || !added.is_empty() {
            let starts: Vec<u32> = (0..self.paragraphs.len())
                .map(|index| paragraph_start(self.base, &self.paragraphs, index))
                .collect();
            let leads = leading_blocks(&self.story, &*txn, &starts);
            for (paragraph, lead) in self.paragraphs.iter_mut().zip(leads) {
                paragraph.lead = lead;
            }
        }
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

    /// The paragraph holding story position `at`, and where it starts.
    fn holder(&self, at: u32) -> Option<(usize, u32)> {
        let mut start = self.base;
        for (index, paragraph) in self.paragraphs.iter().enumerate() {
            let pilcrow = paragraph.pilcrow?;
            if at <= pilcrow {
                return Some((index, start));
            }
            start = pilcrow + 1;
        }
        None
    }

    fn insert(&mut self, at: u32, len: u32, split: bool) {
        if len == 0 {
            return;
        }
        if split {
            self.stale = true;
        } else if !self.stale
            && let Some((index, start)) = self.holder(at)
        {
            let lead = &mut self.paragraphs[index].lead;
            *lead = (*lead).min(at.saturating_sub(start));
        }
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
        let boundary = self.paragraphs.iter().any(|paragraph| {
            paragraph
                .pilcrow
                .is_some_and(|pilcrow| (from..to).contains(&pilcrow))
        });
        if boundary {
            self.stale = true;
        } else if !self.stale
            && let Some((index, start)) = self.holder(from)
        {
            let lead = &mut self.paragraphs[index].lead;
            let covered = to.min(start + *lead).saturating_sub(from.max(start));
            *lead -= covered;
        }
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

#[cfg(test)]
mod tests {
    use yrs::Transact;

    use super::*;
    use crate::{EditCtx, EditingDoc, StoryRange};

    fn offsets(doc: &EditingDoc) -> Vec<Vec<(String, u32)>> {
        let txn = doc.doc.transact();
        let story = crate::story_ref(&txn, "body").unwrap();
        let mut paragraphs = Vec::new();
        for diff in story.diff(&txn, YChange::identity) {
            match &diff.insert {
                Out::YMap(map) if is_pilcrow(map, &txn) => {
                    let values = entries(map, &txn).unwrap_or_else(|| Arc::from([]));
                    paragraphs.push(values.iter().filter_map(offset).collect());
                }
                _ => {}
            }
        }
        paragraphs
    }

    fn offset(value: &Any) -> Option<(String, u32)> {
        match value {
            Any::Map(value) => match (value.get("kind"), value.get("offset")) {
                (Some(Any::String(kind)), Some(offset)) => {
                    Some((kind.to_string(), u32::try_from(offset.clone()).ok()?))
                }
                _ => None,
            },
            _ => None,
        }
    }

    #[test]
    fn deleting_a_leading_table_by_replacement_keeps_offsets() {
        let table = r#"<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="20000001"><w:r><w:t>T</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
        let paragraph = r#"<w:p w14:paraId="10000001"><w:r><w:t>ab</w:t></w:r><w:bookmarkStart w:id="1" w:name="mark"/><w:r><w:t>cd</w:t></w:r><w:bookmarkEnd w:id="1"/><w:r><w:t>ef</w:t></w:r></w:p>"#;
        let body = format!("{table}{table}{paragraph}");
        let bytes = crate::seed::fixture::Package::new(&body).bytes();
        let doc = EditingDoc::new(41);
        crate::seed_from_docx(&doc, &bytes).unwrap();
        let marked = vec![vec![("start".to_owned(), 2), ("end".to_owned(), 4)]];
        assert_eq!(offsets(&doc), marked);
        let ctx = EditCtx::local(String::new(), String::new());
        doc.replace_range(&ctx, StoryRange::new("body", 0, 1), "")
            .unwrap();
        assert_eq!(offsets(&doc), marked);
    }
}
