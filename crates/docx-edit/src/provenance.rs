use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex, MutexGuard};

use docx_parse::paragraph_patch::DonorV1;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use yrs::branch::{Branch, BranchID};
use yrs::types::text::YChange;
use yrs::types::{DeepObservable, Delta, EntryChange, Event, Events, PathSegment};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Assoc, ClientID, Doc, ID, IndexedSequence, Observable, Out, ReadTxn, StateVector,
    StickyIndex, Subscription, Text, TextRef, Transact, TransactionMut,
};

use crate::identity::{GENERATION, SESSION};
use crate::op::ParaBounds;
use crate::ops::{Chunk, ChunkKind, IndexedStory};
use crate::{EditError, EditResult, EditingDoc, PARA_ID, STORIES, UpdateOrigin};

const OCCURRENCE_BYTES: usize = 8 * 1024;
const OPENING_BYTES: usize = 256 * 1024;
const VERSION: u8 = 2;
const CHECKSUM_BYTES: usize = 32;

#[derive(Clone, Copy, Debug, PartialEq, Deserialize, Serialize)]
struct UnitId(u64, u32);

impl From<ID> for UnitId {
    fn from(id: ID) -> Self {
        Self(id.client.get(), id.clock)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Deserialize, Serialize)]
struct Span {
    first: UnitId,
    len: u32,
}

impl Span {
    fn end(self) -> Option<u32> {
        self.first.1.checked_add(self.len)
    }

    fn offset(self, id: UnitId) -> Option<u32> {
        if id.0 != self.first.0 || id.1 >= self.end()? {
            return None;
        }
        id.1.checked_sub(self.first.1)
    }

    fn intersects(self, other: Self) -> bool {
        self.len > 0
            && other.len > 0
            && self.first.0 == other.first.0
            && self.end().is_some_and(|end| other.first.1 < end)
            && other.end().is_some_and(|end| self.first.1 < end)
    }

    fn matches<T: ReadTxn>(self, story: &TextRef, txn: &T, at: u32) -> bool {
        if self.end().is_none()
            || at
                .checked_add(self.len)
                .is_none_or(|end| end > story.len(txn))
        {
            return false;
        }
        (0..self.len)
            .all(|i| unit_id(story, txn, at + i) == Some(UnitId(self.first.0, self.first.1 + i)))
    }

    fn slice(self, from: u32, to: u32) -> Option<Self> {
        (to <= self.len).then_some(())?;
        Some(Self {
            first: UnitId(self.first.0, self.first.1.checked_add(from)?),
            len: to.checked_sub(from)?,
        })
    }
}

#[derive(Clone, Debug, PartialEq, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct Splice {
    left: Option<UnitId>,
    removed: Option<Span>,
    inserted: Vec<Span>,
    donor: DonorV1,
}

impl Splice {
    fn range(&self, seed: Span) -> Option<(u32, u32)> {
        if self.removed.is_none() && self.inserted.is_empty() {
            return None;
        }
        let from = match self.left {
            Some(left) => seed.offset(left)?.checked_add(1)?,
            None => 0,
        };
        let to = match self.removed {
            Some(removed) if removed.len != 0 && seed.offset(removed.first)? == from => {
                from.checked_add(removed.len)?
            }
            Some(_) => return None,
            None => from,
        };
        (to <= seed.len).then_some((from, to))
    }

    fn conflicts(&self, other: &Self, seed: Span) -> bool {
        let Some(((from, to), (other_from, other_to))) = self.range(seed).zip(other.range(seed))
        else {
            return true;
        };
        from == other_from
            || (from < other_to && other_from < to)
            || (from == to && from == other_to && other.inserted.is_empty())
            || (other_from == other_to && other_from == to && self.inserted.is_empty())
    }
}

#[derive(Clone, Debug, PartialEq, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct Records {
    story: String,
    seed: Span,
    splices: Vec<Splice>,
}

#[derive(Clone, Debug, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
enum Entry {
    Unproven,
    HistoryReplacement,
    Recorded(Records),
}

impl Entry {
    fn records(&self) -> Option<&Records> {
        match self {
            Self::Recorded(records) => Some(records),
            _ => None,
        }
    }
}

#[derive(Clone, Default)]
pub(crate) struct Provenance {
    generation: Option<String>,
    entries: BTreeMap<String, Entry>,
    bytes: usize,
    overflowed: bool,
    outgoing: Option<Outgoing>,
}

#[derive(Clone)]
struct Outgoing {
    epoch: u64,
    base: Vec<u8>,
    digest: [u8; CHECKSUM_BYTES],
}

#[derive(Serialize)]
struct Sidecar<'a> {
    generation: &'a str,
    base: Option<&'a [u8]>,
    digest: [u8; CHECKSUM_BYTES],
    state_digest: Option<[u8; CHECKSUM_BYTES]>,
    overflowed: bool,
    entries: &'a BTreeMap<String, Entry>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Incoming {
    generation: String,
    base: Option<Vec<u8>>,
    digest: [u8; CHECKSUM_BYTES],
    state_digest: Option<[u8; CHECKSUM_BYTES]>,
    overflowed: bool,
    entries: BTreeMap<String, Entry>,
}

fn encoded_entry_size(key: &str, entry: &Entry) -> usize {
    let size = || {
        let key = serde_json::to_vec(key).ok()?;
        let entry = serde_json::to_vec(entry).ok()?;
        key.len().checked_add(entry.len())?.checked_add(2)
    };
    size().unwrap_or(usize::MAX)
}

fn lock_provenance(provenance: &Mutex<Provenance>) -> MutexGuard<'_, Provenance> {
    provenance.lock().unwrap_or_else(|poisoned| {
        let mut proof = poisoned.into_inner();
        proof.overflow();
        provenance.clear_poison();
        proof
    })
}

impl Provenance {
    pub(crate) fn overflow(&mut self) {
        self.entries.clear();
        self.generation = None;
        self.bytes = 0;
        self.overflowed = true;
        self.outgoing = None;
    }

    fn set(&mut self, key: String, mut entry: Entry) {
        if key.is_empty() || self.overflowed {
            return;
        }
        let previous = self.entries.get(&key);
        if matches!(previous, Some(Entry::HistoryReplacement))
            || matches!(previous, Some(Entry::Unproven))
                && !matches!(entry, Entry::HistoryReplacement)
        {
            return;
        }
        let mut size = encoded_entry_size(&key, &entry);
        if size > OCCURRENCE_BYTES {
            entry = Entry::Unproven;
            size = encoded_entry_size(&key, &entry);
        }
        let old_size = previous.map_or(0, |old| encoded_entry_size(&key, old));
        self.bytes = self.bytes.saturating_sub(old_size).saturating_add(size);
        if size > OCCURRENCE_BYTES
            || self
                .bytes
                .saturating_add(self.generation.as_ref().map_or(0, String::len))
                > OPENING_BYTES
        {
            self.overflow();
        } else {
            self.entries.insert(key, entry);
            self.outgoing = None;
        }
    }

    pub(crate) fn mark(&mut self, key: String) {
        self.set(key, Entry::Unproven);
    }

    pub(crate) fn has_records(&self, key: &str) -> bool {
        matches!(self.entries.get(key), Some(Entry::Recorded(_)))
    }

    pub(crate) fn equal_output_candidates(&self) -> impl Iterator<Item = (&str, bool)> {
        self.entries.iter().filter_map(|(key, entry)| match entry {
            Entry::Recorded(_) => Some((key.as_str(), false)),
            Entry::HistoryReplacement => Some((key.as_str(), true)),
            Entry::Unproven => None,
        })
    }

    pub(crate) fn unproven(&self, key: &str) -> bool {
        self.overflowed
            || matches!(
                self.entries.get(key),
                Some(Entry::Unproven | Entry::HistoryReplacement)
            )
    }

    fn invalidate_records(&mut self) {
        let keys: Vec<_> = self
            .entries
            .iter()
            .filter(|(_, entry)| entry.records().is_some())
            .map(|(key, _)| key.clone())
            .collect();
        for key in keys {
            self.mark(key);
        }
    }

    fn record(
        &mut self,
        generation: String,
        key: String,
        story: String,
        seed: Span,
        splice: Splice,
    ) {
        if self.unproven(&key) {
            return;
        }
        if self
            .generation
            .as_ref()
            .is_some_and(|old| *old != generation)
        {
            self.overflow();
            return;
        }
        if generation.len() > OPENING_BYTES {
            self.overflow();
            return;
        }
        self.generation = Some(generation);
        let mut records = self
            .entries
            .get(&key)
            .and_then(Entry::records)
            .cloned()
            .unwrap_or_else(|| Records {
                story: story.clone(),
                seed,
                splices: Vec::new(),
            });
        if splice.range(seed).is_none()
            || records.story != story
            || records.seed != seed
            || records
                .splices
                .iter()
                .any(|other| splice.conflicts(other, seed))
        {
            self.mark(key);
        } else {
            records.splices.push(splice);
            self.set(key, Entry::Recorded(records));
        }
    }

    fn merge(&mut self, incoming: Incoming) {
        if incoming.overflowed {
            self.overflow();
            return;
        }
        for (key, entry) in incoming.entries {
            if let Entry::Recorded(records) = entry {
                for splice in records.splices {
                    if self
                        .entries
                        .get(&key)
                        .and_then(Entry::records)
                        .is_some_and(|old| {
                            old.seed == records.seed
                                && old.story == records.story
                                && old.splices.contains(&splice)
                        })
                    {
                        continue;
                    }
                    self.record(
                        incoming.generation.clone(),
                        key.clone(),
                        records.story.clone(),
                        records.seed,
                        splice,
                    );
                }
            } else {
                self.set(key, entry);
            }
        }
    }
}

fn state_digest(txn: &impl ReadTxn) -> [u8; CHECKSUM_BYTES] {
    Sha256::digest(crate::deterministic::encode_state_as_update_v1(
        txn,
        &StateVector::default(),
    ))
    .into()
}

fn references_older_insertions(
    entries: &BTreeMap<String, Entry>,
    base: &StateVector,
) -> bool {
    entries.values().filter_map(Entry::records).any(|records| {
        records.splices.iter().any(|splice| {
            splice
                .inserted
                .iter()
                .any(|span| span.first.1 < base.get(&ClientID::new(span.first.0)))
        })
    })
}

fn generation<T: ReadTxn>(txn: &T) -> Option<String> {
    crate::map_string(&txn.get_map(SESSION)?, txn, GENERATION)
}

fn unit_id<T: ReadTxn>(story: &TextRef, txn: &T, at: u32) -> Option<UnitId> {
    story
        .sticky_index(txn, at, Assoc::After)?
        .id()
        .copied()
        .map(UnitId::from)
}

fn seeded_span<T: ReadTxn>(
    txn: &T,
    story: &TextRef,
    bounds: &ParaBounds,
    mark: UnitId,
    (clock, len): (u32, u32),
    recorded: Option<Span>,
) -> Option<Span> {
    let BranchID::Nested(origin) = AsRef::<Branch>::as_ref(story).id() else {
        return None;
    };
    let client = origin.client.get();
    let seeded = Span {
        first: UnitId(client, clock),
        len,
    };
    if origin.clock >= clock || seeded.offset(mark).is_none() {
        return None;
    }
    if let Some(recorded) = recorded {
        seeded.offset(recorded.first)?;
        return (recorded.end()? == mark.1).then_some(recorded);
    }
    let first = unit_id(story, txn, bounds.start)?;
    seeded.offset(first)?;
    Some(Span {
        first,
        len: mark.1.checked_sub(first.1)?,
    })
}

pub(crate) struct PendingSplice {
    key: String,
    story: String,
    generation: String,
    seed: Span,
    splice: Splice,
}

impl PendingSplice {
    pub(crate) fn capture(
        doc: &EditingDoc,
        txn: &TransactionMut<'_>,
        story_id: &str,
        story: &TextRef,
        start: u32,
        end: u32,
        chunks: &[Chunk],
        indexed: Option<&IndexedStory>,
        inheriting: bool,
    ) -> Option<Self> {
        let seed_identity = doc.source_metadata()?.seed_identity(story_id)?;
        let generation = generation(txn)?;
        let bounds = indexed
            .and_then(|index| index.bounds_at(txn, start))
            .or_else(|| {
                crate::op::para_bounds(story, txn)
                    .into_iter()
                    .find(|bounds| bounds.start <= start && start <= bounds.pilcrow)
            });
        let Some(mut bounds) = bounds else {
            return None;
        };
        bounds.start = crate::ops::past_leading_blocks_indexed(story, txn, bounds.start, indexed);
        if !inheriting || end > bounds.pilcrow || start < bounds.start || end < start {
            let to = if end > bounds.pilcrow && inheriting {
                end
            } else {
                end.saturating_sub(1).max(start)
            };
            doc.mark_provenance_range(txn, story, start, to);
            return None;
        }
        let recorded = {
            let proof = lock_provenance(&doc.provenance);
            if proof.unproven(&bounds.para_id) {
                return None;
            }
            proof
                .entries
                .get(&bounds.para_id)
                .and_then(Entry::records)
                .map(|records| records.seed)
        };
        let captured = unit_id(story, txn, bounds.pilcrow)
            .and_then(|mark| seeded_span(txn, story, &bounds, mark, seed_identity, recorded))
            .and_then(|seed| {
                let splice = capture_splice(txn, story, &bounds, start, end, chunks, seed)?;
                Some((seed, splice))
            });
        if let Some((seed, splice)) = captured {
            Some(Self {
                key: bounds.para_id,
                story: story_id.to_owned(),
                generation,
                seed,
                splice,
            })
        } else {
            lock_provenance(&doc.provenance).mark(bounds.para_id);
            None
        }
    }

    pub(crate) fn finish(
        mut self,
        doc: &EditingDoc,
        txn: &TransactionMut<'_>,
        story: &TextRef,
        from: u32,
        to: u32,
    ) {
        let client = txn.doc().client_id().get();
        for at in from..to {
            let Some(id) = unit_id(story, txn, at).filter(|id| id.0 == client) else {
                lock_provenance(&doc.provenance).mark(self.key);
                return;
            };
            if let Some(last) = self.splice.inserted.last_mut()
                && last.first.0 == id.0
                && last.end() == Some(id.1)
            {
                let Some(len) = last.len.checked_add(1) else {
                    lock_provenance(&doc.provenance).mark(self.key);
                    return;
                };
                last.len = len;
            } else {
                self.splice.inserted.push(Span { first: id, len: 1 });
                if self.splice.inserted.len() > OCCURRENCE_BYTES / 16 {
                    lock_provenance(&doc.provenance).mark(self.key);
                    return;
                }
            }
        }
        lock_provenance(&doc.provenance).record(
            self.generation,
            self.key,
            self.story,
            self.seed,
            self.splice,
        );
    }
}

fn capture_splice(
    txn: &TransactionMut<'_>,
    story: &TextRef,
    bounds: &ParaBounds,
    start: u32,
    end: u32,
    chunks: &[Chunk],
    seed: Span,
) -> Option<Splice> {
    let (from, removed, donor) = if start < end {
        let first = unit_id(story, txn, start)?;
        let from = seed.offset(first)?;
        let removed = Span {
            first,
            len: end - start,
        };
        if removed.end()? > seed.end()? || !removed.matches(story, txn, start) {
            return None;
        }
        let donor = chunks
            .iter()
            .find_map(|chunk| {
                let at = chunk.start.max(start);
                (at < end.min(chunk.end()) && matches!(chunk.kind, ChunkKind::Text(_)))
                    .then_some(at)
            })
            .map(|at| {
                seed.offset(unit_id(story, txn, at)?)
                    .map(|at| DonorV1::At { at })
            })
            .unwrap_or(Some(DonorV1::Pilcrow))?;
        (from, Some(removed), donor)
    } else if start == bounds.start {
        if start == bounds.pilcrow {
            if seed.len != 0 {
                return None;
            }
            (0, None, DonorV1::Pilcrow)
        } else {
            let at = seed.offset(unit_id(story, txn, start)?)?;
            if at != 0 {
                return None;
            }
            (0, None, DonorV1::At { at })
        }
    } else {
        let from = seed
            .offset(unit_id(story, txn, start.checked_sub(1)?)?)?
            .checked_add(1)?;
        let (_, width) = crate::ops::code_point_boundary(chunks, start, false);
        let donor = if width == 0 {
            DonorV1::Pilcrow
        } else {
            DonorV1::At {
                at: from.checked_sub(width)?,
            }
        };
        (from, None, donor)
    };
    let left = if let Some(offset) = from.checked_sub(1) {
        Some(UnitId(seed.first.0, seed.first.1.checked_add(offset)?))
    } else {
        None
    };
    Some(Splice {
        left,
        removed,
        inserted: Vec::new(),
        donor,
    })
}

pub(crate) struct Proof {
    pub(crate) seed_width: u32,
    pub(crate) live_width: u32,
    pub(crate) edits: Vec<ProvenEdit>,
}

pub(crate) struct ProvenEdit {
    pub(crate) from: u32,
    pub(crate) to: u32,
    pub(crate) live_from: u32,
    pub(crate) live_to: u32,
    pub(crate) donor: DonorV1,
}

impl Provenance {
    pub(crate) fn prove(
        &self,
        doc: &EditingDoc,
        txn: &impl ReadTxn,
        story_id: &str,
        key: &str,
    ) -> Option<Proof> {
        let records = self.entries.get(key)?.records()?;
        if self.overflowed || records.story != story_id || generation(txn) != self.generation {
            return None;
        }
        let story = crate::story_text(txn, story_id)?;
        let mut bounds = crate::op::para_bounds(&story, txn)
            .into_iter()
            .find(|bounds| bounds.para_id == key)?;
        bounds.start = crate::ops::past_leading_blocks(&story, txn, bounds.start);
        let seed = records.seed;
        let mark = unit_id(&story, txn, bounds.pilcrow)?;
        let seed_identity = doc.source_metadata()?.seed_identity(story_id)?;
        seeded_span(txn, &story, &bounds, mark, seed_identity, Some(seed))?;
        let mut ordered = records
            .splices
            .iter()
            .map(|splice| Some((splice.range(seed)?, splice)))
            .collect::<Option<Vec<_>>>()?;
        ordered.sort_unstable_by_key(|((from, _), _)| *from);
        let mut seed_at = 0;
        let mut live_at = bounds.start;
        let mut edits = Vec::new();
        for ((from, to), splice) in ordered {
            if from < seed_at
                || edits
                    .last()
                    .is_some_and(|edit: &ProvenEdit| edit.from == from)
                || !seed.slice(seed_at, from)?.matches(&story, txn, live_at)
            {
                return None;
            }
            live_at = live_at.checked_add(from.checked_sub(seed_at)?)?;
            let live_from = live_at.checked_sub(bounds.start)?;
            for &span in &splice.inserted {
                if span.len == 0 || seed.intersects(span) || !span.matches(&story, txn, live_at) {
                    return None;
                }
                live_at = live_at.checked_add(span.len)?;
            }
            edits.push(ProvenEdit {
                from,
                to,
                live_from,
                live_to: live_at.checked_sub(bounds.start)?,
                donor: splice.donor.clone(),
            });
            seed_at = to;
        }
        if live_at.checked_add(seed.len.checked_sub(seed_at)?)? != bounds.pilcrow
            || !seed.slice(seed_at, seed.len)?.matches(&story, txn, live_at)
        {
            return None;
        }
        Some(Proof {
            seed_width: seed.len,
            live_width: bounds.pilcrow.checked_sub(bounds.start)?,
            edits,
        })
    }
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum WatchKind {
    All,
    Recorded,
    History,
}

#[derive(Default)]
struct Touched {
    keys: BTreeSet<String>,
    history: BTreeSet<String>,
    recorded: BTreeSet<String>,
    recorded_only: bool,
    bytes: usize,
    unknown: bool,
    overflowed: bool,
}

impl Touched {
    fn add(&mut self, key: &str) {
        if self.overflowed
            || key.is_empty()
            || self.keys.contains(key)
            || (self.recorded_only && !self.recorded.contains(key))
        {
            return;
        }
        self.bytes = self.bytes.saturating_add(encoded_entry_size(key, &Entry::Unproven));
        if self.bytes > OPENING_BYTES {
            self.overflowed = true;
            self.keys.clear();
            self.history.clear();
        } else {
            self.keys.insert(key.to_owned());
        }
    }

    fn range(&mut self, bounds: &[ParaBounds], from: u32, to: u32) {
        for bounds in overlapping(bounds, from, to) {
            self.add(&bounds.para_id);
        }
    }

    fn recorded_range(&mut self, bounds: &[ParaBounds], from: u32, to: u32) {
        for bounds in overlapping(bounds, from, to) {
            self.add_recorded(&bounds.para_id);
        }
    }

    fn add_recorded(&mut self, key: &str) {
        if self.recorded.contains(key) {
            self.add(key);
        }
    }
}

fn overlapping(bounds: &[ParaBounds], from: u32, to: u32) -> impl Iterator<Item = &ParaBounds> {
    let first = bounds.partition_point(|bounds| bounds.pilcrow < from);
    bounds
        .iter()
        .skip(first)
        .take_while(move |bounds| bounds.start <= to)
}

fn lock_touched(touched: &Mutex<Touched>) -> MutexGuard<'_, Touched> {
    touched.lock().unwrap_or_else(|poisoned| {
        let mut touched = poisoned.into_inner();
        touched.unknown = true;
        touched
    })
}

#[derive(Default)]
struct StoryBounds {
    paragraphs: Vec<ParaBounds>,
    blocks: Vec<u32>,
}

impl StoryBounds {
    fn new(story: &TextRef, txn: &impl ReadTxn) -> Option<Self> {
        let mut bounds = Self::default();
        let (mut start, mut at) = (0u32, 0u32);
        for diff in story.diff(txn, YChange::identity) {
            let end = at.checked_add(crate::out_len(&diff.insert))?;
            if let Out::YMap(map) = diff.insert {
                if is_block_map(&map, txn) {
                    bounds.blocks.push(at);
                } else if crate::is_pilcrow(&map, txn) {
                    bounds.paragraphs.push(ParaBounds {
                        para_id: crate::map_string(&map, txn, PARA_ID).unwrap_or_default(),
                        start,
                        pilcrow: at,
                    });
                    start = end;
                }
            }
            at = end;
        }
        Some(bounds)
    }

    fn block_only(&self, from: u32, to: u32) -> bool {
        let start = self.blocks.partition_point(|at| *at < from);
        let end = self.blocks.partition_point(|at| *at < to);
        to.checked_sub(from)
            .is_some_and(|width| end - start == width as usize)
    }
}

pub(crate) struct Watch {
    provenance: Arc<Mutex<Provenance>>,
    touched: Arc<Mutex<Touched>>,
    subscriptions: Vec<Subscription>,
}

impl Watch {
    pub(crate) fn new(doc: &Doc, provenance: &Arc<Mutex<Provenance>>, kind: WatchKind) -> Self {
        let touched = Arc::new(Mutex::new(Touched::default()));
        let mut watch = Self {
            provenance: Arc::clone(provenance),
            touched,
            subscriptions: Vec::new(),
        };
        let recorded = {
            let proof = lock_provenance(provenance);
            if proof.overflowed {
                return watch;
            }
            proof
                .entries
                .iter()
                .filter_map(|(key, entry)| Some((key.clone(), entry.records()?.story.clone())))
                .collect::<BTreeMap<_, _>>()
        };
        if kind == WatchKind::Recorded && recorded.is_empty() {
            return watch;
        }
        {
            let mut touched = lock_touched(&watch.touched);
            touched.recorded = recorded.keys().cloned().collect();
            touched.recorded_only = kind == WatchKind::Recorded;
        }
        let Ok(txn) = doc.try_transact() else {
            lock_provenance(provenance).invalidate_records();
            return watch;
        };
        if generation(&txn).is_none() {
            return watch;
        }
        let Some(stories) = txn.get_map(STORIES) else {
            return watch;
        };
        let before = {
            let recorded_stories: BTreeSet<_> = recorded.values().map(String::as_str).collect();
            crate::live_stories(&stories, &txn)
                .filter(|(story, _)| {
                    kind != WatchKind::Recorded || recorded_stories.contains(story)
                })
                .filter_map(|(key, value)| match value {
                    Out::YText(text) => Some((key.to_owned(), StoryBounds::new(&text, &txn)?)),
                    _ => None,
                })
                .collect::<BTreeMap<_, _>>()
        };
        let removed = txn.get_map(crate::REMOVED_STORIES);
        let marks = txn.get_map(crate::RAW_INLINES);
        let comments = txn.get_map(crate::COMMENTS).filter(|_| !recorded.is_empty());
        let comment_anchors = comments
            .as_ref()
            .map(|_| crate::canonical::comment_anchors_by_story(&txn));
        drop(txn);
        let before = Arc::new(before);
        let touched = Arc::clone(&watch.touched);
        let prior = Arc::clone(&before);
        watch
            .subscriptions
            .push(stories.observe_deep(move |txn, events| {
                let mut touched = lock_touched(&touched);
                collect_touched(txn, events, &prior, kind, &mut touched);
            }));
        if let Some(removed) = removed {
            let touched = Arc::clone(&watch.touched);
            watch.subscriptions.push(removed.observe(move |txn, event| {
                let mut touched = lock_touched(&touched);
                for story in event.keys(txn).keys() {
                    if let Some(bounds) = before.get(story.as_ref()) {
                        touched.range(&bounds.paragraphs, 0, u32::MAX);
                    }
                }
            }));
        }
        if let Some(marks) = marks.filter(|_| !recorded.is_empty()) {
            let touched = Arc::clone(&watch.touched);
            watch.subscriptions.push(marks.observe(move |txn, event| {
                let mut touched = lock_touched(&touched);
                for change in event.keys(txn).values() {
                    match change {
                        EntryChange::Inserted(value) | EntryChange::Removed(value) => {
                            touch_raw_mark(txn, value, &mut touched);
                        }
                        EntryChange::Updated(before, after) => {
                            touch_raw_mark(txn, before, &mut touched);
                            touch_raw_mark(txn, after, &mut touched);
                        }
                    }
                }
            }));
        }
        if let Some((comments, before)) = comments.zip(comment_anchors) {
            let touched = Arc::clone(&watch.touched);
            watch
                .subscriptions
                .push(comments.observe_deep(move |txn, events| {
                    let keys = changed_comment_anchors(txn, events);
                    if keys.is_empty() {
                        return;
                    }
                    let after = crate::canonical::comment_anchors_by_story(txn);
                    let mut touched = lock_touched(&touched);
                    for (story, anchors) in before.iter().chain(&after) {
                        let Some(text) = crate::story_text(txn, story) else {
                            continue;
                        };
                        let bounds = crate::op::para_bounds(&text, txn);
                        for anchor in keys.iter().filter_map(|key| anchors.get(key)).flatten() {
                            if let (Some(start), Some(end)) =
                                (anchor.start.get_offset(txn), anchor.end.get_offset(txn))
                            {
                                touched.recorded_range(&bounds, start.index, end.index);
                            }
                        }
                    }
                }));
        }
        watch
    }

    fn disarm(&mut self) {
        self.subscriptions.clear();
    }
}

impl Drop for Watch {
    fn drop(&mut self) {
        if self.subscriptions.is_empty() {
            return;
        }
        self.disarm();
        let touched = std::mem::take(&mut *lock_touched(&self.touched));
        let mut proof = lock_provenance(&self.provenance);
        if touched.overflowed {
            proof.overflow();
            return;
        }
        if touched.unknown {
            proof.invalidate_records();
        }
        for key in touched.keys {
            if touched.history.contains(&key) {
                proof.set(key, Entry::HistoryReplacement);
            } else {
                proof.mark(key);
            }
        }
    }
}

fn changed_comment_anchors(txn: &TransactionMut<'_>, events: &Events) -> BTreeSet<String> {
    let mut keys = BTreeSet::new();
    for event in events.iter() {
        let Event::Map(event) = event else {
            continue;
        };
        let path = event.path();
        if path.is_empty() {
            keys.extend(event.keys(txn).keys().map(ToString::to_string));
        } else if path.len() == 1
            && event.keys(txn).contains_key("anchors")
            && let Some(PathSegment::Key(key)) = path.front()
        {
            keys.insert(key.to_string());
        }
    }
    keys
}

fn touch_raw_mark(txn: &TransactionMut<'_>, value: &Out, touched: &mut Touched) {
    let mark = (|| {
        let Out::Any(Any::Map(mark)) = value else {
            return None;
        };
        let (Some(Any::String(story)), Some(Any::Buffer(at))) = (mark.get("story"), mark.get("at"))
        else {
            return None;
        };
        let text = crate::story_text(txn, story)?;
        let offset = StickyIndex::decode_v1(at).ok()?.get_offset(txn)?;
        let bounds = crate::op::para_bounds(&text, txn);
        Some((bounds, offset.index))
    })();
    if let Some((bounds, at)) = mark {
        touched.recorded_range(&bounds, at, at);
    }
}

fn collect_touched(
    txn: &TransactionMut<'_>,
    events: &Events,
    before: &BTreeMap<String, StoryBounds>,
    kind: WatchKind,
    touched: &mut Touched,
) {
    let mut after = BTreeMap::new();
    for event in events.iter() {
        let path = event.path();
        let Some(PathSegment::Key(story)) = path.front() else {
            if let Event::Map(event) = event {
                for story in event.keys(txn).keys() {
                    if let Some(bounds) = before.get(story.as_ref()) {
                        touched.range(&bounds.paragraphs, 0, u32::MAX);
                    }
                }
            }
            continue;
        };
        let old_story = before.get(story.as_ref());
        let old = old_story.map_or(&[][..], |bounds| bounds.paragraphs.as_slice());
        let now = after.entry(story.to_string()).or_insert_with(|| {
            crate::story_text(txn, story.as_ref())
                .map_or_else(Vec::new, |text| crate::op::para_bounds(&text, txn))
        });
        match event {
            Event::Text(event) if path.len() == 1 => {
                if let Some(old_story) = old_story
                    && collect_text_touches(txn, event.delta(txn), old_story, now, kind, touched)
                        .is_none()
                {
                    touched.recorded_range(old, 0, u32::MAX);
                }
            }
            Event::Map(event) => {
                let keys = event.keys(txn);
                let pilcrow = crate::is_pilcrow(event.target(), txn);
                if is_block_map(event.target(), txn) {
                    continue;
                }
                if pilcrow && !keys.contains_key(PARA_ID) && !keys.contains_key(crate::KIND_KEY) {
                    if kind == WatchKind::History
                        && let Some(key) = crate::map_string(event.target(), txn, PARA_ID)
                    {
                        touched.add_recorded(&key);
                    }
                    continue;
                }
                if let Some(
                    EntryChange::Updated(Out::Any(Any::String(key)), _)
                    | EntryChange::Removed(Out::Any(Any::String(key))),
                ) = keys.get(PARA_ID)
                {
                    touched.add(key);
                }
                if pilcrow && let Some(key) = crate::map_string(event.target(), txn, PARA_ID)
                {
                    if old.iter().any(|bounds| bounds.para_id == key) {
                        touched.add(&key);
                    }
                } else if let Some(PathSegment::Index(at)) = path.get(1) {
                    touched.range(now, *at, *at);
                } else {
                    touched.recorded_range(old, 0, u32::MAX);
                }
            }
            _ => {
                if let Some(PathSegment::Index(at)) = path.get(1) {
                    touched.range(now, *at, *at);
                } else {
                    touched.recorded_range(old, 0, u32::MAX);
                }
            }
        }
    }
}

fn collect_text_touches(
    txn: &TransactionMut<'_>,
    deltas: &[Delta<Out>],
    before: &StoryBounds,
    now: &[ParaBounds],
    kind: WatchKind,
    touched: &mut Touched,
) -> Option<()> {
    let old = &before.paragraphs;
    let old_keys: BTreeSet<_> = old.iter().map(|bounds| bounds.para_id.as_str()).collect();
    let (mut was, mut is) = (0u32, 0u32);
    let mut deleted = BTreeSet::new();
    let mut inserted = BTreeSet::new();
    for delta in deltas {
        match delta {
            Delta::Retain(len, attrs) => {
                let old_end = was.checked_add(*len)?;
                let new_end = is.checked_add(*len)?;
                if let Some(attrs) = attrs.as_ref().filter(|_| *len != 0) {
                    if attrs.contains_key(crate::INS) || attrs.contains_key(crate::DEL) {
                        touched.range(old, was, old_end - 1);
                        touched.range(now, is, new_end - 1);
                    } else if kind == WatchKind::History {
                        touched.recorded_range(old, was, old_end - 1);
                        touched.recorded_range(now, is, new_end - 1);
                    }
                }
                was = old_end;
                is = new_end;
            }
            Delta::Deleted(len) if *len != 0 => {
                let end = was.checked_add(*len)?;
                for bounds in overlapping(old, was, end - 1) {
                    let from = was.max(bounds.start);
                    let to = end.min(bounds.pilcrow.checked_add(1)?);
                    if !before.block_only(from, to) {
                        touched.add(&bounds.para_id);
                    }
                    if kind == WatchKind::History
                        && from < end.min(bounds.pilcrow)
                        && !before.block_only(from, end.min(bounds.pilcrow))
                    {
                        deleted.insert(bounds.para_id.as_str());
                    }
                }
                if overlapping(old, was, was).any(|bounds| {
                    bounds.pilcrow < end && !before.block_only(bounds.start, was)
                }) {
                    touched.range(now, is, is);
                }
                was = end;
            }
            Delta::Deleted(_) => {}
            Delta::Inserted(value, _) => {
                let len = crate::out_len(value);
                let end = is.checked_add(len)?;
                if let Out::YMap(map) = value
                    && crate::is_pilcrow(map, txn)
                {
                    for bounds in overlapping(old, was, was) {
                        if !before.block_only(bounds.start, was) {
                            touched.add(&bounds.para_id);
                        }
                    }
                } else if len != 0 && !is_block(value, txn) {
                    for bounds in overlapping(now, is, end - 1) {
                        if old_keys.contains(bounds.para_id.as_str()) {
                            touched.add(&bounds.para_id);
                            if kind == WatchKind::History
                                && matches!(value, Out::Any(Any::String(_)))
                            {
                                inserted.insert(bounds.para_id.as_str());
                            }
                        }
                    }
                }
                is = end;
            }
        }
    }
    for key in deleted.intersection(&inserted) {
        if touched.keys.contains(*key) {
            touched.history.insert((*key).to_owned());
        }
    }
    Some(())
}

fn is_block(value: &Out, txn: &impl ReadTxn) -> bool {
    matches!(value, Out::YMap(map) if is_block_map(map, txn))
}

fn is_block_map(map: &yrs::MapRef, txn: &impl ReadTxn) -> bool {
    crate::map_string(map, txn, crate::KIND_KEY)
        .is_some_and(|kind| crate::segments::is_block_embed(&kind))
}

impl EditingDoc {
    pub(crate) fn provenance_snapshot(&self) -> Provenance {
        lock_provenance(&self.provenance).clone()
    }

    pub(crate) fn clear_edit_provenance(&self) {
        *lock_provenance(&self.provenance) = Provenance::default();
    }

    pub(crate) fn watch_provenance(&self, kind: WatchKind) -> Option<Watch> {
        self.source_metadata()?;
        Some(Watch::new(&self.doc, &self.provenance, kind))
    }

    pub(crate) fn mark_provenance_range(
        &self,
        txn: &impl ReadTxn,
        story: &TextRef,
        from: u32,
        to: u32,
    ) {
        if self.source_metadata().is_none() || generation(txn).is_none() {
            return;
        }
        let to = from.max(to);
        let ranges = crate::op::para_bounds(story, txn);
        let mut proof = lock_provenance(&self.provenance);
        for bounds in overlapping(&ranges, from, to) {
            proof.mark(bounds.para_id.clone());
        }
    }

    pub(crate) fn mark_provenance_key(&self, txn: &impl ReadTxn, key: &str) {
        if self.source_metadata().is_some() && generation(txn).is_some() {
            lock_provenance(&self.provenance).mark(key.to_owned());
        }
    }

    pub(crate) fn invalidate_fork_provenance(&self, fork: &Self) {
        let original = self.provenance_snapshot();
        let mut proof = lock_provenance(&fork.provenance);
        if original.overflowed {
            proof.overflow();
        } else {
            for (key, entry) in original.entries {
                if matches!(entry, Entry::HistoryReplacement) {
                    proof.set(key, entry);
                } else {
                    proof.mark(key);
                }
            }
        }
    }

    pub(crate) fn bind_edit_provenance(&self, base: &StateVector, update: &[u8]) {
        let mut proof = lock_provenance(&self.provenance);
        if proof.entries.is_empty() && !proof.overflowed {
            return;
        }
        let base = base.encode_v1();
        if base
            .len()
            .saturating_add(proof.bytes)
            .saturating_add(proof.generation.as_ref().map_or(0, String::len))
            > OPENING_BYTES
        {
            proof.outgoing = None;
            return;
        }
        proof.outgoing = Some(Outgoing {
            epoch: self.committed_epoch(),
            base,
            digest: Sha256::digest(update).into(),
        });
    }

    /// Encodes provenance for the latest diff of the current state.
    #[doc(hidden)]
    pub fn encode_edit_provenance(&self) -> Vec<u8> {
        let Ok(txn) = self.doc.try_transact() else {
            return Vec::new();
        };
        let proof = self.provenance_snapshot();
        let Some(outgoing) = proof
            .outgoing
            .as_ref()
            .filter(|out| out.epoch == self.committed_epoch())
        else {
            return Vec::new();
        };
        let Ok(base) = StateVector::decode_v1(&outgoing.base) else {
            lock_provenance(&self.provenance).invalidate_records();
            return Vec::new();
        };
        encode_sidecar(&txn, &proof, Some((&outgoing.base, &base)), outgoing.digest)
    }

    /// Encodes provenance bound to a diff or full bootstrap state.
    #[doc(hidden)]
    pub fn encode_edit_provenance_for_transfer(&self, bytes: &[u8], base: Option<&[u8]>) -> Vec<u8> {
        let Ok(txn) = self.doc.try_transact() else {
            return Vec::new();
        };
        if txn.has_missing_updates() {
            return Vec::new();
        }
        let digest = Sha256::digest(bytes).into();
        let base = match base {
            Some(base) if base.len() <= OPENING_BYTES => {
                let Ok(vector) = StateVector::decode_v1(base) else {
                    return Vec::new();
                };
                if crate::deterministic::encode_diff_v1(&txn, &vector).as_slice() != bytes {
                    return Vec::new();
                }
                Some((base, vector))
            }
            Some(_) => return Vec::new(),
            None if state_digest(&txn) == digest => None,
            None => return Vec::new(),
        };
        let proof = self.provenance_snapshot();
        encode_sidecar(
            &txn,
            &proof,
            base.as_ref().map(|(bytes, vector)| (*bytes, vector)),
            digest,
        )
    }

    pub(crate) fn install_bootstrap_provenance(&self, state: &[u8], sidecar: &[u8]) {
        let txn = self.doc.transact();
        let incoming = decode_sidecar(&txn, state, sidecar)
            .filter(|incoming| incoming.base.is_none() && incoming.state_digest.is_none());
        let mut proof = lock_provenance(&self.provenance);
        match incoming {
            Some(incoming) => proof.merge(incoming),
            None => proof.overflow(),
        }
    }

    /// Applies a local update and its bound provenance.
    #[doc(hidden)]
    pub fn apply_local_update_with_provenance(
        &self,
        update: &[u8],
        sidecar: &[u8],
    ) -> EditResult<()> {
        let decoded = yrs::Update::decode_v1(update)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        let incoming = {
            let txn = self.doc.transact();
            decode_update_sidecar(&txn, update, sidecar)
        };
        let mut watch = self.watch_provenance(WatchKind::All);
        let result = self.integrate_updates_inner(vec![decoded], UpdateOrigin::Local);
        if result.is_ok()
            && !self.holds_pending_update()
            && let Some(incoming) = incoming
        {
            let txn = self.doc.transact();
            if generation(&txn).as_deref() == Some(incoming.generation.as_str())
                && incoming
                    .state_digest
                    .is_none_or(|digest| digest == state_digest(&txn))
            {
                if let Some(watch) = watch.as_mut() {
                    watch.disarm();
                }
                lock_provenance(&self.provenance).merge(incoming);
            }
        }
        result
    }
}

fn encode_sidecar(
    txn: &impl ReadTxn,
    proof: &Provenance,
    base: Option<(&[u8], &StateVector)>,
    digest: [u8; CHECKSUM_BYTES],
) -> Vec<u8> {
    let Some(generation) = generation(txn).filter(|value| value.len() <= OPENING_BYTES) else {
        return Vec::new();
    };
    if proof.generation.as_ref().is_some_and(|recorded| recorded != &generation) {
        return Vec::new();
    }
    let encode = |entries: &BTreeMap<String, Entry>, overflowed| {
        let mut bytes = vec![VERSION];
        let sidecar = Sidecar {
            generation: &generation,
            base: base.map(|(bytes, _)| bytes),
            digest,
            state_digest: base
                .is_some_and(|(_, vector)| references_older_insertions(entries, vector))
                .then(|| state_digest(txn)),
            overflowed,
            entries,
        };
        serde_json::to_writer(&mut bytes, &sidecar).ok()?;
        let checksum = Sha256::digest(&bytes);
        bytes.extend_from_slice(&checksum);
        Some(bytes)
    };
    encode(&proof.entries, proof.overflowed)
        .filter(|bytes| bytes.len() <= OPENING_BYTES)
        .or_else(|| encode(&BTreeMap::new(), true).filter(|bytes| bytes.len() <= OPENING_BYTES))
        .unwrap_or_default()
}

fn decode_update_sidecar(txn: &impl ReadTxn, update: &[u8], sidecar: &[u8]) -> Option<Incoming> {
    let incoming = decode_sidecar(txn, update, sidecar)?;
    let base = StateVector::decode_v1(incoming.base.as_deref()?).ok()?;
    if base != txn.state_vector()
        || (incoming.state_digest.is_none()
            && references_older_insertions(&incoming.entries, &base))
    {
        return None;
    }
    Some(incoming)
}

fn decode_sidecar(txn: &impl ReadTxn, bytes: &[u8], sidecar: &[u8]) -> Option<Incoming> {
    if sidecar.first() != Some(&VERSION) || sidecar.len() > OPENING_BYTES {
        return None;
    }
    let (data, checksum) = sidecar.split_at(sidecar.len().checked_sub(CHECKSUM_BYTES)?);
    if data.is_empty() || &Sha256::digest(data)[..] != checksum {
        return None;
    }
    let incoming: Incoming = serde_json::from_slice(&data[1..]).ok()?;
    if generation(txn)?.as_str() != incoming.generation
        || incoming.digest != <[u8; CHECKSUM_BYTES]>::from(Sha256::digest(bytes))
        || txn.store().pending_update().is_some()
        || txn.store().pending_ds().is_some()
        || incoming.entries.iter().any(|(key, entry)| {
            key.is_empty()
                || encoded_entry_size(key, entry) > OCCURRENCE_BYTES
                || entry.records().is_some_and(|records| {
                    records.seed.end().is_none()
                        || records.splices.is_empty()
                        || records.splices.iter().any(|splice| {
                            splice.range(records.seed).is_none()
                                || splice.inserted.iter().any(|span| {
                                    span.first.0 >= (1 << 53)
                                        || span.len == 0
                                        || span.end().is_none()
                                })
                        })
                })
        })
    {
        return None;
    }
    Some(incoming)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_provenance_watch_inside_a_write_transaction_fails_closed() {
        let mut proof = Provenance::default();
        proof.record(
            "generation".to_owned(),
            "paragraph".to_owned(),
            "body".to_owned(),
            Span { first: UnitId(1, 0), len: 1 },
            Splice {
                left: None,
                removed: Some(Span { first: UnitId(1, 0), len: 1 }),
                inserted: Vec::new(),
                donor: DonorV1::At { at: 0 },
            },
        );
        let proof = Arc::new(Mutex::new(proof));
        let doc = Doc::new();
        let _txn = doc.transact_mut();

        let watch = Watch::new(&doc, &proof, WatchKind::Recorded);
        assert!(watch.subscriptions.is_empty());
        let proof = lock_provenance(&proof);
        assert!(proof.unproven("paragraph"));
        assert!(!proof.unproven("neighbor"));
    }
}
