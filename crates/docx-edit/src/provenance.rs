use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex};

use docx_parse::paragraph_patch::DonorV1;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use yrs::branch::{Branch, BranchID};
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
const VERSION: u8 = 1;
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

    fn slice(self, from: u32, to: u32) -> Self {
        Self {
            first: UnitId(self.first.0, self.first.1 + from),
            len: to - from,
        }
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

#[derive(Default)]
pub(crate) struct Provenance {
    generation: Option<String>,
    entries: BTreeMap<String, Option<Records>>,
    bytes: usize,
    overflowed: bool,
    outgoing: Option<Outgoing>,
}

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
    entries: &'a BTreeMap<String, Option<Records>>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct Incoming {
    generation: String,
    base: Option<Vec<u8>>,
    digest: [u8; CHECKSUM_BYTES],
    state_digest: Option<[u8; CHECKSUM_BYTES]>,
    overflowed: bool,
    entries: BTreeMap<String, Option<Records>>,
}

fn encoded_entry_size(key: &str, records: &Option<Records>) -> usize {
    serde_json::to_vec(key)
        .expect("encode provenance key")
        .len()
        + serde_json::to_vec(records)
            .expect("encode provenance entry")
            .len()
        + 2
}

impl Provenance {
    pub(crate) fn overflow(&mut self) {
        self.entries.clear();
        self.generation = None;
        self.bytes = 0;
        self.overflowed = true;
        self.outgoing = None;
    }

    fn set(&mut self, key: String, mut records: Option<Records>) {
        if key.is_empty() || self.overflowed || self.entries.get(&key).is_some_and(Option::is_none)
        {
            return;
        }
        let mut size = encoded_entry_size(&key, &records);
        if size > OCCURRENCE_BYTES {
            records = None;
            size = encoded_entry_size(&key, &records);
        }
        if let Some(old) = self.entries.get(&key) {
            self.bytes -= encoded_entry_size(&key, old);
        }
        self.bytes += size;
        if size > OCCURRENCE_BYTES
            || self.bytes + self.generation.as_ref().map_or(0, String::len) > OPENING_BYTES
        {
            self.overflow();
        } else {
            self.entries.insert(key, records);
            self.outgoing = None;
        }
    }

    pub(crate) fn mark(&mut self, key: String) {
        self.set(key, None);
    }

    pub(crate) fn has_entry(&self, key: &str) -> bool {
        self.entries.contains_key(key)
    }

    pub(crate) fn unproven(&self, key: &str) -> bool {
        self.overflowed || self.entries.get(key).is_some_and(Option::is_none)
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
            .cloned()
            .flatten()
            .unwrap_or_else(|| Records {
                story: story.clone(),
                seed,
                splices: Vec::new(),
            });
        if records.story != story
            || records.seed != seed
            || records
                .splices
                .iter()
                .any(|other| splice.conflicts(other, seed))
        {
            self.mark(key);
        } else {
            records.splices.push(splice);
            self.set(key, Some(records));
        }
    }

    fn merge(&mut self, incoming: Incoming) {
        if incoming.overflowed {
            self.overflow();
            return;
        }
        for (key, records) in incoming.entries {
            if let Some(records) = records {
                for splice in records.splices {
                    if self
                        .entries
                        .get(&key)
                        .and_then(Option::as_ref)
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
                self.mark(key);
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
    entries: &BTreeMap<String, Option<Records>>,
    base: &StateVector,
) -> bool {
    entries.values().flatten().any(|records| {
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
    doc: &EditingDoc,
    txn: &T,
    story_id: &str,
    story: &TextRef,
    bounds: &ParaBounds,
    mark: UnitId,
) -> Option<Span> {
    let (clock, len) = doc.source_metadata()?.seed_identity(story_id)?;
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
    let start = if bounds.start == 0 {
        clock
    } else {
        let left = unit_id(story, txn, bounds.start - 1)?;
        seeded.offset(left)?;
        left.1.checked_add(1)?
    };
    Some(Span {
        first: UnitId(client, start),
        len: mark.1.checked_sub(start)?,
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
        doc.source_metadata()?;
        let generation = generation(txn)?;
        if doc.provenance.lock().unwrap().overflowed {
            return None;
        }
        let bounds = indexed
            .and_then(|index| index.bounds_at(txn, start))
            .or_else(|| {
                crate::op::para_bounds(story, txn)
                    .into_iter()
                    .find(|bounds| bounds.start <= start && start <= bounds.pilcrow)
            })?;
        if !inheriting || end > bounds.pilcrow || start < bounds.start {
            doc.mark_provenance_range(txn, story, start, end);
            return None;
        }
        if doc.provenance.lock().unwrap().unproven(&bounds.para_id) {
            return None;
        }
        if let Some((seed, splice)) =
            capture_splice(doc, txn, story_id, story, &bounds, start, end, chunks)
        {
            Some(Self {
                key: bounds.para_id,
                story: story_id.to_owned(),
                generation,
                seed,
                splice,
            })
        } else {
            doc.provenance.lock().unwrap().mark(bounds.para_id);
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
                doc.provenance.lock().unwrap().mark(self.key);
                return;
            };
            if let Some(last) = self.splice.inserted.last_mut()
                && last.first.0 == id.0
                && last.end() == Some(id.1)
            {
                last.len += 1;
            } else {
                self.splice.inserted.push(Span { first: id, len: 1 });
                if self.splice.inserted.len() * 16 > OCCURRENCE_BYTES {
                    doc.provenance.lock().unwrap().mark(self.key);
                    return;
                }
            }
        }
        doc.provenance.lock().unwrap().record(
            self.generation,
            self.key,
            self.story,
            self.seed,
            self.splice,
        );
    }
}

fn capture_splice(
    doc: &EditingDoc,
    txn: &TransactionMut<'_>,
    story_id: &str,
    story: &TextRef,
    bounds: &ParaBounds,
    start: u32,
    end: u32,
    chunks: &[Chunk],
) -> Option<(Span, Splice)> {
    let mark = unit_id(story, txn, bounds.pilcrow)?;
    let seed = seeded_span(doc, txn, story_id, story, bounds, mark)?;
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
        let at = seed.offset(unit_id(story, txn, start - 1)?)?;
        let (_, width) = crate::ops::code_point_boundary(chunks, start, false);
        let donor = if width == 0 {
            DonorV1::Pilcrow
        } else {
            DonorV1::At {
                at: (at + 1).checked_sub(width)?,
            }
        };
        (at + 1, None, donor)
    };
    let left = (from > 0).then(|| UnitId(seed.first.0, seed.first.1 + from - 1));
    Some((
        seed,
        Splice {
            left,
            removed,
            inserted: Vec::new(),
            donor,
        },
    ))
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
        let records = self.entries.get(key)?.as_ref()?;
        if self.overflowed || records.story != story_id || generation(txn) != self.generation {
            return None;
        }
        let story = crate::story_text(txn, story_id)?;
        let bounds = crate::op::para_bounds(&story, txn)
            .into_iter()
            .find(|bounds| bounds.para_id == key)?;
        let seed = records.seed;
        let mark = unit_id(&story, txn, bounds.pilcrow)?;
        if seeded_span(doc, txn, story_id, &story, &bounds, mark)? != seed {
            return None;
        }
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
                || !seed.slice(seed_at, from).matches(&story, txn, live_at)
            {
                return None;
            }
            live_at += from - seed_at;
            let live_from = live_at - bounds.start;
            for &span in &splice.inserted {
                if span.len == 0 || seed.intersects(span) || !span.matches(&story, txn, live_at) {
                    return None;
                }
                live_at += span.len;
            }
            edits.push(ProvenEdit {
                from,
                to,
                live_from,
                live_to: live_at - bounds.start,
                donor: splice.donor.clone(),
            });
            seed_at = to;
        }
        if live_at.checked_add(seed.len - seed_at)? != bounds.pilcrow
            || !seed.slice(seed_at, seed.len).matches(&story, txn, live_at)
        {
            return None;
        }
        Some(Proof {
            seed_width: seed.len,
            live_width: bounds.len(),
            edits,
        })
    }
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum WatchKind {
    All,
    Recorded,
    Batch,
}

#[derive(Default)]
struct Touched {
    keys: BTreeSet<String>,
    recorded_only: Option<BTreeSet<String>>,
    bytes: usize,
    unknown: bool,
}

impl Touched {
    fn add(&mut self, key: &str) {
        if self.unknown
            || key.is_empty()
            || self.keys.contains(key)
            || self
                .recorded_only
                .as_ref()
                .is_some_and(|recorded| !recorded.contains(key))
        {
            return;
        }
        self.bytes += encoded_entry_size(key, &None);
        if self.bytes > OPENING_BYTES {
            self.unknown = true;
            self.keys.clear();
        } else {
            self.keys.insert(key.to_owned());
        }
    }

    fn range(&mut self, bounds: &[ParaBounds], from: u32, to: u32) {
        for bounds in bounds {
            if bounds.start <= to && from <= bounds.pilcrow {
                self.add(&bounds.para_id);
            }
        }
    }
}

type StoryBounds = BTreeMap<String, Vec<ParaBounds>>;

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
        {
            let proof = provenance.lock().unwrap();
            if proof.overflowed {
                return watch;
            }
            if kind == WatchKind::Recorded {
                let keys: BTreeSet<_> = proof
                    .entries
                    .iter()
                    .filter(|(_, records)| records.is_some())
                    .map(|(key, _)| key.clone())
                    .collect();
                if keys.is_empty() {
                    return watch;
                }
                watch.touched.lock().unwrap().recorded_only = Some(keys);
            }
        }
        let txn = doc.transact();
        if generation(&txn).is_none() {
            return watch;
        }
        let Some(stories) = txn.get_map(STORIES) else {
            return watch;
        };
        let before = {
            let proof = provenance.lock().unwrap();
            let recorded_stories: BTreeSet<_> = proof
                .entries
                .values()
                .filter_map(|records| Some(records.as_ref()?.story.as_str()))
                .collect();
            crate::live_stories(&stories, &txn)
                .filter(|(story, _)| {
                    kind != WatchKind::Recorded || recorded_stories.contains(story)
                })
                .filter_map(|(key, value)| match value {
                    Out::YText(text) => Some((key.to_owned(), crate::op::para_bounds(&text, &txn))),
                    _ => None,
                })
                .collect::<StoryBounds>()
        };
        let before = Arc::new(before);
        let touched = Arc::clone(&watch.touched);
        let prior = Arc::clone(&before);
        watch
            .subscriptions
            .push(stories.observe_deep(move |txn, events| {
                let mut touched = touched.lock().unwrap();
                collect_touched(txn, events, &prior, kind, &mut touched);
            }));
        if let Some(removed) = txn.get_map(crate::REMOVED_STORIES) {
            let touched = Arc::clone(&watch.touched);
            watch.subscriptions.push(removed.observe(move |txn, event| {
                let mut touched = touched.lock().unwrap();
                for story in event.keys(txn).keys() {
                    if let Some(bounds) = before.get(story.as_ref()) {
                        touched.range(bounds, 0, u32::MAX);
                    }
                    if let Some(text) = crate::story_text(txn, story.as_ref()) {
                        touched.range(&crate::op::para_bounds(&text, txn), 0, u32::MAX);
                    }
                }
            }));
        }
        if let Some(marks) = txn.get_map(crate::RAW_INLINES) {
            let touched = Arc::clone(&watch.touched);
            watch.subscriptions.push(marks.observe(move |txn, event| {
                let mut touched = touched.lock().unwrap();
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
        if let Some(comments) = txn.get_map(crate::COMMENTS) {
            let touched = Arc::clone(&watch.touched);
            let before = crate::canonical::comment_anchors_by_story(&txn);
            watch
                .subscriptions
                .push(comments.observe_deep(move |txn, events| {
                    let keys = changed_comment_anchors(txn, events);
                    if keys.is_empty() {
                        return;
                    }
                    let after = crate::canonical::comment_anchors_by_story(txn);
                    let mut touched = touched.lock().unwrap();
                    for (story, anchors) in before.iter().chain(&after) {
                        let Some(text) = crate::story_text(txn, story) else {
                            continue;
                        };
                        let bounds = crate::op::para_bounds(&text, txn);
                        for anchor in keys.iter().filter_map(|key| anchors.get(key)).flatten() {
                            if let (Some(start), Some(end)) =
                                (anchor.start.get_offset(txn), anchor.end.get_offset(txn))
                            {
                                touched.range(&bounds, start.index, end.index);
                            }
                        }
                    }
                }));
        }
        watch
    }

    pub(crate) fn is_active(&self) -> bool {
        !self.subscriptions.is_empty()
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
        let touched = std::mem::take(&mut *self.touched.lock().unwrap());
        let mut proof = self.provenance.lock().unwrap();
        if touched.unknown {
            proof.overflow();
            return;
        }
        for key in touched.keys {
            proof.mark(key);
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
        touched.range(&bounds, at, at);
    } else {
        touched.unknown = true;
    }
}

fn collect_touched(
    txn: &TransactionMut<'_>,
    events: &Events,
    before: &StoryBounds,
    kind: WatchKind,
    touched: &mut Touched,
) {
    let mut after = StoryBounds::new();
    for event in events.iter() {
        let path = event.path();
        let Some(PathSegment::Key(story)) = path.front() else {
            if let Event::Map(event) = event {
                for story in event.keys(txn).keys() {
                    if let Some(bounds) = before.get(story.as_ref()) {
                        touched.range(bounds, 0, u32::MAX);
                    }
                    if let Some(text) = crate::story_text(txn, story.as_ref()) {
                        touched.range(&crate::op::para_bounds(&text, txn), 0, u32::MAX);
                    }
                }
            }
            continue;
        };
        let old = before.get(story.as_ref()).map_or(&[][..], Vec::as_slice);
        let now = after.entry(story.to_string()).or_insert_with(|| {
            crate::story_text(txn, story.as_ref())
                .map_or_else(Vec::new, |text| crate::op::para_bounds(&text, txn))
        });
        match event {
            Event::Text(event) if path.len() == 1 => {
                let (mut was, mut is) = (0u32, 0u32);
                for delta in event.delta(txn) {
                    match delta {
                        Delta::Retain(len, attrs) => {
                            if *len != 0
                                && attrs.as_ref().is_some_and(|attrs| {
                                    kind != WatchKind::Batch
                                        || attrs.contains_key(crate::INS)
                                        || attrs.contains_key(crate::DEL)
                                })
                            {
                                touched.range(old, was, was + len - 1);
                                touched.range(now, is, is + len - 1);
                            }
                            was += len;
                            is += len;
                        }
                        Delta::Deleted(len) if *len != 0 => {
                            touched.range(old, was, was + len - 1);
                            let closes_paragraph = old
                                .iter()
                                .any(|bounds| was <= bounds.pilcrow && bounds.pilcrow < was + len);
                            if closes_paragraph {
                                touched.range(now, is, is);
                            }
                            was += len;
                        }
                        Delta::Deleted(_) => {}
                        Delta::Inserted(value, _) => {
                            let len = crate::out_len(value);
                            touched.range(old, was, was);
                            if len != 0 {
                                touched.range(now, is, is + len - 1);
                            }
                            is += len;
                        }
                    }
                }
            }
            Event::Map(event) => {
                let keys = event.keys(txn);
                if kind == WatchKind::Batch
                    && crate::is_pilcrow(event.target(), txn)
                    && !keys.contains_key(PARA_ID)
                    && !keys.contains_key(crate::KIND_KEY)
                {
                    continue;
                }
                if let Some(
                    EntryChange::Updated(Out::Any(Any::String(key)), _)
                    | EntryChange::Removed(Out::Any(Any::String(key))),
                ) = keys.get(PARA_ID)
                {
                    touched.add(key);
                }
                if crate::is_pilcrow(event.target(), txn)
                    && let Some(key) = crate::map_string(event.target(), txn, PARA_ID)
                {
                    touched.add(&key);
                } else if let Some(PathSegment::Index(at)) = path.get(1) {
                    touched.range(now, *at, *at);
                } else {
                    touched.unknown = true;
                }
            }
            _ => {
                if let Some(PathSegment::Index(at)) = path.get(1) {
                    touched.range(now, *at, *at);
                } else {
                    touched.unknown = true;
                }
            }
        }
    }
}

impl EditingDoc {
    pub(crate) fn clear_edit_provenance(&self) {
        *self.provenance.lock().unwrap() = Provenance::default();
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
        let mut proof = self.provenance.lock().unwrap();
        for bounds in ranges {
            if bounds.start <= to && from <= bounds.pilcrow {
                proof.mark(bounds.para_id);
            }
        }
    }

    pub(crate) fn mark_provenance_key(&self, txn: &impl ReadTxn, key: &str) {
        if self.source_metadata().is_some() && generation(txn).is_some() {
            self.provenance.lock().unwrap().mark(key.to_owned());
        }
    }

    pub(crate) fn invalidate_fork_provenance(&self, fork: &Self) {
        let original = self.provenance.lock().unwrap();
        let mut proof = fork.provenance.lock().unwrap();
        if original.overflowed {
            proof.overflow();
        } else {
            for key in original.entries.keys() {
                proof.mark(key.clone());
            }
        }
    }

    pub(crate) fn bind_edit_provenance(&self, base: &StateVector, update: &[u8]) {
        let mut proof = self.provenance.lock().unwrap();
        if proof.entries.is_empty() && !proof.overflowed {
            return;
        }
        let base = base.encode_v1();
        if base.len() + proof.bytes + proof.generation.as_ref().map_or(0, String::len)
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
        let txn = self.doc.transact();
        let proof = self.provenance.lock().unwrap();
        let Some(outgoing) = proof
            .outgoing
            .as_ref()
            .filter(|out| out.epoch == self.committed_epoch())
        else {
            return Vec::new();
        };
        let base = StateVector::decode_v1(&outgoing.base).expect("encoded provenance base");
        encode_sidecar(&txn, &proof, Some((&outgoing.base, &base)), outgoing.digest)
    }

    /// Encodes provenance bound to a diff or full bootstrap state.
    #[doc(hidden)]
    pub fn encode_edit_provenance_for_transfer(&self, bytes: &[u8], base: Option<&[u8]>) -> Vec<u8> {
        let txn = self.doc.transact();
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
        let proof = self.provenance.lock().unwrap();
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
        let mut proof = self.provenance.lock().unwrap();
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
        if self.holds_pending_update() || result.is_err() {
            if self.source_metadata().is_some() {
                self.provenance.lock().unwrap().overflow();
            }
        } else if let Some(incoming) = incoming {
            let txn = self.doc.transact();
            if generation(&txn).as_deref() == Some(incoming.generation.as_str())
                && incoming
                    .state_digest
                    .is_none_or(|digest| digest == state_digest(&txn))
            {
                if let Some(watch) = watch.as_mut() {
                    watch.disarm();
                }
                self.provenance.lock().unwrap().merge(incoming);
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
    let encode = |entries: &BTreeMap<String, Option<Records>>, overflowed| {
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
        serde_json::to_writer(&mut bytes, &sidecar).expect("encode edit provenance");
        let checksum = Sha256::digest(&bytes);
        bytes.extend_from_slice(&checksum);
        bytes
    };
    let bytes = encode(&proof.entries, proof.overflowed);
    if bytes.len() <= OPENING_BYTES {
        return bytes;
    }
    let bytes = encode(&BTreeMap::new(), true);
    if bytes.len() <= OPENING_BYTES {
        bytes
    } else {
        Vec::new()
    }
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
        || incoming.entries.iter().any(|(key, records)| {
            key.is_empty()
                || encoded_entry_size(key, records) > OCCURRENCE_BYTES
                || records.as_ref().is_some_and(|records| {
                    records.splices.is_empty()
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
