//! Splits a yrs v1 diff into ordered updates for an initial collaboration sync. Applied in order,
//! every update integrates completely (none waits on a later one), and together they rebuild
//! exactly the state the whole diff carries. Item structs travel first, causally ordered, then
//! the delete set travels alone, so the last update completes the sync.

use std::collections::{BTreeMap, HashMap, VecDeque};

use yrs::block::{
    BLOCK_GC_REF_NUMBER, BLOCK_ITEM_ANY_REF_NUMBER, BLOCK_ITEM_BINARY_REF_NUMBER,
    BLOCK_ITEM_DELETED_REF_NUMBER, BLOCK_ITEM_DOC_REF_NUMBER, BLOCK_ITEM_EMBED_REF_NUMBER,
    BLOCK_ITEM_FORMAT_REF_NUMBER, BLOCK_ITEM_JSON_REF_NUMBER, BLOCK_ITEM_STRING_REF_NUMBER,
    BLOCK_ITEM_TYPE_REF_NUMBER, BLOCK_SKIP_REF_NUMBER, HAS_ORIGIN, HAS_PARENT_SUB,
    HAS_RIGHT_ORIGIN,
};
use yrs::types::{
    TYPE_REFS_ARRAY, TYPE_REFS_DOC, TYPE_REFS_MAP, TYPE_REFS_TEXT, TYPE_REFS_UNDEFINED,
    TYPE_REFS_XML_ELEMENT, TYPE_REFS_XML_FRAGMENT, TYPE_REFS_XML_HOOK, TYPE_REFS_XML_TEXT,
};
use yrs::updates::decoder::Decode;
use yrs::{ClientID, ID, IdSet, ReadTxn, StateVector, Update};

const MAX_NESTING: u8 = 64;

/// Why a diff was not split; the caller then sends it whole.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SplitError {
    Malformed,
    Unsupported,
    Unordered,
}

/// One struct of the diff: where its bytes start, its first clock and its clock length.
#[derive(Clone, Copy)]
struct Struct {
    offset: usize,
    clock: u32,
    len: u32,
}

struct Section {
    client: u64,
    structs: Vec<Struct>,
    end: usize,
}

/// Splits `diff` into updates of at most `max_bytes` each; only an update holding a single struct,
/// or a single delete range, larger than the limit exceeds it. Struct updates come first and the
/// trailing updates carry only the delete set (one, possibly empty, when it fits).
pub(crate) fn split_diff_v1(diff: &[u8], max_bytes: usize) -> Result<Vec<Vec<u8>>, SplitError> {
    if max_bytes == 0 {
        return Err(SplitError::Malformed);
    }
    let (sections, delete_set) = parse(diff)?;
    let mut packer = Packer::new(diff, max_bytes);
    schedule(diff, &sections, |section, index| {
        packer.push(&sections[section], index)
    })?;
    let mut parts = packer.finish()?;
    parts.extend(delete_set.encode(max_bytes));
    Ok(parts)
}

fn parse(diff: &[u8]) -> Result<(Vec<Section>, DeleteSet), SplitError> {
    let mut scanner = Scanner::new(diff);
    let mut sections = Vec::new();
    for _ in 0..scanner.count()? {
        let count = scanner.count()?;
        let client = scanner.var()?;
        let mut clock = scanner.var_u32()?;
        let mut structs = Vec::with_capacity(count as usize);
        for _ in 0..count {
            let offset = scanner.pos;
            let len = scanner.block()?.len;
            structs.push(Struct { offset, clock, len });
            clock = clock.checked_add(len).ok_or(SplitError::Malformed)?;
        }
        if structs.is_empty() {
            return Err(SplitError::Malformed);
        }
        sections.push(Section {
            client,
            structs,
            end: scanner.pos,
        });
    }
    let delete_set = DeleteSet::scan(&mut scanner)?;
    if scanner.pos != diff.len() {
        return Err(SplitError::Malformed);
    }
    Ok((sections, delete_set))
}

/// Emits every struct once, in an order where each struct's origins and parent are either
/// already emitted or below the peer's clock (the clock a section starts at).
fn schedule(
    diff: &[u8],
    sections: &[Section],
    mut emit: impl FnMut(usize, usize) -> Result<(), SplitError>,
) -> Result<(), SplitError> {
    let index: HashMap<u64, usize> = sections
        .iter()
        .enumerate()
        .map(|(position, section)| (section.client, position))
        .collect();
    if index.len() != sections.len() {
        return Err(SplitError::Malformed);
    }
    let mut next = vec![0_usize; sections.len()];
    let mut emitted: Vec<u32> = sections
        .iter()
        .map(|section| section.structs[0].clock)
        .collect();
    let mut ready: VecDeque<usize> = (0..sections.len()).collect();
    let mut waiting: Vec<BTreeMap<u32, Vec<usize>>> = vec![BTreeMap::new(); sections.len()];
    while let Some(position) = ready.pop_front() {
        let section = &sections[position];
        let before = next[position];
        while let Some(item) = section.structs.get(next[position]) {
            let mut scanner = Scanner::new(diff);
            scanner.pos = item.offset;
            let block = scanner.block()?;
            let missing = block
                .dependencies
                .into_iter()
                .flatten()
                .find(|(client, clock)| {
                    index
                        .get(client)
                        .is_some_and(|&owner| *clock >= emitted[owner])
                });
            if let Some((client, clock)) = missing {
                if client == section.client && clock >= item.clock {
                    return Err(SplitError::Unordered);
                }
                waiting[index[&client]]
                    .entry(clock)
                    .or_default()
                    .push(position);
                break;
            }
            emit(position, next[position])?;
            emitted[position] = item.clock + item.len;
            next[position] += 1;
        }
        if next[position] == before {
            continue;
        }
        let woken = waiting[position].split_off(&emitted[position]);
        for (_, dependents) in std::mem::replace(&mut waiting[position], woken) {
            ready.extend(dependents);
        }
    }
    if next
        .iter()
        .zip(sections)
        .any(|(next, section)| *next != section.structs.len())
    {
        return Err(SplitError::Unordered);
    }
    Ok(())
}

/// One client's contiguous run inside the update being filled.
struct Run {
    clock: u32,
    end_clock: u32,
    count: u32,
    bytes: std::ops::Range<usize>,
}

struct Packer<'a> {
    diff: &'a [u8],
    max_bytes: usize,
    runs: BTreeMap<u64, Run>,
    size: usize,
    parts: Vec<Vec<u8>>,
}

impl<'a> Packer<'a> {
    fn new(diff: &'a [u8], max_bytes: usize) -> Self {
        Self {
            diff,
            max_bytes,
            runs: BTreeMap::new(),
            size: EMPTY_UPDATE_BYTES,
            parts: Vec::new(),
        }
    }

    fn push(&mut self, section: &Section, index: usize) -> Result<(), SplitError> {
        let item = section.structs[index];
        let end = section
            .structs
            .get(index + 1)
            .map_or(section.end, |next| next.offset);
        let bytes = item.offset..end;
        if !self.runs.is_empty()
            && self.size + self.growth(section.client, &item, &bytes) > self.max_bytes
        {
            self.flush()?;
        }
        let growth = self.growth(section.client, &item, &bytes);
        self.size += growth;
        match self.runs.get_mut(&section.client) {
            Some(run) => {
                if run.end_clock != item.clock || run.bytes.end != bytes.start {
                    return Err(SplitError::Unordered);
                }
                run.end_clock += item.len;
                run.count += 1;
                run.bytes.end = bytes.end;
            }
            None => {
                self.runs.insert(
                    section.client,
                    Run {
                        clock: item.clock,
                        end_clock: item.clock + item.len,
                        count: 1,
                        bytes,
                    },
                );
            }
        }
        Ok(())
    }

    /// Bytes the update grows by when `item` joins it.
    fn growth(&self, client: u64, item: &Struct, bytes: &std::ops::Range<usize>) -> usize {
        let len = bytes.len();
        match self.runs.get(&client) {
            Some(run) => len + var_len(u64::from(run.count) + 1) - var_len(u64::from(run.count)),
            None => {
                let runs = self.runs.len() as u64;
                len + var_len(1)
                    + var_len(client)
                    + var_len(u64::from(item.clock))
                    + var_len(runs + 1)
                    - var_len(runs)
            }
        }
    }

    fn flush(&mut self) -> Result<(), SplitError> {
        let runs = std::mem::take(&mut self.runs);
        let mut part = Vec::with_capacity(self.size);
        write_var(&mut part, runs.len() as u64);
        for (client, run) in runs.iter().rev() {
            write_var(&mut part, u64::from(run.count));
            write_var(&mut part, *client);
            write_var(&mut part, u64::from(run.clock));
            part.extend_from_slice(&self.diff[run.bytes.clone()]);
        }
        write_var(&mut part, 0);
        verify(&part, &runs)?;
        self.size = EMPTY_UPDATE_BYTES;
        self.parts.push(part);
        Ok(())
    }

    fn finish(mut self) -> Result<Vec<Vec<u8>>, SplitError> {
        if !self.runs.is_empty() {
            self.flush()?;
        }
        Ok(self.parts)
    }
}

/// Decodes a part with yrs and checks it holds exactly the clock ranges the split assigned.
fn verify(part: &[u8], runs: &BTreeMap<u64, Run>) -> Result<(), SplitError> {
    let update = Update::decode_v1(part).map_err(|_| SplitError::Malformed)?;
    let mut expected = IdSet::new();
    for (client, run) in runs {
        expected.insert(
            ID::new(ClientID::new(*client), run.clock),
            run.end_clock - run.clock,
        );
    }
    let actual = update.insertions(true);
    if !expected.diff(&actual).is_empty() || !actual.diff(&expected).is_empty() {
        return Err(SplitError::Malformed);
    }
    Ok(())
}

/// `0` struct sections and `0` delete-set clients.
const EMPTY_UPDATE_BYTES: usize = 2;

struct DeleteSet {
    clients: Vec<(u64, Vec<(u32, u32)>)>,
    bytes: Vec<u8>,
}

impl DeleteSet {
    fn scan(scanner: &mut Scanner<'_>) -> Result<Self, SplitError> {
        let start = scanner.pos;
        let mut clients = Vec::new();
        for _ in 0..scanner.count()? {
            let client = scanner.var()?;
            let mut ranges = Vec::new();
            for _ in 0..scanner.count()? {
                ranges.push((scanner.var_u32()?, scanner.var_u32()?));
            }
            clients.push((client, ranges));
        }
        Ok(Self {
            clients,
            bytes: scanner.bytes[start..scanner.pos].to_vec(),
        })
    }

    /// Delete-set-only updates of at most `max_bytes` each; one update when it fits.
    fn encode(self, max_bytes: usize) -> Vec<Vec<u8>> {
        if 1 + self.bytes.len() <= max_bytes {
            let mut part = Vec::with_capacity(1 + self.bytes.len());
            part.push(0);
            part.extend_from_slice(&self.bytes);
            return vec![part];
        }
        let mut parts = Vec::new();
        let mut current: Vec<(u64, Vec<(u32, u32)>)> = Vec::new();
        let mut size = EMPTY_UPDATE_BYTES;
        for (client, ranges) in self.clients {
            for (clock, len) in ranges {
                let range = var_len(u64::from(clock)) + var_len(u64::from(len));
                let same_client = current.last().is_some_and(|(last, _)| *last == client);
                let growth = if same_client {
                    let count = current.last().map_or(0, |(_, ranges)| ranges.len()) as u64;
                    range + var_len(count + 1) - var_len(count)
                } else {
                    let count = current.len() as u64;
                    range + var_len(client) + var_len(1) + var_len(count + 1) - var_len(count)
                };
                if !current.is_empty() && size + growth > max_bytes {
                    parts.push(encode_delete_set(&current));
                    current.clear();
                    size = EMPTY_UPDATE_BYTES;
                }
                match current.last_mut() {
                    Some((last, ranges)) if *last == client => {
                        size +=
                            range + var_len(ranges.len() as u64 + 1) - var_len(ranges.len() as u64);
                        ranges.push((clock, len));
                    }
                    _ => {
                        let count = current.len() as u64;
                        size += range + var_len(client) + var_len(1) + var_len(count + 1)
                            - var_len(count);
                        current.push((client, vec![(clock, len)]));
                    }
                }
            }
        }
        if !current.is_empty() || parts.is_empty() {
            parts.push(encode_delete_set(&current));
        }
        parts
    }
}

fn encode_delete_set(clients: &[(u64, Vec<(u32, u32)>)]) -> Vec<u8> {
    let mut part = vec![0];
    write_var(&mut part, clients.len() as u64);
    for (client, ranges) in clients {
        write_var(&mut part, *client);
        write_var(&mut part, ranges.len() as u64);
        for (clock, len) in ranges {
            write_var(&mut part, u64::from(*clock));
            write_var(&mut part, u64::from(*len));
        }
    }
    part
}

fn var_len(mut value: u64) -> usize {
    let mut len = 1;
    while value >= 128 {
        len += 1;
        value >>= 7;
    }
    len
}

fn write_var(bytes: &mut Vec<u8>, mut value: u64) {
    while value >= 128 {
        bytes.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    bytes.push(value as u8);
}

struct Block {
    len: u32,
    dependencies: [Option<(u64, u32)>; 3],
}

struct Scanner<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> Scanner<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn byte(&mut self) -> Result<u8, SplitError> {
        let byte = *self.bytes.get(self.pos).ok_or(SplitError::Malformed)?;
        self.pos += 1;
        Ok(byte)
    }

    fn take(&mut self, len: usize) -> Result<&'a [u8], SplitError> {
        let end = self.pos.checked_add(len).ok_or(SplitError::Malformed)?;
        let bytes = self.bytes.get(self.pos..end).ok_or(SplitError::Malformed)?;
        self.pos = end;
        Ok(bytes)
    }

    fn var(&mut self) -> Result<u64, SplitError> {
        let mut value = 0_u64;
        let mut shift = 0;
        loop {
            let byte = self.byte()?;
            let payload = u64::from(byte & 0x7f);
            if shift >= 64 || payload > (u64::MAX >> shift) {
                return Err(SplitError::Malformed);
            }
            value |= payload << shift;
            if byte & 0x80 == 0 {
                return Ok(value);
            }
            shift += 7;
        }
    }

    fn var_u32(&mut self) -> Result<u32, SplitError> {
        u32::try_from(self.var()?).map_err(|_| SplitError::Malformed)
    }

    fn count(&mut self) -> Result<u32, SplitError> {
        let count = self.var_u32()?;
        if count as usize > self.bytes.len() - self.pos {
            return Err(SplitError::Malformed);
        }
        Ok(count)
    }

    fn id(&mut self) -> Result<(u64, u32), SplitError> {
        Ok((self.var()?, self.var_u32()?))
    }

    fn buffer(&mut self) -> Result<&'a [u8], SplitError> {
        let len = usize::try_from(self.var()?).map_err(|_| SplitError::Malformed)?;
        self.take(len)
    }

    fn string(&mut self) -> Result<&'a str, SplitError> {
        std::str::from_utf8(self.buffer()?).map_err(|_| SplitError::Malformed)
    }

    fn signed_var(&mut self) -> Result<(), SplitError> {
        let mut byte = self.byte()?;
        let mut shift = 6;
        while byte & 0x80 != 0 {
            if shift >= 64 {
                return Err(SplitError::Malformed);
            }
            byte = self.byte()?;
            shift += 7;
        }
        Ok(())
    }

    fn any(&mut self, depth: u8) -> Result<(), SplitError> {
        if depth >= MAX_NESTING {
            return Err(SplitError::Malformed);
        }
        match self.byte()? {
            127 | 126 | 121 | 120 => {}
            125 => self.signed_var()?,
            124 => {
                self.take(4)?;
            }
            123 | 122 => {
                self.take(8)?;
            }
            119 => {
                self.string()?;
            }
            118 => {
                for _ in 0..self.count()? {
                    self.string()?;
                    self.any(depth + 1)?;
                }
            }
            117 => {
                for _ in 0..self.count()? {
                    self.any(depth + 1)?;
                }
            }
            116 => {
                self.buffer()?;
            }
            _ => return Err(SplitError::Malformed),
        }
        Ok(())
    }

    /// Reads one struct, returning its clock length and the ids it depends on.
    fn block(&mut self) -> Result<Block, SplitError> {
        let info = self.byte()?;
        let mut dependencies = [None; 3];
        if info == BLOCK_SKIP_REF_NUMBER {
            return Err(SplitError::Unsupported);
        }
        if info == BLOCK_GC_REF_NUMBER {
            let len = self.var_u32()?;
            return if len == 0 {
                Err(SplitError::Malformed)
            } else {
                Ok(Block { len, dependencies })
            };
        }
        if info & HAS_ORIGIN != 0 {
            dependencies[0] = Some(self.id()?);
        }
        if info & HAS_RIGHT_ORIGIN != 0 {
            dependencies[1] = Some(self.id()?);
        }
        if info & (HAS_ORIGIN | HAS_RIGHT_ORIGIN) == 0 {
            match self.var()? {
                0 => dependencies[2] = Some(self.id()?),
                1 => {
                    self.string()?;
                }
                _ => return Err(SplitError::Malformed),
            }
            if info & HAS_PARENT_SUB != 0 {
                self.string()?;
            }
        }
        let len = match info & 0x1f {
            BLOCK_ITEM_DELETED_REF_NUMBER => self.var_u32()?,
            BLOCK_ITEM_JSON_REF_NUMBER => {
                let count = self.count()?;
                for _ in 0..count {
                    self.string()?;
                }
                count
            }
            BLOCK_ITEM_BINARY_REF_NUMBER => {
                self.buffer()?;
                1
            }
            BLOCK_ITEM_STRING_REF_NUMBER => u32::try_from(self.string()?.encode_utf16().count())
                .map_err(|_| SplitError::Malformed)?,
            BLOCK_ITEM_EMBED_REF_NUMBER => {
                self.string()?;
                1
            }
            BLOCK_ITEM_FORMAT_REF_NUMBER => {
                self.string()?;
                self.string()?;
                1
            }
            BLOCK_ITEM_TYPE_REF_NUMBER => {
                match self.byte()? {
                    TYPE_REFS_ARRAY
                    | TYPE_REFS_MAP
                    | TYPE_REFS_TEXT
                    | TYPE_REFS_XML_FRAGMENT
                    | TYPE_REFS_XML_HOOK
                    | TYPE_REFS_XML_TEXT
                    | TYPE_REFS_DOC
                    | TYPE_REFS_UNDEFINED => {}
                    TYPE_REFS_XML_ELEMENT => {
                        self.string()?;
                    }
                    _ => return Err(SplitError::Unsupported),
                }
                1
            }
            BLOCK_ITEM_ANY_REF_NUMBER => {
                let count = self.count()?;
                for _ in 0..count {
                    self.any(0)?;
                }
                count
            }
            BLOCK_ITEM_DOC_REF_NUMBER => {
                self.string()?;
                self.any(0)?;
                1
            }
            _ => return Err(SplitError::Unsupported),
        };
        if len == 0 {
            return Err(SplitError::Malformed);
        }
        Ok(Block { len, dependencies })
    }
}

/// The ordered updates that bring a peer at `state_vector` up to `txn`'s state.
pub(crate) fn encode_diff_chunks_v1<T: ReadTxn>(
    txn: &T,
    state_vector: &StateVector,
    max_bytes: usize,
) -> Vec<Vec<u8>> {
    let diff = crate::deterministic::encode_diff_v1(txn, state_vector);
    if diff.len() <= max_bytes {
        return vec![diff];
    }
    match split_diff_v1(&diff, max_bytes) {
        Ok(parts) => parts,
        Err(_) => vec![diff],
    }
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use yrs::{Any, Map, Transact};

    use super::*;
    use crate::{
        EditCtx, EditError, EditingDoc, FormatPolicy, InlineFormatDelta, MAX_UPDATE_BYTES, Patch,
        Position, StoryRange, StorySegment,
    };

    const DATE: &str = "2026-10-16T12:00:00Z";
    const LIMITS: [usize; 6] = [1, 64, 512, 4096, 65536, usize::MAX];

    fn ctx(author: &str) -> EditCtx {
        EditCtx::local(author, DATE)
    }

    fn replica(client_id: u64, state: &[u8]) -> EditingDoc {
        let doc = EditingDoc::new(client_id);
        doc.apply_verbatim_v1(state).unwrap();
        doc
    }

    fn document(client_id: u64) -> EditingDoc {
        let doc = EditingDoc::new(client_id);
        doc.create_story("body", "first paragraph", "Normal", "left")
            .unwrap();
        doc.create_story("header:rId1", "Header", "Header", "center")
            .unwrap();
        doc
    }

    fn write_paragraphs(doc: &EditingDoc, author: &str, count: usize) {
        for n in 0..count {
            let text = format!("{author} paragraph {n} with some text to carry ");
            doc.insert_text(
                &ctx(author),
                Position::new("body", 0),
                &text,
                FormatPolicy::Plain,
            )
            .unwrap();
            doc.split_paragraph(
                &ctx(author),
                Position::new("body", text.encode_utf16().count() as u32),
                None,
            )
            .unwrap();
        }
    }

    fn bold(doc: &EditingDoc, author: &str, start: u32, end: u32) {
        let delta = InlineFormatDelta {
            bold: Patch::Set(true),
            ..InlineFormatDelta::default()
        };
        doc.format_range(&ctx(author), StoryRange::new("body", start, end), &delta)
            .unwrap();
    }

    fn exchange(from: &EditingDoc, to: &EditingDoc) {
        let diff = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
        to.apply_update_v1(&diff).unwrap();
    }

    fn stories(doc: &EditingDoc) -> BTreeMap<String, Vec<StorySegment>> {
        let ids: Vec<String> = {
            let txn = doc.yrs_doc().transact();
            txn.get_map(crate::STORIES)
                .map(|stories| {
                    crate::live_stories(&stories, &txn)
                        .map(|(id, _)| id.to_owned())
                        .collect()
                })
                .unwrap_or_default()
        };
        ids.into_iter()
            .map(|id| {
                let segments = doc.story_segments(&id).unwrap();
                (id, segments)
            })
            .collect()
    }

    fn assert_same(actual: &EditingDoc, expected: &EditingDoc) {
        assert_eq!(
            actual.encode_state_vector_v1(),
            expected.encode_state_vector_v1()
        );
        assert_eq!(stories(actual), stories(expected));
        assert_eq!(
            actual.encode_state_as_update_v1(),
            expected.encode_state_as_update_v1()
        );
    }

    fn assert_integrated(doc: &EditingDoc) {
        let txn = doc.yrs_doc().transact();
        assert!(txn.store().pending_update().is_none());
        assert!(txn.store().pending_ds().is_none());
    }

    fn structs_in(part: &[u8]) -> usize {
        let mut scanner = Scanner::new(part);
        let mut total = 0;
        for _ in 0..scanner.count().unwrap() {
            let count = scanner.count().unwrap();
            scanner.var().unwrap();
            scanner.var_u32().unwrap();
            for _ in 0..count {
                scanner.block().unwrap();
            }
            total += count as usize;
        }
        total
    }

    fn same_ids(left: &IdSet, right: &IdSet) -> bool {
        left.diff(right).is_empty() && right.diff(left).is_empty()
    }

    /// Splits `source`'s diff for `peer` at every limit and checks the parts against the whole
    /// diff: shape, size, stepwise integration, and the atomic path. Returns the parts per limit.
    fn assert_split(source: &EditingDoc, peer: &EditingDoc) -> Vec<(usize, Vec<Vec<u8>>)> {
        let vector = peer.encode_state_vector_v1();
        let base = peer.encode_state_as_update_v1();
        let whole = source.encode_diff_v1(&vector).unwrap();
        let expected = replica(9_001, &base);
        expected.apply_verbatim_v1(&whole).unwrap();
        let whole_delete_set = Update::decode_v1(&whole).unwrap().delete_set().clone();
        let mut results = Vec::new();
        for limit in LIMITS {
            let parts = source.encode_diff_chunks_v1(&vector, limit).unwrap();
            if whole.len() <= limit {
                assert_eq!(parts, vec![whole.clone()], "limit {limit}");
                results.push((limit, parts));
                continue;
            }
            let structs = parts.iter().take_while(|part| part[0] != 0).count();
            let mut deletes = IdSet::new();
            for (index, part) in parts.iter().enumerate() {
                let update = Update::decode_v1(part).unwrap();
                if index < structs {
                    assert!(update.delete_set().is_empty(), "limit {limit} part {index}");
                    assert!(
                        part.len() <= limit || structs_in(part) == 1,
                        "limit {limit} part {index} holds {} bytes",
                        part.len()
                    );
                } else {
                    assert_eq!(structs_in(part), 0, "limit {limit} part {index}");
                    assert!(part.len() <= limit.max(16), "limit {limit} part {index}");
                    deletes.merge_with(update.delete_set().clone());
                }
            }
            assert!(structs < parts.len(), "limit {limit}: no delete-set part");
            assert!(same_ids(&deletes, &whole_delete_set), "limit {limit}");

            let stepwise = replica(9_001, &base);
            for part in &parts {
                stepwise.apply_verbatim_v1(part).unwrap();
                assert_integrated(&stepwise);
            }
            assert_same(&stepwise, &expected);

            let whole_peer = replica(9_002, &base);
            whole_peer.apply_update_v1(&whole).unwrap();
            let atomic = replica(9_002, &base);
            atomic.apply_update_chunks_v1(&parts).unwrap();
            assert_integrated(&atomic);
            assert_same(&atomic, &whole_peer);
            results.push((limit, parts));
        }
        results
    }

    #[test]
    fn text_formatting_and_tables_split_into_complete_parts() {
        let source = document(1);
        write_paragraphs(&source, "A", 12);
        bold(&source, "A", 2, 9);
        let table = source
            .insert_table(&ctx("A"), Position::new("body", 0), 2, 2)
            .unwrap();
        for (n, cell) in table.created_story_ids.iter().enumerate() {
            source
                .insert_text(
                    &ctx("A"),
                    Position::new(cell.as_str(), 0),
                    &format!("cell {n}"),
                    FormatPolicy::Plain,
                )
                .unwrap();
        }
        source
            .insert_text(
                &ctx("A"),
                Position::new("header:rId1", 0),
                "Draft ",
                FormatPolicy::Plain,
            )
            .unwrap();
        let fresh = EditingDoc::new(2);
        let parts = assert_split(&source, &fresh);
        assert!(parts.iter().any(|(_, parts)| parts.len() > 8));

        let midway = replica(3, &source.encode_state_as_update_v1());
        write_paragraphs(&source, "A", 6);
        bold(&source, "A", 0, 20);
        assert_split(&source, &midway);
    }

    #[test]
    fn deletions_and_undo_history_travel_in_the_final_parts() {
        let source = document(1);
        let mut undo = source.undo_manager();
        write_paragraphs(&source, "A", 8);
        undo.add_undo_barrier();
        source
            .delete_range(&ctx("A"), StoryRange::new("body", 4, 40))
            .unwrap();
        undo.add_undo_barrier();
        write_paragraphs(&source, "A", 3);
        assert!(undo.undo());
        assert!(undo.undo());
        assert!(undo.redo());
        let midway = replica(3, &source.encode_state_as_update_v1());
        source
            .delete_range(&ctx("A"), StoryRange::new("body", 2, 12))
            .unwrap();
        let whole = source.encode_diff_v1(&EditingDoc::new(2).encode_state_vector_v1());
        assert!(
            !Update::decode_v1(&whole.unwrap())
                .unwrap()
                .delete_set()
                .is_empty()
        );
        assert_split(&source, &EditingDoc::new(2));
        assert_split(&source, &midway);
    }

    #[test]
    fn a_struct_over_the_limit_travels_alone() {
        let source = document(1);
        write_paragraphs(&source, "A", 4);
        let picture = "data:image/png;base64,".to_owned() + &"QUJD".repeat(50_000);
        source
            .insert_embed(
                &ctx("A"),
                Position::new("body", 3),
                "image",
                vec![("src".to_owned(), Any::from(picture))],
            )
            .unwrap();
        write_paragraphs(&source, "A", 4);
        let results = assert_split(&source, &EditingDoc::new(2));
        let (_, parts) = results.iter().find(|(limit, _)| *limit == 65_536).unwrap();
        let oversized: Vec<_> = parts.iter().filter(|part| part.len() > 65_536).collect();
        assert_eq!(oversized.len(), 1);
        assert_eq!(structs_in(oversized[0]), 1);
    }

    #[test]
    fn multi_client_histories_split_into_complete_parts() {
        let seed = document(1).encode_state_as_update_v1();
        let (a, b, c) = (replica(11, &seed), replica(12, &seed), replica(13, &seed));
        write_paragraphs(&a, "A", 4);
        write_paragraphs(&b, "B", 3);
        exchange(&a, &b);
        write_paragraphs(&b, "B", 2);
        bold(&b, "B", 1, 30);
        exchange(&b, &c);
        write_paragraphs(&c, "C", 3);
        c.delete_range(&ctx("C"), StoryRange::new("body", 5, 25))
            .unwrap();
        exchange(&c, &a);
        write_paragraphs(&a, "A", 2);
        let partial = replica(14, &b.encode_state_as_update_v1());
        exchange(&a, &c);
        exchange(&b, &c);
        assert_split(&c, &EditingDoc::new(15));
        assert_split(&c, &partial);
        assert_split(&c, &replica(16, &seed));
    }

    #[test]
    fn large_delete_sets_split_by_range() {
        let source = document(1);
        source
            .insert_text(
                &ctx("A"),
                Position::new("body", 0),
                &"abcdefghij".repeat(300),
                FormatPolicy::Plain,
            )
            .unwrap();
        for index in (0..900).rev() {
            source
                .delete_range(&ctx("A"), StoryRange::new("body", index * 3, index * 3 + 1))
                .unwrap();
        }
        let results = assert_split(&source, &EditingDoc::new(2));
        let (_, parts) = results.iter().find(|(limit, _)| *limit == 512).unwrap();
        assert!(parts.iter().filter(|part| part[0] == 0).count() > 2);
    }

    #[test]
    fn diffs_that_cannot_be_split_are_refused() {
        assert_eq!(
            split_diff_v1(&[1, 1, 5, 0, BLOCK_SKIP_REF_NUMBER, 3, 0], 1),
            Err(SplitError::Unsupported)
        );
        let forward = [
            1,
            1,
            5,
            0,
            HAS_ORIGIN | BLOCK_ITEM_STRING_REF_NUMBER,
            5,
            9,
            1,
            b'a',
            0,
        ];
        assert_eq!(split_diff_v1(&forward, 1), Err(SplitError::Unordered));
        assert_eq!(split_diff_v1(&forward[..7], 1), Err(SplitError::Malformed));
        assert_eq!(split_diff_v1(&[0, 0, 0], 1), Err(SplitError::Malformed));
        assert_eq!(split_diff_v1(&[0, 0], 0), Err(SplitError::Malformed));
        assert_eq!(split_diff_v1(&[0, 0], 1), Ok(vec![vec![0, 0]]));
    }

    #[test]
    fn truncated_or_mutated_diffs_never_panic_or_hang() {
        let source = document(1);
        write_paragraphs(&source, "A", 6);
        bold(&source, "A", 1, 12);
        source
            .delete_range(&ctx("A"), StoryRange::new("body", 2, 9))
            .unwrap();
        let diff = source
            .encode_diff_v1(&EditingDoc::new(2).encode_state_vector_v1())
            .unwrap();
        assert!(split_diff_v1(&diff, 64).unwrap().len() > 2);
        for end in 0..diff.len() {
            assert!(split_diff_v1(&diff[..end], 64).is_err(), "prefix {end}");
        }
        let mut seed = 0x9e37_79b9_7f4a_7c15_u64;
        let mut next = || {
            seed ^= seed << 13;
            seed ^= seed >> 7;
            seed ^= seed << 17;
            seed
        };
        for _ in 0..5_000 {
            let mut mutated = diff.clone();
            for _ in 0..1 + next() % 4 {
                let index = (next() % mutated.len() as u64) as usize;
                mutated[index] = next() as u8;
            }
            if let Ok((sections, _)) = parse(&mutated) {
                let _ = schedule(&mutated, &sections, |_, _| Ok(()));
            }
        }
    }

    #[test]
    fn every_yrs_content_kind_splits_into_complete_parts() {
        use std::sync::Arc;
        use yrs::types::Attrs;
        use yrs::{
            Array, ClientID, Doc, MapPrelim, OffsetKind, Options, Text, WriteTxn, Xml,
            XmlElementPrelim, XmlFragment, XmlTextPrelim,
        };

        let replica = |client_id: u64| {
            let mut options = Options::with_client_id(ClientID::new(client_id));
            options.offset_kind = OffsetKind::Utf16;
            Doc::with_options(options)
        };
        let source = replica(7);
        {
            let mut txn = source.transact_mut();
            let map = txn.get_or_insert_map("map");
            map.insert(&mut txn, "float", 1.5);
            map.insert(&mut txn, "big", Any::BigInt(1 << 40));
            map.insert(&mut txn, "buffer", Any::Buffer(Arc::from([1, 2, 3])));
            map.insert(
                &mut txn,
                "nested",
                Any::Map(Arc::new(std::collections::HashMap::from([(
                    "a".to_owned(),
                    Any::Array(Arc::from([Any::from(2.5), Any::Null, Any::Bool(true)])),
                )]))),
            );
            map.insert(&mut txn, "sub", Doc::new());
            let inner = map.insert(&mut txn, "inner", MapPrelim::default());
            inner.insert(&mut txn, "x", "y");
            let array = txn.get_or_insert_array("array");
            array.insert_range(
                &mut txn,
                0,
                (0..200).map(|n| Any::from(f64::from(n) + 0.25)),
            );
            let text = txn.get_or_insert_text("text");
            text.insert(&mut txn, 0, "hello world 😀 and more text");
            text.format(
                &mut txn,
                0,
                5,
                Attrs::from([("bold".into(), Any::Bool(true))]),
            );
            text.insert_embed(&mut txn, 3, Any::from("embed"));
            let xml = txn.get_or_insert_xml_fragment("xml");
            let element = xml.insert(&mut txn, 0, XmlElementPrelim::empty("p"));
            element.insert_attribute(&mut txn, "class", "c");
            element.insert(&mut txn, 0, XmlTextPrelim::new("xml text"));
        }
        let other = replica(8);
        other
            .transact_mut()
            .apply_update(
                Update::decode_v1(
                    &source
                        .transact()
                        .encode_state_as_update_v1(&StateVector::default()),
                )
                .unwrap(),
            )
            .unwrap();
        {
            let mut txn = other.transact_mut();
            txn.get_or_insert_text("text")
                .insert(&mut txn, 6, "concurrent ");
            txn.get_or_insert_array("array")
                .remove_range(&mut txn, 10, 50);
        }
        {
            let mut txn = source.transact_mut();
            txn.get_or_insert_map("map").remove(&mut txn, "inner");
            txn.get_or_insert_text("text").remove_range(&mut txn, 6, 3);
        }
        let from_other = other
            .transact()
            .encode_diff_v1(&source.transact().state_vector());
        source
            .transact_mut()
            .apply_update(Update::decode_v1(&from_other).unwrap())
            .unwrap();
        let diff = source.transact().encode_diff_v1(&StateVector::default());
        let whole = replica(9);
        whole
            .transact_mut()
            .apply_update(Update::decode_v1(&diff).unwrap())
            .unwrap();
        let expected = crate::deterministic::encode_state_as_update_v1(
            &whole.transact(),
            &StateVector::default(),
        );
        for limit in [1, 64, 512, 4096] {
            let parts = split_diff_v1(&diff, limit).unwrap();
            assert!(parts.len() > 2 || limit == 4096, "limit {limit}");
            let stepwise = replica(9);
            for part in &parts {
                stepwise
                    .transact_mut()
                    .apply_update(Update::decode_v1(part).unwrap())
                    .unwrap();
                let txn = stepwise.transact();
                assert!(txn.store().pending_update().is_none(), "limit {limit}");
                assert!(txn.store().pending_ds().is_none(), "limit {limit}");
            }
            assert_eq!(
                crate::deterministic::encode_state_as_update_v1(
                    &stepwise.transact(),
                    &StateVector::default(),
                ),
                expected,
                "limit {limit}"
            );
        }
    }

    #[test]
    fn chunk_limits_and_oversized_chunks_are_refused() {
        let source = document(1);
        let peer = EditingDoc::new(2);
        assert!(matches!(
            source.encode_diff_chunks_v1(&peer.encode_state_vector_v1(), 0),
            Err(EditError::InvalidUpdate(_))
        ));
        assert!(matches!(
            source.encode_diff_chunks_v1(&[0xff], 64),
            Err(EditError::InvalidStateVector(_))
        ));
        let parts = source
            .encode_diff_chunks_v1(&peer.encode_state_vector_v1(), 64)
            .unwrap();
        let mut oversized = parts[0].clone();
        oversized.resize(MAX_UPDATE_BYTES + 1, 0);
        let before = peer.encode_state_as_update_v1();
        assert_eq!(
            peer.apply_update_chunks_v1(&[parts[0].clone(), oversized]),
            Err(EditError::InvalidUpdate(format!(
                "update exceeds {MAX_UPDATE_BYTES} bytes"
            )))
        );
        assert_eq!(peer.encode_state_as_update_v1(), before);
        peer.apply_update_chunks_v1(&parts).unwrap();
        assert_eq!(
            peer.encode_state_vector_v1(),
            source.encode_state_vector_v1()
        );
    }
}
