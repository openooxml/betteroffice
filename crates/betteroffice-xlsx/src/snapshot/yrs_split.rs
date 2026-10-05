use std::collections::{BTreeMap, BTreeSet, HashMap, VecDeque};

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

const MAX_CLIENT_ID: u64 = (1_u64 << 53) - 1;
const MAX_CLOCK: u32 = i32::MAX as u32;
const MAX_NESTING: u8 = 64;
pub(crate) const SHEET_ORDER_MAX_ITEMS: usize = 16_384;

#[cfg(test)]
thread_local! {
    pub(crate) static WHOLE_UPDATE_LIMIT: std::cell::Cell<usize> = const { std::cell::Cell::new(64 * 1024 * 1024 - 12) };
    static ADMISSION_ATTEMPTS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SplitError {
    InvalidLimit,
    Malformed,
    UnsupportedContent(u8),
    UnsupportedType(u8),
    UnsupportedMap,
    JsonLengthMismatch,
    MissingDependency,
    RetainedDeletion,
    OversizedStruct,
    SharedText,
    SheetOrderLimit,
}

impl SplitError {
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn reason(self) -> &'static str {
        match self {
            Self::InvalidLimit => "invalid_limit",
            Self::Malformed => "malformed",
            Self::UnsupportedContent(_) => "unsupported_content",
            Self::UnsupportedType(_) => "unsupported_type",
            Self::UnsupportedMap => "multi_key_map",
            Self::JsonLengthMismatch => "json_length_mismatch",
            Self::MissingDependency => "missing_dependency",
            Self::RetainedDeletion => "retained_deletion",
            Self::OversizedStruct => "oversized_struct",
            Self::SharedText => "shared_text_is_not_supported",
            Self::SheetOrderLimit => "sheet_order_exceeds_item_or_clock_limit",
        }
    }
}

#[derive(Default)]
struct Client {
    clock: u32,
    ranges: Vec<(u32, u32, u8)>,
    delete_offset: Option<usize>,
    delete_count: u32,
}

struct Struct<'a> {
    len: u32,
    kind: u8,
    dependencies: [Option<(u64, u32)>; 3],
    root: Option<&'a str>,
    key: Option<&'a str>,
    sheet_keys: bool,
}

pub(crate) fn split_update_v1(
    update: &[u8],
    max_part_bytes: usize,
) -> Result<Vec<Vec<u8>>, SplitError> {
    split_update_v1_parts(update, usize::MAX, max_part_bytes)
}

pub(crate) fn split_update_v1_bounded(
    update: &[u8],
    max_records: usize,
    max_part_bytes: usize,
) -> Result<Vec<Vec<u8>>, SplitError> {
    if max_records == 0 || max_part_bytes == 0 {
        return Err(SplitError::InvalidLimit);
    }
    if max_part_bytes < 2 {
        return Err(SplitError::OversizedStruct);
    }
    let parts = split_update_v1_parts(update, max_records, max_part_bytes)?;
    let mut bounded = Vec::new();
    for part in parts {
        let mut cursor = UpdateCursor::default();
        while let Some(part) = cursor.next_oversized(&part, max_records, max_part_bytes)? {
            bounded.push(part.bytes);
        }
    }
    if bounded.is_empty() {
        bounded.push(vec![0, 0]);
    }
    Ok(bounded)
}

pub(crate) fn whole_update_limit() -> usize {
    #[cfg(test)]
    {
        WHOLE_UPDATE_LIMIT.get()
    }
    #[cfg(not(test))]
    {
        64 * 1024 * 1024 - 12
    }
}

pub(crate) fn split_fallback_v1_bounded(
    update: &[u8],
    max_records: usize,
    max_bytes: usize,
) -> Result<Vec<Vec<u8>>, SplitError> {
    let mut scanner = Scanner::new(update);
    scanner.allow_maps = true;
    let clients = scanner.count()?;
    for _ in 0..clients {
        let count = scanner.count()?;
        scanner.client()?;
        scanner.clock()?;
        for _ in 0..count {
            scanner.block()?;
        }
    }
    let clients = scanner.count()?;
    let mut oversized_deletion = false;
    for _ in 0..clients {
        scanner.client()?;
        let count = scanner.count()?;
        for _ in 0..count {
            let (start, end) = scanner.delete_range()?;
            oversized_deletion |= (end - start) as usize > max_records;
        }
    }
    if scanner.pos != update.len() {
        return Err(SplitError::Malformed);
    }
    if !oversized_deletion
        && update.len() <= whole_update_limit()
        && CausalState::default().admit(update).is_ok()
    {
        return Ok(vec![update.to_vec()]);
    }
    let mut cursor = UpdateCursor::default();
    let mut clients: BTreeMap<u64, VecDeque<Vec<u8>>> = BTreeMap::new();
    let mut deletes = Vec::new();
    while let Some(part) = cursor.next_oversized(update, 1, max_bytes)? {
        if part.bytes[0] == 0 {
            deletes.push(part.bytes);
        } else {
            let mut scanner = Scanner::new(&part.bytes);
            scanner.count()?;
            scanner.count()?;
            let client = scanner.client()?;
            clients.entry(client).or_default().push_back(part.bytes);
        }
    }
    let mut causal = CausalState::default();
    let mut parts = Vec::new();
    let mut ready: BTreeSet<(usize, u64)> = clients.keys().map(|&client| (0, client)).collect();
    let mut waiting: BTreeMap<(u64, u32), Vec<u64>> = BTreeMap::new();
    while let Some((round, client)) = ready.pop_first() {
        let queue = clients.get_mut(&client).ok_or(SplitError::Malformed)?;
        let part = queue.front().ok_or(SplitError::Malformed)?;
        match causal.admit(part) {
            Ok(()) => {
                parts.push(queue.pop_front().ok_or(SplitError::Malformed)?);
                if queue.is_empty() {
                    clients.remove(&client);
                } else {
                    ready.insert((round + 1, client));
                }
                let clock = causal.clocks[&client];
                while let Some((&dependency, _)) =
                    waiting.range((client, 0)..(client, clock)).next()
                {
                    for dependent in waiting.remove(&dependency).ok_or(SplitError::Malformed)? {
                        ready.insert((round + usize::from(dependent <= client), dependent));
                    }
                }
            }
            Err(SplitError::MissingDependency) => {
                let dependency = causal.waiting_on(part)?;
                waiting.entry(dependency).or_default().push(client);
            }
            Err(error) => return Err(error),
        }
    }
    if !clients.is_empty() {
        return Err(SplitError::MissingDependency);
    }
    for part in deletes {
        causal.admit(&part)?;
        parts.push(part);
    }
    if parts.is_empty() {
        parts.push(vec![0, 0]);
    }
    Ok(parts)
}

fn split_update_v1_parts(
    update: &[u8],
    max_records: usize,
    max_part_bytes: usize,
) -> Result<Vec<Vec<u8>>, SplitError> {
    if max_records == 0 || max_part_bytes == 0 {
        return Err(SplitError::InvalidLimit);
    }
    let mut scanner = Scanner::new(update);
    let client_count = scanner.count()?;
    let mut clients: HashMap<u64, Client> = HashMap::new();
    let mut previous_client = None;
    let mut parts = Vec::new();
    let mut has_deleted_structs = false;
    for _ in 0..client_count {
        let struct_count = scanner.count()?;
        let client = scanner.client()?;
        let mut clock = scanner.clock()?;
        if struct_count == 0 || previous_client.is_some_and(|previous| previous <= client) {
            return Err(SplitError::Malformed);
        }
        if clock != 0 {
            return Err(SplitError::MissingDependency);
        }
        previous_client = Some(client);
        clients.insert(client, Client::default());
        let mut run_start = scanner.pos;
        let mut run_clock = clock;
        let mut run_count = 0;
        for _ in 0..struct_count {
            let struct_start = scanner.pos;
            let block = scanner.block()?;
            if block.kind == BLOCK_SKIP_REF_NUMBER {
                return Err(SplitError::MissingDependency);
            }
            has_deleted_structs |= matches!(
                block.kind,
                BLOCK_GC_REF_NUMBER | BLOCK_ITEM_DELETED_REF_NUMBER
            );
            for (index, dependency) in block.dependencies.into_iter().enumerate() {
                let Some((dependency_client, dependency_clock)) = dependency else {
                    continue;
                };
                let known = clients
                    .get(&dependency_client)
                    .ok_or(SplitError::MissingDependency)?;
                if dependency_clock >= known.clock {
                    return Err(SplitError::MissingDependency);
                }
                let end = known
                    .ranges
                    .partition_point(|(start, _, _)| *start <= dependency_clock);
                let kind = known.ranges[..end]
                    .last()
                    .filter(|(start, len, _)| dependency_clock - *start < *len)
                    .map(|(_, _, kind)| *kind);
                if (index == 2 && kind != Some(BLOCK_ITEM_TYPE_REF_NUMBER))
                    || (index < 2 && kind == Some(BLOCK_GC_REF_NUMBER))
                {
                    return Err(SplitError::MissingDependency);
                }
            }
            let next_clock = clock.checked_add(block.len).ok_or(SplitError::Malformed)?;
            if next_clock > MAX_CLOCK {
                return Err(SplitError::Malformed);
            }
            let candidate_count = run_count + 1;
            let candidate_size = run_size(client, run_clock, candidate_count)
                .checked_add(scanner.pos - run_start)
                .ok_or(SplitError::Malformed)?;
            if run_count != 0
                && (candidate_size > max_part_bytes || run_count as usize == max_records)
            {
                parts.push(encode_run(
                    client,
                    run_clock,
                    run_count,
                    &update[run_start..struct_start],
                ));
                run_start = struct_start;
                run_clock = clock;
                run_count = 0;
            }
            run_count += 1;
            let known = clients.get_mut(&client).ok_or(SplitError::Malformed)?;
            if matches!(block.kind, BLOCK_GC_REF_NUMBER | BLOCK_ITEM_TYPE_REF_NUMBER) {
                if let Some((start, len, kind)) = known.ranges.last_mut()
                    && *kind == block.kind
                    && *start + *len == clock
                {
                    *len += block.len;
                } else {
                    known.ranges.push((clock, block.len, block.kind));
                }
            }
            clock = next_clock;
            known.clock = clock;
        }
        parts.push(encode_run(
            client,
            run_clock,
            run_count,
            &update[run_start..scanner.pos],
        ));
    }
    let delete_start = scanner.pos;
    let delete_clients = scanner.count()?;
    for _ in 0..delete_clients {
        let client = scanner.client()?;
        let count = scanner.count()?;
        let known = clients
            .get_mut(&client)
            .ok_or(SplitError::MissingDependency)?;
        if count == 0 || known.delete_offset.is_some() {
            return Err(SplitError::Malformed);
        }
        known.delete_offset = Some(scanner.pos);
        known.delete_count = count;
        let mut previous_end = None;
        for _ in 0..count {
            let (start, end) = scanner.delete_range()?;
            if end > known.clock {
                return Err(SplitError::MissingDependency);
            }
            if previous_end.is_some_and(|previous| start <= previous) {
                return Err(SplitError::Malformed);
            }
            previous_end = Some(end);
        }
    }
    if scanner.pos != update.len() {
        return Err(SplitError::Malformed);
    }
    if delete_clients != 0 || has_deleted_structs {
        validate_deletions(update, &clients)?;
    }
    let delete_set = &update[delete_start..];
    if let Some(last) = parts.last_mut()
        && last
            .len()
            .checked_add(delete_set.len() - 1)
            .is_some_and(|len| len <= max_part_bytes)
    {
        last.pop();
        last.extend_from_slice(delete_set);
    } else if delete_clients != 0 || parts.is_empty() {
        let mut last = Vec::with_capacity(1 + delete_set.len());
        last.push(0);
        last.extend_from_slice(delete_set);
        parts.push(last);
    }
    Ok(parts)
}

#[derive(Clone, Copy, Default)]
pub(crate) struct UpdateCursor {
    position: usize,
    phase: u8,
    clients: u32,
    client: Option<(u64, u32, u32)>,
    delete: Option<(u32, u32)>,
}

pub(crate) struct UpdatePart {
    pub(crate) bytes: Vec<u8>,
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) records: usize,
}

#[cfg(test)]
#[doc(hidden)]
pub(crate) fn deleted_struct_clocks_for_test(
    update: &[u8],
    deleted_client: u64,
    start: u32,
    end: u32,
) -> Result<usize, SplitError> {
    let mut scanner = Scanner::new(update);
    scanner.allow_maps = true;
    let mut deleted = 0;
    for _ in 0..scanner.count()? {
        let count = scanner.count()?;
        let client = scanner.client()?;
        let mut clock = scanner.clock()?;
        for _ in 0..count {
            let block = scanner.block()?;
            let stop = clock.checked_add(block.len).ok_or(SplitError::Malformed)?;
            if client == deleted_client && clock < end && stop > start {
                if !matches!(
                    block.kind,
                    BLOCK_GC_REF_NUMBER | BLOCK_ITEM_DELETED_REF_NUMBER
                ) {
                    return Err(SplitError::RetainedDeletion);
                }
                deleted += (stop.min(end) - clock.max(start)) as usize;
            }
            clock = stop;
        }
    }
    Ok(deleted)
}

impl UpdateCursor {
    pub(crate) fn next(
        &mut self,
        update: &[u8],
        max_records: usize,
        max_bytes: usize,
    ) -> Result<Option<UpdatePart>, SplitError> {
        self.next_inner(update, max_records, max_bytes, false, false)
    }

    pub(crate) fn next_admitted(
        &mut self,
        update: &[u8],
        max_records: usize,
        max_bytes: usize,
    ) -> Result<Option<UpdatePart>, SplitError> {
        self.next_inner(update, max_records, max_bytes, false, true)
    }

    fn next_oversized(
        &mut self,
        update: &[u8],
        max_records: usize,
        max_bytes: usize,
    ) -> Result<Option<UpdatePart>, SplitError> {
        self.next_inner(update, max_records, max_bytes, true, false)
    }

    fn next_inner(
        &mut self,
        update: &[u8],
        max_records: usize,
        max_bytes: usize,
        allow_oversized: bool,
        refuse_text: bool,
    ) -> Result<Option<UpdatePart>, SplitError> {
        if max_records == 0 || max_bytes == 0 {
            return Err(SplitError::InvalidLimit);
        }
        let mut next = *self;
        let mut scanner = Scanner::new(update);
        scanner.allow_maps = true;
        scanner.pos = next.position;
        if !allow_oversized {
            scanner.limit = scanner.pos.saturating_add(max_bytes).min(update.len());
            scanner.refuse_text = refuse_text;
            if refuse_text {
                scanner.allocation_left = max_bytes;
            }
            scanner.measured = refuse_text;
        }
        if next.phase == 0 {
            next.clients = scanner.count()?;
            next.phase = 1;
        }
        loop {
            if next.client.is_none() {
                if next.clients == 0 {
                    if next.phase == 1 {
                        next.clients = scanner.count()?;
                        next.phase = 2;
                        continue;
                    }
                    if scanner.pos != update.len() {
                        return Err(SplitError::Malformed);
                    }
                    next.position = scanner.pos;
                    *self = next;
                    return Ok(None);
                }
                next.client = Some(if next.phase == 1 {
                    let count = scanner.count()?;
                    let client = scanner.client()?;
                    (client, scanner.clock()?, count)
                } else {
                    let client = scanner.client()?;
                    (client, 0, scanner.count()?)
                });
                next.clients -= 1;
            }
            let (client, clock, remaining) = next.client.ok_or(SplitError::Malformed)?;
            if remaining == 0 {
                return Err(SplitError::Malformed);
            }
            if next.phase == 2 {
                let (start, end) = match next.delete.take() {
                    Some(range) => range,
                    None => scanner.delete_range()?,
                };
                let length = (end - start) as usize;
                let records = length.min(max_records);
                let mut bytes = vec![0, 1];
                write_var(&mut bytes, client);
                write_var(&mut bytes, 1);
                write_var(&mut bytes, u64::from(start));
                write_var(&mut bytes, records as u64);
                if bytes.len() > max_bytes && !allow_oversized {
                    return Err(SplitError::OversizedStruct);
                }
                next.position = scanner.pos;
                if records < length {
                    next.delete = Some((start + records as u32, end));
                } else {
                    next.client = (remaining > 1).then_some((client, 0, remaining - 1));
                }
                *self = next;
                return Ok(Some(UpdatePart { bytes, records }));
            }
            let start = scanner.pos;
            let mut end = start;
            let mut count = 0u32;
            let mut next_clock = clock;
            while count < remaining && (count as usize) < max_records {
                let length = if next.phase == 1 {
                    let overhead = run_size(client, clock, count + 1);
                    if !allow_oversized {
                        scanner.limit = start
                            .saturating_add(max_bytes.saturating_sub(overhead))
                            .min(update.len());
                    }
                    match scanner.block() {
                        Ok(block) => block.len,
                        Err(SplitError::OversizedStruct) if count != 0 => break,
                        Err(failure) => return Err(failure),
                    }
                } else {
                    scanner.delete_range()?;
                    0
                };
                let size = if next.phase == 1 {
                    run_size(client, clock, count + 1)
                } else {
                    2 + var_len(client) + var_len(u64::from(count + 1))
                };
                let oversized = size.saturating_add(scanner.pos - start) > max_bytes;
                if oversized && (count != 0 || !allow_oversized) {
                    if count == 0 {
                        return Err(SplitError::OversizedStruct);
                    }
                    break;
                }
                end = scanner.pos;
                count += 1;
                next_clock = next_clock
                    .checked_add(length)
                    .ok_or(SplitError::Malformed)?;
                if next_clock > MAX_CLOCK {
                    return Err(SplitError::Malformed);
                }
                if oversized {
                    break;
                }
            }
            let bytes = if next.phase == 1 {
                encode_run(client, clock, count, &update[start..end])
            } else {
                let mut bytes = vec![0, 1];
                write_var(&mut bytes, client);
                write_var(&mut bytes, u64::from(count));
                bytes.extend_from_slice(&update[start..end]);
                bytes
            };
            next.position = end;
            next.client = (count < remaining).then_some((client, next_clock, remaining - count));
            *self = next;
            return Ok(Some(UpdatePart {
                bytes,
                records: count as usize,
            }));
        }
    }

    pub(crate) fn is_complete(&mut self, update: &[u8]) -> Result<bool, SplitError> {
        if self.client.is_some() || self.delete.is_some() || self.clients != 0 || self.phase == 0 {
            return Ok(false);
        }
        let mut scanner = Scanner::new(update);
        scanner.pos = self.position;
        if self.phase == 1 {
            let count = scanner.count()?;
            if count != 0 {
                return Ok(false);
            }
        }
        if scanner.pos != update.len() {
            return Err(SplitError::Malformed);
        }
        self.phase = 2;
        self.position = scanner.pos;
        Ok(true)
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum SnapshotParent {
    Root(std::sync::Arc<str>),
    Nested(u64, u32),
}

impl SnapshotParent {
    pub(crate) fn branch(&self) -> yrs::BranchID {
        match self {
            Self::Root(name) => yrs::BranchID::Root(name.clone()),
            Self::Nested(client, clock) => {
                yrs::BranchID::Nested(yrs::ID::new(yrs::block::ClientID::new(*client), *clock))
            }
        }
    }
}

pub(crate) type SnapshotKeys = BTreeSet<(SnapshotParent, std::sync::Arc<str>)>;

type SnapshotLocations = BTreeMap<(u64, u32), (u32, SnapshotParent, Option<std::sync::Arc<str>>)>;

#[derive(Default)]
pub(crate) struct CausalState {
    clocks: BTreeMap<u64, u32>,
    kinds: BTreeMap<(u64, u32), (u32, u8)>,
    locations: SnapshotLocations,
    keys: SnapshotKeys,
    order_items: usize,
    order_clocks: usize,
}

#[derive(Default)]
struct AdmissionCheckpoint {
    clocks: BTreeMap<u64, Option<u32>>,
    keys: Vec<(SnapshotParent, std::sync::Arc<str>)>,
}

impl CausalState {
    pub(crate) fn drain(&mut self, max_records: usize, max_bytes: usize) -> (bool, usize, usize) {
        let mut records = 0;
        let mut bytes = 0;
        while records < max_records {
            let cost = if !self.locations.is_empty() {
                96
            } else if !self.kinds.is_empty() {
                std::mem::size_of::<((u64, u32), (u32, u8))>()
            } else if !self.clocks.is_empty() {
                std::mem::size_of::<(u64, u32)>()
            } else {
                break;
            };
            if cost > max_bytes.saturating_sub(bytes) {
                break;
            }
            if self.locations.pop_first().is_none() && self.kinds.pop_first().is_none() {
                self.clocks.pop_first();
            }
            records += 1;
            bytes += cost;
        }
        (self.is_empty(), records, bytes)
    }

    pub(crate) fn final_clock(&self) -> Option<(u64, u32)> {
        if self.kinds.is_empty() && self.locations.is_empty() {
            self.clocks
                .first_key_value()
                .map(|(&client, &clock)| (client, clock))
        } else {
            None
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.kinds.is_empty() && self.clocks.is_empty() && self.locations.is_empty()
    }

    pub(crate) fn take_keys(&mut self) -> SnapshotKeys {
        std::mem::take(&mut self.keys)
    }

    fn waiting_on(&self, update: &[u8]) -> Result<(u64, u32), SplitError> {
        let mut scanner = Scanner::new(update);
        scanner.allow_maps = true;
        if scanner.count()? != 1 || scanner.count()? != 1 {
            return Err(SplitError::Malformed);
        }
        let client = scanner.client()?;
        let clock = scanner.clock()?;
        let known = self.clocks.get(&client).copied().unwrap_or_default();
        if clock > known {
            return Ok((client, clock - 1));
        }
        if clock != known {
            return Err(SplitError::MissingDependency);
        }
        for (client, clock) in scanner.block()?.dependencies.into_iter().flatten() {
            if clock >= self.clocks.get(&client).copied().unwrap_or_default() {
                return Ok((client, clock));
            }
        }
        Err(SplitError::MissingDependency)
    }

    pub(crate) fn admit(&mut self, update: &[u8]) -> Result<(), SplitError> {
        #[cfg(test)]
        ADMISSION_ATTEMPTS.set(ADMISSION_ATTEMPTS.get() + 1);
        let mut checkpoint = AdmissionCheckpoint::default();
        let order_items = self.order_items;
        let order_clocks = self.order_clocks;
        let result = self.admit_inner(update, &mut checkpoint);
        if result.is_err() {
            self.order_items = order_items;
            self.order_clocks = order_clocks;
            for (client, clock) in checkpoint.clocks {
                let start = clock.unwrap_or_default();
                while let Some((&key, _)) = self
                    .locations
                    .range((client, start)..=(client, u32::MAX))
                    .next()
                {
                    self.locations.remove(&key);
                }
                while let Some((&key, _)) = self
                    .kinds
                    .range((client, start)..=(client, u32::MAX))
                    .next()
                {
                    self.kinds.remove(&key);
                }
                if let Some(clock) = clock {
                    self.clocks.insert(client, clock);
                } else {
                    self.clocks.remove(&client);
                }
            }
            for key in checkpoint.keys {
                self.keys.remove(&key);
            }
        }
        result
    }

    fn admit_inner(
        &mut self,
        update: &[u8],
        checkpoint: &mut AdmissionCheckpoint,
    ) -> Result<(), SplitError> {
        let mut scanner = Scanner::new(update);
        scanner.allow_maps = true;
        let clients = scanner.count()?;
        for _ in 0..clients {
            let count = scanner.count()?;
            let client = scanner.client()?;
            let mut clock = scanner.clock()?;
            if count == 0 || clock != self.clocks.get(&client).copied().unwrap_or_default() {
                return Err(SplitError::MissingDependency);
            }
            checkpoint
                .clocks
                .entry(client)
                .or_insert_with(|| self.clocks.get(&client).copied());
            for _ in 0..count {
                let block = scanner.block()?;
                if block.kind == BLOCK_SKIP_REF_NUMBER {
                    return Err(SplitError::MissingDependency);
                }
                for (index, dependency) in block.dependencies.into_iter().enumerate() {
                    let Some((dependency_client, dependency_clock)) = dependency else {
                        continue;
                    };
                    if dependency_clock
                        >= self
                            .clocks
                            .get(&dependency_client)
                            .copied()
                            .unwrap_or_default()
                    {
                        return Err(SplitError::MissingDependency);
                    }
                    let kind = self
                        .kinds
                        .range(..=(dependency_client, dependency_clock))
                        .next_back()
                        .filter(|((known_client, start), (len, _))| {
                            *known_client == dependency_client && dependency_clock - *start < *len
                        })
                        .map(|(_, &(_, kind))| kind);
                    if (index == 2 && kind != Some(BLOCK_ITEM_TYPE_REF_NUMBER))
                        || (index < 2 && kind == Some(BLOCK_GC_REF_NUMBER))
                    {
                        return Err(SplitError::MissingDependency);
                    }
                }
                let location = if let Some(root) = block.root {
                    Some((SnapshotParent::Root(root.into()), block.key.map(Into::into)))
                } else if let Some((parent, clock)) = block.dependencies[2] {
                    Some((
                        SnapshotParent::Nested(parent, clock),
                        block.key.map(Into::into),
                    ))
                } else {
                    block.dependencies[..2]
                        .iter()
                        .flatten()
                        .find_map(|&(parent, at)| {
                            self.locations
                                .range(..=(parent, at))
                                .next_back()
                                .filter(|((client, start), (len, _, _))| {
                                    *client == parent && at - *start < *len
                                })
                                .map(|(_, (_, parent, key))| (parent.clone(), key.clone()))
                        })
                };
                if let Some((parent, key)) = location {
                    if matches!(&parent, SnapshotParent::Root(name) if name.as_ref() == "xlsx:sheet-order")
                    {
                        if key.is_some()
                            || !matches!(
                                block.kind,
                                BLOCK_ITEM_ANY_REF_NUMBER
                                    | BLOCK_ITEM_DELETED_REF_NUMBER
                                    | BLOCK_GC_REF_NUMBER
                            )
                            || (block.kind == BLOCK_ITEM_ANY_REF_NUMBER && !block.sheet_keys)
                        {
                            return Err(SplitError::UnsupportedContent(block.kind));
                        }
                        self.order_items += usize::from(block.kind != BLOCK_GC_REF_NUMBER);
                        let clocks = if block.kind == BLOCK_ITEM_ANY_REF_NUMBER {
                            1
                        } else {
                            block.len as usize
                        };
                        self.order_clocks = self.order_clocks.saturating_add(clocks);
                        if self.order_items > SHEET_ORDER_MAX_ITEMS
                            || self.order_clocks > SHEET_ORDER_MAX_ITEMS
                        {
                            return Err(SplitError::SheetOrderLimit);
                        }
                    }
                    if let Some(key) = &key {
                        let entry = (parent.clone(), key.clone());
                        if self.keys.insert(entry.clone()) {
                            checkpoint.keys.push(entry);
                        }
                    }
                    self.locations
                        .insert((client, clock), (block.len, parent, key));
                }
                if matches!(block.kind, BLOCK_GC_REF_NUMBER | BLOCK_ITEM_TYPE_REF_NUMBER) {
                    self.kinds.insert((client, clock), (block.len, block.kind));
                }
                clock = clock
                    .checked_add(block.len)
                    .filter(|clock| *clock <= MAX_CLOCK)
                    .ok_or(SplitError::Malformed)?;
                self.clocks.insert(client, clock);
            }
        }
        let clients = scanner.count()?;
        for _ in 0..clients {
            let client = scanner.client()?;
            let count = scanner.count()?;
            if count == 0 {
                return Err(SplitError::Malformed);
            }
            for _ in 0..count {
                let (start, end) = scanner.delete_range()?;
                if end > self.clocks.get(&client).copied().unwrap_or_default() {
                    return Err(SplitError::MissingDependency);
                }
                if self
                    .kinds
                    .range((client, start)..(client, end))
                    .any(|(_, &(_, kind))| kind == BLOCK_ITEM_TYPE_REF_NUMBER)
                {
                    return Err(SplitError::RetainedDeletion);
                }
            }
        }
        if scanner.pos != update.len() {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }
}

#[cfg_attr(not(test), allow(dead_code))]
pub(crate) struct SplitParts {
    pub(crate) parts: Vec<Vec<u8>>,
    pub(crate) fallback: Option<SplitError>,
}

#[cfg_attr(not(test), allow(dead_code))]
pub(crate) fn split_or_whole_v1(update: &[u8], max_part_bytes: usize) -> SplitParts {
    match split_update_v1(update, max_part_bytes) {
        Ok(parts) => SplitParts {
            parts,
            fallback: None,
        },
        Err(error) => SplitParts {
            parts: vec![update.to_vec()],
            fallback: Some(error),
        },
    }
}

fn run_size(client: u64, clock: u32, count: u32) -> usize {
    2 + var_len(u64::from(count)) + var_len(client) + var_len(u64::from(clock))
}

fn encode_run(client: u64, clock: u32, count: u32, structs: &[u8]) -> Vec<u8> {
    let mut part = Vec::with_capacity(run_size(client, clock, count) + structs.len());
    part.push(1);
    write_var(&mut part, u64::from(count));
    write_var(&mut part, client);
    write_var(&mut part, u64::from(clock));
    part.extend_from_slice(structs);
    part.push(0);
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

fn validate_deletions(update: &[u8], clients: &HashMap<u64, Client>) -> Result<(), SplitError> {
    let mut scanner = Scanner::new(update);
    let count = scanner.count()?;
    for _ in 0..count {
        let count = scanner.count()?;
        let client = scanner.client()?;
        let mut clock = scanner.clock()?;
        let known = clients.get(&client).ok_or(SplitError::Malformed)?;
        let mut deletes = Scanner::new(update);
        deletes.pos = known.delete_offset.unwrap_or(update.len());
        let mut remaining = known.delete_count;
        let mut range = None;
        for _ in 0..count {
            let block = scanner.block()?;
            let end = clock + block.len;
            while range.is_none_or(|(_, end)| end <= clock) && remaining != 0 {
                range = Some(deletes.delete_range()?);
                remaining -= 1;
            }
            let deleted = matches!(
                block.kind,
                BLOCK_GC_REF_NUMBER | BLOCK_ITEM_DELETED_REF_NUMBER
            );
            let overlaps = range.is_some_and(|(start, stop)| start < end && stop > clock);
            let covered = range.is_some_and(|(start, stop)| start <= clock && stop >= end);
            if (deleted && !covered) || (!deleted && overlaps) {
                return Err(SplitError::RetainedDeletion);
            }
            clock = end;
        }
    }
    Ok(())
}

struct Scanner<'a> {
    bytes: &'a [u8],
    pos: usize,
    allow_maps: bool,
    limit: usize,
    refuse_text: bool,
    allocation_left: usize,
    measured: bool,
}

impl<'a> Scanner<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self {
            bytes,
            pos: 0,
            allow_maps: false,
            limit: bytes.len(),
            refuse_text: false,
            allocation_left: usize::MAX,
            measured: false,
        }
    }

    fn byte(&mut self) -> Result<u8, SplitError> {
        let byte = *self.bytes.get(self.pos).ok_or(SplitError::Malformed)?;
        if self.pos >= self.limit {
            return Err(SplitError::OversizedStruct);
        }
        self.pos += 1;
        #[cfg(test)]
        if self.measured {
            crate::snapshot::step::scan(1);
        }
        Ok(byte)
    }

    fn take(&mut self, len: usize) -> Result<&'a [u8], SplitError> {
        let end = self.pos.checked_add(len).ok_or(SplitError::Malformed)?;
        let bytes = self.bytes.get(self.pos..end).ok_or(SplitError::Malformed)?;
        if end > self.limit {
            return Err(SplitError::OversizedStruct);
        }
        self.pos = end;
        #[cfg(test)]
        if self.measured {
            crate::snapshot::step::scan(len);
        }
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
                if shift != 0 && payload == 0 {
                    return Err(SplitError::Malformed);
                }
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

    fn client(&mut self) -> Result<u64, SplitError> {
        let client = self.var()?;
        if client > MAX_CLIENT_ID {
            return Err(SplitError::Malformed);
        }
        Ok(client)
    }

    fn clock(&mut self) -> Result<u32, SplitError> {
        let clock = self.var_u32()?;
        if clock > MAX_CLOCK {
            return Err(SplitError::Malformed);
        }
        Ok(clock)
    }

    fn id(&mut self) -> Result<(u64, u32), SplitError> {
        Ok((self.client()?, self.clock()?))
    }

    fn buffer(&mut self) -> Result<&'a [u8], SplitError> {
        let len = self.var_u32()? as usize;
        self.take(len)
    }

    fn string(&mut self) -> Result<&'a str, SplitError> {
        std::str::from_utf8(self.buffer()?).map_err(|_| SplitError::Malformed)
    }

    fn signed_var(&mut self) -> Result<(), SplitError> {
        let first = self.byte()?;
        let mut value = u64::from(first & 0x3f);
        let mut byte = first;
        let mut shift = 6;
        while byte & 0x80 != 0 {
            byte = self.byte()?;
            let payload = u64::from(byte & 0x7f);
            if shift >= 63 || payload > (i64::MAX as u64 >> shift) {
                return Err(SplitError::Malformed);
            }
            value |= payload << shift;
            if byte & 0x80 == 0 && payload == 0 {
                return Err(SplitError::Malformed);
            }
            shift += 7;
        }
        if first & 0x40 != 0 && value == 0 {
            return Err(SplitError::Malformed);
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
            tag @ (118 | 117) => {
                let count = usize::try_from(self.var()?).map_err(|_| SplitError::Malformed)?;
                if tag == 118 && count > 1 && !self.allow_maps {
                    return Err(SplitError::UnsupportedMap);
                }
                if count > self.bytes.len() - self.pos {
                    return Err(SplitError::Malformed);
                }
                if tag == 118 {
                    let allocation = count.checked_mul(256).ok_or(SplitError::OversizedStruct)?;
                    if allocation > self.allocation_left {
                        return Err(SplitError::OversizedStruct);
                    }
                    self.allocation_left -= allocation;
                }
                let mut keys = BTreeSet::new();
                for _ in 0..count {
                    if tag == 118 {
                        let key = self.string()?;
                        #[cfg(test)]
                        if self.measured {
                            crate::snapshot::step::allocate(256);
                        }
                        if !keys.insert(key) {
                            return Err(SplitError::Malformed);
                        }
                    }
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

    fn block(&mut self) -> Result<Struct<'a>, SplitError> {
        let info = self.byte()?;
        let kind = info & 0x0f;
        let mut dependencies = [None; 3];
        let mut root = None;
        let mut key = None;
        let mut sheet_keys = true;
        let len = if info == BLOCK_GC_REF_NUMBER || info == BLOCK_SKIP_REF_NUMBER {
            self.var_u32()?
        } else {
            if info & 0x10 != 0 {
                return Err(SplitError::UnsupportedContent(kind));
            }
            if info & HAS_ORIGIN != 0 {
                dependencies[0] = Some(self.id()?);
            }
            if info & HAS_RIGHT_ORIGIN != 0 {
                dependencies[1] = Some(self.id()?);
            }
            if info & (HAS_ORIGIN | HAS_RIGHT_ORIGIN) == 0 {
                match self.var_u32()? {
                    0 => dependencies[2] = Some(self.id()?),
                    1 => {
                        root = Some(self.string()?);
                    }
                    _ => return Err(SplitError::Malformed),
                }
                if info & HAS_PARENT_SUB != 0 {
                    key = Some(self.string()?);
                }
            }
            match kind {
                BLOCK_ITEM_DELETED_REF_NUMBER => self.var_u32()?,
                BLOCK_ITEM_JSON_REF_NUMBER => {
                    let count = self.count()?;
                    for _ in 0..count {
                        self.string()?;
                    }
                    return Err(SplitError::JsonLengthMismatch);
                }
                BLOCK_ITEM_BINARY_REF_NUMBER => {
                    self.buffer()?;
                    1
                }
                BLOCK_ITEM_STRING_REF_NUMBER => {
                    if self.refuse_text {
                        return Err(SplitError::SharedText);
                    }
                    u32::try_from(self.string()?.encode_utf16().count())
                        .map_err(|_| SplitError::Malformed)?
                }
                BLOCK_ITEM_EMBED_REF_NUMBER => {
                    if self.refuse_text {
                        return Err(SplitError::SharedText);
                    }
                    self.json()?;
                    1
                }
                BLOCK_ITEM_FORMAT_REF_NUMBER => {
                    if self.refuse_text {
                        return Err(SplitError::SharedText);
                    }
                    self.string()?;
                    self.json()?;
                    1
                }
                BLOCK_ITEM_TYPE_REF_NUMBER => {
                    let type_ref = self.byte()?;
                    if self.refuse_text && matches!(type_ref, TYPE_REFS_TEXT | TYPE_REFS_XML_TEXT) {
                        return Err(SplitError::SharedText);
                    }
                    match type_ref {
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
                        kind => return Err(SplitError::UnsupportedType(kind)),
                    }
                    1
                }
                BLOCK_ITEM_ANY_REF_NUMBER => {
                    let len = self.count()?;
                    for _ in 0..len {
                        if self.bytes.get(self.pos) == Some(&119) {
                            self.byte()?;
                            sheet_keys &= self.string()?.len() <= 64;
                        } else {
                            sheet_keys = false;
                            self.any(0)?;
                        }
                    }
                    len
                }
                BLOCK_ITEM_DOC_REF_NUMBER => {
                    self.string()?;
                    self.any(0)?;
                    1
                }
                _ => return Err(SplitError::UnsupportedContent(kind)),
            }
        };
        if len == 0 {
            return Err(SplitError::Malformed);
        }
        Ok(Struct {
            len,
            kind,
            dependencies,
            root,
            key,
            sheet_keys,
        })
    }

    fn delete_range(&mut self) -> Result<(u32, u32), SplitError> {
        let clock = self.clock()?;
        let len = self.var_u32()?;
        let end = clock.checked_add(len).ok_or(SplitError::Malformed)?;
        if len == 0 || end > MAX_CLOCK {
            return Err(SplitError::Malformed);
        }
        Ok((clock, end))
    }

    fn json(&mut self) -> Result<(), SplitError> {
        let mut scanner = Scanner::new(self.string()?.as_bytes());
        scanner.allow_maps = self.allow_maps;
        let mut json = Json { scanner };
        json.value(0)?;
        json.space();
        if json.scanner.pos != json.scanner.bytes.len() {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }
}

struct Json<'a> {
    scanner: Scanner<'a>,
}

impl Json<'_> {
    fn peek(&self) -> Option<u8> {
        self.scanner.bytes.get(self.scanner.pos).copied()
    }

    fn consume(&mut self, byte: u8) -> bool {
        if self.peek() == Some(byte) {
            self.scanner.pos += 1;
            true
        } else {
            false
        }
    }

    fn space(&mut self) {
        while matches!(self.peek(), Some(b' ' | b'\r' | b'\n' | b'\t')) {
            self.scanner.pos += 1;
        }
    }

    fn quoted(&mut self) -> Result<(), SplitError> {
        if !self.consume(b'"') {
            return Err(SplitError::Malformed);
        }
        loop {
            match self.scanner.byte()? {
                b'"' => return Ok(()),
                b'\\' => match self.scanner.byte()? {
                    b'"' | b'\\' | b'/' | b'b' | b'f' | b'n' | b'r' | b't' => {}
                    b'u' => {
                        let code = self.hex()?;
                        if (0xd800..=0xdbff).contains(&code) {
                            if self.scanner.take(2)? != b"\\u"
                                || !(0xdc00..=0xdfff).contains(&self.hex()?)
                            {
                                return Err(SplitError::Malformed);
                            }
                        } else if (0xdc00..=0xdfff).contains(&code) {
                            return Err(SplitError::Malformed);
                        }
                    }
                    _ => return Err(SplitError::Malformed),
                },
                0..=31 => return Err(SplitError::Malformed),
                _ => {}
            }
        }
    }

    fn hex(&mut self) -> Result<u16, SplitError> {
        let bytes = self.scanner.take(4)?;
        if !bytes.iter().all(u8::is_ascii_hexdigit) {
            return Err(SplitError::Malformed);
        }
        let text = std::str::from_utf8(bytes).map_err(|_| SplitError::Malformed)?;
        u16::from_str_radix(text, 16).map_err(|_| SplitError::Malformed)
    }

    fn digits(&mut self) -> Result<(), SplitError> {
        let start = self.scanner.pos;
        while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
            self.scanner.pos += 1;
        }
        if self.scanner.pos == start {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }

    fn number(&mut self) -> Result<(), SplitError> {
        let start = self.scanner.pos;
        self.consume(b'-');
        if !self.consume(b'0') {
            self.digits()?;
        }
        if self.consume(b'.') {
            self.digits()?;
        }
        if self.consume(b'e') || self.consume(b'E') {
            if !self.consume(b'+') {
                self.consume(b'-');
            }
            self.digits()?;
        }
        let text = std::str::from_utf8(&self.scanner.bytes[start..self.scanner.pos])
            .map_err(|_| SplitError::Malformed)?;
        let value = text.parse::<f64>().map_err(|_| SplitError::Malformed)?;
        if !value.is_finite() {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }

    fn value(&mut self, depth: u8) -> Result<(), SplitError> {
        if depth >= MAX_NESTING {
            return Err(SplitError::Malformed);
        }
        self.space();
        match self.peek() {
            Some(b'"') => self.quoted()?,
            Some(b'-' | b'0'..=b'9') => self.number()?,
            Some(b'n' | b't' | b'f') => {
                let expected: &[u8] = match self.peek() {
                    Some(b'n') => b"null",
                    Some(b't') => b"true",
                    _ => b"false",
                };
                if self.scanner.take(expected.len())? != expected {
                    return Err(SplitError::Malformed);
                }
            }
            Some(open @ (b'[' | b'{')) => {
                self.scanner.pos += 1;
                let close = if open == b'[' { b']' } else { b'}' };
                self.space();
                if self.consume(close) {
                    return Ok(());
                }
                loop {
                    if open == b'{' {
                        self.quoted()?;
                        self.space();
                        if !self.consume(b':') {
                            return Err(SplitError::Malformed);
                        }
                    }
                    self.value(depth + 1)?;
                    self.space();
                    if self.consume(close) {
                        break;
                    }
                    if !self.consume(b',') {
                        return Err(SplitError::Malformed);
                    }
                    if open == b'{' && !self.scanner.allow_maps {
                        return Err(SplitError::UnsupportedMap);
                    }
                    self.space();
                }
            }
            _ => return Err(SplitError::Malformed),
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::authority::hydrate_snapshot_part;
    use crate::{
        CalculationOptions, Cell, CellRange, CellRef, CellValue, DefinedName, Sheet, SheetId,
        Workbook, WorkbookModel,
    };
    use xlsx_model::{CellFormat, NumberFormat};
    use yrs::encoding::write::Write;
    use yrs::types::{AsPrelim, ToJson};
    use yrs::updates::decoder::Decode;
    use yrs::updates::encoder::{Encoder, EncoderV1};
    use yrs::{
        Any, Array, Doc, GetString, In, Map, MapPrelim, Out, ReadTxn, StateVector, Transact, Update,
    };

    const LIMITS: [usize; 5] = [1, 64, 4096, 65536, usize::MAX];

    fn ordinary_doc() -> Doc {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        let array = doc.get_or_insert_array("array");
        {
            let mut txn = doc.transact_mut();
            for index in 0..50 {
                map.insert(&mut txn, format!("key{index}"), format!("value{index}"));
            }
            array.insert_range(&mut txn, 0, 0..50);
        }
        doc
    }

    fn canonical_any(value: &Any, output: &mut String) {
        match value {
            Any::Map(entries) => {
                let mut entries = entries.iter().collect::<Vec<_>>();
                entries.sort_unstable_by_key(|(key, _)| *key);
                output.push('{');
                for (index, (key, value)) in entries.into_iter().enumerate() {
                    if index != 0 {
                        output.push(',');
                    }
                    output.push_str(&format!("{key:?}:"));
                    canonical_any(value, output);
                }
                output.push('}');
            }
            Any::Array(values) => {
                output.push('[');
                for (index, value) in values.iter().enumerate() {
                    if index != 0 {
                        output.push(',');
                    }
                    canonical_any(value, output);
                }
                output.push(']');
            }
            _ => output.push_str(&format!("{value:?}")),
        }
    }

    fn canonical_snapshot(doc: &Doc) -> String {
        let txn = doc.transact();
        let mut roots = txn.root_refs().collect::<Vec<_>>();
        roots.sort_unstable_by_key(|(name, _)| *name);
        let mut output = String::from("{");
        for (index, (name, root)) in roots.into_iter().enumerate() {
            if index != 0 {
                output.push(',');
            }
            output.push_str(&format!("{name:?}:"));
            let root = match root {
                Out::UndefinedRef(branch) => {
                    let text = yrs::TextRef::from(branch);
                    if txn.get_map(name).unwrap().len(&txn) == 0
                        && !text.get_string(&txn).is_empty()
                    {
                        Out::YText(text)
                    } else {
                        match root.as_prelim(&txn) {
                            In::Map(_) => Out::YMap(branch.into()),
                            In::Array(_) => Out::YArray(branch.into()),
                            In::Text(_) => Out::YText(branch.into()),
                            In::XmlElement(_) => Out::YXmlElement(branch.into()),
                            In::XmlText(_) => Out::YXmlText(branch.into()),
                            _ => unreachable!("unexpected inferred root type"),
                        }
                    }
                }
                root => root,
            };
            canonical_any(&root.to_json(&txn), &mut output);
        }
        output.push('}');
        output
    }

    fn small_model() -> WorkbookModel {
        let mut sheet = Sheet::new("Small");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Text {
                    value: "hello".into(),
                },
                ..Cell::default()
            },
        );
        WorkbookModel {
            sheets: vec![sheet],
            ..WorkbookModel::default()
        }
    }

    fn rich_model() -> WorkbookModel {
        let mut model = WorkbookModel {
            shared_strings: vec!["shared 🦀".into(), "second".into()],
            ..WorkbookModel::default()
        };
        let mut format = CellFormat::default();
        format.font.bold = true;
        format.alignment.wrap_text = true;
        format.number_format = NumberFormat::Custom {
            pattern: "0.000".into(),
        };
        let style = model.styles.intern_cell_format(&format).unwrap();
        for index in 0..3 {
            let mut sheet = Sheet::new(format!("Data{index}"));
            for (col, value) in [
                CellValue::Text {
                    value: model.shared_strings[0].clone(),
                },
                CellValue::Number {
                    value: 12.25 + f64::from(index),
                },
                CellValue::Bool {
                    value: index % 2 == 0,
                },
                CellValue::Number {
                    value: 24.5 + f64::from(index) * 2.0,
                },
            ]
            .into_iter()
            .enumerate()
            {
                sheet.set_cell(
                    CellRef::new(0, col as u32),
                    Cell {
                        value,
                        formula: (col == 3).then(|| "B1*2".into()),
                        style,
                    },
                );
            }
            sheet
                .merges
                .push(CellRange::new(CellRef::new(2, 0), CellRef::new(2, 2)));
            sheet.col_widths.insert(0, 24.0);
            sheet.row_heights.insert(0, 28.0);
            model.sheets.push(sheet);
        }
        model.defined_names.push(DefinedName {
            name: "Amount".into(),
            formula: "Data0!$B$1".into(),
            local_sheet: None,
            hidden: false,
        });
        model.defined_names.push(DefinedName {
            name: "LocalAmount".into(),
            formula: "Data1!$D$1".into(),
            local_sheet: Some(SheetId(1)),
            hidden: true,
        });
        model
    }

    fn append_structs(
        update: &[u8],
        output: &mut Vec<u8>,
        clocks: &mut HashMap<u64, u32>,
    ) -> (usize, usize) {
        let mut scanner = Scanner::new(update);
        let clients = scanner.count().unwrap();
        let mut structs = 0;
        for _ in 0..clients {
            let count = scanner.count().unwrap();
            let client = scanner.client().unwrap();
            let mut clock = scanner.clock().unwrap();
            assert_eq!(clock, clocks.get(&client).copied().unwrap_or_default());
            for _ in 0..count {
                let start = scanner.pos;
                let block = scanner.block().unwrap();
                output.extend_from_slice(&update[start..scanner.pos]);
                clock += block.len;
                structs += 1;
            }
            clocks.insert(client, clock);
        }
        (scanner.pos, structs)
    }

    fn vector_bytes(vector: &StateVector) -> Vec<u8> {
        let mut entries = vector
            .iter()
            .map(|(client, clock)| (client.get(), *clock))
            .collect::<Vec<_>>();
        entries.sort_unstable();
        let mut encoder = EncoderV1::new();
        encoder.write_var(entries.len());
        for (client, clock) in entries {
            encoder.write_var(client);
            encoder.write_var(clock);
        }
        encoder.to_vec()
    }

    fn assert_parts(update: &[u8], vector: &[u8], client_id: u64, limit: usize, parts: &[Vec<u8>]) {
        assert!(!parts.is_empty());
        let mut original_structs = Vec::new();
        let (delete_start, _) = append_structs(update, &mut original_structs, &mut HashMap::new());
        let mut split_structs = Vec::with_capacity(original_structs.len());
        let mut clocks = HashMap::new();
        for (index, part) in parts.iter().enumerate() {
            let (delete_offset, count) = append_structs(part, &mut split_structs, &mut clocks);
            assert!(part.len() <= limit || count <= 1);
            if index + 1 == parts.len() {
                assert_eq!(&part[delete_offset..], &update[delete_start..]);
            } else {
                assert_eq!(&part[delete_offset..], &[0]);
            }
        }
        assert_eq!(split_structs, original_structs);
        assert_applied_parts(update, vector, client_id, parts, |doc, part| {
            hydrate_snapshot_part(doc, part).unwrap();
        });
    }

    fn apply_raw_part(doc: &Doc, part: &[u8]) {
        doc.transact_mut()
            .apply_update(Update::decode_v1(part).unwrap())
            .unwrap();
    }

    fn assert_applied_parts(
        update: &[u8],
        vector: &[u8],
        client_id: u64,
        parts: &[Vec<u8>],
        apply_part: fn(&Doc, &[u8]),
    ) {
        let doc = Doc::with_client_id(client_id);
        for (index, part) in parts.iter().enumerate() {
            apply_part(&doc, part);
            let txn = doc.transact();
            assert!(
                txn.store().pending_update().is_none(),
                "part {index} has missing structs"
            );
            assert!(
                txn.store().pending_ds().is_none(),
                "part {index} has missing deletes"
            );
        }
        let txn = doc.transact();
        assert_eq!(txn.state_vector(), StateVector::decode_v1(vector).unwrap());
        assert_eq!(vector_bytes(&txn.state_vector()), vector);
        assert_eq!(
            txn.encode_state_as_update_v1(&StateVector::default()),
            update
        );
        drop(txn);
        let source = Doc::with_client_id(client_id);
        apply_part(&source, update);
        assert_eq!(canonical_snapshot(&doc), canonical_snapshot(&source));
    }

    fn assert_workbook(workbook: &Workbook, allow_refusal: bool) {
        let update = workbook.encode_state_as_update_v1();
        let vector = workbook.encode_state_vector_v1();
        for limit in LIMITS {
            match split_update_v1(&update, limit) {
                Ok(parts) => assert_parts(&update, &vector, workbook.client_id(), limit, &parts),
                Err(error) => {
                    assert!(allow_refusal, "split at {limit} failed: {error:?}");
                    assert!(matches!(
                        error,
                        SplitError::MissingDependency | SplitError::RetainedDeletion
                    ));
                }
            }
        }
    }

    fn assert_doc(doc: &Doc) {
        let txn = doc.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = vector_bytes(&txn.state_vector());
        drop(txn);
        for limit in LIMITS {
            let parts = split_update_v1(&update, limit).unwrap();
            assert_parts(&update, &vector, doc.client_id().get(), limit, &parts);
        }
    }

    #[test]
    fn small_sheet_snapshot_is_exact() {
        let workbook = Workbook::from_model(small_model()).unwrap();
        assert_workbook(&workbook, false);
    }

    #[test]
    fn rich_workbook_snapshot_is_exact() {
        let saved = Workbook::from_model(rich_model()).unwrap().save().unwrap();
        let workbook = Workbook::open(&saved).unwrap();
        assert_eq!(workbook.sheet_count(), 3);
        assert!(!workbook.model().shared_strings.is_empty());
        assert_eq!(workbook.model().defined_names.len(), 2);
        assert_workbook(&workbook, false);
    }

    #[test]
    fn twenty_thousand_cell_snapshot_is_exact() {
        let mut sheet = Sheet::new("Large");
        for row in 0..1000 {
            for col in 0..20 {
                sheet.set_cell(
                    CellRef::new(row, col),
                    Cell {
                        value: CellValue::Number {
                            value: f64::from(row * 20 + col),
                        },
                        ..Cell::default()
                    },
                );
            }
        }
        assert_eq!(sheet.iter_cells().count(), 20_000);
        let workbook = Workbook::from_model(WorkbookModel {
            sheets: vec![sheet],
            ..WorkbookModel::default()
        })
        .unwrap();
        assert_workbook(&workbook, false);
    }

    #[test]
    fn edited_and_undone_snapshot_is_exact_or_refused() {
        let mut workbook = Workbook::from_model_collaborative(small_model(), 1).unwrap();
        let options = CalculationOptions::default();
        for (col, input) in [(0, "changed"), (1, "42"), (0, "again")] {
            workbook
                .edit_cell(SheetId(0), CellRef::new(0, col), input, options)
                .unwrap();
        }
        assert!(workbook.can_undo());
        workbook.undo(options).unwrap();
        let update = workbook.encode_state_as_update_v1();
        let (delete_start, _) = append_structs(&update, &mut Vec::new(), &mut HashMap::new());
        assert!(Scanner::new(&update[delete_start..]).count().unwrap() > 0);
        assert_eq!(
            StateVector::decode_v1(&workbook.encode_state_vector_v1())
                .unwrap()
                .len(),
            2
        );
        assert_workbook(&workbook, true);
    }

    #[test]
    fn deleted_structs_keep_the_original_delete_set() {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        {
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "first", "old");
            let child = map.insert(&mut txn, "child", MapPrelim::default());
            child.insert(&mut txn, "nested", "garbage collect");
            map.insert(&mut txn, "last", "keep");
        }
        {
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "first", "new");
            map.remove(&mut txn, "child");
        }
        assert_doc(&doc);
    }

    #[test]
    fn failed_admission_restores_existing_clocks_kinds_locations_and_keys() {
        use yrs::WriteTxn;

        let doc = Doc::with_client_id(7);
        let root = doc.get_or_insert_map("journal");
        root.insert(&mut doc.transact_mut(), "existing", 0_i64);
        root.insert(&mut doc.transact_mut(), "nested", MapPrelim::default());
        let baseline = doc.transact().state_vector();
        let mut causal = CausalState::default();
        causal
            .admit(
                &doc.transact()
                    .encode_state_as_update_v1(&StateVector::default()),
            )
            .unwrap();
        let clocks = causal.clocks.clone();
        let kinds = causal.kinds.clone();
        let locations = causal.locations.clone();
        let keys = causal.keys.clone();
        {
            let mut txn = doc.transact_mut();
            root.insert(&mut txn, "existing", 1_i64);
            root.insert(&mut txn, "new", MapPrelim::default());
            txn.get_or_insert_array("xlsx:sheet-order")
                .insert(&mut txn, 0, "sheet:0");
        }
        let delta = doc.transact().encode_diff_v1(&baseline);
        let mut malformed = delta.clone();
        malformed.push(0);
        assert_eq!(causal.admit(&malformed), Err(SplitError::Malformed));
        assert_eq!(causal.clocks, clocks);
        assert_eq!(causal.kinds, kinds);
        assert_eq!(causal.locations, locations);
        assert_eq!(causal.keys, keys);
        assert_eq!(causal.order_items, 0);
        assert_eq!(causal.order_clocks, 0);
        causal.admit(&delta).unwrap();
        assert_eq!(causal.order_items, 1);
        assert_eq!(causal.order_clocks, 1);
        assert!(
            causal
                .keys
                .contains(&(SnapshotParent::Root("journal".into()), "new".into()))
        );
    }

    #[test]
    fn fallback_orders_cross_client_dependencies_before_emission() {
        let first = Doc::with_client_id(7);
        let map = first.get_or_insert_map("map");
        map.insert(&mut first.transact_mut(), "nested", MapPrelim::default());
        let second = Doc::with_client_id(99);
        apply_raw_part(
            &second,
            &first
                .transact()
                .encode_state_as_update_v1(&StateVector::default()),
        );
        let nested = {
            let txn = second.transact();
            let Some(Out::YMap(nested)) = txn.get_map("map").unwrap().get(&txn, "nested") else {
                panic!("missing nested map");
            };
            nested
        };
        nested.insert(&mut second.transact_mut(), "value", 42);
        let txn = second.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = vector_bytes(&txn.state_vector());
        drop(txn);
        assert_eq!(
            split_update_v1_bounded(&update, 1, 384),
            Err(SplitError::MissingDependency),
        );
        let parts = split_fallback_v1_bounded(&update, 1, 384).unwrap();
        let mut causal = CausalState::default();
        for part in &parts {
            causal.admit(part).unwrap();
        }
        assert_applied_parts(&update, &vector, 99, &parts, |doc, part| {
            hydrate_snapshot_part(doc, part).unwrap();
        });
    }

    #[test]
    fn fallback_nested_client_admission_attempts_are_linear() {
        const CLIENTS: u64 = 2_000;
        let mut encoder = EncoderV1::new();
        encoder.write_var(CLIENTS);
        for client in (1..=CLIENTS).rev() {
            encoder.write_var(if client == CLIENTS { 2_u32 } else { 1_u32 });
            encoder.write_var(client);
            encoder.write_var(0_u32);
            encoder.write_info(BLOCK_ITEM_TYPE_REF_NUMBER | HAS_PARENT_SUB);
            encoder.write_parent_info(client == CLIENTS);
            if client == CLIENTS {
                encoder.write_string("map");
            } else {
                encoder.write_var(client + 1);
                encoder.write_var(0_u32);
            }
            encoder.write_string("nested");
            encoder.write_type_ref(TYPE_REFS_MAP);
            if client == CLIENTS {
                encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | HAS_PARENT_SUB);
                encoder.write_parent_info(false);
                encoder.write_var(1_u64);
                encoder.write_var(0_u32);
                encoder.write_string("value");
                encoder.write_len(1);
                encoder.write_any(&Any::from(42));
            }
        }
        encoder.write_var(0_u32);
        let source = Doc::with_client_id(CLIENTS);
        source.get_or_insert_map("map");
        apply_raw_part(&source, &encoder.to_vec());
        let txn = source.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = txn.state_vector();
        drop(txn);
        assert_eq!(
            split_update_v1_bounded(&update, 7, 16_384),
            Err(SplitError::MissingDependency),
        );
        ADMISSION_ATTEMPTS.set(0);
        let parts = split_fallback_v1_bounded(&update, 7, 16_384).unwrap();
        let attempts = ADMISSION_ATTEMPTS.get();
        assert!(attempts > CLIENTS as usize);
        assert!(attempts <= 4 * CLIENTS as usize + 4, "{attempts}");
        assert_eq!(parts.len(), CLIENTS as usize + 1);
        let peer = Doc::with_client_id(CLIENTS);
        peer.get_or_insert_map("map");
        for (index, part) in parts.iter().enumerate() {
            let mut scanner = Scanner::new(part);
            assert_eq!(scanner.count().unwrap(), 1);
            assert_eq!(scanner.count().unwrap(), 1);
            let expected = if index < CLIENTS as usize {
                CLIENTS - index as u64
            } else {
                CLIENTS
            };
            assert_eq!(scanner.client().unwrap(), expected);
            assert_eq!(
                scanner.clock().unwrap(),
                u32::from(index == CLIENTS as usize),
            );
            hydrate_snapshot_part(&peer, part).unwrap();
            let txn = peer.transact();
            assert!(txn.store().pending_update().is_none());
            assert!(txn.store().pending_ds().is_none());
        }
        assert_eq!(peer.transact().state_vector(), vector);
        assert_eq!(
            peer.transact()
                .encode_state_as_update_v1(&StateVector::default()),
            update,
        );
    }

    #[test]
    fn fallback_splits_oversized_deletions_with_multi_key_maps() {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        map.insert(
            &mut doc.transact_mut(),
            "value",
            Any::Map(std::sync::Arc::new(HashMap::from([
                ("first".into(), Any::Bool(true)),
                ("second".into(), Any::Bool(false)),
            ]))),
        );
        let array = doc.get_or_insert_array("deleted");
        array.insert_range(&mut doc.transact_mut(), 0, 0..100);
        array.remove_range(&mut doc.transact_mut(), 0, 100);
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        assert_eq!(
            split_update_v1_bounded(&update, 1, 384),
            Err(SplitError::UnsupportedMap),
        );
        let parts = split_fallback_v1_bounded(&update, 1, 384).unwrap();
        let peer = Doc::with_client_id(7);
        let mut causal = CausalState::default();
        let mut deleted = 0;
        for part in &parts {
            let mut cursor = UpdateCursor::default();
            let admitted = cursor.next(part, 1, 384).unwrap().unwrap();
            assert_eq!(admitted.records, 1);
            deleted += usize::from(admitted.bytes[0] == 0);
            causal.admit(&admitted.bytes).unwrap();
            hydrate_snapshot_part(&peer, &admitted.bytes).unwrap();
            assert!(cursor.is_complete(part).unwrap());
            let txn = peer.transact();
            assert!(txn.store().pending_update().is_none());
            assert!(txn.store().pending_ds().is_none());
        }
        assert_eq!(deleted, 100);
        assert_eq!(
            peer.transact().state_vector(),
            doc.transact().state_vector()
        );
        assert_eq!(canonical_snapshot(&peer), canonical_snapshot(&doc));
    }

    #[test]
    fn bounded_delete_ranges_charge_deleted_clocks_and_preserve_update() {
        let doc = Doc::with_client_id(7);
        let array = doc.get_or_insert_array("deleted");
        array.insert_range(&mut doc.transact_mut(), 0, 0..20_000);
        array.remove_range(&mut doc.transact_mut(), 0, 20_000);
        let txn = doc.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = vector_bytes(&txn.state_vector());
        drop(txn);
        for records in [1, 7, 256] {
            let parts = split_update_v1_bounded(&update, records, 384).unwrap();
            let mut deleted = 0;
            let mut causal = CausalState::default();
            for part in &parts {
                let mut cursor = UpdateCursor::default();
                let mut work = 0;
                while let Some(admitted) = cursor.next(part, records, 384).unwrap() {
                    causal.admit(&admitted.bytes).unwrap();
                    if admitted.bytes[0] == 0 {
                        deleted += admitted.records;
                    }
                    work += admitted.records;
                }
                assert!(work <= records);
                assert!(part.len() <= 384);
            }
            assert_eq!(deleted, 20_000);
            assert_applied_parts(&update, &vector, 7, &parts, |doc, part| {
                hydrate_snapshot_part(doc, part).unwrap();
            });
        }
    }

    #[test]
    fn bounded_subupdates_preserve_structs_deletions_and_full_encode() {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        {
            let mut txn = doc.transact_mut();
            for index in 0..500 {
                map.insert(&mut txn, format!("key{index}"), index);
            }
        }
        {
            let mut txn = doc.transact_mut();
            for index in (0..500).step_by(2) {
                map.remove(&mut txn, &format!("key{index}"));
            }
        }
        let txn = doc.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = vector_bytes(&txn.state_vector());
        drop(txn);
        for records in [1, 7, 256] {
            let parts = split_update_v1_bounded(&update, records, 384).unwrap();
            assert!(parts.len() > 1);
            for part in &parts {
                assert!(part.len() <= 384);
                let mut cursor = UpdateCursor::default();
                let mut processed = 0;
                while let Some(part) = cursor.next(part, records, 384).unwrap() {
                    processed += part.records;
                }
                assert!(processed <= records);
            }
            assert_applied_parts(&update, &vector, 7, &parts, |doc, part| {
                hydrate_snapshot_part(doc, part).unwrap();
            });
        }
    }

    #[test]
    fn subupdate_cursor_rejects_truncated_input_without_panicking() {
        let workbook = Workbook::from_model(small_model()).unwrap();
        let update = workbook.encode_state_as_update_v1();
        for end in 0..update.len() {
            let mut cursor = UpdateCursor::default();
            loop {
                match cursor.next(&update[..end], 7, 512) {
                    Ok(Some(_)) => {}
                    Err(_) => break,
                    Ok(None) => panic!("accepted truncated Yrs update at {end}"),
                }
            }
        }
    }

    #[test]
    fn ordered_client_dependencies_integrate_immediately() {
        let source = Doc::with_client_id(9);
        let map = source.get_or_insert_map("map");
        map.insert(&mut source.transact_mut(), "key", "first");
        let update = source
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let peer = Doc::with_client_id(1);
        hydrate_snapshot_part(&peer, &update).unwrap();
        let map = peer.get_or_insert_map("map");
        map.insert(&mut peer.transact_mut(), "key", "second");
        assert_doc(&peer);
    }

    #[test]
    fn forward_client_dependencies_are_refused() {
        let source = Doc::with_client_id(1);
        let map = source.get_or_insert_map("map");
        map.insert(&mut source.transact_mut(), "key", "first");
        let update = source
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let peer = Doc::with_client_id(9);
        hydrate_snapshot_part(&peer, &update).unwrap();
        let map = peer.get_or_insert_map("map");
        map.insert(&mut peer.transact_mut(), "key", "second");
        let update = peer
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        for limit in LIMITS {
            assert_eq!(
                split_update_v1(&update, limit),
                Err(SplitError::MissingDependency)
            );
        }
    }

    fn raw_update(count: u32, content: impl FnOnce(&mut EncoderV1)) -> Vec<u8> {
        let mut encoder = EncoderV1::new();
        encoder.write_var(1_u32);
        encoder.write_var(count);
        encoder.write_var(7_u64);
        encoder.write_var(0_u32);
        content(&mut encoder);
        encoder.write_var(0_u32);
        encoder.to_vec()
    }

    fn item(encoder: &mut EncoderV1, kind: u8) {
        encoder.write_info(kind);
        encoder.write_parent_info(true);
        encoder.write_string("root");
    }

    fn assert_boundaries(update: &[u8], clock: u32) {
        let source = Doc::with_client_id(7);
        apply_raw_part(&source, update);
        let txn = source.transact();
        assert_eq!(txn.state_vector().get(&source.client_id()), clock);
        let vector = vector_bytes(&txn.state_vector());
        let expected = txn.encode_state_as_update_v1(&StateVector::default());
        drop(txn);
        for limit in LIMITS {
            let parts = split_update_v1(update, limit).unwrap();
            let mut original = Vec::new();
            append_structs(update, &mut original, &mut HashMap::new());
            let mut actual = Vec::new();
            let mut clocks = HashMap::new();
            for part in &parts {
                append_structs(part, &mut actual, &mut clocks);
            }
            assert_eq!(actual, original);
            assert_eq!(clocks[&7], clock);
            assert_applied_parts(
                &expected,
                &vector,
                source.client_id().get(),
                &parts,
                apply_raw_part,
            );
        }
    }

    #[test]
    fn content_boundaries_and_utf16_clocks_are_preserved() {
        let update = raw_update(13, |encoder| {
            item(encoder, BLOCK_ITEM_BINARY_REF_NUMBER);
            encoder.write_buf([0, 1, 128, 255]);
            item(encoder, BLOCK_ITEM_STRING_REF_NUMBER);
            encoder.write_string("a🦀é");
            item(encoder, BLOCK_ITEM_EMBED_REF_NUMBER);
            encoder.write_string(r#"{"x":[null,true,false,-1.25e2,"\uD83E\uDD80"]}"#);
            item(encoder, BLOCK_ITEM_FORMAT_REF_NUMBER);
            encoder.write_key("bold");
            encoder.write_json(&Any::Bool(true));
            for kind in [0_u8, 1, 2, 3, 4, 5, 6, 9, 15] {
                item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                encoder.write_type_ref(kind);
                if kind == 3 {
                    encoder.write_key("element");
                }
            }
        });
        assert_boundaries(&update, 16);
    }

    #[test]
    fn production_cursor_parses_multi_key_json_object_keys_and_colons() {
        let update = raw_update(2, |encoder| {
            item(encoder, BLOCK_ITEM_EMBED_REF_NUMBER);
            encoder.write_string(r#"{"a":0,"b":1}"#);
            item(encoder, BLOCK_ITEM_FORMAT_REF_NUMBER);
            encoder.write_key("format");
            encoder.write_string(r#"{"a":0,"b":1}"#);
        });
        assert_eq!(
            split_update_v1_bounded(&update, 1, 384),
            Err(SplitError::UnsupportedMap),
        );
        let mut cursor = UpdateCursor::default();
        let mut causal = CausalState::default();
        let mut records = 0;
        while let Some(part) = cursor.next(&update, 1, 384).unwrap() {
            assert_eq!(part.records, 1);
            assert!(part.bytes.len() <= 384);
            causal.admit(&part.bytes).unwrap();
            records += part.records;
        }
        assert_eq!(records, 2);
        assert!(cursor.is_complete(&update).unwrap());
        for json in [r#"{"a"0}"#, r#"{"a":}"#, "{0:0}"] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_EMBED_REF_NUMBER);
                encoder.write_string(json);
            });
            assert_eq!(
                UpdateCursor::default().next(&update, 1, 384).err(),
                Some(SplitError::Malformed),
            );
            assert_eq!(
                CausalState::default().admit(&update),
                Err(SplitError::Malformed),
            );
        }
    }

    #[test]
    fn item_origins_and_parent_subtitles_are_scanned() {
        let update = raw_update(5, |encoder| {
            item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
            encoder.write_type_ref(TYPE_REFS_MAP);
            encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | HAS_PARENT_SUB);
            encoder.write_parent_info(false);
            encoder.write_var(7_u64);
            encoder.write_var(0_u32);
            encoder.write_string("key");
            encoder.write_len(1);
            encoder.write_any(&Any::from("first"));
            for flags in [HAS_ORIGIN, HAS_RIGHT_ORIGIN, HAS_ORIGIN | HAS_RIGHT_ORIGIN] {
                encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | HAS_PARENT_SUB | flags);
                if flags & HAS_ORIGIN != 0 {
                    encoder.write_var(7_u64);
                    encoder.write_var(1_u32);
                }
                if flags & HAS_RIGHT_ORIGIN != 0 {
                    encoder.write_var(7_u64);
                    encoder.write_var(1_u32);
                }
                encoder.write_len(1);
                encoder.write_any(&Any::Bool(true));
            }
        });
        assert_boundaries(&update, 5);
    }

    #[test]
    fn parent_ids_must_reference_shared_type_items() {
        let update = [1, 2, 7, 0, 8, 1, 1, 109, 1, 126, 8, 0, 7, 0, 1, 126, 0];
        for limit in LIMITS {
            assert_eq!(
                split_update_v1(&update, limit),
                Err(SplitError::MissingDependency)
            );
        }
        for kind in [
            BLOCK_GC_REF_NUMBER,
            BLOCK_SKIP_REF_NUMBER,
            BLOCK_ITEM_DELETED_REF_NUMBER,
            BLOCK_ITEM_ANY_REF_NUMBER,
        ] {
            for parent_clock in [1_u32, 2] {
                let update = raw_update(3, |encoder| {
                    item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                    encoder.write_type_ref(TYPE_REFS_MAP);
                    if matches!(kind, BLOCK_GC_REF_NUMBER | BLOCK_SKIP_REF_NUMBER) {
                        encoder.write_info(kind);
                    } else {
                        item(encoder, kind);
                    }
                    encoder.write_len(2);
                    if kind == BLOCK_ITEM_ANY_REF_NUMBER {
                        encoder.write_any(&Any::Null);
                        encoder.write_any(&Any::Null);
                    }
                    encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER);
                    encoder.write_parent_info(false);
                    encoder.write_var(7_u64);
                    encoder.write_var(parent_clock);
                    encoder.write_len(1);
                    encoder.write_any(&Any::Null);
                });
                for limit in LIMITS {
                    assert_eq!(
                        split_update_v1(&update, limit),
                        Err(SplitError::MissingDependency),
                        "accepted parent kind {kind} at clock {parent_clock}"
                    );
                }
            }
        }
    }

    #[test]
    fn origins_inside_gc_ranges_are_refused() {
        for flags in [HAS_ORIGIN, HAS_RIGHT_ORIGIN, HAS_ORIGIN | HAS_RIGHT_ORIGIN] {
            for origin_clock in [0_u32, 2] {
                let update = raw_update(3, |encoder| {
                    encoder.write_info(BLOCK_GC_REF_NUMBER);
                    encoder.write_len(3);
                    item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                    encoder.write_type_ref(TYPE_REFS_MAP);
                    encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | flags);
                    for flag in [HAS_ORIGIN, HAS_RIGHT_ORIGIN] {
                        if flags & flag != 0 {
                            encoder.write_var(7_u64);
                            encoder.write_var(origin_clock);
                        }
                    }
                    encoder.write_len(1);
                    encoder.write_any(&Any::Null);
                });
                for limit in LIMITS {
                    assert_eq!(
                        split_update_v1(&update, limit),
                        Err(SplitError::MissingDependency),
                        "accepted origin flags {flags} at GC clock {origin_clock}"
                    );
                }
            }
        }
    }

    #[test]
    fn adjacent_shared_type_parents_are_exact() {
        let update = raw_update(4, |encoder| {
            for _ in 0..2 {
                item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                encoder.write_type_ref(TYPE_REFS_MAP);
            }
            for parent_clock in [0_u32, 1] {
                encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | HAS_PARENT_SUB);
                encoder.write_parent_info(false);
                encoder.write_var(7_u64);
                encoder.write_var(parent_clock);
                encoder.write_string("key");
                encoder.write_len(1);
                encoder.write_any(&Any::Null);
            }
        });
        assert_boundaries(&update, 4);
    }

    #[test]
    fn production_cursor_refuses_duplicate_any_map_keys() {
        let update = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
            encoder.write_len(1);
            encoder.write_all(&[118, 2, 1, b'a', 126, 1, b'a', 126]);
        });
        let failure = UpdateCursor::default()
            .next(&update, 1, 16 * 1024)
            .err()
            .unwrap();
        assert_eq!(failure, SplitError::Malformed);
    }

    #[test]
    fn multi_key_any_maps_are_refused() {
        let value = Any::Map(std::sync::Arc::new(HashMap::from([
            ("a".into(), Any::Number(0.0)),
            ("b".into(), Any::Number(1.0)),
        ])));
        for value in [
            value.clone(),
            Any::Array(std::sync::Arc::from([value.clone()])),
            Any::Map(std::sync::Arc::new(HashMap::from([(
                "outer".into(),
                Any::Array(std::sync::Arc::from([value])),
            )]))),
        ] {
            let doc = Doc::with_client_id(7);
            let map = doc.get_or_insert_map("map");
            map.insert(&mut doc.transact_mut(), "value", value);
            let update = doc
                .transact()
                .encode_state_as_update_v1(&StateVector::default());
            for limit in LIMITS {
                assert_eq!(
                    split_update_v1(&update, limit),
                    Err(SplitError::UnsupportedMap)
                );
            }
        }
    }

    #[test]
    fn fallback_keeps_semantic_equality() {
        let doc = ordinary_doc();
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        assert!(split_update_v1(&update, 64).unwrap().len() > 1);
        let map = doc.get_or_insert_map("map");
        let value = Any::Map(std::sync::Arc::new(HashMap::from([
            ("a".into(), Any::Number(1.0)),
            ("b".into(), Any::Number(2.0)),
            ("c".into(), Any::from("x")),
        ])));
        {
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "value", value.clone());
            map.insert(
                &mut txn,
                "nested",
                Any::Array(std::sync::Arc::from([value])),
            );
        }
        let txn = doc.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = txn.state_vector();
        drop(txn);
        let content = canonical_snapshot(&doc);
        for limit in LIMITS {
            assert_eq!(
                split_update_v1(&update, limit),
                Err(SplitError::UnsupportedMap)
            );
            let split = split_or_whole_v1(&update, limit);
            assert_eq!(split.parts, vec![update.clone()]);
            assert_eq!(split.fallback, Some(SplitError::UnsupportedMap));
            assert_eq!(split.fallback.unwrap().reason(), "multi_key_map");
            let hydrated = Doc::with_client_id(doc.client_id().get());
            hydrate_snapshot_part(&hydrated, &split.parts[0]).unwrap();
            let txn = hydrated.transact();
            assert!(txn.store().pending_update().is_none());
            assert!(txn.store().pending_ds().is_none());
            assert_eq!(txn.state_vector(), vector);
            drop(txn);
            assert_eq!(canonical_snapshot(&hydrated), content);
        }
    }

    #[test]
    fn split_or_whole_matches_split_without_maps() {
        let doc = ordinary_doc();
        let txn = doc.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = vector_bytes(&txn.state_vector());
        drop(txn);
        for limit in LIMITS {
            let parts = split_update_v1(&update, limit).unwrap();
            let split = split_or_whole_v1(&update, limit);
            assert_eq!(split.parts, parts);
            assert_eq!(split.fallback, None);
            if limit == 64 {
                assert!(parts.len() > 1);
            }
            assert_parts(&update, &vector, doc.client_id().get(), limit, &parts);
        }
    }

    #[test]
    fn empty_and_single_key_any_maps_are_exact() {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        let empty = Any::Map(std::sync::Arc::new(HashMap::new()));
        let single = Any::Map(std::sync::Arc::new(HashMap::from([(
            "a".into(),
            Any::Number(0.0),
        )])));
        {
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "empty", empty.clone());
            map.insert(&mut txn, "single", single.clone());
            map.insert(
                &mut txn,
                "nested",
                Any::Map(std::sync::Arc::new(HashMap::from([(
                    "outer".into(),
                    Any::Array(std::sync::Arc::from([empty, single])),
                )]))),
            );
        }
        assert_doc(&doc);
    }

    #[test]
    fn multi_key_json_objects_are_refused() {
        for kind in [BLOCK_ITEM_EMBED_REF_NUMBER, BLOCK_ITEM_FORMAT_REF_NUMBER] {
            for json in [
                r#"{"a":0,"b":1}"#,
                r#"[{"a":0,"b":1}]"#,
                r#"{"outer":[{"a":0,"b":1}]}"#,
            ] {
                let update = raw_update(1, |encoder| {
                    item(encoder, kind);
                    if kind == BLOCK_ITEM_FORMAT_REF_NUMBER {
                        encoder.write_key("format");
                    }
                    encoder.write_string(json);
                });
                for limit in LIMITS {
                    assert_eq!(
                        split_update_v1(&update, limit),
                        Err(SplitError::UnsupportedMap)
                    );
                }
            }
        }
    }

    #[test]
    fn multi_key_subdocument_options_are_refused() {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        map.insert(&mut doc.transact_mut(), "subdoc", Doc::with_client_id(8));
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        for limit in LIMITS {
            assert_eq!(
                split_update_v1(&update, limit),
                Err(SplitError::UnsupportedMap)
            );
        }
    }

    #[test]
    fn empty_and_single_key_json_objects_are_exact() {
        let update = raw_update(8, |encoder| {
            for kind in [BLOCK_ITEM_EMBED_REF_NUMBER, BLOCK_ITEM_FORMAT_REF_NUMBER] {
                for json in ["{}", r#"{"a":0}"#, r#"[{"a":0}]"#, r#"{"outer":[{"a":0}]}"#] {
                    item(encoder, kind);
                    if kind == BLOCK_ITEM_FORMAT_REF_NUMBER {
                        encoder.write_key("format");
                    }
                    encoder.write_string(json);
                }
            }
        });
        assert_boundaries(&update, 8);
    }

    #[test]
    fn every_lib0_any_tag_is_scanned() {
        let values = vec![
            Any::Undefined,
            Any::Null,
            Any::Number(-12345.0),
            Any::Number(1.5),
            Any::Number(1.0 / 3.0),
            Any::BigInt(i64::MIN),
            Any::Bool(false),
            Any::Bool(true),
            Any::from("🦀"),
            Any::Buffer(std::sync::Arc::from([0_u8, 128, 255])),
            Any::Array(std::sync::Arc::from([Any::Null, Any::from("nested")])),
            Any::Map(std::sync::Arc::new(HashMap::from([(
                "key".into(),
                Any::Array(std::sync::Arc::from([Any::Bool(true)])),
            )]))),
        ];
        let update = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
            encoder.write_len(values.len() as u32);
            for value in &values {
                encoder.write_any(value);
            }
        });
        for limit in LIMITS {
            let parts = split_update_v1(&update, limit).unwrap();
            assert_eq!(parts, vec![update.clone()]);
            Update::decode_v1(&parts[0]).unwrap();
        }
        assert_boundaries(&update, values.len() as u32);
    }

    #[test]
    fn unsupported_or_inconsistent_content_is_refused() {
        let json = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_JSON_REF_NUMBER);
            encoder.write_len(1);
            encoder.write_string("null");
        });
        assert_eq!(
            split_update_v1(&json, 64),
            Err(SplitError::JsonLengthMismatch)
        );
        for kind in [11, 12, 13, 14, 15] {
            let update = raw_update(1, |encoder| item(encoder, kind));
            assert_eq!(
                split_update_v1(&update, 64),
                Err(SplitError::UnsupportedContent(kind))
            );
        }
        for kind in [7, 8, 10, 16, 128] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                encoder.write_type_ref(kind);
            });
            assert_eq!(
                split_update_v1(&update, 64),
                Err(SplitError::UnsupportedType(kind))
            );
        }
        let skip = raw_update(1, |encoder| {
            encoder.write_info(BLOCK_SKIP_REF_NUMBER);
            encoder.write_var(3_u32);
        });
        assert_eq!(
            split_update_v1(&skip, 64),
            Err(SplitError::MissingDependency)
        );
    }

    #[test]
    fn malformed_updates_are_refused() {
        let update = Workbook::from_model(small_model())
            .unwrap()
            .encode_state_as_update_v1();
        for end in 0..update.len() {
            assert!(
                split_update_v1(&update[..end], 64).is_err(),
                "accepted prefix {end}"
            );
        }
        for suffix in [vec![0], vec![255], vec![0, 0]] {
            let mut trailing = update.clone();
            trailing.extend(suffix);
            assert!(split_update_v1(&trailing, 64).is_err());
        }
        for bytes in [
            vec![128, 0, 0],
            vec![1, 0, 7, 0, 0],
            vec![1, 255, 255, 255, 255, 15],
            vec![0, 1, 7, 1, 0, 0],
            vec![255; 20],
        ] {
            assert!(split_update_v1(&bytes, 64).is_err());
        }
        assert_eq!(split_update_v1(&[0, 0], 1).unwrap(), vec![vec![0, 0]]);
        assert_eq!(split_update_v1(&update, 0), Err(SplitError::InvalidLimit));
    }

    #[test]
    fn malformed_values_are_refused() {
        for json in [
            "",
            "[1,]",
            "{\"x\":}",
            "01",
            "1e",
            "1e999",
            "true false",
            "\"\\uD800\"",
            "\"\\uDC00\"",
            "\"\\u+123\"",
            "\"\\q\"",
        ] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_EMBED_REF_NUMBER);
                encoder.write_string(json);
            });
            assert!(split_update_v1(&update, 64).is_err(), "accepted {json}");
        }
        for value in [
            vec![0],
            vec![125, 64],
            vec![125, 128, 0],
            vec![119, 1, 255],
            vec![117, 2, 126],
            vec![118, 1, 0],
            vec![116, 2, 0],
            vec![125, 255, 255, 255, 255, 255, 255, 255, 255, 255, 127],
        ] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
                encoder.write_len(1);
                encoder.write_all(&value);
            });
            assert!(split_update_v1(&update, 64).is_err());
        }
        let deeply_nested = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
            encoder.write_len(1);
            for _ in 0..64 {
                encoder.write_u8(117);
                encoder.write_var(1_u32);
            }
            encoder.write_u8(126);
        });
        assert!(split_update_v1(&deeply_nested, 64).is_err());
    }
}
