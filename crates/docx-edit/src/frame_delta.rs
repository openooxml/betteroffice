//! Binary `FrameDelta` v1 encoder for the resident display list.
//!
//! The wire format is deliberately independent of wasm-bindgen and JSON. A
//! fixed header and fixed-size page-operation table point at aligned primitive
//! id arrays and a compact, typed value stream. Containers carry both element
//! counts and byte lengths; the browser decoder rejects any mismatch before a
//! page reaches canvas replay.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use docx_layout::display_list::{DisplayList, DisplayPage, DocAttrs, Primitive};
#[cfg(test)]
use serde_json::Value;

mod typed_page;
use typed_page::encode_page;
#[cfg(test)]
use typed_page::hash_page;

pub const FRAME_DELTA_VERSION: u16 = 1;
pub const FRAME_HEADER_LEN: usize = 80;
pub const PAGE_OP_LEN: usize = 48;
pub const FRAME_FLAG_FULL: u32 = 1;
pub const PAGE_OP_UPSERT: u8 = 1;
pub const PAGE_OP_REMOVE: u8 = 2;
pub const PAGE_OP_MOVE: u8 = 3;
pub const PAGE_OP_PATCH_POSITIONS: u8 = 4;
pub const PAGE_OP_SHIFT_POSITIONS: u8 = 5;
/// Shifts body positions and spans on a contiguous range of pages.
pub const PAGE_OP_SHIFT_RANGE: u8 = 6;
const SHIFT_SPAN_PRESENT: u32 = 1;

const POSITION_DOC_START: u8 = 1 << 0;
const POSITION_DOC_END: u8 = 1 << 1;
const POSITION_FRAGMENT_START: u8 = 1 << 2;
const POSITION_FRAGMENT_END: u8 = 1 << 3;
const POSITION_INLINE_WIDGET: u8 = 1 << 4;
/// Run flag: shift only the masked fields each primitive has, so one run
/// covers primitives that carry different position fields.
const POSITION_PRESENT_ONLY: u8 = 1 << 7;
const POSITION_FIELDS: [u8; 5] = [
    POSITION_DOC_START,
    POSITION_DOC_END,
    POSITION_FRAGMENT_START,
    POSITION_FRAGMENT_END,
    POSITION_INLINE_WIDGET,
];

const MAGIC: [u8; 4] = *b"FDV1";
const MAX_U32: usize = u32::MAX as usize;
const MAX_SAFE_INTEGER: i64 = (1 << 53) - 1;
const FNV_OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
const FNV_PRIME: u64 = 0x0000_0100_0000_01b3;

#[derive(Clone, Debug, PartialEq)]
pub struct FramePageSnapshot {
    pub page_id: u64,
    pub anchor: String,
    pub fingerprint: u64,
    pub visual_fingerprint: u64,
    pub page_index: u32,
    /// Shared with the next frame's snapshot when the page's primitive
    /// identity is unchanged — cloning a snapshot never copies the id array.
    pub primitive_ids: Rc<[u64]>,
    pub positions: Vec<PrimitivePositionSnapshot>,
    /// Every note region note's anchor, in area then note order.
    pub note_anchors: Vec<NoteAnchorSnapshot>,
    placeholder: Option<Rc<DisplayPage>>,
    placeholder_hash: u64,
    position_span: Option<[i64; 2]>,
    position_base: i64,
    body_primitives: u32,
}

impl FramePageSnapshot {
    /// Folds deferred body position shifts into the stored positions.
    pub fn materialize_positions(&mut self) {
        let delta = std::mem::take(&mut self.position_base);
        if delta == 0 {
            return;
        }
        for position in self
            .positions
            .iter_mut()
            .take(self.body_primitives as usize)
        {
            for value in [
                &mut position.doc_start,
                &mut position.doc_end,
                &mut position.fragment_doc_start,
                &mut position.fragment_doc_end,
                &mut position.inline_widget_pos,
            ]
            .into_iter()
            .flatten()
            {
                *value += delta;
            }
        }
    }
}

/// One note region note's backlink anchor, addressed by area and note index.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct NoteAnchorSnapshot {
    pub area: u32,
    pub note: u32,
    pub start: Option<i64>,
    pub end: Option<i64>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct PrimitivePositionSnapshot {
    pub doc_start: Option<i64>,
    pub doc_end: Option<i64>,
    pub fragment_doc_start: Option<i64>,
    pub fragment_doc_end: Option<i64>,
    pub inline_widget_pos: Option<i64>,
}

#[derive(Clone, Copy, Debug)]
pub struct FrameEpochs {
    pub doc_epoch: u64,
    pub layout_epoch: u64,
    pub frame_epoch: u64,
    pub base_frame_epoch: u64,
}

/// A contiguous range of pages whose body positions move by one delta.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PageShiftRun {
    pub start: usize,
    pub end: usize,
    pub delta: i64,
}

/// Pages changed against an existing frame with stable page indices.
pub struct DisplayChanges<'a> {
    /// Pages whose display page was rebuilt, built, released or replaced: prepared afresh.
    pub rebuilt: &'a [usize],
    /// Clean pages whose positions moved non-uniformly: compared position by position.
    pub repositioned: &'a [usize],
    /// Clean pages whose positions all moved by one delta.
    pub shifts: &'a [PageShiftRun],
}

/// A page or range to process in ascending index order.
enum PageUpdate<'a> {
    Rebuilt(usize),
    Repositioned(usize),
    Shift(&'a PageShiftRun),
}

impl PageUpdate<'_> {
    /// Returns the page indices covered by this update.
    fn range(&self) -> std::ops::Range<usize> {
        match self {
            Self::Rebuilt(index) | Self::Repositioned(index) => *index..*index + 1,
            Self::Shift(run) => run.start..run.end,
        }
    }
}

#[derive(Debug)]
struct PreparedPage<'a> {
    snapshot: FramePageSnapshot,
    page: &'a DisplayPage,
    change: PageChange,
    /// Where the page's primitive ids and payload went in the data section
    /// when it was emitted while it was fingerprinted and then upserts.
    emitted: Option<EmittedPage>,
}

impl<'a> PreparedPage<'a> {
    /// Returns the operation for a page whose snapshot changed.
    fn op(&self) -> Option<PageOp<'_, 'a>> {
        match &self.change {
            PageChange::Upsert => Some(PageOp::Upsert(self)),
            PageChange::Move => Some(PageOp::Move(self)),
            PageChange::PatchPositions(patches) => Some(PageOp::PatchPositions(self, patches)),
            PageChange::ShiftPositions(runs, anchors, delta) => {
                Some(PageOp::ShiftPositions(self, runs, anchors, *delta))
            }
            PageChange::Retain => None,
        }
    }
}

#[derive(Debug)]
struct EmittedPage {
    primitive_ids_offset: usize,
    payload_offset: usize,
    payload_len: usize,
}

/// What a frame sends for a page, against the page's previous snapshot.
#[derive(Debug)]
enum PageChange {
    Upsert,
    Move,
    PatchPositions(Vec<PositionPatch>),
    ShiftPositions(Vec<PositionShiftRun>, Vec<NoteAnchorSnapshot>, Option<i64>),
    Retain,
}

/// A frame's string table and data section, which pages emit into while
/// they are prepared.
#[derive(Default)]
struct FrameData {
    strings: StringTable,
    out: Vec<u8>,
}

#[derive(Debug)]
enum PageOp<'a, 'b> {
    Upsert(&'a PreparedPage<'b>),
    Remove(&'a FramePageSnapshot),
    Move(&'a PreparedPage<'b>),
    PatchPositions(&'a PreparedPage<'b>, &'a [PositionPatch]),
    ShiftPositions(
        &'a PreparedPage<'b>,
        &'a [PositionShiftRun],
        &'a [NoteAnchorSnapshot],
        Option<i64>,
    ),
    ShiftRange(&'a PageShiftRun, u64),
}

#[derive(Debug)]
struct PositionPatch {
    primitive_id: u64,
    changed_mask: u8,
    present_mask: u8,
    values: [Option<i64>; 5],
}

#[derive(Debug)]
struct PositionShiftRun {
    start: u32,
    count: u32,
    changed_mask: u8,
    delta: i64,
}

/// Encode one full recovery frame or a delta against `previous`.
///
/// `next_page_id` is session-owned and monotonic. Page ids are matched first
/// by their semantic page-start anchor and then by the old surface index, so a
/// page keeps its identity across ordinary edits and most pagination shifts.
pub fn encode_frame_delta(
    list: &DisplayList,
    previous: &[FramePageSnapshot],
    epochs: FrameEpochs,
    full: bool,
    next_page_id: &mut u64,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    encode_frame_delta_inner(list, previous, epochs, full, next_page_id, None)
}

/// Incremental encoder that fully prepares only display-rebuilt pages. Clean
/// pages already retain stable visual content and primitive identity; walking
/// their positions is enough to emit geometry patches without serializing and
/// hashing the complete page value again.
pub fn encode_frame_delta_incremental(
    list: &DisplayList,
    previous: &[FramePageSnapshot],
    epochs: FrameEpochs,
    next_page_id: &mut u64,
    rebuilt_pages: &HashSet<usize>,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    encode_frame_delta_inner(
        list,
        previous,
        epochs,
        false,
        next_page_id,
        Some(&|index| rebuilt_pages.contains(&index)),
    )
}

/// [`encode_frame_delta_incremental`] for pages built into an unchanged
/// layout: every page keeps its index, so pages match their previous snapshot
/// by index. Semantic anchors count occurrences across built pages, and
/// building an earlier page renumbers the ones after it.
pub fn encode_frame_delta_pages(
    list: &DisplayList,
    previous: &[FramePageSnapshot],
    epochs: FrameEpochs,
    next_page_id: &mut u64,
    rebuilt: &dyn Fn(usize) -> bool,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    let mut data = FrameData::default();
    let prepared = prepare_pages(
        list,
        previous,
        next_page_id,
        Some(rebuilt),
        PrepareOptions {
            match_anchors: false,
            full: false,
        },
        &mut data,
    )?;
    encode_prepared(list, previous, epochs, false, prepared, data)
}

/// Encodes only changed pages and applies uniform range shifts to snapshots in place.
pub fn encode_frame_delta_changes(
    list: &DisplayList,
    snapshots: &mut [FramePageSnapshot],
    epochs: FrameEpochs,
    changes: DisplayChanges<'_>,
) -> Result<Vec<u8>, String> {
    if snapshots.len() != list.pages.len() {
        return Err("FrameDelta snapshot count does not match the display list".to_owned());
    }
    checked_u32(list.pages.len(), "page count")?;
    let mut rebuilt = changes.rebuilt.to_vec();
    rebuilt.sort_unstable();
    rebuilt.dedup();
    let mut repositioned = changes.repositioned.to_vec();
    repositioned.sort_unstable();
    repositioned.dedup();
    if rebuilt
        .iter()
        .chain(&repositioned)
        .any(|&index| index >= list.pages.len())
    {
        return Err("FrameDelta changed page index is out of range".to_owned());
    }
    for run in changes.shifts {
        if run.start >= run.end || run.end > list.pages.len() {
            return Err("FrameDelta page shift range is invalid".to_owned());
        }
        if run.delta == 0 || !(-MAX_SAFE_INTEGER..=MAX_SAFE_INTEGER).contains(&run.delta) {
            return Err("FrameDelta page shift delta is invalid".to_owned());
        }
    }
    let mut updates: Vec<_> = rebuilt
        .into_iter()
        .map(PageUpdate::Rebuilt)
        .chain(repositioned.into_iter().map(PageUpdate::Repositioned))
        .chain(changes.shifts.iter().map(PageUpdate::Shift))
        .collect();
    updates.sort_unstable_by_key(|update| update.range().start);
    let mut end = 0;
    for update in &updates {
        let range = update.range();
        if range.start < end {
            return Err("FrameDelta changed pages and shift ranges overlap".to_owned());
        }
        end = range.end;
    }

    let mut data = FrameData::default();
    let mut placeholder_data = FrameData::default();
    let mut prepared = Vec::new();
    for update in &updates {
        match update {
            PageUpdate::Rebuilt(index) | PageUpdate::Repositioned(index) => {
                let old = &mut snapshots[*index];
                old.materialize_positions();
                let options = PagePreparation {
                    page_id: old.page_id,
                    page_index: checked_u32(*index, "page index")?,
                    anchor: old.anchor.clone(),
                    is_new: false,
                    moved: false,
                    full: false,
                    rebuild: matches!(update, PageUpdate::Rebuilt(_)),
                };
                prepared.push(prepare_page(
                    &list.pages[*index],
                    Some(old),
                    options,
                    &mut data,
                    &mut placeholder_data,
                )?);
            }
            PageUpdate::Shift(run) => {
                if snapshots[run.start].page_id == 0 {
                    return Err("FrameDelta shifted page id is zero".to_owned());
                }
                for old in &snapshots[run.start..run.end] {
                    old.position_base
                        .checked_add(run.delta)
                        .ok_or_else(|| "FrameDelta position base overflow".to_owned())?;
                    if let Some(span) = old.position_span {
                        for value in span {
                            value
                                .checked_add(run.delta)
                                .ok_or_else(|| "FrameDelta position span overflow".to_owned())?;
                        }
                    }
                }
            }
        }
    }
    let mut ops = Vec::new();
    let mut pages = prepared.iter();
    for update in &updates {
        match update {
            PageUpdate::Shift(run) => {
                ops.push(PageOp::ShiftRange(run, snapshots[run.start].page_id));
            }
            _ => {
                if let Some(op) = pages.next().expect("prepared changed page").op() {
                    ops.push(op);
                }
            }
        }
    }
    let bytes = encode_ops(list, epochs, false, &ops, data)?;
    drop(ops);
    for page in prepared {
        let index = page.snapshot.page_index as usize;
        snapshots[index] = page.snapshot;
    }
    for run in changes.shifts {
        let salt = shift_range_salt(run.delta, epochs.frame_epoch);
        for old in &mut snapshots[run.start..run.end] {
            old.position_base += run.delta;
            if let Some(span) = &mut old.position_span {
                span[0] += run.delta;
                span[1] += run.delta;
            }
            old.fingerprint ^= salt;
        }
    }
    Ok(bytes)
}

fn encode_frame_delta_inner(
    list: &DisplayList,
    previous: &[FramePageSnapshot],
    epochs: FrameEpochs,
    full: bool,
    next_page_id: &mut u64,
    rebuilt_pages: Option<&dyn Fn(usize) -> bool>,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    let mut data = FrameData::default();
    let prepared = prepare_pages(
        list,
        previous,
        next_page_id,
        rebuilt_pages,
        PrepareOptions {
            match_anchors: true,
            full,
        },
        &mut data,
    )?;
    encode_prepared(list, previous, epochs, full, prepared, data)
}

fn encode_prepared(
    list: &DisplayList,
    previous: &[FramePageSnapshot],
    epochs: FrameEpochs,
    full: bool,
    prepared: Vec<PreparedPage<'_>>,
    data: FrameData,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    let mut ops = Vec::new();
    if !full {
        let next_ids: HashSet<u64> = prepared.iter().map(|page| page.snapshot.page_id).collect();
        for old in previous {
            if !next_ids.contains(&old.page_id) {
                ops.push(PageOp::Remove(old));
            }
        }
    }
    ops.extend(prepared.iter().filter_map(PreparedPage::op));
    let bytes = encode_ops(list, epochs, full, &ops, data)?;
    drop(ops);
    let next_snapshots = prepared.into_iter().map(|page| page.snapshot).collect();
    Ok((bytes, next_snapshots))
}

/// Writes operation records and assembles the frame's data and header.
fn encode_ops(
    list: &DisplayList,
    epochs: FrameEpochs,
    full: bool,
    ops: &[PageOp<'_, '_>],
    data: FrameData,
) -> Result<Vec<u8>, String> {
    let op_count = checked_u32(ops.len(), "page operation count")?;
    let page_count = checked_u32(list.pages.len(), "page count")?;
    let ops_bytes = ops
        .len()
        .checked_mul(PAGE_OP_LEN)
        .ok_or_else(|| "FrameDelta operation table overflow".to_owned())?;
    let strings_offset = FRAME_HEADER_LEN
        .checked_add(ops_bytes)
        .ok_or_else(|| "FrameDelta header overflow".to_owned())?;

    // Page payloads intern their strings while they are written, so the data
    // section is built first and placed after the finished string table.
    // Offsets recorded here are relative to the data section, whose start is
    // 8-byte aligned, so relative alignment is absolute alignment.
    let mut records = vec![0_u8; ops_bytes];
    let mut data_offsets = Vec::new();
    let FrameData {
        mut strings,
        mut out,
    } = data;
    for (op_index, op) in ops.iter().enumerate() {
        let record = op_index * PAGE_OP_LEN;
        match op {
            PageOp::Upsert(page) => {
                records[record] = PAGE_OP_UPSERT;
                patch_u32(&mut records, record + 4, page.snapshot.page_index);
                patch_u64(&mut records, record + 8, page.snapshot.page_id);
                patch_u64(&mut records, record + 16, page.snapshot.fingerprint);
                patch_u32(
                    &mut records,
                    record + 24,
                    checked_u32(page.snapshot.primitive_ids.len(), "primitive id count")?,
                );

                let emitted = match &page.emitted {
                    Some(emitted) => emitted,
                    None => {
                        &emit_page(
                            page.page,
                            &page.snapshot.primitive_ids,
                            &mut strings,
                            &mut out,
                        )?
                        .0
                    }
                };
                data_offsets.push(record + 28);
                patch_u32(
                    &mut records,
                    record + 28,
                    checked_u32(emitted.primitive_ids_offset, "primitive id offset")?,
                );
                let (payload_offset, payload_len) = (emitted.payload_offset, emitted.payload_len);
                data_offsets.push(record + 32);
                patch_u32(
                    &mut records,
                    record + 32,
                    checked_u32(payload_offset, "page payload offset")?,
                );
                patch_u32(
                    &mut records,
                    record + 36,
                    checked_u32(payload_len, "page payload length")?,
                );
            }
            PageOp::Remove(page) => {
                records[record] = PAGE_OP_REMOVE;
                patch_u32(&mut records, record + 4, page.page_index);
                patch_u64(&mut records, record + 8, page.page_id);
            }
            PageOp::Move(page) => {
                records[record] = PAGE_OP_MOVE;
                patch_u32(&mut records, record + 4, page.snapshot.page_index);
                patch_u64(&mut records, record + 8, page.snapshot.page_id);
                patch_u64(&mut records, record + 16, page.snapshot.fingerprint);
            }
            PageOp::PatchPositions(page, patches) => {
                records[record] = PAGE_OP_PATCH_POSITIONS;
                patch_u32(&mut records, record + 4, page.snapshot.page_index);
                patch_u64(&mut records, record + 8, page.snapshot.page_id);
                patch_u64(&mut records, record + 16, page.snapshot.fingerprint);
                patch_u32(
                    &mut records,
                    record + 24,
                    checked_u32(patches.len(), "position patch count")?,
                );
                align(&mut out, 8);
                let payload_offset = out.len();
                write_u32(
                    &mut out,
                    checked_u32(patches.len(), "position patch count")?,
                );
                write_u32(&mut out, 0);
                for patch in patches.iter() {
                    write_u64(&mut out, patch.primitive_id);
                    out.push(patch.changed_mask);
                    out.push(patch.present_mask);
                    write_u16(&mut out, 0);
                    for (index, field) in POSITION_FIELDS.iter().enumerate() {
                        if patch.present_mask & field != 0 {
                            write_i64(
                                &mut out,
                                patch.values[index]
                                    .expect("present position-patch field carries a value"),
                            );
                        }
                    }
                }
                data_offsets.push(record + 32);
                patch_u32(
                    &mut records,
                    record + 32,
                    checked_u32(payload_offset, "position patch payload offset")?,
                );
                let payload_length = out.len() - payload_offset;
                patch_u32(
                    &mut records,
                    record + 36,
                    checked_u32(payload_length, "position patch payload length")?,
                );
            }
            PageOp::ShiftPositions(page, runs, anchors, span_delta) => {
                records[record] = PAGE_OP_SHIFT_POSITIONS;
                patch_u32(&mut records, record + 4, page.snapshot.page_index);
                patch_u64(&mut records, record + 8, page.snapshot.page_id);
                patch_u64(&mut records, record + 16, page.snapshot.fingerprint);
                patch_u32(
                    &mut records,
                    record + 24,
                    checked_u32(runs.len(), "position shift run count")?,
                );
                patch_u32(
                    &mut records,
                    record + 40,
                    checked_u32(anchors.len(), "note anchor count")?,
                );
                align(&mut out, 8);
                let payload_offset = out.len();
                write_u32(
                    &mut out,
                    checked_u32(runs.len(), "position shift run count")?,
                );
                write_u32(
                    &mut out,
                    if span_delta.is_some() {
                        SHIFT_SPAN_PRESENT
                    } else {
                        0
                    },
                );
                if let Some(delta) = span_delta {
                    write_i64(&mut out, *delta);
                }
                for run in runs.iter() {
                    write_u32(&mut out, run.start);
                    write_u32(&mut out, run.count);
                    out.push(run.changed_mask);
                    out.extend_from_slice(&[0; 7]);
                    write_i64(&mut out, run.delta);
                }
                if !anchors.is_empty() {
                    write_u32(&mut out, checked_u32(anchors.len(), "note anchor count")?);
                    write_u32(&mut out, 0);
                    for anchor in anchors.iter() {
                        write_u32(&mut out, anchor.area);
                        write_u32(&mut out, anchor.note);
                        write_i64(&mut out, anchor.start.unwrap_or(i64::MIN));
                        write_i64(&mut out, anchor.end.unwrap_or(i64::MIN));
                    }
                }
                data_offsets.push(record + 32);
                patch_u32(
                    &mut records,
                    record + 32,
                    checked_u32(payload_offset, "position shift payload offset")?,
                );
                let payload_length = out.len() - payload_offset;
                patch_u32(
                    &mut records,
                    record + 36,
                    checked_u32(payload_length, "position shift payload length")?,
                );
            }
            PageOp::ShiftRange(run, page_id) => {
                records[record] = PAGE_OP_SHIFT_RANGE;
                patch_u32(
                    &mut records,
                    record + 4,
                    checked_u32(run.start, "page index")?,
                );
                patch_u64(&mut records, record + 8, *page_id);
                patch_u64(
                    &mut records,
                    record + 16,
                    shift_range_salt(run.delta, epochs.frame_epoch),
                );
                patch_u32(
                    &mut records,
                    record + 24,
                    checked_u32(run.end - run.start, "page shift count")?,
                );
                records[record + 32..record + 40].copy_from_slice(&run.delta.to_le_bytes());
            }
        }
    }

    let mut prefix = vec![0; FRAME_HEADER_LEN];
    prefix.extend_from_slice(&records);
    let strings = strings.into_strings();
    write_u32(&mut prefix, checked_u32(strings.len(), "string count")?);
    for value in &strings {
        write_u32(&mut prefix, checked_u32(value.len(), "string byte length")?);
        prefix.extend_from_slice(value.as_bytes());
    }
    let strings_len = prefix.len() - strings_offset;
    align(&mut prefix, 8);
    let data_offset = prefix.len();
    for offset in data_offsets {
        let at = FRAME_HEADER_LEN + offset;
        let relative = u32::from_le_bytes(prefix[at..at + 4].try_into().expect("u32 field"));
        patch_u32(
            &mut prefix,
            at,
            checked_u32(data_offset + relative as usize, "data offset")?,
        );
    }
    out.reserve_exact(prefix.len());
    out.splice(0..0, prefix);

    if out.len() > MAX_U32 {
        return Err("FrameDelta exceeds the v1 u32 byte-length limit".to_owned());
    }
    out[0..4].copy_from_slice(&MAGIC);
    patch_u16(&mut out, 4, FRAME_DELTA_VERSION);
    patch_u16(&mut out, 6, FRAME_HEADER_LEN as u16);
    let total_len = out.len() as u32;
    patch_u32(&mut out, 8, total_len);
    patch_u32(&mut out, 12, if full { FRAME_FLAG_FULL } else { 0 });
    patch_u64(&mut out, 16, epochs.doc_epoch);
    patch_u64(&mut out, 24, epochs.layout_epoch);
    patch_u64(&mut out, 32, epochs.frame_epoch);
    patch_u64(&mut out, 40, if full { 0 } else { epochs.base_frame_epoch });
    patch_u32(&mut out, 48, page_count);
    patch_u32(&mut out, 52, op_count);
    patch_u32(&mut out, 56, FRAME_HEADER_LEN as u32);
    patch_u32(
        &mut out,
        60,
        checked_u32(strings_offset, "string table offset")?,
    );
    patch_u32(
        &mut out,
        64,
        checked_u32(strings_len, "string table length")?,
    );
    patch_u32(
        &mut out,
        68,
        checked_u32(data_offset, "data section offset")?,
    );
    patch_u32(&mut out, 72, list.contract_version.unwrap_or_default());

    Ok(out)
}

struct PrepareOptions {
    match_anchors: bool,
    full: bool,
}

/// `snapshots` with every deferred position shift folded in, borrowed when
/// none is deferred.
fn materialized_snapshots(snapshots: &[FramePageSnapshot]) -> Cow<'_, [FramePageSnapshot]> {
    if snapshots.iter().all(|snapshot| snapshot.position_base == 0) {
        return Cow::Borrowed(snapshots);
    }
    Cow::Owned(
        snapshots
            .iter()
            .cloned()
            .map(|mut snapshot| {
                snapshot.materialize_positions();
                snapshot
            })
            .collect(),
    )
}

fn prepare_pages<'a>(
    list: &'a DisplayList,
    previous: &[FramePageSnapshot],
    next_page_id: &mut u64,
    rebuilt_pages: Option<&dyn Fn(usize) -> bool>,
    options: PrepareOptions,
    data: &mut FrameData,
) -> Result<Vec<PreparedPage<'a>>, String> {
    let PrepareOptions {
        match_anchors,
        full,
    } = options;
    let materialized = materialized_snapshots(previous);
    let previous: &[FramePageSnapshot] = &materialized;
    let anchors = page_anchors(list);
    // Anchors are unique within one snapshot list (page_anchors suffixes an
    // occurrence counter), so keyed lookups replace the old per-page scans.
    let mut previous_by_anchor: HashMap<&str, usize> = previous
        .iter()
        .enumerate()
        .map(|(index, old)| (old.anchor.as_str(), index))
        .collect();
    let mut previous_by_index: HashMap<u32, usize> = previous
        .iter()
        .enumerate()
        .map(|(index, old)| (old.page_index, index))
        .collect();
    let mut claimed = HashSet::new();
    let mut matched_previous = vec![None; list.pages.len()];
    // Reserve every semantic anchor before considering the index fallback. A
    // newly inserted leading page must not steal the id of the old page at
    // index zero and shift every retained surface identity after it.
    if match_anchors {
        for (next_index, anchor) in anchors.iter().enumerate() {
            if let Some(previous_index) = previous_by_anchor.remove(anchor.as_str()) {
                claimed.insert(previous[previous_index].page_id);
                matched_previous[next_index] = Some(previous_index);
            }
        }
    }
    for (next_index, matched) in matched_previous.iter_mut().enumerate() {
        if matched.is_some() {
            continue;
        }
        let page_index = checked_u32(next_index, "page index")?;
        if let Some(previous_index) = previous_by_index.remove(&page_index)
            && claimed.insert(previous[previous_index].page_id)
        {
            *matched = Some(previous_index);
        }
    }

    // Clean pages keep their index before the first rebuilt page and shift by
    // the page-count change after it. A page built in this frame can renumber
    // the anchors after it, so a clean page whose anchor match disagrees is
    // prepared afresh.
    let first_rebuilt =
        rebuilt_pages.and_then(|rebuilt| (0..list.pages.len()).find(|&index| rebuilt(index)));
    let page_count_change = list.pages.len() as i64 - previous.len() as i64;
    let retained_index = |index: usize| -> i64 {
        if first_rebuilt.is_some_and(|first| index >= first) {
            index as i64 - page_count_change
        } else {
            index as i64
        }
    };
    let mut prepared = Vec::with_capacity(list.pages.len());
    let mut placeholder_data = FrameData::default();
    for ((index, page), anchor) in list.pages.iter().enumerate().zip(anchors) {
        let page_index = checked_u32(index, "page index")?;
        let matched = matched_previous[index].map(|previous_index| &previous[previous_index]);
        let (page_id, is_new, moved) = if let Some(old) = matched {
            (old.page_id, false, old.page_index != page_index)
        } else {
            *next_page_id = next_page_id
                .checked_add(1)
                .ok_or_else(|| "FrameDelta page id space exhausted".to_owned())?;
            (*next_page_id, true, false)
        };
        prepared.push(prepare_page(
            page,
            matched,
            PagePreparation {
                page_id,
                page_index,
                anchor,
                is_new,
                moved,
                full,
                rebuild: rebuilt_pages.is_none_or(|rebuilt| rebuilt(index))
                    || matched.is_none_or(|old| i64::from(old.page_index) != retained_index(index)),
            },
            data,
            &mut placeholder_data,
        )?);
    }
    Ok(prepared)
}

/// Identity and preparation mode for one display page.
struct PagePreparation {
    page_id: u64,
    page_index: u32,
    anchor: String,
    is_new: bool,
    moved: bool,
    full: bool,
    rebuild: bool,
}

/// Prepares one page and keeps emitted bytes only when it upserts.
fn prepare_page<'a>(
    page: &'a DisplayPage,
    old: Option<&FramePageSnapshot>,
    options: PagePreparation,
    data: &mut FrameData,
    placeholder_data: &mut FrameData,
) -> Result<PreparedPage<'a>, String> {
    let PagePreparation {
        page_id,
        page_index,
        anchor,
        is_new,
        moved,
        full,
        rebuild,
    } = options;
    let positions = primitive_positions(page);
    let note_anchors = note_anchors(page)?;
    let placeholder = page.unbuilt.then(|| {
        let mut normalized = page.clone();
        normalized.page_index = 0;
        normalized.position_span = None;
        Rc::new(normalized)
    });
    let same_placeholder = placeholder
        .as_ref()
        .is_some_and(|next| old.and_then(|old| old.placeholder.as_ref()) == Some(next));
    let placeholder_hash = if same_placeholder {
        old.expect("matched placeholder").placeholder_hash
    } else if let Some(normalized) = &placeholder {
        placeholder_data.strings.rollback(0);
        placeholder_data.out.clear();
        encode_page(
            normalized,
            &mut placeholder_data.strings,
            &mut placeholder_data.out,
        )?
        .fingerprint
    } else {
        0
    };
    let full_prepare = !same_placeholder && (is_new || page.unbuilt || rebuild);
    let mut emitted = None;
    let (fingerprint, visual_fingerprint, primitive_ids) = if full_prepare {
        let primitive_ids: Rc<[u64]> = primitive_ids(page, page_id).into();
        let mark = (data.out.len(), data.strings.mark());
        let (page_emitted, hashes) =
            emit_page(page, &primitive_ids, &mut data.strings, &mut data.out)?;
        emitted = Some((page_emitted, mark));
        (hashes.fingerprint, hashes.visual_fingerprint, primitive_ids)
    } else {
        let old = old.expect("clean incremental pages retain a previous snapshot");
        let fingerprint = if positions == old.positions && note_anchors == old.note_anchors {
            old.fingerprint
        } else {
            hash_positions(old.visual_fingerprint, &positions, &note_anchors)
        };
        (
            fingerprint,
            old.visual_fingerprint,
            Rc::clone(&old.primitive_ids),
        )
    };
    let (fingerprint, visual_fingerprint) = if placeholder.is_some() {
        let mut fingerprint = mix(placeholder_hash, u64::from(page.position_span.is_some()));
        if let Some([start, end]) = page.position_span {
            fingerprint = mix(mix(fingerprint, start as u64), end as u64);
        }
        (fingerprint, placeholder_hash)
    } else {
        (fingerprint, visual_fingerprint)
    };
    let mut snapshot = FramePageSnapshot {
        page_id,
        anchor,
        fingerprint,
        visual_fingerprint,
        page_index,
        primitive_ids,
        positions,
        note_anchors,
        placeholder: if same_placeholder {
            old.and_then(|old| old.placeholder.clone())
        } else {
            placeholder
        },
        placeholder_hash,
        position_span: page.position_span,
        position_base: 0,
        body_primitives: checked_u32(page.primitives.len(), "body primitive count")?,
    };
    let mut change = page_change(full, is_new, moved, old, &snapshot);
    if matches!(change, PageChange::Upsert)
        && !full
        && !is_new
        && !moved
        && let Some(old) = old
        && old.visual_fingerprint == snapshot.visual_fingerprint
        && old.primitive_ids == snapshot.primitive_ids
        && old.positions == snapshot.positions
        && old.note_anchors == snapshot.note_anchors
        && old.placeholder == snapshot.placeholder
        && old.position_span == snapshot.position_span
    {
        change = PageChange::Retain;
        snapshot.fingerprint = old.fingerprint;
    }
    let emitted = match emitted {
        Some((emitted, _)) if matches!(change, PageChange::Upsert) => Some(emitted),
        Some((_, (end, mark))) => {
            data.out.truncate(end);
            data.strings.rollback(mark);
            None
        }
        None => None,
    };
    Ok(PreparedPage {
        snapshot,
        page,
        change,
        emitted,
    })
}

/// Emits `page`'s aligned primitive ids and then its payload.
fn emit_page(
    page: &DisplayPage,
    primitive_ids: &[u64],
    strings: &mut StringTable,
    out: &mut Vec<u8>,
) -> Result<(EmittedPage, typed_page::PageHashes), String> {
    align(out, 8);
    let primitive_ids_offset = out.len();
    for id in primitive_ids {
        write_u64(out, *id);
    }
    let payload_offset = out.len();
    let hashes = encode_page(page, strings, out)?;
    Ok((
        EmittedPage {
            primitive_ids_offset,
            payload_offset,
            payload_len: out.len() - payload_offset,
        },
        hashes,
    ))
}

fn page_change(
    full: bool,
    is_new: bool,
    moved: bool,
    old: Option<&FramePageSnapshot>,
    next: &FramePageSnapshot,
) -> PageChange {
    let Some(old) = old.filter(|_| !full && !is_new) else {
        return PageChange::Upsert;
    };
    if old.placeholder.is_some() || next.placeholder.is_some() {
        if old.placeholder != next.placeholder {
            return PageChange::Upsert;
        }
        if old.position_span != next.position_span {
            if let (Some([start, end]), Some([next_start, next_end])) =
                (old.position_span, next.position_span)
                && let Some(delta) = next_start.checked_sub(start)
                && next_end.checked_sub(end) == Some(delta)
                && delta != 0
                && [start, end, next_start, next_end, delta]
                    .into_iter()
                    .all(|value| (-MAX_SAFE_INTEGER..=MAX_SAFE_INTEGER).contains(&value))
            {
                return PageChange::ShiftPositions(Vec::new(), Vec::new(), Some(delta));
            }
            return PageChange::Upsert;
        }
    }
    if old.fingerprint == next.fingerprint {
        return if moved {
            PageChange::Move
        } else {
            PageChange::Retain
        };
    }
    if old.visual_fingerprint != next.visual_fingerprint || old.primitive_ids != next.primitive_ids
    {
        return PageChange::Upsert;
    }
    let patches = position_patches(old, next);
    let anchors = changed_note_anchors(&old.note_anchors, &next.note_anchors);
    let runs = if patches.is_empty() {
        Some(Vec::new())
    } else {
        position_shift_runs(&old.positions, &next.positions)
    };
    match (anchors, runs) {
        (Some(anchors), Some(runs)) if !anchors.is_empty() || !runs.is_empty() => {
            PageChange::ShiftPositions(runs, anchors, None)
        }
        (Some(anchors), None) if anchors.is_empty() => PageChange::PatchPositions(patches),
        _ => PageChange::Upsert,
    }
}

fn hash_positions(
    visual_fingerprint: u64,
    positions: &[PrimitivePositionSnapshot],
    note_anchors: &[NoteAnchorSnapshot],
) -> u64 {
    let mut hash = visual_fingerprint;
    for anchor in note_anchors {
        for value in [anchor.start, anchor.end] {
            hash ^= value.unwrap_or(i64::MIN) as u64;
            hash = hash.wrapping_mul(FNV_PRIME);
        }
    }
    for position in positions {
        for value in [
            position.doc_start,
            position.doc_end,
            position.fragment_doc_start,
            position.fragment_doc_end,
            position.inline_widget_pos,
        ] {
            hash ^= value.unwrap_or(i64::MIN) as u64;
            hash = hash.wrapping_mul(FNV_PRIME);
        }
    }
    hash
}

fn page_anchors(list: &DisplayList) -> Vec<String> {
    let mut occurrences: HashMap<String, usize> = HashMap::new();
    list.pages
        .iter()
        .map(|page| {
            let semantic = visit_primitives(page)
                .find_map(|(_, primitive)| primitive_owner(primitive))
                .unwrap_or_else(|| format!("empty:{}", page.page_index));
            let section = page.section_id.as_deref().unwrap_or("");
            let raw = format!("{section}|{semantic}");
            let occurrence = occurrences.entry(raw.clone()).or_default();
            let anchor = format!("{raw}#{}", *occurrence);
            *occurrence += 1;
            anchor
        })
        .collect()
}

fn primitive_ids(page: &DisplayPage, page_id: u64) -> Vec<u64> {
    let mut occurrences: HashMap<String, usize> = HashMap::new();
    let mut used = HashSet::new();
    visit_primitives(page)
        .map(|(region, primitive)| {
            let kind = primitive_kind(primitive);
            let owner = primitive_owner(primitive).unwrap_or_else(|| format!("page:{page_id}"));
            let raw = format!("{region}|{kind}|{owner}");
            let occurrence = occurrences.entry(raw.clone()).or_default();
            let key = format!("{raw}|{}", *occurrence);
            *occurrence += 1;
            let mut id = hash_bytes(key.as_bytes());
            let mut salt = 0_u64;
            while id == 0 || !used.insert(id) {
                salt = salt.wrapping_add(1);
                id = hash_bytes(format!("{key}|collision:{salt}").as_bytes());
            }
            id
        })
        .collect()
}

fn primitive_positions(page: &DisplayPage) -> Vec<PrimitivePositionSnapshot> {
    visit_primitives(page)
        .map(|(_, primitive)| {
            let attrs = primitive_attrs(primitive);
            PrimitivePositionSnapshot {
                doc_start: attrs.doc_start,
                doc_end: attrs.doc_end,
                fragment_doc_start: attrs.fragment_doc_start,
                fragment_doc_end: attrs.fragment_doc_end,
                inline_widget_pos: attrs.inline_sdt_widget.as_ref().map(|widget| widget.pos),
            }
        })
        .collect()
}

fn note_anchors(page: &DisplayPage) -> Result<Vec<NoteAnchorSnapshot>, String> {
    let mut anchors = Vec::new();
    for (area, region) in page.note_areas.iter().enumerate() {
        for (note, entry) in region.notes.iter().enumerate() {
            anchors.push(NoteAnchorSnapshot {
                area: checked_u32(area, "note area index")?,
                note: checked_u32(note, "note index")?,
                start: entry.anchor_doc_start,
                end: entry.anchor_doc_end,
            });
        }
    }
    Ok(anchors)
}

/// The anchors that changed between two snapshots of the same note regions,
/// with their new values; None when the regions hold different notes.
fn changed_note_anchors(
    previous: &[NoteAnchorSnapshot],
    next: &[NoteAnchorSnapshot],
) -> Option<Vec<NoteAnchorSnapshot>> {
    if previous.len() != next.len()
        || previous
            .iter()
            .zip(next)
            .any(|(before, after)| (before.area, before.note) != (after.area, after.note))
    {
        return None;
    }
    Some(
        previous
            .iter()
            .zip(next)
            .filter(|(before, after)| before != after)
            .map(|(_, after)| *after)
            .collect(),
    )
}

fn position_patches(previous: &FramePageSnapshot, next: &FramePageSnapshot) -> Vec<PositionPatch> {
    previous
        .positions
        .iter()
        .zip(&next.positions)
        .zip(next.primitive_ids.iter())
        .filter_map(|((previous, next), primitive_id)| {
            let before = [
                previous.doc_start,
                previous.doc_end,
                previous.fragment_doc_start,
                previous.fragment_doc_end,
                previous.inline_widget_pos,
            ];
            let after = [
                next.doc_start,
                next.doc_end,
                next.fragment_doc_start,
                next.fragment_doc_end,
                next.inline_widget_pos,
            ];
            let mut changed_mask = 0;
            let mut present_mask = 0;
            for (index, field) in POSITION_FIELDS.iter().enumerate() {
                if before[index] != after[index] {
                    changed_mask |= field;
                    if after[index].is_some() {
                        present_mask |= field;
                    }
                }
            }
            (changed_mask != 0).then_some(PositionPatch {
                primitive_id: *primitive_id,
                changed_mask,
                present_mask,
                values: after,
            })
        })
        .collect()
}

fn position_shift_runs(
    previous: &[PrimitivePositionSnapshot],
    next: &[PrimitivePositionSnapshot],
) -> Option<Vec<PositionShiftRun>> {
    if previous.len() != next.len() {
        return None;
    }
    let mut runs: Vec<PositionShiftRun> = Vec::new();
    for (index, (previous, next)) in previous.iter().zip(next).enumerate() {
        let before = [
            previous.doc_start,
            previous.doc_end,
            previous.fragment_doc_start,
            previous.fragment_doc_end,
            previous.inline_widget_pos,
        ];
        let after = [
            next.doc_start,
            next.doc_end,
            next.fragment_doc_start,
            next.fragment_doc_end,
            next.inline_widget_pos,
        ];
        let mut changed_mask = 0;
        let mut present_mask = 0;
        let mut common_delta = None;
        for (field_index, field) in POSITION_FIELDS.iter().enumerate() {
            if after[field_index].is_some() {
                present_mask |= field;
            }
            if before[field_index] == after[field_index] {
                continue;
            }
            let (Some(before), Some(after)) = (before[field_index], after[field_index]) else {
                return None;
            };
            let delta = after.checked_sub(before)?;
            if common_delta.is_some_and(|common| common != delta) {
                return None;
            }
            common_delta = Some(delta);
            changed_mask |= field;
        }
        let index = checked_u32(index, "position shift primitive index").ok()?;
        let last = runs
            .last_mut()
            .filter(|last| last.start.checked_add(last.count) == Some(index));
        let Some(delta) = common_delta else {
            // A primitive without positions changes nothing inside a
            // present-only run, so it does not have to end one.
            if present_mask == 0
                && let Some(last) = last
                && last.changed_mask & POSITION_PRESENT_ONLY != 0
            {
                last.count = last.count.checked_add(1)?;
            }
            continue;
        };
        if delta == 0 {
            return None;
        }
        let changed_mask = if changed_mask == present_mask {
            changed_mask | POSITION_PRESENT_ONLY
        } else {
            changed_mask
        };
        match last {
            Some(last)
                if last.delta == delta
                    && last.changed_mask & POSITION_PRESENT_ONLY != 0
                    && changed_mask & POSITION_PRESENT_ONLY != 0 =>
            {
                last.count = last.count.checked_add(1)?;
                last.changed_mask |= changed_mask;
            }
            Some(last) if last.delta == delta && last.changed_mask == changed_mask => {
                last.count = last.count.checked_add(1)?;
            }
            _ => runs.push(PositionShiftRun {
                start: index,
                count: 1,
                changed_mask,
                delta,
            }),
        }
    }
    (!runs.is_empty()).then_some(runs)
}

fn visit_primitives(page: &DisplayPage) -> impl Iterator<Item = (&'static str, &Primitive)> {
    let body = page.primitives.iter().map(|primitive| ("body", primitive));
    let notes = page.note_areas.iter().flat_map(|area| {
        area.separator_primitives
            .iter()
            .map(|primitive| ("note-separator", primitive))
            .chain(area.primitives.iter().map(|primitive| ("note", primitive)))
    });
    let header = page.header.iter().flat_map(|region| {
        region
            .primitives
            .iter()
            .map(|primitive| ("header", primitive))
    });
    let footer = page.footer.iter().flat_map(|region| {
        region
            .primitives
            .iter()
            .map(|primitive| ("footer", primitive))
    });
    body.chain(notes).chain(header).chain(footer)
}

fn primitive_kind(primitive: &Primitive) -> &'static str {
    match primitive {
        Primitive::Text(_) => "text",
        Primitive::GlyphRun(_) => "glyphRun",
        Primitive::Rect(_) => "rect",
        Primitive::Line(_) => "line",
        Primitive::Image(_) => "image",
        Primitive::Shape(_) => "shape",
        Primitive::Decoration(_) => "decoration",
    }
}

fn primitive_attrs(primitive: &Primitive) -> &DocAttrs {
    match primitive {
        Primitive::Text(value) => &value.attrs,
        Primitive::GlyphRun(value) => &value.attrs,
        Primitive::Rect(value) => &value.attrs,
        Primitive::Line(value) => &value.attrs,
        Primitive::Image(value) => &value.attrs,
        Primitive::Shape(value) => &value.attrs,
        Primitive::Decoration(value) => &value.attrs,
    }
}

fn primitive_owner(primitive: &Primitive) -> Option<String> {
    let attrs = primitive_attrs(primitive);
    attrs
        .para_id
        .as_ref()
        .map(|value| format!("para:{value}"))
        .or_else(|| {
            attrs
                .block_key
                .as_ref()
                .map(|value| format!("block:{value}"))
        })
        .or_else(|| {
            attrs
                .block_id
                .as_ref()
                .map(|value| format!("block:{value}"))
        })
        .or_else(|| {
            attrs
                .cell
                .as_ref()
                .and_then(|cell| cell.cell_id.as_ref())
                .map(|value| format!("cell:{value}"))
        })
}

#[cfg(test)]
fn collect_strings(value: &Value, strings: &mut std::collections::BTreeSet<String>) {
    match value {
        Value::String(value) => {
            strings.insert(value.clone());
        }
        Value::Array(values) => {
            for value in values {
                collect_strings(value, strings);
            }
        }
        Value::Object(values) => {
            for (key, value) in values {
                strings.insert(key.clone());
                if key == "glyphs" && compact_glyphs(value).is_some() {
                    continue;
                }
                collect_strings(value, strings);
            }
        }
        Value::Null | Value::Bool(_) | Value::Number(_) => {}
    }
}

// Typed value opcodes. Array/object payloads are `[byte_len, count, ...]`.
const VALUE_NULL: u8 = 0;
const VALUE_FALSE: u8 = 1;
const VALUE_TRUE: u8 = 2;
const VALUE_I64: u8 = 3;
const VALUE_U64: u8 = 4;
const VALUE_F64: u8 = 5;
const VALUE_STRING: u8 = 6;
const VALUE_ARRAY: u8 = 7;
const VALUE_OBJECT: u8 = 8;
const VALUE_GLYPH_ARRAY: u8 = 9;

const GLYPH_LOGICAL_ORDER: u8 = 1 << 0;
const GLYPH_BIDI_LEVEL: u8 = 1 << 1;

#[cfg(test)]
fn string_id(ids: &HashMap<&str, u32>, value: &str) -> Result<u32, String> {
    ids.get(value)
        .copied()
        .ok_or_else(|| "FrameDelta string table missed a value".to_owned())
}

#[cfg(test)]
fn encode_value(
    value: &Value,
    string_ids: &HashMap<&str, u32>,
    out: &mut Vec<u8>,
    parent_key: Option<&str>,
) -> Result<(), String> {
    if parent_key == Some("glyphs")
        && let Some(glyphs) = compact_glyphs(value)
    {
        return encode_glyph_array(glyphs, out);
    }
    match value {
        Value::Null => out.push(VALUE_NULL),
        Value::Bool(false) => out.push(VALUE_FALSE),
        Value::Bool(true) => out.push(VALUE_TRUE),
        Value::Number(value) => {
            if let Some(value) = value.as_i64() {
                out.push(VALUE_I64);
                write_i64(out, value);
            } else if let Some(value) = value.as_u64() {
                out.push(VALUE_U64);
                write_u64(out, value);
            } else if let Some(value) = value.as_f64() {
                out.push(VALUE_F64);
                write_f64(out, value);
            } else {
                return Err("FrameDelta contains an unrepresentable number".to_owned());
            }
        }
        Value::String(value) => {
            out.push(VALUE_STRING);
            write_u32(out, string_id(string_ids, value)?);
        }
        Value::Array(values) => {
            out.push(VALUE_ARRAY);
            let length_at = out.len();
            write_u32(out, 0);
            write_u32(out, checked_u32(values.len(), "array element count")?);
            let payload_at = out.len();
            for value in values {
                encode_value(value, string_ids, out, None)?;
            }
            let payload_len = out.len() - payload_at;
            patch_u32(
                out,
                length_at,
                checked_u32(payload_len, "array payload length")?,
            );
        }
        Value::Object(values) => {
            out.push(VALUE_OBJECT);
            let length_at = out.len();
            write_u32(out, 0);
            write_u32(out, checked_u32(values.len(), "object field count")?);
            let payload_at = out.len();
            for (key, value) in values {
                write_u32(out, string_id(string_ids, key)?);
                encode_value(value, string_ids, out, Some(key))?;
            }
            let payload_len = out.len() - payload_at;
            patch_u32(
                out,
                length_at,
                checked_u32(payload_len, "object payload length")?,
            );
        }
    }
    Ok(())
}

#[cfg(test)]
fn compact_glyphs(value: &Value) -> Option<&[Value]> {
    let Value::Array(glyphs) = value else {
        return None;
    };
    glyphs
        .iter()
        .all(|glyph| {
            let Some(fields) = glyph.as_object() else {
                return false;
            };
            fields.len() >= 5
                && fields.len() <= 7
                && fields.get("id").and_then(Value::as_u64).is_some()
                && fields.get("x").and_then(Value::as_f64).is_some()
                && fields.get("y").and_then(Value::as_f64).is_some()
                && fields.get("cluster").and_then(Value::as_u64).is_some()
                && fields.get("advance").and_then(Value::as_f64).is_some()
                && fields.keys().all(|key| {
                    matches!(
                        key.as_str(),
                        "id" | "x" | "y" | "cluster" | "advance" | "logicalOrder" | "bidiLevel"
                    )
                })
                && fields
                    .get("logicalOrder")
                    .is_none_or(|value| value.as_u64().is_some())
                && fields
                    .get("bidiLevel")
                    .is_none_or(|value| value.as_u64().is_some_and(|value| value <= u8::MAX as u64))
        })
        .then_some(glyphs)
}

#[cfg(test)]
fn encode_glyph_array(glyphs: &[Value], out: &mut Vec<u8>) -> Result<(), String> {
    out.push(VALUE_GLYPH_ARRAY);
    let length_at = out.len();
    write_u32(out, 0);
    write_u32(out, checked_u32(glyphs.len(), "glyph array count")?);
    let payload_at = out.len();
    for glyph in glyphs {
        let fields = glyph
            .as_object()
            .expect("compact glyph validation guarantees an object");
        write_u32(
            out,
            u32::try_from(fields["id"].as_u64().expect("validated glyph id"))
                .map_err(|_| "glyph id exceeds u32".to_owned())?,
        );
        write_f64(out, fields["x"].as_f64().expect("validated glyph x"));
        write_f64(out, fields["y"].as_f64().expect("validated glyph y"));
        write_u32(
            out,
            u32::try_from(fields["cluster"].as_u64().expect("validated glyph cluster"))
                .map_err(|_| "glyph cluster exceeds u32".to_owned())?,
        );
        write_f64(
            out,
            fields["advance"].as_f64().expect("validated glyph advance"),
        );
        let flags = if fields.contains_key("logicalOrder") {
            GLYPH_LOGICAL_ORDER
        } else {
            0
        } | if fields.contains_key("bidiLevel") {
            GLYPH_BIDI_LEVEL
        } else {
            0
        };
        out.push(flags);
        if let Some(value) = fields.get("logicalOrder") {
            write_u64(out, value.as_u64().expect("validated glyph logical order"));
        }
        if let Some(value) = fields.get("bidiLevel") {
            out.push(value.as_u64().expect("validated glyph bidi level") as u8);
        }
    }
    let payload_len = out.len() - payload_at;
    patch_u32(
        out,
        length_at,
        checked_u32(payload_len, "glyph array payload length")?,
    );
    Ok(())
}

/// A frame's string table, filled in first-use order while pages are encoded.
/// The strings come from the document, so the table keeps std's randomized
/// hashing.
#[derive(Default)]
pub(crate) struct StringTable {
    ids: HashMap<String, (u32, u64)>,
    strings: Vec<String>,
}

impl StringTable {
    /// `value`'s id and content hash.
    pub(crate) fn intern(&mut self, value: &str) -> Result<(u32, u64), String> {
        if let Some(entry) = self.ids.get(value) {
            return Ok(*entry);
        }
        let id = u32::try_from(self.ids.len())
            .map_err(|_| "FrameDelta string table exceeds u32".to_owned())?;
        let entry = (id, string_hash(value));
        self.ids.insert(value.to_owned(), entry);
        self.strings.push(value.to_owned());
        Ok(entry)
    }

    fn mark(&self) -> usize {
        self.strings.len()
    }

    /// Forgets the strings interned since `mark`.
    fn rollback(&mut self, mark: usize) {
        for value in self.strings.drain(mark..) {
            self.ids.remove(&value);
        }
    }

    /// The strings indexed by their ids.
    pub(crate) fn into_strings(self) -> Vec<String> {
        self.strings
    }
}

#[cfg(test)]
fn hash_page_value(value: &Value) -> u64 {
    fn visit(value: &Value, root: bool, hash: &mut u64) {
        match value {
            Value::Null => hash_write(hash, &[VALUE_NULL]),
            Value::Bool(false) => hash_write(hash, &[VALUE_FALSE]),
            Value::Bool(true) => hash_write(hash, &[VALUE_TRUE]),
            Value::Number(value) => {
                if let Some(value) = value.as_i64() {
                    hash_write(hash, &[VALUE_I64]);
                    hash_write(hash, &value.to_le_bytes());
                } else if let Some(value) = value.as_u64() {
                    hash_write(hash, &[VALUE_U64]);
                    hash_write(hash, &value.to_le_bytes());
                } else if let Some(value) = value.as_f64() {
                    hash_write(hash, &[VALUE_F64]);
                    hash_write(hash, &value.to_bits().to_le_bytes());
                }
            }
            Value::String(value) => {
                hash_write(hash, &[VALUE_STRING]);
                hash_write(hash, value.as_bytes());
            }
            Value::Array(values) => {
                hash_write(hash, &[VALUE_ARRAY]);
                hash_write(hash, &(values.len() as u64).to_le_bytes());
                for value in values {
                    visit(value, false, hash);
                }
            }
            Value::Object(values) => {
                hash_write(hash, &[VALUE_OBJECT]);
                let mut fields: Vec<_> = values
                    .iter()
                    .filter(|(key, _)| !(root && key.as_str() == "pageIndex"))
                    .collect();
                fields.sort_unstable_by(|a, b| a.0.cmp(b.0));
                for (key, value) in fields {
                    hash_write(hash, key.as_bytes());
                    visit(value, false, hash);
                }
            }
        }
    }
    let mut hash = FNV_OFFSET;
    visit(value, true, &mut hash);
    hash
}

/// Paint/a11y structure hash with only absolute document-position metadata
/// removed. Equality means the browser can retain the raster and receive a
/// compact stable-primitive position patch for its mirror/overlays.
#[cfg(test)]
fn hash_visual_page_value(value: &Value) -> u64 {
    fn visit(value: &Value, parent_key: Option<&str>, root: bool, hash: &mut u64) {
        match value {
            Value::Null => hash_write(hash, &[VALUE_NULL]),
            Value::Bool(false) => hash_write(hash, &[VALUE_FALSE]),
            Value::Bool(true) => hash_write(hash, &[VALUE_TRUE]),
            Value::Number(value) => {
                if let Some(value) = value.as_i64() {
                    hash_write(hash, &[VALUE_I64]);
                    hash_write(hash, &value.to_le_bytes());
                } else if let Some(value) = value.as_u64() {
                    hash_write(hash, &[VALUE_U64]);
                    hash_write(hash, &value.to_le_bytes());
                } else if let Some(value) = value.as_f64() {
                    hash_write(hash, &[VALUE_F64]);
                    hash_write(hash, &value.to_bits().to_le_bytes());
                }
            }
            Value::String(value) => {
                hash_write(hash, &[VALUE_STRING]);
                hash_write(hash, value.as_bytes());
            }
            Value::Array(values) => {
                hash_write(hash, &[VALUE_ARRAY]);
                hash_write(hash, &(values.len() as u64).to_le_bytes());
                for value in values {
                    visit(value, parent_key, false, hash);
                }
            }
            Value::Object(values) => {
                hash_write(hash, &[VALUE_OBJECT]);
                let mut fields: Vec<_> = values
                    .iter()
                    .filter(|(key, _)| {
                        !(root && key.as_str() == "pageIndex")
                            && !matches!(
                                key.as_str(),
                                "docStart" | "docEnd" | "fragmentDocStart" | "fragmentDocEnd"
                            )
                            && !(parent_key == Some("inlineSdtWidget") && key.as_str() == "pos")
                    })
                    .collect();
                fields.sort_unstable_by(|a, b| a.0.cmp(b.0));
                for (key, value) in fields {
                    hash_write(hash, key.as_bytes());
                    visit(value, Some(key), false, hash);
                }
            }
        }
    }
    let mut hash = FNV_OFFSET;
    visit(value, None, true, &mut hash);
    hash
}

/// One step of the page fingerprints: the SplitMix64 finalizer over
/// `state ^ word`. For a fixed `word` it is a bijection of `state`, so changing
/// any single word changes the result, and its avalanche keeps a change in one
/// word from being cancelled by a change in the next.
fn mix(state: u64, word: u64) -> u64 {
    let mut state = state ^ word;
    state = (state ^ (state >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    state = (state ^ (state >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    state ^ (state >> 31)
}

/// The nonzero fingerprint salt of a shift range, distinct per frame so that
/// shifts which cancel out in position never cancel out in fingerprint.
fn shift_range_salt(delta: i64, frame_epoch: u64) -> u64 {
    let salt = mix(
        mix(mix(FNV_OFFSET, 0x5348_4946_545f_5247), delta as u64),
        frame_epoch,
    );
    if salt == 0 { 1 } else { salt }
}

fn string_hash(value: &str) -> u64 {
    let bytes = value.as_bytes();
    let (chunks, remainder) = bytes.as_chunks::<8>();
    let mut hash = mix(FNV_OFFSET, bytes.len() as u64);
    for chunk in chunks {
        hash = mix(hash, u64::from_le_bytes(*chunk));
    }
    let mut tail = [0; 8];
    tail[..remainder.len()].copy_from_slice(remainder);
    mix(hash, u64::from_le_bytes(tail))
}

fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut hash = FNV_OFFSET;
    hash_write(&mut hash, bytes);
    hash
}

fn hash_write(hash: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(FNV_PRIME);
    }
}

fn checked_u32(value: usize, label: &str) -> Result<u32, String> {
    u32::try_from(value).map_err(|_| format!("FrameDelta {label} exceeds u32"))
}

fn align(out: &mut Vec<u8>, alignment: usize) {
    let padding = (alignment - out.len() % alignment) % alignment;
    out.resize(out.len() + padding, 0);
}

fn write_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn write_u16(out: &mut Vec<u8>, value: u16) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn write_u64(out: &mut Vec<u8>, value: u64) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn write_i64(out: &mut Vec<u8>, value: i64) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn write_f64(out: &mut Vec<u8>, value: f64) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn patch_u16(out: &mut [u8], offset: usize, value: u16) {
    out[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn patch_u32(out: &mut [u8], offset: usize, value: u32) {
    out[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn patch_u64(out: &mut [u8], offset: usize, value: u64) {
    out[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
}

/// A decoded page and the identity tokens retained by the test host.
#[cfg(test)]
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct TestFramePage {
    pub(crate) page: DisplayPage,
    pub(crate) fingerprint: u64,
    pub(crate) primitive_ids: Vec<u64>,
}

/// Visits mutable primitives in their encoded identity order.
#[cfg(test)]
fn test_page_primitives_mut(page: &mut DisplayPage) -> impl Iterator<Item = &mut Primitive> {
    let body = page.primitives.iter_mut();
    let notes = page.note_areas.iter_mut().flat_map(|area| {
        area.separator_primitives
            .iter_mut()
            .chain(&mut area.primitives)
    });
    let header = page
        .header
        .iter_mut()
        .flat_map(|region| &mut region.primitives);
    let footer = page
        .footer
        .iter_mut()
        .flat_map(|region| &mut region.primitives);
    body.chain(notes).chain(header).chain(footer)
}

/// Returns the mutable position attributes of a test primitive.
#[cfg(test)]
fn test_primitive_attrs_mut(primitive: &mut Primitive) -> &mut DocAttrs {
    match primitive {
        Primitive::Text(value) => &mut value.attrs,
        Primitive::GlyphRun(value) => &mut value.attrs,
        Primitive::Rect(value) => &mut value.attrs,
        Primitive::Line(value) => &mut value.attrs,
        Primitive::Image(value) => &mut value.attrs,
        Primitive::Shape(value) => &mut value.attrs,
        Primitive::Decoration(value) => &mut value.attrs,
    }
}

/// Applies a masked position delta as the test host would.
#[cfg(test)]
fn shift_test_primitive(primitive: &mut Primitive, mask: u8, delta: i64) {
    let attrs = test_primitive_attrs_mut(primitive);
    for (index, value) in [
        &mut attrs.doc_start,
        &mut attrs.doc_end,
        &mut attrs.fragment_doc_start,
        &mut attrs.fragment_doc_end,
    ]
    .into_iter()
    .enumerate()
    {
        if mask & POSITION_FIELDS[index] != 0 {
            if let Some(value) = value {
                *value = value.checked_add(delta).unwrap();
            } else {
                assert_ne!(mask & POSITION_PRESENT_ONLY, 0);
            }
        }
    }
    if mask & POSITION_INLINE_WIDGET != 0 {
        if let Some(widget) = &mut attrs.inline_sdt_widget {
            widget.pos = widget.pos.checked_add(delta).unwrap();
        } else {
            assert_ne!(mask & POSITION_PRESENT_ONLY, 0);
        }
    }
}

/// Applies a frame to the test host's retained pages and fingerprints.
#[cfg(test)]
pub(crate) fn apply_placeholder_test_frame(
    bytes: &[u8],
    retained: &mut HashMap<u64, TestFramePage>,
) -> DisplayList {
    let u32_at = |offset| u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap());
    let u64_at = |offset| u64::from_le_bytes(bytes[offset..offset + 8].try_into().unwrap());
    let i64_at = |offset| i64::from_le_bytes(bytes[offset..offset + 8].try_into().unwrap());
    if u32_at(12) & FRAME_FLAG_FULL != 0 {
        retained.clear();
    }
    let mut at = u32_at(60) as usize;
    let string_count = u32_at(at);
    at += 4;
    let mut strings = Vec::new();
    for _ in 0..string_count {
        let length = u32_at(at) as usize;
        at += 4;
        strings.push(String::from_utf8(bytes[at..at + length].to_vec()).unwrap());
        at += length;
    }
    for index in 0..u32_at(52) as usize {
        let record = FRAME_HEADER_LEN + index * PAGE_OP_LEN;
        let page_id = u64_at(record + 8);
        let page_index = u32_at(record + 4);
        match bytes[record] {
            PAGE_OP_UPSERT => {
                let mut cursor = u32_at(record + 32) as usize;
                let value = tests::decode_typed(bytes, &mut cursor, &strings);
                assert_eq!(
                    cursor,
                    u32_at(record + 32) as usize + u32_at(record + 36) as usize
                );
                let ids = u32_at(record + 28) as usize;
                let primitive_ids = (0..u32_at(record + 24) as usize)
                    .map(|index| u64_at(ids + index * 8))
                    .collect();
                retained.insert(
                    page_id,
                    TestFramePage {
                        page: serde_json::from_value(value).unwrap(),
                        fingerprint: u64_at(record + 16),
                        primitive_ids,
                    },
                );
            }
            PAGE_OP_REMOVE => {
                retained.remove(&page_id).unwrap();
            }
            PAGE_OP_MOVE => {
                let retained = retained.get_mut(&page_id).unwrap();
                retained.page.page_index = u64::from(page_index);
                retained.fingerprint = u64_at(record + 16);
            }
            PAGE_OP_PATCH_POSITIONS => {
                let retained = retained.get_mut(&page_id).unwrap();
                let mut cursor = u32_at(record + 32) as usize;
                let count = u32_at(cursor);
                assert_eq!(count, u32_at(record + 24));
                cursor += 8;
                for _ in 0..count {
                    let primitive_id = u64_at(cursor);
                    let changed = bytes[cursor + 8];
                    let present = bytes[cursor + 9];
                    cursor += 12;
                    let mut values = [None; 5];
                    for (index, field) in POSITION_FIELDS.iter().enumerate() {
                        if present & field != 0 {
                            values[index] = Some(i64_at(cursor));
                            cursor += 8;
                        }
                    }
                    let index = retained
                        .primitive_ids
                        .iter()
                        .position(|&id| id == primitive_id)
                        .unwrap();
                    let primitive = test_page_primitives_mut(&mut retained.page)
                        .nth(index)
                        .unwrap();
                    let attrs = test_primitive_attrs_mut(primitive);
                    for (index, value) in [
                        &mut attrs.doc_start,
                        &mut attrs.doc_end,
                        &mut attrs.fragment_doc_start,
                        &mut attrs.fragment_doc_end,
                    ]
                    .into_iter()
                    .enumerate()
                    {
                        if changed & POSITION_FIELDS[index] != 0 {
                            *value = values[index];
                        }
                    }
                    if changed & POSITION_INLINE_WIDGET != 0 {
                        if let Some(pos) = values[4] {
                            attrs.inline_sdt_widget.as_mut().unwrap().pos = pos;
                        } else {
                            attrs.inline_sdt_widget = None;
                        }
                    }
                }
                assert_eq!(
                    cursor,
                    u32_at(record + 32) as usize + u32_at(record + 36) as usize
                );
                retained.page.page_index = u64::from(page_index);
                retained.fingerprint = u64_at(record + 16);
            }
            PAGE_OP_SHIFT_POSITIONS => {
                let retained = retained.get_mut(&page_id).unwrap();
                let mut cursor = u32_at(record + 32) as usize;
                let count = u32_at(cursor);
                let flags = u32_at(cursor + 4);
                assert_eq!(count, u32_at(record + 24));
                cursor += 8;
                if flags & SHIFT_SPAN_PRESENT != 0 {
                    let delta = i64_at(cursor);
                    cursor += 8;
                    let span = retained.page.position_span.as_mut().unwrap();
                    for value in span {
                        *value = value.checked_add(delta).unwrap();
                    }
                }
                for _ in 0..count {
                    let start = u32_at(cursor) as usize;
                    let count = u32_at(cursor + 4) as usize;
                    let mask = bytes[cursor + 8];
                    let delta = i64_at(cursor + 16);
                    cursor += 24;
                    for primitive in test_page_primitives_mut(&mut retained.page)
                        .skip(start)
                        .take(count)
                    {
                        shift_test_primitive(primitive, mask, delta);
                    }
                }
                let anchors = u32_at(record + 40);
                if anchors != 0 {
                    assert_eq!(u32_at(cursor), anchors);
                    cursor += 8;
                    for _ in 0..anchors {
                        let area = u32_at(cursor) as usize;
                        let note = u32_at(cursor + 4) as usize;
                        let start = i64_at(cursor + 8);
                        let end = i64_at(cursor + 16);
                        cursor += 24;
                        let note = &mut retained.page.note_areas[area].notes[note];
                        note.anchor_doc_start = (start != i64::MIN).then_some(start);
                        note.anchor_doc_end = (end != i64::MIN).then_some(end);
                    }
                }
                assert_eq!(
                    cursor,
                    u32_at(record + 32) as usize + u32_at(record + 36) as usize
                );
                retained.page.page_index = u64::from(page_index);
                retained.fingerprint = u64_at(record + 16);
            }
            PAGE_OP_SHIFT_RANGE => {
                let count = u32_at(record + 24);
                let salt = u64_at(record + 16);
                let delta = i64_at(record + 32);
                assert!(count >= 1 && page_index + count <= u32_at(48));
                assert_ne!(page_id, 0);
                assert_ne!(salt, 0);
                assert!(delta != 0 && (-MAX_SAFE_INTEGER..=MAX_SAFE_INTEGER).contains(&delta));
                assert_eq!(bytes[record + 1..record + 4], [0; 3]);
                assert_eq!(u32_at(record + 28), 0);
                assert_eq!(u64_at(record + 40), 0);
                assert_eq!(retained[&page_id].page.page_index, u64::from(page_index));
                let mut shifted = 0;
                for retained in retained.values_mut().filter(|retained| {
                    (u64::from(page_index)..u64::from(page_index + count))
                        .contains(&retained.page.page_index)
                }) {
                    retained.fingerprint ^= salt;
                    if let Some(span) = &mut retained.page.position_span {
                        for value in span {
                            *value = value.checked_add(delta).unwrap();
                        }
                    }
                    for primitive in &mut retained.page.primitives {
                        shift_test_primitive(primitive, POSITION_PRESENT_ONLY | 0x1f, delta);
                    }
                    shifted += 1;
                }
                assert_eq!(shifted, count);
            }
            opcode => panic!("unexpected test frame opcode {opcode}"),
        }
    }
    let mut pages: Vec<_> = retained
        .values()
        .map(|retained| retained.page.clone())
        .collect();
    pages.sort_by_key(|page| page.page_index);
    assert_eq!(pages.len(), u32_at(48) as usize);
    DisplayList {
        contract_version: (u32_at(72) != 0).then(|| u32_at(72)),
        pages,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::DisplayList;
    use std::collections::BTreeSet;

    fn list(text: &str) -> DisplayList {
        list_pages(&[("P1", text)])
    }

    fn list_pages(paragraphs: &[(&str, &str)]) -> DisplayList {
        let pages: Vec<_> = paragraphs
            .iter()
            .enumerate()
            .map(|(page_index, (para_id, text))| {
                serde_json::json!({
                    "pageIndex": page_index,
                    "width": 816,
                    "height": 1056,
                    "primitives": [{
                        "kind": "text",
                        "text": text,
                        "x": 96,
                        "baselineY": 120,
                        "width": 40,
                        "font": "16px serif",
                        "color": "#000000",
                        "docStart": 1,
                        "docEnd": 5,
                        "blockId": 7,
                        "paraId": para_id
                    }]
                })
            })
            .collect();
        serde_json::from_value(serde_json::json!({
            "contractVersion": 1,
            "pages": pages
        }))
        .unwrap()
    }

    fn list_at_position(doc_start: i64) -> DisplayList {
        let mut value = serde_json::to_value(list("hello")).unwrap();
        let primitive = &mut value["pages"][0]["primitives"][0];
        primitive["docStart"] = serde_json::json!(doc_start);
        primitive["docEnd"] = serde_json::json!(doc_start + 4);
        primitive["fragmentDocStart"] = serde_json::json!(doc_start - 1);
        primitive["fragmentDocEnd"] = serde_json::json!(doc_start + 5);
        serde_json::from_value(value).unwrap()
    }

    fn u32_at(bytes: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
    }

    fn u64_at(bytes: &[u8], offset: usize) -> u64 {
        u64::from_le_bytes(bytes[offset..offset + 8].try_into().unwrap())
    }

    /// Test-only decoder over the typed value stream, mirroring the browser
    /// decoder's structure so typed emission can be checked against the
    /// `Value` reference implementation.
    pub(super) fn decode_typed(bytes: &[u8], cursor: &mut usize, strings: &[String]) -> Value {
        let tag = bytes[*cursor];
        *cursor += 1;
        match tag {
            VALUE_NULL => Value::Null,
            VALUE_FALSE => Value::Bool(false),
            VALUE_TRUE => Value::Bool(true),
            VALUE_I64 => {
                let value = i64::from_le_bytes(bytes[*cursor..*cursor + 8].try_into().unwrap());
                *cursor += 8;
                Value::from(value)
            }
            VALUE_U64 => {
                let value = u64_at(bytes, *cursor);
                *cursor += 8;
                Value::from(value)
            }
            VALUE_F64 => {
                let value = f64::from_le_bytes(bytes[*cursor..*cursor + 8].try_into().unwrap());
                *cursor += 8;
                Value::from(value)
            }
            VALUE_STRING => {
                let id = u32_at(bytes, *cursor) as usize;
                *cursor += 4;
                Value::from(strings[id].as_str())
            }
            VALUE_ARRAY => {
                let length = u32_at(bytes, *cursor) as usize;
                let count = u32_at(bytes, *cursor + 4) as usize;
                *cursor += 8;
                let end = *cursor + length;
                let values = (0..count)
                    .map(|_| decode_typed(bytes, cursor, strings))
                    .collect();
                assert_eq!(*cursor, end, "array byte length/count mismatch");
                Value::Array(values)
            }
            VALUE_OBJECT => {
                let length = u32_at(bytes, *cursor) as usize;
                let count = u32_at(bytes, *cursor + 4) as usize;
                *cursor += 8;
                let end = *cursor + length;
                let mut fields = serde_json::Map::new();
                for _ in 0..count {
                    let key = strings[u32_at(bytes, *cursor) as usize].clone();
                    *cursor += 4;
                    let value = decode_typed(bytes, cursor, strings);
                    assert!(fields.insert(key, value).is_none(), "duplicate object key");
                }
                assert_eq!(*cursor, end, "object byte length/count mismatch");
                Value::Object(fields)
            }
            VALUE_GLYPH_ARRAY => {
                let length = u32_at(bytes, *cursor) as usize;
                let count = u32_at(bytes, *cursor + 4) as usize;
                *cursor += 8;
                let end = *cursor + length;
                let mut glyphs = Vec::new();
                for _ in 0..count {
                    let mut glyph = serde_json::Map::new();
                    glyph.insert("id".into(), Value::from(u32_at(bytes, *cursor) as u64));
                    let x =
                        f64::from_le_bytes(bytes[*cursor + 4..*cursor + 12].try_into().unwrap());
                    let y =
                        f64::from_le_bytes(bytes[*cursor + 12..*cursor + 20].try_into().unwrap());
                    glyph.insert("x".into(), Value::from(x));
                    glyph.insert("y".into(), Value::from(y));
                    glyph.insert(
                        "cluster".into(),
                        Value::from(u32_at(bytes, *cursor + 20) as u64),
                    );
                    let advance =
                        f64::from_le_bytes(bytes[*cursor + 24..*cursor + 32].try_into().unwrap());
                    glyph.insert("advance".into(), Value::from(advance));
                    let flags = bytes[*cursor + 32];
                    *cursor += 33;
                    if flags & GLYPH_LOGICAL_ORDER != 0 {
                        glyph.insert("logicalOrder".into(), Value::from(u64_at(bytes, *cursor)));
                        *cursor += 8;
                    }
                    if flags & GLYPH_BIDI_LEVEL != 0 {
                        glyph.insert("bidiLevel".into(), Value::from(bytes[*cursor] as u64));
                        *cursor += 1;
                    }
                    glyphs.push(Value::Object(glyph));
                }
                assert_eq!(*cursor, end, "glyph array byte length/count mismatch");
                Value::Array(glyphs)
            }
            other => panic!("unknown typed value opcode {other}"),
        }
    }

    /// Rewrites compact-eligible glyph arrays into the canonical numeric forms
    /// the wire round-trips (ids as u64, coordinates as f64), so a decoded
    /// stream compares equal to the `Value` reference.
    fn canonicalize_compact_glyphs(value: &mut Value) {
        match value {
            Value::Array(values) => {
                for value in values {
                    canonicalize_compact_glyphs(value);
                }
            }
            Value::Object(fields) => {
                for (key, value) in fields.iter_mut() {
                    if key == "glyphs" && compact_glyphs(value).is_some() {
                        let Value::Array(glyphs) = value else {
                            unreachable!()
                        };
                        for glyph in glyphs {
                            let Value::Object(fields) = glyph else {
                                unreachable!()
                            };
                            for (key, field) in fields.iter_mut() {
                                *field = match key.as_str() {
                                    "id" | "cluster" | "logicalOrder" | "bidiLevel" => {
                                        Value::from(field.as_u64().unwrap())
                                    }
                                    _ => Value::from(field.as_f64().unwrap()),
                                };
                            }
                        }
                    } else {
                        canonicalize_compact_glyphs(value);
                    }
                }
            }
            _ => {}
        }
    }

    fn rich_list() -> DisplayList {
        serde_json::from_value(serde_json::json!({
            "contractVersion": 1,
            "pages": [{
                "pageIndex": 0,
                "width": 816,
                "height": 1056,
                "sectionId": "s1",
                "pageLabel": "1",
                "primitives": [
                    {
                        "kind": "glyphRun",
                        "fontId": 3,
                        "size": 16.0,
                        "color": "#112233",
                        "text": "ab",
                        "glyphs": [
                            {"id": 42, "x": 0.0, "y": 120.0, "cluster": 0, "advance": 8.5},
                            {"id": 7, "x": 8.5, "y": 120.0, "cluster": 1, "advance": 8.0,
                             "logicalOrder": 1, "bidiLevel": 1}
                        ],
                        "docStart": 1,
                        "docEnd": 3,
                        "paraId": "P1",
                        "effects": [{"kind": "glow", "radius": 2.5}],
                        "border": {"style": "single", "width": 0.5}
                    },
                    {
                        "kind": "text",
                        "text": "plain",
                        "x": 96,
                        "baselineY": 160,
                        "width": 40,
                        "font": "16px serif",
                        "color": "#000000",
                        "docStart": 4,
                        "docEnd": 9,
                        "blockId": 7
                    }
                ]
            }]
        }))
        .unwrap()
    }

    /// Shifts only the serialized body position fields and page spans.
    fn shifted_body_list(
        list: &DisplayList,
        range: std::ops::Range<usize>,
        delta: i64,
    ) -> DisplayList {
        let mut next = serde_json::to_value(list).unwrap();
        for page in &mut next["pages"].as_array_mut().unwrap()[range] {
            if let Some(span) = page.get_mut("positionSpan") {
                for value in span.as_array_mut().unwrap() {
                    *value = Value::from(value.as_i64().unwrap() + delta);
                }
            }
            for primitive in page["primitives"].as_array_mut().unwrap() {
                for field in ["docStart", "docEnd", "fragmentDocStart", "fragmentDocEnd"] {
                    if let Some(value) = primitive.get_mut(field) {
                        *value = Value::from(value.as_i64().unwrap() + delta);
                    }
                }
                if let Some(widget) = primitive.get_mut("inlineSdtWidget") {
                    widget["pos"] = Value::from(widget["pos"].as_i64().unwrap() + delta);
                }
            }
        }
        serde_json::from_value(next).unwrap()
    }

    /// Includes positioned body, note, header and footer primitives.
    fn range_shift_built_list() -> DisplayList {
        let mut value = serde_json::to_value(list_pages(&[
            ("P1", "first"),
            ("P2", "second"),
            ("P3", "third"),
        ]))
        .unwrap();
        let primitive = value["pages"][0]["primitives"][0].clone();
        let page = &mut value["pages"][0];
        page["header"] = serde_json::json!({
            "rId": "header", "kind": "header", "y": 0, "height": 20,
            "primitives": [primitive.clone()]
        });
        page["footer"] = serde_json::json!({
            "rId": "footer", "kind": "footer", "y": 1000, "height": 20,
            "primitives": [primitive.clone()]
        });
        page["noteAreas"] = serde_json::json!([{
            "separatorPrimitives": [primitive.clone()], "primitives": [primitive],
            "notes": [{"id": 1, "anchorDocStart": 2, "anchorDocEnd": 3}]
        }]);
        page["primitives"][0]["fragmentDocStart"] = serde_json::json!(0);
        page["primitives"][0]["fragmentDocEnd"] = serde_json::json!(6);
        page["primitives"][0]["inlineSdtWidget"] = serde_json::json!({
            "kind": "checkbox", "groupId": "widget", "pos": 3
        });
        page["primitives"]
            .as_array_mut()
            .unwrap()
            .push(serde_json::json!({
                "kind": "rect", "x": 0, "y": 0, "w": 1, "h": 1,
                "fill": "#000", "fragmentDocEnd": 6
            }));
        serde_json::from_value(value).unwrap()
    }

    /// The general encoders compare against range-shifted snapshots as the host holds them.
    #[test]
    fn general_encoders_fold_deferred_shifts_before_comparing() {
        let before = range_shift_built_list();
        let count = before.pages.len();
        let mut next_id = 0;
        let (full, mut snapshots) =
            encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id).unwrap();
        let mut retained = HashMap::new();
        apply_placeholder_test_frame(&full, &mut retained);
        let shifted = shifted_body_list(&before, 0..count, 1);
        let bytes = encode_frame_delta_changes(
            &shifted,
            &mut snapshots,
            placeholder_epochs(2),
            DisplayChanges {
                rebuilt: &[],
                repositioned: &[],
                shifts: &[PageShiftRun {
                    start: 0,
                    end: count,
                    delta: 1,
                }],
            },
        )
        .unwrap();
        assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), shifted);
        let rebuilt: HashSet<usize> = (0..count).collect();
        let (bytes, _) = encode_frame_delta_incremental(
            &before,
            &snapshots,
            placeholder_epochs(3),
            &mut next_id,
            &rebuilt,
        )
        .unwrap();
        assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), before);
    }

    /// Built pages, placeholders and inert extensions decode identically through both encoders.
    #[test]
    fn uniform_page_ranges_encode_as_one_operation() {
        let built = range_shift_built_list();
        let mut placeholders = built.clone();
        for (index, page) in placeholders.pages.iter_mut().enumerate() {
            *page = placeholder_list(Some([10 + index as i64 * 20, 20 + index as i64 * 20]))
                .pages
                .remove(0);
            page.page_index = index as u64;
        }
        let mut inert = built.clone();
        let last = inert.pages.last_mut().unwrap();
        last.primitives.clear();
        for before in [built, placeholders, inert] {
            let count = before.pages.len();
            let mut next_id = 0;
            let (full, mut snapshots) =
                encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id)
                    .unwrap();
            let mut retained = HashMap::new();
            assert_eq!(apply_placeholder_test_frame(&full, &mut retained), before);
            let mut current = before;
            for (epoch, delta) in [(2, 6), (3, -4), (4, 6)] {
                let previous = snapshots.clone();
                let mut legacy_snapshots = previous.clone();
                for page in &mut legacy_snapshots {
                    page.materialize_positions();
                }
                let mut legacy_retained = retained.clone();
                let after = shifted_body_list(&current, 0..count, delta);
                let bytes = encode_frame_delta_changes(
                    &after,
                    &mut snapshots,
                    placeholder_epochs(epoch),
                    DisplayChanges {
                        rebuilt: &[],
                        repositioned: &[],
                        shifts: &[PageShiftRun {
                            start: 0,
                            end: count,
                            delta,
                        }],
                    },
                )
                .unwrap();
                let record = FRAME_HEADER_LEN;
                assert_eq!(u32_at(&bytes, 12), 0);
                assert_eq!(u64_at(&bytes, 40), epoch - 1);
                assert_eq!(u32_at(&bytes, 48) as usize, count);
                assert_eq!(u32_at(&bytes, 52), 1);
                assert_eq!(bytes[record], PAGE_OP_SHIFT_RANGE);
                assert_eq!(u32_at(&bytes, record + 4), 0);
                assert_eq!(u64_at(&bytes, record + 8), previous[0].page_id);
                assert_eq!(u64_at(&bytes, record + 16), shift_range_salt(delta, epoch));
                assert_eq!(u32_at(&bytes, record + 24) as usize, count);
                assert_eq!(
                    i64::from_le_bytes(bytes[record + 32..record + 40].try_into().unwrap()),
                    delta
                );
                assert_eq!(u32_at(&bytes, 68) as usize, bytes.len());
                assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), after);
                let (legacy, _) = encode_frame_delta_incremental(
                    &after,
                    &legacy_snapshots,
                    placeholder_epochs(epoch),
                    &mut next_id,
                    &HashSet::new(),
                )
                .unwrap();
                assert_eq!(
                    apply_placeholder_test_frame(&legacy, &mut legacy_retained),
                    after
                );
                for (index, snapshot) in snapshots.iter().enumerate() {
                    assert_eq!(
                        snapshot.fingerprint,
                        previous[index].fingerprint ^ shift_range_salt(delta, epoch)
                    );
                    assert_eq!(
                        snapshot.fingerprint,
                        retained[&snapshot.page_id].fingerprint
                    );
                    assert_eq!(
                        snapshot.position_base,
                        previous[index].position_base + delta
                    );
                    assert_eq!(snapshot.positions, previous[index].positions);
                    assert_eq!(snapshot.note_anchors, previous[index].note_anchors);
                    assert_eq!(snapshot.position_span, after.pages[index].position_span);
                    assert!(Rc::ptr_eq(
                        &snapshot.primitive_ids,
                        &previous[index].primitive_ids
                    ));
                    let mut materialized = snapshot.clone();
                    materialized.materialize_positions();
                    assert_eq!(
                        materialized.positions,
                        primitive_positions(&after.pages[index])
                    );
                }
                current = after;
            }
        }
    }

    /// Rebuilding identical shifted content retains the shared fingerprint token.
    #[test]
    fn identical_rebuilt_pages_after_a_range_shift_emit_nothing() {
        for before in [range_shift_built_list(), placeholder_list(Some([10, 20]))] {
            let count = before.pages.len();
            let mut next_id = 0;
            let (_, mut snapshots) =
                encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id)
                    .unwrap();
            let after = shifted_body_list(&before, 0..count, 6);
            encode_frame_delta_changes(
                &after,
                &mut snapshots,
                placeholder_epochs(2),
                DisplayChanges {
                    rebuilt: &[],
                    repositioned: &[],
                    shifts: &[PageShiftRun {
                        start: 0,
                        end: count,
                        delta: 6,
                    }],
                },
            )
            .unwrap();
            let mut shifted = snapshots.clone();
            let bytes = encode_frame_delta_changes(
                &after,
                &mut snapshots,
                placeholder_epochs(3),
                DisplayChanges {
                    rebuilt: &[0, 0],
                    repositioned: &[],
                    shifts: &[],
                },
            )
            .unwrap();
            assert_eq!(u32_at(&bytes, 52), 0);
            assert_eq!(snapshots[0].fingerprint, shifted[0].fingerprint);
            assert_eq!(snapshots[0].position_base, 0);
            for (snapshot, old) in snapshots[1..].iter().zip(&shifted[1..]) {
                assert_eq!(snapshot.position_base, old.position_base);
                assert_eq!(snapshot.positions, old.positions);
            }
            let rebuilt: HashSet<_> = (0..count).collect();
            for snapshot in &mut snapshots {
                snapshot.materialize_positions();
            }
            for snapshot in &mut shifted {
                snapshot.materialize_positions();
            }
            assert_eq!(snapshots, shifted);
            for (bytes, next) in [
                encode_frame_delta(
                    &after,
                    &snapshots,
                    placeholder_epochs(4),
                    false,
                    &mut next_id,
                )
                .unwrap(),
                encode_frame_delta_incremental(
                    &after,
                    &snapshots,
                    placeholder_epochs(4),
                    &mut next_id,
                    &rebuilt,
                )
                .unwrap(),
                encode_frame_delta_pages(
                    &after,
                    &snapshots,
                    placeholder_epochs(4),
                    &mut next_id,
                    &|_| true,
                )
                .unwrap(),
            ] {
                assert_eq!(u32_at(&bytes, 52), 0);
                assert_eq!(next, snapshots);
            }
        }
    }

    /// Deferred shifts round trip without changing positions outside the body.
    #[test]
    fn materialize_positions_folds_only_present_body_positions() {
        let before = range_shift_built_list();
        let (_, mut snapshots) =
            encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut 0).unwrap();
        let original = snapshots.remove(0);
        let mut snapshot = original.clone();
        let after = shifted_body_list(&before, 0..1, 6);
        snapshot.position_base = 6;
        snapshot.materialize_positions();
        assert_eq!(snapshot.positions, primitive_positions(&after.pages[0]));
        assert_eq!(snapshot.position_base, 0);
        assert_eq!(snapshot.fingerprint, original.fingerprint);
        assert_eq!(snapshot.note_anchors, original.note_anchors);
        let materialized = snapshot.clone();
        snapshot.materialize_positions();
        assert_eq!(snapshot, materialized);
        snapshot.position_base = -6;
        snapshot.materialize_positions();
        assert_eq!(snapshot, original);
    }

    /// Change sets reject invalid indices, overlapping pages and invalid ranges.
    #[test]
    fn display_changes_validate_indices_ranges_and_snapshot_length() {
        let list = range_shift_built_list();
        let (_, snapshots) =
            encode_frame_delta(&list, &[], placeholder_epochs(1), true, &mut 0).unwrap();
        let attempt = |rebuilt: &[usize], repositioned: &[usize], shifts: &[PageShiftRun]| {
            encode_frame_delta_changes(
                &list,
                &mut snapshots.clone(),
                placeholder_epochs(2),
                DisplayChanges {
                    rebuilt,
                    repositioned,
                    shifts,
                },
            )
        };
        assert!(attempt(&[0], &[0], &[]).is_err());
        assert!(attempt(&[3], &[], &[]).is_err());
        assert!(attempt(&[], &[usize::MAX], &[]).is_err());
        let range = PageShiftRun {
            start: 0,
            end: 2,
            delta: 1,
        };
        assert!(attempt(&[1], &[], &[range]).is_err());
        assert!(attempt(&[], &[0], &[range]).is_err());
        assert!(
            attempt(
                &[],
                &[],
                &[
                    range,
                    PageShiftRun {
                        start: 1,
                        end: 3,
                        delta: 1
                    }
                ]
            )
            .is_err()
        );
        for run in [
            PageShiftRun {
                start: 0,
                end: 4,
                delta: 1,
            },
            PageShiftRun {
                start: 2,
                end: 1,
                delta: 1,
            },
            PageShiftRun {
                start: 1,
                end: 1,
                delta: 1,
            },
            PageShiftRun {
                start: 0,
                end: 1,
                delta: 0,
            },
            PageShiftRun {
                start: 0,
                end: 1,
                delta: MAX_SAFE_INTEGER + 1,
            },
            PageShiftRun {
                start: 0,
                end: 1,
                delta: -MAX_SAFE_INTEGER - 1,
            },
        ] {
            assert!(attempt(&[], &[], &[run]).is_err());
        }
        assert!(
            encode_frame_delta_changes(
                &list,
                &mut snapshots[..2].to_vec(),
                placeholder_epochs(2),
                DisplayChanges {
                    rebuilt: &[],
                    repositioned: &[],
                    shifts: &[]
                },
            )
            .is_err()
        );
    }

    /// Sorting and deduplication preserve existing upsert and patch bytes.
    #[test]
    fn display_changes_reuse_per_page_encoding_and_leave_unreported_snapshots_untouched() {
        let before = list_pages(&[
            ("P1", "one"),
            ("P2", "two"),
            ("P3", "three"),
            ("P4", "four"),
        ]);
        let mut after = before.clone();
        let Primitive::Text(first) = &mut after.pages[0].primitives[0] else {
            unreachable!()
        };
        first.text.push('!');
        let Primitive::Text(second) = &mut after.pages[1].primitives[0] else {
            unreachable!()
        };
        second.attrs.doc_start = Some(2);
        second.attrs.doc_end = Some(7);
        let Primitive::Text(last) = &mut after.pages[3].primitives[0] else {
            unreachable!()
        };
        last.text.push('!');
        let mut next_id = 0;
        let (full, previous) =
            encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id).unwrap();
        let mut snapshots = previous.clone();
        let bytes = encode_frame_delta_changes(
            &after,
            &mut snapshots,
            placeholder_epochs(2),
            DisplayChanges {
                rebuilt: &[3, 0, 3],
                repositioned: &[1, 1],
                shifts: &[],
            },
        )
        .unwrap();
        let (legacy, legacy_snapshots) = encode_frame_delta_pages(
            &after,
            &previous,
            placeholder_epochs(2),
            &mut next_id,
            &|index| index == 0 || index == 3,
        )
        .unwrap();
        assert_eq!(bytes, legacy);
        assert_eq!(snapshots, legacy_snapshots);
        assert_eq!(snapshots[2], previous[2]);
        assert_eq!(u32_at(&bytes, 52), 3);
        assert_eq!(
            bytes[FRAME_HEADER_LEN + PAGE_OP_LEN],
            PAGE_OP_PATCH_POSITIONS
        );
        for (op, index) in [0, 1, 3].into_iter().enumerate() {
            assert_eq!(
                u32_at(&bytes, FRAME_HEADER_LEN + op * PAGE_OP_LEN + 4),
                index
            );
        }
        let mut retained = HashMap::new();
        apply_placeholder_test_frame(&full, &mut retained);
        assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), after);
        let before_empty = snapshots.clone();
        let bytes = encode_frame_delta_changes(
            &after,
            &mut snapshots,
            placeholder_epochs(3),
            DisplayChanges {
                rebuilt: &[],
                repositioned: &[],
                shifts: &[],
            },
        )
        .unwrap();
        assert_eq!(u32_at(&bytes, 52), 0);
        assert_eq!(snapshots, before_empty);
    }

    /// Mixed changes and unsorted shift ranges emit operations in page order.
    #[test]
    fn display_changes_order_rebuilds_repositions_and_shift_ranges() {
        let before = list_pages(&[
            ("P1", "one"),
            ("P2", "two"),
            ("P3", "three"),
            ("P4", "four"),
            ("P5", "five"),
            ("P6", "six"),
        ]);
        let mut after = shifted_body_list(&before, 1..3, 4);
        after = shifted_body_list(&after, 5..6, -1);
        for index in [0, 4] {
            let Primitive::Text(value) = &mut after.pages[index].primitives[0] else {
                unreachable!()
            };
            value.text.push('!');
        }
        let Primitive::Text(value) = &mut after.pages[3].primitives[0] else {
            unreachable!()
        };
        value.attrs.doc_start = Some(2);
        value.attrs.doc_end = Some(7);
        let (full, mut snapshots) =
            encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut 0).unwrap();
        let previous = snapshots.clone();
        let bytes = encode_frame_delta_changes(
            &after,
            &mut snapshots,
            placeholder_epochs(2),
            DisplayChanges {
                rebuilt: &[4, 0, 4],
                repositioned: &[3, 3],
                shifts: &[
                    PageShiftRun {
                        start: 5,
                        end: 6,
                        delta: -1,
                    },
                    PageShiftRun {
                        start: 1,
                        end: 3,
                        delta: 4,
                    },
                ],
            },
        )
        .unwrap();
        assert_eq!(u32_at(&bytes, 52), 5);
        for (op, (index, opcode)) in [
            (0, PAGE_OP_UPSERT),
            (1, PAGE_OP_SHIFT_RANGE),
            (3, PAGE_OP_PATCH_POSITIONS),
            (4, PAGE_OP_UPSERT),
            (5, PAGE_OP_SHIFT_RANGE),
        ]
        .into_iter()
        .enumerate()
        {
            let record = FRAME_HEADER_LEN + op * PAGE_OP_LEN;
            assert_eq!(u32_at(&bytes, record + 4), index);
            assert_eq!(bytes[record], opcode);
            assert_eq!(u64_at(&bytes, record + 8), previous[index as usize].page_id);
        }
        let mut retained = HashMap::new();
        apply_placeholder_test_frame(&full, &mut retained);
        assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), after);
        for (index, snapshot) in snapshots.iter_mut().enumerate() {
            assert_eq!(
                snapshot.fingerprint,
                retained[&snapshot.page_id].fingerprint
            );
            snapshot.materialize_positions();
            assert_eq!(snapshot.positions, primitive_positions(&after.pages[index]));
        }
    }

    #[test]
    fn typed_emission_matches_the_value_reference() {
        let mut watermark = list("watermark");
        watermark.pages[0].watermark_primitive_count = Some(1);
        for list in [list("hello"), list_at_position(2), rich_list(), watermark] {
            for page in &list.pages {
                let value = serde_json::to_value(page).unwrap();

                let mut table = StringTable::default();
                let mut out = Vec::new();
                encode_page(page, &mut table, &mut out).unwrap();
                let strings = table.into_strings();
                let mut reference_strings = BTreeSet::new();
                collect_strings(&value, &mut reference_strings);
                assert_eq!(
                    strings.iter().cloned().collect::<BTreeSet<_>>(),
                    reference_strings,
                    "string tables differ"
                );
                let mut cursor = 0;
                let decoded = decode_typed(&out, &mut cursor, &strings);
                assert_eq!(cursor, out.len(), "typed stream has trailing bytes");
                let mut expected = value;
                canonicalize_compact_glyphs(&mut expected);
                assert_eq!(decoded, expected, "typed stream decodes differently");
            }
        }
    }

    #[test]
    fn changing_the_watermark_prefix_requires_a_page_upsert() {
        let before = list("watermark");
        let mut after = before.clone();
        after.pages[0].watermark_primitive_count = Some(1);
        let epochs = FrameEpochs {
            doc_epoch: 1,
            layout_epoch: 1,
            frame_epoch: 1,
            base_frame_epoch: 0,
        };
        let mut next_id = 0;
        let (_, snapshot) = encode_frame_delta(&before, &[], epochs, true, &mut next_id).unwrap();
        let (bytes, _) = encode_frame_delta(
            &after,
            &snapshot,
            FrameEpochs {
                frame_epoch: 2,
                base_frame_epoch: 1,
                ..epochs
            },
            false,
            &mut next_id,
        )
        .unwrap();
        assert_eq!(bytes[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
    }

    #[test]
    fn duplicate_flattened_keys_compact_to_the_last_write() {
        // a decorative shape sets both the named primitive field and the
        // flattened DocAttrs member; serde_json::Map deduped this on the old
        // path, and the streaming emitter must match (the browser decoder
        // rejects duplicate object keys outright)
        let mut list: DisplayList = serde_json::from_value(serde_json::json!({
            "pages": [{
                "pageIndex": 0,
                "width": 816,
                "height": 1056,
                "primitives": [{
                    "kind": "shape",
                    "x": 10, "y": 20, "w": 30, "h": 40,
                    "geometryPath": [],
                    "docStart": 1, "docEnd": 2
                }]
            }]
        }))
        .unwrap();
        let docx_layout::display_list::Primitive::Shape(shape) = &mut list.pages[0].primitives[0]
        else {
            panic!("shape expected");
        };
        shape.decorative = true;
        shape.attrs.decorative = Some(true);

        let page = &list.pages[0];
        let mut table = StringTable::default();
        let mut out = Vec::new();
        encode_page(page, &mut table, &mut out).unwrap();
        let strings = table.into_strings();
        let mut cursor = 0;
        let decoded = decode_typed(&out, &mut cursor, &strings);
        assert_eq!(cursor, out.len());
        let mut expected = serde_json::to_value(page).unwrap();
        canonicalize_compact_glyphs(&mut expected);
        assert_eq!(
            decoded, expected,
            "duplicate keys must compact to serde_json's last-write value"
        );
        assert_eq!(
            expected["primitives"][0]["decorative"],
            serde_json::json!(true)
        );
    }

    #[test]
    fn reference_hashes_agree_with_the_streaming_exclusion_semantics() {
        // the retained Value-based hashes and the streaming hashes must agree
        // on WHAT is excluded, even though their accumulation orders differ
        let before = serde_json::to_value(&list_at_position(2).pages[0]).unwrap();
        let after = serde_json::to_value(&list_at_position(12).pages[0]).unwrap();
        assert_ne!(hash_page_value(&before), hash_page_value(&after));
        assert_eq!(
            hash_visual_page_value(&before),
            hash_visual_page_value(&after)
        );
        let streamed_before = hash_page(&list_at_position(2).pages[0]).unwrap();
        let streamed_after = hash_page(&list_at_position(12).pages[0]).unwrap();
        assert_ne!(streamed_before.fingerprint, streamed_after.fingerprint);
        assert_eq!(
            streamed_before.visual_fingerprint,
            streamed_after.visual_fingerprint
        );
    }

    #[test]
    fn typed_hashes_are_deterministic_and_position_scoped() {
        let before = hash_page(&list_at_position(2).pages[0]).unwrap();
        let again = hash_page(&list_at_position(2).pages[0]).unwrap();
        assert_eq!(before.fingerprint, again.fingerprint);
        assert_eq!(before.visual_fingerprint, again.visual_fingerprint);

        let after = hash_page(&list_at_position(12).pages[0]).unwrap();
        assert_ne!(
            before.fingerprint, after.fingerprint,
            "position changes alter the structural fingerprint"
        );
        assert_eq!(
            before.visual_fingerprint, after.visual_fingerprint,
            "position changes preserve the visual fingerprint"
        );

        let content = hash_page(&list("other").pages[0]).unwrap();
        assert_ne!(before.visual_fingerprint, content.visual_fingerprint);
    }

    #[test]
    fn a_page_fingerprints_the_same_whatever_was_emitted_before_it() {
        let list = list_pages(&[("P1", "alpha"), ("P2", "beta")]);
        let alone = hash_page(&list.pages[1]).unwrap();
        let mut table = StringTable::default();
        let mut out = Vec::new();
        encode_page(&list.pages[0], &mut table, &mut out).unwrap();
        let after = encode_page(&list.pages[1], &mut table, &mut out).unwrap();
        assert_eq!(alone.fingerprint, after.fingerprint);
        assert_eq!(alone.visual_fingerprint, after.visual_fingerprint);
    }

    #[test]
    fn a_moved_glyph_changes_both_fingerprints() {
        let list = rich_list();
        let mut value = serde_json::to_value(&list).unwrap();
        value["pages"][0]["primitives"][0]["glyphs"][1]["x"] = serde_json::json!(9.5);
        let moved: DisplayList = serde_json::from_value(value).unwrap();
        let before = hash_page(&list.pages[0]).unwrap();
        let after = hash_page(&moved.pages[0]).unwrap();
        assert_ne!(before.fingerprint, after.fingerprint);
        assert_ne!(before.visual_fingerprint, after.visual_fingerprint);
    }

    #[test]
    fn a_glyph_change_spread_over_two_coordinates_changes_both_fingerprints() {
        let glyph_at = |x: f64, y: f64| {
            let mut value = serde_json::to_value(rich_list()).unwrap();
            let glyph = &mut value["pages"][0]["primitives"][0]["glyphs"][1];
            glyph["x"] = serde_json::json!(x);
            glyph["y"] = serde_json::json!(y);
            let list: DisplayList = serde_json::from_value(value).unwrap();
            hash_page(&list.pages[0]).unwrap()
        };
        let before = glyph_at(1.0, 1.0);
        let after = glyph_at(-1.0, f64::from_bits(0xbff0_0000_8000_0000));
        assert_ne!(before.fingerprint, after.fingerprint);
        assert_ne!(before.visual_fingerprint, after.visual_fingerprint);
    }

    /// The upsert bounds the browser decoder checks: aligned primitive ids,
    /// then a non-empty payload, all inside the frame.
    fn assert_upserts_in_bounds(bytes: &[u8]) {
        let data_offset = u32_at(bytes, 68) as usize;
        for op in 0..u32_at(bytes, 52) as usize {
            let record = FRAME_HEADER_LEN + op * PAGE_OP_LEN;
            if bytes[record] != PAGE_OP_UPSERT {
                continue;
            }
            let ids = u32_at(bytes, record + 28) as usize;
            let ids_end = ids + 8 * u32_at(bytes, record + 24) as usize;
            let payload = u32_at(bytes, record + 32) as usize;
            let payload_end = payload + u32_at(bytes, record + 36) as usize;
            assert!(ids >= data_offset && ids % 8 == 0);
            assert!(ids_end <= payload && payload < payload_end && payload_end <= bytes.len());
        }
    }

    #[test]
    fn a_rebuilt_page_that_did_not_change_adds_nothing_to_the_frame() {
        let before = list_pages(&[("P1", "first"), ("P2", "second"), ("P3", "third")]);
        let after = list_pages(&[("P1", "first!"), ("P2", "second"), ("P3", "third!")]);
        let epochs = |frame_epoch| FrameEpochs {
            doc_epoch: frame_epoch,
            layout_epoch: frame_epoch,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        };
        let mut next_id = 0;
        let (_, snapshot) =
            encode_frame_delta(&before, &[], epochs(1), true, &mut next_id).unwrap();
        let delta = |rebuilt: &[usize]| {
            let mut next_id = next_id;
            encode_frame_delta_incremental(
                &after,
                &snapshot,
                epochs(2),
                &mut next_id,
                &rebuilt.iter().copied().collect(),
            )
            .unwrap()
        };
        let changed_only = delta(&[0, 2]);
        assert_eq!(u32_at(&changed_only.0, 52), 2);
        assert_upserts_in_bounds(&changed_only.0);
        assert_eq!(delta(&[0, 1, 2]), changed_only);
    }

    #[test]
    fn glyph_arrays_use_the_compact_fixed_field_payload() {
        let value = serde_json::json!([{
            "id": 42,
            "x": 10.5,
            "y": 20.25,
            "cluster": 3,
            "advance": 7.75,
            "logicalOrder": 4,
            "bidiLevel": 1
        }]);
        let mut out = Vec::new();
        encode_value(&value, &HashMap::new(), &mut out, Some("glyphs")).unwrap();
        assert_eq!(out[0], VALUE_GLYPH_ARRAY);
        assert_eq!(u32_at(&out, 1), 42);
        assert_eq!(u32_at(&out, 5), 1);
        assert_eq!(out.len(), 51);
        assert_eq!(u32_at(&out, 9), 42);
        assert_eq!(u32_at(&out, 29), 3);
        assert_eq!(out[41], GLYPH_LOGICAL_ORDER | GLYPH_BIDI_LEVEL);
        assert_eq!(u64_at(&out, 42), 4);
        assert_eq!(out[50], 1);
    }

    #[test]
    fn pages_built_out_of_order_keep_their_own_content() {
        // Every built page opens with the same header paragraph, so semantic
        // anchors count occurrences, and building page 5 renumbers 7 and 8.
        let pages = |built: &[usize]| -> DisplayList {
            let pages: Vec<_> = (0..10)
                .map(|index| {
                    if built.contains(&index) {
                        serde_json::json!({
                            "pageIndex": index, "width": 816, "height": 1056,
                            "primitives": [{
                                "kind": "text", "text": format!("page {index}"), "x": 96,
                                "baselineY": 120, "width": 40, "font": "16px serif",
                                "color": "#000000", "blockId": 7, "paraId": "HEADER"
                            }]
                        })
                    } else {
                        serde_json::json!({
                            "pageIndex": index, "width": 816, "height": 1056,
                            "primitives": [], "unbuilt": true
                        })
                    }
                })
                .collect();
            serde_json::from_value(serde_json::json!({ "contractVersion": 1, "pages": pages }))
                .unwrap()
        };
        let epochs = |frame_epoch| FrameEpochs {
            doc_epoch: 1,
            layout_epoch: 1,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        };
        let mut next_id = 0;
        let (_, snapshots) =
            encode_frame_delta(&pages(&[0, 1, 2, 3, 4]), &[], epochs(1), true, &mut next_id)
                .unwrap();
        let (_, snapshots) = encode_frame_delta_pages(
            &pages(&[0, 1, 2, 3, 4, 7, 8]),
            &snapshots,
            epochs(2),
            &mut next_id,
            &|index| index == 7 || index == 8,
        )
        .unwrap();
        let built = pages(&[0, 1, 2, 3, 4, 5, 7, 8]);
        let (_, snapshots) =
            encode_frame_delta_pages(&built, &snapshots, epochs(3), &mut next_id, &|index| {
                index == 5
            })
            .unwrap();
        let (_, fresh) = encode_frame_delta(&built, &[], epochs(3), true, &mut 0).unwrap();
        for (retained, fresh) in snapshots.iter().zip(&fresh) {
            assert_eq!(retained.page_index, fresh.page_index);
            assert_eq!(retained.fingerprint, fresh.fingerprint);
            assert_eq!(retained.primitive_ids, fresh.primitive_ids);
        }
    }

    #[test]
    fn an_incremental_frame_that_builds_a_placeholder_keeps_later_pages_their_own_content() {
        // Every built page opens with the same header paragraph, so building
        // page 6 in an edit frame renumbers the anchors of pages 8 and 9.
        let pages = |built: &[usize]| -> DisplayList {
            let pages: Vec<_> = (0..10)
                .map(|index| {
                    if built.contains(&index) {
                        serde_json::json!({
                            "pageIndex": index, "width": 816, "height": 1056,
                            "primitives": [{
                                "kind": "text", "text": format!("page {index}"), "x": 96,
                                "baselineY": 120, "width": 40, "font": "16px serif",
                                "color": "#000000", "blockId": 7, "paraId": "HEADER"
                            }]
                        })
                    } else {
                        serde_json::json!({
                            "pageIndex": index, "width": 816, "height": 1056,
                            "primitives": [], "unbuilt": true
                        })
                    }
                })
                .collect();
            serde_json::from_value(serde_json::json!({ "contractVersion": 1, "pages": pages }))
                .unwrap()
        };
        let epochs = |frame_epoch| FrameEpochs {
            doc_epoch: frame_epoch,
            layout_epoch: frame_epoch,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        };
        let mut next_id = 0;
        let (_, snapshots) =
            encode_frame_delta(&pages(&[0, 1, 2, 3, 4]), &[], epochs(1), true, &mut next_id)
                .unwrap();
        let (_, snapshots) = encode_frame_delta_pages(
            &pages(&[0, 1, 2, 3, 4, 8, 9]),
            &snapshots,
            epochs(2),
            &mut next_id,
            &|index| index == 8 || index == 9,
        )
        .unwrap();
        let built = pages(&[0, 1, 2, 3, 4, 6, 8, 9]);
        let (_, snapshots) = encode_frame_delta_incremental(
            &built,
            &snapshots,
            epochs(3),
            &mut next_id,
            &HashSet::from([6]),
        )
        .unwrap();
        let (_, fresh) = encode_frame_delta(&built, &[], epochs(3), true, &mut 0).unwrap();
        for (retained, fresh) in snapshots.iter().zip(&fresh) {
            assert_eq!(retained.page_index, fresh.page_index);
            assert_eq!(retained.fingerprint, fresh.fingerprint);
            assert_eq!(retained.primitive_ids, fresh.primitive_ids);
        }
    }

    #[test]
    fn releasing_a_repeated_header_page_preserves_identity_through_rebuilds_and_edits() {
        let list = |released: bool, suffix: &str| -> DisplayList {
            let pages: Vec<_> = (0..4).map(|index| {
                if released && index == 0 {
                    serde_json::json!({
                        "pageIndex": index, "width": 816, "height": 1056,
                        "primitives": [], "unbuilt": true, "positionSpan": [1, 9]
                    })
                } else {
                    serde_json::json!({
                        "pageIndex": index, "width": 816, "height": 1056,
                        "primitives": [{
                            "kind": "text", "text": format!("page {index}{}", if index == 3 { suffix } else { "" }),
                            "x": 96, "baselineY": 120, "width": 40, "font": "16px serif",
                            "color": "#000000", "blockId": 7, "paraId": "HEADER"
                        }]
                    })
                }
            }).collect();
            serde_json::from_value(serde_json::json!({ "contractVersion": 1, "pages": pages }))
                .unwrap()
        };
        let epochs = |frame_epoch| FrameEpochs {
            doc_epoch: frame_epoch,
            layout_epoch: frame_epoch,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        };
        let mut next_id = 0;
        let (_, before) =
            encode_frame_delta(&list(false, ""), &[], epochs(1), true, &mut next_id).unwrap();
        let (release, mut snapshots) = encode_frame_delta_pages(
            &list(true, ""),
            &before,
            epochs(2),
            &mut next_id,
            &|index| index == 0,
        )
        .unwrap();
        assert_eq!(u32_at(&release, 52), 1);
        assert_eq!(release[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
        assert_eq!(u64_at(&release, FRAME_HEADER_LEN + 8), before[0].page_id);
        assert!(snapshots[0].primitive_ids.is_empty());
        assert!(snapshots[0].positions.is_empty());
        for (actual, old) in snapshots.iter().zip(&before) {
            assert_eq!(actual.page_id, old.page_id);
        }
        for (epoch, released, suffix, rebuilt) in [
            (3, true, " edited", 3),
            (4, false, " edited", 0),
            (5, false, " edited again", 3),
        ] {
            let current = list(released, suffix);
            let (_, next) = if rebuilt == 0 {
                encode_frame_delta_pages(
                    &current,
                    &snapshots,
                    epochs(epoch),
                    &mut next_id,
                    &|index| index == 0,
                )
                .unwrap()
            } else {
                encode_frame_delta_incremental(
                    &current,
                    &snapshots,
                    epochs(epoch),
                    &mut next_id,
                    &HashSet::from([rebuilt]),
                )
                .unwrap()
            };
            let (_, fresh) =
                encode_frame_delta(&current, &[], epochs(epoch), true, &mut 0).unwrap();
            for (actual, expected) in next.iter().zip(&fresh) {
                assert_eq!(actual.page_id, expected.page_id);
                assert_eq!(actual.fingerprint, expected.fingerprint);
                assert_eq!(actual.primitive_ids, expected.primitive_ids);
            }
            snapshots = next;
        }
    }

    #[test]
    fn an_unbuilt_page_whose_position_span_moves_is_sent_again() {
        let list = |span: [i64; 2]| -> DisplayList {
            serde_json::from_value(serde_json::json!({ "contractVersion": 1, "pages": [
                {
                    "pageIndex": 0, "width": 816, "height": 1056,
                    "primitives": [{
                        "kind": "text", "text": "built", "x": 96, "baselineY": 120,
                        "width": 40, "font": "16px serif", "color": "#000000",
                        "blockId": 7, "docStart": 1, "docEnd": 6
                    }]
                },
                {
                    "pageIndex": 1, "width": 816, "height": 1056,
                    "primitives": [], "unbuilt": true, "positionSpan": span
                }
            ]}))
            .unwrap()
        };
        let epochs = |frame_epoch| FrameEpochs {
            doc_epoch: frame_epoch,
            layout_epoch: frame_epoch,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        };
        let mut next_id = 0;
        let (_, snapshots) =
            encode_frame_delta(&list([8, 30]), &[], epochs(1), true, &mut next_id).unwrap();

        let (unchanged, snapshots) = encode_frame_delta_incremental(
            &list([8, 30]),
            &snapshots,
            epochs(2),
            &mut next_id,
            &HashSet::new(),
        )
        .unwrap();
        assert_eq!(u32_at(&unchanged, 52), 0);

        let (moved, _) = encode_frame_delta_incremental(
            &list([14, 36]),
            &snapshots,
            epochs(3),
            &mut next_id,
            &HashSet::new(),
        )
        .unwrap();
        assert_eq!(u32_at(&moved, 52), 1);
        assert_eq!(moved[FRAME_HEADER_LEN], PAGE_OP_SHIFT_POSITIONS);
    }

    fn placeholder_list(span: Option<[i64; 2]>) -> DisplayList {
        let mut list: DisplayList = serde_json::from_value(serde_json::json!({ "pages": [{
            "pageIndex": 0, "width": 816, "height": 1056,
            "primitives": [], "unbuilt": true
        }]}))
        .unwrap();
        list.pages[0].position_span = span;
        list
    }

    /// Range shifts that cancel out in position never restore an earlier fingerprint.
    #[test]
    fn cancelling_range_shifts_never_restore_an_earlier_fingerprint() {
        let mut fingerprint = 0x1234_u64;
        let mut seen = HashSet::from([fingerprint]);
        for (epoch, delta) in [(2, 1), (3, 1), (4, -1), (5, -1), (6, 1), (7, -1)] {
            fingerprint ^= shift_range_salt(delta, epoch);
            assert!(seen.insert(fingerprint), "frame {epoch}");
        }
    }

    fn placeholder_epochs(frame_epoch: u64) -> FrameEpochs {
        FrameEpochs {
            doc_epoch: frame_epoch,
            layout_epoch: frame_epoch,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        }
    }

    #[test]
    fn an_unbuilt_page_whose_position_span_moves_uses_a_compact_shift() {
        let before = placeholder_list(Some([8, 30]));
        let mut next_id = 0;
        let (full, mut snapshots) =
            encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id).unwrap();
        let mut retained = HashMap::new();
        assert_eq!(apply_placeholder_test_frame(&full, &mut retained), before);
        for (epoch, span) in [(2, [14, 36]), (3, [10, 32]), (4, [10, 32])] {
            let next = placeholder_list(Some(span));
            let (bytes, next_snapshots) = encode_frame_delta_incremental(
                &next,
                &snapshots,
                placeholder_epochs(epoch),
                &mut next_id,
                &HashSet::new(),
            )
            .unwrap();
            if epoch == 4 {
                assert_eq!(u32_at(&bytes, 52), 0);
            } else {
                assert_eq!(u32_at(&bytes, 52), 1);
                assert_eq!(bytes[FRAME_HEADER_LEN], PAGE_OP_SHIFT_POSITIONS);
                assert_eq!(u32_at(&bytes, FRAME_HEADER_LEN + 24), 0);
                assert_eq!(u32_at(&bytes, FRAME_HEADER_LEN + 40), 0);
                let payload = u32_at(&bytes, FRAME_HEADER_LEN + 32) as usize;
                assert_eq!(u32_at(&bytes, payload + 4), SHIFT_SPAN_PRESENT);
                assert_eq!(u32_at(&bytes, FRAME_HEADER_LEN + 36), 16);
            }
            assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), next);
            snapshots = next_snapshots;
        }
    }

    #[test]
    fn compact_placeholders_upsert_unsafe_position_span_shifts() {
        let max = MAX_SAFE_INTEGER;
        for (before_span, after_span) in [
            ([-max, -max + 1], [1, 2]),
            ([1, 2], [-max, -max + 1]),
            ([-max, -max + 1], [-max - 1, -max]),
            ([-max - 1, -max], [-max, -max + 1]),
            ([max - 1, max], [max, max + 1]),
            ([max, max + 1], [max - 1, max]),
        ] {
            let before = placeholder_list(Some(before_span));
            let after = placeholder_list(Some(after_span));
            let mut next_id = 0;
            let (full, snapshots) =
                encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id)
                    .unwrap();
            let mut retained = HashMap::new();
            assert_eq!(apply_placeholder_test_frame(&full, &mut retained), before);
            let (bytes, _) = encode_frame_delta_incremental(
                &after,
                &snapshots,
                placeholder_epochs(2),
                &mut next_id,
                &HashSet::new(),
            )
            .unwrap();
            assert_eq!(u32_at(&bytes, 52), 1);
            assert_eq!(bytes[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
            assert_eq!(apply_placeholder_test_frame(&bytes, &mut retained), after);
        }
    }

    #[test]
    fn compact_placeholders_upsert_nonuniform_or_missing_spans_and_metadata_changes() {
        let before = placeholder_list(Some([8, 30]));
        let moved = placeholder_list(Some([14, 36]));
        let mut changed = moved.clone();
        changed.pages[0].page_label = Some("ii".to_owned());
        let mut geometry = moved.clone();
        geometry.pages[0].width = 800.into();
        let mut built = moved.clone();
        built.pages[0].unbuilt = false;
        built.pages[0].primitives = list_at_position(14).pages[0].primitives.clone();
        let cases = [
            (placeholder_list(None), moved.clone()),
            (before.clone(), placeholder_list(None)),
            (before.clone(), placeholder_list(Some([14, 37]))),
            (before.clone(), changed),
            (before.clone(), geometry),
            (before.clone(), built.clone()),
            (built, moved.clone()),
            (
                placeholder_list(Some([i64::MIN, 0])),
                placeholder_list(Some([0, i64::MAX])),
            ),
        ];
        for (before, after) in cases {
            let mut next_id = 0;
            let (full, snapshots) =
                encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id)
                    .unwrap();
            let mut retained = HashMap::new();
            assert_eq!(apply_placeholder_test_frame(&full, &mut retained), before);
            let rebuilt = if before.pages[0].unbuilt && !after.pages[0].unbuilt {
                HashSet::from([0])
            } else {
                HashSet::new()
            };
            let (bytes, next_snapshots) = encode_frame_delta_incremental(
                &after,
                &snapshots,
                placeholder_epochs(2),
                &mut next_id,
                &rebuilt,
            )
            .unwrap();
            assert_eq!(u32_at(&bytes, 52), 1);
            assert_eq!(bytes[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
            let applied = apply_placeholder_test_frame(&bytes, &mut retained);
            assert_eq!(applied, after);
            if !rebuilt.is_empty() {
                assert!(!applied.pages[0].unbuilt);
                assert_eq!(applied.pages[0].position_span, Some([14, 36]));
                assert_eq!(applied.pages[0].primitives, after.pages[0].primitives);
                assert_ne!(next_snapshots[0].fingerprint, snapshots[0].fingerprint);
                assert_eq!(next_snapshots[0].primitive_ids.len(), 1);
            }
        }
        let mut next_id = 0;
        let (_, snapshots) =
            encode_frame_delta(&before, &[], placeholder_epochs(1), true, &mut next_id).unwrap();
        let (full, _) = encode_frame_delta(
            &moved,
            &snapshots,
            placeholder_epochs(2),
            true,
            &mut next_id,
        )
        .unwrap();
        assert_eq!(full[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
        let (new_page, _) =
            encode_frame_delta(&moved, &[], placeholder_epochs(2), false, &mut next_id).unwrap();
        assert_eq!(new_page[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
    }

    #[test]
    fn full_frame_has_fixed_header_real_typed_payload_and_stable_ids() {
        let mut next_id = 0;
        let epochs = FrameEpochs {
            doc_epoch: 4,
            layout_epoch: 5,
            frame_epoch: 6,
            base_frame_epoch: 0,
        };
        let (bytes, snapshot) =
            encode_frame_delta(&list("hello"), &[], epochs, true, &mut next_id).unwrap();
        assert_eq!(&bytes[0..4], b"FDV1");
        assert_eq!(u32_at(&bytes, 8) as usize, bytes.len());
        assert_eq!(u32_at(&bytes, 12), FRAME_FLAG_FULL);
        assert_eq!(u64_at(&bytes, 16), 4);
        assert_eq!(u64_at(&bytes, 24), 5);
        assert_eq!(u64_at(&bytes, 32), 6);
        assert_eq!(u32_at(&bytes, 48), 1);
        assert_eq!(u32_at(&bytes, 52), 1);
        assert_eq!(u32_at(&bytes, 72), 1);
        assert_eq!(bytes[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
        assert_eq!(u32_at(&bytes, FRAME_HEADER_LEN + 24), 1);
        let payload_offset = u32_at(&bytes, FRAME_HEADER_LEN + 32) as usize;
        let payload_len = u32_at(&bytes, FRAME_HEADER_LEN + 36) as usize;
        assert_eq!(payload_offset + payload_len, bytes.len());
        assert_upserts_in_bounds(&bytes);
        assert_eq!(bytes[payload_offset], VALUE_OBJECT);
        assert_eq!(snapshot[0].page_id, 1);

        let (again, next) = encode_frame_delta(
            &list("hello"),
            &snapshot,
            FrameEpochs {
                frame_epoch: 7,
                base_frame_epoch: 6,
                ..epochs
            },
            false,
            &mut next_id,
        )
        .unwrap();
        assert_eq!(u32_at(&again, 12), 0);
        assert_eq!(u64_at(&again, 40), 6);
        assert_eq!(u32_at(&again, 52), 0, "unchanged page emits no payload");
        assert_eq!(next[0].page_id, snapshot[0].page_id);
    }

    #[test]
    fn changed_page_is_one_upsert_and_keeps_page_and_primitive_identity() {
        let mut next_id = 0;
        let epochs = FrameEpochs {
            doc_epoch: 1,
            layout_epoch: 1,
            frame_epoch: 1,
            base_frame_epoch: 0,
        };
        let (first, snapshot) =
            encode_frame_delta(&list("hello"), &[], epochs, true, &mut next_id).unwrap();
        let first_primitive_offset = u32_at(&first, FRAME_HEADER_LEN + 28) as usize;
        let primitive_id = u64_at(&first, first_primitive_offset);

        let (delta, next) = encode_frame_delta(
            &list("hello!"),
            &snapshot,
            FrameEpochs {
                doc_epoch: 2,
                layout_epoch: 2,
                frame_epoch: 2,
                base_frame_epoch: 1,
            },
            false,
            &mut next_id,
        )
        .unwrap();
        assert_eq!(u32_at(&delta, 52), 1);
        assert_eq!(delta[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
        assert_eq!(u64_at(&delta, FRAME_HEADER_LEN + 8), snapshot[0].page_id);
        let next_primitive_offset = u32_at(&delta, FRAME_HEADER_LEN + 28) as usize;
        assert_eq!(u64_at(&delta, next_primitive_offset), primitive_id);
        assert_eq!(next[0].page_id, snapshot[0].page_id);
        assert_ne!(next[0].fingerprint, snapshot[0].fingerprint);
    }

    #[test]
    fn position_only_changes_shift_stable_primitives_without_page_damage() {
        let mut next_id = 0;
        let epochs = FrameEpochs {
            doc_epoch: 1,
            layout_epoch: 1,
            frame_epoch: 1,
            base_frame_epoch: 0,
        };
        let (_, snapshot) =
            encode_frame_delta(&list_at_position(2), &[], epochs, true, &mut next_id).unwrap();

        let (delta, next) = encode_frame_delta(
            &list_at_position(12),
            &snapshot,
            FrameEpochs {
                doc_epoch: 2,
                layout_epoch: 2,
                frame_epoch: 2,
                base_frame_epoch: 1,
            },
            false,
            &mut next_id,
        )
        .unwrap();

        assert_eq!(u32_at(&delta, 52), 1);
        assert_eq!(delta[FRAME_HEADER_LEN], PAGE_OP_SHIFT_POSITIONS);
        assert_eq!(u32_at(&delta, FRAME_HEADER_LEN + 24), 1);
        assert_eq!(next[0].page_id, snapshot[0].page_id);
        assert_eq!(next[0].primitive_ids, snapshot[0].primitive_ids);
        assert_eq!(next[0].visual_fingerprint, snapshot[0].visual_fingerprint);
        assert_ne!(next[0].fingerprint, snapshot[0].fingerprint);

        let payload = u32_at(&delta, FRAME_HEADER_LEN + 32) as usize;
        assert_eq!(u32_at(&delta, payload), 1);
        assert_eq!(u32_at(&delta, payload + 8), 0);
        assert_eq!(u32_at(&delta, payload + 12), 1);
        assert_eq!(
            delta[payload + 16],
            POSITION_DOC_START
                | POSITION_DOC_END
                | POSITION_FRAGMENT_START
                | POSITION_FRAGMENT_END
                | POSITION_PRESENT_ONLY
        );
        assert_eq!(
            i64::from_le_bytes(delta[payload + 24..payload + 32].try_into().unwrap()),
            10
        );
    }

    #[test]
    fn a_uniform_shift_of_mixed_primitives_is_one_present_only_run() {
        let at = |doc: Option<(i64, i64)>, fragment: Option<(i64, i64)>, widget: Option<i64>| {
            PrimitivePositionSnapshot {
                doc_start: doc.map(|range| range.0),
                doc_end: doc.map(|range| range.1),
                fragment_doc_start: fragment.map(|range| range.0),
                fragment_doc_end: fragment.map(|range| range.1),
                inline_widget_pos: widget,
            }
        };
        let moved = |delta: i64| {
            vec![
                at(Some((10 + delta, 14 + delta)), None, None),
                at(None, Some((9 + delta, 20 + delta)), None),
                at(None, None, None),
                at(Some((15 + delta, 16 + delta)), None, Some(15 + delta)),
                // A header primitive in its own story does not move.
                at(Some((3, 4)), None, None),
            ]
        };
        for delta in [7, -5] {
            let runs = position_shift_runs(&moved(0), &moved(delta)).unwrap();
            assert_eq!(runs.len(), 1, "delta {delta}");
            assert_eq!((runs[0].start, runs[0].count, runs[0].delta), (0, 4, delta));
            assert_eq!(
                runs[0].changed_mask,
                POSITION_DOC_START
                    | POSITION_DOC_END
                    | POSITION_FRAGMENT_START
                    | POSITION_FRAGMENT_END
                    | POSITION_INLINE_WIDGET
                    | POSITION_PRESENT_ONLY
            );
        }

        // A present field that stays put keeps its primitive on an exact run.
        let before = [at(Some((10, 14)), Some((9, 20)), None)];
        let after = [at(Some((12, 16)), Some((9, 20)), None)];
        let runs = position_shift_runs(&before, &after).unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].changed_mask, POSITION_DOC_START | POSITION_DOC_END);
    }

    fn list_with_note(body_start: i64, anchor: Option<i64>, label: &str) -> DisplayList {
        let mut value = serde_json::to_value(list_at_position(body_start)).unwrap();
        let mut note = serde_json::json!({"id": 7, "label": label});
        if let Some(anchor) = anchor {
            note["anchorDocStart"] = serde_json::json!(anchor);
            note["anchorDocEnd"] = serde_json::json!(anchor + 1);
        }
        value["pages"][0]["noteAreas"] = serde_json::json!([{
            "kind": "footnote",
            "primitives": [{"kind": "rect", "x": 96, "y": 900, "w": 10, "h": 10, "fill": "#000"}],
            "noteIds": [7],
            "notes": [note],
        }]);
        serde_json::from_value(value).unwrap()
    }

    fn delta_between(before: &DisplayList, after: &DisplayList) -> Vec<u8> {
        let mut next_id = 0;
        let epochs = |frame_epoch| FrameEpochs {
            doc_epoch: frame_epoch,
            layout_epoch: frame_epoch,
            frame_epoch,
            base_frame_epoch: frame_epoch - 1,
        };
        let (_, snapshot) = encode_frame_delta(before, &[], epochs(1), true, &mut next_id).unwrap();
        encode_frame_delta(after, &snapshot, epochs(2), false, &mut next_id)
            .unwrap()
            .0
    }

    #[test]
    fn a_moved_note_anchor_is_a_shift_not_an_upsert() {
        let record = FRAME_HEADER_LEN;
        let i64_at = |bytes: &[u8], offset: usize| {
            i64::from_le_bytes(bytes[offset..offset + 8].try_into().unwrap())
        };

        // The body and the note's reference move together.
        let delta = delta_between(
            &list_with_note(2, Some(3), "1"),
            &list_with_note(12, Some(13), "1"),
        );
        assert_eq!(delta[record], PAGE_OP_SHIFT_POSITIONS);
        assert_eq!(u32_at(&delta, record + 24), 1);
        assert_eq!(u32_at(&delta, record + 40), 1);
        let payload = u32_at(&delta, record + 32) as usize;
        let anchors = payload + 8 + 24;
        assert_eq!(u32_at(&delta, anchors), 1);
        assert_eq!(u32_at(&delta, anchors + 4), 0);
        assert_eq!(
            (u32_at(&delta, anchors + 8), u32_at(&delta, anchors + 12)),
            (0, 0)
        );
        assert_eq!(
            (i64_at(&delta, anchors + 16), i64_at(&delta, anchors + 24)),
            (13, 14)
        );
        assert_eq!(u32_at(&delta, record + 36) as usize, 8 + 24 + 8 + 24);

        // Only the reference moves, then it goes away.
        for (anchor, expected) in [(Some(40), (40, 41)), (None, (i64::MIN, i64::MIN))] {
            let delta = delta_between(
                &list_with_note(2, Some(3), "1"),
                &list_with_note(2, anchor, "1"),
            );
            assert_eq!(delta[record], PAGE_OP_SHIFT_POSITIONS);
            assert_eq!(u32_at(&delta, record + 24), 0);
            assert_eq!(u32_at(&delta, record + 40), 1);
            let payload = u32_at(&delta, record + 32) as usize;
            assert_eq!(u32_at(&delta, payload), 0);
            let anchors = payload + 8;
            assert_eq!(
                (i64_at(&delta, anchors + 16), i64_at(&delta, anchors + 24)),
                expected
            );
        }

        // A relabelled note redraws its page.
        let delta = delta_between(
            &list_with_note(2, Some(3), "1"),
            &list_with_note(12, Some(13), "2"),
        );
        assert_eq!(delta[record], PAGE_OP_UPSERT);
        assert_eq!(u32_at(&delta, record + 40), 0);
    }

    #[test]
    fn paragraph_damage_emits_only_its_page() {
        let mut next_id = 0;
        let epochs = FrameEpochs {
            doc_epoch: 1,
            layout_epoch: 1,
            frame_epoch: 1,
            base_frame_epoch: 0,
        };
        let (_, snapshot) = encode_frame_delta(
            &list_pages(&[("P1", "first"), ("P2", "second")]),
            &[],
            epochs,
            true,
            &mut next_id,
        )
        .unwrap();

        let (delta, next) = encode_frame_delta(
            &list_pages(&[("P1", "first!"), ("P2", "second")]),
            &snapshot,
            FrameEpochs {
                doc_epoch: 2,
                layout_epoch: 2,
                frame_epoch: 2,
                base_frame_epoch: 1,
            },
            false,
            &mut next_id,
        )
        .unwrap();

        assert_eq!(u32_at(&delta, 52), 1, "only one page operation crosses");
        assert_eq!(delta[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
        assert_eq!(u64_at(&delta, FRAME_HEADER_LEN + 8), snapshot[0].page_id);
        assert_eq!(
            next[1], snapshot[1],
            "the untouched page is retained exactly"
        );
    }

    #[test]
    fn inserted_leading_page_preserves_semantic_page_ids_and_moves_surfaces() {
        let mut next_id = 0;
        let (_, snapshot) = encode_frame_delta(
            &list_pages(&[("A", "first"), ("B", "second")]),
            &[],
            FrameEpochs {
                doc_epoch: 1,
                layout_epoch: 1,
                frame_epoch: 1,
                base_frame_epoch: 0,
            },
            true,
            &mut next_id,
        )
        .unwrap();

        let (delta, next) = encode_frame_delta(
            &list_pages(&[("X", "inserted"), ("A", "first"), ("B", "second")]),
            &snapshot,
            FrameEpochs {
                doc_epoch: 2,
                layout_epoch: 2,
                frame_epoch: 2,
                base_frame_epoch: 1,
            },
            false,
            &mut next_id,
        )
        .unwrap();

        assert_eq!(next[0].page_id, 3, "new leading page receives a new id");
        assert_eq!(next[1].page_id, snapshot[0].page_id);
        assert_eq!(next[2].page_id, snapshot[1].page_id);
        assert_eq!(u32_at(&delta, 52), 3);
        assert_eq!(delta[FRAME_HEADER_LEN], PAGE_OP_UPSERT);
        assert_eq!(delta[FRAME_HEADER_LEN + PAGE_OP_LEN], PAGE_OP_MOVE);
        assert_eq!(delta[FRAME_HEADER_LEN + PAGE_OP_LEN * 2], PAGE_OP_MOVE);
    }
}
