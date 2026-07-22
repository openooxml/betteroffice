//! Binary `FrameDelta` v1 encoder for the resident display list.
//!
//! The wire format is deliberately independent of wasm-bindgen and JSON. A
//! fixed header and fixed-size page-operation table point at aligned primitive
//! id arrays and a compact, typed value stream. Containers carry both element
//! counts and byte lengths; the browser decoder rejects any mismatch before a
//! page reaches canvas replay.

use std::collections::{BTreeSet, HashMap, HashSet};
use std::ops::Range;
use std::rc::Rc;

use docx_layout::display_list::{DisplayList, DisplayPage, DocAttrs, Primitive};
#[cfg(test)]
use serde_json::Value;

mod typed_page;
use typed_page::{collect_page_strings, encode_page, hash_page};

pub const FRAME_DELTA_VERSION: u16 = 1;
pub const FRAME_HEADER_LEN: usize = 80;
pub const PAGE_OP_LEN: usize = 48;
pub const FRAME_FLAG_FULL: u32 = 1;
pub const PAGE_OP_UPSERT: u8 = 1;
pub const PAGE_OP_REMOVE: u8 = 2;
pub const PAGE_OP_MOVE: u8 = 3;
pub const PAGE_OP_PATCH_POSITIONS: u8 = 4;
pub const PAGE_OP_SHIFT_POSITIONS: u8 = 5;

const POSITION_DOC_START: u8 = 1 << 0;
const POSITION_DOC_END: u8 = 1 << 1;
const POSITION_FRAGMENT_START: u8 = 1 << 2;
const POSITION_FRAGMENT_END: u8 = 1 << 3;
const POSITION_INLINE_WIDGET: u8 = 1 << 4;
const POSITION_FIELDS: [u8; 5] = [
    POSITION_DOC_START,
    POSITION_DOC_END,
    POSITION_FRAGMENT_START,
    POSITION_FRAGMENT_END,
    POSITION_INLINE_WIDGET,
];

const MAGIC: [u8; 4] = *b"FDV1";
const MAX_U32: usize = u32::MAX as usize;
const FNV_OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
const FNV_PRIME: u64 = 0x0000_0100_0000_01b3;

#[derive(Clone, Debug, Eq, PartialEq)]
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

#[derive(Debug)]
struct PreparedPage<'a> {
    snapshot: FramePageSnapshot,
    page: &'a DisplayPage,
    is_new: bool,
    moved: bool,
}

#[derive(Debug)]
enum PageOp<'a, 'b> {
    Upsert(&'a PreparedPage<'b>),
    Remove(&'a FramePageSnapshot),
    Move(&'a PreparedPage<'b>),
    PatchPositions(&'a PreparedPage<'b>, Vec<PositionPatch>),
    ShiftPositions(&'a PreparedPage<'b>, Vec<PositionShiftRun>),
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
    rebuilt_pages: Range<usize>,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    encode_frame_delta_inner(
        list,
        previous,
        epochs,
        false,
        next_page_id,
        Some(rebuilt_pages),
    )
}

fn encode_frame_delta_inner(
    list: &DisplayList,
    previous: &[FramePageSnapshot],
    epochs: FrameEpochs,
    full: bool,
    next_page_id: &mut u64,
    rebuilt_pages: Option<Range<usize>>,
) -> Result<(Vec<u8>, Vec<FramePageSnapshot>), String> {
    let prepared = prepare_pages(list, previous, next_page_id, rebuilt_pages.as_ref())?;
    let next_ids: HashSet<u64> = prepared.iter().map(|page| page.snapshot.page_id).collect();
    let previous_by_id: HashMap<u64, &FramePageSnapshot> =
        previous.iter().map(|old| (old.page_id, old)).collect();

    let mut ops = Vec::new();
    if full {
        ops.extend(prepared.iter().map(PageOp::Upsert));
    } else {
        for old in previous {
            if !next_ids.contains(&old.page_id) {
                ops.push(PageOp::Remove(old));
            }
        }
        for page in &prepared {
            let old = previous_by_id.get(&page.snapshot.page_id).copied();
            if page.is_new || old.is_none() {
                ops.push(PageOp::Upsert(page));
            } else if let Some(old) = old
                && old.fingerprint != page.snapshot.fingerprint
            {
                if old.visual_fingerprint == page.snapshot.visual_fingerprint
                    && old.primitive_ids == page.snapshot.primitive_ids
                {
                    let patches = position_patches(old, &page.snapshot);
                    if patches.is_empty() {
                        ops.push(PageOp::Upsert(page));
                    } else if let Some(runs) = position_shift_runs(old, &page.snapshot) {
                        ops.push(PageOp::ShiftPositions(page, runs));
                    } else {
                        ops.push(PageOp::PatchPositions(page, patches));
                    }
                } else {
                    ops.push(PageOp::Upsert(page));
                }
            } else if page.moved {
                ops.push(PageOp::Move(page));
            }
        }
    }

    let mut strings = BTreeSet::new();
    for op in &ops {
        if let PageOp::Upsert(page) = op {
            collect_page_strings(page.page, &mut strings)?;
        }
    }
    let strings: Vec<String> = strings.into_iter().collect();
    let string_ids: HashMap<&str, u32> = strings
        .iter()
        .enumerate()
        .map(|(index, value)| {
            u32::try_from(index)
                .map(|id| (value.as_str(), id))
                .map_err(|_| "FrameDelta string table exceeds u32")
        })
        .collect::<Result<_, _>>()?;

    let op_count = checked_u32(ops.len(), "page operation count")?;
    let page_count = checked_u32(list.pages.len(), "page count")?;
    let ops_bytes = ops
        .len()
        .checked_mul(PAGE_OP_LEN)
        .ok_or_else(|| "FrameDelta operation table overflow".to_owned())?;
    let strings_offset = FRAME_HEADER_LEN
        .checked_add(ops_bytes)
        .ok_or_else(|| "FrameDelta header overflow".to_owned())?;

    let mut out = vec![0; strings_offset];
    write_u32(&mut out, checked_u32(strings.len(), "string count")?);
    for value in &strings {
        write_u32(&mut out, checked_u32(value.len(), "string byte length")?);
        out.extend_from_slice(value.as_bytes());
    }
    let strings_len = out.len() - strings_offset;
    align(&mut out, 8);
    let data_offset = out.len();

    for (op_index, op) in ops.iter().enumerate() {
        let record = FRAME_HEADER_LEN + op_index * PAGE_OP_LEN;
        match op {
            PageOp::Upsert(page) => {
                out[record] = PAGE_OP_UPSERT;
                patch_u32(&mut out, record + 4, page.snapshot.page_index);
                patch_u64(&mut out, record + 8, page.snapshot.page_id);
                patch_u64(&mut out, record + 16, page.snapshot.fingerprint);
                patch_u32(
                    &mut out,
                    record + 24,
                    checked_u32(page.snapshot.primitive_ids.len(), "primitive id count")?,
                );

                align(&mut out, 8);
                let primitive_id_offset = checked_u32(out.len(), "primitive id offset")?;
                patch_u32(&mut out, record + 28, primitive_id_offset);
                for id in page.snapshot.primitive_ids.iter() {
                    write_u64(&mut out, *id);
                }

                let payload_offset = out.len();
                encode_page(page.page, &string_ids, &mut out)?;
                let payload_len = out.len() - payload_offset;
                patch_u32(
                    &mut out,
                    record + 32,
                    checked_u32(payload_offset, "page payload offset")?,
                );
                patch_u32(
                    &mut out,
                    record + 36,
                    checked_u32(payload_len, "page payload length")?,
                );
            }
            PageOp::Remove(page) => {
                out[record] = PAGE_OP_REMOVE;
                patch_u32(&mut out, record + 4, page.page_index);
                patch_u64(&mut out, record + 8, page.page_id);
            }
            PageOp::Move(page) => {
                out[record] = PAGE_OP_MOVE;
                patch_u32(&mut out, record + 4, page.snapshot.page_index);
                patch_u64(&mut out, record + 8, page.snapshot.page_id);
                patch_u64(&mut out, record + 16, page.snapshot.fingerprint);
            }
            PageOp::PatchPositions(page, patches) => {
                out[record] = PAGE_OP_PATCH_POSITIONS;
                patch_u32(&mut out, record + 4, page.snapshot.page_index);
                patch_u64(&mut out, record + 8, page.snapshot.page_id);
                patch_u64(&mut out, record + 16, page.snapshot.fingerprint);
                patch_u32(
                    &mut out,
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
                for patch in patches {
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
                patch_u32(
                    &mut out,
                    record + 32,
                    checked_u32(payload_offset, "position patch payload offset")?,
                );
                let payload_length = out.len() - payload_offset;
                patch_u32(
                    &mut out,
                    record + 36,
                    checked_u32(payload_length, "position patch payload length")?,
                );
            }
            PageOp::ShiftPositions(page, runs) => {
                out[record] = PAGE_OP_SHIFT_POSITIONS;
                patch_u32(&mut out, record + 4, page.snapshot.page_index);
                patch_u64(&mut out, record + 8, page.snapshot.page_id);
                patch_u64(&mut out, record + 16, page.snapshot.fingerprint);
                patch_u32(
                    &mut out,
                    record + 24,
                    checked_u32(runs.len(), "position shift run count")?,
                );
                align(&mut out, 8);
                let payload_offset = out.len();
                write_u32(
                    &mut out,
                    checked_u32(runs.len(), "position shift run count")?,
                );
                write_u32(&mut out, 0);
                for run in runs {
                    write_u32(&mut out, run.start);
                    write_u32(&mut out, run.count);
                    out.push(run.changed_mask);
                    out.extend_from_slice(&[0; 7]);
                    write_i64(&mut out, run.delta);
                }
                patch_u32(
                    &mut out,
                    record + 32,
                    checked_u32(payload_offset, "position shift payload offset")?,
                );
                let payload_length = out.len() - payload_offset;
                patch_u32(
                    &mut out,
                    record + 36,
                    checked_u32(payload_length, "position shift payload length")?,
                );
            }
        }
    }

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

    drop(ops);
    let next_snapshots = prepared.into_iter().map(|page| page.snapshot).collect();
    Ok((out, next_snapshots))
}

fn prepare_pages<'a>(
    list: &'a DisplayList,
    previous: &[FramePageSnapshot],
    next_page_id: &mut u64,
    rebuilt_pages: Option<&Range<usize>>,
) -> Result<Vec<PreparedPage<'a>>, String> {
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
    for (next_index, anchor) in anchors.iter().enumerate() {
        if let Some(previous_index) = previous_by_anchor.remove(anchor.as_str()) {
            claimed.insert(previous[previous_index].page_id);
            matched_previous[next_index] = Some(previous_index);
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

    let mut prepared = Vec::with_capacity(list.pages.len());
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
        let positions = primitive_positions(page);
        let full_prepare =
            is_new || rebuilt_pages.is_none_or(|rebuilt_pages| rebuilt_pages.contains(&index));
        let (fingerprint, visual_fingerprint, primitive_ids) = if full_prepare {
            let hashes = hash_page(page)?;
            let primitive_ids: Rc<[u64]> = primitive_ids(page, page_id).into();
            (hashes.fingerprint, hashes.visual_fingerprint, primitive_ids)
        } else {
            let old = matched.expect("clean incremental pages retain a previous snapshot");
            let fingerprint = if positions == old.positions {
                old.fingerprint
            } else {
                hash_positions(old.visual_fingerprint, &positions)
            };
            (
                fingerprint,
                old.visual_fingerprint,
                Rc::clone(&old.primitive_ids),
            )
        };
        prepared.push(PreparedPage {
            snapshot: FramePageSnapshot {
                page_id,
                anchor,
                fingerprint,
                visual_fingerprint,
                page_index,
                primitive_ids,
                positions,
            },
            page,
            is_new,
            moved,
        });
    }
    Ok(prepared)
}

fn hash_positions(visual_fingerprint: u64, positions: &[PrimitivePositionSnapshot]) -> u64 {
    let mut hash = visual_fingerprint;
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
    previous: &FramePageSnapshot,
    next: &FramePageSnapshot,
) -> Option<Vec<PositionShiftRun>> {
    if previous.positions.len() != next.positions.len() {
        return None;
    }
    let mut runs: Vec<PositionShiftRun> = Vec::new();
    for (index, (previous, next)) in previous.positions.iter().zip(&next.positions).enumerate() {
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
        let mut common_delta = None;
        for (field_index, field) in POSITION_FIELDS.iter().enumerate() {
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
        let Some(delta) = common_delta else {
            continue;
        };
        if delta == 0 {
            return None;
        }
        let index = checked_u32(index, "position shift primitive index").ok()?;
        if let Some(last) = runs.last_mut()
            && last.start.checked_add(last.count) == Some(index)
            && last.changed_mask == changed_mask
            && last.delta == delta
        {
            last.count = last.count.checked_add(1)?;
        } else {
            runs.push(PositionShiftRun {
                start: index,
                count: 1,
                changed_mask,
                delta,
            });
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
fn collect_strings(value: &Value, strings: &mut BTreeSet<String>) {
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

fn string_id(ids: &HashMap<&str, u32>, value: &str) -> Result<u32, String> {
    ids.get(value)
        .copied()
        .ok_or_else(|| "FrameDelta string table missed a value".to_owned())
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

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::DisplayList;

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
    fn decode_typed(bytes: &[u8], cursor: &mut usize, strings: &[String]) -> Value {
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

    #[test]
    fn typed_emission_matches_the_value_reference() {
        for list in [list("hello"), list_at_position(2), rich_list()] {
            for page in &list.pages {
                let value = serde_json::to_value(page).unwrap();

                let mut typed_strings = BTreeSet::new();
                collect_page_strings(page, &mut typed_strings).unwrap();
                let mut reference_strings = BTreeSet::new();
                collect_strings(&value, &mut reference_strings);
                assert_eq!(typed_strings, reference_strings, "string tables differ");

                let strings: Vec<String> = typed_strings.into_iter().collect();
                let ids: HashMap<&str, u32> = strings
                    .iter()
                    .enumerate()
                    .map(|(index, value)| (value.as_str(), index as u32))
                    .collect();
                let mut out = Vec::new();
                encode_page(page, &ids, &mut out).unwrap();
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
        let mut typed_strings = BTreeSet::new();
        collect_page_strings(page, &mut typed_strings).unwrap();
        let strings: Vec<String> = typed_strings.into_iter().collect();
        let ids: HashMap<&str, u32> = strings
            .iter()
            .enumerate()
            .map(|(index, value)| (value.as_str(), index as u32))
            .collect();
        let mut out = Vec::new();
        encode_page(page, &ids, &mut out).unwrap();
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
            POSITION_DOC_START | POSITION_DOC_END | POSITION_FRAGMENT_START | POSITION_FRAGMENT_END
        );
        assert_eq!(
            i64::from_le_bytes(delta[payload + 24..payload + 32].try_into().unwrap()),
            10
        );
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
