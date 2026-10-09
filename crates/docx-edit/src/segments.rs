// Segment/paragraph accessors used only by the wasm helpers read as dead code in
// native builds; the invalidation and paragraph paths below are live on both.
#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use yrs::types::text::YChange;
use yrs::{Any, Out, ReadTxn, Text, TextRef};

use crate::{KIND_KEY, PARA_ID, is_pilcrow, map_string, out_len};

/// Whether the layout gives an embed its own block.
pub(crate) fn is_block_embed(kind: &str) -> bool {
    matches!(kind, "table" | "blockSdt" | "pageBreak" | "columnBreak")
}

/// One paragraph's resolved geometry inside a story.
#[derive(Clone)]
pub(crate) struct ParaEntry {
    pub para_id: Box<str>,
    /// Story index of the paragraph's first unit (after the previous pilcrow).
    pub start: u32,
    /// Story index of the paragraph's own pilcrow embed.
    pub pilcrow: u32,
    /// First content unit after the paragraph's leading block embeds.
    pub node_start: u32,
}

#[derive(Clone)]
pub(crate) enum SegKind {
    Text(Arc<str>),
    Pilcrow,
    Embed,
}

#[derive(Clone)]
pub(crate) struct Seg {
    pub start: u32,
    pub kind: SegKind,
}

/// Text segments may split anywhere; cold builds follow formatting runs, advanced indexes preserve only per-unit kind and text.
#[derive(Clone)]
pub(crate) struct SegmentIndex {
    len: u32,
    segs: Vec<Seg>,
}

#[derive(Clone)]
pub(crate) struct ParagraphIndex {
    paras: Vec<ParaEntry>,
    by_para: HashMap<Box<str>, u32>,
    /// Paragraph ids more than one paragraph of the story carries.
    repeated: HashSet<Box<str>>,
}

pub(crate) fn build_indexes<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
) -> (SegmentIndex, ParagraphIndex) {
    let mut len = 0_u32;
    let mut para_start = 0_u32;
    let mut node_start = 0_u32;
    let mut segs = Vec::new();
    let mut paras: Vec<ParaEntry> = Vec::new();
    let mut by_para = HashMap::new();
    let mut repeated = HashSet::new();
    for diff in story.diff(txn, YChange::identity) {
        let units = out_len(&diff.insert);
        let kind = match diff.insert {
            Out::Any(Any::String(text)) => SegKind::Text(text),
            Out::YMap(map) if is_pilcrow(&map, txn) => {
                let para_id: Box<str> = map_string(&map, txn, PARA_ID)
                    .unwrap_or_default()
                    .into_boxed_str();
                let slot = paras.len() as u32;
                if by_para.contains_key(&para_id) {
                    repeated.insert(para_id.clone());
                } else {
                    by_para.insert(para_id.clone(), slot);
                }
                paras.push(ParaEntry {
                    para_id,
                    start: para_start,
                    pilcrow: len,
                    node_start,
                });
                para_start = len + 1;
                node_start = len + 1;
                SegKind::Pilcrow
            }
            insert => {
                let kind = match insert {
                    Out::YMap(map) => map_string(&map, txn, KIND_KEY).unwrap_or_default(),
                    _ => String::new(),
                };
                let block = is_block_embed(&kind);
                if len == node_start && block {
                    node_start = len + 1;
                }
                SegKind::Embed
            }
        };
        if units == 0 {
            continue;
        }
        segs.push(Seg { start: len, kind });
        len += units;
    }
    (
        SegmentIndex { len, segs },
        ParagraphIndex {
            paras,
            by_para,
            repeated,
        },
    )
}

impl ParagraphIndex {
    /// `(start, pilcrow)` span of `para_id`, matching the first paragraph with that id.
    pub(crate) fn para_span(&self, para_id: &str) -> Option<(u32, u32)> {
        let para = self.paras.get(*self.by_para.get(para_id)? as usize)?;
        Some((para.start, para.pilcrow))
    }

    /// How many paragraphs carry `para_id`: 0, 1, or 2 for two or more.
    pub(crate) fn para_id_count(&self, para_id: &str) -> u32 {
        if self.repeated.contains(para_id) {
            2
        } else {
            u32::from(self.by_para.contains_key(para_id))
        }
    }

    /// First paragraph whose pilcrow sits at or after `index` — the paragraph `index`
    /// resolves into.
    pub(crate) fn para_at(&self, index: u32) -> Option<&ParaEntry> {
        self.paras
            .get(self.paras.partition_point(|para| para.pilcrow < index))
    }

    pub(crate) fn shift_for_text_insert(&mut self, index: u32, units: u32) -> bool {
        let slot = self.paras.partition_point(|para| para.pilcrow < index);
        let Some(para) = self.paras.get_mut(slot) else {
            return false;
        };
        if index < para.node_start {
            return false;
        }
        para.pilcrow += units;
        for para in &mut self.paras[slot + 1..] {
            para.start += units;
            para.pilcrow += units;
            para.node_start += units;
        }
        true
    }

    /// Shifts paragraph positions after deleting content within one paragraph.
    pub(crate) fn shift_for_text_delete(&mut self, start: u32, end: u32) -> bool {
        if start >= end {
            return false;
        }
        let slot = self.paras.partition_point(|para| para.pilcrow < start);
        let Some(para) = self.paras.get_mut(slot) else {
            return false;
        };
        if start < para.node_start || end > para.pilcrow {
            return false;
        }
        let units = end - start;
        para.pilcrow -= units;
        for para in &mut self.paras[slot + 1..] {
            para.start -= units;
            para.pilcrow -= units;
            para.node_start -= units;
        }
        true
    }
}

impl SegmentIndex {
    pub(crate) fn shift_for_text_insert(&mut self, index: u32, text: &str) -> bool {
        let units = text.encode_utf16().count() as u32;
        if text.is_empty() || index > self.len || self.len.checked_add(units).is_none() {
            return false;
        }
        let slot = self.segs.partition_point(|seg| seg.start < index);
        let text_slot = slot
            .checked_sub(1)
            .filter(|&slot| matches!(self.segs[slot].kind, SegKind::Text(_)))
            .or_else(|| {
                self.segs.get(slot).and_then(|seg| {
                    (seg.start == index && matches!(seg.kind, SegKind::Text(_))).then_some(slot)
                })
            });
        let shifted_from = if let Some(slot) = text_slot {
            let seg = &mut self.segs[slot];
            let SegKind::Text(existing) = &seg.kind else {
                return false;
            };
            let Some(offset) = utf16_byte_offset(existing, index - seg.start) else {
                return false;
            };
            let mut replacement = existing.to_string();
            replacement.insert_str(offset, text);
            seg.kind = SegKind::Text(replacement.into());
            slot + 1
        } else {
            self.segs.insert(
                slot,
                Seg {
                    start: index,
                    kind: SegKind::Text(text.into()),
                },
            );
            slot + 1
        };
        for seg in &mut self.segs[shifted_from..] {
            seg.start += units;
        }
        self.len += units;
        true
    }

    pub(crate) fn shift_for_text_delete(&mut self, start: u32, end: u32) -> bool {
        if !self.is_text_range(start, end) {
            return false;
        }
        let first = self.segs.partition_point(|seg| seg.start <= start) - 1;
        let last = self.segs.partition_point(|seg| seg.start < end);
        let mut cuts = Vec::with_capacity(last - first);
        for slot in first..last {
            let seg = &self.segs[slot];
            let SegKind::Text(text) = &seg.kind else {
                return false;
            };
            let seg_end = self.segs.get(slot + 1).map_or(self.len, |seg| seg.start);
            let Some(from) = utf16_byte_offset(text, start.max(seg.start) - seg.start) else {
                return false;
            };
            let Some(to) = utf16_byte_offset(text, end.min(seg_end) - seg.start) else {
                return false;
            };
            cuts.push((slot, from, to));
        }
        for (slot, from, to) in cuts {
            let seg = &mut self.segs[slot];
            if let SegKind::Text(text) = &seg.kind {
                let mut replacement = text.to_string();
                replacement.replace_range(from..to, "");
                seg.kind = SegKind::Text(replacement.into());
            }
            seg.start = seg.start.min(start);
        }
        let units = end - start;
        for seg in &mut self.segs[last..] {
            seg.start -= units;
        }
        let mut slot = first;
        for _ in first..last {
            if matches!(&self.segs[slot].kind, SegKind::Text(text) if text.is_empty()) {
                self.segs.remove(slot);
            } else {
                slot += 1;
            }
        }
        self.len -= units;
        true
    }

    /// Whether a nonempty range contains only plain text units and splits no surrogate pair.
    pub(crate) fn is_text_range(&self, start: u32, end: u32) -> bool {
        if start >= end || end > self.len {
            return false;
        }
        let slot = self.segs.partition_point(|seg| seg.start <= start);
        slot > 0
            && self.segs[slot - 1..]
                .iter()
                .take_while(|seg| seg.start < end)
                .all(|seg| matches!(seg.kind, SegKind::Text(_)))
            && self.is_char_boundary(start)
            && self.is_char_boundary(end)
    }

    /// Whether `pos` does not fall between the two units of a surrogate pair.
    pub(crate) fn is_char_boundary(&self, pos: u32) -> bool {
        match self.segment_at(pos) {
            Some(Seg {
                start,
                kind: SegKind::Text(text),
            }) => utf16_byte_offset(text, pos - start).is_some(),
            _ => true,
        }
    }

    /// The segment covering `pos`, if any.
    pub(crate) fn segment_at(&self, pos: u32) -> Option<&Seg> {
        if pos >= self.len {
            return None;
        }
        self.segs
            .get(self.segs.partition_point(|seg| seg.start <= pos) - 1)
    }
}

fn utf16_byte_offset(text: &str, offset: u32) -> Option<usize> {
    let mut units = 0;
    for (byte, ch) in text.char_indices() {
        if units == offset {
            return Some(byte);
        }
        units += ch.len_utf16() as u32;
        if units > offset {
            return None;
        }
    }
    (units == offset).then_some(text.len())
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::{EditCtx, EditingDoc, FormatPolicy, Position, RawOp, SegmentContent, StoryRange};
    use yrs::{Map, Transact};

    #[derive(Debug, PartialEq, Eq)]
    pub(crate) enum Unit {
        Text(u16),
        Pilcrow,
        Embed,
    }

    pub(crate) fn units(index: &SegmentIndex) -> Vec<Unit> {
        let mut result = Vec::new();
        for seg in &index.segs {
            match &seg.kind {
                SegKind::Text(text) => result.extend(text.encode_utf16().map(Unit::Text)),
                SegKind::Pilcrow => result.push(Unit::Pilcrow),
                SegKind::Embed => result.push(Unit::Embed),
            }
        }
        result
    }

    pub(crate) fn assert_segment_invariants(index: &SegmentIndex) {
        let mut end = 0;
        let mut previous = None;
        for seg in &index.segs {
            assert_eq!(seg.start, end);
            assert!(previous.is_none_or(|start| start < seg.start));
            let len = match &seg.kind {
                SegKind::Text(text) => text.encode_utf16().count() as u32,
                _ => 1,
            };
            assert!(len > 0);
            previous = Some(seg.start);
            end += len;
        }
        assert_eq!(end, index.len);
    }

    fn assert_segment_indexes_unchanged(actual: &SegmentIndex, expected: &SegmentIndex) {
        assert_eq!(actual.len, expected.len);
        assert_eq!(actual.segs.len(), expected.segs.len());
        for (actual, expected) in actual.segs.iter().zip(&expected.segs) {
            assert_eq!(actual.start, expected.start);
            match (&actual.kind, &expected.kind) {
                (SegKind::Text(actual), SegKind::Text(expected)) => assert_eq!(actual, expected),
                (SegKind::Pilcrow, SegKind::Pilcrow) | (SegKind::Embed, SegKind::Embed) => {}
                _ => panic!("segment kind changed"),
            }
        }
    }

    pub(crate) fn next_random(state: &mut u64) -> u32 {
        *state ^= *state << 13;
        *state ^= *state >> 7;
        *state ^= *state << 17;
        *state as u32
    }

    pub(crate) fn random_text(state: &mut u64) -> &'static str {
        let texts = ["x", "é", "😀", "ab", "xé", "é😀x"];
        texts[next_random(state) as usize % texts.len()]
    }

    /// The pre-index paragraph walk, kept as the reference oracle.
    fn reference_para_at(doc: &EditingDoc, story: &str, index: u32) -> Option<(String, u32, u32)> {
        let mut cursor = 0_u32;
        let mut para_start = 0_u32;
        let mut node_start = 0_u32;
        for segment in doc.story_segments(story).unwrap() {
            match segment.content {
                SegmentContent::Text(text) => cursor += text.encode_utf16().count() as u32,
                SegmentContent::Pilcrow(properties) => {
                    if index <= cursor {
                        return Some((
                            properties.para_id,
                            index.saturating_sub(para_start),
                            index.saturating_sub(node_start),
                        ));
                    }
                    cursor += 1;
                    para_start = cursor;
                    node_start = cursor;
                }
                SegmentContent::OtherEmbed { ref kind, .. } => {
                    if cursor == node_start && is_block_embed(kind) {
                        node_start = cursor + 1;
                    }
                    cursor += 1;
                }
            }
        }
        None
    }

    /// The pre-index unit classification (kind only; text widths stay in the caller).
    fn reference_seg_kind(doc: &EditingDoc, story: &str, pos: u32) -> Option<&'static str> {
        let mut cursor = 0_u32;
        for segment in doc.story_segments(story).unwrap() {
            let units = match &segment.content {
                SegmentContent::Text(text) => text.encode_utf16().count() as u32,
                _ => 1,
            };
            if pos >= cursor && pos < cursor + units {
                return Some(match segment.content {
                    SegmentContent::Text(_) => "text",
                    SegmentContent::Pilcrow(_) => "pilcrow",
                    SegmentContent::OtherEmbed { .. } => "embed",
                });
            }
            cursor += units;
        }
        None
    }

    fn assert_index_matches_segments(doc: &EditingDoc, story: &str) {
        let index = doc.segment_index(story).unwrap();
        let paragraphs = doc.paragraph_index(story).unwrap();
        let len = doc.story_len(story).unwrap();
        for pos in 0..len {
            let expected = reference_seg_kind(doc, story, pos);
            let seg = index.segment_at(pos).unwrap();
            let actual = match seg.kind {
                SegKind::Text(_) => "text",
                SegKind::Pilcrow => "pilcrow",
                SegKind::Embed => "embed",
            };
            assert_eq!(expected, Some(actual), "pos {pos}");
        }
        assert!(index.segment_at(len).is_none());
        let para_count = doc.paragraphs(story).unwrap().len() as u32;
        for index_pos in 0..=len {
            let expected = reference_para_at(doc, story, index_pos);
            let actual = paragraphs.para_at(index_pos).map(|para| {
                (
                    para.para_id.to_string(),
                    index_pos.saturating_sub(para.start),
                    index_pos.saturating_sub(para.node_start),
                )
            });
            assert_eq!(expected, actual, "index {index_pos}");
        }
        assert_eq!(paragraphs.paras.len() as u32, para_count);
        for paragraph in doc.paragraphs(story).unwrap() {
            let span = paragraphs.para_span(&paragraph.para_id).unwrap();
            let reference = {
                // The pre-index para-span walk, kept as the reference oracle.
                let mut offset = 0_u32;
                let mut para_start = 0_u32;
                let mut found = (u32::MAX, u32::MAX);
                for segment in doc.story_segments(story).unwrap() {
                    match segment.content {
                        SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
                        SegmentContent::Pilcrow(properties) => {
                            if properties.para_id == paragraph.para_id {
                                found = (para_start, offset);
                                break;
                            }
                            offset += 1;
                            para_start = offset;
                        }
                        SegmentContent::OtherEmbed { .. } => offset += 1,
                    }
                }
                found
            };
            assert_eq!(span, reference, "para {}", paragraph.para_id);
        }
    }

    /// `A B [sdt] pilcrow(p1)` | `pilcrow(p2)` (empty paragraph) | `[table] [pageBreak] C pilcrow(p3)`
    fn seeded_doc() -> EditingDoc {
        let doc = EditingDoc::new(7);
        seed_story(&doc);
        doc
    }

    fn seed_story(doc: &EditingDoc) {
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::InsertEmbed {
                    index: 2,
                    kind: "sdt".into(),
                    payload: vec![("embedId".into(), Any::from("control-1"))],
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 4,
                    kind: "pilcrow".into(),
                    payload: vec![("paraId".into(), Any::from("p-2"))],
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 5,
                    kind: "table".into(),
                    payload: vec![("embedId".into(), Any::from("t-1"))],
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 6,
                    kind: "pageBreak".into(),
                    payload: vec![("embedId".into(), Any::from("p-1"))],
                    attrs: Default::default(),
                },
                RawOp::Insert {
                    index: 7,
                    text: "C".into(),
                    attrs: Default::default(),
                },
                RawOp::InsertEmbed {
                    index: 8,
                    kind: "pilcrow".into(),
                    payload: vec![("paraId".into(), Any::from("p-3"))],
                    attrs: Default::default(),
                },
            ],
            &EditCtx::local(String::new(), String::new()),
        )
        .unwrap();
    }

    pub(crate) fn seed_text_stream(doc: &EditingDoc) {
        seed_story(doc);
        doc.apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: 1,
                text: "é😀Bold".into(),
                attrs: yrs::types::Attrs::from([("bold".into(), Any::Bool(true))]),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    }

    fn plain_doc_with_repeated_ids() -> EditingDoc {
        let doc = EditingDoc::new(11);
        doc.create_story("body", "AlphaBravoCharlie", "Normal", "left")
            .unwrap();
        let ctx = EditCtx::local("", "");
        for index in [5, 11] {
            doc.split_paragraph(&ctx, Position::new("body", index), None)
                .unwrap();
        }
        let mut txn = doc.yrs_doc().transact_mut();
        let story = crate::story_ref(&txn, "body").unwrap();
        let (_, second) = crate::next_pilcrow(&story, &txn, 6).unwrap();
        let (_, last) = crate::next_pilcrow(&story, &txn, 12).unwrap();
        let para_id = map_string(&second, &txn, PARA_ID).unwrap();
        last.insert(&mut txn, PARA_ID, para_id);
        drop(txn);
        doc
    }

    fn fresh_paragraph_index(doc: &EditingDoc, story: &str) -> ParagraphIndex {
        let txn = doc.yrs_doc().transact();
        let story = crate::story_ref(&txn, story).unwrap();
        build_indexes(&story, &txn).1
    }

    pub(crate) fn assert_paragraph_indexes_eq(actual: &ParagraphIndex, expected: &ParagraphIndex) {
        assert_eq!(actual.paras.len(), expected.paras.len());
        for (slot, (actual, expected)) in actual.paras.iter().zip(&expected.paras).enumerate() {
            assert_eq!(
                (
                    actual.start,
                    actual.pilcrow,
                    actual.node_start,
                    actual.para_id.as_ref(),
                ),
                (
                    expected.start,
                    expected.pilcrow,
                    expected.node_start,
                    expected.para_id.as_ref(),
                ),
                "paragraph {slot}"
            );
        }
        assert_eq!(actual.by_para, expected.by_para);
        assert_eq!(actual.repeated, expected.repeated);
    }

    #[test]
    fn segment_index_shifts_preserve_units_and_reject_surrogate_cuts() {
        let original = SegmentIndex {
            len: 12,
            segs: vec![
                Seg {
                    start: 0,
                    kind: SegKind::Text("ab".into()),
                },
                Seg {
                    start: 2,
                    kind: SegKind::Text("é😀z".into()),
                },
                Seg {
                    start: 6,
                    kind: SegKind::Pilcrow,
                },
                Seg {
                    start: 7,
                    kind: SegKind::Embed,
                },
                Seg {
                    start: 8,
                    kind: SegKind::Text("tail".into()),
                },
            ],
        };
        assert_segment_invariants(&original);
        for index in 0..=original.len {
            let mut shifted = original.clone();
            if index == 4 {
                assert!(!shifted.shift_for_text_insert(index, "xé😀"));
                assert_segment_indexes_unchanged(&shifted, &original);
                continue;
            }
            assert!(shifted.shift_for_text_insert(index, "xé😀"));
            let mut expected = units(&original);
            expected.splice(
                index as usize..index as usize,
                "xé😀".encode_utf16().map(Unit::Text),
            );
            assert_eq!(units(&shifted), expected, "insert {index}");
            assert_segment_invariants(&shifted);
        }
        for start in 0..=original.len {
            for end in start..=original.len + 1 {
                let mut shifted = original.clone();
                let eligible = original.is_text_range(start, end);
                assert!(
                    !eligible || (start != 4 && end != 4),
                    "range {start}..{end}"
                );
                assert_eq!(
                    shifted.shift_for_text_delete(start, end),
                    eligible,
                    "delete {start}..{end}"
                );
                if eligible {
                    let mut expected = units(&original);
                    expected.drain(start as usize..end as usize);
                    assert_eq!(units(&shifted), expected, "delete {start}..{end}");
                    assert_segment_invariants(&shifted);
                } else {
                    assert_segment_indexes_unchanged(&shifted, &original);
                }
            }
        }
        assert!(!original.is_char_boundary(4));
        assert!(!original.is_text_range(2, 4) && !original.is_text_range(4, 6));
        assert!(original.is_text_range(2, 6) && original.is_text_range(3, 5));
        for (index, text) in [(0, ""), (original.len + 1, "x")] {
            let mut shifted = original.clone();
            assert!(!shifted.shift_for_text_insert(index, text));
            assert_segment_indexes_unchanged(&shifted, &original);
        }
        let mut empty = SegmentIndex {
            len: 0,
            segs: Vec::new(),
        };
        assert!(empty.shift_for_text_insert(0, "é😀"));
        assert_segment_invariants(&empty);
        assert!(empty.shift_for_text_delete(0, 3));
        assert!(empty.segs.is_empty());
        assert_segment_invariants(&empty);
    }

    fn run_text_stream(advance_cached: bool) {
        for seed in [1, 7, 42, 0x1234_5678] {
            let doc = EditingDoc::new(31);
            seed_text_stream(&doc);
            let mut maintained = (*doc.segment_index("body").unwrap()).clone();
            doc.paragraph_index("body").unwrap();
            let mut random = seed;
            let mut steps = 0;
            let mut shifts = 0;
            let mut cache_shifts = 0;
            for step in 0..200 {
                let before = doc.committed_epoch();
                let len = doc.story_len("body").unwrap();
                let insert = next_random(&mut random) % 3 != 0;
                let start = next_random(&mut random) % (len + u32::from(insert));
                let start = start - u32::from(!maintained.is_char_boundary(start));
                let ctx = EditCtx::local("", "");
                let advanced = if insert {
                    let text = random_text(&mut random);
                    if doc
                        .insert_text(
                            &ctx,
                            Position::new("body", start),
                            text,
                            FormatPolicy::Inherit,
                        )
                        .is_err()
                    {
                        continue;
                    }
                    let exact = doc.story_len("body").unwrap().checked_sub(len)
                        == Some(text.encode_utf16().count() as u32);
                    if advance_cached && exact {
                        doc.advance_indexes_after_text_insert(
                            "body",
                            before,
                            doc.committed_epoch(),
                            start,
                            text,
                        );
                    }
                    exact && maintained.shift_for_text_insert(start, text)
                } else {
                    let end = (start + 1 + next_random(&mut random) % 3).min(len);
                    let end = end + u32::from(!maintained.is_char_boundary(end));
                    let plain = maintained.is_text_range(start, end);
                    let paragraph_safe = maintained
                        .segment_at(end)
                        .is_none_or(|seg| !matches!(seg.kind, SegKind::Embed))
                        || doc
                            .paragraph_index("body")
                            .unwrap()
                            .para_at(start)
                            .is_some_and(|para| start > para.node_start);
                    if doc
                        .delete_range(&ctx, StoryRange::new("body", start, end))
                        .is_err()
                    {
                        continue;
                    }
                    let exact =
                        len.checked_sub(doc.story_len("body").unwrap()) == Some(end - start);
                    if advance_cached && plain && exact && paragraph_safe {
                        doc.advance_indexes_after_text_delete(
                            "body",
                            before,
                            doc.committed_epoch(),
                            start,
                            end,
                        );
                    }
                    plain && exact && maintained.shift_for_text_delete(start, end)
                };
                steps += 1;
                shifts += usize::from(advanced);
                let txn = doc.yrs_doc().transact();
                let story = crate::story_ref(&txn, "body").unwrap();
                let (cold_segments, cold_paragraphs) = build_indexes(&story, &txn);
                drop(txn);
                if !advanced {
                    maintained = cold_segments.clone();
                }
                assert_eq!(
                    units(&maintained),
                    units(&cold_segments),
                    "seed {seed}, step {step}"
                );
                assert_segment_invariants(&maintained);
                assert_segment_invariants(&cold_segments);
                let after = doc.committed_epoch();
                let cached_segments = doc.segment_indexes.lock().unwrap().get("body", after);
                cache_shifts += usize::from(advanced && cached_segments.is_some());
                if let Some(cached) = cached_segments {
                    assert_eq!(
                        units(&cached),
                        units(&cold_segments),
                        "seed {seed}, step {step}"
                    );
                    assert_segment_invariants(&cached);
                }
                if let Some(cached) = doc.paragraph_indexes.lock().unwrap().get("body", after) {
                    assert_paragraph_indexes_eq(&cached, &cold_paragraphs);
                }
                doc.segment_index("body").unwrap();
                doc.paragraph_index("body").unwrap();
            }
            assert!(steps > 150, "seed {seed}: {steps} accepted steps");
            assert!(
                shifts * 5 > steps * 2,
                "seed {seed}: {shifts}/{steps} shifts"
            );
            if advance_cached {
                assert!(
                    cache_shifts * 5 > steps * 2,
                    "seed {seed}: {cache_shifts}/{steps} cache shifts"
                );
            }
        }
    }

    #[test]
    fn segment_index_random_text_stream_matches_cold_builds() {
        run_text_stream(false);
    }

    #[test]
    fn cached_story_indexes_random_text_stream_matches_cold_builds() {
        run_text_stream(true);
    }

    #[test]
    fn paragraph_index_text_insert_matches_full_build_at_every_index() {
        for doc in [seeded_doc(), plain_doc_with_repeated_ids()] {
            let original = fresh_paragraph_index(&doc, "body");
            let state = doc.encode_state_as_update_v1();
            let ctx = EditCtx::local("", "");
            for index in 0..=doc.story_len("body").unwrap() {
                let mut shifted = original.clone();
                let Some(paragraph) = original.para_at(index) else {
                    assert!(!shifted.shift_for_text_insert(index, 2));
                    assert_paragraph_indexes_eq(&shifted, &original);
                    continue;
                };
                let edited = EditingDoc::new(13);
                edited.apply_verbatim_v1(&state).unwrap();
                edited
                    .insert_text(
                        &ctx,
                        Position::new("body", index),
                        "xy",
                        FormatPolicy::Inherit,
                    )
                    .unwrap();
                if shifted.shift_for_text_insert(index, 2) {
                    assert_paragraph_indexes_eq(&shifted, &fresh_paragraph_index(&edited, "body"));
                } else {
                    assert!(index < paragraph.node_start);
                    assert_paragraph_indexes_eq(&shifted, &original);
                }
            }
        }
    }

    /// Short text deletions preserve the full paragraph index geometry.
    #[test]
    fn paragraph_index_text_delete_matches_full_build_at_every_range() {
        for doc in [seeded_doc(), plain_doc_with_repeated_ids()] {
            let original = fresh_paragraph_index(&doc, "body");
            let state = doc.encode_state_as_update_v1();
            let ctx = EditCtx::local("", "");
            let len = doc.story_len("body").unwrap();
            for start in 0..len {
                for end in start + 1..=start.saturating_add(3).min(len) {
                    let mut shifted = original.clone();
                    let edited = EditingDoc::new(17);
                    edited.apply_verbatim_v1(&state).unwrap();
                    edited
                        .delete_range(&ctx, StoryRange::new("body", start, end))
                        .unwrap();
                    let eligible = original
                        .para_at(start)
                        .is_some_and(|para| para.node_start <= start && end <= para.pilcrow);
                    let advanced = shifted.shift_for_text_delete(start, end);
                    assert_eq!(advanced, eligible, "range {start}..{end}");
                    if advanced {
                        assert_paragraph_indexes_eq(
                            &shifted,
                            &fresh_paragraph_index(&edited, "body"),
                        );
                    } else {
                        assert_paragraph_indexes_eq(&shifted, &original);
                    }
                }
            }
            for (start, end) in [(0, 0), (2, 1), (len, len + 1)] {
                let mut shifted = original.clone();
                assert!(!shifted.shift_for_text_delete(start, end));
                assert_paragraph_indexes_eq(&shifted, &original);
            }
        }
    }

    #[test]
    fn story_indexes_advance_after_text_insert() {
        let doc = seeded_doc();
        let original = doc.paragraph_index("body").unwrap();
        let original_segments = doc.segment_index("body").unwrap();
        let expected_segments = units(&original_segments);
        let expected_original = fresh_paragraph_index(&doc, "body");
        let ctx = EditCtx::local("", "");
        let before = doc.committed_epoch();
        doc.insert_text(&ctx, Position::new("body", 1), "xy", FormatPolicy::Inherit)
            .unwrap();
        let after = doc.committed_epoch();
        assert_eq!(after, before + 1);
        doc.advance_indexes_after_text_insert("body", before, after, 1, "xy");
        let shifted = doc
            .paragraph_indexes
            .lock()
            .unwrap()
            .get("body", after)
            .unwrap();
        let cached = doc.paragraph_index("body").unwrap();
        assert!(Arc::ptr_eq(&shifted, &cached));
        assert_paragraph_indexes_eq(&cached, &fresh_paragraph_index(&doc, "body"));
        assert_paragraph_indexes_eq(&original, &expected_original);
        let shifted_segments = doc
            .segment_indexes
            .lock()
            .unwrap()
            .get("body", after)
            .unwrap();
        assert!(Arc::ptr_eq(
            &shifted_segments,
            &doc.segment_index("body").unwrap()
        ));
        assert_eq!(units(&original_segments), expected_segments);
        let txn = doc.yrs_doc().transact();
        let story = crate::story_ref(&txn, "body").unwrap();
        assert_eq!(
            units(&shifted_segments),
            units(&build_indexes(&story, &txn).0)
        );
        assert_segment_invariants(&shifted_segments);
        drop(txn);
        assert_index_matches_segments(&doc, "body");
    }

    #[test]
    fn story_indexes_advance_after_text_delete() {
        let doc = seeded_doc();
        let original = doc.paragraph_index("body").unwrap();
        let original_segments = doc.segment_index("body").unwrap();
        let expected_segments = units(&original_segments);
        let expected_original = fresh_paragraph_index(&doc, "body");
        let ctx = EditCtx::local("", "");
        let before = doc.committed_epoch();
        doc.delete_range(&ctx, StoryRange::new("body", 1, 2))
            .unwrap();
        let after = doc.committed_epoch();
        assert_eq!(after, before + 1);
        doc.advance_indexes_after_text_delete("body", before, after, 1, 2);
        let shifted = doc
            .paragraph_indexes
            .lock()
            .unwrap()
            .get("body", after)
            .unwrap();
        let cached = doc.paragraph_index("body").unwrap();
        assert!(Arc::ptr_eq(&shifted, &cached));
        assert_paragraph_indexes_eq(&cached, &fresh_paragraph_index(&doc, "body"));
        assert_paragraph_indexes_eq(&original, &expected_original);
        let shifted_segments = doc
            .segment_indexes
            .lock()
            .unwrap()
            .get("body", after)
            .unwrap();
        assert!(Arc::ptr_eq(
            &shifted_segments,
            &doc.segment_index("body").unwrap()
        ));
        assert_eq!(units(&original_segments), expected_segments);
        let txn = doc.yrs_doc().transact();
        let story = crate::story_ref(&txn, "body").unwrap();
        assert_eq!(
            units(&shifted_segments),
            units(&build_indexes(&story, &txn).0)
        );
        assert_segment_invariants(&shifted_segments);
        drop(txn);
        assert_index_matches_segments(&doc, "body");
    }

    #[test]
    fn story_indexes_advance_independently_when_one_cache_is_missing() {
        for insert in [false, true] {
            for missing_segments in [false, true] {
                let doc = seeded_doc();
                doc.paragraph_index("body").unwrap();
                let before = doc.committed_epoch();
                if missing_segments {
                    doc.segment_indexes.lock().unwrap().take("body", before);
                } else {
                    doc.paragraph_indexes.lock().unwrap().take("body", before);
                }
                let ctx = EditCtx::local("", "");
                if insert {
                    doc.insert_text(&ctx, Position::new("body", 1), "xy", FormatPolicy::Inherit)
                        .unwrap();
                    doc.advance_indexes_after_text_insert(
                        "body",
                        before,
                        doc.committed_epoch(),
                        1,
                        "xy",
                    );
                } else {
                    doc.delete_range(&ctx, StoryRange::new("body", 1, 2))
                        .unwrap();
                    doc.advance_indexes_after_text_delete(
                        "body",
                        before,
                        doc.committed_epoch(),
                        1,
                        2,
                    );
                }
                let after = doc.committed_epoch();
                let segments = doc.segment_indexes.lock().unwrap().get("body", after);
                let paragraphs = doc.paragraph_indexes.lock().unwrap().get("body", after);
                assert_eq!(segments.is_some(), !missing_segments);
                assert_eq!(paragraphs.is_some(), missing_segments);
                let txn = doc.yrs_doc().transact();
                let story = crate::story_ref(&txn, "body").unwrap();
                let (cold_segments, cold_paragraphs) = build_indexes(&story, &txn);
                if let Some(segments) = segments {
                    assert_eq!(units(&segments), units(&cold_segments));
                    assert_segment_invariants(&segments);
                }
                if let Some(paragraphs) = paragraphs {
                    assert_paragraph_indexes_eq(&paragraphs, &cold_paragraphs);
                }
            }
        }
    }

    #[test]
    fn segment_index_advances_when_paragraph_shift_is_rejected() {
        let doc = seeded_doc();
        doc.paragraph_index("body").unwrap();
        let before = doc.committed_epoch();
        doc.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 5),
            "xy",
            FormatPolicy::Inherit,
        )
        .unwrap();
        let after = doc.committed_epoch();
        doc.advance_indexes_after_text_insert("body", before, after, 5, "xy");
        assert!(
            doc.paragraph_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .is_none()
        );
        let segments = doc
            .segment_indexes
            .lock()
            .unwrap()
            .get("body", after)
            .unwrap();
        let txn = doc.yrs_doc().transact();
        let story = crate::story_ref(&txn, "body").unwrap();
        assert_eq!(units(&segments), units(&build_indexes(&story, &txn).0));
        assert_segment_invariants(&segments);
    }

    #[test]
    fn story_indexes_rebuild_when_text_insert_spans_two_epochs() {
        let doc = seeded_doc();
        doc.paragraph_index("body").unwrap();
        let ctx = EditCtx::local("", "");
        let before = doc.committed_epoch();
        for index in [1, 2] {
            doc.insert_text(
                &ctx,
                Position::new("body", index),
                "xy",
                FormatPolicy::Inherit,
            )
            .unwrap();
        }
        let after = doc.committed_epoch();
        assert_eq!(after, before + 2);
        doc.advance_indexes_after_text_insert("body", before, after, 1, "xy");
        assert!(
            doc.paragraph_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .is_none()
        );
        assert!(
            doc.segment_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .is_none()
        );
        assert_paragraph_indexes_eq(
            &doc.paragraph_index("body").unwrap(),
            &fresh_paragraph_index(&doc, "body"),
        );
        assert_index_matches_segments(&doc, "body");
    }

    /// Two commits cannot advance an index built before either deletion.
    #[test]
    fn story_indexes_rebuild_when_text_delete_spans_two_epochs() {
        let doc = seeded_doc();
        doc.paragraph_index("body").unwrap();
        let ctx = EditCtx::local("", "");
        let before = doc.committed_epoch();
        for _ in 0..2 {
            doc.delete_range(&ctx, StoryRange::new("body", 0, 1))
                .unwrap();
        }
        let after = doc.committed_epoch();
        assert_eq!(after, before + 2);
        doc.advance_indexes_after_text_delete("body", before, after, 0, 1);
        assert!(
            doc.paragraph_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .is_none()
        );
        assert!(
            doc.segment_indexes
                .lock()
                .unwrap()
                .get("body", after)
                .is_none()
        );
        assert_paragraph_indexes_eq(
            &doc.paragraph_index("body").unwrap(),
            &fresh_paragraph_index(&doc, "body"),
        );
        assert_index_matches_segments(&doc, "body");
    }

    /// Plain text ranges exclude embeds, pilcrows and invalid bounds.
    #[test]
    fn segment_index_plain_text_ranges_match_segment_kinds() {
        let doc = seeded_doc();
        let segments = doc.segment_index("body").unwrap();
        let len = doc.story_len("body").unwrap();
        for start in 0..=len {
            for end in start..=len + 1 {
                let expected = start < end
                    && (start..end)
                        .all(|pos| reference_seg_kind(&doc, "body", pos) == Some("text"));
                assert_eq!(
                    segments.is_text_range(start, end),
                    expected,
                    "range {start}..{end}"
                );
            }
        }
        assert!(!segments.is_text_range(2, 1));
    }

    #[test]
    fn segment_index_matches_segment_walks() {
        let doc = seeded_doc();
        assert_index_matches_segments(&doc, "body");
    }

    #[test]
    fn segment_index_counts_the_paragraphs_carrying_an_id() {
        use yrs::{Doc, MapPrelim};
        // Built on a bare story: the editing paths repair a copied id before it is committed.
        let doc = Doc::new();
        let story = doc.get_or_insert_text("story");
        {
            let mut txn = doc.transact_mut();
            for (index, para_id) in ["p-1", "p-2", "p-2"].into_iter().enumerate() {
                let entries = [
                    (KIND_KEY.to_owned(), Any::from("pilcrow")),
                    (PARA_ID.to_owned(), Any::from(para_id)),
                ];
                story.insert_embed(&mut txn, index as u32, MapPrelim::from_iter(entries));
            }
        }
        let (_, index) = build_indexes(&story, &doc.transact());
        assert_eq!(index.para_id_count("p-1"), 1);
        assert_eq!(index.para_id_count("p-2"), 2);
        assert_eq!(index.para_id_count("absent"), 0);
        assert_eq!(index.para_span("p-2"), Some((1, 1)));
    }

    #[test]
    fn segment_index_rebuilds_after_local_and_remote_changes() {
        let doc = seeded_doc();
        let ctx = EditCtx::local(String::new(), String::new());
        let remote = EditingDoc::new(9);
        remote
            .apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        doc.insert_text(&ctx, Position::new("body", 0), "Z", FormatPolicy::Plain)
            .unwrap();
        assert_index_matches_segments(&doc, "body");
        remote
            .apply_update_v1(
                &doc.encode_diff_v1(&remote.encode_state_vector_v1())
                    .unwrap(),
            )
            .unwrap();
        assert_index_matches_segments(&remote, "body");
        doc.delete_range(&ctx, crate::StoryRange::new("body", 0, 4))
            .unwrap();
        assert_index_matches_segments(&doc, "body");
    }
}
