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

pub(crate) enum SegKind {
    Text(Arc<str>),
    Pilcrow,
    Embed,
}

pub(crate) struct Seg {
    pub start: u32,
    pub kind: SegKind,
}

/// Materialized segment geometry for one story at one committed epoch.
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
}

impl SegmentIndex {
    /// The segment covering `pos`, if any.
    pub(crate) fn segment_at(&self, pos: u32) -> Option<&Seg> {
        if pos >= self.len {
            return None;
        }
        self.segs
            .get(self.segs.partition_point(|seg| seg.start <= pos) - 1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, EditingDoc, FormatPolicy, Position, RawOp, SegmentContent};
    use yrs::{Map, Transact};

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
        doc
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

    fn assert_paragraph_indexes_eq(actual: &ParagraphIndex, expected: &ParagraphIndex) {
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

    #[test]
    fn paragraph_index_advances_after_text_insert_and_rebuilds_segments() {
        let doc = seeded_doc();
        let original = doc.paragraph_index("body").unwrap();
        let expected_original = fresh_paragraph_index(&doc, "body");
        let ctx = EditCtx::local("", "");
        let before = doc.committed_epoch();
        doc.insert_text(&ctx, Position::new("body", 1), "xy", FormatPolicy::Inherit)
            .unwrap();
        let after = doc.committed_epoch();
        assert_eq!(after, before + 1);
        doc.advance_paragraph_index_after_text_insert("body", before, after, 1, 2);
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
        assert_index_matches_segments(&doc, "body");
    }

    #[test]
    fn paragraph_index_rebuilds_when_text_insert_spans_two_epochs() {
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
        doc.advance_paragraph_index_after_text_insert("body", before, after, 1, 2);
        assert!(
            doc.paragraph_indexes
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
