// Segment/paragraph accessors used only by the wasm helpers read as dead code in
// native builds; the invalidation and paragraph paths below are live on both.
#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::HashMap;
use std::sync::Arc;

use yrs::types::text::YChange;
use yrs::{Any, Out, ReadTxn, Text, TextRef};

use crate::{KIND_KEY, PARA_ID, is_pilcrow, map_string, out_len};

/// Whether the layout gives an embed its own block.
fn is_block_embed(kind: &str) -> bool {
    matches!(kind, "table" | "blockSdt" | "pageBreak" | "columnBreak")
}

/// One paragraph's resolved geometry inside a story.
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

/// Materialized segment geometry for one story at one committed epoch, built by a
/// single `diff` walk that keeps only what the location/index helpers read. Lookups
/// then answer by binary search or hash lookup instead of re-walking
/// `story_segments` on every call.
pub(crate) struct SegmentIndex {
    len: u32,
    segs: Vec<Seg>,
    paras: Vec<ParaEntry>,
    by_para: HashMap<Box<str>, u32>,
}

impl SegmentIndex {
    pub(crate) fn build<T: ReadTxn>(story: &TextRef, txn: &T) -> Self {
        let mut len = 0_u32;
        let mut para_start = 0_u32;
        let mut node_start = 0_u32;
        let mut segs = Vec::new();
        let mut paras: Vec<ParaEntry> = Vec::new();
        let mut by_para = HashMap::new();
        for diff in story.diff(txn, YChange::identity) {
            let units = out_len(&diff.insert);
            let kind = match diff.insert {
                Out::Any(Any::String(text)) => SegKind::Text(text),
                Out::YMap(map) if is_pilcrow(&map, txn) => {
                    let para_id: Box<str> = map_string(&map, txn, PARA_ID)
                        .unwrap_or_default()
                        .into_boxed_str();
                    let slot = paras.len() as u32;
                    by_para.entry(para_id.clone()).or_insert(slot);
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
        Self {
            len,
            segs,
            paras,
            by_para,
        }
    }

    /// `(start, pilcrow)` span of `para_id`, matching the first paragraph with that id.
    pub(crate) fn para_span(&self, para_id: &str) -> Option<(u32, u32)> {
        let para = self.paras.get(*self.by_para.get(para_id)? as usize)?;
        Some((para.start, para.pilcrow))
    }

    /// First paragraph whose pilcrow sits at or after `index` — the paragraph `index`
    /// resolves into.
    pub(crate) fn para_at(&self, index: u32) -> Option<&ParaEntry> {
        self.paras
            .get(self.paras.partition_point(|para| para.pilcrow < index))
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, EditingDoc, FormatPolicy, Position, RawOp, SegmentContent};

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
            let actual = index.para_at(index_pos).map(|para| {
                (
                    para.para_id.to_string(),
                    index_pos.saturating_sub(para.start),
                    index_pos.saturating_sub(para.node_start),
                )
            });
            assert_eq!(expected, actual, "index {index_pos}");
        }
        assert_eq!(index.paras.len() as u32, para_count);
        for paragraph in doc.paragraphs(story).unwrap() {
            let span = index.para_span(&paragraph.para_id).unwrap();
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

    #[test]
    fn segment_index_matches_segment_walks() {
        let doc = seeded_doc();
        assert_index_matches_segments(&doc, "body");
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
