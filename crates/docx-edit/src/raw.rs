//! Raw story mutations using UTF-16 story indices.

use std::sync::Arc;

use yrs::types::text::YChange;
use yrs::types::{Attrs, Delta};
use yrs::{
    Any, Assoc, In, IndexedSequence, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef,
    TransactionMut,
};

use crate::op::{OpError, OpResult};
use crate::{COMMENTS, EditCtx, EditingDoc, KIND_KEY, anchor_value, out_len, story_ref};

/// One low-level story mutation. Indices are UTF-16 story units (every embed = 1).
#[derive(Clone, Debug, PartialEq)]
pub enum RawOp {
    /// Inserts text with explicit attributes.
    Insert {
        index: u32,
        text: String,
        attrs: Attrs,
    },
    /// Remove `len` units starting at `index` (text, embeds, or pilcrows alike).
    Delete { index: u32, len: u32 },
    /// Re-format `len` units at `index`; `Any::Null` values clear an attribute.
    Format { index: u32, len: u32, attrs: Attrs },
    /// Insert a map-backed embed at `index` with discriminator `kind` (`pilcrow`,
    /// `break`, `opaque`, …) and `payload` map entries. `attrs` are the embed's
    /// text-level formatting (usually just tracked-change stamps, if any).
    InsertEmbed {
        index: u32,
        kind: String,
        payload: Vec<(String, Any)>,
        attrs: Attrs,
    },
    /// Sets one key on the map-backed embed at the index.
    SetEmbedAttr { index: u32, key: String, value: Any },
    /// Upserts a side-map comment with sticky UTF-16 ranges.
    SetComment {
        id: String,
        ranges: Vec<(u32, u32)>,
        author: String,
        date: String,
        body: Any,
    },
    /// Remove the side-map comment keyed by `id`. Errors when it does not exist.
    RemoveComment { id: String },
}

/// Finds the map-backed embed sitting exactly at story `index`.
fn embed_at<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32) -> OpResult<MapRef> {
    let mut offset = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        if offset == index {
            if let Out::YMap(map) = diff.insert {
                return Ok(map);
            }
            break;
        }
        offset += out_len(&diff.insert);
        if offset > index {
            break;
        }
    }
    Err(OpError::OutOfBounds {
        index,
        len: story.len(txn),
    })
}

impl EditingDoc {
    /// Applies raw story operations in one transaction.
    pub fn apply_raw_ops(&self, story_id: &str, ops: Vec<RawOp>, ctx: &EditCtx) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        apply_raw_ops_to_story(&mut txn, story_id, ops)
    }

    pub(crate) fn apply_raw_story_batches(
        &self,
        batches: Vec<(String, Vec<RawOp>)>,
        ctx: &EditCtx,
    ) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        for (story_id, ops) in batches {
            apply_raw_ops_to_story(&mut txn, &story_id, ops)?;
        }
        Ok(())
    }
}

/// Forward-moving ops batched into one yrs delta: a run walks each block once,
/// where per-op absolute-index resolution walked O(index) and made seeding O(n²).
struct DeltaRun {
    deltas: Vec<Delta<In>>,
    cursor: u32,
    story_len: u32,
}

impl DeltaRun {
    fn new(story: &TextRef, txn: &TransactionMut<'_>) -> Self {
        Self {
            deltas: Vec::new(),
            cursor: 0,
            story_len: story.len(txn),
        }
    }

    fn flush(&mut self, story: &TextRef, txn: &mut TransactionMut<'_>) {
        if !self.deltas.is_empty() {
            story.apply_delta(txn, std::mem::take(&mut self.deltas));
        }
        self.cursor = 0;
    }

    fn seek(&mut self, story: &TextRef, txn: &mut TransactionMut<'_>, index: u32) {
        if index < self.cursor {
            self.flush(story, txn);
        }
        if index > self.cursor {
            self.deltas.push(Delta::Retain(index - self.cursor, None));
        }
        self.cursor = index;
    }

    fn guard_index(&self, index: u32) -> OpResult<()> {
        if index <= self.story_len {
            Ok(())
        } else {
            Err(OpError::OutOfBounds {
                index,
                len: self.story_len,
            })
        }
    }

    fn guard_range(&self, index: u32, len: u32) -> OpResult<()> {
        if index
            .checked_add(len)
            .is_some_and(|end| end <= self.story_len)
        {
            Ok(())
        } else {
            Err(OpError::OutOfBounds {
                index: index.saturating_add(len),
                len: self.story_len,
            })
        }
    }

    fn insert(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        text: String,
        attrs: Attrs,
    ) -> OpResult<()> {
        self.guard_index(index)?;
        if text.is_empty() {
            return Ok(());
        }
        let len = text.encode_utf16().count() as u32;
        self.seek(story, txn, index);
        self.deltas.push(Delta::Inserted(
            In::Any(Any::String(Arc::from(text))),
            (!attrs.is_empty()).then(|| Box::new(attrs)),
        ));
        self.cursor += len;
        self.story_len += len;
        Ok(())
    }

    fn insert_embed(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        kind: String,
        payload: Vec<(String, Any)>,
        attrs: Attrs,
    ) -> OpResult<()> {
        self.guard_index(index)?;
        self.seek(story, txn, index);
        let entries = std::iter::once((KIND_KEY.to_owned(), Any::from(kind))).chain(payload);
        self.deltas.push(Delta::Inserted(
            In::Map(MapPrelim::from_iter(entries)),
            (!attrs.is_empty()).then(|| Box::new(attrs)),
        ));
        self.cursor += 1;
        self.story_len += 1;
        Ok(())
    }

    fn delete(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        len: u32,
    ) -> OpResult<()> {
        self.guard_range(index, len)?;
        if len > 0 {
            self.seek(story, txn, index);
            self.deltas.push(Delta::Deleted(len));
            self.story_len -= len;
        }
        Ok(())
    }

    fn format(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        len: u32,
        attrs: Attrs,
    ) -> OpResult<()> {
        self.guard_range(index, len)?;
        if len > 0 {
            self.seek(story, txn, index);
            self.deltas.push(Delta::Retain(len, Some(Box::new(attrs))));
            self.cursor += len;
        }
        Ok(())
    }
}

fn apply_raw_ops_to_story(
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    ops: Vec<RawOp>,
) -> OpResult<()> {
    let story = story_ref(txn, story_id).map_err(OpError::from)?;
    let mut run = DeltaRun::new(&story, txn);
    let mut result = Ok(());
    for op in ops {
        result = match op {
            RawOp::Insert { index, text, attrs } => run.insert(&story, txn, index, text, attrs),
            RawOp::Delete { index, len } => run.delete(&story, txn, index, len),
            RawOp::Format { index, len, attrs } => run.format(&story, txn, index, len, attrs),
            RawOp::InsertEmbed {
                index,
                kind,
                payload,
                attrs,
            } => run.insert_embed(&story, txn, index, kind, payload, attrs),
            RawOp::SetEmbedAttr { index, key, value } => {
                run.flush(&story, txn);
                embed_at(&story, txn, index).map(|embed| {
                    embed.insert(txn, key, value);
                })
            }
            RawOp::SetComment {
                id,
                ranges,
                author,
                date,
                body,
            } => {
                run.flush(&story, txn);
                set_comment(txn, &story, story_id, id, ranges, author, date, body)
            }
            RawOp::RemoveComment { id } => {
                run.flush(&story, txn);
                let comments = txn
                    .get_map(COMMENTS)
                    .expect("comments root is declared by EditingDoc::new");
                if comments.remove(txn, &id).is_none() {
                    Err(OpError::UnknownComment(id))
                } else {
                    Ok(())
                }
            }
        };
        if result.is_err() {
            break;
        }
    }
    run.flush(&story, txn);
    result
}

#[allow(clippy::too_many_arguments)]
fn set_comment(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    id: String,
    ranges: Vec<(u32, u32)>,
    author: String,
    date: String,
    body: Any,
) -> OpResult<()> {
    if ranges.is_empty() {
        return Err(OpError::InvalidComment(
            "at least one anchored range is required".into(),
        ));
    }
    let mut anchors = Vec::with_capacity(ranges.len());
    for (start, end) in ranges {
        let len = end
            .checked_sub(start)
            .filter(|len| *len > 0)
            .ok_or(OpError::InvalidRange { start, end })?;
        guard_range(story, txn, start, len)?;
        let start_anchor = story
            .sticky_index(txn, start, Assoc::After)
            .ok_or_else(|| OpError::InvalidComment("start anchor could not be made".into()))?;
        let end_anchor = story
            .sticky_index(txn, end, Assoc::Before)
            .ok_or_else(|| OpError::InvalidComment("end anchor could not be made".into()))?;
        anchors.push(anchor_value(story_id, &start_anchor, &end_anchor));
    }
    let comments = txn
        .get_map(COMMENTS)
        .expect("comments root is declared by EditingDoc::new");
    let comment = comments.insert(txn, id.as_str(), MapPrelim::default());
    comment.insert(txn, "author", author.as_str());
    comment.insert(txn, "date", date.as_str());
    comment.insert(txn, "parentId", Any::Null);
    comment.insert(txn, "done", false);
    comment.insert(txn, "body", body);
    comment.insert(txn, "anchors", Any::Array(Arc::from(anchors)));
    Ok(())
}

fn guard_range<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32, len: u32) -> OpResult<()> {
    let story_len = story.len(txn);
    if index.checked_add(len).is_some_and(|end| end <= story_len) {
        Ok(())
    } else {
        Err(OpError::OutOfBounds {
            index: index.saturating_add(len),
            len: story_len,
        })
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use super::*;
    use crate::canonical::project_story;
    use crate::{EditCtx, EditingDoc, PILCROW_KIND, seed_from_docx};

    fn attrs(pairs: &[(&str, Any)]) -> Attrs {
        pairs
            .iter()
            .map(|(key, value)| (Arc::from(*key), value.clone()))
            .collect()
    }

    /// The replaced per-op absolute-index applier, kept verbatim as the equivalence oracle.
    fn apply_raw_ops_legacy(
        doc: &EditingDoc,
        story_id: &str,
        ops: Vec<RawOp>,
        ctx: &EditCtx,
    ) -> OpResult<()> {
        let mut txn = doc.transact_for(ctx);
        let story = story_ref(&txn, story_id).map_err(OpError::from)?;
        let guard_index = |txn: &TransactionMut<'_>, index: u32| -> OpResult<()> {
            if index <= story.len(txn) {
                Ok(())
            } else {
                Err(OpError::OutOfBounds {
                    index,
                    len: story.len(txn),
                })
            }
        };
        for op in ops {
            match op {
                RawOp::Insert { index, text, attrs } => {
                    guard_index(&txn, index)?;
                    story.insert_with_attributes(&mut txn, index, &text, attrs);
                }
                RawOp::Delete { index, len } => {
                    guard_range(&story, &txn, index, len)?;
                    story.remove_range(&mut txn, index, len);
                }
                RawOp::Format { index, len, attrs } => {
                    guard_range(&story, &txn, index, len)?;
                    if len > 0 {
                        story.format(&mut txn, index, len, attrs);
                    }
                }
                RawOp::InsertEmbed {
                    index,
                    kind,
                    payload,
                    attrs,
                } => {
                    guard_index(&txn, index)?;
                    let embed = story.insert_embed_with_attributes(
                        &mut txn,
                        index,
                        MapPrelim::default(),
                        attrs,
                    );
                    embed.insert(&mut txn, KIND_KEY, kind.as_str());
                    for (key, value) in payload {
                        embed.insert(&mut txn, key, value);
                    }
                }
                RawOp::SetEmbedAttr { index, key, value } => {
                    let embed = embed_at(&story, &txn, index)?;
                    embed.insert(&mut txn, key, value);
                }
                RawOp::SetComment {
                    id,
                    ranges,
                    author,
                    date,
                    body,
                } => set_comment(&mut txn, &story, story_id, id, ranges, author, date, body)?,
                RawOp::RemoveComment { id } => {
                    let comments = txn
                        .get_map(COMMENTS)
                        .expect("comments root is declared by EditingDoc::new");
                    if comments.remove(&mut txn, &id).is_none() {
                        return Err(OpError::UnknownComment(id));
                    }
                }
            }
        }
        Ok(())
    }

    fn assert_stories_equivalent(left: &EditingDoc, right: &EditingDoc, story_id: &str) {
        assert_eq!(
            project_story(left, story_id).unwrap(),
            project_story(right, story_id).unwrap()
        );
        assert_eq!(
            left.story_segments(story_id).unwrap(),
            right.story_segments(story_id).unwrap()
        );
    }

    fn pilcrow_payload(para_id: &str) -> Vec<(String, Any)> {
        vec![
            ("paraId".into(), Any::from(para_id)),
            ("pStyle".into(), Any::from("Normal")),
            ("alignment".into(), Any::from("left")),
        ]
    }

    /// What `seed_parsed_docx` emits: placeholder delete, appending inserts, trailing comment.
    fn seed_shaped_ops() -> Vec<RawOp> {
        let font = attrs(&[(
            "fontFamily",
            Any::from_json(r#"{"ascii":"Calibri","hAnsi":"Calibri"}"#).unwrap(),
        )]);
        let mut bold = font.clone();
        bold.insert("bold".into(), Any::Bool(true));
        struct Builder {
            ops: Vec<RawOp>,
            index: u32,
        }
        impl Builder {
            fn text(&mut self, value: &str, attrs: &Attrs) {
                self.ops.push(RawOp::Insert {
                    index: self.index,
                    text: value.into(),
                    attrs: attrs.clone(),
                });
                self.index += value.encode_utf16().count() as u32;
            }
            fn embed(&mut self, kind: &str, payload: Vec<(String, Any)>) {
                self.ops.push(RawOp::InsertEmbed {
                    index: self.index,
                    kind: kind.into(),
                    payload,
                    attrs: Attrs::new(),
                });
                self.index += 1;
            }
        }
        let mut builder = Builder {
            ops: vec![RawOp::Delete { index: 0, len: 1 }],
            index: 0,
        };
        builder.text("Heading with stars ★★", &bold);
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p0"));
        builder.text("Body ", &font);
        builder.text("bold span", &bold);
        builder.text(" and surrogate pairs 𝔘𝔫𝔦", &font);
        builder.embed("break", Vec::new());
        builder.text("after the break", &font);
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p1"));
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p2"));
        builder.text("tail paragraph", &font);
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p3"));
        let mut ops = builder.ops;
        ops.push(RawOp::SetComment {
            id: "9".into(),
            ranges: vec![(22, 30), (40, 45)],
            author: String::new(),
            date: String::new(),
            body: Any::Null,
        });
        ops
    }

    #[test]
    fn delta_application_matches_legacy_for_a_seed_shaped_batch() {
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy.create_empty_stories(&["body".into()]).unwrap();
        apply_raw_ops_legacy(&legacy, "body", seed_shaped_ops(), &ctx).unwrap();

        let delta = EditingDoc::new(7);
        delta.create_empty_stories(&["body".into()]).unwrap();
        delta
            .apply_raw_ops("body", seed_shaped_ops(), &ctx)
            .unwrap();

        assert_stories_equivalent(&legacy, &delta, "body");
        assert_eq!(
            legacy.paragraphs("body").unwrap().len(),
            delta.paragraphs("body").unwrap().len()
        );
        let legacy_anchor = legacy.resolve_comment("9").unwrap();
        let delta_anchor = delta.resolve_comment("9").unwrap();
        assert_eq!(
            legacy_anchor
                .iter()
                .map(|anchor| (anchor.start, anchor.end))
                .collect::<Vec<_>>(),
            delta_anchor
                .iter()
                .map(|anchor| (anchor.start, anchor.end))
                .collect::<Vec<_>>()
        );
    }

    #[test]
    fn delta_application_matches_legacy_for_backward_jumps() {
        let ops = vec![
            RawOp::Insert {
                index: 4,
                text: " tail".into(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 2,
                text: "XY".into(),
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
            RawOp::Format {
                index: 1,
                len: 5,
                attrs: attrs(&[("italic", Any::Bool(true))]),
            },
            RawOp::Delete { index: 3, len: 2 },
            RawOp::InsertEmbed {
                index: 0,
                kind: "break".into(),
                payload: Vec::new(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 10,
                text: "end".into(),
                attrs: Attrs::new(),
            },
            RawOp::SetEmbedAttr {
                index: 0,
                key: "cleared".into(),
                value: Any::Bool(true),
            },
        ];
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy
            .create_story("body", "seed", "Normal", "left")
            .unwrap();
        apply_raw_ops_legacy(&legacy, "body", ops.clone(), &ctx).unwrap();

        let delta = EditingDoc::new(7);
        delta
            .create_story("body", "seed", "Normal", "left")
            .unwrap();
        delta.apply_raw_ops("body", ops, &ctx).unwrap();

        assert_stories_equivalent(&legacy, &delta, "body");
    }

    #[test]
    fn a_failing_op_leaves_the_same_applied_prefix_as_legacy() {
        let ops = vec![
            RawOp::Insert {
                index: 2,
                text: "one".into(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 0,
                text: "two".into(),
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
            RawOp::Delete { index: 90, len: 5 },
            RawOp::Insert {
                index: 0,
                text: "never applied".into(),
                attrs: Attrs::new(),
            },
        ];
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy.create_story("body", "AB", "Normal", "left").unwrap();
        let legacy_err = apply_raw_ops_legacy(&legacy, "body", ops.clone(), &ctx).unwrap_err();

        let delta = EditingDoc::new(7);
        delta.create_story("body", "AB", "Normal", "left").unwrap();
        let delta_err = delta.apply_raw_ops("body", ops, &ctx).unwrap_err();

        assert_eq!(format!("{legacy_err:?}"), format!("{delta_err:?}"));
        assert_stories_equivalent(&legacy, &delta, "body");
    }

    #[test]
    fn legacy_seeded_state_loads_and_merges_under_the_delta_applier() {
        let ctx = EditCtx::local(String::new(), String::new());
        let original = EditingDoc::new(1);
        original.create_empty_stories(&["body".into()]).unwrap();
        apply_raw_ops_legacy(&original, "body", seed_shaped_ops(), &ctx).unwrap();

        let restored = EditingDoc::new(2);
        restored
            .apply_update_v1(&original.encode_state_as_update_v1())
            .unwrap();
        assert_stories_equivalent(&original, &restored, "body");

        restored
            .apply_raw_ops(
                "body",
                vec![RawOp::Insert {
                    index: 5,
                    text: "restored edit ".into(),
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                }],
                &ctx,
            )
            .unwrap();
        original
            .apply_update_v1(&restored.encode_state_as_update_v1())
            .unwrap();
        restored
            .apply_update_v1(&original.encode_state_as_update_v1())
            .unwrap();
        assert_stories_equivalent(&original, &restored, "body");
        assert!(
            original
                .story_segments("body")
                .unwrap()
                .iter()
                .any(|segment| matches!(
                    &segment.content,
                    crate::SegmentContent::Text(text) if text.contains("restored edit")
                ))
        );
    }

    #[test]
    fn seeded_docx_state_round_trips_through_encode_and_load() {
        let bytes = include_bytes!("../tests/fixtures/footnote-anchor.docx");
        let seeded = EditingDoc::new(1);
        seed_from_docx(&seeded, bytes).unwrap();

        let loaded = EditingDoc::new(2);
        loaded
            .apply_update_v1(&seeded.encode_state_as_update_v1())
            .unwrap();
        let txn = yrs::Transact::transact(seeded.yrs_doc());
        let stories = txn.get_map(crate::STORIES).unwrap();
        let story_ids: Vec<String> = stories.keys(&txn).map(|key| key.to_string()).collect();
        drop(txn);
        assert!(!story_ids.is_empty());
        for story_id in story_ids {
            assert_stories_equivalent(&seeded, &loaded, &story_id);
        }
    }

    #[test]
    fn insert_format_and_delete_apply_in_one_batch() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        // story: A(0) B(1) pilcrow(2)
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Insert {
                    index: 1,
                    text: "X".into(),
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                // story: A(0) X(1) B(2) pilcrow(3)
                RawOp::Format {
                    index: 0,
                    len: 1,
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                },
                RawOp::Delete { index: 2, len: 1 }, // remove B
                                                    // story: A(0) X(1) pilcrow(2)
            ],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].text, "AX");
    }

    #[test]
    fn insert_embed_pilcrow_splits_a_paragraph() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 1,
                kind: PILCROW_KIND.to_owned(),
                payload: vec![
                    ("paraId".into(), Any::from("7:seed")),
                    ("pStyle".into(), Any::from("Normal")),
                    ("alignment".into(), Any::from("left")),
                ],
                attrs: Attrs::new(),
            }],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text, "A");
        assert_eq!(paras[1].text, "B");
    }

    #[test]
    fn set_embed_attr_updates_a_pilcrow_property() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        // The pilcrow sits at index 2 (after "AB").
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetEmbedAttr {
                index: 2,
                key: "alignment".into(),
                value: Any::from("center"),
            }],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(
            paras[0].properties.get("alignment"),
            Some(&Any::from("center"))
        );
    }

    #[test]
    fn set_comment_upserts_and_remove_comment_deletes() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "ABCDE", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "9".into(),
                ranges: vec![(1, 3)],
                author: "Ada".into(),
                date: "2026-07-13T12:00:00Z".into(),
                body: Any::from("hi"),
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!((resolved[0].start, resolved[0].end), (1, 3));

        // Sticky anchors ride an insertion before the range.
        doc.apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: 0,
                text: "xx".into(),
                attrs: Attrs::new(),
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!((resolved[0].start, resolved[0].end), (3, 5));

        // Upsert on the same key re-anchors.
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "9".into(),
                ranges: vec![(0, 2), (4, 6)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!(resolved.len(), 2);
        assert_eq!((resolved[0].start, resolved[0].end), (0, 2));
        assert_eq!((resolved[1].start, resolved[1].end), (4, 6));

        doc.apply_raw_ops("body", vec![RawOp::RemoveComment { id: "9".into() }], &ctx)
            .unwrap();
        assert!(doc.resolve_comment("9").is_err());
        // Removing an unknown comment is a typed error, not silence.
        let missing =
            doc.apply_raw_ops("body", vec![RawOp::RemoveComment { id: "9".into() }], &ctx);
        assert!(matches!(missing, Err(OpError::UnknownComment(_))));
    }

    #[test]
    fn set_comment_rejects_empty_and_out_of_bounds_ranges() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let empty = doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "1".into(),
                ranges: vec![(1, 1)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        );
        assert!(matches!(empty, Err(OpError::InvalidRange { .. })));
        let outside = doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "1".into(),
                ranges: vec![(0, 99)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        );
        assert!(matches!(outside, Err(OpError::OutOfBounds { .. })));
    }

    #[test]
    fn out_of_bounds_index_is_a_typed_error() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let result = doc.apply_raw_ops("body", vec![RawOp::Delete { index: 99, len: 1 }], &ctx);
        assert!(matches!(result, Err(OpError::OutOfBounds { .. })));
    }
}
