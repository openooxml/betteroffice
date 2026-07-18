//! Raw story-mutation primitives for the coexistence bridge.
//!
//! The Route-B coexistence harness mirrors ProseMirror state into yrs by lowering
//! each PM transaction to a list of low-level story operations applied in ONE yrs
//! transaction (op-contract R1). Unlike the typed op surface (`insert_text`,
//! `split_paragraph`, …) which encodes user-intent semantics (format inheritance,
//! tracked-change 3-way splits, pPr survival), these primitives are a faithful
//! mirror: they write exactly the units and attributes the bridge computed from
//! PM, so the yrs shadow matches PM byte-for-byte under the canonical invariant.
//! Tracked-change attributes (`ins`/`del`) arrive already lowered inside `attrs`
//! when the PM run carried an insertion/deletion mark; the primitives never stamp
//! their own.
//!
//! Addressing is story-global UTF-16 indices, transient to the batch: each op's
//! index is interpreted against the story state AFTER all prior ops in the batch.

use std::sync::Arc;

use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Assoc, IndexedSequence, Map, MapPrelim, MapRef, Out, ReadTxn, Text};

use crate::op::{OpError, OpResult};
use crate::{COMMENTS, EditCtx, EditingDoc, KIND_KEY, anchor_value, out_len, story_ref};

/// One low-level story mutation. Indices are UTF-16 story units (every embed = 1).
#[derive(Clone, Debug, PartialEq)]
pub enum RawOp {
    /// Insert `text` at `index`, each character carrying `attrs` (a faithful mirror
    /// of the PM run's lowered formatting; no ins/del is added automatically).
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
    /// Set one key on the map-backed embed currently at `index` (e.g. a pilcrow's
    /// paragraph property from a PM `AttrStep`). Errors if no embed sits there.
    SetEmbedAttr { index: u32, key: String, value: Any },
    /// Upsert the side-map comment keyed by `id`, (re-)anchoring it to `ranges` —
    /// non-empty `[start, end)` UTF-16 story-unit spans in the batch's story, read
    /// (like every raw index) against the story state after all prior ops. Starts
    /// anchor with [`Assoc::After`], ends with [`Assoc::Before`], mirroring
    /// [`EditingDoc::add_comment`]. The coexistence bridge keys by the PM comment
    /// id so identity survives the mirror (and the render bridge's `numeric_id`).
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
    /// Applies a batch of raw story ops in ONE yrs transaction under `ctx`'s origin.
    ///
    /// The coexistence bridge is the sole caller; ops are pre-validated by the
    /// translator against its position map, but bounds are re-checked here so a
    /// bridge bug fails loudly (typed error) rather than corrupting the CRDT.
    pub fn apply_raw_ops(&self, story_id: &str, ops: Vec<RawOp>, ctx: &EditCtx) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, story_id).map_err(OpError::from)?;
        for op in ops {
            match op {
                RawOp::Insert { index, text, attrs } => {
                    guard_index(&story, &txn, index)?;
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
                    guard_index(&story, &txn, index)?;
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
                } => {
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
                        guard_range(&story, &txn, start, len)?;
                        let start_anchor = story
                            .sticky_index(&mut txn, start, Assoc::After)
                            .ok_or_else(|| {
                                OpError::InvalidComment("start anchor could not be made".into())
                            })?;
                        let end_anchor = story
                            .sticky_index(&mut txn, end, Assoc::Before)
                            .ok_or_else(|| {
                                OpError::InvalidComment("end anchor could not be made".into())
                            })?;
                        anchors.push(anchor_value(story_id, &start_anchor, &end_anchor));
                    }
                    let comments = txn
                        .get_map(COMMENTS)
                        .expect("comments root is declared by EditingDoc::new");
                    let comment = comments.insert(&mut txn, id.as_str(), MapPrelim::default());
                    comment.insert(&mut txn, "author", author.as_str());
                    comment.insert(&mut txn, "date", date.as_str());
                    comment.insert(&mut txn, "parentId", Any::Null);
                    comment.insert(&mut txn, "done", false);
                    comment.insert(&mut txn, "body", body);
                    comment.insert(&mut txn, "anchors", Any::Array(Arc::from(anchors)));
                }
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
}

fn guard_index<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32) -> OpResult<()> {
    if index <= story.len(txn) {
        Ok(())
    } else {
        Err(OpError::OutOfBounds {
            index,
            len: story.len(txn),
        })
    }
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
    use crate::{EditCtx, PILCROW_KIND};

    fn attrs(pairs: &[(&str, Any)]) -> Attrs {
        pairs
            .iter()
            .map(|(key, value)| (Arc::from(*key), value.clone()))
            .collect()
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
