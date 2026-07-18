//! Embed ops: stable-id/position attribute updates and exact-position inserts — the typed write
//! surface for map-backed embeds (images, inline/block content controls,
//! block-level breaks).
//!
//! These are user-intent ops (undo-tracked under the ctx origin), unlike the
//! raw [`crate::RawOp::SetEmbedAttr`] / [`crate::RawOp::InsertEmbed`] mirror
//! primitives which faithfully replay bridge-lowered PM state. The payload
//! vocabulary is the same one the coexistence bridge authors (PM attr names,
//! `dropNulls`-style absent-as-default), so a value written here reads back
//! identically through `story_segments` and the render bridge.

use yrs::types::text::YChange;
use yrs::{Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef};

use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::{adjacent_paragraph_change_revision_id, adjacent_revision_id, snapshot};
use crate::{
    EditCtx, EditingDoc, INS, KIND_KEY, PARA_ID, PILCROW_KIND, Position, check_position,
    insertion_attrs, is_pilcrow, out_len, revision_value, story_ref,
};

/// Finds the map-backed embed sitting exactly at story `index` (any kind,
/// pilcrows included).
fn embed_map_at<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> OpResult<MapRef> {
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

fn id_value_matches(value: Option<Out>, embed_id: &str) -> bool {
    match value {
        Some(Out::Any(Any::String(value))) => value.as_ref() == embed_id,
        Some(Out::Any(Any::Number(value))) => embed_id
            .parse::<f64>()
            .is_ok_and(|candidate| candidate == value),
        Some(Out::Any(Any::BigInt(value))) => embed_id
            .parse::<i64>()
            .is_ok_and(|candidate| candidate == value),
        _ => false,
    }
}

fn embed_has_id<T: ReadTxn>(map: &MapRef, txn: &T, embed_id: &str) -> bool {
    ["embedId", "id", "rId"]
        .into_iter()
        .any(|key| id_value_matches(map.get(txn, key), embed_id))
}

/// Finds one map-backed embed by its stable authored payload identity. New yrs
/// inserts use `embedId`; `id` (SDTs) and `rId` (images) keep mirrored PM
/// embeds addressable without rewriting their payload vocabulary.
fn embed_map_by_id<T: ReadTxn>(txn: &T, embed_id: &str) -> OpResult<MapRef> {
    let stories = txn
        .get_map(crate::STORIES)
        .ok_or_else(|| OpError::UnknownEmbed(embed_id.to_owned()))?;
    let mut story_ids: Vec<String> = stories.keys(txn).map(|key| key.to_string()).collect();
    story_ids.sort();
    for story_id in story_ids {
        let Some(Out::YText(story)) = stories.get(txn, &story_id) else {
            continue;
        };
        for diff in story.diff(txn, YChange::identity) {
            if let Out::YMap(map) = diff.insert
                && !is_pilcrow(&map, txn)
                && embed_has_id(&map, txn, embed_id)
            {
                return Ok(map);
            }
        }
    }
    Err(OpError::UnknownEmbed(embed_id.to_owned()))
}

impl EditingDoc {
    /// Sets (or, with [`Any::Null`], removes) payload entries on the map-backed
    /// embed at `at` in ONE transaction — the mutation behind image geometry
    /// commits and content-control state changes. The `_kind` discriminator is
    /// reserved on every embed; a pilcrow's `paraId` is reserved too (identity
    /// is managed by the schema — use `set_paragraph_attr` for pilcrow
    /// properties anyway). Errors before any mutation when `at` does not hold
    /// a map-backed embed or a reserved key is supplied.
    pub fn set_embed_attrs(
        &self,
        ctx: &EditCtx,
        at: Position,
        entries: Vec<(String, Any)>,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        let map = embed_map_at(&story, &txn, at.index)?;
        let pilcrow = is_pilcrow(&map, &txn);
        for (key, _) in &entries {
            if key == KIND_KEY || (pilcrow && key == PARA_ID) {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        for (key, value) in entries {
            if value == Any::Null {
                map.remove(&mut txn, &key);
            } else {
                map.insert(&mut txn, key, value);
            }
        }
        Ok(Receipt::default())
    }

    /// Stable-id variant of [`EditingDoc::set_embed_attrs`], used by React
    /// image and SDT commands whose handles survive concurrent story edits.
    /// Every entry is applied in one transaction; [`Any::Null`] removes it.
    pub fn set_embed_attrs_by_id(
        &self,
        ctx: &EditCtx,
        embed_id: &str,
        entries: Vec<(String, Any)>,
    ) -> OpResult<Receipt> {
        for (key, _) in &entries {
            if key == KIND_KEY || key == PARA_ID {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        let mut txn = self.transact_for(ctx);
        let map = embed_map_by_id(&txn, embed_id)?;
        for (key, value) in entries {
            if value == Any::Null {
                map.remove(&mut txn, &key);
            } else {
                map.insert(&mut txn, key, value);
            }
        }
        Ok(Receipt::default())
    }

    /// Inserts a non-pilcrow map-backed embed at the exact story position.
    /// The caller supplies the typed kind/payload vocabulary; suggesting mode
    /// stamps the embed with one insertion revision.
    pub fn insert_embed(
        &self,
        ctx: &EditCtx,
        at: Position,
        kind: &str,
        payload: Vec<(String, Any)>,
    ) -> OpResult<Receipt> {
        if kind == PILCROW_KIND {
            return Err(OpError::ReservedKey(kind.to_owned()));
        }
        for (key, _) in &payload {
            if key == KIND_KEY || key == PARA_ID {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        let chunks = snapshot(&story, &txn);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, INS, &ctx.author)
                .or_else(|| {
                    adjacent_paragraph_change_revision_id(&chunks, at.index, &txn, &ctx.author)
                })
                .unwrap_or_else(|| self.next_id())
        });
        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let embed = story.insert_embed_with_attributes(
            &mut txn,
            at.index,
            MapPrelim::default(),
            insertion_attrs(ins, None),
        );
        embed.insert(&mut txn, KIND_KEY, kind);
        for (key, value) in payload {
            embed.insert(&mut txn, key, value);
        }
        let range = loc_range_in_txn(&at.story, &story, &txn, at.index, at.index + 1)?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: revision_id.into_iter().collect(),
            range: Some(range),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::SegmentContent;

    fn ctx() -> EditCtx {
        EditCtx::local(String::new(), String::new())
    }

    fn seed_sdt(doc: &EditingDoc) -> u32 {
        // story: A(0) [sdt](1) B(2) pilcrow(3)
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![crate::RawOp::InsertEmbed {
                index: 1,
                kind: "sdt".into(),
                payload: vec![
                    ("embedId".into(), Any::from("control-1")),
                    ("sdtType".into(), Any::from("checkbox")),
                    ("checked".into(), Any::Bool(false)),
                ],
                attrs: Default::default(),
            }],
            &ctx(),
        )
        .unwrap();
        1
    }

    fn embed_payload(doc: &EditingDoc, index: u32) -> std::collections::BTreeMap<String, Any> {
        let mut offset = 0u32;
        for segment in doc.story_segments("body").unwrap() {
            let len = match &segment.content {
                SegmentContent::Text(text) => text.encode_utf16().count() as u32,
                _ => 1,
            };
            if offset == index {
                if let SegmentContent::OtherEmbed { payload, .. } = segment.content {
                    return payload;
                }
                panic!("expected an embed at {index}");
            }
            offset += len;
        }
        panic!("no segment at {index}");
    }

    #[test]
    fn set_embed_attrs_writes_and_null_removes() {
        let doc = EditingDoc::new(7);
        let index = seed_sdt(&doc);
        doc.set_embed_attrs(
            &ctx(),
            Position::new("body", index),
            vec![
                ("checked".into(), Any::Bool(true)),
                ("tag".into(), Any::from("consent")),
            ],
        )
        .unwrap();
        let payload = embed_payload(&doc, index);
        assert_eq!(payload.get("checked"), Some(&Any::Bool(true)));
        assert_eq!(payload.get("tag"), Some(&Any::from("consent")));

        doc.set_embed_attrs(
            &ctx(),
            Position::new("body", index),
            vec![("tag".into(), Any::Null)],
        )
        .unwrap();
        assert!(!embed_payload(&doc, index).contains_key("tag"));
    }

    #[test]
    fn set_embed_attrs_rejects_reserved_keys_and_non_embeds() {
        let doc = EditingDoc::new(7);
        let index = seed_sdt(&doc);
        let reserved = doc.set_embed_attrs(
            &ctx(),
            Position::new("body", index),
            vec![("_kind".into(), Any::from("image"))],
        );
        assert!(matches!(reserved, Err(OpError::ReservedKey(_))));
        // A pilcrow's paraId is schema-managed identity.
        let pilcrow = doc.set_embed_attrs(
            &ctx(),
            Position::new("body", 3),
            vec![("paraId".into(), Any::from("x"))],
        );
        assert!(matches!(pilcrow, Err(OpError::ReservedKey(_))));
        // Index 0 holds text, not an embed.
        let text = doc.set_embed_attrs(&ctx(), Position::new("body", 0), vec![]);
        assert!(matches!(text, Err(OpError::OutOfBounds { .. })));
    }

    #[test]
    fn stable_id_updates_and_null_clears() {
        let doc = EditingDoc::new(7);
        let index = seed_sdt(&doc);
        doc.set_embed_attrs_by_id(
            &ctx(),
            "control-1",
            vec![("value".into(), Any::from("yes"))],
        )
        .unwrap();
        assert_eq!(
            embed_payload(&doc, index).get("value"),
            Some(&Any::from("yes"))
        );
        doc.set_embed_attrs_by_id(&ctx(), "control-1", vec![("value".into(), Any::Null)])
            .unwrap();
        assert!(!embed_payload(&doc, index).contains_key("value"));
        assert!(matches!(
            doc.set_embed_attrs_by_id(&ctx(), "missing", vec![]),
            Err(OpError::UnknownEmbed(_))
        ));
    }

    #[test]
    fn set_embed_attrs_is_one_undoable_step() {
        let doc = EditingDoc::new(7);
        let index = seed_sdt(&doc);
        let mut undo = doc.undo_scope(&["body"]).unwrap();
        doc.set_embed_attrs(
            &ctx(),
            Position::new("body", index),
            vec![("checked".into(), Any::Bool(true))],
        )
        .unwrap();
        assert_eq!(
            embed_payload(&doc, index).get("checked"),
            Some(&Any::Bool(true))
        );
        assert!(undo.undo());
        assert_eq!(
            embed_payload(&doc, index).get("checked"),
            Some(&Any::Bool(false))
        );
    }

    #[test]
    fn exact_position_insert_is_addressable_and_suggesting_stamps_ins() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let suggesting = EditCtx::local("Ada", "2026-07-14T00:00:00Z").suggesting();
        let receipt = doc
            .insert_embed(&suggesting, Position::new("body", 1), "pageBreak", vec![])
            .unwrap();
        assert_eq!(receipt.revision_ids.len(), 1);
        let range = receipt.range.expect("insert receipt carries a range");
        let located = doc.locate_range(&range).unwrap();
        assert_eq!((located.start, located.end), (1, 2));
        let inserted = doc.story_segments("body").unwrap().into_iter().any(|segment| {
            matches!(segment.content, SegmentContent::OtherEmbed { ref kind, .. } if kind == "pageBreak")
                && matches!(segment.attributes.get("ins"), Some(value) if *value != Any::Null)
        });
        assert!(inserted);
    }
}
