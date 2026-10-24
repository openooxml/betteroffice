//! Map-backed embed operations.

use yrs::types::AsPrelim;
use yrs::types::text::YChange;
use yrs::{Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef, Transact};

use crate::control_values::{guard_embed_insert, guard_embed_write};
use crate::op::{OpError, OpResult, Receipt, loc_range_in_txn};
use crate::ops::bookmarks::Carry;
use crate::ops::{
    ChunkKind, adjacent_paragraph_change_revision_id, adjacent_revision_id, past_leading_blocks,
    position_chunks,
};
use crate::{
    EditCtx, EditingDoc, INS, KIND_KEY, PILCROW_KIND, Position, check_position, insertion_attrs,
    is_pilcrow, out_len, revision_value, story_ref,
};

/// Finds the map-backed embed sitting exactly at story `index` (any kind,
/// pilcrows included).
pub(crate) fn embed_map_at<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> OpResult<MapRef> {
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

/// Finds one map-backed embed, and the story holding it, by its stable authored
/// payload identity. New yrs inserts use `embedId`; `id` (SDTs) and `rId`
/// (images) keep mirrored embeds addressable without rewriting their payload
/// vocabulary.
fn embed_by_id<T: ReadTxn>(txn: &T, embed_id: &str) -> OpResult<(String, MapRef)> {
    let stories = txn
        .get_map(crate::STORIES)
        .ok_or_else(|| OpError::UnknownEmbed(embed_id.to_owned()))?;
    let mut story_ids: Vec<String> = crate::live_stories(&stories, txn)
        .map(|(id, _)| id.to_owned())
        .collect();
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
                return Ok((story_id, map));
            }
        }
    }
    Err(OpError::UnknownEmbed(embed_id.to_owned()))
}

/// [`embed_by_id`] with the embed's story index.
#[cfg(feature = "wasm")]
fn embed_position_by_id<T: ReadTxn>(txn: &T, embed_id: &str) -> OpResult<(String, u32)> {
    let (story_id, target) = embed_by_id(txn, embed_id)?;
    let story = story_ref(txn, &story_id)?;
    let mut offset = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        if let Out::YMap(map) = &diff.insert
            && *map == target
        {
            return Ok((story_id, offset));
        }
        offset += out_len(&diff.insert);
    }
    Err(OpError::UnknownEmbed(embed_id.to_owned()))
}

impl EditingDoc {
    /// The story and story index of the embed carrying `embed_id`.
    #[cfg(feature = "wasm")]
    pub(crate) fn embed_position(&self, embed_id: &str) -> OpResult<(String, u32)> {
        embed_position_by_id(&self.yrs_doc().transact(), embed_id)
    }

    /// The story holding the embed carrying `embed_id`.
    pub fn embed_story(&self, embed_id: &str) -> OpResult<String> {
        embed_by_id(&self.yrs_doc().transact(), embed_id).map(|(story, _)| story)
    }

    /// The `_kind` of the map-backed embed at `at`. Errors when `at` does not
    /// hold a map-backed embed.
    pub fn embed_kind(&self, at: &Position) -> OpResult<Option<String>> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &at.story)?;
        let map = embed_map_at(&story, &txn, at.index)?;
        Ok(crate::map_string(&map, &txn, KIND_KEY))
    }

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
            if key == KIND_KEY || (pilcrow && crate::is_identity_key(key)) {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        let retyped = guard_embed_write(
            &map,
            &txn,
            entries.iter().map(|(key, value)| (key.as_str(), value)),
        )?;
        for (key, value) in entries {
            if value == Any::Null {
                map.remove(&mut txn, &key);
            } else {
                map.insert(&mut txn, key, value);
            }
        }
        if retyped {
            map.remove(&mut txn, "value");
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
            if crate::is_identity_key(key) {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        let mut txn = self.transact_for(ctx);
        let (_, map) = embed_by_id(&txn, embed_id)?;
        let retyped = guard_embed_write(
            &map,
            &txn,
            entries.iter().map(|(key, value)| (key.as_str(), value)),
        )?;
        for (key, value) in entries {
            if value == Any::Null {
                map.remove(&mut txn, &key);
            } else {
                map.insert(&mut txn, key, value);
            }
        }
        if retyped {
            map.remove(&mut txn, "value");
        }
        Ok(Receipt::default())
    }

    /// Moves one inline embed, preserving its payload and run attributes.
    pub fn move_embed(&self, ctx: &EditCtx, at: Position, mut to: Position) -> OpResult<Receipt> {
        if ctx.is_suggesting() {
            return Err(OpError::OverlapsTrackedChange);
        }
        let root = |story: &str| {
            if story == "body" || story.starts_with("body:") {
                "body".to_owned()
            } else {
                story.split(':').take(2).collect::<Vec<_>>().join(":")
            }
        };
        if root(&at.story) != root(&to.story) {
            return Err(OpError::NotTopLevel);
        }
        let mut txn = self.transact_for(ctx);
        let source = story_ref(&txn, &at.story)?;
        let destination = story_ref(&txn, &to.story)?;
        let embed = embed_map_at(&source, &txn, at.index)?;
        let kind = crate::map_string(&embed, &txn, KIND_KEY).unwrap_or_default();
        if is_pilcrow(&embed, &txn) || crate::segments::is_block_embed(&kind) {
            return Err(OpError::NotTopLevel);
        }
        let attrs = source
            .diff(&txn, YChange::identity)
            .into_iter()
            .find(|diff| matches!(&diff.insert, Out::YMap(map) if *map == embed))
            .expect("located embed has a text diff")
            .attributes
            .map(|attrs| *attrs)
            .unwrap_or_default();
        if attrs
            .get(crate::DEL)
            .is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
        {
            return Err(OpError::LockedSdt);
        }
        let ownership = crate::policy::Ownership::build(self, &txn);
        for story in [&at.story, &to.story] {
            let chain = ownership.chain(story).map_err(|_| OpError::LockedSdt)?;
            if chain.iter().any(|owner| owner.content_locked()) {
                return Err(OpError::LockedSdt);
            }
            if chain.iter().any(|owner| owner.revision) {
                return Err(OpError::OverlapsTrackedChange);
            }
        }
        check_position(&destination, &txn, to.index)?;
        to.index = past_leading_blocks(&destination, &txn, to.index)
            .min(destination.len(&txn).saturating_sub(1));
        let _ = position_chunks(&destination, &txn, &mut to.index);
        if at.story == to.story && (to.index == at.index || to.index == at.index + 1) {
            return Ok(Receipt {
                range: Some(loc_range_in_txn(
                    &at.story,
                    &source,
                    &txn,
                    at.index,
                    at.index + 1,
                )?),
                ..Receipt::default()
            });
        }
        let payload = embed.as_prelim(&txn);
        crate::identity::promote_at(self, &mut txn, &at.story, &source, at.index);
        crate::identity::promote_at(self, &mut txn, &to.story, &destination, to.index);
        let carry = Carry::capture(&source, &txn, at.index, at.index + 1, None);
        crate::ops::remove_units(
            &mut txn,
            &source,
            at.index,
            1,
            &[],
            crate::ops::After::Changed,
        );
        if let Some(carry) = carry {
            carry.removed(&mut txn, at.index, at.index + 1);
        }
        if at.story == to.story && to.index > at.index {
            to.index -= 1;
        }
        let carry = Carry::capture(&destination, &txn, to.index, to.index, None);
        destination.insert_embed_with_attributes(&mut txn, to.index, payload, attrs);
        if let Some(carry) = carry {
            carry.inserted(&mut txn, to.index, 1);
        }
        Ok(Receipt {
            range: Some(loc_range_in_txn(
                &to.story,
                &destination,
                &txn,
                to.index,
                to.index + 1,
            )?),
            ..Receipt::default()
        })
    }

    /// Inserts a non-pilcrow map-backed embed at the exact story position.
    /// The caller supplies the typed kind/payload vocabulary; suggesting mode
    /// stamps the embed with one insertion revision.
    pub fn insert_embed(
        &self,
        ctx: &EditCtx,
        mut at: Position,
        kind: &str,
        payload: Vec<(String, Any)>,
    ) -> OpResult<Receipt> {
        if kind == PILCROW_KIND {
            return Err(OpError::ReservedKey(kind.to_owned()));
        }
        for (key, _) in &payload {
            if crate::is_identity_key(key) {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        guard_embed_insert(
            kind,
            payload.iter().map(|(key, value)| (key.as_str(), value)),
        )?;
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        if !crate::segments::is_block_embed(kind) {
            at.index = past_leading_blocks(&story, &txn, at.index);
        }
        crate::identity::promote_at(self, &mut txn, &at.story, &story, at.index);
        let chunks = position_chunks(&story, &txn, &mut at.index);
        let append_paragraph = kind == "pageBreak"
            && at.story == "body"
            && at.index + 1 == story.len(&txn)
            && chunks.iter().any(|chunk| {
                chunk.start == at.index && matches!(chunk.kind, ChunkKind::Pilcrow(_))
            });
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
        let carry = (!crate::segments::is_block_embed(kind))
            .then(|| Carry::capture(&story, &txn, at.index, at.index, None))
            .flatten();
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
        if let Some(carry) = carry {
            carry.inserted(&mut txn, at.index, 1);
        }
        let new_para_ids = if append_paragraph {
            let split = self.split_paragraph_in_transaction(
                ctx,
                &mut txn,
                Position::new(&at.story, at.index + 1),
                None,
                revision_id.clone(),
            )?;
            vec![split.second_para_id]
        } else {
            Vec::new()
        };
        let range = loc_range_in_txn(&at.story, &story, &txn, at.index, at.index + 1)?;
        Ok(Receipt {
            new_para_ids,
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
    fn move_embed_refuses_tracked_owner_without_mutation() {
        for from_cell in [false, true] {
            let doc = EditingDoc::new(7);
            doc.create_story("body", "AB", "Normal", "left").unwrap();
            doc.create_story("body:t0:r0c0", "CD", "Normal", "left")
                .unwrap();
            let Any::Map(table) = Any::from_json(
                r#"{"rows":[{"trPr":{"trIns":{"id":"pending","author":"Ada"}},"cells":[{"story":"body:t0:r0c0"}]}]}"#,
            )
            .unwrap() else {
                unreachable!()
            };
            doc.apply_raw_ops(
                "body",
                vec![crate::RawOp::InsertEmbed {
                    index: 0,
                    kind: "table".into(),
                    payload: table.iter().map(|(k, v)| (k.clone(), v.clone())).collect(),
                    attrs: Default::default(),
                }],
                &ctx(),
            )
            .unwrap();
            let (source, destination) = if from_cell {
                ("body:t0:r0c0", "body")
            } else {
                ("body", "body:t0:r0c0")
            };
            doc.apply_raw_ops(
                source,
                vec![crate::RawOp::InsertEmbed {
                    index: 1,
                    kind: "image".into(),
                    payload: vec![("rId".into(), Any::from("rIdImage"))],
                    attrs: Default::default(),
                }],
                &ctx(),
            )
            .unwrap();
            let before = doc.yrs_doc().transact().state_vector();
            assert_eq!(
                doc.move_embed(
                    &ctx(),
                    Position::new(source, 1),
                    Position::new(destination, 2)
                ),
                Err(OpError::OverlapsTrackedChange)
            );
            assert_eq!(doc.yrs_doc().transact().state_vector(), before);
        }
    }

    #[test]
    fn move_embed_preserves_payload_attrs_and_single_undo() {
        for (to, landed) in [(5, 4), (0, 0), (1, 1), (2, 1)] {
            let doc = EditingDoc::new(7);
            doc.create_story("body", "ABCD", "Normal", "left").unwrap();
            let payload =
                Any::from_json(r#"{"rId":"rIdImage","locks":[true,{"aspect":false}]}"#).unwrap();
            let Any::Map(entries) = payload else {
                unreachable!()
            };
            doc.apply_raw_ops(
                "body",
                vec![crate::RawOp::InsertEmbed {
                    index: 1,
                    kind: "image".into(),
                    payload: entries
                        .iter()
                        .map(|(k, v)| (k.clone(), v.clone()))
                        .collect(),
                    attrs: [
                        ("bold".into(), Any::Bool(true)),
                        (
                            crate::INS.into(),
                            Any::from_json(r#"{"id":"pending","author":"Ada"}"#).unwrap(),
                        ),
                        (
                            "custom".into(),
                            Any::from_json(r#"{"nested":[1,"x"]}"#).unwrap(),
                        ),
                    ]
                    .into(),
                }],
                &ctx(),
            )
            .unwrap();
            let original = doc.story_segments("body").unwrap();
            let vector = doc.yrs_doc().transact().state_vector();
            let mut undo = doc.undo_manager();
            let receipt = doc
                .move_embed(&ctx(), Position::new("body", 1), Position::new("body", to))
                .unwrap();
            let range = doc.locate_range(&receipt.range.unwrap()).unwrap();
            assert_eq!((range.start, range.end), (landed, landed + 1));
            let moved = doc.story_segments("body").unwrap().into_iter().find(|s| matches!(&s.content, SegmentContent::OtherEmbed { kind, .. } if kind == "image")).unwrap();
            let before = original.iter().find(|s| matches!(&s.content, SegmentContent::OtherEmbed { kind, .. } if kind == "image")).unwrap();
            assert_eq!(&moved, before);
            if to == 1 || to == 2 {
                assert_eq!(doc.yrs_doc().transact().state_vector(), vector);
                assert!(!undo.undo());
            } else {
                assert!(undo.undo());
                assert_eq!(doc.story_segments("body").unwrap(), original);
                assert!(!undo.undo());
            }
        }
    }

    #[test]
    fn move_embed_across_paragraphs_and_stories() {
        let doc = EditingDoc::new(7);
        seed_sdt(&doc);
        doc.create_story("body:t0:r0c0", "CD", "Normal", "left")
            .unwrap();
        let original = embed_payload(&doc, 1);
        doc.split_paragraph(&ctx(), Position::new("body", 3), None)
            .unwrap();
        doc.move_embed(&ctx(), Position::new("body", 1), Position::new("body", 4))
            .unwrap();
        assert_eq!(embed_payload(&doc, 3), original);
        doc.move_embed(
            &ctx(),
            Position::new("body", 3),
            Position::new("body:t0:r0c0", 1),
        )
        .unwrap();
        assert_eq!(doc.embed_story("control-1").unwrap(), "body:t0:r0c0");
    }

    #[test]
    fn move_embed_copies_nested_shared_maps() {
        use yrs::types::ToJson;
        let doc = EditingDoc::new(7);
        seed_sdt(&doc);
        let before = {
            let mut txn = doc.transact_for(&ctx());
            let story = story_ref(&txn, "body").unwrap();
            let map = embed_map_at(&story, &txn, 1).unwrap();
            let child = map.insert(&mut txn, "nested", MapPrelim::default());
            child.insert(&mut txn, "value", Any::from("kept"));
            map.to_json(&txn)
        };
        doc.move_embed(&ctx(), Position::new("body", 1), Position::new("body", 3))
            .unwrap();
        let txn = doc.yrs_doc().transact();
        let story = story_ref(&txn, "body").unwrap();
        let map = embed_map_at(&story, &txn, 2).unwrap();
        assert_eq!(map.to_json(&txn), before);
        assert!(matches!(map.get(&txn, "nested"), Some(Out::YMap(_))));
    }

    #[test]
    fn move_embed_refuses_tracked_deletion_with_locked_sdt_before_mutation() {
        let doc = EditingDoc::new(7);
        seed_sdt(&doc);
        {
            let mut txn = doc.transact_for(&ctx());
            let story = story_ref(&txn, "body").unwrap();
            story.format(
                &mut txn,
                1,
                1,
                [(crate::DEL.into(), Any::from("pending"))].into(),
            );
        }
        let before = doc.story_segments("body").unwrap();
        assert_eq!(
            doc.move_embed(&ctx(), Position::new("body", 1), Position::new("body", 3)),
            Err(OpError::LockedSdt)
        );
        assert_eq!(doc.story_segments("body").unwrap(), before);
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
    fn embed_kind_reads_the_embed_at_a_position() {
        let doc = EditingDoc::new(7);
        let index = seed_sdt(&doc);
        let kind = |index| doc.embed_kind(&Position::new("body", index));
        assert_eq!(kind(index).unwrap().as_deref(), Some("sdt"));
        assert_eq!(kind(3).unwrap().as_deref(), Some(PILCROW_KIND));
        assert!(matches!(kind(0), Err(OpError::OutOfBounds { .. })));
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
    fn embed_story_names_the_story_holding_the_embed() {
        let doc = EditingDoc::new(7);
        seed_sdt(&doc);
        assert_eq!(doc.embed_story("control-1").unwrap(), "body");
        assert!(matches!(
            doc.embed_story("missing"),
            Err(OpError::UnknownEmbed(id)) if id == "missing"
        ));
    }

    #[test]
    fn set_embed_attrs_is_one_undoable_step() {
        let doc = EditingDoc::new(7);
        let index = seed_sdt(&doc);
        let mut undo = doc.undo_manager();
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
