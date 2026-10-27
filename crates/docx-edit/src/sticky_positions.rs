#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::HashMap;

use yrs::branch::{Branch, BranchPtr};
use yrs::updates::decoder::Decode;
use yrs::{Assoc, IndexedSequence, Offset, ReadTxn, StickyIndex, TextRef, Transact};

use crate::{EditingDoc, story_ref};

const POSITIONS_PER_STORY: usize = 32;

pub(crate) fn sticky_index<T: ReadTxn>(
    text: &TextRef,
    txn: &T,
    index: u32,
    assoc: Assoc,
) -> Option<StickyIndex> {
    #[cfg(test)]
    STICKY_INDEX_WALKS.with(|walks| walks.set(walks.get() + 1));
    text.sticky_index(txn, index, assoc)
}

pub(crate) fn sticky_offset<T: ReadTxn>(position: &StickyIndex, txn: &T) -> Option<Offset> {
    #[cfg(test)]
    STICKY_OFFSET_WALKS.with(|walks| walks.set(walks.get() + 1));
    position.get_offset(txn)
}

#[derive(Default)]
pub(super) struct StickyPositions {
    epoch: u64,
    stories: HashMap<String, StoryPositions>,
}

#[derive(Default)]
struct StoryPositions {
    indexes: HashMap<(u32, Assoc), Option<StickyIndex>>,
    offsets: HashMap<StickyIndex, Option<u32>>,
}

impl StickyPositions {
    fn story(&mut self, story: &str, epoch: u64) -> &mut StoryPositions {
        if self.epoch != epoch {
            self.stories.clear();
            self.epoch = epoch;
        }
        self.stories.entry(story.to_owned()).or_default()
    }
}

impl EditingDoc {
    pub(crate) fn cached_sticky_index(
        &self,
        story: &str,
        index: u32,
        assoc: Assoc,
    ) -> Result<Option<StickyIndex>, String> {
        let txn = self
            .yrs_doc()
            .try_transact()
            .map_err(|error| error.to_string())?;
        let text = story_ref(&txn, story).map_err(|error| error.to_string())?;
        let mut cache = self.sticky_positions.lock().unwrap();
        let cache = cache.story(story, self.committed_epoch());
        let key = (index, assoc);
        if let Some(position) = cache.indexes.get(&key) {
            return Ok(position.clone());
        }
        let position = sticky_index(&text, &txn, index, assoc);
        if cache.indexes.len() >= POSITIONS_PER_STORY {
            cache.indexes.clear();
        }
        cache.indexes.insert(key, position.clone());
        Ok(position)
    }

    pub(crate) fn cached_sticky_offset(
        &self,
        story: &str,
        position: &StickyIndex,
    ) -> Result<Option<u32>, String> {
        let txn = self
            .yrs_doc()
            .try_transact()
            .map_err(|error| error.to_string())?;
        let text = story_ref(&txn, story).map_err(|error| error.to_string())?;
        let mut cache = self.sticky_positions.lock().unwrap();
        let cache = cache.story(story, self.committed_epoch());
        if let Some(index) = cache.offsets.get(position) {
            return Ok(*index);
        }
        let offset = sticky_offset(position, &txn);
        let expected = BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(&text));
        if offset
            .as_ref()
            .is_some_and(|offset| offset.branch != expected)
        {
            return Err("selection sticky index belongs to another story".to_owned());
        }
        let index = offset.map(|offset| offset.index);
        if cache.offsets.len() >= POSITIONS_PER_STORY {
            cache.offsets.clear();
        }
        cache.offsets.insert(position.clone(), index);
        Ok(index)
    }

    pub(crate) fn resolve_cached_sticky_position(
        &self,
        story: &str,
        position: &[u8],
    ) -> Result<u32, String> {
        let position = StickyIndex::decode_v1(position).map_err(|error| error.to_string())?;
        self.cached_sticky_offset(story, &position)?
            .ok_or_else(|| "selection anchor no longer resolves".to_owned())
    }

    pub(crate) fn resolve_cached_sticky_selection(
        &self,
        story: &str,
        anchor: &StickyIndex,
        head: &StickyIndex,
    ) -> Result<(u32, u32), String> {
        let anchor_index = self.cached_sticky_offset(story, anchor);
        let head_index = if anchor == head {
            anchor_index.clone()
        } else {
            self.cached_sticky_offset(story, head)
        };
        match (anchor_index, head_index) {
            (Ok(None), _) => Err("selection anchor no longer resolves".to_owned()),
            (_, Ok(None)) => Err("selection head no longer resolves".to_owned()),
            (Ok(Some(anchor)), Ok(Some(head))) => Ok((anchor, head)),
            (Err(error), _) | (_, Err(error)) => Err(error),
        }
    }
}

#[cfg(test)]
thread_local! {
    pub(crate) static STICKY_INDEX_WALKS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    pub(crate) static STICKY_OFFSET_WALKS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;

    use yrs::updates::encoder::Encode;
    use yrs::{Any, ClientID, ID, Text};

    use crate::{EditCtx, FormatPolicy, Position, RawOp, StoryRange};

    fn resolve_uncached_sticky_position(
        doc: &EditingDoc,
        story: &str,
        bytes: &[u8],
    ) -> Result<u32, String> {
        let position = StickyIndex::decode_v1(bytes).map_err(|error| error.to_string())?;
        let txn = doc.yrs_doc().transact();
        let text = story_ref(&txn, story).map_err(|error| error.to_string())?;
        let offset = position
            .get_offset(&txn)
            .ok_or_else(|| "selection anchor no longer resolves".to_owned())?;
        if offset.branch != BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(&text)) {
            return Err("selection sticky index belongs to another story".to_owned());
        }
        Ok(offset.index)
    }

    fn assert_matches_uncached(doc: &EditingDoc, story: &str) {
        let txn = doc.yrs_doc().transact();
        let text = story_ref(&txn, story).unwrap();
        for index in 0..=text.len(&txn) + 1 {
            for assoc in [Assoc::Before, Assoc::After] {
                let expected = text.sticky_index(&txn, index, assoc);
                for _ in 0..2 {
                    let actual = doc.cached_sticky_index(story, index, assoc).unwrap();
                    assert_eq!(
                        actual.as_ref().map(|position| position.encode_v1()),
                        expected.as_ref().map(|position| position.encode_v1()),
                        "{story} at {index}, {assoc:?}"
                    );
                    if let Some(position) = actual {
                        assert_eq!(
                            doc.cached_sticky_offset(story, &position).unwrap(),
                            position.get_offset(&txn).map(|offset| offset.index)
                        );
                        let bytes = position.encode_v1();
                        assert_eq!(
                            doc.resolve_cached_sticky_position(story, &bytes),
                            resolve_uncached_sticky_position(doc, story, &bytes)
                        );
                    }
                }
            }
        }
    }

    fn assert_retained_positions(doc: &EditingDoc, story: &str, positions: &[StickyIndex]) {
        for position in positions {
            let bytes = position.encode_v1();
            for _ in 0..2 {
                assert_eq!(
                    doc.resolve_cached_sticky_position(story, &bytes),
                    resolve_uncached_sticky_position(doc, story, &bytes)
                );
            }
        }
    }

    #[test]
    fn transaction_positions_bypass_cached_epoch() {
        let doc = EditingDoc::new(1);
        doc.create_story("body", "abcdef", "Normal", "left")
            .unwrap();
        let position = doc
            .cached_sticky_index("body", 3, Assoc::After)
            .unwrap()
            .unwrap();
        assert_eq!(
            doc.cached_sticky_offset("body", &position).unwrap(),
            Some(3)
        );
        let epoch = doc.committed_epoch();
        STICKY_INDEX_WALKS.with(|walks| walks.set(0));
        STICKY_OFFSET_WALKS.with(|walks| walks.set(0));
        let mut txn = doc.yrs_doc().transact_mut();
        assert!(doc.cached_sticky_index("body", 3, Assoc::After).is_err());
        assert!(doc.cached_sticky_index("body", 2, Assoc::Before).is_err());
        assert!(doc.cached_sticky_offset("body", &position).is_err());
        let text = story_ref(&txn, "body").unwrap();
        text.insert(&mut txn, 0, "x");
        assert_eq!(doc.committed_epoch(), epoch);
        assert_eq!(sticky_offset(&position, &txn).unwrap().index, 4);
        assert_ne!(
            sticky_index(&text, &txn, 3, Assoc::After).unwrap(),
            position
        );
        STICKY_INDEX_WALKS.with(|walks| assert_eq!(walks.get(), 1));
        STICKY_OFFSET_WALKS.with(|walks| assert_eq!(walks.get(), 1));
        drop(txn);
        assert_eq!(
            doc.cached_sticky_offset("body", &position).unwrap(),
            Some(4)
        );
    }

    #[test]
    fn cached_positions_invalidate_after_forced_gc() {
        let doc = EditingDoc::new(1);
        doc.keep_deleted();
        doc.create_story("body", "abcdef", "Normal", "left")
            .unwrap();
        let position = doc
            .cached_sticky_index("body", 3, Assoc::After)
            .unwrap()
            .unwrap();
        doc.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 1, 5))
            .unwrap();
        assert_matches_uncached(&doc, "body");
        assert_retained_positions(&doc, "body", std::slice::from_ref(&position));
        let epoch = doc.committed_epoch();
        let before = doc.encode_state_as_update_v1();
        {
            let mut txn = doc.yrs_doc().transact_mut();
            txn.gc(None);
        }
        assert_eq!(doc.committed_epoch(), epoch);
        assert_ne!(doc.encode_state_as_update_v1(), before);
        doc.invalidate_projections();
        assert_eq!(doc.committed_epoch(), epoch + 1);
        let txn = doc.yrs_doc().transact();
        let text = story_ref(&txn, "body").unwrap();
        let expected = text.sticky_index(&txn, 1, Assoc::After);
        let expected_offset = position.get_offset(&txn).map(|offset| offset.index);
        drop(txn);
        STICKY_INDEX_WALKS.with(|walks| walks.set(0));
        STICKY_OFFSET_WALKS.with(|walks| walks.set(0));
        assert_eq!(
            doc.cached_sticky_index("body", 1, Assoc::After).unwrap(),
            expected
        );
        assert_eq!(
            doc.cached_sticky_offset("body", &position).unwrap(),
            expected_offset
        );
        STICKY_INDEX_WALKS.with(|walks| assert_eq!(walks.get(), 1));
        STICKY_OFFSET_WALKS.with(|walks| assert_eq!(walks.get(), 1));
        assert_matches_uncached(&doc, "body");
        assert_retained_positions(&doc, "body", &[position]);
    }

    #[test]
    fn cached_positions_match_yrs_for_both_associations_and_story_boundaries() {
        let long = "abcdefgh".repeat(12);
        for initial in ["", "abc", "A😀e\u{301}中Z", long.as_str()] {
            let doc = EditingDoc::new(1);
            for story in ["body", "header:rId7", "body:table:cell"] {
                doc.create_story(story, initial, "Normal", "left").unwrap();
            }
            for story in ["body", "header:rId7", "body:table:cell"] {
                assert_matches_uncached(&doc, story);
            }
        }
    }

    #[test]
    fn cached_positions_match_yrs_with_formatting_embeds_and_deleted_items() {
        for keep_deleted in [false, true] {
            let doc = EditingDoc::new(1);
            if keep_deleted {
                doc.keep_deleted();
            }
            doc.create_story("body", "ab😀cdef", "Normal", "left")
                .unwrap();
            doc.insert_text(
                &EditCtx::local("", ""),
                Position::new("body", 1),
                "XY",
                FormatPolicy::Explicit(BTreeMap::from([("bold".into(), Any::Bool(true))])),
            )
            .unwrap();
            doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 3), None)
                .unwrap();
            doc.apply_raw_ops(
                "body",
                vec![RawOp::InsertEmbed {
                    index: 1,
                    kind: "break".into(),
                    payload: Vec::new(),
                    attrs: Default::default(),
                }],
                &EditCtx::local("", ""),
            )
            .unwrap();
            assert_matches_uncached(&doc, "body");
            let positions: Vec<_> = [Assoc::Before, Assoc::After]
                .into_iter()
                .map(|assoc| doc.cached_sticky_index("body", 2, assoc).unwrap().unwrap())
                .collect();
            assert_retained_positions(&doc, "body", &positions);
            doc.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 1, 3))
                .unwrap();
            assert_retained_positions(&doc, "body", &positions);
            assert_matches_uncached(&doc, "body");
        }
    }

    #[test]
    fn cached_positions_invalidate_after_local_remote_and_other_story_edits() {
        let doc = EditingDoc::new(1);
        doc.create_story("body", "abcdef", "Normal", "left")
            .unwrap();
        doc.create_story("header:rId7", "abcdef", "Normal", "left")
            .unwrap();
        let peer = EditingDoc::new(2);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        let positions: Vec<_> = [Assoc::Before, Assoc::After]
            .into_iter()
            .map(|assoc| doc.cached_sticky_index("body", 3, assoc).unwrap().unwrap())
            .collect();
        assert_retained_positions(&doc, "body", &positions);
        doc.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 3),
            "L",
            FormatPolicy::Plain,
        )
        .unwrap();
        assert_eq!(
            doc.cached_sticky_offset("body", &positions[0]).unwrap(),
            Some(3)
        );
        assert_eq!(
            doc.cached_sticky_offset("body", &positions[1]).unwrap(),
            Some(4)
        );
        assert_retained_positions(&doc, "body", &positions);
        assert_matches_uncached(&doc, "body");
        peer.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 1),
            "R😀",
            FormatPolicy::Plain,
        )
        .unwrap();
        doc.apply_update_v1(&peer.encode_state_as_update_v1())
            .unwrap();
        assert_retained_positions(&doc, "body", &positions);
        assert_matches_uncached(&doc, "body");
        peer.insert_text(
            &EditCtx::local("", ""),
            Position::new("header:rId7", 0),
            "H",
            FormatPolicy::Plain,
        )
        .unwrap();
        doc.apply_update_v1(&peer.encode_state_as_update_v1())
            .unwrap();
        assert_retained_positions(&doc, "body", &positions);
        assert_matches_uncached(&doc, "body");
        assert_matches_uncached(&doc, "header:rId7");
    }

    #[test]
    fn cached_positions_invalidate_after_undo_and_redo() {
        let doc = EditingDoc::new(1);
        doc.create_story("body", "abcdef", "Normal", "left")
            .unwrap();
        let mut undo = doc.undo_manager();
        doc.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 3),
            "XY",
            FormatPolicy::Plain,
        )
        .unwrap();
        let positions: Vec<_> = [Assoc::Before, Assoc::After]
            .into_iter()
            .map(|assoc| doc.cached_sticky_index("body", 4, assoc).unwrap().unwrap())
            .collect();
        assert_retained_positions(&doc, "body", &positions);
        assert_matches_uncached(&doc, "body");
        assert!(undo.undo());
        assert_retained_positions(&doc, "body", &positions);
        assert_matches_uncached(&doc, "body");
        assert!(undo.redo());
        assert_retained_positions(&doc, "body", &positions);
        assert_matches_uncached(&doc, "body");
    }

    #[test]
    fn cached_positions_preserve_resolution_errors_and_story_identity() {
        let doc = EditingDoc::new(1);
        doc.create_story("body", "abcdef", "Normal", "left")
            .unwrap();
        doc.create_story("header:rId7", "abcdef", "Normal", "left")
            .unwrap();
        let position = doc
            .cached_sticky_index("body", 3, Assoc::After)
            .unwrap()
            .unwrap();
        let unknown = StickyIndex::from_id(ID::new(ClientID::new(999), 0), Assoc::After);
        for bytes in [position.encode_v1(), unknown.encode_v1(), vec![], vec![255]] {
            for story in ["body", "header:rId7", "missing"] {
                for _ in 0..2 {
                    assert_eq!(
                        doc.resolve_cached_sticky_position(story, &bytes),
                        resolve_uncached_sticky_position(&doc, story, &bytes)
                    );
                }
            }
        }
        doc.delete_story("body").unwrap();
        assert_retained_positions(&doc, "body", std::slice::from_ref(&position));
        assert!(doc.cached_sticky_index("body", 3, Assoc::After).is_err());
        doc.create_story("body", "replacement", "Normal", "left")
            .unwrap();
        assert_retained_positions(&doc, "body", &[position]);
        assert_matches_uncached(&doc, "body");
    }

    #[test]
    fn repeated_positions_walk_once_per_document_version() {
        let doc = EditingDoc::new(1);
        doc.create_story("body", "abcdef", "Normal", "left")
            .unwrap();
        doc.create_story("header:rId7", "header", "Normal", "left")
            .unwrap();
        let peer = EditingDoc::new(2);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        STICKY_INDEX_WALKS.with(|walks| walks.set(0));
        STICKY_OFFSET_WALKS.with(|walks| walks.set(0));
        let read = |expected_walks| {
            for _ in 0..5 {
                let position = doc
                    .cached_sticky_index("body", 3, Assoc::After)
                    .unwrap()
                    .unwrap();
                let bytes = position.encode_v1();
                assert_eq!(
                    doc.resolve_cached_sticky_position("body", &bytes).unwrap(),
                    3
                );
                assert_eq!(
                    doc.cached_sticky_offset("body", &position).unwrap(),
                    Some(3)
                );
            }
            STICKY_INDEX_WALKS.with(|walks| assert_eq!(walks.get(), expected_walks));
            STICKY_OFFSET_WALKS.with(|walks| assert_eq!(walks.get(), expected_walks));
        };
        read(1);
        doc.insert_text(
            &EditCtx::local("", ""),
            Position::new("header:rId7", 0),
            "H",
            FormatPolicy::Plain,
        )
        .unwrap();
        read(2);
        peer.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 5),
            "R",
            FormatPolicy::Plain,
        )
        .unwrap();
        doc.apply_update_v1(&peer.encode_state_as_update_v1())
            .unwrap();
        read(3);
    }
}
