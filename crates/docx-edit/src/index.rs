//! Cached paragraph and embed geometry over the committed document.
//!
//! Every `paraId`- and index-keyed lookup that used to re-materialize a
//! story's segments resolves against this snapshot instead; the yrs doc
//! remains the source of truth. A `Doc` update observer advances the epoch
//! on each committed transaction — local ops, remote updates, and undo/redo
//! alike — and the next lookup rebuilds lazily. The cache cannot see a
//! transaction's uncommitted writes, so mutating ops may only consult it at
//! validation time, before their first write, or from a read transaction.

use std::collections::HashMap;

use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef};

use crate::{KIND_KEY, PARA_ID, STORIES, is_pilcrow, map_string, out_len};

/// `_kind` values whose embeds get their own layout block.
fn is_block_embed(kind: &str) -> bool {
    matches!(kind, "table" | "blockSdt" | "pageBreak" | "columnBreak")
}

/// Keys a non-pilcrow map embed may be addressed by.
const EMBED_ID_KEYS: [&str; 3] = ["embedId", "id", "rId"];

/// One paragraph's geometry inside its story, plus its pilcrow map.
#[derive(Debug)]
pub(crate) struct IndexedPara {
    pub para_id: String,
    /// Story index of the paragraph's first unit.
    pub start: u32,
    /// Story index of the paragraph's pilcrow embed.
    pub pilcrow: u32,
    /// `start` advanced past leading consecutive block embeds — the
    /// node-offset base `index_loc` reports (read by the wasm surface only).
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub node_start: u32,
    pub map: MapRef,
}

impl IndexedPara {
    /// UTF-16 units between `start` and `pilcrow`, exclusive of the mark.
    pub fn len(&self) -> u32 {
        self.pilcrow - self.start
    }
}

/// One non-pilcrow map embed, with the values its authored ids may take.
#[derive(Debug)]
pub(crate) struct IndexedEmbed {
    pub map: MapRef,
    /// Present values of the [`EMBED_ID_KEYS`] entries.
    pub ids: Vec<Any>,
}

/// Paragraph and embed table for one story.
#[derive(Debug)]
pub(crate) struct StoryIndex {
    pub story_id: String,
    pub story: TextRef,
    /// Paragraphs in document order, each ending at its `pilcrow`.
    pub paras: Vec<IndexedPara>,
    /// Non-pilcrow map embeds in document order.
    pub embeds: Vec<IndexedEmbed>,
}

impl StoryIndex {
    /// The paragraph containing story `index` — the first whose pilcrow
    /// is >= it.
    pub fn para_at(&self, index: u32) -> Option<&IndexedPara> {
        let slot = self.paras.partition_point(|para| index > para.pilcrow);
        self.paras.get(slot)
    }
}

/// Paragraph and embed geometry snapshot at one committed document epoch.
#[derive(Debug)]
pub(crate) struct ParaIndex {
    /// [`EditingDoc::doc_epoch`] value the snapshot was built against.
    pub epoch: u64,
    /// Stories sorted by id — the iteration order the previous full scans
    /// used for every resolved set.
    pub stories: Vec<StoryIndex>,
    /// `para_id` -> (story slot, para slot); first occurrence wins when a
    /// divergent merge left duplicate ids behind.
    by_para: HashMap<String, (usize, usize)>,
}

impl ParaIndex {
    pub(crate) fn build<T: ReadTxn>(txn: &T, epoch: u64) -> Self {
        let mut stories = Vec::new();
        if let Some(story_map) = txn.get_map(STORIES) {
            let mut story_ids: Vec<String> =
                story_map.keys(txn).map(|key| key.to_string()).collect();
            story_ids.sort();
            for story_id in story_ids {
                let Some(Out::YText(story)) = story_map.get(txn, &story_id) else {
                    continue;
                };
                stories.push(StoryIndex::build(story_id, &story, txn));
            }
        }
        let mut by_para = HashMap::new();
        for (story_slot, story) in stories.iter().enumerate() {
            for (para_slot, para) in story.paras.iter().enumerate() {
                by_para
                    .entry(para.para_id.clone())
                    .or_insert((story_slot, para_slot));
            }
        }
        Self {
            epoch,
            stories,
            by_para,
        }
    }

    /// The story's geometry, if `story_id` is a story of the doc.
    pub fn story(&self, story_id: &str) -> Option<&StoryIndex> {
        self.stories.iter().find(|story| story.story_id == story_id)
    }

    /// `para_id` resolved inside `story_id` — a para that lives in another
    /// story is not found, matching the story-scoped scans this replaces.
    pub fn para_in(&self, story_id: &str, para_id: &str) -> Option<&IndexedPara> {
        self.story(story_id)?
            .paras
            .iter()
            .find(|para| para.para_id == para_id)
    }

    /// `para_id` resolved across every story — the first in sorted-story
    /// order, as the previous whole-document scans returned.
    pub fn para_anywhere(&self, para_id: &str) -> Option<(&StoryIndex, &IndexedPara)> {
        let (story_slot, para_slot) = *self.by_para.get(para_id)?;
        let story = &self.stories[story_slot];
        Some((story, &story.paras[para_slot]))
    }
}

impl StoryIndex {
    fn build<T: ReadTxn>(story_id: String, story: &TextRef, txn: &T) -> Self {
        let mut paras = Vec::new();
        let mut embeds = Vec::new();
        let mut offset = 0_u32;
        let mut para_start = 0_u32;
        let mut node_start = 0_u32;
        for diff in story.diff(txn, YChange::identity) {
            let len = out_len(&diff.insert);
            if let Out::YMap(map) = &diff.insert {
                if is_pilcrow(map, txn) {
                    paras.push(IndexedPara {
                        para_id: map_string(map, txn, PARA_ID).unwrap_or_default(),
                        start: para_start,
                        pilcrow: offset,
                        node_start,
                        map: map.clone(),
                    });
                    offset += 1;
                    para_start = offset;
                    node_start = offset;
                    continue;
                }
                let kind = map_string(map, txn, KIND_KEY).unwrap_or_default();
                if offset == node_start && is_block_embed(&kind) {
                    node_start = offset + 1;
                }
                let mut ids = Vec::new();
                for key in EMBED_ID_KEYS {
                    if let Some(Out::Any(value)) = map.get(txn, key) {
                        ids.push(value);
                    }
                }
                embeds.push(IndexedEmbed {
                    map: map.clone(),
                    ids,
                });
            }
            offset += len;
        }
        Self {
            story_id,
            story: story.clone(),
            paras,
            embeds,
        }
    }
}
