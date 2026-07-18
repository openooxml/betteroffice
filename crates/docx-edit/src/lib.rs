// Ported from openooxml/docx, which did not gate on clippy style lints;
// burning these down is tracked follow-up work, not a merge blocker.
#![allow(
    clippy::cloned_ref_to_slice_refs,
    clippy::collapsible_if,
    clippy::collapsible_match,
    clippy::doc_lazy_continuation,
    clippy::excessive_precision,
    clippy::field_reassign_with_default,
    clippy::if_same_then_else,
    clippy::inconsistent_digit_grouping,
    clippy::items_after_test_module,
    clippy::large_enum_variant,
    clippy::manual_contains,
    clippy::manual_is_multiple_of,
    clippy::manual_pattern_char_comparison,
    clippy::manual_repeat_n,
    clippy::manual_unwrap_or,
    clippy::map_clone,
    clippy::int_plus_one,
    clippy::needless_lifetimes,
    clippy::nonminimal_bool,
    clippy::unnecessary_mut_passed,
    clippy::useless_asref,
    clippy::obfuscated_if_else,
    clippy::too_many_arguments,
    clippy::trim_split_whitespace,
    clippy::type_complexity,
    clippy::unnecessary_filter_map,
    clippy::unnecessary_lazy_evaluations,
    clippy::unnecessary_sort_by
)]

//! The collaborative editing schema used by every DOCX editing slice.
//!
//! The load-bearing rule is that a Word *story* is one continuous [`yrs::TextRef`]. A story is
//! the body flow, one header/footer part, one table cell, one footnote, and so on. [`StoryId`] is
//! deliberately opaque so later slices can use package relationship IDs or structural cell IDs
//! without changing this crate.
//!
//! OOXML maps to yrs as follows:
//!
//! - a story's ordered `w:p` stream -> one Y.Text stored under its ID in the `stories` Y.Map;
//! - each `w:p` boundary -> one countable Y.Text embed whose nested Y.Map carries `paraId`,
//!   `pStyle`, `alignment`, and future `pPrIns` / `pPrDel` / `pPrChange` values;
//! - adjacent `w:r` properties -> Y.Text formatting attributes (`bold`, `italic`, `fontFamily`,
//!   `fontSize`, `color`, plus opaque attributes added by later slices);
//! - `w:ins` / `w:del` -> `ins` / `del` Y.Text formatting attributes. A suggested deletion is
//!   retained text with a `del` attribute, never a CRDT deletion;
//! - `w:commentRangeStart` / `w:commentRangeEnd` -> encoded [`StickyIndex`] pairs in the side
//!   `comments` Y.Map. Starts use [`Assoc::After`], ends use [`Assoc::Before`].
//!
//! Internal IDs are `{clientID}:{counter}`. Dense integer `w:id` values are an export concern and
//! must be minted only while serializing OOXML.

use std::collections::{BTreeMap, HashMap};
use std::fmt;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Assoc, ClientID, Doc, IndexedSequence, Map, MapPrelim, MapRef, OffsetKind, Options, Out,
    ReadTxn, StateVector, StickyIndex, Text, TextPrelim, TextRef, Transact, Update,
};

mod ctx;
mod format;
mod op;
mod ops;
mod queries;
mod raw;
mod read_state;
mod undo;

pub mod canonical;
pub mod engine;
pub mod frame_delta;

pub use canonical::{CanonicalItem, checksum, project_story, story_checksum, to_canonical_bytes};
pub use ctx::{EditCtx, EditOrigin, SuggestCtx};
pub use engine::{EngineSession, EngineStats};
pub use format::{
    ColorPatch, FontFamilyPatch, FormatPolicy, HYPERLINK, InlineFormatDelta, Patch, SimpleFormat,
    StrikePatch, UnderlinePatch, highlight_color_name,
};
pub use op::{Loc, LocRange, OpError, OpResult, Receipt, SplitReceipt};
pub use ops::paragraph::{
    INDENT_STEP_TWIPS, MergeDirection, ParaAttrDelta, ParaSelector, ResolvedStyleProjection,
    STYLE_CONTROLLED_MARKS, STYLE_CONTROLLED_PARA_ATTRS, TabStop,
};
pub use ops::resolve::ChangeTarget;
pub use ops::table::{CellLoc, TableLocator, TableRange, TableReceipt};
pub use ops::text::RichRun;
pub use queries::{
    ChangeInfo, ChangeKind, CommentInfo, FindMatch, FindOptions, LayoutBridge, NavDirection,
    NavUnit, PageContent, PageParagraph, SelectionInfo, TextView,
};
pub use raw::RawOp;
pub use read_state::{RevisionInfo, SelectionContextInfo, TriState};
pub use undo::{DocUndoManager, UNDO_CAPTURE_TIMEOUT_MS, UNDO_DEPTH};

// The JS boundary for this crate (a wasm-bindgen session API over
// [`EditingDoc`]), compiled only for wasm-pack builds (`--features wasm`).
// COORDINATION: keep this cfg-mod HERE, next to the imports — the
// render-bridge track appends `pub mod bridge;` as the FINAL line of this
// file, and the two declarations must never collide in a merge.
#[cfg(feature = "wasm")]
pub mod wasm;

const STORIES: &str = "stories";
const COMMENTS: &str = "comments";
const PILCROW_KIND: &str = "pilcrow";
const KIND_KEY: &str = "_kind";
const PARA_ID: &str = "paraId";
const INS: &str = "ins";
const DEL: &str = "del";
/// Paragraph-mark insertion revision (suggested split), stored on the pilcrow map.
const PPR_INS: &str = "pPrIns";
/// Paragraph-mark deletion revision (suggested merge/delete), stored on the pilcrow map.
const PPR_DEL: &str = "pPrDel";
/// Paragraph-property revision, stored as OOXML-compatible change records on the pilcrow map.
const PPR_CHANGE: &str = "pPrChange";
/// `_kind` of a hard-break embed.
const BREAK_KIND: &str = "break";

pub type EditResult<T> = Result<T, EditError>;
pub type StoryId = String;
pub type ParagraphId = String;
pub type CommentId = String;
pub type RevisionId = String;

/// Cooperative durable authorship metadata for one suggested operation.
///
/// `date` is supplied by the host so native, WASM, and agent peers share one clock policy. yrs
/// transaction origins remain a separate concern used for local-only undo and authority policy.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Author {
    pub name: String,
    pub date: String,
}

impl Author {
    pub fn new(name: impl Into<String>, date: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            date: date.into(),
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Position {
    pub story: StoryId,
    /// UTF-16 units, with every embed (including a pilcrow) counting as one unit.
    pub index: u32,
}

impl Position {
    pub fn new(story: impl Into<StoryId>, index: u32) -> Self {
        Self {
            story: story.into(),
            index,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct StoryRange {
    pub story: StoryId,
    pub start: u32,
    pub end: u32,
}

impl StoryRange {
    pub fn new(story: impl Into<StoryId>, start: u32, end: u32) -> Self {
        Self {
            story: story.into(),
            start,
            end,
        }
    }

    pub(crate) fn len(&self) -> EditResult<u32> {
        self.end
            .checked_sub(self.start)
            .ok_or(EditError::InvalidRange {
                start: self.start,
                end: self.end,
            })
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct ParagraphProperties {
    pub para_id: ParagraphId,
    pub values: BTreeMap<String, Any>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum SegmentContent {
    Text(String),
    Pilcrow(ParagraphProperties),
    /// A non-pilcrow map embed. The public read surface exposes the discriminator
    /// and payload so save can reconstruct structural tables and inline atoms.
    OtherEmbed {
        kind: String,
        payload: BTreeMap<String, Any>,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub struct StorySegment {
    pub content: SegmentContent,
    pub attributes: BTreeMap<String, Any>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ParagraphSnapshot {
    pub para_id: ParagraphId,
    pub text: String,
    pub properties: BTreeMap<String, Any>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CommentAnchor {
    pub story: StoryId,
    pub start: StickyIndex,
    pub end: StickyIndex,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ResolvedCommentAnchor {
    pub story: StoryId,
    pub start: u32,
    pub end: u32,
}

#[derive(Debug, Clone, Eq, PartialEq)]
pub enum EditError {
    StoryExists(String),
    StoryNotFound(String),
    CommentNotFound(String),
    ParagraphNotFound(String),
    InvalidRange { start: u32, end: u32 },
    OutOfBounds { index: u32, len: u32 },
    ExpectedPilcrow { story: String, index: u32 },
    CannotMergeFinalParagraph { story: String, index: u32 },
    InvalidComment(String),
    InvalidUpdate(String),
    ReservedParagraphKey(String),
}

impl fmt::Display for EditError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::StoryExists(id) => write!(f, "story {id:?} already exists"),
            Self::StoryNotFound(id) => write!(f, "story {id:?} was not found"),
            Self::CommentNotFound(id) => write!(f, "comment {id:?} was not found"),
            Self::ParagraphNotFound(id) => write!(f, "paragraph {id:?} was not found"),
            Self::InvalidRange { start, end } => write!(f, "invalid range {start}..{end}"),
            Self::OutOfBounds { index, len } => {
                write!(f, "index {index} is outside the story length {len}")
            }
            Self::ExpectedPilcrow { story, index } => {
                write!(f, "expected a pilcrow embed at {story}:{index}")
            }
            Self::CannotMergeFinalParagraph { story, index } => {
                write!(f, "cannot merge the final paragraph at {story}:{index}")
            }
            Self::InvalidComment(message) => write!(f, "invalid comment: {message}"),
            Self::InvalidUpdate(message) => write!(f, "invalid yrs update: {message}"),
            Self::ReservedParagraphKey(key) => {
                write!(f, "paragraph property {key:?} is managed by the schema")
            }
        }
    }
}

impl std::error::Error for EditError {}

/// A single yrs replica of the DOCX editing model.
pub struct EditingDoc {
    doc: Doc,
    client_id: u64,
    id_counter: AtomicU64,
}

impl EditingDoc {
    /// Creates a browser-compatible replica. All public positions use UTF-16 offsets.
    pub fn new(client_id: u64) -> Self {
        let mut options = Options::with_client_id(ClientID::new(client_id));
        options.offset_kind = OffsetKind::Utf16;
        let doc = Doc::with_options(options);
        // Root shared types are schema declarations; their contents are still changed only in the
        // explicit transactions below.
        doc.get_or_insert_map(STORIES);
        doc.get_or_insert_map(COMMENTS);
        Self {
            doc,
            client_id,
            id_counter: AtomicU64::new(0),
        }
    }

    pub fn client_id(&self) -> u64 {
        self.client_id
    }

    /// Low-level access for the transport, awareness, and undo bridges.
    pub fn yrs_doc(&self) -> &Doc {
        &self.doc
    }

    /// Adds an arbitrary story with one paragraph and its final pilcrow.
    ///
    /// The text insertion carries explicit `ins:null,del:null` attributes; it is never a bare yrs
    /// insertion. The returned ID belongs to the final pilcrow.
    pub fn create_story(
        &self,
        story_id: impl Into<StoryId>,
        initial_text: &str,
        p_style: &str,
        alignment: &str,
    ) -> EditResult<ParagraphId> {
        let story_id = story_id.into();
        let para_id = self.next_id();
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        if stories.contains_key(&txn, &story_id) {
            return Err(EditError::StoryExists(story_id));
        }
        let story = stories.insert(&mut txn, story_id, TextPrelim::new(""));
        if !initial_text.is_empty() {
            story.insert_with_attributes(&mut txn, 0, initial_text, insertion_attrs(None, None));
        }
        let at = story.len(&txn);
        let pilcrow = story.insert_embed_with_attributes(
            &mut txn,
            at,
            MapPrelim::default(),
            insertion_attrs(None, None),
        );
        write_pilcrow_properties(&pilcrow, &mut txn, &para_id, p_style, alignment);
        Ok(para_id)
    }

    /// Removes one complete story from the document's story map.
    ///
    /// The coexistence table mirror uses this when a PM table/cell disappears;
    /// leaving an unreachable cell story behind would make the per-story
    /// canonical watchdog correctly report a divergent document.
    pub fn delete_story(&self, story_id: &str) -> EditResult<()> {
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        if stories.remove(&mut txn, story_id).is_some() {
            Ok(())
        } else {
            Err(EditError::StoryNotFound(story_id.to_owned()))
        }
    }

    /// Updates one independently-convergent property on the pilcrow identified by `para_id`.
    ///
    /// Arbitrary values leave room for `pPrIns`, `pPrDel`, `pPrChange`, and passive OOXML property
    /// bags. `paraId` and the embed discriminator are immutable schema identity.
    pub fn set_paragraph_attr(
        &self,
        para_id: &str,
        key: impl Into<String>,
        value: Any,
    ) -> EditResult<()> {
        let key = key.into();
        if key == PARA_ID || key == KIND_KEY {
            return Err(EditError::ReservedParagraphKey(key));
        }
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        for (_, value_ref) in stories.iter(&txn) {
            let Out::YText(story) = value_ref else {
                continue;
            };
            for (_, pilcrow) in pilcrows(&story, &txn) {
                if map_string(&pilcrow, &txn, PARA_ID).as_deref() == Some(para_id) {
                    pilcrow.insert(&mut txn, key, value);
                    return Ok(());
                }
            }
        }
        Err(EditError::ParagraphNotFound(para_id.to_owned()))
    }

    /// Creates a side-map comment whose anchors are sticky positions, in one transaction.
    pub fn add_comment(
        &self,
        ranges: &[StoryRange],
        author: &str,
        date: &str,
        body: Any,
    ) -> EditResult<CommentId> {
        if ranges.is_empty() {
            return Err(EditError::InvalidComment(
                "at least one anchored range is required".into(),
            ));
        }
        let comment_id = self.next_id();
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let mut anchors = Vec::with_capacity(ranges.len());
        for range in ranges {
            let len = range.len()?;
            let story = story_ref(&txn, &range.story)?;
            check_range(&story, &txn, range.start, len)?;
            let start = story
                .sticky_index(&txn, range.start, Assoc::After)
                .ok_or_else(|| {
                    EditError::InvalidComment("start anchor could not be made".into())
                })?;
            let end = story
                .sticky_index(&txn, range.end, Assoc::Before)
                .ok_or_else(|| EditError::InvalidComment("end anchor could not be made".into()))?;
            anchors.push(anchor_value(&range.story, &start, &end));
        }
        let comments = txn
            .get_map(COMMENTS)
            .expect("comments root is declared by EditingDoc::new");
        let comment = comments.insert(&mut txn, comment_id.as_str(), MapPrelim::default());
        comment.insert(&mut txn, "author", author);
        comment.insert(&mut txn, "date", date);
        comment.insert(&mut txn, "parentId", Any::Null);
        comment.insert(&mut txn, "done", false);
        comment.insert(&mut txn, "body", body);
        comment.insert(&mut txn, "anchors", Any::Array(Arc::from(anchors)));
        Ok(comment_id)
    }

    pub fn comment_anchors(&self, comment_id: &str) -> EditResult<Vec<CommentAnchor>> {
        let txn = self.doc.transact();
        let comments = txn
            .get_map(COMMENTS)
            .expect("comments root is declared by EditingDoc::new");
        let comment = comments
            .get(&txn, comment_id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .ok_or_else(|| EditError::CommentNotFound(comment_id.to_owned()))?;
        let anchors = match comment.get(&txn, "anchors") {
            Some(Out::Any(Any::Array(values))) => values,
            _ => {
                return Err(EditError::InvalidComment("anchors must be an array".into()));
            }
        };
        anchors.iter().map(decode_anchor).collect()
    }

    /// Resolves comment anchors for the current repaint.
    ///
    /// yrs follows an item's `redone` chain while undo/redo replaces deleted items, which is why
    /// the required undo test recovers the range. This is not an unlimited durability promise:
    /// `get_offset` returns `None` if the referenced tombstone has been garbage-collected, or if an
    /// importer rebuilds content with unrelated CRDT identities. UndoManager keeps its scoped
    /// deleted items from GC while they remain on an undo/redo stack.
    pub fn resolve_comment(&self, comment_id: &str) -> EditResult<Vec<ResolvedCommentAnchor>> {
        let anchors = self.comment_anchors(comment_id)?;
        let txn = self.doc.transact();
        anchors
            .into_iter()
            .map(|anchor| {
                let start = anchor.start.get_offset(&txn).ok_or_else(|| {
                    EditError::InvalidComment("start anchor no longer resolves".into())
                })?;
                let end = anchor.end.get_offset(&txn).ok_or_else(|| {
                    EditError::InvalidComment("end anchor no longer resolves".into())
                })?;
                Ok(ResolvedCommentAnchor {
                    story: anchor.story,
                    start: start.index,
                    end: end.index,
                })
            })
            .collect()
    }

    /// Builds a local-origin undo manager scoped to one story.
    ///
    /// Routed through the WASM-safe constructor in [`crate::undo`] — `UndoManager::new` and
    /// `Options::default()` do not exist on `wasm32-unknown-unknown`.
    pub fn undo_manager(&self, story_id: &str) -> EditResult<yrs::undo::UndoManager<()>> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        drop(txn);
        Ok(undo::build_manager(self, &[story], undo::default_clock()))
    }

    pub fn story_len(&self, story_id: &str) -> EditResult<u32> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        Ok(story.len(&txn))
    }

    pub fn story_segments(&self, story_id: &str) -> EditResult<Vec<StorySegment>> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        Ok(story
            .diff(&txn, YChange::identity)
            .into_iter()
            .map(|diff| StorySegment {
                content: segment_content(diff.insert, &txn),
                attributes: ordered_attrs(diff.attributes.as_deref()),
            })
            .collect())
    }

    pub fn paragraphs(&self, story_id: &str) -> EditResult<Vec<ParagraphSnapshot>> {
        let mut paragraphs = Vec::new();
        let mut text = String::new();
        for segment in self.story_segments(story_id)? {
            match segment.content {
                SegmentContent::Text(value) => text.push_str(&value),
                SegmentContent::Pilcrow(properties) => {
                    paragraphs.push(ParagraphSnapshot {
                        para_id: properties.para_id,
                        text: std::mem::take(&mut text),
                        properties: properties.values,
                    });
                }
                SegmentContent::OtherEmbed { .. } => {}
            }
        }
        Ok(paragraphs)
    }

    pub fn paragraph_mark_position(&self, para_id: &str) -> EditResult<Position> {
        let txn = self.doc.transact();
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        for (story_id, value) in stories.iter(&txn) {
            let Out::YText(story) = value else {
                continue;
            };
            for (index, pilcrow) in pilcrows(&story, &txn) {
                if map_string(&pilcrow, &txn, PARA_ID).as_deref() == Some(para_id) {
                    return Ok(Position::new(story_id.to_string(), index));
                }
            }
        }
        Err(EditError::ParagraphNotFound(para_id.to_owned()))
    }

    pub fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default())
    }

    pub fn apply_update_v1(&self, bytes: &[u8]) -> EditResult<()> {
        let update = Update::decode_v1(bytes)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.doc
            .transact_mut()
            .apply_update(update)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))
    }

    /// Applies a v1 update using this replica's local transaction origin.
    ///
    /// This is reserved for a local worker replica executing an edit on behalf
    /// of this document. Ordinary collaboration updates must continue through
    /// [`Self::apply_update_v1`] so local undo never captures remote work.
    pub fn apply_local_update_v1(&self, bytes: &[u8]) -> EditResult<()> {
        let update = Update::decode_v1(bytes)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.doc
            .transact_mut_with(self.client_id)
            .apply_update(update)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))
    }

    fn next_id(&self) -> String {
        let counter = self.id_counter.fetch_add(1, Ordering::Relaxed);
        format!("{}:{counter}", self.client_id)
    }
}

fn story_ref<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<TextRef> {
    txn.get_map(STORIES)
        .and_then(|stories| stories.get(txn, story_id))
        .and_then(|value| value.cast::<TextRef>().ok())
        .ok_or_else(|| EditError::StoryNotFound(story_id.to_owned()))
}

fn check_position<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> EditResult<()> {
    let len = story.len(txn);
    if index <= len {
        Ok(())
    } else {
        Err(EditError::OutOfBounds { index, len })
    }
}

fn check_range<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, len: u32) -> EditResult<()> {
    let story_len = story.len(txn);
    if start.checked_add(len).is_some_and(|end| end <= story_len) {
        Ok(())
    } else {
        Err(EditError::OutOfBounds {
            index: start.saturating_add(len),
            len: story_len,
        })
    }
}

fn insertion_attrs(ins: Option<Any>, del: Option<Any>) -> Attrs {
    Attrs::from([
        (Arc::from(INS), ins.unwrap_or(Any::Null)),
        (Arc::from(DEL), del.unwrap_or(Any::Null)),
    ])
}

fn revision_value(id: &str, author: &Author) -> Any {
    Any::Map(Arc::new(HashMap::from([
        ("id".into(), Any::from(id)),
        ("author".into(), Any::from(author.name.as_str())),
        ("date".into(), Any::from(author.date.as_str())),
    ])))
}

fn write_pilcrow_properties(
    pilcrow: &MapRef,
    txn: &mut yrs::TransactionMut<'_>,
    para_id: &str,
    p_style: &str,
    alignment: &str,
) {
    pilcrow.insert(txn, KIND_KEY, PILCROW_KIND);
    pilcrow.insert(txn, PARA_ID, para_id);
    pilcrow.insert(txn, "pStyle", p_style);
    pilcrow.insert(txn, "alignment", alignment);
}

fn map_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<String> {
    match map.get(txn, key) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        _ => None,
    }
}

fn is_pilcrow<T: ReadTxn>(map: &MapRef, txn: &T) -> bool {
    map_string(map, txn, KIND_KEY).as_deref() == Some(PILCROW_KIND)
}

fn pilcrows<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<(u32, MapRef)> {
    let mut offset = 0;
    let mut result = Vec::new();
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert
            && is_pilcrow(&map, txn)
        {
            result.push((offset, map));
        }
        offset += len;
    }
    result
}

fn next_pilcrow<T: ReadTxn>(story: &TextRef, txn: &T, from: u32) -> Option<(u32, MapRef)> {
    pilcrows(story, txn)
        .into_iter()
        .find(|(offset, _)| *offset >= from)
}

fn out_len(value: &Out) -> u32 {
    match value {
        Out::Any(Any::String(value)) => value.encode_utf16().count() as u32,
        _ => 1,
    }
}

fn ordered_attrs(attrs: Option<&Attrs>) -> BTreeMap<String, Any> {
    attrs
        .into_iter()
        .flat_map(|attrs| attrs.iter())
        .map(|(key, value)| (key.to_string(), value.clone()))
        .collect()
}

fn segment_content<T: ReadTxn>(value: Out, txn: &T) -> SegmentContent {
    match value {
        Out::Any(Any::String(value)) => SegmentContent::Text(value.to_string()),
        Out::YMap(map) if is_pilcrow(&map, txn) => {
            let para_id = map_string(&map, txn, PARA_ID).unwrap_or_default();
            let values = map
                .iter(txn)
                .filter_map(|(key, value)| {
                    if matches!(key, KIND_KEY | PARA_ID) {
                        return None;
                    }
                    let Out::Any(value) = value else {
                        return None;
                    };
                    Some((key.to_string(), value))
                })
                .collect();
            SegmentContent::Pilcrow(ParagraphProperties { para_id, values })
        }
        Out::YMap(map) => {
            let kind = map_string(&map, txn, KIND_KEY).unwrap_or_default();
            let payload = map
                .iter(txn)
                .filter_map(|(key, value)| {
                    if key == KIND_KEY {
                        return None;
                    }
                    let Out::Any(value) = value else {
                        return None;
                    };
                    Some((key.to_string(), value))
                })
                .collect();
            SegmentContent::OtherEmbed { kind, payload }
        }
        _ => SegmentContent::OtherEmbed {
            kind: String::new(),
            payload: BTreeMap::new(),
        },
    }
}

fn anchor_value(story: &str, start: &StickyIndex, end: &StickyIndex) -> Any {
    Any::Map(Arc::new(HashMap::from([
        ("story".into(), Any::from(story)),
        ("start".into(), Any::from(start.encode_v1())),
        ("end".into(), Any::from(end.encode_v1())),
    ])))
}

fn decode_anchor(value: &Any) -> EditResult<CommentAnchor> {
    let Any::Map(value) = value else {
        return Err(EditError::InvalidComment("anchor must be a map".into()));
    };
    let story = match value.get("story") {
        Some(Any::String(story)) => story.to_string(),
        _ => return Err(EditError::InvalidComment("anchor story is missing".into())),
    };
    let decode_sticky = |key: &str| -> EditResult<StickyIndex> {
        let Some(Any::Buffer(bytes)) = value.get(key) else {
            return Err(EditError::InvalidComment(format!(
                "anchor {key} is missing"
            )));
        };
        StickyIndex::decode_v1(bytes).map_err(|error| EditError::InvalidComment(error.to_string()))
    };
    Ok(CommentAnchor {
        story,
        start: decode_sticky("start")?,
        end: decode_sticky("end")?,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const DATE: &str = "2026-07-13T12:00:00Z";

    fn local(author: &str) -> EditCtx {
        EditCtx::local(author, DATE)
    }

    fn suggesting(author: &str) -> EditCtx {
        EditCtx::local(author, DATE).suggesting()
    }

    fn seed(text: &str) -> EditingDoc {
        let doc = EditingDoc::new(100);
        doc.create_story("body", text, "Normal", "left").unwrap();
        // The same schema primitive backs a second story; body ops must never cross this boundary.
        doc.create_story("header:rId7", "Header", "Header", "center")
            .unwrap();
        doc
    }

    fn peers(text: &str, a_id: u64, b_id: u64) -> (EditingDoc, EditingDoc) {
        let baseline = seed(text);
        let update = baseline.encode_state_as_update_v1();
        let a = EditingDoc::new(a_id);
        let b = EditingDoc::new(b_id);
        a.apply_update_v1(&update).unwrap();
        b.apply_update_v1(&update).unwrap();
        (a, b)
    }

    fn sync(a: &EditingDoc, b: &EditingDoc) {
        let from_a = a.encode_state_as_update_v1();
        let from_b = b.encode_state_as_update_v1();
        a.apply_update_v1(&from_b).unwrap();
        b.apply_update_v1(&from_a).unwrap();
    }

    fn resolved(doc: &EditingDoc, comment_id: &str) -> ResolvedCommentAnchor {
        doc.resolve_comment(comment_id).unwrap().remove(0)
    }

    fn marker_attributes(doc: &EditingDoc, marker: &str) -> BTreeMap<String, Any> {
        doc.story_segments("body")
            .unwrap()
            .into_iter()
            .find_map(|segment| match segment.content {
                SegmentContent::Text(value) if value.contains(marker) => Some(segment.attributes),
                _ => None,
            })
            .unwrap_or_else(|| panic!("marker {marker:?} was not found"))
    }

    fn revision_author(attributes: &BTreeMap<String, Any>, key: &str) -> Option<String> {
        let Any::Map(revision) = attributes.get(key)? else {
            return None;
        };
        let Any::String(author) = revision.get("author")? else {
            return None;
        };
        Some(author.to_string())
    }

    #[test]
    fn local_worker_update_remains_owned_by_main_undo() {
        let main = seed("before");
        let worker = EditingDoc::new(200);
        worker
            .apply_update_v1(&main.encode_state_as_update_v1())
            .unwrap();
        let mut undo = main.undo_scope(&["body"]).unwrap();

        worker
            .insert_text(
                &local("worker"),
                Position::new("body", 6),
                " after",
                FormatPolicy::Plain,
            )
            .unwrap();
        main.apply_local_update_v1(&worker.encode_state_as_update_v1())
            .unwrap();

        assert_eq!(main.paragraphs("body").unwrap()[0].text, "before after");
        assert!(undo.undo());
        assert_eq!(main.paragraphs("body").unwrap()[0].text, "before");
    }

    #[test]
    fn assoc_orientation_lock() {
        let (a, b) = peers("ab", 1, 2);
        let story = story_ref(&a.doc.transact(), "body").unwrap();
        let after = story
            .sticky_index(&a.doc.transact(), 1, Assoc::After)
            .unwrap();
        let before = story
            .sticky_index(&a.doc.transact(), 1, Assoc::Before)
            .unwrap();
        b.insert_text(
            &local("B"),
            Position::new("body", 1),
            "X",
            FormatPolicy::Plain,
        )
        .unwrap();
        a.apply_update_v1(&b.encode_state_as_update_v1()).unwrap();
        // Actual behavior: After follows the concurrent insertion; Before stays before it.
        let txn = a.doc.transact();
        assert_eq!(
            (
                after.get_offset(&txn).unwrap().index,
                before.get_offset(&txn).unwrap().index
            ),
            (2, 1)
        );
    }

    #[test]
    fn split_merge_are_clean_sequence_ops_under_concurrency() {
        let (a, b) = peers("left suffix", 1, 2);
        let split = a
            .split_paragraph(&local("A"), Position::new("body", 5), None)
            .unwrap();
        b.insert_text(
            &local("B"),
            Position::new("body", 8),
            "REMOTE ",
            FormatPolicy::Plain,
        )
        .unwrap();
        sync(&a, &b);

        let paragraphs = a.paragraphs("body").unwrap();
        assert_eq!(paragraphs[0].text, "left ");
        assert_eq!(paragraphs[1].text, "sufREMOTE fix");
        assert_eq!(paragraphs, b.paragraphs("body").unwrap());
        assert_eq!(a.paragraphs("header:rId7").unwrap()[0].text, "Header");

        // The contract split gives the FIRST half the original paraId; its mark is the new
        // pilcrow at the split point.
        assert_eq!(
            a.paragraph_mark_position(&split.first_para_id).unwrap(),
            Position::new("body", 5)
        );
        a.merge_paragraphs(&local("A"), &split.first_para_id, MergeDirection::Forward)
            .unwrap();
        sync(&a, &b);
        assert_eq!(a.paragraphs("body").unwrap()[0].text, "left sufREMOTE fix");

        let merged_para_id = a.paragraphs("body").unwrap()[0].para_id.clone();
        a.toggle_format(
            &local("A"),
            StoryRange::new("body", 0, 4),
            SimpleFormat::Bold,
        )
        .unwrap();
        a.set_paragraph_attr(&merged_para_id, "alignment", Any::from("right"))
            .unwrap();
        sync(&a, &b);
        assert_eq!(
            marker_attributes(&a, "left").get("bold"),
            Some(&Any::Bool(true))
        );
        assert_eq!(
            a.paragraphs("body").unwrap()[0].properties.get("alignment"),
            Some(&Any::from("right"))
        );
        assert_eq!(a.paragraphs("body").unwrap(), b.paragraphs("body").unwrap());
        assert_eq!(
            a.doc.transact().state_vector(),
            b.doc.transact().state_vector(),
            "replicas contain identical CRDT history"
        );
    }

    #[test]
    fn comment_anchor_survives_insert_split_and_delete_undo_redo() {
        let baseline = seed("alpha beta gamma omega");
        let comment_id = baseline
            .add_comment(
                &[StoryRange::new("body", 6, 16)],
                "Reviewer",
                DATE,
                Any::from("comment body"),
            )
            .unwrap();
        let update = baseline.encode_state_as_update_v1();
        let a = EditingDoc::new(1);
        let b = EditingDoc::new(2);
        a.apply_update_v1(&update).unwrap();
        b.apply_update_v1(&update).unwrap();

        b.insert_text(
            &local("B"),
            Position::new("body", 11),
            "REMOTE ",
            FormatPolicy::Plain,
        )
        .unwrap();
        sync(&a, &b);
        assert_eq!(
            (
                resolved(&a, &comment_id).start,
                resolved(&a, &comment_id).end
            ),
            (6, 23)
        );

        a.split_paragraph(&local("A"), Position::new("body", 11), None)
            .unwrap();
        let before_delete = resolved(&a, &comment_id);
        assert_eq!((before_delete.start, before_delete.end), (6, 24));

        let mut undo = a.undo_manager("body").unwrap();
        // Delete strictly inside the annotation, leaving both boundary identities alive. yrs can
        // follow ordinary redone chains, but does not promise the exact original side when the
        // boundary item itself is deleted and recreated (and cannot recover it after GC).
        a.delete_range(&local("A"), StoryRange::new("body", 13, 19))
            .unwrap();
        let after_delete = resolved(&a, &comment_id);
        assert_eq!((after_delete.start, after_delete.end), (6, 18));

        assert!(undo.undo_blocking());
        let after_undo = resolved(&a, &comment_id);
        assert_eq!((after_undo.start, after_undo.end), (6, 24));
        assert!(undo.redo_blocking());
        let after_redo = resolved(&a, &comment_id);
        assert_eq!((after_redo.start, after_redo.end), (6, 18));
    }

    #[test]
    fn case_e_plain_insert_inherits_delete_but_suggested_insert_is_legal() {
        let (plain_a, plain_b) = peers("abcdef", 1, 2);
        plain_a
            .delete_range(&suggesting("Alice"), StoryRange::new("body", 2, 4))
            .unwrap();
        plain_b
            .insert_text(
                &local("Bob"),
                Position::new("body", 3),
                "PLAIN",
                FormatPolicy::Plain,
            )
            .unwrap();
        sync(&plain_a, &plain_b);
        let wrong = marker_attributes(&plain_a, "PLAIN");
        // Wrong but unavoidable case E: Alice's concurrent range format captures Bob's plain text.
        assert_eq!(revision_author(&wrong, DEL).as_deref(), Some("Alice"));
        assert_eq!(revision_author(&wrong, INS), None);

        let (suggest_a, suggest_b) = peers("abcdef", 11, 12);
        suggest_a
            .delete_range(&suggesting("Alice"), StoryRange::new("body", 2, 4))
            .unwrap();
        suggest_b
            .insert_text(
                &suggesting("Bob"),
                Position::new("body", 3),
                "SUGGEST",
                FormatPolicy::Plain,
            )
            .unwrap();
        sync(&suggest_a, &suggest_b);
        let legal = marker_attributes(&suggest_a, "SUGGEST");
        // Suggesting mode makes the same race legal and reviewable as nested w:ins > w:del.
        assert_eq!(revision_author(&legal, DEL).as_deref(), Some("Alice"));
        assert_eq!(revision_author(&legal, INS).as_deref(), Some("Bob"));
    }
}

pub mod bridge;
