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
//! The load-bearing rule is that a Word story is one continuous [`yrs::TextRef`].
//!
//! OOXML maps to yrs as follows:
//!
//! - a story's ordered `w:p` stream -> one Y.Text stored under its ID in the `stories` Y.Map;
//! - each `w:p` boundary -> one countable Y.Text embed whose nested Y.Map carries `paraId`,
//!   `pStyle`, `alignment`, and paragraph-change values;
//! - adjacent `w:r` properties -> Y.Text formatting attributes (`bold`, `italic`, `fontFamily`,
//!   `fontSize`, `color`, plus opaque attributes);
//! - `w:ins` / `w:del` -> `ins` / `del` Y.Text formatting attributes. A suggested deletion is
//!   retained text with a `del` attribute, never a CRDT deletion;
//! - `w:commentRangeStart` / `w:commentRangeEnd` -> encoded [`StickyIndex`] pairs in the side
//!   `comments` Y.Map. Starts use [`Assoc::After`], ends use [`Assoc::Before`].
//!
//! Internal IDs are `{clientID}:{counter}`. Dense integer `w:id` values are an export concern and
//! must be minted only while serializing OOXML. A paragraph's Word `w14:paraId` is a separate
//! binding on its pilcrow, never derived from its internal ID; see [`ParagraphIdentity`].

#[cfg(test)]
extern crate self as docx_edit;

#[cfg(feature = "wasm")]
use std::collections::HashSet;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::fmt;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use yrs::types::text::YChange;
use yrs::types::{Attrs, Delta};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Assoc, ClientID, DeepObservable, Doc, In, IndexedSequence, Map, MapPrelim, MapRef,
    OffsetKind, Options, Out, ReadTxn, StateVector, StickyIndex, Subscription, Text, TextPrelim,
    TextRef, Transact, Update,
};

mod batch;
mod comment_references;
#[cfg_attr(not(feature = "wasm"), allow(dead_code))]
mod compare;
pub mod content_controls;
mod control_source;
mod control_values;
mod ctx;
mod deterministic;
mod fingerprint;
mod format;
mod geometry;
mod heading;
mod identity;
mod inline_content;
mod list_marker;
pub mod media;
mod op;
mod ops;
mod policy;
mod presence;
mod queries;
mod raw;
mod read_state;
pub mod read_types;
mod script_fonts;
mod search;
mod seed;
mod segments;
pub mod structured;
mod target;
mod undo;

pub mod canonical;
pub mod engine;
pub mod frame_delta;

pub use batch::{
    DocumentVersion, EditApplication, EditFailure, EditFailureCode, EditFailureReason, EditGuard,
    EditHistory, EditOperation, EditPreview, EditReceipt, EditRefusal, EditRequest, EditSource,
    EditStep, EditSuggestion, EditTarget, EditValidation, ParagraphInput, ResolvedControl,
    TargetEdge,
};
pub use canonical::{CanonicalItem, checksum, project_story, story_checksum, to_canonical_bytes};
pub use ctx::{EditCtx, EditOrigin, SuggestCtx};
pub use engine::{EngineSession, EngineStats, RegionLayoutProgress};
pub use format::{
    ColorPatch, FontFamilyPatch, FormatPolicy, HYPERLINK, InlineFormatDelta, Patch, SimpleFormat,
    StrikePatch, UnderlinePatch, highlight_color_name,
};
pub use geometry::{
    GeometryEndpoint, GeometryParagraphOutline, GeometryPositionOutline, GeometryRange,
    GeometryRead, GeometrySentinel, GeometryStoryOutline, OwnedRevisionRange,
};
pub use identity::{
    AnchorResolution, AnchorUnsupported, ParagraphAnchor, ParagraphIdAssignment,
    ParagraphIdDiagnostic, ParagraphIdOrigin, ParagraphIdRefusal, ParagraphIdentities,
    ParagraphIdentity, ParagraphOrigin, ParagraphRef, ParagraphSavePlan, PersistedParagraphIds,
    SourceParagraphRef, SourceStory, SourceStoryKind, SplicedPart,
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
pub use search::{TextSearchError, TextSearchMatch};
pub use seed::{seed_docx_preview, seed_from_docx, seed_from_docx_with_generation};
use segments::{ParagraphIndex, SegmentIndex, build_indexes};
pub use target::{
    AtomKind, EditTextView, FindTextRequest, FindTextResponse, ParagraphTarget, ParagraphText,
    ReadParagraphsRequest, ReadParagraphsResponse, SearchScope, TextAtom, TextMatch, TextPosition,
    TextRange, TextTarget,
};
pub use undo::{DocUndoManager, UNDO_CAPTURE_TIMEOUT_MS, UNDO_DEPTH, UndoCaptureMode, UndoSession};

#[cfg(feature = "wasm")]
pub mod wasm;
#[cfg(feature = "wasm")]
pub mod wasm_memory;

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

/// `segments` split after each pilcrow into units.
fn split_segment_units(segments: Vec<StorySegment>) -> Vec<Vec<StorySegment>> {
    let mut units = Vec::new();
    let mut unit = Vec::new();
    for segment in segments {
        let closes = matches!(segment.content, SegmentContent::Pilcrow(_));
        unit.push(segment);
        if closes {
            units.push(std::mem::take(&mut unit));
        }
    }
    if !unit.is_empty() {
        units.push(unit);
    }
    units
}

/// A 128-bit digest of `segments`' content: equal digests mean equal
/// segments.
pub fn segments_digest(segments: &[StorySegment]) -> u128 {
    use std::hash::{Hash, Hasher};

    /// Two independent word-wise multiply-rotate lanes over the same bytes,
    /// for 128 bits in one walk.
    #[derive(Default)]
    struct Digest {
        low: u64,
        high: u64,
    }
    impl Digest {
        fn add(&mut self, word: u64) {
            self.low = (self.low.rotate_left(5) ^ word).wrapping_mul(0x517c_c1b7_2722_0a95);
            self.high = (self.high ^ word)
                .wrapping_mul(0x9e37_79b9_7f4a_7c15)
                .rotate_left(29);
        }
        fn finish128(&self) -> u128 {
            fn mix(mut value: u64) -> u64 {
                value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
                value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
                value ^ (value >> 31)
            }
            (u128::from(mix(self.high)) << 64) | u128::from(mix(self.low))
        }
    }
    impl Hasher for Digest {
        fn write(&mut self, bytes: &[u8]) {
            let (words, rest) = bytes.as_chunks::<8>();
            for word in words {
                self.add(u64::from_le_bytes(*word));
            }
            let mut tail = [0u8; 8];
            tail[..rest.len()].copy_from_slice(rest);
            self.add(u64::from_le_bytes(tail) ^ ((bytes.len() as u64) << 56));
        }
        fn finish(&self) -> u64 {
            self.finish128() as u64
        }
    }

    fn any(value: &Any, hasher: &mut Digest) {
        match value {
            Any::Null => 0u8.hash(hasher),
            Any::Undefined => 1u8.hash(hasher),
            Any::Bool(value) => (2u8, value).hash(hasher),
            Any::Number(value) => (3u8, value.to_bits()).hash(hasher),
            Any::BigInt(value) => (4u8, value).hash(hasher),
            Any::String(value) => (5u8, &**value).hash(hasher),
            Any::Buffer(value) => (6u8, &**value).hash(hasher),
            Any::Array(values) => {
                (7u8, values.len()).hash(hasher);
                for value in values.iter() {
                    any(value, hasher);
                }
            }
            Any::Map(entries) => {
                // Map iteration order is arbitrary, so entries combine commutatively.
                let mut combined = 0u128;
                for (key, value) in entries.iter() {
                    let mut entry = Digest::default();
                    key.hash(&mut entry);
                    any(value, &mut entry);
                    combined = combined.wrapping_add(entry.finish128());
                }
                (8u8, entries.len(), combined).hash(hasher);
            }
        }
    }
    fn map(entries: &BTreeMap<String, Any>, hasher: &mut Digest) {
        entries.len().hash(hasher);
        for (key, value) in entries {
            key.hash(hasher);
            any(value, hasher);
        }
    }

    let mut hasher = Digest::default();
    segments.len().hash(&mut hasher);
    for segment in segments {
        match &segment.content {
            SegmentContent::Text(text) => (0u8, text).hash(&mut hasher),
            SegmentContent::Pilcrow(properties) => {
                (1u8, &properties.para_id).hash(&mut hasher);
                map(&properties.values, &mut hasher);
            }
            SegmentContent::OtherEmbed { kind, payload } => {
                (2u8, kind).hash(&mut hasher);
                map(payload, &mut hasher);
            }
        }
        map(&segment.attributes, &mut hasher);
    }
    hasher.finish128()
}

#[derive(Clone, Debug, PartialEq)]
pub struct ParagraphSnapshot {
    pub para_id: ParagraphId,
    pub text: String,
    pub properties: BTreeMap<String, Any>,
}

/// One paragraph of a [`EditingDoc::seed_story`] batch.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SeedParagraph {
    pub text: String,
    pub p_style: String,
    pub alignment: String,
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
    InvalidStateVector(String),
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
            Self::InvalidStateVector(message) => write!(f, "invalid yrs state vector: {message}"),
            Self::InvalidUpdate(message) => write!(f, "invalid yrs update: {message}"),
            Self::ReservedParagraphKey(key) => {
                write!(f, "paragraph property {key:?} is managed by the schema")
            }
        }
    }
}

impl std::error::Error for EditError {}

static DOC_INSTANCES: AtomicU64 = AtomicU64::new(1);

/// Per-story values built at one committed epoch. Older values are never served, so the
/// first value of a newer epoch drops them all at once.
struct EpochCache<T> {
    epoch: u64,
    entries: HashMap<Box<str>, Arc<T>>,
}

impl<T> Default for EpochCache<T> {
    fn default() -> Self {
        Self {
            epoch: 0,
            entries: HashMap::new(),
        }
    }
}

/// The revision each story last changed at. Every committed change to the
/// stories map stamps the stories it touched (content, embedded maps, the
/// story entry itself) with the next revision.
#[derive(Default)]
struct StoryRevisions {
    current: u64,
    stamped: HashMap<Arc<str>, u64>,
}

impl StoryRevisions {
    fn stamp(&mut self, txn: &yrs::TransactionMut, events: &yrs::types::Events) {
        self.current += 1;
        for event in events.iter() {
            match event.path().front() {
                Some(yrs::types::PathSegment::Key(story)) => {
                    self.stamped.insert(Arc::clone(story), self.current);
                }
                Some(yrs::types::PathSegment::Index(_)) => {}
                None => {
                    if let yrs::types::Event::Map(entries) = event {
                        for story in entries.keys(txn).keys() {
                            self.stamped.insert(Arc::clone(story), self.current);
                        }
                    }
                }
            }
        }
    }
}

impl<T> EpochCache<T> {
    fn get(&self, story_id: &str, epoch: u64) -> Option<Arc<T>> {
        if self.epoch != epoch {
            return None;
        }
        self.entries.get(story_id).cloned()
    }

    fn take(&mut self, story_id: &str, epoch: u64) -> Option<Arc<T>> {
        if self.epoch != epoch {
            return None;
        }
        self.entries.remove(story_id)
    }

    fn insert(&mut self, story_id: &str, epoch: u64, value: Arc<T>) {
        if epoch < self.epoch {
            return;
        }
        if epoch > self.epoch {
            self.entries.clear();
            self.epoch = epoch;
        }
        self.entries.insert(story_id.into(), value);
    }
}

pub(crate) enum UpdateOrigin {
    Remote,
    Local,
    Host,
}

struct SourceNoteSeparators {
    source: Arc<seed::SourceMetadata>,
    state: Result<Option<Arc<[u8]>>, String>,
}

/// A single yrs replica of the DOCX editing model.
pub struct EditingDoc {
    doc: Doc,
    client_id: u64,
    id_counter: AtomicU64,
    direct_batches: AtomicBool,
    direct_batches_applied: AtomicU64,
    host_edit_depth: Arc<AtomicU32>,
    /// Bumped once per committed update (local ops, remote merges, undo/redo); segment
    /// indexes and chunk snapshots older than the current value are rebuilt on next lookup.
    epoch: Arc<AtomicU64>,
    /// Process-unique identity of this replica object.
    instance: u64,
    /// Rotated whenever the replica's content or retained source is replaced.
    version_nonce: AtomicU64,
    metadata: Mutex<Option<Arc<seed::SourceMetadata>>>,
    segment_indexes: Mutex<EpochCache<SegmentIndex>>,
    paragraph_indexes: Mutex<EpochCache<ParagraphIndex>>,
    chunk_snapshots: Mutex<EpochCache<Vec<ops::Chunk>>>,
    shared_read_depth: AtomicU32,
    /// Story projections held only inside a shared-read scope.
    story_views: Mutex<EpochCache<target::StoryView>>,
    source: Mutex<Option<identity::SourcePackage>>,
    media: Mutex<Option<Arc<docx_parse::media::MediaTable>>>,
    media_sources: Mutex<media::MediaSources>,
    source_note_separators: Mutex<Option<SourceNoteSeparators>>,
    loaded_note_separator_state: Mutex<Option<Arc<[u8]>>>,
    seen: identity::SeenCell,
    scan_cache: identity::ScanCache,
    story_revisions: Arc<Mutex<StoryRevisions>>,
    _update_sub: Subscription,
    _story_revision_sub: Subscription,
    _seen_subs: Vec<Subscription>,
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
        doc.get_or_insert_map(identity::PARAGRAPH_IDS);
        doc.get_or_insert_map(identity::SOURCE_PARAGRAPH_IDS);
        doc.get_or_insert_map(identity::SESSION);
        let epoch = Arc::new(AtomicU64::new(0));
        let observed = Arc::clone(&epoch);
        // after_transaction: bumps on any store-changing commit without encoding an update.
        let update_sub = doc
            .observe_after_transaction(move |txn| {
                if !txn.delete_set().is_empty() || txn.after_state() != txn.before_state() {
                    observed.fetch_add(1, Ordering::Relaxed);
                }
            })
            .expect("a fresh doc accepts an update observer");
        let story_revisions = Arc::new(Mutex::new(StoryRevisions::default()));
        let stamped = Arc::clone(&story_revisions);
        let story_revision_sub = doc
            .get_or_insert_map(STORIES)
            .observe_deep(move |txn, events| stamped.lock().unwrap().stamp(txn, events));
        let seen = identity::SeenCell::default();
        let seen_subs = identity::observe_seen(&doc, &seen);
        Self {
            doc,
            client_id,
            id_counter: AtomicU64::new(0),
            direct_batches: AtomicBool::new(false),
            direct_batches_applied: AtomicU64::new(0),
            host_edit_depth: Arc::new(AtomicU32::new(0)),
            epoch,
            instance: DOC_INSTANCES.fetch_add(1, Ordering::Relaxed),
            version_nonce: AtomicU64::new(batch::mint_nonce(client_id, 0)),
            metadata: Mutex::new(None),
            segment_indexes: Mutex::default(),
            paragraph_indexes: Mutex::default(),
            chunk_snapshots: Mutex::default(),
            shared_read_depth: AtomicU32::new(0),
            story_views: Mutex::default(),
            source: Mutex::new(None),
            media: Mutex::new(None),
            media_sources: Mutex::default(),
            source_note_separators: Mutex::default(),
            loaded_note_separator_state: Mutex::default(),
            seen,
            scan_cache: identity::ScanCache::default(),
            story_revisions,
            _update_sub: update_sub,
            _story_revision_sub: story_revision_sub,
            _seen_subs: seen_subs,
        }
    }

    #[doc(hidden)]
    pub fn set_direct_batches(&self, on: bool) {
        self.direct_batches.store(on, Ordering::Relaxed);
    }

    #[doc(hidden)]
    pub fn direct_batches_applied(&self) -> u64 {
        self.direct_batches_applied.load(Ordering::Relaxed)
    }

    /// The optimistic-concurrency token of this replica's committed state.
    ///
    /// It changes with every committed change (local, remote, undo and redo) and whenever the
    /// document or its retained source is replaced. It is scoped to this replica: equal tokens
    /// from different replicas mean nothing.
    pub fn version(&self) -> DocumentVersion {
        batch::version_token(
            self.version_nonce.load(Ordering::Relaxed),
            self.epoch.load(Ordering::Relaxed),
        )
    }

    /// Invalidates every version handed out so far. `entropy` is mixed into the new nonce.
    pub(crate) fn rotate_version(&self, entropy: u64) {
        self.version_nonce.store(
            batch::mint_nonce(self.client_id, entropy),
            Ordering::Relaxed,
        );
    }

    /// Retains the opened package's style and structure context and rotates the version.
    pub(crate) fn install_source(&self, source: seed::SourceMetadata, entropy: u64) {
        *self.metadata.lock().unwrap() = Some(Arc::new(source));
        // Story projections read the source, so none built before it is served again.
        self.epoch.fetch_add(1, Ordering::Relaxed);
        self.rotate_version(entropy);
    }

    pub(crate) fn source_metadata(&self) -> Option<Arc<seed::SourceMetadata>> {
        self.metadata.lock().unwrap().clone()
    }

    #[cfg(feature = "wasm")]
    pub(crate) fn rebase_comment_writes(&self, mut written: HashSet<(String, Option<String>)>) {
        let Some(source) = self.source_metadata() else {
            return;
        };
        let read = source.read();
        let txn = self.doc.transact();
        if let Some(comments) = txn.get_map(COMMENTS) {
            for source_comment in &read.comments {
                let Some(comment) = comments
                    .get(&txn, &source_comment.id)
                    .and_then(|value| value.cast::<MapRef>().ok())
                else {
                    continue;
                };
                for (key, placeholder) in [
                    ("author", Any::String("".into())),
                    ("date", Any::String("".into())),
                    ("parentId", Any::Null),
                    ("body", Any::Null),
                    ("done", Any::Bool(false)),
                ] {
                    let authored = match comment.get(&txn, key) {
                        Some(Out::Any(value)) => value != placeholder,
                        _ => true,
                    };
                    if authored {
                        written.insert((source_comment.id.clone(), Some(key.to_owned())));
                    }
                }
            }
        }
        read.comment_writes.replace(written);
    }

    #[doc(hidden)]
    pub fn committed_epoch(&self) -> u64 {
        self.epoch.load(Ordering::Relaxed)
    }

    /// Cached segment geometry for `story_id`, rebuilt when the doc changes.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn segment_index(&self, story_id: &str) -> EditResult<Arc<SegmentIndex>> {
        let epoch = self.committed_epoch();
        if let Some(index) = self.segment_indexes.lock().unwrap().get(story_id, epoch) {
            return Ok(index);
        }
        self.build_story_indexes(story_id)
            .map(|(segments, _)| segments)
    }

    pub(crate) fn paragraph_index(&self, story_id: &str) -> EditResult<Arc<ParagraphIndex>> {
        let epoch = self.committed_epoch();
        if let Some(index) = self.paragraph_indexes.lock().unwrap().get(story_id, epoch) {
            return Ok(index);
        }
        self.build_story_indexes(story_id)
            .map(|(_, paragraphs)| paragraphs)
    }

    fn build_story_indexes(
        &self,
        story_id: &str,
    ) -> EditResult<(Arc<SegmentIndex>, Arc<ParagraphIndex>)> {
        let txn = self.doc.transact();
        // Commits bump the epoch while they hold the store's write lock, so this read txn pins it
        // to the snapshot being indexed.
        let epoch = self.committed_epoch();
        let story = story_ref(&txn, story_id)?;
        let (segments, paragraphs) = build_indexes(&story, &txn);
        drop(txn);
        let segments = Arc::new(segments);
        let paragraphs = Arc::new(paragraphs);
        self.segment_indexes
            .lock()
            .unwrap()
            .insert(story_id, epoch, Arc::clone(&segments));
        self.paragraph_indexes
            .lock()
            .unwrap()
            .insert(story_id, epoch, Arc::clone(&paragraphs));
        Ok((segments, paragraphs))
    }

    pub(crate) fn advance_indexes_after_text_insert(
        &self,
        story_id: &str,
        before: u64,
        after: u64,
        index: u32,
        text: &str,
    ) {
        if before.checked_add(1) != Some(after) {
            return;
        }
        {
            let mut indexes = self.paragraph_indexes.lock().unwrap();
            if let Some(mut paragraphs) = indexes.take(story_id, before) {
                if Arc::make_mut(&mut paragraphs)
                    .shift_for_text_insert(index, text.encode_utf16().count() as u32)
                {
                    indexes.insert(story_id, after, paragraphs);
                }
            }
        }
        let mut indexes = self.segment_indexes.lock().unwrap();
        if let Some(mut segments) = indexes.take(story_id, before) {
            if Arc::make_mut(&mut segments).shift_for_text_insert(index, text) {
                indexes.insert(story_id, after, segments);
            }
        }
    }

    /// Advances cached story indexes after a single plain text deletion.
    pub(crate) fn advance_indexes_after_text_delete(
        &self,
        story_id: &str,
        before: u64,
        after: u64,
        start: u32,
        end: u32,
    ) {
        if before.checked_add(1) != Some(after) {
            return;
        }
        {
            let mut indexes = self.paragraph_indexes.lock().unwrap();
            if let Some(mut paragraphs) = indexes.take(story_id, before) {
                if Arc::make_mut(&mut paragraphs).shift_for_text_delete(start, end) {
                    indexes.insert(story_id, after, paragraphs);
                }
            }
        }
        let mut indexes = self.segment_indexes.lock().unwrap();
        if let Some(mut segments) = indexes.take(story_id, before) {
            if Arc::make_mut(&mut segments).shift_for_text_delete(start, end) {
                indexes.insert(story_id, after, segments);
            }
        }
    }

    /// Shared `ops::snapshot` for `story_id`, rebuilt per committed epoch.
    pub(crate) fn chunk_snapshot<T: ReadTxn>(
        &self,
        story_id: &str,
        story: &TextRef,
        txn: &T,
    ) -> Arc<Vec<ops::Chunk>> {
        let epoch = self.epoch.load(Ordering::Relaxed);
        if let Some(chunks) = self.chunk_snapshots.lock().unwrap().get(story_id, epoch) {
            return chunks;
        }
        let chunks = Arc::new(ops::snapshot(story, txn));
        self.chunk_snapshots
            .lock()
            .unwrap()
            .insert(story_id, epoch, Arc::clone(&chunks));
        chunks
    }

    pub fn client_id(&self) -> u64 {
        self.client_id
    }

    /// Until the matching `end_shared_reads`, committed reads share story projections of each document state.
    pub fn begin_shared_reads(&self) {
        self.shared_read_depth.fetch_add(1, Ordering::Relaxed);
    }

    /// Ends a shared-read scope, dropping shared story projections when the last scope ends.
    pub fn end_shared_reads(&self) {
        let mut views = self.story_views.lock().unwrap();
        #[allow(deprecated)]
        let previous =
            self.shared_read_depth
                .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |depth| {
                    depth.checked_sub(1)
                });
        if previous == Ok(1) {
            *views = EpochCache::default();
        }
    }

    #[cfg(all(test, feature = "wasm"))]
    pub(crate) fn source_indexed(&self) -> bool {
        matches!(
            *self.source.lock().unwrap(),
            Some(identity::SourcePackage::Ready(_))
        )
    }

    /// Retains the package the stories were, or will be, seeded from.
    pub(crate) fn retain_source(&self, source: identity::SourcePackage) {
        *self.media.lock().unwrap() = None;
        *self.source.lock().unwrap() = Some(source);
    }

    /// Keeps `media`, read from the retained package, as the table the
    /// stories' `media:{n}` image sources name.
    pub(crate) fn install_media(&self, media: docx_parse::media::MediaTable) {
        *self.media.lock().unwrap() = Some(Arc::new(media));
    }

    /// The fingerprints of the `data:` image sources seeding wrote in place
    /// of `media:{n}` tokens; see [`media::MediaSources`].
    pub fn media_sources(&self) -> media::MediaSources {
        self.media_sources.lock().unwrap().clone()
    }

    /// Replaces the media sources, keeping the current ones when equal so
    /// what was lowered with them stays valid.
    pub(crate) fn set_media_sources(&self, sources: media::MediaSources) {
        let mut current = self.media_sources.lock().unwrap();
        if *current != sources {
            *current = sources;
        }
    }

    /// The separator notes as a yrs v1 update for replicas without the source.
    #[doc(hidden)]
    pub fn note_separator_state(&self) -> Result<Option<Arc<[u8]>>, String> {
        let Some(source) = self.source_metadata() else {
            return Ok(self.loaded_note_separator_state.lock().unwrap().clone());
        };
        let mut cache = self.source_note_separators.lock().unwrap();
        if cache
            .as_ref()
            .is_none_or(|cached| !Arc::ptr_eq(&cached.source, &source))
        {
            let state = Self::derive_note_separator_state(&source);
            *cache = Some(SourceNoteSeparators { source, state });
        }
        cache.as_ref().unwrap().state.clone()
    }

    fn derive_note_separator_state(
        source: &seed::SourceMetadata,
    ) -> Result<Option<Arc<[u8]>>, String> {
        fn replace_marks(value: &mut serde_json::Value) {
            if matches!(
                value.get("type").and_then(serde_json::Value::as_str),
                Some("separator" | "continuationSeparator")
            ) {
                *value = serde_json::json!({"type": "text", "text": "\u{200b}"});
            } else {
                match value {
                    serde_json::Value::Array(values) => values.iter_mut().for_each(replace_marks),
                    serde_json::Value::Object(values) => {
                        values.values_mut().for_each(replace_marks)
                    }
                    _ => {}
                }
            }
        }
        let mut stories = Vec::new();
        for kind in ["footnote", "endnote"] {
            if let Some(paragraphs) = source.read().note_separator_paragraphs.get(kind)
                && !paragraphs.is_empty()
            {
                let mut paragraphs = paragraphs.clone();
                paragraphs.iter_mut().for_each(replace_marks);
                stories.push((kind.to_owned(), paragraphs));
            }
        }
        if stories.is_empty() {
            return Ok(None);
        }
        let scratch = EditingDoc::new(0);
        seed::seed_blocks(
            &scratch,
            Some(source),
            &stories
                .iter()
                .map(|(kind, paragraphs)| (kind.clone(), paragraphs.as_slice()))
                .collect::<Vec<_>>(),
        )?;
        Ok(Some(scratch.encode_state_as_update_v1().into()))
    }

    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn has_note_separator_state(&self, state: &[u8]) -> bool {
        self.loaded_note_separator_state
            .lock()
            .unwrap()
            .as_deref()
            .unwrap_or_default()
            == state
    }

    /// Replaces loaded separator notes, preserving their Arc when the bytes match.
    #[doc(hidden)]
    pub fn set_note_separator_state(&self, state: Option<Arc<[u8]>>) {
        let state = state.filter(|state| !state.is_empty());
        let mut current = self.loaded_note_separator_state.lock().unwrap();
        if current.as_deref() != state.as_deref() {
            *current = state;
        }
    }

    /// The media behind the `media:{n}` image sources of the stories seeded
    /// from the retained package, read from that package on first use.
    pub fn media_table(&self) -> Option<Arc<docx_parse::media::MediaTable>> {
        let mut media = self.media.lock().unwrap();
        if media.is_none() {
            let bytes = match self.source.lock().unwrap().as_ref()? {
                identity::SourcePackage::Pending(bytes, _) => bytes.clone(),
                identity::SourcePackage::Ready(index) => index.bytes(),
            };
            let package = ooxml_opc::RetainedPackage::from_bytes(bytes).ok()?;
            *media = Some(Arc::new(docx_parse::media::MediaTable::new(package).ok()?));
        }
        media.clone()
    }

    /// Retains the DOCX package another replica seeded this document from, so
    /// a replica hydrated from state resolves source and persisted anchors and
    /// reserves the package's paragraph IDs. Indexed on first identity use.
    pub fn retain_source_docx(&self, bytes: impl Into<Arc<[u8]>>) {
        let bytes: Arc<[u8]> = bytes.into();
        self.retain_source(identity::SourcePackage::Pending(bytes.into(), None));
    }

    /// Runs `f` over the identities this replica has seen, building them from
    /// `txn` on first use. `f` must not commit a transaction.
    pub(crate) fn with_seen<T: ReadTxn, R>(
        &self,
        txn: &T,
        f: impl FnOnce(&mut identity::Seen) -> R,
    ) -> R {
        let mut seen = self.seen.lock().unwrap();
        f(seen.get_or_insert_with(|| identity::Seen::scan(txn)))
    }

    pub(crate) fn seen_cell(&self) -> identity::SeenCell {
        Arc::clone(&self.seen)
    }

    /// Drops the seen identities after seeding wrote stories wholesale, so
    /// the next use rebuilds them from the seeded state.
    pub(crate) fn forget_seen(&self) {
        *self.seen.lock().unwrap() = None;
    }

    /// The retained package's identity index, built on first use.
    pub(crate) fn source_index(&self) -> Option<Arc<identity::SourceIndex>> {
        let mut source = self.source.lock().unwrap();
        let index = match source.as_ref()? {
            identity::SourcePackage::Ready(index) => return Some(Arc::clone(index)),
            identity::SourcePackage::Pending(bytes, digest) => {
                seed::source_index(bytes.clone(), digest.clone())
                    .ok()
                    .map(Arc::new)
            }
        };
        *source = index.clone().map(identity::SourcePackage::Ready);
        index
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
        let para_id = self.next_id();
        self.create_story_with_paragraph_id(story_id, para_id, initial_text, p_style, alignment)
    }

    pub(crate) fn create_empty_stories(&self, story_ids: &[String]) -> EditResult<()> {
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        for story_id in story_ids {
            let para_id = self.next_id();
            if stories.contains_key(&txn, story_id) {
                return Err(EditError::StoryExists(story_id.clone()));
            }
            let story = stories.insert(&mut txn, story_id.clone(), TextPrelim::new(""));
            let pilcrow = story.insert_embed_with_attributes(
                &mut txn,
                0,
                MapPrelim::default(),
                insertion_attrs(None, None),
            );
            write_pilcrow_properties(&pilcrow, &mut txn, &para_id, "Normal", "left");
        }
        Ok(())
    }

    /// Adds a one-paragraph story with a caller-supplied paragraph ID.
    pub fn create_story_with_paragraph_id(
        &self,
        story_id: impl Into<StoryId>,
        para_id: impl Into<ParagraphId>,
        initial_text: &str,
        p_style: &str,
        alignment: &str,
    ) -> EditResult<ParagraphId> {
        let story_id = story_id.into();
        let para_id = para_id.into();
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

    /// Seeds a story and returns its paragraph IDs in document order.
    pub fn seed_story(
        &self,
        story_id: impl Into<StoryId>,
        paragraphs: &[SeedParagraph],
    ) -> OpResult<Vec<ParagraphId>> {
        if paragraphs.is_empty() {
            return Err(OpError::EmptyRange);
        }
        for paragraph in &paragraphs[1..] {
            ops::text::validate_text(&paragraph.text)?;
        }
        let story_id = story_id.into();
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        if stories.contains_key(&txn, &story_id) {
            return Err(OpError::StoryExists(story_id));
        }
        let story = stories.insert(&mut txn, story_id, TextPrelim::new(""));
        let mut deltas = Vec::with_capacity(paragraphs.len() * 2);
        let mut para_ids = Vec::with_capacity(paragraphs.len());
        let attrs = Box::new(insertion_attrs(None, None));
        for paragraph in paragraphs.iter() {
            if !paragraph.text.is_empty() {
                deltas.push(Delta::Inserted(
                    In::Any(Any::String(Arc::from(paragraph.text.as_str()))),
                    Some(attrs.clone()),
                ));
            }
            let para_id = self.next_id();
            deltas.push(Delta::Inserted(
                In::Map(MapPrelim::from_iter([
                    (KIND_KEY.to_owned(), Any::from(PILCROW_KIND)),
                    (PARA_ID.to_owned(), Any::from(para_id.as_str())),
                    ("pStyle".to_owned(), Any::from(paragraph.p_style.as_str())),
                    (
                        "alignment".to_owned(),
                        Any::from(paragraph.alignment.as_str()),
                    ),
                ])),
                Some(attrs.clone()),
            ));
            para_ids.push(para_id);
        }
        story.apply_delta(&mut txn, deltas);
        Ok(para_ids)
    }

    /// Removes one complete story from the document map.
    pub fn delete_story(&self, story_id: &str) -> EditResult<()> {
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        if stories.remove(&mut txn, story_id).is_some() {
            self.chunk_snapshots
                .lock()
                .unwrap()
                .entries
                .remove(story_id);
            Ok(())
        } else {
            Err(EditError::StoryNotFound(story_id.to_owned()))
        }
    }

    /// Updates one independently-convergent property on the pilcrow identified by `para_id`.
    ///
    /// Arbitrary values leave room for `pPrIns`, `pPrDel`, `pPrChange`, and passive OOXML property
    /// bags. `paraId`, the paragraph's identity bindings and the embed discriminator are schema
    /// identity.
    pub fn set_paragraph_attr(
        &self,
        para_id: &str,
        key: impl Into<String>,
        value: Any,
    ) -> EditResult<()> {
        let key = key.into();
        if is_identity_key(&key) {
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
                    identity::promote(self, &mut txn, &pilcrow);
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
            let (from, to) = crate::ops::code_point_range(&story, &txn, range.start, range.end);
            let start = story
                .sticky_index(&txn, from, Assoc::After)
                .ok_or_else(|| {
                    EditError::InvalidComment("start anchor could not be made".into())
                })?;
            let end = story
                .sticky_index(&txn, to, Assoc::Before)
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

    /// Replaces an existing comment's non-empty ranges, preserving all metadata.
    pub fn set_comment_ranges(&self, comment_id: &str, ranges: &[StoryRange]) -> EditResult<()> {
        if ranges.is_empty() {
            return Err(EditError::InvalidComment(
                "at least one anchored range is required".into(),
            ));
        }
        let mut txn = self.doc.transact_mut_with(self.client_id);
        let comments = txn.get_map(COMMENTS).expect("comments root is declared");
        let comment = comments
            .get(&txn, comment_id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .ok_or_else(|| EditError::CommentNotFound(comment_id.to_owned()))?;
        let mut anchors = Vec::with_capacity(ranges.len());
        for range in ranges {
            let len = range.len()?;
            if len == 0 {
                return Err(EditError::InvalidComment(
                    "comment ranges must be non-empty".into(),
                ));
            }
            let story = story_ref(&txn, &range.story)?;
            check_range(&story, &txn, range.start, len)?;
            let (from, to) = crate::ops::code_point_range(&story, &txn, range.start, range.end);
            let start = story
                .sticky_index(&txn, from, Assoc::After)
                .ok_or_else(|| {
                    EditError::InvalidComment("start anchor could not be made".into())
                })?;
            let end = story
                .sticky_index(&txn, to, Assoc::Before)
                .ok_or_else(|| EditError::InvalidComment("end anchor could not be made".into()))?;
            anchors.push(anchor_value(&range.story, &start, &end));
        }
        comment.insert(&mut txn, "anchors", Any::Array(Arc::from(anchors)));
        comment_references::reconcile(&mut txn, &BTreeSet::from([comment_id.to_owned()]), true);
        Ok(())
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

    pub fn story_len(&self, story_id: &str) -> EditResult<u32> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        Ok(story.len(&txn))
    }

    /// The current story revision, and the stories that changed after `since`
    /// (created, edited, or deleted), sorted.
    pub fn stories_changed_since(&self, since: u64) -> (u64, Vec<String>) {
        let revisions = self.story_revisions.lock().unwrap();
        let mut stories: Vec<String> = revisions
            .stamped
            .iter()
            .filter(|(_, revision)| **revision > since)
            .map(|(story, _)| story.to_string())
            .collect();
        stories.sort();
        (revisions.current, stories)
    }

    /// [`Self::story_segments`] split after each pilcrow into units.
    pub fn story_segment_units(&self, story_id: &str) -> EditResult<Vec<Vec<StorySegment>>> {
        Ok(split_segment_units(self.story_segments(story_id)?))
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
        deterministic::encode_state_as_update_v1(&self.doc.transact(), &StateVector::default())
    }

    pub fn encode_state_vector_v1(&self) -> Vec<u8> {
        self.doc.transact().state_vector().encode_v1()
    }

    pub fn encode_diff_v1(&self, remote_state_vector: &[u8]) -> EditResult<Vec<u8>> {
        let state_vector = StateVector::decode_v1(remote_state_vector)
            .map_err(|error| EditError::InvalidStateVector(error.to_string()))?;
        Ok(deterministic::encode_diff_v1(
            &self.doc.transact(),
            &state_vector,
        ))
    }

    pub fn apply_update_v1(&self, bytes: &[u8]) -> EditResult<()> {
        let update = Update::decode_v1(bytes)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.integrate_update(update, UpdateOrigin::Remote)
    }

    /// Applies an update, then repairs any paragraph identities it duplicated.
    pub(crate) fn integrate_update(&self, update: Update, origin: UpdateOrigin) -> EditResult<()> {
        let _host_edit = matches!(origin, UpdateOrigin::Host)
            .then(|| batch::HostEditGuard::new(&self.host_edit_depth));
        let watch = identity::IdentityWatch::new(self);
        let reanchored = comment_references::CommentWatch::new(self);
        let result = match origin {
            UpdateOrigin::Remote => self.doc.transact_mut().apply_update(update),
            UpdateOrigin::Local => self
                .doc
                .transact_mut_with(self.client_id)
                .apply_update(update),
            UpdateOrigin::Host => self
                .doc
                .transact_mut_with(batch::HOST_ORIGIN)
                .apply_update(update),
        };
        result.map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        if watch.changed() {
            drop(watch);
            self.repair_paragraph_identities();
        }
        let reanchored = reanchored.take();
        if !reanchored.is_empty() {
            let mut txn = self.transact_for(&EditCtx::system(""));
            comment_references::reconcile(&mut txn, &reanchored, false);
        }
        Ok(())
    }

    /// Applies an update as it is, without the identity repair [`Self::apply_update_v1`] runs.
    pub(crate) fn apply_verbatim_v1(&self, bytes: &[u8]) -> EditResult<()> {
        let update = Update::decode_v1(bytes)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.doc
            .transact_mut()
            .apply_update(update)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))
    }

    /// A private replica of `state`, this document's committed state, that allocates identities
    /// as this document would: the same client id and key counter, the retained source index,
    /// and every key and Word paragraph ID this replica has seen, deleted ones included. What
    /// the fork allocates stays its own until its update is adopted.
    pub(crate) fn fork(&self, state: &[u8]) -> EditResult<Self> {
        let fork = Self::new(self.client_id);
        fork.apply_verbatim_v1(state)?;
        fork.id_counter
            .store(self.id_counter.load(Ordering::Relaxed), Ordering::Relaxed);
        if let Some(index) = self.source_index() {
            fork.retain_source(identity::SourcePackage::Ready(index));
        }
        {
            let (txn, forked) = (self.doc.transact(), fork.doc.transact());
            self.with_seen(&txn, |seen| {
                fork.with_seen(&forked, |copy| copy.inherit(seen));
            });
        }
        Ok(fork)
    }

    /// Applies a v1 update using this replica's local transaction origin.
    ///
    /// This is reserved for a local worker replica executing an edit on behalf
    /// of this document. Ordinary collaboration updates must continue through
    /// [`Self::apply_update_v1`] so local undo never captures remote work.
    pub fn apply_local_update_v1(&self, bytes: &[u8]) -> EditResult<()> {
        let update = Update::decode_v1(bytes)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.integrate_update(update, UpdateOrigin::Local)
    }

    /// Integrates another replica's host batch outside local undo history.
    pub fn apply_host_update_v1(&self, bytes: &[u8]) -> EditResult<()> {
        let update = Update::decode_v1(bytes)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.integrate_update(update, UpdateOrigin::Host)
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

/// Pilcrow keys the schema manages: identity and the embed discriminator.
fn is_identity_key(key: &str) -> bool {
    matches!(
        key,
        PARA_ID
            | KIND_KEY
            | identity::OOXML_PARA_ID
            | identity::SOURCE_PARA_ID
            | identity::PARA_ORIGIN
    )
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
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let len = out_len(&diff.insert);
        if offset >= from
            && let Out::YMap(map) = diff.insert
            && is_pilcrow(&map, txn)
        {
            return Some((offset, map));
        }
        offset += len;
    }
    None
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
                    if is_identity_key(key) {
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
        Out::YMap(map) => SegmentContent::OtherEmbed {
            kind: map_string(&map, txn, KIND_KEY).unwrap_or_default(),
            payload: embed_payload(&map, txn),
        },
        _ => SegmentContent::OtherEmbed {
            kind: String::new(),
            payload: BTreeMap::new(),
        },
    }
}

/// An embed's plain values other than its kind: a story segment's `payload`.
pub(crate) fn embed_payload<T: ReadTxn>(map: &MapRef, txn: &T) -> BTreeMap<String, Any> {
    map.iter(txn)
        .filter_map(|(key, value)| {
            if key == KIND_KEY {
                return None;
            }
            let Out::Any(value) = value else {
                return None;
            };
            Some((key.to_string(), value))
        })
        .collect()
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
    use yrs::types::ToJson;

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

    #[test]
    fn epoch_cache_serves_only_the_current_epoch() {
        let mut cache = EpochCache::default();
        cache.insert("body", 1, Arc::new(1));
        cache.insert("hdr", 1, Arc::new(2));
        assert_eq!(cache.get("body", 1).as_deref(), Some(&1));
        assert_eq!(cache.get("body", 2), None);
        cache.insert("body", 2, Arc::new(3));
        assert_eq!(
            cache.get("hdr", 2),
            None,
            "a newer epoch drops the older values"
        );
        cache.insert("hdr", 1, Arc::new(4));
        assert_eq!(
            cache.get("hdr", 2),
            None,
            "a value built before a commit is not kept"
        );
        assert_eq!(cache.entries.len(), 1);
    }

    #[test]
    fn epoch_cache_takes_only_the_current_epoch() {
        let mut cache = EpochCache::default();
        cache.insert("body", 1, Arc::new(1));
        assert_eq!(cache.take("body", 0), None);
        assert_eq!(cache.take("body", 2), None);
        assert_eq!(cache.get("body", 1).as_deref(), Some(&1));
        assert_eq!(cache.take("body", 1).as_deref(), Some(&1));
        assert_eq!(cache.get("body", 1), None);
        assert_eq!(cache.take("body", 1), None);
    }

    #[test]
    fn story_revisions_name_the_stories_each_change_touched() {
        let (a, b) = peers("one two", 1, 2);
        let (since, stories) = b.stories_changed_since(0);
        assert_eq!(stories, ["body", "header:rId7"]);
        assert_eq!(b.stories_changed_since(since).1, Vec::<String>::new());

        a.insert_text(
            &local("A"),
            Position::new("body", 3),
            "!",
            FormatPolicy::Plain,
        )
        .unwrap();
        b.apply_update_v1(&a.encode_state_as_update_v1()).unwrap();
        let (since, stories) = b.stories_changed_since(since);
        assert_eq!(stories, ["body"], "a remote text edit");

        let header = b.paragraphs("header:rId7").unwrap()[0].para_id.clone();
        b.set_paragraph_attr(&header, "keepNext", Any::Bool(true))
            .unwrap();
        let (since, stories) = b.stories_changed_since(since);
        assert_eq!(stories, ["header:rId7"], "a pilcrow property");

        b.create_story("fn:1", "note", "Normal", "left").unwrap();
        b.delete_story("header:rId7").unwrap();
        let (_, stories) = b.stories_changed_since(since);
        assert_eq!(
            stories,
            ["fn:1", "header:rId7"],
            "created and deleted stories"
        );
    }

    #[test]
    fn segment_units_split_a_story_after_each_pilcrow() {
        let doc = seed("alpha beta");
        doc.split_paragraph(&local("A"), Position::new("body", 5), None)
            .unwrap();
        let units = doc.story_segment_units("body").unwrap();
        assert_eq!(units.len(), 2);
        assert_eq!(units.concat(), doc.story_segments("body").unwrap());

        doc.insert_text(
            &local("A"),
            Position::new("body", 8),
            "x",
            FormatPolicy::Plain,
        )
        .unwrap();
        let edited = doc.story_segment_units("body").unwrap();
        let digest = |unit: &[StorySegment]| segments_digest(unit);
        assert_eq!(
            digest(&edited[0]),
            digest(&units[0]),
            "an untouched paragraph"
        );
        assert_ne!(digest(&edited[1]), digest(&units[1]));
        assert_eq!(
            digest(&doc.story_segment_units("header:rId7").unwrap()[0]),
            digest(
                &seed("alpha beta")
                    .story_segment_units("header:rId7")
                    .unwrap()[0]
            ),
            "equal content, equal digest"
        );
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

    fn seed_paragraph(text: &str) -> SeedParagraph {
        SeedParagraph {
            text: text.to_owned(),
            p_style: "Normal".to_owned(),
            alignment: "left".to_owned(),
        }
    }

    #[test]
    fn comment_reanchoring_survives_whole_text_replacement_history() {
        let doc = EditingDoc::new(801);
        doc.create_story("body", "Antes achado depois", "Normal", "left")
            .unwrap();
        let id = doc
            .add_comment(
                &[StoryRange::new("body", 6, 12)],
                "Ada",
                DATE,
                Any::from("Review body"),
            )
            .unwrap();
        let mut undo = doc.undo_manager();
        doc.replace_range(
            &local("Ada"),
            StoryRange::new("body", 0, 19),
            "Novo achado fim",
        )
        .unwrap();
        doc.set_comment_ranges(&id, &[StoryRange::new("body", 5, 11)])
            .unwrap();
        for _ in 0..3 {
            let anchor = resolved(&doc, &id);
            assert_eq!((anchor.start, anchor.end), (5, 11));
            assert!(undo.undo());
            let anchor = resolved(&doc, &id);
            assert_eq!((anchor.start, anchor.end), (6, 12));
            assert!(undo.redo());
        }
    }

    #[test]
    fn comment_reanchoring_preserves_unicode_in_history() {
        let doc = EditingDoc::new(802);
        doc.create_story("body", "Antes 🦀 depois", "Normal", "left")
            .unwrap();
        let id = doc
            .add_comment(
                &[StoryRange::new("body", 6, 8)],
                "Ada",
                DATE,
                Any::from("Review body"),
            )
            .unwrap();
        let mut undo = doc.undo_manager();
        doc.replace_range(&local("Ada"), StoryRange::new("body", 0, 15), "Novo 🦀 fim")
            .unwrap();
        doc.set_comment_ranges(&id, &[StoryRange::new("body", 5, 7)])
            .unwrap();
        let text = || {
            doc.story_segments("body")
                .unwrap()
                .into_iter()
                .filter_map(|segment| match segment.content {
                    SegmentContent::Text(text) => Some(text),
                    _ => None,
                })
                .collect::<String>()
        };
        for _ in 0..3 {
            assert_eq!(text(), "Novo 🦀 fim");
            assert_eq!(doc.story_len("body").unwrap(), 12);
            let anchor = resolved(&doc, &id);
            assert_eq!((anchor.start, anchor.end), (5, 7));
            assert!(undo.undo());
            let anchor = resolved(&doc, &id);
            assert_eq!((anchor.start, anchor.end), (6, 8));
            assert_eq!(text(), "Antes 🦀 depois");
            assert_eq!(doc.story_len("body").unwrap(), 16);
            assert!(undo.redo());
        }
    }

    #[test]
    fn comment_reanchoring_preserves_metadata_and_undo() {
        let doc = EditingDoc::new(800);
        doc.create_story("body", "first second", "Normal", "left")
            .unwrap();
        let id = doc
            .add_comment(
                &[StoryRange::new("body", 0, 5)],
                "Ada",
                DATE,
                Any::from("Review body"),
            )
            .unwrap();
        let metadata = || {
            let txn = doc.doc.transact();
            let comment = txn
                .get_map(COMMENTS)
                .unwrap()
                .get(&txn, &id)
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            comment
                .iter(&txn)
                .filter(|(key, _)| *key != "anchors")
                .map(|(key, value)| (key.to_owned(), value.to_json(&txn)))
                .collect::<BTreeMap<_, _>>()
        };
        {
            let mut txn = doc.doc.transact_mut_with(doc.client_id);
            let comment = txn
                .get_map(COMMENTS)
                .unwrap()
                .get(&txn, &id)
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            comment.insert(&mut txn, "parentId", "parent");
            comment.insert(&mut txn, "done", true);
            comment.insert(&mut txn, "custom", "retained");
        }
        let before = metadata();
        let mut undo = doc.undo_manager();
        doc.set_comment_ranges(&id, &[StoryRange::new("body", 6, 12)])
            .unwrap();
        for _ in 0..3 {
            assert_eq!(metadata(), before);
            let anchor = resolved(&doc, &id);
            assert_eq!((anchor.start, anchor.end), (6, 12));
            assert!(undo.undo());
            let anchor = resolved(&doc, &id);
            assert_eq!((anchor.start, anchor.end), (0, 5));
            assert_eq!(metadata(), before);
            assert_eq!(undo.changed_stories(), ["body"]);
            assert!(undo.redo());
        }
    }

    /// Seeds comment 1 over "first" with its reference right after it, a header, and a synced peer.
    fn referenced(doc: &EditingDoc) -> EditingDoc {
        doc.create_story("body", "first second", "Normal", "left")
            .unwrap();
        doc.create_story(HEADER, "Header", "Header", "left")
            .unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::InsertEmbed {
                    index: 5,
                    kind: "field".into(),
                    payload: vec![
                        ("modelKind".into(), Any::from("commentReference")),
                        ("commentId".into(), Any::from(1.0)),
                    ],
                    attrs: Attrs::new(),
                },
                RawOp::SetComment {
                    id: "1".into(),
                    ranges: vec![(0, 5)],
                    author: "Ada".into(),
                    date: DATE.into(),
                    body: Any::Null,
                },
            ],
            &local("Ada"),
        )
        .unwrap();
        let peer = EditingDoc::new(doc.client_id + 1);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        peer
    }

    fn reference_offsets(doc: &EditingDoc, story: &str) -> Vec<u32> {
        let mut offset = 0;
        let mut found = Vec::new();
        for segment in doc.story_segments(story).unwrap() {
            match segment.content {
                SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
                SegmentContent::OtherEmbed { payload, .. } => {
                    if payload.get("modelKind") == Some(&Any::from("commentReference")) {
                        found.push(offset);
                    }
                    offset += 1;
                }
                SegmentContent::Pilcrow(_) => offset += 1,
            }
        }
        found
    }

    #[test]
    fn comment_reanchoring_moves_its_reference_with_undo() {
        let doc = EditingDoc::new(803);
        referenced(&doc);
        let mut undo = doc.undo_manager();
        doc.set_comment_ranges("1", &[StoryRange::new("body", 7, 13)])
            .unwrap();
        for _ in 0..3 {
            assert_eq!(reference_offsets(&doc, "body"), [12]);
            let anchor = resolved(&doc, "1");
            assert_eq!((anchor.start, anchor.end), (6, 12));
            assert!(undo.undo());
            assert_eq!(reference_offsets(&doc, "body"), [5]);
            let anchor = resolved(&doc, "1");
            assert_eq!((anchor.start, anchor.end), (0, 5));
            assert!(undo.redo());
        }
    }

    #[test]
    fn comment_reanchoring_keeps_a_reference_already_after_the_range() {
        let doc = EditingDoc::new(804);
        referenced(&doc);
        doc.set_comment_ranges("1", &[StoryRange::new("body", 1, 5)])
            .unwrap();
        assert_eq!(reference_offsets(&doc, "body"), [5]);
        doc.set_comment_ranges(
            "1",
            &[StoryRange::new("body", 0, 2), StoryRange::new("body", 3, 5)],
        )
        .unwrap();
        assert_eq!(reference_offsets(&doc, "body"), [5]);
    }

    #[test]
    fn comment_reanchoring_converges_with_a_concurrent_edit() {
        let doc = EditingDoc::new(805);
        let peer = referenced(&doc);
        doc.set_comment_ranges("1", &[StoryRange::new("body", 7, 13)])
            .unwrap();
        peer.apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: 0,
                text: "xx".into(),
                attrs: Attrs::new(),
            }],
            &local("Ada"),
        )
        .unwrap();
        sync(&doc, &peer);
        for replica in [&doc, &peer] {
            assert_eq!(reference_offsets(replica, "body"), [14]);
            let anchor = resolved(replica, "1");
            assert_eq!((anchor.start, anchor.end), (8, 14));
        }
        assert_eq!(doc.story_segments("body"), peer.story_segments("body"));
    }

    const HEADER: &str = "header:rId7";

    #[test]
    fn comment_reanchoring_into_another_story_leaves_no_reference_behind() {
        let doc = EditingDoc::new(806);
        referenced(&doc);
        let mut undo = doc.undo_manager();
        doc.set_comment_ranges("1", &[StoryRange::new(HEADER, 0, 4)])
            .unwrap();
        for _ in 0..2 {
            assert!(reference_offsets(&doc, "body").is_empty());
            assert_eq!(reference_offsets(&doc, HEADER), [4]);
            assert!(undo.undo());
            assert_eq!(reference_offsets(&doc, "body"), [5]);
            assert!(reference_offsets(&doc, HEADER).is_empty());
            assert!(undo.redo());
        }
    }

    #[test]
    fn raw_comment_replacement_moves_its_reference() {
        let doc = EditingDoc::new(807);
        referenced(&doc);
        let replace = |story: &str, ranges| {
            doc.apply_raw_ops(
                story,
                vec![RawOp::SetComment {
                    id: "1".into(),
                    ranges,
                    author: "Ada".into(),
                    date: DATE.into(),
                    body: Any::Null,
                }],
                &local("Ada"),
            )
            .unwrap();
        };
        replace("body", vec![(7, 13)]);
        assert_eq!(reference_offsets(&doc, "body"), [12]);
        assert_eq!(
            (resolved(&doc, "1").start, resolved(&doc, "1").end),
            (6, 12)
        );
        replace(HEADER, vec![(0, 4)]);
        assert!(reference_offsets(&doc, "body").is_empty());
        assert_eq!(reference_offsets(&doc, HEADER), [4]);
    }

    #[test]
    fn concurrent_comment_reanchors_keep_one_reference_in_either_sync_order() {
        let moves = [
            StoryRange::new("body", 7, 13),
            StoryRange::new(HEADER, 0, 4),
        ];
        for peer_first in [false, true] {
            for (left, right) in [(&moves[0], &moves[1]), (&moves[0], &moves[0])] {
                let doc = EditingDoc::new(808);
                let peer = referenced(&doc);
                doc.set_comment_ranges("1", std::slice::from_ref(left))
                    .unwrap();
                peer.set_comment_ranges("1", std::slice::from_ref(right))
                    .unwrap();
                let (first, second) = if peer_first {
                    (&peer, &doc)
                } else {
                    (&doc, &peer)
                };
                first
                    .apply_update_v1(&second.encode_state_as_update_v1())
                    .unwrap();
                second
                    .apply_update_v1(&first.encode_state_as_update_v1())
                    .unwrap();
                sync(&doc, &peer);
                let anchor = resolved(&doc, "1");
                for replica in [&doc, &peer] {
                    let references: Vec<_> = ["body", HEADER]
                        .into_iter()
                        .flat_map(|story| {
                            reference_offsets(replica, story)
                                .into_iter()
                                .map(move |offset| (story.to_owned(), offset))
                        })
                        .collect();
                    assert_eq!(references, [(anchor.story.clone(), anchor.end)]);
                    assert_eq!(resolved(replica, "1"), anchor);
                }
                for story in ["body", HEADER] {
                    assert_eq!(doc.story_segments(story), peer.story_segments(story));
                }
            }
        }
    }

    #[test]
    fn loading_and_metadata_updates_keep_a_lone_reference_away_from_its_range() {
        let doc = EditingDoc::new(809);
        doc.create_story("body", "first second", "Normal", "left")
            .unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::InsertEmbed {
                    index: 12,
                    kind: "field".into(),
                    payload: vec![
                        ("modelKind".into(), Any::from("commentReference")),
                        ("commentId".into(), Any::from(1.0)),
                    ],
                    attrs: Attrs::new(),
                },
                RawOp::SetComment {
                    id: "1".into(),
                    ranges: vec![(0, 5)],
                    author: "Ada".into(),
                    date: DATE.into(),
                    body: Any::Null,
                },
            ],
            &local("Ada"),
        )
        .unwrap();
        let peer = EditingDoc::new(810);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(reference_offsets(&peer, "body"), [12]);
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let Some(Out::YMap(comment)) = txn
                .get_map(COMMENTS)
                .and_then(|comments| comments.get(&txn, "1"))
            else {
                panic!("comment");
            };
            comment.insert(&mut txn, "done", true);
        }
        sync(&doc, &peer);
        for replica in [&doc, &peer] {
            assert_eq!(reference_offsets(replica, "body"), [12]);
        }

        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 12, len: 1 },
                RawOp::InsertEmbed {
                    index: 8,
                    kind: "field".into(),
                    payload: vec![
                        ("modelKind".into(), Any::from("commentReference")),
                        ("commentId".into(), Any::from(1.0)),
                    ],
                    attrs: Attrs::new(),
                },
            ],
            &local("Ada"),
        )
        .unwrap();
        sync(&doc, &peer);
        for replica in [&doc, &peer] {
            assert_eq!(reference_offsets(replica, "body"), [8]);
        }
    }

    #[test]
    fn loading_and_metadata_updates_keep_every_reference_of_a_comment_across_stories() {
        let doc = EditingDoc::new(811);
        doc.create_story("body", "first second", "Normal", "left")
            .unwrap();
        doc.create_story(HEADER, "Header", "Header", "left")
            .unwrap();
        let reference = |index| RawOp::InsertEmbed {
            index,
            kind: "field".into(),
            payload: vec![
                ("modelKind".into(), Any::from("commentReference")),
                ("commentId".into(), Any::from(1.0)),
            ],
            attrs: Attrs::new(),
        };
        doc.apply_raw_ops(
            "body",
            vec![
                reference(12),
                RawOp::SetComment {
                    id: "1".into(),
                    ranges: vec![(0, 5)],
                    author: "Ada".into(),
                    date: DATE.into(),
                    body: Any::Null,
                },
            ],
            &local("Ada"),
        )
        .unwrap();
        doc.apply_raw_ops(HEADER, vec![reference(4)], &local("Ada"))
            .unwrap();
        {
            let mut txn = doc.doc.transact_mut_with(doc.client_id);
            let anchors: Vec<Any> = [("body", 0, 5), (HEADER, 0, 4)]
                .into_iter()
                .map(|(story, start, end)| {
                    let text = story_ref(&txn, story).unwrap();
                    let start = text.sticky_index(&txn, start, Assoc::After).unwrap();
                    let end = text.sticky_index(&txn, end, Assoc::Before).unwrap();
                    anchor_value(story, &start, &end)
                })
                .collect();
            let comment = txn
                .get_map(COMMENTS)
                .and_then(|comments| comments.get(&txn, "1"))
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            comment.insert(&mut txn, "anchors", Any::Array(Arc::from(anchors)));
        }
        let placed = |replica: &EditingDoc| {
            (
                reference_offsets(replica, "body"),
                reference_offsets(replica, HEADER),
            )
        };
        assert_eq!(placed(&doc), (vec![12], vec![4]));
        let peer = EditingDoc::new(812);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(placed(&peer), (vec![12], vec![4]));
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let Some(Out::YMap(comment)) = txn
                .get_map(COMMENTS)
                .and_then(|comments| comments.get(&txn, "1"))
            else {
                panic!("comment");
            };
            comment.insert(&mut txn, "done", true);
        }
        sync(&doc, &peer);
        for replica in [&doc, &peer] {
            assert_eq!(placed(replica), (vec![12], vec![4]));
        }
    }

    #[test]
    fn a_winning_start_only_reanchor_drops_the_losing_moves_reference() {
        let start_only = StoryRange::new("body", 2, 5);
        let moved = StoryRange::new(HEADER, 0, 4);
        for peer_first in [false, true] {
            for (left, right) in [(&start_only, &moved), (&moved, &start_only)] {
                let doc = EditingDoc::new(808);
                let peer = referenced(&doc);
                doc.set_comment_ranges("1", std::slice::from_ref(left))
                    .unwrap();
                peer.set_comment_ranges("1", std::slice::from_ref(right))
                    .unwrap();
                let (first, second) = if peer_first {
                    (&peer, &doc)
                } else {
                    (&doc, &peer)
                };
                first
                    .apply_update_v1(&second.encode_state_as_update_v1())
                    .unwrap();
                second
                    .apply_update_v1(&first.encode_state_as_update_v1())
                    .unwrap();
                sync(&doc, &peer);
                let anchor = resolved(&doc, "1");
                for replica in [&doc, &peer] {
                    let references: Vec<_> = ["body", HEADER]
                        .into_iter()
                        .flat_map(|story| {
                            reference_offsets(replica, story)
                                .into_iter()
                                .map(move |offset| (story.to_owned(), offset))
                        })
                        .collect();
                    assert!(references.len() <= 1, "{references:?}");
                    assert!(
                        references
                            .iter()
                            .all(|(story, offset)| *story == anchor.story && *offset == anchor.end),
                        "{references:?} vs {anchor:?}"
                    );
                    assert_eq!(resolved(replica, "1"), anchor);
                }
                for story in ["body", HEADER] {
                    assert_eq!(doc.story_segments(story), peer.story_segments(story));
                }
            }
        }
    }

    #[test]
    fn seed_story_returns_ids_in_order_and_marks_paragraphs() {
        let doc = EditingDoc::new(100);
        let ids = doc
            .seed_story(
                "body",
                &[
                    seed_paragraph("one"),
                    seed_paragraph("two"),
                    seed_paragraph(""),
                ],
            )
            .unwrap();
        assert_eq!(ids.len(), 3);
        let segments = doc.story_segments("body").unwrap();
        let joined: String = segments
            .iter()
            .filter_map(|segment| match &segment.content {
                SegmentContent::Text(text) => Some(text.clone()),
                _ => None,
            })
            .collect();
        assert_eq!(joined, "onetwo");
        assert_eq!(doc.story_len("body").unwrap(), 3 + "onetwo".len() as u32);
    }

    #[test]
    fn seed_story_rejects_breaks_without_committing() {
        let doc = EditingDoc::new(100);
        let result = doc.seed_story("body", &[seed_paragraph("ok"), seed_paragraph("bad\ntext")]);
        assert!(matches!(result, Err(OpError::TextContainsBreak)));
        assert!(matches!(
            doc.story_len("body"),
            Err(EditError::StoryNotFound(_))
        ));
    }

    #[test]
    fn seed_story_rejects_empty_and_existing() {
        let doc = EditingDoc::new(100);
        assert!(matches!(
            doc.seed_story("body", &[]),
            Err(OpError::EmptyRange)
        ));
        doc.seed_story("body", &[seed_paragraph("a")]).unwrap();
        assert!(matches!(
            doc.seed_story("body", &[seed_paragraph("b")]),
            Err(OpError::StoryExists(_))
        ));
    }

    #[test]
    fn local_worker_update_remains_owned_by_main_undo() {
        let main = seed("before");
        let worker = EditingDoc::new(200);
        worker
            .apply_update_v1(&main.encode_state_as_update_v1())
            .unwrap();
        let mut undo = main.undo_manager();

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
    fn host_worker_update_preserves_earlier_local_undo() {
        let main = seed("before");
        let mut undo = main.undo_manager();
        main.insert_text(
            &local("main"),
            Position::new("body", 6),
            " local",
            FormatPolicy::Plain,
        )
        .unwrap();
        let worker = EditingDoc::new(200);
        worker
            .apply_update_v1(&main.encode_state_as_update_v1())
            .unwrap();
        worker
            .insert_text(
                &local("worker"),
                Position::new("body", 0),
                "host ",
                FormatPolicy::Plain,
            )
            .unwrap();
        main.apply_host_update_v1(&worker.encode_state_as_update_v1())
            .unwrap();

        assert_eq!(
            main.paragraphs("body").unwrap()[0].text,
            "host before local"
        );
        assert!(undo.undo());
        assert_eq!(main.paragraphs("body").unwrap()[0].text, "host before");
        assert!(!undo.undo());
    }

    #[test]
    fn state_vector_diff_converges_and_rejects_invalid_vectors() {
        let left = seed("before");
        let right = EditingDoc::new(200);
        right
            .apply_update_v1(&left.encode_state_as_update_v1())
            .unwrap();
        let baseline = right.encode_state_vector_v1();

        left.insert_text(
            &local("left"),
            Position::new("body", 6),
            " after",
            FormatPolicy::Plain,
        )
        .unwrap();
        let update = left.encode_diff_v1(&baseline).unwrap();
        right.apply_update_v1(&update).unwrap();

        assert_eq!(left.paragraphs("body"), right.paragraphs("body"));
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert!(matches!(
            left.encode_diff_v1(&[0xff]),
            Err(EditError::InvalidStateVector(_))
        ));
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

        let mut undo = a.undo_manager();
        // Delete strictly inside the annotation, leaving both boundary identities alive. yrs can
        // follow ordinary redone chains, but does not promise the exact original side when the
        // boundary item itself is deleted and recreated (and cannot recover it after GC).
        a.delete_range(&local("A"), StoryRange::new("body", 13, 19))
            .unwrap();
        let after_delete = resolved(&a, &comment_id);
        assert_eq!((after_delete.start, after_delete.end), (6, 18));

        assert!(undo.undo());
        let after_undo = resolved(&a, &comment_id);
        assert_eq!((after_undo.start, after_undo.end), (6, 24));
        assert!(undo.redo());
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
