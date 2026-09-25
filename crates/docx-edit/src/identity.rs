//! Paragraph identities: session keys, Word paragraph IDs and anchors.
//!
//! A pilcrow's `paraId` is the opaque session key every op addresses; it is
//! never saved. A paragraph's Word paragraph ID (`w14:paraId`) is its
//! `ooxmlParaId` binding when the session allocated one, else the immutable
//! `sourceParaId` seeding recorded from the source package. Identity moves
//! with the paragraph: a split's first half keeps it and the second half is
//! allocated a fresh one, a plain merge keeps the earlier paragraph's and a
//! tracked join the later one's. Editor-only paragraphs carry
//! `paraOrigin: synthetic` until authoring promotes them.
//!
//! Allocated IDs are recorded as claims `{id}/{owner}` in the replicated
//! [`PARAGRAPH_IDS`] map and published once a save wrote them; a claim keeps
//! its ID reserved after the paragraph is deleted. Source paragraphs outside
//! the stories (separators, comments, retained XML) take IDs by occurrence in
//! [`SOURCE_PARAGRAPH_IDS`]. Duplicates are repaired by ownership: source and
//! published claims keep their IDs over unpublished ones and copies, and only
//! an explicit persistence repairs duplicated source IDs.

use std::collections::hash_map::Entry;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, OnceLock};

use docx_parse::paragraph_identity::{
    ParagraphOccurrence, allocate_paragraph_id_where, format_paragraph_id, parse_paragraph_id,
};
use yrs::branch::Branch;
use yrs::types::text::YChange;
use yrs::types::{Delta, EntryChange, Event};
use yrs::{
    Any, BranchID, DeepObservable, Doc, Map, MapRef, Observable, Out, ReadTxn, Subscription, Text,
    TextRef, Transact, TransactionMut,
};

use crate::{
    EditCtx, EditResult, EditingDoc, KIND_KEY, PARA_ID, ParagraphId, STORIES, StoryId, map_string,
    pilcrows, story_ref,
};

/// Pilcrow key: a Word paragraph ID the session allocated.
pub(crate) const OOXML_PARA_ID: &str = "ooxmlParaId";
/// Pilcrow key: the source paragraph ID seeding recorded.
pub(crate) const SOURCE_PARA_ID: &str = "sourceParaId";
/// Pilcrow key: `synthetic` for a paragraph the editor added.
pub(crate) const PARA_ORIGIN: &str = "paraOrigin";
pub(crate) const SYNTHETIC: &str = "synthetic";
/// Root map of allocation claims.
pub(crate) const PARAGRAPH_IDS: &str = "paragraphIds";
/// Root map of IDs assigned to source paragraphs outside the stories.
pub(crate) const SOURCE_PARAGRAPH_IDS: &str = "sourceParagraphIds";
/// Root map holding the opening generation.
pub(crate) const SESSION: &str = "session";
const GENERATION: &str = "generation";

/// A new opening generation: 128 bits, unique per call.
fn fresh_generation() -> String {
    let [high, low] = entropy();
    format!("{high:016x}{low:016x}")
}

#[cfg(all(feature = "wasm", target_arch = "wasm32"))]
fn entropy() -> [u64; 2] {
    let draw = || (js_sys::Math::random() * 9_007_199_254_740_992.0) as u64;
    [draw() ^ ((js_sys::Date::now() as u64) << 11), draw()]
}

#[cfg(not(all(feature = "wasm", target_arch = "wasm32")))]
fn entropy() -> [u64; 2] {
    use std::hash::{BuildHasher, Hasher};
    static OPENINGS: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
    let opening = OPENINGS.fetch_add(1, Ordering::Relaxed);
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0, |elapsed| elapsed.as_nanos());
    [0u64, 1].map(|lane| {
        let mut hasher = std::collections::hash_map::RandomState::new().build_hasher();
        hasher.write_u64(lane);
        hasher.write_u128(nanos);
        hasher.write_u64(opening);
        hasher.finish()
    })
}

/// Kind of Word story a source part holds.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum SourceStoryKind {
    Body,
    Header,
    Footer,
    Footnote,
    Endnote,
    Comment,
}

/// A Word story qualified by the package part it is read from.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct SourceStory {
    /// The OPC part name, such as `/word/document.xml`.
    pub part_uri: String,
    pub kind: SourceStoryKind,
    /// The note or comment ID, for the stories that share a part.
    pub item_id: Option<String>,
}

/// A paragraph occurrence in the exact source package.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct SourceParagraphRef {
    pub package_sha256: String,
    pub part_uri: String,
    /// Zero-based among the part's `w:p` elements in document order.
    pub paragraph_ordinal: u32,
}

/// Where a paragraph came from.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ParagraphOrigin {
    Source,
    Authored,
    /// Added by the editor; it saves without an ID until content is authored into it.
    Synthetic,
}

/// Where a paragraph's Word paragraph ID comes from.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ParagraphIdOrigin {
    /// Present in the source package.
    Source,
    /// Allocated when the paragraph was authored.
    Authored,
    /// Allocated by [`EditingDoc::persist_paragraph_ids`].
    Persisted,
    /// Allocated to replace a duplicate.
    Repaired,
}

impl ParagraphIdOrigin {
    /// The lowercase wire name.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Source => "source",
            Self::Authored => "authored",
            Self::Persisted => "persisted",
            Self::Repaired => "repaired",
        }
    }

    fn parse(value: &str) -> Self {
        match value {
            "persisted" => Self::Persisted,
            "repaired" => Self::Repaired,
            _ => Self::Authored,
        }
    }
}

/// Addresses one paragraph.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ParagraphAnchor {
    /// A session key in one story of one collaborative document session.
    Session {
        session_id: String,
        story: StoryId,
        para_id: ParagraphId,
    },
    /// A paragraph occurrence in the exact source package.
    Source(SourceParagraphRef),
    /// A saved Word paragraph ID within its source story.
    Persisted { story: SourceStory, para_id: String },
}

/// The paragraph an anchor, assignment or diagnostic names.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub enum ParagraphRef {
    Session {
        story: StoryId,
        para_id: ParagraphId,
    },
    /// A source paragraph outside the session stories.
    Source(SourceParagraphRef),
}

/// One paragraph's identities.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphIdentity {
    pub paragraph: ParagraphRef,
    pub origin: ParagraphOrigin,
    /// The Word paragraph ID the paragraph saves with.
    pub ooxml_para_id: Option<String>,
    pub id_origin: Option<ParagraphIdOrigin>,
    /// The story its persisted anchor is qualified by.
    pub source_story: Option<SourceStory>,
    /// Its occurrence in the retained source package.
    pub source: Option<SourceParagraphRef>,
}

/// Every paragraph's identities: session paragraphs with stories sorted and in
/// document order, then source paragraphs outside the stories in package order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphIdentities {
    pub session_id: String,
    pub package_sha256: Option<String>,
    pub paragraphs: Vec<ParagraphIdentity>,
}

/// Why an anchor cannot be resolved here.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AnchorUnsupported {
    ForeignSession,
    ForeignPackage,
    NoSourcePackage,
}

/// Result of [`EditingDoc::resolve_paragraph_anchor`].
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum AnchorResolution {
    Found(ParagraphRef),
    Missing,
    Ambiguous(Vec<ParagraphRef>),
    Unsupported(AnchorUnsupported),
}

/// A paragraph [`EditingDoc::persist_paragraph_ids`] gave a Word paragraph ID.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphIdAssignment {
    pub paragraph: ParagraphRef,
    /// The duplicated session key the paragraph carried before repair.
    pub replaced_para_id: Option<ParagraphId>,
    /// The duplicated Word paragraph ID it replaces.
    pub previous_ooxml_para_id: Option<String>,
    pub ooxml_para_id: String,
    pub origin: ParagraphIdOrigin,
    /// The story its persisted anchor is qualified by.
    pub source_story: Option<SourceStory>,
}

/// A condition persistence reports without failing.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ParagraphIdDiagnostic {
    /// Saved claims to one ID conflict; the paragraphs keep it and resolve as ambiguous.
    ConflictingSavedIds {
        ooxml_para_id: String,
        paragraphs: Vec<ParagraphRef>,
    },
    /// No retained source package, so source paragraphs outside the stories were not covered.
    NoSourcePackage,
}

/// What [`EditingDoc::persist_paragraph_ids`] applied.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct PersistedParagraphIds {
    pub assignments: Vec<ParagraphIdAssignment>,
    pub diagnostics: Vec<ParagraphIdDiagnostic>,
}

/// Why [`EditingDoc::persist_paragraph_ids`] changed nothing.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ParagraphIdRefusal {
    /// A duplicated comment paragraph ID needs repair, but the comment
    /// companion parts reference it, so which comment they mean is ambiguous.
    AmbiguousCommentReference {
        ooxml_para_id: String,
        comment_ids: Vec<String>,
    },
}

/// Paragraph IDs a save applies: IDs for source paragraphs outside the
/// stories, by part and occurrence, and the story parts unchanged since
/// seeding, which it writes as their source bytes with IDs patched in.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct ParagraphSavePlan {
    /// `(part path, ordinal, ID)`.
    pub assignments: Vec<(String, u32, String)>,
    /// `(part path, [(ordinal, ID)])`.
    pub patched_parts: Vec<(String, Vec<(u32, String)>)>,
}

/// One paragraph as the seeder lowers it, in document order.
pub(crate) struct SeededParagraph {
    pub(crate) root: String,
    pub(crate) key: String,
    pub(crate) source_para_id: Option<String>,
    /// The paragraph's `w:p` occurrence in its part.
    pub(crate) ordinal: Option<u32>,
    /// `false` for the paragraphs seeding adds.
    pub(crate) source: bool,
}

/// One story part: its path, kind, XML and the root stories seeded from it.
pub(crate) struct SourcePartInput {
    pub(crate) path: String,
    pub(crate) kind: SourceStoryKind,
    pub(crate) xml: String,
    /// `(root story, note ID)`.
    pub(crate) roots: Vec<(String, Option<String>)>,
}

struct SourcePart {
    uri: String,
    kind: SourceStoryKind,
    occurrences: Vec<ParagraphOccurrence>,
    /// Occurrence ordinal to the session keys seeded from it: one per root
    /// story sharing the part, each a view of the same paragraph, first seeded first.
    backed: HashMap<u32, Vec<String>>,
    roots: Vec<String>,
}

impl SourcePart {
    fn story(&self, item_id: Option<&str>) -> SourceStory {
        let itemized = matches!(
            self.kind,
            SourceStoryKind::Footnote | SourceStoryKind::Endnote | SourceStoryKind::Comment
        );
        SourceStory {
            part_uri: self.uri.clone(),
            kind: self.kind,
            item_id: item_id.filter(|_| itemized).map(str::to_owned),
        }
    }

    fn path(&self) -> &str {
        &self.uri[1..]
    }

    fn occurrence(&self, ordinal: u32) -> Option<&ParagraphOccurrence> {
        self.occurrences.get(ordinal as usize)
    }
}

struct Seeded {
    root: String,
    part: Option<usize>,
    ordinal: Option<u32>,
    source_para_id: Option<String>,
}

/// Identity index of the retained source package, reconstructible from its bytes.
pub(crate) struct SourceIndex {
    package_sha256: String,
    bytes: Arc<[u8]>,
    occupied: BTreeSet<u32>,
    /// Story parts in package order.
    parts: Vec<SourcePart>,
    roots: HashMap<String, SourceStory>,
    seeded: HashMap<String, Seeded>,
    /// IDs the comment companion parts reference.
    comment_references: BTreeSet<u32>,
    seed_states: OnceLock<HashMap<String, StoryState>>,
}

impl SourceIndex {
    pub(crate) fn new(
        package_sha256: String,
        bytes: Arc<[u8]>,
        occupied: BTreeSet<u32>,
        inputs: Vec<SourcePartInput>,
        comment_references: BTreeSet<u32>,
        lowered: Vec<SeededParagraph>,
    ) -> Self {
        let mut parts = Vec::new();
        let mut roots = HashMap::new();
        let mut part_of_root = HashMap::new();
        for input in inputs {
            let part = SourcePart {
                uri: format!("/{}", input.path),
                kind: input.kind,
                occurrences: docx_parse::paragraph_identity::paragraph_occurrences(&input.xml)
                    .unwrap_or_default(),
                backed: HashMap::new(),
                roots: input.roots.iter().map(|(root, _)| root.clone()).collect(),
            };
            for (root, item_id) in &input.roots {
                roots.insert(root.clone(), part.story(item_id.as_deref()));
                part_of_root.insert(root.clone(), parts.len());
            }
            parts.push(part);
        }
        let mut seeded = HashMap::new();
        for paragraph in lowered.into_iter().filter(|paragraph| paragraph.source) {
            let part = part_of_root.get(&paragraph.root).copied();
            let mut ordinal = None;
            if let (Some(index), Some(candidate)) = (part, paragraph.ordinal) {
                let views = parts[index].backed.entry(candidate).or_default();
                let viewed = |key: &String| {
                    seeded
                        .get(key)
                        .is_some_and(|seed: &Seeded| seed.root == paragraph.root)
                };
                if !views.iter().any(viewed) {
                    views.push(paragraph.key.clone());
                    ordinal = Some(candidate);
                }
            }
            seeded.insert(
                paragraph.key,
                Seeded {
                    root: paragraph.root,
                    part,
                    ordinal,
                    source_para_id: paragraph.source_para_id,
                },
            );
        }
        Self {
            package_sha256,
            bytes,
            occupied,
            parts,
            roots,
            seeded,
            comment_references,
            seed_states: OnceLock::new(),
        }
    }

    fn part(&self, uri: &str) -> Option<&SourcePart> {
        self.parts.iter().find(|part| part.uri == uri)
    }

    fn reference(&self, part: &SourcePart, ordinal: u32) -> SourceParagraphRef {
        SourceParagraphRef {
            package_sha256: self.package_sha256.clone(),
            part_uri: part.uri.clone(),
            paragraph_ordinal: ordinal,
        }
    }

    /// Source paragraphs no seeded story holds, in package order.
    fn unbacked(&self) -> impl Iterator<Item = (usize, &ParagraphOccurrence)> {
        self.parts.iter().enumerate().flat_map(|(index, part)| {
            part.occurrences
                .iter()
                .filter(move |occurrence| !part.backed.contains_key(&occurrence.ordinal))
                .map(move |occurrence| (index, occurrence))
        })
    }

    /// The part and occurrence `key` was seeded from, and the rank of its
    /// view among the stories sharing the part, the first seeded first.
    fn view_of(&self, key: &str) -> Option<((usize, u32), usize)> {
        let seed = self.seeded.get(key)?;
        let (part, ordinal) = seed.part.zip(seed.ordinal)?;
        let rank = self.parts[part]
            .backed
            .get(&ordinal)?
            .iter()
            .position(|view| view == key)?;
        Some(((part, ordinal), rank))
    }

    /// Keeps, of the paragraphs that view one occurrence, the first-ranked.
    fn first_views<'a>(&self, pilcrows: impl Iterator<Item = &'a Pilcrow>) -> Vec<&'a Pilcrow> {
        let mut kept: Vec<&Pilcrow> = Vec::new();
        let mut viewed: HashMap<(usize, u32), (usize, usize)> = HashMap::new();
        for pilcrow in pilcrows {
            if let Some((occurrence, rank)) = self.view_of(&pilcrow.key) {
                match viewed.entry(occurrence) {
                    Entry::Occupied(mut entry) => {
                        let (best, at) = *entry.get();
                        if rank < best {
                            kept[at] = pilcrow;
                            entry.insert((rank, at));
                        }
                        continue;
                    }
                    Entry::Vacant(entry) => {
                        entry.insert((rank, kept.len()));
                    }
                }
            }
            kept.push(pilcrow);
        }
        kept
    }

    /// Whether `key` names the paragraph seeded from this package, rather than a copy.
    fn seeds(&self, key: &str, id: u32) -> bool {
        self.seeded
            .get(key)
            .and_then(|seed| seed.source_para_id.as_deref())
            .and_then(parse_paragraph_id)
            == Some(id)
    }

    /// Every story as seeded; see [`StoryState`].
    fn seed_states(&self) -> &HashMap<String, StoryState> {
        self.seed_states.get_or_init(|| {
            let scratch = EditingDoc::new(0);
            if crate::seed::seed_stories(&scratch, &self.bytes).is_err() {
                return HashMap::new();
            }
            story_states(&scratch)
        })
    }
}

/// The retained source package: its bytes until an identity read needs the index.
pub(crate) enum SourcePackage {
    Pending(Arc<[u8]>),
    Ready(Arc<SourceIndex>),
}

/// What a save projects from one story, compared with the story as seeded
/// to decide whether its part saves as its source bytes.
struct StoryState {
    root: String,
    /// Digest of every segment the projection reads: see [`story_fingerprint`].
    fingerprint: [u8; 32],
    /// Where each comment anchored in the story starts and ends, carets and
    /// ranges over embeds included: the projection's comment markers.
    comments: BTreeMap<String, Vec<(u32, u32)>>,
}

fn story_states(doc: &EditingDoc) -> HashMap<String, StoryState> {
    let (scan, mut comments) = {
        let txn = doc.yrs_doc().transact();
        let scan = Scan::new(&txn);
        let comments: HashMap<String, BTreeMap<String, Vec<(u32, u32)>>> = scan
            .stories
            .iter()
            .map(|story| {
                let anchors = crate::canonical::story_comment_anchors(&txn, story);
                (story.clone(), anchors)
            })
            .collect();
        (scan, comments)
    };
    scan.stories
        .iter()
        .filter_map(|story| {
            Some((
                story.clone(),
                StoryState {
                    root: scan.root(story).to_owned(),
                    fingerprint: story_fingerprint(doc, story)?,
                    comments: comments.remove(story).unwrap_or_default(),
                },
            ))
        })
        .collect()
}

/// SHA-256 over every segment of a story exactly as
/// [`EditingDoc::story_segments`] hands it to the save projection: each
/// text, paragraph mark and embed with its full payload and its full
/// attribute map (tracked insertions and deletions, formatting), map keys
/// sorted and nulls kept. The one exclusion is the paragraph identity the
/// save plan patches in place: the segments already leave out the session
/// key, Word paragraph ID, source ID and editor-only marker.
fn story_fingerprint(doc: &EditingDoc, story: &str) -> Option<[u8; 32]> {
    use sha2::{Digest, Sha256};
    fn ordered(value: &Any) -> serde_json::Value {
        match value {
            Any::Map(map) => ordered_map(map.iter()),
            Any::Array(values) => serde_json::Value::Array(values.iter().map(ordered).collect()),
            value => serde_json::to_value(value).unwrap_or(serde_json::Value::Null),
        }
    }
    fn ordered_map<'a>(entries: impl Iterator<Item = (&'a String, &'a Any)>) -> serde_json::Value {
        let sorted: BTreeMap<&String, serde_json::Value> =
            entries.map(|(key, value)| (key, ordered(value))).collect();
        serde_json::Value::Object(
            sorted
                .into_iter()
                .map(|(key, value)| (key.clone(), value))
                .collect(),
        )
    }
    let mut hasher = Sha256::new();
    for segment in doc.story_segments(story).ok()? {
        let content = match &segment.content {
            crate::SegmentContent::Text(text) => serde_json::json!({ "text": text }),
            crate::SegmentContent::Pilcrow(properties) => {
                serde_json::json!({ "pilcrow": ordered_map(properties.values.iter()) })
            }
            crate::SegmentContent::OtherEmbed { kind, payload } => {
                serde_json::json!({ "embed": kind, "payload": ordered_map(payload.iter()) })
            }
        };
        let entry = serde_json::json!([content, ordered_map(segment.attributes.iter())]);
        hasher.update(serde_json::to_vec(&entry).ok()?);
        hasher.update(b"\n");
    }
    Some(hasher.finalize().into())
}

fn valid(value: Option<String>) -> Option<String> {
    value.filter(|value| parse_paragraph_id(value).is_some())
}

/// A renamed duplicate's key, derived from its mark's CRDT identity so every
/// replica repairing the same state picks the same one.
fn derived_key(map: &MapRef) -> String {
    match AsRef::<Branch>::as_ref(map).id() {
        BranchID::Nested(id) => format!("{}~{}", id.client, id.clock),
        BranchID::Root(name) => name.to_string(),
    }
}

fn occurrence_owner(part: &SourcePart, ordinal: u32) -> String {
    format!("{}#{ordinal}", part.uri)
}

struct Pilcrow {
    story: String,
    key: String,
    map: MapRef,
    allocated: Option<String>,
    source: Option<String>,
    synthetic: bool,
    /// The paragraph holds text or inline content, not only block embeds.
    content: bool,
}

impl Pilcrow {
    fn read<T: ReadTxn>(story: &str, map: MapRef, txn: &T) -> Self {
        Self {
            story: story.to_owned(),
            key: map_string(&map, txn, PARA_ID).unwrap_or_default(),
            allocated: valid(map_string(&map, txn, OOXML_PARA_ID)),
            source: valid(map_string(&map, txn, SOURCE_PARA_ID)),
            synthetic: map_string(&map, txn, PARA_ORIGIN).as_deref() == Some(SYNTHETIC),
            content: false,
            map,
        }
    }

    fn session_ref(&self) -> ParagraphRef {
        ParagraphRef::Session {
            story: self.story.clone(),
            para_id: self.key.clone(),
        }
    }

    /// The ID it saves with, and whether it is the source ID seeding recorded.
    fn saved_id(&self, source: Option<&SourceIndex>) -> Option<(String, bool)> {
        self.allocated
            .clone()
            .map(|id| (id, false))
            .or_else(|| self.source.clone().map(|id| (id, true)))
            .or_else(|| {
                valid(source?.seeded.get(&self.key)?.source_para_id.clone()).map(|id| (id, true))
            })
    }
}

/// Every story and pilcrow, stories in sorted order, plus each nested story's parent.
struct Scan {
    stories: Vec<String>,
    pilcrows: Vec<Pilcrow>,
    parents: HashMap<String, String>,
}

impl Scan {
    fn new<T: ReadTxn>(txn: &T) -> Self {
        let mut scan = Self {
            stories: Vec::new(),
            pilcrows: Vec::new(),
            parents: HashMap::new(),
        };
        let Some(stories) = txn.get_map(STORIES) else {
            return scan;
        };
        let mut texts: Vec<(String, TextRef)> = stories
            .iter(txn)
            .filter_map(|(id, value)| match value {
                Out::YText(text) => Some((id.to_string(), text)),
                _ => None,
            })
            .collect();
        texts.sort_by(|left, right| left.0.cmp(&right.0));
        for (story, text) in texts {
            let mut content = false;
            for diff in text.diff(txn, YChange::identity) {
                let Out::YMap(map) = diff.insert else {
                    content = true;
                    continue;
                };
                match map_string(&map, txn, KIND_KEY).as_deref() {
                    Some(crate::PILCROW_KIND) => {
                        let mut pilcrow = Pilcrow::read(&story, map, txn);
                        pilcrow.content = std::mem::take(&mut content);
                        scan.pilcrows.push(pilcrow);
                    }
                    Some("table") => {
                        let Some(Out::Any(Any::Array(rows))) = map.get(txn, "rows") else {
                            continue;
                        };
                        for row in rows.iter() {
                            let Any::Map(row) = row else { continue };
                            let Some(Any::Array(cells)) = row.get("cells") else {
                                continue;
                            };
                            for cell in cells.iter() {
                                if let Any::Map(cell) = cell
                                    && let Some(Any::String(child)) = cell.get("story")
                                {
                                    scan.parents.insert(child.to_string(), story.clone());
                                }
                            }
                        }
                    }
                    Some("blockSdt") => {
                        if let Some(child) = map_string(&map, txn, "story") {
                            scan.parents.insert(child, story.clone());
                        }
                    }
                    _ => content = true,
                }
            }
            scan.stories.push(story);
        }
        scan
    }

    /// The root story a story is nested in, itself when it is a root.
    fn root<'a>(&'a self, mut story: &'a str) -> &'a str {
        for _ in 0..=self.parents.len() {
            match self.parents.get(story) {
                Some(parent) => story = parent,
                None => break,
            }
        }
        story
    }
}

struct Claim {
    owner: String,
    origin: ParagraphIdOrigin,
    published: bool,
}

fn claims<T: ReadTxn>(txn: &T) -> HashMap<u32, Vec<Claim>> {
    let mut claims: HashMap<u32, Vec<Claim>> = HashMap::new();
    let Some(map) = txn.get_map(PARAGRAPH_IDS) else {
        return claims;
    };
    for (key, value) in map.iter(txn) {
        let Some((id, owner)) = key.split_once('/') else {
            continue;
        };
        let (Some(id), Out::Any(Any::Map(value))) = (parse_paragraph_id(id), value) else {
            continue;
        };
        claims.entry(id).or_default().push(Claim {
            owner: owner.to_owned(),
            origin: match value.get("origin") {
                Some(Any::String(origin)) => ParagraphIdOrigin::parse(origin),
                _ => ParagraphIdOrigin::Authored,
            },
            published: matches!(value.get("published"), Some(Any::Bool(true))),
        });
    }
    claims
}

fn claim<'a>(claims: &'a HashMap<u32, Vec<Claim>>, id: &str, owner: &str) -> Option<&'a Claim> {
    claims
        .get(&parse_paragraph_id(id)?)?
        .iter()
        .find(|claim| claim.owner == owner)
}

/// Claims `id` for `owner` and binds it to `pilcrow`.
fn attach(
    txn: &mut TransactionMut<'_>,
    pilcrow: &MapRef,
    owner: &str,
    id: &str,
    origin: ParagraphIdOrigin,
) {
    record_claim(txn, id, owner, origin, false);
    pilcrow.insert(txn, OOXML_PARA_ID, id);
}

fn source_assignments<T: ReadTxn>(txn: &T) -> HashMap<String, String> {
    txn.get_map(SOURCE_PARAGRAPH_IDS)
        .map(|map| {
            map.iter(txn)
                .filter_map(|(key, value)| match value {
                    Out::Any(Any::String(id)) if parse_paragraph_id(&id).is_some() => {
                        Some((key.to_owned(), id.to_string()))
                    }
                    _ => None,
                })
                .collect()
        })
        .unwrap_or_default()
}

fn record_claim(
    txn: &mut TransactionMut<'_>,
    id: &str,
    owner: &str,
    origin: ParagraphIdOrigin,
    published: bool,
) {
    txn.get_map(PARAGRAPH_IDS)
        .expect("paragraph IDs root is declared by EditingDoc::new")
        .insert(
            txn,
            format!("{id}/{owner}"),
            Any::Map(Arc::new(HashMap::from([
                ("origin".to_owned(), Any::from(origin.as_str())),
                ("published".to_owned(), Any::Bool(published)),
            ]))),
        );
}

/// Every Word paragraph ID and session key this replica has seen, whether
/// current, claimed, assigned or since deleted, and every pilcrow seeded as
/// editor-only. Built on first use and kept current by [`observe_seen`];
/// entries only accumulate, so a deleted identity stays reserved and a
/// promotion that undo reverts is found again.
#[derive(Default)]
pub(crate) struct Seen {
    ids: HashSet<u32>,
    keys: HashSet<String>,
    /// Pilcrows seeded as editor-only, with their stories.
    synthetic: Vec<(String, MapRef)>,
}

pub(crate) type SeenCell = Arc<Mutex<Option<Seen>>>;

impl Seen {
    pub(crate) fn scan<T: ReadTxn>(txn: &T) -> Self {
        let mut seen = Self::default();
        for pilcrow in Scan::new(txn).pilcrows {
            seen.add(&pilcrow);
        }
        for (id, owners) in claims(txn) {
            seen.ids.insert(id);
            seen.keys
                .extend(owners.into_iter().map(|claim| claim.owner));
        }
        seen.ids.extend(
            source_assignments(txn)
                .values()
                .filter_map(|id| parse_paragraph_id(id)),
        );
        seen
    }

    fn add(&mut self, pilcrow: &Pilcrow) {
        self.keys.insert(pilcrow.key.clone());
        for id in [&pilcrow.allocated, &pilcrow.source].into_iter().flatten() {
            self.ids.extend(parse_paragraph_id(id));
        }
        if pilcrow.synthetic && !self.synthetic.iter().any(|(_, map)| *map == pilcrow.map) {
            self.synthetic
                .push((pilcrow.story.clone(), pilcrow.map.clone()));
        }
    }

    fn add_story<T: ReadTxn>(&mut self, story: &str, text: &TextRef, txn: &T) {
        for diff in text.diff(txn, YChange::identity) {
            if let Out::YMap(map) = diff.insert
                && map_string(&map, txn, KIND_KEY).as_deref() == Some(crate::PILCROW_KIND)
            {
                self.add(&Pilcrow::read(story, map, txn));
            }
        }
    }
}

/// Keeps a replica's [`Seen`] identities current once they are first read:
/// stories as they are added, and every claim and assignment. A pilcrow
/// inserted into an existing story or re-keyed carries a copied identity or
/// one allocated, and so claimed, already.
pub(crate) fn observe_seen(doc: &Doc, cell: &SeenCell) -> Vec<Subscription> {
    let (stories, claimed, assigned) = {
        let txn = doc.transact();
        (
            txn.get_map(STORIES),
            txn.get_map(PARAGRAPH_IDS),
            txn.get_map(SOURCE_PARAGRAPH_IDS),
        )
    };
    let mut subscriptions = Vec::new();
    if let Some(stories) = stories {
        let cell = Arc::clone(cell);
        subscriptions.push(stories.observe(move |txn, event| {
            if let Some(seen) = cell.lock().unwrap().as_mut() {
                for (story, change) in event.keys(txn) {
                    if let EntryChange::Inserted(Out::YText(text))
                    | EntryChange::Updated(_, Out::YText(text)) = change
                    {
                        seen.add_story(story, text, txn);
                    }
                }
            }
        }));
    }
    if let Some(claimed) = claimed {
        let cell = Arc::clone(cell);
        subscriptions.push(claimed.observe(move |txn, event| {
            if let Some(seen) = cell.lock().unwrap().as_mut() {
                for key in event.keys(txn).keys() {
                    if let Some((id, owner)) = key.split_once('/') {
                        seen.ids.extend(parse_paragraph_id(id));
                        seen.keys.insert(owner.to_owned());
                    }
                }
            }
        }));
    }
    if let Some(assigned) = assigned {
        let cell = Arc::clone(cell);
        subscriptions.push(assigned.observe(move |txn, event| {
            if let Some(seen) = cell.lock().unwrap().as_mut() {
                for change in event.keys(txn).values() {
                    if let EntryChange::Inserted(Out::Any(Any::String(id)))
                    | EntryChange::Updated(_, Out::Any(Any::String(id))) = change
                    {
                        seen.ids.extend(parse_paragraph_id(id));
                    }
                }
            }
        }));
    }
    subscriptions
}

/// Allocates Word paragraph IDs and session keys that no current, reserved,
/// assigned or source identity uses.
pub(crate) struct IdAllocator {
    seen: SeenCell,
    source: Option<Arc<SourceIndex>>,
}

impl IdAllocator {
    pub(crate) fn new<T: ReadTxn>(doc: &EditingDoc, txn: &T) -> Self {
        doc.with_seen(txn, |_| ());
        Self {
            seen: doc.seen_cell(),
            source: doc.source_index(),
        }
    }

    fn allocate(&mut self, owner: &str) -> String {
        let mut guard = self.seen.lock().unwrap();
        let seen = guard
            .as_mut()
            .expect("IdAllocator::new builds the seen identities");
        let source = self.source.as_deref();
        let id = allocate_paragraph_id_where(owner, |id| {
            seen.ids.contains(&id) || source.is_some_and(|source| source.occupied.contains(&id))
        })
        .expect("the paragraph ID space outnumbers every document's paragraphs");
        seen.ids.insert(id);
        format_paragraph_id(id)
    }

    /// The next `clientId:counter` key no paragraph has carried.
    pub(crate) fn session_key(&mut self, doc: &EditingDoc) -> String {
        let mut guard = self.seen.lock().unwrap();
        let seen = guard
            .as_mut()
            .expect("IdAllocator::new builds the seen identities");
        loop {
            let key = doc.next_id();
            if seen.keys.insert(key.clone()) {
                return key;
            }
        }
    }

    /// Allocates, claims and binds a fresh ID to `pilcrow`.
    pub(crate) fn bind(
        &mut self,
        txn: &mut TransactionMut<'_>,
        pilcrow: &MapRef,
        owner: &str,
        origin: ParagraphIdOrigin,
    ) -> String {
        let id = self.allocate(owner);
        attach(txn, pilcrow, owner, &id, origin);
        id
    }

    fn assign_source(
        &mut self,
        txn: &mut TransactionMut<'_>,
        owner: &str,
        origin: ParagraphIdOrigin,
    ) -> String {
        let id = self.allocate(owner);
        record_claim(txn, &id, owner, origin, false);
        txn.get_map(SOURCE_PARAGRAPH_IDS)
            .expect("source paragraph IDs root is declared by EditingDoc::new")
            .insert(txn, owner, id.as_str());
        id
    }
}

/// Makes an editor-only paragraph authored content in the mutation that
/// authors into it: it loses the synthetic marker and is allocated its own
/// Word paragraph ID.
pub(crate) fn promote(doc: &EditingDoc, txn: &mut TransactionMut<'_>, pilcrow: &MapRef) {
    if map_string(pilcrow, txn, PARA_ORIGIN).as_deref() != Some(SYNTHETIC) {
        return;
    }
    pilcrow.remove(txn, PARA_ORIGIN);
    let key = map_string(pilcrow, txn, PARA_ID).unwrap_or_default();
    IdAllocator::new(doc, txn).bind(txn, pilcrow, &key, ParagraphIdOrigin::Authored);
}

/// [`promote`] for an edit at story unit `index`. Editor-only paragraphs are
/// empty and end their story, so only an edit at a story's final unit can
/// author into one; whether a seeded one still is editor-only is read from
/// its replicated marker.
pub(crate) fn promote_at(
    doc: &EditingDoc,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    story: &TextRef,
    index: u32,
) {
    if index + 1 != story.len(txn) {
        return;
    }
    let seeded: Vec<MapRef> = doc.with_seen(&*txn, |seen| {
        seen.synthetic
            .iter()
            .filter(|(story, _)| story == story_id)
            .map(|(_, pilcrow)| pilcrow.clone())
            .collect()
    });
    for pilcrow in seeded {
        promote(doc, txn, &pilcrow);
    }
}

/// How a claimant holds the ID it shares with others.
#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
enum Standing {
    /// Neither the seeded paragraph nor holding its own claim: a copy.
    Unowned,
    /// Claimed by its own key, not yet saved.
    Unpublished,
    Published,
    /// Recorded from the source package.
    Source,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Holder {
    Pilcrow(usize),
    Occurrence(usize, u32),
    /// A source ID or saved claim no paragraph holds any more: it stays reserved.
    Retired,
}

struct Claimant {
    holder: Holder,
    id: u32,
    text: String,
    owner: String,
    standing: Standing,
    /// Other stories' views of the holder's source paragraph, sharing its ID.
    views: Vec<usize>,
}

fn claim_standing(claims: &HashMap<u32, Vec<Claim>>, id: &str, owner: &str) -> Standing {
    match claim(claims, id, owner) {
        Some(claim) if claim.published => Standing::Published,
        Some(_) => Standing::Unpublished,
        None => Standing::Unowned,
    }
}

/// How a pilcrow holds the ID it saves with under its own key.
fn pilcrow_standing(
    pilcrow: &Pilcrow,
    source: Option<&SourceIndex>,
    claims: &HashMap<u32, Vec<Claim>>,
) -> Standing {
    let Some((text, from_source)) = pilcrow.saved_id(source) else {
        return Standing::Unowned;
    };
    match (from_source, source) {
        (true, Some(source)) => match parse_paragraph_id(&text) {
            Some(id) if source.seeds(&pilcrow.key, id) => Standing::Source,
            _ => Standing::Unowned,
        },
        (true, None) => Standing::Source,
        (false, _) => claim_standing(claims, &text, &pilcrow.key),
    }
}

/// Groups the holders of each shared ID in package order: seeded paragraphs
/// and source occurrences by part and ordinal, then other paragraphs by story,
/// then the source IDs and saved claims no paragraph holds any more, which
/// stay reserved. Views of one source paragraph in stories sharing its part
/// are one holder. A `copies` pilcrow holds no identity it carries, and a
/// `renamed` one gives up its key. Only groups a live holder shares are
/// returned.
fn collisions(
    scan: &Scan,
    source: Option<&SourceIndex>,
    claims: &HashMap<u32, Vec<Claim>>,
    assignments: &HashMap<String, String>,
    copies: &HashSet<usize>,
    renamed: &HashMap<usize, &str>,
) -> Vec<Vec<Claimant>> {
    let mut ordered: Vec<((usize, u32, usize), Claimant)> = Vec::new();
    let mut live: HashSet<(String, u32)> = HashSet::new();
    let mut viewed: HashMap<((usize, u32), u32), (usize, usize)> = HashMap::new();
    for (index, pilcrow) in scan.pilcrows.iter().enumerate() {
        let Some((text, _)) = pilcrow.saved_id(source) else {
            continue;
        };
        let Some(id) = parse_paragraph_id(&text) else {
            continue;
        };
        let copy = copies.contains(&index);
        let rename = renamed.get(&index).copied();
        let standing = if copy || rename.is_some() {
            Standing::Unowned
        } else {
            pilcrow_standing(pilcrow, source, claims)
        };
        let view = source
            .filter(|_| !copy && rename.is_none())
            .and_then(|source| Some((source, source.view_of(&pilcrow.key)?)));
        if !copy {
            live.insert((pilcrow.key.clone(), id));
        }
        if let Some((source, ((part, ordinal), rank))) = view {
            for key in &source.parts[part].backed[&ordinal] {
                live.insert((key.clone(), id));
            }
            if let Some(&(best, at)) = viewed.get(&((part, ordinal), id)) {
                let claimant = &mut ordered[at].1;
                claimant.standing = claimant.standing.max(standing);
                if rank < best {
                    if let Holder::Pilcrow(previous) = claimant.holder {
                        claimant.views.push(previous);
                    }
                    claimant.holder = Holder::Pilcrow(index);
                    claimant.owner = pilcrow.key.clone();
                    viewed.insert(((part, ordinal), id), (rank, at));
                } else {
                    claimant.views.push(index);
                }
                continue;
            }
            viewed.insert(((part, ordinal), id), (rank, ordered.len()));
        }
        let order = match view {
            Some((_, ((part, ordinal), _))) => (part, ordinal, 0),
            None => (usize::MAX, 0, index),
        };
        ordered.push((
            order,
            Claimant {
                holder: Holder::Pilcrow(index),
                id,
                text,
                owner: rename.unwrap_or(&pilcrow.key).to_owned(),
                standing,
                views: Vec::new(),
            },
        ));
    }
    if let Some(source) = source {
        for (part_index, occurrence) in source.unbacked() {
            let part = &source.parts[part_index];
            let owner = occurrence_owner(part, occurrence.ordinal);
            let (text, standing) = match assignments.get(&owner) {
                Some(id) => (id.clone(), claim_standing(claims, id, &owner)),
                None => match valid(occurrence.para_id.clone()) {
                    Some(id) => (id, Standing::Source),
                    None => continue,
                },
            };
            let Some(id) = parse_paragraph_id(&text) else {
                continue;
            };
            live.insert((owner.clone(), id));
            ordered.push((
                (part_index, occurrence.ordinal, 0),
                Claimant {
                    holder: Holder::Occurrence(part_index, occurrence.ordinal),
                    id,
                    text,
                    owner,
                    standing,
                    views: Vec::new(),
                },
            ));
        }
    }
    let mut retired = Vec::new();
    let mut retire = |id: u32, owner: &str, standing: Standing| {
        if !live.contains(&(owner.to_owned(), id)) {
            retired.push((
                (usize::MAX, u32::MAX, usize::MAX),
                Claimant {
                    holder: Holder::Retired,
                    id,
                    text: format_paragraph_id(id),
                    owner: owner.to_owned(),
                    standing,
                    views: Vec::new(),
                },
            ));
        }
    };
    for (id, owners) in claims {
        for claim in owners.iter().filter(|claim| claim.published) {
            retire(*id, &claim.owner, Standing::Published);
        }
    }
    if let Some(source) = source {
        for (key, seed) in &source.seeded {
            if let Some(id) = seed.source_para_id.as_deref().and_then(parse_paragraph_id) {
                retire(id, key, Standing::Source);
            }
        }
    }
    retired.sort_by(|left, right| (left.1.id, &left.1.owner).cmp(&(right.1.id, &right.1.owner)));
    ordered.sort_by_key(|(order, _)| *order);
    ordered.extend(retired);
    let mut groups: BTreeMap<u32, Vec<Claimant>> = BTreeMap::new();
    for (_, claimant) in ordered {
        groups.entry(claimant.id).or_default().push(claimant);
    }
    groups
        .into_values()
        .filter(|group| {
            group.len() > 1
                && group
                    .iter()
                    .any(|claimant| claimant.holder != Holder::Retired)
        })
        .collect()
}

/// Which claimants of one ID take new IDs, and whether saved claims conflict.
/// Every claimant below the strongest standing loses. Among the strongest,
/// source duplicates are repaired only when `persist` asks, the first in
/// package order keeping the ID; unpublished claims and copies keep it for
/// the smallest owner; published claims all keep it and conflict.
fn losers(group: &[Claimant], persist: bool) -> (Vec<usize>, bool) {
    let best = group
        .iter()
        .map(|claimant| claimant.standing)
        .max()
        .unwrap_or(Standing::Unowned);
    let top: Vec<usize> = (0..group.len())
        .filter(|index| group[*index].standing == best)
        .collect();
    let mut losers: Vec<usize> = (0..group.len())
        .filter(|index| group[*index].standing != best)
        .collect();
    let keeper = match best {
        Standing::Source if persist => Some(top[0]),
        Standing::Unpublished | Standing::Unowned => top
            .iter()
            .copied()
            .min_by(|left, right| group[*left].owner.cmp(&group[*right].owner)),
        Standing::Source | Standing::Published => None,
    };
    if let Some(keeper) = keeper {
        losers.extend(top.iter().copied().filter(|index| *index != keeper));
    }
    losers.sort_unstable();
    (losers, best == Standing::Published && top.len() > 1)
}

/// Flags identity-relevant changes while an update integrates.
pub(crate) struct IdentityWatch {
    changed: Arc<AtomicBool>,
    _subscriptions: Vec<Subscription>,
}

impl IdentityWatch {
    pub(crate) fn new(doc: &EditingDoc) -> Self {
        let changed = Arc::new(AtomicBool::new(false));
        let txn = doc.yrs_doc().transact();
        let (stories, assignments) = (txn.get_map(STORIES), txn.get_map(SOURCE_PARAGRAPH_IDS));
        drop(txn);
        let mut subscriptions = Vec::new();
        if let Some(stories) = stories {
            let flag = Arc::clone(&changed);
            subscriptions.push(stories.observe_deep(move |txn, events| {
                let relevant = events.iter().any(|event| match event {
                    Event::Map(event) => event.keys(txn).iter().any(|(key, change)| {
                        matches!(key.as_ref(), PARA_ID | OOXML_PARA_ID | SOURCE_PARA_ID)
                            || matches!(change, EntryChange::Inserted(Out::YText(_)))
                    }),
                    Event::Text(event) => event
                        .delta(txn)
                        .iter()
                        .any(|delta| matches!(delta, Delta::Inserted(Out::YMap(_), _))),
                    _ => false,
                });
                if relevant {
                    flag.store(true, Ordering::Relaxed);
                }
            }));
        }
        if let Some(assignments) = assignments {
            let flag = Arc::clone(&changed);
            subscriptions.push(assignments.observe(move |_, _| {
                flag.store(true, Ordering::Relaxed);
            }));
        }
        Self {
            changed,
            _subscriptions: subscriptions,
        }
    }

    pub(crate) fn changed(&self) -> bool {
        self.changed.load(Ordering::Relaxed)
    }
}

struct Reassignment {
    holder: Holder,
    previous: String,
    /// The pilcrow carries a source ID that is not its own: a copy's.
    copied_source: bool,
    /// Other stories' views of the holder, which take its new ID too.
    views: Vec<usize>,
}

struct Repairs {
    /// `(pilcrow, previous key, new key)`.
    renamed: Vec<(usize, String, String)>,
    reassigned: Vec<Reassignment>,
    conflicts: Vec<(String, Vec<Holder>)>,
}

impl Repairs {
    fn is_empty(&self) -> bool {
        self.renamed.is_empty() && self.reassigned.is_empty()
    }
}

impl EditingDoc {
    /// Starts a new opening of this document: writes its generation into
    /// replicated state, identical on every replica that syncs it, so session
    /// anchors from any other opening never resolve here, even one seeded
    /// alike by the same client. Every path that establishes a document
    /// session calls it; `generation` fixes the value, as a deterministic
    /// shared seed needs, and is otherwise freshly minted. System-origin and
    /// outside undo.
    pub fn begin_opening(&self, generation: Option<&str>) {
        let generation = generation.map_or_else(fresh_generation, str::to_owned);
        let mut txn = self.transact_for(&EditCtx::system(""));
        txn.get_map(SESSION)
            .expect("session root is declared by EditingDoc::new")
            .insert(&mut txn, GENERATION, generation);
    }

    /// Whether an opening has written this document's generation.
    #[cfg(feature = "wasm")]
    pub(crate) fn has_opening(&self) -> bool {
        let txn = self.yrs_doc().transact();
        txn.get_map(SESSION)
            .is_some_and(|session| session.contains_key(&txn, GENERATION))
    }

    /// The collaborative document session: the opening generation (see
    /// [`Self::begin_opening`]) and the CRDT identity of the body story. Both
    /// are replicated, so every replica of one opening shares it, and a fresh
    /// opening or a replaced body is a new session.
    pub fn session_id(&self) -> String {
        let txn = self.yrs_doc().transact();
        let generation = txn
            .get_map(SESSION)
            .and_then(|session| map_string(&session, &txn, GENERATION));
        let story = txn.get_map(STORIES).and_then(|stories| {
            stories.get(&txn, "body").or_else(|| {
                let mut names: Vec<String> = stories.keys(&txn).map(str::to_owned).collect();
                names.sort();
                names.first().and_then(|name| stories.get(&txn, name))
            })
        });
        let story = match story {
            Some(Out::YText(text)) => match AsRef::<Branch>::as_ref(&text).id() {
                BranchID::Nested(id) => format!("{}.{}", id.client, id.clock),
                BranchID::Root(name) => name.to_string(),
            },
            _ => String::new(),
        };
        match generation {
            Some(generation) => format!("{generation}.{story}"),
            None => story,
        }
    }

    /// Plans the repairs of duplicated identities. A pilcrow `copy` accepts
    /// holds none of the identity it carries: it never keeps a duplicate, nor
    /// a key or ID a paragraph since deleted held, whatever key it names.
    fn plan_repairs<T: ReadTxn>(
        &self,
        txn: &T,
        persist: bool,
        copy: &dyn Fn(&MapRef) -> bool,
    ) -> (Scan, Repairs) {
        let source = self.source_index();
        let scan = Scan::new(txn);
        let claims = claims(txn);
        let copies: HashSet<usize> = (0..scan.pilcrows.len())
            .filter(|index| copy(&scan.pilcrows[*index].map))
            .collect();
        let standing = |index: usize| {
            if copies.contains(&index) {
                Standing::Unowned
            } else {
                pilcrow_standing(&scan.pilcrows[index], source.as_deref(), &claims)
            }
        };
        let claimed: HashSet<&str> = claims
            .values()
            .flatten()
            .map(|claim| claim.owner.as_str())
            .collect();
        let reserved = |key: &str| {
            claimed.contains(key)
                || source
                    .as_deref()
                    .is_some_and(|source| source.seeded.contains_key(key))
                || self.with_seen(txn, |seen| seen.keys.contains(key))
        };
        let mut holders: BTreeMap<&str, Vec<usize>> = BTreeMap::new();
        for (index, pilcrow) in scan.pilcrows.iter().enumerate() {
            holders.entry(pilcrow.key.as_str()).or_default().push(index);
        }
        let mut renamed = Vec::new();
        for (key, indexes) in &holders {
            if indexes.len() == 1 && !copies.contains(&indexes[0]) {
                continue;
            }
            let keeper = indexes
                .iter()
                .copied()
                .filter(|index| !copies.contains(index) || !reserved(key))
                .max_by_key(|index| {
                    (
                        !copies.contains(index),
                        standing(*index),
                        std::cmp::Reverse(*index),
                    )
                });
            for index in indexes
                .iter()
                .copied()
                .filter(|index| Some(*index) != keeper)
            {
                let pilcrow = &scan.pilcrows[index];
                renamed.push((index, pilcrow.key.clone(), derived_key(&pilcrow.map)));
            }
        }
        renamed.sort_by_key(|(index, _, _)| *index);
        let renamed_keys: HashMap<usize, &str> = renamed
            .iter()
            .map(|(index, _, key)| (*index, key.as_str()))
            .collect();
        let mut repairs = Repairs {
            renamed: Vec::new(),
            reassigned: Vec::new(),
            conflicts: Vec::new(),
        };
        let groups = collisions(
            &scan,
            source.as_deref(),
            &claims,
            &source_assignments(txn),
            &copies,
            &renamed_keys,
        );
        for group in groups {
            let (losers, conflict) = losers(&group, persist);
            let saved_live = group.iter().any(|claimant| {
                claimant.standing == Standing::Published && claimant.holder != Holder::Retired
            });
            if conflict && saved_live {
                repairs.conflicts.push((
                    format_paragraph_id(group[0].id),
                    group
                        .iter()
                        .filter(|claimant| claimant.standing == Standing::Published)
                        .map(|claimant| claimant.holder)
                        .collect(),
                ));
            }
            for loser in losers {
                let claimant = &group[loser];
                if claimant.holder == Holder::Retired {
                    continue;
                }
                let copied_source = match claimant.holder {
                    Holder::Pilcrow(index) => {
                        let pilcrow = &scan.pilcrows[index];
                        pilcrow.source.is_some()
                            && (renamed_keys.contains_key(&index)
                                || copies.contains(&index)
                                || source.as_deref().is_some_and(|source| {
                                    !source.seeded.contains_key(&pilcrow.key)
                                }))
                    }
                    Holder::Occurrence(..) | Holder::Retired => false,
                };
                repairs.reassigned.push(Reassignment {
                    holder: claimant.holder,
                    previous: claimant.text.clone(),
                    copied_source,
                    views: claimant.views.clone(),
                });
            }
        }
        repairs.renamed = renamed;
        (scan, repairs)
    }

    /// Deterministically restores unique session keys and Word paragraph IDs
    /// after divergent replicas merged; see [`Self::persist_paragraph_ids`]
    /// for the ownership rules. Duplicated source IDs and conflicting saved
    /// claims are left for an explicit persistence. Writes nothing when
    /// nothing collides. Returns the `(previous, new)` session keys.
    pub(crate) fn repair_paragraph_identities(&self) -> Vec<(ParagraphId, ParagraphId)> {
        let (_, plan) = self.plan_repairs(&self.yrs_doc().transact(), false, &|_| false);
        if plan.is_empty() {
            return Vec::new();
        }
        let mut txn = self.transact_for(&EditCtx::system(""));
        let (scan, plan) = self.plan_repairs(&txn, false, &|_| false);
        let mut allocator = IdAllocator::new(self, &txn);
        self.apply_repairs(&mut txn, &scan, &mut allocator, &plan);
        plan.renamed
            .into_iter()
            .map(|(_, previous, key)| (previous, key))
            .collect()
    }

    /// Repairs, within the transaction that made them, the duplicates that
    /// raw ops introduced: a pilcrow they inserted (created by this replica
    /// from clock `since` on) or re-keyed is the copy, and the paragraph it
    /// copied keeps its identity. The identity each copy ends with is then
    /// its own: its ID is claimed for its key, so a save publishes it and a
    /// deletion leaves it reserved, and its key is never allocated again.
    pub(crate) fn repair_copies(
        &self,
        txn: &mut TransactionMut<'_>,
        since: u32,
        rekeyed: &[MapRef],
    ) {
        let client = yrs::ClientID::new(self.client_id());
        let copy = |map: &MapRef| {
            rekeyed.contains(map)
                || matches!(
                    AsRef::<Branch>::as_ref(map).id(),
                    BranchID::Nested(id) if id.client == client && id.clock >= since
                )
        };
        let (scan, plan) = self.plan_repairs(&*txn, false, &copy);
        if !plan.is_empty() {
            let mut allocator = IdAllocator::new(self, &*txn);
            self.apply_repairs(txn, &scan, &mut allocator, &plan);
        }
        let keys: HashMap<usize, &str> = plan
            .renamed
            .iter()
            .map(|(index, _, key)| (*index, key.as_str()))
            .collect();
        let reassigned: HashSet<usize> = plan
            .reassigned
            .iter()
            .filter_map(|reassignment| match reassignment.holder {
                Holder::Pilcrow(index) => Some(index),
                Holder::Occurrence(..) | Holder::Retired => None,
            })
            .collect();
        let claims = claims(&*txn);
        for (index, pilcrow) in scan.pilcrows.iter().enumerate() {
            if !copy(&pilcrow.map) {
                continue;
            }
            let key = keys.get(&index).copied().unwrap_or(&pilcrow.key);
            self.with_seen(&*txn, |seen| seen.keys.insert(key.to_owned()));
            let id = pilcrow.allocated.as_deref().and_then(parse_paragraph_id);
            if let Some(id) = id.filter(|_| !reassigned.contains(&index))
                && claim(&claims, &format_paragraph_id(id), key).is_none()
            {
                record_claim(
                    txn,
                    &format_paragraph_id(id),
                    key,
                    ParagraphIdOrigin::Authored,
                    false,
                );
            }
        }
    }

    fn source_story_of(
        &self,
        scan: &Scan,
        source: Option<&SourceIndex>,
        holder: Holder,
    ) -> Option<SourceStory> {
        let source = source?;
        match holder {
            Holder::Pilcrow(index) => source
                .roots
                .get(scan.root(&scan.pilcrows[index].story))
                .cloned(),
            Holder::Occurrence(part, ordinal) => {
                let part = &source.parts[part];
                Some(part.story(part.occurrence(ordinal)?.item_id.as_deref()))
            }
            Holder::Retired => None,
        }
    }

    fn apply_repairs(
        &self,
        txn: &mut TransactionMut<'_>,
        scan: &Scan,
        allocator: &mut IdAllocator,
        plan: &Repairs,
    ) -> Vec<ParagraphIdAssignment> {
        let source = self.source_index();
        let mut keys: BTreeMap<usize, (&str, &str)> = BTreeMap::new();
        for (index, previous, key) in &plan.renamed {
            scan.pilcrows[*index].map.insert(txn, PARA_ID, key.as_str());
            keys.insert(*index, (previous.as_str(), key.as_str()));
        }
        let mut assignments = Vec::new();
        for reassignment in &plan.reassigned {
            let source_story = self.source_story_of(scan, source.as_deref(), reassignment.holder);
            match reassignment.holder {
                Holder::Pilcrow(index) => {
                    let pilcrow = &scan.pilcrows[index];
                    if reassignment.copied_source {
                        pilcrow.map.remove(txn, SOURCE_PARA_ID);
                    }
                    let (replaced, key) = match keys.remove(&index) {
                        Some((previous, key)) => (Some(previous.to_owned()), key.to_owned()),
                        None => (None, pilcrow.key.clone()),
                    };
                    let id = allocator.bind(txn, &pilcrow.map, &key, ParagraphIdOrigin::Repaired);
                    assignments.push(ParagraphIdAssignment {
                        paragraph: ParagraphRef::Session {
                            story: pilcrow.story.clone(),
                            para_id: key,
                        },
                        replaced_para_id: replaced,
                        previous_ooxml_para_id: Some(reassignment.previous.clone()),
                        ooxml_para_id: id.clone(),
                        origin: ParagraphIdOrigin::Repaired,
                        source_story: source_story.clone(),
                    });
                    for view in &reassignment.views {
                        let view = &scan.pilcrows[*view];
                        attach(txn, &view.map, &view.key, &id, ParagraphIdOrigin::Repaired);
                        assignments.push(ParagraphIdAssignment {
                            paragraph: view.session_ref(),
                            replaced_para_id: None,
                            previous_ooxml_para_id: Some(reassignment.previous.clone()),
                            ooxml_para_id: id.clone(),
                            origin: ParagraphIdOrigin::Repaired,
                            source_story: source_story.clone(),
                        });
                    }
                }
                Holder::Retired => {}
                Holder::Occurrence(part, ordinal) => {
                    let Some(source) = source.as_deref() else {
                        continue;
                    };
                    let part = &source.parts[part];
                    let id = allocator.assign_source(
                        txn,
                        &occurrence_owner(part, ordinal),
                        ParagraphIdOrigin::Repaired,
                    );
                    assignments.push(ParagraphIdAssignment {
                        paragraph: ParagraphRef::Source(source.reference(part, ordinal)),
                        replaced_para_id: None,
                        previous_ooxml_para_id: Some(reassignment.previous.clone()),
                        ooxml_para_id: id,
                        origin: ParagraphIdOrigin::Repaired,
                        source_story,
                    });
                }
            }
        }
        for (index, (previous, key)) in keys {
            let pilcrow = &scan.pilcrows[index];
            if let Some((id, _)) = pilcrow.saved_id(source.as_deref()) {
                assignments.push(ParagraphIdAssignment {
                    paragraph: ParagraphRef::Session {
                        story: pilcrow.story.clone(),
                        para_id: key.to_owned(),
                    },
                    replaced_para_id: Some(previous.to_owned()),
                    previous_ooxml_para_id: None,
                    ooxml_para_id: id,
                    origin: ParagraphIdOrigin::Repaired,
                    source_story: self.source_story_of(
                        scan,
                        source.as_deref(),
                        Holder::Pilcrow(index),
                    ),
                });
            }
        }
        assignments
    }

    /// Gives every paragraph that saves without a Word paragraph ID a fresh
    /// one, across the session stories and the retained package's source
    /// paragraphs outside them; an editor-only paragraph gets one only once
    /// it holds content. Repairs duplicates by ownership: source IDs and
    /// saved claims keep their IDs over unsaved claims and copies, and the
    /// first occurrence in package order keeps a duplicated source ID. Saved
    /// claims that conflict stay and are reported. A refusal leaves the
    /// document unchanged. The change is replicated, system-origin and
    /// outside undo, and later saves keep it.
    pub fn persist_paragraph_ids(&self) -> Result<PersistedParagraphIds, ParagraphIdRefusal> {
        let source = self.source_index();
        let mut txn = self.transact_for(&EditCtx::system(""));
        let (scan, plan) = self.plan_repairs(&txn, true, &|_| false);
        if let Some(source) = source.as_deref() {
            for reassignment in &plan.reassigned {
                let Holder::Occurrence(part, _) = reassignment.holder else {
                    continue;
                };
                let part = &source.parts[part];
                let previous = parse_paragraph_id(&reassignment.previous);
                if part.kind != SourceStoryKind::Comment
                    || !previous.is_some_and(|id| source.comment_references.contains(&id))
                {
                    continue;
                }
                let holders: Vec<&ParagraphOccurrence> = part
                    .occurrences
                    .iter()
                    .filter(|occurrence| {
                        occurrence.para_id.as_deref().and_then(parse_paragraph_id) == previous
                    })
                    .collect();
                if holders.len() > 1 {
                    let mut comment_ids: Vec<String> = holders
                        .iter()
                        .filter_map(|occurrence| occurrence.item_id.clone())
                        .collect();
                    comment_ids.dedup();
                    return Err(ParagraphIdRefusal::AmbiguousCommentReference {
                        ooxml_para_id: reassignment.previous.clone(),
                        comment_ids,
                    });
                }
            }
        }
        let mut allocator = IdAllocator::new(self, &txn);
        let mut report = PersistedParagraphIds {
            assignments: self.apply_repairs(&mut txn, &scan, &mut allocator, &plan),
            diagnostics: Vec::new(),
        };
        for (id, holders) in &plan.conflicts {
            report
                .diagnostics
                .push(ParagraphIdDiagnostic::ConflictingSavedIds {
                    ooxml_para_id: id.clone(),
                    paragraphs: holders
                        .iter()
                        .filter_map(|holder| holder_ref(&scan, source.as_deref(), *holder))
                        .collect(),
                });
        }
        let handled: HashSet<usize> = plan
            .reassigned
            .iter()
            .filter_map(|reassignment| match reassignment.holder {
                Holder::Pilcrow(index) => Some(index),
                Holder::Occurrence(..) | Holder::Retired => None,
            })
            .chain(plan.renamed.iter().map(|(index, _, _)| *index))
            .collect();
        let mut shared: HashMap<(usize, u32), String> = HashMap::new();
        for (index, pilcrow) in scan.pilcrows.iter().enumerate() {
            if (pilcrow.synthetic && !pilcrow.content)
                || handled.contains(&index)
                || pilcrow.saved_id(source.as_deref()).is_some()
            {
                continue;
            }
            if pilcrow.synthetic {
                pilcrow.map.remove(&mut txn, PARA_ORIGIN);
            }
            let occurrence = source
                .as_deref()
                .and_then(|source| source.view_of(&pilcrow.key))
                .map(|(occurrence, _)| occurrence);
            let id = match occurrence.and_then(|occurrence| shared.get(&occurrence)) {
                Some(id) => {
                    attach(
                        &mut txn,
                        &pilcrow.map,
                        &pilcrow.key,
                        id,
                        ParagraphIdOrigin::Persisted,
                    );
                    id.clone()
                }
                None => allocator.bind(
                    &mut txn,
                    &pilcrow.map,
                    &pilcrow.key,
                    ParagraphIdOrigin::Persisted,
                ),
            };
            if let Some(occurrence) = occurrence {
                shared.entry(occurrence).or_insert_with(|| id.clone());
            }
            report.assignments.push(ParagraphIdAssignment {
                paragraph: pilcrow.session_ref(),
                replaced_para_id: None,
                previous_ooxml_para_id: None,
                ooxml_para_id: id,
                origin: ParagraphIdOrigin::Persisted,
                source_story: self.source_story_of(
                    &scan,
                    source.as_deref(),
                    Holder::Pilcrow(index),
                ),
            });
        }
        match source.as_deref() {
            Some(source) => {
                let assigned = source_assignments(&txn);
                for (part_index, occurrence) in source.unbacked() {
                    let part = &source.parts[part_index];
                    let owner = occurrence_owner(part, occurrence.ordinal);
                    if assigned.contains_key(&owner) || valid(occurrence.para_id.clone()).is_some()
                    {
                        continue;
                    }
                    let id =
                        allocator.assign_source(&mut txn, &owner, ParagraphIdOrigin::Persisted);
                    report.assignments.push(ParagraphIdAssignment {
                        paragraph: ParagraphRef::Source(source.reference(part, occurrence.ordinal)),
                        replaced_para_id: None,
                        previous_ooxml_para_id: None,
                        ooxml_para_id: id,
                        origin: ParagraphIdOrigin::Persisted,
                        source_story: Some(part.story(occurrence.item_id.as_deref())),
                    });
                }
            }
            None => report
                .diagnostics
                .push(ParagraphIdDiagnostic::NoSourcePackage),
        }
        Ok(report)
    }

    /// Reconciles what a save wrote with the live state and publishes it:
    /// `saved` holds the `(owner, ID)` pairs the save captured, an owner being
    /// a session key or a source occurrence's `{partUri}#{ordinal}`. A pair
    /// whose owner now saves with another ID, or whose ID another paragraph
    /// took since the capture, is stale: it is returned and not published.
    /// The rest are marked published, so they keep their IDs against unsaved
    /// claims from other replicas. Writes nothing when nothing changes.
    pub fn record_saved_paragraph_ids(&self, saved: &[(String, String)]) -> Vec<(String, String)> {
        let source = self.source_index();
        let (stale, unpublished) = {
            let txn = self.yrs_doc().transact();
            let claims = claims(&txn);
            let assignments = source_assignments(&txn);
            let mut current: HashMap<String, u32> = HashMap::new();
            for pilcrow in Scan::new(&txn).pilcrows {
                if let Some(id) = pilcrow
                    .saved_id(source.as_deref())
                    .and_then(|(id, _)| parse_paragraph_id(&id))
                {
                    current.insert(pilcrow.key, id);
                }
            }
            if let Some(source) = source.as_deref() {
                for (part_index, occurrence) in source.unbacked() {
                    let owner = occurrence_owner(&source.parts[part_index], occurrence.ordinal);
                    let id = assignments
                        .get(&owner)
                        .cloned()
                        .or_else(|| valid(occurrence.para_id.clone()));
                    if let Some(id) = id.as_deref().and_then(parse_paragraph_id) {
                        current.insert(owner, id);
                    }
                }
            }
            let captured: HashSet<(&str, u32)> = saved
                .iter()
                .filter_map(|(owner, id)| Some((owner.as_str(), parse_paragraph_id(id)?)))
                .collect();
            let occurrence = |key: &str| {
                source
                    .as_deref()
                    .and_then(|source| source.view_of(key))
                    .map(|(occurrence, _)| occurrence)
            };
            let mut stale = Vec::new();
            let mut unpublished = Vec::new();
            for (owner, text) in saved {
                let Some(id) = parse_paragraph_id(text) else {
                    continue;
                };
                let moved = current.get(owner).is_some_and(|now| *now != id);
                let viewed = occurrence(owner);
                let taken = current.iter().any(|(holder, now)| {
                    *now == id
                        && !captured.contains(&(holder.as_str(), id))
                        && (viewed.is_none() || occurrence(holder) != viewed)
                });
                if moved || taken {
                    stale.push((owner.clone(), text.clone()));
                } else if let Some(found) =
                    claim(&claims, text, owner).filter(|claim| !claim.published)
                {
                    unpublished.push((format_paragraph_id(id), owner.as_str(), found.origin));
                }
            }
            (stale, unpublished)
        };
        if !unpublished.is_empty() {
            let mut txn = self.transact_for(&EditCtx::system(""));
            for (id, owner, origin) in unpublished {
                record_claim(&mut txn, &id, owner, origin, true);
            }
        }
        stale
    }

    /// The Word paragraph ID each paragraph of `story_id` saves with, in
    /// document order.
    pub fn story_paragraph_ids(&self, story_id: &str) -> EditResult<Vec<Option<String>>> {
        let source = self.source_index();
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, story_id)?;
        Ok(pilcrows(&story, &txn)
            .into_iter()
            .map(|(_, map)| {
                Pilcrow::read(story_id, map, &txn)
                    .saved_id(source.as_deref())
                    .map(|(id, _)| id)
            })
            .collect())
    }

    /// Every paragraph's session key, Word paragraph ID and anchors. Reads only.
    pub fn paragraph_identities(&self) -> ParagraphIdentities {
        let source = self.source_index();
        let txn = self.yrs_doc().transact();
        let claims = claims(&txn);
        let assignments = source_assignments(&txn);
        let scan = Scan::new(&txn);
        let id_origin = |id: &str, owner: &str| {
            claim(&claims, id, owner).map_or(ParagraphIdOrigin::Source, |claim| claim.origin)
        };
        let mut paragraphs: Vec<ParagraphIdentity> = scan
            .pilcrows
            .iter()
            .map(|pilcrow| {
                let saved = pilcrow.saved_id(source.as_deref());
                let seed = source
                    .as_deref()
                    .and_then(|source| source.seeded.get(&pilcrow.key));
                let root = scan.root(&pilcrow.story);
                ParagraphIdentity {
                    paragraph: pilcrow.session_ref(),
                    origin: if pilcrow.synthetic {
                        ParagraphOrigin::Synthetic
                    } else if seed.is_some() || (source.is_none() && pilcrow.source.is_some()) {
                        ParagraphOrigin::Source
                    } else {
                        ParagraphOrigin::Authored
                    },
                    id_origin: saved.as_ref().map(|(id, from_source)| {
                        if *from_source {
                            ParagraphIdOrigin::Source
                        } else {
                            id_origin(id, &pilcrow.key)
                        }
                    }),
                    ooxml_para_id: saved.map(|(id, _)| id),
                    source_story: source
                        .as_deref()
                        .and_then(|source| source.roots.get(root))
                        .cloned(),
                    source: source.as_deref().and_then(|source| {
                        let seed = seed.filter(|seed| seed.root == root)?;
                        Some(source.reference(&source.parts[seed.part?], seed.ordinal?))
                    }),
                }
            })
            .collect();
        if let Some(source) = source.as_deref() {
            for (part_index, occurrence) in source.unbacked() {
                let part = &source.parts[part_index];
                let owner = occurrence_owner(part, occurrence.ordinal);
                let (id, origin) = match assignments.get(&owner) {
                    Some(id) => (Some(id.clone()), Some(id_origin(id, &owner))),
                    None => {
                        let id = valid(occurrence.para_id.clone());
                        let origin = id.as_ref().map(|_| ParagraphIdOrigin::Source);
                        (id, origin)
                    }
                };
                let reference = source.reference(part, occurrence.ordinal);
                paragraphs.push(ParagraphIdentity {
                    paragraph: ParagraphRef::Source(reference.clone()),
                    origin: ParagraphOrigin::Source,
                    ooxml_para_id: id,
                    id_origin: origin,
                    source_story: Some(part.story(occurrence.item_id.as_deref())),
                    source: Some(reference),
                });
            }
        }
        drop(txn);
        ParagraphIdentities {
            session_id: self.session_id(),
            package_sha256: source.map(|source| source.package_sha256.clone()),
            paragraphs,
        }
    }

    /// Resolves an anchor to the paragraph that currently holds it: a session
    /// anchor by its session, story and key; a source anchor by its package,
    /// part and occurrence; a persisted anchor by its Word paragraph ID in its
    /// source story, nested tables and content controls included. Reads only.
    pub fn resolve_paragraph_anchor(&self, anchor: &ParagraphAnchor) -> AnchorResolution {
        if let ParagraphAnchor::Session { session_id, .. } = anchor
            && *session_id != self.session_id()
        {
            return AnchorResolution::Unsupported(AnchorUnsupported::ForeignSession);
        }
        let source = self.source_index();
        let txn = self.yrs_doc().transact();
        let scan = Scan::new(&txn);
        let by_key = |key: &str, story: Option<&str>| -> Vec<ParagraphRef> {
            scan.pilcrows
                .iter()
                .filter(|pilcrow| {
                    pilcrow.key == key && story.is_none_or(|story| pilcrow.story == story)
                })
                .map(Pilcrow::session_ref)
                .collect()
        };
        let matches = match anchor {
            ParagraphAnchor::Session { story, para_id, .. } => by_key(para_id, Some(story)),
            ParagraphAnchor::Source(reference) => {
                let Some(source) = source.as_deref() else {
                    return AnchorResolution::Unsupported(AnchorUnsupported::NoSourcePackage);
                };
                if source.package_sha256 != reference.package_sha256 {
                    return AnchorResolution::Unsupported(AnchorUnsupported::ForeignPackage);
                }
                match source.part(&reference.part_uri) {
                    Some(part) => match part.backed.get(&reference.paragraph_ordinal) {
                        Some(views) => views
                            .iter()
                            .map(|key| by_key(key, None))
                            .find(|matches| !matches.is_empty())
                            .unwrap_or_default(),
                        None if part.occurrence(reference.paragraph_ordinal).is_some() => {
                            vec![ParagraphRef::Source(reference.clone())]
                        }
                        None => Vec::new(),
                    },
                    None => Vec::new(),
                }
            }
            ParagraphAnchor::Persisted { story, para_id } => {
                let Some(source) = source.as_deref() else {
                    return AnchorResolution::Unsupported(AnchorUnsupported::NoSourcePackage);
                };
                let Some(target) = parse_paragraph_id(para_id) else {
                    return AnchorResolution::Missing;
                };
                let mut matches: Vec<ParagraphRef> = source
                    .first_views(scan.pilcrows.iter().filter(|pilcrow| {
                        source.roots.get(scan.root(&pilcrow.story)) == Some(story)
                            && pilcrow
                                .saved_id(Some(source))
                                .and_then(|(id, _)| parse_paragraph_id(&id))
                                == Some(target)
                    }))
                    .into_iter()
                    .map(Pilcrow::session_ref)
                    .collect();
                let assignments = source_assignments(&txn);
                for (part_index, occurrence) in source.unbacked() {
                    let part = &source.parts[part_index];
                    if part.story(occurrence.item_id.as_deref()) != *story {
                        continue;
                    }
                    let id = assignments
                        .get(&occurrence_owner(part, occurrence.ordinal))
                        .cloned()
                        .or_else(|| valid(occurrence.para_id.clone()));
                    if id.as_deref().and_then(parse_paragraph_id) == Some(target) {
                        matches.push(ParagraphRef::Source(
                            source.reference(part, occurrence.ordinal),
                        ));
                    }
                }
                matches
            }
        };
        match <[ParagraphRef; 1]>::try_from(matches) {
            Ok([paragraph]) => AnchorResolution::Found(paragraph),
            Err(matches) if matches.is_empty() => AnchorResolution::Missing,
            Err(matches) => AnchorResolution::Ambiguous(matches),
        }
    }

    /// The paragraph IDs a save of this session applies; see
    /// [`ParagraphSavePlan`]. A part is patched in place only when every
    /// story seeded from it, nested ones included, gives the save projection
    /// exactly what it gave as seeded (see [`StoryState`]), so that its source
    /// bytes are what a save would write, IDs aside.
    pub fn paragraph_save_plan(&self) -> ParagraphSavePlan {
        let Some(source) = self.source_index() else {
            return ParagraphSavePlan::default();
        };
        let txn = self.yrs_doc().transact();
        let assignments = source_assignments(&txn);
        let scan = Scan::new(&txn);
        drop(txn);
        let mut plan = ParagraphSavePlan {
            assignments: source
                .unbacked()
                .filter_map(|(part_index, occurrence)| {
                    let part = &source.parts[part_index];
                    let id = assignments.get(&occurrence_owner(part, occurrence.ordinal))?;
                    Some((part.path().to_owned(), occurrence.ordinal, id.clone()))
                })
                .collect(),
            patched_parts: Vec::new(),
        };
        let seeded = source.seed_states();
        let live = story_states(self);
        let stories_of = |states: &HashMap<String, StoryState>, part: &SourcePart| {
            let prefix = match part.kind {
                SourceStoryKind::Footnote => Some("fn:"),
                SourceStoryKind::Endnote => Some("en:"),
                _ => None,
            };
            states
                .iter()
                .filter(|(_, state)| {
                    part.roots.contains(&state.root)
                        || prefix.is_some_and(|prefix| state.root.starts_with(prefix))
                })
                .map(|(story, state)| (story.clone(), (state.fingerprint, state.comments.clone())))
                .collect::<BTreeMap<_, _>>()
        };
        let by_key: HashMap<&str, &Pilcrow> = scan
            .pilcrows
            .iter()
            .map(|pilcrow| (pilcrow.key.as_str(), pilcrow))
            .collect();
        'parts: for part in &source.parts {
            if seeded.is_empty() || stories_of(seeded, part) != stories_of(&live, part) {
                continue;
            }
            let mut patches = Vec::new();
            for occurrence in &part.occurrences {
                let desired = match part.backed.get(&occurrence.ordinal) {
                    Some(views) => match views.iter().find_map(|key| by_key.get(key.as_str())) {
                        Some(pilcrow) => pilcrow.saved_id(Some(&source)).map(|(id, _)| id),
                        None => continue 'parts,
                    },
                    None => assignments
                        .get(&occurrence_owner(part, occurrence.ordinal))
                        .cloned(),
                };
                if let Some(desired) = desired
                    && occurrence.para_id.as_deref() != Some(desired.as_str())
                {
                    patches.push((occurrence.ordinal, desired));
                }
            }
            plan.patched_parts.push((part.path().to_owned(), patches));
        }
        plan
    }
}

fn holder_ref(scan: &Scan, source: Option<&SourceIndex>, holder: Holder) -> Option<ParagraphRef> {
    match holder {
        Holder::Pilcrow(index) => Some(scan.pilcrows[index].session_ref()),
        Holder::Occurrence(part, ordinal) => {
            let source = source?;
            Some(ParagraphRef::Source(
                source.reference(&source.parts[part], ordinal),
            ))
        }
        Holder::Retired => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{FormatPolicy, Position, StoryRange};

    const DATE: &str = "2026-09-24T00:00:00Z";

    fn ctx() -> EditCtx {
        EditCtx::local("", DATE)
    }

    fn session(story: &str, key: &str) -> ParagraphRef {
        ParagraphRef::Session {
            story: story.into(),
            para_id: key.into(),
        }
    }

    fn ids(doc: &EditingDoc) -> Vec<(ParagraphRef, Option<String>)> {
        doc.paragraph_identities()
            .paragraphs
            .into_iter()
            .map(|identity| (identity.paragraph, identity.ooxml_para_id))
            .collect()
    }

    /// Binds `id` to the paragraph `key` as its own claim, as an allocation would.
    fn claim_id(doc: &EditingDoc, key: &str, id: &str, published: bool) {
        let mut txn = doc.transact_for(&EditCtx::system(""));
        let pilcrow = Scan::new(&txn)
            .pilcrows
            .into_iter()
            .find(|pilcrow| pilcrow.key == key)
            .unwrap();
        record_claim(&mut txn, id, key, ParagraphIdOrigin::Authored, published);
        pilcrow.map.insert(&mut txn, OOXML_PARA_ID, id);
    }

    #[test]
    fn a_split_keeps_the_first_half_id_and_allocates_the_second_half_one() {
        let doc = EditingDoc::new(7);
        let first = doc.create_story("body", "abcd", "Normal", "left").unwrap();
        assert_eq!(doc.paragraph_identities().paragraphs[0].ooxml_para_id, None);
        let first_id = doc.persist_paragraph_ids().unwrap().assignments[0]
            .ooxml_para_id
            .clone();
        let split = doc
            .split_paragraph(&ctx(), Position::new("body", 2), None)
            .unwrap();
        let identities = doc.paragraph_identities().paragraphs;
        assert_eq!(identities[0].paragraph, session("body", &first));
        assert_eq!(
            identities[0].ooxml_para_id.as_deref(),
            Some(first_id.as_str())
        );
        assert_eq!(identities[0].id_origin, Some(ParagraphIdOrigin::Persisted));
        let second = identities[1].ooxml_para_id.clone().unwrap();
        assert_ne!(second, first_id);
        assert_eq!(identities[1].id_origin, Some(ParagraphIdOrigin::Authored));
        assert_eq!(identities[1].origin, ParagraphOrigin::Authored);
        assert_eq!(
            second,
            format_paragraph_id(docx_parse::paragraph_identity::paragraph_id_candidate(
                &split.second_para_id,
                0
            ))
        );
    }

    #[test]
    fn deleted_identities_stay_reserved_and_undo_restores_them() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "abcd", "Normal", "left").unwrap();
        let mut undo = doc.undo_manager();
        let split = doc
            .split_paragraph(&ctx(), Position::new("body", 2), None)
            .unwrap();
        let second = doc.paragraph_identities().paragraphs[1]
            .ooxml_para_id
            .clone()
            .unwrap();
        assert!(undo.undo());
        assert_eq!(doc.paragraph_identities().paragraphs.len(), 1);
        let txn = doc.yrs_doc().transact();
        assert!(
            Seen::scan(&txn)
                .ids
                .contains(&parse_paragraph_id(&second).unwrap())
        );
        drop(txn);
        assert!(undo.redo());
        let restored = &doc.paragraph_identities().paragraphs[1];
        assert_eq!(restored.paragraph, session("body", &split.second_para_id));
        assert_eq!(restored.ooxml_para_id.as_deref(), Some(second.as_str()));
    }

    #[test]
    fn session_keys_skip_counters_a_restored_client_already_used() {
        let first = EditingDoc::new(7);
        first
            .create_story("body", "abcd", "Normal", "left")
            .unwrap();
        first
            .split_paragraph(&ctx(), Position::new("body", 2), None)
            .unwrap();
        let restored = EditingDoc::new(7);
        restored
            .apply_update_v1(&first.encode_state_as_update_v1())
            .unwrap();
        restored
            .insert_text(&ctx(), Position::new("body", 0), "x", FormatPolicy::Plain)
            .unwrap();
        let split = restored
            .split_paragraph(&ctx(), Position::new("body", 1), None)
            .unwrap();
        let keys: Vec<_> = restored
            .paragraphs("body")
            .unwrap()
            .into_iter()
            .map(|paragraph| paragraph.para_id)
            .collect();
        assert_eq!(keys.len(), 3);
        assert_eq!(keys.iter().collect::<HashSet<_>>().len(), 3);
        assert!(keys.contains(&split.second_para_id));
        restored.persist_paragraph_ids().unwrap();
        let ids: HashSet<_> = ids(&restored)
            .into_iter()
            .map(|(_, id)| id.unwrap())
            .collect();
        assert_eq!(ids.len(), 3);
    }

    #[test]
    fn repair_and_persistence_are_idempotent() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "abcd", "Normal", "left").unwrap();
        let before = doc.encode_state_vector_v1();
        assert!(doc.repair_paragraph_identities().is_empty());
        assert_eq!(doc.encode_state_vector_v1(), before);
        let report = doc.persist_paragraph_ids().unwrap();
        assert_eq!(report.assignments.len(), 1);
        assert_eq!(report.assignments[0].origin, ParagraphIdOrigin::Persisted);
        assert_eq!(report.diagnostics, [ParagraphIdDiagnostic::NoSourcePackage]);
        doc.delete_range(&ctx(), StoryRange::new("body", 0, 1))
            .unwrap();
        let persisted = doc.encode_state_vector_v1();
        assert!(doc.persist_paragraph_ids().unwrap().assignments.is_empty());
        assert_eq!(doc.encode_state_vector_v1(), persisted);
    }

    #[test]
    fn ownership_keeps_the_stronger_claim_and_ties_break_by_owner() {
        let claimant = |owner: &str, standing| Claimant {
            holder: Holder::Pilcrow(0),
            id: 1,
            text: "00000001".into(),
            owner: owner.into(),
            standing,
            views: Vec::new(),
        };
        let group = [
            claimant("b", Standing::Unpublished),
            claimant("a", Standing::Unpublished),
            claimant("c", Standing::Unowned),
        ];
        assert_eq!(losers(&group, false), (vec![0, 2], false));
        let group = [
            claimant("copy", Standing::Unowned),
            claimant("seed", Standing::Source),
        ];
        assert_eq!(losers(&group, false), (vec![0], false));
        let group = [
            claimant("first", Standing::Source),
            claimant("second", Standing::Source),
        ];
        assert_eq!(losers(&group, false), (vec![], false));
        assert_eq!(losers(&group, true), (vec![1], false));
        let group = [
            claimant("x", Standing::Published),
            claimant("y", Standing::Published),
            claimant("z", Standing::Unpublished),
        ];
        assert_eq!(losers(&group, true), (vec![2], true));
    }

    #[test]
    fn replicas_keep_the_published_claim_when_ids_collide() {
        let base = EditingDoc::new(1);
        base.create_story("body", "one", "Normal", "left").unwrap();
        base.create_story("note", "two", "Normal", "left").unwrap();
        let update = base.encode_state_as_update_v1();
        let (left, right) = (EditingDoc::new(2), EditingDoc::new(3));
        left.apply_update_v1(&update).unwrap();
        right.apply_update_v1(&update).unwrap();
        let body = base.paragraphs("body").unwrap()[0].para_id.clone();
        let note = base.paragraphs("note").unwrap()[0].para_id.clone();
        claim_id(&left, &body, "0000AAAA", false);
        claim_id(&right, &note, "0000AAAA", true);
        let (from_left, from_right) = (
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1(),
        );
        left.apply_update_v1(&from_right).unwrap();
        right.apply_update_v1(&from_left).unwrap();
        for doc in [&left, &right] {
            doc.apply_update_v1(&left.encode_state_as_update_v1())
                .unwrap();
            doc.apply_update_v1(&right.encode_state_as_update_v1())
                .unwrap();
        }
        assert_eq!(ids(&left), ids(&right));
        let converged: HashMap<_, _> = ids(&left).into_iter().collect();
        assert_eq!(
            converged[&session("note", &note)].as_deref(),
            Some("0000AAAA")
        );
        let body_id = converged[&session("body", &body)].clone().unwrap();
        assert_ne!(body_id, "0000AAAA");
        let identity = doc_identity(&left, &body);
        assert_eq!(identity.id_origin, Some(ParagraphIdOrigin::Repaired));
    }

    fn doc_identity(doc: &EditingDoc, key: &str) -> ParagraphIdentity {
        doc.paragraph_identities()
            .paragraphs
            .into_iter()
            .find(|identity| matches!(&identity.paragraph, ParagraphRef::Session { para_id, .. } if para_id == key))
            .unwrap()
    }

    #[test]
    fn unpublished_collisions_keep_the_smallest_owner_on_every_replica() {
        let base = EditingDoc::new(1);
        base.create_story("a", "one", "Normal", "left").unwrap();
        base.create_story("b", "two", "Normal", "left").unwrap();
        let update = base.encode_state_as_update_v1();
        let (left, right, late) = (EditingDoc::new(2), EditingDoc::new(3), EditingDoc::new(4));
        for doc in [&left, &right, &late] {
            doc.apply_update_v1(&update).unwrap();
        }
        let a = base.paragraphs("a").unwrap()[0].para_id.clone();
        let b = base.paragraphs("b").unwrap()[0].para_id.clone();
        claim_id(&left, &a, "0000BBBB", false);
        claim_id(&right, &b, "0000BBBB", false);
        let (from_left, from_right) = (
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1(),
        );
        late.apply_update_v1(&from_left).unwrap();
        left.apply_update_v1(&from_right).unwrap();
        right.apply_update_v1(&from_left).unwrap();
        late.apply_update_v1(&right.encode_state_as_update_v1())
            .unwrap();
        late.apply_update_v1(&left.encode_state_as_update_v1())
            .unwrap();
        for doc in [&left, &right] {
            doc.apply_update_v1(&late.encode_state_as_update_v1())
                .unwrap();
        }
        let keeper = a.clone().min(b.clone());
        for doc in [&left, &right, &late] {
            let converged: HashMap<_, _> = ids(doc).into_iter().collect();
            let values: HashSet<_> = converged.values().cloned().collect();
            assert_eq!(values.len(), 2);
            let story = if keeper == a { "a" } else { "b" };
            assert_eq!(
                converged[&session(story, &keeper)].as_deref(),
                Some("0000BBBB")
            );
        }
        assert_eq!(ids(&left), ids(&late));
        assert_eq!(ids(&right), ids(&late));
    }

    #[test]
    fn conflicting_saved_claims_are_reported_and_resolve_as_ambiguous() {
        let base = EditingDoc::new(1);
        base.create_story("a", "one", "Normal", "left").unwrap();
        base.create_story("b", "two", "Normal", "left").unwrap();
        let update = base.encode_state_as_update_v1();
        let (left, right) = (EditingDoc::new(2), EditingDoc::new(3));
        left.apply_update_v1(&update).unwrap();
        right.apply_update_v1(&update).unwrap();
        let a = base.paragraphs("a").unwrap()[0].para_id.clone();
        let b = base.paragraphs("b").unwrap()[0].para_id.clone();
        claim_id(&left, &a, "0000CCCC", true);
        claim_id(&right, &b, "0000CCCC", true);
        left.apply_update_v1(&right.encode_state_as_update_v1())
            .unwrap();
        let report = left.persist_paragraph_ids().unwrap();
        assert!(report.assignments.is_empty());
        assert_eq!(
            report.diagnostics,
            [
                ParagraphIdDiagnostic::ConflictingSavedIds {
                    ooxml_para_id: "0000CCCC".into(),
                    paragraphs: vec![session("a", &a), session("b", &b)],
                },
                ParagraphIdDiagnostic::NoSourcePackage,
            ]
        );
        assert!(
            ids(&left)
                .iter()
                .all(|(_, id)| id.as_deref() == Some("0000CCCC"))
        );
    }

    #[test]
    fn a_deleted_paragraph_keeps_its_saved_id_from_an_offline_claim() {
        let base = EditingDoc::new(1);
        base.create_story("body", "abc", "Normal", "left").unwrap();
        base.create_story("note", "n", "Normal", "left").unwrap();
        let update = base.encode_state_as_update_v1();
        let (left, offline) = (EditingDoc::new(2), EditingDoc::new(3));
        left.apply_update_v1(&update).unwrap();
        offline.apply_update_v1(&update).unwrap();
        left.split_paragraph(&ctx(), Position::new("body", 1), None)
            .unwrap();
        let saved = left
            .split_paragraph(&ctx(), Position::new("body", 3), None)
            .unwrap()
            .first_para_id;
        let id = doc_identity(&left, &saved).ooxml_para_id.unwrap();
        assert!(
            left.record_saved_paragraph_ids(&[(saved.clone(), id.clone())])
                .is_empty()
        );
        left.delete_range(&ctx(), StoryRange::new("body", 1, 3))
            .unwrap();
        assert!(
            ids(&left)
                .iter()
                .all(|(_, current)| current.as_deref() != Some(id.as_str()))
        );

        let note = base.paragraphs("note").unwrap()[0].para_id.clone();
        claim_id(&offline, &note, &id, false);
        left.apply_update_v1(&offline.encode_state_as_update_v1())
            .unwrap();
        let reassigned = doc_identity(&left, &note);
        assert_ne!(reassigned.ooxml_para_id.as_deref(), Some(id.as_str()));
        assert_eq!(reassigned.id_origin, Some(ParagraphIdOrigin::Repaired));
        assert!(
            ids(&left)
                .iter()
                .all(|(_, current)| current.as_deref() != Some(id.as_str()))
        );
    }

    #[test]
    fn a_hydrated_replica_never_reuses_a_deleted_paragraph_key() {
        let first = EditingDoc::new(7);
        first.create_story("body", "abc", "Normal", "left").unwrap();
        first
            .split_paragraph(&ctx(), Position::new("body", 1), None)
            .unwrap();
        let deleted = first
            .split_paragraph(&ctx(), Position::new("body", 3), None)
            .unwrap()
            .first_para_id;
        first
            .delete_range(&ctx(), StoryRange::new("body", 1, 3))
            .unwrap();
        let stale = ParagraphAnchor::Session {
            session_id: first.session_id(),
            story: "body".into(),
            para_id: deleted.clone(),
        };
        assert_eq!(
            first.resolve_paragraph_anchor(&stale),
            AnchorResolution::Missing
        );

        let hydrated = EditingDoc::new(7);
        hydrated
            .apply_update_v1(&first.encode_state_as_update_v1())
            .unwrap();
        let split = hydrated
            .split_paragraph(&ctx(), Position::new("body", 1), None)
            .unwrap();
        assert_ne!(split.second_para_id, deleted);
        assert_eq!(
            hydrated.resolve_paragraph_anchor(&stale),
            AnchorResolution::Missing
        );
    }

    #[test]
    fn a_save_reports_captured_ids_the_live_session_reassigned() {
        let base = EditingDoc::new(1);
        base.create_story("a", "one", "Normal", "left").unwrap();
        base.create_story("b", "two", "Normal", "left").unwrap();
        let update = base.encode_state_as_update_v1();
        let (left, right) = (EditingDoc::new(2), EditingDoc::new(3));
        left.apply_update_v1(&update).unwrap();
        right.apply_update_v1(&update).unwrap();
        let a = base.paragraphs("a").unwrap()[0].para_id.clone();
        let b = base.paragraphs("b").unwrap()[0].para_id.clone();
        claim_id(&left, &a, "0000DDDD", false);
        claim_id(&left, &b, "0000EEEE", false);
        let captured = vec![
            (a.clone(), "0000DDDD".to_owned()),
            (b.clone(), "0000EEEE".to_owned()),
        ];
        claim_id(&right, &b, "0000DDDD", true);
        left.apply_update_v1(&right.encode_state_as_update_v1())
            .unwrap();
        assert_ne!(
            doc_identity(&left, &a).ooxml_para_id.as_deref(),
            Some("0000DDDD")
        );

        assert_eq!(left.record_saved_paragraph_ids(&captured), captured);
        let txn = left.yrs_doc().transact();
        assert!(!claim(&claims(&txn), "0000DDDD", &a).unwrap().published);
        drop(txn);
        let current = doc_identity(&left, &a).ooxml_para_id.unwrap();
        assert!(
            left.record_saved_paragraph_ids(&[(a.clone(), current.clone())])
                .is_empty()
        );
        let txn = left.yrs_doc().transact();
        assert!(claim(&claims(&txn), &current, &a).unwrap().published);
    }

    #[test]
    fn a_synthetic_paragraph_is_promoted_by_authoring_into_it() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "", "Normal", "left").unwrap();
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let story = story_ref(&txn, "body").unwrap();
            let (_, map) = crate::next_pilcrow(&story, &txn, 0).unwrap();
            map.insert(&mut txn, PARA_ORIGIN, SYNTHETIC);
        }
        assert_eq!(
            doc.paragraph_identities().paragraphs[0].origin,
            ParagraphOrigin::Synthetic
        );
        assert!(doc.persist_paragraph_ids().unwrap().assignments.is_empty());
        doc.insert_text(
            &ctx(),
            Position::new("body", 0),
            "typed",
            FormatPolicy::Plain,
        )
        .unwrap();
        let identity = &doc.paragraph_identities().paragraphs[0];
        assert_eq!(identity.origin, ParagraphOrigin::Authored);
        assert_eq!(identity.id_origin, Some(ParagraphIdOrigin::Authored));
        assert!(identity.ooxml_para_id.is_some());
    }

    #[test]
    fn session_ids_identify_the_replicated_opening() {
        let first = EditingDoc::new(7);
        first.create_story("body", "a", "Normal", "left").unwrap();
        first.begin_opening(None);
        let replica = EditingDoc::new(8);
        replica
            .apply_update_v1(&first.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(first.session_id(), replica.session_id());
        let again = EditingDoc::new(7);
        again.create_story("body", "a", "Normal", "left").unwrap();
        again.begin_opening(None);
        assert_ne!(again.session_id(), first.session_id());
    }
}
