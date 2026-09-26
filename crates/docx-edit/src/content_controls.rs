//! Content-control discovery: every control of a document in document order, with the metadata
//! the structured export reports, its canonical text, placement, anchor, nesting and effective
//! lock.
//!
//! One inventory serves a live session ([`EditingDoc::list_content_controls`]), DOCX bytes
//! ([`list_docx_content_controls`]) and a parsed package ([`list_package_content_controls`], which
//! the native facade uses for its current model), and resolves the targets of
//! `setContentControlText` batch steps.

use std::collections::{BTreeMap, HashMap, HashSet};

use serde::{Deserialize, Serialize};
use yrs::branch::{Branch, BranchID};
use yrs::{Any, Map, MapRef, Out, ReadTxn, Transact};

use crate::control_source::{ControlSafety, SourceControl, may_hold_controls, safety_key};
use crate::control_values::{is_text_type, resolved_type};
use crate::ops::ChunkKind;
use crate::policy::Ownership;
use crate::read_types::{block_control_id, control_metadata, inline_control_id, nested_control_id};
use crate::segments::is_block_embed;
use crate::structured::source::{ReadSource, SourceParts, holds_control, story_root};
use crate::structured::{
    ExportError, ExportFailure, ExportFailureCode, ExportRead, ExportRefusal, StoryKind,
    byte_limit, invalid_options, limit_exceeded,
};
use crate::{DEL, EditingDoc, INS, KIND_KEY, PARA_ID, PPR_DEL, PPR_INS, map_string, story_ref};

pub use crate::read_types::{Anchor, ControlMetadata, StorySelection};
pub use crate::structured::{AnchorScope, Diagnostic, DiagnosticCode, Severity};

/// The only content-control snapshot schema version this crate reads and writes.
pub const CONTROLS_SCHEMA_VERSION: u8 = 1;
/// Controls a read returns when the options name no limit.
pub const DEFAULT_MAX_CONTROLS: u32 = 10_000;
/// The largest control limit a read accepts.
pub const MAX_CONTROLS_LIMIT: u32 = 1_000_000;
/// Nesting levels of tables and controls a read descends.
const MAX_DEPTH: usize = 32;
/// Stream chunks one read visits.
const MAX_VISITED: usize = 16_000_000;
/// Diagnostics of one code a read returns before it summarizes the rest.
const MAX_DIAGNOSTICS_PER_CODE: usize = 100;
/// The client id of the private sessions bytes and packages seed; reads never reveal it.
const SNAPSHOT_CLIENT_ID: u64 = 1;

/// Whether a control sits inside a paragraph or wraps whole blocks.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ControlPlacement {
    Inline,
    Block,
}

/// A control's canonical displayed text, or why it has no faithful text projection.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum ControlValue {
    /// Tabs stay tabs; line breaks and paragraph boundaries read as LF.
    Text {
        text: String,
    },
    Unavailable {
        reason: ValueUnavailable,
    },
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum ValueUnavailable {
    NonTextContent,
    TrackedRevisions,
    ProvenanceUnavailable,
    UnsupportedStory,
}

/// The restrictions that apply to a control, its own lock and its ancestors' combined. `known`
/// is false when any of them carries a lock value this crate does not know.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EffectiveLock {
    pub content: bool,
    pub control: bool,
    pub known: bool,
}

/// One content control. `metadata.lock` is the control's own authored lock.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ContentControl {
    #[serde(flatten)]
    pub metadata: ControlMetadata,
    pub placement: ControlPlacement,
    pub anchor: Anchor,
    pub parent_control_id: Option<String>,
    pub value: ControlValue,
    /// Whether a plain-text control accepts line breaks; `None` for other control types.
    pub multi_line: Option<bool>,
    pub effective_lock: EffectiveLock,
}

/// What to list. `stories` defaults to every category.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ContentControlsOptions {
    #[serde(default)]
    pub stories: Option<Vec<StorySelection>>,
    #[serde(default)]
    pub max_controls: Option<u32>,
    #[serde(default)]
    pub max_bytes: Option<u32>,
}

/// The control a write addresses: its engine id, or its tag, which must then be unique.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum ContentControlSelector {
    Id { control_id: String },
    Tag { tag: String },
}

/// Which controls a find returns: every exact, case-sensitive match.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum ContentControlQuery {
    Id { control_id: String },
    Tag { tag: String },
    Alias { alias: String },
}

/// The controls of a document. `complete` is false when coverage gaps are diagnosed.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ContentControlsSnapshot {
    pub schema_version: u8,
    pub anchor_scope: AnchorScope,
    pub included_stories: Vec<StorySelection>,
    pub controls: Vec<ContentControl>,
    pub complete: bool,
    pub diagnostics: Vec<Diagnostic>,
}

/// Where a session control lives.
#[derive(Clone, Debug)]
pub(crate) enum Site {
    /// A control embed at story index `raw` in paragraph `para_id`.
    Inline { raw: u32, para_id: String },
    /// A control inside another inline control's content.
    Nested,
    /// A block control embed at story index `raw` owning story `child`.
    Block { raw: u32, child: String },
    /// Only the source package holds it.
    Source,
}

/// One control of the inventory with what writes need to know about it.
#[derive(Clone, Debug)]
pub(crate) struct ControlRecord {
    pub control: ContentControl,
    /// `None` for a story that cannot be classified without the source package.
    pub category: Option<StorySelection>,
    /// Listed by reads; false for a second story reading an already listed part.
    pub listed: bool,
    /// The story holding the control's embed.
    pub story: String,
    pub site: Site,
    /// Controls nested inside it.
    pub children: bool,
    /// The embed carries a tracked insertion or deletion.
    pub stamped: bool,
    pub payload: HashMap<String, Any>,
    /// Its source content, when the package was read for it.
    pub safety: Option<ControlSafety>,
    /// The control is a copy another story holds because it reads the same part; writes
    /// address the first story's control.
    pub alias: bool,
    /// The control's embed, for a control that is one.
    pub embed: Option<BranchID>,
    /// The paragraph holding the control, or a block control's first paragraph.
    pub paragraph: Option<String>,
}

impl ControlRecord {
    pub fn id(&self) -> &str {
        &self.control.metadata.control_id
    }
}

/// A coverage gap and the story category it concerns (`None`: every read).
struct Gap {
    category: Option<StorySelection>,
    diagnostic: Diagnostic,
}

/// Every control of a document, listed or not, with its coverage gaps.
pub(crate) struct Inventory {
    pub records: Vec<ControlRecord>,
    gaps: Vec<Gap>,
    notes: Vec<Diagnostic>,
    /// Content controls may exist that the inventory cannot see or identify.
    pub unidentified: bool,
    /// The copies of each control that other stories reading its part hold.
    pub companions: HashMap<usize, Vec<usize>>,
    /// Controls whose part another story reads with different controls.
    pub divergent: HashSet<usize>,
}

#[derive(Clone)]
struct Ancestor {
    control_id: String,
    lock: Option<String>,
}

/// The locks one lock value imposes: `(content, control, known)`.
fn lock_state(lock: Option<&str>) -> (bool, bool, bool) {
    match lock {
        None | Some("unlocked") => (false, false, true),
        Some("sdtLocked") => (false, true, true),
        Some("contentLocked") => (true, false, true),
        Some("sdtContentLocked") => (true, true, true),
        Some(_) => (false, false, false),
    }
}

fn effective_lock(own: Option<&str>, ancestors: &[Ancestor]) -> EffectiveLock {
    let (content, control, known) = lock_state(own);
    let mut lock = EffectiveLock {
        content,
        control,
        known,
    };
    for ancestor in ancestors {
        let (content, _, known) = lock_state(ancestor.lock.as_deref());
        lock.content |= content;
        lock.control |= content;
        lock.known &= known;
    }
    lock
}

fn active(value: Option<&Any>) -> bool {
    value.is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
}

fn any_map(value: Option<&Any>) -> Option<&HashMap<String, Any>> {
    match value? {
        Any::Map(map) => Some(map),
        _ => None,
    }
}

fn any_str(value: Option<&Any>) -> Option<&str> {
    match value? {
        Any::String(value) => Some(value),
        _ => None,
    }
}

fn payload_of<T: ReadTxn>(map: &MapRef, txn: &T) -> HashMap<String, Any> {
    map.iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_owned(), value)),
            _ => None,
        })
        .collect()
}

pub(crate) fn content_items(payload: &HashMap<String, Any>) -> &[Any] {
    match payload.get("content") {
        Some(Any::Array(items)) => items,
        _ => &[],
    }
}

pub(crate) fn item_stamped(item: &HashMap<String, Any>) -> bool {
    any_map(item.get("attrs")).is_some_and(|attrs| active(attrs.get(INS)) || active(attrs.get(DEL)))
}

/// Whether a plain-text control accepts line breaks, read from its typed properties first (the
/// flat field, then the parsed properties) and its captured `w:sdtPr` last; false when absent.
fn multi_line(payload: &HashMap<String, Any>, control_type: &str) -> Option<bool> {
    if control_type != "plainText" {
        return None;
    }
    let flat = match payload.get("multiLine") {
        Some(Any::Bool(value)) => Some(*value),
        _ => None,
    };
    let typed = flat.or_else(|| {
        any_str(payload.get("propertiesJson"))
            .and_then(|json| serde_json::from_str::<serde_json::Value>(json).ok())
            .and_then(|properties| properties.get("multiLine")?.as_bool())
    });
    let parsed = typed.or_else(|| {
        any_str(payload.get("rawPropertiesXml"))
            .and_then(|raw| docx_parse::parse_sdt_properties_xml(raw).ok())
            .and_then(|properties| properties.multi_line)
    });
    Some(parsed.unwrap_or(false))
}

/// The canonical text of an inline control's content and whether it is faithful.
fn inline_text(items: &[Any]) -> Result<String, ValueUnavailable> {
    let mut text = String::new();
    let mut non_text = false;
    for item in items {
        let Some(item) = any_map(Some(item)) else {
            continue;
        };
        if item_stamped(item) {
            return Err(ValueUnavailable::TrackedRevisions);
        }
        match any_str(item.get("kind")).unwrap_or_default() {
            "text" => text.push_str(any_str(item.get("text")).unwrap_or_default()),
            "tab" => text.push('\t'),
            "break" => text.push('\n'),
            _ => non_text = true,
        }
    }
    if non_text {
        Err(ValueUnavailable::NonTextContent)
    } else {
        Ok(text)
    }
}

/// What a story walk found, for the value of the block control owning it.
#[derive(Default)]
struct StorySummary {
    paragraphs: Vec<String>,
    first_paragraph: Option<String>,
    non_text: bool,
    revisions: bool,
    controls: bool,
}

fn source_metadata(control: &SourceControl, control_id: String) -> ControlMetadata {
    let properties = &control.properties;
    ControlMetadata {
        control_id,
        ooxml_id: properties.id.map(|id| {
            if id.is_finite() && id.fract() == 0.0 {
                format!("{id:.0}")
            } else {
                id.to_string()
            }
        }),
        control_type: properties.sdt_type.clone(),
        tag: properties.tag.clone(),
        alias: properties.alias.clone(),
        lock: properties.lock.clone(),
        showing_placeholder: properties.showing_placeholder == Some(true),
        data_bound: properties.data_binding.is_some(),
    }
}

/// The control id of a control only the source holds: its part and element path.
fn source_control_id(anchor: &Anchor) -> String {
    match anchor {
        Anchor::SourcePart { part, path, .. } => format!(
            "{part}#{}",
            path.iter()
                .map(u32::to_string)
                .collect::<Vec<_>>()
                .join("/")
        ),
        Anchor::Control { control_id, .. } => control_id.clone(),
        _ => String::new(),
    }
}

fn diagnostic(
    code: DiagnosticCode,
    severity: Severity,
    anchor: Option<Anchor>,
    message: impl Into<String>,
) -> Diagnostic {
    Diagnostic {
        code,
        severity,
        anchor,
        message: message.into(),
    }
}

struct Builder<'a, T: ReadTxn> {
    doc: &'a EditingDoc,
    txn: &'a T,
    read: Option<&'a ReadSource>,
    records: Vec<ControlRecord>,
    visited: usize,
    walked: HashSet<String>,
    /// The source controls of each story and raw block, in document order.
    sources: HashMap<(&'a str, Option<usize>), Vec<usize>>,
    /// The source controls that hold other controls.
    parents: HashSet<usize>,
}

impl<'a, T: ReadTxn> Builder<'a, T> {
    fn part(&self, story: &str) -> Option<String> {
        self.read?.story_part(story)
    }

    /// The safety of the control's own source occurrence when seeding paired its embed with it,
    /// or else of every occurrence with its captured properties in its part.
    fn safety(
        &self,
        story: &str,
        payload: &HashMap<String, Any>,
        embed: Option<&BranchID>,
    ) -> Option<ControlSafety> {
        let read = self.read?;
        if let Some(safety) = embed.and_then(|embed| read.embed_safety.get()?.get(embed)) {
            return Some(safety.clone());
        }
        let table = read.control_safety.as_ref()?;
        let key = safety_key(any_str(payload.get("rawPropertiesXml")));
        table.get(&(self.part(story)?, key)).cloned()
    }

    /// Records a control before its descendants and returns its index.
    #[allow(clippy::too_many_arguments)]
    fn push(
        &mut self,
        metadata: ControlMetadata,
        placement: ControlPlacement,
        story: &str,
        category: Option<StorySelection>,
        listed: bool,
        site: Site,
        stamped: bool,
        payload: HashMap<String, Any>,
        ancestors: &[Ancestor],
        embed: Option<BranchID>,
    ) -> usize {
        let anchor = Anchor::Control {
            story: story.to_owned(),
            control_id: metadata.control_id.clone(),
        };
        let multi_line = multi_line(&payload, &metadata.control_type);
        let safety = match site {
            Site::Source => None,
            _ => self.safety(story, &payload, embed.as_ref()),
        };
        self.records.push(ControlRecord {
            control: ContentControl {
                effective_lock: effective_lock(metadata.lock.as_deref(), ancestors),
                parent_control_id: ancestors.last().map(|parent| parent.control_id.clone()),
                metadata,
                placement,
                anchor,
                value: ControlValue::Unavailable {
                    reason: ValueUnavailable::NonTextContent,
                },
                multi_line,
            },
            category,
            listed,
            story: story.to_owned(),
            site,
            children: false,
            stamped,
            payload,
            safety,
            alias: false,
            embed,
            paragraph: None,
        });
        self.records.len() - 1
    }

    fn settle(&mut self, index: usize, text: Result<String, ValueUnavailable>, children: bool) {
        let record = &mut self.records[index];
        let text = text.and_then(|text| {
            if record.stamped {
                return Err(ValueUnavailable::TrackedRevisions);
            }
            match &record.safety {
                Some(safety) if safety.revisions => Err(ValueUnavailable::TrackedRevisions),
                Some(safety) if safety.non_text => Err(ValueUnavailable::NonTextContent),
                _ => Ok(text),
            }
        });
        record.control.value = match text {
            Ok(text) => ControlValue::Text { text },
            Err(reason) => ControlValue::Unavailable { reason },
        };
        record.children = children;
    }

    fn visit(&mut self) -> Result<(), ExportFailure> {
        self.visited += 1;
        if self.visited > MAX_VISITED {
            return Err(limit_exceeded(format!(
                "a content-control read visits at most {MAX_VISITED} stream chunks"
            )));
        }
        Ok(())
    }

    fn deep(depth: usize) -> Result<(), ExportFailure> {
        if depth > MAX_DEPTH {
            return Err(limit_exceeded(format!(
                "content is nested more than {MAX_DEPTH} levels deep"
            )));
        }
        Ok(())
    }

    /// Records the source-only controls of raw block `raw_block` of `story`.
    fn source_controls(
        &mut self,
        story: &str,
        raw_block: usize,
        category: Option<StorySelection>,
        listed: bool,
        ancestors: &[Ancestor],
    ) {
        let indices = self
            .sources
            .get(&(story, Some(raw_block)))
            .cloned()
            .unwrap_or_default();
        self.push_source(&indices, category, listed, ancestors);
    }

    fn push_source(
        &mut self,
        indices: &[usize],
        category: Option<StorySelection>,
        listed: bool,
        ancestors: &[Ancestor],
    ) {
        let Some(read) = self.read else {
            return;
        };
        let mut chains: HashMap<usize, Vec<Ancestor>> = HashMap::new();
        for &index in indices {
            let control = &read.source_controls[index];
            let control_id = source_control_id(&control.anchor);
            let metadata = source_metadata(control, control_id.clone());
            let chain = control
                .parent
                .and_then(|parent| chains.get(&parent).cloned())
                .unwrap_or_else(|| ancestors.to_vec());
            let lock = metadata.lock.clone();
            let multi_line = (metadata.control_type == "plainText")
                .then(|| control.properties.multi_line.unwrap_or(false));
            let record = self.records.len();
            self.records.push(ControlRecord {
                control: ContentControl {
                    effective_lock: effective_lock(lock.as_deref(), &chain),
                    parent_control_id: chain.last().map(|parent| parent.control_id.clone()),
                    metadata,
                    placement: if control.block {
                        ControlPlacement::Block
                    } else {
                        ControlPlacement::Inline
                    },
                    anchor: control.anchor.clone(),
                    value: ControlValue::Unavailable {
                        reason: ValueUnavailable::UnsupportedStory,
                    },
                    multi_line,
                },
                category,
                listed,
                story: control.story.clone(),
                site: Site::Source,
                children: self.parents.contains(&index),
                stamped: false,
                payload: HashMap::new(),
                safety: None,
                alias: false,
                embed: None,
                paragraph: None,
            });
            let mut inner = chain;
            inner.push(Ancestor {
                control_id: self.records[record].id().to_owned(),
                lock,
            });
            chains.insert(index, inner);
        }
    }

    /// Walks one story in document order, recording its controls and those of every story it
    /// owns at their occurrence.
    fn story(
        &mut self,
        story: &str,
        category: Option<StorySelection>,
        listed: bool,
        ancestors: &[Ancestor],
        depth: usize,
    ) -> Result<StorySummary, ExportFailure> {
        Self::deep(depth)?;
        let mut summary = StorySummary::default();
        if !self.walked.insert(story.to_owned()) {
            return Ok(summary);
        }
        let Ok(text) = story_ref(self.txn, story) else {
            return Ok(summary);
        };
        let chunks = self.doc.chunk_snapshot(story, &text, self.txn);
        let (mut raw_before, trailing) = self.raw_positions(story);
        let mut paragraph = String::new();
        let mut node_start = 0u32;
        let mut pending: Vec<(u32, MapRef, bool)> = Vec::new();
        for chunk in chunks.iter() {
            self.visit()?;
            let stamped = chunk.attr_active(INS) || chunk.attr_active(DEL);
            match &chunk.kind {
                ChunkKind::Pilcrow(map) => {
                    let para_id = map_string(map, self.txn, PARA_ID).unwrap_or_default();
                    if stamped
                        || [PPR_INS, PPR_DEL].into_iter().any(|key| {
                            matches!(map.get(self.txn, key), Some(Out::Any(value)) if active(Some(&value)))
                        })
                    {
                        summary.revisions = true;
                    }
                    for raw in raw_before.remove(&para_id).unwrap_or_default() {
                        summary.non_text = true;
                        self.source_controls(story, raw, category, listed, ancestors);
                    }
                    for (ordinal, (raw, map, stamped)) in
                        std::mem::take(&mut pending).into_iter().enumerate()
                    {
                        summary.controls = true;
                        let control_id = inline_control_id(story, &para_id, ordinal);
                        self.inline(
                            story, &para_id, raw, &map, stamped, control_id, category, listed,
                            ancestors, depth,
                        )?;
                    }
                    summary
                        .first_paragraph
                        .get_or_insert_with(|| para_id.clone());
                    summary.paragraphs.push(std::mem::take(&mut paragraph));
                    node_start = chunk.end();
                }
                ChunkKind::Text(text) => {
                    summary.revisions |= stamped;
                    paragraph.push_str(text);
                }
                ChunkKind::Embed(map) => {
                    summary.revisions |= stamped;
                    let kind = map
                        .as_ref()
                        .and_then(|map| map_string(map, self.txn, KIND_KEY))
                        .unwrap_or_default();
                    if chunk.start == node_start && is_block_embed(&kind) {
                        node_start = chunk.end();
                        summary.non_text = true;
                        let Some(map) = map else { continue };
                        match kind.as_str() {
                            "table" => {
                                let cells = table_cells(map, self.txn);
                                if let Some(id) = cells.first().and_then(|cell| {
                                    crate::structured::source::cell_table(cell).map(str::to_owned)
                                }) {
                                    for raw in raw_before.remove(&id).unwrap_or_default() {
                                        self.source_controls(
                                            story, raw, category, listed, ancestors,
                                        );
                                    }
                                }
                                for cell in cells {
                                    let inner =
                                        self.story(&cell, category, listed, ancestors, depth + 1)?;
                                    summary.controls |= inner.controls;
                                }
                            }
                            "blockSdt" => {
                                summary.controls = true;
                                self.block(
                                    story,
                                    chunk.start,
                                    map,
                                    stamped,
                                    category,
                                    listed,
                                    ancestors,
                                    depth,
                                    &mut raw_before,
                                )?;
                            }
                            _ => {}
                        }
                        continue;
                    }
                    match kind.as_str() {
                        "break" => paragraph.push('\n'),
                        "sdt" => {
                            summary.non_text = true;
                            if let Some(map) = map {
                                pending.push((chunk.start, map.clone(), stamped));
                            }
                        }
                        _ => summary.non_text = true,
                    }
                }
            }
        }
        for raw in raw_before.into_values().flatten().chain(trailing) {
            summary.non_text = true;
            self.source_controls(story, raw, category, listed, ancestors);
        }
        Ok(summary)
    }

    /// Which raw blocks of `story` precede each block id, and which follow every block.
    fn raw_positions(&self, story: &str) -> (HashMap<String, Vec<usize>>, Vec<usize>) {
        let mut before: HashMap<String, Vec<usize>> = HashMap::new();
        let mut pending = Vec::new();
        let Some(order) = self
            .read
            .and_then(|read| read.provenance.block_order.get(story))
        else {
            return (before, pending);
        };
        let mut raw = 0usize;
        for entry in order {
            match entry {
                None => {
                    pending.push(raw);
                    raw += 1;
                }
                Some(id) if !pending.is_empty() => {
                    before
                        .entry(id.clone())
                        .or_default()
                        .extend(std::mem::take(&mut pending));
                }
                Some(_) => {}
            }
        }
        (before, pending)
    }

    #[allow(clippy::too_many_arguments)]
    fn block(
        &mut self,
        story: &str,
        raw: u32,
        map: &MapRef,
        stamped: bool,
        category: Option<StorySelection>,
        listed: bool,
        ancestors: &[Ancestor],
        depth: usize,
        raw_before: &mut HashMap<String, Vec<usize>>,
    ) -> Result<(), ExportFailure> {
        let payload = payload_of(map, self.txn);
        let Some(child) = any_str(payload.get("story")).map(str::to_owned) else {
            return Ok(());
        };
        for raw in raw_before.remove(&child).unwrap_or_default() {
            self.source_controls(story, raw, category, listed, ancestors);
        }
        let metadata = control_metadata(block_control_id(&child), |key| payload.get(key));
        let lock = metadata.lock.clone();
        let index = self.push(
            metadata,
            ControlPlacement::Block,
            story,
            category,
            listed,
            Site::Block {
                raw,
                child: child.clone(),
            },
            stamped,
            payload,
            ancestors,
            Some(AsRef::<Branch>::as_ref(map).id()),
        );
        let mut inner = ancestors.to_vec();
        inner.push(Ancestor {
            control_id: self.records[index].id().to_owned(),
            lock,
        });
        let summary = self.story(&child, category, listed, &inner, depth + 1)?;
        self.records[index].paragraph = summary.first_paragraph;
        let text = if summary.revisions {
            Err(ValueUnavailable::TrackedRevisions)
        } else if summary.non_text {
            Err(ValueUnavailable::NonTextContent)
        } else {
            Ok(summary.paragraphs.join("\n"))
        };
        self.settle(index, text, summary.controls);
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn inline(
        &mut self,
        story: &str,
        para_id: &str,
        raw: u32,
        map: &MapRef,
        stamped: bool,
        control_id: String,
        category: Option<StorySelection>,
        listed: bool,
        ancestors: &[Ancestor],
        depth: usize,
    ) -> Result<(), ExportFailure> {
        let payload = payload_of(map, self.txn);
        self.inline_payload(
            story,
            Site::Inline {
                raw,
                para_id: para_id.to_owned(),
            },
            payload,
            stamped,
            control_id,
            category,
            listed,
            ancestors,
            depth,
            (Some(AsRef::<Branch>::as_ref(map).id()), para_id),
        )
    }

    #[allow(clippy::too_many_arguments)]
    fn inline_payload(
        &mut self,
        story: &str,
        site: Site,
        payload: HashMap<String, Any>,
        stamped: bool,
        control_id: String,
        category: Option<StorySelection>,
        listed: bool,
        ancestors: &[Ancestor],
        depth: usize,
        (embed, paragraph): (Option<BranchID>, &str),
    ) -> Result<(), ExportFailure> {
        Self::deep(depth)?;
        let mut metadata = control_metadata(control_id, |key| payload.get(key));
        metadata.control_type = resolved_type(Some("sdt"), |key| payload.get(key));
        let lock = metadata.lock.clone();
        let items = content_items(&payload).to_vec();
        let text = inline_text(&items);
        let index = self.push(
            metadata,
            ControlPlacement::Inline,
            story,
            category,
            listed,
            site,
            stamped,
            payload,
            ancestors,
            embed,
        );
        self.records[index].paragraph = Some(paragraph.to_owned());
        let mut inner = ancestors.to_vec();
        inner.push(Ancestor {
            control_id: self.records[index].id().to_owned(),
            lock,
        });
        let mut nested = 0usize;
        for item in &items {
            self.visit()?;
            let Some(item) = any_map(Some(item)) else {
                continue;
            };
            if any_str(item.get("kind")) != Some("sdt")
                || any_map(item.get("attrs")).is_some_and(|attrs| active(attrs.get(DEL)))
            {
                continue;
            }
            let payload = any_map(item.get("payload")).cloned().unwrap_or_default();
            let id = nested_control_id(self.records[index].id(), nested);
            nested += 1;
            self.inline_payload(
                story,
                Site::Nested,
                payload,
                stamped || item_stamped(item),
                id,
                category,
                listed,
                &inner,
                depth + 1,
                (None, paragraph),
            )?;
        }
        self.settle(index, text, nested > 0);
        Ok(())
    }
}

fn table_cells<T: ReadTxn>(map: &MapRef, txn: &T) -> Vec<String> {
    let Some(Out::Any(Any::Array(rows))) = map.get(txn, "rows") else {
        return Vec::new();
    };
    let mut cells = Vec::new();
    for row in rows.iter() {
        let Some(Any::Array(row_cells)) = any_map(Some(row)).and_then(|row| row.get("cells"))
        else {
            continue;
        };
        for cell in row_cells.iter() {
            if let Some(story) = any_map(Some(cell)).and_then(|cell| any_str(cell.get("story"))) {
                cells.push(story.to_owned());
            }
        }
    }
    cells
}

/// Whether two stories reading one part hold the same control at the same position.
fn same_control(copy: &ControlRecord, control: &ControlRecord) -> bool {
    let (copy_control, original) = (&copy.control, &control.control);
    let metadata = ControlMetadata {
        control_id: original.metadata.control_id.clone(),
        ..copy_control.metadata.clone()
    };
    metadata == original.metadata
        && copy_control.value == original.value
        && copy_control.placement == original.placement
        && copy_control.multi_line == original.multi_line
        && copy_control.effective_lock == original.effective_lock
        && copy_control.parent_control_id.is_some() == original.parent_control_id.is_some()
        && copy.children == control.children
        && copy.stamped == control.stamped
        && std::mem::discriminant(&copy.site) == std::mem::discriminant(&control.site)
}

fn category_of(kind: StoryKind) -> StorySelection {
    match kind {
        StoryKind::Header => StorySelection::Headers,
        StoryKind::Footer => StorySelection::Footers,
        StoryKind::Footnote => StorySelection::Footnotes,
        StoryKind::Endnote => StorySelection::Endnotes,
        StoryKind::Body => StorySelection::Body,
        StoryKind::Comment => StorySelection::Comments,
    }
}

impl Inventory {
    /// Pairs the controls of each properties key whose occurrences differ in safety with those
    /// occurrences in document order, where the counts match and every pair sits in the same
    /// paragraph. Read on the state seeding left, the pairs follow each embed through later edits.
    pub(crate) fn occurrence_safety(&self, read: &ReadSource) -> HashMap<BranchID, ControlSafety> {
        let mut ordered: HashMap<(String, String), Vec<&ControlRecord>> = HashMap::new();
        for record in &self.records {
            if record.alias || matches!(record.site, Site::Source) {
                continue;
            }
            let Some(part) = read.story_part(&record.story) else {
                continue;
            };
            let key = (
                part,
                safety_key(any_str(record.payload.get("rawPropertiesXml"))),
            );
            if read.ambiguous_safety.contains_key(&key) {
                ordered.entry(key).or_default().push(record);
            }
        }
        let mut paired = HashMap::new();
        for (key, occurrences) in &read.ambiguous_safety {
            let records = ordered.get(key).map_or(&[][..], Vec::as_slice);
            let agree = records.len() == occurrences.len()
                && records
                    .iter()
                    .zip(occurrences)
                    .all(|(record, (paragraph, _))| {
                        paragraph.is_some() && record.paragraph == *paragraph
                    });
            if !agree {
                continue;
            }
            for (record, (_, safety)) in records.iter().zip(occurrences) {
                if let Some(embed) = &record.embed {
                    paired.insert(embed.clone(), safety.clone());
                }
            }
        }
        paired
    }

    /// Reads every control of `doc`'s committed state.
    pub(crate) fn build<T: ReadTxn>(doc: &EditingDoc, txn: &T) -> Result<Self, ExportFailure> {
        let source = doc.source_metadata();
        let read = source.as_deref().map(crate::seed::SourceMetadata::read);
        let story_ids: Vec<String> = txn
            .get_map(crate::STORIES)
            .map(|stories| stories.keys(txn).map(str::to_owned).collect())
            .unwrap_or_default();
        let known: HashSet<&str> = story_ids.iter().map(String::as_str).collect();
        let mut sources: HashMap<(&str, Option<usize>), Vec<usize>> = HashMap::new();
        let mut parents = HashSet::new();
        for (index, control) in read
            .iter()
            .flat_map(|read| read.source_controls.iter().enumerate())
        {
            sources
                .entry((control.story.as_str(), control.raw_block))
                .or_default()
                .push(index);
            parents.extend(control.parent);
        }
        let mut builder = Builder {
            doc,
            txn,
            read,
            records: Vec::new(),
            visited: 0,
            walked: HashSet::new(),
            sources,
            parents,
        };
        let mut gaps = Vec::new();
        let mut notes = Vec::new();
        if known.contains("body") {
            builder.story("body", Some(StorySelection::Body), true, &[], 0)?;
        }
        let mut roots: Vec<(StorySelection, String, Option<String>)> = Vec::new();
        if let Some(read) = read {
            let mut parts: HashMap<&str, &str> = HashMap::new();
            for source in &read.stories {
                if !known.contains(source.story.as_str()) {
                    continue;
                }
                let alias_of = match (&source.part, source.kind) {
                    (Some(part), StoryKind::Header | StoryKind::Footer) => {
                        match parts.get(part.as_str()) {
                            Some(first) => {
                                notes.push(diagnostic(
                                    DiagnosticCode::AmbiguousIdentity,
                                    Severity::Info,
                                    None,
                                    format!("Story {} reads part {part} as story {first} does, so only the controls of {first} are listed and fills write both.", source.story),
                                ));
                                Some((*first).to_owned())
                            }
                            None => {
                                parts.insert(part, &source.story);
                                None
                            }
                        }
                    }
                    _ => None,
                };
                roots.push((category_of(source.kind), source.story.clone(), alias_of));
            }
        }
        let rooted: HashSet<&str> = roots.iter().map(|(_, story, _)| story.as_str()).collect();
        let mut notes_by_id: Vec<(StorySelection, String)> = story_ids
            .iter()
            .filter(|story| story_root(story) == story.as_str() && !rooted.contains(story.as_str()))
            .filter_map(|story| {
                if story.starts_with("fn:") {
                    Some((StorySelection::Footnotes, story.clone()))
                } else if story.starts_with("en:") {
                    Some((StorySelection::Endnotes, story.clone()))
                } else {
                    None
                }
            })
            .collect();
        notes_by_id.sort_by_key(|(category, story)| {
            let id = &story[3..];
            (*category, id.parse::<i64>().ok(), id.to_owned())
        });
        roots.extend(
            notes_by_id
                .into_iter()
                .map(|(category, story)| (category, story, None)),
        );
        roots.sort_by_key(|(category, _, _)| *category);
        let mut ranges: HashMap<&str, std::ops::Range<usize>> = HashMap::new();
        for (category, story, alias_of) in &roots {
            let start = builder.records.len();
            builder.story(story, Some(*category), alias_of.is_none(), &[], 0)?;
            ranges.insert(story, start..builder.records.len());
        }
        let mut companions: HashMap<usize, Vec<usize>> = HashMap::new();
        let mut divergent = HashSet::new();
        let mut diverged = false;
        for (category, story, alias_of) in &roots {
            let Some(first) = alias_of else {
                continue;
            };
            let (Some(own), Some(canonical)) =
                (ranges.get(story.as_str()), ranges.get(first.as_str()))
            else {
                continue;
            };
            for record in &mut builder.records[own.clone()] {
                record.alias = true;
            }
            let records = &builder.records;
            if own.len() == canonical.len()
                && own
                    .clone()
                    .zip(canonical.clone())
                    .all(|(copy, control)| same_control(&records[copy], &records[control]))
            {
                for (copy, control) in own.clone().zip(canonical.clone()) {
                    companions.entry(control).or_default().push(copy);
                }
            } else {
                diverged = true;
                divergent.extend(canonical.clone());
                gaps.push(Gap {
                    category: Some(*category),
                    diagnostic: diagnostic(
                        DiagnosticCode::AmbiguousIdentity,
                        Severity::Warning,
                        None,
                        format!("Stories {story} and {first} read the same part but now hold different content controls; the controls of {story} are not listed and those of {first} cannot be filled."),
                    ),
                });
            }
        }
        let ownership = Ownership::build(doc, txn);
        let mut unclassified: Vec<&String> = story_ids
            .iter()
            .filter(|story| {
                story.as_str() != "body"
                    && !builder.walked.contains(story.as_str())
                    && !ownership.owns(story)
            })
            .collect();
        unclassified.sort();
        let mut unclassified_controls = 0usize;
        for story in unclassified {
            let before = builder.records.len();
            builder.story(story, None, false, &[], 0)?;
            unclassified_controls += builder.records.len() - before;
        }
        if unclassified_controls > 0 {
            gaps.push(Gap {
                category: Some(StorySelection::Headers),
                diagnostic: diagnostic(
                    DiagnosticCode::ProvenanceUnavailable,
                    Severity::Warning,
                    None,
                    format!("{unclassified_controls} content controls sit in header or footer stories that cannot be classified without the source package and are not listed."),
                ),
            });
        }
        let mut unidentified = diverged;
        match read {
            Some(read) => {
                for comment in &read.comments {
                    let story = format!("comment:{}", comment.id);
                    let indices = builder
                        .sources
                        .get(&(story.as_str(), None))
                        .cloned()
                        .unwrap_or_default();
                    builder.push_source(&indices, Some(StorySelection::Comments), true, &[]);
                }
                let unlocated = read.unlocated_controls
                    + if read.control_safety.is_none() {
                        read.provenance
                            .raw_sources
                            .iter()
                            .filter(|source| may_hold_controls(&source.xml))
                            .count()
                            + read
                                .comments
                                .iter()
                                .filter(|comment| comment.body.iter().any(holds_control))
                                .count()
                    } else {
                        0
                    };
                if read.unrepresented_controls > 0 {
                    unidentified = true;
                    let anchored = &read.unrepresented_anchors[..read
                        .unrepresented_anchors
                        .len()
                        .min(MAX_DIAGNOSTICS_PER_CODE)];
                    for anchor in anchored {
                        gaps.push(Gap {
                            category: None,
                            diagnostic: diagnostic(
                                DiagnosticCode::ProvenanceUnavailable,
                                Severity::Warning,
                                Some(anchor.clone()),
                                "This content control of the package is not held as a control in this session, as inside a text box or a tracked change; it is not listed.",
                            ),
                        });
                    }
                    let rest = read.unrepresented_controls - anchored.len();
                    if rest > 0 {
                        gaps.push(Gap {
                            category: None,
                            diagnostic: diagnostic(
                                DiagnosticCode::ProvenanceUnavailable,
                                Severity::Warning,
                                None,
                                format!("{rest} more content controls of the package are not held as controls in this session, as inside text boxes or tracked changes; they are not listed."),
                            ),
                        });
                    }
                }
                if unlocated > 0 {
                    unidentified = true;
                    gaps.push(Gap {
                        category: None,
                        diagnostic: diagnostic(
                            DiagnosticCode::ProvenanceUnavailable,
                            Severity::Warning,
                            None,
                            format!("{unlocated} raw XML blocks or comment bodies may hold content controls that could not be located in the package; they are not listed."),
                        ),
                    });
                }
            }
            None => {
                unidentified = true;
                gaps.push(Gap {
                    category: None,
                    diagnostic: diagnostic(
                        DiagnosticCode::ProvenanceUnavailable,
                        Severity::Info,
                        None,
                        "The document was not opened from DOCX bytes, so controls kept only as source XML (raw blocks, comment bodies) cannot be listed and no control can be filled.",
                    ),
                });
            }
        }
        let mut seen: HashMap<String, usize> = HashMap::new();
        for record in &builder.records {
            *seen.entry(record.id().to_owned()).or_default() += 1;
        }
        for (id, count) in seen.into_iter().filter(|(_, count)| *count > 1) {
            notes.push(diagnostic(
                DiagnosticCode::AmbiguousIdentity,
                Severity::Warning,
                None,
                format!("{count} controls share control id {id:?} because their paragraphs share an id; writes cannot address them by id."),
            ));
        }
        notes.sort_by(|left, right| left.message.cmp(&right.message));
        for record in &builder.records {
            if active(record.payload.get("value"))
                && is_text_type(&record.control.metadata.control_type)
            {
                notes.push(diagnostic(
                    DiagnosticCode::LegacyControlValue,
                    Severity::Warning,
                    Some(record.control.anchor.clone()),
                    "This text control carries an authored value from an older version, which saving ignores; its value here is the control's content. Filling the control drops it.",
                ));
            }
        }
        Ok(Self {
            records: builder.records,
            gaps,
            notes,
            unidentified,
            companions,
            divergent,
        })
    }

    /// The inventory as a read with `options`, keeping the controls `keep` selects.
    fn snapshot(
        &self,
        options: &Resolved,
        scope: AnchorScope,
        keep: impl Fn(&ControlRecord) -> bool,
    ) -> Result<ContentControlsSnapshot, ExportFailure> {
        let selected: HashSet<StorySelection> = options.stories.iter().copied().collect();
        let mut complete = true;
        let mut diagnostics = Vec::new();
        for gap in &self.gaps {
            let relevant = match gap.category {
                None => true,
                Some(StorySelection::Headers | StorySelection::Footers) => {
                    selected.contains(&StorySelection::Headers)
                        || selected.contains(&StorySelection::Footers)
                }
                Some(category) => selected.contains(&category),
            };
            if relevant {
                complete = false;
                diagnostics.push(gap.diagnostic.clone());
            }
        }
        let mut omitted: BTreeMap<StorySelection, usize> = BTreeMap::new();
        let mut controls = Vec::new();
        for record in &self.records {
            if !record.listed {
                continue;
            }
            let Some(category) = record.category else {
                continue;
            };
            if !selected.contains(&category) {
                *omitted.entry(category).or_default() += 1;
                continue;
            }
            if keep(record) {
                controls.push(record.control.clone());
            }
        }
        for (category, count) in omitted {
            let option = serde_json::to_value(category)
                .ok()
                .and_then(|value| value.as_str().map(str::to_owned))
                .unwrap_or_default();
            diagnostics.push(diagnostic(
                DiagnosticCode::StoriesOmitted,
                Severity::Info,
                None,
                format!("{count} content controls in {option} stories are not included; select \"{option}\" to include them."),
            ));
        }
        let mut per_code: BTreeMap<DiagnosticCode, usize> = BTreeMap::new();
        let mut summarized: BTreeMap<DiagnosticCode, usize> = BTreeMap::new();
        for note in &self.notes {
            let listed = match &note.anchor {
                Some(anchor) => controls.iter().any(|control| &control.anchor == anchor),
                None => true,
            };
            if !listed {
                continue;
            }
            let count = per_code.entry(note.code).or_default();
            *count += 1;
            if *count > MAX_DIAGNOSTICS_PER_CODE {
                *summarized.entry(note.code).or_default() += 1;
            } else {
                diagnostics.push(note.clone());
            }
        }
        for (code, count) in summarized {
            diagnostics.push(diagnostic(
                code,
                Severity::Info,
                None,
                format!("{count} more diagnostics of this kind are not listed."),
            ));
        }
        if controls.len() > options.max_controls {
            return Err(limit_exceeded(format!(
                "the read holds {} content controls, more than maxControls {}",
                controls.len(),
                options.max_controls
            )));
        }
        let snapshot = ContentControlsSnapshot {
            schema_version: CONTROLS_SCHEMA_VERSION,
            anchor_scope: scope,
            included_stories: options.stories.clone(),
            controls,
            complete,
            diagnostics,
        };
        let mut counter = Counter(0, options.max_bytes);
        if serde_json::to_writer(&mut counter, &snapshot).is_err() || counter.0 > options.max_bytes
        {
            return Err(limit_exceeded(format!(
                "the read serializes to more than maxBytes {} bytes",
                options.max_bytes
            )));
        }
        Ok(snapshot)
    }
}

/// Counts written bytes and fails past a limit.
struct Counter(usize, usize);

impl std::io::Write for Counter {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.0 += bytes.len();
        if self.0 > self.1 {
            return Err(std::io::Error::other("limit"));
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// Read options with their defaults applied and validated.
struct Resolved {
    stories: Vec<StorySelection>,
    max_controls: usize,
    max_bytes: usize,
}

impl Resolved {
    fn new(options: &ContentControlsOptions) -> Result<Self, ExportFailure> {
        let mut stories = options.stories.clone().unwrap_or_else(|| {
            vec![
                StorySelection::Body,
                StorySelection::Headers,
                StorySelection::Footers,
                StorySelection::Footnotes,
                StorySelection::Endnotes,
                StorySelection::Comments,
            ]
        });
        if stories.is_empty() {
            return Err(invalid_options("stories must name at least one category"));
        }
        stories.sort_unstable();
        stories.dedup();
        let max_controls = options.max_controls.unwrap_or(DEFAULT_MAX_CONTROLS);
        if max_controls == 0 {
            return Err(invalid_options("maxControls must be at least 1"));
        }
        if max_controls > MAX_CONTROLS_LIMIT {
            return Err(limit_exceeded(format!(
                "maxControls may be at most {MAX_CONTROLS_LIMIT}"
            )));
        }
        Ok(Self {
            stories,
            max_controls: max_controls as usize,
            max_bytes: byte_limit(options.max_bytes)?,
        })
    }
}

impl ContentControlQuery {
    pub(crate) fn matches(&self, control: &ContentControl) -> bool {
        match self {
            Self::Id { control_id } => &control.metadata.control_id == control_id,
            Self::Tag { tag } => control.metadata.tag.as_deref() == Some(tag.as_str()),
            Self::Alias { alias } => control.metadata.alias.as_deref() == Some(alias.as_str()),
        }
    }
}

fn read_doc(
    doc: &EditingDoc,
    options: &ContentControlsOptions,
    scope: AnchorScope,
    query: Option<&ContentControlQuery>,
) -> Result<ContentControlsSnapshot, ExportFailure> {
    let resolved = Resolved::new(options)?;
    let txn = doc.yrs_doc().transact();
    if txn
        .get_map(crate::STORIES)
        .is_none_or(|stories| stories.len(&txn) == 0)
    {
        return Err(ExportFailure {
            code: ExportFailureCode::Unsupported,
            target: None,
            message: "The session holds no document content to read.".to_owned(),
        });
    }
    let inventory = Inventory::build(doc, &txn)?;
    inventory.snapshot(&resolved, scope, |record| {
        query.is_none_or(|query| query.matches(&record.control))
    })
}

impl EditingDoc {
    /// Lists the content controls of the committed state with the version they were read at.
    /// Control ids and anchors are scoped to that version. Nothing is mutated or minted.
    pub fn list_content_controls(
        &self,
        options: &ContentControlsOptions,
    ) -> Result<ExportRead<ContentControlsSnapshot>, ExportRefusal> {
        self.read_controls(options, None)
    }

    /// The controls [`EditingDoc::list_content_controls`] lists that match `query` exactly. Zero
    /// and several matches are both results.
    pub fn find_content_controls(
        &self,
        query: &ContentControlQuery,
        options: &ContentControlsOptions,
    ) -> Result<ExportRead<ContentControlsSnapshot>, ExportRefusal> {
        self.read_controls(options, Some(query))
    }

    fn read_controls(
        &self,
        options: &ContentControlsOptions,
        query: Option<&ContentControlQuery>,
    ) -> Result<ExportRead<ContentControlsSnapshot>, ExportRefusal> {
        let version = self.version();
        match read_doc(self, options, AnchorScope::Session, query) {
            Ok(content) => Ok(ExportRead { version, content }),
            Err(failure) => Err(ExportRefusal { version, failure }),
        }
    }
}

fn snapshot(
    envelope: docx_parse::S9WireEnvelope,
    parts: &SourceParts,
    options: &ContentControlsOptions,
    query: Option<&ContentControlQuery>,
) -> Result<ContentControlsSnapshot, ExportError> {
    Resolved::new(options)?;
    let doc = EditingDoc::new(SNAPSHOT_CLIENT_ID);
    crate::seed::seed_parsed_docx_with(&doc, envelope, Some(parts)).map_err(ExportError::Parse)?;
    Ok(read_doc(&doc, options, AnchorScope::Snapshot, query)?)
}

/// Lists the content controls of DOCX bytes as a snapshot: ids and anchors address the returned
/// content only.
pub fn list_docx_content_controls(
    bytes: &[u8],
    options: &ContentControlsOptions,
) -> Result<ContentControlsSnapshot, ExportError> {
    Resolved::new(options)?;
    let (envelope, parts) =
        crate::seed::parse_docx_with_parts(bytes).map_err(ExportError::Parse)?;
    snapshot(envelope, &parts, options, None)
}

/// The controls of DOCX bytes that match `query` exactly.
pub fn find_docx_content_controls(
    bytes: &[u8],
    query: &ContentControlQuery,
    options: &ContentControlsOptions,
) -> Result<ContentControlsSnapshot, ExportError> {
    Resolved::new(options)?;
    let (envelope, parts) =
        crate::seed::parse_docx_with_parts(bytes).map_err(ExportError::Parse)?;
    snapshot(envelope, &parts, options, Some(query))
}

/// Lists the content controls of a parsed package, reading its current content. `parts` are the
/// inflated parts of the DOCX it was parsed from.
pub fn list_package_content_controls(
    envelope: docx_parse::S9WireEnvelope,
    parts: &[(String, Vec<u8>)],
    options: &ContentControlsOptions,
) -> Result<ContentControlsSnapshot, ExportError> {
    snapshot(envelope, &SourceParts::new(parts.to_vec()), options, None)
}

/// The controls of a parsed package that match `query` exactly.
pub fn find_package_content_controls(
    envelope: docx_parse::S9WireEnvelope,
    parts: &[(String, Vec<u8>)],
    query: &ContentControlQuery,
    options: &ContentControlsOptions,
) -> Result<ContentControlsSnapshot, ExportError> {
    snapshot(
        envelope,
        &SourceParts::new(parts.to_vec()),
        options,
        Some(query),
    )
}
