//! The one walker behind every structured export. It reads committed editing-stream state
//! through the batch projections and the retained source context, classifies headings and lists,
//! projects revisions, records diagnostics and admits root blocks one at a time within the budget.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::rc::Rc;
use std::sync::Arc;

use serde::Serialize;
use serde_json::Value;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Transact};

use super::source::{
    COMMENTS_PART, ENDNOTES_PART, FOOTNOTES_PART, InlineSource, Provenance, ReadSource,
    RelationshipTarget, SourceComment, SourceMerge, TableLayout, Witness, cell_table, move_key,
    story_root,
};
use super::*;
use crate::heading::{OutlineInputs, current_inputs, resolve_heading};
use crate::list_marker::{
    ListState, Unrenderable, list_level_formats, list_marker, numbering_level, renders_format,
};
use crate::ops::{Chunk, ChunkKind};
use crate::queries::revision_parts;
use crate::read_types::{block_control_id, control_metadata, inline_control_id, nested_control_id};
use crate::seed::{SourceMetadata, numeric_field_instruction};
use crate::target::{EditTextView, StoryView, TextPosition, Views};
use crate::{
    COMMENTS, DEL, EditingDoc, INS, KIND_KEY, TextRange, decode_anchor, map_string, story_ref,
};

/// Nesting levels of tables, block controls and field results the walk descends.
const MAX_DEPTH: usize = 32;
/// Stream units one export visits before it stops.
const MAX_VISITED: usize = 4_000_000;
/// Diagnostics one export returns.
const MAX_DIAGNOSTICS: usize = 1_000;
/// Diagnostics of one code an export returns before it summarizes the rest.
const MAX_DIAGNOSTICS_PER_CODE: usize = 100;
/// The client id of the private documents field results and comment bodies are read through.
const SCRATCH_CLIENT_ID: u64 = 1;

/// A lower bound of the serialized size of one inline beyond its text.
const INLINE_OVERHEAD: usize = 48;

const TRUNCATED_MESSAGE: &str =
    "The export stopped at its limits; content after this point is not included.";
const OVERSIZED_MESSAGE: &str = "The export stopped before a story too large to project within maxBytes; raise maxBytes to include it.";
/// Stream units of one story projected per byte of remaining output budget, beyond
/// [`PROJECTION_SLACK`]. A story past that is never projected: the export stops before it.
const PROJECTION_RATIO: usize = 4;
const PROJECTION_SLACK: usize = 1 << 20;
const SUMMARY_MESSAGE: &str = "Further diagnostics with this code were omitted.";

type Payload = HashMap<String, Any>;

/// Counts written bytes without keeping them.
struct Counter(usize);

impl std::io::Write for Counter {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.0 += bytes.len();
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// The compact JSON size of `value`, measured without building the JSON.
fn json_len<T: Serialize + ?Sized>(value: &T) -> usize {
    let mut counter = Counter(0);
    serde_json::to_writer(&mut counter, value).map_or(0, |_| counter.0)
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

fn present(value: Option<&Any>) -> Option<&Any> {
    value.filter(|value| !matches!(value, Any::Null | Any::Undefined))
}

fn any_text(value: Option<&Any>) -> Option<String> {
    match present(value)? {
        Any::String(value) => Some(value.to_string()),
        Any::Number(value) if value.is_finite() && value.fract() == 0.0 => {
            Some(format!("{value:.0}"))
        }
        Any::Number(value) => Some(value.to_string()),
        Any::BigInt(value) => Some(value.to_string()),
        _ => None,
    }
}

fn any_str(value: Option<&Any>) -> Option<&str> {
    match present(value)? {
        Any::String(value) => Some(value),
        _ => None,
    }
}

fn any_number(value: Option<&Any>) -> Option<f64> {
    match present(value)? {
        Any::Number(value) => Some(*value),
        Any::BigInt(value) => Some(*value as f64),
        Any::String(value) => value.parse().ok(),
        _ => None,
    }
}

/// A toggle the way the render bridge reads it: a map is on unless it says `enabled: false`.
fn any_flag(value: Option<&Any>) -> bool {
    match present(value) {
        Some(Any::Bool(value)) => *value,
        Some(Any::Number(value)) => *value != 0.0,
        Some(Any::BigInt(value)) => *value != 0,
        Some(Any::Map(map)) => !matches!(map.get("enabled"), Some(Any::Bool(false))),
        Some(_) => true,
        None => false,
    }
}

fn any_map(value: Option<&Any>) -> Option<&HashMap<String, Any>> {
    match present(value)? {
        Any::Map(map) => Some(map),
        _ => None,
    }
}

fn payload_of<T: ReadTxn>(map: &MapRef, txn: &T) -> Payload {
    map.iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_string(), value)),
            _ => None,
        })
        .collect()
}

fn marks_of(attr: impl Fn(&str) -> Option<Any>) -> Vec<FormattingMark> {
    let mut marks = Vec::new();
    if any_flag(attr("bold").as_ref()) {
        marks.push(FormattingMark::Bold);
    }
    if any_flag(attr("italic").as_ref()) {
        marks.push(FormattingMark::Italic);
    }
    match attr("underline") {
        Some(Any::Map(map)) => {
            let style = any_text(map.get("style")).unwrap_or_else(|| "single".to_owned());
            if style != "none" {
                marks.push(FormattingMark::Underline { style });
            }
        }
        value if any_flag(value.as_ref()) => marks.push(FormattingMark::Underline {
            style: "single".to_owned(),
        }),
        _ => {}
    }
    if any_flag(attr("strike").as_ref()) {
        marks.push(FormattingMark::Strike);
    }
    if any_flag(attr("subscript").as_ref()) {
        marks.push(FormattingMark::Subscript);
    }
    if any_flag(attr("superscript").as_ref()) {
        marks.push(FormattingMark::Superscript);
    }
    if any_flag(attr("hidden").as_ref()) {
        marks.push(FormattingMark::Hidden);
    }
    marks
}

fn link_of(value: Option<&Any>) -> Option<Link> {
    match present(value)? {
        Any::String(href) if !href.is_empty() => Some(Link {
            href: href.to_string(),
            title: None,
        }),
        Any::Map(map) => Some(Link {
            href: any_text(map.get("href")).filter(|href| !href.is_empty())?,
            title: any_text(map.get("tooltip")),
        }),
        _ => None,
    }
}

/// The attribution of a tracked insertion (`inserted`) or deletion: a `moveTo` or `moveFrom`
/// when the source recorded the stamp as a move.
fn revision_of(value: &Any, inserted: bool, moves: Option<&HashSet<String>>) -> Revision {
    let (id, author, date) = revision_parts(value).unwrap_or_default();
    let moved = moves.is_some_and(|moves| moves.contains(&move_key(inserted, &id, &author, &date)));
    let set = |value: String| (!value.is_empty()).then_some(value);
    Revision {
        kind: match (inserted, moved) {
            (true, false) => RevisionKind::Insertion,
            (true, true) => RevisionKind::MoveTo,
            (false, false) => RevisionKind::Deletion,
            (false, true) => RevisionKind::MoveFrom,
        },
        id: set(id),
        author: set(author),
        date: set(date),
    }
}

/// Whether stream content is a pending insertion and a pending deletion, and its attribution.
fn revisions_of(
    attrs: impl Fn(&str) -> Option<Any>,
    moves: Option<&HashSet<String>>,
) -> (bool, bool, Vec<Revision>) {
    let active = |key: &str| attrs(key).filter(|value| present(Some(value)).is_some());
    let (ins, del) = (active(INS), active(DEL));
    let revisions = ins
        .iter()
        .map(|value| revision_of(value, true, moves))
        .chain(del.iter().map(|value| revision_of(value, false, moves)))
        .collect();
    (ins.is_some(), del.is_some(), revisions)
}

fn section_break_type(value: Option<&str>) -> SectionBreakType {
    match value {
        Some("continuous") => SectionBreakType::Continuous,
        Some("nextColumn") => SectionBreakType::NextColumn,
        Some("oddPage") => SectionBreakType::OddPage,
        Some("evenPage") => SectionBreakType::EvenPage,
        _ => SectionBreakType::NextPage,
    }
}

fn count_blocks(blocks: &[Block]) -> usize {
    blocks.iter().map(count_block).sum()
}

fn count_block(block: &Block) -> usize {
    1 + match &block.content {
        BlockKind::Table { table } => table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .map(|cell| count_blocks(&cell.blocks))
            .sum(),
        BlockKind::ContentControl { blocks, .. } => count_blocks(blocks),
        BlockKind::Paragraph { paragraph }
        | BlockKind::Heading { paragraph, .. }
        | BlockKind::ListItem { paragraph, .. } => count_inline_blocks(&paragraph.inlines),
        _ => 0,
    }
}

fn count_inline_blocks(inlines: &[Inline]) -> usize {
    inlines
        .iter()
        .map(|inline| match &inline.content {
            InlineKind::Field {
                cached_result: CachedResult::Blocks { blocks },
                ..
            } => count_blocks(blocks),
            InlineKind::Field {
                cached_result: CachedResult::Inline { inlines },
                ..
            }
            | InlineKind::ContentControl { inlines, .. } => count_inline_blocks(inlines),
            _ => 0,
        })
        .sum()
}

/// Gives every node below `blocks` its export-tree id.
fn number_blocks(blocks: &mut [Block], prefix: &str) {
    for (index, block) in blocks.iter_mut().enumerate() {
        number_block(block, format!("{prefix}.b{index}"));
    }
}

fn number_block(block: &mut Block, id: String) {
    match &mut block.content {
        BlockKind::Paragraph { paragraph }
        | BlockKind::Heading { paragraph, .. }
        | BlockKind::ListItem { paragraph, .. } => number_inlines(&mut paragraph.inlines, &id),
        BlockKind::Table { table } => {
            for (row_index, row) in table.rows.iter_mut().enumerate() {
                for (cell_index, cell) in row.cells.iter_mut().enumerate() {
                    number_blocks(
                        &mut cell.blocks,
                        &format!("{id}.r{row_index}.c{cell_index}"),
                    );
                }
            }
        }
        BlockKind::ContentControl { blocks, .. } => number_blocks(blocks, &id),
        _ => {}
    }
    block.id = id;
}

fn number_inlines(inlines: &mut [Inline], prefix: &str) {
    for (index, inline) in inlines.iter_mut().enumerate() {
        let id = format!("{prefix}.i{index}");
        match &mut inline.content {
            InlineKind::ContentControl { inlines, .. } => number_inlines(inlines, &id),
            InlineKind::Field { cached_result, .. } => match cached_result {
                CachedResult::Inline { inlines } => number_inlines(inlines, &id),
                CachedResult::Blocks { blocks } => number_blocks(blocks, &id),
                CachedResult::Missing => {}
            },
            _ => {}
        }
        inline.id = id;
    }
}

/// Joins adjacent text of one attribution whose ranges continue each other, so the stream's
/// chunking never shows.
fn canonicalize(inlines: Vec<Inline>) -> Vec<Inline> {
    let mut output: Vec<Inline> = Vec::with_capacity(inlines.len());
    for inline in inlines {
        if let Some(previous) = output.last_mut()
            && let (InlineKind::Text { text: left }, InlineKind::Text { text: right }) =
                (&mut previous.content, &inline.content)
            && previous.marks == inline.marks
            && previous.link == inline.link
            && previous.revisions == inline.revisions
        {
            match (&mut previous.anchor, &inline.anchor) {
                (Anchor::Range(before), Anchor::Range(after))
                    if before.story == after.story
                        && before.view == after.view
                        && before.end == after.start =>
                {
                    before.end = after.end.clone();
                }
                (before, after) if !matches!(before, Anchor::Range(_)) && before == after => {}
                _ => {
                    output.push(inline);
                    continue;
                }
            }
            left.push_str(right);
            continue;
        }
        output.push(inline);
    }
    output
}

/// Splits `text` after `units` UTF-16 units, at the nearest character boundary.
fn split_utf16(text: &str, units: u32) -> (&str, &str) {
    let mut counted = 0u32;
    for (byte, ch) in text.char_indices() {
        if counted >= units {
            return text.split_at(byte);
        }
        counted += ch.len_utf16() as u32;
    }
    (text, "")
}

/// One section of the body: the paragraph ending it, its start and its properties.
struct Section {
    end: Option<String>,
    start: Option<String>,
    properties: Option<docx_parse::SectionProperties>,
}

/// Where a story's provenance comes from: the retained source, or the private document a
/// comment body or field result was lowered into.
#[derive(Clone)]
enum Prov {
    Main,
    Scratch(Rc<ScratchProv>),
}

struct ScratchProv {
    provenance: Provenance,
    /// Source provenance of each story's raw blocks, when the content is the package's.
    raw_anchors: HashMap<String, Vec<Option<Anchor>>>,
}

/// Inline source content placed back into a paragraph.
#[derive(Clone)]
enum Record {
    Omitted {
        element: String,
        in_control: bool,
    },
    Break {
        kind: BreakType,
        revision: Option<Revision>,
        in_control: bool,
        /// Its UTF-16 offsets into the content control's content and each nested control's,
        /// when inside one.
        control_offset: Option<Vec<u32>>,
    },
}

impl Record {
    fn in_control(&self) -> bool {
        match self {
            Self::Omitted { in_control, .. } | Self::Break { in_control, .. } => *in_control,
        }
    }

    /// The record's offset into the control content being read.
    fn control_offset(&self) -> Option<u32> {
        match self {
            Self::Break { control_offset, .. } => control_offset.as_ref()?.first().copied(),
            Self::Omitted { .. } => None,
        }
    }

    /// Whether the record belongs to a control nested at its offset.
    fn nested(&self) -> bool {
        matches!(self, Self::Break { control_offset: Some(path), .. } if path.len() > 1)
    }

    /// The record as the nested control at its offset reads it.
    fn descend(self) -> Self {
        match self {
            Self::Break {
                kind,
                revision,
                in_control,
                control_offset: Some(path),
            } => Self::Break {
                kind,
                revision,
                in_control,
                control_offset: Some(path[1..].to_vec()),
            },
            record => record,
        }
    }
}

/// Where the source content of one story that the stream leaves out goes back in.
#[derive(Default)]
struct Placements {
    /// Raw XML block elements with their provenance, keyed by the block they precede.
    before: HashMap<String, Vec<(String, Option<Anchor>)>>,
    /// Inline records per paragraph, at the story index they precede.
    inline: HashMap<String, Vec<(u32, Record)>>,
    /// Story indices of break embeds seeding moved out of paragraphs that still stand for them.
    relocated: HashSet<u32>,
}

/// Per-story state of the walk.
struct StoryCtx {
    story: String,
    /// The package part the story belongs to.
    part: Option<String>,
    prov: Prov,
    accepted: Rc<StoryView>,
    /// The original-view projection; the accepted view, which never reads it, reuses `accepted`.
    original: Rc<StoryView>,
    chunks: Arc<Vec<Chunk>>,
    /// The anchor every node uses instead of its own, for content without a session location.
    owner: Option<Anchor>,
    /// The same for the block being built: a field's anchor for its result blocks, no location
    /// for a paragraph whose id another shares.
    block_owner: Option<Anchor>,
    /// Block ids a numeric field hides as its cached result: the field's paragraph and ordinal.
    hidden: HashMap<String, (usize, usize)>,
    /// Paragraphs owning hidden result blocks.
    owners: HashSet<usize>,
    /// The anchors of the fields owning hidden blocks, once built.
    field_anchors: HashMap<(usize, usize), Anchor>,
    placements: Placements,
    duplicates: HashSet<String>,
    /// The style of the paragraph being built, which its fields' results inherit.
    paragraph_style: Option<String>,
    /// Per paragraph: its mark is a pending insertion, a pending deletion.
    marks: Vec<(bool, bool)>,
    /// When the projection holds only the story's leading paragraphs, the paragraph and block
    /// ids after them.
    beyond: Option<HashSet<String>>,
    /// Where the projection ends: after the leading paragraphs and the complete blocks that
    /// follow them, or `u32::MAX` for the whole story.
    cut: u32,
}

impl StoryCtx {
    fn owner(&self) -> Option<&Anchor> {
        self.block_owner.as_ref().or(self.owner.as_ref())
    }

    fn anchor(&self, anchor: Anchor) -> Anchor {
        self.owner().cloned().unwrap_or(anchor)
    }

    /// A paragraph's anchor; one whose id another paragraph shares has no location.
    fn paragraph_anchor(&self, para_id: &str) -> Anchor {
        self.anchor(if self.duplicates.contains(para_id) {
            self.unlocated(UnlocatedReason::DuplicateParagraphId)
        } else {
            Anchor::Paragraph {
                story: self.story.clone(),
                para_id: para_id.to_owned(),
            }
        })
    }

    fn unlocated(&self, reason: UnlocatedReason) -> Anchor {
        Anchor::Unlocated {
            story: self.story.clone(),
            reason,
        }
    }

    /// A child story name, when it names a story of the session.
    fn session_story(&self, story: String) -> Option<String> {
        matches!(self.prov, Prov::Main).then_some(story)
    }
}

/// A block of a story with the diagnostics it raised and, for a paragraph, its index.
struct Built {
    block: Block,
    diagnostics: Vec<Diagnostic>,
    paragraph: Option<usize>,
}

/// A paragraph held back until the result blocks its numeric fields hide have been attached.
struct Held {
    built: Built,
    /// Blocks the stream places after it before those results, such as its section break.
    trailing: Vec<Built>,
}

/// A comment body the export reads, converted only once the export reaches it.
enum CommentBody<'a> {
    /// Blocks the comment store holds.
    Stored(Arc<[Any]>),
    /// Plain text the comment store holds.
    Text(String),
    /// The package's body, whose source provenance applies.
    Source(&'a [Value]),
    Empty,
}

impl CommentBody<'_> {
    fn is_empty(&self) -> bool {
        match self {
            Self::Stored(blocks) => blocks.is_empty(),
            Self::Text(_) => false,
            Self::Source(blocks) => blocks.is_empty(),
            Self::Empty => true,
        }
    }

    fn size(&self) -> usize {
        match self {
            Self::Stored(blocks) => json_len(&blocks[..]),
            Self::Text(text) => text.len(),
            Self::Source(blocks) => json_len(blocks),
            Self::Empty => 0,
        }
    }

    fn blocks(&self) -> Vec<Value> {
        match self {
            Self::Stored(blocks) => blocks
                .iter()
                .filter_map(|block| serde_json::to_value(block).ok())
                .collect(),
            Self::Text(text) => vec![serde_json::json!({
                "type": "paragraph",
                "content": [{"type": "run", "content": [{"type": "text", "text": text}]}]
            })],
            Self::Source(blocks) => blocks.to_vec(),
            Self::Empty => Vec::new(),
        }
    }
}

/// A story the export visits.
struct Planned<'a> {
    story: ExportStory,
    comment_body: Option<CommentBody<'a>>,
}

struct Budget {
    max_bytes: usize,
    max_blocks: usize,
    max_visited: usize,
    bytes: usize,
    blocks: usize,
    visited: usize,
}

struct Exporter<'a> {
    options: &'a Resolved,
    source: Option<&'a SourceMetadata>,
    content: DocxStructuredContent,
    budget: Budget,
    /// The root subtree under construction: its blocks and a lower bound of its size.
    building: (usize, usize),
    /// Diagnostics raised since the last commit.
    pending: Vec<Diagnostic>,
    /// Diagnostics raised since the last commit, per code, including those not kept.
    pending_counts: BTreeMap<DiagnosticCode, usize>,
    counts: BTreeMap<DiagnosticCode, usize>,
    stopped: bool,
    excluded_revision: bool,
    sections: Vec<Section>,
    /// The index in `sections` of the section each paragraph ends.
    section_ends: HashMap<String, usize>,
    story_ids: BTreeSet<String>,
    comment_ids: BTreeSet<String>,
    /// Private documents for comment bodies and field results, one per walk reading at once.
    scratch: Vec<Rc<EditingDoc>>,
    /// How many private documents a walk is reading now: the first that many, which seeding
    /// never writes to.
    reading: usize,
    scratch_stories: usize,
    /// The stream content projected out of the field being built, by result index.
    projected_result: Option<Vec<(i64, Vec<Inline>)>>,
    /// The export stopped before a story too large to project.
    oversized: bool,
    /// Blocks under construction that only wrap a field result and are never output.
    wrappers: usize,
}

/// Exports `doc`. A walk that has to stop is repeated with room kept for the truncation
/// diagnostic, so content that fits exactly is never truncated.
pub(super) fn export(
    doc: &EditingDoc,
    options: &Resolved,
    scope: AnchorScope,
) -> DocxStructuredContent {
    let content = walk(doc, options, scope, false);
    if content.truncated {
        walk(doc, options, scope, true)
    } else {
        content
    }
}

/// Whether `doc` holds any story to export.
pub(super) fn has_stories(doc: &EditingDoc) -> bool {
    let txn = doc.yrs_doc().transact();
    txn.get_map(crate::STORIES)
        .is_some_and(|stories| stories.len(&txn) > 0)
}

fn walk(
    doc: &EditingDoc,
    options: &Resolved,
    scope: AnchorScope,
    reserve: bool,
) -> DocxStructuredContent {
    let source = doc.source_metadata();
    walk_with(
        doc,
        Exporter::new(options, source.as_deref(), scope, reserve),
    )
}

fn walk_with(doc: &EditingDoc, mut exporter: Exporter<'_>) -> DocxStructuredContent {
    let txn = doc.yrs_doc().transact();
    exporter.story_ids = txn
        .get_map(crate::STORIES)
        .map(|stories| stories.keys(&txn).map(str::to_owned).collect())
        .unwrap_or_default();
    let mut views = Views::new(doc, &txn);
    exporter.read_sections(&mut views);
    let planned = exporter.plan(&mut views);
    if exporter.commit_notes() {
        exporter.run(&mut views, planned);
    }
    exporter.finish()
}

impl<'a> Exporter<'a> {
    fn new(
        options: &'a Resolved,
        source: Option<&'a SourceMetadata>,
        scope: AnchorScope,
        reserve: bool,
    ) -> Self {
        let content = DocxStructuredContent {
            schema_version: SCHEMA_VERSION,
            revision_view: options.view,
            anchor_scope: scope,
            included_stories: options.stories.clone(),
            include_formatting: options.include_formatting,
            stories: Vec::new(),
            diagnostics: Vec::new(),
            truncated: false,
        };
        let reserve = if reserve {
            json_len(&diagnostic(
                DiagnosticCode::Truncated,
                Severity::Warning,
                None,
                OVERSIZED_MESSAGE,
            )) + 1
        } else {
            0
        };
        Self {
            options,
            source,
            budget: Budget {
                max_bytes: options.max_bytes,
                max_blocks: options.max_blocks,
                max_visited: MAX_VISITED,
                bytes: json_len(&content) + reserve,
                blocks: 0,
                visited: 0,
            },
            content,
            building: (0, 0),
            pending: Vec::new(),
            pending_counts: BTreeMap::new(),
            counts: BTreeMap::new(),
            stopped: false,
            excluded_revision: false,
            sections: Vec::new(),
            section_ends: HashMap::new(),
            story_ids: BTreeSet::new(),
            comment_ids: BTreeSet::new(),
            scratch: Vec::new(),
            reading: 0,
            scratch_stories: 0,
            projected_result: None,
            oversized: false,
            wrappers: 0,
        }
    }

    fn view(&self) -> RevisionView {
        self.options.view
    }

    fn read(&self) -> Option<&'a ReadSource> {
        self.source.map(SourceMetadata::read)
    }

    fn numbering(&self) -> Option<Arc<docx_parse::NumberingMap>> {
        self.source.map(SourceMetadata::numbering)
    }

    /// The move revisions of a story's provenance.
    fn moves<'p>(&self, prov: &'p Prov) -> Option<&'p HashSet<String>>
    where
        'a: 'p,
    {
        match prov {
            Prov::Main => self.read().map(|read| &read.provenance.moves),
            Prov::Scratch(scratch) => Some(&scratch.provenance.moves),
        }
    }

    /// Raises a diagnostic, keeping only as many of one code as the export can return.
    fn note(
        &mut self,
        code: DiagnosticCode,
        severity: Severity,
        anchor: Option<Anchor>,
        message: impl Into<String>,
    ) {
        let raised = self.pending_counts.entry(code).or_default();
        *raised += 1;
        let seen = self.counts.get(&code).copied().unwrap_or_default() + *raised;
        if seen <= MAX_DIAGNOSTICS_PER_CODE
            && self.content.diagnostics.len() + self.pending.len() < MAX_DIAGNOSTICS
        {
            self.pending
                .push(diagnostic(code, severity, anchor, message));
        }
    }

    /// Counts visited stream units; stops the export once it has visited too many.
    fn visit(&mut self, units: usize) -> bool {
        self.budget.visited = self.budget.visited.saturating_add(units);
        if self.budget.visited > self.budget.max_visited {
            self.stop();
        }
        !self.stopped
    }

    /// Adds to the root subtree under construction; stops the export once the subtree cannot
    /// fit, so no subtree is built far past the budget.
    fn grow(&mut self, blocks: usize, bytes: usize) -> bool {
        self.building.0 += blocks;
        self.building.1 = self.building.1.saturating_add(bytes);
        if self.budget.blocks + self.building.0.saturating_sub(self.wrappers)
            > self.budget.max_blocks
            || self.budget.bytes.saturating_add(self.building.1) > self.budget.max_bytes
        {
            self.stop();
        }
        !self.stopped
    }

    /// Accounts for one inline of the subtree under construction.
    fn grow_inline(&mut self, inline: &Inline) -> bool {
        let leaf = match &inline.content {
            InlineKind::ContentControl { .. } | InlineKind::Field { .. } => 0,
            _ => json_len(inline),
        };
        self.grow(0, leaf)
    }

    /// `diagnostics` after the per-code and total caps, with the counts that would result.
    fn admit(
        &self,
        diagnostics: Vec<Diagnostic>,
    ) -> (Vec<Diagnostic>, BTreeMap<DiagnosticCode, usize>) {
        let mut counts = self.counts.clone();
        let mut total = self.content.diagnostics.len();
        let mut admitted = Vec::new();
        for item in diagnostics {
            if total >= MAX_DIAGNOSTICS {
                break;
            }
            let count = counts.entry(item.code).or_default();
            *count += 1;
            if *count < MAX_DIAGNOSTICS_PER_CODE {
                admitted.push(item);
                total += 1;
            } else if *count == MAX_DIAGNOSTICS_PER_CODE {
                admitted.push(diagnostic(item.code, item.severity, None, SUMMARY_MESSAGE));
                total += 1;
            }
        }
        (admitted, counts)
    }

    fn diagnostics_cost(&self, diagnostics: &[Diagnostic]) -> usize {
        let existing = self.content.diagnostics.len();
        diagnostics
            .iter()
            .enumerate()
            .map(|(index, item)| json_len(item) + usize::from(existing + index > 0))
            .sum()
    }

    fn stop(&mut self) {
        self.stopped = true;
        self.content.truncated = true;
    }

    /// Forgets what was raised for a subtree that will not be committed.
    fn reset_building(&mut self) {
        self.building = (0, 0);
        self.pending_counts.clear();
    }

    /// Commits the diagnostics raised outside any block, if they fit.
    fn commit_notes(&mut self) -> bool {
        if self.stopped {
            return false;
        }
        let pending = std::mem::take(&mut self.pending);
        let (admitted, counts) = self.admit(pending);
        let cost = self.diagnostics_cost(&admitted);
        self.reset_building();
        if self.budget.bytes + cost > self.budget.max_bytes {
            self.stop();
            return false;
        }
        self.budget.bytes += cost;
        self.counts = counts;
        self.content.diagnostics.extend(admitted);
        true
    }

    /// Adds a story with no blocks yet, if it fits.
    fn open_story(&mut self, story: ExportStory) -> bool {
        let cost = json_len(&story) + usize::from(!self.content.stories.is_empty());
        if self.budget.bytes + cost > self.budget.max_bytes {
            self.stop();
            return false;
        }
        self.budget.bytes += cost;
        self.content.stories.push(story);
        true
    }

    /// Adds one root block to the last story with its diagnostics, or stops the export.
    fn commit(&mut self, mut built: Built) -> bool {
        if self.stopped {
            return false;
        }
        let story_index = self.content.stories.len() - 1;
        let block_index = self.content.stories[story_index].blocks.len();
        number_block(&mut built.block, format!("s{story_index}.b{block_index}"));
        let blocks = count_block(&built.block);
        let (admitted, counts) = self.admit(built.diagnostics);
        let cost = json_len(&built.block)
            + usize::from(block_index > 0)
            + self.diagnostics_cost(&admitted);
        self.reset_building();
        if self.budget.blocks + blocks > self.budget.max_blocks
            || self.budget.bytes + cost > self.budget.max_bytes
        {
            self.stop();
            return false;
        }
        self.budget.blocks += blocks;
        self.budget.bytes += cost;
        self.counts = counts;
        self.content.diagnostics.extend(admitted);
        self.content.stories[story_index].blocks.push(built.block);
        true
    }

    fn finish(mut self) -> DocxStructuredContent {
        if self.content.truncated {
            self.content.diagnostics.push(diagnostic(
                DiagnosticCode::Truncated,
                Severity::Warning,
                None,
                if self.oversized {
                    OVERSIZED_MESSAGE
                } else {
                    TRUNCATED_MESSAGE
                },
            ));
        }
        self.content
    }

    /// Whether content of `size` stream units or source bytes can be read into the export's
    /// working structures within the remaining budget.
    fn fits(&self, size: usize) -> bool {
        size <= self.capacity()
    }

    fn capacity(&self) -> usize {
        let remaining = self.budget.max_bytes.saturating_sub(self.budget.bytes);
        remaining.saturating_mul(PROJECTION_RATIO) + PROJECTION_SLACK
    }

    /// Stops the export before reading content too large for the remaining budget.
    fn oversize(&mut self) {
        self.oversized = true;
        self.stop();
    }

    fn story_units<T: ReadTxn>(views: &Views<'_, T>, story: &str) -> usize {
        let txn = views.txn();
        story_ref(txn, story).map_or(0, |text| yrs::Text::len(&text, txn) as usize)
    }

    /// Stops the export before a story whose projection would outgrow the remaining budget.
    fn projectable<T: ReadTxn>(&mut self, views: &Views<'_, T>, story: &str) -> bool {
        if self.fits(Self::story_units(views, story)) {
            return true;
        }
        self.oversize();
        false
    }

    /// Lowers `blocks` into a new story of the first private document no walk is reading.
    fn seed_scratch(
        &mut self,
        story: &str,
        blocks: &[Value],
    ) -> Result<(Rc<EditingDoc>, Provenance), String> {
        while self.scratch.len() <= self.reading {
            self.scratch
                .push(Rc::new(EditingDoc::new(SCRATCH_CLIENT_ID)));
        }
        let doc = Rc::clone(&self.scratch[self.reading]);
        let provenance =
            crate::seed::seed_blocks(&doc, self.source, &[(story.to_owned(), blocks)])?;
        Ok((doc, provenance))
    }

    fn run<T: ReadTxn>(&mut self, views: &mut Views<'_, T>, planned: Vec<Planned<'a>>) {
        for planned in planned {
            if self.stopped {
                return;
            }
            let story = planned.story.story.clone();
            let owner = planned.story.comment.as_ref().and_then(|comment| {
                self.read()
                    .and_then(|read| read.comment_anchors.get(&comment.id).cloned())
                    .or_else(|| comment.anchors.first().cloned())
            });
            if !self.open_story(planned.story) {
                return;
            }
            match planned.comment_body {
                None => {
                    let part = self.read().and_then(|read| read.story_part(&story));
                    self.root_story(views, &story, Prov::Main, None, part);
                }
                Some(body) => self.comment_story(&story, body, owner),
            }
        }
        if self.excluded_revision {
            let message = match self.view() {
                RevisionView::Original => "Pending insertions are excluded from the original view.",
                _ => "Pending deletions are excluded from the accepted view.",
            };
            self.note(
                DiagnosticCode::RevisionContentExcluded,
                Severity::Info,
                None,
                message,
            );
            self.commit_notes();
        }
    }

    /// Walks a comment body lowered into a private document, every node anchored to the comment.
    fn comment_story(&mut self, story: &str, body: CommentBody<'a>, owner: Option<Anchor>) {
        if body.is_empty() {
            return;
        }
        let Some(owner) = owner else {
            self.note(
                DiagnosticCode::ProvenanceUnavailable,
                Severity::Info,
                None,
                format!("Comment story {story} has no source element and annotates no remaining text, so its body is not exported."),
            );
            self.commit_notes();
            return;
        };
        if !self.fits(body.size()) {
            self.oversize();
            return;
        }
        let raw_anchors = match (matches!(body, CommentBody::Source(_)), self.read()) {
            (true, Some(read)) => read
                .raw_block_anchors
                .iter()
                .filter(|(key, _)| story_root(key) == story)
                .map(|(key, anchors)| (key.clone(), anchors.clone()))
                .collect(),
            _ => HashMap::new(),
        };
        match self.seed_scratch(story, &body.blocks()) {
            Ok((doc, provenance)) => {
                let prov = Prov::Scratch(Rc::new(ScratchProv {
                    provenance,
                    raw_anchors,
                }));
                self.reading += 1;
                {
                    let txn = doc.yrs_doc().transact();
                    let mut views = Views::new(&doc, &txn);
                    self.root_story(
                        &mut views,
                        story,
                        prov,
                        Some(owner),
                        Some(COMMENTS_PART.to_owned()),
                    );
                }
                self.reading -= 1;
            }
            Err(_) => {
                self.note(
                    DiagnosticCode::UnsupportedContent,
                    Severity::Warning,
                    Some(owner),
                    format!("The body of comment story {story} could not be read."),
                );
                self.commit_notes();
            }
        }
    }

    /// Reads the body's sections and indexes them by the paragraph ending each.
    fn read_sections<T: ReadTxn>(&mut self, views: &mut Views<'_, T>) {
        self.sections = self.body_sections(views);
        self.section_ends = HashMap::new();
        for (index, section) in self.sections.iter().enumerate() {
            if let Some(end) = &section.end {
                self.section_ends.entry(end.clone()).or_insert(index);
            }
        }
    }

    /// The body's sections in order, each with the paragraph ending it, then the final section.
    fn body_sections<T: ReadTxn>(&mut self, views: &mut Views<'_, T>) -> Vec<Section> {
        let mut sections = Vec::new();
        let txn = views.txn();
        if let Ok(text) = story_ref(txn, "body") {
            let chunks = views.doc().chunk_snapshot("body", &text, txn);
            for chunk in chunks.iter() {
                let ChunkKind::Pilcrow(map) = &chunk.kind else {
                    continue;
                };
                let value = |key: &str| match map.get(txn, key) {
                    Some(Out::Any(value)) if !matches!(value, Any::Null | Any::Undefined) => {
                        Some(value)
                    }
                    _ => None,
                };
                let (section, break_type) = (value("sectPr"), value("sectionBreakType"));
                if section.is_none() && break_type.is_none() {
                    continue;
                }
                let properties = section
                    .filter(|value| self.fits(json_len(value)))
                    .and_then(|value| serde_json::to_value(value).ok())
                    .and_then(|value| {
                        serde_json::from_value::<docx_parse::SectionProperties>(value).ok()
                    });
                let start = properties
                    .as_ref()
                    .and_then(|section| section.section_start.clone())
                    .or_else(|| any_text(break_type.as_ref()));
                sections.push(Section {
                    end: Some(map_string(map, txn, crate::PARA_ID).unwrap_or_default()),
                    start,
                    properties,
                });
            }
        }
        let last = self
            .read()
            .and_then(|read| read.final_section.clone())
            .and_then(|value| serde_json::from_value::<docx_parse::SectionProperties>(value).ok());
        sections.push(Section {
            end: None,
            start: last
                .as_ref()
                .and_then(|section| section.section_start.clone()),
            properties: last,
        });
        sections
    }

    /// The sections referencing each header and footer story, with inherited references.
    fn story_uses(&self) -> HashMap<String, Vec<StoryUse>> {
        let mut properties: Vec<docx_parse::SectionProperties> = self
            .sections
            .iter()
            .map(|section| section.properties.clone().unwrap_or_default())
            .collect();
        docx_parse::apply_section_inheritance(&mut properties);
        let mut uses: HashMap<String, Vec<StoryUse>> = HashMap::new();
        let mut seen: HashSet<(String, u32, u8)> = HashSet::new();
        for (index, section) in properties.iter().enumerate() {
            for reference in section
                .header_references
                .iter()
                .chain(section.footer_references.iter())
                .flatten()
            {
                let variant = match reference.reference_type.as_str() {
                    "first" => HeaderFooterVariant::First,
                    "even" => HeaderFooterVariant::Even,
                    _ => HeaderFooterVariant::Default,
                };
                let story = format!("hf:{}", reference.relationship_id);
                if seen.insert((story.clone(), index as u32, variant as u8)) {
                    uses.entry(story).or_default().push(StoryUse {
                        section_index: index as u32,
                        variant,
                    });
                }
            }
        }
        uses
    }

    /// The stories to export in output order, plus the envelope diagnostics.
    fn plan<T: ReadTxn>(&mut self, views: &mut Views<'_, T>) -> Vec<Planned<'a>> {
        let selected: BTreeSet<StorySelection> = self.options.stories.iter().copied().collect();
        if !self.options.include_formatting {
            self.note(
                DiagnosticCode::FormattingOmitted,
                Severity::Info,
                None,
                "Formatting marks are omitted from this export.",
            );
        }
        match self.read() {
            None => self.note(
                DiagnosticCode::ProvenanceUnavailable,
                Severity::Info,
                None,
                "The document was not opened from DOCX bytes, so styles, numbering definitions, header and footer roles, comment metadata, source grids and omitted source content are unavailable.",
            ),
            Some(read) => {
                for warning in &read.warnings {
                    self.note(
                        DiagnosticCode::ParseWarning,
                        Severity::Warning,
                        None,
                        warning.clone(),
                    );
                }
                if !read.pinned && !read.provenance.inline.is_empty() {
                    self.note(
                        DiagnosticCode::ProvenanceUnavailable,
                        Severity::Info,
                        None,
                        "The stories were not seeded from the opened package here, so omitted inline source content and source break positions cannot be located.",
                    );
                }
            }
        }
        let mut members: BTreeMap<StorySelection, usize> = BTreeMap::new();
        let mut planned = Vec::new();
        if self.story_ids.contains("body") {
            members.insert(StorySelection::Body, 1);
            if selected.contains(&StorySelection::Body) {
                planned.push(Planned {
                    story: ExportStory {
                        story: "body".to_owned(),
                        kind: StoryKind::Body,
                        part: self.read().map(|read| read.document_part.clone()),
                        note_id: None,
                        comment: None,
                        uses: Vec::new(),
                        blocks: Vec::new(),
                    },
                    comment_body: None,
                });
            }
        }
        let uses = self.story_uses();
        let mut stories: Vec<(StorySelection, ExportStory)> = Vec::new();
        let mut known: HashSet<&str> = HashSet::new();
        let mut unclassified = 0usize;
        if let Some(read) = self.read() {
            let mut by_part: HashMap<String, usize> = HashMap::new();
            for source in &read.stories {
                if !self.story_ids.contains(&source.story) {
                    continue;
                }
                known.insert(&source.story);
                let category = match source.kind {
                    StoryKind::Header => StorySelection::Headers,
                    StoryKind::Footer => StorySelection::Footers,
                    StoryKind::Footnote => StorySelection::Footnotes,
                    _ => StorySelection::Endnotes,
                };
                let story_uses = uses.get(&source.story).cloned().unwrap_or_default();
                let shared = matches!(source.kind, StoryKind::Header | StoryKind::Footer)
                    .then(|| source.part.clone())
                    .flatten()
                    .and_then(|part| by_part.get(&part).copied().map(|index| (part, index)));
                if let Some((part, index)) = shared {
                    let first = stories[index].1.story.clone();
                    let capacity = self.capacity();
                    if same_content(views, &self.story_ids, capacity, &first, &source.story) {
                        self.note(
                            DiagnosticCode::AmbiguousIdentity,
                            Severity::Info,
                            None,
                            format!("Story {} reads part {part} as story {first} does and holds the same content, so it is exported once as {first}.", source.story),
                        );
                        let merged = &mut stories[index].1.uses;
                        merged.extend(story_uses);
                        merged.sort_by_key(|used| (used.section_index, used.variant as u8));
                        merged.dedup();
                        continue;
                    }
                    self.note(
                        DiagnosticCode::AmbiguousIdentity,
                        Severity::Warning,
                        None,
                        format!("Stories {first} and {} both read part {part} but now hold different content, so both are exported.", source.story),
                    );
                }
                if let Some(part) = source
                    .part
                    .clone()
                    .filter(|_| matches!(source.kind, StoryKind::Header | StoryKind::Footer))
                {
                    by_part.entry(part).or_insert(stories.len());
                }
                stories.push((
                    category,
                    ExportStory {
                        story: source.story.clone(),
                        kind: source.kind,
                        part: source.part.clone(),
                        note_id: source.note_id.clone(),
                        comment: None,
                        uses: story_uses,
                        blocks: Vec::new(),
                    },
                ));
            }
        }
        let mut extra = Vec::new();
        for story in &self.story_ids {
            if known.contains(story.as_str()) || story_root(story) != story {
                continue;
            }
            let (category, kind, part) = if story.starts_with("fn:") {
                (
                    StorySelection::Footnotes,
                    StoryKind::Footnote,
                    FOOTNOTES_PART,
                )
            } else if story.starts_with("en:") {
                (StorySelection::Endnotes, StoryKind::Endnote, ENDNOTES_PART)
            } else {
                unclassified += usize::from(story.starts_with("hf:"));
                continue;
            };
            extra.push((
                category,
                ExportStory {
                    story: story.clone(),
                    kind,
                    part: self.read().map(|_| part.to_owned()),
                    note_id: Some(story[3..].to_owned()),
                    comment: None,
                    uses: Vec::new(),
                    blocks: Vec::new(),
                },
            ));
        }
        extra.sort_by_key(|(category, story)| {
            let id = story.note_id.clone().unwrap_or_default();
            (*category, id.parse::<i64>().ok(), id)
        });
        stories.extend(extra);
        stories.sort_by_key(|(category, _)| *category);
        for (category, story) in stories {
            *members.entry(category).or_default() += 1;
            if selected.contains(&category) {
                planned.push(Planned {
                    story,
                    comment_body: None,
                });
            }
        }
        if unclassified > 0
            && (selected.contains(&StorySelection::Headers)
                || selected.contains(&StorySelection::Footers))
        {
            self.note(
                DiagnosticCode::ProvenanceUnavailable,
                Severity::Warning,
                None,
                format!("{unclassified} header or footer stories cannot be classified without the source package and are not exported."),
            );
        }
        let comments = self.plan_comments(views, selected.contains(&StorySelection::Comments));
        if !comments.is_empty() {
            members.insert(StorySelection::Comments, comments.len());
        }
        if selected.contains(&StorySelection::Comments) {
            planned.extend(comments);
        }
        for (category, count) in members {
            if selected.contains(&category) {
                continue;
            }
            let option = match category {
                StorySelection::Body => "body",
                StorySelection::Headers => "headers",
                StorySelection::Footers => "footers",
                StorySelection::Footnotes => "footnotes",
                StorySelection::Endnotes => "endnotes",
                StorySelection::Comments => "comments",
            };
            self.note(
                DiagnosticCode::StoriesOmitted,
                Severity::Info,
                None,
                format!("{count} {option} stories are not included; select \"{option}\" to include them."),
            );
        }
        if let Some(read) = self.read() {
            for (category, count, noun) in [
                (
                    StorySelection::Footnotes,
                    read.footnote_separators,
                    "footnote",
                ),
                (StorySelection::Endnotes, read.endnote_separators, "endnote"),
            ] {
                if count > 0 && selected.contains(&category) {
                    self.note(
                        DiagnosticCode::StoriesOmitted,
                        Severity::Info,
                        None,
                        format!("{count} {noun} separator notes are not exported."),
                    );
                }
            }
        }
        planned
    }

    /// Comments from the session's comment store in source order and then by id. A field of a
    /// source comment reads the store once it has been written there since the source was
    /// retained, and the source until then.
    fn plan_comments<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        locate: bool,
    ) -> Vec<Planned<'a>> {
        let txn = views.txn();
        let mut stored: BTreeMap<String, (Payload, Vec<(String, u32, u32)>)> = BTreeMap::new();
        if let Some(comments) = txn.get_map(COMMENTS) {
            for (id, value) in comments.iter(txn) {
                let Ok(comment) = value.cast::<MapRef>() else {
                    continue;
                };
                let values = payload_of(&comment, txn);
                let ranges = match values.get("anchors").filter(|_| locate) {
                    Some(Any::Array(anchors)) => anchors
                        .iter()
                        .filter_map(|anchor| decode_anchor(anchor).ok())
                        .filter_map(|anchor| {
                            let start = anchor.start.get_offset(txn)?.index;
                            let end = anchor.end.get_offset(txn)?.index;
                            Some((anchor.story, start, end))
                        })
                        .collect(),
                    _ => Vec::new(),
                };
                stored.insert(id.to_owned(), (values, ranges));
            }
        }
        let mut order: Vec<String> = Vec::new();
        let mut sources: HashMap<String, &SourceComment> = HashMap::new();
        if let Some(read) = self.read() {
            for comment in &read.comments {
                if stored.contains_key(&comment.id) || !read.seeded_comments.contains(&comment.id) {
                    order.push(comment.id.clone());
                    sources.insert(comment.id.clone(), comment);
                }
            }
        }
        let mut rest: Vec<String> = stored
            .keys()
            .filter(|id| !sources.contains_key(*id))
            .cloned()
            .collect();
        rest.sort_by_key(|id| (id.parse::<i64>().ok(), id.clone()));
        order.extend(rest);
        let view = match self.view() {
            RevisionView::Original => EditTextView::Original,
            _ => EditTextView::Accepted,
        };
        let read = self.read();
        let mut planned = Vec::new();
        let mut ambiguous: Vec<(String, Anchor)> = Vec::new();
        let mut shared_ids: HashMap<String, HashSet<String>> = HashMap::new();
        for id in order {
            let entry = stored.get(&id);
            let source = sources.get(&id).copied();
            let current = |key: &str| -> Option<Option<&Any>> {
                let (values, _) = entry?;
                let written = match (source, read) {
                    (Some(_), Some(read)) => read.comment_writes.written(&id, key),
                    _ => true,
                };
                written.then(|| values.get(key))
            };
            let text = |key: &str, fallback: Option<&String>| match current(key) {
                Some(value) => any_text(value).filter(|value| !value.is_empty()),
                None => fallback.cloned(),
            };
            let mut anchors = Vec::new();
            let ranges = entry.map(|(_, ranges)| ranges.clone()).unwrap_or_default();
            for (story, start, end) in ranges {
                if !self.fits(Self::story_units(views, &story)) {
                    let unlocated = Anchor::Unlocated {
                        story: story.clone(),
                        reason: UnlocatedReason::StoryTooLarge,
                    };
                    self.note(
                        DiagnosticCode::ProvenanceUnavailable,
                        Severity::Warning,
                        Some(unlocated.clone()),
                        format!("A range of comment {id} lies in story {story}, which is too large to read within maxBytes, so it has no location."),
                    );
                    anchors.push(unlocated);
                    continue;
                }
                if let Some(projection) = views.story(&story, view)
                    && let Some(range) = projection.range_of_raw(start, end)
                {
                    let shared = shared_ids
                        .entry(story.clone())
                        .or_insert_with(|| duplicate_ids(&projection));
                    if shared.contains(&range.start.para_id) || shared.contains(&range.end.para_id)
                    {
                        let unlocated = Anchor::Unlocated {
                            story,
                            reason: UnlocatedReason::DuplicateParagraphId,
                        };
                        ambiguous.push((id.clone(), unlocated.clone()));
                        anchors.push(unlocated);
                        continue;
                    }
                    anchors.push(Anchor::Range(range));
                }
            }
            let body = match current("body") {
                Some(Some(Any::Array(blocks))) => CommentBody::Stored(Arc::clone(blocks)),
                Some(Some(Any::String(text))) => CommentBody::Text(text.to_string()),
                Some(_) if source.is_none() => CommentBody::Empty,
                _ => match source {
                    Some(source) => CommentBody::Source(&source.body),
                    None => CommentBody::Empty,
                },
            };
            let resolved = match current("done") {
                Some(value) => matches!(value, Some(Any::Bool(true))),
                None => source.is_some_and(|source| source.done),
            };
            self.comment_ids.insert(id.clone());
            planned.push(Planned {
                story: ExportStory {
                    story: format!("comment:{id}"),
                    kind: StoryKind::Comment,
                    part: source.map(|_| COMMENTS_PART.to_owned()),
                    note_id: None,
                    comment: Some(CommentMetadata {
                        id: id.clone(),
                        author: text("author", source.and_then(|source| source.author.as_ref())),
                        date: text("date", source.and_then(|source| source.date.as_ref())),
                        parent_id: text(
                            "parentId",
                            source.and_then(|source| source.parent_id.as_ref()),
                        ),
                        resolved,
                        anchors,
                    }),
                    uses: Vec::new(),
                    blocks: Vec::new(),
                },
                comment_body: Some(body),
            });
        }
        for (id, anchor) in ambiguous {
            self.note(
                DiagnosticCode::AmbiguousIdentity,
                Severity::Warning,
                Some(anchor),
                format!("A range of comment {id} lies in a paragraph whose id another paragraph shares, so it has no location of its own."),
            );
        }
        planned
    }

    /// Walks a root story, committing its blocks one at a time until the budget runs out.
    fn root_story<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        story: &str,
        prov: Prov,
        owner: Option<Anchor>,
        part: Option<String>,
    ) {
        let mut list = ListState::new(self.numbering());
        self.story_blocks(views, story, prov, owner, part, 0, &mut list, true);
        if !self.pending.is_empty() && !self.stopped {
            self.commit_notes();
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn story_ctx<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        story: &str,
        prov: Prov,
        owner: Option<Anchor>,
        part: Option<String>,
        limit: u32,
    ) -> Option<StoryCtx> {
        let (accepted, complete) = views.story_within(story, EditTextView::Accepted, limit)?;
        let original = match self.view() {
            RevisionView::Accepted => Rc::clone(&accepted),
            _ => {
                let limit = if complete { u32::MAX } else { limit };
                views.story_within(story, EditTextView::Original, limit)?.0
            }
        };
        let txn = views.txn();
        let text = story_ref(txn, story).ok()?;
        let chunks = views.doc().chunk_snapshot(story, &text, txn);
        let (cut, beyond) = if complete {
            (u32::MAX, None)
        } else {
            let (cut, ids) = later_ids(&accepted, &chunks, limit, txn);
            (cut, Some(ids))
        };
        let mut hidden = HashMap::new();
        let mut owners = HashSet::new();
        for (index, paragraph) in accepted.paragraphs.iter().enumerate() {
            let mut ordinal = 0;
            for (kind, map) in &paragraph.embeds {
                let Some(map) = map else { continue };
                if kind != "field"
                    || !numeric_field_instruction(
                        &map_string(map, txn, "instruction").unwrap_or_default(),
                    )
                {
                    continue;
                }
                if let Some(Out::Any(Any::Array(ids))) = map.get(txn, "fieldResultBlocks") {
                    for id in ids.iter() {
                        if let Any::String(id) = id {
                            hidden.insert(id.to_string(), (index, ordinal));
                            owners.insert(index);
                        }
                    }
                    ordinal += 1;
                }
            }
        }
        let mut seen = HashSet::new();
        let duplicates = accepted
            .paragraphs
            .iter()
            .filter(|paragraph| {
                !seen.insert(paragraph.para_id.as_str())
                    || beyond
                        .as_ref()
                        .is_some_and(|beyond| beyond.contains(&paragraph.para_id))
            })
            .map(|paragraph| paragraph.para_id.clone())
            .collect();
        let placements = self.placements(story, &prov, &accepted, &chunks, beyond.as_ref(), txn);
        let marks = accepted
            .paragraphs
            .iter()
            .map(|paragraph| {
                let chunk = chunks
                    .binary_search_by_key(&paragraph.pilcrow, |chunk| chunk.start)
                    .ok()
                    .map(|index| &chunks[index]);
                let key = |key: &str| present(paragraph.mark.properties.get(key)).is_some();
                (
                    chunk.is_some_and(|chunk| chunk.attr_active(INS)) || key("pPrIns"),
                    chunk.is_some_and(|chunk| chunk.attr_active(DEL)) || key("pPrDel"),
                )
            })
            .collect();
        Some(StoryCtx {
            story: story.to_owned(),
            part,
            prov,
            accepted,
            original,
            chunks,
            owner,
            block_owner: None,
            hidden,
            owners,
            field_anchors: HashMap::new(),
            placements,
            duplicates,
            paragraph_style: None,
            marks,
            beyond,
            cut,
        })
    }

    /// Where the source content of `story` the stream leaves out goes back in, reconciled
    /// through the blocks and paragraphs that still exist.
    fn placements<T: ReadTxn>(
        &mut self,
        story: &str,
        prov: &Prov,
        accepted: &StoryView,
        chunks: &[Chunk],
        beyond: Option<&HashSet<String>>,
        txn: &T,
    ) -> Placements {
        let mut placements = Placements::default();
        let read = self.read();
        let (provenance, pinned) = match prov {
            Prov::Main => match read {
                Some(read) => (&read.provenance, read.pinned),
                None => return placements,
            },
            Prov::Scratch(scratch) => (&scratch.provenance, true),
        };
        let raw_anchor = |index: usize| match prov {
            Prov::Main => read.and_then(|read| read.raw_block_anchor(story, index)),
            Prov::Scratch(scratch) => scratch
                .raw_anchors
                .get(story)
                .and_then(|anchors| anchors.get(index).cloned().flatten()),
        };
        if let (Some(elements), Some(order)) = (
            provenance.raw_blocks.get(story),
            provenance.block_order.get(story),
        ) {
            let mut alive = alive_blocks(accepted, chunks, txn);
            alive.extend(beyond.into_iter().flatten().cloned());
            let mut targets = vec![None; order.len()];
            let mut next: Option<&String> = None;
            for (position, entry) in order.iter().enumerate().rev() {
                targets[position] = next;
                if let Some(id) = entry.as_ref().filter(|id| alive.contains(id.as_str())) {
                    next = Some(id);
                }
            }
            let mut index = 0;
            for (position, entry) in order.iter().enumerate() {
                if entry.is_some() {
                    continue;
                }
                let element = elements.get(index).cloned().unwrap_or_default();
                let anchor = raw_anchor(index);
                index += 1;
                match targets[position] {
                    Some(id) => placements
                        .before
                        .entry(id.clone())
                        .or_default()
                        .push((element, anchor)),
                    None => self.note(
                        DiagnosticCode::ProvenanceUnavailable,
                        Severity::Warning,
                        anchor,
                        format!("A {element} block of story {story} lost every block after it, so its position is unknown."),
                    ),
                }
            }
        }
        if !pinned {
            return placements;
        }
        let embed_kind = |raw: u32| {
            let index = chunks
                .binary_search_by_key(&raw, |chunk| chunk.start)
                .ok()?;
            match &chunks[index].kind {
                ChunkKind::Embed(Some(map)) => map_string(map, txn, KIND_KEY),
                _ => None,
            }
        };
        let Some((records, relocated)) = provenance.by_story.get(story) else {
            return placements;
        };
        let mut by_id: HashMap<&str, &crate::target::ParagraphView> = HashMap::new();
        for paragraph in &accepted.paragraphs {
            by_id.entry(paragraph.para_id.as_str()).or_insert(paragraph);
        }
        let paragraph = |para_id: &str| by_id.get(para_id).copied();
        let mut resolved = HashMap::new();
        for record in records.iter().map(|index| &provenance.inline[*index]) {
            let Some(owner) = paragraph(&record.para_id) else {
                if matches!(record.content, InlineSource::Omitted { .. })
                    && !beyond.is_some_and(|beyond| beyond.contains(&record.para_id))
                {
                    self.note(
                        DiagnosticCode::ProvenanceUnavailable,
                        Severity::Warning,
                        None,
                        "Omitted source content lost its paragraph to an edit.",
                    );
                }
                continue;
            };
            let Some(raw) = record
                .pin
                .resolve(txn, &mut resolved)
                .filter(|raw| owner.node_start <= *raw && *raw <= owner.pilcrow)
            else {
                self.note(
                    DiagnosticCode::ProvenanceUnavailable,
                    Severity::Warning,
                    None,
                    format!(
                        "Source content of paragraph {} can no longer be placed.",
                        record.para_id
                    ),
                );
                continue;
            };
            let entry = match &record.content {
                InlineSource::Omitted { element } => Record::Omitted {
                    element: element.clone(),
                    in_control: record.in_control,
                },
                InlineSource::Break { kind, revision } => {
                    let kept = match record.witness {
                        Witness::Invisible => true,
                        Witness::Leading => {
                            any_flag(owner.mark.properties.get("pageBreakBeforeRun"))
                        }
                        Witness::Embed(index) => provenance
                            .relocated
                            .get(index)
                            .and_then(|relocated| relocated.pin.resolve(txn, &mut resolved))
                            .filter(|raw| {
                                embed_kind(*raw).as_deref()
                                    == Some(match kind {
                                        BreakType::Column => "columnBreak",
                                        _ => "pageBreak",
                                    })
                            })
                            .map(|raw| placements.relocated.insert(raw))
                            .is_some(),
                    };
                    if !kept {
                        continue;
                    }
                    Record::Break {
                        kind: *kind,
                        revision: revision.clone(),
                        in_control: record.in_control,
                        control_offset: record.control_offset.clone(),
                    }
                }
            };
            placements
                .inline
                .entry(record.para_id.clone())
                .or_default()
                .push((raw, entry));
        }
        for index in relocated {
            let relocated = &provenance.relocated[*index];
            if provenance.witnessed.contains(index) {
                continue;
            }
            if let Some(raw) = relocated.pin.resolve(txn, &mut resolved)
                && matches!(
                    embed_kind(raw).as_deref(),
                    Some("pageBreak" | "columnBreak")
                )
                && paragraph(&relocated.para_id).is_some()
            {
                placements.relocated.insert(raw);
            }
        }
        placements
    }

    /// The blocks of `story` in order. A root story commits each root block as it completes and
    /// returns nothing; a nested one returns its blocks.
    #[allow(clippy::too_many_arguments)]
    fn story_blocks<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        story: &str,
        prov: Prov,
        owner: Option<Anchor>,
        part: Option<String>,
        depth: usize,
        list: &mut ListState,
        root: bool,
    ) -> Vec<Built> {
        let limit = if root {
            u32::try_from(self.capacity()).unwrap_or(u32::MAX)
        } else if self.projectable(views, story) {
            u32::MAX
        } else {
            return Vec::new();
        };
        let Some(mut ctx) = self.story_ctx(views, story, prov, owner.clone(), part, limit) else {
            let anchor = owner.unwrap_or_else(|| Anchor::Unlocated {
                story: story.to_owned(),
                reason: UnlocatedReason::MissingStory,
            });
            self.note(
                DiagnosticCode::UnresolvedReference,
                Severity::Warning,
                Some(anchor.clone()),
                format!("Story {story} is missing."),
            );
            let mut output = Vec::new();
            if self.grow(1, 0) {
                let built = Built {
                    block: Block {
                        id: String::new(),
                        anchor,
                        content: BlockKind::Unsupported {
                            element: "story".to_owned(),
                        },
                    },
                    diagnostics: std::mem::take(&mut self.pending),
                    paragraph: None,
                };
                self.deliver(&mut output, root, built);
            }
            return output;
        };
        let chunks = Arc::clone(&ctx.chunks);
        let accepted = Rc::clone(&ctx.accepted);
        let mut output: Vec<Built> = Vec::new();
        let mut held: Option<Held> = None;
        let mut cursor = 0usize;
        let mut table_index = 0u32;
        for (index, paragraph) in accepted.paragraphs.iter().enumerate() {
            if self.stopped {
                break;
            }
            while cursor < chunks.len() && chunks[cursor].start < paragraph.start {
                cursor += 1;
            }
            while cursor < chunks.len() && chunks[cursor].start < paragraph.node_start {
                let chunk = &chunks[cursor];
                cursor += 1;
                self.stream_block(
                    views,
                    &mut ctx,
                    chunk,
                    depth,
                    list,
                    &mut table_index,
                    (&mut output, &mut held, root),
                );
                if self.stopped {
                    break;
                }
            }
            if self.stopped {
                break;
            }
            self.raw_blocks(&mut ctx, &paragraph.para_id, (&mut output, &mut held, root));
            let end = chunks[cursor..]
                .iter()
                .position(|chunk| chunk.start >= paragraph.pilcrow)
                .map_or(chunks.len(), |offset| cursor + offset);
            let hidden_by = ctx.hidden.get(&paragraph.para_id).copied();
            self.settle(&mut output, &mut held, root, hidden_by);
            if !self.visit(1) {
                break;
            }
            ctx.block_owner = hidden_by.map(|owner| self.field_anchor(&ctx, owner));
            let block = self.paragraph(views, &mut ctx, index, &chunks[cursor..end], depth, list);
            ctx.block_owner = None;
            cursor = end;
            if self.stopped {
                break;
            }
            let built = Built {
                block,
                diagnostics: std::mem::take(&mut self.pending),
                paragraph: Some(index),
            };
            self.place(&ctx, &mut output, &mut held, root, built, hidden_by);
            if let Some(section) = self.section_break(&ctx, &paragraph.para_id) {
                let built = Built {
                    block: section,
                    diagnostics: Vec::new(),
                    paragraph: None,
                };
                match held.as_mut() {
                    Some(held) if held.built.paragraph == Some(index) => held.trailing.push(built),
                    _ => self.place(&ctx, &mut output, &mut held, root, built, None),
                }
            }
        }
        let tail = ctx
            .accepted
            .paragraphs
            .last()
            .map_or(0, |paragraph| paragraph.pilcrow + 1);
        let cut = ctx.cut;
        for chunk in chunks
            .iter()
            .filter(|chunk| chunk.start >= tail && chunk.start < cut)
        {
            if self.stopped {
                break;
            }
            self.stream_block(
                views,
                &mut ctx,
                chunk,
                depth,
                list,
                &mut table_index,
                (&mut output, &mut held, root),
            );
        }
        if let Some(beyond) = ctx.beyond.as_ref()
            && !self.stopped
        {
            if let Some(owner) = held.take()
                && !ctx.hidden.iter().any(|(id, (paragraph, _))| {
                    owner.built.paragraph == Some(*paragraph) && beyond.contains(id)
                })
            {
                self.release(&mut output, root, owner);
            }
            self.oversize();
        }
        if let Some(held) = held.take() {
            self.release(&mut output, root, held);
        }
        for (_, raws) in std::mem::take(&mut ctx.placements.before) {
            for (element, provenance) in raws {
                self.note(
                    DiagnosticCode::ProvenanceUnavailable,
                    Severity::Warning,
                    provenance,
                    format!("A {element} block of story {story} could not be placed."),
                );
            }
        }
        if !root && let Some(last) = output.last_mut() {
            last.diagnostics.append(&mut self.pending);
        }
        output
    }

    /// The anchor of the field whose result hides a block, or its paragraph's until the field
    /// has been built.
    fn field_anchor(&self, ctx: &StoryCtx, owner: (usize, usize)) -> Anchor {
        ctx.field_anchors
            .get(&owner)
            .cloned()
            .unwrap_or_else(|| ctx.paragraph_anchor(&ctx.accepted.paragraphs[owner.0].para_id))
    }

    /// Hands a completed block of a story on: to its hidden-result owner, to the held owner, or
    /// out as the story's next block.
    fn place(
        &mut self,
        ctx: &StoryCtx,
        output: &mut Vec<Built>,
        held: &mut Option<Held>,
        root: bool,
        built: Built,
        hidden_by: Option<(usize, usize)>,
    ) {
        let mut built = built;
        if let Some((paragraph, ordinal)) = hidden_by
            && let Some(owner) = held
                .as_mut()
                .filter(|owner| owner.built.paragraph == Some(paragraph))
        {
            match attach_result(&mut owner.built.block, ordinal, built.block) {
                None => {
                    owner.built.diagnostics.extend(built.diagnostics);
                    return;
                }
                Some(block) => built.block = block,
            }
        }
        if let Some(previous) = held.take() {
            self.release(output, root, previous);
        }
        if built
            .paragraph
            .is_some_and(|paragraph| ctx.owners.contains(&paragraph))
            && hidden_by.is_none()
        {
            *held = Some(Held {
                built,
                trailing: Vec::new(),
            });
            return;
        }
        self.deliver(output, root, built);
    }

    /// Hands a held paragraph on before building a block that is not one of its results, so a
    /// completed owner never waits on a block that may not fit.
    fn settle(
        &mut self,
        output: &mut Vec<Built>,
        held: &mut Option<Held>,
        root: bool,
        upcoming: Option<(usize, usize)>,
    ) {
        let belongs = held.as_ref().is_some_and(|owner| {
            upcoming.is_some_and(|(paragraph, _)| owner.built.paragraph == Some(paragraph))
        });
        if !belongs && let Some(previous) = held.take() {
            self.release(output, root, previous);
        }
    }

    fn release(&mut self, output: &mut Vec<Built>, root: bool, held: Held) {
        self.deliver(output, root, held.built);
        for built in held.trailing {
            self.deliver(output, root, built);
        }
    }

    fn deliver(&mut self, output: &mut Vec<Built>, root: bool, built: Built) {
        if root {
            self.commit(built);
        } else {
            output.push(built);
        }
    }

    /// The raw XML blocks placed in front of the block `id`, as placeholders.
    fn raw_blocks(
        &mut self,
        ctx: &mut StoryCtx,
        id: &str,
        (output, held, root): (&mut Vec<Built>, &mut Option<Held>, bool),
    ) {
        let Some(raws) = ctx.placements.before.remove(id) else {
            return;
        };
        self.settle(output, held, root, None);
        for (element, provenance) in raws {
            let anchor = provenance.unwrap_or_else(|| {
                ctx.anchor(ctx.unlocated(UnlocatedReason::ProvenanceUnavailable))
            });
            self.note(
                DiagnosticCode::UnsupportedContent,
                Severity::Warning,
                Some(anchor.clone()),
                format!("A {element} block is not represented in the export."),
            );
            if !self.grow(1, 0) {
                return;
            }
            let built = Built {
                block: Block {
                    id: String::new(),
                    anchor,
                    content: BlockKind::Unsupported { element },
                },
                diagnostics: std::mem::take(&mut self.pending),
                paragraph: None,
            };
            self.place(ctx, output, held, root, built, None);
        }
    }

    /// One block embed of the stream, placed after the raw blocks that precede it.
    #[allow(clippy::too_many_arguments)]
    fn stream_block<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &mut StoryCtx,
        chunk: &Chunk,
        depth: usize,
        list: &mut ListState,
        table_index: &mut u32,
        (output, held, root): (&mut Vec<Built>, &mut Option<Held>, bool),
    ) {
        if ctx.placements.relocated.contains(&chunk.start) {
            return;
        }
        if let Some(id) = self.block_identity(ctx, chunk, views, *table_index) {
            self.raw_blocks(ctx, &id, (output, held, root));
        }
        let hidden_by = self.hidden_owner(ctx, chunk, views);
        self.settle(output, held, root, hidden_by);
        if !self.visit(1) {
            return;
        }
        ctx.block_owner = hidden_by.map(|owner| self.field_anchor(ctx, owner));
        let block = self.block_embed(views, ctx, chunk, depth, list, table_index);
        ctx.block_owner = None;
        if let Some(block) = block {
            let built = Built {
                block,
                diagnostics: std::mem::take(&mut self.pending),
                paragraph: None,
            };
            self.place(ctx, output, held, root, built, hidden_by);
        }
    }

    /// The id seeding gives the block embed `chunk`: `{story}:t{n}` for a table, the child story
    /// for a block control.
    fn block_identity<T: ReadTxn>(
        &self,
        ctx: &StoryCtx,
        chunk: &Chunk,
        views: &Views<'_, T>,
        table_index: u32,
    ) -> Option<String> {
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            return None;
        };
        let txn = views.txn();
        match map_string(map, txn, KIND_KEY)?.as_str() {
            "table" => Some(
                table_identity(&payload_of(map, txn))
                    .unwrap_or_else(|| format!("{}:t{table_index}", ctx.story)),
            ),
            "blockSdt" => map_string(map, txn, "story"),
            _ => None,
        }
    }

    /// The paragraph and ordinal of the numeric field hiding the block embed `chunk` as its
    /// result.
    fn hidden_owner<T: ReadTxn>(
        &self,
        ctx: &StoryCtx,
        chunk: &Chunk,
        views: &Views<'_, T>,
    ) -> Option<(usize, usize)> {
        if ctx.hidden.is_empty() {
            return None;
        }
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            return None;
        };
        let txn = views.txn();
        let id = map_string(map, txn, "blockId").or_else(|| map_string(map, txn, "story"))?;
        ctx.hidden.get(&id).copied()
    }

    fn section_break(&mut self, ctx: &StoryCtx, para_id: &str) -> Option<Block> {
        if ctx.story != "body" || ctx.owner.is_some() || !matches!(ctx.prov, Prov::Main) {
            return None;
        }
        let index = *self.section_ends.get(para_id)?;
        let start = self
            .sections
            .get(index + 1)
            .and_then(|next| next.start.clone());
        if !self.grow(1, 0) {
            return None;
        }
        Some(Block {
            id: String::new(),
            anchor: ctx.paragraph_anchor(para_id),
            content: BlockKind::SectionBreak {
                section_index: index as u32 + 1,
                break_type: section_break_type(start.as_deref()),
            },
        })
    }

    /// Whether the current view hides stream content with these revision stamps.
    fn hides(&mut self, ins: bool, del: bool) -> bool {
        let hidden = match self.view() {
            RevisionView::Accepted => del,
            RevisionView::Original => ins,
            RevisionView::Markup => false,
        };
        self.excluded_revision |= hidden;
        hidden
    }

    /// An anchored placeholder for content whose historical state this view cannot reconstruct.
    fn unreconstructed(&mut self, anchor: Anchor, element: &str, message: &str) -> Block {
        self.note(
            DiagnosticCode::UnsupportedRevision,
            Severity::Warning,
            Some(anchor.clone()),
            message,
        );
        Block {
            id: String::new(),
            anchor,
            content: BlockKind::Unsupported {
                element: element.to_owned(),
            },
        }
    }

    fn block_embed<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &StoryCtx,
        chunk: &Chunk,
        depth: usize,
        list: &mut ListState,
        table_index: &mut u32,
    ) -> Option<Block> {
        if !self.grow(1, 0) {
            return None;
        }
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            return Some(self.unsupported_block(ctx, "embed", chunk));
        };
        let txn = views.txn();
        let kind = map_string(map, txn, KIND_KEY).unwrap_or_default();
        let payload = payload_of(map, txn);
        let (ins, del, _) = revisions_of(|key| chunk.attrs.get(key).cloned(), None);
        match kind.as_str() {
            "table" => {
                let index = *table_index;
                *table_index += 1;
                if self.hides(ins, del) {
                    return None;
                }
                let anchor = ctx.anchor(Anchor::Table {
                    story: ctx.story.clone(),
                    table_index: index,
                });
                if ins || del {
                    return Some(self.unreconstructed(
                        anchor,
                        "w:tbl",
                        "The table's tracked insertion or deletion cannot be attributed in the markup view.",
                    ));
                }
                Some(self.table(views, ctx, &payload, anchor, depth, list))
            }
            "blockSdt" => {
                if self.hides(ins, del) {
                    return None;
                }
                let child = any_text(payload.get("story"))?;
                let control_id = block_control_id(&child);
                let anchor = ctx.anchor(Anchor::Control {
                    story: ctx.story.clone(),
                    control_id: control_id.clone(),
                });
                if ins || del {
                    return Some(self.unreconstructed(
                        anchor,
                        "w:sdt",
                        "The content control's tracked insertion or deletion cannot be attributed in the markup view.",
                    ));
                }
                let control = control_metadata(control_id, |key| payload.get(key));
                if present(payload.get("value")).is_some() {
                    self.note(
                        DiagnosticCode::LegacyControlValue,
                        Severity::Warning,
                        Some(anchor.clone()),
                        "An authored value replaces this control's content when the document is saved; the export shows its current content.",
                    );
                }
                let blocks = self.nested(views, ctx, &child, depth, list, &anchor);
                Some(Block {
                    id: String::new(),
                    anchor,
                    content: BlockKind::ContentControl {
                        control,
                        story: ctx.session_story(child),
                        blocks,
                    },
                })
            }
            "pageBreak" | "columnBreak" => {
                if self.hides(ins, del) {
                    return None;
                }
                let paragraphs = &ctx.accepted.paragraphs;
                let para_id = paragraphs
                    .get(
                        paragraphs.partition_point(|paragraph| paragraph.node_start <= chunk.start),
                    )
                    .or_else(|| paragraphs.last())
                    .map(|paragraph| paragraph.para_id.clone())
                    .unwrap_or_default();
                Some(Block {
                    id: String::new(),
                    anchor: ctx.paragraph_anchor(&para_id),
                    content: BlockKind::Break {
                        break_type: if kind == "pageBreak" {
                            BreakType::Page
                        } else {
                            BreakType::Column
                        },
                    },
                })
            }
            _ => Some(self.unsupported_block(ctx, &kind, chunk)),
        }
    }

    fn unsupported_block(&mut self, ctx: &StoryCtx, element: &str, chunk: &Chunk) -> Block {
        let paragraphs = &ctx.accepted.paragraphs;
        let para_id = paragraphs
            .get(paragraphs.partition_point(|paragraph| paragraph.pilcrow < chunk.start))
            .or_else(|| paragraphs.last())
            .map(|paragraph| paragraph.para_id.clone())
            .unwrap_or_default();
        let anchor = ctx.paragraph_anchor(&para_id);
        self.note(
            DiagnosticCode::UnsupportedContent,
            Severity::Warning,
            Some(anchor.clone()),
            format!("A {element} block is not represented in the export."),
        );
        Block {
            id: String::new(),
            anchor,
            content: BlockKind::Unsupported {
                element: element.to_owned(),
            },
        }
    }

    /// The blocks of a child story owned by a table cell or block control.
    fn nested<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &StoryCtx,
        story: &str,
        depth: usize,
        list: &mut ListState,
        container: &Anchor,
    ) -> Vec<Block> {
        if depth + 1 > MAX_DEPTH {
            self.note(
                DiagnosticCode::UnsupportedContent,
                Severity::Warning,
                Some(container.clone()),
                format!("Content nested more than {MAX_DEPTH} levels deep is not exported."),
            );
            return Vec::new();
        }
        let built = self.story_blocks(
            views,
            story,
            ctx.prov.clone(),
            ctx.owner().cloned(),
            ctx.part.clone(),
            depth + 1,
            list,
            false,
        );
        built
            .into_iter()
            .map(|built| {
                self.pending.extend(built.diagnostics);
                built.block
            })
            .collect()
    }

    fn table<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &StoryCtx,
        payload: &Payload,
        anchor: Anchor,
        depth: usize,
        list: &mut ListState,
    ) -> Block {
        let rows: Vec<HashMap<String, Any>> = match payload.get("rows") {
            Some(Any::Array(rows)) => rows
                .iter()
                .filter_map(|row| any_map(Some(row)).cloned())
                .collect(),
            _ => Vec::new(),
        };
        let row_revision = |row: &HashMap<String, Any>, key: &str| {
            any_map(row.get("trPr"))
                .is_some_and(|properties| present(properties.get(key)).is_some())
        };
        let view = self.view();
        let tracked_rows = rows
            .iter()
            .any(|row| row_revision(row, "trIns") || row_revision(row, "trDel"));
        if tracked_rows && view == RevisionView::Markup {
            return self.unreconstructed(
                anchor,
                "w:tbl",
                "Tracked row insertions and deletions cannot be attributed in the markup view.",
            );
        }
        let hidden_rows: Vec<bool> = rows
            .iter()
            .map(|row| match view {
                RevisionView::Accepted => row_revision(row, "trDel"),
                RevisionView::Original => row_revision(row, "trIns"),
                RevisionView::Markup => false,
            })
            .collect();
        let cells = |row: &HashMap<String, Any>| -> Vec<HashMap<String, Any>> {
            match row.get("cells") {
                Some(Any::Array(cells)) => cells
                    .iter()
                    .filter_map(|cell| any_map(Some(cell)).cloned())
                    .collect(),
                _ => Vec::new(),
            }
        };
        if let Some(reason) = unreconstructed_table(&rows, &cells, view) {
            return self.unreconstructed(anchor, "w:tbl", reason);
        }
        let table_id = table_identity(payload);
        let layout = match (&ctx.prov, table_id.as_deref()) {
            (Prov::Main, Some(id)) => self.read().and_then(|read| read.provenance.tables.get(id)),
            (Prov::Scratch(scratch), Some(id)) => scratch.provenance.tables.get(id),
            _ => None,
        };
        let stories: Vec<Vec<Option<String>>> = rows
            .iter()
            .map(|row| {
                cells(row)
                    .iter()
                    .map(|cell| any_text(cell.get("story")))
                    .collect()
            })
            .collect();
        let layout = layout.filter(|layout| {
            layout.rows.len() == rows.len()
                && layout.rows.iter().zip(&stories).all(|(row, stories)| {
                    let kept: Vec<Option<String>> = row
                        .cells
                        .iter()
                        .filter(|cell| cell.story.is_some())
                        .map(|cell| cell.story.clone())
                        .collect();
                    &kept == stories
                })
        });
        if hidden_rows.iter().any(|hidden| *hidden) {
            self.excluded_revision = true;
            let merged = match layout {
                Some(layout) => layout
                    .rows
                    .iter()
                    .any(|row| row.cells.iter().any(|cell| cell.merge != SourceMerge::None)),
                None => rows.iter().any(|row| {
                    cells(row).iter().any(|cell| {
                        any_map(cell.get("tcPr"))
                            .and_then(|properties| any_number(properties.get("rowspan")))
                            .is_some_and(|span| span > 1.0)
                    })
                }),
            };
            if merged {
                return self.unreconstructed(
                    anchor,
                    "w:tbl",
                    "Vertically merged cells cross tracked row changes, which this view cannot reconstruct.",
                );
            }
        }
        let header = |row: &HashMap<String, Any>| {
            any_map(row.get("trPr")).is_some_and(|properties| {
                matches!(properties.get("isHeader"), Some(Any::Bool(true)))
            })
        };
        let grid = match payload.get("grid") {
            Some(Any::Array(grid)) => grid.len() as u32,
            _ => 0,
        };
        let table = match layout {
            Some(layout) => {
                let layout: &TableLayout = layout;
                self.table_from_layout(
                    views,
                    ctx,
                    layout,
                    &rows,
                    &hidden_rows,
                    header,
                    &anchor,
                    depth,
                    list,
                )
            }
            None => {
                if table_id.is_some() && matches!(ctx.prov, Prov::Main) && self.read().is_some() {
                    self.note(
                        DiagnosticCode::ProvenanceUnavailable,
                        Severity::Warning,
                        Some(anchor.clone()),
                        "The table's rows or cells changed since it was read from the package, so its grid comes from the editing model, which does not keep every vertical merge or skipped grid column.",
                    );
                }
                self.table_from_stream(
                    views,
                    ctx,
                    &rows,
                    &hidden_rows,
                    header,
                    &anchor,
                    depth,
                    list,
                )
            }
        };
        Block {
            id: String::new(),
            anchor,
            content: BlockKind::Table {
                table: TableData {
                    grid_columns: table.grid_columns.max(grid),
                    rows: table.rows,
                },
            },
        }
    }

    /// A cell's anchor: its first paragraph, or the table's when it has none.
    fn cell_anchor<T: ReadTxn>(
        &self,
        views: &mut Views<'_, T>,
        ctx: &StoryCtx,
        story: Option<&str>,
        table: &Anchor,
    ) -> Anchor {
        story
            .filter(|story| self.fits(Self::story_units(views, story)))
            .and_then(|story| views.story(story, EditTextView::Accepted))
            .and_then(|view| {
                let first = view.paragraphs.first()?;
                unique_id(&view, &first.para_id).then(|| Anchor::Paragraph {
                    story: view.story.clone(),
                    para_id: first.para_id.clone(),
                })
            })
            .map_or_else(|| table.clone(), |anchor| ctx.anchor(anchor))
    }

    /// The table on its source grid: merges and skipped grid columns as the package defines them,
    /// content from the cell stories.
    #[allow(clippy::too_many_arguments)]
    fn table_from_layout<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &StoryCtx,
        layout: &TableLayout,
        rows: &[HashMap<String, Any>],
        hidden_rows: &[bool],
        header: impl Fn(&HashMap<String, Any>) -> bool,
        anchor: &Anchor,
        depth: usize,
        list: &mut ListState,
    ) -> TableData {
        let mut output: Vec<TableRow> = Vec::new();
        // Open merges by grid column: the origin's row and cell index in the output.
        let mut open: BTreeMap<u32, (usize, usize, u32)> = BTreeMap::new();
        for ((source, row), hidden) in layout.rows.iter().zip(rows).zip(hidden_rows) {
            if *hidden || self.stopped {
                continue;
            }
            let row_index = output.len();
            let mut cells = Vec::new();
            let mut continued = BTreeSet::new();
            for cell in &source.cells {
                let origin = open
                    .get(&cell.column)
                    .copied()
                    .filter(|(_, _, span)| *span == cell.span);
                if cell.merge == SourceMerge::Continue
                    && let Some((origin_row, origin_cell, _)) = origin
                {
                    continued.insert(cell.column);
                    let omitted = cell.content
                        || cell.story.as_deref().is_some_and(|story| {
                            !self.fits(Self::story_units(views, story))
                                || views
                                    .story(story, EditTextView::Accepted)
                                    .is_some_and(|view| {
                                        view.paragraphs
                                            .iter()
                                            .any(|paragraph| !paragraph.text.is_empty())
                                    })
                        });
                    if omitted {
                        self.note(
                            DiagnosticCode::MergeContinuationContentOmitted,
                            Severity::Warning,
                            Some(anchor.clone()),
                            format!(
                                "The merged continuation cell at row {row_index}, column {} holds content the export omits.",
                                cell.column
                            ),
                        );
                    }
                    output[origin_row].cells[origin_cell].row_span += 1;
                    let origin_column = output[origin_row].cells[origin_cell].column;
                    cells.push(TableCell {
                        anchor: anchor.clone(),
                        story: None,
                        column: cell.column,
                        grid_span: cell.span,
                        row_span: 0,
                        vertical_merge: VerticalMerge::Continue,
                        merge_origin: Some(CellPosition {
                            row: origin_row as u32,
                            column: origin_column,
                        }),
                        blocks: Vec::new(),
                    });
                    continue;
                }
                if cell.merge == SourceMerge::Continue {
                    self.note(
                        DiagnosticCode::UnsupportedContent,
                        Severity::Warning,
                        Some(anchor.clone()),
                        format!("The cell at row {row_index}, column {} continues a vertical merge no cell above starts, so it is exported as a cell of its own.", cell.column),
                    );
                }
                let cell_anchor = self.cell_anchor(views, ctx, cell.story.as_deref(), anchor);
                let blocks = match &cell.story {
                    Some(story) => self.nested(views, ctx, story, depth, list, anchor),
                    None => Vec::new(),
                };
                let merge = if cell.merge == SourceMerge::Restart {
                    VerticalMerge::Restart
                } else {
                    VerticalMerge::None
                };
                if merge == VerticalMerge::Restart {
                    open.insert(cell.column, (row_index, cells.len(), cell.span));
                }
                cells.push(TableCell {
                    anchor: cell_anchor,
                    story: cell
                        .story
                        .clone()
                        .and_then(|story| ctx.session_story(story)),
                    column: cell.column,
                    grid_span: cell.span,
                    row_span: 1,
                    vertical_merge: merge,
                    merge_origin: None,
                    blocks,
                });
            }
            open.retain(|column, (origin_row, _, _)| {
                *origin_row == row_index || continued.contains(column)
            });
            output.push(TableRow {
                header: header(row),
                grid_before: source.grid_before,
                grid_after: source.grid_after,
                cells,
            });
        }
        TableData {
            grid_columns: layout.grid_columns,
            rows: output,
        }
    }

    /// The table as the editing model holds it, for tables without a matching source grid.
    #[allow(clippy::too_many_arguments)]
    fn table_from_stream<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &StoryCtx,
        rows: &[HashMap<String, Any>],
        hidden_rows: &[bool],
        header: impl Fn(&HashMap<String, Any>) -> bool,
        anchor: &Anchor,
        depth: usize,
        list: &mut ListState,
    ) -> TableData {
        let span_of = |cell: &HashMap<String, Any>, key: &str| {
            any_map(cell.get("tcPr"))
                .and_then(|properties| any_number(properties.get(key)))
                .filter(|span| span.is_finite())
                .unwrap_or(1.0)
                .clamp(1.0, f64::from(u16::MAX)) as u32
        };
        let mut grid_columns = 0;
        // Per grid column: origin row, origin column, span and rows still covered.
        let mut covered: BTreeMap<u32, (u32, u32, u32, u32)> = BTreeMap::new();
        let mut output_rows = Vec::new();
        for (row, hidden) in rows.iter().zip(hidden_rows) {
            if *hidden || self.stopped {
                continue;
            }
            let row_index = output_rows.len() as u32;
            let original = any_map(row.get("trPr"))
                .and_then(|properties| any_map(properties.get("_originalFormatting")));
            let grid = |key: &str| {
                original
                    .and_then(|original| any_number(original.get(key)))
                    .filter(|value| value.is_finite())
                    .unwrap_or(0.0)
                    .clamp(0.0, f64::from(u16::MAX)) as u32
            };
            let (grid_before, grid_after) = (grid("gridBefore"), grid("gridAfter"));
            let mut cells = Vec::new();
            let mut column = grid_before;
            let row_cells: Vec<HashMap<String, Any>> = match row.get("cells") {
                Some(Any::Array(cells)) => cells
                    .iter()
                    .filter_map(|cell| any_map(Some(cell)).cloned())
                    .collect(),
                _ => Vec::new(),
            };
            for cell in row_cells {
                continuations(&mut covered, &mut column, anchor, &mut cells);
                let span = span_of(&cell, "colspan");
                let rows_spanned = span_of(&cell, "rowspan");
                let story = any_text(cell.get("story"));
                let cell_anchor = self.cell_anchor(views, ctx, story.as_deref(), anchor);
                let blocks = match &story {
                    Some(story) => self.nested(views, ctx, story, depth, list, anchor),
                    None => Vec::new(),
                };
                if rows_spanned > 1 {
                    for slot in column..column.saturating_add(span) {
                        covered.insert(slot, (row_index, column, span, rows_spanned - 1));
                    }
                }
                cells.push(TableCell {
                    anchor: cell_anchor,
                    story: story.and_then(|story| ctx.session_story(story)),
                    column,
                    grid_span: span,
                    row_span: rows_spanned,
                    vertical_merge: if rows_spanned > 1 {
                        VerticalMerge::Restart
                    } else {
                        VerticalMerge::None
                    },
                    merge_origin: None,
                    blocks,
                });
                column = column.saturating_add(span);
            }
            continuations(&mut covered, &mut column, anchor, &mut cells);
            grid_columns = u32::max(grid_columns, column.saturating_add(grid_after));
            output_rows.push(TableRow {
                header: header(row),
                grid_before,
                grid_after,
                cells,
            });
        }
        TableData {
            grid_columns,
            rows: output_rows,
        }
    }

    /// One paragraph as a paragraph, heading or list item, or an anchored placeholder where this
    /// view cannot reconstruct it.
    fn paragraph<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &mut StoryCtx,
        index: usize,
        chunks: &[Chunk],
        depth: usize,
        list: &mut ListState,
    ) -> Block {
        let accepted = Rc::clone(&ctx.accepted);
        let paragraph = &accepted.paragraphs[index];
        if !self.grow(1, 0) {
            return Block {
                id: String::new(),
                anchor: ctx.paragraph_anchor(&paragraph.para_id),
                content: BlockKind::Unsupported {
                    element: "w:p".to_owned(),
                },
            };
        }
        if ctx.duplicates.contains(&paragraph.para_id) && ctx.owner().is_none() {
            let unlocated = ctx.unlocated(UnlocatedReason::DuplicateParagraphId);
            self.note(
                DiagnosticCode::AmbiguousIdentity,
                Severity::Warning,
                Some(unlocated.clone()),
                format!(
                    "Paragraph id {} occurs more than once in story {}, so this paragraph and its text have no location of their own.",
                    paragraph.para_id, ctx.story
                ),
            );
            ctx.block_owner = Some(unlocated);
        }
        let anchor = ctx.paragraph_anchor(&paragraph.para_id);
        let view = self.view();
        let moved = |mark: fn(&(bool, bool)) -> bool| {
            mark(&ctx.marks[index])
                || index
                    .checked_sub(1)
                    .is_some_and(|previous| mark(&ctx.marks[previous]))
        };
        let revised = match view {
            RevisionView::Accepted => moved(|(_, deleted)| *deleted),
            RevisionView::Original => moved(|(inserted, _)| *inserted),
            RevisionView::Markup => accepted.structurally_revised(index),
        };
        if revised {
            return self.unreconstructed(
                anchor,
                "w:p",
                "A tracked paragraph-mark change moves this paragraph's boundaries, which this view cannot reconstruct.",
            );
        }
        let properties = &paragraph.mark.properties;
        let change = match properties.get("pPrChange") {
            Some(Any::Array(changes)) => changes.first().and_then(|change| any_map(Some(change))),
            _ => None,
        };
        let previous = change.and_then(|change| any_map(change.get("previousFormatting")));
        let empty = HashMap::new();
        if view != RevisionView::Accepted
            && let Some(change) = change
        {
            let previous = previous.unwrap_or(&empty);
            let recorded = any_map(change.get("previousFormatting")).is_some();
            let direct = any_map(properties.get("_originalFormatting"));
            let differs = |key: &str| {
                present(previous.get(key)) != present(direct.and_then(|direct| direct.get(key)))
            };
            let original_style = if recorded {
                ctx.original.paragraphs[index].style_id.as_deref()
            } else {
                None
            };
            let style_differs = paragraph.style_id.as_deref() != original_style;
            if !recorded && style_differs {
                return self.unreconstructed(
                    anchor,
                    "w:p",
                    "A tracked paragraph formatting change records no earlier properties, so the paragraph's earlier style cannot be reconstructed.",
                );
            }
            let marks_differ = style_differs
                && self.options.include_formatting
                && self.source.is_none_or(|source| {
                    !source.same_run_marks(paragraph.style_id.as_deref(), original_style)
                });
            let reason = if differs("numPr") {
                Some(
                    "A tracked paragraph formatting change alters this paragraph's numbering, which this view cannot reconstruct.",
                )
            } else if view == RevisionView::Markup && (style_differs || differs("outlineLevel")) {
                Some(
                    "A tracked paragraph formatting change alters this paragraph's style or outline level, which the markup view cannot attribute.",
                )
            } else if marks_differ {
                Some(
                    "A tracked paragraph style change alters the style's run formatting, which this view cannot reconstruct in the paragraph's marks.",
                )
            } else {
                None
            };
            if let Some(reason) = reason {
                return self.unreconstructed(anchor, "w:p", reason);
            }
        }
        if view != RevisionView::Accepted
            && self.options.include_formatting
            && (paragraph.run_revisions
                || self
                    .source
                    .is_some_and(|source| source.run_revision(&ctx.story, &paragraph.para_id)))
            && run_marks_changed(properties)
        {
            return self.unreconstructed(
                anchor,
                "w:p",
                "Tracked run formatting changes alter this paragraph's marks, which this view cannot reconstruct.",
            );
        }
        let style_id = match view {
            RevisionView::Original => ctx.original.paragraphs[index].style_id.clone(),
            _ => paragraph.style_id.clone(),
        };
        if let (Some(source), Some(style)) = (self.source, style_id.as_deref())
            && !source.has_style(style)
        {
            self.note(
                DiagnosticCode::UnresolvedStyle,
                Severity::Warning,
                Some(anchor.clone()),
                format!("Style {style} is not defined in the document; the default paragraph style applies."),
            );
        }
        let inputs = match (view, previous) {
            (RevisionView::Original, Some(previous)) => OutlineInputs {
                effective: None,
                direct: any_number(previous.get("outlineLevel")),
                style_id: style_id.as_deref(),
            },
            _ => OutlineInputs {
                style_id: style_id.as_deref(),
                ..current_inputs(properties)
            },
        };
        let heading = resolve_heading(&inputs, self.source);
        let list_info = self.list_info(properties, list, &anchor);
        ctx.paragraph_style = style_id.clone();
        let inlines = self.inlines(views, ctx, index, chunks, depth);
        let paragraph = ParagraphData { style_id, inlines };
        let content = match (list_info, heading) {
            (Some(list), heading) => BlockKind::ListItem {
                paragraph,
                list,
                heading,
            },
            (None, Some(heading)) => BlockKind::Heading { paragraph, heading },
            (None, None) => BlockKind::Paragraph { paragraph },
        };
        Block {
            id: String::new(),
            anchor,
            content,
        }
    }

    /// The paragraph's numbering, advancing the story's counters the way the renderer does.
    fn list_info(
        &mut self,
        values: &BTreeMap<String, Any>,
        list: &mut ListState,
        anchor: &Anchor,
    ) -> Option<ListInfo> {
        let num_pr = any_map(values.get("numPr"))?;
        let num_id = any_number(num_pr.get("numId")).filter(|id| *id != 0.0)?;
        let num_id = any_text(num_pr.get("numId")).unwrap_or_else(|| num_id.to_string());
        let abstract_num_id = any_text(values.get("listAbstractNumId"));
        let suffix = match any_text(values.get("listMarkerSuffix")).as_deref() {
            Some("space") => MarkerSuffix::Space,
            Some("nothing") => MarkerSuffix::Nothing,
            _ => MarkerSuffix::Tab,
        };
        let marker_hidden = matches!(values.get("listMarkerHidden"), Some(Any::Bool(true)));
        let level_value = any_number(num_pr.get("ilvl")).unwrap_or(0.0);
        let Some(level) = numbering_level(level_value) else {
            self.note(
                DiagnosticCode::UnsupportedNumbering,
                Severity::Warning,
                Some(anchor.clone()),
                format!(
                    "Numbering level {level_value} is outside 0..8, so the marker is unresolved."
                ),
            );
            return Some(ListInfo {
                num_id,
                abstract_num_id,
                level: if level_value > 8.0 { 8 } else { 0 },
                format: String::new(),
                marker: None,
                suffix,
                marker_hidden,
            });
        };
        let bullet = matches!(values.get("listIsBullet"), Some(Any::Bool(true)));
        let mut formats = list_level_formats(values);
        formats.truncate(9);
        let format = if bullet {
            Some("bullet".to_owned())
        } else {
            formats
                .get(level)
                .cloned()
                .or_else(|| any_text(values.get("listNumFmt")))
        };
        let rendered = list_marker(values, list);
        let template = any_text(values.get("listMarker"));
        let referenced: Vec<String> = match &template {
            Some(template) if !bullet => (1..=9usize)
                .filter(|level| template.contains(&format!("%{level}")))
                .filter_map(|level| formats.get(level - 1).cloned().or_else(|| format.clone()))
                .collect(),
            _ => format.iter().cloned().collect(),
        };
        let unsupported = referenced
            .iter()
            .find(|format| !renders_format(format))
            .cloned();
        let marker = match (&format, &unsupported) {
            (None, _) => {
                self.note(
                    DiagnosticCode::UnsupportedNumbering,
                    Severity::Warning,
                    Some(anchor.clone()),
                    format!("Numbering instance {num_id} has no resolved format, so its marker is unresolved."),
                );
                None
            }
            (_, Some(unsupported)) => {
                self.note(
                    DiagnosticCode::UnsupportedNumbering,
                    Severity::Warning,
                    Some(anchor.clone()),
                    format!("Numbering format {unsupported} is not rendered, so the marker is unresolved."),
                );
                None
            }
            _ => match rendered {
                Some(Err(Unrenderable { value, format })) => {
                    self.note(
                        DiagnosticCode::UnsupportedNumbering,
                        Severity::Warning,
                        Some(anchor.clone()),
                        format!("Number {value} cannot be written in format {format}, so the marker is unresolved."),
                    );
                    None
                }
                Some(Ok(marker)) => Some(marker),
                None => Some(String::new()),
            },
        };
        Some(ListInfo {
            num_id,
            abstract_num_id,
            level: level as u8,
            format: format.unwrap_or_default(),
            marker,
            suffix,
            marker_hidden,
        })
    }

    /// The inline for a record of source content placed back into a paragraph, if this view
    /// shows it.
    fn record_inline(&mut self, anchor: &Anchor, record: Record) -> Option<Inline> {
        let inline = match record {
            Record::Omitted { element, .. } => self.omitted_inline(anchor, element),
            Record::Break { kind, revision, .. } => {
                let (ins, del) = match revision.as_ref().map(|revision| revision.kind) {
                    Some(RevisionKind::Insertion | RevisionKind::MoveTo) => (true, false),
                    Some(RevisionKind::Deletion | RevisionKind::MoveFrom) => (false, true),
                    None => (false, false),
                };
                if self.hides(ins, del) {
                    return None;
                }
                Inline {
                    id: String::new(),
                    anchor: anchor.clone(),
                    marks: self.options.include_formatting.then(Vec::new),
                    link: None,
                    revisions: match self.view() {
                        RevisionView::Markup => revision.into_iter().collect(),
                        _ => Vec::new(),
                    },
                    content: InlineKind::Break { break_type: kind },
                }
            }
        };
        self.grow_inline(&inline);
        Some(inline)
    }

    /// The inlines of one paragraph's inline region.
    fn inlines<T: ReadTxn>(
        &mut self,
        views: &mut Views<'_, T>,
        ctx: &mut StoryCtx,
        index: usize,
        chunks: &[Chunk],
        depth: usize,
    ) -> Vec<Inline> {
        let accepted = Rc::clone(&ctx.accepted);
        let original = Rc::clone(&ctx.original);
        let paragraph = &accepted.paragraphs[index];
        let paragraph_anchor = ctx.paragraph_anchor(&paragraph.para_id);
        let mut records = ctx
            .placements
            .inline
            .remove(&paragraph.para_id)
            .unwrap_or_default();
        records.sort_by_key(|(raw, record)| (*raw, record.in_control()));
        let mut records = records.into_iter().peekable();
        let mut inlines = Vec::new();
        let mut controls = 0usize;
        let mut fields = 0usize;
        let mut projected: HashMap<i64, Vec<(i64, Vec<Inline>)>> = HashMap::new();
        for chunk in chunks {
            if !self.visit(1) {
                break;
            }
            let projected_as = any_map(chunk.attrs.get("fieldResult")).and_then(|projection| {
                Some((
                    any_number(projection.get("id"))? as i64,
                    any_number(projection.get("index"))? as i64,
                ))
            });
            let before = inlines.len();
            while let Some((raw, record)) = records.peek()
                && !record.in_control()
                && *raw <= chunk.start
            {
                let (_, record) = records.next().expect("peeked");
                inlines.extend(self.record_inline(&paragraph_anchor, record));
            }
            let mut in_control = Vec::new();
            while let Some((raw, record)) = records.peek()
                && record.in_control()
                && *raw == chunk.start
            {
                in_control.push(records.next().expect("peeked").1);
            }
            let (ins, del, stamps) =
                revisions_of(|key| chunk.attrs.get(key).cloned(), self.moves(&ctx.prov));
            if let ChunkKind::Embed(Some(map)) = &chunk.kind {
                let kind = map_string(map, views.txn(), KIND_KEY);
                controls += usize::from(kind.as_deref() == Some("sdt"));
                if kind.as_deref() == Some("field")
                    && numeric_field_instruction(
                        &map_string(map, views.txn(), "instruction").unwrap_or_default(),
                    )
                    && matches!(
                        map.get(views.txn(), "fieldResultBlocks"),
                        Some(Out::Any(Any::Array(_)))
                    )
                {
                    fields += 1;
                }
            }
            if self.hides(ins, del) || projected_as.is_some_and(|(_, index)| index < 0) {
                continue;
            }
            let view = match self.view() {
                RevisionView::Accepted => {
                    Some((&accepted.paragraphs[index], EditTextView::Accepted))
                }
                RevisionView::Original => {
                    Some((&original.paragraphs[index], EditTextView::Original))
                }
                RevisionView::Markup if ins && del => None,
                RevisionView::Markup if del => {
                    Some((&original.paragraphs[index], EditTextView::Original))
                }
                RevisionView::Markup => Some((&accepted.paragraphs[index], EditTextView::Accepted)),
            };
            let revisions = if self.view() == RevisionView::Markup {
                stamps
            } else {
                Vec::new()
            };
            let marks = self
                .options
                .include_formatting
                .then(|| marks_of(|key| chunk.attrs.get(key).cloned()));
            let link = link_of(chunk.attrs.get("hyperlink"));
            let range = |start: u32, end: u32| -> Anchor {
                match view {
                    Some((projection, view)) => ctx.anchor(Anchor::Range(TextRange {
                        story: ctx.story.clone(),
                        start: TextPosition {
                            para_id: projection.para_id.clone(),
                            offset: projection.offset_of_raw(start),
                        },
                        end: TextPosition {
                            para_id: projection.para_id.clone(),
                            offset: projection.offset_of_raw(end),
                        },
                        view,
                    })),
                    None => paragraph_anchor.clone(),
                }
            };
            match &chunk.kind {
                ChunkKind::Text(text) => {
                    let mut raw = chunk.start;
                    let mut rest = text.as_str();
                    loop {
                        let cut = match records.peek() {
                            Some((at, record))
                                if !record.in_control() && *at > raw && *at < chunk.end() =>
                            {
                                Some(*at)
                            }
                            _ => None,
                        };
                        let (piece, tail) = match cut {
                            Some(at) => split_utf16(rest, at - raw),
                            None => (rest, ""),
                        };
                        for (position, segment) in piece.split('\t').enumerate() {
                            if position > 0 {
                                if !self.grow(0, INLINE_OVERHEAD) {
                                    break;
                                }
                                inlines.push(Inline {
                                    id: String::new(),
                                    anchor: range(raw, raw + 1),
                                    marks: marks.clone(),
                                    link: link.clone(),
                                    revisions: revisions.clone(),
                                    content: InlineKind::Tab,
                                });
                                raw += 1;
                            }
                            if segment.is_empty() {
                                continue;
                            }
                            if !self.grow(0, segment.len() + INLINE_OVERHEAD) {
                                break;
                            }
                            let width = segment.encode_utf16().count() as u32;
                            inlines.push(Inline {
                                id: String::new(),
                                anchor: range(raw, raw + width),
                                marks: marks.clone(),
                                link: link.clone(),
                                revisions: revisions.clone(),
                                content: InlineKind::Text {
                                    text: segment.to_owned(),
                                },
                            });
                            raw += width;
                        }
                        rest = tail;
                        if self.stopped {
                            break;
                        }
                        let Some((_, record)) = cut.and_then(|_| records.next()) else {
                            break;
                        };
                        inlines.extend(self.record_inline(&paragraph_anchor, record));
                    }
                }
                ChunkKind::Pilcrow(_) => {}
                ChunkKind::Embed(map) => {
                    let anchor = range(chunk.start, chunk.start + 1);
                    let payload = map
                        .as_ref()
                        .map(|map| payload_of(map, views.txn()))
                        .unwrap_or_default();
                    let kind =
                        any_text(payload.get(KIND_KEY)).unwrap_or_else(|| "embed".to_owned());
                    let control_id = (kind == "sdt")
                        .then(|| inline_control_id(&ctx.story, &paragraph.para_id, controls - 1));
                    let hides_result = kind == "field"
                        && numeric_field_instruction(
                            &any_text(payload.get("instruction")).unwrap_or_default(),
                        )
                        && matches!(payload.get("fieldResultBlocks"), Some(Any::Array(_)));
                    if hides_result {
                        ctx.field_anchors
                            .insert((index, fields - 1), anchor.clone());
                    }
                    self.projected_result = any_map(payload.get("resultProjection"))
                        .and_then(|projection| any_number(projection.get("id")))
                        .map(|id| projected.remove(&(id as i64)).unwrap_or_default());
                    let atom = self.atom(
                        ctx,
                        &kind,
                        &payload,
                        Atom {
                            anchor,
                            marks: marks.clone(),
                            link: link.clone(),
                            revisions: revisions.clone(),
                        },
                        control_id,
                        in_control,
                        depth,
                    );
                    self.grow_inline(&atom);
                    inlines.push(atom);
                }
            }
            if let Some((id, result_index)) = projected_as {
                let moved = inlines.split_off(before);
                projected.entry(id).or_default().push((result_index, moved));
            }
            if self.stopped {
                break;
            }
        }
        for (_, record) in records {
            inlines.extend(self.record_inline(&paragraph_anchor, record));
        }
        canonicalize(inlines)
    }

    fn omitted_inline(&mut self, anchor: &Anchor, element: String) -> Inline {
        self.note(
            DiagnosticCode::UnsupportedContent,
            Severity::Warning,
            Some(anchor.clone()),
            format!(
                "An inline {element} is not represented in the export; it stays in the package."
            ),
        );
        Inline {
            id: String::new(),
            anchor: anchor.clone(),
            marks: self.options.include_formatting.then(Vec::new),
            link: None,
            revisions: Vec::new(),
            content: InlineKind::Unsupported {
                element,
                alt_text: None,
            },
        }
    }

    /// One inline atom from an embed of the stream or of a control's frozen content.
    #[allow(clippy::too_many_arguments)]
    fn atom(
        &mut self,
        ctx: &StoryCtx,
        kind: &str,
        payload: &Payload,
        base: Atom,
        control_id: Option<String>,
        records: Vec<Record>,
        depth: usize,
    ) -> Inline {
        let anchor = base.anchor.clone();
        let content = match kind {
            "break" => InlineKind::Break {
                break_type: BreakType::Line,
            },
            "noteRef" => {
                let (note_kind, id, prefix) = match any_text(payload.get("footnoteRefId")) {
                    Some(id) => (NoteKind::Footnote, id, "fn"),
                    None => (
                        NoteKind::Endnote,
                        any_text(payload.get("endnoteRefId")).unwrap_or_default(),
                        "en",
                    ),
                };
                let story = format!("{prefix}:{id}");
                let story = self.story_ids.contains(&story).then_some(story);
                if story.is_none() {
                    self.note(
                        DiagnosticCode::UnresolvedReference,
                        Severity::Warning,
                        Some(anchor.clone()),
                        format!(
                            "The note reference names note {id}, which the document does not hold."
                        ),
                    );
                }
                InlineKind::NoteReference {
                    note_kind,
                    note_id: id,
                    story,
                }
            }
            "field"
                if any_text(payload.get("modelKind")).as_deref() == Some("commentReference") =>
            {
                let id = any_text(payload.get("commentId")).unwrap_or_default();
                let story = self
                    .comment_ids
                    .contains(&id)
                    .then(|| format!("comment:{id}"));
                if story.is_none() {
                    self.note(
                        DiagnosticCode::UnresolvedReference,
                        Severity::Warning,
                        Some(anchor.clone()),
                        format!("The comment reference names comment {id}, which the document does not hold."),
                    );
                }
                InlineKind::CommentReference {
                    comment_id: id,
                    story,
                }
            }
            "field" => self.field(ctx, payload, &base, depth),
            "image" => self.image(ctx, payload, &anchor),
            "sdt" => {
                let control_id = control_id.unwrap_or_default();
                return self.inline_control(ctx, payload, base, control_id, records, depth);
            }
            _ => {
                let (element, alt_text) = match kind {
                    "shape" => {
                        let shape = any_str(payload.get("shapeJson"))
                            .filter(|json| self.fits(json.len()))
                            .and_then(|json| serde_json::from_str::<Value>(json).ok());
                        let text_box = shape.as_ref().is_some_and(|shape| {
                            shape
                                .pointer("/textBody/content")
                                .and_then(Value::as_array)
                                .is_some_and(|content| !content.is_empty())
                        });
                        let alt = shape.as_ref().and_then(|shape| {
                            ["description", "title", "name"]
                                .into_iter()
                                .find_map(|key| shape.get(key)?.as_str().map(str::to_owned))
                                .filter(|alt| !alt.is_empty())
                        });
                        (if text_box { "wps:txbx" } else { "wps:wsp" }, alt)
                    }
                    "chart" => ("c:chart", any_text(payload.get("title"))),
                    "math" => (
                        if any_text(payload.get("display")).as_deref() == Some("block") {
                            "m:oMathPara"
                        } else {
                            "m:oMath"
                        },
                        any_text(payload.get("plainText")).filter(|text| !text.is_empty()),
                    ),
                    "horizontalRule" => ("v:rect", None),
                    other => (other, None),
                };
                self.note(
                    DiagnosticCode::UnsupportedContent,
                    Severity::Warning,
                    Some(anchor.clone()),
                    format!("An inline {element} is not represented in the export."),
                );
                InlineKind::Unsupported {
                    element: element.to_owned(),
                    alt_text,
                }
            }
        };
        Inline {
            id: String::new(),
            anchor: base.anchor,
            marks: base.marks,
            link: base.link,
            revisions: base.revisions,
            content,
        }
    }

    /// An image, its relationship resolved against the part that owns the story it is in.
    fn image(&mut self, ctx: &StoryCtx, payload: &Payload, anchor: &Anchor) -> InlineKind {
        let relationship_id = any_text(payload.get("rId"));
        let mut target = None;
        if let (Some(id), Some(read), Some(part)) =
            (relationship_id.as_deref(), self.read(), ctx.part.as_deref())
        {
            match read.relationship(part, id) {
                Some(Some(found)) => target = Some(found.clone()),
                Some(None) => self.note(
                    DiagnosticCode::UnresolvedReference,
                    Severity::Warning,
                    Some(anchor.clone()),
                    format!("Part {part} has no relationship {id} for this image."),
                ),
                None => self.note(
                    DiagnosticCode::ProvenanceUnavailable,
                    Severity::Info,
                    Some(anchor.clone()),
                    format!("The relationships of part {part} are unavailable, so the image's target is unknown."),
                ),
            }
        }
        self.note(
            DiagnosticCode::ImageDataOmitted,
            Severity::Info,
            Some(anchor.clone()),
            "Image data is not exported; the image is described by its alt text and relationship.",
        );
        InlineKind::Image {
            alt_text: any_text(payload.get("alt"))
                .or_else(|| any_text(payload.get("title")))
                .filter(|alt| !alt.is_empty()),
            relationship_id,
            part: match &target {
                Some(RelationshipTarget::Part(part)) => Some(part.clone()),
                _ => None,
            },
            external_target: match target {
                Some(RelationshipTarget::External(url)) => Some(url),
                _ => None,
            },
        }
    }

    /// A field with its cached result, which is never evaluated. The result is read from the
    /// field's parsed content through the same lowering and walk as the stream, every node
    /// anchored to the field; result blocks a numeric field hides are attached as they are read.
    fn field(
        &mut self,
        ctx: &StoryCtx,
        payload: &Payload,
        base: &Atom,
        depth: usize,
    ) -> InlineKind {
        let field_type = any_text(payload.get("fieldType")).unwrap_or_default();
        let instruction = any_text(payload.get("instruction")).unwrap_or_default();
        let display = any_str(payload.get("displayText")).unwrap_or_default();
        let data = any_str(payload.get("fieldData"));
        if !self.fits(display.len()) || data.is_some_and(|json| !self.fits(json.len())) {
            self.oversize();
            return InlineKind::Field {
                field_type,
                instruction,
                cached_result: CachedResult::Missing,
                dirty: false,
                locked: false,
            };
        }
        let data = data.and_then(|json| serde_json::from_str::<Value>(json).ok());
        let has_nodes = |key: &str| {
            data.as_ref()
                .and_then(|data| data.get(key))
                .and_then(Value::as_array)
                .is_some_and(|nodes| !nodes.is_empty())
        };
        let has_result = match data.as_ref() {
            Some(data) => {
                data.get("structuredResult")
                    .is_some_and(|result| !result.is_null())
                    || has_nodes("fieldResult")
                    || (data.get("type").and_then(Value::as_str) == Some("simpleField")
                        && has_nodes("content"))
            }
            None => matches!(payload.get("hasCachedResult"), Some(Any::Bool(true))),
        } || present(payload.get("resultProjection")).is_some();
        let nodes = data.as_ref().map(result_nodes).unwrap_or_default();
        let projected = self.projected_result.take();
        let hides_blocks = numeric_field_instruction(&instruction)
            && matches!(payload.get("fieldResultBlocks"), Some(Any::Array(ids)) if !ids.is_empty());
        let cached_result = if hides_blocks {
            let mut blocks = Vec::new();
            if !nodes.is_empty()
                && let Some(block) = self.result_paragraph(ctx, &nodes, &base.anchor, depth, true)
            {
                blocks.push(block);
            }
            CachedResult::Blocks { blocks }
        } else if !has_result {
            CachedResult::Missing
        } else if let Some(projected) = projected {
            CachedResult::Inline {
                inlines: self.projected_result_inlines(ctx, &nodes, projected, &base.anchor, depth),
            }
        } else if simple_result(&nodes) {
            CachedResult::Inline {
                inlines: if display.is_empty() || !self.grow(0, display.len() + INLINE_OVERHEAD) {
                    Vec::new()
                } else {
                    vec![Inline {
                        id: String::new(),
                        anchor: ctx.anchor(base.anchor.clone()),
                        marks: base.marks.clone(),
                        link: base.link.clone(),
                        revisions: Vec::new(),
                        content: InlineKind::Text {
                            text: display.to_owned(),
                        },
                    }]
                },
            }
        } else {
            CachedResult::Inline {
                inlines: self.result_inlines(ctx, &nodes, &base.anchor, depth),
            }
        };
        let named = if field_type.is_empty() {
            "A field"
        } else {
            &field_type
        };
        self.note(
            DiagnosticCode::FieldCachedResult,
            Severity::Info,
            Some(base.anchor.clone()),
            format!("{named} field is exported with its cached result and is not evaluated."),
        );
        if matches!(cached_result, CachedResult::Missing) {
            self.note(
                DiagnosticCode::MissingFieldResult,
                Severity::Warning,
                Some(base.anchor.clone()),
                format!("{named} field has no cached result."),
            );
        }
        InlineKind::Field {
            field_type,
            instruction,
            cached_result,
            dirty: any_flag(payload.get("dirty")),
            locked: any_flag(payload.get("fldLock")),
        }
    }

    /// The inlines of a field result whose hyperlinks and nested fields the stream holds before
    /// the field: those come from the stream as they are now, the other result nodes from the
    /// field's parsed content, in result order and anchored to the field.
    fn projected_result_inlines(
        &mut self,
        ctx: &StoryCtx,
        nodes: &[Value],
        projected: Vec<(i64, Vec<Inline>)>,
        owner: &Anchor,
        depth: usize,
    ) -> Vec<Inline> {
        let mut by_index: BTreeMap<i64, Vec<Inline>> = BTreeMap::new();
        for (index, inlines) in projected {
            by_index.entry(index).or_default().extend(inlines);
        }
        let mut output = Vec::new();
        let mut pending: Vec<Value> = Vec::new();
        for (index, node) in nodes.iter().enumerate() {
            let kind = node.get("type").and_then(Value::as_str);
            if !matches!(kind, Some("hyperlink" | "simpleField")) {
                pending.push(node.clone());
                continue;
            }
            if !pending.is_empty() {
                output.extend(self.result_inlines(
                    ctx,
                    &std::mem::take(&mut pending),
                    owner,
                    depth,
                ));
            }
            for mut inline in by_index.remove(&(index as i64)).unwrap_or_default() {
                reanchor(&mut inline, owner);
                output.push(inline);
            }
        }
        if !pending.is_empty() {
            output.extend(self.result_inlines(ctx, &pending, owner, depth));
        }
        canonicalize(output)
    }

    fn result_inlines(
        &mut self,
        ctx: &StoryCtx,
        nodes: &[Value],
        owner: &Anchor,
        depth: usize,
    ) -> Vec<Inline> {
        self.result_paragraph(ctx, nodes, owner, depth, false)
            .map(|block| match block.content {
                BlockKind::Paragraph { paragraph }
                | BlockKind::Heading { paragraph, .. }
                | BlockKind::ListItem { paragraph, .. } => paragraph.inlines,
                _ => Vec::new(),
            })
            .unwrap_or_default()
    }

    /// The paragraph a field's inline result nodes lower to, read in a private document with
    /// every node anchored to `owner`.
    fn result_paragraph(
        &mut self,
        ctx: &StoryCtx,
        nodes: &[Value],
        owner: &Anchor,
        depth: usize,
        output: bool,
    ) -> Option<Block> {
        if depth + 1 > MAX_DEPTH {
            self.note(
                DiagnosticCode::UnsupportedContent,
                Severity::Warning,
                Some(owner.clone()),
                format!("A field result nested more than {MAX_DEPTH} levels deep is not exported."),
            );
            return None;
        }
        let story = format!("field:{}", self.scratch_stories);
        self.scratch_stories += 1;
        let paragraph = serde_json::json!({
            "type": "paragraph",
            "formatting": { "styleId": ctx.paragraph_style },
            "content": nodes,
        });
        let (doc, provenance) = match self.seed_scratch(&story, &[paragraph]) {
            Ok(seeded) => seeded,
            Err(_) => {
                self.note(
                    DiagnosticCode::UnsupportedContent,
                    Severity::Warning,
                    Some(owner.clone()),
                    "The field's cached result could not be read.",
                );
                return None;
            }
        };
        let mut list = ListState::new(self.numbering());
        let prov = Prov::Scratch(Rc::new(ScratchProv {
            provenance,
            raw_anchors: HashMap::new(),
        }));
        let before = self.building.0;
        self.wrappers += usize::from(!output);
        self.reading += 1;
        let built = {
            let txn = doc.yrs_doc().transact();
            let mut views = Views::new(&doc, &txn);
            self.story_blocks(
                &mut views,
                &story,
                prov,
                Some(ctx.anchor(owner.clone())),
                ctx.part.clone(),
                depth + 1,
                &mut list,
                false,
            )
        };
        self.reading -= 1;
        if !output {
            self.wrappers -= 1;
            self.building.0 = before;
        }
        let mut blocks = built.into_iter().map(|built| {
            self.pending.extend(built.diagnostics);
            built.block
        });
        blocks.find(|block| {
            matches!(
                block.content,
                BlockKind::Paragraph { .. }
                    | BlockKind::Heading { .. }
                    | BlockKind::ListItem { .. }
            )
        })
    }

    /// An inline content control and its frozen content, whose children carry the control's
    /// anchor and follow the view's revision projection.
    fn inline_control(
        &mut self,
        ctx: &StoryCtx,
        payload: &Payload,
        base: Atom,
        control_id: String,
        records: Vec<Record>,
        depth: usize,
    ) -> Inline {
        let control_anchor = ctx.anchor(Anchor::Control {
            story: ctx.story.clone(),
            control_id: control_id.clone(),
        });
        let control = control_metadata(control_id.clone(), |key| payload.get(key));
        if present(payload.get("value")).is_some() {
            self.note(
                DiagnosticCode::LegacyControlValue,
                Severity::Warning,
                Some(control_anchor.clone()),
                "An authored value replaces this control's content when the document is saved; the export shows its current content.",
            );
        }
        let mut inlines = Vec::new();
        let mut nested = 0usize;
        let items: &[Any] = match payload.get("content") {
            Some(Any::Array(items)) => items,
            _ => &[],
        };
        let (positioned, trailing): (Vec<Record>, Vec<Record>) = records
            .into_iter()
            .partition(|record| record.control_offset().is_some());
        let mut positioned = positioned;
        positioned.sort_by_key(|record| record.control_offset());
        let mut positioned = positioned.into_iter().peekable();
        let mut offset = 0u32;
        for item in items {
            if self.stopped {
                break;
            }
            let Some(item) = any_map(Some(item)) else {
                continue;
            };
            let kind = any_text(item.get("kind")).unwrap_or_default();
            let text = any_str(item.get("text")).unwrap_or_default();
            let width = if kind == "text" {
                text.encode_utf16().count() as u32
            } else {
                1
            };
            let mut nested_records = Vec::new();
            while let Some(record) =
                positioned.next_if(|record| record.control_offset() <= Some(offset))
            {
                if record.nested() && record.control_offset() == Some(offset) {
                    nested_records.push(record.descend());
                } else {
                    inlines.extend(self.record_inline(&control_anchor, record));
                }
            }
            let attrs = any_map(item.get("attrs"));
            let attr = |key: &str| attrs.and_then(|attrs| attrs.get(key)).cloned();
            let (ins, del, stamps) = revisions_of(attr, self.moves(&ctx.prov));
            if self.hides(ins, del) {
                offset += width;
                continue;
            }
            let child = Atom {
                anchor: control_anchor.clone(),
                marks: self.options.include_formatting.then(|| marks_of(attr)),
                link: link_of(attrs.and_then(|attrs| attrs.get("hyperlink"))),
                revisions: match self.view() {
                    RevisionView::Markup => stamps,
                    _ => Vec::new(),
                },
            };
            match kind.as_str() {
                "text" => {
                    let mut at = offset;
                    let mut rest = text;
                    loop {
                        let cut = positioned
                            .peek()
                            .and_then(Record::control_offset)
                            .filter(|cut| *cut > at && *cut < offset + width);
                        let (piece, tail) = match cut {
                            Some(cut) => split_utf16(rest, cut - at),
                            None => (rest, ""),
                        };
                        if !self.grow(0, piece.len() + INLINE_OVERHEAD) {
                            break;
                        }
                        for (position, segment) in piece.split('\t').enumerate() {
                            if position > 0 {
                                inlines.push(child.inline(InlineKind::Tab));
                            }
                            if !segment.is_empty() {
                                inlines.push(child.inline(InlineKind::Text {
                                    text: segment.to_owned(),
                                }));
                            }
                        }
                        let Some(cut) = cut else {
                            break;
                        };
                        at = cut;
                        rest = tail;
                        while let Some(record) = positioned.next_if(|record| {
                            record.control_offset() == Some(cut) && !record.nested()
                        }) {
                            inlines.extend(self.record_inline(&control_anchor, record));
                        }
                    }
                }
                "tab" => inlines.push(child.inline(InlineKind::Tab)),
                kind => {
                    let payload: Payload =
                        any_map(item.get("payload")).cloned().unwrap_or_default();
                    let nested_id = (kind == "sdt").then(|| {
                        nested += 1;
                        nested_control_id(&control_id, nested - 1)
                    });
                    if depth < MAX_DEPTH {
                        let records = std::mem::take(&mut nested_records);
                        let atom =
                            self.atom(ctx, kind, &payload, child, nested_id, records, depth + 1);
                        self.grow_inline(&atom);
                        inlines.push(atom);
                    }
                }
            }
            for record in nested_records {
                inlines.extend(self.record_inline(&control_anchor, record));
            }
            offset += width;
        }
        for record in positioned.chain(trailing) {
            inlines.extend(self.record_inline(&control_anchor, record));
        }
        Inline {
            id: String::new(),
            anchor: base.anchor,
            marks: base.marks,
            link: base.link,
            revisions: base.revisions,
            content: InlineKind::ContentControl {
                control,
                inlines: canonicalize(inlines),
            },
        }
    }
}

/// The shared attributes of an atom and its children.
#[derive(Clone)]
struct Atom {
    anchor: Anchor,
    marks: Option<Vec<FormattingMark>>,
    link: Option<Link>,
    revisions: Vec<Revision>,
}

impl Atom {
    fn inline(&self, content: InlineKind) -> Inline {
        Inline {
            id: String::new(),
            anchor: self.anchor.clone(),
            marks: self.marks.clone(),
            link: self.link.clone(),
            revisions: self.revisions.clone(),
            content,
        }
    }
}

/// Emits the continuation records of the vertical merges covering `column` onwards.
fn continuations(
    covered: &mut BTreeMap<u32, (u32, u32, u32, u32)>,
    column: &mut u32,
    anchor: &Anchor,
    cells: &mut Vec<TableCell>,
) {
    while let Some((origin_row, origin_column, span, remaining)) = covered.get(column).copied() {
        if remaining == 0 || origin_column != *column {
            break;
        }
        for slot in *column..column.saturating_add(span) {
            if let Some(entry) = covered.get_mut(&slot) {
                entry.3 -= 1;
            }
        }
        cells.push(TableCell {
            anchor: anchor.clone(),
            story: None,
            column: *column,
            grid_span: span,
            row_span: 0,
            vertical_merge: VerticalMerge::Continue,
            merge_origin: Some(CellPosition {
                row: origin_row,
                column: origin_column,
            }),
            blocks: Vec::new(),
        });
        *column = column.saturating_add(span);
    }
}

/// The table id seeding gives a table embed, read from its first cell story.
fn table_identity(payload: &Payload) -> Option<String> {
    let Some(Any::Array(rows)) = payload.get("rows") else {
        return None;
    };
    rows.iter().find_map(|row| {
        let Some(Any::Array(cells)) = any_map(Some(row))?.get("cells") else {
            return None;
        };
        cells.iter().find_map(|cell| {
            any_text(any_map(Some(cell))?.get("story"))
                .and_then(|story| cell_table(&story).map(str::to_owned))
        })
    })
}

/// The ids of the blocks of a story: its paragraphs, tables and block controls.
/// Where a projection of the leading paragraphs of a story ends once the complete blocks
/// within `limit` after them are included, and the ids of the paragraphs and blocks after that.
fn later_ids<T: ReadTxn>(
    prefix: &StoryView,
    chunks: &[Chunk],
    limit: u32,
    txn: &T,
) -> (u32, HashSet<String>) {
    let mut cut = prefix
        .paragraphs
        .last()
        .map_or(0, |paragraph| paragraph.pilcrow + 1);
    let mut first = chunks.partition_point(|chunk| chunk.start < cut);
    for chunk in &chunks[first..] {
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            break;
        };
        let kind = map_string(map, txn, KIND_KEY).unwrap_or_default();
        if chunk.start != cut || chunk.end() > limit || !crate::segments::is_block_embed(&kind) {
            break;
        }
        cut = chunk.end();
        first += 1;
    }
    let mut ids = HashSet::new();
    for chunk in &chunks[first..] {
        match &chunk.kind {
            ChunkKind::Pilcrow(map) => ids.extend(map_string(map, txn, crate::PARA_ID)),
            ChunkKind::Embed(Some(map)) => {
                ids.extend(map_string(map, txn, "blockId"));
                ids.extend(map_string(map, txn, "story"));
                if map_string(map, txn, KIND_KEY).as_deref() == Some("table") {
                    ids.extend(table_identity(&payload_of(map, txn)));
                }
            }
            _ => {}
        }
    }
    (cut, ids)
}

fn alive_blocks<T: ReadTxn>(accepted: &StoryView, chunks: &[Chunk], txn: &T) -> HashSet<String> {
    let mut alive: HashSet<String> = accepted
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.para_id.clone())
        .collect();
    for chunk in chunks {
        let ChunkKind::Embed(Some(map)) = &chunk.kind else {
            continue;
        };
        match map_string(map, txn, KIND_KEY).as_deref() {
            Some("table") => alive.extend(table_identity(&payload_of(map, txn))),
            Some("blockSdt") => alive.extend(map_string(map, txn, "story")),
            _ => {}
        }
    }
    alive
}

/// Why a table's tracked cell or grid changes keep this view from reconstructing it.
fn unreconstructed_table(
    rows: &[HashMap<String, Any>],
    cells: &impl Fn(&HashMap<String, Any>) -> Vec<HashMap<String, Any>>,
    view: RevisionView,
) -> Option<&'static str> {
    let changed = |properties: Option<&HashMap<String, Any>>, key: &str, keys: &[&str]| {
        let Some(Any::Array(changes)) = properties.and_then(|properties| properties.get(key))
        else {
            return false;
        };
        let current =
            properties.and_then(|properties| any_map(properties.get("_originalFormatting")));
        changes.iter().any(|change| {
            let previous =
                any_map(Some(change)).and_then(|change| any_map(change.get("previousFormatting")));
            keys.iter().any(|key| {
                present(previous.and_then(|previous| previous.get(*key)))
                    != present(current.and_then(|current| current.get(*key)))
            })
        })
    };
    for row in rows {
        let row_properties = any_map(row.get("trPr"));
        if view != RevisionView::Accepted
            && changed(
                row_properties,
                "trPrChange",
                &["gridBefore", "gridAfter", "header"],
            )
        {
            return Some(
                "A tracked row property change alters the table's grid or header rows, which this view cannot reconstruct.",
            );
        }
        for cell in cells(row) {
            let properties = any_map(cell.get("tcPr"));
            if let Some(marker) =
                properties.and_then(|properties| any_map(properties.get("cellMarker")))
            {
                let inserted = any_text(marker.get("kind")).as_deref() == Some("ins");
                if view != RevisionView::Accepted || !inserted {
                    return Some(
                        "A tracked cell insertion, deletion or merge changes the table's grid, which this view cannot reconstruct.",
                    );
                }
            }
            if view != RevisionView::Accepted
                && changed(properties, "tcPrChange", &["gridSpan", "vMerge"])
            {
                return Some(
                    "A tracked cell property change alters the table's merges, which this view cannot reconstruct.",
                );
            }
        }
    }
    None
}

/// Whether tracked run formatting changes of the paragraph alter the marks the export reports:
/// unknown unless every changed run's previous formatting can be compared.
fn run_marks_changed(properties: &BTreeMap<String, Any>) -> bool {
    let Some(Any::Array(runs)) = properties.get("_originalRunBoundaries") else {
        return true;
    };
    let exported = |formatting: Option<&HashMap<String, Any>>| {
        marks_of(|key| {
            let formatting = formatting?;
            match key {
                "subscript" => (any_text(formatting.get("vertAlign")).as_deref()
                    == Some("subscript"))
                .then_some(Any::Bool(true)),
                "superscript" => (any_text(formatting.get("vertAlign")).as_deref()
                    == Some("superscript"))
                .then_some(Any::Bool(true)),
                key => formatting.get(key).cloned(),
            }
        })
    };
    let mut compared = false;
    for run in runs.iter() {
        let Some(run) = any_map(Some(run)) else {
            continue;
        };
        let Some(Any::Array(changes)) = run.get("propertyChanges") else {
            continue;
        };
        compared = true;
        let current = exported(any_map(run.get("formatting")));
        for change in changes.iter() {
            let previous =
                any_map(Some(change)).and_then(|change| any_map(change.get("previousFormatting")));
            if exported(previous) != current {
                return true;
            }
        }
    }
    !compared
}

/// The inline nodes of a field's cached result.
fn result_nodes(data: &Value) -> Vec<Value> {
    let nodes = |key: &str| {
        data.get(key)
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default()
    };
    let structured = data
        .pointer("/structuredResult/inline")
        .and_then(Value::as_array)
        .cloned();
    match data.get("type").and_then(Value::as_str) {
        Some("simpleField") => nodes("content"),
        _ => structured.unwrap_or_else(|| nodes("fieldResult")),
    }
}

/// Anchors an inline and everything inside it to `anchor`.
fn reanchor(inline: &mut Inline, anchor: &Anchor) {
    inline.anchor = anchor.clone();
    match &mut inline.content {
        InlineKind::ContentControl { inlines, .. }
        | InlineKind::Field {
            cached_result: CachedResult::Inline { inlines },
            ..
        } => {
            for inline in inlines {
                reanchor(inline, anchor);
            }
        }
        InlineKind::Field {
            cached_result: CachedResult::Blocks { blocks },
            ..
        } => {
            for block in blocks {
                reanchor_block(block, anchor);
            }
        }
        _ => {}
    }
}

fn reanchor_block(block: &mut Block, anchor: &Anchor) {
    block.anchor = anchor.clone();
    match &mut block.content {
        BlockKind::Paragraph { paragraph }
        | BlockKind::Heading { paragraph, .. }
        | BlockKind::ListItem { paragraph, .. } => {
            for inline in &mut paragraph.inlines {
                reanchor(inline, anchor);
            }
        }
        BlockKind::Table { table } => {
            for cell in table.rows.iter_mut().flat_map(|row| &mut row.cells) {
                cell.anchor = anchor.clone();
                for block in &mut cell.blocks {
                    reanchor_block(block, anchor);
                }
            }
        }
        BlockKind::ContentControl { blocks, .. } => {
            for block in blocks {
                reanchor_block(block, anchor);
            }
        }
        _ => {}
    }
}

/// Whether a result is one run of plain text, which the field's display text renders exactly.
fn simple_result(nodes: &[Value]) -> bool {
    match nodes {
        [] => true,
        [run] => {
            run.get("type").and_then(Value::as_str) == Some("run")
                && run
                    .get("content")
                    .and_then(Value::as_array)
                    .is_some_and(|content| {
                        content.iter().all(|item| {
                            item.get("type").and_then(Value::as_str) == Some("text")
                                && item
                                    .get("text")
                                    .and_then(Value::as_str)
                                    .is_some_and(|text| !text.is_empty() && !text.contains('\t'))
                        })
                    })
        }
        _ => false,
    }
}

/// Adds `block` to the `ordinal`-th hidden-block field result of `owner`, or hands it back.
fn attach_result(owner: &mut Block, ordinal: usize, block: Block) -> Option<Block> {
    let (BlockKind::Paragraph { paragraph }
    | BlockKind::Heading { paragraph, .. }
    | BlockKind::ListItem { paragraph, .. }) = &mut owner.content
    else {
        return Some(block);
    };
    let Some(InlineKind::Field { cached_result, .. }) = paragraph
        .inlines
        .iter_mut()
        .map(|inline| &mut inline.content)
        .filter(|content| {
            matches!(
                content,
                InlineKind::Field {
                    cached_result: CachedResult::Blocks { .. },
                    ..
                }
            )
        })
        .nth(ordinal)
    else {
        return Some(block);
    };
    let CachedResult::Blocks { blocks } = cached_result else {
        return Some(block);
    };
    blocks.push(block);
    None
}

/// Whether no other paragraph of `view` has the id `para_id`.
fn unique_id(view: &StoryView, para_id: &str) -> bool {
    view.paragraphs
        .iter()
        .filter(|paragraph| paragraph.para_id == para_id)
        .count()
        == 1
}

/// The paragraph ids more than one paragraph of `view` carries.
fn duplicate_ids(view: &StoryView) -> HashSet<String> {
    let mut seen = HashSet::new();
    view.paragraphs
        .iter()
        .filter(|paragraph| !seen.insert(paragraph.para_id.as_str()))
        .map(|paragraph| paragraph.para_id.clone())
        .collect()
}

/// Whether two stories and their table-cell and control stories hold the same content: every
/// unit, attribute and embed payload, revisions included. Only identities differ between two
/// copies of one part: the story ids that name child stories and blocks, and the paragraph ids
/// parsing mints afresh for the second copy, so those are set aside and nothing else is.
fn same_content<T: ReadTxn>(
    views: &Views<'_, T>,
    story_ids: &BTreeSet<String>,
    capacity: usize,
    left: &str,
    right: &str,
) -> bool {
    let (Some(left), Some(right)) = (
        content_fingerprint(views, story_ids, capacity, left),
        content_fingerprint(views, story_ids, capacity, right),
    ) else {
        return false;
    };
    left == right
}

/// The content of `story` and its child stories with identities set aside, or `None` when it
/// cannot be read within `capacity` stream units.
fn content_fingerprint<T: ReadTxn>(
    views: &Views<'_, T>,
    story_ids: &BTreeSet<String>,
    capacity: usize,
    story: &str,
) -> Option<Value> {
    let txn = views.txn();
    let prefix = format!("{story}:");
    let ids: Vec<&str> = std::iter::once(story)
        .chain(
            story_ids
                .iter()
                .map(String::as_str)
                .filter(|id| id.starts_with(&prefix)),
        )
        .collect();
    let mut units = 0usize;
    for id in &ids {
        let text = story_ref(txn, id).ok()?;
        units = units.saturating_add(yrs::Text::len(&text, txn) as usize);
    }
    if units > capacity {
        return None;
    }
    let mut stories = Vec::new();
    for id in ids {
        let text = story_ref(txn, id).ok()?;
        let chunks = views.doc().chunk_snapshot(id, &text, txn);
        let units: Vec<Value> = chunks
            .iter()
            .map(|chunk| {
                let content = match &chunk.kind {
                    ChunkKind::Text(text) => Value::String(text.clone()),
                    ChunkKind::Pilcrow(map) | ChunkKind::Embed(Some(map)) => {
                        let mut payload =
                            serde_json::to_value(payload_of(map, txn)).unwrap_or(Value::Null);
                        set_identities_aside(&mut payload, &prefix);
                        payload
                    }
                    ChunkKind::Embed(None) => Value::Null,
                };
                serde_json::json!([
                    content,
                    serde_json::to_value(&chunk.attrs).unwrap_or(Value::Null)
                ])
            })
            .collect();
        stories.push(serde_json::json!([&id[story.len()..], units]));
    }
    Some(Value::Array(stories))
}

/// Blanks the identity fields of an embed payload: paragraph ids, the block ids a field hides,
/// and the child story and block ids under `prefix`, which name the story they belong to.
fn set_identities_aside(value: &mut Value, prefix: &str) {
    match value {
        Value::Object(object) => {
            object.remove(crate::PARA_ID);
            if let Some(Value::Array(ids)) = object.get_mut("fieldResultBlocks") {
                ids.iter_mut().for_each(|id| *id = Value::Null);
            }
            for (key, value) in object.iter_mut() {
                match value {
                    Value::String(id) if matches!(key.as_str(), "story" | "blockId") => {
                        if let Some(rest) = id.strip_prefix(prefix) {
                            *id = format!(":{rest}");
                        }
                    }
                    value => set_identities_aside(value, prefix),
                }
            }
        }
        Value::Array(values) => {
            for value in values {
                set_identities_aside(value, prefix);
            }
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::SeedParagraph;

    fn package(body: &str) -> Vec<u8> {
        let rels = |entries: &str| {
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
            )
        };
        let parts = [
            (
                "[Content_Types].xml",
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned(),
            ),
            (
                "_rels/.rels",
                rels(r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>"#),
            ),
            ("word/_rels/document.xml.rels", rels("")),
            (
                "word/document.xml",
                format!(
                    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}</w:body></w:document>"#
                ),
            ),
        ];
        ooxml_opc::rezip_parts(
            &parts
                .into_iter()
                .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
                .collect::<Vec<_>>(),
        )
        .unwrap()
    }

    #[test]
    fn the_visit_limit_keeps_a_completed_field_owner() {
        let paragraph =
            |id: &str, content: &str| format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#);
        let body = [
            paragraph(
                "01000001",
                r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> 7 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>first</w:t></w:r>"#,
            ),
            paragraph("01000002", r#"<w:r><w:t>second</w:t></w:r>"#),
            paragraph("01000003", r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#),
            paragraph("01000004", r#"<w:r><w:t>after</w:t></w:r>"#),
        ]
        .concat();
        let doc = EditingDoc::new(2);
        crate::seed_from_docx(&doc, &package(&body)).unwrap();
        let options = Resolved::new(&ExportOptions::new(RevisionView::Accepted)).unwrap();
        let full = walk(&doc, &options, AnchorScope::Session, false);
        let full = &full.stories[0].blocks;
        assert_eq!(full.len(), 2);
        let mut counts = Vec::new();
        for limit in 1..40 {
            let source = doc.source_metadata();
            let mut exporter =
                Exporter::new(&options, source.as_deref(), AnchorScope::Session, false);
            exporter.budget.max_visited = limit;
            let content = walk_with(&doc, exporter);
            let blocks = &content.stories[0].blocks;
            assert_eq!(blocks.as_slice(), &full[..blocks.len()]);
            counts.push(blocks.len());
        }
        assert!(counts.windows(2).all(|pair| pair[0] <= pair[1]));
        assert!(
            counts.contains(&1),
            "some limit stops before the next paragraph yet keeps the completed owner: {counts:?}"
        );
    }

    #[test]
    fn a_story_past_the_budget_keeps_its_leading_blocks_without_a_full_projection() {
        let doc = EditingDoc::new(2);
        let paragraph = |text: String| SeedParagraph {
            text,
            p_style: "Normal".to_owned(),
            alignment: "left".to_owned(),
        };
        let mut paragraphs = vec![paragraph("Manual".to_owned())];
        paragraphs
            .extend((0..1_200).map(|index| paragraph(format!("{index:04} {}", "m".repeat(1_000)))));
        doc.seed_story("body", &paragraphs).unwrap();
        doc.add_comment(
            &[crate::StoryRange::new("body", 0, 6)],
            "Ann",
            "",
            Any::from("Note"),
        )
        .unwrap();
        for stories in [
            vec![StorySelection::Body],
            vec![StorySelection::Comments],
            vec![StorySelection::Body, StorySelection::Comments],
        ] {
            let options = Resolved::new(&ExportOptions {
                max_bytes: Some(4_096),
                stories: Some(stories.clone()),
                ..ExportOptions::new(RevisionView::Accepted)
            })
            .unwrap();
            let txn = doc.yrs_doc().transact();
            let mut exporter = Exporter::new(&options, None, AnchorScope::Session, false);
            exporter.story_ids = txn
                .get_map(crate::STORIES)
                .map(|stories| stories.keys(&txn).map(str::to_owned).collect())
                .unwrap_or_default();
            let mut views = Views::new(&doc, &txn);
            exporter.read_sections(&mut views);
            let planned = exporter.plan(&mut views);
            assert!(exporter.commit_notes());
            exporter.run(&mut views, planned);
            assert!(
                !views.built("body", EditTextView::Accepted),
                "{stories:?} projected the whole body"
            );
            let content = exporter.finish();
            assert_eq!(content.truncated, stories.contains(&StorySelection::Body));
            let comment = content
                .stories
                .iter()
                .find(|story| story.kind == StoryKind::Comment);
            if stories.contains(&StorySelection::Comments) && !content.truncated {
                assert!(matches!(
                    comment
                        .and_then(|story| story.comment.as_ref())
                        .map(|comment| comment.anchors.as_slice()),
                    Some([Anchor::Unlocated {
                        reason: UnlocatedReason::StoryTooLarge,
                        ..
                    }])
                ));
                assert_eq!(comment.map(|story| story.blocks.len()), Some(1));
            }
            if stories.contains(&StorySelection::Body) {
                let first = &content.stories[0].blocks[0];
                assert!(matches!(
                    &first.content,
                    BlockKind::Paragraph { paragraph }
                        if matches!(&paragraph.inlines[..], [Inline { content: InlineKind::Text { text }, .. }] if text == "Manual")
                ));
            }
        }
    }

    #[test]
    fn stopping_at_the_visit_limit_keeps_the_committed_prefix() {
        let doc = EditingDoc::new(2);
        let paragraph = |text: &str| SeedParagraph {
            text: text.to_owned(),
            p_style: "Normal".to_owned(),
            alignment: "left".to_owned(),
        };
        doc.seed_story("body", &["One", "Two", "Three", "Four"].map(paragraph))
            .unwrap();
        let options = Resolved::new(&ExportOptions::new(RevisionView::Accepted)).unwrap();
        let mut exporter = Exporter::new(&options, None, AnchorScope::Session, false);
        exporter.budget.max_visited = 5;
        let content = walk_with(&doc, exporter);
        assert!(content.truncated);
        let texts: Vec<String> = content.stories[0]
            .blocks
            .iter()
            .map(|block| match &block.content {
                BlockKind::Paragraph { paragraph } => paragraph
                    .inlines
                    .iter()
                    .map(|inline| match &inline.content {
                        InlineKind::Text { text } => text.as_str(),
                        _ => "",
                    })
                    .collect(),
                _ => String::new(),
            })
            .collect();
        assert!(!texts.is_empty() && texts.len() < 4, "{texts:?}");
        assert_eq!(texts[0], "One");
    }
}
