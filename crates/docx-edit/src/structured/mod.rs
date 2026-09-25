//! Read-only structured export of DOCX content: ordered stories of typed blocks and inlines, each
//! carrying the location it was read from, plus diagnostics for everything the export omits or
//! cannot represent. [`markdown`] renders the same content as Markdown.
//!
//! One walker reads every source: a live session ([`EditingDoc::export_structured`]), DOCX bytes
//! ([`export_docx_structured`]) and a parsed package ([`export_package_structured`], which the
//! native facade uses for its current model). Bytes and packages seed a private session, so all
//! three read the same editing stream and retained package context.

mod markdown;
pub(crate) mod source;
mod walk;

use std::fmt;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize};

use crate::EditingDoc;
use crate::batch::DocumentVersion;

pub use crate::read_types::{Anchor, ControlMetadata, HeadingInfo, OutlineSource, StorySelection};
pub use markdown::render_docx_markdown;

/// The only structured-content schema version this crate reads and writes.
pub const SCHEMA_VERSION: u8 = 1;
/// Blocks an export returns when the options name no limit.
pub const DEFAULT_MAX_BLOCKS: u32 = 10_000;
/// Serialized bytes an export or rendering returns when the options name no limit.
pub const DEFAULT_MAX_BYTES: u32 = 8_388_608;
/// The largest block limit an export accepts.
pub const MAX_BLOCKS_LIMIT: u32 = 1_000_000;
/// The largest byte limit an export or rendering accepts.
pub const MAX_BYTES_LIMIT: u32 = 67_108_864;
/// The smallest byte limit an export or rendering accepts: room for the envelope and one
/// truncation diagnostic.
pub const MIN_BYTES_LIMIT: u32 = 1_024;

/// Which revision projection an export shows.
#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum RevisionView {
    /// Pending insertions included, pending deletions excluded.
    Accepted,
    /// Pending deletions included, pending insertions excluded.
    Original,
    /// Both, each inline carrying its revision attribution.
    Markup,
}

/// Whether anchors resolve against a live session version or only against the returned
/// snapshot.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum AnchorScope {
    Session,
    Snapshot,
}

/// A structured export. Ids are deterministic export-tree paths, not document identities.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocxStructuredContent {
    #[serde(deserialize_with = "schema_version")]
    pub schema_version: u8,
    pub revision_view: RevisionView,
    pub anchor_scope: AnchorScope,
    pub included_stories: Vec<StorySelection>,
    pub include_formatting: bool,
    pub stories: Vec<ExportStory>,
    pub diagnostics: Vec<Diagnostic>,
    pub truncated: bool,
}

fn schema_version<'de, D: Deserializer<'de>>(deserializer: D) -> Result<u8, D::Error> {
    let version = u8::deserialize(deserializer)?;
    if version != SCHEMA_VERSION {
        return Err(D::Error::custom(format!(
            "structured content schema version {version} is not supported; expected {SCHEMA_VERSION}"
        )));
    }
    Ok(version)
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum StoryKind {
    Body,
    Header,
    Footer,
    Footnote,
    Endnote,
    Comment,
}

/// One exported story. Table cells and block controls are nested inside the blocks that own
/// them, never exported as stories of their own.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportStory {
    pub story: String,
    pub kind: StoryKind,
    pub part: Option<String>,
    pub note_id: Option<String>,
    pub comment: Option<CommentMetadata>,
    /// The sections that reference a header or footer part, with the reference's variant.
    pub uses: Vec<StoryUse>,
    pub blocks: Vec<Block>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryUse {
    pub section_index: u32,
    pub variant: HeaderFooterVariant,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum HeaderFooterVariant {
    Default,
    First,
    Even,
}

/// A comment's metadata; `anchors` are the ranges it annotates.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CommentMetadata {
    pub id: String,
    pub author: Option<String>,
    pub date: Option<String>,
    pub parent_id: Option<String>,
    pub resolved: bool,
    pub anchors: Vec<Anchor>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Block {
    pub id: String,
    pub anchor: Anchor,
    #[serde(flatten)]
    pub content: BlockKind,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum BlockKind {
    Paragraph {
        paragraph: ParagraphData,
    },
    Heading {
        paragraph: ParagraphData,
        heading: HeadingInfo,
    },
    /// A numbered heading is a list item with a heading.
    ListItem {
        paragraph: ParagraphData,
        list: ListInfo,
        heading: Option<HeadingInfo>,
    },
    Table {
        table: TableData,
    },
    ContentControl {
        control: ControlMetadata,
        story: Option<String>,
        blocks: Vec<Block>,
    },
    /// The break that starts section `section_index`, typed by that section's start.
    SectionBreak {
        section_index: u32,
        break_type: SectionBreakType,
    },
    /// A page or column break between blocks.
    Break {
        break_type: BreakType,
    },
    Unsupported {
        element: String,
    },
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphData {
    pub style_id: Option<String>,
    pub inlines: Vec<Inline>,
}

/// A list paragraph's numbering. `format` is the OOXML numbering format; `marker` is the
/// rendered marker, `None` when it could not be resolved and `Some("")` when it is empty.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ListInfo {
    pub num_id: String,
    pub abstract_num_id: Option<String>,
    pub level: u8,
    pub format: String,
    pub marker: Option<String>,
    pub suffix: MarkerSuffix,
    pub marker_hidden: bool,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum MarkerSuffix {
    Tab,
    Space,
    Nothing,
}

/// A table on its zero-based grid. Every covered grid position of a row appears once: as a cell
/// that owns content, or as a vertical-merge continuation (`row_span: 0`) naming its origin.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableData {
    pub grid_columns: u32,
    pub rows: Vec<TableRow>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    pub header: bool,
    pub grid_before: u32,
    pub grid_after: u32,
    pub cells: Vec<TableCell>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    pub anchor: Anchor,
    pub story: Option<String>,
    pub column: u32,
    pub grid_span: u32,
    pub row_span: u32,
    pub vertical_merge: VerticalMerge,
    pub merge_origin: Option<CellPosition>,
    pub blocks: Vec<Block>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellPosition {
    pub row: u32,
    pub column: u32,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum VerticalMerge {
    None,
    Restart,
    Continue,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum SectionBreakType {
    Continuous,
    NextPage,
    NextColumn,
    OddPage,
    EvenPage,
}

/// One inline. Text, tabs and atoms carry range anchors into the projected text of their
/// view: tabs are `\t`, every other atom one U+FFFC. Children of an atom (a control's content,
/// a field's cached result) carry the atom's own anchor.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Inline {
    pub id: String,
    pub anchor: Anchor,
    /// `None` when the export excludes formatting.
    pub marks: Option<Vec<FormattingMark>>,
    pub link: Option<Link>,
    /// Revision attribution; only the markup view carries any.
    pub revisions: Vec<Revision>,
    #[serde(flatten)]
    pub content: InlineKind,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum InlineKind {
    Text {
        text: String,
    },
    Tab,
    Break {
        break_type: BreakType,
    },
    NoteReference {
        note_kind: NoteKind,
        note_id: String,
        story: Option<String>,
    },
    CommentReference {
        comment_id: String,
        story: Option<String>,
    },
    /// A field, exported with its cached result and never evaluated.
    Field {
        field_type: String,
        instruction: String,
        cached_result: CachedResult,
        dirty: bool,
        locked: bool,
    },
    Image {
        alt_text: Option<String>,
        relationship_id: Option<String>,
        part: Option<String>,
        external_target: Option<String>,
    },
    ContentControl {
        control: ControlMetadata,
        inlines: Vec<Inline>,
    },
    Unsupported {
        element: String,
        alt_text: Option<String>,
    },
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum BreakType {
    Line,
    Page,
    Column,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum NoteKind {
    Footnote,
    Endnote,
}

/// A field's cached result: absent, inline content, or the blocks it spans.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum CachedResult {
    Missing,
    Inline { inlines: Vec<Inline> },
    Blocks { blocks: Vec<Block> },
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Link {
    pub href: String,
    pub title: Option<String>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum FormattingMark {
    Bold,
    Italic,
    Underline { style: String },
    Strike,
    Subscript,
    Superscript,
    Hidden,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Revision {
    pub kind: RevisionKind,
    pub id: Option<String>,
    pub author: Option<String>,
    pub date: Option<String>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum RevisionKind {
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Diagnostic {
    pub code: DiagnosticCode,
    pub severity: Severity,
    pub anchor: Option<Anchor>,
    pub message: String,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum Severity {
    Info,
    Warning,
    Error,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum DiagnosticCode {
    UnsupportedContent,
    UnsupportedRevision,
    UnsupportedNumbering,
    UnresolvedStyle,
    UnresolvedReference,
    ProvenanceUnavailable,
    AmbiguousIdentity,
    ImageDataOmitted,
    FieldCachedResult,
    MissingFieldResult,
    FormattingOmitted,
    StoriesOmitted,
    RevisionContentExcluded,
    MergeContinuationContentOmitted,
    LegacyControlValue,
    ParseWarning,
    Truncated,
    MarkdownLossy,
}

/// What to export. `stories` defaults to the body alone, `include_formatting` to true.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ExportOptions {
    pub revision_view: RevisionView,
    #[serde(default)]
    pub stories: Option<Vec<StorySelection>>,
    #[serde(default)]
    pub include_formatting: Option<bool>,
    #[serde(default)]
    pub max_blocks: Option<u32>,
    #[serde(default)]
    pub max_bytes: Option<u32>,
}

impl ExportOptions {
    pub fn new(revision_view: RevisionView) -> Self {
        Self {
            revision_view,
            stories: None,
            include_formatting: None,
            max_blocks: None,
            max_bytes: None,
        }
    }
}

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MarkdownOptions {
    #[serde(default)]
    pub max_bytes: Option<u32>,
}

/// Markdown rendered from structured content. Each `<!-- docx-export:N -->` marker precedes the
/// block it names; `anchors` maps markers to source anchors. Markdown does not preserve Word
/// pagination, typography or layout.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MarkdownContent {
    pub markdown: String,
    pub anchors: Vec<MarkdownAnchor>,
    pub diagnostics: Vec<Diagnostic>,
    pub truncated: bool,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MarkdownAnchor {
    pub marker: String,
    pub anchor: Anchor,
}

/// Why an export was refused: options out of range (`InvalidOptions`) or above a hard maximum
/// (`LimitExceeded`), or a session with no document content to export (`Unsupported`).
/// Unsupported document content is diagnosed, never refused.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum ExportFailureCode {
    InvalidOptions,
    LimitExceeded,
    Unsupported,
}

/// Why an export was refused, as data. `target` is the anchor the refusal concerns, if any.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportFailure {
    pub code: ExportFailureCode,
    pub target: Option<Box<Anchor>>,
    pub message: String,
}

impl fmt::Display for ExportFailure {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let code = match self.code {
            ExportFailureCode::InvalidOptions => "invalid-options",
            ExportFailureCode::LimitExceeded => "limit-exceeded",
            ExportFailureCode::Unsupported => "unsupported",
        };
        write!(f, "{code}: {}", self.message)
    }
}

impl std::error::Error for ExportFailure {}

/// A session export refusal; the document is at `version`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportRefusal {
    pub version: DocumentVersion,
    pub failure: ExportFailure,
}

/// Content read from a session together with the version it was read at.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportRead<T> {
    pub version: DocumentVersion,
    pub content: T,
}

/// Why a bytes or package export produced nothing.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExportError {
    /// The input is not a readable DOCX package.
    Parse(String),
    Refused(ExportFailure),
}

impl fmt::Display for ExportError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Parse(message) => f.write_str(message),
            Self::Refused(failure) => failure.fmt(f),
        }
    }
}

impl std::error::Error for ExportError {}

impl From<ExportFailure> for ExportError {
    fn from(failure: ExportFailure) -> Self {
        Self::Refused(failure)
    }
}

pub(crate) fn invalid_options(message: impl Into<String>) -> ExportFailure {
    ExportFailure {
        code: ExportFailureCode::InvalidOptions,
        target: None,
        message: message.into(),
    }
}

pub(crate) fn limit_exceeded(message: impl Into<String>) -> ExportFailure {
    ExportFailure {
        code: ExportFailureCode::LimitExceeded,
        target: None,
        message: message.into(),
    }
}

/// Validates a byte limit shared by exports and renderings.
pub(crate) fn byte_limit(requested: Option<u32>) -> Result<usize, ExportFailure> {
    let limit = requested.unwrap_or(DEFAULT_MAX_BYTES);
    if limit < MIN_BYTES_LIMIT {
        return Err(invalid_options(format!(
            "maxBytes must be at least {MIN_BYTES_LIMIT}"
        )));
    }
    if limit > MAX_BYTES_LIMIT {
        return Err(limit_exceeded(format!(
            "maxBytes may be at most {MAX_BYTES_LIMIT}"
        )));
    }
    Ok(limit as usize)
}

/// Export options with their defaults applied and validated.
pub(crate) struct Resolved {
    pub view: RevisionView,
    pub stories: Vec<StorySelection>,
    pub include_formatting: bool,
    pub max_blocks: usize,
    pub max_bytes: usize,
}

impl Resolved {
    pub fn new(options: &ExportOptions) -> Result<Self, ExportFailure> {
        let mut stories = options
            .stories
            .clone()
            .unwrap_or_else(|| vec![StorySelection::Body]);
        if stories.is_empty() {
            return Err(invalid_options("stories must name at least one category"));
        }
        stories.sort_unstable();
        stories.dedup();
        let max_blocks = options.max_blocks.unwrap_or(DEFAULT_MAX_BLOCKS);
        if max_blocks == 0 {
            return Err(invalid_options("maxBlocks must be at least 1"));
        }
        if max_blocks > MAX_BLOCKS_LIMIT {
            return Err(limit_exceeded(format!(
                "maxBlocks may be at most {MAX_BLOCKS_LIMIT}"
            )));
        }
        Ok(Self {
            view: options.revision_view,
            stories,
            include_formatting: options.include_formatting.unwrap_or(true),
            max_blocks: max_blocks as usize,
            max_bytes: byte_limit(options.max_bytes)?,
        })
    }
}

impl EditingDoc {
    /// Exports the committed state with the version it was read at. Anchors are scoped to that
    /// version. Nothing is mutated, minted or published.
    pub fn export_structured(
        &self,
        options: &ExportOptions,
    ) -> Result<ExportRead<DocxStructuredContent>, ExportRefusal> {
        let version = self.version();
        let refuse = |failure| ExportRefusal {
            version: version.clone(),
            failure,
        };
        let resolved = Resolved::new(options).map_err(refuse)?;
        if !walk::has_stories(self) {
            return Err(refuse(ExportFailure {
                code: ExportFailureCode::Unsupported,
                target: None,
                message: "The session holds no document content to export.".to_owned(),
            }));
        }
        let content = walk::export(self, &resolved, AnchorScope::Session);
        Ok(ExportRead { version, content })
    }

    /// [`EditingDoc::export_structured`] rendered as Markdown from that one read.
    pub fn export_markdown(
        &self,
        options: &ExportOptions,
    ) -> Result<ExportRead<MarkdownContent>, ExportRefusal> {
        let read = self.export_structured(options)?;
        let rendered = render_docx_markdown(
            &read.content,
            &MarkdownOptions {
                max_bytes: options.max_bytes,
            },
        )
        .map_err(|failure| ExportRefusal {
            version: read.version.clone(),
            failure,
        })?;
        Ok(ExportRead {
            version: read.version,
            content: rendered,
        })
    }
}

/// Exports DOCX bytes as a snapshot: anchors address the returned content only.
pub fn export_docx_structured(
    bytes: &[u8],
    options: &ExportOptions,
) -> Result<DocxStructuredContent, ExportError> {
    let resolved = Resolved::new(options)?;
    let (envelope, parts) =
        crate::seed::parse_docx_with_parts(bytes).map_err(ExportError::Parse)?;
    snapshot(envelope, Some(&parts), &resolved)
}

/// [`export_docx_structured`] rendered as Markdown.
pub fn export_docx_markdown(
    bytes: &[u8],
    options: &ExportOptions,
) -> Result<MarkdownContent, ExportError> {
    let content = export_docx_structured(bytes, options)?;
    Ok(render_docx_markdown(
        &content,
        &MarkdownOptions {
            max_bytes: options.max_bytes,
        },
    )?)
}

/// Exports a parsed package as a snapshot, reading its current content. `parts` are the inflated
/// parts of the DOCX it was parsed from; provenance that no longer matches them is diagnosed.
pub fn export_package_structured(
    envelope: docx_parse::S9WireEnvelope,
    parts: &[(String, Vec<u8>)],
    options: &ExportOptions,
) -> Result<DocxStructuredContent, ExportError> {
    let resolved = Resolved::new(options)?;
    let parts = source::SourceParts::new(parts.to_vec());
    snapshot(envelope, Some(&parts), &resolved)
}

/// The client id of the private sessions bytes and packages seed; exports never reveal it.
const SNAPSHOT_CLIENT_ID: u64 = 1;

fn snapshot(
    envelope: docx_parse::S9WireEnvelope,
    parts: Option<&source::SourceParts>,
    options: &Resolved,
) -> Result<DocxStructuredContent, ExportError> {
    let doc = EditingDoc::new(SNAPSHOT_CLIENT_ID);
    crate::seed::seed_parsed_docx_with(&doc, envelope, parts).map_err(ExportError::Parse)?;
    Ok(walk::export(&doc, options, AnchorScope::Snapshot))
}
