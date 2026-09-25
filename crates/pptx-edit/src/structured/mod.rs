//! Read-only structured export of PPTX content: slides in deck order, each with its shape tree,
//! text stories, tables and placeholders for content the export cannot represent, every record
//! carrying the location it was read from, plus diagnostics for everything omitted.
//! [`render_pptx_markdown`] renders the same content as Markdown.
//!
//! One walker reads every source: a live session ([`DeckSession::export_structured`]) and PPTX
//! bytes ([`export_pptx_structured`]), which open a private session first.
//!
//! Reading order is the current shape tree, depth first: slides in deck order, shapes in their
//! tree order, a group's descendants at the group's position, and table cells row by row. It is
//! the authored order, not a visual or accessibility reading order inferred from geometry.

mod markdown;
mod walk;

use std::fmt;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize};

use crate::batch::DocumentVersion;
use crate::{DeckSession, EditResult};

pub use markdown::render_pptx_markdown;

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

/// Whether anchors resolve against a live session version or only against the returned
/// snapshot.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum AnchorScope {
    Session,
    Snapshot,
}

/// The order records are exported in; see the module documentation.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ReadingOrder {
    ShapeTree,
}

/// A structured export. Ids are deterministic export-tree paths, not document identities.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PptxStructuredContent {
    #[serde(deserialize_with = "schema_version")]
    pub schema_version: u8,
    pub anchor_scope: AnchorScope,
    pub reading_order: ReadingOrder,
    pub included: IncludedContent,
    pub slides: Vec<ExportSlide>,
    pub diagnostics: Vec<ExportDiagnostic>,
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

/// The optional content an export was asked to include.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct IncludedContent {
    pub hidden_slides: bool,
    pub hidden_shapes: bool,
    pub notes: bool,
    pub comments: bool,
    pub formatting: bool,
}

/// A half-open range of UTF-16 offsets.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextSpan {
    pub start: u32,
    pub end: u32,
}

/// Where a record was read from. `text` ranges use the story offsets of
/// [`DeckSession::read_content`]; `notes` and `comment` ranges index the plain text they carry;
/// `sourcePart` addresses retained XML by element-child ordinals from the part's root.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum PptxAnchor {
    Slide {
        slide_id: String,
    },
    Shape {
        slide_id: String,
        shape_id: String,
    },
    Text {
        slide_id: String,
        shape_id: String,
        story_id: String,
        range: TextSpan,
    },
    Notes {
        slide_id: String,
        range: TextSpan,
    },
    Comment {
        slide_id: String,
        comment_id: String,
        range: TextSpan,
    },
    SourcePart {
        part: String,
        part_sha256: String,
        path: Vec<u32>,
    },
}

/// The retained source XML a record was seeded from. `path` holds element-child ordinals from
/// the part's root; `sld_id` is the owning slide's `p:sldId/@id` and `source_id` a shape's
/// `cNvPr/@id`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceProvenance {
    pub part: String,
    pub part_sha256: String,
    pub path: Vec<u32>,
    pub sld_id: Option<u32>,
    pub source_id: Option<u32>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportSlide {
    pub id: String,
    /// Zero-based position in the current deck, hidden slides counted.
    pub index: u32,
    pub anchor: PptxAnchor,
    pub name: Option<String>,
    /// `None` when the deck does not record it.
    pub hidden: Option<bool>,
    pub provenance: Option<SourceProvenance>,
    pub shapes: Vec<ExportShape>,
    pub notes: Option<ExportNotes>,
    pub comments: Vec<ExportComment>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ExportShapeKind {
    Shape,
    Picture,
    GraphicFrame,
    Group,
    /// A shape-tree element only the source holds, anchored by its source location.
    Unknown,
}

/// A shape, group or unrepresented shape-tree element, with its stories, table, placeholder
/// object and children.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportShape {
    pub id: String,
    pub anchor: PptxAnchor,
    pub kind: ExportShapeKind,
    pub name: String,
    /// `cNvPr/@title`, the alternative-text title.
    pub title: Option<String>,
    /// `cNvPr/@descr`, the alternative-text description.
    pub description: Option<String>,
    /// Hidden itself or through a hidden group.
    pub hidden: bool,
    pub placeholder: Option<ExportPlaceholder>,
    pub provenance: Option<SourceProvenance>,
    pub stories: Vec<ExportStory>,
    pub table: Option<ExportTable>,
    pub object: Option<ExportObject>,
    pub children: Vec<ExportShape>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportPlaceholder {
    #[serde(rename = "type")]
    pub placeholder_type: Option<String>,
    pub index: Option<u32>,
}

/// One text story; `anchor` spans all of it.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportStory {
    pub id: String,
    pub anchor: PptxAnchor,
    pub paragraphs: Vec<ExportParagraph>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportParagraph {
    pub id: String,
    pub anchor: PptxAnchor,
    pub paragraph_id: String,
    pub level: u32,
    /// The authored alignment, else the inherited one.
    pub alignment: Option<String>,
    /// The paragraph's own `a:buChar`, `a:buAutoNum` or `a:buNone`, as the session stores it.
    pub bullet_json: Option<String>,
    /// The list marker after inheritance and numbering; `None` for a paragraph that is not a
    /// list item.
    pub list: Option<ExportList>,
    pub runs: Vec<ExportRun>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum ExportList {
    /// `character` is the authored `a:buChar`; `font` its `a:buFont` typeface, a symbol font
    /// when the character names a glyph slot.
    Bullet {
        character: String,
        font: Option<String>,
    },
    /// `marker` is the formatted number, `None` for a scheme that cannot be formatted.
    Number {
        scheme: String,
        start_at: u32,
        value: u32,
        marker: Option<String>,
    },
}

/// A run of one kind and formatting. Line breaks are one `\n` unit and fields cover their cached
/// result; an unsupported inline is zero-width.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportRun {
    pub anchor: PptxAnchor,
    /// `None` when the export excludes formatting.
    pub marks: Option<Vec<ExportMark>>,
    pub link: Option<ExportLink>,
    #[serde(flatten)]
    pub content: ExportRunKind,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum ExportRunKind {
    Text {
        text: String,
    },
    LineBreak,
    /// A field's cached result; fields are never evaluated.
    Field {
        field_type: Option<String>,
        text: String,
    },
    Unsupported {
        element: String,
    },
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum ExportMark {
    Bold,
    Italic,
    Underline { style: String },
    Superscript,
    Subscript,
    SmallCaps,
    AllCaps,
}

/// A click hyperlink: an external URL, or the internal part it jumps to.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportLink {
    pub href: String,
    pub external: bool,
}

/// A table on its grid. Every source cell appears once; `merged` marks a continuation covered
/// by the cell at `merge_origin`, whose content PowerPoint does not show.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportTable {
    pub columns: u32,
    pub rows: Vec<ExportTableRow>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportTableRow {
    pub cells: Vec<ExportTableCell>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportTableCell {
    pub id: String,
    pub anchor: PptxAnchor,
    pub row: u32,
    pub column: u32,
    pub grid_span: u32,
    pub row_span: u32,
    pub merged: bool,
    pub merge_origin: Option<CellPosition>,
    pub story: Option<ExportStory>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellPosition {
    pub row: u32,
    pub column: u32,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ExportObjectKind {
    Picture,
    Video,
    Audio,
    Chart,
    SmartArt,
    EmbeddedObject,
    Unknown,
}

/// A placeholder for content the export does not represent; its data is never included.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportObject {
    pub id: String,
    pub kind: ExportObjectKind,
    /// The source element's qualified name, such as `p:graphicFrame`.
    pub element: String,
    /// `a:graphicData/@uri`.
    pub uri: Option<String>,
    pub relationship_ids: Vec<String>,
    /// The package parts the object references.
    pub parts: Vec<String>,
    /// Targets of the object's external relationships, such as a linked video.
    pub external_targets: Vec<String>,
}

/// A slide's speaker notes, as plain text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportNotes {
    pub id: String,
    pub anchor: PptxAnchor,
    pub text: String,
    pub provenance: Option<SourceProvenance>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportComment {
    pub id: String,
    pub anchor: PptxAnchor,
    pub comment_id: String,
    pub author: Option<String>,
    pub date: Option<String>,
    pub parent_id: Option<String>,
    pub resolved: bool,
    pub text: String,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportDiagnostic {
    pub code: ExportDiagnosticCode,
    pub severity: ExportSeverity,
    pub anchor: Option<PptxAnchor>,
    pub message: String,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ExportSeverity {
    Info,
    Warning,
    Error,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum ExportDiagnosticCode {
    UnsupportedContent,
    ImageDataOmitted,
    UnsupportedNumbering,
    FieldCachedResult,
    ProvenanceUnavailable,
    VisibilityUnknown,
    HiddenContentExcluded,
    StoriesOmitted,
    FormattingOmitted,
    InheritedContentOmitted,
    NotesStructureOmitted,
    MergeContinuationContentOmitted,
    Truncated,
    MarkdownLossy,
}

/// What to export. Hidden slides and shapes, notes and comments are excluded unless requested;
/// formatting is included unless turned off.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PptxExportOptions {
    #[serde(default)]
    pub include_hidden_slides: Option<bool>,
    #[serde(default)]
    pub include_hidden_shapes: Option<bool>,
    #[serde(default)]
    pub include_notes: Option<bool>,
    #[serde(default)]
    pub include_comments: Option<bool>,
    #[serde(default)]
    pub include_formatting: Option<bool>,
    #[serde(default)]
    pub max_blocks: Option<u32>,
    #[serde(default)]
    pub max_bytes: Option<u32>,
}

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PptxMarkdownOptions {
    #[serde(default)]
    pub max_bytes: Option<u32>,
}

/// Markdown rendered from structured content. Each `<!-- pptx-export:N -->` marker precedes the
/// record `anchors` maps it to. Markdown keeps no slide layout, geometry or typography.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PptxMarkdownContent {
    pub markdown: String,
    pub anchors: Vec<MarkdownAnchor>,
    pub diagnostics: Vec<ExportDiagnostic>,
    pub truncated: bool,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MarkdownAnchor {
    pub marker: String,
    pub anchor: PptxAnchor,
}

/// Why an export was refused: options out of range (`InvalidOptions`), above a hard maximum
/// (`LimitExceeded`), or structured content handed to the renderer that contradicts itself
/// (`InvalidContent`). Unsupported deck content is diagnosed, never refused.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum ExportFailureCode {
    InvalidOptions,
    LimitExceeded,
    InvalidContent,
}

/// Why an export was refused, as data. `target` is the anchor the refusal concerns, if any.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportFailure {
    pub code: ExportFailureCode,
    pub target: Option<Box<PptxAnchor>>,
    pub message: String,
}

impl fmt::Display for ExportFailure {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let code = match self.code {
            ExportFailureCode::InvalidOptions => "invalid-options",
            ExportFailureCode::LimitExceeded => "limit-exceeded",
            ExportFailureCode::InvalidContent => "invalid-content",
        };
        write!(f, "{code}: {}", self.message)
    }
}

impl std::error::Error for ExportFailure {}

/// A session export refusal; the deck is at `version`.
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

/// A session export: the content with its version, or a refusal.
pub type PptxExportResult<T> = Result<ExportRead<T>, ExportRefusal>;

/// Why a bytes export produced nothing.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExportError {
    /// The input is not a readable PPTX package.
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

fn invalid_options(message: impl Into<String>) -> ExportFailure {
    ExportFailure {
        code: ExportFailureCode::InvalidOptions,
        target: None,
        message: message.into(),
    }
}

fn limit_exceeded(message: impl Into<String>) -> ExportFailure {
    ExportFailure {
        code: ExportFailureCode::LimitExceeded,
        target: None,
        message: message.into(),
    }
}

/// Validates a byte limit shared by exports and renderings.
fn byte_limit(requested: Option<u32>) -> Result<usize, ExportFailure> {
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
    pub included: IncludedContent,
    pub max_blocks: usize,
    pub max_bytes: usize,
}

impl Resolved {
    fn new(options: &PptxExportOptions) -> Result<Self, ExportFailure> {
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
            included: IncludedContent {
                hidden_slides: options.include_hidden_slides.unwrap_or(false),
                hidden_shapes: options.include_hidden_shapes.unwrap_or(false),
                notes: options.include_notes.unwrap_or(false),
                comments: options.include_comments.unwrap_or(false),
                formatting: options.include_formatting.unwrap_or(true),
            },
            max_blocks: max_blocks as usize,
            max_bytes: byte_limit(options.max_bytes)?,
        })
    }
}

impl DeckSession {
    /// Exports the committed deck with the version it was read at; anchors are scoped to that
    /// version. Nothing is mutated, minted, published or laid out, and pending input is not
    /// flushed. Refusals are the inner `Err`.
    pub fn export_structured(
        &self,
        options: &PptxExportOptions,
    ) -> EditResult<PptxExportResult<PptxStructuredContent>> {
        let version = self.version();
        let resolved = match Resolved::new(options) {
            Ok(resolved) => resolved,
            Err(failure) => return Ok(Err(ExportRefusal { version, failure })),
        };
        let content = walk::export(self, &resolved, AnchorScope::Session)?;
        Ok(Ok(ExportRead { version, content }))
    }

    /// [`DeckSession::export_structured`] rendered as Markdown from that one read.
    pub fn export_markdown(
        &self,
        options: &PptxExportOptions,
    ) -> EditResult<PptxExportResult<PptxMarkdownContent>> {
        let read = match self.export_structured(options)? {
            Ok(read) => read,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let markdown_options = PptxMarkdownOptions {
            max_bytes: options.max_bytes,
        };
        Ok(
            match render_pptx_markdown(&read.content, &markdown_options) {
                Ok(content) => Ok(ExportRead {
                    version: read.version,
                    content,
                }),
                Err(failure) => Err(ExportRefusal {
                    version: read.version,
                    failure,
                }),
            },
        )
    }
}

/// The client id of the private session a bytes export opens; exports never reveal it.
const SNAPSHOT_CLIENT_ID: u64 = 1;

/// Exports PPTX bytes as a snapshot: anchors address the returned content only.
pub fn export_pptx_structured(
    bytes: &[u8],
    options: &PptxExportOptions,
) -> Result<PptxStructuredContent, ExportError> {
    let resolved = Resolved::new(options)?;
    let session = DeckSession::open(bytes, SNAPSHOT_CLIENT_ID)
        .map_err(|error| ExportError::Parse(error.to_string()))?;
    walk::export(&session, &resolved, AnchorScope::Snapshot)
        .map_err(|error| ExportError::Parse(error.to_string()))
}

/// JSON of a session export: `{"ok":true,"version","content"}` or
/// `{"ok":false,"version","failure"}`.
pub fn export_outcome_json<T: Serialize>(
    outcome: &PptxExportResult<T>,
) -> Result<String, serde_json::Error> {
    tagged(
        outcome.is_ok(),
        match outcome {
            Ok(read) => serde_json::to_value(read)?,
            Err(refusal) => serde_json::to_value(refusal)?,
        },
    )
}

/// JSON of a bytes export or rendering: `{"ok":true,"content"}` or `{"ok":false,"failure"}`.
/// `Err` is the message for input that is not a readable PPTX.
pub fn snapshot_outcome_json<T: Serialize>(
    outcome: Result<T, ExportError>,
) -> Result<String, String> {
    let value = match &outcome {
        Ok(content) => serde_json::json!({ "content": content }),
        Err(ExportError::Refused(failure)) => serde_json::json!({ "failure": failure }),
        Err(ExportError::Parse(message)) => return Err(message.clone()),
    };
    tagged(outcome.is_ok(), value).map_err(|error| error.to_string())
}

fn tagged(ok: bool, mut value: serde_json::Value) -> Result<String, serde_json::Error> {
    if let serde_json::Value::Object(object) = &mut value {
        object.insert("ok".to_owned(), serde_json::Value::Bool(ok));
    }
    serde_json::to_string(&value)
}

/// [`export_pptx_structured`] rendered as Markdown.
pub fn export_pptx_markdown(
    bytes: &[u8],
    options: &PptxExportOptions,
) -> Result<PptxMarkdownContent, ExportError> {
    let content = export_pptx_structured(bytes, options)?;
    Ok(render_pptx_markdown(
        &content,
        &PptxMarkdownOptions {
            max_bytes: options.max_bytes,
        },
    )?)
}
