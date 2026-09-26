//! Read-only structured export of workbook content with positional anchors, and its
//! Markdown projection. One walker serves live and bytes exports; neither recalculates.
//!
//! Formula results are exported as stored (`calculation.policy: "asStored"`): a stored
//! result may be stale, so freshness is `"unverified"` and every sheet exporting formulas
//! carries a `formula-cache-unverified` diagnostic. Cells report a missing result, a cycle
//! or an evaluation limit when the workbook knows of one.

mod markdown;
mod walk;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde_json::Value;

use crate::workbook::batch::{DocumentVersion, MAX_REQUEST_BYTES, MAX_RESPONSE_BYTES};
use crate::{Error, Result, Workbook};
use xlsx_model::ErrorValue;

pub use markdown::render_xlsx_markdown;
pub(crate) use walk::ExportSource;

pub const DEFAULT_EXPORT_MAX_CELLS: u32 = 100_000;
pub const MAX_EXPORT_CELLS: u32 = 1_000_000;
pub const DEFAULT_EXPORT_MAX_BYTES: u32 = 8_388_608;
/// The largest `maxBytes` an export or Markdown rendering accepts: 16 MiB.
pub const MAX_EXPORT_BYTES: u32 = 16 * 1024 * 1024;
pub const DEFAULT_MARKDOWN_MAX_ROWS: u32 = 200;
pub const DEFAULT_MARKDOWN_MAX_COLUMNS: u32 = 50;
pub const DEFAULT_MARKDOWN_MAX_CELLS: u32 = 10_000;
pub const MAX_MARKDOWN_ROWS: u32 = 100_000;
pub const MAX_MARKDOWN_CELLS: u32 = 1_000_000;
/// The smallest Markdown `maxBytes` accepted.
pub const MIN_MARKDOWN_BYTES: u32 = 4_096;
/// Largest content JSON [`render_xlsx_markdown_json`] decodes.
const MAX_RENDER_REQUEST_BYTES: usize = 2 * MAX_EXPORT_BYTES as usize;

/// Always `1` on the wire; any other value fails to decode.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct XlsxSchemaVersion;

impl Serialize for XlsxSchemaVersion {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        serializer.serialize_u8(1)
    }
}

impl<'de> Deserialize<'de> for XlsxSchemaVersion {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> std::result::Result<Self, D::Error> {
        match u64::deserialize(deserializer)? {
            1 => Ok(Self),
            other => Err(D::Error::custom(format!(
                "unsupported schemaVersion {other}; expected 1"
            ))),
        }
    }
}

/// One sheet to export: a zero-based index and at most one rectangular A1 range, which
/// defaults to the sheet's used range.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportScope {
    pub sheet: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<String>,
}

/// What to export. Hidden sheets, rows, columns and names are excluded unless asked for.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportOptions {
    /// Defaults to every sheet.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scope: Option<Vec<XlsxExportScope>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub include_hidden_sheets: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub include_hidden_rows: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub include_hidden_columns: Option<bool>,
    /// Defaults to `true`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub include_defined_names: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub include_hidden_names: Option<bool>,
    /// Stored cells to export, 1 to 1,000,000; defaults to 100,000.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_cells: Option<u32>,
    /// Compact UTF-8 JSON bytes of the content, diagnostics included; at most 16 MiB,
    /// defaults to 8 MiB.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_bytes: Option<u32>,
}

/// Bounds of a Markdown rendering. Empty grid positions count toward `max_cells`.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxMarkdownOptions {
    /// Grid rows per sheet; defaults to 200.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_rows: Option<u32>,
    /// Grid columns per sheet; defaults to 50.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_columns: Option<u32>,
    /// Grid positions across all sheets; defaults to 10,000.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_cells: Option<u32>,
    /// UTF-8 bytes of Markdown; defaults to 8 MiB.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_bytes: Option<u32>,
}

/// Whether anchors belong to the live session that produced them or to the snapshot a
/// bytes export read.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxAnchorScope {
    Session,
    Snapshot,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
pub enum XlsxDateSystem {
    #[serde(rename = "1900")]
    V1900,
    #[serde(rename = "1904")]
    V1904,
}

#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxCalculationPolicy {
    /// Formula results are the stored values; export never recalculates.
    #[default]
    AsStored,
}

#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxCalculationFreshness {
    /// Nothing establishes that stored results match their formulas.
    #[default]
    Unverified,
}

#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxCalculationState {
    pub policy: XlsxCalculationPolicy,
    pub freshness: XlsxCalculationFreshness,
}

/// The options an export applied.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportIncluded {
    pub hidden_sheets: bool,
    pub hidden_rows: bool,
    pub hidden_columns: bool,
    pub defined_names: bool,
    pub hidden_names: bool,
}

/// A sheet by the id edit batches take and, descriptively, its current position and name.
/// None is a durable identity.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxSheetIdentity {
    /// The exporting session's catalog id; `sheet:{index}` for bytes.
    pub sheet_id: String,
    pub index: u32,
    pub name: String,
}

/// Where exported content sits. A1 anchors name positions in the exported version or
/// snapshot, not cells that follow later row or column edits. A cell or range anchor's
/// `sheet.sheet_id` and `a1` are its batch [`RangeTarget`](crate::RangeTarget) at that version.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum XlsxAnchor {
    Sheet {
        sheet: XlsxSheetIdentity,
    },
    Cell {
        sheet: XlsxSheetIdentity,
        a1: String,
    },
    Range {
        sheet: XlsxSheetIdentity,
        a1: String,
    },
    DefinedName {
        name: String,
        local_sheet: Option<XlsxSheetIdentity>,
        /// Position in the workbook's defined-name list.
        ordinal: u32,
    },
    /// Retained source XML: the part, the SHA-256 of its bytes, and element-child ordinals
    /// from its root. Provenance, not an editable target.
    SourcePart {
        part: String,
        part_sha256: String,
        path: Vec<u32>,
    },
}

/// A retained source part.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxSourcePart {
    pub part: String,
    pub part_sha256: String,
    pub path: Vec<u32>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxSheetKind {
    Worksheet,
    Chartsheet,
    Dialogsheet,
    Macrosheet,
    Other,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxSheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
    Unknown,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportSheet {
    pub id: String,
    pub anchor: XlsxAnchor,
    pub kind: XlsxSheetKind,
    pub visibility: XlsxSheetVisibility,
    /// Stored cells and hyperlinks; drawings, merges and tables do not extend it.
    pub used_range: Option<String>,
    /// The requested range, else the used range.
    pub selected_range: Option<String>,
    /// Hidden row spans (`"5:7"`) inside the selected range, whether or not exported.
    pub hidden_rows: Vec<String>,
    /// Hidden column spans (`"C:D"`) inside the selected range, whether or not exported.
    pub hidden_columns: Vec<String>,
    pub merges: Vec<XlsxExportMerge>,
    pub tables: Vec<XlsxExportTable>,
    pub hyperlinks: Vec<XlsxExportHyperlink>,
    pub objects: Vec<XlsxExportObject>,
    /// Stored cells in row-major order. A position without a record is empty only before
    /// the last record of a truncated sheet.
    pub cells: Vec<XlsxExportCell>,
    pub source: Option<XlsxSourcePart>,
    /// Whether export stopped inside this sheet; its lists are complete in the order
    /// hidden spans, merges, tables, hyperlinks, objects, cells up to where it stopped.
    pub truncated: bool,
}

/// A cell value. Numbers stay numbers; dates are numbers read with `dateSystem`.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum XlsxExportValue {
    Empty,
    Number {
        value: f64,
    },
    Text {
        value: String,
    },
    Bool {
        value: bool,
    },
    Error {
        value: ErrorValue,
    },
    /// A value JSON cannot carry, such as a non-finite number.
    Unavailable {
        reason: String,
    },
}

/// What is known about a formula cell's stored result.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxFormulaResult {
    /// A result is stored; nothing verifies it is current.
    Unverified,
    /// The file stored no result, and nothing has calculated or changed since it was read.
    Missing,
    /// Whether a result is stored cannot be established: something has calculated or
    /// changed since the file was read, the cell cannot be traced to it, or it is empty
    /// inside an array formula's range, where empty can be a result.
    Uncertain,
    /// The last calculation settled it as part of a circular reference.
    Cycle,
    /// The last calculation left it at an evaluation limit.
    Limited,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxCellMerge {
    /// The whole merged range, which can extend past the selected range.
    pub range: String,
    /// Whether this is the top-left cell that displays the merge.
    pub origin: bool,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportCell {
    pub id: String,
    pub anchor: XlsxAnchor,
    pub value: XlsxExportValue,
    /// Formula source without the leading `=`.
    pub formula: Option<String>,
    /// The engine's formatting of the value, without column-width clipping or overflow.
    pub display_text: String,
    pub number_format: String,
    /// `None` for cells without a formula.
    pub formula_result: Option<XlsxFormulaResult>,
    pub merge: Option<XlsxCellMerge>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportMerge {
    pub id: String,
    pub anchor: XlsxAnchor,
    /// Whether the merge extends past the selected range.
    pub clipped: bool,
}

/// A ListObject.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportTable {
    pub id: String,
    pub anchor: XlsxAnchor,
    pub name: String,
    pub header_rows: u32,
    pub totals_rows: u32,
    pub columns: Vec<String>,
    pub clipped: bool,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportHyperlink {
    pub id: String,
    pub anchor: XlsxAnchor,
    pub external_target: Option<String>,
    /// A location inside the workbook.
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
    pub clipped: bool,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxObjectKind {
    Chart,
    Picture,
    Shape,
    Group,
    Connector,
    Diagram,
    GraphicFrame,
    ContentPart,
    Unknown,
}

/// A drawing object, exported as a placeholder without chart data, image bytes or shape
/// text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportObject {
    pub id: String,
    pub kind: XlsxObjectKind,
    /// The cells holding the object's corners; `None` when its grid position is unknown.
    pub anchor: Option<XlsxAnchor>,
    pub name: Option<String>,
    pub alt_text: Option<String>,
    pub title: Option<String>,
    pub hidden: bool,
    /// The chart or image part.
    pub part: Option<String>,
    pub source: Option<XlsxSourcePart>,
}

/// A defined name, listed read-only.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportDefinedName {
    pub id: String,
    pub anchor: XlsxAnchor,
    pub name: String,
    /// Without a leading `=`.
    pub formula: String,
    pub hidden: bool,
    pub local_sheet: Option<XlsxSheetIdentity>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum XlsxExportSeverity {
    Info,
    Warning,
    Error,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum XlsxExportDiagnosticCode {
    HiddenContentExcluded,
    VisibilityUnknown,
    UnsupportedSheet,
    UnsupportedContent,
    ObjectPlaceholder,
    UnreadableObject,
    ProvenanceUnavailable,
    CommentsOmitted,
    RichTextOmitted,
    FormattingApproximate,
    FormulaCacheUnverified,
    FormulaResultMissing,
    CalculationFailure,
    ValueUnavailable,
    Truncated,
    MarkdownLossy,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxExportDiagnostic {
    pub code: XlsxExportDiagnosticCode,
    pub severity: XlsxExportSeverity,
    pub anchor: Option<XlsxAnchor>,
    pub message: String,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxStructuredContent {
    pub schema_version: XlsxSchemaVersion,
    pub anchor_scope: XlsxAnchorScope,
    pub date_system: XlsxDateSystem,
    pub calculation: XlsxCalculationState,
    pub included: XlsxExportIncluded,
    /// In sheet order.
    pub sheets: Vec<XlsxExportSheet>,
    /// In workbook order.
    pub defined_names: Vec<XlsxExportDefinedName>,
    pub diagnostics: Vec<XlsxExportDiagnostic>,
    pub truncated: bool,
}

/// A Markdown marker comment and the anchor of the content it precedes.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxMarkdownAnchor {
    pub marker: String,
    pub anchor: XlsxAnchor,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct XlsxMarkdownContent {
    pub markdown: String,
    pub anchors: Vec<XlsxMarkdownAnchor>,
    /// The structured export's diagnostics, then what Markdown itself lost.
    pub diagnostics: Vec<XlsxExportDiagnostic>,
    pub truncated: bool,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum XlsxExportFailureCode {
    InvalidOptions,
    InvalidScope,
    LimitExceeded,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct XlsxExportFailure {
    pub code: XlsxExportFailureCode,
    /// The sheet a scope failure names, when it names one.
    pub target: Option<XlsxAnchor>,
    pub message: String,
}

/// A refused export; the workbook is untouched at `version`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct XlsxExportRefusal {
    pub version: DocumentVersion,
    pub failure: XlsxExportFailure,
}

/// Content read at `version`.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct XlsxExport<T> {
    pub version: DocumentVersion,
    pub content: T,
}

pub type XlsxExportResult<T> = std::result::Result<XlsxExport<T>, XlsxExportRefusal>;

/// The JSON wire form of a live export: the export or the refusal, tagged with `ok`.
pub fn export_result_json<T: Serialize>(result: &XlsxExportResult<T>) -> Result<String> {
    let mut value = match result {
        Ok(export) => serde_json::to_value(export),
        Err(refusal) => serde_json::to_value(refusal),
    }
    .map_err(|error| Error::InvalidOperation(error.to_string()))?;
    if let Value::Object(object) = &mut value {
        object.insert("ok".to_owned(), Value::Bool(result.is_ok()));
    }
    serde_json::to_string(&value).map_err(|error| Error::InvalidOperation(error.to_string()))
}

impl Workbook {
    /// Exports the committed workbook with the version it was read at. Nothing is
    /// recalculated, flushed or published.
    pub fn export_structured(
        &self,
        options: &XlsxExportOptions,
    ) -> Result<XlsxExportResult<XlsxStructuredContent>> {
        let version = self.version();
        Ok(
            match walk::export(&self.export_source(), options, XlsxAnchorScope::Session) {
                Ok(content) => Ok(XlsxExport { version, content }),
                Err(failure) => Err(XlsxExportRefusal { version, failure }),
            },
        )
    }

    /// [`Workbook::export_structured`] rendered as Markdown from the one captured read.
    pub fn export_markdown(
        &self,
        options: &XlsxExportOptions,
        markdown_options: &XlsxMarkdownOptions,
    ) -> Result<XlsxExportResult<XlsxMarkdownContent>> {
        let version = self.version();
        if let Err(failure) = markdown::limits(markdown_options) {
            return Ok(Err(XlsxExportRefusal { version, failure }));
        }
        Ok(match self.export_structured(options)? {
            Ok(export) => Ok(XlsxExport {
                version: export.version,
                content: render_xlsx_markdown(&export.content, markdown_options)?,
            }),
            Err(refusal) => Err(refusal),
        })
    }

    /// [`Workbook::export_structured`] over JSON. An oversized request or response is
    /// refused with `limit-exceeded`; a malformed request is [`Error::InvalidRequest`].
    pub fn export_structured_json(&self, options: &str) -> Result<String> {
        if let Some(refused) = self.oversized_export(options.len()) {
            return Ok(refused);
        }
        let result = self.export_structured(&decode(options)?)?;
        self.bounded_export(&result)
    }

    /// [`Workbook::export_markdown`] over JSON, bounded like
    /// [`Workbook::export_structured_json`].
    pub fn export_markdown_json(&self, options: &str, markdown_options: &str) -> Result<String> {
        if let Some(refused) = self.oversized_export(options.len() + markdown_options.len()) {
            return Ok(refused);
        }
        let result = self.export_markdown(&decode(options)?, &decode(markdown_options)?)?;
        self.bounded_export(&result)
    }

    fn oversized_export(&self, bytes: usize) -> Option<String> {
        (bytes > MAX_REQUEST_BYTES).then(|| {
            self.export_limit_refusal(format!(
                "a request carries at most {MAX_REQUEST_BYTES} bytes"
            ))
        })
    }

    fn bounded_export<T: Serialize>(&self, result: &XlsxExportResult<T>) -> Result<String> {
        let json = export_result_json(result)?;
        if json.len() <= MAX_RESPONSE_BYTES {
            return Ok(json);
        }
        Ok(self.export_limit_refusal(format!("the result exceeds {MAX_RESPONSE_BYTES} bytes")))
    }

    fn export_limit_refusal(&self, message: String) -> String {
        let refused: XlsxExportResult<()> = Err(XlsxExportRefusal {
            version: self.version(),
            failure: XlsxExportFailure {
                code: XlsxExportFailureCode::LimitExceeded,
                target: None,
                message,
            },
        });
        export_result_json(&refused).expect("a refusal encodes")
    }
}

/// Exports `.xlsx` bytes as read, without recalculating: stored formula results, no clock,
/// `anchorScope: "snapshot"`. Options a live export would refuse are
/// [`Error::InvalidRequest`] here.
pub fn export_xlsx_structured(
    bytes: &[u8],
    options: &XlsxExportOptions,
) -> Result<XlsxStructuredContent> {
    let workbook = Workbook::open_for_read(bytes)?;
    walk::export(
        &workbook.export_source(),
        options,
        XlsxAnchorScope::Snapshot,
    )
    .map_err(|failure| Error::InvalidRequest(failure.message))
}

/// [`export_xlsx_structured`] rendered as Markdown.
pub fn export_xlsx_markdown(
    bytes: &[u8],
    options: &XlsxExportOptions,
    markdown_options: &XlsxMarkdownOptions,
) -> Result<XlsxMarkdownContent> {
    markdown::limits(markdown_options).map_err(|failure| Error::InvalidRequest(failure.message))?;
    render_xlsx_markdown(&export_xlsx_structured(bytes, options)?, markdown_options)
}

/// [`export_xlsx_structured`] over JSON.
pub fn export_xlsx_structured_json(bytes: &[u8], options: &str) -> Result<String> {
    bounded_content(&export_xlsx_structured(
        bytes,
        &decode_bounded(options, MAX_REQUEST_BYTES)?,
    )?)
}

/// [`export_xlsx_markdown`] over JSON.
pub fn export_xlsx_markdown_json(
    bytes: &[u8],
    options: &str,
    markdown_options: &str,
) -> Result<String> {
    bounded_content(&export_xlsx_markdown(
        bytes,
        &decode_bounded(options, MAX_REQUEST_BYTES)?,
        &decode_bounded(markdown_options, MAX_REQUEST_BYTES)?,
    )?)
}

/// [`render_xlsx_markdown`] over JSON.
pub fn render_xlsx_markdown_json(content: &str, options: &str) -> Result<String> {
    bounded_content(&render_xlsx_markdown(
        &decode_bounded(content, MAX_RENDER_REQUEST_BYTES)?,
        &decode_bounded(options, MAX_REQUEST_BYTES)?,
    )?)
}

fn decode<T: serde::de::DeserializeOwned>(request: &str) -> Result<T> {
    serde_json::from_str(request).map_err(|error| Error::InvalidRequest(error.to_string()))
}

fn decode_bounded<T: serde::de::DeserializeOwned>(request: &str, max: usize) -> Result<T> {
    if request.len() > max {
        return Err(Error::InvalidRequest(format!(
            "a request carries at most {max} bytes"
        )));
    }
    decode(request)
}

fn bounded_content<T: Serialize>(content: &T) -> Result<String> {
    let json = serde_json::to_string(content)
        .map_err(|error| Error::InvalidOperation(error.to_string()))?;
    if json.len() > MAX_RESPONSE_BYTES {
        return Err(Error::InvalidOperation(format!(
            "the result exceeds {MAX_RESPONSE_BYTES} bytes"
        )));
    }
    Ok(json)
}
