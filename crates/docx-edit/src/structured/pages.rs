//! Page fragments of a structured export: the physical page, region and story occurrence that
//! shows each exported block and inline, read from the layout the export was captured with.
//!
//! Layout placement (line windows, row windows, bands and note areas) comes from
//! [`docx_layout::placement`]; the lowering map recorded with the laid-out blocks turns display
//! positions back into story units, and the batch projection turns those into the anchors the
//! export already carries. No other offset space is exposed.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::rc::Rc;

use docx_layout::display_list::DisplayList;
use docx_layout::footnotes::{NoteContent, NoteKind as LayoutNoteKind};
use docx_layout::header_footer::{HeaderFooterKind, HeaderFooterPayload, HeaderFooterType};
use docx_layout::hit::{RangeRect, RangeRectIndex, note_primitives};
use docx_layout::placement::{
    PlacedItem, PlacedParagraph, PlacedRegion, PlacedTable, PlacementInput, PlacementIssueKind,
    line_window_slices, place_layout,
};
use docx_layout::types::{Layout, LayoutBlock, MeasuredBlock, Page};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use yrs::{Any, Map, Out, ReadTxn, Transact};

use super::*;
use crate::bridge::LoweringMap;
use crate::ops::ChunkKind;
use crate::target::{EditTextView, StoryView, Views};
use crate::{DEL, INS, KIND_KEY, PPR_CHANGE, PPR_DEL, PPR_INS, TextPosition, TextRange, story_ref};

/// Fragments a page map returns when the options name no limit.
pub const DEFAULT_MAX_FRAGMENTS: u32 = 100_000;
/// The largest fragment limit a page map accepts.
pub const MAX_FRAGMENTS_LIMIT: u32 = 1_000_000;
/// Diagnostics one page map returns.
const MAX_PAGE_DIAGNOSTICS: usize = 1_000;
/// Diagnostics of one code a page map returns before it summarizes the rest.
const MAX_PAGE_DIAGNOSTICS_PER_CODE: usize = 100;

/// A paged export: structured content with the page map captured with it.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocxPagedStructuredContent<L> {
    pub structured: DocxStructuredContent,
    pub layout: L,
}

/// What a paged export reads. The export options are [`ExportOptions`]'; `include_geometry`
/// defaults to false, `max_fragments` to [`DEFAULT_MAX_FRAGMENTS`] and `max_layout_bytes` to
/// [`DEFAULT_MAX_BYTES`], bounding the page map independently of the content.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PageExportOptions {
    pub revision_view: RevisionView,
    #[serde(default)]
    pub stories: Option<Vec<StorySelection>>,
    #[serde(default)]
    pub include_formatting: Option<bool>,
    #[serde(default)]
    pub max_blocks: Option<u32>,
    #[serde(default)]
    pub max_bytes: Option<u32>,
    #[serde(default)]
    pub include_geometry: Option<bool>,
    #[serde(default)]
    pub expect_layout_version: Option<String>,
    #[serde(default)]
    pub max_fragments: Option<u32>,
    #[serde(default)]
    pub max_layout_bytes: Option<u32>,
}

impl PageExportOptions {
    pub fn new(revision_view: RevisionView) -> Self {
        Self {
            revision_view,
            stories: None,
            include_formatting: None,
            max_blocks: None,
            max_bytes: None,
            include_geometry: None,
            expect_layout_version: None,
            max_fragments: None,
            max_layout_bytes: None,
        }
    }

    /// The options of the structured export the page map attaches to.
    pub fn export_options(&self) -> ExportOptions {
        ExportOptions {
            revision_view: self.revision_view,
            stories: self.stories.clone(),
            include_formatting: self.include_formatting,
            max_blocks: self.max_blocks,
            max_bytes: self.max_bytes,
        }
    }
}

/// Markdown options for a paged export; `page_markers` defaults to false.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PageMarkdownOptions {
    #[serde(default)]
    pub max_bytes: Option<u32>,
    #[serde(default)]
    pub page_markers: Option<bool>,
}

/// The revision view pages are laid out in.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum LayoutRevisionView {
    Markup,
}

/// Which layout a session page map was read from.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutProvenance {
    pub engine_version: String,
    pub font_set_fingerprint: String,
    pub options_fingerprint: String,
    pub layout_epoch: String,
}

/// Which layout a snapshot page map was read from; every value is deterministic.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SnapshotLayoutProvenance {
    pub engine_version: String,
    pub font_set_fingerprint: String,
    pub options_fingerprint: String,
}

/// The page map of a session export, valid for `document_version` and `layout_version`.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocxLayoutMap {
    pub document_version: crate::batch::DocumentVersion,
    pub layout_version: String,
    pub export_fingerprint: String,
    pub revision_view: RevisionView,
    pub layout_revision_view: LayoutRevisionView,
    pub provenance: LayoutProvenance,
    pub pages: Vec<ExportPage>,
    pub occurrences: Vec<StoryOccurrence>,
    pub fragments: Vec<PageFragment>,
    pub diagnostics: Vec<PageDiagnostic>,
    pub truncated: bool,
}

/// The page map of a bytes export: the session map with its live tokens replaced by a
/// deterministic fingerprint.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocxSnapshotLayoutMap {
    pub snapshot_fingerprint: String,
    pub export_fingerprint: String,
    pub revision_view: RevisionView,
    pub layout_revision_view: LayoutRevisionView,
    pub provenance: SnapshotLayoutProvenance,
    pub pages: Vec<ExportPage>,
    pub occurrences: Vec<StoryOccurrence>,
    pub fragments: Vec<PageFragment>,
    pub diagnostics: Vec<PageDiagnostic>,
    pub truncated: bool,
}

impl DocxSnapshotLayoutMap {
    /// Drops the session tokens of `map`.
    pub fn from_session(map: DocxLayoutMap) -> Self {
        let provenance = SnapshotLayoutProvenance {
            engine_version: map.provenance.engine_version,
            font_set_fingerprint: map.provenance.font_set_fingerprint,
            options_fingerprint: map.provenance.options_fingerprint,
        };
        let snapshot_fingerprint = sha256_hex(
            [
                map.export_fingerprint.as_str(),
                provenance.engine_version.as_str(),
                provenance.font_set_fingerprint.as_str(),
                provenance.options_fingerprint.as_str(),
            ]
            .join("\n")
            .as_bytes(),
        );
        Self {
            snapshot_fingerprint,
            export_fingerprint: map.export_fingerprint,
            revision_view: map.revision_view,
            layout_revision_view: map.layout_revision_view,
            provenance,
            pages: map.pages,
            occurrences: map.occurrences,
            fragments: map.fragments,
            diagnostics: map.diagnostics,
            truncated: map.truncated,
        }
    }
}

/// One physical page. `page_index` counts every sheet, blank parity fillers included; the
/// displayed number follows the section's PAGE numbering and may restart or repeat.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportPage {
    pub page_index: u32,
    pub section_index: u32,
    pub section_page_index: u32,
    pub displayed_number: u64,
    pub displayed_label: String,
    pub numbering_format: String,
    pub numbering_status: NumberingStatus,
    pub parity_filler: bool,
    pub size: PageSize,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum NumberingStatus {
    Resolved,
    Fallback,
}

/// A page size in CSS pixels (96 per inch).
#[derive(Clone, Copy, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSize {
    pub width: f64,
    pub height: f64,
}

/// One appearance of an exported story on a page. A header or footer part appears once per page
/// that shows it; its content is exported once.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryOccurrence {
    pub id: String,
    pub page_index: u32,
    pub story: String,
    pub part: Option<String>,
    pub section_index: Option<u32>,
    pub region: OccurrenceRegion,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum OccurrenceRegion {
    Body,
    Header {
        variant: HeaderFooterVariant,
    },
    Footer {
        variant: HeaderFooterVariant,
    },
    Footnote {
        note_id: String,
        placement: NotePlacement,
    },
    Endnote {
        note_id: String,
        placement: NotePlacement,
    },
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum NotePlacement {
    PageBottom,
    BeneathText,
    SectionEnd,
    DocumentEnd,
}

/// The part of one exported node shown in one story occurrence.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageFragment {
    pub id: String,
    pub page_index: u32,
    pub occurrence_id: String,
    pub node_id: String,
    pub block_id: String,
    pub anchor: Anchor,
    pub slice: FragmentSlice,
    pub continued_from_previous: bool,
    pub continued_on_next: bool,
    pub repeated_table_header: bool,
    pub geometry: Option<FragmentGeometry>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum FragmentSlice {
    /// Text of the node's view, as a paragraph-local range.
    Text {
        range: TextRange,
    },
    /// An atom (field, control, image, note mark, break), whose display content has no offsets
    /// of its own; `partial` when only some of it is on this page.
    Atom {
        range: Option<TextRange>,
        coverage: AtomCoverage,
    },
    Block,
    Table {
        rows: Vec<FragmentRow>,
    },
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum AtomCoverage {
    Whole,
    Partial,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FragmentRow {
    pub row_index: u32,
    pub continued_from_previous: bool,
    pub continued_on_next: bool,
    pub repeated_header: bool,
}

/// Where a fragment is painted: separate rectangles in unzoomed CSS pixels from the physical
/// page's top-left corner.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FragmentGeometry {
    pub unit: GeometryUnit,
    pub origin: GeometryOrigin,
    pub rects: Vec<GeometryRect>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum GeometryUnit {
    CssPx,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum GeometryOrigin {
    PageTopLeft,
}

#[derive(Clone, Copy, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GeometryRect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageDiagnostic {
    pub code: PageDiagnosticCode,
    pub node_id: Option<String>,
    pub page_index: Option<u32>,
    pub message: String,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum PageDiagnosticCode {
    UnmappedContent,
    AnchorOnly,
    NotLaidOut,
    UnsupportedNumbering,
    UnsupportedNoteLayout,
    GeometryUnavailable,
    Truncated,
}

/// Page options with their defaults applied and validated.
pub(crate) struct PageLimits {
    pub include_geometry: bool,
    pub max_fragments: usize,
    pub max_bytes: usize,
}

impl PageLimits {
    pub fn new(options: &PageExportOptions) -> Result<Self, ExportFailure> {
        let max_fragments = options.max_fragments.unwrap_or(DEFAULT_MAX_FRAGMENTS);
        if max_fragments == 0 {
            return Err(invalid_options("maxFragments must be at least 1"));
        }
        if max_fragments > MAX_FRAGMENTS_LIMIT {
            return Err(limit_exceeded(format!(
                "maxFragments may be at most {MAX_FRAGMENTS_LIMIT}"
            )));
        }
        let max_bytes = byte_limit(options.max_layout_bytes).map_err(|failure| ExportFailure {
            message: failure.message.replace("maxBytes", "maxLayoutBytes"),
            ..failure
        })?;
        Ok(Self {
            include_geometry: options.include_geometry.unwrap_or(false),
            max_fragments: max_fragments as usize,
            max_bytes,
        })
    }
}

/// A refusal with no target.
pub(crate) fn failure(code: ExportFailureCode, message: impl Into<String>) -> ExportFailure {
    ExportFailure {
        code,
        target: None,
        message: message.into(),
    }
}

fn hex(digest: &[u8]) -> String {
    let mut hex = String::with_capacity(digest.len() * 2);
    for byte in digest {
        let _ = std::fmt::Write::write_fmt(&mut hex, format_args!("{byte:02x}"));
    }
    hex
}

pub(crate) fn sha256_hex(bytes: &[u8]) -> String {
    hex(&Sha256::digest(bytes))
}

/// Hashes the compact JSON of `value` without building it.
fn json_sha256<T: Serialize + ?Sized>(value: &T) -> String {
    struct Hasher(Sha256);
    impl std::io::Write for Hasher {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.0.update(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }
    let mut hasher = Hasher(Sha256::new());
    let _ = serde_json::to_writer(&mut hasher, value);
    hex(&hasher.0.finalize())
}

/// The fingerprint a page map records for the content it attaches to.
pub fn export_fingerprint(content: &DocxStructuredContent) -> String {
    json_sha256(content)
}

/// Why the section or settings metadata a region layout was requested with does not describe
/// `doc`, if it does not. Every section record layout reads is compared, index for index: the
/// sections the body's paragraph marks end, held by the editing stream, and, unless an editor
/// owns them, the final section and settings of the source package. A repeat of the final
/// record, which the request builder appends, is the only record allowed past them. Sections
/// compare as layout reads them, with header and footer references inherited.
pub(crate) fn metadata_mismatch(
    doc: &EditingDoc,
    request: &serde_json::Value,
    editor_owned: bool,
) -> Option<String> {
    let mut sections: Vec<docx_parse::SectionProperties> = Vec::new();
    {
        let txn = doc.yrs_doc().transact();
        if let Ok(text) = story_ref(&txn, "body") {
            for chunk in doc.chunk_snapshot("body", &text, &txn).iter() {
                let ChunkKind::Pilcrow(map) = &chunk.kind else {
                    continue;
                };
                let value = |key: &str| match map.get(&txn, key) {
                    Some(Out::Any(value)) if !matches!(value, Any::Null | Any::Undefined) => {
                        Some(value)
                    }
                    _ => None,
                };
                let (section, break_type) = (value("sectPr"), value("sectionBreakType"));
                if section.is_none() && break_type.is_none() {
                    continue;
                }
                let mut properties: docx_parse::SectionProperties = section
                    .and_then(|value| serde_json::to_value(value).ok())
                    .and_then(|value| serde_json::from_value(value).ok())
                    .unwrap_or_default();
                if properties.section_start.is_none()
                    && let Some(Any::String(start)) = break_type
                {
                    properties.section_start = Some(start.to_string());
                }
                sections.push(properties);
            }
        }
    }
    let requested: Vec<docx_parse::SectionProperties> = request
        .pointer("/regions/sections")
        .and_then(serde_json::Value::as_array)
        .map(|sections| {
            sections
                .iter()
                .map(|section| {
                    section
                        .get("properties")
                        .cloned()
                        .and_then(|value| serde_json::from_value(value).ok())
                        .unwrap_or_default()
                })
                .collect()
        })
        .unwrap_or_default();
    let authored = |section: &docx_parse::SectionProperties| {
        serde_json::to_value(section)
            .ok()
            .and_then(|value| {
                serde_json::from_value::<docx_layout::regions::AuthoredSectionProperties>(value)
                    .ok()
            })
            .and_then(|properties| serde_json::to_value(properties).ok())
    };
    let inner = sections.len();
    let repeated_final = requested.len() == inner + 2
        && authored(&requested[inner]) == authored(&requested[inner + 1]);
    if requested.len() != inner + 1 && !repeated_final {
        return Some(format!(
            "The layout's section metadata has {} records for the document's {} sections; lay it out with its current sections.",
            requested.len(),
            inner + 1
        ));
    }
    let source = doc.source_metadata();
    let read = source.as_deref().map(|source| source.read());
    sections.push(
        read.and_then(|read| read.final_section.clone())
            .and_then(|value| serde_json::from_value(value).ok())
            .unwrap_or_default(),
    );
    docx_parse::apply_section_inheritance(&mut sections);
    let mut laid_out = requested[..=inner].to_vec();
    docx_parse::apply_section_inheritance(&mut laid_out);
    let compared = if editor_owned { inner } else { inner + 1 };
    if let Some(index) =
        (0..compared).find(|index| authored(&sections[*index]) != authored(&laid_out[*index]))
    {
        return Some(if index == inner {
            "The layout's final section differs from the document's; lay it out with its current sections.".to_owned()
        } else {
            format!(
                "The layout's metadata for section {index} differs from the section the document ends there; lay it out with its current sections."
            )
        });
    }
    let settings = |value: Option<&serde_json::Value>| {
        serde_json::to_value(
            value
                .cloned()
                .and_then(|value| {
                    serde_json::from_value::<docx_layout::regions::AuthoredRegionSettings>(value)
                        .ok()
                })
                .unwrap_or_default(),
        )
        .ok()
    };
    if !editor_owned
        && settings(request.pointer("/regions/settings"))
            != settings(read.and_then(|read| read.settings.as_ref()))
    {
        return Some(
            "The layout's note and header settings differ from the document's; lay it out with its current settings.".to_owned(),
        );
    }
    None
}

/// Whether any story that contributes to layout carries a pending revision, which would make
/// the accepted and original views reflow differently from the markup the pages show. Returns
/// the story holding the first one.
pub(crate) fn revised_story(doc: &EditingDoc, stories: &BTreeSet<String>) -> Option<String> {
    if let Some(source) = doc.source_metadata()
        && let Some(story) = source
            .run_revision_stories()
            .find(|story| stories.contains(*story))
    {
        return Some(story.to_owned());
    }
    let txn = doc.yrs_doc().transact();
    for story_id in stories {
        let Ok(story) = story_ref(&txn, story_id) else {
            continue;
        };
        let chunks = doc.chunk_snapshot(story_id, &story, &txn);
        for chunk in chunks.iter() {
            if chunk.attr_active(INS) || chunk.attr_active(DEL) {
                return Some(story_id.clone());
            }
            match &chunk.kind {
                ChunkKind::Pilcrow(map) => {
                    let revised = [PPR_INS, PPR_DEL, PPR_CHANGE]
                        .iter()
                        .any(|key| present_out(map.get(&txn, key)))
                        || matches!(
                            map.get(&txn, "_originalRunBoundaries"),
                            Some(Out::Any(Any::Array(runs))) if runs.iter().any(|run| matches!(
                                run,
                                Any::Map(run) if matches!(
                                    run.get("propertyChanges"),
                                    Some(Any::Array(changes)) if !changes.is_empty()
                                )
                            ))
                        );
                    if revised {
                        return Some(story_id.clone());
                    }
                }
                ChunkKind::Embed(Some(map))
                    if crate::map_string(map, &txn, KIND_KEY).as_deref() == Some("table")
                        && table_revised(map, &txn) =>
                {
                    return Some(story_id.clone());
                }
                _ => {}
            }
        }
    }
    None
}

fn present_out(value: Option<Out>) -> bool {
    match value {
        None | Some(Out::Any(Any::Null | Any::Undefined)) => false,
        Some(Out::Any(Any::Array(items))) => !items.is_empty(),
        Some(_) => true,
    }
}

/// Whether a table embed's rows, cells or table properties carry tracked changes.
fn table_revised<T: ReadTxn>(map: &yrs::MapRef, txn: &T) -> bool {
    let present = |value: Option<&Any>| match value {
        None | Some(Any::Null | Any::Undefined) => false,
        Some(Any::Array(items)) => !items.is_empty(),
        Some(_) => true,
    };
    let as_map = |value: Option<&Any>| match value {
        Some(Any::Map(map)) => Some(map.clone()),
        _ => None,
    };
    let payload = |key: &str| match map.get(txn, key) {
        Some(Out::Any(value)) => Some(value),
        _ => None,
    };
    if let Some(Any::Map(properties)) = payload("tblPr")
        && present(properties.get("tblPrChange"))
    {
        return true;
    }
    let Some(Any::Array(rows)) = payload("rows") else {
        return false;
    };
    rows.iter().any(|row| {
        let Any::Map(row) = row else {
            return false;
        };
        let row_properties = as_map(row.get("trPr"));
        if row_properties.as_ref().is_some_and(|properties| {
            ["trIns", "trDel", "trPrChange"]
                .iter()
                .any(|key| present(properties.get(*key)))
        }) {
            return true;
        }
        let Some(Any::Array(cells)) = row.get("cells") else {
            return false;
        };
        cells.iter().any(|cell| {
            let Any::Map(cell) = cell else {
                return false;
            };
            as_map(cell.get("tcPr")).is_some_and(|properties| {
                ["cellMarker", "tcPrChange"]
                    .iter()
                    .any(|key| present(properties.get(*key)))
            })
        })
    })
}

/// The retained layout a page map is read from, with the lowering map of every laid-out root
/// story and, for geometry, the display list it paints.
pub(crate) struct CapturedLayout<'a> {
    pub layout: &'a Layout,
    pub measured: &'a [MeasuredBlock],
    pub headers_footers: Option<&'a HeaderFooterPayload>,
    pub bands_composed: bool,
    pub notes: &'a [NoteContent],
    /// Lowering maps by root story.
    pub maps: &'a HashMap<String, Rc<LoweringMap>>,
    pub display: Option<&'a DisplayList>,
}

/// The identity a session page map records.
pub(crate) struct MapIdentity {
    pub document_version: crate::batch::DocumentVersion,
    pub layout_version: String,
    pub layout_epoch: String,
    pub font_set_fingerprint: String,
    pub options_fingerprint: String,
}

/// Maps the content of one capture onto its pages. Refuses a layout that shows a footnote on
/// another page than its reference, whatever the export selects.
pub(crate) fn build_layout_map(
    doc: &EditingDoc,
    content: &DocxStructuredContent,
    captured: &CapturedLayout<'_>,
    limits: &PageLimits,
    identity: MapIdentity,
) -> Result<DocxLayoutMap, ExportFailure> {
    let placements = place_layout(&PlacementInput {
        layout: captured.layout,
        measured: captured.measured,
        headers_footers: captured.headers_footers,
        bands_composed: captured.bands_composed,
        notes: captured.notes,
    });
    if let Some((id, reference_page, note_page)) =
        placements.issues.iter().find_map(|issue| match issue.kind {
            PlacementIssueKind::NoteReferenceElsewhere { id, reference_page } => {
                Some((id, reference_page, issue.page_index))
            }
            _ => None,
        })
    {
        return Err(failure(
            ExportFailureCode::Unsupported,
            format!(
                "Footnote {id} is referenced from the part of a table row on page index {reference_page}, but the layout places it on page index {note_page}; pages of this layout cannot be exported."
            ),
        ));
    }
    let index = NodeIndex::new(content);
    let txn = doc.yrs_doc().transact();
    let mut mapper = Mapper {
        index: &index,
        content,
        maps: captured.maps,
        views: Views::new(doc, &txn),
        paragraph_indexes: HashMap::new(),
        atom_units: HashMap::new(),
        aliases: HashMap::new(),
        export_views: match content.revision_view {
            RevisionView::Accepted => &[EditTextView::Accepted],
            RevisionView::Original => &[EditTextView::Original],
            RevisionView::Markup => &[EditTextView::Accepted, EditTextView::Original],
        },
        items: Vec::new(),
        fragment_count: 0,
        diagnostics: Diagnostics::default(),
        placed_nodes: HashSet::new(),
        wrapped: HashSet::new(),
        laid_out_roots: HashSet::new(),
    };
    let selected: HashSet<StorySelection> = content.included_stories.iter().copied().collect();
    let mut pages = Vec::with_capacity(placements.pages.len());
    for placed in &placements.pages {
        let page_index = placed.page_index as u32;
        let (record, unsupported) = export_page(placed.page_index, placed.page);
        if let Some(format) = unsupported {
            mapper.diagnostics.push(PageDiagnostic {
                code: PageDiagnosticCode::UnsupportedNumbering,
                node_id: None,
                page_index: Some(page_index),
                message: format!(
                    "The page number format {format:?} is not supported; the page is numbered in decimal."
                ),
            });
        }
        pages.push(record);
        for region in &placed.regions {
            let occurrence = match &region.region {
                PlacedRegion::Body { section_index } => {
                    if !selected.contains(&StorySelection::Body) {
                        continue;
                    }
                    mapper.laid_out_roots.insert("body".to_owned());
                    Occurrence {
                        record: StoryOccurrence {
                            id: format!("p{page_index}.body.s{section_index}"),
                            page_index,
                            story: "body".to_owned(),
                            part: index.part("body"),
                            section_index: Some(*section_index as u32),
                            region: OccurrenceRegion::Body,
                        },
                        root: "body".to_owned(),
                        flow: "body".to_owned(),
                    }
                }
                PlacedRegion::Band {
                    kind,
                    hf_type,
                    r_id,
                    section_index,
                } => {
                    let category = match kind {
                        HeaderFooterKind::Header => StorySelection::Headers,
                        HeaderFooterKind::Footer => StorySelection::Footers,
                    };
                    if !selected.contains(&category) {
                        continue;
                    }
                    let variant = match hf_type {
                        HeaderFooterType::Default => HeaderFooterVariant::Default,
                        HeaderFooterType::First => HeaderFooterVariant::First,
                        HeaderFooterType::Even => HeaderFooterVariant::Even,
                    };
                    let root = format!("hf:{r_id}");
                    let story = if index.stories.contains_key(&root) {
                        root.clone()
                    } else {
                        index
                            .band_story(*kind, *section_index as u32, variant)
                            .unwrap_or_else(|| root.clone())
                    };
                    mapper.laid_out_roots.insert(story.clone());
                    let kind_name = match kind {
                        HeaderFooterKind::Header => "header",
                        HeaderFooterKind::Footer => "footer",
                    };
                    let id = format!("p{page_index}.{kind_name}");
                    Occurrence {
                        record: StoryOccurrence {
                            id: id.clone(),
                            page_index,
                            part: index.part(&story),
                            story,
                            section_index: Some(*section_index as u32),
                            region: match kind {
                                HeaderFooterKind::Header => OccurrenceRegion::Header { variant },
                                HeaderFooterKind::Footer => OccurrenceRegion::Footer { variant },
                            },
                        },
                        root,
                        flow: id,
                    }
                }
                PlacedRegion::Note {
                    kind,
                    id,
                    placement,
                } => {
                    let (category, prefix, name) = match kind {
                        LayoutNoteKind::Footnote => (StorySelection::Footnotes, "fn", "footnote"),
                        LayoutNoteKind::Endnote => (StorySelection::Endnotes, "en", "endnote"),
                    };
                    if !selected.contains(&category) {
                        continue;
                    }
                    let story = format!("{prefix}:{id}");
                    mapper.laid_out_roots.insert(story.clone());
                    let placement = match *placement {
                        "beneathText" => NotePlacement::BeneathText,
                        "sectEnd" => NotePlacement::SectionEnd,
                        "docEnd" => NotePlacement::DocumentEnd,
                        _ => NotePlacement::PageBottom,
                    };
                    let occurrence_id = format!("p{page_index}.{name}.{id}");
                    Occurrence {
                        record: StoryOccurrence {
                            id: occurrence_id.clone(),
                            page_index,
                            part: index.part(&story),
                            story: story.clone(),
                            section_index: None,
                            region: match kind {
                                LayoutNoteKind::Footnote => OccurrenceRegion::Footnote {
                                    note_id: id.to_string(),
                                    placement,
                                },
                                LayoutNoteKind::Endnote => OccurrenceRegion::Endnote {
                                    note_id: id.to_string(),
                                    placement,
                                },
                            },
                        },
                        root: story,
                        flow: occurrence_id,
                    }
                }
            };
            mapper.region(&occurrence, &region.items);
        }
    }
    for issue in &placements.issues {
        let (code, message) = match &issue.kind {
            PlacementIssueKind::UnresolvedFragment => (
                PageDiagnosticCode::UnmappedContent,
                "A laid-out fragment names no measured block.".to_owned(),
            ),
            PlacementIssueKind::MissingNote { kind, id } => (
                PageDiagnosticCode::UnsupportedNoteLayout,
                format!(
                    "The {} {id} placed on this page has no laid-out content.",
                    match kind {
                        LayoutNoteKind::Footnote => "footnote",
                        LayoutNoteKind::Endnote => "endnote",
                    }
                ),
            ),
            PlacementIssueKind::UnmeasuredBand => (
                PageDiagnosticCode::NotLaidOut,
                "The page's header or footer was composed from content this layout did not measure."
                    .to_owned(),
            ),
            PlacementIssueKind::NoteOverflow { kind, id } => (
                PageDiagnosticCode::UnsupportedNoteLayout,
                format!(
                    "The {} {id} does not fit in the note area of this page, and the layout does not continue notes onto the next page, so where its content shows is not reported.",
                    match kind {
                        LayoutNoteKind::Footnote => "footnote",
                        LayoutNoteKind::Endnote => "endnote",
                    }
                ),
            ),
            PlacementIssueKind::NoteReferenceElsewhere { .. } => continue,
        };
        mapper.diagnostics.push(PageDiagnostic {
            code,
            node_id: None,
            page_index: Some(issue.page_index as u32),
            message,
        });
    }
    mapper.unplaced(&selected);
    let mut items = std::mem::take(&mut mapper.items);
    link_continuations(&mut items);
    if limits.include_geometry {
        match captured.display {
            Some(display) => add_geometry(&mut items, display),
            None => mapper.diagnostics.push(PageDiagnostic {
                code: PageDiagnosticCode::GeometryUnavailable,
                node_id: None,
                page_index: None,
                message: "Page geometry could not be computed for this layout.".to_owned(),
            }),
        }
    }
    let diagnostics = std::mem::take(&mut mapper.diagnostics).finish();
    drop(mapper);
    Ok(assemble(
        DocxLayoutMap {
            document_version: identity.document_version,
            layout_version: identity.layout_version,
            export_fingerprint: export_fingerprint(content),
            revision_view: content.revision_view,
            layout_revision_view: LayoutRevisionView::Markup,
            provenance: LayoutProvenance {
                engine_version: env!("CARGO_PKG_VERSION").to_owned(),
                font_set_fingerprint: identity.font_set_fingerprint,
                options_fingerprint: identity.options_fingerprint,
                layout_epoch: identity.layout_epoch,
            },
            pages: Vec::new(),
            occurrences: Vec::new(),
            fragments: Vec::new(),
            diagnostics: Vec::new(),
            truncated: false,
        },
        pages,
        items,
        diagnostics,
        limits,
    ))
}

/// A page's record, and the number format it could not write.
fn export_page(page_index: usize, page: &Page) -> (ExportPage, Option<String>) {
    let displayed_number = page
        .section_page_number
        .unwrap_or_else(|| u64::from(page.number.max(1)));
    let numbering_format = page
        .page_numbering
        .as_ref()
        .and_then(|numbering| numbering.get("format"))
        .and_then(|format| format.as_str())
        .unwrap_or("decimal")
        .to_owned();
    let resolved =
        docx_layout::regions::number_format_resolves(displayed_number as i64, &numbering_format);
    (
        ExportPage {
            page_index: page_index as u32,
            section_index: page
                .section_index
                .unwrap_or(page.region_section_index as u64) as u32,
            section_page_index: page.section_page_index.unwrap_or(0) as u32,
            displayed_number,
            displayed_label: page
                .page_label
                .clone()
                .unwrap_or_else(|| displayed_number.to_string()),
            numbering_status: if resolved {
                NumberingStatus::Resolved
            } else {
                NumberingStatus::Fallback
            },
            numbering_format: numbering_format.clone(),
            parity_filler: page.parity_filler == Some(true),
            size: PageSize {
                width: page.size.w,
                height: page.size.h,
            },
        },
        (!resolved).then_some(numbering_format),
    )
}

struct Occurrence {
    record: StoryOccurrence,
    /// The laid-out root story whose lowering map the region's positions belong to.
    root: String,
    /// Fragments in one flow continue one another; each band and note is a flow of its own.
    flow: String,
}

/// An occurrence or fragment of the map, in output order.
enum MapItem {
    Occurrence(StoryOccurrence),
    Fragment(Box<DraftFragment>),
}

struct DraftFragment {
    fragment: PageFragment,
    flow: String,
    /// Whether continuation flags come from the layout rather than from the fragments' order.
    placed_flags: bool,
    geometry: Geometry,
}

/// Where to read a fragment's geometry from.
enum Geometry {
    None,
    /// The page fragment's own box.
    Frame([f64; 4]),
    /// Display ranges of a region of the page.
    Ranges(RegionKey, Vec<(i64, i64)>),
}

/// A page region's primitives in the display list: the body, a band by relationship id, or a
/// note by endnote flag and id.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
enum RegionKey {
    Body,
    Header(String),
    Footer(String),
    Note(bool, i64),
}

#[derive(Default)]
struct Diagnostics {
    items: Vec<PageDiagnostic>,
    counts: BTreeMap<PageDiagnosticCode, usize>,
}

impl Diagnostics {
    fn push(&mut self, diagnostic: PageDiagnostic) {
        let count = self.counts.entry(diagnostic.code).or_default();
        *count += 1;
        if *count <= MAX_PAGE_DIAGNOSTICS_PER_CODE {
            self.items.push(diagnostic);
        }
    }

    fn finish(self) -> Vec<PageDiagnostic> {
        let mut items = self.items;
        for (code, count) in self.counts {
            if count > MAX_PAGE_DIAGNOSTICS_PER_CODE {
                items.push(PageDiagnostic {
                    code,
                    node_id: None,
                    page_index: None,
                    message: format!(
                        "{} further diagnostics with this code were omitted.",
                        count - MAX_PAGE_DIAGNOSTICS_PER_CODE
                    ),
                });
            }
        }
        items.truncate(MAX_PAGE_DIAGNOSTICS);
        items
    }
}

/// An export control a block sits inside, which gets a block fragment wherever its content is.
#[derive(Clone, Debug)]
struct ControlRef {
    id: String,
    anchor: Anchor,
}

struct TextEntry {
    id: String,
    start: u32,
    end: u32,
    anchor: Anchor,
}

struct ParagraphNode {
    id: String,
    anchor: Anchor,
    /// The exported story the paragraph is part of.
    root: String,
    controls: Rc<[ControlRef]>,
    /// Text and tab inlines per view (accepted, original), by start.
    texts: [Vec<TextEntry>; 2],
    /// Atom inlines by view and offset.
    atoms: HashMap<(u8, u32), (String, TextRange)>,
    /// Note marks by note id.
    notes: HashMap<String, String>,
}

struct TableNode {
    id: String,
    anchor: Anchor,
    rows: usize,
    controls: Rc<[ControlRef]>,
}

/// The exported nodes the page map can attach to, by the anchors they carry.
struct NodeIndex {
    paragraphs: HashMap<(String, String), ParagraphNode>,
    tables: HashMap<(String, u32), TableNode>,
    /// The part of every exported story.
    stories: HashMap<String, Option<String>>,
    band_uses: HashMap<(u8, u32, u8), String>,
}

fn view_slot(view: EditTextView) -> u8 {
    match view {
        EditTextView::Accepted => 0,
        EditTextView::Original => 1,
    }
}

impl NodeIndex {
    fn new(content: &DocxStructuredContent) -> Self {
        let mut index = Self {
            paragraphs: HashMap::new(),
            tables: HashMap::new(),
            stories: HashMap::new(),
            band_uses: HashMap::new(),
        };
        for story in &content.stories {
            index
                .stories
                .insert(story.story.clone(), story.part.clone());
            let kind = match story.kind {
                StoryKind::Header => Some(0),
                StoryKind::Footer => Some(1),
                _ => None,
            };
            if let Some(kind) = kind {
                for used in &story.uses {
                    index.band_uses.insert(
                        (kind, used.section_index, used.variant as u8),
                        story.story.clone(),
                    );
                }
            }
            index.blocks(&story.blocks, &story.story, &Rc::from(Vec::new()));
        }
        index
    }

    fn part(&self, story: &str) -> Option<String> {
        self.stories.get(story).cloned().flatten()
    }

    fn band_story(
        &self,
        kind: HeaderFooterKind,
        section_index: u32,
        variant: HeaderFooterVariant,
    ) -> Option<String> {
        let kind = match kind {
            HeaderFooterKind::Header => 0,
            HeaderFooterKind::Footer => 1,
        };
        self.band_uses
            .get(&(kind, section_index, variant as u8))
            .cloned()
    }

    fn blocks(&mut self, blocks: &[Block], root: &str, controls: &Rc<[ControlRef]>) {
        for block in blocks {
            match &block.content {
                BlockKind::Paragraph { paragraph }
                | BlockKind::Heading { paragraph, .. }
                | BlockKind::ListItem { paragraph, .. } => {
                    let Anchor::Paragraph { story, para_id } = &block.anchor else {
                        continue;
                    };
                    if para_id.is_empty() {
                        continue;
                    }
                    let mut node = ParagraphNode {
                        id: block.id.clone(),
                        anchor: block.anchor.clone(),
                        root: root.to_owned(),
                        controls: Rc::clone(controls),
                        texts: [Vec::new(), Vec::new()],
                        atoms: HashMap::new(),
                        notes: HashMap::new(),
                    };
                    for inline in &paragraph.inlines {
                        let Anchor::Range(range) = &inline.anchor else {
                            continue;
                        };
                        if range.start.para_id != range.end.para_id {
                            continue;
                        }
                        let (start, end) = (range.start.offset, range.end.offset);
                        let slot = view_slot(range.view);
                        match &inline.content {
                            InlineKind::Text { .. } | InlineKind::Tab if end > start => {
                                node.texts[slot as usize].push(TextEntry {
                                    id: inline.id.clone(),
                                    start,
                                    end,
                                    anchor: inline.anchor.clone(),
                                });
                            }
                            InlineKind::Text { .. } | InlineKind::Tab => {}
                            content if end == start + 1 => {
                                if let InlineKind::NoteReference { note_id, .. } = content {
                                    node.notes.insert(note_id.clone(), inline.id.clone());
                                }
                                node.atoms
                                    .insert((slot, start), (inline.id.clone(), range.clone()));
                            }
                            _ => {}
                        }
                    }
                    for texts in &mut node.texts {
                        texts.sort_by_key(|entry| entry.start);
                    }
                    self.paragraphs
                        .insert((story.clone(), para_id.clone()), node);
                }
                BlockKind::Table { table } => {
                    if let Anchor::Table { story, table_index } = &block.anchor {
                        self.tables.insert(
                            (story.clone(), *table_index),
                            TableNode {
                                id: block.id.clone(),
                                anchor: block.anchor.clone(),
                                rows: table.rows.len(),
                                controls: Rc::clone(controls),
                            },
                        );
                    }
                    for row in &table.rows {
                        for cell in &row.cells {
                            self.blocks(&cell.blocks, root, controls);
                        }
                    }
                }
                BlockKind::ContentControl { blocks, .. } => {
                    let mut chain = controls.to_vec();
                    chain.push(ControlRef {
                        id: block.id.clone(),
                        anchor: block.anchor.clone(),
                    });
                    self.blocks(blocks, root, &Rc::from(chain));
                }
                _ => {}
            }
        }
    }
}

/// A laid-out paragraph's source: its story and paragraph id, in the exported story's names.
struct Source {
    story: String,
    para_id: String,
}

/// The stories, paragraphs and tables of a header or footer the export merged into another with
/// the same content, in that exported story's names, cell and control stories included.
#[derive(Default)]
struct AliasNames {
    stories: HashMap<String, String>,
    paragraphs: HashMap<(String, String), String>,
    tables: HashMap<(String, u32), u32>,
}

impl AliasNames {
    /// Pairs the lowerings of two stories with the same content, which record the same stories,
    /// paragraphs and tables in the same order. Empty when they do not.
    fn new(alias: &LoweringMap, exported: &LoweringMap) -> Self {
        let mut names = Self::default();
        if alias.stories.len() != exported.stories.len()
            || alias.paragraphs.len() != exported.paragraphs.len()
            || alias.tables.len() != exported.tables.len()
        {
            return names;
        }
        names.stories = alias
            .stories
            .iter()
            .cloned()
            .zip(exported.stories.iter().cloned())
            .collect();
        for ((story, para_id), (exported_story, exported_para_id)) in
            alias.paragraphs.iter().zip(&exported.paragraphs)
        {
            if story != exported_story {
                return Self::default();
            }
            names.paragraphs.insert(
                (alias.stories[*story as usize].clone(), para_id.clone()),
                exported_para_id.clone(),
            );
        }
        for (&(_, story, ordinal), &(_, exported_story, exported_ordinal)) in
            alias.tables.iter().zip(&exported.tables)
        {
            if story != exported_story {
                return Self::default();
            }
            names.tables.insert(
                (alias.stories[story as usize].clone(), ordinal),
                exported_ordinal,
            );
        }
        names
    }
}

struct Mapper<'a, 't, T: ReadTxn> {
    index: &'a NodeIndex,
    content: &'a DocxStructuredContent,
    maps: &'a HashMap<String, Rc<LoweringMap>>,
    views: Views<'t, T>,
    /// Paragraph positions by id, per story and view.
    paragraph_indexes: HashMap<(String, u8), Option<(Rc<StoryView>, Rc<HashMap<String, usize>>)>>,
    /// Display units of every atom, per root, source paragraph and story unit.
    atom_units: HashMap<String, Rc<HashMap<(u32, u32), u64>>>,
    /// The names of alias headers and footers in the exported stories'.
    aliases: HashMap<(String, String), Rc<AliasNames>>,
    export_views: &'static [EditTextView],
    items: Vec<MapItem>,
    fragment_count: usize,
    diagnostics: Diagnostics,
    /// Exported paragraph ids with a block fragment.
    placed_nodes: HashSet<String>,
    /// Controls given a block fragment, per occurrence.
    wrapped: HashSet<(String, String)>,
    /// Exported stories laid out on some page.
    laid_out_roots: HashSet<String>,
}

impl<'a, 't, T: ReadTxn> Mapper<'a, 't, T> {
    fn region(&mut self, occurrence: &Occurrence, items: &[PlacedItem<'_>]) {
        self.items
            .push(MapItem::Occurrence(occurrence.record.clone()));
        let Some(map) = self.maps.get(&occurrence.root).cloned() else {
            if !items.is_empty() {
                self.unmapped(occurrence, "Content of this region has no lowering map.");
            }
            return;
        };
        for item in items {
            match item {
                PlacedItem::Paragraph(paragraph) => self.paragraph(occurrence, &map, paragraph),
                PlacedItem::Table(table) => self.table(occurrence, &map, table),
                PlacedItem::Block(block) => self.drawing(occurrence, &map, block),
            }
        }
    }

    fn unmapped(&mut self, occurrence: &Occurrence, message: &str) {
        if self.content.truncated {
            return;
        }
        self.diagnostics.push(PageDiagnostic {
            code: PageDiagnosticCode::UnmappedContent,
            node_id: None,
            page_index: Some(occurrence.record.page_index),
            message: message.to_owned(),
        });
    }

    fn fragment(
        &mut self,
        occurrence: &Occurrence,
        node_id: String,
        block_id: String,
        anchor: Anchor,
        slice: FragmentSlice,
        flags: Option<(bool, bool)>,
        repeated: bool,
        geometry: Geometry,
    ) {
        let (from, next) = flags.unwrap_or((false, false));
        let id = format!("f{}", self.fragment_count);
        self.fragment_count += 1;
        self.items.push(MapItem::Fragment(Box::new(DraftFragment {
            fragment: PageFragment {
                id,
                page_index: occurrence.record.page_index,
                occurrence_id: occurrence.record.id.clone(),
                node_id,
                block_id,
                anchor,
                slice,
                continued_from_previous: from,
                continued_on_next: next,
                repeated_table_header: repeated,
                geometry: None,
            },
            flow: occurrence.flow.clone(),
            placed_flags: flags.is_some(),
            geometry,
        })));
    }

    /// Block fragments for the controls around a node, once per occurrence.
    fn wrap(&mut self, occurrence: &Occurrence, controls: &Rc<[ControlRef]>, repeated: bool) {
        for control in controls.iter() {
            if self
                .wrapped
                .insert((occurrence.record.id.clone(), control.id.clone()))
            {
                self.fragment(
                    occurrence,
                    control.id.clone(),
                    control.id.clone(),
                    control.anchor.clone(),
                    FragmentSlice::Block,
                    None,
                    repeated,
                    Geometry::None,
                );
            }
        }
    }

    fn region_key(occurrence: &Occurrence) -> RegionKey {
        match &occurrence.record.region {
            OccurrenceRegion::Body => RegionKey::Body,
            OccurrenceRegion::Header { .. } => {
                RegionKey::Header(occurrence.root.trim_start_matches("hf:").to_owned())
            }
            OccurrenceRegion::Footer { .. } => {
                RegionKey::Footer(occurrence.root.trim_start_matches("hf:").to_owned())
            }
            OccurrenceRegion::Footnote { note_id, .. } => {
                RegionKey::Note(false, note_id.parse().unwrap_or_default())
            }
            OccurrenceRegion::Endnote { note_id, .. } => {
                RegionKey::Note(true, note_id.parse().unwrap_or_default())
            }
        }
    }

    /// The source of the paragraph block starting at `pm`, translated into the exported
    /// story's names when the region shows an alias of it.
    fn source(&mut self, occurrence: &Occurrence, map: &LoweringMap, pm: u64) -> Option<Source> {
        let paragraph = map.paragraph_at(pm)?;
        self.source_of(occurrence, map, paragraph)
    }

    fn source_of(
        &mut self,
        occurrence: &Occurrence,
        map: &LoweringMap,
        paragraph: u32,
    ) -> Option<Source> {
        let (story, para_id) = map.paragraphs.get(paragraph as usize)?;
        let story = map.stories.get(*story as usize)?;
        if occurrence.root == occurrence.record.story {
            return Some(Source {
                story: story.clone(),
                para_id: para_id.clone(),
            });
        }
        let names = self.alias(&occurrence.root, &occurrence.record.story);
        Some(Source {
            story: names.stories.get(story)?.clone(),
            para_id: names
                .paragraphs
                .get(&(story.clone(), para_id.clone()))?
                .clone(),
        })
    }

    /// The names of `alias`'s content in the exported story holding the same content.
    fn alias(&mut self, alias: &str, exported: &str) -> Rc<AliasNames> {
        let key = (alias.to_owned(), exported.to_owned());
        if let Some(found) = self.aliases.get(&key) {
            return Rc::clone(found);
        }
        let names = Rc::new(match (self.maps.get(alias), self.maps.get(exported)) {
            (Some(alias), Some(exported)) => AliasNames::new(alias, exported),
            _ => AliasNames::default(),
        });
        self.aliases.insert(key, Rc::clone(&names));
        names
    }

    /// The projection of `story` in `view` with its paragraphs by id; duplicate ids resolve to
    /// nothing.
    fn projection(
        &mut self,
        story: &str,
        view: EditTextView,
    ) -> Option<(Rc<StoryView>, Rc<HashMap<String, usize>>)> {
        let key = (story.to_owned(), view_slot(view));
        if let Some(found) = self.paragraph_indexes.get(&key) {
            return found.clone();
        }
        let built = self.views.story(story, view).map(|projection| {
            let mut by_id: HashMap<String, usize> = HashMap::new();
            let mut duplicates = HashSet::new();
            for (position, paragraph) in projection.paragraphs.iter().enumerate() {
                if by_id.insert(paragraph.para_id.clone(), position).is_some() {
                    duplicates.insert(paragraph.para_id.clone());
                }
            }
            for duplicate in duplicates {
                by_id.remove(&duplicate);
            }
            (projection, Rc::new(by_id))
        });
        self.paragraph_indexes.insert(key, built.clone());
        built
    }

    fn atom_units(&mut self, root: &str, map: &LoweringMap) -> Rc<HashMap<(u32, u32), u64>> {
        if let Some(found) = self.atom_units.get(root) {
            return Rc::clone(found);
        }
        let mut units: HashMap<(u32, u32), u64> = HashMap::new();
        for span in map.spans.iter().filter(|span| span.atom) {
            *units.entry((span.paragraph, span.raw_start)).or_default() +=
                span.pm_end - span.pm_start;
        }
        let units = Rc::new(units);
        self.atom_units.insert(root.to_owned(), Rc::clone(&units));
        units
    }

    fn paragraph(
        &mut self,
        occurrence: &Occurrence,
        map: &LoweringMap,
        placed: &PlacedParagraph<'_>,
    ) {
        let Some(start) = placed.block.pm_start.map(|value| value as u64) else {
            return self.unmapped(occurrence, "A laid-out paragraph has no document position.");
        };
        let Some(source) = self.source(occurrence, map, start) else {
            return self.unmapped(occurrence, "A laid-out paragraph has no recorded source.");
        };
        let index: &'a NodeIndex = self.index;
        let Some(node) = index
            .paragraphs
            .get(&(source.story.clone(), source.para_id.clone()))
        else {
            if index.stories.contains_key(story_root_of(&source.story)) {
                self.unmapped(
                    occurrence,
                    "A laid-out paragraph matches no exported paragraph.",
                );
            }
            return;
        };
        self.wrap(occurrence, &node.controls, placed.repeated_header);
        self.placed_nodes.insert(node.id.clone());
        let region = Self::region_key(occurrence);
        let slices = line_window_slices(placed.block, placed.measure, placed.lines.clone());
        let geometry = match placed.frame {
            Some(frame) => Geometry::Frame(frame),
            None => Geometry::Ranges(
                region.clone(),
                vec![(
                    slices
                        .first()
                        .map_or(start as i64, |slice| slice.pm_start as i64),
                    slices
                        .last()
                        .map_or(start as i64 + 1, |slice| slice.pm_end as i64),
                )],
            ),
        };
        self.fragment(
            occurrence,
            node.id.clone(),
            node.id.clone(),
            node.anchor.clone(),
            FragmentSlice::Block,
            Some((placed.continued_from_previous, placed.continued_on_next)),
            placed.repeated_header,
            geometry,
        );
        let units = self.atom_units(&occurrence.root, map);
        // Text runs as contiguous story intervals with their display start, atoms with the
        // display units shown here.
        let mut texts: Vec<(u32, u32, u64)> = Vec::new();
        let mut atoms: Vec<(u32, u32, u64, Vec<(i64, i64)>)> = Vec::new();
        for slice in &slices {
            let (from, to) = (slice.pm_start as u64, slice.pm_end as u64);
            let first = map.spans.partition_point(|span| span.pm_end <= from);
            for span in map.spans[first..]
                .iter()
                .take_while(|span| span.pm_start < to)
            {
                let lo = from.max(span.pm_start);
                let hi = to.min(span.pm_end);
                if hi <= lo {
                    continue;
                }
                if span.atom {
                    match atoms.last_mut() {
                        Some(last) if last.0 == span.paragraph && last.1 == span.raw_start => {
                            last.2 += hi - lo;
                            last.3.push((lo as i64, hi as i64));
                        }
                        _ => atoms.push((
                            span.paragraph,
                            span.raw_start,
                            hi - lo,
                            vec![(lo as i64, hi as i64)],
                        )),
                    }
                    continue;
                }
                let raw_lo = span.raw_start + (lo - span.pm_start) as u32;
                let raw_hi = span.raw_start + (hi - span.pm_start) as u32;
                match texts.last_mut() {
                    Some(last) if last.1 == raw_lo && last.2 + u64::from(last.1 - last.0) == lo => {
                        last.1 = raw_hi;
                    }
                    _ => texts.push((raw_lo, raw_hi, lo)),
                }
            }
        }
        let mut inlines: Vec<(i64, String, Anchor, FragmentSlice, Geometry)> = Vec::new();
        for &(raw_start, raw_end, pm_start) in &texts {
            for &view in self.export_views {
                let Some((projection, by_id)) = self.projection(&source.story, view) else {
                    continue;
                };
                let Some(&position) = by_id.get(&source.para_id) else {
                    continue;
                };
                let paragraph = &projection.paragraphs[position];
                let low = paragraph.offset_of_raw(raw_start);
                let high = paragraph.offset_of_raw(raw_end);
                if high <= low {
                    continue;
                }
                let entries = &node.texts[view_slot(view) as usize];
                let first = entries.partition_point(|entry| entry.end <= low);
                for entry in entries[first..]
                    .iter()
                    .take_while(|entry| entry.start < high)
                {
                    let (start, end) = (entry.start.max(low), entry.end.min(high));
                    if end <= start {
                        continue;
                    }
                    let raw_lo = paragraph.raw_at(start);
                    let raw_hi = paragraph.raw_after(end);
                    let pm_lo = pm_start as i64 + i64::from(raw_lo.saturating_sub(raw_start));
                    let pm_hi = pm_start as i64 + i64::from(raw_hi.saturating_sub(raw_start));
                    let range = TextRange {
                        story: projection.story.clone(),
                        start: TextPosition {
                            para_id: paragraph.para_id.clone(),
                            offset: start,
                        },
                        end: TextPosition {
                            para_id: paragraph.para_id.clone(),
                            offset: end,
                        },
                        view,
                    };
                    inlines.push((
                        pm_lo,
                        entry.id.clone(),
                        entry.anchor.clone(),
                        FragmentSlice::Text { range },
                        Geometry::Ranges(region.clone(), vec![(pm_lo, pm_hi.max(pm_lo + 1))]),
                    ));
                }
            }
        }
        for (paragraph_index, raw, shown, ranges) in atoms {
            let whole = units
                .get(&(paragraph_index, raw))
                .is_none_or(|total| shown >= *total);
            let mut found = None;
            for &view in self.export_views {
                let Some((projection, by_id)) = self.projection(&source.story, view) else {
                    continue;
                };
                let Some(&position) = by_id.get(&source.para_id) else {
                    continue;
                };
                let paragraph = &projection.paragraphs[position];
                let offset = paragraph.offset_of_raw(raw);
                if offset < paragraph.len() && paragraph.raw_at(offset) == raw {
                    found = node.atoms.get(&(view_slot(view), offset)).cloned();
                    if found.is_some() {
                        break;
                    }
                }
            }
            let Some((id, range)) = found else {
                continue;
            };
            if !whole {
                self.diagnostics.push(PageDiagnostic {
                    code: PageDiagnosticCode::AnchorOnly,
                    node_id: Some(id.clone()),
                    page_index: Some(occurrence.record.page_index),
                    message: "Only part of this atom's content is on this page; its anchor is the whole atom.".to_owned(),
                });
            }
            inlines.push((
                ranges.first().map_or(0, |range| range.0),
                id,
                Anchor::Range(range.clone()),
                FragmentSlice::Atom {
                    range: Some(range),
                    coverage: if whole {
                        AtomCoverage::Whole
                    } else {
                        AtomCoverage::Partial
                    },
                },
                Geometry::Ranges(region.clone(), ranges),
            ));
        }
        inlines.sort_by_key(|(position, ..)| *position);
        for (_, id, anchor, slice, geometry) in inlines {
            self.fragment(
                occurrence,
                id,
                node.id.clone(),
                anchor,
                slice,
                None,
                placed.repeated_header,
                geometry,
            );
        }
    }

    fn table(&mut self, occurrence: &Occurrence, map: &LoweringMap, placed: &PlacedTable<'_>) {
        let Some(start) = placed.block.pm_start.map(|value| value as u64) else {
            return self.unmapped(occurrence, "A laid-out table has no document position.");
        };
        let Some((story, ordinal)) = map.table_at(start) else {
            return self.unmapped(occurrence, "A laid-out table has no recorded source.");
        };
        let Some(story) = map.stories.get(story as usize) else {
            return;
        };
        let (story, ordinal) = if occurrence.root == occurrence.record.story {
            (story.clone(), ordinal)
        } else {
            let names = self.alias(&occurrence.root, &occurrence.record.story);
            match (
                names.stories.get(story),
                names.tables.get(&(story.clone(), ordinal)),
            ) {
                (Some(story), Some(ordinal)) => (story.clone(), *ordinal),
                _ => {
                    return self.unmapped(
                        occurrence,
                        "A laid-out table of a merged header or footer matches no exported table.",
                    );
                }
            }
        };
        let index: &'a NodeIndex = self.index;
        let Some(node) = index.tables.get(&(story.clone(), ordinal)) else {
            if index.stories.contains_key(story_root_of(&story)) {
                self.unmapped(occurrence, "A laid-out table matches no exported table.");
            }
            return;
        };
        self.wrap(occurrence, &node.controls, placed.repeated_header);
        let rows = placed
            .rows
            .iter()
            .filter(|row| row.row_index < node.rows)
            .map(|row| FragmentRow {
                row_index: row.row_index as u32,
                continued_from_previous: row.continued_from_previous,
                continued_on_next: row.continued_on_next,
                repeated_header: row.repeated_header,
            })
            .collect();
        let geometry = match placed.frame {
            Some(frame) => Geometry::Frame(frame),
            None => Geometry::Ranges(
                Self::region_key(occurrence),
                vec![(
                    start as i64,
                    placed
                        .block
                        .pm_end
                        .map_or(start as i64 + 1, |end| end as i64),
                )],
            ),
        };
        self.fragment(
            occurrence,
            node.id.clone(),
            node.id.clone(),
            node.anchor.clone(),
            FragmentSlice::Table { rows },
            Some((placed.continued_from_previous, placed.continued_on_next)),
            placed.repeated_header,
            geometry,
        );
    }

    /// An image, shape, chart or text box laid out as a block of its own: the atom it stands
    /// for in its paragraph.
    fn drawing(&mut self, occurrence: &Occurrence, map: &LoweringMap, block: &LayoutBlock) {
        let Some(start) = block.pm_start().map(|value| value as u64) else {
            return;
        };
        let Some(span) = map.span_at(start).copied().filter(|span| span.atom) else {
            return;
        };
        let Some(source) = self.source_of(occurrence, map, span.paragraph) else {
            return;
        };
        let index: &'a NodeIndex = self.index;
        let Some(node) = index
            .paragraphs
            .get(&(source.story.clone(), source.para_id.clone()))
        else {
            return;
        };
        let mut found = None;
        for &view in self.export_views {
            let Some((projection, by_id)) = self.projection(&source.story, view) else {
                continue;
            };
            let Some(&position) = by_id.get(&source.para_id) else {
                continue;
            };
            let paragraph = &projection.paragraphs[position];
            let offset = paragraph.offset_of_raw(span.raw_start);
            if offset < paragraph.len() && paragraph.raw_at(offset) == span.raw_start {
                found = node.atoms.get(&(view_slot(view), offset)).cloned();
                if found.is_some() {
                    break;
                }
            }
        }
        let Some((id, range)) = found else {
            return;
        };
        self.wrap(occurrence, &node.controls, false);
        let block_id = node.id.clone();
        self.fragment(
            occurrence,
            id,
            block_id,
            Anchor::Range(range.clone()),
            FragmentSlice::Atom {
                range: Some(range),
                coverage: AtomCoverage::Whole,
            },
            None,
            false,
            Geometry::Ranges(
                Self::region_key(occurrence),
                vec![(span.pm_start as i64, span.pm_end as i64)],
            ),
        );
    }

    /// Diagnoses exported content no page shows.
    fn unplaced(&mut self, selected: &HashSet<StorySelection>) {
        for story in &self.content.stories {
            if story.kind == StoryKind::Comment {
                continue;
            }
            if !self.laid_out_roots.contains(&story.story) {
                self.diagnostics.push(PageDiagnostic {
                    code: PageDiagnosticCode::NotLaidOut,
                    node_id: None,
                    page_index: None,
                    message: format!("Story {} is not shown on any page.", story.story),
                });
            }
        }
        if selected.contains(&StorySelection::Comments) {
            self.diagnostics.push(PageDiagnostic {
                code: PageDiagnosticCode::NotLaidOut,
                node_id: None,
                page_index: None,
                message: "Comments are not laid out on pages.".to_owned(),
            });
        }
        let mut missing: Vec<&str> = self
            .index
            .paragraphs
            .values()
            .filter(|node| {
                self.laid_out_roots.contains(&node.root) && !self.placed_nodes.contains(&node.id)
            })
            .map(|node| node.id.as_str())
            .collect();
        missing.sort_unstable();
        for id in missing {
            self.diagnostics.push(PageDiagnostic {
                code: PageDiagnosticCode::NotLaidOut,
                node_id: Some(id.to_owned()),
                page_index: None,
                message: "This paragraph is not shown on any page.".to_owned(),
            });
        }
    }
}

fn story_root_of(story: &str) -> &str {
    source::story_root(story)
}

/// Sets continuation flags on fragments whose node the layout did not flag: each continues the
/// fragments of its node before it in the same flow, and is continued by those after.
fn link_continuations(items: &mut [MapItem]) {
    let mut first: HashMap<(String, String), usize> = HashMap::new();
    let mut last: HashMap<(String, String), usize> = HashMap::new();
    for (position, item) in items.iter().enumerate() {
        let MapItem::Fragment(draft) = item else {
            continue;
        };
        if draft.placed_flags || draft.fragment.repeated_table_header {
            continue;
        }
        let key = (draft.flow.clone(), draft.fragment.node_id.clone());
        first.entry(key.clone()).or_insert(position);
        last.insert(key, position);
    }
    for (position, item) in items.iter_mut().enumerate() {
        let MapItem::Fragment(draft) = item else {
            continue;
        };
        if draft.placed_flags || draft.fragment.repeated_table_header {
            continue;
        }
        let key = (draft.flow.clone(), draft.fragment.node_id.clone());
        draft.fragment.continued_from_previous = first.get(&key) != Some(&position);
        draft.fragment.continued_on_next = last.get(&key) != Some(&position);
    }
}

/// Fills every fragment's geometry from the display list the layout paints.
fn add_geometry(items: &mut [MapItem], display: &DisplayList) {
    let mut page_indexes: HashMap<(u32, RegionKey), Option<RangeRectIndex<'_>>> = HashMap::new();
    for item in items.iter_mut() {
        let MapItem::Fragment(draft) = item else {
            continue;
        };
        let page_index = draft.fragment.page_index;
        let rects: Vec<GeometryRect> = match &draft.geometry {
            Geometry::None => continue,
            Geometry::Frame([x, y, width, height]) => vec![GeometryRect {
                x: *x,
                y: *y,
                width: *width,
                height: *height,
            }],
            Geometry::Ranges(region, ranges) => {
                let key = (page_index, region.clone());
                let index = page_indexes
                    .entry(key)
                    .or_insert_with(|| region_index(display, page_index as usize, region));
                let Some(index) = index else {
                    continue;
                };
                ranges
                    .iter()
                    .flat_map(|(from, to)| index.rects(*from, *to))
                    .map(|rect: RangeRect| GeometryRect {
                        x: rect.x,
                        y: rect.y,
                        width: rect.width,
                        height: rect.height,
                    })
                    .collect()
            }
        };
        draft.fragment.geometry = Some(FragmentGeometry {
            unit: GeometryUnit::CssPx,
            origin: GeometryOrigin::PageTopLeft,
            rects,
        });
    }
}

fn region_index<'a>(
    display: &'a DisplayList,
    page_index: usize,
    region: &RegionKey,
) -> Option<RangeRectIndex<'a>> {
    let page = display.pages.get(page_index)?;
    let primitives = match region {
        RegionKey::Body => &page.primitives[..],
        RegionKey::Header(r_id) => {
            let band = page.header.as_ref().filter(|band| band.r_id == *r_id)?;
            &band.primitives[..]
        }
        RegionKey::Footer(r_id) => {
            let band = page.footer.as_ref().filter(|band| band.r_id == *r_id)?;
            &band.primitives[..]
        }
        RegionKey::Note(endnote, id) => page
            .note_areas
            .iter()
            .filter(|area| (area.kind.as_deref() == Some("endnote")) == *endnote)
            .flat_map(note_primitives)
            .find(|(note, _)| note == id)
            .map(|(_, primitives)| primitives)?,
    };
    Some(RangeRectIndex::new(primitives, page_index))
}

/// Admits pages, then occurrences and fragments in page order, then diagnostics, within the
/// fragment and byte limits. Whatever does not fit is left out and the map says so.
fn assemble(
    mut map: DocxLayoutMap,
    pages: Vec<ExportPage>,
    items: Vec<MapItem>,
    diagnostics: Vec<PageDiagnostic>,
    limits: &PageLimits,
) -> DocxLayoutMap {
    let stop = |page_index: Option<u32>| {
        PageDiagnostic {
        code: PageDiagnosticCode::Truncated,
        node_id: None,
        page_index,
        message: "The page map stopped at its limits; pages, occurrences and fragments after this point are not included.".to_owned(),
    }
    };
    let reserve = json_len(&stop(Some(u32::MAX))) + 1;
    let mut used = json_len(&map);
    let budget = limits.max_bytes.saturating_sub(reserve);
    let mut stopped_at: Option<Option<u32>> = None;
    let fits = |used: &mut usize, size: usize| {
        if *used + size + 1 > budget {
            false
        } else {
            *used += size + 1;
            true
        }
    };
    for page in pages {
        if !fits(&mut used, json_len(&page)) {
            stopped_at = Some(Some(page.page_index));
            break;
        }
        map.pages.push(page);
    }
    if stopped_at.is_none() {
        for item in items {
            match item {
                MapItem::Occurrence(occurrence) => {
                    if !fits(&mut used, json_len(&occurrence)) {
                        stopped_at = Some(Some(occurrence.page_index));
                        break;
                    }
                    map.occurrences.push(occurrence);
                }
                MapItem::Fragment(draft) => {
                    if map.fragments.len() >= limits.max_fragments
                        || !fits(&mut used, json_len(&draft.fragment))
                    {
                        stopped_at = Some(Some(draft.fragment.page_index));
                        break;
                    }
                    map.fragments.push(draft.fragment);
                }
            }
        }
    }
    for diagnostic in diagnostics {
        if !fits(&mut used, json_len(&diagnostic)) {
            stopped_at.get_or_insert(None);
            break;
        }
        map.diagnostics.push(diagnostic);
    }
    if let Some(page_index) = stopped_at {
        map.truncated = true;
        map.diagnostics.push(stop(page_index));
    }
    map
}
