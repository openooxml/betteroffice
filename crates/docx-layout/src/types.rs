//! Serde twins of the pagination type vocabulary.
//!
//! Ports `packages/core/src/layout/pagination/types.ts` (plus `measuredBlock.ts`
//! and `LayoutOptions`) field-for-field. Every name round-trips the TS
//! camelCase JSON exactly (`serde(rename_all = "camelCase")`), optional TS
//! fields are `Option<T>` and are omitted from output when `None` — matching
//! how `JSON.stringify` drops `undefined`. Unknown `kind` tags degrade to an
//! `Unsupported` variant so the engine can refuse gracefully instead of
//! failing to parse; unknown extra fields are ignored on input, mirroring the
//! TS structural contract.
//!
//! Types referenced from outside `types.ts` (`InlineSdtWidget`, `RevisionInfo`,
//! `CellMarker`) pass through as raw `serde_json::Value` — pagination never
//! reads them, and passthrough preserves them verbatim.

use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::BTreeMap;

// ---------------------------------------------------------------------------
// shared scalars
// ---------------------------------------------------------------------------

/// TS `BlockId = string | number` — passed through verbatim to fragments.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum BlockId {
    Num(f64),
    Str(String),
}

/// `{ w, h }` page-size pair.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Size {
    pub w: f64,
    pub h: f64,
}

/// TS `PageMargins` — `header`/`footer` are the only optional keys.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PageMargins {
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
    pub left: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer: Option<f64>,
}

/// One authored unequal-width column.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ColumnDefinition {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<f64>,
}

/// TS `ColumnLayout` (w:cols). `count` stays `f64` so arithmetic matches TS.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnLayout {
    pub count: f64,
    pub gap: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub equal_width: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub separator: Option<bool>,
}

/// TS `SectionBreakBlock['type']` / `LayoutOptions['bodyBreakType']`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SectionBreakType {
    Continuous,
    NextPage,
    EvenPage,
    OddPage,
}

// ---------------------------------------------------------------------------
// runs
// ---------------------------------------------------------------------------

/// TS `underline?: boolean | { style?, color? }`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum UnderlineSpec {
    Flag(bool),
    Styled {
        #[serde(skip_serializing_if = "Option::is_none")]
        style: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
    },
}

/// TS `HyperlinkInfo`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HyperlinkInfo {
    pub href: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_default_style: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub history: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_location: Option<String>,
}

/// TS `RunFontSlots`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunFontSlots {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ascii: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_ansi: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ascii_theme: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_ansi_theme: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia_theme: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs_theme: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hint: Option<String>,
}

/// TS `RunLanguageSlots`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunLanguageSlots {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latin: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<String>,
}

/// TS `RunFormatting` — character formatting shared by text/tab/field runs.
/// Pagination itself reads none of these; they ride along so resolved-line
/// run slices round-trip like the TS engine's.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunFormatting {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_slots: Option<RunFontSlots>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold_cs: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic_cs: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub complex_script: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub language: Option<RunLanguageSlots>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub letter_spacing: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub superscript: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subscript: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub all_caps: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position_px: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kerning_min_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub imprint: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emboss: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_shadow: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_outline: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_effect: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub modern_effects: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<HyperlinkInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_ref_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnote_ref_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comment_ids: Option<Vec<f64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_insertion: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_deletion: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub change_author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub change_date: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub change_revision_id: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
}

/// TS `TextRun`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    #[serde(flatten)]
    pub fmt: RunFormatting,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inline_sdt_widget: Option<Value>,
}

/// TS `TabRun`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TabRun {
    #[serde(flatten)]
    pub fmt: RunFormatting,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_glyphs: Option<Value>,
}

/// One axis of TS `ImageRunPosition`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AxisPosition {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pos_offset: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relative_to: Option<String>,
}

/// TS `ImageRunPosition`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageRunPosition {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<AxisPosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical: Option<AxisPosition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_simple_pos: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub simple_pos: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub behind_doc: Option<bool>,
}

/// TS `ImageRun` (no shared `RunFormatting`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageRun {
    pub src: String,
    pub width: f64,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImageRunPosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_mode: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub css_float: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop_right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop_bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_deg: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip_h: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip_v: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_bounds: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_polygon: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_overlap: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_in_cell: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_extent: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<CellBorderSpec>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<HyperlinkInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_insertion: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_deletion: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub change_author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub change_date: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub change_revision_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `LineBreakRun`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LineBreakRun {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `FieldRun`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FieldRun {
    #[serde(flatten)]
    pub fmt: RunFormatting,
    pub field_type: String,
    /// raw Word field type token when `field_type` collapsed it to a painter
    /// category — inert a11y identity, never evaluated
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub raw_type: Option<String>,
    /// raw field instruction text carried INERT for a11y announcement only
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub instruction: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fallback: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `Run` union. Unknown kinds degrade to `Unsupported` (the engine refuses
/// the document rather than mangling it).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind")]
pub enum Run {
    #[serde(rename = "text")]
    Text(TextRun),
    #[serde(rename = "tab")]
    Tab(TabRun),
    #[serde(rename = "image")]
    Image(ImageRun),
    #[serde(rename = "lineBreak")]
    LineBreak(LineBreakRun),
    #[serde(rename = "field")]
    Field(FieldRun),
    #[serde(other, rename = "unsupported")]
    Unsupported,
}

impl Run {
    /// PM start offset, regardless of run flavor.
    pub fn pm_start(&self) -> Option<f64> {
        match self {
            Run::Text(r) => r.pm_start,
            Run::Tab(r) => r.pm_start,
            Run::Image(r) => r.pm_start,
            Run::LineBreak(r) => r.pm_start,
            Run::Field(r) => r.pm_start,
            Run::Unsupported => None,
        }
    }

    /// PM end offset, regardless of run flavor.
    pub fn pm_end(&self) -> Option<f64> {
        match self {
            Run::Text(r) => r.pm_end,
            Run::Tab(r) => r.pm_end,
            Run::Image(r) => r.pm_end,
            Run::LineBreak(r) => r.pm_end,
            Run::Field(r) => r.pm_end,
            Run::Unsupported => None,
        }
    }
}

// ---------------------------------------------------------------------------
// paragraph attributes
// ---------------------------------------------------------------------------

/// TS `ParagraphSpacing`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphSpacing {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_unit: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_rule: Option<String>,
}

/// TS `ParagraphAttrs['spacingExplicit']`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SpacingExplicit {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after: Option<bool>,
}

/// TS `ParagraphIndent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphIndent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_line: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hanging: Option<f64>,
}

/// TS `TabStop`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TabStop {
    pub val: String,
    pub pos: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub leader: Option<String>,
}

/// TS `BorderStyle` (one paragraph border edge).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BorderStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space: Option<f64>,
}

/// TS `ParagraphBorders`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParagraphBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub between: Option<BorderStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar: Option<BorderStyle>,
}

/// TS `ListNumPr`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ListNumPr {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ilvl: Option<f64>,
}

/// TS `ParagraphAttrs`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphAttrs {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing: Option<ParagraphSpacing>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing_explicit: Option<SpacingExplicit>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<ParagraphIndent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub contextual_spacing: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<ParagraphBorders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tabs: Option<Vec<TabStop>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_pr: Option<ListNumPr>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_marker: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_is_bullet: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_marker_hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_marker_font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_marker_font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_marker_suffix: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_marker_revision: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_tab_stop_twips: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suppress_empty_paragraph_height: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub p_pr_ins: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub p_pr_del: Option<Value>,
}

/// TS `SdtGroup`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SdtGroup {
    pub id: String,
    pub sdt_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alias: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lock: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checked: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bound: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub repeating_item: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_id: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_state: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub properties: Option<Value>,
}

// ---------------------------------------------------------------------------
// flow blocks
// ---------------------------------------------------------------------------

/// TS `ParagraphBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    pub runs: Vec<Run>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub attrs: Option<ParagraphAttrs>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `CellBorderSpec`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellBorderSpec {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
}

/// TS `CellBorders`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<CellBorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<CellBorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<CellBorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<CellBorderSpec>,
}

/// TS `TableCell['padding']` / `TextBoxBlock['margins']` box.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BoxEdges {
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
    pub left: f64,
}

/// Typed `tblW`/`tcW`/row-before/after preferred width.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PreferredWidth {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<String>,
}

/// TS `TableCell`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    pub id: BlockId,
    pub blocks: Vec<LayoutBlock>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub col_span: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_span: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_value: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preferred_width: Option<PreferredWidth>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_start: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_content_width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_content_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<CellBorders>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub padding: Option<BoxEdges>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_wrap: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tracked_marker: Option<Value>,
}

/// TS `TableRow`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    pub id: BlockId,
    pub cells: Vec<TableCell>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_rule: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_header: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cant_split: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_before: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_after: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_before: Option<PreferredWidth>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_after: Option<PreferredWidth>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tracked_ins: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tracked_del: Option<Value>,
}

/// TS `FloatingTablePosition` (w:tblpPr, px).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FloatingTablePosition {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horz_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_x_spec: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_y_spec: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top_from_text: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_from_text: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom_from_text: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_from_text: Option<f64>,
}

/// TS `TableBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    pub rows: Vec<TableRow>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_widths: Option<Vec<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_widths: Option<Vec<f64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preferred_width: Option<PreferredWidth>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_mode: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_algorithm: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_cascade: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub justification: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub floating: Option<FloatingTablePosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `ImageBlock['anchor']`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageAnchor {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_anchored: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_h: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset_v: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub behind_doc: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<ImageRunPosition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_overlap: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_in_cell: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_polygon: Option<Vec<Value>>,
}

/// TS `ImageBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    pub src: String,
    pub width: f64,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_deg: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip_h: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip_v: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotation_bounds: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor: Option<ImageAnchor>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hlink_href: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crop: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<CellBorderSpec>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `ShapeBlock`. Complex DrawingML paint/scene payloads pass through as
/// JSON; pagination only needs the bbox and placement metadata.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    pub shape_type: String,
    pub geometry_path: Vec<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub transform: Option<Value>,
    pub width: f64,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_text: Option<Vec<ParagraphBlock>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_measures: Option<Vec<ParagraphExtent>>,
    pub children: Vec<ShapeBlock>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scene: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<Value>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_body_properties: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImageRunPosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub behind_doc: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `ChartBlock`; the normalized chart model is renderer-owned JSON.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    pub chart: Value,
    pub width: f64,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImageRunPosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub relative_height: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub behind_doc: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `SectionBreakBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionBreakBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
    pub break_type: Option<SectionBreakType>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_size: Option<Size>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<PageMargins>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
}

/// TS `PageBreakBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageBreakBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `ColumnBreakBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnBreakBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `TextBoxBlock`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBoxBlock {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sdt_groups: Option<Vec<SdtGroup>>,
    pub id: BlockId,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<BoxEdges>,
    pub content: Vec<ParagraphBlock>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_mode: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub css_float: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_target: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ImageRunPosition>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dist_right: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
}

/// TS `LayoutBlock` union. Unknown kinds parse to `Unsupported` so the
/// engine can return the fallback signal instead of a parse error.
// paragraph attrs make the variant big, but blocks are deserialized once and
// walked — not a hot allocation path; boxing would only obscure the TS mirror
#[allow(clippy::large_enum_variant)]
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind")]
pub enum LayoutBlock {
    #[serde(rename = "paragraph")]
    Paragraph(ParagraphBlock),
    #[serde(rename = "table")]
    Table(TableBlock),
    #[serde(rename = "image")]
    Image(ImageBlock),
    #[serde(rename = "shape")]
    Shape(ShapeBlock),
    #[serde(rename = "chart")]
    Chart(ChartBlock),
    #[serde(rename = "textBox")]
    TextBox(TextBoxBlock),
    #[serde(rename = "sectionBreak")]
    SectionBreak(SectionBreakBlock),
    #[serde(rename = "pageBreak")]
    PageBreak(PageBreakBlock),
    #[serde(rename = "columnBreak")]
    ColumnBreak(ColumnBreakBlock),
    #[serde(other, rename = "unsupported")]
    Unsupported,
}

// ---------------------------------------------------------------------------
// extents (measurement results)
// ---------------------------------------------------------------------------

/// TS `TypesetRowSegment`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRowSegment {
    pub head_run: usize,
    pub head_char: usize,
    pub tail_run: usize,
    pub tail_char: usize,
    pub left_offset: f64,
    pub available_width: f64,
    pub width: f64,
}

/// TS `TypesetRunAdvance`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRunAdvance {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_char: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_char: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub advance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
}

/// TS `TypesetClusterAdvance`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetClusterAdvance {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_char: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_char: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub advance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_offset: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
}

/// TS `TypesetBidiSlice`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetBidiSlice {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_char: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_char: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub advance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visual_order: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
}

/// TS `TypesetRow` — one measured paragraph line.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRow {
    pub head_run: usize,
    pub head_char: usize,
    pub tail_run: usize,
    pub tail_char: usize,
    pub width: f64,
    pub ascent: f64,
    pub descent: f64,
    pub line_height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_offset: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_offset: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub segments: Option<Vec<TypesetRowSegment>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub float_skip_before: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_advances: Option<Vec<TypesetRunAdvance>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cluster_advances: Option<Vec<TypesetClusterAdvance>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_slices: Option<Vec<TypesetBidiSlice>>,
}

/// TS `ParagraphExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphExtent {
    pub lines: Vec<TypesetRow>,
    pub total_height: f64,
}

/// TS `ImageExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ImageExtent {
    pub width: f64,
    pub height: f64,
}

/// TS `ShapeExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeExtent {
    pub width: f64,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_measures: Option<Vec<ParagraphExtent>>,
}

/// TS `ChartExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartExtent {
    pub width: f64,
    pub height: f64,
}

/// TS `TableCellExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellExtent {
    pub blocks: Vec<BlockExtent>,
    pub width: f64,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub col_span: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_span: Option<f64>,
}

/// TS `TableRowExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRowExtent {
    pub cells: Vec<TableCellExtent>,
    pub height: f64,
}

/// TS `TableExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableExtent {
    pub rows: Vec<TableRowExtent>,
    pub column_widths: Vec<f64>,
    pub total_width: f64,
    pub total_height: f64,
}

/// TS `TextBoxExtent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBoxExtent {
    pub width: f64,
    pub height: f64,
    pub inner_measures: Vec<ParagraphExtent>,
}

/// TS `BlockExtent` union (break placeholders carry no fields).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind")]
pub enum BlockExtent {
    #[serde(rename = "paragraph")]
    Paragraph(ParagraphExtent),
    #[serde(rename = "image")]
    Image(ImageExtent),
    #[serde(rename = "shape")]
    Shape(ShapeExtent),
    #[serde(rename = "chart")]
    Chart(ChartExtent),
    #[serde(rename = "table")]
    Table(TableExtent),
    #[serde(rename = "textBox")]
    TextBox(TextBoxExtent),
    #[serde(rename = "sectionBreak")]
    SectionBreak,
    #[serde(rename = "pageBreak")]
    PageBreak,
    #[serde(rename = "columnBreak")]
    ColumnBreak,
    #[serde(other, rename = "unsupported")]
    Unsupported,
}

/// TS `MeasuredBlock` (measuredBlock.ts) — a block paired with its measure.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MeasuredBlock {
    pub block: LayoutBlock,
    pub measure: BlockExtent,
}

// ---------------------------------------------------------------------------
// layout options
// ---------------------------------------------------------------------------

/// TS `LayoutOptions`. `footnoteReservedHeights` is a `Map` in TS; the JSON
/// boundary carries it as an object keyed by decimal page number (see
/// `scripts/export-golden-fixtures.ts`).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", default)]
pub struct LayoutOptions {
    #[serde(default)]
    pub contract_version: Option<u32>,
    pub page_size: Option<Size>,
    pub margins: Option<PageMargins>,
    pub final_page_size: Option<Size>,
    pub final_margins: Option<PageMargins>,
    pub columns: Option<ColumnLayout>,
    pub page_gap: Option<f64>,
    pub default_line_height: Option<f64>,
    pub header_content_heights: Option<Value>,
    pub footer_content_heights: Option<Value>,
    pub title_page: Option<bool>,
    pub even_and_odd_headers: Option<bool>,
    pub footnote_reserved_heights: Option<BTreeMap<String, f64>>,
    pub body_break_type: Option<SectionBreakType>,
    #[serde(default)]
    pub sections: Option<Vec<SectionLayoutContract>>,
}

/// Additive TS `LayoutOptions.sections[]` contract. The spine ignores it until
/// Batch D consumes per-section state; serde still preserves the typed seam.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionLayoutContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_size: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margins: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer_refs: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_settings: Option<Value>,
}

/// The `{ measured, options }` JSON envelope the engine consumes — the same
/// pair `layoutDocument(measured, options)` takes in TS.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Input {
    pub measured: Vec<MeasuredBlock>,
    #[serde(default)]
    pub options: LayoutOptions,
}

// ---------------------------------------------------------------------------
// fragments and pages (output)
// ---------------------------------------------------------------------------

use crate::resolve_lines::ResolvedLine;

/// TS `ParagraphFragment`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphFragment {
    pub block_id: BlockId,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    pub from_line: usize,
    pub to_line: usize,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub carried_from_prev: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub carried_to_next: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub resolved_lines: Option<Vec<ResolvedLine>>,
}

/// TS `TableFragment`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableFragment {
    pub block_id: BlockId,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    pub row_start: usize,
    pub row_end: usize,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_floating: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub carried_from_prev: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub carried_to_next: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_row_count: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub clip_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub clip_bottom: Option<f64>,
}

/// TS `ImageFragment`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageFragment {
    pub block_id: BlockId,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_anchored: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub z_index: Option<f64>,
}

/// TS `ShapeFragment`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeFragment {
    pub block_id: BlockId,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_anchored: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub z_index: Option<f64>,
}

/// TS `ChartFragment`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartFragment {
    pub block_id: BlockId,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_anchored: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub z_index: Option<f64>,
}

/// TS `TextBoxFragment`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBoxFragment {
    pub block_id: BlockId,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    pub height: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub is_floating: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub z_index: Option<f64>,
}

/// TS `Fragment` union.
#[derive(Debug, Clone, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum Fragment {
    Paragraph(ParagraphFragment),
    Table(TableFragment),
    Image(ImageFragment),
    Shape(ShapeFragment),
    Chart(ChartFragment),
    TextBox(TextBoxFragment),
}

impl Fragment {
    /// Position the fragment (the paginator writes `x`/`y` on placement,
    /// mirroring `addFragment`'s mutation of the TS object).
    pub fn set_xy(&mut self, x: f64, y: f64) {
        match self {
            Fragment::Paragraph(f) => {
                f.x = x;
                f.y = y;
            }
            Fragment::Table(f) => {
                f.x = x;
                f.y = y;
            }
            Fragment::Image(f) => {
                f.x = x;
                f.y = y;
            }
            Fragment::Shape(f) => {
                f.x = x;
                f.y = y;
            }
            Fragment::Chart(f) => {
                f.x = x;
                f.y = y;
            }
            Fragment::TextBox(f) => {
                f.x = x;
                f.y = y;
            }
        }
    }
}

/// TS `Page['headerFooterRefs']`. Never set by the spine; ships for the
/// header/footer port.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterRefs {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_default: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_first: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_even: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_default: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_first: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_even: Option<String>,
}

/// TS `Page`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Page {
    pub number: u32,
    pub fragments: Vec<Fragment>,
    pub margins: PageMargins,
    pub size: Size,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_footer_refs: Option<HeaderFooterRefs>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_ids: Option<Vec<f64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_reserved_height: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_columns: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
}

/// TS `HeaderFooterLayout`. Never emitted by the spine; ships for the
/// header/footer port.
#[derive(Debug, Clone, Serialize)]
pub struct HeaderFooterLayout {
    pub height: f64,
    pub fragments: Vec<Fragment>,
}

/// TS `Layout` — the paginator's complete result. `checkpoints` (derived
/// resume bookmarks, omitted from golden serialization in TS) are not
/// produced by this port; they affect no layout output.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Layout {
    pub page_size: Size,
    pub pages: Vec<Page>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub headers: Option<BTreeMap<String, HeaderFooterLayout>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footers: Option<BTreeMap<String, HeaderFooterLayout>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_gap: Option<f64>,
}

// ---------------------------------------------------------------------------
// additive Word-parity contracts
// ---------------------------------------------------------------------------

/// Optional page metadata added by the versioned TS `Page` contract. Kept
/// separate from the actively-produced spine `Page` until Batch D populates it,
/// so unchanged layout output remains byte-identical.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_page_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_page_number: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_label: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_distance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_distance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_areas: Option<Vec<NoteAreaContract>>,
}

/// Additive TS `PageMargins` metadata.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct PageMarginsContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gutter: Option<f64>,
}

/// Additive TS `ColumnLayout` metadata.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct ColumnLayoutContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<Vec<ColumnDefinition>>,
}

/// Additive TS `SectionBreakBlock` metadata. Kept separate to avoid changing
/// active placement constructors before Batch D consumes these fields.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionBreakContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer_refs: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_settings: Option<Value>,
}

/// TS `NoteLayoutItem` / display-list note item shared payload.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteLayoutItemContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_label: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blocks: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub measures: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor_doc_start: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor_doc_end: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_mark_follows: Option<bool>,
}

/// TS `NoteAreaContract` / `DisplayListNoteArea` serde twin.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteAreaContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub placement: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<NoteLayoutItemContract>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes: Option<Vec<NoteLayoutItemContract>>,
}

/// TS display output `NoteRegion` (primitive arrays remain opaque until the
/// display-list producer owns the active primitive enum arm).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayNoteRegionContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator_primitives: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub primitives: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_ids: Option<Vec<i64>>,
}

/// Comment/reviewer presentation attached to display primitives.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayCommentMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub palette_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selected: Option<bool>,
}

/// Scoped clip/group metadata. Batch F activates the group primitive arm.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayClipGroupMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clip: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
}

/// Standalone TS `ClipGroupPrimitive` contract. Batch F activates it in the
/// live primitive union and producer after updating replay/mirror switches.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayClipGroupPrimitiveContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clip: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub primitives: Option<Vec<Value>>,
    #[serde(default, flatten)]
    pub attrs: DisplayPrimitiveMetadata,
}

/// New additive members of TS `DocAttrs`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayPrimitiveMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_target: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_history: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_doc_location: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub aria_label: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub aria_description: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden_object: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<DisplayCommentMetadata>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clip_group: Option<DisplayClipGroupMetadata>,
}

/// Additive TS `DisplayList`/`DisplayPage` metadata.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayListContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contract_version: Option<u32>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayPageContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_areas: Option<Vec<NoteAreaContract>>,
}

/// New members on `PositionedGlyph`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PositionedGlyphContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub logical_order: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_level: Option<u8>,
}

/// New members on `TextRunPrimitive`/`GlyphRunPrimitive`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayTextContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader_glyphs: Option<LeaderGlyphContract>,
    /// Modern w14 text effects payload (glow/shadow/reflection/textFill/
    /// textOutline), mirrored from `TextModernEffects` in displayList.ts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub modern_effects: Option<Value>,
    /// GlyphRun only: resolved CSS font shorthand for the fillText safety net
    /// when glyph outlines are unavailable.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fallback_font: Option<String>,
}

/// New members on `RectPrimitive`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct DisplayRectContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
}

/// New members on `LinePrimitive`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayLineContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    /// Owner class of the retained border recipe (`cell`/`fragment`/...).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_owner: Option<String>,
    /// Owning table grid cell for table-border/table-cut lines (mirrors the
    /// `TableCellRef` the line carries through its flattened DocAttrs).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell: Option<Value>,
    /// Enclosing table-fragment identity for table-border/table-cut lines.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table: Option<Value>,
}

/// New members on `ImagePrimitive`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayImageContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip_h: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip_v: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content_frame: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<Value>,
}

/// New members on `ShapePrimitive`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayShapeContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_paint: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effects: Option<Vec<Value>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect_extent: Option<Value>,
}

/// New members on `DecorationPrimitive`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayDecorationContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight_slice: Option<HighlightSliceContract>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
}

/// TS `LeaderGlyphPrimitive` / `leaderGlyphs` metadata.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LeaderGlyphContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glyph: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline_y: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub advance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
}

/// TS `DecorationPrimitive.highlightSlice`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HighlightSliceContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_start: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_end: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ascent: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub descent: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub includes_trailing_whitespace: Option<bool>,
}

/// TS display-list build-envelope additions.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayListBuildContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contract_version: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_areas: Option<Vec<NoteAreaContract>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub resolved_comment_ids: Option<Vec<i64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment_authors: Option<Vec<DisplayCommentAuthorContract>>,
}

/// Additive header/footer envelope and watermark metadata.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayHeaderFooterContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_index: Option<u64>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct DisplayWatermarkContractMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decorative: Option<bool>,
}

/// TS `DisplayListCommentAuthor`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DisplayCommentAuthorContract {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub palette_index: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[cfg(test)]
mod parity_contract_tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn old_optional_contracts_deserialize_to_noop_defaults() {
        let primitive: DisplayPrimitiveMetadata = serde_json::from_value(json!({})).unwrap();
        let build: DisplayListBuildContractMetadata = serde_json::from_value(json!({})).unwrap();
        let row: TypesetRow = serde_json::from_value(json!({
            "headRun": 0,
            "headChar": 0,
            "tailRun": 0,
            "tailChar": 0,
            "width": 0.0,
            "ascent": 0.0,
            "descent": 0.0,
            "lineHeight": 0.0
        }))
        .unwrap();

        assert_eq!(primitive, DisplayPrimitiveMetadata::default());
        assert_eq!(build, DisplayListBuildContractMetadata::default());
        assert!(row.run_advances.is_none());
        assert!(row.cluster_advances.is_none());
        assert!(row.bidi_slices.is_none());
    }

    #[test]
    fn present_optional_contracts_round_trip_camel_case() {
        let value = json!({
            "contractVersion": 1,
            "noteAreas": [{
                "pageIndex": 0,
                "kind": "footnote",
                "placement": "pageBottom",
                "notes": [{ "id": 7, "displayLabel": "1" }]
            }],
            "resolvedCommentIds": [4]
        });
        let contract: DisplayListBuildContractMetadata =
            serde_json::from_value(value.clone()).unwrap();
        assert_eq!(serde_json::to_value(contract).unwrap(), value);

        let attrs_value = json!({
            "linkTitle": "ScreenTip",
            "logicalOrder": 3,
            "bidiLevel": 1,
            "decorative": true,
            "clipGroup": { "id": "cell-1", "opacity": 0.5 }
        });
        let attrs: DisplayPrimitiveMetadata = serde_json::from_value(attrs_value.clone()).unwrap();
        assert_eq!(serde_json::to_value(attrs).unwrap(), attrs_value);
    }
}
