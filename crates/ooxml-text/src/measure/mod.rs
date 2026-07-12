mod floats;
mod input;
mod line_filler;
mod list_marker;
mod prepare;
mod tabs;

pub use input::{
    AttrsIn, BlockIn, CompatIn, DefaultsIn, FloatSegmentIn, FloatZoneIn, IndentIn, MeasureInput,
    RunIn, SpacingIn, TabStopIn,
};

use crate::font_store::{FontId, FontStore};

/// Why measurement refused an input.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MeasureError {
    Unsupported(String),
    /// The input violates the measurement contract (malformed JSON, unknown
    /// font ids, non-finite numbers).
    Invalid(String),
}

impl std::fmt::Display for MeasureError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MeasureError::Unsupported(reason) => write!(f, "UNSUPPORTED: {reason}"),
            MeasureError::Invalid(reason) => write!(f, "invalid: {reason}"),
        }
    }
}

impl std::error::Error for MeasureError {}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRowOut {
    pub head_run: u32,
    pub head_char: u32,
    pub tail_run: u32,
    pub tail_char: u32,
    pub width: f32,
    pub ascent: f32,
    pub descent: f32,
    pub line_height: f32,
    /// Px from the content left edge (floats); `Some` only when > 0.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_offset: Option<f32>,
    /// Px from the content right edge (floats); `Some` only when > 0.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_offset: Option<f32>,
    /// Split strips for centered floating exclusions.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub segments: Option<Vec<TypesetRowSegmentOut>>,
    /// Vertical space inserted before an obstructed line.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub float_skip_before: Option<f32>,
    /// Exact advances for run slices, emitted in visual paint order.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_advances: Option<Vec<TypesetRunAdvanceOut>>,
    /// Exact shaped cluster advances and visual x offsets.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cluster_advances: Option<Vec<TypesetClusterAdvanceOut>>,
    /// Bidi slices keep logical identity separate from visual paint order.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi_slices: Option<Vec<TypesetBidiSliceOut>>,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRunAdvanceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub logical_order: u32,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetClusterAdvanceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub x_offset: f32,
    pub bidi_level: u8,
    pub logical_order: u32,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetBidiSliceOut {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub advance: f32,
    pub bidi_level: u8,
    pub visual_order: u32,
    pub logical_order: u32,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TypesetRowSegmentOut {
    pub head_run: u32,
    pub head_char: u32,
    pub tail_run: u32,
    pub tail_char: u32,
    pub left_offset: f32,
    pub available_width: f32,
    pub width: f32,
}

#[derive(Debug, Clone, PartialEq, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphExtentOut {
    /// Always `"paragraph"`.
    pub kind: &'static str,
    pub lines: Vec<TypesetRowOut>,
    pub total_height: f32,
}

/// Security caps on file-derived counts.
pub(crate) const MAX_RUNS: usize = 10_000;
pub(crate) const MAX_RUN_TEXT_BYTES: usize = 1_000_000;
pub(crate) const MAX_LINES: usize = 100_000;
pub(crate) const MAX_TAB_STOPS: usize = 1_000;
pub(crate) const MAX_FLOAT_ZONES: usize = 200;
pub(crate) const MAX_ZONE_SEGMENTS: usize = 100;

pub(crate) fn pt_to_px(pt: f32) -> f32 {
    pt * 96.0 / 72.0
}

/// Measure one paragraph.
pub fn measure_paragraph(
    store: &FontStore,
    input: &MeasureInput,
) -> Result<ParagraphExtentOut, MeasureError> {
    if input.block.kind != "paragraph" {
        return Err(MeasureError::Unsupported(format!(
            "block kind {:?}",
            input.block.kind
        )));
    }
    if !input.max_width.is_finite() {
        return Err(MeasureError::Unsupported("non-finite maxWidth".to_string()));
    }
    input::validate_pt_size(input.defaults.font_size, "defaults.fontSize")?;

    let runs = &input.block.runs;
    if runs.len() > MAX_RUNS {
        return Err(MeasureError::Unsupported(format!(
            "too many runs ({} > {MAX_RUNS})",
            runs.len()
        )));
    }

    let attrs = input.block.attrs.as_ref();
    let spacing = attrs.and_then(|a| a.spacing.as_ref());
    if let Some(sp) = spacing {
        sp.validate()?;
    }
    if let Some(ind) = attrs.and_then(|a| a.indent.as_ref()) {
        ind.validate()?;
    }
    if let Some(tabs) = attrs.and_then(|a| a.tabs.as_deref()) {
        input::validate_tabs(tabs)?;
    }
    let zones = input.floating_zones.as_deref().unwrap_or(&[]);
    let paragraph_y_offset = input.paragraph_y_offset.unwrap_or(0.0);
    input::validate_float_context(zones, paragraph_y_offset)?;

    if runs.is_empty() {
        if attrs.is_some_and(|a| a.suppress_empty_paragraph_height) {
            return Ok(ParagraphExtentOut {
                kind: "paragraph",
                lines: vec![zero_row()],
                total_height: 0.0,
            });
        }
        let size_pt = attrs
            .and_then(|a| a.default_font_size)
            .unwrap_or(input.defaults.font_size);
        input::validate_pt_size(size_pt, "attrs.defaultFontSize")?;
        let family = attrs
            .and_then(|a| a.default_font_family.as_deref())
            .unwrap_or(&input.defaults.font_family);
        let font = regular_chain_head(store, input, family)?;
        return line_filler::empty_paragraph_extent(store, font, size_pt, spacing, &input.compat);
    }

    if runs.len() == 1 && runs[0].kind == "text" && is_whitespace_only(&runs[0]) {
        let run = &runs[0];
        let size_pt = run
            .font_size
            .or_else(|| attrs.and_then(|a| a.default_font_size))
            .unwrap_or(input.defaults.font_size);
        input::validate_pt_size(size_pt, "run.fontSize")?;
        let family = run
            .font_family
            .as_deref()
            .or_else(|| attrs.and_then(|a| a.default_font_family.as_deref()))
            .unwrap_or(&input.defaults.font_family);
        let font = regular_chain_head(store, input, family)?;
        return line_filler::empty_paragraph_extent(store, font, size_pt, spacing, &input.compat);
    }

    let marker_inline_width = match attrs {
        Some(a) if a.indent.as_ref().and_then(|i| i.hanging).unwrap_or(0.0) == 0.0 => {
            list_marker::list_marker_inline_width(store, input, a)?
        }
        _ => 0.0,
    };

    let prepared = prepare::prepare_runs(store, input)?;

    let indent = attrs.and_then(|a| a.indent.as_ref());
    let indent_left = indent.and_then(|i| i.left).unwrap_or(0.0);
    let indent_right = indent.and_then(|i| i.right).unwrap_or(0.0);
    let first_line_offset = indent.and_then(|i| i.first_line).unwrap_or(0.0)
        - indent.and_then(|i| i.hanging).unwrap_or(0.0);
    let body_width = (input.max_width - indent_left - indent_right).max(1.0);
    let first_line_width = (body_width - first_line_offset - marker_inline_width).max(1.0);

    line_filler::fill(line_filler::FillParams {
        store,
        prepared: &prepared,
        spacing,
        body_width,
        first_line_width,
        default_font_size_pt: input.defaults.font_size,
        compat: &input.compat,
        tabs: attrs.and_then(|a| a.tabs.as_deref()).unwrap_or(&[]),
        indent_left_px: indent_left,
        first_line_offset_px: first_line_offset,
        zones,
        paragraph_y_offset,
        authoritative_shaping: input.authoritative_shaping,
    })
}

pub fn measure_paragraph_json(store: &FontStore, input: &str) -> Result<String, String> {
    let parsed: MeasureInput =
        serde_json::from_str(input).map_err(|e| format!("invalid: parse: {e}"))?;
    let extent = measure_paragraph(store, &parsed).map_err(|e| e.to_string())?;
    serde_json::to_string(&extent).map_err(|e| format!("invalid: serialize: {e}"))
}

fn zero_row() -> TypesetRowOut {
    TypesetRowOut {
        head_run: 0,
        head_char: 0,
        tail_run: 0,
        tail_char: 0,
        width: 0.0,
        ascent: 0.0,
        descent: 0.0,
        line_height: 0.0,
        left_offset: None,
        right_offset: None,
        segments: None,
        float_skip_before: None,
        run_advances: None,
        cluster_advances: None,
        bidi_slices: None,
    }
}

fn is_whitespace_only(run: &RunIn) -> bool {
    match run.text.as_deref() {
        None => true,
        Some(t) => t.chars().all(|c| c == '\u{00a0}' || c.is_whitespace()),
    }
}

fn regular_chain_head(
    store: &FontStore,
    input: &MeasureInput,
    family: &str,
) -> Result<FontId, MeasureError> {
    let chain = input.chain_for(family, false, false)?;
    prepare::validate_chain(store, &chain)?;
    Ok(chain[0])
}

#[cfg(test)]
mod authoritative_tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn authoritative_json_uses_one_advance_source_for_rows_runs_clusters_and_bidi() {
        let mut store = FontStore::new();
        store.register(FIXTURE.to_vec()).unwrap();
        let input: MeasureInput = serde_json::from_value(serde_json::json!({
            "block": {
                "kind": "paragraph",
                "runs": [{
                    "kind": "text",
                    "text": "Latin e\u{301} ffi אב",
                    "letterSpacing": 1.25,
                    "allCaps": false,
                    "kerningMinPt": 14.0
                }]
            },
            "maxWidth": 1000.0,
            "fontChains": { "liberation sans|0|0": [0] },
            "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" },
            "authoritativeShaping": true
        }))
        .unwrap();
        let extent = measure_paragraph(&store, &input).unwrap();
        let line = &extent.lines[0];
        let clusters = line.cluster_advances.as_ref().unwrap();
        let runs = line.run_advances.as_ref().unwrap();
        let slices = line.bidi_slices.as_ref().unwrap();
        let cluster_sum: f32 = clusters.iter().map(|cluster| cluster.advance).sum();
        let run_sum: f32 = runs.iter().map(|run| run.advance).sum();
        let slice_sum: f32 = slices.iter().map(|slice| slice.advance).sum();
        assert!((cluster_sum - line.width).abs() < 0.001);
        assert!((run_sum - line.width).abs() < 0.001);
        assert!((slice_sum - line.width).abs() < 0.001);
        assert!(
            clusters
                .iter()
                .any(|cluster| cluster.end_char - cluster.start_char > 1)
        );
        assert!(slices.iter().any(|slice| slice.bidi_level % 2 == 1));
    }

    #[test]
    fn rotated_inline_image_uses_transformed_footprint_for_flow() {
        let store = FontStore::new();
        let input: MeasureInput = serde_json::from_value(serde_json::json!({
            "block": {
                "kind": "paragraph",
                "runs": [{
                    "kind": "image",
                    "width": 80.0,
                    "height": 20.0,
                    "rotationBounds": { "width": 20.0, "height": 80.0 }
                }]
            },
            "maxWidth": 200.0,
            "defaults": { "fontSize": 12.0, "fontFamily": "Fallback" },
            "authoritativeShaping": true
        }))
        .unwrap();
        let extent = measure_paragraph(&store, &input).unwrap();
        assert!((extent.lines[0].width - 20.0).abs() < 0.001);
        assert!(extent.lines[0].line_height >= 80.0);
    }
}
