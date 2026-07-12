use serde::Deserialize;

use crate::display_list::{
    BlockIn, BlockRef, FieldWidthEntry, FieldWidthMap, FloatingTablePositionIn, HfKind, HfRegion,
    MeasureIn, MeasuredBlockIn, PageIn, ParagraphFragmentIn, Primitive, RenderCtx, ShapeFonts,
    TableFragmentIn, WatermarkIn, capped_alt_text, emit_paragraph_fragment, emit_table_fragment,
    px, rotation_degrees, sanitized_href,
};
use crate::display_list::{Crop, ImagePrimitive};

const DEFAULT_HF_DISTANCE_PX: f64 = 48.0;

const MIN_BAND_HEIGHT_PX: f64 = 24.0;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HeadersFootersIn {
    #[serde(default)]
    title_pg: Option<bool>,
    #[serde(default)]
    even_and_odd_headers: Option<bool>,
    #[serde(default)]
    header_distance: Option<f64>,
    #[serde(default)]
    footer_distance: Option<f64>,
    #[serde(default)]
    pub(crate) watermark: Option<WatermarkIn>,
    #[serde(default)]
    variants: Vec<HfVariantIn>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct HfVariantIn {
    r_id: String,
    kind: HfKind,
    #[serde(rename = "type", default)]
    hf_type: HfType,
    #[serde(default)]
    measured: Vec<MeasuredBlockIn>,
    /// HeaderFooterContent.height (total in-flow stack)
    #[serde(default)]
    height: Option<f64>,
    /// HeaderFooterContent.flowHeight (in-flow band height, excludes floats)
    #[serde(default)]
    flow_height: Option<f64>,
    #[serde(default)]
    visual_top: Option<f64>,
    #[serde(default)]
    visual_bottom: Option<f64>,
    /// Per-page widths for dynamic page-number fields.
    #[serde(default)]
    field_widths: Vec<FieldWidthsIn>,
}

/// Per-field widths for one header or footer variant.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct FieldWidthsIn {
    /// field run's pm position in this HF doc — the key the builder matches
    pm_start: i64,
    /// width the measure baked into `line.width` (the field's fallback text)
    fallback_width: f64,
    /// resolved-text width per layout page index
    #[serde(default)]
    per_page: Vec<f64>,
}

#[derive(Deserialize, Default, Clone, Copy, PartialEq, Eq)]
enum HfType {
    #[serde(rename = "default")]
    #[default]
    Default,
    #[serde(rename = "first")]
    First,
    #[serde(rename = "even")]
    Even,
}

fn resolve_variant(hf: &HeadersFootersIn, kind: HfKind, page_number: u64) -> Option<&HfVariantIn> {
    let get = |t: HfType| {
        hf.variants
            .iter()
            .rfind(|v| v.kind == kind && v.hf_type == t)
    };
    if page_number == 1 && hf.title_pg == Some(true) {
        // `titlePg` selects a distinct story. Word treats an absent first-page
        // relationship as an intentionally blank band; falling through here
        // would incorrectly repeat the default header/footer on page one.
        return get(HfType::First);
    }
    if hf.even_and_odd_headers == Some(true)
        && page_number.is_multiple_of(2)
        && let Some(v) = get(HfType::Even)
    {
        return Some(v);
    }
    get(HfType::Default)
}

/// Sum measured block heights.
fn stacked_height(measured: &[MeasuredBlockIn]) -> f64 {
    measured
        .iter()
        .map(|mb| match &mb.measure {
            MeasureIn::Paragraph(p) => p.total_height,
            MeasureIn::Table(t) => t.total_height,
            MeasureIn::Image(i) => i.height,
            MeasureIn::TextBox(t) => t.height,
            MeasureIn::Shape(s) => s.height,
            MeasureIn::Chart(c) => c.height,
            MeasureIn::Unsupported => 0.0,
        })
        .sum()
}

pub(crate) fn compose_page_regions<'a>(
    hf: &HeadersFootersIn,
    page: &PageIn,
    page_index: usize,
    total_pages: u64,
    shape: Option<&'a ShapeFonts<'a>>,
) -> (Option<HfRegion>, Option<HfRegion>) {
    let page_number = page.number.unwrap_or(page_index as u64 + 1);
    let header = resolve_variant(hf, HfKind::Header, page_number).map(|v| {
        compose_region(
            v,
            HfKind::Header,
            hf,
            page,
            page_index,
            page_number,
            total_pages,
            shape,
        )
    });
    let footer = resolve_variant(hf, HfKind::Footer, page_number).map(|v| {
        compose_region(
            v,
            HfKind::Footer,
            hf,
            page,
            page_index,
            page_number,
            total_pages,
            shape,
        )
    });
    (header, footer)
}

/// Build the dynamic field-width map for one variant.
fn field_width_map(v: &HfVariantIn) -> Option<FieldWidthMap> {
    if v.field_widths.is_empty() {
        return None;
    }
    Some(
        v.field_widths
            .iter()
            .map(|fw| {
                (
                    fw.pm_start,
                    FieldWidthEntry {
                        fallback: fw.fallback_width,
                        per_page: fw.per_page.clone(),
                    },
                )
            })
            .collect(),
    )
}

fn resolve_hf_floating_table_position(
    floating: &FloatingTablePositionIn,
    page: &PageIn,
    flow_top: f64,
    flow_left: f64,
) -> (f64, f64) {
    let mut top = floating.tblp_y.unwrap_or(0.0);
    match floating.vert_anchor.as_deref() {
        Some("page") => top -= flow_top,
        Some("margin") => top += page.margins.top - flow_top,
        _ => {}
    }

    let mut left = floating.tblp_x.unwrap_or(0.0);
    match floating.horz_anchor.as_deref() {
        Some("page") => left -= flow_left,
        Some("margin") => left += page.margins.left - flow_left,
        _ => {}
    }

    (left, top)
}

#[allow(clippy::too_many_arguments)]
fn compose_region(
    v: &HfVariantIn,
    kind: HfKind,
    hf: &HeadersFootersIn,
    page: &PageIn,
    page_index: usize,
    page_number: u64,
    total_pages: u64,
    shape: Option<&ShapeFonts<'_>>,
) -> HfRegion {
    let field_widths = field_width_map(v);
    let ctx = RenderCtx {
        page_number,
        page_index,
        total_pages,
        shape,
        field_widths: field_widths.as_ref(),
    };
    let ctx = &ctx;
    let content_width = page.size.w - page.margins.left - page.margins.right;
    let height = v.height.unwrap_or_else(|| stacked_height(&v.measured));
    let visual_top = v.visual_top.unwrap_or(0.0);
    let visual_bottom = v.visual_bottom.unwrap_or(height);
    let flow_height = v.flow_height.unwrap_or(height);
    let interactive = (flow_height - visual_top.min(0.0)).max(MIN_BAND_HEIGHT_PX);

    let (band_y, band_h, origin_y, flow_top) = match kind {
        HfKind::Header => {
            let distance = hf
                .header_distance
                .or(page.margins.header)
                .unwrap_or(DEFAULT_HF_DISTANCE_PX);
            (distance + visual_top, interactive, distance, distance)
        }
        HfKind::Footer => {
            let distance = hf
                .footer_distance
                .or(page.margins.footer)
                .unwrap_or(DEFAULT_HF_DISTANCE_PX);
            let actual = (visual_bottom - visual_top).max(MIN_BAND_HEIGHT_PX);
            let band_y = page.size.h - distance - interactive;
            let origin_y = page.size.h - distance - actual - visual_top;
            let flow_top = page.size.h - distance - height;
            (band_y, interactive, origin_y, flow_top)
        }
    };

    let mut prims: Vec<Primitive> = Vec::new();
    let origin_x = page.margins.left;
    let mut cursor = 0.0_f64;

    for mb in &v.measured {
        match (&mb.block, &mb.measure) {
            (BlockIn::Paragraph(block), MeasureIn::Paragraph(measure)) => {
                let spacing_before = block
                    .attrs
                    .as_ref()
                    .and_then(|a| a.spacing)
                    .and_then(|s| s.before)
                    .unwrap_or(0.0);
                let frag = ParagraphFragmentIn {
                    block_id: block.id.clone(),
                    x: origin_x,
                    y: origin_y + cursor + spacing_before,
                    width: content_width,
                    height: measure.total_height,
                    from_line: 0,
                    to_line: measure.lines.len(),
                    pm_start: block.pm_start,
                    pm_end: block.pm_end,
                    carried_from_prev: None,
                    carried_to_next: None,
                };
                emit_paragraph_fragment(
                    &mut prims, &frag, block, measure, ctx, frag.x, frag.y, None, None, true, true,
                );
                cursor += measure.total_height;
            }
            (BlockIn::Table(block), MeasureIn::Table(measure)) => {
                let (x, y, advance_cursor) = if let Some(floating) = block.floating.as_ref() {
                    let (left, top) =
                        resolve_hf_floating_table_position(floating, page, flow_top, origin_x);
                    (origin_x + left, origin_y + top, false)
                } else {
                    (origin_x, origin_y + cursor, true)
                };
                let frag = TableFragmentIn {
                    block_id: block.id.clone(),
                    x,
                    y,
                    height: measure.total_height,
                    row_start: 0,
                    row_end: block.rows.len(),
                    clip_top: None,
                    clip_bottom: None,
                    header_row_count: None,
                    carried_from_prev: None,
                    carried_to_next: None,
                };
                emit_table_fragment(&mut prims, &frag, block, measure, ctx);
                if advance_cursor {
                    cursor += measure.total_height;
                }
            }
            (BlockIn::Image(block), MeasureIn::Image(measure)) => {
                let rot = rotation_degrees(block.transform.as_deref());
                let mut attrs = BlockRef::of(&block.id).attrs();
                attrs.doc_start = block.pm_start;
                attrs.doc_end = block.pm_end;
                attrs.href = sanitized_href(block.hlink_href.as_deref());
                attrs.sdt = crate::display_list::sdt_attrs_from_groups(&block.sdt_groups);
                prims.push(Primitive::Image(ImagePrimitive {
                    rel_id: block.src.clone(),
                    x: px(origin_x),
                    y: px(origin_y + cursor),
                    w: px(measure.width),
                    h: px(measure.height),
                    rotation_deg: if rot != 0.0 { Some(px(rot)) } else { None },
                    opacity: None,
                    filter: None,
                    decorative: false,
                    crop: None::<Crop>,
                    alt_text: capped_alt_text(block.alt.as_deref()),
                    attrs,
                }));
                cursor += measure.height;
            }
            _ => {}
        }
    }

    HfRegion {
        r_id: v.r_id.clone(),
        kind,
        y: px(band_y),
        height: px(band_h),
        primitives: prims,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn variant(kind: HfKind, hf_type: HfType, r_id: &str) -> HfVariantIn {
        HfVariantIn {
            r_id: r_id.to_owned(),
            kind,
            hf_type,
            measured: Vec::new(),
            height: None,
            flow_height: None,
            visual_top: None,
            visual_bottom: None,
            field_widths: Vec::new(),
        }
    }

    fn envelope(title_pg: bool, variants: Vec<HfVariantIn>) -> HeadersFootersIn {
        HeadersFootersIn {
            title_pg: Some(title_pg),
            even_and_odd_headers: None,
            header_distance: None,
            footer_distance: None,
            watermark: None,
            variants,
        }
    }

    #[test]
    fn title_page_without_first_variant_is_blank() {
        let hf = envelope(
            true,
            vec![variant(HfKind::Header, HfType::Default, "default-header")],
        );

        assert!(resolve_variant(&hf, HfKind::Header, 1).is_none());
        assert_eq!(
            resolve_variant(&hf, HfKind::Header, 2).map(|v| v.r_id.as_str()),
            Some("default-header")
        );
    }

    #[test]
    fn title_page_uses_first_variant_when_present() {
        let hf = envelope(
            true,
            vec![
                variant(HfKind::Header, HfType::Default, "default-header"),
                variant(HfKind::Header, HfType::First, "first-header"),
            ],
        );

        assert_eq!(
            resolve_variant(&hf, HfKind::Header, 1).map(|v| v.r_id.as_str()),
            Some("first-header")
        );
    }
}
