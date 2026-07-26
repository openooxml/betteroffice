//! Header/footer band composition for the display list.
//!
//! A page gains `header` / `footer` regions only when the build envelope
//! carries a `headersFooters` payload. Band content reuses the body paragraph
//! and table emitters, so runs, decorations and PAGE/NUMPAGES fields behave the
//! same as in the body — but the `docStart` / `docEnd` on those primitives
//! address the header/footer document named by the region's `rId`, never the
//! body document. Hit testing and selection must therefore scope by region.
//!
//! # Envelope
//!
//! The payload carries the section flags `titlePg` (`w:titlePg`) and
//! `evenAndOddHeaders` (`w:evenAndOddHeaders`), each optionally narrowed to
//! individual sections via `titlePageSections` / `evenAndOddSections`; optional
//! `headerDistance` / `footerDistance` overrides (`w:pgMar` `w:header` /
//! `w:footer`); an optional watermark; and one `variants` entry per part in
//! play. A variant names its `rId`, `kind`, `type`, optional `sectionIndex`,
//! its own `measured` blocks in the body schema, its heights, and optional
//! per-page `fieldWidths`.
//!
//! # Variant selection
//!
//! A page's own `headerFooterRefs` wins when present: the relationship id for
//! the selected type resolves the variant directly, and a `first` selection
//! that names no relationship leaves the band blank. Otherwise, by 1-based page
//! number:
//!
//! 1. the section's first page (`sectionPageIndex == 0`) under `titlePg`
//!    selects `first`. An absent `first` variant means a deliberately blank
//!    band; it must not fall through to `default`.
//! 2. an even page number under `evenAndOddHeaders` selects `even` when that
//!    variant exists;
//! 3. otherwise `default`.
//!
//! Where several variants match a (kind, type, section), the last one wins.
//!
//! # Band geometry
//!
//! The distance resolves as the envelope override, then the page's
//! `margins.header` / `margins.footer`, then [`DEFAULT_HF_DISTANCE_PX`]. Both
//! kinds take the interactive height `max(flowHeight - min(0, visualTop), 24)`.
//! A header band sits at `distance + visualTop` and flows content from
//! `distance`. A footer band is bottom-anchored: it sits at
//! `pageHeight - distance - bandHeight`, flows content from
//! `pageHeight - distance - max(visualBottom - visualTop, 24) - visualTop`, and
//! anchors floating tables against `pageHeight - distance - height`. Content
//! starts at `margins.left` horizontally in both cases.
//!
//! # Stacking inside a band
//!
//! A band-local cursor starts at zero. A paragraph paints at
//! `cursor + spacing.before` as a single unsplit fragment and advances the
//! cursor by its measured total height, which already accounts for its own
//! spacing. An inline table paints whole at the cursor and advances by its
//! total height; a `w:tblpPr` floating table paints at its resolved anchor and
//! does not advance. An image paints at the cursor and advances by its measured
//! height. Every other block kind emits nothing and leaves the cursor alone.

use serde::Deserialize;

use crate::display_list::{
    BlockIn, BlockRef, FieldWidthEntry, FieldWidthMap, FloatingTablePositionIn, HfKind, HfRegion,
    MeasureIn, MeasuredBlockIn, PageIn, ParagraphFragmentIn, Primitive, RenderCtx, ShapeFonts,
    TableFragmentIn, WatermarkIn, capped_alt_text, emit_paragraph_fragment, emit_table_fragment,
    px, rotation_degrees, sanitized_href, table_total_width,
};
use crate::display_list::{Crop, ImagePrimitive};

/// Word's default header/footer distance: 0.5 inches at 96 DPI.
const DEFAULT_HF_DISTANCE_PX: f64 = 48.0;

/// Minimum interactive band height, so a near-empty band stays clickable.
const MIN_BAND_HEIGHT_PX: f64 = 24.0;

/// The optional `headersFooters` envelope field.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HeadersFootersIn {
    #[serde(default)]
    title_pg: Option<bool>,
    #[serde(default)]
    even_and_odd_headers: Option<bool>,
    #[serde(default)]
    title_page_sections: Vec<usize>,
    #[serde(default)]
    even_and_odd_sections: Vec<usize>,
    #[serde(default)]
    header_distance: Option<f64>,
    #[serde(default)]
    footer_distance: Option<f64>,
    #[serde(default)]
    pub(crate) watermark: Option<WatermarkIn>,
    #[serde(default)]
    variants: Vec<HfVariantIn>,
}

/// One header/footer part, with the four heights the band geometry needs.
/// `visualTop` and `visualBottom` describe the painted extent, which may sit
/// outside the in-flow stack when content is negatively offset.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct HfVariantIn {
    r_id: String,
    kind: HfKind,
    #[serde(rename = "type", default)]
    hf_type: HfType,
    #[serde(default)]
    section_index: Option<usize>,
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
    /// Per-page resolved PAGE/NUMPAGES field widths.
    #[serde(default)]
    field_widths: Vec<FieldWidthsIn>,
}

/// Per-field-run header/footer widths.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct FieldWidthsIn {
    /// field run's document position in this HF doc — the key the builder matches
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

/// Selects the header or footer variant for a 1-based page number.
fn resolve_variant<'a>(
    hf: &'a HeadersFootersIn,
    page: &PageIn,
    kind: HfKind,
    page_number: u64,
) -> Option<&'a HfVariantIn> {
    let section_index = page.section_index.map(|value| value as usize);
    // Later variants take precedence.
    let get = |t: HfType| {
        hf.variants.iter().rfind(|v| {
            v.kind == kind
                && v.hf_type == t
                && (v.section_index.is_none() || v.section_index == section_index)
        })
    };
    let title_page = section_index.is_some_and(|index| hf.title_page_sections.contains(&index))
        || hf.title_pg == Some(true);
    let even_and_odd = section_index.is_some_and(|index| hf.even_and_odd_sections.contains(&index))
        || hf.even_and_odd_headers == Some(true);
    let first_page = page
        .section_page_index
        .unwrap_or(page_number.saturating_sub(1))
        == 0;
    let selected_type = if first_page && title_page {
        HfType::First
    } else if even_and_odd && page_number.is_multiple_of(2) {
        HfType::Even
    } else {
        HfType::Default
    };
    if let Some(refs) = &page.header_footer_refs {
        let r_id = match (kind, selected_type) {
            (HfKind::Header, HfType::Default) => refs.header_default.as_deref(),
            (HfKind::Header, HfType::First) => refs.header_first.as_deref(),
            (HfKind::Header, HfType::Even) => refs.header_even.as_deref(),
            (HfKind::Footer, HfType::Default) => refs.footer_default.as_deref(),
            (HfKind::Footer, HfType::First) => refs.footer_first.as_deref(),
            (HfKind::Footer, HfType::Even) => refs.footer_even.as_deref(),
        };
        if let Some(r_id) = r_id
            && let Some(variant) = hf.variants.iter().rfind(|variant| {
                variant.kind == kind
                    && variant.r_id == r_id
                    && (variant.section_index.is_none() || variant.section_index == section_index)
            })
        {
            return Some(variant);
        }
        if selected_type == HfType::First {
            return None;
        }
    }
    if first_page && title_page {
        // `titlePg` selects a distinct story. Word treats an absent first-page
        // relationship as an intentionally blank band; falling through here
        // would incorrectly repeat the default header/footer on page one.
        return get(HfType::First);
    }
    if even_and_odd
        && page_number.is_multiple_of(2)
        && let Some(v) = get(HfType::Even)
    {
        return Some(v);
    }
    get(HfType::Default)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn variant(kind: HfKind, hf_type: HfType, r_id: &str) -> HfVariantIn {
        HfVariantIn {
            r_id: r_id.to_owned(),
            kind,
            hf_type,
            section_index: None,
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
            title_page_sections: Vec::new(),
            even_and_odd_sections: Vec::new(),
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

        let page = PageIn::default();
        assert!(resolve_variant(&hf, &page, HfKind::Header, 1).is_none());
        assert_eq!(
            resolve_variant(&hf, &page, HfKind::Header, 2).map(|v| v.r_id.as_str()),
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
            resolve_variant(&hf, &PageIn::default(), HfKind::Header, 1).map(|v| v.r_id.as_str()),
            Some("first-header")
        );
    }

    #[test]
    fn page_relationship_selects_the_matching_section_variant() {
        let mut first = variant(HfKind::Header, HfType::Default, "section-zero");
        first.section_index = Some(0);
        let mut second = variant(HfKind::Header, HfType::Default, "section-one");
        second.section_index = Some(1);
        let hf = envelope(false, vec![first, second]);
        let mut page = PageIn::default();
        page.section_index = Some(1);
        page.header_footer_refs = Some(crate::display_list::PageHeaderFooterRefsIn {
            header_default: Some("section-one".to_owned()),
            ..Default::default()
        });

        assert_eq!(
            resolve_variant(&hf, &page, HfKind::Header, 2).map(|value| value.r_id.as_str()),
            Some("section-one")
        );
    }
}

/// Computes stacked in-flow height when the payload omits it.
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

/// Resolves both bands for one page.
///
/// `page_number` is the layout's own 1-based `Page.number` (falling back to
/// `page_index + 1`) and drives both variant selection and PAGE field text, so
/// it restarts wherever a section restarts numbering.
pub(crate) fn compose_page_regions<'a>(
    hf: &HeadersFootersIn,
    page: &PageIn,
    page_index: usize,
    total_pages: u64,
    shape: Option<&'a ShapeFonts<'a>>,
) -> (Option<HfRegion>, Option<HfRegion>) {
    let page_number = page.number.unwrap_or(page_index as u64 + 1);
    let header = resolve_variant(hf, page, HfKind::Header, page_number).map(|v| {
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
    let footer = resolve_variant(hf, page, HfKind::Footer, page_number).map(|v| {
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

/// Returns per-page PAGE/NUMPAGES widths when supplied.
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

/// Resolves a `w:tblpPr` anchor into offsets from the band's flow origin.
///
/// `page` and `margin` anchors are expressed against the page or the body
/// margin box, so both are rebased onto the band; any other anchor is already
/// band-relative. The caller adds the returned offsets to the region origin.
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

/// Builds one band: geometry first, then the block stack, per the module rules.
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
    // Field widths are scoped to one header/footer variant.
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
                    // HF paragraphs surface as mirror paragraph wrappers inside
                    // the header/footer region — stamp the line range
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
                    width: table_total_width(measure),
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
