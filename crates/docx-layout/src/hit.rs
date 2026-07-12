use crate::display_list::{DisplayList, HfRegion, Primitive};
use serde::Serialize;

const BAND_SLACK: f64 = 4.0;

/// width of the selection sliver drawn for a blank line (BLANK_LINE_SELECTION_WIDTH_PX)
const BLANK_LINE_SELECTION_WIDTH: f64 = 4.0;

/// parse the px size out of a CSS font shorthand ("700 16px Calibri, ...")
fn font_px(font: &str) -> f64 {
    for token in font.split_whitespace() {
        if let Some(num) = token.strip_suffix("px")
            && let Ok(v) = num.parse::<f64>()
        {
            return v;
        }
    }
    14.666667 // 11pt default
}

/// Positioned text geometry used for hit testing.
struct TextHit<'a> {
    text: &'a str,
    x: f64,
    width: f64,
    top: f64,
    bottom: f64,
    doc_start: i64,
    doc_end: i64,
}

fn text_hits(prims: &[Primitive]) -> Vec<TextHit<'_>> {
    let mut out = Vec::new();
    for p in prims {
        match p {
            Primitive::Text(t) => {
                let (Some(ds), Some(de)) = (t.attrs.doc_start, t.attrs.doc_end) else {
                    continue;
                };
                let fp = font_px(&t.font);
                let baseline = t.baseline_y.as_f64().unwrap_or(0.0);
                out.push(TextHit {
                    text: &t.text,
                    x: t.x.as_f64().unwrap_or(0.0),
                    width: t.width.as_f64().unwrap_or(0.0),
                    top: baseline - fp,
                    bottom: baseline + fp * 0.25,
                    doc_start: ds,
                    doc_end: de,
                });
            }
            Primitive::GlyphRun(g) => {
                let (Some(ds), Some(de)) = (g.attrs.doc_start, g.attrs.doc_end) else {
                    continue;
                };
                if g.glyphs.is_empty() {
                    continue;
                }
                let fp = g.size;
                // Marks sit above the baseline, and the trailing edge includes
                // the final glyph advance.
                let baseline = g
                    .glyphs
                    .iter()
                    .map(|gl| gl.y)
                    .fold(f64::NEG_INFINITY, f64::max);
                let min_x = g.glyphs.iter().map(|gl| gl.x).fold(f64::INFINITY, f64::min);
                let right = g
                    .glyphs
                    .iter()
                    .map(|gl| gl.x + gl.advance)
                    .fold(f64::NEG_INFINITY, f64::max);
                let width = (right - min_x).max(0.0);
                out.push(TextHit {
                    text: &g.text,
                    x: min_x,
                    width,
                    top: baseline - fp,
                    bottom: baseline + fp * 0.25,
                    doc_start: ds,
                    doc_end: de,
                });
            }
            _ => {}
        }
    }
    out
}

fn position_in_run(hit: &TextHit<'_>, x: f64) -> i64 {
    let chars = hit.text.chars().count() as i64;
    let span = hit.doc_end - hit.doc_start;
    let units = if chars > 0 { chars } else { span };
    if units <= 0 || hit.width <= 0.0 {
        return hit.doc_start;
    }
    let ratio = ((x - hit.x) / hit.width).clamp(0.0, 1.0);
    let unit = (ratio * units as f64).round() as i64;
    (hit.doc_start + unit).min(hit.doc_end)
}

pub fn hit_test(dl: &DisplayList, page_index: usize, x: f64, y: f64) -> Option<i64> {
    let page = dl.pages.get(page_index)?;
    resolve_point(&page.primitives, x, y)
}

/// the shared point resolver over one primitive list (a page body or one
/// HF region — both use page coordinates)
fn resolve_point(prims: &[Primitive], x: f64, y: f64) -> Option<i64> {
    let hits = text_hits(prims);

    for h in &hits {
        if h.width <= 0.0 {
            continue;
        }
        if x >= h.x && x <= h.x + h.width && y >= h.top - BAND_SLACK && y <= h.bottom + BAND_SLACK {
            return Some(position_in_run(h, x));
        }
    }

    // 2. direct hit on an image with a doc position (caret parks before it)
    for p in prims {
        let Primitive::Image(img) = p else { continue };
        let Some(ds) = img.attrs.doc_start else {
            continue;
        };
        let (ix, iy) = (img.x.as_f64().unwrap_or(0.0), img.y.as_f64().unwrap_or(0.0));
        let (iw, ih) = (img.w.as_f64().unwrap_or(0.0), img.h.as_f64().unwrap_or(0.0));
        if x >= ix && x <= ix + iw && y >= iy && y <= iy + ih {
            return Some(ds);
        }
    }

    if hits.is_empty() {
        return None;
    }

    // 3. nearest line by vertical center distance (blank-line markers included)
    let mut best_center = f64::INFINITY;
    for h in &hits {
        let center = (h.top + h.bottom) / 2.0;
        let d = (y - center).abs();
        if d < best_center {
            best_center = d;
        }
    }
    // collect the hits on that nearest band (same center within epsilon)
    let mut line: Vec<&TextHit> = Vec::new();
    for h in &hits {
        let center = (h.top + h.bottom) / 2.0;
        if ((y - center).abs() - best_center).abs() < 0.5 {
            line.push(h);
        }
    }

    // 4. nearest primitive on the line; inside -> interpolate, outside -> snap
    // to the closer edge's position
    let mut best: Option<(f64, i64)> = None;
    for h in &line {
        let (dist, pos) = if x < h.x {
            (h.x - x, h.doc_start)
        } else if x > h.x + h.width {
            (x - (h.x + h.width), h.doc_end)
        } else {
            (0.0, position_in_run(h, x))
        };
        if best.is_none() || dist < best.unwrap().0 {
            best = Some((dist, pos));
        }
    }
    best.map(|(_, pos)| pos)
}

/// region-aware hit result: which part of the page owns the point, and the
/// resolved position inside that part's PM doc. For `header`/`footer` the
/// position refers to the HF ProseMirror doc identified by `rId`, NOT the
/// body doc (the caller must route the selection to that HF editor, the way
/// `usePagesPointer` scopes clicks to `.layout-page-header|footer`).
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RegionHit {
    pub region: HitRegion,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r_id: Option<String>,
    pub pos: Option<i64>,
}

#[derive(Serialize, Debug, Clone, Copy, PartialEq, Eq)]
pub enum HitRegion {
    #[serde(rename = "body")]
    Body,
    #[serde(rename = "header")]
    Header,
    #[serde(rename = "footer")]
    Footer,
}

pub(crate) fn parse_region(region: &str) -> Result<HitRegion, String> {
    match region {
        "body" => Ok(HitRegion::Body),
        "header" => Ok(HitRegion::Header),
        "footer" => Ok(HitRegion::Footer),
        other => Err(format!("unknown region {other:?}")),
    }
}

/// Resolve a point against header, footer, and body regions.
pub fn hit_test_regions(dl: &DisplayList, page_index: usize, x: f64, y: f64) -> Option<RegionHit> {
    let page = dl.pages.get(page_index)?;

    let in_band = |r: &HfRegion| -> bool {
        let top = r.y.as_f64().unwrap_or(0.0);
        let bottom = top + r.height.as_f64().unwrap_or(0.0);
        y >= top && y <= bottom
    };

    if let Some(h) = &page.header
        && in_band(h)
    {
        return Some(RegionHit {
            region: HitRegion::Header,
            r_id: Some(h.r_id.clone()),
            pos: resolve_point(&h.primitives, x, y),
        });
    }
    if let Some(f) = &page.footer
        && in_band(f)
    {
        return Some(RegionHit {
            region: HitRegion::Footer,
            r_id: Some(f.r_id.clone()),
            pos: resolve_point(&f.primitives, x, y),
        });
    }
    Some(RegionHit {
        region: HitRegion::Body,
        r_id: None,
        pos: resolve_point(&page.primitives, x, y),
    })
}

/// one highlight rectangle of a PM range, page-local coordinates
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RangeRect {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

/// append the highlight rects covering `[from, to)` over ONE primitive list (a
/// page body or one HF band — both page-local) to `out`, stamping `page_index`
/// on each. Shared by the body-only [`range_rects`] and the region-aware
/// [`range_rects_in_region`]: per overlapped text primitive a proportional
/// sub-span rect over the line band, a 4px sliver for blank-line markers, and
/// the full box for images inside the range.
fn collect_range_rects(
    prims: &[Primitive],
    page_index: usize,
    from: i64,
    to: i64,
    out: &mut Vec<RangeRect>,
) {
    for h in text_hits(prims) {
        // blank-line marker: zero-length span selects as a thin sliver
        if h.doc_start == h.doc_end {
            if h.doc_start >= from && h.doc_start < to {
                out.push(RangeRect {
                    page_index,
                    x: h.x,
                    y: h.top,
                    width: BLANK_LINE_SELECTION_WIDTH,
                    height: h.bottom - h.top,
                });
            }
            continue;
        }
        if h.doc_end <= from || h.doc_start >= to {
            continue;
        }
        let span = (h.doc_end - h.doc_start) as f64;
        let start_ratio = ((from - h.doc_start).max(0) as f64 / span).clamp(0.0, 1.0);
        let end_ratio =
            ((to - h.doc_start).min(h.doc_end - h.doc_start) as f64 / span).clamp(0.0, 1.0);
        let x0 = h.x + h.width * start_ratio;
        let x1 = h.x + h.width * end_ratio;
        out.push(RangeRect {
            page_index,
            x: x0,
            y: h.top,
            // degenerate overlaps keep a 1px floor like lineSpanRect
            width: (x1 - x0).max(1.0),
            height: h.bottom - h.top,
        });
    }

    for p in prims {
        let Primitive::Image(img) = p else { continue };
        let (Some(ds), Some(de)) = (img.attrs.doc_start, img.attrs.doc_end) else {
            continue;
        };
        if de <= from || ds >= to {
            continue;
        }
        out.push(RangeRect {
            page_index,
            x: img.x.as_f64().unwrap_or(0.0),
            y: img.y.as_f64().unwrap_or(0.0),
            width: img.w.as_f64().unwrap_or(0.0),
            height: img.h.as_f64().unwrap_or(0.0),
        });
    }
}

pub fn range_rects(dl: &DisplayList, from: i64, to: i64) -> Vec<RangeRect> {
    range_rects_in_region(dl, HitRegion::Body, None, from, to)
}

pub fn range_rects_in_region(
    dl: &DisplayList,
    region: HitRegion,
    r_id: Option<&str>,
    from: i64,
    to: i64,
) -> Vec<RangeRect> {
    let (from, to) = (from.min(to), from.max(to));
    let mut rects = Vec::new();
    if from == to {
        return rects;
    }

    let matches = |band: &HfRegion| r_id.is_none_or(|id| id == band.r_id);
    for (page_index, page) in dl.pages.iter().enumerate() {
        let prims: Option<&[Primitive]> = match region {
            HitRegion::Body => Some(&page.primitives),
            HitRegion::Header => page
                .header
                .as_ref()
                .filter(|h| matches(h))
                .map(|h| h.primitives.as_slice()),
            HitRegion::Footer => page
                .footer
                .as_ref()
                .filter(|f| matches(f))
                .map(|f| f.primitives.as_slice()),
        };
        if let Some(prims) = prims {
            collect_range_rects(prims, page_index, from, to, &mut rects);
        }
    }

    rects
}

/// `hit_test` over serialized inputs; returns `"null"` or the position as JSON
pub fn hit_test_json(
    display_list: &str,
    page_index: usize,
    x: f64,
    y: f64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    match hit_test(&dl, page_index, x, y) {
        Some(pos) => Ok(pos.to_string()),
        None => Ok("null".to_string()),
    }
}

/// `range_rects` over serialized inputs; returns a JSON array of rects
pub fn range_rects_json(display_list: &str, from: i64, to: i64) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    serde_json::to_string(&range_rects(&dl, from, to)).map_err(|e| format!("serialize: {e}"))
}

/// `range_rects_in_region` over serialized inputs. `region` is
/// `"body" | "header" | "footer"`; `r_id` is the HF part's relationship id
/// (ignored for `body`; empty string ⇒ match any band of the kind). Returns a
/// JSON array of rects, or a `parse:`/`unknown region` error string.
pub fn range_rects_region_json(
    display_list: &str,
    region: &str,
    r_id: &str,
    from: i64,
    to: i64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    let region = parse_region(region)?;
    let r_id = if r_id.is_empty() { None } else { Some(r_id) };
    serde_json::to_string(&range_rects_in_region(&dl, region, r_id, from, to))
        .map_err(|e| format!("serialize: {e}"))
}

/// `hit_test_regions` over serialized inputs; returns
/// `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` or `"null"`
/// for an out-of-range page
pub fn hit_test_regions_json(
    display_list: &str,
    page_index: usize,
    x: f64,
    y: f64,
) -> Result<String, String> {
    let dl: DisplayList = serde_json::from_str(display_list).map_err(|e| format!("parse: {e}"))?;
    match hit_test_regions(&dl, page_index, x, y) {
        Some(hit) => serde_json::to_string(&hit).map_err(|e| format!("serialize: {e}")),
        None => Ok("null".to_string()),
    }
}
