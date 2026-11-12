//! Text painting and per-character font fallback for PDF export.
//!
//! `TextRunPrimitive`s carry Unicode text but no glyph ids, so this module
//! resolves every scalar against the font registry that produced the display
//! list and paints the resulting CIDs with embedded subsets. `GlyphRun`s
//! already carry ids; they are painted as positioned CIDs and may be rotated.

use std::collections::BTreeMap;

use docx_layout::display_list::{DisplayPage, GlyphRunPrimitive, PlacedGlyph, TextRunPrimitive};
use ooxml_text::{FontId, FontStore};
use pdf_writer::{Content, Name, Str, TextStr};
use serde_json::Number;
use ttf_parser::{Face, GlyphId};

use crate::PdfError;
use crate::font::EmbeddedFont;
use crate::units::{PX_TO_PT, f64_to_points, font_size_px, parse_color, points};

/// Base-14 resource registered by the writer for text without embedded fonts.
const BASE_FONT: Name<'static> = Name(b"F0");
/// One parsed face per font id, in registry order.
///
/// Painting and collection must agree on which font supplies each scalar, so
/// both go through [`FontFaces::resolve`]. Faces whose declared family matches
/// the run's CSS family are tried before every other face.
struct FontFaces<'a> {
    faces: Vec<Option<Face<'a>>>,
    /// Registry indices of faces declaring the run's requested family.
    preferred: Vec<u32>,
}

impl<'a> FontFaces<'a> {
    fn new(store: &'a FontStore) -> Self {
        Self::with_family(store, None)
    }

    fn with_family(store: &'a FontStore, family: Option<&str>) -> Self {
        let mut faces = Vec::with_capacity(store.font_count());
        for index in 0..store.font_count() {
            let id = FontId::from_u32(index as u32);
            let face = store
                .font_bytes(id)
                .ok()
                .and_then(|bytes| Face::parse(bytes, 0).ok());
            faces.push(face);
        }
        let preferred = match family.map(str::to_lowercase) {
            Some(requested) => (0..faces.len())
                .filter(|&index| faces[index].is_some())
                .filter(|&index| {
                    store
                        .declared_families(FontId::from_u32(index as u32))
                        .is_ok_and(|declared| {
                            declared
                                .iter()
                                .any(|name| name.trim().to_lowercase() == requested)
                        })
                })
                .map(|index| index as u32)
                .collect(),
            None => Vec::new(),
        };
        Self { faces, preferred }
    }

    fn is_empty(&self) -> bool {
        self.faces.iter().all(Option::is_none)
    }

    /// First font whose cmap maps `ch` to a real glyph, preferring a face that
    /// declares the run's requested family.
    ///
    /// Uncovered scalars fall back to the first parseable face's `.notdef`
    /// (glyph 0), which every subset keeps.
    fn resolve(&self, ch: char) -> Option<(u32, u16)> {
        for &index in &self.preferred {
            if let Some(glyph) = self.covering(index as usize, ch) {
                return Some((index, glyph));
            }
        }
        for index in 0..self.faces.len() {
            if let Some(glyph) = self.covering(index, ch) {
                return Some((index as u32, glyph));
            }
        }
        self.first_face().map(|index| (index, 0))
    }

    fn covering(&self, index: usize, ch: char) -> Option<u16> {
        let glyph = self.faces.get(index)?.as_ref()?.glyph_index(ch)?;
        (glyph.0 != 0).then_some(glyph.0)
    }

    fn first_face(&self) -> Option<u32> {
        // Prefer a TrueType face for `.notdef` fallback: its CIDToGIDMap can
        // duplicate one empty glyph into several searchable CIDs, while CFF
        // CIDs cannot be duplicated. The requested family only breaks ties
        // between TrueType faces: a preferred CFF face must not displace a
        // registry TrueType face, or extraction loses every duplicate.
        let truetype = |index: usize| {
            self.faces
                .get(index)
                .and_then(Option::as_ref)
                .is_some_and(|face| face.tables().cff.is_none())
        };
        self.preferred
            .iter()
            .copied()
            .find(|&index| truetype(index as usize))
            .or_else(|| {
                (0..self.faces.len())
                    .find(|&index| truetype(index))
                    .map(|index| index as u32)
            })
            .or_else(|| {
                self.preferred
                    .iter()
                    .copied()
                    .find(|&index| self.faces[index as usize].is_some())
            })
            .or_else(|| {
                self.faces
                    .iter()
                    .position(Option::is_some)
                    .map(|index| index as u32)
            })
    }

    /// Character advance in CSS pixels: `glyph_hor_advance * size / upem`.
    fn advance_px(&self, font_id: u32, glyph: u16, size_px: f64) -> f64 {
        let face = self
            .faces
            .get(font_id as usize)
            .and_then(Option::as_ref)
            .or_else(|| self.faces.iter().find_map(Option::as_ref));
        let Some(face) = face else {
            return size_px * 0.5;
        };
        let upem = face.units_per_em().max(1);
        let advance = face.glyph_hor_advance(GlyphId(glyph)).unwrap_or(upem / 2);
        f64::from(advance) * size_px / f64::from(upem)
    }
}

/// Record the glyphs a fallback text run needs, keyed by font id and glyph id.
///
/// Scalar resolution follows [`FontFaces::resolve`], so the painter can rely on
/// every selected glyph being present in the subset. Casing (all caps and
/// synthesized small caps) is applied here exactly as [`paint_text`] applies
/// it, including the uppercased glyphs small caps draw.
pub(crate) fn store_has_usable_face(store: &FontStore) -> bool {
    !FontFaces::new(store).is_empty()
}

pub(crate) fn display_text(run: &TextRunPrimitive) -> String {
    if run.all_caps {
        run.text.to_uppercase()
    } else {
        run.text.clone()
    }
}

/// Primary family of a CSS font shorthand such as
/// `italic small-caps 700 16px "Times New Roman", serif`.
fn css_family(css: &str) -> Option<String> {
    let mut cursor = 0;
    let mut family_start = None;
    for token in css.split_whitespace() {
        let start = cursor + css.get(cursor..)?.find(token)?;
        cursor = start + token.len();
        if token.ends_with("px") {
            family_start = Some(cursor);
            break;
        }
    }
    let families = css.get(family_start?..)?.trim();
    let family = docx_layout::measure_blocks::css_primary_family(families);
    (!family.is_empty()).then(|| family.to_lowercase())
}

/// Whether the shorthand asks for `font-variant: small-caps`.
fn css_small_caps(css: &str) -> bool {
    css.split_whitespace()
        .take_while(|token| !token.ends_with("px"))
        .any(|token| token == "small-caps")
}

/// Small caps are synthesized only when the run is not already all caps.
fn small_caps(run: &TextRunPrimitive) -> bool {
    (run.small_caps || css_small_caps(&run.font)) && !run.all_caps
}

/// One source scalar as painted, after all-caps / synthesized small-caps
/// casing.
struct PaintedChar {
    /// Display scalars in paint order. One source scalar can uppercase into
    /// several (`ß` → `SS`).
    chars: Vec<char>,
    /// Font-size and advance multiplier: `0.7` for a synthesized small cap.
    scale: f64,
}

/// Casing shared by collection and painting. Mirrors the raster's
/// `paint_cased_text`: lowercase scalars become uppercase glyphs at
/// [`ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE`] of the run size.
fn painted_chars(run: &TextRunPrimitive) -> Vec<PaintedChar> {
    let small_caps = small_caps(run);
    let mut painted = Vec::with_capacity(run.text.len());
    for ch in run.text.chars() {
        if run.all_caps {
            painted.push(PaintedChar {
                chars: ch.to_uppercase().collect(),
                scale: 1.0,
            });
        } else if small_caps && ch.is_lowercase() {
            painted.push(PaintedChar {
                chars: ooxml_text::uppercase_for_language(ch, run.attrs.lang.as_deref()),
                scale: f64::from(ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE),
            });
        } else {
            painted.push(PaintedChar {
                chars: vec![ch],
                scale: 1.0,
            });
        }
    }
    painted
}

pub(crate) fn collect_text_glyphs(
    store: &FontStore,
    run: &TextRunPrimitive,
    out: &mut BTreeMap<u32, BTreeMap<u16, std::collections::BTreeSet<String>>>,
) {
    if run.hidden || run.text.is_empty() {
        return;
    }
    let faces = FontFaces::with_family(store, css_family(&run.font).as_deref());
    for painted in painted_chars(run) {
        for ch in painted.chars {
            let Some((font_id, glyph)) = faces.resolve(ch) else {
                continue;
            };
            let entry = out.entry(font_id).or_default();
            // Even `.notdef` keeps the source scalar: the text layer must stay
            // searchable when no registered face covers the glyph.
            entry.entry(glyph).or_default().insert(ch.to_string());
        }
    }
}

/// Collect fallback glyphs for `.notdef` clusters in a glyph run.
///
/// A shaped glyph run is single-font, but its font may lack a script entirely
/// (emoji shaped with a Latin run font is the common case). Each gid-0 cluster
/// is re-resolved through the registry so the painter can emit the fallback
/// font's glyph at the same position instead of collapsing extraction to the
/// run's `.notdef`.
pub(crate) fn collect_glyph_run_fallbacks(
    store: &FontStore,
    run: &GlyphRunPrimitive,
    out: &mut BTreeMap<u32, BTreeMap<u16, std::collections::BTreeSet<String>>>,
) {
    if !run.glyphs.iter().any(|glyph| glyph.id == 0) {
        return;
    }
    let faces = FontFaces::new(store);
    if faces.is_empty() {
        return;
    }
    let starts = crate::font::cluster_starts(run);
    for glyph in &run.glyphs {
        if glyph.id != 0 {
            continue;
        }
        let text = crate::font::cluster_source_text(run, &starts, glyph.cluster);
        let Some(ch) = text.chars().next() else {
            continue;
        };
        let Some((font_id, glyph_id)) = faces.resolve(ch) else {
            continue;
        };
        let entry = out.entry(font_id).or_default();
        // `.notdef` still carries its whole source cluster so extraction can
        // recover characters no registered face covers.
        entry.entry(glyph_id).or_default().insert(text);
    }
}

/// A maximal run of consecutive scalars painted with the same font and size.
struct Segment {
    font_id: u32,
    /// Font-size multiplier: `0.7` for a synthesized small cap.
    scale: f64,
    /// `(glyph id, advance in PDF points, source text)` in source order.
    glyphs: Vec<(u16, f64, String)>,
}

fn shape_text(faces: &FontFaces<'_>, run: &TextRunPrimitive, size_px: f64) -> Vec<Segment> {
    let mut segments: Vec<Segment> = Vec::new();
    for painted in painted_chars(run) {
        for ch in painted.chars {
            let Some((font_id, glyph)) = faces.resolve(ch) else {
                continue;
            };
            let advance =
                faces.advance_px(font_id, glyph, size_px * painted.scale) * f64::from(PX_TO_PT);
            match segments.last_mut() {
                Some(segment) if segment.font_id == font_id && segment.scale == painted.scale => {
                    segment.glyphs.push((glyph, advance, ch.to_string()));
                }
                _ => segments.push(Segment {
                    font_id,
                    scale: painted.scale,
                    glyphs: vec![(glyph, advance, ch.to_string())],
                }),
            }
        }
    }
    segments
}

/// Build a text matrix for a baseline origin plus an offset along the
/// (possibly rotated) baseline direction.
///
/// CSS rotates clockwise in display space (y down); PDF space is y up, so the
/// angle is negated. The matrix is `[cos, sin, -sin, cos, x, y]` with that
/// negated angle.
fn text_matrix(
    origin_x: f32,
    origin_y: f32,
    offset: f64,
    rotation_deg: f64,
    center_x: f32,
    center_y: f32,
) -> [f32; 6] {
    let point_x = origin_x + offset as f32;
    if rotation_deg.abs() <= f64::EPSILON || !rotation_deg.is_finite() {
        return [1.0, 0.0, 0.0, 1.0, point_x, origin_y];
    }
    let theta = -rotation_deg.to_radians();
    let (sin, cos) = theta.sin_cos();
    let (a, b, c, d) = (cos as f32, sin as f32, -sin as f32, cos as f32);
    let dx = point_x - center_x;
    let dy = origin_y - center_y;
    [
        a,
        b,
        c,
        d,
        center_x + a * dx + c * dy,
        center_y + b * dx + d * dy,
    ]
}

fn optional_number(value: Option<&Number>, default: f64) -> f64 {
    value
        .and_then(Number::as_f64)
        .filter(|value| value.is_finite())
        .unwrap_or(default)
}

fn checked_scale(value: Option<&Number>) -> Result<f32, PdfError> {
    // Absent means the CSS default. A present value must be a positive,
    // finite percentage: `0` would collapse every advance and NaN has no
    // visual meaning, so both are refused instead of silently defaulted.
    let Some(value) = value else {
        return Ok(100.0);
    };
    let scale = value
        .as_f64()
        .ok_or(PdfError::InvalidNumber("horizontal scale"))? as f32;
    if scale.is_finite() && scale > 0.0 {
        Ok(scale)
    } else {
        Err(PdfError::InvalidNumber("horizontal scale"))
    }
}

fn cid(font: &EmbeddedFont, glyph: u32, text: &str) -> Option<u16> {
    font.cid(glyph, text).or_else(|| font.glyph(0))
}

/// Paint a fallback text run. Never rejects its text; only a malformed
/// coordinate or a non-positive horizontal scale is a PDF number error.
pub(crate) fn paint_text(
    content: &mut Content,
    page: &DisplayPage,
    text: &TextRunPrimitive,
    fonts: &BTreeMap<u32, EmbeddedFont>,
    store: &FontStore,
) -> Result<(), PdfError> {
    if text.hidden || text.text.is_empty() {
        return Ok(());
    }

    let size_px = f64::from(font_size_px(&text.font));
    let size = (size_px * f64::from(PX_TO_PT)) as f32;
    let origin_x = points(&text.x, "text x")?;
    let page_height = points(&page.height, "page height")?;
    let baseline = points(&text.baseline_y, "text baseline")?;
    let origin_y = page_height - baseline;
    let width = points(&text.width, "text width")?;
    let center_x = origin_x + width / 2.0;
    let center_y = origin_y + size * 0.2;
    let color = parse_color(&text.color).unwrap_or([0.0, 0.0, 0.0]);
    let rotation = optional_number(text.rotation_deg.as_ref(), 0.0);
    let horizontal_scale = checked_scale(text.horizontal_scale.as_ref())?;
    let rtl = text.rtl == Some(true);

    let actual_text = !text.text.is_empty();
    if actual_text {
        content
            .begin_marked_content_with_properties(Name(b"Span"))
            .properties()
            .actual_text(TextStr(text.text.as_str()));
    }

    let faces = FontFaces::with_family(store, css_family(&text.font).as_deref());
    let segments = shape_text(&faces, text, size_px);
    let usable = !fonts.is_empty()
        && segments
            .iter()
            .any(|segment| fonts.contains_key(&segment.font_id));
    if faces.is_empty() || !usable {
        paint_builtin_text(
            content,
            &display_text(text),
            origin_x,
            origin_y,
            size,
            color,
            rotation,
            horizontal_scale,
            width,
            rtl,
        );
        if actual_text {
            content.end_marked_content();
        }
        return Ok(());
    }

    let scale = f64::from(horizontal_scale) / 100.0;
    // `width` is the measured, already horizontally scaled extent, so a
    // right-to-left run starts at its measured right edge and only the
    // per-glyph advances below are scaled again.
    let mut cursor = if rtl {
        f64::from(origin_x) + f64::from(width)
    } else {
        f64::from(origin_x)
    };
    let mut painted = 0usize;

    content.begin_text();
    content.set_fill_rgb(color[0], color[1], color[2]);
    // Tz is text state and persists across BT/ET. Always reset it so one
    // expanded run cannot stretch every later run on the page.
    content.set_horizontal_scaling(horizontal_scale);

    for segment in &segments {
        let Some(font) = fonts.get(&segment.font_id) else {
            continue;
        };
        content.set_font(Name(font.resource.as_bytes()), size * segment.scale as f32);
        if rtl {
            for (glyph, advance, text) in &segment.glyphs {
                cursor -= advance * scale;
                let Some(glyph) = cid(font, u32::from(*glyph), text) else {
                    continue;
                };
                content.set_text_matrix(text_matrix(
                    origin_x,
                    origin_y,
                    cursor - f64::from(origin_x),
                    rotation,
                    center_x,
                    center_y,
                ));
                content.show(Str(&glyph.to_be_bytes()));
                painted += 1;
            }
        } else {
            let start = cursor;
            let mut bytes = Vec::with_capacity(segment.glyphs.len() * 2);
            for (glyph, _, text) in &segment.glyphs {
                let Some(glyph) = cid(font, u32::from(*glyph), text) else {
                    continue;
                };
                bytes.extend_from_slice(&glyph.to_be_bytes());
            }
            if bytes.is_empty() {
                continue;
            }
            content.set_text_matrix(text_matrix(
                origin_x,
                origin_y,
                start - f64::from(origin_x),
                rotation,
                center_x,
                center_y,
            ));
            content.show(Str(&bytes));
            cursor += segment
                .glyphs
                .iter()
                .map(|(_, advance, _)| *advance * scale)
                .sum::<f64>();
            painted += 1;
        }
    }
    content.end_text();

    if painted == 0 {
        paint_builtin_text(
            content,
            &display_text(text),
            origin_x,
            origin_y,
            size,
            color,
            rotation,
            horizontal_scale,
            width,
            rtl,
        );
    }
    if actual_text {
        content.end_marked_content();
    }
    Ok(())
}

/// Paint with the Base-14 resource. Non-ASCII scalars become `?`; this path is
/// only reachable when the caller registered no usable font.
#[allow(clippy::too_many_arguments)]
fn paint_builtin_text(
    content: &mut Content,
    text: &str,
    origin_x: f32,
    origin_y: f32,
    size: f32,
    color: [f32; 3],
    rotation: f64,
    horizontal_scale: f32,
    width: f32,
    rtl: bool,
) {
    if text.is_empty() {
        return;
    }
    let bytes: Vec<u8> = text
        .chars()
        .map(|ch| if ch.is_ascii() { ch as u8 } else { b'?' })
        .collect();
    let center_x = origin_x + width / 2.0;
    let center_y = origin_y + size * 0.2;

    content.begin_text();
    content.set_font(BASE_FONT, size);
    content.set_fill_rgb(color[0], color[1], color[2]);
    content.set_horizontal_scaling(horizontal_scale);
    if rtl {
        // `width` already includes the horizontal scale; `Tz` never applies
        // to a text matrix translation, so each byte needs its measured
        // advance here and the run must start at its measured right edge.
        let advance = f64::from(width) / bytes.len().max(1) as f64;
        let mut cursor = f64::from(origin_x) + f64::from(width);
        for byte in bytes {
            cursor -= advance;
            content.set_text_matrix(text_matrix(
                origin_x,
                origin_y,
                cursor - f64::from(origin_x),
                rotation,
                center_x,
                center_y,
            ));
            content.show(Str(&[byte]));
        }
    } else {
        content.set_text_matrix(text_matrix(
            origin_x, origin_y, 0.0, rotation, center_x, center_y,
        ));
        content.show(Str(&bytes));
    }
    content.end_text();
}

/// Paint a positioned glyph run, honouring rotation and horizontal scaling.
///
/// Kept for callers that have no font registry; missing-glyph clusters paint
/// the run font's `.notdef`.
pub(crate) fn paint_glyph_run(
    content: &mut Content,
    page: &DisplayPage,
    run: &GlyphRunPrimitive,
    fonts: &BTreeMap<u32, EmbeddedFont>,
) -> Result<(), PdfError> {
    paint_glyph_run_impl(content, page, run, fonts, None)
}

/// Paint a positioned glyph run, resolving gid-0 clusters against `store`.
///
/// The shaped run names one font, but that font may not cover the cluster's
/// script (emoji in a Latin run is the common case). Each missing cluster is
/// re-resolved through the registry and painted with the fallback font's
/// resource at the shaped glyph position.
pub(crate) fn paint_glyph_run_with_store(
    content: &mut Content,
    page: &DisplayPage,
    run: &GlyphRunPrimitive,
    fonts: &BTreeMap<u32, EmbeddedFont>,
    store: &FontStore,
) -> Result<(), PdfError> {
    paint_glyph_run_impl(content, page, run, fonts, Some(store))
}

fn paint_glyph_run_impl(
    content: &mut Content,
    page: &DisplayPage,
    run: &GlyphRunPrimitive,
    fonts: &BTreeMap<u32, EmbeddedFont>,
    store: Option<&FontStore>,
) -> Result<(), PdfError> {
    if run.hidden || run.glyphs.is_empty() {
        return Ok(());
    }
    if !fonts.contains_key(&run.font_id) {
        return Err(PdfError::Unsupported("glyph run font was not embedded"));
    }

    let size = (run.size * f64::from(PX_TO_PT)) as f32;
    if !size.is_finite() || size <= 0.0 {
        return Err(PdfError::InvalidNumber("glyph run size"));
    }
    let page_height = points(&page.height, "page height")?;
    let color = parse_color(&run.color).unwrap_or([0.0, 0.0, 0.0]);
    let horizontal_scale = checked_scale(run.horizontal_scale.as_ref())?;
    let scale = f64::from(horizontal_scale) / 100.0;
    let rotation = optional_number(run.rotation_deg.as_ref(), 0.0);
    let rotated = rotation.abs() > f64::EPSILON && rotation.is_finite();

    let has_missing_glyph = run.glyphs.iter().any(|glyph| glyph.id == 0);
    let faces = store.filter(|_| has_missing_glyph).map(FontFaces::new);
    let starts = crate::font::cluster_starts(run);
    let mut bases: BTreeMap<u32, (f64, usize)> = BTreeMap::new();
    for (index, glyph) in run.glyphs.iter().enumerate() {
        if bases
            .get(&glyph.cluster)
            .is_none_or(|(advance, _)| glyph.advance > *advance)
        {
            bases.insert(glyph.cluster, (glyph.advance, index));
        }
    }
    let is_base = |index: usize, glyph: &PlacedGlyph| {
        bases
            .get(&glyph.cluster)
            .is_some_and(|(_, base)| *base == index)
    };
    let has_non_base = run
        .glyphs
        .iter()
        .enumerate()
        .any(|(index, glyph)| !is_base(index, glyph));
    let target_for = |index: usize, glyph: &PlacedGlyph| -> (u32, u32, String) {
        let text = crate::font::cluster_source_text(run, &starts, glyph.cluster);
        if !is_base(index, glyph) {
            return (run.font_id, glyph.id, "\u{200B}".to_owned());
        }
        if glyph.id == 0
            && let Some(faces) = faces.as_ref()
            && let Some(ch) = text.chars().next()
            && let Some((font_id, glyph_id)) = faces.resolve(ch)
            && fonts.contains_key(&font_id)
        {
            return (font_id, u32::from(glyph_id), text);
        }
        (run.font_id, glyph.id, text)
    };
    let targets: Vec<Option<(u32, u16)>> = run
        .glyphs
        .iter()
        .enumerate()
        .map(|(index, glyph)| {
            let (font_id, glyph_id, text) = target_for(index, glyph);
            let glyph_id = fonts
                .get(&font_id)
                .and_then(|font| cid(font, glyph_id, &text))?;
            Some((font_id, glyph_id))
        })
        .collect();

    // Match the raster oracle: rotate around the run box's centre, not the
    // first glyph. `glyph_rect` there spans every advance and a 1.2em line box.
    let pivot = run.glyphs.first().expect("non-empty glyph run");
    let mut left = f64::INFINITY;
    let mut right = f64::NEG_INFINITY;
    for glyph in &run.glyphs {
        if !glyph.x.is_finite() || !glyph.advance.is_finite() {
            continue;
        }
        left = left.min(glyph.x);
        right = right.max(glyph.x + glyph.advance);
    }
    let (pivot_x_px, pivot_y_px) = if left.is_finite() && right.is_finite() {
        ((left + right) / 2.0, pivot.y - run.size * 0.2)
    } else {
        (pivot.x, pivot.y)
    };
    // `docx-raster::font::glyph_rect` scales about the run's left edge.
    let left_px = if left.is_finite() { left } else { pivot.x };
    let pivot_x = f64_to_points(pivot_x_px, "glyph x")?;
    let pivot_y = page_height - f64_to_points(pivot_y_px, "glyph y")?;
    let theta = -rotation.to_radians();
    let (sin, cos) = theta.sin_cos();
    let (a, b, c, d) = if rotated {
        (cos as f32, sin as f32, -sin as f32, cos as f32)
    } else {
        (1.0, 0.0, 0.0, 1.0)
    };
    let (display_sin, display_cos) = rotation.to_radians().sin_cos();

    content.begin_text();
    content.set_fill_rgb(color[0], color[1], color[2]);
    // Tz persists across BT/ET; reset it for every run.
    content.set_horizontal_scaling(horizontal_scale);

    let uniform_baseline = run
        .glyphs
        .windows(2)
        .all(|pair| (pair[0].y - pair[1].y).abs() <= f64::EPSILON);
    let mut font_ids = targets.iter().flatten().map(|(font_id, _)| *font_id);
    let uniform_font = match font_ids.next() {
        Some(first) if font_ids.all(|font_id| font_id == first) => Some(first),
        _ => None,
    };
    if !rotated
        && uniform_baseline
        && !has_non_base
        && let Some(font_id) = uniform_font
    {
        // One text object and one TJ: concatenated glyphs advance from the
        // embedded widths, and TJ corrections pin each glyph to its shaped
        // position. This also keeps Poppler from inferring word breaks between
        // separately positioned glyphs.
        let font = fonts.get(&font_id).expect("target font is embedded");
        content.set_font(Name(font.resource.as_bytes()), size);
        let left = f64::from(f64_to_points(left_px, "glyph x")?);
        let first = run.glyphs.first().expect("non-empty glyph run");
        let first_x = f64::from(f64_to_points(first.x, "glyph x")?);
        let first_x = left + (first_x - left) * scale;
        let baseline = f64::from(page_height) - f64::from(f64_to_points(first.y, "glyph y")?);
        content.set_text_matrix([1.0, 0.0, 0.0, 1.0, first_x as f32, baseline as f32]);
        let mut positioned = content.show_positioned();
        let mut items = positioned.items();
        let mut pen = first_x;
        for (index, target) in targets.iter().enumerate() {
            let Some((_, glyph_id)) = target else {
                continue;
            };
            items.show(Str(&glyph_id.to_be_bytes()));
            let width = font
                .widths
                .get(usize::from(*glyph_id))
                .copied()
                .unwrap_or(1000.0) as f64
                * f64::from(size)
                / 1000.0;
            pen += width * scale;
            if let Some(next) = run.glyphs.get(index + 1) {
                let next_x = f64::from(f64_to_points(next.x, "glyph x")?);
                let desired = left + (next_x - left) * scale;
                // TJ adjustments are thousandths of the *scaled* text
                // space advance, so the horizontal scale divides out here.
                let adjustment = (pen - desired) * 1000.0 / (f64::from(size) * scale);
                items.adjust(adjustment as f32);
                pen = desired;
            }
        }
        drop(items);
        drop(positioned);
        content.end_text();
        return Ok(());
    }

    let mut current_font = None;
    for (index, (glyph, target)) in run.glyphs.iter().zip(&targets).enumerate() {
        let Some((font_id, glyph_id)) = target else {
            continue;
        };
        let is_base = bases
            .get(&glyph.cluster)
            .is_some_and(|(_, base)| *base == index);
        // Dots and marks are part of the cluster's visual glyphs, not extra
        // characters. An empty ActualText keeps them out of the text layer
        // while the glyph still paints.
        if !is_base {
            content
                .begin_marked_content_with_properties(Name(b"Span"))
                .properties()
                .actual_text(TextStr(""));
        }
        let target_font = fonts.get(font_id).expect("target font is embedded");
        if current_font != Some(*font_id) {
            content.set_font(Name(target_font.resource.as_bytes()), size);
            current_font = Some(*font_id);
        }
        let (x, y) = if rotated {
            // Accepted rotation semantics (HEAD 4a8ae8c2): the matrix is
            // anchored at the glyph-box centre, but the offsets are measured
            // from the run's first glyph. Horizontal scale multiplies the
            // offset once, before the rotation.
            let dx = (glyph.x - pivot.x) * scale;
            let dy = glyph.y - pivot.y;
            let rotated_x = dx * display_cos - dy * display_sin;
            let rotated_y = dx * display_sin + dy * display_cos;
            (
                pivot_x + f64_to_points(rotated_x, "glyph x")?,
                pivot_y - f64_to_points(rotated_y, "glyph y")?,
            )
        } else {
            // Scale about the run's left edge, as `docx-raster::font::glyph_rect`
            // does through `primitive_visual_transform`.
            let scaled_x = left_px + (glyph.x - left_px) * scale;
            (
                f64_to_points(scaled_x, "glyph x")?,
                page_height - f64_to_points(glyph.y, "glyph y")?,
            )
        };
        content.set_text_matrix([a, b, c, d, x, y]);
        content.show(Str(&glyph_id.to_be_bytes()));
        if !is_base {
            content.end_marked_content();
        }
    }
    content.end_text();
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::{
        DisplayList, DocAttrs, GlyphRunPrimitive, PlacedGlyph, Primitive,
    };
    use std::collections::BTreeSet;

    fn font_bytes(name: &str) -> Vec<u8> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../packages/fonts/assets")
            .join(name);
        std::fs::read(path).expect("font fixture")
    }

    fn page(primitives: Vec<Primitive>) -> DisplayPage {
        DisplayPage {
            page_index: 0,
            width: Number::from(816),
            height: Number::from(1056),
            content_bounds: None,
            column_bounds: Vec::new(),
            section_id: None,
            section_index: None,
            section_page_index: None,
            section_page_number: None,
            page_label: None,
            primitives,
            watermark_primitive_count: None,
            background: None,
            page_borders: Vec::new(),
            header: None,
            footer: None,
            note_areas: Vec::new(),
            reservations: Vec::new(),
            table_gridlines: Vec::new(),
            unbuilt: false,
            position_span: None,
            hf_parts: None,
        }
    }

    fn text_run(value: &str) -> TextRunPrimitive {
        TextRunPrimitive {
            text: value.to_owned(),
            paint_text: None,
            x: Number::from(72),
            baseline_y: Number::from(96),
            width: Number::from(200),
            paint_clip: None,
            font: "16px Roboto".to_owned(),
            scripts: None,
            color: "#112233".to_owned(),
            letter_spacing: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            substitute_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: DocAttrs::default(),
        }
    }

    fn glyph_run(font_id: u32) -> GlyphRunPrimitive {
        GlyphRunPrimitive {
            font_id,
            size: 16.0,
            color: "#112233".to_owned(),
            text: "Hi".to_owned(),
            paint_text: None,
            glyphs: vec![
                PlacedGlyph {
                    id: u32::from('H'),
                    x: 72.0,
                    y: 96.0,
                    cluster: 0,
                    advance: 10.0,
                    logical_order: None,
                    bidi_level: None,
                },
                PlacedGlyph {
                    id: u32::from('i'),
                    x: 82.0,
                    y: 96.0,
                    cluster: 1,
                    advance: 5.0,
                    logical_order: None,
                    bidi_level: None,
                },
            ],
            echo: None,
            paint_clip: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: DocAttrs::default(),
        }
    }

    #[test]
    fn collects_per_character_fallbacks() {
        let mut store = FontStore::default();
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let hebrew = store
            .register(font_bytes("NotoSansHebrew-Regular.ttf"))
            .expect("hebrew");
        let mut out = BTreeMap::new();
        collect_text_glyphs(&store, &text_run("AשלוםB"), &mut out);
        assert!(out.contains_key(&roboto.to_u32()));
        assert!(out.contains_key(&hebrew.to_u32()));
        assert!(
            out.values()
                .flat_map(BTreeMap::values)
                .all(|text| !text.is_empty())
        );

        let mut missing = BTreeMap::new();
        collect_text_glyphs(&store, &text_run("你好🙂"), &mut missing);
        assert!(!missing.is_empty());
        assert!(
            missing
                .values()
                .flat_map(BTreeMap::values)
                .all(|text| !text.is_empty())
        );
        // No fixture covers CJK or emoji, so every scalar paints `.notdef`.
        assert!(missing.values().all(|glyphs| glyphs.contains_key(&0)));
    }

    #[test]
    fn embeds_a_font_for_missing_glyphs_only() {
        let mut store = FontStore::default();
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let mut out = BTreeMap::new();
        collect_text_glyphs(&store, &text_run("你好🙂"), &mut out);
        let uses = out.get(&roboto.to_u32()).expect("fallback font");
        let font =
            EmbeddedFont::new("F1".to_owned(), &store, roboto.to_u32(), uses).expect("embed");
        assert_eq!(font.glyph(0), Some(0));
        assert!(!font.tounicode.is_empty());
    }

    #[test]
    fn paints_missing_glyph_text_without_rejecting() {
        let mut store = FontStore::default();
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let text = text_run("A你🙂שלוםB");
        let page = page(vec![Primitive::Text(text.clone())]);
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![page.clone()],
        };
        let uses = crate::font::collect_with_text(&display_list, &store);
        let mut fonts = BTreeMap::new();
        for (index, (font_id, glyphs)) in uses.iter().enumerate() {
            fonts.insert(
                *font_id,
                EmbeddedFont::new(format!("F{}", index + 1), &store, *font_id, glyphs)
                    .expect("embed"),
            );
        }
        assert!(fonts.contains_key(&roboto.to_u32()));
        let mut content = Content::new();
        paint_text(&mut content, &page, &text, &fonts, &store).expect("paint fallback text");
        let stream = content.finish();
        assert!(
            stream
                .windows(2)
                .any(|window| window == b"Tj" || window == b"TJ")
        );
        assert!(stream.windows(2).any(|window| window == b"BT" as &[u8]));
    }

    #[test]
    fn rotates_glyph_runs() {
        let mut store = FontStore::default();
        let bytes = font_bytes("Roboto-Regular.ttf");
        let roboto = store.register(bytes.clone()).expect("roboto");
        let face = Face::parse(&bytes, 0).expect("parse");
        let mut glyphs = Vec::new();
        for (cluster, ch) in "Hi".chars().enumerate() {
            let id = face.glyph_index(ch).expect("glyph").0;
            glyphs.push(PlacedGlyph {
                id: u32::from(id),
                x: 72.0 + cluster as f64 * 10.0,
                y: 96.0,
                cluster: cluster as u32,
                advance: 10.0,
                logical_order: None,
                bidi_level: None,
            });
        }
        let mut run = glyph_run(roboto.to_u32());
        run.glyphs = glyphs;
        run.rotation_deg = Number::from_f64(45.0);
        let uses = BTreeMap::from_iter(run.glyphs.iter().map(|glyph| {
            (
                glyph.id as u16,
                BTreeSet::from([run.text.chars().next().unwrap().to_string()]),
            )
        }));
        let font =
            EmbeddedFont::new("F1".to_owned(), &store, roboto.to_u32(), &uses).expect("embed");
        let fonts = BTreeMap::from([(roboto.to_u32(), font)]);
        let page = page(vec![Primitive::GlyphRun(run.clone())]);
        let mut content = Content::new();
        paint_glyph_run(&mut content, &page, &run, &fonts).expect("rotated glyph run");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        assert!(stream.contains("Tm"));
        assert!(
            stream.contains("0.7071"),
            "rotated text matrix missing: {stream}"
        );
    }

    #[test]
    fn paints_notdef_only_glyph_runs() {
        let mut store = FontStore::default();
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let uses = BTreeMap::from([(0, BTreeSet::from(["\u{FFFD}".to_owned()]))]);
        let font =
            EmbeddedFont::new("F1".to_owned(), &store, roboto.to_u32(), &uses).expect("embed");
        let fonts = BTreeMap::from([(roboto.to_u32(), font)]);
        let mut run = glyph_run(roboto.to_u32());
        for glyph in &mut run.glyphs {
            glyph.id = 0;
        }
        let page = page(vec![Primitive::GlyphRun(run.clone())]);
        let mut content = Content::new();
        paint_glyph_run(&mut content, &page, &run, &fonts).expect("notdef run");
        let stream = content.finish();
        assert!(
            stream
                .windows(2)
                .any(|window| window == b"Tj" || window == b"TJ")
        );
    }

    fn tm_xs(stream: &str) -> Vec<f64> {
        tm_matrices(stream).iter().map(|matrix| matrix[4]).collect()
    }

    fn tm_matrices(stream: &str) -> Vec<[f64; 6]> {
        stream
            .lines()
            .filter_map(|line| {
                let tokens: Vec<&str> = line.split_whitespace().collect();
                if tokens.last() != Some(&"Tm") {
                    return None;
                }
                let start = tokens.len().checked_sub(7)?;
                let mut matrix = [0.0; 6];
                for (slot, token) in matrix.iter_mut().zip(&tokens[start..start + 6]) {
                    *slot = token.parse().ok()?;
                }
                Some(matrix)
            })
            .collect()
    }

    fn tf_sizes(stream: &str) -> Vec<f64> {
        stream
            .lines()
            .filter_map(|line| {
                let tokens: Vec<&str> = line.split_whitespace().collect();
                if tokens.last() != Some(&"Tf") {
                    return None;
                }
                tokens.get(tokens.len().checked_sub(2)?)?.parse().ok()
            })
            .collect()
    }

    /// Numeric TJ items, i.e. the per-glyph adjustments.
    fn tj_adjustments(stream: &str) -> Vec<f64> {
        stream
            .lines()
            .filter(|line| line.trim_end().ends_with("TJ"))
            .flat_map(|line| {
                line.split(|character: char| {
                    character.is_whitespace() || matches!(character, '[' | ']' | '(' | ')')
                })
                .filter_map(|token| token.parse::<f64>().ok())
                .collect::<Vec<_>>()
            })
            .collect()
    }

    /// Advance in CSS pixels for `ch` at `size` px, as the painter computes it.
    fn advance_px(bytes: &[u8], ch: char, size: f64) -> f64 {
        let face = Face::parse(bytes, 0).expect("parse font");
        let glyph = face.glyph_index(ch).expect("glyph");
        f64::from(face.glyph_hor_advance(glyph).expect("advance")) * size
            / f64::from(face.units_per_em())
    }

    fn placed_glyphs(bytes: &[u8], text: &str, positions: &[(f64, f64)]) -> Vec<PlacedGlyph> {
        let face = Face::parse(bytes, 0).expect("parse font");
        text.chars()
            .enumerate()
            .map(|(index, ch)| {
                let (x, advance) = positions[index];
                PlacedGlyph {
                    id: u32::from(face.glyph_index(ch).expect("glyph").0),
                    x,
                    y: 96.0,
                    cluster: index as u32,
                    advance,
                    logical_order: None,
                    bidi_level: None,
                }
            })
            .collect()
    }

    fn embed_fonts(
        store: &FontStore,
        uses: &BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>>,
    ) -> BTreeMap<u32, EmbeddedFont> {
        uses.iter()
            .enumerate()
            .map(|(index, (font_id, glyphs))| {
                (
                    *font_id,
                    EmbeddedFont::new(format!("F{}", index + 1), store, *font_id, glyphs)
                        .expect("embed"),
                )
            })
            .collect()
    }

    /// Paint a glyph run over a fresh Roboto store and return the content
    /// stream plus the embedded fonts, whose widths the TJ path relies on.
    fn painted_glyph_run(
        glyphs: Vec<PlacedGlyph>,
        rotation: Option<f64>,
        horizontal_scale: Option<f64>,
    ) -> (String, FontStore, BTreeMap<u32, EmbeddedFont>) {
        let bytes = font_bytes("Roboto-Regular.ttf");
        let mut store = FontStore::default();
        let roboto = store.register(bytes).expect("roboto");
        let mut run = glyph_run(roboto.to_u32());
        run.glyphs = glyphs;
        run.rotation_deg = rotation.and_then(Number::from_f64);
        run.horizontal_scale = horizontal_scale.and_then(Number::from_f64);
        let page = page(vec![Primitive::GlyphRun(run.clone())]);
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![page.clone()],
        };
        let uses = crate::font::collect(&display_list);
        let fonts = embed_fonts(&store, &uses);
        let mut content = Content::new();
        paint_glyph_run(&mut content, &page, &run, &fonts).expect("paint glyph run");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        (stream, store, fonts)
    }

    fn paint_fallback(
        value: &str,
        rtl: bool,
        rotation: Option<f64>,
        horizontal_scale: Option<f64>,
    ) -> String {
        let mut store = FontStore::default();
        store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        store
            .register(font_bytes("NotoSansHebrew-Regular.ttf"))
            .expect("hebrew");
        let mut text = text_run(value);
        text.rtl = rtl.then_some(true);
        text.rotation_deg = rotation.and_then(Number::from_f64);
        text.horizontal_scale = horizontal_scale.and_then(Number::from_f64);
        let page = page(vec![Primitive::Text(text.clone())]);
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![page.clone()],
        };
        let uses = crate::font::collect_with_text(&display_list, &store);
        let mut fonts = BTreeMap::new();
        for (index, (font_id, glyphs)) in uses.iter().enumerate() {
            fonts.insert(
                *font_id,
                EmbeddedFont::new(format!("F{}", index + 1), &store, *font_id, glyphs)
                    .expect("embed"),
            );
        }
        let mut content = Content::new();
        paint_text(&mut content, &page, &text, &fonts, &store).expect("paint text");
        String::from_utf8_lossy(&content.finish()).into_owned()
    }

    #[test]
    fn paints_rtl_text_right_to_left() {
        let stream = paint_fallback("שלום", true, None, None);
        let xs = tm_xs(&stream);
        assert!(xs.len() >= 2, "expected one matrix per glyph: {stream}");
        assert!(
            xs.windows(2).all(|pair| pair[0] > pair[1]),
            "RTL text did not walk leftward: {xs:?}"
        );
    }

    #[test]
    fn paints_rotated_and_scaled_text() {
        let stream = paint_fallback("Hi", false, Some(30.0), Some(150.0));
        assert!(
            stream.contains("150 Tz"),
            "horizontal scale missing: {stream}"
        );
        assert!(
            !stream.contains("15000 Tz"),
            "scale is percent, not a ratio: {stream}"
        );
        assert!(stream.contains("0.866"), "rotation missing: {stream}");
    }

    #[test]
    fn horizontal_scale_scales_text_segment_origins() {
        let stream = paint_fallback("Aש", false, None, Some(200.0));
        assert!(
            stream.contains("200 Tz"),
            "horizontal scale missing: {stream}"
        );
        let matrices = tm_matrices(&stream);
        assert_eq!(
            matrices.len(),
            2,
            "expected one matrix per font segment: {stream}"
        );
        let left = 54.0;
        let advance_a =
            advance_px(&font_bytes("Roboto-Regular.ttf"), 'A', 16.0) * f64::from(PX_TO_PT) * 2.0;
        assert!(
            (matrices[0][4] - left).abs() < 1e-3,
            "first segment must start at the unscaled left edge: {matrices:?}"
        );
        assert!(
            (matrices[1][4] - (left + advance_a)).abs() < 1e-3,
            "second segment origin must scale about the left edge: {matrices:?}"
        );
    }

    #[test]
    fn horizontal_scale_scales_builtin_text_origins() {
        let store = FontStore::default();
        let fonts = BTreeMap::new();
        let mut run = text_run("Hi");
        run.horizontal_scale = Number::from_f64(200.0);
        // `width` is the measured extent *after* horizontal scale, as layout
        // reports it: a 200px run at 200% measures 400px.
        run.width = Number::from(400);

        let sheet = page(vec![Primitive::Text(run.clone())]);
        let mut content = Content::new();
        paint_text(&mut content, &sheet, &run, &fonts, &store).expect("paint builtin");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 1, "{stream}");
        assert!(
            (matrices[0][4] - 54.0).abs() < 1e-3,
            "left edge must stay fixed under Tz: {matrices:?}"
        );

        run.rtl = Some(true);
        let sheet = page(vec![Primitive::Text(run.clone())]);
        let mut content = Content::new();
        paint_text(&mut content, &sheet, &run, &fonts, &store).expect("paint builtin rtl");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        let mut xs = tm_xs(&stream);
        xs.sort_by(f64::total_cmp);
        assert_eq!(xs.len(), 2, "{stream}");
        assert!((xs[0] - 54.0).abs() < 1e-3, "{xs:?}");
        assert!(
            (xs[1] - 204.0).abs() < 1e-3,
            "RTL must start at the measured right edge without scaling it twice: {xs:?}"
        );
    }

    #[test]
    fn builtin_rtl_treats_width_as_already_scaled() {
        let store = FontStore::default();
        let fonts = BTreeMap::new();
        let mut run = text_run("Hi");
        run.horizontal_scale = Number::from_f64(200.0);
        // 200px measured at 100%; layout's scaled extent stays 200px at 200%.
        run.width = Number::from(200);
        run.rtl = Some(true);

        let sheet = page(vec![Primitive::Text(run.clone())]);
        let mut content = Content::new();
        paint_text(&mut content, &sheet, &run, &fonts, &store).expect("paint builtin rtl");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        let mut xs = tm_xs(&stream);
        xs.sort_by(f64::total_cmp);
        assert_eq!(xs.len(), 2, "{stream}");
        assert!((xs[0] - 54.0).abs() < 1e-3, "{xs:?}");
        assert!(
            (xs[1] - 129.0).abs() < 1e-3,
            "width must be used as the measured extent, not multiplied by Tz: {xs:?}"
        );
    }

    #[test]
    fn horizontal_scale_scales_tj_glyph_origins_and_adjustments() {
        let bytes = font_bytes("Roboto-Regular.ttf");
        // Visual RTL order: the first glyph is not at `glyph_rect.x`, so the
        // single Tm translation itself must be scaled about the left edge.
        let glyphs = placed_glyphs(&bytes, "Hi", &[(100.0, 10.0), (72.0, 5.0)]);
        let (stream, _store, fonts) = painted_glyph_run(glyphs, None, Some(200.0));
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 1, "TJ fast path expected: {stream}");
        assert!(
            (matrices[0][4] - 96.0).abs() < 1e-3,
            "first origin must scale about glyph_rect.x: {matrices:?}"
        );
        assert!((matrices[0][5] - 720.0).abs() < 1e-3, "{matrices:?}");

        // The TJ adjustment must land the second glyph at its scaled shaped
        // position in points; Tz also scales the adjustment itself.
        let font = fonts.values().next().expect("embedded font");
        let face = Face::parse(&bytes, 0).expect("face");
        let first = face.glyph_index('H').expect("H").0;
        let cid = font.cid(u32::from(first), "H").expect("cid");
        let width = f64::from(font.widths[usize::from(cid)]) * 12.0 / 1000.0;
        let pen = 96.0 + width * 2.0;
        let desired = 54.0;
        let expected = (pen - desired) * 1000.0 / (12.0 * 2.0);
        let adjustments = tj_adjustments(&stream);
        assert_eq!(adjustments.len(), 1, "{stream}");
        assert!(
            (adjustments[0] - expected).abs() < 1e-2,
            "adjustment {} does not land at {desired} (expected {expected})",
            adjustments[0]
        );
    }

    #[test]
    fn horizontal_scale_scales_per_glyph_origins() {
        let bytes = font_bytes("Roboto-Regular.ttf");
        let mut glyphs = placed_glyphs(&bytes, "Hi", &[(72.0, 10.0), (100.0, 5.0)]);
        // Different baselines disable the TJ fast path.
        glyphs[1].y = 94.0;
        let (stream, _store, _fonts) = painted_glyph_run(glyphs, None, Some(200.0));
        assert!(stream.contains("200 Tz"), "{stream}");
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 2, "{stream}");
        assert!((matrices[0][4] - 54.0).abs() < 1e-3, "{matrices:?}");
        assert!(
            (matrices[1][4] - 96.0).abs() < 1e-3,
            "per-glyph origin must scale about the left edge: {matrices:?}"
        );
        assert!((matrices[0][5] - 720.0).abs() < 1e-3, "{matrices:?}");
        assert!((matrices[1][5] - 721.5).abs() < 1e-3, "{matrices:?}");
    }

    #[test]
    fn rotated_glyph_run_pivots_on_the_glyph_rect_centre() {
        let bytes = font_bytes("Roboto-Regular.ttf");
        let glyphs = placed_glyphs(&bytes, "Hi", &[(72.0, 10.0), (90.0, 5.0)]);
        let (stream, _store, _fonts) = painted_glyph_run(glyphs, Some(90.0), None);
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 2, "{stream}");

        // Accepted behaviour (HEAD 4a8ae8c2): the matrix translation is
        // anchored at `docx-raster::font::glyph_rect`'s centre — x = min
        // glyph.x + half the span, y = first.y - 0.2em — while the offsets
        // rotated are measured from the run's first glyph.
        let left = 72.0;
        let right = (72.0f64 + 10.0).max(90.0 + 5.0);
        let center_x = (left + right) / 2.0;
        let center_y = 96.0 - 16.0 * 0.2;
        let first_x = 72.0;
        let page_height = 792.0;
        for (matrix, glyph_x) in matrices.iter().zip([72.0, 90.0]) {
            assert!(matrix[0].abs() < 1e-4, "{matrix:?}");
            assert!((matrix[1] + 1.0).abs() < 1e-4, "{matrix:?}");
            assert!((matrix[2] - 1.0).abs() < 1e-4, "{matrix:?}");
            assert!(matrix[3].abs() < 1e-4, "{matrix:?}");
            let dx = glyph_x - first_x;
            let dy = 0.0;
            // CSS rotate(90°) in display space, y down.
            let expected_x = center_x - dy;
            let expected_y = center_y + dx;
            assert!(
                (matrix[4] - expected_x * f64::from(PX_TO_PT)).abs() < 1e-3,
                "glyph {glyph_x} x: {matrix:?}"
            );
            assert!(
                (matrix[5] - (page_height - expected_y * f64::from(PX_TO_PT))).abs() < 1e-3,
                "glyph {glyph_x} y: {matrix:?}"
            );
        }
    }

    #[test]
    fn rotated_glyph_run_applies_horizontal_scale_once() {
        let bytes = font_bytes("Roboto-Regular.ttf");
        let glyphs = placed_glyphs(&bytes, "Hi", &[(72.0, 10.0), (90.0, 5.0)]);
        let (stream, _store, _fonts) = painted_glyph_run(glyphs, Some(90.0), Some(200.0));
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 2, "{stream}");

        // Rotation keeps the accepted first-glyph offsets; horizontal scale
        // multiplies each offset once before the rotation.
        let center_x = (72.0 + 95.0) / 2.0;
        let center_y = 96.0 - 16.0 * 0.2;
        let first_x = 72.0;
        let page_height = 792.0;
        for (matrix, glyph_x) in matrices.iter().zip([72.0, 90.0]) {
            let dx = (glyph_x - first_x) * 2.0;
            let dy = 0.0;
            let expected_x = center_x - dy;
            let expected_y = center_y + dx;
            assert!(
                (matrix[4] - expected_x * f64::from(PX_TO_PT)).abs() < 1e-3,
                "glyph at {glyph_x}px: {matrix:?}"
            );
            assert!(
                (matrix[5] - (page_height - expected_y * f64::from(PX_TO_PT))).abs() < 1e-3,
                "glyph at {glyph_x}px: {matrix:?}"
            );
        }
    }

    #[test]
    fn rotated_glyph_run_rotates_first_glyph_offsets_by_the_run_angle() {
        let bytes = font_bytes("Roboto-Regular.ttf");
        let mut glyphs = placed_glyphs(&bytes, "Hi", &[(72.0, 10.0), (90.0, 5.0)]);
        glyphs[1].y = 98.0;
        let (stream, _store, _fonts) = painted_glyph_run(glyphs, Some(30.0), None);
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 2, "{stream}");

        // Offsets are (glyph - first glyph) in display space, rotated by the
        // run angle and anchored at the glyph_rect centre. Both terms matter:
        // a swapped sin/cos pair would move the second glyph right instead of
        // down-right.
        let center_x = (72.0 + 95.0) / 2.0;
        let center_y = 96.0 - 16.0 * 0.2;
        let page_height = 792.0;
        let (display_sin, display_cos) = 30.0f64.to_radians().sin_cos();
        for (matrix, (glyph_x, glyph_y)) in matrices.iter().zip([(72.0, 96.0), (90.0, 98.0)]) {
            let dx = glyph_x - 72.0;
            let dy = glyph_y - 96.0;
            let expected_x = center_x + dx * display_cos - dy * display_sin;
            let expected_y = center_y + dx * display_sin + dy * display_cos;
            assert!(
                (matrix[4] - expected_x * f64::from(PX_TO_PT)).abs() < 1e-3,
                "glyph at {glyph_x}px: {matrix:?}"
            );
            assert!(
                (matrix[5] - (page_height - expected_y * f64::from(PX_TO_PT))).abs() < 1e-3,
                "glyph at {glyph_x}px: {matrix:?}"
            );
        }
    }

    #[test]
    fn fallback_prefers_the_run_family() {
        let roboto_bytes = font_bytes("Roboto-Regular.ttf");
        let carlito_bytes = font_bytes("Carlito-Regular.ttf");
        let mut store = FontStore::default();
        let roboto = store.register(roboto_bytes.clone()).expect("roboto");
        let carlito = store.register(carlito_bytes.clone()).expect("carlito");
        let mut run = text_run("A");
        run.font = "16px Carlito".to_owned();

        let mut collected = BTreeMap::new();
        collect_text_glyphs(&store, &run, &mut collected);
        assert!(
            collected.contains_key(&carlito.to_u32()),
            "run family was ignored: {collected:?}"
        );
        assert!(
            !collected.contains_key(&roboto.to_u32()),
            "first covering face won over the run family: {collected:?}"
        );

        let glyph_a = |bytes: &[u8]| {
            Face::parse(bytes, 0)
                .expect("face")
                .glyph_index('A')
                .expect("A")
                .0
        };
        let uses = BTreeMap::from([
            (
                roboto.to_u32(),
                BTreeMap::from([(glyph_a(&roboto_bytes), BTreeSet::from(["A".to_owned()]))]),
            ),
            (
                carlito.to_u32(),
                BTreeMap::from([(glyph_a(&carlito_bytes), BTreeSet::from(["A".to_owned()]))]),
            ),
        ]);
        let fonts = BTreeMap::from([
            (
                roboto.to_u32(),
                EmbeddedFont::new(
                    "FA".to_owned(),
                    &store,
                    roboto.to_u32(),
                    &uses[&roboto.to_u32()],
                )
                .expect("embed roboto"),
            ),
            (
                carlito.to_u32(),
                EmbeddedFont::new(
                    "FB".to_owned(),
                    &store,
                    carlito.to_u32(),
                    &uses[&carlito.to_u32()],
                )
                .expect("embed carlito"),
            ),
        ]);
        let page = page(vec![Primitive::Text(run.clone())]);
        let mut content = Content::new();
        paint_text(&mut content, &page, &run, &fonts, &store).expect("paint text");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        assert!(
            stream.contains("/FB 12 Tf"),
            "painting must resolve through the run family: {stream}"
        );
        assert!(!stream.contains("/FA "), "wrong face selected: {stream}");
    }

    #[test]
    fn small_caps_paints_uppercase_glyphs_at_a_scaled_size() {
        let mut store = FontStore::default();
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let mut run = text_run("Hi");
        run.small_caps = true;

        let mut collected = BTreeMap::new();
        collect_text_glyphs(&store, &run, &mut collected);
        let texts: BTreeSet<String> = collected
            .values()
            .flat_map(BTreeMap::values)
            .flatten()
            .cloned()
            .collect();
        assert_eq!(texts, BTreeSet::from(["H".to_owned(), "I".to_owned()]));

        let page = page(vec![Primitive::Text(run.clone())]);
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![page.clone()],
        };
        let uses = crate::font::collect_with_text(&display_list, &store);
        assert!(uses.contains_key(&roboto.to_u32()));
        let fonts = embed_fonts(&store, &uses);
        let mut content = Content::new();
        paint_text(&mut content, &page, &run, &fonts, &store).expect("paint small caps");
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();

        let sizes = tf_sizes(&stream);
        assert_eq!(sizes.len(), 2, "{stream}");
        assert!((sizes[0] - 12.0).abs() < 1e-3, "{sizes:?}");
        assert!(
            (sizes[1] - 12.0 * 0.7).abs() < 1e-3,
            "synthesized small caps keep the run size: {sizes:?}"
        );

        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 2, "{stream}");
        let advance_h =
            advance_px(&font_bytes("Roboto-Regular.ttf"), 'H', 16.0) * f64::from(PX_TO_PT);
        assert!((matrices[0][4] - 54.0).abs() < 1e-3, "{matrices:?}");
        assert!(
            (matrices[1][4] - (54.0 + advance_h)).abs() < 1e-3,
            "small cap must start after the regular glyph's advance: {matrices:?}"
        );
    }

    #[test]
    fn rtl_text_run_starts_at_the_measured_right_edge() {
        // `text_run` width is 200px, i.e. the already horizontally scaled
        // extent at 200% (150pt); it must not be scaled a second time.
        let stream = paint_fallback("שלום", true, None, Some(200.0));
        assert!(stream.contains("200 Tz"), "{stream}");
        let xs = tm_xs(&stream);
        assert_eq!(xs.len(), 4, "{stream}");
        let advance = advance_px(&font_bytes("NotoSansHebrew-Regular.ttf"), 'ש', 16.0)
            * f64::from(PX_TO_PT)
            * 2.0;
        assert!(
            (xs[0] - (54.0 + 150.0 - advance)).abs() < 1e-3,
            "RTL must start at origin_x + measured width: {xs:?}"
        );
        assert!(
            xs.windows(2).all(|pair| pair[0] > pair[1]),
            "RTL text did not walk leftward: {xs:?}"
        );
    }

    #[test]
    fn rotated_text_run_pivots_on_the_measured_rect() {
        let stream = paint_fallback("Hi", false, Some(30.0), Some(200.0));
        let matrices = tm_matrices(&stream);
        assert_eq!(matrices.len(), 1, "{stream}");

        // text_rect centre: x = run.x + measured width / 2, y = baseline -
        // 0.2em. The measured width is 150pt at 200%; using width * scale
        // would pivot 75pt to the right.
        let center_x = 54.0 + 150.0 / 2.0;
        let origin_y = 792.0 - 96.0 * f64::from(PX_TO_PT);
        let center_y = origin_y + 12.0 * 0.2;
        let (sin, cos) = (-30.0f64).to_radians().sin_cos();
        let dx = 54.0 - center_x;
        let dy = origin_y - center_y;
        let expected_x = center_x + cos * dx - sin * dy;
        let expected_y = center_y + sin * dx + cos * dy;
        assert!(
            (matrices[0][4] - expected_x).abs() < 1e-3,
            "pivot must use the measured width: {matrices:?}"
        );
        assert!(
            (matrices[0][5] - expected_y).abs() < 1e-3,
            "pivot must use the measured rect centre: {matrices:?}"
        );
    }

    #[test]
    fn tz_is_set_once_per_text_object() {
        let count = |stream: &str| {
            stream
                .lines()
                .filter(|line| line.trim_end().ends_with("Tz"))
                .count()
        };
        let text = paint_fallback("Aש", false, None, Some(150.0));
        assert_eq!(count(&text), 1, "{text}");
        let bytes = font_bytes("Roboto-Regular.ttf");
        let glyphs = placed_glyphs(&bytes, "Hi", &[(72.0, 10.0), (82.0, 5.0)]);
        let (run, _store, _fonts) = painted_glyph_run(glyphs, None, Some(150.0));
        assert_eq!(count(&run), 1, "{run}");
    }

    #[test]
    fn css_family_parses_the_primary_family() {
        assert_eq!(
            css_family("italic small-caps 700 16px \"Times New Roman\", serif").as_deref(),
            Some("times new roman")
        );
        assert_eq!(
            css_family("400 12.5px Carlito, sans-serif").as_deref(),
            Some("carlito")
        );
        assert_eq!(
            css_family("16px 'Last, First', serif").as_deref(),
            Some("last, first")
        );
        // No size token, empty input and an empty primary family all mean
        // "no preference", not a bogus family match.
        assert_eq!(css_family(""), None);
        assert_eq!(css_family("16px"), None);
        assert_eq!(css_family("12pt Carlito, sans-serif"), None);
        assert_eq!(css_family("16px , serif"), None);
        assert!(!css_small_caps("400 16px small-caps, sans-serif"));
        assert!(css_small_caps("italic small-caps 700 16px Carlito"));
    }

    #[test]
    fn small_caps_cases_only_lowercase_and_never_overrides_all_caps() {
        let mut run = text_run("Hi ß");
        run.small_caps = true;
        let painted: Vec<(char, f64)> = painted_chars(&run)
            .into_iter()
            .flat_map(|painted| painted.chars.into_iter().map(move |ch| (ch, painted.scale)))
            .collect();
        assert_eq!(
            painted,
            vec![
                ('H', 1.0),
                ('I', f64::from(ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE)),
                (' ', 1.0),
                ('S', f64::from(ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE)),
                ('S', f64::from(ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE)),
            ]
        );

        // allCaps wins: small caps is not synthesized on top of it.
        run.all_caps = true;
        assert!(!small_caps(&run));
        let painted: String = painted_chars(&run)
            .into_iter()
            .flat_map(|painted| painted.chars)
            .collect();
        assert_eq!(painted, "HI SS");
        assert!(
            painted_chars(&run)
                .iter()
                .all(|painted| painted.scale == 1.0)
        );
    }

    #[test]
    fn family_preference_keeps_script_fallback() {
        let mut store = FontStore::default();
        let hebrew = store
            .register(font_bytes("NotoSansHebrew-Regular.ttf"))
            .expect("hebrew");
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let mut run = text_run("Aש");
        run.font = "16px Roboto".to_owned();
        let mut collected = BTreeMap::new();
        collect_text_glyphs(&store, &run, &mut collected);
        assert!(
            collected.contains_key(&roboto.to_u32()),
            "run family was ignored: {collected:?}"
        );
        assert!(
            collected.contains_key(&hebrew.to_u32()),
            "preference broke the Hebrew fallback: {collected:?}"
        );

        // A family nobody declares leaves registry order untouched, including
        // the `.notdef` face choice.
        run.font = "16px No Such Family".to_owned();
        let preferred = FontFaces::with_family(&store, css_family(&run.font).as_deref());
        assert!(preferred.preferred.is_empty());
        let plain = FontFaces::new(&store);
        for ch in ['A', 'ש', '\u{10FFFD}'] {
            assert_eq!(preferred.resolve(ch), plain.resolve(ch), "{ch}");
        }
    }

    #[test]
    fn notdef_fallback_prefers_truetype_over_a_cff_family() {
        let path = "/usr/share/fonts/adobe-source-han-sans/SourceHanSansCN-Normal.otf";
        let Ok(bytes) = std::fs::read(path) else {
            return;
        };
        let mut store = FontStore::default();
        let cff = store.register(bytes).expect("cff");
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let cff_face = Face::parse(store.font_bytes(cff).expect("bytes"), 0).expect("face");
        let roboto_bytes = font_bytes("Roboto-Regular.ttf");
        let roboto_face = Face::parse(&roboto_bytes, 0).expect("face");
        assert!(cff_face.tables().cff.is_some(), "fixture must be CFF");
        assert!(
            roboto_face.tables().cff.is_none(),
            "fixture must be TrueType"
        );
        // Plane-16 private use is uncovered by both fixtures.
        assert!(cff_face.glyph_index('\u{10FFFD}').is_none());
        assert!(roboto_face.glyph_index('\u{10FFFD}').is_none());

        let faces = FontFaces::with_family(&store, Some("source han sans cn"));
        assert_eq!(faces.preferred, vec![cff.to_u32()]);
        assert_eq!(
            faces.resolve('\u{10FFFD}'),
            Some((roboto.to_u32(), 0)),
            "a preferred CFF face must not displace the TrueType .notdef"
        );
    }

    #[test]
    fn horizontal_scale_rejects_invalid_values() {
        assert_eq!(checked_scale(None).expect("default"), 100.0);
        let valid = Number::from_f64(150.0).expect("finite");
        assert_eq!(checked_scale(Some(&valid)).expect("valid"), 150.0);
        for invalid in [0.0, -100.0, 1e300] {
            let value = Number::from_f64(invalid).expect("finite");
            assert!(
                matches!(
                    checked_scale(Some(&value)),
                    Err(PdfError::InvalidNumber("horizontal scale"))
                ),
                "{invalid} was accepted"
            );
        }

        let store = FontStore::default();
        let fonts = BTreeMap::new();
        let mut run = text_run("Hi");
        run.horizontal_scale = Number::from_f64(0.0);
        let sheet = page(vec![Primitive::Text(run.clone())]);
        let mut content = Content::new();
        assert!(paint_text(&mut content, &sheet, &run, &fonts, &store).is_err());
    }

    #[test]
    fn collect_includes_gid_zero_for_glyph_runs() {
        let mut store = FontStore::default();
        let roboto = store
            .register(font_bytes("Roboto-Regular.ttf"))
            .expect("roboto");
        let mut run = glyph_run(roboto.to_u32());
        for glyph in &mut run.glyphs {
            glyph.id = 0;
        }
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![page(vec![Primitive::GlyphRun(run)])],
        };
        let uses = crate::font::collect(&display_list);
        let glyphs = uses.get(&roboto.to_u32()).expect("glyph run font");
        assert!(glyphs.contains_key(&0), "gid 0 was dropped: {glyphs:?}");
        assert!(glyphs.values().all(|texts| !texts.is_empty()));
    }
}
