//! Text painting and per-character font fallback for PDF export.
//!
//! `TextRunPrimitive`s carry Unicode text but no glyph ids, so this module
//! resolves every scalar against the font registry that produced the display
//! list and paints the resulting CIDs with embedded subsets. `GlyphRun`s
//! already carry ids; they are painted as positioned CIDs and may be rotated.

use std::collections::BTreeMap;

use docx_layout::display_list::{DisplayPage, GlyphRunPrimitive, PlacedGlyph, TextRunPrimitive};
use ooxml_text::{FontId, FontStore};
use pdf_writer::{Content, Name, Str};
use serde_json::Number;
use ttf_parser::{Face, GlyphId};

use crate::PdfError;
use crate::font::EmbeddedFont;
use crate::units::{PX_TO_PT, f64_to_points, font_size_px, parse_color, points};

/// Base-14 resource registered by the writer for text without embedded fonts.
const BASE_FONT: Name<'static> = Name(b"F0");
/// Source text reported for `.notdef`: the scalar itself is unavailable.
const REPLACEMENT: char = '\u{FFFD}';

/// One parsed face per font id, in registry order.
///
/// Painting and collection must agree on which font supplies each scalar, so
/// both go through [`FontFaces::resolve`].
struct FontFaces<'a> {
    faces: Vec<Option<Face<'a>>>,
}

impl<'a> FontFaces<'a> {
    fn new(store: &'a FontStore) -> Self {
        let mut faces = Vec::with_capacity(store.font_count());
        for index in 0..store.font_count() {
            let id = FontId::from_u32(index as u32);
            let face = store
                .font_bytes(id)
                .ok()
                .and_then(|bytes| Face::parse(bytes, 0).ok());
            faces.push(face);
        }
        Self { faces }
    }

    fn is_empty(&self) -> bool {
        self.faces.iter().all(Option::is_none)
    }

    /// First registered font whose cmap maps `ch` to a real glyph.
    ///
    /// Uncovered scalars fall back to the first parseable face's `.notdef`
    /// (glyph 0), which every subset keeps.
    fn resolve(&self, ch: char) -> Option<(u32, u16)> {
        for (index, face) in self.faces.iter().enumerate() {
            let Some(face) = face else {
                continue;
            };
            if let Some(glyph) = face.glyph_index(ch)
                && glyph.0 != 0
            {
                return Some((index as u32, glyph.0));
            }
        }
        self.first_face().map(|index| (index, 0))
    }

    fn first_face(&self) -> Option<u32> {
        self.faces
            .iter()
            .position(Option::is_some)
            .map(|index| index as u32)
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
/// every selected glyph being present in the subset.
pub(crate) fn collect_text_glyphs(
    store: &FontStore,
    text: &str,
    out: &mut BTreeMap<u32, BTreeMap<u16, String>>,
) {
    let faces = FontFaces::new(store);
    for ch in text.chars() {
        let Some((font_id, glyph)) = faces.resolve(ch) else {
            continue;
        };
        let entry = out.entry(font_id).or_default();
        if glyph == 0 {
            entry.entry(0).or_insert_with(|| REPLACEMENT.to_string());
        } else {
            entry.entry(glyph).or_insert_with(|| ch.to_string());
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
    out: &mut BTreeMap<u32, BTreeMap<u16, String>>,
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
        let Some(ch) = crate::font::cluster_source_text(run, &starts, glyph.cluster)
            .chars()
            .next()
        else {
            continue;
        };
        let Some((font_id, glyph_id)) = faces.resolve(ch) else {
            continue;
        };
        let entry = out.entry(font_id).or_default();
        if glyph_id == 0 {
            entry.entry(0).or_insert_with(|| REPLACEMENT.to_string());
        } else {
            entry.entry(glyph_id).or_insert_with(|| ch.to_string());
        }
    }
}

/// A maximal run of consecutive scalars painted with the same font.
struct Segment {
    font_id: u32,
    /// `(glyph id, advance in PDF points)` in source order.
    glyphs: Vec<(u16, f64)>,
}

fn shape_text(faces: &FontFaces<'_>, text: &str, size_px: f64) -> Vec<Segment> {
    let mut segments: Vec<Segment> = Vec::new();
    for ch in text.chars() {
        let Some((font_id, glyph)) = faces.resolve(ch) else {
            continue;
        };
        let advance = faces.advance_px(font_id, glyph, size_px) * f64::from(PX_TO_PT);
        match segments.last_mut() {
            Some(segment) if segment.font_id == font_id => {
                segment.glyphs.push((glyph, advance));
            }
            _ => segments.push(Segment {
                font_id,
                glyphs: vec![(glyph, advance)],
            }),
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
fn text_matrix(origin_x: f32, origin_y: f32, offset: f64, rotation_deg: f64) -> [f32; 6] {
    if rotation_deg.abs() <= f64::EPSILON || !rotation_deg.is_finite() {
        return [1.0, 0.0, 0.0, 1.0, origin_x + offset as f32, origin_y];
    }
    let theta = -rotation_deg.to_radians();
    let (sin, cos) = theta.sin_cos();
    let (a, b, c, d) = (cos as f32, sin as f32, -sin as f32, cos as f32);
    let offset = offset as f32;
    [a, b, c, d, origin_x + offset * a, origin_y + offset * b]
}

fn optional_number(value: Option<&Number>, default: f64) -> f64 {
    value
        .and_then(Number::as_f64)
        .filter(|value| value.is_finite())
        .unwrap_or(default)
}

fn cid(font: &EmbeddedFont, glyph: u32) -> Option<u16> {
    font.glyph(glyph).or_else(|| font.glyph(0))
}

/// Paint a fallback text run. Never rejects its content.
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
    let color = parse_color(&text.color).unwrap_or([0.0, 0.0, 0.0]);
    let rotation = optional_number(text.rotation_deg.as_ref(), 0.0);
    let horizontal_scale = optional_number(text.horizontal_scale.as_ref(), 100.0) as f32;
    let rtl = text.rtl == Some(true);

    let faces = FontFaces::new(store);
    let segments = shape_text(&faces, &text.text, size_px);
    let usable = !fonts.is_empty()
        && segments
            .iter()
            .any(|segment| fonts.contains_key(&segment.font_id));
    if faces.is_empty() || !usable {
        paint_builtin_text(
            content,
            &text.text,
            origin_x,
            origin_y,
            size,
            color,
            rotation,
            horizontal_scale,
            width,
            rtl,
        );
        return Ok(());
    }

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
        content.set_font(Name(font.resource.as_bytes()), size);
        if rtl {
            for &(glyph, advance) in &segment.glyphs {
                cursor -= advance;
                let Some(glyph) = cid(font, u32::from(glyph)) else {
                    continue;
                };
                content.set_text_matrix(text_matrix(
                    origin_x,
                    origin_y,
                    cursor - f64::from(origin_x),
                    rotation,
                ));
                content.show(Str(&glyph.to_be_bytes()));
                painted += 1;
            }
        } else {
            let start = cursor;
            let mut bytes = Vec::with_capacity(segment.glyphs.len() * 2);
            for &(glyph, _) in &segment.glyphs {
                let Some(glyph) = cid(font, u32::from(glyph)) else {
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
            ));
            content.show(Str(&bytes));
            cursor += segment
                .glyphs
                .iter()
                .map(|(_, advance)| *advance)
                .sum::<f64>();
            painted += 1;
        }
    }
    content.end_text();

    if painted == 0 {
        paint_builtin_text(
            content,
            &text.text,
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

    content.begin_text();
    content.set_font(BASE_FONT, size);
    content.set_fill_rgb(color[0], color[1], color[2]);
    content.set_horizontal_scaling(horizontal_scale);
    if rtl {
        let advance = f64::from(width) / bytes.len().max(1) as f64;
        let mut cursor = f64::from(origin_x) + f64::from(width);
        for byte in bytes {
            cursor -= advance;
            content.set_text_matrix(text_matrix(
                origin_x,
                origin_y,
                cursor - f64::from(origin_x),
                rotation,
            ));
            content.show(Str(&[byte]));
        }
    } else {
        content.set_text_matrix(text_matrix(origin_x, origin_y, 0.0, rotation));
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
    if !size.is_finite() {
        return Err(PdfError::InvalidNumber("glyph run size"));
    }
    let page_height = points(&page.height, "page height")?;
    let color = parse_color(&run.color).unwrap_or([0.0, 0.0, 0.0]);
    let horizontal_scale = optional_number(run.horizontal_scale.as_ref(), 100.0) as f32;
    let rotation = optional_number(run.rotation_deg.as_ref(), 0.0);
    let rotated = rotation.abs() > f64::EPSILON && rotation.is_finite();

    let has_missing_glyph = run.glyphs.iter().any(|glyph| glyph.id == 0);
    let faces = store.filter(|_| has_missing_glyph).map(FontFaces::new);
    let fallback_starts = faces.as_ref().map(|_| crate::font::cluster_starts(run));
    let target_for = |glyph: &PlacedGlyph| -> (u32, u32) {
        if glyph.id == 0 {
            if let (Some(faces), Some(starts)) = (faces.as_ref(), fallback_starts.as_ref()) {
                if let Some(ch) = crate::font::cluster_source_text(run, starts, glyph.cluster)
                    .chars()
                    .next()
                {
                    if let Some((font_id, glyph_id)) = faces.resolve(ch) {
                        if fonts.contains_key(&font_id) {
                            return (font_id, u32::from(glyph_id));
                        }
                    }
                }
            }
        }
        (run.font_id, glyph.id)
    };
    let targets: Vec<Option<(u32, u16)>> = run
        .glyphs
        .iter()
        .map(|glyph| {
            let (font_id, glyph_id) = target_for(glyph);
            let glyph_id = fonts.get(&font_id).and_then(|font| cid(font, glyph_id))?;
            Some((font_id, glyph_id))
        })
        .collect();

    // Rotate every glyph around the first glyph's origin so a run keeps its
    // internal offsets and baseline shifts.
    let pivot = run.glyphs.first().expect("non-empty glyph run");
    let pivot_x = f64_to_points(pivot.x, "glyph x")?;
    let pivot_y = page_height - f64_to_points(pivot.y, "glyph y")?;
    let theta = -rotation.to_radians();
    let (sin, cos) = theta.sin_cos();
    let (a, b, c, d) = if rotated {
        (cos as f32, sin as f32, -sin as f32, cos as f32)
    } else {
        (1.0, 0.0, 0.0, 1.0)
    };
    let (display_cos, display_sin) = rotation.to_radians().sin_cos();

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
    if !rotated && uniform_baseline {
        if let Some(font_id) = uniform_font {
            // One text object and one TJ: concatenated glyphs advance from the
            // embedded widths, and TJ corrections pin each glyph to its shaped
            // position. This also keeps Poppler from inferring word breaks between
            // separately positioned glyphs.
            let font = fonts.get(&font_id).expect("target font is embedded");
            content.set_font(Name(font.resource.as_bytes()), size);
            let first = run.glyphs.first().expect("non-empty glyph run");
            let first_x = f64_to_points(first.x, "glyph x")?;
            let baseline = page_height - f64_to_points(first.y, "glyph y")?;
            content.set_text_matrix([1.0, 0.0, 0.0, 1.0, first_x, baseline]);
            let mut positioned = content.show_positioned();
            let mut items = positioned.items();
            let mut pen = f64::from(first_x);
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
                pen += width;
                if let Some(next) = run.glyphs.get(index + 1) {
                    let desired = f64::from(f64_to_points(next.x, "glyph x")?);
                    let adjustment = (pen - desired) * 1000.0 / f64::from(size);
                    items.adjust(adjustment as f32);
                    pen = desired;
                }
            }
            drop(items);
            drop(positioned);
            content.end_text();
            return Ok(());
        }
    }

    let mut current_font = None;
    for (glyph, target) in run.glyphs.iter().zip(&targets) {
        let Some((font_id, glyph_id)) = target else {
            continue;
        };
        let target_font = fonts.get(font_id).expect("target font is embedded");
        if current_font != Some(*font_id) {
            content.set_font(Name(target_font.resource.as_bytes()), size);
            current_font = Some(*font_id);
        }
        let (x, y) = if rotated {
            let dx = glyph.x - pivot.x;
            let dy = glyph.y - pivot.y;
            let rotated_x = dx * display_cos - dy * display_sin;
            let rotated_y = dx * display_sin + dy * display_cos;
            (
                pivot_x + f64_to_points(rotated_x, "glyph x")?,
                pivot_y - f64_to_points(rotated_y, "glyph y")?,
            )
        } else {
            (
                f64_to_points(glyph.x, "glyph x")?,
                page_height - f64_to_points(glyph.y, "glyph y")?,
            )
        };
        content.set_text_matrix([a, b, c, d, x, y]);
        content.show(Str(&glyph_id.to_be_bytes()));
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
        collect_text_glyphs(&store, "AשלוםB", &mut out);
        assert!(out.get(&roboto.to_u32()).is_some());
        assert!(out.get(&hebrew.to_u32()).is_some());
        assert!(
            out.values()
                .flat_map(BTreeMap::values)
                .all(|text| !text.is_empty())
        );

        let mut missing = BTreeMap::new();
        collect_text_glyphs(&store, "你好🙂", &mut missing);
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
        collect_text_glyphs(&store, "你好🙂", &mut out);
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
                run.text.chars().next().unwrap().to_string(),
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
        let uses = BTreeMap::from([(0, "\u{FFFD}".to_owned())]);
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
        stream
            .lines()
            .filter_map(|line| {
                let tokens: Vec<&str> = line.split_whitespace().collect();
                if tokens.last() != Some(&"Tm") {
                    return None;
                }
                tokens.get(tokens.len().checked_sub(3)?)?.parse().ok()
            })
            .collect()
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
        assert!(glyphs.values().all(|text| !text.is_empty()));
    }
}
