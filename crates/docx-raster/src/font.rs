//! Shared-store text shaping and glyph rasterization.

use std::collections::HashMap;
use std::collections::hash_map::Entry;

use docx_layout::display_list::{
    GlyphRunPrimitive, LeaderGlyphMetadata, PlacedGlyph, TextRunPrimitive,
};
use ooxml_text::{
    FontId, FontStore, PathCmd, ShapeDirection, ShapeFeature, shape, shape::shape_with_properties,
};
use tiny_skia::{FillRule, Mask, Paint, Path, PathBuilder, Pixmap, Transform};

use crate::{FRect, RenderResources, color_with_opacity, number_f32, primitive_visual_transform};

const MAX_PAINTED_GLYPHS: usize = 1_000_000;

/// Measures one font run in pixels.
pub fn measure_text(
    fonts: &FontStore,
    font: FontId,
    text: &str,
    size_px: f32,
) -> Result<f32, String> {
    if !size_px.is_finite() || size_px <= 0.0 {
        return Err("font size must be finite and positive".to_string());
    }
    let glyphs = shape(fonts, font, text, size_px, &[]).map_err(|error| error.to_string())?;
    Ok(glyphs.iter().map(|glyph| glyph.x_advance).sum())
}

#[derive(Default)]
pub(crate) struct GlyphCache {
    entries: HashMap<(u32, u32), CachedGlyph>,
}

struct CachedGlyph {
    path: Option<Path>,
    upem: f32,
}

#[derive(Clone)]
struct FontSpec {
    family: String,
    size: f32,
    bold: bool,
    italic: bool,
    small_caps: bool,
}

struct FontSegment {
    start: usize,
    end: usize,
    font: FontId,
}

pub(crate) struct PaintContext<'a, 'b> {
    pub pixmap: &'a mut Pixmap,
    pub resources: &'a RenderResources<'b>,
    pub cache: &'a mut GlyphCache,
    pub base_transform: Transform,
    pub mask: Option<&'a Mask>,
    pub opacity: f32,
}

pub(crate) fn paint_text(
    context: &mut PaintContext<'_, '_>,
    run: &TextRunPrimitive,
) -> Result<(), String> {
    validate_text_effects(run)?;
    if run.text.is_empty() {
        return Ok(());
    }
    let spec = parse_font(&run.font)?;
    if run.small_caps || spec.small_caps {
        return Err("unsupported text field: smallCaps".to_string());
    }
    let rect = text_rect(run, spec.size)?;
    let visual = primitive_visual_transform(
        rect,
        run.rotation_deg.as_ref(),
        run.horizontal_scale.as_ref(),
    )?;
    let transform = context.base_transform.pre_concat(visual);
    let color = color_with_opacity(&run.color, context.opacity)?;
    if let Some(leader) = active_leader(&run.attrs.leader_glyphs) {
        return paint_leader(context, run, leader, &spec, color, transform);
    }
    let text = if run.all_caps {
        run.text.to_uppercase()
    } else {
        run.text.clone()
    };
    let letter_spacing = run
        .letter_spacing
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(0.0);
    let word_spacing = run
        .word_spacing
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(0.0);
    paint_resolved_text(
        context,
        &text,
        number_f32(&run.x)?,
        number_f32(&run.baseline_y)?,
        &spec,
        run.rtl == Some(true),
        run.attrs.lang.as_deref(),
        letter_spacing,
        word_spacing,
        &[],
        color,
        transform,
    )
}

pub(crate) fn paint_glyph_run(
    context: &mut PaintContext<'_, '_>,
    run: &GlyphRunPrimitive,
) -> Result<(), String> {
    validate_glyph_effects(run)?;
    if run.glyphs.len() > MAX_PAINTED_GLYPHS {
        return Err(format!(
            "glyph run exceeds the {MAX_PAINTED_GLYPHS} glyph limit"
        ));
    }
    if run.glyphs.is_empty() {
        return Ok(());
    }
    if !run.size.is_finite() || run.size <= 0.0 || run.size > f32::MAX as f64 {
        return Err("glyph run size must be finite and positive".to_string());
    }
    let rect = glyph_rect(run)?;
    let visual = primitive_visual_transform(
        rect,
        run.rotation_deg.as_ref(),
        run.horizontal_scale.as_ref(),
    )?;
    let transform = context.base_transform.pre_concat(visual);
    let color = color_with_opacity(&run.color, context.opacity)?;
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    let font = FontId::from_u32(run.font_id);
    for glyph in &run.glyphs {
        paint_placed_glyph(context, font, glyph, run.size as f32, &paint, transform)?;
    }
    Ok(())
}

fn validate_text_effects(run: &TextRunPrimitive) -> Result<(), String> {
    if run.text_shadow.is_some() {
        return Err("unsupported text field: textShadow".to_string());
    }
    if run.text_outline {
        return Err("unsupported text field: textOutline".to_string());
    }
    if run.emphasis_mark.is_some() {
        return Err("unsupported text field: emphasisMark".to_string());
    }
    if run.text_effect.is_some() {
        return Err("unsupported text field: textEffect".to_string());
    }
    if run.attrs.modern_effects.is_some() {
        return Err("unsupported text field: modernEffects".to_string());
    }
    Ok(())
}

fn validate_glyph_effects(run: &GlyphRunPrimitive) -> Result<(), String> {
    if run.all_caps {
        return Err("unsupported glyphRun field: allCaps".to_string());
    }
    if run.small_caps {
        return Err("unsupported glyphRun field: smallCaps".to_string());
    }
    if run.text_shadow.is_some() {
        return Err("unsupported glyphRun field: textShadow".to_string());
    }
    if run.text_outline {
        return Err("unsupported glyphRun field: textOutline".to_string());
    }
    if run.emphasis_mark.is_some() {
        return Err("unsupported glyphRun field: emphasisMark".to_string());
    }
    if run.text_effect.is_some() {
        return Err("unsupported glyphRun field: textEffect".to_string());
    }
    if run.attrs.modern_effects.is_some() {
        return Err("unsupported glyphRun field: modernEffects".to_string());
    }
    if run.attrs.leader_glyphs.is_some() {
        return Err("unsupported glyphRun field: leaderGlyphs".to_string());
    }
    Ok(())
}

fn active_leader(leader: &Option<LeaderGlyphMetadata>) -> Option<&LeaderGlyphMetadata> {
    leader.as_ref().filter(|leader| {
        leader
            .glyph
            .as_deref()
            .is_some_and(|glyph| !glyph.is_empty())
            && leader.count.unwrap_or(0) > 0
    })
}

fn paint_leader(
    context: &mut PaintContext<'_, '_>,
    run: &TextRunPrimitive,
    leader: &LeaderGlyphMetadata,
    run_spec: &FontSpec,
    fallback_color: tiny_skia::Color,
    transform: Transform,
) -> Result<(), String> {
    let count = usize::try_from(leader.count.unwrap_or(0))
        .map_err(|_| "leader glyph count is too large".to_string())?;
    if count > MAX_PAINTED_GLYPHS {
        return Err(format!(
            "leader glyph count exceeds the {MAX_PAINTED_GLYPHS} glyph limit"
        ));
    }
    let advance = leader
        .advance
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(0.0);
    if advance <= 0.0 {
        return Ok(());
    }
    let glyph = leader
        .glyph
        .as_deref()
        .ok_or_else(|| "leader glyph is missing".to_string())?;
    let x = leader
        .x
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(number_f32(&run.x)?);
    let baseline = leader
        .baseline_y
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(number_f32(&run.baseline_y)?);
    let mut spec = run_spec.clone();
    if let Some(family) = leader.font.as_deref() {
        spec.family = family.to_string();
    }
    if let Some(size) = &leader.size {
        spec.size = number_f32(size)?;
        if spec.size <= 0.0 {
            return Err("leader glyph size must be positive".to_string());
        }
    }
    let color = leader
        .color
        .as_deref()
        .map(|value| color_with_opacity(value, context.opacity))
        .transpose()?
        .unwrap_or(fallback_color);
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    let direction = leader.rtl == Some(true);
    for index in 0..count {
        let visual_index = if direction { count - 1 - index } else { index };
        let origin = x + visual_index as f32 * advance;
        if let Some(raw) = leader.font_id {
            paint_single_font_text(
                context,
                glyph,
                FontId::from_u32(raw),
                spec.size,
                origin,
                baseline,
                direction,
                None,
                &[],
                &paint,
                transform,
            )?;
        } else {
            paint_resolved_text(
                context,
                glyph,
                origin,
                baseline,
                &spec,
                direction,
                None,
                0.0,
                0.0,
                &[],
                color,
                transform,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn paint_resolved_text(
    context: &mut PaintContext<'_, '_>,
    text: &str,
    x: f32,
    baseline: f32,
    spec: &FontSpec,
    rtl: bool,
    language: Option<&str>,
    letter_spacing: f32,
    word_spacing: f32,
    features: &[ShapeFeature],
    color: tiny_skia::Color,
    transform: Transform,
) -> Result<(), String> {
    let key = format!(
        "{}|{}|{}",
        spec.family.to_lowercase(),
        u8::from(spec.bold),
        u8::from(spec.italic)
    );
    let chain = context
        .resources
        .font_chains
        .get(&key)
        .ok_or_else(|| format!("missing font chain for `{key}`"))?;
    if chain.is_empty() {
        return Err(format!("font chain `{key}` is empty"));
    }
    let segments = font_segments(context.resources.fonts, chain, text)?;
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    let mut pen = x;
    let order: Box<dyn Iterator<Item = usize>> = if rtl {
        Box::new((0..segments.len()).rev())
    } else {
        Box::new(0..segments.len())
    };
    let direction = if rtl {
        ShapeDirection::Rtl
    } else {
        ShapeDirection::Ltr
    };
    let mut painted = 0usize;
    for segment_index in order {
        let segment = &segments[segment_index];
        let segment_text = &text[segment.start..segment.end];
        let glyphs = shape_with_properties(
            context.resources.fonts,
            segment.font,
            segment_text,
            spec.size,
            features,
            direction,
            language,
        )
        .map_err(|error| error.to_string())?;
        painted = painted
            .checked_add(glyphs.len())
            .ok_or_else(|| "painted glyph count overflow".to_string())?;
        if painted > MAX_PAINTED_GLYPHS {
            return Err(format!(
                "text run exceeds the {MAX_PAINTED_GLYPHS} glyph limit"
            ));
        }
        for (glyph_index, glyph) in glyphs.iter().enumerate() {
            let placed = PlacedGlyph {
                id: glyph.glyph_id,
                x: f64::from(pen + glyph.x_offset),
                y: f64::from(baseline - glyph.y_offset),
                cluster: glyph.cluster,
                advance: f64::from(glyph.x_advance),
                logical_order: None,
                bidi_level: None,
            };
            paint_placed_glyph(context, segment.font, &placed, spec.size, &paint, transform)?;
            pen += glyph.x_advance;
            let cluster_ends = glyphs
                .get(glyph_index + 1)
                .is_none_or(|next| next.cluster != glyph.cluster);
            if cluster_ends {
                let source = segment_text
                    .get(glyph.cluster as usize..)
                    .and_then(|tail| tail.chars().next());
                if source == Some(' ') {
                    pen += word_spacing;
                }
                let has_more = glyph_index + 1 < glyphs.len()
                    || if rtl {
                        segment_index > 0
                    } else {
                        segment_index + 1 < segments.len()
                    };
                if has_more {
                    pen += letter_spacing;
                }
            }
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn paint_single_font_text(
    context: &mut PaintContext<'_, '_>,
    text: &str,
    font: FontId,
    size: f32,
    x: f32,
    baseline: f32,
    rtl: bool,
    language: Option<&str>,
    features: &[ShapeFeature],
    paint: &Paint<'_>,
    transform: Transform,
) -> Result<(), String> {
    let direction = if rtl {
        ShapeDirection::Rtl
    } else {
        ShapeDirection::Ltr
    };
    let glyphs = shape_with_properties(
        context.resources.fonts,
        font,
        text,
        size,
        features,
        direction,
        language,
    )
    .map_err(|error| error.to_string())?;
    let mut pen = x;
    for glyph in glyphs {
        let placed = PlacedGlyph {
            id: glyph.glyph_id,
            x: f64::from(pen + glyph.x_offset),
            y: f64::from(baseline - glyph.y_offset),
            cluster: glyph.cluster,
            advance: f64::from(glyph.x_advance),
            logical_order: None,
            bidi_level: None,
        };
        paint_placed_glyph(context, font, &placed, size, paint, transform)?;
        pen += glyph.x_advance;
    }
    Ok(())
}

fn paint_placed_glyph(
    context: &mut PaintContext<'_, '_>,
    font: FontId,
    glyph: &PlacedGlyph,
    size: f32,
    paint: &Paint<'_>,
    transform: Transform,
) -> Result<(), String> {
    let cached = cached_glyph(context.cache, context.resources.fonts, font, glyph.id)?;
    let Some(path) = &cached.path else {
        return Ok(());
    };
    let scale = size / cached.upem;
    let glyph_transform = Transform::from_row(
        scale,
        0.0,
        0.0,
        -scale,
        f64_to_f32(glyph.x, "glyph x")?,
        f64_to_f32(glyph.y, "glyph y")?,
    );
    context.pixmap.fill_path(
        path,
        paint,
        FillRule::Winding,
        transform.pre_concat(glyph_transform),
        context.mask,
    );
    Ok(())
}

fn cached_glyph<'a>(
    cache: &'a mut GlyphCache,
    fonts: &FontStore,
    font: FontId,
    glyph_id: u32,
) -> Result<&'a CachedGlyph, String> {
    let key = (font.to_u32(), glyph_id);
    match cache.entries.entry(key) {
        Entry::Occupied(entry) => Ok(entry.into_mut()),
        Entry::Vacant(entry) => {
            let id = u16::try_from(glyph_id)
                .map_err(|_| format!("glyph id {glyph_id} exceeds the font outline range"))?;
            let outline = fonts
                .outline_glyph(font, id)
                .map_err(|error| error.to_string())?;
            let path = outline_path(&outline.cmds)?;
            Ok(entry.insert(CachedGlyph {
                path,
                upem: f32::from(outline.upem),
            }))
        }
    }
}

fn outline_path(commands: &[PathCmd]) -> Result<Option<Path>, String> {
    if commands.is_empty() {
        return Ok(None);
    }
    let mut builder = PathBuilder::new();
    for command in commands {
        match *command {
            PathCmd::MoveTo { x, y } => builder.move_to(x, y),
            PathCmd::LineTo { x, y } => builder.line_to(x, y),
            PathCmd::QuadTo { cx, cy, x, y } => builder.quad_to(cx, cy, x, y),
            PathCmd::CubicTo {
                c1x,
                c1y,
                c2x,
                c2y,
                x,
                y,
            } => builder.cubic_to(c1x, c1y, c2x, c2y, x, y),
            PathCmd::Close => builder.close(),
        }
    }
    builder
        .finish()
        .map(Some)
        .ok_or_else(|| "glyph outline has no drawable commands".to_string())
}

fn font_segments(
    fonts: &FontStore,
    chain: &[FontId],
    text: &str,
) -> Result<Vec<FontSegment>, String> {
    let mut segments: Vec<FontSegment> = Vec::new();
    for (start, character) in text.char_indices() {
        let end = start + character.len_utf8();
        let font = fonts
            .resolve(chain, character)
            .or_else(|| chain.last().copied())
            .ok_or_else(|| "font chain is empty".to_string())?;
        if let Some(last) = segments.last_mut()
            && last.font == font
        {
            last.end = end;
        } else {
            segments.push(FontSegment { start, end, font });
        }
    }
    Ok(segments)
}

fn parse_font(font: &str) -> Result<FontSpec, String> {
    let mut size_match = None;
    let mut cursor = 0usize;
    for token in font.split_whitespace() {
        let relative = font[cursor..]
            .find(token)
            .ok_or_else(|| format!("unsupported CSS font shorthand: {font}"))?;
        let index = cursor + relative;
        cursor = index + token.len();
        if token.strip_suffix("px").is_some() {
            size_match = Some((index, token));
            break;
        }
    }
    let (size_index, size_token) =
        size_match.ok_or_else(|| format!("unsupported CSS font shorthand: {font}"))?;
    let size = size_token
        .strip_suffix("px")
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value > 0.0)
        .ok_or_else(|| format!("unsupported CSS font size: {size_token}"))?;
    let mut bold = false;
    let mut italic = false;
    let mut small_caps = false;
    for token in font[..size_index].split_whitespace() {
        match token {
            "normal" => {}
            "italic" | "oblique" => italic = true,
            "small-caps" => small_caps = true,
            "bold" | "bolder" => bold = true,
            value => {
                let weight = value
                    .parse::<u16>()
                    .map_err(|_| format!("unsupported CSS font token: {value}"))?;
                if !(1..=1000).contains(&weight) {
                    return Err(format!("unsupported CSS font weight: {value}"));
                }
                bold = weight >= 600;
            }
        }
    }
    let family_start = size_index + size_token.len();
    let family_list = font[family_start..].trim();
    let family = first_font_family(family_list)
        .trim()
        .trim_matches(['\'', '"'])
        .trim()
        .to_string();
    if family.is_empty() {
        return Err(format!("CSS font shorthand has no family: {font}"));
    }
    Ok(FontSpec {
        family,
        size,
        bold,
        italic,
        small_caps,
    })
}

fn first_font_family(value: &str) -> &str {
    let mut quote = None;
    for (index, character) in value.char_indices() {
        match character {
            '\'' | '"' if quote == Some(character) => quote = None,
            '\'' | '"' if quote.is_none() => quote = Some(character),
            ',' if quote.is_none() => return &value[..index],
            _ => {}
        }
    }
    value
}

fn text_rect(run: &TextRunPrimitive, size: f32) -> Result<FRect, String> {
    Ok(FRect {
        x: number_f32(&run.x)?,
        y: number_f32(&run.baseline_y)? - size * 0.8,
        w: number_f32(&run.width)?.max(0.0),
        h: size * 1.2,
    })
}

fn glyph_rect(run: &GlyphRunPrimitive) -> Result<FRect, String> {
    let first = run
        .glyphs
        .first()
        .ok_or_else(|| "glyph run is empty".to_string())?;
    let mut left = f32::INFINITY;
    let mut right = f32::NEG_INFINITY;
    for glyph in &run.glyphs {
        let x = f64_to_f32(glyph.x, "glyph x")?;
        let advance = f64_to_f32(glyph.advance, "glyph advance")?;
        left = left.min(x);
        right = right.max(x + advance);
    }
    let size = run.size as f32;
    Ok(FRect {
        x: left,
        y: f64_to_f32(first.y, "glyph y")? - size * 0.8,
        w: (right - left).max(0.0),
        h: size * 1.2,
    })
}

fn f64_to_f32(value: f64, field: &str) -> Result<f32, String> {
    let value = value as f32;
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| format!("{field} is outside the raster coordinate range"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_layout_font_shorthands() {
        let regular = parse_font("400 16px Liberation Sans, sans-serif").unwrap();
        assert_eq!(regular.family, "Liberation Sans");
        assert_eq!(regular.size, 16.0);
        assert!(!regular.bold);
        assert!(!regular.italic);

        let styled =
            parse_font("italic small-caps 700 12.5px \"Aptos Display\", sans-serif").unwrap();
        assert_eq!(styled.family, "Aptos Display");
        assert_eq!(styled.size, 12.5);
        assert!(styled.bold);
        assert!(styled.italic);
        assert!(styled.small_caps);
    }
}
