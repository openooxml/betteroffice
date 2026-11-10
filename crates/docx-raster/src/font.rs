//! Shared-store text shaping and glyph rasterization.

use std::collections::{BTreeMap, HashMap};
use std::rc::Rc;

use docx_layout::display_list::{
    ClipRect, GlyphRunPrimitive, LeaderGlyphMetadata, PlacedGlyph, TextRunPrimitive,
};
use ooxml_text::{
    FontId, FontStore, PathCmd, ShapeDirection, ShapeFeature, ShapedGlyph, shape,
    shape::shape_with_properties,
};
use tiny_skia::{FillRule, Mask, Paint, Path, PathBuilder, Pixmap, Rect, Transform};

use crate::shadow::{self, Shadow, TextShadow};
use crate::{
    FRect, PageBudget, RenderResources, color_with_opacity, mask_bytes, number_f32,
    primitive_visual_transform, skip_effect,
};

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

/// Store-bound glyph outlines reusable across pages.
#[derive(Default)]
pub struct GlyphCache {
    inner: ooxml_raster::GlyphCache,
}

#[derive(Clone)]
struct FontSpec {
    family: String,
    size: f32,
    bold: bool,
    italic: bool,
    small_caps: bool,
}

/// A parsed shorthand plus the font-chain key its family resolves under.
struct ParsedFont {
    spec: FontSpec,
    chain_key: String,
}

/// Distinct shorthands one page keeps parsed, and the raw bytes their keys may
/// copy out of the display list. A `font` string is untrusted and unbounded, so
/// the cache stops keeping them rather than growing with the page.
const MAX_CACHED_FONT_SPECS: usize = 256;
const MAX_FONT_SPEC_KEY_BYTES: usize = 65_536;

/// Parsed font shorthands keyed by the raw display-list string.
#[derive(Default)]
pub(crate) struct FontSpecCache {
    entries: HashMap<Box<str>, Rc<ParsedFont>>,
    key_bytes: usize,
}

impl FontSpecCache {
    fn font(&mut self, font: &str) -> Result<Rc<ParsedFont>, String> {
        if let Some(parsed) = self.entries.get(font) {
            return Ok(parsed.clone());
        }
        let spec = parse_font(font)?;
        let parsed = Rc::new(ParsedFont {
            chain_key: chain_key(&spec),
            spec,
        });
        self.remember(font, &parsed);
        Ok(parsed)
    }

    /// Keeps one parse against its raw shorthand, while the map has room. The
    /// key is a copy of display-list bytes, so it is the copy that is budgeted.
    fn remember(&mut self, font: &str, parsed: &Rc<ParsedFont>) {
        if self.entries.len() >= MAX_CACHED_FONT_SPECS
            || self.key_bytes + font.len() > MAX_FONT_SPEC_KEY_BYTES
        {
            return;
        }
        self.key_bytes += font.len();
        self.entries.insert(font.into(), parsed.clone());
    }
}

/// The built-in faces the layout names when only they cover a run, for hosts
/// that registered no chain under their family.
fn backstop_chain(spec: &FontSpec, key: &str) -> Result<[FontId; 1], String> {
    if spec
        .family
        .eq_ignore_ascii_case(ooxml_text::BASE_LATIN_FAMILY)
    {
        Ok([FontId::base_latin(spec.bold, spec.italic)])
    } else {
        Err(format!("missing font chain for `{key}`"))
    }
}

fn chain_key(spec: &FontSpec) -> String {
    format!(
        "{}|{}|{}",
        spec.family.to_lowercase(),
        u8::from(spec.bold),
        u8::from(spec.italic)
    )
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
    pub specs: &'a mut FontSpecCache,
    pub budget: &'a mut PageBudget,
    pub base_transform: Transform,
    pub mask: Option<&'a Mask>,
    pub retained_mask_bytes: u64,
    pub opacity: f32,
    pub skipped: &'a mut BTreeMap<String, usize>,
}

impl<'a, 'b> PaintContext<'a, 'b> {
    fn reborrow_with_mask<'c>(&'c mut self, mask: &'c Mask) -> PaintContext<'c, 'b> {
        PaintContext {
            pixmap: self.pixmap,
            resources: self.resources,
            cache: self.cache,
            specs: self.specs,
            budget: self.budget,
            base_transform: self.base_transform,
            mask: Some(mask),
            retained_mask_bytes: self.retained_mask_bytes,
            opacity: self.opacity,
            skipped: self.skipped,
        }
    }

    /// The same resources, painting into `surface` without a clip.
    fn on_surface<'c>(&'c mut self, surface: &'c mut Pixmap) -> PaintContext<'c, 'b> {
        PaintContext {
            pixmap: surface,
            resources: self.resources,
            cache: self.cache,
            specs: self.specs,
            budget: self.budget,
            base_transform: self.base_transform,
            mask: None,
            retained_mask_bytes: self.retained_mask_bytes,
            opacity: self.opacity,
            skipped: self.skipped,
        }
    }

    /// Paints `shadow` under what `draw` paints, by running `draw` once into a
    /// scratch surface around `rect` (in `transform`'s space).
    fn paint_shadow(
        &mut self,
        shadow: Option<Shadow>,
        rect: FRect,
        transform: Transform,
        draw: impl FnOnce(&mut PaintContext<'_, 'b>, Transform) -> Result<f32, String>,
    ) -> Result<(), String> {
        let Some(shadow) = shadow else {
            return Ok(());
        };
        let Some(bounds) = Rect::from_xywh(rect.x, rect.y, rect.w.max(1.0), rect.h.max(1.0))
            .and_then(|rect| rect.transform(transform))
        else {
            return Ok(());
        };
        let page = (self.pixmap.width(), self.pixmap.height());
        let clip = self
            .mask
            .map_or(0, |mask| mask_bytes(mask.width(), mask.height()));
        let retained = self.retained_mask_bytes.saturating_add(clip);
        let Some(mut surface) =
            shadow::shadow_surface(page, bounds, &shadow, self.budget, retained)?
        else {
            return Ok(());
        };
        let offset = surface.offset;
        draw(
            &mut self.on_surface(&mut surface.pixmap),
            offset.pre_concat(transform),
        )?;
        shadow::composite_shadow(self.pixmap, self.mask, surface, &shadow);
        Ok(())
    }
}

pub(crate) fn paint_text(
    context: &mut PaintContext<'_, '_>,
    run: &TextRunPrimitive,
) -> Result<(), String> {
    let mask = paint_clip_mask(context, run.paint_clip.as_ref())?;
    if let Some(mask) = &mask {
        paint_text_core(&mut context.reborrow_with_mask(mask), run)
    } else {
        paint_text_core(context, run)
    }
}

fn paint_text_core(
    context: &mut PaintContext<'_, '_>,
    run: &TextRunPrimitive,
) -> Result<(), String> {
    skip_text_effects(
        context.skipped,
        "text",
        run.text_outline,
        run.emphasis_mark.is_some(),
        run.text_effect.is_some(),
        run.attrs.modern_effects.is_some(),
    );
    if run.text.is_empty() {
        return Ok(());
    }
    let parsed = context.specs.font(&run.font)?;
    let small_caps = (run.small_caps || parsed.spec.small_caps) && !run.all_caps;
    let rect = text_rect(run, parsed.spec.size)?;
    let visual = primitive_visual_transform(
        rect,
        run.rotation_deg.as_ref(),
        run.horizontal_scale.as_ref(),
    )?;
    let transform = context.base_transform.pre_concat(visual);
    let color = color_with_opacity(&run.color, context.opacity)?;
    if let Some(leader) = active_leader(run.attrs.leader_glyphs.as_deref()) {
        return paint_leader(context, run, leader, &parsed, color, transform);
    }
    context.budget.charge_glyphs(painted_chars(run))?;
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
    let x = number_f32(&run.x)?;
    let baseline = number_f32(&run.baseline_y)?;
    let draw = |context: &mut PaintContext<'_, '_>, transform: Transform| {
        paint_cased_text(
            context,
            &text,
            x,
            baseline,
            &parsed.spec,
            &parsed.chain_key,
            run.rtl == Some(true),
            run.attrs.lang.as_deref(),
            letter_spacing,
            word_spacing,
            small_caps,
            color,
            transform,
        )
    };
    let shadow = drawn_text_shadow(
        context.skipped,
        "text",
        run.text_shadow.as_deref(),
        parsed.spec.size,
    );
    context.paint_shadow(shadow, rect, transform, draw)?;
    draw(context, transform).map(|_| ())
}

/// A run of text painted in one size: synthesized small caps or as is. An
/// expanded segment is one lowercase character whose uppercase form is several
/// (`ß` to `SS`), so it advances as one character.
struct CaseSegment {
    text: String,
    small: bool,
    expanded: bool,
}

/// Paints text and returns its advance. Small caps are drawn the way the web
/// canvas synthesizes them for a face without `smcp`: a lowercase letter as its
/// uppercase form at [`ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE`] of the
/// size, everything else as is, which is the advance layout measured.
#[allow(clippy::too_many_arguments)]
fn paint_cased_text(
    context: &mut PaintContext<'_, '_>,
    text: &str,
    x: f32,
    baseline: f32,
    spec: &FontSpec,
    key: &str,
    rtl: bool,
    language: Option<&str>,
    letter_spacing: f32,
    word_spacing: f32,
    small_caps: bool,
    color: tiny_skia::Color,
    transform: Transform,
) -> Result<f32, String> {
    if !small_caps {
        return paint_resolved_text(
            context,
            text,
            x,
            baseline,
            spec,
            key,
            rtl,
            language,
            letter_spacing,
            word_spacing,
            &[],
            color,
            transform,
        );
    }
    let mut segments: Vec<CaseSegment> = Vec::new();
    for ch in text.chars() {
        if !ch.is_lowercase() {
            match segments.last_mut() {
                Some(segment) if !segment.small => segment.text.push(ch),
                _ => segments.push(CaseSegment {
                    text: ch.to_string(),
                    small: false,
                    expanded: false,
                }),
            }
            continue;
        }
        let upper = ooxml_text::uppercase_for_language(ch, language);
        if upper.len() > 1 {
            segments.push(CaseSegment {
                text: upper.into_iter().collect(),
                small: true,
                expanded: true,
            });
            continue;
        }
        match segments.last_mut() {
            Some(segment) if segment.small && !segment.expanded => segment.text.extend(upper),
            _ => segments.push(CaseSegment {
                text: upper.into_iter().collect(),
                small: true,
                expanded: false,
            }),
        }
    }
    if rtl {
        segments.reverse();
    }
    let small_spec = FontSpec {
        size: spec.size * ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE,
        ..spec.clone()
    };
    let mut pen = x;
    let count = segments.len();
    for (index, segment) in segments.into_iter().enumerate() {
        let spacing = if segment.expanded {
            0.0
        } else {
            letter_spacing
        };
        pen += paint_resolved_text(
            context,
            &segment.text,
            pen,
            baseline,
            if segment.small { &small_spec } else { spec },
            key,
            rtl,
            language,
            spacing,
            word_spacing,
            &[],
            color,
            transform,
        )?;
        if index + 1 < count {
            pen += letter_spacing;
        }
    }
    Ok(pen - x)
}

/// The shadow a `textShadow` value draws, counting a value no backend draws yet.
fn drawn_text_shadow(
    skipped: &mut BTreeMap<String, usize>,
    primitive: &str,
    value: Option<&str>,
    size: f32,
) -> Option<Shadow> {
    match shadow::text_shadow(value, size) {
        TextShadow::Drawn(shadow) => Some(shadow),
        TextShadow::Undrawn => {
            skip_effect(skipped, &format!("{primitive}.textShadow"));
            None
        }
        TextShadow::None => None,
    }
}

/// Counts the text effects this backend leaves out; the web canvas draws them.
fn skip_text_effects(
    skipped: &mut BTreeMap<String, usize>,
    primitive: &str,
    outline: bool,
    emphasis_mark: bool,
    text_effect: bool,
    modern_effects: bool,
) {
    for (present, field) in [
        (outline, "textOutline"),
        (emphasis_mark, "emphasisMark"),
        (text_effect, "textEffect"),
        (modern_effects, "modernEffects"),
    ] {
        if present {
            skip_effect(skipped, &format!("{primitive}.{field}"));
        }
    }
}

pub(crate) fn paint_glyph_run(
    context: &mut PaintContext<'_, '_>,
    run: &GlyphRunPrimitive,
) -> Result<(), String> {
    let mask = paint_clip_mask(context, run.paint_clip.as_ref())?;
    if let Some(mask) = &mask {
        paint_glyph_run_core(&mut context.reborrow_with_mask(mask), run)
    } else {
        paint_glyph_run_core(context, run)
    }
}

fn paint_glyph_run_core(
    context: &mut PaintContext<'_, '_>,
    run: &GlyphRunPrimitive,
) -> Result<(), String> {
    skip_text_effects(
        context.skipped,
        "glyphRun",
        run.text_outline,
        run.emphasis_mark.is_some(),
        run.text_effect.is_some(),
        run.attrs.modern_effects.is_some(),
    );
    if run.attrs.leader_glyphs.is_some() {
        skip_effect(context.skipped, "glyphRun.leaderGlyphs");
    }
    context.budget.charge_glyphs(run.glyphs.len() as u64)?;
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
    let draw = |context: &mut PaintContext<'_, '_>, transform: Transform| {
        for glyph in &run.glyphs {
            paint_placed_glyph(context, font, glyph, run.size as f32, &paint, transform)?;
        }
        Ok(0.0)
    };
    let shadow = drawn_text_shadow(
        context.skipped,
        "glyphRun",
        run.text_shadow.as_deref(),
        run.size as f32,
    );
    context.paint_shadow(shadow, rect, transform, draw)?;
    draw(context, transform).map(|_| ())
}

fn paint_clip_mask(
    context: &mut PaintContext<'_, '_>,
    clip: Option<&ClipRect>,
) -> Result<Option<Mask>, String> {
    let Some(clip) = clip else {
        return Ok(None);
    };
    debug_assert!(context.base_transform.kx == 0.0 && context.base_transform.ky == 0.0);
    let width = context.pixmap.width();
    let height = context.pixmap.height();
    let bytes = mask_bytes(width, height);
    context
        .budget
        .charge_masks(context.retained_mask_bytes.saturating_add(bytes))?;
    let mut mask =
        Mask::new(width, height).ok_or_else(|| "invalid text paint mask size".to_string())?;
    let x = clip.x.as_ref().map(number_f32).transpose()?.unwrap_or(0.0);
    let slot_width = clip
        .w
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(0.0)
        .max(0.0);
    let scale = f64::from(context.base_transform.sx);
    let offset = f64::from(context.base_transform.tx);
    let left = (f64::from(x) * scale + offset).clamp(0.0, f64::from(width));
    let right =
        ((f64::from(x) + f64::from(slot_width)) * scale + offset).clamp(0.0, f64::from(width));
    if right > left {
        let rect = Rect::from_xywh(left as f32, 0.0, (right - left) as f32, height as f32)
            .ok_or_else(|| "invalid text paint clip rectangle".to_string())?;
        mask.fill_path(
            &PathBuilder::from_rect(rect),
            FillRule::Winding,
            true,
            Transform::identity(),
        );
    }
    if let Some(outer) = context.mask {
        for (value, clip) in mask.data_mut().iter_mut().zip(outer.data()) {
            *value = ((u16::from(*value) * u16::from(*clip) + 127) / 255) as u8;
        }
    }
    Ok(Some(mask))
}

/// Characters a run paints. `allCaps` uppercases first, and one character can
/// uppercase into three, so the charge is counted without building the string.
fn painted_chars(run: &TextRunPrimitive) -> u64 {
    if run.all_caps {
        run.text
            .chars()
            .map(|character| character.to_uppercase().count() as u64)
            .sum()
    } else {
        run.text.chars().count() as u64
    }
}

fn active_leader(leader: Option<&LeaderGlyphMetadata>) -> Option<&LeaderGlyphMetadata> {
    leader.filter(|leader| {
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
    run_font: &ParsedFont,
    fallback_color: tiny_skia::Color,
    transform: Transform,
) -> Result<(), String> {
    let count = usize::try_from(leader.count.unwrap_or(0))
        .map_err(|_| "leader glyph count is too large".to_string())?;
    let glyph = leader
        .glyph
        .as_deref()
        .ok_or_else(|| "leader glyph is missing".to_string())?;
    context
        .budget
        .charge_glyphs((count as u64).saturating_mul(glyph.chars().count() as u64))?;
    let advance = leader
        .advance
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(0.0);
    if advance <= 0.0 {
        return Ok(());
    }
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
    let mut spec = run_font.spec.clone();
    if let Some(font) = leader.font.as_deref() {
        match parse_font(font) {
            Ok(parsed) => {
                spec.family = parsed.family;
                spec.bold = parsed.bold;
                spec.italic = parsed.italic;
            }
            Err(_) => spec.family = font.to_string(),
        }
    }
    if let Some(size) = &leader.size {
        spec.size = number_f32(size)?;
        if spec.size <= 0.0 {
            return Err("leader glyph size must be positive".to_string());
        }
    }
    let chain_key = match leader.font.as_deref() {
        Some(_) => chain_key(&spec),
        None => run_font.chain_key.clone(),
    };
    let color = leader
        .color
        .as_deref()
        .map(|value| color_with_opacity(value, context.opacity))
        .transpose()?
        .unwrap_or(fallback_color);
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    let rtl = leader.rtl == Some(true);
    let direction = if rtl {
        ShapeDirection::Rtl
    } else {
        ShapeDirection::Ltr
    };
    let stamp = shape_stamp(
        context,
        glyph,
        &spec,
        &chain_key,
        direction,
        count,
        leader.font_id,
    )?;
    for index in 0..count {
        let visual_index = if rtl { count - 1 - index } else { index };
        let origin = x + visual_index as f32 * advance;
        paint_shaped_stamp(context, &stamp, origin, baseline, &paint, transform)?;
    }
    Ok(())
}

/// One shaped leader text, ready to be stamped at every repeat position.
struct ShapedStamp {
    size: f32,
    segments: Vec<ShapedSegment>,
}

struct ShapedSegment {
    font: FontId,
    glyphs: Vec<ShapedGlyph>,
}

/// Shapes the leader text once in the order its glyphs are painted, charging
/// what every repeat will paint. Shaping never depends on the pen origin, so
/// one pass covers every repeat.
#[allow(clippy::too_many_arguments)]
fn shape_stamp(
    context: &mut PaintContext<'_, '_>,
    glyph: &str,
    spec: &FontSpec,
    key: &str,
    direction: ShapeDirection,
    count: usize,
    font_id: Option<u32>,
) -> Result<ShapedStamp, String> {
    let mut segments = Vec::new();
    let mut excess = 0u64;
    if let Some(raw) = font_id {
        let font = FontId::from_u32(raw);
        let glyphs = shape_segment(context, glyph, font, spec.size, direction)?;
        excess = (glyphs.len() as u64).saturating_sub(glyph.chars().count() as u64);
        segments.push(ShapedSegment { font, glyphs });
    } else {
        let backstop;
        let chain: &[FontId] = match context.resources.font_chains.get(key) {
            Some(chain) => chain.as_slice(),
            None => {
                backstop = backstop_chain(spec, key)?;
                &backstop
            }
        };
        if chain.is_empty() {
            return Err(format!("font chain `{key}` is empty"));
        }
        for segment in font_segments(
            context.resources.fonts,
            chain,
            glyph,
            spec.bold,
            spec.italic,
            &spec.family,
        )? {
            let text = &glyph[segment.start..segment.end];
            let glyphs = shape_segment(context, text, segment.font, spec.size, direction)?;
            excess = excess
                .saturating_add((glyphs.len() as u64).saturating_sub(text.chars().count() as u64));
            segments.push(ShapedSegment {
                font: segment.font,
                glyphs,
            });
        }
        if matches!(direction, ShapeDirection::Rtl) {
            segments.reverse();
        }
    }
    context
        .budget
        .charge_glyphs(excess.saturating_mul(count as u64))?;
    Ok(ShapedStamp {
        size: spec.size,
        segments,
    })
}

fn shape_segment(
    context: &PaintContext<'_, '_>,
    text: &str,
    font: FontId,
    size: f32,
    direction: ShapeDirection,
) -> Result<Vec<ShapedGlyph>, String> {
    shape_with_properties(
        context.resources.fonts,
        font,
        text,
        size,
        &[],
        direction,
        None,
    )
    .map_err(|error| error.to_string())
}

/// Replays one shaping at a stamp origin, painting exactly what a fresh shape
/// at that origin would.
fn paint_shaped_stamp(
    context: &mut PaintContext<'_, '_>,
    stamp: &ShapedStamp,
    origin: f32,
    baseline: f32,
    paint: &Paint<'_>,
    transform: Transform,
) -> Result<(), String> {
    let mut pen = origin;
    for segment in &stamp.segments {
        for glyph in &segment.glyphs {
            let placed = PlacedGlyph {
                id: glyph.glyph_id,
                x: f64::from(pen + glyph.x_offset),
                y: f64::from(baseline - glyph.y_offset),
                cluster: glyph.cluster,
                advance: f64::from(glyph.x_advance),
                logical_order: None,
                bidi_level: None,
            };
            paint_placed_glyph(context, segment.font, &placed, stamp.size, paint, transform)?;
            pen += glyph.x_advance;
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
    key: &str,
    rtl: bool,
    language: Option<&str>,
    letter_spacing: f32,
    word_spacing: f32,
    features: &[ShapeFeature],
    color: tiny_skia::Color,
    transform: Transform,
) -> Result<f32, String> {
    let backstop;
    let chain: &[FontId] = match context.resources.font_chains.get(key) {
        Some(chain) => chain.as_slice(),
        None => {
            backstop = backstop_chain(spec, key)?;
            &backstop
        }
    };
    if chain.is_empty() {
        return Err(format!("font chain `{key}` is empty"));
    }
    let segments = font_segments(
        context.resources.fonts,
        chain,
        text,
        spec.bold,
        spec.italic,
        &spec.family,
    )?;
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
        charge_shaped_excess(context, segment_text, glyphs.len())?;
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
    Ok(pen - x)
}

/// Shaping can place more glyphs than the text had characters. The caller
/// charged the characters before the shaper ran, so only the excess is left.
fn charge_shaped_excess(
    context: &mut PaintContext<'_, '_>,
    text: &str,
    glyphs: usize,
) -> Result<(), String> {
    let charged = text.chars().count() as u64;
    context
        .budget
        .charge_glyphs((glyphs as u64).saturating_sub(charged))
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
    ooxml_raster::fill_glyph(
        context.pixmap,
        path,
        scale,
        f64_to_f32(glyph.x, "glyph x")?,
        f64_to_f32(glyph.y, "glyph y")?,
        transform,
        paint,
        context.mask,
    );
    Ok(())
}

fn cached_glyph<'a>(
    cache: &'a mut GlyphCache,
    fonts: &FontStore,
    font: FontId,
    glyph_id: u32,
) -> Result<&'a ooxml_raster::CachedGlyph, String> {
    cache.inner.lookup(
        fonts,
        font,
        glyph_id,
        || {
            u16::try_from(glyph_id)
                .map_err(|_| format!("glyph id {glyph_id} exceeds the font outline range"))
        },
        outline_path,
    )
}

fn outline_path(commands: &[PathCmd]) -> Result<Option<Path>, String> {
    if commands.is_empty() {
        return Ok(None);
    }
    ooxml_raster::outline_path(commands)
        .map(Some)
        .ok_or_else(|| "glyph outline has no drawable commands".to_string())
}

fn font_segments(
    fonts: &FontStore,
    chain: &[FontId],
    text: &str,
    bold: bool,
    italic: bool,
    family: &str,
) -> Result<Vec<FontSegment>, String> {
    let mut segments: Vec<FontSegment> = Vec::new();
    for (start, character) in text.char_indices() {
        let end = start + character.len_utf8();
        let font = fonts
            .resolve_with_family(chain, character, bold, italic, family)
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
    let family = docx_layout::measure_blocks::css_primary_family(family_list).to_string();
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
    fn small_caps_advance_by_layouts_synthesized_measure() {
        let mut fonts = FontStore::new();
        let id = fonts
            .register(include_bytes!("../tests/assets/Carlito-Regular.ttf").to_vec())
            .unwrap();
        let chains = crate::FontChains::from([("carlito|0|0".to_string(), vec![id])]);
        let images = crate::ImageMap::new();
        let resources = RenderResources::new(&fonts, &chains, &images);
        let mut pixmap = Pixmap::new(200, 40).unwrap();
        let mut cache = GlyphCache::default();
        let mut specs = FontSpecCache::default();
        let mut budget = PageBudget::new(200, 40);
        let mut skipped = BTreeMap::new();
        let mut context = PaintContext {
            pixmap: &mut pixmap,
            resources: &resources,
            cache: &mut cache,
            specs: &mut specs,
            budget: &mut budget,
            base_transform: Transform::identity(),
            mask: None,
            retained_mask_bytes: 0,
            opacity: 1.0,
            skipped: &mut skipped,
        };
        let parsed = context.specs.font("16px Carlito").unwrap();
        let advance = paint_cased_text(
            &mut context,
            "Small Caps",
            2.0,
            30.0,
            &parsed.spec,
            &parsed.chain_key,
            false,
            None,
            0.0,
            0.0,
            true,
            tiny_skia::Color::BLACK,
            Transform::identity(),
        )
        .unwrap();
        let measure = |text: &str| measure_text(&fonts, id, text, 16.0).unwrap();
        let scale = ooxml_text::BROWSER_SMALL_CAPS_ADVANCE_SCALE;
        let expected =
            measure("S") + measure("MALL") * scale + measure(" C") + measure("APS") * scale;
        assert!(
            (advance - expected).abs() < 0.01,
            "{advance} against {expected}"
        );
        assert!(advance < measure("SMALL CAPS"));
        let sharp = paint_cased_text(
            &mut context,
            "\u{df}",
            2.0,
            30.0,
            &parsed.spec,
            &parsed.chain_key,
            false,
            None,
            2.0,
            0.0,
            true,
            tiny_skia::Color::BLACK,
            Transform::identity(),
        )
        .unwrap();
        assert!(
            (sharp - measure("SS") * scale).abs() < 0.01,
            "one character, no spacing inside"
        );
    }

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

    #[test]
    fn spec_cache_reuses_one_parse_per_raw_string() {
        let mut cache = FontSpecCache::default();
        let first = cache
            .font("400 16px Liberation Sans, sans-serif")
            .expect("parse");
        let second = cache
            .font("400 16px Liberation Sans, sans-serif")
            .expect("parse");
        assert!(Rc::ptr_eq(&first, &second));
        assert_eq!(first.spec.family, "Liberation Sans");
        assert_eq!(first.chain_key, "liberation sans|0|0");
        assert_eq!(cache.entries.len(), 1);

        let bold = cache
            .font("700 16px Liberation Sans, sans-serif")
            .expect("parse");
        assert!(!Rc::ptr_eq(&bold, &first));
        assert!(bold.spec.bold && !bold.spec.italic);
        assert_ne!(bold.chain_key, first.chain_key);

        let italic = cache
            .font("italic 16px Liberation Sans, sans-serif")
            .expect("parse");
        assert!(italic.spec.italic && !italic.spec.bold);
        assert_ne!(italic.chain_key, bold.chain_key);
        assert_eq!(cache.entries.len(), 3);

        assert!(cache.font("16 Liberation Sans").is_err());
        assert!(cache.font("px").is_err());
        assert_eq!(cache.entries.len(), 3);
    }

    /// Shorthands are untrusted display-list strings, so the cache parses every
    /// one but stops copying them once either budget is spent.
    #[test]
    fn the_spec_cache_stops_keeping_shorthands_past_its_budgets() {
        let mut counted = FontSpecCache::default();
        for index in 0..MAX_CACHED_FONT_SPECS + 8 {
            let font = format!("400 {}12px Carlito, sans-serif", " ".repeat(index));
            assert_eq!(counted.font(&font).expect("parse").spec.family, "Carlito");
        }
        assert_eq!(counted.entries.len(), MAX_CACHED_FONT_SPECS);
        assert!(counted.key_bytes <= MAX_FONT_SPEC_KEY_BYTES);

        let mut oversized = FontSpecCache::default();
        let padding = " ".repeat(MAX_FONT_SPEC_KEY_BYTES);
        for index in 0..4 {
            let font = format!("400 {padding}{}12px Carlito, sans-serif", " ".repeat(index));
            assert_eq!(oversized.font(&font).expect("parse").spec.family, "Carlito");
        }
        assert!(oversized.entries.is_empty());
        assert_eq!(oversized.key_bytes, 0);
    }
}

#[cfg(test)]
mod shared_helper_tests {
    use super::*;

    #[test]
    fn outline_keeps_empty_and_undrawable_results() {
        assert!(outline_path(&[]).unwrap().is_none());
        assert_eq!(
            outline_path(&[PathCmd::MoveTo { x: 0.0, y: 0.0 }]).unwrap_err(),
            "glyph outline has no drawable commands"
        );
    }
}
