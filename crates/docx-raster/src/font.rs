//! Shared-store text shaping and glyph rasterization.

use std::collections::HashMap;
use std::collections::VecDeque;
use std::rc::Rc;

use docx_layout::display_list::{
    ClipRect, GlyphRunPrimitive, LeaderGlyphMetadata, PlacedGlyph, TextRunPrimitive,
};
use ooxml_text::{
    FontId, FontStore, PathCmd, ShapeDirection, ShapeFeature, ShapedGlyph, shape,
    shape::shape_with_properties,
};
use tiny_skia::{FillRule, Mask, Paint, Path, PathBuilder, Pixmap, Rect, Transform};

use crate::{
    FRect, PageBudget, RenderResources, color_with_opacity, mask_bytes, number_f32,
    primitive_visual_transform,
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

/// Most glyph outlines one cache retains; the oldest insertion evicts first.
const MAX_CACHED_GLYPH_OUTLINES: usize = 8_192;

/// Extracted glyph outlines keyed by font and glyph id, reusable across pages.
/// Ids are store-local, so the cache binds to the store that fills it and
/// refuses any other; entries past the cap evict oldest-first.
#[derive(Default)]
pub struct GlyphCache {
    entries: HashMap<(u32, u32), CachedGlyph>,
    order: VecDeque<(u32, u32)>,
    store: Option<u64>,
}

impl GlyphCache {
    fn remember(&mut self, key: (u32, u32), glyph: CachedGlyph) {
        if self.entries.insert(key, glyph).is_none() {
            self.order.push_back(key);
        }
        while self.order.len() > MAX_CACHED_GLYPH_OUTLINES {
            let oldest = self
                .order
                .pop_front()
                .expect("the queue tracks every cached entry");
            self.entries.remove(&oldest);
        }
    }
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
        }
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
    validate_text_effects(run)?;
    if run.text.is_empty() {
        return Ok(());
    }
    let parsed = context.specs.font(&run.font)?;
    if run.small_caps || parsed.spec.small_caps {
        return Err("unsupported text field: smallCaps".to_string());
    }
    let rect = text_rect(run, parsed.spec.size)?;
    let visual = primitive_visual_transform(
        rect,
        run.rotation_deg.as_ref(),
        run.horizontal_scale.as_ref(),
    )?;
    let transform = context.base_transform.pre_concat(visual);
    let color = color_with_opacity(&run.color, context.opacity)?;
    if let Some(leader) = active_leader(&run.attrs.leader_glyphs) {
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
    paint_resolved_text(
        context,
        &text,
        number_f32(&run.x)?,
        number_f32(&run.baseline_y)?,
        &parsed.spec,
        &parsed.chain_key,
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
    validate_glyph_effects(run)?;
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
    for glyph in &run.glyphs {
        paint_placed_glyph(context, font, glyph, run.size as f32, &paint, transform)?;
    }
    Ok(())
}

fn paint_clip_mask(
    context: &mut PaintContext<'_, '_>,
    clip: Option<&ClipRect>,
) -> Result<Option<Mask>, String> {
    let Some(clip) = clip else {
        return Ok(None);
    };
    debug_assert!(context.base_transform.is_identity() || context.base_transform.is_translate());
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
    let left = (f64::from(x) + f64::from(context.base_transform.tx)).clamp(0.0, f64::from(width));
    let right = (f64::from(x) + f64::from(slot_width) + f64::from(context.base_transform.tx))
        .clamp(0.0, f64::from(width));
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
    if let Some(family) = leader.font.as_deref() {
        spec.family = family.to_string();
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
        let chain = context
            .resources
            .font_chains
            .get(key)
            .ok_or_else(|| format!("missing font chain for `{key}`"))?;
        if chain.is_empty() {
            return Err(format!("font chain `{key}` is empty"));
        }
        for segment in font_segments(context.resources.fonts, chain, glyph)? {
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
) -> Result<(), String> {
    let chain = context
        .resources
        .font_chains
        .get(key)
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
    Ok(())
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
    if cache.store.is_some_and(|bound| bound != fonts.id()) {
        return Err("glyph cache is bound to another font store".to_string());
    }
    let key = (font.to_u32(), glyph_id);
    if !cache.entries.contains_key(&key) {
        let id = u16::try_from(glyph_id)
            .map_err(|_| format!("glyph id {glyph_id} exceeds the font outline range"))?;
        let outline = fonts
            .outline_glyph(font, id)
            .map_err(|error| error.to_string())?;
        let path = outline_path(&outline.cmds)?;
        cache.store = Some(fonts.id());
        cache.remember(
            key,
            CachedGlyph {
                path,
                upem: f32::from(outline.upem),
            },
        );
    }
    Ok(cache
        .entries
        .get(&key)
        .expect("a just-inserted or occupied key is present"))
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

    const CARLITO: &[u8] = include_bytes!("../tests/assets/Carlito-Regular.ttf");

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

    #[test]
    fn the_glyph_cache_evicts_its_oldest_entries_past_the_cap() {
        let mut cache = GlyphCache::default();
        for index in 0..MAX_CACHED_GLYPH_OUTLINES as u32 + 2 {
            cache.remember(
                (index, 0),
                CachedGlyph {
                    path: None,
                    upem: 1000.0,
                },
            );
        }
        assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
        assert_eq!(cache.order.len(), MAX_CACHED_GLYPH_OUTLINES);
        assert!(!cache.entries.contains_key(&(0, 0)));
        assert!(cache.entries.contains_key(&(2, 0)));
        assert!(
            cache
                .entries
                .contains_key(&(MAX_CACHED_GLYPH_OUTLINES as u32 + 1, 0))
        );
    }

    /// Distinct outlines past the cap evict earlier ones; a render that asks
    /// for an evicted outline gets a fresh extraction, not a stale or missing
    /// path.
    #[test]
    fn an_evicted_glyph_outline_is_extracted_again_on_demand() {
        let mut fonts = FontStore::new();
        let mut ids = Vec::new();
        for _ in 0..4 {
            ids.push(fonts.register(CARLITO.to_vec()).expect("font"));
        }
        let mut cache = GlyphCache::default();
        let mut inserted = 0usize;
        for id in &ids {
            for gid in 0..u16::MAX {
                if cached_glyph(&mut cache, &fonts, *id, u32::from(gid)).is_err() {
                    break;
                }
                inserted += 1;
            }
        }
        assert!(inserted > MAX_CACHED_GLYPH_OUTLINES);
        assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
        assert!(!cache.entries.contains_key(&(ids[0].to_u32(), 0)));
        let refilled =
            cached_glyph(&mut cache, &fonts, ids[0], 0).expect("re-extract after eviction");
        let upem = fonts.metrics(ids[0]).expect("metrics").units_per_em;
        assert_eq!(refilled.upem, f32::from(upem));
        assert!(cache.entries.contains_key(&(ids[0].to_u32(), 0)));
    }
}
