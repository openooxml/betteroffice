//! Native DOCX display-list raster backend.

#[cfg(target_arch = "wasm32")]
compile_error!("betteroffice-docx-raster is server-side only");

mod font;

pub use font::{GlyphCache, measure_text};

use std::borrow::Cow;
use std::collections::HashMap;
use std::io::Cursor;
use std::rc::Rc;

use base64::Engine as _;
use docx_layout::display_list::{
    DecorationPrimitive, DisplayBorderStyle, DisplayList, DisplayPage, DocAttrs, ImagePrimitive,
    LinePrimitive, PageBorderPrimitive, PageBorderSide, PageBorderZOrder, Primitive, RectPrimitive,
    RevisionKind, ShapePathCommand, ShapePrimitive,
};
use ooxml_text::{FontId, FontStore};
use serde_json::{Number, Value};
use tiny_skia::{
    Color, ColorU8, FillRule, FilterQuality, GradientStop, IntSize, LineCap, LineJoin,
    LinearGradient, Mask, Paint, Path, PathBuilder, PathSegment, Pixmap, PixmapPaint, Point,
    RadialGradient, Rect, SpreadMode, Stroke, StrokeDash, Transform,
};

use font::{FontSpecCache, PaintContext};

/// Layout font chains keyed as `"<family lowercase>|<bold>|<italic>"`.
pub type FontChains = HashMap<String, Vec<FontId>>;

/// Embedded image bytes keyed by display-list relationship ID.
pub type ImageMap = HashMap<String, Vec<u8>>;

/// One decoded image, at twice the surface a page may allocate.
pub const MAX_IMAGE_PIXELS: u64 = 33_554_432;
/// Every image decoded for one page, at four times that surface.
pub const MAX_PAGE_IMAGE_PIXELS: u64 = 67_108_864;
/// One decoded image's source buffer plus the pixmap it converts to, at
/// [`MAX_IMAGE_PIXELS`] of RGBA8. A deeper source costs more per pixel, so
/// pixels alone do not bound the memory a decode needs.
pub const MAX_IMAGE_BYTES: u64 = 268_435_456;
/// The same, summed across every image one page decodes.
pub const MAX_PAGE_IMAGE_BYTES: u64 = 536_870_912;
/// One `data:` payload, matching the bytes a facade registers for a
/// relationship id.
pub const MAX_DATA_URL_BYTES: u64 = 33_554_432;
/// One rendered page's longest side.
pub const MAX_PAGE_DIM: u32 = 16_384;
/// One rendered page's surface.
pub const MAX_PAGE_PIXELS: u64 = 16_777_216;
/// Scratch one page render may allocate for crop masks, clip surfaces and
/// generated paths, per pixel of the page being drawn. A mask and a clip
/// surface are both page-sized, so the page is what their cost scales with:
/// the crop-mask cache accounts for 8 of these bytes, a clip surface 5, the one
/// mask live outside the cache 1, and generated paths the rest.
pub const MAX_PAGE_SCRATCH_BYTES_PER_PIXEL: u64 = 32;
/// The floor under that, so a small page still affords its clips and paths.
pub const MIN_PAGE_SCRATCH_BYTES: u64 = 33_554_432;
/// Glyphs one page render may paint across every run on it.
pub const MAX_PAGE_GLYPHS: u64 = 1_000_000;

const MASK_BYTES_PER_PIXEL: u64 = 1;
const CLIP_BYTES_PER_PIXEL: u64 = 5;
// Keeps the affordable wave step count under f32 stagnation: raising the scratch
// budget per pixel or lowering this turns `paint_wave`'s loop into a hang.
const PATH_SEGMENT_BYTES: u64 = 40;
const MAX_CACHED_CROP_MASKS: usize = 16;
const MAX_CACHED_CROP_MASK_BYTES_PER_PIXEL: u64 = 8;
const CLIP_SURFACE_SLACK: u64 = 4;
const MAX_IMAGE_CACHE_ENTRIES: usize = 256;
const MAX_IMAGE_CACHE_KEY_BYTES: usize = 65_536;

/// The part whose relationships own a display-list relationship id.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ImageScope<'a> {
    /// `word/document.xml`.
    Body,
    /// A header or footer part, named by its `HfRegion` `r_id`.
    HeaderFooter(&'a str),
    /// `word/footnotes.xml` or `word/endnotes.xml`.
    Notes(NoteKind),
}

/// The notes part a `NoteRegion` belongs to. The display list leaves `kind`
/// out for footnotes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NoteKind {
    Footnote,
    Endnote,
}

impl NoteKind {
    fn from_region(kind: Option<&str>) -> Self {
        match kind {
            Some("endnote") => Self::Endnote,
            _ => Self::Footnote,
        }
    }
}

/// Separates a scope from the relationship id it owns.
const SCOPE_SEPARATOR: char = '\u{1f}';

/// Part-scoped [`ImageMap`] key. A header and the body can both use `rId9` for
/// different media, so bytes are keyed by owning part and resolve in that part
/// alone.
pub fn scoped_image_key(scope: ImageScope<'_>, rel_id: &str) -> String {
    match scope {
        ImageScope::Body => format!("{SCOPE_SEPARATOR}{rel_id}"),
        ImageScope::HeaderFooter(part) => format!("hf:{part}{SCOPE_SEPARATOR}{rel_id}"),
        ImageScope::Notes(NoteKind::Footnote) => format!("footnotes{SCOPE_SEPARATOR}{rel_id}"),
        ImageScope::Notes(NoteKind::Endnote) => format!("endnotes{SCOPE_SEPARATOR}{rel_id}"),
    }
}

/// Shared resources used by layout and rasterization.
pub struct RenderResources<'a> {
    pub fonts: &'a FontStore,
    pub font_chains: &'a FontChains,
    pub images: &'a ImageMap,
}

impl<'a> RenderResources<'a> {
    /// Creates a render resource view.
    pub fn new(fonts: &'a FontStore, font_chains: &'a FontChains, images: &'a ImageMap) -> Self {
        Self {
            fonts,
            font_chains,
            images,
        }
    }
}

/// A rendered page and the image references it left undrawn.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RenderedPage {
    pub bytes: Vec<u8>,
    pub skipped_images: usize,
}

/// Renders one display-list page to deterministic PNG bytes.
pub fn render_png(
    display_list: &DisplayList,
    page_ordinal: usize,
    resources: &RenderResources<'_>,
) -> Result<Vec<u8>, String> {
    render_page(display_list, page_ordinal, resources).map(|page| page.bytes)
}

/// Renders one display-list page, reporting the image references it skipped.
/// Each call starts from a fresh glyph cache; pass a shared one to
/// [`render_page_cached`] when rendering many pages of one export.
pub fn render_page(
    display_list: &DisplayList,
    page_ordinal: usize,
    resources: &RenderResources<'_>,
) -> Result<RenderedPage, String> {
    render_page_cached(
        display_list,
        page_ordinal,
        resources,
        &mut GlyphCache::default(),
    )
}

/// Renders one display-list page into a caller-owned glyph cache, so a
/// multi-page export reuses the outlines it still holds. The cache is bounded:
/// an export past its cap re-extracts what it evicted.
pub fn render_page_cached(
    display_list: &DisplayList,
    page_ordinal: usize,
    resources: &RenderResources<'_>,
    glyphs: &mut GlyphCache,
) -> Result<RenderedPage, String> {
    let page = display_list
        .pages
        .get(page_ordinal)
        .ok_or_else(|| format!("page ordinal {page_ordinal} is out of range"))?;
    let width = page_dimension(&page.width, "page width")?;
    let height = page_dimension(&page.height, "page height")?;
    validate_page_surface(width, height)?;
    let mut pixmap = Pixmap::new(width, height).ok_or_else(|| "invalid pixmap size".to_string())?;
    pixmap.fill(Color::WHITE);
    let mut renderer = Renderer::new(width, height, glyphs);
    renderer.paint_page(&mut pixmap, page, resources)?;
    Ok(RenderedPage {
        bytes: encode_png(pixmap, width, height)?,
        skipped_images: renderer.scratch.images.skipped,
    })
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct FRect {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
}

/// What one page render may spend beyond the page surface itself. Masks, clip
/// surfaces and generated paths are charged in the bytes they are about to
/// allocate, which is the one unit they share; glyphs keep their own counter,
/// since a glyph costs its rasterization rather than its footprint. Both are
/// charged before the allocation, so a page cannot overspend and then refuse.
///
/// A surface is held rather than consumed: one clip surface and one crop-mask
/// cache are alive at a time, so they are charged their high-water mark and
/// refilling one allocates nothing further. A generated path is charged every
/// time it is built, because one primitive expands into an unbounded one.
pub(crate) struct PageBudget {
    scratch: u64,
    limit: u64,
    masks: u64,
    clips: u64,
    glyphs: u64,
}

impl PageBudget {
    fn new(width: u32, height: u32) -> Self {
        let limit = u64::from(width)
            .saturating_mul(u64::from(height))
            .saturating_mul(MAX_PAGE_SCRATCH_BYTES_PER_PIXEL)
            .max(MIN_PAGE_SCRATCH_BYTES);
        Self {
            scratch: 0,
            limit,
            masks: 0,
            clips: 0,
            glyphs: 0,
        }
    }

    fn charge_scratch(&mut self, bytes: u64) -> Result<(), String> {
        let spent = self.scratch.saturating_add(bytes);
        if spent > self.limit {
            return Err(format!(
                "page exceeds its {} byte render work budget",
                self.limit
            ));
        }
        self.scratch = spent;
        Ok(())
    }

    /// Charges the growth of a held surface, leaving what it already cost
    /// charged once.
    fn charge_high_water(&mut self, held: u64, bytes: u64) -> Result<u64, String> {
        let growth = bytes.saturating_sub(held);
        if growth == 0 {
            return Ok(held);
        }
        self.charge_scratch(growth)?;
        Ok(bytes)
    }

    fn charge_masks(&mut self, bytes: u64) -> Result<(), String> {
        self.masks = self.charge_high_water(self.masks, bytes)?;
        Ok(())
    }

    fn charge_clips(&mut self, bytes: u64) -> Result<(), String> {
        self.clips = self.charge_high_water(self.clips, bytes)?;
        Ok(())
    }

    fn charge_path(&mut self, segments: u64) -> Result<(), String> {
        self.charge_scratch(segments.saturating_mul(PATH_SEGMENT_BYTES))
    }

    pub(crate) fn charge_glyphs(&mut self, glyphs: u64) -> Result<(), String> {
        self.glyphs = self.glyphs.saturating_add(glyphs);
        if self.glyphs > MAX_PAGE_GLYPHS {
            return Err(format!(
                "page exceeds the {MAX_PAGE_GLYPHS} painted glyph budget"
            ));
        }
        Ok(())
    }
}

/// The geometry a crop mask was filled for. A mask covers the whole page
/// whatever it crops, so identical geometry has to reuse one.
#[derive(Clone, Copy, PartialEq, Eq)]
struct CropMaskKey {
    size: (u32, u32),
    frame: [u32; 4],
    transform: [u32; 6],
}

impl CropMaskKey {
    fn new(pixmap: &Pixmap, frame: FRect, transform: Transform) -> Self {
        Self {
            size: (pixmap.width(), pixmap.height()),
            frame: [
                frame.x.to_bits(),
                frame.y.to_bits(),
                frame.w.to_bits(),
                frame.h.to_bits(),
            ],
            transform: [
                transform.sx.to_bits(),
                transform.kx.to_bits(),
                transform.ky.to_bits(),
                transform.sy.to_bits(),
                transform.tx.to_bits(),
                transform.ty.to_bits(),
            ],
        }
    }
}

/// Filled crop masks, most recently used first. A page walks its geometries in
/// whatever order the display list lists them, so a cache that evicts by age of
/// insertion drops the entry a round-robin is about to ask for again.
struct MaskCache {
    entries: Vec<(CropMaskKey, Rc<Mask>)>,
    bytes: u64,
    limit: u64,
}

impl MaskCache {
    fn new(width: u32, height: u32) -> Self {
        Self {
            entries: Vec::new(),
            bytes: 0,
            limit: u64::from(width)
                .saturating_mul(u64::from(height))
                .saturating_mul(MAX_CACHED_CROP_MASK_BYTES_PER_PIXEL),
        }
    }

    fn take(&mut self, key: CropMaskKey) -> Option<Rc<Mask>> {
        let index = self.entries.iter().position(|(cached, _)| *cached == key)?;
        let entry = self.entries.remove(index);
        let mask = entry.1.clone();
        self.entries.insert(0, entry);
        Some(mask)
    }

    fn remember(&mut self, key: CropMaskKey, mask: Rc<Mask>, bytes: u64) {
        if bytes > self.limit {
            return;
        }
        self.entries.insert(0, (key, mask));
        self.bytes += bytes;
        while self.entries.len() > MAX_CACHED_CROP_MASKS || self.bytes > self.limit {
            let Some((key, _)) = self.entries.pop() else {
                break;
            };
            self.bytes = self
                .bytes
                .saturating_sub(mask_bytes(key.size.0, key.size.1));
        }
    }
}

fn mask_bytes(width: u32, height: u32) -> u64 {
    u64::from(width)
        .saturating_mul(u64::from(height))
        .saturating_mul(MASK_BYTES_PER_PIXEL)
}

/// Everything one page render accumulates: the caches that keep repeated work
/// from repeating, and the budget that bounds the work left over.
struct Scratch<'k> {
    glyphs: &'k mut GlyphCache,
    specs: FontSpecCache,
    images: ImageCache<'k>,
    masks: MaskCache,
    budget: PageBudget,
}

struct Renderer<'k> {
    clips: ClipSurface,
    scratch: Scratch<'k>,
}

impl<'k> Renderer<'k> {
    fn new(width: u32, height: u32, glyphs: &'k mut GlyphCache) -> Self {
        Self {
            clips: ClipSurface::default(),
            scratch: Scratch {
                glyphs,
                specs: FontSpecCache::default(),
                images: ImageCache::default(),
                masks: MaskCache::new(width, height),
                budget: PageBudget::new(width, height),
            },
        }
    }

    fn paint_page(
        &mut self,
        pixmap: &mut Pixmap,
        page: &'k DisplayPage,
        resources: &RenderResources<'_>,
    ) -> Result<(), String> {
        if let Some(background) = &page.background {
            let rect = Rect::from_xywh(0.0, 0.0, pixmap.width() as f32, pixmap.height() as f32)
                .ok_or_else(|| "invalid page background rectangle".to_string())?;
            let mut paint = Paint::default();
            paint.set_color(parse_color(background)?);
            pixmap.fill_rect(rect, &paint, Transform::identity(), None);
        }
        for border in page
            .page_borders
            .iter()
            .filter(|border| border.z_order == Some(PageBorderZOrder::Back))
        {
            paint_page_border(pixmap, border, &mut self.scratch.budget)?;
        }
        for primitive in &page.primitives {
            self.paint_primitive(pixmap, primitive, resources, ImageScope::Body)?;
        }
        for area in &page.note_areas {
            let scope = ImageScope::Notes(NoteKind::from_region(area.kind.as_deref()));
            for primitive in &area.separator_primitives {
                self.paint_primitive(pixmap, primitive, resources, scope)?;
            }
            for primitive in &area.primitives {
                self.paint_primitive(pixmap, primitive, resources, scope)?;
            }
        }
        for region in [&page.header, &page.footer].into_iter().flatten() {
            let scope = ImageScope::HeaderFooter(&region.r_id);
            for primitive in &region.primitives {
                self.paint_primitive(pixmap, primitive, resources, scope)?;
            }
        }
        for border in page
            .page_borders
            .iter()
            .filter(|border| border.z_order != Some(PageBorderZOrder::Back))
        {
            paint_page_border(pixmap, border, &mut self.scratch.budget)?;
        }
        Ok(())
    }

    fn paint_primitive(
        &mut self,
        pixmap: &mut Pixmap,
        primitive: &'k Primitive,
        resources: &RenderResources<'_>,
        scope: ImageScope<'k>,
    ) -> Result<(), String> {
        let attrs = primitive_attrs(primitive);
        validate_visual_attrs(primitive, attrs)?;
        let opacity = primitive_opacity(primitive)?;
        let clip = clip_rect(attrs)?;
        let clips = &mut self.clips;
        let scratch = &mut self.scratch;
        if let Some(clip) = clip {
            clips.paint(pixmap, clip, scratch, |target, scratch, transform, mask| {
                paint_primitive_core(
                    target, primitive, resources, scratch, transform, mask, opacity, scope,
                )
            })
        } else {
            paint_primitive_core(
                pixmap,
                primitive,
                resources,
                scratch,
                Transform::identity(),
                None,
                opacity,
                scope,
            )
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn paint_primitive_core<'k>(
    pixmap: &mut Pixmap,
    primitive: &'k Primitive,
    resources: &RenderResources<'_>,
    scratch: &mut Scratch<'k>,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    scope: ImageScope<'k>,
) -> Result<(), String> {
    match primitive {
        Primitive::Text(run) => {
            let retained_mask_bytes = scratch.masks.bytes;
            let mut context = PaintContext {
                pixmap,
                resources,
                cache: &mut *scratch.glyphs,
                specs: &mut scratch.specs,
                budget: &mut scratch.budget,
                base_transform: transform,
                mask,
                retained_mask_bytes,
                opacity,
            };
            font::paint_text(&mut context, run)
        }
        Primitive::GlyphRun(run) => {
            let retained_mask_bytes = scratch.masks.bytes;
            let mut context = PaintContext {
                pixmap,
                resources,
                cache: &mut *scratch.glyphs,
                specs: &mut scratch.specs,
                budget: &mut scratch.budget,
                base_transform: transform,
                mask,
                retained_mask_bytes,
                opacity,
            };
            font::paint_glyph_run(&mut context, run)
        }
        Primitive::Rect(rect) => paint_rect(pixmap, rect, transform, mask, opacity),
        Primitive::Line(line) => {
            paint_line_primitive(pixmap, line, transform, mask, opacity, &mut scratch.budget)
        }
        Primitive::Image(image) => paint_image(
            pixmap, image, resources, scratch, transform, mask, opacity, scope,
        ),
        Primitive::Shape(shape) => {
            paint_shape(pixmap, shape, transform, mask, opacity, &mut scratch.budget)
        }
        Primitive::Decoration(decoration) => paint_decoration(
            pixmap,
            decoration,
            transform,
            mask,
            opacity,
            &mut scratch.budget,
        ),
    }
}

fn paint_rect(
    pixmap: &mut Pixmap,
    primitive: &RectPrimitive,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
) -> Result<(), String> {
    let rect = primitive_rect(
        &primitive.x,
        &primitive.y,
        &primitive.w,
        &primitive.h,
        "rectangle",
    )?;
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(&primitive.fill, opacity)?);
    paint.anti_alias = true;
    pixmap.fill_rect(rect, &paint, transform, mask);
    Ok(())
}

fn paint_line_primitive(
    pixmap: &mut Pixmap,
    line: &LinePrimitive,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let x1 = number_f32(&line.x1)?;
    let y1 = number_f32(&line.y1)?;
    let x2 = number_f32(&line.x2)?;
    let y2 = number_f32(&line.y2)?;
    let width = number_f32(&line.stroke_width)?.max(0.5);
    let dash = line.dash.as_deref().map(number_slice).transpose()?;
    paint_border_recipe(
        pixmap,
        Segment { x1, y1, x2, y2 },
        width,
        &line.color,
        line.border_style.unwrap_or(DisplayBorderStyle::Solid),
        line.secondary_color.as_deref(),
        dash.as_deref(),
        transform,
        mask,
        opacity,
        budget,
    )
}

fn paint_decoration(
    pixmap: &mut Pixmap,
    decoration: &DecorationPrimitive,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let x = number_f32(&decoration.x)?;
    let y = number_f32(&decoration.y)?;
    let w = number_f32(&decoration.w)?;
    let h = number_f32(&decoration.h)?;
    if w < 0.0 || h < 0.0 {
        return Err("decoration dimensions must be non-negative".to_string());
    }
    if let Some(style) = decoration.attrs.style
        && style != DisplayBorderStyle::Solid
    {
        return paint_border_recipe(
            pixmap,
            Segment {
                x1: x,
                y1: y + h / 2.0,
                x2: x + w,
                y2: y + h / 2.0,
            },
            h.max(1.0),
            &decoration.color,
            style,
            None,
            None,
            transform,
            mask,
            opacity,
            budget,
        );
    }
    if decoration.dashed || decoration.dotted {
        let thickness = h.max(1.0);
        let unit = if decoration.dotted {
            thickness
        } else {
            (thickness * 2.0).round().max(2.0)
        };
        let gap = if decoration.dotted {
            (thickness * 2.0).round().max(2.0)
        } else {
            unit
        };
        return stroke_segment(
            pixmap,
            Segment {
                x1: x,
                y1: y + h / 2.0,
                x2: x + w,
                y2: y + h / 2.0,
            },
            thickness,
            &decoration.color,
            Some(&[unit, gap]),
            transform,
            mask,
            opacity,
            StrokeStyle::default(),
            budget,
        );
    }
    let rect =
        Rect::from_xywh(x, y, w, h).ok_or_else(|| "invalid decoration rectangle".to_string())?;
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(&decoration.color, opacity)?);
    paint.anti_alias = true;
    pixmap.fill_rect(rect, &paint, transform, mask);
    Ok(())
}

fn paint_shape(
    pixmap: &mut Pixmap,
    shape: &ShapePrimitive,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    if !shape.attrs.effects.is_empty() {
        return Err("unsupported shape field: effects".to_string());
    }
    budget.charge_path(shape.geometry_path.len() as u64)?;
    let path = build_shape_path(&shape.geometry_path)?;
    let visual = shape_transform(shape)?;
    let transform = transform.pre_concat(visual);
    if let Some(paint) = shape_fill(shape, opacity)? {
        pixmap.fill_path(&path, &paint, FillRule::Winding, transform, mask);
    }
    if let Some(stroked) = shape_stroke(shape, opacity)? {
        if let Some(dash) = &stroked.dash {
            budget.charge_path(dash_segments(path_length(&path), dash))?;
        }
        pixmap.stroke_path(&path, &stroked.paint, &stroked.stroke, transform, mask);
    }
    Ok(())
}

fn shape_fill(shape: &ShapePrimitive, opacity: f32) -> Result<Option<Paint<'static>>, String> {
    let Some(value) = &shape.attrs.fill_paint else {
        return shape
            .fill
            .as_deref()
            .map(|color| solid_paint(color, opacity))
            .transpose();
    };
    let object = value
        .as_object()
        .ok_or_else(|| "shape fillPaint must be an object".to_string())?;
    ensure_keys(
        object,
        &[
            "kind",
            "color",
            "angle",
            "stops",
            "gradientType",
            "patternPreset",
            "foregroundColor",
            "backgroundColor",
            "pictureRelId",
            "pictureFillMode",
            "pictureSrc",
            "pictureSrcRect",
            "pictureTile",
            "pictureStretchRect",
            "pictureOpacity",
            "themeRefIndex",
        ],
        "shape fillPaint",
    )?;
    let kind = object
        .get("kind")
        .and_then(Value::as_str)
        .unwrap_or("solid");
    match kind {
        "none" => {
            reject_shape_fill_fields(
                object,
                &[
                    "color",
                    "angle",
                    "stops",
                    "gradientType",
                    "patternPreset",
                    "foregroundColor",
                    "backgroundColor",
                    "pictureRelId",
                    "pictureFillMode",
                    "pictureSrc",
                    "pictureSrcRect",
                    "pictureTile",
                    "pictureStretchRect",
                    "pictureOpacity",
                ],
            )?;
            Ok(None)
        }
        "solid" | "theme" => {
            reject_shape_fill_fields(
                object,
                &[
                    "angle",
                    "stops",
                    "gradientType",
                    "patternPreset",
                    "foregroundColor",
                    "backgroundColor",
                    "pictureRelId",
                    "pictureFillMode",
                    "pictureSrc",
                    "pictureSrcRect",
                    "pictureTile",
                    "pictureStretchRect",
                    "pictureOpacity",
                ],
            )?;
            let Some(color) = object
                .get("color")
                .and_then(Value::as_str)
                .or(shape.fill.as_deref())
            else {
                return Ok(None);
            };
            Ok(Some(solid_paint(color, opacity)?))
        }
        "gradient" => gradient_paint(shape, object, opacity).map(Some),
        "pattern" => Err("unsupported shape fillPaint kind: pattern".to_string()),
        "picture" => Err("unsupported shape fillPaint kind: picture".to_string()),
        other => Err(format!("unsupported shape fillPaint kind: {other}")),
    }
}

fn reject_shape_fill_fields(
    object: &serde_json::Map<String, Value>,
    fields: &[&str],
) -> Result<(), String> {
    if let Some(field) = fields.iter().find(|field| object.contains_key(**field)) {
        return Err(format!(
            "unsupported shape {field} field for this fill kind"
        ));
    }
    Ok(())
}

fn gradient_paint(
    shape: &ShapePrimitive,
    object: &serde_json::Map<String, Value>,
    opacity: f32,
) -> Result<Paint<'static>, String> {
    for field in [
        "patternPreset",
        "foregroundColor",
        "backgroundColor",
        "pictureRelId",
        "pictureFillMode",
        "pictureSrc",
        "pictureSrcRect",
        "pictureTile",
        "pictureStretchRect",
        "pictureOpacity",
    ] {
        if object.contains_key(field) {
            return Err(format!("unsupported gradient fill field: {field}"));
        }
    }
    let fallback = object
        .get("color")
        .and_then(Value::as_str)
        .or(shape.fill.as_deref())
        .unwrap_or("#000000");
    let stops = gradient_stops(object.get("stops"), fallback, opacity)?;
    let x = number_f32(&shape.x)?;
    let y = number_f32(&shape.y)?;
    let w = number_f32(&shape.w)?;
    let h = number_f32(&shape.h)?;
    let gradient_type = object
        .get("gradientType")
        .and_then(Value::as_str)
        .unwrap_or("linear");
    let shader = match gradient_type {
        "linear" => {
            let angle = object
                .get("angle")
                .map(value_f32)
                .transpose()?
                .unwrap_or(0.0);
            let radians = angle.to_radians();
            let dx = radians.cos();
            let dy = radians.sin();
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let reach = dx.abs() * w / 2.0 + dy.abs() * h / 2.0;
            LinearGradient::new(
                Point::from_xy(cx - dx * reach, cy - dy * reach),
                Point::from_xy(cx + dx * reach, cy + dy * reach),
                stops,
                SpreadMode::Pad,
                Transform::identity(),
            )
        }
        "radial" => {
            let center = Point::from_xy(x + w / 2.0, y + h / 2.0);
            RadialGradient::new(
                center,
                0.0,
                center,
                w.max(h) / 2.0,
                stops,
                SpreadMode::Pad,
                Transform::identity(),
            )
        }
        other => return Err(format!("unsupported gradient type: {other}")),
    }
    .ok_or_else(|| "invalid shape gradient".to_string())?;
    Ok(Paint {
        shader,
        anti_alias: true,
        ..Paint::default()
    })
}

fn gradient_stops(
    value: Option<&Value>,
    fallback: &str,
    opacity: f32,
) -> Result<Vec<GradientStop>, String> {
    let mut stops = Vec::new();
    if let Some(value) = value {
        let values = value
            .as_array()
            .ok_or_else(|| "gradient stops must be an array".to_string())?;
        for value in values {
            let object = value
                .as_object()
                .ok_or_else(|| "gradient stop must be an object".to_string())?;
            ensure_keys(object, &["position", "color"], "gradient stop")?;
            let mut position = object
                .get("position")
                .map(value_f32)
                .transpose()?
                .unwrap_or(0.0);
            if position > 1.0 {
                position /= 100_000.0;
            }
            let color = object
                .get("color")
                .and_then(Value::as_str)
                .unwrap_or(fallback);
            stops.push((
                position.clamp(0.0, 1.0),
                color_with_opacity(color, opacity)?,
            ));
        }
    }
    if stops.is_empty() {
        let color = color_with_opacity(fallback, opacity)?;
        stops.push((0.0, color));
        stops.push((1.0, color));
    } else if stops.len() == 1 {
        let color = stops[0].1;
        stops.clear();
        stops.push((0.0, color));
        stops.push((1.0, color));
    }
    stops.sort_by(|left, right| left.0.total_cmp(&right.0));
    Ok(stops
        .into_iter()
        .map(|(position, color)| GradientStop::new(position, color))
        .collect())
}

/// A resolved shape stroke, keeping the dash array the [`Stroke`] hides so the
/// path it expands into can be charged.
struct ShapeStroke {
    paint: Paint<'static>,
    stroke: Stroke,
    dash: Option<Vec<f32>>,
}

fn shape_stroke(shape: &ShapePrimitive, opacity: f32) -> Result<Option<ShapeStroke>, String> {
    let mut color = shape.stroke.as_ref().map(|stroke| stroke.color.as_str());
    let mut width = shape
        .stroke
        .as_ref()
        .map(|stroke| number_f32(&stroke.width))
        .transpose()?
        .unwrap_or(1.0);
    let mut dash_name = shape
        .stroke
        .as_ref()
        .and_then(|stroke| stroke.dash.as_deref());
    let mut custom_dash = None;
    let mut style = StrokeStyle::default();
    if let Some(value) = &shape.attrs.stroke_paint {
        let object = value
            .as_object()
            .ok_or_else(|| "shape strokePaint must be an object".to_string())?;
        ensure_keys(
            object,
            &[
                "color",
                "width",
                "dash",
                "customDash",
                "compound",
                "alignment",
                "cap",
                "join",
                "miterLimit",
                "headEnd",
                "tailEnd",
            ],
            "shape strokePaint",
        )?;
        color = object.get("color").and_then(Value::as_str).or(color);
        if let Some(value) = object.get("width") {
            width = value_f32(value)?;
        }
        dash_name = object.get("dash").and_then(Value::as_str).or(dash_name);
        custom_dash = object
            .get("customDash")
            .map(value_number_array)
            .transpose()?;
        if let Some(compound) = object.get("compound").and_then(Value::as_str)
            && compound != "single"
        {
            return Err(format!("unsupported shape stroke compound: {compound}"));
        }
        if let Some(alignment) = object.get("alignment").and_then(Value::as_str)
            && alignment != "center"
        {
            return Err(format!("unsupported shape stroke alignment: {alignment}"));
        }
        if object.get("headEnd").is_some_and(|value| !value.is_null()) {
            return Err("unsupported shape stroke field: headEnd".to_string());
        }
        if object.get("tailEnd").is_some_and(|value| !value.is_null()) {
            return Err("unsupported shape stroke field: tailEnd".to_string());
        }
        if let Some(cap) = object.get("cap").and_then(Value::as_str) {
            style.cap = match cap {
                "flat" => LineCap::Butt,
                "round" => LineCap::Round,
                "square" => LineCap::Square,
                other => return Err(format!("unsupported shape stroke cap: {other}")),
            };
        }
        if let Some(join) = object.get("join").and_then(Value::as_str) {
            style.join = match join {
                "miter" => LineJoin::Miter,
                "round" => LineJoin::Round,
                "bevel" => LineJoin::Bevel,
                other => return Err(format!("unsupported shape stroke join: {other}")),
            };
        }
        if let Some(limit) = object.get("miterLimit") {
            style.miter_limit = value_f32(limit)?.max(1.0);
        }
    }
    if shape.stroke.is_none() && shape.attrs.stroke_paint.is_none() {
        return Ok(None);
    }
    if width < 0.0 {
        return Err("shape stroke width must be non-negative".to_string());
    }
    if width == 0.0 {
        return Ok(None);
    }
    let color = color.unwrap_or("#000000");
    let dash = if let Some(values) = custom_dash {
        Some(
            values
                .into_iter()
                .map(|value| value.max(0.0) * width)
                .collect(),
        )
    } else {
        shape_dash_pattern(dash_name, width)?
    };
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(color, opacity)?);
    paint.anti_alias = true;
    let stroke = stroke_style(width, dash.as_deref(), style)?;
    Ok(Some(ShapeStroke {
        paint,
        stroke,
        dash,
    }))
}

fn shape_transform(shape: &ShapePrimitive) -> Result<Transform, String> {
    let Some(value) = &shape.transform else {
        return Ok(Transform::identity());
    };
    let x = number_f32(&shape.x)?;
    let y = number_f32(&shape.y)?;
    let w = number_f32(&shape.w)?;
    let h = number_f32(&shape.h)?;
    let center_x = x + w / 2.0;
    let center_y = y + h / 2.0;
    let mut transform = Transform::identity();
    if let Some(rotation) = &value.rotation {
        transform = transform.pre_concat(Transform::from_rotate_at(
            number_f32(rotation)?,
            center_x,
            center_y,
        ));
    }
    if value.flip_h || value.flip_v {
        transform = transform.pre_concat(scale_at(
            if value.flip_h { -1.0 } else { 1.0 },
            if value.flip_v { -1.0 } else { 1.0 },
            center_x,
            center_y,
        ));
    }
    Ok(transform)
}

fn build_shape_path(commands: &[ShapePathCommand]) -> Result<Path, String> {
    let mut builder = PathBuilder::new();
    for command in commands {
        match command {
            ShapePathCommand::Move { x, y } => {
                builder.move_to(number_f32(x)?, number_f32(y)?);
            }
            ShapePathCommand::Line { x, y } => {
                builder.line_to(number_f32(x)?, number_f32(y)?);
            }
            ShapePathCommand::Quad { cpx, cpy, x, y } => {
                builder.quad_to(
                    number_f32(cpx)?,
                    number_f32(cpy)?,
                    number_f32(x)?,
                    number_f32(y)?,
                );
            }
            ShapePathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                builder.cubic_to(
                    number_f32(cp1x)?,
                    number_f32(cp1y)?,
                    number_f32(cp2x)?,
                    number_f32(cp2y)?,
                    number_f32(x)?,
                    number_f32(y)?,
                );
            }
            ShapePathCommand::Close => builder.close(),
        }
    }
    builder
        .finish()
        .ok_or_else(|| "shape path has no drawable commands".to_string())
}

#[allow(clippy::too_many_arguments)]
fn paint_image<'k>(
    pixmap: &mut Pixmap,
    image: &'k ImagePrimitive,
    resources: &RenderResources<'_>,
    scratch: &mut Scratch<'k>,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    scope: ImageScope<'k>,
) -> Result<(), String> {
    if image.filter.is_some() {
        return Err("unsupported image field: filter".to_string());
    }
    if !image.attrs.effects.is_empty() {
        return Err("unsupported image field: effects".to_string());
    }
    let Some((key, source_bytes)) = image_source(&image.rel_id, scope, resources) else {
        scratch.images.skipped += 1;
        return Ok(());
    };
    let Some(decoded) = scratch.images.resolve(key, source_bytes) else {
        return Ok(());
    };
    let source: &Pixmap = &decoded;
    let frame = image_frame(image)?;
    if frame.w == 0.0 || frame.h == 0.0 {
        return Ok(());
    }
    let visual = image_transform(image, frame)?;
    let image_transform = transform.pre_concat(visual);
    let crop = image.crop.as_ref();
    let source_width = source.width() as f32;
    let source_height = source.height() as f32;
    let left = crop
        .map(|crop| number_f32(&crop.left))
        .transpose()?
        .unwrap_or(0.0);
    let top = crop
        .map(|crop| number_f32(&crop.top))
        .transpose()?
        .unwrap_or(0.0);
    let right = crop
        .map(|crop| number_f32(&crop.right))
        .transpose()?
        .unwrap_or(0.0);
    let bottom = crop
        .map(|crop| number_f32(&crop.bottom))
        .transpose()?
        .unwrap_or(0.0);
    let requested_x = left * source_width;
    let requested_y = top * source_height;
    let requested_w = source_width * (1.0 - left - right);
    let requested_h = source_height * (1.0 - top - bottom);
    if requested_w <= 0.0 || requested_h <= 0.0 {
        return Ok(());
    }
    let source_to_frame = Transform::from_row(
        frame.w / requested_w,
        0.0,
        0.0,
        frame.h / requested_h,
        frame.x - requested_x * frame.w / requested_w,
        frame.y - requested_y * frame.h / requested_h,
    );
    let mut frame_mask = crop
        .is_some()
        .then(|| {
            crop_mask(
                pixmap,
                frame,
                image_transform,
                mask,
                &mut scratch.masks,
                &mut scratch.budget,
            )
        })
        .transpose()?;
    let paint = PixmapPaint {
        opacity,
        quality: FilterQuality::Bicubic,
        ..PixmapPaint::default()
    };
    pixmap.draw_pixmap(
        0,
        0,
        source.as_ref(),
        &paint,
        image_transform.pre_concat(source_to_frame),
        frame_mask.as_deref().or(mask),
    );
    frame_mask.take();
    paint_image_border(
        pixmap,
        image,
        frame,
        image_transform,
        mask,
        opacity,
        &mut scratch.budget,
    )?;
    paint_image_revision(
        pixmap,
        image,
        frame,
        image_transform,
        mask,
        opacity,
        &mut scratch.budget,
    )
}

/// Where a reference's bytes come from. A `data:` payload stays encoded so the
/// cache is consulted before base64 expands it.
enum ImageSource<'k, 'b> {
    Data(&'k str),
    Registered(&'b [u8]),
}

/// The identity a reference resolved to and its bytes, or `None` when the
/// reference resolves to nothing.
fn image_source<'k, 'b>(
    rel_id: &'k str,
    scope: ImageScope<'_>,
    resources: &RenderResources<'b>,
) -> Option<(Cow<'k, str>, ImageSource<'k, 'b>)> {
    if rel_id.starts_with("data:") {
        let (metadata, payload) = rel_id.split_once(',')?;
        if !metadata.ends_with(";base64") {
            return None;
        }
        return Some((Cow::Borrowed(rel_id), ImageSource::Data(payload)));
    }
    let key = scoped_image_key(scope, rel_id);
    if let Some(bytes) = resources.images.get(&key) {
        return Some((Cow::Owned(key), ImageSource::Registered(bytes.as_slice())));
    }
    let bytes = legacy_body_image(rel_id, scope, resources)?;
    Some((Cow::Borrowed(rel_id), ImageSource::Registered(bytes)))
}

/// [`ImageMap`] shipped keyed by bare relationship id, so a body image still
/// resolves that way. The body alone: a header reaching a bare key is the
/// cross-part lookup scoping exists to prevent. A relationship id carrying the
/// separator would spell another part's scoped key, so it resolves to nothing.
fn legacy_body_image<'b>(
    rel_id: &str,
    scope: ImageScope<'_>,
    resources: &RenderResources<'b>,
) -> Option<&'b [u8]> {
    if !matches!(scope, ImageScope::Body) || rel_id.contains(SCOPE_SEPARATOR) {
        return None;
    }
    resources.images.get(rel_id).map(Vec::as_slice)
}

/// One decode per resolved image, and at most [`MAX_PAGE_IMAGE_PIXELS`]
/// decoded for the whole page.
#[derive(Default)]
struct ImageCache<'k> {
    decoded: HashMap<Cow<'k, str>, Option<Rc<Pixmap>>>,
    key_bytes: usize,
    pixels: u64,
    bytes: u64,
    skipped: usize,
}

impl<'k> ImageCache<'k> {
    fn resolve(&mut self, key: Cow<'k, str>, source: ImageSource<'_, '_>) -> Option<Rc<Pixmap>> {
        let decoded = match self.decoded.get(key.as_ref()) {
            Some(entry) => entry.clone(),
            None => {
                let decoded = self.materialize(source).map(Rc::new);
                self.remember(key, decoded.clone());
                decoded
            }
        };
        if decoded.is_none() {
            self.skipped += 1;
        }
        decoded
    }

    /// Keeps a decode against its reference, while the map has room. Only an
    /// owned key costs bytes: a `data:` URL is borrowed from the display list,
    /// so the cache holds a pointer into it rather than a copy of it.
    fn remember(&mut self, key: Cow<'k, str>, decoded: Option<Rc<Pixmap>>) {
        let owned = match &key {
            Cow::Owned(value) => value.len(),
            Cow::Borrowed(_) => 0,
        };
        if self.decoded.len() >= MAX_IMAGE_CACHE_ENTRIES
            || self.key_bytes + owned > MAX_IMAGE_CACHE_KEY_BYTES
        {
            return;
        }
        self.key_bytes += owned;
        self.decoded.insert(key, decoded);
    }

    /// The bytes behind a reference the cache has not seen. A `data:` payload
    /// is bounded and charged from its encoded length, before base64 expands
    /// it into a buffer.
    fn materialize(&mut self, source: ImageSource<'_, '_>) -> Option<Pixmap> {
        match source {
            ImageSource::Registered(bytes) => self.decode(bytes),
            ImageSource::Data(payload) => {
                let declared = payload.len() as u64 / 4 * 3;
                if declared > MAX_DATA_URL_BYTES || self.bytes + declared > MAX_PAGE_IMAGE_BYTES {
                    return None;
                }
                self.bytes += declared;
                let bytes = base64::engine::general_purpose::STANDARD
                    .decode(payload)
                    .ok()?;
                self.decode(&bytes)
            }
        }
    }

    /// Decoded pixels, or `None` for content this backend will not draw: bytes
    /// it cannot decode, an image past [`MAX_IMAGE_PIXELS`], or one the page
    /// has no budget left for. Declared pixels are charged before the decoder
    /// allocates, so a stream that fails late still costs what it claimed.
    fn decode(&mut self, bytes: &[u8]) -> Option<Pixmap> {
        use image::ImageDecoder as _;

        let mut decoder = image::ImageReader::new(Cursor::new(bytes))
            .with_guessed_format()
            .ok()?
            .into_decoder()
            .ok()?;
        let (declared_width, declared_height) = decoder.dimensions();
        let declared = u64::from(declared_width) * u64::from(declared_height);
        if declared > MAX_IMAGE_PIXELS || self.pixels + declared > MAX_PAGE_IMAGE_PIXELS {
            return None;
        }
        let cost = decoder
            .total_bytes()
            .saturating_add(declared.saturating_mul(4));
        if cost > MAX_IMAGE_BYTES || self.bytes + cost > MAX_PAGE_IMAGE_BYTES {
            return None;
        }
        self.pixels += declared;
        self.bytes += cost;
        let orientation = decoder.orientation().ok()?;
        let mut decoded = image::DynamicImage::from_decoder(decoder).ok()?;
        decoded.apply_orientation(orientation);
        let size = IntSize::from_wh(decoded.width(), decoded.height())?;
        let mut data = decoded.into_rgba8().into_raw();
        let (pixels, _) = data.as_chunks_mut::<4>();
        for pixel in pixels {
            let color = ColorU8::from_rgba(pixel[0], pixel[1], pixel[2], pixel[3]).premultiply();
            *pixel = [color.red(), color.green(), color.blue(), color.alpha()];
        }
        Pixmap::from_vec(data, size)
    }
}

fn image_frame(image: &ImagePrimitive) -> Result<FRect, String> {
    let x = image
        .attrs
        .content_frame
        .as_ref()
        .and_then(|frame| frame.x.as_ref())
        .unwrap_or(&image.x);
    let y = image
        .attrs
        .content_frame
        .as_ref()
        .and_then(|frame| frame.y.as_ref())
        .unwrap_or(&image.y);
    let w = image
        .attrs
        .content_frame
        .as_ref()
        .and_then(|frame| frame.w.as_ref())
        .unwrap_or(&image.w);
    let h = image
        .attrs
        .content_frame
        .as_ref()
        .and_then(|frame| frame.h.as_ref())
        .unwrap_or(&image.h);
    Ok(FRect {
        x: number_f32(x)?,
        y: number_f32(y)?,
        w: number_f32(w)?.max(0.0),
        h: number_f32(h)?.max(0.0),
    })
}

fn image_transform(image: &ImagePrimitive, frame: FRect) -> Result<Transform, String> {
    let center_x = frame.x + frame.w / 2.0;
    let center_y = frame.y + frame.h / 2.0;
    let mut transform = Transform::identity();
    if let Some(rotation) = &image.rotation_deg {
        transform = transform.pre_concat(Transform::from_rotate_at(
            number_f32(rotation)?,
            center_x,
            center_y,
        ));
    }
    let flip_h = image.attrs.image_flip_h == Some(true);
    let flip_v = image.attrs.image_flip_v == Some(true);
    if flip_h || flip_v {
        transform = transform.pre_concat(scale_at(
            if flip_h { -1.0 } else { 1.0 },
            if flip_v { -1.0 } else { 1.0 },
            center_x,
            center_y,
        ));
    }
    Ok(transform)
}

/// The crop mask for one image reference, reused where the geometry repeats.
/// An outer clip mixes into the fill, so a clipped crop is built fresh rather
/// than shared. Only one mask is ever live outside the cache, so the page is
/// charged what the cache holds plus that one, not one charge per reference.
fn crop_mask(
    pixmap: &Pixmap,
    frame: FRect,
    transform: Transform,
    outer: Option<&Mask>,
    masks: &mut MaskCache,
    budget: &mut PageBudget,
) -> Result<Rc<Mask>, String> {
    let bytes = mask_bytes(pixmap.width(), pixmap.height());
    if outer.is_some() {
        budget.charge_masks(masks.bytes.saturating_add(bytes))?;
        return Ok(Rc::new(transformed_rect_mask(
            pixmap, frame, transform, outer,
        )?));
    }
    let key = CropMaskKey::new(pixmap, frame, transform);
    if let Some(mask) = masks.take(key) {
        return Ok(mask);
    }
    budget.charge_masks(masks.bytes.saturating_add(bytes))?;
    let mask = Rc::new(transformed_rect_mask(pixmap, frame, transform, None)?);
    masks.remember(key, mask.clone(), bytes);
    Ok(mask)
}

fn transformed_rect_mask(
    pixmap: &Pixmap,
    frame: FRect,
    transform: Transform,
    outer: Option<&Mask>,
) -> Result<Mask, String> {
    let mut mask = Mask::new(pixmap.width(), pixmap.height())
        .ok_or_else(|| "invalid image crop mask size".to_string())?;
    let rect = Rect::from_xywh(frame.x, frame.y, frame.w, frame.h)
        .ok_or_else(|| "invalid image crop rectangle".to_string())?;
    let path = PathBuilder::from_rect(rect);
    mask.fill_path(&path, FillRule::Winding, true, transform);
    if let Some(outer) = outer {
        for (value, clip) in mask.data_mut().iter_mut().zip(outer.data()) {
            *value = ((u16::from(*value) * u16::from(*clip) + 127) / 255) as u8;
        }
    }
    Ok(mask)
}

#[allow(clippy::too_many_arguments)]
fn paint_image_border(
    pixmap: &mut Pixmap,
    image: &ImagePrimitive,
    frame: FRect,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let Some(value) = &image.attrs.border else {
        return Ok(());
    };
    let object = value
        .as_object()
        .ok_or_else(|| "image border must be an object".to_string())?;
    ensure_keys(object, &["width", "color", "style", "dash"], "image border")?;
    let width = object
        .get("width")
        .map(value_f32)
        .transpose()?
        .unwrap_or(0.0);
    if width <= 0.0 {
        return Ok(());
    }
    let color = object
        .get("color")
        .and_then(Value::as_str)
        .unwrap_or("#000000");
    let dash = if let Some(value) = object.get("dash") {
        Some(value_number_array(value)?)
    } else {
        shape_dash_pattern(object.get("style").and_then(Value::as_str), width)?
    };
    let rect = Rect::from_xywh(frame.x, frame.y, frame.w, frame.h)
        .ok_or_else(|| "invalid image border rectangle".to_string())?;
    let path = PathBuilder::from_rect(rect);
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(color, opacity)?);
    paint.anti_alias = true;
    let stroke = stroke_style(width, dash.as_deref(), StrokeStyle::default())?;
    if let Some(dash) = &dash {
        budget.charge_path(dash_segments((frame.w + frame.h) * 2.0, dash))?;
    }
    pixmap.stroke_path(&path, &paint, &stroke, transform, mask);
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn paint_image_revision(
    pixmap: &mut Pixmap,
    image: &ImagePrimitive,
    frame: FRect,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let Some(revision) = &image.attrs.revision else {
        return Ok(());
    };
    let color = match revision.kind {
        RevisionKind::Ins => "rgb(46, 125, 50)",
        RevisionKind::Del => "rgb(198, 40, 40)",
    };
    let rect = Rect::from_xywh(frame.x, frame.y, frame.w, frame.h)
        .ok_or_else(|| "invalid image revision rectangle".to_string())?;
    let path = PathBuilder::from_rect(rect);
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(color, opacity)?);
    paint.anti_alias = true;
    pixmap.stroke_path(
        &path,
        &paint,
        &Stroke {
            width: 2.0,
            ..Stroke::default()
        },
        transform,
        mask,
    );
    if revision.kind == RevisionKind::Del {
        stroke_segment(
            pixmap,
            Segment {
                x1: frame.x,
                y1: frame.y + frame.h,
                x2: frame.x + frame.w,
                y2: frame.y,
            },
            2.0,
            color,
            None,
            transform,
            mask,
            opacity,
            StrokeStyle::default(),
            budget,
        )?;
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct Segment {
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
}

#[derive(Clone, Copy)]
struct StrokeStyle {
    cap: LineCap,
    join: LineJoin,
    miter_limit: f32,
}

impl Default for StrokeStyle {
    fn default() -> Self {
        Self {
            cap: LineCap::Butt,
            join: LineJoin::Miter,
            miter_limit: 4.0,
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn paint_border_recipe(
    pixmap: &mut Pixmap,
    segment: Segment,
    width: f32,
    color: &str,
    style: DisplayBorderStyle,
    secondary_color: Option<&str>,
    dash_override: Option<&[f32]>,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    match style {
        DisplayBorderStyle::Wave | DisplayBorderStyle::DoubleWave => paint_wave(
            pixmap,
            segment,
            width,
            color,
            style == DisplayBorderStyle::DoubleWave,
            transform,
            mask,
            opacity,
            budget,
        ),
        DisplayBorderStyle::Double
        | DisplayBorderStyle::Triple
        | DisplayBorderStyle::ThinThick
        | DisplayBorderStyle::ThickThin => paint_compound_border(
            pixmap, segment, width, color, style, transform, mask, opacity, budget,
        ),
        DisplayBorderStyle::Groove
        | DisplayBorderStyle::Ridge
        | DisplayBorderStyle::Inset
        | DisplayBorderStyle::Outset => paint_three_d_border(
            pixmap,
            segment,
            width,
            color,
            secondary_color,
            style,
            transform,
            mask,
            opacity,
        ),
        _ => {
            let dash = if let Some(dash) = dash_override {
                Some(dash.to_vec())
            } else {
                border_dash(style, width)
            };
            stroke_segment(
                pixmap,
                segment,
                width,
                color,
                dash.as_deref(),
                transform,
                mask,
                opacity,
                StrokeStyle::default(),
                budget,
            )
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn stroke_segment(
    pixmap: &mut Pixmap,
    segment: Segment,
    width: f32,
    color: &str,
    dash: Option<&[f32]>,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    style: StrokeStyle,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let stroke = stroke_style(width, dash, style)?;
    if let Some(dash) = dash {
        budget.charge_path(dash_segments(segment_length(segment), dash))?;
    }
    let mut builder = PathBuilder::new();
    builder.move_to(segment.x1, segment.y1);
    builder.line_to(segment.x2, segment.y2);
    let path = builder
        .finish()
        .ok_or_else(|| "line has no drawable segment".to_string())?;
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(color, opacity)?);
    paint.anti_alias = true;
    pixmap.stroke_path(&path, &paint, &stroke, transform, mask);
    Ok(())
}

fn segment_length(segment: Segment) -> f32 {
    (segment.x2 - segment.x1).hypot(segment.y2 - segment.y1)
}

/// Subpaths a dash pattern expands a path of this length into, counted the way
/// tiny-skia counts them before it builds them. The length comes off the
/// display list, so it is charged before the expansion allocates.
fn dash_segments(length: f32, dash: &[f32]) -> u64 {
    let period = f64::from(dash.iter().sum::<f32>());
    let intervals = (dash.len() / 2).max(1) as f64;
    let count = (f64::from(length) * intervals / period).ceil();
    if count.is_nan() || count == f64::INFINITY {
        return u64::MAX;
    }
    if count <= 0.0 {
        return 0;
    }
    count.min(u64::MAX as f64) as u64
}

/// An upper bound on the length a stroke walks, since a curve is no longer
/// than the control polygon it is drawn from.
fn path_length(path: &Path) -> f32 {
    let mut total = 0.0;
    let mut cursor = Point::zero();
    let mut start = Point::zero();
    let mut step = |from: Point, to: Point| {
        total += (to.x - from.x).hypot(to.y - from.y);
    };
    for segment in path.segments() {
        match segment {
            PathSegment::MoveTo(point) => {
                cursor = point;
                start = point;
            }
            PathSegment::LineTo(point) => {
                step(cursor, point);
                cursor = point;
            }
            PathSegment::QuadTo(control, point) => {
                step(cursor, control);
                step(control, point);
                cursor = point;
            }
            PathSegment::CubicTo(first, second, point) => {
                step(cursor, first);
                step(first, second);
                step(second, point);
                cursor = point;
            }
            PathSegment::Close => {
                step(cursor, start);
                cursor = start;
            }
        }
    }
    total
}

fn stroke_style(width: f32, dash: Option<&[f32]>, style: StrokeStyle) -> Result<Stroke, String> {
    if !width.is_finite() || width < 0.0 {
        return Err("stroke width must be finite and non-negative".to_string());
    }
    let dash = dash
        .map(|values| {
            StrokeDash::new(values.to_vec(), 0.0)
                .ok_or_else(|| "invalid stroke dash pattern".to_string())
        })
        .transpose()?;
    Ok(Stroke {
        width,
        miter_limit: style.miter_limit,
        line_cap: style.cap,
        line_join: style.join,
        dash,
    })
}

fn border_dash(style: DisplayBorderStyle, width: f32) -> Option<Vec<f32>> {
    match style {
        DisplayBorderStyle::Dotted => Some(vec![width.max(1.0), (width * 1.5).max(1.5)]),
        DisplayBorderStyle::Dashed => Some(vec![(width * 4.0).max(2.0), (width * 2.0).max(2.0)]),
        DisplayBorderStyle::DashDot => Some(vec![width * 4.0, width * 1.5, width, width * 1.5]),
        DisplayBorderStyle::DashDotDot => Some(vec![
            width * 4.0,
            width * 1.5,
            width,
            width * 1.5,
            width,
            width * 1.5,
        ]),
        _ => None,
    }
}

#[allow(clippy::too_many_arguments)]
fn paint_compound_border(
    pixmap: &mut Pixmap,
    segment: Segment,
    width: f32,
    color: &str,
    style: DisplayBorderStyle,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let (normal_x, normal_y) = segment_normal(segment);
    let strokes = match style {
        DisplayBorderStyle::Triple => vec![
            (-width, (width / 3.0).max(0.5)),
            (0.0, (width / 3.0).max(0.5)),
            (width, (width / 3.0).max(0.5)),
        ],
        DisplayBorderStyle::ThinThick => vec![
            (-width * 0.45, (width * 0.25).max(0.5)),
            (width * 0.3, (width * 0.55).max(0.75)),
        ],
        DisplayBorderStyle::ThickThin => vec![
            (-width * 0.3, (width * 0.55).max(0.75)),
            (width * 0.45, (width * 0.25).max(0.5)),
        ],
        DisplayBorderStyle::Double => vec![
            (-width / 2.0, (width / 3.0).max(0.5)),
            (width / 2.0, (width / 3.0).max(0.5)),
        ],
        _ => return Err("invalid compound border style".to_string()),
    };
    for (offset, stroke_width) in strokes {
        stroke_segment(
            pixmap,
            Segment {
                x1: segment.x1 + normal_x * offset,
                y1: segment.y1 + normal_y * offset,
                x2: segment.x2 + normal_x * offset,
                y2: segment.y2 + normal_y * offset,
            },
            stroke_width,
            color,
            None,
            transform,
            mask,
            opacity,
            StrokeStyle::default(),
            budget,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn paint_wave(
    pixmap: &mut Pixmap,
    segment: Segment,
    width: f32,
    color: &str,
    double: bool,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let dx = segment.x2 - segment.x1;
    let dy = segment.y2 - segment.y1;
    let length = dx.hypot(dy);
    if length == 0.0 {
        return Ok(());
    }
    let tangent_x = dx / length;
    let tangent_y = dy / length;
    let normal_x = -tangent_y;
    let normal_y = tangent_x;
    let wavelength = (width * 4.0).max(4.0);
    let amplitude = width.max(1.0);
    let lanes: &[f32] = if double { &[-1.0, 1.0] } else { &[0.0] };
    budget.charge_path(wave_segments(length, wavelength).saturating_mul(lanes.len() as u64))?;
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(color, opacity)?);
    paint.anti_alias = true;
    for lane_factor in lanes {
        let lane = lane_factor * amplitude;
        let mut builder = PathBuilder::new();
        builder.move_to(segment.x1 + normal_x * lane, segment.y1 + normal_y * lane);
        let mut distance = 0.0;
        let mut sign = 1.0;
        while distance < length {
            let next = length.min(distance + wavelength / 2.0);
            let midpoint = (distance + next) / 2.0;
            builder.quad_to(
                segment.x1 + tangent_x * midpoint + normal_x * (lane + amplitude * sign),
                segment.y1 + tangent_y * midpoint + normal_y * (lane + amplitude * sign),
                segment.x1 + tangent_x * next + normal_x * lane,
                segment.y1 + tangent_y * next + normal_y * lane,
            );
            distance = next;
            sign = -sign;
        }
        let path = builder
            .finish()
            .ok_or_else(|| "wave border has no drawable path".to_string())?;
        pixmap.stroke_path(
            &path,
            &paint,
            &Stroke {
                width: (width / if double { 2.0 } else { 1.0 }).max(0.5),
                ..Stroke::default()
            },
            transform,
            mask,
        );
    }
    Ok(())
}

/// Quadratics one lane of a wave expands into. The length is a number off the
/// display list, so it is charged before the path builder grows to hold it.
/// Two individually finite coordinates subtract to an infinite length, which
/// nothing can afford: the loop that walks it never reaches its end.
fn wave_segments(length: f32, wavelength: f32) -> u64 {
    let steps = (f64::from(length) / f64::from(wavelength / 2.0)).ceil();
    if steps.is_nan() || steps == f64::INFINITY {
        return u64::MAX;
    }
    if steps <= 0.0 {
        return 0;
    }
    steps.min(u64::MAX as f64) as u64
}

#[allow(clippy::too_many_arguments)]
fn paint_three_d_border(
    pixmap: &mut Pixmap,
    segment: Segment,
    width: f32,
    color: &str,
    secondary_color: Option<&str>,
    style: DisplayBorderStyle,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
) -> Result<(), String> {
    let base = parse_color(color)?;
    let light = secondary_color
        .map(parse_color)
        .transpose()?
        .unwrap_or_else(|| mix_color(base, Color::WHITE, 0.55));
    let dark = mix_color(base, Color::BLACK, 0.45);
    let reverse = matches!(
        style,
        DisplayBorderStyle::Ridge | DisplayBorderStyle::Outset
    );
    let first = if reverse { light } else { dark };
    let second = if reverse { dark } else { light };
    let (normal_x, normal_y) = segment_normal(segment);
    let offset = width / 4.0;
    for (lane, lane_color) in [(-offset, first), (offset, second)] {
        stroke_segment_color(
            pixmap,
            Segment {
                x1: segment.x1 + normal_x * lane,
                y1: segment.y1 + normal_y * lane,
                x2: segment.x2 + normal_x * lane,
                y2: segment.y2 + normal_y * lane,
            },
            (width / 2.0).max(0.5),
            lane_color,
            transform,
            mask,
            opacity,
        )?;
    }
    Ok(())
}

fn stroke_segment_color(
    pixmap: &mut Pixmap,
    segment: Segment,
    width: f32,
    mut color: Color,
    transform: Transform,
    mask: Option<&Mask>,
    opacity: f32,
) -> Result<(), String> {
    color.apply_opacity(opacity);
    let mut builder = PathBuilder::new();
    builder.move_to(segment.x1, segment.y1);
    builder.line_to(segment.x2, segment.y2);
    let path = builder
        .finish()
        .ok_or_else(|| "line has no drawable segment".to_string())?;
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    pixmap.stroke_path(
        &path,
        &paint,
        &Stroke {
            width,
            ..Stroke::default()
        },
        transform,
        mask,
    );
    Ok(())
}

fn segment_normal(segment: Segment) -> (f32, f32) {
    let dx = segment.x2 - segment.x1;
    let dy = segment.y2 - segment.y1;
    let length = dx.hypot(dy).max(f32::EPSILON);
    (-dy / length, dx / length)
}

fn shape_dash_pattern(name: Option<&str>, width: f32) -> Result<Option<Vec<f32>>, String> {
    let Some(name) = name else {
        return Ok(None);
    };
    let pattern = match name {
        "" | "solid" => Vec::new(),
        "dot" | "dotted" | "sysDot" => vec![width.max(1.0), (width * 2.0).max(2.0)],
        "dash" | "dashed" | "dashSmallGap" | "sysDash" => {
            vec![(width * 3.0).max(2.0), (width * 2.0).max(2.0)]
        }
        "lgDash" | "dashLong" | "dashLongHeavy" => {
            vec![(width * 6.0).max(4.0), (width * 2.0).max(2.0)]
        }
        "dashDot" | "lgDashDot" | "sysDashDot" | "dashDotHeavy" => vec![
            (width * 3.0).max(2.0),
            (width * 2.0).max(2.0),
            width,
            (width * 2.0).max(2.0),
        ],
        "dashDotDot" | "lgDashDotDot" | "sysDashDotDot" | "dashDotDotHeavy" => vec![
            (width * 3.0).max(2.0),
            (width * 2.0).max(2.0),
            width,
            (width * 2.0).max(2.0),
            width,
            (width * 2.0).max(2.0),
        ],
        other => return Err(format!("unsupported shape dash style: {other}")),
    };
    Ok((!pattern.is_empty()).then_some(pattern))
}

fn paint_page_border(
    pixmap: &mut Pixmap,
    border: &PageBorderPrimitive,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let x = number_f32(&border.x)?;
    let y = number_f32(&border.y)?;
    let w = number_f32(&border.w)?;
    let h = number_f32(&border.h)?;
    if let Some(side) = &border.top {
        paint_page_border_side(
            pixmap,
            Segment {
                x1: x,
                y1: y,
                x2: x + w,
                y2: y,
            },
            side,
            budget,
        )?;
    }
    if let Some(side) = &border.right {
        paint_page_border_side(
            pixmap,
            Segment {
                x1: x + w,
                y1: y,
                x2: x + w,
                y2: y + h,
            },
            side,
            budget,
        )?;
    }
    if let Some(side) = &border.bottom {
        paint_page_border_side(
            pixmap,
            Segment {
                x1: x,
                y1: y + h,
                x2: x + w,
                y2: y + h,
            },
            side,
            budget,
        )?;
    }
    if let Some(side) = &border.left {
        paint_page_border_side(
            pixmap,
            Segment {
                x1: x,
                y1: y,
                x2: x,
                y2: y + h,
            },
            side,
            budget,
        )?;
    }
    Ok(())
}

fn paint_page_border_side(
    pixmap: &mut Pixmap,
    segment: Segment,
    side: &PageBorderSide,
    budget: &mut PageBudget,
) -> Result<(), String> {
    let style = match side.style.as_str() {
        "solid" => DisplayBorderStyle::Solid,
        "double" => DisplayBorderStyle::Double,
        "dotted" => DisplayBorderStyle::Dotted,
        "dashed" => DisplayBorderStyle::Dashed,
        "groove" => DisplayBorderStyle::Groove,
        "ridge" => DisplayBorderStyle::Ridge,
        "inset" => DisplayBorderStyle::Inset,
        "outset" => DisplayBorderStyle::Outset,
        other => return Err(format!("unsupported page border style: {other}")),
    };
    paint_border_recipe(
        pixmap,
        segment,
        number_f32(&side.width)?.max(0.5),
        &side.color,
        style,
        None,
        None,
        Transform::identity(),
        None,
        1.0,
        budget,
    )
}

fn primitive_attrs(primitive: &Primitive) -> &DocAttrs {
    match primitive {
        Primitive::Text(value) => &value.attrs,
        Primitive::GlyphRun(value) => &value.attrs,
        Primitive::Rect(value) => &value.attrs,
        Primitive::Line(value) => &value.attrs,
        Primitive::Image(value) => &value.attrs,
        Primitive::Shape(value) => &value.attrs,
        Primitive::Decoration(value) => &value.attrs,
    }
}

fn validate_visual_attrs(primitive: &Primitive, attrs: &DocAttrs) -> Result<(), String> {
    let is_text = matches!(primitive, Primitive::Text(_));
    let is_glyph = matches!(primitive, Primitive::GlyphRun(_));
    let is_image = matches!(primitive, Primitive::Image(_));
    let is_shape = matches!(primitive, Primitive::Shape(_));
    let is_decoration = matches!(primitive, Primitive::Decoration(_));
    if !is_image
        && (attrs.image_flip_h == Some(true)
            || attrs.image_flip_v == Some(true)
            || attrs.content_frame.is_some()
            || attrs.border.is_some())
    {
        return Err("image visual fields are attached to a non-image primitive".to_string());
    }
    if !is_shape
        && (attrs.fill_paint.is_some()
            || attrs.stroke_paint.is_some()
            || attrs.effect_extent.is_some()
            || attrs.drawing_scene.is_some()
            || attrs.text_body_properties.is_some())
    {
        return Err("shape visual fields are attached to a non-shape primitive".to_string());
    }
    if !is_image && !is_shape && !attrs.effects.is_empty() {
        return Err("effects are attached to an unsupported primitive".to_string());
    }
    if !is_decoration && attrs.style.is_some() {
        return Err("decoration style is attached to a non-decoration primitive".to_string());
    }
    if !is_text && attrs.leader_glyphs.is_some() {
        return Err("leaderGlyphs are attached to a non-text primitive".to_string());
    }
    if !is_text && !is_glyph && attrs.modern_effects.is_some() {
        return Err("modernEffects are attached to a non-text primitive".to_string());
    }
    if !is_glyph && attrs.fallback_font.is_some() {
        return Err("fallbackFont is attached to a non-glyph primitive".to_string());
    }
    Ok(())
}

fn primitive_opacity(primitive: &Primitive) -> Result<f32, String> {
    let attrs = primitive_attrs(primitive);
    let own = match primitive {
        Primitive::Text(value) => {
            if value.hidden && value.opacity.is_none() {
                Some(0.4)
            } else {
                value.opacity.as_ref().map(number_f32).transpose()?
            }
        }
        Primitive::GlyphRun(value) => {
            if value.hidden && value.opacity.is_none() {
                Some(0.4)
            } else {
                value.opacity.as_ref().map(number_f32).transpose()?
            }
        }
        Primitive::Line(value) => value.opacity.as_ref().map(number_f32).transpose()?,
        Primitive::Image(value) => value.opacity.as_ref().map(number_f32).transpose()?,
        Primitive::Rect(_) | Primitive::Shape(_) | Primitive::Decoration(_) => None,
    }
    .unwrap_or(1.0)
    .clamp(0.0, 1.0);
    let flattened = attrs
        .primitive_opacity
        .as_ref()
        .map(number_f32)
        .transpose()?
        .unwrap_or(1.0)
        .clamp(0.0, 1.0);
    let group = attrs
        .clip_group
        .as_ref()
        .and_then(|group| group.opacity.as_ref())
        .map(number_f32)
        .transpose()?
        .unwrap_or(1.0)
        .clamp(0.0, 1.0);
    Ok(own * flattened * group)
}

fn clip_rect(attrs: &DocAttrs) -> Result<Option<FRect>, String> {
    let Some(clip) = attrs
        .clip_group
        .as_ref()
        .and_then(|group| group.clip.as_ref())
    else {
        return Ok(None);
    };
    Ok(Some(FRect {
        x: clip.x.as_ref().map(number_f32).transpose()?.unwrap_or(0.0),
        y: clip.y.as_ref().map(number_f32).transpose()?.unwrap_or(0.0),
        w: clip
            .w
            .as_ref()
            .map(number_f32)
            .transpose()?
            .unwrap_or(0.0)
            .max(0.0),
        h: clip
            .h
            .as_ref()
            .map(number_f32)
            .transpose()?
            .unwrap_or(0.0)
            .max(0.0),
    }))
}

#[derive(Default)]
struct ClipSurface {
    rect: Option<FRect>,
    origin: (i32, i32),
    size: (u32, u32),
    pixmap: Option<Pixmap>,
    mask: Option<Mask>,
}

impl ClipSurface {
    /// The surface to allocate for a clip this size, or `None` to keep the one
    /// already there. Every clipped primitive clears and blits the whole
    /// surface, so a surface kept far larger than the clip it now serves costs
    /// more than reallocating it: one page-sized clip early on would otherwise
    /// charge every small clip after it for the whole page.
    fn resize(&self, width: u32, height: u32) -> Option<(u32, u32)> {
        let Some(pixmap) = &self.pixmap else {
            return Some((width, height));
        };
        if self.size.0 < width || self.size.1 < height {
            return Some((self.size.0.max(width), self.size.1.max(height)));
        }
        let held = u64::from(pixmap.width()).saturating_mul(u64::from(pixmap.height()));
        let needed = u64::from(width).saturating_mul(u64::from(height));
        (held > needed.saturating_mul(CLIP_SURFACE_SLACK)).then_some((width, height))
    }

    fn paint<'k, F>(
        &mut self,
        target: &mut Pixmap,
        clip: FRect,
        scratch: &mut Scratch<'k>,
        painter: F,
    ) -> Result<(), String>
    where
        F: FnOnce(&mut Pixmap, &mut Scratch<'k>, Transform, Option<&Mask>) -> Result<(), String>,
    {
        let Some(clip) = clipped_to_target(clip, target.width(), target.height()) else {
            return Ok(());
        };
        if self.rect != Some(clip) {
            let origin_x = clip.x.floor() as i32;
            let origin_y = clip.y.floor() as i32;
            let width = ((clip.x + clip.w).ceil() as i32 - origin_x) as u32;
            let height = ((clip.y + clip.h).ceil() as i32 - origin_y) as u32;
            if let Some(size) = self.resize(width, height) {
                scratch.budget.charge_clips(
                    u64::from(size.0)
                        .saturating_mul(u64::from(size.1))
                        .saturating_mul(CLIP_BYTES_PER_PIXEL),
                )?;
                self.pixmap = Some(
                    Pixmap::new(size.0, size.1).ok_or_else(|| "invalid clip size".to_string())?,
                );
                self.mask = Some(
                    Mask::new(size.0, size.1)
                        .ok_or_else(|| "invalid clip mask size".to_string())?,
                );
                self.size = size;
            }
            let mask = self
                .mask
                .as_mut()
                .ok_or_else(|| "clip mask is unavailable".to_string())?;
            let rect = Rect::from_xywh(
                clip.x - origin_x as f32,
                clip.y - origin_y as f32,
                clip.w,
                clip.h,
            )
            .ok_or_else(|| "invalid clip rectangle".to_string())?;
            let path = PathBuilder::from_rect(rect);
            mask.clear();
            mask.fill_path(&path, FillRule::Winding, true, Transform::identity());
            self.rect = Some(clip);
            self.origin = (origin_x, origin_y);
        }
        let pixmap = self
            .pixmap
            .as_mut()
            .ok_or_else(|| "clip surface is unavailable".to_string())?;
        pixmap.fill(Color::TRANSPARENT);
        let transform = Transform::from_translate(-(self.origin.0 as f32), -(self.origin.1 as f32));
        painter(pixmap, scratch, transform, self.mask.as_ref())?;
        target.draw_pixmap(
            self.origin.0,
            self.origin.1,
            pixmap.as_ref(),
            &PixmapPaint::default(),
            Transform::identity(),
            None,
        );
        Ok(())
    }
}

fn clipped_to_target(clip: FRect, width: u32, height: u32) -> Option<FRect> {
    let left = clip.x.max(0.0);
    let top = clip.y.max(0.0);
    let right = (clip.x + clip.w).min(width as f32);
    let bottom = (clip.y + clip.h).min(height as f32);
    (left.is_finite()
        && top.is_finite()
        && right.is_finite()
        && bottom.is_finite()
        && right > left
        && bottom > top)
        .then_some(FRect {
            x: left,
            y: top,
            w: right - left,
            h: bottom - top,
        })
}

pub(crate) fn primitive_visual_transform(
    rect: FRect,
    rotation: Option<&Number>,
    horizontal_scale: Option<&Number>,
) -> Result<Transform, String> {
    let mut transform = Transform::identity();
    if let Some(rotation) = rotation {
        transform = transform.pre_concat(Transform::from_rotate_at(
            number_f32(rotation)?,
            rect.x + rect.w / 2.0,
            rect.y + rect.h / 2.0,
        ));
    }
    if let Some(horizontal_scale) = horizontal_scale {
        let scale = number_f32(horizontal_scale)? / 100.0;
        if scale <= 0.0 {
            return Err("horizontalScale must be positive".to_string());
        }
        transform = transform.pre_concat(scale_at(scale, 1.0, rect.x, rect.y + rect.h / 2.0));
    }
    Ok(transform)
}

fn scale_at(scale_x: f32, scale_y: f32, x: f32, y: f32) -> Transform {
    Transform::from_translate(x, y)
        .pre_concat(Transform::from_scale(scale_x, scale_y))
        .pre_concat(Transform::from_translate(-x, -y))
}

fn primitive_rect(
    x: &Number,
    y: &Number,
    w: &Number,
    h: &Number,
    name: &str,
) -> Result<Rect, String> {
    let w = number_f32(w)?;
    let h = number_f32(h)?;
    if w < 0.0 || h < 0.0 {
        return Err(format!("{name} dimensions must be non-negative"));
    }
    Rect::from_xywh(number_f32(x)?, number_f32(y)?, w, h).ok_or_else(|| format!("invalid {name}"))
}

pub(crate) fn number_f32(number: &Number) -> Result<f32, String> {
    let value = number
        .as_f64()
        .ok_or_else(|| "display-list number is not finite".to_string())? as f32;
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| "display-list number is outside the raster coordinate range".to_string())
}

fn value_f32(value: &Value) -> Result<f32, String> {
    let value = value
        .as_f64()
        .ok_or_else(|| "visual field must be a number".to_string())? as f32;
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| "visual field is outside the raster coordinate range".to_string())
}

fn number_slice(values: &[Number]) -> Result<Vec<f32>, String> {
    values.iter().map(number_f32).collect()
}

fn value_number_array(value: &Value) -> Result<Vec<f32>, String> {
    value
        .as_array()
        .ok_or_else(|| "visual field must be a number array".to_string())?
        .iter()
        .map(value_f32)
        .collect()
}

/// The surface cap the published entry points enforce for themselves, so a
/// consumer of this crate is bounded without a facade in front of it.
fn validate_page_surface(width: u32, height: u32) -> Result<(), String> {
    if width > MAX_PAGE_DIM || height > MAX_PAGE_DIM {
        return Err(format!(
            "page is {width}x{height}px, past the {MAX_PAGE_DIM}px per-side cap"
        ));
    }
    if u64::from(width) * u64::from(height) > MAX_PAGE_PIXELS {
        return Err(format!(
            "page is {width}x{height}px, past the {MAX_PAGE_PIXELS}-pixel surface cap"
        ));
    }
    Ok(())
}

fn page_dimension(number: &Number, name: &str) -> Result<u32, String> {
    let value = number_f32(number)?;
    if value <= 0.0 || value.ceil() > u32::MAX as f32 {
        return Err(format!("{name} must be finite and positive"));
    }
    Ok((value.ceil() as u32).max(1))
}

fn solid_paint(color: &str, opacity: f32) -> Result<Paint<'static>, String> {
    let mut paint = Paint::default();
    paint.set_color(color_with_opacity(color, opacity)?);
    paint.anti_alias = true;
    Ok(paint)
}

pub(crate) fn color_with_opacity(value: &str, opacity: f32) -> Result<Color, String> {
    let mut color = parse_color(value)?;
    color.apply_opacity(opacity.clamp(0.0, 1.0));
    Ok(color)
}

fn parse_color(value: &str) -> Result<Color, String> {
    if value.eq_ignore_ascii_case("transparent") {
        return Ok(Color::TRANSPARENT);
    }
    if let Some(hex) = value.strip_prefix('#') {
        return parse_hex_color(hex).ok_or_else(|| format!("bad color: {value}"));
    }
    if let Some(body) = value
        .strip_prefix("rgb(")
        .and_then(|value| value.strip_suffix(')'))
    {
        return parse_rgb_function(body, false).ok_or_else(|| format!("bad color: {value}"));
    }
    if let Some(body) = value
        .strip_prefix("rgba(")
        .and_then(|value| value.strip_suffix(')'))
    {
        return parse_rgb_function(body, true).ok_or_else(|| format!("bad color: {value}"));
    }
    Err(format!("bad color: {value}"))
}

fn parse_hex_color(hex: &str) -> Option<Color> {
    let expanded;
    let hex = match hex.len() {
        3 | 4 => {
            expanded = hex
                .chars()
                .flat_map(|character| [character, character])
                .collect::<String>();
            expanded.as_str()
        }
        6 | 8 => hex,
        _ => return None,
    };
    let byte = |index: usize| u8::from_str_radix(hex.get(index..index + 2)?, 16).ok();
    Some(Color::from_rgba8(
        byte(0)?,
        byte(2)?,
        byte(4)?,
        if hex.len() == 8 { byte(6)? } else { 255 },
    ))
}

fn parse_rgb_function(body: &str, alpha: bool) -> Option<Color> {
    let values = body.split(',').map(str::trim).collect::<Vec<_>>();
    if values.len() != if alpha { 4 } else { 3 } {
        return None;
    }
    let channel = |value: &str| -> Option<u8> {
        if let Some(percent) = value.strip_suffix('%') {
            let value = percent.parse::<f32>().ok()?.clamp(0.0, 100.0);
            Some((value * 2.55).round() as u8)
        } else {
            Some(value.parse::<f32>().ok()?.clamp(0.0, 255.0).round() as u8)
        }
    };
    let alpha = if alpha {
        if let Some(percent) = values[3].strip_suffix('%') {
            (percent.parse::<f32>().ok()?.clamp(0.0, 100.0) * 2.55).round() as u8
        } else {
            (values[3].parse::<f32>().ok()?.clamp(0.0, 1.0) * 255.0).round() as u8
        }
    } else {
        255
    };
    Some(Color::from_rgba8(
        channel(values[0])?,
        channel(values[1])?,
        channel(values[2])?,
        alpha,
    ))
}

fn mix_color(left: Color, right: Color, amount: f32) -> Color {
    let amount = amount.clamp(0.0, 1.0);
    Color::from_rgba(
        left.red() + (right.red() - left.red()) * amount,
        left.green() + (right.green() - left.green()) * amount,
        left.blue() + (right.blue() - left.blue()) * amount,
        left.alpha() + (right.alpha() - left.alpha()) * amount,
    )
    .unwrap_or(left)
}

fn ensure_keys(
    object: &serde_json::Map<String, Value>,
    allowed: &[&str],
    context: &str,
) -> Result<(), String> {
    if let Some(key) = object.keys().find(|key| !allowed.contains(&key.as_str())) {
        return Err(format!("unsupported {context} field: {key}"));
    }
    Ok(())
}

fn encode_png(pixmap: Pixmap, width: u32, height: u32) -> Result<Vec<u8>, String> {
    let pixels = pixmap.take_demultiplied();
    let mut data = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut data, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().map_err(|error| error.to_string())?;
        writer
            .write_image_data(&pixels)
            .map_err(|error| error.to_string())?;
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn clip(x: f32, y: f32, w: f32, h: f32) -> FRect {
        FRect { x, y, w, h }
    }

    /// Every clipped primitive clears its surface and blits it whole, so a
    /// surface grown to a running maximum makes one page-sized clip charge
    /// every small clip after it for the whole page.
    #[test]
    fn a_clip_surface_shrinks_back_to_the_clip_it_serves() {
        let mut pixmap = Pixmap::new(800, 1120).expect("page");
        let mut glyphs = GlyphCache::default();
        let mut renderer = Renderer::new(800, 1120, &mut glyphs);
        let mut surface = ClipSurface::default();
        for rect in [clip(0.0, 0.0, 800.0, 1120.0), clip(8.0, 8.0, 64.0, 16.0)] {
            surface
                .paint(
                    &mut pixmap,
                    rect,
                    &mut renderer.scratch,
                    |_, _, _, _| Ok(()),
                )
                .expect("clipped paint");
        }
        assert_eq!(surface.size, (64, 16));
    }

    /// The surface is one reused allocation, so a page that keeps resizing it
    /// is charged its high-water mark rather than every resize.
    #[test]
    fn a_resized_clip_surface_is_charged_once_at_its_high_water_mark() {
        let mut pixmap = Pixmap::new(400, 560).expect("page");
        let mut glyphs = GlyphCache::default();
        let mut renderer = Renderer::new(400, 560, &mut glyphs);
        let mut surface = ClipSurface::default();
        for index in 0..32 {
            let rect = if index % 2 == 0 {
                clip(0.0, 0.0, 400.0, 560.0)
            } else {
                clip(8.0, index as f32, 64.0, 16.0)
            };
            surface
                .paint(
                    &mut pixmap,
                    rect,
                    &mut renderer.scratch,
                    |_, _, _, _| Ok(()),
                )
                .expect("clipped paint");
        }
        assert_eq!(
            renderer.scratch.budget.scratch,
            400 * 560 * CLIP_BYTES_PER_PIXEL
        );
    }
}
