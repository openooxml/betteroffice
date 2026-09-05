//! Native raster backend: paints a slide display list to a png via tiny-skia.
//! Server-side twin of the browser's canvas backend.

#[cfg(target_arch = "wasm32")]
compile_error!("betteroffice-pptx-raster is server-side only");

mod font;

pub use font::GlyphCache;

use std::collections::HashMap;
use std::io::Cursor;

use ooxml_drawingml::GeometryPathCommand;
use ooxml_text::{FontId, FontStore};
use pptx_render::{
    GradientType, ImageCrop, Paint as SlidePaint, Primitive, Stroke as SlideStroke,
    SurfaceDisplayList, Transform as SlideTransform,
};
use tiny_skia::{
    Color, ColorU8, FillRule, FilterQuality, GradientStop, IntSize, LinearGradient, Mask, Paint,
    Path, PathBuilder, Pixmap, PixmapPaint, Point, RadialGradient, Rect, SpreadMode, Stroke,
    StrokeDash, Transform,
};

/// Media bytes keyed by the `asset_id` an image primitive carries — an OPC part
/// path such as `ppt/media/image1.png`. Unlike a DOCX relationship id, that path
/// is already package-absolute, so no scoping key is needed. Entries borrow the
/// package, so building one costs pointers rather than a copy of every picture.
pub type AssetMap<'a> = HashMap<&'a str, &'a [u8]>;

/// One rendered slide's longest side.
pub const MAX_SLIDE_DIM: u32 = 16_384;
/// One rendered slide's surface.
pub const MAX_SLIDE_PIXELS: u64 = 16_777_216;
/// One decoded image, at twice the surface a slide may allocate.
pub const MAX_IMAGE_PIXELS: u64 = 33_554_432;
/// Every image decoded for one slide, at four times that surface.
pub const MAX_SLIDE_IMAGE_PIXELS: u64 = 67_108_864;
/// One decoded image's source buffer plus the pixmap it converts to, at
/// [`MAX_IMAGE_PIXELS`] of RGBA8.
pub const MAX_IMAGE_BYTES: u64 = 268_435_456;
/// The same, summed across every image one slide decodes.
pub const MAX_SLIDE_IMAGE_BYTES: u64 = 536_870_912;

const PLACEHOLDER_STROKE: &str = "#8a94a6";
const PLACEHOLDER_LABEL: &str = "#5d6675";
const PLACEHOLDER_LABEL_PX: f32 = 12.0;

/// Fonts and media the display list refers to by id.
pub struct RenderResources<'a> {
    pub fonts: &'a FontStore,
    pub images: &'a AssetMap<'a>,
    /// Face for placeholder labels. Without one the dashed box still draws, but
    /// its label does not.
    pub label_font: Option<FontId>,
}

impl<'a> RenderResources<'a> {
    pub fn new(fonts: &'a FontStore, images: &'a AssetMap<'a>) -> Self {
        Self {
            fonts,
            images,
            label_font: None,
        }
    }

    pub fn with_label_font(mut self, font: Option<FontId>) -> Self {
        self.label_font = font;
        self
    }
}

/// What fills the pixels the slide's own primitives leave uncovered.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum Background {
    /// Opaque white under the display list's own background — what the editor
    /// shows.
    #[default]
    Slide,
    /// Fully transparent, so the slide composites onto whatever is behind it.
    Transparent,
    /// A solid color under the display list's own background.
    Color(String),
}

#[derive(Clone, Debug, PartialEq)]
pub struct RenderOptions {
    /// Output scale, e.g. `2.0` for hidpi. The display list stays in CSS px;
    /// only the pixmap and its transform grow.
    pub scale: f32,
    pub background: Background,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self {
            scale: 1.0,
            background: Background::default(),
        }
    }
}

/// Png bytes plus what the render could not draw.
#[derive(Clone, PartialEq, Eq)]
pub struct RenderedSlide {
    pub bytes: Vec<u8>,
    pub width: u32,
    pub height: u32,
    /// Images whose bytes were missing, undecodable, or over budget. These are
    /// skipped and counted rather than failing the render.
    pub skipped_images: usize,
}

impl std::fmt::Debug for RenderedSlide {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("RenderedSlide")
            .field("width", &self.width)
            .field("height", &self.height)
            .field("bytes", &self.bytes.len())
            .field("skipped_images", &self.skipped_images)
            .finish()
    }
}

/// Paint a slide display list and encode it as png bytes.
pub fn render_png(
    dl: &SurfaceDisplayList,
    resources: &RenderResources<'_>,
) -> Result<Vec<u8>, String> {
    let mut cache = GlyphCache::default();
    Ok(render_slide_cached(dl, resources, &RenderOptions::default(), &mut cache)?.bytes)
}

/// Paint a slide display list at `options.scale`.
pub fn render_slide(
    dl: &SurfaceDisplayList,
    resources: &RenderResources<'_>,
    options: &RenderOptions,
) -> Result<RenderedSlide, String> {
    let mut cache = GlyphCache::default();
    render_slide_cached(dl, resources, options, &mut cache)
}

/// The same, reusing glyph outlines an earlier slide extracted. The cache binds
/// to the font store that fills it and refuses any other.
pub fn render_slide_cached(
    dl: &SurfaceDisplayList,
    resources: &RenderResources<'_>,
    options: &RenderOptions,
    glyphs: &mut GlyphCache,
) -> Result<RenderedSlide, String> {
    if !options.scale.is_finite() || options.scale <= 0.0 {
        return Err("render scale must be finite and positive".to_string());
    }
    let width = surface_dimension(dl.width, options.scale, "slide width")?;
    let height = surface_dimension(dl.height, options.scale, "slide height")?;
    if width > MAX_SLIDE_DIM || height > MAX_SLIDE_DIM {
        return Err(format!(
            "slide is {width}x{height}px, past the {MAX_SLIDE_DIM}px limit"
        ));
    }
    if u64::from(width) * u64::from(height) > MAX_SLIDE_PIXELS {
        return Err(format!(
            "slide is {width}x{height}px, past the {MAX_SLIDE_PIXELS}px limit"
        ));
    }

    let mut pixmap = Pixmap::new(width, height).ok_or("invalid pixmap size".to_string())?;
    match &options.background {
        Background::Slide => pixmap.fill(Color::WHITE),
        Background::Transparent => pixmap.fill(Color::TRANSPARENT),
        Background::Color(color) => pixmap.fill(parse_color(color)?),
    }

    let base = Transform::from_scale(options.scale, options.scale);
    if let Some(background) = &dl.background {
        let rect = frame_rect(0.0, 0.0, dl.width, dl.height)?;
        let paint = shader_paint(background, 0.0, 0.0, dl.width, dl.height)?;
        pixmap.fill_rect(rect, &paint, base, None);
    }

    let mut painter = Painter {
        pixmap: &mut pixmap,
        resources,
        glyphs,
        images: ImageCache::default(),
        skipped_images: 0,
    };
    for primitive in &dl.primitives {
        painter.paint(primitive, base, None)?;
    }
    let skipped_images = painter.skipped_images;

    Ok(RenderedSlide {
        bytes: encode_png(pixmap, width, height)?,
        width,
        height,
        skipped_images,
    })
}

fn surface_dimension(value: f32, scale: f32, label: &str) -> Result<u32, String> {
    let scaled = value * scale;
    if !scaled.is_finite() || scaled <= 0.0 {
        return Err(format!("{label} must be finite and positive"));
    }
    Ok((scaled.ceil() as u32).max(1))
}

fn encode_png(pixmap: Pixmap, width: u32, height: u32) -> Result<Vec<u8>, String> {
    let pixels = pixmap.take_demultiplied();
    let mut data = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut data, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().map_err(|e| e.to_string())?;
        writer
            .write_image_data(&pixels)
            .map_err(|e| e.to_string())?;
    }
    Ok(data)
}

struct Painter<'a, 'b> {
    pixmap: &'a mut Pixmap,
    resources: &'a RenderResources<'b>,
    glyphs: &'a mut GlyphCache,
    images: ImageCache,
    skipped_images: usize,
}

impl Painter<'_, '_> {
    fn paint(
        &mut self,
        primitive: &Primitive,
        base: Transform,
        clip: Option<&Mask>,
    ) -> Result<(), String> {
        let transform = base.pre_concat(local_transform(primitive));
        match primitive {
            Primitive::Shape {
                x,
                y,
                w,
                h,
                path,
                fill,
                stroke,
                ..
            } => self.paint_shape(
                *x,
                *y,
                *w,
                *h,
                path,
                fill.as_ref(),
                stroke.as_ref(),
                transform,
                clip,
            ),
            Primitive::Image {
                x,
                y,
                w,
                h,
                asset_id,
                crop,
                path,
                stroke,
                ..
            } => self.paint_image(
                *x,
                *y,
                *w,
                *h,
                asset_id.as_deref(),
                *crop,
                path.as_deref(),
                stroke.as_ref(),
                transform,
                clip,
            ),
            Primitive::TextBox {
                x, y, w, h, lines, ..
            } => {
                if lines.is_empty() {
                    return Ok(());
                }
                let Some(inner) = self.clipped(clip, *x, *y, *w, *h, transform)? else {
                    return Ok(());
                };
                font::paint_lines(
                    self.pixmap,
                    self.resources,
                    self.glyphs,
                    lines,
                    transform,
                    Some(&inner),
                )
            }
            Primitive::Placeholder {
                x, y, w, h, label, ..
            } => self.paint_placeholder(*x, *y, *w, *h, label.as_deref(), transform, clip),
            Primitive::Chart {
                x,
                y,
                w,
                h,
                primitives,
                ..
            } => {
                let Some(inner) = self.clipped(clip, *x, *y, *w, *h, transform)? else {
                    return Ok(());
                };
                for child in primitives {
                    self.paint(child, transform, Some(&inner))?;
                }
                Ok(())
            }
        }
    }

    /// The clip a child paints under: `clip` narrowed to this primitive's box.
    /// `None` means the box lands off the surface, so nothing inside it can
    /// draw and the mask is never allocated.
    fn clipped(
        &self,
        clip: Option<&Mask>,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        transform: Transform,
    ) -> Result<Option<Mask>, String> {
        let Ok(rect) = frame_rect(x, y, w, h) else {
            return Ok(None);
        };
        self.clipped_path(clip, PathBuilder::from_rect(rect), transform)
    }

    fn clipped_path(
        &self,
        clip: Option<&Mask>,
        path: Path,
        transform: Transform,
    ) -> Result<Option<Mask>, String> {
        let Some(path) = path.transform(transform) else {
            return Ok(None);
        };
        let bounds = path.bounds();
        let (width, height) = (self.pixmap.width() as f32, self.pixmap.height() as f32);
        if bounds.right() <= 0.0
            || bounds.bottom() <= 0.0
            || bounds.left() >= width
            || bounds.top() >= height
        {
            return Ok(None);
        }
        let mut mask = match clip {
            Some(existing) => existing.clone(),
            None => Mask::new(self.pixmap.width(), self.pixmap.height())
                .ok_or("invalid clip mask size".to_string())?,
        };
        // The path already carries the transform, so the mask takes identity.
        if clip.is_some() {
            mask.intersect_path(&path, FillRule::Winding, true, Transform::identity());
        } else {
            mask.fill_path(&path, FillRule::Winding, true, Transform::identity());
        }
        Ok(Some(mask))
    }

    #[allow(clippy::too_many_arguments)]
    fn paint_shape(
        &mut self,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        commands: &[GeometryPathCommand],
        fill: Option<&SlidePaint>,
        stroke: Option<&SlideStroke>,
        transform: Transform,
        clip: Option<&Mask>,
    ) -> Result<(), String> {
        let Some(path) = geometry_path(commands, x, y, w, h) else {
            return Ok(());
        };
        if let Some(fill) = fill {
            let paint = shader_paint(fill, x, y, w, h)?;
            self.pixmap
                .fill_path(&path, &paint, FillRule::Winding, transform, clip);
        }
        if let Some(stroke) = stroke {
            self.stroke_path(&path, stroke, transform, clip)?;
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn paint_image(
        &mut self,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        asset_id: Option<&str>,
        crop: ImageCrop,
        commands: Option<&[GeometryPathCommand]>,
        stroke: Option<&SlideStroke>,
        transform: Transform,
        clip: Option<&Mask>,
    ) -> Result<(), String> {
        let Ok(frame) = frame_rect(x, y, w, h) else {
            return Ok(());
        };
        let outline = match commands {
            Some(commands) => geometry_path(commands, x, y, w, h),
            None => Some(PathBuilder::from_rect(frame)),
        };
        let Some(outline) = outline else {
            return Ok(());
        };
        let (kept_x, kept_y) = crop.kept();
        if kept_x > 0.0 && kept_y > 0.0 {
            let mask = if !crop.is_whole() || commands.is_some() {
                let Some(mask) = self.clipped_path(clip, outline.clone(), transform)? else {
                    return Ok(());
                };
                Some(mask)
            } else {
                None
            };
            match asset_id.and_then(|asset_id| self.decode(asset_id)) {
                Some(source) => {
                    let fit = Transform::from_row(
                        frame.width() / (source.width() as f32 * kept_x),
                        0.0,
                        0.0,
                        frame.height() / (source.height() as f32 * kept_y),
                        frame.x() - crop.left * frame.width() / kept_x,
                        frame.y() - crop.top * frame.height() / kept_y,
                    );
                    self.pixmap.draw_pixmap(
                        0,
                        0,
                        source.as_ref(),
                        &PixmapPaint {
                            quality: FilterQuality::Bicubic,
                            ..PixmapPaint::default()
                        },
                        transform.pre_concat(fit),
                        mask.as_ref().or(clip),
                    );
                }
                None => self.skipped_images += 1,
            }
        }
        if let Some(stroke) = stroke {
            self.stroke_path(&outline, stroke, transform, clip)?;
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn paint_placeholder(
        &mut self,
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        label: Option<&str>,
        transform: Transform,
        clip: Option<&Mask>,
    ) -> Result<(), String> {
        let Ok(rect) = frame_rect(x, y, w, h) else {
            return Ok(());
        };
        let mut paint = Paint::default();
        paint.set_color(parse_color(PLACEHOLDER_STROKE)?);
        paint.anti_alias = true;
        let stroke = Stroke {
            width: 1.0,
            dash: StrokeDash::new(vec![5.0, 4.0], 0.0),
            ..Stroke::default()
        };
        self.pixmap.stroke_path(
            &PathBuilder::from_rect(rect),
            &paint,
            &stroke,
            transform,
            clip,
        );

        let (Some(label), Some(font)) = (label, self.resources.label_font) else {
            return Ok(());
        };
        font::paint_centered_label(
            self.pixmap,
            self.resources.fonts,
            self.glyphs,
            font,
            label,
            PLACEHOLDER_LABEL_PX,
            parse_color(PLACEHOLDER_LABEL)?,
            rect,
            transform,
            clip,
        )
    }

    fn stroke_path(
        &mut self,
        path: &Path,
        stroke: &SlideStroke,
        transform: Transform,
        clip: Option<&Mask>,
    ) -> Result<(), String> {
        let Some(paint) = stroke_paint(stroke)? else {
            return Ok(());
        };
        let (paint, stroke) = paint;
        self.pixmap
            .stroke_path(path, &paint, &stroke, transform, clip);
        Ok(())
    }

    fn decode(&mut self, asset_id: &str) -> Option<Pixmap> {
        let bytes = self.resources.images.get(asset_id)?;
        self.images.decode(bytes)
    }
}

/// The rotate-and-flip a primitive applies about its own centre, matching
/// `applyTransform` in the canvas backend.
fn local_transform(primitive: &Primitive) -> Transform {
    let (x, y, w, h, transform) = match primitive {
        Primitive::Shape {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::Image {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::TextBox {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::Placeholder {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::Chart {
            x,
            y,
            w,
            h,
            transform,
            ..
        } => (*x, *y, *w, *h, transform),
    };
    if transform.is_identity() {
        return Transform::identity();
    }
    let SlideTransform {
        rotation_deg,
        flip_h,
        flip_v,
    } = *transform;
    let center_x = x + w / 2.0;
    let center_y = y + h / 2.0;
    Transform::from_translate(center_x, center_y)
        .pre_concat(Transform::from_rotate(rotation_deg))
        .pre_concat(Transform::from_scale(
            if flip_h { -1.0 } else { 1.0 },
            if flip_v { -1.0 } else { 1.0 },
        ))
        .pre_concat(Transform::from_translate(-center_x, -center_y))
}

fn frame_rect(x: f32, y: f32, w: f32, h: f32) -> Result<Rect, String> {
    Rect::from_xywh(x, y, w, h).ok_or_else(|| format!("invalid rectangle {x},{y} {w}x{h}"))
}

/// Geometry commands are fractions of the primitive's box, so each coordinate
/// scales by the box before it is placed — the same mapping `buildPath` does in
/// the canvas backend.
fn geometry_path(commands: &[GeometryPathCommand], x: f32, y: f32, w: f32, h: f32) -> Option<Path> {
    if commands.is_empty() {
        return None;
    }
    let px = |value: f64| x + value as f32 * w;
    let py = |value: f64| y + value as f32 * h;
    let mut builder = PathBuilder::new();
    for command in commands {
        match *command {
            GeometryPathCommand::Move { x, y } => builder.move_to(px(x), py(y)),
            GeometryPathCommand::Line { x, y } => builder.line_to(px(x), py(y)),
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                builder.quad_to(px(cpx), py(cpy), px(x), py(y))
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => builder.cubic_to(px(cp1x), py(cp1y), px(cp2x), py(cp2y), px(x), py(y)),
            GeometryPathCommand::Close => builder.close(),
        }
    }
    builder.finish()
}

fn shader_paint(
    paint: &SlidePaint,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) -> Result<Paint<'static>, String> {
    match paint {
        SlidePaint::Solid { color } => {
            let mut solid = Paint::default();
            solid.set_color(parse_color(color)?);
            solid.anti_alias = true;
            Ok(solid)
        }
        SlidePaint::Gradient {
            gradient_type,
            angle_deg,
            stops,
        } => gradient_paint(*gradient_type, *angle_deg, stops, x, y, w, h),
    }
}

/// Gradient geometry matches the canvas backend: both kinds reach the corner of
/// the box, so raster and browser place the same stops.
fn gradient_paint(
    gradient_type: GradientType,
    angle_deg: Option<f32>,
    stops: &[pptx_render::GradientStop],
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) -> Result<Paint<'static>, String> {
    let mut colors = Vec::with_capacity(stops.len());
    for stop in stops {
        colors.push((stop.position.clamp(0.0, 1.0), parse_color(&stop.color)?));
    }
    let Some((_, first)) = colors.first().copied() else {
        return Err("gradient has no stops".to_string());
    };
    if colors.len() == 1 {
        let mut solid = Paint::default();
        solid.set_color(first);
        solid.anti_alias = true;
        return Ok(solid);
    }
    let converted = colors
        .into_iter()
        .map(|(position, color)| GradientStop::new(position, color))
        .collect::<Vec<_>>();

    let center_x = x + w / 2.0;
    let center_y = y + h / 2.0;
    let radius = w.hypot(h) / 2.0;
    let shader = match gradient_type {
        GradientType::Linear => {
            let radians = angle_deg.unwrap_or(0.0).to_radians();
            LinearGradient::new(
                Point::from_xy(
                    center_x - radians.cos() * radius,
                    center_y - radians.sin() * radius,
                ),
                Point::from_xy(
                    center_x + radians.cos() * radius,
                    center_y + radians.sin() * radius,
                ),
                converted,
                SpreadMode::Pad,
                Transform::identity(),
            )
        }
        GradientType::Radial | GradientType::Rectangular | GradientType::Path => {
            let center = Point::from_xy(center_x, center_y);
            RadialGradient::new(
                center,
                0.0,
                center,
                radius,
                converted,
                SpreadMode::Pad,
                Transform::identity(),
            )
        }
    }
    .ok_or_else(|| "invalid gradient".to_string())?;
    Ok(Paint {
        shader,
        anti_alias: true,
        ..Paint::default()
    })
}

/// The dash pattern mirrors `strokeCurrentPath` in the canvas backend, so a
/// dashed outline breaks at the same places in both.
fn stroke_paint(stroke: &SlideStroke) -> Result<Option<(Paint<'static>, Stroke)>, String> {
    if !stroke.width.is_finite() || stroke.width <= 0.0 {
        return Ok(None);
    }
    let mut paint = Paint::default();
    paint.set_color(parse_color(&stroke.color)?);
    paint.anti_alias = true;
    let dash = stroke.dashed.then(|| {
        StrokeDash::new(
            vec![3.0_f32.max(stroke.width * 2.0), 2.0_f32.max(stroke.width)],
            0.0,
        )
    });
    Ok(Some((
        paint,
        Stroke {
            width: stroke.width,
            dash: dash.flatten(),
            ..Stroke::default()
        },
    )))
}

/// Decoded images for one slide, budgeted so a hostile deck cannot allocate its
/// way out of the surface caps.
#[derive(Default)]
struct ImageCache {
    pixels: u64,
    bytes: u64,
}

impl ImageCache {
    /// Decoded pixels, or `None` for content this backend will not draw: bytes
    /// it cannot decode, an image past [`MAX_IMAGE_PIXELS`], or one the slide
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
        if declared > MAX_IMAGE_PIXELS || self.pixels + declared > MAX_SLIDE_IMAGE_PIXELS {
            return None;
        }
        let cost = decoder
            .total_bytes()
            .saturating_add(declared.saturating_mul(4));
        if cost > MAX_IMAGE_BYTES || self.bytes + cost > MAX_SLIDE_IMAGE_BYTES {
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

/// Colors arrive resolved to `#rrggbb` by the layout pass; the longer CSS forms
/// are accepted so a hand-built display list reads the same as the canvas one.
pub(crate) fn parse_color(value: &str) -> Result<Color, String> {
    if value.eq_ignore_ascii_case("transparent") {
        return Ok(Color::TRANSPARENT);
    }
    if let Some(hex) = value.strip_prefix('#') {
        return parse_hex_color(hex).ok_or_else(|| format!("bad color: {value}"));
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

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_list(width: f32, height: f32) -> SurfaceDisplayList {
        SurfaceDisplayList {
            contract_version: pptx_render::CONTRACT_VERSION,
            width,
            height,
            background: None,
            primitives: Vec::new(),
        }
    }

    fn resources<'a>(fonts: &'a FontStore, images: &'a AssetMap<'a>) -> RenderResources<'a> {
        RenderResources::new(fonts, images)
    }

    #[test]
    fn png_dimensions_follow_the_scale() {
        let fonts = FontStore::new();
        let images = AssetMap::default();
        let rendered = render_slide(
            &empty_list(100.0, 50.0),
            &resources(&fonts, &images),
            &RenderOptions {
                scale: 2.0,
                ..RenderOptions::default()
            },
        )
        .expect("render");
        assert_eq!((rendered.width, rendered.height), (200, 100));
        assert_eq!(
            rendered.bytes[..8],
            [0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]
        );
    }

    #[test]
    fn a_slide_past_the_surface_cap_is_refused_before_allocating() {
        let fonts = FontStore::new();
        let images = AssetMap::default();
        let error = render_slide(
            &empty_list(20_000.0, 10.0),
            &resources(&fonts, &images),
            &RenderOptions::default(),
        )
        .expect_err("refuse");
        assert!(error.contains("past the"), "{error}");
    }

    #[test]
    fn a_cropped_masked_picture_preserves_pixels_outline_and_parent_clip() {
        let fonts = FontStore::new();
        let mut source = Pixmap::new(100, 100).unwrap();
        for (i, pixel) in source.pixels_mut().iter_mut().enumerate() {
            *pixel =
                ColorU8::from_rgba((i % 100 * 2) as u8, (i / 100 * 2) as u8, 40, 255).premultiply();
        }
        let bytes = source.encode_png().unwrap();
        let images = AssetMap::from([("photo", bytes.as_slice())]);
        for flip_h in [false, true] {
            for parent_clip in [false, true] {
                let image = Primitive::Image {
                    object_id: 1,
                    shape_id: None,
                    name: "Photo".into(),
                    x: 100.0,
                    y: 50.0,
                    w: 200.0,
                    h: 100.0,
                    asset_id: Some("photo".into()),
                    crop: ImageCrop {
                        left: 0.1,
                        top: 0.2,
                        right: 0.3,
                        bottom: 0.1,
                    },
                    path: ooxml_drawingml::preset_geometry_to_path(
                        "ellipse",
                        &Default::default(),
                        2.0,
                    ),
                    stroke: Some(SlideStroke {
                        color: "#ff00ff".into(),
                        width: 2.0,
                        dashed: false,
                        head_end: None,
                        tail_end: None,
                    }),
                    transform: SlideTransform {
                        flip_h,
                        ..Default::default()
                    },
                };
                let mut list = empty_list(400.0, 200.0);
                list.primitives.push(if parent_clip {
                    Primitive::Chart {
                        object_id: 2,
                        shape_id: None,
                        name: "Parent".into(),
                        x: 175.0,
                        y: 0.0,
                        w: 225.0,
                        h: 200.0,
                        label: String::new(),
                        primitives: vec![image],
                        transform: SlideTransform::default(),
                    }
                } else {
                    image
                });
                let rendered = render_slide(
                    &list,
                    &resources(&fonts, &images),
                    &RenderOptions {
                        background: Background::Transparent,
                        ..Default::default()
                    },
                )
                .unwrap();
                assert_eq!(rendered.skipped_images, 0);
                let pixels = Pixmap::decode_png(&rendered.bytes).unwrap();
                let kept = pixels.pixel(250, 125).unwrap();
                let expected_red = if flip_h { 50 } else { 110 };
                assert!(kept.red().abs_diff(expected_red) <= 2, "{kept:?}");
                assert!(kept.green().abs_diff(145) <= 2, "{kept:?}");
                assert_eq!(kept.blue(), 40);
                assert_eq!(kept.alpha(), 255);
                assert_eq!(
                    pixels.pixel(150, 100).unwrap().alpha(),
                    if parent_clip { 0 } else { 255 }
                );
                assert_eq!(pixels.pixel(299, 51).unwrap().alpha(), 0);
                assert_eq!(pixels.pixel(250, 45).unwrap().alpha(), 0);
                let border = pixels.pixel(287, 75).unwrap();
                assert!(
                    border.red() > 100 && border.green() < 10 && border.blue() > 100,
                    "{border:?}"
                );
            }
        }
    }

    #[test]
    fn a_missing_asset_is_skipped_and_counted() {
        let fonts = FontStore::new();
        let images = AssetMap::default();
        for asset_id in [None, Some("ppt/media/image1.png")] {
            let mut list = empty_list(100.0, 100.0);
            list.primitives.push(Primitive::Image {
                object_id: 1,
                shape_id: None,
                name: "picture".into(),
                x: 0.0,
                y: 0.0,
                w: 50.0,
                h: 50.0,
                asset_id: asset_id.map(str::to_owned),
                crop: ImageCrop::default(),
                path: None,
                stroke: None,
                transform: SlideTransform::default(),
            });
            let rendered = render_slide(
                &list,
                &resources(&fonts, &images),
                &RenderOptions::default(),
            )
            .expect("render");
            assert_eq!(rendered.skipped_images, 1, "asset_id: {asset_id:?}");
        }
    }

    #[test]
    fn a_transparent_background_leaves_uncovered_pixels_clear() {
        let fonts = FontStore::new();
        let images = AssetMap::default();
        let rendered = render_slide(
            &empty_list(4.0, 4.0),
            &resources(&fonts, &images),
            &RenderOptions {
                background: Background::Transparent,
                ..RenderOptions::default()
            },
        )
        .expect("render");
        let decoded = Pixmap::decode_png(&rendered.bytes).expect("decode");
        assert!(decoded.pixels().iter().all(|pixel| pixel.alpha() == 0));
    }

    /// A clipped primitive off the surface must cost no mask at all, and must
    /// not smuggle pixels onto a surface it does not overlap.
    #[test]
    fn a_clipped_primitive_off_the_surface_draws_nothing() {
        let fonts = FontStore::new();
        let images = AssetMap::default();
        let mut list = empty_list(64.0, 64.0);
        list.primitives.push(Primitive::Chart {
            object_id: 1,
            shape_id: None,
            name: "chart".into(),
            x: 500.0,
            y: 500.0,
            w: 100.0,
            h: 100.0,
            label: "offscreen".into(),
            primitives: vec![Primitive::Shape {
                object_id: 2,
                shape_id: None,
                name: "bar".into(),
                x: 0.0,
                y: 0.0,
                w: 64.0,
                h: 64.0,
                geometry: "rect".into(),
                path: vec![
                    GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                    GeometryPathCommand::Close,
                ],
                adjust_values: Default::default(),
                fill: Some(SlidePaint::Solid {
                    color: "#ff0000".into(),
                }),
                stroke: None,
                transform: SlideTransform::default(),
            }],
            transform: SlideTransform::default(),
        });
        let rendered = render_slide(
            &list,
            &resources(&fonts, &images),
            &RenderOptions::default(),
        )
        .expect("render");
        let decoded = Pixmap::decode_png(&rendered.bytes).expect("decode");
        assert!(
            decoded
                .pixels()
                .iter()
                .all(|pixel| pixel.demultiply() == ColorU8::from_rgba(255, 255, 255, 255)),
            "an off-surface chart painted onto the slide"
        );
    }

    #[test]
    fn geometry_commands_scale_by_the_primitive_box() {
        let path = geometry_path(
            &[
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                GeometryPathCommand::Close,
            ],
            10.0,
            20.0,
            100.0,
            50.0,
        )
        .expect("path");
        let bounds = path.bounds();
        assert_eq!((bounds.left(), bounds.top()), (10.0, 20.0));
        assert_eq!((bounds.right(), bounds.bottom()), (110.0, 70.0));
    }

    #[test]
    fn colors_parse_the_forms_the_layout_pass_emits() {
        assert_eq!(
            parse_color("#ff0000").expect("red"),
            Color::from_rgba8(255, 0, 0, 255)
        );
        assert_eq!(
            parse_color("#f00").expect("red"),
            Color::from_rgba8(255, 0, 0, 255)
        );
        assert_eq!(
            parse_color("transparent").expect("clear"),
            Color::TRANSPARENT
        );
        assert!(parse_color("rebeccapurple").is_err());
    }
}
