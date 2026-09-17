//! Raster backend: paints a VSDX display-list page to PNG via tiny-skia.

use std::collections::HashMap;
use std::io::Cursor;

use ooxml_drawingml::GeometryPathCommand;
use ooxml_text::{FontId, FontStore, PathCmd};
use rustybuzz::ttf_parser;
use tiny_skia::{
    Color, FillRule, FilterQuality, GradientStop, IntSize, LinearGradient, Paint, Path,
    PathBuilder, Pixmap, PixmapPaint, Point, Rect, SpreadMode, Stroke, StrokeDash, Transform,
};
use vsdx_parse::VsdxPackage;
use vsdx_render::{
    Affine, Paint as VsPaint, Primitive, Renderer, Stroke as VsStroke, TextFragment,
    VsdxDisplayList, linear_gradient, text_fragments, z_order,
};

/// Carlito Regular (OFL), metric-compatible with Calibri.
const FONT_BYTES: &[u8] = include_bytes!("../../xlsx-raster/assets/Carlito-Regular.ttf");

/// Last-resort face when the caller registered nothing for a run's family.
fn fallback_face() -> Option<(FontStore, FontId)> {
    let mut store = FontStore::new();
    let id = store.register(FONT_BYTES.to_vec()).ok()?;
    Some((store, id))
}

/// One rendered page's longest side.
pub const MAX_PAGE_DIM: u32 = 16_384;
/// One rendered page's surface.
pub const MAX_PAGE_PIXELS: u64 = 16_777_216;
/// One decoded image.
pub const MAX_IMAGE_PIXELS: u64 = 33_554_432;
/// Synthetic bold's second pass, as a fraction of the em.
const BOLD_OFFSET_EM: f32 = 0.35 / 12.0;
/// Synthetic italic shear.
const ITALIC_SHEAR: f32 = 0.21;
/// Glyphs above this device size are skipped rather than rasterized.
const MAX_GLYPH_PX: f32 = 8192.0;
/// Baseline below a line's top edge, as a fraction of the em.
const BASELINE_EM: f32 = 0.8;

/// PNG bytes plus what the render could not draw.
pub struct RenderedPage {
    pub bytes: Vec<u8>,
    pub width: u32,
    pub height: u32,
    /// Images missing, undecodable, over budget, or in an unsupported format.
    pub skipped_images: usize,
}

/// Renders one diagram page to PNG at `scale` times the 96 dpi display list.
pub fn render_page(
    renderer: &Renderer,
    package: &VsdxPackage,
    page_index: usize,
    scale: f32,
) -> Result<RenderedPage, String> {
    let part = package
        .page_part_paths
        .get(page_index)
        .ok_or_else(|| format!("page index {page_index} is out of range"))?;
    let list = renderer
        .layout_page(package, part)
        .map_err(|error| error.to_string())?;
    let mut assets = Vec::new();
    for primitive in &list.primitives {
        collect_images(primitive, &mut assets);
    }
    let images: HashMap<&str, &[u8]> = assets
        .into_iter()
        .filter_map(|asset_id| package.part_bytes(asset_id).map(|bytes| (asset_id, bytes)))
        .collect();
    render_list(Some(renderer), &list, &images, scale)
}

fn collect_images<'a>(primitive: &'a Primitive, out: &mut Vec<&'a str>) {
    match primitive {
        Primitive::Image { asset_id, .. } => out.push(asset_id),
        Primitive::Group { primitives, .. } => {
            for child in primitives {
                collect_images(child, out);
            }
        }
        _ => {}
    }
}

/// Renders an already laid-out page with the faces `fonts` measured it with;
/// `images` maps asset ids to file bytes.
pub fn render_list(
    fonts: Option<&Renderer>,
    list: &VsdxDisplayList,
    images: &HashMap<&str, &[u8]>,
    scale: f32,
) -> Result<RenderedPage, String> {
    if !scale.is_finite() || scale <= 0.0 {
        return Err("png scale must be a finite positive number".to_owned());
    }
    let width = (list.width * scale).ceil().max(1.0) as u32;
    let height = (list.height * scale).ceil().max(1.0) as u32;
    if width > MAX_PAGE_DIM || height > MAX_PAGE_DIM {
        return Err(format!(
            "png page is {width}x{height}px, over the {MAX_PAGE_DIM}px limit on either side"
        ));
    }
    if u64::from(width) * u64::from(height) > MAX_PAGE_PIXELS {
        return Err(format!(
            "png page is {}px, over the {MAX_PAGE_PIXELS}px surface limit",
            u64::from(width) * u64::from(height)
        ));
    }
    let mut pixmap = Pixmap::new(width, height).ok_or_else(|| "invalid pixmap size".to_owned())?;
    pixmap.fill(Color::WHITE);
    let paint = list.paint_transform;
    let base = Affine {
        a: scale,
        b: 0.0,
        c: 0.0,
        d: scale,
        e: 0.0,
        f: 0.0,
    }
    .compose(Affine {
        a: paint.a,
        b: paint.b,
        c: paint.c,
        d: paint.d,
        e: paint.e,
        f: paint.f,
    });
    let fallback = fallback_face();
    let skipped = {
        let mut painter = Painter {
            pixmap: &mut pixmap,
            images,
            fonts,
            fallback: fallback.as_ref(),
            skipped_images: 0,
        };
        let mut top: Vec<&Primitive> = list.primitives.iter().collect();
        top.sort_by_key(|primitive| z_order(primitive));
        for primitive in top {
            painter.paint(primitive, base);
        }
        painter.skipped_images
    };
    let bytes = encode_png(pixmap, width, height)?;
    Ok(RenderedPage {
        bytes,
        width,
        height,
        skipped_images: skipped,
    })
}

fn tiny(transform: Affine) -> Transform {
    Transform::from_row(
        transform.a,
        transform.b,
        transform.c,
        transform.d,
        transform.e,
        transform.f,
    )
}

fn translate(x: f32, y: f32) -> Affine {
    Affine {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: x,
        f: y,
    }
}

const FLIP_Y: Affine = Affine {
    a: 1.0,
    b: 0.0,
    c: 0.0,
    d: -1.0,
    e: 0.0,
    f: 0.0,
};

/// Shadow colours carry an optional alpha byte; the blur radius is dropped.
fn shadow_color(value: &str) -> Option<Color> {
    let hex = value.strip_prefix('#')?;
    let (digits, alpha) = match hex.len() {
        6 => (hex, 255),
        8 => (&hex[..6], u8::from_str_radix(&hex[6..], 16).ok()?),
        _ => return None,
    };
    let channel = |range: std::ops::Range<usize>| u8::from_str_radix(digits.get(range)?, 16).ok();
    Some(Color::from_rgba8(
        channel(0..2)?,
        channel(2..4)?,
        channel(4..6)?,
        alpha,
    ))
}

fn parse_color(value: &str) -> Option<Color> {
    let hex = value.strip_prefix('#')?;
    if hex.len() != 6 {
        return None;
    }
    let channel = |range: std::ops::Range<usize>| u8::from_str_radix(hex.get(range)?, 16).ok();
    Some(Color::from_rgba8(
        channel(0..2)?,
        channel(2..4)?,
        channel(4..6)?,
        255,
    ))
}

struct Painter<'a, 'b> {
    pixmap: &'a mut Pixmap,
    images: &'b HashMap<&'b str, &'b [u8]>,
    fonts: Option<&'b Renderer>,
    fallback: Option<&'b (FontStore, FontId)>,
    skipped_images: usize,
}

/// A registered face plus the store that measured with it.
struct Face<'a> {
    store: &'a FontStore,
    id: FontId,
}

impl<'b> Painter<'_, 'b> {
    /// The face layout measured this family with, else the vendored fallback.
    fn face(&self, family: &str, bold: bool, italic: bool) -> Option<Face<'b>> {
        if let Some(renderer) = self.fonts
            && let Some(id) = renderer.font_id(family, bold, italic)
        {
            return Some(Face {
                store: renderer.fonts(),
                id,
            });
        }
        self.fallback.map(|(store, id)| Face { store, id: *id })
    }

    fn paint(&mut self, primitive: &Primitive, outer: Affine) {
        match primitive {
            Primitive::Shape {
                path,
                fill,
                stroke,
                shadow,
                transform,
                ..
            } => {
                let Some(shape) = build_path(path) else {
                    return;
                };
                let composed = tiny(outer.compose(*transform));
                if let Some(shadow) = shadow
                    && let Some(color) = shadow_color(&shadow.color)
                {
                    let mut paint = Paint::default();
                    paint.set_color(color);
                    paint.anti_alias = true;
                    let offset = outer
                        .compose(*transform)
                        .compose(translate(shadow.offset_x_in, shadow.offset_y_in));
                    self.pixmap
                        .fill_path(&shape, &paint, FillRule::Winding, tiny(offset), None);
                }
                if let Some(paint) = fill_paint(fill, path) {
                    self.pixmap
                        .fill_path(&shape, &paint, FillRule::Winding, composed, None);
                }
                if let Some((paint, stroke)) = stroke_paint(stroke) {
                    self.pixmap
                        .stroke_path(&shape, &paint, &stroke, composed, None);
                }
            }
            Primitive::TextBox {
                y,
                height,
                paragraphs,
                lines,
                transform,
                ..
            } => {
                let frame = outer
                    .compose(*transform)
                    .compose(translate(0.0, y * 2.0 + height))
                    .compose(FLIP_Y);
                for fragment in text_fragments(paragraphs, lines) {
                    self.text(&fragment, frame);
                }
            }
            Primitive::Image {
                asset_id,
                x,
                y,
                width,
                height,
                transform,
                ..
            } => {
                let composed = outer.compose(*transform);
                match self.decode(asset_id) {
                    Some((source, bitmap)) => {
                        let fit = Affine {
                            a: width / bitmap.0,
                            b: 0.0,
                            c: 0.0,
                            d: -height / bitmap.1,
                            e: *x,
                            f: *y + height,
                        };
                        self.pixmap.draw_pixmap(
                            0,
                            0,
                            source.as_ref(),
                            &PixmapPaint {
                                quality: FilterQuality::Bilinear,
                                ..PixmapPaint::default()
                            },
                            tiny(composed.compose(fit)),
                            None,
                        );
                    }
                    None => {
                        self.skipped_images += 1;
                        self.placeholder_rect(*x, *y, *width, *height, composed);
                    }
                }
            }
            Primitive::Placeholder {
                x,
                y,
                width,
                height,
                reason,
                ..
            } => {
                self.placeholder_rect(*x, *y, *width, *height, outer);
                if !reason.is_empty() {
                    self.label(reason, *x, *y, outer);
                }
            }
            Primitive::Group {
                primitives,
                transform,
                ..
            } => {
                let composed = outer.compose(*transform);
                let mut children: Vec<&Primitive> = primitives.iter().collect();
                children.sort_by_key(|primitive| z_order(primitive));
                for child in children {
                    self.paint(child, composed);
                }
            }
        }
    }

    fn text(&mut self, fragment: &TextFragment, frame: Affine) {
        if fragment.size_in <= 0.0 || !fragment.size_in.is_finite() {
            return;
        }
        let device = device_scale(frame);
        let size_px = fragment.size_in * device;
        if !size_px.is_finite() || size_px <= 0.0 || size_px > MAX_GLYPH_PX {
            return;
        }
        let Some(color) = parse_color(&fragment.color) else {
            return;
        };
        let Some(face) = self.face(&fragment.family, fragment.bold, fragment.italic) else {
            return;
        };
        let space = frame.compose(translate(fragment.x, fragment.line_y));
        let baseline = fragment.size_in * BASELINE_EM;
        let mut paint = Paint::default();
        paint.set_color(color);
        paint.anti_alias = true;
        let shear = if fragment.italic { ITALIC_SHEAR } else { 0.0 };
        let spacing = if fragment.letter_spacing.is_finite() {
            fragment.letter_spacing
        } else {
            0.0
        };
        let bold_shift = fragment.bold.then_some(BOLD_OFFSET_EM * fragment.size_in);
        for extra in [0.0].into_iter().chain(bold_shift) {
            self.glyphs(
                &face,
                &fragment.text,
                extra,
                baseline,
                fragment.size_in,
                shear,
                spacing,
                &paint,
                space,
            );
        }
        if fragment.underline {
            let width = advance(&face, &fragment.text, fragment.size_in)
                + spacing * fragment.text.chars().count().max(1) as f32;
            self.underline(
                &face,
                width,
                baseline,
                fragment.size_in,
                device,
                &paint,
                space,
            );
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn glyphs(
        &mut self,
        face: &Face<'_>,
        text: &str,
        anchor_x: f32,
        baseline: f32,
        size_in: f32,
        shear: f32,
        spacing: f32,
        paint: &Paint,
        space: Affine,
    ) {
        let Ok(shaped) = ooxml_text::shape(face.store, face.id, text, size_in, &[]) else {
            return;
        };
        let mut pen = anchor_x;
        for glyph in &shaped {
            let outline = face
                .store
                .outline_glyph(face.id, glyph.glyph_id as u16)
                .ok()
                .filter(|outline| outline.upem > 0);
            if let Some(outline) = outline {
                let scale = size_in / f32::from(outline.upem);
                if let Some(path) = glyph_path(&outline.cmds) {
                    let placed = Affine {
                        a: scale,
                        b: 0.0,
                        c: shear * scale,
                        d: -scale,
                        e: pen + glyph.x_offset,
                        f: baseline - glyph.y_offset,
                    };
                    self.pixmap.fill_path(
                        &path,
                        paint,
                        FillRule::Winding,
                        tiny(space.compose(placed)),
                        None,
                    );
                }
            }
            pen += glyph.x_advance + spacing;
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn underline(
        &mut self,
        face: &Face<'_>,
        width: f32,
        baseline: f32,
        size_in: f32,
        device: f32,
        paint: &Paint,
        space: Affine,
    ) {
        let (position, thickness) = underline_metrics(face);
        let height = (thickness * size_in).max(0.5 / device);
        let cy = baseline - position * size_in;
        if width > 0.0
            && let Some(rect) = Rect::from_xywh(0.0, cy - height / 2.0, width, height)
        {
            self.pixmap.fill_path(
                &PathBuilder::from_rect(rect),
                paint,
                FillRule::Winding,
                tiny(space),
                None,
            );
        }
    }

    fn label(&mut self, reason: &str, x: f32, y: f32, outer: Affine) {
        let Some(face) = self.face("sans-serif", false, false) else {
            return;
        };
        let size_in = 10.0 / 72.0;
        let mut paint = Paint::default();
        paint.set_color(Color::from_rgba8(0x5d, 0x66, 0x75, 255));
        paint.anti_alias = true;
        let space = outer.compose(translate(x, y)).compose(FLIP_Y);
        self.glyphs(
            &face,
            reason,
            0.0,
            size_in * BASELINE_EM,
            size_in,
            0.0,
            0.0,
            &paint,
            space,
        );
    }

    fn placeholder_rect(&mut self, x: f32, y: f32, width: f32, height: f32, outer: Affine) {
        let (xa, ya) = outer.apply_point(x, y);
        let (xb, yb) = outer.apply_point(x + width, y + height);
        let (x0, x1) = (xa.min(xb), xa.max(xb));
        let (y0, y1) = (ya.min(yb), ya.max(yb));
        let Some(rect) = Rect::from_xywh(x0, y0, (x1 - x0).max(0.0), (y1 - y0).max(0.0)) else {
            return;
        };
        let unit = device_scale(outer).max(1.0);
        let mut paint = Paint::default();
        paint.set_color(Color::from_rgba8(0x8a, 0x94, 0xa6, 255));
        paint.anti_alias = true;
        let stroke = Stroke {
            width: unit,
            dash: StrokeDash::new(vec![5.0 * unit, 4.0 * unit], 0.0),
            ..Stroke::default()
        };
        self.pixmap.stroke_path(
            &PathBuilder::from_rect(rect),
            &paint,
            &stroke,
            Transform::identity(),
            None,
        );
    }

    fn decode(&mut self, asset_id: &str) -> Option<(Pixmap, (f32, f32))> {
        use image::ImageDecoder as _;
        let bytes = self.images.get(asset_id)?;
        let decoder = image::ImageReader::new(Cursor::new(bytes))
            .with_guessed_format()
            .ok()?
            .into_decoder()
            .ok()?;
        let (declared_width, declared_height) = decoder.dimensions();
        let declared = u64::from(declared_width) * u64::from(declared_height);
        if declared == 0 || declared > MAX_IMAGE_PIXELS {
            return None;
        }
        let decoded = image::load_from_memory(bytes).ok()?.into_rgba8();
        let (width, height) = (decoded.width(), decoded.height());
        let size = IntSize::from_wh(width, height)?;
        let mut data = decoded.into_raw();
        for pixel in data.as_chunks_mut::<4>().0 {
            let color =
                tiny_skia::ColorU8::from_rgba(pixel[0], pixel[1], pixel[2], pixel[3]).premultiply();
            *pixel = [color.red(), color.green(), color.blue(), color.alpha()];
        }
        Pixmap::from_vec(data, size).map(|pixmap| (pixmap, (width as f32, height as f32)))
    }
}

fn device_scale(transform: Affine) -> f32 {
    transform.a.hypot(transform.b).max(0.0)
}

fn advance(face: &Face<'_>, text: &str, size_in: f32) -> f32 {
    ooxml_text::shape(face.store, face.id, text, size_in, &[])
        .map(|glyphs| glyphs.iter().map(|glyph| glyph.x_advance).sum())
        .unwrap_or(0.0)
}

/// Underline position and thickness as fractions of the em.
fn underline_metrics(face: &Face<'_>) -> (f32, f32) {
    face.store
        .font_bytes(face.id)
        .ok()
        .and_then(|bytes| ttf_parser::Face::parse(bytes, 0).ok())
        .and_then(|parsed| {
            let em = f32::from(parsed.units_per_em());
            parsed.underline_metrics().map(|metrics| {
                (
                    f32::from(metrics.position) / em,
                    f32::from(metrics.thickness) / em,
                )
            })
        })
        .unwrap_or((-0.1, 0.05))
}

fn build_path(path: &[GeometryPathCommand]) -> Option<Path> {
    let mut builder = PathBuilder::new();
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } => builder.move_to(*x as f32, *y as f32),
            GeometryPathCommand::Line { x, y } => builder.line_to(*x as f32, *y as f32),
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                builder.quad_to(*cpx as f32, *cpy as f32, *x as f32, *y as f32);
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => builder.cubic_to(
                *cp1x as f32,
                *cp1y as f32,
                *cp2x as f32,
                *cp2y as f32,
                *x as f32,
                *y as f32,
            ),
            GeometryPathCommand::Close => builder.close(),
        }
    }
    builder.finish()
}

fn solid_paint(color: Color) -> Paint<'static> {
    let mut paint = Paint::default();
    paint.set_color(color);
    paint.anti_alias = true;
    paint
}

fn fill_paint(fill: &Option<VsPaint>, path: &[GeometryPathCommand]) -> Option<Paint<'static>> {
    let fill = fill.as_ref()?;
    if let Some(gradient) = linear_gradient(fill, path) {
        let stops = gradient
            .stops
            .iter()
            .filter_map(|(position, color)| {
                parse_color(color).map(|color| GradientStop::new(*position, color))
            })
            .collect();
        return Some(Paint {
            shader: LinearGradient::new(
                Point::from_xy(gradient.start.0, gradient.start.1),
                Point::from_xy(gradient.end.0, gradient.end.1),
                stops,
                SpreadMode::Pad,
                Transform::identity(),
            )?,
            anti_alias: true,
            ..Paint::default()
        });
    }
    match fill {
        VsPaint::Solid { color } => parse_color(color).map(solid_paint),
        VsPaint::Gradient { stops, .. } => stops
            .iter()
            .find_map(|stop| parse_color(&stop.color))
            .map(solid_paint),
    }
}

fn stroke_paint(stroke: &Option<VsStroke>) -> Option<(Paint<'_>, Stroke)> {
    let stroke = stroke.as_ref()?;
    let mut paint = Paint::default();
    paint.set_color(parse_color(&stroke.color)?);
    paint.anti_alias = true;
    let width = f64::from(stroke.width).max(0.0) as f32;
    if !width.is_finite() {
        return None;
    }
    Some((
        paint,
        Stroke {
            width,
            dash: stroke
                .dashed
                .then(|| StrokeDash::new(vec![width * 2.0, width * 2.0], 0.0))
                .flatten(),
            ..Stroke::default()
        },
    ))
}

fn glyph_path(cmds: &[PathCmd]) -> Option<Path> {
    let mut builder = PathBuilder::new();
    for cmd in cmds {
        match *cmd {
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
    builder.finish()
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
    use vsdx_render::{PaintTransform, Primitive};

    fn package(path: &str) -> VsdxPackage {
        let root = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let source = std::fs::read(root.join("../vsdx-parse/tests/fixtures").join(path)).unwrap();
        vsdx_parse::parse_vsdx(&source).unwrap()
    }

    fn empty_list(width_in: f32, height_in: f32) -> VsdxDisplayList {
        let (width, height) = (
            width_in * vsdx_render::PIXELS_PER_INCH,
            height_in * vsdx_render::PIXELS_PER_INCH,
        );
        VsdxDisplayList {
            contract_version: vsdx_render::CONTRACT_VERSION,
            width,
            height,
            print_width: width,
            print_height: height,
            paint_transform: PaintTransform {
                a: 96.0,
                b: 0.0,
                c: 0.0,
                d: -96.0,
                e: 0.0,
                f: height,
            },
            primitives: Vec::new(),
            connectors: Vec::new(),
        }
    }

    fn png_size(bytes: &[u8]) -> (u32, u32) {
        assert_eq!(&bytes[0..8], &[137, 80, 78, 71, 13, 10, 26, 10]);
        assert_eq!(&bytes[12..16], b"IHDR");
        (
            u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]),
            u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]),
        )
    }

    #[test]
    fn renders_white_pages_at_display_list_size_times_scale() {
        let images = HashMap::new();
        let first = render_list(None, &empty_list(8.5, 11.0), &images, 1.0).unwrap();
        assert_eq!((first.width, first.height), (816, 1056));
        assert_eq!(png_size(&first.bytes), (816, 1056));
        let second = render_list(None, &empty_list(8.5, 11.0), &images, 2.0).unwrap();
        assert_eq!((second.width, second.height), (1632, 2112));
        assert_eq!(png_size(&second.bytes), (1632, 2112));
    }

    #[test]
    fn rejects_bad_scales_and_pages() {
        let package = package("foundation.vsdx");
        let renderer = Renderer::default();
        for scale in [0.0, -1.0, f32::NAN, f32::INFINITY] {
            assert!(render_page(&renderer, &package, 0, scale).is_err());
        }
        assert!(render_page(&renderer, &package, 99, 1.0).is_err());
        assert!(render_page(&renderer, &package, 0, 64.0).is_err());
    }

    #[test]
    fn renders_are_deterministic() {
        let package = package("text-accounting.vsdx");
        let renderer = Renderer::default();
        let first = render_page(&renderer, &package, 0, 1.0).unwrap();
        let second = render_page(&renderer, &package, 0, 1.0).unwrap();
        assert_eq!(first.bytes, second.bytes);
    }

    #[test]
    fn a_shape_shadow_paints_an_offset_silhouette_behind_the_shape() {
        let rect = |shadow| Primitive::Shape {
            id: "rect".into(),
            z_order: 0,
            path: vec![
                GeometryPathCommand::Move { x: 0.0, y: 1.0 },
                GeometryPathCommand::Line { x: 2.0, y: 1.0 },
                GeometryPathCommand::Line { x: 2.0, y: 2.0 },
                GeometryPathCommand::Line { x: 0.0, y: 2.0 },
                GeometryPathCommand::Close,
            ],
            fill: Some(vsdx_render::Paint::Solid {
                color: "#111111".into(),
            }),
            stroke: None,
            shadow,
            transform: Affine::identity(),
            diagnostics: Vec::new(),
        };
        let ink = |primitive| {
            let mut list = empty_list(4.0, 4.0);
            list.primitives.push(primitive);
            let page = render_list(None, &list, &HashMap::new(), 1.0).unwrap();
            ink_bounds(&page)
        };
        let plain = ink(rect(None));
        let shaded = ink(rect(Some(vsdx_render::Shadow {
            color: "#33445580".into(),
            blur_in: 0.0,
            offset_x_in: 0.25,
            offset_y_in: -0.25,
        })));
        assert_eq!(shaded.0, plain.0, "the shadow never moves the left edge");
        assert_eq!(shaded.1, plain.1, "the shadow never moves the top edge");
        assert_eq!(shaded.2, plain.2 + 24, "the shadow extends right by 0.25in");
        assert_eq!(shaded.3, plain.3 + 24, "the shadow extends down by 0.25in");
    }

    #[test]
    fn dark_geometry_leaves_marks() {
        let mut list = empty_list(4.0, 4.0);
        list.primitives.push(Primitive::Shape {
            id: "rect".into(),
            z_order: 0,
            path: vec![
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 2.0, y: 0.0 },
                GeometryPathCommand::Line { x: 2.0, y: 1.0 },
                GeometryPathCommand::Line { x: 0.0, y: 1.0 },
                GeometryPathCommand::Close,
            ],
            fill: Some(vsdx_render::Paint::Solid {
                color: "#111111".into(),
            }),
            stroke: Some(vsdx_render::Stroke {
                color: "#222222".into(),
                width: 0.02,
                dashed: false,
            }),
            shadow: None,
            transform: Affine::identity(),
            diagnostics: Vec::new(),
        });
        let page = render_list(None, &list, &HashMap::new(), 1.0).unwrap();
        let decoded = image::load_from_memory(&page.bytes).unwrap().into_rgba8();
        let marked = decoded
            .pixels()
            .filter(|pixel| pixel.0 != [255, 255, 255, 255])
            .count();
        assert!(marked > 100, "expected painted pixels, found {marked}");
    }

    #[test]
    fn text_fixture_leaves_marks() {
        let package = package("text-accounting.vsdx");
        let renderer = Renderer::default();
        let page = render_page(&renderer, &package, 0, 1.0).unwrap();
        let decoded = image::load_from_memory(&page.bytes).unwrap().into_rgba8();
        let marked = decoded
            .pixels()
            .filter(|pixel| pixel.0 != [255, 255, 255, 255])
            .count();
        assert!(marked > 100, "expected painted text, found {marked}");
    }

    #[test]
    fn counts_missing_images_without_failing() {
        let mut list = empty_list(4.0, 4.0);
        list.primitives.push(Primitive::Image {
            id: "missing".into(),
            z_order: 0,
            asset_id: "visio/media/gone.png".into(),
            x: 0.0,
            y: 0.0,
            width: 1.0,
            height: 1.0,
            transform: Affine::identity(),
        });
        let page = render_list(None, &list, &HashMap::new(), 1.0).unwrap();
        assert_eq!(page.skipped_images, 1);
    }

    #[test]
    fn draws_a_decoded_image() {
        let mut data = Vec::new();
        {
            let mut encoder = png::Encoder::new(&mut data, 2, 2);
            encoder.set_color(png::ColorType::Rgba);
            encoder.set_depth(png::BitDepth::Eight);
            let mut writer = encoder.write_header().unwrap();
            writer
                .write_image_data(&[200u8, 30, 30, 255].repeat(4))
                .unwrap();
        }
        let mut list = empty_list(4.0, 4.0);
        list.primitives.push(Primitive::Image {
            id: "red".into(),
            z_order: 0,
            asset_id: "a".into(),
            x: 0.0,
            y: 0.0,
            width: 2.0,
            height: 2.0,
            transform: Affine::identity(),
        });
        let mut images: HashMap<&str, &[u8]> = HashMap::new();
        images.insert("a", &data);
        let page = render_list(None, &list, &images, 1.0).unwrap();
        assert_eq!(page.skipped_images, 0);
        let decoded = image::load_from_memory(&page.bytes).unwrap().into_rgba8();
        let red = decoded
            .pixels()
            .filter(|pixel| pixel[0] > 150 && pixel[1] < 100)
            .count();
        assert!(red > 100, "expected red pixels, found {red}");
    }

    const IDENTITY: &str = r#"{"a":1,"b":0,"c":0,"d":1,"e":0,"f":0}"#;

    /// One 4x4 inch page holding a single-line text box under `transform`.
    fn text_list(text: &str, family: &str, transform: &str) -> VsdxDisplayList {
        serde_json::from_str(&format!(
            r##"{{"contractVersion":7,"width":384,"height":384,
                "printWidth":384,"printHeight":384,
                "paintTransform":{{"a":96,"b":0,"c":0,"d":-96,"e":0,"f":384}},
                "primitives":[{{"kind":"textBox","id":"t","zOrder":0,
                  "x":1,"y":1,"width":2,"height":1,
                  "paragraphs":[{{"runs":[{{"text":"{text}","family":"{family}",
                    "sizeIn":0.2,"bold":false,"italic":false,"color":"#000000"}}]}}],
                  "lines":[{{"x":1,"y":1.1,"width":2,"height":0.24,
                    "start":0,"end":{end},"caretStops":[]}}],
                  "transform":{transform}}}]}}"##,
            end = text.len()
        ))
        .unwrap()
    }

    /// Bounds of the painted pixels as `(min_x, min_y, max_x, max_y)`.
    fn ink_bounds(page: &RenderedPage) -> (u32, u32, u32, u32) {
        let decoded = image::load_from_memory(&page.bytes).unwrap().into_rgba8();
        let (mut min_x, mut min_y, mut max_x, mut max_y) = (u32::MAX, u32::MAX, 0, 0);
        for (x, y, pixel) in decoded.enumerate_pixels() {
            if pixel.0 != [255, 255, 255, 255] {
                min_x = min_x.min(x);
                min_y = min_y.min(y);
                max_x = max_x.max(x);
                max_y = max_y.max(y);
            }
        }
        assert!(min_x <= max_x && min_y <= max_y, "page has no painted text");
        (min_x, min_y, max_x, max_y)
    }

    #[test]
    fn text_paints_inside_its_own_box() {
        let list = text_list("Hamburgefonstiv", "Arial", IDENTITY);
        let page = render_list(None, &list, &HashMap::new(), 1.0).unwrap();
        let (min_x, min_y, max_x, max_y) = ink_bounds(&page);
        assert!(
            min_x >= 96 && max_x <= 288,
            "x ran outside the box: {min_x}..{max_x}"
        );
        assert!(
            min_y >= 192 && max_y <= 288,
            "y ran outside the box: {min_y}..{max_y}"
        );
    }

    #[test]
    fn rotated_text_boxes_paint_rotated_glyphs() {
        let upright = render_list(
            None,
            &text_list("Hamburgefonstiv", "Arial", IDENTITY),
            &HashMap::new(),
            1.0,
        )
        .unwrap();
        let quarter = render_list(
            None,
            &text_list(
                "Hamburgefonstiv",
                "Arial",
                r#"{"a":0,"b":1,"c":-1,"d":0,"e":3,"f":0}"#,
            ),
            &HashMap::new(),
            1.0,
        )
        .unwrap();
        let (ux0, uy0, ux1, uy1) = ink_bounds(&upright);
        let (rx0, ry0, rx1, ry1) = ink_bounds(&quarter);
        assert!(
            ux1 - ux0 > uy1 - uy0,
            "upright text should read wider than tall"
        );
        assert!(
            ry1 - ry0 > rx1 - rx0,
            "a quarter turn should make the same run taller than wide"
        );
    }

    #[test]
    fn paints_runs_with_the_registered_family() {
        let list = text_list("\u{5e9}\u{5dc}\u{5d5}\u{5dd}", "Noto Sans Hebrew", IDENTITY);
        let mut renderer = Renderer::default();
        renderer
            .register_font(
                "Noto Sans Hebrew",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/NotoSansHebrew-Regular.ttf")
                    .to_vec(),
            )
            .unwrap();
        let fallback = render_list(None, &list, &HashMap::new(), 1.0).unwrap();
        let registered = render_list(Some(&renderer), &list, &HashMap::new(), 1.0).unwrap();
        assert!(
            fallback.bytes != registered.bytes,
            "the registered face must reach the glyph painter"
        );
        ink_bounds(&registered);
    }
}
