//! Decoded image data, placement math and PDF XObject painting.
//!
//! Display-list images use CSS pixels with a top-left origin; PDF images are
//! unit-square XObjects placed with a transform and a bottom-left origin, so
//! every placement here converts px to pt and flips Y.

use base64::Engine;
use docx_layout::display_list::ImagePrimitive;
use pdf_writer::{Content, Name};
use serde_json::Number;

use crate::PdfError;
use crate::units::points;

/// A decoded raster image: interleaved RGB plus an optional 8-bit alpha mask.
#[derive(Debug)]
pub struct DecodedImage {
    pub width: i32,
    pub height: i32,
    pub rgb: Vec<u8>,
    pub alpha: Option<Vec<u8>>,
}

impl DecodedImage {
    /// Strictly decode an image, reporting why it failed.
    ///
    /// Raster formats go through the `image` crate. SVG and EMF/WMF metafile
    /// pictures go through the shared sandboxed vector decoders, which bound
    /// the document, the render work and the raster before allocating; a
    /// metafile is replayed to SVG first. Both hand back straight RGBA.
    ///
    /// Metafile text needs the document's fonts; this entry point has none, so
    /// it renders metafile text without glyphs. Use
    /// [`Self::from_bytes_with_fonts`] when a `FontStore` is available.
    ///
    /// Kept as the entry point for callers that only have image bytes; the
    /// export path uses [`Self::from_bytes_with_fonts`].
    #[allow(dead_code)]
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        Self::from_bytes_with_fonts(bytes, None)
    }

    /// [`Self::from_bytes`], laying metafile text out with the document's
    /// fonts when `fonts` carries them. Plain SVG, raster and metafiles
    /// without text decode exactly as before.
    pub fn from_bytes_with_fonts(
        bytes: &[u8],
        fonts: Option<&ooxml_text::FontStore>,
    ) -> Result<Self, String> {
        if ooxml_metafile::is_metafile(bytes) {
            return Self::from_metafile(bytes, fonts);
        }
        if ooxml_raster::svg::looks_like_svg(bytes) {
            return Self::from_svg(bytes);
        }
        let image = image::load_from_memory(bytes).map_err(|error| error.to_string())?;
        let image = image.to_rgba8();
        let (width, height) = image.dimensions();
        Self::from_rgba(width, height, image.into_raw())
    }

    /// Rasterize an SVG picture within [`ooxml_raster::svg::MAX_IMAGE_PIXELS`]
    /// of raster and [`ooxml_raster::svg::MAX_SVG_RENDER_WORK`] of drawing
    /// work, which caps the supersampling it can afford.
    fn from_svg(bytes: &[u8]) -> Result<Self, String> {
        let image = ooxml_raster::svg::parse(bytes)
            .map_err(|refusal| format!("SVG picture refused: {refusal:?}"))?;
        let (data, size) = image
            .render()
            .map_err(|refusal| format!("SVG picture could not render: {refusal:?}"))?;
        Self::from_rgba(size.width(), size.height(), data)
    }

    /// Replay an EMF, EMF+ or WMF metafile into SVG and rasterize it with the
    /// same bounded decoder as [`Self::from_svg`].
    ///
    /// Metafile text records name document fonts. When `fonts` is present and
    /// the replay actually contains text, the parser gets
    /// `ooxml_raster::svg::metafile_fontdb(fonts)` so those runs lay out with
    /// real glyphs; otherwise it falls back to the default empty font
    /// database, where geometry, fills and bitmaps still paint and text runs
    /// render without glyphs.
    fn from_metafile(bytes: &[u8], fonts: Option<&ooxml_text::FontStore>) -> Result<Self, String> {
        let mut budget = ooxml_metafile::ReplayBudget::default();
        let markup = ooxml_metafile::to_svg_with_budget(bytes, &mut budget)
            .map_err(|refusal| format!("metafile picture refused: {refusal}"))?;
        let fontdb = if ooxml_raster::svg::has_text_element(markup.markup.as_bytes()) {
            fonts.map(ooxml_raster::svg::metafile_fontdb)
        } else {
            None
        };
        let image = ooxml_raster::svg::parse_metafile(
            markup.markup.as_bytes(),
            ooxml_raster::svg::MAX_IMAGE_PIXELS,
            fontdb,
        )
        .map_err(|refusal| format!("metafile picture could not parse: {refusal:?}"))?;
        let (data, size) = image
            .render()
            .map_err(|refusal| format!("metafile picture could not render: {refusal:?}"))?;
        Self::from_rgba(size.width(), size.height(), data)
    }

    /// Split straight RGBA into interleaved RGB and an optional alpha plane,
    /// the representation the PDF image writer consumes.
    fn from_rgba(width: u32, height: u32, data: Vec<u8>) -> Result<Self, String> {
        let width_usize = usize::try_from(width).map_err(|_| "image width is too large")?;
        let height_usize = usize::try_from(height).map_err(|_| "image height is too large")?;
        let pixels = width_usize
            .checked_mul(height_usize)
            .ok_or("image dimensions are too large")?;
        let rgba_length = pixels
            .checked_mul(4)
            .ok_or("image dimensions are too large")?;
        if data.len() != rgba_length {
            return Err("image pixel buffer has the wrong length".to_owned());
        }
        let rgb_length = pixels
            .checked_mul(3)
            .ok_or("image dimensions are too large")?;
        let mut rgb = Vec::with_capacity(rgb_length);
        let mut alpha = Vec::with_capacity(pixels);
        let mut has_alpha = false;
        for pixel in data.chunks_exact(4) {
            rgb.extend_from_slice(&pixel[..3]);
            alpha.push(pixel[3]);
            has_alpha |= pixel[3] != 255;
        }
        Ok(Self {
            width: i32::try_from(width).map_err(|_| "image width is too large")?,
            height: i32::try_from(height).map_err(|_| "image height is too large")?,
            rgb,
            alpha: has_alpha.then_some(alpha),
        })
    }
}

/// Decode an image, treating a truncated or corrupt source as a non-fatal
/// missing picture: a bad logo cannot fail the whole document.
#[allow(dead_code)]
pub(crate) fn try_decode(bytes: &[u8]) -> Option<DecodedImage> {
    DecodedImage::from_bytes(bytes).ok()
}

/// [`try_decode`] with the document's fonts for metafile text.
pub(crate) fn try_decode_with_fonts(
    bytes: &[u8],
    fonts: Option<&ooxml_text::FontStore>,
) -> Option<DecodedImage> {
    DecodedImage::from_bytes_with_fonts(bytes, fonts).ok()
}

/// Decode a `data:` URL into its payload.
pub fn data_url_bytes(value: &str) -> Option<Vec<u8>> {
    let (metadata, payload) = value.strip_prefix("data:")?.split_once(',')?;
    if metadata.contains(";base64") {
        base64::engine::general_purpose::STANDARD
            .decode(payload)
            .ok()
    } else {
        Some(payload.as_bytes().to_vec())
    }
}

const IDENTITY: [f32; 6] = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];

/// Where and how one image primitive paints onto a page.
///
/// `matrix` is the complete image-unit-square-to-page transform, including the
/// crop scale/translate, the mirror about the frame centre and the display
/// rotation. `crop`, `flip` and `rotation` are the same pieces kept apart so
/// the painter can clip in the frame's own, still-unrotated coordinate system.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct ImageTransform {
    /// Image unit square → page points, crop, flip and rotation composed. The
    /// painter draws the same matrix as `crop` inside `flip` inside
    /// `rotation`; it is kept whole for callers that hit-test or test the
    /// resulting placement.
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) matrix: [f32; 6],
    /// Crop scale/translate in unrotated page points: `[sx, 0, 0, sy, e, f]`.
    pub(crate) crop: [f32; 6],
    /// Mirror about the frame centre in unrotated page points, or
    /// [`IDENTITY`] when the picture is not flipped: `[±1, 0, 0, ±1, tx, ty]`.
    pub(crate) flip: [f32; 6],
    /// Rotation about the frame centre in page points.
    pub(crate) rotation: [f32; 6],
    /// Frame rectangle `[x, y, w, h]` in page points, before flip or rotation.
    pub(crate) clip: [f32; 4],
}

impl ImageTransform {
    /// The crop made nothing visible (crop fractions cover the whole source).
    fn is_empty(&self) -> bool {
        self.crop[0] == 0.0 || self.crop[3] == 0.0
    }
}

/// Compute the placement of an image primitive on a page.
///
/// `page_height` is in PDF points; the primitive's frame fields are CSS pixels
/// and are converted to points. A `contentFrame` overrides the primitive's own
/// `x`, `y`, `w` and `h` field by field, exactly as `docx-raster`'s
/// `image_frame` does; the primitive's fields remain the fallback. The visible
/// source rectangle is `[left, 1 - right] × [bottom, 1 - top]` in image unit
/// coordinates (PDF's bottom-left origin), scaled so it exactly covers the
/// frame. A `flipH` or `flipV` attribute mirrors that placement about the
/// frame centre before the rotation, matching the native raster, and leaves
/// the frame itself in place.
pub(crate) fn image_transform(
    image: &ImagePrimitive,
    page_height: f32,
) -> Result<ImageTransform, PdfError> {
    if !page_height.is_finite() {
        return Err(PdfError::InvalidNumber("page height"));
    }
    let content_frame = image.attrs.content_frame.as_deref();
    let frame_x = content_frame
        .and_then(|frame| frame.x.as_ref())
        .unwrap_or(&image.x);
    let frame_y = content_frame
        .and_then(|frame| frame.y.as_ref())
        .unwrap_or(&image.y);
    let frame_w = content_frame
        .and_then(|frame| frame.w.as_ref())
        .unwrap_or(&image.w);
    let frame_h = content_frame
        .and_then(|frame| frame.h.as_ref())
        .unwrap_or(&image.h);
    let x = points(frame_x, "image x")?;
    let y = points(frame_y, "image y")?;
    let w = points(frame_w, "image width")?;
    let h = points(frame_h, "image height")?;
    let frame_y = page_height - y - h;

    let (left, bottom, visible_w, visible_h) = crop_bounds(image)?;
    // A crop covering the whole image leaves nothing to paint; a zero scale is
    // a degenerate matrix the painter skips rather than a divide by zero.
    let sx = if visible_w > 0.0 { w / visible_w } else { 0.0 };
    let sy = if visible_h > 0.0 { h / visible_h } else { 0.0 };
    let e = x - left * sx;
    let f = frame_y - bottom * sy;
    let crop = [sx, 0.0, 0.0, sy, e, f];

    let degrees = match image.rotation_deg.as_ref() {
        Some(number) => number
            .as_f64()
            .ok_or(PdfError::InvalidNumber("image rotation"))? as f32,
        None => 0.0,
    };
    if !degrees.is_finite() {
        return Err(PdfError::InvalidNumber("image rotation"));
    }
    // Display rotation is clockwise on a Y-down page; PDF's Y-up space turns
    // the same visual angle counterclockwise, so negate it.
    let (sin, cos) = (-degrees).to_radians().sin_cos();
    let cx = x + w / 2.0;
    let cy = frame_y + h / 2.0;
    let rotation = [
        cos,
        sin,
        -sin,
        cos,
        cx - cos * cx + sin * cy,
        cy - sin * cx - cos * cy,
    ];

    // The native raster mirrors the frame about its centre before rotating it
    // (`docx-raster`'s `image_transform` composes `rotation.pre_concat(flip)`),
    // so the frame stays where it was and its content is mirrored.
    let (flip_x, flip_y) = (
        if image.attrs.image_flip_h == Some(true) {
            -1.0
        } else {
            1.0
        },
        if image.attrs.image_flip_v == Some(true) {
            -1.0
        } else {
            1.0
        },
    );
    let flip = [flip_x, 0.0, 0.0, flip_y, cx - flip_x * cx, cy - flip_y * cy];

    // Crop first, then mirror, then rotate about the frame centre (row-vector
    // composition: crop × flip × rotation).
    let matrix = compose(compose(crop, flip), rotation);
    if !matrix.iter().all(|value| value.is_finite()) {
        return Err(PdfError::InvalidNumber("image transform"));
    }

    Ok(ImageTransform {
        matrix,
        crop,
        flip,
        rotation,
        clip: [x, frame_y, w, h],
    })
}

/// Compose two row-vector PDF matrices: the result applies `inner` first and
/// `outer` second, so `compose(crop, rotation)` is crop then rotate.
fn compose(outer: [f32; 6], inner: [f32; 6]) -> [f32; 6] {
    let [a, b, c, d, e, f] = outer;
    let [g, h, i, j, k, l] = inner;
    [
        a * g + b * i,
        a * h + b * j,
        c * g + d * i,
        c * h + d * j,
        e * g + f * i + k,
        e * h + f * j + l,
    ]
}

/// Visible crop fractions as `(left, bottom, width, height)`, clamped to the
/// source. Numbers that do not fit an `f64` are rejected.
fn crop_bounds(image: &ImagePrimitive) -> Result<(f32, f32, f32, f32), PdfError> {
    let Some(crop) = image.crop.as_ref() else {
        return Ok((0.0, 0.0, 1.0, 1.0));
    };
    let fraction = |value: &Number, what: &'static str| -> Result<f32, PdfError> {
        let value = value.as_f64().ok_or(PdfError::InvalidNumber(what))?;
        if !value.is_finite() {
            return Err(PdfError::InvalidNumber(what));
        }
        Ok(value.clamp(0.0, 1.0) as f32)
    };
    let left = fraction(&crop.left, "image crop left")?;
    let right = fraction(&crop.right, "image crop right")?;
    let top = fraction(&crop.top, "image crop top")?;
    let bottom = fraction(&crop.bottom, "image crop bottom")?;
    Ok((
        left,
        bottom,
        (1.0 - left - right).max(0.0),
        (1.0 - top - bottom).max(0.0),
    ))
}

/// The alpha the integration owner should apply through an ExtGState, or
/// `None` when the image is fully opaque. This module deliberately does not
/// write graphics-state resources.
pub(crate) fn opacity_alpha(image: &ImagePrimitive) -> Option<f32> {
    let value = image.opacity.as_ref()?.as_f64()? as f32;
    if !value.is_finite() {
        return None;
    }
    let value = value.clamp(0.0, 1.0);
    (value < 1.0).then_some(value)
}

/// Paint the image XObject named by `refs`, applying its frame, crop, flip and
/// rotation.
///
/// `page_height` is in PDF points; opacity is ignored here by design — see
/// [`opacity_alpha`].
pub(crate) fn paint_image(
    content: &mut Content,
    page_height: f32,
    image: &ImagePrimitive,
    refs: &crate::ImageRefs,
) -> Result<(), PdfError> {
    let placement = image_transform(image, page_height)?;
    let [clip_x, clip_y, clip_w, clip_h] = placement.clip;
    if clip_w <= 0.0 || clip_h <= 0.0 || placement.is_empty() {
        return Ok(());
    }

    content.save_state();
    // The clip is defined in the frame's own coordinates. Applying the
    // rotation and mirror first clips the rotated frame, not its axis-aligned
    // bound; the mirror maps the frame rectangle onto itself.
    if placement.rotation != IDENTITY {
        content.transform(placement.rotation);
    }
    if placement.flip != IDENTITY {
        content.transform(placement.flip);
    }
    content.rect(clip_x, clip_y, clip_w, clip_h);
    content.clip_nonzero();
    content.end_path();
    // Crop scaling is applied inside the clip; with the rotation and mirror
    // already on the CTM this composes to `crop × flip × rotation`, exactly
    // `matrix`.
    content.transform(placement.crop);
    content.x_object(Name(refs.resource.as_bytes()));
    content.restore_state();
    Ok(())
}
