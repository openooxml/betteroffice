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
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        let image = image::load_from_memory(bytes).map_err(|error| error.to_string())?;
        let image = image.to_rgba8();
        let (width, height) = image.dimensions();
        let mut rgb = Vec::with_capacity(width as usize * height as usize * 3);
        let mut alpha = Vec::with_capacity(width as usize * height as usize);
        let mut has_alpha = false;
        for pixel in image.pixels() {
            rgb.extend_from_slice(&pixel.0[..3]);
            alpha.push(pixel.0[3]);
            has_alpha |= pixel.0[3] != 255;
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
pub(crate) fn try_decode(bytes: &[u8]) -> Option<DecodedImage> {
    DecodedImage::from_bytes(bytes).ok()
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
/// crop scale/translate and the display rotation. `crop` and `rotation` are the
/// same pieces kept apart so the painter can clip in the frame's own,
/// still-unrotated coordinate system.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct ImageTransform {
    /// Image unit square → page points, crop and rotation composed. The
    /// painter draws the same matrix as `crop` inside `rotation`; it is kept
    /// whole for callers that hit-test or test the resulting placement.
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) matrix: [f32; 6],
    /// Crop scale/translate in unrotated page points: `[sx, 0, 0, sy, e, f]`.
    pub(crate) crop: [f32; 6],
    /// Rotation about the frame centre in page points.
    pub(crate) rotation: [f32; 6],
    /// Frame rectangle `[x, y, w, h]` in page points, before rotation.
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
/// `page_height` is in PDF points; the primitive's `x`, `y`, `w` and `h` are
/// CSS pixels and are converted to points. The visible source rectangle is
/// `[left, 1 - right] × [bottom, 1 - top]` in image unit coordinates (PDF's
/// bottom-left origin), scaled so it exactly covers the frame.
pub(crate) fn image_transform(
    image: &ImagePrimitive,
    page_height: f32,
) -> Result<ImageTransform, PdfError> {
    if !page_height.is_finite() {
        return Err(PdfError::InvalidNumber("page height"));
    }
    let x = points(&image.x, "image x")?;
    let y = points(&image.y, "image y")?;
    let w = points(&image.w, "image width")?;
    let h = points(&image.h, "image height")?;
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

    // Crop first, then rotate about the frame centre (row-vector composition:
    // crop × rotation).
    let matrix = [
        sx * cos,
        sx * sin,
        -sy * sin,
        sy * cos,
        e * cos - f * sin + rotation[4],
        e * sin + f * cos + rotation[5],
    ];
    if !matrix.iter().all(|value| value.is_finite()) {
        return Err(PdfError::InvalidNumber("image transform"));
    }

    Ok(ImageTransform {
        matrix,
        crop,
        rotation,
        clip: [x, frame_y, w, h],
    })
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

/// Paint the image XObject named by `refs`, applying its frame, crop and
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
    // rotation first clips the rotated frame, not its axis-aligned bound.
    if placement.rotation != IDENTITY {
        content.transform(placement.rotation);
    }
    content.rect(clip_x, clip_y, clip_w, clip_h);
    content.clip_nonzero();
    content.end_path();
    // Crop scaling is applied inside the clip; with the rotation already on
    // the CTM this composes to `crop × rotation`, exactly `matrix`.
    content.transform(placement.crop);
    content.x_object(Name(refs.resource.as_bytes()));
    content.restore_state();
    Ok(())
}
