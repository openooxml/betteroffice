//! Canvas-style shadows: the alpha a primitive paints, offset, blurred and filled
//! with one colour, drawn under the primitive.

use tiny_skia::{Color, Mask, Pixmap, PixmapPaint, PremultipliedColorU8, Rect, Transform};

use crate::{PageBudget, mask_bytes};

/// The widest blur and furthest offset a shadow is drawn with, in pixels.
const MAX_SHADOW_EXTENT: f32 = 256.0;

/// A shadow as the web canvas draws one: `shadowColor`, `shadowBlur` and
/// `shadowOffsetX/Y`, in device pixels and, as on a canvas, unaffected by the
/// current transform.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct Shadow {
    pub color: Color,
    pub blur: f32,
    pub dx: f32,
    pub dy: f32,
}

impl Shadow {
    /// A shadow with its blur and offset bounded, so a hostile effect cannot
    /// ask for an unbounded surface or blur.
    pub(crate) fn new(color: Color, blur: f32, dx: f32, dy: f32) -> Self {
        let bound = |value: f32, low: f32| {
            if value.is_finite() {
                value.clamp(low, MAX_SHADOW_EXTENT)
            } else {
                0.0
            }
        };
        Self {
            color,
            blur: bound(blur, 0.0),
            dx: bound(dx, -MAX_SHADOW_EXTENT),
            dy: bound(dy, -MAX_SHADOW_EXTENT),
        }
    }
}

/// What a `textShadow` value draws. The web backend draws only `"shadow"`: 30%
/// black, blurred by an eighth of the font size and offset by a sixteenth.
pub(crate) enum TextShadow {
    None,
    Drawn(Shadow),
    /// A value neither backend draws yet, such as `emboss` or `imprint`.
    Undrawn,
}

pub(crate) fn text_shadow(value: Option<&str>, size: f32) -> TextShadow {
    match value {
        None => TextShadow::None,
        Some("shadow") => TextShadow::Drawn(Shadow::new(
            Color::from_rgba8(0, 0, 0, 77),
            (size / 8.0).max(1.0),
            (size / 16.0).max(0.5),
            (size / 16.0).max(0.5),
        )),
        Some(_) => TextShadow::Undrawn,
    }
}

/// A scratch surface the shadowed content is painted into, and where it lands
/// on the page.
pub(crate) struct ShadowSurface {
    pub pixmap: Pixmap,
    /// Maps page device space onto the surface, shadow offset included.
    pub offset: Transform,
    left: i32,
    top: i32,
}

/// A surface covering `bounds` (device pixels) moved by the shadow offset and
/// grown by the blur's reach. It is clipped one reach beyond the page, so
/// coverage just off the page still blurs inward. `None` when nothing of the
/// shadow can land on the page.
pub(crate) fn shadow_surface(
    page: (u32, u32),
    bounds: Rect,
    shadow: &Shadow,
    budget: &mut PageBudget,
    retained_mask_bytes: u64,
) -> Result<Option<ShadowSurface>, String> {
    let reach = (shadow.blur / 2.0 * 3.0).ceil();
    let left = (bounds.left() + shadow.dx - reach).floor().max(-reach);
    let top = (bounds.top() + shadow.dy - reach).floor().max(-reach);
    let right = (bounds.right() + shadow.dx + reach)
        .ceil()
        .min(page.0 as f32 + reach);
    let bottom = (bounds.bottom() + shadow.dy + reach)
        .ceil()
        .min(page.1 as f32 + reach);
    if !(right > left && bottom > top) || right <= 0.0 || bottom <= 0.0 {
        return Ok(None);
    }
    let (width, height) = ((right - left) as u32, (bottom - top) as u32);
    let surface = mask_bytes(width, height).saturating_mul(5);
    let line = u64::from(width.max(height)).saturating_mul(4);
    budget.charge_masks(
        retained_mask_bytes
            .saturating_add(surface)
            .saturating_add(line),
    )?;
    let pixmap =
        Pixmap::new(width, height).ok_or_else(|| "invalid shadow surface size".to_string())?;
    Ok(Some(ShadowSurface {
        pixmap,
        offset: Transform::from_translate(shadow.dx - left, shadow.dy - top),
        left: left as i32,
        top: top as i32,
    }))
}

/// Blurs the surface's coverage, fills it with the shadow colour and draws it
/// onto `target` under `clip`. The coverage already carries the content's own
/// opacity, as a canvas shadow's does.
pub(crate) fn composite_shadow(
    target: &mut Pixmap,
    clip: Option<&Mask>,
    mut surface: ShadowSurface,
    shadow: &Shadow,
) {
    let width = surface.pixmap.width() as usize;
    let height = surface.pixmap.height() as usize;
    let mut alpha: Vec<u8> = surface
        .pixmap
        .pixels()
        .iter()
        .map(|pixel| pixel.alpha())
        .collect();
    blur(&mut alpha, width, height, shadow.blur / 2.0);
    let strength = shadow.color.alpha();
    for (pixel, coverage) in surface.pixmap.pixels_mut().iter_mut().zip(&alpha) {
        let a = f32::from(*coverage) / 255.0 * strength;
        let channel = |value: f32| (value * a * 255.0).round().clamp(0.0, 255.0) as u8;
        let alpha = (a * 255.0).round().clamp(0.0, 255.0) as u8;
        *pixel = PremultipliedColorU8::from_rgba(
            channel(shadow.color.red()).min(alpha),
            channel(shadow.color.green()).min(alpha),
            channel(shadow.color.blue()).min(alpha),
            alpha,
        )
        .unwrap_or(PremultipliedColorU8::TRANSPARENT);
    }
    target.draw_pixmap(
        surface.left,
        surface.top,
        surface.pixmap.as_ref(),
        &PixmapPaint::default(),
        Transform::identity(),
        clip,
    );
}

/// Approximates a Gaussian blur of standard deviation `sigma` with three box
/// blurs in each direction.
fn blur(data: &mut [u8], width: usize, height: usize, sigma: f32) {
    if sigma < 0.5 || width == 0 || height == 0 {
        return;
    }
    for radius in box_radii(sigma) {
        box_pass(data, height, width, width, 1, radius);
        box_pass(data, width, height, 1, width, radius);
    }
}

/// The radii of three box blurs whose sum approximates a Gaussian of `sigma`.
fn box_radii(sigma: f32) -> [usize; 3] {
    let sigma = sigma.min(MAX_SHADOW_EXTENT / 2.0);
    let passes = 3.0;
    let ideal = (12.0 * sigma * sigma / passes + 1.0).sqrt();
    let mut lower = ideal.floor() as i64;
    if lower % 2 == 0 {
        lower -= 1;
    }
    let lower = lower.max(1);
    let upper = lower + 2;
    let lower_f = lower as f32;
    let small = ((12.0 * sigma * sigma
        - passes * lower_f * lower_f
        - 4.0 * passes * lower_f
        - 3.0 * passes)
        / (-4.0 * lower_f - 4.0))
        .round() as i64;
    [0, 1, 2].map(|pass| (((if pass < small { lower } else { upper }) - 1) / 2).max(0) as usize)
}

/// One box blur of `radius` along `lines` lines of `len` samples, zero beyond
/// the edges.
fn box_pass(
    data: &mut [u8],
    lines: usize,
    len: usize,
    line_stride: usize,
    step: usize,
    radius: usize,
) {
    if radius == 0 || len == 0 {
        return;
    }
    let window = (2 * radius + 1) as u32;
    let mut line = vec![0u32; len];
    for index in 0..lines {
        let start = index * line_stride;
        for (at, sample) in line.iter_mut().enumerate() {
            *sample = u32::from(data[start + at * step]);
        }
        let mut sum: u32 = line.iter().take(radius + 1).sum();
        for at in 0..len {
            data[start + at * step] = ((sum + window / 2) / window) as u8;
            if let Some(next) = line.get(at + radius + 1) {
                sum += next;
            }
            if at >= radius {
                sum -= line[at - radius];
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_blur_spreads_coverage_and_keeps_its_total() {
        let (width, height) = (31, 31);
        let mut data = vec![0u8; width * height];
        data[15 * width + 15] = 255;
        for row in 13..18 {
            for col in 13..18 {
                data[row * width + col] = 255;
            }
        }
        let before: u32 = data.iter().map(|&value| u32::from(value)).sum();
        blur(&mut data, width, height, 2.0);
        let after: u32 = data.iter().map(|&value| u32::from(value)).sum();
        assert!(data[15 * width + 15] < 255 && data[15 * width + 11] > 0);
        assert!((f64::from(after) - f64::from(before)).abs() / f64::from(before) < 0.05);
    }

    #[test]
    fn the_text_shadow_is_the_web_recipe() {
        let TextShadow::Drawn(shadow) = text_shadow(Some("shadow"), 16.0) else {
            panic!("a drawn shadow");
        };
        assert_eq!((shadow.blur, shadow.dx, shadow.dy), (2.0, 1.0, 1.0));
        assert_eq!(shadow.color, Color::from_rgba8(0, 0, 0, 77));
        assert!(matches!(
            text_shadow(Some("emboss"), 16.0),
            TextShadow::Undrawn
        ));
        assert!(matches!(text_shadow(None, 16.0), TextShadow::None));
    }

    #[test]
    fn a_shadow_is_bounded_whatever_it_asks_for() {
        let shadow = Shadow::new(Color::BLACK, 1e20, f32::NAN, -1e9);
        assert_eq!((shadow.blur, shadow.dx, shadow.dy), (256.0, 0.0, -256.0));
        assert!(box_radii(1e20).iter().all(|radius| *radius < 1024));
    }
}
