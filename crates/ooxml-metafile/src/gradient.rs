//! GDI raster operations without a source bitmap, and `GRADIENTFILL`.

use crate::drawing::{Paint, PathCommand, Rgba};
use crate::player::Player;
use crate::read::{i32_at, u16_at, u32_at};

pub(crate) const ROP_BLACKNESS: u32 = 0x0000_0042;
pub(crate) const ROP_PATCOPY: u32 = 0x00F0_0021;
pub(crate) const ROP_DSTCOPY: u32 = 0x00AA_0029;
pub(crate) const ROP_WHITENESS: u32 = 0x00FF_0062;

/// Replays the BITBLT raster operations that need no source bitmap.
pub(crate) fn bitblt<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    body: usize,
) -> Option<()> {
    if u32_at(bytes, body + 80)? != 0 || u32_at(bytes, body + 88)? != 0 {
        return None;
    }
    let rop = u32_at(bytes, body + 32)?;
    if rop == ROP_DSTCOPY {
        return Some(());
    }
    let fill = match rop {
        ROP_PATCOPY => player.brush_fill(),
        ROP_BLACKNESS => Some(Paint::Solid(Rgba::BLACK)),
        ROP_WHITENESS => Some(Paint::Solid(Rgba::WHITE)),
        _ => return None,
    };
    if player.bracketed {
        return None;
    }
    let (Some(x), Some(y), Some(width), Some(height)) = (
        i32_at(bytes, body + 16),
        i32_at(bytes, body + 20),
        i32_at(bytes, body + 24),
        i32_at(bytes, body + 28),
    ) else {
        return None;
    };
    player.flush_pending();
    if width == 0 || height == 0 {
        return Some(());
    }
    let (x, y) = (f64::from(x), f64::from(y));
    player.append_rect((x, y, x + f64::from(width), y + f64::from(height)));
    player.emit(fill, None);
    Some(())
}

pub(crate) const GRADIENT_RECT_H: u32 = 0;
pub(crate) const GRADIENT_RECT_V: u32 = 1;
pub(crate) const GRADIENT_TRIANGLE: u32 = 2;
pub(crate) const MAX_GRADIENT_VERTICES: usize = 4_096;
pub(crate) const MAX_GRADIENT_BANDS: usize = 64;

#[derive(Clone, Copy, Default)]
struct GradientVertex {
    at: (f64, f64),
    color: [f64; 3],
}

/// Replays GRADIENTFILL as flat-shaded bands across each rectangle or triangle.
pub(crate) fn gradient_fill<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    body: usize,
) -> Option<()> {
    if player.bracketed {
        return None;
    }
    let count = u32_at(bytes, body + 16)? as usize;
    let shapes = u32_at(bytes, body + 20)? as usize;
    let mode = u32_at(bytes, body + 24)?;
    let corners = match mode {
        GRADIENT_TRIANGLE => 3usize,
        GRADIENT_RECT_H | GRADIENT_RECT_V => 2,
        _ => return None,
    };
    if count > MAX_GRADIENT_VERTICES || shapes > MAX_GRADIENT_VERTICES {
        return None;
    }
    let indexes = body + 28 + count * 16;
    if bytes.len() < indexes + shapes * corners * 4 {
        return None;
    }
    let mut vertices = Vec::with_capacity(count);
    for index in 0..count {
        let at = body + 28 + index * 16;
        vertices.push(GradientVertex {
            at: (
                f64::from(i32_at(bytes, at)?),
                f64::from(i32_at(bytes, at + 4)?),
            ),
            color: [
                f64::from(u16_at(bytes, at + 8)? >> 8),
                f64::from(u16_at(bytes, at + 10)? >> 8),
                f64::from(u16_at(bytes, at + 12)? >> 8),
            ],
        });
    }
    player.flush_pending();
    for shape in 0..shapes {
        if player.overflowed {
            return Some(());
        }
        let mut corner = [GradientVertex::default(); 3];
        for (slot, vertex) in corner.iter_mut().enumerate().take(corners) {
            let index = u32_at(bytes, indexes + (shape * corners + slot) * 4)? as usize;
            *vertex = *vertices.get(index)?;
        }
        if corners == 3 {
            shade_triangle(player, corner);
        } else {
            shade_rect(player, corner[0], corner[1], mode == GRADIENT_RECT_V);
        }
    }
    Some(())
}

fn band_count(spread: f64) -> usize {
    (spread.ceil() as usize).clamp(1, MAX_GRADIENT_BANDS)
}

fn mix(from: [f64; 3], to: [f64; 3], at: f64) -> Paint {
    let channel = |index: usize| {
        let value = from[index] + (to[index] - from[index]) * at;
        (value.round() as u32).min(255)
    };
    Paint::Solid(Rgba::from_colorref(
        channel(0) | channel(1) << 8 | channel(2) << 16,
    ))
}

fn shade_rect<const FULL: bool>(
    player: &mut Player<FULL>,
    from: GradientVertex,
    to: GradientVertex,
    vertical: bool,
) {
    let (x0, y0) = from.at;
    let (x1, y1) = to.at;
    if x0 == x1 || y0 == y1 {
        return;
    }
    let spread = (0..3)
        .map(|channel| (to.color[channel] - from.color[channel]).abs())
        .fold(0.0f64, f64::max);
    let bands = band_count(spread);
    for band in 0..bands {
        let start = band as f64 / bands as f64;
        let rect = if vertical {
            (x0, y0 + (y1 - y0) * start, x1, y1)
        } else {
            (x0 + (x1 - x0) * start, y0, x1, y1)
        };
        player.append_rect(rect);
        player.emit(
            Some(mix(from.color, to.color, start + 0.5 / bands as f64)),
            None,
        );
        if player.overflowed {
            return;
        }
    }
}

fn shade_triangle<const FULL: bool>(player: &mut Player<FULL>, corner: [GradientVertex; 3]) {
    let points = [corner[0].at, corner[1].at, corner[2].at];
    let first = (points[1].0 - points[0].0, points[1].1 - points[0].1);
    let second = (points[2].0 - points[0].0, points[2].1 - points[0].1);
    let area = first.0 * second.1 - first.1 * second.0;
    if area == 0.0 {
        return;
    }
    let mut slope = [(0.0, 0.0); 3];
    for (channel, gradient) in slope.iter_mut().enumerate() {
        let along_first = corner[1].color[channel] - corner[0].color[channel];
        let along_second = corner[2].color[channel] - corner[0].color[channel];
        *gradient = (
            (second.1 * along_first - first.1 * along_second) / area,
            (first.0 * along_second - second.0 * along_first) / area,
        );
    }
    let spread = |channel: usize| {
        let values = corner.map(|vertex| vertex.color[channel]);
        values.iter().fold(f64::MIN, |a, b| a.max(*b))
            - values.iter().fold(f64::MAX, |a, b| a.min(*b))
    };
    let dominant = (0..3)
        .max_by(|a, b| spread(*a).total_cmp(&spread(*b)))
        .unwrap_or(0);
    let colour_at = |point: (f64, f64)| {
        let channel = |index: usize| {
            let value = corner[0].color[index]
                + slope[index].0 * (point.0 - points[0].0)
                + slope[index].1 * (point.1 - points[0].1);
            (value.round().max(0.0) as u32).min(255)
        };
        Paint::Solid(Rgba::from_colorref(
            channel(0) | channel(1) << 8 | channel(2) << 16,
        ))
    };
    let length = slope[dominant].0.hypot(slope[dominant].1);
    if !length.is_normal() {
        fill_polygon(player, &points, colour_at(points[0]));
        return;
    }
    let axis = (slope[dominant].0 / length, slope[dominant].1 / length);
    let project = |point: (f64, f64)| point.0 * axis.0 + point.1 * axis.1;
    let projected = points.map(project);
    let low = projected.iter().fold(f64::MAX, |a, b| a.min(*b));
    let high = projected.iter().fold(f64::MIN, |a, b| a.max(*b));
    let bands = band_count(spread(dominant));
    let step = (high - low) / bands as f64;
    for band in 0..bands {
        let start = low + step * band as f64;
        let kept = clip_beyond(&points, axis, start);
        if kept.len() < 3 {
            continue;
        }
        let mean = kept.iter().fold((0.0, 0.0), |sum, point| {
            (
                sum.0 + point.0 / kept.len() as f64,
                sum.1 + point.1 / kept.len() as f64,
            )
        });
        let shift = start + step / 2.0 - project(mean);
        let sample = (mean.0 + axis.0 * shift, mean.1 + axis.1 * shift);
        fill_polygon(player, &kept, colour_at(sample));
        if player.overflowed {
            return;
        }
    }
}

fn fill_polygon<const FULL: bool>(player: &mut Player<FULL>, points: &[(f64, f64)], color: Paint) {
    let Some(first) = points.first().copied() else {
        return;
    };
    player.move_to(first.0, first.1);
    for point in &points[1..] {
        player.line_to(point.0, point.1);
    }
    player.push(PathCommand::Close);
    player.emit(Some(color), None);
}

fn clip_beyond(points: &[(f64, f64)], axis: (f64, f64), limit: f64) -> Vec<(f64, f64)> {
    let project = |point: (f64, f64)| point.0 * axis.0 + point.1 * axis.1;
    let inside = |value: f64| value >= limit;
    let mut kept = Vec::with_capacity(points.len() + 2);
    for (index, point) in points.iter().enumerate() {
        let previous = points[(index + points.len() - 1) % points.len()];
        let (here, there) = (project(*point), project(previous));
        if inside(here) {
            if !inside(there) {
                kept.push(cross(previous, *point, there, here, limit));
            }
            kept.push(*point);
        } else if inside(there) {
            kept.push(cross(previous, *point, there, here, limit));
        }
    }
    kept
}

fn cross(from: (f64, f64), to: (f64, f64), from_at: f64, to_at: f64, limit: f64) -> (f64, f64) {
    let span = to_at - from_at;
    if span == 0.0 {
        return to;
    }
    let ratio = (limit - from_at) / span;
    (
        from.0 + (to.0 - from.0) * ratio,
        from.1 + (to.1 - from.1) * ratio,
    )
}
