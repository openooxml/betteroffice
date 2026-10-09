//! EMF (`[MS-EMF]`) records.

use std::sync::Arc;

use crate::drawing::PathCommand;
use crate::player::{
    Brush, Combine, GdiObject, IDENTITY, LogFont, Pen, Player, axis_aligned_rect, concat,
    stock_object,
};
use crate::read::{finite_at, i16_at, i32_at, u16_at, u32_at};

pub(crate) const EMF_SIGNATURE: u32 = 0x464D_4520;
pub(crate) const EMF_EOF: u32 = 14;
pub(crate) const EMF_COMMENT: u32 = 70;
pub(crate) const EMF_PLUS: u32 = 0x2B46_4D45;

/// The only clip combination mode the strict replay represents: replace.
pub(crate) const RGN_COPY: u32 = 5;

pub(crate) fn is_emf(bytes: &[u8]) -> bool {
    u32_at(bytes, 0) == Some(1) && u32_at(bytes, 40) == Some(EMF_SIGNATURE)
}

/// The frame in device units, from the header's `rclFrame` or, when that is
/// empty, its `rclBounds`.
pub(crate) fn emf_frame(bytes: &[u8]) -> Option<(f64, f64, f64, f64)> {
    let bounds = (
        i32_at(bytes, 8)? as f64,
        i32_at(bytes, 12)? as f64,
        i32_at(bytes, 16)? as f64 + 1.0,
        i32_at(bytes, 20)? as f64 + 1.0,
    );
    let frame = (
        i32_at(bytes, 24)? as f64,
        i32_at(bytes, 28)? as f64,
        i32_at(bytes, 32)? as f64,
        i32_at(bytes, 36)? as f64,
    );
    let device = (i32_at(bytes, 72)? as f64, i32_at(bytes, 76)? as f64);
    let millimetres = (i32_at(bytes, 80)? as f64, i32_at(bytes, 84)? as f64);
    let rect = if millimetres.0 > 0.0 && millimetres.1 > 0.0 {
        let per_unit = (
            device.0 / (millimetres.0 * 100.0),
            device.1 / (millimetres.1 * 100.0),
        );
        (
            frame.0 * per_unit.0,
            frame.1 * per_unit.1,
            frame.2 * per_unit.0,
            frame.3 * per_unit.1,
        )
    } else {
        bounds
    };
    let (width, height) = (rect.2 - rect.0, rect.3 - rect.1);
    if width.abs() < f64::EPSILON || height.abs() < f64::EPSILON {
        let (width, height) = (bounds.2 - bounds.0, bounds.3 - bounds.1);
        if width.abs() < f64::EPSILON || height.abs() < f64::EPSILON {
            return None;
        }
        return Some((bounds.0, bounds.1, width, height));
    }
    Some((rect.0, rect.1, width, height))
}

/// Device units per millimetre, from the reference device the header names.
pub(crate) fn device_per_mm(bytes: &[u8]) -> (f64, f64) {
    let read = |at| i32_at(bytes, at).map(f64::from).unwrap_or(0.0);
    let (device, millimetres) = ((read(72), read(76)), (read(80), read(84)));
    let axis = |device: f64, millimetres: f64| {
        if device > 0.0 && millimetres > 0.0 {
            device / millimetres
        } else {
            96.0 / 25.4
        }
    };
    (axis(device.0, millimetres.0), axis(device.1, millimetres.1))
}

pub(crate) fn read_points(
    bytes: &[u8],
    offset: usize,
    count: usize,
    small: bool,
    limit: usize,
) -> Option<Vec<(f64, f64)>> {
    if count > limit {
        return None;
    }
    let stride = if small { 4 } else { 8 };
    crate::read::span(bytes, offset, count, stride)?;
    let mut points = Vec::with_capacity(count);
    for index in 0..count {
        let at = offset + index * stride;
        let point = if small {
            (i16_at(bytes, at)? as f64, i16_at(bytes, at + 2)? as f64)
        } else {
            (i32_at(bytes, at)? as f64, i32_at(bytes, at + 4)? as f64)
        };
        points.push(point);
    }
    Some(points)
}

pub(crate) fn brush_from_style(style: u32, color: u32) -> Option<Brush> {
    if !matches!(style, 0 | 1) {
        return None;
    }
    Some(Brush::solid(color, style != 1))
}

pub(crate) fn pen_from_style(style: u32, width: f64, color: u32) -> Option<Pen> {
    if !matches!(style & 0x0f, 0 | 5) {
        return None;
    }
    Some(Pen::solid(color, width, style & 0x0f != 5))
}

/// Any `PS_*` pen; `geometric` is `None` for a `LOGPEN`, whose type follows its width.
pub(crate) fn full_pen(style: u32, width: f64, color: u32, geometric: Option<bool>) -> Pen {
    let kind = style & 0x0f;
    Pen {
        color,
        width,
        visible: kind != 5,
        style,
        cosmetic: match geometric {
            Some(geometric) => !geometric,
            None => width <= 1.0,
        },
        user_dash: None,
    }
}

pub(crate) fn full_brush(style: u32, color: u32, hatch: u32) -> Brush {
    match style {
        1 => Brush::solid(color, false),
        2 => Brush {
            hatch: Some(hatch),
            ..Brush::solid(color, true)
        },
        _ => Brush::solid(color, true),
    }
}

fn default_font(index: u32) -> LogFont {
    let fixed = matches!(index, 10 | 11 | 16);
    LogFont {
        height: 16,
        escapement: 0,
        weight: if index == 13 || index == 16 { 700 } else { 400 },
        italic: false,
        underline: false,
        strike: false,
        charset: 0,
        face: if fixed { "Courier New" } else { "Arial" }.to_owned(),
    }
}

pub(crate) fn logfont(bytes: &[u8], at: usize, wide: bool) -> Option<LogFont> {
    let face = if wide {
        let mut units = Vec::with_capacity(32);
        for index in 0..32 {
            let unit = u16_at(bytes, at + 28 + index * 2)?;
            if unit == 0 {
                break;
            }
            units.push(unit);
        }
        String::from_utf16_lossy(&units)
    } else {
        let raw = bytes.get(at + 18..(at + 50).min(bytes.len()))?;
        let end = raw.iter().position(|byte| *byte == 0).unwrap_or(raw.len());
        raw[..end].iter().map(|byte| char::from(*byte)).collect()
    };
    let (height, escapement, weight, flags) = if wide {
        (
            i32_at(bytes, at)?,
            i32_at(bytes, at + 8)?,
            i32_at(bytes, at + 16)?,
            at + 20,
        )
    } else {
        (
            i32::from(i16_at(bytes, at)?),
            i32::from(i16_at(bytes, at + 4)?),
            i32::from(i16_at(bytes, at + 8)?),
            at + 10,
        )
    };
    Some(LogFont {
        height,
        escapement,
        weight,
        italic: bytes.get(flags).copied().unwrap_or(0) != 0,
        underline: bytes.get(flags + 1).copied().unwrap_or(0) != 0,
        strike: bytes.get(flags + 2).copied().unwrap_or(0) != 0,
        charset: bytes.get(flags + 3).copied().unwrap_or(0),
        face,
    })
}

pub(crate) fn emf_record<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    kind: u32,
    body: usize,
) -> Option<()> {
    let limit = player.limits.points_per_record;
    let poly_points = |small: bool| -> Option<Vec<(f64, f64)>> {
        let count = u32_at(bytes, body + 16)? as usize;
        read_points(bytes, body + 20, count, small, limit)
    };
    let box_rect = || -> Option<(f64, f64, f64, f64)> {
        Some((
            i32_at(bytes, body)? as f64,
            i32_at(bytes, body + 4)? as f64,
            i32_at(bytes, body + 8)? as f64,
            i32_at(bytes, body + 12)? as f64,
        ))
    };
    match kind {
        9..=12 => {
            let (x, y) = i32_at(bytes, body).zip(i32_at(bytes, body + 4))?;
            let value = (f64::from(x), f64::from(y));
            if matches!(kind, 9 | 11) && (x == 0 || y == 0) {
                return FULL.then_some(());
            }
            match kind {
                9 => {
                    player.dc.window_ext = value;
                    player.dc.window_ext_set = true;
                }
                10 => player.dc.window_org = value,
                11 => {
                    player.dc.viewport_ext = value;
                    player.dc.viewport_ext_set = true;
                }
                _ => player.dc.viewport_org = value,
            }
            player.dc.scaled = player.dc.map_mode == 8
                && player.dc.window_ext_set
                && player.dc.viewport_ext_set
                && player.dc.window_ext.0 != 0.0
                && player.dc.window_ext.1 != 0.0;
        }
        17 => {
            let mode = i32_at(bytes, body)?;
            let accepted = if FULL {
                (1..=8).contains(&mode)
            } else {
                matches!(mode, 1 | 8)
            };
            if !accepted {
                return player.refuse(format!("mapping mode {mode} is not defined"));
            }
            player.dc.map_mode = mode;
            player.dc.scaled = mode == 8 && player.dc.window_ext_set && player.dc.viewport_ext_set;
        }
        19 => {
            player.dc.even_odd = match u32_at(bytes, body)? {
                1 => true,
                2 => false,
                _ => return None,
            };
        }
        57 => {
            player.dc.clockwise = match u32_at(bytes, body)? {
                1 => false,
                2 => true,
                _ => return None,
            };
        }
        33 => player.save()?,
        34 => player.restore(i32_at(bytes, body)?)?,
        35 | 36 => {
            let mut matrix = IDENTITY;
            for (index, slot) in matrix.iter_mut().enumerate() {
                *slot = finite_at(bytes, body + index * 4)?;
            }
            let mode = if kind == 36 {
                u32_at(bytes, body + 24)?
            } else {
                4
            };
            player.dc.xform = match mode {
                1 => IDENTITY,
                2 => concat(matrix, player.dc.xform),
                3 => concat(player.dc.xform, matrix),
                4 => matrix,
                _ => return None,
            };
        }
        37 => {
            let handle = u32_at(bytes, body)?;
            player.flush_pending();
            let object = if handle & 0x8000_0000 != 0 {
                let index = handle & 0x7fff_ffff;
                match stock_object(index) {
                    Some(object) => Some(object),
                    None if FULL && (10..=17).contains(&index) && index != 15 => {
                        Some(GdiObject::Font(Arc::new(default_font(index))))
                    }
                    None => None,
                }
            } else {
                player.objects.get(handle as usize).cloned().flatten()
            };
            if let Some(object) = object {
                player.select(object);
            }
        }
        38 => {
            let (Some(handle), Some(style), Some(width), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 4),
                i32_at(bytes, body + 8),
                u32_at(bytes, body + 16),
            ) else {
                return None;
            };
            let pen = if FULL {
                full_pen(style, f64::from(width), color, None)
            } else {
                pen_from_style(style, f64::from(width), color)?
            };
            player.store(handle as usize, GdiObject::Pen(pen));
        }
        95 => {
            let (Some(handle), Some(style), Some(width), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 20),
                u32_at(bytes, body + 24),
                u32_at(bytes, body + 32),
            ) else {
                return None;
            };
            let pen = if FULL {
                let mut pen = full_pen(
                    style,
                    f64::from(width),
                    color,
                    Some(style & 0x0001_0000 != 0),
                );
                if style & 0x0f == 7 {
                    let count = (u32_at(bytes, body + 40)? as usize).min(16);
                    let mut entries = Vec::with_capacity(count);
                    for index in 0..count {
                        entries.push(f64::from(u32_at(bytes, body + 44 + index * 4)?));
                    }
                    if entries.iter().any(|entry| *entry > 0.0) {
                        pen.user_dash = Some(entries.into());
                    }
                }
                if u32_at(bytes, body + 28)? == 1 {
                    pen.visible = false;
                }
                pen
            } else {
                pen_from_style(style, f64::from(width), color)?
            };
            player.store(handle as usize, GdiObject::Pen(pen));
        }
        39 => {
            let (Some(handle), Some(style), Some(color)) = (
                u32_at(bytes, body),
                u32_at(bytes, body + 4),
                u32_at(bytes, body + 8),
            ) else {
                return None;
            };
            let brush = if FULL {
                full_brush(style, color, u32_at(bytes, body + 12)?)
            } else {
                brush_from_style(style, color)?
            };
            player.store(handle as usize, GdiObject::Brush(brush));
        }
        40 => {
            let handle = u32_at(bytes, body);
            if FULL && handle.is_none() {
                return None;
            }
            if let Some(handle) = handle
                && let Some(slot) = player.objects.get_mut(handle as usize)
            {
                *slot = None;
            }
        }
        59 => {
            player.flush_pending();
            player.path.clear();
            player.selected_path.clear();
            player.bracketed = true;
            player.widened = false;
        }
        60 => {
            player.bracketed = false;
            player.selected_path = std::mem::take(&mut player.path);
        }
        61 => player.close_figure(),
        62 => {
            player.flush_pending();
            player.path = std::mem::take(&mut player.selected_path);
            if FULL && std::mem::take(&mut player.widened) {
                let stroke = player.widened_stroke();
                player.emit(None, stroke);
            } else {
                let fill = player.brush_fill();
                player.emit(fill, None);
            }
        }
        63 => {
            player.flush_pending();
            player.path = std::mem::take(&mut player.selected_path);
            let (fill, stroke) = (player.brush_fill(), player.pen_stroke());
            player.emit(fill, stroke);
        }
        64 => {
            player.flush_pending();
            player.path = std::mem::take(&mut player.selected_path);
            let stroke = player.pen_stroke();
            player.emit(None, stroke);
        }
        67 => {
            let mode = u32_at(bytes, body)?;
            if !FULL && mode != RGN_COPY {
                return None;
            }
            player.flush_pending();
            let path = std::mem::take(&mut player.selected_path);
            if FULL {
                let mode = Combine::from_gdi(mode)?;
                let even_odd = player.dc.even_odd;
                player.combine_clip(path, even_odd, mode)?;
            } else {
                let rect = axis_aligned_rect(&path)?;
                player.combine_clip(crate::player::rect_path(rect), false, Combine::Copy)?;
            }
        }
        75 => {
            let (size, mode) = (u32_at(bytes, body)?, u32_at(bytes, body + 4)?);
            if !FULL && (size != 0 || mode != RGN_COPY) {
                return None;
            }
            player.flush_pending();
            if FULL {
                let mode = Combine::from_gdi(mode)?;
                if size == 0 {
                    if mode != Combine::Copy {
                        return player.refuse("a clip region combines with no region");
                    }
                    player.reset_clip();
                } else {
                    let path = region_path(player, bytes, body + 8, size as usize)?;
                    player.combine_clip(path, false, mode)?;
                }
            } else {
                player.reset_clip();
            }
        }
        68 => {
            player.path.clear();
            player.selected_path.clear();
            player.bracketed = false;
            player.pending_stroke = false;
        }
        27 => {
            let (x, y) = i32_at(bytes, body).zip(i32_at(bytes, body + 4))?;
            if !player.bracketed {
                player.flush_pending();
            }
            player.move_to(f64::from(x), f64::from(y));
        }
        54 => {
            let (x, y) = i32_at(bytes, body).zip(i32_at(bytes, body + 4))?;
            player.line_to(f64::from(x), f64::from(y));
            player.pending_stroke = !player.bracketed;
        }
        5 | 6 | 88 | 89 => {
            let small = kind >= 88;
            let points = poly_points(small)?;
            if matches!(kind, 5 | 88) {
                for chunk in points.as_chunks::<3>().0 {
                    player.cubic_to([chunk[0], chunk[1], chunk[2]]);
                }
            } else {
                for point in points {
                    player.line_to(point.0, point.1);
                }
            }
            player.pending_stroke = !player.bracketed;
        }
        2 | 3 | 4 | 85 | 86 | 87 => {
            let small = kind >= 85;
            let points = poly_points(small)?;
            let first = points.first().copied()?;
            if !player.bracketed {
                player.flush_pending();
            }
            player.move_to(first.0, first.1);
            if matches!(kind, 2 | 85) {
                for chunk in points[1..].as_chunks::<3>().0 {
                    player.cubic_to([chunk[0], chunk[1], chunk[2]]);
                }
            } else {
                for point in &points[1..] {
                    player.line_to(point.0, point.1);
                }
            }
            let closed = matches!(kind, 3 | 86);
            if closed {
                player.push(PathCommand::Close);
            }
            player.paint_figure(closed);
        }
        7 | 8 | 90 | 91 => {
            let small = kind >= 90;
            let (Some(polygons), Some(total)) =
                (u32_at(bytes, body + 16), u32_at(bytes, body + 20))
            else {
                return None;
            };
            let (polygons, total) = (polygons as usize, total as usize);
            if polygons > limit || total > limit {
                return None;
            }
            crate::read::span(bytes, body + 24, polygons, 4)?;
            let mut counts = Vec::with_capacity(polygons);
            for index in 0..polygons {
                let count = u32_at(bytes, body + 24 + index * 4)?;
                counts.push(count as usize);
            }
            if counts
                .iter()
                .try_fold(0usize, |sum, count| sum.checked_add(*count))
                != Some(total)
            {
                return None;
            }
            let points = read_points(bytes, body + 24 + polygons * 4, total, small, limit)?;
            if !player.bracketed {
                player.flush_pending();
            }
            let closed = matches!(kind, 8 | 91);
            let mut start = 0usize;
            for count in counts {
                let Some(run) = points.get(start..start + count) else {
                    break;
                };
                start += count;
                let Some(first) = run.first().copied() else {
                    continue;
                };
                player.move_to(first.0, first.1);
                for point in &run[1..] {
                    player.line_to(point.0, point.1);
                }
                if closed {
                    player.push(PathCommand::Close);
                }
            }
            player.paint_figure(closed);
        }
        42 | 43 => {
            let rect = box_rect()?;
            if !player.bracketed {
                player.flush_pending();
            }
            if kind == 42 {
                player.append_ellipse(rect);
            } else {
                player.append_rect(rect);
            }
            player.paint_figure(true);
        }
        47 => {
            let (Some(rect), Some(sx), Some(sy), Some(ex), Some(ey)) = (
                box_rect(),
                i32_at(bytes, body + 16),
                i32_at(bytes, body + 20),
                i32_at(bytes, body + 24),
                i32_at(bytes, body + 28),
            ) else {
                return None;
            };
            if !player.bracketed {
                player.flush_pending();
            }
            let centre = ((rect.0 + rect.2) / 2.0, (rect.1 + rect.3) / 2.0);
            player.move_to(centre.0, centre.1);
            player.append_arc(
                rect,
                (f64::from(sx), f64::from(sy)),
                (f64::from(ex), f64::from(ey)),
            );
            player.push(PathCommand::Close);
            player.paint_figure(true);
        }
        76 if !FULL => crate::gradient::bitblt(player, bytes, body)?,
        118 => crate::gradient::gradient_fill(player, bytes, body)?,
        // 82 EXTCREATEFONTINDIRECTW takes a handle the strict replay does not
        // model; parking an opaque object keeps later selections addressing
        // the right slot.
        82 => {
            let handle = u32_at(bytes, body)?;
            let object = if FULL {
                GdiObject::Font(Arc::new(logfont(bytes, body + 4, true)?))
            } else {
                GdiObject::Opaque
            };
            player.store(handle as usize, object);
        }
        20 if !FULL && u32_at(bytes, body)? == 13 => {}
        // 28 SETMETARGN only sets state the strict replay never reads.
        1 | 13 | 16 | 18 | 21 | 22 | 24 | 25 | 28 | 58 | 69 | 70 | 98 if !FULL => {}
        _ if !FULL => return None,
        _ => full_record(player, bytes, kind, body)?,
    }
    Some(())
}

/// Records only the full profile replays.
fn full_record<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    kind: u32,
    body: usize,
) -> Option<()> {
    let int = |at: usize| i32_at(bytes, body + at).map(f64::from);
    let box_rect =
        || -> Option<(f64, f64, f64, f64)> { Some((int(0)?, int(4)?, int(8)?, int(12)?)) };
    match kind {
        1 | 13 | 16 | 21 | 23 | 48 | 50 | 51 | 52 | 65 | 69 | 98 | 100 | 104 | 105 | 106 | 109
        | 110 | 111 | 112 | 113 | 119 | 120 | 121 => {}
        70 => crate::emfplus::comment(player, bytes)?,
        18 => player.dc.bk_opaque = u32_at(bytes, body)? == 2,
        20 => {
            let mode = u32_at(bytes, body)?;
            if !(1..=16).contains(&mode) {
                return player.refuse(format!("raster operation {mode} is not defined"));
            }
            player.dc.rop2 = mode;
            if !matches!(mode, 1 | 11 | 13 | 16) {
                player.omit("mixing raster operations drawn as copies")?;
            }
        }
        22 => player.dc.text_align = u32_at(bytes, body)?,
        24 => player.dc.text_color = u32_at(bytes, body)?,
        25 => player.dc.bk_color = u32_at(bytes, body)?,
        26 => {
            let (dx, dy) = (int(0)?, int(4)?);
            let m = player.logical_to_output();
            player.offset_clip(dx * m[0] + dy * m[2], dx * m[1] + dy * m[3])?;
        }
        28 => player.set_meta_region(),
        29 | 30 => {
            player.flush_pending();
            let rect = box_rect()?;
            let path = player.logical_rect_path(rect);
            if kind == 30 {
                match axis_aligned_rect(&path) {
                    Some(rect) => player.intersect_rect(rect)?,
                    None => player.combine_clip(path, false, Combine::And)?,
                }
            } else {
                player.combine_clip(path, false, Combine::Diff)?;
            }
        }
        31 | 32 => {
            let (xn, xd, yn, yd) = (int(0)?, int(4)?, int(8)?, int(12)?);
            if xd == 0.0 || yd == 0.0 || xn == 0.0 || yn == 0.0 {
                return Some(());
            }
            let ext = if kind == 31 {
                &mut player.dc.viewport_ext
            } else {
                &mut player.dc.window_ext
            };
            *ext = (ext.0 * xn / xd, ext.1 * yn / yd);
        }
        15 => {
            let (x, y, color) = (int(0)?, int(4)?, u32_at(bytes, body + 8)?);
            player.flush_pending();
            let (px, py) = player.point(x, y);
            let pixel = player.device_pixel();
            player.path = crate::player::rect_path([px, py, px + pixel, py + pixel]);
            player.emit(
                Some(crate::drawing::Paint::Solid(
                    crate::drawing::Rgba::from_colorref(color),
                )),
                None,
            );
        }
        58 => {
            let raw = u32_at(bytes, body)?;
            let limit = if raw > 0x00ff_ffff {
                f64::from(f32::from_bits(raw))
            } else {
                f64::from(raw)
            };
            if limit.is_finite() && limit > 0.0 {
                player.dc.miter_limit = limit;
            }
        }
        49 | 99 | 122 => {
            let handle = u32_at(bytes, body)?;
            player.store(handle as usize, GdiObject::Opaque);
        }
        101 => {
            if let Some(slot) = player.objects.get_mut(u32_at(bytes, body)? as usize) {
                *slot = None;
            }
        }
        115 => {
            if u32_at(bytes, body)? & 1 != 0 {
                player.omit("right-to-left layouts drawn left to right")?;
            }
        }
        41 => {
            let (cx, cy) = (int(0)?, int(4)?);
            let radius = f64::from(u32_at(bytes, body + 8)?);
            let (start, sweep) = (finite_at(bytes, body + 12)?, finite_at(bytes, body + 16)?);
            let (from, sweep) = (-start.to_radians(), -sweep.to_radians());
            let first = (cx + radius * from.cos(), cy + radius * from.sin());
            player.line_to(first.0, first.1);
            player.append_sweep((cx, cy), (radius, radius), from, sweep);
            player.pending_stroke = !player.bracketed;
        }
        44 => {
            let rect = box_rect()?;
            let corner = (int(16)?, int(20)?);
            if !player.bracketed {
                player.flush_pending();
            }
            player.append_round_rect(rect, corner);
            player.paint_figure(true);
        }
        45 | 46 | 55 => {
            let rect = box_rect()?;
            let (start, end) = ((int(16)?, int(20)?), (int(24)?, int(28)?));
            match kind {
                55 => {
                    let before = player.path.len();
                    player.append_arc(rect, start, end);
                    if let Some(PathCommand::Move { x, y }) = player.path.get(before).cloned() {
                        player.path[before] = PathCommand::Line { x, y };
                        if before == 0 {
                            let (cx, cy) = player.current;
                            let (px, py) = player.point(cx, cy);
                            player.path.insert(0, PathCommand::Move { x: px, y: py });
                        }
                    }
                    player.pending_stroke = !player.bracketed;
                }
                _ => {
                    if !player.bracketed {
                        player.flush_pending();
                    }
                    player.append_arc(rect, start, end);
                    if kind == 46 {
                        player.push(PathCommand::Close);
                    }
                    player.paint_figure(kind == 46);
                }
            }
        }
        56 | 92 => {
            let small = kind == 92;
            let count = u32_at(bytes, body + 16)? as usize;
            let points = read_points(
                bytes,
                body + 20,
                count,
                small,
                player.limits.points_per_record,
            )?;
            let types = crate::read::span(
                bytes,
                body + 20 + count * if small { 4 } else { 8 },
                count,
                1,
            )?;
            if !player.bracketed {
                player.flush_pending();
            }
            let mut index = 0;
            while index < count {
                let (point, kind) = (points[index], types[index]);
                match kind & !1 {
                    6 => {
                        player.move_to(point.0, point.1);
                        index += 1;
                    }
                    2 => {
                        player.line_to(point.0, point.1);
                        index += 1;
                    }
                    4 if index + 2 < count => {
                        player.cubic_to([points[index], points[index + 1], points[index + 2]]);
                        index += 3;
                    }
                    _ => return player.refuse("a POLYDRAW point type is not defined"),
                }
                if types[index - 1] & 1 != 0 {
                    player.close_figure();
                }
            }
            player.pending_stroke = !player.bracketed;
        }
        66 => player.widened = true,
        53 => player.omit("flood fills")?,
        73 => player.omit("inverted regions")?,
        102 | 103 => player.omit("OpenGL records")?,
        71 | 72 | 74 => {
            let size = u32_at(bytes, body + 16)? as usize;
            let (brush, data) = match kind {
                74 => (None, body + 20),
                71 => (Some(u32_at(bytes, body + 20)?), body + 24),
                _ => (Some(u32_at(bytes, body + 20)?), body + 32),
            };
            player.flush_pending();
            let path = region_path(player, bytes, data, size)?;
            player.charge(path.len(), 0)?;
            let selected = match brush {
                Some(handle) if handle & 0x8000_0000 != 0 => stock_object(handle & 0x7fff_ffff),
                Some(handle) => player.objects.get(handle as usize).cloned().flatten(),
                None => None,
            };
            let saved = player.dc.brush.clone();
            if let Some(GdiObject::Brush(brush)) = selected {
                player.dc.brush = Some(brush);
            }
            player.path = path;
            if kind == 72 {
                let width = f64::from(i32_at(bytes, body + 24)?.max(1));
                let fill = player.brush_fill();
                let stroke = fill.map(|paint| crate::drawing::Stroke {
                    paint,
                    width: width * player.device_pixel(),
                    dash: None,
                    cap: crate::drawing::LineCap::Butt,
                    join: crate::drawing::LineJoin::Miter,
                    miter_limit: 10.0,
                });
                player.emit(None, stroke);
            } else {
                let fill = player.brush_fill();
                player.emit(fill, None);
            }
            player.dc.brush = saved;
        }
        93 | 94 => crate::blit::pattern_brush(player, bytes, kind, body)?,
        76 | 77 | 78 | 79 | 80 | 81 | 114 | 116 => crate::blit::blit(player, bytes, kind, body)?,
        83 | 84 | 96 | 97 | 108 => crate::text::emf_text(player, bytes, kind, body)?,
        _ => return player.refuse(format!("EMF record type {kind} is not supported")),
    }
    Some(())
}

/// An `RGNDATA` block's rectangles, in device units, as one path in output units.
pub(crate) fn region_path<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    at: usize,
    size: usize,
) -> Option<Vec<PathCommand>> {
    let data = bytes.get(at..at.checked_add(size)?)?;
    let count = u32_at(data, 8)? as usize;
    let header = u32_at(data, 0)? as usize;
    if header < 32 || count > player.limits.clip_rects {
        return player.refuse("a region holds more rectangles than the limit");
    }
    crate::read::span(data, header, count, 16)?;
    let mut path = Vec::with_capacity(count * 5);
    for index in 0..count {
        let at = header + index * 16;
        let (l, t, r, b) = (
            f64::from(i32_at(data, at)?),
            f64::from(i32_at(data, at + 4)?),
            f64::from(i32_at(data, at + 8)?),
            f64::from(i32_at(data, at + 12)?),
        );
        let (l, t) = player.device_point(l.min(r), t.min(b));
        let (r, b) = player.device_point(
            f64::from(i32_at(data, at)?.max(i32_at(data, at + 8)?)),
            f64::from(i32_at(data, at + 4)?.max(i32_at(data, at + 12)?)),
        );
        path.extend(crate::player::rect_path([l, t, r, b]));
    }
    if path.is_empty() {
        path.extend(crate::player::rect_path([0.0, 0.0, 0.0, 0.0]));
    }
    Some(path)
}
