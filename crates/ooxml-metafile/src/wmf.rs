//! WMF (`[MS-WMF]`) records.

use std::sync::Arc;

use crate::drawing::PathCommand;
use crate::emf::{brush_from_style, full_brush, full_pen, logfont, pen_from_style, read_points};
use crate::player::{Combine, GdiObject, MAX_HANDLES, Player, SharedBudget, axis_aligned_rect};
use crate::read::{i16_at, u16_at, u32_at};

pub(crate) const WMF_PLACEABLE_KEY: u32 = 0x9AC6_CDD7;

pub(crate) fn is_wmf(bytes: &[u8]) -> bool {
    u32_at(bytes, 0) == Some(WMF_PLACEABLE_KEY)
        || (matches!(u16_at(bytes, 0), Some(1 | 2)) && u16_at(bytes, 2) == Some(9))
}

pub(crate) struct FreeObjects {
    slots: [u64; MAX_HANDLES / 64],
    groups: u64,
}

impl FreeObjects {
    pub(crate) fn new() -> Self {
        Self {
            slots: [u64::MAX; MAX_HANDLES / 64],
            groups: u64::MAX,
        }
    }

    fn take(&mut self) -> Option<usize> {
        if self.groups == 0 {
            return None;
        }
        let group = self.groups.trailing_zeros() as usize;
        let slot = self.slots[group].trailing_zeros() as usize;
        self.slots[group] &= !(1 << slot);
        if self.slots[group] == 0 {
            self.groups &= !(1 << group);
        }
        Some(group * 64 + slot)
    }

    fn release(&mut self, slot: usize) {
        self.slots[slot / 64] |= 1 << (slot % 64);
        self.groups |= 1 << (slot / 64);
    }
}

pub(crate) fn store_object<const FULL: bool>(player: &mut Player<FULL>, object: GdiObject) {
    if let Some(index) = player.free_objects.take() {
        player.store(index, object);
    }
}

/// The records of a WMF, located and bounded, with the frame they draw in.
pub(crate) struct WmfRecords {
    pub records: Vec<(usize, usize, usize)>,
    /// `(x, y, width, height)` in logical units.
    pub frame: (f64, f64, f64, f64),
    /// The placeable header's picture size in CSS pixels.
    pub size: Option<(f64, f64)>,
    /// Whether a placeable header or a window extent sizes the picture.
    pub framed: bool,
    pub handles: usize,
}

pub(crate) fn wmf_records(
    bytes: &[u8],
    max_records: usize,
    budget: Option<&SharedBudget>,
    default_origin: bool,
) -> Option<WmfRecords> {
    let placeable = u32_at(bytes, 0)? == WMF_PLACEABLE_KEY;
    let header = if placeable { 22 } else { 0 };
    let kind = u16_at(bytes, header)?;
    if kind != 1 && kind != 2 {
        return None;
    }
    if u16_at(bytes, header + 2)? != 9 {
        return None;
    }
    let placeable_frame = placeable
        .then(|| {
            let left = f64::from(i16_at(bytes, 6)?);
            let top = f64::from(i16_at(bytes, 8)?);
            let right = f64::from(i16_at(bytes, 10)?);
            let bottom = f64::from(i16_at(bytes, 12)?);
            Some((left, top, right - left, bottom - top))
        })
        .flatten();
    let inch = placeable
        .then(|| u16_at(bytes, 14))
        .flatten()
        .filter(|inch| *inch > 0)
        .map(f64::from);
    let size = placeable_frame
        .zip(inch)
        .map(|(frame, inch)| (frame.2.abs() / inch * 96.0, frame.3.abs() / inch * 96.0))
        .filter(|(width, height)| *width > 0.0 && *height > 0.0);
    let handles = u16_at(bytes, header + 10)? as usize;
    let mut offset = header + 18;
    let mut eof = false;
    let mut count = 0usize;
    let mut records: Vec<(usize, usize, usize)> = Vec::new();
    while offset + 6 <= bytes.len() {
        if budget.is_some_and(|budget| !budget.spend(1, 0)) {
            return None;
        }
        let size = (u32_at(bytes, offset)? as usize).checked_mul(2)?;
        let function = u16_at(bytes, offset + 4)?;
        let end = offset.checked_add(size)?;
        if size < 6 || end > bytes.len() {
            return None;
        }
        count += 1;
        if count > max_records {
            return None;
        }
        if function == 0 {
            eof = true;
            break;
        }
        records.push((function as usize, offset, end));
        offset = end;
    }
    if !eof {
        return None;
    }
    let mut org = None;
    let mut ext = None;
    for (function, start, end) in &records {
        let bytes = &bytes[*start..*end];
        let body = 6;
        match function {
            0x020B => org = org.or(i16_at(bytes, body + 2).zip(i16_at(bytes, body))),
            0x020C => ext = ext.or(i16_at(bytes, body + 2).zip(i16_at(bytes, body))),
            _ => {}
        }
    }
    let org = if placeable || !default_origin {
        org
    } else {
        org.or(Some((0, 0)))
    };
    let mut frame = placeable_frame.unwrap_or((0.0, 0.0, 1.0, 1.0));
    let windowed = if let (Some(org), Some((width, height))) = (org, ext)
        && width != 0
        && height != 0
    {
        frame = (
            f64::from(org.0),
            f64::from(org.1),
            f64::from(width),
            f64::from(height),
        );
        true
    } else {
        false
    };
    if frame.2.abs() < f64::EPSILON || frame.3.abs() < f64::EPSILON {
        return None;
    }
    Some(WmfRecords {
        records,
        frame,
        size,
        framed: placeable || windowed,
        handles,
    })
}

pub(crate) fn play_wmf<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    records: &[(usize, usize, usize)],
) -> Option<()> {
    player.dc.window_org = (player.frame.0, player.frame.1);
    player.dc.window_ext = (player.frame.2, player.frame.3);
    player.dc.viewport_ext = player.dc.window_ext;
    player.dc.scaled = true;
    player.dc.map_mode = 8;
    player.frame.0 = 0.0;
    player.frame.1 = 0.0;
    for (index, (function, start, end)) in records.iter().enumerate() {
        if wmf_record(player, &bytes[*start..*end], *function, 6).is_none() {
            if player.refusal.is_none() {
                player.refusal = Some(format!(
                    "WMF record {} (function {function:#06x}) could not be replayed",
                    index + 1
                ));
            }
            return None;
        }
        if player.overflowed {
            return player.refuse("the metafile draws more than the replay limits");
        }
    }
    player.flush_pending();
    if player.overflowed {
        return player.refuse("the metafile draws more than the replay limits");
    }
    Some(())
}

fn wmf_record<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    function: usize,
    body: usize,
) -> Option<()> {
    let point_at = |offset: usize| -> Option<(f64, f64)> {
        Some((
            f64::from(i16_at(bytes, offset + 2)?),
            f64::from(i16_at(bytes, offset)?),
        ))
    };
    let rect_at = |offset: usize| -> Option<(f64, f64, f64, f64)> {
        let (Some(bottom), Some(right), Some(top), Some(left)) = (
            i16_at(bytes, offset),
            i16_at(bytes, offset + 2),
            i16_at(bytes, offset + 4),
            i16_at(bytes, offset + 6),
        ) else {
            return None;
        };
        Some((
            f64::from(left),
            f64::from(top),
            f64::from(right),
            f64::from(bottom),
        ))
    };
    match function {
        0x020B => player.dc.window_org = point_at(body)?,
        0x020C => {
            let ext = point_at(body)?;
            if ext.0 == 0.0 || ext.1 == 0.0 {
                return FULL.then_some(());
            }
            player.dc.window_ext = ext;
        }
        0x020D => player.dc.viewport_org = point_at(body)?,
        0x020E => player.dc.viewport_ext = point_at(body)?,
        0x0106 => {
            player.dc.even_odd = match u16_at(bytes, body)? {
                1 => true,
                2 => false,
                _ => return None,
            }
        }
        0x001E => player.save()?,
        0x0127 => player.restore(i32::from(i16_at(bytes, body)?))?,
        0x0214 => {
            let point = point_at(body)?;
            player.flush_pending();
            player.move_to(point.0, point.1);
        }
        0x0213 => {
            let point = point_at(body)?;
            player.line_to(point.0, point.1);
            player.pending_stroke = true;
        }
        0x02FA => {
            let (Some(style), Some(width), Some(color)) = (
                u16_at(bytes, body),
                i16_at(bytes, body + 2),
                u32_at(bytes, body + 6),
            ) else {
                return None;
            };
            let pen = if FULL {
                full_pen(u32::from(style), f64::from(width), color, None)
            } else {
                pen_from_style(u32::from(style), f64::from(width), color)?
            };
            store_object(player, GdiObject::Pen(pen));
        }
        0x02FC => {
            let (Some(style), Some(color)) = (u16_at(bytes, body), u32_at(bytes, body + 2)) else {
                return None;
            };
            let brush = if FULL {
                full_brush(
                    u32::from(style),
                    color,
                    u32::from(u16_at(bytes, body + 6).unwrap_or(0)),
                )
            } else {
                brush_from_style(u32::from(style), color)?
            };
            store_object(player, GdiObject::Brush(brush));
        }
        0x012D => {
            let index = u16_at(bytes, body)?;
            player.flush_pending();
            if let Some(object) = player.objects.get(index as usize).cloned().flatten() {
                player.select(object);
            }
        }
        0x01F0 => {
            if let Some(index) = u16_at(bytes, body)
                && let Some(slot) = player.objects.get_mut(index as usize)
            {
                *slot = None;
                player.free_objects.release(index as usize);
            }
        }
        0x0324 | 0x0325 => {
            let count = u16_at(bytes, body)?;
            let limit = player.limits.points_per_record;
            let points = read_points(bytes, body + 2, count as usize, true, limit)?;
            let first = points.first().copied()?;
            player.flush_pending();
            player.move_to(first.0, first.1);
            for point in &points[1..] {
                player.line_to(point.0, point.1);
            }
            let closed = function == 0x0324;
            if closed {
                player.push(PathCommand::Close);
            }
            player.paint_figure(closed);
        }
        0x0538 => {
            let polygons = u16_at(bytes, body)? as usize;
            let limit = player.limits.points_per_record;
            if polygons > limit {
                return None;
            }
            let mut counts = Vec::with_capacity(polygons);
            let mut total = 0usize;
            for index in 0..polygons {
                let count = u16_at(bytes, body + 2 + index * 2)?;
                total += count as usize;
                counts.push(count as usize);
            }
            let points = read_points(bytes, body + 2 + polygons * 2, total, true, limit)?;
            player.flush_pending();
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
                player.push(PathCommand::Close);
            }
            player.paint_figure(true);
        }
        0x041B | 0x0418 => {
            let rect = rect_at(body)?;
            player.flush_pending();
            if function == 0x0418 {
                player.append_ellipse(rect);
            } else {
                player.append_rect(rect);
            }
            player.paint_figure(true);
        }
        // CREATEPALETTE and CREATEFONTINDIRECT consume the next handle slot,
        // which WMF assigns by position, so they have to be stored rather than
        // skipped even though the strict replay cannot draw with them.
        0x00F7 => store_object(player, GdiObject::Opaque),
        0x02FB => {
            let object = if FULL {
                GdiObject::Font(Arc::new(logfont(bytes, body, false)?))
            } else {
                GdiObject::Opaque
            };
            store_object(player, object);
        }
        0x0103 if FULL => {
            if (2..=6).contains(&u16_at(bytes, body)?) {
                player.omit("WMF fixed mapping modes drawn in the picture frame")?;
            }
        }
        0x0103 if u16_at(bytes, body)? == 8 => {}
        // SETROP2, SETRELABS, SETTEXTALIGN and ESCAPE set state the strict
        // replay never reads.
        0x0102 | 0x0104 | 0x0105 | 0x0107 | 0x0108 | 0x012E | 0x0201 | 0x0209 | 0x020A | 0x0626
            if !FULL => {}
        _ if !FULL => return None,
        _ => full_record(player, bytes, function, body)?,
    }
    Some(())
}

fn full_record<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    function: usize,
    body: usize,
) -> Option<()> {
    let word = |at: usize| i16_at(bytes, body + at).map(f64::from);
    let rect_at = |at: usize| -> Option<(f64, f64, f64, f64)> {
        Some((word(at + 6)?, word(at + 4)?, word(at + 2)?, word(at)?))
    };
    match function {
        0x0035 | 0x0037 | 0x0105 | 0x0107 | 0x0108 | 0x0139 | 0x0149 | 0x020A | 0x0231 | 0x0234
        | 0x0436 | 0x0626 => {}
        0x0102 => player.dc.bk_opaque = u16_at(bytes, body)? == 2,
        0x0104 => {
            let mode = u32::from(u16_at(bytes, body)?);
            if !(1..=16).contains(&mode) {
                return player.refuse(format!("raster operation {mode} is not defined"));
            }
            player.dc.rop2 = mode;
            if !matches!(mode, 1 | 11 | 13 | 16) {
                player.omit("mixing raster operations drawn as copies")?;
            }
        }
        0x012E => player.dc.text_align = u32::from(u16_at(bytes, body)?),
        0x0201 => player.dc.bk_color = u32_at(bytes, body)?,
        0x0209 => player.dc.text_color = u32_at(bytes, body)?,
        0x020F | 0x0211 => {
            let (dy, dx) = (word(0)?, word(2)?);
            let origin = if function == 0x020F {
                &mut player.dc.window_org
            } else {
                &mut player.dc.viewport_org
            };
            *origin = (origin.0 + dx, origin.1 + dy);
        }
        0x0410 | 0x0412 => {
            let (yd, yn, xd, xn) = (word(0)?, word(2)?, word(4)?, word(6)?);
            if xd == 0.0 || yd == 0.0 || xn == 0.0 || yn == 0.0 {
                return Some(());
            }
            let ext = if function == 0x0410 {
                &mut player.dc.window_ext
            } else {
                &mut player.dc.viewport_ext
            };
            *ext = (ext.0 * xn / xd, ext.1 * yn / yd);
        }
        0x0415 | 0x0416 => {
            player.flush_pending();
            let path = player.logical_rect_path(rect_at(0)?);
            if function == 0x0416 {
                match axis_aligned_rect(&path) {
                    Some(rect) => player.intersect_rect(rect)?,
                    None => player.combine_clip(path, false, Combine::And)?,
                }
            } else {
                player.combine_clip(path, false, Combine::Diff)?;
            }
        }
        0x0220 => {
            let (dy, dx) = (word(0)?, word(2)?);
            let m = player.logical_to_output();
            player.offset_clip(dx * m[0] + dy * m[2], dx * m[1] + dy * m[3])?;
        }
        0x061C => {
            let (height, width) = (word(0)?, word(2)?);
            let rect = rect_at(4)?;
            player.flush_pending();
            player.append_round_rect(rect, (width, height));
            player.paint_figure(true);
        }
        0x0817 | 0x081A | 0x0830 => {
            let end = (word(2)?, word(0)?);
            let start = (word(6)?, word(4)?);
            let rect = rect_at(8)?;
            player.flush_pending();
            let pie = function == 0x081A;
            if pie {
                let centre = ((rect.0 + rect.2) / 2.0, (rect.1 + rect.3) / 2.0);
                player.move_to(centre.0, centre.1);
            }
            player.append_arc(rect, start, end);
            let closed = function != 0x0817;
            if closed {
                player.push(PathCommand::Close);
            }
            player.paint_figure(closed);
        }
        0x041F => {
            let color = u32_at(bytes, body)?;
            let (y, x) = (word(4)?, word(6)?);
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
        0x0419 | 0x0548 => player.omit("flood fills")?,
        0x012A => player.omit("inverted regions")?,
        0x06FF => {
            let region = region_rects(bytes, body, player.limits.clip_rects);
            store_object(player, region.map_or(GdiObject::Opaque, GdiObject::Region));
        }
        0x012C => {
            let index = u16_at(bytes, body)? as usize;
            player.flush_pending();
            match player.objects.get(index).cloned().flatten() {
                Some(GdiObject::Region(rects)) => {
                    let path = region_path(player, &rects);
                    player.combine_clip(path, false, Combine::Copy)?;
                }
                _ => player.reset_clip(),
            }
        }
        0x0228 | 0x0429 | 0x012B => {
            let region = u16_at(bytes, body)? as usize;
            let brush = match function {
                0x012B => None,
                _ => Some(u16_at(bytes, body + 2)? as usize),
            };
            player.flush_pending();
            let Some(GdiObject::Region(rects)) = player.objects.get(region).cloned().flatten()
            else {
                return Some(());
            };
            let saved = player.dc.brush.clone();
            if let Some(index) = brush
                && let Some(GdiObject::Brush(brush)) = player.objects.get(index).cloned().flatten()
            {
                player.dc.brush = Some(brush);
            }
            let path = region_path(player, &rects);
            player.charge(path.len(), 0)?;
            player.path = path;
            let fill = player.brush_fill();
            if function == 0x0429 {
                let width = word(6).unwrap_or(1.0).max(1.0) * player.device_pixel();
                let stroke = fill.map(|paint| crate::drawing::Stroke {
                    paint,
                    width,
                    dash: None,
                    cap: crate::drawing::LineCap::Butt,
                    join: crate::drawing::LineJoin::Miter,
                    miter_limit: 10.0,
                });
                player.emit(None, stroke);
            } else {
                player.emit(fill, None);
            }
            player.dc.brush = saved;
        }
        0x01F9 | 0x0142 => crate::blit::wmf_pattern_brush(player, bytes, function, body)?,
        0x061D | 0x0922 | 0x0B23 | 0x0940 | 0x0B41 | 0x0F43 | 0x0D33 => {
            crate::blit::wmf_blit(player, bytes, function, body)?
        }
        0x0521 | 0x0A32 => crate::text::wmf_text(player, bytes, function, body)?,
        _ => {
            return player.refuse(format!(
                "WMF record function {function:#06x} is not supported"
            ));
        }
    }
    Some(())
}

/// A `META_CREATEREGION` object's scan rectangles in logical units.
fn region_rects(bytes: &[u8], body: usize, limit: usize) -> Option<Arc<[[f64; 4]]>> {
    let scans = u16_at(bytes, body + 10)? as usize;
    let mut at = body + 22;
    let mut rects = Vec::new();
    for _ in 0..scans {
        let count = u16_at(bytes, at)? as usize;
        let (top, bottom) = (i16_at(bytes, at + 2)?, i16_at(bytes, at + 4)?);
        for pair in 0..count / 2 {
            let left = i16_at(bytes, at + 6 + pair * 4)?;
            let right = i16_at(bytes, at + 8 + pair * 4)?;
            if rects.len() >= limit {
                return None;
            }
            rects.push([
                f64::from(left),
                f64::from(top),
                f64::from(right),
                f64::from(bottom),
            ]);
        }
        at += 8 + count * 2;
    }
    Some(rects.into())
}

fn region_path<const FULL: bool>(player: &Player<FULL>, rects: &[[f64; 4]]) -> Vec<PathCommand> {
    let mut path = Vec::with_capacity(rects.len() * 5 + 5);
    for rect in rects {
        path.extend(player.logical_rect_path((rect[0], rect[1], rect[2], rect[3])));
    }
    if path.is_empty() {
        path.extend(crate::player::rect_path([0.0, 0.0, 0.0, 0.0]));
    }
    path
}

/// The EMF a WMF carries in `META_ESCAPE_ENHANCED_METAFILE` comments, when
/// its chunks arrive complete and in order.
pub(crate) fn embedded_emf(bytes: &[u8], records: &[(usize, usize, usize)]) -> Option<Vec<u8>> {
    const WMFC: u32 = 0x4346_4D57;
    let mut emf: Vec<u8> = Vec::new();
    let mut expected = None;
    for (function, start, end) in records {
        if *function != 0x0626 {
            continue;
        }
        let record = &bytes[*start..*end];
        if u16_at(record, 6)? != 0x000F || u32_at(record, 10) != Some(WMFC) {
            continue;
        }
        if u32_at(record, 14)? != 1 {
            continue;
        }
        let total = u32_at(record, 40)? as usize;
        let size = u32_at(record, 32)? as usize;
        if *expected.get_or_insert(total) != total || emf.len().checked_add(size)? > total {
            return None;
        }
        emf.extend_from_slice(record.get(44..44usize.checked_add(size)?)?);
    }
    (expected == Some(emf.len()) && crate::emf::is_emf(&emf)).then_some(emf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::player::Brush;

    #[test]
    fn object_churn_reuses_the_lowest_free_slot() {
        for handles in [0, MAX_HANDLES] {
            let mut player = Player::<true>::new((0.0, 0.0, 10.0, 10.0), handles, (10.0, 10.0));
            for index in 0..MAX_HANDLES {
                store_object(
                    &mut player,
                    GdiObject::Brush(Brush::solid(index as u32, true)),
                );
                let Some(GdiObject::Brush(brush)) = &player.objects[index] else {
                    panic!("a brush in the next free slot");
                };
                assert_eq!(brush.color, index as u32);
            }
            store_object(&mut player, GdiObject::Opaque);
            assert_eq!(player.objects.len(), MAX_HANDLES);
            assert_eq!(player.free_objects.groups, 0);
            for index in [4095u16, 64, 0, 63] {
                wmf_record(&mut player, &index.to_le_bytes(), 0x01F0, 0).unwrap();
            }
            for index in [0, 63, 64, 4095] {
                store_object(&mut player, GdiObject::Opaque);
                assert!(matches!(player.objects[index], Some(GdiObject::Opaque)));
            }
            for _ in 0..32 {
                wmf_record(&mut player, &0u16.to_le_bytes(), 0x01F0, 0).unwrap();
                assert_eq!(player.free_objects.groups, 1);
                assert_eq!(player.free_objects.slots[0], 1);
                store_object(&mut player, GdiObject::Opaque);
                assert!(matches!(player.objects[0], Some(GdiObject::Opaque)));
                assert_eq!(player.free_objects.groups, 0);
            }
        }
    }
}
