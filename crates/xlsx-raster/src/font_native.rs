use super::*;

pub(crate) struct FontFace {
    pub data: std::sync::Arc<[u8]>,
    pub bold: bool,
    pub italic: bool,
}

fn segments<'a>(faces: &[FontFace], text: &'a str) -> Vec<(usize, &'a str)> {
    let parsed: Vec<_> = faces
        .iter()
        .map(|face| rustybuzz::Face::from_slice(&face.data, 0).unwrap())
        .collect();
    let mut segments: Vec<(usize, &str)> = Vec::new();
    let mut start = 0;
    for (offset, cluster) in text.grapheme_indices(true) {
        let index = parsed
            .iter()
            .position(|face| {
                cluster.chars().all(|ch| {
                    ch.is_control()
                        || ch.is_whitespace()
                        || matches!(ch as u32, 0x200b..=0x200f | 0xfe00..=0xfe0f)
                        || face.glyph_index(ch).is_some()
                })
            })
            .unwrap_or(usize::MAX);
        if let Some((previous, _)) = segments.last()
            && *previous == index
        {
            segments.last_mut().unwrap().1 = &text[start..offset + cluster.len()];
        } else {
            start = offset;
            segments.push((index, &text[offset..offset + cluster.len()]));
        }
    }
    segments
}

fn face_at(faces: &[FontFace], index: usize) -> rustybuzz::Face<'_> {
    rustybuzz::Face::from_slice(
        faces
            .get(index)
            .map_or(FONT_BYTES, |face| face.data.as_ref()),
        0,
    )
    .unwrap()
}

pub(crate) fn measure_text_with_fonts(faces: &[FontFace], text: &str, size: f32) -> f32 {
    if faces.is_empty() {
        return measure_text(text, size);
    }
    segments(faces, text)
        .into_iter()
        .map(|(index, text)| {
            let face = face_at(faces, index);
            let glyphs = shape(&face, text);
            glyphs
                .glyph_positions()
                .iter()
                .map(|position| position.x_advance)
                .sum::<i32>() as f32
                * size
                * PX_PER_PT
                / face.units_per_em() as f32
        })
        .sum()
}

/// a single-line text run; `x`/`y` are the align anchor and alphabetic baseline.
#[derive(Clone, Copy)]
pub struct TextRun<'a> {
    pub faces: &'a [FontFace],
    pub x: f32,
    pub y: f32,
    pub text: &'a str,
    pub font_size_pt: f32,
    pub color: Color,
    pub align: Align,
    pub clip: &'a DlRect,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub highlight: Option<Color>,
    pub dashed_underline: bool,
}

/// paint one shaped, single-line run, clipped strictly to the run's clip rect.
pub fn paint_text(pixmap: &mut Pixmap, run: &TextRun) {
    if run.faces.is_empty() {
        paint_text_face(pixmap, run, &FACE);
        return;
    }
    let width = measure_text_with_fonts(run.faces, run.text, run.font_size_pt);
    let mut x = match run.align {
        Align::Left => run.x,
        Align::Center => run.x - width / 2.0,
        Align::Right => run.x - width,
    };
    for (index, text) in segments(run.faces, run.text) {
        let face = face_at(run.faces, index);
        let selected = run.faces.get(index);
        let part = TextRun {
            x,
            text,
            align: Align::Left,
            bold: run.bold && selected.is_none_or(|face| !face.bold),
            italic: run.italic && selected.is_none_or(|face| !face.italic),
            ..*run
        };
        paint_text_face(pixmap, &part, &face);
        let glyphs = shape(&face, text);
        x += glyphs
            .glyph_positions()
            .iter()
            .map(|position| position.x_advance)
            .sum::<i32>() as f32
            * run.font_size_pt
            * PX_PER_PT
            / face.units_per_em() as f32;
    }
}

fn paint_text_face(pixmap: &mut Pixmap, run: &TextRun, face: &rustybuzz::Face<'_>) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 {
        return;
    }

    let scale = run.font_size_pt * PX_PER_PT / face.units_per_em() as f32;
    let glyphs = shape(face, run.text);

    let total: i32 = glyphs.glyph_positions().iter().map(|p| p.x_advance).sum();
    let pen = match run.align {
        Align::Left => run.x,
        Align::Center => run.x - total as f32 * scale / 2.0,
        Align::Right => run.x - total as f32 * scale,
    };

    let mut paint = Paint::default();
    paint.set_color(run.color);
    paint.anti_alias = true;

    // tiny-skia mask clipping costs O(pixmap) per run; runs inside the clip
    // paint directly, overflowing runs paint into a clip-sized scratch and blit.
    let ascent = face.ascender() as f32 * scale;
    let descent = face.descender() as f32 * scale;
    let run_width = total as f32 * scale;
    let fully_inside = pen >= run.clip.x
        && pen + run_width <= run.clip.x + run.clip.w
        && run.y - ascent >= run.clip.y
        && run.y - descent <= run.clip.y + run.clip.h;

    if fully_inside {
        paint_run(
            pixmap, face, &glyphs, pen, run, scale, &paint, run_width, 0.0, 0.0,
        );
        return;
    }

    let ox = run.clip.x.floor();
    let oy = run.clip.y.floor();
    let sw = ((run.clip.x + run.clip.w).ceil() - ox).max(1.0) as u32;
    let sh = ((run.clip.y + run.clip.h).ceil() - oy).max(1.0) as u32;
    let Some(mut scratch) = Pixmap::new(sw, sh) else {
        return;
    };
    paint_run(
        &mut scratch,
        face,
        &glyphs,
        pen,
        run,
        scale,
        &paint,
        run_width,
        ox,
        oy,
    );
    pixmap.draw_pixmap(
        ox as i32,
        oy as i32,
        scratch.as_ref(),
        &tiny_skia::PixmapPaint::default(),
        Transform::identity(),
        None,
    );
}

pub fn paint_rotated_text(
    pixmap: &mut Pixmap,
    run: &TextRun,
    deg: f32,
    cell: &DlRect,
    valign: TextValign,
    number_format: Option<&NumberFormatLayout>,
) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 || cell.w <= 0.0 || cell.h <= 0.0 {
        return;
    }
    let primary = face_at(run.faces, 0);
    let face = &primary;
    let scale = run.font_size_pt * PX_PER_PT / face.units_per_em() as f32;
    let (segments, run_width) = match number_format {
        Some(layout) => {
            let mut effective = layout.clone();
            effective.left = 0.0;
            effective.width = rotated_available_length(cell, deg, run.font_size_pt);
            let mut raw: Vec<(String, f32)> = Vec::new();
            xlsx_render::paint_number_format(
                run.text,
                &effective,
                run.align,
                |literal| measure_text_with_fonts(run.faces, literal, run.font_size_pt),
                |literal, x| raw.push((literal.to_owned(), x)),
            );
            if raw.is_empty() {
                (
                    vec![(run.text.to_owned(), 0.0)],
                    measure_text_with_fonts(run.faces, run.text, run.font_size_pt),
                )
            } else {
                let min_x = raw.iter().map(|(_, x)| *x).fold(f32::INFINITY, f32::min);
                let max_x = raw
                    .iter()
                    .map(|(text, x)| x + measure_text_with_fonts(run.faces, text, run.font_size_pt))
                    .fold(f32::NEG_INFINITY, f32::max);
                let segments = raw
                    .into_iter()
                    .map(|(text, x)| (text, x - min_x))
                    .collect::<Vec<_>>();
                (segments, (max_x - min_x).max(0.0))
            }
        }
        None => (
            vec![(run.text.to_owned(), 0.0)],
            measure_text_with_fonts(run.faces, run.text, run.font_size_pt),
        ),
    };
    let ascent = face.ascender() as f32 * scale;
    let descent = -face.descender() as f32 * scale;
    let line_h = ascent + descent;
    let a = -deg.to_radians();
    let (sin, cos) = fmath::sin_cos(a);
    let bbox_w = (run_width * cos).abs() + (line_h * sin).abs();
    let bbox_h = (run_width * sin).abs() + (line_h * cos).abs();
    let bbox_left = match run.align {
        Align::Left => cell.x,
        Align::Center => cell.x + (cell.w - bbox_w) / 2.0,
        Align::Right => cell.x + cell.w - bbox_w,
    };
    let bbox_top = match valign {
        TextValign::Top => cell.y,
        TextValign::Center => cell.y + (cell.h - bbox_h) / 2.0,
        TextValign::Bottom => cell.y + cell.h - bbox_h,
    };
    let cx = bbox_left + bbox_w / 2.0;
    let cy = bbox_top + bbox_h / 2.0;
    let ox = run.clip.x.floor();
    let oy = run.clip.y.floor();
    let sw = ((run.clip.x + run.clip.w).ceil() - ox).max(1.0) as u32;
    let sh = ((run.clip.y + run.clip.h).ceil() - oy).max(1.0) as u32;
    let Some(mut scratch) = Pixmap::new(sw, sh) else {
        return;
    };
    let mut paint = Paint::default();
    paint.set_color(run.color);
    paint.anti_alias = true;
    let center_y_local = (-ascent + descent) / 2.0;
    if let Some(color) = run.highlight {
        let mut hp = Paint::default();
        hp.set_color(color);
        if let Some(rect) = Rect::from_xywh(
            -run_width / 2.0 - 2.0,
            -ascent - center_y_local - 1.0,
            run_width + 4.0,
            line_h + 2.0,
        ) {
            let transform = Transform::from_row(cos, sin, -sin, cos, cx - ox, cy - oy);
            scratch.fill_rect(rect, &hp, transform, None);
        }
    }
    for (segment, segment_offset) in &segments {
        let mut span_offset = *segment_offset;
        for (index, text) in self::segments(run.faces, segment) {
            let selected = run.faces.get(index);
            let selected_face = face_at(run.faces, index);
            let face = &selected_face;
            let scale = run.font_size_pt * PX_PER_PT / face.units_per_em() as f32;
            let shear = if run.italic && selected.is_none_or(|face| !face.italic) {
                ITALIC_SHEAR * scale
            } else {
                0.0
            };
            let glyphs = shape(face, text);
            let mut pen = 0.0_f32;
            for (info, pos) in glyphs.glyph_infos().iter().zip(glyphs.glyph_positions()) {
                if let Some(path) = glyph_path(face, info.glyph_id as u16) {
                    let tx0 = span_offset + pen + pos.x_offset as f32 * scale - run_width / 2.0;
                    let ty0 = -pos.y_offset as f32 * scale - center_y_local;
                    let e = cos * tx0 - sin * ty0 + cx - ox;
                    let f = sin * tx0 + cos * ty0 + cy - oy;
                    let transform = Transform::from_row(
                        cos * scale,
                        sin * scale,
                        cos * shear + sin * scale,
                        sin * shear - cos * scale,
                        e,
                        f,
                    );
                    scratch.fill_path(&path, &paint, FillRule::Winding, transform, None);
                    if run.bold && selected.is_none_or(|face| !face.bold) {
                        let bold = Transform::from_row(
                            cos * scale,
                            sin * scale,
                            cos * shear + sin * scale,
                            sin * shear - cos * scale,
                            e + cos * BOLD_OFFSET_PX,
                            f + sin * BOLD_OFFSET_PX,
                        );
                        scratch.fill_path(&path, &paint, FillRule::Winding, bold, None);
                    }
                }
                pen += pos.x_advance as f32 * scale;
            }
            span_offset += pen;
        }
    }
    if run.underline || run.strike || run.dashed_underline {
        let em = face.units_per_em() as f32;
        let mut bar = |center_up: f32, thickness: f32| {
            let h = (thickness * scale).max(0.5);
            let y_local = -center_up * scale - center_y_local;
            if let Some(rect) = Rect::from_xywh(-run_width / 2.0, y_local - h / 2.0, run_width, h) {
                let transform = Transform::from_row(cos, sin, -sin, cos, cx - ox, cy - oy);
                scratch.fill_rect(rect, &paint, transform, None);
            }
        };
        if run.underline || run.dashed_underline {
            let m = face.underline_metrics();
            bar(
                m.map(|m| m.position as f32).unwrap_or(-0.1 * em),
                m.map(|m| m.thickness as f32).unwrap_or(0.05 * em),
            );
        }
        if run.strike {
            let m = face.strikeout_metrics();
            bar(
                m.map(|m| m.position as f32).unwrap_or(0.26 * em),
                m.map(|m| m.thickness as f32).unwrap_or(0.05 * em),
            );
        }
    }
    pixmap.draw_pixmap(
        ox as i32,
        oy as i32,
        scratch.as_ref(),
        &tiny_skia::PixmapPaint::default(),
        Transform::identity(),
        None,
    );
}

pub fn paint_stacked_text(pixmap: &mut Pixmap, run: &TextRun, cell: &DlRect, valign: TextValign) {
    if run.clip.w <= 0.0 || run.clip.h <= 0.0 {
        return;
    }
    let line_h = stacked_line_height(run.font_size_pt);
    let ascent = run.font_size_pt * PX_PER_PT * STACKED_ASCENT_RATIO;
    let count = run.text.graphemes(true).count();
    let total_h = count as f32 * line_h;
    let top = match valign {
        TextValign::Top => cell.y,
        TextValign::Center => cell.y + (cell.h - total_h) / 2.0,
        TextValign::Bottom => cell.y + cell.h - total_h,
    };
    let clip_top = run.clip.y;
    let clip_bottom = run.clip.y + run.clip.h;
    for (i, cluster) in run.text.graphemes(true).enumerate() {
        let line_top = top + i as f32 * line_h;
        if line_top + line_h <= clip_top || line_top >= clip_bottom {
            continue;
        }
        let w = measure_text_with_fonts(run.faces, cluster, run.font_size_pt);
        let x = match run.align {
            Align::Left => cell.x,
            Align::Center => cell.x + (cell.w - w) / 2.0,
            Align::Right => cell.x + cell.w - w,
        };
        if x + w <= run.clip.x || x >= run.clip.x + run.clip.w {
            continue;
        }
        let sub = TextRun {
            faces: run.faces,
            x,
            y: line_top + ascent,
            text: cluster,
            font_size_pt: run.font_size_pt,
            color: run.color,
            align: Align::Left,
            clip: run.clip,
            bold: run.bold,
            italic: run.italic,
            underline: run.underline,
            strike: run.strike,
            highlight: run.highlight,
            dashed_underline: run.dashed_underline,
        };
        paint_text(pixmap, &sub);
    }
}
