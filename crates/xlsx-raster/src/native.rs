use super::*;
use std::sync::Arc;

const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONT_FACES: usize = 256;
const MAX_FONT_REGISTRATIONS: usize = 4_096;

#[derive(Default)]
pub struct RasterFonts {
    faces: HashMap<(String, bool, bool), Vec<font::FontFace>>,
    data: Vec<Arc<[u8]>>,
    registrations: usize,
}

impl RasterFonts {
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        data: Vec<u8>,
        face_bold: bool,
        face_italic: bool,
    ) -> Result<(), String> {
        if data.len() > MAX_FONT_BYTES {
            return Err(format!("font exceeds {MAX_FONT_BYTES} bytes"));
        }
        self.register_shared_font(family, bold, italic, data.into(), face_bold, face_italic)
    }

    pub fn register_shared_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        data: Arc<[u8]>,
        face_bold: bool,
        face_italic: bool,
    ) -> Result<(), String> {
        if data.len() > MAX_FONT_BYTES {
            return Err(format!("font exceeds {MAX_FONT_BYTES} bytes"));
        }
        let key = (family.trim().to_lowercase(), bold, italic);
        if self.faces.get(&key).is_some_and(|chain| {
            chain.iter().any(|face| {
                (face.bold, face.italic) == (face_bold, face_italic)
                    && (Arc::ptr_eq(&face.data, &data) || face.data == data)
            })
        }) {
            return Ok(());
        }
        if self.registrations >= MAX_FONT_REGISTRATIONS {
            return Err(format!(
                "more than {MAX_FONT_REGISTRATIONS} font registrations"
            ));
        }
        let data = if let Some(known) = self
            .data
            .iter()
            .find(|known| Arc::ptr_eq(known, &data) || known.as_ref() == data.as_ref())
        {
            Arc::clone(known)
        } else {
            if self.data.len() >= MAX_FONT_FACES {
                return Err(format!("more than {MAX_FONT_FACES} font faces"));
            }
            if rustybuzz::Face::from_slice(&data, 0).is_none() {
                return Err("invalid raster font".into());
            }
            self.data.push(Arc::clone(&data));
            data
        };
        self.faces.entry(key).or_default().push(font::FontFace {
            data,
            bold: face_bold,
            italic: face_italic,
        });
        self.registrations += 1;
        Ok(())
    }

    pub fn face_count(&self) -> usize {
        self.data.len()
    }

    pub fn registration_count(&self) -> usize {
        self.registrations
    }

    #[doc(hidden)]
    pub fn shared_font_data(&self, data: &Arc<[u8]>) -> Option<Arc<[u8]>> {
        self.data
            .iter()
            .find(|known| Arc::ptr_eq(known, data) || known.as_ref() == data.as_ref())
            .map(Arc::clone)
    }

    pub fn font_data(&self, family: &str, bold: bool, italic: bool) -> Option<&[u8]> {
        self.chain(family, bold, italic)
            .first()
            .map(|face| face.data.as_ref())
    }

    fn chain(&self, family: &str, bold: bool, italic: bool) -> &[font::FontFace] {
        self.faces
            .get(&(family.trim().to_lowercase(), bold, italic))
            .map_or(&[], Vec::as_slice)
    }
}

pub fn render_png(dl: &DisplayList) -> Result<Vec<u8>, String> {
    render_png_with_fonts(dl, &RasterFonts::default())
}

pub fn render_png_with_fonts(dl: &DisplayList, fonts: &RasterFonts) -> Result<Vec<u8>, String> {
    let w = (dl.width.ceil() as u32).max(1);
    let h = (dl.height.ceil() as u32).max(1);
    let mut pixmap = Pixmap::new(w, h).ok_or_else(|| "invalid pixmap size".to_string())?;
    pixmap.fill(Color::WHITE);
    let mut clip_surface = ClipSurface::default();
    let mut patterns = HashMap::new();

    for cmd in &dl.commands {
        match cmd {
            DrawCmd::PatternFill {
                x,
                y,
                w,
                h,
                pattern,
                fg,
                bg,
                origin_x,
                origin_y,
                clip,
            } => {
                let tile = match patterns.entry((*pattern, fg.clone(), bg.clone())) {
                    Entry::Occupied(entry) => entry.into_mut(),
                    Entry::Vacant(entry) => {
                        let mut tile =
                            Pixmap::new(8, 8).ok_or_else(|| "invalid pattern tile".to_string())?;
                        tile.fill(parse_color(bg)?);
                        let foreground = parse_color(fg)?.to_color_u8().premultiply();
                        for (y, row) in pattern.mask().into_iter().enumerate() {
                            for x in 0..8 {
                                if row & (1 << x) != 0 {
                                    tile.pixels_mut()[y * 8 + x] = foreground;
                                }
                            }
                        }
                        entry.insert(tile)
                    }
                };
                let Some(rect) = Rect::from_xywh(
                    pixel_round(*x),
                    pixel_round(*y),
                    pixel_round(x + w) - pixel_round(*x),
                    pixel_round(y + h) - pixel_round(*y),
                ) else {
                    continue;
                };
                let paint = Paint {
                    shader: tiny_skia::Pattern::new(
                        tile.as_ref(),
                        tiny_skia::SpreadMode::Repeat,
                        tiny_skia::FilterQuality::Nearest,
                        1.0,
                        Transform::from_translate(pixel_round(*origin_x), pixel_round(*origin_y)),
                    ),
                    anti_alias: false,
                    ..Paint::default()
                };
                if let Some(clip) = clip {
                    let bounds = reach(rect.left(), rect.top(), rect.right(), rect.bottom(), 1.0);
                    clip_surface.paint(&mut pixmap, *clip, bounds, |target, transform, mask| {
                        target.fill_rect(rect, &paint, transform, mask);
                        Ok(())
                    })?;
                } else {
                    pixmap.fill_rect(rect, &paint, Transform::identity(), None);
                }
            }
            DrawCmd::FillRect {
                x,
                y,
                w,
                h,
                color,
                clip,
            } => {
                let rect = Rect::from_xywh(*x, *y, *w, *h)
                    .ok_or_else(|| "invalid fill rectangle".to_string())?;
                let mut paint = Paint::default();
                paint.set_color(parse_color(color)?);
                paint.anti_alias = true;
                if let Some(clip) = clip {
                    let bounds = reach(rect.left(), rect.top(), rect.right(), rect.bottom(), 1.0);
                    clip_surface.paint(&mut pixmap, *clip, bounds, |target, transform, mask| {
                        target.fill_rect(rect, &paint, transform, mask);
                        Ok(())
                    })?;
                } else {
                    pixmap.fill_rect(rect, &paint, Transform::identity(), None);
                }
            }
            DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color,
                style,
                rail,
                clip,
            } => {
                let paint =
                    |target: &mut Pixmap, transform: Transform, mask: Option<&Mask>| match rail {
                        Some(rail) if style.as_deref() == Some("double") => paint_double(
                            target, *x1, *y1, *x2, *y2, *width, *rail, color, transform, mask,
                        ),
                        _ => paint_line(
                            target, *x1, *y1, *x2, *y2, *width, color, style, transform, mask,
                        ),
                    };
                if let Some(clip) = clip {
                    let pen = width.abs().max(rail.map_or(0.0, f32::abs)).max(1.0);
                    let bounds = reach(*x1, *y1, *x2, *y2, 2.0 * pen + 2.0);
                    clip_surface.paint(&mut pixmap, *clip, bounds, paint)?;
                } else {
                    paint(&mut pixmap, Transform::identity(), None)?;
                }
            }
            DrawCmd::Path {
                commands,
                fill,
                stroke,
                clip,
            } => {
                let path = build_path(commands)?;
                let mut paint = Paint::default();
                paint.set_color(parse_color(fill)?);
                paint.anti_alias = true;
                let stroke = if let Some(stroke) = stroke {
                    let mut paint = Paint::default();
                    paint.set_color(parse_color(&stroke.color)?);
                    paint.anti_alias = true;
                    Some((paint, stroke.width))
                } else {
                    None
                };
                let paint_path =
                    |target: &mut Pixmap, transform: Transform, mask: Option<&Mask>| {
                        target.fill_path(&path, &paint, FillRule::Winding, transform, mask);
                        if let Some((stroke_paint, width)) = &stroke {
                            target.stroke_path(
                                &path,
                                stroke_paint,
                                &Stroke {
                                    width: *width,
                                    ..Stroke::default()
                                },
                                transform,
                                mask,
                            );
                        }
                        Ok(())
                    };
                if let Some(clip) = clip {
                    let pen = stroke.as_ref().map_or(0.0, |(_, width)| width.abs());
                    let area = path.bounds();
                    let bounds = reach(
                        area.left(),
                        area.top(),
                        area.right(),
                        area.bottom(),
                        2.0 * pen + 2.0,
                    );
                    clip_surface.paint(&mut pixmap, *clip, bounds, paint_path)?;
                } else {
                    pixmap.fill_path(
                        &path,
                        &paint,
                        FillRule::Winding,
                        Transform::identity(),
                        None,
                    );
                    if let Some((stroke_paint, width)) = &stroke {
                        pixmap.stroke_path(
                            &path,
                            stroke_paint,
                            &Stroke {
                                width: *width,
                                ..Stroke::default()
                            },
                            Transform::identity(),
                            None,
                        );
                    }
                }
            }
            DrawCmd::Text {
                x,
                y,
                text,
                font_size,
                color,
                clip,
                align,
                bold,
                italic,
                underline,
                strike,
                highlight,
                dashed_underline,
                font_family,
                number_format,
                ghost: _,
                chart: _,
                rotation_deg,
                rotation_box,
                rotation_valign,
                stacked,
                wrap,
            } => {
                let run = font::TextRun {
                    x: *x,
                    y: *y,
                    text,
                    font_size_pt: *font_size,
                    color: parse_color(color)?,
                    align: *align,
                    clip,
                    bold: *bold,
                    italic: *italic,
                    underline: *underline,
                    strike: *strike,
                    highlight: highlight.as_deref().map(parse_color).transpose()?,
                    dashed_underline: *dashed_underline,
                    faces: fonts.chain(font_family.as_deref().unwrap_or("Calibri"), *bold, *italic),
                };
                if *stacked {
                    let cell = rotation_box.as_ref().unwrap_or(clip);
                    font::paint_stacked_text(
                        &mut pixmap,
                        &run,
                        cell,
                        (*rotation_valign).unwrap_or(xlsx_render::TextValign::Bottom),
                    );
                } else if let (Some(deg), Some(wrap)) = (rotation_deg, wrap) {
                    let cell = rotation_box.as_ref().unwrap_or(clip);
                    let lines = xlsx_render::wrap::wrap_lines(text, wrap.width, |line| {
                        font::measure_text_with_fonts(run.faces, line, *font_size)
                    });
                    for (index, line) in lines.iter().enumerate() {
                        let strip = xlsx_render::wrap::rotated_line_box(
                            cell,
                            *deg,
                            index,
                            lines.len(),
                            wrap,
                            *align,
                        );
                        if strip.x + strip.w <= clip.x || strip.x >= clip.x + clip.w {
                            continue;
                        }
                        font::paint_rotated_text(
                            &mut pixmap,
                            &font::TextRun {
                                text: line,
                                align: Align::Center,
                                ..run
                            },
                            *deg,
                            &strip,
                            (*rotation_valign).unwrap_or(xlsx_render::TextValign::Bottom),
                            None,
                        );
                    }
                } else if let Some(deg) = rotation_deg {
                    let cell = rotation_box.as_ref().unwrap_or(clip);
                    font::paint_rotated_text(
                        &mut pixmap,
                        &run,
                        *deg,
                        cell,
                        (*rotation_valign).unwrap_or(xlsx_render::TextValign::Bottom),
                        number_format.as_ref(),
                    );
                } else if let Some(layout) = number_format {
                    xlsx_render::paint_number_format(
                        text,
                        layout,
                        *align,
                        |literal| font::measure_text_with_fonts(run.faces, literal, *font_size),
                        |literal, x| {
                            font::paint_text(
                                &mut pixmap,
                                &font::TextRun {
                                    x,
                                    text: literal,
                                    align: Align::Left,
                                    ..run
                                },
                            )
                        },
                    );
                } else if let Some(wrap) = wrap {
                    let lines = xlsx_render::wrap::wrap_lines(text, wrap.width, |line| {
                        font::measure_text_with_fonts(run.faces, line, *font_size)
                    });
                    for (index, line) in lines.iter().enumerate() {
                        font::paint_text(
                            &mut pixmap,
                            &font::TextRun {
                                y: xlsx_render::wrap::wrapped_line_y(*y, index, lines.len(), wrap),
                                text: line,
                                ..run
                            },
                        );
                    }
                } else {
                    font::paint_text(&mut pixmap, &run);
                }
            }
        }
    }

    let pixels = pixmap.take_demultiplied();
    let mut data = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut data, w, h);
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
