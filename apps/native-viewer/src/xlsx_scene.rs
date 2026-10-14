use std::collections::HashMap;
use std::sync::{Arc, LazyLock};

use anyhow::{Result, bail};
use rustybuzz::{Face, GlyphBuffer, UnicodeBuffer};
use vello::kurbo::{Affine, BezPath, Line, Rect, Stroke};
use vello::peniko::{
    Blob, Color, Extend, Fill, FontData, ImageAlphaType, ImageBrush, ImageData, ImageFormat,
    ImageQuality,
};
use vello::{Glyph, Scene};
use xlsx_render::{
    Align, DisplayList, DrawCmd, GeometryPathCommand, PatternType, Rect as DisplayRect, border_dash,
};

use crate::scene_shared::{
    PageScene, SkipStats, draw_placeholder, strict_color, with_clip_layers, with_dashes,
};

const BOLD_OFFSET_PX: f32 = 0.35;
const ITALIC_SHEAR: f64 = 0.21;
const PX_PER_PT: f32 = 96.0 / 72.0;
const FONT_BYTES: &[u8] = include_bytes!("../../../crates/xlsx-raster/assets/Carlito-Regular.ttf");

static FACE: LazyLock<Face<'static>> =
    LazyLock::new(|| Face::from_slice(FONT_BYTES, 0).expect("embedded Carlito is a valid font"));

pub fn translate_sheet(display_list: &DisplayList) -> Result<PageScene> {
    let width = frame_dimension(display_list.width, "width")?;
    let height = frame_dimension(display_list.height, "height")?;
    let font = FontData::new(Blob::from(FONT_BYTES.to_vec()), 0);
    let mut scene = Scene::new();
    scene.fill(
        Fill::NonZero,
        Affine::IDENTITY,
        Color::WHITE,
        None,
        &Rect::new(0.0, 0.0, width, height),
    );
    let mut skipped = SkipStats::default();
    let mut tiles = PatternTiles::default();
    for command in &display_list.commands {
        translate_or_placeholder(&mut scene, command, &font, &mut tiles, &mut skipped);
    }
    Ok(PageScene {
        background: Scene::new(),
        scene,
        width,
        height,
        skipped,
    })
}

pub fn paint_cell_editor(scene: &mut Scene, rect: DisplayRect, value: &str) -> Result<()> {
    let bounds =
        rectangle(rect.x, rect.y, rect.w, rect.h, "cell editor").map_err(anyhow::Error::msg)?;
    scene.fill(Fill::NonZero, Affine::IDENTITY, Color::WHITE, None, &bounds);
    let font = FontData::new(Blob::from(FONT_BYTES.to_vec()), 0);
    let text = PreparedText::new(
        rect.x + 2.0,
        rect.y + (rect.h + 11.0 * 0.7) / 2.0,
        value,
        11.0,
        "#000000",
        Align::Left,
        false,
        false,
        false,
        false,
        None,
        false,
    )
    .map_err(anyhow::Error::msg)?;
    paint_clipped(scene, Some(&rect), |scene| text.paint(scene, &font)).map_err(anyhow::Error::msg)
}

fn translate_or_placeholder(
    scene: &mut Scene,
    command: &DrawCmd,
    font: &FontData,
    tiles: &mut PatternTiles,
    skipped: &mut SkipStats,
) {
    if let Err(reason) = translate_command(scene, command, font, tiles) {
        skipped.record(command_kind(command), reason);
        draw_placeholder(scene, command_bounds(command));
    }
}

/// One 8x8 tile per pattern and color pair, shared by every cell that uses it.
#[derive(Default)]
struct PatternTiles(HashMap<(PatternType, Arc<str>, Arc<str>), ImageData>);

impl PatternTiles {
    fn tile(
        &mut self,
        pattern: PatternType,
        fg: &Arc<str>,
        bg: &Arc<str>,
    ) -> Result<ImageData, String> {
        let key = (pattern, fg.clone(), bg.clone());
        if let Some(tile) = self.0.get(&key) {
            return Ok(tile.clone());
        }
        let fg = strict_color(fg)?.to_rgba8();
        let bg = strict_color(bg)?.to_rgba8();
        let mut pixels = Vec::with_capacity(8 * 8 * 4);
        for row in pattern.mask() {
            for x in 0..8 {
                let color = if row & (1 << x) != 0 { fg } else { bg };
                pixels.extend([color.r, color.g, color.b, color.a]);
            }
        }
        let tile = ImageData {
            data: Blob::from(pixels),
            format: ImageFormat::Rgba8,
            alpha_type: ImageAlphaType::Alpha,
            width: 8,
            height: 8,
        };
        self.0.insert(key, tile.clone());
        Ok(tile)
    }
}

fn translate_command(
    scene: &mut Scene,
    command: &DrawCmd,
    font: &FontData,
    tiles: &mut PatternTiles,
) -> Result<(), String> {
    match command {
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
            let rect = rectangle(*x, *y, *w, *h, "pattern fill")?;
            let origin = Affine::translate((
                f64::from(finite(*origin_x, "pattern origin x")?),
                f64::from(finite(*origin_y, "pattern origin y")?),
            ));
            let brush = ImageBrush::new(tiles.tile(*pattern, fg, bg)?)
                .with_extend(Extend::Repeat)
                .with_quality(ImageQuality::Low);
            paint_clipped(scene, clip.as_ref(), |scene| {
                scene.fill(Fill::NonZero, Affine::IDENTITY, &brush, Some(origin), &rect);
            })
        }
        DrawCmd::FillRect {
            x,
            y,
            w,
            h,
            color,
            clip,
        } => {
            let rect = rectangle(*x, *y, *w, *h, "fill rectangle")?;
            let color = strict_color(color)?;
            paint_clipped(scene, clip.as_ref(), |scene| {
                scene.fill(Fill::NonZero, Affine::IDENTITY, color, None, &rect);
            })
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
            let line = PreparedLine::new(*x1, *y1, *x2, *y2, *width, color, style, *rail)?;
            paint_clipped(scene, clip.as_ref(), |scene| line.paint(scene))
        }
        DrawCmd::Path {
            commands,
            fill,
            stroke,
            clip,
        } => {
            let path = build_path(commands)?;
            let fill = strict_color(fill)?;
            let stroke = stroke
                .as_ref()
                .map(|stroke| -> Result<_, String> {
                    Ok((
                        Stroke::new(nonnegative(stroke.width, "path stroke width")?),
                        strict_color(&stroke.color)?,
                    ))
                })
                .transpose()?;
            paint_clipped(scene, clip.as_ref(), |scene| {
                scene.fill(Fill::NonZero, Affine::IDENTITY, fill, None, &path);
                if let Some((stroke, color)) = &stroke {
                    scene.stroke(stroke, Affine::IDENTITY, *color, None, &path);
                }
            })
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
            font_family: _,
            ghost: _,
            chart: _,
        } => {
            let text = PreparedText::new(
                *x,
                *y,
                text,
                *font_size,
                color,
                *align,
                *bold,
                *italic,
                *underline,
                *strike,
                highlight.as_deref(),
                *dashed_underline,
            )?;
            paint_clipped(scene, Some(clip), |scene| text.paint(scene, font))
        }
    }
}

fn paint_clipped(
    scene: &mut Scene,
    clip: Option<&DisplayRect>,
    paint: impl FnOnce(&mut Scene),
) -> Result<(), String> {
    let Some(clip) = clip else {
        paint(scene);
        return Ok(());
    };
    let Some(rect) = clip_rectangle(*clip)? else {
        return Ok(());
    };
    with_clip_layers(scene, &[(Affine::IDENTITY, rect)], |scene| {
        paint(scene);
        Ok(())
    })
}

struct PreparedLine {
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    width: f32,
    color: Color,
    style: Option<String>,
    rail: Option<f32>,
}

impl PreparedLine {
    #[allow(clippy::too_many_arguments)]
    fn new(
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        width: f32,
        color: &str,
        style: &Option<String>,
        rail: Option<f32>,
    ) -> Result<Self, String> {
        if !matches!(style.as_deref(), None | Some("solid" | "double"))
            && border_dash(style.as_deref().unwrap_or_default()).is_empty()
        {
            return Err(format!(
                "line style {} is not translated",
                style.as_deref().unwrap_or_default()
            ));
        }
        Ok(Self {
            x1: finite(x1, "line x1")?,
            y1: finite(y1, "line y1")?,
            x2: finite(x2, "line x2")?,
            y2: finite(y2, "line y2")?,
            width: nonnegative_f32(width, "line width")?,
            color: strict_color(color)?,
            style: style.clone(),
            rail: rail
                .map(|rail| nonnegative_f32(rail, "line rail"))
                .transpose()?,
        })
    }

    fn paint(&self, scene: &mut Scene) {
        match self.style.as_deref() {
            Some("double") => {
                let (offset, width) = match self.rail {
                    Some(rail) => ((self.width - rail) / 2.0, rail),
                    None => ((self.width * 0.8).max(0.8), (self.width * 0.6).max(0.5)),
                };
                let horizontal = (self.y1 - self.y2).abs() <= (self.x1 - self.x2).abs();
                let (dx, dy) = if horizontal {
                    (0.0, offset)
                } else {
                    (offset, 0.0)
                };
                stroke_segment(
                    scene,
                    self.x1 - dx,
                    self.y1 - dy,
                    self.x2 - dx,
                    self.y2 - dy,
                    Stroke::new(f64::from(width)),
                    self.color,
                );
                stroke_segment(
                    scene,
                    self.x1 + dx,
                    self.y1 + dy,
                    self.x2 + dx,
                    self.y2 + dy,
                    Stroke::new(f64::from(width)),
                    self.color,
                );
            }
            Some(style) if !border_dash(style).is_empty() => {
                let scale = if style.starts_with("medium") {
                    self.width / 2.0
                } else {
                    self.width
                };
                stroke_segment(
                    scene,
                    self.x1,
                    self.y1,
                    self.x2,
                    self.y2,
                    with_dashes(
                        Stroke::new(f64::from(self.width)),
                        border_dash(style).iter().map(|run| f64::from(run * scale)),
                    ),
                    self.color,
                );
            }
            None | Some("solid") => stroke_segment(
                scene,
                self.x1,
                self.y1,
                self.x2,
                self.y2,
                Stroke::new(f64::from(self.width)),
                self.color,
            ),
            Some(_) => unreachable!(),
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn stroke_segment(
    scene: &mut Scene,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    stroke: Stroke,
    color: Color,
) {
    scene.stroke(
        &stroke,
        Affine::IDENTITY,
        color,
        None,
        &Line::new(
            (f64::from(x1), f64::from(y1)),
            (f64::from(x2), f64::from(y2)),
        ),
    );
}

struct PreparedText {
    baseline: f32,
    pen: f32,
    font_size_px: f32,
    scale: f32,
    run_width: f32,
    glyphs: Vec<Glyph>,
    color: Color,
    highlight: Option<Color>,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    dashed_underline: bool,
}

impl PreparedText {
    #[allow(clippy::too_many_arguments)]
    fn new(
        x: f32,
        y: f32,
        text: &str,
        font_size_pt: f32,
        color: &str,
        align: Align,
        bold: bool,
        italic: bool,
        underline: bool,
        strike: bool,
        highlight: Option<&str>,
        dashed_underline: bool,
    ) -> Result<Self, String> {
        let x = finite(x, "text x")?;
        let baseline = finite(y, "text y")?;
        let font_size_pt = positive_f32(font_size_pt, "text font size")?;
        let font_size_px = font_size_pt * PX_PER_PT;
        let scale = font_size_px / FACE.units_per_em() as f32;
        let shaped = shape(text);
        let total: i32 = shaped
            .glyph_positions()
            .iter()
            .map(|position| position.x_advance)
            .sum();
        let run_width = total as f32 * scale;
        let pen = match align {
            Align::Left => x,
            Align::Center => x - run_width / 2.0,
            Align::Right => x - run_width,
        };
        let mut glyph_pen = 0.0;
        let glyphs = shaped
            .glyph_infos()
            .iter()
            .zip(shaped.glyph_positions())
            .map(|(info, position)| {
                let glyph = Glyph {
                    id: info.glyph_id,
                    x: glyph_pen + position.x_offset as f32 * scale,
                    y: -(position.y_offset as f32) * scale,
                };
                glyph_pen += position.x_advance as f32 * scale;
                glyph
            })
            .collect();
        Ok(Self {
            baseline,
            pen,
            font_size_px,
            scale,
            run_width,
            glyphs,
            color: strict_color(color)?,
            highlight: highlight.map(strict_color).transpose()?,
            bold,
            italic,
            underline,
            strike,
            dashed_underline,
        })
    }

    fn paint(&self, scene: &mut Scene, font: &FontData) {
        self.paint_highlight(scene);
        self.paint_glyphs(scene, font, self.pen);
        if self.bold {
            self.paint_glyphs(scene, font, self.pen + BOLD_OFFSET_PX);
        }
        self.paint_decorations(scene);
    }

    fn paint_highlight(&self, scene: &mut Scene) {
        let Some(color) = self.highlight else {
            return;
        };
        let ascent = FACE.ascender() as f32 * self.scale;
        let descent = -FACE.descender() as f32 * self.scale;
        let rect = Rect::new(
            f64::from(self.pen - 2.0),
            f64::from(self.baseline - ascent - 1.0),
            f64::from(self.pen + self.run_width + 2.0),
            f64::from(self.baseline + descent + 1.0),
        );
        scene.fill(Fill::NonZero, Affine::IDENTITY, color, None, &rect);
    }

    fn paint_glyphs(&self, scene: &mut Scene, font: &FontData, pen: f32) {
        let glyph_transform = self.italic.then_some(Affine::skew(ITALIC_SHEAR, 0.0));
        scene
            .draw_glyphs(font)
            .font_size(self.font_size_px)
            .brush(self.color)
            .glyph_transform(glyph_transform)
            .transform(Affine::translate((
                f64::from(pen),
                f64::from(self.baseline),
            )))
            .draw(Fill::NonZero, self.glyphs.iter().copied());
    }

    fn paint_decorations(&self, scene: &mut Scene) {
        if self.run_width <= 0.0 {
            return;
        }
        let em = FACE.units_per_em() as f32;
        if self.underline {
            let metrics = FACE.underline_metrics();
            let position = metrics
                .map(|metrics| metrics.position as f32)
                .unwrap_or(-0.1 * em);
            let thickness = metrics
                .map(|metrics| metrics.thickness as f32)
                .unwrap_or(0.05 * em);
            self.paint_bar(scene, position, thickness);
        }
        if self.strike {
            let metrics = FACE.strikeout_metrics();
            let position = metrics
                .map(|metrics| metrics.position as f32)
                .unwrap_or(0.26 * em);
            let thickness = metrics
                .map(|metrics| metrics.thickness as f32)
                .unwrap_or(0.05 * em);
            self.paint_bar(scene, position, thickness);
        }
        if self.dashed_underline {
            let metrics = FACE.underline_metrics();
            let position = metrics
                .map(|metrics| metrics.position as f32)
                .unwrap_or(-0.1 * em);
            let thickness = metrics
                .map(|metrics| metrics.thickness as f32)
                .unwrap_or(0.05 * em);
            let height = (thickness * self.scale).max(0.5);
            let y = self.baseline - position * self.scale - height / 2.0;
            let dash = 3.0_f32.max(height * 2.0);
            let gap = 2.0_f32.max(height);
            let mut x = self.pen;
            let end = x + self.run_width;
            while x < end {
                let right = x + dash.min(end - x);
                let rect = Rect::new(
                    f64::from(x),
                    f64::from(y),
                    f64::from(right),
                    f64::from(y + height),
                );
                scene.fill(Fill::NonZero, Affine::IDENTITY, self.color, None, &rect);
                x += dash + gap;
            }
        }
    }

    fn paint_bar(&self, scene: &mut Scene, center_up: f32, thickness: f32) {
        let height = (thickness * self.scale).max(0.5);
        let center = self.baseline - center_up * self.scale;
        let rect = Rect::new(
            f64::from(self.pen),
            f64::from(center - height / 2.0),
            f64::from(self.pen + self.run_width),
            f64::from(center + height / 2.0),
        );
        scene.fill(Fill::NonZero, Affine::IDENTITY, self.color, None, &rect);
    }
}

fn shape(text: &str) -> GlyphBuffer {
    let mut buffer = UnicodeBuffer::new();
    buffer.push_str(text);
    rustybuzz::shape(&FACE, &[], buffer)
}

fn build_path(commands: &[GeometryPathCommand]) -> Result<BezPath, String> {
    if commands.is_empty() {
        return Err("path has no drawable commands".to_owned());
    }
    let mut path = BezPath::new();
    for command in commands {
        match command {
            GeometryPathCommand::Move { x, y } => {
                path.move_to((path_coord(*x)?, path_coord(*y)?));
            }
            GeometryPathCommand::Line { x, y } => {
                path.line_to((path_coord(*x)?, path_coord(*y)?));
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                path.quad_to(
                    (path_coord(*cpx)?, path_coord(*cpy)?),
                    (path_coord(*x)?, path_coord(*y)?),
                );
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                path.curve_to(
                    (path_coord(*cp1x)?, path_coord(*cp1y)?),
                    (path_coord(*cp2x)?, path_coord(*cp2y)?),
                    (path_coord(*x)?, path_coord(*y)?),
                );
            }
            GeometryPathCommand::Close => path.close_path(),
        }
    }
    Ok(path)
}

fn path_coord(value: f64) -> Result<f64, String> {
    let value = value as f32;
    finite(value, "path coordinate").map(f64::from)
}

fn rectangle(x: f32, y: f32, w: f32, h: f32, name: &str) -> Result<Rect, String> {
    let x = finite(x, name)?;
    let y = finite(y, name)?;
    let w = positive_f32(w, name)?;
    let h = positive_f32(h, name)?;
    let right = finite(x + w, name)?;
    let bottom = finite(y + h, name)?;
    Ok(Rect::new(
        f64::from(x),
        f64::from(y),
        f64::from(right),
        f64::from(bottom),
    ))
}

fn clip_rectangle(rect: DisplayRect) -> Result<Option<Rect>, String> {
    let x = finite(rect.x, "clip x")?;
    let y = finite(rect.y, "clip y")?;
    let w = finite(rect.w, "clip width")?;
    let h = finite(rect.h, "clip height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(None);
    }
    let right = finite(x + w, "clip right")?;
    let bottom = finite(y + h, "clip bottom")?;
    Ok(Some(Rect::new(
        f64::from(x),
        f64::from(y),
        f64::from(right),
        f64::from(bottom),
    )))
}

fn frame_dimension(value: f32, name: &str) -> Result<f64> {
    if !value.is_finite() || value <= 0.0 {
        bail!("XLSX display-list {name} is invalid");
    }
    Ok(f64::from(value))
}

fn finite(value: f32, name: &str) -> Result<f32, String> {
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| format!("{name} is not finite"))
}

fn positive_f32(value: f32, name: &str) -> Result<f32, String> {
    let value = finite(value, name)?;
    (value > 0.0)
        .then_some(value)
        .ok_or_else(|| format!("{name} is not positive"))
}

fn nonnegative_f32(value: f32, name: &str) -> Result<f32, String> {
    let value = finite(value, name)?;
    (value >= 0.0)
        .then_some(value)
        .ok_or_else(|| format!("{name} is negative"))
}

fn nonnegative(value: f32, name: &str) -> Result<f64, String> {
    nonnegative_f32(value, name).map(f64::from)
}

fn command_kind(command: &DrawCmd) -> &'static str {
    match command {
        DrawCmd::FillRect { .. } => "fillRect",
        DrawCmd::PatternFill { .. } => "patternFill",
        DrawCmd::Line { .. } => "line",
        DrawCmd::Path { .. } => "path",
        DrawCmd::Text { .. } => "text",
    }
}

fn command_bounds(command: &DrawCmd) -> Option<Rect> {
    match command {
        DrawCmd::FillRect { x, y, w, h, .. } | DrawCmd::PatternFill { x, y, w, h, .. } => {
            loose_rectangle(*x, *y, *w, *h)
        }
        DrawCmd::Line { x1, y1, x2, y2, .. } => loose_extents(*x1, *y1, *x2, *y2),
        DrawCmd::Path { commands, .. } => path_bounds(commands),
        DrawCmd::Text { clip, .. } => loose_rectangle(clip.x, clip.y, clip.w, clip.h),
    }
}

fn path_bounds(commands: &[GeometryPathCommand]) -> Option<Rect> {
    let mut points = Vec::new();
    for command in commands {
        match command {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                points.push((*x, *y));
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                points.extend([(*cpx, *cpy), (*x, *y)]);
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                points.extend([(*cp1x, *cp1y), (*cp2x, *cp2y), (*x, *y)]);
            }
            GeometryPathCommand::Close => {}
        }
    }
    let mut points = points
        .into_iter()
        .filter(|(x, y)| x.is_finite() && y.is_finite());
    let (first_x, first_y) = points.next()?;
    let (mut left, mut top, mut right, mut bottom) = (first_x, first_y, first_x, first_y);
    for (x, y) in points {
        left = left.min(x);
        top = top.min(y);
        right = right.max(x);
        bottom = bottom.max(y);
    }
    Some(Rect::new(left, top, right, bottom))
}

fn loose_rectangle(x: f32, y: f32, w: f32, h: f32) -> Option<Rect> {
    let (right, bottom) = (x + w, y + h);
    if [x, y, right, bottom].iter().any(|value| !value.is_finite()) {
        return None;
    }
    Some(Rect::new(
        f64::from(x.min(right)),
        f64::from(y.min(bottom)),
        f64::from(x.max(right)),
        f64::from(y.max(bottom)),
    ))
}

fn loose_extents(x1: f32, y1: f32, x2: f32, y2: f32) -> Option<Rect> {
    if [x1, y1, x2, y2].iter().any(|value| !value.is_finite()) {
        return None;
    }
    Some(Rect::new(
        f64::from(x1.min(x2)),
        f64::from(y1.min(y2)),
        f64::from(x1.max(x2)),
        f64::from(y1.max(y2)),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_render::{GridMeta, PathStroke};

    fn display_list(commands: Vec<DrawCmd>) -> DisplayList {
        DisplayList {
            width: 100.0,
            height: 60.0,
            commands,
            grid: GridMeta::default(),
            hyperlinks: Vec::new(),
            charts: Vec::new(),
        }
    }

    #[test]
    fn translates_every_command_and_line_style() {
        let clip = DisplayRect {
            x: 0.0,
            y: 0.0,
            w: 100.0,
            h: 60.0,
        };
        let mut commands = vec![
            DrawCmd::FillRect {
                x: 0.0,
                y: 0.0,
                w: 100.0,
                h: 60.0,
                color: "#ffffff".into(),
                clip: Some(clip),
            },
            DrawCmd::PatternFill {
                x: 0.0,
                y: 0.0,
                w: 100.0,
                h: 20.0,
                pattern: PatternType::LightUp,
                fg: "#123456".into(),
                bg: "#fedcba".into(),
                origin_x: -3.0,
                origin_y: -2.0,
                clip: Some(clip),
            },
        ];
        for (index, style) in [
            None,
            Some("dashed"),
            Some("dotted"),
            Some("double"),
            Some("mediumDashed"),
            Some("dashDot"),
            Some("mediumDashDot"),
            Some("dashDotDot"),
            Some("mediumDashDotDot"),
            Some("slantDashDot"),
        ]
        .into_iter()
        .enumerate()
        {
            commands.push(DrawCmd::Line {
                x1: 0.0,
                y1: 4.0 + index as f32 * 4.0,
                x2: 100.0,
                y2: 4.0 + index as f32 * 4.0,
                width: 1.0,
                color: "#123456".into(),
                style: style.map(str::to_owned),
                rail: None,
                clip: Some(clip),
            });
        }
        commands.push(DrawCmd::Path {
            commands: vec![
                GeometryPathCommand::Move { x: 2.0, y: 20.0 },
                GeometryPathCommand::Line { x: 20.0, y: 20.0 },
                GeometryPathCommand::Line { x: 20.0, y: 38.0 },
                GeometryPathCommand::Close,
            ],
            fill: "#abcdef80".into(),
            stroke: Some(PathStroke {
                color: "#102030".into(),
                width: 1.0,
            }),
            clip: Some(clip),
        });
        commands.push(DrawCmd::Text {
            x: 50.0,
            y: 48.0,
            text: "Styled".into(),
            font_size: 11.0,
            color: "#000000".into(),
            clip,
            align: Align::Center,
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            highlight: Some("#ffff0080".into()),
            dashed_underline: true,
            font_family: Some("Ignored like xlsx-raster".into()),
            ghost: true,
            chart: true,
        });
        let page = translate_sheet(&display_list(commands)).unwrap();
        assert_eq!(page.skipped.total(), 0);
    }

    #[test]
    fn malformed_color_becomes_a_counted_placeholder() {
        let page = translate_sheet(&display_list(vec![DrawCmd::FillRect {
            x: 2.0,
            y: 2.0,
            w: 20.0,
            h: 10.0,
            color: "red".into(),
            clip: None,
        }]))
        .unwrap();
        assert_eq!(page.skipped.total(), 1);
        assert_eq!(page.skipped.counts.get("fillRect"), Some(&1));
        assert!(
            page.skipped
                .reasons
                .keys()
                .any(|reason| reason.contains("invalid XLSX color"))
        );
    }

    #[test]
    fn unknown_line_style_becomes_a_counted_placeholder() {
        let page = translate_sheet(&display_list(vec![DrawCmd::Line {
            x1: 2.0,
            y1: 2.0,
            x2: 20.0,
            y2: 2.0,
            width: 1.0,
            color: "#000000".into(),
            style: Some("future-style".into()),
            rail: None,
            clip: None,
        }]))
        .unwrap();
        assert_eq!(page.skipped.total(), 1);
        assert_eq!(page.skipped.counts.get("line"), Some(&1));
        assert!(
            page.skipped
                .reasons
                .keys()
                .any(|reason| reason.contains("future-style"))
        );
    }
}
