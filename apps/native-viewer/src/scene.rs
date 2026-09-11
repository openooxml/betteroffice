use anyhow::{Context, Result};
use docx_layout::display_list::{
    DisplayBorderStyle, DisplayList, DisplayPage, DocAttrs, GlyphRunPrimitive, ImagePrimitive,
    LinePrimitive, PageBorderPrimitive, PageBorderSide, PageBorderZOrder, Primitive,
    ShapePathCommand, ShapePrimitive, TextRunPrimitive,
};
use docx_raster::{ImageScope, NoteKind};
use serde_json::Number;
use vello::kurbo::{Affine, BezPath, Line, Rect, Stroke};
use vello::peniko::{Color, Fill, ImageBrush};
use vello::{Glyph, Scene};

use crate::fonts::FontRegistry;
use crate::images::ImageRegistry;
use crate::scene_shared::{
    PageScene, SkipStats, color, draw_placeholder, with_clip_layers, with_dashes,
};

pub fn translate_document(
    display_list: &DisplayList,
    fonts: &FontRegistry,
    images: &ImageRegistry,
) -> Result<Vec<PageScene>> {
    display_list
        .pages
        .iter()
        .map(|page| translate_page(page, fonts, images))
        .collect()
}

pub fn translate_page(
    page: &DisplayPage,
    fonts: &FontRegistry,
    images: &ImageRegistry,
) -> Result<PageScene> {
    let width = number(&page.width).map_err(anyhow::Error::msg)?;
    let height = number(&page.height).map_err(anyhow::Error::msg)?;
    let mut background_scene = Scene::new();
    let background = page.background.as_deref().unwrap_or("#ffffff");
    background_scene.fill(
        Fill::NonZero,
        Affine::IDENTITY,
        color(background, 1.0).map_err(anyhow::Error::msg)?,
        None,
        &Rect::new(0.0, 0.0, width, height),
    );
    let mut scene = Scene::new();
    let mut skipped = SkipStats::default();
    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order == Some(PageBorderZOrder::Back))
    {
        translate_border_or_placeholder(&mut scene, border, &mut skipped);
    }
    for primitive in &page.primitives {
        translate_or_placeholder(
            &mut scene,
            primitive,
            fonts,
            images,
            ImageScope::Body,
            width,
            height,
            &mut skipped,
        );
    }
    for area in &page.note_areas {
        let note_kind = if area.kind.as_deref() == Some("endnote") {
            NoteKind::Endnote
        } else {
            NoteKind::Footnote
        };
        let scope = ImageScope::Notes(note_kind);
        for primitive in area.separator_primitives.iter().chain(&area.primitives) {
            translate_or_placeholder(
                &mut scene,
                primitive,
                fonts,
                images,
                scope,
                width,
                height,
                &mut skipped,
            );
        }
    }
    for region in [&page.header, &page.footer].into_iter().flatten() {
        for primitive in &region.primitives {
            translate_or_placeholder(
                &mut scene,
                primitive,
                fonts,
                images,
                ImageScope::HeaderFooter(&region.r_id),
                width,
                height,
                &mut skipped,
            );
        }
    }
    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order != Some(PageBorderZOrder::Back))
    {
        translate_border_or_placeholder(&mut scene, border, &mut skipped);
    }
    Ok(PageScene {
        background: background_scene,
        scene,
        width,
        height,
        skipped,
    })
}

#[allow(clippy::too_many_arguments)]
fn translate_or_placeholder(
    scene: &mut Scene,
    primitive: &Primitive,
    fonts: &FontRegistry,
    images: &ImageRegistry,
    scope: ImageScope<'_>,
    page_width: f64,
    page_height: f64,
    skipped: &mut SkipStats,
) {
    if let Err(reason) = translate_primitive(
        scene,
        primitive,
        fonts,
        images,
        scope,
        page_width,
        page_height,
    ) {
        skipped.record(primitive_kind(primitive), reason);
        draw_placeholder(scene, primitive_bounds(primitive));
    }
}

#[allow(clippy::too_many_arguments)]
fn translate_primitive(
    scene: &mut Scene,
    primitive: &Primitive,
    fonts: &FontRegistry,
    images: &ImageRegistry,
    scope: ImageScope<'_>,
    page_width: f64,
    page_height: f64,
) -> Result<(), String> {
    validate_visual_fields(primitive)?;
    let attrs = primitive_attrs(primitive);
    let mut layers = Vec::with_capacity(2);
    if let Some(clip) = attrs
        .clip_group
        .as_ref()
        .and_then(|group| group.clip.as_ref())
    {
        let rect = clip_rect(clip, page_width, page_height)?;
        layers.push((Affine::IDENTITY, rect));
    }
    match primitive {
        Primitive::Text(run) => {
            if let Some(clip) = &run.paint_clip {
                let rect = text_clip_rect(clip, page_width, page_height)?;
                layers.push((Affine::IDENTITY, rect));
            }
        }
        Primitive::GlyphRun(run) => {
            if let Some(clip) = &run.paint_clip {
                let rect = text_clip_rect(clip, page_width, page_height)?;
                layers.push((Affine::IDENTITY, rect));
            }
        }
        _ => {}
    }
    with_clip_layers(scene, &layers, |scene| match primitive {
        Primitive::Text(run) => draw_text(scene, run, fonts, primitive_opacity(primitive)?),
        Primitive::GlyphRun(run) => {
            draw_glyph_run(scene, run, fonts, primitive_opacity(primitive)?)
        }
        Primitive::Rect(rect) => {
            let bounds = rect_from_numbers(&rect.x, &rect.y, &rect.w, &rect.h)?;
            scene.fill(
                Fill::NonZero,
                Affine::IDENTITY,
                color(&rect.fill, primitive_opacity(primitive)?)?,
                None,
                &bounds,
            );
            Ok(())
        }
        Primitive::Line(line) => draw_line(scene, line, primitive_opacity(primitive)?),
        Primitive::Image(image) => {
            draw_image(scene, image, images, scope, primitive_opacity(primitive)?)
        }
        Primitive::Shape(shape) => draw_shape(scene, shape, primitive_opacity(primitive)?),
        Primitive::Decoration(decoration) => {
            let rect =
                rect_from_numbers(&decoration.x, &decoration.y, &decoration.w, &decoration.h)?;
            let opacity = primitive_opacity(primitive)?;
            let styled = decoration
                .attrs
                .style
                .filter(|style| *style != DisplayBorderStyle::Solid);
            if styled.is_some() || decoration.dashed || decoration.dotted {
                let thickness = rect.height().max(1.0);
                let mut stroke = Stroke::new(thickness);
                if decoration.dotted {
                    stroke = with_dashes(stroke, [thickness, thickness * 2.0]);
                } else if decoration.dashed {
                    stroke = with_dashes(stroke, [thickness * 2.0, thickness * 2.0]);
                } else if let Some(style) = styled {
                    stroke = apply_border_dash(stroke, style, thickness)?;
                }
                let y = rect.y0 + rect.height() / 2.0;
                scene.stroke(
                    &stroke,
                    Affine::IDENTITY,
                    color(&decoration.color, opacity)?,
                    None,
                    &Line::new((rect.x0, y), (rect.x1, y)),
                );
            } else {
                scene.fill(
                    Fill::NonZero,
                    Affine::IDENTITY,
                    color(&decoration.color, opacity)?,
                    None,
                    &rect,
                );
            }
            Ok(())
        }
    })
}

fn draw_text(
    scene: &mut Scene,
    run: &TextRunPrimitive,
    fonts: &FontRegistry,
    opacity: f32,
) -> Result<(), String> {
    let text = if run.all_caps {
        run.text.to_uppercase()
    } else {
        run.text.clone()
    };
    if text.is_empty() {
        return Ok(());
    }
    let letter_spacing = optional_number(run.letter_spacing.as_ref())?.unwrap_or(0.0) as f32;
    let word_spacing = optional_number(run.word_spacing.as_ref())?.unwrap_or(0.0) as f32;
    let (font_id, size, glyphs) = fonts
        .shape_fallback(
            &text,
            &run.font,
            run.rtl == Some(true),
            run.small_caps,
            letter_spacing,
            word_spacing,
        )
        .map_err(|error| error.to_string())?;
    let face = fonts
        .face(font_id)
        .ok_or_else(|| format!("text fallback font id {font_id} is unavailable"))?;
    let x = number(&run.x)?;
    let baseline = number(&run.baseline_y)?;
    let width = number(&run.width)?.max(0.0);
    let bounds = Rect::new(
        x,
        baseline - f64::from(size) * 0.8,
        x + width,
        baseline + f64::from(size) * 0.4,
    );
    let visual = visual_transform(
        bounds,
        run.rotation_deg.as_ref(),
        run.horizontal_scale.as_ref(),
    )?;
    scene
        .draw_glyphs(&face.data)
        .font_size(size)
        .brush(color(&run.color, 1.0)?)
        .brush_alpha(opacity)
        .transform(visual * Affine::translate((x, baseline)))
        .draw(Fill::NonZero, glyphs.into_iter());
    Ok(())
}

fn draw_glyph_run(
    scene: &mut Scene,
    run: &GlyphRunPrimitive,
    fonts: &FontRegistry,
    opacity: f32,
) -> Result<(), String> {
    if run.glyphs.is_empty() {
        return Ok(());
    }
    if !run.size.is_finite() || run.size <= 0.0 || run.size > f32::MAX as f64 {
        return Err("glyphRun size is invalid".to_owned());
    }
    let face = fonts
        .face(run.font_id)
        .ok_or_else(|| format!("glyphRun font id {} is unavailable", run.font_id))?;
    let glyphs = run
        .glyphs
        .iter()
        .map(|glyph| {
            Ok(Glyph {
                id: glyph.id,
                x: finite_f32(glyph.x, "glyph x")?,
                y: finite_f32(glyph.y, "glyph y")?,
            })
        })
        .collect::<Result<Vec<_>, String>>()?;
    let bounds = glyph_bounds(run)?;
    let transform = visual_transform(
        bounds,
        run.rotation_deg.as_ref(),
        run.horizontal_scale.as_ref(),
    )?;
    scene
        .draw_glyphs(&face.data)
        .font_size(run.size as f32)
        .brush(color(&run.color, 1.0)?)
        .brush_alpha(opacity)
        .transform(transform)
        .draw(Fill::NonZero, glyphs.into_iter());
    Ok(())
}

fn draw_line(scene: &mut Scene, line: &LinePrimitive, opacity: f32) -> Result<(), String> {
    if line.secondary_color.is_some() {
        return Err("line secondary colors are not translated".to_owned());
    }
    let width = number(&line.stroke_width)?.max(0.5);
    let mut stroke = Stroke::new(width);
    if let Some(dash) = &line.dash {
        let pattern = dash.iter().map(number).collect::<Result<Vec<_>, _>>()?;
        if !pattern.is_empty() {
            stroke = with_dashes(stroke, pattern);
        }
    } else if let Some(style) = line.border_style {
        stroke = apply_border_dash(stroke, style, width)?;
    }
    scene.stroke(
        &stroke,
        Affine::IDENTITY,
        color(&line.color, opacity)?,
        None,
        &Line::new(
            (number(&line.x1)?, number(&line.y1)?),
            (number(&line.x2)?, number(&line.y2)?),
        ),
    );
    Ok(())
}

fn draw_shape(scene: &mut Scene, shape: &ShapePrimitive, opacity: f32) -> Result<(), String> {
    let mut path = BezPath::new();
    for command in &shape.geometry_path {
        match command {
            ShapePathCommand::Move { x, y } => path.move_to((number(x)?, number(y)?)),
            ShapePathCommand::Line { x, y } => path.line_to((number(x)?, number(y)?)),
            ShapePathCommand::Quad { cpx, cpy, x, y } => {
                path.quad_to((number(cpx)?, number(cpy)?), (number(x)?, number(y)?));
            }
            ShapePathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => path.curve_to(
                (number(cp1x)?, number(cp1y)?),
                (number(cp2x)?, number(cp2y)?),
                (number(x)?, number(y)?),
            ),
            ShapePathCommand::Close => path.close_path(),
        }
    }
    let bounds = rect_from_numbers(&shape.x, &shape.y, &shape.w, &shape.h)?;
    let mut transform = Affine::IDENTITY;
    if let Some(value) = &shape.transform {
        let center = bounds.center();
        if value.flip_h || value.flip_v {
            transform = Affine::translate(center.to_vec2())
                * Affine::scale_non_uniform(
                    if value.flip_h { -1.0 } else { 1.0 },
                    if value.flip_v { -1.0 } else { 1.0 },
                )
                * Affine::translate(-center.to_vec2());
        }
        if let Some(rotation) = &value.rotation {
            transform = Affine::rotate_about(number(rotation)?.to_radians(), center) * transform;
        }
    }
    let fill = shape
        .fill
        .as_deref()
        .map(|fill| color(fill, opacity))
        .transpose()?;
    let stroke = shape
        .stroke
        .as_ref()
        .map(|stroke_spec| -> Result<_, String> {
            let width = number(&stroke_spec.width)?.max(0.5);
            let stroke = shape_dash(Stroke::new(width), stroke_spec.dash.as_deref(), width)?;
            Ok((stroke, color(&stroke_spec.color, opacity)?))
        })
        .transpose()?;
    if let Some(fill) = fill {
        scene.fill(Fill::NonZero, transform, fill, None, &path);
    }
    if let Some((stroke, stroke_color)) = stroke {
        scene.stroke(&stroke, transform, stroke_color, None, &path);
    }
    Ok(())
}

fn image_frame(primitive: &ImagePrimitive) -> Result<Rect, String> {
    let frame = primitive.attrs.content_frame.as_ref();
    rect_from_numbers(
        frame
            .and_then(|frame| frame.x.as_ref())
            .unwrap_or(&primitive.x),
        frame
            .and_then(|frame| frame.y.as_ref())
            .unwrap_or(&primitive.y),
        frame
            .and_then(|frame| frame.w.as_ref())
            .unwrap_or(&primitive.w),
        frame
            .and_then(|frame| frame.h.as_ref())
            .unwrap_or(&primitive.h),
    )
}

fn draw_image(
    scene: &mut Scene,
    primitive: &ImagePrimitive,
    images: &ImageRegistry,
    scope: ImageScope<'_>,
    opacity: f32,
) -> Result<(), String> {
    if primitive
        .filter
        .as_deref()
        .is_some_and(|filter| filter != "none")
    {
        return Err(format!(
            "image filter {} is not translated",
            primitive.filter.as_deref().unwrap_or_default()
        ));
    }
    let image = images
        .get(scope, &primitive.rel_id)
        .ok_or_else(|| format!("image {} is missing or undecodable", primitive.rel_id))?;
    let target = image_frame(primitive)?;
    if target.width() == 0.0 || target.height() == 0.0 {
        return Ok(());
    }
    let (left, top, right, bottom) = if let Some(crop) = &primitive.crop {
        (
            number(&crop.left)?,
            number(&crop.top)?,
            number(&crop.right)?,
            number(&crop.bottom)?,
        )
    } else {
        (0.0, 0.0, 0.0, 0.0)
    };
    if [left, top, right, bottom]
        .iter()
        .any(|value| !value.is_finite() || *value < 0.0 || *value > 1.0)
        || left + right >= 1.0
        || top + bottom >= 1.0
    {
        return Err("image crop is invalid".to_owned());
    }
    let source_width = f64::from(image.width);
    let source_height = f64::from(image.height);
    let Some((crop_x, crop_y, crop_width, crop_height)) =
        cropped_source_rect(source_width, source_height, left, top, right, bottom)
    else {
        return Ok(());
    };

    let mut visual = Affine::IDENTITY;
    let center = target.center();
    if primitive.attrs.image_flip_h == Some(true) || primitive.attrs.image_flip_v == Some(true) {
        visual = Affine::translate(center.to_vec2())
            * Affine::scale_non_uniform(
                if primitive.attrs.image_flip_h == Some(true) {
                    -1.0
                } else {
                    1.0
                },
                if primitive.attrs.image_flip_v == Some(true) {
                    -1.0
                } else {
                    1.0
                },
            )
            * Affine::translate(-center.to_vec2());
    }
    if let Some(rotation) = &primitive.rotation_deg {
        visual = Affine::rotate_about(number(rotation)?.to_radians(), center) * visual;
    }
    let mapping = Affine::translate((target.x0, target.y0))
        * Affine::scale_non_uniform(target.width() / crop_width, target.height() / crop_height)
        * Affine::translate((-crop_x, -crop_y));
    scene.push_clip_layer(Fill::NonZero, visual, &target);
    let brush = ImageBrush::new(image).with_alpha(opacity);
    scene.draw_image(&brush, visual * mapping);
    scene.pop_layer();
    Ok(())
}

fn cropped_source_rect(
    source_width: f64,
    source_height: f64,
    left: f64,
    top: f64,
    right: f64,
    bottom: f64,
) -> Option<(f64, f64, f64, f64)> {
    let crop_x = left * source_width;
    let crop_y = top * source_height;
    let crop_width = (1.0 - left - right) * source_width;
    let crop_height = (1.0 - top - bottom) * source_height;
    if crop_width <= 0.0 || crop_height <= 0.0 {
        return None;
    }
    Some((crop_x, crop_y, crop_width, crop_height))
}

fn translate_border_or_placeholder(
    scene: &mut Scene,
    border: &PageBorderPrimitive,
    skipped: &mut SkipStats,
) {
    if let Err(reason) = draw_page_border(scene, border) {
        skipped.record("pageBorder", reason);
        draw_placeholder(
            scene,
            rect_from_numbers(&border.x, &border.y, &border.w, &border.h).ok(),
        );
    }
}

fn draw_page_border(scene: &mut Scene, border: &PageBorderPrimitive) -> Result<(), String> {
    let rect = rect_from_numbers(&border.x, &border.y, &border.w, &border.h)?;
    let mut prepared = Vec::new();
    for (side, line) in [
        (
            &border.top,
            Line::new((rect.x0, rect.y0), (rect.x1, rect.y0)),
        ),
        (
            &border.right,
            Line::new((rect.x1, rect.y0), (rect.x1, rect.y1)),
        ),
        (
            &border.bottom,
            Line::new((rect.x0, rect.y1), (rect.x1, rect.y1)),
        ),
        (
            &border.left,
            Line::new((rect.x0, rect.y0), (rect.x0, rect.y1)),
        ),
    ] {
        if let Some(side) = side {
            prepared.push(prepare_page_border_side(side, line)?);
        }
    }
    for (stroke, color, line) in prepared {
        scene.stroke(&stroke, Affine::IDENTITY, color, None, &line);
    }
    Ok(())
}

fn prepare_page_border_side(
    side: &PageBorderSide,
    line: Line,
) -> Result<(Stroke, Color, Line), String> {
    let width = number(&side.width)?.max(0.5);
    let style = match side.style.as_str() {
        "solid" => DisplayBorderStyle::Solid,
        "dotted" => DisplayBorderStyle::Dotted,
        "dashed" => DisplayBorderStyle::Dashed,
        "dashDot" => DisplayBorderStyle::DashDot,
        "dashDotDot" => DisplayBorderStyle::DashDotDot,
        other => return Err(format!("page border style {other} is not translated")),
    };
    let stroke = apply_border_dash(Stroke::new(width), style, width)?;
    Ok((stroke, color(&side.color, 1.0)?, line))
}

fn validate_visual_fields(primitive: &Primitive) -> Result<(), String> {
    let attrs = primitive_attrs(primitive);
    let is_text = matches!(primitive, Primitive::Text(_));
    let is_image = matches!(primitive, Primitive::Image(_));
    let is_shape = matches!(primitive, Primitive::Shape(_));
    let is_decoration = matches!(primitive, Primitive::Decoration(_));
    let is_glyph = matches!(primitive, Primitive::GlyphRun(_));
    if !is_image
        && (attrs.image_flip_h == Some(true)
            || attrs.image_flip_v == Some(true)
            || attrs.content_frame.is_some()
            || attrs.border.is_some())
    {
        return Err("image visual fields are attached to another primitive".to_owned());
    }
    if !is_shape
        && (attrs.fill_paint.is_some()
            || attrs.stroke_paint.is_some()
            || attrs.effect_extent.is_some()
            || attrs.drawing_scene.is_some()
            || attrs.text_body_properties.is_some())
    {
        return Err("shape visual fields are attached to another primitive".to_owned());
    }
    if !is_decoration && attrs.style.is_some() {
        return Err("decoration style is attached to another primitive".to_owned());
    }
    if !is_glyph && attrs.fallback_font.is_some() {
        return Err("glyph fallback font is attached to another primitive".to_owned());
    }
    if !is_text && attrs.leader_glyphs.is_some() {
        return Err("leader glyphs are attached to another primitive".to_owned());
    }
    if !is_text && !is_glyph && attrs.modern_effects.is_some() {
        return Err("modern text effects are attached to another primitive".to_owned());
    }
    if is_shape
        && (attrs.fill_paint.is_some()
            || attrs.stroke_paint.is_some()
            || attrs.effect_extent.is_some()
            || attrs.drawing_scene.is_some()
            || attrs.text_body_properties.is_some())
    {
        return Err("advanced shape paint is not translated".to_owned());
    }
    if is_image && attrs.border.is_some() {
        return Err("image border is not translated".to_owned());
    }
    if !attrs.effects.is_empty() {
        return Err("DrawingML effects are not translated".to_owned());
    }
    match primitive {
        Primitive::Text(run)
            if run.text_shadow.is_some()
                || run.text_outline
                || run.emphasis_mark.is_some()
                || run.text_effect.is_some()
                || attrs.modern_effects.is_some()
                || attrs.leader_glyphs.is_some() =>
        {
            return Err("advanced text effects are not translated".to_owned());
        }
        Primitive::GlyphRun(run)
            if run.all_caps
                || run.small_caps
                || run.text_shadow.is_some()
                || run.text_outline
                || run.emphasis_mark.is_some()
                || run.text_effect.is_some()
                || attrs.modern_effects.is_some()
                || attrs.leader_glyphs.is_some() =>
        {
            return Err("advanced glyphRun effects are not translated".to_owned());
        }
        _ => {}
    }
    Ok(())
}

fn primitive_opacity(primitive: &Primitive) -> Result<f32, String> {
    let attrs = primitive_attrs(primitive);
    let own = match primitive {
        Primitive::Text(value) => value.opacity.as_ref(),
        Primitive::GlyphRun(value) => value.opacity.as_ref(),
        Primitive::Line(value) => value.opacity.as_ref(),
        Primitive::Image(value) => value.opacity.as_ref(),
        _ => None,
    }
    .map(number)
    .transpose()?
    .unwrap_or(match primitive {
        Primitive::Text(value) if value.hidden => 0.4,
        Primitive::GlyphRun(value) if value.hidden => 0.4,
        _ => 1.0,
    });
    let flattened = attrs
        .primitive_opacity
        .as_ref()
        .map(number)
        .transpose()?
        .unwrap_or(1.0);
    let group = attrs
        .clip_group
        .as_ref()
        .and_then(|group| group.opacity.as_ref())
        .map(number)
        .transpose()?
        .unwrap_or(1.0);
    Ok((own * flattened * group).clamp(0.0, 1.0) as f32)
}

fn apply_border_dash(
    stroke: Stroke,
    style: DisplayBorderStyle,
    width: f64,
) -> Result<Stroke, String> {
    let dotted = [width.max(1.0), (width * 2.0).max(2.0)];
    let dashed = [(width * 3.0).max(2.0), (width * 2.0).max(2.0)];
    match style {
        DisplayBorderStyle::Solid => Ok(stroke),
        DisplayBorderStyle::Dotted => Ok(with_dashes(stroke, dotted)),
        DisplayBorderStyle::Dashed => Ok(with_dashes(stroke, dashed)),
        DisplayBorderStyle::DashDot => Ok(with_dashes(
            stroke,
            [dashed[0], dashed[1], width, dashed[1]],
        )),
        DisplayBorderStyle::DashDotDot => Ok(with_dashes(
            stroke,
            [dashed[0], dashed[1], width, dashed[1], width, dashed[1]],
        )),
        other => Err(format!("line border style {other:?} is not translated")),
    }
}

fn shape_dash(stroke: Stroke, name: Option<&str>, width: f64) -> Result<Stroke, String> {
    match name.unwrap_or("") {
        "" | "solid" => Ok(stroke),
        "dot" | "dotted" | "sysDot" => Ok(with_dashes(
            stroke,
            [width.max(1.0), (width * 2.0).max(2.0)],
        )),
        "dash" | "dashed" | "dashSmallGap" | "sysDash" => Ok(with_dashes(
            stroke,
            [(width * 3.0).max(2.0), (width * 2.0).max(2.0)],
        )),
        "lgDash" | "dashLong" | "dashLongHeavy" => Ok(with_dashes(
            stroke,
            [(width * 6.0).max(4.0), (width * 2.0).max(2.0)],
        )),
        "dashDot" | "lgDashDot" | "sysDashDot" | "dashDotHeavy" => Ok(with_dashes(
            stroke,
            [
                (width * 3.0).max(2.0),
                (width * 2.0).max(2.0),
                width,
                (width * 2.0).max(2.0),
            ],
        )),
        "dashDotDot" | "lgDashDotDot" | "sysDashDotDot" | "dashDotDotHeavy" => Ok(with_dashes(
            stroke,
            [
                (width * 3.0).max(2.0),
                (width * 2.0).max(2.0),
                width,
                (width * 2.0).max(2.0),
                width,
                (width * 2.0).max(2.0),
            ],
        )),
        other => Err(format!("shape dash style {other} is not translated")),
    }
}

fn visual_transform(
    bounds: Rect,
    rotation: Option<&Number>,
    horizontal_scale: Option<&Number>,
) -> Result<Affine, String> {
    let mut transform = Affine::IDENTITY;
    if let Some(rotation) = rotation {
        transform =
            Affine::rotate_about(number(rotation)?.to_radians(), bounds.center()) * transform;
    }
    if let Some(horizontal_scale) = horizontal_scale {
        let scale = number(horizontal_scale)? / 100.0;
        if scale <= 0.0 {
            return Err("horizontalScale must be positive".to_owned());
        }
        let origin = (bounds.x0, bounds.center().y);
        transform = Affine::translate(origin)
            * Affine::scale_non_uniform(scale, 1.0)
            * Affine::translate((-origin.0, -origin.1))
            * transform;
    }
    Ok(transform)
}

fn primitive_attrs(primitive: &Primitive) -> &DocAttrs {
    match primitive {
        Primitive::Text(value) => &value.attrs,
        Primitive::GlyphRun(value) => &value.attrs,
        Primitive::Rect(value) => &value.attrs,
        Primitive::Line(value) => &value.attrs,
        Primitive::Image(value) => &value.attrs,
        Primitive::Shape(value) => &value.attrs,
        Primitive::Decoration(value) => &value.attrs,
    }
}

fn primitive_kind(primitive: &Primitive) -> &'static str {
    match primitive {
        Primitive::Text(_) => "text",
        Primitive::GlyphRun(_) => "glyphRun",
        Primitive::Rect(_) => "rect",
        Primitive::Line(_) => "line",
        Primitive::Image(_) => "image",
        Primitive::Shape(_) => "shape",
        Primitive::Decoration(_) => "decoration",
    }
}

fn primitive_bounds(primitive: &Primitive) -> Option<Rect> {
    match primitive {
        Primitive::Text(run) => {
            let x = number(&run.x).ok()?;
            let baseline = number(&run.baseline_y).ok()?;
            let width = number(&run.width).ok()?.max(1.0);
            let size = run
                .font
                .split_whitespace()
                .find_map(|token| token.strip_suffix("px")?.parse::<f64>().ok())
                .unwrap_or(12.0);
            Some(Rect::new(
                x,
                baseline - size,
                x + width,
                baseline + size * 0.3,
            ))
        }
        Primitive::GlyphRun(run) => glyph_bounds(run).ok(),
        Primitive::Rect(value) => rect_from_numbers(&value.x, &value.y, &value.w, &value.h).ok(),
        Primitive::Line(value) => {
            let x1 = number(&value.x1).ok()?;
            let y1 = number(&value.y1).ok()?;
            let x2 = number(&value.x2).ok()?;
            let y2 = number(&value.y2).ok()?;
            Some(Rect::new(x1.min(x2), y1.min(y2), x1.max(x2), y1.max(y2)))
        }
        Primitive::Image(value) => rect_from_numbers(&value.x, &value.y, &value.w, &value.h).ok(),
        Primitive::Shape(value) => rect_from_numbers(&value.x, &value.y, &value.w, &value.h).ok(),
        Primitive::Decoration(value) => {
            rect_from_numbers(&value.x, &value.y, &value.w, &value.h).ok()
        }
    }
}

fn glyph_bounds(run: &GlyphRunPrimitive) -> Result<Rect, String> {
    let first = run
        .glyphs
        .first()
        .context("glyphRun is empty")
        .map_err(|e| e.to_string())?;
    let mut left = f64::INFINITY;
    let mut right = f64::NEG_INFINITY;
    for glyph in &run.glyphs {
        left = left.min(glyph.x);
        right = right.max(glyph.x + glyph.advance);
    }
    Ok(Rect::new(
        left,
        first.y - run.size * 0.8,
        right.max(left + 1.0),
        first.y + run.size * 0.4,
    ))
}

fn rect_from_numbers(x: &Number, y: &Number, w: &Number, h: &Number) -> Result<Rect, String> {
    let x = number(x)?;
    let y = number(y)?;
    let w = number(w)?;
    let h = number(h)?;
    if w < 0.0 || h < 0.0 {
        return Err("rectangle dimensions are negative".to_owned());
    }
    Ok(Rect::new(x, y, x + w, y + h))
}

fn clip_rect(
    clip: &docx_layout::display_list::ClipRect,
    page_width: f64,
    page_height: f64,
) -> Result<Rect, String> {
    let x = optional_number(clip.x.as_ref())?.unwrap_or(0.0);
    let y = optional_number(clip.y.as_ref())?.unwrap_or(0.0);
    let width = optional_number(clip.w.as_ref())?.unwrap_or(page_width);
    let height = optional_number(clip.h.as_ref())?.unwrap_or(page_height);
    if width < 0.0 || height < 0.0 {
        return Err("clip dimensions are negative".to_owned());
    }
    Ok(Rect::new(x, y, x + width, y + height))
}

fn text_clip_rect(
    clip: &docx_layout::display_list::ClipRect,
    page_width: f64,
    page_height: f64,
) -> Result<Rect, String> {
    let x = optional_number(clip.x.as_ref())?.unwrap_or(0.0);
    let y = optional_number(clip.y.as_ref())?.unwrap_or(0.0);
    let width = optional_number(clip.w.as_ref())?.unwrap_or(page_width);
    let height = optional_number(clip.h.as_ref())?.unwrap_or(page_height);
    Ok(Rect::new(x, y, x + width.max(0.0), y + height.max(0.0)))
}

fn number(value: &Number) -> Result<f64, String> {
    value
        .as_f64()
        .filter(|value| value.is_finite())
        .ok_or_else(|| "display-list number is invalid".to_owned())
}

fn optional_number(value: Option<&Number>) -> Result<Option<f64>, String> {
    value.map(number).transpose()
}

fn finite_f32(value: f64, name: &str) -> Result<f32, String> {
    let value = value as f32;
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| format!("{name} is outside the Vello coordinate range"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::ShapeStrokePrimitive;

    fn n(value: i64) -> Number {
        Number::from(value)
    }

    #[test]
    fn invalid_shape_stroke_does_not_leave_a_fill() {
        let shape = ShapePrimitive {
            x: n(0),
            y: n(0),
            w: n(20),
            h: n(20),
            geometry_path: vec![
                ShapePathCommand::Move { x: n(0), y: n(0) },
                ShapePathCommand::Line { x: n(20), y: n(0) },
                ShapePathCommand::Line { x: n(20), y: n(20) },
                ShapePathCommand::Close,
            ],
            fill: Some("#00ff00".to_owned()),
            stroke: Some(ShapeStrokePrimitive {
                color: "#000000".to_owned(),
                width: n(1),
                dash: Some("unsupported".to_owned()),
            }),
            transform: None,
            decorative: false,
            attrs: DocAttrs::default(),
        };
        let mut scene = Scene::new();
        assert!(draw_shape(&mut scene, &shape, 1.0).is_err());
        assert!(scene.encoding().is_empty());
    }

    #[test]
    fn invalid_later_page_border_does_not_leave_an_earlier_side() {
        let border = PageBorderPrimitive {
            x: n(0),
            y: n(0),
            w: n(20),
            h: n(20),
            z_order: None,
            top: Some(PageBorderSide {
                width: n(1),
                color: "#000000".to_owned(),
                style: "solid".to_owned(),
            }),
            right: Some(PageBorderSide {
                width: n(1),
                color: "#000000".to_owned(),
                style: "unsupported".to_owned(),
            }),
            bottom: None,
            left: None,
        };
        let mut scene = Scene::new();
        assert!(draw_page_border(&mut scene, &border).is_err());
        assert!(scene.encoding().is_empty());
    }

    #[test]
    fn zero_width_image_crop_is_empty() {
        assert!(cropped_source_rect(0.0, 80.0, 0.0, 0.0, 0.0, 0.0).is_none());
    }
}
