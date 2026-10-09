//! Media, scaling and DrawingML conversions shared by both exporters.

use std::collections::BTreeMap;

use ooxml_drawingml::{
    EmitGradientStop, EmitParagraph, EmitRun, Fill, GeometryPathCommand, Line, Matrix,
    dml_paragraphs, fill_xml, line_xml,
};
use vsdx_parse::VsdxPackage;
use vsdx_render::{Affine, Paint, Primitive, Stroke, TextParagraph, VsdxDisplayList};

pub struct Media {
    pub part: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

pub fn collect_media(
    primitives: &[Primitive],
    package: &VsdxPackage,
    media: &mut Vec<Media>,
    index_by_asset: &mut BTreeMap<String, usize>,
    prefix: &str,
) {
    for primitive in primitives {
        match primitive {
            Primitive::Group { primitives, .. } => {
                collect_media(primitives, package, media, index_by_asset, prefix);
            }
            Primitive::Image { asset_id, .. } => {
                if index_by_asset.contains_key(asset_id) {
                    continue;
                }
                let Some(bytes) = package.part_bytes(asset_id) else {
                    continue;
                };
                let Some((ext, content_type)) = sniff_image(asset_id, bytes) else {
                    continue;
                };
                let part = format!("{prefix}/image{}.{ext}", media.len() + 1);
                index_by_asset.insert(asset_id.clone(), media.len());
                media.push(Media {
                    part,
                    content_type: content_type.to_owned(),
                    bytes: bytes.to_vec(),
                });
            }
            _ => {}
        }
    }
}

pub fn used_media(
    primitives: &[Primitive],
    index_by_asset: &BTreeMap<String, usize>,
) -> Vec<usize> {
    let mut used = Vec::new();
    let mut visit: Vec<&[Primitive]> = vec![primitives];
    while let Some(list) = visit.pop() {
        for primitive in list {
            match primitive {
                Primitive::Group { primitives, .. } => visit.push(primitives),
                Primitive::Image { asset_id, .. } => {
                    if let Some(index) = index_by_asset.get(asset_id)
                        && !used.contains(index)
                    {
                        used.push(*index);
                    }
                }
                _ => {}
            }
        }
    }
    used.sort_unstable();
    used
}

pub fn sniff_image(asset_id: &str, bytes: &[u8]) -> Option<(&'static str, &'static str)> {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G', b'\r', b'\n', 0x1A, b'\n']) {
        Some(("png", "image/png"))
    } else if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        Some(("jpg", "image/jpeg"))
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        Some(("gif", "image/gif"))
    } else if bytes.starts_with(b"BM") {
        Some(("bmp", "image/bmp"))
    } else if has_extension(asset_id, "png") {
        Some(("png", "image/png"))
    } else if has_extension(asset_id, "jpg") || has_extension(asset_id, "jpeg") {
        Some(("jpg", "image/jpeg"))
    } else if has_extension(asset_id, "emf") {
        Some(("emf", "image/x-emf"))
    } else if has_extension(asset_id, "wmf") {
        Some(("wmf", "image/x-wmf"))
    } else {
        None
    }
}

fn has_extension(asset_id: &str, extension: &str) -> bool {
    asset_id
        .rsplit_once('.')
        .is_some_and(|(_, actual)| actual.eq_ignore_ascii_case(extension))
}

pub fn scale_list(list: &VsdxDisplayList, scale: f64) -> VsdxDisplayList {
    let mut list = list.clone();
    list.width *= scale as f32;
    list.height *= scale as f32;
    scale_primitives(&mut list.primitives, scale);
    list
}

fn scale_primitives(primitives: &mut [Primitive], scale: f64) {
    for primitive in primitives {
        match primitive {
            Primitive::Shape { path, stroke, .. } => {
                for command in path.iter_mut() {
                    scale_command(command, scale);
                }
                if let Some(stroke) = stroke {
                    stroke.width *= scale as f32;
                }
            }
            Primitive::Image {
                x,
                y,
                width,
                height,
                transform,
                ..
            } => {
                *x *= scale as f32;
                *y *= scale as f32;
                *width *= scale as f32;
                *height *= scale as f32;
                scale_affine(transform, scale);
            }
            Primitive::TextBox {
                x,
                y,
                width,
                height,
                paragraphs,
                lines,
                transform,
                ..
            } => {
                *x *= scale as f32;
                *y *= scale as f32;
                *width *= scale as f32;
                *height *= scale as f32;
                for paragraph in paragraphs {
                    for run in &mut paragraph.runs {
                        run.size_in *= scale as f32;
                        run.letter_spacing *= scale as f32;
                    }
                }
                for line in lines {
                    line.x *= scale as f32;
                    line.y *= scale as f32;
                    line.width *= scale as f32;
                    line.height *= scale as f32;
                    for stop in &mut line.caret_stops {
                        stop.x *= scale as f32;
                        stop.y *= scale as f32;
                    }
                }
                scale_affine(transform, scale);
            }
            Primitive::Placeholder {
                x,
                y,
                width,
                height,
                ..
            } => {
                *x *= scale as f32;
                *y *= scale as f32;
                *width *= scale as f32;
                *height *= scale as f32;
            }
            Primitive::Group {
                primitives,
                transform,
                ..
            } => {
                scale_primitives(primitives, scale);
                scale_affine(transform, scale);
            }
        }
    }
}

fn scale_command(command: &mut GeometryPathCommand, scale: f64) {
    match command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
            *x *= scale;
            *y *= scale;
        }
        GeometryPathCommand::Quad { cpx, cpy, x, y } => {
            *cpx *= scale;
            *cpy *= scale;
            *x *= scale;
            *y *= scale;
        }
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            *cp1x *= scale;
            *cp1y *= scale;
            *cp2x *= scale;
            *cp2y *= scale;
            *x *= scale;
            *y *= scale;
        }
        GeometryPathCommand::Close => {}
    }
}

fn scale_affine(transform: &mut Affine, scale: f64) {
    transform.e *= scale as f32;
    transform.f *= scale as f32;
}

/// Shifts a scaled list in scene inches so a fitted page centres on canvas.
pub fn translate_primitives(primitives: &mut [Primitive], dx: f32, dy: f32) {
    for primitive in primitives {
        match primitive {
            Primitive::Shape { path, .. } => {
                for command in path.iter_mut() {
                    shift_command(command, f64::from(dx), f64::from(dy));
                }
            }
            Primitive::Image { transform, .. } | Primitive::TextBox { transform, .. } => {
                transform.e += dx;
                transform.f += dy;
            }
            Primitive::Placeholder { x, y, .. } => {
                *x += dx;
                *y += dy;
            }
            Primitive::Group { transform, .. } => {
                transform.e += dx;
                transform.f += dy;
            }
        }
    }
}

fn shift_command(command: &mut GeometryPathCommand, dx: f64, dy: f64) {
    match command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
            *x += dx;
            *y += dy;
        }
        GeometryPathCommand::Quad { cpx, cpy, x, y } => {
            *cpx += dx;
            *cpy += dy;
            *x += dx;
            *y += dy;
        }
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            *cp1x += dx;
            *cp1y += dy;
            *cp2x += dx;
            *cp2y += dy;
            *x += dx;
            *y += dy;
        }
        GeometryPathCommand::Close => {}
    }
}

/// Bounding frame of the finite path points, in EMUs Y-down from page top.
pub fn fallback_rect(
    path: &[GeometryPathCommand],
    page_height_in: f64,
) -> Option<(i64, i64, i64, i64)> {
    let mut points: Vec<(f64, f64)> = Vec::new();
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                points.push((*x, page_height_in - *y));
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                points.push((*cpx, page_height_in - *cpy));
                points.push((*x, page_height_in - *y));
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                points.push((*cp1x, page_height_in - *cp1y));
                points.push((*cp2x, page_height_in - *cp2y));
                points.push((*x, page_height_in - *y));
            }
            GeometryPathCommand::Close => {}
        }
    }
    points.retain(|(x, y)| x.is_finite() && y.is_finite());
    if points.is_empty() {
        return None;
    }
    let min_x = points
        .iter()
        .map(|point| point.0)
        .fold(f64::INFINITY, f64::min);
    let max_x = points
        .iter()
        .map(|point| point.0)
        .fold(f64::NEG_INFINITY, f64::max);
    let min_y = points
        .iter()
        .map(|point| point.1)
        .fold(f64::INFINITY, f64::min);
    let max_y = points
        .iter()
        .map(|point| point.1)
        .fold(f64::NEG_INFINITY, f64::max);
    Some((
        ooxml_drawingml::emu(min_x),
        ooxml_drawingml::emu(min_y),
        ooxml_drawingml::emu(max_x - min_x).max(1),
        ooxml_drawingml::emu(max_y - min_y).max(1),
    ))
}

/// Renders a fill element for a display-list paint, degrading to no fill.
pub fn fill_string(fill: &Option<Paint>) -> String {
    match fill {
        None => fill_xml(&Fill::NoFill),
        Some(Paint::Solid { color }) => fill_xml(&Fill::Solid { hex: color }),
        Some(Paint::Gradient { stops, angle_deg }) => {
            let stops: Vec<EmitGradientStop> = stops
                .iter()
                .map(|stop| EmitGradientStop {
                    position: stop.position,
                    color: stop.color.as_str(),
                })
                .collect();
            fill_xml(&Fill::Gradient {
                stops: &stops,
                angle_deg: *angle_deg,
            })
        }
    }
}

/// Renders an outline element; absent for no stroke, `noFill` for bad colour.
pub fn line_string(stroke: &Option<Stroke>) -> String {
    line_xml(
        stroke
            .as_ref()
            .map(|stroke| Line {
                color: stroke.color.as_str(),
                width_in: f64::from(stroke.width),
                dashed: stroke.dashed,
            })
            .as_ref(),
    )
}

/// Renders DrawingML paragraphs for display-list text.
pub fn paragraph_string(paragraphs: &[TextParagraph]) -> String {
    let runs: Vec<Vec<EmitRun>> = paragraphs
        .iter()
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .map(|run| EmitRun {
                    text: run.text.as_str(),
                    family: run.family.as_str(),
                    size_in: run.size_in,
                    bold: run.bold,
                    italic: run.italic,
                    underline: run.underline,
                    small_caps: run.small_caps,
                    superscript: run.superscript,
                    subscript: run.subscript,
                    color: run.color.as_str(),
                })
                .collect()
        })
        .collect();
    let paragraphs: Vec<EmitParagraph> = runs.iter().map(|runs| EmitParagraph { runs }).collect();
    dml_paragraphs(&paragraphs)
}

/// Reports whether any paragraph carries text.
pub fn has_text(paragraphs: &[TextParagraph]) -> bool {
    paragraphs
        .iter()
        .any(|paragraph| paragraph.runs.iter().any(|run| !run.text.is_empty()))
}

/// Converts a display-list affine into a DrawingML placement matrix.
pub fn emit_matrix(transform: Affine) -> Matrix {
    Matrix {
        a: transform.a,
        b: transform.b,
        c: transform.c,
        d: transform.d,
        e: transform.e,
        f: transform.f,
    }
}

pub fn flat(primitives: &[Primitive]) -> Vec<Primitive> {
    let mut out = Vec::new();
    flat_with_transform(primitives, Affine::identity(), &mut out);
    out
}

fn flat_with_transform(primitives: &[Primitive], parent: Affine, out: &mut Vec<Primitive>) {
    let mut ordered: Vec<&Primitive> = primitives.iter().collect();
    ordered.sort_by_key(z_order);
    for primitive in ordered {
        let mut primitive = primitive.clone();
        match &mut primitive {
            Primitive::Group {
                primitives,
                transform,
                ..
            } => flat_with_transform(primitives, parent.compose(*transform), out),
            _ => {
                bake_transform(&mut primitive, parent);
                out.push(primitive);
            }
        }
    }
}

fn z_order(primitive: &&Primitive) -> u32 {
    match primitive {
        Primitive::Shape { z_order, .. }
        | Primitive::Image { z_order, .. }
        | Primitive::TextBox { z_order, .. }
        | Primitive::Placeholder { z_order, .. }
        | Primitive::Group { z_order, .. } => *z_order,
    }
}

fn bake_transform(primitive: &mut Primitive, matrix: Affine) {
    match primitive {
        Primitive::Shape {
            path, transform, ..
        } => {
            for command in path {
                transform_command(command, matrix);
            }
            *transform = Affine::identity();
        }
        Primitive::Image { transform, .. } | Primitive::TextBox { transform, .. } => {
            *transform = matrix.compose(*transform);
        }
        Primitive::Placeholder {
            x,
            y,
            width,
            height,
            ..
        } => transform_rect(x, y, width, height, matrix),
        Primitive::Group { .. } => {}
    }
}

fn transform_command(command: &mut GeometryPathCommand, matrix: Affine) {
    match command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
            let (px, py) = matrix.apply_point(*x as f32, *y as f32);
            (*x, *y) = (f64::from(px), f64::from(py));
        }
        GeometryPathCommand::Quad { cpx, cpy, x, y } => {
            let (px, py) = matrix.apply_point(*cpx as f32, *cpy as f32);
            (*cpx, *cpy) = (f64::from(px), f64::from(py));
            let (px, py) = matrix.apply_point(*x as f32, *y as f32);
            (*x, *y) = (f64::from(px), f64::from(py));
        }
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            let (px, py) = matrix.apply_point(*cp1x as f32, *cp1y as f32);
            (*cp1x, *cp1y) = (f64::from(px), f64::from(py));
            let (px, py) = matrix.apply_point(*cp2x as f32, *cp2y as f32);
            (*cp2x, *cp2y) = (f64::from(px), f64::from(py));
            let (px, py) = matrix.apply_point(*x as f32, *y as f32);
            (*x, *y) = (f64::from(px), f64::from(py));
        }
        GeometryPathCommand::Close => {}
    }
}

fn transform_rect(x: &mut f32, y: &mut f32, width: &mut f32, height: &mut f32, matrix: Affine) {
    let corners = [
        matrix.apply_point(*x, *y),
        matrix.apply_point(*x + *width, *y),
        matrix.apply_point(*x, *y + *height),
        matrix.apply_point(*x + *width, *y + *height),
    ];
    let min_x = corners
        .iter()
        .map(|(x, _)| *x)
        .fold(f32::INFINITY, f32::min);
    let max_x = corners
        .iter()
        .map(|(x, _)| *x)
        .fold(f32::NEG_INFINITY, f32::max);
    let min_y = corners
        .iter()
        .map(|(_, y)| *y)
        .fold(f32::INFINITY, f32::min);
    let max_y = corners
        .iter()
        .map(|(_, y)| *y)
        .fold(f32::NEG_INFINITY, f32::max);
    *x = min_x;
    *y = min_y;
    *width = max_x - min_x;
    *height = max_y - min_y;
}

#[cfg(test)]
mod tests {
    use super::*;
    use vsdx_render::Affine;

    #[test]
    fn flat_composes_group_transforms() {
        let mut primitives = vec![Primitive::Group {
            id: "group".into(),
            z_order: 0,
            transform: Affine {
                a: 0.0,
                b: 1.0,
                c: -1.0,
                d: 0.0,
                e: 4.0,
                f: 5.0,
            },
            primitives: vec![Primitive::TextBox {
                id: "text".into(),
                z_order: 0,
                x: 1.0,
                y: 2.0,
                width: 3.0,
                height: 4.0,
                paragraphs: Vec::new(),
                lines: Vec::new(),
                transform: Affine::identity(),
            }],
        }];
        let flattened = flat(&primitives);
        let [Primitive::TextBox { transform, .. }] = flattened.as_slice() else {
            panic!("text expected");
        };
        assert_eq!(transform.apply_point(1.0, 2.0), (2.0, 6.0));
        translate_primitives(&mut primitives, 0.5, 1.0);
        let flattened = flat(&primitives);
        let [Primitive::TextBox { transform, .. }] = flattened.as_slice() else {
            panic!("text expected");
        };
        assert_eq!(transform.apply_point(1.0, 2.0), (2.5, 7.0));
    }

    #[test]
    fn emf_uses_renderable_content_type() {
        assert_eq!(
            sniff_image("visio/media/image.emf", &[]),
            Some(("emf", "image/x-emf"))
        );
    }

    #[test]
    fn wmf_uses_renderable_content_type() {
        assert_eq!(
            sniff_image("visio/media/image.WMF", &[]),
            Some(("wmf", "image/x-wmf"))
        );
    }

    #[test]
    fn translate_moves_scene_positions() {
        let mut primitives = vec![Primitive::Placeholder {
            id: "box".into(),
            z_order: 0,
            x: 1.0,
            y: 2.0,
            width: 3.0,
            height: 4.0,
            reason: "test".into(),
        }];
        translate_primitives(&mut primitives, 0.5, 1.0);
        let [Primitive::Placeholder { x, y, .. }] = primitives.as_slice() else {
            panic!("box expected");
        };
        assert_eq!((*x, *y), (1.5, 3.0));
    }
}
