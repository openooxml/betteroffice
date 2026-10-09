//! Places a nested metafile's drawing inside the picture that embeds it.

use std::collections::HashMap;
use std::sync::Arc;

use crate::drawing::{Clip, ClipChain, LinearGradient, Op, Paint, PathCommand, Stroke};
use crate::player::{Xform, apply, chain, concat};

/// `op`, drawn in its own drawing units, mapped by `m` and clipped by `outer` too.
pub(crate) fn op(
    op: Op,
    m: Xform,
    outer: &Clip,
    cache: &mut HashMap<*const ClipChain, Clip>,
) -> Op {
    let scale = (m[0] * m[3] - m[1] * m[2]).abs().sqrt();
    match op {
        Op::Shape(mut shape) => {
            path(&mut shape.path, m);
            shape.fill = shape.fill.map(|fill| paint(fill, m, scale));
            shape.stroke = shape.stroke.map(|stroke| Stroke {
                paint: paint(stroke.paint, m, scale),
                width: stroke.width * scale,
                dash: stroke
                    .dash
                    .map(|dash| dash.into_iter().map(|length| length * scale).collect()),
                ..stroke
            });
            shape.clip = clip(&shape.clip, m, outer, cache);
            Op::Shape(shape)
        }
        Op::Text(mut text) => {
            text.transform = concat(text.transform, m);
            text.fill = paint(text.fill, m, scale);
            text.clip = clip(&text.clip, m, outer, cache);
            Op::Text(text)
        }
        Op::Image(mut image) => {
            image.transform = concat(image.transform, m);
            image.clip = clip(&image.clip, m, outer, cache);
            Op::Image(image)
        }
    }
}

fn path(path: &mut [PathCommand], m: Xform) {
    let map = |x: &mut f64, y: &mut f64| {
        let (px, py) = apply(m, (*x, *y));
        *x = px;
        *y = py;
    };
    for command in path {
        match command {
            PathCommand::Move { x, y } | PathCommand::Line { x, y } => map(x, y),
            PathCommand::Quad { cpx, cpy, x, y } => {
                map(cpx, cpy);
                map(x, y);
            }
            PathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                map(cp1x, cp1y);
                map(cp2x, cp2y);
                map(x, y);
            }
            PathCommand::Close => {}
        }
    }
}

fn paint(paint: Paint, m: Xform, scale: f64) -> Paint {
    match paint {
        Paint::Solid(_) => paint,
        Paint::Hatch {
            style,
            color,
            background,
            cell,
        } => Paint::Hatch {
            style,
            color,
            background,
            cell: cell * scale,
        },
        Paint::Pattern {
            tile,
            width,
            height,
        } => Paint::Pattern {
            tile,
            width: width * scale,
            height: height * scale,
        },
        Paint::Linear(gradient) => Paint::Linear(Arc::new(linear(&gradient, m))),
    }
}

/// `gradient` mapped by `m` with every point keeping its colour: the start
/// maps with `m`, and the gradient vector with its inverse transpose, which
/// keeps the colour bands parallel under shear and unequal scaling.
pub(crate) fn linear(gradient: &LinearGradient, m: Xform) -> LinearGradient {
    let start = apply(m, gradient.start);
    let (dx, dy) = (
        gradient.end.0 - gradient.start.0,
        gradient.end.1 - gradient.start.1,
    );
    let scale = (m[0] * m[3] - m[1] * m[2]) * (dx * dx + dy * dy);
    let normal = (
        (m[3] * dx - m[1] * dy) / scale,
        (m[0] * dy - m[2] * dx) / scale,
    );
    let length = normal.0 * normal.0 + normal.1 * normal.1;
    let end = if length.is_finite() && length > 0.0 {
        (start.0 + normal.0 / length, start.1 + normal.1 / length)
    } else {
        apply(m, gradient.end)
    };
    LinearGradient {
        start,
        end,
        ..gradient.clone()
    }
}

fn clip(clip: &Clip, m: Xform, outer: &Clip, cache: &mut HashMap<*const ClipChain, Clip>) -> Clip {
    let Some(link) = clip else {
        return outer.clone();
    };
    if let Some(done) = cache.get(&Arc::as_ptr(link)) {
        return done.clone();
    }
    let parent = self::clip(&link.parent, m, outer, cache);
    let mut region = link.region.clone();
    path(&mut region.path, m);
    let placed = chain(parent, region);
    cache.insert(Arc::as_ptr(link), placed.clone());
    placed
}
