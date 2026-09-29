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
        Paint::Linear(gradient) => Paint::Linear(Arc::new(LinearGradient {
            start: apply(m, gradient.start),
            end: apply(m, gradient.end),
            ..(*gradient).clone()
        })),
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
