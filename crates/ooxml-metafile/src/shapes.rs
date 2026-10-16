//! The strict vector replay PowerPoint pictures draw from.

use crate::drawing::{Op, Paint, PathCommand};
use crate::player::{Player, axis_aligned_rect};

#[derive(Debug, Clone, PartialEq)]
pub struct MetafileOp {
    pub path: Vec<PathCommand>,
    pub fill: Option<String>,
    pub stroke: Option<MetafileStroke>,
    pub even_odd: bool,
    /// Clip rectangle in frame coordinates, as `[left, top, right, bottom]`.
    pub clip: Option<[f64; 4]>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct MetafileStroke {
    pub color: String,
    pub width: f64,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct MetafileDrawing {
    pub ops: Vec<MetafileOp>,
}

/// Solid fills and strokes in fractions of the metafile's frame, or `None`
/// when the metafile holds anything else.
pub fn decode(bytes: &[u8]) -> Option<MetafileDrawing> {
    let player = crate::play_emf::<false>(bytes, 0, false, None)
        .or_else(|_| crate::play_wmf::<false>(bytes, 0, None))
        .ok()?;
    let drawing = finish(player)?;
    (!drawing.ops.is_empty()).then_some(drawing)
}

fn finish(player: Player<false>) -> Option<MetafileDrawing> {
    let solid = |paint: Paint| match paint {
        Paint::Solid(color) => Some(color.hex()),
        _ => None,
    };
    let mut ops = Vec::with_capacity(player.ops.len());
    for op in player.ops {
        let Op::Shape(shape) = op else {
            return None;
        };
        let clip = match shape.clip.as_deref() {
            None => None,
            Some(chain) if chain.parent.is_none() && !chain.region.exclude => {
                Some(axis_aligned_rect(&chain.region.path)?)
            }
            Some(_) => return None,
        };
        let stroke = match shape.stroke {
            Some(stroke) => Some(MetafileStroke {
                color: solid(stroke.paint)?,
                width: stroke.width,
            }),
            None => None,
        };
        let fill = match shape.fill {
            Some(fill) => Some(solid(fill)?),
            None => None,
        };
        ops.push(MetafileOp {
            path: shape.path,
            fill,
            stroke,
            even_odd: shape.even_odd,
            clip,
        });
    }
    Some(MetafileDrawing { ops })
}

#[cfg(test)]
#[path = "shapes_tests.rs"]
mod tests;
