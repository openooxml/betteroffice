use ooxml_drawingml::{PresetPathFill, preset_geometry_layers};

use crate::{Paint, Primitive};

pub(crate) fn preset_primitives(mut primitive: Primitive) -> Vec<(Primitive, bool)> {
    if let Primitive::Shape {
        geometry,
        stroke: Some(stroke),
        ..
    } = &mut primitive
        && stroke.join.is_none()
        && matches!(
            geometry.as_str(),
            "arc"
                | "leftBrace"
                | "rightBrace"
                | "donut"
                | "bentArrow"
                | "cloudCallout"
                | "wedgeRectCallout"
                | "wedgeRoundRectCallout"
                | "ribbon2"
                | "swooshArrow"
                | "circularArrow"
                | "foldedCorner"
                | "ellipse"
                | "roundRect"
                | "triangle"
                | "homePlate"
        )
    {
        stroke.join = Some("round".to_owned());
    }
    let Primitive::Shape {
        geometry,
        adjust_values,
        w,
        h,
        ..
    } = &primitive
    else {
        return vec![(primitive, true)];
    };
    let adjustments = adjust_values
        .iter()
        .map(|(key, value)| (key.clone(), f64::from(*value)))
        .collect();
    let Some(layers) =
        preset_geometry_layers(geometry, &adjustments, f64::from(*w) / f64::from(*h))
    else {
        return vec![(primitive, true)];
    };
    layers
        .into_iter()
        .map(|layer| {
            let mut part = primitive.clone();
            if let Primitive::Shape {
                path,
                fill,
                stroke,
                shadow,
                geometry_fallback,
                ..
            } = &mut part
            {
                *path = layer.commands;
                *geometry_fallback = false;
                if layer.fill == PresetPathFill::None {
                    *fill = None;
                } else if let Some(fill) = fill {
                    shade_fill(fill, layer.fill);
                }
                if !layer.stroke {
                    *stroke = None;
                }
                if fill.is_none() && stroke.is_none() {
                    *shadow = None;
                }
            }
            (part, layer.fill != PresetPathFill::None)
        })
        .collect()
}

fn shade_fill(paint: &mut Paint, mode: PresetPathFill) {
    if !matches!(
        mode,
        PresetPathFill::DarkenLess | PresetPathFill::LightenLess
    ) {
        return;
    }
    let shade = |color: &mut String| {
        let Some(hex) = color
            .strip_prefix('#')
            .filter(|hex| hex.is_ascii() && (hex.len() == 6 || hex.len() == 8))
        else {
            return;
        };
        let channels: Option<Vec<_>> = (0..3)
            .map(|index| u8::from_str_radix(&hex[index * 2..index * 2 + 2], 16).ok())
            .collect();
        let Some(channels) = channels else {
            return;
        };
        let adjusted: Vec<_> = channels
            .into_iter()
            .map(|value| {
                let value = f64::from(value) * 0.8;
                (value
                    + if mode == PresetPathFill::LightenLess {
                        51.0
                    } else {
                        0.0
                    })
                .round() as u8
            })
            .collect();
        *color = format!(
            "#{:02x}{:02x}{:02x}{}",
            adjusted[0],
            adjusted[1],
            adjusted[2],
            &hex[6..]
        );
    };
    match paint {
        Paint::Solid { color } => shade(color),
        Paint::Gradient { stops, .. } => {
            for stop in stops {
                shade(&mut stop.color);
            }
        }
    }
}
