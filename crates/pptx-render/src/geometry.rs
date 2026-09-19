use ooxml_drawingml::{PresetPathFill, preset_geometry_layers};

use crate::{Paint, Primitive, ShadowPath};

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
    let shadow_paths = if let Primitive::Shape {
        shadow: Some(_),
        stroke,
        ..
    } = &primitive
        && layers.len() > 1
    {
        layers
            .iter()
            .map(|layer| ShadowPath {
                path: layer.commands.clone(),
                fill: layer.fill != PresetPathFill::None,
                stroke: stroke.clone().filter(|_| layer.stroke),
            })
            .collect()
    } else {
        Vec::new()
    };
    layers
        .into_iter()
        .enumerate()
        .map(|(index, layer)| {
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
                if index > 0 {
                    *shadow = None;
                } else if let Some(shadow) = shadow {
                    shadow.paths.clone_from(&shadow_paths);
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn layered_presets_cast_one_shadow_with_every_fill_and_open_stroke() {
        for geometry in [
            "arc",
            "cube",
            "leftBrace",
            "rightBrace",
            "ribbon2",
            "foldedCorner",
            "cloudCallout",
        ] {
            let source = serde_json::from_value(serde_json::json!({
                "kind": "shape", "objectId": 1, "name": "shadow", "geometry": geometry,
                "x": 10, "y": 20, "w": 100, "h": 70, "path": [],
                "fill": { "kind": "solid", "color": "#DCE9F780" },
                "stroke": { "color": "#000000", "width": 1 },
                "shadow": { "color": "#00000066", "blur": 8, "dx": 10 }
            }))
            .unwrap();
            let parts = preset_primitives(source);
            let Primitive::Shape {
                shadow: Some(shadow),
                ..
            } = &parts[0].0
            else {
                panic!("{geometry}")
            };
            assert_eq!(shadow.paths.len(), parts.len(), "{geometry}");
            for (index, (mask, (part, has_fill))) in shadow.paths.iter().zip(&parts).enumerate() {
                let Primitive::Shape {
                    path,
                    stroke,
                    shadow,
                    ..
                } = part
                else {
                    panic!()
                };
                assert_eq!(&mask.path, path);
                assert_eq!(&mask.stroke, stroke);
                assert_eq!(mask.fill, *has_fill);
                assert_eq!(shadow.is_some(), index == 0);
            }
        }
    }
}
