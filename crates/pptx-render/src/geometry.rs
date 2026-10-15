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
                | "cloud"
                | "cloudCallout"
                | "wedgeRectCallout"
                | "wedgeRoundRectCallout"
                | "ribbon2"
                | "swooshArrow"
                | "circularArrow"
                | "leftCircularArrow"
                | "curvedLeftArrow"
                | "curvedRightArrow"
                | "curvedUpArrow"
                | "uturnArrow"
                | "foldedCorner"
                | "ellipse"
                | "roundRect"
                | "round1Rect"
                | "round2SameRect"
                | "round2DiagRect"
                | "can"
                | "blockArc"
                | "pie"
                | "chord"
                | "flowChartDelay"
                | "flowChartOnlineStorage"
                | "gear6"
                | "gear9"
                | "moon"
                | "heart"
                | "leftBracket"
                | "rightBracket"
                | "teardrop"
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
    let (factor, offset) = match mode {
        PresetPathFill::DarkenLess => (0.8, 0.0),
        PresetPathFill::LightenLess => (0.8, 51.0),
        PresetPathFill::Lighten => (0.6, 102.0),
        _ => return,
    };
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
            .map(|value| (f64::from(value) * factor + offset).round() as u8)
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
    fn can_lightens_its_top_and_keeps_alpha() {
        let source = serde_json::from_value(serde_json::json!({
            "kind": "shape", "objectId": 1, "name": "can", "geometry": "can",
            "x": 0, "y": 0, "w": 200, "h": 100, "path": [],
            "fill": { "kind": "solid", "color": "#32649680" },
            "stroke": { "color": "#000000", "width": 1 }
        }))
        .unwrap();
        let parts = preset_primitives(source);
        assert_eq!(parts.len(), 3);
        for (index, expected) in [Some("#32649680"), Some("#84a2c080"), None]
            .into_iter()
            .enumerate()
        {
            let Primitive::Shape { fill, stroke, .. } = &parts[index].0 else {
                panic!()
            };
            match (fill, expected) {
                (Some(Paint::Solid { color }), Some(expected)) => assert_eq!(color, expected),
                (None, None) => {}
                _ => panic!("unexpected fill"),
            }
            assert_eq!(stroke.is_some(), index == 2);
            if let Some(stroke) = stroke {
                assert_eq!(stroke.join.as_deref(), Some("round"));
            }
        }
    }

    #[test]
    fn callout_leader_lines_are_stroke_only() {
        for (geometry, body_stroked) in [
            ("callout1", false),
            ("borderCallout1", true),
            ("accentCallout1", false),
            ("accentCallout2", false),
        ] {
            let source = serde_json::from_value(serde_json::json!({
                "kind": "shape", "objectId": 1, "name": geometry, "geometry": geometry,
                "x": 0, "y": 0, "w": 200, "h": 100, "path": [],
                "fill": { "kind": "solid", "color": "#326496" },
                "stroke": { "color": "#000000", "width": 1 }
            }))
            .unwrap();
            let parts = preset_primitives(source);
            assert_eq!(parts.len(), 2, "{geometry}");
            let Primitive::Shape {
                fill: body_fill,
                stroke: body_stroke,
                ..
            } = &parts[0].0
            else {
                panic!()
            };
            assert!(body_fill.is_some() && parts[0].1);
            assert_eq!(body_stroke.is_some(), body_stroked);

            let Primitive::Shape {
                fill: leader_fill,
                stroke: leader_stroke,
                ..
            } = &parts[1].0
            else {
                panic!()
            };
            assert!(leader_fill.is_none() && !parts[1].1);
            assert!(leader_stroke.is_some());
        }
    }

    #[test]
    fn layered_presets_cast_one_shadow_with_every_fill_and_open_stroke() {
        for geometry in [
            "arc",
            "cube",
            "leftBrace",
            "rightBrace",
            "ribbon2",
            "foldedCorner",
            "cloud",
            "cloudCallout",
            "can",
            "leftBracket",
            "rightBracket",
            "callout1",
            "borderCallout1",
            "accentCallout1",
            "accentCallout2",
            "curvedLeftArrow",
            "curvedRightArrow",
            "curvedUpArrow",
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
