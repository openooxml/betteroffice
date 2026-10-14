use std::collections::BTreeSet;

use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer, SurfaceDisplayList};

const FIXTURE: &[u8] = include_bytes!("fixtures/prstgeom.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render() -> SurfaceDisplayList {
    let session = DeckSession::open(FIXTURE, 600).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer
        .register_font("Calibri", false, false, FONT)
        .unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

#[test]
fn issue_600_presets_render_without_rectangle_substitutions() {
    let rendered = render();
    let mut geometries = BTreeSet::new();
    for primitive in &rendered.primitives {
        if let Primitive::Shape {
            geometry,
            geometry_fallback,
            path,
            ..
        } = primitive
        {
            assert!(!geometry_fallback, "{geometry}");
            if geometry != "rect" {
                geometries.insert(geometry.as_str());
                assert!(!path.is_empty(), "{geometry}");
            }
        }
    }
    assert_eq!(
        geometries,
        BTreeSet::from([
            "arc",
            "cube",
            "rightBrace",
            "leftBrace",
            "donut",
            "bentArrow",
            "cloudCallout",
            "wedgeRectCallout",
            "wedgeRoundRectCallout",
            "ribbon2",
            "swooshArrow",
            "circularArrow",
            "foldedCorner",
            "ellipse",
            "roundRect",
            "chevron",
            "triangle",
            "homePlate",
        ])
    );
}

#[test]
fn arcs_and_braces_fill_closed_regions_but_stroke_open_outlines() {
    let rendered = render();
    for kind in ["arc", "rightBrace", "leftBrace"] {
        let layers: Vec<_> = rendered.primitives.iter().filter(|primitive| {
            matches!(primitive, Primitive::Shape { geometry, .. } if geometry == kind)
        }).collect();
        assert_eq!(layers.len(), 2, "{kind}");
        let Primitive::Shape {
            path, fill, stroke, ..
        } = layers[0]
        else {
            panic!()
        };
        assert!(fill.is_some());
        assert!(stroke.is_none());
        assert_eq!(path.last(), Some(&C::Close));
        let Primitive::Shape {
            path, fill, stroke, ..
        } = layers[1]
        else {
            panic!()
        };
        assert!(fill.is_none());
        assert!(stroke.is_some());
        assert!(!path.contains(&C::Close));
    }
}

#[test]
fn cubes_and_folds_keep_separately_shaded_faces() {
    let rendered = render();
    for (kind, expected) in [
        (
            "cube",
            vec![Some("#DCE9F7"), Some("#b0bac6"), Some("#e3edf9"), None],
        ),
        ("foldedCorner", vec![Some("#DCE9F7"), Some("#b0bac6"), None]),
        ("ribbon2", vec![Some("#DCE9F7"), Some("#b0bac6"), None]),
    ] {
        let layers: Vec<_> = rendered
            .primitives
            .iter()
            .filter_map(|primitive| {
                if let Primitive::Shape {
                    geometry,
                    fill,
                    stroke,
                    ..
                } = primitive
                    && geometry == kind
                {
                    return Some((fill, stroke));
                }
                None
            })
            .collect();
        assert_eq!(layers.len(), expected.len());
        for ((fill, stroke), color) in layers.iter().zip(expected) {
            assert_eq!(
                fill.as_ref().map(|paint| match paint {
                    Paint::Solid { color } => color.as_str(),
                    _ => panic!(),
                }),
                color
            );
            assert_eq!(stroke.is_some(), color.is_none());
        }
    }
}

#[test]
fn composed_arcs_keep_adjustments_and_transforms_on_both_paths() {
    let json = r##"{"widthPx":500,"heightPx":500,"shapes":[{"kind":"shape","id":7,"name":"Arc","rect":{"x":10,"y":20,"w":282,"h":306},"rotationDeg":-90,"geometry":"arc","adjustValues":{"adj1":173.68684,"adj2":207.9893},"fill":{"kind":"solid","color":"#abcdef"}}]}"##;
    let rendered: SurfaceDisplayList =
        serde_json::from_str(&pptx_render::compile_json(json).unwrap()).unwrap();
    assert_eq!(rendered.primitives.len(), 2);
    for primitive in rendered.primitives {
        let Primitive::Shape {
            x,
            y,
            w,
            h,
            path,
            transform,
            geometry_fallback,
            ..
        } = primitive
        else {
            panic!()
        };
        assert_eq!((x, y, w, h), (10.0, 20.0, 282.0, 306.0));
        assert_eq!(transform.rotation_deg, -90.0);
        assert!(!geometry_fallback);
        assert!(
            path.iter()
                .any(|command| matches!(command, C::Cubic { .. }))
        );
    }
}

#[test]
fn preset_outlines_match_powerpoints_export_within_a_tenth_of_a_pixel() {
    let reference: serde_json::Value =
        serde_json::from_str(include_str!("fixtures/prstgeom-powerpoint.json")).unwrap();
    let outlines: std::collections::BTreeMap<String, Vec<C>> =
        serde_json::from_value(reference["outlines"].clone()).unwrap();
    let rendered = render();
    for (name, expected) in outlines {
        let expected = segments(&expected, (0.0, 0.0, 1.0, 1.0));
        let actual: Vec<_> = rendered
            .primitives
            .iter()
            .flat_map(|primitive| match primitive {
                Primitive::Shape {
                    geometry,
                    path,
                    x,
                    y,
                    w,
                    h,
                    stroke: Some(stroke),
                    ..
                } if geometry == &name => {
                    assert_eq!(
                        stroke.join.as_deref().unwrap_or("miter"),
                        reference["joins"][&name].as_str().unwrap(),
                        "{name}"
                    );
                    segments(
                        path,
                        (f64::from(*x), f64::from(*y), f64::from(*w), f64::from(*h)),
                    )
                }
                _ => Vec::new(),
            })
            .collect();
        assert!(!actual.is_empty(), "{name}");
        for (source, target) in [(&actual, &expected), (&expected, &actual)] {
            let distance = source
                .iter()
                .flat_map(|(a, b)| [a, b])
                .map(|point| {
                    target
                        .iter()
                        .map(|segment| distance_to_segment(*point, *segment))
                        .fold(f64::INFINITY, f64::min)
                })
                .fold(0.0, f64::max);
            assert!(distance < 0.1, "{name}: outline differs by {distance} px");
        }
    }
}

type Point = (f64, f64);
type Segment = (Point, Point);

fn segments(path: &[C], (x, y, w, h): (f64, f64, f64, f64)) -> Vec<Segment> {
    let map = |(px, py): Point| (x + px * w, y + py * h);
    let mut result = Vec::new();
    let mut cursor = (0.0, 0.0);
    let mut start = cursor;
    for command in path {
        let mut points = Vec::new();
        match *command {
            C::Move { x, y } => {
                cursor = (x, y);
                start = cursor;
            }
            C::Line { x, y } => points.push((x, y)),
            C::Close => points.push(start),
            C::Quad { cpx, cpy, x, y } => {
                for step in 1..=64 {
                    let t = f64::from(step) / 64.0;
                    let u = 1.0 - t;
                    points.push((
                        u * u * cursor.0 + 2.0 * u * t * cpx + t * t * x,
                        u * u * cursor.1 + 2.0 * u * t * cpy + t * t * y,
                    ));
                }
            }
            C::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                for step in 1..=64 {
                    let t = f64::from(step) / 64.0;
                    let u = 1.0 - t;
                    points.push((
                        u * u * u * cursor.0
                            + 3.0 * u * u * t * cp1x
                            + 3.0 * u * t * t * cp2x
                            + t * t * t * x,
                        u * u * u * cursor.1
                            + 3.0 * u * u * t * cp1y
                            + 3.0 * u * t * t * cp2y
                            + t * t * t * y,
                    ));
                }
            }
        }
        for point in points {
            result.push((map(cursor), map(point)));
            cursor = point;
        }
    }
    result
}

fn distance_to_segment(point: Point, (a, b): Segment) -> f64 {
    let delta = (b.0 - a.0, b.1 - a.1);
    let length = delta.0 * delta.0 + delta.1 * delta.1;
    let t = if length > 0.0 {
        (((point.0 - a.0) * delta.0 + (point.1 - a.1) * delta.1) / length).clamp(0.0, 1.0)
    } else {
        0.0
    };
    (point.0 - a.0 - t * delta.0).hypot(point.1 - a.1 - t * delta.1)
}
