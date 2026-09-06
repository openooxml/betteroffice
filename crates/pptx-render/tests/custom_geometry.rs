use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::DeckSession;
use pptx_parse::ShapeNode;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/custom-geometry.pptx");

#[test]
fn custom_paths_keep_coordinates_paints_and_fallbacks() {
    let session = DeckSession::open(FIXTURE, 285).unwrap();
    let renderer = SlideRenderer::new();
    let deck = session.snapshot().unwrap();
    let rendered = renderer.layout_slide(session.package(), &deck, 0).unwrap();
    let primitives = &rendered.display_list.primitives;
    assert_eq!(primitives.len(), 8);
    for (index, primitive) in primitives[..4].iter().enumerate() {
        let Primitive::Shape {
            x,
            y,
            w,
            h,
            path,
            fill,
            stroke,
            shape_id,
            ..
        } = primitive
        else {
            panic!()
        };
        assert_eq!((*x, *y, *w, *h), (20.0, 20.0, 200.0, 100.0));
        assert_eq!(
            shape_id.as_deref(),
            Some(deck.slides[0].shapes[0].id.as_str())
        );
        assert_eq!(fill.is_none(), index == 1 || index == 3);
        assert_eq!(stroke.is_none(), index == 2 || index == 3);
        if let Some(fill) = fill {
            assert_eq!(
                *fill,
                Paint::Solid {
                    color: "#DC2626".into()
                }
            );
        }
        if let Some(stroke) = stroke {
            assert_eq!(stroke.color, "#2563EB");
            assert_eq!(stroke.width, 2.0);
        }
        let expected = match index {
            0 => vec![
                C::Move { x: 0.1, y: 0.1 },
                C::Line { x: 0.9, y: 0.1 },
                C::Cubic {
                    cp1x: 0.9,
                    cp1y: 0.4,
                    cp2x: 0.6,
                    cp2y: 1.0,
                    x: 0.5,
                    y: 0.8,
                },
                C::Quad {
                    cpx: 0.2,
                    cpy: 1.0,
                    x: 0.1,
                    y: 0.5,
                },
                C::Close,
            ],
            1 => vec![C::Move { x: 0.1, y: 0.5 }, C::Line { x: 0.9, y: 0.5 }],
            2 => vec![
                C::Move { x: 0.5, y: 0.1 },
                C::Line { x: 0.9, y: 0.9 },
                C::Line { x: 0.1, y: 0.9 },
                C::Close,
            ],
            _ => vec![C::Move { x: 0.0, y: 0.0 }, C::Line { x: 1.0, y: 1.0 }],
        };
        assert_eq!(*path, expected);
    }
    let Primitive::Shape { path, .. } = &primitives[4] else {
        panic!()
    };
    assert_eq!(path.len(), 4);
    assert!(
        matches!(path[1], C::Cubic { x, y, .. } if (x - 0.5).abs() < 1e-12 && (y - 0.5).abs() < 1e-12)
    );
    let rectangle = vec![
        C::Move { x: 0.0, y: 0.0 },
        C::Line { x: 1.0, y: 0.0 },
        C::Line { x: 1.0, y: 1.0 },
        C::Line { x: 0.0, y: 1.0 },
        C::Close,
    ];
    for primitive in &primitives[5..7] {
        let Primitive::Shape { path, fill, .. } = primitive else {
            panic!()
        };
        assert_eq!(*path, rectangle);
        assert!(fill.is_some());
    }
    let preset =
        ooxml_drawingml::preset_geometry_to_path("ellipse", &Default::default(), 100.0 / 70.0)
            .unwrap();
    let Primitive::Shape { path, .. } = &primitives[7] else {
        panic!()
    };
    assert_eq!(*path, preset);

    let mut deck = deck;
    deck.slides[0].shapes[0].geometry = "ellipse".into();
    let rendered = renderer.layout_slide(session.package(), &deck, 0).unwrap();
    let Primitive::Shape { path, .. } = &rendered.display_list.primitives[0] else {
        panic!()
    };
    assert_eq!(
        *path,
        ooxml_drawingml::preset_geometry_to_path("ellipse", &Default::default(), 2.0).unwrap()
    );
}

#[test]
fn layout_and_master_paths_survive_snapshot_hydration() {
    let custom = pptx_parse::parse_pptx(FIXTURE)
        .unwrap()
        .slides
        .remove(0)
        .shapes
        .remove(0);
    let demo = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    let mut package = pptx_parse::parse_pptx(demo).unwrap();
    package.slides[0].shapes.clear();
    package.slides[0].layout_part_path = Some(package.layouts[0].part_path.clone());
    package.layouts[0].master_part_path = Some(package.masters[0].part_path.clone());
    package.layouts[0].shapes = vec![custom.clone()];
    package.masters[0].shapes = vec![custom];
    let session = DeckSession::from_package(package, 285).unwrap();
    let renderer = SlideRenderer::new();
    let before = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(before.primitives.len(), 8);
    for primitives in before.primitives.chunks(4) {
        let Primitive::Shape { path, .. } = &primitives[0] else {
            panic!()
        };
        assert_eq!(path[0], C::Move { x: 0.1, y: 0.1 });
        assert!(matches!(
            &primitives[1],
            Primitive::Shape {
                fill: None,
                stroke: Some(_),
                ..
            }
        ));
        assert!(matches!(
            &primitives[2],
            Primitive::Shape {
                fill: Some(_),
                stroke: None,
                ..
            }
        ));
    }
    let restored =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 286).unwrap();
    let after = renderer
        .layout_slide(restored.package(), &restored.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(before, after);
    let ShapeNode::Shape(shape) = &restored.package().layouts[0].shapes[0] else {
        panic!()
    };
    assert_eq!(shape.paths.len(), 4);
}

#[test]
fn custom_paths_share_the_slide_shape_budget() {
    let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[0] else {
        panic!()
    };
    shape.paths.resize(20_001, shape.paths[0].clone());
    let session = DeckSession::from_package(package, 285).unwrap();
    assert!(matches!(
        SlideRenderer::new().layout_slide(session.package(), &session.snapshot().unwrap(), 0),
        Err(pptx_render::RenderError::ResourceLimit(_))
    ));
}
