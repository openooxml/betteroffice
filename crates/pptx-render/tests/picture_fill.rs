use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/picture-fill.pptx");

#[test]
fn a_picture_filled_shape_paints_its_blip_through_its_own_outline() {
    let session = DeckSession::open(FIXTURE, 286).unwrap();
    let renderer = SlideRenderer::new();
    let deck = session.snapshot().unwrap();
    let rendered = renderer.layout_slide(session.package(), &deck, 0).unwrap();
    let primitives = &rendered.display_list.primitives;
    assert_eq!(primitives.len(), 3);

    let Primitive::Image {
        x,
        y,
        w,
        h,
        asset_id,
        crop,
        path,
        stroke,
        ..
    } = &primitives[0]
    else {
        panic!("the freeform paints its picture fill")
    };
    assert_eq!((*x, *y, *w, *h), (20.0, 20.0, 200.0, 200.0));
    assert_eq!(asset_id.as_deref(), Some("ppt/media/image1.png"));
    assert!(crop.is_whole());
    assert!(stroke.is_none());
    assert_eq!(
        path.as_deref(),
        Some(
            [
                C::Move { x: 0.0, y: 0.0 },
                C::Line { x: 1.0, y: 0.0 },
                C::Line { x: 1.0, y: 1.0 },
                C::Line { x: 0.0, y: 1.0 },
                C::Close,
                C::Move { x: 0.25, y: 0.25 },
                C::Line { x: 0.25, y: 0.75 },
                C::Line { x: 0.75, y: 0.75 },
                C::Line { x: 0.75, y: 0.25 },
                C::Close,
            ]
            .as_slice()
        )
    );

    let Primitive::Image {
        x,
        crop,
        path,
        stroke,
        ..
    } = &primitives[1]
    else {
        panic!("the ellipse paints its picture fill")
    };
    assert!((*x - 250.0).abs() < 0.01);
    assert!((crop.left - (0.1 + 0.7 / 3.0)).abs() < 1e-6);
    assert!((crop.right - 0.2).abs() < 1e-6);
    assert_eq!((crop.top, crop.bottom), (0.0, 0.0));
    assert_eq!(
        stroke.as_ref().map(|stroke| stroke.color.as_str()),
        Some("#2563EB")
    );
    assert_eq!(
        path.as_deref(),
        Some(
            ooxml_drawingml::preset_geometry_to_path("ellipse", &Default::default(), 1.0)
                .unwrap()
                .as_slice()
        )
    );

    assert!(matches!(
        &primitives[2],
        Primitive::Shape {
            fill: None,
            stroke: None,
            ..
        }
    ));
}

#[test]
fn a_picture_fill_survives_a_snapshot_round_trip() {
    let session = DeckSession::open(FIXTURE, 286).unwrap();
    let renderer = SlideRenderer::new();
    let before = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(
        before
            .primitives
            .iter()
            .filter(|primitive| matches!(primitive, Primitive::Image { .. }))
            .count(),
        2
    );
    let restored =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 287).unwrap();
    let after = renderer
        .layout_slide(restored.package(), &restored.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(before, after);
}

#[test]
fn an_explicit_tiled_fill_blocks_a_placeholders_inherited_picture() {
    let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    let demo = pptx_parse::parse_pptx(include_bytes!(
        "../../../apps/demo/public/betteroffice-demo.pptx"
    ))
    .unwrap();
    let pptx_parse::ShapeNode::Shape(mut inherited) = package.slides[0].shapes[0].clone() else {
        panic!("expected ring")
    };
    let pptx_parse::ShapeNode::Shape(mut tiled) = package.slides[0].shapes[2].clone() else {
        panic!("expected tiled shape")
    };
    let placeholder = pptx_parse::Placeholder {
        placeholder_type: Some("body".to_owned()),
        index: Some(33),
        orientation: None,
        size: None,
    };
    inherited.base.placeholder = Some(placeholder.clone());
    tiled.base.placeholder = Some(placeholder);
    let mut layout = demo.layouts[0].clone();
    layout.shapes = vec![pptx_parse::ShapeNode::Shape(inherited)];
    package.slides[0].layout_part_path = Some(layout.part_path.clone());
    package.layouts = vec![layout];
    for explicit in [false, true] {
        let mut shape = tiled.clone();
        if !explicit {
            shape.fill = None;
        }
        package.slides[0].shapes = vec![pptx_parse::ShapeNode::Shape(shape)];
        let session =
            DeckSession::from_package_with_source(package.clone(), FIXTURE, 33610).unwrap();
        let rendered = SlideRenderer::new()
            .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
            .unwrap();
        assert_eq!(rendered.display_list.primitives.len(), 1);
        assert_eq!(
            matches!(rendered.display_list.primitives[0], Primitive::Image { .. }),
            !explicit
        );
    }
}
