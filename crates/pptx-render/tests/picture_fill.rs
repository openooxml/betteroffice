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
    let restored =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 287).unwrap();
    let after = renderer
        .layout_slide(restored.package(), &restored.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(before, after);
}
