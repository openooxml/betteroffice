//! `a:effectLst/a:outerShdw` reaches the display list, and only for shapes that
//! have something to cast a shadow.

use pptx_edit::DeckSession;
use pptx_render::{Primitive, Shadow, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/outer-shadow.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn shadows() -> Vec<(u32, Option<Shadow>)> {
    let session = DeckSession::open(FIXTURE, 31).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap();
    rendered
        .display_list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                object_id, shadow, ..
            } => Some((*object_id, shadow.clone())),
            _ => None,
        })
        .collect()
}

#[test]
fn a_filled_shape_carries_its_shadow_into_the_display_list() {
    let shadows = shadows();
    let (_, shadow) = shadows
        .iter()
        .find(|(object_id, _)| *object_id == 2)
        .expect("the filled card is drawn");
    let shadow = shadow.as_ref().expect("the filled card casts a shadow");
    assert_eq!(shadow.color, "#00000066");
    assert!((shadow.blur - 8.0).abs() < 0.01);
    assert!((shadow.dx - 2.828).abs() < 0.01);
    assert!((shadow.dy - 2.828).abs() < 0.01);
}

#[test]
fn an_unfilled_shape_casts_its_outline_shadow() {
    let shadows = shadows();
    let (_, shadow) = shadows
        .iter()
        .find(|(object_id, _)| *object_id == 3)
        .expect("the unfilled card is drawn");
    assert_eq!(shadow.as_ref().unwrap().color, "#00000066");
}
