//! `a:effectLst/a:outerShdw` reaches the display list, and only for shapes that
//! have something to cast a shadow.

use pptx_edit::DeckSession;
use pptx_render::{Primitive, Shadow, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/outer-shadow.pptx");
const PICTURE_FIXTURE: &[u8] = include_bytes!("fixtures/outer-shadow-picture.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn primitives(fixture: &[u8]) -> Vec<Primitive> {
    let session = DeckSession::open(fixture, 31).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
        .display_list
        .primitives
}

fn shadows() -> Vec<(u32, Option<Shadow>)> {
    primitives(FIXTURE)
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                object_id, shadow, ..
            } => Some((*object_id, shadow.clone())),
            _ => None,
        })
        .collect()
}

fn image_shadows() -> Vec<(u32, Option<Shadow>)> {
    primitives(PICTURE_FIXTURE)
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Image {
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

#[test]
fn a_picture_carries_its_shadow_into_the_display_list() {
    let shadows = image_shadows();
    let (_, shadow) = shadows
        .iter()
        .find(|(object_id, _)| *object_id == 4)
        .expect("the shadowed mark is drawn");
    let shadow = shadow.as_ref().expect("the shadowed mark casts a shadow");
    assert_eq!(shadow.color, "#00000066");
    assert!((shadow.blur - 8.0).abs() < 0.01);
    assert!((shadow.dx - 2.828).abs() < 0.01);
    assert!((shadow.dy - 2.828).abs() < 0.01);
}

#[test]
fn a_picture_without_an_effect_list_casts_nothing() {
    let shadows = image_shadows();
    let (_, shadow) = shadows
        .iter()
        .find(|(object_id, _)| *object_id == 5)
        .expect("the plain mark is drawn");
    assert_eq!(*shadow, None);
}

#[test]
fn a_picture_filled_shape_keeps_its_shadow_through_the_image_rewrite() {
    let shadows = image_shadows();
    let (_, shadow) = shadows
        .iter()
        .find(|(object_id, _)| *object_id == 6)
        .expect("the picture-filled card is drawn as an image");
    let shadow = shadow
        .as_ref()
        .expect("the picture-filled card keeps its shadow");
    assert_eq!(shadow.color, "#00000066");
    assert!((shadow.blur - 8.0).abs() < 0.01);
    assert!((shadow.dx - 2.828).abs() < 0.01);
    assert!((shadow.dy - 2.828).abs() < 0.01);
}

#[test]
fn an_ole_fallback_picture_resolves_its_shadow_in_the_frame_space() {
    let shadows = image_shadows();
    let (_, shadow) = shadows
        .iter()
        .find(|(object_id, _)| *object_id == 7)
        .expect("the OLE fallback picture is drawn");
    let shadow = shadow
        .as_ref()
        .expect("the OLE fallback picture casts a shadow");
    assert_eq!(shadow.color, "#00000066");
    assert!((shadow.blur - 8.0).abs() < 0.01);
    assert!((shadow.dx - 2.828).abs() < 0.01);
    assert!((shadow.dy - 2.828).abs() < 0.01);
}
