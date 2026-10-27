#[path = "../../pptx-parse/tests/common/ole_vml.rs"]
mod fixture;

use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

#[test]
fn a_legacy_ole_frame_paints_its_vml_preview_without_a_placeholder() {
    let bytes = fixture::deck(fixture::VML, fixture::OBJECT);
    let session = DeckSession::open(&bytes, 319).unwrap();
    let rendered = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let images: Vec<_> = rendered
        .display_list
        .primitives
        .iter()
        .filter(|primitive| matches!(primitive, Primitive::Image { .. }))
        .collect();
    assert_eq!(images.len(), 1);
    let Primitive::Image {
        asset_id,
        x,
        y,
        w,
        h,
        crop,
        ..
    } = images[0]
    else {
        unreachable!();
    };
    assert_eq!(asset_id.as_deref(), Some(fixture::MEDIA));
    for (actual, expected) in [
        (*x, 5868144.0),
        (*y, 4365104.0),
        (*w, 1152128.0),
        (*h, 747326.0),
    ] {
        assert!((actual - expected / 9525.0).abs() < 0.01);
    }
    assert!(crop.is_whole());
    assert!(
        !rendered
            .display_list
            .primitives
            .iter()
            .any(|primitive| matches!(primitive, Primitive::Placeholder { .. }))
    );
}
