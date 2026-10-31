use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/metafile-pictures.pptx");

#[test]
fn emf_wmf_and_ole_pictures_keep_their_frames_and_source_crop() {
    let session = DeckSession::open(FIXTURE, 318).unwrap();
    let rendered = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    assert_eq!(rendered.display_list.primitives.len(), 4);
    for (index, primitive) in rendered.display_list.primitives.iter().enumerate() {
        let Primitive::Shape {
            x,
            y,
            w,
            h,
            path,
            fill,
            clip,
            even_odd,
            ..
        } = primitive
        else {
            panic!("picture did not become vector artwork: {primitive:?}");
        };
        assert_eq!(
            (*x, *y, *w, *h),
            (96.0 + index as f32 * 120.0, 96.0, 96.0, 96.0)
        );
        assert_eq!(
            fill,
            &Some(Paint::Solid {
                color: "#ff0000".into()
            })
        );
        assert!(*even_odd);
        assert_eq!(clip.as_ref().unwrap().len(), 5);
        let expected = if index == 3 { 1.0 / 6.0 } else { 0.0 };
        assert!(matches!(path[0], C::Move { x, y: 0.0 } if (x - expected).abs() < 1e-6));
    }
}

#[test]
fn full_replay_pictures_resolve_lazily_without_changing_saved_bytes() {
    let fixture = include_bytes!("fixtures/metafile-svg-pictures.pptx");
    let session = DeckSession::open(fixture, 319).unwrap();
    let saved = session.save().unwrap();
    let state = session.encode_state_as_update_v1();
    let rendered = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    assert_eq!(rendered.display_list.primitives.len(), 2);
    for (index, primitive) in rendered.display_list.primitives.iter().enumerate() {
        let Primitive::Image {
            asset_id: Some(part),
            x,
            y,
            w,
            h,
            ..
        } = primitive
        else {
            panic!("expected an unresolved picture: {primitive:?}");
        };
        assert_eq!(
            (*x, *y, *w, *h),
            (96.0 + index as f32 * 120.0, 96.0, 96.0, 96.0)
        );
        let source = session.media_bytes(part).unwrap();
        assert!(ooxml_metafile::decode(&source).is_none());
        let svg = session.metafile_svg(part).unwrap();
        let markup = std::str::from_utf8(&svg).unwrap();
        assert!(markup.contains(if index == 0 { "<text" } else { "<image" }));
        assert!(std::sync::Arc::ptr_eq(
            &svg,
            &session.metafile_svg(part).unwrap()
        ));
        assert_eq!(source, session.media_bytes(part).unwrap());
    }
    assert_eq!(session.save().unwrap(), saved);
    assert_eq!(session.encode_state_as_update_v1(), state);
}
