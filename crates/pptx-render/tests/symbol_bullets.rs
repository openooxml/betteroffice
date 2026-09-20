use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/symbol-bullets.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
}

fn marker(list: &SurfaceDisplayList, id: u32) -> String {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id, lines, ..
            } if *object_id == id => Some(lines[0].runs[0].text.clone()),
            _ => None,
        })
        .unwrap()
}

#[test]
fn a_symbol_bullet_font_draws_the_shape_its_slot_addresses() {
    let session = DeckSession::open(DECK, 5_318).unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(marker(&list, 10), "▪");
    assert_eq!(marker(&list, 11), "►");
    assert_eq!(marker(&list, 12), "●");
    assert_eq!(marker(&list, 13), "•");
    assert_eq!(marker(&list, 14), "»");
}

#[test]
fn a_substituted_bullet_is_drawn_with_a_glyph_the_face_covers() {
    let session = DeckSession::open(DECK, 5_318).unwrap();
    let renderer = renderer();
    let list = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    for id in [10, 11, 12, 13, 14] {
        let glyphs = match list.primitives.iter().find(|primitive| {
            matches!(primitive, Primitive::TextBox { object_id, .. } if *object_id == id)
        }) {
            Some(Primitive::TextBox { lines, .. }) => lines[0].runs[0].glyphs.clone(),
            _ => panic!("missing text box {id}"),
        };
        assert!(glyphs.iter().all(|glyph| glyph.glyph_id != 0), "{id}");
    }
}
