use pptx_edit::DeckSession;
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/theme-fonts.pptx");
const SANS: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf");
const SERIF: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSerif-Regular.ttf");

fn render(slide: usize) -> SurfaceDisplayList {
    let session = DeckSession::open(DECK, 291).unwrap();
    let mut renderer = SlideRenderer::new();
    for (family, bytes) in [
        ("Arial", SANS),
        ("Liberation Sans", SANS),
        ("Liberation Serif", SERIF),
    ] {
        renderer.register_font(family, false, false, bytes).unwrap();
    }
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), slide)
        .unwrap()
        .display_list
}

fn lines(list: &SurfaceDisplayList, id: u32) -> &[PositionedTextLine] {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id, lines, ..
            } if *object_id == id => Some(lines.as_slice()),
            _ => None,
        })
        .unwrap()
}

#[test]
fn major_theme_runs_use_the_major_face_and_metrics() {
    let list = render(0);
    let title = &lines(&list, 2)[0];
    assert_eq!(title.runs[0].font_family, "Liberation Serif");
    assert_eq!(title.runs[0].font_id, 2);
    assert_eq!(title.runs[0].color, "#17365D");
    assert!((title.width - 437.0625).abs() < 0.001);
    let body = lines(&list, 3);
    assert_eq!(body.len(), 3);
    for (line, expected) in body.iter().zip([
        "Theme major font chooses the ",
        "heading face and its own ",
        "wrapping metrics.",
    ]) {
        assert_eq!(line.runs.len(), 1);
        assert_eq!(line.runs[0].text, expected);
        assert_eq!(line.runs[0].font_id, 2);
        assert_eq!(line.runs[0].color, "#17365D");
    }
    let minor = &lines(&list, 4)[0].runs[0];
    assert_eq!(minor.font_family, "Liberation Sans");
    assert_eq!(minor.font_id, 1);
    assert_eq!(minor.color, "#008080");
    let explicit = &lines(&list, 5)[0].runs[0];
    assert_eq!(explicit.font_family, "Arial");
    assert_eq!(explicit.font_id, 0);
    assert_eq!(explicit.color, "#7F3F00");
}

#[test]
fn literal_font_names_keep_the_configured_fallback() {
    let list = render(1);
    for (id, width) in [(2, 690.15625), (3, 754.1719), (4, 725.6875)] {
        let line = &lines(&list, id)[0];
        assert_eq!(line.runs[0].font_family, "Arial");
        assert_eq!(line.runs[0].font_id, 0);
        assert!(!line.runs[0].bold);
        assert!((line.width - width).abs() < 0.001);
    }
    let registered = &lines(&list, 5)[0].runs[0];
    assert_eq!(registered.font_family, "Liberation Serif");
    assert_eq!(registered.font_id, 2);
}
