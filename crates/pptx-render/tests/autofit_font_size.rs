use pptx_edit::DeckSession;
use pptx_parse::{ShapeNode, TextAutofit};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer};

const DECK: &[u8] = include_bytes!("fixtures/line-spacing.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render(family: &str, size_pt: f64, scale: f64) -> (f32, Vec<PositionedTextLine>) {
    let mut package = pptx_parse::parse_pptx(DECK).unwrap();
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[2] else {
        panic!("text shape");
    };
    let body = shape.text.as_mut().unwrap();
    for paragraph in &mut body.paragraphs {
        for run in &mut paragraph.runs {
            run.properties.font_family = Some(family.to_owned());
            run.properties.font_size_pt = Some(size_pt);
        }
    }
    body.autofit = Some(TextAutofit::Normal {
        font_scale: Some(scale),
        line_space_reduction: None,
    });
    let session = DeckSession::from_package(package, 8_021).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: 4,
                paragraphs,
                lines,
                ..
            } => Some((paragraphs[0].runs[0].font_size_pt, lines)),
            _ => None,
        })
        .unwrap()
}

#[test]
fn fractional_autofit_matches_an_authored_whole_point_size() {
    for family in ["Arial", "Trebuchet MS"] {
        let (size, lines) = render(family, 32.0, 0.85);
        let (_, authored) = render(family, 27.0, 1.0);
        assert_eq!(size, 27.0);
        assert_eq!(lines, authored);
        assert_eq!(lines[0].runs[0].font_size_px, 36.0);
    }
}

#[test]
fn unscaled_text_keeps_its_authored_fractional_size() {
    let (size, lines) = render("Arial", 27.5, 1.0);
    assert_eq!(size, 27.5);
    assert!((lines[0].runs[0].font_size_px - 27.5 * 4.0 / 3.0).abs() < 0.001);
}
