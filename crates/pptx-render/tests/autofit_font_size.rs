use pptx_edit::DeckSession;
use pptx_parse::{ShapeNode, TextAutofit};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer};

const DECK: &[u8] = include_bytes!("fixtures/line-spacing.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render(size_pt: f64, scale: f64) -> (f32, Vec<PositionedTextLine>) {
    let session = DeckSession::open(DECK, 8_021).unwrap();
    let mut snapshot = session.snapshot().unwrap();
    let shape = snapshot.slides[0]
        .shapes
        .iter_mut()
        .find(|shape| shape.source_id == 4)
        .unwrap();
    for paragraph in &mut shape.text_stories[0].paragraphs {
        for run in &mut paragraph.runs {
            run.style.font_size_pt = Some(size_pt);
        }
    }
    let mut package = session.package().clone();
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[2] else {
        panic!("text shape");
    };
    shape.text.as_mut().unwrap().autofit = Some(TextAutofit::Normal {
        font_scale: Some(scale),
        line_space_reduction: None,
    });
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(&package, &snapshot, 0)
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
    let (size, lines) = render(32.0, 0.85);
    let (_, authored) = render(27.0, 1.0);
    assert_eq!(size, 27.0);
    assert_eq!(lines, authored);
    assert_eq!(lines[0].runs[0].font_size_px, 36.0);
}

#[test]
fn unscaled_text_keeps_its_authored_fractional_size() {
    let (size, lines) = render(27.5, 1.0);
    assert_eq!(size, 27.5);
    assert!((lines[0].runs[0].font_size_px - 27.5 * 4.0 / 3.0).abs() < 0.001);
}
