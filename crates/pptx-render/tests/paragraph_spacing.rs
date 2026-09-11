use pptx_edit::DeckSession;
use pptx_parse::ShapeNode;
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/paragraph-spacing.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    renderer
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

fn gaps(lines: &[PositionedTextLine]) -> Vec<f32> {
    lines
        .windows(2)
        .map(|pair| pair[1].y - pair[0].y - pair[0].height)
        .collect()
}

fn points(value: f32) -> f32 {
    value * 96.0 / 72.0
}

fn assert_gaps(lines: &[PositionedTextLine], expected: f32) {
    assert_eq!(lines.len(), 3);
    for gap in gaps(lines) {
        assert!(
            (gap - expected).abs() < 0.01,
            "gap {gap}, expected {expected}"
        );
    }
}

#[test]
fn spacing_opens_between_paragraphs_and_cascades_from_the_master_and_layout() {
    let session = DeckSession::open(DECK, 4_240).unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;

    assert_gaps(lines(&list, 8), 0.0);
    assert_gaps(lines(&list, 2), points(10.0));
    assert_gaps(lines(&list, 3), points(28.0));
    assert_gaps(lines(&list, 4), points(16.0));
    assert_gaps(lines(&list, 5), 0.0);
    assert_gaps(lines(&list, 6), points(12.0));

    for id in [2, 3, 4, 5] {
        let first = lines(&list, id)[0].y;
        assert!((first - 40.0).abs() < 0.01, "shape {id} starts at {first}");
    }
    let bottom = lines(&list, 7).last().unwrap();
    let end = bottom.y + bottom.height;
    assert!(
        (end - 640.0).abs() < 0.01,
        "bottom anchored text ends at {end}"
    );
}

#[test]
fn spacing_shrinks_with_the_autofit_scale() {
    let session = DeckSession::open(DECK, 4_241).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut package = session.package().clone();
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[1] else {
        panic!("text shape")
    };
    shape.text.as_mut().unwrap().autofit = Some(pptx_parse::TextAutofit::Normal {
        font_scale: Some(0.5),
        line_space_reduction: None,
    });
    let list = renderer()
        .layout_slide(&package, &snapshot, 0)
        .unwrap()
        .display_list;

    assert_gaps(lines(&list, 3), points(14.0));
}
