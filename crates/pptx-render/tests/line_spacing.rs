use pptx_edit::DeckSession;
use pptx_parse::ShapeNode;
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/line-spacing.pptx");
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

#[test]
fn exact_spacing_preserves_the_first_baseline_and_scales_with_autofit() {
    let session = DeckSession::open(DECK, 3141).unwrap();
    let renderer = renderer();
    let snapshot = session.snapshot().unwrap();
    let mut package = session.package().clone();
    let exact = renderer
        .layout_slide(&package, &snapshot, 0)
        .unwrap()
        .display_list;
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[2] else {
        panic!("text shape")
    };
    shape.text.as_mut().unwrap().paragraphs[0]
        .properties
        .line_spacing = None;
    let natural = renderer
        .layout_slide(&package, &snapshot, 0)
        .unwrap()
        .display_list;
    let exact = lines(&exact, 4);
    let natural = lines(&natural, 4);
    assert_eq!(exact.len(), 2);
    assert!(
        (exact[0].baseline - natural[0].baseline).abs() < 0.001,
        "exact baseline {}, natural {}",
        exact[0].baseline,
        natural[0].baseline
    );
    assert!((exact[1].baseline - exact[0].baseline - 96.0).abs() < 0.001);
    assert_eq!(
        exact[0].runs[0].font_size_px,
        natural[0].runs[0].font_size_px
    );

    let mut package = session.package().clone();
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[2] else {
        panic!("text shape")
    };
    shape.text.as_mut().unwrap().autofit = Some(pptx_parse::TextAutofit::Normal {
        font_scale: Some(0.5),
        line_space_reduction: None,
    });
    let scaled = renderer
        .layout_slide(&package, &snapshot, 0)
        .unwrap()
        .display_list;
    let scaled = lines(&scaled, 4);
    assert!((scaled[1].baseline - scaled[0].baseline - 48.0).abs() < 0.001);
    assert!(
        (scaled[0].baseline - scaled[0].y - (natural[0].baseline - natural[0].y) * 0.5).abs()
            < 0.001
    );
}

#[test]
fn fixture_spacing_cascades_from_master_layout_and_shape_lists() {
    let session = DeckSession::open(DECK, 3142).unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    for (id, pitch) in [(2, 40.96), (5, 61.44), (6, 76.8), (8, 17.28)] {
        let lines = lines(&list, id);
        assert_eq!(lines.len(), 2, "shape {id}");
        assert!(
            (lines[1].y - lines[0].y - pitch).abs() < 0.001,
            "shape {id}: {}",
            lines[1].y - lines[0].y
        );
        assert_eq!(lines[0].runs[0].color, "#2040B0");
    }
    for id in [5, 6] {
        assert!((lines(&list, id)[0].baseline - 338.625).abs() < 0.001);
    }
    let font_based = lines(&list, 3);
    let control = lines(&list, 7);
    let natural_32 = control[0].height * 32.0 / 24.0;
    assert!((font_based[0].height - natural_32 * 0.8).abs() < 0.001);
}

#[test]
fn zero_spacing_overrides_inherited_spacing() {
    let session = DeckSession::open(DECK, 3143).unwrap();
    let renderer = renderer();
    for spacing in [
        pptx_parse::LineSpacing::Percent { value: 0.0 },
        pptx_parse::LineSpacing::Points { value: 0.0 },
    ] {
        let mut package = session.package().clone();
        let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[0] else {
            panic!("text shape")
        };
        shape.text.as_mut().unwrap().paragraphs[0]
            .properties
            .line_spacing = Some(spacing);
        let list = renderer
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list;
        let lines = lines(&list, 2);
        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].height, 0.0);
        assert_eq!(lines[1].baseline, lines[0].baseline);
    }
}
