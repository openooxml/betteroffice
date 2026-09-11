use pptx_edit::{DeckSession, EditCtx};
use pptx_parse::{ShapeNode, TextAutofit};
use pptx_render::{Primitive, RenderedSlide, SlideRenderer};

const DECK: &[u8] = include_bytes!("fixtures/vertical-writing-modes.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render() -> RenderedSlide {
    let session = DeckSession::open(DECK, 308).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
}

fn text(slide: &RenderedSlide, id: u32) -> &Primitive {
    slide
        .display_list
        .primitives
        .iter()
        .find(|primitive| {
            matches!(primitive, Primitive::TextBox { object_id, .. } if *object_id == id)
        })
        .unwrap()
}

#[test]
fn east_asian_vertical_turns_the_box_like_vert() {
    let slide = render();
    let Primitive::TextBox {
        transform: vert_transform,
        w: vert_w,
        h: vert_h,
        lines: vert_lines,
        ..
    } = text(&slide, 44)
    else {
        unreachable!()
    };
    let Primitive::TextBox {
        transform,
        w,
        h,
        lines,
        ..
    } = text(&slide, 40)
    else {
        unreachable!()
    };
    assert_eq!(transform.rotation_deg, 90.0);
    assert_eq!((transform, w, h), (vert_transform, vert_w, vert_h));
    assert_eq!(lines.len(), 2);
    assert_eq!(
        lines.iter().map(|line| line.y).collect::<Vec<_>>(),
        vert_lines.iter().map(|line| line.y).collect::<Vec<_>>()
    );
    assert!(lines[0].y < lines[1].y);
}

#[test]
fn mongolian_vertical_turns_the_box_and_runs_its_lines_the_other_way() {
    let slide = render();
    let Primitive::TextBox {
        transform, lines, ..
    } = text(&slide, 41)
    else {
        unreachable!()
    };
    let Primitive::TextBox {
        lines: east_asian, ..
    } = text(&slide, 40)
    else {
        unreachable!()
    };
    assert_eq!(transform.rotation_deg, 90.0);
    assert_eq!(lines.len(), 2);
    assert!(lines[0].y > lines[1].y);
    let span = |lines: &[pptx_render::PositionedTextLine]| {
        let top = lines.iter().map(|line| line.y).fold(f32::MAX, f32::min);
        let bottom = lines
            .iter()
            .map(|line| line.y + line.height)
            .fold(f32::MIN, f32::max);
        (top, bottom)
    };
    assert_eq!(span(lines), span(east_asian));
}

#[test]
fn word_art_vertical_stacks_one_cluster_a_line_without_turning_the_box() {
    let slide = render();
    for id in [42, 43] {
        let Primitive::TextBox {
            transform,
            w,
            h,
            lines,
            ..
        } = text(&slide, id)
        else {
            unreachable!()
        };
        assert_eq!(transform.rotation_deg, 0.0, "shape {id}");
        assert!(*w < *h, "shape {id}");
        assert_eq!(lines.len(), 5, "shape {id}");
        let stacked = lines
            .iter()
            .map(|line| line.runs[0].text.as_str())
            .collect::<String>();
        assert_eq!(stacked, "STACK", "shape {id}");
        for line in lines {
            assert_eq!(line.runs.len(), 1, "shape {id}");
            assert_eq!(line.runs[0].glyphs.len(), 1, "shape {id}");
        }
        for pair in lines.windows(2) {
            assert!(pair[0].y + pair[0].height <= pair[1].y, "shape {id}");
        }
    }
}

/// Point size of `source_id` after squeezing its box, with `normAutofit` on or off.
fn squeezed_font_size(source_id: u32, autofit: bool) -> f32 {
    let mut package = pptx_parse::parse_pptx(DECK).unwrap();
    let session = DeckSession::open(DECK, 309).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[0].id.clone();
    let shape_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.source_id == source_id)
        .unwrap()
        .id
        .clone();
    session
        .resize_shape(
            &EditCtx::local("test"),
            &slide_id,
            &shape_id,
            1_600_200,
            200_000,
        )
        .unwrap();
    if autofit {
        let ShapeNode::Shape(parsed) = package.slides[0]
            .shapes
            .iter_mut()
            .find(|shape| shape.id() == source_id)
            .unwrap()
        else {
            unreachable!()
        };
        parsed.text.as_mut().unwrap().autofit = Some(TextAutofit::Normal {
            font_scale: None,
            line_space_reduction: None,
        });
    }
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let slide = renderer
        .layout_slide(&package, &session.snapshot().unwrap(), 0)
        .unwrap();
    let Primitive::TextBox { paragraphs, .. } = text(&slide, source_id) else {
        unreachable!()
    };
    paragraphs[0].runs[0].font_size_pt
}

#[test]
fn autofit_shrinks_a_squeezed_horizontal_box_but_never_a_stack() {
    assert_eq!(squeezed_font_size(42, true), squeezed_font_size(42, false));
    assert!(squeezed_font_size(45, true) < squeezed_font_size(45, false));
}

#[test]
fn a_horizontal_box_on_the_same_slide_is_untouched() {
    let slide = render();
    let Primitive::TextBox {
        transform,
        w,
        h,
        lines,
        ..
    } = text(&slide, 45)
    else {
        unreachable!()
    };
    assert_eq!(transform.rotation_deg, 0.0);
    assert!(*w < *h);
    assert_eq!(lines.len(), 2);
    assert!(lines[0].y < lines[1].y);
}
