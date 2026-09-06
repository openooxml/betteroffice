use pptx_edit::DeckSession;
use pptx_render::{HitTestResult, Primitive, RenderedSlide, SlideRenderer};

const DECK: &[u8] = include_bytes!("fixtures/text-overflow.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn slide(index: usize) -> RenderedSlide {
    let session = DeckSession::open(DECK, 295).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), index)
        .unwrap()
}

fn text(slide: &RenderedSlide, id: u32) -> &Primitive {
    slide
        .display_list
        .primitives
        .iter()
        .find(|primitive| matches!(primitive, Primitive::TextBox { object_id, .. } if *object_id == id))
        .unwrap()
}

#[test]
fn shape_autofit_preserves_explicit_and_inherited_font_sizes() {
    for (index, id) in [(0, 3), (1, 4)] {
        let rendered = slide(index);
        let Primitive::TextBox {
            lines,
            paragraphs,
            overflow,
            ..
        } = text(&rendered, id)
        else {
            unreachable!()
        };
        assert_eq!(lines[0].runs[0].font_size_px, 64.0);
        assert_eq!(paragraphs[0].runs[0].font_size_pt, 48.0);
        assert!(*overflow);
    }
}

#[test]
fn overflowing_lines_respect_top_center_and_bottom_anchors() {
    let rendered = slide(0);
    for (id, fraction) in [(2, 0.0), (4, 0.5), (5, 1.0)] {
        let Primitive::TextBox {
            y,
            h,
            lines,
            overflow,
            ..
        } = text(&rendered, id)
        else {
            unreachable!()
        };
        assert!(*overflow);
        assert!(lines[0].height > *h);
        assert!((lines[0].y - (y + (h - lines[0].height) * fraction)).abs() < 0.001);
    }
}

#[test]
fn overflow_hit_testing_uses_the_original_rotation_and_flip_pivot() {
    for (index, id) in [(0, 2), (0, 4), (0, 5), (1, 2), (1, 3)] {
        let rendered = slide(index);
        let Primitive::TextBox {
            shape_id,
            story_id,
            x,
            y,
            w,
            h,
            lines,
            transform,
            ..
        } = text(&rendered, id)
        else {
            unreachable!()
        };
        let line = &lines[0];
        let outside_y = if line.y < *y {
            line.y + 1.0
        } else {
            line.y + line.height - 1.0
        };
        assert!(outside_y < *y || outside_y > y + h);
        for stop in &line.caret_stops {
            let cx = x + w / 2.0;
            let cy = y + h / 2.0;
            let dx = (stop.x - cx) * if transform.flip_h { -1.0 } else { 1.0 };
            let dy = (outside_y - cy) * if transform.flip_v { -1.0 } else { 1.0 };
            let (sin, cos) = transform.rotation_deg.to_radians().sin_cos();
            let hit = rendered.hit_test(cx + dx * cos - dy * sin, cy + dx * sin + dy * cos);
            assert_eq!(
                hit,
                Some(HitTestResult::Text {
                    shape_id: shape_id.clone().unwrap(),
                    story_id: story_id.clone().unwrap(),
                    position: stop.position,
                }),
                "slide {index}, shape {id}, position {}",
                stop.position
            );
        }
    }
}

#[test]
fn explicit_and_inherited_clipping_keep_hidden_text_out_of_hit_testing() {
    let rendered = slide(2);
    for id in [2, 3, 4, 5] {
        let Primitive::TextBox {
            x,
            y,
            h,
            lines,
            overflow,
            ..
        } = text(&rendered, id)
        else {
            unreachable!()
        };
        assert!(!*overflow, "shape {id}");
        assert!(lines[0].height > *h);
        assert_eq!(rendered.hit_test(x + 10.0, y + h + 5.0), None);
    }
    let Primitive::TextBox { overflow, .. } = text(&rendered, 6) else {
        unreachable!()
    };
    assert!(*overflow);
}

#[test]
fn normal_autofit_still_shrinks_and_fitting_text_keeps_its_size() {
    let rendered = slide(0);
    let Primitive::TextBox {
        lines, overflow, ..
    } = text(&rendered, 6)
    else {
        unreachable!()
    };
    assert!(!*overflow);
    assert!(lines[0].runs[0].font_size_px < 64.0);
    assert!(lines[0].height <= 48.0);
    let Primitive::TextBox {
        lines, overflow, ..
    } = text(&rendered, 7)
    else {
        unreachable!()
    };
    assert!(!*overflow);
    assert_eq!(lines[0].runs[0].font_size_px, 64.0);
}

#[test]
fn a_line_wider_than_its_shape_remains_clickable() {
    let session = DeckSession::open(DECK, 295).unwrap();
    let mut snapshot = session.snapshot().unwrap();
    snapshot.slides[0].shapes[0].width = 95_250;
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap();
    let Primitive::TextBox {
        shape_id,
        story_id,
        x,
        w,
        lines,
        ..
    } = text(&rendered, 2)
    else {
        unreachable!()
    };
    let stop = lines[0].caret_stops.last().unwrap();
    assert!(stop.x > x + w);
    assert_eq!(
        rendered.hit_test(stop.x, lines[0].baseline),
        Some(HitTestResult::Text {
            shape_id: shape_id.clone().unwrap(),
            story_id: story_id.clone().unwrap(),
            position: stop.position,
        })
    );
}
