use pptx_edit::DeckSession;
use pptx_render::{HitTestResult, Primitive, RenderedSlide, SlideRenderer, Transform};

const DECK: &[u8] = include_bytes!("fixtures/text-orientation.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn render(index: usize) -> RenderedSlide {
    let session = DeckSession::open(DECK, 308).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), index)
        .unwrap()
}

fn text(slide: &RenderedSlide, id: u32) -> &Primitive {
    slide.display_list.primitives.iter().find(|primitive| {
        matches!(primitive, Primitive::TextBox { object_id, .. } if *object_id == id)
    }).unwrap()
}

#[test]
fn parsed_flips_and_vertical_bars_keep_text_and_carets_readable() {
    let slide = render(0);
    for (id, rotation) in [
        (2, 0.0),
        (3, 0.0),
        (4, 0.0),
        (5, 180.0),
        (6, 180.0),
        (7, 0.0),
    ] {
        let Primitive::TextBox {
            transform,
            lines,
            shape_id,
            story_id,
            ..
        } = text(&slide, id)
        else {
            unreachable!()
        };
        assert_eq!(
            *transform,
            Transform {
                rotation_deg: rotation,
                ..Transform::default()
            },
            "shape {id}"
        );
        assert_eq!(lines.len(), 1, "shape {id}");
        assert_eq!(lines[0].runs[0].color, "#17365D");
        if rotation == 0.0 {
            for caret in &lines[0].caret_stops {
                assert_eq!(
                    slide.hit_test(caret.x, lines[0].baseline),
                    Some(HitTestResult::Text {
                        shape_id: shape_id.clone().unwrap(),
                        story_id: story_id.clone().unwrap(),
                        position: caret.position,
                    }),
                    "shape {id}, caret {}",
                    caret.position
                );
            }
        }
    }
    let Primitive::TextBox { w, h, .. } = text(&slide, 2) else {
        unreachable!()
    };
    assert!((*w - 810.29205).abs() < 0.001);
    assert!((*h - 67.19055).abs() < 0.001);
    for (id, rotation, flip_h, flip_v) in [
        (3, 180.0, false, true),
        (4, 0.0, true, false),
        (5, 0.0, false, true),
        (6, 0.0, true, true),
    ] {
        assert!(slide.display_list.primitives.iter().any(|primitive| matches!(primitive,
            Primitive::Shape { object_id, transform, .. } if *object_id == id && *transform == Transform { rotation_deg: rotation, flip_h, flip_v }
        )));
    }
}

#[test]
fn vertical_text_keeps_insets_on_the_shape_edges() {
    let slide = render(1);
    for (id, x, y, rotation) in [(10, -65.0, 295.0, 90.0), (11, 135.0, 275.0, 270.0)] {
        let Primitive::TextBox {
            lines, transform, ..
        } = text(&slide, id)
        else {
            unreachable!()
        };
        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].x, x, "shape {id}");
        assert_eq!(lines[0].y, y, "shape {id}");
        assert_eq!(transform.rotation_deg, rotation);
    }
}

#[test]
fn vertical_direction_cascades_and_explicit_horizontal_wins() {
    let slide = render(1);
    for (id, rotation, dimensions) in [
        (12, 270.0, (450.0, 120.0)),
        (13, 90.0, (450.0, 120.0)),
        (14, 0.0, (320.0, 80.0)),
    ] {
        let Primitive::TextBox {
            w,
            h,
            transform,
            lines,
            ..
        } = text(&slide, id)
        else {
            unreachable!()
        };
        assert_eq!((*w, *h), dimensions, "shape {id}");
        assert_eq!(transform.rotation_deg, rotation, "shape {id}");
        assert_eq!(lines.len(), 1, "shape {id}");
    }
}

#[test]
fn horizontal_text_preserves_the_original_rotation() {
    let slide = render(2);
    for (id, rotation) in [(20, -15.0), (21, 390.0), (22, 0.0), (23, 0.0)] {
        let Primitive::TextBox { transform, .. } = text(&slide, id) else {
            unreachable!()
        };
        assert_eq!(transform.rotation_deg, rotation, "shape {id}");
    }
}

#[test]
fn vertical_autofit_and_anchoring_use_the_inset_text_area() {
    let session = DeckSession::open(DECK, 308).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let snapshot = session.snapshot().unwrap();
    for autofit in [
        pptx_parse::TextAutofit::Normal {
            font_scale: Some(0.85),
            line_space_reduction: None,
        },
        pptx_parse::TextAutofit::Shape,
    ] {
        let shape_autofit = matches!(autofit, pptx_parse::TextAutofit::Shape);
        let mut package = session.package().clone();
        let shape = package.slides[1]
            .shapes
            .iter_mut()
            .find_map(|node| match node {
                pptx_parse::ShapeNode::Shape(shape) if shape.base.id == 15 => Some(shape),
                _ => None,
            })
            .unwrap();
        shape.text.as_mut().unwrap().autofit = Some(autofit);
        let slide = renderer.layout_slide(&package, &snapshot, 1).unwrap();
        let Primitive::TextBox {
            lines,
            overflow,
            transform,
            ..
        } = text(&slide, 15)
        else {
            unreachable!()
        };
        assert_eq!(transform.rotation_deg, 90.0);
        assert_eq!(lines.len(), 2);
        assert_eq!(*overflow, shape_autofit);
        let top = lines[0].y;
        let last = lines.last().unwrap();
        let bottom = last.y + last.height;
        if shape_autofit {
            assert!(top < 427.1);
            assert!(bottom > 472.9);
        } else {
            assert!(top >= 427.1);
            assert!(bottom <= 472.9);
        }
        assert!(((top + bottom) / 2.0 - 450.0).abs() < 0.001);
        for line in lines {
            assert_eq!(line.x, 827.3);
            if shape_autofit {
                assert_eq!(line.runs[0].font_size_px, 32.0);
            } else {
                assert!(line.runs[0].font_size_px < 20.0);
            }
            assert!(line.width <= 330.4);
        }
    }
}
