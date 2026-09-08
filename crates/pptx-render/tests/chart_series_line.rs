use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/chart-series-line.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn chart_primitives(
    session: &DeckSession,
    renderer: &mut SlideRenderer,
    index: usize,
) -> Vec<Primitive> {
    let snapshot = session.snapshot().unwrap();
    renderer
        .layout_slide(session.package(), &snapshot, index)
        .unwrap()
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Chart { primitives, .. } => Some(primitives.clone()),
            _ => None,
        })
        .unwrap()
}

/// The width of every segment stroked in `color`.
fn segments(primitives: &[Primitive], color: &str) -> Vec<f32> {
    primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                geometry,
                stroke: Some(stroke),
                ..
            } if geometry == "line" && stroke.color == color => Some(stroke.width),
            _ => None,
        })
        .collect()
}

/// The markers and legend key filled in `color`.
fn swatches(primitives: &[Primitive], color: &str) -> usize {
    primitives
        .iter()
        .filter(|primitive| {
            matches!(
                primitive,
                Primitive::Shape {
                    fill: Some(Paint::Solid { color: actual }),
                    ..
                } if actual == color
            )
        })
        .count()
}

fn assert_widths(actual: &[f32], count: usize, expected: f32) {
    assert_eq!(actual.len(), count, "{actual:?}");
    for width in actual {
        assert!((width - expected).abs() < 0.001, "{width} != {expected}");
    }
}

#[test]
fn a_series_is_stroked_at_the_width_its_own_sp_pr_declares() {
    let session = DeckSession::open(FIXTURE, 351).unwrap();
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    assert_eq!(session.snapshot().unwrap().slides.len(), 3);
    let slides: Vec<_> = (0..3)
        .map(|index| chart_primitives(&session, &mut renderer, index))
        .collect();

    let line = &slides[0];
    assert_widths(&segments(line, "#D64550"), 3, 41275.0 / 9525.0);
    assert_widths(&segments(line, "#1B4F9C"), 3, 2.0);
    assert!(segments(line, "#2F8F5B").is_empty());
    assert_eq!(swatches(line, "#2F8F5B"), 5);

    let scatter = &slides[1];
    assert_widths(&segments(scatter, "#7A3EA1"), 3, 3.0);
    assert!(segments(scatter, "#C77B12").is_empty());
    assert_eq!(swatches(scatter, "#C77B12"), 5);

    let radar = &slides[2];
    assert_widths(&segments(radar, "#0E7C86"), 4, 12700.0 / 9525.0);
    assert_widths(&segments(radar, "#B02E6F"), 4, 2.0);
    assert_widths(&segments(radar, "#666666"), 4, 0.5);
}
