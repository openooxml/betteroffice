use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/value-axis-autoscale.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn close(actual: f32, expected: f32) {
    assert!((actual - expected).abs() < 0.001, "{actual} != {expected}");
}

#[test]
fn an_unpinned_value_axis_counts_in_round_steps_and_leaves_headroom() {
    let session = DeckSession::open(FIXTURE, 320).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    let list = renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
        .display_list;
    let primitives = list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Chart { primitives, .. } => Some(primitives),
            _ => None,
        })
        .unwrap();

    let mut ticks: Vec<(String, f32)> = primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .filter_map(|line| {
            let text: String = line.runs.iter().map(|run| run.text.as_str()).collect();
            text.parse::<f32>().ok().map(|_| (text, line.baseline))
        })
        .collect();
    ticks.sort_by(|a, b| b.1.total_cmp(&a.1));
    let labels: Vec<&str> = ticks.iter().map(|(text, _)| text.as_str()).collect();
    assert_eq!(labels, ["0", "2", "4", "6", "8", "10", "12", "14"]);
    for (index, (_, baseline)) in ticks.iter().enumerate() {
        close(*baseline, 401.0 - 39.142857 * index as f32);
    }

    let top = primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                y,
                h,
                fill: Some(Paint::Solid { color }),
                ..
            } if color == "#1FA97A" && *h > 8.0 => Some(*y),
            _ => None,
        })
        .fold(f32::MAX, f32::min);
    close(top, 157.27142);
    let (tick_14, tick_12) = (ticks[7].1 - 3.0, ticks[6].1 - 3.0);
    assert!(
        tick_14 < top && top < tick_12,
        "{top} not in {tick_14}..{tick_12}"
    );
}
