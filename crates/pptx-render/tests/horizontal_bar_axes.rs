use pptx_edit::DeckSession;
use pptx_render::{Paint, PositionedTextLine, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/horizontal-bar-axes.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn label<'a>(primitives: &'a [Primitive], text: &str) -> &'a PositionedTextLine {
    primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .find(|line| {
            line.runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>()
                == text
        })
        .unwrap()
}

fn close(actual: f32, expected: f32) {
    assert!((actual - expected).abs() < 0.001, "{actual} != {expected}");
}

#[test]
fn horizontal_bar_fixture_transposes_axes_without_losing_series_or_labels() {
    let session = DeckSession::open(FIXTURE, 310).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    assert_eq!(snapshot.slides.len(), 5);
    for index in [0, 1, 3, 4] {
        let list = renderer
            .layout_slide(session.package(), &snapshot, index)
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
        for (tick, x) in [
            ("0", 156.0),
            ("10", 252.5),
            ("20", 349.0),
            ("30", 445.5),
            ("40", 542.0),
        ] {
            let line = label(primitives, tick);
            close(line.x, if index == 4 { 698.0 - x } else { x });
            close(line.baseline, 412.0);
            assert!(line.runs.iter().all(|run| run.color == "#222222"));
        }
        close(label(primitives, "Quarter").baseline, 137.0);
        close(label(primitives, "Millions").baseline, 424.0);
        for (category, baseline) in [
            ("Category 1", 359.6),
            ("Category 2", 274.26666),
            ("Category 3", 188.93333),
        ] {
            let line = label(primitives, category);
            close(line.x, 100.0);
            close(
                line.baseline,
                if index == 1 {
                    548.5333 - baseline
                } else {
                    baseline
                },
            );
            assert!(line.x + line.width < 172.0);
        }
        for (color, widths) in [
            ("#6254E7", [115.8, 183.35, 67.55]),
            ("#1FA97A", [77.2, 135.1, 202.65]),
        ] {
            let bars: Vec<_> = primitives
                .iter()
                .filter_map(|primitive| match primitive {
                    Primitive::Shape {
                        x,
                        y,
                        w,
                        h,
                        fill: Some(Paint::Solid { color: actual }),
                        ..
                    } if actual == color && *h > 8.0 => Some((*x, *y, *w, *h)),
                    _ => None,
                })
                .collect();
            assert_eq!(bars.len(), 3);
            for (bar, width) in bars.iter().zip(widths) {
                close(bar.2, width);
                close(bar.3, 34.133335);
            }
            assert_eq!(bars[0].1 > bars[2].1, index != 1);
        }
        if index == 3 {
            close(label(primitives, "80").x, 542.0);
            close(label(primitives, "80").baseline, 136.0);
            assert!(primitives.iter().any(|primitive| matches!(primitive,
                Primitive::Shape { x, y, w, h, stroke: Some(stroke), .. }
                    if stroke.color == "#D9D9D9" && stroke.width == 0.25
                        && *x > 172.0 && *y == 142.0 && *w == 0.0 && *h == 256.0
            )));
        }
    }
}
