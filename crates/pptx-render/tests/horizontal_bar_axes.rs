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
            ("0", 181.41602),
            ("10", 288.26343),
            ("20", 395.11084),
            ("30", 501.95825),
            ("40", 608.80566),
        ] {
            let line = label(primitives, tick);
            close(
                line.x + line.width / 2.0,
                if index == 4 { 790.2217 - x } else { x },
            );
            close(line.baseline, 406.0);
            assert!(line.runs.iter().all(|run| run.color == "#222222"));
        }
        close(label(primitives, "Quarter").baseline, 141.6);
        close(label(primitives, "Millions").baseline, 418.5);
        for (category, baseline) in [
            ("Category 1", 348.18332),
            ("Category 2", 268.55),
            ("Category 3", 188.91667),
        ] {
            let line = label(primitives, category);
            close(line.x + line.width, 168.41602);
            close(
                line.baseline,
                if index == 1 {
                    537.1 - baseline
                } else {
                    baseline
                },
            );
            assert!(line.x >= 96.0);
        }
        for (color, widths) in [
            ("#6254E7", [128.21689, 203.01009, 74.79319]),
            ("#1FA97A", [85.47793, 149.58638, 224.37956]),
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
                close(bar.3, 31.853333);
            }
            assert_eq!(bars[0].1 > bars[2].1, index != 1);
        }
        if index == 3 {
            let top = label(primitives, "80");
            close(top.x + top.width / 2.0, 608.80566);
            close(top.baseline, 131.1);
            assert!(primitives.iter().any(|primitive| matches!(primitive,
                Primitive::Shape { x, y, w, h, stroke: Some(stroke), .. }
                    if stroke.color == "#D9D9D9" && stroke.width == 0.25
                        && *x > 182.0 && (*y - 146.6).abs() < 0.001 && *w == 0.0
                        && (*h - 238.9).abs() < 0.001
            )));
        }
    }
}
