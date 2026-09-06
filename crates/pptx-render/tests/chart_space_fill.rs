use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/chart-space-fill.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const FRAME: (f32, f32, f32, f32) = (96.0, 96.0, 576.0, 336.0);

fn close(actual: f32, expected: f32) -> bool {
    (actual - expected).abs() < 0.001
}

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

/// The fill of the rectangle covering the whole chart frame, if one is drawn.
fn ground(primitives: &[Primitive]) -> Option<String> {
    primitives.iter().find_map(|primitive| match primitive {
        Primitive::Shape {
            x, y, w, h, fill, ..
        } if close(*x, FRAME.0)
            && close(*y, FRAME.1)
            && close(*w, FRAME.2)
            && close(*h, FRAME.3) =>
        {
            match fill {
                Some(Paint::Solid { color }) => Some(color.clone()),
                _ => Some(String::new()),
            }
        }
        _ => None,
    })
}

/// `(x, y, w, h, color, width)` of every stroked horizontal or vertical rule.
fn rules(primitives: &[Primitive]) -> Vec<(f32, f32, f32, f32, String, f32)> {
    primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                x,
                y,
                w,
                h,
                stroke: Some(stroke),
                ..
            } if *w == 0.0 || *h == 0.0 => {
                Some((*x, *y, *w, *h, stroke.color.clone(), stroke.width))
            }
            _ => None,
        })
        .collect()
}

fn bars(primitives: &[Primitive], color: &str) -> Vec<(f32, f32, f32, f32)> {
    primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                x,
                y,
                w,
                h,
                fill: Some(Paint::Solid { color: actual }),
                ..
            } if actual == color && *h > 8.0 && *w < 100.0 => Some((*x, *y, *w, *h)),
            _ => None,
        })
        .collect()
}

fn texts(primitives: &[Primitive]) -> Vec<String> {
    let mut texts: Vec<String> = primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { paragraphs, .. } => Some(
                paragraphs
                    .iter()
                    .flat_map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()))
                    .collect(),
            ),
            _ => None,
        })
        .collect();
    texts.sort();
    texts
}

fn assert_rule(
    rule: &(f32, f32, f32, f32, String, f32),
    expected: (f32, f32, f32, f32, &str, f32),
) {
    assert!(
        close(rule.0, expected.0)
            && close(rule.1, expected.1)
            && close(rule.2, expected.2)
            && close(rule.3, expected.3)
            && rule.4 == expected.4
            && close(rule.5, expected.5),
        "{rule:?} != {expected:?}"
    );
}

#[test]
fn a_chart_space_fill_and_axis_lines_reach_the_display_list() {
    let session = DeckSession::open(FIXTURE, 327).unwrap();
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    assert_eq!(session.snapshot().unwrap().slides.len(), 3);
    let slides: Vec<_> = (0..3)
        .map(|index| chart_primitives(&session, &mut renderer, index))
        .collect();

    let pattern = &slides[0];
    assert_eq!(ground(pattern).as_deref(), Some("#01BFBD"));
    let pattern_rules = rules(pattern);
    assert_eq!(pattern_rules.len(), 1, "{pattern_rules:?}");
    assert_rule(
        &pattern_rules[0],
        (138.0, 398.0, 420.0, 0.0, "#D9D9D9", 1.0),
    );
    assert!(!pattern.iter().any(|primitive| matches!(
        primitive,
        Primitive::Shape { stroke: Some(stroke), .. } if stroke.color == "#666666"
    )));

    let bare = &slides[1];
    assert_eq!(ground(bare), None);
    let bare_rules = rules(bare);
    assert_eq!(bare_rules.len(), 2, "{bare_rules:?}");
    assert_rule(&bare_rules[0], (138.0, 124.0, 0.0, 274.0, "#666666", 1.0));
    assert_rule(&bare_rules[1], (138.0, 398.0, 420.0, 0.0, "#666666", 1.0));

    let solid = &slides[2];
    assert_eq!(ground(solid).as_deref(), Some("#F5F7FA"));
    let solid_rules = rules(solid);
    assert_eq!(solid_rules.len(), 2, "{solid_rules:?}");
    assert_rule(&solid_rules[0], (138.0, 124.0, 0.0, 274.0, "#6254E7", 1.0));
    assert_rule(&solid_rules[1], (138.0, 398.0, 420.0, 0.0, "#BFBFBF", 2.0));

    let expected_texts = texts(pattern);
    assert_eq!(expected_texts.len(), 14, "{expected_texts:?}");
    for primitives in &slides {
        assert_eq!(texts(primitives), expected_texts);
        for color in ["#FFFFFF", "#1FA97A"] {
            assert_eq!(bars(primitives, color), bars(pattern, color));
            assert_eq!(bars(primitives, color).len(), 3, "three bars for {color}");
        }
    }
}
