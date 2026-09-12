use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/chart-text-properties.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// `(family, bold, italic, size_px, letter_spacing_px)` of every chart run.
type RunStyle = (String, bool, bool, f32, f32);

fn chart_runs(
    session: &DeckSession,
    renderer: &mut SlideRenderer,
    index: usize,
) -> Vec<(String, RunStyle)> {
    let snapshot = session.snapshot().unwrap();
    let list = renderer
        .layout_slide(session.package(), &snapshot, index)
        .unwrap()
        .display_list;
    let chart = list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Chart { primitives, .. } => Some(primitives.clone()),
            _ => None,
        })
        .unwrap();
    let mut runs: Vec<(String, RunStyle)> = chart
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => lines.first()?.runs.first().map(|run| {
                (
                    run.text.clone(),
                    (
                        run.font_family.clone(),
                        run.bold,
                        run.italic,
                        run.font_size_px,
                        run.letter_spacing_px,
                    ),
                )
            }),
            _ => None,
        })
        .collect();
    runs.sort_by(|left, right| left.0.cmp(&right.0));
    runs
}

fn style(runs: &[(String, RunStyle)], text: &str) -> RunStyle {
    runs.iter()
        .find(|(run, _)| run == text)
        .unwrap_or_else(|| panic!("no {text:?} in {runs:?}"))
        .1
        .clone()
}

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    for family in ["Arial", "Georgia", "Verdana"] {
        for bold in [false, true] {
            for italic in [false, true] {
                renderer.register_font(family, bold, italic, FONT).unwrap();
            }
        }
    }
    renderer
}

#[test]
fn chart_text_takes_its_family_slant_size_and_tracking_from_c_txpr() {
    let session = DeckSession::open(FIXTURE, 1).unwrap();
    let mut renderer = renderer();

    let styled = chart_runs(&session, &mut renderer, 0);
    assert_eq!(
        style(&styled, "Revenue"),
        ("Georgia".to_owned(), true, true, 13.0, 8.0)
    );
    assert_eq!(
        style(&styled, "Q1"),
        ("Verdana".to_owned(), true, false, 10.0, 0.0)
    );
    assert_eq!(
        style(&styled, "North"),
        ("Georgia".to_owned(), false, false, 18.666_666, 0.0)
    );
    assert_eq!(
        style(&styled, "Millions"),
        ("Georgia".to_owned(), false, false, 10.0, 0.0)
    );

    let bare = chart_runs(&session, &mut renderer, 1);
    for (text, _) in &bare {
        assert_eq!(
            style(&bare, text).0,
            "Arial",
            "{text:?} should inherit the theme minor font"
        );
    }
    assert_eq!(
        style(&bare, "Revenue"),
        ("Arial".to_owned(), true, false, 13.0, 0.0)
    );
}

#[test]
fn tracking_spaces_the_glyphs_of_a_chart_title_without_a_trailing_gap() {
    let session = DeckSession::open(FIXTURE, 1).unwrap();
    let renderer = &mut renderer();
    let snapshot = session.snapshot().unwrap();
    let title = |index: usize| {
        renderer
            .layout_slide(session.package(), &snapshot, index)
            .unwrap()
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Chart { primitives, .. } => primitives.iter().find_map(|part| {
                    match part {
                        Primitive::TextBox { lines, .. } => lines.first(),
                        _ => None,
                    }
                    .filter(|line| line.runs.first().is_some_and(|run| run.text == "Revenue"))
                    .cloned()
                }),
                _ => None,
            })
            .unwrap()
    };
    let tracked = title(0);
    let plain = title(1);
    let run = &tracked.runs[0];
    let gaps: Vec<f32> = run
        .glyphs
        .windows(2)
        .map(|pair| pair[1].x - pair[0].x - pair[0].advance)
        .collect();
    assert!(gaps.iter().all(|gap| (gap - 8.0).abs() < 0.001), "{gaps:?}");
    assert!(
        (run.width - plain.runs[0].width - 8.0 * (run.glyphs.len() - 1) as f32).abs() < 0.5,
        "{} vs {}",
        run.width,
        plain.runs[0].width
    );
}
