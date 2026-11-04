#[path = "../../pptx-parse/tests/common/strikethrough.rs"]
mod fixture;

use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn positioned_runs_inherit_strike_and_respect_no_strike() {
    let session = DeckSession::open(&fixture::deck(), 6407).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let layout = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let runs = layout
        .display_list
        .primitives
        .iter()
        .flat_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => {
                lines.iter().flat_map(|line| &line.runs).collect::<Vec<_>>()
            }
            _ => Vec::new(),
        })
        .collect::<Vec<_>>();
    for (text, strike) in [
        ("A", Some("sngStrike")),
        ("B", Some("dblStrike")),
        ("C", Some("noStrike")),
        ("D", Some("dblStrike")),
        ("E", Some("noStrike")),
        ("M", Some("sngStrike")),
        ("N", Some("noStrike")),
        ("L", Some("dblStrike")),
        ("O", Some("noStrike")),
        ("P", None),
        ("Q", Some("sngStrike")),
    ] {
        let run = runs.iter().find(|run| run.text.contains(text)).unwrap();
        assert_eq!(run.strike.as_deref(), strike, "{text}");
    }
}
