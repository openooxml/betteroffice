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
        ("C", None),
        ("D", Some("dblStrike")),
        ("E", None),
        ("M", Some("sngStrike")),
        ("N", None),
        ("L", Some("dblStrike")),
        ("O", None),
        ("P", None),
        ("Q", Some("sngStrike")),
    ] {
        let run = runs.iter().find(|run| run.text.contains(text)).unwrap();
        assert_eq!(run.strike.as_deref(), strike, "{text}");
    }
}

#[test]
fn adjacent_plain_and_no_strike_runs_coalesce() {
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    r#"<a:rPr strike="sngStrike"/><a:t>Q</a:t>"#,
                    r#"<a:rPr strike="noStrike"/><a:t>Q</a:t>"#,
                )
                .into_bytes();
        }
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 6441).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let layout = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let runs = layout
        .display_list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .flat_map(|line| &line.runs)
        .collect::<Vec<_>>();
    assert!(
        runs.iter()
            .any(|run| run.text == "PQ" && run.strike.is_none())
    );
    assert!(runs.iter().all(|run| matches!(
        run.strike.as_deref(),
        None | Some("sngStrike" | "dblStrike")
    )));
}
