use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

#[test]
fn run_gradients_render_the_lowest_stop_without_changing_controls() {
    let session = DeckSession::open(
        include_bytes!("../../pptx-parse/tests/fixtures/run-gradfill.pptx"),
        299,
    )
    .unwrap();
    let mut renderer = SlideRenderer::new();
    renderer
        .register_font(
            "Arial",
            false,
            false,
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"),
        )
        .unwrap();
    let list = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    let colors: Vec<_> = list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { paragraphs, .. } => Some(paragraphs),
            _ => None,
        })
        .flat_map(|paragraphs| paragraphs.iter().flat_map(|paragraph| &paragraph.runs))
        .map(|run| (run.text.as_str(), run.color.as_str()))
        .collect();
    assert_eq!(
        colors,
        [
            ("RGB gradient", "#FFFFFF"),
            ("Theme gradient", "#FFFFFF"),
            ("Modified gradient", "#90A0B0"),
            ("Unsorted ramp", "#00FF00"),
            ("First second", "#FFFFFF"),
            ("Solid control", "#FFFFFF"),
            ("Inherited control", "#505050"),
            ("Empty gradient", "#505050"),
        ],
    );
}
