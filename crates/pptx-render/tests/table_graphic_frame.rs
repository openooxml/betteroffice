use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/table-basic.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn session() -> DeckSession {
    DeckSession::open(DECK, 4101).unwrap()
}

fn display_list(session: &DeckSession) -> SurfaceDisplayList {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

fn text_of(primitive: &Primitive) -> Option<String> {
    match primitive {
        Primitive::TextBox { lines, .. } => Some(
            lines
                .iter()
                .flat_map(|line| line.runs.iter().map(|run| run.text.as_str()))
                .collect(),
        ),
        _ => None,
    }
}

#[test]
fn a_table_frame_does_not_spill_its_first_cell_onto_the_slide() {
    let list = display_list(&session());
    let stray: Vec<String> = list.primitives.iter().filter_map(text_of).collect();
    assert!(
        stray.is_empty(),
        "a table frame must not paint a cell as loose slide text, but painted {stray:?}"
    );
}

#[test]
fn the_frame_still_paints_and_every_cell_stays_in_the_model() {
    let session = session();
    let labels: Vec<String> = display_list(&session)
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Placeholder { label, .. } => label.clone(),
            _ => None,
        })
        .collect();
    assert_eq!(labels, ["Table"]);
    let snapshot = session.snapshot().unwrap();
    let cells: Vec<String> = snapshot.slides[0].shapes[0]
        .text_stories
        .iter()
        .map(|story| story.plain_text())
        .collect();
    assert_eq!(
        cells,
        [
            "Header spans two",
            "",
            "Third",
            "Tall",
            "Middle",
            "Edge",
            "",
            "Below",
            "Last"
        ]
    );
}
