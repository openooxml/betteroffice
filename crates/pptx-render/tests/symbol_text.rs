use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/symbol-text.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
}

fn drawn(list: &SurfaceDisplayList, id: u32) -> String {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id, lines, ..
            } if *object_id == id => Some(
                lines
                    .iter()
                    .flat_map(|line| &line.runs)
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .unwrap()
}

fn stored(session: &DeckSession) -> Vec<String> {
    let snapshot = session.snapshot().unwrap();
    [10, 11, 12]
        .into_iter()
        .map(|id| {
            snapshot.slides[0]
                .shapes
                .iter()
                .find(|shape| shape.source_id == id)
                .unwrap()
                .text_stories[0]
                .plain_text()
        })
        .collect()
}

#[test]
fn symbol_face_run_text_draws_the_characters_its_slots_address() {
    let session = DeckSession::open(DECK, 5_319).unwrap();
    let saved = session.save().unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    assert_eq!(drawn(&list, 10).trim(), "α+β ≥ π");
    assert_eq!(drawn(&list, 11).trim(), "●•");
    assert_eq!(drawn(&list, 12).trim(), "a+b");

    let expected = ["a+b \u{f0b3} p", "l\u{f0e0}", "a+b"];
    assert_eq!(stored(&session), expected);
    assert_eq!(
        session.save().unwrap(),
        saved,
        "drawing changes nothing saved"
    );
    let reopened = DeckSession::open(&saved, 5_320).unwrap();
    assert_eq!(stored(&reopened), expected);
}
