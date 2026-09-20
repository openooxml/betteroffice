use pptx_edit::{DeckSession, TextCaps};
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/run-caps.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn renderer() -> SlideRenderer {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
}

fn lines(list: &SurfaceDisplayList, id: u32) -> &[PositionedTextLine] {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id, lines, ..
            } if *object_id == id => Some(lines.as_slice()),
            _ => None,
        })
        .unwrap()
}

fn drawn(list: &SurfaceDisplayList, id: u32) -> String {
    lines(list, id)[0]
        .runs
        .iter()
        .map(|run| run.text.as_str())
        .collect()
}

#[test]
fn cap_cases_a_run_for_drawing_only() {
    let session = DeckSession::open(DECK, 4_271).unwrap();
    let snapshot = session.snapshot().unwrap();
    let list = renderer()
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
        .display_list;
    assert_eq!(drawn(&list, 10), "MIXED CASE TITLE");
    assert_eq!(drawn(&list, 11), "SMALL CAPS RUN");
    assert_eq!(drawn(&list, 12), "INHERITED CAPS");
    assert_eq!(drawn(&list, 13), "Kept as authored");

    let stored = snapshot.slides[0]
        .shapes
        .iter()
        .map(|shape| shape.text_stories[0].plain_text())
        .collect::<Vec<_>>();
    assert_eq!(
        stored,
        [
            "Mixed Case Title",
            "Small Caps Run",
            "Inherited Caps",
            "Kept as authored"
        ]
    );
    assert_eq!(
        snapshot.slides[0].shapes[0].text_stories[0].paragraphs[0].runs[0]
            .style
            .caps,
        Some(TextCaps::All)
    );
}

#[test]
fn small_caps_draw_their_lowercase_stretches_smaller() {
    let session = DeckSession::open(DECK, 4_271).unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    let sizes = lines(&list, 11)[0]
        .runs
        .iter()
        .map(|run| (run.text.clone(), run.font_size_px))
        .collect::<Vec<_>>();
    assert_eq!(
        sizes,
        [
            ("S".to_owned(), 128.0 / 3.0),
            ("MALL".to_owned(), 128.0 / 3.0 * 0.8),
            (" C".to_owned(), 128.0 / 3.0),
            ("APS".to_owned(), 128.0 / 3.0 * 0.8),
            (" R".to_owned(), 128.0 / 3.0),
            ("UN".to_owned(), 128.0 / 3.0 * 0.8),
        ]
    );
    let all_caps = &lines(&list, 10)[0];
    assert_eq!(all_caps.runs.len(), 1);
    assert_eq!(all_caps.runs[0].font_size_px, 128.0 / 3.0);
    assert_eq!(lines(&list, 11)[0].height, all_caps.height);
}

#[test]
fn caret_offsets_and_saved_bytes_ignore_the_cased_drawing() {
    let session = DeckSession::open(DECK, 4_271).unwrap();
    let list = renderer()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    for id in [10, 11, 12, 13] {
        let line = &lines(&list, id)[0];
        assert_eq!(line.runs.first().unwrap().start, 0);
        assert_eq!(
            line.runs.last().unwrap().end,
            line.runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
        );
    }
    let saved = ooxml_opc::unzip_parts(&session.save().unwrap())
        .unwrap()
        .into_iter()
        .collect::<std::collections::HashMap<_, _>>();
    for (path, bytes) in ooxml_opc::unzip_parts(DECK).unwrap() {
        assert_eq!(saved.get(&path).map(Vec::as_slice), Some(bytes.as_slice()));
    }
}
