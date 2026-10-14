use pptx_edit::DeckSession;
use pptx_render::{
    PositionedTextLine, Primitive, RenderedSlide, SlideRenderer, SurfaceDisplayList,
};

const DECK: &[u8] = include_bytes!("fixtures/theme-fonts.pptx");
const SANS: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf");
const SERIF: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSerif-Regular.ttf");
const REGISTERED: [(&str, &[u8]); 3] = [
    ("Arial", SANS),
    ("Liberation Sans", SANS),
    ("Liberation Serif", SERIF),
];

fn rendered(slide: usize, families: &[(&str, &[u8])]) -> RenderedSlide {
    let session = DeckSession::open(DECK, 291).unwrap();
    let mut renderer = SlideRenderer::new();
    for (family, bytes) in families {
        renderer.register_font(family, false, false, bytes).unwrap();
    }
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), slide)
        .unwrap()
}

fn render(slide: usize) -> SurfaceDisplayList {
    rendered(slide, &REGISTERED).display_list
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

#[test]
fn major_theme_runs_use_the_major_face_and_metrics() {
    let list = render(0);
    let title = &lines(&list, 2)[0];
    assert_eq!(title.runs[0].font_family, "Liberation Serif");
    assert_eq!(title.runs[0].font_id, 2);
    assert_eq!(title.runs[0].color, "#17365D");
    assert!((title.width - 437.0625).abs() < 0.001);
    let body = lines(&list, 3);
    assert_eq!(body.len(), 3);
    for (line, expected) in body.iter().zip([
        "Theme major font chooses the ",
        "heading face and its own ",
        "wrapping metrics.",
    ]) {
        assert_eq!(line.runs.len(), 1);
        assert_eq!(line.runs[0].text, expected);
        assert_eq!(line.runs[0].font_id, 2);
        assert_eq!(line.runs[0].color, "#17365D");
    }
    let minor = &lines(&list, 4)[0].runs[0];
    assert_eq!(minor.font_family, "Liberation Sans");
    assert_eq!(minor.font_id, 1);
    assert_eq!(minor.color, "#008080");
    let explicit = &lines(&list, 5)[0].runs[0];
    assert_eq!(explicit.font_family, "Arial");
    assert_eq!(explicit.font_id, 0);
    assert_eq!(explicit.color, "#7F3F00");
}

fn shape_id(list: &SurfaceDisplayList, id: u32) -> &str {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id,
                shape_id,
                ..
            } if *object_id == id => shape_id.as_deref(),
            _ => None,
        })
        .unwrap()
}

fn run_texts(list: &SurfaceDisplayList) -> Vec<String> {
    list.primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .flat_map(|line| &line.runs)
        .map(|run| run.text.clone())
        .filter(|text| !text.trim().is_empty())
        .collect()
}

#[test]
fn a_registered_family_reports_no_substitution() {
    assert!(rendered(0, &REGISTERED).font_substitutions.is_empty());
}

#[test]
fn a_missing_family_reports_the_family_it_drew_instead() {
    let slide = rendered(0, &[("Arial", SANS), ("Liberation Sans", SANS)]);
    assert_eq!(slide.font_substitutions.len(), 1);
    let reported = &slide.font_substitutions[0];
    assert_eq!(reported.requested_family, "Liberation Serif");
    assert_eq!(reported.selected_family, "Arial");
    assert_eq!(reported.shape_id, shape_id(&slide.display_list, 2));
}

#[test]
fn two_shapes_missing_one_family_report_it_once() {
    let slide = rendered(0, &[("Arial", SANS), ("Liberation Sans", SANS)]);
    for id in [2, 3] {
        assert_eq!(
            lines(&slide.display_list, id)[0].runs[0].font_family,
            "Arial"
        );
    }
    assert_eq!(slide.font_substitutions.len(), 1);
}

#[test]
fn a_substitution_carries_no_document_text() {
    let slide = rendered(1, &[("Arial", SANS)]);
    let reported = format!("{:?}", slide.font_substitutions);
    assert!(reported.contains("Liberation Serif Light"));
    assert!(reported.contains("Arial"));
    let texts = run_texts(&slide.display_list);
    assert!(texts.iter().any(|text| text.contains("keeps the")));
    for text in texts {
        assert!(!reported.contains(&text), "{reported} leaked {text:?}");
    }
}

#[test]
fn literal_font_names_keep_the_configured_fallback() {
    let list = render(1);
    for (id, width) in [(2, 690.15625), (3, 754.1719), (4, 725.6875)] {
        let line = &lines(&list, id)[0];
        assert_eq!(line.runs[0].font_family, "Arial");
        assert_eq!(line.runs[0].font_id, 0);
        assert!(!line.runs[0].bold);
        assert!((line.width - width).abs() < 0.001);
    }
    let registered = &lines(&list, 5)[0].runs[0];
    assert_eq!(registered.font_family, "Liberation Serif");
    assert_eq!(registered.font_id, 2);
}
