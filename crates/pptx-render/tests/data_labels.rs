use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/data-labels.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn explicit_data_labels_preserve_axes_bars_and_overrides() {
    let session = DeckSession::open(FIXTURE, 305).unwrap();
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    let mut reference_shapes = None;
    let label_texts: &[&[&str]] = &[
        &[],
        &[],
        &["Q1", "Q2", "Q3", "Q1", "Q2", "Q3"],
        &["19"],
        &[],
    ];
    assert_eq!(snapshot.slides.len(), label_texts.len());
    for (index, labels) in label_texts.iter().enumerate() {
        let list = renderer
            .layout_slide(session.package(), &snapshot, index)
            .unwrap()
            .display_list;
        let chart = list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Chart { primitives, .. } => Some(primitives),
                _ => None,
            })
            .unwrap();
        let mut texts: Vec<_> = chart
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::TextBox { paragraphs, .. } => Some(
                    paragraphs
                        .iter()
                        .flat_map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()))
                        .collect::<String>(),
                ),
                _ => None,
            })
            .collect();
        let mut expected = vec![
            "Revenue", "0", "6", "13", "19", "25", "Q1", "Q2", "Q3", "Quarter", "Millions",
            "North", "South",
        ];
        expected.extend_from_slice(labels);
        texts.sort();
        expected.sort();
        assert_eq!(texts, expected, "slide {}", index + 1);
        let shapes: Vec<_> = chart
            .iter()
            .filter(|primitive| matches!(primitive, Primitive::Shape { .. }))
            .cloned()
            .collect();
        for color in ["#6254E7", "#1FA97A"] {
            assert_eq!(
                shapes
                    .iter()
                    .filter(|primitive| matches!(primitive, Primitive::Shape {
                        fill: Some(Paint::Solid { color: actual }), ..
                    } if actual == color))
                    .count(),
                4,
                "three bars and one legend swatch for {color}"
            );
        }
        if let Some(reference) = &reference_shapes {
            assert_eq!(&shapes, reference, "slide {} geometry", index + 1);
        } else {
            reference_shapes = Some(shapes);
        }
    }
}
