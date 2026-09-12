use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_parse::{LineSpacing, PptxPackage, ShapeNode};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/paragraph-spacing.pptx");

fn spacing(package: &PptxPackage) -> Vec<(Option<LineSpacing>, Option<LineSpacing>)> {
    package.slides[0]
        .shapes
        .iter()
        .filter_map(|shape| match shape {
            ShapeNode::Shape(shape) => shape.text.as_ref(),
            _ => None,
        })
        .flat_map(|text| &text.paragraphs)
        .map(|paragraph| {
            (
                paragraph.properties.space_before,
                paragraph.properties.space_after,
            )
        })
        .collect()
}

#[test]
fn saving_an_edited_deck_keeps_the_paragraph_spacing() {
    let session = DeckSession::open(DECK, 4_243).unwrap();
    let story = session.snapshot().unwrap().slides[0].shapes[4].text_stories[0]
        .id
        .clone();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "Edited ",
            &TextStyle::default(),
        )
        .unwrap();
    session
        .set_paragraph_alignment(&EditCtx::local("test"), &story, 0, 1, Some("ctr"))
        .unwrap();
    let saved = session.save().unwrap();

    let before = spacing(&pptx_parse::parse_pptx(DECK).unwrap());
    assert!(before.iter().any(|pair| pair.0.is_some()));
    assert!(before.iter().any(|pair| pair.1.is_some()));
    assert_eq!(spacing(&pptx_parse::parse_pptx(&saved).unwrap()), before);
}
