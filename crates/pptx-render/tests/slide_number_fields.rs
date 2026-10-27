#[path = "../../pptx-parse/tests/support/header_footer.rs"]
mod support;

use pptx_edit::{DeckSession, EditCtx, ReadRequest, TextParagraphDraft, TextRunDraft, TextStyle};
use pptx_render::{Primitive, SlideRenderer};
use support::{deck, field, placeholder};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn number_shape(paragraphs: &str) -> String {
    placeholder(2, "sldNum", paragraphs, false)
}

fn field_paragraph() -> String {
    format!("<a:p>{}</a:p>", field("slidenum", "9"))
}

fn open(bytes: &[u8]) -> DeckSession {
    DeckSession::open(bytes, 920).unwrap()
}

fn read_number(session: &DeckSession, index: usize) -> String {
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    read.slides[index]
        .header_footer
        .as_ref()
        .unwrap()
        .slide_number
        .as_ref()
        .unwrap()
        .text
        .clone()
}

fn rendered_number(session: &DeckSession, index: usize) -> String {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot.slides[index]
        .shapes
        .iter()
        .find(|shape| {
            shape
                .placeholder
                .as_ref()
                .and_then(|placeholder| placeholder.placeholder_type.as_deref())
                == Some("sldNum")
        })
        .unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &snapshot, index)
        .unwrap();
    rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                shape_id: Some(id),
                paragraphs,
                ..
            } if id == &shape.id => Some(
                paragraphs
                    .iter()
                    .flat_map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()))
                    .collect::<String>(),
            ),
            _ => None,
        })
        .unwrap()
}

#[test]
fn stale_slide_owned_number_renders_live_and_agrees_with_read() {
    let shape = number_shape(&field_paragraph());
    let bytes = deck(&[("", false), (&shape, false)], "", "", "", Some(1));
    let session = open(&bytes);
    assert_eq!(rendered_number(&session, 1), "2");
    assert_eq!(read_number(&session, 1), "2");
}

#[test]
fn number_next_to_literal_text_renders_with_the_literal() {
    let paragraphs = format!(
        "<a:p><a:r><a:t>Seite </a:t></a:r>{}</a:p>",
        field("slidenum", "9")
    );
    let shape = number_shape(&paragraphs);
    let bytes = deck(&[("", false), (&shape, false)], "", "", "", Some(1));
    let session = open(&bytes);
    assert_eq!(rendered_number(&session, 1), "Seite 2");
    assert_eq!(read_number(&session, 1), "Seite 2");
}

#[test]
fn moved_slide_renders_and_reads_its_new_number() {
    let shape = number_shape(&field_paragraph());
    let bytes = deck(&[("", false), (&shape, false)], "", "", "", Some(1));
    let session = open(&bytes);
    let slide_id = session.slide_ids().unwrap()[1].clone();
    session
        .move_slide(&EditCtx::local("test"), &slide_id, 0)
        .unwrap();
    assert_eq!(rendered_number(&session, 0), "1");
    assert_eq!(read_number(&session, 0), "1");
}

#[test]
fn edit_inside_the_field_keeps_the_edited_text() {
    let shape = number_shape(&field_paragraph());
    let bytes = deck(&[("", false), (&shape, false)], "", "", "", Some(1));
    let session = open(&bytes);
    let snapshot = session.snapshot().unwrap();
    let story = snapshot.slides[1].shapes[0].text_stories[0].clone();
    session
        .set_story_paragraphs(
            &EditCtx::local("test"),
            &story.id,
            &[TextParagraphDraft {
                alignment: None,
                bullet: None,
                runs: vec![TextRunDraft {
                    text: "8".to_owned(),
                    style: TextStyle::default(),
                }],
            }],
        )
        .unwrap();
    assert_eq!(rendered_number(&session, 1), "8");
    assert_eq!(read_number(&session, 1), "8");
}

#[test]
fn first_slide_num_offsets_the_rendered_number() {
    let shape = number_shape(&field_paragraph());
    let bytes = deck(&[("", false), (&shape, false)], "", "", "", Some(10));
    let session = open(&bytes);
    assert_eq!(rendered_number(&session, 1), "11");
    assert_eq!(read_number(&session, 1), "11");
}
