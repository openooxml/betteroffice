use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_parse::{Bullet, ShapeNode};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/autonumber-bullets.pptx");

fn restart_bullets(session: &DeckSession) -> Vec<Bullet> {
    let shape = session.package().slides[0]
        .shapes
        .iter()
        .find_map(|node| match node {
            ShapeNode::Shape(shape) if shape.base.id == 12 => Some(shape),
            _ => None,
        })
        .unwrap();
    shape
        .text
        .as_ref()
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.properties.bullet.clone().unwrap())
        .collect()
}

#[test]
fn explicit_numbering_restarts_survive_text_edits_and_save() {
    let session = DeckSession::open(DECK, 30010).unwrap();
    let expected = restart_bullets(&session);
    assert!(matches!(
        expected[2],
        Bullet::AutoNumber {
            start_at: 1,
            restart: true,
            ..
        }
    ));
    let snapshot = session.snapshot().unwrap();
    let story = &snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "List 12")
        .unwrap()
        .text_stories[0];
    let restart_offset = story.paragraphs[..2]
        .iter()
        .map(|paragraph| {
            1 + paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
        })
        .sum::<u32>();
    session
        .set_paragraph_alignment(
            &EditCtx::local("numbered"),
            &story.id,
            restart_offset,
            restart_offset,
            Some("ctr"),
        )
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("numbered"),
            &story.id,
            0,
            "Edited ",
            &TextStyle::default(),
        )
        .unwrap();
    let reopened = DeckSession::open(&session.save().unwrap(), 30011).unwrap();
    assert_eq!(restart_bullets(&reopened), expected);
    let snapshot = reopened.snapshot().unwrap();
    assert!(
        snapshot.slides[0]
            .shapes
            .iter()
            .find(|shape| shape.name == "List 12")
            .unwrap()
            .text_stories[0]
            .paragraphs[0]
            .runs[0]
            .text
            .starts_with("Edited ")
    );
}
