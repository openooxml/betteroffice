use pptx_edit::{DeckSession, EditCtx, TextStylePatch};

const FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/run-gradfill.pptx");

fn gradients(bytes: &[u8]) -> Vec<String> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, slide) = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    String::from_utf8(slide.clone())
        .unwrap()
        .split("<a:gradFill ")
        .skip(1)
        .map(|tail| tail.split_once("</a:gradFill>").unwrap().0.to_owned())
        .collect()
}

fn assert_text_edits_preserve_run_gradients(bytes: &[u8], indices: std::ops::Range<usize>) {
    let expected = gradients(bytes);
    assert_eq!(expected.len(), 6);
    for index in indices {
        for insertion in [None, Some("X"), Some("😀")] {
            let session = DeckSession::open(bytes, 299).unwrap();
            let snapshot = session.snapshot().unwrap();
            let shape = &snapshot.slides[0].shapes[index];
            let story = &shape.text_stories[0];
            let context = EditCtx::local("test");
            if let Some(insertion) = insertion {
                session
                    .insert_text(
                        &context,
                        &story.id,
                        1,
                        insertion,
                        &story.paragraphs[0].runs[0].style,
                    )
                    .unwrap();
            } else {
                session
                    .format_text(
                        &context,
                        &story.id,
                        0,
                        story.length - 1,
                        &TextStylePatch {
                            bold: Some(true),
                            ..Default::default()
                        },
                    )
                    .unwrap();
            }
            let saved = session.save().unwrap();
            assert_eq!(
                gradients(&saved),
                expected,
                "{}: insertion={insertion:?}",
                shape.name
            );
            let reopened = DeckSession::open(&saved, 300).unwrap();
            assert_eq!(
                reopened.story(&story.id).unwrap(),
                session.story(&story.id).unwrap()
            );
        }
    }
}

#[test]
fn text_edits_preserve_run_gradients() {
    assert_text_edits_preserve_run_gradients(FIXTURE, 0..4);
}

#[test]
fn text_edits_preserve_adjacent_run_gradients() {
    assert_text_edits_preserve_run_gradients(FIXTURE, 4..5);
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let (_, slide) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    *slide = String::from_utf8(slide.clone())
        .unwrap()
        .replace(
            "<a:t>First </a:t></a:r><a:r>",
            "<a:t>First </a:t></a:r><a:br/><a:r>",
        )
        .into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    assert_text_edits_preserve_run_gradients(&bytes, 4..5);
}

#[test]
fn color_edits_replace_run_gradients() {
    let expected = gradients(FIXTURE);
    for index in 0..5 {
        let session = DeckSession::open(FIXTURE, 299).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story = &snapshot.slides[0].shapes[index].text_stories[0];
        session
            .format_text(
                &EditCtx::local("test"),
                &story.id,
                0,
                story.length - 1,
                &TextStylePatch {
                    color: Some("#123456".into()),
                    ..Default::default()
                },
            )
            .unwrap();
        let saved = session.save().unwrap();
        let mut remaining = expected.clone();
        remaining.drain(index..if index == 4 { 6 } else { index + 1 });
        assert_eq!(gradients(&saved), remaining);
        let reopened = DeckSession::open(&saved, 300).unwrap();
        assert_eq!(
            reopened.story(&story.id).unwrap(),
            session.story(&story.id).unwrap()
        );
    }
}

#[test]
fn main_gradient_snapshot_saves_source_parts_unchanged() {
    let session = DeckSession::open_from_update_with_source(
        include_bytes!("fixtures/run-gradfill-main-v8.update.bin"),
        FIXTURE,
        300,
    )
    .unwrap();
    assert_eq!(
        ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap(),
        ooxml_opc::unzip_parts(FIXTURE).unwrap(),
    );
}
