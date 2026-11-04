#[path = "../../pptx-parse/tests/common/strikethrough.rs"]
mod fixture;

use pptx_edit::{DeckSession, EditCtx, TextStylePatch};

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "ppt/slides/slide1.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

#[test]
fn unchanged_strikes_round_trip_byte_identically() {
    let bytes = fixture::deck();
    let session = DeckSession::open(&bytes, 6401).unwrap();
    assert_eq!(slide_xml(&session.save().unwrap()), slide_xml(&bytes));
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    for (shape, expected) in [
        (
            0,
            vec![
                Some("sngStrike"),
                Some("dblStrike"),
                Some("noStrike"),
                Some("dblStrike"),
                Some("noStrike"),
            ],
        ),
        (1, vec![Some("sngStrike"), Some("noStrike")]),
        (2, vec![Some("dblStrike"), Some("noStrike")]),
    ] {
        let shape = &slide.shapes[shape];
        let story = session
            .story_formatting(&slide.id, &shape.id, &shape.text_stories[0].id)
            .unwrap();
        assert_eq!(
            story
                .paragraphs
                .iter()
                .flat_map(|paragraph| &paragraph.runs)
                .map(|run| run.style.strike.as_deref())
                .collect::<Vec<_>>(),
            expected
        );
    }
}

#[test]
fn sets_and_clears_strike_without_rewriting_untouched_runs() {
    let bytes = fixture::deck();
    let session = DeckSession::open(&bytes, 6402).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    for strike in ["sngStrike", "dblStrike"] {
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                0,
                1,
                &TextStylePatch {
                    strike: Some(strike.to_owned()),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        let remote =
            DeckSession::open_from_update(&session.encode_state_as_update_v1(), 6409).unwrap();
        assert_eq!(
            remote.story(story_id).unwrap(),
            session.story(story_id).unwrap()
        );
        let saved = session.save().unwrap();
        let reopened = DeckSession::open(&saved, 6403).unwrap();
        assert_eq!(
            reopened.story(story_id).unwrap().paragraphs[0].runs[0]
                .style
                .strike
                .as_deref(),
            Some(strike)
        );
        assert!(slide_xml(&saved).contains(
            r#"<a:r><a:rPr lang="en-US" strike="sngStrike" dirty="0"/><a:t>A</a:t></a:r>"#
        ));
        reopened
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                0,
                1,
                &TextStylePatch {
                    strike: Some("noStrike".to_owned()),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        let cleared = reopened.save().unwrap();
        let reopened = DeckSession::open(&cleared, 6404).unwrap();
        assert_eq!(
            reopened.story(story_id).unwrap().paragraphs[0].runs[0]
                .style
                .strike,
            None
        );
        let xml = slide_xml(&cleared);
        let plain = xml.split(r#"name="Strike 5""#).nth(1).unwrap();
        assert!(!plain.split("</a:r>").next().unwrap().contains("strike="));
    }
}

#[test]
fn clearing_inherited_strike_saves_an_explicit_override() {
    let session = DeckSession::open(&fixture::deck(), 6405).unwrap();
    let snapshot = session.snapshot().unwrap();
    for index in 0..3 {
        let shape = &snapshot.slides[0].shapes[index];
        let story_id = &shape.text_stories[0].id;
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                0,
                1,
                &TextStylePatch {
                    strike: Some("noStrike".to_owned()),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 6406).unwrap();
        let story = reopened
            .story_formatting(&snapshot.slides[0].id, &shape.id, story_id)
            .unwrap();
        assert_eq!(
            story.paragraphs[0].runs[0].style.strike.as_deref(),
            Some("noStrike")
        );
    }
}

#[test]
fn changing_a_sibling_run_keeps_authored_strike_markup() {
    let session = DeckSession::open(&fixture::deck(), 6408).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[0].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            1,
            2,
            &TextStylePatch {
                strike: Some("sngStrike".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let xml = slide_xml(&session.save().unwrap());
    assert!(
        xml.contains(
            r#"<a:r><a:rPr lang="en-US" strike="sngStrike" dirty="0"/><a:t>A</a:t></a:r>"#
        )
    );
    assert!(xml.contains(r#"<a:r><a:rPr strike="noStrike"/><a:t>C</a:t></a:r>"#));
    let before = session.encode_state_as_update_v1();
    assert!(
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                0,
                1,
                &TextStylePatch {
                    strike: Some("invalid".to_owned()),
                    ..TextStylePatch::default()
                }
            )
            .is_err()
    );
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn changing_another_mark_preserves_authored_no_strike() {
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "<a:t>P</a:t>",
                    r#"<a:t>P</a:t></a:r><a:r><a:rPr strike="noStrike"/><a:t>off</a:t>"#,
                )
                .into_bytes();
        }
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 6410).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            1,
            2,
            &TextStylePatch {
                bold: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let reopened = DeckSession::open(&session.save().unwrap(), 6411).unwrap();
    let story = reopened.story(story_id).unwrap();
    assert_eq!(
        story.paragraphs[0]
            .runs
            .iter()
            .map(|run| (run.text.as_str(), run.style.strike.as_deref()))
            .collect::<Vec<_>>(),
        [
            ("P", None),
            ("o", Some("noStrike")),
            ("ff", Some("noStrike")),
            ("Q", Some("sngStrike"))
        ]
    );
}
