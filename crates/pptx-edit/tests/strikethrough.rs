#[path = "../../pptx-parse/tests/common/strikethrough.rs"]
mod fixture;

use pptx_edit::engine::{Document, Host};
use pptx_edit::{DeckSession, EditCtx, EditRequest, TextStyle, TextStylePatch};

fn authored_properties(strike: &str) -> String {
    format!(
        r#"<a:rPr lang="en-GB" strike="{strike}"><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/></a:rPr>"#
    )
}

fn authored_deck(strike: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "<a:r><a:t>P</a:t></a:r>",
                    &format!(
                        "<a:r><a:t>left</a:t></a:r><a:r>{}<a:t>author</a:t></a:r>",
                        authored_properties(strike)
                    ),
                )
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

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

fn assert_authored_properties(bytes: &[u8], strike: &str, expected_text: &str, added: &str) {
    let xml = slide_xml(bytes);
    let mut text = String::new();
    let mut added_seen = added.is_empty();
    for run in xml.split("<a:r>").skip(1) {
        let run = run.split("</a:r>").next().unwrap();
        if !run.contains(r#"lang="en-GB""#) {
            continue;
        }
        let (properties, rest) = run.split_once("<a:t>").unwrap();
        assert_eq!(properties.replace(added, ""), authored_properties(strike));
        added_seen |= properties.contains(added);
        text.push_str(rest.strip_suffix("</a:t>").unwrap());
    }
    assert_eq!(text, expected_text);
    assert!(added_seen);
}

#[test]
fn inserting_inside_authored_strike_preserves_run_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        for underline in [None, Some("none".to_owned())] {
            let bytes = authored_deck(strike);
            let session = DeckSession::open(&bytes, 6412).unwrap();
            let snapshot = session.snapshot().unwrap();
            let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
            session
                .insert_text(
                    &EditCtx::local("strike"),
                    story_id,
                    6,
                    "X",
                    &TextStyle {
                        underline: underline.clone(),
                        ..TextStyle::default()
                    },
                )
                .unwrap();
            let added = if underline.is_some() {
                r#" u="none""#
            } else {
                ""
            };
            let saved = session.save().unwrap();
            assert_authored_properties(&saved, strike, "auXthor", added);
        }
    }
}

#[test]
fn raw_json_insert_preserves_authored_no_strike() {
    for style in [
        serde_json::json!({}),
        serde_json::json!({ "underline": "none" }),
    ] {
        let added = if style.get("underline").is_some() {
            r#" u="none""#
        } else {
            ""
        };
        let bytes = authored_deck("noStrike");
        let document = Document::open_collaborative(
            &bytes,
            6413.0,
            Host {
                entropy: || 0,
                timer_ms: || 0.0,
            },
        )
        .unwrap();
        let snapshot: serde_json::Value =
            serde_json::from_str(&document.snapshot_json().unwrap()).unwrap();
        let story_id = &snapshot["slides"][0]["shapes"][3]["textStories"][0]["id"];
        document
            .insert_text_json(
                &serde_json::json!({
                    "storyId": story_id,
                    "index": 6,
                    "text": "X",
                    "style": style,
                })
                .to_string(),
            )
            .unwrap();
        let saved = document.save_bytes().unwrap();
        assert_authored_properties(&saved, "noStrike", "auXthor", added);
    }
}

#[test]
fn deleting_inside_authored_strike_preserves_run_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6414).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        session
            .delete_text(&EditCtx::local("strike"), story_id, 6, 7)
            .unwrap();
        assert_authored_properties(&session.save().unwrap(), strike, "auhor", "");
    }
}

#[test]
fn bold_format_preserves_authored_strike_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6415).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                4,
                10,
                &TextStylePatch {
                    bold: Some(true),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        assert_authored_properties(&session.save().unwrap(), strike, "author", r#" b="1""#);
    }
}

#[test]
fn inserting_into_neighbour_preserves_authored_strike_properties() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6416).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        session
            .insert_text(
                &EditCtx::local("strike"),
                story_id,
                2,
                "X",
                &TextStyle::default(),
            )
            .unwrap();
        assert_authored_properties(&session.save().unwrap(), strike, "author", "");
    }
}

#[test]
fn explicit_clear_survives_text_edits_and_update_restore() {
    let bytes = authored_deck("sngStrike");
    let session = DeckSession::open(&bytes, 6417).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            4,
            10,
            &TextStylePatch {
                strike: Some("noStrike".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("strike"),
            story_id,
            6,
            "X",
            &TextStyle::default(),
        )
        .unwrap();
    let restored = DeckSession::open_from_update_with_source(
        &session.encode_state_as_update_v1(),
        &bytes,
        6418,
    )
    .unwrap();
    let saved = restored.save().unwrap();
    let xml = slide_xml(&saved);
    let expected = authored_properties("sngStrike").replace(r#" strike="sngStrike""#, "");
    assert!(xml.contains(&format!("<a:r>{expected}<a:t>auXthor</a:t></a:r>")));
    let reopened = DeckSession::open(&saved, 6419).unwrap();
    assert!(
        reopened.story(story_id).unwrap().paragraphs[0]
            .runs
            .iter()
            .filter(|run| run.text != "Q")
            .all(|run| run.style.strike.is_none())
    );
}

#[test]
fn clearing_neighbour_preserves_authored_no_strike_after_insert() {
    let session = DeckSession::open(&authored_deck("noStrike"), 6420).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            10,
            11,
            &TextStylePatch {
                strike: Some("noStrike".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("strike"),
            story_id,
            6,
            "X",
            &TextStyle::default(),
        )
        .unwrap();
    let saved = session.save().unwrap();
    assert_authored_properties(&saved, "noStrike", "auXthor", "");
    assert!(slide_xml(&saved).contains("<a:r><a:rPr/><a:t>Q</a:t></a:r>"));
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
        assert_eq!(
            reopened.story(story_id).unwrap().paragraphs[0].runs[0]
                .style
                .strike,
            None
        );
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
        assert_eq!(
            session.story(story_id).unwrap().paragraphs[0].runs[0]
                .style
                .strike
                .as_deref(),
            Some("noStrike")
        );
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

#[test]
fn clearing_across_paragraphs_uses_each_paragraphs_inheritance() {
    let session = DeckSession::open(&fixture::deck(), 6421).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[0].text_stories[0].id;
    session
        .format_text(
            &EditCtx::local("strike"),
            story_id,
            0,
            6,
            &TextStylePatch {
                strike: Some("noStrike".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let story = session.story(story_id).unwrap();
    assert!(
        story.paragraphs[0]
            .runs
            .iter()
            .all(|run| run.style.strike.as_deref() == Some("noStrike"))
    );
    assert!(
        story.paragraphs[1]
            .runs
            .iter()
            .all(|run| run.style.strike.is_none())
    );
    let reopened = DeckSession::open(&session.save().unwrap(), 6422).unwrap();
    assert_eq!(reopened.story(story_id).unwrap(), story);
}

#[test]
fn undoing_a_clear_restores_the_exact_stored_strike() {
    for strike in ["noStrike", "sngStrike", "dblStrike"] {
        let session = DeckSession::open(&authored_deck(strike), 6423).unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
        let before = session.story(story_id).unwrap();
        session.clear_undo_history();
        session
            .format_text(
                &EditCtx::local("strike"),
                story_id,
                4,
                10,
                &TextStylePatch {
                    strike: Some("noStrike".to_owned()),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        let cleared = session.story(story_id).unwrap();
        assert!(
            cleared.paragraphs[0]
                .runs
                .iter()
                .filter(|run| run.text != "Q")
                .all(|run| run.style.strike.is_none())
        );
        assert!(session.undo());
        assert_eq!(session.story(story_id).unwrap(), before);
        assert_authored_properties(&session.save().unwrap(), strike, "author", "");
        assert!(session.redo());
        assert_eq!(session.story(story_id).unwrap(), cleared);
    }
}

#[test]
fn batch_clear_removes_authored_no_strike_without_inherited_strike() {
    let session = DeckSession::open(&authored_deck("noStrike"), 6424).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[3];
    let story_id = &shape.text_stories[0].id;
    let before = session.story(story_id).unwrap();
    let request: EditRequest = serde_json::from_value(serde_json::json!({
        "expectVersion": session.version(),
        "steps": [{
            "op": "formatText",
            "target": {
                "kind": "range",
                "slideId": slide.id,
                "shapeId": shape.id,
                "storyId": story_id,
                "start": 4,
                "end": 10,
            },
            "patch": {"strike": "noStrike"},
        }],
    }))
    .unwrap();
    assert!(session.apply_edits(&request).unwrap().unwrap().applied);
    let cleared = session.story(story_id).unwrap();
    assert!(
        cleared.paragraphs[0]
            .runs
            .iter()
            .filter(|run| run.text != "Q")
            .all(|run| run.style.strike.is_none())
    );
    let reopened = DeckSession::open(&session.save().unwrap(), 6425).unwrap();
    assert_eq!(reopened.story(story_id).unwrap(), cleared);
    assert!(session.undo());
    assert_eq!(session.story(story_id).unwrap(), before);
    assert!(session.redo());
    assert_eq!(session.story(story_id).unwrap(), cleared);
}

#[test]
fn replay_clear_and_history_preserve_literal_strike_values() {
    for shape_index in [0, 1, 2, 3] {
        let bytes = authored_deck("noStrike");
        let session = DeckSession::open_replay_baseline(&bytes, Some(6426), None).unwrap();
        let identity = session.peer_identity(Vec::new()).unwrap();
        let peer = DeckSession::open_peer_deck(&bytes, &identity, None).unwrap();
        peer.register_peer_fonts(Vec::new()).unwrap();
        peer.adopt_peer_identity().unwrap();
        let snapshot = session.snapshot().unwrap();
        let story_id = &snapshot.slides[0].shapes[shape_index].text_stories[0].id;
        let (start, end, expected) = if shape_index == 3 {
            (4, 10, None)
        } else {
            (0, 1, Some("noStrike"))
        };
        let before = session.story(story_id).unwrap();
        session.clear_undo_history();
        let mut cleared = None;
        for (sequence, op) in [
            serde_json::json!({
                "method": "formatText",
                "args": [story_id, start, end, {"strike": "noStrike"}],
            }),
            serde_json::json!({"method": "undo", "args": []}),
            serde_json::json!({"method": "redo", "args": []}),
        ]
        .into_iter()
        .enumerate()
        {
            let envelope = serde_json::json!({
                "sequence": sequence + 1,
                "baseVersion": session.version(),
                "op": op,
            })
            .to_string();
            assert_eq!(
                session.replay_json(&envelope).unwrap(),
                peer.replay_json(&envelope).unwrap()
            );
            assert_eq!(
                peer.story(story_id).unwrap(),
                session.story(story_id).unwrap()
            );
            let story = session.story(story_id).unwrap();
            if sequence == 0 {
                let run = story.paragraphs[0]
                    .runs
                    .iter()
                    .find(|run| run.text != "left")
                    .unwrap();
                assert_eq!(run.style.strike.as_deref(), expected);
                cleared = Some(story);
            } else if sequence == 1 {
                assert_eq!(story, before);
            } else {
                assert_eq!(Some(story), cleared);
            }
            let restored = DeckSession::open_from_update_with_source(
                &session.encode_state_as_update_v1(),
                &bytes,
                6427,
            )
            .unwrap();
            assert_eq!(
                restored.story(story_id).unwrap(),
                session.story(story_id).unwrap()
            );
            let reopened = DeckSession::open(&session.save().unwrap(), 6428).unwrap();
            assert_eq!(
                reopened.story(story_id).unwrap(),
                session.story(story_id).unwrap()
            );
        }
    }
}

#[test]
fn inserted_explicit_clear_only_overrides_inherited_strike() {
    for shape_index in 0..4 {
        for authored in ["sngStrike", "noStrike"] {
            let bytes = authored_deck(authored);
            let session = DeckSession::open(&bytes, 6433).unwrap();
            let snapshot = session.snapshot().unwrap();
            let shape = &snapshot.slides[0].shapes[shape_index];
            let story_id = &shape.text_stories[0].id;
            let at = if shape_index == 3 { 6 } else { 1 };
            session
                .insert_text(
                    &EditCtx::local("strike"),
                    story_id,
                    at,
                    "X",
                    &TextStyle {
                        strike: Some("noStrike".to_owned()),
                        ..TextStyle::default()
                    },
                )
                .unwrap();
            let expected = (shape_index != 3).then_some("noStrike");
            let reopened = DeckSession::open(&session.save().unwrap(), 6434).unwrap();
            for session in [&session, &reopened] {
                let story = session.story(story_id).unwrap();
                let inserted = story.paragraphs[0]
                    .runs
                    .iter()
                    .find(|run| run.text.contains('X'))
                    .unwrap();
                assert_eq!(inserted.style.strike.as_deref(), expected);
            }
        }
    }
}

#[test]
fn replacement_runs_do_not_inherit_a_retained_prefix_strike() {
    let session = DeckSession::open(&fixture::deck(), 6435).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    let context = EditCtx::local("strike");
    session
        .format_text(
            &context,
            story_id,
            0,
            2,
            &TextStylePatch {
                strike: Some("sngStrike".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let drafts: Vec<pptx_edit::TextParagraphDraft> = serde_json::from_value(serde_json::json!([{
        "runs": [
            {"text": "P", "style": {"strike": "sngStrike"}},
            {"text": "Q", "style": {}},
        ],
    }]))
    .unwrap();
    session
        .set_story_paragraphs(&context, story_id, &drafts)
        .unwrap();
    let reopened = DeckSession::open(&session.save().unwrap(), 6436).unwrap();
    for session in [&session, &reopened] {
        assert_eq!(
            session.story(story_id).unwrap().paragraphs[0]
                .runs
                .iter()
                .map(|run| (run.text.clone(), run.style.strike.clone()))
                .collect::<Vec<_>>(),
            [
                ("P".to_owned(), Some("sngStrike".to_owned())),
                ("Q".to_owned(), None)
            ]
        );
    }
}

fn paragraph_strike_deck(first_struck: bool) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    let properties = |struck| {
        if struck {
            r#"<a:pPr><a:defRPr strike="sngStrike"/></a:pPr>"#
        } else {
            ""
        }
    };
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    r#"<a:p><a:r><a:t>P</a:t></a:r><a:r><a:rPr strike="sngStrike"/><a:t>Q</a:t></a:r></a:p>"#,
                    &format!(
                        "<a:p>{}<a:r><a:t>AB</a:t></a:r></a:p><a:p>{}<a:r><a:t>C</a:t></a:r></a:p>",
                        properties(first_struck),
                        properties(!first_struck),
                    ),
                )
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn inserted_paragraph_uses_preserved_defaults_and_shifted_source_identity() {
    for first_struck in [true, false] {
        let session = DeckSession::open(&paragraph_strike_deck(first_struck), 6437).unwrap();
        let snapshot = session.snapshot().unwrap();
        let slide = &snapshot.slides[0];
        let shape = &slide.shapes[3];
        let story_id = &shape.text_stories[0].id;
        let context = EditCtx::local("strike");
        let drafts: Vec<pptx_edit::TextParagraphDraft> =
            serde_json::from_value(serde_json::json!([
                {"runs": [{"text": "AB"}]},
                {"runs": [{"text": "X"}]},
                {"runs": [{"text": "C"}]},
            ]))
            .unwrap();
        session
            .set_story_paragraphs(&context, story_id, &drafts)
            .unwrap();
        let formatted = session
            .story_formatting(&slide.id, &shape.id, story_id)
            .unwrap();
        for (index, struck) in [first_struck, first_struck, !first_struck]
            .into_iter()
            .enumerate()
        {
            assert_eq!(
                formatted.paragraphs[index].runs[0].style.strike.as_deref(),
                struck.then_some("sngStrike")
            );
        }
        session
            .format_text(
                &context,
                story_id,
                3,
                4,
                &TextStylePatch {
                    strike: Some("noStrike".to_owned()),
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 6438).unwrap();
        for session in [&session, &reopened] {
            let story = session.story(story_id).unwrap();
            assert_eq!(
                story.paragraphs[1].runs[0].style.strike.as_deref(),
                first_struck.then_some("noStrike")
            );
            let formatted = session
                .story_formatting(&slide.id, &shape.id, story_id)
                .unwrap();
            assert!(!matches!(
                formatted.paragraphs[1].runs[0].style.strike.as_deref(),
                Some("sngStrike" | "dblStrike")
            ));
        }
    }
}

#[test]
fn split_and_deleted_paragraphs_resolve_defaults_by_identity() {
    for first_struck in [true, false] {
        let bytes = paragraph_strike_deck(first_struck);
        for split in [true, false] {
            let session = DeckSession::open(&bytes, 6439).unwrap();
            let snapshot = session.snapshot().unwrap();
            let slide = &snapshot.slides[0];
            let shape = &slide.shapes[3];
            let story_id = &shape.text_stories[0].id;
            let context = EditCtx::local("strike");
            let (index, expected) = if split {
                session
                    .insert_paragraph_break(&context, story_id, 1)
                    .unwrap();
                (1, first_struck)
            } else {
                let drafts: Vec<pptx_edit::TextParagraphDraft> =
                    serde_json::from_value(serde_json::json!([
                        {"runs": [{"text": "C"}]},
                    ]))
                    .unwrap();
                session
                    .set_story_paragraphs(&context, story_id, &drafts)
                    .unwrap();
                (0, !first_struck)
            };
            let reopened = DeckSession::open(&session.save().unwrap(), 6440).unwrap();
            for session in [&session, &reopened] {
                let story = session
                    .story_formatting(&slide.id, &shape.id, story_id)
                    .unwrap();
                assert_eq!(
                    story.paragraphs[index].runs[0].style.strike.as_deref(),
                    expected.then_some("sngStrike")
                );
            }
        }
    }
}

#[test]
fn typing_under_no_strike_paragraph_defaults_does_not_duplicate_the_override() {
    let mut parts = ooxml_opc::unzip_parts(&paragraph_strike_deck(true)).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    r#"<a:pPr><a:defRPr strike="sngStrike"/></a:pPr>"#,
                    r#"<a:pPr><a:defRPr strike="noStrike"/></a:pPr>"#,
                )
                .into_bytes();
        }
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 6442).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[3].text_stories[0].id;
    session
        .insert_text(
            &EditCtx::local("strike"),
            story_id,
            1,
            "X",
            &TextStyle {
                strike: Some("noStrike".to_owned()),
                ..TextStyle::default()
            },
        )
        .unwrap();
    let reopened = DeckSession::open(&session.save().unwrap(), 6443).unwrap();
    for session in [&session, &reopened] {
        let story = session.story(story_id).unwrap();
        assert!(
            story.paragraphs[0]
                .runs
                .iter()
                .all(|run| run.style.strike.is_none())
        );
        assert!(story.plain_text().contains('X'));
    }
}
