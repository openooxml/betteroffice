#[path = "../../pptx-parse/tests/common/strikethrough.rs"]
mod fixture;

use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};

const AUTHORED: [&str; 5] = [
    r#"strike="noStrike""#,
    r#"u="none""#,
    r#"b="0""#,
    r#"i="0""#,
    r#"strike="noStrike" u="none" b="0" i="0" lang="en-GB" dirty="0""#,
];
const PARAGRAPH_LENGTH: u32 = 16;
const AUTHOR_START: u32 = 4;
const AUTHOR_END: u32 = 10;

fn authored_deck() -> Vec<u8> {
    let paragraphs = AUTHORED
        .iter()
        .map(|attributes| {
            format!(
                r#"<a:p><a:r><a:t>left</a:t></a:r><a:r><a:rPr {attributes}/><a:t>author</a:t></a:r><a:r><a:t>right</a:t></a:r></a:p>"#
            )
        })
        .collect::<String>();
    let mut parts = ooxml_opc::unzip_parts(&fixture::deck()).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let original =
        r#"<a:p><a:r><a:t>P</a:t></a:r><a:r><a:rPr strike="sngStrike"/><a:t>Q</a:t></a:r></a:p>"#;
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert_eq!(xml.matches(original).count(), 1);
    *bytes = xml.replace(original, &paragraphs).into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open_deck(bytes: &[u8]) -> (DeckSession, String) {
    let session = DeckSession::open(bytes, 6701).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.source_id == 5)
        .unwrap();
    let story_id = shape.text_stories[0].id.clone();
    (session, story_id)
}

fn insertion_style(case: usize) -> TextStyle {
    TextStyle {
        bold: AUTHORED[case].contains(r#"b="0""#).then_some(false),
        italic: AUTHORED[case].contains(r#"i="0""#).then_some(false),
        underline: AUTHORED[case]
            .contains(r#"u="none""#)
            .then(|| "none".to_owned()),
        ..TextStyle::default()
    }
}

fn without_attribute(properties: &str, name: &str) -> String {
    let prefix = format!(" {name}=\"");
    let Some(start) = properties.find(&prefix) else {
        return properties.to_owned();
    };
    let value_start = start + prefix.len();
    let end = value_start + properties[value_start..].find('"').unwrap() + 1;
    format!("{}{}", &properties[..start], &properties[end..])
}

fn assert_authored_run_properties(
    bytes: &[u8],
    edited_case: usize,
    edited_left: &str,
    edited_author: &str,
    target: Option<(&str, &str)>,
) {
    let (_, bytes) = ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let xml = String::from_utf8(bytes).unwrap();
    let shape = xml.rsplit_once("<p:sp>").unwrap().1;
    let shape = shape.split("</p:sp>").next().unwrap();
    let paragraphs = shape.split("<a:p>").skip(1).collect::<Vec<_>>();
    assert_eq!(paragraphs.len(), AUTHORED.len());
    for (case, paragraph) in paragraphs.iter().enumerate() {
        let paragraph = paragraph.split("</a:p>").next().unwrap();
        let (left, author, target) = if case == edited_case {
            (edited_left, edited_author, target)
        } else {
            ("left", "author", None)
        };
        let authored = format!("<a:rPr {}/>", AUTHORED[case]);
        let mut text = String::new();
        let mut author_runs = 0;
        for run in paragraph.split("<a:r>").skip(1) {
            let run = run.split("</a:r>").next().unwrap();
            let (properties, run_text) = run.split_once("<a:t>").unwrap();
            let run_text = run_text.strip_suffix("</a:t>").unwrap();
            let start = text.len();
            text.push_str(run_text);
            let end = text.len();
            if start < left.len() + author.len() && end > left.len() {
                assert!(start >= left.len() && end <= left.len() + author.len());
                author_runs += 1;
                if let Some((name, value)) = target {
                    assert!(properties.contains(&format!(" {name}=\"{value}\"")));
                    assert_eq!(
                        without_attribute(properties, name),
                        without_attribute(&authored, name),
                        "case {case}: untargeted attributes must remain byte-identical"
                    );
                } else {
                    assert_eq!(
                        properties, authored,
                        "case {case}: authored defaults must remain byte-identical"
                    );
                }
            }
        }
        assert!(author_runs > 0, "case {case}: authored run must survive");
        assert_eq!(text, format!("{left}{author}right"), "case {case}");
    }
}

#[test]
fn inserting_inside_a_run_preserves_every_untargeted_authored_property() {
    let bytes = authored_deck();
    for case in 0..AUTHORED.len() {
        let (session, story_id) = open_deck(&bytes);
        session
            .insert_text(
                &EditCtx::local("authored"),
                &story_id,
                case as u32 * PARAGRAPH_LENGTH + AUTHOR_START + 3,
                "X",
                &insertion_style(case),
            )
            .unwrap();
        assert_authored_run_properties(&session.save().unwrap(), case, "left", "autXhor", None);
    }
}

#[test]
fn deleting_inside_a_run_preserves_every_untargeted_authored_property() {
    let bytes = authored_deck();
    for case in 0..AUTHORED.len() {
        let (session, story_id) = open_deck(&bytes);
        let index = case as u32 * PARAGRAPH_LENGTH + AUTHOR_START + 3;
        session
            .delete_text(&EditCtx::local("authored"), &story_id, index, index + 1)
            .unwrap();
        assert_authored_run_properties(&session.save().unwrap(), case, "left", "autor", None);
    }
}

#[test]
fn inserting_into_a_neighbour_preserves_every_untargeted_authored_property() {
    let bytes = authored_deck();
    for case in 0..AUTHORED.len() {
        let (session, story_id) = open_deck(&bytes);
        session
            .insert_text(
                &EditCtx::local("authored"),
                &story_id,
                case as u32 * PARAGRAPH_LENGTH + 2,
                "X",
                &TextStyle::default(),
            )
            .unwrap();
        assert_authored_run_properties(&session.save().unwrap(), case, "leXft", "author", None);
    }
}

#[test]
fn formatting_one_property_preserves_every_other_authored_property() {
    let bytes = authored_deck();
    for case in 0..AUTHORED.len() {
        let (patch, target) = match case {
            2 => (
                TextStylePatch {
                    italic: Some(true),
                    ..TextStylePatch::default()
                },
                ("i", "1"),
            ),
            3 => (
                TextStylePatch {
                    underline: Some("sng".to_owned()),
                    ..TextStylePatch::default()
                },
                ("u", "sng"),
            ),
            _ => (
                TextStylePatch {
                    bold: Some(true),
                    ..TextStylePatch::default()
                },
                ("b", "1"),
            ),
        };
        let (session, story_id) = open_deck(&bytes);
        let offset = case as u32 * PARAGRAPH_LENGTH;
        session
            .format_text(
                &EditCtx::local("authored"),
                &story_id,
                offset + AUTHOR_START,
                offset + AUTHOR_END,
                &patch,
            )
            .unwrap();
        assert_authored_run_properties(
            &session.save().unwrap(),
            case,
            "left",
            "author",
            Some(target),
        );
    }
}

#[test]
fn formatting_strike_preserves_every_other_authored_property_branch() {
    let bytes = authored_deck();
    let patch: TextStylePatch =
        serde_json::from_value(serde_json::json!({ "strike": "sngStrike" })).unwrap();
    for case in 1..AUTHORED.len() {
        let (session, story_id) = open_deck(&bytes);
        let offset = case as u32 * PARAGRAPH_LENGTH;
        session
            .format_text(
                &EditCtx::local("authored"),
                &story_id,
                offset + AUTHOR_START,
                offset + AUTHOR_END,
                &patch,
            )
            .unwrap();
        assert_authored_run_properties(
            &session.save().unwrap(),
            case,
            "left",
            "author",
            Some(("strike", "sngStrike")),
        );
    }
}
