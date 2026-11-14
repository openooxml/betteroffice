use pptx_edit::{DeckSession, EditCtx, TextParagraphDraft, TextRunDraft, TextStylePatch};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/run-spacing.pptx");
const STORY: &str = "story:slide:0:256:shape:0:0";
const SLIDE: &str = "ppt/slides/slide1.xml";

fn with_title(text: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
    let (_, xml) = parts.iter_mut().find(|(path, _)| path == SLIDE).unwrap();
    let source = String::from_utf8(xml.clone()).unwrap();
    let title = "<a:t>OUR GREEN INITIATIVES</a:t>";
    assert!(source.contains(title));
    *xml = source
        .replace(title, &format!("<a:t>{text}</a:t>"))
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn slide_xml(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, xml) = parts.iter().find(|(path, _)| path == SLIDE).unwrap();
    String::from_utf8(xml.clone()).unwrap()
}

#[test]
fn a_literal_cr_lf_or_cr_in_run_text_reads_as_one_line_feed() {
    for (source, expected) in [
        ("OUR GREEN\r\nINITIATIVES", "OUR GREEN\nINITIATIVES"),
        ("OUR GREEN\rINITIATIVES", "OUR GREEN\nINITIATIVES"),
        ("OUR GREEN&#13;INITIATIVES", "OUR GREEN\rINITIATIVES"),
    ] {
        let session = DeckSession::open(&with_title(source), 41001).unwrap();
        assert_eq!(
            session.story(STORY).unwrap().plain_text(),
            expected,
            "{source:?}"
        );
    }
}

#[test]
fn run_text_read_from_a_cr_lf_deck_can_be_written_back() {
    let session = DeckSession::open(&with_title("OUR GREEN\r\nINITIATIVES"), 41002).unwrap();
    let paragraphs: Vec<TextParagraphDraft> = session
        .story(STORY)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| TextParagraphDraft {
            alignment: None,
            bullet: None,
            runs: paragraph
                .runs
                .iter()
                .map(|run| TextRunDraft {
                    text: run.text.clone(),
                    style: run.style.clone(),
                })
                .collect(),
        })
        .collect();
    let written = session
        .set_story_paragraphs(&EditCtx::local("test"), STORY, &paragraphs)
        .unwrap();
    assert_eq!(written.plain_text(), "OUR GREEN\nINITIATIVES");
}

#[test]
fn restyling_a_cr_lf_run_saves_a_line_break_not_a_carriage_return() {
    let session = DeckSession::open(&with_title("OUR GREEN\r\nINITIATIVES"), 41003).unwrap();
    let end = session.story(STORY).unwrap().length - 1;
    session
        .format_text(
            &EditCtx::local("test"),
            STORY,
            0,
            end,
            &TextStylePatch {
                italic: Some(true),
                ..Default::default()
            },
        )
        .unwrap();
    let saved = session.save().unwrap();
    let xml = slide_xml(&saved);
    assert!(!xml.contains("&#xD;") && !xml.contains('\r'), "{xml}");
    let reopened = DeckSession::open(&saved, 41004).unwrap();
    assert_eq!(
        reopened.story(STORY).unwrap().plain_text(),
        "OUR GREEN\nINITIATIVES"
    );
}
