#[allow(dead_code)]
#[path = "../../pptx-parse/tests/support/header_footer.rs"]
mod support;

use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_render::{Primitive, SlideRenderer, TextLanguage, TextParagraph};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn run(text: &str, language: &str) -> String {
    format!(r#"<a:r><a:rPr lang="{language}"/><a:t>{text}</a:t></a:r>"#)
}

fn open(paragraphs: &str, list: &str) -> DeckSession {
    let shape = support::placeholder(2, "body", paragraphs, false)
        .replace(
            "<p:spPr/>",
            r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="8000000" cy="4000000"/></a:xfrm></p:spPr>"#,
        )
        .replace("<a:lstStyle/>", &format!("<a:lstStyle>{list}</a:lstStyle>"));
    let bytes = support::deck(&[(&shape, false)], "", "", "", Some(12));
    DeckSession::open(&bytes, 842).unwrap()
}

fn story_id(session: &DeckSession) -> String {
    session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone()
}

fn rendered(session: &DeckSession) -> Vec<TextParagraph> {
    fn paragraphs(primitive: Primitive) -> Vec<TextParagraph> {
        match primitive {
            Primitive::TextBox { paragraphs, .. } => paragraphs,
            Primitive::Table { primitives, .. } => {
                primitives.into_iter().flat_map(paragraphs).collect()
            }
            _ => Vec::new(),
        }
    }
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .flat_map(paragraphs)
        .collect()
}

fn language(start: u32, end: u32, lang: &str) -> TextLanguage {
    TextLanguage {
        start,
        end,
        lang: lang.to_owned(),
    }
}

#[test]
fn coalesced_formatting_keeps_source_language_boundaries_and_merges_equal_neighbours() {
    let paragraphs = format!(
        "<a:p>{}{}{}</a:p>",
        run("Hello ", "en-US"),
        run("world ", "en-US"),
        run("bonjour", "fr-FR"),
    );
    let session = open(&paragraphs, "");
    let story = session.story(&story_id(&session)).unwrap();
    assert_eq!(story.paragraphs[0].runs.len(), 1);
    let paragraph = rendered(&session).remove(0);
    assert_eq!(paragraph.runs.len(), 1);
    assert_eq!(paragraph.runs[0].text, "Hello world bonjour");
    assert_eq!(
        paragraph.languages,
        [language(0, 12, "en-US"), language(12, 19, "fr-FR")]
    );
}

#[test]
fn offsets_count_surrogate_pairs_as_two_utf16_units() {
    let paragraphs = format!(
        "<a:p>{}{}</a:p>",
        run("A😀 ", "en-US"),
        run("salut", "fr-FR"),
    );
    let paragraph = rendered(&open(&paragraphs, "")).remove(0);
    assert_eq!(
        paragraph.languages,
        [language(0, 4, "en-US"), language(4, 9, "fr-FR")]
    );
}

#[test]
fn unknown_languages_are_omitted_and_old_json_defaults_to_no_languages() {
    let paragraphs = format!(
        "<a:p>{}{}{}</a:p>",
        run("A", "x-none"),
        run("B", ""),
        run("C", "X-NONE"),
    );
    let paragraph = rendered(&open(&paragraphs, "")).remove(0);
    assert!(paragraph.languages.is_empty());
    let json = serde_json::to_value(&paragraph).unwrap();
    assert!(json.get("languages").is_none());
    let decoded: TextParagraph = serde_json::from_value(json).unwrap();
    assert!(decoded.languages.is_empty());
}

#[test]
fn run_language_precedes_the_paragraph_end_then_list_default() {
    let paragraphs = format!(
        r#"<a:p>{}{}<a:r><a:t>world</a:t></a:r><a:endParaRPr lang="en-US"/></a:p><a:p>{}<a:endParaRPr lang="x-none"/></a:p>"#,
        run("salut ", "fr-FR"),
        run("hello ", "x-none"),
        run("Guten Tag", ""),
    );
    let list = r#"<a:lvl1pPr><a:defRPr lang="de-DE"/></a:lvl1pPr>"#;
    let paragraphs = rendered(&open(&paragraphs, list));
    assert_eq!(
        paragraphs[0].languages,
        [language(0, 6, "fr-FR"), language(6, 17, "en-US")]
    );
    assert_eq!(paragraphs[1].languages, [language(0, 9, "de-DE")]);
}

#[test]
fn edited_text_uses_one_source_language_without_guessing_mixed_ranges() {
    for (second, expected) in [("en-US", vec![language(0, 15, "en-US")]), ("fr-FR", vec![])] {
        let paragraphs = format!(
            "<a:p>{}{}</a:p>",
            run("Hello ", "en-US"),
            run("world", second),
        );
        let session = open(&paragraphs, "");
        session
            .insert_text(
                &EditCtx::local("test"),
                &story_id(&session),
                0,
                "New ",
                &TextStyle::default(),
            )
            .unwrap();
        let paragraph = rendered(&session).remove(0);
        assert_eq!(
            paragraph
                .runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>(),
            "New Hello world"
        );
        assert_eq!(paragraph.languages, expected);
    }
}

#[test]
fn edited_text_with_unmarked_runs_is_not_given_a_tagged_runs_language() {
    let paragraphs = format!(
        "<a:p><a:r><a:t>Please say </a:t></a:r>{}</a:p>",
        run("bonjour", "fr-FR"),
    );
    let list = r#"<a:lvl1pPr><a:defRPr lang="en-US"/></a:lvl1pPr>"#;
    let session = open(&paragraphs, list);
    session
        .insert_text(
            &EditCtx::local("test"),
            &story_id(&session),
            0,
            "New ",
            &TextStyle::default(),
        )
        .unwrap();
    assert!(rendered(&session).remove(0).languages.is_empty());
}

#[test]
fn joined_paragraphs_require_one_language_across_all_sources() {
    for (second, expected) in [("en-US", vec![language(0, 10, "en-US")]), ("fr-FR", vec![])] {
        let paragraphs = format!(
            "<a:p>{}</a:p><a:p>{}</a:p>",
            run("Hello", "en-US"),
            run("world", second),
        );
        let session = open(&paragraphs, "");
        session
            .delete_text_joining_paragraphs(&EditCtx::local("test"), &story_id(&session), 5, 6)
            .unwrap();
        let paragraphs = rendered(&session);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].languages, expected);
    }
}

#[test]
fn new_paragraphs_do_not_borrow_languages_from_neighbours() {
    let session = open(&format!("<a:p>{}</a:p>", run("Hello", "en-US")), "");
    let story = story_id(&session);
    session
        .insert_paragraph_break(&EditCtx::local("test"), &story, 0)
        .unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "New",
            &TextStyle::default(),
        )
        .unwrap();
    let paragraphs = rendered(&session);
    assert_eq!(paragraphs.len(), 2);
    assert!(paragraphs[0].languages.is_empty());
    assert_eq!(paragraphs[1].languages, [language(0, 5, "en-US")]);
}

#[test]
fn slide_number_fields_keep_their_language_and_shift_later_offsets() {
    for cached in ["9", ""] {
        let paragraphs = format!(
            r#"<a:p>{}<a:fld id="{{00000000-0000-0000-0000-000000000001}}" type="slidenum"><a:rPr lang="de-DE"/><a:t>{cached}</a:t></a:fld>{}</a:p>"#,
            run("Page ", "en-US"),
            run(" suite", "fr-FR"),
        );
        let paragraph = rendered(&open(&paragraphs, "")).remove(0);
        assert_eq!(
            paragraph
                .runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>(),
            "Page 12 suite"
        );
        assert_eq!(
            paragraph.languages,
            [
                language(0, 5, "en-US"),
                language(5, 7, "de-DE"),
                language(7, 13, "fr-FR")
            ]
        );
    }
}

#[test]
fn text_typed_into_an_empty_source_uses_its_end_language() {
    let session = open(r#"<a:p><a:endParaRPr lang="fr-FR"/></a:p>"#, "");
    session
        .insert_text(
            &EditCtx::local("test"),
            &story_id(&session),
            0,
            "Bonjour",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(rendered(&session)[0].languages, [language(0, 7, "fr-FR")]);
}

#[test]
fn table_cells_keep_source_languages() {
    let table = format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="0" y="0"/><a:ext cx="8000000" cy="1200000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="8000000"/></a:tblGrid>
<a:tr h="1200000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>{}{}</a:p></a:txBody><a:tcPr/></a:tc></a:tr>
</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        run("Hello ", "en-US"),
        run("salut", "fr-FR"),
    );
    let bytes = support::deck(&[(&table, false)], "", "", "", None);
    let session = DeckSession::open(&bytes, 843).unwrap();
    let paragraphs = rendered(&session);
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(
        paragraphs[0].languages,
        [language(0, 6, "en-US"), language(6, 11, "fr-FR")]
    );
}
