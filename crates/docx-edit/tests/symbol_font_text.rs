#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditingDoc, seed_from_docx};
use docx_layout::types::{LayoutBlock, Run};

fn body() -> String {
    let symbol = r#"<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>"#;
    let east_asia_only =
        r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Symbol"/></w:rPr>"#;
    [
        fixture::p(
            "00000001",
            &format!(r#"<w:r>{symbol}<w:t xml:space="preserve">a+b &#xF0B3; p</w:t></w:r>"#),
        ),
        fixture::p(
            "00000002",
            r#"<w:r><w:sym w:font="Symbol" w:char="F061"/><w:sym w:font="Wingdings" w:char="F0E0"/><w:sym w:font="Wingdings" w:char="F0D8"/><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>"#,
        ),
        fixture::p(
            "00000003",
            &format!(r#"<w:r>{east_asia_only}<w:t>ab</w:t></w:r>"#),
        ),
    ]
    .concat()
}

fn layout_texts(doc: &EditingDoc) -> Vec<String> {
    yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default())
        .unwrap()
        .iter()
        .filter_map(|block| match block {
            LayoutBlock::Paragraph(paragraph) => Some(
                paragraph
                    .runs
                    .iter()
                    .filter_map(|run| match run {
                        Run::Text(text) => Some(text.text.as_str()),
                        _ => None,
                    })
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect()
}

#[test]
fn symbol_face_slots_draw_their_covered_characters_and_keep_the_stored_text() {
    let bytes = fixture::with_body(&body());
    let doc = EditingDoc::new(5100);
    seed_from_docx(&doc, &bytes).unwrap();
    let stored: Vec<String> = doc
        .paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect();
    let plan = doc.paragraph_save_plan();

    let drawn = layout_texts(&doc);
    assert_eq!(drawn, ["α+β ≥ π", "α\u{f0e0}►\u{f0fc}", "ab"]);
    for (drawn, stored) in drawn.iter().zip(&stored) {
        assert_eq!(drawn.encode_utf16().count(), stored.encode_utf16().count());
    }

    assert_eq!(
        stored,
        ["a+b \u{f0b3} p", "\u{f061}\u{f0e0}\u{f0d8}\u{f0fc}", "ab"],
        "the story keeps the stored characters"
    );
    assert_eq!(
        doc.paragraph_save_plan(),
        plan,
        "laying out changes nothing to save"
    );
}

fn first_run(body: &str) -> (String, [Option<String>; 3]) {
    let bytes = fixture::with_body(body);
    let doc = EditingDoc::new(5101);
    seed_from_docx(&doc, &bytes).unwrap();
    let stored = |doc: &EditingDoc| {
        doc.story_segments("body")
            .unwrap()
            .into_iter()
            .map(|segment| segment.attributes)
            .collect::<Vec<_>>()
    };
    let before = stored(&doc);
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    assert_eq!(stored(&doc), before, "the stored run fonts are untouched");
    assert!(
        format!("{before:?}").contains("Calibri"),
        "the stored hAnsi font stays Calibri: {before:?}"
    );
    let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
        panic!("paragraph");
    };
    let Run::Text(run) = &paragraph.runs[0] else {
        panic!("text run");
    };
    let slots = run.fmt.font_slots.clone().unwrap_or_default();
    (
        run.text.clone(),
        [slots.ascii, slots.h_ansi, slots.east_asia],
    )
}

#[test]
fn a_replacement_in_another_slot_still_draws_with_the_symbol_face() {
    let run = |text: &str| {
        first_run(&fixture::p(
            "00000004",
            &format!(
                r#"<w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Calibri"/></w:rPr><w:t>{text}</w:t></w:r>"#
            ),
        ))
    };
    let name = |value: &str| Some(value.to_owned());
    assert_eq!(
        run("a+b"),
        (
            "α+β".to_owned(),
            [name("Symbol"), name("Symbol"), name("Calibri")]
        )
    );
    assert_eq!(
        run("a\u{301}"),
        (
            "α\u{301}".to_owned(),
            [name("Symbol"), name("Symbol"), name("Calibri")]
        ),
        "a combining mark keeps its base's family"
    );
    assert_eq!(
        run("a日"),
        (
            "α日".to_owned(),
            [name("Symbol"), name("Symbol"), name("Calibri")]
        ),
        "an unnamed slot is pinned to the family it fell back to"
    );
    let (text, slots) = run("a\u{e9}");
    assert_eq!(text, "α\u{e9}");
    assert_eq!(
        slots[1],
        name("Calibri"),
        "a character that keeps the slot's own family leaves the slots alone"
    );
}
