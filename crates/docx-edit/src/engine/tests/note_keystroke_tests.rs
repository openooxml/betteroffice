use super::*;
use yrs::{ReadTxn, Transact};

/// A layout request over `sections` and `contents`, every font the document asks for measured
/// with `font`.
fn note_request(
    engine: &EngineSession,
    font: u32,
    sections: serde_json::Value,
    contents: Vec<serde_json::Value>,
) -> serde_json::Value {
    let mut request = json!({
        "bodyStory": "body",
        "renderEnv": {},
        "options": {"pageGap": 24},
        "regions": {"sections": sections},
        "notes": {"contents": contents},
    });
    let requirements: Vec<serde_json::Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let chains: serde_json::Map<String, serde_json::Value> = requirements
        .iter()
        .map(|requirement| {
            (
                requirement["key"].as_str().unwrap().to_owned(),
                json!([font]),
            )
        })
        .collect();
    request["measurement"] = json!({
        "fontChains": chains,
        "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        "authoritativeShaping": true,
    });
    request
}

fn open_engine(bytes: &[u8], font: u32, sections: serde_json::Value) -> (EngineSession, String) {
    let engine = EngineSession::new(95104);
    crate::seed_from_docx(engine.doc(), bytes).unwrap();
    let contents = {
        let txn = engine.doc().yrs_doc().transact();
        let stories = txn.get_map(crate::STORIES).unwrap();
        let mut ids: Vec<i64> = crate::live_stories(&stories, &txn)
            .filter_map(|(id, _)| id.strip_prefix("fn:")?.parse().ok())
            .filter(|id| *id > 0)
            .collect();
        ids.sort_unstable();
        ids.into_iter()
            .map(|id| json!({"id": id, "noteKind": "footnote", "height": 0}))
            .collect()
    };
    let request = note_request(&engine, font, sections, contents).to_string();
    engine.set_local_lowering(true);
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    let request_value: serde_json::Value = serde_json::from_str(&request).unwrap();
    engine
        .build_display_list_frame(
            &json!({"fontChains": request_value["measurement"]["fontChains"]}).to_string(),
            0,
        )
        .unwrap();
    (engine, request)
}

/// The start of paragraph `index` of the body.
fn paragraph_start(engine: &EngineSession, index: usize) -> u32 {
    let paragraphs = engine.doc().paragraphs("body").unwrap();
    engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_content_start(&paragraphs[index].para_id)
        .unwrap()
}

const NOTES: usize = 400;
const PARAGRAPHS: usize = 500;

/// Paragraph `index` of the note fixture holds a note reference unless it is a multiple of
/// five; the references number the notes from 1 in order.
fn note_of(index: usize) -> Option<usize> {
    (index % 5 != 0).then(|| index - index / 5)
}

/// A package of [`PARAGRAPHS`] paragraphs, [`NOTES`] of them referencing a footnote of their
/// own.
fn notes_docx() -> Vec<u8> {
    use crate::seed::fixture::{NS, Package, para, run};
    let spacing = r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/><w:widowControl w:val="0"/></w:pPr>"#;
    let body: String = (0..PARAGRAPHS)
        .map(|index| {
            let mut content = format!(
                "{spacing}{}",
                run(&format!(
                    "Paragraph {index} has words that fill a line or two"
                ))
            );
            if let Some(note) = note_of(index) {
                content.push_str(&format!(
                    r#"<w:r><w:footnoteReference w:id="{note}"/></w:r>{}"#,
                    run(" and then some")
                ));
            }
            para(&format!("{:08X}", 0x5100_0000 + index), &content)
        })
        .collect();
    let mut notes = format!(
        r#"<w:footnotes {NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#
    );
    for note in 1..=NOTES {
        let words = if note % 3 == 0 {
            " that runs on past the end of its first line"
        } else {
            ""
        };
        notes.push_str(&format!(
            r#"<w:footnote w:id="{note}">{}</w:footnote>"#,
            para(
                &format!("{:08X}", 0x5200_0000 + note),
                &format!(
                    r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/></w:pPr>{}"#,
                    run(&format!("Note {note}{words}"))
                )
            )
        ));
    }
    notes.push_str("</w:footnotes>");
    Package::new(&body)
        .part(
            "footnotes.xml",
            "rIdFootnotes",
            "footnotes",
            "footnotes",
            &notes,
        )
        .bytes()
}

fn note_sections(restart: Option<&str>) -> serde_json::Value {
    let mut section = json!({
        "sectionId": "main",
        "pageSize": {"w": 360, "h": 480},
        "margins": {"top": 24, "right": 24, "bottom": 24, "left": 24}
    });
    if let Some(restart) = restart {
        section["noteSettings"] = json!({"footnote": {"numRestart": restart}});
    }
    json!([section])
}

/// The notes each page of the retained layout shows, with their labels.
fn page_notes(engine: &EngineSession) -> Vec<Vec<(i64, String)>> {
    let pagination = engine.pagination.borrow();
    pagination
        .layout
        .as_ref()
        .unwrap()
        .pages
        .iter()
        .map(|page| {
            page.note_areas
                .iter()
                .flatten()
                .flat_map(|area| area.notes.iter().flatten())
                .map(|note| (note.id.unwrap(), note.display_label.clone().unwrap()))
                .collect()
        })
        .collect()
}

/// Edits `story` as a keystroke does, lays the document out again as the worker does after
/// one, and checks the result against a cold layout. Returns the note stories measured afresh.
fn keystroke(
    engine: &EngineSession,
    request: &str,
    story: &str,
    range: std::ops::Range<u32>,
    text: Option<&str>,
    label: &str,
) -> usize {
    engine
        .edit_resident_text(
            crate::StoryRange::new(story, range.start, range.end),
            text,
            true,
        )
        .unwrap();
    assert_eq!(
        engine.region_relayout_trigger(request).unwrap(),
        RelayoutTrigger::Interactive,
        "{label}"
    );
    NOTE_STORY_MEASUREMENTS.with(|count| count.set(0));
    let epoch = engine.display.borrow().binary_frame_epoch;
    engine.apply_and_layout(story, epoch).unwrap();
    let measured = NOTE_STORY_MEASUREMENTS.with(Cell::get);
    super::assert_region_state_matches_cold(engine, request, label);
    measured
}

fn notes_engine(restart: Option<&str>) -> (EngineSession, String) {
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let (engine, request) = open_engine(&notes_docx(), font, note_sections(restart));
    let pages = page_notes(&engine);
    assert!(pages.len() > 40, "{} pages", pages.len());
    assert_eq!(pages.iter().map(Vec::len).sum::<usize>(), NOTES);
    (engine, request)
}

/// The paragraph whose note `note` references, and the page showing it.
fn note_page(pages: &[Vec<(i64, String)>], note: usize) -> usize {
    pages
        .iter()
        .position(|notes| notes.iter().any(|(id, _)| *id == note as i64))
        .unwrap()
}

#[test]
fn keys_beside_notes_reuse_note_measurements_and_match_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let (engine, request) = notes_engine(None);
    let free = PARAGRAPHS / 2;
    assert_eq!(note_of(free), None);
    for (index, text) in [Some("x"), None, Some("y")].into_iter().enumerate() {
        let at = paragraph_start(&engine, free) + 3;
        let measured = keystroke(
            &engine,
            &request,
            "body",
            at..at + text.is_none() as u32,
            text,
            &format!("key {index} beside the notes"),
        );
        assert_eq!(measured, 0, "key {index} beside the notes");
    }
}

#[test]
fn a_key_in_a_note_remeasures_that_note_alone_and_matches_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let (engine, request) = notes_engine(None);
    let before = page_notes(&engine);
    for (index, text) in [Some("x"), Some("y"), None].into_iter().enumerate() {
        let measured = keystroke(
            &engine,
            &request,
            "fn:123",
            2..2 + text.is_none() as u32,
            text,
            &format!("key {index} in a note"),
        );
        assert_eq!(measured, 1, "key {index} in a note");
    }
    assert_eq!(page_notes(&engine), before);
    let long = " and many more words".repeat(12);
    let measured = keystroke(
        &engine,
        &request,
        "fn:123",
        4..4,
        Some(&long),
        "a note growing by lines",
    );
    assert_eq!(measured, 1);
}

#[test]
fn a_key_that_moves_a_note_reference_to_the_next_page_matches_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for restart in [None, Some("eachPage")] {
        let (engine, request) = notes_engine(restart);
        let before = page_notes(&engine);
        // The last note of a page some way in, and a paragraph before it on that page.
        let page = before.len() / 2;
        let (last, _) = before[page].last().cloned().unwrap();
        let paragraph = (0..PARAGRAPHS)
            .find(|&index| note_of(index) == Some(last as usize))
            .unwrap();
        let mut moved = false;
        for index in 0..8 {
            let at = paragraph_start(&engine, paragraph - 1) + 2;
            keystroke(
                &engine,
                &request,
                "body",
                at..at,
                Some(" growing words"),
                &format!("{restart:?}: key {index} before a page's last note"),
            );
            let after = page_notes(&engine);
            if note_page(&after, last as usize) != page {
                if restart.is_some() {
                    assert_ne!(after[page + 1], before[page + 1], "{restart:?}");
                }
                moved = true;
                break;
            }
        }
        assert!(moved, "{restart:?}: the note stayed on page {page}");
    }
}

#[test]
fn a_key_that_deletes_a_note_reference_renumbers_the_notes_after_it_as_a_cold_layout_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let (engine, request) = notes_engine(None);
    let paragraph = 101;
    let note = note_of(paragraph).unwrap();
    let before = page_notes(&engine);
    let start = paragraph_start(&engine, paragraph);
    let reference = (start..start + 80)
        .find(|&at| {
            engine
                .doc()
                .segment_index("body")
                .unwrap()
                .segment_at(at)
                .is_some_and(|segment| matches!(segment.kind, crate::segments::SegKind::Embed))
        })
        .unwrap();
    keystroke(
        &engine,
        &request,
        "body",
        reference..reference + 1,
        None,
        "deleting a note reference",
    );
    let after = page_notes(&engine);
    assert_eq!(after.iter().map(Vec::len).sum::<usize>(), NOTES - 1);
    let label = |pages: &[Vec<(i64, String)>]| {
        pages[note_page(pages, note + 1)]
            .iter()
            .find(|(id, _)| *id == note as i64 + 1)
            .unwrap()
            .1
            .clone()
    };
    assert_eq!(label(&before), (note + 1).to_string());
    assert_eq!(label(&after), note.to_string());
    let at = paragraph_start(&engine, PARAGRAPHS / 2) + 3;
    let measured = keystroke(&engine, &request, "body", at..at, Some("x"), "a key after");
    assert_eq!(measured, 0);
}
