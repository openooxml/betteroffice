//! Note passes of keystroke relayouts resume from the passes of the relayout before, and
//! every keystroke still lays out as a cold layout does.

use super::*;
use yrs::{ReadTxn, Transact};

/// A layout request over `sections` and the document's footnotes, every font it asks for
/// measured with `font`.
fn note_request(engine: &EngineSession, font: u32, sections: serde_json::Value) -> String {
    let contents: Vec<serde_json::Value> = {
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
    request.to_string()
}

fn sections(restart: Option<&str>) -> serde_json::Value {
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

fn open_engine(bytes: &[u8], restart: Option<&str>) -> (EngineSession, String) {
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let engine = EngineSession::new(95112);
    crate::seed_from_docx(engine.doc(), bytes).unwrap();
    let request = note_request(&engine, font, sections(restart));
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

/// A body of paragraphs; paragraph `index` holds `references(index)` footnote references
/// after its words, numbering the notes from 1 in order, and note `note` reads `note_text`.
fn notes_docx(
    paragraphs: usize,
    words: impl Fn(usize) -> String,
    references: impl Fn(usize) -> usize,
    note_text: impl Fn(usize) -> String,
    exact_lines: bool,
) -> Vec<u8> {
    use crate::seed::fixture::{NS, Package, para, run};
    let spacing = if exact_lines {
        r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/><w:widowControl w:val="0"/></w:pPr>"#
    } else {
        r#"<w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr>"#
    };
    let mut note = 0;
    let body: String = (0..paragraphs)
        .map(|index| {
            let mut content = format!("{spacing}{}", run(&words(index)));
            for _ in 0..references(index) {
                note += 1;
                content.push_str(&format!(
                    r#"<w:r><w:footnoteReference w:id="{note}"/></w:r>{}"#,
                    run(" and")
                ));
            }
            para(&format!("{:08X}", 0x5300_0000 + index), &content)
        })
        .collect();
    let mut notes = format!(
        r#"<w:footnotes {NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#
    );
    for id in 1..=note {
        notes.push_str(&format!(
            r#"<w:footnote w:id="{id}">{}</w:footnote>"#,
            para(
                &format!("{:08X}", 0x5400_0000 + id),
                &format!(
                    r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/></w:pPr>{}"#,
                    run(&note_text(id))
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

const PARAGRAPHS: usize = 200;

/// Paragraph `index` of the regular fixture references a note unless it is a multiple of five.
fn regular_references(index: usize) -> usize {
    usize::from(index % 5 != 0)
}

fn regular_docx() -> Vec<u8> {
    notes_docx(
        PARAGRAPHS,
        |index| format!("Paragraph {index} has words that fill a line or two"),
        regular_references,
        |note| {
            let words = if note % 3 == 0 {
                " that runs on past the end of its first line"
            } else {
                ""
            };
            format!("Note {note}{words}")
        },
        true,
    )
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

/// Each page's reserved note space and the position its body opens at.
fn page_reservations(engine: &EngineSession) -> Vec<(Option<f64>, Option<f64>)> {
    let pagination = engine.pagination.borrow();
    pagination
        .layout
        .as_ref()
        .unwrap()
        .pages
        .iter()
        .map(|page| {
            let opening = page.fragments.first().and_then(|fragment| match fragment {
                Fragment::Paragraph(fragment) => fragment.pm_start,
                _ => None,
            });
            (page.footnote_reserved_height, opening)
        })
        .collect()
}

fn notes_converged(engine: &EngineSession) -> bool {
    engine.regions.borrow().as_ref().unwrap().notes_converged
}

/// Edits `story` as a keystroke does, lays the document out again as the worker does after
/// one, and checks the result against a cold layout. Returns the note passes placed in full
/// and resumed.
fn keystroke(
    engine: &EngineSession,
    request: &str,
    story: &str,
    range: std::ops::Range<u32>,
    text: Option<&str>,
    label: &str,
) -> (usize, usize) {
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
    NOTE_PASS_PLACEMENTS.with(|counts| counts.set((0, 0)));
    let epoch = engine.display.borrow().binary_frame_epoch;
    engine.apply_and_layout(story, epoch).unwrap();
    let placements = NOTE_PASS_PLACEMENTS.with(Cell::get);
    super::assert_region_state_matches_cold(engine, request, label);
    placements
}

/// A fixture whose note fixpoint settles: every fifth paragraph references a note.
fn settling_docx() -> Vec<u8> {
    notes_docx(
        PARAGRAPHS,
        |index| format!("Paragraph {index} has words that fill a line or two"),
        |index| usize::from(index % 5 == 1),
        |note| format!("Note {note}"),
        true,
    )
}

#[test]
fn keys_beside_notes_resume_every_note_pass_and_match_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    // The regular fixture's notes only settle through the covering fallback.
    for (bytes, settles) in [(settling_docx(), true), (regular_docx(), false)] {
        let (engine, request) = open_engine(&bytes, None);
        assert_eq!(notes_converged(&engine), settles);
        let pages = page_notes(&engine);
        assert!(pages.len() > 10, "{} pages", pages.len());
        let free = PARAGRAPHS / 2;
        assert_eq!(regular_references(free), 0);
        for (index, text) in [Some("x"), None, Some("y")].into_iter().enumerate() {
            let label = format!("settles {settles}: key {index} beside the notes");
            let at = paragraph_start(&engine, free) + 3;
            let (full, resumed) = keystroke(
                &engine,
                &request,
                "body",
                at..at + u32::from(text.is_none()),
                text,
                &label,
            );
            assert_eq!(full, 0, "{label}");
            assert!(resumed >= 2, "{label}: {resumed} passes resumed");
        }
    }
}

#[test]
fn keys_in_a_note_and_a_note_that_moves_page_breaks_match_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let (engine, request) = open_engine(&regular_docx(), None);
    for (index, text) in [Some("x"), None].into_iter().enumerate() {
        keystroke(
            &engine,
            &request,
            "fn:40",
            2..2 + u32::from(text.is_none()),
            text,
            &format!("key {index} in a note"),
        );
    }
    let before = page_reservations(&engine);
    let long = " and many more words".repeat(12);
    keystroke(
        &engine,
        &request,
        "fn:40",
        4..4,
        Some(&long),
        "a note growing by lines",
    );
    let after = page_reservations(&engine);
    let page = page_notes(&engine)
        .iter()
        .position(|notes| notes.iter().any(|(id, _)| *id == 40))
        .unwrap();
    assert_ne!(
        before[page].0, after[page].0,
        "the note's page reserves more"
    );
    assert_ne!(
        before[page + 1].1,
        after[page + 1].1,
        "the next page opens elsewhere"
    );
    // The keystroke after it resumes from the passes that reached the new reservations.
    let at = paragraph_start(&engine, PARAGRAPHS / 2) + 3;
    let (full, _) = keystroke(&engine, &request, "body", at..at, Some("x"), "a key after");
    assert_eq!(full, 0);
    keystroke(
        &engine,
        &request,
        "fn:40",
        4..4 + long.encode_utf16().count() as u32,
        None,
        "the note shrinking back",
    );
    assert_eq!(page_reservations(&engine).len(), before.len());
}

#[test]
fn a_key_that_moves_a_note_reference_to_the_next_page_matches_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for restart in [None, Some("eachPage")] {
        let (engine, request) = open_engine(&regular_docx(), restart);
        let before = page_notes(&engine);
        let page = before.len() / 2;
        let (last, _) = before[page].last().cloned().unwrap();
        let mut note = 0;
        let paragraph = (0..PARAGRAPHS)
            .find(|&index| {
                note += regular_references(index);
                regular_references(index) == 1 && note == last as usize
            })
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
            if !after[page].iter().any(|(id, _)| *id == last) {
                moved = true;
                break;
            }
        }
        assert!(moved, "{restart:?}: the note stayed on page {page}");
        let at = paragraph_start(&engine, paragraph - 1) + 2;
        keystroke(
            &engine,
            &request,
            "body",
            at..at + " growing words".len() as u32,
            None,
            &format!("{restart:?}: deleting the words again"),
        );
    }
}

#[test]
fn keys_beside_notes_crowding_one_page_match_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let crowded = 30;
    let bytes = notes_docx(
        60,
        |index| format!("Paragraph {index} has words that fill a line or two"),
        |index| match index {
            index if index == crowded => 9,
            index if index % 4 == 1 => 1,
            _ => 0,
        },
        |note| format!("Note {note} with words enough to take two or three lines of note text"),
        true,
    );
    let (engine, request) = open_engine(&bytes, None);
    let pages = page_notes(&engine);
    assert!(
        pages.iter().any(|notes| notes.len() >= 8),
        "{:?}",
        pages.iter().map(Vec::len).collect::<Vec<_>>()
    );
    for (paragraph, text) in [
        (crowded - 2, Some("x")),
        (crowded, Some("y")),
        (crowded, None),
        (
            crowded - 1,
            Some(" more words to push the crowded paragraph"),
        ),
    ] {
        let at = paragraph_start(&engine, paragraph) + 3;
        let (_, resumed) = keystroke(
            &engine,
            &request,
            "body",
            at..at + u32::from(text.is_none()),
            text,
            &format!("key in paragraph {paragraph}"),
        );
        assert!(resumed > 0, "paragraph {paragraph}");
    }
}

/// Long paragraphs ending in references to notes of a few lines, so reference lines keep
/// landing at page bottoms where reserving their notes pushes them off.
fn unsettled_docx() -> Vec<u8> {
    notes_docx(
        60,
        |index| {
            (0..18 + (index * 37) % 40)
                .map(|word| format!("w{}x{word}", index % 97))
                .collect::<Vec<_>>()
                .join(" ")
        },
        |index| usize::from(index % 2 == 0),
        |note| {
            (0..6 + (note * 53) % 29)
                .map(|word| format!("n{note}x{word}"))
                .collect::<Vec<_>>()
                .join(" ")
        },
        false,
    )
}

#[test]
fn keys_where_notes_never_settle_match_a_cold_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let (engine, request) = open_engine(&unsettled_docx(), None);
    assert!(!notes_converged(&engine), "the fixture's notes settle");
    for (index, (paragraph, text)) in [(31, Some("x")), (31, None), (3, Some("words "))]
        .into_iter()
        .enumerate()
    {
        let at = paragraph_start(&engine, paragraph) + 3;
        keystroke(
            &engine,
            &request,
            "body",
            at..at + u32::from(text.is_none()),
            text,
            &format!("key {index} in an unsettled document"),
        );
        assert!(!notes_converged(&engine), "key {index}");
    }
}
