//! Keystrokes that only move the note references after them keep those pages' display.
use super::*;
use yrs::{ReadTxn, Transact};

const PARAGRAPHS: usize = 300;

/// Paragraph `index` of the fixture holds a note reference unless it is a multiple of five;
/// the references number the notes from 1 in order.
fn note_of(index: usize) -> Option<usize> {
    (index % 5 != 0).then(|| index - index / 5)
}

const NOTES: usize = PARAGRAPHS - PARAGRAPHS / 5;

/// [`PARAGRAPHS`] paragraphs on small pages, [`NOTES`] of them referencing a footnote of
/// their own, so most pages show notes anchored after an edit near the start.
fn notes_docx() -> Vec<u8> {
    use crate::seed::fixture::{NS, Package, para, run};
    let spacing = r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/><w:widowControl w:val="0"/></w:pPr>"#;
    let body: String = (0..PARAGRAPHS)
        .map(|index| {
            let mut content = format!("{spacing}{}", run(&format!("Paragraph {index} words")));
            if let Some(note) = note_of(index) {
                content.push_str(&format!(
                    r#"<w:r><w:footnoteReference w:id="{note}"/></w:r>{}"#,
                    run(" and then some")
                ));
            }
            para(&format!("{:08X}", 0x5400_0000 + index), &content)
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
                &format!("{:08X}", 0x5500_0000 + note),
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

/// Opens the fixture as the worker does, its display limited to `window` when given.
fn notes_engine(window: Option<std::ops::Range<usize>>) -> (EngineSession, String) {
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let engine = EngineSession::new(95114);
    crate::seed_from_docx(engine.doc(), &notes_docx()).unwrap();
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
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 360, "h": 480},
            "margins": {"top": 24, "right": 24, "bottom": 24, "left": 24}
        }]},
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
    let request = request.to_string();
    engine.set_local_lowering(true);
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    if let Some(window) = window {
        engine.set_display_window(Some(window));
        engine.set_windowed_incremental_builds(true);
    }
    engine
        .build_display_list_frame(&json!({"fontChains": chains}).to_string(), 0)
        .unwrap();
    let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
    assert!(pages > 20, "{pages} pages");
    (engine, request)
}

/// The start of body paragraph `index`.
fn paragraph_start(engine: &EngineSession, index: usize) -> u32 {
    let paragraphs = engine.doc().paragraphs("body").unwrap();
    engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_content_start(&paragraphs[index].para_id)
        .unwrap()
}

/// The position of the note reference in body paragraph `index`.
fn reference_in(engine: &EngineSession, index: usize) -> u32 {
    let start = paragraph_start(engine, index);
    (start..start + 120)
        .find(|&at| {
            engine
                .doc()
                .segment_index("body")
                .unwrap()
                .segment_at(at)
                .is_some_and(|segment| matches!(segment.kind, crate::segments::SegKind::Embed))
        })
        .unwrap()
}

/// What the frames sent so far leave the client holding for each page's note backlinks
/// equals what the display list holds: a page shifted as a uniform range would keep its
/// old anchors there.
fn assert_frames_carry_note_anchors(engine: &EngineSession, label: &str) {
    let display = engine.display.borrow();
    let list = display.list.as_ref().unwrap();
    assert_eq!(display.pages.len(), list.pages.len(), "{label}");
    for (index, (snapshot, page)) in display.pages.iter().zip(&list.pages).enumerate() {
        let shown: Vec<_> = page
            .note_areas
            .iter()
            .flat_map(|area| &area.notes)
            .map(|note| (note.anchor_doc_start, note.anchor_doc_end))
            .collect();
        let sent: Vec<_> = snapshot
            .note_anchors
            .iter()
            .map(|anchor| (anchor.start, anchor.end))
            .collect();
        assert_eq!(sent, shown, "{label}: page {index}");
    }
}

/// Every page the windowed display has built equals that page of a cold display build
/// over the retained layout.
fn assert_built_pages_match_cold(engine: &EngineSession, extras: &str, label: &str) {
    let cold = {
        let pagination = engine.pagination.borrow();
        docx_layout::build_resident_display_list_observed(
            pagination.input.as_ref().unwrap(),
            pagination.layout.as_ref().unwrap(),
            extras,
            &mut || {},
        )
        .unwrap()
        .1
    };
    let display = engine.display.borrow();
    let list = display.list.as_ref().unwrap();
    assert_eq!(list.pages.len(), cold.pages.len(), "{label}");
    let mut built = 0;
    for (index, (page, cold)) in list.pages.iter().zip(&cold.pages).enumerate() {
        if !page.unbuilt {
            assert_eq!(page, cold, "{label}: page {index}");
            built += 1;
        }
    }
    assert!(built > 0, "{label}");
}

/// Edits `story` as a keystroke does, lays out and builds the display as the worker does
/// after one, and checks the result against a cold layout and display. Returns the display
/// pages the key rebuilt and the pages whose note anchors moved in place.
fn keystroke(
    engine: &EngineSession,
    request: &str,
    story: &str,
    range: std::ops::Range<u32>,
    text: Option<&str>,
    label: &str,
) -> (u64, usize) {
    engine
        .edit_resident_text(
            crate::StoryRange::new(story, range.start, range.end),
            text,
            true,
        )
        .unwrap();
    let before = engine.stats().rebuilt_display_pages;
    let epoch = engine.display.borrow().binary_frame_epoch;
    engine.set_relayout_trigger(RelayoutTrigger::Interactive);
    if !engine
        .apply_and_layout_regions_resident(story, &mut |_| {}, RelayoutTrigger::Interactive)
        .unwrap()
    {
        engine
            .apply_and_layout_regions_full(RelayoutTrigger::Interactive)
            .unwrap();
    }
    let moved = engine.pagination.borrow().note_anchor_pages.len();
    let extras = engine.resident_region_display_extras().unwrap();
    engine.build_display_list_frame(&extras, epoch).unwrap();
    let rebuilt = engine.stats().rebuilt_display_pages - before;
    assert_frames_carry_note_anchors(engine, label);
    if engine.display.borrow().window.is_some() {
        assert_built_pages_match_cold(engine, &extras, label);
    } else {
        super::assert_region_state_matches_cold(engine, request, label);
    }
    (rebuilt, moved)
}

#[test]
fn body_keys_before_note_references_rebuild_only_the_edited_page() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let (engine, request) = notes_engine(None);
    let pages = engine.with_display_list(|list| list.pages.len()).unwrap();
    let free = 10;
    assert_eq!(note_of(free), None);
    for (index, text) in [Some("x"), None, Some("y"), None].into_iter().enumerate() {
        let at = paragraph_start(&engine, free) + 3;
        let (rebuilt, moved) = keystroke(
            &engine,
            &request,
            "body",
            at..at + text.is_none() as u32,
            text,
            &format!("key {index} early in the body"),
        );
        assert!(
            rebuilt <= 2,
            "key {index}: rebuilt {rebuilt} of {pages} pages"
        );
        assert!(
            moved > pages / 2,
            "key {index}: anchors moved on {moved} of {pages} pages"
        );
    }
}

#[test]
fn keys_beside_and_inside_notes_match_a_cold_display() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for window in [None, Some(2..5)] {
        let (engine, request) = notes_engine(window.clone());
        let paragraph = 12;
        let note = note_of(paragraph).unwrap();
        let reference = reference_in(&engine, paragraph);
        let keys = [
            (
                "body".to_owned(),
                reference..reference,
                Some("b"),
                "before a reference",
            ),
            (
                "body".to_owned(),
                reference + 2..reference + 2,
                Some("a"),
                "after a reference",
            ),
            (
                "body".to_owned(),
                reference + 2..reference + 3,
                None,
                "deleting after a reference",
            ),
            (
                "body".to_owned(),
                reference..reference + 1,
                None,
                "deleting before a reference",
            ),
            (format!("fn:{note}"), 2..2, Some("n"), "in a note"),
            (
                format!("fn:{}", note + 40),
                2..3,
                None,
                "deleting in a later note",
            ),
        ];
        for (story, range, text, label) in keys {
            keystroke(
                &engine,
                &request,
                &story,
                range,
                text,
                &format!("{window:?}: {label}"),
            );
        }
    }
}

#[test]
fn deleting_a_note_reference_renumbers_later_notes_as_a_cold_display_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for window in [None, Some(3..5)] {
        let (engine, request) = notes_engine(window.clone());
        let reference = reference_in(&engine, 21);
        keystroke(
            &engine,
            &request,
            "body",
            reference..reference + 1,
            None,
            &format!("{window:?}: deleting a note reference"),
        );
        let at = paragraph_start(&engine, 6) + 2;
        let (rebuilt, moved) = keystroke(
            &engine,
            &request,
            "body",
            at..at,
            Some("x"),
            &format!("{window:?}: a key after renumbering"),
        );
        assert!(rebuilt <= 2, "{window:?}: rebuilt {rebuilt}");
        assert!(moved > 0, "{window:?}");
    }
}
