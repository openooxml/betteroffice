#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{
    EditCtx, EngineSession, FormatPolicy, Position, SegmentContent, StoryRange, seed_from_docx,
};
use serde_json::{Value, json};

const EDIT_ID: &str = "71001001";
const REF_ID: &str = "71001002";
const FOLLOWER_ID: &str = "71001004";
const NOTE_PARA_ID: &str = "71003001";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum EditClass {
    Interactive,
    Bulk,
}

struct Snapshot {
    output: Value,
    inputs: Value,
}

fn paragraph(id: &str, line_height: u32, properties: &str, content: &str) -> String {
    fixture::p(
        id,
        &format!(
            r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{}" w:lineRule="exact"/><w:widowControl w:val="0"/><w:keepLines/>{properties}</w:pPr>{content}"#,
            line_height * 15,
        ),
    )
}

fn reference(id: u32, text: &str) -> String {
    format!(
        r#"{}<w:r><w:footnoteReference w:id="{id}"/></w:r>"#,
        fixture::r(text),
    )
}

fn prefix() -> String {
    (1..=10)
        .map(|id| {
            paragraph(
                &format!("{:08X}", 0x7100_0000 + id),
                20,
                "",
                &fixture::r("Prefix"),
            )
        })
        .collect()
}

fn tail() -> String {
    (0..5)
        .map(|id| {
            paragraph(
                &format!("{:08X}", 0x7100_0200 + id),
                20,
                if id == 0 { "<w:pageBreakBefore/>" } else { "" },
                &fixture::r("Tail"),
            )
        })
        .collect()
}

fn added_words() -> String {
    " word".repeat(14)
}

fn moving_body(long: bool, section_properties: &str) -> String {
    let text = if long {
        format!("Grow{}", added_words())
    } else {
        "Grow".to_owned()
    };
    format!(
        "{}{}{}{}",
        prefix(),
        paragraph("71001000", 40, "", &fixture::r("Filler")),
        paragraph(EDIT_ID, 20, "", &fixture::r(&text)),
        paragraph(REF_ID, 10, section_properties, &reference(1, "Reference")),
    )
}

fn small_sections() -> Value {
    json!([{
        "sectionId": "main",
        "pageSize": {"w": 200, "h": 120},
        "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10},
    }])
}

fn note() -> String {
    paragraph(NOTE_PARA_ID, 10, "", &fixture::r("Note"))
}

fn snapshot(engine: &EngineSession, output: &str) -> Snapshot {
    let output: Value = serde_json::from_str(output).unwrap();
    assert_eq!(output["notesConverged"], true);
    Snapshot {
        output,
        inputs: serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap(),
    }
}

fn edit_and_compare(
    bytes: &[u8],
    sections: Value,
    class: EditClass,
    edit: impl FnOnce(&EngineSession, EditClass),
) -> (Snapshot, Snapshot) {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let engine = EngineSession::new(75401);
    engine.layout_document_with_regions_retained("{}").unwrap();
    let seed = docx_edit::EditingDoc::new(75400);
    seed_from_docx(&seed, bytes).unwrap();
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    engine
        .doc()
        .set_note_separator_state(seed.note_separator_state().unwrap());
    let mut request = fixture::region_request(&engine, bytes, font);
    request["regions"]["sections"] = sections;
    request["notes"]["contents"] = json!([
        {"id": 1, "noteKind": "footnote", "height": 0},
        {"id": 2, "noteKind": "footnote", "height": 0},
    ]);
    let request = request.to_string();
    let original = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let original = snapshot(&engine, &original);
    let before = engine.stats();
    edit(&engine, class);
    let output = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let after = engine.stats();
    let edited = snapshot(&engine, &output);

    let fresh = EngineSession::new(75402);
    if class == EditClass::Bulk {
        fresh.layout_document_with_regions_retained("{}").unwrap();
    }
    let update = engine.doc().encode_state_as_update_v1();
    match class {
        EditClass::Interactive => fresh.doc().apply_update_v1(&update),
        EditClass::Bulk => fresh.doc().apply_host_update_v1(&update),
    }
    .unwrap();
    fresh
        .doc()
        .set_note_separator_state(engine.doc().note_separator_state().unwrap());
    let expected = fresh
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(output.as_bytes(), expected.as_bytes());
    assert_eq!(engine.retained_layout_json().unwrap(), expected);

    assert_eq!(after.pagination_calls, before.pagination_calls + 1);
    if class == EditClass::Bulk {
        assert_eq!(
            after.incremental_pagination_calls,
            before.incremental_pagination_calls + 1,
            "{class:?}",
        );
        let placed = after.pagination_blocks_placed - before.pagination_blocks_placed;
        let full_placement = fresh.stats().pagination_blocks_placed;
        assert!(
            placed > 0 && placed < full_placement,
            "{placed} of {full_placement}",
        );
        assert_eq!(fresh.stats().incremental_pagination_calls, 0);
    } else {
        assert_eq!(
            after.incremental_pagination_calls, before.incremental_pagination_calls,
            "{class:?}",
        );
        assert_eq!(after.rebuilt_pages, pages(&edited).len());
        assert_eq!(
            after.pagination_blocks_placed - before.pagination_blocks_placed,
            fresh.stats().pagination_blocks_placed,
        );
    }
    (original, edited)
}

fn paragraph_range(engine: &EngineSession, story: &str, para_id: &str) -> StoryRange {
    let mut start = 0;
    let mut end = 0;
    for segment in engine.doc().story_segments(story).unwrap() {
        match segment.content {
            SegmentContent::Text(text) => end += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(properties) => {
                end += 1;
                if properties.para_id == para_id {
                    return StoryRange::new(story, start, end);
                }
                start = end;
            }
            SegmentContent::OtherEmbed { .. } => end += 1,
        }
    }
    panic!("missing paragraph {para_id} in {story}");
}

fn apply_host_edit(engine: &EngineSession, step: Value) {
    let request = serde_json::from_value(json!({
        "expectVersion": engine.doc().version(),
        "history": "none",
        "steps": [step],
    }))
    .unwrap();
    assert!(
        engine
            .doc()
            .apply_edits(&request, &docx_edit::UndoSession::new())
            .unwrap()
            .unwrap()
            .applied
    );
}

fn grow(engine: &EngineSession, class: EditClass) {
    if class == EditClass::Bulk {
        apply_host_edit(
            engine,
            json!({
                "op": "insertText", "at": "end", "text": added_words(),
                "target": {"kind": "paragraph", "story": "body", "paraId": EDIT_ID},
            }),
        );
        return;
    }
    let range = paragraph_range(engine, "body", EDIT_ID);
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", range.end - 1),
            &added_words(),
            FormatPolicy::Inherit,
        )
        .unwrap();
}

fn pages(snapshot: &Snapshot) -> &[Value] {
    snapshot.output["layout"]["pages"].as_array().unwrap()
}

fn block_page(snapshot: &Snapshot, block_id: &Value) -> usize {
    let matches: Vec<_> = pages(snapshot)
        .iter()
        .enumerate()
        .filter(|(_, page)| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .any(|fragment| &fragment["blockId"] == block_id)
        })
        .map(|(index, _)| index)
        .collect();
    assert_eq!(matches.len(), 1, "block {block_id}: {matches:?}");
    matches[0]
}

fn reference_page(snapshot: &Snapshot, note_id: u32) -> Option<usize> {
    snapshot.inputs["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["runs"].as_array().is_some_and(|runs| {
                runs.iter()
                    .any(|run| run["footnoteRefId"].as_f64() == Some(f64::from(note_id)))
            })
        })
        .map(|measured| block_page(snapshot, &measured["block"]["id"]))
}

fn note_pages(snapshot: &Snapshot, note_id: u32) -> Vec<usize> {
    pages(snapshot)
        .iter()
        .enumerate()
        .filter(|(_, page)| {
            page["noteAreas"].as_array().is_some_and(|areas| {
                areas.iter().any(|area| {
                    area["kind"] == "footnote"
                        && area["notes"]
                            .as_array()
                            .unwrap()
                            .iter()
                            .any(|note| note["id"].as_u64() == Some(u64::from(note_id)))
                })
            })
        })
        .map(|(index, _)| index)
        .collect()
}

fn note_area(snapshot: &Snapshot, note_id: u32) -> &Value {
    let locations = note_pages(snapshot, note_id);
    assert_eq!(locations.len(), 1);
    pages(snapshot)[locations[0]]["noteAreas"]
        .as_array()
        .unwrap()
        .iter()
        .find(|area| {
            area["kind"] == "footnote"
                && area["notes"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|note| note["id"].as_u64() == Some(u64::from(note_id)))
        })
        .unwrap()
}

fn assert_note_on_page(snapshot: &Snapshot, note_id: u32, page: usize) {
    assert_eq!(reference_page(snapshot, note_id), Some(page));
    assert_eq!(note_pages(snapshot, note_id), [page]);
    assert!(
        pages(snapshot)[page]["footnoteReservedHeight"]
            .as_f64()
            .unwrap()
            > 0.0
    );
}

#[test]
fn lengthening_a_paragraph_moves_a_footnote_forward_and_resumes() {
    let body = format!("{}{}<w:sectPr/>", moving_body(false, ""), tail());
    let bytes = fixture::with_body_and_note(&body, &note());
    for class in [EditClass::Interactive, EditClass::Bulk] {
        let (before, after) = edit_and_compare(&bytes, small_sections(), class, grow);
        assert_note_on_page(&before, 1, 2);
        assert_note_on_page(&after, 1, 3);
        assert!(
            pages(&after)[2]["noteAreas"]
                .as_array()
                .is_none_or(Vec::is_empty)
        );
    }
}

#[test]
fn moving_a_footnote_reuses_the_unchanged_tail_for_layout_and_display() {
    for class in [EditClass::Interactive, EditClass::Bulk] {
        assert_moving_footnote_tail_reuse(class);
    }
}

fn assert_moving_footnote_tail_reuse(class: EditClass) {
    let tail: String = (0..20)
        .map(|id| {
            paragraph(
                &format!("{:08X}", 0x7100_0200 + id),
                20,
                if id == 0 { "<w:pageBreakBefore/>" } else { "" },
                &fixture::r("Tail"),
            )
        })
        .collect();
    let body = format!(
        "{}{}{tail}<w:sectPr/>",
        moving_body(false, ""),
        paragraph(FOLLOWER_ID, 40, "", &fixture::r("Follower")),
    );
    let bytes = fixture::with_body_and_note(&body, &note());
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let engine = EngineSession::new(75403);
    engine.layout_document_with_regions_retained("{}").unwrap();
    let seed = docx_edit::EditingDoc::new(75405);
    seed_from_docx(&seed, &bytes).unwrap();
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    engine
        .doc()
        .set_note_separator_state(seed.note_separator_state().unwrap());
    let mut request = fixture::region_request(&engine, &bytes, font);
    request["regions"]["sections"] = small_sections();
    request["notes"]["contents"] = json!([
        {"id": 1, "noteKind": "footnote", "height": 0},
        {"id": 2, "noteKind": "footnote", "height": 0},
    ]);
    let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
    let request = request.to_string();
    let original = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let original = snapshot(&engine, &original);
    engine.build_display_list_frame(&extras, 0).unwrap();
    let before = engine.stats();

    grow(&engine, class);
    let output = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let edited = snapshot(&engine, &output);
    engine
        .build_display_list_frame(&extras, before.frame_epoch)
        .unwrap();
    let after = engine.stats();
    assert_note_on_page(&original, 1, 2);
    assert_note_on_page(&edited, 1, 3);
    assert_eq!(block_page(&original, &json!(FOLLOWER_ID)), 3);
    assert_eq!(block_page(&edited, &json!(FOLLOWER_ID)), 3);
    assert_eq!(pages(&edited).len(), pages(&original).len());
    assert!(pages(&edited).len() > 5);
    for id in 0..20 {
        let block_id = json!(format!("{:08X}", 0x7100_0200 + id));
        assert_eq!(block_page(&original, &block_id), 4 + id / 5);
        assert_eq!(block_page(&edited, &block_id), 4 + id / 5);
    }
    assert_eq!(after.pagination_calls, before.pagination_calls + 1);
    if class == EditClass::Bulk {
        assert_eq!(
            after.incremental_pagination_calls,
            before.incremental_pagination_calls + 1,
            "{class:?}",
        );
        assert_eq!(after.rebuilt_pages, 3);
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds + 1,
        );
        assert_eq!(
            after.rebuilt_display_pages - before.rebuilt_display_pages,
            3
        );
    } else {
        assert_eq!(
            after.incremental_pagination_calls, before.incremental_pagination_calls,
            "{class:?}",
        );
        assert_eq!(after.rebuilt_pages, pages(&edited).len());
        assert_eq!(
            after.incremental_display_builds,
            before.incremental_display_builds,
        );
        assert_eq!(
            after.rebuilt_display_pages - before.rebuilt_display_pages,
            pages(&edited).len() as u64,
        );
    }

    let fresh = EngineSession::new(75404);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    fresh
        .doc()
        .set_note_separator_state(engine.doc().note_separator_state().unwrap());
    let expected = fresh
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(output.as_bytes(), expected.as_bytes());
    assert_eq!(engine.retained_layout_json().unwrap(), expected);
    fresh.build_display_list_frame(&extras, 0).unwrap();
    let display_pages = |engine: &EngineSession| {
        engine
            .with_display_list(|list| serde_json::to_vec(&list.pages).unwrap())
            .unwrap()
    };
    assert_eq!(display_pages(&engine), display_pages(&fresh));
    if class == EditClass::Interactive {
        assert_eq!(
            after.pagination_blocks_placed - before.pagination_blocks_placed,
            fresh.stats().pagination_blocks_placed,
        );
    }
}

#[test]
fn shortening_a_paragraph_moves_a_footnote_back_and_resumes() {
    let body = format!("{}{}<w:sectPr/>", moving_body(true, ""), tail());
    let bytes = fixture::with_body_and_note(&body, &note());
    for class in [EditClass::Interactive, EditClass::Bulk] {
        let (before, after) = edit_and_compare(&bytes, small_sections(), class, |engine, class| {
            if class == EditClass::Bulk {
                apply_host_edit(
                    engine,
                    json!({
                        "op": "deleteText",
                        "target": {
                            "kind": "range", "story": "body", "view": "accepted",
                            "start": {"paraId": EDIT_ID, "offset": 4},
                            "end": {"paraId": EDIT_ID, "offset": 4 + added_words().encode_utf16().count()},
                        },
                    }),
                );
                return;
            }
            let range = paragraph_range(engine, "body", EDIT_ID);
            engine
                .doc()
                .delete_range(
                    &EditCtx::local("", ""),
                    StoryRange::new("body", range.start + 4, range.end - 1),
                )
                .unwrap();
        });
        assert_note_on_page(&before, 1, 3);
        assert_note_on_page(&after, 1, 2);
        assert!(
            pages(&after)[3]["noteAreas"]
                .as_array()
                .is_none_or(Vec::is_empty)
        );
    }
}

#[test]
fn lengthening_section_one_moves_a_footnote_onto_the_old_boundary_page_and_resumes() {
    let break_properties = r#"<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="3000" w:h="1800"/><w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>"#;
    let second_section = format!(
        r#"{}{}<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="3600" w:h="2100"/><w:pgMar w:top="150" w:right="300" w:bottom="150" w:left="300"/></w:sectPr>"#,
        paragraph("71001003", 10, "", &reference(2, "Second reference")),
        (0..3)
            .map(|id| {
                paragraph(
                    &format!("{:08X}", 0x7100_0200 + id),
                    20,
                    "",
                    &fixture::r("Second tail"),
                )
            })
            .collect::<String>(),
    );
    let body = format!("{}{second_section}", moving_body(false, break_properties));
    let bytes = fixture::with_body_and_note(&body, &note());
    let sections = json!([
        {
            "sectionId": "first",
            "pageSize": {"w": 200, "h": 120},
            "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10},
            "sectionStart": "nextPage",
        },
        {
            "sectionId": "second",
            "pageSize": {"w": 240, "h": 140},
            "margins": {"top": 10, "right": 20, "bottom": 10, "left": 20},
            "sectionStart": "nextPage",
        },
    ]);
    for class in [EditClass::Interactive, EditClass::Bulk] {
        let (before, after) = edit_and_compare(&bytes, sections.clone(), class, grow);
        assert_note_on_page(&before, 1, 2);
        assert_note_on_page(&after, 1, 3);
        assert_note_on_page(&before, 2, 3);
        assert_note_on_page(&after, 2, 4);
        assert_eq!(pages(&before)[3]["sectionIndex"], 1);
        assert_eq!(pages(&after)[3]["sectionIndex"], 0);
        assert_eq!(pages(&after)[4]["sectionIndex"], 1);
        for snapshot in [&before, &after] {
            let first = note_pages(snapshot, 1)[0];
            let second = note_pages(snapshot, 2)[0];
            assert_eq!(note_area(snapshot, 1)["sectionId"], "first");
            assert_eq!(note_area(snapshot, 2)["sectionId"], "second");
            assert_eq!(
                pages(snapshot)[first]["size"],
                json!({"w": 200.0, "h": 120.0}),
            );
            assert_eq!(
                pages(snapshot)[second]["size"],
                json!({"w": 240.0, "h": 140.0}),
            );
            assert_eq!(pages(snapshot)[first]["margins"]["left"], 10.0);
            assert_eq!(pages(snapshot)[second]["margins"]["left"], 20.0);
        }
    }
}

#[test]
fn deleting_a_reference_paragraph_removes_its_note_and_reflows_a_later_note() {
    let body = format!(
        "{}{}{}{}{}<w:sectPr/>",
        prefix(),
        paragraph("71001000", 10, "", &fixture::r("Filler")),
        paragraph(REF_ID, 50, "", &reference(1, "Removed reference")),
        paragraph("71001003", 50, "", &reference(2, "Later reference")),
        tail(),
    );
    let bytes = fixture::with_body_and_note(&body, &note());
    let (before, after) = edit_and_compare(
        &bytes,
        small_sections(),
        EditClass::Interactive,
        |engine, _| {
            engine
                .doc()
                .delete_range(
                    &EditCtx::local("", ""),
                    paragraph_range(engine, "body", REF_ID),
                )
                .unwrap();
        },
    );
    assert_note_on_page(&before, 1, 2);
    assert_note_on_page(&before, 2, 3);
    assert_eq!(reference_page(&after, 1), None);
    assert!(note_pages(&after, 1).is_empty());
    assert_note_on_page(&after, 2, 2);
    assert_eq!(
        after.inputs["measured"].as_array().unwrap().len() + 1,
        before.inputs["measured"].as_array().unwrap().len(),
    );
    assert_eq!(pages(&after).len() + 1, pages(&before).len());
}

#[test]
fn growing_a_multiline_note_keeps_its_reference_page_and_pushes_body_text() {
    let body = format!(
        "{}{}{}{}<w:sectPr/>",
        prefix(),
        paragraph(REF_ID, 10, "", &reference(1, "Reference")),
        paragraph(FOLLOWER_ID, 50, "", &fixture::r("Follower")),
        tail(),
    );
    let note = paragraph(
        NOTE_PARA_ID,
        10,
        "",
        &format!(
            "{}<w:r><w:br/></w:r>{}",
            fixture::r("First line"),
            fixture::r("Second line"),
        ),
    );
    let bytes = fixture::with_body_and_note(&body, &note);
    let (before, after) = edit_and_compare(
        &bytes,
        small_sections(),
        EditClass::Interactive,
        |engine, _| {
            let range = paragraph_range(engine, "fn:1", NOTE_PARA_ID);
            engine
                .doc()
                .insert_text(
                    &EditCtx::local("", ""),
                    Position::new("fn:1", range.end - 1),
                    &added_words(),
                    FormatPolicy::Inherit,
                )
                .unwrap();
        },
    );
    assert_note_on_page(&before, 1, 2);
    assert_note_on_page(&after, 1, 2);
    assert_eq!(block_page(&before, &json!(FOLLOWER_ID)), 2);
    assert_eq!(block_page(&after, &json!(FOLLOWER_ID)), 3);
    assert!(
        note_area(&after, 1)["height"].as_f64().unwrap()
            > note_area(&before, 1)["height"].as_f64().unwrap()
    );
    assert!(
        pages(&after)[2]["footnoteReservedHeight"].as_f64().unwrap()
            > pages(&before)[2]["footnoteReservedHeight"]
                .as_f64()
                .unwrap()
    );
}
