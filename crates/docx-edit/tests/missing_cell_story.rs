//! A table cell that a live table names while its story is missing: one replica splits a cell,
//! another inserts a row, the first undoes the split. Every replica reads the cell as one empty
//! paragraph, lays the document out, and refuses every write into the cell without changing
//! anything, until a redo, a row deletion or a table deletion makes the state whole again.

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::canonical::{project_story, story_checksum};
use docx_edit::*;
use serde_json::json;
use yrs::Any;

const DATE: &str = "2026-11-10T00:00:00Z";
const LEFT: u64 = 71;
const RIGHT: u64 = 72;

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn exchange(left: &EditingDoc, right: &EditingDoc) {
    let from_left = left.encode_state_as_update_v1();
    let from_right = right.encode_state_as_update_v1();
    right.apply_update_v1(&from_left).unwrap();
    left.apply_update_v1(&from_right).unwrap();
}

fn origin() -> EditingDoc {
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    let table = origin
        .insert_table(&ctx(), Position::new("body", 0), 1, 2)
        .unwrap();
    for (index, story) in table.created_story_ids.iter().enumerate() {
        origin
            .insert_text(
                &ctx(),
                Position::new(story, 0),
                &format!("cell{index}"),
                FormatPolicy::Plain,
            )
            .unwrap();
    }
    origin
}

fn replica(origin: &EditingDoc, client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    doc.apply_update_v1(&origin.encode_state_as_update_v1())
        .unwrap();
    doc
}

/// Every cell the body's first table names, as `(row, column, story)`.
fn named_cells(doc: &EditingDoc) -> Vec<(u32, u32, String)> {
    let payload = doc.table_payload(&TableLocator::new("body", 0)).unwrap();
    let Some(Any::Array(rows)) = payload.get("rows") else {
        panic!("table without rows: {payload:?}");
    };
    let mut named = Vec::new();
    for (row_index, row) in rows.iter().enumerate() {
        let Any::Map(row) = row else {
            panic!("row is not a map: {row:?}");
        };
        let Some(Any::Array(cells)) = row.get("cells") else {
            panic!("row without cells: {row:?}");
        };
        for (column, cell) in cells.iter().enumerate() {
            let Any::Map(cell) = cell else {
                panic!("cell is not a map: {cell:?}");
            };
            let Some(Any::String(story)) = cell.get("story") else {
                panic!("cell without story: {cell:?}");
            };
            named.push((row_index as u32, column as u32, story.to_string()));
        }
    }
    named
}

fn listed(doc: &EditingDoc, story: &str) -> bool {
    doc.list_stories()
        .stories
        .iter()
        .any(|info| info.story == story)
}

/// The named cells without a story, as `(row, column, story)`.
fn dangling(doc: &EditingDoc) -> Vec<(u32, u32, String)> {
    named_cells(doc)
        .into_iter()
        .filter(|(_, _, story)| !listed(doc, story))
        .collect()
}

/// `(splitter, peer, splitter's undo manager)` after the splitter split cell (0, 1) into two
/// columns, the peer inserted a row below row 0, and the splitter undid its split: after the
/// row reached it when `synced`, else concurrently with the row.
fn dangling_state(synced: bool) -> (EditingDoc, EditingDoc, DocUndoManager) {
    let origin = origin();
    let splitter = replica(&origin, LEFT);
    let peer = replica(&origin, RIGHT);
    let mut undo = splitter.undo_manager();
    splitter
        .split_cell_grid(&ctx(), &CellLoc::new("body", 0, 0, 1), Some(1), Some(2))
        .unwrap();
    exchange(&splitter, &peer);
    peer.insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    if synced {
        exchange(&splitter, &peer);
    }
    assert!(undo.undo());
    exchange(&splitter, &peer);
    (splitter, peer, undo)
}

fn state(doc: &EditingDoc) -> (Vec<u8>, Vec<u8>) {
    (
        doc.encode_state_vector_v1(),
        doc.encode_state_as_update_v1(),
    )
}

fn assert_reads_empty(doc: &EditingDoc, cell: &str) {
    let key = format!("{cell}#p0");
    let paragraphs = doc.paragraphs(cell).unwrap();
    assert_eq!(paragraphs.len(), 1, "{cell}");
    assert_eq!(
        (paragraphs[0].para_id.as_str(), paragraphs[0].text.as_str()),
        (key.as_str(), "")
    );
    assert_eq!(doc.story_len(cell).unwrap(), 1);
    assert_eq!(doc.story_segments(cell).unwrap().len(), 1);
    assert_eq!(doc.story_segment_units(cell).unwrap().len(), 1);
    assert_eq!(project_story(doc, cell).unwrap().len(), 1);
    story_checksum(doc, cell).unwrap();
    assert_eq!(doc.story_paragraph_ids(cell).unwrap(), vec![None]);
    assert!(doc.list_changes(cell).unwrap().is_empty());
    assert_eq!(
        doc.paragraph_mark_position(&key).unwrap(),
        Position::new(cell, 0)
    );

    let loc = Loc::new(cell, key.as_str(), 0);
    let range = LocRange {
        start: loc.clone(),
        end: loc.clone(),
    };
    assert_eq!(doc.locate(&loc).unwrap(), Position::new(cell, 0));
    assert_eq!(doc.loc_at(&Position::new(cell, 0)).unwrap(), loc);
    assert_eq!(
        doc.locate_range(&range).unwrap(),
        StoryRange::new(cell, 0, 0)
    );
    assert_eq!(
        doc.loc_range_of(&StoryRange::new(cell, 0, 0)).unwrap(),
        range
    );
    let context = doc.selection_context(&StoryRange::new(cell, 0, 0)).unwrap();
    assert_eq!(context.para_id, key);
    assert!(context.in_table);
    assert_eq!(doc.selection_info(&loc, &loc).unwrap().paragraph_text, "");
    assert_eq!(
        doc.selection_text(&range, EditTextView::Accepted)
            .unwrap()
            .paragraph_text,
        ""
    );
    assert_eq!(doc.text_between(&range, TextView::Raw).unwrap(), "");
    assert_eq!(
        doc.nav_boundary(&loc, NavUnit::Grapheme, NavDirection::Next)
            .unwrap(),
        loc
    );
    assert!(
        doc.find_in_document(cell, "cell", FindOptions::default())
            .unwrap()
            .is_empty()
    );
    let read = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: Some(cell.to_owned()),
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap();
    assert_eq!(read.paragraphs.len(), 1);
    assert_eq!(read.paragraphs[0].para_id, key);
    assert_eq!(read.paragraphs[0].text, "");

    let blocks = yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
    let lowered = serde_json::to_string(&blocks).unwrap();
    assert!(lowered.contains("\"table\""), "{lowered}");
}

fn assert_refused<T: std::fmt::Debug, E: std::fmt::Display>(
    result: Result<T, E>,
    cell: &str,
    what: &str,
) {
    match result {
        Err(error) => {
            let message = error.to_string();
            assert!(
                message.contains("has no story") && message.contains(cell),
                "{what}: {message}"
            );
        }
        Ok(value) => panic!("{what} was not refused: {value:?}"),
    }
}

fn assert_writes_refused(doc: &EditingDoc, row: u32, column: u32, cell: &str) {
    let key = format!("{cell}#p0");
    let at = || Position::new(cell, 0);
    let mark = || StoryRange::new(cell, 0, 1);
    let before = state(doc);
    assert_refused(
        doc.insert_text(&ctx(), at(), "x", FormatPolicy::Plain),
        cell,
        "insert_text",
    );
    assert_refused(
        doc.insert_text(&ctx().suggesting(), at(), "x", FormatPolicy::Plain),
        cell,
        "suggested insert_text",
    );
    assert_refused(
        doc.replace_range(&ctx(), StoryRange::new(cell, 0, 0), "x"),
        cell,
        "replace_range",
    );
    assert_refused(doc.delete_range(&ctx(), mark()), cell, "delete_range");
    assert_refused(
        doc.toggle_format_keeping(&ctx(), mark(), SimpleFormat::Bold),
        cell,
        "toggle_format_keeping",
    );
    assert_refused(
        doc.clear_formatting_keeping(&ctx(), mark()),
        cell,
        "clear_formatting_keeping",
    );
    assert_refused(
        doc.set_hyperlink(&ctx(), mark(), Some(Any::from("https://example.com"))),
        cell,
        "set_hyperlink",
    );
    assert_refused(
        doc.split_paragraph(&ctx(), at(), None),
        cell,
        "split_paragraph",
    );
    assert_refused(
        doc.insert_hard_break(&ctx(), at()),
        cell,
        "insert_hard_break",
    );
    assert_refused(doc.insert_tab(&ctx(), at()), cell, "insert_tab");
    assert_refused(
        doc.merge_paragraphs(&ctx(), &key, MergeDirection::Forward),
        cell,
        "merge_paragraphs",
    );
    assert_refused(
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::One(key.clone()),
            &ParaAttrDelta::default(),
        ),
        cell,
        "set_paragraph_attrs by id",
    );
    assert_refused(
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::Range(StoryRange::new(cell, 0, 0)),
            &ParaAttrDelta::default(),
        ),
        cell,
        "set_paragraph_attrs by range",
    );
    assert_refused(
        doc.set_paragraph_attr(&key, "keepNext", Any::Bool(true)),
        cell,
        "set_paragraph_attr",
    );
    assert_refused(
        doc.insert_embed(&ctx(), at(), "pageBreak", Vec::new()),
        cell,
        "insert_embed",
    );
    assert_refused(doc.insert_table(&ctx(), at(), 1, 1), cell, "insert_table");
    assert_refused(
        doc.add_comment(&[mark()], "Ada", DATE, Any::from("note")),
        cell,
        "add_comment",
    );
    assert_refused(
        doc.apply_raw_ops(
            cell,
            vec![RawOp::Insert {
                index: 0,
                text: "x".to_owned(),
                attrs: Default::default(),
            }],
            &ctx(),
        ),
        cell,
        "apply_raw_ops",
    );
    assert_refused(
        doc.merge_cells(
            &ctx(),
            &TableRange::new(
                CellLoc::new("body", 0, row, column - 1),
                CellLoc::new("body", 0, row, column),
            ),
        ),
        cell,
        "merge_cells into the left neighbour",
    );
    if named_cells(doc)
        .iter()
        .any(|(at_row, at_column, _)| *at_row == row + 1 && *at_column == column)
    {
        assert_refused(
            doc.merge_cells(
                &ctx(),
                &TableRange::new(
                    CellLoc::new("body", 0, row, column),
                    CellLoc::new("body", 0, row + 1, column),
                ),
            ),
            cell,
            "merge_cells with the cell below",
        );
    }
    assert_refused(
        doc.split_cell_grid(
            &ctx(),
            &CellLoc::new("body", 0, row, column),
            Some(2),
            Some(1),
        ),
        cell,
        "split_cell_grid",
    );
    assert_refused(doc.delete_story(cell), cell, "delete_story");
    assert_refused(
        doc.create_story(cell, "x", "Normal", "left"),
        cell,
        "create_story",
    );
    assert_refused(
        doc.create_story_with_paragraph_id(cell, "p", "x", "Normal", "left"),
        cell,
        "create_story_with_paragraph_id",
    );
    let seed = SeedParagraph {
        text: "x".to_owned(),
        p_style: "Normal".to_owned(),
        alignment: "left".to_owned(),
    };
    assert_refused(doc.seed_story(cell, &[seed]), cell, "seed_story");

    let collapsed = || StoryRange::new(cell, 0, 0);
    assert_refused(
        doc.insert_text(&ctx(), at(), "a\nb", FormatPolicy::Plain),
        cell,
        "insert_text with a break",
    );
    assert_refused(
        doc.delete_range(&ctx(), collapsed()),
        cell,
        "collapsed delete_range",
    );
    assert_refused(
        doc.replace_range(&ctx(), collapsed(), ""),
        cell,
        "collapsed replace_range with no text",
    );
    assert_refused(
        doc.toggle_format_keeping(&ctx(), collapsed(), SimpleFormat::Bold),
        cell,
        "collapsed toggle_format_keeping",
    );
    assert_refused(
        doc.clear_formatting_keeping(&ctx(), collapsed()),
        cell,
        "collapsed clear_formatting_keeping",
    );
    assert_refused(
        doc.set_hyperlink(&ctx(), collapsed(), None),
        cell,
        "collapsed set_hyperlink",
    );
    assert_refused(
        doc.merge_paragraphs(&ctx(), &key, MergeDirection::Backward),
        cell,
        "merge_paragraphs backward from the first paragraph",
    );
    assert_refused(
        doc.add_comment(&[collapsed()], "Ada", DATE, Any::from("note")),
        cell,
        "collapsed add_comment",
    );
    assert_refused(
        doc.set_comment_ranges("missing", &[collapsed()]),
        cell,
        "collapsed set_comment_ranges",
    );
    for kind in ["sectionBreak", "watermark", "pilcrow"] {
        assert_refused(doc.insert_embed(&ctx(), at(), kind, Vec::new()), cell, kind);
    }
    assert_refused(
        doc.move_embed(&ctx(), Position::new("body", 0), at()),
        cell,
        "move_embed into the cell",
    );

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": doc.version().as_str(),
        "steps": [{
            "op": "replaceText",
            "target": {"kind": "paragraph", "story": cell, "paraId": key},
            "text": "x",
        }],
    }))
    .unwrap();
    let refusal = doc
        .apply_edits(&request, &UndoSession::new())
        .unwrap()
        .expect_err("a batch step into the cell is refused");
    assert_eq!(refusal.failure.code, EditFailureCode::Unsupported);
    assert_eq!(
        refusal.failure.reason,
        Some(EditFailureReason::CellHasNoStory)
    );
    assert!(
        refusal.failure.message.contains("has no story"),
        "{}",
        refusal.failure.message
    );
    assert_eq!(state(doc), before, "refused writes change nothing");
}

#[test]
fn a_cell_an_undo_left_without_its_story_reads_empty_and_refuses_writes() {
    for synced in [true, false] {
        let (left, right, _undo) = dangling_state(synced);
        let cells = dangling(&left);
        assert!(!cells.is_empty(), "synced={synced}");
        assert_eq!(dangling(&right), cells);
        let (row, column, cell) = cells[0].clone();
        assert!(column > 0);
        for doc in [&left, &right] {
            let before = state(doc);
            assert_reads_empty(doc, &cell);
            assert_eq!(state(doc), before, "reads never write");
            assert_writes_refused(doc, row, column, &cell);
        }
        let (_, _, neighbour) = named_cells(&left)
            .into_iter()
            .find(|(at_row, at_column, _)| *at_row == row && *at_column == 0)
            .unwrap();
        left.insert_text(
            &ctx(),
            Position::new(neighbour.as_str(), 0),
            "kept",
            FormatPolicy::Plain,
        )
        .unwrap();
        exchange(&left, &right);
        for doc in [&left, &right] {
            let paragraphs = doc.paragraphs(&neighbour).unwrap();
            assert!(paragraphs[0].text.starts_with("kept"));
            assert_reads_empty(doc, &cell);
        }
    }
}

#[test]
fn a_redo_or_removing_the_row_or_table_makes_the_state_whole() {
    for recovery in ["redo", "row", "table"] {
        let (splitter, peer, mut undo) = dangling_state(true);
        let (row, _, cell) = dangling(&splitter).remove(0);
        match recovery {
            "redo" => assert!(undo.redo()),
            "row" => {
                let at = CellLoc::new("body", 0, row, 0);
                peer.delete_row(&ctx(), &TableRange::new(at.clone(), at))
                    .unwrap();
            }
            _ => {
                peer.delete_table(&ctx(), &TableLocator::new("body", 0))
                    .unwrap();
            }
        }
        exchange(&splitter, &peer);
        for doc in [&splitter, &peer] {
            yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
            if recovery == "redo" {
                assert!(dangling(doc).is_empty(), "{recovery}");
                assert!(listed(doc, &cell), "{recovery}");
                doc.insert_text(
                    &ctx(),
                    Position::new(cell.as_str(), 0),
                    "x",
                    FormatPolicy::Plain,
                )
                .unwrap();
            } else {
                if recovery == "row" {
                    assert!(dangling(doc).is_empty(), "{recovery}");
                }
                assert!(
                    matches!(doc.story_len(&cell), Err(EditError::StoryNotFound(_))),
                    "{recovery}"
                );
            }
        }
    }
}

#[test]
fn resident_lowering_and_font_requirements_read_the_cell_as_one_empty_paragraph() {
    let (splitter, _peer, _undo) = dangling_state(true);
    let (_, _, cell) = dangling(&splitter).remove(0);
    let engine = EngineSession::new(73);
    engine
        .doc()
        .apply_update_v1(&splitter.encode_state_as_update_v1())
        .unwrap();
    let before = state(engine.doc());
    let count = engine
        .with_lowered_story(&cell, &RenderEnv::default(), |blocks| blocks.len())
        .unwrap();
    assert_eq!(count, 1);
    let lowered = engine
        .lower_story_json(&cell, &RenderEnv::default())
        .unwrap();
    let lowered: serde_json::Value = serde_json::from_str(&lowered).unwrap();
    assert_eq!(lowered.as_array().map(Vec::len), Some(1), "{lowered}");
    let request = json!({"bodyStory": cell, "renderEnv": {}}).to_string();
    engine.layout_font_requirements_json(&request).unwrap();
    assert_eq!(state(engine.doc()), before);
}

#[test]
fn a_comment_anchored_in_the_cell_before_the_undo_resolves_to_its_empty_paragraph() {
    let origin = origin();
    let splitter = replica(&origin, LEFT);
    let peer = replica(&origin, RIGHT);
    let mut undo = splitter.undo_manager();
    splitter
        .split_cell_grid(&ctx(), &CellLoc::new("body", 0, 0, 1), Some(1), Some(2))
        .unwrap();
    exchange(&splitter, &peer);
    let (_, _, cell) = named_cells(&peer)
        .into_iter()
        .find(|(row, column, _)| (*row, *column) == (0, 2))
        .unwrap();
    let range = StoryRange::new(cell.as_str(), 0, 1);
    let comment = peer
        .add_comment(&[range], "Ada", DATE, Any::from("note"))
        .unwrap();
    peer.insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    exchange(&splitter, &peer);
    assert!(undo.undo());
    exchange(&splitter, &peer);
    let key = format!("{cell}#p0");
    let at = Loc::new(cell.as_str(), key.as_str(), 0);
    let empty = LocRange {
        start: at.clone(),
        end: at,
    };
    for doc in [&splitter, &peer] {
        assert!(dangling(doc).iter().any(|(_, _, story)| *story == cell));
        let before = state(doc);
        let expected = ResolvedCommentAnchor {
            story: cell.clone(),
            start: 0,
            end: 0,
        };
        assert_eq!(doc.resolve_comment(&comment).unwrap(), [expected]);
        assert_eq!(doc.find_comment_range(&comment).unwrap(), empty);
        let listed = doc.list_comments().unwrap();
        let info = listed.iter().find(|info| info.id == comment).unwrap();
        assert_eq!(info.ranges, std::slice::from_ref(&empty));
        assert_eq!(state(doc), before);
    }
}

#[test]
fn table_edits_after_the_undo_never_reuse_the_id_of_the_cell_without_its_story() {
    for synced in [true, false] {
        let (left, right, _undo) = dangling_state(synced);
        let missing = dangling(&left);
        let (_, _, cell) = missing[0].clone();
        let at = |column| CellLoc::new("body", 0, 0, column);
        let split = left
            .split_cell_grid(&ctx(), &at(0), Some(1), Some(2))
            .unwrap();
        assert!(!split.created_story_ids.contains(&cell), "synced={synced}");
        left.merge_cells(&ctx(), &TableRange::new(at(0), at(1)))
            .unwrap();
        left.insert_row(&ctx(), &at(0), true).unwrap();
        left.insert_column(&ctx(), &at(0), true).unwrap();
        left.split_cell_grid(&ctx(), &at(0), Some(1), Some(2))
            .unwrap();
        exchange(&left, &right);
        for doc in [&left, &right] {
            let named: Vec<String> = named_cells(doc)
                .into_iter()
                .map(|(_, _, story)| story)
                .collect();
            let unique: std::collections::HashSet<&String> = named.iter().collect();
            assert_eq!(unique.len(), named.len(), "{named:?}");
            assert_eq!(dangling(doc).len(), missing.len(), "synced={synced}");
            assert_reads_empty(doc, &cell);
        }
    }
}

/// The split, row and undo sequence of [`dangling_state`] in a header's table: the header cell
/// that the undo left without its story.
fn header_cell_without_story(synced: bool) -> (EditingDoc, EditingDoc, String) {
    const HEADER: &str = "hf:rId7";
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    origin
        .create_story(HEADER, "Header", "Normal", "left")
        .unwrap();
    origin
        .insert_table(&ctx(), Position::new(HEADER, 0), 1, 2)
        .unwrap();
    let splitter = replica(&origin, LEFT);
    let peer = replica(&origin, RIGHT);
    let mut undo = splitter.undo_manager();
    splitter
        .split_cell_grid(&ctx(), &CellLoc::new(HEADER, 0, 0, 1), Some(1), Some(2))
        .unwrap();
    exchange(&splitter, &peer);
    peer.insert_row(&ctx(), &CellLoc::new(HEADER, 0, 0, 0), true)
        .unwrap();
    if synced {
        exchange(&splitter, &peer);
    }
    assert!(undo.undo());
    exchange(&splitter, &peer);
    let payload = splitter
        .table_payload(&TableLocator::new(HEADER, 0))
        .unwrap();
    let rows = serde_json::to_value(&payload["rows"]).unwrap();
    let cell = rows[0]["cells"][2]["story"].as_str().unwrap().to_owned();
    assert!(!listed(&splitter, &cell), "{cell}");
    assert_eq!(splitter.paragraphs(&cell).unwrap()[0].text, "");
    (splitter, peer, cell)
}

#[test]
fn metadata_setters_never_create_a_header_cell_without_its_story() {
    for synced in [true, false] {
        let (left, right, cell) = header_cell_without_story(synced);
        let part = cell.strip_prefix("hf:").unwrap();
        let property = json!({"headers": [[part, {}]]}).to_string();
        let metadata = json!({"body": {}, "headers": [[part, {}]]}).to_string();
        for doc in [&left, &right] {
            let before = state(doc);
            assert_refused(
                doc.set_document_property(&ctx(), "metadata", &property),
                &cell,
                "set_document_property",
            );
            assert_refused(
                doc.set_part_metadata(&ctx(), &metadata, &[]),
                &cell,
                "set_part_metadata",
            );
            assert_eq!(state(doc), before, "synced={synced}");
            assert!(!listed(doc, &cell));
        }
    }
}

#[test]
fn host_batches_with_formatted_runs_aimed_at_the_cell_are_refused() {
    for synced in [true, false] {
        let (left, right, _undo) = dangling_state(synced);
        let (_, _, cell) = dangling(&left).remove(0);
        let key = format!("{cell}#p0");
        let target = json!({"kind": "paragraph", "story": cell, "paraId": key});
        let runs = json!([
            {"text": "bold", "marks": {"bold": true}},
            {"text": " plain"},
        ]);
        for doc in [&left, &right] {
            for step in [
                json!({"op": "insertText", "target": target, "at": "end", "runs": runs}),
                json!({"op": "replaceText", "target": target, "runs": runs}),
                json!({
                    "op": "insertParagraphs",
                    "target": {"story": cell, "paraId": key},
                    "at": "end",
                    "paragraphs": [{"runs": runs}],
                }),
            ] {
                let before = state(doc);
                let request: EditRequest = serde_json::from_value(json!({
                    "expectVersion": doc.version().as_str(),
                    "steps": [step],
                }))
                .unwrap();
                let refusal = doc
                    .apply_edits(&request, &UndoSession::new())
                    .unwrap()
                    .expect_err("a step into the cell is refused");
                assert_eq!(refusal.failure.code, EditFailureCode::Unsupported);
                assert_eq!(
                    refusal.failure.reason,
                    Some(EditFailureReason::CellHasNoStory),
                    "synced={synced}"
                );
                assert_eq!(state(doc), before, "synced={synced}");
            }
        }
    }
}
