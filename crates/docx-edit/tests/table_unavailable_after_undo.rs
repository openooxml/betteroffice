//! Two collaborators each change one table and undo in crossing order: the table embed loses its
//! properties, grid and rows on every replica. The rest of the document still lays out, table
//! operations on that table refuse with a coded error and change nothing, the table can be
//! deleted, and the second undoer's redo restores it. Undos in synced order are unaffected.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::*;

const DATE: &str = "2026-11-10T00:00:00Z";
const CASES: [&str; 3] = ["S1", "S2", "S3"];

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn at(row: u32, column: u32) -> CellLoc {
    CellLoc::new("body", 0, row, column)
}

fn table() -> TableLocator {
    TableLocator::new("body", 0)
}

fn body() -> String {
    let cell = |id: &str, text: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>{}</w:tc>"#,
            fixture::p(id, &fixture::r(text))
        )
    };
    format!(
        concat!(
            "{}",
            r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>"#,
            r#"<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>"#,
            "<w:tr>{}{}</w:tr></w:tbl>{}"
        ),
        fixture::p("10000001", &fixture::r("Before the table")),
        cell("10000002", "First cell"),
        cell("10000003", "Second cell"),
        fixture::p("10000004", &fixture::r("After the table")),
    )
}

struct Run {
    a: EngineSession,
    b: EngineSession,
    undo_a: DocUndoManager,
    undo_b: DocUndoManager,
    request: String,
    origin: EngineSession,
}

fn exchange(left: &EngineSession, right: &EngineSession) {
    let from_left = left.doc().encode_state_as_update_v1();
    let from_right = right.doc().encode_state_as_update_v1();
    right.doc().apply_update_v1(&from_left).unwrap();
    left.doc().apply_update_v1(&from_right).unwrap();
}

fn joined(state: &[u8], client: u64) -> EngineSession {
    let engine = EngineSession::new(client);
    engine.doc().apply_update_v1(state).unwrap();
    engine
}

/// Both collaborators change the table, then A undoes and B undoes. `synced`: B's change has
/// reached A before A undoes; otherwise A undoes first and B's change arrives after.
fn run(case: &str, synced: bool) -> Run {
    let (origin, request) = fixture::laid_out(&fixture::with_body(&body()), 100);
    let state = origin.doc().encode_state_as_update_v1();
    let a = joined(&state, 61);
    let b = joined(&state, 62);
    let mut undo_a = a.doc().undo_manager();
    let mut undo_b = b.doc().undo_manager();
    match case {
        "S1" => a.doc().split_cell_grid(&ctx(), &at(0, 1), Some(1), Some(2)),
        "S2" => a
            .doc()
            .set_cell_shading(&ctx(), &TableRange::cell(at(0, 0)), Some("FF0000")),
        _ => a.doc().insert_row(&ctx(), &at(0, 0), true),
    }
    .unwrap();
    exchange(&a, &b);
    match case {
        "S1" => b.doc().insert_row(&ctx(), &at(0, 0), true),
        "S2" => b
            .doc()
            .set_cell_shading(&ctx(), &TableRange::cell(at(0, 1)), Some("00FF00")),
        _ => b.doc().set_column_width(&ctx(), &at(0, 0), 2000.0),
    }
    .unwrap();
    if synced {
        exchange(&a, &b);
    }
    undo_a.undo();
    exchange(&a, &b);
    undo_b.undo();
    exchange(&a, &b);
    Run {
        a,
        b,
        undo_a,
        undo_b,
        request,
        origin,
    }
}

fn keys(engine: &EngineSession) -> [bool; 3] {
    let payload = engine.doc().table_payload(&table()).unwrap();
    ["tblPr", "grid", "rows"].map(|key| payload.contains_key(key))
}

/// Lays out `engine` fully and builds its frame; returns the laid out text.
fn laid_out_text(engine: &EngineSession, request: &str) -> String {
    yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default()).unwrap();
    engine.layout_document_with_regions_json(request).unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
    engine.retained_kernel_inputs_json().unwrap()
}

fn table_ops(doc: &EditingDoc) -> Vec<(&'static str, OpResult<TableReceipt>)> {
    vec![
        ("insert_row", doc.insert_row(&ctx(), &at(0, 0), true)),
        ("insert_column", doc.insert_column(&ctx(), &at(0, 0), true)),
        (
            "delete_row",
            doc.delete_row(&ctx(), &TableRange::cell(at(0, 0))),
        ),
        (
            "set_cell_shading",
            doc.set_cell_shading(&ctx(), &TableRange::cell(at(0, 0)), Some("0000FF")),
        ),
        (
            "set_column_width",
            doc.set_column_width(&ctx(), &at(0, 0), 3000.0),
        ),
        (
            "split_cell_grid",
            doc.split_cell_grid(&ctx(), &at(0, 0), Some(2), None),
        ),
        (
            "merge_cells",
            doc.merge_cells(&ctx(), &TableRange::new(at(0, 0), at(0, 1))),
        ),
    ]
}

#[test]
fn crossing_undos_leave_a_table_every_replica_and_a_late_joiner_lay_out_without() {
    for case in CASES {
        let run = run(case, false);
        let late = joined(&run.a.doc().encode_state_as_update_v1(), 70);
        for (name, engine) in [("A", &run.a), ("B", &run.b), ("late", &late)] {
            assert_eq!(
                keys(engine),
                [false; 3],
                "{case} {name}: the table lost its payload"
            );
            let text = laid_out_text(engine, &run.request);
            assert!(text.contains("Before the table"), "{case} {name}");
            assert!(text.contains("After the table"), "{case} {name}");
            assert!(
                !text.contains("First cell"),
                "{case} {name}: the table is skipped"
            );
            let blocks =
                yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default()).unwrap();
            let paragraphs: Vec<_> = blocks
                .iter()
                .filter_map(|block| match block {
                    docx_layout::types::LayoutBlock::Paragraph(paragraph) => {
                        Some((paragraph.pm_start.unwrap(), paragraph.pm_end.unwrap()))
                    }
                    _ => None,
                })
                .collect();
            assert!(paragraphs.len() >= 2, "{case} {name}");
            assert_eq!(
                paragraphs[1].0,
                paragraphs[0].1 + 2.0,
                "{case} {name}: the table keeps an empty table's two display units"
            );
        }
        drop(run.origin);
    }
}

#[test]
fn table_operations_on_an_unavailable_table_refuse_with_a_coded_error_and_change_nothing() {
    for case in CASES {
        let run = run(case, false);
        for engine in [&run.a, &run.b] {
            let before = engine.doc().encode_state_as_update_v1();
            for (name, result) in table_ops(engine.doc()) {
                let error = result.unwrap_err();
                assert_eq!(format!("{error:?}"), "TableUnavailable", "{case} {name}");
                assert!(
                    error
                        .to_string()
                        .starts_with("a table is unavailable after a concurrent undo"),
                    "{case} {name}: {error}"
                );
                assert_eq!(
                    engine.doc().encode_state_as_update_v1(),
                    before,
                    "{case} {name}: state unchanged"
                );
            }
        }
        let refused = run.a.doc().insert_row(&ctx(), &at(0, 0), true).unwrap_err();
        let coded: docx_edit::session::EngineError = refused.clone().into();
        assert_eq!(format!("{:?}", coded.kind), "TableUnavailable");
        let by_text = docx_edit::session::EngineError::message(refused);
        assert_eq!(format!("{:?}", by_text.kind), "TableUnavailable");
    }
}

#[test]
fn an_unavailable_table_can_be_deleted_and_the_rest_keeps_laying_out() {
    for case in CASES {
        let run = run(case, false);
        run.a.doc().delete_table(&ctx(), &table()).unwrap();
        exchange(&run.a, &run.b);
        for (name, engine) in [("A", &run.a), ("B", &run.b)] {
            assert!(
                matches!(
                    engine.doc().table_payload(&table()),
                    Err(OpError::UnknownTable { .. })
                ),
                "{case} {name}: the table is gone"
            );
            let text = laid_out_text(engine, &run.request);
            assert!(text.contains("Before the table") && text.contains("After the table"));
        }
    }
}

#[test]
fn the_second_undoer_redo_restores_the_table_and_the_first_undoer_redo_does_not() {
    for case in CASES {
        let mut run = run(case, false);
        run.undo_a.redo();
        exchange(&run.a, &run.b);
        assert_eq!(
            keys(&run.a),
            [false; 3],
            "{case}: A's redo leaves it unavailable"
        );
        run.undo_b.redo();
        exchange(&run.a, &run.b);
        for (name, engine) in [("A", &run.a), ("B", &run.b)] {
            assert_eq!(
                keys(engine),
                [true; 3],
                "{case} {name}: B's redo restores it"
            );
            assert!(laid_out_text(engine, &run.request).contains("First cell"));
            engine.doc().insert_row(&ctx(), &at(0, 0), true).unwrap();
        }
    }
}

#[test]
fn undos_in_synced_order_keep_the_table() {
    for case in CASES {
        let run = run(case, true);
        for (name, engine) in [("A", &run.a), ("B", &run.b)] {
            assert_eq!(keys(engine), [true; 3], "{case} {name}");
            if let Err(error) = engine.doc().set_column_width(&ctx(), &at(0, 0), 2500.0) {
                assert_ne!(format!("{error:?}"), "TableUnavailable", "{case} {name}");
            }
            if case != "S2" {
                // The undone split or row leaves a cell story the other replica's rows still
                // name; a table cell without its story is a separate fix.
                continue;
            }
            assert!(laid_out_text(engine, &run.request).contains("First cell"));
            engine.doc().insert_row(&ctx(), &at(0, 0), true).unwrap();
        }
    }
}

/// Display starts of the body's paragraphs as the lowering gives them and as the geometry outline
/// gives them.
fn paragraph_starts(engine: &EngineSession) -> (Vec<f64>, Vec<f64>) {
    let lowered = yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default())
        .unwrap()
        .iter()
        .filter_map(|block| match block {
            docx_layout::types::LayoutBlock::Paragraph(paragraph) => paragraph.pm_start,
            _ => None,
        })
        .collect();
    let GeometryRead::Value(outline) = engine.doc().geometry_position_outline("body") else {
        panic!("no geometry outline");
    };
    let outlined = serde_json::to_value(&outline["body"]).unwrap()["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .map(|paragraph| paragraph["displayStart"].as_f64().unwrap())
        .collect();
    (lowered, outlined)
}

#[test]
fn a_table_that_lost_only_its_properties_takes_two_units_in_layout_and_geometry_alike() {
    let (origin, request) = fixture::laid_out(&fixture::with_body(&body()), 100);
    let state = origin.doc().encode_state_as_update_v1();
    let a = joined(&state, 61);
    let b = joined(&state, 62);
    let mut undo_a = a.doc().undo_manager();
    let mut undo_b = b.doc().undo_manager();
    let patch =
        |key: &str, value: yrs::Any| std::collections::HashMap::from([(key.to_owned(), value)]);
    a.doc()
        .set_table_properties(&ctx(), &table(), &patch("width", yrs::Any::from(4000.0)))
        .unwrap();
    exchange(&a, &b);
    b.doc()
        .set_table_properties(
            &ctx(),
            &table(),
            &patch("justification", yrs::Any::from("center")),
        )
        .unwrap();
    assert!(undo_a.undo());
    exchange(&a, &b);
    assert!(undo_b.undo());
    exchange(&a, &b);
    for (name, engine) in [("A", &a), ("B", &b)] {
        assert_eq!(
            keys(engine),
            [false, true, true],
            "{name}: only tblPr is gone"
        );
        let text = laid_out_text(engine, &request);
        assert!(
            text.contains("After the table") && !text.contains("First cell"),
            "{name}"
        );
        let (lowered, outlined) = paragraph_starts(engine);
        assert_eq!(lowered.len(), outlined.len(), "{name}");
        assert!(lowered.len() >= 2, "{name}");
        assert_eq!(
            lowered[1] - lowered[0],
            outlined[1] - outlined[0],
            "{name}: layout and geometry agree after the table"
        );
        let error = engine
            .doc()
            .insert_row(&ctx(), &at(0, 0), true)
            .unwrap_err();
        assert_eq!(format!("{error:?}"), "TableUnavailable", "{name}");
    }
}
