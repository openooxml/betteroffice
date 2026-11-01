use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellAddress, CellRange, CellRef, CellState, CellValue, Op, Sheet,
    SheetId, UnevaluatedCell, UpdateOrigin, Workbook, WorkbookModel,
};
use serde_json::{Value, json};
use xlsx_model::CellProvider;

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.to_owned(),
    }
}

/// A1 asks a web service, B1 reads it through `b1`, C1 reads only the plain
/// input D1. B1's cache differs from what the engine would compute from A1, so
/// a kept cache shows.
fn data_sheet(b1: &str) -> Sheet {
    let mut sheet = Sheet::new("Data");
    for (col, formula, value) in [
        (
            0,
            Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#),
            text("live"),
        ),
        (1, Some(b1), number(9.0)),
        (2, Some("D1+1"), number(2.0)),
        (3, None, number(1.0)),
    ] {
        sheet.set_cell(
            CellRef::new(0, col),
            Cell {
                value,
                formula: formula.map(str::to_owned),
                style: None,
            },
        );
    }
    sheet
}

fn package(sheets: Vec<Sheet>) -> Vec<u8> {
    let model = WorkbookModel {
        sheets,
        ..WorkbookModel::default()
    };
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

fn source() -> Vec<u8> {
    package(vec![data_sheet("LEN(A1)*2")])
}

fn edit(workbook: &mut Workbook, col: u32, input: &str) {
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, col),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

fn value(workbook: &Workbook, col: u32) -> CellValue {
    workbook.model().value(SheetId(0), CellRef::new(0, col))
}

fn reported(cols: &[u32]) -> Vec<UnevaluatedCell> {
    cols.iter()
        .map(|&col| UnevaluatedCell {
            address: CellAddress {
                sheet: SheetId(0),
                cell: CellRef::new(0, col),
            },
            functions: vec!["WEBSERVICE".to_owned()],
        })
        .collect()
}

fn workbook_xml(bytes: &[u8]) -> String {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "xl/workbook.xml")
        .map(|(_, xml)| String::from_utf8(xml).unwrap())
        .unwrap()
}

/// A1 and B1 listed with `WEBSERVICE` in the report, the status and a read.
fn assert_reported(workbook: &Workbook) {
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
    let status: Value = serde_json::from_str(&workbook.calculation_status_json().unwrap()).unwrap();
    assert_eq!(
        status["unevaluatedCells"],
        json!([
            {"address": "A1", "functions": ["WEBSERVICE"]},
            {"address": "B1", "functions": ["WEBSERVICE"]},
        ])
    );
    let read: Value =
        serde_json::from_str(&workbook.read_cells_json(r#"{"ranges":[]}"#).unwrap()).unwrap();
    assert_eq!(
        read["calculation"]["unevaluatedCells"],
        json!([
            {"sheetId": read["sheets"][0]["sheetId"], "row": 0, "col": 0, "a1": "A1", "functions": ["WEBSERVICE"]},
            {"sheetId": read["sheets"][0]["sheetId"], "row": 0, "col": 1, "a1": "B1", "functions": ["WEBSERVICE"]},
        ])
    );
}

#[test]
fn uncomputable_formulas_keep_excel_values_through_recalculation_and_save() {
    let bytes = source();
    assert!(!workbook_xml(&bytes).contains("fullCalcOnLoad"));
    let mut workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_eq!(value(&workbook, 0), text("live"));
    assert_eq!(value(&workbook, 1), number(9.0));
    assert_reported(&workbook);
    assert_eq!(workbook.save().unwrap(), bytes);

    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 3),
            "5",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(value(&workbook, 2), number(6.0));
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
    let saved = workbook.save().unwrap();
    assert!(workbook_xml(&saved).contains(r#"fullCalcOnLoad="1""#));
    let parsed = xlsx_parse::parse_workbook(&ooxml_opc::unzip_parts(&saved).unwrap()).unwrap();
    assert_eq!(parsed.value(SheetId(0), CellRef::new(0, 0)), text("live"));
    assert_eq!(parsed.value(SheetId(0), CellRef::new(0, 1)), number(9.0));
}

#[test]
fn a_computable_workbook_reports_no_unevaluated_cells() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "=D1*3",
            CalculationOptions::default(),
        )
        .unwrap();
    let result = workbook.recalculate_all(CalculationOptions::default());
    assert!(result.unevaluated_cells.is_empty());
    assert_eq!(value(&workbook, 1), number(2.0));
    let status: Value = serde_json::from_str(&workbook.calculation_status_json().unwrap()).unwrap();
    assert_eq!(status["unevaluatedCells"], json!([]));
}

/// snapshots and recovery checkpoints carry no list, so a restored session
/// re-derives it from the formulas and keeps to it through later edits.
#[test]
fn restored_sessions_rederive_uncomputable_cells() {
    use betteroffice_xlsx::{SnapshotBudget, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder};

    let workbook = Workbook::open_recalculated(&source(), CalculationOptions::default()).unwrap();
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (hydrated, _) = builder.finish().unwrap().into_parts();
    assert_reported(&hydrated);

    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (mut recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_reported(&recovered);
    recovered
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 3),
            "5",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(value(&recovered, 1), number(9.0));
    assert_reported(&recovered);
}

/// an open without recalculation derives the list, so the first edit keeps the
/// cache of a cell reading an uncomputable one.
#[test]
fn a_plain_open_keeps_readers_of_uncomputable_cells_through_edits() {
    let mut workbook = Workbook::open(&package(vec![data_sheet("LEN(A1)*2+D1")])).unwrap();
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
    edit(&mut workbook, 3, "5");
    assert_eq!(value(&workbook, 1), number(9.0));
    assert_eq!(value(&workbook, 2), number(6.0));
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
}

#[test]
fn a_sheet_rename_keeps_unevaluated_cells() {
    let mut other = Sheet::new("Other");
    other.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: number(1.0),
            ..Cell::default()
        },
    );
    let bytes = package(vec![data_sheet("LEN(A1)*2+Other!A1"), other]);
    let mut workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(1),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(value(&workbook, 1), number(9.0));
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
    edit(&mut workbook, 3, "5");
    assert_eq!(value(&workbook, 1), number(9.0));
    assert_eq!(
        workbook.last_calculation().unevaluated_cells,
        reported(&[0, 1])
    );
}

/// reads report the list, so a recalculation changing only it moves the
/// version and notifies observers.
#[test]
fn a_recalculation_changing_only_unevaluated_cells_moves_the_version() {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: number(7.0),
            formula: Some("REPRO_UNKNOWN_FUNCTION(1)".into()),
            style: None,
        },
    );
    let mut workbook = Workbook::open(&package(vec![sheet])).unwrap();
    assert!(workbook.last_calculation().unevaluated_cells.is_empty());
    let version = workbook.version();
    let origins = Arc::new(Mutex::new(Vec::new()));
    let seen = Arc::clone(&origins);
    let _subscription = workbook
        .observe_update_v1(move |event| seen.lock().unwrap().push(event.origin))
        .unwrap();
    let result = workbook.recalculate_all(CalculationOptions::default());
    assert!(result.changed.is_empty());
    assert_eq!(result.unevaluated_cells.len(), 1);
    assert!(result.unevaluated_cells[0].functions.is_empty());
    assert_ne!(workbook.version(), version);
    assert_eq!(*origins.lock().unwrap(), [UpdateOrigin::Recalculation]);
}

/// a one-sheet package of `(row, col, formula, value)` cells.
fn cells(cells: &[(u32, u32, Option<&str>, CellValue)]) -> Vec<u8> {
    let mut sheet = Sheet::new("Data");
    for (row, col, formula, value) in cells {
        sheet.set_cell(
            CellRef::new(*row, *col),
            Cell {
                value: value.clone(),
                formula: formula.map(str::to_owned),
                style: None,
            },
        );
    }
    package(vec![sheet])
}

fn at(workbook: &Workbook, row: u32, col: u32) -> CellValue {
    workbook.model().value(SheetId(0), CellRef::new(row, col))
}

fn set(workbook: &mut Workbook, row: u32, col: u32, input: &str) {
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(row, col),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

fn unevaluated(workbook: &Workbook) -> Vec<String> {
    workbook
        .last_calculation()
        .unevaluated_cells
        .iter()
        .map(|cell| cell.address.cell.to_a1())
        .collect()
}

fn hydrated(workbook: &Workbook) -> Workbook {
    use betteroffice_xlsx::{SnapshotBudget, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder};

    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    builder.finish().unwrap().into_parts().0
}

/// A1, B1 and C1 hold no saved value and an uncomputable call reaches each,
/// so all three stay blank and listed, whatever recalculates or restores them.
fn assert_blank_and_listed(workbook: &Workbook, stage: &str) {
    for col in 0..3 {
        assert_eq!(at(workbook, 0, col), CellValue::Empty, "{stage} col {col}");
    }
    assert_eq!(unevaluated(workbook), ["A1", "B1", "C1"], "{stage}");
}

#[test]
fn uncached_gap_cells_stay_blank_through_every_path() {
    let bytes = cells(&[
        (
            0,
            0,
            Some(r#"_xlfn.WEBSERVICE("http://x")"#),
            CellValue::Empty,
        ),
        (
            0,
            1,
            Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#),
            CellValue::Empty,
        ),
        (0, 2, Some("LEN(A1)"), CellValue::Empty),
        (0, 3, None, number(1.0)),
        (0, 4, Some("D1+1"), number(2.0)),
    ]);
    let options = CalculationOptions::default();
    assert_blank_and_listed(&Workbook::open(&bytes).unwrap(), "open");
    let mut workbook = Workbook::open_recalculated(&bytes, options).unwrap();
    assert_blank_and_listed(&workbook, "open recalculated");
    let mut restored = hydrated(&workbook);
    assert_blank_and_listed(&restored, "snapshot");
    set(&mut restored, 0, 3, "7");
    assert_blank_and_listed(&restored, "snapshot edit");
    let mut peer =
        Workbook::open_with_peer_hydration_json(&bytes, &workbook.peer_hydration_json().unwrap())
            .unwrap();
    set(&mut peer, 0, 3, "8");
    assert_blank_and_listed(&peer, "peer edit");
    workbook.recalculate_all(options);
    assert_blank_and_listed(&workbook, "recalculate all");
    set(&mut workbook, 0, 3, "5");
    assert_eq!(at(&workbook, 0, 4), number(6.0));
    assert_blank_and_listed(&workbook, "edit");
    workbook.undo(options).unwrap();
    assert_eq!(at(&workbook, 0, 4), number(2.0));
    assert_blank_and_listed(&workbook, "undo");
    workbook.redo(options).unwrap();
    assert_eq!(at(&workbook, 0, 4), number(6.0));
    assert_blank_and_listed(&workbook, "redo");

    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (mut recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_blank_and_listed(&recovered, "recovery");
    recovered.recalculate_all(options);
    assert_blank_and_listed(&recovered, "recovery recalculate all");
}

/// a restore takes every value it holds as the one to keep: B1 and C1 keep
/// theirs, though the engine would compute C1's from A1.
#[test]
fn a_restore_keeps_every_held_value() {
    let workbook = Workbook::open(&cells(&[
        (
            0,
            0,
            Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#),
            text("live"),
        ),
        (0, 1, Some("LEN(A1)+D1"), number(9.0)),
        (0, 2, Some("LEN(A1)+D1"), number(5.0)),
        (0, 3, None, number(1.0)),
    ]))
    .unwrap();
    let mut restored = hydrated(&workbook);
    assert_eq!(unevaluated(&restored), ["A1", "B1", "C1"]);
    set(&mut restored, 0, 3, "2");
    assert_eq!(at(&restored, 0, 0), text("live"));
    assert_eq!(at(&restored, 0, 1), number(9.0));
    assert_eq!(at(&restored, 0, 2), number(5.0));
    assert_eq!(unevaluated(&restored), ["A1", "B1", "C1"]);
}

/// a read-only open derives the list too, so the first edit after it keeps
/// the cache of a cell reading an uncomputable one.
#[test]
fn a_read_only_open_derives_unevaluated_cells() {
    let mut workbook = Workbook::open_for_read(&package(vec![data_sheet("LEN(A1)*2+D1")])).unwrap();
    assert_reported_through(&workbook, &["A1", "B1"]);
    set(&mut workbook, 0, 3, "5");
    assert_eq!(at(&workbook, 0, 1), number(9.0));
    assert_eq!(at(&workbook, 0, 2), number(6.0));
    assert_eq!(unevaluated(&workbook), ["A1", "B1"]);

    let plain = Workbook::open_for_read(&cells(&[(0, 0, Some("SUM(1,2)"), number(3.0))])).unwrap();
    assert!(unevaluated(&plain).is_empty());
}

fn assert_reported_through(workbook: &Workbook, cells: &[&str]) {
    assert_eq!(unevaluated(workbook), cells);
    let status: Value = serde_json::from_str(&workbook.calculation_status_json().unwrap()).unwrap();
    let listed: Vec<_> = status["unevaluatedCells"]
        .as_array()
        .unwrap()
        .iter()
        .map(|cell| cell["address"].as_str().unwrap().to_owned())
        .collect();
    assert_eq!(listed, cells);
}

/// B1 stores no value, so nothing reads as kept there, yet C1 reads the
/// uncomputable A1 through it and keeps its cache.
#[test]
fn taint_passes_through_a_formula_without_a_stored_value() {
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some(r#"_xlfn.WEBSERVICE("http://x")"#), text("live")),
        (0, 1, Some("LEN(A1)"), CellValue::Empty),
        (0, 2, Some("B1+D1"), number(99.0)),
        (0, 3, None, number(0.0)),
    ]))
    .unwrap();
    assert_eq!(unevaluated(&workbook), ["A1", "B1", "C1"]);
    set(&mut workbook, 0, 3, "1");
    assert_eq!(at(&workbook, 0, 2), number(99.0));
    assert_eq!(unevaluated(&workbook), ["A1", "B1", "C1"]);
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), text("live"));
    assert_eq!(at(&workbook, 0, 1), CellValue::Empty);
    assert_eq!(at(&workbook, 0, 2), number(99.0));
    assert_eq!(unevaluated(&workbook), ["A1", "B1", "C1"]);
}

/// an uncomputable call the file stored no value for still taints the
/// formulas reading it; recalculated, it stays blank.
#[test]
fn an_uncached_uncomputable_call_taints_its_readers() {
    let mut workbook = Workbook::open(&cells(&[
        (
            0,
            0,
            Some(r#"_xlfn.WEBSERVICE("http://x")"#),
            CellValue::Empty,
        ),
        (0, 1, Some(r#"IFERROR(A1,"fallback")&D1"#), text("live1")),
        (0, 3, None, number(1.0)),
    ]))
    .unwrap();
    assert_eq!(unevaluated(&workbook), ["A1", "B1"]);
    set(&mut workbook, 0, 3, "2");
    assert_eq!(at(&workbook, 0, 1), text("live1"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), CellValue::Empty);
    assert_eq!(at(&workbook, 0, 1), text("live1"));
    assert_eq!(unevaluated(&workbook), ["A1", "B1"]);
}

/// the parser lets whitespace stand before a call's `(`, so the derivation
/// finds such a call too.
#[test]
fn whitespace_before_the_paren_still_marks_an_uncomputable_call() {
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some(r#"_xlfn.WEBSERVICE ("http://x")"#), text("live")),
        (0, 1, Some("LEN(A1)+D1"), number(9.0)),
        (0, 3, None, number(0.0)),
    ]))
    .unwrap();
    assert_eq!(unevaluated(&workbook), ["A1", "B1"]);
    set(&mut workbook, 0, 3, "1");
    assert_eq!(at(&workbook, 0, 1), number(9.0));
}

/// no element of the condition picks the branch holding the call, so the
/// formula computes and nothing is reported.
#[test]
fn an_array_if_branch_no_element_picks_is_not_evaluated() {
    let bytes = cells(&[
        (
            0,
            0,
            Some(r#"SUM(IF({TRUE,TRUE},{1,2},_xlfn.WEBSERVICE("http://x")))"#),
            number(999.0),
        ),
        (
            1,
            0,
            Some(r#"SUM(IF({TRUE,FALSE},{1,2},_xlfn.WEBSERVICE("http://x")))"#),
            number(7.0),
        ),
    ]);
    let workbook = Workbook::open(&bytes).unwrap();
    assert_eq!(unevaluated(&workbook), ["A2"]);
    let workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_eq!(at(&workbook, 0, 0), number(3.0));
    assert_eq!(at(&workbook, 1, 0), number(7.0));
    assert_eq!(unevaluated(&workbook), ["A2"]);
}

/// an edit to A1 leaves it blank while the new call reaches it, an undo puts
/// excel's value back, and a later recalculation keeps that value, as it
/// keeps one an op pastes.
#[test]
fn an_undo_restores_excels_value_and_keeps_it() {
    let mut workbook = Workbook::open(&cells(&[(
        0,
        0,
        Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#),
        text("live"),
    )]))
    .unwrap();
    set(&mut workbook, 0, 0, r#"=_xlfn.WEBSERVICE("http://y")"#);
    assert_eq!(at(&workbook, 0, 0), CellValue::Empty);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(at(&workbook, 0, 0), text("live"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), text("live"));
    assert_eq!(unevaluated(&workbook), ["A1"]);

    set(&mut workbook, 0, 0, r#"=_xlfn.WEBSERVICE("http://y")"#);
    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: text("pasted"),
                    formula: Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#.into()),
                    style: None,
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(at(&workbook, 0, 0), text("pasted"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), text("pasted"));
}

/// a moved cell keeps what it holds: B1's saved value moved onto A1 stays,
/// and D1's blank moved to E1 stays blank as its input changes.
#[test]
fn a_moved_gap_cell_keeps_what_it_holds() {
    let mut workbook = Workbook::open(&cells(&[
        (
            0,
            1,
            Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#),
            text("live"),
        ),
        (
            0,
            3,
            Some("IFERROR(_xlfn.WEBSERVICE(C2),C2)"),
            CellValue::Empty,
        ),
        (1, 2, None, text("v1")),
    ]))
    .unwrap();
    let options = CalculationOptions::default();
    workbook.recalculate_all(options);
    assert_eq!(at(&workbook, 0, 3), CellValue::Empty);
    set(&mut workbook, 0, 0, r#"=_xlfn.WEBSERVICE("http://y")"#);
    let range = |a1: &str| CellRange::parse_a1(a1).unwrap();
    workbook
        .move_range(SheetId(0), range("B1"), CellRef::new(0, 0), options)
        .unwrap();
    assert_eq!(at(&workbook, 0, 0), text("live"));
    workbook.recalculate_all(options);
    assert_eq!(at(&workbook, 0, 0), text("live"));

    workbook
        .move_range(SheetId(0), range("D1"), CellRef::new(0, 4), options)
        .unwrap();
    set(&mut workbook, 1, 2, "v2");
    assert_eq!(at(&workbook, 0, 4), CellValue::Empty);
    assert_eq!(unevaluated(&workbook), ["A1", "E1"]);
}

/// deleting a row, a column or the sheet and undoing it brings B2 back as
/// it was, blank, and it stays blank and listed as its input changes.
#[test]
fn an_undone_delete_brings_back_a_blank_gap_cell() {
    let mut data = Sheet::new("Data");
    data.set_cell(
        CellRef::new(1, 1),
        Cell {
            formula: Some("IFERROR(_xlfn.WEBSERVICE(Inputs!A1),Inputs!A1)".into()),
            ..Cell::default()
        },
    );
    let mut inputs = Sheet::new("Inputs");
    inputs.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: text("v1"),
            ..Cell::default()
        },
    );
    let bytes = package(vec![data, inputs]);
    let options = CalculationOptions::default();
    for op in [
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        Op::RemoveSheet { index: 0 },
    ] {
        let mut workbook = Workbook::open_recalculated(&bytes, options).unwrap();
        assert_eq!(at(&workbook, 1, 1), CellValue::Empty, "{op:?}");
        workbook.apply_ops(vec![op.clone()], options).unwrap();
        workbook.undo(options).unwrap();
        assert_eq!(at(&workbook, 1, 1), CellValue::Empty, "{op:?}");
        workbook
            .edit_cell(SheetId(1), CellRef::new(0, 0), "v2", options)
            .unwrap();
        assert_eq!(at(&workbook, 1, 1), CellValue::Empty, "{op:?}");
        assert_eq!(unevaluated(&workbook), ["B2"], "{op:?}");
    }
}

/// an op writing a value over a blank gap cell takes effect even when the
/// value is the one the formula's fallback would give, and later edits to
/// that fallback's input leave it.
#[test]
fn a_value_an_op_writes_over_a_gap_cell_is_kept() {
    let bytes = cells(&[
        (
            0,
            0,
            Some("IFERROR(_xlfn.WEBSERVICE(C1),C1)"),
            CellValue::Empty,
        ),
        (0, 2, None, text("v1")),
    ]);
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open_recalculated(&bytes, options).unwrap();
    assert_eq!(at(&workbook, 0, 0), CellValue::Empty);
    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: text("v1"),
                    formula: Some("IFERROR(_xlfn.WEBSERVICE(C1),C1)".into()),
                    style: None,
                },
            }],
            options,
        )
        .unwrap();
    assert_eq!(at(&workbook, 0, 0), text("v1"));
    set(&mut workbook, 0, 2, "v2");
    assert_eq!(at(&workbook, 0, 0), text("v1"));
    assert_eq!(unevaluated(&workbook), ["A1"]);
}

/// the shared document keeps a formula cell's value while its formula stays,
/// so a local op pasting a value over a gap cell with the same formula, or
/// moving a saved one onto it, leaves every replica holding the same blank.
#[test]
fn collaboration_keeps_every_replica_on_the_same_gap_value() {
    let formula = r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#;
    let mut sheet = Sheet::new("Data");
    for (col, value) in [(0, CellValue::Empty), (1, text("live"))] {
        sheet.set_cell(
            CellRef::new(0, col),
            Cell {
                value,
                formula: Some(formula.into()),
                style: None,
            },
        );
    }
    let model = WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    };
    let options = CalculationOptions::default();
    let mut local = Workbook::from_model_collaborative(model.clone(), 101).unwrap();
    let mut peer = Workbook::from_model_collaborative(model, 102).unwrap();
    peer.apply_update_v1(&local.encode_state_as_update_v1(), options)
        .unwrap();
    local
        .apply_update_v1(&peer.encode_state_as_update_v1(), options)
        .unwrap();
    let sync = |local: &mut Workbook, peer: &mut Workbook| {
        let update = local
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap();
        peer.apply_update_v1(&update, options).unwrap();
    };

    local
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: text("pasted"),
                    formula: Some(formula.into()),
                    style: None,
                },
            }],
            options,
        )
        .unwrap();
    sync(&mut local, &mut peer);
    assert_eq!(at(&local, 0, 0), CellValue::Empty);
    assert_eq!(at(&peer, 0, 0), CellValue::Empty);

    let range = CellRange::parse_a1("B1").unwrap();
    local
        .move_range(SheetId(0), range, CellRef::new(0, 0), options)
        .unwrap();
    sync(&mut local, &mut peer);
    for replica in [&local, &peer] {
        assert_eq!(at(replica, 0, 0), CellValue::Empty);
        assert_eq!(at(replica, 0, 1), CellValue::Empty);
    }
    local.recalculate_all(options);
    peer.recalculate_all(options);
    for replica in [&local, &peer] {
        assert_eq!(at(replica, 0, 0), CellValue::Empty);
    }
    assert_eq!(unevaluated(&local), ["A1"]);
    assert_eq!(unevaluated(&peer), ["A1"]);
}

/// A1 is an array formula recorded over its own cell whose result would
/// fill A2, and E1's uncomputable call keeps it, so B1 reading A2 and C1
/// reading B1 keep their values and are listed however the session opens or
/// restores.
#[test]
fn readers_of_cells_a_kept_spill_would_fill_are_listed_through_every_open() {
    let mut sheet = Sheet::new("Data");
    for (col, formula, value) in [
        (0, Some(r#"IF(E1="x",{10;20},0)"#), CellValue::Empty),
        (1, Some("A2+D1"), number(99.0)),
        (2, Some("B1+1"), number(100.0)),
        (3, None, number(1.0)),
        (4, Some(r#"_xlfn.WEBSERVICE("x")"#), text("x")),
    ] {
        sheet.set_cell(
            CellRef::new(0, col),
            Cell {
                value,
                formula: formula.map(str::to_owned),
                style: None,
            },
        );
    }
    let anchor = CellRef::new(0, 0);
    sheet.set_array_formula(anchor, CellRange::new(anchor, anchor));
    let bytes = package(vec![sheet]);
    let listed = ["A1", "B1", "C1", "E1"];
    let workbook = Workbook::open(&bytes).unwrap();
    assert_eq!(unevaluated(&workbook), listed, "open");
    let read_only = Workbook::open_for_read(&bytes).unwrap();
    assert_eq!(unevaluated(&read_only), listed, "read-only open");
    assert_eq!(unevaluated(&hydrated(&workbook)), listed, "snapshot");
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(unevaluated(&recovered), listed, "recovery");
    let mut peer =
        Workbook::open_with_peer_hydration_json(&bytes, &workbook.peer_hydration_json().unwrap())
            .unwrap();
    set(&mut peer, 0, 3, "8");
    assert_eq!(at(&peer, 0, 1), number(99.0), "peer edit");
    assert_eq!(unevaluated(&peer), listed, "peer edit");
    let recalculated = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_eq!(at(&recalculated, 0, 1), number(99.0), "recalculated");
    assert_eq!(unevaluated(&recalculated), listed, "recalculated");
}
