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
            formula: Some("TODAY()".into()),
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
                    value: text("written"),
                    formula: Some(r#"IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")"#.into()),
                    style: None,
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(at(&workbook, 0, 0), text("written"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), text("written"));
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
                    value: text("written"),
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

/// B1 reads the uncomputable A1 only through `INDIRECT`, so no edge of the
/// graph leads there; what it reads keeps it at Excel's value and listed,
/// whatever opens, recalculates or restores it. C1 reads plain cells the same
/// way and computes.
#[test]
fn a_reader_through_a_computed_reference_keeps_excels_value_through_every_path() {
    let bytes = cells(&[
        (0, 0, Some(r#"_xlfn.WEBSERVICE("http://x")"#), text("live")),
        (0, 1, Some(r#"INDIRECT("A1")&"!""#), text("cached")),
        (0, 2, Some(r#"SUM(INDIRECT("D1:E1"))"#), number(0.0)),
        (0, 3, None, number(1.0)),
        (0, 4, None, number(2.0)),
    ]);
    let check = |workbook: &Workbook, stage: &str, sum: f64| {
        assert_eq!(at(workbook, 0, 0), text("live"), "{stage}");
        assert_eq!(at(workbook, 0, 1), text("cached"), "{stage}");
        assert_eq!(unevaluated(workbook), ["A1", "B1"], "{stage}");
        if sum > 0.0 {
            assert_eq!(at(workbook, 0, 2), number(sum), "{stage}");
        }
    };
    let options = CalculationOptions::default();
    let mut opened = Workbook::open(&bytes).unwrap();
    check(&opened, "open", 0.0);
    set(&mut opened, 0, 3, "4");
    check(&opened, "open edit", 6.0);
    check(
        &Workbook::open_for_read(&bytes).unwrap(),
        "read-only open",
        0.0,
    );
    let mut workbook = Workbook::open_recalculated(&bytes, options).unwrap();
    check(&workbook, "open recalculated", 3.0);
    let mut restored = hydrated(&workbook);
    check(&restored, "snapshot", 0.0);
    set(&mut restored, 0, 4, "5");
    check(&restored, "snapshot edit", 6.0);
    workbook.recalculate_all(options);
    check(&workbook, "recalculate all", 3.0);
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (mut recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    check(&recovered, "recovery", 0.0);
    recovered.recalculate_all(options);
    check(&recovered, "recovery recalculate all", 3.0);
}

fn names(workbook: &Workbook) -> Vec<(String, Vec<String>)> {
    workbook
        .last_calculation()
        .unevaluated_cells
        .iter()
        .map(|cell| (cell.address.cell.to_a1(), cell.functions.clone()))
        .collect()
}

fn named(cells: &[&str], function: &str) -> Vec<(String, Vec<String>)> {
    cells
        .iter()
        .map(|cell| ((*cell).to_owned(), vec![function.to_owned()]))
        .collect()
}

fn name_error() -> CellValue {
    CellValue::Error {
        value: xlsx_model::ErrorValue::Name,
    }
}

/// B1 calls a function the engine does not know under a handler and C1 reads
/// it; the plain formula in E1 keeps working.
fn unknown_source() -> Vec<u8> {
    cells(&[
        (
            0,
            1,
            Some("IFERROR(REPRO_UNKNOWN_FUNCTION(D1),0)"),
            number(42.0),
        ),
        (0, 2, Some("B1*2"), number(5.0)),
        (0, 3, None, number(1.0)),
        (0, 4, Some("D1+1"), number(2.0)),
    ])
}

fn assert_unknown_kept(workbook: &Workbook, stage: &str) {
    assert_eq!(at(workbook, 0, 1), number(42.0), "{stage}");
    assert_eq!(at(workbook, 0, 2), number(5.0), "{stage}");
    assert_eq!(
        names(workbook),
        named(&["B1", "C1"], "REPRO_UNKNOWN_FUNCTION"),
        "{stage}"
    );
}

#[test]
fn an_unknown_call_keeps_excels_value_through_every_path() {
    let bytes = unknown_source();
    let options = CalculationOptions::default();
    assert_unknown_kept(&Workbook::open(&bytes).unwrap(), "open");
    assert_unknown_kept(&Workbook::open_for_read(&bytes).unwrap(), "read-only open");
    let mut workbook = Workbook::open_recalculated(&bytes, options).unwrap();
    assert_unknown_kept(&workbook, "open recalculated");
    assert_unknown_kept(&hydrated(&workbook), "snapshot");
    let peer =
        Workbook::open_with_peer_hydration_json(&bytes, &workbook.peer_hydration_json().unwrap())
            .unwrap();
    assert_unknown_kept(&peer, "peer");
    set(&mut workbook, 0, 3, "9");
    assert_eq!(at(&workbook, 0, 4), number(10.0));
    assert_unknown_kept(&workbook, "edit");
    workbook.undo(options).unwrap();
    assert_unknown_kept(&workbook, "undo");
    workbook.recalculate_all(options);
    assert_unknown_kept(&workbook, "recalculate all");
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (recovered, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_unknown_kept(&recovered, "recovery");
}

/// typing a typo shows what excel shows, in the cell and in what reads it,
/// and the report says so; fixing it clears both.
#[test]
fn a_typed_unknown_call_shows_name_and_its_readers_follow() {
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some("1+1"), number(2.0)),
        (0, 1, Some("A1*2"), number(4.0)),
    ]))
    .unwrap();
    set(&mut workbook, 0, 0, "=SUMM(1,2)");
    assert_eq!(at(&workbook, 0, 0), name_error());
    assert_eq!(at(&workbook, 0, 1), name_error());
    assert_eq!(names(&workbook), named(&["A1", "B1"], "SUMM"));
    set(&mut workbook, 0, 0, "=SUM(1,2)");
    assert_eq!(at(&workbook, 0, 0), number(3.0));
    assert_eq!(at(&workbook, 0, 1), number(6.0));
    assert!(names(&workbook).is_empty());
}

/// a handler typed around the call answers, as in excel, and goes on
/// answering as its input moves.
#[test]
fn a_typed_handler_answers_and_follows_its_input() {
    let mut workbook = Workbook::open(&cells(&[(0, 2, None, number(1.0))])).unwrap();
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),C1)");
    assert_eq!(at(&workbook, 0, 0), number(1.0));
    set(&mut workbook, 0, 2, "2");
    assert_eq!(at(&workbook, 0, 0), number(2.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), number(2.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
}

/// an edit replaces a saved value with the engine's answer; undo brings the
/// saved value back and it stays, even though the graph answered this cell.
#[test]
fn an_undo_restores_a_saved_value_the_engine_had_answered_over() {
    let mut workbook =
        Workbook::open(&cells(&[(0, 0, Some("IFERROR(TYPO(1),0)"), number(42.0))])).unwrap();
    let options = CalculationOptions::default();
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(2),0)");
    assert_eq!(at(&workbook, 0, 0), number(0.0));
    workbook.undo(options).unwrap();
    assert_eq!(at(&workbook, 0, 0), number(42.0));
    workbook.recalculate_all(options);
    assert_eq!(at(&workbook, 0, 0), number(42.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
    workbook.redo(options).unwrap();
    assert_eq!(at(&workbook, 0, 0), number(0.0));
}

/// a sheet rename builds a new graph; the cell goes on answering.
#[test]
fn an_answer_survives_a_rebuilt_graph() {
    let mut workbook = Workbook::open(&cells(&[(0, 2, None, number(1.0))])).unwrap();
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),C1)");
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    set(&mut workbook, 0, 2, "2");
    assert_eq!(at(&workbook, 0, 0), number(2.0));
}

/// a full recalculation builds a new graph too.
#[test]
fn an_answer_survives_a_full_recalculation() {
    let mut workbook = Workbook::open(&cells(&[(0, 2, None, number(1.0))])).unwrap();
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),C1)");
    workbook.recalculate_all(CalculationOptions::default());
    set(&mut workbook, 0, 2, "2");
    assert_eq!(at(&workbook, 0, 0), number(2.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
}

/// an edit batch builds a new graph for each batch.
#[test]
fn an_answer_survives_edit_batches() {
    let mut workbook = Workbook::open(&cells(&[(0, 2, None, number(1.0))])).unwrap();
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),C1)");
    for input in ["2", "3"] {
        let request = json!({
            "expectVersion": workbook.version(),
            "steps": [{
                "op": "setCellInputs",
                "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "C1" } },
                "inputs": [[input]]
            }]
        });
        let result: Value =
            serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap())
                .unwrap();
        assert_eq!(result["applied"], true);
        assert_eq!(
            at(&workbook, 0, 0),
            number(input.parse().unwrap()),
            "{input}"
        );
    }
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
}

fn write_keeping_formula(workbook: &mut Workbook, col: u32, formula: &str, value: CellValue) {
    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, col),
                cell: CellState {
                    value,
                    formula: Some(formula.into()),
                    style: None,
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

/// an op writing a value beside the formula, as a move or a replica does: the
/// record of the answer no longer matches, so the value came from elsewhere
/// and stays, through edits and recalculations.
#[test]
fn a_value_an_op_writes_over_an_answer_is_kept() {
    let formula = "IFERROR(TYPO(1),C1)";
    let mut workbook = Workbook::open(&cells(&[(0, 2, None, number(1.0))])).unwrap();
    set(&mut workbook, 0, 0, &format!("={formula}"));
    assert_eq!(at(&workbook, 0, 0), number(1.0));
    write_keeping_formula(&mut workbook, 0, formula, text("written"));
    assert_eq!(at(&workbook, 0, 0), text("written"));
    set(&mut workbook, 0, 2, "2");
    assert_eq!(at(&workbook, 0, 0), text("written"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), text("written"));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
}

/// a reader of an answer rests on nothing of excel's, so it computes like any
/// formula, and a value an op writes beside its formula lasts no longer than
/// beside any other.
#[test]
fn a_reader_of_an_answer_computes_over_what_an_op_wrote() {
    let formula = "IFERROR(A1,C1)";
    let mut workbook = Workbook::open(&cells(&[(0, 2, None, number(1.0))])).unwrap();
    set(&mut workbook, 0, 0, "=TYPO(1)");
    set(&mut workbook, 0, 1, &format!("={formula}"));
    assert_eq!(at(&workbook, 0, 1), number(1.0));
    set(&mut workbook, 0, 2, "2");
    assert_eq!(at(&workbook, 0, 1), number(2.0));
    write_keeping_formula(&mut workbook, 1, formula, text("written"));
    assert_eq!(at(&workbook, 0, 1), number(2.0));
    set(&mut workbook, 0, 2, "3");
    assert_eq!(at(&workbook, 0, 1), number(3.0));
    assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"));
}

/// A1 holds no value and no recalculation has answered it yet, so nothing
/// tells B1 what it reads: B1 keeps what the file saved until one does.
#[test]
fn a_reader_keeps_its_value_until_its_source_is_answered() {
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some("NOPE(1)"), CellValue::Empty),
        (0, 1, Some("A1*2+C1"), number(5.0)),
        (0, 2, None, number(1.0)),
    ]))
    .unwrap();
    assert_eq!(names(&workbook), named(&["A1", "B1"], "NOPE"));
    set(&mut workbook, 0, 2, "2");
    assert_eq!(at(&workbook, 0, 0), CellValue::Empty);
    assert_eq!(at(&workbook, 0, 1), number(5.0));
    assert_eq!(names(&workbook), named(&["A1", "B1"], "NOPE"));
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(at(&workbook, 0, 0), name_error());
    assert_eq!(at(&workbook, 0, 1), name_error());
}

#[derive(Clone, Copy, Debug)]
enum Overwrite {
    EditCell,
    EditCells,
    OpWithoutFormula,
}

const OVERWRITES: [Overwrite; 3] = [
    Overwrite::EditCell,
    Overwrite::EditCells,
    Overwrite::OpWithoutFormula,
];

fn overwrite(workbook: &mut Workbook, col: u32, how: Overwrite) {
    let options = CalculationOptions::default();
    let cell = CellRef::new(0, col);
    match how {
        Overwrite::EditCell => {
            workbook.edit_cell(SheetId(0), cell, "9", options).unwrap();
        }
        Overwrite::EditCells => {
            let edit = betteroffice_xlsx::CellInput {
                cell,
                input: "9".into(),
            };
            workbook.edit_cells(SheetId(0), &[edit], options).unwrap();
        }
        Overwrite::OpWithoutFormula => {
            workbook
                .apply_ops(
                    vec![Op::SetCell {
                        sheet: SheetId(0),
                        at: cell,
                        cell: CellState {
                            value: number(9.0),
                            formula: None,
                            style: None,
                        },
                    }],
                    options,
                )
                .unwrap();
        }
    }
}

fn formula_at_row(workbook: &Workbook, row: u32, col: u32) -> Option<String> {
    workbook.model().sheets[0]
        .cell(CellRef::new(row, col))
        .and_then(|cell| cell.formula.clone())
}

fn formula_at(workbook: &Workbook, col: u32) -> Option<String> {
    formula_at_row(workbook, 0, col)
}

/// a value the user puts over the cell at `col` replaces the formula and
/// stays through recalculations and an edit to the input in D1. two undos
/// take back that edit and the value, and the formula comes back holding
/// `restored`; an edit to D1 then leaves `after`.
fn assert_overwrite_sticks(
    setup: impl Fn() -> Workbook,
    col: u32,
    restored: CellValue,
    after: CellValue,
) {
    for how in OVERWRITES {
        let mut workbook = setup();
        let options = CalculationOptions::default();
        let formula = formula_at(&workbook, col);
        let listed = names(&workbook);
        let at_col = CellRef::new(0, col).to_a1();
        let others: Vec<_> = listed
            .iter()
            .filter(|(cell, _)| *cell != at_col)
            .cloned()
            .collect();
        let stays = |workbook: &Workbook, stage: &str| {
            assert_eq!(at(workbook, 0, col), number(9.0), "{how:?} {stage}");
            assert_eq!(formula_at(workbook, col), None, "{how:?} {stage}");
            assert_eq!(names(workbook), others, "{how:?} {stage}");
        };
        overwrite(&mut workbook, col, how);
        stays(&workbook, "overwritten");
        workbook.recalculate_all(options);
        stays(&workbook, "recalculated");
        set(&mut workbook, 0, 3, "2");
        stays(&workbook, "input edited");
        workbook.recalculate_all(options);
        stays(&workbook, "recalculated again");

        workbook.undo(options).unwrap();
        workbook.undo(options).unwrap();
        assert_eq!(formula_at(&workbook, col), formula, "{how:?}");
        assert_eq!(at(&workbook, 0, col), restored, "{how:?}");
        assert_eq!(names(&workbook), listed, "{how:?}");
        set(&mut workbook, 0, 3, "5");
        assert_eq!(at(&workbook, 0, col), after, "{how:?}");
        assert_eq!(names(&workbook), listed, "{how:?}");
    }
}

#[test]
fn a_value_replaces_a_cell_holding_an_answer_until_undone() {
    let setup = || {
        let mut workbook = Workbook::open(&cells(&[(0, 3, None, number(1.0))])).unwrap();
        set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),D1)");
        workbook
    };
    assert_overwrite_sticks(setup, 0, number(1.0), number(5.0));
}

#[test]
fn a_value_replaces_a_reader_of_an_answer_until_undone() {
    let setup = || {
        let mut workbook = Workbook::open(&cells(&[(0, 3, None, number(1.0))])).unwrap();
        set(&mut workbook, 0, 0, "=TYPO(1)");
        set(&mut workbook, 0, 1, "=IFERROR(A1,D1)");
        workbook
    };
    assert_overwrite_sticks(setup, 1, number(1.0), number(5.0));
}

#[test]
fn a_value_replaces_a_cell_keeping_excels_value_until_undone() {
    let setup = || {
        Workbook::open(&cells(&[
            (0, 0, Some("IFERROR(TYPO(D1),0)"), number(42.0)),
            (0, 3, None, number(1.0)),
        ]))
        .unwrap()
    };
    assert_overwrite_sticks(setup, 0, number(42.0), number(42.0));
}

#[test]
fn a_value_replaces_a_reader_of_a_kept_value_until_undone() {
    let setup = || {
        Workbook::open(&cells(&[
            (0, 0, Some("TYPO(1)"), number(21.0)),
            (0, 1, Some("A1*2+D1"), number(5.0)),
            (0, 3, None, number(1.0)),
        ]))
        .unwrap()
    };
    assert_overwrite_sticks(setup, 1, number(5.0), number(5.0));
}

/// the batch ends on a formula with no value, so it is a fresh one: the
/// engine answers it, whatever excel saved for the formula it replaced.
#[test]
fn overwriting_and_retyping_in_one_op_batch_counts_as_a_fresh_formula() {
    let formula = "IFERROR(TYPO(1),D1)";
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some(formula), number(42.0)),
        (0, 3, None, number(1.0)),
    ]))
    .unwrap();
    let state = |value, formula: Option<&str>| Op::SetCell {
        sheet: SheetId(0),
        at: CellRef::new(0, 0),
        cell: CellState {
            value,
            formula: formula.map(str::to_owned),
            style: None,
        },
    };
    workbook
        .apply_ops(
            vec![
                state(number(9.0), None),
                state(CellValue::Empty, Some(formula)),
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(at(&workbook, 0, 0), number(1.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
    set(&mut workbook, 0, 3, "2");
    assert_eq!(at(&workbook, 0, 0), number(2.0));
}

/// an edit batch refuses two steps writing one cell, so it can't reach that.
#[test]
fn an_edit_batch_refuses_overwriting_and_retyping_one_cell() {
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some("IFERROR(TYPO(1),D1)"), number(42.0)),
        (0, 3, None, number(1.0)),
    ]))
    .unwrap();
    let step = |input: &str| {
        json!({
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "A1" } },
            "inputs": [[input]]
        })
    };
    let version = workbook.version();
    let request = json!({
        "expectVersion": version,
        "steps": [step("9"), step("=IFERROR(TYPO(1),D1)")]
    });
    let result: Value =
        serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap()).unwrap();
    assert_eq!(result["failure"]["code"], "overlapping-steps", "{result}");
    assert_eq!(workbook.version(), version);
    assert_eq!(at(&workbook, 0, 0), number(42.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
}

/// an answer that happens to equal the value an undo restores is no answer
/// for the formula undone to: that keeps what excel saved.
#[test]
fn a_saved_value_equal_to_an_answer_stays_saved() {
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some("IFERROR(TYPO(2),0)"), number(7.0)),
        (0, 3, None, number(7.0)),
    ]))
    .unwrap();
    let options = CalculationOptions::default();
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),D1)");
    assert_eq!(at(&workbook, 0, 0), number(7.0));
    workbook.undo(options).unwrap();
    assert_eq!(at(&workbook, 0, 0), number(7.0));
    workbook.recalculate_all(options);
    assert_eq!(at(&workbook, 0, 0), number(7.0));
    assert_eq!(names(&workbook), named(&["A1"], "TYPO"));
}

/// Data holds `data`; Sheet2 holds the number 1 in C1.
fn renamable(data: &[(u32, u32, Option<&str>, CellValue)]) -> Workbook {
    let mut first = Sheet::new("Data");
    for (row, col, formula, value) in data {
        first.set_cell(
            CellRef::new(*row, *col),
            Cell {
                value: value.clone(),
                formula: formula.map(str::to_owned),
                style: None,
            },
        );
    }
    let mut second = Sheet::new("Sheet2");
    second.set_cell(
        CellRef::new(0, 2),
        Cell {
            value: number(1.0),
            ..Cell::default()
        },
    );
    Workbook::open(&package(vec![first, second])).unwrap()
}

fn rename(workbook: &mut Workbook, name: &str) {
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(1),
                name: name.into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

fn set_other(workbook: &mut Workbook, input: &str) {
    workbook
        .edit_cell(
            SheetId(1),
            CellRef::new(0, 2),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

/// a rename runs on the same graph with only what it re-resolves as seeds, so
/// A1 is not evaluated while B1's text changes: B1 goes on computing.
#[test]
fn a_rename_keeps_a_reader_of_an_answer_computing() {
    let mut workbook = renamable(&[]);
    set(&mut workbook, 0, 0, "=TYPO(1)");
    set(&mut workbook, 0, 1, "=IFERROR(A1,Sheet2!C1)");
    assert_eq!(at(&workbook, 0, 1), number(1.0));
    rename(&mut workbook, "Renamed");
    assert_eq!(
        formula_at(&workbook, 1).as_deref(),
        Some("IFERROR(A1,Renamed!C1)")
    );
    assert_eq!(at(&workbook, 0, 0), name_error());
    assert_eq!(at(&workbook, 0, 1), number(1.0));
    assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"));
    set_other(&mut workbook, "2");
    assert_eq!(at(&workbook, 0, 0), name_error());
    assert_eq!(at(&workbook, 0, 1), number(2.0));
    assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"));
}

/// a sheet name no sheet held starts resolving, which seeds B1 alone.
#[test]
fn a_rename_seeding_only_a_reader_of_an_answer_computes_it() {
    let mut workbook = renamable(&[]);
    set(&mut workbook, 0, 0, "=TYPO(1)");
    set(&mut workbook, 0, 1, "=IFERROR(A1,Ghost!C1)");
    rename(&mut workbook, "Ghost");
    assert_eq!(at(&workbook, 0, 0), name_error());
    assert_eq!(at(&workbook, 0, 1), number(1.0));
    assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"));
}

#[test]
fn a_rename_leaves_kept_values_alone() {
    let mut workbook = renamable(&[
        (0, 0, Some("TYPO(1)"), number(21.0)),
        (0, 1, Some("IFERROR(A1,Sheet2!C1)"), number(5.0)),
    ]);
    for stage in ["open", "renamed", "edited"] {
        match stage {
            "renamed" => rename(&mut workbook, "Renamed"),
            "edited" => set_other(&mut workbook, "2"),
            _ => {}
        }
        assert_eq!(at(&workbook, 0, 0), number(21.0), "{stage}");
        assert_eq!(at(&workbook, 0, 1), number(5.0), "{stage}");
        assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"), "{stage}");
    }
}

/// the rename rewrites the answering cell, not its reader; the answer stays
/// the rewritten formula's, so both go on following the input.
#[test]
fn a_rename_rewriting_an_answering_cell_keeps_it_answering() {
    let mut workbook = renamable(&[]);
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(Sheet2!C1),Sheet2!C1)");
    set(&mut workbook, 0, 1, "=A1+1");
    assert_eq!(at(&workbook, 0, 1), number(2.0));
    rename(&mut workbook, "Renamed");
    set_other(&mut workbook, "2");
    assert_eq!(at(&workbook, 0, 0), number(2.0));
    assert_eq!(at(&workbook, 0, 1), number(3.0));
    assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"));
}

#[test]
fn a_rename_rewriting_a_bare_unknown_call_keeps_its_error_and_reader() {
    let mut workbook = renamable(&[]);
    set(&mut workbook, 0, 0, "=TYPO(Sheet2!C1)");
    set(&mut workbook, 0, 1, "=A1+1");
    rename(&mut workbook, "Renamed");
    set_other(&mut workbook, "2");
    assert_eq!(
        formula_at(&workbook, 0).as_deref(),
        Some("TYPO(Renamed!C1)")
    );
    assert_eq!(at(&workbook, 0, 0), name_error());
    assert_eq!(at(&workbook, 0, 1), name_error());
    assert_eq!(names(&workbook), named(&["A1", "B1"], "TYPO"));
}

/// A5 answers from the input in D5 and B5 reads it.
fn answering_in_row_five() -> Workbook {
    let mut workbook = Workbook::open(&cells(&[(4, 3, None, number(1.0))])).unwrap();
    set(&mut workbook, 4, 0, "=IFERROR(TYPO(1),D5)");
    set(&mut workbook, 4, 1, "=A5+1");
    workbook
}

/// with the answering cell at `(row, col)`, its reader beside it and the input
/// three columns on, a new input reaches both.
fn assert_follows(workbook: &mut Workbook, row: u32, col: u32, input: f64, stage: &str) {
    set(workbook, row, col + 3, &input.to_string());
    assert_eq!(at(workbook, row, col), number(input), "{stage}");
    assert_eq!(at(workbook, row, col + 1), number(input + 1.0), "{stage}");
    let a = CellRef::new(row, col).to_a1();
    let b = CellRef::new(row, col + 1).to_a1();
    let listed = named(&[a.as_str(), b.as_str()], "TYPO");
    assert_eq!(names(workbook), listed, "{stage}");
}

/// inserting and deleting shift cells and rewrite the formulas reading them;
/// an answer goes with its cell, through undo and redo too.
#[test]
fn a_shift_keeps_an_answer_following_its_input() {
    let sheet = SheetId(0);
    for (op, row, col) in [
        (
            Op::InsertRows {
                sheet,
                at: 1,
                count: 2,
            },
            6,
            0,
        ),
        (
            Op::DeleteRows {
                sheet,
                at: 1,
                count: 2,
            },
            2,
            0,
        ),
        (
            Op::InsertCols {
                sheet,
                at: 0,
                count: 1,
            },
            4,
            1,
        ),
    ] {
        let stage = format!("{op:?}");
        let options = CalculationOptions::default();
        let mut workbook = answering_in_row_five();
        workbook.apply_ops(vec![op], options).unwrap();
        assert_follows(&mut workbook, row, col, 2.0, &stage);
        workbook.undo(options).unwrap();
        workbook.undo(options).unwrap();
        assert_follows(&mut workbook, 4, 0, 3.0, &stage);
        workbook.undo(options).unwrap();
        workbook.redo(options).unwrap();
        assert_follows(&mut workbook, 4, 0, 4.0, &stage);
    }
}

#[test]
fn a_redone_shift_keeps_an_answer_following_its_input() {
    let options = CalculationOptions::default();
    let mut workbook = answering_in_row_five();
    let op = Op::InsertRows {
        sheet: SheetId(0),
        at: 1,
        count: 2,
    };
    workbook.apply_ops(vec![op], options).unwrap();
    workbook.undo(options).unwrap();
    workbook.redo(options).unwrap();
    assert_follows(&mut workbook, 6, 0, 2.0, "redo");
}

/// the row holding the cell goes, and its answer with it: undo brings the
/// formula back beside the value it held, which from then on stands.
#[test]
fn an_answer_goes_with_its_deleted_cell() {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open(&cells(&[(0, 3, None, number(1.0))])).unwrap();
    set(&mut workbook, 4, 0, "=IFERROR(TYPO(1),D1)");
    assert_eq!(at(&workbook, 4, 0), number(1.0));
    let op = Op::DeleteRows {
        sheet: SheetId(0),
        at: 4,
        count: 1,
    };
    workbook.apply_ops(vec![op], options).unwrap();
    assert!(names(&workbook).is_empty());
    workbook.undo(options).unwrap();
    assert_eq!(names(&workbook), named(&["A5"], "TYPO"));
    set(&mut workbook, 0, 3, "2");
    assert_eq!(at(&workbook, 4, 0), number(1.0));
}

/// A6 holds what excel saved, with the text and value of the answer in A5. it
/// moves up as A5 goes, and A5's answer must not take it for itself.
#[test]
fn a_deleted_answer_leaves_nothing_for_the_cell_moving_up() {
    let formula = "IFERROR(TYPO(1),$D$1)";
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open(&cells(&[
        (0, 3, None, number(1.0)),
        (4, 0, Some(formula), CellValue::Empty),
        (5, 0, Some(formula), number(1.0)),
    ]))
    .unwrap();
    workbook.recalculate_all(options);
    let op = Op::DeleteRows {
        sheet: SheetId(0),
        at: 4,
        count: 1,
    };
    workbook.apply_ops(vec![op], options).unwrap();
    set(&mut workbook, 0, 3, "2");
    assert_eq!(at(&workbook, 4, 0), number(1.0));
}

/// A3 answers; A1 moves over it with the same text and value, but what A1
/// holds excel saved, and A3's answer must not take it for itself.
#[test]
fn a_cell_moved_over_an_answer_keeps_what_excel_saved() {
    let formula = "IFERROR(TYPO(1),$D$1)";
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open(&cells(&[
        (0, 0, Some(formula), number(1.0)),
        (0, 3, None, number(1.0)),
        (2, 0, Some(formula), CellValue::Empty),
    ]))
    .unwrap();
    workbook.recalculate_all(options);
    let source = CellRange::parse_a1("A1").unwrap();
    workbook
        .move_range(SheetId(0), source, CellRef::new(2, 0), options)
        .unwrap();
    set(&mut workbook, 0, 3, "2");
    assert_eq!(at(&workbook, 2, 0), number(1.0));
}

/// the move's source and target overlap, so A6 is both.
#[test]
fn an_overlapping_move_keeps_each_answer_with_its_cell() {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open(&cells(&[(0, 3, None, number(1.0))])).unwrap();
    set(&mut workbook, 4, 0, "=IFERROR(TYPO(1),$D$1)");
    set(&mut workbook, 5, 0, "=IFERROR(TYPO(2),$D$1*2)");
    let source = CellRange::parse_a1("A5:A6").unwrap();
    workbook
        .move_range(SheetId(0), source, CellRef::new(5, 0), options)
        .unwrap();
    set(&mut workbook, 0, 3, "2");
    assert_eq!(at(&workbook, 5, 0), number(2.0));
    assert_eq!(at(&workbook, 6, 0), number(4.0));
}

#[test]
fn a_move_keeps_an_answer_following_its_input() {
    let options = CalculationOptions::default();
    let mut workbook = answering_in_row_five();
    let source = CellRange::parse_a1("A5:B5").unwrap();
    workbook
        .move_range(SheetId(0), source, CellRef::new(7, 0), options)
        .unwrap();
    set(&mut workbook, 4, 3, "2");
    assert_eq!(at(&workbook, 7, 0), number(2.0));
    assert_eq!(at(&workbook, 7, 1), number(3.0));
    assert_eq!(names(&workbook), named(&["A8", "B8"], "TYPO"));
    workbook.undo(options).unwrap();
    workbook.undo(options).unwrap();
    assert_follows(&mut workbook, 4, 0, 3.0, "undone");
}

/// a move of what the formula reads rewrites its text where it stands.
#[test]
fn moving_an_input_keeps_the_answer_reading_it() {
    let options = CalculationOptions::default();
    let mut workbook = answering_in_row_five();
    let source = CellRange::parse_a1("D5").unwrap();
    workbook
        .move_range(SheetId(0), source, CellRef::new(4, 5), options)
        .unwrap();
    assert_eq!(
        formula_at_row(&workbook, 4, 0).as_deref(),
        Some("IFERROR(TYPO(1),F5)")
    );
    set(&mut workbook, 4, 5, "2");
    assert_eq!(at(&workbook, 4, 0), number(2.0));
    assert_eq!(at(&workbook, 4, 1), number(3.0));
}

/// answers are held by sheet position, which these ops change.
#[test]
fn sheet_ops_keep_an_answer_on_its_sheet() {
    let options = CalculationOptions::default();
    let mut workbook = renamable(&[(0, 2, None, number(1.0))]);
    set(&mut workbook, 0, 0, "=IFERROR(TYPO(1),C1)");
    workbook
        .apply_ops(vec![Op::MoveSheet { from: 0, to: 1 }], options)
        .unwrap();
    workbook
        .edit_cell(SheetId(1), CellRef::new(0, 2), "2", options)
        .unwrap();
    assert_eq!(
        workbook.model().value(SheetId(1), CellRef::new(0, 0)),
        number(2.0)
    );
    workbook.undo(options).unwrap();
    workbook.undo(options).unwrap();
    set(&mut workbook, 0, 2, "3");
    assert_eq!(at(&workbook, 0, 0), number(3.0));
}

/// the insert reaches the answer on every sheet: on the first it moves it, and
/// on the others it leaves it where it stands.
#[test]
fn an_insert_rebases_answers_across_many_sheets() {
    let sheets = (0..200).map(|index| {
        let mut sheet = Sheet::new(format!("S{index}"));
        for (col, formula, value) in [
            (0, Some("IFERROR(TYPO(1),$C$2)"), CellValue::Empty),
            (2, None, number(1.0)),
        ] {
            sheet.set_cell(
                CellRef::new(1, col),
                Cell {
                    value,
                    formula: formula.map(str::to_owned),
                    style: None,
                },
            );
        }
        sheet
    });
    let options = CalculationOptions::default();
    let bytes = package(sheets.collect());
    let mut workbook = Workbook::open_recalculated(&bytes, options).unwrap();
    let op = Op::InsertRows {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    };
    workbook.apply_ops(vec![op], options).unwrap();
    for (sheet, row) in [(SheetId(0), 2), (SheetId(150), 1)] {
        workbook
            .edit_cell(sheet, CellRef::new(row, 2), "2", options)
            .unwrap();
        let answer = workbook.model().value(sheet, CellRef::new(row, 0));
        assert_eq!(answer, number(2.0), "{sheet:?}");
    }
}

/// an edited save keeps the formula as stored and excel's value beside it,
/// and asks excel to calculate on open, as for an uncomputable call.
#[test]
fn an_edited_save_keeps_an_unknown_formula_and_its_value() {
    let bytes = cells(&[
        (0, 0, Some("IFERROR(_xlfn.FUTURE(D1),0)"), number(42.0)),
        (0, 1, Some("[1]!Foo(D1)"), number(7.0)),
        (0, 3, None, number(1.0)),
    ]);
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert_eq!(
        names(&workbook),
        [
            ("A1".to_owned(), vec!["FUTURE".to_owned()]),
            ("B1".to_owned(), vec!["[1]!Foo".to_owned()]),
        ]
    );
    assert_eq!(workbook.save().unwrap(), bytes);
    set(&mut workbook, 0, 3, "2");
    let saved = workbook.save().unwrap();
    assert!(workbook_xml(&saved).contains(r#"fullCalcOnLoad="1""#));
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(at(&reopened, 0, 0), number(42.0));
    assert_eq!(at(&reopened, 0, 1), number(7.0));
    let model = reopened.model();
    let formula = |col| {
        model.sheets[0]
            .cell(CellRef::new(0, col))
            .and_then(|cell| cell.formula.clone())
    };
    assert_eq!(formula(0).as_deref(), Some("IFERROR(_xlfn.FUTURE(D1),0)"));
    assert_eq!(formula(1).as_deref(), Some("[1]!Foo(D1)"));
}

#[test]
fn text_that_only_looks_like_a_call_lists_nothing() {
    let bytes = cells(&[
        (0, 0, Some("LOG10 (100)&\"TYPO(1)\""), text("2TYPO(1)")),
        (0, 1, Some("_xlfn.LET(_xlpm.x,1,_xlpm.x)"), number(1.0)),
    ]);
    let workbook = Workbook::open(&bytes).unwrap();
    assert!(names(&workbook).is_empty());
    assert_eq!(at(&workbook, 0, 0), text("2TYPO(1)"));
}
