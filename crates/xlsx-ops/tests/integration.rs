//! public-api integration: serde wire round-trips and a full edit/undo cycle.

use xlsx_model::{CellProvider, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook};
use xlsx_ops::{CellState, Op, Provenance, Transaction, UndoStack, apply};

fn r(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

#[test]
fn transaction_json_round_trip() {
    let tx = Transaction::proposal(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: r("B2"),
                cell: CellState {
                    value: CellValue::Text { value: "hi".into() },
                    formula: Some("A1&\"i\"".into()),
                    style: Some(3),
                },
            },
            Op::MergeCells {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:C1").unwrap(),
            },
            Op::InsertRows {
                sheet: SheetId(0),
                at: 4,
                count: 2,
            },
        ],
        Provenance::Agent {
            id: "agent-7".into(),
        },
    );

    let json = serde_json::to_string(&tx).unwrap();
    let back: Transaction = serde_json::from_str(&json).unwrap();
    assert_eq!(tx, back);

    let value: serde_json::Value = serde_json::from_str(&json).unwrap();
    assert_eq!(value["proposed"], true);
    assert_eq!(value["author"]["kind"], "agent");
    assert_eq!(value["author"]["id"], "agent-7");
    assert_eq!(value["ops"][0]["type"], "setCell");
    assert_eq!(value["ops"][2]["type"], "insertRows");
}

#[test]
fn full_workbook_round_trip_via_undo_stack() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    wb.sheet_mut(SheetId(0)).unwrap().set_cell(
        r("A1"),
        xlsx_model::Cell {
            value: CellValue::Number { value: 1.0 },
            ..Default::default()
        },
    );
    let baseline = wb.clone();

    let tx = Transaction::new(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: r("A2"),
                cell: CellState {
                    value: CellValue::Number { value: 2.0 },
                    ..Default::default()
                },
            },
            Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::MergeCells {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B1").unwrap(),
            },
        ],
        Provenance::User,
    );

    let mut stack = UndoStack::new();
    stack.commit(&mut wb, &tx).unwrap();
    assert_ne!(wb.sheets[0].used_range(), baseline.sheets[0].used_range());

    stack.undo(&mut wb).unwrap();
    assert_eq!(wb.sheets[0].used_range(), baseline.sheets[0].used_range());
    assert_eq!(
        wb.value(SheetId(0), r("A1")),
        CellValue::Number { value: 1.0 }
    );
    assert!(wb.sheets[0].merges.is_empty());
}

#[test]
fn apply_returns_replayable_inverse() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));

    let op = Op::SetCell {
        sheet: SheetId(0),
        at: r("D4"),
        cell: CellState {
            value: CellValue::Bool { value: true },
            ..Default::default()
        },
    };
    let inverse = apply(&mut wb, &op).unwrap();
    assert_eq!(
        wb.value(SheetId(0), r("D4")),
        CellValue::Bool { value: true }
    );
    for iop in &inverse.0 {
        apply(&mut wb, iop).unwrap();
    }
    assert_eq!(wb.value(SheetId(0), r("D4")), CellValue::Empty);
}
