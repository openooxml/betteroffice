//! public-api integration: serde wire round-trips and a full edit/undo cycle.

use xlsx_model::{CellProvider, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook};
use xlsx_ops::{
    BorderLineStyle, BorderPatch, BorderPreset, CellState, HorizontalAlignment,
    NumberFormatMutation, Op, Provenance, StylePatch, TextWrapping, Transaction, UndoStack, apply,
};

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

#[test]
fn range_style_and_number_format_are_undoable() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    let tx = Transaction::new(
        vec![
            Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                patch: StylePatch {
                    bold: Some(true),
                    text_color: Some("#123456".into()),
                    ..StylePatch::default()
                },
            },
            Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                format: NumberFormatMutation::Percent,
            },
            Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                format: NumberFormatMutation::IncreaseDecimal,
            },
            Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                patch: StylePatch {
                    fill_color: Some("#abcdef".into()),
                    horizontal_alignment: Some(HorizontalAlignment::Center),
                    text_wrapping: Some(TextWrapping::Wrap),
                    border: Some(BorderPatch {
                        preset: Some(BorderPreset::Outer),
                        style: Some(BorderLineStyle::Double),
                        color: Some("#654321".into()),
                    }),
                    ..StylePatch::default()
                },
            },
        ],
        Provenance::User,
    );
    let mut stack = UndoStack::new();
    stack.commit(&mut wb, &tx).unwrap();
    for address in ["A1", "B1", "A2", "B2"] {
        let cell = wb.sheets[0].cell(r(address)).unwrap();
        let format = wb.styles.cell_format(cell.style);
        assert!(format.font.bold);
        assert_eq!(
            format.number_format,
            xlsx_model::NumberFormat::Custom {
                pattern: "0.000%".into()
            }
        );
        assert_eq!(format.alignment.h, Some(xlsx_model::HAlign::Center));
        assert!(format.alignment.wrap_text);
        assert_eq!(
            format.fill,
            xlsx_model::Fill::Solid(xlsx_model::Color::Rgb("#abcdef".into()))
        );
    }
    let top_left = wb
        .styles
        .cell_format(wb.sheets[0].cell(r("A1")).unwrap().style);
    assert_eq!(
        top_left.border.top.as_ref().unwrap().style,
        xlsx_model::BorderStyle::Double
    );
    assert!(top_left.border.left.is_some());
    assert!(top_left.border.right.is_none());
    assert!(top_left.border.bottom.is_none());
    stack.undo(&mut wb).unwrap();
    assert!(wb.sheets[0].iter_cells().next().is_none());
    stack.redo(&mut wb).unwrap();
    assert!(
        wb.styles
            .cell_format(wb.sheets[0].cell(r("A1")).unwrap().style)
            .font
            .bold
    );
}
