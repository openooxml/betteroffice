use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRef, CellState, CellValue, Error, Op, Sheet, SheetId,
    Workbook, WorkbookModel,
};
use yrs::Update;
use yrs::updates::decoder::Decode;

const SEEDED_CELLS: [(&str, &str); 6] = [
    ("B3", "row 3"),
    ("B4", "row 4"),
    ("B5", "row 5"),
    ("C2", "column C"),
    ("D2", "column D"),
    ("E2", "column E"),
];

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Data");
    for (address, input) in SEEDED_CELLS {
        sheet.set_cell(
            cell(address),
            Cell {
                value: CellValue::Text {
                    value: input.into(),
                },
                ..Default::default()
            },
        );
    }
    model.sheets.push(sheet);
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

fn assert_empty_undo(workbook: &mut Workbook) {
    let model = workbook.model().clone();
    let state = workbook.encode_state_as_update_v1();
    let history = workbook.history_state();
    assert!(!history.can_undo);
    assert_eq!(history.undo_depth, 0);
    assert!(
        !workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.model(), &model);
    assert_eq!(workbook.encode_state_as_update_v1(), state);
    assert_eq!(workbook.history_state(), history);
}

fn assert_refusal_policy_preserves_peer_edits(
    ops: Vec<Op>,
    inside: &str,
    outside: &str,
    undo_before_exchange: bool,
) {
    let bytes = source();
    let options = CalculationOptions::default();
    for left_first in [true, false] {
        let mut left = Workbook::open_collaborative(&bytes, 101).unwrap();
        let mut right = Workbook::open_collaborative(&bytes, 202).unwrap();
        let model = left.model().clone();
        let state = left.encode_state_as_update_v1();
        let history = left.history_state();

        assert!(matches!(
            left.apply_ops(ops.clone(), options),
            Err(Error::CollaborativeStructureOperation)
        ));
        assert_eq!(left.model(), &model);
        assert_eq!(left.encode_state_as_update_v1(), state);
        assert_eq!(left.history_state(), history);

        assert!(
            right
                .edit_cells(
                    SheetId(0),
                    &[
                        CellInput {
                            cell: cell(outside),
                            input: "peer outside deletion".into(),
                        },
                        CellInput {
                            cell: cell(inside),
                            input: "peer inside deletion".into(),
                        },
                    ],
                    options,
                )
                .unwrap()
                .applied
        );
        let right_history = right.history_state();
        assert!(right_history.can_undo);

        if undo_before_exchange {
            assert_empty_undo(&mut left);
        }

        let left_update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let right_update = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        if left_first {
            right.apply_update_v1(&left_update, options).unwrap();
            left.apply_update_v1(&right_update, options).unwrap();
        } else {
            left.apply_update_v1(&right_update, options).unwrap();
            right.apply_update_v1(&left_update, options).unwrap();
        }

        assert_eq!(left.model(), right.model());
        assert_empty_undo(&mut left);
        let update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        right.apply_update_v1(&update, options).unwrap();

        assert_eq!(left.model(), right.model());
        assert_eq!(left.history_state(), history);
        assert_eq!(right.history_state(), right_history);
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert!(
            Update::decode_v1(&left.encode_state_as_update_v1()).unwrap()
                == Update::decode_v1(&right.encode_state_as_update_v1()).unwrap()
        );
        for workbook in [&left, &right] {
            for (address, seeded) in SEEDED_CELLS {
                let expected = if address == outside {
                    "peer outside deletion"
                } else if address == inside {
                    "peer inside deletion"
                } else {
                    seeded
                };
                assert_eq!(
                    workbook
                        .sheet(SheetId(0))
                        .unwrap()
                        .cell(cell(address))
                        .unwrap()
                        .value,
                    CellValue::Text {
                        value: expected.into(),
                    },
                    "{ops:?}, left_first={left_first}, undo_before_exchange={undo_before_exchange}, {address}"
                );
            }
        }
    }
}

#[test]
fn row_deletion_pins_refusal_policy_with_peer_edits_and_empty_undo() {
    assert_refusal_policy_preserves_peer_edits(
        vec![Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        }],
        "B3",
        "B5",
        false,
    );
}

#[test]
fn column_deletion_pins_refusal_policy_with_peer_edits_and_empty_undo() {
    assert_refusal_policy_preserves_peer_edits(
        vec![Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        }],
        "C2",
        "E2",
        false,
    );
}

#[test]
fn row_deletion_pins_refusal_policy_with_peer_edits_delivered_after_empty_undo() {
    assert_refusal_policy_preserves_peer_edits(
        vec![Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        }],
        "B3",
        "B5",
        true,
    );
}

#[test]
fn column_deletion_pins_refusal_policy_with_peer_edits_delivered_after_empty_undo() {
    assert_refusal_policy_preserves_peer_edits(
        vec![Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        }],
        "C2",
        "E2",
        true,
    );
}

#[test]
fn set_cell_then_delete_rows_pins_atomic_refusal_policy() {
    assert_refusal_policy_preserves_peer_edits(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: cell("B4"),
                cell: CellState {
                    value: CellValue::Text {
                        value: "refused batch".into(),
                    },
                    ..Default::default()
                },
            },
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            },
        ],
        "B3",
        "B5",
        false,
    );
}

#[test]
fn delete_cols_then_set_cell_pins_atomic_refusal_policy() {
    assert_refusal_policy_preserves_peer_edits(
        vec![
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            },
            Op::SetCell {
                sheet: SheetId(0),
                at: cell("D2"),
                cell: CellState {
                    value: CellValue::Text {
                        value: "refused batch".into(),
                    },
                    ..Default::default()
                },
            },
        ],
        "C2",
        "E2",
        false,
    );
}
