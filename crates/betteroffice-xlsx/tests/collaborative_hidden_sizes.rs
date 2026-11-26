use betteroffice_xlsx::{
    CalculationOptions, CellRef, CellState, CellValue, EditOperation, EditRequest, EditStep, Op,
    RangeAddress, RangeTarget, Sheet, SheetId, Workbook, WorkbookModel,
};
use yrs::Update;
use yrs::updates::{decoder::Decode, encoder::Encode};

const DATA: SheetId = SheetId(0);
type Sizes = (Option<f64>, Option<f64>, Option<f64>);
const AUTHORED: Sizes = (Some(25.0), Some(30.0), Some(18.0));
const DEFAULT: Sizes = (None, None, None);
const HIDDEN: Sizes = (Some(0.0), Some(0.0), Some(0.0));

fn source() -> Vec<u8> {
    source_with_worksheet(concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<cols><col min="2" max="2" width="18" customWidth="1" hidden="1"/></cols>"#,
        r#"<sheetData><row r="1" ht="25" customHeight="1" hidden="1"><c r="A1"><v>7</v></c></row>"#,
        r#"<row r="2" ht="30" customHeight="1" hidden="1"><c r="B2"><v>2</v></c></row>"#,
        r#"<row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>"#,
    ))
}

fn source_with_worksheet(xml: &str) -> Vec<u8> {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..Default::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = xml.as_bytes().to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn stripped_model(source: &[u8]) -> WorkbookModel {
    let mut model = Workbook::open(source).unwrap().into_model();
    for sheet in &mut model.sheets {
        sheet.hidden_row_heights.clear();
        sheet.hidden_col_widths.clear();
    }
    model
}

fn edit(workbook: &mut Workbook) {
    workbook
        .edit_cell(DATA, CellRef::new(2, 0), "4", CalculationOptions::default())
        .unwrap();
}

fn apply(workbook: &mut Workbook, ops: Vec<Op>) {
    workbook
        .apply_ops(ops, CalculationOptions::default())
        .unwrap();
}

fn row_visibility(row: u32, hidden: bool) -> Op {
    Op::SetRowVisibility {
        sheet: DATA,
        row,
        hidden,
    }
}

fn row_height(row: u32, height: f64) -> Op {
    Op::SetRowHeight {
        sheet: DATA,
        row,
        height: Some(height),
    }
}

fn col_visibility(hidden: bool) -> Op {
    Op::SetColVisibility {
        sheet: DATA,
        col: 1,
        hidden,
    }
}

fn col_width(width: f64) -> Op {
    Op::SetColWidth {
        sheet: DATA,
        col: 1,
        width: Some(width),
    }
}

fn unhide(workbook: &mut Workbook) {
    apply(
        workbook,
        vec![
            row_visibility(0, false),
            row_visibility(1, false),
            col_visibility(false),
        ],
    );
}

fn diff(from: &Workbook, to: &Workbook) -> Vec<u8> {
    from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap()
}

fn receive(workbook: &mut Workbook, update: &[u8]) {
    workbook
        .apply_update_v1(update, CalculationOptions::default())
        .unwrap();
}

fn reopen(workbook: &Workbook) -> Workbook {
    Workbook::open(&workbook.save().unwrap()).unwrap()
}

fn sizes(workbook: &Workbook) -> Sizes {
    let sheet = workbook.sheet(DATA).unwrap();
    (
        sheet.row_heights.get(&0).copied(),
        sheet.row_heights.get(&1).copied(),
        sheet.col_widths.get(&1).copied(),
    )
}

fn hidden_sizes(workbook: &Workbook) -> Sizes {
    let sheet = workbook.sheet(DATA).unwrap();
    (
        sheet.hidden_row_heights.get(&0).copied(),
        sheet.hidden_row_heights.get(&1).copied(),
        sheet.hidden_col_widths.get(&1).copied(),
    )
}

fn assert_saved_hidden(workbook: &Workbook) {
    let saved = reopen(workbook);
    assert_eq!(
        (hidden_sizes(workbook), sizes(&saved), hidden_sizes(&saved)),
        (AUTHORED, HIDDEN, AUTHORED),
    );
}

fn visible_sizes(workbook: &Workbook) -> (Sizes, Sizes, Sizes, Sizes) {
    let saved = reopen(workbook);
    (
        sizes(workbook),
        hidden_sizes(workbook),
        sizes(&saved),
        hidden_sizes(&saved),
    )
}

fn assert_visible(workbook: &Workbook) {
    assert_eq!(
        visible_sizes(workbook),
        (AUTHORED, DEFAULT, AUTHORED, DEFAULT),
    );
}

#[test]
fn unrelated_cell_edit_keeps_authored_sizes_on_both_replicas_and_through_history() {
    let source = source();
    let mut left = Workbook::open_collaborative(&source, 7).unwrap();
    let mut right = Workbook::open_collaborative(&source, 8).unwrap();
    assert_eq!(hidden_sizes(&left), AUTHORED);
    edit(&mut left);
    let update = diff(&left, &right);
    receive(&mut right, &update);
    assert_saved_hidden(&left);
    assert_saved_hidden(&right);
    let mut from_snapshot = Workbook::open_collaborative(&source, 9).unwrap();
    receive(&mut from_snapshot, &left.encode_state_as_update_v1());
    assert_saved_hidden(&from_snapshot);
    left.undo(CalculationOptions::default()).unwrap();
    assert_saved_hidden(&left);
    left.redo(CalculationOptions::default()).unwrap();
    assert_saved_hidden(&left);
}

#[test]
fn unhide_replicates_authored_sizes_and_undo_restores_hidden_state() {
    let source = source();
    let mut left = Workbook::open_collaborative(&source, 7).unwrap();
    let mut right = Workbook::open_collaborative(&source, 8).unwrap();
    let before = left.encode_state_as_update_v1();
    unhide(&mut left);
    let update = diff(&left, &right);
    receive(&mut right, &update);
    assert_visible(&left);
    assert_visible(&right);

    left.undo(CalculationOptions::default()).unwrap();
    let mut before_peer = Workbook::from_model_collaborative(stripped_model(&source), 9).unwrap();
    let mut undone_peer = Workbook::from_model_collaborative(stripped_model(&source), 10).unwrap();
    receive(&mut before_peer, &before);
    receive(&mut undone_peer, &left.encode_state_as_update_v1());
    assert_eq!(
        (undone_peer.model(), hidden_sizes(&left)),
        (before_peer.model(), AUTHORED),
    );
    assert_saved_hidden(&left);
    left.redo(CalculationOptions::default()).unwrap();
    assert_visible(&left);
}

#[cfg(feature = "unstable-track-visibility")]
#[test]
fn range_unhide_restores_authored_and_default_sizes_on_both_replicas() {
    use std::collections::BTreeMap;

    let source = source_with_worksheet(concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<cols><col min="1" max="1" width="18" customWidth="1" hidden="1"/>"#,
        r#"<col min="2" max="2" hidden="1"/>"#,
        r#"<col min="3" max="3" width="24" customWidth="1" hidden="1"/></cols>"#,
        r#"<sheetData><row r="1" ht="25" customHeight="1" hidden="1"></row>"#,
        r#"<row r="2" hidden="1"></row>"#,
        r#"<row r="3" ht="30" customHeight="1" hidden="1"></row></sheetData></worksheet>"#,
    ));
    let mut left = Workbook::open_collaborative(&source, 7).unwrap();
    let mut right = Workbook::from_model_collaborative(stripped_model(&source), 8).unwrap();
    for workbook in [&left, &right] {
        let sheet = workbook.sheet(DATA).unwrap();
        for track in 0..3 {
            assert!(sheet.row_hidden(track));
            assert!(sheet.col_hidden(track));
        }
    }
    apply(
        &mut left,
        vec![
            Op::SetRowsVisibility {
                sheet: DATA,
                row: 0,
                count: 3,
                hidden: false,
            },
            Op::SetColsVisibility {
                sheet: DATA,
                col: 0,
                count: 3,
                hidden: false,
            },
        ],
    );
    let update = diff(&left, &right);
    receive(&mut right, &update);
    for workbook in [&left, &right] {
        let sheet = workbook.sheet(DATA).unwrap();
        assert_eq!(sheet.row_heights, BTreeMap::from([(0, 25.0), (2, 30.0)]));
        assert_eq!(sheet.col_widths, BTreeMap::from([(0, 18.0), (2, 24.0)]));
    }
}

#[test]
fn undo_of_a_default_row_resize_leaves_no_size_for_a_later_peer_hide() {
    let source = source();
    let mut left = Workbook::open_collaborative(&source, 7).unwrap();
    let mut right = Workbook::open_collaborative(&source, 8).unwrap();
    apply(&mut left, vec![row_height(2, 40.0)]);
    assert!(left.undo(CalculationOptions::default()).unwrap().applied);
    assert_eq!(left.sheet(DATA).unwrap().row_heights.get(&2), None);
    apply(&mut right, vec![row_visibility(2, true)]);
    let update = diff(&right, &left);
    receive(&mut left, &update);
    let sheet = left.sheet(DATA).unwrap();
    assert!(sheet.row_hidden(2));
    assert_eq!(sheet.hidden_row_heights.get(&2), None);
}

#[test]
fn received_undo_remembers_the_last_visible_size_seen() {
    let source = source();
    let mut left = Workbook::open_collaborative(&source, 7).unwrap();
    let mut right = Workbook::open_collaborative(&source, 8).unwrap();
    assert_eq!(
        (hidden_sizes(&left).0, hidden_sizes(&right).0),
        (Some(25.0), Some(25.0)),
    );
    apply(
        &mut left,
        vec![row_visibility(0, false), row_height(0, 40.0)],
    );
    let update = diff(&left, &right);
    receive(&mut right, &update);
    assert_eq!(sizes(&right).0, Some(40.0));
    left.undo(CalculationOptions::default()).unwrap();
    let update = diff(&left, &right);
    receive(&mut right, &update);
    assert_eq!(sizes(&left), HIDDEN);
    assert_eq!(
        (hidden_sizes(&left).0, hidden_sizes(&right).0),
        (Some(40.0), Some(40.0)),
    );
    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(sizes(&right), HIDDEN);
}

#[test]
fn replicas_without_local_sizes_receive_authored_unhides_and_send_default_unhides() {
    let source = source();
    for new_peer_unhides in [true, false] {
        let mut current = Workbook::open_collaborative(&source, 7).unwrap();
        let mut released = Workbook::from_model_collaborative(stripped_model(&source), 8).unwrap();
        let known_at_open = hidden_sizes(&current);
        let expected = if new_peer_unhides {
            unhide(&mut current);
            let update = diff(&current, &released);
            receive(&mut released, &update);
            (AUTHORED, DEFAULT, AUTHORED, DEFAULT)
        } else {
            unhide(&mut released);
            let update = diff(&released, &current);
            receive(&mut current, &update);
            (DEFAULT, DEFAULT, DEFAULT, DEFAULT)
        };
        assert_eq!(
            (
                known_at_open,
                visible_sizes(&current),
                visible_sizes(&released)
            ),
            (AUTHORED, expected, expected),
        );
    }
}

#[test]
fn concurrent_unhide_converges_with_resize_or_rehide_on_both_axes() {
    let source = source();
    let reveal = || vec![row_visibility(0, false), col_visibility(false)];
    for (left_id, right_id) in [(7, 8), (8, 7)] {
        for rehide in [false, true] {
            for left_first in [false, true] {
                let mut left = Workbook::open_collaborative(&source, left_id).unwrap();
                let mut right = Workbook::open_collaborative(&source, right_id).unwrap();
                let other = |workbook: &mut Workbook| {
                    if rehide {
                        apply(workbook, reveal());
                        apply(
                            workbook,
                            vec![row_visibility(0, true), col_visibility(true)],
                        );
                    } else {
                        apply(workbook, vec![row_height(0, 40.0), col_width(28.0)]);
                    }
                };
                if left_first {
                    apply(&mut left, reveal());
                    other(&mut right);
                } else {
                    other(&mut right);
                    apply(&mut left, reveal());
                }
                assert_eq!(sizes(&left), (Some(25.0), Some(0.0), Some(18.0)));
                let to_left = diff(&right, &left);
                let to_right = diff(&left, &right);
                if left_first {
                    receive(&mut left, &to_left);
                    receive(&mut right, &to_right);
                } else {
                    receive(&mut right, &to_right);
                    receive(&mut left, &to_left);
                }
                let left = left.sheet(DATA).unwrap();
                let right = right.sheet(DATA).unwrap();
                assert_eq!(left.row_heights, right.row_heights);
                assert_eq!(left.col_widths, right.col_widths);
                let visible = |sheet: &Sheet| {
                    (
                        (!sheet.row_hidden(0)).then(|| sheet.row_heights[&0]),
                        (!sheet.col_hidden(1)).then(|| sheet.col_widths[&1]),
                    )
                };
                let (row, col) = visible(left);
                assert_eq!((row, col), visible(right));
                if rehide {
                    assert!([None, Some(25.0)].contains(&row));
                    assert!([None, Some(18.0)].contains(&col));
                } else {
                    assert!([Some(25.0), Some(40.0)].contains(&row));
                    assert!([Some(18.0), Some(28.0)].contains(&col));
                }
            }
        }
    }
}

#[test]
fn remote_hide_remembers_the_size_of_a_previously_visible_row() {
    let source = source();
    let mut left = Workbook::open_collaborative(&source, 7).unwrap();
    let mut right = Workbook::open_collaborative(&source, 8).unwrap();
    for op in [row_height(2, 22.0), row_visibility(2, true)] {
        apply(&mut right, vec![op]);
        let update = diff(&right, &left);
        receive(&mut left, &update);
    }
    let saved = reopen(&left);
    let sheet = saved.sheet(DATA).unwrap();
    assert_eq!(
        (
            sheet.row_hidden(2),
            sheet.hidden_row_heights.get(&2).copied()
        ),
        (true, Some(22.0)),
    );
}

#[test]
fn retried_remote_hide_remembers_and_republishes_the_predecessor_height() {
    let source = source();
    let mut delayed = Workbook::open_collaborative(&source, 7).unwrap();
    let mut ready = Workbook::open_collaborative(&source, 8).unwrap();
    let mut target = Workbook::open_collaborative(&source, 9).unwrap();
    apply(&mut delayed, vec![row_height(2, 22.0)]);
    let predecessor = diff(&delayed, &target);
    let before_hide = delayed.encode_state_vector_v1();
    apply(&mut delayed, vec![row_visibility(2, true)]);
    let hide = delayed.encode_diff_v1(&before_hide).unwrap();
    edit(&mut ready);
    let merged = Update::merge_updates([
        Update::decode_v1(&hide).unwrap(),
        Update::decode_v1(&diff(&ready, &target)).unwrap(),
    ])
    .encode_v1();

    receive(&mut target, &merged);
    assert_eq!(target.cell(DATA, CellRef::new(2, 0)).unwrap().input, "4");
    assert!(!target.sheet(DATA).unwrap().row_hidden(2));

    receive(&mut target, &predecessor);
    let sheet = target.sheet(DATA).unwrap();
    assert!(sheet.row_hidden(2));
    assert_eq!(sheet.hidden_row_heights.get(&2), Some(&22.0));

    apply(&mut target, vec![row_visibility(2, false)]);
    assert_eq!(
        target.sheet(DATA).unwrap().row_heights.get(&2),
        Some(&22.0),
    );
    let update = diff(&target, &delayed);
    receive(&mut delayed, &update);
    assert_eq!(
        delayed.sheet(DATA).unwrap().row_heights.get(&2),
        Some(&22.0),
    );
}

#[test]
fn hidden_row_resize_save_keeps_the_source_until_hidden_only_changes_patch() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    apply(&mut workbook, vec![row_height(0, 40.0)]);
    apply(&mut workbook, vec![row_visibility(0, true)]);
    let resized = (Some(40.0), Some(30.0), Some(18.0));
    assert_eq!(hidden_sizes(&workbook), resized);
    assert_eq!(hidden_sizes(&reopen(&workbook)), AUTHORED);

    edit(&mut workbook);
    let saved = reopen(&workbook);
    assert_eq!((sizes(&saved), hidden_sizes(&saved)), (HIDDEN, resized));
}

#[test]
fn hidden_size_only_batch_keeps_the_source_until_hidden_only_changes_patch() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    let version = workbook.version();
    let before = workbook.encode_state_as_update_v1();
    let result = workbook
        .apply_ops(
            vec![row_height(0, 40.0), row_visibility(0, true)],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        (
            result.applied,
            workbook.can_undo(),
            workbook.encode_state_as_update_v1(),
            hidden_sizes(&workbook),
        ),
        (true, false, before, (Some(40.0), Some(30.0), Some(18.0))),
    );
    assert_ne!(workbook.version(), version);
    assert_eq!(hidden_sizes(&reopen(&workbook)), AUTHORED);
    apply(&mut workbook, vec![row_visibility(0, false)]);
    assert_eq!(sizes(&workbook).0, Some(40.0));
}

#[test]
fn no_op_reset_then_hide_forgets_the_local_size_without_committing() {
    let mut workbook = Workbook::open_collaborative(&source(), 7).unwrap();
    let options = CalculationOptions::default();
    let version = workbook.version();
    let history = workbook.history_state();
    let state = workbook.encode_state_as_update_v1();
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    assert_eq!(hidden_sizes(&workbook), AUTHORED);
    let result = workbook
        .apply_ops(
            vec![
                Op::SetRowHeight {
                    sheet: DATA,
                    row: 0,
                    height: None,
                },
                row_visibility(0, true),
            ],
            options,
        )
        .unwrap();
    assert!(!result.applied);
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.history_state(), history);
    assert_eq!(workbook.encode_state_as_update_v1(), state);
    assert_eq!(workbook.recovery_checkpoint(1, None).unwrap(), checkpoint);
    assert_eq!(sizes(&workbook), HIDDEN);
    assert_eq!(hidden_sizes(&workbook), (None, Some(30.0), Some(18.0)));
    apply(&mut workbook, vec![row_visibility(0, false)]);
    assert_eq!(sizes(&workbook).0, None);
}

#[test]
fn reset_then_hide_save_keeps_the_source_until_hidden_only_changes_patch() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    assert_eq!(workbook.save().unwrap(), source);
    let version = workbook.version();
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let result = workbook
        .apply_ops(
            vec![
                Op::SetRowHeight {
                    sheet: DATA,
                    row: 0,
                    height: None,
                },
                row_visibility(0, true),
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(!result.applied);
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.recovery_checkpoint(1, None).unwrap(), checkpoint);
    let cleared = (None, Some(30.0), Some(18.0));
    assert_eq!(hidden_sizes(&workbook), cleared);
    assert_eq!(workbook.save().unwrap(), source);

    edit(&mut workbook);
    let saved = reopen(&workbook);
    assert_eq!((sizes(&saved), hidden_sizes(&saved)), (HIDDEN, cleared));
}

#[test]
fn local_reset_then_hide_keeps_the_staged_absence_on_both_axes() {
    for height in [25.0, 40.0] {
        let mut model = Workbook::open(&source()).unwrap().into_model();
        model.sheets[0].hidden_row_heights.remove(&0);
        model.sheets[0].hidden_col_widths.remove(&1);
        model.sheets[0].row_heights.insert(0, height);
        model.sheets[0].col_widths.insert(1, height);
        let mut workbook = Workbook::from_model_collaborative(model, 7).unwrap();
        apply(
            &mut workbook,
            vec![
                Op::SetRowHeight {
                    sheet: DATA,
                    row: 0,
                    height: None,
                },
                row_visibility(0, true),
                Op::SetColWidth {
                    sheet: DATA,
                    col: 1,
                    width: None,
                },
                col_visibility(true),
            ],
        );
        assert_eq!(sizes(&workbook), HIDDEN);
        assert_eq!(
            (hidden_sizes(&workbook).0, hidden_sizes(&workbook).2),
            (None, None),
        );
        apply(
            &mut workbook,
            vec![row_visibility(0, false), col_visibility(false)],
        );
        assert_eq!((sizes(&workbook).0, sizes(&workbook).2), (None, None));
    }
}

#[test]
fn cell_edit_preserves_hidden_sizes_on_an_untouched_sheet_with_many_overrides() {
    let mut dimensions = Sheet::new("Dimensions");
    dimensions
        .row_heights
        .extend((0..5_000).map(|row| (row, 0.0)));
    dimensions
        .hidden_row_heights
        .extend((0..5_000).map(|row| (row, 25.0)));
    dimensions.col_widths.insert(1, 0.0);
    dimensions.hidden_col_widths.insert(1, 18.0);
    let mut workbook = Workbook::from_model_collaborative(
        WorkbookModel {
            sheets: vec![Sheet::new("Data"), dimensions.clone()],
            ..Default::default()
        },
        7,
    )
    .unwrap();
    apply(
        &mut workbook,
        vec![Op::SetCell {
            sheet: DATA,
            at: CellRef::new(0, 0),
            cell: CellState {
                value: CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        }],
    );
    let sheet = workbook.sheet(SheetId(1)).unwrap();
    assert_eq!(sheet.hidden_row_heights, dimensions.hidden_row_heights);
    assert_eq!(sheet.hidden_col_widths, dimensions.hidden_col_widths);
}

#[test]
fn validating_applying_and_receiving_cell_batches_keep_untouched_size_maps() {
    let mut dimensions = Sheet::new("Dimensions");
    dimensions
        .row_heights
        .extend((0..5_000).map(|row| (row, 0.0)));
    dimensions
        .hidden_row_heights
        .extend((0..5_000).map(|row| (row, 25.0)));
    dimensions.col_widths.insert(1, 0.0);
    dimensions.hidden_col_widths.insert(1, 18.0);
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data"), dimensions.clone()],
        ..Default::default()
    };
    let mut left = Workbook::from_model_collaborative(model.clone(), 7).unwrap();
    let mut right = Workbook::from_model_collaborative(model, 8).unwrap();
    let request = EditRequest {
        expect_version: left.version(),
        source: Default::default(),
        history: Default::default(),
        calculation: Default::default(),
        steps: vec![EditStep::new(EditOperation::SetCellInputs {
            target: RangeTarget {
                sheet_id: left.sheet_ids()[0].clone(),
                range: RangeAddress::A1 { a1: "A1".into() },
            },
            inputs: vec![vec!["4".into()]],
        })],
    };
    let row_entry = |workbook: &Workbook| {
        workbook
            .sheet(SheetId(1))
            .unwrap()
            .hidden_row_heights
            .get(&4_999)
            .map(std::ptr::from_ref)
    };
    let left_entry = row_entry(&left);
    let right_entry = row_entry(&right);
    assert!(left_entry.is_some());
    assert!(right_entry.is_some());
    assert!(left.validate_edits(&request).unwrap().unwrap().would_apply);
    assert_eq!(row_entry(&left), left_entry);
    assert!(left.apply_edits(&request).unwrap().unwrap().applied);
    let update = diff(&left, &right);
    receive(&mut right, &update);
    assert_eq!(row_entry(&left), left_entry);
    assert_eq!(row_entry(&right), right_entry);
    for workbook in [&left, &right] {
        let sheet = workbook.sheet(SheetId(1)).unwrap();
        assert_eq!(sheet.hidden_row_heights, dimensions.hidden_row_heights);
        assert_eq!(sheet.hidden_col_widths, dimensions.hidden_col_widths);
    }
}

#[test]
fn local_hidden_sizes_leave_shared_state_bytes_unchanged() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    let mut stripped = Workbook::from_model_collaborative(stripped_model(&source), 7).unwrap();
    edit(&mut workbook);
    edit(&mut stripped);
    let saved = reopen(&workbook);
    assert_eq!(
        (workbook.encode_state_as_update_v1(), hidden_sizes(&saved)),
        (stripped.encode_state_as_update_v1(), AUTHORED),
    );
}

#[test]
fn collaborative_checkpoint_excludes_local_sizes_and_recovers_them_from_source() {
    let source = source();
    let model = Workbook::open(&source).unwrap().into_model();
    let mut sized = Workbook::from_model_collaborative(model, 7).unwrap();
    let mut stripped = Workbook::from_model_collaborative(stripped_model(&source), 7).unwrap();
    stripped
        .adopt_peer_version(sized.version().as_str())
        .unwrap();
    let mut source_backed = Workbook::open_collaborative(&source, 7).unwrap();
    for workbook in [&mut sized, &mut stripped, &mut source_backed] {
        edit(workbook);
    }
    let checkpoint = source_backed.recovery_checkpoint(1, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(
        (
            sized.recovery_checkpoint(1, None).unwrap(),
            hidden_sizes(&restored)
        ),
        (stripped.recovery_checkpoint(1, None).unwrap(), AUTHORED),
    );
    assert_saved_hidden(&restored);
}

#[test]
fn collaborative_checkpoint_recovers_a_hidden_size_from_the_visible_source_row() {
    let mut source = Workbook::open(&source()).unwrap();
    apply(&mut source, vec![row_height(2, 22.0)]);
    let source = source.save().unwrap();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    apply(&mut workbook, vec![row_visibility(2, true)]);
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    for book in [&restored, &reopen(&restored)] {
        let sheet = book.sheet(DATA).unwrap();
        assert!(sheet.row_hidden(2));
        assert_eq!(sheet.hidden_row_heights.get(&2), Some(&22.0));
    }
}
