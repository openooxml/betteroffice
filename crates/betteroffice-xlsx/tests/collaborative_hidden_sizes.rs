use betteroffice_xlsx::{CalculationOptions, CellRef, Op, Sheet, SheetId, Workbook, WorkbookModel};

const DATA: SheetId = SheetId(0);
type Sizes = (Option<f64>, Option<f64>, Option<f64>);
const AUTHORED: Sizes = (Some(25.0), Some(30.0), Some(18.0));
const DEFAULT: Sizes = (None, None, None);
const HIDDEN: Sizes = (Some(0.0), Some(0.0), Some(0.0));

fn source() -> Vec<u8> {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..Default::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<cols><col min="2" max="2" width="18" customWidth="1" hidden="1"/></cols>"#,
        r#"<sheetData><row r="1" ht="25" customHeight="1" hidden="1"><c r="A1"><v>7</v></c></row>"#,
        r#"<row r="2" ht="30" customHeight="1" hidden="1"><c r="B2"><v>2</v></c></row>"#,
        r#"<row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>"#,
    )
    .as_bytes()
    .to_vec();
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
        .edit_cell(
            DATA,
            CellRef::new(2, 0),
            "4",
            CalculationOptions::default(),
        )
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

#[test]
fn undo_of_unhide_and_resize_restores_pre_step_hidden_sizes() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    apply(
        &mut workbook,
        vec![row_visibility(0, false), row_height(0, 40.0)],
    );
    assert_eq!(
        (sizes(&workbook).0, hidden_sizes(&workbook).0),
        (Some(40.0), None),
    );
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_saved_hidden(&workbook);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(
        (sizes(&workbook).0, hidden_sizes(&workbook).0),
        (Some(40.0), None),
    );
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
            (known_at_open, visible_sizes(&current), visible_sizes(&released)),
            (AUTHORED, expected, expected),
        );
    }
}

#[test]
fn released_unhide_then_undo_restores_remembered_hidden_sizes() {
    let source = source();
    let mut current = Workbook::open_collaborative(&source, 7).unwrap();
    let mut released = Workbook::from_model_collaborative(stripped_model(&source), 8).unwrap();
    unhide(&mut released);
    let update = diff(&released, &current);
    receive(&mut current, &update);
    assert_eq!((sizes(&current), hidden_sizes(&current)), (DEFAULT, DEFAULT));
    released.undo(CalculationOptions::default()).unwrap();
    let update = diff(&released, &current);
    receive(&mut current, &update);
    assert_saved_hidden(&current);
}

#[test]
fn visible_resize_supersedes_a_remembered_hidden_size() {
    let source = source();
    let mut current = Workbook::open_collaborative(&source, 7).unwrap();
    let mut released = Workbook::from_model_collaborative(stripped_model(&source), 8).unwrap();
    for ops in [
        vec![row_visibility(0, false), col_visibility(false)],
        vec![row_height(0, 40.0), col_width(28.0)],
        vec![
            Op::SetRowHeight {
                sheet: DATA,
                row: 0,
                height: None,
            },
            Op::SetColWidth {
                sheet: DATA,
                col: 1,
                width: None,
            },
        ],
        vec![row_visibility(0, true), col_visibility(true)],
    ] {
        apply(&mut released, ops);
        let update = diff(&released, &current);
        receive(&mut current, &update);
    }
    assert_eq!(hidden_sizes(&current), (Some(40.0), Some(30.0), Some(28.0)));
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
                        apply(workbook, vec![row_visibility(0, true), col_visibility(true)]);
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
        (sheet.row_hidden(2), sheet.hidden_row_heights.get(&2).copied()),
        (true, Some(22.0)),
    );
}

#[test]
fn saving_rewrites_a_worksheet_when_only_its_hidden_sizes_differ() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    apply(&mut workbook, vec![row_height(0, 40.0)]);
    apply(&mut workbook, vec![row_visibility(0, true)]);
    let saved = reopen(&workbook);
    assert_eq!(
        (sizes(&saved), hidden_sizes(&saved)),
        (HIDDEN, (Some(40.0), Some(30.0), Some(18.0))),
    );
}

#[test]
fn hidden_size_only_batch_commits_without_a_shared_update_and_keeps_the_resize() {
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
    assert_eq!(hidden_sizes(&reopen(&workbook)).0, Some(40.0));
    apply(&mut workbook, vec![row_visibility(0, false)]);
    assert_eq!(sizes(&workbook).0, Some(40.0));
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
    stripped.adopt_peer_version(sized.version().as_str()).unwrap();
    let mut source_backed = Workbook::open_collaborative(&source, 7).unwrap();
    for workbook in [&mut sized, &mut stripped, &mut source_backed] {
        edit(workbook);
    }
    let checkpoint = source_backed.recovery_checkpoint(1, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(
        (sized.recovery_checkpoint(1, None).unwrap(), hidden_sizes(&restored)),
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
