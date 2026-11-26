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

fn unhide(workbook: &mut Workbook) {
    apply(
        workbook,
        vec![
            row_visibility(0, false),
            row_visibility(1, false),
            Op::SetColVisibility {
                sheet: DATA,
                col: 1,
                hidden: false,
            },
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
fn concurrent_unhide_converges_with_resize_or_rehide() {
    let source = source();
    for rehide in [false, true] {
        let mut left = Workbook::open_collaborative(&source, 7).unwrap();
        let mut right = Workbook::open_collaborative(&source, 8).unwrap();
        apply(&mut left, vec![row_visibility(0, false)]);
        let unhidden = sizes(&left).0;
        let other_intent = if rehide {
            apply(&mut right, vec![row_visibility(0, false)]);
            apply(&mut right, vec![row_visibility(0, true)]);
            0.0
        } else {
            apply(&mut right, vec![row_height(0, 40.0)]);
            40.0
        };
        let to_left = diff(&right, &left);
        let to_right = diff(&left, &right);
        receive(&mut left, &to_left);
        receive(&mut right, &to_right);
        let row_size = sizes(&left).0;
        assert_eq!(
            (
                left.model(),
                unhidden,
                row_size == Some(25.0) || row_size == Some(other_intent),
                left.sheet(DATA).unwrap().row_hidden(0),
            ),
            (right.model(), Some(25.0), true, row_size == Some(0.0)),
        );
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
fn hidden_size_only_batch_creates_no_shared_update_or_history() {
    let source = source();
    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
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
        (false, false, before, AUTHORED),
    );
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
