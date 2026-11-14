use super::*;

use std::panic::{AssertUnwindSafe, catch_unwind};

use xlsx_model::{Cell, Table};

const NOW_SERIAL: f64 = 45_000.25;
const RAND_SEED: u32 = 17;
const SEEDS: std::ops::RangeInclusive<u64> = 1..=4;

fn options() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(NOW_SERIAL),
    }
}

fn at(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

fn number(value: f64) -> Cell {
    Cell {
        value: CellValue::Number { value },
        ..Cell::default()
    }
}

fn text(value: &str) -> Cell {
    Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Cell::default()
    }
}

fn formula(source: &str) -> Cell {
    Cell {
        formula: Some(source.into()),
        ..Cell::default()
    }
}

fn defined_name(name: &str, formula: &str) -> DefinedName {
    DefinedName {
        name: name.into(),
        formula: formula.into(),
        local_sheet: None,
        hidden: false,
    }
}

/// a seeded workbook of formulas, names, a cycle and volatiles; `extras` adds a table and a spill.
fn generated(seed: u64, extras: bool) -> WorkbookModel {
    let mut state = seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1;
    let mut draw = |bound: u64| -> u64 {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state % bound
    };
    let mut model = WorkbookModel::default();
    let factor = 1 + draw(3);
    for name in ["Sheet1", "Data Two"] {
        let mut sheet = Sheet::new(name);
        let rows = 6 + draw(6) as u32;
        for row in 1..=rows {
            let value = (1 + draw(50)) as f64;
            let chain = match row {
                1 => "B1".to_owned(),
                _ => format!("B{row}+D{}", row - 1),
            };
            let scaled = format!("A{row}*{factor}+1");
            sheet.set_cell(at(&format!("A{row}")), number(value));
            sheet.set_cell(at(&format!("B{row}")), formula(&scaled));
            sheet.set_cell(at(&format!("C{row}")), formula("SUM($A$1:$A$4)+$A$1"));
            sheet.set_cell(at(&format!("D{row}")), formula(&chain));
        }
        model.sheets.push(sheet);
    }
    let mut calc = Sheet::new("Calc");
    for row in 1..=4 + draw(4) as u32 {
        let cross = format!("'Data Two'!B{row}+Sheet1!$C$1");
        calc.set_cell(at(&format!("A{row}")), formula(&cross));
    }
    calc.set_cell(at("B1"), formula("SUM(Sheet1!A1:B6)"));
    calc.set_cell(at("B2"), formula("SUM(Total)*Rate"));
    calc.set_cell(at("C1"), formula("RAND()"));
    calc.set_cell(at("C2"), formula("NOW()-NOW()+A1"));
    calc.set_cell(at("C3"), formula("INDIRECT(\"Sheet1!A2\")+1"));
    calc.set_cell(at("F1"), formula("G1+1"));
    calc.set_cell(at("G1"), formula("F1+1"));
    model.sheets.push(calc);
    let total = defined_name("Total", "Sheet1!$A$1:$A$4");
    let rate = defined_name("Rate", "'Data Two'!$A$2");
    model.defined_names.extend([total, rate]);
    if !extras {
        return model;
    }
    let mut table = Sheet::new("Table");
    table.set_cell(at("A1"), text("Item"));
    table.set_cell(at("B1"), text("Amount"));
    let last = 4 + draw(3) as u32;
    for row in 2..=last {
        let amount = (1 + draw(20)) as f64;
        table.set_cell(at(&format!("A{row}")), text(&format!("item {row}")));
        table.set_cell(at(&format!("B{row}")), number(amount));
    }
    table.set_cell(at("D1"), formula("SUM(Sales[Amount])"));
    model.tables.push(Table {
        name: "Sales".into(),
        sheet: SheetId(3),
        range: CellRange::new(at("A1"), at(&format!("B{last}"))),
        header_rows: 1,
        totals_rows: 0,
        columns: vec!["Item".into(), "Amount".into()],
    });
    model.sheets.push(table);
    let mut spill = Sheet::new("Spill");
    spill.set_cell(at("B1"), number((2 + draw(3)) as f64));
    spill.set_cell(at("A2"), formula("_xlfn.SEQUENCE(B1)"));
    spill.set_array_formula(at("A2"), CellRange::new(at("A2"), at("A2")));
    spill.set_cell(at("C1"), formula("SUM(A2:A9)"));
    model.sheets.push(spill);
    model
}

/// a workbook over `model`, recalculated once.
fn workbook(model: WorkbookModel) -> Workbook {
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook.set_rand_seed(Some(RAND_SEED));
    workbook.recalculate_all(options());
    workbook
}

fn edit(workbook: &mut Workbook, sheet: u32, a1: &str, input: &str) {
    let sheet = SheetId(sheet);
    workbook.edit_cell(sheet, at(a1), input, options()).unwrap();
}

/// the next full recalculation keeps the graph exactly when `keeps`, and agrees with a fresh one.
fn assert_next_full_recalculation(workbook: &mut Workbook, keeps: bool, label: &str) {
    let (kept, rebuilt) = workbook.graph_reuse_counts();
    let difference = workbook.recalculate_all_beside_fresh(options());
    let expected = if keeps {
        (kept + 1, rebuilt)
    } else {
        (kept, rebuilt + 1)
    };
    let counts = workbook.graph_reuse_counts();
    assert_eq!(counts, expected, "{label}: (keeps, rebuilds)");
    assert_eq!(difference, None, "{label}");
}

fn check(extras: bool, keeps: bool, change: impl Fn(&mut Workbook)) {
    for seed in SEEDS {
        let mut workbook = workbook(generated(seed, extras));
        change(&mut workbook);
        let label = format!("seed {seed}");
        assert_next_full_recalculation(&mut workbook, keeps, &label);
    }
}

fn keeps_after(op: Op) {
    check(false, true, |workbook| {
        workbook.apply_ops(vec![op.clone()], options()).unwrap();
    });
}

fn rebuilds_after(op: Op) {
    check(false, false, |workbook| {
        workbook.apply_ops(vec![op.clone()], options()).unwrap();
    });
}

#[test]
fn two_full_recalculations_in_a_row_keep_the_graph() {
    for seed in SEEDS {
        let mut workbook = Workbook::from_model(generated(seed, true)).unwrap();
        workbook.set_rand_seed(Some(RAND_SEED));
        for run in ["first", "second"] {
            let label = format!("seed {seed}, {run}");
            assert_next_full_recalculation(&mut workbook, true, &label);
        }
    }
}

#[test]
fn a_value_edit_keeps_the_graph() {
    check(true, true, |workbook| edit(workbook, 0, "A2", "1000"));
}

#[test]
fn a_spill_source_change_keeps_the_graph() {
    check(true, true, |workbook| edit(workbook, 4, "B1", "6"));
}

#[test]
fn a_formula_set_rebuilds_the_graph() {
    check(true, false, |workbook| edit(workbook, 0, "B3", "=A3*10"));
}

#[test]
fn a_formula_clear_rebuilds_the_graph() {
    check(true, false, |workbook| edit(workbook, 0, "B3", ""));
}

#[test]
fn an_array_formula_rebuilds_the_graph_once() {
    for seed in SEEDS {
        let mut workbook = workbook(generated(seed, true));
        edit(&mut workbook, 2, "E1", "=SEQUENCE(3)");
        for (keeps, run) in [(false, "entered"), (true, "again")] {
            let label = format!("seed {seed}, {run}");
            assert_next_full_recalculation(&mut workbook, keeps, &label);
        }
    }
}

#[test]
fn a_sheet_rename_rebuilds_the_graph() {
    rebuilds_after(Op::RenameSheet {
        sheet: SheetId(1),
        name: "Renamed Two".into(),
    });
}

#[test]
fn adding_a_defined_name_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::SetDefinedName {
        name: "Later".into(),
        refers_to: "Sheet1!$A$1".into(),
        scope: None,
        hidden: None,
    });
}

#[test]
fn changing_a_defined_name_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::SetDefinedName {
        name: "Rate".into(),
        refers_to: "Sheet1!$A$3".into(),
        scope: None,
        hidden: None,
    });
}

#[test]
fn removing_a_defined_name_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::DeleteDefinedName {
        name: "Total".into(),
        scope: None,
    });
}

#[test]
fn a_table_resize_swapped_in_rebuilds_the_graph() {
    check(true, false, |workbook| {
        let mut model = workbook.model().clone();
        model.tables[0].range.end.row -= 1;
        workbook.install_model(model).unwrap();
    });
}

#[test]
fn adding_a_sheet_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::AddSheet {
        index: 1,
        name: "Extra".into(),
    });
}

#[test]
fn deleting_a_sheet_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::RemoveSheet { index: 1 });
}

#[test]
fn moving_a_sheet_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::MoveSheet { from: 2, to: 0 });
}

#[test]
fn inserting_a_row_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::InsertRows {
        sheet: SheetId(0),
        at: 2,
        count: 1,
    });
}

#[test]
fn deleting_a_row_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::DeleteRows {
        sheet: SheetId(0),
        at: 1,
        count: 1,
    });
}

#[test]
fn inserting_a_column_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::InsertCols {
        sheet: SheetId(0),
        at: 1,
        count: 1,
    });
}

#[test]
fn deleting_a_column_keeps_the_graph_its_recalculation_rebuilt() {
    keeps_after(Op::DeleteCols {
        sheet: SheetId(0),
        at: 1,
        count: 1,
    });
}

#[test]
fn an_undo_keeps_the_graph_its_recalculation_rebuilt() {
    check(false, true, |workbook| {
        edit(workbook, 0, "A2", "1000");
        workbook.undo(options()).unwrap();
    });
}

#[test]
fn a_redo_keeps_the_graph_its_recalculation_rebuilt() {
    check(false, true, |workbook| {
        edit(workbook, 0, "A2", "1000");
        workbook.undo(options()).unwrap();
        workbook.redo(options()).unwrap();
    });
}

#[test]
fn a_model_swap_rebuilds_the_graph() {
    check(true, false, |workbook| {
        let mut model = workbook.model().clone();
        model.sheets[0].set_cell(at("A2"), number(1000.0));
        workbook.install_model(model).unwrap();
    });
}

#[test]
fn a_remote_update_keeps_the_graph_its_install_built() {
    for seed in SEEDS {
        let source = Workbook::from_model(generated(seed, false)).unwrap();
        let bytes = source.save().unwrap();
        let mut local = Workbook::open_collaborative(&bytes, 1_001).unwrap();
        let mut remote = Workbook::open_collaborative(&bytes, 2_002).unwrap();
        remote.set_rand_seed(Some(RAND_SEED));
        remote.recalculate_all(options());
        edit(&mut local, 0, "A2", "1000");
        let state = remote.encode_state_vector_v1();
        let update = local.encode_diff_v1(&state).unwrap();
        remote.apply_update_v1(&update, options()).unwrap();
        let label = format!("seed {seed}");
        assert_next_full_recalculation(&mut remote, true, &label);
    }
}

/// marks the graph current after formula edits too until dropped.
struct RestampingFormulaEdits;

impl RestampingFormulaEdits {
    fn on() -> Self {
        RESTAMP_FORMULA_EDITS.set(true);
        Self
    }
}

impl Drop for RestampingFormulaEdits {
    fn drop(&mut self) {
        RESTAMP_FORMULA_EDITS.set(false);
    }
}

#[test]
fn a_graph_wrongly_kept_after_a_formula_edit_differs_from_a_fresh_build() {
    for seed in SEEDS {
        let mut workbook = workbook(generated(seed, true));
        let restamping = RestampingFormulaEdits::on();
        edit(&mut workbook, 0, "A1", "=B2+1");
        let outcome = catch_unwind(AssertUnwindSafe(|| {
            workbook.recalculate_all_beside_fresh(options())
        }));
        drop(restamping);
        let reported = match outcome {
            Ok(difference) => difference,
            Err(panic) => panic.downcast_ref::<String>().cloned(),
        };
        assert!(reported.is_some(), "seed {seed}: no difference reported");
    }
}

/// edits that enter formulas, recalculated once with a rebuild, before the value edits under test.
fn check_after_formulas(formulas: &[(u32, &str, &str)], values: &[(u32, &str, &str)]) {
    for seed in SEEDS {
        let mut workbook = workbook(generated(seed, true));
        for (sheet, a1, input) in formulas {
            edit(&mut workbook, *sheet, a1, input);
        }
        let label = format!("seed {seed}");
        assert_next_full_recalculation(&mut workbook, false, &format!("{label}, formulas"));
        for (sheet, a1, input) in values {
            edit(&mut workbook, *sheet, a1, input);
        }
        assert_next_full_recalculation(&mut workbook, true, &format!("{label}, values"));
    }
}

#[test]
fn a_value_over_a_formula_rebuilds_the_graph() {
    check(true, false, |workbook| edit(workbook, 0, "B2", "5"));
}

#[test]
fn a_table_header_edit_keeps_the_graph() {
    check(true, true, |workbook| {
        edit(workbook, 3, "B1", "Total Amount")
    });
}

#[test]
fn a_value_blocking_a_spill_keeps_the_graph() {
    check(true, true, |workbook| {
        edit(workbook, 4, "A8", "9");
        edit(workbook, 4, "B1", "7");
    });
}

#[test]
fn clearing_a_spill_blocker_keeps_the_graph() {
    for seed in SEEDS {
        let mut workbook = workbook(generated(seed, true));
        edit(&mut workbook, 4, "A8", "9");
        edit(&mut workbook, 4, "B1", "7");
        assert_next_full_recalculation(&mut workbook, true, &format!("seed {seed}, blocked"));
        edit(&mut workbook, 4, "A8", "");
        assert_next_full_recalculation(&mut workbook, true, &format!("seed {seed}, cleared"));
    }
}

#[test]
fn a_computed_reference_target_edit_keeps_the_graph() {
    let formulas = [
        (2, "H1", "Sheet1!A2"),
        (2, "H2", "=INDIRECT(H1)+1"),
        (2, "H3", "=OFFSET(Sheet1!$A$1,H4,0)"),
        (2, "H4", "1"),
    ];
    check_after_formulas(&formulas, &[(2, "H1", "Sheet1!A3"), (2, "H4", "2")]);
}

#[test]
fn a_type_change_under_a_criteria_range_keeps_the_graph() {
    let formulas = [
        (2, "H5", "=SUMIF(Sheet1!A1:A6,\">10\")"),
        (2, "H6", "=COUNTIF(Sheet1!A1:A6,\"x\")"),
    ];
    check_after_formulas(&formulas, &[(0, "A3", "x")]);
}

#[test]
fn a_value_read_by_a_defined_name_keeps_the_graph() {
    check(true, true, |workbook| edit(workbook, 1, "A2", "77"));
}
