use super::*;
use xlsx_model::{Cell, Sheet};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn put(wb: &mut Workbook, row: u32, col: u32, value: CellValue) {
    wb.sheets[0].set_cell(
        CellRef::new(row, col),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn formula(wb: &mut Workbook, row: u32, col: u32, source: &str) {
    wb.sheets[0].set_cell(
        CellRef::new(row, col),
        Cell {
            formula: Some(source.into()),
            ..Cell::default()
        },
    );
}

fn array(wb: &mut Workbook, row: u32, col: u32, source: &str, rows: u32, cols: u32) {
    formula(wb, row, col, source);
    wb.sheets[0].set_array_formula(
        CellRef::new(row, col),
        CellRange::new(
            CellRef::new(row, col),
            CellRef::new(row + rows - 1, col + cols - 1),
        ),
    );
}

fn empty() -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Data"));
    wb
}

fn recalc(wb: &mut Workbook, enabled: bool, limit: u64) -> (RecalcResult, Rc<EvaluationBudget>) {
    recalc_with(wb, enabled, EvaluationBudget::new(limit))
}

fn recalc_with(
    wb: &mut Workbook,
    enabled: bool,
    budget: EvaluationBudget,
) -> (RecalcResult, Rc<EvaluationBudget>) {
    let graph = DepGraph::build(wb);
    let recompute = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
    let budget = Rc::new(budget);
    budget.criteria_index_enabled.set(enabled);
    let result = run_recalc_with_budget(
        wb,
        &graph,
        recompute,
        RecalcContext {
            now_serial: Some(45_000.25),
            rand_seed: Some(42),
        },
        None,
        Rc::clone(&budget),
    );
    (result, budget)
}

/// with a budget the workbook fits in, the index and the streamed path
/// compute the same cells, values, spills and cycles.
fn compare(input: &Workbook, limit: u64) -> Workbook {
    let mut indexed = input.clone();
    let mut streamed = input.clone();
    let (actual, _) = recalc(&mut indexed, true, limit);
    let (expected, _) = recalc(&mut streamed, false, limit);
    assert!(!expected.budget_exhausted, "budget {limit} is not ample");
    assert_eq!(actual.budget_exhausted, expected.budget_exhausted);
    assert_eq!(actual.cycle_cells, expected.cycle_cells);
    for (id, (a, b)) in indexed.sheets.iter().zip(&streamed.sheets).enumerate() {
        let sheet = SheetId(id as u32);
        let cells: HashSet<CellRef> = a
            .iter_cells()
            .chain(b.iter_cells())
            .map(|(at, _)| at)
            .collect();
        for at in cells {
            let av = indexed.value(sheet, at);
            let bv = streamed.value(sheet, at);
            match (&av, &bv) {
                (CellValue::Number { value: a }, CellValue::Number { value: b }) => {
                    assert_eq!(a.to_bits(), b.to_bits(), "{sheet:?} {at:?}, budget {limit}");
                }
                _ => assert_eq!(av, bv, "{sheet:?} {at:?}, budget {limit}"),
            }
        }
        assert_eq!(
            a.array_formulas().collect::<Vec<_>>(),
            b.array_formulas().collect::<Vec<_>>()
        );
    }
    indexed
}

fn refused(value: &CellValue) -> bool {
    matches!(
        value,
        CellValue::Error {
            value: ErrorValue::Num
        }
    )
}

/// with a tight budget the index may compute cells the streamed path refuses
/// and so leave less for later cells, so refused cells can differ either way.
/// cells both compute agree, neither run charges past the budget, and the
/// cells index builds read, rolled back ones included, stay within it.
fn compare_budget(input: &Workbook, limit: u64) -> Workbook {
    let mut indexed = input.clone();
    let mut streamed = input.clone();
    let (_, on) = recalc(&mut indexed, true, limit);
    let (_, off) = recalc(&mut streamed, false, limit);
    assert!(on.remaining() <= limit && off.remaining() <= limit);
    assert!(
        on.criteria_index_reads.get() <= limit,
        "index read {} cells, budget {limit}",
        on.criteria_index_reads.get()
    );
    for (id, (a, b)) in indexed.sheets.iter().zip(&streamed.sheets).enumerate() {
        let sheet = SheetId(id as u32);
        let cells: HashSet<CellRef> = a
            .iter_cells()
            .chain(b.iter_cells())
            .map(|(at, _)| at)
            .collect();
        for at in cells {
            let av = indexed.value(sheet, at);
            let bv = streamed.value(sheet, at);
            if refused(&av) || refused(&bv) {
                continue;
            }
            match (&av, &bv) {
                (CellValue::Number { value: a }, CellValue::Number { value: b }) => {
                    assert_eq!(a.to_bits(), b.to_bits(), "{sheet:?} {at:?}, budget {limit}");
                }
                _ => assert_eq!(av, bv, "{sheet:?} {at:?}, budget {limit}"),
            }
        }
        let spills: HashMap<CellRef, CellRange> = b.array_formulas().collect();
        for (anchor, range) in a.array_formulas() {
            if let Some(other) = spills.get(&anchor)
                && !refused(&indexed.value(sheet, anchor))
                && !refused(&streamed.value(sheet, anchor))
            {
                assert_eq!(range, *other, "{sheet:?} {anchor:?}, budget {limit}");
            }
        }
    }
    indexed
}

fn spill_workbook() -> Workbook {
    let mut wb = empty();
    for (row, value) in [(1, 10.0), (2, 20.0), (3, 30.0)] {
        put(&mut wb, row, 2, number(value));
    }
    put(&mut wb, 0, 4, number(3.0));
    formula(&mut wb, 0, 0, "SUMPRODUCT((B2:B4=2)*C2:C4)");
    array(&mut wb, 0, 1, "_xlfn.SEQUENCE(E1)", 1, 1);
    wb
}

#[test]
fn criteria_index_differential_review_workbooks() {
    for source in [
        "SUMPRODUCT((B2:B4=2)*C2:C4)",
        "SUMPRODUCT((B2:B4=E1-1)*C2:C4)",
        "SUMPRODUCT((C2:C4=10)*B2:B4)",
    ] {
        let mut wb = spill_workbook();
        formula(&mut wb, 0, 0, source);
        compare(&wb, 100_000);
        put(&mut wb, 0, 4, number(1.0));
        let mut wb = compare(&wb, 100_000);
        put(&mut wb, 0, 4, number(3.0));
        compare(&wb, 100_000);
    }
    let mut a = spill_workbook();
    array(
        &mut a,
        0,
        6,
        "_xlfn.SEQUENCE(IF(SUMPRODUCT((B2:B4=2)*C2:C4)=10,1,2))",
        1,
        1,
    );
    array(&mut a, 1, 5, "_xlfn.SEQUENCE(1,2)", 1, 1);
    let result = compare(&a, 100_000);
    assert_eq!(result.value(SheetId(0), CellRef::new(1, 5)), number(1.0));
    assert_eq!(result.value(SheetId(0), CellRef::new(1, 6)), number(2.0));
    let mut b = spill_workbook();
    formula(
        &mut b,
        0,
        3,
        "IF(SUMPRODUCT((B2:B4=2)*C2:C4)=10,WEBSERVICE(\"x\"),99)",
    );
    b.sheets[0].cell_mut(CellRef::new(0, 3)).unwrap().value = number(42.0);
    let result = compare(&b, 100_000);
    assert_eq!(result.value(SheetId(0), CellRef::new(0, 3)), number(42.0));
    for whole_column in [false, true] {
        let mut wb = empty();
        let mut factors = Vec::new();
        for col in 0..16 {
            put(&mut wb, 0, col, number(1.0));
            let name = char::from(b'A' + col as u8);
            factors.push(if whole_column {
                format!("({name}:{name}=1)")
            } else {
                format!("({name}1:{name}1=1)")
            });
        }
        put(&mut wb, 0, 16, number(10.0));
        let values = if whole_column { "Q:Q" } else { "Q1:Q1" };
        formula(
            &mut wb,
            0,
            17,
            &format!("SUMPRODUCT({}*{values})", factors.join("*")),
        );
        formula(&mut wb, 0, 18, "A1");
        compare_budget(&wb, 18);
    }
}

#[test]
fn criteria_index_differential_settle_reads_spill_sources() {
    let mut wb = spill_workbook();
    let consumer = "IF(TRUE,SUMPRODUCT((B2:B4=2)*C2:C4),D1)";
    formula(&mut wb, 0, 0, consumer);
    formula(&mut wb, 1, 0, consumer);
    array(&mut wb, 0, 1, "_xlfn.SEQUENCE(3,1,A2+1)", 1, 1);
    formula(&mut wb, 0, 3, "A1+B1+A2+D1");
    compare(&wb, 100_000);
}

#[test]
fn criteria_index_differential_500_unread_ranges_small_budget() {
    let mut wb = empty();
    put(&mut wb, 0, 0, number(1.0));
    for row in 0..500 {
        formula(
            &mut wb,
            row,
            2,
            &format!(
                "SUMPRODUCT((A{}:A1000000=1)*B{}:B1000000)",
                row + 2,
                row + 2
            ),
        );
    }
    compare_budget(&wb, 1_000);
    compare(&wb, 10_000_000);
}

#[test]
fn criteria_index_needs_the_streamed_headroom() {
    // the streamed path needs 22 cells of headroom here; the index alone
    // would charge 20 and leave nothing for C2
    let mut wb = empty();
    put(&mut wb, 9, 0, number(2.0));
    formula(&mut wb, 0, 2, "SUMPRODUCT((A1:A21=1)*B1:B21)");
    formula(&mut wb, 1, 2, "A10");
    let result = compare(&wb, 1_000);
    assert_eq!(result.value(SheetId(0), CellRef::new(0, 2)), number(0.0));
    let mut indexed = wb.clone();
    let mut streamed = wb.clone();
    recalc(&mut indexed, true, 20);
    recalc(&mut streamed, false, 20);
    for row in 0..2 {
        let at = CellRef::new(row, 2);
        assert_eq!(
            indexed.value(SheetId(0), at),
            streamed.value(SheetId(0), at)
        );
    }
    assert_eq!(indexed.value(SheetId(0), CellRef::new(1, 2)), number(2.0));
}

#[test]
fn criteria_index_ranges_below_the_used_rows_need_the_streamed_headroom() {
    let mut wb = empty();
    put(&mut wb, 0, 0, number(1.0));
    formula(&mut wb, 0, 2, "SUMPRODUCT((A2:AMJ2=0)*A3:AMJ3)");
    formula(&mut wb, 0, 3, "A1");
    compare(&wb, 100_000);
    let mut indexed = wb.clone();
    let mut streamed = wb.clone();
    recalc(&mut indexed, true, 64);
    recalc(&mut streamed, false, 64);
    for col in 2..4 {
        let at = CellRef::new(0, col);
        assert_eq!(
            indexed.value(SheetId(0), at),
            streamed.value(SheetId(0), at)
        );
    }
    assert_eq!(indexed.value(SheetId(0), CellRef::new(0, 3)), number(1.0));
}

#[test]
fn criteria_index_refusals_read_no_ranges() {
    // a pair the budget cannot cover is refused before either range is read,
    // so a thousand such formulas never reread the affordable one
    let mut wb = empty();
    put(&mut wb, 9_999, 0, number(1.0));
    for row in 0..1_000 {
        formula(&mut wb, row, 2, "SUMPRODUCT((A1:A10001=1)*B9001:B19001)");
    }
    compare_budget(&wb, 10_000);
    let mut indexed = wb.clone();
    let (_, budget) = recalc(&mut indexed, true, 10_000);
    assert_eq!(budget.criteria_index_reads.get(), 0);
}

/// the index leaves this workbook to the existing path: every cell matches
/// the index off run and no range is read for an index.
fn compare_untouched(input: &Workbook, budget: impl Fn() -> EvaluationBudget) {
    let mut indexed = input.clone();
    let mut streamed = input.clone();
    let (_, on) = recalc_with(&mut indexed, true, budget());
    recalc_with(&mut streamed, false, budget());
    assert_eq!(on.criteria_index_reads.get(), 0);
    for (a, b) in indexed.sheets.iter().zip(&streamed.sheets) {
        let cells: HashSet<CellRef> = a
            .iter_cells()
            .chain(b.iter_cells())
            .map(|(at, _)| at)
            .collect();
        for at in cells {
            assert_eq!(
                indexed.value(SheetId(0), at),
                streamed.value(SheetId(0), at),
                "{at:?}"
            );
        }
    }
}

#[test]
fn criteria_index_needs_a_scalar_key_before_reading_ranges() {
    for source in [
        "SUMPRODUCT((D1001:X51000=A1:A50000)*B1:B50000)",
        "SUMPRODUCT((A1:A600000=C1:D1)*B1:B600000)",
    ] {
        let mut wb = empty();
        put(&mut wb, 999, 0, number(1.0));
        for row in 0..100 {
            formula(&mut wb, row, 4, source);
        }
        compare_untouched(&wb, || EvaluationBudget::new(10_000));
    }
    let mut wb = empty();
    put(&mut wb, 0, 0, number(1.0));
    formula(&mut wb, 0, 2, "SUMPRODUCT((+D2:D600=A1:AMJ1)*B1:AMJ1)");
    compare_untouched(&wb, || EvaluationBudget::new(2_500));
}

#[test]
fn criteria_index_leaves_computed_keys_to_the_existing_path() {
    let mut wb = empty();
    put(&mut wb, 48, 0, number(1.0));
    formula(
        &mut wb,
        0,
        2,
        "SUMPRODUCT((SUM(D1:D49)=A1:A10)*(A1:A10=0)*B1:B10)",
    );
    formula(&mut wb, 1, 2, "A49");
    compare_untouched(&wb, || EvaluationBudget::new(70));
    let mut wb = empty();
    put(&mut wb, 9, 0, CellValue::Text { value: "ab".into() });
    put(&mut wb, 9, 1, number(5.0));
    formula(&mut wb, 0, 2, "SUMPRODUCT(--(\"a\"&\"b\"=A1:A10)*B1:B10)");
    compare_untouched(&wb, || EvaluationBudget::new(40));
}

#[test]
fn criteria_index_charges_no_key_text_the_streamed_path_does_not() {
    let mut wb = empty();
    for row in 0..2 {
        put(&mut wb, row, 0, CellValue::Text { value: "x".into() });
        put(&mut wb, row, 1, number(row as f64 + 1.0));
    }
    put(&mut wb, 0, 3, CellValue::Text { value: "x".into() });
    formula(&mut wb, 0, 2, "SUMPRODUCT((A1:A2=+D1)*B1:B2)");
    formula(&mut wb, 1, 2, "\"y\"");
    compare_untouched(&wb, || EvaluationBudget::with_text(100, 1));
}

struct Lcg(u32);

impl Lcg {
    fn next(&mut self) -> u32 {
        self.0 = self.0.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
        self.0
    }

    fn value(&mut self) -> CellValue {
        match self.next() % 9 {
            0 => CellValue::Empty,
            1 => CellValue::Text {
                value: "Apple".into(),
            },
            2 => CellValue::Text {
                value: "APPLE".into(),
            },
            3 => CellValue::Text { value: "".into() },
            4 => CellValue::Bool { value: false },
            5 => CellValue::Bool { value: true },
            6 => CellValue::Error {
                value: ErrorValue::Div0,
            },
            7 => CellValue::Text { value: "2".into() },
            _ => number((self.next() % 7) as f64 - 3.0),
        }
    }
}

#[test]
fn criteria_index_differential_generated() {
    let mut rng = Lcg(0x5eed_1234);
    for case in 0..320 {
        let mut wb = empty();
        for row in 1..9 {
            let key = rng.value();
            let value = if case % 3 == 0 {
                rng.value()
            } else {
                number((rng.next() % 17) as f64)
            };
            put(&mut wb, row, 1, key);
            put(&mut wb, row, 2, value);
            put(&mut wb, row, 3, number((row % 2) as f64));
        }
        let key = ["1", "0", "\"\"", "\"apple\"", "FALSE", "TRUE", "1/0", "2+3"][case % 8];
        let end = if case % 5 == 0 { 1_000_000 } else { 9 };
        let product = format!("SUMPRODUCT((B2:B{end}={key})*(D2:D{end}=1)*C2:C{end})");
        formula(&mut wb, 0, 0, &product);
        formula(&mut wb, 0, 5, &product);
        formula(&mut wb, 10, 0, &product);
        match case % 8 {
            0 | 1 => {
                let col = if case % 8 == 0 { 1 } else { 2 };
                for row in 1..4 {
                    put(&mut wb, row, col, CellValue::Empty);
                }
                array(&mut wb, 0, col, "_xlfn.SEQUENCE(4)", 1, 1);
            }
            2 => {
                array(&mut wb, 0, 1, "_xlfn.SEQUENCE(1)", 4, 1);
            }
            3 => {
                array(&mut wb, 0, 1, "_xlfn.SEQUENCE(4)", 1, 1);
            }
            4 => {
                formula(&mut wb, 2, 1, "1+0");
                formula(&mut wb, 3, 2, "2+3");
            }
            5 => {
                formula(&mut wb, 1, 1, "IF(TRUE,1,F1)");
                formula(&mut wb, 0, 5, &format!("{product}+F1"));
            }
            6 => {
                formula(
                    &mut wb,
                    0,
                    0,
                    "SUMPRODUCT((B20:B1000000=\"\")*C20:C1000000)",
                );
                formula(&mut wb, 0, 5, "SUMPRODUCT((B20:B1000000=1)*C20:C1000000)");
            }
            _ => {
                for col in 7..15 {
                    let key = rng.value();
                    put(&mut wb, 1, col, key);
                    put(&mut wb, 2, col, number((rng.next() % 9) as f64));
                }
                formula(&mut wb, 0, 0, &format!("SUMPRODUCT((H2:O2={key})*H3:O3)"));
                if case % 16 == 7 {
                    for col in 7..10 {
                        put(&mut wb, 1, col, CellValue::Empty);
                    }
                }
                array(&mut wb, 1, 6, "_xlfn.SEQUENCE(1,4)", 1, 1);
                formula(&mut wb, 3, 0, &format!("SUMPRODUCT((H2:O2={key})*H3:O3)"));
            }
        }
        if case % 4 != 0 {
            for at in [CellRef::new(0, 0), CellRef::new(0, 5), CellRef::new(10, 0)] {
                wb.sheets[0].set_array_formula(at, CellRange::new(at, at));
            }
        }
        if case % 4 == 1 {
            array(&mut wb, 12, 5, &format!("{product}*{{1,2,3}}"), 1, 3);
            for (row, value) in [(0, 1.0), (1, 0.0), (2, -1.0)] {
                put(&mut wb, row, 4, number(value));
            }
            array(&mut wb, 14, 5, &format!("IF(E1:E3>0,{product},0)"), 3, 1);
        }
        let mut next = compare(&wb, 100_000);
        compare_budget(&wb, 1);
        if case % 8 == 4 {
            compare_budget(&wb, 8 + (rng.next() % 40) as u64);
        }
        if case % 8 <= 3 {
            let col = if case % 8 == 1 { 2 } else { 1 };
            array(&mut next, 0, col, "_xlfn.SEQUENCE(1)", 4, 1);
            let mut next = compare(&next, 100_000);
            array(&mut next, 0, col, "_xlfn.SEQUENCE(4)", 1, 1);
            compare(&next, 100_000);
        }
    }
}

#[test]
fn criteria_index_cse_differential_types_errors_and_arrays() {
    let mixed = [
        number(1.0),
        CellValue::Text { value: "x".into() },
        CellValue::Text { value: "1".into() },
        CellValue::Bool { value: true },
        CellValue::Empty,
        number(0.0),
        CellValue::Bool { value: false },
        CellValue::Text { value: "".into() },
    ];
    for profile in 0..6 {
        for end in [9, 50] {
            for key in &mixed {
                let mut input = empty();
                for (offset, value) in mixed.iter().enumerate() {
                    let row = offset as u32 + 1;
                    put(
                        &mut input,
                        row,
                        0,
                        if profile == 0 || profile == 2 {
                            number((offset % 2) as f64)
                        } else {
                            value.clone()
                        },
                    );
                    put(
                        &mut input,
                        row,
                        1,
                        if profile == 2 || profile == 3 {
                            value.clone()
                        } else if offset == 4 {
                            CellValue::Empty
                        } else {
                            number(row as f64 / 8.0)
                        },
                    );
                }
                if profile == 4 || profile == 5 {
                    put(
                        &mut input,
                        3,
                        u32::from(profile == 5),
                        CellValue::Error {
                            value: if profile == 4 {
                                ErrorValue::Div0
                            } else {
                                ErrorValue::NA
                            },
                        },
                    );
                }
                put(&mut input, 0, 3, key.clone());
                for (row, value) in [(0, 1.0), (1, 0.0), (2, -1.0)] {
                    put(&mut input, row, 4, number(value));
                }
                let product = format!("SUMPRODUCT((A2:A{end}=D1)*B2:B{end})");
                for (source, rows, cols) in [
                    (product.clone(), 1, 1),
                    (product.clone(), 3, 1),
                    (format!("{product}*{{1,2,3}}"), 1, 3),
                    (format!("{product}+E1:E3"), 3, 1),
                    (format!("IF(E1:E3>0,{product},0)"), 3, 1),
                ] {
                    let mut wb = input.clone();
                    array(&mut wb, 0, 6, &source, rows, cols);
                    compare(&wb, 1_000_000);
                    if profile <= 1 && !matches!(key, CellValue::Empty) {
                        let (_, budget) = recalc(&mut wb, true, 1_000_000);
                        assert!(
                            budget.criteria_index_reads.get() > 0,
                            "{source}, profile {profile}, key {key:?}"
                        );
                    }
                }
            }
        }
    }
}

#[test]
fn criteria_index_cse_single_cell_errors_and_non_numeric_values() {
    for (col, value, expected) in [
        (
            0,
            CellValue::Error {
                value: ErrorValue::Div0,
            },
            CellValue::Error {
                value: ErrorValue::Div0,
            },
        ),
        (
            1,
            CellValue::Error {
                value: ErrorValue::NA,
            },
            CellValue::Error {
                value: ErrorValue::NA,
            },
        ),
        (
            1,
            CellValue::Text { value: "x".into() },
            CellValue::Error {
                value: ErrorValue::Value,
            },
        ),
        (1, CellValue::Bool { value: true }, number(6.0)),
        (1, CellValue::Bool { value: false }, number(5.0)),
        (1, CellValue::Text { value: "7".into() }, number(12.0)),
        (1, CellValue::Empty, number(5.0)),
    ] {
        let mut wb = empty();
        for row in 1..4 {
            put(&mut wb, row, 0, number(1.0));
            put(&mut wb, row, 1, number(if row == 1 { 2.0 } else { 3.0 }));
        }
        put(&mut wb, 2, col, value);
        array(&mut wb, 0, 5, "SUMPRODUCT((A2:A4=1)*B2:B4)", 1, 1);
        let result = compare(&wb, 100_000);
        assert_eq!(result.value(SheetId(0), CellRef::new(0, 5)), expected);
    }
    for value in [
        CellValue::Text { value: "x".into() },
        CellValue::Text { value: "7".into() },
    ] {
        let mut wb = empty();
        for row in 1..4 {
            put(&mut wb, row, 0, number(1.0));
            put(&mut wb, row, 1, number(if row == 1 { 2.0 } else { 3.0 }));
        }
        put(&mut wb, 2, 1, value);
        array(&mut wb, 0, 5, "SUMPRODUCT((A2:A4=1),B2:B4)", 1, 1);
        let result = compare(&wb, 100_000);
        assert_eq!(result.value(SheetId(0), CellRef::new(0, 5)), number(5.0));
    }
}

#[test]
fn criteria_index_cse_scalar_and_broadcast_parity() {
    for key in ["D1", "D1:D1", "2", "\"x\""] {
        for rows in [0, 1, 3] {
            let mut wb = empty();
            for row in 1..9 {
                put(&mut wb, row, 0, number((row % 3) as f64));
                put(&mut wb, row, 2, number(row as f64 / 8.0));
            }
            put(&mut wb, 0, 3, number(2.0));
            let source = format!("SUMPRODUCT((A2:A9={key})*C2:C9)");
            if rows == 0 {
                formula(&mut wb, 5, 1, &source);
            } else {
                array(&mut wb, 5, 1, &source, rows, 1);
            }
            let result = compare(&wb, 100_000);
            let expected = if key == "\"x\"" { 0.0 } else { 15.0 / 8.0 };
            for row in 5..5 + rows.max(1) {
                assert_eq!(
                    result.value(SheetId(0), CellRef::new(row, 1)),
                    number(expected),
                    "{key}, {rows} rows"
                );
            }
            let mut indexed = wb.clone();
            let (_, budget) = recalc(&mut indexed, true, 100_000);
            assert!(budget.criteria_index_reads.get() > 0, "{key}, {rows} rows");
        }
    }
}

#[test]
fn criteria_index_cse_refusals_preserve_existing_results() {
    for source in [
        "SUMPRODUCT((A2:A9=D1:D8)*C2:C9)",
        "_xlfn.MAP(D1:D3,_xlfn.LAMBDA(_xlpm.k,SUMPRODUCT((A2:A9=_xlpm.k)*C2:C9)))",
        "_xlfn.LET(_xlpm.k,D1,SUMPRODUCT((A2:A9=_xlpm.k)*C2:C9))",
        "SUMPRODUCT((A2:A9=D1),C2:C9)",
        "SUMPRODUCT((A2:A9=D1),C2:C9,C2:C9)",
    ] {
        let mut wb = empty();
        for row in 0..9 {
            put(&mut wb, row, 0, number((row % 3) as f64));
            put(&mut wb, row, 2, number(row as f64));
            put(&mut wb, row, 3, number((row % 3) as f64));
        }
        array(&mut wb, 0, 5, source, 3, 1);
        compare(&wb, 100_000);
        compare_untouched(&wb, || EvaluationBudget::new(100_000));
    }
}

#[test]
fn criteria_index_cse_result_supports_further_array_arithmetic() {
    let mut wb = empty();
    for row in 1..9 {
        put(&mut wb, row, 0, number((row % 3) as f64));
        put(&mut wb, row, 2, number(row as f64));
    }
    put(&mut wb, 0, 3, number(2.0));
    array(&mut wb, 0, 5, "SUMPRODUCT((A2:A9=D1)*C2:C9)+D1:D3", 3, 1);
    let result = compare(&wb, 100_000);
    for (row, expected) in [(0, 17.0), (1, 15.0), (2, 15.0)] {
        assert_eq!(
            result.value(SheetId(0), CellRef::new(row, 5)),
            number(expected)
        );
    }
    let mut indexed = wb.clone();
    let (_, budget) = recalc(&mut indexed, true, 100_000);
    assert!(budget.criteria_index_reads.get() > 0);
}

#[test]
fn criteria_index_cse_needs_the_streamed_headroom() {
    let mut wb = empty();
    put(&mut wb, 9, 0, number(2.0));
    array(&mut wb, 0, 2, "SUMPRODUCT((A1:A21=1)*B1:B21)", 1, 1);
    formula(&mut wb, 1, 2, "A10");
    compare(&wb, 1_000);
    let mut indexed = wb.clone();
    let mut streamed = wb.clone();
    let (_, budget) = recalc(&mut indexed, true, 20);
    recalc(&mut streamed, false, 20);
    assert_eq!(budget.criteria_index_reads.get(), 0);
    for row in 0..2 {
        let at = CellRef::new(row, 2);
        assert_eq!(
            indexed.value(SheetId(0), at),
            streamed.value(SheetId(0), at)
        );
    }
    assert_eq!(indexed.value(SheetId(0), CellRef::new(1, 2)), number(2.0));
}

#[test]
fn criteria_index_cse_refusals_read_no_ranges() {
    let mut wb = empty();
    put(&mut wb, 9_999, 0, number(1.0));
    for row in 0..1_000 {
        array(
            &mut wb,
            row,
            2,
            "SUMPRODUCT((A1:A10001=1)*B9001:B19001)",
            1,
            1,
        );
    }
    compare_budget(&wb, 10_000);
    let mut indexed = wb.clone();
    let (_, budget) = recalc(&mut indexed, true, 10_000);
    assert_eq!(budget.criteria_index_reads.get(), 0);
}

#[test]
fn criteria_index_cse_needs_literal_text_headroom() {
    let mut wb = empty();
    for row in 1..9 {
        put(
            &mut wb,
            row,
            0,
            CellValue::Text {
                value: "apple".into(),
            },
        );
        put(&mut wb, row, 1, number(row as f64));
    }
    array(&mut wb, 0, 2, "SUMPRODUCT((A2:A9=\"apple\")*B2:B9)", 1, 1);
    compare_untouched(&wb, || EvaluationBudget::with_text(100_000, 4));
}

#[test]
fn criteria_indexes_cover_final_anchors_and_invalidate_on_writes() {
    let mut wb = empty();
    for row in 1..30 {
        put(&mut wb, row, 0, number((row % 4) as f64));
        put(&mut wb, row, 2, number(row as f64 / 4.0));
    }
    array(&mut wb, 7, 0, "1+2", 1, 1);
    array(&mut wb, 11, 2, "SUM(C2:C5)", 1, 1);
    formula(&mut wb, 12, 2, "1+1");
    put(&mut wb, 0, 3, number(3.0));
    for row in 0..12 {
        array(
            &mut wb,
            row,
            4,
            &format!("IF(D1>0,SUMPRODUCT((A2:A30=D1)*C2:C30)+{row},0)"),
            1,
            1,
        );
        array(
            &mut wb,
            row,
            5,
            &format!("SUMPRODUCT((A13:A30=D1)*C13:C30)+{row}"),
            1,
            1,
        );
    }
    compare(&wb, 1_000_000);
    let mut indexed = wb.clone();
    let (_, budget) = recalc(&mut indexed, true, 1_000_000);
    assert!(budget.criteria_index_reads.get() > 0);
    assert!(budget.criteria_index_covers(SheetId(0), CellRef::new(7, 0)));
    assert!(budget.criteria_index_covers(SheetId(0), CellRef::new(11, 2)));
    let at = CellRef::new(7, 0);
    let mut changed = Vec::new();
    write_spill(
        &mut indexed,
        key(SheetId(0), at),
        Spill {
            range: CellRange::new(at, at),
            values: vec![number(9.0)],
        },
        &budget,
        &mut changed,
    )
    .unwrap();
    assert!(!budget.criteria_index_covers(SheetId(0), at));
    let (_, budget) = recalc(&mut indexed, true, 1_000_000);
    let at = CellRef::new(12, 2);
    assert!(budget.criteria_index_covers(SheetId(0), at));
    assert!(write_if_changed(
        &mut indexed,
        key(SheetId(0), at),
        number(99.0),
        &budget,
    ));
    assert!(!budget.criteria_index_covers(SheetId(0), at));
}

#[test]
fn criteria_index_differential_formula_dependencies_generated() {
    let mut rng = Lcg(0xf012_3456);
    for case in 0..80 {
        let mut wb = empty();
        for row in 1..9 {
            put(&mut wb, row, 0, number((rng.next() % 3) as f64));
            put(&mut wb, row, 1, number((rng.next() % 13) as f64));
        }
        put(&mut wb, 0, 7, number(2.0));
        formula(&mut wb, 0, 6, "H1+1");
        formula(&mut wb, 1, 0, "1+1");
        match case % 8 {
            0 => formula(&mut wb, 2, 1, "2+3"),
            1 => formula(&mut wb, 2, 1, "G1+H1"),
            2 => {
                formula(&mut wb, 0, 5, "SUMPRODUCT((A6:A9=1)*B6:B9)");
                formula(&mut wb, 2, 1, "F1+1");
            }
            3 => array(&mut wb, 2, 1, "G1+1", 1, 1),
            4 => {
                put(&mut wb, 3, 1, CellValue::Empty);
                array(&mut wb, 2, 1, "_xlfn.SEQUENCE(2)", 2, 1);
            }
            5 => formula(&mut wb, 2, 1, "INDIRECT(\"H1\")"),
            6 => formula(&mut wb, 2, 1, "OFFSET(H1,0,G1-3)"),
            _ => {
                formula(&mut wb, 2, 1, "E1+1");
                formula(&mut wb, 2, 2, "B3+1");
            }
        }
        for row in 0..3 {
            let source = "SUMPRODUCT((A2:A9=2)*B2:B9)";
            if case % 2 == 0 {
                formula(&mut wb, row, 4, source);
            } else {
                array(&mut wb, row, 4, source, 1, 1);
            }
        }
        let mut indexed = compare(&wb, 100_000);
        if case % 8 <= 3 {
            let (_, budget) = recalc(&mut indexed, true, 100_000);
            assert!(budget.criteria_index_covers(SheetId(0), CellRef::new(1, 0)));
            assert!(budget.criteria_index_covers(SheetId(0), CellRef::new(2, 1)));
        }
    }
}

#[test]
fn criteria_index_refuses_volatile_formula_ranges() {
    for source in ["NOW()", "TODAY()", "INDIRECT(\"H1\")", "OFFSET(H1,0,G1)"] {
        let mut wb = empty();
        put(&mut wb, 0, 7, number(2.0));
        put(&mut wb, 0, 6, number(0.0));
        formula(&mut wb, 1, 0, source);
        put(&mut wb, 2, 0, number(2.0));
        put(&mut wb, 1, 1, number(3.0));
        put(&mut wb, 2, 1, number(4.0));
        formula(&mut wb, 0, 4, "SUMPRODUCT((A2:A3=2)*B2:B3)");
        let indexed = compare(&wb, 100_000);
        if source == "NOW()" || source == "TODAY()" {
            assert_eq!(
                indexed.value(SheetId(0), CellRef::new(1, 0)),
                number(if source == "NOW()" {
                    45_000.25
                } else {
                    45_000.0
                }),
            );
        }
        let (_, budget) = recalc(&mut wb, true, 100_000);
        assert!(!budget.criteria_index_covers(SheetId(0), CellRef::new(1, 0)));
        assert!(!budget.criteria_index_covers(SheetId(0), CellRef::new(1, 1)));
    }
}

#[test]
fn criteria_index_formula_names_tables_and_cross_sheet_dependencies() {
    let mut wb = empty();
    let mut inputs = Sheet::new("Inputs");
    inputs.set_cell(
        CellRef::new(1, 0),
        Cell {
            formula: Some("1+1".into()),
            ..Cell::default()
        },
    );
    wb.sheets.push(inputs);
    wb.tables.push(xlsx_model::Table {
        name: "InputsTable".into(),
        sheet: SheetId(1),
        range: CellRange::new(CellRef::new(0, 0), CellRef::new(1, 0)),
        header_rows: 1,
        totals_rows: 0,
        columns: vec!["Value".into()],
    });
    wb.defined_names.push(xlsx_model::DefinedName {
        name: "InputValue".into(),
        formula: "SUM(InputsTable[Value])".into(),
        local_sheet: None,
        hidden: false,
    });
    formula(&mut wb, 1, 0, "Inputs!A2");
    formula(&mut wb, 2, 0, "InputValue");
    formula(&mut wb, 1, 1, "SUM(InputsTable[Value])+1");
    formula(&mut wb, 2, 1, "InputValue+2");
    formula(&mut wb, 0, 4, "SUMPRODUCT((A2:A3=2)*B2:B3)");
    let mut indexed = compare(&wb, 100_000);
    assert_eq!(indexed.value(SheetId(0), CellRef::new(0, 4)), number(7.0));
    let (_, budget) = recalc(&mut indexed, true, 100_000);
    assert!(budget.criteria_index_covers(SheetId(0), CellRef::new(1, 0)));
    assert!(budget.criteria_index_covers(SheetId(0), CellRef::new(1, 1)));
}

fn spill_reader_workbook(size: f64, reader_first: bool) -> Workbook {
    let mut wb = empty();
    for row in 0..3 {
        put(&mut wb, row, 0, number(1.0));
    }
    put(&mut wb, 1, 1, number(10.0));
    put(&mut wb, 2, 1, number(20.0));
    put(&mut wb, 0, 5, number(size));
    formula(&mut wb, 0, 1, "D2+1");
    array(
        &mut wb,
        0,
        3,
        if reader_first {
            "_xlfn.SEQUENCE(F1+B1*0,1,100)"
        } else {
            "_xlfn.SEQUENCE(F1,1,100)"
        },
        1,
        1,
    );
    formula(&mut wb, 0, 4, "F1+SUMPRODUCT((A1:A3=1)*B1:B3)");
    wb
}

#[test]
fn criteria_index_spill_reader_fresh_and_incremental_match_streamed() {
    let sheet = SheetId(0);
    let dashboard = CellRef::new(0, 4);
    for reader_first in [false, true] {
        let fresh = compare(&spill_reader_workbook(2.0, reader_first), 100_000);
        assert_eq!(fresh.value(sheet, dashboard), number(134.0));
        let mut results = Vec::new();
        for enabled in [false, true] {
            let mut wb = spill_reader_workbook(1.0, reader_first);
            let (initial, _) = recalc(&mut wb, enabled, 100_000);
            assert!(!initial.budget_exhausted);
            assert_eq!(wb.value(sheet, dashboard), number(32.0));
            let graph = DepGraph::build(&wb);
            put(&mut wb, 0, 5, number(2.0));
            let pending = collect_recompute(&graph, &[(sheet, CellRef::new(0, 5))]);
            if !reader_first {
                assert!(!pending.contains(&key(sheet, CellRef::new(0, 1))));
            }
            let budget = Rc::new(EvaluationBudget::new(100_000));
            budget.criteria_index_enabled.set(enabled);
            let result = run_recalc_with_budget(
                &mut wb,
                &graph,
                pending,
                RecalcContext {
                    now_serial: None,
                    rand_seed: None,
                },
                None,
                budget,
            );
            assert!(!result.budget_exhausted);
            assert_eq!(wb.value(sheet, CellRef::new(0, 1)), number(102.0));
            assert_eq!(wb.value(sheet, dashboard), number(134.0));
            results.push(wb);
        }
        let cells: HashSet<CellRef> = results[0].sheets[0]
            .iter_cells()
            .chain(results[1].sheets[0].iter_cells())
            .map(|(at, _)| at)
            .collect();
        for at in cells {
            assert_eq!(results[0].value(sheet, at), results[1].value(sheet, at));
        }
        assert_eq!(
            results[0].sheets[0].array_formulas().collect::<Vec<_>>(),
            results[1].sheets[0].array_formulas().collect::<Vec<_>>()
        );
    }
}

#[test]
fn criteria_index_spill_schedules_guard_before_later_readers() {
    for reader_first in [false, true] {
        let mut wb = spill_reader_workbook(2.0, reader_first);
        let graph = DepGraph::build(&wb);
        let budget = Rc::new(EvaluationBudget::new(100_000));
        let mut pass = Pass {
            graph: &graph,
            name_index: Some(&graph),
            context: RecalcContext {
                now_serial: None,
                rand_seed: None,
            },
            budget: Rc::clone(&budget),
            changed: Vec::new(),
            cycle_cells: Vec::new(),
            limited_cells: Vec::new(),
            spilled: Vec::new(),
        };
        let sheet = SheetId(0);
        let reader = CellRef::new(0, 1);
        if reader_first {
            pass.evaluate(&mut wb, key(sheet, reader));
        }
        pass.evaluate(&mut wb, key(sheet, CellRef::new(0, 3)));
        assert!(!pass.spilled.is_empty());
        pass.evaluate(&mut wb, key(sheet, CellRef::new(0, 4)));
        assert!(!budget.criteria_index_covers(sheet, reader));
    }
}
