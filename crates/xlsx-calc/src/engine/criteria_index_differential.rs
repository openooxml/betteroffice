use super::*;
use xlsx_model::Sheet;

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
            now_serial: None,
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
        match self.next() % 8 {
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
