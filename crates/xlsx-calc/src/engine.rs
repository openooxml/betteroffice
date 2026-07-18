//! recalc driver: given edited cells, re-evaluate exactly the formulas that
//! could have changed, in dependency order, and report what moved.

use std::collections::{HashMap, HashSet, VecDeque};
use std::rc::Rc;

use xlsx_model::{CellProvider, CellRef, CellValue, ColId, RowId, SheetId, Workbook};

use crate::eval::{EvalContext, EvaluationBudget, MAX_RECALCULATION_CELL_VISITS, evaluate};
use crate::graph::DepGraph;
use crate::parser::parse_formula;

/// the outcome of a recalc: cells whose displayed value changed, and cells
/// forced to `0` by cycle participation.
pub struct RecalcResult {
    pub changed: Vec<(SheetId, CellRef)>,
    pub cycle_cells: Vec<(SheetId, CellRef)>,
    pub limited_cells: Vec<(SheetId, CellRef)>,
}

/// normalized cell identity (avoids `$`-anchor `Hash`/`Eq` mismatches).
type Key = (SheetId, RowId, ColId);

fn key(sheet: SheetId, cell: CellRef) -> Key {
    (sheet, cell.row, cell.col)
}

fn cell_of(k: Key) -> CellRef {
    CellRef::new(k.1, k.2)
}

/// incremental recalc after `dirty_seeds` were edited: re-evaluates their
/// transitive dependents plus volatile cells, returns what moved.
pub fn recalc_after(
    wb: &mut Workbook,
    graph: &mut DepGraph,
    dirty_seeds: &[(SheetId, CellRef)],
    now_serial: Option<f64>,
) -> RecalcResult {
    let recompute = collect_recompute(graph, dirty_seeds);
    run_recalc(wb, graph, recompute, now_serial)
}

/// rebuild the graph from scratch and recalc every formula in dependency order.
pub fn rebuild_and_recalc_all(
    wb: &mut Workbook,
    now_serial: Option<f64>,
) -> (DepGraph, RecalcResult) {
    let graph = DepGraph::build(wb);
    let recompute: HashSet<Key> = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
    let result = run_recalc(wb, &graph, recompute, now_serial);
    (graph, result)
}

/// formula cells to re-evaluate: transitive dependents of the seeds, plus
/// volatile cells and their dependents.
fn collect_recompute(graph: &DepGraph, seeds: &[(SheetId, CellRef)]) -> HashSet<Key> {
    let mut recompute: HashSet<Key> = HashSet::new();
    let mut worklist: Vec<Key> = Vec::new();

    for &(sheet, cell) in seeds {
        let k = key(sheet, cell);
        worklist.push(k);
        if graph.is_formula(sheet, cell) {
            recompute.insert(k);
        }
    }
    for (sheet, cell) in graph.volatile_cells() {
        let k = key(sheet, cell);
        recompute.insert(k);
        worklist.push(k);
    }

    let mut expanded: HashSet<Key> = HashSet::new();
    while let Some(k) = worklist.pop() {
        if !expanded.insert(k) {
            continue;
        }
        for (ds, dc) in graph.dependents_of(k.0, cell_of(k)) {
            let dk = key(ds, dc);
            recompute.insert(dk);
            worklist.push(dk);
        }
    }
    recompute
}

/// topologically order `recompute` and evaluate it, writing changed values into
/// `wb`. cells caught in a cycle are zeroed and reported separately.
fn run_recalc(
    wb: &mut Workbook,
    graph: &DepGraph,
    recompute: HashSet<Key>,
    now_serial: Option<f64>,
) -> RecalcResult {
    let (order, cycle) = topo_order(graph, &recompute);
    let budget = Rc::new(EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS));

    let mut changed: Vec<(SheetId, CellRef)> = Vec::new();
    let mut limited_cells = Vec::new();
    for u in &order {
        let (value, limited) = eval_node(wb, *u, now_serial, Rc::clone(&budget));
        if limited {
            limited_cells.push((u.0, cell_of(*u)));
        }
        if let Some(value) = value
            && write_if_changed(wb, *u, value)
        {
            changed.push((u.0, cell_of(*u)));
        }
    }

    let mut cycle_cells: Vec<(SheetId, CellRef)> = Vec::new();
    for u in &cycle {
        if write_if_changed(wb, *u, CellValue::Number { value: 0.0 }) {
            changed.push((u.0, cell_of(*u)));
        }
        cycle_cells.push((u.0, cell_of(*u)));
    }

    changed.sort_by(sort_key);
    RecalcResult {
        changed,
        cycle_cells,
        limited_cells,
    }
}

/// kahn's sort over the sub-graph induced by `recompute`: returns the evaluable
/// order and, separately, the cells caught in (or only reachable through) a cycle.
fn topo_order(graph: &DepGraph, recompute: &HashSet<Key>) -> (Vec<Key>, Vec<Key>) {
    let mut adj: HashMap<Key, Vec<Key>> = HashMap::new();
    let mut indegree: HashMap<Key, usize> = recompute.iter().map(|k| (*k, 0)).collect();

    for &u in recompute {
        let mut seen: HashSet<Key> = HashSet::new();
        for (ds, dc) in graph.dependents_of(u.0, cell_of(u)) {
            let v = key(ds, dc);
            if recompute.contains(&v) && seen.insert(v) {
                adj.entry(u).or_default().push(v);
                *indegree.get_mut(&v).unwrap() += 1;
            }
        }
    }

    let mut queue: VecDeque<Key> = {
        let mut ready: Vec<Key> = recompute
            .iter()
            .copied()
            .filter(|k| indegree.get(k) == Some(&0))
            .collect();
        ready.sort();
        ready.into_iter().collect()
    };

    let mut order: Vec<Key> = Vec::new();
    while let Some(u) = queue.pop_front() {
        order.push(u);
        if let Some(children) = adj.get(&u) {
            let mut ready: Vec<Key> = Vec::new();
            for &v in children {
                let d = indegree.get_mut(&v).unwrap();
                *d -= 1;
                if *d == 0 {
                    ready.push(v);
                }
            }
            ready.sort();
            queue.extend(ready);
        }
    }

    let ordered: HashSet<Key> = order.iter().copied().collect();
    let mut cycle: Vec<Key> = recompute
        .iter()
        .copied()
        .filter(|k| !ordered.contains(k))
        .collect();
    cycle.sort();
    (order, cycle)
}

/// evaluate one formula node; `None` when the cell has no formula or it no
/// longer parses (cached value left untouched).
fn eval_node(
    wb: &Workbook,
    u: Key,
    now_serial: Option<f64>,
    budget: Rc<EvaluationBudget>,
) -> (Option<CellValue>, bool) {
    let Some(src) = wb.formula(u.0, cell_of(u)).map(str::to_string) else {
        return (None, false);
    };
    let Ok(expr) = parse_formula(&src) else {
        return (None, false);
    };
    let mut ctx = EvalContext::with_budget(wb, u.0, budget);
    ctx.now_serial = now_serial;
    let value = evaluate(&expr, &ctx);
    if !ctx.has_unhandled_budget_error() {
        return (Some(value), ctx.exhausted());
    }
    if matches!(wb.value(u.0, cell_of(u)), CellValue::Empty) {
        (Some(value), true)
    } else {
        (None, true)
    }
}

/// write `value` only if it differs from the stored value; returns whether
/// anything changed. formula and style are preserved.
fn write_if_changed(wb: &mut Workbook, u: Key, value: CellValue) -> bool {
    if wb.value(u.0, cell_of(u)) == value {
        return false;
    }
    if let Some(sheet) = wb.sheet_mut(u.0) {
        let mut cell = sheet.cell(cell_of(u)).cloned().unwrap_or_default();
        cell.value = value;
        sheet.set_cell(cell_of(u), cell);
    }
    true
}

fn sort_key(a: &(SheetId, CellRef), b: &(SheetId, CellRef)) -> std::cmp::Ordering {
    (a.0.0, a.1.row, a.1.col).cmp(&(b.0.0, b.1.row, b.1.col))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, Sheet};

    fn a1(s: &str) -> CellRef {
        CellRef::parse_a1(s).unwrap()
    }

    fn num(v: f64) -> CellValue {
        CellValue::Number { value: v }
    }

    /// set a literal number cell.
    fn put_num(wb: &mut Workbook, sheet: SheetId, cell: &str, v: f64) {
        wb.sheet_mut(sheet).unwrap().set_cell(
            a1(cell),
            Cell {
                value: num(v),
                ..Cell::default()
            },
        );
    }

    /// set a formula cell with an (initially blank) cached value.
    fn put_formula(wb: &mut Workbook, sheet: SheetId, cell: &str, f: &str) {
        wb.sheet_mut(sheet).unwrap().set_cell(
            a1(cell),
            Cell {
                value: CellValue::Empty,
                formula: Some(f.to_string()),
                style: None,
            },
        );
    }

    fn value(wb: &Workbook, sheet: SheetId, cell: &str) -> CellValue {
        wb.value(sheet, a1(cell))
    }

    fn changed_a1(r: &RecalcResult) -> Vec<String> {
        r.changed.iter().map(|(_, c)| c.to_a1()).collect()
    }

    fn one_sheet() -> (Workbook, SheetId) {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        (wb, SheetId(0))
    }

    #[test]
    fn chain_propagates_transitively() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 1.0);
        put_formula(&mut wb, s, "B1", "A1+1");
        put_formula(&mut wb, s, "C1", "B1+1");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(2.0));
        assert_eq!(value(&wb, s, "C1"), num(3.0));

        put_num(&mut wb, s, "A1", 10.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        assert_eq!(value(&wb, s, "B1"), num(11.0));
        assert_eq!(value(&wb, s, "C1"), num(12.0));
        assert_eq!(changed_a1(&r), vec!["B1", "C1"]);
    }

    #[test]
    fn diamond_evaluates_each_cell_once_in_order() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 5.0);
        put_formula(&mut wb, s, "B1", "A1*2");
        put_formula(&mut wb, s, "C1", "A1+3");
        put_formula(&mut wb, s, "D1", "B1+C1");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "D1"), num(18.0));

        put_num(&mut wb, s, "A1", 6.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        assert_eq!(value(&wb, s, "D1"), num(21.0));
        assert_eq!(changed_a1(&r), vec!["B1", "C1", "D1"]);
    }

    #[test]
    fn range_dependency_recalcs_on_interior_edit() {
        let (mut wb, s) = one_sheet();
        for (i, cell) in ["A1", "A2", "A3", "A4", "A5"].iter().enumerate() {
            put_num(&mut wb, s, cell, (i + 1) as f64);
        }
        put_formula(&mut wb, s, "B1", "SUM(A1:A10)");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(15.0));

        put_num(&mut wb, s, "A5", 100.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A5"))], None);
        assert_eq!(value(&wb, s, "B1"), num(110.0));
        assert_eq!(changed_a1(&r), vec!["B1"]);
    }

    #[test]
    fn cross_sheet_chain() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb.sheets.push(Sheet::new("Data"));
        let (s1, s2) = (SheetId(0), SheetId(1));
        put_num(&mut wb, s1, "A1", 7.0);
        put_formula(&mut wb, s2, "A1", "sheet1!A1 * 2");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s2, "A1"), num(14.0));

        put_num(&mut wb, s1, "A1", 8.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s1, a1("A1"))], None);
        assert_eq!(value(&wb, s2, "A1"), num(16.0));
        assert_eq!(r.changed, vec![(s2, a1("A1"))]);
    }

    #[test]
    fn cycle_zeros_cells_and_recovers_when_broken() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "B1+1");
        put_formula(&mut wb, s, "B1", "A1+1");
        let (mut graph, r) = rebuild_and_recalc_all(&mut wb, None);
        let mut cyc: Vec<String> = r.cycle_cells.iter().map(|(_, c)| c.to_a1()).collect();
        cyc.sort();
        assert_eq!(cyc, vec!["A1", "B1"]);
        assert_eq!(value(&wb, s, "A1"), num(0.0));
        assert_eq!(value(&wb, s, "B1"), num(0.0));

        put_formula(&mut wb, s, "B1", "5");
        graph.set_formula(s, a1("B1"), Some("5"));
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("B1"))], None);
        assert!(r.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "B1"), num(5.0));
        assert_eq!(value(&wb, s, "A1"), num(6.0));
    }

    #[test]
    fn self_reference_is_a_cycle() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "A1+1");
        let (_, r) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(r.cycle_cells, vec![(s, a1("A1"))]);
        assert_eq!(value(&wb, s, "A1"), num(0.0));
    }

    #[test]
    fn incremental_set_formula_updates_live_edges() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 1.0);
        put_num(&mut wb, s, "B1", 100.0);
        put_formula(&mut wb, s, "C1", "A1+1");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "C1"), num(2.0));

        put_formula(&mut wb, s, "C1", "B1+1");
        graph.set_formula(s, a1("C1"), Some("B1+1"));
        recalc_after(&mut wb, &mut graph, &[(s, a1("C1"))], None);
        assert_eq!(value(&wb, s, "C1"), num(101.0));

        put_num(&mut wb, s, "A1", 50.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        assert!(r.changed.is_empty());
        assert_eq!(value(&wb, s, "C1"), num(101.0));

        put_num(&mut wb, s, "B1", 200.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("B1"))], None);
        assert_eq!(changed_a1(&r), vec!["C1"]);
        assert_eq!(value(&wb, s, "C1"), num(201.0));
    }

    /// overwrite a cell's cached value while keeping its formula.
    fn set_cached(wb: &mut Workbook, sheet: SheetId, cell: &str, v: CellValue) {
        let mut c = wb.sheet(sheet).unwrap().cell(a1(cell)).cloned().unwrap();
        c.value = v;
        wb.sheet_mut(sheet).unwrap().set_cell(a1(cell), c);
    }

    #[test]
    fn volatile_cell_reevaluates_on_unrelated_edit() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 1.0);
        put_formula(&mut wb, s, "B1", "A1+1");
        put_formula(&mut wb, s, "C1", "NOW()");
        put_formula(&mut wb, s, "D1", "C1+1");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, Some(45000.0));

        let sentinel = num(-999_999.0);
        set_cached(&mut wb, s, "C1", sentinel.clone());
        set_cached(&mut wb, s, "D1", sentinel.clone());

        put_num(&mut wb, s, "A1", 2.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], Some(45000.0));
        assert_ne!(value(&wb, s, "C1"), sentinel);
        assert_ne!(value(&wb, s, "D1"), sentinel);
        assert_eq!(changed_a1(&r), vec!["B1", "C1", "D1"]);
    }

    #[test]
    fn changed_list_excludes_unmoved_dependents() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 1.0);
        put_num(&mut wb, s, "A2", 2.0);
        put_num(&mut wb, s, "A3", 3.0);
        put_formula(&mut wb, s, "B1", "MIN(A1:A3)");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(1.0));

        put_num(&mut wb, s, "A2", 5.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A2"))], None);
        assert!(r.changed.is_empty(), "unmoved MIN must not be reported");
        assert_eq!(value(&wb, s, "B1"), num(1.0));
    }

    #[test]
    fn rebuild_corrects_stale_cached_value() {
        let (mut wb, s) = one_sheet();
        wb.sheet_mut(s).unwrap().set_cell(
            a1("A1"),
            Cell {
                value: num(999.0),
                formula: Some("1+1".to_string()),
                style: None,
            },
        );
        let (_, r) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(2.0));
        assert_eq!(r.changed, vec![(s, a1("A1"))]);
    }

    #[test]
    fn exhausted_formula_budget_preserves_cached_value() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            a1("A1"),
            Cell {
                value: num(123.0),
                formula: Some("SUM(Data!A1:XFD1048576)".into()),
                style: None,
            },
        );
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, SheetId(1), "A1"), num(123.0));
        assert!(result.changed.is_empty());
        assert_eq!(result.limited_cells, vec![(SheetId(1), a1("A1"))]);
    }

    #[test]
    fn handled_budget_error_updates_the_cached_value() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.sheets.push(Sheet::new("Formula"));
        put_formula(
            &mut wb,
            SheetId(1),
            "A1",
            "IFERROR(SUM(Data!A1:XFD1048576),42)",
        );
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, SheetId(1), "A1"), num(42.0));
        assert_eq!(result.limited_cells, vec![(SheetId(1), a1("A1"))]);
    }

    #[test]
    fn handled_budget_error_can_produce_an_explicit_num_error() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            a1("A1"),
            Cell {
                value: num(123.0),
                formula: Some("IFERROR(SUM(Data!A1:XFD1048576),#NUM!)".into()),
                style: None,
            },
        );
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(
            value(&wb, SheetId(1), "A1"),
            CellValue::Error {
                value: xlsx_model::ErrorValue::Num
            }
        );
        assert_eq!(result.limited_cells, vec![(SheetId(1), a1("A1"))]);
    }

    #[test]
    #[ignore = "perf smoke; run with --release --ignored"]
    fn ten_thousand_cell_chain_is_fast() {
        use std::time::Instant;
        const N: u32 = 10_000;
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 0.0);
        for row in 1..N {
            let cell = CellRef::new(row, 0);
            let prev = CellRef::new(row - 1, 0).to_a1();
            wb.sheet_mut(s).unwrap().set_cell(
                cell,
                Cell {
                    value: CellValue::Empty,
                    formula: Some(format!("{prev}+1")),
                    style: None,
                },
            );
        }
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(wb.value(s, CellRef::new(N - 1, 0)), num((N - 1) as f64));

        let start = Instant::now();
        put_num(&mut wb, s, "A1", 1.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        let elapsed = start.elapsed();
        assert_eq!(wb.value(s, CellRef::new(N - 1, 0)), num(N as f64));
        assert_eq!(r.changed.len(), (N - 1) as usize);
        assert!(elapsed.as_secs_f64() < 1.0, "recalc took {elapsed:?}");
    }
}
