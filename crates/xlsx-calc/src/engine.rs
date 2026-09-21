//! recalc driver: given edited cells, re-evaluate exactly the formulas that
//! could have changed, in dependency order, and report what moved.

use std::borrow::Cow;
use std::cell::Cell as Flag;
use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::rc::Rc;

use xlsx_model::{
    Cell, CellProvider, CellRange, CellRef, CellValue, ColId, DefinedName, ErrorValue, RowId,
    SheetId, Table, Workbook,
};

use crate::array::{Spill, evaluate_spill, spill_at};
use crate::eval::{EvalContext, EvaluationBudget, MAX_RECALCULATION_CELL_VISITS, evaluate};
use crate::graph::DepGraph;

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
    let result = run_recalc(wb, graph, recompute, now_serial);
    graph.refresh_spills(wb);
    result
}

/// rebuild the graph from scratch and recalc every formula in dependency order.
pub fn rebuild_and_recalc_all(
    wb: &mut Workbook,
    now_serial: Option<f64>,
) -> (DepGraph, RecalcResult) {
    let mut graph = DepGraph::build(wb);
    let recompute: HashSet<Key> = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
    let result = run_recalc(wb, &graph, recompute, now_serial);
    graph.refresh_spills(wb);
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

/// a spilled rectangle that grows or shrinks uncovers cells whose readers were
/// not in the recompute set, so each round seeds the next. bounded so no
/// workbook can loop here.
const MAX_SPILL_ROUNDS: usize = 4;

/// cells one deferred pass will try to settle. a larger apparent cycle keeps
/// the range graph's answer rather than paying a quadratic scan for it.
const MAX_DEFERRED_CYCLE_CELLS: usize = 4096;

/// topologically order `recompute` and evaluate it, writing changed values into
/// `wb`. cells caught in a cycle are zeroed and reported separately.
fn run_recalc(
    wb: &mut Workbook,
    graph: &DepGraph,
    recompute: HashSet<Key>,
    now_serial: Option<f64>,
) -> RecalcResult {
    let budget = Rc::new(EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS));
    let mut changed: Vec<(SheetId, CellRef)> = Vec::new();
    let mut limited_cells = Vec::new();
    let mut cycle_cells: Vec<(SheetId, CellRef)> = Vec::new();
    let mut pending = recompute;

    for _ in 0..MAX_SPILL_ROUNDS {
        let (order, cycle) = topo_order(graph, &pending);
        let mut spilled: Vec<(SheetId, CellRef)> = Vec::new();
        for u in &order {
            let (value, limited) = eval_node(wb, *u, now_serial, Rc::clone(&budget), graph);
            if limited {
                limited_cells.push((u.0, cell_of(*u)));
            }
            match value {
                Some(NodeValue::Scalar(value)) => {
                    if write_if_changed(wb, *u, value) {
                        changed.push((u.0, cell_of(*u)));
                    }
                }
                Some(NodeValue::Spill(spill)) => {
                    spilled.extend(write_spill(wb, *u, spill, &mut changed));
                }
                None => {}
            }
        }
        let circular = settle_deferred(
            wb,
            &cycle,
            graph,
            now_serial,
            &budget,
            &mut changed,
            &mut limited_cells,
        );
        for u in &circular {
            if write_if_changed(wb, *u, CellValue::Number { value: 0.0 }) {
                changed.push((u.0, cell_of(*u)));
            }
            cycle_cells.push((u.0, cell_of(*u)));
        }
        if spilled.is_empty() {
            break;
        }
        pending = dependents_closure(graph, &spilled);
        if pending.is_empty() {
            break;
        }
    }

    changed.sort_by(sort_key);
    changed.dedup();
    RecalcResult {
        changed,
        cycle_cells,
        limited_cells,
    }
}

/// formula cells that transitively read any of `seeds`.
fn dependents_closure(graph: &DepGraph, seeds: &[(SheetId, CellRef)]) -> HashSet<Key> {
    let mut found: HashSet<Key> = HashSet::new();
    let mut expanded: HashSet<Key> = HashSet::new();
    let mut worklist: Vec<Key> = seeds.iter().map(|&(s, c)| key(s, c)).collect();
    while let Some(node) = worklist.pop() {
        if !expanded.insert(node) {
            continue;
        }
        for (sheet, cell) in graph.dependents_of(node.0, cell_of(node)) {
            let dependent = key(sheet, cell);
            if found.insert(dependent) {
                worklist.push(dependent);
            }
        }
    }
    found
}

/// kahn's sort over the sub-graph induced by `recompute`: returns the evaluable
/// order and, separately, the cells caught in (or only reachable through) a cycle.
fn topo_order(graph: &DepGraph, recompute: &HashSet<Key>) -> (Vec<Key>, Vec<Key>) {
    let mut points: HashMap<SheetId, BTreeMap<RowId, Vec<ColId>>> = HashMap::new();
    for &(sheet, row, col) in recompute {
        points
            .entry(sheet)
            .or_default()
            .entry(row)
            .or_default()
            .push(col);
    }

    let mut adj: HashMap<Key, Vec<Key>> = HashMap::new();
    let mut indegree: HashMap<Key, usize> = recompute.iter().map(|k| (*k, 0)).collect();
    let mut seen: HashSet<(Key, Key)> = HashSet::new();

    for (es, range, vs, vc) in graph.edges() {
        let v = key(vs, vc);
        if !recompute.contains(&v) {
            continue;
        }
        let Some(rows) = points.get(&es) else {
            continue;
        };
        for (&row, cols) in rows.range(range.start.row..=range.end.row) {
            for &col in cols {
                if col < range.start.col || col > range.end.col {
                    continue;
                }
                let u = (es, row, col);
                if seen.insert((u, v)) {
                    adj.entry(u).or_default().push(v);
                    *indegree.get_mut(&v).unwrap() += 1;
                }
            }
        }
        for (sheet, anchor) in graph.spill_sources(es, range) {
            let u = key(sheet, anchor);
            if u == v || !recompute.contains(&u) {
                continue;
            }
            if seen.insert((u, v)) {
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

/// what a formula node produced: one value, or a rectangle an array formula
/// fills from its anchor.
enum NodeValue {
    Scalar(CellValue),
    Spill(Spill),
}

/// evaluate one formula node; `None` keeps the cached value, because the cell
/// has no formula, it no longer parses, or an engine gap reached the result.
/// the range graph makes every cell of a range a precedent, so a table whose
/// column reads its own earlier rows through `INDEX(range, k)` looks circular
/// even though no cell reads itself. settle what the reads really allow: a
/// cell whose evaluation touches nothing still unsettled was never in a
/// cycle. what is left over is.
fn settle_deferred(
    wb: &mut Workbook,
    cycle: &[Key],
    graph: &DepGraph,
    now_serial: Option<f64>,
    budget: &Rc<EvaluationBudget>,
    changed: &mut Vec<(SheetId, CellRef)>,
    limited_cells: &mut Vec<(SheetId, CellRef)>,
) -> Vec<Key> {
    if cycle.is_empty() || cycle.len() > MAX_DEFERRED_CYCLE_CELLS {
        return cycle.to_vec();
    }
    let mut unsettled: HashSet<Key> = cycle.iter().copied().collect();
    for _ in 0..cycle.len() {
        let mut settled_any = false;
        for u in cycle {
            if !unsettled.contains(u) {
                continue;
            }
            let log = ReadLog {
                inner: wb,
                watch: &unsettled,
                touched: Flag::new(false),
            };
            let (value, limited) =
                eval_node_with(&log, wb, *u, now_serial, Rc::clone(budget), graph);
            // a spill rewrites a rectangle, which the ordered pass owns
            if log.touched.get() || matches!(value, Some(NodeValue::Spill(_))) {
                continue;
            }
            unsettled.remove(u);
            settled_any = true;
            if limited {
                limited_cells.push((u.0, cell_of(*u)));
            }
            if let Some(NodeValue::Scalar(value)) = value
                && write_if_changed(wb, *u, value)
            {
                changed.push((u.0, cell_of(*u)));
            }
        }
        if !settled_any {
            break;
        }
    }
    cycle
        .iter()
        .copied()
        .filter(|u| unsettled.contains(u))
        .collect()
}

/// a provider that notes whether a formula read any of the cells still
/// waiting to settle, without recording the rest.
struct ReadLog<'a> {
    inner: &'a Workbook,
    watch: &'a HashSet<Key>,
    touched: Flag<bool>,
}

impl ReadLog<'_> {
    fn note(&self, sheet: SheetId, at: CellRef) {
        if self.watch.contains(&key(sheet, at)) {
            self.touched.set(true);
        }
    }
}

impl CellProvider for ReadLog<'_> {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue {
        self.note(sheet, at);
        self.inner.value(sheet, at)
    }

    fn value_cow(&self, sheet: SheetId, at: CellRef) -> Cow<'_, CellValue> {
        self.note(sheet, at);
        self.inner.value_cow(sheet, at)
    }

    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str> {
        self.inner.formula(sheet, at)
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.inner.sheet_id(name)
    }

    fn defined_name(&self, sheet: SheetId, name: &str) -> Option<&DefinedName> {
        self.inner.defined_name(sheet, name)
    }

    fn table(&self, name: &str) -> Option<&Table> {
        self.inner.table(name)
    }

    fn used_rows(&self, sheet: SheetId) -> RowId {
        self.inner.used_rows(sheet)
    }

    fn used_cols(&self, sheet: SheetId) -> ColId {
        self.inner.used_cols(sheet)
    }

    fn spill_range(&self, sheet: SheetId, at: CellRef) -> Option<CellRange> {
        self.inner.spill_range(sheet, at)
    }
}

fn eval_node(
    wb: &Workbook,
    u: Key,
    now_serial: Option<f64>,
    budget: Rc<EvaluationBudget>,
    graph: &DepGraph,
) -> (Option<NodeValue>, bool) {
    eval_node_with(wb, wb, u, now_serial, budget, graph)
}

/// evaluate one node, reading cells through `provider` while `wb` supplies the
/// formula's own shape.
fn eval_node_with(
    provider: &dyn CellProvider,
    wb: &Workbook,
    u: Key,
    now_serial: Option<f64>,
    budget: Rc<EvaluationBudget>,
    graph: &DepGraph,
) -> (Option<NodeValue>, bool) {
    let Some(expr) = graph.ast(u.0, cell_of(u)) else {
        return (None, false);
    };
    let cell = cell_of(u);
    let authored = wb.sheet(u.0).and_then(|sheet| sheet.array_formula(cell));
    let mut ctx = EvalContext::with_budget(provider, u.0, budget);
    ctx.cell = Some(cell);
    ctx.now_serial = now_serial;
    ctx.parse_cache = Some(graph.asts());
    let value = match authored {
        Some(_) => NodeValue::Spill(evaluate_spill(&expr, &ctx, cell, authored)),
        None => NodeValue::Scalar(computed(evaluate(&expr, &ctx))),
    };
    let unsupported = ctx.has_unhandled_unsupported_function();
    let incomplete = ctx.has_unhandled_budget_error() || unsupported;
    if incomplete && !matches!(wb.value_cow(u.0, cell).as_ref(), CellValue::Empty) {
        return (None, ctx.exhausted());
    }
    // a rectangle an engine gap reshaped would retire cells the real result
    // still covers, so report the gap over the recorded rectangle instead.
    let value = match value {
        NodeValue::Spill(spill)
            if unsupported && authored.is_some_and(|range| range != spill.range) =>
        {
            let name = CellValue::Error {
                value: ErrorValue::Name,
            };
            NodeValue::Spill(spill_at(cell, authored, name.into()))
        }
        value => value,
    };
    (Some(value), ctx.exhausted())
}

/// a formula never leaves its cell blank: excel reads an empty cell as the
/// number zero, so a result that resolves to one shows `0`.
pub(crate) fn computed(value: CellValue) -> CellValue {
    match value {
        CellValue::Empty => CellValue::Number { value: 0.0 },
        value => value,
    }
}

/// lay a spilled result out from its anchor: retire the cells the previous
/// result reached and no longer fills, then write the new ones. returns the
/// spilled cells whose value moved, so their readers can be rescheduled.
fn write_spill(
    wb: &mut Workbook,
    u: Key,
    spill: Spill,
    changed: &mut Vec<(SheetId, CellRef)>,
) -> Vec<(SheetId, CellRef)> {
    let anchor = cell_of(u);
    let previous = wb.sheet(u.0).and_then(|sheet| sheet.array_formula(anchor));
    let (range, values) = if blocked(wb, u.0, anchor, spill.range, previous) {
        (
            CellRange::new(anchor, anchor),
            vec![CellValue::Error {
                value: ErrorValue::Spill,
            }],
        )
    } else {
        (spill.range, spill.values)
    };
    let mut moved = Vec::new();
    if let Some(previous) = previous {
        for (row, col) in cells_of(previous) {
            let at = CellRef::new(row, col);
            if at.row == anchor.row && at.col == anchor.col || range.contains(at) {
                continue;
            }
            write_spilled_cell(wb, (u.0, row, col), CellValue::Empty, changed, &mut moved);
        }
    }
    for ((row, col), value) in cells_of(range).zip(values) {
        if row == anchor.row && col == anchor.col {
            if write_if_changed(wb, (u.0, row, col), value) {
                changed.push((u.0, anchor));
            }
            continue;
        }
        write_spilled_cell(wb, (u.0, row, col), value, changed, &mut moved);
    }
    if let Some(sheet) = wb.sheet_mut(u.0) {
        sheet.set_array_formula(anchor, range);
    }
    moved
}

/// whether anything the author put in the way stops a result spilling. only
/// cells beyond the rectangle this anchor already records count: what is inside
/// it is the previous result, and a file may record one over its own formulas.
fn blocked(
    wb: &Workbook,
    sheet: SheetId,
    anchor: CellRef,
    range: CellRange,
    previous: Option<CellRange>,
) -> bool {
    if range.start == range.end {
        return false;
    }
    cells_of(range).any(|(row, col)| {
        let at = CellRef::new(row, col);
        if at.row == anchor.row && at.col == anchor.col {
            return false;
        }
        if previous.is_some_and(|previous| previous.contains(at)) {
            return false;
        }
        owns_formula(wb, sheet, at) || !matches!(wb.value_cow(sheet, at).as_ref(), CellValue::Empty)
    })
}

/// whether a cell carries a formula of its own. producers mark the interior
/// of a spilled rectangle with an empty `<f/>`, which is the anchor's formula
/// reaching that cell, not one the cell owns.
fn owns_formula(wb: &Workbook, sheet: SheetId, at: CellRef) -> bool {
    wb.formula(sheet, at)
        .is_some_and(|formula| !formula.trim().is_empty())
}

fn write_spilled_cell(
    wb: &mut Workbook,
    u: Key,
    value: CellValue,
    changed: &mut Vec<(SheetId, CellRef)>,
    moved: &mut Vec<(SheetId, CellRef)>,
) {
    if owns_formula(wb, u.0, cell_of(u)) {
        return;
    }
    if write_if_changed(wb, u, value) {
        changed.push((u.0, cell_of(u)));
        moved.push((u.0, cell_of(u)));
    }
}

fn cells_of(range: CellRange) -> impl Iterator<Item = (RowId, ColId)> {
    (range.start.row..=range.end.row)
        .flat_map(move |row| (range.start.col..=range.end.col).map(move |col| (row, col)))
}

/// write `value` only if it differs from the stored value; returns whether
/// anything changed. formula and style are preserved.
fn write_if_changed(wb: &mut Workbook, u: Key, value: CellValue) -> bool {
    if *wb.value_cow(u.0, cell_of(u)) == value {
        return false;
    }
    if let Some(sheet) = wb.sheet_mut(u.0) {
        let at = cell_of(u);
        if let Some(cell) = sheet.cell_mut(at) {
            cell.value = value;
            if *cell == Cell::default() {
                sheet.set_cell(at, Cell::default());
            }
        } else {
            sheet.set_cell(
                at,
                Cell {
                    value,
                    ..Cell::default()
                },
            );
        }
    }
    true
}

fn sort_key(a: &(SheetId, CellRef), b: &(SheetId, CellRef)) -> std::cmp::Ordering {
    (a.0.0, a.1.row, a.1.col).cmp(&(b.0.0, b.1.row, b.1.col))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, DefinedName, Sheet};

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

    /// a formula cell that already carries its authoring app's value.
    fn put_cached_formula(wb: &mut Workbook, sheet: SheetId, cell: &str, f: &str, v: CellValue) {
        wb.sheet_mut(sheet).unwrap().set_cell(
            a1(cell),
            Cell {
                value: v,
                formula: Some(f.to_string()),
                style: None,
            },
        );
    }

    /// producers mark the interior of a spilled rectangle with an empty
    /// `<f/>`; that is the anchor's formula reaching the cell, so the result
    /// still lands there.
    #[test]
    fn an_empty_formula_element_does_not_block_a_spill() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 5.0);
        put_num(&mut wb, s, "A2", 7.0);
        wb.sheet_mut(s).unwrap().set_cell(
            a1("C1"),
            Cell {
                value: CellValue::Empty,
                formula: Some("A1:A2".into()),
                style: None,
            },
        );
        wb.sheet_mut(s).unwrap().set_cell(
            a1("C2"),
            Cell {
                value: CellValue::Empty,
                formula: Some(String::new()),
                style: None,
            },
        );
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("C1"), xlsx_model::CellRange::parse_a1("C1:C2").unwrap());
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "C2"), num(7.0));
    }

    /// excel reads a blank cell as zero, so a formula that resolves to one
    /// shows `0` rather than leaving its own cell blank.
    #[test]
    fn a_formula_reading_a_blank_cell_shows_zero() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "B1", "A1");
        put_formula(&mut wb, s, "B2", "A1&\"\"");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(0.0));
        assert_eq!(
            value(&wb, s, "B2"),
            CellValue::Text {
                value: String::new()
            }
        );
    }

    /// the same rule inside a spill: the positions the block reaches show `0`
    /// for a blank source, the ones beyond it stay blank.
    #[test]
    fn a_spilled_blank_shows_zero() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 5.0);
        put_num(&mut wb, s, "A3", 7.0);
        wb.sheet_mut(s).unwrap().set_cell(
            a1("C1"),
            Cell {
                value: CellValue::Empty,
                formula: Some("A1:A3".into()),
                style: None,
            },
        );
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("C1"), xlsx_model::CellRange::parse_a1("C1:C3").unwrap());
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "C1"), num(5.0));
        assert_eq!(value(&wb, s, "C2"), num(0.0));
        assert_eq!(value(&wb, s, "C3"), num(7.0));
    }

    /// `WEBSERVICE` stands in for any function the engine does not implement,
    /// and is one it never will.
    #[test]
    fn an_unimplemented_function_keeps_the_cached_value() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_cached_formula(&mut wb, s, "B1", "WEBSERVICE(A1)/10", num(4.9));
        put_formula(&mut wb, s, "C1", "B1*2");
        let r = rebuild_and_recalc_all(&mut wb, None).1;
        assert_eq!(value(&wb, s, "B1"), num(4.9));
        assert_eq!(value(&wb, s, "C1"), num(9.8));
        assert!(!changed_a1(&r).contains(&"B1".to_string()));
    }

    #[test]
    fn an_unimplemented_function_without_a_cached_value_reports_name() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "WEBSERVICE(1)");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(
            value(&wb, s, "A1"),
            CellValue::Error {
                value: xlsx_model::ErrorValue::Name
            }
        );
    }

    /// the array form of a supported function is still a gap in the engine, so
    /// it must leave the cached value alone exactly as an unknown name does.
    /// `nanogpt-excel` reaches this through
    /// `SUMPRODUCT(OFFSET(...), TRANSPOSE(OFFSET(...)))`.
    #[test]
    fn an_array_result_the_engine_cannot_represent_keeps_the_cached_value() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 2.0);
        put_num(&mut wb, s, "A2", 3.0);
        put_cached_formula(&mut wb, s, "B1", "SUMPRODUCT(TRANSPOSE(A1:A2))", num(5.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(5.0));
    }

    /// IFERROR answers for the call it wraps, so the gap never reaches the
    /// result and the computed value must replace the cache.
    #[test]
    fn a_handler_that_answers_an_unimplemented_call_writes_its_result() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "B1", "IFERROR(WEBSERVICE(1),0)", num(99.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(0.0));
    }

    #[test]
    fn a_handler_beside_an_unimplemented_call_still_keeps_the_cache() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 2.0);
        put_cached_formula(&mut wb, s, "B1", "IFERROR(A1,0)+WEBSERVICE(1)", num(99.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(99.0));
    }

    /// A memoized name must replay the gap its evaluation recorded, or a
    /// second use returns the memo without it and the cache is overwritten.
    #[test]
    fn a_handled_gap_behind_a_defined_name_still_marks_its_second_use() {
        let (mut wb, s) = one_sheet();
        wb.defined_names.push(xlsx_model::DefinedName {
            name: "Gap".to_string(),
            formula: "WEBSERVICE(1)".to_string(),
            local_sheet: None,
            hidden: false,
        });
        put_cached_formula(&mut wb, s, "B1", "IFERROR(Gap,0)+Gap", num(99.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(99.0));
    }

    #[test]
    fn a_supported_function_still_overwrites_a_stale_cached_value() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 2.0);
        put_cached_formula(&mut wb, s, "B1", "SUM(A1,A1)", num(99.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(4.0));
    }

    /// OFFSET reads its anchor's coordinates, not its value, so a cell may
    /// offset from itself: Greptile flagged `A1=SUM(OFFSET(A1,1,0,3,1))`
    /// reporting A1 as cyclic and zeroing a valid result.
    #[test]
    fn a_cell_may_offset_from_its_own_position() {
        let (mut wb, s) = one_sheet();
        for (cell, v) in [("A2", 1.0), ("A3", 2.0), ("A4", 3.0)] {
            put_num(&mut wb, s, cell, v);
        }
        put_formula(&mut wb, s, "A1", "SUM(OFFSET(A1,1,0,3,1))");
        let report = rebuild_and_recalc_all(&mut wb, None).1;
        assert!(report.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "A1"), num(6.0));
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
    fn static_offset_recalcs_when_its_target_changes() {
        let (mut wb, s) = one_sheet();
        for (i, cell) in ["A1", "A2", "A3", "A4"].iter().enumerate() {
            put_num(&mut wb, s, cell, (i + 1) as f64);
        }
        put_formula(&mut wb, s, "B1", "SUM(OFFSET(A1, 1, 0, 3, 1))");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(9.0));

        put_num(&mut wb, s, "A3", 100.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A3"))], None);
        assert_eq!(value(&wb, s, "B1"), num(106.0));
        assert_eq!(changed_a1(&r), vec!["B1"]);
    }

    #[test]
    fn unresolvable_offset_recalcs_when_its_target_changes() {
        let (mut wb, s) = one_sheet();
        for (i, cell) in ["A1", "A2", "A3"].iter().enumerate() {
            put_num(&mut wb, s, cell, (i + 1) as f64);
        }
        put_num(&mut wb, s, "D1", 2.0);
        put_formula(&mut wb, s, "B1", "OFFSET(A1, D1, 0)");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(3.0));

        put_num(&mut wb, s, "A3", 30.0);
        recalc_after(&mut wb, &mut graph, &[(s, a1("A3"))], None);
        assert_eq!(value(&wb, s, "B1"), num(30.0));

        put_num(&mut wb, s, "D1", 1.0);
        recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        assert_eq!(value(&wb, s, "B1"), num(2.0));
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
    fn defined_name_range_recalculates_from_its_dependencies() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheets.push(Sheet::new("Data"));
        wb.defined_names.push(DefinedName {
            name: "Inputs".into(),
            formula: "Data!$A$1:$A$3".into(),
            local_sheet: None,
            hidden: false,
        });
        for (cell, value) in [("A1", 1.0), ("A2", 2.0), ("A3", 3.0)] {
            put_num(&mut wb, SheetId(1), cell, value);
        }
        put_formula(&mut wb, SheetId(0), "A1", "SUM(Inputs)");

        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, SheetId(0), "A1"), num(6.0));

        put_num(&mut wb, SheetId(1), "A2", 20.0);
        let result = recalc_after(&mut wb, &mut graph, &[(SheetId(1), a1("A2"))], None);
        assert_eq!(value(&wb, SheetId(0), "A1"), num(24.0));
        assert_eq!(result.changed, vec![(SheetId(0), a1("A1"))]);
    }

    /// recalc runs on open, so a `Chain_0=Chain_1+Chain_1` ladder in an
    /// untrusted workbook must not cost 2^n expansions.
    #[test]
    fn doubling_defined_name_chain_recalculates_without_blowing_up() {
        const DEPTH: usize = 60;
        let (mut wb, sheet) = one_sheet();
        for index in 0..DEPTH {
            let next = index + 1;
            wb.defined_names.push(DefinedName {
                name: format!("Chain_{index}"),
                formula: format!("Chain_{next}+Chain_{next}"),
                local_sheet: None,
                hidden: false,
            });
        }
        wb.defined_names.push(DefinedName {
            name: format!("Chain_{DEPTH}"),
            formula: "1".into(),
            local_sheet: None,
            hidden: false,
        });
        put_formula(&mut wb, sheet, "A1", "Chain_0");

        let (_, result) = rebuild_and_recalc_all(&mut wb, None);

        assert_eq!(value(&wb, sheet, "A1"), num(2f64.powi(DEPTH as i32)));
        assert!(result.limited_cells.is_empty());
    }

    #[test]
    fn unknown_defined_name_replaces_stale_cache_with_name_error() {
        let (mut wb, sheet) = one_sheet();
        wb.sheet_mut(sheet).unwrap().set_cell(
            a1("A1"),
            Cell {
                value: num(99.0),
                formula: Some("MissingName+1".into()),
                style: None,
            },
        );

        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(
            value(&wb, sheet, "A1"),
            CellValue::Error {
                value: xlsx_model::ErrorValue::Name
            }
        );
        assert_eq!(result.changed, vec![(sheet, a1("A1"))]);
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

    /// the range graph makes `INDEX(range, k)` a precedent of every cell of
    /// `range`, so a column reading its own earlier rows looks circular. what
    /// the cells actually read says otherwise.
    #[test]
    fn a_range_wide_precedent_is_not_a_cycle_when_the_reads_are_not() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 5.0);
        put_formula(&mut wb, s, "B1", "A1");
        put_formula(&mut wb, s, "B2", "INDEX(B1:B3,1)+1");
        put_formula(&mut wb, s, "B3", "INDEX(B1:B3,2)+1");
        let (_, r) = rebuild_and_recalc_all(&mut wb, None);
        assert!(r.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "B1"), num(5.0));
        assert_eq!(value(&wb, s, "B2"), num(6.0));
        assert_eq!(value(&wb, s, "B3"), num(7.0));
    }

    /// settling the cells whose reads allow it leaves the ones whose reads do
    /// not, so a real cycle beside a false one is still reported.
    #[test]
    fn a_real_cycle_beside_a_false_one_is_still_reported() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 2.0);
        put_formula(&mut wb, s, "B1", "A1");
        put_formula(&mut wb, s, "B2", "INDEX(B1:B3,1)+1");
        put_formula(&mut wb, s, "B3", "INDEX(B1:B3,2)+1");
        put_formula(&mut wb, s, "D1", "D2+1");
        put_formula(&mut wb, s, "D2", "D1+1");
        let (_, r) = rebuild_and_recalc_all(&mut wb, None);
        let mut cyc: Vec<String> = r.cycle_cells.iter().map(|(_, c)| c.to_a1()).collect();
        cyc.sort();
        assert_eq!(cyc, vec!["D1", "D2"]);
        assert_eq!(value(&wb, s, "B3"), num(4.0));
        assert_eq!(value(&wb, s, "D1"), num(0.0));
    }

    /// a settled cell feeds the ones that read it, however deep the chain,
    /// and an edit still propagates through it afterwards.
    #[test]
    fn settling_walks_a_chain_and_survives_a_later_edit() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 1.0);
        put_formula(&mut wb, s, "B1", "A1");
        for row in 2..=8 {
            put_formula(
                &mut wb,
                s,
                &format!("B{row}"),
                &format!("INDEX(B1:B8,{})+1", row - 1),
            );
        }
        let (mut graph, r) = rebuild_and_recalc_all(&mut wb, None);
        assert!(r.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "B8"), num(8.0));

        put_num(&mut wb, s, "A1", 10.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        assert!(r.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "B8"), num(17.0));
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
    fn referenceless_row_and_column_resolve_against_the_calling_cell() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "C7", "ROW()");
        put_formula(&mut wb, s, "D8", "COLUMN()");
        put_formula(&mut wb, s, "E9", "ROW()*100+COLUMN()");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "C7"), num(7.0));
        assert_eq!(value(&wb, s, "D8"), num(4.0));
        assert_eq!(value(&wb, s, "E9"), num(905.0));
    }

    #[test]
    fn row_of_a_reference_is_positional_not_a_read() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "X1", "ROW($X$1)+COLUMN($X$1)+ROWS($X$1:$X$4)");
        put_num(&mut wb, s, "A1", 1.0);
        put_formula(&mut wb, s, "B1", "ROW()-ROW($A$1)");
        let (_, r) = rebuild_and_recalc_all(&mut wb, None);
        assert!(r.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "X1"), num(1.0 + 24.0 + 4.0));
        assert_eq!(value(&wb, s, "B1"), num(0.0));
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
    fn randbetween_redraws_on_every_recalc() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 1.0);
        put_formula(&mut wb, s, "B1", "RANDBETWEEN(1, 1000000)");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(graph.volatile_cells().count(), 1);

        let sentinel = num(-1.0);
        set_cached(&mut wb, s, "B1", sentinel.clone());
        put_num(&mut wb, s, "A1", 2.0);
        let r = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        assert_eq!(changed_a1(&r), vec!["B1"]);
        match value(&wb, s, "B1") {
            CellValue::Number { value } => assert!((1.0..=1_000_000.0).contains(&value)),
            other => panic!("expected a number, got {other:?}"),
        }
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
        // a cell in the far corner gives Data the extent the guard is for
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            a1("XFD1048576"),
            Cell {
                value: num(1.0),
                formula: None,
                style: None,
            },
        );
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
        // a cell in the far corner gives Data the extent the guard is for
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            a1("XFD1048576"),
            Cell {
                value: num(1.0),
                formula: None,
                style: None,
            },
        );
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
        // a cell in the far corner gives Data the extent the guard is for
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            a1("XFD1048576"),
            Cell {
                value: num(1.0),
                formula: None,
                style: None,
            },
        );
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
