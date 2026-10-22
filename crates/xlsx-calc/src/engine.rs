//! recalc driver: given edited cells, re-evaluate exactly the formulas that
//! could have changed, in dependency order, and report what moved.

mod order;

use order::{Step, topo_order};

use std::borrow::Cow;
use std::cell::Cell as Flag;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use xlsx_model::progress::tick as progress;
use xlsx_model::{
    Cell, CellProvider, CellRange, CellRef, CellValue, ColId, DefinedName, ErrorValue, RowId,
    SheetId, Table, Workbook,
};

use crate::array::{Spill, charged_spill, evaluate_spill};
use crate::eval::{EvalContext, EvaluationBudget, MAX_RECALCULATION_CELL_VISITS, evaluate};
use crate::graph::{DepGraph, ReaderIndex};

#[cfg(feature = "test-counters")]
thread_local! {
    static EVALUATED_FORMULAS: std::cell::RefCell<Option<Vec<(SheetId, CellRef)>>> =
        const { std::cell::RefCell::new(None) };
}

/// start recording the formula cells recalculation evaluates on this thread.
#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn track_formula_evaluations() {
    EVALUATED_FORMULAS.with(|cells| *cells.borrow_mut() = Some(Vec::new()));
}

/// the cells evaluated since [`track_formula_evaluations`]; stops recording.
#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn take_evaluated_formulas() -> Vec<(SheetId, CellRef)> {
    EVALUATED_FORMULAS.with(|cells| cells.borrow_mut().take().unwrap_or_default())
}

/// the outcome of a recalc: cells whose displayed value changed, and cells
/// settled by cycle participation rather than by their own formula.
pub struct RecalcResult {
    pub changed: Vec<(SheetId, CellRef)>,
    pub cycle_cells: Vec<(SheetId, CellRef)>,
    pub limited_cells: Vec<(SheetId, CellRef)>,
}

#[derive(Clone, Copy)]
struct RecalcContext {
    now_serial: Option<f64>,
    rand_seed: Option<u32>,
}

/// registers a function that parsing, building a workbook's authority and
/// recalculation call every few thousand units of work, so a host blocked on
/// a long one can report that it is alive.
pub use xlsx_model::progress::set_hook as set_progress_hook;

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
    recalc_after_with_seed(wb, graph, dirty_seeds, now_serial, None)
}

pub fn recalc_after_with_seed(
    wb: &mut Workbook,
    graph: &mut DepGraph,
    dirty_seeds: &[(SheetId, CellRef)],
    now_serial: Option<f64>,
    rand_seed: Option<u32>,
) -> RecalcResult {
    let context = RecalcContext {
        now_serial,
        rand_seed,
    };
    recalc_seeded(wb, graph, dirty_seeds, context, None)
}

/// [`recalc_after_with_seed`], also returning every formula cell it
/// re-evaluated, in address order: the cells its cycle and limited lists
/// speak for.
pub fn recalc_after_with_recomputed(
    wb: &mut Workbook,
    graph: &mut DepGraph,
    dirty_seeds: &[(SheetId, CellRef)],
    now_serial: Option<f64>,
    rand_seed: Option<u32>,
) -> (RecalcResult, Vec<(SheetId, CellRef)>) {
    let context = RecalcContext {
        now_serial,
        rand_seed,
    };
    let mut recomputed = HashSet::new();
    let result = recalc_seeded(wb, graph, dirty_seeds, context, Some(&mut recomputed));
    let mut recomputed: Vec<Key> = recomputed.into_iter().collect();
    recomputed.sort_unstable();
    let cells = recomputed.iter().map(|&k| (k.0, cell_of(k)));
    (result, cells.collect())
}

fn recalc_seeded(
    wb: &mut Workbook,
    graph: &mut DepGraph,
    dirty_seeds: &[(SheetId, CellRef)],
    context: RecalcContext,
    recomputed: Option<&mut HashSet<Key>>,
) -> RecalcResult {
    let recompute = collect_recompute(graph, dirty_seeds);
    let result = run_recalc(wb, graph, recompute, context, recomputed);
    graph.refresh_spills(wb);
    result
}

/// rebuild the graph from scratch and recalc every formula in dependency order.
pub fn rebuild_and_recalc_all(
    wb: &mut Workbook,
    now_serial: Option<f64>,
) -> (DepGraph, RecalcResult) {
    rebuild_and_recalc_all_with_seed(wb, now_serial, None)
}

pub fn rebuild_and_recalc_all_with_seed(
    wb: &mut Workbook,
    now_serial: Option<f64>,
    rand_seed: Option<u32>,
) -> (DepGraph, RecalcResult) {
    let mut graph = DepGraph::build(wb);
    let recompute: HashSet<Key> = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
    let context = RecalcContext {
        now_serial,
        rand_seed,
    };
    let result = run_recalc(wb, &graph, recompute, context, None);
    xlsx_model::progress::pulse();
    graph.refresh_spills(wb);
    xlsx_model::progress::pulse();
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
    let mut readers = Readers::new(graph);
    let mut found = Vec::new();
    while let Some(k) = worklist.pop() {
        if !expanded.insert(k) {
            continue;
        }
        progress();
        readers.dependents_of(k, &mut found);
        for (ds, dc) in found.drain(..) {
            let dk = key(ds, dc);
            recompute.insert(dk);
            worklist.push(dk);
        }
    }
    recompute
}

/// lookups a walk makes into one sheet by scanning its edges before it
/// indexes that sheet.
const SCANNED_LOOKUPS: usize = 16;

/// the readers of a cell for a walk through the graph: a scan per lookup for
/// short walks, and one [`ReaderIndex`] per sheet the walk keeps returning to,
/// so a chain of `n` formulas costs `O(n + edges)` rather than `O(n * edges)`.
struct Readers<'g> {
    graph: &'g DepGraph,
    sheets: HashMap<SheetId, SheetReaders<'g>>,
}

enum SheetReaders<'g> {
    Scanned(usize),
    Indexed(ReaderIndex<'g>),
}

impl<'g> Readers<'g> {
    fn new(graph: &'g DepGraph) -> Self {
        Self {
            graph,
            sheets: HashMap::new(),
        }
    }

    fn dependents_of(&mut self, k: Key, out: &mut Vec<(SheetId, CellRef)>) {
        let readers = self.sheets.entry(k.0).or_insert(SheetReaders::Scanned(0));
        if let SheetReaders::Scanned(lookups) = readers {
            *lookups += 1;
            if *lookups > SCANNED_LOOKUPS {
                *readers = SheetReaders::Indexed(self.graph.reader_index(k.0));
            }
        }
        match readers {
            SheetReaders::Indexed(index) => index.dependents_of(cell_of(k), out),
            SheetReaders::Scanned(_) => out.extend(self.graph.dependents_of(k.0, cell_of(k))),
        }
    }
}

/// a spilled rectangle that grows or shrinks uncovers cells whose readers were
/// not in the recompute set, so each round seeds the next. bounded so no
/// workbook can loop here.
const MAX_SPILL_ROUNDS: usize = 4;

/// topologically order `recompute` and evaluate it, writing changed values into
/// `wb`. cells caught in a cycle are settled and reported separately. every
/// cell a round takes up joins `recomputed`.
fn run_recalc(
    wb: &mut Workbook,
    graph: &DepGraph,
    recompute: HashSet<Key>,
    context: RecalcContext,
    mut recomputed: Option<&mut HashSet<Key>>,
) -> RecalcResult {
    let mut pass = Pass {
        graph,
        context,
        budget: Rc::new(EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS)),
        changed: Vec::new(),
        cycle_cells: Vec::new(),
        limited_cells: Vec::new(),
        spilled: Vec::new(),
    };
    let mut pending = recompute;

    for _ in 0..MAX_SPILL_ROUNDS {
        if let Some(recomputed) = &mut recomputed {
            recomputed.extend(pending.iter().copied());
        }
        for step in topo_order(graph, &pending) {
            match step {
                Step::Cell(u) => pass.evaluate(wb, u),
                Step::Component(members) => pass.settle(wb, &members),
            }
        }
        if pass.spilled.is_empty() {
            break;
        }
        pending = dependents_closure(graph, &std::mem::take(&mut pass.spilled));
        if pending.is_empty() {
            break;
        }
    }

    let Pass {
        mut changed,
        mut cycle_cells,
        limited_cells,
        ..
    } = pass;
    changed.sort_by(sort_key);
    changed.dedup();
    cycle_cells.sort_by(sort_key);
    cycle_cells.dedup();
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
    let mut readers = Readers::new(graph);
    let mut dependents = Vec::new();
    while let Some(node) = worklist.pop() {
        if !expanded.insert(node) {
            continue;
        }
        progress();
        readers.dependents_of(node, &mut dependents);
        for (sheet, cell) in dependents.drain(..) {
            let dependent = key(sheet, cell);
            if found.insert(dependent) {
                worklist.push(dependent);
            }
        }
    }
    found
}

#[cfg(test)]
thread_local! {
    /// seeks retiring spilled cells took, so a test can see the search skip
    /// the empty positions of a rectangle.
    static RETIREMENT_SEEKS: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
    /// formula evaluations a recalc ran, the ones it discarded included.
    static FORMULA_EVALUATIONS: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
    /// waits settling components registered, plus the members their search
    /// for cycles of waits visited.
    static SCHEDULER_STEPS: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
    /// reads settling components made, and the spilled rectangles those reads
    /// were compared against.
    static SETTLING_READS: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
    static SPILL_CHECKS: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
}

/// what a formula node produced: one value, or a rectangle an array formula
/// fills from its anchor.
enum NodeValue {
    Scalar(CellValue),
    Spill(Spill),
    /// an array formula the budget refused: its anchor shows `#NUM!`, and its
    /// rectangle keeps the extent and the cells it had.
    Refused,
}

/// one recalc: the budget its evaluations share and what they found.
struct Pass<'g> {
    graph: &'g DepGraph,
    context: RecalcContext,
    budget: Rc<EvaluationBudget>,
    changed: Vec<(SheetId, CellRef)>,
    cycle_cells: Vec<(SheetId, CellRef)>,
    limited_cells: Vec<(SheetId, CellRef)>,
    /// spilled cells whose value moved; the next round runs their readers.
    spilled: Vec<(SheetId, CellRef)>,
}

impl Pass<'_> {
    /// evaluate a formula no cycle of the range graph passes through.
    fn evaluate(&mut self, wb: &mut Workbook, u: Key) {
        #[cfg(test)]
        FORMULA_EVALUATIONS.with(|count| count.set(count.get() + 1));
        progress();
        let (value, limited) = eval_node(wb, u, self.context, Rc::clone(&self.budget), self.graph);
        self.write(wb, u, value, limited);
    }

    /// settle the members of one component of the range graph by what they
    /// actually read. a member that reads another before it settles waits on
    /// it: that one runs next, and the waiting member runs again once it has
    /// settled. the members waiting form one chain, so a member that reads
    /// one on the chain closes a cycle of reads: the members from that one on
    /// are circular and settle at `circular_value`, and the member waiting on
    /// the first of them runs again with that value.
    fn settle(&mut self, wb: &mut Workbook, members: &[Key]) {
        let ids: HashMap<Key, usize> = members.iter().enumerate().map(|(id, &u)| (u, id)).collect();
        let spills = Spills::new(wb, members);
        let mut settled = vec![false; members.len()];
        let mut on_chain: Vec<Option<usize>> = vec![None; members.len()];
        let mut chain: Vec<usize> = Vec::new();
        for start in 0..members.len() {
            if settled[start] {
                continue;
            }
            on_chain[start] = Some(0);
            chain.push(start);
            while let Some(&id) = chain.last() {
                let u = members[id];
                let log = ReadLog {
                    inner: wb,
                    ids: &ids,
                    spills: &spills,
                    settled: &settled,
                    first: Flag::new(None),
                };
                #[cfg(test)]
                FORMULA_EVALUATIONS.with(|count| count.set(count.get() + 1));
                progress();
                let (value, limited) = eval_node_with(
                    &log,
                    wb,
                    u,
                    self.context,
                    Rc::clone(&self.budget),
                    self.graph,
                );
                let Some(target) = log.first.get() else {
                    self.write(wb, u, value, limited);
                    settled[id] = true;
                    on_chain[id] = None;
                    chain.pop();
                    continue;
                };
                #[cfg(test)]
                SCHEDULER_STEPS.with(|count| count.set(count.get() + 1));
                let Some(at) = on_chain[target] else {
                    on_chain[target] = Some(chain.len());
                    chain.push(target);
                    continue;
                };
                for member in chain.drain(at..) {
                    #[cfg(test)]
                    SCHEDULER_STEPS.with(|count| count.set(count.get() + 1));
                    let u = members[member];
                    if write_if_changed(wb, u, circular_value(wb, u)) {
                        self.changed.push((u.0, cell_of(u)));
                    }
                    self.cycle_cells.push((u.0, cell_of(u)));
                    settled[member] = true;
                    on_chain[member] = None;
                }
            }
        }
    }

    fn write(&mut self, wb: &mut Workbook, u: Key, value: Option<NodeValue>, limited: bool) {
        let at = (u.0, cell_of(u));
        if limited {
            self.limited_cells.push(at);
        }
        match value {
            Some(NodeValue::Scalar(value)) => {
                if write_if_changed(wb, u, value) {
                    self.changed.push(at);
                }
            }
            Some(NodeValue::Spill(spill)) => {
                match write_spill(wb, u, spill, &self.budget, &mut self.changed) {
                    Some(moved) => self.spilled.extend(moved),
                    None => {
                        if !limited {
                            self.limited_cells.push(at);
                        }
                        refuse(wb, u, &mut self.changed);
                    }
                }
            }
            Some(NodeValue::Refused) => refuse(wb, u, &mut self.changed),
            None => {}
        }
    }
}

/// a provider that notes the first unsettled member of the settling component
/// a formula reads, through the member's own cell or the rectangle it spills
/// over. what the formula reads after that may follow from the stale value, so
/// only the first such read is certain to be needed.
struct ReadLog<'a> {
    inner: &'a Workbook,
    ids: &'a HashMap<Key, usize>,
    spills: &'a Spills,
    settled: &'a [bool],
    first: Flag<Option<usize>>,
}

impl ReadLog<'_> {
    fn note(&self, sheet: SheetId, at: CellRef) {
        #[cfg(test)]
        SETTLING_READS.with(|count| count.set(count.get() + 1));
        if self.first.get().is_some() {
            return;
        }
        let read = self
            .ids
            .get(&key(sheet, at))
            .copied()
            .filter(|&id| !self.settled[id])
            .or_else(|| self.spills.holding(sheet, at, self.settled));
        self.first.set(read);
    }
}

/// a grid block of one size class: sheet, size class, row block, column block.
type Block = (SheetId, u32, u32, u32, u32);

/// the rectangles the members of a component spill over. one between `2^j`
/// and `2^(j+1)` rows high and between `2^k` and `2^(k+1)` columns wide is
/// filed under the block of a `2^j` by `2^k` grid that holds its top-left
/// cell, so a cell lies in at most nine blocks a rectangle of that size can
/// start in, and two rectangles in one block overlap.
struct Spills {
    sizes: Vec<(u32, u32)>,
    blocks: HashMap<Block, Vec<(usize, CellRange)>>,
}

impl Spills {
    fn new(wb: &Workbook, members: &[Key]) -> Self {
        let rectangles: Vec<(usize, SheetId, CellRange)> = members
            .iter()
            .enumerate()
            .filter_map(|(id, &u)| {
                let range = wb.sheet(u.0)?.array_formula(cell_of(u))?;
                (range.start != range.end).then_some((id, u.0, range))
            })
            .collect();
        let mut spills = Self {
            sizes: Vec::new(),
            blocks: HashMap::new(),
        };
        for (id, sheet, range) in rectangles {
            let size = (
                (range.end.row - range.start.row).saturating_add(1).ilog2(),
                (range.end.col - range.start.col).saturating_add(1).ilog2(),
            );
            if !spills.sizes.contains(&size) {
                spills.sizes.push(size);
            }
            let block = (
                sheet,
                size.0,
                size.1,
                range.start.row >> size.0,
                range.start.col >> size.1,
            );
            spills.blocks.entry(block).or_default().push((id, range));
        }
        spills
    }

    /// the first member, in key order, not settled yet whose rectangle holds
    /// `at`.
    fn holding(&self, sheet: SheetId, at: CellRef, settled: &[bool]) -> Option<usize> {
        let mut first: Option<usize> = None;
        for &(rows, cols) in &self.sizes {
            for row in starts(at.row, rows) {
                for col in starts(at.col, cols) {
                    let Some(spills) = self.blocks.get(&(sheet, rows, cols, row, col)) else {
                        continue;
                    };
                    for &(id, range) in spills {
                        #[cfg(test)]
                        SPILL_CHECKS.with(|count| count.set(count.get() + 1));
                        if range.contains(at)
                            && !settled[id]
                            && first.is_none_or(|found| id < found)
                        {
                            first = Some(id);
                        }
                    }
                }
            }
        }
        first
    }
}

/// the blocks of a `2^size` grid where a span shorter than `2^(size+1)` that
/// holds `at` can start.
fn starts(at: u32, size: u32) -> std::ops::RangeInclusive<u32> {
    let reach = (2u64 << size) - 2;
    let first = u64::from(at).saturating_sub(reach) >> size;
    first as u32..=at >> size
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

fn mix_random_seed(mut state: u64) -> u64 {
    state = state.wrapping_add(0x9E37_79B9_7F4A_7C15);
    state = (state ^ (state >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
    state = (state ^ (state >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
    state ^ (state >> 31)
}

fn cell_random_seed(seed: u32, cell: Key) -> u64 {
    let seed_sheet = (u64::from(seed) << 32) | u64::from(cell.0.0);
    let row_col = (u64::from(cell.1) << 32) | u64::from(cell.2);
    mix_random_seed(mix_random_seed(seed_sheet) ^ row_col)
}

/// evaluate one formula node; `None` keeps the cached value, because the cell
/// has no formula, it no longer parses, or an engine gap reached the result.
fn eval_node(
    wb: &Workbook,
    u: Key,
    context: RecalcContext,
    budget: Rc<EvaluationBudget>,
    graph: &DepGraph,
) -> (Option<NodeValue>, bool) {
    eval_node_with(wb, wb, u, context, budget, graph)
}

/// evaluate one node, reading cells through `provider` while `wb` supplies the
/// formula's own shape.
fn eval_node_with(
    provider: &dyn CellProvider,
    wb: &Workbook,
    u: Key,
    context: RecalcContext,
    budget: Rc<EvaluationBudget>,
    graph: &DepGraph,
) -> (Option<NodeValue>, bool) {
    let Some(expr) = graph.ast(u.0, cell_of(u)) else {
        return (None, false);
    };
    #[cfg(feature = "test-counters")]
    EVALUATED_FORMULAS.with(|cells| {
        if let Some(cells) = &mut *cells.borrow_mut() {
            cells.push((u.0, cell_of(u)));
        }
    });
    let cell = cell_of(u);
    let authored = wb.sheet(u.0).and_then(|sheet| sheet.array_formula(cell));
    let mut ctx = EvalContext::with_budget(provider, u.0, budget);
    ctx.cell = Some(cell);
    ctx.now_serial = context.now_serial;
    ctx.rand_seed = context.rand_seed.map(|seed| cell_random_seed(seed, u));
    ctx.date_system = wb.date_system;
    ctx.parse_cache = Some(graph.asts());
    let value = match authored {
        Some(_) => NodeValue::Spill(evaluate_spill(&expr, &ctx, cell, authored)),
        None => NodeValue::Scalar(computed(evaluate(&expr, &ctx))),
    };
    let unsupported = ctx.has_unhandled_unsupported_function();
    let refused = ctx.has_unhandled_budget_error();
    if (refused || unsupported) && !matches!(wb.value_cow(u.0, cell).as_ref(), CellValue::Empty) {
        return (None, ctx.exhausted());
    }
    // a refusal is no error value a wrapper may answer, so what `IFERROR` or
    // `ISERROR` made of it never stands in for the result
    if refused {
        let refused = match authored {
            Some(_) => NodeValue::Refused,
            None => NodeValue::Scalar(CellValue::Error {
                value: ErrorValue::Num,
            }),
        };
        return (Some(refused), ctx.exhausted());
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
            charged_spill(&ctx, cell, authored, name.into())
                .map_or(NodeValue::Refused, NodeValue::Spill)
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

/// what a cell caught in a cycle settles at. excel settles an ordinary
/// circular reference at `0`, but an array formula has no rectangle to settle
/// into and caches an empty string instead.
fn circular_value(wb: &Workbook, u: Key) -> CellValue {
    match wb
        .sheet(u.0)
        .and_then(|sheet| sheet.array_formula(cell_of(u)))
    {
        Some(_) => CellValue::Text {
            value: String::new(),
        },
        None => CellValue::Number { value: 0.0 },
    }
}

/// show an array formula the budget refused: an uncached anchor shows `#NUM!`,
/// and a cached one, like the rest of its rectangle, keeps what it had.
fn refuse(wb: &mut Workbook, u: Key, changed: &mut Vec<(SheetId, CellRef)>) {
    let refused = CellValue::Error {
        value: ErrorValue::Num,
    };
    if matches!(wb.value_cow(u.0, cell_of(u)).as_ref(), CellValue::Empty)
        && write_if_changed(wb, u, refused)
    {
        changed.push((u.0, cell_of(u)));
    }
}

/// lay a spilled result out from its anchor: retire the cells the previous
/// result reached and no longer fills, then write the new ones. returns the
/// spilled cells whose value moved, so their readers can be rescheduled, or
/// `None`, having written nothing, when the budget refuses the retirement.
fn write_spill(
    wb: &mut Workbook,
    u: Key,
    spill: Spill,
    budget: &EvaluationBudget,
    changed: &mut Vec<(SheetId, CellRef)>,
) -> Option<Vec<(SheetId, CellRef)>> {
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
    let retired = match (previous, wb.sheet(u.0)) {
        (Some(previous), Some(sheet)) => retired_cells(sheet, previous, anchor, range, budget)?,
        _ => Vec::new(),
    };
    let mut moved = Vec::new();
    for at in retired {
        write_spilled_cell(
            wb,
            (u.0, at.row, at.col),
            CellValue::Empty,
            changed,
            &mut moved,
        );
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
    Some(moved)
}

/// the stored cells of `previous` a result over `range` from `anchor` no
/// longer covers. the search seeks stored cells rather than walking every
/// position, and pays before it starts for the two seeks each row of
/// `previous` may take, then for each cell it finds; `None` once the budget
/// refuses.
fn retired_cells(
    sheet: &xlsx_model::Sheet,
    previous: CellRange,
    anchor: CellRef,
    range: CellRange,
    budget: &EvaluationBudget,
) -> Option<Vec<CellRef>> {
    if range.contains(previous.start) && range.contains(previous.end) {
        return Some(Vec::new());
    }
    let rows = u64::from(previous.end.row - previous.start.row) + 1;
    if !budget.consume(2 * rows + 1) {
        return None;
    }
    let mut cells = sheet.cells_in_range(previous);
    let mut retired = Vec::new();
    let mut paid = true;
    for (at, _) in cells.by_ref() {
        if !budget.consume(1) {
            paid = false;
            break;
        }
        if (at.row, at.col) != (anchor.row, anchor.col) && !range.contains(at) {
            retired.push(at);
        }
    }
    #[cfg(test)]
    RETIREMENT_SEEKS.with(|seeks| seeks.set(seeks.get() + cells.seeks()));
    paid.then_some(retired)
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

    /// formulas ahead of every other in the recalculation order that spend its
    /// whole budget.
    fn spend_the_budget(wb: &mut Workbook, s: SheetId) {
        for row in 1..=11 {
            put_formula(wb, s, &format!("A{row}"), "ROWS(_xlfn.SEQUENCE(1000000))");
        }
    }

    /// a legacy array formula at C20 over C20:D29, its anchor uncached and one
    /// interior cell holding a cached 7.
    fn put_array(wb: &mut Workbook, s: SheetId, formula: &str) {
        put_formula(wb, s, "C20", formula);
        put_num(wb, s, "D25", 7.0);
        wb.sheet_mut(s).unwrap().set_array_formula(
            a1("C20"),
            xlsx_model::CellRange::parse_a1("C20:D29").unwrap(),
        );
    }

    /// an array formula the budget refuses keeps its rectangle and what it
    /// holds, its anchor showing `#NUM!`, and fills in full once it fits.
    #[test]
    fn a_refused_array_keeps_its_rectangle() {
        let refused = CellValue::Error {
            value: xlsx_model::ErrorValue::Num,
        };
        for formula in ["1", "WEBSERVICE(1)"] {
            let (mut wb, s) = one_sheet();
            spend_the_budget(&mut wb, s);
            put_array(&mut wb, s, formula);
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert!(result.limited_cells.contains(&(s, a1("C20"))), "{formula}");
            assert_eq!(value(&wb, s, "C20"), refused, "{formula}");
            assert_eq!(value(&wb, s, "D25"), num(7.0), "{formula}");
            assert_eq!(value(&wb, s, "D29"), CellValue::Empty, "{formula}");
            assert_eq!(
                wb.sheet(s).unwrap().array_formula(a1("C20")),
                Some(xlsx_model::CellRange::parse_a1("C20:D29").unwrap()),
                "{formula}"
            );
        }
        let (mut wb, s) = one_sheet();
        spend_the_budget(&mut wb, s);
        put_array(&mut wb, s, "1");
        rebuild_and_recalc_all(&mut wb, None);
        for row in 1..=11 {
            put_num(&mut wb, s, &format!("A{row}"), 0.0);
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.limited_cells.is_empty());
        for address in ["C20", "D25", "D29"] {
            assert_eq!(value(&wb, s, address), num(1.0), "{address}");
        }
    }

    /// an array formula at `anchor` recorded over `range`.
    fn put_recorded(wb: &mut Workbook, s: SheetId, anchor: &str, formula: &str, range: &str) {
        put_formula(wb, s, anchor, formula);
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1(anchor), xlsx_model::CellRange::parse_a1(range).unwrap());
    }

    /// an engine gap whose small result fits the budget, but whose report
    /// over the recorded rectangle does not, is refused like any other.
    #[test]
    fn a_recovery_the_budget_refuses_keeps_the_rectangle() {
        let (mut wb, s) = one_sheet();
        for row in 1..=10 {
            put_formula(
                &mut wb,
                s,
                &format!("A{row}"),
                "ROWS(_xlfn.SEQUENCE(990000))",
            );
        }
        put_recorded(&mut wb, s, "C20", "{1;2}+WEBSERVICE(1)", "C20:C262163");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.limited_cells.contains(&(s, a1("C20"))));
        assert_eq!(
            value(&wb, s, "C20"),
            CellValue::Error {
                value: xlsx_model::ErrorValue::Num
            }
        );
        assert_eq!(value(&wb, s, "C21"), CellValue::Empty);
        assert_eq!(
            wb.sheet(s).unwrap().array_formula(a1("C20")),
            Some(xlsx_model::CellRange::parse_a1("C20:C262163").unwrap())
        );
    }

    /// a result that shrinks retires its old rectangle by seeking the cells
    /// the sheet stores there, however tall the rectangle is.
    #[test]
    fn a_shrinking_spill_seeks_only_stored_cells() {
        let (mut wb, s) = one_sheet();
        for row in 1..=1000 {
            put_num(&mut wb, s, &format!("A{row}"), 1.0);
        }
        put_recorded(&mut wb, s, "C1", "_xlfn.SEQUENCE(2)", "C1:C262144");
        put_num(&mut wb, s, "C3", 5.0);
        put_num(&mut wb, s, "C262144", 6.0);
        RETIREMENT_SEEKS.with(|seeks| seeks.set(0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.limited_cells.is_empty());
        assert_eq!(value(&wb, s, "C2"), num(2.0));
        assert_eq!(value(&wb, s, "C3"), CellValue::Empty);
        assert_eq!(value(&wb, s, "C262144"), CellValue::Empty);
        assert_eq!(
            wb.sheet(s).unwrap().array_formula(a1("C1")),
            Some(xlsx_model::CellRange::parse_a1("C1:C2").unwrap())
        );
        let seeks = RETIREMENT_SEEKS.with(|seeks| seeks.get());
        assert!((1..=2 * 1000 + 8).contains(&seeks), "{seeks} seeks");
    }

    /// retiring the rectangles of shrinking results shares the recalculation
    /// budget, whether the result shrank or an obstruction cut it to
    /// `#SPILL!`; a retirement it refuses leaves the rectangle as it was.
    #[test]
    fn shrinking_spills_share_the_recalculation_budget() {
        for (formula, obstructed) in [("_xlfn.SEQUENCE(2)", false), ("_xlfn.SEQUENCE(2,2)", true)] {
            let (mut wb, s) = one_sheet();
            let anchors: Vec<CellRef> = (0..40).map(|index| CellRef::new(0, index * 2)).collect();
            for &at in &anchors {
                let sheet = wb.sheet_mut(s).unwrap();
                sheet.set_cell(
                    at,
                    Cell {
                        formula: Some(formula.into()),
                        ..Cell::default()
                    },
                );
                sheet.set_array_formula(
                    at,
                    xlsx_model::CellRange::new(at, CellRef::new(262_143, at.col)),
                );
                if obstructed {
                    sheet.set_cell(
                        CellRef::new(0, at.col + 1),
                        Cell {
                            value: num(1.0),
                            ..Cell::default()
                        },
                    );
                }
            }
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            let mut kept = 0;
            for &at in &anchors {
                let range = wb.sheet(s).unwrap().array_formula(at).unwrap();
                if range.end.row == 262_143 {
                    kept += 1;
                    assert!(result.limited_cells.contains(&(s, at)), "{formula}");
                    assert_eq!(
                        wb.value(s, at),
                        CellValue::Error {
                            value: xlsx_model::ErrorValue::Num
                        },
                        "{formula}"
                    );
                }
            }
            assert!(kept > 0 && kept < anchors.len(), "{formula}: {kept} kept");
        }
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

    #[test]
    fn clockless_today_and_now_keep_caches_and_recalculate_dependents() {
        for (formula, clocked) in [("TODAY()", 46_000.0), ("NOW()", 46_000.25)] {
            let (mut wb, s) = one_sheet();
            put_cached_formula(&mut wb, s, "A1", formula, num(45_000.75));
            put_cached_formula(&mut wb, s, "B1", "A1+1", num(-1.0));
            let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "A1"), num(45_000.75));
            assert_eq!(value(&wb, s, "B1"), num(45_001.75));
            assert_eq!(changed_a1(&result), vec!["B1"]);
            recalc_after(&mut wb, &mut graph, &[], Some(46_000.25));
            assert_eq!(value(&wb, s, "A1"), num(clocked));
            assert_eq!(value(&wb, s, "B1"), num(clocked + 1.0));
        }
    }

    #[test]
    fn clockless_today_without_a_cache_reports_value() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "TODAY()");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(
            value(&wb, s, "A1"),
            CellValue::Error {
                value: ErrorValue::Value
            }
        );
        rebuild_and_recalc_all(&mut wb, Some(46_000.25));
        assert_eq!(value(&wb, s, "A1"), num(46_000.0));
    }

    #[test]
    fn clockless_today_expression_keeps_its_stale_cache_after_an_input_changes() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A2", 44_999.0);
        put_cached_formula(&mut wb, s, "A1", "TODAY()-A2", num(1.0));
        put_formula(&mut wb, s, "B1", "A1+1");
        let mut graph = DepGraph::build(&wb);
        put_num(&mut wb, s, "A2", 44_998.0);
        recalc_after(&mut wb, &mut graph, &[(s, a1("A2"))], None);
        assert_eq!(value(&wb, s, "A1"), num(1.0));
        assert_eq!(value(&wb, s, "B1"), num(2.0));
        recalc_after(&mut wb, &mut graph, &[], Some(46_000.25));
        assert_eq!(value(&wb, s, "A1"), num(1_002.0));
        assert_eq!(value(&wb, s, "B1"), num(1_003.0));
    }

    #[test]
    fn iferror_handles_clockless_today_and_overwrites_the_cache() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "A1", "IFERROR(TODAY(),0)", num(99.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(0.0));
        rebuild_and_recalc_all(&mut wb, Some(46_000.25));
        assert_eq!(value(&wb, s, "A1"), num(46_000.0));
    }

    #[test]
    fn clockless_date_text_keeps_caches_only_when_it_needs_the_current_year() {
        for (formula, clocked) in [
            ("DATEVALUE(\"1/1\")", 43_831.0),
            ("YEAR(\"1/1\")", 2020.0),
            ("MONTH(\"1/1\")", 1.0),
            ("DAYS(\"1/1\",1)", 43_830.0),
        ] {
            let (mut wb, s) = one_sheet();
            put_cached_formula(&mut wb, s, "A1", formula, num(-1.0));
            put_cached_formula(&mut wb, s, "B1", "DATEVALUE(\"1/1/2020\")", num(-1.0));
            rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "A1"), num(-1.0), "{formula}");
            assert_eq!(value(&wb, s, "B1"), num(43_831.0));
            rebuild_and_recalc_all(&mut wb, Some(43_831.5));
            assert_eq!(value(&wb, s, "A1"), num(clocked), "{formula}");
            assert_eq!(value(&wb, s, "B1"), num(43_831.0));
        }
    }

    #[test]
    fn clockless_today_array_keeps_its_cached_rectangle() {
        let (mut wb, s) = one_sheet();
        put_recorded(&mut wb, s, "A1", "{1;2}+TODAY()", "A1:A2");
        set_cached(&mut wb, s, "A1", num(45_001.0));
        put_num(&mut wb, s, "A2", 45_002.0);
        put_formula(&mut wb, s, "B1", "SUM(A1:A2)");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(45_001.0));
        assert_eq!(value(&wb, s, "A2"), num(45_002.0));
        assert_eq!(value(&wb, s, "B1"), num(90_003.0));
        assert_eq!(
            wb.sheet(s).unwrap().array_formula(a1("A1")),
            Some(CellRange::parse_a1("A1:A2").unwrap())
        );
        rebuild_and_recalc_all(&mut wb, Some(46_000.25));
        assert_eq!(value(&wb, s, "A1"), num(46_001.0));
        assert_eq!(value(&wb, s, "A2"), num(46_002.0));
        assert_eq!(value(&wb, s, "B1"), num(92_003.0));
    }

    fn put_text(wb: &mut Workbook, sheet: SheetId, cell: &str, text: &str) {
        wb.sheet_mut(sheet).unwrap().set_cell(
            a1(cell),
            Cell {
                value: CellValue::Text { value: text.into() },
                ..Cell::default()
            },
        );
    }

    fn defined(wb: &mut Workbook, name: &str, formula: &str) {
        wb.defined_names.push(DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: None,
            hidden: false,
        });
    }

    #[test]
    fn references_into_another_workbook_keep_their_cached_results() {
        let (mut wb, s) = one_sheet();
        defined(&mut wb, "Rate", "'[2]Rates'!$A$1");
        for (address, formula, cached) in [
            ("A1", "'[1]Other Book'!B2", 91.5),
            ("A2", "SUM('[1]Data'!A1:A4)*2", 10.0),
            ("A3", "[1]Data!A1+1", 4.0),
            ("A4", "Rate*2", 0.5),
            ("A5", "INDIRECT(\"'[1]Data'!A1\")", 3.0),
        ] {
            put_cached_formula(&mut wb, s, address, formula, num(cached));
        }
        let cached_false = CellValue::Bool { value: false };
        put_cached_formula(
            &mut wb,
            s,
            "A6",
            "ISERROR('[1]Data'!A1)",
            cached_false.clone(),
        );
        put_formula(&mut wb, s, "B1", "A1+1");
        put_cached_formula(&mut wb, s, "C1", "Missing!A1", num(7.0));
        rebuild_and_recalc_all(&mut wb, None);
        for (address, cached) in [
            ("A1", 91.5),
            ("A2", 10.0),
            ("A3", 4.0),
            ("A4", 0.5),
            ("A5", 3.0),
        ] {
            assert_eq!(value(&wb, s, address), num(cached), "{address}");
        }
        assert_eq!(value(&wb, s, "A6"), cached_false);
        assert_eq!(value(&wb, s, "B1"), num(92.5));
        assert_eq!(
            value(&wb, s, "C1"),
            CellValue::Error {
                value: ErrorValue::Ref
            }
        );
    }

    #[test]
    fn indirect_in_r1c1_style_keeps_its_cached_result() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 5.0);
        put_cached_formula(&mut wb, s, "B1", "INDIRECT(\"R1C1\",FALSE)", num(8.0));
        put_formula(&mut wb, s, "B2", "B1+1");
        put_formula(&mut wb, s, "B3", "INDIRECT(\"R1C1\",FALSE)");
        put_cached_formula(&mut wb, s, "B4", "INDIRECT(\"A1\")", num(-1.0));
        put_cached_formula(&mut wb, s, "B5", "INDIRECT(\"Nowhere!A1\")", num(-1.0));
        put_cached_formula(&mut wb, s, "B6", "INDIRECT(\"'[1]Book'!Rate\")", num(0.25));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(8.0));
        assert_eq!(value(&wb, s, "B2"), num(9.0));
        let reference = CellValue::Error {
            value: ErrorValue::Ref,
        };
        assert_eq!(value(&wb, s, "B3"), reference);
        assert_eq!(value(&wb, s, "B4"), num(5.0));
        assert_eq!(value(&wb, s, "B5"), reference);
        assert_eq!(value(&wb, s, "B6"), num(0.25));
    }

    #[test]
    fn a_range_in_a_scalar_formula_reads_the_cell_on_the_formulas_row_or_column() {
        let (mut wb, s) = one_sheet();
        wb.sheets.push(Sheet::new("Data"));
        let data = SheetId(1);
        for row in 1..=5 {
            put_num(&mut wb, s, &format!("A{row}"), row as f64 * 10.0);
            put_num(&mut wb, data, &format!("A{row}"), row as f64);
        }
        put_num(&mut wb, s, "D10", 4.0);
        put_num(&mut wb, s, "E10", 5.0);
        for (address, index) in [("E1", 1.0), ("E2", 4.0), ("E3", 2.0)] {
            put_num(&mut wb, s, address, index);
        }
        defined(&mut wb, "Pay", "Sheet1!$A$1:$A$5");
        put_recorded(&mut wb, s, "H1", "A1:A5*2", "H1:H5");
        for (address, formula) in [
            ("C3", "A1:A5"),
            ("C4", "Pay*2"),
            ("C5", "IF(Pay<>\"\",DATE(2020,Pay/10,1),\"\")"),
            ("C2", "A:A"),
            ("C1", "SUM(A1:A5)"),
            ("G4", "Data!A1:A5"),
            ("E12", "D10:F10"),
            ("D2", "INDEX(A1:A5,E1:E3)"),
            ("F3", "Pay+SUMPRODUCT(Pay*1)"),
            ("F4", "SUMPRODUCT(Pay*1)+Pay"),
            ("C7", "A1:A5"),
            ("C8", "A1:B5"),
            ("H12", "D10:F10"),
        ] {
            put_formula(&mut wb, s, address, formula);
        }
        rebuild_and_recalc_all(&mut wb, None);
        for (address, expected) in [
            ("C3", num(30.0)),
            ("C4", num(80.0)),
            ("C5", num(43_952.0)),
            ("C2", num(20.0)),
            ("C1", num(150.0)),
            ("G4", num(4.0)),
            ("E12", num(5.0)),
            ("D2", num(40.0)),
            ("F3", num(180.0)),
            ("F4", num(190.0)),
            ("H3", num(60.0)),
            ("H5", num(100.0)),
        ] {
            assert_eq!(value(&wb, s, address), expected, "{address}");
        }
        for address in ["C7", "C8", "H12"] {
            assert_eq!(
                value(&wb, s, address),
                CellValue::Error {
                    value: ErrorValue::Value
                },
                "{address}"
            );
        }
    }

    #[test]
    fn arithmetic_on_date_text_keeps_its_cached_result() {
        let (mut wb, s) = one_sheet();
        put_text(&mut wb, s, "A1", "19/11/2021");
        put_text(&mut wb, s, "A2", "no date");
        put_text(&mut wb, s, "A3", "5");
        put_cached_formula(&mut wb, s, "B1", "A1+1", num(44_520.0));
        put_formula(&mut wb, s, "B2", "B1+1");
        put_cached_formula(&mut wb, s, "B3", "A2+1", num(-1.0));
        put_cached_formula(&mut wb, s, "B4", "A3+1", num(-1.0));
        put_cached_formula(&mut wb, s, "B5", "-A1", num(-44_519.0));
        put_cached_formula(&mut wb, s, "B6", "A1%", num(445.19));
        put_text(&mut wb, s, "A7", "31/04/2021");
        put_cached_formula(&mut wb, s, "B7", "A7+1", num(-1.0));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(44_520.0));
        assert_eq!(value(&wb, s, "B5"), num(-44_519.0));
        assert_eq!(value(&wb, s, "B6"), num(445.19));
        assert_eq!(value(&wb, s, "B2"), num(44_521.0));
        assert_eq!(
            value(&wb, s, "B3"),
            CellValue::Error {
                value: ErrorValue::Value
            }
        );
        assert_eq!(value(&wb, s, "B4"), num(6.0));
        assert_eq!(
            value(&wb, s, "B7"),
            CellValue::Error {
                value: ErrorValue::Value
            }
        );
    }

    #[test]
    fn an_array_formula_over_date_text_keeps_its_cached_rectangle() {
        for (formula, first, second) in [
            ("-A1:A2", -44_519.0, -44_520.0),
            ("A1:A2%", 445.19, 445.2),
            ("A1:A2+1", 44_520.0, 44_521.0),
        ] {
            let (mut wb, s) = one_sheet();
            put_text(&mut wb, s, "A1", "19/11/2021");
            put_text(&mut wb, s, "A2", "20/11/2021");
            put_recorded(&mut wb, s, "C1", formula, "C1:C2");
            set_cached(&mut wb, s, "C1", num(first));
            put_num(&mut wb, s, "C2", second);
            rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "C1"), num(first), "{formula}");
            assert_eq!(value(&wb, s, "C2"), num(second), "{formula}");
        }
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

    #[test]
    fn cycle_classification_selective_index() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        for row in 2..=10_001 {
            put_formula(
                &mut wb,
                s,
                &format!("A{row}"),
                &format!("INDEX($A$1:A{row},1)"),
            );
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        assert!(result.limited_cells.is_empty());
        for row in 2..=10_001 {
            assert_eq!(value(&wb, s, &format!("A{row}")), num(7.0));
        }
    }

    #[test]
    fn cycle_classification_lazy_if() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        for row in 2..=10_001 {
            put_formula(
                &mut wb,
                s,
                &format!("A{row}"),
                &format!("IF(TRUE,7,A{row})"),
            );
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        assert!(result.limited_cells.is_empty());
        for row in 2..=10_001 {
            assert_eq!(value(&wb, s, &format!("A{row}")), num(7.0));
        }
    }

    #[test]
    fn cycle_classification_selective_vlookup() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_num(&mut wb, s, "B1", 7.0);
        for row in 2..=10_001 {
            put_formula(
                &mut wb,
                s,
                &format!("B{row}"),
                &format!("VLOOKUP($A$1,$A$1:B{row},2,FALSE)"),
            );
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        assert!(result.limited_cells.is_empty());
        for row in 2..=10_001 {
            assert_eq!(value(&wb, s, &format!("B{row}")), num(7.0));
        }
    }

    #[test]
    fn cycle_classification_whole_column_index() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        for row in 2..=10_001 {
            put_formula(&mut wb, s, &format!("A{row}"), "INDEX(A:A,1)");
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        assert!(result.limited_cells.is_empty());
        for row in 2..=10_001 {
            assert_eq!(value(&wb, s, &format!("A{row}")), num(7.0));
        }
    }

    fn check_cycle_dependents(other_sheet: bool) {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_formula(&mut wb, s, "B1", "C1+1");
        put_formula(&mut wb, s, "C1", "B1+1");
        let dependent_sheet = if other_sheet {
            wb.sheets.push(Sheet::new("Sheet2"));
            SheetId(1)
        } else {
            s
        };
        for row in 2..=10_001 {
            put_formula(
                &mut wb,
                dependent_sheet,
                &format!("D{row}"),
                "Sheet1!B1+Sheet1!C1+7",
            );
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        let mut cycles = result.cycle_cells.clone();
        cycles.sort_by(sort_key);
        assert_eq!(cycles, vec![(s, a1("B1")), (s, a1("C1"))]);
        assert!(result.limited_cells.is_empty());
        assert_eq!(value(&wb, s, "B1"), num(0.0));
        assert_eq!(value(&wb, s, "C1"), num(0.0));
        for row in 2..=10_001 {
            assert_eq!(value(&wb, dependent_sheet, &format!("D{row}")), num(7.0));
        }
        assert_eq!(result.changed.len(), 10_002);
    }

    #[test]
    fn cycle_classification_same_sheet_dependents() {
        check_cycle_dependents(false);
    }

    #[test]
    fn cycle_classification_cross_sheet_dependents() {
        check_cycle_dependents(true);
    }

    #[test]
    fn cycle_classification_sum_self_reference() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_formula(&mut wb, s, "A2", "SUM(A1:A2)");
        for row in 2..=10_001 {
            put_formula(&mut wb, s, &format!("B{row}"), "$A$2+7");
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(result.cycle_cells, vec![(s, a1("A2"))]);
        assert_eq!(value(&wb, s, "A2"), num(0.0));
        for row in 2..=10_001 {
            assert_eq!(value(&wb, s, &format!("B{row}")), num(7.0));
        }
    }

    #[test]
    fn cycle_classification_long_lazy_chain() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        for row in 2..=20_001 {
            put_formula(
                &mut wb,
                s,
                &format!("A{row}"),
                &format!("IF(TRUE,A{}+1,$A$20001)", row - 1),
            );
        }
        FORMULA_EVALUATIONS.with(|count| count.set(0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        let evaluations = FORMULA_EVALUATIONS.with(|count| count.get());
        assert!(result.cycle_cells.is_empty());
        assert!(result.limited_cells.is_empty());
        assert_eq!(value(&wb, s, "A20001"), num(20_007.0));
        assert!(evaluations <= 3 * 20_000, "{evaluations} evaluations");
    }

    /// a chain that reads against key order waits once per cell, then settles
    /// back along the waits.
    #[test]
    fn cycle_classification_reversed_lazy_chain() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A20002", 7.0);
        for row in 2..=20_001 {
            put_formula(
                &mut wb,
                s,
                &format!("A{row}"),
                &format!("IF(TRUE,A{}+1,$A$2)", row + 1),
            );
        }
        FORMULA_EVALUATIONS.with(|count| count.set(0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        let evaluations = FORMULA_EVALUATIONS.with(|count| count.get());
        assert!(result.cycle_cells.is_empty());
        assert!(result.limited_cells.is_empty());
        assert_eq!(value(&wb, s, "A2"), num(20_007.0));
        assert!(evaluations <= 3 * 20_000, "{evaluations} evaluations");
    }

    #[test]
    fn cycle_classification_wait_cycle_excludes_static_component_dependents() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_formula(&mut wb, s, "B1", "IF(TRUE,C1+1,D1)");
        put_formula(&mut wb, s, "C1", "B1+1");
        put_formula(&mut wb, s, "D1", "B1+7");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(result.cycle_cells, vec![(s, a1("B1")), (s, a1("C1"))]);
        assert_eq!(value(&wb, s, "D1"), num(7.0));
    }

    #[test]
    fn cycle_classification_waits_before_selecting_a_branch() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_formula(&mut wb, s, "B1", "IF(C1=7,A1,B1)");
        put_formula(&mut wb, s, "C1", "IF(TRUE,A1,B1)");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "B1"), num(7.0));
        assert_eq!(value(&wb, s, "C1"), num(7.0));
    }

    /// A1 reads itself, and each later cell reads itself once the one before
    /// it shows `0`, so every cell is a cycle of its own, found one after the
    /// other. each evaluation registers at most one wait and each member is
    /// collected into a cycle once, so the work stays within one step per
    /// evaluation plus one per cell.
    #[test]
    fn cycle_classification_self_cycles_found_in_turn() {
        const CELLS: u32 = 20_000;
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", &format!("IF(TRUE,A1,A{CELLS})"));
        for row in 2..=CELLS {
            put_formula(
                &mut wb,
                s,
                &format!("A{row}"),
                &format!("IF(A{}=0,A{row},1)", row - 1),
            );
        }
        FORMULA_EVALUATIONS.with(|count| count.set(0));
        SCHEDULER_STEPS.with(|count| count.set(0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        let evaluations = FORMULA_EVALUATIONS.with(|count| count.get());
        let steps = SCHEDULER_STEPS.with(|count| count.get());
        assert!(
            evaluations <= 3 * u64::from(CELLS),
            "{evaluations} evaluations"
        );
        assert!(steps <= 4 * u64::from(CELLS), "{steps} steps");
        let every: Vec<_> = (0..CELLS).map(|row| (s, CellRef::new(row, 0))).collect();
        assert!(
            result.cycle_cells == every,
            "{} cycle cells",
            result.cycle_cells.len()
        );
        for row in 1..=CELLS {
            assert_eq!(value(&wb, s, &format!("A{row}")), num(0.0));
        }
    }

    /// arrays tied into one component only by branches they never take, each
    /// reading a constant outside it. each read is compared with the
    /// rectangles of at most nine grid blocks of the one size class here, and
    /// rectangles that do not overlap never share a block.
    #[test]
    fn cycle_classification_spill_lookup_scales_with_reads() {
        const ARRAYS: u32 = 5_000;
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        for row in 1..=ARRAYS {
            put_recorded(
                &mut wb,
                s,
                &format!("B{row}"),
                &format!("IF($A$1=7,{{1,2}},B{})", row % ARRAYS + 1),
                &format!("B{row}:C{row}"),
            );
        }
        SETTLING_READS.with(|count| count.set(0));
        SPILL_CHECKS.with(|count| count.set(0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        let reads = SETTLING_READS.with(|count| count.get());
        let checks = SPILL_CHECKS.with(|count| count.get());
        assert!(reads >= u64::from(ARRAYS), "{reads} reads");
        assert!(checks <= 9 * reads, "{checks} checks over {reads} reads");
        assert!(result.cycle_cells.is_empty());
        for row in 1..=ARRAYS {
            assert_eq!(value(&wb, s, &format!("B{row}")), num(1.0));
            assert_eq!(value(&wb, s, &format!("C{row}")), num(2.0));
        }
    }

    /// the range graph puts A1 and the array on one cycle, but A1 only reads
    /// inside the rectangle the array spills over: it waits for the array,
    /// which spills its real result.
    #[test]
    fn cycle_classification_waits_on_a_spill_interior() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "B2");
        put_recorded(&mut wb, s, "B1", "IF(TRUE,{1;2},A1)", "B1:B2");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        assert_eq!(value(&wb, s, "A1"), num(2.0));
        assert_eq!(value(&wb, s, "B1"), num(1.0));
        assert_eq!(value(&wb, s, "B2"), num(2.0));
        assert_eq!(changed_a1(&result), vec!["A1", "B1", "B2"]);
    }

    /// a range reaches the formulas it covers through a tree over each row or
    /// column, and must still run after every one of them.
    #[test]
    fn ranges_over_many_formulas_order_every_precedent_first() {
        const N: u32 = 40;
        let (mut wb, s) = one_sheet();
        let column = |at: u32| CellRef::new(at, 0).to_a1();
        let row = |at: u32| CellRef::new(N + 1, at).to_a1();
        put_num(&mut wb, s, &column(N - 1), 1.0);
        put_num(&mut wb, s, &row(N - 1), 1.0);
        for at in 0..N - 1 {
            let below = format!("SUM({}:{})", column(at + 1), column(N - 1));
            put_formula(&mut wb, s, &column(at), &below);
            let right = format!("SUM({}:{})", row(at + 1), row(N - 1));
            put_formula(&mut wb, s, &row(at), &right);
        }
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.cycle_cells.is_empty());
        for at in 0..N - 1 {
            let expected = num(2f64.powi((N - 2 - at) as i32));
            assert_eq!(value(&wb, s, &column(at)), expected);
            assert_eq!(value(&wb, s, &row(at)), expected);
        }
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

    /// an array formula caught in a cycle has no rectangle to settle into, so
    /// it caches an empty string where an ordinary one settles at `0`.
    #[test]
    fn a_circular_array_formula_settles_at_empty_text() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "A1+1");
        put_formula(&mut wb, s, "B1", "B1+1");
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("B1"), xlsx_model::CellRange::parse_a1("B1").unwrap());
        let (_, r) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(r.cycle_cells, vec![(s, a1("A1")), (s, a1("B1"))]);
        assert_eq!(value(&wb, s, "A1"), num(0.0));
        assert_eq!(
            value(&wb, s, "B1"),
            CellValue::Text {
                value: String::new()
            }
        );
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
    fn seeded_draws_distinguish_seed_sheet_and_row_col_pairs() {
        let wb = Workbook::default();
        for (first, second) in [
            ((0, (SheetId(1), 1, 2)), (1, (SheetId(0), 1, 2))),
            ((42, (SheetId(0), 1, 2)), (42, (SheetId(0), 2, 1))),
        ] {
            let first_seed = cell_random_seed(first.0, first.1);
            let second_seed = cell_random_seed(second.0, second.1);
            assert_ne!(first_seed, second_seed);
            let mut first_ctx = EvalContext::new(&wb, first.1.0);
            first_ctx.rand_seed = Some(first_seed);
            let mut second_ctx = EvalContext::new(&wb, second.1.0);
            second_ctx.rand_seed = Some(second_seed);
            let first_draws: [f64; 4] = std::array::from_fn(|_| first_ctx.next_random_unit());
            let second_draws: [f64; 4] = std::array::from_fn(|_| second_ctx.next_random_unit());
            for (first, second) in first_draws.into_iter().zip(second_draws) {
                assert_ne!(first, second);
            }
            assert!(first_draws.windows(2).all(|draws| draws[0] != draws[1]));
            assert!(second_draws.windows(2).all(|draws| draws[0] != draws[1]));
        }
    }

    #[test]
    fn earlier_independent_random_cells_preserve_seeded_draws() {
        let (mut wb, sheet) = one_sheet();
        put_formula(&mut wb, sheet, "B3", "RANDBETWEEN(1,1000000)");
        put_formula(
            &mut wb,
            sheet,
            "C3",
            "RANDBETWEEN(1,1000000)*1000000+RANDBETWEEN(1,1000000)",
        );
        put_formula(
            &mut wb,
            sheet,
            "E3",
            "MAKEARRAY(2,2,LAMBDA(r,c,RANDBETWEEN(1,1000000)))",
        );
        wb.sheet_mut(sheet)
            .unwrap()
            .set_array_formula(a1("E3"), CellRange::parse_a1("E3:F4").unwrap());
        let (mut graph, _) = rebuild_and_recalc_all_with_seed(&mut wb, Some(45_000.5), Some(42));
        let addresses = ["B3", "C3", "E3", "E4", "F3", "F4"];
        let before = addresses.map(|address| value(&wb, sheet, address));
        assert!(
            before
                .iter()
                .all(|value| matches!(value, CellValue::Number { .. }))
        );

        put_formula(&mut wb, sheet, "A1", "RANDBETWEEN(1,1000000)");
        graph.set_formula(sheet, a1("A1"), Some("RANDBETWEEN(1,1000000)"));
        let seeds = [(sheet, a1("A1"))];
        let recompute = collect_recompute(&graph, &seeds);
        let order = topo_order(&graph, &recompute);
        assert!(order.iter().all(|step| matches!(step, Step::Cell(_))));
        assert_eq!(order[0], Step::Cell(key(sheet, a1("A1"))));
        for address in ["B3", "C3", "E3"] {
            assert!(order.contains(&Step::Cell(key(sheet, a1(address)))));
        }
        recalc_after_with_seed(&mut wb, &mut graph, &seeds, Some(45_000.5), Some(42));
        assert!(matches!(value(&wb, sheet, "A1"), CellValue::Number { .. }));
        assert_eq!(addresses.map(|address| value(&wb, sheet, address)), before);
        let formulas = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
        assert_eq!(
            topo_order(&graph, &formulas)[0],
            Step::Cell(key(sheet, a1("A1")))
        );
        rebuild_and_recalc_all_with_seed(&mut wb, Some(45_000.5), Some(42));
        assert_eq!(addresses.map(|address| value(&wb, sheet, address)), before);
    }

    #[test]
    fn seeded_draws_survive_deferred_and_spill_recalculation() {
        let (mut wb, sheet) = one_sheet();
        put_formula(
            &mut wb,
            sheet,
            "A1",
            "IF(ROW()=1,RANDBETWEEN(1,1000000),INDEX(A1:A2,ROW()-1))",
        );
        put_formula(&mut wb, sheet, "A2", "A1*2");
        put_formula(
            &mut wb,
            sheet,
            "B1",
            "MAKEARRAY(2,2,LAMBDA(r,c,RANDBETWEEN(1,1000000)))",
        );
        wb.sheet_mut(sheet)
            .unwrap()
            .set_array_formula(a1("B1"), CellRange::parse_a1("B1:C2").unwrap());
        let (mut graph, result) =
            rebuild_and_recalc_all_with_seed(&mut wb, Some(45_000.5), Some(7));
        assert!(result.cycle_cells.is_empty());
        let addresses = ["A1", "A2", "B1", "B2", "C1", "C2"];
        let before = addresses.map(|address| value(&wb, sheet, address));
        for index in [0, 2, 3, 4, 5] {
            match before[index] {
                CellValue::Number { value } => {
                    assert!((1.0..=1_000_000.0).contains(&value));
                    assert_eq!(value.fract(), 0.0);
                }
                ref other => panic!("expected random number, got {other:?}"),
            }
        }
        assert_ne!(before[2], before[3]);
        assert_ne!(before[2], before[4]);
        rebuild_and_recalc_all_with_seed(&mut wb, Some(45_000.5), Some(7));
        assert_eq!(addresses.map(|address| value(&wb, sheet, address)), before);
        put_num(&mut wb, sheet, "Z10", 1.0);
        recalc_after_with_seed(
            &mut wb,
            &mut graph,
            &[(sheet, a1("Z10"))],
            Some(45_000.5),
            Some(7),
        );
        assert_eq!(addresses.map(|address| value(&wb, sheet, address)), before);
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

    /// a far corner cell gives `Data!A1:XFD1048576` the extent the
    /// per-formula cap refuses; formulas go on the second sheet.
    fn over_the_formula_cap() -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.sheets.push(Sheet::new("Formula"));
        put_num(&mut wb, SheetId(0), "XFD1048576", 1.0);
        wb
    }

    /// an error handler over a formula's refused work answers nothing: the
    /// cell keeps its cached value, or shows `#NUM!` without one.
    #[test]
    fn error_handlers_do_not_answer_a_formula_over_its_cap() {
        let refused = CellValue::Error {
            value: ErrorValue::Num,
        };
        let s = SheetId(1);
        for formula in [
            "IFERROR(SUM(Data!A1:XFD1048576),42)",
            "IFERROR(SUM(Data!A1:XFD1048576),#NUM!)",
            "ISERROR(SUM(Data!A1:XFD1048576))",
            "ISNUMBER(SUM(Data!A1:XFD1048576))",
            "COUNT(1,SUM(Data!A1:XFD1048576))",
        ] {
            let mut wb = over_the_formula_cap();
            put_cached_formula(&mut wb, s, "A1", formula, num(123.0));
            put_formula(&mut wb, s, "A2", formula);
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "A1"), num(123.0), "{formula}");
            assert_eq!(value(&wb, s, "A2"), refused, "{formula}");
            assert_eq!(result.limited_cells.len(), 2, "{formula}");
            for at in ["A1", "A2"] {
                assert!(result.limited_cells.contains(&(s, a1(at))), "{formula}");
            }
        }
    }

    /// a lookup that runs out of its formula's budget before the match keeps
    /// the cached match under `IFERROR` instead of taking the fallback.
    #[test]
    fn a_lookup_over_its_cap_keeps_its_cached_match_under_iferror() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A600000", 600_000.0);
        put_num(&mut wb, s, "C600000", 1.0);
        put_num(&mut wb, s, "D600000", 42.0);
        let formula = "IFERROR(_xlfn.XLOOKUP(1,OFFSET(C1,0,0,SUM(A:A)),OFFSET(D1,0,0,A600000)),0)";
        put_cached_formula(&mut wb, s, "B1", formula, num(42.0));
        put_formula(&mut wb, s, "B2", formula);
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(42.0));
        assert_eq!(
            value(&wb, s, "B2"),
            CellValue::Error {
                value: ErrorValue::Num
            }
        );
        assert!(result.limited_cells.contains(&(s, a1("B1"))));
        assert!(result.limited_cells.contains(&(s, a1("B2"))));
    }

    /// an array formula whose refused work an error handler would answer
    /// keeps its cached anchor and rectangle.
    #[test]
    fn a_refused_array_keeps_its_cache_under_an_error_handler() {
        let s = SheetId(1);
        for formula in [
            "IFERROR(SUM(Data!A1:XFD1048576),0)",
            "IFERROR(Data!A1:A10*SUM(Data!A1:XFD1048576),0)",
            "ISERROR(SUM(Data!A1:XFD1048576))",
        ] {
            let mut wb = over_the_formula_cap();
            put_cached_formula(&mut wb, s, "C20", formula, num(5.0));
            put_num(&mut wb, s, "D25", 7.0);
            wb.sheet_mut(s)
                .unwrap()
                .set_array_formula(a1("C20"), CellRange::parse_a1("C20:D29").unwrap());
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "C20"), num(5.0), "{formula}");
            assert_eq!(value(&wb, s, "D25"), num(7.0), "{formula}");
            assert_eq!(value(&wb, s, "D29"), CellValue::Empty, "{formula}");
            assert!(result.limited_cells.contains(&(s, a1("C20"))), "{formula}");
        }
    }

    /// once the recalculation's budget is spent, error handlers over the
    /// refused work keep the cached values and uncached cells show `#NUM!`,
    /// while a formula that charges nothing still computes.
    #[test]
    fn a_spent_recalculation_keeps_cached_values_under_error_handlers() {
        let (mut wb, s) = one_sheet();
        spend_the_budget(&mut wb, s);
        for row in 1..=3 {
            put_num(&mut wb, s, &format!("E{row}"), f64::from(row));
        }
        let cached = [
            ("B20", "IFERROR(SUMPRODUCT(E1:E3,E1:E3),0)", num(14.0)),
            (
                "B21",
                "ISERROR(SUMPRODUCT(E1:E3))",
                CellValue::Bool { value: false },
            ),
            (
                "B22",
                "ISNUMBER(SUMPRODUCT(E1:E3))",
                CellValue::Bool { value: true },
            ),
            ("B23", "ISERROR(B20)", CellValue::Bool { value: false }),
            ("B24", "IF(ISERROR(SUM(E1:E3)),-1,SUM(E1:E3))", num(6.0)),
        ];
        for (at, formula, cache) in &cached {
            put_cached_formula(&mut wb, s, at, formula, cache.clone());
        }
        put_formula(&mut wb, s, "B30", "IFERROR(SUMPRODUCT(E1:E3),0)");
        put_formula(&mut wb, s, "B31", "ISERROR(SUMPRODUCT(E1:E3))");
        put_formula(&mut wb, s, "B32", "1+2");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        for (at, formula, cache) in &cached {
            assert_eq!(value(&wb, s, at), *cache, "{formula}");
            assert!(result.limited_cells.contains(&(s, a1(at))), "{formula}");
        }
        let refused = CellValue::Error {
            value: ErrorValue::Num,
        };
        assert_eq!(value(&wb, s, "B30"), refused);
        assert_eq!(value(&wb, s, "B31"), refused);
        assert_eq!(value(&wb, s, "B32"), num(3.0));
        assert!(!result.limited_cells.contains(&(s, a1("B32"))));
    }

    /// an inner handler that answers an error after work was spent there
    /// leaves a later refusal standing, so the outer handler cannot answer it.
    #[test]
    fn an_outer_handler_keeps_the_cache_after_an_inner_one_answers() {
        let inputs = |spent: bool| {
            let (mut wb, s) = one_sheet();
            if spent {
                spend_the_budget(&mut wb, s);
            }
            put_num(&mut wb, s, "A600000", 1.0);
            put_num(&mut wb, s, "C1", 1.0);
            put_num(&mut wb, s, "D1", 42.0);
            (wb, s)
        };
        let lookup =
            "IFERROR(IFERROR(_xlfn.XLOOKUP(NA(),OFFSET(C1,0,0,SUM(A:A)),D1),42)+SUM(A:A),0)";
        // reads `A:A` twice, which is over the formula's cap
        let twice = "IFERROR(IFERROR(INDEX(OFFSET(C1,0,0,SUM(A:A)),1)+NA(),42)+SUM(A:A),0)";
        let refused = CellValue::Error {
            value: ErrorValue::Num,
        };
        for (formula, spent) in [(lookup, true), (twice, false), (twice, true)] {
            let (mut wb, s) = inputs(spent);
            put_cached_formula(&mut wb, s, "E1", formula, num(43.0));
            put_formula(&mut wb, s, "E2", formula);
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "E1"), num(43.0), "{formula} {spent}");
            assert_eq!(value(&wb, s, "E2"), refused, "{formula} {spent}");
            assert!(result.limited_cells.contains(&(s, a1("E1"))), "{formula}");
        }
        let (mut wb, s) = inputs(false);
        put_formula(&mut wb, s, "E1", lookup);
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "E1"), num(43.0));
        assert!(result.limited_cells.is_empty());
    }

    #[test]
    fn a_long_recalculation_reports_progress() {
        thread_local! {
            static CALLS: std::cell::Cell<u32> = const { std::cell::Cell::new(0) };
        }
        fn count() {
            CALLS.with(|calls| calls.set(calls.get() + 1));
        }
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 0.0);
        for row in 1..20_000u32 {
            let prev = CellRef::new(row - 1, 0).to_a1();
            put_formula(
                &mut wb,
                s,
                &CellRef::new(row, 0).to_a1(),
                &format!("{prev}+1"),
            );
        }
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        set_progress_hook(Some(count));
        put_num(&mut wb, s, "A1", 1.0);
        recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        set_progress_hook(None);
        assert!(CALLS.with(|calls| calls.get()) >= 4);
    }

    #[test]
    fn progress_follows_the_cells_a_recalculation_reads() {
        thread_local! {
            static CALLS: std::cell::Cell<u32> = const { std::cell::Cell::new(0) };
        }
        fn count() {
            CALLS.with(|calls| calls.set(calls.get() + 1));
        }
        const ROWS: u32 = 100_000;
        let (mut wb, s) = one_sheet();
        for row in 0..ROWS {
            put_num(&mut wb, s, &CellRef::new(row, 0).to_a1(), f64::from(row));
        }
        for (at, last) in [("B1", ROWS), ("B2", ROWS - 1), ("B3", ROWS - 2)] {
            put_formula(&mut wb, s, at, &format!("SUM(A1:A{last})"));
        }
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        set_progress_hook(Some(count));
        put_num(&mut wb, s, "A1", 1.0);
        recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        set_progress_hook(None);
        let read = 3 * (ROWS - 1);
        assert!(CALLS.with(|calls| calls.get()) >= read / xlsx_model::progress::EVERY);
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
