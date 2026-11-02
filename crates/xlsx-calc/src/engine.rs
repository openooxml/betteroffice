//! recalc driver: given edited cells, re-evaluate exactly the formulas that
//! could have changed, in dependency order, and report what moved.

mod order;

#[cfg(test)]
mod criteria_index_differential;

use order::{Step, topo_order};

use std::borrow::Cow;
use std::cell::{Cell as Flag, RefCell};
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use xlsx_model::progress::tick as progress;
use xlsx_model::{
    CellProvider, CellRange, CellRef, CellValue, ColId, DefinedName, ErrorValue, RowId, SheetId,
    Table, Workbook,
};

use crate::array::{Spill, charged_spill, evaluate_spill};
use crate::deps::reads_other_workbook;
use crate::eval::{
    EvalContext, EvaluationBudget, HeldCells, MAX_RECALCULATION_CELL_VISITS, evaluate,
};
use crate::graph::{DepGraph, ReaderIndex, UnevaluatedCell};

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
    /// the recalculation's shared budget ran out, so formulas it refused kept
    /// their cached values or show `#NUM!` without one.
    pub budget_exhausted: bool,
    /// every formula cell the graph holds as unevaluated after this recalc,
    /// not only those it evaluated, in address order.
    pub unevaluated_cells: Vec<UnevaluatedCell>,
}

#[derive(Clone, Copy)]
struct RecalcContext {
    now_serial: Option<f64>,
    rand_seed: Option<u32>,
}

/// what every evaluation of one recalc or derivation shares besides its
/// budget.
#[derive(Clone, Copy)]
struct Scope<'g> {
    graph: &'g DepGraph,
    /// the graph again when its names are the workbook's, so evaluation looks
    /// names up there rather than scanning the workbook's table.
    name_index: Option<&'g DepGraph>,
    context: RecalcContext,
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

/// marks what a recalculation would leave unevaluated on a graph built over
/// values no recalculation over it produced, writing nothing: each formula
/// whose evaluation over the current values reaches an uncomputable function,
/// and every formula reading one or a cell its array result would fill,
/// through formulas the file stored no value for too, or reading one through
/// a reference it computes. other gaps of the engine show only when a
/// recalculation meets them. without a clock, a volatile formula counts every
/// uncomputable call it holds, since what it reaches may depend on the date.
pub fn derive_unevaluated(
    wb: &Workbook,
    graph: &DepGraph,
    now_serial: Option<f64>,
    rand_seed: Option<u32>,
) {
    let mut derivation = Derivation::new(now_serial, rand_seed);
    while !derivation.advance(wb, graph, usize::MAX).done {}
}

/// [`derive_unevaluated`] in steps a host bounds, over a workbook and graph
/// that stay the same between them. however the work is sliced, it marks the
/// same cells: the slices spend the one budget of cell visits a one-shot
/// derivation would, in the same order.
pub struct Derivation {
    context: RecalcContext,
    /// what the budget of cell visits and text bytes has left, and whether it
    /// turned work away.
    budget: (u64, u64, bool),
    currency: Option<crate::coercion::CurrencyCoercion>,
    /// whether the graph's names are the workbook's, and whether one of them
    /// may call an uncomputable function; known once the first step runs.
    names: Option<(bool, bool)>,
    phase: Phase,
    found: HashMap<Key, Vec<String>>,
    reaches: HashMap<Key, Option<CellRange>>,
    worklist: Vec<Key>,
    expanded: HashSet<Key>,
    spill_readers: HashMap<Key, Vec<(SheetId, CellRef)>>,
    /// the volatile formulas, in address order: the ones that may read cells
    /// through references they compute.
    dynamic: Vec<Key>,
}

enum Phase {
    /// scanning each sheet's formulas after the last one scanned for those
    /// whose own evaluation reaches an uncomputable call.
    Sources {
        sheet: usize,
        after: Option<CellRef>,
    },
    /// passing what is found on to the readers the graph knows.
    Spread,
    /// evaluating the volatile formulas not found yet against what is; one
    /// that reads a found cell is found and spreads in turn.
    Dynamic {
        next: usize,
        added: bool,
    },
    Done,
}

/// what one [`Derivation::advance`] did.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DerivationStep {
    /// the derivation finished and marked the graph.
    pub done: bool,
    /// the steps it took: each looks at one cell, evaluates one formula or
    /// passes one found cell on to its readers, within the per-formula and
    /// per-derivation budgets of cell visits a recalculation has.
    pub work: usize,
}

impl Derivation {
    pub fn new(now_serial: Option<f64>, rand_seed: Option<u32>) -> Self {
        Self {
            context: RecalcContext {
                now_serial,
                rand_seed,
            },
            budget: EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS).left(),
            currency: None,
            names: None,
            phase: Phase::Sources {
                sheet: 0,
                after: None,
            },
            found: HashMap::new(),
            reaches: HashMap::new(),
            worklist: Vec::new(),
            expanded: HashSet::new(),
            spill_readers: HashMap::new(),
            dynamic: Vec::new(),
        }
    }

    /// runs up to `max_work` steps, at least one, or until the derivation is
    /// done.
    pub fn advance(&mut self, wb: &Workbook, graph: &DepGraph, max_work: usize) -> DerivationStep {
        let (matches, _) = *self.names.get_or_insert_with(|| {
            (
                graph.names_match(&wb.defined_names),
                wb.defined_names
                    .iter()
                    .any(|defined| crate::functions::may_call_uncomputable(&defined.formula)),
            )
        });
        let scope = Scope {
            graph,
            name_index: matches.then_some(graph),
            context: self.context,
        };
        let (cells, text, refused) = self.budget;
        let mut budget = EvaluationBudget::resumed(cells, text, refused);
        budget.currency_coercion = self
            .currency
            .get_or_insert_with(|| crate::coercion::CurrencyCoercion::from_workbook(wb))
            .clone();
        let budget = Rc::new(budget);
        let mut readers = Readers::new(graph);
        let mut work = 0usize;
        while !matches!(self.phase, Phase::Done) && (work == 0 || work < max_work) {
            self.step(wb, scope, &budget, &mut readers);
            work += 1;
        }
        self.budget = budget.left();
        let done = matches!(self.phase, Phase::Done);
        if done {
            graph.assume_unevaluated(
                std::mem::take(&mut self.found)
                    .into_iter()
                    .filter(|&(at, _)| graph.is_formula(at.0, cell_of(at)))
                    .map(|(at, functions)| UnevaluatedCell {
                        sheet: at.0,
                        cell: cell_of(at),
                        functions,
                    }),
            );
        }
        DerivationStep { done, work }
    }

    /// one step of the derivation.
    fn step(
        &mut self,
        wb: &Workbook,
        scope: Scope<'_>,
        budget: &Rc<EvaluationBudget>,
        readers: &mut Readers<'_>,
    ) {
        match self.phase {
            Phase::Sources { sheet, after } => self.scan(wb, scope, budget, sheet, after),
            Phase::Spread => self.spread(wb, scope, budget, readers),
            Phase::Dynamic { next, added } => self.revisit(wb, scope, budget, next, added),
            Phase::Done => {}
        }
    }

    /// looks at the cell on `sheet` after `after`: a volatile formula joins
    /// [`Derivation::dynamic`], and one that may call an uncomputable function
    /// is evaluated.
    fn scan(
        &mut self,
        wb: &Workbook,
        scope: Scope<'_>,
        budget: &Rc<EvaluationBudget>,
        sheet: usize,
        after: Option<CellRef>,
    ) {
        let graph = scope.graph;
        let Some(cells) = wb.sheets.get(sheet) else {
            self.phase = Phase::Spread;
            return;
        };
        let start = after.map_or(CellRef::new(0, 0), |at| {
            CellRef::new(at.row, at.col.saturating_add(1))
        });
        let row = cells.cells_in_range(CellRange::new(start, CellRef::new(start.row, ColId::MAX)));
        let later = cells.cells_in_range(CellRange::new(
            CellRef::new(start.row.saturating_add(1), 0),
            CellRef::new(RowId::MAX, ColId::MAX),
        ));
        let Some((at, cell)) = row.chain(later).find(|&(at, _)| Some(at) != after) else {
            self.phase = Phase::Sources {
                sheet: sheet + 1,
                after: None,
            };
            return;
        };
        self.phase = Phase::Sources {
            sheet,
            after: Some(at),
        };
        let id = SheetId(sheet as u32);
        if cell.formula.is_none() || !graph.is_formula(id, at) {
            return;
        }
        let u = key(id, at);
        if graph.is_volatile_cell(id, at) {
            self.dynamic.push(u);
        }
        let names_may_call = self.names.is_some_and(|(_, may_call)| may_call);
        let may_call = cell
            .formula
            .as_deref()
            .is_some_and(crate::functions::may_call_uncomputable);
        if !(may_call || names_may_call) || graph.uncomputable_calls_at(id, at).is_empty() {
            return;
        }
        budget.begin_spill_round();
        let evaluated = eval_node(wb, u, scope, Rc::clone(budget), None);
        if let Some(functions) = evaluated.gap.filter(|functions| !functions.is_empty()) {
            self.reaches.insert(u, evaluated.reach);
            self.found.insert(u, functions);
            self.worklist.push(u);
        }
    }

    /// passes one found cell's functions on to its readers, and to the
    /// readers of the cells its array result would fill.
    fn spread(
        &mut self,
        wb: &Workbook,
        scope: Scope<'_>,
        budget: &Rc<EvaluationBudget>,
        readers: &mut Readers<'_>,
    ) {
        let graph = scope.graph;
        let Some(at) = self.worklist.pop() else {
            self.phase = if self.found.is_empty() || self.dynamic.is_empty() {
                Phase::Done
            } else {
                Phase::Dynamic {
                    next: 0,
                    added: false,
                }
            };
            return;
        };
        if !self.expanded.insert(at) && !budget.consume(1) {
            return;
        }
        progress();
        let functions = self.found.get(&at).cloned().unwrap_or_default();
        let mut dependents = Vec::new();
        readers.dependents_of(at, &mut dependents);
        let reaches = &mut self.reaches;
        let filled = self.spill_readers.entry(at).or_insert_with(|| {
            budget.begin_spill_round();
            reaches
                .remove(&at)
                .unwrap_or_else(|| would_fill(wb, at, scope, budget))
                .filter(|range| {
                    range.start != range.end && graph.spill_of(at.0, cell_of(at)) != Some(*range)
                })
                .map(|range| {
                    graph
                        .readers_of_range(at.0, range)
                        .filter(|&(sheet, cell)| key(sheet, cell) != at)
                        .collect()
                })
                .unwrap_or_default()
        });
        dependents.extend(filled.iter().copied());
        for (sheet, cell) in dependents {
            let reader = key(sheet, cell);
            let inherited = self.found.entry(reader).or_default();
            let known = inherited.len();
            inherited.extend(functions.iter().cloned());
            inherited.sort_unstable();
            inherited.dedup();
            if inherited.len() != known {
                self.worklist.push(reader);
            }
        }
    }

    /// evaluates the next volatile formula not found yet, watching what it
    /// reads; after the last, spreads what this pass found and starts another
    /// pass, until one finds nothing or the budget runs out.
    fn revisit(
        &mut self,
        wb: &Workbook,
        scope: Scope<'_>,
        budget: &Rc<EvaluationBudget>,
        next: usize,
        added: bool,
    ) {
        let Some(&u) = self.dynamic.get(next) else {
            self.phase = if added { Phase::Spread } else { Phase::Done };
            return;
        };
        self.phase = Phase::Dynamic {
            next: next + 1,
            added,
        };
        if self.found.contains_key(&u) {
            return;
        }
        if !budget.consume(1) {
            self.phase = Phase::Done;
            return;
        }
        budget.begin_spill_round();
        let watch = DerivedWatch {
            graph: scope.graph,
            found: &self.found,
        };
        let evaluated = eval_node(wb, u, scope, Rc::clone(budget), Some(&watch));
        if let Some(functions) = evaluated.gap.filter(|functions| !functions.is_empty()) {
            self.reaches.insert(u, evaluated.reach);
            self.found.insert(u, functions);
            self.worklist.push(u);
            self.phase = Phase::Dynamic {
                next: next + 1,
                added: true,
            };
        }
    }
}

/// what a derivation has found so far, as a formula reading through a
/// reference it computes sees it.
struct DerivedWatch<'w> {
    graph: &'w DepGraph,
    found: &'w HashMap<Key, Vec<String>>,
}

impl HeldCells for DerivedWatch<'_> {
    fn held(&self, sheet: SheetId, at: CellRef) -> Option<Vec<String>> {
        let mut functions = Vec::new();
        for source in holders(self.graph, sheet, at) {
            if let Some(found) = self.found.get(&source) {
                functions.extend(found.iter().cloned());
            }
        }
        functions.sort_unstable();
        functions.dedup();
        (!functions.is_empty()).then_some(functions)
    }
}

/// the formulas whose value a read of the cell at `at` takes: its own, and
/// any array formula the graph knows to spill over it.
fn holders(graph: &DepGraph, sheet: SheetId, at: CellRef) -> impl Iterator<Item = Key> + '_ {
    let spilled = graph
        .spill_sources(sheet, CellRange::new(at, at))
        .filter(move |&(_, anchor)| anchor != at)
        .map(|(sheet, anchor)| key(sheet, anchor));
    std::iter::once(key(sheet, at)).chain(spilled)
}

/// the rectangle an array formula's result would fill over the current
/// values, when `at` anchors one and the budget pays for finding out.
fn would_fill(
    wb: &Workbook,
    at: Key,
    scope: Scope<'_>,
    budget: &Rc<EvaluationBudget>,
) -> Option<CellRange> {
    wb.sheet(at.0)?.array_formula(cell_of(at))?;
    let evaluated = eval_node(wb, at, scope, Rc::clone(budget), None);
    match evaluated.value {
        Some(NodeValue::Spill(spill)) => Some(spill.range),
        _ => evaluated.reach,
    }
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
/// cell a round takes up joins `recomputed`. an array formula an uncomputable
/// call left unevaluated runs again, so the formulas reading the cells its
/// result would newly fill learn of the call.
fn run_recalc(
    wb: &mut Workbook,
    graph: &DepGraph,
    recompute: HashSet<Key>,
    context: RecalcContext,
    recomputed: Option<&mut HashSet<Key>>,
) -> RecalcResult {
    let mut budget = EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS);
    budget.currency_coercion = crate::coercion::CurrencyCoercion::from_workbook(wb);
    run_recalc_with_budget(wb, graph, recompute, context, recomputed, Rc::new(budget))
}

fn run_recalc_with_budget(
    wb: &mut Workbook,
    graph: &DepGraph,
    mut recompute: HashSet<Key>,
    context: RecalcContext,
    mut recomputed: Option<&mut HashSet<Key>>,
    budget: Rc<EvaluationBudget>,
) -> RecalcResult {
    {
        let mut previous = graph
            .currency_coercion
            .lock()
            .unwrap_or_else(|err| err.into_inner());
        // every builder records the evidence it was built over; a graph that
        // recorded none takes the current evidence as the one its values
        // were computed with
        if previous
            .as_ref()
            .is_some_and(|previous| *previous != budget.currency_coercion)
        {
            recompute.extend(graph.formula_cells().map(|(sheet, cell)| key(sheet, cell)));
        }
        *previous = Some(budget.currency_coercion.clone());
    }
    let mut pass = Pass {
        graph,
        name_index: graph.names_match(&wb.defined_names).then_some(graph),
        context,
        budget,
        changed: Vec::new(),
        cycle_cells: Vec::new(),
        limited_cells: Vec::new(),
        spilled: Vec::new(),
        readers: Readers::new(graph),
        inherited: HashMap::new(),
        replaced: HashMap::new(),
        rerun: Vec::new(),
        early: HashMap::new(),
        watching: false,
        unwatched: Vec::new(),
    };
    let mut pending = recompute;
    let mut anchors = Vec::new();
    for cell in graph.unevaluated_cells() {
        let at = key(cell.sheet, cell.cell);
        pass.watching |= !cell.functions.is_empty();
        if pending.contains(&at) || cell.functions.is_empty() {
            continue;
        }
        let anchor = wb
            .sheet(cell.sheet)
            .is_some_and(|sheet| sheet.array_formula(cell.cell).is_some());
        if anchor {
            anchors.push(at);
        } else {
            pass.taint_readers(at, &cell.functions);
        }
    }
    pending.extend(anchors);

    for round in 0..MAX_SPILL_ROUNDS {
        if round > 0 {
            pass.budget.begin_spill_round();
        }
        if let Some(recomputed) = &mut recomputed {
            recomputed.extend(pending.iter().copied());
        }
        *pass.budget.pending_formulas.borrow_mut() =
            pending.iter().map(|&u| (u.0, cell_of(u))).collect();
        for step in topo_order(graph, &pending) {
            match step {
                Step::Cell(u) => {
                    pass.evaluate(wb, u);
                    pass.budget
                        .pending_formulas
                        .borrow_mut()
                        .remove(&(u.0, cell_of(u)));
                }
                Step::Component(members) => {
                    pass.settle(wb, &members);
                    for u in members {
                        pass.budget
                            .pending_formulas
                            .borrow_mut()
                            .remove(&(u.0, cell_of(u)));
                    }
                }
            }
        }
        if pass.spilled.is_empty() && pass.rerun.is_empty() {
            break;
        }
        let rerun = std::mem::take(&mut pass.rerun);
        let mut seeds = std::mem::take(&mut pass.spilled);
        seeds.extend(rerun.iter().map(|&u| (u.0, cell_of(u))));
        pending = dependents_closure(graph, &seeds);
        pending.extend(rerun);
        if pending.is_empty() {
            break;
        }
    }

    let budget_exhausted = pass.budget.refused();
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
        budget_exhausted,
        unevaluated_cells: graph.unevaluated_cells(),
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
    /// the graph again when its names are the workbook's, so evaluation looks
    /// names up there rather than scanning the workbook's table.
    name_index: Option<&'g DepGraph>,
    context: RecalcContext,
    budget: Rc<EvaluationBudget>,
    changed: Vec<(SheetId, CellRef)>,
    cycle_cells: Vec<(SheetId, CellRef)>,
    limited_cells: Vec<(SheetId, CellRef)>,
    /// spilled cells whose value moved; the next round runs their readers.
    spilled: Vec<(SheetId, CellRef)>,
    readers: Readers<'g>,
    /// formulas reading a cell an uncomputable function left unevaluated,
    /// with those functions.
    inherited: HashMap<Key, Vec<String>>,
    /// what writes no gap reached replaced, by formula, as it was before this
    /// recalc: a later round that finds a gap reaching the formula puts it
    /// back. one entry per cell written, whatever the rounds.
    replaced: HashMap<Key, Replaced>,
    /// formulas reading cells an array result an uncomputable call kept from
    /// being written would newly fill, or reading through a computed
    /// reference a formula this round had not settled yet; the next round
    /// runs them again, so they learn what keeps those cells.
    rerun: Vec<Key>,
    /// by formula not yet evaluated this round, each formula that read it
    /// through a computed reference, with what the formula was held for then.
    early: HashMap<Key, HashMap<Key, Option<Vec<String>>>>,
    /// whether an uncomputable call holds any cell yet, so formulas reading
    /// through computed references watch their reads; until one does, they
    /// read as every formula does.
    watching: bool,
    /// the formulas reading through computed references this recalc
    /// evaluated before it watched; they run again once it does.
    unwatched: Vec<Key>,
}

/// what a formula the graph cannot know the reads of finds during a recalc:
/// the cells the graph holds unevaluated so far. it notes each formula it
/// reads before this round evaluated it, so the reader runs again should the
/// formula settle held otherwise.
struct RecalcWatch<'w> {
    graph: &'w DepGraph,
    budget: &'w EvaluationBudget,
    reader: Key,
    early: RefCell<Vec<(Key, Option<Vec<String>>)>>,
}

impl<'w> RecalcWatch<'w> {
    /// a watch for `reader` when the graph cannot know what it reads, a
    /// volatile formula: `INDIRECT`, a computed `OFFSET`, or one using them.
    fn over(graph: &'w DepGraph, budget: &'w EvaluationBudget, reader: Key) -> Option<Self> {
        graph
            .is_volatile_cell(reader.0, cell_of(reader))
            .then(|| Self {
                graph,
                budget,
                reader,
                early: RefCell::new(Vec::new()),
            })
    }
}

impl HeldCells for RecalcWatch<'_> {
    fn held(&self, sheet: SheetId, at: CellRef) -> Option<Vec<String>> {
        let mut functions = Vec::new();
        for source in holders(self.graph, sheet, at) {
            let held = self
                .graph
                .unevaluated_functions(source.0, cell_of(source))
                .filter(|functions| !functions.is_empty());
            if source != self.reader
                && self
                    .budget
                    .pending_formulas
                    .borrow()
                    .contains(&(source.0, cell_of(source)))
            {
                self.early.borrow_mut().push((source, held.clone()));
            }
            functions.extend(held.into_iter().flatten());
        }
        functions.sort_unstable();
        functions.dedup();
        (!functions.is_empty()).then_some(functions)
    }
}

/// what one formula's writes replaced: each cell's previous value, and the
/// rectangle an array formula recorded before.
struct Replaced {
    cells: HashMap<Key, CellValue>,
    range: Option<CellRange>,
}

/// what evaluating one formula node produced.
struct Evaluated {
    /// `None` keeps the cached value.
    value: Option<NodeValue>,
    /// the shared budget ran out during the evaluation.
    limited: bool,
    /// the uncomputable functions called, when an engine gap reached the
    /// result.
    gap: Option<Vec<String>>,
    /// the rectangle an array result would fill, when an uncomputable call
    /// kept it from being written.
    reach: Option<CellRange>,
}

impl<'g> Pass<'g> {
    /// evaluate a formula no cycle of the range graph passes through.
    fn evaluate(&mut self, wb: &mut Workbook, u: Key) {
        #[cfg(test)]
        FORMULA_EVALUATIONS.with(|count| count.set(count.get() + 1));
        progress();
        let authored = self.authored(wb, u);
        let budget = Rc::clone(&self.budget);
        let watch = self.watch(&budget, u);
        let evaluated = eval_node_with(
            wb,
            wb,
            u,
            authored,
            self.scope(),
            Rc::clone(&budget),
            watch.as_ref().map(|watch| watch as &dyn HeldCells),
        );
        self.note_early(watch);
        self.write(wb, u, evaluated);
    }

    /// the watch over what `u` reads, when it reads through computed
    /// references and an uncomputable call holds some cell.
    fn watch<'w>(&mut self, budget: &'w EvaluationBudget, u: Key) -> Option<RecalcWatch<'w>>
    where
        'g: 'w,
    {
        let watch = RecalcWatch::over(self.graph, budget, u)?;
        if self.watching {
            return Some(watch);
        }
        self.unwatched.push(u);
        None
    }

    /// keeps what `watch` saw of formulas this round had not evaluated yet.
    fn note_early(&mut self, watch: Option<RecalcWatch<'_>>) {
        let Some(watch) = watch else {
            return;
        };
        for (source, held) in watch.early.into_inner() {
            self.early
                .entry(source)
                .or_default()
                .entry(watch.reader)
                .or_insert(held);
        }
    }

    /// `u` settled, held for `held`: each formula that read it through a
    /// computed reference before, when it was held otherwise, runs again.
    fn release_early(&mut self, u: Key, held: Option<&[String]>) {
        let Some(readers) = self.early.remove(&u) else {
            return;
        };
        let before = self.rerun.len();
        for (reader, seen) in readers {
            if seen.as_deref() != held {
                self.rerun.push(reader);
            }
        }
        if self.rerun.len() != before {
            self.budget.schedule_spill_round();
        }
    }

    fn scope(&self) -> Scope<'g> {
        Scope {
            graph: self.graph,
            name_index: self.name_index,
            context: self.context,
        }
    }

    /// the rectangle an array formula at `u` recorded before this recalc,
    /// whatever an earlier round wrote since.
    fn authored(&self, wb: &Workbook, u: Key) -> Option<CellRange> {
        match self.replaced.get(&u) {
            Some(replaced) => replaced.range,
            None => wb.sheet(u.0)?.array_formula(cell_of(u)),
        }
    }

    /// passes a cell's uncomputable functions on to the formulas that read it.
    /// a gap of this engine alone passes nothing: what excel saved for the
    /// cell is its value, so its readers recalculate from that.
    fn taint_readers(&mut self, u: Key, functions: &[String]) {
        if functions.is_empty() {
            return;
        }
        let mut readers = Vec::new();
        self.readers.dependents_of(u, &mut readers);
        for (sheet, cell) in readers {
            let inherited = self.inherited.entry(key(sheet, cell)).or_default();
            inherited.extend(functions.iter().cloned());
            inherited.sort_unstable();
            inherited.dedup();
        }
    }

    /// passes an anchor's uncomputable functions on to the formulas reading
    /// the cells its result would fill, when the graph does not know that
    /// rectangle, and runs them again in the next round.
    fn taint_spill_readers(&mut self, u: Key, range: CellRange, functions: &[String]) {
        let graph = self.graph;
        if functions.is_empty()
            || range.start == range.end
            || graph.spill_of(u.0, cell_of(u)) == Some(range)
        {
            return;
        }
        for (sheet, cell) in graph.readers_of_range(u.0, range) {
            let reader = key(sheet, cell);
            if reader == u {
                continue;
            }
            let inherited = self.inherited.entry(reader).or_default();
            inherited.extend(functions.iter().cloned());
            inherited.sort_unstable();
            inherited.dedup();
            self.rerun.push(reader);
        }
        if !self.rerun.is_empty() {
            self.budget.schedule_spill_round();
        }
    }

    /// show an array formula the budget refused: an uncached anchor shows
    /// `#NUM!`, and a cached one, like the rest of its rectangle, keeps what
    /// it had.
    fn refuse(&mut self, wb: &mut Workbook, u: Key, log: Option<&mut Vec<(Key, CellValue)>>) {
        let refused = CellValue::Error {
            value: ErrorValue::Num,
        };
        if matches!(wb.value_cow(u.0, cell_of(u)).as_ref(), CellValue::Empty)
            && write_if_changed(wb, u, refused, &self.budget)
        {
            self.changed.push((u.0, cell_of(u)));
            if let Some(log) = log {
                log.push((u, CellValue::Empty));
            }
        }
    }

    /// puts back what an earlier evaluation of `u` replaced, as it was.
    fn restore(&mut self, wb: &mut Workbook, u: Key, replaced: Replaced) {
        let Some(sheet) = wb.sheet_mut(u.0) else {
            return;
        };
        for (at, value) in replaced.cells {
            self.budget.invalidate_criteria_cell(at.0, cell_of(at));
            sheet.set_calculated_value(cell_of(at), value);
            self.changed.push((at.0, cell_of(at)));
            if at != u {
                self.spilled.push((at.0, cell_of(at)));
            }
        }
        match replaced.range {
            Some(range) => sheet.set_array_formula(cell_of(u), range),
            None => sheet.clear_array_formula(cell_of(u)),
        }
        if !self.spilled.is_empty() {
            self.budget.schedule_spill_round();
        }
    }

    /// settle the members of one component of the range graph by what they
    /// actually read. a member that reads another before it settles waits on
    /// it: that one runs next, and the waiting member runs again once it has
    /// settled. the members waiting form one chain, so a member that reads
    /// one on the chain closes a cycle of reads: the members from that one on
    /// are circular and settle at `circular_value`, and the member waiting on
    /// the first of them runs again with that value. a cycle one of whose
    /// members holds or reached an uncomputable call settles at nothing: what
    /// it would settle at hangs on values no engine computes, so every member
    /// keeps what it holds and is listed with the calls.
    fn settle(&mut self, wb: &mut Workbook, members: &[Key]) {
        self.budget.set_settling(true);
        let ids: HashMap<Key, usize> = members.iter().enumerate().map(|(id, &u)| (u, id)).collect();
        let spills = Spills::new(wb, members);
        let mut reached: Vec<Vec<String>> = vec![Vec::new(); members.len()];
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
                let budget = Rc::clone(&self.budget);
                let watch = self.watch(&budget, u);
                let evaluated = eval_node_with(
                    &log,
                    wb,
                    u,
                    self.authored(wb, u),
                    self.scope(),
                    Rc::clone(&budget),
                    watch.as_ref().map(|watch| watch as &dyn HeldCells),
                );
                let target = log.first.get();
                self.note_early(watch);
                reached[id].extend(evaluated.gap.iter().flatten().cloned());
                let Some(target) = target else {
                    self.write(wb, u, evaluated);
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
                let circular: Vec<usize> = chain.drain(at..).collect();
                let mut functions: Vec<String> = Vec::new();
                for &member in &circular {
                    let u = members[member];
                    functions.extend(reached[member].iter().cloned());
                    functions.extend(self.graph.uncomputable_calls_at(u.0, cell_of(u)));
                    functions.extend(self.inherited.get(&u).into_iter().flatten().cloned());
                }
                functions.sort_unstable();
                functions.dedup();
                for member in circular {
                    #[cfg(test)]
                    SCHEDULER_STEPS.with(|count| count.set(count.get() + 1));
                    let u = members[member];
                    settled[member] = true;
                    on_chain[member] = None;
                    if !functions.is_empty() {
                        let kept = Evaluated {
                            value: None,
                            limited: false,
                            gap: Some(functions.clone()),
                            reach: None,
                        };
                        self.write(wb, u, kept);
                        continue;
                    }
                    if write_if_changed(wb, u, circular_value(wb, u), &self.budget) {
                        self.changed.push((u.0, cell_of(u)));
                    }
                    self.graph.set_unevaluated(u.0, cell_of(u), None);
                    self.release_early(u, None);
                    self.cycle_cells.push((u.0, cell_of(u)));
                }
            }
        }
        self.budget.set_settling(false);
    }

    /// a formula an uncomputable call reaches, through its own evaluation or
    /// a cell it reads, is never written: it keeps what it holds, excel's
    /// value or a blank, and its readers learn of the call. once a gap keeps a
    /// formula's value, what an earlier round of this recalc wrote for it is
    /// put back as it was before the recalc.
    fn write(&mut self, wb: &mut Workbook, u: Key, evaluated: Evaluated) {
        let Evaluated {
            mut value,
            limited,
            mut gap,
            mut reach,
        } = evaluated;
        if let Some(inherited) = self.inherited.get(&u) {
            let functions = gap.get_or_insert_with(Vec::new);
            functions.extend(inherited.iter().cloned());
            functions.sort_unstable();
            functions.dedup();
        }
        if gap.as_ref().is_some_and(|functions| !functions.is_empty())
            && let Some(NodeValue::Spill(spill)) = value.take()
        {
            reach = Some(spill.range);
        }
        if gap.is_some()
            && value.is_none()
            && let Some(replaced) = self.replaced.remove(&u)
        {
            self.restore(wb, u, replaced);
        }
        if let Some(functions) = &gap {
            self.taint_readers(u, functions);
            if let Some(range) = reach {
                self.taint_spill_readers(u, range, functions);
            }
        }
        let at = (u.0, cell_of(u));
        let mut log = (gap.is_none() && value.is_some()).then(Vec::new);
        if limited {
            self.limited_cells.push(at);
        }
        let range = wb.sheet(u.0).and_then(|sheet| sheet.array_formula(at.1));
        match value {
            Some(NodeValue::Scalar(value)) => {
                if let Some(previous) = replace_if_changed(wb, u, value, &self.budget) {
                    self.changed.push(at);
                    if let Some(log) = &mut log {
                        log.push((u, previous));
                    }
                }
            }
            Some(NodeValue::Spill(spill)) => {
                match write_spill(wb, u, spill, &self.budget, &mut self.changed, log.as_mut()) {
                    Some(moved) => {
                        if !moved.is_empty() {
                            self.budget.schedule_spill_round();
                            self.spilled.extend(moved);
                        }
                    }
                    None => {
                        if !limited {
                            self.limited_cells.push(at);
                        }
                        if let Some(held) = held_gap(self.graph, u, true, false) {
                            let kept = Evaluated {
                                value: None,
                                limited: false,
                                gap: Some(held),
                                reach: None,
                            };
                            self.write(wb, u, kept);
                            return;
                        }
                        self.refuse(wb, u, log.as_mut());
                    }
                }
            }
            Some(NodeValue::Refused) => self.refuse(wb, u, log.as_mut()),
            None => {}
        }
        let held = gap.clone().filter(|functions| !functions.is_empty());
        self.graph.set_unevaluated(u.0, at.1, gap);
        self.release_early(u, held.as_deref());
        if held.is_some() && !self.watching {
            self.watching = true;
            let unwatched = std::mem::take(&mut self.unwatched);
            self.rerun
                .extend(unwatched.into_iter().filter(|&reader| reader != u));
            if !self.rerun.is_empty() {
                self.budget.schedule_spill_round();
            }
        }
        let reshaped = wb.sheet(u.0).and_then(|sheet| sheet.array_formula(at.1)) != range;
        if let Some(log) = log.filter(|log| !log.is_empty() || reshaped) {
            let replaced = self.replaced.entry(u).or_insert_with(|| Replaced {
                cells: HashMap::new(),
                range,
            });
            for (at, previous) in log {
                replaced.cells.entry(at).or_insert(previous);
            }
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
    fn workbook(&self) -> Option<&Workbook> {
        Some(self.inner)
    }

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
        let spilled = self.inner.sheet(sheet).map_or(0, |sheet| {
            sheet
                .array_formulas()
                .map(|(_, range)| range.end.row.saturating_add(1))
                .max()
                .unwrap_or(0)
        });
        self.inner.used_rows(sheet).max(spilled)
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
/// has no formula, it no longer parses, it reads another workbook, or an engine
/// gap reached the result.
fn eval_node(
    wb: &Workbook,
    u: Key,
    scope: Scope<'_>,
    budget: Rc<EvaluationBudget>,
    watch: Option<&dyn HeldCells>,
) -> Evaluated {
    let authored = wb
        .sheet(u.0)
        .and_then(|sheet| sheet.array_formula(cell_of(u)));
    eval_node_with(wb, wb, u, authored, scope, budget, watch)
}

/// evaluate one node, reading cells through `provider` while `wb` supplies the
/// formula's own shape; `authored` is the rectangle an array formula recorded.
/// `watch` answers, for every cell the formula reads, which uncomputable calls
/// keep that cell, for a formula the graph cannot know the reads of.
fn eval_node_with(
    provider: &dyn CellProvider,
    wb: &Workbook,
    u: Key,
    authored: Option<CellRange>,
    scope: Scope<'_>,
    budget: Rc<EvaluationBudget>,
    watch: Option<&dyn HeldCells>,
) -> Evaluated {
    let Scope {
        graph,
        name_index,
        context,
    } = scope;
    let Some(expr) = graph
        .ast(u.0, cell_of(u))
        .filter(|expr| !reads_other_workbook(expr))
    else {
        return Evaluated {
            value: None,
            limited: false,
            gap: None,
            reach: None,
        };
    };
    #[cfg(feature = "test-counters")]
    EVALUATED_FORMULAS.with(|cells| {
        if let Some(cells) = &mut *cells.borrow_mut() {
            cells.push((u.0, cell_of(u)));
        }
    });
    let cell = cell_of(u);
    let mut ctx = EvalContext::with_budget(provider, u.0, budget);
    ctx.cell = Some(cell);
    ctx.now_serial = context.now_serial;
    ctx.rand_seed = context.rand_seed.map(|seed| cell_random_seed(seed, u));
    ctx.date_system = wb.date_system;
    ctx.parse_cache = Some(graph.asts());
    ctx.name_index = name_index;
    ctx.spill_index = Some(graph);
    ctx.held_cells = watch;
    let value = match authored {
        Some(_) => NodeValue::Spill(evaluate_spill(&expr, &ctx, cell, authored)),
        None => NodeValue::Scalar(computed(evaluate(&expr, &ctx))),
    };
    if ctx.read_other_workbook() {
        return Evaluated {
            value: None,
            limited: ctx.exhausted(),
            gap: None,
            reach: None,
        };
    }
    let unsupported = ctx.has_unhandled_unsupported_function();
    let refused = ctx.has_unhandled_budget_error();
    let mut gap = unsupported.then(|| ctx.uncomputable_functions());
    let undated = context.now_serial.is_none() && graph.is_volatile_cell(u.0, cell);
    let held = held_gap(graph, u, refused, undated);
    if held.is_some() || gap.as_ref().is_some_and(|functions| !functions.is_empty()) {
        if let Some(held) = held {
            let functions = gap.get_or_insert_with(Vec::new);
            functions.extend(held);
            functions.sort_unstable();
            functions.dedup();
        }
        let reach = match value {
            NodeValue::Spill(spill) => Some(spill.range),
            NodeValue::Scalar(_) | NodeValue::Refused => None,
        };
        return Evaluated {
            value: None,
            limited: ctx.exhausted(),
            gap,
            reach,
        };
    }
    if (refused || unsupported) && !matches!(wb.value_cow(u.0, cell).as_ref(), CellValue::Empty) {
        return Evaluated {
            value: None,
            limited: ctx.exhausted(),
            gap,
            reach: None,
        };
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
        return Evaluated {
            value: Some(refused),
            limited: ctx.exhausted(),
            gap,
            reach: None,
        };
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
    Evaluated {
        value: Some(value),
        limited: ctx.exhausted(),
        gap,
        reach: None,
    }
}

/// what keeps a formula whose evaluation may not have reached its
/// uncomputable calls: when the budget stopped it, every call it holds and
/// what it was already held for; when it is volatile and there is no clock,
/// every call it holds, since what it reaches may depend on the date. `None`
/// when nothing does.
fn held_gap(graph: &DepGraph, u: Key, refused: bool, undated: bool) -> Option<Vec<String>> {
    if !refused && !undated {
        return None;
    }
    let mut held = graph.uncomputable_calls_at(u.0, cell_of(u));
    if refused && let Some(listed) = graph.unevaluated_functions(u.0, cell_of(u)) {
        held.extend(listed);
        held.sort_unstable();
        held.dedup();
        return Some(held);
    }
    (!held.is_empty()).then_some(held)
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
    mut log: Option<&mut Vec<(Key, CellValue)>>,
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
            budget,
            changed,
            &mut moved,
            log.as_deref_mut(),
        );
    }
    for ((row, col), value) in cells_of(range).zip(values) {
        if row == anchor.row && col == anchor.col {
            if let Some(previous) = replace_if_changed(wb, u, value, budget) {
                changed.push((u.0, anchor));
                if let Some(log) = log.as_deref_mut() {
                    log.push((u, previous));
                }
            }
            continue;
        }
        write_spilled_cell(
            wb,
            (u.0, row, col),
            value,
            budget,
            changed,
            &mut moved,
            log.as_deref_mut(),
        );
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
    budget: &EvaluationBudget,
    changed: &mut Vec<(SheetId, CellRef)>,
    moved: &mut Vec<(SheetId, CellRef)>,
    log: Option<&mut Vec<(Key, CellValue)>>,
) {
    if owns_formula(wb, u.0, cell_of(u)) {
        return;
    }
    if let Some(previous) = replace_if_changed(wb, u, value, budget) {
        changed.push((u.0, cell_of(u)));
        moved.push((u.0, cell_of(u)));
        if let Some(log) = log {
            log.push((u, previous));
        }
    }
}

fn cells_of(range: CellRange) -> impl Iterator<Item = (RowId, ColId)> {
    (range.start.row..=range.end.row)
        .flat_map(move |row| (range.start.col..=range.end.col).map(move |col| (row, col)))
}

/// write `value` only if it differs from the stored value; returns whether
/// anything changed. formula and style are preserved.
fn write_if_changed(
    wb: &mut Workbook,
    u: Key,
    value: CellValue,
    budget: &EvaluationBudget,
) -> bool {
    replace_if_changed(wb, u, value, budget).is_some()
}

/// [`write_if_changed`], returning the value it replaced.
fn replace_if_changed(
    wb: &mut Workbook,
    u: Key,
    value: CellValue,
    budget: &EvaluationBudget,
) -> Option<CellValue> {
    let value = precision_as_displayed(wb, u, stored(value));
    let previous = wb.value_cow(u.0, cell_of(u));
    let negative_zero = !wb.full_precision
        && matches!(previous.as_ref(), CellValue::Number { value } if *value == 0.0 && value.is_sign_negative())
        && matches!(&value, CellValue::Number { value } if *value == 0.0 && value.is_sign_positive());
    if *previous == value && !negative_zero {
        return None;
    }
    budget.invalidate_criteria_cell(u.0, cell_of(u));
    Some(wb.sheet_mut(u.0).map_or(CellValue::Empty, |sheet| {
        sheet.set_calculated_value(cell_of(u), value)
    }))
}

/// a result as excel stores it: below the smallest normal double it is 0,
/// while the math that produced it kept every bit.
fn stored(value: CellValue) -> CellValue {
    match value {
        CellValue::Number { value } if value != 0.0 && value.abs() < f64::MIN_POSITIVE => {
            CellValue::Number { value: 0.0 }
        }
        value => value,
    }
}

fn precision_as_displayed(wb: &Workbook, u: Key, value: CellValue) -> CellValue {
    if wb.full_precision {
        return value;
    }
    let CellValue::Number { value: number } = value else {
        return value;
    };
    let style = wb
        .sheet(u.0)
        .and_then(|sheet| sheet.effective_style(cell_of(u)));
    let code = match wb.styles.resolved_format(style).number_format {
        xlsx_model::styles::FormatCode::Custom(code) => Some(code),
        xlsx_model::styles::FormatCode::Builtin(id) => xlsx_model::numfmt::builtin_format_code(id),
    };
    let Some(digits) = code.and_then(|code| xlsx_model::numfmt::displayed_precision(number, code))
    else {
        return value;
    };
    let rounded = crate::functions::math::round_number(number, digits);
    if rounded.is_finite() {
        CellValue::Number {
            value: if rounded == 0.0 { 0.0 } else { rounded },
        }
    } else {
        value
    }
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
        for formula in ["1", "REPRO_UNKNOWN_FUNCTION(1)"] {
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
        put_recorded(
            &mut wb,
            s,
            "C20",
            "{1;2}+REPRO_UNKNOWN_FUNCTION(1)",
            "C20:C262163",
        );
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

    /// `WEBSERVICE` is a function no offline engine computes, so neither the
    /// formula calling it nor one reading that formula is written.
    #[test]
    fn an_unimplemented_function_keeps_the_cached_value() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 7.0);
        put_cached_formula(&mut wb, s, "B1", "WEBSERVICE(A1)/10", num(4.9));
        put_formula(&mut wb, s, "C1", "B1*2");
        let r = rebuild_and_recalc_all(&mut wb, None).1;
        assert_eq!(value(&wb, s, "B1"), num(4.9));
        assert_eq!(value(&wb, s, "C1"), CellValue::Empty);
        assert!(!changed_a1(&r).contains(&"B1".to_string()));
    }

    #[test]
    fn an_unknown_function_without_a_cached_value_reports_name() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "REPRO_UNKNOWN_FUNCTION(1)");
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
    fn indirect_in_r1c1_style_resolves_like_a1() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 5.0);
        put_cached_formula(&mut wb, s, "B1", "INDIRECT(\"R1C1\",FALSE)", num(8.0));
        put_formula(&mut wb, s, "B2", "B1+1");
        put_formula(&mut wb, s, "B3", "INDIRECT(\"R1C1\",FALSE)");
        put_cached_formula(&mut wb, s, "B4", "INDIRECT(\"A1\")", num(-1.0));
        put_cached_formula(&mut wb, s, "B5", "INDIRECT(\"Nowhere!A1\")", num(-1.0));
        put_cached_formula(&mut wb, s, "B6", "INDIRECT(\"'[1]Book'!Rate\")", num(0.25));
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(5.0));
        assert_eq!(value(&wb, s, "B2"), num(6.0));
        let reference = CellValue::Error {
            value: ErrorValue::Ref,
        };
        assert_eq!(value(&wb, s, "B3"), num(5.0));
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

    /// a year-first date reads alike in every locale, so arithmetic coerces it
    /// to its serial, time of day included, under the workbook's date system.
    #[test]
    fn arithmetic_on_year_first_date_text_reads_its_serial() {
        let (mut wb, s) = one_sheet();
        put_text(&mut wb, s, "A1", "2017-04-27");
        put_num(&mut wb, s, "A2", 41_309.0);
        put_text(&mut wb, s, "A3", "2017/4/27 18:00");
        put_text(&mut wb, s, "A4", "2017-02-30");
        put_formula(&mut wb, s, "B1", "A1-A2");
        put_formula(&mut wb, s, "B2", "A3-A1");
        put_formula(&mut wb, s, "B3", "-A1");
        put_formula(&mut wb, s, "B4", "A1%");
        put_formula(&mut wb, s, "B5", "A4+1");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(1_543.0));
        assert_eq!(value(&wb, s, "B2"), num(0.75));
        assert_eq!(value(&wb, s, "B3"), num(-42_852.0));
        assert_eq!(value(&wb, s, "B4"), num(428.52));
        assert_eq!(
            value(&wb, s, "B5"),
            CellValue::Error {
                value: ErrorValue::Value
            }
        );

        let (mut wb, s) = one_sheet();
        wb.date_system = xlsx_model::DateSystem::V1904;
        put_text(&mut wb, s, "A1", "2017-04-27");
        put_formula(&mut wb, s, "B1", "A1+0");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(41_390.0));
    }

    /// a TEXT conversion that fails must not poison the cell with an
    /// unsupported flag: the cell and its dependents recalc normally.
    #[test]
    fn a_failed_text_conversion_recalcs_after_an_edit() {
        let (mut wb, s) = one_sheet();
        let text = |v: &str| CellValue::Text { value: v.into() };
        put_text(&mut wb, s, "A1", "99-1");
        put_num(&mut wb, s, "B1", 5.0);
        put_cached_formula(&mut wb, s, "C1", "TEXT(A1,\"0\")&B1", text("99-15"));
        put_cached_formula(&mut wb, s, "D1", "C1", text("99-15"));
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "C1"), text("99-15"));

        put_num(&mut wb, s, "B1", 10.0);
        recalc_after(&mut wb, &mut graph, &[(s, a1("B1"))], None);
        assert_eq!(value(&wb, s, "C1"), text("99-110"));
        assert_eq!(value(&wb, s, "D1"), text("99-110"));
    }

    /// date text TEXT converts in a 1904 workbook lands on the same date.
    #[test]
    fn text_converts_date_text_in_a_1904_workbook() {
        let (mut wb, s) = one_sheet();
        wb.date_system = xlsx_model::DateSystem::V1904;
        put_text(&mut wb, s, "A1", "November 2025");
        put_formula(&mut wb, s, "B1", "TEXT(A1,\"MMMM YYYY \")");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(
            value(&wb, s, "B1"),
            CellValue::Text {
                value: "November 2025 ".into()
            }
        );
        put_text(&mut wb, s, "A2", "November 1902");
        put_formula(&mut wb, s, "B2", "TEXT(A2,\"MMMM YYYY \")");
        rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(
            value(&wb, s, "B2"),
            CellValue::Text {
                value: "November 1902".into()
            }
        );
    }

    /// excel trims only ascii spaces off date text, so other unicode spaces,
    /// full-width digits and number-like clock junk leave it text.
    #[test]
    fn year_first_date_text_with_foreign_spacing_stays_text() {
        let (mut wb, s) = one_sheet();
        let texts = [
            "\u{a0}2024-03-05",
            "2024-03-05\u{a0}",
            "\u{3000}2024-03-05",
            "2024-03-05\u{2009}18:00",
            "2024-03-05 \u{a0}18:00",
            "2024-03-05 18:00\u{a0}",
            "\u{ff12}024-03-05",
            "2024-03-05 1e1:00",
            "2024-03-05 +6:00",
        ];
        for (row, text) in texts.iter().enumerate() {
            let row = row + 1;
            put_text(&mut wb, s, &format!("A{row}"), text);
            put_cached_formula(
                &mut wb,
                s,
                &format!("B{row}"),
                &format!("A{row}+0"),
                CellValue::Error {
                    value: ErrorValue::Value,
                },
            );
        }
        put_text(&mut wb, s, "C1", "  2024-03-05  ");
        put_text(&mut wb, s, "C2", "2024-03-05   18:00 ");
        put_formula(&mut wb, s, "D1", "C1+0");
        put_formula(&mut wb, s, "D2", "C2+0");
        rebuild_and_recalc_all(&mut wb, None);
        for (row, text) in texts.iter().enumerate() {
            assert_eq!(
                value(&wb, s, &format!("B{}", row + 1)),
                CellValue::Error {
                    value: ErrorValue::Value
                },
                "{text:?}"
            );
        }
        assert_eq!(value(&wb, s, "D1"), num(45_356.0));
        assert_eq!(value(&wb, s, "D2"), num(45_356.75));
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

    /// IFERROR answers for the unknown name it wraps, as excel's `#NAME?`
    /// does, so the gap never reaches the result and the computed value must
    /// replace the cache.
    #[test]
    fn a_handler_that_answers_an_unknown_call_writes_its_result() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(
            &mut wb,
            s,
            "B1",
            "IFERROR(REPRO_UNKNOWN_FUNCTION(1),0)",
            num(99.0),
        );
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(0.0));
        assert!(result.unevaluated_cells.is_empty());
    }

    fn unevaluated(result: &RecalcResult) -> Vec<(String, Vec<String>)> {
        result
            .unevaluated_cells
            .iter()
            .map(|cell| (cell.cell.to_a1(), cell.functions.clone()))
            .collect()
    }

    fn reported(cells: &[(&str, &[&str])]) -> Vec<(String, Vec<String>)> {
        cells
            .iter()
            .map(|(at, functions)| {
                (
                    (*at).to_string(),
                    functions.iter().map(|name| (*name).to_string()).collect(),
                )
            })
            .collect()
    }

    fn text(value: &str) -> CellValue {
        CellValue::Text {
            value: value.to_string(),
        }
    }

    /// no handler knows what an uncomputable call returned in excel, so the
    /// cache survives every wrapper and the cell is reported.
    #[test]
    fn a_handler_around_an_uncomputable_call_keeps_the_cache() {
        for formula in [
            "IFERROR(WEBSERVICE(\"http://x\"),\"fallback\")",
            "IFNA(_xlfn.WEBSERVICE(\"http://x\"),\"fallback\")",
            "IF(ISERROR(WEBSERVICE(\"http://x\")),\"fallback\",\"ok\")",
        ] {
            let (mut wb, s) = one_sheet();
            put_cached_formula(&mut wb, s, "B1", formula, text("live"));
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "B1"), text("live"), "{formula}");
            assert_eq!(
                unevaluated(&result),
                reported(&[("B1", &["WEBSERVICE"])]),
                "{formula}"
            );
        }
    }

    /// without a saved value there is nothing to keep, so an uncomputable
    /// call, a handler around one and a formula reading one stay blank.
    #[test]
    fn an_uncached_uncomputable_call_stays_blank_and_is_reported() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "A1", "WEBSERVICE(\"http://x\")");
        put_formula(
            &mut wb,
            s,
            "A2",
            "IFERROR(WEBSERVICE(\"http://x\"),\"fallback\")",
        );
        put_formula(&mut wb, s, "A3", "LEN(A1)");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        for at in ["A1", "A2", "A3"] {
            assert_eq!(value(&wb, s, at), CellValue::Empty, "{at}");
        }
        assert!(result.changed.is_empty());
        assert_eq!(
            unevaluated(&result),
            reported(&[
                ("A1", &["WEBSERVICE"]),
                ("A2", &["WEBSERVICE"]),
                ("A3", &["WEBSERVICE"]),
            ])
        );
    }

    /// a formula reading an unevaluated cell keeps its own cache, even after
    /// an edit recomputes it without recomputing the cell it reads.
    #[test]
    fn dependents_of_an_uncomputable_call_keep_their_caches() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "A1", "WEBSERVICE(\"http://x\")", num(5.0));
        put_cached_formula(&mut wb, s, "B1", "A1*2", num(11.0));
        put_num(&mut wb, s, "D1", 1.0);
        put_cached_formula(&mut wb, s, "C1", "B1+D1", num(12.0));
        put_formula(&mut wb, s, "E1", "A1*2");
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(11.0));
        assert_eq!(value(&wb, s, "C1"), num(12.0));
        assert_eq!(value(&wb, s, "E1"), CellValue::Empty);
        let all = reported(&[
            ("A1", &["WEBSERVICE"]),
            ("B1", &["WEBSERVICE"]),
            ("C1", &["WEBSERVICE"]),
            ("E1", &["WEBSERVICE"]),
        ]);
        assert_eq!(unevaluated(&result), all);

        put_num(&mut wb, s, "D1", 2.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        assert_eq!(value(&wb, s, "C1"), num(12.0));
        assert_eq!(unevaluated(&result), all);

        put_num(&mut wb, s, "A1", 3.0);
        graph.set_formula(s, a1("A1"), None);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("A1"))], None);
        assert_eq!(value(&wb, s, "B1"), num(6.0));
        assert_eq!(value(&wb, s, "C1"), num(8.0));
        assert!(result.unevaluated_cells.is_empty());
    }

    /// an unknown name is excel's own `#NAME?`: unhandled, the file's cache
    /// is the only value known for it, so it stays and is reported.
    #[test]
    fn an_unknown_function_keeps_its_cache_and_is_reported() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "A1", "REPRO_UNKNOWN_FUNCTION(1)", num(7.0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(7.0));
        assert_eq!(unevaluated(&result), reported(&[("A1", &[])]));
    }

    /// add-in and user-defined functions run code the file does not carry.
    #[test]
    fn add_in_functions_are_uncomputable() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "A1", "IFERROR(_xll.Quote(\"x\"),0)", num(3.0));
        put_cached_formula(&mut wb, s, "A2", "IFERROR(_xludf.Mine(1),0)", num(4.0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(3.0));
        assert_eq!(value(&wb, s, "A2"), num(4.0));
        assert_eq!(
            unevaluated(&result),
            reported(&[("A1", &["_xll.Quote"]), ("A2", &["_xludf.Mine"])])
        );
    }

    #[test]
    fn a_memoized_name_replays_its_uncomputable_call() {
        let (mut wb, s) = one_sheet();
        wb.defined_names.push(xlsx_model::DefinedName {
            name: "Live".to_string(),
            formula: "STOCKHISTORY(\"X\",1)".to_string(),
            local_sheet: None,
            hidden: false,
        });
        put_cached_formula(
            &mut wb,
            s,
            "B1",
            "IFERROR(Live,0)+IFERROR(Live,0)",
            num(99.0),
        );
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), num(99.0));
        assert_eq!(unevaluated(&result), reported(&[("B1", &["STOCKHISTORY"])]));
    }

    /// a graph built without evaluating learns the uncomputable calls its
    /// formulas reach and their readers, and a later edit keeps to them.
    #[test]
    fn unevaluated_cells_derive_without_writing() {
        let (mut wb, s) = one_sheet();
        wb.defined_names.push(DefinedName {
            name: "Live".to_string(),
            formula: "RTD(\"x\",,\"y\")".to_string(),
            local_sheet: None,
            hidden: false,
        });
        put_cached_formula(
            &mut wb,
            s,
            "A1",
            "IFERROR(_xlfn.WEBSERVICE(\"http://x\"),0)",
            num(5.0),
        );
        put_cached_formula(&mut wb, s, "B1", "A1*2", num(11.0));
        put_cached_formula(&mut wb, s, "C1", "Live+1", num(3.0));
        put_formula(&mut wb, s, "D1", "B1+1");
        put_cached_formula(&mut wb, s, "E1", "SUM(1,2)", num(3.0));
        put_cached_formula(&mut wb, s, "F1", "REPRO_UNKNOWN_FUNCTION(1)", num(1.0));
        put_num(&mut wb, s, "H1", 1.0);
        put_cached_formula(&mut wb, s, "G1", "B1+H1", num(12.0));
        put_cached_formula(
            &mut wb,
            s,
            "I1",
            "IF(H1>0,10,WEBSERVICE(\"http://x\"))",
            num(10.0),
        );
        put_cached_formula(&mut wb, s, "J1", "I1+H1", num(11.0));
        let mut graph = DepGraph::build(&wb);
        derive_unevaluated(&wb, &graph, None, None);
        let all = reported(&[
            ("A1", &["WEBSERVICE"]),
            ("B1", &["WEBSERVICE"]),
            ("C1", &["RTD"]),
            ("D1", &["WEBSERVICE"]),
            ("G1", &["WEBSERVICE"]),
        ]);
        let derived: Vec<_> = graph
            .unevaluated_cells()
            .iter()
            .map(|cell| (cell.cell.to_a1(), cell.functions.clone()))
            .collect();
        assert_eq!(derived, all);

        put_num(&mut wb, s, "H1", 2.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("H1"))], None);
        assert_eq!(value(&wb, s, "G1"), num(12.0));
        assert_eq!(value(&wb, s, "J1"), num(12.0));
        assert_eq!(unevaluated(&result), all);
    }

    /// a fallback the file stored no value for stays blank as its inputs
    /// change, and so does the formula reading it.
    #[test]
    fn an_uncached_fallback_stays_blank_as_its_inputs_change() {
        let (mut wb, s) = one_sheet();
        put_text(&mut wb, s, "C1", "v1");
        put_formula(&mut wb, s, "A1", "IFERROR(_xlfn.WEBSERVICE(C1),C1)");
        put_formula(&mut wb, s, "B1", "A1&\"!\"");
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), CellValue::Empty);
        assert_eq!(value(&wb, s, "B1"), CellValue::Empty);
        let both = reported(&[("A1", &["WEBSERVICE"]), ("B1", &["WEBSERVICE"])]);
        assert_eq!(unevaluated(&result), both);

        put_text(&mut wb, s, "C1", "v2");
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("C1"))], None);
        assert_eq!(value(&wb, s, "A1"), CellValue::Empty);
        assert_eq!(value(&wb, s, "B1"), CellValue::Empty);
        assert_eq!(unevaluated(&result), both);
    }

    /// an array condition evaluates only the branches its elements pick, so
    /// a gap in a branch none picks reaches nothing; the values are those
    /// both branches evaluated would give.
    #[test]
    fn an_array_if_evaluates_only_picked_branches() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(
            &mut wb,
            s,
            "A1",
            "SUM(IF({TRUE,TRUE},{1,2},WEBSERVICE(\"x\")))",
            num(999.0),
        );
        put_cached_formula(
            &mut wb,
            s,
            "A2",
            "SUM(IF({FALSE,FALSE},REPRO_UNKNOWN_FUNCTION(1),{3,4}))",
            num(999.0),
        );
        put_cached_formula(
            &mut wb,
            s,
            "A3",
            "SUM(IF({TRUE,FALSE},{1,2},WEBSERVICE(\"x\")))",
            num(999.0),
        );
        put_formula(
            &mut wb,
            s,
            "A4",
            "SUM(IF({TRUE,FALSE;FALSE,TRUE},{1,2;3,4},{10,20;30,40}))",
        );
        put_formula(&mut wb, s, "A5", "SUM(IF({1,0,\"x\"},{1,2,3},{10,20,30}))");
        put_formula(&mut wb, s, "A6", "SUM(IFERROR(IF({TRUE,FALSE},{1,2}),0))");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(3.0));
        assert_eq!(value(&wb, s, "A2"), num(7.0));
        assert_eq!(value(&wb, s, "A3"), num(999.0));
        assert_eq!(value(&wb, s, "A4"), num(55.0));
        assert_eq!(
            value(&wb, s, "A5"),
            CellValue::Error {
                value: ErrorValue::Value
            }
        );
        assert_eq!(value(&wb, s, "A6"), num(1.0));
        assert_eq!(unevaluated(&result), reported(&[("A3", &["WEBSERVICE"])]));
    }

    /// a cell that learns a function after it passed its others on passes it
    /// on again, through every formula of a long chain.
    #[test]
    fn derivation_carries_every_function_down_a_long_chain() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "A1", "WEBSERVICE(\"x\")", num(1.0));
        put_cached_formula(&mut wb, s, "B1", "RTD(\"x\",,\"y\")", num(1.0));
        put_cached_formula(&mut wb, s, "A2", "A1+B1", num(2.0));
        for row in 3..=5000 {
            let formula = format!("A{}+1", row - 1);
            put_cached_formula(&mut wb, s, &format!("A{row}"), &formula, num(row as f64));
        }
        let graph = DepGraph::build(&wb);
        derive_unevaluated(&wb, &graph, None, None);
        let cells = graph.unevaluated_cells();
        assert_eq!(cells.len(), 5001);
        let last = cells
            .iter()
            .find(|cell| cell.cell == a1("A5000"))
            .expect("chain end reported");
        assert_eq!(last.functions, ["RTD", "WEBSERVICE"]);
    }

    /// a formula computes again once no gap reaches it, and keeps what it
    /// then holds when one reaches it again.
    #[test]
    fn a_formula_no_gap_reaches_computes_again() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "D1", 1.0);
        put_formula(
            &mut wb,
            s,
            "A1",
            "IF(D1=1,IFERROR(_xlfn.WEBSERVICE(\"x\"),\"fb\"),\"plain\")",
        );
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), CellValue::Empty);
        assert_eq!(unevaluated(&result), reported(&[("A1", &["WEBSERVICE"])]));
        put_num(&mut wb, s, "D1", 0.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        assert_eq!(value(&wb, s, "A1"), text("plain"));
        assert!(result.unevaluated_cells.is_empty());
        put_num(&mut wb, s, "D1", 1.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        assert_eq!(value(&wb, s, "A1"), text("plain"));
        assert_eq!(unevaluated(&result), reported(&[("A1", &["WEBSERVICE"])]));
    }

    /// E1 is kept, so the uncached array A1 reading it is never written, nor
    /// the cells it would newly fill past the one it recorded, A2 among them.
    /// B1 and F1 read A2 and run before A1, B2 reads it and runs after: each
    /// keeps what it held, and C1 reading B1 keeps its cache too.
    #[test]
    fn a_tainted_spill_taints_readers_of_the_cells_it_would_newly_fill() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "E1", "_xlfn.WEBSERVICE(\"x\")", text("x"));
        put_formula(&mut wb, s, "A1", "IF(E1=\"x\",{10;20},0)");
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("A1"), xlsx_model::CellRange::parse_a1("A1").unwrap());
        put_num(&mut wb, s, "D1", 1.0);
        put_cached_formula(&mut wb, s, "B1", "A2+D1", num(99.0));
        put_formula(&mut wb, s, "F1", "A2*D1");
        put_formula(&mut wb, s, "G1", "D1*1");
        put_formula(&mut wb, s, "H1", "G1*1");
        put_cached_formula(&mut wb, s, "B2", "A2*H1", num(77.0));
        put_cached_formula(&mut wb, s, "C1", "B1+1", num(100.0));
        let graph = DepGraph::build(&wb);
        let all = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
        let order: Vec<_> = topo_order(&graph, &all)
            .into_iter()
            .map(|step| match step {
                Step::Cell(u) => cell_of(u).to_a1(),
                Step::Component(_) => unreachable!("no cycle"),
            })
            .collect();
        let position = |at: &str| order.iter().position(|cell| cell == at).unwrap();
        assert!(position("B1") < position("A1"));
        assert!(position("F1") < position("A1"));
        assert!(position("A1") < position("B2"));

        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), CellValue::Empty);
        assert_eq!(value(&wb, s, "A2"), CellValue::Empty);
        assert_eq!(value(&wb, s, "B1"), num(99.0));
        assert_eq!(value(&wb, s, "F1"), CellValue::Empty);
        assert_eq!(value(&wb, s, "B2"), num(77.0));
        assert_eq!(value(&wb, s, "C1"), num(100.0));
        let all = reported(&[
            ("A1", &["WEBSERVICE"]),
            ("B1", &["WEBSERVICE"]),
            ("C1", &["WEBSERVICE"]),
            ("E1", &["WEBSERVICE"]),
            ("F1", &["WEBSERVICE"]),
            ("B2", &["WEBSERVICE"]),
        ]);
        assert_eq!(unevaluated(&result), all);

        put_num(&mut wb, s, "D1", 2.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        assert_eq!(value(&wb, s, "B1"), num(99.0));
        assert_eq!(value(&wb, s, "F1"), CellValue::Empty);
        assert_eq!(value(&wb, s, "B2"), num(77.0));
        assert_eq!(value(&wb, s, "H1"), num(2.0));
        assert_eq!(unevaluated(&result), all);
    }

    /// A1 runs before the array B1 spills into B2, so its first evaluation
    /// takes the plain branch and writes 99 over excel's 77. once B2 holds 20
    /// the next round reaches the uncomputable call, and 77 comes back.
    #[test]
    fn a_later_round_reaching_a_call_puts_back_the_value_before_the_recalc() {
        let (mut wb, s) = one_sheet();
        put_cached_formula(
            &mut wb,
            s,
            "A1",
            "IF(B2=20,_xlfn.WEBSERVICE(\"http://x\"),99)",
            num(77.0),
        );
        put_formula(&mut wb, s, "B1", "_xlfn.SEQUENCE(2)*10");
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("B1"), xlsx_model::CellRange::parse_a1("B1").unwrap());
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B2"), num(20.0));
        assert_eq!(value(&wb, s, "A1"), num(77.0));
        assert_eq!(unevaluated(&result), reported(&[("A1", &["WEBSERVICE"])]));
    }

    /// the same race for a formula the file stored no value for: the round
    /// that reaches the call puts the blank back, and once B2 moves away from
    /// 20 the formula computes again, so nothing froze.
    #[test]
    fn a_later_round_reaching_a_call_puts_back_a_blank() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "D1", 1.0);
        put_formula(
            &mut wb,
            s,
            "A1",
            "IF(B2=20,IFERROR(_xlfn.WEBSERVICE(\"x\"),42),99)",
        );
        put_formula(&mut wb, s, "B1", "IF(D1=1,{10;20},{10;30})");
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("B1"), xlsx_model::CellRange::parse_a1("B1").unwrap());
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B2"), num(20.0));
        assert_eq!(value(&wb, s, "A1"), CellValue::Empty);
        assert_eq!(unevaluated(&result), reported(&[("A1", &["WEBSERVICE"])]));

        put_num(&mut wb, s, "D1", 2.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        assert_eq!(value(&wb, s, "B2"), num(30.0));
        assert_eq!(value(&wb, s, "A1"), num(99.0));
        assert!(result.unevaluated_cells.is_empty());
    }

    /// without the clock a dated formula counts every uncomputable call it
    /// holds, and the formulas reading it are found too.
    #[test]
    fn derivation_without_a_clock_counts_a_dated_formulas_calls() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "D1", 1.0);
        put_cached_formula(
            &mut wb,
            s,
            "A1",
            "IF(TODAY()>0,D1,_xlfn.WEBSERVICE(\"x\"))",
            num(5.0),
        );
        put_cached_formula(&mut wb, s, "B1", "A1+D1", num(6.0));
        let graph = DepGraph::build(&wb);
        derive_unevaluated(&wb, &graph, None, None);
        let derived: Vec<_> = graph
            .unevaluated_cells()
            .iter()
            .map(|cell| cell.cell.to_a1())
            .collect();
        assert_eq!(derived, ["A1", "B1"]);
        let graph = DepGraph::build(&wb);
        derive_unevaluated(&wb, &graph, Some(46000.0), None);
        assert!(graph.unevaluated_cells().is_empty());
    }

    fn derived(wb: &Workbook) -> Vec<(String, Vec<String>)> {
        let graph = DepGraph::build(wb);
        derive_unevaluated(wb, &graph, None, None);
        graph
            .unevaluated_cells()
            .iter()
            .map(|cell| (cell.cell.to_a1(), cell.functions.clone()))
            .collect()
    }

    /// derivation evaluates an array formula a kept cell taints for the
    /// rectangle its result would fill, as a recalculation does, so the
    /// formulas reading the cells it would newly fill are found at open too,
    /// over a large spill as well.
    #[test]
    fn derivation_follows_a_tainted_spill_to_the_readers_of_the_cells_it_would_fill() {
        let (mut inherited, s) = one_sheet();
        put_cached_formula(
            &mut inherited,
            s,
            "E1",
            "_xlfn.WEBSERVICE(\"x\")",
            text("x"),
        );
        put_recorded(&mut inherited, s, "A1", "IF(E1=\"x\",{10;20},0)", "A1");
        put_num(&mut inherited, s, "D1", 1.0);
        put_cached_formula(&mut inherited, s, "B1", "A2+D1", num(99.0));
        put_formula(&mut inherited, s, "F1", "A2*D1");
        put_cached_formula(&mut inherited, s, "B2", "A2*D1", num(77.0));
        put_cached_formula(&mut inherited, s, "C1", "B1+1", num(100.0));
        let ws = &["WEBSERVICE"][..];
        let expected = reported(&[
            ("A1", ws),
            ("B1", ws),
            ("C1", ws),
            ("E1", ws),
            ("F1", ws),
            ("B2", ws),
        ]);
        assert_eq!(derived(&inherited), expected);

        let (mut own, s) = one_sheet();
        let handled = "IFERROR(_xlfn.WEBSERVICE(\"x\"),{10;20})";
        put_recorded(&mut own, s, "A1", handled, "A1");
        put_cached_formula(&mut own, s, "B1", "A2+1", num(99.0));
        put_cached_formula(&mut own, s, "C1", "B1+1", num(100.0));
        let expected = reported(&[("A1", ws), ("B1", ws), ("C1", ws)]);
        assert_eq!(derived(&own), expected);

        let (mut large, s) = one_sheet();
        put_cached_formula(&mut large, s, "E1", "_xlfn.WEBSERVICE(\"x\")", text("x"));
        let spill = "IF(E1=\"x\",_xlfn.SEQUENCE(200000),0)";
        put_recorded(&mut large, s, "A1", spill, "A1");
        put_cached_formula(&mut large, s, "K1", "SUM(A150000:A200000)", num(7.0));
        let expected = reported(&[("A1", ws), ("E1", ws), ("K1", ws)]);
        assert_eq!(derived(&large), expected);

        for mut wb in [inherited, own, large] {
            let before = derived(&wb);
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(unevaluated(&result), before);
        }
    }

    /// without a clock a volatile formula holding an uncomputable call keeps
    /// what it holds, under a handler too, its readers keep theirs, and an
    /// uncached one stays blank; with a clock it computes.
    #[test]
    fn a_volatile_formula_without_a_clock_keeps_its_uncomputable_calls() {
        let workbook = || {
            let (mut wb, s) = one_sheet();
            put_num(&mut wb, s, "D1", 1.0);
            let dated = "IF(TODAY()>0,D1,_xlfn.WEBSERVICE(\"x\"))";
            put_cached_formula(&mut wb, s, "A1", dated, num(5.0));
            put_cached_formula(&mut wb, s, "B1", "A1+D1", num(99.0));
            put_formula(&mut wb, s, "A2", dated);
            put_formula(&mut wb, s, "B2", "A2+D1");
            let handled = "IFERROR(IF(TODAY()<0,\"old\",_xlfn.WEBSERVICE(\"x\")),\"fallback\")";
            put_cached_formula(&mut wb, s, "A3", handled, text("live"));
            put_cached_formula(&mut wb, s, "B3", "LEN(A3)*2", num(9.0));
            (wb, s)
        };
        let (mut wb, s) = workbook();
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        let kept = [
            ("A1", num(5.0)),
            ("B1", num(99.0)),
            ("A2", CellValue::Empty),
            ("B2", CellValue::Empty),
            ("A3", text("live")),
            ("B3", num(9.0)),
        ];
        let ws = &["WEBSERVICE"][..];
        let all = reported(&[
            ("A1", ws),
            ("B1", ws),
            ("A2", ws),
            ("B2", ws),
            ("A3", ws),
            ("B3", ws),
        ]);
        for (at, expected) in &kept {
            assert_eq!(value(&wb, s, at), *expected, "{at}");
        }
        assert_eq!(unevaluated(&result), all);
        put_num(&mut wb, s, "D1", 2.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("D1"))], None);
        for (at, expected) in &kept {
            assert_eq!(value(&wb, s, at), *expected, "{at}");
        }
        assert_eq!(unevaluated(&result), all);

        let (mut wb, s) = workbook();
        let (_, result) = rebuild_and_recalc_all(&mut wb, Some(46000.0));
        assert_eq!(value(&wb, s, "A1"), num(1.0));
        assert_eq!(value(&wb, s, "B1"), num(2.0));
        assert_eq!(value(&wb, s, "A2"), num(1.0));
        assert_eq!(value(&wb, s, "B2"), num(2.0));
        assert_eq!(value(&wb, s, "A3"), text("live"));
        assert_eq!(unevaluated(&result), reported(&[("A3", ws), ("B3", ws)]));
    }

    /// a formula holding an uncomputable call that the formula's own cap or
    /// the recalculation's spent budget stops before the call keeps what it
    /// holds, blank without a saved value, never `#NUM!`, and its readers keep
    /// theirs.
    #[test]
    fn a_budget_stop_before_an_uncomputable_call_keeps_the_formula() {
        let ws = &["WEBSERVICE"][..];
        let mut wb = over_the_formula_cap();
        let s = SheetId(1);
        let capped = "IF(SUM(Data!A1:XFD1048576)>0,_xlfn.WEBSERVICE(\"http://x\"),1)";
        put_cached_formula(&mut wb, s, "A1", capped, num(10.0));
        put_cached_formula(&mut wb, s, "B1", "A1+1", num(99.0));
        put_formula(&mut wb, s, "A2", capped);
        put_formula(&mut wb, s, "B2", "A2+1");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "A1"), num(10.0));
        assert_eq!(value(&wb, s, "B1"), num(99.0));
        assert_eq!(value(&wb, s, "A2"), CellValue::Empty);
        assert_eq!(value(&wb, s, "B2"), CellValue::Empty);
        assert!(result.limited_cells.contains(&(s, a1("A1"))));
        assert!(result.limited_cells.contains(&(s, a1("A2"))));
        let all = reported(&[("A1", ws), ("B1", ws), ("A2", ws), ("B2", ws)]);
        assert_eq!(unevaluated(&result), all);

        let (mut wb, s) = one_sheet();
        spend_the_budget(&mut wb, s);
        for row in 1..=3 {
            put_num(&mut wb, s, &format!("E{row}"), f64::from(row));
        }
        let spent = "IF(SUM(E1:E3)>0,_xlfn.WEBSERVICE(\"x\"),1)";
        put_cached_formula(&mut wb, s, "B20", spent, num(5.0));
        put_formula(&mut wb, s, "B21", spent);
        put_cached_formula(&mut wb, s, "C20", "B20+1", num(99.0));
        put_formula(&mut wb, s, "C21", "B21+1");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.budget_exhausted);
        assert_eq!(value(&wb, s, "B20"), num(5.0));
        assert_eq!(value(&wb, s, "C20"), num(99.0));
        assert_eq!(value(&wb, s, "B21"), CellValue::Empty);
        assert_eq!(value(&wb, s, "C21"), CellValue::Empty);
        let all = reported(&[("B20", ws), ("C20", ws), ("B21", ws), ("C21", ws)]);
        assert_eq!(unevaluated(&result), all);
    }

    /// A1 runs before the array B1 spills into B2 and writes its plain
    /// branch. once B2 holds 20 the next round reaches the call: it measures
    /// its reach over the rectangle recorded before the recalc, so D1 reading
    /// A2 is not taken for a reader of its result, and the rectangle a round
    /// reshaped without moving a value comes back too.
    #[test]
    fn a_later_round_keeps_the_rectangle_recorded_before_the_recalc() {
        let (mut wb, s) = one_sheet();
        let grows = "IF(B2=20,_xlfn.WEBSERVICE(\"http://x\"),{100;200})";
        put_recorded(&mut wb, s, "A1", grows, "A1");
        set_cached(&mut wb, s, "A1", num(77.0));
        put_formula(&mut wb, s, "D1", "A2+1");
        put_recorded(&mut wb, s, "B1", "_xlfn.SEQUENCE(2)*10", "B1");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B2"), num(20.0));
        assert_eq!(value(&wb, s, "A1"), num(77.0));
        assert_eq!(value(&wb, s, "A2"), CellValue::Empty);
        assert_eq!(value(&wb, s, "D1"), num(1.0));
        assert_eq!(
            wb.sheet(s).unwrap().array_formula(a1("A1")),
            Some(CellRange::parse_a1("A1").unwrap())
        );
        assert_eq!(unevaluated(&result), reported(&[("A1", &["WEBSERVICE"])]));

        let (mut wb, s) = one_sheet();
        let shrinks = "IF(B2=20,_xlfn.WEBSERVICE(\"http://x\"),{77;88})";
        put_recorded(&mut wb, s, "A1", shrinks, "A1:A3");
        set_cached(&mut wb, s, "A1", num(77.0));
        put_num(&mut wb, s, "A2", 88.0);
        put_recorded(&mut wb, s, "B1", "_xlfn.SEQUENCE(2)*10", "B1");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B2"), num(20.0));
        assert_eq!(value(&wb, s, "A1"), num(77.0));
        assert_eq!(value(&wb, s, "A2"), num(88.0));
        assert_eq!(
            wb.sheet(s).unwrap().array_formula(a1("A1")),
            Some(CellRange::parse_a1("A1:A3").unwrap())
        );
        assert_eq!(unevaluated(&result), reported(&[("A1", &["WEBSERVICE"])]));
    }

    fn sorted(mut cells: Vec<(String, Vec<String>)>) -> Vec<(String, Vec<String>)> {
        cells.sort();
        cells
    }

    /// B1, C1, E1, F1 and G1 read the uncomputable A6 through a reference
    /// they compute, so each keeps its cache and is listed, at open, through
    /// a full recalculation and after an edit. B1 to G1 come before A6, so a
    /// recalculation reads A6 before it settles; B10 comes after. H1 reads a
    /// plain cell the same way and computes.
    fn computed_reference_readers() -> (Workbook, SheetId) {
        let (mut wb, s) = one_sheet();
        put_cached_formula(&mut wb, s, "A6", "_xlfn.WEBSERVICE(\"http://x\")", num(5.0));
        put_num(&mut wb, s, "A7", 1.0);
        put_num(&mut wb, s, "A8", 2.0);
        put_num(&mut wb, s, "D1", 0.0);
        put_cached_formula(&mut wb, s, "B1", "INDIRECT(\"A6\")*2", num(99.0));
        put_cached_formula(&mut wb, s, "C1", "OFFSET(A6,D1,0)+0", num(98.0));
        put_cached_formula(&mut wb, s, "E1", "INDEX(A6:A8,D1+1)+0", num(97.0));
        put_cached_formula(&mut wb, s, "F1", "SUM(INDIRECT(\"A6:A8\"))", num(96.0));
        put_cached_formula(&mut wb, s, "G1", "CHOOSE(D1+1,A6,A7)+0", num(95.0));
        put_cached_formula(&mut wb, s, "B10", "INDIRECT(\"A6\")+1", num(94.0));
        put_cached_formula(&mut wb, s, "H1", "INDIRECT(\"A7\")*2", num(50.0));
        (wb, s)
    }

    #[test]
    fn readers_through_computed_references_keep_their_caches() {
        let listed = sorted(reported(&[
            ("A6", &["WEBSERVICE"]),
            ("B1", &["WEBSERVICE"]),
            ("B10", &["WEBSERVICE"]),
            ("C1", &["WEBSERVICE"]),
            ("E1", &["WEBSERVICE"]),
            ("F1", &["WEBSERVICE"]),
            ("G1", &["WEBSERVICE"]),
        ]));
        let kept = |wb: &Workbook, s: SheetId| {
            for (at, cached) in [
                ("A6", 5.0),
                ("B1", 99.0),
                ("C1", 98.0),
                ("E1", 97.0),
                ("F1", 96.0),
                ("G1", 95.0),
                ("B10", 94.0),
            ] {
                assert_eq!(value(wb, s, at), num(cached), "{at}");
            }
        };
        let (wb, _) = computed_reference_readers();
        assert_eq!(sorted(derived(&wb)), listed);

        let (mut wb, s) = computed_reference_readers();
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, Some(46_000.0));
        kept(&wb, s);
        assert_eq!(value(&wb, s, "H1"), num(2.0));
        assert_eq!(sorted(unevaluated(&result)), listed);

        put_num(&mut wb, s, "A7", 3.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("A7"))], Some(46_000.0));
        kept(&wb, s);
        assert_eq!(value(&wb, s, "H1"), num(6.0));
        assert_eq!(sorted(unevaluated(&result)), listed);
    }

    /// a cycle through a fallback, a cycle only one of whose members holds the
    /// call, and one whose members hold no value: what each would settle at
    /// hangs on what the service returns, so every member keeps what it holds
    /// and is listed, and none settles as circular. a plain cycle still does.
    fn uncomputable_cycles() -> (Workbook, SheetId) {
        let (mut wb, s) = one_sheet();
        put_cached_formula(
            &mut wb,
            s,
            "A1",
            "IFERROR(_xlfn.WEBSERVICE(\"http://x\"),B1)",
            text("live"),
        );
        put_cached_formula(&mut wb, s, "B1", "A1", text("live"));
        put_cached_formula(
            &mut wb,
            s,
            "A3",
            "IF(B3>=0,_xlfn.WEBSERVICE(\"http://x\"),1)",
            num(7.0),
        );
        put_cached_formula(&mut wb, s, "B3", "A3+1", num(8.0));
        put_formula(
            &mut wb,
            s,
            "A5",
            "IFERROR(_xlfn.WEBSERVICE(\"http://x\"),B5)",
        );
        put_formula(&mut wb, s, "B5", "A5");
        put_cached_formula(&mut wb, s, "C5", "B5&\"!\"", text("kept"));
        put_cached_formula(&mut wb, s, "D1", "E1", num(4.0));
        put_cached_formula(&mut wb, s, "E1", "D1", num(4.0));
        (wb, s)
    }

    #[test]
    fn a_cycle_an_uncomputable_call_reaches_keeps_every_member() {
        let listed = sorted(reported(&[
            ("A1", &["WEBSERVICE"]),
            ("A3", &["WEBSERVICE"]),
            ("A5", &["WEBSERVICE"]),
            ("B1", &["WEBSERVICE"]),
            ("B3", &["WEBSERVICE"]),
            ("B5", &["WEBSERVICE"]),
            ("C5", &["WEBSERVICE"]),
        ]));
        let (wb, _) = uncomputable_cycles();
        assert_eq!(sorted(derived(&wb)), listed);

        let (mut wb, s) = uncomputable_cycles();
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, Some(46_000.0));
        let kept = |wb: &Workbook| {
            assert_eq!(value(wb, s, "A1"), text("live"));
            assert_eq!(value(wb, s, "B1"), text("live"));
            assert_eq!(value(wb, s, "A3"), num(7.0));
            assert_eq!(value(wb, s, "B3"), num(8.0));
            assert_eq!(value(wb, s, "A5"), CellValue::Empty);
            assert_eq!(value(wb, s, "B5"), CellValue::Empty);
            assert_eq!(value(wb, s, "C5"), text("kept"));
            assert_eq!(value(wb, s, "D1"), num(0.0));
            assert_eq!(value(wb, s, "E1"), num(0.0));
        };
        kept(&wb);
        assert_eq!(sorted(unevaluated(&result)), listed);
        let circular: Vec<String> = result.cycle_cells.iter().map(|(_, c)| c.to_a1()).collect();
        assert_eq!(circular, ["D1", "E1"]);

        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("B3"))], Some(46_000.0));
        kept(&wb);
        assert_eq!(sorted(unevaluated(&result)), listed);
    }

    /// stepping a derivation one step at a time, so a host can bound each
    /// slice, marks exactly what one run to the end does: over chains,
    /// spills, cycles and readers through computed references.
    #[test]
    fn a_derivation_in_single_steps_marks_what_a_one_shot_one_does() {
        let (mut wb, s) = computed_reference_readers();
        let (cycles, _) = uncomputable_cycles();
        let mut cycles = cycles.sheets.into_iter().next().unwrap();
        cycles.name = "Cycles".to_string();
        wb.sheets.push(cycles);
        put_cached_formula(&mut wb, s, "J1", "_xlfn.WEBSERVICE(\"x\")", text("x"));
        put_formula(&mut wb, s, "K1", "IF(J1=\"x\",{10;20},0)");
        wb.sheet_mut(s)
            .unwrap()
            .set_array_formula(a1("K1"), xlsx_model::CellRange::parse_a1("K1").unwrap());
        put_cached_formula(&mut wb, s, "L1", "K2+D1", num(9.0));
        for row in 20u32..60 {
            put_cached_formula(
                &mut wb,
                s,
                &format!("M{row}"),
                &format!("M{}+1", row - 1),
                num(1.0),
            );
            put_num(&mut wb, s, &format!("N{row}"), f64::from(row));
        }
        put_cached_formula(&mut wb, s, "M19", "LEN(B10)", num(1.0));

        let whole = DepGraph::build(&wb);
        derive_unevaluated(&wb, &whole, None, None);
        let stepped = DepGraph::build(&wb);
        let mut derivation = Derivation::new(None, None);
        let mut steps = 0;
        loop {
            let step = derivation.advance(&wb, &stepped, 1);
            assert_eq!(step.work, 1);
            steps += 1;
            if step.done {
                break;
            }
        }
        assert!(steps > 100, "{steps}");
        let marked = whole.unevaluated_cells();
        assert!(marked.len() > 50, "{}", marked.len());
        assert_eq!(stepped.unevaluated_cells(), marked);
    }

    #[test]
    fn a_computable_workbook_reports_no_unevaluated_cells() {
        let (mut wb, s) = one_sheet();
        put_num(&mut wb, s, "A1", 2.0);
        put_cached_formula(&mut wb, s, "B1", "IFERROR(A1/0,\"x\")", num(1.0));
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, s, "B1"), text("x"));
        assert!(result.unevaluated_cells.is_empty());
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

    #[test]
    fn names_edited_after_the_graph_was_built_still_evaluate_from_the_workbook() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        for index in 0..3 {
            wb.defined_names.push(DefinedName {
                name: format!("Rate_{index}"),
                formula: format!("{index}"),
                local_sheet: None,
                hidden: index == 1,
            });
        }
        put_formula(&mut wb, SheetId(0), "A1", "Rate_1*2");
        let (mut graph, _) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(value(&wb, SheetId(0), "A1"), num(2.0));

        wb.defined_names[1].formula = "5".into();
        recalc_after(&mut wb, &mut graph, &[(SheetId(0), a1("A1"))], None);
        assert_eq!(value(&wb, SheetId(0), "A1"), num(10.0));
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
            let expected = num(crate::fmath::powi(2.0, (N - 2 - at) as i32));
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

    /// the result says when the recalculation's own budget ran out, which a
    /// single formula over its cap does not.
    #[test]
    fn a_spent_recalculation_reports_its_budget_exhausted() {
        let (mut wb, s) = one_sheet();
        put_formula(&mut wb, s, "B1", "1+2");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(!result.budget_exhausted);

        let mut wb = over_the_formula_cap();
        put_formula(&mut wb, SheetId(1), "A1", "SUM(Data!A1:XFD1048576)");
        let (_, result) = rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(result.limited_cells, vec![(SheetId(1), a1("A1"))]);
        assert!(!result.budget_exhausted);

        let (mut wb, s) = one_sheet();
        spend_the_budget(&mut wb, s);
        put_cached_formula(&mut wb, s, "B1", "SUM(A1:A11)", num(5.0));
        let (mut graph, result) = rebuild_and_recalc_all(&mut wb, None);
        assert!(result.budget_exhausted);
        assert_eq!(value(&wb, s, "B1"), num(5.0));

        put_num(&mut wb, s, "C1", 1.0);
        let result = recalc_after(&mut wb, &mut graph, &[(s, a1("C1"))], None);
        assert!(!result.budget_exhausted);
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
        let missing =
            "IFERROR(IFERROR(_xlfn.XLOOKUP(1,NA(),OFFSET(D1,0,0,SUM(A:A))),42)+SUM(A:A),0)";
        for formula in [lookup, missing] {
            let (mut wb, s) = inputs(false);
            put_formula(&mut wb, s, "E1", formula);
            let (_, result) = rebuild_and_recalc_all(&mut wb, None);
            assert_eq!(value(&wb, s, "E1"), num(43.0), "{formula}");
            assert!(result.limited_cells.is_empty(), "{formula}");
        }
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
