//! Conditional-format effects of the cells a display list draws.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::sync::Arc;

use xlsx_calc::eval::MAX_EVALUATION_CELL_VISITS;
use xlsx_calc::functions::criteria::CountTable;
use xlsx_calc::{BinaryOp, ColumnRange, EvalContext, Expr, RowRange, evaluate, parse_formula};
use xlsx_model::styles::DifferentialFormat;
use xlsx_model::value::CellValue;
use xlsx_model::{CellRange, CellRef, ErrorValue, MAX_COLS, MAX_ROWS, SheetId, Workbook};
use xlsx_render::CellFormats;

/// Rule checks plus cell visits one frame on screen spends deciding cells.
pub const CONDITIONAL_FRAME_BUDGET: u64 = 200_000;
/// Milliseconds one timed frame on screen spends deciding cells.
pub const CONDITIONAL_FRAME_MS: f64 = 10.0;
/// A timed frame's unit cap in untimed frames, which still ends a frame whose
/// clock does not advance.
const TIMED_FRAMES: u64 = 10;
/// Cells decided between two reads of a frame's clock.
const CLOCK_STRIDE: u64 = 8;
/// Units spent since the last read that make the next cell read it anyway.
const CLOCK_UNITS: u64 = 4_096;
/// Decided cells kept per model epoch; past it, cells outside the frame go.
pub const CONDITIONAL_RESULT_CAP: usize = 262_144;
/// Parsed shifted formulas kept across epochs; cleared when full.
const PARSED_CAP: usize = 16_384;

/// Functions whose call over absolute references reads the same for every cell.
const AGGREGATES: &[&str] = &[
    "AVERAGE",
    "AVERAGEA",
    "AVERAGEIF",
    "AVERAGEIFS",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "COUNTIF",
    "COUNTIFS",
    "LARGE",
    "MAX",
    "MAXA",
    "MAXIFS",
    "MEDIAN",
    "MIN",
    "MINA",
    "MINIFS",
    "PRODUCT",
    "SMALL",
    "STDEV",
    "STDEV.P",
    "STDEV.S",
    "STDEVA",
    "STDEVP",
    "SUM",
    "SUMIF",
    "SUMIFS",
    "VAR",
    "VAR.P",
    "VAR.S",
];

/// Volatile functions reading the clock, decided with the date the cells were
/// last calculated with when there is one.
const CLOCK: &[&str] = &["TODAY", "NOW"];
/// The calc engine's other volatile functions less INDIRECT, whose value
/// follows the cells it reads: no date decides them.
const RANDOM: &[&str] = &["RAND", "RANDBETWEEN", "RANDARRAY"];
/// Defined names one volatility scan expands, the calc engine's name depth.
const NAME_DEPTH: usize = 256;

/// A monotonic clock in milliseconds that bounds a frame's conditional-format
/// work; the host supplies it, since the engine reads no clock of its own.
pub trait FrameClock {
    fn now_ms(&self) -> f64;
}

impl<F: Fn() -> f64> FrameClock for F {
    fn now_ms(&self) -> f64 {
        self()
    }
}

/// How much evaluation a display list spends on conditional formats.
#[derive(Clone, Copy)]
pub(crate) enum Budget<'c> {
    /// One frame's budget, ended early once the clock passes the frame's
    /// time; later frames decide the rest. A cell it leaves undecided paints
    /// its result from before the last direct edit of one cell or
    /// recalculation when it has one, and no conditional format otherwise.
    Frame(Option<&'c dyn FrameClock>),
    /// Every visible cell, for output no later frame replaces.
    Complete,
}

/// Counters of the conditional-format cache.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ConditionalStats {
    /// Visible cells read back decided.
    pub hits: u64,
    /// Visible cells whose rules were evaluated.
    pub evaluated_cells: u64,
    /// Whole-range statistics and constant calls computed.
    pub aggregate_builds: u64,
    /// Visible cells the last display list left undecided.
    pub pending: u32,
    /// Rule formula texts parsed into expressions.
    pub formula_parses: u64,
    /// Rules built whose cells parse their translated formula text instead of
    /// shifting the anchor's parse.
    pub slow_rules: u64,
    /// Rule checks, cell visits, formulas compiled and rule range tests the
    /// last display list spent.
    pub units: u64,
    /// Decided cells held.
    pub cached_cells: usize,
    /// Results of the epoch before held for cells not yet decided again.
    pub previous_cells: usize,
    /// Model epoch the cache holds.
    pub epoch: u64,
    /// Cells withheld because a between rule's bound is an error.
    pub withheld_error_bound: u64,
    /// Cells withheld because a rule formula did not translate or parse there.
    pub withheld_unparsable: u64,
    /// Cells withheld because the calc engine refused a rule formula for size.
    pub withheld_refused: u64,
    /// Cells withheld because a rule's kind, operator or rank is unsupported.
    pub withheld_unsupported: u64,
    /// Cells withheld because a rule reads random numbers, or the clock in a
    /// workbook not calculated with a date.
    pub withheld_volatile: u64,
    /// Cells withheld because a range rule's areas overlap.
    pub withheld_overlapping: u64,
    /// Cells whose conditional fill is withheld because its format names two
    /// different fill colors.
    pub withheld_fill: u64,
}

impl ConditionalStats {
    fn withhold(&mut self, reason: Unknown) {
        let counter = match reason {
            Unknown::ErrorBound => &mut self.withheld_error_bound,
            Unknown::Unparsable => &mut self.withheld_unparsable,
            Unknown::Refused => &mut self.withheld_refused,
            Unknown::Unsupported => &mut self.withheld_unsupported,
            Unknown::Volatile => &mut self.withheld_volatile,
            Unknown::Overlapping => &mut self.withheld_overlapping,
        };
        *counter += 1;
    }
}

/// Why a rule's result at a cell cannot be decided; a cell whose rules reach
/// one paints no conditional format.
#[derive(Clone, Copy)]
enum Unknown {
    ErrorBound,
    Unparsable,
    Refused,
    Unsupported,
    Volatile,
    Overlapping,
}

/// Each sheet's rules, aggregates and decided cells for one model epoch, the
/// decided cells of the epoch before, and shifted formulas parsed in any.
pub(crate) struct Cache {
    /// The model epoch and the bits of the calculation date the cache holds.
    epoch: Option<(u64, Option<u64>)>,
    /// The calculation date rules on the clock are decided with.
    now_serial: Option<f64>,
    rules: HashMap<SheetId, Arc<[Rule]>>,
    aggregates: HashMap<(SheetId, usize), Arc<Aggregate>>,
    results: HashMap<(SheetId, u32, u32), Option<Arc<DifferentialFormat>>>,
    /// The epoch before's results of the cells not yet decided in this one.
    previous: HashMap<(SheetId, u32, u32), Option<Arc<DifferentialFormat>>>,
    /// Whether a change since the epoch was other than a direct edit of one
    /// cell or a recalculation, so that no result of the epoch is painted in
    /// the next.
    moved: bool,
    parsed: HashMap<String, Option<Arc<Expr>>>,
    budget: u64,
    stats: ConditionalStats,
}

impl Default for Cache {
    fn default() -> Self {
        Self {
            epoch: None,
            now_serial: None,
            rules: HashMap::new(),
            aggregates: HashMap::new(),
            results: HashMap::new(),
            previous: HashMap::new(),
            moved: false,
            parsed: HashMap::new(),
            budget: CONDITIONAL_FRAME_BUDGET,
            stats: ConditionalStats::default(),
        }
    }
}

pub(crate) struct Rule {
    ranges: Vec<CellRange>,
    anchor: CellRef,
    test: Test,
    formulas: Vec<Formula>,
    format: Arc<DifferentialFormat>,
    fill_unknown: bool,
    stop: bool,
}

enum Test {
    CellIs(Op),
    Expression,
    Top {
        rank: u32,
        percent: bool,
        bottom: bool,
    },
    Average {
        above: bool,
        equal: bool,
        deviations: f64,
    },
    Duplicates {
        unique: bool,
    },
    Undecidable(Unknown),
}

/// A cell's conditional format, or a whole-range build the frame has no
/// room left for, which waits for a frame it starts.
enum Decision {
    Paint(Option<Arc<DifferentialFormat>>),
    Later,
}

#[derive(Clone, Copy)]
enum Op {
    Compare(BinaryOp),
    Between(bool),
}

/// A rule formula as written at the anchor.
struct Formula {
    source: String,
    base: Option<Arc<Expr>>,
    /// Whether a cell other than the anchor reads it shifted.
    relative: bool,
    /// Whether moving the references of `base` gives what parsing the
    /// translated text gives, so a cell shifts `base` instead of parsing.
    shifts: bool,
    /// Paths in `base` of the calls whose value is the same for every cell.
    constants: Vec<Vec<usize>>,
    /// Paths in `base` of `COUNTIF` calls over a fixed range whose criterion
    /// is a reference that moves with the cell.
    counts: Vec<Vec<usize>>,
    /// Texts parsed to compile it.
    parses: u8,
}

/// What one rule reads from its whole range, once per epoch.
#[derive(Default)]
struct Aggregate {
    stats: Option<Stats>,
    /// Each formula's constant calls, in `Formula::constants` order.
    constants: Vec<Vec<CellValue>>,
    /// Each formula's `base` with its constant calls read in.
    bases: Vec<Option<Arc<Expr>>>,
    /// Each formula's count tables, in `Formula::counts` order.
    tables: Vec<Vec<Option<CountTable>>>,
    /// Whether the calc engine refused a constant call for size.
    refused: bool,
}

enum Stats {
    Threshold(Option<f64>),
    Average { mean: f64, deviation: f64 },
    Counts(HashMap<Key, u32>),
}

#[derive(PartialEq, Eq, Hash)]
enum Key {
    Number(u64),
    Text(String),
    Bool(bool),
}

impl Cache {
    pub(crate) fn set_budget(&mut self, visits: u64) {
        self.budget = visits.max(1);
    }

    pub(crate) fn stats(&self) -> ConditionalStats {
        ConditionalStats {
            cached_cells: self.results.len(),
            previous_cells: self.previous.len(),
            epoch: self.epoch.map_or(0, |(epoch, _)| epoch),
            ..self.stats
        }
    }

    /// Marks the change that ends the epoch as one whose earlier results are
    /// not painted.
    pub(crate) fn cells_moved(&mut self) {
        self.moved = true;
    }

    /// The effects on the cells of `sheet` in the `rows` by `cols` runs, and
    /// how many of them `budget` left undecided. A cell takes each property
    /// from the highest-priority matching rule that sets it, and no rule
    /// below a matching `stopIfTrue` rule applies. `epoch` is the model epoch
    /// and the date its cells were last calculated with; a change of either
    /// starts a new epoch of the cache.
    pub(crate) fn effects(
        &mut self,
        epoch: (u64, Option<f64>),
        model: &Workbook,
        sheet: SheetId,
        rows: &[Range<u32>],
        cols: &[Range<u32>],
        budget: Budget<'_>,
    ) -> (CellFormats, u32) {
        let (epoch, now_serial) = epoch;
        let key = (epoch, now_serial.map(f64::to_bits));
        if self.epoch != Some(key) {
            self.epoch = Some(key);
            self.now_serial = now_serial;
            self.rules.clear();
            self.aggregates.clear();
            if std::mem::take(&mut self.moved) {
                self.previous.clear();
                self.results.clear();
            } else {
                self.previous = std::mem::take(&mut self.results);
            }
        }
        let (limit, clock) = match budget {
            Budget::Frame(None) => (self.budget, None),
            Budget::Frame(Some(clock)) => (self.budget.saturating_mul(TIMED_FRAMES), Some(clock)),
            Budget::Complete => (u64::MAX, None),
        };
        let ruled = match self.rules.get(&sheet) {
            Some(rules) => !rules.is_empty(),
            None => model
                .sheet(sheet)
                .is_some_and(|sheet| !sheet.conditional_formats.is_empty()),
        };
        let deadline = clock
            .filter(|_| ruled)
            .map(|clock| clock.now_ms() + CONDITIONAL_FRAME_MS);
        let mut scanned = 0u64;
        if !self.rules.contains_key(&sheet) {
            let built: Arc<[Rule]> = rules(model, sheet, self.now_serial.is_some()).into();
            scanned += built
                .iter()
                .map(|rule| rule.formulas.len() as u64)
                .sum::<u64>();
            self.stats.formula_parses += built
                .iter()
                .flat_map(|rule| &rule.formulas)
                .map(|formula| u64::from(formula.parses))
                .sum::<u64>();
            self.stats.slow_rules += built
                .iter()
                .filter(|rule| {
                    rule.formulas
                        .iter()
                        .any(|formula| formula.relative && !formula.shifts)
                })
                .count() as u64;
            self.rules.insert(sheet, built);
        }
        let rules = Arc::clone(&self.rules[&sheet]);
        let mut formats = CellFormats::new();
        self.stats.pending = 0;
        self.stats.units = scanned;
        if rules.is_empty() {
            return (formats, 0);
        }
        let range_tests = rules
            .iter()
            .map(|rule| rule.ranges.len() as u64)
            .sum::<u64>();
        let mut decided = 0u64;
        let mut checked = 0u64;
        let mut stopped = false;
        let mut context = match self.now_serial {
            Some(now_serial) => EvalContext::with_now(model, sheet, now_serial),
            None => EvalContext::new(model, sheet),
        };
        let mut spent = 0;
        let mut pending = 0u32;
        let mut caching = true;
        for row in rows.iter().cloned().flatten() {
            if !stopped {
                scanned += range_tests;
            }
            let candidates: Vec<usize> = (0..rules.len())
                .filter(|&index| {
                    rules[index]
                        .ranges
                        .iter()
                        .any(|range| (range.start.row..=range.end.row).contains(&row))
                })
                .collect();
            if candidates.is_empty() {
                continue;
            }
            for col in cols.iter().cloned().flatten() {
                let at = CellRef::new(row, col);
                if !stopped {
                    scanned += candidates.len() as u64;
                }
                let Some(first) = candidates
                    .iter()
                    .map(|&index| &rules[index])
                    .find(|rule| rule.covers(at))
                else {
                    continue;
                };
                let key = (sheet, row, col);
                let cached = self.results.get(&key).cloned();
                if !stopped && decided > 0 {
                    if spent >= limit {
                        stopped = true;
                    } else if let (Some(clock), Some(deadline)) = (clock, deadline)
                        && (spent + scanned >= checked + CLOCK_UNITS
                            || (cached.is_none() && decided.is_multiple_of(CLOCK_STRIDE)))
                    {
                        checked = spent + scanned;
                        stopped = clock.now_ms() >= deadline;
                    }
                }
                let result = match (cached, first.undecidable()) {
                    (Some(result), _) => {
                        self.stats.hits += 1;
                        result
                    }
                    (None, Some(reason)) => {
                        self.stats.withhold(reason);
                        if caching {
                            caching = self.keep(sheet, at, None, rows, cols);
                        }
                        None
                    }
                    (None, None) if stopped => {
                        pending += 1;
                        match self.previous.get(&key) {
                            Some(result) => result.clone(),
                            None => continue,
                        }
                    }
                    (None, None) => {
                        let decision = self.decide(
                            model,
                            sheet,
                            &rules,
                            &candidates,
                            at,
                            &mut context,
                            &mut spent,
                            (decided > 0).then_some(limit),
                        );
                        match decision {
                            Decision::Paint(result) => {
                                decided += 1;
                                self.stats.evaluated_cells += 1;
                                if caching {
                                    caching = self.keep(sheet, at, result.clone(), rows, cols);
                                }
                                result
                            }
                            Decision::Later => {
                                stopped = true;
                                pending += 1;
                                match self.previous.get(&key) {
                                    Some(result) => result.clone(),
                                    None => continue,
                                }
                            }
                        }
                    }
                };
                if let Some(format) = result {
                    formats.insert((row, col), format);
                }
            }
        }
        self.stats.pending = pending;
        self.stats.units = spent + scanned;
        (formats, pending)
    }

    /// Holds a decided cell in place of its result from the epoch before,
    /// first dropping the cells outside the frame and every whole-range build
    /// when full; false when the frame alone fills the cap.
    fn keep(
        &mut self,
        sheet: SheetId,
        at: CellRef,
        result: Option<Arc<DifferentialFormat>>,
        rows: &[Range<u32>],
        cols: &[Range<u32>],
    ) -> bool {
        if self.results.len() >= CONDITIONAL_RESULT_CAP {
            self.aggregates.clear();
            self.results.retain(|&(owner, row, col), _| {
                owner == sheet && within(rows, row) && within(cols, col)
            });
            if self.results.len() >= CONDITIONAL_RESULT_CAP {
                return false;
            }
        }
        let key = (sheet, at.row, at.col);
        self.previous.remove(&key);
        self.results.insert(key, result);
        true
    }

    #[allow(clippy::too_many_arguments)]
    fn decide(
        &mut self,
        model: &Workbook,
        sheet: SheetId,
        rules: &[Rule],
        candidates: &[usize],
        at: CellRef,
        context: &mut EvalContext<'_>,
        spent: &mut u64,
        limit: Option<u64>,
    ) -> Decision {
        let value = model
            .sheet(sheet)
            .and_then(|owner| owner.cell(at))
            .map_or(CellValue::Empty, |cell| cell.value.clone());
        let mut merged: Option<Arc<DifferentialFormat>> = None;
        let mut fill_unknown = false;
        for &index in candidates {
            let rule = &rules[index];
            if !rule.covers(at) {
                continue;
            }
            if let Some(limit) = limit
                && rule.undecidable().is_none()
                && !self.aggregates.contains_key(&(sheet, index))
                && build_cost(model, sheet, rule) > limit.saturating_sub(*spent)
            {
                return Decision::Later;
            }
            *spent += 1;
            match self.matches(model, sheet, index, rule, &value, at, context, spent) {
                Ok(true) => {}
                Ok(false) => continue,
                Err(reason) => {
                    self.stats.withhold(reason);
                    return Decision::Paint(None);
                }
            }
            merged = Some(match merged {
                None => Arc::clone(&rule.format),
                Some(format) => Arc::new(DifferentialFormat {
                    font_color: format
                        .font_color
                        .clone()
                        .or_else(|| rule.format.font_color.clone()),
                    bold: format.bold.or(rule.format.bold),
                    italic: format.italic.or(rule.format.italic),
                    fill: format
                        .fill
                        .clone()
                        .or_else(|| rule.format.fill.clone().filter(|_| !fill_unknown)),
                }),
            });
            fill_unknown |=
                rule.fill_unknown && merged.as_ref().is_some_and(|format| format.fill.is_none());
            if rule.stop {
                break;
            }
        }
        if fill_unknown {
            self.stats.withheld_fill += 1;
        }
        Decision::Paint(merged)
    }

    #[allow(clippy::too_many_arguments)]
    fn matches(
        &mut self,
        model: &Workbook,
        sheet: SheetId,
        index: usize,
        rule: &Rule,
        value: &CellValue,
        at: CellRef,
        context: &mut EvalContext<'_>,
        spent: &mut u64,
    ) -> Result<bool, Unknown> {
        if let Some(reason) = rule.undecidable() {
            return Err(reason);
        }
        let aggregate = self.aggregate(model, sheet, index, rule, context, spent);
        if aggregate.refused {
            return Err(Unknown::Refused);
        }
        context.restart();
        context.cell = Some(at);
        let ctx: &EvalContext<'_> = context;
        let matched = match &rule.test {
            Test::CellIs(op) => {
                if is_error(value) {
                    Ok(false)
                } else {
                    match *op {
                        Op::Compare(op) => match self.formula(rule, 0, &aggregate, at, ctx) {
                            Some(rhs) => Ok(compare(ctx, op, value, &rhs)),
                            None => Err(Unknown::Unparsable),
                        },
                        Op::Between(inside) => {
                            let low = self.formula(rule, 0, &aggregate, at, ctx);
                            let high = self.formula(rule, 1, &aggregate, at, ctx);
                            match (low, high) {
                                (Some(low), Some(high)) if is_error(&low) || is_error(&high) => {
                                    Err(Unknown::ErrorBound)
                                }
                                (Some(low), Some(high)) => {
                                    let within = |low: &CellValue, high: &CellValue| {
                                        compare(ctx, BinaryOp::Ge, value, low)
                                            && compare(ctx, BinaryOp::Le, value, high)
                                    };
                                    Ok((within(&low, &high) || within(&high, &low)) == inside)
                                }
                                _ => Err(Unknown::Unparsable),
                            }
                        }
                    }
                }
            }
            Test::Expression => match self.formula(rule, 0, &aggregate, at, ctx) {
                Some(CellValue::Bool { value }) => Ok(value),
                Some(CellValue::Number { value }) => Ok(value != 0.0),
                Some(_) => Ok(false),
                None => Err(Unknown::Unparsable),
            },
            Test::Top { bottom, .. } => Ok(match (value, &aggregate.stats) {
                (CellValue::Number { value }, Some(Stats::Threshold(Some(threshold)))) => {
                    if *bottom {
                        value <= threshold
                    } else {
                        value >= threshold
                    }
                }
                _ => false,
            }),
            Test::Average {
                above,
                equal,
                deviations,
            } => Ok(match (value, &aggregate.stats) {
                (CellValue::Number { value }, Some(Stats::Average { mean, deviation })) => {
                    let bound = if *above {
                        mean + deviations * deviation
                    } else {
                        mean - deviations * deviation
                    };
                    match (above, equal) {
                        (true, true) => *value >= bound,
                        (true, false) => *value > bound,
                        (false, true) => *value <= bound,
                        (false, false) => *value < bound,
                    }
                }
                _ => false,
            }),
            Test::Duplicates { unique } => Ok(match (key(value), &aggregate.stats) {
                (Some(key), Some(Stats::Counts(counts))) => {
                    let count = counts.get(&key).copied().unwrap_or(0);
                    if *unique { count == 1 } else { count > 1 }
                }
                _ => false,
            }),
            Test::Undecidable(reason) => Err(*reason),
        };
        *spent += MAX_EVALUATION_CELL_VISITS.saturating_sub(ctx.remaining_cells());
        if ctx.exhausted() {
            return Err(Unknown::Refused);
        }
        matched
    }

    /// The value of a rule's formula at `at`: its parse at the anchor with the
    /// references moved to `at`, its constant calls and count tables read from
    /// `aggregate`.
    fn formula(
        &mut self,
        rule: &Rule,
        slot: usize,
        aggregate: &Aggregate,
        at: CellRef,
        ctx: &EvalContext<'_>,
    ) -> Option<CellValue> {
        let formula = rule.formulas.get(slot)?;
        if !formula.relative {
            return Some(evaluate(aggregate.bases.get(slot)?.as_deref()?, ctx));
        }
        if !formula.shifts {
            let shifted = xlsx_parse::translate_formula(&formula.source, rule.anchor, at).ok()?;
            let parsed = self.parse(shifted)?;
            let values = aggregate.constants.get(slot).map_or(&[][..], Vec::as_slice);
            if values.is_empty() {
                return Some(evaluate(&parsed, ctx));
            }
            let mut expr = (*parsed).clone();
            read_in(&mut expr, &formula.constants, values);
            return Some(evaluate(&expr, ctx));
        }
        let mut expr = shift(
            aggregate.bases.get(slot)?.as_deref()?,
            i64::from(at.row) - i64::from(rule.anchor.row),
            i64::from(at.col) - i64::from(rule.anchor.col),
        );
        let tables = aggregate.tables.get(slot).map_or(&[][..], Vec::as_slice);
        for (path, table) in formula.counts.iter().zip(tables) {
            let (Some(table), Some(call)) = (table, node_mut(&mut expr, path)) else {
                continue;
            };
            let count = match &*call {
                Expr::FuncCall { args, .. } => args
                    .get(1)
                    .and_then(|criterion| table.count(&evaluate(criterion, ctx), ctx)),
                _ => None,
            };
            if let Some(count) = count {
                *call = Expr::Literal(CellValue::Number {
                    value: count as f64,
                });
            }
        }
        Some(evaluate(&expr, ctx))
    }

    fn parse(&mut self, text: String) -> Option<Arc<Expr>> {
        if let Some(parsed) = self.parsed.get(&text) {
            return parsed.clone();
        }
        if self.parsed.len() >= PARSED_CAP {
            self.parsed.clear();
        }
        self.stats.formula_parses += 1;
        let parsed = parse_formula(&text).ok().map(Arc::new);
        self.parsed.insert(text, parsed.clone());
        parsed
    }

    /// The rule's whole-range statistics and constant calls, computed the
    /// first time a cell of this epoch needs them.
    fn aggregate(
        &mut self,
        model: &Workbook,
        sheet: SheetId,
        index: usize,
        rule: &Rule,
        context: &mut EvalContext<'_>,
        spent: &mut u64,
    ) -> Arc<Aggregate> {
        let key = (sheet, index);
        if let Some(aggregate) = self.aggregates.get(&key) {
            return Arc::clone(aggregate);
        }
        let mut aggregate = Aggregate::default();
        let mut built = false;
        if matches!(
            rule.test,
            Test::Top { .. } | Test::Average { .. } | Test::Duplicates { .. }
        ) {
            let (stats, read) = stats(model, sheet, rule);
            aggregate.stats = Some(stats);
            *spent += read;
            built = true;
        }
        for formula in &rule.formulas {
            let mut values = Vec::with_capacity(formula.constants.len());
            for path in &formula.constants {
                let Some(call) = formula.base.as_deref().and_then(|base| node(base, path)) else {
                    break;
                };
                context.restart();
                context.cell = Some(rule.anchor);
                values.push(evaluate(call, context));
                *spent += MAX_EVALUATION_CELL_VISITS.saturating_sub(context.remaining_cells());
                aggregate.refused |= context.exhausted();
                built = true;
            }
            let mut tables = Vec::with_capacity(formula.counts.len());
            for path in &formula.counts {
                let range = formula
                    .base
                    .as_deref()
                    .and_then(|base| node(base, path))
                    .and_then(|call| match call {
                        Expr::FuncCall { args, .. } => args.first(),
                        _ => None,
                    });
                context.restart();
                context.cell = Some(rule.anchor);
                tables.push(range.and_then(|range| CountTable::build(range, &*context)));
                *spent += MAX_EVALUATION_CELL_VISITS.saturating_sub(context.remaining_cells());
                built = true;
            }
            let base = formula.base.as_ref().map(|base| {
                if values.is_empty() {
                    Arc::clone(base)
                } else {
                    let mut base = (**base).clone();
                    read_in(&mut base, &formula.constants, &values);
                    Arc::new(base)
                }
            });
            aggregate.bases.push(base);
            aggregate.constants.push(values);
            aggregate.tables.push(tables);
        }
        self.stats.aggregate_builds += u64::from(built);
        let aggregate = Arc::new(aggregate);
        self.aggregates.insert(key, Arc::clone(&aggregate));
        aggregate
    }
}

impl Rule {
    fn covers(&self, at: CellRef) -> bool {
        self.ranges.iter().any(|range| range.contains(at))
    }

    /// Why no cell can decide this rule, when that is known without a cell.
    fn undecidable(&self) -> Option<Unknown> {
        match self.test {
            Test::Undecidable(reason) => Some(reason),
            _ => None,
        }
    }
}

fn within(runs: &[Range<u32>], index: u32) -> bool {
    runs.iter().any(|run| run.contains(&index))
}

fn is_error(value: &CellValue) -> bool {
    matches!(value, CellValue::Error { .. })
}

fn compare(ctx: &EvalContext<'_>, op: BinaryOp, lhs: &CellValue, rhs: &CellValue) -> bool {
    let expr = Expr::Binary {
        op,
        lhs: Box::new(Expr::Literal(lhs.clone())),
        rhs: Box::new(Expr::Literal(rhs.clone())),
    };
    evaluate(&expr, ctx) == CellValue::Bool { value: true }
}

/// The sheet's rules that paint a differential format, highest priority first;
/// `dated` when the cells were calculated with a date.
fn rules(model: &Workbook, sheet: SheetId, dated: bool) -> Vec<Rule> {
    let Some(owner) = model.sheet(sheet) else {
        return Vec::new();
    };
    let mut rules: Vec<(u32, Rule)> = owner
        .conditional_formats
        .iter()
        .flat_map(|group| {
            let ranges: Vec<CellRange> = group
                .sqref
                .split_whitespace()
                .filter_map(|range| CellRange::parse_a1(range).ok())
                .collect();
            group.rules.iter().filter_map(move |rule| {
                let attribute = |name: &str| {
                    rule.raw_xml
                        .as_deref()
                        .and_then(|xml| xlsx_parse::cf_rule_attribute(xml, name))
                };
                let flag = |name: &str, default: bool| {
                    attribute(name).map_or(default, |value| value == "1" || value == "true")
                };
                let stop = flag("stopIfTrue", false);
                let supported = match rule.rule_type.as_str() {
                    "cellIs" => rule
                        .operator
                        .as_deref()
                        .and_then(operator)
                        .map(|op| (Test::CellIs(op), rule.formulas.clone())),
                    "expression" | "containsText" | "notContainsText" | "beginsWith"
                    | "endsWith" | "containsBlanks" | "notContainsBlanks" | "containsErrors"
                    | "notContainsErrors" => Some((
                        Test::Expression,
                        rule.formulas.iter().take(1).cloned().collect(),
                    )),
                    "top10" => attribute("rank")
                        .and_then(|rank| rank.parse().ok())
                        .map(|rank| {
                            let test = Test::Top {
                                rank,
                                percent: flag("percent", false),
                                bottom: flag("bottom", false),
                            };
                            (test, Vec::new())
                        }),
                    "aboveAverage" => Some((
                        Test::Average {
                            above: flag("aboveAverage", true),
                            equal: flag("equalAverage", false),
                            deviations: attribute("stdDev")
                                .and_then(|deviations| deviations.parse().ok())
                                .unwrap_or(0.0),
                        },
                        Vec::new(),
                    )),
                    "duplicateValues" => Some((Test::Duplicates { unique: false }, Vec::new())),
                    "uniqueValues" => Some((Test::Duplicates { unique: true }, Vec::new())),
                    "colorScale" | "dataBar" | "iconSet" => return None,
                    _ => None,
                };
                let paints =
                    stop || rule.dxf_id.is_some() || rule.format != DifferentialFormat::default();
                let (test, sources) = match supported {
                    Some(supported) => supported,
                    None if paints => (Test::Undecidable(Unknown::Unsupported), Vec::new()),
                    None => return None,
                };
                let anchor = ranges.first()?.start;
                let formulas: Vec<Formula> = sources
                    .iter()
                    .map(|source| compile(source, anchor))
                    .collect();
                let reads_volatile = formulas.iter().any(|formula| {
                    formula
                        .base
                        .as_deref()
                        .is_some_and(|base| volatile(model, sheet, base, dated))
                });
                let whole_range = matches!(
                    test,
                    Test::Top { .. } | Test::Average { .. } | Test::Duplicates { .. }
                );
                let test = if reads_volatile {
                    Test::Undecidable(Unknown::Volatile)
                } else if whole_range && overlapping(&ranges) {
                    Test::Undecidable(Unknown::Overlapping)
                } else {
                    test
                };
                let fill_unknown = rule.dxf_id.is_some_and(|id| {
                    model.styles.dxf_fill_colors_differ(id)
                        && model.styles.dxfs.get(id as usize) == Some(&rule.format)
                });
                let format = if fill_unknown {
                    DifferentialFormat {
                        fill: None,
                        ..rule.format.clone()
                    }
                } else {
                    rule.format.clone()
                };
                Some((
                    rule.priority,
                    Rule {
                        anchor,
                        ranges: ranges.clone(),
                        test,
                        formulas,
                        format: Arc::new(format),
                        fill_unknown,
                        stop,
                    },
                ))
            })
        })
        .collect();
    rules.sort_by_key(|(priority, _)| *priority);
    rules.into_iter().map(|(_, rule)| rule).collect()
}

fn overlapping(ranges: &[CellRange]) -> bool {
    ranges
        .iter()
        .enumerate()
        .any(|(index, range)| ranges[..index].iter().any(|seen| seen.overlaps(range)))
}

/// Whether `expr` at `sheet`, or a defined name it reads at any depth, calls
/// a function whose value changes with no edit and nothing decides: one in
/// `RANDOM`, or in `CLOCK` unless `dated`. Names past `NAME_DEPTH` count as
/// such a call.
fn volatile(model: &Workbook, sheet: SheetId, expr: &Expr, dated: bool) -> bool {
    let mut names = Vec::new();
    if calls_volatile(model, sheet, expr, &mut names, dated) {
        return true;
    }
    let mut expanded = HashSet::new();
    while let Some((owner, name)) = names.pop() {
        if !expanded.insert((owner, name.to_lowercase())) {
            continue;
        }
        if expanded.len() > NAME_DEPTH {
            return true;
        }
        let Some(defined) = model.defined_name(owner, &name) else {
            continue;
        };
        let formula = &defined.formula;
        let Ok(definition) = parse_formula(formula.strip_prefix('=').unwrap_or(formula)) else {
            continue;
        };
        let scope = defined.local_sheet.unwrap_or(owner);
        if calls_volatile(model, scope, &definition, &mut names, dated) {
            return true;
        }
    }
    false
}

/// Whether `expr` calls a function in `RANDOM`, or in `CLOCK` unless
/// `dated`, collecting the defined names it reads into `names`.
fn calls_volatile(
    model: &Workbook,
    sheet: SheetId,
    expr: &Expr,
    names: &mut Vec<(SheetId, String)>,
    dated: bool,
) -> bool {
    match expr {
        Expr::FuncCall { name, args, .. } => {
            let upper = name.to_ascii_uppercase();
            let bare = upper.strip_prefix("_XLFN.").unwrap_or(&upper);
            RANDOM.contains(&bare)
                || (!dated && CLOCK.contains(&bare))
                || args
                    .iter()
                    .any(|arg| calls_volatile(model, sheet, arg, names, dated))
        }
        Expr::Name { scope, name } => {
            let owner = match scope {
                Some(scope) => model.sheet_by_name(scope).map(|(id, _)| id),
                None => Some(sheet),
            };
            names.extend(owner.map(|owner| (owner, name.clone())));
            false
        }
        Expr::RangeJoin { start, end } => {
            calls_volatile(model, sheet, start, names, dated)
                || calls_volatile(model, sheet, end, names, dated)
        }
        other => children(other)
            .into_iter()
            .any(|child| calls_volatile(model, sheet, child, names, dated)),
    }
}

fn operator(name: &str) -> Option<Op> {
    Some(match name {
        "lessThan" => Op::Compare(BinaryOp::Lt),
        "lessThanOrEqual" => Op::Compare(BinaryOp::Le),
        "equal" => Op::Compare(BinaryOp::Eq),
        "notEqual" => Op::Compare(BinaryOp::Ne),
        "greaterThanOrEqual" => Op::Compare(BinaryOp::Ge),
        "greaterThan" => Op::Compare(BinaryOp::Gt),
        "between" => Op::Between(true),
        "notBetween" => Op::Between(false),
        _ => return None,
    })
}

/// `source` parsed at `anchor`, with what shifting it by a row and a column
/// leaves unchanged: the whole text, or calls over absolute references.
fn compile(source: &str, anchor: CellRef) -> Formula {
    let step = |index: u32, max: u32| {
        if index + 1 < max {
            index + 1
        } else {
            index - 1
        }
    };
    let probe = CellRef::new(step(anchor.row, MAX_ROWS), step(anchor.col, MAX_COLS));
    let at_anchor = xlsx_parse::translate_formula(source, anchor, anchor).ok();
    let shifted = xlsx_parse::translate_formula(source, anchor, probe).ok();
    let relative = at_anchor.is_none() || shifted != at_anchor;
    let base = at_anchor
        .as_deref()
        .and_then(|text| parse_formula(text).ok())
        .map(Arc::new);
    let moved = shifted.as_deref().and_then(|text| parse_formula(text).ok());
    let mut constants = Vec::new();
    let mut counts = Vec::new();
    let mut shifts = false;
    if let (Some(base), Some(moved)) = (&base, &moved) {
        fixed_calls(base, moved, &mut Vec::new(), &mut constants, &mut counts);
        let (down, right) = (
            i64::from(probe.row) - i64::from(anchor.row),
            i64::from(probe.col) - i64::from(anchor.col),
        );
        shifts = !relative_whole(base) && shift(base, down, right) == *moved;
    }
    if !shifts {
        counts.clear();
    }
    Formula {
        source: source.to_owned(),
        base,
        relative,
        shifts,
        constants,
        counts,
        parses: u8::from(at_anchor.is_some()) + u8::from(shifted.is_some()),
    }
}

/// Collects the paths of the calls `base` and `shifted` share that aggregate
/// absolute references alone, and of the `COUNTIF` calls whose range stays
/// while their criterion reference moves.
fn fixed_calls(
    base: &Expr,
    shifted: &Expr,
    path: &mut Vec<usize>,
    constants: &mut Vec<Vec<usize>>,
    counts: &mut Vec<Vec<usize>>,
) {
    if base == shifted && aggregates_absolute(base) {
        constants.push(path.clone());
        return;
    }
    if counts_each_cell(base, shifted) {
        counts.push(path.clone());
        return;
    }
    if std::mem::discriminant(base) != std::mem::discriminant(shifted) {
        return;
    }
    let (left, right) = (children(base), children(shifted));
    if left.len() != right.len() {
        return;
    }
    for (index, (left, right)) in left.into_iter().zip(right).enumerate() {
        path.push(index);
        fixed_calls(left, right, path, constants, counts);
        path.pop();
    }
}

/// Whether `base` is `COUNTIF(range, reference)` and `shifted` keeps the range
/// but moves the reference.
fn counts_each_cell(base: &Expr, shifted: &Expr) -> bool {
    let Expr::FuncCall { name, args, .. } = base else {
        return false;
    };
    let Expr::FuncCall { args: moved, .. } = shifted else {
        return false;
    };
    let upper = name.to_ascii_uppercase();
    upper.strip_prefix("_XLFN.").unwrap_or(&upper) == "COUNTIF"
        && args.len() == 2
        && moved.len() == 2
        && args[0] == moved[0]
        && is_reference(&args[0])
        && matches!(args[1], Expr::Ref { .. })
        && args[1] != moved[1]
}

/// Whether `expr` holds a whole-row or whole-column reference with a relative
/// end, which `shift` and `translate_formula` move apart near the sheet's edge.
fn relative_whole(expr: &Expr) -> bool {
    match expr {
        Expr::ColumnRange { range, .. } => !range.abs_start || !range.abs_end,
        Expr::RowRange { range, .. } => !range.abs_start || !range.abs_end,
        Expr::RangeJoin { start, end } => relative_whole(start) || relative_whole(end),
        other => children(other).into_iter().any(relative_whole),
    }
}

fn is_reference(arg: &Expr) -> bool {
    matches!(
        arg,
        Expr::Ref { .. } | Expr::Range { .. } | Expr::ColumnRange { .. } | Expr::RowRange { .. }
    )
}

fn aggregates_absolute(expr: &Expr) -> bool {
    let Expr::FuncCall { name, args, .. } = expr else {
        return false;
    };
    let upper = name.to_ascii_uppercase();
    let name = upper.strip_prefix("_XLFN.").unwrap_or(&upper);
    AGGREGATES.contains(&name)
        && args.iter().any(is_reference)
        && args.iter().enumerate().all(|(index, arg)| match arg {
            Expr::Range { .. } | Expr::ColumnRange { .. } | Expr::RowRange { .. } => {
                reads_a_range(name, index)
            }
            Expr::Ref { .. } | Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) => true,
            _ => false,
        })
}

/// Whether `name` reads its argument at `index` as a range rather than as
/// one value, which a multi-cell reference would give per cell.
fn reads_a_range(name: &str, index: usize) -> bool {
    match name {
        "LARGE" | "SMALL" => index == 0,
        "COUNTIF" | "SUMIF" | "AVERAGEIF" => index != 1,
        "COUNTIFS" => index.is_multiple_of(2),
        "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" => index == 0 || !index.is_multiple_of(2),
        _ => true,
    }
}

/// `expr` with its relative references moved `down` rows and `right` columns,
/// as `translate_formula` moves them in text; one moved off the sheet reads
/// `#REF!`.
fn shift(expr: &Expr, down: i64, right: i64) -> Expr {
    let moved = |index: u32, absolute: bool, delta: i64, limit: u32| {
        let index = i64::from(index) + if absolute { 0 } else { delta };
        u32::try_from(index).ok().filter(|&index| index < limit)
    };
    let cell = |at: CellRef| {
        Some(CellRef {
            row: moved(at.row, at.abs_row, down, MAX_ROWS)?,
            col: moved(at.col, at.abs_col, right, MAX_COLS)?,
            ..at
        })
    };
    let all = |values: &[Expr]| -> Vec<Expr> {
        values
            .iter()
            .map(|value| shift(value, down, right))
            .collect()
    };
    let boxed = |inner: &Expr| Box::new(shift(inner, down, right));
    match expr {
        Expr::Ref { sheet, cell: at } => match cell(*at) {
            Some(at) => Expr::Ref {
                sheet: sheet.clone(),
                cell: at,
            },
            None => Expr::Error(ErrorValue::Ref),
        },
        Expr::Range { sheet, range } => match (cell(range.start), cell(range.end)) {
            (Some(start), Some(end)) => Expr::Range {
                sheet: sheet.clone(),
                range: CellRange::new(start, end),
            },
            _ => Expr::Error(ErrorValue::Ref),
        },
        Expr::ColumnRange { sheet, range } => match (
            moved(range.start, range.abs_start, right, MAX_COLS),
            moved(range.end, range.abs_end, right, MAX_COLS),
        ) {
            (Some(start), Some(end)) => Expr::ColumnRange {
                sheet: sheet.clone(),
                range: ColumnRange {
                    start: start.min(end),
                    end: start.max(end),
                    ..*range
                },
            },
            _ => Expr::Error(ErrorValue::Ref),
        },
        Expr::RowRange { sheet, range } => match (
            moved(range.start, range.abs_start, down, MAX_ROWS),
            moved(range.end, range.abs_end, down, MAX_ROWS),
        ) {
            (Some(start), Some(end)) => Expr::RowRange {
                sheet: sheet.clone(),
                range: RowRange {
                    start: start.min(end),
                    end: start.max(end),
                    ..*range
                },
            },
            _ => Expr::Error(ErrorValue::Ref),
        },
        Expr::ArrayLiteral { cols, values } => Expr::ArrayLiteral {
            cols: *cols,
            values: all(values),
        },
        Expr::Unary { op, expr: inner } => Expr::Unary {
            op: *op,
            expr: boxed(inner),
        },
        Expr::Binary { op, lhs, rhs } => Expr::Binary {
            op: *op,
            lhs: boxed(lhs),
            rhs: boxed(rhs),
        },
        Expr::Percent(inner) => Expr::Percent(boxed(inner)),
        Expr::RangeJoin { start, end } => Expr::RangeJoin {
            start: boxed(start),
            end: boxed(end),
        },
        Expr::FuncCall { name, func, args } => Expr::FuncCall {
            name: name.clone(),
            func: *func,
            args: all(args),
        },
        other => other.clone(),
    }
}

/// Writes `values` over the calls at `paths` in `expr`.
fn read_in(expr: &mut Expr, paths: &[Vec<usize>], values: &[CellValue]) {
    for (path, value) in paths.iter().zip(values) {
        if let Some(call) = node_mut(expr, path)
            && matches!(call, Expr::FuncCall { .. })
        {
            *call = Expr::Literal(value.clone());
        }
    }
}

/// The operands `fixed_calls` descends into; a range join's ends stay
/// whole, since they are references.
fn children(expr: &Expr) -> Vec<&Expr> {
    match expr {
        Expr::ArrayLiteral { values, .. } | Expr::FuncCall { args: values, .. } => {
            values.iter().collect()
        }
        Expr::Unary { expr: inner, .. } | Expr::Percent(inner) => vec![&**inner],
        Expr::Binary { lhs, rhs, .. } => vec![&**lhs, &**rhs],
        _ => Vec::new(),
    }
}

fn node<'e>(expr: &'e Expr, path: &[usize]) -> Option<&'e Expr> {
    path.iter()
        .try_fold(expr, |at, &index| children(at).get(index).copied())
}

fn node_mut<'e>(expr: &'e mut Expr, path: &[usize]) -> Option<&'e mut Expr> {
    let Some((&first, rest)) = path.split_first() else {
        return Some(expr);
    };
    let child = match expr {
        Expr::ArrayLiteral { values, .. } | Expr::FuncCall { args: values, .. } => {
            values.get_mut(first)?
        }
        Expr::Unary { expr: inner, .. } | Expr::Percent(inner) if first == 0 => &mut **inner,
        Expr::Binary { lhs, .. } if first == 0 => &mut **lhs,
        Expr::Binary { rhs, .. } if first == 1 => &mut **rhs,
        _ => return None,
    };
    node_mut(child, rest)
}

/// Statistics over the rule's whole range, and how many cells they read.
fn stats(model: &Workbook, sheet: SheetId, rule: &Rule) -> (Stats, u64) {
    let cells: Vec<&CellValue> = model
        .sheet(sheet)
        .into_iter()
        .flat_map(|owner| {
            rule.ranges
                .iter()
                .flat_map(move |range| owner.cells_in_range(*range))
                .map(|(_, cell)| &cell.value)
        })
        .collect();
    let read = cells.len() as u64;
    let numbers = || {
        cells.iter().filter_map(|value| match value {
            CellValue::Number { value } => Some(*value),
            _ => None,
        })
    };
    let stats = match rule.test {
        Test::Top {
            rank,
            percent,
            bottom,
        } => {
            let mut values: Vec<f64> = numbers().collect();
            values.sort_by(|a, b| {
                if bottom {
                    a.total_cmp(b)
                } else {
                    b.total_cmp(a)
                }
            });
            let count = if percent {
                (values.len() as u64 * u64::from(rank) / 100).max(1) as usize
            } else {
                rank as usize
            };
            Stats::Threshold(
                count
                    .checked_sub(1)
                    .and_then(|last| values.get(last.min(values.len().saturating_sub(1))))
                    .copied(),
            )
        }
        Test::Average { .. } => {
            let values: Vec<f64> = numbers().collect();
            let n = values.len() as f64;
            let mean = values.iter().sum::<f64>() / n;
            let variance =
                values.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / (n - 1.0).max(1.0);
            Stats::Average {
                mean,
                deviation: variance.sqrt(),
            }
        }
        _ => {
            let mut counts = HashMap::new();
            for key in cells.iter().copied().filter_map(key) {
                *counts.entry(key).or_insert(0) += 1;
            }
            Stats::Counts(counts)
        }
    };
    (stats, read)
}

/// Cells a rule's whole-range statistics, constant calls and count tables
/// read at most, each area cut to the rows its sheet uses.
fn build_cost(model: &Workbook, sheet: SheetId, rule: &Rule) -> u64 {
    let whole_range = matches!(
        rule.test,
        Test::Top { .. } | Test::Average { .. } | Test::Duplicates { .. }
    );
    let ranges: u64 = if whole_range {
        rule.ranges
            .iter()
            .map(|range| cells_used(model, sheet, range))
            .sum()
    } else {
        0
    };
    let calls: u64 = rule
        .formulas
        .iter()
        .flat_map(|formula| {
            formula
                .constants
                .iter()
                .chain(&formula.counts)
                .filter_map(move |path| formula.base.as_deref().and_then(|base| node(base, path)))
        })
        .flat_map(children)
        .map(|arg| reference_cells(model, sheet, arg))
        .sum();
    ranges + calls
}

fn reference_cells(model: &Workbook, sheet: SheetId, arg: &Expr) -> u64 {
    let (name, area) = match arg {
        Expr::Ref { .. } => return 1,
        Expr::Range { sheet: name, range } => (name, *range),
        Expr::ColumnRange { sheet: name, range } => (
            name,
            CellRange::new(
                CellRef::new(0, range.start),
                CellRef::new(MAX_ROWS - 1, range.end),
            ),
        ),
        Expr::RowRange { sheet: name, range } => (
            name,
            CellRange::new(
                CellRef::new(range.start, 0),
                CellRef::new(range.end, MAX_COLS - 1),
            ),
        ),
        _ => return 0,
    };
    let owner = match name {
        Some(name) => model.sheet_by_name(name).map_or(sheet, |(id, _)| id),
        None => sheet,
    };
    cells_used(model, owner, &area)
}

/// The cells of `range` on `sheet` above the last row the sheet uses.
fn cells_used(model: &Workbook, sheet: SheetId, range: &CellRange) -> u64 {
    let used = xlsx_model::CellProvider::used_rows(model, sheet);
    let rows = used
        .min(range.end.row.saturating_add(1))
        .saturating_sub(range.start.row);
    u64::from(rows) * u64::from(range.end.col - range.start.col + 1)
}

fn key(value: &CellValue) -> Option<Key> {
    Some(match value {
        CellValue::Number { value } => Key::Number((value + 0.0).to_bits()),
        CellValue::Text { value } => Key::Text(value.to_lowercase()),
        CellValue::Bool { value } => Key::Bool(*value),
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shifting_a_parsed_formula_matches_parsing_its_translated_text() {
        let anchor = CellRef::new(4, 3);
        for source in [
            "A1+$B$1>50",
            "COUNTIF($A$1:$A$50000,A1)>1",
            "SUM(A$1:B$20)+$C3",
            "AND(MOD(D5,3)=0,Sheet2!E6<>\"x\")",
            "SUM(A:A)+SUM($B:C)+SUM(2:3)",
            "A1:INDEX($B$1:$B$9,2)",
            "{1,2;3,4}",
            "-A1%",
        ] {
            let text = xlsx_parse::translate_formula(source, anchor, anchor).unwrap();
            let base = parse_formula(&text).unwrap();
            for at in [
                CellRef::new(4, 3),
                CellRef::new(9, 3),
                CellRef::new(4, 8),
                CellRef::new(1000, 50),
            ] {
                let text = xlsx_parse::translate_formula(source, anchor, at).unwrap();
                let shifted = shift(&base, i64::from(at.row) - 4, i64::from(at.col) - 3);
                assert_eq!(shifted, parse_formula(&text).unwrap(), "{source} at {at:?}");
            }
        }
    }

    #[test]
    fn a_countif_over_a_fixed_range_compiles_to_a_count_table() {
        let anchor = CellRef::new(0, 0);
        assert_eq!(
            compile("COUNTIF($A$1:$A$50000,A1)>1", anchor).counts,
            [vec![0]]
        );
        for source in [
            "COUNTIF(A1:A50000,A1)>1",
            "COUNTIF($A$1:$A$50000,$A$1)>1",
            "COUNTIF($A$1:$A$50000,A1&\"*\")>1",
        ] {
            assert!(compile(source, anchor).counts.is_empty(), "{source}");
        }
    }

    fn number(value: f64) -> CellValue {
        CellValue::Number { value }
    }

    fn text(value: &str) -> CellValue {
        CellValue::Text {
            value: value.to_owned(),
        }
    }

    fn error(value: ErrorValue) -> CellValue {
        CellValue::Error { value }
    }

    fn put(sheet: &mut xlsx_model::Sheet, at: CellRef, value: CellValue, formula: Option<&str>) {
        sheet.set_cell(
            at,
            xlsx_model::Cell {
                value,
                formula: formula.map(str::to_owned),
                ..xlsx_model::Cell::default()
            },
        );
    }

    fn expression_rule(sqref: &str, source: &str) -> Rule {
        let ranges: Vec<CellRange> = sqref
            .split_whitespace()
            .map(|range| CellRange::parse_a1(range).unwrap())
            .collect();
        let anchor = ranges[0].start;
        Rule {
            formulas: vec![compile(source, anchor)],
            anchor,
            ranges,
            test: Test::Expression,
            format: Arc::new(DifferentialFormat::default()),
            fill_unknown: false,
            stop: false,
        }
    }

    /// Each cell of `rule` whose formula value through `cache` differs from
    /// evaluating its text translated to that cell, `None` where that text
    /// does not parse.
    fn disagreements(cache: &mut Cache, model: &Workbook, rule: &Rule) -> Vec<String> {
        let mut context = EvalContext::new(model, SheetId(0));
        let mut spent = 0;
        let aggregate = cache.aggregate(model, SheetId(0), 0, rule, &mut context, &mut spent);
        let source = &rule.formulas[0].source;
        let mut out = Vec::new();
        for range in &rule.ranges {
            for row in range.start.row..=range.end.row {
                for col in range.start.col..=range.end.col {
                    let at = CellRef::new(row, col);
                    context.restart();
                    context.cell = Some(at);
                    let cached = cache.formula(rule, 0, &aggregate, at, &context);
                    context.restart();
                    let reference = xlsx_parse::translate_formula(source, rule.anchor, at)
                        .ok()
                        .and_then(|text| parse_formula(&text).ok())
                        .map(|expr| evaluate(&expr, &context));
                    if cached != reference {
                        out.push(format!(
                            "{source} at {}: {cached:?} != {reference:?}",
                            at.to_a1()
                        ));
                    }
                }
            }
        }
        out
    }

    #[test]
    fn a_shifted_rule_formula_evaluates_as_its_translated_text_at_every_cell() {
        let mut model = Workbook::default();
        for name in ["Sheet1", "Other", "Sheet2"] {
            let mut sheet = xlsx_model::Sheet::new(name);
            for row in 0..12 {
                for col in 0..6 {
                    put(
                        &mut sheet,
                        CellRef::new(row, col),
                        number(f64::from(row * 10 + col + 1)),
                        None,
                    );
                }
            }
            model.sheets.push(sheet);
        }
        put(
            &mut model.sheets[1],
            CellRef::new(MAX_ROWS - 3, 0),
            number(7.0),
            None,
        );
        for (name, formula) in [("Rate", "Sheet1!$C$1"), ("Left", "Sheet1!A1")] {
            model.defined_names.push(xlsx_model::DefinedName {
                name: name.to_owned(),
                formula: formula.to_owned(),
                local_sheet: None,
                hidden: false,
            });
        }
        let mut mismatches = Vec::new();
        for (sqref, source) in [
            ("B2:D5", "A1+$B1+C$1+$D$2"),
            ("B2:D5 F1:G3", "A1+$B1+C$1"),
            ("B2:D5", "SUM(1:1)+SUM(A:A)+SUM($A:$A)+SUM($C:A)"),
            ("C3:E8", "'Other'!A1+SUM(Sheet2!$B$2:B10)"),
            ("B2:D5", "Rate*A1+Left"),
            ("A1048570:B1048576", "A1048574+B$1048576"),
            ("A1048570:B1048576", "SUM(A1048574:B1048575)"),
            ("A1048570:B1048576", "'Other'!A1048574+1"),
            ("A1048576:C1048576", "A1+B1048576"),
            ("XFA1:XFD3", "XFC1+SUM(XFB1:XFC2)"),
            ("XFD1:XFD4", "A1+XFD1"),
            ("A1048570:A1048576", "SUM(1048574:1048575)"),
            ("XFA1:XFD3", "SUM(XFB:XFC)"),
            ("B1:C3 A5:A6", "SUM(A:B)"),
        ] {
            let rule = expression_rule(sqref, source);
            mismatches.extend(disagreements(&mut Cache::default(), &model, &rule));
        }
        assert!(mismatches.is_empty(), "{mismatches:#?}");
    }

    #[test]
    fn an_absolute_call_reading_a_range_as_one_value_is_evaluated_per_cell() {
        let mut sheet = xlsx_model::Sheet::new("Sheet1");
        for (row, (b, c)) in [(1.0, 1.0), (2.0, 3.0)].into_iter().enumerate() {
            put(&mut sheet, CellRef::new(row as u32, 1), number(b), None);
            put(&mut sheet, CellRef::new(row as u32, 2), number(c), None);
        }
        let mut model = Workbook::default();
        model.sheets.push(sheet);
        let anchor = CellRef::new(0, 0);
        for source in [
            "COUNTIF($B$1:$B$2,$C$1:$C$2)>0",
            "SUMIF($B$1:$B$2,$C$1:$C$2)>0",
            "LARGE($B$1:$B$2,$C$1:$C$2)>0",
            "COUNTIFS($B$1:$B$2,$C$1:$C$2)>0",
        ] {
            assert!(compile(source, anchor).constants.is_empty(), "{source}");
        }
        for source in [
            "COUNTIF($B$1:$B$2,$C$1)>0",
            "SUMIF($B$1:$B$2,\">0\",$C$1:$C$2)>0",
            "SUM($B$1:$B$2,$C$1:$C$2)>0",
        ] {
            assert_eq!(compile(source, anchor).constants, [vec![0]], "{source}");
        }
        let rule = expression_rule("A1:A2", "COUNTIF($B$1:$B$2,$C$1:$C$2)>0");
        let mismatches = disagreements(&mut Cache::default(), &model, &rule);
        assert!(mismatches.is_empty(), "{mismatches:#?}");
    }

    /// Checks `rule` cell by cell against the general path in a new `epoch`,
    /// and its counts at the first and thirteenth cells against `expected`.
    fn check_counts(
        cache: &mut Cache,
        model: &Workbook,
        rule: &Rule,
        epoch: u64,
        expected: [f64; 2],
    ) {
        let rows = rule.ranges[0].end.row + 1;
        cache.effects(
            (epoch, None),
            model,
            SheetId(0),
            std::slice::from_ref(&(0..rows)),
            std::slice::from_ref(&(1..2)),
            Budget::Complete,
        );
        let mismatches = disagreements(cache, model, rule);
        assert!(mismatches.is_empty(), "{mismatches:#?}");
        let mut context = EvalContext::new(model, SheetId(0));
        let mut spent = 0;
        let aggregate = cache.aggregate(model, SheetId(0), 0, rule, &mut context, &mut spent);
        let table = aggregate.tables[0][0].as_ref().unwrap();
        let answered = (0..rows)
            .filter(|&row| {
                let criterion = model
                    .sheet(SheetId(0))
                    .and_then(|sheet| sheet.cell(CellRef::new(row, 1)))
                    .map_or(CellValue::Empty, |cell| cell.value.clone());
                table.count(&criterion, &context).is_some()
            })
            .count();
        assert!(answered >= 10, "{answered}");
        for (row, count) in [0, 12].into_iter().zip(expected) {
            let at = CellRef::new(row, 1);
            context.restart();
            context.cell = Some(at);
            assert_eq!(
                cache.formula(rule, 0, &aggregate, at, &context),
                Some(number(count)),
                "{}",
                at.to_a1()
            );
        }
    }

    #[test]
    fn a_count_table_counts_as_countif_does_before_and_after_an_edit() {
        let mut data = xlsx_model::Sheet::new("Sheet1");
        let values = [
            text("abc"),
            text("ABC"),
            text("Abc "),
            number(1.0),
            text("1"),
            text("1.0"),
            CellValue::Bool { value: true },
            text("TRUE"),
            error(ErrorValue::NA),
            error(ErrorValue::Div0),
            CellValue::Empty,
            CellValue::Bool { value: false },
            number(0.0),
        ];
        for (row, value) in values.into_iter().enumerate() {
            if value != CellValue::Empty {
                put(&mut data, CellRef::new(row as u32, 0), value, None);
            }
        }
        put(&mut data, CellRef::new(13, 0), number(5.0), Some("C1"));
        put(&mut data, CellRef::new(14, 0), text("X"), Some("UPPER(C2)"));
        put(&mut data, CellRef::new(0, 2), number(5.0), None);
        put(&mut data, CellRef::new(1, 2), text("x"), None);
        let criteria = [
            text("abc"),
            text("ABC"),
            number(1.0),
            text("1"),
            text("1.0"),
            CellValue::Bool { value: true },
            CellValue::Bool { value: false },
            error(ErrorValue::NA),
            error(ErrorValue::Div0),
            CellValue::Empty,
            text(""),
            text("x"),
            number(5.0),
            number(0.0),
            text("true"),
            text("abc*"),
            text(">0"),
        ];
        let rows = criteria.len();
        for (row, value) in criteria.into_iter().enumerate() {
            if value != CellValue::Empty {
                put(&mut data, CellRef::new(row as u32, 1), value, None);
            }
        }
        let mut model = Workbook::default();
        model.sheets.push(data);
        let mut workbook = crate::Workbook::from_model(model).unwrap();
        let rule = expression_rule(&format!("B1:B{rows}"), "COUNTIF($A$1:$A$15,B1)");
        assert!(rule.formulas[0].shifts);
        assert_eq!(rule.formulas[0].counts, [Vec::<usize>::new()]);
        let mut cache = Cache::default();
        check_counts(&mut cache, workbook.model(), &rule, 0, [2.0, 1.0]);
        workbook
            .edit_cells(
                SheetId(0),
                &[
                    crate::CellInput {
                        cell: CellRef::new(0, 2),
                        input: "ABC".into(),
                    },
                    crate::CellInput {
                        cell: CellRef::new(1, 2),
                        input: "abc".into(),
                    },
                ],
                crate::CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(
            workbook
                .model()
                .sheet(SheetId(0))
                .and_then(|sheet| sheet.cell(CellRef::new(14, 0)))
                .map(|cell| cell.value.clone()),
            Some(text("ABC"))
        );
        check_counts(&mut cache, workbook.model(), &rule, 1, [4.0, 0.0]);
    }

    #[test]
    fn a_rule_on_the_clock_is_decided_again_when_only_the_calculation_date_changes() {
        let mut sheet = xlsx_model::Sheet::new("Sheet1");
        for row in 0..8 {
            put(
                &mut sheet,
                CellRef::new(row, 0),
                number(45_999.0 + f64::from(row)),
                None,
            );
        }
        sheet
            .conditional_formats
            .push(xlsx_model::ConditionalFormatting {
                sqref: "A1:A8".to_owned(),
                rules: vec![xlsx_model::ConditionalFormatRule {
                    id: String::new(),
                    rule_type: "expression".to_owned(),
                    operator: None,
                    formulas: vec!["A1=TODAY()".to_owned()],
                    priority: 1,
                    dxf_id: None,
                    format: DifferentialFormat {
                        bold: Some(true),
                        ..DifferentialFormat::default()
                    },
                    raw_xml: None,
                }],
            });
        let mut model = Workbook::default();
        model.sheets.push(sheet);
        let mut cache = Cache::default();
        let mut painted = |now_serial: Option<f64>| {
            let (formats, pending) = cache.effects(
                (1, now_serial),
                &model,
                SheetId(0),
                std::slice::from_ref(&(0..8)),
                std::slice::from_ref(&(0..1)),
                Budget::Complete,
            );
            assert_eq!(pending, 0);
            formats.keys().map(|&(row, _)| row).collect::<Vec<u32>>()
        };
        assert_eq!(painted(Some(46_000.25)), [1]);
        assert_eq!(painted(Some(46_003.75)), [4]);
        assert!(painted(None).is_empty());
        assert_eq!(painted(Some(46_000.25)), [1]);
    }
}
