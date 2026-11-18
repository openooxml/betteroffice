//! Conditional-format effects of the cells a display list draws.

use std::collections::HashMap;
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
    /// time; a cell it leaves undecided paints its result from the epoch
    /// before when it has one, and later frames decide the rest.
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
    /// Rule checks plus cell visits the last display list spent.
    pub units: u64,
    /// Decided cells held.
    pub cached_cells: usize,
    /// Results of the epoch before held for cells not yet decided again.
    pub previous_cells: usize,
    /// Model epoch the cache holds.
    pub epoch: u64,
}

/// Each sheet's rules, aggregates and decided cells for one model epoch, the
/// decided cells of the epoch before, and shifted formulas parsed in any.
pub(crate) struct Cache {
    epoch: Option<u64>,
    rules: HashMap<SheetId, Arc<[Rule]>>,
    aggregates: HashMap<(SheetId, usize), Arc<Aggregate>>,
    results: HashMap<(SheetId, u32, u32), Option<Arc<DifferentialFormat>>>,
    /// The epoch before's results of the cells not yet decided in this one.
    previous: HashMap<(SheetId, u32, u32), Option<Arc<DifferentialFormat>>>,
    /// Whether a change since the epoch may have moved a cell or changed a
    /// rule, so that no result of the epoch belongs to the next.
    moved: bool,
    parsed: HashMap<String, Option<Arc<Expr>>>,
    budget: u64,
    stats: ConditionalStats,
}

impl Default for Cache {
    fn default() -> Self {
        Self {
            epoch: None,
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
            epoch: self.epoch.unwrap_or_default(),
            ..self.stats
        }
    }

    /// Marks the change that ends the epoch as one that may move cells or
    /// change rules.
    pub(crate) fn cells_moved(&mut self) {
        self.moved = true;
    }

    /// The effects on the cells of `sheet` in the `rows` by `cols` runs, and
    /// how many of them `budget` left undecided. A cell takes each property
    /// from the highest-priority matching rule that sets it, and no rule
    /// below a matching `stopIfTrue` rule applies.
    pub(crate) fn effects(
        &mut self,
        epoch: u64,
        model: &Workbook,
        sheet: SheetId,
        rows: &[Range<u32>],
        cols: &[Range<u32>],
        budget: Budget<'_>,
    ) -> (CellFormats, u32) {
        if self.epoch != Some(epoch) {
            self.epoch = Some(epoch);
            self.rules.clear();
            self.aggregates.clear();
            if std::mem::take(&mut self.moved) {
                self.previous.clear();
                self.results.clear();
            } else {
                self.previous = std::mem::take(&mut self.results);
            }
        }
        if !self.rules.contains_key(&sheet) {
            let built: Arc<[Rule]> = rules(model, sheet).into();
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
        self.stats.units = 0;
        if rules.is_empty() {
            return (formats, 0);
        }
        let (limit, clock) = match budget {
            Budget::Frame(None) => (self.budget, None),
            Budget::Frame(Some(clock)) => (self.budget.saturating_mul(TIMED_FRAMES), Some(clock)),
            Budget::Complete => (u64::MAX, None),
        };
        let mut deadline = None;
        let mut decided = 0u64;
        let mut checked = 0u64;
        let mut stopped = false;
        let mut context = EvalContext::new(model, sheet);
        let mut spent = 0;
        let mut pending = 0u32;
        let mut caching = true;
        for row in rows.iter().cloned().flatten() {
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
                if !candidates.iter().any(|&index| rules[index].covers(at)) {
                    continue;
                }
                let key = (sheet, row, col);
                let result = match self.results.get(&key).cloned() {
                    Some(result) => {
                        self.stats.hits += 1;
                        result
                    }
                    None => {
                        if !stopped
                            && spent < limit
                            && let Some(clock) = clock
                            && (decided.is_multiple_of(CLOCK_STRIDE)
                                || spent >= checked + CLOCK_UNITS)
                        {
                            let now = clock.now_ms();
                            checked = spent;
                            stopped = now >= *deadline.get_or_insert(now + CONDITIONAL_FRAME_MS);
                        }
                        if stopped || spent >= limit {
                            pending += 1;
                            match self.previous.get(&key) {
                                Some(result) => result.clone(),
                                None => continue,
                            }
                        } else {
                            decided += 1;
                            self.stats.evaluated_cells += 1;
                            let result = self.decide(
                                model,
                                sheet,
                                &rules,
                                &candidates,
                                at,
                                &mut context,
                                &mut spent,
                            );
                            if caching {
                                caching = self.keep(sheet, at, result.clone(), rows, cols);
                            }
                            result
                        }
                    }
                };
                if let Some(format) = result {
                    formats.insert((row, col), format);
                }
            }
        }
        self.stats.pending = pending;
        self.stats.units = spent;
        (formats, pending)
    }

    /// Holds a decided cell in place of its result from the epoch before,
    /// first dropping the cells outside the frame when full; false when the
    /// frame alone fills the cap.
    fn keep(
        &mut self,
        sheet: SheetId,
        at: CellRef,
        result: Option<Arc<DifferentialFormat>>,
        rows: &[Range<u32>],
        cols: &[Range<u32>],
    ) -> bool {
        if self.results.len() >= CONDITIONAL_RESULT_CAP {
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
    ) -> Option<Arc<DifferentialFormat>> {
        let value = model
            .sheet(sheet)
            .and_then(|owner| owner.cell(at))
            .map_or(CellValue::Empty, |cell| cell.value.clone());
        let mut merged: Option<Arc<DifferentialFormat>> = None;
        for &index in candidates {
            let rule = &rules[index];
            if !rule.covers(at) {
                continue;
            }
            *spent += 1;
            if !self.matches(model, sheet, index, rule, &value, at, context, spent) {
                continue;
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
                    fill: format.fill.clone().or_else(|| rule.format.fill.clone()),
                }),
            });
            if rule.stop {
                break;
            }
        }
        merged
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
    ) -> bool {
        let aggregate = self.aggregate(model, sheet, index, rule, context, spent);
        context.restart();
        context.cell = Some(at);
        let ctx: &EvalContext<'_> = context;
        let matched = match &rule.test {
            Test::CellIs(op) => {
                if matches!(value, CellValue::Error { .. }) {
                    false
                } else {
                    match *op {
                        Op::Compare(op) => self
                            .formula(rule, 0, &aggregate, at, ctx)
                            .is_some_and(|rhs| compare(ctx, op, value, &rhs)),
                        Op::Between(inside) => {
                            let low = self.formula(rule, 0, &aggregate, at, ctx);
                            let high = self.formula(rule, 1, &aggregate, at, ctx);
                            match (low, high) {
                                (Some(low), Some(high)) => {
                                    let within = |low: &CellValue, high: &CellValue| {
                                        compare(ctx, BinaryOp::Ge, value, low)
                                            && compare(ctx, BinaryOp::Le, value, high)
                                    };
                                    (within(&low, &high) || within(&high, &low)) == inside
                                }
                                _ => false,
                            }
                        }
                    }
                }
            }
            Test::Expression => match self.formula(rule, 0, &aggregate, at, ctx) {
                Some(CellValue::Bool { value }) => value,
                Some(CellValue::Number { value }) => value != 0.0,
                _ => false,
            },
            Test::Top { bottom, .. } => match (value, &aggregate.stats) {
                (CellValue::Number { value }, Some(Stats::Threshold(Some(threshold)))) => {
                    if *bottom {
                        value <= threshold
                    } else {
                        value >= threshold
                    }
                }
                _ => false,
            },
            Test::Average {
                above,
                equal,
                deviations,
            } => match (value, &aggregate.stats) {
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
            },
            Test::Duplicates { unique } => match (key(value), &aggregate.stats) {
                (Some(key), Some(Stats::Counts(counts))) => {
                    let count = counts.get(&key).copied().unwrap_or(0);
                    if *unique { count == 1 } else { count > 1 }
                }
                _ => false,
            },
        };
        *spent += MAX_EVALUATION_CELL_VISITS.saturating_sub(ctx.remaining_cells());
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
}

fn within(runs: &[Range<u32>], index: u32) -> bool {
    runs.iter().any(|run| run.contains(&index))
}

fn compare(ctx: &EvalContext<'_>, op: BinaryOp, lhs: &CellValue, rhs: &CellValue) -> bool {
    let expr = Expr::Binary {
        op,
        lhs: Box::new(Expr::Literal(lhs.clone())),
        rhs: Box::new(Expr::Literal(rhs.clone())),
    };
    evaluate(&expr, ctx) == CellValue::Bool { value: true }
}

/// The sheet's rules that paint a differential format, highest priority first.
fn rules(model: &Workbook, sheet: SheetId) -> Vec<Rule> {
    let Some(sheet) = model.sheet(sheet) else {
        return Vec::new();
    };
    let mut rules: Vec<(u32, Rule)> = sheet
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
                let (test, sources) = match rule.rule_type.as_str() {
                    "cellIs" => (
                        Test::CellIs(operator(rule.operator.as_deref()?)?),
                        rule.formulas.clone(),
                    ),
                    "expression" | "containsText" | "notContainsText" | "beginsWith"
                    | "endsWith" | "containsBlanks" | "notContainsBlanks" | "containsErrors"
                    | "notContainsErrors" => {
                        (Test::Expression, vec![rule.formulas.first()?.clone()])
                    }
                    "top10" => (
                        Test::Top {
                            rank: attribute("rank").and_then(|rank| rank.parse().ok())?,
                            percent: flag("percent", false),
                            bottom: flag("bottom", false),
                        },
                        Vec::new(),
                    ),
                    "aboveAverage" => (
                        Test::Average {
                            above: flag("aboveAverage", true),
                            equal: flag("equalAverage", false),
                            deviations: attribute("stdDev")
                                .and_then(|deviations| deviations.parse().ok())
                                .unwrap_or(0.0),
                        },
                        Vec::new(),
                    ),
                    "duplicateValues" => (Test::Duplicates { unique: false }, Vec::new()),
                    "uniqueValues" => (Test::Duplicates { unique: true }, Vec::new()),
                    _ => return None,
                };
                let anchor = ranges.first()?.start;
                Some((
                    rule.priority,
                    Rule {
                        anchor,
                        ranges: ranges.clone(),
                        test,
                        formulas: sources
                            .iter()
                            .map(|source| compile(source, anchor))
                            .collect(),
                        format: Arc::new(rule.format.clone()),
                        stop: flag("stopIfTrue", false),
                    },
                ))
            })
        })
        .collect();
    rules.sort_by_key(|(priority, _)| *priority);
    rules.into_iter().map(|(_, rule)| rule).collect()
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
        && args.iter().all(|arg| {
            is_reference(arg) || matches!(arg, Expr::Number(_) | Expr::Text(_) | Expr::Bool(_))
        })
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
            epoch,
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
}
