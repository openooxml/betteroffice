//! Conditional-format effects of the cells a display list draws.

use std::collections::HashMap;
use std::ops::Range;
use std::sync::Arc;

use xlsx_calc::eval::MAX_EVALUATION_CELL_VISITS;
use xlsx_calc::{BinaryOp, EvalContext, Expr, evaluate, parse_formula};
use xlsx_model::styles::DifferentialFormat;
use xlsx_model::value::CellValue;
use xlsx_model::{CellRange, CellRef, MAX_COLS, MAX_ROWS, SheetId, Workbook};
use xlsx_render::CellFormats;

/// Rule checks plus cell visits one frame on screen spends deciding cells.
pub const CONDITIONAL_FRAME_BUDGET: u64 = 200_000;
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

/// How much evaluation a display list spends on conditional formats.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Budget {
    /// One frame's budget; later frames decide the rest.
    Frame,
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
    /// Decided cells held.
    pub cached_cells: usize,
    /// Model epoch the cache holds.
    pub epoch: u64,
}

/// Each sheet's rules, aggregates and decided cells for one model epoch, and
/// shifted formulas parsed in any.
pub(crate) struct Cache {
    epoch: Option<u64>,
    rules: HashMap<SheetId, Arc<[Rule]>>,
    aggregates: HashMap<(SheetId, usize), Arc<Aggregate>>,
    results: HashMap<(SheetId, u32, u32), Option<Arc<DifferentialFormat>>>,
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
    /// Paths in `base` of the calls whose value is the same for every cell.
    constants: Vec<Vec<usize>>,
}

/// What one rule reads from its whole range, once per epoch.
#[derive(Default)]
struct Aggregate {
    stats: Option<Stats>,
    /// Each formula's constant calls, in `Formula::constants` order.
    constants: Vec<Vec<CellValue>>,
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
            epoch: self.epoch.unwrap_or_default(),
            ..self.stats
        }
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
        budget: Budget,
    ) -> (CellFormats, u32) {
        if self.epoch != Some(epoch) {
            self.epoch = Some(epoch);
            self.rules.clear();
            self.aggregates.clear();
            self.results.clear();
        }
        let rules = Arc::clone(
            self.rules
                .entry(sheet)
                .or_insert_with(|| rules(model, sheet).into()),
        );
        let mut formats = CellFormats::new();
        self.stats.pending = 0;
        if rules.is_empty() {
            return (formats, 0);
        }
        let limit = match budget {
            Budget::Frame => self.budget,
            Budget::Complete => u64::MAX,
        };
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
                let result = match self.results.get(&(sheet, row, col)).cloned() {
                    Some(result) => {
                        self.stats.hits += 1;
                        result
                    }
                    None if spent >= limit => {
                        pending += 1;
                        continue;
                    }
                    None => {
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
                };
                if let Some(format) = result {
                    formats.insert((row, col), format);
                }
            }
        }
        self.stats.pending = pending;
        (formats, pending)
    }

    /// Holds a decided cell, first dropping the cells outside the frame when
    /// full; false when the frame alone fills the cap.
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
        self.results.insert((sheet, at.row, at.col), result);
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

    /// The value of a rule's formula at `at`, its constant calls read from
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
        let expr = if formula.relative {
            let shifted = xlsx_parse::translate_formula(&formula.source, rule.anchor, at).ok()?;
            self.parse(shifted)?
        } else {
            Arc::clone(formula.base.as_ref()?)
        };
        let values = aggregate.constants.get(slot).map_or(&[][..], Vec::as_slice);
        if values.is_empty() {
            return Some(evaluate(&expr, ctx));
        }
        let mut expr = (*expr).clone();
        for (path, value) in formula.constants.iter().zip(values) {
            if let Some(call) = node_mut(&mut expr, path)
                && matches!(call, Expr::FuncCall { .. })
            {
                *call = Expr::Literal(value.clone());
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
            aggregate.constants.push(values);
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
    let mut constants = Vec::new();
    if let (Some(base), Some(shifted)) = (
        &base,
        shifted.as_deref().and_then(|text| parse_formula(text).ok()),
    ) {
        constant_calls(base, &shifted, &mut Vec::new(), &mut constants);
    }
    Formula {
        source: source.to_owned(),
        base,
        relative,
        constants,
    }
}

/// Collects the paths of the calls `base` and `shifted` share that aggregate
/// absolute references alone.
fn constant_calls(base: &Expr, shifted: &Expr, path: &mut Vec<usize>, out: &mut Vec<Vec<usize>>) {
    if base == shifted && aggregates_absolute(base) {
        out.push(path.clone());
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
        constant_calls(left, right, path, out);
        path.pop();
    }
}

fn aggregates_absolute(expr: &Expr) -> bool {
    let Expr::FuncCall { name, args, .. } = expr else {
        return false;
    };
    let upper = name.to_ascii_uppercase();
    let name = upper.strip_prefix("_XLFN.").unwrap_or(&upper);
    let reference = |arg: &Expr| {
        matches!(
            arg,
            Expr::Ref { .. }
                | Expr::Range { .. }
                | Expr::ColumnRange { .. }
                | Expr::RowRange { .. }
        )
    };
    AGGREGATES.contains(&name)
        && args.iter().any(reference)
        && args.iter().all(|arg| {
            reference(arg) || matches!(arg, Expr::Number(_) | Expr::Text(_) | Expr::Bool(_))
        })
}

/// The operands `constant_calls` descends into; a range join's ends stay
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
