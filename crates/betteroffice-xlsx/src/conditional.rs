//! Conditional-format effects of the cells a display list draws.

use std::collections::{BTreeMap, HashMap};
use std::ops::Range;
use std::sync::{Arc, OnceLock};

use xlsx_calc::{BinaryOp, EvalContext, Expr, evaluate, parse_formula};
use xlsx_model::styles::DifferentialFormat;
use xlsx_model::value::CellValue;
use xlsx_model::{CellRange, CellRef, SheetId, Workbook};
use xlsx_render::CellFormats;

/// Each sheet's rules by the model epoch they were read at.
#[derive(Default)]
pub(crate) struct Cache {
    epoch: u64,
    sheets: HashMap<SheetId, Arc<[Rule]>>,
}

impl Cache {
    pub(crate) fn rules(&mut self, epoch: u64, model: &Workbook, sheet: SheetId) -> Arc<[Rule]> {
        if self.epoch != epoch {
            self.epoch = epoch;
            self.sheets.clear();
        }
        Arc::clone(
            self.sheets
                .entry(sheet)
                .or_insert_with(|| rules(model, sheet).into()),
        )
    }
}

pub(crate) struct Rule {
    ranges: Vec<CellRange>,
    anchor: CellRef,
    test: Test,
    format: DifferentialFormat,
    stop: bool,
    stats: OnceLock<Stats>,
}

enum Test {
    CellIs {
        op: Op,
        formulas: Vec<String>,
    },
    Expression(String),
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
                let first = rule.formulas.first().cloned();
                let test = match rule.rule_type.as_str() {
                    "cellIs" => Test::CellIs {
                        op: operator(rule.operator.as_deref()?)?,
                        formulas: rule.formulas.clone(),
                    },
                    "expression" | "containsText" | "notContainsText" | "beginsWith"
                    | "endsWith" | "containsBlanks" | "notContainsBlanks" | "containsErrors"
                    | "notContainsErrors" => Test::Expression(first?),
                    "top10" => Test::Top {
                        rank: attribute("rank").and_then(|rank| rank.parse().ok())?,
                        percent: flag("percent", false),
                        bottom: flag("bottom", false),
                    },
                    "aboveAverage" => Test::Average {
                        above: flag("aboveAverage", true),
                        equal: flag("equalAverage", false),
                        deviations: attribute("stdDev")
                            .and_then(|deviations| deviations.parse().ok())
                            .unwrap_or(0.0),
                    },
                    "duplicateValues" => Test::Duplicates { unique: false },
                    "uniqueValues" => Test::Duplicates { unique: true },
                    _ => return None,
                };
                Some((
                    rule.priority,
                    Rule {
                        anchor: ranges.first()?.start,
                        ranges: ranges.clone(),
                        test,
                        format: rule.format.clone(),
                        stop: flag("stopIfTrue", false),
                        stats: OnceLock::new(),
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

/// The effects of `rules` on the cells in the `rows` by `cols` runs. A cell
/// takes each property from the highest-priority matching rule that sets it,
/// and no rule below a matching `stopIfTrue` rule applies.
pub(crate) fn cell_formats(
    model: &Workbook,
    sheet: SheetId,
    rules: &[Rule],
    rows: &[Range<u32>],
    cols: &[Range<u32>],
) -> CellFormats {
    let mut cells: BTreeMap<(u32, u32), (DifferentialFormat, bool)> = BTreeMap::new();
    for rule in rules {
        for range in &rule.ranges {
            for row in runs(rows, range.start.row..range.end.row + 1) {
                for col in runs(cols, range.start.col..range.end.col + 1) {
                    if cells.get(&(row, col)).is_some_and(|(_, stopped)| *stopped)
                        || !rule.matches(model, sheet, CellRef::new(row, col))
                    {
                        continue;
                    }
                    let (format, stopped) = cells.entry((row, col)).or_default();
                    format.fill = format.fill.take().or_else(|| rule.format.fill.clone());
                    format.font_color = format
                        .font_color
                        .take()
                        .or_else(|| rule.format.font_color.clone());
                    format.bold = format.bold.or(rule.format.bold);
                    format.italic = format.italic.or(rule.format.italic);
                    *stopped = rule.stop;
                }
            }
        }
    }
    cells
        .into_iter()
        .map(|(at, (format, _))| (at, format))
        .collect()
}

fn runs<'a>(runs: &'a [Range<u32>], within: Range<u32>) -> impl Iterator<Item = u32> + 'a {
    runs.iter()
        .flat_map(move |run| run.start.max(within.start)..run.end.min(within.end))
}

impl Rule {
    fn matches(&self, model: &Workbook, sheet: SheetId, at: CellRef) -> bool {
        let value = model
            .sheet(sheet)
            .and_then(|owner| owner.cell(at))
            .map_or(CellValue::Empty, |cell| cell.value.clone());
        let mut ctx = EvalContext::new(model, sheet);
        ctx.cell = Some(at);
        let formula = |index: usize| {
            let source = self.formulas().get(index)?;
            let shifted = xlsx_parse::translate_formula(source, self.anchor, at).ok()?;
            Some(evaluate(&parse_formula(&shifted).ok()?, &ctx))
        };
        let compare = |op: BinaryOp, lhs: &CellValue, rhs: &CellValue| {
            let expr = Expr::Binary {
                op,
                lhs: Box::new(Expr::Literal(lhs.clone())),
                rhs: Box::new(Expr::Literal(rhs.clone())),
            };
            evaluate(&expr, &ctx) == CellValue::Bool { value: true }
        };
        match &self.test {
            Test::CellIs { op, .. } => {
                if matches!(value, CellValue::Error { .. }) {
                    return false;
                }
                match *op {
                    Op::Compare(op) => formula(0).is_some_and(|rhs| compare(op, &value, &rhs)),
                    Op::Between(inside) => {
                        let (Some(low), Some(high)) = (formula(0), formula(1)) else {
                            return false;
                        };
                        let within = |low: &CellValue, high: &CellValue| {
                            compare(BinaryOp::Ge, &value, low)
                                && compare(BinaryOp::Le, &value, high)
                        };
                        (within(&low, &high) || within(&high, &low)) == inside
                    }
                }
            }
            Test::Expression(_) => match formula(0) {
                Some(CellValue::Bool { value }) => value,
                Some(CellValue::Number { value }) => value != 0.0,
                _ => false,
            },
            Test::Top { bottom, .. } => {
                let (CellValue::Number { value }, Stats::Threshold(Some(threshold))) =
                    (value, self.stats(model, sheet))
                else {
                    return false;
                };
                if *bottom {
                    value <= *threshold
                } else {
                    value >= *threshold
                }
            }
            Test::Average {
                above,
                equal,
                deviations,
            } => {
                let (CellValue::Number { value }, Stats::Average { mean, deviation }) =
                    (value, self.stats(model, sheet))
                else {
                    return false;
                };
                let bound = if *above {
                    mean + deviations * deviation
                } else {
                    mean - deviations * deviation
                };
                match (above, equal) {
                    (true, true) => value >= bound,
                    (true, false) => value > bound,
                    (false, true) => value <= bound,
                    (false, false) => value < bound,
                }
            }
            Test::Duplicates { unique } => {
                let (Some(key), Stats::Counts(counts)) = (key(&value), self.stats(model, sheet))
                else {
                    return false;
                };
                let count = counts.get(&key).copied().unwrap_or(0);
                if *unique { count == 1 } else { count > 1 }
            }
        }
    }

    fn formulas(&self) -> &[String] {
        match &self.test {
            Test::CellIs { formulas, .. } => formulas,
            Test::Expression(formula) => std::slice::from_ref(formula),
            _ => &[],
        }
    }

    /// Statistics over the rule's whole range, computed once.
    fn stats(&self, model: &Workbook, sheet: SheetId) -> &Stats {
        self.stats.get_or_init(|| {
            let cells = || {
                model.sheet(sheet).into_iter().flat_map(|owner| {
                    self.ranges
                        .iter()
                        .flat_map(|range| owner.cells_in_range(*range))
                        .map(|(_, cell)| &cell.value)
                })
            };
            let numbers = || {
                cells().filter_map(|value| match value {
                    CellValue::Number { value } => Some(*value),
                    _ => None,
                })
            };
            match self.test {
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
                    for key in cells().filter_map(key) {
                        *counts.entry(key).or_insert(0) += 1;
                    }
                    Stats::Counts(counts)
                }
            }
        })
    }
}

fn key(value: &CellValue) -> Option<Key> {
    Some(match value {
        CellValue::Number { value } => Key::Number((value + 0.0).to_bits()),
        CellValue::Text { value } => Key::Text(value.to_lowercase()),
        CellValue::Bool { value } => Key::Bool(*value),
        _ => return None,
    })
}
