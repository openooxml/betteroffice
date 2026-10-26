use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use xlsx_model::{CellRange, CellRef, CellValue, SheetId};

use crate::array::{Value, evaluate_array, step_units};
use crate::eval::{Area, EvalContext, as_area, num};
use crate::parser::{BinaryOp, Expr, UnaryOp};

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
enum Key {
    Number(u64),
    Text(String),
    Bool(bool),
}

impl Key {
    fn of(value: &CellValue) -> Option<Self> {
        match value {
            CellValue::Number { value } if value.is_finite() => {
                Some(Self::Number(if *value == 0.0 {
                    0
                } else {
                    value.to_bits()
                }))
            }
            CellValue::Text { value } => Some(Self::Text(value.to_lowercase())),
            CellValue::Bool { value } => Some(Self::Bool(*value)),
            _ => None,
        }
    }

    fn matches_blank(&self) -> bool {
        match self {
            Self::Number(bits) => *bits == 0,
            Self::Text(value) => value.is_empty(),
            Self::Bool(value) => !value,
        }
    }
}

type RangeKey = (SheetId, u32, u32, usize, usize);

#[derive(Default)]
pub(crate) struct CriteriaCache {
    ranges: HashMap<RangeKey, Rc<IndexedRange>>,
    insertions: Vec<RangeKey>,
    unusable: HashSet<RangeKey>,
    non_numeric: HashSet<RangeKey>,
}

impl CriteriaCache {
    #[cfg(test)]
    pub(super) fn holds_ranges(&self) -> bool {
        !self.ranges.is_empty()
    }

    /// forget every index and refusal: a spill round may have written new
    /// values into ranges an earlier round indexed.
    pub(crate) fn clear(&mut self) {
        self.ranges.clear();
        self.insertions.clear();
        self.unusable.clear();
        self.non_numeric.clear();
    }

    pub(crate) fn checkpoint(&self) -> usize {
        self.insertions.len()
    }

    pub(crate) fn restore(&mut self, checkpoint: usize) {
        while self.insertions.len() > checkpoint {
            if let Some(key) = self.insertions.pop() {
                self.ranges.remove(&key);
            }
        }
    }
}

struct IndexedRange {
    postings: HashMap<Key, Posting>,
    row_keys: Vec<u32>,
    read_cells: usize,
    blank_tail: usize,
    blanks: Vec<usize>,
    numbers: Option<Vec<f64>>,
}

struct Posting {
    id: u32,
    rows: Vec<usize>,
}

struct Plan<'a> {
    criteria: Vec<(Area, &'a Expr)>,
    values: Option<Area>,
}

fn plain_area(expr: &Expr, ctx: &EvalContext<'_>) -> Option<Area> {
    if !matches!(
        expr,
        Expr::Ref { .. } | Expr::Range { .. } | Expr::ColumnRange { .. } | Expr::RowRange { .. }
    ) {
        return None;
    }
    let mut area = as_area(expr, ctx)?;
    if area.rows >= xlsx_model::MAX_ROWS as usize {
        area.rows = (ctx.provider.used_rows(area.sheet) as usize)
            .saturating_sub(area.start.row as usize)
            .max(1);
    }
    if area.cols >= xlsx_model::MAX_COLS as usize {
        area.cols = (ctx.provider.used_cols(area.sheet) as usize)
            .saturating_sub(area.start.col as usize)
            .max(1);
    }
    (area.rows == 1 || area.cols == 1).then_some(area)
}

fn factors<'a>(expr: &'a Expr, ctx: &EvalContext<'_>, plan: &mut Plan<'a>) -> Option<()> {
    if let Expr::Binary {
        op: BinaryOp::Mul,
        lhs,
        rhs,
    } = expr
    {
        factors(lhs, ctx, plan)?;
        return factors(rhs, ctx, plan);
    }
    let mut comparison = expr;
    if let Expr::Unary {
        op: UnaryOp::Neg,
        expr: inner,
    } = comparison
        && let Expr::Unary {
            op: UnaryOp::Neg,
            expr: inner,
        } = inner.as_ref()
    {
        comparison = inner;
    }
    if let Expr::Binary {
        op: BinaryOp::Eq,
        lhs,
        rhs,
    } = comparison
    {
        let left = plain_area(lhs, ctx);
        let right = plain_area(rhs, ctx);
        let (area, key) = match (left, right) {
            (Some(left), Some(right)) if left.rows * left.cols == 1 => (right, lhs.as_ref()),
            (Some(left), _) => (left, rhs.as_ref()),
            (_, Some(right)) => (right, lhs.as_ref()),
            _ => return None,
        };
        if !constant_key(key, ctx) {
            return None;
        }
        plan.criteria.push((area, key));
        return Some(());
    }
    let area = plain_area(expr, ctx)?;
    if plan.values.replace(area).is_some() {
        return None;
    }
    Some(())
}

/// a key the index can read before it touches a range: a literal or one
/// cell. anything computed takes the existing path, which meters it.
fn constant_key(expr: &Expr, ctx: &EvalContext<'_>) -> bool {
    match expr {
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) => true,
        Expr::Unary {
            op: UnaryOp::Neg,
            expr,
        } => matches!(expr.as_ref(), Expr::Number(_)),
        _ => plain_area(expr, ctx).is_some_and(|area| (area.rows, area.cols) == (1, 1)),
    }
}

fn range_key(area: Area) -> RangeKey {
    (
        area.sheet,
        area.start.row,
        area.start.col,
        area.rows,
        area.cols,
    )
}

/// rows of `area` inside its sheet's used rows; the rest are a blank tail
fn read_rows(area: Area, ctx: &EvalContext<'_>) -> usize {
    let used = ctx.provider.used_rows(area.sheet) as usize;
    if used == 0 {
        area.rows
    } else {
        used.saturating_sub(area.start.row as usize).min(area.rows)
    }
}

fn indexed_range(area: Area, ctx: &EvalContext<'_>) -> Option<Rc<IndexedRange>> {
    let key = range_key(area);
    if ctx.criteria_cache.borrow().unusable.contains(&key) {
        return None;
    }
    if let Some(range) = ctx.criteria_cache.borrow().ranges.get(&key) {
        return Some(Rc::clone(range));
    }
    let bounds = CellRange::new(
        area.start,
        CellRef::new(
            area.start.row + area.rows as u32 - 1,
            area.start.col + area.cols as u32 - 1,
        ),
    );
    if ctx
        .spill_index
        .is_some_and(|graph| graph.spill_sources(area.sheet, bounds).next().is_some())
    {
        ctx.criteria_cache.borrow_mut().unusable.insert(key);
        return None;
    }
    let read_rows = read_rows(area, ctx);
    let cost = (read_rows as u64).checked_mul(area.cols as u64)?;
    if !ctx.can_consume(cost) {
        return None;
    }
    if !ctx.consume_cells(cost) {
        return None;
    }
    #[cfg(test)]
    ctx.note_criteria_index_reads(cost);
    let len = read_rows.checked_mul(area.cols)?;
    let mut range = IndexedRange {
        postings: HashMap::new(),
        row_keys: Vec::with_capacity(len),
        read_cells: cost as usize,
        blank_tail: (area.rows - read_rows).checked_mul(area.cols)?,
        blanks: Vec::new(),
        numbers: Some(Vec::with_capacity(len)),
    };
    for offset in 0..len {
        let (row, col) = (offset / area.cols, offset % area.cols);
        let at = CellRef::new(area.start.row + row as u32, area.start.col + col as u32);
        if ctx.provider.formula(area.sheet, at).is_some() {
            ctx.criteria_cache.borrow_mut().unusable.insert(key);
            return None;
        }
        let value = area.get_unmetered_ref(ctx, row, col);
        if matches!(value.as_ref(), CellValue::Empty) {
            range.blanks.push(offset);
            range.row_keys.push(0);
        } else {
            let Some(value_key) = Key::of(&value) else {
                ctx.criteria_cache.borrow_mut().unusable.insert(key);
                return None;
            };
            let next_id = u32::try_from(range.postings.len() + 1).ok()?;
            let posting = range.postings.entry(value_key).or_insert_with(|| Posting {
                id: next_id,
                rows: Vec::new(),
            });
            posting.rows.push(offset);
            range.row_keys.push(posting.id);
        }
        match value.as_ref() {
            CellValue::Number { value } => {
                if let Some(numbers) = &mut range.numbers {
                    numbers.push(*value);
                }
            }
            CellValue::Empty => {
                if let Some(numbers) = &mut range.numbers {
                    numbers.push(0.0);
                }
            }
            _ => {
                if range.numbers.take().is_some() {
                    ctx.criteria_cache.borrow_mut().non_numeric.insert(key);
                }
            }
        }
    }
    let range = Rc::new(range);
    let mut cache = ctx.criteria_cache.borrow_mut();
    cache.ranges.insert(key, Rc::clone(&range));
    cache.insertions.push(key);
    Some(range)
}

pub(super) fn sumproduct(expr: &Expr, ctx: &EvalContext<'_>) -> Option<CellValue> {
    if ctx.is_settling() {
        return None;
    }
    #[cfg(test)]
    if !ctx.criteria_index_enabled() {
        return None;
    }
    let mut plan = Plan {
        criteria: Vec::new(),
        values: None,
    };
    factors(expr, ctx, &mut plan)?;
    let values = plan.values?;
    let len = values.rows.checked_mul(values.cols)?;
    if plan.criteria.is_empty()
        || len <= 1
        || len > crate::array::MAX_ARRAY_CELLS
        || plan
            .criteria
            .iter()
            .any(|(area, _)| (area.rows, area.cols) != (values.rows, values.cols))
    {
        return None;
    }
    {
        let cache = ctx.criteria_cache.borrow();
        let key = range_key(values);
        if cache.unusable.contains(&key)
            || cache.non_numeric.contains(&key)
            || plan
                .criteria
                .iter()
                .any(|(area, _)| cache.unusable.contains(&range_key(*area)))
        {
            return None;
        }
    }
    // the index replaces only a SUMPRODUCT the streamed path would run now:
    // it needs the same headroom and never charges more
    let (streamed_cost, text) = super::math::criteria_streamed_cost(expr, ctx)?;
    if !ctx.can_consume(streamed_cost) || !ctx.can_consume_text(text) {
        return None;
    }
    ctx.try_indexed(|| probe(&plan.criteria, values, streamed_cost, ctx))
}

fn probe(
    criteria: &[(Area, &Expr)],
    values: Area,
    cost_limit: u64,
    ctx: &EvalContext<'_>,
) -> Option<CellValue> {
    let remaining = ctx.remaining_cells();
    // keys are literals or single cells, read before any range so a key the
    // index cannot use never leaves a range read and rolled back
    let mut keys = Vec::with_capacity(criteria.len());
    for (_, expr) in criteria {
        let value = {
            let _scope = ctx.array_scope();
            if let Some(area) = plain_area(expr, ctx) {
                if (area.rows, area.cols) != (1, 1) {
                    return None;
                }
                area.get_ref(ctx, 0, 0).ok()?.into_owned()
            } else {
                match evaluate_array(expr, ctx) {
                    Value::Scalar(value) => value,
                    _ => return None,
                }
            }
        };
        keys.push(Key::of(&value)?);
    }
    let values_cached = ctx
        .criteria_cache
        .borrow()
        .ranges
        .contains_key(&range_key(values));
    // refuse before reading anything when the ranges together are
    // unaffordable, so a fallback never reads ranges it does not charge
    let mut unbuilt: Vec<RangeKey> = Vec::with_capacity(criteria.len() + 1);
    let mut build_cost = 0u64;
    for area in criteria.iter().map(|(area, _)| *area).chain([values]) {
        let key = range_key(area);
        if unbuilt.contains(&key) || ctx.criteria_cache.borrow().ranges.contains_key(&key) {
            continue;
        }
        let read_rows = read_rows(area, ctx) as u64;
        build_cost = build_cost.checked_add(read_rows.checked_mul(area.cols as u64)?)?;
        unbuilt.push(key);
    }
    if !ctx.can_consume(build_cost) {
        return None;
    }
    let mut ranges = Vec::with_capacity(criteria.len());
    for (area, _) in criteria {
        ranges.push(indexed_range(*area, ctx)?);
    }
    let values = indexed_range(values, ctx)?;
    let numbers = values.numbers.as_ref()?;
    // A newly built value range has already paid for every cell it read.
    let prepaid_values = !values_cached && values.read_cells == numbers.len();
    let mut matches = Vec::with_capacity(criteria.len());
    for (key, range) in keys.into_iter().zip(&ranges) {
        let posting = range.postings.get(&key);
        let id = posting.map(|posting| posting.id);
        let posting = posting.map_or(&[][..], |posting| posting.rows.as_slice());
        let blanks = if key.matches_blank() {
            &range.blanks[..]
        } else {
            &[]
        };
        matches.push(Matches {
            posting,
            blanks,
            row_keys: &range.row_keys,
            id,
            matches_blank: key.matches_blank(),
            blank_tail: if key.matches_blank() {
                range.blank_tail
            } else {
                0
            },
        });
    }
    matches.sort_by_key(Matches::len);
    let first = matches.first()?;
    let steps = (first.len() as u64).checked_mul(matches.len() as u64)?;
    let value_cost = if prepaid_values {
        0
    } else {
        first.len() as u64
    };
    let cost = step_units(steps).checked_add(value_cost)?;
    let spent = remaining.checked_sub(ctx.remaining_cells())?;
    if spent.checked_add(cost)? > cost_limit {
        return None;
    }
    if !ctx.consume_cells(step_units(steps)) {
        return None;
    }
    let mut posting = first.posting.iter().peekable();
    let mut blanks = first.blanks.iter().peekable();
    let read = std::iter::from_fn(|| match (posting.peek(), blanks.peek()) {
        (Some(a), Some(b)) if a < b => posting.next(),
        (Some(_), Some(_)) | (None, Some(_)) => blanks.next(),
        (Some(_), None) => posting.next(),
        (None, None) => None,
    });
    // rows past the first range's used rows are blank there and come last
    let tail_start = first.row_keys.len();
    let tail = tail_start..tail_start + first.blank_tail;
    let mut total = 0.0;
    for row in read.copied().chain(tail) {
        if matches[1..].iter().all(|other| other.contains(row)) {
            if !prepaid_values && !ctx.consume_cells(1) {
                return None;
            }
            total += numbers.get(row).copied().unwrap_or(0.0);
        }
    }
    Some(num(total))
}

struct Matches<'a> {
    posting: &'a [usize],
    blanks: &'a [usize],
    row_keys: &'a [u32],
    id: Option<u32>,
    matches_blank: bool,
    blank_tail: usize,
}

impl Matches<'_> {
    fn len(&self) -> usize {
        self.posting.len() + self.blanks.len() + self.blank_tail
    }

    fn contains(&self, row: usize) -> bool {
        // rows past this range's used rows are blank
        let id = self.row_keys.get(row).copied().unwrap_or(0);
        self.id == Some(id) || (self.matches_blank && id == 0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::eval::{EvaluationBudget, cmp_values};
    use xlsx_model::{Cell, Sheet, Workbook};

    #[test]
    fn canonical_keys_match_scalar_equality() {
        let mut values = vec![
            CellValue::Empty,
            CellValue::Bool { value: false },
            CellValue::Bool { value: true },
        ];
        for value in [-5.0, -0.0, 0.0, 5.0, f64::MIN_POSITIVE, f64::MAX] {
            values.push(CellValue::Number { value });
        }
        for value in [
            "", "5", "Apple", "APPLE", "İ", "i\u{307}", "Σ", "σ", "ς", "[a]", "[A]",
        ] {
            values.push(CellValue::Text {
                value: value.into(),
            });
        }
        for a in &values {
            for b in &values {
                let indexed = match (Key::of(a), Key::of(b)) {
                    (Some(a), Some(b)) => a == b,
                    (Some(key), None) | (None, Some(key)) => key.matches_blank(),
                    (None, None) => true,
                };
                assert_eq!(indexed, cmp_values(a, b).is_eq(), "{a:?} {b:?}");
            }
        }
    }

    fn small_workbook() -> Workbook {
        let mut sheet = Sheet::new("Data");
        for row in 0..3 {
            for col in 0..2 {
                sheet.set_cell(
                    CellRef::new(row, col),
                    Cell {
                        value: CellValue::Number { value: 5.0 },
                        ..Cell::default()
                    },
                );
            }
        }
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        wb
    }

    #[test]
    fn ranges_with_different_used_rows_match_the_existing_path() {
        // Keys uses 1,000 rows and Vals 1,600, so Keys!A1:A1600 ends in a blank tail
        let mut keys = Sheet::new("Keys");
        for row in 0..1_000 {
            if row % 100 != 50 {
                keys.set_cell(
                    CellRef::new(row, 0),
                    Cell {
                        value: num(row as f64),
                        ..Cell::default()
                    },
                );
            }
        }
        let mut vals = Sheet::new("Vals");
        for row in 0..1_600 {
            let key = if row % 200 == 0 {
                "x".to_string()
            } else {
                format!("k{row}")
            };
            for (col, value) in [
                (0, num((row % 7) as f64)),
                (1, CellValue::Text { value: key }),
            ] {
                vals.set_cell(
                    CellRef::new(row, col),
                    Cell {
                        value,
                        ..Cell::default()
                    },
                );
            }
        }
        let mut wb = Workbook::default();
        wb.sheets.extend([keys, vals]);
        for argument in [
            "(Vals!B1:B1600=\"x\")*Keys!A1:A1600",
            "(Keys!A1:A1600=\"\")*(Vals!B1:B1600=\"x\")*Vals!A1:A1600",
            "(Keys!A1:A1600=\"\")*Vals!A1:A1600",
        ] {
            let ctx = EvalContext::new(&wb, SheetId(0));
            let indexed = sumproduct(&crate::parse_formula(argument).unwrap(), &ctx);
            let fallback = crate::parse_formula(&format!("SUMPRODUCT(({argument})+0)")).unwrap();
            let expected = crate::eval::evaluate(&fallback, &EvalContext::new(&wb, SheetId(0)));
            if argument.starts_with("(Vals") {
                assert_eq!(indexed, Some(num(2_000.0)));
            }
            if let Some(indexed) = indexed {
                assert_eq!(indexed, expected, "{argument}");
            }
        }
    }

    #[test]
    fn criteria_index_unread_million_row_tail_uses_only_read_entries() {
        let mut wb = small_workbook();
        wb.sheets[0].set_cell(
            CellRef::new(9, 0),
            Cell {
                value: num(5.0),
                ..Cell::default()
            },
        );
        let ctx = EvalContext::new(&wb, SheetId(0));
        let area = Area {
            sheet: SheetId(0),
            start: CellRef::new(0, 0),
            rows: 1_000_000,
            cols: 1,
        };
        // the unread rows are one count, not an entry per row
        let range = indexed_range(area, &ctx).unwrap();
        assert_eq!(range.blank_tail, 999_990);
        assert_eq!(range.row_keys.len(), 10);
        assert_eq!(range.numbers.as_ref().unwrap().len(), 10);
        assert_eq!(range.read_cells, 10);
    }

    #[test]
    fn criteria_ranges_are_charged_once_per_pass() {
        let wb = small_workbook();
        let budget = Rc::new(EvaluationBudget::new(100));
        let area = Area {
            sheet: SheetId(0),
            start: CellRef::new(0, 0),
            rows: 3,
            cols: 1,
        };
        let first = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        assert!(indexed_range(area, &first).is_some());
        assert_eq!(budget.remaining(), 97);
        let second = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        assert!(indexed_range(area, &second).is_some());
        assert_eq!(budget.remaining(), 97);
    }

    #[test]
    fn refused_plans_restore_budget_and_cache_before_fallback() {
        let wb = small_workbook();
        for limit in [2, 100] {
            for source in [
                "(A1:A3=C1)*B1:B3",
                "(A1:A3=1/0)*B1:B3",
                "(A1:A3={5;5;5})*B1:B3",
            ] {
                let budget = Rc::new(EvaluationBudget::new(limit));
                let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
                let expr = crate::parse_formula(source).unwrap();
                assert!(sumproduct(&expr, &ctx).is_none(), "{source}");
                assert_eq!(budget.remaining(), limit);
                assert!(!budget.refused());
                assert!(!ctx.exhausted());
                assert!(ctx.criteria_cache.borrow().ranges.is_empty());
            }
        }
    }

    #[test]
    fn single_cell_shapes_keep_the_existing_path() {
        let wb = small_workbook();
        let budget = Rc::new(EvaluationBudget::new(100));
        let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        for source in ["(A1:A1=5)*B1:B1", "(A1=5)*(A2=5)*B3"] {
            let expr = crate::parse_formula(source).unwrap();
            assert!(sumproduct(&expr, &ctx).is_none(), "{source}");
        }
        assert_eq!(budget.remaining(), 100);
        assert!(ctx.criteria_cache.borrow().ranges.is_empty());
    }

    #[test]
    fn expensive_probe_falls_back_without_caching_a_refusal() {
        let wb = small_workbook();
        let budget = Rc::new(EvaluationBudget::new(1_000));
        let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        let expensive = crate::parse_formula("(A1:A1000=0)*B1:B1000").unwrap();
        assert!(sumproduct(&expensive, &ctx).is_none());
        assert_eq!(budget.remaining(), 1_000);
        assert!(!budget.refused());
        assert!(ctx.criteria_cache.borrow().ranges.is_empty());
        assert!(ctx.criteria_cache.borrow().unusable.is_empty());
        assert!(ctx.criteria_cache.borrow().non_numeric.is_empty());
        let selective = crate::parse_formula("(A1:A1000=6)*B1:B1000").unwrap();
        assert_eq!(sumproduct(&selective, &ctx), Some(num(0.0)));
    }

    struct CountingProvider {
        workbook: Workbook,
        reads: std::cell::Cell<usize>,
    }

    impl xlsx_model::CellProvider for CountingProvider {
        fn value(&self, sheet: SheetId, at: CellRef) -> CellValue {
            self.reads.set(self.reads.get() + 1);
            self.workbook.value(sheet, at)
        }

        fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str> {
            self.reads.set(self.reads.get() + 1);
            self.workbook.formula(sheet, at)
        }

        fn sheet_id(&self, name: &str) -> Option<SheetId> {
            self.workbook.sheet_id(name)
        }
    }

    #[test]
    fn unindexable_ranges_survive_rollback_without_repeated_reads() {
        use xlsx_model::ErrorValue;

        for cell in [
            Cell {
                formula: Some("5".into()),
                value: num(5.0),
                ..Cell::default()
            },
            Cell {
                value: CellValue::Error {
                    value: ErrorValue::Div0,
                },
                ..Cell::default()
            },
            Cell {
                value: CellValue::Number {
                    value: f64::INFINITY,
                },
                ..Cell::default()
            },
            Cell {
                value: CellValue::Text {
                    value: "bad".into(),
                },
                ..Cell::default()
            },
            Cell {
                value: CellValue::Bool { value: false },
                ..Cell::default()
            },
        ] {
            let mut wb = small_workbook();
            wb.sheets[0].set_cell(CellRef::new(2, 1), cell);
            let provider = CountingProvider {
                workbook: wb,
                reads: std::cell::Cell::new(0),
            };
            let budget = Rc::new(EvaluationBudget::new(100));
            let expr = crate::parse_formula("(A1:A3=5)*B1:B3").unwrap();
            let first = EvalContext::with_budget(&provider, SheetId(0), Rc::clone(&budget));
            assert!(sumproduct(&expr, &first).is_none());
            assert_eq!(budget.remaining(), 100);
            assert!(first.criteria_cache.borrow().ranges.is_empty());
            let reads = provider.reads.get();
            assert!(reads > 0);
            let second = EvalContext::with_budget(&provider, SheetId(0), Rc::clone(&budget));
            assert!(sumproduct(&expr, &second).is_none());
            assert_eq!(provider.reads.get(), reads);
            assert_eq!(budget.remaining(), 100);
            assert!(!budget.refused());
        }
    }

    #[test]
    fn non_numeric_value_refusal_still_allows_criteria_use() {
        let mut wb = small_workbook();
        wb.sheets[0].set_cell(
            CellRef::new(2, 1),
            Cell {
                value: CellValue::Text { value: "x".into() },
                ..Cell::default()
            },
        );
        let budget = Rc::new(EvaluationBudget::new(100));
        let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        let refused = crate::parse_formula("(A1:A3=5)*B1:B3").unwrap();
        assert!(sumproduct(&refused, &ctx).is_none());
        assert_eq!(budget.remaining(), 100);
        let accepted = crate::parse_formula("(B1:B3=\"x\")*A1:A3").unwrap();
        assert_eq!(sumproduct(&accepted, &ctx), Some(num(5.0)));
    }

    #[test]
    fn spill_refusal_is_remembered_across_rollback() {
        let mut wb = small_workbook();
        let start = CellRef::new(0, 1);
        wb.sheets[0].set_cell(
            start,
            Cell {
                formula: Some("SEQUENCE(3)".into()),
                ..Cell::default()
            },
        );
        wb.sheets[0].set_array_formula(start, CellRange::new(start, CellRef::new(2, 1)));
        let graph = crate::graph::DepGraph::build(&wb);
        let budget = Rc::new(EvaluationBudget::new(100));
        let mut ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        ctx.spill_index = Some(&graph);
        let expr = crate::parse_formula("(A1:A3=5)*B1:B3").unwrap();
        assert!(sumproduct(&expr, &ctx).is_none());
        assert_eq!(ctx.criteria_cache.borrow().unusable.len(), 1);
        assert!(ctx.criteria_cache.borrow().ranges.is_empty());
        assert_eq!(budget.remaining(), 100);
        assert!(sumproduct(&expr, &ctx).is_none());
        assert_eq!(budget.remaining(), 100);
    }

    #[test]
    fn budget_refusal_does_not_poison_range_after_refund() {
        let wb = small_workbook();
        let budget = Rc::new(EvaluationBudget::new(100));
        let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
        let area = Area {
            sheet: SheetId(0),
            start: CellRef::new(0, 0),
            rows: 3,
            cols: 1,
        };
        assert!(
            ctx.try_indexed(|| {
                assert!(ctx.consume_cells(98));
                indexed_range(area, &ctx)
            })
            .is_none()
        );
        assert_eq!(budget.remaining(), 100);
        assert!(ctx.criteria_cache.borrow().unusable.is_empty());
        assert!(indexed_range(area, &ctx).is_some());
        assert_eq!(budget.remaining(), 97);
    }

    #[test]
    fn literal_and_single_cell_keys_use_the_index_and_computed_keys_do_not() {
        let wb = small_workbook();
        for source in ["(B1=A1:A3)*B1:B3", "(A1:A3=5)*B1:B3", "--(A1:A3=B1)*B1:B3"] {
            let ctx = EvalContext::new(&wb, SheetId(0));
            let expr = crate::parse_formula(source).unwrap();
            assert_eq!(sumproduct(&expr, &ctx), Some(num(15.0)), "{source}");
        }
        for source in [
            "(A1:A3=2+3)*B1:B3",
            "--(A1:A3=ABS(-5))*B1:B3",
            "(A1:A3=+B1)*B1:B3",
            "(A1:A3=B1:B3)*B1:B3",
        ] {
            let ctx = EvalContext::new(&wb, SheetId(0));
            let expr = crate::parse_formula(source).unwrap();
            assert_eq!(sumproduct(&expr, &ctx), None, "{source}");
            assert!(ctx.criteria_cache.borrow().ranges.is_empty());
        }
    }

    #[test]
    fn indexed_probes_charge_builds_once_and_stay_far_below_recalc_limit() {
        let mut sheet = Sheet::new("Data");
        for row in 0..20_000 {
            for col in 0..4 {
                let value = match col {
                    0 => CellValue::Text {
                        value: format!("a{}", row % 100),
                    },
                    1 => CellValue::Text {
                        value: format!("b{}", row % 37),
                    },
                    2 => CellValue::Number {
                        value: (row % 13) as f64,
                    },
                    _ => CellValue::Number {
                        value: (row % 17) as f64 / 8.0,
                    },
                };
                sheet.set_cell(
                    CellRef::new(row, col),
                    Cell {
                        value,
                        ..Cell::default()
                    },
                );
            }
        }
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        let budget = Rc::new(EvaluationBudget::new(10_000_000));
        for key in 0..1_500 {
            let expr = crate::parse_formula(&format!(
                "(A1:A20000=\"a{}\")*(B1:B20000=\"b{}\")*(C1:C20000={})*D1:D20000",
                key % 100,
                key % 37,
                key % 13
            ))
            .unwrap();
            let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
            assert!(sumproduct(&expr, &ctx).is_some());
            assert_eq!(ctx.criteria_cache.borrow().ranges.len(), 4);
        }
        assert!(!budget.refused());
        assert!(budget.remaining() > 5_000_000, "{}", budget.remaining());
    }
}
