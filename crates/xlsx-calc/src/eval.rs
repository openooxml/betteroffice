//! tree-walking evaluator; reads cells only through `xlsx_model::CellProvider`.
//! coercion follows excel; errors propagate leftmost-first.

use std::cell::Cell;
use std::rc::Rc;

use xlsx_model::{CellProvider, CellRange, CellRef, CellValue, ErrorValue, SheetId};

use crate::parser::{BinaryOp, Expr, UnaryOp};

pub const MAX_EVALUATION_CELL_VISITS: u64 = 1_100_000;
pub const MAX_RECALCULATION_CELL_VISITS: u64 = 10_000_000;
pub const MAX_CELL_TEXT_CHARS: usize = 32_767;

pub(crate) struct EvaluationBudget {
    remaining: Cell<u64>,
}

impl EvaluationBudget {
    pub(crate) fn new(limit: u64) -> Self {
        Self {
            remaining: Cell::new(limit),
        }
    }

    fn consume(&self, count: u64) -> bool {
        let remaining = self.remaining.get();
        if count > remaining {
            return false;
        }
        self.remaining.set(remaining - count);
        true
    }
}

/// evaluation environment: the cell source plus the sheet unqualified refs
/// resolve against.
pub struct EvalContext<'a> {
    pub provider: &'a dyn CellProvider,
    pub sheet: SheetId,
    /// wall-clock as an excel date serial; `None` -> TODAY()/NOW() return #VALUE!.
    pub now_serial: Option<f64>,
    remaining_cell_visits: Cell<u64>,
    exhausted: Cell<bool>,
    unhandled_budget_errors: Cell<u64>,
    shared_budget: Option<Rc<EvaluationBudget>>,
}

impl<'a> EvalContext<'a> {
    pub fn new(provider: &'a dyn CellProvider, sheet: SheetId) -> Self {
        Self {
            provider,
            sheet,
            now_serial: None,
            remaining_cell_visits: Cell::new(MAX_EVALUATION_CELL_VISITS),
            exhausted: Cell::new(false),
            unhandled_budget_errors: Cell::new(0),
            shared_budget: None,
        }
    }

    pub fn with_now(provider: &'a dyn CellProvider, sheet: SheetId, now_serial: f64) -> Self {
        Self {
            provider,
            sheet,
            now_serial: Some(now_serial),
            remaining_cell_visits: Cell::new(MAX_EVALUATION_CELL_VISITS),
            exhausted: Cell::new(false),
            unhandled_budget_errors: Cell::new(0),
            shared_budget: None,
        }
    }

    pub(crate) fn with_budget(
        provider: &'a dyn CellProvider,
        sheet: SheetId,
        budget: Rc<EvaluationBudget>,
    ) -> Self {
        Self {
            provider,
            sheet,
            now_serial: None,
            remaining_cell_visits: Cell::new(MAX_EVALUATION_CELL_VISITS),
            exhausted: Cell::new(false),
            unhandled_budget_errors: Cell::new(0),
            shared_budget: Some(budget),
        }
    }

    pub(crate) fn consume_cells(&self, count: u64) -> bool {
        let remaining = self.remaining_cell_visits.get();
        if count > remaining {
            self.record_budget_error();
            return false;
        }
        if self
            .shared_budget
            .as_ref()
            .is_some_and(|budget| !budget.consume(count))
        {
            self.record_budget_error();
            return false;
        }
        self.remaining_cell_visits.set(remaining - count);
        true
    }

    pub(crate) fn exhausted(&self) -> bool {
        self.exhausted.get()
    }

    pub(crate) fn budget_error_checkpoint(&self) -> u64 {
        self.unhandled_budget_errors.get()
    }

    pub(crate) fn handle_budget_errors_since(&self, checkpoint: u64) {
        self.unhandled_budget_errors
            .set(self.unhandled_budget_errors.get().min(checkpoint));
    }

    pub(crate) fn has_unhandled_budget_error(&self) -> bool {
        self.unhandled_budget_errors.get() != 0
    }

    fn record_budget_error(&self) {
        self.exhausted.set(true);
        self.unhandled_budget_errors
            .set(self.unhandled_budget_errors.get().saturating_add(1));
    }
}

pub(crate) fn err(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

pub(crate) fn num(value: f64) -> CellValue {
    if value.is_finite() {
        CellValue::Number { value }
    } else {
        err(ErrorValue::Num)
    }
}

pub(crate) fn text(value: impl Into<String>) -> CellValue {
    let value = value.into();
    if value.chars().count() > MAX_CELL_TEXT_CHARS {
        err(ErrorValue::Value)
    } else {
        CellValue::Text { value }
    }
}

pub(crate) fn boolean(value: bool) -> CellValue {
    CellValue::Bool { value }
}

/// evaluate an expression against a cell provider.
pub fn evaluate(expr: &Expr, ctx: &EvalContext<'_>) -> CellValue {
    match expr {
        Expr::Number(n) => num(*n),
        Expr::Text(s) => CellValue::Text { value: s.clone() },
        Expr::Bool(b) => CellValue::Bool { value: *b },
        Expr::Error(e) => err(*e),
        Expr::Ref { sheet, cell } => resolve_ref(sheet, *cell, ctx),
        // no implicit intersection: a bare range in scalar context is #VALUE!
        Expr::Range { .. } => err(ErrorValue::Value),
        Expr::Unary { op, expr } => eval_unary(*op, expr, ctx),
        Expr::Binary { op, lhs, rhs } => eval_binary(*op, lhs, rhs, ctx),
        Expr::Percent(inner) => match to_number(&evaluate(inner, ctx)) {
            Ok(n) => num(n / 100.0),
            Err(e) => err(e),
        },
        Expr::FuncCall { name, args } => match crate::functions::lookup(name) {
            Some(f) => f(args, ctx),
            None => err(ErrorValue::Name),
        },
    }
}

/// resolve a possibly sheet-qualified cell reference to its stored value.
pub(crate) fn resolve_ref(
    sheet: &Option<String>,
    cell: CellRef,
    ctx: &EvalContext<'_>,
) -> CellValue {
    let sid = match sheet {
        Some(name) => match ctx.provider.sheet_id(name) {
            Some(id) => id,
            None => return err(ErrorValue::Ref),
        },
        None => ctx.sheet,
    };
    if !ctx.consume_cells(1) {
        return err(ErrorValue::Num);
    }
    normalize_provider_value(ctx.provider.value(sid, cell))
}

/// resolve a possibly sheet-qualified name to its sheet id (`None` -> the
/// context sheet). returns `None` only when a named sheet does not exist.
pub(crate) fn resolve_sheet(sheet: &Option<String>, ctx: &EvalContext<'_>) -> Option<SheetId> {
    match sheet {
        Some(name) => ctx.provider.sheet_id(name),
        None => Some(ctx.sheet),
    }
}

fn eval_unary(op: UnaryOp, expr: &Expr, ctx: &EvalContext<'_>) -> CellValue {
    let v = evaluate(expr, ctx);
    match to_number(&v) {
        Ok(n) => match op {
            UnaryOp::Neg => num(-n),
            UnaryOp::Plus => num(n),
        },
        Err(e) => err(e),
    }
}

fn eval_binary(op: BinaryOp, lhs: &Expr, rhs: &Expr, ctx: &EvalContext<'_>) -> CellValue {
    let lv = evaluate(lhs, ctx);
    if let CellValue::Error { value } = lv {
        return err(value);
    }
    let rv = evaluate(rhs, ctx);
    if let CellValue::Error { value } = rv {
        return err(value);
    }
    match op {
        BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div | BinaryOp::Pow => {
            let a = match to_number(&lv) {
                Ok(n) => n,
                Err(e) => return err(e),
            };
            let b = match to_number(&rv) {
                Ok(n) => n,
                Err(e) => return err(e),
            };
            arithmetic(op, a, b)
        }
        BinaryOp::Concat => {
            let a = match to_text(&lv) {
                Ok(s) => s,
                Err(e) => return err(e),
            };
            let b = match to_text(&rv) {
                Ok(s) => s,
                Err(e) => return err(e),
            };
            text(a + &b)
        }
        _ => compare(op, &lv, &rv),
    }
}

fn arithmetic(op: BinaryOp, a: f64, b: f64) -> CellValue {
    match op {
        BinaryOp::Add => num(a + b),
        BinaryOp::Sub => num(a - b),
        BinaryOp::Mul => num(a * b),
        BinaryOp::Div => {
            if b == 0.0 {
                err(ErrorValue::Div0)
            } else {
                num(a / b)
            }
        }
        BinaryOp::Pow => {
            let r = a.powf(b);
            if r.is_finite() {
                num(r)
            } else {
                err(ErrorValue::Num)
            }
        }
        _ => unreachable!("non-arithmetic op"),
    }
}

fn compare(op: BinaryOp, lv: &CellValue, rv: &CellValue) -> CellValue {
    use std::cmp::Ordering::*;
    let ord = cmp_values(lv, rv);
    let result = match op {
        BinaryOp::Eq => ord == Equal,
        BinaryOp::Ne => ord != Equal,
        BinaryOp::Lt => ord == Less,
        BinaryOp::Le => ord != Greater,
        BinaryOp::Gt => ord == Greater,
        BinaryOp::Ge => ord != Less,
        _ => unreachable!("non-comparison op"),
    };
    CellValue::Bool { value: result }
}

/// excel cross-type ordering: number < text < bool; blanks adopt the other
/// operand's type (0 / "" / false); text compares case-insensitively.
pub(crate) fn cmp_values(a: &CellValue, b: &CellValue) -> std::cmp::Ordering {
    use CellValue::*;
    use std::cmp::Ordering::Equal;
    match (a, b) {
        (Number { value: x }, Number { value: y }) => x.partial_cmp(y).unwrap_or(Equal),
        (Text { value: x }, Text { value: y }) => cmp_text(x, y),
        (Bool { value: x }, Bool { value: y }) => x.cmp(y),
        (Empty, Empty) => Equal,
        (Empty, Number { value: y }) => 0.0_f64.partial_cmp(y).unwrap_or(Equal),
        (Number { value: x }, Empty) => x.partial_cmp(&0.0).unwrap_or(Equal),
        (Empty, Text { value: y }) => cmp_text("", y),
        (Text { value: x }, Empty) => cmp_text(x, ""),
        (Empty, Bool { value: y }) => false.cmp(y),
        (Bool { value: x }, Empty) => x.cmp(&false),
        _ => type_rank(a).cmp(&type_rank(b)),
    }
}

fn cmp_text(a: &str, b: &str) -> std::cmp::Ordering {
    a.to_lowercase().cmp(&b.to_lowercase())
}

fn type_rank(v: &CellValue) -> u8 {
    match v {
        CellValue::Empty | CellValue::Number { .. } => 0,
        CellValue::Text { .. } => 1,
        CellValue::Bool { .. } => 2,
        CellValue::Error { .. } => 3,
    }
}

/// coerce a value to a number for arithmetic. numeric text coerces (excel),
/// non-numeric text is #VALUE!, errors propagate.
pub(crate) fn to_number(v: &CellValue) -> Result<f64, ErrorValue> {
    match v {
        CellValue::Empty => Ok(0.0),
        CellValue::Number { value } if value.is_finite() => Ok(*value),
        CellValue::Number { .. } => Err(ErrorValue::Num),
        CellValue::Bool { value } => Ok(if *value { 1.0 } else { 0.0 }),
        CellValue::Text { value } => parse_num(value).ok_or(ErrorValue::Value),
        CellValue::Error { value } => Err(*value),
    }
}

pub(crate) fn to_text(v: &CellValue) -> Result<String, ErrorValue> {
    match v {
        CellValue::Empty => Ok(String::new()),
        CellValue::Number { value } => Ok(format_number(*value)),
        CellValue::Bool { value } => Ok(if *value { "TRUE" } else { "FALSE" }.to_string()),
        CellValue::Text { value } => Ok(value.clone()),
        CellValue::Error { value } => Err(*value),
    }
}

pub(crate) fn to_bool(v: &CellValue) -> Result<bool, ErrorValue> {
    match v {
        CellValue::Empty => Ok(false),
        CellValue::Bool { value } => Ok(*value),
        CellValue::Number { value } => Ok(*value != 0.0),
        CellValue::Text { value } => {
            if value.eq_ignore_ascii_case("true") {
                Ok(true)
            } else if value.eq_ignore_ascii_case("false") {
                Ok(false)
            } else {
                Err(ErrorValue::Value)
            }
        }
        CellValue::Error { value } => Err(*value),
    }
}

pub(crate) fn parse_num(s: &str) -> Option<f64> {
    s.trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

/// excel "general" number formatting, good enough for text coercion: integers
/// print without a decimal, everything else uses rust's shortest round-trip.
pub(crate) fn format_number(n: f64) -> String {
    if n == 0.0 {
        return "0".to_string();
    }
    if n == n.trunc() && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    format!("{n}")
}

/// expand a range argument into its cell values (bounded rectangle, row-major).
pub(crate) fn range_values(
    sheet: &Option<String>,
    range: &CellRange,
    ctx: &EvalContext<'_>,
) -> Result<Vec<CellValue>, ErrorValue> {
    let sid = match sheet {
        Some(name) => ctx.provider.sheet_id(name).ok_or(ErrorValue::Ref)?,
        None => ctx.sheet,
    };
    let count = range_cell_count(range).ok_or(ErrorValue::Num)?;
    if !ctx.consume_cells(count) {
        return Err(ErrorValue::Num);
    }
    let mut out = Vec::with_capacity(usize::try_from(count).map_err(|_| ErrorValue::Num)?);
    for row in range.start.row..=range.end.row {
        for col in range.start.col..=range.end.col {
            out.push(normalize_provider_value(
                ctx.provider.value(sid, CellRef::new(row, col)),
            ));
        }
    }
    Ok(out)
}

/// a resolved rectangular reference: absolute top-left plus dimensions on a
/// known sheet, for positional access by function modules.
pub(crate) struct Area {
    pub sheet: SheetId,
    pub start: CellRef,
    pub rows: usize,
    pub cols: usize,
}

impl Area {
    /// value at 0-based `(row, col)` within the area.
    pub(crate) fn get(
        &self,
        ctx: &EvalContext<'_>,
        row: usize,
        col: usize,
    ) -> Result<CellValue, ErrorValue> {
        if !ctx.consume_cells(1) {
            return Err(ErrorValue::Num);
        }
        Ok(self.get_unmetered(ctx, row, col))
    }

    pub(crate) fn get_unmetered(&self, ctx: &EvalContext<'_>, row: usize, col: usize) -> CellValue {
        let cell = CellRef::new(self.start.row + row as u32, self.start.col + col as u32);
        normalize_provider_value(ctx.provider.value(self.sheet, cell))
    }

    /// all values in row-major order.
    pub(crate) fn values(&self, ctx: &EvalContext<'_>) -> Result<Vec<CellValue>, ErrorValue> {
        let count = self.cell_count().ok_or(ErrorValue::Num)?;
        if !ctx.consume_cells(count) {
            return Err(ErrorValue::Num);
        }
        let capacity = usize::try_from(count).map_err(|_| ErrorValue::Num)?;
        let mut out = Vec::with_capacity(capacity);
        for row in 0..self.rows {
            for col in 0..self.cols {
                out.push(self.get_unmetered(ctx, row, col));
            }
        }
        Ok(out)
    }

    pub(crate) fn cell_count(&self) -> Option<u64> {
        u64::try_from(self.rows)
            .ok()?
            .checked_mul(u64::try_from(self.cols).ok()?)
    }
}

/// interpret an argument as a rectangular reference (1x1 for single cells);
/// `None` for non-references or unknown sheets.
pub(crate) fn as_area(arg: &Expr, ctx: &EvalContext<'_>) -> Option<Area> {
    match arg {
        Expr::Ref { sheet, cell } => Some(Area {
            sheet: resolve_sheet(sheet, ctx)?,
            start: *cell,
            rows: 1,
            cols: 1,
        }),
        Expr::Range { sheet, range } => Some(Area {
            sheet: resolve_sheet(sheet, ctx)?,
            start: range.start,
            rows: (range.end.row - range.start.row + 1) as usize,
            cols: (range.end.col - range.start.col + 1) as usize,
        }),
        _ => None,
    }
}

fn range_cell_count(range: &CellRange) -> Option<u64> {
    let rows = u64::from(range.end.row - range.start.row + 1);
    let cols = u64::from(range.end.col - range.start.col + 1);
    rows.checked_mul(cols)
}

fn normalize_provider_value(value: CellValue) -> CellValue {
    match value {
        CellValue::Number { value } if !value.is_finite() => err(ErrorValue::Num),
        CellValue::Text { value } if value.chars().count() > MAX_CELL_TEXT_CHARS => {
            err(ErrorValue::Value)
        }
        value => value,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse_formula;
    use xlsx_model::{Sheet, Workbook};

    #[test]
    fn rejects_ranges_over_the_evaluation_budget() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        workbook.sheets.push(Sheet::new("Formula"));
        let expression = parse_formula("SUM(Data!A1:XFD1048576)").unwrap();
        let context = EvalContext::new(&workbook, SheetId(1));
        assert_eq!(
            evaluate(&expression, &context),
            CellValue::Error {
                value: ErrorValue::Num
            }
        );
        assert!(context.exhausted());
    }

    #[test]
    fn evaluates_large_ranges_within_the_cumulative_budget() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        workbook.sheets.push(Sheet::new("Formula"));
        let expression = parse_formula("SUM(Data!A1:A100001)").unwrap();
        let context = EvalContext::new(&workbook, SheetId(1));
        assert_eq!(
            evaluate(&expression, &context),
            CellValue::Number { value: 0.0 }
        );
        assert!(!context.exhausted());
    }

    #[test]
    fn metadata_functions_do_not_consume_the_referenced_area() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        let expression = parse_formula("ROWS(A1:XFD1048576)").unwrap();
        let context = EvalContext::new(&workbook, SheetId(0));
        assert_eq!(
            evaluate(&expression, &context),
            CellValue::Number { value: 1_048_576.0 }
        );
        assert!(!context.exhausted());
    }

    #[test]
    fn shared_budget_applies_across_contexts() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        let budget = Rc::new(EvaluationBudget::new(1));
        let first = EvalContext::with_budget(&workbook, SheetId(0), Rc::clone(&budget));
        let second = EvalContext::with_budget(&workbook, SheetId(0), budget);
        let expression = parse_formula("A1").unwrap();
        assert_eq!(evaluate(&expression, &first), CellValue::Empty);
        assert_eq!(
            evaluate(&expression, &second),
            CellValue::Error {
                value: ErrorValue::Num
            }
        );
        assert!(second.exhausted());
    }

    #[test]
    fn criteria_short_circuit_charges_only_actual_reads() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        let expression = parse_formula("COUNTIFS(A1:A600000,\"x\",B1:B600000,\"y\")").unwrap();
        let context = EvalContext::new(&workbook, SheetId(0));
        assert_eq!(
            evaluate(&expression, &context),
            CellValue::Number { value: 0.0 }
        );
        assert!(!context.exhausted());
    }

    #[test]
    fn exact_lookup_stops_after_the_first_match() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        let expression = parse_formula("VLOOKUP(0,A1:B1048576,2,FALSE)").unwrap();
        let context = EvalContext::new(&workbook, SheetId(0));
        assert_eq!(evaluate(&expression, &context), CellValue::Empty);
        assert!(!context.exhausted());

        let expression = parse_formula("XLOOKUP(0,A1:A1048576,B1:B1048576)").unwrap();
        let context = EvalContext::new(&workbook, SheetId(0));
        assert_eq!(evaluate(&expression, &context), CellValue::Empty);
        assert!(!context.exhausted());
    }

    #[test]
    fn non_finite_arithmetic_becomes_num_error() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        for formula in ["MEDIAN(1e308*1e308,1)", "SMALL(1e308*1e308,1)"] {
            let expression = parse_formula(formula).unwrap();
            let context = EvalContext::new(&workbook, SheetId(0));
            assert_eq!(
                evaluate(&expression, &context),
                CellValue::Error {
                    value: ErrorValue::Num
                }
            );
        }
    }

    #[test]
    fn generated_text_respects_the_excel_cell_limit() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        for formula in [
            "REPT(\"xx\",1e20)",
            "TEXTJOIN(REPT(\"x\",32767),FALSE,\"a\",\"b\")",
            "CONCAT(REPT(\"x\",20000),REPT(\"y\",20000))",
            "REPT(\"x\",32767)&\"x\"",
        ] {
            let expression = parse_formula(formula).unwrap();
            let context = EvalContext::new(&workbook, SheetId(0));
            assert_eq!(
                evaluate(&expression, &context),
                CellValue::Error {
                    value: ErrorValue::Value
                }
            );
        }
    }
}
