//! tree-walking evaluator; reads cells only through `xlsx_model::CellProvider`.
//! coercion follows excel; errors propagate leftmost-first.

use std::borrow::Cow;
use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::rc::Rc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use xlsx_model::{CellProvider, CellRef, CellValue, ErrorValue, SheetId};

use crate::TableSpec;
use crate::array::{Binding, evaluate_array};
use crate::parser::{BinaryOp, Expr, UnaryOp, parse_formula};
use crate::reference::table_rect;

pub const MAX_EVALUATION_CELL_VISITS: u64 = 1_100_000;
pub const MAX_RECALCULATION_CELL_VISITS: u64 = 10_000_000;
pub const MAX_CELL_TEXT_CHARS: usize = 32_767;
const MAX_DEFINED_NAME_DEPTH: usize = 256;
/// names one formula may have bound at once across nested `LET`/`LAMBDA`.
const MAX_BINDINGS: usize = 1024;
/// nested `LAMBDA` invocations, so a callback chain cannot exhaust the stack.
/// the parser's depth cap already keeps real formulas well under this.
const MAX_LAMBDA_DEPTH: usize = 64;

#[cfg(test)]
thread_local! {
    static DEFINED_NAME_EXPANSIONS: Cell<usize> = const { Cell::new(0) };
}

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

/// evaluation environment: the cell source, the sheet unqualified refs resolve
/// against, and the cell the formula belongs to.
pub struct EvalContext<'a> {
    pub provider: &'a dyn CellProvider,
    pub sheet: SheetId,
    /// the cell this formula belongs to; `None` -> referenceless ROW()/COLUMN()
    /// return #VALUE!.
    pub cell: Option<CellRef>,
    /// wall-clock as an excel date serial; `None` -> TODAY()/NOW() return #VALUE!.
    pub now_serial: Option<f64>,
    /// seed for the volatile random functions; `None` takes a fresh stream from
    /// a process-local counter. set before the first draw; later writes are
    /// ignored because the stream has already started.
    pub rand_seed: Option<u64>,
    random_state: Rc<Cell<Option<u64>>>,
    remaining_cell_visits: Rc<Cell<u64>>,
    exhausted: Rc<Cell<bool>>,
    unhandled_budget_errors: Rc<Cell<u64>>,
    unsupported_functions: Rc<Cell<u64>>,
    defined_name_stack: Rc<RefCell<Vec<DefinedNameKey>>>,
    defined_name_values: Rc<RefCell<HashMap<DefinedNameKey, (CellValue, bool)>>>,
    bindings: Rc<RefCell<Vec<Binding>>>,
    lambda_depth: Rc<Cell<usize>>,
    shared_budget: Option<Rc<EvaluationBudget>>,
    /// recalc-wide parse memo; `None` for one-off `evaluate` calls.
    pub(crate) parse_cache: Option<&'a ParseCache>,
}

impl<'a> EvalContext<'a> {
    pub fn new(provider: &'a dyn CellProvider, sheet: SheetId) -> Self {
        Self {
            provider,
            sheet,
            cell: None,
            now_serial: None,
            rand_seed: None,
            random_state: Rc::new(Cell::new(None)),
            remaining_cell_visits: Rc::new(Cell::new(MAX_EVALUATION_CELL_VISITS)),
            exhausted: Rc::new(Cell::new(false)),
            unhandled_budget_errors: Rc::new(Cell::new(0)),
            unsupported_functions: Rc::new(Cell::new(0)),
            defined_name_stack: Rc::new(RefCell::new(Vec::new())),
            defined_name_values: Rc::new(RefCell::new(HashMap::new())),
            bindings: Rc::new(RefCell::new(Vec::new())),
            lambda_depth: Rc::new(Cell::new(0)),
            shared_budget: None,
            parse_cache: None,
        }
    }

    pub fn with_now(provider: &'a dyn CellProvider, sheet: SheetId, now_serial: f64) -> Self {
        Self {
            provider,
            sheet,
            cell: None,
            now_serial: Some(now_serial),
            rand_seed: None,
            random_state: Rc::new(Cell::new(None)),
            remaining_cell_visits: Rc::new(Cell::new(MAX_EVALUATION_CELL_VISITS)),
            exhausted: Rc::new(Cell::new(false)),
            unhandled_budget_errors: Rc::new(Cell::new(0)),
            unsupported_functions: Rc::new(Cell::new(0)),
            defined_name_stack: Rc::new(RefCell::new(Vec::new())),
            defined_name_values: Rc::new(RefCell::new(HashMap::new())),
            bindings: Rc::new(RefCell::new(Vec::new())),
            lambda_depth: Rc::new(Cell::new(0)),
            shared_budget: None,
            parse_cache: None,
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
            cell: None,
            now_serial: None,
            rand_seed: None,
            random_state: Rc::new(Cell::new(None)),
            remaining_cell_visits: Rc::new(Cell::new(MAX_EVALUATION_CELL_VISITS)),
            exhausted: Rc::new(Cell::new(false)),
            unhandled_budget_errors: Rc::new(Cell::new(0)),
            unsupported_functions: Rc::new(Cell::new(0)),
            defined_name_stack: Rc::new(RefCell::new(Vec::new())),
            defined_name_values: Rc::new(RefCell::new(HashMap::new())),
            bindings: Rc::new(RefCell::new(Vec::new())),
            lambda_depth: Rc::new(Cell::new(0)),
            shared_budget: Some(budget),
            parse_cache: None,
        }
    }

    fn for_sheet(&self, sheet: SheetId) -> Self {
        Self {
            provider: self.provider,
            sheet,
            cell: self.cell,
            now_serial: self.now_serial,
            rand_seed: self.rand_seed,
            random_state: Rc::clone(&self.random_state),
            remaining_cell_visits: Rc::clone(&self.remaining_cell_visits),
            exhausted: Rc::clone(&self.exhausted),
            unhandled_budget_errors: Rc::clone(&self.unhandled_budget_errors),
            unsupported_functions: Rc::clone(&self.unsupported_functions),
            defined_name_stack: Rc::clone(&self.defined_name_stack),
            defined_name_values: Rc::clone(&self.defined_name_values),
            bindings: Rc::clone(&self.bindings),
            lambda_depth: Rc::clone(&self.lambda_depth),
            shared_budget: self.shared_budget.clone(),
            parse_cache: self.parse_cache,
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

    /// next draw in [0, 1) from this context's stream; advances it.
    pub(crate) fn next_random_unit(&self) -> f64 {
        let seed = self
            .random_state
            .get()
            .or(self.rand_seed)
            .unwrap_or_else(next_random_stream)
            .wrapping_add(0x9E37_79B9_7F4A_7C15);
        self.random_state.set(Some(seed));
        let mut z = seed;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        ((z ^ (z >> 31)) >> 11) as f64 / (1u64 << 53) as f64
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

    pub(crate) fn unsupported_checkpoint(&self) -> u64 {
        self.unsupported_functions.get()
    }

    pub(crate) fn handle_unsupported_since(&self, checkpoint: u64) {
        self.unsupported_functions
            .set(self.unsupported_functions.get().min(checkpoint));
    }

    /// whether a function the engine does not implement reached this
    /// evaluation's result, rather than being answered by a handler.
    pub(crate) fn has_unhandled_unsupported_function(&self) -> bool {
        self.unsupported_functions.get() != 0
    }

    pub(crate) fn record_unsupported_function(&self) {
        self.unsupported_functions
            .set(self.unsupported_functions.get().saturating_add(1));
    }

    fn record_budget_error(&self) {
        self.exhausted.set(true);
        self.unhandled_budget_errors
            .set(self.unhandled_budget_errors.get().saturating_add(1));
    }

    pub(crate) fn binding_depth(&self) -> usize {
        self.bindings.borrow().len()
    }

    /// bind a `LET` name or `LAMBDA` parameter; `false` once the formula holds
    /// as many as it may.
    pub(crate) fn push_binding(&self, binding: Binding) -> bool {
        let mut bindings = self.bindings.borrow_mut();
        if bindings.len() >= MAX_BINDINGS {
            return false;
        }
        bindings.push(binding);
        true
    }

    pub(crate) fn truncate_bindings(&self, depth: usize) {
        self.bindings.borrow_mut().truncate(depth);
    }

    /// innermost binding for `name`, so a `LAMBDA` parameter shadows an outer
    /// `LET` name and both shadow a defined name.
    pub(crate) fn binding(&self, name: &str) -> Option<Binding> {
        self.bindings
            .borrow()
            .iter()
            .rev()
            .find(|binding| binding.matches(name))
            .cloned()
    }

    pub(crate) fn enter_lambda(&self) -> bool {
        let depth = self.lambda_depth.get();
        if depth >= MAX_LAMBDA_DEPTH {
            return false;
        }
        self.lambda_depth.set(depth + 1);
        true
    }

    pub(crate) fn leave_lambda(&self) {
        self.lambda_depth
            .set(self.lambda_depth.get().saturating_sub(1));
    }

    /// evaluate a bound definition with the name held on the re-entry stack.
    fn inside_defined_name<T>(
        &self,
        binding: &DefinedNameBinding,
        evaluate_definition: impl FnOnce(&Expr, &EvalContext<'_>) -> T,
    ) -> T {
        self.defined_name_stack
            .borrow_mut()
            .push(binding.key.clone());
        let value = evaluate_definition(&binding.expression, &self.for_sheet(binding.sheet));
        self.defined_name_stack.borrow_mut().pop();
        value
    }
}

/// one stream per unseeded context, so sibling cells draw different values and
/// each recalc re-draws; a process that evaluates in the same order replays the
/// same draws.
fn next_random_stream() -> u64 {
    static NEXT_STREAM: AtomicU64 = AtomicU64::new(0);
    NEXT_STREAM.fetch_add(0x2545_F491_4F6C_DD1D, Ordering::Relaxed)
}

/// parsed-formula memo shared by graph build and every recalc. `parse_formula`
/// is pure over the source text, so entries never invalidate — a changed
/// formula is simply a different key.
pub(crate) type ParseCache = Mutex<HashMap<String, Arc<Expr>>>;

/// parse `src` once per unique text instead of once per cell that stores it.
pub(crate) fn parse_cached(cache: &ParseCache, src: &str) -> Option<Arc<Expr>> {
    if let Some(expr) = cache.lock().expect("parse cache poisoned").get(src) {
        return Some(Arc::clone(expr));
    }
    let expr = Arc::new(parse_formula(src).ok()?);
    cache
        .lock()
        .expect("parse cache poisoned")
        .insert(src.to_string(), Arc::clone(&expr));
    Some(expr)
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
        Expr::Range { .. } | Expr::ColumnRange { .. } | Expr::RowRange { .. } => {
            err(ErrorValue::Value)
        }
        Expr::RangeJoin { start, end } => match range_join_area(start, end, ctx) {
            Ok(area) if area.rows == 1 && area.cols == 1 => match area.get(ctx, 0, 0) {
                Ok(value) => value,
                Err(error) => err(error),
            },
            Ok(_) => err(ErrorValue::Value),
            Err(error) => err(error),
        },
        Expr::TableRef { table, spec } => match table_area(table, spec, ctx) {
            Ok(area) if area.rows == 1 && area.cols == 1 => match area.get(ctx, 0, 0) {
                Ok(value) => value,
                Err(error) => err(error),
            },
            Ok(_) => err(ErrorValue::Value),
            Err(error) => err(error),
        },
        Expr::Name { scope, name } => match bound(scope, name, ctx) {
            Some(binding) => binding.scalar(),
            None => evaluate_defined_name(scope, name, ctx),
        },
        Expr::Unary { op, expr } => eval_unary(*op, expr, ctx),
        Expr::Binary { op, lhs, rhs } => eval_binary(*op, lhs, rhs, ctx),
        Expr::Percent(inner) => apply_percent(&evaluate(inner, ctx)),
        Expr::Literal(value) => normalize_provider_value(value.clone()),
        Expr::ArrayLiteral { .. } => crate::array::evaluate_array(expr, ctx).into_scalar(),
        Expr::FuncCall { func, name, args } => match func {
            Some(_)
                if crate::array::is_array_builtin(name) && crate::array::args_need_array(args) =>
            {
                evaluate_array(expr, ctx).into_scalar()
            }
            Some(f) => f.call(args, ctx),
            None if crate::array::is_array_builtin(name) => evaluate_array(expr, ctx).into_scalar(),
            None => {
                ctx.record_unsupported_function();
                err(ErrorValue::Name)
            }
        },
    }
}

/// a defined name identified by the sheet its lookup resolved against.
type DefinedNameKey = (SheetId, String);

/// a defined name resolved to its parsed definition plus the sheet that
/// definition's unqualified refs bind to.
struct DefinedNameBinding {
    key: DefinedNameKey,
    expression: Arc<Expr>,
    sheet: SheetId,
}

/// a name's value is stable for the life of a context, so each one is expanded
/// at most once: without the memo a chain of `A=B+B` definitions costs 2^n. a
/// hit replays the engine gap the expansion recorded, which a handler may have
/// cleared since.
fn evaluate_defined_name(scope: &Option<String>, name: &str, ctx: &EvalContext<'_>) -> CellValue {
    let key = match defined_name_key(scope, name, ctx) {
        Ok(key) => key,
        Err(error) => return err(error),
    };
    if let Some((cached, gap)) = ctx.defined_name_values.borrow().get(&key) {
        if *gap {
            ctx.record_unsupported_function();
        }
        return cached.clone();
    }
    let binding = match bind_defined_name(key, name, ctx) {
        Ok(binding) => binding,
        Err(error) => return err(error),
    };
    let checkpoint = ctx.unsupported_checkpoint();
    let value = ctx.inside_defined_name(&binding, evaluate);
    let gap = ctx.unsupported_checkpoint() > checkpoint;
    ctx.defined_name_values
        .borrow_mut()
        .insert(binding.key, (value.clone(), gap));
    value
}

fn defined_name_key(
    scope: &Option<String>,
    name: &str,
    ctx: &EvalContext<'_>,
) -> Result<DefinedNameKey, ErrorValue> {
    let lookup_sheet = match scope {
        Some(scope) => ctx.provider.sheet_id(scope).ok_or(ErrorValue::Ref)?,
        None => ctx.sheet,
    };
    Ok((lookup_sheet, name.to_lowercase()))
}

/// Parses a name's definition, refusing re-entry, and charges the work budget.
fn bind_defined_name(
    key: DefinedNameKey,
    name: &str,
    ctx: &EvalContext<'_>,
) -> Result<DefinedNameBinding, ErrorValue> {
    {
        let stack = ctx.defined_name_stack.borrow();
        if stack.contains(&key) {
            return Err(ErrorValue::Name);
        }
        if stack.len() >= MAX_DEFINED_NAME_DEPTH {
            return Err(ErrorValue::Num);
        }
    }
    if !ctx.consume_cells(1) {
        return Err(ErrorValue::Num);
    }
    let defined = ctx
        .provider
        .defined_name(key.0, name)
        .ok_or(ErrorValue::Name)?;
    let formula = defined
        .formula
        .strip_prefix('=')
        .unwrap_or(&defined.formula);
    let expression = match ctx.parse_cache {
        Some(cache) => parse_cached(cache, formula),
        None => parse_formula(formula).ok().map(Arc::new),
    }
    .ok_or(ErrorValue::Name)?;
    let sheet = defined.local_sheet.unwrap_or(key.0);
    #[cfg(test)]
    DEFINED_NAME_EXPANSIONS.with(|count| count.set(count.get() + 1));
    Ok(DefinedNameBinding {
        key,
        expression,
        sheet,
    })
}

/// the `LET`/`LAMBDA` binding a bare name stands for, if any. a sheet
/// qualifier always means a defined name.
pub(crate) fn bound(scope: &Option<String>, name: &str, ctx: &EvalContext<'_>) -> Option<Binding> {
    match scope {
        Some(_) => None,
        None => ctx.binding(name),
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
    normalize_cow(ctx.provider.value_cow(sid, cell)).into_owned()
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
    apply_unary(op, &evaluate(expr, ctx))
}

pub(crate) fn apply_unary(op: UnaryOp, v: &CellValue) -> CellValue {
    // the lotus-style leading `+` passes its operand through whatever it is,
    // so `+A1` on text is that text; only negation needs a number
    if op == UnaryOp::Plus {
        return v.clone();
    }
    match to_number(v) {
        Ok(n) => num(-n),
        Err(e) => err(e),
    }
}

pub(crate) fn apply_percent(v: &CellValue) -> CellValue {
    match to_number(v) {
        Ok(n) => num(n / 100.0),
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
    apply_binary(op, &lv, &rv)
}

/// the operator itself, over values already evaluated; errors propagate
/// leftmost-first.
pub(crate) fn apply_binary(op: BinaryOp, lv: &CellValue, rv: &CellValue) -> CellValue {
    if let CellValue::Error { value } = lv {
        return err(*value);
    }
    if let CellValue::Error { value } = rv {
        return err(*value);
    }
    match op {
        BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div | BinaryOp::Pow => {
            let a = match to_number(lv) {
                Ok(n) => n,
                Err(e) => return err(e),
            };
            let b = match to_number(rv) {
                Ok(n) => n,
                Err(e) => return err(e),
            };
            arithmetic(op, a, b)
        }
        BinaryOp::Concat => {
            let a = match to_text(lv) {
                Ok(s) => s,
                Err(e) => return err(e),
            };
            let b = match to_text(rv) {
                Ok(s) => s,
                Err(e) => return err(e),
            };
            text(a + &b)
        }
        _ => compare(op, lv, rv),
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

pub(crate) fn cmp_text(a: &str, b: &str) -> std::cmp::Ordering {
    a.to_lowercase()
        .chars()
        .map(collation_key)
        .cmp(b.to_lowercase().chars().map(collation_key))
}

/// excel orders text by a collation rather than by code point: punctuation
/// and symbols come before digits, and digits before letters. that is what
/// puts `[Person_1]` above `[Person_10]`, where `]` against `0` would not.
fn collation_key(ch: char) -> (u8, char) {
    let class = if ch.is_alphabetic() {
        2
    } else if ch.is_numeric() {
        1
    } else {
        0
    };
    (class, ch)
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
    let s = s.trim();
    if let Ok(value) = s.parse::<f64>() {
        return value.is_finite().then_some(value);
    }
    if let Some(rest) = s.strip_suffix('%') {
        return parse_num(rest).map(|value| value / 100.0);
    }
    parse_clock(s)
}

/// a text time — `"0:15"`, `"12:30:45"`, either with a meridiem — as the
/// fraction of a day excel coerces it to.
pub(crate) fn parse_clock(s: &str) -> Option<f64> {
    let (body, pm) = match s.to_ascii_uppercase() {
        upper if upper.ends_with("AM") => (s[..s.len() - 2].trim_end(), Some(false)),
        upper if upper.ends_with("PM") => (s[..s.len() - 2].trim_end(), Some(true)),
        _ => (s, None),
    };
    let mut parts = body.split(':');
    let hour: f64 = parts.next()?.trim().parse().ok()?;
    let minute: f64 = parts.next()?.trim().parse().ok()?;
    let second: f64 = match parts.next() {
        Some(text) => text.trim().parse().ok()?,
        None => 0.0,
    };
    if parts.next().is_some() || !(0.0..60.0).contains(&minute) || !(0.0..60.0).contains(&second) {
        return None;
    }
    if hour < 0.0 || hour.fract() != 0.0 || minute.fract() != 0.0 {
        return None;
    }
    let hour = match pm {
        Some(_) if !(1.0..=12.0).contains(&hour) => return None,
        Some(true) if hour < 12.0 => hour + 12.0,
        Some(false) if hour == 12.0 => 0.0,
        _ => hour,
    };
    Some((hour * 3600.0 + minute * 60.0 + second) / 86400.0)
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

/// a reference cut to the extent the sheet uses, so a whole-column or
/// whole-row argument costs the authored data rather than a million blanks.
pub(crate) fn bound_area(mut area: Area, ctx: &EvalContext<'_>) -> Area {
    if area.rows >= xlsx_model::MAX_ROWS as usize {
        area.rows = used_height(&area, ctx);
    }
    if area.cols >= xlsx_model::MAX_COLS as usize {
        area.cols = used_width(&area, ctx);
    }
    area
}

pub(crate) fn used_height(area: &Area, ctx: &EvalContext<'_>) -> usize {
    (ctx.provider.used_rows(area.sheet) as usize)
        .saturating_sub(area.start.row as usize)
        .max(1)
}

pub(crate) fn used_width(area: &Area, ctx: &EvalContext<'_>) -> usize {
    (ctx.provider.used_cols(area.sheet) as usize)
        .saturating_sub(area.start.col as usize)
        .max(1)
}

/// a resolved rectangular reference: absolute top-left plus dimensions on a
/// known sheet, for positional access by function modules.
#[derive(Clone, Copy)]
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
        self.get_ref(ctx, row, col).map(Cow::into_owned)
    }

    /// borrowed variant of `get` for callers that only inspect the value.
    pub(crate) fn get_ref<'p>(
        &self,
        ctx: &EvalContext<'p>,
        row: usize,
        col: usize,
    ) -> Result<Cow<'p, CellValue>, ErrorValue> {
        if !ctx.consume_cells(1) {
            return Err(ErrorValue::Num);
        }
        Ok(self.get_unmetered_ref(ctx, row, col))
    }

    pub(crate) fn get_unmetered_ref<'p>(
        &self,
        ctx: &EvalContext<'p>,
        row: usize,
        col: usize,
    ) -> Cow<'p, CellValue> {
        let cell = CellRef::new(self.start.row + row as u32, self.start.col + col as u32);
        normalize_cow(ctx.provider.value_cow(self.sheet, cell))
    }

    /// all values in row-major order, borrowed where the provider can lend them.
    pub(crate) fn values_ref<'p>(
        &self,
        ctx: &EvalContext<'p>,
    ) -> Result<Vec<Cow<'p, CellValue>>, ErrorValue> {
        // reading a whole-column or whole-row band for its values costs the
        // extent the sheet reaches; the blanks past it contribute nothing and
        // would spend the recalculation's budget on the address space
        let area = bound_area(*self, ctx);
        let count = area.cell_count().ok_or(ErrorValue::Num)?;
        if !ctx.consume_cells(count) {
            return Err(ErrorValue::Num);
        }
        let capacity = usize::try_from(count).map_err(|_| ErrorValue::Num)?;
        let mut out = Vec::with_capacity(capacity);
        for row in 0..area.rows {
            for col in 0..area.cols {
                out.push(area.get_unmetered_ref(ctx, row, col));
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

/// the rectangle a reference argument designates, or the error excel reports
/// for it: an unresolved name is #NAME? wherever it appears, not a bad value.
pub(crate) fn required_area(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Area, ErrorValue> {
    match as_area(arg, ctx) {
        Some(area) => Ok(area),
        None => match evaluate(arg, ctx) {
            CellValue::Error { value } => Err(value),
            _ => Err(ErrorValue::Value),
        },
    }
}

/// interpret an argument as a rectangular reference (1x1 for single cells);
/// `None` for non-references, unknown sheets, and reference functions whose
/// result is #REF!.
pub(crate) fn as_area(arg: &Expr, ctx: &EvalContext<'_>) -> Option<Area> {
    match arg {
        Expr::FuncCall { name, args, .. } if name.eq_ignore_ascii_case("OFFSET") => {
            crate::functions::lookups::offset_area(args, ctx).ok()
        }
        Expr::FuncCall { name, args, .. } if name.eq_ignore_ascii_case("INDIRECT") => {
            crate::functions::lookups::indirect_area(args, ctx).ok()
        }
        Expr::FuncCall { name, args, .. } if name.eq_ignore_ascii_case("INDEX") => {
            crate::functions::lookups::index_area(args, ctx).ok()
        }
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
        Expr::ColumnRange { sheet, range } => Some(Area {
            sheet: resolve_sheet(sheet, ctx)?,
            start: CellRef::new(0, range.start),
            rows: xlsx_model::addr::MAX_ROWS as usize,
            cols: (range.end - range.start + 1) as usize,
        }),
        Expr::RowRange { sheet, range } => Some(Area {
            sheet: resolve_sheet(sheet, ctx)?,
            start: CellRef::new(range.start, 0),
            rows: (range.end - range.start + 1) as usize,
            cols: xlsx_model::addr::MAX_COLS as usize,
        }),
        Expr::RangeJoin { start, end } => range_join_area(start, end, ctx).ok(),
        Expr::TableRef { table, spec } => table_area(table, spec, ctx).ok(),
        Expr::Name { scope, name } => {
            if let Some(binding) = bound(scope, name, ctx) {
                return binding.reference().and_then(|expr| as_area(expr, ctx));
            }
            let key = defined_name_key(scope, name, ctx).ok()?;
            let definition = bind_defined_name(key, name, ctx).ok()?;
            ctx.inside_defined_name(&definition, as_area)
        }
        _ => None,
    }
}

/// the rectangle `start:end` designates, on the sheet both ends share. it
/// reports why an end has no rectangle, so a failed `MATCH` inside
/// `A1:INDEX(..)` surfaces as that end's own error.
pub(crate) fn range_join_area(
    start: &Expr,
    end: &Expr,
    ctx: &EvalContext<'_>,
) -> Result<Area, ErrorValue> {
    let a = endpoint_area(start, ctx)?;
    let b = endpoint_area(end, ctx)?;
    if a.sheet != b.sheet {
        return Err(ErrorValue::Ref);
    }
    let last = |area: &Area| {
        (
            area.start.row.saturating_add(area.rows as u32 - 1),
            area.start.col.saturating_add(area.cols as u32 - 1),
        )
    };
    let (a_bottom, a_right) = last(&a);
    let (b_bottom, b_right) = last(&b);
    let top = a.start.row.min(b.start.row);
    let left = a.start.col.min(b.start.col);
    Ok(Area {
        sheet: a.sheet,
        start: CellRef::new(top, left),
        rows: (a_bottom.max(b_bottom) - top + 1) as usize,
        cols: (a_right.max(b_right) - left + 1) as usize,
    })
}

/// one end of a `:` join. the reference-returning builtins already report
/// their own errors, so those pass straight through.
fn endpoint_area(expr: &Expr, ctx: &EvalContext<'_>) -> Result<Area, ErrorValue> {
    match expr {
        Expr::FuncCall { name, args, .. } if name.eq_ignore_ascii_case("OFFSET") => {
            crate::functions::lookups::offset_area(args, ctx)
        }
        Expr::FuncCall { name, args, .. } if name.eq_ignore_ascii_case("INDEX") => {
            crate::functions::lookups::index_area(args, ctx)
        }
        Expr::FuncCall { name, args, .. } if name.eq_ignore_ascii_case("INDIRECT") => {
            crate::functions::lookups::indirect_area(args, ctx)
        }
        Expr::TableRef { table, spec } => table_area(table, spec, ctx),
        Expr::RangeJoin { start, end } => range_join_area(start, end, ctx),
        Expr::Error(value) => Err(*value),
        _ => as_area(expr, ctx).ok_or(ErrorValue::Ref),
    }
}

/// the rectangle a structured reference designates, on the sheet its table
/// lives on. an unknown table name is #REF!.
pub(crate) fn table_area(
    table: &str,
    spec: &TableSpec,
    ctx: &EvalContext<'_>,
) -> Result<Area, ErrorValue> {
    let definition = ctx.provider.table(table).ok_or(ErrorValue::Ref)?;
    let rect = table_rect(definition, spec, ctx.cell)?;
    Ok(Area {
        sheet: definition.sheet,
        start: rect.start,
        rows: (rect.end.row - rect.start.row + 1) as usize,
        cols: (rect.end.col - rect.start.col + 1) as usize,
    })
}

/// the value a stored cell surfaces as, with invalid stored values mapped to
/// errors.
pub(crate) fn normalize_provider_value(value: CellValue) -> CellValue {
    match provider_error(&value) {
        Some(error) => err(error),
        None => value,
    }
}

/// Map invalid stored values to errors.
fn normalize_cow(value: Cow<'_, CellValue>) -> Cow<'_, CellValue> {
    match provider_error(&value) {
        Some(error) => Cow::Owned(err(error)),
        None => value,
    }
}

/// `ErrorValue` a stored value must surface as, if any.
fn provider_error(value: &CellValue) -> Option<ErrorValue> {
    match value {
        CellValue::Number { value } if !value.is_finite() => Some(ErrorValue::Num),
        CellValue::Text { value }
            if value.len() > MAX_CELL_TEXT_CHARS && value.chars().count() > MAX_CELL_TEXT_CHARS =>
        {
            Some(ErrorValue::Value)
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {

    /// excel coerces a text time or percentage to a number, so `"0:15"+0` and
    /// `MROUND(x,"0:15")` work on the strings a schedule is written with.
    #[test]
    fn text_coercion_reads_a_clock_and_a_percentage() {
        assert_eq!(parse_num("12"), Some(12.0));
        assert_eq!(parse_num(" 1.5 "), Some(1.5));
        assert_eq!(parse_num("50%"), Some(0.5));
        assert_eq!(parse_num("0:15"), Some(0.25 / 24.0));
        assert_eq!(parse_num("12:00"), Some(0.5));
        assert_eq!(
            parse_num("12:30:45"),
            Some((12.0 * 3600.0 + 30.0 * 60.0 + 45.0) / 86400.0)
        );
        assert_eq!(parse_num("1:30 PM"), Some(13.5 / 24.0));
        assert_eq!(parse_num("12:00 AM"), Some(0.0));
        assert_eq!(parse_num("abc"), None);
        assert_eq!(parse_num("1:60"), None);
        assert_eq!(parse_num("1:2:3:4"), None);
        assert_eq!(parse_num("13:00 PM"), None);
    }
    use super::*;
    use crate::parse_formula;
    use xlsx_model::{Cell, DefinedName, Sheet, Workbook};

    fn number(value: f64) -> Cell {
        Cell {
            value: CellValue::Number { value },
            ..Cell::default()
        }
    }

    fn defined_name(name: &str, formula: &str) -> DefinedName {
        DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: None,
            hidden: false,
        }
    }

    #[test]
    fn evaluates_workbook_and_sheet_scoped_names() {
        let mut workbook = Workbook::default();
        let mut data = Sheet::new("Data");
        data.set_cell(CellRef::parse_a1("A1").unwrap(), number(7.0));
        let mut other = Sheet::new("Other");
        other.set_cell(CellRef::parse_a1("A1").unwrap(), number(11.0));
        workbook.sheets.extend([data, other]);
        workbook.defined_names.extend([
            DefinedName {
                name: "Input".into(),
                formula: "Data!$A$1".into(),
                local_sheet: None,
                hidden: false,
            },
            DefinedName {
                name: "Input".into(),
                formula: "$A$1".into(),
                local_sheet: Some(SheetId(1)),
                hidden: false,
            },
        ]);

        let workbook_value = parse_formula("Input*2").unwrap();
        assert_eq!(
            evaluate(&workbook_value, &EvalContext::new(&workbook, SheetId(0))),
            CellValue::Number { value: 14.0 }
        );
        assert_eq!(
            evaluate(&workbook_value, &EvalContext::new(&workbook, SheetId(1))),
            CellValue::Number { value: 22.0 }
        );
        let qualified = parse_formula("Data!Input").unwrap();
        assert_eq!(
            evaluate(&qualified, &EvalContext::new(&workbook, SheetId(1))),
            CellValue::Number { value: 7.0 }
        );
    }

    #[test]
    fn unknown_and_recursive_names_return_name_error() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        workbook.defined_names.push(DefinedName {
            name: "Loop".into(),
            formula: "Loop+1".into(),
            local_sheet: None,
            hidden: false,
        });
        for formula in ["Missing", "Loop"] {
            assert_eq!(
                evaluate(
                    &parse_formula(formula).unwrap(),
                    &EvalContext::new(&workbook, SheetId(0))
                ),
                CellValue::Error {
                    value: ErrorValue::Name
                }
            );
        }
    }

    #[test]
    fn mutually_recursive_names_return_name_error() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        workbook
            .defined_names
            .extend([defined_name("Ping", "Pong"), defined_name("Pong", "Ping")]);
        let context = EvalContext::new(&workbook, SheetId(0));
        for formula in ["Ping", "Pong"] {
            assert_eq!(
                evaluate(&parse_formula(formula).unwrap(), &context),
                CellValue::Error {
                    value: ErrorValue::Name
                }
            );
        }
    }

    /// `Chain_0=Chain_1+Chain_1`, `Chain_1=Chain_2+Chain_2`, ... doubles the
    /// work per level, so re-expanding a name would cost 2^n.
    #[test]
    fn repeated_defined_name_expansion_stays_linear() {
        const DEPTH: usize = 60;
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        for index in 0..DEPTH {
            let next = index + 1;
            workbook.defined_names.push(defined_name(
                &format!("Chain_{index}"),
                &format!("Chain_{next}+Chain_{next}"),
            ));
        }
        workbook
            .defined_names
            .push(defined_name(&format!("Chain_{DEPTH}"), "1"));

        DEFINED_NAME_EXPANSIONS.with(|count| count.set(0));
        let value = evaluate(
            &parse_formula("Chain_0").unwrap(),
            &EvalContext::new(&workbook, SheetId(0)),
        );
        let expansions = DEFINED_NAME_EXPANSIONS.with(std::cell::Cell::get);

        assert_eq!(
            value,
            CellValue::Number {
                value: 2f64.powi(DEPTH as i32)
            }
        );
        assert_eq!(expansions, DEPTH + 1);
    }

    #[test]
    fn deep_acyclic_name_chains_resolve_to_the_terminal_value() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        for index in 0..MAX_DEFINED_NAME_DEPTH - 1 {
            let next = index + 1;
            workbook.defined_names.push(defined_name(
                &format!("Link_{index}"),
                &format!("Link_{next}"),
            ));
        }
        workbook.defined_names.push(defined_name(
            &format!("Link_{}", MAX_DEFINED_NAME_DEPTH - 1),
            "7",
        ));

        assert_eq!(
            evaluate(
                &parse_formula("Link_0").unwrap(),
                &EvalContext::new(&workbook, SheetId(0)),
            ),
            CellValue::Number { value: 7.0 }
        );
    }

    #[test]
    fn name_chains_past_the_depth_limit_report_a_number_error() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        for index in 0..MAX_DEFINED_NAME_DEPTH + 4 {
            let next = index + 1;
            workbook.defined_names.push(defined_name(
                &format!("Link_{index}"),
                &format!("Link_{next}"),
            ));
        }
        workbook.defined_names.push(defined_name(
            &format!("Link_{}", MAX_DEFINED_NAME_DEPTH + 4),
            "7",
        ));

        assert_eq!(
            evaluate(
                &parse_formula("Link_0").unwrap(),
                &EvalContext::new(&workbook, SheetId(0)),
            ),
            err(ErrorValue::Num)
        );
    }

    #[test]
    fn defined_name_expansion_charges_the_work_budget() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        for index in 0..4 {
            workbook.defined_names.push(defined_name(
                &format!("Step_{index}"),
                &format!("Step_{}", index + 1),
            ));
        }
        workbook.defined_names.push(defined_name("Step_4", "1"));

        let context =
            EvalContext::with_budget(&workbook, SheetId(0), Rc::new(EvaluationBudget::new(2)));
        assert_eq!(
            evaluate(&parse_formula("Step_0").unwrap(), &context),
            CellValue::Error {
                value: ErrorValue::Num
            }
        );
        assert!(context.exhausted());
    }

    #[test]
    fn rejects_ranges_over_the_evaluation_budget() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        workbook.sheets.push(Sheet::new("Formula"));
        // a cell in the far corner gives Data the extent the guard is for
        workbook.sheet_mut(SheetId(0)).unwrap().set_cell(
            CellRef::parse_a1("XFD1048576").unwrap(),
            xlsx_model::Cell {
                value: CellValue::Number { value: 1.0 },
                ..xlsx_model::Cell::default()
            },
        );
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
