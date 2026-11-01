//! database functions: one field of the records in a database range that
//! satisfy a criteria range, summed, counted, averaged and so on. a database
//! is a range whose first row holds the field names; a criteria range has
//! field names in its first row and conditions below, the cells of one row
//! all holding and any one row sufficing.

use std::sync::Arc;

use xlsx_model::{CellRange, CellRef, CellValue, ErrorValue, MAX_COLS, MAX_ROWS, SheetId};

use crate::eval::{
    Area, EvalContext, as_area, bound_area, err, evaluate, format_number, num, parse_cached,
    resolve_sheet, used_height,
};
use crate::parser::{Expr, parse_formula};
use crate::reference::RowRange;

use super::criteria::Criterion;
use super::omitted;
use super::stats::numbers_value;

/// DAVERAGE(database, field, criteria): the mean of the field's numbers in
/// the matching records; none -> `#DIV/0!`.
pub(crate) fn daverage(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 1)
}

/// DMAX(database, field, criteria): the largest number; none -> 0.
pub(crate) fn dmax(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 4)
}

/// DMIN(database, field, criteria): the smallest number; none -> 0.
pub(crate) fn dmin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 5)
}

/// DPRODUCT(database, field, criteria): the product of the numbers; none -> 0.
pub(crate) fn dproduct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 6)
}

/// DSTDEV(database, field, criteria): the sample standard deviation.
pub(crate) fn dstdev(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 7)
}

/// DSTDEVP(database, field, criteria): the population standard deviation.
pub(crate) fn dstdevp(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 8)
}

/// DSUM(database, field, criteria): the sum of the numbers.
pub(crate) fn dsum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 9)
}

/// DVAR(database, field, criteria): the sample variance.
pub(crate) fn dvar(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 10)
}

/// DVARP(database, field, criteria): the population variance.
pub(crate) fn dvarp(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    aggregate(args, ctx, 11)
}

/// DCOUNT(database, [field], criteria): the numbers in the field of the
/// matching records, or the matching records when the field is omitted.
pub(crate) fn dcount(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    count(args, ctx, |value| matches!(value, CellValue::Number { .. }))
}

/// DCOUNTA(database, [field], criteria): the nonblank cells in the field of
/// the matching records, or the matching records when the field is omitted.
pub(crate) fn dcounta(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    count(args, ctx, |value| !matches!(value, CellValue::Empty))
}

/// DGET(database, field, criteria): the field of the one matching record;
/// no match -> `#VALUE!`, more than one -> `#NUM!`.
pub(crate) fn dget(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let query = Query::new(args, ctx)?;
        let column = query.field.ok_or(ErrorValue::Value)?;
        let matches = query.matching(ctx)?;
        match (&matches.records[..], matches.blank) {
            ([], 0) => Err(ErrorValue::Value),
            ([], 1) => Ok(num(0.0)),
            ([record], 0) => Ok(match query.database.get(ctx, *record, column)? {
                CellValue::Empty => num(0.0),
                value => value,
            }),
            _ => Err(ErrorValue::Num),
        }
    })
}

/// the numeric aggregates, by their SUBTOTAL code: text, logicals and blanks
/// in the field are skipped, an error in a matching record propagates.
fn aggregate(args: &[Expr], ctx: &EvalContext<'_>, code: i64) -> CellValue {
    compute(|| {
        let query = Query::new(args, ctx)?;
        let column = query.field.ok_or(ErrorValue::Value)?;
        let mut numbers = Vec::new();
        for record in query.matching(ctx)?.records {
            match query.database.get_ref(ctx, record, column)?.as_ref() {
                CellValue::Number { value } => numbers.push(*value),
                CellValue::Error { value } => return Err(*value),
                _ => {}
            }
        }
        Ok(numbers_value(code, &numbers))
    })
}

fn count(args: &[Expr], ctx: &EvalContext<'_>, counted: fn(&CellValue) -> bool) -> CellValue {
    compute(|| {
        let query = Query::new(args, ctx)?;
        let matches = query.matching(ctx)?;
        let Some(column) = query.field else {
            return Ok(num((matches.records.len() + matches.blank) as f64));
        };
        let mut total = 0usize;
        for record in matches.records {
            if counted(query.database.get_ref(ctx, record, column)?.as_ref()) {
                total += 1;
            }
        }
        Ok(num(total as f64))
    })
}

fn compute(f: impl FnOnce() -> Result<CellValue, ErrorValue>) -> CellValue {
    f().unwrap_or_else(err)
}

/// one condition of a criteria row.
enum Condition {
    /// a criterion on one database column.
    Field(usize, Criterion),
    /// a computed criterion: a formula under a label that names no field,
    /// written against the first record and true for the records it holds for.
    /// it is evaluated where its cell is, moved down with the record.
    Computed { formula: Arc<Expr>, at: CellRef },
    /// a constant under a label that names no field.
    Constant(bool),
}

/// a parsed database call: the database, the field column, the criteria rows.
struct Query {
    /// the database, cut to the rows the sheet uses.
    database: Area,
    /// records the cut left out, all of them blank.
    blank: usize,
    field: Option<usize>,
    rows: Vec<Vec<Condition>>,
    criteria_sheet: SheetId,
}

/// the records a query matched: area rows, plus how many of the blank
/// records past the cut matched.
struct Matches {
    records: Vec<usize>,
    blank: usize,
}

impl Query {
    fn new(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Query, ErrorValue> {
        if args.len() != 3 {
            return Err(ErrorValue::Value);
        }
        let mut database = reference(&args[0], ctx)?;
        database.cols = bound_area(database, ctx).cols;
        let labels = (0..database.cols)
            .map(|col| Ok(label(database.get_ref(ctx, 0, col)?.as_ref())))
            .collect::<Result<Vec<_>, ErrorValue>>()?;
        let field = field(&args[1], &labels, ctx)?;
        let criteria = reference(&args[2], ctx)?;
        if criteria.rows < 2 {
            return Err(ErrorValue::Value);
        }
        let criteria_rows = criteria.rows;
        let criteria = bound_area(criteria, ctx);
        let columns = (0..criteria.cols)
            .map(|col| {
                let name = label(criteria.get_ref(ctx, 0, col)?.as_ref());
                Ok(name.and_then(|name| labels.iter().position(|l| l.as_ref() == Some(&name))))
            })
            .collect::<Result<Vec<_>, ErrorValue>>()?;
        let mut rows = Vec::with_capacity(criteria.rows - 1);
        for row in 1..criteria.rows {
            let mut conditions = Vec::new();
            for (col, column) in columns.iter().enumerate() {
                let value = criteria.get_ref(ctx, row, col)?;
                let condition = match column {
                    Some(column) => criterion(*column, value.as_ref(), ctx),
                    None => computed(&criteria, row, col, value.as_ref(), ctx),
                };
                if let Some(Condition::Computed { formula, .. }) = &condition
                    && !tracked(formula, &database, criteria.sheet, ctx)
                {
                    ctx.record_unsupported_function();
                    return Err(ErrorValue::Name);
                }
                conditions.extend(condition);
            }
            rows.push(conditions);
        }
        if criteria.rows < criteria_rows {
            rows.push(Vec::new());
        }
        let computes = rows
            .iter()
            .flatten()
            .any(|condition| matches!(condition, Condition::Computed { .. }));
        let mut blank = 0;
        let used = used_height(&database, ctx);
        let reported = ctx.provider.used_rows(database.sheet) > 0;
        if !computes && reported && used < database.rows {
            blank = database.rows - used;
            database.rows = used;
        }
        Ok(Query {
            database,
            blank,
            field,
            rows,
            criteria_sheet: criteria.sheet,
        })
    }

    /// the records any criteria row holds for. every record a criteria row
    /// is tried on is charged to the budget.
    fn matching(&self, ctx: &EvalContext<'_>) -> Result<Matches, ErrorValue> {
        let mut records = Vec::new();
        for record in 1..self.database.rows {
            for row in &self.rows {
                if !ctx.consume_cells(1) {
                    return Err(ErrorValue::Num);
                }
                if self.holds(row, record, ctx)? {
                    records.push(record);
                    break;
                }
            }
        }
        let blank_holds = self.rows.iter().any(|row| {
            row.iter().all(|condition| match condition {
                Condition::Field(_, criterion) => criterion.matches(&CellValue::Empty, ctx),
                Condition::Constant(holds) => *holds,
                Condition::Computed { .. } => false,
            })
        });
        Ok(Matches {
            records,
            blank: if blank_holds { self.blank } else { 0 },
        })
    }

    fn holds(
        &self,
        row: &[Condition],
        record: usize,
        ctx: &EvalContext<'_>,
    ) -> Result<bool, ErrorValue> {
        for condition in row {
            let holds = match condition {
                Condition::Field(column, criterion) => {
                    criterion.matches(self.database.get_ref(ctx, record, *column)?.as_ref(), ctx)
                }
                Condition::Constant(holds) => *holds,
                Condition::Computed { formula, at } => {
                    if !ctx.consume_cells(1) {
                        return Err(ErrorValue::Num);
                    }
                    self.computed_holds(formula, *at, record, ctx)?
                }
            };
            if !holds {
                return Ok(false);
            }
        }
        Ok(true)
    }

    /// a computed criterion for one record: the formula moved down to it and
    /// evaluated at its own cell, moved likewise. a function the engine lacks
    /// leaves the answer unknown, so the call reports the gap with `#NAME?`.
    fn computed_holds(
        &self,
        formula: &Expr,
        at: CellRef,
        record: usize,
        ctx: &EvalContext<'_>,
    ) -> Result<bool, ErrorValue> {
        let Some(by) = u32::try_from(record - 1).ok() else {
            return Ok(false);
        };
        let (Some(moved), Some(row)) = (
            shifted(formula, by),
            at.row.checked_add(by).filter(|row| *row < MAX_ROWS),
        ) else {
            return Ok(false);
        };
        let mut local = ctx.for_sheet(self.criteria_sheet);
        local.cell = Some(CellRef::new(row, at.col));
        let unsupported = ctx.unsupported_checkpoint();
        let value = evaluate(&moved, &local);
        if ctx.unsupported_checkpoint() > unsupported {
            return Err(ErrorValue::Name);
        }
        Ok(truthy(&value))
    }
}

/// a database or criteria argument: a reference at its full size, so blank
/// records and blank criteria rows past the used cells still count; anything
/// else is `#VALUE!` or its own error.
fn reference(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Area, ErrorValue> {
    match as_area(arg, ctx) {
        Some(area) => Ok(area),
        None => match evaluate(arg, ctx) {
            CellValue::Error { value } => Err(value),
            _ => Err(ErrorValue::Value),
        },
    }
}

/// a header cell as a case-insensitive label; a blank header names nothing.
fn label(value: &CellValue) -> Option<String> {
    match value {
        CellValue::Text { value } if !value.is_empty() => Some(value.to_lowercase()),
        CellValue::Number { value } => Some(format_number(*value)),
        CellValue::Bool { value } => Some(if *value { "true" } else { "false" }.to_string()),
        _ => None,
    }
}

/// the field column: a label matched without regard to case or a 1-based
/// position, truncated. `None` when omitted, which only the counts accept.
fn field(
    arg: &Expr,
    labels: &[Option<String>],
    ctx: &EvalContext<'_>,
) -> Result<Option<usize>, ErrorValue> {
    if omitted(arg) {
        return Ok(None);
    }
    match evaluate(arg, ctx) {
        CellValue::Empty => Ok(None),
        CellValue::Number { value } => {
            let position = value.trunc();
            if position >= 1.0 && position <= labels.len() as f64 {
                Ok(Some(position as usize - 1))
            } else {
                Err(ErrorValue::Value)
            }
        }
        CellValue::Text { value } => {
            let name = value.to_lowercase();
            labels
                .iter()
                .position(|l| l.as_ref() == Some(&name))
                .map(Some)
                .ok_or(ErrorValue::Value)
        }
        CellValue::Bool { .. } => Err(ErrorValue::Value),
        CellValue::Error { value } => Err(value),
    }
}

/// the condition a cell under a field label sets, `None` for a blank cell,
/// which sets none. text without a leading operator that is not a number
/// asks for text beginning with it, the way excel reads a criteria range;
/// `=text` asks for the whole text. an error holds for no record.
fn criterion(column: usize, value: &CellValue, ctx: &EvalContext<'_>) -> Option<Condition> {
    let criterion = match value {
        CellValue::Empty => return None,
        CellValue::Text { value } if value.is_empty() => return None,
        CellValue::Text { value }
            if !value.starts_with(['<', '>', '='])
                && ctx.parse_number(value).is_none()
                && !value.eq_ignore_ascii_case("true")
                && !value.eq_ignore_ascii_case("false") =>
        {
            Criterion::parse(&format!("{value}*"), ctx)
        }
        CellValue::Error { .. } => return Some(Condition::Constant(false)),
        value => Criterion::from_value(value, ctx),
    };
    Some(Condition::Field(column, criterion))
}

/// a cell under a label that names no field: a formula there is a computed
/// criterion, a constant holds when it is true, a blank sets nothing.
fn computed(
    criteria: &Area,
    row: usize,
    col: usize,
    value: &CellValue,
    ctx: &EvalContext<'_>,
) -> Option<Condition> {
    let at = CellRef::new(
        criteria.start.row + row as u32,
        criteria.start.col + col as u32,
    );
    if let Some(source) = ctx.provider.formula(criteria.sheet, at) {
        let source = source.strip_prefix('=').unwrap_or(source);
        let formula = match ctx.parse_cache {
            Some(cache) => parse_cached(cache, source),
            None => parse_formula(source).ok().map(Arc::new),
        };
        return Some(match formula {
            Some(formula) => Condition::Computed { formula, at },
            None => Condition::Constant(false),
        });
    }
    match value {
        CellValue::Empty => None,
        CellValue::Text { value } if value.is_empty() => None,
        value => Some(Condition::Constant(truthy(value))),
    }
}

/// whether every reference a computed criterion moves stays inside the
/// database for every record it is moved to. the database argument's own
/// dependency then covers what the moved formula reads; what does not move is
/// read by the criteria cell itself, which the call depends on. a range join,
/// a function that builds a reference and a range with a relative top row and
/// an anchored bottom one, which may have been written upside down, are reads
/// the check cannot bound. a criterion reading past
/// the database is an engine gap: the cell keeps the value
/// excel cached, or shows `#NAME?` without one.
fn tracked(
    formula: &Expr,
    database: &Area,
    criteria_sheet: SheetId,
    ctx: &EvalContext<'_>,
) -> bool {
    let last = database.rows.saturating_sub(2) as u64;
    if last == 0 {
        return true;
    }
    let span = |start: u32, abs_start: bool, end: u32, abs_end: bool| {
        let reach = |row: u32, anchored: bool| {
            let row = u64::from(row);
            if anchored {
                (row, row)
            } else {
                (row + 1, row + last)
            }
        };
        let (start_low, start_high) = reach(start, abs_start);
        let (end_low, end_high) = reach(end, abs_end);
        (start_low.min(end_low), start_high.max(end_high))
    };
    let top = u64::from(database.start.row);
    let bottom = top + database.rows as u64;
    let left = u64::from(database.start.col);
    let right = left + database.cols as u64;
    let inside = |sheet: &Option<String>, rows: (u64, u64), cols: (u64, u64)| {
        let sheet = match sheet {
            None => Some(criteria_sheet),
            Some(_) => resolve_sheet(sheet, ctx),
        };
        sheet == Some(database.sheet)
            && rows.0 >= top
            && rows.1 < bottom
            && cols.0 >= left
            && cols.1 < right
    };
    let nested = |expr: &Expr| tracked(expr, database, criteria_sheet, ctx);
    match formula {
        Expr::Ref { sheet, cell } => {
            cell.abs_row
                || inside(
                    sheet,
                    span(cell.row, false, cell.row, false),
                    (u64::from(cell.col), u64::from(cell.col)),
                )
        }
        Expr::Range { range, .. }
            if !range.start.abs_row && range.end.abs_row && range.start.row < range.end.row =>
        {
            false
        }
        Expr::Range { sheet, range } => {
            (range.start.abs_row && range.end.abs_row)
                || inside(
                    sheet,
                    span(
                        range.start.row,
                        range.start.abs_row,
                        range.end.row,
                        range.end.abs_row,
                    ),
                    (u64::from(range.start.col), u64::from(range.end.col)),
                )
        }
        Expr::RowRange { sheet, range } => {
            (range.abs_start && range.abs_end)
                || inside(
                    sheet,
                    span(range.start, range.abs_start, range.end, range.abs_end),
                    (0, u64::from(MAX_COLS) - 1),
                )
        }
        Expr::ArrayLiteral { values, .. } => values.iter().all(nested),
        Expr::Unary { expr, .. } | Expr::Percent(expr) => nested(expr),
        Expr::Binary { lhs, rhs, .. } => nested(lhs) && nested(rhs),
        Expr::RangeJoin { .. } => false,
        Expr::FuncCall {
            func:
                Some(
                    super::Func::Offset
                    | super::Func::Indirect
                    | super::Func::Index
                    | super::Func::Choose
                    | super::Func::Xlookup,
                ),
            ..
        } => false,
        Expr::FuncCall { args, .. } => args.iter().all(nested),
        _ => true,
    }
}

fn truthy(value: &CellValue) -> bool {
    match value {
        CellValue::Bool { value } => *value,
        CellValue::Number { value } => *value != 0.0,
        _ => false,
    }
}

/// a formula moved down `by` rows: relative row references follow, anchored
/// ones stay. `None` when a reference would leave the sheet.
fn shifted(expr: &Expr, by: u32) -> Option<Expr> {
    let line = |row: u32, anchored: bool| {
        if anchored {
            Some(row)
        } else {
            row.checked_add(by).filter(|row| *row < MAX_ROWS)
        }
    };
    let cell = |cell: CellRef| {
        Some(CellRef {
            row: line(cell.row, cell.abs_row)?,
            ..cell
        })
    };
    let boxed = |expr: &Expr| shifted(expr, by).map(Box::new);
    Some(match expr {
        Expr::Ref { sheet, cell: at } => Expr::Ref {
            sheet: sheet.clone(),
            cell: cell(*at)?,
        },
        Expr::Range { sheet, range } => Expr::Range {
            sheet: sheet.clone(),
            range: CellRange::new(cell(range.start)?, cell(range.end)?),
        },
        Expr::RowRange { sheet, range } => {
            let start = line(range.start, range.abs_start)?;
            let end = line(range.end, range.abs_end)?;
            let (start, abs_start, end, abs_end) = if start <= end {
                (start, range.abs_start, end, range.abs_end)
            } else {
                (end, range.abs_end, start, range.abs_start)
            };
            Expr::RowRange {
                sheet: sheet.clone(),
                range: RowRange {
                    start,
                    end,
                    abs_start,
                    abs_end,
                },
            }
        }
        Expr::ArrayLiteral { cols, values } => Expr::ArrayLiteral {
            cols: *cols,
            values: values
                .iter()
                .map(|value| shifted(value, by))
                .collect::<Option<_>>()?,
        },
        Expr::Unary { op, expr } => Expr::Unary {
            op: *op,
            expr: boxed(expr)?,
        },
        Expr::Binary { op, lhs, rhs } => Expr::Binary {
            op: *op,
            lhs: boxed(lhs)?,
            rhs: boxed(rhs)?,
        },
        Expr::Percent(inner) => Expr::Percent(boxed(inner)?),
        Expr::RangeJoin { start, end } => Expr::RangeJoin {
            start: boxed(start)?,
            end: boxed(end)?,
        },
        Expr::FuncCall { name, func, args } => Expr::FuncCall {
            name: name.clone(),
            func: *func,
            args: args
                .iter()
                .map(|arg| shifted(arg, by))
                .collect::<Option<_>>()?,
        },
        other => other.clone(),
    })
}
