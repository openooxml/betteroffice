//! array evaluation: blocks, elementwise lifting, and the dynamic-array
//! builtins. only cells the file marks `<f t="array">` take this path, and
//! anything without an array-aware implementation falls back to `evaluate`.

use std::collections::HashMap;
use std::collections::hash_map::Entry;

use xlsx_model::{CellRange, CellRef, CellValue, ErrorValue, MAX_SPILL_CELLS};

use crate::eval::{
    Area, EvalContext, apply_binary, apply_percent, apply_unary, as_area, cmp_values, err,
    evaluate, normalize_provider_value, num, to_bool, to_number,
};
use crate::functions::Func;
use crate::parser::Expr;

/// cells one intermediate array may hold. the per-formula evaluation budget
/// normally bites first; this bounds a single allocation on its own.
pub const MAX_ARRAY_CELLS: usize = 1 << 20;

/// a rectangular block of values, row-major.
#[derive(Debug, Clone, PartialEq)]
pub struct Array {
    rows: usize,
    cols: usize,
    values: Vec<CellValue>,
}

impl Array {
    fn new(rows: usize, cols: usize, values: Vec<CellValue>) -> Result<Self, ErrorValue> {
        if rows == 0 || cols == 0 || rows.checked_mul(cols) != Some(values.len()) {
            return Err(ErrorValue::Value);
        }
        if values.len() > MAX_ARRAY_CELLS {
            return Err(ErrorValue::Num);
        }
        Ok(Self { rows, cols, values })
    }

    pub fn rows(&self) -> usize {
        self.rows
    }

    pub fn cols(&self) -> usize {
        self.cols
    }

    /// value at an exact position; `#N/A` outside the block, as excel pads.
    pub fn at(&self, row: usize, col: usize) -> CellValue {
        if row >= self.rows || col >= self.cols {
            return err(ErrorValue::NA);
        }
        self.values[row * self.cols + col].clone()
    }

    /// value for a broadcast position: a single row repeats down, a single
    /// column repeats across.
    fn broadcast(&self, row: usize, col: usize) -> CellValue {
        self.at(
            if self.rows == 1 { 0 } else { row },
            if self.cols == 1 { 0 } else { col },
        )
    }

    fn row_values(&self, row: usize) -> &[CellValue] {
        &self.values[row * self.cols..(row + 1) * self.cols]
    }
}

/// what an expression evaluates to in array mode.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Scalar(CellValue),
    Array(Array),
}

impl Value {
    fn error(value: ErrorValue) -> Self {
        Value::Scalar(err(value))
    }

    fn dims(&self) -> (usize, usize) {
        match self {
            Value::Scalar(_) => (1, 1),
            Value::Array(array) => (array.rows, array.cols),
        }
    }

    fn broadcast(&self, row: usize, col: usize) -> CellValue {
        match self {
            Value::Scalar(value) => value.clone(),
            Value::Array(array) => array.broadcast(row, col),
        }
    }

    /// the value a single cell shows: the top-left element.
    pub fn into_scalar(self) -> CellValue {
        match self {
            Value::Scalar(value) => value,
            Value::Array(array) => array.at(0, 0),
        }
    }

    /// this value as a block; a single value becomes a 1x1 block.
    pub fn into_array(self) -> Array {
        match self {
            Value::Scalar(value) => Array {
                rows: 1,
                cols: 1,
                values: vec![value],
            },
            Value::Array(array) => array,
        }
    }

    pub fn as_error(&self) -> Option<ErrorValue> {
        match self {
            Value::Scalar(CellValue::Error { value }) => Some(*value),
            _ => None,
        }
    }
}

impl From<CellValue> for Value {
    fn from(value: CellValue) -> Self {
        Value::Scalar(value)
    }
}

/// the cell count of an output block, refused before anything is reserved so a
/// product of two in-budget inputs cannot ask for an out-of-budget result.
fn output_cells(rows: usize, cols: usize) -> Result<usize, ErrorValue> {
    match rows.checked_mul(cols) {
        Some(count) if count <= MAX_ARRAY_CELLS => Ok(count),
        _ => Err(ErrorValue::Num),
    }
}

/// charge the evaluation budget for a block and build it.
fn block(ctx: &EvalContext<'_>, rows: usize, cols: usize, values: Vec<CellValue>) -> Value {
    let Some(count) = rows.checked_mul(cols) else {
        return Value::error(ErrorValue::Num);
    };
    if count > MAX_ARRAY_CELLS || !ctx.consume_cells(count as u64) {
        return Value::error(ErrorValue::Num);
    }
    match Array::new(rows, cols, values) {
        Ok(array) if count == 1 => Value::Scalar(array.at(0, 0)),
        Ok(array) => Value::Array(array),
        Err(error) => Value::error(error),
    }
}

/// evaluate an expression as a possibly rectangular value.
pub fn evaluate_array(expr: &Expr, ctx: &EvalContext<'_>) -> Value {
    match expr {
        Expr::Range { .. } | Expr::ColumnRange { .. } => match as_array_area(expr, ctx) {
            Some(area) => area_values(&area, ctx),
            None => Value::error(ErrorValue::Ref),
        },
        Expr::Name { .. } => match as_array_area(expr, ctx) {
            Some(area) => area_values(&area, ctx),
            None => Value::Scalar(evaluate(expr, ctx)),
        },
        Expr::ArrayLiteral { cols, values } => array_literal(*cols, values, ctx),
        Expr::Unary { op, expr } => map1(evaluate_array(expr, ctx), ctx, |v| apply_unary(*op, v)),
        Expr::Percent(inner) => map1(evaluate_array(inner, ctx), ctx, apply_percent),
        Expr::Binary { op, lhs, rhs } => {
            let left = evaluate_array(lhs, ctx);
            let right = evaluate_array(rhs, ctx);
            map2(left, right, ctx, |a, b| apply_binary(*op, a, b))
        }
        // OFFSET yields a reference, so a multi-cell result is a block here
        Expr::FuncCall { name, .. } if name.eq_ignore_ascii_case("OFFSET") => {
            match as_array_area(expr, ctx) {
                Some(area) => area_values(&area, ctx),
                None => Value::Scalar(evaluate(expr, ctx)),
            }
        }
        Expr::FuncCall { name, func, args } => call(name, *func, args, ctx),
        _ => Value::Scalar(evaluate(expr, ctx)),
    }
}

fn array_literal(cols: usize, values: &[Expr], ctx: &EvalContext<'_>) -> Value {
    if cols == 0 || !values.len().is_multiple_of(cols) {
        return Value::error(ErrorValue::Value);
    }
    let cells: Vec<CellValue> = values
        .iter()
        .map(|value| evaluate_array(value, ctx).into_scalar())
        .collect();
    block(ctx, values.len() / cols, cols, cells)
}

/// elementwise over one value.
fn map1(value: Value, ctx: &EvalContext<'_>, f: impl Fn(&CellValue) -> CellValue) -> Value {
    match value {
        Value::Scalar(value) => Value::Scalar(f(&value)),
        Value::Array(array) => {
            let cells: Vec<CellValue> = array.values.iter().map(&f).collect();
            block(ctx, array.rows, array.cols, cells)
        }
    }
}

/// elementwise over two values, broadcasting single rows and columns.
fn map2(
    left: Value,
    right: Value,
    ctx: &EvalContext<'_>,
    f: impl Fn(&CellValue, &CellValue) -> CellValue,
) -> Value {
    if let (Value::Scalar(a), Value::Scalar(b)) = (&left, &right) {
        return Value::Scalar(f(a, b));
    }
    let (lr, lc) = left.dims();
    let (rr, rc) = right.dims();
    let (rows, cols) = (lr.max(rr), lc.max(rc));
    let Some(count) = rows.checked_mul(cols) else {
        return Value::error(ErrorValue::Num);
    };
    if count > MAX_ARRAY_CELLS {
        return Value::error(ErrorValue::Num);
    }
    let mut cells = Vec::with_capacity(count);
    for row in 0..rows {
        for col in 0..cols {
            cells.push(f(&left.broadcast(row, col), &right.broadcast(row, col)));
        }
    }
    block(ctx, rows, cols, cells)
}

/// a rectangular reference, with whole-column ranges cut to the rows the sheet
/// actually uses so `A:A` costs the authored data, not a million blanks.
fn as_array_area(expr: &Expr, ctx: &EvalContext<'_>) -> Option<Area> {
    let mut area = as_area(expr, ctx)?;
    if area.rows >= xlsx_model::MAX_ROWS as usize {
        let used = ctx.provider.used_rows(area.sheet) as usize;
        area.rows = used.saturating_sub(area.start.row as usize).max(1);
    }
    Some(area)
}

fn area_values(area: &Area, ctx: &EvalContext<'_>) -> Value {
    match area.values_ref(ctx) {
        Ok(values) => {
            let values = values
                .into_iter()
                .map(std::borrow::Cow::into_owned)
                .collect();
            match Array::new(area.rows, area.cols, values) {
                Ok(array) if area.rows * area.cols == 1 => Value::Scalar(array.at(0, 0)),
                Ok(array) => Value::Array(array),
                Err(error) => Value::error(error),
            }
        }
        Err(error) => Value::error(error),
    }
}

/// read one argument as a block, whatever shape it has.
fn argument(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<Array, ErrorValue> {
    let value = args
        .get(index)
        .map(|arg| evaluate_array(arg, ctx))
        .ok_or(ErrorValue::Value)?;
    match value {
        Value::Scalar(CellValue::Error { value }) => Err(value),
        Value::Scalar(value) => Array::new(1, 1, vec![value]),
        Value::Array(array) => Ok(array),
    }
}

/// an optional scalar argument; omitted and blank arguments give `None`.
fn optional(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
) -> Result<Option<CellValue>, ErrorValue> {
    let Some(arg) = args.get(index) else {
        return Ok(None);
    };
    match evaluate_array(arg, ctx).into_scalar() {
        CellValue::Error { value } => Err(value),
        CellValue::Empty => Ok(None),
        value => Ok(Some(value)),
    }
}

fn optional_number(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: f64,
) -> Result<f64, ErrorValue> {
    match optional(args, ctx, index)? {
        Some(value) => to_number(&value),
        None => Ok(default),
    }
}

fn optional_bool(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: bool,
) -> Result<bool, ErrorValue> {
    match optional(args, ctx, index)? {
        Some(value) => to_bool(&value),
        None => Ok(default),
    }
}

/// a count argument as a usize index, rejecting anything outside the sheet.
fn index_of(value: f64) -> Result<usize, ErrorValue> {
    if !value.is_finite() || value < 1.0 || value > MAX_ARRAY_CELLS as f64 {
        return Err(ErrorValue::Value);
    }
    Ok(value.trunc() as usize)
}

fn blank(value: &CellValue) -> bool {
    matches!(value, CellValue::Empty)
}

/// excel's dynamic-array ordering: blanks sort last whichever way the order
/// runs, everything else by the usual cross-type ranking.
fn order_values(a: &CellValue, b: &CellValue, descending: bool) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (blank(a), blank(b)) {
        (true, true) => Ordering::Equal,
        (true, false) => Ordering::Greater,
        (false, true) => Ordering::Less,
        (false, false) => {
            let ordering = cmp_values(a, b);
            if descending {
                ordering.reverse()
            } else {
                ordering
            }
        }
    }
}

/// the key `UNIQUE` groups by; hashing it keeps dedup linear rather than
/// comparing every slice against every kept one.
fn identity(slice: &[CellValue]) -> String {
    let mut key = String::with_capacity(slice.len() * 8);
    for value in slice {
        match value {
            CellValue::Empty => key.push_str("n:0"),
            CellValue::Number { value } => {
                key.push_str("n:");
                key.push_str(&crate::eval::format_number(*value + 0.0));
            }
            CellValue::Text { value } => {
                key.push_str("t:");
                key.push_str(&value.to_lowercase());
            }
            CellValue::Bool { value } => key.push_str(if *value { "b:1" } else { "b:0" }),
            CellValue::Error { value } => {
                key.push_str("e:");
                key.push_str(value.as_str());
            }
        }
        key.push('\u{1f}');
    }
    key
}

fn rows_of(array: &Array, row: usize) -> Vec<CellValue> {
    array.row_values(row).to_vec()
}

fn column_of(array: &Array, col: usize) -> Vec<CellValue> {
    (0..array.rows).map(|row| array.at(row, col)).collect()
}

fn from_rows(ctx: &EvalContext<'_>, cols: usize, rows: Vec<Vec<CellValue>>) -> Value {
    let count = rows.len();
    block(ctx, count, cols, rows.into_iter().flatten().collect())
}

fn from_columns(ctx: &EvalContext<'_>, rows: usize, columns: Vec<Vec<CellValue>>) -> Value {
    let cols = columns.len();
    let Ok(count) = output_cells(rows, cols) else {
        return Value::error(ErrorValue::Num);
    };
    let mut cells = Vec::with_capacity(count);
    for row in 0..rows {
        for column in &columns {
            cells.push(column.get(row).cloned().unwrap_or(CellValue::Empty));
        }
    }
    block(ctx, rows, cols, cells)
}

fn result(value: Result<Value, ErrorValue>) -> Value {
    value.unwrap_or_else(Value::error)
}

// ---------------------------------------------------------------- dispatch

/// an array-aware builtin: lazy arguments in, one rectangular value out.
type ArrayFn = fn(&[Expr], &EvalContext<'_>) -> Value;

fn call(name: &str, func: Option<Func>, args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let upper = crate::functions::bare_name(name).to_ascii_uppercase();
    if let Some(f) = lookup_array(&upper) {
        return f(args, ctx);
    }
    match func {
        Some(f) if lifts(&upper) => lift(f, args, ctx),
        Some(f) => Value::Scalar(f.call(args, ctx)),
        None => {
            ctx.record_unsupported_function();
            Value::error(ErrorValue::Name)
        }
    }
}

/// call a scalar builtin once per output position, splicing each element in as
/// a literal argument. with no array argument the call is left untouched, so
/// laziness and reference arguments behave exactly as in scalar mode.
fn lift(f: Func, args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let values: Vec<Value> = args.iter().map(|arg| evaluate_array(arg, ctx)).collect();
    let mut rows = 1usize;
    let mut cols = 1usize;
    for value in &values {
        let (r, c) = value.dims();
        rows = rows.max(r);
        cols = cols.max(c);
    }
    if rows == 1 && cols == 1 {
        return Value::Scalar(f.call(args, ctx));
    }
    let Some(count) = rows.checked_mul(cols) else {
        return Value::error(ErrorValue::Num);
    };
    if count > MAX_ARRAY_CELLS {
        return Value::error(ErrorValue::Num);
    }
    let mut spliced: Vec<Expr> = args.iter().map(|_| Expr::Number(0.0)).collect();
    let mut cells = Vec::with_capacity(count);
    for row in 0..rows {
        for col in 0..cols {
            for (slot, value) in spliced.iter_mut().zip(&values) {
                *slot = Expr::Literal(value.broadcast(row, col));
            }
            cells.push(f.call(&spliced, ctx));
        }
    }
    block(ctx, rows, cols, cells)
}

/// scalar builtins whose every argument is a single value, so an array
/// argument means "do this once per element".
fn lifts(name: &str) -> bool {
    matches!(
        name,
        "ABS"
            | "CEILING"
            | "CHAR"
            | "CLEAN"
            | "CODE"
            | "DATE"
            | "DATEDIF"
            | "DAY"
            | "EDATE"
            | "EOMONTH"
            | "EXACT"
            | "EXP"
            | "FIND"
            | "FLOOR"
            | "HOUR"
            | "INT"
            | "ISBLANK"
            | "ISERR"
            | "ISERROR"
            | "ISLOGICAL"
            | "ISNA"
            | "ISNUMBER"
            | "ISTEXT"
            | "LEFT"
            | "LEN"
            | "LN"
            | "LOG"
            | "LOG10"
            | "LOWER"
            | "MID"
            | "MINUTE"
            | "MOD"
            | "MONTH"
            | "MROUND"
            | "N"
            | "NOT"
            | "NUMBERVALUE"
            | "POWER"
            | "PROPER"
            | "REPLACE"
            | "REPT"
            | "RIGHT"
            | "ROUND"
            | "ROUNDDOWN"
            | "ROUNDUP"
            | "SEARCH"
            | "SECOND"
            | "SIGN"
            | "SQRT"
            | "SUBSTITUTE"
            | "T"
            | "TANH"
            | "TEXT"
            | "TIME"
            | "TRIM"
            | "TRUNC"
            | "UPPER"
            | "VALUE"
            | "WEEKDAY"
            | "YEAR"
    )
}

fn lookup_array(name: &str) -> Option<ArrayFn> {
    Some(match name {
        "ANCHORARRAY" => anchorarray,
        "AVERAGE" => average,
        "COUNT" => count,
        "COUNTA" => counta,
        "MATCH" => match_,
        "MMULT" => mmult,
        "MAX" => max,
        "MIN" => min,
        "PRODUCT" => product,
        "SUM" => sum,
        "CHOOSECOLS" => choosecols,
        "CHOOSEROWS" => chooserows,
        "COLUMN" => column,
        "COLUMNS" => columns,
        "DROP" => drop,
        "EXPAND" => expand,
        "FILTER" => filter,
        "HSTACK" => hstack,
        "IF" => if_,
        "IFERROR" => iferror,
        "IFNA" => ifna,
        "INDEX" => index,
        "ROW" => row,
        "ROWS" => rows_fn,
        "SUMPRODUCT" => sumproduct,
        "SEQUENCE" => sequence,
        "SORT" => sort,
        "SORTBY" => sortby,
        "TAKE" => take,
        "TOCOL" => tocol,
        "TOROW" => torow,
        "TRANSPOSE" => transpose,
        "UNIQUE" => unique,
        "VSTACK" => vstack,
        _ => return None,
    })
}

// --------------------------------------------------------------- builtins

/// the block a spilled formula produced, addressed by its anchor cell.
fn anchorarray(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let [Expr::Ref { sheet, cell }] = args else {
        return Value::error(ErrorValue::Ref);
    };
    let Some(sid) = crate::eval::resolve_sheet(sheet, ctx) else {
        return Value::error(ErrorValue::Ref);
    };
    let Some(range) = ctx.provider.spill_range(sid, *cell) else {
        return Value::Scalar(normalize_provider_value(ctx.provider.value(sid, *cell)));
    };
    let area = Area {
        sheet: sid,
        start: range.start,
        rows: (range.end.row - range.start.row + 1) as usize,
        cols: (range.end.col - range.start.col + 1) as usize,
    };
    area_values(&area, ctx)
}

fn filter(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() < 2 || args.len() > 3 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let include = argument(args, ctx, 1)?;
        let by_rows = include.rows == data.rows && include.cols == 1;
        let by_cols = include.cols == data.cols && include.rows == 1;
        if !by_rows && !by_cols {
            return Err(ErrorValue::Value);
        }
        let mut keep = Vec::new();
        for position in 0..include.values.len() {
            match to_bool(&include.values[position]) {
                Ok(true) => keep.push(position),
                Ok(false) => {}
                Err(error) => return Err(error),
            }
        }
        if keep.is_empty() {
            return Ok(match optional(args, ctx, 2)? {
                Some(value) => Value::Scalar(value),
                None => Value::error(ErrorValue::NA),
            });
        }
        Ok(if by_rows {
            from_rows(
                ctx,
                data.cols,
                keep.into_iter().map(|row| rows_of(&data, row)).collect(),
            )
        } else {
            from_columns(
                ctx,
                data.rows,
                keep.into_iter().map(|col| column_of(&data, col)).collect(),
            )
        })
    })())
}

fn sort(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.is_empty() || args.len() > 4 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let key = index_of(optional_number(args, ctx, 1, 1.0)?)?;
        let descending = optional_number(args, ctx, 2, 1.0)? < 0.0;
        let by_col = optional_bool(args, ctx, 3, false)?;
        Ok(if by_col {
            if key > data.rows {
                return Err(ErrorValue::Value);
            }
            let mut columns: Vec<Vec<CellValue>> =
                (0..data.cols).map(|col| column_of(&data, col)).collect();
            columns.sort_by(|a, b| order_values(&a[key - 1], &b[key - 1], descending));
            from_columns(ctx, data.rows, columns)
        } else {
            if key > data.cols {
                return Err(ErrorValue::Value);
            }
            let mut rows: Vec<Vec<CellValue>> =
                (0..data.rows).map(|row| rows_of(&data, row)).collect();
            rows.sort_by(|a, b| order_values(&a[key - 1], &b[key - 1], descending));
            from_rows(ctx, data.cols, rows)
        })
    })())
}

fn sortby(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() < 2 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let mut keys: Vec<(Array, bool)> = Vec::new();
        let mut index = 1;
        while index < args.len() {
            let by = argument(args, ctx, index)?;
            if by.rows != data.rows || by.cols != 1 {
                return Err(ErrorValue::Value);
            }
            let descending = optional_number(args, ctx, index + 1, 1.0)? < 0.0;
            keys.push((by, descending));
            index += 2;
        }
        let mut order: Vec<usize> = (0..data.rows).collect();
        order.sort_by(|a, b| {
            for (by, descending) in &keys {
                let ordering = order_values(&by.at(*a, 0), &by.at(*b, 0), *descending);
                if ordering != std::cmp::Ordering::Equal {
                    return ordering;
                }
            }
            std::cmp::Ordering::Equal
        });
        Ok(from_rows(
            ctx,
            data.cols,
            order.into_iter().map(|row| rows_of(&data, row)).collect(),
        ))
    })())
}

fn unique(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.is_empty() || args.len() > 3 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let by_col = optional_bool(args, ctx, 1, false)?;
        let exactly_once = optional_bool(args, ctx, 2, false)?;
        let count = if by_col { data.cols } else { data.rows };
        let slices: Vec<Vec<CellValue>> = (0..count)
            .map(|index| {
                if by_col {
                    column_of(&data, index)
                } else {
                    rows_of(&data, index)
                }
            })
            .collect();
        let mut seen: HashMap<String, usize> = HashMap::with_capacity(count);
        let mut occurrences: Vec<usize> = vec![0; count];
        let mut first: Vec<usize> = Vec::new();
        for index in 0..count {
            match seen.entry(identity(&slices[index])) {
                Entry::Occupied(entry) => occurrences[*entry.get()] += 1,
                Entry::Vacant(entry) => {
                    entry.insert(index);
                    occurrences[index] = 1;
                    first.push(index);
                }
            }
        }
        let kept: Vec<usize> = first
            .into_iter()
            .filter(|index| !exactly_once || occurrences[*index] == 1)
            .collect();
        if kept.is_empty() {
            return Ok(Value::error(ErrorValue::NA));
        }
        Ok(if by_col {
            from_columns(
                ctx,
                data.rows,
                kept.into_iter()
                    .map(|index| slices[index].clone())
                    .collect(),
            )
        } else {
            from_rows(
                ctx,
                data.cols,
                kept.into_iter()
                    .map(|index| slices[index].clone())
                    .collect(),
            )
        })
    })())
}

fn sequence(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.is_empty() || args.len() > 4 {
            return Err(ErrorValue::Value);
        }
        let rows = index_of(optional_number(args, ctx, 0, 1.0)?)?;
        let cols = index_of(optional_number(args, ctx, 1, 1.0)?)?;
        let start = optional_number(args, ctx, 2, 1.0)?;
        let step = optional_number(args, ctx, 3, 1.0)?;
        let cells = (0..output_cells(rows, cols)?)
            .map(|index| num(start + step * index as f64))
            .collect();
        Ok(block(ctx, rows, cols, cells))
    })())
}

fn choosecols(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    choose(args, ctx, false)
}

fn chooserows(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    choose(args, ctx, true)
}

/// `CHOOSEROWS`/`CHOOSECOLS`: each index argument may itself be an array, and
/// a negative index counts from the end.
fn choose(args: &[Expr], ctx: &EvalContext<'_>, by_row: bool) -> Value {
    result((|| {
        if args.len() < 2 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let limit = if by_row { data.rows } else { data.cols };
        let mut picks = Vec::new();
        for position in 1..args.len() {
            for value in argument(args, ctx, position)?.values {
                let index = to_number(&value)?.trunc();
                let resolved = if index < 0.0 {
                    limit as f64 + index
                } else {
                    index - 1.0
                };
                if resolved < 0.0 || resolved >= limit as f64 {
                    return Err(ErrorValue::Value);
                }
                picks.push(resolved as usize);
            }
        }
        if picks.is_empty() {
            return Err(ErrorValue::Value);
        }
        Ok(if by_row {
            from_rows(
                ctx,
                data.cols,
                picks.into_iter().map(|row| rows_of(&data, row)).collect(),
            )
        } else {
            from_columns(
                ctx,
                data.rows,
                picks.into_iter().map(|col| column_of(&data, col)).collect(),
            )
        })
    })())
}

fn hstack(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        let blocks = stack_arguments(args, ctx)?;
        let rows = blocks.iter().map(|block| block.rows).max().unwrap_or(1);
        let cols = blocks
            .iter()
            .try_fold(0usize, |total, block| total.checked_add(block.cols))
            .ok_or(ErrorValue::Num)?;
        let mut cells = Vec::with_capacity(output_cells(rows, cols)?);
        for row in 0..rows {
            for block in &blocks {
                for col in 0..block.cols {
                    cells.push(block.at(row, col));
                }
            }
        }
        Ok(block(ctx, rows, cols, cells))
    })())
}

fn vstack(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        let blocks = stack_arguments(args, ctx)?;
        let cols = blocks.iter().map(|block| block.cols).max().unwrap_or(1);
        let rows = blocks
            .iter()
            .try_fold(0usize, |total, block| total.checked_add(block.rows))
            .ok_or(ErrorValue::Num)?;
        let mut cells = Vec::with_capacity(output_cells(rows, cols)?);
        for block in &blocks {
            for row in 0..block.rows {
                for col in 0..cols {
                    cells.push(block.at(row, col));
                }
            }
        }
        Ok(block(ctx, rows, cols, cells))
    })())
}

fn stack_arguments(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Vec<Array>, ErrorValue> {
    if args.is_empty() {
        return Err(ErrorValue::Value);
    }
    (0..args.len())
        .map(|index| argument(args, ctx, index))
        .collect()
}

fn tocol(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    flatten(args, ctx, false)
}

fn torow(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    flatten(args, ctx, true)
}

fn flatten(args: &[Expr], ctx: &EvalContext<'_>, by_row: bool) -> Value {
    result((|| {
        if args.is_empty() || args.len() > 3 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let ignore = optional_number(args, ctx, 1, 0.0)?.trunc() as i64;
        if !(0..=3).contains(&ignore) {
            return Err(ErrorValue::Value);
        }
        let scan_by_column = optional_bool(args, ctx, 2, false)?;
        let mut cells = Vec::new();
        let mut push = |value: CellValue| {
            let skip = match ignore {
                1 => blank(&value),
                2 => matches!(value, CellValue::Error { .. }),
                3 => blank(&value) || matches!(value, CellValue::Error { .. }),
                _ => false,
            };
            if !skip {
                cells.push(value);
            }
        };
        if scan_by_column {
            for col in 0..data.cols {
                for row in 0..data.rows {
                    push(data.at(row, col));
                }
            }
        } else {
            for value in &data.values {
                push(value.clone());
            }
        }
        if cells.is_empty() {
            return Ok(Value::error(ErrorValue::NA));
        }
        let count = cells.len();
        Ok(if by_row {
            block(ctx, 1, count, cells)
        } else {
            block(ctx, count, 1, cells)
        })
    })())
}

fn expand(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() < 2 || args.len() > 4 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let rows = match optional(args, ctx, 1)? {
            Some(value) => index_of(to_number(&value)?)?,
            None => data.rows,
        };
        let cols = match optional(args, ctx, 2)? {
            Some(value) => index_of(to_number(&value)?)?,
            None => data.cols,
        };
        if rows < data.rows || cols < data.cols {
            return Err(ErrorValue::Value);
        }
        let pad = optional(args, ctx, 3)?.unwrap_or(err(ErrorValue::NA));
        let mut cells = Vec::with_capacity(output_cells(rows, cols)?);
        for row in 0..rows {
            for col in 0..cols {
                cells.push(if row < data.rows && col < data.cols {
                    data.at(row, col)
                } else {
                    pad.clone()
                });
            }
        }
        Ok(block(ctx, rows, cols, cells))
    })())
}

fn drop(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    slice(args, ctx, true)
}

fn take(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    slice(args, ctx, false)
}

/// `DROP` removes the first or last n rows/columns, `TAKE` keeps them.
fn slice(args: &[Expr], ctx: &EvalContext<'_>, dropping: bool) -> Value {
    result((|| {
        if args.len() < 2 || args.len() > 3 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let row_span = span(optional(args, ctx, 1)?, data.rows, dropping)?;
        let col_span = span(optional(args, ctx, 2)?, data.cols, dropping)?;
        if row_span.is_empty() || col_span.is_empty() {
            return Err(ErrorValue::Value);
        }
        let (rows, cols) = (row_span.len(), col_span.len());
        let mut cells = Vec::with_capacity(output_cells(rows, cols)?);
        for row in row_span {
            for col in col_span.clone() {
                cells.push(data.at(row, col));
            }
        }
        Ok(block(ctx, rows, cols, cells))
    })())
}

fn span(
    count: Option<CellValue>,
    length: usize,
    dropping: bool,
) -> Result<std::ops::Range<usize>, ErrorValue> {
    let Some(count) = count else {
        return Ok(0..length);
    };
    let count = to_number(&count)?.trunc();
    if !count.is_finite() || count.abs() > length as f64 {
        return Ok(if dropping { 0..0 } else { 0..length });
    }
    let magnitude = count.abs() as usize;
    Ok(match (dropping, count < 0.0) {
        (true, false) => magnitude..length,
        (true, true) => 0..length - magnitude,
        (false, false) => 0..magnitude,
        (false, true) => length - magnitude..length,
    })
}

fn transpose(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() != 1 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let mut cells = Vec::with_capacity(data.values.len());
        for col in 0..data.cols {
            for row in 0..data.rows {
                cells.push(data.at(row, col));
            }
        }
        Ok(block(ctx, data.cols, data.rows, cells))
    })())
}

fn rows_fn(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    dimension(args, ctx, true)
}

fn columns(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    dimension(args, ctx, false)
}

/// `ROWS`/`COLUMNS`: a reference keeps the cheap scalar answer, a computed
/// block is measured directly.
fn dimension(args: &[Expr], ctx: &EvalContext<'_>, by_row: bool) -> Value {
    if args.len() == 1 && as_area(&args[0], ctx).is_some() {
        let name = if by_row { "ROWS" } else { "COLUMNS" };
        if let Some(f) = crate::functions::resolve(name) {
            return Value::Scalar(f.call(args, ctx));
        }
    }
    result((|| {
        let data = argument(args, ctx, 0)?;
        Ok(Value::Scalar(num(if by_row {
            data.rows as f64
        } else {
            data.cols as f64
        })))
    })())
}

fn index(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    if args.first().is_some_and(|arg| as_area(arg, ctx).is_some())
        && let Some(f) = crate::functions::resolve("INDEX")
    {
        return Value::Scalar(f.call(args, ctx));
    }
    result((|| {
        if args.len() < 2 || args.len() > 3 {
            return Err(ErrorValue::Value);
        }
        let data = argument(args, ctx, 0)?;
        let first = optional_number(args, ctx, 1, 0.0)?.trunc();
        let second = match args.len() {
            3 => Some(optional_number(args, ctx, 2, 0.0)?.trunc()),
            _ => None,
        };
        // one index against a single row addresses that row's columns
        let (row, col) = match second {
            Some(col) => (first, col),
            None if data.rows == 1 && data.cols > 1 => (0.0, first),
            None => (first, 0.0),
        };
        if row < 0.0 || col < 0.0 || row > data.rows as f64 || col > data.cols as f64 {
            return Err(ErrorValue::Ref);
        }
        Ok(match (row as usize, col as usize) {
            (0, 0) => Value::Array(data),
            (0, col) => block(ctx, data.rows, 1, column_of(&data, col - 1)),
            (row, 0) => block(ctx, 1, data.cols, rows_of(&data, row - 1)),
            (row, col) => Value::Scalar(data.at(row - 1, col - 1)),
        })
    })())
}

/// a single condition still picks one branch lazily; an array condition
/// evaluates both and chooses per element.
fn if_(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    if !(2..=3).contains(&args.len()) {
        return Value::error(ErrorValue::Value);
    }
    let condition = evaluate_array(&args[0], ctx);
    if let Some(error) = condition.as_error() {
        return Value::error(error);
    }
    if let Value::Scalar(value) = &condition {
        return match to_bool(value) {
            Ok(true) => evaluate_array(&args[1], ctx),
            Ok(false) => match args.get(2) {
                Some(arg) => evaluate_array(arg, ctx),
                None => Value::Scalar(CellValue::Bool { value: false }),
            },
            Err(error) => Value::error(error),
        };
    }
    let whenever = evaluate_array(&args[1], ctx);
    let otherwise = match args.get(2) {
        Some(arg) => evaluate_array(arg, ctx),
        None => Value::Scalar(CellValue::Bool { value: false }),
    };
    let (rows, cols) = condition.dims();
    let mut cells = Vec::with_capacity(rows.saturating_mul(cols));
    for row in 0..rows {
        for col in 0..cols {
            cells.push(match to_bool(&condition.broadcast(row, col)) {
                Ok(true) => whenever.broadcast(row, col),
                Ok(false) => otherwise.broadcast(row, col),
                Err(error) => err(error),
            });
        }
    }
    block(ctx, rows, cols, cells)
}

fn iferror(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    fallback(args, ctx, |value| matches!(value, CellValue::Error { .. }))
}

fn ifna(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    fallback(args, ctx, |value| {
        matches!(
            value,
            CellValue::Error {
                value: ErrorValue::NA
            }
        )
    })
}

fn fallback(args: &[Expr], ctx: &EvalContext<'_>, caught: fn(&CellValue) -> bool) -> Value {
    if args.len() != 2 {
        return Value::error(ErrorValue::Value);
    }
    let budget = ctx.budget_error_checkpoint();
    let unsupported = ctx.unsupported_checkpoint();
    let value = evaluate_array(&args[0], ctx);
    let needs_fallback = match &value {
        Value::Scalar(value) => caught(value),
        Value::Array(array) => array.values.iter().any(caught),
    };
    if !needs_fallback {
        return value;
    }
    ctx.handle_budget_errors_since(budget);
    ctx.handle_unsupported_since(unsupported);
    let other = evaluate_array(&args[1], ctx);
    match value {
        Value::Scalar(_) => other,
        Value::Array(array) => {
            let (rows, cols) = (array.rows, array.cols);
            let cells = array
                .values
                .iter()
                .enumerate()
                .map(|(position, value)| {
                    if caught(value) {
                        other.broadcast(position / cols, position % cols)
                    } else {
                        value.clone()
                    }
                })
                .collect();
            block(ctx, rows, cols, cells)
        }
    }
}

/// aggregates over a computed block; with no block the scalar builtin answers,
/// so references and literals behave exactly as before.
fn aggregate(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    name: &str,
    f: fn(&[f64]) -> CellValue,
) -> Value {
    fn empty_is_zero(name: &str) -> bool {
        matches!(name, "MIN" | "MAX")
    }
    let values: Vec<Value> = args.iter().map(|arg| evaluate_array(arg, ctx)).collect();
    if !values.iter().any(|value| matches!(value, Value::Array(_)))
        && let Some(scalar) = crate::functions::resolve(name)
    {
        return Value::Scalar(scalar.call(args, ctx));
    }
    let mut numbers = Vec::new();
    for (arg, value) in args.iter().zip(values) {
        match value {
            Value::Array(array) => {
                for value in &array.values {
                    match value {
                        CellValue::Number { value } => numbers.push(*value),
                        CellValue::Error { value } => return Value::error(*value),
                        _ => {}
                    }
                }
            }
            Value::Scalar(value) if as_area(arg, ctx).is_some() => match value {
                CellValue::Number { value } => numbers.push(value),
                CellValue::Error { value } => return Value::error(value),
                _ => {}
            },
            Value::Scalar(value) => match to_number(&value) {
                Ok(number) => numbers.push(number),
                Err(_) if matches!(value, CellValue::Empty) => {}
                Err(error) => return Value::error(error),
            },
        }
    }
    if numbers.is_empty() && empty_is_zero(name) {
        return Value::Scalar(num(0.0));
    }
    Value::Scalar(f(&numbers))
}

fn sum(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    aggregate(args, ctx, "SUM", |values| num(values.iter().sum()))
}

fn product(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    aggregate(args, ctx, "PRODUCT", |values| {
        num(values.iter().product::<f64>())
    })
}

fn average(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    aggregate(args, ctx, "AVERAGE", |values| {
        if values.is_empty() {
            err(ErrorValue::Div0)
        } else {
            num(values.iter().sum::<f64>() / values.len() as f64)
        }
    })
}

fn min(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    aggregate(args, ctx, "MIN", |values| {
        num(values.iter().copied().fold(f64::INFINITY, f64::min))
    })
}

fn max(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    aggregate(args, ctx, "MAX", |values| {
        num(values.iter().copied().fold(f64::NEG_INFINITY, f64::max))
    })
}

fn count(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    aggregate(args, ctx, "COUNT", |values| num(values.len() as f64))
}

/// COUNTA counts every non-blank cell, so it reads the block itself rather than
/// the numbers `aggregate` extracts.
fn counta(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let values: Vec<Value> = args.iter().map(|arg| evaluate_array(arg, ctx)).collect();
    if !values.iter().any(|value| matches!(value, Value::Array(_)))
        && let Some(scalar) = crate::functions::resolve("COUNTA")
    {
        return Value::Scalar(scalar.call(args, ctx));
    }
    let mut total = 0usize;
    for value in values {
        match value {
            Value::Array(array) => total += array.values.iter().filter(|v| !blank(v)).count(),
            Value::Scalar(value) => total += usize::from(!blank(&value)),
        }
    }
    Value::Scalar(num(total as f64))
}

/// MATCH over a computed block; a plain reference keeps the scalar path, which
/// can stop early on a big range.
fn match_(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    if args.len() >= 2
        && as_area(&args[1], ctx).is_some()
        && let Some(scalar) = crate::functions::resolve("MATCH")
    {
        return Value::Scalar(scalar.call(args, ctx));
    }
    result((|| {
        if args.len() < 2 || args.len() > 3 {
            return Err(ErrorValue::Value);
        }
        let target = match evaluate_array(&args[0], ctx).into_scalar() {
            CellValue::Error { value } => return Err(value),
            value => value,
        };
        let data = argument(args, ctx, 1)?;
        let kind = optional_number(args, ctx, 2, 1.0)?.trunc();
        let mut best = None;
        for (position, value) in data.values.iter().enumerate() {
            let ordering = cmp_values(value, &target);
            let hit = match kind {
                0.0 => ordering == std::cmp::Ordering::Equal,
                k if k > 0.0 => ordering != std::cmp::Ordering::Greater,
                _ => ordering != std::cmp::Ordering::Less,
            };
            if !hit {
                continue;
            }
            best = Some(position + 1);
            if kind == 0.0 {
                break;
            }
        }
        best.map(|position| Value::Scalar(num(position as f64)))
            .ok_or(ErrorValue::NA)
    })())
}

/// `ROW`/`COLUMN` over a multi-cell reference give its indices as a block.
fn row(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    indices(args, ctx, "ROW", true)
}

fn column(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    indices(args, ctx, "COLUMN", false)
}

fn indices(args: &[Expr], ctx: &EvalContext<'_>, name: &str, by_row: bool) -> Value {
    let area = match args {
        [arg] => as_array_area(arg, ctx),
        _ => None,
    };
    let Some(area) = area.filter(|area| if by_row { area.rows } else { area.cols } > 1) else {
        return match crate::functions::resolve(name) {
            Some(scalar) => Value::Scalar(scalar.call(args, ctx)),
            None => Value::error(ErrorValue::Name),
        };
    };
    if by_row {
        let cells = (0..area.rows)
            .map(|row| num(f64::from(area.start.row) + row as f64 + 1.0))
            .collect();
        block(ctx, area.rows, 1, cells)
    } else {
        let cells = (0..area.cols)
            .map(|col| num(f64::from(area.start.col) + col as f64 + 1.0))
            .collect();
        block(ctx, 1, area.cols, cells)
    }
}

/// MMULT(a, b): matrix product; the inner dimensions must agree and every cell
/// must be numeric.
fn mmult(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() != 2 {
            return Err(ErrorValue::Value);
        }
        let left = argument(args, ctx, 0)?;
        let right = argument(args, ctx, 1)?;
        if left.cols != right.rows {
            return Err(ErrorValue::Value);
        }
        let count = output_cells(left.rows, right.cols)?;
        let mut cells = Vec::with_capacity(count);
        for row in 0..left.rows {
            for col in 0..right.cols {
                let mut total = 0.0;
                for inner in 0..left.cols {
                    total += to_number(&left.at(row, inner))? * to_number(&right.at(inner, col))?;
                }
                cells.push(num(total));
            }
        }
        Ok(block(ctx, left.rows, right.cols, cells))
    })())
}

/// SUMPRODUCT: elementwise product of every argument, summed; non-numeric
/// cells count as zero.
fn sumproduct(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let values: Vec<Value> = args.iter().map(|arg| evaluate_array(arg, ctx)).collect();
    if !values.iter().any(|value| matches!(value, Value::Array(_)))
        && let Some(scalar) = crate::functions::resolve("SUMPRODUCT")
    {
        return Value::Scalar(scalar.call(args, ctx));
    }
    let mut rows = 1usize;
    let mut cols = 1usize;
    for value in &values {
        let (r, c) = value.dims();
        rows = rows.max(r);
        cols = cols.max(c);
    }
    let mut total = 0.0;
    for row in 0..rows {
        for col in 0..cols {
            let mut product = 1.0;
            for value in &values {
                product *= match value.broadcast(row, col) {
                    CellValue::Number { value } => value,
                    CellValue::Bool { value } => f64::from(value),
                    CellValue::Error { value } => return Value::error(value),
                    _ => 0.0,
                };
            }
            total += product;
        }
    }
    Value::Scalar(num(total))
}

// ------------------------------------------------------------------ spill

/// where an array formula's result lands and what it holds.
pub struct Spill {
    pub range: CellRange,
    pub values: Vec<CellValue>,
}

/// place a result at `anchor`: a block fills its own rectangle, a single value
/// repeats across the rectangle the file recorded, as ctrl-shift-enter does.
pub fn spill_at(anchor: CellRef, authored: Option<CellRange>, value: Value) -> Spill {
    let (rows, cols, values) = match value {
        Value::Array(array) => (array.rows, array.cols, array.values),
        Value::Scalar(value) => {
            let (rows, cols) = match authored {
                Some(range) => (
                    (range.end.row.saturating_sub(range.start.row) + 1) as usize,
                    (range.end.col.saturating_sub(range.start.col) + 1) as usize,
                ),
                None => (1, 1),
            };
            (rows, cols, vec![value; rows.saturating_mul(cols).max(1)])
        }
    };
    let rows = rows.clamp(1, MAX_SPILL_CELLS);
    let cols = cols.clamp(1, MAX_SPILL_CELLS.div_euclid(rows).max(1));
    let end = CellRef::new(
        anchor
            .row
            .saturating_add(rows as u32 - 1)
            .min(xlsx_model::MAX_ROWS - 1),
        anchor
            .col
            .saturating_add(cols as u32 - 1)
            .min(xlsx_model::MAX_COLS - 1),
    );
    let range = CellRange::new(CellRef::new(anchor.row, anchor.col), end);
    let rows = (range.end.row - range.start.row + 1) as usize;
    let cols = (range.end.col - range.start.col + 1) as usize;
    let mut out = Vec::with_capacity(rows * cols);
    for row in 0..rows {
        for col in 0..cols {
            out.push(
                values
                    .get(row * cols + col)
                    .cloned()
                    .unwrap_or(CellValue::Empty),
            );
        }
    }
    Spill { range, values: out }
}

/// evaluate a formula as an array formula and lay its result out from `anchor`.
pub fn evaluate_spill(
    expr: &Expr,
    ctx: &EvalContext<'_>,
    anchor: CellRef,
    authored: Option<CellRange>,
) -> Spill {
    spill_at(anchor, authored, evaluate_array(expr, ctx))
}

/// the single value an array formula shows in its anchor when it does not
/// spill, used where only one cell is being asked about.
pub fn evaluate_scalar(expr: &Expr, ctx: &EvalContext<'_>) -> CellValue {
    evaluate_array(expr, ctx).into_scalar()
}
