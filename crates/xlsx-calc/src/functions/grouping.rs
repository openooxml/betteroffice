//! GROUPBY and PIVOTBY, dynamic-array summaries grouped by row and column
//! fields, and TRIMRANGE, which cuts the blank edges off a range.

use std::collections::HashMap;
use std::collections::hash_map::Entry;
use std::rc::Rc;

use xlsx_model::{CellRef, CellValue, ErrorValue};

use crate::array::{
    Array, Cells, Lambda, MAX_ARRAY_CELLS, STEPS_PER_UNIT, Value, area_values, argument,
    charge_sort, charge_steps, copied, evaluate_array, invoke, optional, optional_block,
    optional_number, order_values, plain, result, step,
};
use crate::eval::{Area, EvalContext, resolve_area, text, to_bool, to_number};
use crate::parser::Expr;

/// the aggregation argument as a list of lambdas, and whether a vector of
/// them lays its results out across columns (`HSTACK`) or down rows.
fn callbacks(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
) -> Result<(bool, Vec<Rc<Lambda>>), ErrorValue> {
    let list = match args.get(index).map(|arg| evaluate_array(arg, ctx)) {
        Some(Value::Lambda(lambda)) => match &lambda.vector {
            Some((by_col, list)) => (*by_col, list.clone()),
            None => (true, vec![lambda]),
        },
        Some(Value::Scalar(CellValue::Error { value })) => return Err(value),
        _ => return Err(ErrorValue::Value),
    };
    if list.1.is_empty() || list.1.iter().any(|l| !matches!(l.params.len(), 1 | 2)) {
        return Err(ErrorValue::Value);
    }
    Ok(list)
}

/// a key under which equal field values group: text ignores case, as in
/// Excel, and carries its length, so no text can run into the next field.
fn group_identity(slice: &[CellValue]) -> String {
    let mut key = String::with_capacity(slice.len() * 8);
    for value in slice {
        match value {
            CellValue::Empty => key.push_str("z:"),
            CellValue::Number { value } => {
                key.push_str("n:");
                key.push_str(&crate::eval::exact_number_text(*value + 0.0));
            }
            CellValue::Text { value } => {
                let lower = value.to_lowercase();
                key.push('t');
                key.push_str(&lower.len().to_string());
                key.push(':');
                key.push_str(&lower);
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

/// the automatic header rule: the first row of `values` is text and the
/// second holds a number.
fn looks_like_header(block: &Array) -> bool {
    if block.rows < 2 {
        return false;
    }
    let row0 = block.row_values(0);
    let row1 = block.row_values(1);
    let row0_text = row0.iter().any(|v| matches!(v, CellValue::Text { .. }))
        && !row0
            .iter()
            .any(|v| matches!(v, CellValue::Number { .. } | CellValue::Bool { .. }));
    let row1_non_text = row1
        .iter()
        .any(|v| matches!(v, CellValue::Number { .. } | CellValue::Bool { .. }));
    row0_text && row1_non_text
}

/// `field_headers` as (the data has a header row, the output shows headers).
fn resolve_headers(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    values: &Array,
    group_levels: usize,
) -> Result<(bool, bool), ErrorValue> {
    match optional(args, ctx, index)? {
        None => {
            let has = looks_like_header(values);
            Ok((has, has && group_levels > 1))
        }
        Some(value) => match to_number(&value)?.trunc() as i64 {
            0 => Ok((false, false)),
            1 => Ok((true, false)),
            2 => Ok((false, true)),
            3 => Ok((true, true)),
            _ => Err(ErrorValue::Value),
        },
    }
}

/// `total_depth`, defaulting to a grand total as Excel does; a table
/// relationship allows no subtotals.
fn resolve_total_depth(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    field_cols: usize,
    field_relationship: i64,
) -> Result<i64, ErrorValue> {
    match optional(args, ctx, index)? {
        None => Ok(i64::from(field_cols > 0)),
        Some(value) => {
            let depth = to_number(&value)?.trunc() as i64;
            if depth.unsigned_abs() > field_cols as u64
                || (field_relationship == 1 && depth.unsigned_abs() >= 2)
            {
                return Err(ErrorValue::Value);
            }
            Ok(depth)
        }
    }
}

/// the header row for a block: its own first row, or generated names.
fn header_labels(
    block: &Array,
    has_headers: bool,
    prefix: &str,
    ctx: &EvalContext<'_>,
) -> Vec<CellValue> {
    (0..block.cols)
        .map(|col| {
            if has_headers {
                match block.at(0, col) {
                    CellValue::Empty => text(ctx, String::new()),
                    other => other,
                }
            } else {
                text(ctx, format!("{prefix} {}", col + 1))
            }
        })
        .collect()
}

/// value headers repeated once per aggregation laid out across columns.
fn expanded_val_headers(
    values: &Array,
    has_headers: bool,
    col_funcs: usize,
    ctx: &EvalContext<'_>,
) -> Vec<CellValue> {
    let base = header_labels(values, has_headers, "Value", ctx);
    if col_funcs == 1 {
        return base;
    }
    let mut out = Vec::with_capacity(base.len() * col_funcs);
    for label in base {
        for _ in 0..col_funcs {
            out.push(label.clone());
        }
    }
    out
}

/// the aggregation names shown above each value column.
fn calc_headers(lambdas: &[Rc<Lambda>], val_cols: usize, ctx: &EvalContext<'_>) -> Vec<CellValue> {
    let mut out = Vec::with_capacity(val_cols * lambdas.len());
    for _ in 0..val_cols {
        for lambda in lambdas {
            out.push(text(ctx, lambda.name.clone().unwrap_or_default()));
        }
    }
    out
}

/// the data rows `filter_array` keeps; none left is #CALC!.
fn kept_data_rows(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    filter_index: usize,
    total_rows: usize,
    data_start: usize,
) -> Result<Vec<usize>, ErrorValue> {
    if data_start >= total_rows {
        return Err(ErrorValue::Calc);
    }
    charge_steps(ctx, (total_rows - data_start) as u64)?;
    let Some(filter) = optional_block(args, ctx, filter_index)? else {
        return Ok((data_start..total_rows).collect());
    };
    if filter.cols != 1 || filter.rows != total_rows {
        return Err(ErrorValue::Value);
    }
    let mut kept = Vec::with_capacity(total_rows - data_start);
    for row in data_start..total_rows {
        if to_bool(&filter.values[row])? {
            kept.push(row);
        }
    }
    if kept.is_empty() {
        return Err(ErrorValue::Calc);
    }
    Ok(kept)
}

/// `sort_order` as (output column, descending) pairs; the default sorts by
/// every field ascending.
fn parse_sort_specs(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    field_cols: usize,
    val_cols: usize,
) -> Result<Vec<(usize, bool)>, ErrorValue> {
    let Some(block) = optional_block(args, ctx, index)? else {
        return Ok((0..field_cols).map(|col| (col, false)).collect());
    };
    if block.rows > 1 && block.cols > 1 {
        return Err(ErrorValue::Value);
    }
    let max_col = field_cols + val_cols;
    let mut specs = Vec::with_capacity(block.values.len());
    for cell in &block.values {
        let raw = to_number(cell)?.trunc() as i64;
        let col = raw.unsigned_abs();
        if raw == 0 || col > max_col as u64 {
            return Err(ErrorValue::Value);
        }
        specs.push((col as usize - 1, raw < 0));
    }
    if specs.is_empty() {
        return Err(ErrorValue::Value);
    }
    Ok(specs)
}

/// one value column restricted to `rows`, as a column block; the text of
/// each copied cell is paid for before it is copied.
fn slice_column(
    ctx: &EvalContext<'_>,
    values: &Array,
    rows: &[usize],
    col: usize,
) -> Result<Value, ErrorValue> {
    let mut out = Vec::with_capacity(rows.len());
    for &row in rows {
        let cell = &values.values[row * values.cols + col];
        if !ctx.keep_text(cell) {
            return Err(ErrorValue::Num);
        }
        out.push(copied(cell));
    }
    Ok(Value::Array(Array::new(rows.len(), 1, out)?))
}

/// one distinct combination of field values and the rows that carry it.
struct GroupEntry {
    key: Vec<CellValue>,
    rows: Vec<usize>,
    vals: Vec<CellValue>,
}

/// a line of the row or column axis: a group or a total.
enum AxisItem {
    Group(GroupEntry),
    Total {
        depth: usize,
        label_prefix: Vec<CellValue>,
        rows: Vec<usize>,
        vals: Vec<CellValue>,
        grand_label: &'static str,
    },
}

impl AxisItem {
    fn rows(&self) -> &[usize] {
        match self {
            Self::Group(group) => &group.rows,
            Self::Total { rows, .. } => rows,
        }
    }

    fn parent_prefix(&self) -> &[CellValue] {
        match self {
            Self::Group(group) => &group.key[..group.key.len().saturating_sub(1)],
            Self::Total {
                depth,
                label_prefix,
                ..
            } => &label_prefix[..depth.saturating_sub(1)],
        }
    }

    fn label_at(&self, col: usize, ctx: &EvalContext<'_>) -> CellValue {
        match self {
            Self::Group(group) => group.key[col].clone(),
            Self::Total {
                depth,
                label_prefix,
                grand_label,
                ..
            } => {
                if col < *depth {
                    label_prefix[col].clone()
                } else if *depth == 0 && col == 0 {
                    text(ctx, (*grand_label).to_string())
                } else {
                    text(ctx, String::new())
                }
            }
        }
    }
}

/// applies the aggregation lambdas to row subsets of `values`.
struct GroupAggregator<'a> {
    lambdas: &'a [Rc<Lambda>],
    by_col: bool,
    values: &'a Array,
    kept_rows: &'a [usize],
    ctx: &'a EvalContext<'a>,
}

impl GroupAggregator<'_> {
    fn eval_cell_with(
        &self,
        lambda: &Lambda,
        subset_rows: &[usize],
        total_rows: &[usize],
        col: usize,
    ) -> Result<CellValue, ErrorValue> {
        step();
        charge_steps(self.ctx, (subset_rows.len() as u64).max(STEPS_PER_UNIT))?;
        let subset = slice_column(self.ctx, self.values, subset_rows, col)?;
        let argv = match lambda.params.len() {
            1 => vec![plain(subset)],
            2 => {
                charge_steps(self.ctx, (total_rows.len() as u64).max(STEPS_PER_UNIT))?;
                vec![
                    plain(subset),
                    plain(slice_column(self.ctx, self.values, total_rows, col)?),
                ]
            }
            _ => return Err(ErrorValue::Value),
        };
        let out = invoke(lambda, argv, self.ctx);
        self.ctx.within_budget()?;
        Ok(out.into_scalar())
    }

    fn eval_all_cols(
        &self,
        subset_rows: &[usize],
        total_rows: &[usize],
    ) -> Result<Vec<CellValue>, ErrorValue> {
        let mut out = Vec::with_capacity(self.values.cols * self.lambdas.len());
        if self.by_col {
            for col in 0..self.values.cols {
                for lambda in self.lambdas {
                    out.push(self.eval_cell_with(lambda, subset_rows, total_rows, col)?);
                }
            }
        } else {
            for lambda in self.lambdas {
                for col in 0..self.values.cols {
                    out.push(self.eval_cell_with(lambda, subset_rows, total_rows, col)?);
                }
            }
        }
        Ok(out)
    }
}

/// how one axis groups, totals and sorts.
struct AxisConfig<'a> {
    field_cols: usize,
    total_depth: i64,
    field_relationship: i64,
    sort_specs: &'a [(usize, bool)],
    need_group_vals: bool,
    max_groups: usize,
}

/// the distinct field combinations among the kept rows, in first-seen order.
fn collect_groups(
    fields: &Array,
    kept_rows: &[usize],
    ctx: &EvalContext<'_>,
) -> Result<Vec<GroupEntry>, ErrorValue> {
    charge_steps(
        ctx,
        (kept_rows.len() as u64).saturating_mul(fields.cols.max(1) as u64),
    )?;
    let mut seen: HashMap<String, usize> = HashMap::new();
    let mut groups: Vec<GroupEntry> = Vec::new();
    for &row in kept_rows {
        step();
        let slice = fields.row_values(row);
        let key = group_identity(slice);
        match seen.entry(key) {
            Entry::Occupied(entry) => groups[*entry.get()].rows.push(row),
            Entry::Vacant(entry) => {
                entry.insert(groups.len());
                groups.push(GroupEntry {
                    key: slice.iter().map(copied).collect(),
                    rows: vec![row],
                    vals: Vec::new(),
                });
            }
        }
    }
    Ok(groups)
}

/// the sort order of two groups, falling back to their fields ascending.
fn compare_group_entries(
    a: &GroupEntry,
    b: &GroupEntry,
    field_cols: usize,
    sort_specs: &[(usize, bool)],
) -> std::cmp::Ordering {
    for &(col, desc) in sort_specs {
        let ordering = if col < field_cols {
            order_values(&a.key[col], &b.key[col], desc)
        } else {
            let vc = col - field_cols;
            order_values(&a.vals[vc], &b.vals[vc], desc)
        };
        if ordering != std::cmp::Ordering::Equal {
            return ordering;
        }
    }
    for col in 0..field_cols {
        let ordering = order_values(&a.key[col], &b.key[col], false);
        if ordering != std::cmp::Ordering::Equal {
            return ordering;
        }
    }
    std::cmp::Ordering::Equal
}

/// groups nested level by level with subtotals placed above or below
/// each parent, sorting within every parent.
fn build_hierarchy(
    mut groups: Vec<GroupEntry>,
    depth: usize,
    agg: &GroupAggregator<'_>,
    cfg: &AxisConfig<'_>,
) -> Result<Vec<AxisItem>, ErrorValue> {
    if depth + 1 >= cfg.field_cols {
        charge_sort(agg.ctx, groups.len(), cfg.sort_specs.len())?;
        groups.sort_by(|a, b| {
            step();
            compare_group_entries(a, b, cfg.field_cols, cfg.sort_specs)
        });
        return Ok(groups.into_iter().map(AxisItem::Group).collect());
    }
    charge_steps(
        agg.ctx,
        (groups.len() as u64).saturating_mul(cfg.field_cols as u64),
    )?;
    struct Bucket {
        key_prefix: Vec<CellValue>,
        rows: Vec<usize>,
        vals: Vec<CellValue>,
        children: Vec<GroupEntry>,
    }
    let mut seen: HashMap<String, usize> = HashMap::new();
    let mut buckets: Vec<Bucket> = Vec::new();
    for group in groups {
        step();
        let id = group_identity(&group.key[depth..=depth]);
        match seen.entry(id) {
            Entry::Occupied(entry) => {
                let bucket = &mut buckets[*entry.get()];
                bucket.rows.extend_from_slice(&group.rows);
                bucket.children.push(group);
            }
            Entry::Vacant(entry) => {
                entry.insert(buckets.len());
                buckets.push(Bucket {
                    key_prefix: group.key[..=depth].to_vec(),
                    rows: group.rows.clone(),
                    vals: Vec::new(),
                    children: vec![group],
                });
            }
        }
    }
    let show_subtotals = depth + 1 < cfg.total_depth.unsigned_abs() as usize;
    let sort_uses_vals = cfg.sort_specs.iter().any(|&(col, _)| col >= cfg.field_cols);
    for bucket in &mut buckets {
        charge_sort(agg.ctx, bucket.rows.len(), 1)?;
        bucket.rows.sort_unstable();
        if (show_subtotals && cfg.need_group_vals) || sort_uses_vals {
            bucket.vals = agg.eval_all_cols(&bucket.rows, agg.kept_rows)?;
        }
    }
    charge_sort(agg.ctx, buckets.len(), cfg.sort_specs.len())?;
    buckets.sort_by(|a, b| {
        step();
        for &(col, desc) in cfg.sort_specs {
            let ordering = if col <= depth {
                order_values(&a.key_prefix[col], &b.key_prefix[col], desc)
            } else if col >= cfg.field_cols {
                let vc = col - cfg.field_cols;
                order_values(&a.vals[vc], &b.vals[vc], desc)
            } else {
                std::cmp::Ordering::Equal
            };
            if ordering != std::cmp::Ordering::Equal {
                return ordering;
            }
        }
        order_values(&a.key_prefix[depth], &b.key_prefix[depth], false)
    });
    let mut items = Vec::new();
    for bucket in buckets {
        let children = build_hierarchy(bucket.children, depth + 1, agg, cfg)?;
        if show_subtotals {
            let subtotal = AxisItem::Total {
                depth: depth + 1,
                label_prefix: bucket.key_prefix,
                rows: bucket.rows,
                vals: bucket.vals,
                grand_label: "",
            };
            if cfg.total_depth < 0 {
                items.push(subtotal);
                items.extend(children);
            } else {
                items.extend(children);
                items.push(subtotal);
            }
        } else {
            items.extend(children);
        }
    }
    Ok(items)
}

/// every line of one axis, grand total included.
fn build_axis_items(
    fields: Option<&Array>,
    agg: &GroupAggregator<'_>,
    cfg: &AxisConfig<'_>,
) -> Result<Vec<AxisItem>, ErrorValue> {
    let Some(fields) = fields else {
        let vals = if cfg.need_group_vals {
            agg.eval_all_cols(agg.kept_rows, agg.kept_rows)?
        } else {
            Vec::new()
        };
        return Ok(vec![AxisItem::Group(GroupEntry {
            key: Vec::new(),
            rows: agg.kept_rows.to_vec(),
            vals,
        })]);
    };
    let mut groups = collect_groups(fields, agg.kept_rows, agg.ctx)?;
    if groups.len() > cfg.max_groups {
        return Err(ErrorValue::Num);
    }
    let sort_uses_vals = cfg.sort_specs.iter().any(|&(col, _)| col >= cfg.field_cols);
    if cfg.need_group_vals || sort_uses_vals {
        for group in &mut groups {
            group.vals = agg.eval_all_cols(&group.rows, agg.kept_rows)?;
        }
    }
    let mut items = if cfg.field_relationship == 1 || cfg.field_cols <= 1 {
        charge_sort(agg.ctx, groups.len(), cfg.sort_specs.len())?;
        groups.sort_by(|a, b| {
            step();
            compare_group_entries(a, b, cfg.field_cols, cfg.sort_specs)
        });
        groups.into_iter().map(AxisItem::Group).collect()
    } else {
        build_hierarchy(groups, 0, agg, cfg)?
    };
    if cfg.total_depth != 0 {
        let grand_vals = if cfg.need_group_vals {
            agg.eval_all_cols(agg.kept_rows, agg.kept_rows)?
        } else {
            Vec::new()
        };
        let grand = AxisItem::Total {
            depth: 0,
            label_prefix: Vec::new(),
            rows: agg.kept_rows.to_vec(),
            vals: grand_vals,
            grand_label: if cfg.total_depth.unsigned_abs() >= 2 {
                "Grand Total"
            } else {
                "Total"
            },
        };
        if cfg.total_depth < 0 {
            items.insert(0, grand);
        } else {
            items.push(grand);
        }
    }
    Ok(items)
}

/// GROUPBY(row_fields, values, function, [field_headers], [total_depth],
/// [sort_order], [filter_array], [field_relationship]).
pub(crate) fn groupby(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() < 3 || args.len() > 8 {
            return Err(ErrorValue::Value);
        }
        let row_fields = argument(args, ctx, 0)?;
        let values = argument(args, ctx, 1)?;
        if row_fields.rows != values.rows {
            return Err(ErrorValue::Value);
        }
        let (by_col, lambdas) = callbacks(args, ctx, 2)?;
        let (has_headers, show_headers) = resolve_headers(args, ctx, 3, &values, row_fields.cols)?;
        let field_relationship = optional_number(args, ctx, 7, 0.0)?.trunc() as i64;
        if !matches!(field_relationship, 0 | 1) {
            return Err(ErrorValue::Value);
        }
        let total_depth = resolve_total_depth(args, ctx, 4, row_fields.cols, field_relationship)?;
        let col_funcs = if by_col { lambdas.len() } else { 1 };
        let row_funcs = if by_col { 1 } else { lambdas.len() };
        let val_width = values.cols.checked_mul(col_funcs).ok_or(ErrorValue::Num)?;
        let out_cols = row_fields
            .cols
            .checked_add(val_width)
            .ok_or(ErrorValue::Num)?;
        let sort_specs = parse_sort_specs(args, ctx, 5, row_fields.cols, val_width)?;
        let data_start = usize::from(has_headers);
        let kept_rows = kept_data_rows(args, ctx, 6, row_fields.rows, data_start)?;
        let agg = GroupAggregator {
            lambdas: &lambdas,
            by_col,
            values: &values,
            kept_rows: &kept_rows,
            ctx,
        };
        let cfg = AxisConfig {
            field_cols: row_fields.cols,
            total_depth,
            field_relationship,
            sort_specs: &sort_specs,
            need_group_vals: true,
            max_groups: MAX_ARRAY_CELLS / out_cols / row_funcs,
        };
        let items = build_axis_items(Some(&row_fields), &agg, &cfg)?;
        let calc_header_row = col_funcs > 1;
        let out_rows = items
            .len()
            .checked_mul(row_funcs)
            .and_then(|rows| rows.checked_add(usize::from(calc_header_row)))
            .and_then(|rows| rows.checked_add(usize::from(show_headers)))
            .ok_or(ErrorValue::Num)?;
        let mut cells = Cells::reserve(ctx, out_rows, out_cols)?;
        if calc_header_row {
            for _ in 0..row_fields.cols {
                cells.push(text(ctx, String::new()))?;
            }
            for label in calc_headers(&lambdas, values.cols, ctx) {
                cells.push(label)?;
            }
        }
        if show_headers {
            for label in header_labels(&row_fields, has_headers, "Row Field", ctx) {
                cells.push(label)?;
            }
            for label in expanded_val_headers(&values, has_headers, col_funcs, ctx) {
                cells.push(label)?;
            }
        }
        for item in &items {
            let vals = match item {
                AxisItem::Group(group) => &group.vals,
                AxisItem::Total { vals, .. } => vals,
            };
            for rf in 0..row_funcs {
                for col in 0..row_fields.cols {
                    cells.push(item.label_at(col, ctx))?;
                }
                for val in &vals[rf * val_width..(rf + 1) * val_width] {
                    cells.push(val.clone())?;
                }
            }
        }
        Ok(cells.block(out_rows, out_cols))
    })())
}

/// rows present in both ascending lists.
fn intersect_sorted(
    ctx: &EvalContext<'_>,
    a: &[usize],
    b: &[usize],
) -> Result<Vec<usize>, ErrorValue> {
    charge_steps(ctx, (a.len() + b.len()) as u64)?;
    let mut out = Vec::new();
    let (mut i, mut j) = (0, 0);
    while i < a.len() && j < b.len() {
        step();
        match a[i].cmp(&b[j]) {
            std::cmp::Ordering::Less => i += 1,
            std::cmp::Ordering::Greater => j += 1,
            std::cmp::Ordering::Equal => {
                out.push(a[i]);
                i += 1;
                j += 1;
            }
        }
    }
    Ok(out)
}

/// the kept rows whose leading fields equal `prefix`.
fn matching_prefix_rows(
    ctx: &EvalContext<'_>,
    fields: Option<&Array>,
    kept_rows: &[usize],
    prefix: &[CellValue],
) -> Result<Vec<usize>, ErrorValue> {
    let Some(fields) = fields else {
        return Ok(kept_rows.to_vec());
    };
    if prefix.is_empty() {
        return Ok(kept_rows.to_vec());
    }
    charge_steps(
        ctx,
        (kept_rows.len() as u64).saturating_mul(prefix.len() as u64),
    )?;
    let target = group_identity(prefix);
    Ok(kept_rows
        .iter()
        .copied()
        .filter(|&row| {
            step();
            group_identity(&fields.row_values(row)[..prefix.len()]) == target
        })
        .collect())
}

/// PIVOTBY(row_fields, col_fields, values, function, [field_headers],
/// [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order],
/// [filter_array], [relative_to]).
pub(crate) fn pivotby(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    result((|| {
        if args.len() < 4 || args.len() > 11 {
            return Err(ErrorValue::Value);
        }
        let row_fields = optional_block(args, ctx, 0)?;
        let col_fields = optional_block(args, ctx, 1)?;
        let values = argument(args, ctx, 2)?;
        if row_fields.is_none() && col_fields.is_none() {
            return Err(ErrorValue::Value);
        }
        if row_fields.as_ref().is_some_and(|rf| rf.rows != values.rows)
            || col_fields.as_ref().is_some_and(|cf| cf.rows != values.rows)
        {
            return Err(ErrorValue::Value);
        }
        let row_field_cols = row_fields.as_ref().map_or(0, |rf| rf.cols);
        let col_field_cols = col_fields.as_ref().map_or(0, |cf| cf.cols);
        let (by_col, lambdas) = callbacks(args, ctx, 3)?;
        let (has_headers, show_headers) =
            resolve_headers(args, ctx, 4, &values, row_field_cols + col_field_cols)?;
        let row_total_depth = resolve_total_depth(args, ctx, 5, row_field_cols, 0)?;
        let col_total_depth = resolve_total_depth(args, ctx, 7, col_field_cols, 0)?;
        let col_funcs = if by_col { lambdas.len() } else { 1 };
        let row_funcs = if by_col { 1 } else { lambdas.len() };
        let val_width = values.cols.checked_mul(col_funcs).ok_or(ErrorValue::Num)?;
        let row_sort_specs = parse_sort_specs(args, ctx, 6, row_field_cols, val_width)?;
        let col_sort_specs = parse_sort_specs(args, ctx, 8, col_field_cols, val_width)?;
        let relative_to = optional_number(args, ctx, 10, 0.0)?.trunc() as i64;
        if !(0..=4).contains(&relative_to) {
            return Err(ErrorValue::Value);
        }
        let data_start = usize::from(has_headers);
        let kept_rows = kept_data_rows(args, ctx, 9, values.rows, data_start)?;
        let agg = GroupAggregator {
            lambdas: &lambdas,
            by_col,
            values: &values,
            kept_rows: &kept_rows,
            ctx,
        };
        let row_cfg = AxisConfig {
            field_cols: row_field_cols,
            total_depth: row_total_depth,
            field_relationship: 0,
            sort_specs: &row_sort_specs,
            need_group_vals: false,
            max_groups: MAX_ARRAY_CELLS / row_funcs / (row_field_cols + val_width),
        };
        let col_cfg = AxisConfig {
            field_cols: col_field_cols,
            total_depth: col_total_depth,
            field_relationship: 0,
            sort_specs: &col_sort_specs,
            need_group_vals: false,
            max_groups: MAX_ARRAY_CELLS / val_width,
        };
        let row_items = build_axis_items(row_fields.as_ref(), &agg, &row_cfg)?;
        let col_items = build_axis_items(col_fields.as_ref(), &agg, &col_cfg)?;
        let calc_header_row = col_funcs > 1;
        let field_header_row = show_headers && col_field_cols > 0;
        let header_rows = col_field_cols
            + usize::from(field_header_row)
            + usize::from(calc_header_row)
            + usize::from(show_headers);
        let out_rows = row_items
            .len()
            .checked_mul(row_funcs)
            .and_then(|rows| rows.checked_add(header_rows))
            .ok_or(ErrorValue::Num)?;
        let col_width = col_items
            .len()
            .checked_mul(val_width)
            .ok_or(ErrorValue::Num)?;
        let out_cols = row_field_cols
            .checked_add(col_width)
            .ok_or(ErrorValue::Num)?;
        let mut cells = Cells::reserve(ctx, out_rows, out_cols)?;
        let row_member_sum: u64 = row_items.iter().map(|i| i.rows().len() as u64).sum();
        let col_member_sum: u64 = col_items.iter().map(|i| i.rows().len() as u64).sum();
        let intersect_steps = row_member_sum
            .saturating_mul(col_items.len() as u64)
            .saturating_add(col_member_sum.saturating_mul(row_items.len() as u64));
        charge_steps(ctx, intersect_steps)?;
        let col_denominators: Vec<Vec<usize>> = if relative_to == 3 {
            col_items
                .iter()
                .map(|item| {
                    matching_prefix_rows(ctx, col_fields.as_ref(), &kept_rows, item.parent_prefix())
                })
                .collect::<Result<_, _>>()?
        } else {
            Vec::new()
        };
        let row_denominators: Vec<Vec<usize>> = if relative_to == 4 {
            row_items
                .iter()
                .map(|item| {
                    matching_prefix_rows(ctx, row_fields.as_ref(), &kept_rows, item.parent_prefix())
                })
                .collect::<Result<_, _>>()?
        } else {
            Vec::new()
        };
        let row_headers = row_fields.as_ref().map_or_else(Vec::new, |rf| {
            header_labels(rf, has_headers, "Row Field", ctx)
        });
        let val_headers = expanded_val_headers(&values, has_headers, col_funcs, ctx);
        if let (true, Some(cf)) = (field_header_row, col_fields.as_ref()) {
            for _ in 0..row_field_cols {
                cells.push(text(ctx, String::new()))?;
            }
            let mut col_headers = header_labels(cf, has_headers, "Col Field", ctx).into_iter();
            for _ in 0..col_width {
                cells.push(
                    col_headers
                        .next()
                        .unwrap_or_else(|| text(ctx, String::new())),
                )?;
            }
        }
        for hr in 0..col_field_cols {
            for _ in 0..row_field_cols {
                cells.push(text(ctx, String::new()))?;
            }
            for col_item in &col_items {
                cells.push(col_item.label_at(hr, ctx))?;
                for _ in 1..val_width {
                    cells.push(text(ctx, String::new()))?;
                }
            }
        }
        if calc_header_row {
            for _ in 0..row_field_cols {
                cells.push(text(ctx, String::new()))?;
            }
            let fn_headers = calc_headers(&lambdas, values.cols, ctx);
            for _ in &col_items {
                for label in &fn_headers {
                    cells.push(label.clone())?;
                }
            }
        }
        if show_headers {
            for row_header in &row_headers {
                cells.push(row_header.clone())?;
            }
            for _ in &col_items {
                for val_header in &val_headers {
                    cells.push(val_header.clone())?;
                }
            }
        }
        for (ri, row_item) in row_items.iter().enumerate() {
            for rf in 0..row_funcs {
                for rc in 0..row_field_cols {
                    cells.push(row_item.label_at(rc, ctx))?;
                }
                for (ci, col_item) in col_items.iter().enumerate() {
                    let subset = intersect_sorted(ctx, row_item.rows(), col_item.rows())?;
                    if subset.is_empty() {
                        for _ in 0..val_width {
                            cells.push(text(ctx, String::new()))?;
                        }
                        continue;
                    }
                    let parent_denom;
                    let denom: &[usize] = match relative_to {
                        0 => col_item.rows(),
                        1 => row_item.rows(),
                        2 => &kept_rows,
                        3 => {
                            parent_denom =
                                intersect_sorted(ctx, row_item.rows(), &col_denominators[ci])?;
                            &parent_denom
                        }
                        _ => {
                            parent_denom =
                                intersect_sorted(ctx, &row_denominators[ri], col_item.rows())?;
                            &parent_denom
                        }
                    };
                    if by_col {
                        for vc in 0..values.cols {
                            for lambda in &lambdas {
                                cells.push(agg.eval_cell_with(lambda, &subset, denom, vc)?)?;
                            }
                        }
                    } else {
                        for vc in 0..values.cols {
                            cells.push(agg.eval_cell_with(&lambdas[rf], &subset, denom, vc)?)?;
                        }
                    }
                }
            }
        }
        Ok(cells.block(out_rows, out_cols))
    })())
}

/// TRIMRANGE's edge modes: 0 keeps the edges, 1 trims leading blanks, 2
/// trailing blanks, 3 both.
const TRIM_LEADING: u8 = 1;
const TRIM_TRAILING: u8 = 2;
const TRIM_BOTH: u8 = 3;

/// the fixed modes of the trim-ref forms Excel writes to a file for `.:.`,
/// `.:` and `:.`, or `None` when `name` is not one of them.
fn trim_ref_mode(upper: &str) -> Option<u8> {
    match upper {
        "_TRO_ALL" => Some(TRIM_BOTH),
        "_TRO_LEADING" => Some(TRIM_LEADING),
        "_TRO_TRAILING" => Some(TRIM_TRAILING),
        _ => None,
    }
}

/// whether `name` is TRIMRANGE or one of its trim-ref forms.
pub(crate) fn trims(name: &str) -> bool {
    let upper = crate::functions::bare_name(name).to_ascii_uppercase();
    upper == "TRIMRANGE" || trim_ref_mode(&upper).is_some()
}

/// the (rows, columns) modes a call asks for. a mode outside 0..=3 is #VALUE!.
fn trim_modes(name: &str, args: &[Expr], ctx: &EvalContext<'_>) -> Result<(u8, u8), ErrorValue> {
    let upper = crate::functions::bare_name(name).to_ascii_uppercase();
    if let Some(mode) = trim_ref_mode(&upper) {
        return if args.len() == 1 {
            Ok((mode, mode))
        } else {
            Err(ErrorValue::Value)
        };
    }
    if args.is_empty() || args.len() > 3 {
        return Err(ErrorValue::Value);
    }
    let mode = |index: usize| -> Result<u8, ErrorValue> {
        let value = optional_number(args, ctx, index, f64::from(TRIM_BOTH))?.trunc();
        if (0.0..=f64::from(TRIM_BOTH)).contains(&value) {
            Ok(value as u8)
        } else {
            Err(ErrorValue::Value)
        }
    };
    Ok((mode(1)?, mode(2)?))
}

/// the lines of `len` a mode keeps around the non-blank span `first..=last`.
fn kept_span(mode: u8, len: usize, (first, last): (usize, usize)) -> (usize, usize) {
    let start = if mode & TRIM_LEADING != 0 { first } else { 0 };
    let end = if mode & TRIM_TRAILING != 0 {
        last
    } else {
        len - 1
    };
    (start, end)
}

/// the first and last index of the lines a block uses along one axis.
type Span = (usize, usize);

/// the extent of the non-blank values of a row-major block, as its row span
/// and its column span; the scan is charged first and a block without a
/// value is #REF!.
fn used_lines(
    ctx: &EvalContext<'_>,
    rows: usize,
    cols: usize,
    blank: impl Fn(usize, usize) -> bool,
) -> Result<(Span, Span), ErrorValue> {
    charge_steps(ctx, (rows as u64).saturating_mul(cols as u64))?;
    let mut extent: Option<(Span, Span)> = None;
    for row in 0..rows {
        for col in 0..cols {
            if blank(row, col) {
                continue;
            }
            extent = Some(match extent {
                None => ((row, row), (col, col)),
                Some(((r0, r1), (c0, c1))) => {
                    ((r0.min(row), r1.max(row)), (c0.min(col), c1.max(col)))
                }
            });
        }
    }
    extent.ok_or(ErrorValue::Ref)
}

/// the rectangle TRIMRANGE leaves of a reference, or `None` when the first
/// argument is no reference and the call trims a computed block instead. a
/// range with nothing in it is #REF!. only the part of the range inside the
/// sheet's used extent is read, as everything past it is blank.
pub(crate) fn trim_area(
    name: &str,
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Option<Result<Area, ErrorValue>> {
    let area = match resolve_area(args.first()?, ctx)? {
        Ok(area) => area,
        Err(error) => return Some(Err(error)),
    };
    Some((|| {
        let (row_mode, col_mode) = trim_modes(name, args, ctx)?;
        if row_mode == 0 && col_mode == 0 {
            return Ok(area);
        }
        let used_rows = (ctx.provider.used_rows(area.sheet) as usize)
            .saturating_sub(area.start.row as usize)
            .min(area.rows);
        let used_cols = (ctx.provider.used_cols(area.sheet) as usize)
            .saturating_sub(area.start.col as usize)
            .min(area.cols);
        if used_rows == 0 || used_cols == 0 {
            return Err(ErrorValue::Ref);
        }
        let read = Area {
            rows: used_rows,
            cols: used_cols,
            ..area
        };
        let cells = read.cells_ref(ctx)?;
        let (row_span, col_span) = used_lines(ctx, used_rows, used_cols, |row, col| {
            matches!(*cells[row * used_cols + col], CellValue::Empty)
        })?;
        let (top, bottom) = kept_span(row_mode, area.rows, row_span);
        let (left, right) = kept_span(col_mode, area.cols, col_span);
        Ok(Area {
            sheet: area.sheet,
            start: CellRef::new(area.start.row + top as u32, area.start.col + left as u32),
            rows: bottom - top + 1,
            cols: right - left + 1,
        })
    })())
}

/// TRIMRANGE over a reference or a computed block, under the modes `name`
/// carries.
fn trim(name: &str, args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    if let Some(area) = trim_area(name, args, ctx) {
        return match area {
            Ok(area) => area_values(&area, ctx),
            Err(error) => Value::error(error),
        };
    }
    result((|| {
        let (row_mode, col_mode) = trim_modes(name, args, ctx)?;
        let block = argument(args, ctx, 0)?;
        if row_mode == 0 && col_mode == 0 {
            return Ok(Value::Array(block));
        }
        let (row_span, col_span) = used_lines(ctx, block.rows, block.cols, |row, col| {
            matches!(block.values[row * block.cols + col], CellValue::Empty)
        })?;
        let (top, bottom) = kept_span(row_mode, block.rows, row_span);
        let (left, right) = kept_span(col_mode, block.cols, col_span);
        let (rows, cols) = (bottom - top + 1, right - left + 1);
        let mut cells = Cells::reserve(ctx, rows, cols)?;
        for row in top..=bottom {
            for col in left..=right {
                cells.push(block.at(row, col))?;
            }
        }
        Ok(cells.block(rows, cols))
    })())
}

/// TRIMRANGE(range, [trim_rows], [trim_cols]).
pub(crate) fn trimrange(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    trim("TRIMRANGE", args, ctx)
}

/// `range.:.range`, stored as `_xlfn._TRO_ALL(range)`.
pub(crate) fn trim_all(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    trim("_TRO_ALL", args, ctx)
}

/// `range.:range`, stored as `_xlfn._TRO_LEADING(range)`.
pub(crate) fn trim_leading(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    trim("_TRO_LEADING", args, ctx)
}

/// `range:.range`, stored as `_xlfn._TRO_TRAILING(range)`.
pub(crate) fn trim_trailing(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    trim("_TRO_TRAILING", args, ctx)
}

#[cfg(test)]
mod tests {
    use xlsx_model::{CellValue, ErrorValue, Sheet, SheetId, Workbook};

    use crate::array::{STEPS, evaluate_array};
    use crate::eval::EvalContext;
    use crate::parser::parse_formula;

    /// 1024 stacked functions leave room for 512 groups; 600 distinct keys
    /// are #NUM! before a single group is aggregated.
    #[test]
    fn too_many_distinct_keys_stop_before_any_group_is_aggregated() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Data"));
        let formula = parse_formula(
            "LET(f,REDUCE(SUM,SEQUENCE(10),LAMBDA(a,i,VSTACK(a,a))),k,SEQUENCE(600),GROUPBY(k,k,f,0,0))",
        )
        .unwrap();
        let context = EvalContext::new(&workbook, SheetId(0));
        STEPS.with(|steps| steps.set(0));
        let value = evaluate_array(&formula, &context).into_scalar();
        let steps = STEPS.with(|steps| steps.get());
        assert_eq!(
            value,
            CellValue::Error {
                value: ErrorValue::Num
            }
        );
        assert!(steps < 5_000, "{steps} steps");
    }
}
