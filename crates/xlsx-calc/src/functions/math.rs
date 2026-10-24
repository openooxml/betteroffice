//! math and trig functions. rounding follows excel: ROUND is half-away-from-zero,
//! ROUNDUP/ROUNDDOWN are directional, non-finite results are `#NUM!`.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, err, evaluate, num, to_number, to_text};
use crate::parser::Expr;

use super::criteria::{self, Criterion};
use super::{collect_numbers, finite, nth_int, nth_number};

pub(crate) fn sum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) => num(nums.iter().sum()),
        Err(e) => err(e),
    }
}

/// PRODUCT of every numeric argument; no numbers -> 0 (excel).
pub(crate) fn product(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) if nums.is_empty() => num(0.0),
        Ok(nums) => num(nums.iter().product()),
        Err(e) => err(e),
    }
}

/// SUMIF(range, criteria, [sum_range]). sum_range is anchored at its top-left
/// with the criteria range's shape, so a mismatched size still aligns.
pub(crate) fn sumif(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 && args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let crit_area = match crate::eval::required_area(&args[0], ctx) {
        Ok(a) => a,
        Err(error) => return err(error),
    };
    let criterion = criteria::criterion_from_arg(&args[1], ctx);
    let values = if args.len() == 3 { &args[2] } else { &args[0] };
    let Some(sum_area) = as_area(values, ctx) else {
        return err(ErrorValue::Value);
    };
    let (crit_area, sum_area) = criteria::single_pair(crit_area, sum_area, ctx);
    let pairs = [(crit_area, criterion)];
    sum_matching(&pairs, &sum_area, ctx)
}

/// SUMIFS(sum_range, crit_range1, crit1, ...). all ranges share dimensions.
pub(crate) fn sumifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    match criteria::collect_pairs(&args[1..], Some(&args[0]), ctx) {
        Some(criteria::Criteria {
            pairs,
            values: Some(sum_area),
            ..
        }) => sum_matching(&pairs, &sum_area, ctx),
        _ => err(ErrorValue::Value),
    }
}

fn sum_matching(
    pairs: &[(Area, Criterion)],
    value_area: &Area,
    ctx: &EvalContext<'_>,
) -> CellValue {
    let cols = pairs[0].0.cols;
    let mut total = 0.0;
    let indices = match criteria::matching_indices(pairs, ctx) {
        Ok(indices) => indices,
        Err(error) => return err(error),
    };
    for i in indices {
        let (r, c) = (i / cols, i % cols);
        match value_area.get_ref(ctx, r, c) {
            Ok(v) => {
                if let CellValue::Number { value } = *v {
                    total += value;
                }
            }
            Err(error) => return err(error),
        }
    }
    num(total)
}

/// SUMPRODUCT(array1, [array2], ...). element-wise product summed; non-numeric
/// cells count as 0. all arrays must share the same length.
pub(crate) fn sumproduct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() {
        return err(ErrorValue::Value);
    }
    if let Some(value) = streamed::sumproduct(args, ctx) {
        return value;
    }
    sumproduct_unstreamed(args, ctx)
}

fn sumproduct_unstreamed(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let areas: Vec<Option<Area>> = args.iter().map(|arg| as_area(arg, ctx)).collect();
    // references on different sheets are cut to one extent, so their cells
    // still pair up by position
    let mut cut = areas.clone();
    criteria::cut_references(&mut cut, ctx);
    let mut arrays: Vec<((usize, usize), Vec<f64>)> = Vec::with_capacity(args.len());
    for ((arg, area), cut) in args.iter().zip(areas).zip(cut) {
        match area.zip(cut) {
            Some((area, cut)) => {
                let values = match cut.cells_ref(ctx) {
                    Ok(values) => values,
                    Err(error) => return err(error),
                };
                let mut col = Vec::with_capacity(values.len());
                for v in values {
                    match v.as_ref() {
                        CellValue::Number { value } => col.push(*value),
                        CellValue::Error { value } => return err(*value),
                        _ => col.push(0.0),
                    }
                }
                arrays.push(((area.rows, area.cols), col));
            }
            // `--(range=key)` and the other computed operands are arrays, not
            // references, so they are read as blocks rather than coerced
            None => match crate::array::evaluate_array(arg, ctx) {
                crate::array::Value::Array(array) => {
                    let mut col = Vec::with_capacity(array.rows() * array.cols());
                    for row in 0..array.rows() {
                        for c in 0..array.cols() {
                            match array.at(row, c) {
                                CellValue::Number { value } => col.push(value),
                                CellValue::Error { value } => return err(value),
                                _ => col.push(0.0),
                            }
                        }
                    }
                    arrays.push(((array.rows(), array.cols()), col));
                }
                value => match to_number(&value.into_scalar()) {
                    Ok(n) => arrays.push(((1, 1), vec![n])),
                    Err(e) => return err(e),
                },
            },
        }
    }
    // excel pairs the operands by position, so they must share a shape: a
    // column and a row of the same length are not the same operand
    let shape = arrays[0].0;
    if arrays.iter().any(|(dims, _)| *dims != shape) {
        return err(ErrorValue::Value);
    }
    // past the extent a reference was cut to its cells are blank, so their
    // products are zero
    let len = arrays.iter().map(|(_, a)| a.len()).min().unwrap_or(0);
    let mut total = 0.0;
    for i in 0..len {
        let mut prod = 1.0;
        for (_, a) in &arrays {
            prod *= a[i];
        }
        total += prod;
    }
    num(total)
}

pub(crate) fn sumproduct_array(args: &[Expr], ctx: &EvalContext<'_>) -> Option<CellValue> {
    streamed::sumproduct_mode(args, ctx, true)
}

mod streamed {
    use std::collections::{HashMap, HashSet};

    use xlsx_model::{CellValue, ErrorValue, SheetId};

    use crate::eval::{
        Area, EvalContext, apply_binary, apply_percent, apply_unary, as_area, coerced, err,
        evaluate, normalize_provider_value,
    };
    use crate::parser::{BinaryOp, Expr, UnaryOp};

    enum Node {
        Leaf(usize),
        Scalar(usize),
        Unary(UnaryOp, Box<Node>),
        Percent(Box<Node>),
        Binary(BinaryOp, Box<Node>, Box<Node>),
    }

    struct Leaf {
        area: Area,
        shape: (usize, usize),
        read_rows: usize,
        charge_cells: u64,
        key: (SheetId, u32, u32, usize, usize),
    }

    struct Plan<'a> {
        root: Node,
        shape: (usize, usize),
        scalar_shape: (usize, usize),
        tail_start: Option<usize>,
        ops: u64,
        leaves: Vec<usize>,
        scalars: Vec<&'a Expr>,
        scalar_values: Vec<CellValue>,
    }

    #[cfg(test)]
    thread_local! {
        /// tests turn streaming off to compare against the unstreamed path.
        static UNSTREAMED: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    }

    #[cfg(test)]
    fn unstreamed() -> bool {
        UNSTREAMED.get()
    }

    #[cfg(not(test))]
    fn unstreamed() -> bool {
        false
    }

    pub(super) fn sumproduct(args: &[Expr], ctx: &EvalContext<'_>) -> Option<CellValue> {
        sumproduct_mode(args, ctx, false)
    }

    pub(super) fn sumproduct_mode(
        args: &[Expr],
        ctx: &EvalContext<'_>,
        array: bool,
    ) -> Option<CellValue> {
        if args.is_empty() || ctx.exhausted() || unstreamed() {
            return None;
        }
        let mut leaves: Vec<Leaf> = Vec::new();
        let mut plans = Vec::with_capacity(args.len());
        for arg in args {
            let bare = matches!(
                arg,
                Expr::Ref { .. }
                    | Expr::Range { .. }
                    | Expr::ColumnRange { .. }
                    | Expr::RowRange { .. }
            );
            let mut builder = Builder {
                ctx,
                leaves: &mut leaves,
                scalars: Vec::new(),
                used_rows: HashMap::new(),
                used_cols: HashMap::new(),
            };
            let (root, shape, ops) = builder.build(arg)?;
            let scalars = builder.scalars;
            let scalar_shape = if bare {
                let area = as_area(arg, ctx)?;
                (area.rows, area.cols)
            } else {
                shape
            };
            let mut leaf_ids = Vec::new();
            collect(&root, &mut leaf_ids);
            let tail_start = tail_start(&root, &leaves, shape.0);
            plans.push(Plan {
                root,
                shape,
                scalar_shape,
                tail_start,
                ops,
                leaves: leaf_ids,
                scalars,
                scalar_values: Vec::new(),
            });
        }
        if array {
            let shape = (
                plans.iter().map(|plan| plan.shape.0).max().unwrap_or(1),
                plans.iter().map(|plan| plan.shape.1).max().unwrap_or(1),
            );
            if shape == (1, 1) {
                return None;
            }
            if (shape.0 as u128) * (shape.1 as u128) > crate::array::MAX_ARRAY_CELLS as u128 {
                return None;
            }
            return run_array(plans, &leaves, shape, ctx);
        }
        let shape = plans[0].shape;
        let scalar_shape = plans[0].scalar_shape;
        if plans
            .iter()
            .any(|plan| plan.shape != shape || plan.scalar_shape != scalar_shape)
        {
            return None;
        }
        if shape == (1, 1) {
            return None;
        }
        if (shape.0 as u128) * (shape.1 as u128) > crate::array::MAX_ARRAY_CELLS as u128 {
            return None;
        }
        run(plans, &leaves, shape, ctx)
    }

    fn collect(node: &Node, out: &mut Vec<usize>) {
        match node {
            Node::Leaf(index) => {
                if !out.contains(index) {
                    out.push(*index);
                }
            }
            Node::Scalar(_) => {}
            Node::Unary(_, child) | Node::Percent(child) => collect(child, out),
            Node::Binary(_, lhs, rhs) => {
                collect(lhs, out);
                collect(rhs, out);
            }
        }
    }

    fn tail_start(node: &Node, leaves: &[Leaf], rows: usize) -> Option<usize> {
        if rows <= 1 {
            return None;
        }
        let mut ids = Vec::new();
        collect(node, &mut ids);
        let mut start = 0usize;
        for index in ids {
            let leaf = &leaves[index];
            if leaf.shape.0 == 1 {
                continue;
            }
            if leaf.shape.0 != rows {
                return None;
            }
            start = start.max(leaf.read_rows.min(rows));
        }
        Some(start)
    }

    /// whether the budget covers `units` cells and the text of every literal
    /// operand, spending nothing. a SUMPRODUCT it cannot cover takes the
    /// unstreamed path, which spends and fails as it always has, instead of
    /// draining what a later cell would have had.
    fn affordable(plans: &[Plan<'_>], units: u64, ctx: &EvalContext<'_>) -> bool {
        let bytes = plans
            .iter()
            .flat_map(|plan| &plan.scalars)
            .map(|expr| text_bytes(expr))
            .fold(0u64, u64::saturating_add);
        ctx.can_consume(units) && ctx.can_consume_text(bytes)
    }

    fn evaluate_literals(plan: &mut Plan<'_>, ctx: &EvalContext<'_>) {
        plan.scalar_values = {
            let _scope = ctx.array_scope();
            plan.scalars
                .iter()
                .map(|expr| evaluate(expr, ctx))
                .collect()
        };
    }

    /// the budget units a plan's arithmetic costs over `positions`; a bare
    /// range is paid for by its cells, as the unstreamed path charges it.
    fn step_cost(plan: &Plan<'_>, positions: u64) -> u64 {
        if plan.ops == 0 {
            return 0;
        }
        crate::array::step_units(plan.ops.saturating_add(1).saturating_mul(positions))
    }

    /// the budget units of adding the repeated tail row once per row past the
    /// streamed ones.
    fn tail_cost(rows: usize, streamed: usize, cols: usize) -> u64 {
        let repeats = rows.saturating_sub(streamed) as u64;
        crate::array::step_units(repeats.saturating_mul(cols as u64))
    }

    /// whether a plan is zero (or an error, which ends the sum first) on the
    /// tail row, where every range running past it is blank: a product with
    /// such a range as a factor is, and so is a quotient of one.
    fn zero_at_tail(node: &Node, leaves: &[Leaf]) -> bool {
        match node {
            Node::Leaf(index) => leaves[*index].shape.0 > 1,
            Node::Scalar(_) => false,
            Node::Unary(_, child) | Node::Percent(child) => zero_at_tail(child, leaves),
            Node::Binary(BinaryOp::Mul, lhs, rhs) => {
                zero_at_tail(lhs, leaves) || zero_at_tail(rhs, leaves)
            }
            Node::Binary(BinaryOp::Div, lhs, _) => zero_at_tail(lhs, leaves),
            Node::Binary(..) => false,
        }
    }

    /// read a plan's ranges leaf by leaf, each row by row, as the unstreamed
    /// path materialises them before any arithmetic, so a settle pass meets
    /// its first unsettled cell in the same order; the rows read them again.
    fn read_in_order(plan: &Plan<'_>, leaves: &[Leaf], ctx: &EvalContext<'_>) {
        for index in &plan.leaves {
            let leaf = &leaves[*index];
            for row in 0..leaf.read_rows {
                for col in 0..leaf.shape.1 {
                    drop(leaf.area.get_unmetered_ref(ctx, row, col));
                }
            }
        }
    }

    /// the text a literal operand charges when it is evaluated.
    fn text_bytes(expr: &Expr) -> u64 {
        match expr {
            Expr::Text(text) => text.len() as u64,
            Expr::Literal(value) => match normalize_provider_value(value.clone()) {
                CellValue::Text { value } => value.len() as u64,
                _ => 0,
            },
            _ => 0,
        }
    }

    fn run(
        mut plans: Vec<Plan<'_>>,
        leaves: &[Leaf],
        shape: (usize, usize),
        ctx: &EvalContext<'_>,
    ) -> Option<CellValue> {
        let (rows, cols) = shape;
        let max_main = plans
            .iter()
            .map(|plan| plan.tail_start.unwrap_or(rows))
            .max()
            .unwrap_or(rows);
        let mut charged: HashSet<(SheetId, u32, u32, usize, usize)> = HashSet::new();
        let mut costs = Vec::with_capacity(plans.len());
        for plan in &plans {
            let mut cost = 0u64;
            for index in &plan.leaves {
                let leaf = &leaves[*index];
                if charged.insert(leaf.key) {
                    cost = cost.saturating_add(leaf.charge_cells);
                }
            }
            let tail_rows = plan.tail_start.is_some_and(|start| start < rows);
            let positions = (max_main as u64)
                .saturating_mul(cols as u64)
                .saturating_add(if tail_rows { cols as u64 } else { 0 });
            costs.push(cost.saturating_add(step_cost(plan, positions)));
        }
        // a tail known to add only zeros is added once, so it costs nothing
        let zero = plans.iter().any(|plan| zero_at_tail(&plan.root, leaves));
        let tail_units = if zero {
            0
        } else {
            tail_cost(rows, max_main, cols)
        };
        let units = costs
            .iter()
            .fold(tail_units, |total, cost| total.saturating_add(*cost));
        if !affordable(&plans, units, ctx) {
            return None;
        }
        let mut products: Option<Vec<f64>> = None;
        let mut tail: Option<Vec<f64>> = None;
        for (plan, cost) in plans.iter_mut().zip(costs) {
            // each argument is paid for as it is reached, so one whose error
            // ends the sum leaves the later ones unpaid, as the unstreamed path
            if !ctx.consume_cells(cost) {
                return Some(err(ErrorValue::Num));
            }
            evaluate_literals(plan, ctx);
            read_in_order(plan, leaves, ctx);
            // an argument is read whole before its first error answers, as the
            // unstreamed path reads it, so cycle and date text checks see it all
            let mut error = None;
            let first = products.is_none();
            if first {
                products = Some(vec![0.0; max_main * cols]);
            }
            let values = products.as_mut().unwrap();
            for row in 0..max_main {
                for col in 0..cols {
                    let value = node_value(&plan.root, plan, leaves, row, col, false, ctx);
                    if error.is_some() {
                        continue;
                    }
                    let number = match as_number(value) {
                        Ok(number) => number,
                        Err(found) => {
                            error = Some(found);
                            continue;
                        }
                    };
                    let slot = &mut values[row * cols + col];
                    if first {
                        *slot = number;
                    } else {
                        *slot *= number;
                    }
                }
            }
            if max_main < rows
                && let Some(start) = plan.tail_start.filter(|start| *start < rows)
            {
                let first = tail.is_none();
                if first {
                    tail = Some(vec![0.0; cols]);
                }
                let values = tail.as_mut().unwrap();
                for (col, slot) in values.iter_mut().enumerate() {
                    let value = node_value(&plan.root, plan, leaves, start, col, true, ctx);
                    if error.is_some() {
                        continue;
                    }
                    let number = match as_number(value) {
                        Ok(number) => number,
                        Err(found) => {
                            error = Some(found);
                            continue;
                        }
                    };
                    if first {
                        *slot = number;
                    } else {
                        *slot *= number;
                    }
                }
            }
            if let Some(error) = error {
                return Some(err(error));
            }
        }
        // a row of zeros added again leaves the sum as its first addition did
        let tail = tail.unwrap_or_default();
        let zeros = tail.iter().all(|value| *value == 0.0);
        if !zeros && !ctx.consume_cells(tail_cost(rows, max_main, cols)) {
            return Some(err(ErrorValue::Num));
        }
        let mut total = 0.0;
        for value in products.unwrap_or_default() {
            total += value;
        }
        if max_main < rows {
            let repeats = if zeros { 1 } else { rows - max_main };
            for _ in 0..repeats {
                for value in &tail {
                    total += value;
                }
            }
        }
        Some(super::num(total))
    }

    fn run_array(
        mut plans: Vec<Plan<'_>>,
        leaves: &[Leaf],
        shape: (usize, usize),
        ctx: &EvalContext<'_>,
    ) -> Option<CellValue> {
        let (rows, cols) = shape;
        let mut main_rows = 0usize;
        let mut tail = true;
        for plan in &plans {
            if plan.shape.0 != 1 && plan.shape.0 != rows {
                tail = false;
            } else if plan.shape.0 == rows {
                match plan.tail_start {
                    Some(start) => main_rows = main_rows.max(start),
                    None => tail = false,
                }
            }
        }
        let main_rows = if tail { main_rows } else { rows };
        let mut charged: HashSet<(SheetId, u32, u32, usize, usize)> = HashSet::new();
        let mut units = 0u64;
        for plan in &plans {
            for index in &plan.leaves {
                let leaf = &leaves[*index];
                if charged.insert(leaf.key) {
                    units = units.saturating_add(leaf.charge_cells);
                }
            }
        }
        let positions = (main_rows as u64)
            .saturating_mul(cols as u64)
            .saturating_add(if main_rows < rows { cols as u64 } else { 0 });
        let units = plans.iter().fold(units, |total, plan| {
            total.saturating_add(step_cost(plan, positions))
        });
        // a tail known to add only zeros is added once, so it costs nothing;
        // any other tail is checked now and paid for once its row is known
        let zero = plans.iter().any(|plan| zero_at_tail(&plan.root, leaves));
        let tail_units = if zero {
            0
        } else {
            tail_cost(rows, main_rows, cols)
        };
        if !affordable(&plans, units.saturating_add(tail_units), ctx) {
            return None;
        }
        if !ctx.consume_cells(units) {
            return Some(err(ErrorValue::Num));
        }
        for plan in &mut plans {
            evaluate_literals(plan, ctx);
        }
        for plan in &plans {
            read_in_order(plan, leaves, ctx);
        }
        // every argument is read whole before the first error answers, as the
        // unstreamed path evaluates every argument first
        let mut error = None;
        let mut total = 0.0;
        for row in 0..main_rows {
            for col in 0..cols {
                let mut product = 1.0;
                for plan in &plans {
                    let value = node_value(&plan.root, plan, leaves, row, col, false, ctx);
                    if error.is_none() {
                        match as_number_array(value) {
                            Ok(number) => product *= number,
                            Err(found) => error = Some(found),
                        }
                    }
                }
                total += product;
            }
        }
        if main_rows < rows {
            let mut tail = Vec::with_capacity(cols);
            for col in 0..cols {
                let mut product = 1.0;
                for plan in &plans {
                    let value = node_value(&plan.root, plan, leaves, main_rows, col, true, ctx);
                    if error.is_none() {
                        match as_number_array(value) {
                            Ok(number) => product *= number,
                            Err(found) => error = Some(found),
                        }
                    }
                }
                tail.push(product);
            }
            if let Some(error) = error {
                return Some(err(error));
            }
            // a row of zeros added again leaves the sum as its first addition did
            let zeros = tail.iter().all(|value| *value == 0.0);
            if !zeros && !ctx.consume_cells(tail_cost(rows, main_rows, cols)) {
                return Some(err(ErrorValue::Num));
            }
            let repeats = if zeros { 1 } else { rows - main_rows };
            for _ in 0..repeats {
                for value in &tail {
                    total += value;
                }
            }
        }
        if let Some(error) = error {
            return Some(err(error));
        }
        Some(super::num(total))
    }

    fn as_number(value: CellValue) -> Result<f64, ErrorValue> {
        match value {
            CellValue::Number { value } => Ok(value),
            CellValue::Error { value } => Err(value),
            _ => Ok(0.0),
        }
    }

    fn as_number_array(value: CellValue) -> Result<f64, ErrorValue> {
        match value {
            CellValue::Number { value } => Ok(value),
            CellValue::Bool { value } => Ok(f64::from(value)),
            CellValue::Error { value } => Err(value),
            _ => Ok(0.0),
        }
    }

    fn node_value(
        node: &Node,
        plan: &Plan<'_>,
        leaves: &[Leaf],
        row: usize,
        col: usize,
        tail: bool,
        ctx: &EvalContext<'_>,
    ) -> CellValue {
        match node {
            Node::Leaf(index) => leaf_value(&leaves[*index], row, col, tail, ctx),
            Node::Scalar(index) => plan.scalar_values[*index].clone(),
            Node::Unary(op, child) => {
                let value = node_value(child, plan, leaves, row, col, tail, ctx);
                coerced(&value, ctx, |value| apply_unary(*op, value))
            }
            Node::Percent(child) => {
                let value = node_value(child, plan, leaves, row, col, tail, ctx);
                coerced(&value, ctx, apply_percent)
            }
            Node::Binary(op, lhs, rhs) => {
                let lhs = node_value(lhs, plan, leaves, row, col, tail, ctx);
                let rhs = node_value(rhs, plan, leaves, row, col, tail, ctx);
                apply_binary(*op, &lhs, &rhs, ctx)
            }
        }
    }

    fn leaf_value(
        leaf: &Leaf,
        row: usize,
        col: usize,
        _tail: bool,
        ctx: &EvalContext<'_>,
    ) -> CellValue {
        let row = if leaf.shape.0 == 1 { 0 } else { row };
        if row >= leaf.shape.0 {
            return err(ErrorValue::NA);
        }
        let col = if leaf.shape.1 == 1 { 0 } else { col };
        if col >= leaf.shape.1 {
            return err(ErrorValue::NA);
        }
        if row >= leaf.read_rows {
            return CellValue::Empty;
        }
        leaf.area.get_unmetered_ref(ctx, row, col).into_owned()
    }

    struct Builder<'a, 'p, 'e> {
        ctx: &'a EvalContext<'p>,
        leaves: &'a mut Vec<Leaf>,
        scalars: Vec<&'e Expr>,
        used_rows: HashMap<SheetId, usize>,
        used_cols: HashMap<SheetId, usize>,
    }

    impl<'a, 'p, 'e> Builder<'a, 'p, 'e> {
        fn used_rows(&mut self, sheet: SheetId) -> usize {
            *self
                .used_rows
                .entry(sheet)
                .or_insert_with(|| self.ctx.provider.used_rows(sheet) as usize)
        }

        fn used_cols(&mut self, sheet: SheetId) -> usize {
            *self
                .used_cols
                .entry(sheet)
                .or_insert_with(|| self.ctx.provider.used_cols(sheet) as usize)
        }

        fn leaf(&mut self, expr: &Expr) -> Option<usize> {
            let area = as_area(expr, self.ctx)?;
            let declared = (area.rows, area.cols);
            let key = (
                area.sheet,
                area.start.row,
                area.start.col,
                declared.0,
                declared.1,
            );
            if let Some(index) = self.leaves.iter().position(|leaf| leaf.key == key) {
                return Some(index);
            }
            let mut shape = declared;
            if area.rows >= xlsx_model::MAX_ROWS as usize {
                let used = self.used_rows(area.sheet);
                shape.0 = used.saturating_sub(area.start.row as usize).max(1);
            }
            if area.cols >= xlsx_model::MAX_COLS as usize {
                let used = self.used_cols(area.sheet);
                shape.1 = used.saturating_sub(area.start.col as usize).max(1);
            }
            let used = self.used_rows(area.sheet);
            let read_rows = if used == 0 {
                shape.0
            } else {
                used.saturating_sub(area.start.row as usize).min(shape.0)
            };
            let charge_cells = (read_rows as u64).saturating_mul(shape.1 as u64);
            let area = Area {
                rows: shape.0,
                cols: shape.1,
                ..area
            };
            self.leaves.push(Leaf {
                area,
                shape,
                read_rows: read_rows.min(shape.0),
                charge_cells,
                key,
            });
            Some(self.leaves.len() - 1)
        }

        fn scalar(&mut self, expr: &'e Expr) -> (Node, (usize, usize), u64) {
            let index = self.scalars.len();
            self.scalars.push(expr);
            (Node::Scalar(index), (1, 1), 0)
        }

        fn build(&mut self, expr: &'e Expr) -> Option<(Node, (usize, usize), u64)> {
            match expr {
                Expr::Ref { .. }
                | Expr::Range { .. }
                | Expr::ColumnRange { .. }
                | Expr::RowRange { .. } => {
                    let index = self.leaf(expr)?;
                    Some((Node::Leaf(index), self.leaves[index].shape, 0))
                }
                Expr::Number(_)
                | Expr::Text(_)
                | Expr::Bool(_)
                | Expr::Error(_)
                | Expr::Literal(_) => Some(self.scalar(expr)),
                Expr::Unary { op, expr } => {
                    let (child, shape, ops) = self.build(expr)?;
                    Some((Node::Unary(*op, Box::new(child)), shape, ops + 1))
                }
                Expr::Percent(inner) => {
                    let (child, shape, ops) = self.build(inner)?;
                    Some((Node::Percent(Box::new(child)), shape, ops + 1))
                }
                Expr::Binary { op, lhs, rhs } if *op != BinaryOp::Concat => {
                    let (lhs, lshape, lops) = self.build(lhs)?;
                    let (rhs, rshape, rops) = self.build(rhs)?;
                    let shape = (lshape.0.max(rshape.0), lshape.1.max(rshape.1));
                    Some((
                        Node::Binary(*op, Box::new(lhs), Box::new(rhs)),
                        shape,
                        lops + rops + 1,
                    ))
                }
                _ => None,
            }
        }
    }

    #[cfg(test)]
    mod tests {
        use std::rc::Rc;

        use xlsx_model::{
            Cell, CellProvider, CellRange, CellRef, CellValue, ErrorValue, Sheet, Workbook,
        };

        use super::{Area, EvalContext, Leaf, Node, SheetId, tail_start};
        use crate::engine::RecalcResult;
        use crate::eval::EvaluationBudget;
        use crate::parser::{BinaryOp, Expr, parse_formula};

        type Cells<'a> = &'a [(&'a str, CellValue, Option<&'a str>)];

        /// `cells` recalculated with streaming on or off; `B1` is an array
        /// formula when `array` is set.
        fn recalc(cells: Cells<'_>, streamed: bool, array: bool) -> (Vec<CellValue>, RecalcResult) {
            let mut wb = Workbook::default();
            let mut sheet = Sheet::new("Sheet1");
            for (a1, value, formula) in cells {
                sheet.set_cell(
                    CellRef::parse_a1(a1).unwrap(),
                    Cell {
                        value: value.clone(),
                        formula: formula.map(str::to_string),
                        ..Cell::default()
                    },
                );
            }
            if array {
                let at = CellRef::parse_a1("B1").unwrap();
                sheet.set_array_formula(at, CellRange::new(at, at));
            }
            wb.sheets.push(sheet);
            super::UNSTREAMED.set(!streamed);
            let (_graph, result) = crate::rebuild_and_recalc_all(&mut wb, None);
            super::UNSTREAMED.set(false);
            let values = cells
                .iter()
                .map(|(a1, ..)| wb.value(SheetId(0), CellRef::parse_a1(a1).unwrap()))
                .collect();
            (values, result)
        }

        /// an argument with an error is still read whole, as the unstreamed
        /// path reads it: a cycle through a later row is found, and date text
        /// on a later row keeps the cached value.
        #[test]
        fn an_argument_error_still_reads_the_rest_of_the_argument() {
            let text = |value: &str| CellValue::Text {
                value: value.into(),
            };
            let div0 = CellValue::Error {
                value: ErrorValue::Div0,
            };
            let cases: [Cells<'_>; 4] = [
                &[
                    ("A1", div0.clone(), None),
                    ("A2", CellValue::Empty, Some("B1+1")),
                    ("B1", CellValue::Empty, Some("SUMPRODUCT(A1:A2)")),
                ],
                &[
                    ("A1", text("x"), None),
                    ("A2", CellValue::Empty, Some("B1+1")),
                    ("B1", CellValue::Empty, Some("SUMPRODUCT(A1:A2*1)")),
                ],
                &[
                    ("A1", text("abc"), None),
                    ("A2", text("2024-01-15"), None),
                    (
                        "B1",
                        CellValue::Number { value: 42.0 },
                        Some("SUMPRODUCT(--A1:A2)"),
                    ),
                ],
                &[
                    ("A1", div0, None),
                    ("A2", CellValue::Number { value: 1.0 }, None),
                    ("C1", CellValue::Number { value: 2.0 }, None),
                    ("C2", CellValue::Number { value: 3.0 }, None),
                    ("D1", CellValue::Empty, Some("B1+1")),
                    (
                        "B1",
                        CellValue::Empty,
                        Some("SUMPRODUCT(A1:A2,C1:C2*ABS(D1))"),
                    ),
                ],
            ];
            for cells in cases {
                same_as_unstreamed(cells);
            }
        }

        /// streamed ranges are read in the unstreamed order, so a settle pass
        /// closes the same cycle and settles the same cells.
        #[test]
        fn ranges_are_read_in_the_unstreamed_order_for_cycles() {
            for formula in ["SUMPRODUCT(A1:A2*C1:C2)", "SUMPRODUCT(A1:A2,C1:C2)"] {
                same_as_unstreamed(&[
                    ("A1", CellValue::Number { value: 1.0 }, None),
                    ("A2", CellValue::Empty, Some("B1+1")),
                    ("C1", CellValue::Empty, Some("A2+10")),
                    ("C2", CellValue::Number { value: 1.0 }, None),
                    ("B1", CellValue::Empty, Some(formula)),
                ]);
            }
        }

        /// values and cycle cells match the unstreamed path, with `B1` as a
        /// plain and as an array formula.
        fn same_as_unstreamed(cells: Cells<'_>) {
            for array in [false, true] {
                let (values, result) = recalc(cells, true, array);
                let (expected, unstreamed) = recalc(cells, false, array);
                let formula = cells.iter().find(|(a1, ..)| *a1 == "B1").unwrap().2;
                let case = format!("{formula:?} array={array}");
                assert_eq!(values, expected, "{case}");
                assert_eq!(result.cycle_cells, unstreamed.cycle_cells, "{case}");
            }
        }

        fn budgeted(wb: &Workbook, args: &[Expr], budget: u64, streamed: bool) -> (CellValue, u64) {
            let budget = Rc::new(EvaluationBudget::new(budget));
            let ctx = EvalContext::with_budget(wb, SheetId(0), Rc::clone(&budget));
            let value = if streamed {
                super::super::sumproduct(args, &ctx)
            } else {
                super::super::sumproduct_unstreamed(args, &ctx)
            };
            (value, budget.remaining())
        }

        /// the tail row repeated past the used rows is paid for per row, on the
        /// scalar and the array route, unless it adds only zeros: then it is
        /// added once and costs nothing, as the sum cannot change again.
        #[test]
        fn repeated_tail_rows_are_charged() {
            let mut wb = Workbook::default();
            let mut sheet = Sheet::new("Data");
            sheet.set_cell(
                CellRef::new(0, 0),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    ..Cell::default()
                },
            );
            wb.sheets.push(sheet);
            let repeats = 1_048_575u64;
            for (formula, total) in [
                ("SUMPRODUCT((A2:A1048576=\"\")*1)", repeats as f64),
                ("SUMPRODUCT((A2:A1048576=5)*1)", 0.0),
            ] {
                let Expr::FuncCall { args, .. } = parse_formula(formula).unwrap() else {
                    panic!("{formula} is not a call");
                };
                for array in [false, true] {
                    let budget = Rc::new(EvaluationBudget::new(10_000_000));
                    let ctx = EvalContext::with_budget(&wb, SheetId(0), Rc::clone(&budget));
                    let value = if array {
                        super::super::sumproduct_array(&args, &ctx)
                    } else {
                        super::sumproduct(&args, &ctx)
                    };
                    let case = format!("{formula} array={array}");
                    assert_eq!(value, Some(CellValue::Number { value: total }), "{case}");
                    let spent = 10_000_000 - budget.remaining();
                    if total == 0.0 {
                        assert!(spent < 16, "{case}: {spent} spent");
                    } else {
                        assert!(spent >= repeats / 16, "{case}: {spent} spent");
                    }
                }
            }
        }

        /// an argument whose error ends the sum leaves the later arguments
        /// unpaid, as the unstreamed path does.
        #[test]
        fn an_argument_error_spends_like_the_unstreamed_path() {
            let mut wb = Workbook::default();
            let mut sheet = Sheet::new("Data");
            for row in 0..1_000u32 {
                let first = if row == 0 {
                    CellValue::Error {
                        value: ErrorValue::Div0,
                    }
                } else {
                    CellValue::Number { value: 1.0 }
                };
                for (col, value) in [first, CellValue::Number { value: 2.0 }]
                    .into_iter()
                    .enumerate()
                {
                    sheet.set_cell(
                        CellRef::new(row, col as u32),
                        Cell {
                            value,
                            ..Cell::default()
                        },
                    );
                }
            }
            wb.sheets.push(sheet);
            let formula = "SUMPRODUCT(A1:A1000,B1:B1000)";
            let Expr::FuncCall { args, .. } = parse_formula(formula).unwrap() else {
                panic!("{formula} is not a call");
            };
            assert_eq!(
                budgeted(&wb, &args, 1_000_000, true),
                budgeted(&wb, &args, 1_000_000, false)
            );
        }

        /// ranges declared past the data, as summary sheets write them: a
        /// streamed SUMPRODUCT the budget cannot cover spends and answers as
        /// the unstreamed one does, so it never drains budget a later cell had.
        #[test]
        fn refused_streamed_sumproduct_spends_like_the_unstreamed_path() {
            let mut wb = Workbook::default();
            let mut sheet = Sheet::new("Data");
            for row in 0..1_000u32 {
                for (col, value) in [1.0, 2.0, 3.0, -2.0].into_iter().enumerate() {
                    sheet.set_cell(
                        CellRef::new(row, col as u32),
                        Cell {
                            value: CellValue::Number { value },
                            ..Cell::default()
                        },
                    );
                }
            }
            wb.sheets.push(sheet);
            for (formula, total) in [
                ("SUMPRODUCT((A1:A3000=1)*(B1:B3000=2)*C1:C3000)", 3_000.0),
                (
                    "SUMPRODUCT((A1:A3000=1)*(B1:B3000=ABS(D1))*C1:C3000)",
                    3_000.0,
                ),
                ("SUMPRODUCT((A2:A3000=\"\")*ISERROR(D1))", 0.0),
            ] {
                let Expr::FuncCall { args, .. } = parse_formula(formula).unwrap() else {
                    panic!("{formula} is not a call");
                };
                let total = CellValue::Number { value: total };
                let (value, left) = budgeted(&wb, &args, 1_000_000, true);
                assert_eq!(value, total);
                // a streamed formula never spends more than the unstreamed one
                let unstreamed = budgeted(&wb, &args, 1_000_000, false).1;
                assert!(
                    left >= unstreamed,
                    "{formula}: {left} left, unstreamed {unstreamed}"
                );
                for budget in (0..=12_000).step_by(250).chain(3_000..3_400) {
                    let streamed = budgeted(&wb, &args, budget, true);
                    if streamed.0 != total {
                        assert_eq!(
                            streamed,
                            budgeted(&wb, &args, budget, false),
                            "{formula} budget {budget}"
                        );
                    }
                }
            }
        }

        fn leaf(shape: (usize, usize), read_rows: usize) -> Leaf {
            Leaf {
                area: Area {
                    sheet: SheetId(0),
                    start: CellRef::new(0, 0),
                    rows: shape.0,
                    cols: shape.1,
                },
                shape,
                read_rows,
                charge_cells: 0,
                key: (SheetId(0), 0, 0, shape.0, shape.1),
            }
        }

        #[test]
        fn tail_starts_past_the_last_read_row() {
            let leaves = vec![leaf((100, 1), 4), leaf((1, 3), 1), leaf((50, 1), 50)];
            let root = Node::Binary(
                BinaryOp::Mul,
                Box::new(Node::Leaf(0)),
                Box::new(Node::Leaf(1)),
            );
            assert_eq!(tail_start(&root, &leaves, 100), Some(4));
            let mismatched = Node::Binary(
                BinaryOp::Mul,
                Box::new(Node::Leaf(0)),
                Box::new(Node::Leaf(2)),
            );
            assert_eq!(tail_start(&mismatched, &leaves, 100), None);
            assert_eq!(tail_start(&root, &leaves, 1), None);
        }
    }
}

/// MMULT(array1, array2): the matrix product; `cols(array1)` must equal
/// `rows(array2)` and any non-numeric operand cell is `#VALUE!`. a cell holds
/// one value, so this is the top-left element excel caches in the anchor.
pub(crate) fn mmult(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let left = match matrix(&args[0], ctx) {
        Ok(m) => m,
        Err(e) => return err(e),
    };
    let right = match matrix(&args[1], ctx) {
        Ok(m) => m,
        Err(e) => return err(e),
    };
    if left.cols != right.rows {
        return err(ErrorValue::Value);
    }
    let mut total = 0.0;
    for k in 0..left.cols {
        total += left.values[k] * right.values[k * right.cols];
    }
    finite(total)
}

/// a rectangle of numbers in row-major order.
struct Matrix {
    rows: usize,
    cols: usize,
    values: Vec<f64>,
}

/// read an argument as a numeric rectangle; a non-reference is 1x1 and any
/// non-numeric cell is `#VALUE!`.
fn matrix(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Matrix, ErrorValue> {
    let Some(area) = as_area(arg, ctx) else {
        return match evaluate(arg, ctx) {
            CellValue::Number { value } => Ok(Matrix {
                rows: 1,
                cols: 1,
                values: vec![value],
            }),
            CellValue::Error { value } => Err(value),
            _ => Err(ErrorValue::Value),
        };
    };
    let cells = area.values_ref(ctx)?;
    // the blanks past the used extent are not numbers
    if area.unread(cells.len() as u64) > 0 {
        return Err(ErrorValue::Value);
    }
    let mut values = Vec::with_capacity(cells.len());
    for cell in cells {
        match *cell {
            CellValue::Number { value } => values.push(value),
            CellValue::Error { value } => return Err(value),
            _ => return Err(ErrorValue::Value),
        }
    }
    Ok(Matrix {
        rows: area.rows,
        cols: area.cols,
        values,
    })
}

pub(crate) fn abs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::abs)
}

pub(crate) fn sign(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x > 0.0 {
            1.0
        } else if x < 0.0 {
            -1.0
        } else {
            0.0
        }
    })
}

pub(crate) fn sqrt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) if x < 0.0 => err(ErrorValue::Num),
        Ok(x) => num(x.sqrt()),
        Err(e) => err(e),
    }
}

pub(crate) fn exp(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => finite(x.exp()),
        Err(e) => err(e),
    }
}

pub(crate) fn ln(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    positive_log(args, ctx, f64::ln)
}

pub(crate) fn log10(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    positive_log(args, ctx, f64::log10)
}

/// LOG(number, [base]); base defaults to 10.
pub(crate) fn log(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let base = if args.len() == 2 {
        match nth_number(args, ctx, 1) {
            Ok(b) => b,
            Err(e) => return err(e),
        }
    } else {
        10.0
    };
    match nth_number(args, ctx, 0) {
        Ok(x) if x <= 0.0 || base <= 0.0 || base == 1.0 => err(ErrorValue::Num),
        Ok(x) => finite(x.log(base)),
        Err(e) => err(e),
    }
}

/// LN / LOG10: positive inputs only. uses the dedicated libm routine (not
/// `x.log(base)`) so exact powers round-trip precisely.
fn positive_log(args: &[Expr], ctx: &EvalContext<'_>, f: fn(f64) -> f64) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) if x <= 0.0 => err(ErrorValue::Num),
        Ok(x) => finite(f(x)),
        Err(e) => err(e),
    }
}

pub(crate) fn pi(args: &[Expr], _ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() {
        num(std::f64::consts::PI)
    } else {
        err(ErrorValue::Value)
    }
}

pub(crate) fn power(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(x), Ok(y)) => finite(x.powf(y)),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// MOD(n, d) = n - d*INT(n/d); sign follows the divisor. d = 0 -> #DIV/0!.
pub(crate) fn mod_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => err(ErrorValue::Div0),
        (Ok(n), Ok(d)) => num(n - d * (n / d).floor()),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

pub(crate) fn int(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::floor)
}

/// every cell of the arguments as a number, refusing anything that is not one.
fn strict_numbers(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = Vec::new();
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                for value in area.values_ref(ctx)? {
                    match value.as_ref() {
                        CellValue::Number { value } => nums.push(*value),
                        CellValue::Error { value } => return Err(*value),
                        CellValue::Empty => {}
                        _ => return Err(ErrorValue::Value),
                    }
                }
            }
            None => match evaluate(arg, ctx) {
                CellValue::Number { value } => nums.push(value),
                CellValue::Error { value } => return Err(value),
                _ => return Err(ErrorValue::Value),
            },
        }
    }
    Ok(nums)
}

/// QUOTIENT(numerator, denominator): the integer part of the division,
/// truncated toward zero.
pub(crate) fn quotient(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => err(ErrorValue::Div0),
        (Ok(n), Ok(d)) => finite((n / d).trunc()),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// SERIESSUM(x, n, m, coefficients): sum of `a_i * x^(n + (i-1) * m)`.
pub(crate) fn seriessum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let (x, n, m) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(x), Ok(n), Ok(m)) => (x, n, m),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    // a non-numeric coefficient is #VALUE!, not a term to skip: dropping one
    // would shift every later power
    let coefficients = match strict_numbers(&args[3..], ctx) {
        Ok(c) => c,
        Err(e) => return err(e),
    };
    let mut total = 0.0;
    for (i, a) in coefficients.iter().enumerate() {
        total += a * x.powf(n + (i as f64) * m);
    }
    finite(total)
}

/// TRUNC(number, [digits]); digits default 0. truncates toward zero.
pub(crate) fn trunc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    directional(args, ctx, |scaled| scaled.trunc())
}

/// ROUNDUP(number, digits): away from zero.
pub(crate) fn roundup(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    directional(args, ctx, |scaled| {
        if scaled < 0.0 {
            scaled.floor()
        } else {
            scaled.ceil()
        }
    })
}

/// ROUNDDOWN(number, digits): toward zero (same as TRUNC with digits).
pub(crate) fn rounddown(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    directional(args, ctx, |scaled| scaled.trunc())
}

/// ROUND(number, digits): half away from zero.
pub(crate) fn round(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let digits = match nth_int(args, ctx, 1) {
        Ok(d) => d as i32,
        Err(e) => return err(e),
    };
    let factor = 10f64.powi(digits);
    finite((x * factor).round() / factor)
}

/// MROUND(number, multiple): nearest multiple, half away from zero. opposite
/// signs -> #NUM!; multiple 0 -> 0.
pub(crate) fn mround(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => num(0.0),
        (Ok(x), Ok(m)) if (x < 0.0) != (m < 0.0) && x != 0.0 => err(ErrorValue::Num),
        (Ok(x), Ok(m)) => num((x / m).round() * m),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// CEILING(number, significance): away from zero to a multiple of significance.
/// positive number with negative significance is #NUM!; significance 0 -> 0.
pub(crate) fn ceiling(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    round_to_multiple(args, ctx, f64::ceil)
}

/// FLOOR(number, significance): toward zero to a multiple of significance.
pub(crate) fn floor(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    round_to_multiple(args, ctx, f64::floor)
}

fn round_to_multiple(args: &[Expr], ctx: &EvalContext<'_>, f: fn(f64) -> f64) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => num(0.0),
        (Ok(x), Ok(s)) if x > 0.0 && s < 0.0 => err(ErrorValue::Num),
        (Ok(x), Ok(s)) => num(f(x / s) * s),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

fn one_number(args: &[Expr], ctx: &EvalContext<'_>) -> Result<f64, ErrorValue> {
    if args.len() != 1 {
        return Err(ErrorValue::Value);
    }
    nth_number(args, ctx, 0)
}

fn unary(args: &[Expr], ctx: &EvalContext<'_>, f: fn(f64) -> f64) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => num(f(x)),
        Err(e) => err(e),
    }
}

/// scale by 10^digits, apply a rounding rule, unscale.
fn directional(args: &[Expr], ctx: &EvalContext<'_>, rule: fn(f64) -> f64) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let digits = if args.len() == 2 {
        match nth_int(args, ctx, 1) {
            Ok(d) => d as i32,
            Err(e) => return err(e),
        }
    } else {
        0
    };
    let factor = 10f64.powi(digits);
    finite(rule(x * factor) / factor)
}

pub(crate) fn tanh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::tanh)
}

pub(crate) fn sin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::sin)
}

pub(crate) fn cos(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::cos)
}

pub(crate) fn tan(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::tan)
}

pub(crate) fn sinh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::sinh)
}

pub(crate) fn cosh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::cosh)
}

/// ASIN/ACOS are `#NUM!` outside [-1, 1]; `finite` turns the NaN into it.
pub(crate) fn asin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => finite(x.asin()),
        Err(e) => err(e),
    }
}

pub(crate) fn acos(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => finite(x.acos()),
        Err(e) => err(e),
    }
}

pub(crate) fn atan(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::atan)
}

/// ATAN2(x, y) — excel takes the x coordinate first, the reverse of `f64::atan2`.
/// both zero is `#DIV/0!`.
pub(crate) fn atan2(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(0.0), Ok(0.0)) => err(ErrorValue::Div0),
        (Ok(x), Ok(y)) => finite(y.atan2(x)),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

pub(crate) fn degrees(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::to_degrees)
}

pub(crate) fn radians(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::to_radians)
}

/// RANDBETWEEN(bottom, top): a volatile integer draw. measured against excel:
/// a raw `bottom > top` is `#NUM!`, the draw spans `ceil(bottom)..=floor(top)`,
/// and a span that rounds away to nothing collapses to `ceil(bottom)`.
pub(crate) fn randbetween(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let (bottom, top) = match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(bottom), Ok(top)) => (bottom, top),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    if bottom > top {
        return err(ErrorValue::Num);
    }
    let (lo, hi) = (bottom.ceil(), top.floor());
    if hi < lo {
        return num(lo);
    }
    let span = hi - lo + 1.0;
    finite(lo + (ctx.next_random_unit() * span).floor().min(hi - lo))
}

/// RAND(): volatile, uniform over [0, 1), drawn from the same seeded stream as
/// RANDBETWEEN.
pub(crate) fn rand(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if !args.is_empty() {
        return err(ErrorValue::Value);
    }
    num(ctx.next_random_unit())
}

/// 2^53: GCD/LCM refuse an argument or a result at or past it with `#NUM!`.
const WHOLE_LIMIT: f64 = 9_007_199_254_740_992.0;

/// the arguments of GCD/LCM as truncated whole numbers. a blank reads as 0,
/// whether a cell, a cell in a range or an omitted argument; references
/// refuse text or logicals; a negative value or one at or past 2^53 is
/// `#NUM!`.
fn whole_arguments(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Vec<u64>, ErrorValue> {
    if args.is_empty() {
        return Err(ErrorValue::Value);
    }
    let mut values = Vec::new();
    for arg in args {
        let (cells, from_sheet) = super::argument_cells(arg, ctx)?;
        for cell in cells {
            let value = match cell {
                CellValue::Number { value } => value,
                CellValue::Empty => 0.0,
                CellValue::Error { value } => return Err(value),
                CellValue::Bool { .. } => return Err(ErrorValue::Value),
                CellValue::Text { .. } if from_sheet => return Err(ErrorValue::Value),
                text @ CellValue::Text { .. } => super::coerce_number(&text, ctx)?,
            };
            if !(0.0..WHOLE_LIMIT).contains(&value) {
                return Err(ErrorValue::Num);
            }
            values.push(value.trunc() as u64);
        }
    }
    Ok(values)
}

fn euclid(mut a: u128, mut b: u128) -> u128 {
    while b != 0 {
        (a, b) = (b, a % b);
    }
    a
}

/// GCD(number, ...): greatest common divisor of the truncated arguments.
pub(crate) fn gcd(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match whole_arguments(args, ctx) {
        Ok(values) => num(values
            .into_iter()
            .fold(0u128, |a, b| euclid(a, u128::from(b))) as f64),
        Err(e) => err(e),
    }
}

/// LCM(number, ...): least common multiple of the truncated arguments; any
/// zero makes it zero, and a multiple at or past 2^53 is `#NUM!`.
pub(crate) fn lcm(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let values = match whole_arguments(args, ctx) {
        Ok(values) => values,
        Err(e) => return err(e),
    };
    if values.is_empty() || values.contains(&0) {
        return num(0.0);
    }
    let mut acc: u128 = 1;
    for value in values {
        let value = u128::from(value);
        match (acc / euclid(acc, value)).checked_mul(value) {
            Some(next) if (next as f64) < WHOLE_LIMIT => acc = next,
            _ => return err(ErrorValue::Num),
        }
    }
    num(acc as f64)
}

/// EVEN(number): away from zero to the nearest even integer.
pub(crate) fn even(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => {
            let magnitude = (x.abs() / 2.0).ceil() * 2.0;
            finite(if x < 0.0 { -magnitude } else { magnitude })
        }
        Err(e) => err(e),
    }
}

/// ODD(number): away from zero to the nearest odd integer; zero gives 1.
pub(crate) fn odd(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => {
            let magnitude = ((x.abs() + 1.0) / 2.0).ceil() * 2.0 - 1.0;
            finite(if x < 0.0 { -magnitude } else { magnitude })
        }
        Err(e) => err(e),
    }
}

/// FACT(number): factorial of the truncated number; negative is `#NUM!`, and
/// so is anything past 170!, which overflows a double.
pub(crate) fn fact(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    // the sign is checked before truncating, so -0.5 is refused, not 0!
    let x = match one_number(args, ctx) {
        Ok(x) if x < 0.0 => return err(ErrorValue::Num),
        Ok(x) => x.trunc(),
        Err(e) => return err(e),
    };
    if x > 170.0 {
        return err(ErrorValue::Num);
    }
    num((2..=x as u32).map(f64::from).product())
}

/// COMBIN(number, chosen): ways to choose `chosen` of `number`, both
/// truncated. a negative argument or `number < chosen` is `#NUM!`.
pub(crate) fn combin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    // signs are checked before truncating, so a negative fraction is refused
    let (n, k) = match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(n), Ok(k)) if n < 0.0 || k < 0.0 => return err(ErrorValue::Num),
        (Ok(n), Ok(k)) => (n.trunc(), k.trunc()),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    if n < k {
        return err(ErrorValue::Num);
    }
    // once k <= n / 2 every factor is at least 2, so both loops end within
    // about a thousand steps however large k is
    let k = k.min(n - k);
    let mut result = 1.0;
    let mut i = 1.0;
    // whole numbers while the running binomial fits, so every result below
    // 2^53 is exact
    if n < 2f64.powi(64) {
        let base = (n - k) as u128;
        let mut exact: u128 = 1;
        while i <= k {
            let Some(product) = exact.checked_mul(base + i as u128) else {
                break;
            };
            exact = product / i as u128;
            i += 1.0;
        }
        result = exact as f64;
    }
    // past that, divide first where multiplying first would overflow, so
    // only a result beyond a double is #NUM!
    while i <= k {
        let product = result * (n - k + i);
        result = if product.is_finite() {
            product / i
        } else {
            result / i * (n - k + i)
        };
        if !result.is_finite() {
            return err(ErrorValue::Num);
        }
        i += 1.0;
    }
    num(result.round())
}

/// SUMSQ(number, ...): the sum of the squares of every numeric argument.
pub(crate) fn sumsq(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match super::collect_numbers_deep(args, ctx) {
        Ok(nums) => finite(nums.iter().map(|x| x * x).sum()),
        Err(e) => err(e),
    }
}

/// CONVERT(number, from, to): excel's unit conversion, over the unit families
/// it defines. A metric unit may carry an SI prefix; the two units must share
/// a family.
pub(crate) fn convert(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let value = match nth_number(args, ctx, 0) {
        Ok(value) => value,
        Err(e) => return err(e),
    };
    let (from, to) = match (
        to_text(&evaluate(&args[1], ctx)),
        to_text(&evaluate(&args[2], ctx)),
    ) {
        (Ok(a), Ok(b)) => (a, b),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    let (Some((from_family, from_scale)), Some((to_family, to_scale))) = (unit(&from), unit(&to))
    else {
        return err(ErrorValue::NA);
    };
    if from_family != to_family {
        return err(ErrorValue::NA);
    }
    if from_family == TEMPERATURE {
        return match to_celsius(&from, value).and_then(|c| from_celsius(&to, c)) {
            Some(out) => finite(out),
            None => err(ErrorValue::NA),
        };
    }
    finite(value * from_scale / to_scale)
}

const TEMPERATURE: u8 = 4;

/// a unit's family and its size in the family's base unit, after any SI
/// prefix. The base units are the metre, the gram, the second and the litre.
fn unit(name: &str) -> Option<(u8, f64)> {
    if let Some(found) = base_unit(name) {
        return Some(found);
    }
    let mut chars = name.chars();
    let head = chars.next()?;
    let rest = chars.as_str();
    let (family, scale) = base_unit(rest)?;
    if !matches!(family, 0..=3) {
        return None;
    }
    Some((family, scale * si_prefix(head)?))
}

fn base_unit(name: &str) -> Option<(u8, f64)> {
    Some(match name {
        // 0: length, in metres
        "m" => (0, 1.0),
        "mi" => (0, 1609.344),
        "Nmi" => (0, 1852.0),
        "in" => (0, 0.0254),
        "ft" => (0, 0.3048),
        "yd" => (0, 0.9144),
        "ang" => (0, 1e-10),
        "Pica" | "pica" => (0, 0.0254 / 72.0),
        // 1: mass, in grams
        "g" => (1, 1.0),
        "sg" => (1, 14593.903),
        "lbm" => (1, 453.59237),
        "u" => (1, 1.660_538_782e-24),
        "ozm" => (1, 28.349523125),
        // 2: time, in seconds
        "sec" | "s" => (2, 1.0),
        "mn" | "min" => (2, 60.0),
        "hr" => (2, 3600.0),
        "day" | "d" => (2, 86400.0),
        "yr" => (2, 31_557_600.0),
        // 3: liquid measure, in litres
        "l" | "L" | "lt" => (3, 1.0),
        "tsp" => (3, 0.004_928_921_593_75),
        "tbs" => (3, 0.014_786_764_781_25),
        "oz" => (3, 0.0295735295625),
        "cup" => (3, 0.2365882365),
        "pt" | "us_pt" => (3, 0.473176473),
        "qt" => (3, 0.946352946),
        "gal" => (3, 3.785411784),
        // 4: temperature, converted rather than scaled
        "C" | "cel" => (TEMPERATURE, 1.0),
        "F" | "fah" => (TEMPERATURE, 1.0),
        "K" | "kel" => (TEMPERATURE, 1.0),
        _ => return None,
    })
}

fn si_prefix(symbol: char) -> Option<f64> {
    Some(match symbol {
        'Y' => 1e24,
        'Z' => 1e21,
        'E' => 1e18,
        'P' => 1e15,
        'T' => 1e12,
        'G' => 1e9,
        'M' => 1e6,
        'k' => 1e3,
        'h' => 1e2,
        'e' => 1e1,
        'd' => 1e-1,
        'c' => 1e-2,
        'm' => 1e-3,
        'u' => 1e-6,
        'n' => 1e-9,
        'p' => 1e-12,
        'f' => 1e-15,
        'a' => 1e-18,
        'z' => 1e-21,
        'y' => 1e-24,
        _ => return None,
    })
}

fn to_celsius(name: &str, value: f64) -> Option<f64> {
    Some(match name {
        "C" | "cel" => value,
        "F" | "fah" => (value - 32.0) * 5.0 / 9.0,
        "K" | "kel" => value - 273.15,
        _ => return None,
    })
}

fn from_celsius(name: &str, value: f64) -> Option<f64> {
    Some(match name {
        "C" | "cel" => value,
        "F" | "fah" => value * 9.0 / 5.0 + 32.0,
        "K" | "kel" => value + 273.15,
        _ => return None,
    })
}
