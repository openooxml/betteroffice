//! whether the engine can evaluate a formula wherever it could run, decided
//! from the formula alone so that no input changes the answer.

use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use xlsx_model::{
    ArraySize, Cell, CellProvider, CellRange, CellRef, CellValue, ErrorValue, MAX_SPILL_CELLS,
    SheetId, Workbook,
};

use crate::array::evaluate_spill;
use crate::eval::{EvalContext, EvaluationBudget, MAX_RECALCULATION_CELL_VISITS};
use crate::graph::DepGraph;
use crate::parser::{Expr, parse_formula};

/// analysis steps one formula's classification takes, each expression node,
/// binding and name reference it visits, before it answers no for that
/// formula alone.
const MAX_WORK: usize = 1 << 20;

/// whether the engine implements everything `formula`, on `sheet`, could
/// run: it parses, and every function it calls, in every branch and in every
/// defined name it reaches, is one the engine evaluates. an `IF` or `IFERROR`
/// around a function the engine lacks answers no whichever way it would go.
pub fn evaluable(wb: &Workbook, sheet: SheetId, formula: &str) -> bool {
    Capability::new(wb).evaluable(sheet, formula)
}

/// Decides, once, which of a workbook's arrays the engine can evaluate. Each
/// such array takes the cells its file cached beside its anchor for its
/// result; every other array is opaque and keeps them as stored cells. An
/// opaque array whose anchor holds no cached value gets what it shows stored
/// over its recorded rectangle the same way.
pub fn classify_arrays(wb: &mut Workbook) {
    let mut capability = Capability::new(wb);
    let mut decisions = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        for (anchor, definition, _) in sheet.array_definitions() {
            if definition.is_opaque() {
                continue;
            }
            let Some(formula) = sheet.cell(anchor).and_then(|cell| cell.formula.as_deref()) else {
                continue;
            };
            let known = capability.evaluable(SheetId(index as u32), formula);
            decisions.push((index, anchor, known));
        }
    }
    for (index, anchor, known) in decisions {
        let sheet = &mut wb.sheets[index];
        if known {
            sheet.adopt_results(anchor);
        } else {
            sheet.set_opaque(anchor);
        }
    }
    fill_uncached(wb);
}

/// an opaque array whose anchor holds no cached value, with the rectangle it
/// fills and whether its formula can be evaluated reproducibly.
struct Uncached {
    sheet: usize,
    anchor: CellRef,
    rectangle: CellRange,
    formula: String,
    deterministic: bool,
}

/// Stores what each opaque array whose anchor holds no cached value shows,
/// over its recorded rectangle. A formula that can read neither the clock nor
/// a random number, in any branch or name, is evaluated once against the
/// workbook as read, within one recalculation's budget and with its date
/// system, a function the engine lacks yielding `#NAME?` as it would at
/// runtime; identical bytes therefore always read the same. Arrays that read
/// one another's rectangles are evaluated in that order, each stored before
/// the ones reading it. Any other array, one in or behind a cycle of them,
/// and one the budget refuses, shows `#NAME?`, as Excel shows an array it
/// cannot evaluate. At most [`MAX_SPILL_CELLS`] cells are filled per
/// workbook; anchors past that fill the anchor alone.
fn fill_uncached(wb: &mut Workbook) {
    let mut room = MAX_SPILL_CELLS;
    let mut capability = Capability::new(wb);
    let mut uncached = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        for (anchor, definition, _) in sheet.array_definitions() {
            let Some(formula) = sheet
                .cell(anchor)
                .filter(|cell| matches!(cell.value, CellValue::Empty))
                .and_then(|cell| cell.formula.as_deref())
                .filter(|_| definition.is_opaque())
            else {
                continue;
            };
            let recorded = definition.entered(anchor);
            let size = ArraySize::of(recorded);
            let cells = size.rows as usize * size.cols as usize;
            let rectangle = if cells <= room {
                room -= cells;
                recorded
            } else {
                CellRange::new(anchor, anchor)
            };
            uncached.push(Uncached {
                sheet: index,
                anchor,
                rectangle,
                formula: formula.to_owned(),
                deterministic: capability.deterministic(SheetId(index as u32), formula),
            });
        }
    }
    if uncached.is_empty() {
        return;
    }
    let budget = Rc::new(EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS));
    let order = evaluation_order(wb, &uncached, &budget);
    let mut ordered = vec![false; uncached.len()];
    for &position in &order {
        ordered[position] = true;
    }
    for (array, _) in uncached
        .iter()
        .zip(&ordered)
        .filter(|(_, ordered)| !**ordered)
    {
        store(wb, array, None);
    }
    for position in order {
        let array = &uncached[position];
        let values = evaluate(wb, array, &budget);
        store(wb, array, values);
    }
}

/// the deterministic arrays among `uncached` in an order where each comes
/// after every one whose rectangle it reads. the work of finding who reads
/// whom is charged to `budget`: each array, and each of its reads against
/// every other array, an array reading nothing costing one. an array in or
/// behind a cycle of them is left out, as is every one once the budget
/// refuses.
fn evaluation_order(wb: &Workbook, uncached: &[Uncached], budget: &EvaluationBudget) -> Vec<usize> {
    let candidates: Vec<usize> = (0..uncached.len())
        .filter(|&position| uncached[position].deterministic)
        .collect();
    let mut graph = DepGraph::empty(wb);
    let mut readers: Vec<Vec<usize>> = vec![Vec::new(); uncached.len()];
    let mut waiting = vec![0usize; uncached.len()];
    for &reader in &candidates {
        let array = &uncached[reader];
        let sheet = SheetId(array.sheet as u32);
        graph.set_formula(sheet, array.anchor, Some(&array.formula));
        let reads = graph.reads(sheet, array.anchor);
        let pairs = reads.len().saturating_mul(candidates.len());
        if !budget.consume(1 + pairs as u64) {
            return Vec::new();
        }
        if reads.is_empty() {
            graph.set_formula(sheet, array.anchor, None);
            continue;
        }
        for &written in &candidates {
            let target = &uncached[written];
            let read = reads.iter().any(|(sheet, range)| {
                sheet.0 as usize == target.sheet && range.overlaps(&target.rectangle)
            });
            if read {
                readers[written].push(reader);
                waiting[reader] += 1;
            }
        }
        graph.set_formula(sheet, array.anchor, None);
    }
    let mut ready: Vec<usize> = candidates
        .iter()
        .copied()
        .filter(|&position| waiting[position] == 0)
        .rev()
        .collect();
    let mut order = Vec::with_capacity(candidates.len());
    while let Some(position) = ready.pop() {
        order.push(position);
        for &reader in readers[position].iter().rev() {
            waiting[reader] -= 1;
            if waiting[reader] == 0 {
                ready.push(reader);
            }
        }
    }
    order
}

/// what `array` shows over its rectangle, `None` when the budget refuses it.
fn evaluate(
    wb: &Workbook,
    array: &Uncached,
    budget: &Rc<EvaluationBudget>,
) -> Option<Vec<CellValue>> {
    let expr = parse_formula(&array.formula).ok()?;
    let sheet = SheetId(array.sheet as u32);
    let mut ctx = EvalContext::with_budget(wb, sheet, Rc::clone(budget));
    ctx.cell = Some(array.anchor);
    ctx.date_system = wb.date_system;
    let spill = evaluate_spill(&expr, &ctx, array.anchor, Some(array.rectangle));
    (!ctx.has_unhandled_budget_error()).then_some(spill.values)
}

/// writes `values`, or `#NAME?` throughout, over `array`'s rectangle as
/// stored cells: its anchor and every cell there nothing else holds.
fn store(wb: &mut Workbook, array: &Uncached, values: Option<Vec<CellValue>>) {
    let size = ArraySize::of(array.rectangle);
    let values = values.unwrap_or_else(|| {
        let name = CellValue::Error {
            value: ErrorValue::Name,
        };
        vec![name; size.rows as usize * size.cols as usize]
    });
    let rectangle = array.rectangle;
    let positions = (rectangle.start.row..=rectangle.end.row).flat_map(|row| {
        (rectangle.start.col..=rectangle.end.col).map(move |col| CellRef::new(row, col))
    });
    let sheet = &mut wb.sheets[array.sheet];
    for (at, value) in positions.zip(values) {
        let stored = sheet.cell(at).cloned().unwrap_or_default();
        let free = matches!(stored.value, CellValue::Empty)
            && stored
                .formula
                .as_deref()
                .is_none_or(|formula| formula.trim().is_empty());
        if (at.row, at.col) == (array.anchor.row, array.anchor.col) || free {
            sheet.set_cell(at, Cell { value, ..stored });
        }
    }
}

/// [`evaluable`] for many formulas of one workbook, looking its defined names
/// up by index. each formula's classification has its own bound on work; a
/// decision about a defined name read by a formula binding no name, and
/// which names reach no binding at all, are settled once and shared.
pub struct Capability<'a> {
    wb: &'a Workbook,
    names: HashMap<(Option<SheetId>, String), usize>,
    parsed: HashMap<usize, Option<Rc<Parsed>>>,
    decided: [HashMap<Node, bool>; 2],
    unbinding: HashSet<Node>,
    work: usize,
    limit: usize,
}

/// what a walk over a formula and the names it reaches asks of each.
#[derive(Clone, Copy)]
enum Check {
    /// every function is one the engine evaluates.
    Supported = 0,
    /// no function reads the clock, draws a random number or is otherwise
    /// volatile to the engine.
    Deterministic = 1,
}

impl Check {
    fn holds(self, parsed: &Parsed) -> bool {
        match self {
            Check::Supported => parsed.supported,
            Check::Deterministic => !parsed.volatile,
        }
    }
}

/// a defined name as the walk expands it: its index and the sheet its
/// unqualified references resolve against.
type Node = (usize, SheetId);

/// what a formula calls, the names it reads and the names its `LET`s and
/// `LAMBDA`s bind, gathered once.
struct Parsed {
    supported: bool,
    volatile: bool,
    reads: Vec<(Option<String>, String)>,
    binds: Vec<String>,
}

/// one defined name being expanded: what its formula reads, how far through
/// that the walk is, where it sits among the names not yet settled, and the
/// earliest name still being expanded that it reaches back to.
struct Frame {
    node: Node,
    order: usize,
    start: usize,
    parsed: Rc<Parsed>,
    next: usize,
    low: usize,
}

impl<'a> Capability<'a> {
    pub fn new(wb: &'a Workbook) -> Self {
        let mut names = HashMap::new();
        for (index, defined) in wb.defined_names.iter().enumerate() {
            names
                .entry((defined.local_sheet, defined.name.to_ascii_lowercase()))
                .or_insert(index);
        }
        Self {
            wb,
            names,
            parsed: HashMap::new(),
            decided: [HashMap::new(), HashMap::new()],
            unbinding: HashSet::new(),
            work: 0,
            limit: MAX_WORK,
        }
    }

    /// [`evaluable`] against this workbook. the evaluator scopes names
    /// dynamically, so a `LAMBDA` or a defined name sees whatever its caller
    /// has bound when it runs: a name that any `LET` or `LAMBDA` in the
    /// formula, or in a defined name it reaches, binds may stand for that
    /// binding and is not followed as a workbook name.
    pub fn evaluable(&mut self, sheet: SheetId, formula: &str) -> bool {
        self.check(sheet, formula, Check::Supported)
    }

    /// whether nothing `formula`, on `sheet`, could run, in any branch or
    /// defined name it reaches, reads the clock, draws a random number or is
    /// otherwise volatile to the engine, so evaluating it twice over the same
    /// cells gives the same result. every name it reads is followed, bound or
    /// not.
    pub fn deterministic(&mut self, sheet: SheetId, formula: &str) -> bool {
        self.check(sheet, formula, Check::Deterministic)
    }

    fn check(&mut self, sheet: SheetId, formula: &str, check: Check) -> bool {
        self.work = 0;
        let Some(parsed) = parse_formula(formula)
            .ok()
            .and_then(|expr| gather(&expr, &mut self.work, self.limit))
        else {
            return false;
        };
        if !check.holds(&parsed) {
            return false;
        }
        let parsed = Rc::new(parsed);
        let bound = match check {
            Check::Supported => match self.bound_anywhere(sheet, &parsed) {
                Some(bound) => bound,
                None => return false,
            },
            Check::Deterministic => HashSet::new(),
        };
        self.names_hold(sheet, parsed, check, &bound)
    }

    /// the index of the name `name` reads as from `sheet`, local before
    /// global and matched ignoring ASCII case, as the evaluator looks it up.
    fn resolve(
        &self,
        sheet: SheetId,
        scope: &Option<String>,
        name: &str,
    ) -> Option<(usize, SheetId)> {
        let lookup = match scope {
            Some(scope) => self.wb.sheet_id(scope)?,
            None => sheet,
        };
        let key = name.to_ascii_lowercase();
        let index = self
            .names
            .get(&(Some(lookup), key.clone()))
            .or_else(|| self.names.get(&(None, key)))
            .copied()?;
        let scope = self.wb.defined_names[index].local_sheet.unwrap_or(lookup);
        Some((index, scope))
    }

    /// the parsed definition of the name at `index`, `Some(None)` when it
    /// does not parse, and `None` when gathering it passes the bound, which
    /// is not kept, so a later formula tries again with its own.
    fn parsed(&mut self, index: usize) -> Option<Option<Rc<Parsed>>> {
        if let Some(parsed) = self.parsed.get(&index) {
            return Some(parsed.clone());
        }
        let formula = &self.wb.defined_names[index].formula;
        let formula = formula.strip_prefix('=').unwrap_or(formula);
        let parsed = match parse_formula(formula) {
            Ok(expr) => Some(Rc::new(gather(&expr, &mut self.work, self.limit)?)),
            Err(_) => None,
        };
        self.parsed.insert(index, parsed.clone());
        Some(parsed)
    }

    /// every name a `LET` or `LAMBDA` binds in `parsed` or in a defined name
    /// reachable from it through any name it reads; `None` past the bound.
    /// the walk keeps one set for the formula, each name added once and
    /// charged, and skips the names already known to reach no binding. a
    /// group of names reaching one another is settled together, as in
    /// [`Capability::names_hold`], and remembered once none of it binds.
    fn bound_anywhere(&mut self, sheet: SheetId, parsed: &Rc<Parsed>) -> Option<HashSet<String>> {
        struct Reach {
            name: Node,
            order: usize,
            start: usize,
            parsed: Rc<Parsed>,
            next: usize,
            low: usize,
            binds: bool,
        }
        let mut bound: HashSet<String> = HashSet::new();
        self.work += parsed.binds.len();
        if self.work > self.limit {
            return None;
        }
        bound.extend(parsed.binds.iter().cloned());
        let mut order: HashMap<Node, usize> = HashMap::new();
        let mut settled: HashSet<Node> = HashSet::new();
        let mut pending: Vec<Node> = Vec::new();
        let mut root = Reach {
            name: (usize::MAX, sheet),
            order: usize::MAX,
            start: 0,
            parsed: Rc::clone(parsed),
            next: 0,
            low: usize::MAX,
            binds: false,
        };
        let mut stack: Vec<Reach> = Vec::new();
        loop {
            let frame = stack.last_mut().unwrap_or(&mut root);
            if frame.next < frame.parsed.reads.len() {
                let (scope, name) = frame.parsed.reads[frame.next].clone();
                frame.next += 1;
                let from = frame.name.1;
                self.work += 1;
                if self.work > self.limit {
                    return None;
                }
                let Some(name) = self.resolve(from, &scope, &name) else {
                    continue;
                };
                if self.unbinding.contains(&name) {
                    continue;
                }
                if settled.contains(&name) {
                    stack.last_mut().unwrap_or(&mut root).binds = true;
                    continue;
                }
                if let Some(&seen) = order.get(&name) {
                    let frame = stack.last_mut().unwrap_or(&mut root);
                    frame.low = frame.low.min(seen);
                    continue;
                }
                let Some(expanded) = self.parsed(name.0)? else {
                    continue;
                };
                self.work += expanded.binds.len();
                if self.work > self.limit {
                    return None;
                }
                bound.extend(expanded.binds.iter().cloned());
                let at = order.len();
                order.insert(name, at);
                let start = pending.len();
                pending.push(name);
                stack.push(Reach {
                    name,
                    order: at,
                    start,
                    binds: !expanded.binds.is_empty(),
                    parsed: expanded,
                    next: 0,
                    low: at,
                });
                continue;
            }
            let Some(done) = stack.pop() else {
                break;
            };
            let parent = stack.last_mut().unwrap_or(&mut root);
            parent.binds |= done.binds;
            if done.low >= done.order {
                for name in pending.drain(done.start..) {
                    order.remove(&name);
                    if done.binds {
                        settled.insert(name);
                    } else {
                        self.unbinding.insert(name);
                    }
                }
            } else {
                parent.low = parent.low.min(done.low);
            }
        }
        Some(bound)
    }

    /// whether every defined name reachable from `parsed`, read on `sheet`,
    /// passes `check`, leaving aside the names in `bound`: a walk without
    /// recursion that settles each group of names reaching one another once
    /// all of it is known, as Tarjan's algorithm does. what it settles for a
    /// formula binding no name is shared with every later formula; anything
    /// else holds for this formula alone.
    fn names_hold(
        &mut self,
        sheet: SheetId,
        parsed: Rc<Parsed>,
        check: Check,
        bound: &HashSet<String>,
    ) -> bool {
        let decided = check as usize;
        let shared = bound.is_empty();
        let mut local: HashMap<Node, bool> = HashMap::new();
        let mut order: HashMap<Node, usize> = HashMap::new();
        let mut pending: Vec<Node> = Vec::new();
        let mut root = Frame {
            node: (usize::MAX, sheet),
            order: usize::MAX,
            start: 0,
            parsed,
            next: 0,
            low: usize::MAX,
        };
        let mut stack: Vec<Frame> = Vec::new();
        loop {
            let frame = stack.last_mut().unwrap_or(&mut root);
            if frame.next < frame.parsed.reads.len() {
                let (scope, name) = frame.parsed.reads[frame.next].clone();
                frame.next += 1;
                let from = frame.node.1;
                self.work += 1;
                if self.work > self.limit {
                    return false;
                }
                if scope.is_none() && bound.contains(&name.to_ascii_lowercase()) {
                    continue;
                }
                let Some(node) = self.resolve(from, &scope, &name) else {
                    continue;
                };
                let known = if shared {
                    self.decided[decided].get(&node).copied()
                } else {
                    local.get(&node).copied()
                };
                match known {
                    Some(true) => continue,
                    Some(false) => return self.refuse(&stack, decided, shared),
                    None => {}
                }
                if let Some(&seen) = order.get(&node) {
                    let frame = stack.last_mut().unwrap_or(&mut root);
                    frame.low = frame.low.min(seen);
                    continue;
                }
                let Some(parsed) = self.parsed(node.0) else {
                    return false;
                };
                let Some(parsed) = parsed.filter(|parsed| check.holds(parsed)) else {
                    if shared {
                        self.decided[decided].insert(node, false);
                    }
                    return self.refuse(&stack, decided, shared);
                };
                let at = order.len();
                order.insert(node, at);
                let start = pending.len();
                pending.push(node);
                stack.push(Frame {
                    node,
                    order: at,
                    start,
                    parsed,
                    next: 0,
                    low: at,
                });
                continue;
            }
            let Some(done) = stack.pop() else {
                break;
            };
            if done.low >= done.order {
                let memo = if shared {
                    &mut self.decided[decided]
                } else {
                    &mut local
                };
                for node in pending.drain(done.start..) {
                    memo.insert(node, true);
                }
            } else {
                let parent = stack.last_mut().unwrap_or(&mut root);
                parent.low = parent.low.min(done.low);
            }
        }
        if shared {
            for node in pending {
                self.decided[decided].insert(node, true);
            }
        }
        true
    }

    /// every name being expanded reaches what fails the check, which is
    /// remembered when the decision is one formulas share.
    fn refuse(&mut self, stack: &[Frame], decided: usize, shared: bool) -> bool {
        if shared {
            for frame in stack {
                self.decided[decided].insert(frame.node, false);
            }
        }
        false
    }
}

/// whether `expr` calls only functions the engine evaluates, whether it calls
/// a volatile one, the names it reads and the names its `LET`s and `LAMBDA`s
/// bind. each expression node is charged to `work`, and the walk gives up
/// once `work` passes `limit`.
fn gather(expr: &Expr, work: &mut usize, limit: usize) -> Option<Parsed> {
    let mut parsed = Parsed {
        supported: true,
        volatile: false,
        reads: Vec::new(),
        binds: Vec::new(),
    };
    let mut pending = vec![expr];
    while let Some(expr) = pending.pop() {
        *work += 1;
        if *work > limit {
            return None;
        }
        match expr {
            Expr::FuncCall { func, name, args } => {
                if func.is_none() && !crate::array::is_array_builtin(name) {
                    parsed.supported = false;
                }
                let binder = crate::functions::bare_name(name).to_ascii_uppercase();
                if crate::graph::VOLATILE_FNS.contains(&binder.as_str())
                    || matches!(binder.as_str(), "RANDARRAY" | "OFFSET")
                {
                    parsed.volatile = true;
                }
                let width = match binder.as_str() {
                    "LET" if args.len() >= 3 && args.len() % 2 == 1 => 2,
                    "LAMBDA" if !args.is_empty() => 1,
                    _ => {
                        pending.extend(args.iter().rev());
                        continue;
                    }
                };
                let (body, heads) = args.split_last().expect("a binder has a body");
                pending.push(body);
                for head in heads.chunks(width).rev() {
                    match &head[0] {
                        Expr::Name { scope: None, name } => {
                            parsed.binds.push(name.to_ascii_lowercase())
                        }
                        other => pending.push(other),
                    }
                    if let Some(value) = head.get(1) {
                        pending.push(value);
                    }
                }
            }
            Expr::Name { scope, name } => parsed.reads.push((scope.clone(), name.clone())),
            Expr::ArrayLiteral { values, .. } => pending.extend(values),
            Expr::Unary { expr, .. } | Expr::Percent(expr) => pending.push(expr),
            Expr::Binary { lhs, rhs, .. } => pending.extend([lhs.as_ref(), rhs.as_ref()]),
            Expr::RangeJoin { start, end } => {
                parsed.volatile = true;
                pending.extend([start.as_ref(), end.as_ref()])
            }
            _ => {}
        }
    }
    Some(parsed)
}

#[cfg(test)]
mod tests {
    use xlsx_model::{DefinedName, Sheet};

    use super::*;

    fn workbook(names: &[(&str, &str)]) -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb.defined_names = names
            .iter()
            .map(|(name, formula)| DefinedName {
                name: (*name).into(),
                formula: (*formula).into(),
                local_sheet: None,
                hidden: false,
            })
            .collect();
        wb
    }

    /// a function the engine lacks makes a formula opaque in a branch no
    /// input takes, behind a fallback and inside a name it reads; an ordinary
    /// error, an unknown name and a cycle of names do not.
    #[test]
    fn a_formula_is_evaluable_only_if_every_branch_is() {
        let wb = workbook(&[
            ("Remote", "WEBSERVICE(\"https://example.com\")"),
            ("Local", "SEQUENCE(3)"),
            ("Loop", "Loop+1"),
            ("Ring", "Back+Remote"),
            ("Back", "Ring"),
        ]);
        for (formula, expected) in [
            ("_xlfn.SEQUENCE(3)", true),
            ("IF(FALSE,WEBSERVICE(\"x\"),_xlfn.SEQUENCE(3))", false),
            ("IFERROR(WEBSERVICE(\"x\"),0)", false),
            ("Remote*2", false),
            ("Local*2", true),
            ("1/0", true),
            ("Missing+1", true),
            ("Loop", true),
            ("Back", false),
            ("LET(x,2,x*_xlfn.SEQUENCE(2))", true),
            ("LET(Remote,A1,Remote*_xlfn.SEQUENCE(3))", true),
            ("LET(x,Remote,x)", false),
            ("LET(Remote,Remote,Remote)", true),
            ("_xlfn.MAP(A1:A2,_xlfn.LAMBDA(Remote,Remote+1))", true),
            ("_xlfn.MAP(A1:A2,_xlfn.LAMBDA(x,Remote+x))", false),
            ("SUM(", false),
        ] {
            assert_eq!(evaluable(&wb, SheetId(0), formula), expected, "{formula}");
        }
    }

    /// a formula is deterministic only if nothing it could run, in a branch
    /// no input takes or a name it reads, reads the clock, draws a random
    /// number or is otherwise volatile to the engine. a name reached through
    /// a binding still counts, since a `LAMBDA` does not keep the bindings
    /// around it once it is called elsewhere.
    #[test]
    fn a_formula_is_deterministic_only_if_every_branch_is() {
        let wb = workbook(&[
            ("Clock", "NOW()"),
            ("Fixed", "SEQUENCE(3)"),
            ("x", "RAND()"),
        ]);
        let mut capability = Capability::new(&wb);
        for (formula, expected) in [
            ("IFERROR(_xlfn.FILTERXML(\"<a/>\",\"//b\"),\"\")", true),
            ("Fixed*2", true),
            ("IF(FALSE,RAND(),1)", false),
            ("IF(FALSE,TODAY(),1)", false),
            ("_xlfn.RANDARRAY(3)", false),
            ("Clock+1", false),
            ("LET(Clock,1,Clock)", false),
            (
                "IF(FALSE,WEBSERVICE(\"x\"),_xlfn.MAP({1},_xlfn.LET(x,0,_xlfn.LAMBDA(y,x))))",
                false,
            ),
            ("LET(y,1,y)", true),
        ] {
            assert_eq!(
                capability.deterministic(SheetId(0), formula),
                expected,
                "{formula}"
            );
        }
    }

    /// the evaluator scopes names dynamically, so a name any `LET` or `LAMBDA`
    /// in the formula binds, even one a `LAMBDA` stored earlier reads when
    /// it is called later, may stand for that binding inside every defined
    /// name the formula reaches. a formula binding no such name still reaches
    /// the workbook name, whichever is asked first.
    #[test]
    fn a_defined_name_is_read_within_the_bindings_around_it() {
        let wb = workbook(&[
            ("Remote", "WEBSERVICE(\"x\")"),
            ("Alias", "Remote"),
            ("Chain", "Alias+1"),
        ]);
        let questions = [
            ("LET(Remote,A1,Alias*_xlfn.SEQUENCE(3))", true),
            ("_xlfn.MAP(A1:A2,_xlfn.LAMBDA(Remote,Alias*2))", true),
            ("LET(Remote,1,Chain)", true),
            (
                "LET(fn,_xlfn.LAMBDA(x,Alias*x),Remote,A1,_xlfn.MAP({1;2;3},fn))",
                true,
            ),
            ("LET(Remote,1,Alias)+Alias", true),
            ("Alias", false),
            ("Chain", false),
            ("LET(Other,1,Alias)", false),
        ];
        for rotation in 0..questions.len() {
            let mut capability = Capability::new(&wb);
            for step in 0..questions.len() {
                let (formula, expected) = questions[(rotation + step) % questions.len()];
                assert_eq!(
                    capability.evaluable(SheetId(0), formula),
                    expected,
                    "{formula} after rotation {rotation}"
                );
            }
        }
    }

    /// a name resolves as the evaluator resolves it: local before global,
    /// and ignoring ASCII case only, so a local `Ö` does not stand in for a
    /// global `ö`.
    #[test]
    fn names_resolve_as_the_evaluator_resolves_them() {
        let mut wb = workbook(&[("ö", "RAND()")]);
        wb.defined_names.push(DefinedName {
            name: "Ö".into(),
            formula: "1".into(),
            local_sheet: Some(SheetId(0)),
            hidden: false,
        });
        let formula = "IF(FALSE,WEBSERVICE(\"x\"),ö)";
        assert_eq!(
            wb.defined_name(SheetId(0), "ö")
                .map(|name| name.formula.as_str()),
            Some("RAND()")
        );
        assert!(!Capability::new(&wb).deterministic(SheetId(0), formula));
        let mut wb = workbook(&[("Mixed", "RAND()")]);
        wb.defined_names.push(DefinedName {
            name: "MIXED".into(),
            formula: "1".into(),
            local_sheet: Some(SheetId(0)),
            hidden: false,
        });
        assert!(Capability::new(&wb).deterministic(SheetId(0), "mixed"));
    }

    /// every step of the walk, each binding and each name it looks up, is
    /// charged to one bound, so a `LET` as long as the parser admits costs
    /// work linear in its size, and a formula past the bound is neither
    /// evaluable nor deterministic.
    #[test]
    fn a_long_let_is_charged_as_it_is_walked() {
        const PAIRS: usize = 2_400;
        let formula = format!("_xlfn.LET(x,1,{}x)", "y,x,".repeat(PAIRS));
        let wb = workbook(&[]);
        let mut capability = Capability::new(&wb);
        let start = std::time::Instant::now();
        assert!(capability.evaluable(SheetId(0), &formula));
        assert!(capability.deterministic(SheetId(0), &formula));
        assert!(start.elapsed().as_secs_f64() < 5.0);
        assert!(capability.work <= 2 * 8 * PAIRS, "{}", capability.work);
        let mut capability = Capability::new(&wb);
        capability.limit = PAIRS;
        assert!(!capability.evaluable(SheetId(0), &formula));
        assert!(!capability.deterministic(SheetId(0), &formula));
    }

    /// deciding a name settles every name it reaches, cycles included, so
    /// the same names read again later cost no expansion, and an answer never
    /// depends on which formula asked first.
    #[test]
    fn names_are_decided_once_whatever_asks_first() {
        let wb = workbook(&[
            ("A", "B+C"),
            ("B", "A+1"),
            ("C", "D"),
            ("D", "1"),
            ("E", "C+Remote"),
            ("Remote", "WEBSERVICE(\"x\")"),
        ]);
        let questions = ["A", "B", "C", "E", "A+E"];
        let answers = [true, true, true, false, false];
        for rotation in 0..questions.len() {
            let mut capability = Capability::new(&wb);
            for step in 0..questions.len() {
                let index = (rotation + step) % questions.len();
                assert_eq!(
                    capability.evaluable(SheetId(0), questions[index]),
                    answers[index],
                    "{}",
                    questions[index]
                );
            }
            let parsed = capability.parsed.len();
            let decided = capability.decided[0].len();
            for (question, answer) in questions.iter().zip(answers) {
                assert_eq!(capability.evaluable(SheetId(0), question), answer);
            }
            assert_eq!(capability.parsed.len(), parsed);
            assert_eq!(capability.decided[0].len(), decided);
        }
    }

    fn uncached(sheet: &mut Sheet, address: &str, branch: &str) {
        let anchor = CellRef::parse_a1(address).unwrap();
        sheet.set_cell(
            anchor,
            Cell {
                formula: Some(format!("IF(FALSE,WEBSERVICE(\"x\"),{branch})")),
                ..Cell::default()
            },
        );
        sheet.set_dynamic_array_formula(anchor, CellRange::new(anchor, anchor));
    }

    fn shown(wb: &Workbook, address: &str) -> CellValue {
        wb.value(SheetId(0), CellRef::parse_a1(address).unwrap())
    }

    /// uncached opaque arrays reading one another's rectangles are evaluated
    /// in that order, each stored before the ones reading it, whatever order
    /// they sit in; arrays reading each other round a cycle show `#NAME?`.
    #[test]
    fn uncached_arrays_are_evaluated_in_dependency_order() {
        let mut wb = workbook(&[]);
        let sheet = &mut wb.sheets[0];
        uncached(sheet, "B1", "C1+1");
        uncached(sheet, "C1", "D1*2");
        uncached(sheet, "D1", "7");
        uncached(sheet, "E1", "F1+1");
        uncached(sheet, "F1", "E1+1");
        classify_arrays(&mut wb);
        let name = CellValue::Error {
            value: ErrorValue::Name,
        };
        for (address, value) in [
            ("B1", CellValue::Number { value: 15.0 }),
            ("C1", CellValue::Number { value: 14.0 }),
            ("D1", CellValue::Number { value: 7.0 }),
            ("E1", name.clone()),
            ("F1", name),
        ] {
            assert_eq!(shown(&wb, address), value, "{address}");
        }
    }

    /// each formula's classification has its own bound on work, so formulas
    /// that each cost much, together past a bound one pass could share, leave
    /// every array after them evaluable.
    #[test]
    fn many_costly_formulas_leave_later_arrays_evaluable() {
        const ANCHORS: u32 = 640;
        let costly = format!("_xlfn.LET(x,1,{}x)", "y,x,".repeat(2_400));
        let mut wb = workbook(&[]);
        let mut capability = Capability::new(&wb);
        let spent: usize = (0..ANCHORS)
            .map(|_| {
                assert!(capability.evaluable(SheetId(0), &costly));
                capability.work
            })
            .sum();
        assert!(spent > 4 << 20, "{spent}");
        let sheet = &mut wb.sheets[0];
        for row in 0..=ANCHORS {
            let anchor = CellRef::new(row, 0);
            let formula = if row < ANCHORS {
                costly.clone()
            } else {
                "_xlfn.SEQUENCE(1)".to_owned()
            };
            sheet.set_cell(
                anchor,
                Cell {
                    formula: Some(formula),
                    value: CellValue::Number { value: 1.0 },
                    style: None,
                },
            );
            sheet.set_dynamic_array_formula(anchor, CellRange::new(anchor, anchor));
        }
        classify_arrays(&mut wb);
        for row in 0..=ANCHORS {
            let definition = wb.sheets[0].array_definition(CellRef::new(row, 0)).unwrap();
            assert!(!definition.is_opaque(), "row {row}");
        }
    }

    /// a chain of names each binding a name of its own is walked with one set
    /// of bound names for the formula, not one per name, so it costs time
    /// and memory linear in its length.
    #[test]
    fn a_chain_of_binding_names_is_walked_with_one_set() {
        const LENGTH: usize = 30_000;
        let names: Vec<(String, String)> = (0..LENGTH)
            .map(|index| {
                let formula = if index + 1 == LENGTH {
                    "1".to_owned()
                } else {
                    format!("LET(bound_{index},0,step_{})", index + 1)
                };
                (format!("step_{index}"), formula)
            })
            .collect();
        let names: Vec<(&str, &str)> = names
            .iter()
            .map(|(name, formula)| (name.as_str(), formula.as_str()))
            .collect();
        let wb = workbook(&names);
        let mut capability = Capability::new(&wb);
        let start = std::time::Instant::now();
        assert!(capability.evaluable(SheetId(0), "step_0"));
        assert!(capability.work <= 16 * LENGTH, "{}", capability.work);
        assert!(capability.decided[0].is_empty());
        assert!(start.elapsed().as_secs_f64() < 5.0, "{:?}", start.elapsed());
    }

    /// many uncached arrays that read nothing are ordered in time linear in
    /// how many there are, and each shows what it evaluates to.
    #[test]
    fn many_uncached_arrays_reading_nothing_are_ordered_promptly() {
        const ANCHORS: u32 = 100_000;
        let mut wb = workbook(&[]);
        let sheet = &mut wb.sheets[0];
        for row in 0..ANCHORS {
            uncached(sheet, &CellRef::new(row, 0).to_a1(), "1");
        }
        let start = std::time::Instant::now();
        classify_arrays(&mut wb);
        assert!(
            start.elapsed().as_secs_f64() < 60.0,
            "{:?}",
            start.elapsed()
        );
        for row in [0, ANCHORS - 1] {
            let at = CellRef::new(row, 0);
            assert_eq!(wb.value(SheetId(0), at), CellValue::Number { value: 1.0 });
        }
    }

    /// a chain of names as long as a workbook may hold is walked without
    /// recursion, in time linear in its length, whether it ends in a function
    /// the engine has or one it lacks, and once only.
    #[test]
    fn a_long_chain_of_names_is_walked_once() {
        const LENGTH: usize = 60_000;
        for last in ["1", "WEBSERVICE(\"x\")"] {
            let names: Vec<(String, String)> = (0..LENGTH)
                .map(|index| {
                    let formula = if index + 1 == LENGTH {
                        last.to_owned()
                    } else {
                        format!("step_{}+1", index + 1)
                    };
                    (format!("step_{index}"), formula)
                })
                .collect();
            let names: Vec<(&str, &str)> = names
                .iter()
                .map(|(name, formula)| (name.as_str(), formula.as_str()))
                .collect();
            let wb = workbook(&names);
            let mut capability = Capability::new(&wb);
            let start = std::time::Instant::now();
            let supported = last == "1";
            assert_eq!(capability.evaluable(SheetId(0), "step_0"), supported);
            assert_eq!(capability.evaluable(SheetId(0), "step_30000"), supported);
            let elapsed = start.elapsed();
            assert!(capability.work <= 8 * LENGTH, "{last}");
            assert!(elapsed.as_secs_f64() < 5.0, "{last}: {elapsed:?}");
        }
    }
}
