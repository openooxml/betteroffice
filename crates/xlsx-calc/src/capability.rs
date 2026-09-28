//! whether the engine can evaluate a formula wherever it could run, decided
//! from the formula alone so that no input changes the answer.

use std::collections::HashMap;
use std::rc::Rc;

use xlsx_model::{
    ArraySize, Cell, CellProvider, CellRange, CellRef, CellValue, ErrorValue, MAX_SPILL_CELLS,
    SheetId, Workbook,
};

use crate::array::evaluate_spill;
use crate::eval::{EvalContext, EvaluationBudget, MAX_RECALCULATION_CELL_VISITS};
use crate::parser::{Expr, parse_formula};

/// analysis steps one [`Capability`] takes, each expression node, binding
/// and name reference it visits, before it answers no for whatever still
/// needs more.
const MAX_WORK: usize = 1 << 22;

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

/// Stores what each opaque array whose anchor holds no cached value shows,
/// over its recorded rectangle. A formula that can read neither the clock nor
/// a random number, in any branch or name, is evaluated once against the
/// workbook as read, within one recalculation's budget and with its date
/// system, a function the engine lacks yielding `#NAME?` as it would at
/// runtime; identical bytes therefore always read the same. Any other, and
/// one the budget refuses, shows `#NAME?`, as Excel shows an array it cannot
/// evaluate. At most [`MAX_SPILL_CELLS`] cells are filled per workbook;
/// anchors past that fill the anchor alone.
fn fill_uncached(wb: &mut Workbook) {
    let mut room = MAX_SPILL_CELLS;
    let mut capability = Capability::new(wb);
    let budget = Rc::new(EvaluationBudget::new(MAX_RECALCULATION_CELL_VISITS));
    let mut fills: Vec<(usize, CellRef, CellRange, Vec<CellValue>)> = Vec::new();
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
            let sheet_id = SheetId(index as u32);
            let evaluated = capability
                .deterministic(sheet_id, formula)
                .then(|| parse_formula(formula).ok())
                .flatten()
                .and_then(|expr| {
                    let mut ctx = EvalContext::with_budget(&*wb, sheet_id, Rc::clone(&budget));
                    ctx.cell = Some(anchor);
                    ctx.date_system = wb.date_system;
                    let spill = evaluate_spill(&expr, &ctx, anchor, Some(rectangle));
                    (!ctx.has_unhandled_budget_error()).then_some(spill.values)
                });
            let name = CellValue::Error {
                value: ErrorValue::Name,
            };
            let positions = ArraySize::of(rectangle);
            let values = evaluated
                .unwrap_or_else(|| vec![name; positions.rows as usize * positions.cols as usize]);
            fills.push((index, anchor, rectangle, values));
        }
    }
    for (index, anchor, rectangle, values) in fills {
        let sheet = &mut wb.sheets[index];
        let positions = (rectangle.start.row..=rectangle.end.row).flat_map(|row| {
            (rectangle.start.col..=rectangle.end.col).map(move |col| CellRef::new(row, col))
        });
        for (at, value) in positions.zip(values) {
            let stored = sheet.cell(at).cloned().unwrap_or_default();
            let free = matches!(stored.value, CellValue::Empty)
                && stored
                    .formula
                    .as_deref()
                    .is_none_or(|formula| formula.trim().is_empty());
            if (at.row, at.col) == (anchor.row, anchor.col) || free {
                sheet.set_cell(at, Cell { value, ..stored });
            }
        }
    }
}

/// [`evaluable`] for many formulas of one workbook, looking its defined names
/// up by index and deciding each name, for the sheet it is read from and the
/// bindings around it, once.
pub struct Capability<'a> {
    wb: &'a Workbook,
    names: HashMap<(Option<SheetId>, String), usize>,
    parsed: HashMap<usize, Option<Rc<Parsed>>>,
    decided: [HashMap<Node, bool>; 2],
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

/// a defined name as the walk expands it: its index, the sheet its
/// unqualified references resolve against, and the names bound where it is
/// read, which the evaluator keeps bound inside it.
type Node = (usize, SheetId, Rc<[String]>);

/// a name a formula reads: the sheet qualifying it, if any, and the
/// innermost of the formula's bindings around it.
#[derive(Clone)]
struct Reference {
    scope: Option<String>,
    name: String,
    within: Option<usize>,
}

/// what a formula calls and names, gathered once: the names no `LET` or
/// `LAMBDA` in it binds, every name it reads, bound or not, and its bindings,
/// each with the one around it.
struct Parsed {
    supported: bool,
    volatile: bool,
    names: Vec<Reference>,
    every: Vec<Reference>,
    bindings: Vec<(Option<usize>, String)>,
}

impl Parsed {
    /// the names `check` follows: a binding shadows a workbook name for
    /// evaluation, but whether a `LAMBDA` sees the binding around it depends
    /// on where it is called, so determinism follows every name.
    fn names(&self, check: Check) -> &[Reference] {
        match check {
            Check::Supported => &self.names,
            Check::Deterministic => &self.every,
        }
    }
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
            work: 0,
            limit: MAX_WORK,
        }
    }

    pub fn evaluable(&mut self, sheet: SheetId, formula: &str) -> bool {
        self.check(sheet, formula, Check::Supported)
    }

    /// whether nothing `formula`, on `sheet`, could run, in any branch or
    /// defined name it reaches, reads the clock, draws a random number or is
    /// otherwise volatile to the engine, so evaluating it twice over the same
    /// cells gives the same result.
    pub fn deterministic(&mut self, sheet: SheetId, formula: &str) -> bool {
        self.check(sheet, formula, Check::Deterministic)
    }

    fn check(&mut self, sheet: SheetId, formula: &str, check: Check) -> bool {
        let Some(parsed) = parse_formula(formula)
            .ok()
            .and_then(|expr| gather(&expr, &mut self.work, self.limit))
        else {
            return false;
        };
        check.holds(&parsed) && self.names_hold(sheet, Rc::new(parsed), check)
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

    fn parsed(&mut self, index: usize) -> Option<Rc<Parsed>> {
        let (wb, work, limit) = (self.wb, &mut self.work, self.limit);
        self.parsed
            .entry(index)
            .or_insert_with(|| {
                let formula = &wb.defined_names[index].formula;
                let formula = formula.strip_prefix('=').unwrap_or(formula);
                parse_formula(formula)
                    .ok()
                    .and_then(|expr| gather(&expr, work, limit))
                    .map(Rc::new)
            })
            .clone()
    }

    /// the names bound where `reference` sits in `parsed`, beside the ones
    /// `outer` already holds, sorted; each is charged, `None` past the bound.
    fn bound_around(
        &mut self,
        outer: &[String],
        parsed: &Parsed,
        reference: &Reference,
    ) -> Option<Rc<[String]>> {
        let mut names = outer.to_vec();
        let mut within = reference.within;
        while let Some(at) = within {
            let (parent, name) = &parsed.bindings[at];
            names.push(name.clone());
            within = *parent;
        }
        self.work += names.len();
        if self.work > self.limit {
            return None;
        }
        names.sort_unstable();
        names.dedup();
        Some(names.into())
    }

    /// whether every defined name reachable from `parsed`, read on `sheet`,
    /// passes `check`: a walk without recursion that settles each group of
    /// names reaching one another once all of it is known, as Tarjan's
    /// algorithm does, so a later formula reading any of them within the same
    /// bindings costs nothing. for evaluability, a name bound where a defined
    /// name is read stays bound inside it, as the evaluator expands a name
    /// within its caller's bindings.
    fn names_hold(&mut self, sheet: SheetId, parsed: Rc<Parsed>, check: Check) -> bool {
        let decided = check as usize;
        let unbound: Rc<[String]> = Rc::from(Vec::new());
        let mut order: HashMap<Node, usize> = HashMap::new();
        let mut pending: Vec<Node> = Vec::new();
        let mut root = Frame {
            node: (usize::MAX, sheet, Rc::clone(&unbound)),
            order: usize::MAX,
            start: 0,
            parsed,
            next: 0,
            low: usize::MAX,
        };
        let mut stack: Vec<Frame> = Vec::new();
        loop {
            let frame = stack.last_mut().unwrap_or(&mut root);
            if frame.next < frame.parsed.names(check).len() {
                let reference = frame.parsed.names(check)[frame.next].clone();
                frame.next += 1;
                let (from, outer, parsed) = (
                    frame.node.1,
                    Rc::clone(&frame.node.2),
                    Rc::clone(&frame.parsed),
                );
                self.work += 1;
                if self.work > self.limit {
                    return false;
                }
                let caller_binds = reference.scope.is_none()
                    && outer
                        .binary_search(&reference.name.to_ascii_lowercase())
                        .is_ok();
                if caller_binds {
                    continue;
                }
                let Some((index, scope)) = self.resolve(from, &reference.scope, &reference.name)
                else {
                    continue;
                };
                let around = match check {
                    Check::Supported if !outer.is_empty() || reference.within.is_some() => {
                        match self.bound_around(&outer, &parsed, &reference) {
                            Some(around) => around,
                            None => return false,
                        }
                    }
                    _ => Rc::clone(&unbound),
                };
                let node = (index, scope, around);
                match self.decided[decided].get(&node) {
                    Some(true) => continue,
                    Some(false) => return self.refuse(&stack, decided),
                    None => {}
                }
                if let Some(&seen) = order.get(&node) {
                    let frame = stack.last_mut().unwrap_or(&mut root);
                    frame.low = frame.low.min(seen);
                    continue;
                }
                let Some(parsed) = self.parsed(index).filter(|parsed| check.holds(parsed)) else {
                    self.decided[decided].insert(node, false);
                    return self.refuse(&stack, decided);
                };
                let at = order.len();
                order.insert(node.clone(), at);
                let start = pending.len();
                pending.push(node.clone());
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
                for node in pending.drain(done.start..) {
                    self.decided[decided].insert(node, true);
                }
            } else {
                let parent = stack.last_mut().unwrap_or(&mut root);
                parent.low = parent.low.min(done.low);
            }
        }
        for node in pending {
            self.decided[decided].insert(node, true);
        }
        true
    }

    /// every name being expanded reaches what fails the check.
    fn refuse(&mut self, stack: &[Frame], decided: usize) -> bool {
        for frame in stack {
            self.decided[decided].insert(frame.node.clone(), false);
        }
        false
    }
}

/// one step of [`gather`]'s walk: an expression to visit, or a `LET` or
/// `LAMBDA` name coming into or going out of scope.
enum Step<'e> {
    Visit(&'e Expr),
    Bind(&'e str),
    Unbind(&'e str),
}

/// whether `expr` calls only functions the engine evaluates, whether it calls
/// a volatile one, and the names it reads. each step is charged to `work`,
/// and the walk gives up once `work` passes `limit`.
fn gather(expr: &Expr, work: &mut usize, limit: usize) -> Option<Parsed> {
    let mut parsed = Parsed {
        supported: true,
        volatile: false,
        names: Vec::new(),
        every: Vec::new(),
        bindings: Vec::new(),
    };
    let mut bound: HashMap<String, usize> = HashMap::new();
    let mut within: Option<usize> = None;
    let mut pending = vec![Step::Visit(expr)];
    while let Some(step) = pending.pop() {
        *work += 1;
        if *work > limit {
            return None;
        }
        let expr = match step {
            Step::Visit(expr) => expr,
            Step::Bind(name) => {
                let name = name.to_ascii_lowercase();
                *bound.entry(name.clone()).or_default() += 1;
                parsed.bindings.push((within, name));
                within = Some(parsed.bindings.len() - 1);
                continue;
            }
            Step::Unbind(name) => {
                if let Some(count) = bound.get_mut(&name.to_ascii_lowercase()) {
                    *count -= 1;
                }
                within = within.and_then(|at| parsed.bindings[at].0);
                continue;
            }
        };
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
                let binds = match binder.as_str() {
                    "LET" => args.len() >= 3 && args.len() % 2 == 1,
                    "LAMBDA" => !args.is_empty(),
                    _ => false,
                };
                let Some((body, heads)) = args.split_last().filter(|_| binds) else {
                    pending.extend(args.iter().rev().map(Step::Visit));
                    continue;
                };
                let width = if binder == "LET" { 2 } else { 1 };
                for head in heads.chunks(width) {
                    if let Expr::Name { scope: None, name } = &head[0] {
                        pending.push(Step::Unbind(name));
                    }
                }
                pending.push(Step::Visit(body));
                for head in heads.chunks(width).rev() {
                    match &head[0] {
                        Expr::Name { scope: None, name } => pending.push(Step::Bind(name)),
                        other => pending.push(Step::Visit(other)),
                    }
                    if let Some(value) = head.get(1) {
                        pending.push(Step::Visit(value));
                    }
                }
            }
            Expr::Name { scope, name } => {
                let reference = Reference {
                    scope: scope.clone(),
                    name: name.clone(),
                    within,
                };
                let shadowed = scope.is_none()
                    && bound
                        .get(&name.to_ascii_lowercase())
                        .is_some_and(|count| *count > 0);
                if !shadowed {
                    parsed.names.push(reference.clone());
                }
                parsed.every.push(reference);
            }
            Expr::ArrayLiteral { values, .. } => pending.extend(values.iter().map(Step::Visit)),
            Expr::Unary { expr, .. } | Expr::Percent(expr) => pending.push(Step::Visit(expr)),
            Expr::Binary { lhs, rhs, .. } => pending.extend([Step::Visit(lhs), Step::Visit(rhs)]),
            Expr::RangeJoin { start, end } => {
                parsed.volatile = true;
                pending.extend([Step::Visit(start), Step::Visit(end)])
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
            ("LET(Remote,Remote,Remote)", false),
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

    /// the evaluator expands a defined name within the bindings where it is
    /// read, so a name a `LET` or `LAMBDA` binds there stays bound inside
    /// the definition, through any chain of aliases; the same name read
    /// unbound still reaches the workbook name, whichever is asked first.
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
            ("Alias", false),
            ("Chain", false),
            ("LET(Remote,1,Alias)+Alias", false),
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
            let work = capability.work;
            for (question, answer) in questions.iter().zip(answers) {
                assert_eq!(capability.evaluable(SheetId(0), question), answer);
            }
            assert!(capability.work - work <= 6 * 4);
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
