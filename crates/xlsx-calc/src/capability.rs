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

/// name references one [`Capability`] visits before it answers no for
/// whatever still needs more.
const MAX_VISITS: usize = 1 << 22;

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
/// up by index and deciding each name, for the sheet it is read from, once.
pub struct Capability<'a> {
    wb: &'a Workbook,
    names: HashMap<(Option<SheetId>, String), usize>,
    parsed: HashMap<usize, Option<Parsed>>,
    decided: [HashMap<Name, bool>; 2],
    visits: usize,
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

/// a defined name, by its index, with the sheet its unqualified references
/// resolve against.
type Name = (usize, SheetId);

/// what a formula calls and names, gathered once.
struct Parsed {
    supported: bool,
    volatile: bool,
    names: Vec<(Option<String>, String)>,
}

/// one defined name being expanded: the names its formula reads, how far
/// through them the walk is, where it sits among the names not yet settled,
/// and the earliest name still being expanded that it reaches back to.
struct Frame {
    name: Name,
    order: usize,
    start: usize,
    names: Vec<(Option<String>, String)>,
    next: usize,
    low: usize,
}

impl<'a> Capability<'a> {
    pub fn new(wb: &'a Workbook) -> Self {
        let mut names = HashMap::new();
        for (index, defined) in wb.defined_names.iter().enumerate() {
            names
                .entry((defined.local_sheet, defined.name.to_lowercase()))
                .or_insert(index);
        }
        Self {
            wb,
            names,
            parsed: HashMap::new(),
            decided: [HashMap::new(), HashMap::new()],
            visits: 0,
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
        match parse_formula(formula) {
            Ok(expr) => {
                let parsed = gather(&expr);
                check.holds(&parsed) && self.names_hold(sheet, parsed.names, check)
            }
            Err(_) => false,
        }
    }

    /// the index of the name `name` reads as from `sheet`, local before
    /// global.
    fn resolve(&self, sheet: SheetId, scope: &Option<String>, name: &str) -> Option<Name> {
        let lookup = match scope {
            Some(scope) => self.wb.sheet_id(scope)?,
            None => sheet,
        };
        let key = name.to_lowercase();
        let index = self
            .names
            .get(&(Some(lookup), key.clone()))
            .or_else(|| self.names.get(&(None, key)))
            .copied()?;
        let scope = self.wb.defined_names[index].local_sheet.unwrap_or(lookup);
        Some((index, scope))
    }

    fn parsed(&mut self, index: usize) -> Option<&Parsed> {
        self.parsed
            .entry(index)
            .or_insert_with(|| {
                let formula = &self.wb.defined_names[index].formula;
                let formula = formula.strip_prefix('=').unwrap_or(formula);
                parse_formula(formula).ok().map(|expr| gather(&expr))
            })
            .as_ref()
    }

    /// whether every defined name reachable from `names`, read on `sheet`,
    /// passes `check`: a walk without recursion that settles each group of
    /// names reaching one another once all of it is known, as Tarjan's
    /// algorithm does, so a later formula reading any of them costs nothing.
    fn names_hold(
        &mut self,
        sheet: SheetId,
        names: Vec<(Option<String>, String)>,
        check: Check,
    ) -> bool {
        let decided = check as usize;
        let mut order: HashMap<Name, usize> = HashMap::new();
        let mut pending: Vec<Name> = Vec::new();
        let mut root = Frame {
            name: (usize::MAX, sheet),
            order: usize::MAX,
            start: 0,
            names,
            next: 0,
            low: usize::MAX,
        };
        let mut stack: Vec<Frame> = Vec::new();
        loop {
            let frame = stack.last_mut().unwrap_or(&mut root);
            if frame.next < frame.names.len() {
                let (scope, name) = &frame.names[frame.next];
                frame.next += 1;
                self.visits += 1;
                if self.visits > MAX_VISITS {
                    return false;
                }
                let from = frame.name.1;
                let Some(name) = self.resolve(from, scope, name) else {
                    continue;
                };
                match self.decided[decided].get(&name) {
                    Some(true) => continue,
                    Some(false) => return self.refuse(&stack, decided),
                    None => {}
                }
                if let Some(&seen) = order.get(&name) {
                    let frame = stack.last_mut().unwrap_or(&mut root);
                    frame.low = frame.low.min(seen);
                    continue;
                }
                let Some(parsed) = self.parsed(name.0).filter(|parsed| check.holds(parsed)) else {
                    self.decided[decided].insert(name, false);
                    return self.refuse(&stack, decided);
                };
                let names = parsed.names.clone();
                let at = order.len();
                order.insert(name, at);
                let start = pending.len();
                pending.push(name);
                stack.push(Frame {
                    name,
                    order: at,
                    start,
                    names,
                    next: 0,
                    low: at,
                });
                continue;
            }
            let Some(done) = stack.pop() else {
                break;
            };
            if done.low >= done.order {
                for name in pending.drain(done.start..) {
                    self.decided[decided].insert(name, true);
                }
            } else {
                let parent = stack.last_mut().unwrap_or(&mut root);
                parent.low = parent.low.min(done.low);
            }
        }
        for name in pending {
            self.decided[decided].insert(name, true);
        }
        true
    }

    /// every name being expanded reaches what fails the check.
    fn refuse(&mut self, stack: &[Frame], decided: usize) -> bool {
        for frame in stack {
            self.decided[decided].insert(frame.name, false);
        }
        false
    }
}

/// whether `expr` calls only functions the engine evaluates, and the names
/// it reads that no `LET` or `LAMBDA` around them binds.
fn gather(expr: &Expr) -> Parsed {
    let mut parsed = Parsed {
        supported: true,
        volatile: false,
        names: Vec::new(),
    };
    let mut bindings: Vec<(Option<usize>, &str)> = Vec::new();
    let mut pending = vec![(expr, None)];
    while let Some((expr, scope)) = pending.pop() {
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
                    pending.extend(args.iter().map(|arg| (arg, scope)));
                    continue;
                };
                let mut inner = scope;
                let step = if binder == "LET" { 2 } else { 1 };
                for head in heads.chunks(step) {
                    if let Some(value) = head.get(1) {
                        pending.push((value, inner));
                    }
                    match &head[0] {
                        Expr::Name { scope: None, name } => {
                            bindings.push((inner, name));
                            inner = Some(bindings.len() - 1);
                        }
                        other => pending.push((other, inner)),
                    }
                }
                pending.push((body, inner));
            }
            Expr::Name { scope: None, name } if bound(&bindings, scope, name) => {}
            Expr::Name { scope, name } => parsed.names.push((scope.clone(), name.clone())),
            Expr::ArrayLiteral { values, .. } => {
                pending.extend(values.iter().map(|value| (value, scope)))
            }
            Expr::Unary { expr, .. } | Expr::Percent(expr) => pending.push((expr, scope)),
            Expr::Binary { lhs, rhs, .. } => {
                pending.extend([(lhs.as_ref(), scope), (rhs.as_ref(), scope)])
            }
            Expr::RangeJoin { start, end } => {
                parsed.volatile = true;
                pending.extend([(start.as_ref(), scope), (end.as_ref(), scope)])
            }
            _ => {}
        }
    }
    parsed
}

/// whether a `LET` or `LAMBDA` enclosing `scope` binds `name`.
fn bound(bindings: &[(Option<usize>, &str)], mut scope: Option<usize>, name: &str) -> bool {
    while let Some(at) = scope {
        let (parent, bound) = bindings[at];
        if bound.eq_ignore_ascii_case(name) {
            return true;
        }
        scope = parent;
    }
    false
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
    /// number or is otherwise volatile to the engine.
    #[test]
    fn a_formula_is_deterministic_only_if_every_branch_is() {
        let wb = workbook(&[("Clock", "NOW()"), ("Fixed", "SEQUENCE(3)")]);
        let mut capability = Capability::new(&wb);
        for (formula, expected) in [
            ("IFERROR(_xlfn.FILTERXML(\"<a/>\",\"//b\"),\"\")", true),
            ("Fixed*2", true),
            ("IF(FALSE,RAND(),1)", false),
            ("IF(FALSE,TODAY(),1)", false),
            ("_xlfn.RANDARRAY(3)", false),
            ("Clock+1", false),
            ("LET(Clock,1,Clock)", true),
        ] {
            assert_eq!(
                capability.deterministic(SheetId(0), formula),
                expected,
                "{formula}"
            );
        }
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
            let visits = capability.visits;
            for (question, answer) in questions.iter().zip(answers) {
                assert_eq!(capability.evaluable(SheetId(0), question), answer);
            }
            assert!(capability.visits - visits <= 6);
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
            assert_eq!(capability.visits, LENGTH + 1, "{last}");
            assert!(elapsed.as_secs_f64() < 5.0, "{last}: {elapsed:?}");
        }
    }
}
