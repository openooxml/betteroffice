//! whether the engine can evaluate a formula wherever it could run, decided
//! from the formula alone so that no input changes the answer.

use std::collections::HashMap;

use xlsx_model::{
    ArrayDefinition, ArrayKind, Cell, CellProvider, CellRange, CellRef, CellValue, SheetId,
    Workbook,
};

use crate::parser::{Expr, parse_formula};

/// defined-name expansions one [`Capability`] performs before it answers no
/// for whatever still needs more.
const MAX_EXPANSIONS: usize = 1 << 22;

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
/// opaque array whose anchor holds no cached value is evaluated once over its
/// recorded rectangle, and what that shows is stored there too, so the
/// workbook reads as Excel shows it without the engine ever owning those
/// cells.
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

/// what the engine shows for each opaque array whose anchor holds no cached
/// value, evaluated over its recorded rectangle as a legacy array would be,
/// written into those cells as stored values.
fn fill_uncached(wb: &mut Workbook) {
    let uncached: Vec<(usize, CellRef, ArrayDefinition)> = wb
        .sheets
        .iter()
        .enumerate()
        .flat_map(|(index, sheet)| {
            sheet
                .array_definitions()
                .filter(|(anchor, definition, _)| {
                    definition.is_opaque()
                        && sheet.cell(*anchor).is_some_and(|cell| {
                            cell.formula.is_some() && matches!(cell.value, CellValue::Empty)
                        })
                })
                .map(move |(anchor, definition, _)| (index, anchor, definition))
        })
        .collect();
    if uncached.is_empty() {
        return;
    }
    let mut evaluated = wb.clone();
    for &(index, anchor, definition) in &uncached {
        let Some(size) = definition.opaque else {
            continue;
        };
        let entered = ArrayDefinition {
            kind: ArrayKind::Legacy {
                rows: size.rows,
                cols: size.cols,
            },
            opaque: None,
            ..definition
        };
        evaluated.sheets[index].set_array_definition(anchor, Some(entered));
    }
    crate::rebuild_and_recalc_all(&mut evaluated, None);
    for (index, anchor, definition) in uncached {
        let rectangle: CellRange = definition.entered(anchor);
        let values: Vec<(CellRef, CellValue)> = evaluated.sheets[index]
            .cells_in_range(rectangle)
            .map(|(at, cell)| (at, cell.value.clone()))
            .collect();
        let sheet = &mut wb.sheets[index];
        for (at, value) in values {
            let stored = sheet.cell(at).cloned().unwrap_or_default();
            let anchored = (at.row, at.col) == (anchor.row, anchor.col);
            let free = matches!(stored.value, CellValue::Empty)
                && stored
                    .formula
                    .as_deref()
                    .is_none_or(|formula| formula.trim().is_empty());
            if anchored || free {
                sheet.set_cell(
                    at,
                    Cell {
                        value,
                        formula: stored.formula,
                        style: stored.style,
                    },
                );
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
    decided: HashMap<Name, bool>,
    expansions: usize,
}

/// a defined name, by its index, with the sheet its unqualified references
/// resolve against.
type Name = (usize, SheetId);

/// what a formula calls and names, gathered once.
struct Parsed {
    supported: bool,
    names: Vec<(Option<String>, String)>,
}

/// one defined name being expanded: the names its formula reads, how far
/// through them the walk is, and the earliest name still being expanded that
/// it reaches back to.
struct Frame {
    name: Name,
    order: usize,
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
            decided: HashMap::new(),
            expansions: 0,
        }
    }

    pub fn evaluable(&mut self, sheet: SheetId, formula: &str) -> bool {
        match parse_formula(formula) {
            Ok(expr) => {
                let parsed = gather(&expr);
                parsed.supported && self.names_evaluable(sheet, parsed.names)
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

    /// whether every defined name reachable from `names`, read on `sheet`, is
    /// evaluable: a walk without recursion that settles each group of names
    /// reaching one another once all of it is known, as Tarjan's algorithm
    /// does, so a later formula reading any of them costs nothing.
    fn names_evaluable(&mut self, sheet: SheetId, names: Vec<(Option<String>, String)>) -> bool {
        let mut order: HashMap<Name, usize> = HashMap::new();
        let mut pending: Vec<Name> = Vec::new();
        let mut root = Frame {
            name: (usize::MAX, sheet),
            order: usize::MAX,
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
                let from = frame.name.1;
                let Some(name) = self.resolve(from, scope, name) else {
                    continue;
                };
                match self.decided.get(&name) {
                    Some(true) => continue,
                    Some(false) => return self.refuse(&stack),
                    None => {}
                }
                if let Some(&seen) = order.get(&name) {
                    let frame = stack.last_mut().unwrap_or(&mut root);
                    frame.low = frame.low.min(seen);
                    continue;
                }
                self.expansions += 1;
                if self.expansions > MAX_EXPANSIONS {
                    return false;
                }
                let Some(parsed) = self.parsed(name.0) else {
                    self.decided.insert(name, false);
                    return self.refuse(&stack);
                };
                if !parsed.supported {
                    self.decided.insert(name, false);
                    return self.refuse(&stack);
                }
                let names = parsed.names.clone();
                let at = order.len();
                order.insert(name, at);
                pending.push(name);
                stack.push(Frame {
                    name,
                    order: at,
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
                let start = pending
                    .iter()
                    .position(|name| *name == done.name)
                    .expect("an expanded name is pending");
                for name in pending.drain(start..) {
                    self.decided.insert(name, true);
                }
            } else {
                let parent = stack.last_mut().unwrap_or(&mut root);
                parent.low = parent.low.min(done.low);
            }
        }
        for name in pending {
            self.decided.insert(name, true);
        }
        true
    }

    /// every name being expanded reaches what the engine lacks.
    fn refuse(&mut self, stack: &[Frame]) -> bool {
        for frame in stack {
            self.decided.insert(frame.name, false);
        }
        false
    }
}

/// whether `expr` calls only functions the engine evaluates, and the names
/// it reads.
fn gather(expr: &Expr) -> Parsed {
    let mut parsed = Parsed {
        supported: true,
        names: Vec::new(),
    };
    let mut pending = vec![expr];
    while let Some(expr) = pending.pop() {
        match expr {
            Expr::FuncCall { func, name, args } => {
                if func.is_none() && !crate::array::is_array_builtin(name) {
                    parsed.supported = false;
                }
                pending.extend(args);
            }
            Expr::Name { scope, name } => parsed.names.push((scope.clone(), name.clone())),
            Expr::ArrayLiteral { values, .. } => pending.extend(values),
            Expr::Unary { expr, .. } | Expr::Percent(expr) => pending.push(expr),
            Expr::Binary { lhs, rhs, .. } => pending.extend([lhs.as_ref(), rhs.as_ref()]),
            Expr::RangeJoin { start, end } => pending.extend([start.as_ref(), end.as_ref()]),
            _ => {}
        }
    }
    parsed
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
            ("SUM(", false),
        ] {
            assert_eq!(evaluable(&wb, SheetId(0), formula), expected, "{formula}");
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
            let expansions = capability.expansions;
            for (question, answer) in questions.iter().zip(answers) {
                assert_eq!(capability.evaluable(SheetId(0), question), answer);
            }
            assert_eq!(capability.expansions, expansions);
        }
    }

    /// a chain of names as long as a workbook may hold is walked without
    /// recursion, once.
    #[test]
    fn a_long_chain_of_names_is_walked_once() {
        const LENGTH: usize = 60_000;
        let names: Vec<(String, String)> = (0..LENGTH)
            .map(|index| {
                let formula = if index + 1 == LENGTH {
                    "WEBSERVICE(\"x\")".to_owned()
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
        assert!(!capability.evaluable(SheetId(0), "step_0"));
        assert!(!capability.evaluable(SheetId(0), "step_30000"));
        assert_eq!(capability.expansions, LENGTH);
    }
}
