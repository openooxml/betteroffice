//! whether the engine can evaluate a formula wherever it could run, decided
//! from the formula alone so that no input changes the answer.

use std::collections::HashSet;

use xlsx_model::{CellProvider, SheetId, Workbook};

use crate::parser::{Expr, parse_formula};

/// whether the engine implements everything `formula`, on `sheet`, could
/// run: it parses, and every function it calls, in every branch and in every
/// defined name it reaches, is one the engine evaluates. an `IF` or `IFERROR`
/// around a function the engine lacks answers no whichever way it would go.
pub fn evaluable(wb: &Workbook, sheet: SheetId, formula: &str) -> bool {
    parse_formula(formula).is_ok_and(|expr| supported(&expr, wb, sheet, &mut HashSet::new()))
}

fn supported(
    expr: &Expr,
    wb: &Workbook,
    sheet: SheetId,
    expanded: &mut HashSet<(SheetId, String)>,
) -> bool {
    match expr {
        Expr::FuncCall { func, name, args } => {
            (func.is_some() || crate::array::is_array_builtin(name))
                && args.iter().all(|arg| supported(arg, wb, sheet, expanded))
        }
        Expr::Name { scope, name } => {
            let lookup = match scope {
                Some(scope) => match wb.sheet_id(scope) {
                    Some(lookup) => lookup,
                    None => return true,
                },
                None => sheet,
            };
            let Some(defined) = wb.defined_name(lookup, name) else {
                return true;
            };
            if !expanded.insert((lookup, name.to_lowercase())) {
                return true;
            }
            let formula = defined
                .formula
                .strip_prefix('=')
                .unwrap_or(&defined.formula);
            let scope = defined.local_sheet.unwrap_or(lookup);
            parse_formula(formula).is_ok_and(|expr| supported(&expr, wb, scope, expanded))
        }
        Expr::ArrayLiteral { values, .. } => values
            .iter()
            .all(|value| supported(value, wb, sheet, expanded)),
        Expr::Unary { expr, .. } | Expr::Percent(expr) => supported(expr, wb, sheet, expanded),
        Expr::Binary { lhs, rhs, .. } => {
            supported(lhs, wb, sheet, expanded) && supported(rhs, wb, sheet, expanded)
        }
        Expr::RangeJoin { start, end } => {
            supported(start, wb, sheet, expanded) && supported(end, wb, sheet, expanded)
        }
        _ => true,
    }
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
            ("LET(x,2,x*_xlfn.SEQUENCE(2))", true),
            ("SUM(", false),
        ] {
            assert_eq!(evaluable(&wb, SheetId(0), formula), expected, "{formula}");
        }
    }
}
