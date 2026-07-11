//! dependency extraction: the set of cells/ranges a formula reads.

use xlsx_model::CellRange;

use crate::parser::Expr;

/// collect every reference a formula reads as `(sheet, range)` pairs;
/// `sheet` is `None` for unqualified refs. order-preserving, de-duplicated.
pub fn references(expr: &Expr) -> Vec<(Option<String>, CellRange)> {
    let mut out = Vec::new();
    walk(expr, &mut out);
    out
}

fn walk(expr: &Expr, out: &mut Vec<(Option<String>, CellRange)>) {
    match expr {
        Expr::Ref { sheet, cell } => {
            push_unique(
                out,
                sheet.clone(),
                CellRange {
                    start: *cell,
                    end: *cell,
                },
            );
        }
        Expr::Range { sheet, range } => {
            push_unique(out, sheet.clone(), *range);
        }
        Expr::Unary { expr, .. } | Expr::Percent(expr) => walk(expr, out),
        Expr::Binary { lhs, rhs, .. } => {
            walk(lhs, out);
            walk(rhs, out);
        }
        Expr::FuncCall { args, .. } => {
            for arg in args {
                walk(arg, out);
            }
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) => {}
    }
}

fn push_unique(
    out: &mut Vec<(Option<String>, CellRange)>,
    sheet: Option<String>,
    range: CellRange,
) {
    let entry = (sheet, range);
    if !out.contains(&entry) {
        out.push(entry);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::parse_formula;

    fn refs(src: &str) -> Vec<String> {
        references(&parse_formula(src).unwrap())
            .into_iter()
            .map(|(sheet, range)| match sheet {
                Some(s) => format!("{s}!{}", range.to_a1()),
                None => range.to_a1(),
            })
            .collect()
    }

    #[test]
    fn extracts_single_cell() {
        assert_eq!(refs("A1+1"), vec!["A1"]);
    }

    #[test]
    fn extracts_range_and_qualified() {
        assert_eq!(refs("SUM(A1:B2) + Sheet2!C3"), vec!["A1:B2", "Sheet2!C3"]);
    }

    #[test]
    fn dedups_repeated_refs() {
        assert_eq!(refs("A1 + A1 * A1"), vec!["A1"]);
    }

    #[test]
    fn no_refs_for_literals() {
        assert!(refs("1 + 2 * 3").is_empty());
    }

    #[test]
    fn walks_nested_expressions() {
        assert_eq!(refs("IF(A1>0, B1, -C1%)"), vec!["A1", "B1", "C1"]);
    }
}
