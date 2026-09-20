//! dependency extraction: the set of cells/ranges a formula reads.

use std::collections::HashSet;

use xlsx_model::{CellRange, CellRef};

use crate::parser::{Expr, UnaryOp};
use crate::reference::offset_rect;

/// collect every reference a formula reads as `(sheet, range)` pairs;
/// `sheet` is `None` for unqualified refs. order-preserving, de-duplicated.
pub fn references(expr: &Expr) -> Vec<(Option<String>, CellRange)> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    walk(expr, &mut out, &mut seen);
    out
}

/// true when argument `index` is read for where it sits, not what it holds,
/// so it carries no data dependency: any reference under `ROW`/`COLUMN`/
/// `ROWS`/`COLUMNS`, and `OFFSET`'s anchor.
pub fn positional_argument(name: &str, index: usize, arg: &Expr) -> bool {
    let positional = match name.to_ascii_uppercase().as_str() {
        "ROW" | "COLUMN" | "ROWS" | "COLUMNS" => true,
        "OFFSET" => index == 0,
        _ => false,
    };
    positional
        && matches!(
            arg,
            Expr::Ref { .. } | Expr::Range { .. } | Expr::ColumnRange { .. } | Expr::Name { .. }
        )
}

fn walk(
    expr: &Expr,
    out: &mut Vec<(Option<String>, CellRange)>,
    seen: &mut HashSet<(Option<String>, CellRange)>,
) {
    match expr {
        Expr::Ref { sheet, cell } => {
            push_unique(
                out,
                seen,
                sheet.clone(),
                CellRange {
                    start: *cell,
                    end: *cell,
                },
            );
        }
        Expr::Range { sheet, range } => {
            push_unique(out, seen, sheet.clone(), *range);
        }
        Expr::ColumnRange { sheet, range } => {
            push_unique(out, seen, sheet.clone(), range.cell_range());
        }
        Expr::Unary { expr, .. } | Expr::Percent(expr) => walk(expr, out, seen),
        Expr::Binary { lhs, rhs, .. } => {
            walk(lhs, out, seen);
            walk(rhs, out, seen);
        }
        Expr::FuncCall { name, args } => {
            for (index, arg) in args.iter().enumerate() {
                if positional_argument(name, index, arg) {
                    continue;
                }
                walk(arg, out, seen);
            }
            if name.eq_ignore_ascii_case("OFFSET")
                && let Some(Some((sheet, range))) = offset_target(args)
            {
                push_unique(out, seen, sheet, range);
            }
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Name { .. } => {}
    }
}

/// what an OFFSET reads, without evaluating it. `None`: an argument is not a
/// literal, so the caller must treat the formula as volatile. `Some(None)`:
/// the literals resolve to #REF!.
pub(crate) fn offset_target(args: &[Expr]) -> Option<Option<(Option<String>, CellRange)>> {
    if args.len() < 3 || args.len() > 5 {
        return Some(None);
    }
    let (sheet, anchor) = anchor_range(&args[0])?;
    let rows = const_int(&args[1])?;
    let cols = const_int(&args[2])?;
    let height = match args.get(3) {
        Some(arg) => Some(const_int(arg)?),
        None => None,
    };
    let width = match args.get(4) {
        Some(arg) => Some(const_int(arg)?),
        None => None,
    };
    Some(offset_rect(anchor, rows, cols, height, width).map(|rect| (sheet, rect)))
}

fn anchor_range(expr: &Expr) -> Option<(Option<String>, CellRange)> {
    match expr {
        Expr::Ref { sheet, cell } => Some((
            sheet.clone(),
            CellRange {
                start: *cell,
                end: *cell,
            },
        )),
        Expr::Range { sheet, range } => {
            Some((sheet.clone(), CellRange::new(range.start, range.end)))
        }
        Expr::ColumnRange { sheet, range } => Some((sheet.clone(), range.cell_range())),
        _ => None,
    }
}

fn const_int(expr: &Expr) -> Option<i64> {
    match expr {
        Expr::Number(n) if n.is_finite() => Some(n.trunc() as i64),
        Expr::Unary {
            op: UnaryOp::Neg,
            expr,
        } => const_int(expr).and_then(i64::checked_neg),
        Expr::Unary {
            op: UnaryOp::Plus,
            expr,
        } => const_int(expr),
        _ => None,
    }
}

fn push_unique(
    out: &mut Vec<(Option<String>, CellRange)>,
    seen: &mut HashSet<(Option<String>, CellRange)>,
    sheet: Option<String>,
    range: CellRange,
) {
    let range = CellRange::new(
        CellRef::new(range.start.row, range.start.col),
        CellRef::new(range.end.row, range.end.col),
    );
    let entry = (sheet, range);
    if seen.insert(entry.clone()) {
        out.push(entry);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::parse_formula;

    fn args(src: &str) -> Vec<Expr> {
        match parse_formula(src).unwrap() {
            Expr::FuncCall { args, .. } => args,
            other => panic!("not a call: {other:?}"),
        }
    }

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
        assert_eq!(refs("A1 + $A1 * A$1 + $A$1"), vec!["A1"]);
    }

    #[test]
    fn no_refs_for_literals() {
        assert!(refs("1 + 2 * 3").is_empty());
    }

    #[test]
    fn positional_queries_read_nothing() {
        assert!(refs("ROW($X$1)").is_empty());
        assert!(refs("COLUMN(Sheet2!C3)").is_empty());
        assert!(refs("ROWS(A1:A9)+COLUMNS(B:D)").is_empty());
        assert!(refs("ROW()").is_empty());
    }

    #[test]
    fn positional_queries_still_read_computed_arguments() {
        assert_eq!(refs("ROW(OFFSET(A1,B1,0))"), vec!["B1"]);
    }

    #[test]
    fn walks_nested_expressions() {
        assert_eq!(refs("IF(A1>0, B1, -C1%)"), vec!["A1", "B1", "C1"]);
    }

    /// the anchor gives coordinates, never a value, so the only cells an
    /// OFFSET reads are the ones in the rectangle it resolves to.
    #[test]
    fn static_offset_reads_its_target_and_not_its_anchor() {
        assert_eq!(refs("OFFSET(A1, 1, 1)"), vec!["B2"]);
        assert_eq!(refs("SUM(OFFSET($A$1, 1, 0, 3, 2))"), vec!["A2:B4"]);
        assert_eq!(refs("OFFSET(C3, -2, -2)"), vec!["A1"]);
        assert_eq!(refs("SUM(OFFSET(A5, 0, 0, -3, 1))"), vec!["A3:A5"]);
        assert_eq!(refs("ROWS(OFFSET(E1:F4, 1, 0))"), vec!["E2:F5"]);
        assert_eq!(refs("OFFSET(Sheet2!A1, 1, 0)"), vec!["Sheet2!A2"]);
    }

    #[test]
    fn offset_without_a_target_reads_only_its_computed_arguments() {
        assert_eq!(refs("OFFSET(A1, B1, 0)"), vec!["B1"]);
        assert!(refs("OFFSET(A1, ROW(), 0)").is_empty());
        assert!(refs("OFFSET(A1, -1, 0)").is_empty());
        assert!(refs("OFFSET(A1, 0, 0, 0, 1)").is_empty());
    }

    #[test]
    fn offset_separates_dynamic_arguments_from_a_static_ref_error() {
        assert!(offset_target(&args("OFFSET(A1, B1, 0)")).is_none());
        assert!(offset_target(&args("OFFSET(Named, 1, 0)")).is_none());
        assert_eq!(offset_target(&args("OFFSET(A1, -1, 0)")), Some(None));
        assert_eq!(offset_target(&args("OFFSET(A1, 0, 0, 0, 1)")), Some(None));
        assert_eq!(offset_target(&args("OFFSET(A1, 1)")), Some(None));
        assert!(offset_target(&args("OFFSET(A1, 1, 1)")).is_some_and(|t| t.is_some()));
    }
}
