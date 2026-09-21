//! ast -> formula text; printing a parsed formula and re-parsing it yields an
//! equivalent ast.

use crate::TableSpec;
use crate::parser::{BinaryOp, Expr, UnaryOp};
use xlsx_model::CellRef;

// mirrors the parser's precedence table so parens are emitted only where
// re-parsing would otherwise change the tree
fn binary_bp(op: &BinaryOp) -> u8 {
    match op {
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
            1
        }
        BinaryOp::Concat => 2,
        BinaryOp::Add | BinaryOp::Sub => 3,
        BinaryOp::Mul | BinaryOp::Div => 4,
        BinaryOp::Pow => 5,
    }
}

fn binary_token(op: &BinaryOp) -> &'static str {
    match op {
        BinaryOp::Add => "+",
        BinaryOp::Sub => "-",
        BinaryOp::Mul => "*",
        BinaryOp::Div => "/",
        BinaryOp::Pow => "^",
        BinaryOp::Concat => "&",
        BinaryOp::Eq => "=",
        BinaryOp::Ne => "<>",
        BinaryOp::Lt => "<",
        BinaryOp::Le => "<=",
        BinaryOp::Gt => ">",
        BinaryOp::Ge => ">=",
    }
}

/// quote a sheet name when it needs it (anything beyond ascii alnum + `_`).
fn sheet_prefix(sheet: &Option<String>) -> String {
    match sheet {
        None => String::new(),
        Some(name) => {
            let simple = !name.is_empty()
                && name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_')
                && !name.chars().next().is_some_and(|c| c.is_ascii_digit())
                && CellRef::parse_a1(name).is_err()
                && !is_r1c1_reference(name);
            if simple {
                format!("{name}!")
            } else {
                format!("'{}'!", name.replace('\'', "''"))
            }
        }
    }
}

fn is_r1c1_reference(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    if matches!(upper.as_str(), "R" | "C") {
        return true;
    }
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
    rest.as_bytes().get(digits) == Some(&b'C')
}

impl Expr {
    /// render the expression as formula text (without a leading `=`).
    pub fn to_formula(&self) -> String {
        self.print(0)
    }

    fn print(&self, parent_bp: u8) -> String {
        match self {
            Expr::Number(n) => n.to_string(),
            Expr::Text(t) => format!("\"{}\"", t.replace('"', "\"\"")),
            Expr::Literal(value) => literal(value),
            Expr::ArrayLiteral { cols, values } => array_literal(*cols, values),
            Expr::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Expr::Error(e) => e.as_str().to_string(),
            Expr::Ref { sheet, cell } => format!("{}{}", sheet_prefix(sheet), cell.to_a1()),
            Expr::Range { sheet, range } => {
                format!("{}{}", sheet_prefix(sheet), range.to_a1())
            }
            Expr::ColumnRange { sheet, range } => {
                format!("{}{}", sheet_prefix(sheet), range.to_a1())
            }
            Expr::RowRange { sheet, range } => {
                format!("{}{}", sheet_prefix(sheet), range.to_a1())
            }
            Expr::TableRef { table, spec } => format!("{table}{}", table_spec(spec)),
            Expr::Name { scope, name } => format!("{}{name}", sheet_prefix(scope)),
            Expr::Unary { op, expr } => {
                // 6 > every binary bp: unary minus binds tighter than all binary ops
                let inner = expr.print(6);
                match op {
                    UnaryOp::Neg => format!("-{inner}"),
                    UnaryOp::Plus => format!("+{inner}"),
                }
            }
            Expr::Percent(expr) => format!("{}%", expr.print(6)),
            // 7 > every other bp: `:` binds tighter than anything around it
            Expr::RangeJoin { start, end } => format!("{}:{}", start.print(7), end.print(7)),
            Expr::Binary { op, lhs, rhs } => {
                let bp = binary_bp(op);
                // left-associative: the right child needs parens at equal bp
                let s = format!("{}{}{}", lhs.print(bp), binary_token(op), rhs.print(bp + 1));
                if bp < parent_bp { format!("({s})") } else { s }
            }
            Expr::FuncCall { name, args, .. } => {
                let inner: Vec<String> = args.iter().map(|a| a.print(0)).collect();
                format!("{}({})", name, inner.join(","))
            }
        }
    }
}

/// print the bracket body in the always-bracketed form, which re-parses to the
/// same spec whatever the source used.
fn table_spec(spec: &TableSpec) -> String {
    let mut items: Vec<String> = spec
        .bands
        .iter()
        .map(|band| format!("[{}]", band.keyword()))
        .collect();
    let columns: Vec<String> = spec
        .first_column
        .iter()
        .chain(spec.last_column.iter())
        .map(|name| format!("[{}]", escape_table_name(name)))
        .collect();
    let span = columns.join(":");
    if !span.is_empty() {
        items.push(span);
    }
    format!("[{}]", items.join(","))
}

fn escape_table_name(name: &str) -> String {
    let mut out = String::with_capacity(name.len());
    for c in name.chars() {
        if matches!(c, '\'' | '[' | ']' | '#' | '@') {
            out.push('\'');
        }
        out.push(c);
    }
    out
}

/// print a spliced value the way the same literal would be written.
fn literal(value: &xlsx_model::CellValue) -> String {
    match value {
        xlsx_model::CellValue::Empty => "\"\"".to_string(),
        xlsx_model::CellValue::Number { value } => crate::eval::format_number(*value),
        xlsx_model::CellValue::Text { value } => format!("\"{}\"", value.replace('"', "\"\"")),
        xlsx_model::CellValue::Bool { value } => if *value { "TRUE" } else { "FALSE" }.to_string(),
        xlsx_model::CellValue::Error { value } => value.as_str().to_string(),
    }
}

fn array_literal(cols: usize, values: &[Expr]) -> String {
    let rows: Vec<String> = values
        .chunks(cols.max(1))
        .map(|row| {
            row.iter()
                .map(|value| value.print(0))
                .collect::<Vec<_>>()
                .join(",")
        })
        .collect();
    format!("{{{}}}", rows.join(";"))
}

#[cfg(test)]
mod tests {
    use crate::parse_formula;

    #[track_caller]
    fn round_trips(src: &str) {
        let ast = parse_formula(src).unwrap();
        let printed = ast.to_formula();
        let reparsed = parse_formula(&printed).unwrap();
        assert_eq!(ast, reparsed, "printed form {printed:?} changed the ast");
    }

    #[test]
    fn round_trips_representative_formulas() {
        for src in [
            "1+2*3",
            "(1+2)*3",
            "-2^2",
            "2^-3",
            "A1+$B$2",
            "SUM(A1:B10,3,\"x\")",
            "IF(A1>=3,\"yes\",\"no\")",
            "Sheet1!A1&'My Sheet'!B2",
            "10%",
            "(1+2)%",
            "NOT(TRUE)",
            "TaxRate*'Input Sheet'!LocalRate",
            "1<=2",
            "\"he said \"\"hi\"\"\"",
            "1-2-3",
            "2^3^2",
            "-(1+2)",
            "A1:INDEX(A1:A9,3)",
            "MIN(AA9:INDEX(AA9:AH9,MATCH(4,AI9:AP9,0)))",
            "SUM($B$4:OFFSET($B$4,0,2))",
            "INDEX(A:A,1):INDEX(A:A,4)",
            "-A1:B2",
            "Sheet1!A1:INDEX(Sheet1!A:A,2)",
            "A1:B2:C3",
        ] {
            round_trips(src);
        }
    }

    #[test]
    fn emits_minimal_parens() {
        let ast = parse_formula("(1+2)*3").unwrap();
        assert_eq!(ast.to_formula(), "(1+2)*3");
        let ast = parse_formula("1+(2*3)").unwrap();
        assert_eq!(ast.to_formula(), "1+2*3");
    }

    #[test]
    fn quotes_cell_like_sheet_names() {
        assert_eq!(parse_formula("'A1'!B2").unwrap().to_formula(), "'A1'!B2");
        assert_eq!(
            parse_formula("'R1C1'!B2").unwrap().to_formula(),
            "'R1C1'!B2"
        );
        assert_eq!(parse_formula("'R'!B2").unwrap().to_formula(), "'R'!B2");
        assert_eq!(parse_formula("'C'!B2").unwrap().to_formula(), "'C'!B2");
        assert_eq!(parse_formula("'R4C'!B2").unwrap().to_formula(), "'R4C'!B2");
        assert_eq!(
            parse_formula("'R1C1b'!B2").unwrap().to_formula(),
            "'R1C1b'!B2"
        );
    }
}
