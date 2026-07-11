//! ast -> formula text; printing a parsed formula and re-parsing it yields an
//! equivalent ast.

use crate::parser::{BinaryOp, Expr, UnaryOp};

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
                && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
            if simple {
                format!("{name}!")
            } else {
                format!("'{}'!", name.replace('\'', "''"))
            }
        }
    }
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
            Expr::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Expr::Error(e) => e.as_str().to_string(),
            Expr::Ref { sheet, cell } => format!("{}{}", sheet_prefix(sheet), cell.to_a1()),
            Expr::Range { sheet, range } => {
                format!("{}{}", sheet_prefix(sheet), range.to_a1())
            }
            Expr::Unary { op, expr } => {
                // 6 > every binary bp: unary minus binds tighter than all binary ops
                let inner = expr.print(6);
                match op {
                    UnaryOp::Neg => format!("-{inner}"),
                    UnaryOp::Plus => format!("+{inner}"),
                }
            }
            Expr::Percent(expr) => format!("{}%", expr.print(6)),
            Expr::Binary { op, lhs, rhs } => {
                let bp = binary_bp(op);
                // left-associative: the right child needs parens at equal bp
                let s = format!("{}{}{}", lhs.print(bp), binary_token(op), rhs.print(bp + 1));
                if bp < parent_bp { format!("({s})") } else { s }
            }
            Expr::FuncCall { name, args } => {
                let inner: Vec<String> = args.iter().map(|a| a.print(0)).collect();
                format!("{}({})", name, inner.join(","))
            }
        }
    }
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
            "1<=2",
            "\"he said \"\"hi\"\"\"",
            "1-2-3",
            "2^3^2",
            "-(1+2)",
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
}
