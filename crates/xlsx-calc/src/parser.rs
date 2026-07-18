//! precedence-climbing (pratt) parser: tokens -> `Expr` ast. never panics;
//! malformed input yields a positioned `ParseError`, nesting is depth-capped.

use xlsx_model::{CellRange, CellRef, ErrorValue};

use crate::lexer::{ParseError, TokKind, Token, lex};

/// maximum expression nesting depth before we bail with a `ParseError`.
pub const MAX_DEPTH: usize = 100;

/// unary prefix binding power; higher than `^`'s left power (9) because excel
/// binds unary minus tighter than exponent: `-2^2 = 4`.
const UNARY_BP: u8 = 10;

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrorValue),
    Ref {
        sheet: Option<String>,
        cell: CellRef,
    },
    Range {
        sheet: Option<String>,
        range: CellRange,
    },
    Unary {
        op: UnaryOp,
        expr: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
    },
    Percent(Box<Expr>),
    FuncCall {
        name: String,
        args: Vec<Expr>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOp {
    Neg,
    Plus,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

/// public entry point: parse a formula body (no leading `=`) into an ast.
pub fn parse_formula(src: &str) -> Result<Expr, ParseError> {
    let tokens = lex(src)?;
    if tokens.is_empty() {
        return Err(ParseError::new(0, "empty formula"));
    }
    let mut p = Parser {
        tokens: &tokens,
        pos: 0,
        src_len: src.len(),
    };
    let expr = p.expr_bp(0, 0)?;
    if let Some(tok) = p.peek() {
        return Err(ParseError::new(tok.start, "unexpected trailing token"));
    }
    Ok(expr.expr)
}

struct ParsedExpr {
    expr: Expr,
    depth: usize,
}

impl ParsedExpr {
    fn leaf(expr: Expr) -> Self {
        Self { expr, depth: 1 }
    }
}

struct Parser<'a> {
    tokens: &'a [Token],
    pos: usize,
    src_len: usize,
}

impl Parser<'_> {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn advance(&mut self) -> Option<&Token> {
        let t = self.tokens.get(self.pos);
        if t.is_some() {
            self.pos += 1;
        }
        t
    }

    /// byte offset for error reporting at the current position (or end).
    fn here(&self) -> usize {
        self.peek().map(|t| t.start).unwrap_or(self.src_len)
    }

    fn expect(&mut self, want: &TokKind, what: &str) -> Result<(), ParseError> {
        match self.peek() {
            Some(t) if &t.kind == want => {
                self.advance();
                Ok(())
            }
            _ => Err(ParseError::new(self.here(), format!("expected {what}"))),
        }
    }

    fn expr_bp(&mut self, min_bp: u8, depth: usize) -> Result<ParsedExpr, ParseError> {
        if depth > MAX_DEPTH {
            return Err(ParseError::new(self.here(), "formula nesting too deep"));
        }
        let mut lhs = self.prefix(depth)?;
        while let Some((op, l_bp, r_bp)) = self.peek().and_then(|t| infix_binding_power(&t.kind)) {
            if l_bp < min_bp {
                break;
            }
            self.advance();
            let rhs = self.expr_bp(r_bp, depth + 1)?;
            let ast_depth = lhs.depth.max(rhs.depth) + 1;
            self.validate_ast_depth(ast_depth)?;
            lhs = ParsedExpr {
                expr: Expr::Binary {
                    op,
                    lhs: Box::new(lhs.expr),
                    rhs: Box::new(rhs.expr),
                },
                depth: ast_depth,
            };
        }
        Ok(lhs)
    }

    fn prefix(&mut self, depth: usize) -> Result<ParsedExpr, ParseError> {
        match self.peek().map(|t| &t.kind) {
            Some(TokKind::Minus) => {
                self.advance();
                let expr = self.expr_bp(UNARY_BP, depth + 1)?;
                let ast_depth = expr.depth + 1;
                self.validate_ast_depth(ast_depth)?;
                Ok(ParsedExpr {
                    expr: Expr::Unary {
                        op: UnaryOp::Neg,
                        expr: Box::new(expr.expr),
                    },
                    depth: ast_depth,
                })
            }
            Some(TokKind::Plus) => {
                self.advance();
                let expr = self.expr_bp(UNARY_BP, depth + 1)?;
                let ast_depth = expr.depth + 1;
                self.validate_ast_depth(ast_depth)?;
                Ok(ParsedExpr {
                    expr: Expr::Unary {
                        op: UnaryOp::Plus,
                        expr: Box::new(expr.expr),
                    },
                    depth: ast_depth,
                })
            }
            _ => self.postfix(depth),
        }
    }

    /// an atom followed by any number of `%` postfixes.
    fn postfix(&mut self, depth: usize) -> Result<ParsedExpr, ParseError> {
        let mut expr = self.atom(depth)?;
        while matches!(self.peek().map(|t| &t.kind), Some(TokKind::Percent)) {
            self.advance();
            let ast_depth = expr.depth + 1;
            self.validate_ast_depth(ast_depth)?;
            expr = ParsedExpr {
                expr: Expr::Percent(Box::new(expr.expr)),
                depth: ast_depth,
            };
        }
        Ok(expr)
    }

    fn atom(&mut self, depth: usize) -> Result<ParsedExpr, ParseError> {
        let pos = self.here();
        let Some(tok) = self.advance() else {
            return Err(ParseError::new(pos, "unexpected end of formula"));
        };
        match tok.kind.clone() {
            TokKind::Num(n) => Ok(ParsedExpr::leaf(Expr::Number(n))),
            TokKind::Str(s) => Ok(ParsedExpr::leaf(Expr::Text(s))),
            TokKind::Bool(b) => Ok(ParsedExpr::leaf(Expr::Bool(b))),
            TokKind::ErrLit(e) => Ok(ParsedExpr::leaf(Expr::Error(e))),
            TokKind::Ref { sheet, cell } => Ok(ParsedExpr::leaf(Expr::Ref { sheet, cell })),
            TokKind::Range { sheet, range } => Ok(ParsedExpr::leaf(Expr::Range { sheet, range })),
            TokKind::LParen => {
                let inner = self.expr_bp(0, depth + 1)?;
                self.expect(&TokKind::RParen, "')'")?;
                Ok(inner)
            }
            TokKind::Ident(name) => self.func_call(name, depth),
            other => Err(ParseError::new(pos, format!("unexpected token {other:?}"))),
        }
    }

    /// a bare name must be a function call; defined names are not supported.
    fn func_call(&mut self, name: String, depth: usize) -> Result<ParsedExpr, ParseError> {
        self.expect(&TokKind::LParen, "'(' after function name")?;
        let mut args = Vec::new();
        if matches!(self.peek().map(|t| &t.kind), Some(TokKind::RParen)) {
            self.advance();
            return Ok(ParsedExpr::leaf(Expr::FuncCall {
                name,
                args: Vec::new(),
            }));
        }
        loop {
            args.push(self.expr_bp(0, depth + 1)?);
            match self.peek().map(|t| &t.kind) {
                Some(TokKind::Comma) => {
                    self.advance();
                }
                Some(TokKind::RParen) => {
                    self.advance();
                    break;
                }
                _ => return Err(ParseError::new(self.here(), "expected ',' or ')'")),
            }
        }
        let ast_depth = args.iter().map(|arg| arg.depth).max().unwrap_or(0) + 1;
        self.validate_ast_depth(ast_depth)?;
        Ok(ParsedExpr {
            expr: Expr::FuncCall {
                name,
                args: args.into_iter().map(|arg| arg.expr).collect(),
            },
            depth: ast_depth,
        })
    }

    fn validate_ast_depth(&self, depth: usize) -> Result<(), ParseError> {
        if depth > MAX_DEPTH {
            Err(ParseError::new(self.here(), "formula nesting too deep"))
        } else {
            Ok(())
        }
    }
}

/// left/right binding powers for infix operators; `None` for non-operators.
/// left-assoc operators use `l < r`; equal precedence stops the climb.
fn infix_binding_power(kind: &TokKind) -> Option<(BinaryOp, u8, u8)> {
    Some(match kind {
        TokKind::Eq => (BinaryOp::Eq, 1, 2),
        TokKind::Ne => (BinaryOp::Ne, 1, 2),
        TokKind::Lt => (BinaryOp::Lt, 1, 2),
        TokKind::Le => (BinaryOp::Le, 1, 2),
        TokKind::Gt => (BinaryOp::Gt, 1, 2),
        TokKind::Ge => (BinaryOp::Ge, 1, 2),
        TokKind::Amp => (BinaryOp::Concat, 3, 4),
        TokKind::Plus => (BinaryOp::Add, 5, 6),
        TokKind::Minus => (BinaryOp::Sub, 5, 6),
        TokKind::Star => (BinaryOp::Mul, 7, 8),
        TokKind::Slash => (BinaryOp::Div, 7, 8),
        TokKind::Caret => (BinaryOp::Pow, 9, 10),
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(src: &str) -> Expr {
        parse_formula(src).unwrap()
    }

    fn bin(op: BinaryOp, l: Expr, r: Expr) -> Expr {
        Expr::Binary {
            op,
            lhs: Box::new(l),
            rhs: Box::new(r),
        }
    }

    #[test]
    fn precedence_mul_over_add() {
        assert_eq!(
            parse("1+2*3"),
            bin(
                BinaryOp::Add,
                Expr::Number(1.0),
                bin(BinaryOp::Mul, Expr::Number(2.0), Expr::Number(3.0))
            )
        );
    }

    #[test]
    fn excel_unary_minus_binds_tighter_than_power() {
        assert_eq!(
            parse("-2^2"),
            bin(
                BinaryOp::Pow,
                Expr::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(Expr::Number(2.0))
                },
                Expr::Number(2.0)
            )
        );
    }

    #[test]
    fn power_is_left_associative() {
        assert_eq!(
            parse("2^3^2"),
            bin(
                BinaryOp::Pow,
                bin(BinaryOp::Pow, Expr::Number(2.0), Expr::Number(3.0)),
                Expr::Number(2.0)
            )
        );
    }

    #[test]
    fn concat_looser_than_arithmetic() {
        assert_eq!(
            parse("1+2&3"),
            bin(
                BinaryOp::Concat,
                bin(BinaryOp::Add, Expr::Number(1.0), Expr::Number(2.0)),
                Expr::Number(3.0)
            )
        );
    }

    #[test]
    fn comparison_loosest() {
        assert_eq!(
            parse("1+1=2"),
            bin(
                BinaryOp::Eq,
                bin(BinaryOp::Add, Expr::Number(1.0), Expr::Number(1.0)),
                Expr::Number(2.0)
            )
        );
    }

    #[test]
    fn percent_binds_tightest() {
        assert_eq!(
            parse("-50%"),
            Expr::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(Expr::Percent(Box::new(Expr::Number(50.0))))
            }
        );
    }

    #[test]
    fn parens_override_precedence() {
        assert_eq!(
            parse("(1+2)*3"),
            bin(
                BinaryOp::Mul,
                bin(BinaryOp::Add, Expr::Number(1.0), Expr::Number(2.0)),
                Expr::Number(3.0)
            )
        );
    }

    #[test]
    fn function_calls_parse_args() {
        match parse("SUM(1, A1, 2+3)") {
            Expr::FuncCall { name, args } => {
                assert_eq!(name, "SUM");
                assert_eq!(args.len(), 3);
            }
            other => panic!("expected func call, got {other:?}"),
        }
        assert!(matches!(parse("PI()"), Expr::FuncCall { args, .. } if args.is_empty()));
    }

    #[test]
    fn rejects_malformed_input() {
        for src in ["", "1+", "(1", "1 2", "SUM(1,)", "FOO", "*1", ")"] {
            assert!(parse_formula(src).is_err(), "should reject {src:?}");
        }
    }

    #[test]
    fn rejects_excessive_nesting() {
        let src = format!("{}1{}", "(".repeat(200), ")".repeat(200));
        let err = parse_formula(&src).unwrap_err();
        assert!(err.message.contains("too deep"));
    }

    #[test]
    fn rejects_deep_postfix_chain() {
        let src = format!("1{}", "%".repeat(MAX_DEPTH + 1));
        assert!(
            parse_formula(&src)
                .unwrap_err()
                .message
                .contains("too deep")
        );
    }

    #[test]
    fn rejects_deep_left_associative_chain() {
        let src = std::iter::repeat_n("1", MAX_DEPTH + 1)
            .collect::<Vec<_>>()
            .join("+");
        assert!(
            parse_formula(&src)
                .unwrap_err()
                .message
                .contains("too deep")
        );
    }
}
