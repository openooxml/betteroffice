//! Formula AST.

use crate::Unit;

#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    Number(f64, Unit),
    String(String),
    Reference(String),
    Unary(Box<Expr>),
    Binary(Box<Expr>, Op, Box<Expr>),
    Call(String, Vec<Expr>),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Op {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}
