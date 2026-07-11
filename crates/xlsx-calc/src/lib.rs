//! formula engine: lexer -> parser -> `Expr` ast -> evaluator, plus dependency
//! extraction. reads cells exclusively through `xlsx_model::CellProvider`.
//!
//! ```
//! use xlsx_calc::{evaluate, parse_formula, EvalContext};
//! use xlsx_model::{Sheet, SheetId, Workbook};
//!
//! let mut wb = Workbook::default();
//! wb.sheets.push(Sheet::new("Sheet1"));
//! let expr = parse_formula("1 + 2 * 3").unwrap();
//! let ctx = EvalContext::new(&wb, SheetId(0));
//! assert_eq!(evaluate(&expr, &ctx), xlsx_model::CellValue::Number { value: 7.0 });
//! ```

pub mod deps;
pub mod engine;
pub mod eval;
pub mod functions;
pub mod graph;
pub mod lexer;
pub mod parser;
pub mod printer;

pub use deps::references;
pub use engine::{RecalcResult, rebuild_and_recalc_all, recalc_after};
pub use eval::{EvalContext, evaluate};
pub use lexer::{ParseError, TokKind, Token, lex};
pub use parser::{BinaryOp, Expr, MAX_DEPTH, UnaryOp, parse_formula};
