//! information functions that look at a value's kind, a reference's sheet or
//! the workbook itself rather than at a cell's contents.

use xlsx_model::{CellValue, ErrorValue, SheetId};

use crate::array::{Value, evaluate_array};
use crate::eval::{EvalContext, as_area, boolean, err, evaluate, num, text, union_areas};
use crate::parser::Expr;

use super::omitted;

/// ERROR.TYPE(error_val): microsoft's number for an error value; anything
/// that is not an error is `#N/A`.
pub(crate) fn error_type(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match evaluate(&args[0], ctx) {
        CellValue::Error { value } => num(error_number(value)),
        _ => err(ErrorValue::NA),
    }
}

fn error_number(error: ErrorValue) -> f64 {
    match error {
        ErrorValue::Null => 1.0,
        ErrorValue::Div0 => 2.0,
        ErrorValue::Value => 3.0,
        ErrorValue::Ref => 4.0,
        ErrorValue::Name => 5.0,
        ErrorValue::Num => 6.0,
        ErrorValue::NA => 7.0,
        ErrorValue::Spill => 9.0,
        ErrorValue::Calc => 14.0,
    }
}

/// INFO(type_text): one fact about the operating environment.
///
/// "recalc" is the engine's own mode, "Automatic". "memavail", "memused" and
/// "totmem" are no longer supported by excel, which answers `#N/A`. the rest
/// describe the machine or the window excel ran in, so they are engine gaps:
/// a cell keeps the value excel cached. uncached, "numfile" counts this
/// workbook's sheets, "origin" is the unscrolled `$A:$A$1`, and the machine
/// facts are `#N/A`.
pub(crate) fn info(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let kind = match evaluate(&args[0], ctx) {
        CellValue::Text { value } => value.to_ascii_lowercase(),
        CellValue::Error { value } => return err(value),
        _ => return err(ErrorValue::Value),
    };
    match kind.as_str() {
        "recalc" => text(ctx, "Automatic"),
        "memavail" | "memused" | "totmem" => err(ErrorValue::NA),
        "numfile" => {
            ctx.record_unsupported_function();
            match sheet_count(ctx) {
                Some(count) => num(count as f64),
                None => err(ErrorValue::NA),
            }
        }
        "origin" => {
            ctx.record_unsupported_function();
            text(ctx, "$A:$A$1")
        }
        "directory" | "osversion" | "release" | "system" => {
            ctx.record_unsupported_function();
            err(ErrorValue::NA)
        }
        _ => err(ErrorValue::Value),
    }
}

/// ISFORMULA(reference): whether the top-left cell of a reference holds a
/// formula. a value that is no reference is `#VALUE!`.
pub(crate) fn isformula(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match as_area(&args[0], ctx) {
        Some(area) => boolean(ctx.provider.formula(area.sheet, area.start).is_some()),
        None => match evaluate(&args[0], ctx) {
            CellValue::Error { value } => err(value),
            _ => err(ErrorValue::Value),
        },
    }
}

/// ISNONTEXT(value): true for anything but text, a blank cell and an error
/// included.
pub(crate) fn isnontext(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    boolean(!matches!(evaluate(&args[0], ctx), CellValue::Text { .. }))
}

/// ISOMITTED(argument): whether an argument was left out. the engine invokes
/// a `LAMBDA` only with every parameter bound, so only a gap written into
/// the call itself counts as omitted.
pub(crate) fn isomitted(args: &[Expr], _ctx: &EvalContext<'_>) -> CellValue {
    match args {
        [arg] => boolean(omitted(arg)),
        _ => err(ErrorValue::Value),
    }
}

/// ISREF(value): whether the argument designates cells; a reference function
/// that fails, such as INDIRECT of bad text, does not.
pub(crate) fn isref(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    if let Some(areas) = union_areas(&args[0], ctx) {
        return boolean(areas.is_ok());
    }
    boolean(as_area(&args[0], ctx).is_some())
}

/// SHEET([value]): the 1-based position of a reference's sheet, or of the
/// sheet a name spells; the formula's own sheet when omitted. an unknown
/// sheet name is `#N/A`. excel also counts chart and dialog sheets, which
/// the workbook model does not hold.
pub(crate) fn sheet(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let arg = match args {
        [] => return position(ctx.sheet),
        [arg] if omitted(arg) => return position(ctx.sheet),
        [arg] => arg,
        _ => return err(ErrorValue::Value),
    };
    if let Some(areas) = union_areas(arg, ctx) {
        return match areas {
            Ok(areas) => areas
                .first()
                .map_or(err(ErrorValue::Ref), |area| position(area.sheet)),
            Err(error) => err(error),
        };
    }
    if let Some(area) = as_area(arg, ctx) {
        return position(area.sheet);
    }
    match evaluate(arg, ctx) {
        CellValue::Text { value } => match ctx.provider.sheet_id(&value) {
            Some(id) => position(id),
            None => err(ErrorValue::NA),
        },
        CellValue::Error { value } => err(value),
        _ => err(ErrorValue::Value),
    }
}

/// SHEETS([reference]): how many sheets a reference spans, or the workbook's
/// sheet count when omitted.
pub(crate) fn sheets(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let arg = match args {
        [] => return count(ctx),
        [arg] if omitted(arg) => return count(ctx),
        [arg] => arg,
        _ => return err(ErrorValue::Value),
    };
    if let Some(areas) = union_areas(arg, ctx) {
        return match areas {
            Ok(areas) => {
                let mut ids: Vec<u32> = areas.iter().map(|area| area.sheet.0).collect();
                ids.sort_unstable();
                ids.dedup();
                num(ids.len() as f64)
            }
            Err(error) => err(error),
        };
    }
    if as_area(arg, ctx).is_some() {
        return num(1.0);
    }
    match evaluate(arg, ctx) {
        CellValue::Error { value } => err(value),
        _ => err(ErrorValue::Value),
    }
}

fn position(sheet: SheetId) -> CellValue {
    num(f64::from(sheet.0) + 1.0)
}

fn count(ctx: &EvalContext<'_>) -> CellValue {
    match sheet_count(ctx) {
        Some(count) => num(count as f64),
        None => err(ErrorValue::NA),
    }
}

fn sheet_count(ctx: &EvalContext<'_>) -> Option<usize> {
    ctx.provider.workbook().map(|book| book.sheets.len())
}

/// TYPE(value): 1 number or blank, 2 text, 4 logical, 16 error, 64 array, 128
/// a lambda. a multi-cell reference is an array and a multi-area one reads
/// as the error it evaluates to; errors do not propagate.
pub(crate) fn type_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let arg = &args[0];
    if union_areas(arg, ctx).is_some() {
        return num(16.0);
    }
    if let Some(area) = as_area(arg, ctx) {
        if area.rows * area.cols > 1 {
            return num(64.0);
        }
        return num(kind(&evaluate(arg, ctx)));
    }
    if matches!(arg, Expr::ArrayLiteral { .. }) {
        return num(64.0);
    }
    match evaluate_array(arg, ctx) {
        Value::Lambda(_) => num(128.0),
        Value::Array(array) if array.rows() * array.cols() > 1 => num(64.0),
        value => num(kind(&value.into_scalar())),
    }
}

fn kind(value: &CellValue) -> f64 {
    match value {
        CellValue::Empty | CellValue::Number { .. } => 1.0,
        CellValue::Text { .. } => 2.0,
        CellValue::Bool { .. } => 4.0,
        CellValue::Error { .. } => 16.0,
    }
}
