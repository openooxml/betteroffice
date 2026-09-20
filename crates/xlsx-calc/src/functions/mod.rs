//! builtin function library: `resolve` interns a case-insensitive name to a
//! `Func` bound at parse time; `Func::call` dispatches. builtins receive
//! arguments unevaluated so control-flow can skip branches.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, as_area, err, evaluate, num, parse_num, to_number};
use crate::parser::Expr;

pub mod criteria;
pub mod datetime;
pub mod info;
pub mod logical;
pub mod lookups;
pub mod math;
pub mod stats;
pub mod text;

/// a builtin's interned identity, bound once at parse time; `call` dispatches
/// on it during evaluation so the hot loop never touches the name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Func {
    Sum,
    SumIf,
    SumIfs,
    SumProduct,
    Product,
    Abs,
    Sign,
    Round,
    RoundUp,
    RoundDown,
    Mround,
    Ceiling,
    Floor,
    Int,
    Trunc,
    Mod,
    Power,
    Sqrt,
    Exp,
    Ln,
    Log,
    Log10,
    Pi,
    Average,
    Count,
    CountA,
    CountBlank,
    CountIf,
    CountIfs,
    AverageIf,
    AverageIfs,
    Min,
    Max,
    Median,
    Mode,
    StdevS,
    StdevP,
    VarS,
    VarP,
    Large,
    Small,
    Rank,
    Len,
    Left,
    Right,
    Mid,
    Find,
    Search,
    Substitute,
    Replace,
    Trim,
    Upper,
    Lower,
    Proper,
    Clean,
    Rept,
    Exact,
    T,
    Char,
    Code,
    Value,
    NumberValue,
    Text,
    TextJoin,
    Concat,
    Date,
    Year,
    Month,
    Day,
    Weekday,
    Edate,
    Eomonth,
    Today,
    Now,
    Hour,
    Minute,
    Second,
    Time,
    DateDif,
    If,
    IfError,
    IfNa,
    Ifs,
    Switch,
    And,
    Or,
    Not,
    Xor,
    Vlookup,
    Hlookup,
    Index,
    Match,
    Xlookup,
    Choose,
    Row,
    Column,
    Rows,
    Columns,
    IsBlank,
    IsNumber,
    IsText,
    IsLogical,
    IsError,
    IsErr,
    IsNa,
    Na,
    N,
}

/// resolve a function name (case-insensitive) to its interned id; aliases map
/// to the same id. every builtin name is at most 11 ascii bytes, so the
/// uppercase fold fits in a stack buffer and never allocates.
pub fn resolve(name: &str) -> Option<Func> {
    const MAX_BUILTIN_LEN: usize = 16;
    let bytes = name.as_bytes();
    if bytes.len() > MAX_BUILTIN_LEN {
        return None;
    }
    let mut buf = [0u8; MAX_BUILTIN_LEN];
    for (i, b) in bytes.iter().enumerate() {
        buf[i] = b.to_ascii_uppercase();
    }
    let upper = std::str::from_utf8(&buf[..bytes.len()]).ok()?;
    Some(match upper {
        "SUM" => Func::Sum,
        "SUMIF" => Func::SumIf,
        "SUMIFS" => Func::SumIfs,
        "SUMPRODUCT" => Func::SumProduct,
        "PRODUCT" => Func::Product,
        "ABS" => Func::Abs,
        "SIGN" => Func::Sign,
        "ROUND" => Func::Round,
        "ROUNDUP" => Func::RoundUp,
        "ROUNDDOWN" => Func::RoundDown,
        "MROUND" => Func::Mround,
        "CEILING" => Func::Ceiling,
        "FLOOR" => Func::Floor,
        "INT" => Func::Int,
        "TRUNC" => Func::Trunc,
        "MOD" => Func::Mod,
        "POWER" => Func::Power,
        "SQRT" => Func::Sqrt,
        "EXP" => Func::Exp,
        "LN" => Func::Ln,
        "LOG" => Func::Log,
        "LOG10" => Func::Log10,
        "PI" => Func::Pi,
        "AVERAGE" => Func::Average,
        "COUNT" => Func::Count,
        "COUNTA" => Func::CountA,
        "COUNTBLANK" => Func::CountBlank,
        "COUNTIF" => Func::CountIf,
        "COUNTIFS" => Func::CountIfs,
        "AVERAGEIF" => Func::AverageIf,
        "AVERAGEIFS" => Func::AverageIfs,
        "MIN" => Func::Min,
        "MAX" => Func::Max,
        "MEDIAN" => Func::Median,
        "MODE" | "MODE.SNGL" => Func::Mode,
        "STDEV" | "STDEV.S" => Func::StdevS,
        "STDEVP" | "STDEV.P" => Func::StdevP,
        "VAR" | "VAR.S" => Func::VarS,
        "VARP" | "VAR.P" => Func::VarP,
        "LARGE" => Func::Large,
        "SMALL" => Func::Small,
        "RANK" | "RANK.EQ" => Func::Rank,
        "LEN" => Func::Len,
        "LEFT" => Func::Left,
        "RIGHT" => Func::Right,
        "MID" => Func::Mid,
        "FIND" => Func::Find,
        "SEARCH" => Func::Search,
        "SUBSTITUTE" => Func::Substitute,
        "REPLACE" => Func::Replace,
        "TRIM" => Func::Trim,
        "UPPER" => Func::Upper,
        "LOWER" => Func::Lower,
        "PROPER" => Func::Proper,
        "CLEAN" => Func::Clean,
        "REPT" => Func::Rept,
        "EXACT" => Func::Exact,
        "T" => Func::T,
        "CHAR" => Func::Char,
        "CODE" => Func::Code,
        "VALUE" => Func::Value,
        "NUMBERVALUE" => Func::NumberValue,
        "TEXT" => Func::Text,
        "TEXTJOIN" => Func::TextJoin,
        "CONCATENATE" | "CONCAT" => Func::Concat,
        "DATE" => Func::Date,
        "YEAR" => Func::Year,
        "MONTH" => Func::Month,
        "DAY" => Func::Day,
        "WEEKDAY" => Func::Weekday,
        "EDATE" => Func::Edate,
        "EOMONTH" => Func::Eomonth,
        "TODAY" => Func::Today,
        "NOW" => Func::Now,
        "HOUR" => Func::Hour,
        "MINUTE" => Func::Minute,
        "SECOND" => Func::Second,
        "TIME" => Func::Time,
        "DATEDIF" => Func::DateDif,
        "IF" => Func::If,
        "IFERROR" => Func::IfError,
        "IFNA" => Func::IfNa,
        "IFS" => Func::Ifs,
        "SWITCH" => Func::Switch,
        "AND" => Func::And,
        "OR" => Func::Or,
        "NOT" => Func::Not,
        "XOR" => Func::Xor,
        "VLOOKUP" => Func::Vlookup,
        "HLOOKUP" => Func::Hlookup,
        "INDEX" => Func::Index,
        "MATCH" => Func::Match,
        "XLOOKUP" => Func::Xlookup,
        "CHOOSE" => Func::Choose,
        "ROW" => Func::Row,
        "COLUMN" => Func::Column,
        "ROWS" => Func::Rows,
        "COLUMNS" => Func::Columns,
        "ISBLANK" => Func::IsBlank,
        "ISNUMBER" => Func::IsNumber,
        "ISTEXT" => Func::IsText,
        "ISLOGICAL" => Func::IsLogical,
        "ISERROR" => Func::IsError,
        "ISERR" => Func::IsErr,
        "ISNA" => Func::IsNa,
        "NA" => Func::Na,
        "N" => Func::N,
        _ => return None,
    })
}

impl Func {
    /// invoke the implementation with unevaluated arguments.
    pub fn call(self, args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
        match self {
            Func::Sum => math::sum(args, ctx),
            Func::SumIf => math::sumif(args, ctx),
            Func::SumIfs => math::sumifs(args, ctx),
            Func::SumProduct => math::sumproduct(args, ctx),
            Func::Product => math::product(args, ctx),
            Func::Abs => math::abs(args, ctx),
            Func::Sign => math::sign(args, ctx),
            Func::Round => math::round(args, ctx),
            Func::RoundUp => math::roundup(args, ctx),
            Func::RoundDown => math::rounddown(args, ctx),
            Func::Mround => math::mround(args, ctx),
            Func::Ceiling => math::ceiling(args, ctx),
            Func::Floor => math::floor(args, ctx),
            Func::Int => math::int(args, ctx),
            Func::Trunc => math::trunc(args, ctx),
            Func::Mod => math::mod_(args, ctx),
            Func::Power => math::power(args, ctx),
            Func::Sqrt => math::sqrt(args, ctx),
            Func::Exp => math::exp(args, ctx),
            Func::Ln => math::ln(args, ctx),
            Func::Log => math::log(args, ctx),
            Func::Log10 => math::log10(args, ctx),
            Func::Pi => math::pi(args, ctx),
            Func::Average => stats::average(args, ctx),
            Func::Count => stats::count(args, ctx),
            Func::CountA => stats::counta(args, ctx),
            Func::CountBlank => stats::countblank(args, ctx),
            Func::CountIf => stats::countif(args, ctx),
            Func::CountIfs => stats::countifs(args, ctx),
            Func::AverageIf => stats::averageif(args, ctx),
            Func::AverageIfs => stats::averageifs(args, ctx),
            Func::Min => stats::min(args, ctx),
            Func::Max => stats::max(args, ctx),
            Func::Median => stats::median(args, ctx),
            Func::Mode => stats::mode(args, ctx),
            Func::StdevS => stats::stdev_s(args, ctx),
            Func::StdevP => stats::stdev_p(args, ctx),
            Func::VarS => stats::var_s(args, ctx),
            Func::VarP => stats::var_p(args, ctx),
            Func::Large => stats::large(args, ctx),
            Func::Small => stats::small(args, ctx),
            Func::Rank => stats::rank(args, ctx),
            Func::Len => text::len(args, ctx),
            Func::Left => text::left(args, ctx),
            Func::Right => text::right(args, ctx),
            Func::Mid => text::mid(args, ctx),
            Func::Find => text::find(args, ctx),
            Func::Search => text::search(args, ctx),
            Func::Substitute => text::substitute(args, ctx),
            Func::Replace => text::replace(args, ctx),
            Func::Trim => text::trim(args, ctx),
            Func::Upper => text::upper(args, ctx),
            Func::Lower => text::lower(args, ctx),
            Func::Proper => text::proper(args, ctx),
            Func::Clean => text::clean(args, ctx),
            Func::Rept => text::rept(args, ctx),
            Func::Exact => text::exact(args, ctx),
            Func::T => text::t(args, ctx),
            Func::Char => text::char_(args, ctx),
            Func::Code => text::code(args, ctx),
            Func::Value => text::value(args, ctx),
            Func::NumberValue => text::numbervalue(args, ctx),
            Func::Text => text::text_fn(args, ctx),
            Func::TextJoin => text::textjoin(args, ctx),
            Func::Concat => text::concat(args, ctx),
            Func::Date => datetime::date(args, ctx),
            Func::Year => datetime::year(args, ctx),
            Func::Month => datetime::month(args, ctx),
            Func::Day => datetime::day(args, ctx),
            Func::Weekday => datetime::weekday(args, ctx),
            Func::Edate => datetime::edate(args, ctx),
            Func::Eomonth => datetime::eomonth(args, ctx),
            Func::Today => datetime::today(args, ctx),
            Func::Now => datetime::now(args, ctx),
            Func::Hour => datetime::hour(args, ctx),
            Func::Minute => datetime::minute(args, ctx),
            Func::Second => datetime::second(args, ctx),
            Func::Time => datetime::time(args, ctx),
            Func::DateDif => datetime::datedif(args, ctx),
            Func::If => logical::if_(args, ctx),
            Func::IfError => logical::iferror(args, ctx),
            Func::IfNa => logical::ifna(args, ctx),
            Func::Ifs => logical::ifs(args, ctx),
            Func::Switch => logical::switch(args, ctx),
            Func::And => logical::and(args, ctx),
            Func::Or => logical::or(args, ctx),
            Func::Not => logical::not(args, ctx),
            Func::Xor => logical::xor(args, ctx),
            Func::Vlookup => lookups::vlookup(args, ctx),
            Func::Hlookup => lookups::hlookup(args, ctx),
            Func::Index => lookups::index(args, ctx),
            Func::Match => lookups::match_(args, ctx),
            Func::Xlookup => lookups::xlookup(args, ctx),
            Func::Choose => lookups::choose(args, ctx),
            Func::Row => lookups::row(args, ctx),
            Func::Column => lookups::column(args, ctx),
            Func::Rows => lookups::rows(args, ctx),
            Func::Columns => lookups::columns(args, ctx),
            Func::IsBlank => info::isblank(args, ctx),
            Func::IsNumber => info::isnumber(args, ctx),
            Func::IsText => info::istext(args, ctx),
            Func::IsLogical => info::islogical(args, ctx),
            Func::IsError => info::iserror(args, ctx),
            Func::IsErr => info::iserr(args, ctx),
            Func::IsNa => info::isna(args, ctx),
            Func::Na => info::na(args, ctx),
            Func::N => info::n(args, ctx),
        }
    }
}

/// collect numbers for aggregation: referenced cells contribute only numeric
/// values, literal/computed arguments coerce, errors propagate.
pub(crate) fn collect_numbers(
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = Vec::new();
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                for value in area.values(ctx)? {
                    push_reference_number(&mut nums, value)?;
                }
            }
            None => match evaluate(arg, ctx) {
                CellValue::Number { value } => nums.push(value),
                CellValue::Bool { value } => nums.push(if value { 1.0 } else { 0.0 }),
                CellValue::Empty => {}
                CellValue::Text { value } => match parse_num(&value) {
                    Some(n) => nums.push(n),
                    None => return Err(ErrorValue::Value),
                },
                CellValue::Error { value } => return Err(value),
            },
        }
    }
    Ok(nums)
}

/// a referenced cell contributes to aggregation only when numeric; errors
/// propagate, text/bool/blank are silently skipped.
fn push_reference_number(nums: &mut Vec<f64>, v: CellValue) -> Result<(), ErrorValue> {
    match v {
        CellValue::Number { value } => nums.push(value),
        CellValue::Error { value } => return Err(value),
        _ => {}
    }
    Ok(())
}

/// evaluate one argument and coerce it to a number, propagating errors.
pub(crate) fn nth_number(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    i: usize,
) -> Result<f64, ErrorValue> {
    to_number(&evaluate(&args[i], ctx))
}

/// evaluate one argument, coerce to a number, truncate toward zero.
pub(crate) fn nth_int(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<i64, ErrorValue> {
    Ok(nth_number(args, ctx, i)?.trunc() as i64)
}

/// finalize a computed float: non-finite results become `#NUM!`.
pub(crate) fn finite(x: f64) -> CellValue {
    if x.is_finite() {
        num(x)
    } else {
        err(ErrorValue::Num)
    }
}
