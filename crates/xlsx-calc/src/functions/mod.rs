//! builtin function library: `resolve` folds a case-insensitive name to a
//! `Func`; `Func::call` dispatches. builtins receive arguments unevaluated so
//! control-flow can skip branches.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, as_area, err, evaluate, num, parse_num, to_number};
use crate::parser::Expr;

pub mod criteria;
pub mod datetime;
pub mod financial;
pub mod info;
pub mod logical;
pub mod lookups;
pub mod math;
pub mod stats;
pub mod text;

/// a builtin's interned identity; `call` dispatches on it so evaluation never
/// touches the name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Func {
    Sum,
    SumIf,
    SumIfs,
    SumProduct,
    Mmult,
    Product,
    Abs,
    Sign,
    Round,
    RoundUp,
    RoundDown,
    Mround,
    Quotient,
    SeriesSum,
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
    Tanh,
    Sin,
    Cos,
    Tan,
    Sinh,
    Cosh,
    Asin,
    Acos,
    Atan,
    Atan2,
    Degrees,
    Radians,
    Correl,
    CovarianceP,
    CovarianceS,
    Slope,
    Intercept,
    Percentile,
    Quartile,
    Randbetween,
    Average,
    Count,
    CountA,
    CountBlank,
    CountIf,
    CountIfs,
    AverageIf,
    AverageIfs,
    MaxIfs,
    MinIfs,
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
    AverageA,
    MaxA,
    MinA,
    StdevA,
    GeoMean,
    HarMean,
    AveDev,
    DevSq,
    Skew,
    Kurt,
    TrimMean,
    SumSq,
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
    WeekNum,
    IsoWeekNum,
    TextBefore,
    TextAfter,
    Edate,
    Eomonth,
    Today,
    Now,
    Hour,
    Minute,
    Second,
    Time,
    DateDif,
    DateValue,
    YearFrac,
    WorkdayIntl,
    NetworkdaysIntl,
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
    Offset,
    Match,
    Xlookup,
    XMatch,
    Subtotal,
    Aggregate,
    Choose,
    Row,
    Column,
    Rows,
    Columns,
    Transpose,
    IsBlank,
    IsNumber,
    IsText,
    IsLogical,
    IsError,
    IsErr,
    IsNa,
    IsEven,
    IsOdd,
    Na,
    N,
    Lookup,
    Indirect,
    Address,
    Hyperlink,
    Fv,
    NormDist,
}

/// drop the `_xlfn.` / `_xlfn._xlws.` prefix excel stores post-2007 functions
/// under; the prefix belongs to the stored name, not to the function.
pub(crate) fn bare_name(name: &str) -> &str {
    match strip_ascii_prefix(name, "_xlfn.") {
        Some(rest) => strip_ascii_prefix(rest, "_xlws.").unwrap_or(rest),
        None => name,
    }
}

fn strip_ascii_prefix<'a>(name: &'a str, prefix: &str) -> Option<&'a str> {
    name.as_bytes()
        .get(..prefix.len())
        .filter(|head| head.eq_ignore_ascii_case(prefix.as_bytes()))
        .and_then(|_| name.get(prefix.len()..))
}

/// resolve a function name (case-insensitive) to its interned id; aliases map
/// to the same id. the stored `_xlfn.` prefix is stripped first, so the
/// uppercase fold only has to hold the bare name and never allocates.
pub fn resolve(name: &str) -> Option<Func> {
    const MAX_BUILTIN_LEN: usize = 16;
    let name = bare_name(name);
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
        "MMULT" => Func::Mmult,
        "PRODUCT" => Func::Product,
        "ABS" => Func::Abs,
        "SIGN" => Func::Sign,
        "ROUND" => Func::Round,
        "ROUNDUP" => Func::RoundUp,
        "ROUNDDOWN" => Func::RoundDown,
        "MROUND" => Func::Mround,
        "QUOTIENT" => Func::Quotient,
        "SERIESSUM" => Func::SeriesSum,
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
        "TANH" => Func::Tanh,
        "SIN" => Func::Sin,
        "COS" => Func::Cos,
        "TAN" => Func::Tan,
        "SINH" => Func::Sinh,
        "COSH" => Func::Cosh,
        "ASIN" => Func::Asin,
        "ACOS" => Func::Acos,
        "ATAN" => Func::Atan,
        "ATAN2" => Func::Atan2,
        "DEGREES" => Func::Degrees,
        "RADIANS" => Func::Radians,
        "CORREL" => Func::Correl,
        "COVAR" | "COVARIANCE.P" => Func::CovarianceP,
        "COVARIANCE.S" => Func::CovarianceS,
        "SLOPE" => Func::Slope,
        "INTERCEPT" => Func::Intercept,
        "PERCENTILE" | "PERCENTILE.INC" => Func::Percentile,
        "QUARTILE" | "QUARTILE.INC" => Func::Quartile,
        "RANDBETWEEN" => Func::Randbetween,
        "AVERAGE" => Func::Average,
        "COUNT" => Func::Count,
        "COUNTA" => Func::CountA,
        "COUNTBLANK" => Func::CountBlank,
        "COUNTIF" => Func::CountIf,
        "COUNTIFS" => Func::CountIfs,
        "AVERAGEIF" => Func::AverageIf,
        "AVERAGEIFS" => Func::AverageIfs,
        "MAXIFS" => Func::MaxIfs,
        "MINIFS" => Func::MinIfs,
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
        "AVERAGEA" => Func::AverageA,
        "MAXA" => Func::MaxA,
        "MINA" => Func::MinA,
        "STDEVA" => Func::StdevA,
        "GEOMEAN" => Func::GeoMean,
        "HARMEAN" => Func::HarMean,
        "AVEDEV" => Func::AveDev,
        "DEVSQ" => Func::DevSq,
        "SKEW" => Func::Skew,
        "KURT" => Func::Kurt,
        "TRIMMEAN" => Func::TrimMean,
        "SUMSQ" => Func::SumSq,
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
        "WEEKNUM" => Func::WeekNum,
        "ISOWEEKNUM" => Func::IsoWeekNum,
        "TEXTBEFORE" => Func::TextBefore,
        "TEXTAFTER" => Func::TextAfter,
        "EDATE" => Func::Edate,
        "EOMONTH" => Func::Eomonth,
        "TODAY" => Func::Today,
        "NOW" => Func::Now,
        "HOUR" => Func::Hour,
        "MINUTE" => Func::Minute,
        "SECOND" => Func::Second,
        "TIME" => Func::Time,
        "DATEDIF" => Func::DateDif,
        "DATEVALUE" => Func::DateValue,
        "YEARFRAC" => Func::YearFrac,
        "WORKDAY.INTL" => Func::WorkdayIntl,
        "NETWORKDAYS.INTL" => Func::NetworkdaysIntl,
        "ADDRESS" => Func::Address,
        "HYPERLINK" => Func::Hyperlink,
        "FV" => Func::Fv,
        "NORM.DIST" | "NORMDIST" => Func::NormDist,
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
        "OFFSET" => Func::Offset,
        "MATCH" => Func::Match,
        "XLOOKUP" => Func::Xlookup,
        "XMATCH" => Func::XMatch,
        "SUBTOTAL" => Func::Subtotal,
        "AGGREGATE" => Func::Aggregate,
        "LOOKUP" => Func::Lookup,
        "INDIRECT" => Func::Indirect,
        "CHOOSE" => Func::Choose,
        "ROW" => Func::Row,
        "COLUMN" => Func::Column,
        "ROWS" => Func::Rows,
        "COLUMNS" => Func::Columns,
        "TRANSPOSE" => Func::Transpose,
        "ISBLANK" => Func::IsBlank,
        "ISNUMBER" => Func::IsNumber,
        "ISTEXT" => Func::IsText,
        "ISLOGICAL" => Func::IsLogical,
        "ISERROR" => Func::IsError,
        "ISERR" => Func::IsErr,
        "ISNA" => Func::IsNa,
        "ISEVEN" => Func::IsEven,
        "ISODD" => Func::IsOdd,
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
            Func::Mmult => math::mmult(args, ctx),
            Func::Product => math::product(args, ctx),
            Func::Abs => math::abs(args, ctx),
            Func::Sign => math::sign(args, ctx),
            Func::Round => math::round(args, ctx),
            Func::RoundUp => math::roundup(args, ctx),
            Func::RoundDown => math::rounddown(args, ctx),
            Func::Mround => math::mround(args, ctx),
            Func::Quotient => math::quotient(args, ctx),
            Func::SeriesSum => math::seriessum(args, ctx),
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
            Func::Tanh => math::tanh(args, ctx),
            Func::Sin => math::sin(args, ctx),
            Func::Cos => math::cos(args, ctx),
            Func::Tan => math::tan(args, ctx),
            Func::Sinh => math::sinh(args, ctx),
            Func::Cosh => math::cosh(args, ctx),
            Func::Asin => math::asin(args, ctx),
            Func::Acos => math::acos(args, ctx),
            Func::Atan => math::atan(args, ctx),
            Func::Atan2 => math::atan2(args, ctx),
            Func::Degrees => math::degrees(args, ctx),
            Func::Radians => math::radians(args, ctx),
            Func::Correl => stats::correl(args, ctx),
            Func::CovarianceP => stats::covariance_p(args, ctx),
            Func::CovarianceS => stats::covariance_s(args, ctx),
            Func::Slope => stats::slope(args, ctx),
            Func::Intercept => stats::intercept(args, ctx),
            Func::Percentile => stats::percentile(args, ctx),
            Func::Quartile => stats::quartile(args, ctx),
            Func::Randbetween => math::randbetween(args, ctx),
            Func::Average => stats::average(args, ctx),
            Func::Count => stats::count(args, ctx),
            Func::CountA => stats::counta(args, ctx),
            Func::CountBlank => stats::countblank(args, ctx),
            Func::CountIf => stats::countif(args, ctx),
            Func::CountIfs => stats::countifs(args, ctx),
            Func::AverageIf => stats::averageif(args, ctx),
            Func::AverageIfs => stats::averageifs(args, ctx),
            Func::MaxIfs => stats::maxifs(args, ctx),
            Func::MinIfs => stats::minifs(args, ctx),
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
            Func::AverageA => stats::averagea(args, ctx),
            Func::MaxA => stats::maxa(args, ctx),
            Func::MinA => stats::mina(args, ctx),
            Func::StdevA => stats::stdeva(args, ctx),
            Func::GeoMean => stats::geomean(args, ctx),
            Func::HarMean => stats::harmean(args, ctx),
            Func::AveDev => stats::avedev(args, ctx),
            Func::DevSq => stats::devsq(args, ctx),
            Func::Skew => stats::skew(args, ctx),
            Func::Kurt => stats::kurt(args, ctx),
            Func::TrimMean => stats::trimmean(args, ctx),
            Func::SumSq => math::sumsq(args, ctx),
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
            Func::WeekNum => datetime::weeknum(args, ctx),
            Func::IsoWeekNum => datetime::isoweeknum(args, ctx),
            Func::TextBefore => text::textbefore(args, ctx),
            Func::TextAfter => text::textafter(args, ctx),
            Func::Edate => datetime::edate(args, ctx),
            Func::Eomonth => datetime::eomonth(args, ctx),
            Func::Today => datetime::today(args, ctx),
            Func::Now => datetime::now(args, ctx),
            Func::Hour => datetime::hour(args, ctx),
            Func::Minute => datetime::minute(args, ctx),
            Func::Second => datetime::second(args, ctx),
            Func::Time => datetime::time(args, ctx),
            Func::DateDif => datetime::datedif(args, ctx),
            Func::DateValue => datetime::datevalue(args, ctx),
            Func::YearFrac => datetime::yearfrac(args, ctx),
            Func::WorkdayIntl => datetime::workday_intl(args, ctx),
            Func::NetworkdaysIntl => datetime::networkdays_intl(args, ctx),
            Func::Address => lookups::address(args, ctx),
            Func::Hyperlink => lookups::hyperlink(args, ctx),
            Func::Fv => financial::fv(args, ctx),
            Func::NormDist => stats::norm_dist(args, ctx),
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
            Func::Offset => lookups::offset(args, ctx),
            Func::Match => lookups::match_(args, ctx),
            Func::Xlookup => lookups::xlookup(args, ctx),
            Func::XMatch => lookups::xmatch(args, ctx),
            Func::Subtotal => stats::subtotal(args, ctx),
            Func::Aggregate => stats::aggregate(args, ctx),
            Func::Lookup => lookups::lookup_fn(args, ctx),
            Func::Indirect => lookups::indirect(args, ctx),
            Func::Choose => lookups::choose(args, ctx),
            Func::Row => lookups::row(args, ctx),
            Func::Column => lookups::column(args, ctx),
            Func::Rows => lookups::rows(args, ctx),
            Func::Columns => lookups::columns(args, ctx),
            Func::Transpose => lookups::transpose(args, ctx),
            Func::IsBlank => info::isblank(args, ctx),
            Func::IsNumber => info::isnumber(args, ctx),
            Func::IsText => info::istext(args, ctx),
            Func::IsLogical => info::islogical(args, ctx),
            Func::IsError => info::iserror(args, ctx),
            Func::IsErr => info::iserr(args, ctx),
            Func::IsNa => info::isna(args, ctx),
            Func::IsEven => info::iseven(args, ctx),
            Func::IsOdd => info::isodd(args, ctx),
            Func::Na => info::na(args, ctx),
            Func::N => info::n(args, ctx),
        }
    }
}

/// what one argument contributes, and whether it came from the sheet: a
/// reference lends its own cells, anything else the block it evaluates to, so
/// `{1;2;3}` reads like `A1:A3`. a lone computed value still counts as typed
/// into the call, where excel coerces instead of skipping.
pub(crate) fn argument_cells(
    arg: &Expr,
    ctx: &EvalContext<'_>,
) -> Result<(Vec<CellValue>, bool), ErrorValue> {
    if let Some(area) = as_area(arg, ctx) {
        let cells = area
            .values_ref(ctx)?
            .into_iter()
            .map(std::borrow::Cow::into_owned)
            .collect();
        return Ok((cells, true));
    }
    let cells = crate::array::evaluate_array(arg, ctx)
        .into_array()
        .cells()
        .to_vec();
    let from_sheet = cells.len() > 1;
    Ok((cells, from_sheet))
}

/// like [`collect_numbers`], but a computed block contributes every numeric
/// cell instead of only its first, so `MEDIAN(IF(..))` sees the whole block.
pub(crate) fn collect_numbers_deep(
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = Vec::new();
    for arg in args {
        let (cells, from_sheet) = argument_cells(arg, ctx)?;
        for value in cells {
            if from_sheet {
                push_reference_number(&mut nums, &value)?;
                continue;
            }
            match value {
                CellValue::Number { value } => nums.push(value),
                CellValue::Bool { value } => nums.push(if value { 1.0 } else { 0.0 }),
                CellValue::Empty => {}
                CellValue::Text { value } => match parse_num(&value) {
                    Some(n) => nums.push(n),
                    None => return Err(ErrorValue::Value),
                },
                CellValue::Error { value } => return Err(value),
            }
        }
    }
    Ok(nums)
}

/// the values the `-A` aggregates see: text counts as zero and logicals as
/// one or zero, where the plain forms skip both. blanks are still skipped,
/// and text typed straight into the call must still be a number.
pub(crate) fn collect_numbers_anytype(
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = Vec::new();
    for arg in args {
        let (cells, from_sheet) = argument_cells(arg, ctx)?;
        for value in cells {
            match value {
                CellValue::Number { value } => nums.push(value),
                CellValue::Bool { value } => nums.push(if value { 1.0 } else { 0.0 }),
                CellValue::Error { value } => return Err(value),
                CellValue::Empty => {}
                CellValue::Text { .. } if from_sheet => nums.push(0.0),
                CellValue::Text { value } => match parse_num(&value) {
                    Some(n) => nums.push(n),
                    None => return Err(ErrorValue::Value),
                },
            }
        }
    }
    Ok(nums)
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
                for value in area.values_ref(ctx)? {
                    push_reference_number(&mut nums, &value)?;
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
fn push_reference_number(nums: &mut Vec<f64>, v: &CellValue) -> Result<(), ErrorValue> {
    match v {
        CellValue::Number { value } => nums.push(*value),
        CellValue::Error { value } => return Err(*value),
        _ => {}
    }
    Ok(())
}

/// whether an argument was written as a gap, as in `OFFSET(a,,,n,)`. excel
/// reads those as absent, not as zero.
pub(crate) fn omitted(arg: &Expr) -> bool {
    matches!(arg, Expr::Literal(CellValue::Empty))
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

#[cfg(test)]
mod tests {
    use super::{Func, resolve};

    #[test]
    fn every_builtin_name_and_alias_resolves_to_its_variant() {
        let canonical: &[(Func, &str)] = &[
            (Func::Sum, "SUM"),
            (Func::SumIf, "SUMIF"),
            (Func::SumIfs, "SUMIFS"),
            (Func::SumProduct, "SUMPRODUCT"),
            (Func::Mmult, "MMULT"),
            (Func::Product, "PRODUCT"),
            (Func::Abs, "ABS"),
            (Func::Sign, "SIGN"),
            (Func::Round, "ROUND"),
            (Func::RoundUp, "ROUNDUP"),
            (Func::RoundDown, "ROUNDDOWN"),
            (Func::Mround, "MROUND"),
            (Func::Quotient, "QUOTIENT"),
            (Func::SeriesSum, "SERIESSUM"),
            (Func::Ceiling, "CEILING"),
            (Func::Floor, "FLOOR"),
            (Func::Int, "INT"),
            (Func::Trunc, "TRUNC"),
            (Func::Mod, "MOD"),
            (Func::Power, "POWER"),
            (Func::Sqrt, "SQRT"),
            (Func::Exp, "EXP"),
            (Func::Ln, "LN"),
            (Func::Log, "LOG"),
            (Func::Log10, "LOG10"),
            (Func::Pi, "PI"),
            (Func::Tanh, "TANH"),
            (Func::Randbetween, "RANDBETWEEN"),
            (Func::Average, "AVERAGE"),
            (Func::Count, "COUNT"),
            (Func::CountA, "COUNTA"),
            (Func::CountBlank, "COUNTBLANK"),
            (Func::CountIf, "COUNTIF"),
            (Func::CountIfs, "COUNTIFS"),
            (Func::AverageIf, "AVERAGEIF"),
            (Func::AverageIfs, "AVERAGEIFS"),
            (Func::MaxIfs, "MAXIFS"),
            (Func::MinIfs, "MINIFS"),
            (Func::Min, "MIN"),
            (Func::Max, "MAX"),
            (Func::Median, "MEDIAN"),
            (Func::Mode, "MODE"),
            (Func::StdevS, "STDEV"),
            (Func::StdevP, "STDEVP"),
            (Func::VarS, "VAR"),
            (Func::VarP, "VARP"),
            (Func::Large, "LARGE"),
            (Func::Small, "SMALL"),
            (Func::Rank, "RANK"),
            (Func::AverageA, "AVERAGEA"),
            (Func::MaxA, "MAXA"),
            (Func::MinA, "MINA"),
            (Func::StdevA, "STDEVA"),
            (Func::GeoMean, "GEOMEAN"),
            (Func::HarMean, "HARMEAN"),
            (Func::AveDev, "AVEDEV"),
            (Func::DevSq, "DEVSQ"),
            (Func::Skew, "SKEW"),
            (Func::Kurt, "KURT"),
            (Func::TrimMean, "TRIMMEAN"),
            (Func::SumSq, "SUMSQ"),
            (Func::Len, "LEN"),
            (Func::Left, "LEFT"),
            (Func::Right, "RIGHT"),
            (Func::Mid, "MID"),
            (Func::Find, "FIND"),
            (Func::Search, "SEARCH"),
            (Func::Substitute, "SUBSTITUTE"),
            (Func::Replace, "REPLACE"),
            (Func::Trim, "TRIM"),
            (Func::Upper, "UPPER"),
            (Func::Lower, "LOWER"),
            (Func::Proper, "PROPER"),
            (Func::Clean, "CLEAN"),
            (Func::Rept, "REPT"),
            (Func::Exact, "EXACT"),
            (Func::T, "T"),
            (Func::Char, "CHAR"),
            (Func::Code, "CODE"),
            (Func::Value, "VALUE"),
            (Func::NumberValue, "NUMBERVALUE"),
            (Func::Text, "TEXT"),
            (Func::TextJoin, "TEXTJOIN"),
            (Func::Concat, "CONCATENATE"),
            (Func::Date, "DATE"),
            (Func::Year, "YEAR"),
            (Func::Month, "MONTH"),
            (Func::Day, "DAY"),
            (Func::Weekday, "WEEKDAY"),
            (Func::Edate, "EDATE"),
            (Func::Eomonth, "EOMONTH"),
            (Func::Today, "TODAY"),
            (Func::Now, "NOW"),
            (Func::Hour, "HOUR"),
            (Func::Minute, "MINUTE"),
            (Func::Second, "SECOND"),
            (Func::Time, "TIME"),
            (Func::DateDif, "DATEDIF"),
            (Func::DateValue, "DATEVALUE"),
            (Func::YearFrac, "YEARFRAC"),
            (Func::WorkdayIntl, "WORKDAY.INTL"),
            (Func::NetworkdaysIntl, "NETWORKDAYS.INTL"),
            (Func::Address, "ADDRESS"),
            (Func::Hyperlink, "HYPERLINK"),
            (Func::Fv, "FV"),
            (Func::NormDist, "NORM.DIST"),
            (Func::If, "IF"),
            (Func::IfError, "IFERROR"),
            (Func::IfNa, "IFNA"),
            (Func::Ifs, "IFS"),
            (Func::Switch, "SWITCH"),
            (Func::And, "AND"),
            (Func::Or, "OR"),
            (Func::Not, "NOT"),
            (Func::Xor, "XOR"),
            (Func::Vlookup, "VLOOKUP"),
            (Func::Hlookup, "HLOOKUP"),
            (Func::Index, "INDEX"),
            (Func::Offset, "OFFSET"),
            (Func::Match, "MATCH"),
            (Func::Xlookup, "XLOOKUP"),
            (Func::Lookup, "LOOKUP"),
            (Func::Indirect, "INDIRECT"),
            (Func::Choose, "CHOOSE"),
            (Func::Row, "ROW"),
            (Func::Column, "COLUMN"),
            (Func::Rows, "ROWS"),
            (Func::Columns, "COLUMNS"),
            (Func::Transpose, "TRANSPOSE"),
            (Func::IsBlank, "ISBLANK"),
            (Func::IsNumber, "ISNUMBER"),
            (Func::IsText, "ISTEXT"),
            (Func::IsLogical, "ISLOGICAL"),
            (Func::IsError, "ISERROR"),
            (Func::IsErr, "ISERR"),
            (Func::IsNa, "ISNA"),
            (Func::IsEven, "ISEVEN"),
            (Func::IsOdd, "ISODD"),
            (Func::Na, "NA"),
            (Func::N, "N"),
        ];
        for &(func, name) in canonical {
            assert_eq!(resolve(name), Some(func), "{name}");
            assert_eq!(resolve(&name.to_ascii_lowercase()), Some(func), "{name}");
        }
        let aliases: &[(&str, Func)] = &[
            ("MODE.SNGL", Func::Mode),
            ("STDEV.S", Func::StdevS),
            ("STDEV.P", Func::StdevP),
            ("VAR.S", Func::VarS),
            ("VAR.P", Func::VarP),
            ("RANK.EQ", Func::Rank),
            ("CONCAT", Func::Concat),
            ("NORMDIST", Func::NormDist),
        ];
        for &(name, func) in aliases {
            assert_eq!(resolve(name), Some(func), "{name}");
        }
    }
}
