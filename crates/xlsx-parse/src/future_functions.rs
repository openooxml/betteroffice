//! The functions Excel stores under a version prefix, and the save-time pass
//! that writes formula text the way Excel does: `_xlfn.` before every function
//! newer than Excel 2007, `_xlfn._xlws.` before the worksheet ones, `_xlpm.`
//! before the names a `LET` or `LAMBDA` binds, and `_xleta.` before a function
//! passed by name to a lambda helper. [`display_formula`] is the reverse, for
//! formula text shown to people.

use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::HashMap;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};

use crate::formula::{MAX_FORMULA_BYTES, skip_brackets, skip_quoted};

/// every function Excel writes behind `_xlfn.`, upper case and sorted.
const FUTURE_FUNCTIONS: &[&str] = &[
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ANCHORARRAY",
    "ARABIC",
    "ARRAYTOTEXT",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DROP",
    "ECMA.CEILING",
    "ENCODEURL",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FIELDVALUE",
    "FILTER",
    "FILTERXML",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "GROUPBY",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "IMAGE",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISO.CEILING",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTOF",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "PIVOTBY",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "QUERYSTRING",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "REGEXEXTRACT",
    "REGEXREPLACE",
    "REGEXTEST",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SINGLE",
    "SKEW.P",
    "SORT",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "STOCKHISTORY",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRIMRANGE",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VALUETOTEXT",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEBSERVICE",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
];

/// every other Excel worksheet function, upper case and sorted. a call of one
/// of these or of a future function stays a function call even where a `LET`
/// or `LAMBDA` binds the same name.
const CLASSIC_FUNCTIONS: &[&str] = &[
    "ABS",
    "ACCRINT",
    "ACCRINTM",
    "ACOS",
    "ACOSH",
    "ADDRESS",
    "AMORDEGRC",
    "AMORLINC",
    "AND",
    "AREAS",
    "ASC",
    "ASIN",
    "ASINH",
    "ATAN",
    "ATAN2",
    "ATANH",
    "AVEDEV",
    "AVERAGE",
    "AVERAGEA",
    "AVERAGEIF",
    "AVERAGEIFS",
    "BAHTTEXT",
    "BESSELI",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "BETADIST",
    "BETAINV",
    "BIN2DEC",
    "BIN2HEX",
    "BIN2OCT",
    "BINOMDIST",
    "CALL",
    "CEILING",
    "CELL",
    "CHAR",
    "CHIDIST",
    "CHIINV",
    "CHITEST",
    "CHOOSE",
    "CLEAN",
    "CODE",
    "COLUMN",
    "COLUMNS",
    "COMBIN",
    "COMPLEX",
    "CONCATENATE",
    "CONFIDENCE",
    "CONVERT",
    "COPILOT",
    "CORREL",
    "COS",
    "COSH",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "COUNTIF",
    "COUNTIFS",
    "COUPDAYBS",
    "COUPDAYS",
    "COUPDAYSNC",
    "COUPNCD",
    "COUPNUM",
    "COUPPCD",
    "COVAR",
    "CRITBINOM",
    "CUBEKPIMEMBER",
    "CUBEMEMBER",
    "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "CUBEVALUE",
    "CUMIPMT",
    "CUMPRINC",
    "DATE",
    "DATEDIF",
    "DATEVALUE",
    "DAVERAGE",
    "DAY",
    "DAYS360",
    "DB",
    "DBCS",
    "DCOUNT",
    "DCOUNTA",
    "DDB",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "DEGREES",
    "DELTA",
    "DETECTLANGUAGE",
    "DEVSQ",
    "DGET",
    "DISC",
    "DMAX",
    "DMIN",
    "DOLLAR",
    "DOLLARDE",
    "DOLLARFR",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DSUM",
    "DURATION",
    "DVAR",
    "DVARP",
    "EDATE",
    "EFFECT",
    "EOMONTH",
    "ERF",
    "ERFC",
    "ERROR.TYPE",
    "EUROCONVERT",
    "EVEN",
    "EXACT",
    "EXP",
    "EXPONDIST",
    "FACT",
    "FACTDOUBLE",
    "FALSE",
    "FDIST",
    "FIND",
    "FINDB",
    "FINV",
    "FISHER",
    "FISHERINV",
    "FIXED",
    "FLOOR",
    "FORECAST",
    "FREQUENCY",
    "FTEST",
    "FV",
    "FVSCHEDULE",
    "GAMMADIST",
    "GAMMAINV",
    "GAMMALN",
    "GCD",
    "GEOMEAN",
    "GESTEP",
    "GETPIVOTDATA",
    "GROWTH",
    "HARMEAN",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "HLOOKUP",
    "HOUR",
    "HYPERLINK",
    "HYPGEOMDIST",
    "IF",
    "IFERROR",
    "IMABS",
    "IMAGINARY",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMCOS",
    "IMDIV",
    "IMEXP",
    "IMLN",
    "IMLOG10",
    "IMLOG2",
    "IMPOWER",
    "IMPRODUCT",
    "IMREAL",
    "IMSIN",
    "IMSQRT",
    "IMSUB",
    "IMSUM",
    "INDEX",
    "INDIRECT",
    "INFO",
    "INT",
    "INTERCEPT",
    "INTRATE",
    "IPMT",
    "IRR",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISEVEN",
    "ISLOGICAL",
    "ISNA",
    "ISNONTEXT",
    "ISNUMBER",
    "ISODD",
    "ISPMT",
    "ISREF",
    "ISTEXT",
    "JIS",
    "KURT",
    "LARGE",
    "LCM",
    "LEFT",
    "LEFTB",
    "LEN",
    "LENB",
    "LINEST",
    "LN",
    "LOG",
    "LOG10",
    "LOGEST",
    "LOGINV",
    "LOGNORMDIST",
    "LOOKUP",
    "LOWER",
    "MATCH",
    "MAX",
    "MAXA",
    "MDETERM",
    "MDURATION",
    "MEDIAN",
    "MID",
    "MIDB",
    "MIN",
    "MINA",
    "MINUTE",
    "MINVERSE",
    "MIRR",
    "MMULT",
    "MOD",
    "MODE",
    "MONTH",
    "MROUND",
    "MULTINOMIAL",
    "N",
    "NA",
    "NEGBINOMDIST",
    "NETWORKDAYS",
    "NOMINAL",
    "NORMDIST",
    "NORMINV",
    "NORMSDIST",
    "NORMSINV",
    "NOT",
    "NOW",
    "NPER",
    "NPV",
    "OCT2BIN",
    "OCT2DEC",
    "OCT2HEX",
    "ODD",
    "ODDFPRICE",
    "ODDFYIELD",
    "ODDLPRICE",
    "ODDLYIELD",
    "OFFSET",
    "OR",
    "PEARSON",
    "PERCENTILE",
    "PERCENTRANK",
    "PERMUT",
    "PHONETIC",
    "PI",
    "PMT",
    "POISSON",
    "POWER",
    "PPMT",
    "PRICE",
    "PRICEDISC",
    "PRICEMAT",
    "PROB",
    "PRODUCT",
    "PROPER",
    "PV",
    "PY",
    "QUARTILE",
    "QUOTIENT",
    "RADIANS",
    "RAND",
    "RANDBETWEEN",
    "RANK",
    "RATE",
    "RECEIVED",
    "REGISTER.ID",
    "REPLACE",
    "REPLACEB",
    "REPT",
    "RIGHT",
    "RIGHTB",
    "ROMAN",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "ROW",
    "ROWS",
    "RSQ",
    "RTD",
    "SEARCH",
    "SEARCHB",
    "SECOND",
    "SERIESSUM",
    "SIGN",
    "SIN",
    "SINH",
    "SKEW",
    "SLN",
    "SLOPE",
    "SMALL",
    "SQRT",
    "SQRTPI",
    "STANDARDIZE",
    "STDEV",
    "STDEVA",
    "STDEVP",
    "STDEVPA",
    "STEYX",
    "SUBSTITUTE",
    "SUBTOTAL",
    "SUM",
    "SUMIF",
    "SUMIFS",
    "SUMPRODUCT",
    "SUMSQ",
    "SUMX2MY2",
    "SUMX2PY2",
    "SUMXMY2",
    "SYD",
    "T",
    "TAN",
    "TANH",
    "TBILLEQ",
    "TBILLPRICE",
    "TBILLYIELD",
    "TDIST",
    "TEXT",
    "TIME",
    "TIMEVALUE",
    "TINV",
    "TODAY",
    "TRANSLATE",
    "TRANSPOSE",
    "TREND",
    "TRIM",
    "TRIMMEAN",
    "TRUE",
    "TRUNC",
    "TTEST",
    "TYPE",
    "UPPER",
    "VALUE",
    "VAR",
    "VARA",
    "VARP",
    "VARPA",
    "VDB",
    "VLOOKUP",
    "WEEKDAY",
    "WEEKNUM",
    "WEIBULL",
    "WORKDAY",
    "XIRR",
    "XNPV",
    "YEAR",
    "YEARFRAC",
    "YIELD",
    "YIELDDISC",
    "YIELDMAT",
    "ZTEST",
];

/// the future functions Excel writes as `_xlfn._xlws.NAME`.
const WORKSHEET_FUNCTIONS: &[&str] = &["FILTER", "SORT"];

const FUNCTION_PREFIX: &str = "_xlfn.";
const WORKSHEET_PREFIX: &str = "_xlws.";
const WORKSHEET_FUNCTION_PREFIX: &str = "_xlfn._xlws.";
const PARAMETER_PREFIX: &str = "_xlpm.";
const ETA_PREFIX: &str = "_xleta.";

/// the argument of a lambda helper that takes its function: the one at an
/// index, or the last of two or more.
#[derive(Clone, Copy)]
enum Callback {
    At(usize),
    Last,
}

/// the functions that take a lambda, with the argument where a bare function
/// name stands for one. Every other argument is data, dimensions or an initial
/// value, where a bare name is a defined name.
const LAMBDA_HELPERS: &[(&str, Callback)] = &[
    ("BYCOL", Callback::At(1)),
    ("BYROW", Callback::At(1)),
    ("GROUPBY", Callback::At(2)),
    ("MAKEARRAY", Callback::At(2)),
    ("MAP", Callback::Last),
    ("PIVOTBY", Callback::At(3)),
    ("REDUCE", Callback::At(2)),
    ("SCAN", Callback::At(2)),
];

/// `formula` as Excel stores it. Text already in that form comes back
/// borrowed, so the pass is idempotent, and so does text the prefixes would
/// push past the formula length limit, which this crate could not read back,
/// or that is already past it.
pub fn excel_formula(formula: &str) -> Cow<'_, str> {
    if formula.len() > MAX_FORMULA_BYTES || !formula.contains('(') {
        return Cow::Borrowed(formula);
    }
    let tokens = tokenize(formula);
    let mut pass = Pass::new(formula, &tokens);
    pass.run();
    let grown = pass
        .edits
        .iter()
        .map(|edit| edit.text.len() - (edit.end - edit.start).min(edit.text.len()))
        .sum::<usize>();
    if formula.len().saturating_add(grown) > MAX_FORMULA_BYTES {
        return Cow::Borrowed(formula);
    }
    splice(formula, &pass.edits)
}

/// `formula` as people read and type it: the storage prefixes `_xlfn.`,
/// `_xlws.`, `_xlpm.` and `_xleta.` dropped from function names and `LET` or
/// `LAMBDA` names, never from string literals, sheet names or structured
/// references. A stripped text is shown only when it [`saves_as`] `formula`,
/// otherwise a text keeping `_xleta.`, otherwise `formula` itself, so the
/// result is idempotent. Text past the formula length limit comes back as it
/// is; below it the work is linear: one tokenizing pass and at most two
/// rewrites.
pub fn display_formula(formula: &str) -> Cow<'_, str> {
    if formula.len() > MAX_FORMULA_BYTES
        || !formula
            .as_bytes()
            .windows(3)
            .any(|window| window.eq_ignore_ascii_case(b"_xl"))
    {
        return Cow::Borrowed(formula);
    }
    let tokens = tokenize(formula);
    let pass = Pass::new(formula, &tokens);
    for keep_eta in [false, true] {
        let shown = splice(formula, &pass.unprefixed(keep_eta));
        if matches!(shown, Cow::Borrowed(_)) || saves_as(&shown, formula) {
            return shown;
        }
    }
    Cow::Borrowed(formula)
}

/// Whether formula text people typed or were shown saves as exactly `stored`,
/// byte for byte. The one test for showing a text in place of the stored one
/// and for keeping the stored text when that text is committed.
pub fn saves_as(typed: &str, stored: &str) -> bool {
    excel_formula(typed) == stored
}

/// `formula` with the source-ordered `edits` applied.
fn splice<'a>(formula: &'a str, edits: &[Edit]) -> Cow<'a, str> {
    if edits.is_empty() {
        return Cow::Borrowed(formula);
    }
    let mut out = String::with_capacity(formula.len() + 8 * edits.len());
    let mut cursor = 0;
    for edit in edits {
        out.push_str(&formula[cursor..edit.start]);
        out.push_str(edit.text);
        cursor = edit.end;
    }
    out.push_str(&formula[cursor..]);
    Cow::Owned(out)
}

/// `name` without its leading storage prefixes, keeping `_xleta.` when asked
/// and never stripping a name down to nothing.
fn unprefixed(name: &str, keep_eta: bool) -> &str {
    let prefixes: &[&str] = if keep_eta {
        &[FUNCTION_PREFIX, WORKSHEET_PREFIX, PARAMETER_PREFIX]
    } else {
        &[
            FUNCTION_PREFIX,
            WORKSHEET_PREFIX,
            PARAMETER_PREFIX,
            ETA_PREFIX,
        ]
    };
    let mut rest = name;
    while let Some(next) = prefixes
        .iter()
        .find_map(|prefix| strip_prefix(rest, prefix))
        .filter(|next| !next.is_empty())
    {
        rest = next;
    }
    rest
}

/// the canonical spelling of a future function and the prefix it is stored
/// under, matched case-insensitively.
fn future_function(name: &str) -> Option<(&'static str, &'static str)> {
    let index = FUTURE_FUNCTIONS
        .binary_search_by(|probe| compare_upper(probe, name))
        .ok()?;
    let canonical = FUTURE_FUNCTIONS[index];
    let prefix = if WORKSHEET_FUNCTIONS.contains(&canonical) {
        WORKSHEET_FUNCTION_PREFIX
    } else {
        FUNCTION_PREFIX
    };
    Some((canonical, prefix))
}

fn builtin(name: &str) -> bool {
    canonical(name).is_some()
}

/// the canonical spelling of any worksheet function, matched case-insensitively.
fn canonical(name: &str) -> Option<&'static str> {
    match future_function(name) {
        Some((canonical, _)) => Some(canonical),
        None => CLASSIC_FUNCTIONS
            .binary_search_by(|probe| compare_upper(probe, name))
            .ok()
            .map(|index| CLASSIC_FUNCTIONS[index]),
    }
}

/// the spelling of a function passed by name to a lambda helper. `TRUE` and
/// `FALSE` stay the booleans they read as.
fn eta_function(name: &str) -> Option<&'static str> {
    canonical(name).filter(|name| !matches!(*name, "TRUE" | "FALSE"))
}

fn compare_upper(probe: &str, name: &str) -> Ordering {
    probe
        .bytes()
        .cmp(name.bytes().map(|byte| byte.to_ascii_uppercase()))
}

fn strip_prefix<'a>(name: &'a str, prefix: &str) -> Option<&'a str> {
    name.get(..prefix.len())
        .filter(|head| head.eq_ignore_ascii_case(prefix))
        .map(|_| &name[prefix.len()..])
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    Name,
    Open,
    Close,
    Comma,
    BraceOpen,
    BraceClose,
    Bang,
    Colon,
    Bracket,
    Space,
    Other,
}

#[derive(Clone, Copy)]
struct Token {
    kind: Kind,
    start: usize,
    end: usize,
}

/// split formula text into the pieces the pass cares about. string literals,
/// quoted sheet names and bracketed references are single opaque tokens.
fn tokenize(source: &str) -> Vec<Token> {
    let bytes = source.as_bytes();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < bytes.len() {
        let start = index;
        let kind = match bytes[index] {
            quote @ (b'"' | b'\'') => {
                index = skip_quoted(source, index, quote);
                Kind::Other
            }
            b'[' => {
                index = skip_brackets(source, index);
                Kind::Bracket
            }
            byte if name_byte(byte) => {
                while index < bytes.len() && name_byte(bytes[index]) {
                    index += 1;
                }
                Kind::Name
            }
            byte if byte.is_ascii_whitespace() => {
                while index < bytes.len() && bytes[index].is_ascii_whitespace() {
                    index += 1;
                }
                Kind::Space
            }
            byte => {
                index += 1;
                match byte {
                    b'(' => Kind::Open,
                    b')' => Kind::Close,
                    b',' => Kind::Comma,
                    b'{' => Kind::BraceOpen,
                    b'}' => Kind::BraceClose,
                    b'!' => Kind::Bang,
                    b':' => Kind::Colon,
                    _ => Kind::Other,
                }
            }
        };
        tokens.push(Token {
            kind,
            start,
            end: index,
        });
    }
    tokens
}

fn name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b'$' | b'\\') || !byte.is_ascii()
}

/// a column a range can end on: one to three letters up to `XFD`.
fn column_shaped(text: &str) -> bool {
    let letters = text.strip_prefix('$').unwrap_or(text);
    (1..=3).contains(&letters.len())
        && letters.bytes().all(|byte| byte.is_ascii_alphabetic())
        && letters.bytes().fold(0u32, |column, byte| {
            column * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1)
        }) <= MAX_COLS
}

/// a name a `LET` or `LAMBDA` may bind: not a number, boolean, A1 or R1C1
/// reference, and free of `$`.
fn parameter_name(text: &str) -> bool {
    let Some(&first) = text.as_bytes().first() else {
        return false;
    };
    if !(first.is_ascii_alphabetic() || matches!(first, b'_' | b'\\') || !first.is_ascii())
        || text.contains('$')
        || text.eq_ignore_ascii_case("TRUE")
        || text.eq_ignore_ascii_case("FALSE")
    {
        return false;
    }
    let letters = text.bytes().take_while(u8::is_ascii_alphabetic).count();
    let (column, row) = text.split_at(letters);
    let cell = column_shaped(column)
        && !row.is_empty()
        && row.bytes().all(|byte| byte.is_ascii_digit())
        && row
            .parse::<u32>()
            .is_ok_and(|row| (1..=MAX_ROWS).contains(&row));
    !cell && !r1c1_shaped(text)
}

/// `RnCn` with either number left out, as in `R1C1`, `RC2` or `RC`.
fn r1c1_shaped(text: &str) -> bool {
    let Some(rest) = strip_prefix(text, "R") else {
        return false;
    };
    let rest = rest.trim_start_matches(|c: char| c.is_ascii_digit());
    strip_prefix(rest, "C").is_some_and(|rest| rest.bytes().all(|byte| byte.is_ascii_digit()))
}

/// text to put in place of `start..end` of the source.
struct Edit {
    start: usize,
    end: usize,
    text: &'static str,
}

#[derive(Clone, Copy)]
enum Binder {
    Let,
    Lambda,
}

/// what the walk does on reaching a token, besides rewriting it.
#[derive(Default)]
enum Action {
    #[default]
    None,
    /// a name a binder declares: prefix it and bring it into scope, at once
    /// or on reaching the given token.
    Declare { visible_at: Option<usize> },
    /// a `LET` value ends here, so its name comes into scope.
    Bring(String),
    /// a binder closes here, so the scope drops back to this many names.
    Close(usize),
}

/// one flat left-to-right walk over the tokens, so edits come out in source
/// order and nesting costs no stack.
struct Pass<'a> {
    source: &'a str,
    tokens: &'a [Token],
    /// the closing token of every opening parenthesis or brace, paired on one
    /// stack so pairs nest; a close that does not match the innermost open
    /// pairs with nothing.
    partners: Vec<Option<usize>>,
    actions: Vec<Action>,
    /// the lower-cased names bound by every enclosing `LET` and `LAMBDA`,
    /// innermost last, with how often each is bound.
    scopes: Vec<String>,
    bound: HashMap<String, usize>,
    /// the name tokens that are the whole callback argument of a lambda helper.
    eta: Vec<bool>,
    edits: Vec<Edit>,
}

impl<'a> Pass<'a> {
    fn new(source: &'a str, tokens: &'a [Token]) -> Self {
        let mut partners = vec![None; tokens.len()];
        let mut open = Vec::new();
        for (index, token) in tokens.iter().enumerate() {
            let opening = match token.kind {
                Kind::Open | Kind::BraceOpen => {
                    open.push(index);
                    continue;
                }
                Kind::Close => Kind::Open,
                Kind::BraceClose => Kind::BraceOpen,
                _ => continue,
            };
            if let Some(&top) = open.last()
                && tokens[top].kind == opening
            {
                open.pop();
                partners[top] = Some(index);
            }
        }
        Self {
            source,
            tokens,
            partners,
            actions: std::iter::repeat_with(Action::default)
                .take(tokens.len())
                .collect(),
            scopes: Vec::new(),
            bound: HashMap::new(),
            eta: vec![false; tokens.len()],
            edits: Vec::new(),
        }
    }

    /// the edits that drop the storage prefixes of every name outside a sheet
    /// qualifier or structured reference.
    fn unprefixed(&self, keep_eta: bool) -> Vec<Edit> {
        self.tokens
            .iter()
            .enumerate()
            .filter(|&(index, token)| {
                token.kind == Kind::Name && !self.qualified(index) && !self.next_to_bracket(index)
            })
            .filter_map(|(_, token)| {
                let text = &self.source[token.start..token.end];
                let cut = text.len() - unprefixed(text, keep_eta).len();
                (cut > 0).then_some(Edit {
                    start: token.start,
                    end: token.start + cut,
                    text: "",
                })
            })
            .collect()
    }

    fn run(&mut self) {
        for index in 0..self.tokens.len() {
            match std::mem::take(&mut self.actions[index]) {
                Action::None => self.term(index),
                Action::Declare { visible_at } => self.declare(index, visible_at),
                Action::Bring(name) => self.bring(name),
                Action::Close(frame) => {
                    for name in self.scopes.drain(frame..) {
                        if let Some(count) = self.bound.get_mut(&name) {
                            *count -= 1;
                            if *count == 0 {
                                self.bound.remove(&name);
                            }
                        }
                    }
                }
            }
        }
    }

    fn bring(&mut self, name: String) {
        *self.bound.entry(name.clone()).or_default() += 1;
        self.scopes.push(name);
    }

    /// rewrite the token at `index`.
    fn term(&mut self, index: usize) {
        let token = self.tokens[index];
        if token.kind != Kind::Name || self.qualified(index) {
            return;
        }
        let text = &self.source[token.start..token.end];
        let parameter = strip_prefix(text, PARAMETER_PREFIX);
        let call = self.kind(index + 1) == Some(Kind::Open);
        let used = if call {
            !builtin(parameter.unwrap_or(text))
        } else {
            !self.reference_part(index, text)
        };
        if used && self.is_bound(text) {
            if parameter.is_none() {
                self.insert(token.start, PARAMETER_PREFIX);
            }
            return;
        }
        if !call {
            if self.eta[index]
                && let Some(name) = eta_function(text)
            {
                self.insert(token.start, ETA_PREFIX);
                self.edits.push(Edit {
                    start: token.start,
                    end: token.end,
                    text: name,
                });
            }
            return;
        }
        let bare = match strip_prefix(text, FUNCTION_PREFIX) {
            Some(rest) => strip_prefix(rest, WORKSHEET_PREFIX).unwrap_or(rest),
            None => {
                if let Some((canonical, prefix)) = future_function(text) {
                    self.insert(token.start, prefix);
                    self.edits.push(Edit {
                        start: token.start,
                        end: token.end,
                        text: canonical,
                    });
                }
                text
            }
        };
        let binder = if bare.eq_ignore_ascii_case("LET") {
            Binder::Let
        } else if bare.eq_ignore_ascii_case("LAMBDA") {
            Binder::Lambda
        } else {
            if let Ok(helper) =
                LAMBDA_HELPERS.binary_search_by(|(probe, _)| compare_upper(probe, bare))
                && let Some(close) = self.partners[index + 1]
            {
                let arguments = self.arguments(index + 1, close);
                let slot = match LAMBDA_HELPERS[helper].1 {
                    Callback::At(slot) => Some(slot),
                    Callback::Last => arguments.len().checked_sub(1).filter(|&last| last > 0),
                };
                if let Some(at) = slot
                    .and_then(|slot| arguments.get(slot))
                    .and_then(|&(lo, hi)| self.single(lo, hi))
                {
                    self.eta[at] = true;
                }
            }
            return;
        };
        if let Some(close) = self.partners[index + 1] {
            let arguments = self.arguments(index + 1, close);
            self.bind(binder, &arguments, close);
        }
    }

    fn kind(&self, index: usize) -> Option<Kind> {
        self.tokens.get(index).map(|token| token.kind)
    }

    /// the `steps`-th token after `index`, not counting whitespace.
    fn after(&self, index: usize, steps: usize) -> Option<usize> {
        let mut at = index;
        for _ in 0..steps {
            at += 1;
            if self.kind(at) == Some(Kind::Space) {
                at += 1;
            }
        }
        (at < self.tokens.len()).then_some(at)
    }

    /// the `steps`-th token before `index`, not counting whitespace.
    fn before(&self, index: usize, steps: usize) -> Option<usize> {
        let mut at = index;
        for _ in 0..steps {
            at = at.checked_sub(1)?;
            if self.kind(at) == Some(Kind::Space) {
                at = at.checked_sub(1)?;
            }
        }
        Some(at)
    }

    fn kind_after(&self, index: usize, steps: usize) -> Option<Kind> {
        self.after(index, steps).and_then(|at| self.kind(at))
    }

    fn kind_before(&self, index: usize, steps: usize) -> Option<Kind> {
        self.before(index, steps).and_then(|at| self.kind(at))
    }

    /// a sheet name before `!` or starting a `First:Last!` span, or a name
    /// scoped to a sheet after it.
    fn qualified(&self, index: usize) -> bool {
        self.kind_after(index, 1) == Some(Kind::Bang)
            || self.kind_before(index, 1) == Some(Kind::Bang)
            || (self.kind_after(index, 1) == Some(Kind::Colon)
                && self.kind_after(index, 2) == Some(Kind::Name)
                && self.kind_after(index, 3) == Some(Kind::Bang))
    }

    /// a table name or R1C1 axis right next to its bracket, or one end of a
    /// whole column range such as `A:A`.
    fn reference_part(&self, index: usize, text: &str) -> bool {
        if self.next_to_bracket(index) {
            return true;
        }
        if !column_shaped(text) {
            return false;
        }
        let column = |at: Option<usize>| {
            at.is_some_and(|at| {
                let token = self.tokens[at];
                token.kind == Kind::Name
                    && self.kind(at + 1) != Some(Kind::Open)
                    && column_shaped(&self.source[token.start..token.end])
            })
        };
        (self.kind_after(index, 1) == Some(Kind::Colon) && column(self.after(index, 2)))
            || (self.kind_before(index, 1) == Some(Kind::Colon) && column(self.before(index, 2)))
    }

    /// a table name or R1C1 axis right next to its bracket.
    fn next_to_bracket(&self, index: usize) -> bool {
        self.kind(index + 1) == Some(Kind::Bracket)
            || index
                .checked_sub(1)
                .is_some_and(|before| self.kind(before) == Some(Kind::Bracket))
    }

    fn is_bound(&self, text: &str) -> bool {
        if self.bound.is_empty() {
            return false;
        }
        let bare = strip_prefix(text, PARAMETER_PREFIX)
            .unwrap_or(text)
            .to_lowercase();
        self.bound.contains_key(&bare)
    }

    fn insert(&mut self, at: usize, text: &'static str) {
        self.edits.push(Edit {
            start: at,
            end: at,
            text,
        });
    }

    /// the top-level argument token ranges inside the parentheses at `open`
    /// and `close`, stepping over nested groups whole.
    fn arguments(&self, open: usize, close: usize) -> Vec<(usize, usize)> {
        let mut arguments = Vec::new();
        let mut start = open + 1;
        let mut index = open + 1;
        while index < close {
            match self.tokens[index].kind {
                Kind::Open | Kind::BraceOpen => {
                    index = self.partners[index]
                        .filter(|&partner| partner < close)
                        .unwrap_or(index);
                }
                Kind::Comma => {
                    arguments.push((start, index));
                    start = index + 1;
                }
                _ => {}
            }
            index += 1;
        }
        arguments.push((start, close));
        arguments
    }

    /// queue the scope changes of one `LET` or `LAMBDA` for the walk. one slot
    /// that is not a valid name leaves every argument a plain expression,
    /// since Excel refuses such a formula.
    fn bind(&mut self, binder: Binder, arguments: &[(usize, usize)], close: usize) {
        let Some((_, names)) = arguments.split_last() else {
            return;
        };
        let (pairs, _) = names.as_chunks::<2>();
        let slots = match binder {
            Binder::Lambda => names
                .iter()
                .map(|&(lo, hi)| self.parameter(lo, hi).map(|slot| (slot, None)))
                .collect::<Option<Vec<_>>>(),
            Binder::Let => pairs
                .iter()
                .map(|&[(lo, hi), (_, value_hi)]| {
                    self.parameter(lo, hi).map(|slot| (slot, Some(value_hi)))
                })
                .collect::<Option<Vec<_>>>(),
        };
        let Some(slots) = slots else {
            return;
        };
        for (slot, visible_at) in slots {
            self.actions[slot] = Action::Declare { visible_at };
        }
        self.actions[close] = Action::Close(self.scopes.len());
    }

    /// the token a binding argument holds, when it is a single name Excel
    /// accepts as a parameter.
    fn parameter(&self, lo: usize, hi: usize) -> Option<usize> {
        let index = self.single(lo, hi)?;
        let token = self.tokens[index];
        let text = &self.source[token.start..token.end];
        parameter_name(strip_prefix(text, PARAMETER_PREFIX).unwrap_or(text)).then_some(index)
    }

    /// the one name token an argument holds, whitespace aside.
    fn single(&self, lo: usize, hi: usize) -> Option<usize> {
        let mut tokens = (lo..hi).filter(|&index| self.tokens[index].kind != Kind::Space);
        let index = tokens.next()?;
        (tokens.next().is_none() && self.tokens[index].kind == Kind::Name).then_some(index)
    }

    /// prefix the name a binder declares at `index` when it is still bare,
    /// and bring it into scope now or once its value ends.
    fn declare(&mut self, index: usize, visible_at: Option<usize>) {
        let token = self.tokens[index];
        let text = &self.source[token.start..token.end];
        let bare = match strip_prefix(text, PARAMETER_PREFIX) {
            Some(rest) => rest,
            None => {
                self.insert(token.start, PARAMETER_PREFIX);
                text
            }
        };
        let name = bare.to_lowercase();
        match visible_at {
            Some(at) => self.actions[at] = Action::Bring(name),
            None => self.bring(name),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[track_caller]
    fn written(source: &str, expected: &str) {
        let once = excel_formula(source);
        assert_eq!(once, expected, "{source}");
        assert!(
            matches!(excel_formula(&once), Cow::Borrowed(_)),
            "not idempotent: {once}"
        );
        let shown = display_formula(&once);
        assert_eq!(excel_formula(&shown), once, "round trip of {once}");
        assert!(
            matches!(display_formula(&shown), Cow::Borrowed(_)),
            "display not idempotent: {shown}"
        );
    }

    #[track_caller]
    fn shown(stored: &str, expected: &str) {
        let display = display_formula(stored);
        assert_eq!(display, expected, "{stored}");
        assert!(
            matches!(display_formula(&display), Cow::Borrowed(_)),
            "display not idempotent: {display}"
        );
        assert_eq!(
            excel_formula(&display),
            stored,
            "{display} does not store as {stored}"
        );
    }

    #[test]
    fn future_function_table_is_sorted_unique_and_upper_case() {
        assert!(
            FUTURE_FUNCTIONS.windows(2).all(|pair| pair[0] < pair[1]),
            "table must be strictly sorted"
        );
        assert!(
            FUTURE_FUNCTIONS
                .iter()
                .all(|name| *name == name.to_ascii_uppercase())
        );
        assert!(
            WORKSHEET_FUNCTIONS
                .iter()
                .all(|name| FUTURE_FUNCTIONS.binary_search(name).is_ok())
        );
        for name in FUTURE_FUNCTIONS {
            assert_eq!(
                future_function(&name.to_ascii_lowercase()).unwrap().0,
                *name
            );
        }
        assert!(LAMBDA_HELPERS.windows(2).all(|pair| pair[0].0 < pair[1].0));
        assert!(
            LAMBDA_HELPERS
                .iter()
                .all(|(name, _)| FUTURE_FUNCTIONS.binary_search(name).is_ok())
        );
        for name in [
            "SUM",
            "IFERROR",
            "COUNTIFS",
            "VLOOKUP",
            "CONCATENATE",
            "SUMIFS",
        ] {
            assert!(future_function(name).is_none(), "{name}");
            assert!(builtin(&name.to_ascii_lowercase()), "{name}");
        }
    }

    #[test]
    fn classic_function_table_is_sorted_unique_and_apart_from_the_future_one() {
        assert!(
            CLASSIC_FUNCTIONS.windows(2).all(|pair| pair[0] < pair[1]),
            "table must be strictly sorted"
        );
        for name in CLASSIC_FUNCTIONS {
            assert_eq!(*name, name.to_ascii_uppercase());
            assert!(future_function(name).is_none(), "{name}");
        }
        for name in ["ROW", "INDEX", "TRUE", "XLOOKUP", "filter"] {
            assert!(builtin(name), "{name}");
        }
        for name in ["f", "MYFUNC", "total"] {
            assert!(!builtin(name), "{name}");
        }
    }

    #[test]
    fn future_functions_take_the_prefix_excel_writes() {
        for (source, expected) in [
            ("XLOOKUP(A1,B1:B3,C1:C3)", "_xlfn.XLOOKUP(A1,B1:B3,C1:C3)"),
            ("IFS(A1>0,1,TRUE,2)", "_xlfn.IFS(A1>0,1,TRUE,2)"),
            ("CONCAT(A1,B1)", "_xlfn.CONCAT(A1,B1)"),
            (
                "TEXTJOIN(\",\",TRUE,A1:A3)",
                "_xlfn.TEXTJOIN(\",\",TRUE,A1:A3)",
            ),
            ("FILTER(A1:A3,B1:B3>0)", "_xlfn._xlws.FILTER(A1:A3,B1:B3>0)"),
            ("SORT(A1:A3)", "_xlfn._xlws.SORT(A1:A3)"),
            ("SEQUENCE(3)", "_xlfn.SEQUENCE(3)"),
            ("UNIQUE(A1:A3)", "_xlfn.UNIQUE(A1:A3)"),
            ("NORM.DIST(1,0,1,TRUE)", "_xlfn.NORM.DIST(1,0,1,TRUE)"),
            ("STDEV.S(A1:A3)", "_xlfn.STDEV.S(A1:A3)"),
            ("T.TEST(A1:A3,B1:B3,2,1)", "_xlfn.T.TEST(A1:A3,B1:B3,2,1)"),
            ("ANCHORARRAY(A1)", "_xlfn.ANCHORARRAY(A1)"),
            ("SINGLE(A1:A3)", "_xlfn.SINGLE(A1:A3)"),
            (
                "SUM(XLOOKUP(1,A1:A3,B1:B3))",
                "SUM(_xlfn.XLOOKUP(1,A1:A3,B1:B3))",
            ),
            (
                "IFERROR(XLOOKUP(1,A1:A3,B1:B3),0)",
                "IFERROR(_xlfn.XLOOKUP(1,A1:A3,B1:B3),0)",
            ),
            ("xlookup(1,a1:a3,b1:b3)", "_xlfn.XLOOKUP(1,a1:a3,b1:b3)"),
            ("Filter(A1:A3,A1:A3>1)", "_xlfn._xlws.FILTER(A1:A3,A1:A3>1)"),
            (
                "SORT(UNIQUE(FILTER(A1:A9,B1:B9=\"x\")))",
                "_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(A1:A9,B1:B9=\"x\")))",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn text_that_only_looks_like_a_call_is_left_alone() {
        for source in [
            "\"XLOOKUP(\"&A1",
            "CONCATENATE(\"IFS(\",A1)",
            "XLOOKUP!A1+SUM(XLOOKUP!B1:B3)",
            "'XLOOKUP(1)'!A1",
            "SUM(Table1[XLOOKUP(x)])",
            "Table1[[#This Row],[SORT(]]",
            "XLOOKUP+1",
            "Sheet1!SEQUENCE",
            "MyXLOOKUP(1)",
            "XLOOKUPX(1)",
            "SUM(A1:A3)",
            "XLOOKUP (1,A1:A3,B1:B3)",
            "",
        ] {
            written(source, source);
        }
    }

    #[test]
    fn prefixed_text_is_unchanged() {
        for source in [
            "_xlfn.XLOOKUP(A1,B1:B3,C1:C3)",
            "_xlfn._xlws.FILTER(A1:A3,B1:B3>0)",
            "_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))",
            "_xlfn.LET(_xlpm.x,1,_xlpm.y,2,_xlpm.x+_xlpm.y)",
            "_xlfn.LAMBDA(_xlpm.a,_xlpm.a*2)(3)",
            "_xlfn.xlookup(1,A1:A3,B1:B3)",
            "IFERROR(_xlfn.XLOOKUP(1,A1:A3,B1:B3),0)",
            "SUM(_xlfn.MAKEARRAY(2,1,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c)))",
            "_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.SUM)",
        ] {
            assert!(
                matches!(excel_formula(source), Cow::Borrowed(_)),
                "{source}"
            );
        }
    }

    #[test]
    fn let_and_lambda_names_take_the_parameter_prefix() {
        for (source, expected) in [
            (
                "LET(x,1,y,x+1,x*y)",
                "_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)",
            ),
            ("LAMBDA(a,a*2)(3)", "_xlfn.LAMBDA(_xlpm.a,_xlpm.a*2)(3)"),
            (
                "LET(f,LAMBDA(v,v+1),f(2))",
                "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v+1),_xlpm.f(2))",
            ),
            ("LET(x,x,x)", "_xlfn.LET(_xlpm.x,x,_xlpm.x)"),
            (
                "LET(x,1,LAMBDA(y,x+y)(2))+x",
                "_xlfn.LET(_xlpm.x,1,_xlfn.LAMBDA(_xlpm.y,_xlpm.x+_xlpm.y)(2))+x",
            ),
            (
                "LET(x,1,LET(x,2,x)+X)",
                "_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.x,2,_xlpm.x)+_xlpm.X)",
            ),
            (
                "LET(total, SUM(A1:A3), rate, 2, total * rate)",
                "_xlfn.LET(_xlpm.total, SUM(A1:A3), _xlpm.rate, 2, _xlpm.total * _xlpm.rate)",
            ),
            (
                "LET(s,{1,2;3,4},SUM(s))",
                "_xlfn.LET(_xlpm.s,{1,2;3,4},SUM(_xlpm.s))",
            ),
            (
                "LET(x,\"x\",x&Sheet1!x)",
                "_xlfn.LET(_xlpm.x,\"x\",_xlpm.x&Sheet1!x)",
            ),
            (
                "MAP(A1:A3,LAMBDA(v,XLOOKUP(v,B1:B3,C1:C3)))",
                "_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlfn.XLOOKUP(_xlpm.v,B1:B3,C1:C3)))",
            ),
            (
                "_xlfn.LET(x,1,_xlpm.x+x)",
                "_xlfn.LET(_xlpm.x,1,_xlpm.x+_xlpm.x)",
            ),
            ("LET(x,1", "_xlfn.LET(x,1"),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn references_that_share_a_bound_name_are_left_alone() {
        for (source, expected) in [
            (
                "LET(a,1,SUM(A:A)+a)",
                "_xlfn.LET(_xlpm.a,1,SUM(A:A)+_xlpm.a)",
            ),
            (
                "LET(a,1,SUM(Data!A:A)+a)",
                "_xlfn.LET(_xlpm.a,1,SUM(Data!A:A)+_xlpm.a)",
            ),
            (
                "LET(x,1,SUM(Data!A1:x)+x)",
                "_xlfn.LET(_xlpm.x,1,SUM(Data!A1:_xlpm.x)+_xlpm.x)",
            ),
            (
                "LET(s,1,SUM(s:Sheet2!A1)+s)",
                "_xlfn.LET(_xlpm.s,1,SUM(s:Sheet2!A1)+_xlpm.s)",
            ),
            (
                "LET(total,1,SUM('My Sheet'!A1:total)+total)",
                "_xlfn.LET(_xlpm.total,1,SUM('My Sheet'!A1:_xlpm.total)+_xlpm.total)",
            ),
            (
                "LET(lo,A1,hi,A3,SUM(lo:hi))",
                "_xlfn.LET(_xlpm.lo,A1,_xlpm.hi,A3,SUM(lo:hi))",
            ),
            (
                "LET(start,A1,finish,A3,SUM(start:finish))",
                "_xlfn.LET(_xlpm.start,A1,_xlpm.finish,A3,SUM(_xlpm.start:_xlpm.finish))",
            ),
            (
                "LET(t,1,SUM(t[Col])+t)",
                "_xlfn.LET(_xlpm.t,1,SUM(t[Col])+_xlpm.t)",
            ),
            (
                "LET(r,1,R[1]C[1]+r)",
                "_xlfn.LET(_xlpm.r,1,R[1]C[1]+_xlpm.r)",
            ),
            (
                "LET(Data,1,Data!A1+Data)",
                "_xlfn.LET(_xlpm.Data,1,Data!A1+_xlpm.Data)",
            ),
            (
                "LET(Rate,2,Rate*Sheet1!Rate)+Rate",
                "_xlfn.LET(_xlpm.Rate,2,_xlpm.Rate*Sheet1!Rate)+Rate",
            ),
            (
                "MAKEARRAY(2,2,LAMBDA(r,c,r+c))",
                "_xlfn.MAKEARRAY(2,2,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c))",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn only_valid_names_bind() {
        for (source, expected) in [
            ("LET(A1,1,A1+B1)", "_xlfn.LET(A1,1,A1+B1)"),
            ("LET(x1,2,x1+1)", "_xlfn.LET(x1,2,x1+1)"),
            ("LET($A$1,2,$A$1)", "_xlfn.LET($A$1,2,$A$1)"),
            ("LET(1,2,3)", "_xlfn.LET(1,2,3)"),
            ("LET(TRUE,2,TRUE)", "_xlfn.LET(TRUE,2,TRUE)"),
            ("LET(R1C1,2,R1C1)", "_xlfn.LET(R1C1,2,R1C1)"),
            ("LET(x,1,A1,2,x+A1)", "_xlfn.LET(x,1,A1,2,x+A1)"),
            ("LAMBDA(a,B2,a+B2)", "_xlfn.LAMBDA(a,B2,a+B2)"),
            (
                "LET(x,1,XLOOKUP(x,A1:A3,B1:B3))",
                "_xlfn.LET(_xlpm.x,1,_xlfn.XLOOKUP(_xlpm.x,A1:A3,B1:B3))",
            ),
            (
                "LET(total1,2,_rate,3,total1*_rate)",
                "_xlfn.LET(_xlpm.total1,2,_xlpm._rate,3,_xlpm.total1*_xlpm._rate)",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn builtin_calls_stay_functions_and_bound_calls_take_the_parameter_prefix() {
        for (source, expected) in [
            (
                "LET(filter,B1:B3>0,FILTER(A1:A3,filter))",
                "_xlfn.LET(_xlpm.filter,B1:B3>0,_xlfn._xlws.FILTER(A1:A3,_xlpm.filter))",
            ),
            (
                "LET(row,1,ROW(A1)+row)",
                "_xlfn.LET(_xlpm.row,1,ROW(A1)+_xlpm.row)",
            ),
            (
                "LET(f,LAMBDA(v,v+1),f(2))",
                "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v+1),_xlpm.f(2))",
            ),
            (
                "LET(g,LAMBDA(n,n*n),g(3)+PI())+g(1)",
                "_xlfn.LET(_xlpm.g,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*_xlpm.n),_xlpm.g(3)+PI())+g(1)",
            ),
            (
                "LET(seq,SEQUENCE,seq(2))",
                "_xlfn.LET(_xlpm.seq,SEQUENCE,_xlpm.seq(2))",
            ),
            (
                "LAMBDA(sum,SUM(sum))(A1:A3)",
                "_xlfn.LAMBDA(_xlpm.sum,SUM(_xlpm.sum))(A1:A3)",
            ),
            (
                "LET(x,1,LET(y,x,y))",
                "_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.y,_xlpm.x,_xlpm.y))",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn range_ends_that_are_calls_or_variables_are_rewritten() {
        for (source, expected) in [
            (
                "SUM(Data!A1:XLOOKUP(2,Data!A1:A3,Data!B1:B3))",
                "SUM(Data!A1:_xlfn.XLOOKUP(2,Data!A1:A3,Data!B1:B3))",
            ),
            (
                "LET(f,LAMBDA(r,r),SUM(Data!A1:f(Data!A3)))",
                "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.r,_xlpm.r),SUM(Data!A1:_xlpm.f(Data!A3)))",
            ),
            (
                "LET(x,A3,SUM(A1:x))",
                "_xlfn.LET(_xlpm.x,A3,SUM(A1:_xlpm.x))",
            ),
            (
                "LET(x,A1,SUM(x:A3))",
                "_xlfn.LET(_xlpm.x,A1,SUM(_xlpm.x:A3))",
            ),
            (
                "LET(x,A3,SUM(INDEX(A1:A3,1):x))",
                "_xlfn.LET(_xlpm.x,A3,SUM(INDEX(A1:A3,1):_xlpm.x))",
            ),
            (
                "LET(x,1,SUM($A:$x)+x)",
                "_xlfn.LET(_xlpm.x,1,SUM($A:$x)+_xlpm.x)",
            ),
            (
                "LET(a,A1,SUM(a:f(1)))",
                "_xlfn.LET(_xlpm.a,A1,SUM(_xlpm.a:f(1)))",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn spaces_around_range_and_sheet_punctuation_keep_references() {
        for (source, expected) in [
            (
                "LET(a,1,SUM(A : A)+a)",
                "_xlfn.LET(_xlpm.a,1,SUM(A : A)+_xlpm.a)",
            ),
            (
                "LET(a,1,SUM(Data!A :A)+a)",
                "_xlfn.LET(_xlpm.a,1,SUM(Data!A :A)+_xlpm.a)",
            ),
            (
                "LET(s,1,SUM(s : Sheet2!A1)+s)",
                "_xlfn.LET(_xlpm.s,1,SUM(s : Sheet2!A1)+_xlpm.s)",
            ),
            (
                "LET(x,A3,SUM(A1 : x))",
                "_xlfn.LET(_xlpm.x,A3,SUM(A1 : _xlpm.x))",
            ),
            (
                "LET(t,A1:A3,SUM(Table1[Col] t))",
                "_xlfn.LET(_xlpm.t,A1:A3,SUM(Table1[Col] _xlpm.t))",
            ),
            (
                "LET(\n  a, 1,\n  SUM(A\t:\tA) + a\n)",
                "_xlfn.LET(\n  _xlpm.a, 1,\n  SUM(A\t:\tA) + _xlpm.a\n)",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn deep_nesting_is_rewritten_without_recursion() {
        for (depth, fits) in [(1_489, true), (200_000, false)] {
            let source = format!("{}a{}", "LAMBDA(a,".repeat(depth), ")".repeat(depth));
            let expected = format!(
                "{}_xlpm.a{}",
                "_xlfn.LAMBDA(_xlpm.a,".repeat(depth),
                ")".repeat(depth)
            );
            assert_eq!(expected.len() <= MAX_FORMULA_BYTES, fits);
            written(&source, if fits { &expected } else { &source });
        }
        for (depth, fits) in [(1_000, true), (200_000, false)] {
            let source = format!("{}1{}", "SUM(LET(x,".repeat(depth), ",x))".repeat(depth));
            let expected = format!(
                "{}1{}",
                "SUM(_xlfn.LET(_xlpm.x,".repeat(depth),
                ",_xlpm.x))".repeat(depth)
            );
            assert_eq!(expected.len() <= MAX_FORMULA_BYTES, fits);
            written(&source, if fits { &expected } else { &source });
        }
        let unbalanced = "LET(x,(".repeat(2_000);
        assert_eq!(excel_formula(&unbalanced), "_xlfn.LET(x,(".repeat(2_000));
        let unbalanced = "LET(x,(".repeat(200_000);
        assert!(matches!(excel_formula(&unbalanced), Cow::Borrowed(_)));
    }

    #[test]
    fn long_formulas_with_many_tokens_and_names_are_rewritten() {
        let call = "XLOOKUP(1,A1:A3,B1:B3)+";
        let count = (MAX_FORMULA_BYTES - 1) / (call.len() + FUNCTION_PREFIX.len());
        let source = format!("{}0", call.repeat(count));
        let expected = format!("{}0", "_xlfn.XLOOKUP(1,A1:A3,B1:B3)+".repeat(count));
        assert!(expected.len() <= MAX_FORMULA_BYTES);
        written(&source, &expected);

        let names = 800;
        let bindings = (0..names).map(|index| format!("name{index},{index},"));
        let uses = (0..names)
            .map(|index| format!("name{index}"))
            .collect::<Vec<_>>();
        let source = format!("LET({}{})", bindings.collect::<String>(), uses.join("+"));
        let bindings = (0..names).map(|index| format!("_xlpm.name{index},{index},"));
        let uses = (0..names)
            .map(|index| format!("_xlpm.name{index}"))
            .collect::<Vec<_>>();
        let expected = format!(
            "_xlfn.LET({}{})",
            bindings.collect::<String>(),
            uses.join("+")
        );
        assert!(expected.len() <= MAX_FORMULA_BYTES);
        written(&source, &expected);

        let array = format!("LET(x,{{{}}},x)", vec!["1,2"; 6_000].join(";"));
        let expected = format!(
            "_xlfn.LET(_xlpm.x,{{{}}},_xlpm.x)",
            vec!["1,2"; 6_000].join(";")
        );
        assert!(expected.len() <= MAX_FORMULA_BYTES);
        written(&array, &expected);
    }

    #[test]
    fn a_rewrite_past_the_formula_limit_keeps_the_text() {
        let call = "IFS(1,2)+";
        let fits = format!(
            "{}0",
            call.repeat((MAX_FORMULA_BYTES - 1) / (call.len() + 6))
        );
        assert!(matches!(excel_formula(&fits), Cow::Owned(_)));
        let over = format!("{}0", call.repeat((MAX_FORMULA_BYTES - 1) / call.len()));
        assert!(over.len() <= MAX_FORMULA_BYTES);
        assert!(matches!(excel_formula(&over), Cow::Borrowed(_)));
    }

    #[test]
    fn functions_passed_by_name_to_a_lambda_helper_take_the_eta_prefix() {
        for (source, expected) in [
            (
                "GROUPBY(A2:A9,B2:B9,SUM)",
                "_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.SUM)",
            ),
            ("BYROW(A1:B2, sum)", "_xlfn.BYROW(A1:B2, _xleta.SUM)"),
            (
                "PIVOTBY(A2:A9,B2:B9,C2:C9,AVERAGE,FALSE)",
                "_xlfn.PIVOTBY(A2:A9,B2:B9,C2:C9,_xleta.AVERAGE,FALSE)",
            ),
            (
                "LET(sum,LAMBDA(x,x),MAP(A1:A3,sum))",
                "_xlfn.LET(_xlpm.sum,_xlfn.LAMBDA(_xlpm.x,_xlpm.x),_xlfn.MAP(A1:A3,_xlpm.sum))",
            ),
            ("MAP(A1:A3,Total)", "_xlfn.MAP(A1:A3,Total)"),
            ("HSTACK(SUM,AVERAGE)", "_xlfn.HSTACK(SUM,AVERAGE)"),
            ("SUM(A1,ABS)", "SUM(A1,ABS)"),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn display_drops_the_storage_prefixes_of_names() {
        for (stored, expected) in [
            ("_xlfn.SEQUENCE(2,2)", "SEQUENCE(2,2)"),
            ("_xlfn.XLOOKUP(A1,B1:B3,C1:C3)", "XLOOKUP(A1,B1:B3,C1:C3)"),
            ("_xlfn._xlws.FILTER(A1:A3,B1:B3>0)", "FILTER(A1:A3,B1:B3>0)"),
            (
                "_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))",
                "SORT(UNIQUE(A1:A3))",
            ),
            (
                "_xlfn.LET(_xlpm.x,1,_xlpm.y,2,_xlpm.x+_xlpm.y)",
                "LET(x,1,y,2,x+y)",
            ),
            ("_xlfn.LAMBDA(_xlpm.a,_xlpm.a*2)(3)", "LAMBDA(a,a*2)(3)"),
            (
                "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.v,_xlpm.v+1),_xlpm.f(2))",
                "LET(f,LAMBDA(v,v+1),f(2))",
            ),
            (
                "SUM(_xlfn.MAKEARRAY(2,1,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c)))",
                "SUM(MAKEARRAY(2,1,LAMBDA(r,c,r+c)))",
            ),
            (
                "_xlfn.GROUPBY(A2:A9,B2:B9,_xleta.SUM)",
                "GROUPBY(A2:A9,B2:B9,SUM)",
            ),
            ("_xlfn.ANCHORARRAY(A1)", "ANCHORARRAY(A1)"),
            ("_xlfn.SINGLE(A1:A3)", "SINGLE(A1:A3)"),
            (
                "IFERROR(_xlfn.XLOOKUP(1,A1:A3,B1:B3),\"_xlfn.\")",
                "IFERROR(XLOOKUP(1,A1:A3,B1:B3),\"_xlfn.\")",
            ),
        ] {
            shown(stored, expected);
        }
    }

    #[test]
    fn display_leaves_literals_sheets_and_tables_alone() {
        for stored in [
            "\"_xlfn.SEQUENCE(2)\"&A1",
            "'_xlfn.Data'!A1",
            "_xlfn.Data!A1+SUM(_xlfn.Data!B1:B3)",
            "SUM(Table1[_xlfn.Col])",
            "SUM(_xlfn.T[Col])",
            "SUM(A1:A3)",
            "_xlpm.x+1",
            "",
        ] {
            assert!(
                matches!(display_formula(stored), Cow::Borrowed(_)),
                "{stored}"
            );
        }
    }

    #[test]
    fn display_keeps_a_prefix_saving_would_not_restore() {
        shown(
            "_xlfn.GROUPBY(A2:A9,B2:B9,_xlfn.HSTACK(_xleta.SUM,_xleta.AVERAGE))",
            "GROUPBY(A2:A9,B2:B9,HSTACK(_xleta.SUM,_xleta.AVERAGE))",
        );
    }

    #[test]
    fn display_of_deep_and_long_formulas_is_linear_and_bounded() {
        for depth in [MAX_FORMULA_BYTES / 30, MAX_FORMULA_BYTES / 10] {
            let stored = format!(
                "{}_xlpm.x{}",
                "_xlfn.LET(_xlpm.x,1,".repeat(depth),
                ")".repeat(depth)
            );
            let shown = format!("{}x{}", "LET(x,1,".repeat(depth), ")".repeat(depth));
            if stored.len() <= MAX_FORMULA_BYTES {
                assert_eq!(display_formula(&stored), shown);
            } else {
                assert!(matches!(display_formula(&stored), Cow::Borrowed(_)));
            }
        }
        let calls = "_xlfn.SEQUENCE(1)+".repeat(MAX_FORMULA_BYTES / 20);
        let stored = format!("{calls}0");
        assert_eq!(
            display_formula(&stored),
            format!("{}0", "SEQUENCE(1)+".repeat(MAX_FORMULA_BYTES / 20))
        );
        let prefixes = format!("{}SUM(1)", "_xlfn.".repeat(MAX_FORMULA_BYTES / 8));
        assert!(matches!(display_formula(&prefixes), Cow::Borrowed(_)));
    }

    #[test]
    fn only_the_callback_slot_of_a_lambda_helper_takes_the_eta_prefix() {
        for (source, expected) in [
            (
                "MAKEARRAY(rows,columns,LAMBDA(r,c,r*c))",
                "_xlfn.MAKEARRAY(rows,columns,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*_xlpm.c))",
            ),
            (
                "BYROW(value,LAMBDA(r,SUM(r)))",
                "_xlfn.BYROW(value,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))",
            ),
            (
                "REDUCE(0,value,LAMBDA(a,v,a+v))",
                "_xlfn.REDUCE(0,value,_xlfn.LAMBDA(_xlpm.a,_xlpm.v,_xlpm.a+_xlpm.v))",
            ),
            ("BYROW(A1:B2,SUM)", "_xlfn.BYROW(A1:B2,_xleta.SUM)"),
            ("BYCOL(rate,MAX)", "_xlfn.BYCOL(rate,_xleta.MAX)"),
            ("MAP(N,T,year,CONCAT)", "_xlfn.MAP(N,T,year,_xleta.CONCAT)"),
            ("MAP(value)", "_xlfn.MAP(value)"),
            (
                "MAKEARRAY(rows,columns,SUM)",
                "_xlfn.MAKEARRAY(rows,columns,_xleta.SUM)",
            ),
            (
                "REDUCE(value,rate,SUM)",
                "_xlfn.REDUCE(value,rate,_xleta.SUM)",
            ),
            ("SCAN(,year,SUM)", "_xlfn.SCAN(,year,_xleta.SUM)"),
            (
                "GROUPBY(year,value,SUM,rate,T)",
                "_xlfn.GROUPBY(year,value,_xleta.SUM,rate,T)",
            ),
            (
                "PIVOTBY(year,rows,value,AVERAGE,N)",
                "_xlfn.PIVOTBY(year,rows,value,_xleta.AVERAGE,N)",
            ),
        ] {
            written(source, expected);
        }
    }

    #[test]
    fn display_shows_stripped_text_only_when_it_saves_as_the_stored_text() {
        for stored in [
            "_xlfn.xlookup(1,A1:A3,B1:B3)",
            "_XLFN.SEQUENCE(2,2)",
            "_xlfn.BYROW(A1:B2,_xleta.sum)",
            "_xlfn.MAP(A1:A3,SUM)",
            "_xlfn.BYROW(value,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))",
            "_xlfn.MAKEARRAY(rows,columns,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*_xlpm.c))",
            "_xlfn.LET(_xlpm.x,1,SEQUENCE(_xlpm.x))",
            "_xlfn.REDUCE(0,value,_xlfn.LAMBDA(_xlpm.a,_xlpm.v,_xlpm.a+_xlpm.v))",
        ] {
            let display = display_formula(stored);
            assert!(
                display == stored || excel_formula(&display) == stored,
                "{stored} shown as {display}"
            );
            assert!(
                matches!(display_formula(&display), Cow::Borrowed(_)),
                "display not idempotent: {display}"
            );
        }
        for stored in [
            "_xlfn.xlookup(1,A1:A3,B1:B3)",
            "_XLFN.SEQUENCE(2,2)",
            "_xlfn.MAP(A1:A3,SUM)",
        ] {
            assert_eq!(display_formula(stored), stored);
        }
        shown("_xlfn.BYROW(A1:B2,_xleta.sum)", "BYROW(A1:B2,_xleta.sum)");
        shown(
            "_xlfn.BYROW(value,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))",
            "BYROW(value,LAMBDA(r,SUM(r)))",
        );
        shown(
            "_xlfn.MAKEARRAY(rows,columns,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*_xlpm.c))",
            "MAKEARRAY(rows,columns,LAMBDA(r,c,r*c))",
        );
    }

    #[test]
    fn interleaved_brackets_pair_on_one_stack_and_stay_linear() {
        let source = format!(
            "{}}}{}{}{}",
            "MAP({".repeat(700),
            ",a".repeat(1500),
            "}".repeat(699),
            ")".repeat(700)
        );
        assert!(source.len() <= MAX_FORMULA_BYTES);
        let tokens = tokenize(&source);
        let pass = Pass::new(&source, &tokens);
        let mut opener = vec![None; tokens.len()];
        for (open, close) in pass.partners.iter().enumerate() {
            if let Some(close) = *close {
                assert!(matches!(
                    (tokens[open].kind, tokens[close].kind),
                    (Kind::Open, Kind::Close) | (Kind::BraceOpen, Kind::BraceClose)
                ));
                opener[close] = Some(open);
            }
        }
        let mut open = Vec::new();
        for (index, (partner, start)) in pass.partners.iter().zip(&opener).enumerate() {
            if partner.is_some() {
                open.push(index);
            }
            if let Some(start) = *start {
                assert_eq!(
                    open.pop(),
                    Some(start),
                    "pair {start}..{index} crosses another"
                );
            }
        }

        let started = std::time::Instant::now();
        let stored = excel_formula(&source).into_owned();
        assert_eq!(stored, source.replace("MAP(", "_xlfn.MAP("));
        assert_eq!(display_formula(&stored), source);
        assert!(started.elapsed() < std::time::Duration::from_secs(5));
    }

    #[test]
    fn text_over_the_formula_limit_is_returned_unchanged() {
        let calls = MAX_FORMULA_BYTES / "_xlfn.SEQUENCE(1)+".len() + 1;
        let stored = format!("{}0", "_xlfn.SEQUENCE(1)+".repeat(calls));
        assert!(stored.len() > MAX_FORMULA_BYTES);
        assert!(matches!(display_formula(&stored), Cow::Borrowed(_)));
        let typed = format!("{}0", "SEQUENCE(1)+".repeat(MAX_FORMULA_BYTES / 12 + 1));
        assert!(typed.len() > MAX_FORMULA_BYTES);
        assert!(matches!(excel_formula(&typed), Cow::Borrowed(_)));
        let huge = "_xlfn.SEQUENCE(1)+".repeat(200_000);
        assert!(matches!(display_formula(&huge), Cow::Borrowed(_)));
        assert!(matches!(excel_formula(&huge), Cow::Borrowed(_)));
    }
}
