//! The functions Excel stores under a version prefix, and the save-time pass
//! that writes formula text the way Excel does: `_xlfn.` before every function
//! newer than Excel 2007, `_xlfn._xlws.` before the worksheet ones, and
//! `_xlpm.` before the names a `LET` or `LAMBDA` binds.

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

/// `formula` as Excel stores it. Text already in that form comes back
/// borrowed, so the pass is idempotent, and so does text the prefixes would
/// push past the formula length limit, which this crate could not read back.
pub(crate) fn excel_formula(formula: &str) -> Cow<'_, str> {
    if !formula.contains('(') {
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
    if pass.edits.is_empty() || formula.len().saturating_add(grown) > MAX_FORMULA_BYTES {
        return Cow::Borrowed(formula);
    }
    let mut out = String::with_capacity(formula.len() + 8 * pass.edits.len());
    let mut cursor = 0;
    for edit in &pass.edits {
        out.push_str(&formula[cursor..edit.start]);
        out.push_str(edit.text);
        cursor = edit.end;
    }
    out.push_str(&formula[cursor..]);
    Cow::Owned(out)
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
    future_function(name).is_some()
        || CLASSIC_FUNCTIONS
            .binary_search_by(|probe| compare_upper(probe, name))
            .is_ok()
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
    /// the closing token of every opening parenthesis or brace.
    partners: Vec<Option<usize>>,
    actions: Vec<Action>,
    /// the lower-cased names bound by every enclosing `LET` and `LAMBDA`,
    /// innermost last, with how often each is bound.
    scopes: Vec<String>,
    bound: HashMap<String, usize>,
    edits: Vec<Edit>,
}

impl<'a> Pass<'a> {
    fn new(source: &'a str, tokens: &'a [Token]) -> Self {
        let mut partners = vec![None; tokens.len()];
        let mut parens = Vec::new();
        let mut braces = Vec::new();
        for (index, token) in tokens.iter().enumerate() {
            let (stack, closing) = match token.kind {
                Kind::Open => (&mut parens, false),
                Kind::BraceOpen => (&mut braces, false),
                Kind::Close => (&mut parens, true),
                Kind::BraceClose => (&mut braces, true),
                _ => continue,
            };
            if !closing {
                stack.push(index);
            } else if let Some(opening) = stack.pop() {
                partners[opening] = Some(index);
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
            edits: Vec::new(),
        }
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
        if self.kind(index + 1) == Some(Kind::Bracket)
            || index
                .checked_sub(1)
                .is_some_and(|before| self.kind(before) == Some(Kind::Bracket))
        {
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
        let mut names = (lo..hi).filter(|&index| self.tokens[index].kind != Kind::Space);
        let index = names.next()?;
        let token = self.tokens[index];
        if names.next().is_some() || token.kind != Kind::Name {
            return None;
        }
        let text = &self.source[token.start..token.end];
        parameter_name(strip_prefix(text, PARAMETER_PREFIX).unwrap_or(text)).then_some(index)
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
}
