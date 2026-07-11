//! excel criteria strings shared by the *IF / *IFS family: optional comparison
//! operator prefix, numeric or case-insensitive text compare, `*`/`?`/`~` wildcards.

use xlsx_model::CellValue;

use crate::eval::{Area, EvalContext, as_area, evaluate, parse_num};
use crate::parser::Expr;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Op {
    Eq,
    Ne,
    Gt,
    Ge,
    Lt,
    Le,
}

/// a parsed criterion ready to test cell values against.
#[derive(Debug, Clone)]
pub(crate) struct Criterion {
    op: Op,
    /// numeric target if the value parsed as a number.
    number: Option<f64>,
    /// the raw comparison text (lowercased), used for text comparisons.
    text: String,
    /// whether the text carries `*`/`?` wildcards (only meaningful for =/<>).
    wildcard: bool,
}

impl Criterion {
    /// build a criterion from an already-evaluated cell value.
    pub(crate) fn from_value(v: &CellValue) -> Criterion {
        match v {
            CellValue::Number { value } => Criterion {
                op: Op::Eq,
                number: Some(*value),
                text: String::new(),
                wildcard: false,
            },
            CellValue::Bool { value } => Criterion::parse(if *value { "TRUE" } else { "FALSE" }),
            CellValue::Text { value } => Criterion::parse(value),
            _ => Criterion::parse(""),
        }
    }

    /// parse a criteria string such as `">=5"`, `"<>x"`, `"5"`, `"a*"`.
    pub(crate) fn parse(raw: &str) -> Criterion {
        let (op, rest) = split_op(raw);
        let number = parse_num(rest);
        let wildcard = has_wildcard(rest);
        Criterion {
            op,
            number,
            text: rest.to_lowercase(),
            wildcard,
        }
    }

    /// does a cell value satisfy this criterion?
    pub(crate) fn matches(&self, v: &CellValue) -> bool {
        match self.op {
            Op::Eq => self.eq_matches(v),
            Op::Ne => !self.eq_matches(v),
            Op::Gt | Op::Ge | Op::Lt | Op::Le => self.ord_matches(v),
        }
    }

    fn eq_matches(&self, v: &CellValue) -> bool {
        if let Some(target) = self.number {
            return cell_number(v).map(|n| n == target).unwrap_or(false);
        }
        if self.text.is_empty() {
            return matches!(v, CellValue::Empty);
        }
        let cell_text = cell_text_lower(v);
        if self.wildcard {
            wildcard_match(&self.text, &cell_text)
        } else {
            cell_text == self.text
        }
    }

    fn ord_matches(&self, v: &CellValue) -> bool {
        use std::cmp::Ordering;
        let ord = if let Some(target) = self.number {
            match cell_number(v) {
                Some(n) => n.partial_cmp(&target),
                None => return false,
            }
        } else {
            // numeric cells never satisfy a text inequality in excel
            match v {
                CellValue::Text { value } => Some(value.to_lowercase().cmp(&self.text)),
                CellValue::Empty if self.text.is_empty() => Some(Ordering::Equal),
                _ => return false,
            }
        };
        match ord {
            Some(o) => match self.op {
                Op::Gt => o == Ordering::Greater,
                Op::Ge => o != Ordering::Less,
                Op::Lt => o == Ordering::Less,
                Op::Le => o != Ordering::Greater,
                _ => unreachable!(),
            },
            None => false,
        }
    }
}

fn split_op(s: &str) -> (Op, &str) {
    if let Some(rest) = s.strip_prefix(">=") {
        (Op::Ge, rest)
    } else if let Some(rest) = s.strip_prefix("<=") {
        (Op::Le, rest)
    } else if let Some(rest) = s.strip_prefix("<>") {
        (Op::Ne, rest)
    } else if let Some(rest) = s.strip_prefix('>') {
        (Op::Gt, rest)
    } else if let Some(rest) = s.strip_prefix('<') {
        (Op::Lt, rest)
    } else if let Some(rest) = s.strip_prefix('=') {
        (Op::Eq, rest)
    } else {
        (Op::Eq, s)
    }
}

fn cell_number(v: &CellValue) -> Option<f64> {
    match v {
        CellValue::Number { value } => Some(*value),
        CellValue::Bool { value } => Some(if *value { 1.0 } else { 0.0 }),
        CellValue::Text { value } => parse_num(value),
        _ => None,
    }
}

fn cell_text_lower(v: &CellValue) -> String {
    match v {
        CellValue::Text { value } => value.to_lowercase(),
        CellValue::Bool { value } => if *value { "true" } else { "false" }.to_string(),
        CellValue::Number { value } => crate::eval::format_number(*value),
        _ => String::new(),
    }
}

/// any metacharacter routes through the glob matcher — even fully-escaped
/// patterns, since `~` still has to be unescaped.
fn has_wildcard(s: &str) -> bool {
    s.contains(['*', '?', '~'])
}

/// glob match with `*` and `?`; `~` escapes the following metacharacter.
/// pattern and text are already lowercased.
pub(crate) fn wildcard_match(pattern: &str, text: &str) -> bool {
    let p: Vec<char> = pattern.chars().collect();
    let t: Vec<char> = text.chars().collect();
    let (mut pi, mut ti) = (0usize, 0usize);
    let (mut star_p, mut star_t): (Option<usize>, usize) = (None, 0);
    while ti < t.len() {
        let (lit, is_star, is_any) = classify(&p, pi);
        if is_star {
            star_p = Some(pi);
            star_t = ti;
            pi += 1;
        } else if pi < p.len() && (is_any || lit == Some(t[ti])) {
            pi += advance(&p, pi);
            ti += 1;
        } else if let Some(sp) = star_p {
            pi = sp + 1;
            star_t += 1;
            ti = star_t;
        } else {
            return false;
        }
    }
    while pi < p.len() {
        let (_, is_star, _) = classify(&p, pi);
        if !is_star {
            return false;
        }
        pi += 1;
    }
    true
}

/// interpret the pattern element at `pi`: (literal char, is `*`, is `?`).
fn classify(p: &[char], pi: usize) -> (Option<char>, bool, bool) {
    match p.get(pi) {
        Some('~') => (p.get(pi + 1).copied(), false, false),
        Some('*') => (None, true, false),
        Some('?') => (None, false, true),
        Some(&c) => (Some(c), false, false),
        None => (None, false, false),
    }
}

/// how many pattern chars a single match consumes (2 for an escape `~x`).
fn advance(p: &[char], pi: usize) -> usize {
    if p.get(pi) == Some(&'~') { 2 } else { 1 }
}

/// build a criterion from a criteria argument's evaluated value.
pub(crate) fn criterion_from_arg(arg: &Expr, ctx: &EvalContext<'_>) -> Criterion {
    Criterion::from_value(&evaluate(arg, ctx))
}

/// parse a `range, criteria, ...` tail into aligned (area, criterion) pairs;
/// `None` on an odd count, a non-reference range, or a shape mismatch.
pub(crate) fn collect_pairs(
    specs: &[Expr],
    ctx: &EvalContext<'_>,
) -> Option<Vec<(Area, Criterion)>> {
    if specs.is_empty() || !specs.len().is_multiple_of(2) {
        return None;
    }
    let mut pairs = Vec::new();
    let mut dims: Option<(usize, usize)> = None;
    for chunk in specs.chunks(2) {
        let area = as_area(&chunk[0], ctx)?;
        match dims {
            Some(d) if d != (area.rows, area.cols) => return None,
            _ => dims = Some((area.rows, area.cols)),
        }
        pairs.push((area, criterion_from_arg(&chunk[1], ctx)));
    }
    Some(pairs)
}

/// flat row-major indices where every criterion matches its aligned cell.
pub(crate) fn matching_indices(pairs: &[(Area, Criterion)], ctx: &EvalContext<'_>) -> Vec<usize> {
    let (rows, cols) = match pairs.first() {
        Some((a, _)) => (a.rows, a.cols),
        None => return Vec::new(),
    };
    (0..rows * cols)
        .filter(|&i| {
            let (r, c) = (i / cols, i % cols);
            pairs
                .iter()
                .all(|(area, crit)| crit.matches(&area.get(ctx, r, c)))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn num(v: f64) -> CellValue {
        CellValue::Number { value: v }
    }
    fn txt(v: &str) -> CellValue {
        CellValue::Text { value: v.into() }
    }

    #[test]
    fn numeric_operators() {
        let c = Criterion::parse(">=5");
        assert!(c.matches(&num(5.0)));
        assert!(c.matches(&num(9.0)));
        assert!(!c.matches(&num(4.0)));
        assert!(!c.matches(&txt("x")));

        let c = Criterion::parse("<>0");
        assert!(c.matches(&num(3.0)));
        assert!(!c.matches(&num(0.0)));

        assert!(Criterion::parse("5").matches(&num(5.0)));
        assert!(Criterion::parse("5").matches(&txt("5")));
    }

    #[test]
    fn text_and_wildcards() {
        assert!(Criterion::parse("apple").matches(&txt("Apple")));
        assert!(!Criterion::parse("apple").matches(&txt("apples")));
        assert!(Criterion::parse("a*").matches(&txt("Apple")));
        assert!(Criterion::parse("a?ple").matches(&txt("apple")));
        assert!(!Criterion::parse("a?ple").matches(&txt("aple")));
        assert!(Criterion::parse("<>apple").matches(&txt("pear")));
        assert!(Criterion::parse("a~*").matches(&txt("a*")));
        assert!(!Criterion::parse("a~*").matches(&txt("ab")));
    }

    #[test]
    fn blank_criteria() {
        assert!(Criterion::parse("").matches(&CellValue::Empty));
        assert!(!Criterion::parse("").matches(&txt("x")));
    }

    #[test]
    fn wildcard_matcher_direct() {
        assert!(wildcard_match("*", "anything"));
        assert!(wildcard_match("h*o", "hello"));
        assert!(wildcard_match("h*llo", "hello"));
        assert!(!wildcard_match("h*x", "hello"));
        assert!(wildcard_match("a*b*c", "aXXbYYc"));
        assert!(wildcard_match("", ""));
        assert!(!wildcard_match("", "x"));
    }
}
