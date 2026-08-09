//! Bounded baseline ShapeSheet evaluation. Unsupported formulas never use cached values.
//!
//! See the crate README for the supported profile and explicit non-goals.

mod colour;
#[path = "tests.rs"]
mod corpus;
mod eval;
mod policy;
pub use policy::{MutationContext, MutationOutcome, decide as decide_mutation};
use vsdx_formula::unit;
pub use vsdx_formula::{Expr, Op, Unit};

use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_drawingml::{ColorValue, Theme, get_theme_color, resolve_color_value_to_hex_with_theme};
use thiserror::Error;
use vsdx_parse::{ParseLimits, VsdxPackage};
use vsdx_resolve::{Lookup, ResolveError, ResolvedShape, Resolver};

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Number {
    pub number: f64,
    pub unit: Unit,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Color {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
    pub alpha: Option<u8>,
}
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Value {
    Number(Number),
    Color(Color),
}
#[derive(Clone, Debug, PartialEq)]
pub struct Evaluated {
    pub value: Value,
    pub guarded: bool,
}
#[derive(Clone, Debug, PartialEq)]
pub enum Evaluation {
    Evaluated(Evaluated),
    Unsupported(String),
    Error(Diagnostic),
}
#[derive(Clone, Debug, Error, PartialEq, Eq)]
#[error("{message}")]
pub struct Diagnostic {
    pub message: String,
}

/// A reference provider normally backed by a `vsdx_resolve::ResolvedShape` so inheritance
/// is applied before formulas are evaluated.
pub trait References {
    fn formula(&self, name: &str) -> Option<&str>;
    fn exhausted_inheritance(&self, _name: &str) -> bool {
        false
    }
    fn value(&self, _name: &str) -> Option<(&str, Option<&str>)> {
        None
    }
    fn formula_in(&self, _sheet: Option<u32>, name: &str) -> Option<&str> {
        self.formula(name)
    }
    fn formula_in_scoped(
        &self,
        sheet: Option<u32>,
        scope: Option<&str>,
        name: &str,
    ) -> Option<&str> {
        let name = scope.map_or_else(|| name.into(), |scope| format!("{scope}!{name}"));
        self.formula_in(sheet, &name)
    }
    fn value_in(&self, _sheet: Option<u32>, name: &str) -> Option<(&str, Option<&str>)> {
        self.value(name)
    }
    fn value_in_scoped(
        &self,
        sheet: Option<u32>,
        scope: Option<&str>,
        name: &str,
    ) -> Option<(&str, Option<&str>)> {
        let name = scope.map_or_else(|| name.into(), |scope| format!("{scope}!{name}"));
        self.value_in(sheet, &name)
    }
    fn reference_key(&self, sheet: Option<u32>, name: &str) -> String {
        sheet.map_or_else(|| name.into(), |sheet| format!("Sheet.{sheet}!{name}"))
    }
}
impl References for ResolvedShape {
    fn formula(&self, name: &str) -> Option<&str> {
        match self.cell(name) {
            Some(Lookup::Found(value)) => value
                .cell
                .formula
                .as_deref()
                .filter(|formula| !formula.eq_ignore_ascii_case("Inh")),
            _ => None,
        }
    }
    fn value(&self, name: &str) -> Option<(&str, Option<&str>)> {
        match self.cell(name) {
            Some(Lookup::Found(cell)) => cell
                .cell
                .value
                .as_deref()
                .map(|value| (value, cell.cell.unit.as_deref())),
            _ => None,
        }
    }
    fn exhausted_inheritance(&self, name: &str) -> bool {
        self.cell(name)
            .and_then(|lookup| match lookup {
                Lookup::Found(value) => value.cell.formula.as_deref(),
                Lookup::Deleted | Lookup::Absent => None,
            })
            .is_some_and(|formula| formula.eq_ignore_ascii_case("Inh"))
    }
}

pub struct DocumentReferences<'a, R> {
    references: &'a R,
    document: Option<&'a ResolvedShape>,
}
impl<'a, R> DocumentReferences<'a, R> {
    pub fn new(references: &'a R, document: Option<&'a ResolvedShape>) -> Self {
        Self {
            references,
            document,
        }
    }
    fn document_name(name: &str) -> Option<&str> {
        name.strip_prefix("TheDoc!")
    }
}
impl<R: References> References for DocumentReferences<'_, R> {
    fn formula(&self, name: &str) -> Option<&str> {
        match Self::document_name(name) {
            Some(name) => self.document?.formula(name),
            None => self.references.formula(name),
        }
    }
    fn exhausted_inheritance(&self, name: &str) -> bool {
        match Self::document_name(name) {
            Some(name) => self
                .document
                .is_some_and(|document| document.exhausted_inheritance(name)),
            None => self.references.exhausted_inheritance(name),
        }
    }
    fn value(&self, name: &str) -> Option<(&str, Option<&str>)> {
        match Self::document_name(name) {
            Some(name) => self.document?.value(name),
            None => self.references.value(name),
        }
    }
    fn formula_in(&self, sheet: Option<u32>, name: &str) -> Option<&str> {
        match Self::document_name(name) {
            Some(name) => self.document?.formula_in(sheet, name),
            None => self.references.formula_in(sheet, name),
        }
    }
    fn formula_in_scoped(
        &self,
        sheet: Option<u32>,
        scope: Option<&str>,
        name: &str,
    ) -> Option<&str> {
        if scope == Some("TheDoc") {
            self.document?.formula_in(sheet, name)
        } else {
            self.references.formula_in_scoped(sheet, scope, name)
        }
    }
    fn value_in(&self, sheet: Option<u32>, name: &str) -> Option<(&str, Option<&str>)> {
        match Self::document_name(name) {
            Some(name) => self.document?.value_in(sheet, name),
            None => self.references.value_in(sheet, name),
        }
    }
    fn value_in_scoped(
        &self,
        sheet: Option<u32>,
        scope: Option<&str>,
        name: &str,
    ) -> Option<(&str, Option<&str>)> {
        if scope == Some("TheDoc") {
            self.document?.value_in(sheet, name)
        } else {
            self.references.value_in_scoped(sheet, scope, name)
        }
    }
    fn reference_key(&self, sheet: Option<u32>, name: &str) -> String {
        if name.starts_with("TheDoc!") {
            name.into()
        } else {
            self.references.reference_key(sheet, name)
        }
    }
}

pub struct PageShapeReferences {
    shapes: BTreeMap<u32, ResolvedShape>,
    page: ResolvedShape,
    document: Option<ResolvedShape>,
}
impl PageShapeReferences {
    pub fn new(resolver: &Resolver<'_>, page: &str) -> Result<Self, ResolveError> {
        let page_sheet = resolver
            .package()
            .page_part_ids
            .get(page)
            .and_then(|id| resolver.package().page_sheets.get(id))
            .ok_or_else(|| ResolveError::MissingPage(page.into()))?;
        Ok(Self {
            shapes: resolver.resolve_page_shapes(page)?,
            page: resolver.resolve_sheet(page_sheet)?,
            document: resolver
                .package()
                .document_sheet
                .as_ref()
                .map(|sheet| resolver.resolve_sheet(sheet))
                .transpose()?,
        })
    }
    pub fn for_shape(&self, current: u32) -> ShapeReferences<'_> {
        ShapeReferences {
            current,
            shapes: &self.shapes,
            page: &self.page,
            document: self.document.as_ref(),
        }
    }
    pub fn shape(&self, id: u32) -> Option<&ResolvedShape> {
        self.shapes.get(&id)
    }
    pub fn shapes(&self) -> &BTreeMap<u32, ResolvedShape> {
        &self.shapes
    }
}

pub struct ShapeReferences<'a> {
    current: u32,
    shapes: &'a BTreeMap<u32, ResolvedShape>,
    page: &'a ResolvedShape,
    document: Option<&'a ResolvedShape>,
}
impl ShapeReferences<'_> {
    fn target(&self, sheet: Option<u32>, name: &str) -> Option<&ResolvedShape> {
        if let Some((sheet, _)) = name
            .strip_prefix("Sheet.")
            .and_then(|name| name.split_once('!'))
        {
            return self.shapes.get(&sheet.parse().ok()?);
        }
        if let Some((scope, _)) = name.split_once('!') {
            let sheet = match scope {
                "ThePage" => self.page,
                "TheDoc" => self.document?,
                _ => return None,
            };
            return Some(sheet);
        }
        let sheet = sheet.unwrap_or(self.current);
        self.shapes.get(&sheet)
    }
    fn cell_name<'a>(&self, name: &'a str) -> &'a str {
        name.split_once('!').map_or(name, |(_, name)| name)
    }
}
impl References for ShapeReferences<'_> {
    fn formula(&self, name: &str) -> Option<&str> {
        self.formula_in(None, name)
    }
    fn value(&self, name: &str) -> Option<(&str, Option<&str>)> {
        self.value_in(None, name)
    }
    fn formula_in(&self, sheet: Option<u32>, name: &str) -> Option<&str> {
        References::formula(self.target(sheet, name)?, self.cell_name(name))
    }
    fn value_in(&self, sheet: Option<u32>, name: &str) -> Option<(&str, Option<&str>)> {
        References::value(self.target(sheet, name)?, self.cell_name(name))
    }
    fn exhausted_inheritance(&self, name: &str) -> bool {
        self.target(None, name)
            .is_some_and(|shape| References::exhausted_inheritance(shape, self.cell_name(name)))
    }
    fn reference_key(&self, sheet: Option<u32>, name: &str) -> String {
        if name.starts_with("ThePage!") || name.starts_with("TheDoc!") {
            return name.into();
        }
        let (sheet, name) = match name
            .strip_prefix("Sheet.")
            .and_then(|name| name.split_once('!'))
        {
            Some((prefix, name)) => (prefix.parse().ok().or(sheet), name),
            None => (sheet, name),
        };
        sheet.map_or_else(|| name.into(), |sheet| format!("Sheet.{sheet}!{name}"))
    }
}
impl References for BTreeMap<String, String> {
    fn formula(&self, name: &str) -> Option<&str> {
        self.get(name).map(String::as_str)
    }
}

pub fn parse(input: &str, limits: &ParseLimits) -> Result<Expr, Diagnostic> {
    let parsed = vsdx_formula::parse(
        input,
        vsdx_formula::Limits {
            max_depth: limits.max_formula_depth,
            max_nodes: limits.max_formula_nodes,
            max_tokens: limits.max_formula_tokens,
        },
    )
    .map_err(|error| Diagnostic {
        message: error.message,
    })?;
    Ok(parsed)
}
pub fn evaluate(input: &str, refs: &impl References, limits: &ParseLimits) -> Evaluation {
    evaluate_with_theme(input, refs, limits, None)
}
/// Evaluates a cell formula with its ShapeSheet host-cell identity.
pub fn evaluate_cell(
    name: &str,
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
) -> Evaluation {
    evaluate_cell_with_theme(name, input, refs, limits, None)
}

fn evaluate_cell_with_theme(
    name: &str,
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    theme: Option<&Theme>,
) -> Evaluation {
    if is_event_cell(name) {
        // Event/recalculation plumbing is outside the display evaluation profile.
        return unsupported("event cell is outside the display evaluation profile");
    }
    if name.eq_ignore_ascii_case("TheText") {
        return unsupported("TheText requires phase-4b text layout");
    }
    if input.trim().eq_ignore_ascii_case("Inh") {
        return match refs.formula(name) {
            Some(formula) if !formula.eq_ignore_ascii_case("Inh") => {
                evaluate_with_theme_at(formula, refs, limits, theme, Some(name))
            }
            _ if refs.exhausted_inheritance(name) => err("Inh has no concrete inherited value"),
            _ => unsupported("Inh requires an inheritance host"),
        };
    }
    evaluate_with_theme_at(input, refs, limits, theme, Some(name))
}

fn is_event_cell(name: &str) -> bool {
    matches!(
        name.rsplit_once('!').map_or(name, |(_, name)| name),
        "EventXFMod"
            | "BegTrigger"
            | "EndTrigger"
            | "EventDblClick"
            | "EventDrop"
            | "EventMultiDrop"
    )
}
/// Evaluates against the active theme selected for the shape/page by the caller.
/// ThemeIndex and ColorSchemeIndex selection belongs to resolution, where the package is available.
pub fn evaluate_with_theme(
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    theme: Option<&Theme>,
) -> Evaluation {
    evaluate_with_theme_at(input, refs, limits, theme, None)
}
fn evaluate_with_theme_at(
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    theme: Option<&Theme>,
    host: Option<&str>,
) -> Evaluation {
    match parse(input, limits) {
        Ok(expr) => Engine {
            refs,
            limits,
            theme,
            active: HashSet::new(),
            memo: HashMap::new(),
            steps: 0,
            sheet: None,
            scope: None,
            host: host.map(str::to_owned),
        }
        .expr(&expr, 0),
        Err(error) => Evaluation::Error(error),
    }
}
/// Evaluates using the shape's ThemeIndex, falling back to ColorSchemeIndex.
pub fn evaluate_with_shape_themes(
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    shape: &ResolvedShape,
    themes: &BTreeMap<u32, Theme>,
) -> Evaluation {
    let theme = shape
        .theme_index()
        .or_else(|| shape.color_scheme_index())
        .and_then(|index| themes.get(&index));
    evaluate_with_theme(input, refs, limits, theme)
}

/// Evaluates with themes parsed from the VSDX package. ThemeIndex is one-based in relationship order;
/// ColorSchemeIndex is used only when ThemeIndex is absent.
pub fn evaluate_with_shape_package_theme(
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    shape: &ResolvedShape,
    package: &VsdxPackage,
) -> Evaluation {
    let theme = shape
        .theme_index()
        .or_else(|| shape.color_scheme_index())
        .and_then(|index| package.themes.get(&index));
    evaluate_with_theme(input, refs, limits, theme)
}

/// Evaluates a host cell using the shape's selected package theme.
pub fn evaluate_cell_with_shape_package_theme(
    name: &str,
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    shape: &ResolvedShape,
    package: &VsdxPackage,
) -> Evaluation {
    let theme = shape
        .theme_index()
        .or_else(|| shape.color_scheme_index())
        .and_then(|index| package.themes.get(&index));
    evaluate_cell_with_theme(name, input, refs, limits, theme)
}

/// Evaluates a sheet cell with the package's default theme when one is available.
pub fn evaluate_cell_with_package_theme(
    name: &str,
    input: &str,
    refs: &impl References,
    limits: &ParseLimits,
    package: &VsdxPackage,
) -> Evaluation {
    evaluate_cell_with_theme(name, input, refs, limits, package.themes.get(&1))
}

struct Engine<'a, R> {
    refs: &'a R,
    limits: &'a ParseLimits,
    theme: Option<&'a Theme>,
    active: HashSet<String>,
    memo: HashMap<String, Evaluation>,
    steps: usize,
    sheet: Option<u32>,
    scope: Option<&'static str>,
    host: Option<String>,
}
impl<R: References> Engine<'_, R> {
    fn expr(&mut self, expr: &Expr, depth: usize) -> Evaluation {
        self.steps += 1;
        if self.steps > self.limits.max_formula_steps {
            return err("formula evaluation step limit exceeded");
        }
        if depth > self.limits.max_formula_depth {
            return err("formula depth limit exceeded");
        }
        match expr {
            Expr::Number(n, u) => number(*n, *u),
            Expr::String(_) => unsupported("string values are not display numbers"),
            Expr::Reference(name) => {
                if name.eq_ignore_ascii_case("Inh") {
                    return err("Inh has no concrete inherited value");
                }
                if is_event_cell(name) {
                    return unsupported("event cell is outside the display evaluation profile");
                }
                if name.eq_ignore_ascii_case("TheText") {
                    return unsupported("TheText requires phase-4b text layout");
                }
                if name.eq_ignore_ascii_case("FALSE") {
                    return number(0., Unit::Bool);
                }
                if name.eq_ignore_ascii_case("TRUE") {
                    return number(1., Unit::Bool);
                }
                let key_name = if name.contains('!') {
                    name.clone()
                } else {
                    self.scope
                        .map_or_else(|| name.clone(), |scope| format!("{scope}!{name}"))
                };
                let key = self.refs.reference_key(self.sheet, &key_name);
                if let Some(value) = self.memo.get(&key) {
                    return value.clone();
                }
                if !self.active.insert(key.clone()) {
                    return err(format!("reference cycle at {key}"));
                }
                let (sheet, scope, lookup) = match name
                    .strip_prefix("Sheet.")
                    .and_then(|name| name.split_once('!'))
                {
                    Some((sheet, name)) => (sheet.parse().ok(), None, name),
                    None => match name.split_once('!') {
                        Some(("ThePage", name)) => (None, Some("ThePage"), name),
                        Some(("TheDoc", name)) => (None, Some("TheDoc"), name),
                        _ => (self.sheet, self.scope, name.as_str()),
                    },
                };
                let result = match self.refs.formula_in_scoped(sheet, scope, lookup) {
                    Some(formula) => match parse(formula, self.limits) {
                        Ok(e) => {
                            let previous = self.sheet;
                            let previous_scope = self.scope;
                            let previous_host = self.host.replace(lookup.to_owned());
                            self.sheet = sheet;
                            self.scope = scope;
                            let result = self.expr(&e, depth + 1);
                            self.sheet = previous;
                            self.scope = previous_scope;
                            self.host = previous_host;
                            result
                        }
                        Err(e) => Evaluation::Error(e),
                    },
                    None => self.refs.value_in_scoped(sheet, scope, lookup).map_or_else(
                        || err(format!("unresolved reference {name}")),
                        |(value, unit)| cell_value(value, unit),
                    ),
                };
                self.active.remove(&key);
                self.memo.insert(key, result.clone());
                result
            }
            Expr::Unary(v) => match numeric(self.expr(v, depth + 1)) {
                Ok((v, guarded)) => result(
                    Value::Number(Number {
                        number: -v.number,
                        unit: v.unit,
                    }),
                    guarded,
                ),
                Err(r) => r,
            },
            Expr::Binary(a, op, b) => self.binary(a, *op, b, depth + 1),
            Expr::Call(name, args) => self.call(name, args, depth + 1),
        }
    }
    fn binary(&mut self, a: &Expr, op: Op, b: &Expr, d: usize) -> Evaluation {
        let a = match numeric(self.expr(a, d)) {
            Ok(v) => v,
            Err(r) => return r,
        };
        let b = match numeric(self.expr(b, d)) {
            Ok(v) => v,
            Err(r) => return r,
        };
        let (a, a_guarded) = a;
        let (b, b_guarded) = b;
        let guarded = a_guarded || b_guarded;
        match op {
            Op::Add | Op::Sub => {
                if a.unit != b.unit {
                    return err("incompatible units");
                }
                numeric_result(
                    if op == Op::Add {
                        a.number + b.number
                    } else {
                        a.number - b.number
                    },
                    a.unit,
                    guarded,
                )
            }
            Op::Mul => {
                if a.unit != Unit::Number && b.unit != Unit::Number {
                    err("cannot multiply dimensional values")
                } else {
                    numeric_result(
                        a.number * b.number,
                        if a.unit == Unit::Number {
                            b.unit
                        } else {
                            a.unit
                        },
                        guarded,
                    )
                }
            }
            Op::Div => {
                if b.number == 0.0 {
                    return err("division by zero");
                };
                if b.unit != Unit::Number && a.unit != b.unit {
                    return err("incompatible units");
                };
                numeric_result(
                    a.number / b.number,
                    if a.unit == b.unit {
                        Unit::Number
                    } else {
                        a.unit
                    },
                    guarded,
                )
            }
            Op::Pow => {
                if b.unit != Unit::Number {
                    return err("exponent must be dimensionless");
                };
                if a.unit != Unit::Number && b.number.fract() != 0.0 {
                    return err("dimensional powers require an integral exponent");
                }
                numeric_result(a.number.powf(b.number), a.unit, guarded)
            }
            Op::Eq | Op::Ne | Op::Lt | Op::Le | Op::Gt | Op::Ge => {
                if a.unit != b.unit {
                    return err("incompatible units");
                };
                let x = match op {
                    Op::Eq => a.number == b.number,
                    Op::Ne => a.number != b.number,
                    Op::Lt => a.number < b.number,
                    Op::Le => a.number <= b.number,
                    Op::Gt => a.number > b.number,
                    Op::Ge => a.number >= b.number,
                    _ => false,
                };
                numeric_result(if x { 1. } else { 0. }, Unit::Bool, guarded)
            }
        }
    }
    fn call(&mut self, name: &str, args: &[Expr], d: usize) -> Evaluation {
        let upper = name.to_ascii_uppercase();
        if matches!(
            upper.as_str(),
            "SETATREF" | "SETATREFEXPR" | "SETATREFEVAL" | "DEPENDSON"
        ) {
            return unsupported(format!("{upper} is outside the phase-4 evaluator"));
        }
        if upper == "GUARD" {
            // Visio GUARD intercepts edits; display evaluation returns its argument. Mutation policy is phase 5.
            return args
                .first()
                .map_or_else(|| err("missing argument"), |arg| guard(self.expr(arg, d)));
        }
        if matches!(upper.as_str(), "THEMEGUARD" | "_XFTRIGGER") {
            // THEMEGUARD protects theme edits and _XFTRIGGER schedules recalculation; both are display-transparent.
            return args
                .first()
                .map_or_else(|| err("missing argument"), |arg| self.expr(arg, d));
        }
        if upper == "THEMEVAL" {
            return self.themeval(args, d);
        }
        if upper == "RGB" {
            return self.rgb(args, d);
        }
        if matches!(
            upper.as_str(),
            "LUMDIFF" | "MSOTINT" | "SHADE" | "TINT" | "SAT"
        ) {
            return self.color_function(&upper, args, d);
        }
        if upper == "IF" {
            if args.len() != 3 {
                return err("IF requires three arguments");
            }
            let (condition, guarded) = match numeric(self.expr(&args[0], d)) {
                Ok(value) => value,
                Err(error) => return error,
            };
            return match self.expr(
                if condition.number != 0. {
                    &args[1]
                } else {
                    &args[2]
                },
                d,
            ) {
                Evaluation::Evaluated(mut value) => {
                    value.guarded |= guarded;
                    Evaluation::Evaluated(value)
                }
                other => other,
            };
        }
        let vals: Result<Vec<_>, _> = args.iter().map(|a| numeric(self.expr(a, d))).collect();
        let vals = match vals {
            Ok(v) => v,
            Err(r) => return r,
        };
        let guarded = vals.iter().any(|(_, guarded)| *guarded);
        let vals: Vec<Number> = vals.into_iter().map(|(value, _)| value).collect();
        let one = || vals.first().copied().ok_or_else(|| err("missing argument"));
        let same = || {
            if vals.iter().map(|v| v.unit).all(|u| u == vals[0].unit) {
                Ok(())
            } else {
                Err(err("incompatible units"))
            }
        };
        match upper.as_str() {
            "AND" => numeric_result(
                if vals.iter().all(|v| v.number != 0.) {
                    1.
                } else {
                    0.
                },
                Unit::Bool,
                guarded,
            ),
            "OR" => numeric_result(
                if vals.iter().any(|v| v.number != 0.) {
                    1.
                } else {
                    0.
                },
                Unit::Bool,
                guarded,
            ),
            "NOT" => one().map_or_else(
                |r| r,
                |v| numeric_result(if v.number == 0. { 1. } else { 0. }, Unit::Bool, guarded),
            ),
            "MIN" | "MAX" | "SUM" => {
                if vals.is_empty() {
                    return err("missing argument");
                }
                if let Err(r) = same() {
                    return r;
                }
                let n = if upper == "SUM" {
                    vals.iter().map(|v| v.number).sum()
                } else if upper == "MIN" {
                    vals.iter().map(|v| v.number).fold(f64::INFINITY, f64::min)
                } else {
                    vals.iter()
                        .map(|v| v.number)
                        .fold(f64::NEG_INFINITY, f64::max)
                };
                numeric_result(n, vals[0].unit, guarded)
            }
            "ABS" => one().map_or_else(|r| r, |v| numeric_result(v.number.abs(), v.unit, guarded)),
            "INT" => {
                one().map_or_else(|r| r, |v| numeric_result(v.number.floor(), v.unit, guarded))
            }
            "TRUNC" => {
                one().map_or_else(|r| r, |v| numeric_result(v.number.trunc(), v.unit, guarded))
            }
            "SIGN" => one().map_or_else(
                |r| r,
                |v| numeric_result(v.number.signum(), Unit::Number, guarded),
            ),
            "ROUND" => one().map_or_else(
                |r| r,
                |v| numeric_result((v.number + 0.5).floor(), v.unit, guarded),
            ),
            "CEILING" => {
                one().map_or_else(|r| r, |v| numeric_result(v.number.ceil(), v.unit, guarded))
            }
            "FLOOR" => {
                one().map_or_else(|r| r, |v| numeric_result(v.number.floor(), v.unit, guarded))
            }
            "SQRT" => one().map_or_else(
                |r| r,
                |v| {
                    if v.number < 0. {
                        err("square root of negative number")
                    } else if v.unit != Unit::Number {
                        unsupported("SQRT of dimensional values is not implemented")
                    } else {
                        numeric_result(v.number.sqrt(), v.unit, guarded)
                    }
                },
            ),
            "SIN" | "COS" | "TAN" => one().map_or_else(
                |r| r,
                |v| {
                    if v.unit != Unit::Radians {
                        err("trigonometric argument must be an angle")
                    } else {
                        numeric_result(
                            if upper == "SIN" {
                                v.number.sin()
                            } else if upper == "COS" {
                                v.number.cos()
                            } else {
                                v.number.tan()
                            },
                            Unit::Number,
                            guarded,
                        )
                    }
                },
            ),
            "ATAN2" if vals.len() == 2 => {
                if vals[0].unit != vals[1].unit {
                    return err("incompatible units");
                }
                numeric_result(vals[0].number.atan2(vals[1].number), Unit::Radians, guarded)
            }
            "PI" if vals.is_empty() => numeric_result(std::f64::consts::PI, Unit::Radians, guarded),
            "MOD" if vals.len() == 2 => {
                if vals[0].unit != vals[1].unit {
                    return err("incompatible units");
                }
                if vals[1].number == 0. {
                    err("division by zero")
                } else {
                    numeric_result(
                        vals[0].number.rem_euclid(vals[1].number),
                        vals[0].unit,
                        guarded,
                    )
                }
            }
            _ => unsupported(format!("unsupported function {upper}")),
        }
    }
    fn rgb(&mut self, args: &[Expr], d: usize) -> Evaluation {
        let components: Result<Vec<_>, _> =
            args.iter().map(|arg| numeric(self.expr(arg, d))).collect();
        let Ok(components) = components else {
            return components.unwrap_err();
        };
        if components.len() != 3 {
            return err("RGB requires three arguments");
        }
        if components
            .iter()
            .any(|(value, _)| value.unit != Unit::Number)
        {
            return err("RGB channels must be dimensionless");
        }
        // Visio RGB takes 8-bit channels; out-of-range inputs are conservatively saturated.
        let channel = |value: f64| value.round().clamp(0.0, 255.0) as u8;
        result(
            Value::Color(Color {
                red: channel(components[0].0.number),
                green: channel(components[1].0.number),
                blue: channel(components[2].0.number),
                alpha: None,
            }),
            components.iter().any(|(_, guarded)| *guarded),
        )
    }
    fn themeval(&mut self, args: &[Expr], d: usize) -> Evaluation {
        let host = self
            .host
            .as_deref()
            .and_then(|host| host.rsplit('.').next());
        let name = match args.first() {
            Some(Expr::String(name)) => name.clone(),
            Some(expr) => match numeric(self.expr(expr, d)) {
                Ok((value, _)) if value.unit == Unit::Number && value.number.fract() == 0.0 => {
                    match value.number as i32 {
                        1 => "Dark".to_owned(),
                        2 => "Light".to_owned(),
                        3 => "AccentColor1".to_owned(),
                        4 => "AccentColor2".to_owned(),
                        5 => "AccentColor3".to_owned(),
                        6 => "AccentColor4".to_owned(),
                        7 => "AccentColor5".to_owned(),
                        8 => "AccentColor6".to_owned(),
                        _ => {
                            return unsupported("THEMEVAL colour-scheme index must be 1 through 8");
                        }
                    }
                }
                Ok(_) => return unsupported("THEMEVAL requires a string or integer theme value"),
                Err(error) => return error,
            },
            None => match host {
                Some("FillForegnd") => "FillColor".to_owned(),
                Some("FillBkgnd") => "FillColor2".to_owned(),
                Some("LineColor") => "LineColor".to_owned(),
                Some("Color") => "TextColor".to_owned(),
                _ => return unsupported("THEMEVAL host-cell lookup requires theme-cell context"),
            },
        };
        let Some(theme) = self.theme else {
            return args.get(1).map_or_else(
                || unsupported("THEMEVAL has no resolvable theme"),
                |arg| self.expr(arg, d),
            );
        };
        let name = name.as_str();
        let slot = match name {
            "BackgroundColor" => "lt1",
            "Light" => "lt1",
            "FillColor" => "accent1",
            "FillColor2" => "accent2",
            "LineColor" => "dk1",
            "TextColor" => "dk1",
            "AccentColor1" => "accent1",
            "AccentColor2" => "accent2",
            "AccentColor3" => "accent3",
            "AccentColor4" => "accent4",
            "AccentColor5" => "accent5",
            "AccentColor6" => "accent6",
            "VariantColor1" => "accent1",
            "VariantColor2" => "accent2",
            "VariantColor3" => "accent3",
            "VariantColor4" => "accent4",
            _ => {
                return args.get(1).map_or_else(
                    || unsupported("unresolvable THEMEVAL value"),
                    |arg| self.expr(arg, d),
                );
            }
        };
        // Theme values use DrawingML colour slots; unknown named Visio theme values intentionally remain unsupported.
        let color = ColorValue {
            theme_color: Some(slot.to_ascii_lowercase()),
            ..ColorValue::default()
        };
        let hex = resolve_color_value_to_hex_with_theme(Some(&color), Some(theme))
            .or_else(|| Some(format!("#{}", get_theme_color(Some(theme), slot))));
        hex.and_then(|hex| parse_color(&hex)).map_or_else(
            || {
                args.get(1).map_or_else(
                    || unsupported("unresolvable THEMEVAL value"),
                    |arg| self.expr(arg, d),
                )
            },
            |color| result(Value::Color(color), false),
        )
    }
    fn color_function(&mut self, name: &str, args: &[Expr], d: usize) -> Evaluation {
        if name == "SAT" {
            let Some(first) = args.first() else {
                return err("SAT requires one argument");
            };
            if args.len() != 1 {
                return err("SAT requires one argument");
            }
            let (color, guarded) = match color_value(self.expr(first, d)) {
                Ok(value) => value,
                Err(error) => return error,
            };
            return numeric_result(saturation(color), Unit::Number, guarded);
        }
        let Some((first, second)) = args.first().zip(args.get(1)) else {
            return err(format!("{name} requires two arguments"));
        };
        let (color, color_guarded) = match color_value(self.expr(first, d)) {
            Ok(value) => value,
            Err(error) => return error,
        };
        if matches!(name, "LUMDIFF" | "SHADE") {
            // Visio does not document enough of these colour-model semantics to render them honestly.
            return unsupported(format!("{name} is not implemented"));
        }
        let (amount, amount_guarded) = match numeric(self.expr(second, d)) {
            Ok(value) => value,
            Err(error) => return error,
        };
        match name {
            "TINT" | "MSOTINT" => {
                if amount.unit != Unit::Number || amount.number.fract() != 0.0 {
                    return err(format!("{name} amount must be a dimensionless integer"));
                }
                let color = if name == "TINT" {
                    tint(color, amount.number)
                } else if !(-100.0..=100.0).contains(&amount.number) {
                    return err("MSOTINT percentage must be between -100 and 100");
                } else {
                    mso_tint(color, amount.number)
                };
                result(Value::Color(color), color_guarded || amount_guarded)
            }
            _ => unreachable!(),
        }
    }
}
fn number(number: f64, unit: Unit) -> Evaluation {
    numeric_result(number, unit, false)
}
fn cell_value(value: &str, unit_name: Option<&str>) -> Evaluation {
    if let Some(color) = parse_color(value) {
        return result(Value::Color(color), false);
    }
    let Ok(value) = value.parse::<f64>() else {
        return unsupported("cell value is not a supported display literal");
    };
    let Some((unit, scale)) = unit(unit_name.unwrap_or("")) else {
        return unsupported("cell value has an unsupported unit");
    };
    number(value * scale, unit)
}
fn numeric_result(number: f64, unit: Unit, guarded: bool) -> Evaluation {
    if number.is_finite() {
        result(Value::Number(Number { number, unit }), guarded)
    } else {
        unsupported("non-finite result")
    }
}
fn err(message: impl Into<String>) -> Evaluation {
    Evaluation::Error(Diagnostic {
        message: message.into(),
    })
}
fn unsupported(message: impl Into<String>) -> Evaluation {
    Evaluation::Unsupported(message.into())
}
fn result(value: Value, guarded: bool) -> Evaluation {
    Evaluation::Evaluated(Evaluated { value, guarded })
}
fn guard(result: Evaluation) -> Evaluation {
    match result {
        Evaluation::Evaluated(mut value) => {
            value.guarded = true;
            Evaluation::Evaluated(value)
        }
        other => other,
    }
}
fn numeric(result: Evaluation) -> Result<(Number, bool), Evaluation> {
    match result {
        Evaluation::Evaluated(Evaluated {
            value: Value::Number(value),
            guarded,
        }) => Ok((value, guarded)),
        Evaluation::Evaluated(_) => Err(err("colour used where a numeric value is required")),
        x => Err(x),
    }
}
fn color_value(result: Evaluation) -> Result<(Color, bool), Evaluation> {
    match result {
        Evaluation::Evaluated(Evaluated {
            value: Value::Color(value),
            guarded,
        }) => Ok((value, guarded)),
        Evaluation::Evaluated(_) => Err(err("numeric value used where a colour is required")),
        x => Err(x),
    }
}
fn parse_color(hex: &str) -> Option<Color> {
    let value = hex.strip_prefix('#').unwrap_or(hex);
    if value.len() != 6 {
        return None;
    }
    let packed = u32::from_str_radix(value, 16).ok()?;
    Some(Color {
        red: (packed >> 16) as u8,
        green: (packed >> 8) as u8,
        blue: packed as u8,
        alpha: None,
    })
}
fn saturation(color: Color) -> f64 {
    rgb_to_hls(color).1
}
fn tint(color: Color, amount: f64) -> Color {
    let (hue, saturation, luminosity) = rgb_to_hls(color);
    hls_to_rgb(hue, saturation, (luminosity + amount).clamp(0.0, 240.0))
}
fn mso_tint(color: Color, percentage: f64) -> Color {
    let (hue, saturation, luminosity) = rgb_to_hls(color);
    let luminosity = if percentage < 0.0 {
        luminosity + (-percentage / 100.0) * (240.0 - luminosity)
    } else {
        luminosity * (1.0 - percentage / 100.0)
    };
    hls_to_rgb(hue, saturation, luminosity)
}
fn rgb_to_hls(color: Color) -> (f64, f64, f64) {
    let red = f64::from(color.red) / 255.0;
    let green = f64::from(color.green) / 255.0;
    let blue = f64::from(color.blue) / 255.0;
    let high = red.max(green).max(blue);
    let low = red.min(green).min(blue);
    let delta = high - low;
    let luminosity = (high + low) / 2.0;
    if delta == 0.0 {
        return (0.0, 0.0, luminosity * 240.0);
    }
    let saturation = delta / (1.0 - (2.0 * luminosity - 1.0).abs());
    let hue = if high == red {
        ((green - blue) / delta).rem_euclid(6.0)
    } else if high == green {
        (blue - red) / delta + 2.0
    } else {
        (red - green) / delta + 4.0
    } / 6.0;
    (hue * 240.0, saturation * 240.0, luminosity * 240.0)
}
fn hls_to_rgb(hue: f64, saturation: f64, luminosity: f64) -> Color {
    let hue = hue / 240.0;
    let saturation = saturation / 240.0;
    let luminosity = luminosity / 240.0;
    let chroma = (1.0 - (2.0 * luminosity - 1.0).abs()) * saturation;
    let x = chroma * (1.0 - ((hue * 6.0).rem_euclid(2.0) - 1.0).abs());
    let (red, green, blue) = match (hue * 6.0).floor() as i32 {
        0 => (chroma, x, 0.0),
        1 => (x, chroma, 0.0),
        2 => (0.0, chroma, x),
        3 => (0.0, x, chroma),
        4 => (x, 0.0, chroma),
        _ => (chroma, 0.0, x),
    };
    let offset = luminosity - chroma / 2.0;
    let channel = |value: f64| ((value + offset) * 255.0).round().clamp(0.0, 255.0) as u8;
    Color {
        red: channel(red),
        green: channel(green),
        blue: channel(blue),
        alpha: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use vsdx_parse::{Cell, Shape, Sheet, parse_vsdx};
    use vsdx_resolve::{Lookup, Provenance, ResolvedCell, ResolvedRow, ResolvedSection, Resolver};

    fn limits() -> ParseLimits {
        ParseLimits::default()
    }
    fn number(formula: &str) -> Number {
        match evaluate(formula, &BTreeMap::new(), &limits()) {
            Evaluation::Evaluated(Evaluated {
                value: Value::Number(value),
                ..
            }) => value,
            value => panic!("expected value, got {value:?}"),
        }
    }

    fn resolved_value(name: &str, value: &str) -> Lookup {
        Lookup::Found(ResolvedCell {
            cell: Cell {
                name: name.into(),
                formula: None,
                value: Some(value.into()),
                unit: None,
                del: false,
                other_attrs: Vec::new(),
            },
            provenance: Provenance::Local,
        })
    }

    #[test]
    fn parses_literals_references_and_precedence() {
        assert_eq!(number("1 + 2 * 3").number, 7.0);
        assert_eq!(number("2^3^2").number, 512.0);
        assert_eq!(number("-(1 + 2)").number, -3.0);
        assert_eq!(number("1.5E-6").number, 1.5E-6);
        assert_eq!(number("2em").number, 120.0);
        assert_eq!(number("3ed").number, 3.0 * 24.0 * 60.0 * 60.0);
        assert_eq!(number("4es").number, 4.0);
        assert_eq!(number("5ew").number, 5.0 * 7.0 * 24.0 * 60.0 * 60.0);
        for formula in ["1e", "1e+"] {
            assert!(
                matches!(parse(formula, &limits()), Err(Diagnostic { message }) if message == "unknown unit suffix e")
            );
        }
        assert!(matches!(
            parse("Geometry1.X2", &limits()),
            Ok(Expr::Reference(_))
        ));
        assert!(matches!(
            parse("Sheet.5!Width", &limits()),
            Ok(Expr::Reference(_))
        ));
        assert!(matches!(
            parse("IF(AND(1, 1), MIN(3, 2), 0)", &limits()),
            Ok(Expr::Call(_, _))
        ));
    }

    #[test]
    fn converts_units_and_rejects_mismatches() {
        let value = number("1 in + 25.4 mm");
        assert_eq!(value.unit, Unit::Inches);
        assert!((value.number - 2.0).abs() < 1e-12);
        assert!(matches!(
            evaluate("1 in + 1 deg", &BTreeMap::new(), &limits()),
            Evaluation::Error(_)
        ));
        assert!(matches!(
            evaluate("4 in ^ 0.5", &BTreeMap::new(), &limits()),
            Evaluation::Error(_)
        ));
    }

    #[test]
    fn resolves_page_and_document_scope_references() {
        fn value(value: &str) -> Lookup {
            Lookup::Found(ResolvedCell {
                cell: Cell {
                    name: "Value".into(),
                    formula: None,
                    value: Some(value.into()),
                    unit: None,
                    del: false,
                    other_attrs: Vec::new(),
                },
                provenance: Provenance::Local,
            })
        }
        fn formula(name: &str, formula: &str) -> Lookup {
            Lookup::Found(ResolvedCell {
                cell: Cell {
                    name: name.into(),
                    formula: Some(formula.into()),
                    value: None,
                    unit: None,
                    del: false,
                    other_attrs: Vec::new(),
                },
                provenance: Provenance::Local,
            })
        }
        let page = ResolvedShape {
            cells: BTreeMap::from([
                ("DrawingScale".into(), value("2")),
                ("PageWidth".into(), formula("PageWidth", "DrawingScale*5")),
            ]),
            sections: BTreeMap::from([(
                "User".into(),
                ResolvedSection {
                    name: "User".into(),
                    rows: BTreeMap::from([(
                        "N:x".into(),
                        ResolvedRow {
                            key: "N:x".into(),
                            cells: BTreeMap::from([("Value".into(), value("7"))]),
                            ..Default::default()
                        },
                    )]),
                    ..Default::default()
                },
            )]),
            ..Default::default()
        };
        let document = ResolvedShape {
            cells: BTreeMap::from([("DocScale".into(), value("3"))]),
            ..Default::default()
        };
        let shape = ResolvedShape {
            cells: BTreeMap::from([("Width".into(), formula("Width", "2"))]),
            ..Default::default()
        };
        let page_refs = PageShapeReferences {
            shapes: BTreeMap::from([(5, shape)]),
            page,
            document: Some(document),
        };
        let refs = page_refs.for_shape(1);
        assert_eq!(number_with(&refs, "ThePage!DrawingScale").number, 2.0);
        assert_eq!(number_with(&refs, "ThePage!PageWidth").number, 10.0);
        assert_eq!(number_with(&refs, "TheDoc!DocScale").number, 3.0);
        assert_eq!(number_with(&refs, "ThePage!User.x").number, 7.0);
        assert_eq!(refs.formula("Sheet.5!Width"), Some("2"));
        for name in ["ThePage!Unknown", "TheDoc!Unknown"] {
            assert!(matches!(
                evaluate(name, &refs, &limits()),
                Evaluation::Error(Diagnostic { message }) if message == format!("unresolved reference {name}")
            ));
        }
    }

    #[test]
    fn resolves_document_references_without_a_page_context() {
        let document = ResolvedShape {
            cells: BTreeMap::from([("DocScale".into(), resolved_value("DocScale", "3"))]),
            sections: BTreeMap::from([(
                "User".into(),
                ResolvedSection {
                    name: "User".into(),
                    rows: BTreeMap::from([(
                        "N:x".into(),
                        ResolvedRow {
                            key: "N:x".into(),
                            cells: BTreeMap::from([("Value".into(), resolved_value("Value", "7"))]),
                            ..Default::default()
                        },
                    )]),
                    ..Default::default()
                },
            )]),
            ..Default::default()
        };
        let master = ResolvedShape::default();
        let master_refs = DocumentReferences::new(&master, Some(&document));
        assert_eq!(
            number_with_cell(&master_refs, "Width", "TheDoc!DocScale").number,
            3.0
        );
        assert_eq!(
            number_with_cell(&master_refs, "Width", "TheDoc!User.x").number,
            7.0
        );
        assert!(matches!(
            evaluate_cell("Width", "ThePage!PageWidth", &master_refs, &limits()),
            Evaluation::Error(Diagnostic { message }) if message == "unresolved reference ThePage!PageWidth"
        ));

        let style_sheet = BTreeMap::from([("Width".into(), "TheDoc!DocScale".into())]);
        let style_refs = DocumentReferences::new(&style_sheet, Some(&document));
        assert_eq!(number_with(&style_refs, "Width").number, 3.0);

        let page_sheet = BTreeMap::from([("Height".into(), "TheDoc!DocScale".into())]);
        let page_refs = DocumentReferences::new(&page_sheet, Some(&document));
        assert_eq!(number_with(&page_refs, "Height").number, 3.0);
    }

    fn number_with(refs: &impl References, formula: &str) -> Number {
        match evaluate(formula, refs, &limits()) {
            Evaluation::Evaluated(Evaluated {
                value: Value::Number(value),
                ..
            }) => value,
            value => panic!("expected value, got {value:?}"),
        }
    }

    fn number_with_cell(refs: &impl References, name: &str, formula: &str) -> Number {
        match evaluate_cell(name, formula, refs, &limits()) {
            Evaluation::Evaluated(Evaluated {
                value: Value::Number(value),
                ..
            }) => value,
            value => panic!("expected value, got {value:?}"),
        }
    }

    #[test]
    fn evaluates_display_functions_and_unsupported_calls() {
        assert_eq!(number("MOD(-3, 2)").number, 1.0);
        assert!((number("ATAN2(1, -1)").number - 3.0 * std::f64::consts::PI / 4.0).abs() < 1e-12);
        assert_eq!(number("ROUND(1.5)").number, 2.0);
        for formula in ["SETATREF(1)", "NotAFunction(1)"] {
            assert!(matches!(
                evaluate(formula, &BTreeMap::new(), &limits()),
                Evaluation::Unsupported(_)
            ));
        }
        assert!(matches!(
            evaluate("GUARD(1)", &BTreeMap::new(), &limits()),
            Evaluation::Evaluated(Evaluated { guarded: true, .. })
        ));
    }

    fn color(formula: &str, theme: Option<&Theme>) -> Color {
        match evaluate_with_theme(formula, &BTreeMap::new(), &limits(), theme) {
            Evaluation::Evaluated(Evaluated {
                value: Value::Color(value),
                ..
            }) => value,
            value => panic!("expected colour, got {value:?}"),
        }
    }

    #[test]
    fn evaluates_colours_and_wrappers() {
        assert_eq!(
            color("RGB(-1, 12.6, 300)", None),
            Color {
                red: 0,
                green: 13,
                blue: 255,
                alpha: None
            }
        );
        assert_eq!(
            color("THEMEGUARD(RGB(1,2,3))", None),
            Color {
                red: 1,
                green: 2,
                blue: 3,
                alpha: None
            }
        );
        assert_eq!(number("_XFTRIGGER(7)").number, 7.0);
        assert!(matches!(
            evaluate("RGB(1,2,3)+1", &BTreeMap::new(), &limits()),
            Evaluation::Error(_)
        ));
    }

    #[test]
    fn resolves_theme_values_or_uses_fallback() {
        let mut theme = Theme::default();
        theme.color_scheme.accent4 = "102030".to_owned();
        assert_eq!(
            color("THEMEVAL(\"AccentColor4\")", Some(&theme)),
            Color {
                red: 16,
                green: 32,
                blue: 48,
                alpha: None
            }
        );
        assert_eq!(
            color("THEMEVAL(6)", Some(&theme)),
            Color {
                red: 16,
                green: 32,
                blue: 48,
                alpha: None
            }
        );
        assert_eq!(
            color("THEMEVAL(\"AccentColor4\",RGB(4,5,6))", None),
            Color {
                red: 4,
                green: 5,
                blue: 6,
                alpha: None
            }
        );
        assert!(matches!(
            evaluate("THEMEVAL(\"AccentColor4\")", &BTreeMap::new(), &limits()),
            Evaluation::Unsupported(_)
        ));
        assert!(matches!(
            evaluate("THEMEVAL()", &BTreeMap::new(), &limits()),
            Evaluation::Unsupported(_)
        ));
        assert_eq!(
            color("THEMEGUARD(THEMEVAL(\"AccentColor4\"))", Some(&theme)),
            Color {
                red: 16,
                green: 32,
                blue: 48,
                alpha: None
            }
        );
    }

    #[test]
    fn resolves_host_relative_theme_value_through_a_reference() {
        let refs = BTreeMap::from([("FillForegnd".into(), "THEMEVAL()".into())]);
        let theme = Theme::default();
        assert!(matches!(
            evaluate_cell_with_theme("LineColor", "FillForegnd", &refs, &limits(), Some(&theme)),
            Evaluation::Evaluated(Evaluated {
                value: Value::Color(Color {
                    red: 68,
                    green: 114,
                    blue: 196,
                    ..
                }),
                ..
            })
        ));
    }

    #[test]
    fn selects_a_theme_from_shape_indices() {
        let mut shape = ResolvedShape::default();
        shape.cells.insert(
            "ThemeIndex".into(),
            Lookup::Found(vsdx_resolve::ResolvedCell {
                cell: vsdx_parse::Cell {
                    name: "ThemeIndex".into(),
                    formula: None,
                    value: Some("7".into()),
                    unit: None,
                    del: false,
                    other_attrs: Vec::new(),
                },
                provenance: vsdx_resolve::Provenance::Local,
            }),
        );
        let mut theme = Theme::default();
        theme.color_scheme.accent1 = "A0B0C0".into();
        let themes = BTreeMap::from([(7, theme)]);
        assert_eq!(
            evaluate_with_shape_themes(
                "THEMEVAL(\"FillColor\")",
                &BTreeMap::new(),
                &limits(),
                &shape,
                &themes
            ),
            Evaluation::Evaluated(Evaluated {
                value: Value::Color(Color {
                    red: 160,
                    green: 176,
                    blue: 192,
                    alpha: None
                }),
                guarded: false
            })
        );
    }

    #[test]
    fn resolves_shape_cell_literals_and_host_inheritance() {
        let mut shape = ResolvedShape::default();
        shape.cells.insert(
            "Width".into(),
            Lookup::Found(vsdx_resolve::ResolvedCell {
                cell: vsdx_parse::Cell {
                    name: "Width".into(),
                    formula: None,
                    value: Some("2".into()),
                    unit: Some("in".into()),
                    del: false,
                    other_attrs: Vec::new(),
                },
                provenance: vsdx_resolve::Provenance::Local,
            }),
        );
        assert_eq!(
            evaluate("Width + 1 in", &shape, &limits()),
            Evaluation::Evaluated(Evaluated {
                value: Value::Number(Number {
                    number: 3.0,
                    unit: Unit::Inches,
                }),
                guarded: false,
            })
        );
        {
            let Lookup::Found(width) = shape.cells.get_mut("Width").unwrap() else {
                panic!("expected Width");
            };
            width.cell.formula = Some("Inh".into());
        }
        assert!(matches!(
            evaluate_cell("Width", "Inh", &shape, &limits()),
            Evaluation::Error(Diagnostic { message }) if message == "Inh has no concrete inherited value"
        ));
        let Lookup::Found(width) = shape.cells.get_mut("Width").unwrap() else {
            panic!("expected Width");
        };
        width.cell.formula = Some("2 in".into());
        assert_eq!(
            evaluate_cell("Width", "Inh", &shape, &limits()),
            Evaluation::Evaluated(Evaluated {
                value: Value::Number(Number {
                    number: 2.0,
                    unit: Unit::Inches
                }),
                guarded: false,
            })
        );
        assert_eq!(number("FALSE").unit, Unit::Bool);
        assert_eq!(number("2 DL").number, 2.0);
        for name in [
            "EventXFMod",
            "Sheet.1!EventXFMod",
            "EventDblClick",
            "EventDrop",
            "EventMultiDrop",
            "BegTrigger",
            "EndTrigger",
            "TheText",
        ] {
            assert!(matches!(
                evaluate_cell(name, "1", &shape, &limits()),
                Evaluation::Unsupported(_)
            ));
        }
    }

    #[test]
    fn evaluates_documented_colour_transforms() {
        for formula in [
            "LUMDIFF(RGB(255,255,255),RGB(0,0,0))",
            "SHADE(RGB(100,150,200),0.5)",
        ] {
            assert!(matches!(
                evaluate(formula, &BTreeMap::new(), &limits()),
                Evaluation::Unsupported(_)
            ));
        }
        assert_eq!(number("SAT(RGB(255,0,0))").number, 240.0);
        assert_eq!(
            color("TINT(RGB(255,0,0),20)", None),
            Color {
                red: 255,
                green: 43,
                blue: 43,
                alpha: None
            }
        );
        assert_eq!(
            color("MSOTINT(RGB(255,0,0),-50)", None),
            Color {
                red: 255,
                green: 128,
                blue: 128,
                alpha: None
            }
        );
        assert_eq!(
            color("MSOTINT(RGB(255,0,0),50)", None),
            Color {
                red: 128,
                green: 0,
                blue: 0,
                alpha: None
            }
        );
    }

    #[test]
    fn detects_reference_cycles_and_depth() {
        let refs = BTreeMap::from([("A".into(), "B".into()), ("B".into(), "A".into())]);
        assert!(matches!(
            evaluate("A", &refs, &limits()),
            Evaluation::Error(_)
        ));
        let limits = ParseLimits {
            max_formula_depth: 2,
            ..limits()
        };
        assert!(parse("(((1)))", &limits).is_err());
    }

    #[test]
    fn bounds_every_formula_recursion_and_node_count() {
        let limits = ParseLimits {
            max_formula_depth: 32,
            max_formula_nodes: 128,
            ..limits()
        };
        assert!(parse(&format!("1{}", "^1".repeat(64)), &limits).is_err());
        assert!(parse(&format!("{}1", "-".repeat(64)), &limits).is_err());
        assert!(
            parse(
                &std::iter::repeat_n("1", 256).collect::<Vec<_>>().join("+"),
                &limits
            )
            .is_err()
        );
    }

    #[test]
    fn bounds_formula_token_count() {
        let limits = ParseLimits {
            max_formula_tokens: 8,
            ..limits()
        };
        assert!(matches!(
            parse(&std::iter::repeat_n("1", 16).collect::<Vec<_>>().join("+"), &limits),
            Err(Diagnostic { message }) if message == "formula token limit exceeded"
        ));
    }

    #[test]
    fn memoizes_diamond_references_with_a_step_budget() {
        let refs = BTreeMap::from([
            ("A".into(), "B+B".into()),
            ("B".into(), "C+C".into()),
            ("C".into(), "D+D".into()),
            ("D".into(), "1".into()),
        ]);
        let limits = ParseLimits {
            max_formula_steps: 12,
            ..limits()
        };
        assert!(matches!(
            evaluate("A", &refs, &limits),
            Evaluation::Evaluated(_)
        ));
    }

    #[test]
    fn corpus_formulas_report_honest_evaluation() {
        let directory = std::env::var("VSDX_CORPUS_DIR")
            .expect("VSDX_CORPUS_DIR must name the required VSDX corpus directory");
        let directory = std::path::PathBuf::from(directory);
        let files = ["lichtsysteme.vsdx", "soundplan.vsdx"];
        for file in files {
            assert!(
                directory.join(file).is_file(),
                "missing corpus file: {file}"
            );
        }
        let mut measurement = CorpusMeasurement::default();
        for file in files {
            let path = directory.join(file);
            let package = parse_vsdx(&fs::read(&path).expect("read corpus file"))
                .expect("parse corpus package");
            let resolver = Resolver::new(&package);
            let document = package
                .document_sheet
                .as_ref()
                .map(|sheet| resolver.resolve_sheet(sheet))
                .transpose()
                .expect("resolve corpus document sheet");
            for sheet in package
                .document_sheet
                .iter()
                .chain(package.style_sheets.iter())
                .chain(package.page_sheets.values())
                .chain(package.master_sheets.values())
            {
                let refs = sheet_references(sheet);
                let refs = DocumentReferences::new(&refs, document.as_ref());
                for (name, cell) in sheet_formula_cells(sheet) {
                    let formula = cell.formula.as_deref().unwrap();
                    let evaluation = evaluate_cell_with_package_theme(
                        &name,
                        formula,
                        &refs,
                        &limits(),
                        &package,
                    );
                    measurement.record(formula, evaluation.clone());
                    measurement.record_oracle(
                        &name,
                        formula,
                        cell.value
                            .as_deref()
                            .map(|value| (value, cell.unit.as_deref())),
                        &evaluation,
                    );
                }
            }
            for (page, sheet) in &package.page_contents {
                let refs = sheet_references(sheet);
                let refs = DocumentReferences::new(&refs, document.as_ref());
                for (name, cell) in sheet_formula_cells(sheet) {
                    let formula = cell.formula.as_deref().unwrap();
                    let evaluation = evaluate_cell_with_package_theme(
                        &name,
                        formula,
                        &refs,
                        &limits(),
                        &package,
                    );
                    measurement.record(formula, evaluation.clone());
                    measurement.record_oracle(
                        &name,
                        formula,
                        cell.value
                            .as_deref()
                            .map(|value| (value, cell.unit.as_deref())),
                        &evaluation,
                    );
                }
                let page_refs =
                    PageShapeReferences::new(&resolver, page).expect("resolve corpus page shapes");
                for shape in shapes(sheet) {
                    let refs = page_refs.for_shape(shape.id);
                    let resolved = page_refs.shape(shape.id).expect("resolve corpus shape");
                    for (name, cell) in shape_formula_cells(shape) {
                        let formula = cell.formula.as_deref().unwrap();
                        let evaluation = evaluate_cell_with_shape_package_theme(
                            &name,
                            formula,
                            &refs,
                            &limits(),
                            resolved,
                            &package,
                        );
                        measurement.record(formula, evaluation.clone());
                        measurement.record_oracle(
                            &name,
                            formula,
                            cell.value
                                .as_deref()
                                .map(|value| (value, cell.unit.as_deref())),
                            &evaluation,
                        );
                    }
                }
            }
            for sheet in package.master_contents.values() {
                let refs = sheet_references(sheet);
                let refs = DocumentReferences::new(&refs, document.as_ref());
                for (name, cell) in sheet_formula_cells(sheet) {
                    let formula = cell.formula.as_deref().unwrap();
                    let evaluation = evaluate_cell_with_package_theme(
                        &name,
                        formula,
                        &refs,
                        &limits(),
                        &package,
                    );
                    measurement.record(formula, evaluation.clone());
                    measurement.record_oracle(
                        &name,
                        formula,
                        cell.value
                            .as_deref()
                            .map(|value| (value, cell.unit.as_deref())),
                        &evaluation,
                    );
                }
                for shape in shapes(sheet) {
                    let resolved = resolver
                        .resolve_shape_in_sheet(shape, sheet)
                        .expect("resolve corpus master shape");
                    for (name, cell) in shape_formula_cells(shape) {
                        let formula = cell.formula.as_deref().unwrap();
                        let evaluation = evaluate_cell_with_shape_package_theme(
                            &name,
                            formula,
                            &DocumentReferences::new(&resolved, document.as_ref()),
                            &limits(),
                            &resolved,
                            &package,
                        );
                        measurement.record(formula, evaluation.clone());
                        measurement.record_oracle(
                            &name,
                            formula,
                            cell.value
                                .as_deref()
                                .map(|value| (value, cell.unit.as_deref())),
                            &evaluation,
                        );
                    }
                }
            }
        }
        eprintln!(
            "VSDX corpus formulas: parse_ast_ok={} static_known_unsupported={} outcomes: evaluated={} unsupported_known={} unsupported_other={} error={} total={}",
            measurement.ast_ok,
            measurement.static_known_unsupported,
            measurement.evaluated,
            measurement.unsupported_known,
            measurement.unsupported_other,
            measurement.error,
            measurement.total,
        );
        let oracle_compared = measurement.oracle_agreement + measurement.oracle_disagreement;
        let agreement_rate = if oracle_compared == 0 {
            0.0
        } else {
            measurement.oracle_agreement as f64 * 100.0 / oracle_compared as f64
        };
        eprintln!(
            "VSDX corpus @V oracle (Visio-produced, potentially stale): agreement={} disagreement={} no-cache-available={} agreement-rate={agreement_rate:.2}%",
            measurement.oracle_agreement,
            measurement.oracle_disagreement,
            measurement.oracle_no_cache,
        );
        measurement.oracle_disagreements.sort_by(|left, right| {
            left.name
                .cmp(&right.name)
                .then_with(|| left.formula.cmp(&right.formula))
                .then_with(|| left.computed.cmp(&right.computed))
                .then_with(|| left.cached.cmp(&right.cached))
        });
        measurement.oracle_disagreements.dedup_by(|left, right| {
            left.name == right.name
                && left.formula == right.formula
                && left.computed == right.computed
                && left.cached == right.cached
        });
        measurement.oracle_disagreements.truncate(20);
        for disagreement in &measurement.oracle_disagreements {
            eprintln!(
                "VSDX corpus @V disagreement: cell={} formula={:?} computed={} cached={}",
                disagreement.name, disagreement.formula, disagreement.computed, disagreement.cached,
            );
        }
        let mut top_unsupported = measurement
            .unsupported_names
            .into_iter()
            .collect::<Vec<_>>();
        top_unsupported
            .sort_by(|left, right| right.1.cmp(&left.1).then_with(|| left.0.cmp(&right.0)));
        eprintln!("VSDX corpus top unsupported constructs: {top_unsupported:?}");
        let mut top_unsupported_other = measurement
            .unsupported_other_kinds
            .into_iter()
            .collect::<Vec<_>>();
        top_unsupported_other
            .sort_by(|left, right| right.1.cmp(&left.1).then_with(|| left.0.cmp(&right.0)));
        top_unsupported_other.truncate(20);
        eprintln!("VSDX corpus top other unsupported kinds: {top_unsupported_other:?}");
        let mut top_errors = measurement.error_kinds.into_iter().collect::<Vec<_>>();
        top_errors.sort_by(|left, right| right.1.cmp(&left.1).then_with(|| left.0.cmp(&right.0)));
        eprintln!("VSDX corpus error histogram: {top_errors:?}");
        let mut top_references = measurement
            .unresolved_references
            .into_iter()
            .collect::<Vec<_>>();
        top_references
            .sort_by(|left, right| right.1.cmp(&left.1).then_with(|| left.0.cmp(&right.0)));
        top_references.truncate(20);
        eprintln!("VSDX corpus top unresolved references: {top_references:?}");
        assert_eq!(
            measurement.total, 6_992,
            "corpus formula denominator changed"
        );
        assert_eq!(
            measurement.evaluated
                + measurement.unsupported_known
                + measurement.unsupported_other
                + measurement.error,
            measurement.total,
            "corpus outcome buckets must be disjoint and exhaustive"
        );
    }

    #[derive(Default)]
    struct CorpusMeasurement {
        ast_ok: usize,
        static_known_unsupported: usize,
        evaluated: usize,
        unsupported_known: usize,
        unsupported_other: usize,
        error: usize,
        total: usize,
        unsupported_names: BTreeMap<String, usize>,
        unsupported_other_kinds: BTreeMap<String, usize>,
        error_kinds: BTreeMap<String, usize>,
        unresolved_references: BTreeMap<String, usize>,
        oracle_agreement: usize,
        oracle_disagreement: usize,
        oracle_no_cache: usize,
        oracle_disagreements: Vec<OracleDisagreement>,
    }
    struct OracleDisagreement {
        name: String,
        formula: String,
        computed: String,
        cached: String,
    }
    impl CorpusMeasurement {
        fn record(&mut self, formula: &str, evaluation: Evaluation) {
            self.total += 1;
            if let Ok(expression) = parse(formula, &limits()) {
                self.ast_ok += 1;
                if has_unsupported(&expression) {
                    self.static_known_unsupported += 1;
                    collect_unsupported(&expression, &mut self.unsupported_names);
                }
            }
            match evaluation {
                Evaluation::Evaluated(_) => self.evaluated += 1,
                Evaluation::Unsupported(reason) => {
                    if is_known_deferred_reason(&reason) {
                        self.unsupported_known += 1;
                    } else {
                        self.unsupported_other += 1;
                        *self.unsupported_other_kinds.entry(reason).or_default() += 1;
                    }
                }
                Evaluation::Error(error) => {
                    self.error += 1;
                    let kind = classify_error(&error.message, &mut self.unresolved_references);
                    *self.error_kinds.entry(kind).or_default() += 1;
                }
            }
        }
        fn record_oracle(
            &mut self,
            name: &str,
            formula: &str,
            cached: Option<(&str, Option<&str>)>,
            evaluation: &Evaluation,
        ) {
            let Evaluation::Evaluated(computed) = evaluation else {
                return;
            };
            let Some((cached, unit)) = cached else {
                self.oracle_no_cache += 1;
                return;
            };
            let Evaluation::Evaluated(cached_value) = cell_value(cached, unit) else {
                self.oracle_no_cache += 1;
                return;
            };
            if oracle_values_agree(computed.value, cached_value.value) {
                self.oracle_agreement += 1;
            } else {
                self.oracle_disagreement += 1;
                self.oracle_disagreements.push(OracleDisagreement {
                    name: name.into(),
                    formula: formula.into(),
                    computed: format_value(computed.value),
                    cached: cached.into(),
                });
            }
        }
    }

    fn oracle_values_agree(computed: Value, cached: Value) -> bool {
        match (computed, cached) {
            (Value::Number(computed), Value::Number(cached)) => {
                (computed.number - cached.number).abs()
                    <= 1e-9_f64.max(computed.number.abs().max(cached.number.abs()) * 1e-9)
            }
            (Value::Color(computed), Value::Color(cached)) => computed == cached,
            _ => false,
        }
    }

    fn format_value(value: Value) -> String {
        match value {
            Value::Number(value) => format!("{} {:?}", value.number, value.unit),
            Value::Color(value) => {
                format!("#{:02X}{:02X}{:02X}", value.red, value.green, value.blue)
            }
        }
    }

    #[test]
    fn corpus_measurement_uses_evaluation_outcomes_not_static_calls() {
        let refs = BTreeMap::new();
        let mut measurement = CorpusMeasurement::default();
        measurement.record(
            "IF(0,PNT(1,2),1)",
            evaluate("IF(0,PNT(1,2),1)", &refs, &limits()),
        );
        measurement.record("1/0+PNT(1,2)", evaluate("1/0+PNT(1,2)", &refs, &limits()));

        assert_eq!(measurement.static_known_unsupported, 2);
        assert_eq!(measurement.evaluated, 1);
        assert_eq!(measurement.error, 1);
        assert_eq!(measurement.unsupported_known, 0);
        assert_eq!(measurement.unsupported_other, 0);
    }

    #[test]
    fn corpus_policy_scanner_reports_local_formula_counts() {
        let Some(directory) = std::env::var_os("VSDX_CORPUS_DIR") else {
            eprintln!("SKIPPED CORPUS POLICY SCANNER: VSDX_CORPUS_DIR is unset");
            return;
        };
        let mut counts = CorpusPolicyCounts::default();
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let package =
                parse_vsdx(&fs::read(std::path::Path::new(&directory).join(file)).unwrap())
                    .unwrap();
            let mut sheets = package
                .document_sheet
                .iter()
                .chain(package.style_sheets.iter())
                .chain(package.page_sheets.values())
                .chain(package.master_sheets.values())
                .chain(package.page_contents.values())
                .chain(package.master_contents.values())
                .collect::<Vec<_>>();
            for sheet in sheets.drain(..) {
                scan_policy_formulas(sheet_formulas(sheet), &mut counts);
                for shape in shapes(sheet) {
                    scan_policy_formulas(shape_formulas(shape), &mut counts);
                }
            }
        }
        eprintln!(
            "VSDX corpus policy (local formulas only; inherited formulas excluded): guarded={} redirected_root={} redirected_nested_or_multiple={} lock_protected={}",
            counts.guarded,
            counts.redirected_root,
            counts.redirected_nested_or_multiple,
            counts.lock_protected
        );
    }

    #[derive(Default)]
    struct CorpusPolicyCounts {
        guarded: usize,
        redirected_root: usize,
        redirected_nested_or_multiple: usize,
        lock_protected: usize,
    }

    fn scan_policy_formulas(formulas: Vec<(String, &str)>, counts: &mut CorpusPolicyCounts) {
        let references = formulas
            .iter()
            .filter(|(_, formula)| !formula.eq_ignore_ascii_case("Inh"))
            .map(|(name, formula)| (name.clone(), (*formula).to_owned()))
            .collect::<BTreeMap<_, _>>();
        for (name, formula) in formulas {
            if formula.eq_ignore_ascii_case("Inh") {
                continue;
            }
            let Ok(expression) = parse(formula, &limits()) else {
                continue;
            };
            if contains_named_call(&expression, "GUARD") {
                counts.guarded += 1;
            }
            let redirects = named_call_count(&expression, "SETATREF");
            if redirects > 0 {
                if redirects == 1 && is_root_call(&expression, "SETATREF") {
                    counts.redirected_root += 1;
                } else {
                    counts.redirected_nested_or_multiple += 1;
                }
            }
            if name.starts_with("Lock")
                && matches!(
                    evaluate(formula, &references, &limits()),
                    Evaluation::Evaluated(Evaluated {
                        value: Value::Number(Number { number: 1.0, .. }),
                        ..
                    })
                )
            {
                counts.lock_protected += 1;
            }
        }
    }

    fn contains_named_call(expression: &Expr, name: &str) -> bool {
        named_call_count(expression, name) > 0
    }

    fn named_call_count(expression: &Expr, name: &str) -> usize {
        match expression {
            Expr::Call(current, arguments) => {
                usize::from(current.eq_ignore_ascii_case(name))
                    + arguments
                        .iter()
                        .map(|argument| named_call_count(argument, name))
                        .sum::<usize>()
            }
            Expr::Unary(value) => named_call_count(value, name),
            Expr::Binary(left, _, right) => {
                named_call_count(left, name) + named_call_count(right, name)
            }
            Expr::Number(_, _) | Expr::String(_) | Expr::Reference(_) => 0,
        }
    }

    fn is_root_call(expression: &Expr, name: &str) -> bool {
        matches!(expression, Expr::Call(current, _) if current.eq_ignore_ascii_case(name))
    }

    fn classify_error(
        message: &str,
        unresolved_references: &mut BTreeMap<String, usize>,
    ) -> String {
        if let Some(name) = message.strip_prefix("unresolved reference ") {
            *unresolved_references.entry(name.to_owned()).or_default() += 1;
            if name.starts_with("Sheet.")
                || matches!(name.split_once('!'), Some(("ThePage" | "TheDoc", _)))
            {
                return format!("unresolved cross-sheet reference: {name}");
            }
            return format!("unresolved cell reference: {name}");
        }
        if message == "colour used where a numeric value is required"
            || message == "numeric value used where a colour is required"
        {
            return format!("type error: {message}");
        }
        if message == "missing argument" || message.contains(" requires ") {
            return format!("arity error: {message}");
        }
        if message.contains("unit")
            || message.contains("dimensional")
            || message.contains("trigonometric argument")
        {
            return format!("unit/dimension error: {message}");
        }
        if message.contains("limit exceeded") {
            return format!("budget/depth/step exceeded: {message}");
        }
        format!("other: {message}")
    }

    fn sheet_formulas(sheet: &Sheet) -> Vec<(String, &str)> {
        let mut values = sheet
            .cells()
            .filter_map(|cell| {
                cell.formula
                    .as_deref()
                    .map(|formula| (cell.name.clone(), formula))
            })
            .collect::<Vec<_>>();
        for section in sheet.sections() {
            for row in section.rows() {
                values.extend(row.cells().filter_map(|cell| {
                    cell.formula
                        .as_deref()
                        .map(|formula| (section_cell_name(section, row, cell), formula))
                }));
            }
        }
        values
    }
    fn sheet_formula_cells(sheet: &Sheet) -> Vec<(String, &Cell)> {
        let mut values = sheet
            .cells()
            .filter(|cell| cell.formula.is_some())
            .map(|cell| (cell.name.clone(), cell))
            .collect::<Vec<_>>();
        for section in sheet.sections() {
            for row in section.rows() {
                values.extend(
                    row.cells()
                        .filter(|cell| cell.formula.is_some())
                        .map(|cell| (section_cell_name(section, row, cell), cell)),
                );
            }
        }
        values
    }
    fn shape_formulas(shape: &Shape) -> Vec<(String, &str)> {
        let mut values = shape
            .cells()
            .filter_map(|cell| {
                cell.formula
                    .as_deref()
                    .map(|formula| (cell.name.clone(), formula))
            })
            .collect::<Vec<_>>();
        for section in shape.sections() {
            for row in section.rows() {
                values.extend(row.cells().filter_map(|cell| {
                    cell.formula
                        .as_deref()
                        .map(|formula| (section_cell_name(section, row, cell), formula))
                }));
            }
        }
        values
    }
    fn shape_formula_cells(shape: &Shape) -> Vec<(String, &Cell)> {
        let mut values = shape
            .cells()
            .filter(|cell| cell.formula.is_some())
            .map(|cell| (cell.name.clone(), cell))
            .collect::<Vec<_>>();
        for section in shape.sections() {
            for row in section.rows() {
                values.extend(
                    row.cells()
                        .filter(|cell| cell.formula.is_some())
                        .map(|cell| (section_cell_name(section, row, cell), cell)),
                );
            }
        }
        values
    }
    fn section_cell_name(
        section: &vsdx_parse::Section,
        row: &vsdx_parse::Row,
        cell: &Cell,
    ) -> String {
        row.name.as_ref().map_or_else(
            || format!("{}.{}", section.name, cell.name),
            |row| format!("{}.{}.{}", section.name, row, cell.name),
        )
    }
    fn shapes(sheet: &Sheet) -> Vec<&Shape> {
        let mut values = Vec::new();
        for shape in sheet.shapes() {
            collect_shapes(shape, &mut values);
        }
        values
    }
    fn collect_shapes<'a>(shape: &'a Shape, values: &mut Vec<&'a Shape>) {
        values.push(shape);
        for child in &shape.children {
            if let vsdx_parse::ShapeChild::Shapes(children) = child {
                for child in children {
                    if let vsdx_parse::ShapesChild::Shape(shape) = child {
                        collect_shapes(shape, values);
                    }
                }
            }
        }
    }
    fn sheet_references(sheet: &Sheet) -> BTreeMap<String, String> {
        let mut refs = references(sheet.cells().map(|cell| (cell.name.clone(), cell)));
        for section in sheet.sections() {
            for row in section.rows() {
                for cell in row.cells() {
                    refs.extend(references(std::iter::once((
                        section_cell_name(section, row, cell),
                        cell,
                    ))));
                }
            }
        }
        refs
    }
    fn references<'a>(cells: impl Iterator<Item = (String, &'a Cell)>) -> BTreeMap<String, String> {
        cells
            .filter_map(|(name, cell)| cell.formula.as_ref().map(|formula| (name, formula.clone())))
            .collect()
    }
    fn has_unsupported(expression: &Expr) -> bool {
        match expression {
            Expr::Call(name, args) => {
                is_known_deferred_call(name) || args.iter().any(has_unsupported)
            }
            Expr::Unary(value) => has_unsupported(value),
            Expr::Binary(left, _, right) => has_unsupported(left) || has_unsupported(right),
            _ => false,
        }
    }
    fn is_known_deferred_call(name: &str) -> bool {
        !matches!(
            name.to_ascii_uppercase().as_str(),
            "IF" | "AND"
                | "OR"
                | "NOT"
                | "MIN"
                | "MAX"
                | "ABS"
                | "INT"
                | "ROUND"
                | "CEILING"
                | "FLOOR"
                | "SQRT"
                | "SIN"
                | "COS"
                | "TAN"
                | "ATAN2"
                | "PI"
                | "MOD"
                | "SUM"
                | "TRUNC"
                | "SIGN"
                | "RGB"
                | "TINT"
                | "MSOTINT"
                | "SAT"
                | "THEMEVAL"
                | "THEMEGUARD"
                | "_XFTRIGGER"
                | "GUARD"
        )
    }
    fn is_known_deferred_reason(reason: &str) -> bool {
        let name = reason
            .strip_prefix("unsupported function ")
            .or_else(|| reason.strip_suffix(" is not implemented"))
            .or_else(|| reason.strip_suffix(" is outside the phase-4 evaluator"));
        name.is_some_and(is_known_deferred_call)
    }
    fn collect_unsupported(expression: &Expr, counts: &mut BTreeMap<String, usize>) {
        match expression {
            Expr::Call(name, args) => {
                if has_unsupported(&Expr::Call(name.clone(), Vec::new())) {
                    *counts.entry(name.to_ascii_uppercase()).or_default() += 1;
                }
                for argument in args {
                    collect_unsupported(argument, counts);
                }
            }
            Expr::Unary(value) => collect_unsupported(value, counts),
            Expr::Binary(left, _, right) => {
                collect_unsupported(left, counts);
                collect_unsupported(right, counts);
            }
            _ => {}
        }
    }
}
