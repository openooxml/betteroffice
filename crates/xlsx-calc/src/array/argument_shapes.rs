//! every builtin the engine resolves must say how array evaluation treats
//! its arguments: lifted element by element (all of them via `lifts`, some
//! via `lifted_positions`), implemented over arrays in `lookup_array`, or
//! listed below as taking its arguments whole. a new function that is in
//! none of these fails here instead of silently answering once.

use super::{lifted_positions, lifts, lookup_array};

/// builtins evaluated once with their arguments as written: no arguments,
/// ranges or arrays read whole (aggregates, regressions, matrices, lookups,
/// holiday lists), references (OFFSET, INDIRECT, AREAS, FORMULATEXT), branches
/// only evaluated when chosen (CHOOSE, SWITCH), and the parser's union.
const WHOLE_ARGUMENTS: &[&str] = &[
    "(UNION)",
    "AGGREGATE",
    "AND",
    "AREAS",
    "AVEDEV",
    "AVERAGEA",
    "CHOOSE",
    "CORREL",
    "COUNTBLANK",
    "COVAR",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "DEVSQ",
    "FALSE",
    "FORMULATEXT",
    "GCD",
    "GEOMEAN",
    "HARMEAN",
    "IMPRODUCT",
    "IMSUM",
    "INDIRECT",
    "INTERCEPT",
    "KURT",
    "LCM",
    "LOOKUP",
    "MAXA",
    "MDETERM",
    "MEDIAN",
    "MINA",
    "MODE",
    "MODE.SNGL",
    "MULTINOMIAL",
    "NA",
    "NETWORKDAYS",
    "NETWORKDAYS.INTL",
    "NOW",
    "NPV",
    "OFFSET",
    "OR",
    "PERCENTOF",
    "PI",
    "RAND",
    "RSQ",
    "SERIESSUM",
    "SKEW",
    "SLOPE",
    "STDEV",
    "STDEV.P",
    "STDEV.S",
    "STDEVA",
    "STDEVP",
    "SUBTOTAL",
    "SUMSQ",
    "SUMX2MY2",
    "SUMX2PY2",
    "SUMXMY2",
    "SWITCH",
    "TODAY",
    "TRIMMEAN",
    "TRUE",
    "VAR",
    "VAR.P",
    "VAR.S",
    "VARP",
    "WORKDAY",
    "WORKDAY.INTL",
    "XOR",
];

/// scalar builtins excel lifts element by element that this engine still
/// evaluates once; listed so the gap is visible rather than accidental.
const NOT_YET_LIFTED: &[&str] = &[
    "ADDRESS",
    "CONVERT",
    "DATEVALUE",
    "DAYS",
    "FV",
    "HYPERLINK",
    "NORM.DIST",
    "NORMDIST",
    "NPER",
    "PMT",
    "PV",
    "RANDBETWEEN",
    "TIMEVALUE",
    "YEARFRAC",
];

const FUNCTIONS_SOURCE: &str = include_str!("../functions/mod.rs");

fn block<'a>(source: &'a str, start: &str) -> &'a str {
    let from = source.find(start).expect("block start");
    let rest = &source[from..];
    &rest[..rest.find("\n}\n").expect("block end")]
}

/// every name `resolve` maps, aliases included.
fn resolved_names() -> Vec<String> {
    block(FUNCTIONS_SOURCE, "pub fn resolve(")
        .lines()
        .map(str::trim)
        .filter(|line| line.starts_with('"') && line.contains("=> Func::"))
        .flat_map(|line| {
            let names = line.split("=>").next().unwrap_or_default();
            names
                .split('|')
                .map(|name| name.trim().trim_matches('"').to_string())
                .collect::<Vec<_>>()
        })
        .collect()
}

/// the `Func` variants, and the variants `resolve` reaches.
fn variants() -> (Vec<String>, Vec<String>) {
    let declared = block(FUNCTIONS_SOURCE, "pub enum Func {")
        .lines()
        .skip(1)
        .map(str::trim)
        .filter(|line| !line.starts_with("//"))
        .filter_map(|line| line.strip_suffix(','))
        .map(str::to_string)
        .collect();
    let reached = block(FUNCTIONS_SOURCE, "pub fn resolve(")
        .lines()
        .filter_map(|line| line.split("=> Func::").nth(1))
        .map(|rest| rest.trim_end_matches(',').trim().to_string())
        .collect();
    (declared, reached)
}

#[test]
fn every_variant_has_a_resolvable_name() {
    let (declared, reached) = variants();
    assert!(declared.len() > 200, "parsed {} variants", declared.len());
    for variant in &declared {
        assert!(reached.contains(variant), "Func::{variant} has no name");
    }
}

#[test]
fn every_builtin_has_an_array_argument_shape() {
    let names = resolved_names();
    assert!(names.len() > 200, "parsed {} names", names.len());
    for name in &names {
        let lifted = lifts(name) || lifted_positions(name).is_some();
        let native = lookup_array(name).is_some();
        let listed =
            WHOLE_ARGUMENTS.contains(&name.as_str()) || NOT_YET_LIFTED.contains(&name.as_str());
        assert!(
            lifted || native || listed,
            "{name} has no array argument shape: add it to `lifts`, \
             `lifted_positions` or WHOLE_ARGUMENTS"
        );
    }
}

#[test]
fn the_lists_name_only_unlifted_builtins() {
    let names = resolved_names();
    for name in WHOLE_ARGUMENTS.iter().chain(NOT_YET_LIFTED) {
        assert!(names.iter().any(|known| known == name), "{name} is unknown");
        assert!(
            !lifts(name) && lifted_positions(name).is_none() && lookup_array(name).is_none(),
            "{name} is listed but also lifted or array-native"
        );
    }
    for name in WHOLE_ARGUMENTS {
        assert!(!NOT_YET_LIFTED.contains(name), "{name} is listed twice");
    }
}
