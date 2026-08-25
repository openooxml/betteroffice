//! The ledger matches the catalogue, every claim names an existing test,
//! and every measured number recomputes exactly; regeneration refuses a
//! number that moved the wrong way.
//! Governed by `openspec/changes/docx-word-fidelity/specs/fidelity-scorecard`.

use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

fn workspace_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .unwrap()
}

fn ledger_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/scorecard/ledger.json")
}

fn ledger() -> serde_json::Value {
    serde_json::from_str(&std::fs::read_to_string(ledger_path()).unwrap()).unwrap()
}

/// Axis vocabularies, weakest first.
const VOCABULARIES: &[(&str, &[&str])] = &[
    ("preserve", &["none", "partial", "gated"]),
    ("model", &["generic", "partial", "typed"]),
    ("edit", &["none", "partial", "word-parity"]),
    ("layout", &["none", "partial", "pinned"]),
    ("paint", &["none", "partial", "golden"]),
    ("evidence", &["none", "spec", "word"]),
];

fn rank(axis: &str, value: &str) -> usize {
    let (_, vocabulary) = VOCABULARIES
        .iter()
        .find(|(name, _)| *name == axis)
        .unwrap_or_else(|| panic!("{axis} has no vocabulary"));
    vocabulary
        .iter()
        .position(|entry| *entry == value)
        .unwrap_or_else(|| panic!("{value:?} is not a {axis} status"))
}

fn criteria(ledger: &serde_json::Value) -> &Vec<serde_json::Value> {
    ledger["criteria"].as_array().unwrap()
}

/// Ids from the normative catalogue tables in the scorecard spec.
fn catalogue_ids() -> BTreeSet<String> {
    let spec = std::fs::read_to_string(
        workspace_root()
            .join("openspec/changes/docx-word-fidelity/specs/fidelity-scorecard/spec.md"),
    )
    .unwrap();
    let mut ids = BTreeSet::new();
    for line in spec.lines() {
        let Some(rest) = line.strip_prefix("| ") else {
            continue;
        };
        let Some((cell, _)) = rest.split_once(" |") else {
            continue;
        };
        let looks_like_id = cell.split_once('.').is_some_and(|(area, key)| {
            !area.is_empty()
                && !key.is_empty()
                && area.chars().all(|character| character.is_ascii_lowercase())
                && key
                    .chars()
                    .all(|character| character.is_ascii_lowercase() || character == '-')
        });
        if looks_like_id {
            ids.insert(cell.to_owned());
        }
    }
    ids
}

#[test]
fn the_ledger_and_the_catalogue_cannot_diverge() {
    let catalogue = catalogue_ids();
    assert!(catalogue.len() > 100, "catalogue tables were not found");
    let ledger = ledger();
    let ledger_ids: BTreeSet<String> = criteria(&ledger)
        .iter()
        .map(|criterion| criterion["id"].as_str().unwrap().to_owned())
        .collect();
    let missing: Vec<_> = catalogue.difference(&ledger_ids).collect();
    let extra: Vec<_> = ledger_ids.difference(&catalogue).collect();
    assert!(
        missing.is_empty() && extra.is_empty(),
        "ledger drifted from the catalogue; missing {missing:?}, extra {extra:?}"
    );
}

#[test]
fn preservation_has_no_target_exemptions() {
    let ledger = ledger();
    for criterion in criteria(&ledger) {
        let target = criterion["targets"]["preserve"].as_str().unwrap();
        assert!(
            target == "gated" || target == "n/a",
            "{}: preserve target must be gated",
            criterion["id"]
        );
    }
}

/// A status above its axis floor must name at least one gate, and every
/// named gate of the form `betteroffice-docx <file>::<test>` must exist.
#[test]
fn every_claimed_status_names_a_test_that_exists() {
    let floors = [
        ("preserve", "none"),
        ("model", "generic"),
        ("edit", "none"),
        ("layout", "none"),
        ("paint", "none"),
        ("evidence", "none"),
    ];
    let ledger = ledger();
    for criterion in criteria(&ledger) {
        let id = criterion["id"].as_str().unwrap();
        for (axis, floor) in floors {
            let status = criterion["statuses"][axis].as_str().unwrap();
            if status == floor || status == "n/a" {
                continue;
            }
            let gates = criterion["gates"][axis]
                .as_array()
                .unwrap_or_else(|| panic!("{id}.{axis} claims {status} without gates"));
            assert!(
                !gates.is_empty(),
                "{id}.{axis} claims {status} without gates"
            );
            for gate in gates {
                assert_gate_exists(gate.as_str().unwrap(), id);
            }
        }
    }
}

fn assert_gate_exists(gate: &str, criterion: &str) {
    let Some(rest) = gate.strip_prefix("betteroffice-docx ") else {
        panic!("{criterion}: unrecognized gate form {gate:?}");
    };
    let Some((file, function)) = rest.split_once("::") else {
        panic!("{criterion}: unrecognized gate form {gate:?}");
    };
    let source = std::fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/{file}.rs")),
    )
    .unwrap_or_else(|_| panic!("{criterion}: gate file tests/{file}.rs does not exist"));
    assert!(
        source.contains(&format!("fn {function}(")),
        "{criterion}: gate test {function} does not exist in tests/{file}.rs"
    );
}

/// A target its status already outranks is a goal that was passed, not a gap.
#[test]
fn no_status_outranks_its_target() {
    let ledger = ledger();
    for criterion in criteria(&ledger) {
        let id = criterion["id"].as_str().unwrap();
        for (axis, _) in VOCABULARIES {
            let status = criterion["statuses"][axis].as_str().unwrap();
            let target = criterion["targets"][axis].as_str().unwrap();
            if status == "n/a" || target == "n/a" {
                continue;
            }
            assert!(
                rank(axis, status) <= rank(axis, target),
                "{id}.{axis}: status {status} outranks target {target}"
            );
        }
    }
}

/// A defect nobody probes is headroom, not a measurement.
#[test]
fn every_defect_names_a_ceiling_that_exists() {
    let ledger = ledger();
    for criterion in criteria(&ledger) {
        let id = criterion["id"].as_str().unwrap();
        if criterion["defects"].as_array().unwrap().is_empty() {
            continue;
        }
        let ceilings = criterion["ceilings"]
            .as_array()
            .unwrap_or_else(|| panic!("{id} enumerates defects but names no ceiling"));
        assert!(
            !ceilings.is_empty(),
            "{id} enumerates defects but names no ceiling"
        );
        for ceiling in ceilings {
            assert_gate_exists(ceiling.as_str().unwrap(), id);
        }
    }
}

/// Measurements improve upward.
const IMPROVES_UPWARD: &[&str] = &[
    "corpusFixtures",
    "corpusFixturesWithoutFindings",
    "corpusFloor",
    "layoutGoldensRequired",
];

/// Measurements improve downward.
const IMPROVES_DOWNWARD: &[&str] = &["knownDefects"];

/// Each measurement recomputed from the artifact that defines it, in the order
/// the ledger records them. `declaredNormalizations` is a count, not a score:
/// it carries no direction, since a reviewed tolerance may legitimately join.
fn recomputed_measurements() -> Vec<(&'static str, u64)> {
    let manifest: serde_json::Value = serde_json::from_str(
        &std::fs::read_to_string(
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/corpus/manifest.json"),
        )
        .unwrap(),
    )
    .unwrap();
    let fixtures = manifest["fixtures"].as_array().unwrap();
    let without_findings = fixtures
        .iter()
        .filter(|entry| {
            let file = entry["file"].as_str().unwrap();
            let findings = Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/corpus/expected")
                .join(format!("{file}.findings.txt"));
            std::fs::read_to_string(findings)
                .unwrap_or_default()
                .trim()
                .is_empty()
        })
        .count();
    let ledger = ledger();
    let defects: usize = criteria(&ledger)
        .iter()
        .map(|criterion| criterion["defects"].as_array().unwrap().len())
        .sum();
    vec![
        ("corpusFixtures", fixtures.len() as u64),
        ("corpusFixturesWithoutFindings", without_findings as u64),
        (
            "corpusFloor",
            pinned_number(
                &Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/corpus.rs"),
                "CORPUS_FLOOR: usize = ",
                ";",
            ),
        ),
        (
            "declaredNormalizations",
            ooxml_fidelity::DECLARED_NORMALIZATIONS.len() as u64,
        ),
        ("knownDefects", defects as u64),
        (
            "layoutGoldensRequired",
            pinned_number(
                &workspace_root().join("crates/docx-layout/tests/goldens.rs"),
                "REQUIRED: [&str; ",
                "]",
            ),
        ),
    ]
}

/// A count another crate pins as a literal, read from its source.
fn pinned_number(source: &Path, prefix: &str, terminator: &str) -> u64 {
    let text = std::fs::read_to_string(source).unwrap();
    text.split(prefix)
        .nth(1)
        .and_then(|rest| rest.split(terminator).next())
        .unwrap_or_else(|| panic!("{} no longer pins {prefix:?}", source.display()))
        .trim()
        .parse()
        .unwrap()
}

#[test]
fn every_measured_number_recomputes_exactly() {
    let ledger = ledger();
    let recorded: BTreeMap<&str, u64> = ledger["measured"]
        .as_object()
        .unwrap()
        .iter()
        .map(|(name, value)| (name.as_str(), value.as_u64().unwrap()))
        .collect();
    let recomputed = recomputed_measurements();
    if std::env::var("GOLDEN_UPDATE").is_ok_and(|value| value == "1") {
        regenerate_measurements(&recomputed, &recorded);
        return;
    }
    assert_eq!(
        recorded,
        recomputed.iter().copied().collect::<BTreeMap<_, _>>(),
        "measured numbers drifted; regenerate deliberately with GOLDEN_UPDATE=1"
    );
}

/// Regeneration is the only moment a measurement moves, so it is where the
/// ratchet bites: a number that moved away from its target refuses to be
/// written until the regression is recorded with the issue tracking it.
fn regenerate_measurements(recomputed: &[(&str, u64)], recorded: &BTreeMap<&str, u64>) {
    for (name, value) in recomputed {
        let Some(previous) = recorded.get(name) else {
            continue;
        };
        let regressed = (IMPROVES_UPWARD.contains(name) && value < previous)
            || (IMPROVES_DOWNWARD.contains(name) && value > previous);
        assert!(
            !regressed,
            "{name} regressed from {previous} to {value}; record the regression and its \
             tracking issue in the entry before regenerating"
        );
    }
    let block = recomputed
        .iter()
        .map(|(name, value)| format!("    \"{name}\": {value}"))
        .collect::<Vec<_>>()
        .join(",\n");
    let source = std::fs::read_to_string(ledger_path()).unwrap();
    let start = source.find("  \"measured\": {").unwrap();
    let end = start + source[start..].find("\n  },").unwrap();
    std::fs::write(
        ledger_path(),
        format!(
            "{}  \"measured\": {{\n{block}{}",
            &source[..start],
            &source[end..]
        ),
    )
    .unwrap();
}
