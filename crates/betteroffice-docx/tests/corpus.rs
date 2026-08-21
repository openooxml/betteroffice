//! Corpus gate: every manifest-pinned fixture through the unedited round-trip
//! oracles, its findings pinned exactly. A finding that appears or disappears
//! without regenerating the golden under `GOLDEN_UPDATE=1` fails, so drift
//! stays a deliberate act and a fixed defect must claim its fix.
//! Governed by `openspec/changes/docx-word-fidelity/specs/fidelity-corpus`.

mod common;

use std::collections::BTreeSet;
use std::path::PathBuf;

use common::{parts_of, roundtrip_report, save_unedited};
use sha2::{Digest, Sha256};

const CORPUS_FLOOR: usize = 2;

fn corpus_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/corpus")
}

fn manifest_fixtures() -> Vec<serde_json::Value> {
    let manifest: serde_json::Value =
        serde_json::from_str(&std::fs::read_to_string(corpus_dir().join("manifest.json")).unwrap())
            .unwrap();
    manifest["fixtures"].as_array().unwrap().clone()
}

#[test]
fn the_corpus_is_pinned_by_the_manifest() {
    let fixtures = manifest_fixtures();
    assert!(fixtures.len() >= CORPUS_FLOOR, "corpus below its floor");
    let mut listed = BTreeSet::new();
    for entry in &fixtures {
        let file = entry["file"].as_str().unwrap();
        listed.insert(file.to_owned());
        let bytes = std::fs::read(corpus_dir().join("fixtures").join(file)).unwrap();
        assert_eq!(
            entry["sha256"].as_str().unwrap(),
            format!("{:x}", Sha256::digest(&bytes)),
            "{file}: checked-in bytes do not match the manifest hash"
        );
        assert!(!entry["provenance"].as_str().unwrap().is_empty());
        assert_eq!(entry["oracle"].as_str().unwrap(), "roundtrip-findings");
        assert!(!entry["features"].as_array().unwrap().is_empty());
    }
    for file in std::fs::read_dir(corpus_dir().join("fixtures")).unwrap() {
        let name = file.unwrap().file_name().to_string_lossy().to_string();
        assert!(listed.contains(&name), "{name} has no manifest entry");
    }
}

#[test]
fn every_fixture_round_trips_within_its_pinned_findings() {
    let fixtures = manifest_fixtures();
    assert!(fixtures.len() >= CORPUS_FLOOR, "corpus below its floor");
    let update = std::env::var("GOLDEN_UPDATE").is_ok_and(|value| value == "1");
    let mut failures = Vec::new();
    for entry in &fixtures {
        let file = entry["file"].as_str().unwrap();
        let original = std::fs::read(corpus_dir().join("fixtures").join(file)).unwrap();
        let saved = save_unedited(&original);
        let mut findings = roundtrip_report(&parts_of(&original), &parts_of(&saved));
        let resaved = save_unedited(&saved);
        if resaved != saved {
            let second: std::collections::BTreeMap<String, Vec<u8>> =
                parts_of(&resaved).into_iter().collect();
            for (name, first) in parts_of(&saved) {
                if second.get(&name).is_none_or(|bytes| *bytes != first) {
                    findings.push(format!("fixed point: unstable part {name}"));
                }
            }
        }
        let expected_path = corpus_dir()
            .join("expected")
            .join(format!("{file}.findings.txt"));
        let actual = if findings.is_empty() {
            String::new()
        } else {
            findings.join("\n") + "\n"
        };
        if update {
            std::fs::create_dir_all(expected_path.parent().unwrap()).unwrap();
            std::fs::write(&expected_path, &actual).unwrap();
            continue;
        }
        let expected = std::fs::read_to_string(&expected_path).unwrap_or_default();
        if actual != expected {
            let expected_count = expected.lines().count();
            failures.push(format!(
                "{file}: {} finding(s) but {} pinned; first divergence:\n  pinned: {}\n  actual: {}\n\
                 regenerate deliberately with GOLDEN_UPDATE=1 and review the diff",
                findings.len(),
                expected_count,
                expected
                    .lines()
                    .zip(actual.lines().chain(std::iter::repeat("absent")))
                    .find(|(e, a)| e != a)
                    .map(|(e, _)| e)
                    .unwrap_or("absent"),
                actual
                    .lines()
                    .zip(expected.lines().chain(std::iter::repeat("absent")))
                    .find(|(a, e)| a != e)
                    .map(|(a, _)| a)
                    .unwrap_or("absent"),
            ));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n\n"));
}
