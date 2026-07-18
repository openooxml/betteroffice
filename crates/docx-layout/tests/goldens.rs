//! Native golden-corpus harness — no wasm, no browser.
//!
//! Replays every fixture pair exported by `scripts/export-golden-fixtures.ts`
//! through the pure engine. A scenario the engine accepts MUST reproduce the
//! committed golden canonical layout byte-for-byte; a scenario that engages a
//! not-yet-ported feature may report UNSUPPORTED — except the required
//! paragraph scenarios, which must pass. Prints a one-line PASS/UNSUPPORTED
//! table (visible via `cargo test -- --nocapture`, or on failure).

use std::fs;
use std::path::PathBuf;

/// Scenarios the spine must handle (kept in sync with `SUPPORTED` in
/// `packages/core/src/layout/pagination/__golden__/rustParity.test.ts`).
const REQUIRED: [&str; 2] = [
    "single-page-multi-paragraph",
    "multi-page-paragraph-overflow",
];

#[test]
fn golden_corpus_byte_identity() {
    let fixtures = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut names: Vec<String> = fs::read_dir(&fixtures)
        .expect("fixtures dir exists — run `bun scripts/export-golden-fixtures.ts`")
        .filter_map(|entry| {
            let name = entry.ok()?.file_name().into_string().ok()?;
            name.strip_suffix(".input.json").map(str::to_string)
        })
        .collect();
    names.sort();
    assert!(!names.is_empty(), "no fixtures found in {fixtures:?}");
    for required in REQUIRED {
        assert!(
            names.iter().any(|n| n == required),
            "required scenario `{required}` has no fixture"
        );
    }

    let mut failures: Vec<String> = Vec::new();
    let width = names.iter().map(|n| n.len()).max().unwrap_or(0);

    println!("\n golden corpus — rust spine");
    for name in &names {
        let input = fs::read_to_string(fixtures.join(format!("{name}.input.json")))
            .expect("fixture input readable");
        let golden = fs::read_to_string(fixtures.join(format!("{name}.golden.json")))
            .expect("fixture golden readable");

        match docx_layout::layout_to_canonical_json(&input) {
            Ok(canonical) => {
                if canonical == golden {
                    println!("  {name:width$}  PASS (byte-identical)");
                } else {
                    println!("  {name:width$}  FAIL (output differs from golden)");
                    failures.push(format!(
                        "{name}: engine accepted the input but output differs from the golden\n\
                         --- golden (first divergence context) ---\n{}\n\
                         --- rust ---\n{}",
                        first_divergence(&golden, &canonical).0,
                        first_divergence(&golden, &canonical).1,
                    ));
                }
            }
            Err(err) => {
                println!("  {name:width$}  UNSUPPORTED ({err})");
                if REQUIRED.contains(&name.as_str()) {
                    failures.push(format!("{name}: required scenario returned {err}"));
                }
            }
        }
    }
    println!();

    assert!(failures.is_empty(), "\n{}", failures.join("\n\n"));
}

/// A few lines of context around the first differing line of two texts.
fn first_divergence(expected: &str, actual: &str) -> (String, String) {
    let expected_lines: Vec<&str> = expected.lines().collect();
    let actual_lines: Vec<&str> = actual.lines().collect();
    let first_diff = expected_lines
        .iter()
        .zip(actual_lines.iter())
        .position(|(a, b)| a != b)
        .unwrap_or(expected_lines.len().min(actual_lines.len()));
    let context = |lines: &[&str]| {
        let start = first_diff.saturating_sub(2);
        let end = (first_diff + 3).min(lines.len());
        lines[start..end]
            .iter()
            .enumerate()
            .map(|(i, l)| format!("{:>5} | {l}", start + i + 1))
            .collect::<Vec<_>>()
            .join("\n")
    };
    (context(&expected_lines), context(&actual_lines))
}
