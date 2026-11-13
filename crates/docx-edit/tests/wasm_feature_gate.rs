//! CI's `--features wasm` pass runs only the lib tests under the modules that the
//! `wasm` feature adds. These tests fail when code elsewhere in the crate starts to
//! depend on the feature, since the plain `cargo test` would then no longer cover
//! what the wasm build runs.

use std::path::{Path, PathBuf};

const WASM_MODULES: [&str; 2] = ["wasm", "wasm_memory"];
const NEEDLE: &str = concat!("feature = \"", "wasm\"");
const WASM_PASS: &str = "cargo test -p betteroffice-docx-edit --features wasm";

fn rust_files(dir: &Path, files: &mut Vec<PathBuf>) {
    for entry in std::fs::read_dir(dir).unwrap() {
        let path = entry.unwrap().path();
        if path.is_dir() {
            rust_files(&path, files);
        } else if path.extension().is_some_and(|ext| ext == "rs") {
            files.push(path);
        }
    }
}

fn inside_wasm_module(src: &Path, path: &Path) -> bool {
    path.strip_prefix(src)
        .ok()
        .and_then(|relative| relative.components().next())
        .and_then(|first| first.as_os_str().to_str())
        .is_some_and(|first| {
            let name = first.strip_suffix(".rs").unwrap_or(first);
            WASM_MODULES.contains(&name)
        })
}

fn allowed(previous: &str, line: &str, next: &str) -> bool {
    let (previous, line, next) = (previous.trim(), line.trim(), next.trim());
    let lint_attr = format!("cfg_attr(not({NEEDLE}), allow(dead_code))");
    let wasm32_only = format!("all({NEEDLE}, target_arch = \"wasm32\")");
    let module_attr = format!("#[cfg({NEEDLE})]");
    let helper_attr = format!("#[cfg(all(test, {NEEDLE}))]");
    let declared = next
        .strip_prefix("pub mod ")
        .and_then(|rest| rest.strip_suffix(';'));
    let gated_module =
        line == module_attr && declared.is_some_and(|module| WASM_MODULES.contains(&module));
    let test_helper = line == helper_attr
        && !previous.starts_with("#[")
        && !next.starts_with("#[")
        && !next.contains("mod ")
        && next.contains("fn ");
    line.contains(lint_attr.as_str())
        || line.contains(wasm32_only.as_str())
        || gated_module
        || test_helper
}

#[test]
fn only_the_wasm_modules_change_under_the_wasm_feature() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let src = root.join("src");
    let mut files = Vec::new();
    rust_files(&src, &mut files);
    rust_files(&root.join("tests"), &mut files);
    let mut offending = Vec::new();
    for path in files.iter().filter(|path| !inside_wasm_module(&src, path)) {
        let text = std::fs::read_to_string(path).unwrap();
        let lines: Vec<&str> = text.lines().collect();
        for (index, line) in lines.iter().enumerate() {
            if !line.contains(NEEDLE) {
                continue;
            }
            let previous = index.checked_sub(1).map_or("", |at| lines[at]);
            let next = lines.get(index + 1).copied().unwrap_or("");
            if !allowed(previous, line, next) {
                offending.push(format!("{}:{}: {}", path.display(), index + 1, line.trim()));
            }
        }
    }
    assert!(
        offending.is_empty(),
        "tests that need the wasm feature belong under src/{{{}}}, which CI's wasm pass \
         selects; any other cfg on the feature changes code that pass does not run:\n{}",
        WASM_MODULES.join(","),
        offending.join("\n")
    );
}

#[test]
fn the_ci_wasm_pass_selects_every_wasm_module() {
    let workflow = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../.github/workflows/ci.yml");
    let Ok(workflow) = std::fs::read_to_string(workflow) else {
        return;
    };
    let pass = workflow
        .lines()
        .find(|line| line.contains(WASM_PASS))
        .unwrap_or_else(|| panic!("ci.yml has no `{WASM_PASS}` line"));
    let words: Vec<&str> = pass.split_whitespace().collect();
    assert!(words.contains(&"--lib"), "{pass}");
    for module in WASM_MODULES {
        let filter = format!("{module}::");
        assert!(words.contains(&filter.as_str()), "{pass} lacks {filter}");
    }
}
