use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

/// How production code reads the list a paragraph is on: its `numPr.numId`, or a list state
/// over numbering definitions.
const PATTERNS: [&str; 3] = ["\"numId\"", "ListState::new(", ".numbering()"];

/// Each source file and pattern allowed outside test code, with why.
const ALLOWED: [(&str, &str, &str); 4] = [
    (
        "list_definitions.rs",
        "\"numId\"",
        "the one resolver of a paragraph's list, and the numPr a created list writes",
    ),
    (
        "list_definitions.rs",
        "ListState::new(",
        "the one list state over the effective definitions, source and created",
    ),
    (
        "list_definitions.rs",
        ".numbering()",
        "the effective definitions start from the source package's",
    ),
    (
        "seed.rs",
        "\"numId\"",
        "seeding reads the source package's numPr, not a paragraph's attributes",
    ),
];

fn sources(directory: &Path, found: &mut Vec<PathBuf>) {
    for entry in std::fs::read_dir(directory).unwrap() {
        let path = entry.unwrap().path();
        if path.is_dir() {
            if path.file_name().is_some_and(|name| name != "tests") {
                sources(&path, found);
            }
        } else if path.extension().is_some_and(|extension| extension == "rs")
            && !path.to_string_lossy().ends_with("tests.rs")
        {
            found.push(path);
        }
    }
}

/// `source` without its top-level `#[cfg(test)]` modules.
fn production(source: &str) -> String {
    let mut kept = String::new();
    let mut lines = source.lines().peekable();
    while let Some(line) = lines.next() {
        if line == "#[cfg(test)]"
            && lines
                .peek()
                .is_some_and(|next| next.starts_with("mod ") && next.ends_with('{'))
        {
            for skipped in lines.by_ref() {
                if skipped == "}" {
                    break;
                }
            }
            continue;
        }
        kept.push_str(line);
        kept.push('\n');
    }
    kept
}

#[test]
fn only_the_list_resolver_reads_the_list_a_paragraph_is_on() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("src");
    let mut files = Vec::new();
    sources(&root, &mut files);
    let mut found = BTreeSet::new();
    for path in files {
        let text = production(&std::fs::read_to_string(&path).unwrap());
        let file = path
            .strip_prefix(&root)
            .unwrap()
            .to_string_lossy()
            .replace('\\', "/");
        for pattern in PATTERNS {
            if text.contains(pattern) {
                found.insert((file.clone(), pattern));
            }
        }
    }
    let allowed: BTreeSet<(String, &str)> = ALLOWED
        .iter()
        .map(|&(file, pattern, _)| (file.to_owned(), pattern))
        .collect();
    let unexpected: Vec<_> = found.difference(&allowed).collect();
    let absent: Vec<_> = allowed.difference(&found).collect();
    assert!(
        unexpected.is_empty() && absent.is_empty(),
        "read outside the resolver: {unexpected:?}; allowed but absent: {absent:?}"
    );
}
