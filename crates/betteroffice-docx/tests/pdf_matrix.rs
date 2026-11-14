#![cfg(all(feature = "raster", feature = "fonts"))]

//! End-to-end smoke test for the batch PDF exporter.
//!
//! This runs the compiled example over a small representative fixture set, so
//! it is ignored by default:
//!
//! ```text
//! cargo +1.99.1 test -p betteroffice-docx --test pdf_matrix -- --ignored --nocapture
//! ```

use std::fs;
use std::path::Path;
use std::process::Command;

use serde_json::Value;

const FIXTURES: &[&str] = &[
    "crates/betteroffice-docx/tests/corpus/fixtures/betteroffice-demo.docx",
    "crates/docx-edit/tests/fixtures/header-footer-reservation/even-and-odd-headers.docx",
    "crates/docx-edit/tests/fixtures/footnote-anchor.docx",
    "crates/docx-edit/tests/fixtures/image-line-spacing/img-body-360.docx",
];

#[test]
#[ignore = "runs the batch exporter over real fixtures"]
fn batch_exports_representative_fixtures() {
    let workspace = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("workspace root");
    let work = workspace
        .join("target")
        .join(format!("pdf-matrix-test-{}", std::process::id()));
    let input = work.join("fixtures");
    let output = work.join("pdfs");
    if work.exists() {
        fs::remove_dir_all(&work).expect("clear previous run");
    }
    fs::create_dir_all(&input).expect("create fixture directory");

    for relative in FIXTURES {
        copy_fixture(&workspace, Path::new(relative), &input);
    }

    let status = Command::new(env!("CARGO"))
        .current_dir(&workspace)
        .args([
            "run",
            "--quiet",
            "-p",
            "betteroffice-docx",
            "--example",
            "export_pdf_batch",
            "--",
        ])
        .arg(&input)
        .arg(&output)
        .status()
        .expect("run export_pdf_batch");
    assert!(status.success(), "export_pdf_batch exited {status}");

    let manifest_path = output.join("manifest.json");
    let manifest: Value =
        serde_json::from_str(&fs::read_to_string(&manifest_path).expect("read manifest.json"))
            .expect("parse manifest.json");
    let fixtures = manifest["fixtures"].as_array().expect("fixtures array");
    assert_eq!(fixtures.len(), FIXTURES.len(), "manifest: {manifest:#}");

    for fixture in fixtures {
        let path = fixture["path"].as_str().expect("fixture path");
        assert_eq!(
            fixture["status"],
            "ok",
            "{path}: {}",
            fixture["error"].as_str().unwrap_or("unknown error")
        );
        assert!(
            fixture["pages"].as_u64().unwrap_or(0) > 0,
            "{path}: no pages"
        );
        assert!(
            fixture["pdf_bytes"].as_u64().unwrap_or(0) > 0,
            "{path}: empty PDF"
        );
        assert!(
            fixture["fonts_embedded"].as_u64().unwrap_or(0) > 0,
            "{path}: no embedded font programs"
        );
        assert!(
            fixture["text_chars"].as_u64().unwrap_or(0) > 0,
            "{path}: no authored text"
        );
        let pdf = output.join(path).with_extension("pdf");
        let bytes = fs::read(&pdf).unwrap_or_else(|error| panic!("{}: {error}", pdf.display()));
        assert!(bytes.starts_with(b"%PDF-"), "{}: not a PDF", pdf.display());
    }

    println!("{}", serde_json::to_string_pretty(&manifest).unwrap());
    fs::remove_dir_all(&work).ok();
}

fn copy_fixture(workspace: &Path, relative: &Path, input: &Path) {
    let source = workspace.join(relative);
    let name = source
        .file_name()
        .unwrap_or_else(|| panic!("{} has no file name", source.display()));
    let target = input.join(name);
    fs::copy(&source, &target).unwrap_or_else(|error| panic!("copy {}: {error}", source.display()));
}
