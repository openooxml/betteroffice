use std::fs;
use std::path::{Path, PathBuf};

use betteroffice_docx::Document;

const CONTEXT: usize = 60;

fn document_xml(package: &[u8]) -> Vec<u8> {
    ooxml_opc::unzip_parts(package)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "word/document.xml")
        .map(|(_, bytes)| bytes)
        .expect("the saved package has no word/document.xml")
}

fn window(bytes: &[u8], at: usize) -> String {
    let start = at.saturating_sub(CONTEXT);
    let end = (at + CONTEXT).min(bytes.len());
    String::from_utf8_lossy(&bytes[start..end]).into_owned()
}

fn mismatch(name: &str, crate_xml: &[u8], product_xml: &[u8]) -> Option<String> {
    if crate_xml == product_xml {
        return None;
    }
    let at = crate_xml
        .iter()
        .zip(product_xml)
        .position(|(left, right)| left != right)
        .unwrap_or_else(|| crate_xml.len().min(product_xml.len()));
    Some(format!(
        "{name}: document.xml differs at byte {at} (crate {} bytes, product {} bytes)\n  crate:   {}\n  product: {}",
        crate_xml.len(),
        product_xml.len(),
        window(crate_xml, at),
        window(product_xml, at),
    ))
}

fn cases(dir: &Path) -> Vec<PathBuf> {
    let mut cases: Vec<PathBuf> = fs::read_dir(dir)
        .unwrap_or_else(|error| panic!("cannot read {}: {error}", dir.display()))
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.join("input.docx").is_file())
        .collect();
    cases.sort();
    cases
}

#[test]
#[ignore = "reads the product's saves from DOCX_CRATE_PARITY_DIR; CI runs it with --ignored"]
fn crate_save_matches_the_product_save() {
    let dir = PathBuf::from(
        std::env::var_os("DOCX_CRATE_PARITY_DIR").expect("DOCX_CRATE_PARITY_DIR is not set"),
    );
    let cases = cases(&dir);
    assert!(!cases.is_empty(), "no cases in {}", dir.display());

    let mut failures = Vec::new();
    for case in &cases {
        let name = case.file_name().unwrap().to_string_lossy().into_owned();
        let input = fs::read(case.join("input.docx")).unwrap();
        let spec: serde_json::Value =
            serde_json::from_slice(&fs::read(case.join("edit.json")).unwrap()).unwrap();

        let unedited = Document::open(&input).unwrap().save().unwrap();
        failures.extend(mismatch(
            &format!("{name} unedited"),
            &document_xml(&unedited),
            &fs::read(case.join("unedited.xml")).unwrap(),
        ));

        let mut document = Document::open(&input).unwrap();
        let index = spec["paragraph"].as_u64().unwrap() as usize;
        let para_id = document.paragraphs()[index]
            .para_id
            .clone()
            .unwrap_or_else(|| panic!("{name}: paragraph {index} has no ID"));
        assert_eq!(
            para_id,
            spec["paraId"].as_str().unwrap(),
            "{name}: paragraph {index}"
        );
        document
            .replace_paragraph_text(&para_id, spec["text"].as_str().unwrap())
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        let edited = document.save().unwrap();
        failures.extend(mismatch(
            &format!("{name} edited"),
            &document_xml(&edited),
            &fs::read(case.join("edited.xml")).unwrap(),
        ));
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// The crate's report in the shape of the product's `report.json`, omitted fields absent.
fn report_json(report: &betteroffice_docx::SaveReport) -> serde_json::Value {
    let mut value = serde_json::Map::new();
    if let Some(kept) = &report.kept_unchanged {
        let paragraphs: Vec<serde_json::Value> = kept
            .paragraphs
            .iter()
            .map(|paragraph| {
                let mut entry = serde_json::json!({
                    "part": paragraph.part,
                    "ordinal": paragraph.ordinal,
                    "text": paragraph.text,
                    "scope": paragraph.scope,
                    "dropped": paragraph.dropped,
                });
                if let Some(para_id) = &paragraph.para_id {
                    entry["paraId"] = serde_json::json!(para_id);
                }
                entry
            })
            .collect();
        value.insert(
            "keptUnchanged".to_owned(),
            serde_json::json!({ "structure": kept.structure, "paragraphs": paragraphs }),
        );
    }
    if let Some(dropped) = &report.structure_dropped {
        let parts: Vec<serde_json::Value> = dropped
            .iter()
            .map(|structure| {
                let paragraphs: Vec<serde_json::Value> = structure
                    .paragraphs
                    .iter()
                    .map(|paragraph| {
                        serde_json::json!({
                            "ordinal": paragraph.ordinal,
                            "paraId": paragraph.para_id,
                            "text": paragraph.text,
                            "dropped": paragraph.dropped,
                        })
                    })
                    .collect();
                serde_json::json!({
                    "part": structure.part,
                    "dropped": structure.dropped,
                    "paragraphs": paragraphs,
                })
            })
            .collect();
        value.insert("structureDropped".to_owned(), serde_json::Value::Array(parts));
    }
    serde_json::Value::Object(value)
}

/// The product's report with each kept paragraph's session key read as its Word paragraph ID.
fn with_word_ids(mut product: serde_json::Value, ids: &serde_json::Value) -> serde_json::Value {
    let paragraphs = product
        .pointer_mut("/keptUnchanged/paragraphs")
        .and_then(serde_json::Value::as_array_mut);
    for paragraph in paragraphs.into_iter().flatten() {
        let word = paragraph
            .get("paraId")
            .and_then(serde_json::Value::as_str)
            .and_then(|key| ids.get(key))
            .cloned();
        if let Some(word) = word {
            paragraph["paraId"] = word;
        }
    }
    product
}

#[test]
#[ignore = "reads the product's save reports from DOCX_CRATE_PARITY_DIR; CI runs it with --ignored"]
fn crate_save_report_matches_the_product_save() {
    let dir = PathBuf::from(
        std::env::var_os("DOCX_CRATE_PARITY_DIR").expect("DOCX_CRATE_PARITY_DIR is not set"),
    )
    .join("report");
    let cases = cases(&dir);
    assert!(!cases.is_empty(), "no cases in {}", dir.display());

    let mut failures = Vec::new();
    for case in &cases {
        let name = case.file_name().unwrap().to_string_lossy().into_owned();
        let input = fs::read(case.join("input.docx")).unwrap();
        let spec: serde_json::Value =
            serde_json::from_slice(&fs::read(case.join("edit.json")).unwrap()).unwrap();
        let read = |file: &str| -> serde_json::Value {
            serde_json::from_slice(&fs::read(case.join(file)).unwrap()).unwrap()
        };

        let mut document = Document::open(&input).unwrap();
        if let Some(para_id) = spec.get("paraId").and_then(serde_json::Value::as_str) {
            document
                .replace_paragraph_text(para_id, spec["text"].as_str().unwrap())
                .unwrap_or_else(|error| panic!("{name}: {error}"));
        }
        let report = document
            .save_report()
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        if report.bytes != document.save().unwrap() {
            failures.push(format!("{name}: save_report bytes differ from save"));
        }
        let crate_report = report_json(&report);
        let product = with_word_ids(read("report.json"), &read("ids.json"));
        if name == "structure-dropped" && product.get("structureDropped").is_none() {
            failures.push(format!("{name}: the product reports no dropped structure"));
        }
        if crate_report != product {
            failures.push(format!(
                "{name}: save report differs\n  crate:   {crate_report}\n  product: {product}"
            ));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
