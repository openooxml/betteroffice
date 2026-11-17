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
        let text = spec["text"].as_str().unwrap();
        if spec.get("paraId").is_none() {
            if let Err(error) = document.replace_paragraph_text_at(index, text) {
                failures.push(format!("{name} edited: {error}"));
                continue;
            }
            failures.extend(mismatch(
                &format!("{name} edited"),
                &document_xml(&document.save().unwrap()),
                &fs::read(case.join("edited.xml")).unwrap(),
            ));
            continue;
        }
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
            .replace_paragraph_text(&para_id, text)
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
