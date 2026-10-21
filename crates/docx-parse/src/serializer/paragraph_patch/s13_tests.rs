use super::*;
use crate::paragraph_patch::{Attrs, UnitV1, result_sha256, unit_sha256};

fn save_patch(source_paragraph: &str, seed_text: &str) -> Result<String, ParseError> {
    save_patch_mode(source_paragraph, seed_text, true)
}

fn save_patch_mode(source_paragraph: &str, seed_text: &str, minimal: bool) -> Result<String, ParseError> {
    let root = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#;
    let keep = "<w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>";
    let source = format!("{root}{source_paragraph}\n{keep}</w:body></w:document>");
    let units = |value: &str| vec![UnitV1::Text {
        text: value.to_owned(),
        attrs: Attrs::new(),
    }];
    let seed = units(seed_text);
    let mut edited = text_paragraph("new", None);
    edited["sourceOrdinal"] = json!(0);
    let mut kept = text_paragraph("keep", None);
    kept["sourceOrdinal"] = json!(1);
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": determinism(),
        "document": {"content": [edited, text_paragraph("inserted", None), kept]},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/document.xml",
            "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
            "paragraphs": [0, 1], "changed": [0], "inserted": [{"after": 0}],
            "paragraphPatches": [{
                "ordinal": 0, "sessionKey": "source:0",
                "seed": {"units": seed, "properties": {}, "unitSha256": unit_sha256(&seed)},
                "textEdits": [{"from": 0, "to": 3, "insert": [{"text": "new", "attrs": {}, "inherit": {}}], "donor": {"kind": "at", "at": 0}}],
                "runEdits": [], "paraDeltas": [], "paraInherit": {},
                "resultSha256": result_sha256(&units("new"), &Attrs::new())
            }]
        }]}
    })).unwrap();
    let mut request = request;
    if !minimal {
        request.paragraph_ids.as_mut().unwrap().spliced_parts[0].paragraph_patches.clear();
    }
    let saved = write_docx_s13(request, &base_package(&source))?;
    Ok(String::from_utf8(part_map(&saved)["word/document.xml"].clone()).unwrap())
}

#[test]
fn paragraph_patch_s13_keeps_source_and_allows_insertions() {
    let source = "<w:p w:rsidR='87654321'><w:r><w:rPr><w:noProof/></w:rPr><w:t>old</w:t></w:r></w:p>";
    let output = save_patch(source, "old").unwrap();
    let root = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#;
    assert_eq!(output, format!(
        "{root}{}<w:p><w:r><w:t>inserted</w:t></w:r></w:p>\n<w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p></w:body></w:document>",
        source.replace("old", "new")
    ));
}

#[test]
fn paragraph_patch_s13_propagates_seed_mismatch() {
    let source = "<w:p><w:r><w:t>old</w:t></w:r></w:p>";
    assert!(save_patch(source, "bad").unwrap_err().to_string().contains(
        "DOCX_MINIMAL_PATCH [SEED_MISMATCH] word/document.xml p#0: source units differ from seed"
    ));
}

#[test]
fn paragraph_patch_s13_unsupported_structure_uses_legacy_replacement() {
    let source = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/><w:t>old</w:t><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
    let output = save_patch(source, "old").unwrap();
    let expected = save_patch_mode(source, "old", false).unwrap();
    assert_eq!(output, expected);
}
