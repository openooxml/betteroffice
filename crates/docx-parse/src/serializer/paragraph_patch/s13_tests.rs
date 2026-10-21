use super::*;
use crate::paragraph_patch::{Attrs, UnitV1, result_sha256, unit_sha256};

fn save_patch(source_paragraph: &str, seed_text: &str) -> Result<String, ParseError> {
    save_patch_mode(source_paragraph, seed_text, true)
}

fn save_patch_mode(
    source_paragraph: &str,
    seed_text: &str,
    minimal: bool,
) -> Result<String, ParseError> {
    let root = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#;
    let keep = "<w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>";
    let source = format!("{root}{source_paragraph}\n{keep}</w:body></w:document>");
    let units = |value: &str| {
        vec![UnitV1::Text {
            text: value.to_owned(),
            attrs: Attrs::new(),
        }]
    };
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
    }))
    .unwrap();
    let mut request = request;
    if !minimal {
        request.paragraph_ids.as_mut().unwrap().spliced_parts[0]
            .paragraph_patches.clear();
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

fn mixed_request(middle: &str) -> (String, S13SaveRequest) {
    let root = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#;
    let patched = "<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>old</w:t></w:r></w:p>";
    let keep = "<w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>";
    let source = format!("{root}{keep}{patched}{middle}{patched}{keep}</w:body></w:document>");
    let units = |value: &str| {
        vec![UnitV1::Text {
            text: value.to_owned(),
            attrs: Attrs::new(),
        }]
    };
    let seed = units("old");
    let plans: Vec<_> = [1, 3]
        .into_iter()
        .map(|ordinal| {
            json!({
                "ordinal": ordinal, "sessionKey": format!("source:{ordinal}"),
                "seed": {"units": seed, "properties": {}, "unitSha256": unit_sha256(&seed)},
                "textEdits": [{"from": 0, "to": 3, "insert": [{"text": "new", "attrs": {}, "inherit": {}}], "donor": {"kind": "at", "at": 0}}],
                "runEdits": [], "paraDeltas": [], "paraInherit": {},
                "resultSha256": result_sha256(&units("new"), &Attrs::new())
            })
        })
        .collect();
    let content: Vec<_> = ["keep", "new", "replaced", "new", "keep"]
        .into_iter()
        .enumerate()
        .map(|(ordinal, value)| {
            let mut paragraph = text_paragraph(value, None);
            paragraph["sourceOrdinal"] = json!(ordinal);
            paragraph
        })
        .collect();
    let request = serde_json::from_value(json!({
        "determinism": determinism(), "document": {"content": content},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/document.xml",
            "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
            "paragraphs": [0, 1, 2, 3, 4], "changed": [1, 2, 3],
            "paragraphPatches": plans,
            "legacyParagraphs": [{"ordinal": 2, "code": "UNSUPPORTED_STRUCTURE", "detail": "field: legacy field"}]
        }]}
    }))
    .unwrap();
    (source, request)
}

#[test]
fn paragraph_patch_s13_replaces_only_the_balanced_field_paragraph() {
    let field = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>field</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
    let (source, request) = mixed_request(field);
    let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
    let output = String::from_utf8(part_map(&saved)["word/document.xml"].clone()).unwrap();
    assert_eq!(
        output,
        source.replace("old", "new").replace(
            field,
            "<w:p><w:r><w:t>replaced</w:t></w:r></w:p>"
        )
    );
    assert!(warnings.contains(&"DOCX_MINIMAL_PATCH replace counts {\"field\":1}".to_owned()));
}

#[test]
fn paragraph_patch_s13_refuses_unadmitted_replacement_with_applied_patches() {
    let field = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>field</w:t></w:r></w:p>"#;
    let (source, request) = mixed_request(field);
    let error = write_docx_s13(request, &base_package(&source))
        .unwrap_err()
        .to_string();
    assert!(
        error.contains("DOCX_MINIMAL_PATCH [PATCH_CONFLICT] word/document.xml p#2:"),
        "{error}"
    );
}

#[test]
fn paragraph_patch_s13_rejects_selective_patch_plans() {
    let (source, mut request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    request.selective = Some(S13SelectiveSave {
        changed_para_ids: vec!["AAAAAAAA".to_owned()],
        source_paragraphs: None,
    });
    let error = write_docx_s13(request, &base_package(&source))
        .unwrap_err()
        .to_string();
    assert!(error.contains("DOCX_MINIMAL_PATCH [INVALID_PLAN]"), "{error}");
}

#[test]
fn paragraph_patch_s13_aggregates_producer_and_writer_replace_counts() {
    let field = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>field</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
    let (source, mut request) = mixed_request(field);
    let source = source.replacen(
        "<w:t>old</w:t>",
        r#"<w:fldChar w:fldCharType="begin"/><w:t>old</w:t><w:fldChar w:fldCharType="end"/>"#,
        1,
    );
    let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
    part.sha256 = format!("{:x}", Sha256::digest(source.as_bytes()));
    let (_, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
    assert!(warnings.contains(&"DOCX_MINIMAL_PATCH replace counts {\"field\":2}".to_owned()));
}
