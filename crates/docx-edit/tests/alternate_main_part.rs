use docx_parse::serializer::s13::{S13SaveRequest, write_docx_s13};
use serde_json::{Value, json};

#[path = "support/quality_fixture.rs"]
mod quality_fixture;

#[test]
fn alternate_main_part_loads_content_relationships_and_saves_edits_in_place() {
    let docx = quality_fixture::document(true);
    let wire = docx_parse::parse_docx_s9_wire(&docx, Default::default()).unwrap();
    assert_eq!(wire.document.package.document.content.len(), 12);
    assert_eq!(
        wire.document.package.header_entries.as_ref().unwrap().len(),
        1
    );
    let projection = docx_parse::parse_docx_s8_projection(&docx).unwrap();
    assert_eq!(projection.body.content.len(), 12);
    let mut body = serde_json::to_value(projection.body).unwrap();
    body["content"][0]["content"][0]["content"][0]["text"] = json!("Edited paragraph");
    for selective in [Value::Null, json!({"changedParaIds": ["00000001"]})] {
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": {
                "seed": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
                "now": "2030-01-02T03:04:05.006Z"
            },
            "document": body,
            "relationshipEntries": wire.document.package.relationship_entries,
            "headerEntries": wire.document.package.header_entries,
            "options": {"updateModifiedDate": false},
            "selective": selective
        }))
        .unwrap();
        let saved = write_docx_s13(request, &docx).unwrap();
        let parts = ooxml_opc::unzip_parts(&saved).unwrap();
        assert!(!parts.iter().any(|(path, _)| path == "word/document.xml"));
        assert!(
            !parts
                .iter()
                .any(|(path, _)| path == "word/_rels/document.xml.rels")
        );
        let xml = &parts
            .iter()
            .find(|(path, _)| path == "word/document2.xml")
            .unwrap()
            .1;
        assert!(String::from_utf8_lossy(xml).contains("Edited paragraph"));
        let reopened = docx_parse::parse_docx_s9_wire(&saved, Default::default()).unwrap();
        assert_eq!(reopened.document.package.document.content.len(), 12);
        assert_eq!(
            reopened
                .document
                .package
                .header_entries
                .as_ref()
                .unwrap()
                .len(),
            1
        );
    }
}

#[test]
fn absolute_strict_root_relationship_wins_over_an_unreferenced_default_part() {
    let mut parts = ooxml_opc::unzip_parts(&quality_fixture::document(true)).unwrap();
    let root = &mut parts
        .iter_mut()
        .find(|(path, _)| path == "_rels/.rels")
        .unwrap()
        .1;
    *root = String::from_utf8(root.clone())
        .unwrap()
        .replace(
            "Target=\"word/document2.xml\"",
            "Target=\"/word/document2.xml\"",
        )
        .replace(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
        )
        .into_bytes();
    parts.push((
        "word/document.xml".to_owned(),
        b"<w:document xmlns:w=\"w\"><w:body/></w:document>".to_vec(),
    ));
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let wire = docx_parse::parse_docx_s9_wire(&bytes, Default::default()).unwrap();
    assert_eq!(wire.document.package.document.content.len(), 12);
}

#[test]
fn relocated_headers_keep_their_paths_and_edits_on_full_and_selective_save() {
    for target in ["../stories/header1.xml", "/stories/header1.xml"] {
        let parts = ooxml_opc::unzip_parts(&quality_fixture::document(true)).unwrap();
        let parts: Vec<_> = parts
            .into_iter()
            .map(|(path, bytes)| {
                let path = path
                    .replace("word/header1.xml", "stories/header1.xml")
                    .replace("word/", "custom/");
                let xml = String::from_utf8(bytes)
                    .unwrap()
                    .replace("word/header1.xml", "stories/header1.xml")
                    .replace("word/", "custom/")
                    .replace("Target=\"header1.xml\"", &format!("Target=\"{target}\""));
                (path, xml.into_bytes())
            })
            .collect();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let wire = docx_parse::parse_docx_s9_wire(&bytes, Default::default()).unwrap();
        let projection = docx_parse::parse_docx_s8_projection(&bytes).unwrap();
        assert_eq!(projection.header_entries.len(), 1);
        let mut headers = serde_json::to_value(&wire.document.package.header_entries).unwrap();
        headers[0][1]["content"] = json!([{
            "type": "paragraph", "content": [{
                "type": "run", "content": [{"type": "text", "text": "Edited header"}]
            }]
        }]);
        for selective in [Value::Null, json!({"changedParaIds": []})] {
            let request: S13SaveRequest = serde_json::from_value(json!({
                "determinism": {
                    "seed": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
                    "now": "2030-01-02T03:04:05.006Z"
                },
                "document": projection.body,
                "relationshipEntries": wire.document.package.relationship_entries,
                "headerEntries": headers,
                "options": {"updateModifiedDate": false},
                "selective": selective
            }))
            .unwrap();
            let saved = write_docx_s13(request, &bytes).unwrap();
            let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
            assert!(
                !saved_parts
                    .iter()
                    .any(|(path, _)| path.starts_with("word/"))
            );
            let reopened = docx_parse::parse_docx_s9_wire(&saved, Default::default()).unwrap();
            let header = serde_json::to_value(reopened.document.package.header_entries).unwrap();
            assert!(header.to_string().contains("Edited header"));
            assert_eq!(reopened.document.package.document.content.len(), 12);
        }
    }
}
