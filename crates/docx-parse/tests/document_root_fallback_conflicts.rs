use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use serde_json::json;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn save(root: &str, para_id: Option<&str>) -> String {
    let original = ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            format!("{root}<w:body><w:p><w:r><w:t>old</w:t></w:r></w:p></w:body></w:document>").into_bytes(),
        ),
    ])
    .unwrap();
    let mut paragraph = json!({
        "type": "paragraph",
        "content": [{"type": "run", "content": [{"type": "text", "text": "edited"}]}],
    });
    if let Some(id) = para_id {
        paragraph["paraId"] = json!(id);
    }
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": [paragraph]},
        "options": {"updateModifiedDate": false},
    }))
    .unwrap();
    let saved = write_docx_s13(request, &original).expect("fallback save");
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    String::from_utf8(xml).unwrap()
}

fn root(xml: &str) -> &str {
    let start = xml.find("<w:document").unwrap();
    &xml[start..start + xml[start..].find('>').unwrap() + 1]
}

#[test]
fn strict_source_root_fallback_keeps_the_written_body_namespace() {
    let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    let xml = save(&format!("<w:document xmlns:w=\"{strict}\">"), None);
    assert!(root(&xml).contains(&format!("xmlns:w=\"{W}\"")));
    assert!(!root(&xml).contains(strict));
    assert!(xml.contains("<w:t>edited</w:t>"));
}

#[test]
fn conflicting_mc_source_root_fallback_keeps_paragraph_ids_consistent() {
    let xml = save(
        &format!("<w:document xmlns:w=\"{W}\" xmlns:mc=\"urn:custom-compatibility\">"),
        Some("1234ABCD"),
    );
    assert!(root(&xml).contains(&format!("xmlns:w=\"{W}\"")));
    assert!(root(&xml).contains(&format!("xmlns:w14=\"{W14}\"")));
    assert!(root(&xml).contains(&format!("xmlns:mc=\"{MC}\"")));
    assert!(!root(&xml).contains("urn:custom-compatibility"));
    assert!(xml.contains("w14:paraId=\"1234ABCD\""));
    assert!(xml.contains("<w:t>edited</w:t>"));
}
