//! Permission ranges, which the model does not keep, through saves spliced into their source.

use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::json;
use sha2::{Digest, Sha256};

const START: &str = r#"<w:permStart w:id="7" w:edGrp="everyone"/>"#;
const END: &str = r#"<w:permEnd w:id="7"/>"#;

fn document(body: &str) -> String {
    format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    )
}

fn package(document: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.as_bytes().to_vec()),
    ])
    .unwrap()
}

fn run(text: &str) -> String {
    format!("<w:r><w:t>{text}</w:t></w:r>")
}

/// Saves the body `source` with its paragraphs as the body `edited` reads, `changed` among them.
fn save(source: &str, edited: &str, changed: &[u32]) -> Result<String, String> {
    let source = document(source);
    let original = package(&source);
    let parsed = parse_docx_s9_wire(&package(&document(edited)), S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    let mut content = serde_json::to_value(&parsed.document.content).unwrap();
    let paragraphs = content.as_array_mut().unwrap();
    for (ordinal, paragraph) in paragraphs.iter_mut().enumerate() {
        paragraph["sourceOrdinal"] = json!(ordinal);
    }
    let ordinals: Vec<usize> = (0..paragraphs.len()).collect();
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": content},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/document.xml",
            "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
            "paragraphs": ordinals,
            "changed": changed,
            "inserted": [],
            "removed": [],
        }]},
    }))
    .unwrap();
    let saved = write_docx_s13(request, &original).map_err(|error| format!("{error:?}"))?;
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    Ok(String::from_utf8(xml).unwrap())
}

#[test]
fn an_edit_beside_a_permission_range_keeps_its_markers_in_place() {
    let edited = format!("<w:p>{}{}</w:p>", run("Final"), run("Signed"));
    let saved = save(
        &format!("<w:p>{}{START}{}{END}</w:p>", run("Draft"), run("Signed")),
        &edited,
        &[0],
    )
    .unwrap();
    let plain = save(
        &format!("<w:p>{}{}</w:p>", run("Draft"), run("Signed")),
        &edited,
        &[0],
    )
    .unwrap();
    assert!(saved.contains(&format!(
        "{}{START}{}{END}</w:p>",
        run("Final"),
        run("Signed")
    )));
    assert_eq!(saved.replace(START, "").replace(END, ""), plain);
}

#[test]
fn edits_at_both_ends_of_a_permission_range_across_paragraphs_keep_both_markers() {
    let saved = save(
        &format!(
            "<w:p>{}{START}{}</w:p><w:p>{}</w:p><w:p>{}{END}{}</w:p>",
            run("One"),
            run("Two"),
            run("Middle"),
            run("Three"),
            run("Four")
        ),
        &format!(
            "<w:p>{}{}</w:p><w:p>{}</w:p><w:p>{}{}</w:p>",
            run("Uno"),
            run("Two"),
            run("Centre"),
            run("Three"),
            run("Four")
        ),
        &[0, 1],
    )
    .unwrap();
    assert!(saved.contains(&format!("{}{START}{}</w:p>", run("Uno"), run("Two"))));
    assert!(saved.contains(&run("Centre")));
    assert!(saved.contains(&format!("{}{END}{}</w:p>", run("Three"), run("Four"))));
    assert_eq!(saved.matches(START).count(), 1);
    assert_eq!(saved.matches(END).count(), 1);
}

#[test]
fn refuses_an_edit_across_a_permission_marker_instead_of_dropping_it() {
    let error = save(
        &format!("<w:p>{}{START}{}{END}</w:p>", run("Before"), run("After")),
        &format!("<w:p>{}</w:p>", run("Something else")),
        &[0],
    )
    .unwrap_err();
    assert!(error.contains("word/document.xml"), "{error}");
}

#[test]
fn a_save_without_edits_keeps_permission_ranges_byte_for_byte() {
    let body = format!(
        "<w:p>{}{START}{}</w:p><w:p>{}{END}</w:p><w:p><w:permStart w:id=\"8\" w:ed=\"someone\"/>{}<w:permEnd w:id=\"8\"/></w:p>",
        run("One"),
        run("Two"),
        run("Three"),
        run("Four")
    );
    assert_eq!(save(&body, &body, &[]).unwrap(), document(&body));
}
