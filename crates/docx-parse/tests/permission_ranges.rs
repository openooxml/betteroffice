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

/// Saves the body `source` with its paragraphs as the body `edited` reads, in order, but for
/// those `removed`; `changed` among them.
fn save(source: &str, edited: &str, changed: &[u32], removed: &[u32]) -> Result<String, String> {
    let ordinals: Vec<Option<u32>> = (0..)
        .filter(|ordinal| !removed.contains(ordinal))
        .take(edited.matches("<w:p>").count())
        .map(Some)
        .collect();
    save_paragraphs(source, edited, &ordinals, changed, removed)
}

/// Saves the body `source` with the paragraphs of the body `edited` written from the source
/// paragraphs `ordinals` gives them, or inserted after the one before them when `None`.
fn save_paragraphs(
    source: &str,
    edited: &str,
    ordinals: &[Option<u32>],
    changed: &[u32],
    removed: &[u32],
) -> Result<String, String> {
    let source = document(source);
    let original = package(&source);
    let parsed = parse_docx_s9_wire(&package(&document(edited)), S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    let mut content = serde_json::to_value(&parsed.document.content).unwrap();
    let paragraphs = content.as_array_mut().unwrap();
    assert_eq!(paragraphs.len(), ordinals.len());
    let mut inserted = Vec::new();
    let mut previous = None;
    for (paragraph, ordinal) in paragraphs.iter_mut().zip(ordinals) {
        match ordinal {
            Some(ordinal) => {
                paragraph["sourceOrdinal"] = json!(ordinal);
                previous = Some(*ordinal);
            }
            None => inserted.push(json!({"after": previous.unwrap()})),
        }
    }
    let written: Vec<u32> = ordinals.iter().flatten().copied().collect();
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": content},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/document.xml",
            "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
            "paragraphs": written,
            "changed": changed,
            "inserted": inserted,
            "removed": removed,
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
        &[],
    )
    .unwrap();
    let plain = save(
        &format!("<w:p>{}{}</w:p>", run("Draft"), run("Signed")),
        &edited,
        &[0],
        &[],
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
        &[],
    )
    .unwrap();
    assert!(saved.contains(&format!("{}{START}{}</w:p>", run("Uno"), run("Two"))));
    assert!(saved.contains(&run("Centre")));
    assert!(saved.contains(&format!("{}{END}{}</w:p>", run("Three"), run("Four"))));
    assert_eq!(saved.matches(START).count(), 1);
    assert_eq!(saved.matches(END).count(), 1);
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
    assert_eq!(save(&body, &body, &[], &[]).unwrap(), document(&body));
}

#[test]
fn replacing_text_across_a_permission_marker_keeps_the_range_around_the_new_text() {
    let saved = save(
        &format!("<w:p>{}{START}{}{END}</w:p>", run("Before"), run("After")),
        &format!("<w:p>{}</w:p>", run("Something else")),
        &[0],
        &[],
    )
    .unwrap();
    assert!(saved.contains(&format!("<w:p>{START}{}{END}</w:p>", run("Something else"))));
}

#[test]
fn text_typed_into_an_empty_permission_range_lands_inside_it() {
    let saved = save(
        &format!("<w:p>{}{START}{END}{}</w:p>", run("A"), run("B")),
        &format!("<w:p>{}{}{}</w:p>", run("A"), run("new"), run("B")),
        &[0],
        &[],
    )
    .unwrap();
    assert!(saved.contains(&format!(
        "{}{START}{}{END}{}",
        run("A"),
        run("new"),
        run("B")
    )));
}

#[test]
fn deleting_the_text_of_a_range_between_bookmarks_keeps_its_start_before_its_end() {
    let start = r#"<w:bookmarkStart w:id="0" w:name="b"/>"#;
    let end = r#"<w:bookmarkEnd w:id="0"/>"#;
    let saved = save(
        &format!(
            "<w:p>{}{start}{START}{}{END}{end}</w:p>",
            run("Keep"),
            run("A")
        ),
        &format!("<w:p>{}{start}{end}</w:p>", run("Keep")),
        &[0],
        &[],
    )
    .unwrap();
    assert_eq!(saved.matches(START).count(), 1);
    assert_eq!(saved.matches(END).count(), 1);
    assert!(saved.find(START) < saved.find(END));
}

#[test]
fn a_permission_range_in_a_hyperlink_moves_beside_it_when_the_paragraph_is_edited() {
    let saved = save(
        &format!(
            r#"<w:p><w:hyperlink w:anchor="t">{START}{}{END}</w:hyperlink>{}</w:p>"#,
            run("Link"),
            run("Tail")
        ),
        &format!(
            r#"<w:p><w:hyperlink w:anchor="t">{}</w:hyperlink>{}</w:p>"#,
            run("Link"),
            run("End")
        ),
        &[0],
        &[],
    )
    .unwrap();
    assert_eq!(saved.matches(START).count(), 1);
    assert_eq!(saved.matches(END).count(), 1);
    assert!(saved.find(START) < saved.find("<w:hyperlink"));
    assert!(saved.find("</w:hyperlink>") < saved.find(END));
    assert!(saved.find(END) < saved.find("End"));
}

#[test]
fn deleting_a_paragraph_moves_its_permission_marker_to_the_paragraph_before() {
    let saved = save(
        &format!(
            "<w:p>{}</w:p><w:p>{START}{}</w:p><w:p>{}{END}</w:p>",
            run("Keep"),
            run("Gone"),
            run("Rest")
        ),
        &format!("<w:p>{}</w:p><w:p>{}</w:p>", run("Keep"), run("Rest")),
        &[],
        &[1],
    )
    .unwrap();
    assert!(!saved.contains("Gone"));
    assert!(saved.contains(&format!("{}{START}</w:p>", run("Keep"))));
    assert!(saved.contains(&format!("{}{END}</w:p>", run("Rest"))));
    assert_eq!(saved.matches(START).count(), 1);
    assert_eq!(saved.matches(END).count(), 1);
}

#[test]
fn merging_a_paragraph_keeps_its_permission_marker_where_its_text_went() {
    let saved = save(
        &format!(
            "<w:p>{}</w:p><w:p>{}{START}{}</w:p><w:p>{}{END}</w:p>",
            run("One"),
            run("Two"),
            run("Three"),
            run("Four")
        ),
        &format!(
            "<w:p>{}{}{}</w:p><w:p>{}</w:p>",
            run("One"),
            run("Two"),
            run("Three"),
            run("Four")
        ),
        &[0],
        &[1],
    )
    .unwrap();
    assert!(saved.contains(&format!("{}{START}{}</w:p>", run("Two"), run("Three"))));
    assert!(saved.contains(&format!("{}{END}</w:p>", run("Four"))));
    assert_eq!(saved.matches(START).count(), 1);
}

#[test]
fn a_save_without_edits_keeps_permission_markers_in_a_text_box_byte_for_byte() {
    let body = format!(
        "<w:p>{}<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:textbox><w:txbxContent><w:p><w:permStart w:id='4' w:edGrp='everyone'></w:permStart>{}<w:permEnd w:id='4'></w:permEnd></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p><w:p>{}</w:p>",
        run("Before"),
        run("Boxed"),
        run("After")
    );
    let saved = save_paragraphs(&body, &body, &[Some(0), Some(2)], &[], &[]).unwrap();
    assert_eq!(saved, document(&body));
}

#[test]
fn separate_edits_on_both_sides_of_a_range_keep_it_around_its_unchanged_text() {
    let saved = save(
        &format!(
            "<w:p>{}{START}{}{END}{}</w:p>",
            run("A"),
            run("M"),
            run("Z")
        ),
        &format!("<w:p>{}{}{}</w:p>", run("B"), run("M"), run("Y")),
        &[0],
        &[],
    )
    .unwrap();
    assert!(saved.contains(&format!("{}{START}{}{END}{}", run("B"), run("M"), run("Y"))));
}

#[test]
fn splitting_a_paragraph_keeps_each_end_of_its_range_with_its_text() {
    let saved = save_paragraphs(
        &format!("<w:p>{START}{}{END}</w:p>", run("AB")),
        &format!("<w:p>{}</w:p><w:p>{}</w:p>", run("A"), run("B")),
        &[Some(0), None],
        &[0],
        &[],
    )
    .unwrap();
    assert!(saved.contains(&format!("<w:p>{START}{}</w:p>", run("A"))));
    assert!(saved.contains(&format!("{}{END}</w:p>", run("B"))));
}

#[test]
fn merging_three_paragraphs_keeps_a_range_on_its_own_text() {
    let saved = save(
        &format!(
            "<w:p>{}</w:p><w:p>{}</w:p><w:p>{START}{}{END}</w:p>",
            run("A"),
            run("B"),
            run("C")
        ),
        &format!("<w:p>{}{}{}</w:p>", run("A"), run("B"), run("C")),
        &[0],
        &[1, 2],
    )
    .unwrap();
    assert!(saved.contains(&format!("{}{}{START}{}{END}", run("A"), run("B"), run("C"))));
}

#[test]
fn a_paragraph_kept_as_raw_markup_does_not_take_the_markers_of_an_identical_one() {
    let block = format!(
        r#"<x:block xmlns:x="urn:example:block"><w:p>{}</w:p></x:block>"#,
        run("One")
    );
    let saved = save_paragraphs(
        &format!(
            "{block}<w:p>{START}{}</w:p><w:p>{}{END}</w:p>",
            run("One"),
            run("Two")
        ),
        &format!("{block}<w:p>{}</w:p><w:p>{}</w:p>", run("One"), run("Deux")),
        &[Some(0), Some(1), Some(2)],
        &[2],
        &[],
    )
    .unwrap();
    assert_eq!(saved.matches(START).count(), 1);
    assert!(saved.find("</x:block>") < saved.find(START));
    assert!(saved.contains(&format!("{START}{}</w:p>", run("One"))));
    assert!(saved.contains(&format!("{}{END}</w:p>", run("Deux"))));
}

#[test]
fn a_source_paragraph_comment_cannot_hide_permission_ranges() {
    for block in [
        "",
        r#"<x:block xmlns:x="urn:example:block"><![CDATA[<!--bo-paragraph _0-->]]></x:block>"#,
    ] {
        let source = format!(
            "<!--bo-paragraph 0-->{}",
            document(&format!("{block}<w:p>{START}{}{END}</w:p>", run("Draft")))
        );
        let original = package(&source);
        let parsed = parse_docx_s9_wire(&original, S9ParseOptions::default())
            .unwrap()
            .document
            .package;
        let mut content = serde_json::to_value(&parsed.document.content).unwrap();
        let paragraph = usize::from(!block.is_empty());
        content[paragraph]["sourceOrdinal"] = json!(0);
        content[paragraph]["content"][0]["content"][0]["text"] = json!("Final");
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
            "document": {"content": content},
            "options": {"updateModifiedDate": false},
            "paragraphIds": {"splicedParts": [{
                "part": "word/document.xml",
                "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
                "paragraphs": [0],
                "changed": [0],
            }]},
        }))
        .unwrap();
        let saved = write_docx_s13(request, &original).unwrap();
        let (_, xml) = ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap();
        assert_eq!(
            String::from_utf8(xml).unwrap(),
            source.replace("Draft", "Final")
        );
    }
}
