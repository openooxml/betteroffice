use docx_parse::document::{DocumentBody, Section};
use docx_parse::s9::{S9ParseOptions, parse_docx_s9_wire};
use docx_parse::serializer::{
    S13SaveOptions, S13SaveRequest, SerializerDeterminism, write_docx_s13,
};
use quick_xml::{Reader, events::Event};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

fn save_request(original: &[u8]) -> S13SaveRequest {
    let package = parse_docx_s9_wire(original, S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    let body = package.document;
    let sections = body.sections.map(|sections| {
        sections
            .into_iter()
            .map(|section| Section {
                id: section.id,
                properties: section.properties,
                content: body.content[section.content_start..section.content_end].to_vec(),
            })
            .collect()
    });
    S13SaveRequest {
        determinism: SerializerDeterminism {
            seed: format!("{:x}", Sha256::digest(original)),
            now: "1970-01-01T00:00:00.000Z".to_owned(),
        },
        document: DocumentBody {
            content: body.content,
            sections,
            final_section_properties: body.final_section_properties,
            custom_root_bindings: body.custom_root_bindings,
            comments: body.comments,
        },
        header_entries: package.header_entries.unwrap_or_default(),
        footer_entries: package.footer_entries.unwrap_or_default(),
        footnotes: package.footnotes.unwrap_or_default(),
        endnotes: package.endnotes.unwrap_or_default(),
        footnote_separators: package.footnote_separators.unwrap_or_default(),
        endnote_separators: package.endnote_separators.unwrap_or_default(),
        relationship_entries: package.relationship_entries,
        numbering: Some(package.numbering),
        options: S13SaveOptions {
            update_modified_date: false,
            modified_by: None,
        },
        selective: None,
    }
}

fn comment_package(content: Value) -> Vec<u8> {
    let original = ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#.to_vec(),
        ),
    ])
    .unwrap();
    let mut request = save_request(&original);
    request.document = serde_json::from_value(json!({
        "content": [{ "type": "paragraph", "paraId": "11111111", "content": content }],
        "comments": [{
            "id": 7, "author": "Reviewer", "initials": "R",
            "date": "2026-01-01T00:00:00Z", "status": "active", "paletteIndex": 0,
            "content": [{ "type": "paragraph", "content": [
                { "type": "run", "content": [{ "type": "text", "text": "New comment" }] }
            ] }]
        }]
    }))
    .unwrap();
    write_docx_s13(request, &original).unwrap()
}

fn comment_range() -> Value {
    json!([
        { "type": "commentRangeStart", "id": 7 },
        { "type": "run", "content": [{ "type": "text", "text": "Hello" }] },
        { "type": "commentRangeEnd", "id": 7 }
    ])
}

fn element_ids(saved: &[u8], part: &str, name: &[u8]) -> Vec<String> {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let xml = &parts.iter().find(|(path, _)| path == part).unwrap().1;
    let mut reader = Reader::from_reader(xml.as_slice());
    let mut ids = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element) if element.name().as_ref() == name => {
                ids.push(
                    element
                        .try_get_attribute("w:id")
                        .unwrap()
                        .unwrap()
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .unwrap()
                        .into_owned(),
                );
            }
            Event::Eof => break,
            _ => {}
        }
    }
    ids
}

fn assert_comment_anchor(saved: &[u8]) {
    for name in ["commentReference", "commentRangeStart", "commentRangeEnd"] {
        assert_eq!(
            element_ids(saved, "word/document.xml", format!("w:{name}").as_bytes()),
            ["7"],
            "{name} must occur exactly once with the comment id"
        );
    }
    assert_eq!(element_ids(saved, "word/comments.xml", b"w:comment"), ["7"]);
}

#[test]
fn a_new_comment_saves_one_reference_and_both_range_markers() {
    let saved = comment_package(comment_range());
    assert_comment_anchor(&saved);
    let resaved = write_docx_s13(save_request(&saved), &saved).unwrap();
    assert_comment_anchor(&resaved);
    assert_eq!(saved, resaved);
}

#[test]
fn an_existing_comment_reference_is_preserved_without_a_duplicate() {
    let reference = json!({
        "type": "run", "formatting": { "styleId": "CommentReference", "bold": true },
        "content": [{ "type": "commentReference", "id": 7 }]
    });
    for wrapper in [
        reference.clone(),
        json!({ "type": "hyperlink", "anchor": "target", "children": [reference.clone()] }),
        json!({
            "type": "inlineSdt", "properties": { "sdtType": "richText", "alias": "Comment" },
            "content": [reference.clone()]
        }),
        json!({
            "type": "insertion", "info": { "id": 1, "author": "Reviewer" },
            "content": [reference]
        }),
    ] {
        for index in [2, 3] {
            let mut content = comment_range();
            content
                .as_array_mut()
                .unwrap()
                .insert(index, wrapper.clone());
            let saved = comment_package(content);
            assert_comment_anchor(&saved);
            let resaved = write_docx_s13(save_request(&saved), &saved).unwrap();
            assert_comment_anchor(&resaved);
            assert_eq!(saved, resaved);
        }
    }
}

#[test]
fn another_comments_reference_does_not_suppress_a_new_reference() {
    let mut content = comment_range();
    content.as_array_mut().unwrap().push(json!({
        "type": "run", "content": [{ "type": "commentReference", "id": 9 }]
    }));
    let saved = comment_package(content);
    assert_eq!(
        element_ids(&saved, "word/document.xml", b"w:commentReference"),
        ["7", "9"]
    );
}
