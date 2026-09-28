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
        paragraph_ids: None,
    }
}

const DOCUMENT_XML: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#;

fn package(document_xml: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document_xml.as_bytes().to_vec()),
    ])
    .unwrap()
}

/// The source package and a request saving `blocks` as its body, with comment 7.
fn comment_request(blocks: Value) -> (Vec<u8>, S13SaveRequest) {
    let original = package(DOCUMENT_XML);
    let mut request = save_request(&original);
    request.document = serde_json::from_value(json!({
        "content": blocks,
        "comments": [{
            "id": 7, "author": "Reviewer", "initials": "R",
            "date": "2026-01-01T00:00:00Z", "status": "active", "paletteIndex": 0,
            "content": [{ "type": "paragraph", "content": [
                { "type": "run", "content": [{ "type": "text", "text": "New comment" }] }
            ] }]
        }]
    }))
    .unwrap();
    (original, request)
}

fn comment_package(content: Value) -> Vec<u8> {
    let (original, request) =
        comment_request(json!([{ "type": "paragraph", "paraId": "11111111", "content": content }]));
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

fn text_run(text: &str) -> Value {
    json!({ "type": "run", "content": [{ "type": "text", "text": text }] })
}

fn reference(id: u32) -> Value {
    json!({ "type": "run", "content": [{ "type": "commentReference", "id": id }] })
}

fn paragraph(para_id: &str, content: Value) -> Value {
    json!({ "type": "paragraph", "paraId": para_id, "content": content })
}

fn document_xml(saved: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let xml = &parts
        .iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap()
        .1;
    String::from_utf8(xml.clone()).unwrap()
}

/// Asserts one anchor for comment 7, its reference right after the range end.
fn assert_reference_after_end(saved: &[u8]) {
    assert_comment_anchor(saved);
    assert_document_reference_after_end(saved);
}

fn assert_document_reference_after_end(saved: &[u8]) {
    for name in ["commentReference", "commentRangeStart", "commentRangeEnd"] {
        let ids = element_ids(saved, "word/document.xml", format!("w:{name}").as_bytes());
        assert_eq!(ids, ["7"], "{name}");
    }
    let xml = document_xml(saved);
    let end = xml.find(r#"<w:commentRangeEnd w:id="7"/>"#).unwrap();
    let reference = xml.find(r#"<w:commentReference w:id="7"/>"#).unwrap();
    assert!(reference > end && !xml[end..reference].contains("</w:p>"));
}

#[test]
fn a_reference_outside_its_range_end_paragraph_moves_there() {
    let stale = json!([text_run("World"), reference(7)]);
    for blocks in [
        json!([
            paragraph("11111111", comment_range()),
            paragraph("22222222", stale.clone())
        ]),
        json!([
            paragraph("11111111", comment_range()),
            { "type": "table", "rows": [{ "type": "tableRow", "cells": [{
                "type": "tableCell", "content": [paragraph("22222222", stale)]
            }] }] }
        ]),
    ] {
        let (original, request) = comment_request(blocks);
        let saved = write_docx_s13(request, &original).unwrap();
        assert_reference_after_end(&saved);
        assert_eq!(write_docx_s13(save_request(&saved), &saved).unwrap(), saved);
    }
}

#[test]
fn only_the_first_reference_of_a_range_end_paragraph_stays() {
    let mut content = comment_range();
    let first = json!({
        "type": "run", "formatting": { "bold": true },
        "content": [{ "type": "commentReference", "id": 7 }]
    });
    content.as_array_mut().unwrap().insert(1, first);
    content.as_array_mut().unwrap().push(reference(7));
    let saved = comment_package(content);
    assert_comment_anchor(&saved);
    let xml = document_xml(&saved);
    let reference = xml.find(r#"<w:commentReference w:id="7"/>"#).unwrap();
    assert!(reference < xml.find(r#"<w:commentRangeEnd w:id="7"/>"#).unwrap());
    assert!(xml[..reference].contains("<w:b/>"));
}

#[test]
fn references_of_a_comment_without_a_range_end_stay() {
    let (original, request) = comment_request(json!([
        paragraph("11111111", json!([text_run("One"), reference(9)])),
        paragraph("22222222", json!([reference(9), reference(9)])),
    ]));
    let saved = write_docx_s13(request, &original).unwrap();
    assert_eq!(
        element_ids(&saved, "word/document.xml", b"w:commentReference"),
        ["9", "9", "9"]
    );
}

#[test]
fn a_note_kept_verbatim_keeps_its_references_and_a_serialized_note_loses_strays() {
    let (original, mut request) = comment_request(json!([paragraph("11111111", comment_range())]));
    let verbatim =
        r#"<w:footnote w:id="1"><w:p><w:r><w:commentReference w:id="7"/></w:r></w:p></w:footnote>"#;
    request.footnotes = serde_json::from_value(json!([
        {
            "type": "footnote", "id": 1, "noteType": "normal",
            "content": [paragraph("33333333", json!([reference(7)]))],
            "verbatimXml": verbatim
        },
        {
            "type": "footnote", "id": 2, "noteType": "normal",
            "content": [paragraph("44444444", json!([text_run("Note"), reference(7)]))]
        }
    ]))
    .unwrap();
    let saved = write_docx_s13(request, &original).unwrap();
    assert_document_reference_after_end(&saved);
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let notes = String::from_utf8(
        parts
            .iter()
            .find(|(path, _)| path == "word/footnotes.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert!(notes.contains(verbatim));
    assert_eq!(notes.matches("<w:commentReference ").count(), 1);
}

#[test]
fn a_part_kept_as_source_bytes_keeps_them_with_a_stray_reference_and_a_hyperlink_field() {
    let source = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="7"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="7"/></w:p><w:p><w:r><w:commentReference w:id="7"/></w:r><w:r><w:t>World</w:t></w:r></w:p><w:p><w:hyperlink w:anchor="target"><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple></w:hyperlink></w:p></w:body></w:document>"#;
    let original = package(source);
    let (_, mut request) = comment_request(json!([
        { "type": "paragraph", "sourceOrdinal": 0, "content": comment_range() },
        { "type": "paragraph", "sourceOrdinal": 1, "content": [reference(7), text_run("World")] },
        { "type": "paragraph", "sourceOrdinal": 2, "content": [
            { "type": "hyperlink", "anchor": "target", "children": [text_run("Page ")] }
        ] },
    ]));
    request.paragraph_ids = serde_json::from_value(json!({
        "patchedParts": [{ "part": "word/document.xml", "paraIds": [] }]
    }))
    .unwrap();
    let saved = write_docx_s13(request, &original).unwrap();
    assert_eq!(document_xml(&saved), source);
}

#[test]
fn a_part_whose_paragraph_id_patch_fails_is_normalized_as_it_is_serialized() {
    let original = package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="7"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="7"/></w:p><w:p><w:r><w:commentReference w:id="7"/></w:r><w:r><w:t>World</w:t></w:r></w:p></w:body></w:document>"#,
    );
    let (_, mut request) = comment_request(json!([
        paragraph("11111111", comment_range()),
        paragraph("22222222", json!([reference(7), text_run("World")])),
    ]));
    request.footnotes = serde_json::from_value(json!([{
        "type": "footnote", "id": 1, "noteType": "normal",
        "content": [paragraph("33333333", json!([text_run("Note"), reference(7)]))]
    }]))
    .unwrap();
    request.paragraph_ids = serde_json::from_value(json!({ "patchedParts": [
        { "part": "word/document.xml", "paraIds": [[9, "1A2B3C4D"]] },
        { "part": "word/footnotes.xml", "paraIds": [[0, "2B3C4D5E"]] }
    ] }))
    .unwrap();
    let saved = write_docx_s13(request, &original).unwrap();
    assert_reference_after_end(&saved);
    assert!(element_ids(&saved, "word/footnotes.xml", b"w:commentReference").is_empty());
}

fn selective_original() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:commentRangeStart w:id="7"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="7"/></w:p><w:p w14:paraId="22222222"><w:r><w:t>World</w:t></w:r><w:r><w:commentReference w:id="7"/></w:r></w:p></w:body></w:document>"#,
    )
}

/// A request saving the selective source with `range_id` on its range paragraph.
fn selective_request(range_id: Value) -> S13SaveRequest {
    let (_, request) = comment_request(json!([
        { "type": "paragraph", "paraId": range_id, "content": comment_range() },
        paragraph("22222222", json!([text_run("World!"), reference(7)])),
    ]));
    request
}

#[test]
fn a_selective_save_of_the_stray_reference_paragraph_keeps_one_reference() {
    let original = selective_original();
    let mut request = selective_request(json!("11111111"));
    request.selective = serde_json::from_value(json!({ "changedParaIds": ["22222222"] })).unwrap();
    let saved = write_docx_s13(request, &original).unwrap();
    let xml = document_xml(&saved);
    assert_eq!(
        element_ids(&saved, "word/document.xml", b"w:commentReference"),
        ["7"]
    );
    assert!(xml.contains("<w:t>World!</w:t></w:r><w:r><w:commentReference w:id=\"7\"/>"));

    let original = package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:commentRangeStart w:id="7"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>World</w:t></w:r><w:r><w:commentReference w:id="7"/></w:r></w:p></w:body></w:document>"#,
    );
    let (_, mut request) = comment_request(json!([
        paragraph(
            "11111111",
            json!([
                { "type": "commentRangeStart", "id": 7 }, text_run("Hello"),
                { "type": "commentRangeEnd", "id": 7 }, reference(7)
            ])
        ),
        paragraph("22222222", json!([text_run("World!"), reference(7)])),
    ]));
    request.selective = serde_json::from_value(json!({ "changedParaIds": ["22222222"] })).unwrap();
    let saved = write_docx_s13(request, &original).unwrap();
    assert_document_reference_after_end(&saved);
    assert!(document_xml(&saved).contains("World!"));
}

#[test]
fn a_source_paragraph_save_keeps_one_reference_whichever_paragraphs_it_replaces() {
    let original = selective_original();
    let digest = format!("{:x}", Sha256::digest(document_xml(&original).as_bytes()));
    let replaced = |blocks: Value| {
        let mut request = selective_request(json!("11111111"));
        request.selective = serde_json::from_value(json!({
            "sourceParagraphs": { "partSha256": digest, "paragraphs": blocks }
        }))
        .unwrap();
        write_docx_s13(request, &original).unwrap()
    };
    let saved = replaced(json!([{ "path": [0, 1], "block": 1 }]));
    assert_eq!(
        element_ids(&saved, "word/document.xml", b"w:commentReference"),
        ["7"]
    );
    assert!(document_xml(&saved).contains("World!"));
    let saved = replaced(json!([
        { "path": [0, 0], "block": 0 },
        { "path": [0, 1], "block": 1 }
    ]));
    assert_document_reference_after_end(&saved);
}

#[test]
fn a_source_paragraph_save_leaves_references_outside_its_paragraphs_as_the_source_has_them() {
    let original = package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="7"/><w:r><w:t>Hello</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p><w:p><w:r><w:t>World</w:t></w:r><w:r><w:commentReference w:id="7"/></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>"#,
    );
    let (_, mut request) = comment_request(json!([
        { "type": "paragraph", "content": [
            { "type": "commentRangeStart", "id": 7 }, text_run("Hello"),
            { "type": "commentRangeEnd", "id": 7 }, reference(7)
        ] },
        { "type": "paragraph", "content": [text_run("World"), reference(7)] },
        { "type": "paragraph", "content": [text_run("Tail edited")] },
    ]));
    let digest = format!("{:x}", Sha256::digest(document_xml(&original).as_bytes()));
    request.selective = serde_json::from_value(json!({
        "sourceParagraphs": { "partSha256": digest, "paragraphs": [{ "path": [0, 2], "block": 2 }] }
    }))
    .unwrap();
    let saved = write_docx_s13(request, &original).unwrap();
    let xml = document_xml(&saved);
    assert_eq!(
        element_ids(&saved, "word/document.xml", b"w:commentReference"),
        ["7", "7"]
    );
    assert!(xml.contains("<w:t>World</w:t></w:r><w:r><w:commentReference w:id=\"7\"/></w:r>"));
    assert!(xml.contains("Tail edited"));
}
