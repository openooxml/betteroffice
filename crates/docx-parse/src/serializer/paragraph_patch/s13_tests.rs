use super::*;
use crate::paragraph_patch::{Attrs, UnitV1, result_sha256, unit_sha256};
use serde_json::Value;

fn save_patch(source_paragraph: &str, seed_text: &str) -> Result<String, ParseError> {
    save_patch_with_warnings(source_paragraph, seed_text).map(|(output, _)| output)
}

fn save_patch_with_warnings(
    source_paragraph: &str,
    seed_text: &str,
) -> Result<(String, Vec<String>), ParseError> {
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
    let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source))?;
    let output = String::from_utf8(part_map(&saved)["word/document.xml"].clone()).unwrap();
    Ok((output, warnings))
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
fn paragraph_patch_s13_preserves_body_run_start_attributes() {
    let source = concat!(
        "<w:p w:rsidR='87654321'><w:r  w:rsidR='00112233'",
        " w:rsidRPr=\"00575701\" w:rsidDel = '00947A26'",
        " xmlns:x='urn:run-metadata' x:id='foreign' x:keep='A&#x26;B' w:unknown='opaque'>",
        "<w:rPr><w:noProof/></w:rPr><w:t>old</w:t></w:r></w:p>"
    );
    let (output, warnings) = save_patch_with_warnings(source, "old").unwrap();
    let expected = format!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
         <w:body>{}<w:p><w:r><w:t>inserted</w:t></w:r></w:p>\n\
         <w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p></w:body></w:document>",
        source.replace("old", "new")
    );
    assert_eq!(output.as_bytes(), expected.as_bytes());
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn paragraph_patch_s13_preserves_footnote_run_start_attributes() {
    let source = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n",
        "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        "<w:footnote w:type='separator' w:id='-1'>",
        "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>\n",
        "<w:footnote w:type='continuationSeparator' w:id='0'>",
        "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>\n",
        "<w:footnote w:id='3'><w:p>",
        "<w:r w:rsidRPr=\"00575701\"><w:rPr><w:rStyle w:val=\"FootnoteReference\"/>",
        "</w:rPr><w:footnoteRef/></w:r>",
        "<w:r w:rsidRPr=\"00947A26\"><w:t xml:space=\"preserve\"> Some note text</w:t></w:r>",
        "</w:p></w:footnote></w:footnotes>"
    );
    let body = concat!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        "<w:body><w:p><w:r><w:t>Body</w:t></w:r>",
        "<w:r><w:footnoteReference w:id=\"3\"/></w:r></w:p></w:body></w:document>"
    );
    let separators: Vec<_> = [(0, -1, "separator"), (1, 0, "continuationSeparator")]
        .into_iter()
        .map(|(ordinal, id, kind)| {
            json!({
                "type": "footnote", "id": id, "noteType": kind,
                "content": [{"type": "paragraph", "sourceOrdinal": ordinal,
                    "content": [{"type": "run", "content": [{"type": kind}]}]}]
            })
        })
        .collect();
    let units = |value: &str| {
        vec![UnitV1::Text {
            text: value.to_owned(),
            attrs: Attrs::new(),
        }]
    };
    let seed = units(" Some note text");
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": determinism(),
        "document": {"content": [{"type": "paragraph", "content": [
            {"type": "run", "content": [{"type": "text", "text": "Body"}]},
            {"type": "run", "content": [{"type": "footnoteRef", "id": 3}]}
        ]}]},
        "footnoteSeparators": separators,
        "footnotes": [{"type": "footnote", "id": 3, "noteType": "normal", "content": [{
            "type": "paragraph", "sourceOrdinal": 2, "content": [
                {"type": "run", "formatting": {"styleId": "FootnoteReference"},
                    "content": [{"type": "footnoteRefMark"}]},
                {"type": "run", "content": [{"type": "text", "text": " Some edited text",
                    "preserveSpace": true}]}
            ]
        }]}],
        "relationshipEntries": [["rIdNotes", {
            "id": "rIdNotes", "type": relationship_types::FOOTNOTES, "target": "footnotes.xml"
        }]],
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [{
            "part": "word/footnotes.xml",
            "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
            "paragraphs": [0, 1, 2], "changed": [2],
            "paragraphPatches": [{
                "ordinal": 2, "sessionKey": "source:2",
                "seed": {"units": seed, "properties": {}, "unitSha256": unit_sha256(&seed)},
                "textEdits": [{"from": 6, "to": 10, "insert": [{
                    "text": "edited", "attrs": {}, "inherit": {}
                }], "donor": {"kind": "at", "at": 6}}],
                "runEdits": [], "paraDeltas": [], "paraInherit": {},
                "resultSha256": result_sha256(&units(" Some edited text"), &Attrs::new())
            }]
        }]}
    }))
    .unwrap();
    let mut original_parts = ooxml_opc::unzip_parts(&base_package(body)).unwrap();
    original_parts.push(("word/footnotes.xml".to_owned(), source.as_bytes().to_vec()));
    let original = ooxml_opc::rezip_parts(&original_parts).unwrap();
    let (saved, warnings) = write_docx_s13_with_warnings(request, &original).unwrap();
    let parts = part_map(&saved);
    let output = &parts["word/footnotes.xml"];
    let expected = source.replace(" Some note text", " Some edited text");
    assert_eq!(output.as_slice(), expected.as_bytes());
    let output = std::str::from_utf8(output).unwrap();
    assert!(output.contains("<w:r w:rsidRPr=\"00575701\">"));
    assert!(output.contains("<w:r w:rsidRPr=\"00947A26\">"));
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn paragraph_patch_s13_seed_mismatch_uses_counted_paragraph_replacement() {
    let source = "<w:p><w:r><w:t>old</w:t></w:r></w:p>";
    let (output, warnings) = save_patch_with_warnings(source, "bad").unwrap();
    let (expected, _) = save_patch_with_warnings(source, "old").unwrap();
    assert_eq!(output.as_bytes(), expected.as_bytes());
    assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"ambiguous\":1}"]);
}

#[test]
fn paragraph_patch_s13_unsupported_structure_uses_legacy_replacement() {
    let source = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/><w:t>old</w:t><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
    let (output, warnings) = save_patch_with_warnings(source, "old").unwrap();
    let root = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#;
    let keep = "<w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>";
    assert_eq!(
        output,
        format!(
            "{root}<w:p><w:r><w:t>new</w:t></w:r></w:p><w:p><w:r><w:t>inserted</w:t></w:r></w:p>\n{keep}</w:body></w:document>"
        )
    );
    assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"field\":1}"]);
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

fn assert_writer_conflict_replaces_only_paragraph(
    middle: &str,
    patch: Value,
    mut paragraph: Value,
    replacement: &str,
) {
    let (source, mut request) = mixed_request(middle);
    let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
    part.legacy_paragraphs.clear();
    part.paragraph_patches.push(serde_json::from_value(patch).unwrap());
    paragraph["sourceOrdinal"] = json!(2);
    request.document.content[2] = serde_json::from_value(paragraph).unwrap();
    assert_ambiguous_paragraph_replacement(&source, request, middle, replacement);
}

fn assert_ambiguous_paragraph_replacement(
    source: &str,
    request: S13SaveRequest,
    middle: &str,
    replacement: &str,
) {
    let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(source)).unwrap();
    let parts = part_map(&saved);
    let expected = source.replace("old", "new").replace(middle, replacement);
    assert_eq!(parts["word/document.xml"].as_slice(), expected.as_bytes());
    assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"ambiguous\":1}"]);
}

#[test]
fn paragraph_patch_s13_whitespace_comment_id_replaces_only_changed_paragraph() {
    let middle = concat!(
        "<w:p w:rsidR='87654321'><w:r><w:t>abc</w:t></w:r>",
        "<w:r><w:commentReference w:id=\" 1 \"/></w:r></w:p>"
    );
    let (source, mut request) = mixed_request(middle);
    let atom = UnitV1::Atom {
        atom_kind: "field".to_owned(),
        payload: serde_json::from_value(json!({
            "fieldType": "COMMENT", "instruction": "", "displayText": "",
            "fieldKind": "simple", "fldLock": false, "dirty": false,
            "displayMode": "result", "hasCachedResult": false,
            "modelKind": "commentReference", "commentId": 1
        }))
        .unwrap(),
        attrs: Attrs::new(),
    };
    let units = |text: &str| {
        vec![
            UnitV1::Text {
                text: text.to_owned(),
                attrs: Attrs::new(),
            },
            atom.clone(),
        ]
    };
    let seed = units("abc");
    let patch = json!({
        "ordinal": 2, "sessionKey": "source:2",
        "seed": {"units": seed, "properties": {}, "unitSha256": unit_sha256(&seed)},
        "textEdits": [{"from": 1, "to": 2, "donor": {"kind": "at", "at": 1},
            "insert": [{"text": "X", "attrs": {}, "inherit": {}}]}],
        "runEdits": [], "paraDeltas": [], "paraInherit": {},
        "resultSha256": result_sha256(&units("aXc"), &Attrs::new())
    });
    let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
    part.legacy_paragraphs.clear();
    part.paragraph_patches.push(serde_json::from_value(patch).unwrap());
    let mut paragraph = text_paragraph("aXc", None);
    paragraph["sourceOrdinal"] = json!(2);
    paragraph["content"].as_array_mut().unwrap().push(json!({
        "type": "run", "content": [{"type": "commentReference", "id": 1}]
    }));
    request.document.content[2] = serde_json::from_value(paragraph).unwrap();
    request.document.comments = Some(vec![
        serde_json::from_value(json!({"id": 1, "author": "A", "content": []})).unwrap(),
    ]);
    assert_ambiguous_paragraph_replacement(
        &source,
        request,
        middle,
        concat!(
            "<w:p><w:r><w:t>aXc</w:t></w:r>",
            "<w:r><w:commentReference w:id=\"1\"/></w:r></w:p>"
        ),
    );
}

#[test]
fn paragraph_patch_s13_hyperlink_bookmarks_replace_only_changed_paragraph() {
    let middle = concat!(
        "<w:p w:rsidR='87654321'><w:bookmarkStart w:id=\"1\" w:name=\"outer\"/>",
        "<w:r><w:t>abc</w:t></w:r><w:bookmarkEnd w:id=\"1\"/>",
        "<w:hyperlink w:anchor=\"outer\"><w:bookmarkStart w:id=\"2\" w:name=\"inner\"/>",
        "<w:r><w:t>link</w:t></w:r><w:bookmarkEnd w:id=\"2\"/></w:hyperlink></w:p>"
    );
    let units = |text: &str| {
        vec![
            UnitV1::Text {
                text: text.to_owned(),
                attrs: Attrs::new(),
            },
            UnitV1::Text {
                text: "link".to_owned(),
                attrs: serde_json::from_value(json!({
                    "hyperlink": {"href": "#outer", "tooltip": null, "anchor": "outer"}
                }))
                .unwrap(),
            },
        ]
    };
    let properties: Attrs = serde_json::from_value(json!({"bookmarks": [
        {"id": 1, "name": "outer", "kind": "start", "offset": 0},
        {"id": 1, "kind": "end", "offset": 3}
    ]}))
    .unwrap();
    let seed = units("abc");
    let patch = json!({
        "ordinal": 2, "sessionKey": "source:2",
        "seed": {"units": seed, "properties": properties, "unitSha256": unit_sha256(&seed)},
        "textEdits": [{"from": 1, "to": 2, "donor": {"kind": "at", "at": 1},
            "insert": [{"text": "X", "attrs": {}, "inherit": {}}]}],
        "runEdits": [], "paraDeltas": [], "paraInherit": {},
        "resultSha256": result_sha256(&units("aXc"), &properties)
    });
    let paragraph = json!({"type": "paragraph", "content": [
        {"type": "bookmarkStart", "id": 1, "name": "outer", "position": {"offset": 0}},
        {"type": "run", "content": [{"type": "text", "text": "aXc"}]},
        {"type": "bookmarkEnd", "id": 1, "position": {"offset": 3}},
        {"type": "hyperlink", "anchor": "outer", "children": [
            {"type": "bookmarkStart", "id": 2, "name": "inner", "position": {"offset": 0}},
            {"type": "run", "content": [{"type": "text", "text": "link"}]},
            {"type": "bookmarkEnd", "id": 2, "position": {"offset": 4}}
        ]}
    ]});
    let replacement = middle
        .replace(" w:rsidR='87654321'", "")
        .replace(">abc<", ">aXc<");
    assert_writer_conflict_replaces_only_paragraph(middle, patch, paragraph, &replacement);
}

#[test]
fn paragraph_patch_s13_property_order_conflict_replaces_only_changed_paragraph() {
    let middle = concat!(
        "<w:p w:rsidR='87654321'><w:r><w:rPr><w:sz w:val=\"24\"/><w:b/></w:rPr>",
        "<w:t>abc</w:t></w:r></w:p>"
    );
    let seed = vec![UnitV1::Text {
        text: "abc".to_owned(),
        attrs: serde_json::from_value(json!({"fontSize": {"size": 24}, "bold": true})).unwrap(),
    }];
    let target = json!({"fontSize": {"size": 24}, "bold": true, "italic": true});
    let result = vec![UnitV1::Text {
        text: "abc".to_owned(),
        attrs: serde_json::from_value(target.clone()).unwrap(),
    }];
    let patch = json!({
        "ordinal": 2, "sessionKey": "source:2",
        "seed": {"units": seed, "properties": {}, "unitSha256": unit_sha256(&seed)},
        "textEdits": [],
        "runEdits": [{
            "from": 0, "to": 3, "target": target, "inherit": {},
            "deltas": [{"path": ["italic"], "before": {"state": "missing"},
                "after": {"state": "present", "value": true}}]
        }],
        "paraDeltas": [], "paraInherit": {},
        "resultSha256": result_sha256(&result, &Attrs::new())
    });
    let mut paragraph = text_paragraph("abc", None);
    paragraph["content"][0]["formatting"] = json!({"fontSize": 24, "bold": true, "italic": true});
    assert_writer_conflict_replaces_only_paragraph(
        middle,
        patch,
        paragraph,
        "<w:p><w:r><w:rPr><w:b/><w:i/><w:sz w:val=\"24\"/></w:rPr><w:t>abc</w:t></w:r></w:p>",
    );
}

#[test]
fn paragraph_patch_s13_insert_then_format_conflict_replaces_only_changed_paragraph() {
    let middle = "<w:p w:rsidR='87654321'><w:r><w:t>abc</w:t></w:r></w:p>";
    let seed = vec![UnitV1::Text {
        text: "abc".to_owned(),
        attrs: Attrs::new(),
    }];
    let result = vec![UnitV1::Text {
        text: "aXbc".to_owned(),
        attrs: serde_json::from_value(json!({"bold": true})).unwrap(),
    }];
    let patch = json!({
        "ordinal": 2, "sessionKey": "source:2",
        "seed": {"units": seed, "properties": {}, "unitSha256": unit_sha256(&seed)},
        "textEdits": [{"from": 1, "to": 1, "donor": {"kind": "at", "at": 1},
            "insert": [{"text": "X", "attrs": {"bold": true}, "inherit": {}}]}],
        "runEdits": [{
            "from": 0, "to": 3, "target": {"bold": true}, "inherit": {},
            "deltas": [{"path": ["bold"], "before": {"state": "missing"},
                "after": {"state": "present", "value": true}}]
        }],
        "paraDeltas": [], "paraInherit": {},
        "resultSha256": result_sha256(&result, &Attrs::new())
    });
    let mut paragraph = text_paragraph("aXbc", None);
    paragraph["content"][0]["formatting"] = json!({"bold": true});
    assert_writer_conflict_replaces_only_paragraph(
        middle,
        patch,
        paragraph,
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>aXbc</w:t></w:r></w:p>",
    );
}

#[test]
fn paragraph_patch_s13_no_edits_keeps_overlapping_parts_unchanged() {
    let source = concat!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        "<w:body>\n <w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>\n",
        "</w:body></w:document>"
    );
    let mut kept = text_paragraph("keep", None);
    kept["sourceOrdinal"] = json!(0);
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": determinism(), "document": {"content": [kept]},
        "options": {"updateModifiedDate": false},
        "paragraphIds": {
            "patchedParts": [{"part": "word/document.xml", "paraIds": []}],
            "splicedParts": [{
                "part": "word/document.xml",
                "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
                "paragraphs": [0], "changed": []
            }]
        }
    }))
    .unwrap();
    let original = base_package(source);
    let (saved, warnings) = write_docx_s13_with_warnings(request, &original).unwrap();
    assert_eq!(part_map(&saved), part_map(&original));
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn paragraph_patch_s13_body_edit_keeps_overlapping_header_and_footer_unchanged() {
    let (source, mut request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    let mut original_parts = ooxml_opc::unzip_parts(&base_package(&source)).unwrap();
    for (kind, root, rel_type) in [
        ("header", "hdr", relationship_types::HEADER),
        ("footer", "ftr", relationship_types::FOOTER),
    ] {
        let story = format!(
            "<w:{root} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n \
             <w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>\n</w:{root}>"
        );
        let path = format!("word/{kind}1.xml");
        let id = format!("rId_{kind}");
        let mut paragraph = text_paragraph("keep", None);
        paragraph["sourceOrdinal"] = json!(0);
        let entry = serde_json::from_value(json!({
            "type": kind, "hdrFtrType": "default", "content": [paragraph]
        }))
        .unwrap();
        if kind == "header" {
            request.header_entries.push((id.clone(), entry));
        } else {
            request.footer_entries.push((id.clone(), entry));
        }
        request.relationship_entries.push((
            id.clone(),
            serde_json::from_value(json!({
                "id": id, "type": rel_type, "target": format!("{kind}1.xml")
            }))
            .unwrap(),
        ));
        let ids = request.paragraph_ids.as_mut().unwrap();
        ids.patched_parts.push(crate::serializer::S13PatchedPart {
            part: path.clone(),
            para_ids: vec![],
        });
        ids.spliced_parts.push(
            serde_json::from_value(json!({
                "part": path,
                "sha256": format!("{:x}", Sha256::digest(story.as_bytes())),
                "paragraphs": [0], "changed": []
            }))
            .unwrap(),
        );
        original_parts.push((path, story.into_bytes()));
    }
    let original = ooxml_opc::rezip_parts(&original_parts).unwrap();
    let (saved, warnings) = write_docx_s13_with_warnings(request, &original).unwrap();
    let before = part_map(&original);
    let after = part_map(&saved);
    for path in ["word/header1.xml", "word/footer1.xml"] {
        assert_eq!(after[path], before[path], "{path}");
    }
    let body = std::str::from_utf8(&after["word/document.xml"]).unwrap();
    assert_eq!(
        body,
        source.replace("old", "new").replace(">field<", ">replaced<")
    );
    assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"field\":1}"]);
}

#[test]
fn paragraph_patch_s13_resolves_retained_attributes_at_ancestor_namespaces() {
    for element in ["w:document", "w:body"] {
        let (source, mut request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
        let source = source
            .replacen(
                &format!("<{element}"),
                &format!("<{element} xmlns:x='urn:metadata'"),
                1,
            )
            .replace("<w:r><w:rPr>", "<w:r x:keep='A&#x26;B'><w:rPr>");
        request.paragraph_ids.as_mut().unwrap().spliced_parts[0].sha256 =
            format!("{:x}", Sha256::digest(source.as_bytes()));
        let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
        let parts = part_map(&saved);
        let output = std::str::from_utf8(&parts["word/document.xml"]).unwrap();
        assert_eq!(
            output,
            source.replace("old", "new").replace(">field<", ">replaced<"),
            "{element}"
        );
        assert_eq!(output.matches("x:keep='A&#x26;B'").count(), 2);
        assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"field\":1}"]);
    }
}

#[test]
fn paragraph_patch_s13_admission_rejection_replaces_only_changed_paragraphs() {
    let (source, mut request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    let source = source.replace("<w:r><w:rPr>", "<w:r xmlns='urn:metadata'><w:rPr>");
    request.paragraph_ids.as_mut().unwrap().spliced_parts[0].sha256 =
        format!("{:x}", Sha256::digest(source.as_bytes()));
    let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
    let parts = part_map(&saved);
    let output = std::str::from_utf8(&parts["word/document.xml"]).unwrap();
    let expected = source.replace("old", "new").replace(">field<", ">replaced<");
    assert_eq!(output, expected);
    assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"field\":1}"]);
}

#[test]
fn paragraph_patch_s13_admission_rejection_keeps_patched_neighbors_exact() {
    let rejected = concat!(
        "<w:p><w:sdt><w:sdtPr><w:id w:val='7'/></w:sdtPr><w:sdtContent>",
        "<w:r><w:t>field</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    );
    for ancestor in [false, true] {
        let (source, mut request) = mixed_request(rejected);
        let source = if ancestor {
            source
                .replace("<w:body>", "<w:body xmlns:x='urn:metadata'>")
                .replace("<w:r><w:rPr>", "<w:r x:keep='A&#x26;B'><w:rPr>")
        } else {
            source.replace(
                "<w:r><w:rPr>",
                "<w:r xmlns:x='urn:metadata' x:keep='A&#x26;B'><w:rPr>",
            )
        };
        let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
        part.sha256 = format!("{:x}", Sha256::digest(source.as_bytes()));
        part.legacy_paragraphs[0].detail = "sdt: inline content control".to_owned();
        let (saved, warnings) =
            write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
        let parts = part_map(&saved);
        let expected = source.replace("old", "new").replace(
            rejected,
            "<w:p><w:r><w:t>replaced</w:t></w:r></w:p>",
        );
        assert_eq!(parts["word/document.xml"].as_slice(), expected.as_bytes());
        assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"sdt\":1}"]);
    }
}

#[test]
fn paragraph_patch_s13_admission_replacement_keeps_insertion_anchors() {
    let source = concat!(
        "<w:p><w:r xmlns='urn:metadata'><w:rPr><w:noProof/></w:rPr>",
        "<w:t>old</w:t></w:r></w:p>"
    );
    let (output, warnings) = save_patch_with_warnings(source, "old").unwrap();
    let root = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#;
    let expected = format!(
        "{root}<w:body>{}<w:p><w:r><w:t>inserted</w:t></w:r></w:p>\n\
         <w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p></w:body></w:document>",
        source.replace("old", "new")
    );
    assert_eq!(output, expected);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn paragraph_patch_s13_seed_content_mismatches_replace_only_changed_paragraph() {
    let (source, request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    let source = source.replacen("<w:noProof/>", "<w:b/>", 1);
    for restyle in [false, true] {
        for kind in ["text", "unitSha256", "bookmarks", "markers"] {
            let mut request = request.clone();
            let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
            part.sha256 = format!("{:x}", Sha256::digest(source.as_bytes()));
            let patch = &mut part.paragraph_patches[0];
            if restyle {
                let inherit = &mut patch.para_inherit;
                inherit.insert("alignment".to_owned(), json!("right"));
            }
            match kind {
                "text" => {
                    patch.seed.units = vec![UnitV1::Text {
                        text: "bad".to_owned(),
                        attrs: Attrs::new(),
                    }];
                    patch.seed.unit_sha256 = unit_sha256(&patch.seed.units);
                }
                "unitSha256" => {
                    patch.seed.unit_sha256 = "0".repeat(64);
                }
                "bookmarks" => {
                    patch.seed.properties.insert(
                        "bookmarks".to_owned(),
                        json!([{"id": 7, "name": "stale", "kind": "start", "offset": 0}]),
                    );
                }
                "markers" => {
                    let properties = &mut patch.seed.properties;
                    properties.insert("markers".to_owned(), json!([[0, "7", false]]));
                }
                _ => unreachable!(),
            }
            let (saved, warnings) =
                write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
            let parts = part_map(&saved);
            let expected = source
                .replacen(
                    "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>old</w:t></w:r></w:p>",
                    "<w:p><w:r><w:t>new</w:t></w:r></w:p>",
                    1,
                )
                .replace("old", "new")
                .replace(">field<", ">replaced<");
            assert_eq!(
                parts["word/document.xml"].as_slice(),
                expected.as_bytes(),
                "{kind}, restyle={restyle}"
            );
            assert_eq!(
                warnings,
                ["DOCX_MINIMAL_PATCH replace counts {\"ambiguous\":1,\"field\":1}"],
                "{kind}, restyle={restyle}"
            );
        }
    }
}

#[test]
fn paragraph_patch_s13_seed_mismatch_does_not_mask_source_digest_mismatch() {
    let (source, mut request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    let patch = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0].paragraph_patches[0];
    patch.seed.unit_sha256 = "0".repeat(64);
    let drifted = source.replace(">keep<", ">newer<");
    let error = write_docx_s13(request, &base_package(&drifted))
        .unwrap_err()
        .to_string();
    assert!(
        error.contains("DOCX_MINIMAL_PATCH [SOURCE_DIGEST_MISMATCH] word/document.xml p#?:"),
        "{error}"
    );
}

#[test]
fn paragraph_patch_s13_legacy_only_and_captured_plans_reject_source_drift() {
    let (source, request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    for legacy in [false, true] {
        let mut request = request.clone();
        let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
        part.paragraph_patches.clear();
        part.changed = vec![2];
        if legacy {
            part.legacy_paragraphs[0].detail = "restyle: style changed".to_owned();
        } else {
            part.legacy_paragraphs.clear();
        }
        for ordinal in [1, 3] {
            let mut paragraph = text_paragraph("old", None);
            paragraph["sourceOrdinal"] = json!(ordinal);
            request.document.content[ordinal] = serde_json::from_value(paragraph).unwrap();
        }
        let (saved, warnings) =
            write_docx_s13_with_warnings(request.clone(), &base_package(&source)).unwrap();
        let output = String::from_utf8(part_map(&saved)["word/document.xml"].clone()).unwrap();
        assert_eq!(output, source.replace(">field<", ">replaced<"));
        if legacy {
            assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"restyle\":1}"]);
        } else {
            assert!(warnings.is_empty(), "{warnings:?}");
        }
        let drifted = source.replace(">keep<", ">newer<");
        let error = write_docx_s13(request, &base_package(&drifted))
            .unwrap_err()
            .to_string();
        assert!(
            error.contains("DOCX_MINIMAL_PATCH [SOURCE_DIGEST_MISMATCH] word/document.xml p#?:"),
            "legacy={legacy}: {error}"
        );
    }
}

#[test]
fn paragraph_patch_s13_legacy_only_plans_cannot_be_filtered_out() {
    let (source, request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    for mode in ["patched", "selective", "unwritten"] {
        let mut request = request.clone();
        let ids = request.paragraph_ids.as_mut().unwrap();
        let part = &mut ids.spliced_parts[0];
        part.paragraph_patches.clear();
        part.changed = vec![2];
        match mode {
            "patched" => ids.patched_parts.push(crate::serializer::S13PatchedPart {
                part: "word/document.xml".to_owned(),
                para_ids: vec![(0, "00000001".to_owned())],
            }),
            "selective" => {
                request.selective = Some(S13SelectiveSave {
                    changed_para_ids: vec![],
                    source_paragraphs: None,
                });
            }
            "unwritten" => {
                part.part = "word/header1.xml".to_owned();
            }
            _ => unreachable!(),
        }
        let mut parts = ooxml_opc::unzip_parts(&base_package(&source)).unwrap();
        if mode == "unwritten" {
            parts.push(("word/header1.xml".to_owned(), source.as_bytes().to_vec()));
        }
        let error = write_docx_s13_parts(request, &parts, None)
            .unwrap_err()
            .to_string();
        assert!(
            error.contains("DOCX_MINIMAL_PATCH [INVALID_PLAN]"),
            "{mode}: {error}"
        );
    }
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
fn paragraph_patch_s13_unadmitted_structure_uses_counted_paragraph_replacement() {
    let field = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>field</w:t></w:r></w:p>"#;
    let (source, request) = mixed_request(field);
    let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
    let parts = part_map(&saved);
    let output = std::str::from_utf8(&parts["word/document.xml"]).unwrap();
    assert_eq!(
        output,
        source
            .replace("old", "new")
            .replace(field, "<w:p><w:r><w:t>replaced</w:t></w:r></w:p>")
    );
    assert_eq!(warnings, ["DOCX_MINIMAL_PATCH replace counts {\"field\":1}"]);
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

#[test]
fn paragraph_patch_s13_replace_only_preserves_untouched_paragraphs() {
    let field = r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>field</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
    let (source, mut request) = mixed_request(field);
    let part = &mut request.paragraph_ids.as_mut().unwrap().spliced_parts[0];
    part.paragraph_patches.clear();
    part.changed = vec![2];
    for ordinal in [1, 3] {
        let mut paragraph = text_paragraph("old", None);
        paragraph["sourceOrdinal"] = json!(ordinal);
        request.document.content[ordinal] = serde_json::from_value(paragraph).unwrap();
    }
    let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
    let output = String::from_utf8(part_map(&saved)["word/document.xml"].clone()).unwrap();
    assert_eq!(
        output,
        source.replace(field, "<w:p><w:r><w:t>replaced</w:t></w:r></w:p>")
    );
    assert!(warnings.contains(&"DOCX_MINIMAL_PATCH replace counts {\"field\":1}".to_owned()));
}

#[test]
fn paragraph_patch_s13_rejects_unwritten_story_parts() {
    let (source, request) = mixed_request("<w:p><w:r><w:t>field</w:t></w:r></w:p>");
    for (path, root) in [
        ("word/footnotes.xml", "footnotes"),
        ("word/endnotes.xml", "endnotes"),
        ("word/header1.xml", "hdr"),
        ("word/footer1.xml", "ftr"),
    ] {
        for seed_text in ["old", "bad"] {
            let story = format!(
                r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>old</w:t></w:r></w:p></w:{root}>"#
            );
            let mut request = request.clone();
            let ids = request.paragraph_ids.as_mut().unwrap();
            let part = &mut ids.spliced_parts[0];
            part.part = path.to_owned();
            part.sha256 = format!("{:x}", Sha256::digest(story.as_bytes()));
            part.paragraphs = vec![0];
            part.changed = vec![0];
            part.legacy_paragraphs.clear();
            part.paragraph_patches.truncate(1);
            let patch = &mut part.paragraph_patches[0];
            patch.ordinal = 0;
            patch.seed.units = vec![UnitV1::Text {
                text: seed_text.to_owned(),
                attrs: Attrs::new(),
            }];
            patch.seed.unit_sha256 = unit_sha256(&patch.seed.units);
            let mut parts = ooxml_opc::unzip_parts(&base_package(&source)).unwrap();
            parts.push((path.to_owned(), story.into_bytes()));
            let error = write_docx_s13_parts(request, &parts, None)
                .unwrap_err()
                .to_string();
            assert!(
                error.contains(&format!("DOCX_MINIMAL_PATCH [INVALID_PLAN] {path} p#?:")),
                "{error}"
            );
        }
    }
}

#[test]
fn paragraph_patch_s13_counts_changed_inheritance_as_restyle() {
    for (scope, before, after) in [
        ("run", json!({"boldCs": true}), json!({"boldCs": false})),
        ("run", json!({"italicCs": true}), json!({"italicCs": false})),
        (
            "mark",
            json!({"fontFamily": {"ascii": "Arial", "hAnsi": "Arial"}}),
            json!({"fontFamily": {"ascii": "Times", "hAnsi": "Times"}}),
        ),
        ("run", json!({}), json!({"highlight": {"color": "yellow"}})),
        ("paragraph", json!({}), json!({"bidi": true})),
        ("paragraph", json!({"widowControl": false}), json!({})),
        (
            "run",
            json!({"strike": {"double": true}}),
            json!({"strike": {"double": false}}),
        ),
        (
            "run",
            json!({"fontFamily": {"hAnsi": "Arial"}}),
            json!({"fontFamily": {"hAnsi": "Times"}}),
        ),
        ("run", json!({}), json!({"fontSize": {"sizeCs": 40}})),
        ("numPr", json!(null), json!({"numId": 2, "ilvl": 0})),
        ("pStyle", json!(null), json!("B")),
        ("runStyle", json!(null), json!({"styleId": "C"})),
        ("insertedRunStyle", json!(null), json!({"styleId": "C"})),
    ] {
        let root = concat!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/",
            "wordprocessingml/2006/main\"><w:body>",
        );
        let previous = "<w:p w:rsidR='87654321'><w:r><w:t>before</w:t></w:r></w:p>";
        let changed = "<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:t>x</w:t></w:r></w:p>";
        let keep = "<w:p w:rsidR='12345678'><w:r><w:t>keep</w:t></w:r></w:p>";
        let untouched = "<w:p w:rsidR='99999999'><w:r><w:t>untouched</w:t></w:r></w:p>";
        let source = format!("{root}{previous}{changed}{keep}{untouched}</w:body></w:document>");
        let units = vec![UnitV1::Text {
            text: "x".to_owned(),
            attrs: Attrs::new(),
        }];
        let inherit = |kind, value: &Value| {
            if scope == kind {
                value.clone()
            } else {
                json!({})
            }
        };
        let mut patch = json!({
            "ordinal": 1, "sessionKey": "source:1",
            "seed": {
                "units": units, "properties": {}, "unitSha256": unit_sha256(&units),
                "inherit": {
                    "paragraph": inherit("paragraph", &before),
                    "runs": [inherit("run", &before)],
                    "paragraphMark": inherit("mark", &before)
                }
            },
            "textEdits": [],
            "runEdits": [{
                "from": 0, "to": 1, "deltas": [], "target": {},
                "inherit": inherit("run", &after)
            }],
            "paraDeltas": [], "paraInherit": inherit("paragraph", &after),
            "markInherit": inherit("mark", &after),
            "resultSha256": result_sha256(&units, &Attrs::new())
        });
        let change = json!({
            "path": [scope], "before": {"state": "missing"},
            "after": {"state": "present", "value": after}
        });
        match scope {
            "numPr" | "pStyle" => patch["paraDeltas"] = json!([change]),
            "runStyle" => patch["runEdits"][0]["deltas"] = json!([change]),
            "insertedRunStyle" => {
                patch["runEdits"] = json!([]);
                patch["textEdits"] = json!([{
                    "from": 0, "to": 1, "donor": {"kind": "at", "at": 0},
                    "insert": [{"text": "x", "attrs": {"runStyle": after}, "inherit": {}}]
                }]);
            }
            _ => {}
        }
        let content: Vec<_> = ["before", "x", "keep", "untouched"]
            .into_iter()
            .enumerate()
            .map(|(ordinal, text)| {
                let mut paragraph = text_paragraph(text, None);
                paragraph["sourceOrdinal"] = json!(ordinal);
                paragraph
            })
            .collect();
        let request = serde_json::from_value(json!({
            "determinism": determinism(),
            "document": {"content": content},
            "options": {"updateModifiedDate": false},
            "paragraphIds": {"splicedParts": [{
                "part": "word/document.xml",
                "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
                "paragraphs": [0, 1, 2, 3], "changed": [1],
                "paragraphPatches": [patch]
            }]}
        }))
        .unwrap();
        let (saved, warnings) = write_docx_s13_with_warnings(request, &base_package(&source)).unwrap();
        let output = String::from_utf8(part_map(&saved)["word/document.xml"].clone()).unwrap();
        assert_eq!(
            output,
            format!(
                "{root}{previous}<w:p><w:r><w:t>x</w:t></w:r></w:p>\
                 {keep}{untouched}</w:body></w:document>"
            ),
            "{scope}"
        );
        assert_eq!(
            warnings,
            ["DOCX_MINIMAL_PATCH replace counts {\"restyle\":1}"],
            "{scope}"
        );
    }
}
