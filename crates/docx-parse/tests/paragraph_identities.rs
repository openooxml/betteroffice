//! Paragraph IDs: identities and occurrences the parser records, and the
//! paragraph ID plan a save applies.

use std::sync::Arc;

use docx_parse::block::BlockContent;
use docx_parse::paragraph_identity::{W14_NAMESPACE, package_paragraph_ids};
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, S9WireEnvelope, parse_docx_s9_wire};
use quick_xml::NsReader;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const BODY: &str = concat!(
    r#"<w:p w14:paraId="0000000A"/><w:p w14:paraId="0000000a"/><w:p w14:paraId="xyz"/>"#,
    r#"<bofx:block><w:p w14:paraId="0000000B"/><w:p/></bofx:block><w:p w14:paraId="80000000"/><w:p/>"#
);

fn package(body: &str, extra: &[(&str, &str)]) -> Vec<u8> {
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_vec()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h" Type="{R}/header" Target="header1.xml"/></Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="{W}" xmlns:w14="{W14_NAMESPACE}" xmlns:r="{R}" xmlns:bofx="urn:bofx"><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="h"/></w:sectPr></w:body></w:document>"#).into_bytes()),
        ("word/header1.xml".to_owned(), format!(r#"<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>"#).into_bytes()),
    ];
    for (path, xml) in extra {
        parts.push(((*path).to_owned(), xml.as_bytes().to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parse(bytes: &[u8], source_ordinals: bool) -> S9WireEnvelope {
    parse_docx_s9_wire(
        bytes,
        S9ParseOptions {
            source_ordinals,
            ..S9ParseOptions::default()
        },
    )
    .unwrap()
}

fn save(bytes: &[u8], paragraph_ids: Value, edit: impl FnOnce(&mut [BlockContent])) -> Vec<u8> {
    let mut package = parse(bytes, true).document.package;
    edit(&mut package.document.content);
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {
            "content": package.document.content,
            "comments": package.document.comments,
        },
        "headerEntries": package.header_entries.unwrap_or_default(),
        "footnotes": package.footnotes.unwrap_or_default(),
        "footnoteSeparators": package.footnote_separators.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": paragraph_ids,
    }))
    .unwrap();
    write_docx_s13(request, bytes).unwrap()
}

fn part(docx: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(docx)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .map(|(_, bytes)| bytes)
        .unwrap_or_else(|| panic!("{name} missing"))
}

/// Each `w:p` element's paragraph ID, read with namespace resolution; every
/// start tag must be well formed with unique attributes.
fn saved_para_ids(xml: &[u8]) -> Vec<Option<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut ids = Vec::new();
    loop {
        match reader.read_resolved_event().unwrap() {
            (
                ResolveResult::Bound(Namespace(namespace)),
                Event::Start(element) | Event::Empty(element),
            ) if namespace == W.as_bytes() && element.local_name().as_ref() == b"p" => {
                let mut id = None;
                for attribute in element.attributes() {
                    let attribute = attribute.expect("attributes are well formed and unique");
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if local.as_ref() == b"paraId" {
                        assert_eq!(
                            namespace,
                            ResolveResult::Bound(Namespace(W14_NAMESPACE.as_bytes()))
                        );
                        assert!(id.is_none(), "one paraId per paragraph");
                        id = Some(String::from_utf8(attribute.value.to_vec()).unwrap());
                    }
                }
                ids.push(id);
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    ids
}

fn paragraph(blocks: &mut [BlockContent], index: usize) -> &mut docx_parse::Paragraph {
    let BlockContent::Paragraph(paragraph) = &mut blocks[index] else {
        panic!("block {index} is not a paragraph");
    };
    Arc::make_mut(paragraph)
}

#[test]
fn every_valid_occurrence_keeps_its_id_and_later_repeats_are_flagged() {
    let content = parse(&package(BODY, &[]), true)
        .document
        .package
        .document
        .content;
    let summary: Vec<_> = content
        .iter()
        .map(|block| match block {
            BlockContent::Paragraph(paragraph) => json!([
                paragraph.para_id,
                paragraph.repeated_para_id,
                paragraph.source_ordinal,
                paragraph
                    .extra_attributes
                    .iter()
                    .map(|attribute| format!("{}={}", attribute.name, attribute.value))
                    .collect::<Vec<_>>(),
            ]),
            BlockContent::RawXml(raw) => json!(["raw", raw.source_ordinal]),
            _ => panic!("unexpected block"),
        })
        .collect();
    assert_eq!(
        summary,
        [
            json!(["0000000A", null, 0, []]),
            json!(["0000000a", true, 1, []]),
            json!([null, null, 2, ["w14:paraId=xyz"]]),
            json!(["raw", 3]),
            json!([null, null, 5, ["w14:paraId=80000000"]]),
            json!([null, null, 6, []]),
        ]
    );
    let plain = parse(&package(BODY, &[]), false)
        .document
        .package
        .document
        .content;
    assert!(plain.iter().all(|block| match block {
        BlockContent::Paragraph(paragraph) => paragraph.source_ordinal.is_none(),
        BlockContent::RawXml(raw) => raw.source_ordinal.is_none(),
        _ => true,
    }));
}

#[test]
fn saving_keeps_authored_ids_until_a_typed_id_replaces_them() {
    let original = package(BODY, &[]);
    let saved = save(&original, json!(null), |_| {});
    let ids = saved_para_ids(&part(&saved, "word/document.xml"));
    assert_eq!(
        ids,
        [
            Some("0000000A"),
            Some("0000000a"),
            Some("xyz"),
            Some("0000000B"),
            None,
            Some("80000000"),
            None,
        ]
        .map(|id| id.map(str::to_owned))
    );
    let repaired = save(&original, json!(null), |blocks| {
        paragraph(blocks, 2).para_id = Some("00000011".to_owned());
    });
    assert_eq!(
        saved_para_ids(&part(&repaired, "word/document.xml"))[2].as_deref(),
        Some("00000011")
    );
}

#[test]
fn assignments_reach_source_paragraphs_and_retained_xml_of_serialized_parts() {
    let original = package(BODY, &[]);
    let saved = save(
        &original,
        json!({"assignments": [
            {"part": "word/document.xml", "ordinal": 2, "paraId": "00000021"},
            {"part": "word/document.xml", "ordinal": 4, "paraId": "00000022"},
            {"part": "word/header1.xml", "ordinal": 0, "paraId": "00000023"},
        ]}),
        |_| {},
    );
    let ids = saved_para_ids(&part(&saved, "word/document.xml"));
    assert_eq!(ids[2].as_deref(), Some("00000021"));
    assert_eq!(ids[4].as_deref(), Some("00000022"));
    assert_eq!(
        saved_para_ids(&part(&saved, "word/header1.xml")),
        [Some("00000023".to_owned())]
    );
    let rejected: Result<S13SaveRequest, _> = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": []},
        "paragraphIds": {"assignments": [{"part": "word/document.xml", "ordinal": 0, "paraId": "\"/><x"}]},
    }));
    assert!(write_docx_s13(rejected.unwrap(), &original).is_err());
}

#[test]
fn patched_parts_change_only_start_tags_and_root_declarations() {
    let original = package(BODY, &[]);
    let saved = save(
        &original,
        json!({"patchedParts": [
            {"part": "word/document.xml", "paraIds": [[1, "00000031"], [4, "00000032"]]},
            {"part": "word/header1.xml", "paraIds": [[0, "00000033"]]},
        ]}),
        |_| {},
    );
    let document = String::from_utf8(part(&original, "word/document.xml")).unwrap();
    assert_eq!(
        String::from_utf8(part(&saved, "word/document.xml")).unwrap(),
        document
            .replacen(r#"w14:paraId="0000000a""#, r#"w14:paraId="00000031""#, 1)
            .replacen(
                r#"<w:p w14:paraId="0000000B"/><w:p/>"#,
                r#"<w:p w14:paraId="0000000B"/><w:p w14:paraId="00000032"/>"#,
                1
            )
    );
    let header = String::from_utf8(part(&original, "word/header1.xml")).unwrap();
    assert_eq!(
        String::from_utf8(part(&saved, "word/header1.xml")).unwrap(),
        header
            .replacen(
                &format!(r#"xmlns:w="{W}""#),
                &format!(
                    r#"xmlns:w="{W}" xmlns:w14="{W14_NAMESPACE}" xmlns:mc="{}" mc:Ignorable="w14""#,
                    docx_parse::paragraph_identity::MC_NAMESPACE
                ),
                1
            )
            .replacen("<w:p>", r#"<w:p w14:paraId="00000033">"#, 1)
    );
    assert_eq!(
        saved_para_ids(&part(&saved, "word/header1.xml")),
        [Some("00000033".to_owned())]
    );
    for name in ["[Content_Types].xml", "word/_rels/document.xml.rels"] {
        assert_eq!(part(&saved, name), part(&original, name), "{name}");
    }
}

const COMMENTS: &str = concat!(
    r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">"#,
    r#"<w:comment w:id="1" w:author="A"><w:p w14:paraId="00000041"><w:r><w:t>Parent</w:t></w:r></w:p></w:comment>"#,
    r#"<w:comment w:id="2" w:author="B"><w:p w14:paraId="0000000A"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment>"#,
    r#"<w:comment w:id="3" w:author="C"><w:p><w:r><w:t>Bare</w:t></w:r></w:p></w:comment>"#,
    r#"</w:comments>"#
);
const COMMENTS_EXTENDED: &str = concat!(
    r#"<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">"#,
    r#"<w15:commentEx w15:paraId="00000041" w15:done="0"/><w15:commentEx w15:paraId="0000000A" w15:paraIdParent="00000041" w15:done="0"/>"#,
    r#"</w15:commentsEx>"#
);

#[test]
fn patched_comments_rename_their_companion_references() {
    let original = package(
        BODY,
        &[
            ("word/comments.xml", COMMENTS),
            ("word/commentsExtended.xml", COMMENTS_EXTENDED),
        ],
    );
    let saved = save(
        &original,
        json!({"patchedParts": [
            {"part": "word/comments.xml", "paraIds": [[0, "00000051"], [1, "00000052"], [2, "00000053"]]},
        ]}),
        |_| {},
    );
    assert_eq!(
        String::from_utf8(part(&saved, "word/comments.xml")).unwrap(),
        COMMENTS
            .replacen("00000041", "00000051", 1)
            .replacen("0000000A", "00000052", 1)
            .replacen(
                "<w:p><w:r><w:t>Bare",
                r#"<w:p w14:paraId="00000053"><w:r><w:t>Bare"#,
                1
            )
    );
    assert_eq!(
        String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap(),
        COMMENTS_EXTENDED
            .replace("00000041", "00000051")
            .replacen("0000000A", "00000052", 1)
    );
}

#[test]
fn ambiguous_comment_references_fall_back_to_serialized_comments() {
    let ambiguous = COMMENTS.replacen("00000041", "0000000A", 1);
    let original = package(
        BODY,
        &[
            ("word/comments.xml", ambiguous.as_str()),
            ("word/commentsExtended.xml", COMMENTS_EXTENDED),
        ],
    );
    let saved = save(
        &original,
        json!({"patchedParts": [{"part": "word/comments.xml", "paraIds": [[1, "00000052"]]}]}),
        |_| {},
    );
    let comments = part(&saved, "word/comments.xml");
    assert_ne!(comments, ambiguous.as_bytes());
    let ids = saved_para_ids(&comments);
    assert_eq!(ids.len(), 3);
    let bare = ids[2].clone().unwrap();
    let parts = ooxml_opc::unzip_parts(&original).unwrap();
    assert!(!package_paragraph_ids(&parts).contains(&u32::from_str_radix(&bare, 16).unwrap()));
}

#[test]
fn the_package_inventory_counts_values_numerically_across_parts() {
    let parts = ooxml_opc::unzip_parts(&package(BODY, &[("word/comments.xml", COMMENTS)])).unwrap();
    assert_eq!(
        package_paragraph_ids(&parts)
            .into_iter()
            .collect::<Vec<_>>(),
        vec![0x0A, 0x0B, 0x41]
    );
}

#[test]
fn paragraph_ids_resolve_their_namespace_on_parse_and_full_save() {
    let bytes = package(
        &format!(
            r#"<w:p xmlns:w14="urn:other" w14:paraId="0000000A"/><w:p xmlns:x="{W14_NAMESPACE}" x:paraId="0000000B"/>"#
        ),
        &[],
    );
    let content = parse(&bytes, true).document.package.document.content;
    let ids: Vec<_> = content
        .iter()
        .map(|block| match block {
            BlockContent::Paragraph(paragraph) => paragraph.para_id.clone(),
            _ => panic!("paragraph"),
        })
        .collect();
    assert_eq!(ids, [None, Some("0000000B".to_owned())]);
    let occurrences = docx_parse::paragraph_identity::paragraph_occurrences(
        std::str::from_utf8(&part(&bytes, "word/document.xml")).unwrap(),
    )
    .unwrap();
    assert_eq!(
        occurrences
            .iter()
            .map(|occurrence| occurrence.para_id.clone())
            .collect::<Vec<_>>(),
        ids
    );

    let saved = save(&bytes, Value::Null, |blocks| {
        paragraph(blocks, 0).para_id = Some("0000000C".to_owned());
    });
    let xml = String::from_utf8(part(&saved, "word/document.xml")).unwrap();
    assert!(xml.contains(&format!(
        r#"<w:p xmlns:w14p0="{W14_NAMESPACE}" w14p0:paraId="0000000C" xmlns:w14="urn:other" w14:paraId="0000000A""#
    )));
    let reopened = parse(&saved, true).document.package.document.content;
    let BlockContent::Paragraph(first) = &reopened[0] else {
        panic!("paragraph");
    };
    assert_eq!(first.para_id.as_deref(), Some("0000000C"));
}

#[test]
fn a_new_id_replaces_an_invalid_one_under_an_alias_prefix() {
    let bytes = package(
        &format!(
            r#"<w:p xmlns:x="{W14_NAMESPACE}" x:paraId="xyz"><w:r><w:t>Alias</w:t></w:r></w:p>"#
        ),
        &[],
    );
    let content = parse(&bytes, true).document.package.document.content;
    let BlockContent::Paragraph(alias) = &content[0] else {
        panic!("paragraph");
    };
    assert_eq!(alias.para_id, None);
    assert_eq!(alias.para_id_attribute.as_deref(), Some("x:paraId"));
    let saved = save(&bytes, Value::Null, |blocks| {
        paragraph(blocks, 0).para_id = Some("0000000C".to_owned());
    });
    assert_eq!(
        saved_para_ids(&part(&saved, "word/document.xml")),
        [Some("0000000C".to_owned())]
    );
}
