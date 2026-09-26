//! Paragraph identity: source bindings, allocation, ownership, structural
//! rules, anchors and the save plan.

use std::collections::HashSet;
use std::path::Path;

use docx_edit::*;
use docx_parse::paragraph_identity::{
    package_paragraph_ids, paragraph_ids_by_part, parse_paragraph_id,
};
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use serde_json::json;
use yrs::Any;

const DATE: &str = "2026-09-24T00:00:00Z";
const BODY: &str = "/word/document.xml";

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn fixture_parts() -> Vec<(String, Vec<u8>)> {
    fn collect(root: &Path, dir: &Path, parts: &mut Vec<(String, Vec<u8>)>) {
        for entry in std::fs::read_dir(dir).unwrap() {
            let path = entry.unwrap().path();
            if path.is_dir() {
                collect(root, &path, parts);
            } else {
                let name = path.strip_prefix(root).unwrap().to_string_lossy();
                parts.push((name.replace('\\', "/"), std::fs::read(&path).unwrap()));
            }
        }
    }
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/paragraph-identities");
    let mut parts = Vec::new();
    collect(&root, &root, &mut parts);
    parts.sort();
    parts
}

fn text(parts: &[(String, Vec<u8>)], name: &str) -> String {
    let (_, bytes) = parts.iter().find(|(path, _)| path == name).unwrap();
    String::from_utf8(bytes.clone()).unwrap()
}

/// The fixture with some parts' XML rewritten.
fn fixture_with(edit: impl FnOnce(&mut Vec<(String, Vec<u8>)>)) -> Vec<u8> {
    let mut parts = fixture_parts();
    edit(&mut parts);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn replace(parts: &mut [(String, Vec<u8>)], name: &str, from: &str, to: &str) {
    let (_, bytes) = parts.iter_mut().find(|(path, _)| path == name).unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert!(xml.contains(from), "{name} lacks {from}");
    *bytes = xml.replacen(from, to, 1).into_bytes();
}

fn fixture() -> Vec<u8> {
    fixture_with(|_| {})
}

fn seeded(bytes: &[u8]) -> EditingDoc {
    seeded_by(41, bytes)
}

fn seeded_by(client_id: u64, bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(client_id);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn session(story: &str, key: &str) -> ParagraphRef {
    ParagraphRef::Session {
        story: story.to_owned(),
        para_id: key.to_owned(),
    }
}

fn source_ref(doc: &EditingDoc, part_uri: &str, ordinal: u32) -> ParagraphRef {
    ParagraphRef::Source(SourceParagraphRef {
        package_sha256: doc.paragraph_identities().package_sha256.unwrap(),
        part_uri: part_uri.to_owned(),
        paragraph_ordinal: ordinal,
    })
}

fn story(part_uri: &str, kind: SourceStoryKind, item_id: Option<&str>) -> SourceStory {
    SourceStory {
        part_uri: part_uri.to_owned(),
        kind,
        item_id: item_id.map(str::to_owned),
    }
}

fn body() -> SourceStory {
    story(BODY, SourceStoryKind::Body, None)
}

fn identity(doc: &EditingDoc, key: &str) -> ParagraphIdentity {
    doc.paragraph_identities()
        .paragraphs
        .into_iter()
        .find(|paragraph| {
            matches!(&paragraph.paragraph, ParagraphRef::Session { para_id, .. } if para_id == key)
        })
        .unwrap_or_else(|| panic!("paragraph {key:?} not found"))
}

fn identity_text(doc: &EditingDoc, key: &str) -> String {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.para_id == key)
        .unwrap()
        .text
}

fn persisted(story: SourceStory, para_id: &str) -> ParagraphAnchor {
    ParagraphAnchor::Persisted {
        story,
        para_id: para_id.to_owned(),
    }
}

fn found(story: &str, key: &str) -> AnchorResolution {
    AnchorResolution::Found(session(story, key))
}

fn saved_ids(doc: &EditingDoc) -> Vec<String> {
    doc.paragraph_identities()
        .paragraphs
        .into_iter()
        .filter_map(|paragraph| paragraph.ooxml_para_id)
        .collect()
}

#[test]
fn seeding_records_every_occurrence_with_its_own_source_id() {
    let doc = seeded(&fixture());
    let identities = doc.paragraph_identities();
    assert_eq!(identities.session_id, doc.session_id());
    let sha = identities.package_sha256.clone().unwrap();
    assert_eq!(sha.len(), 64);
    let summary: Vec<_> = identities
        .paragraphs
        .iter()
        .map(|paragraph| {
            (
                paragraph.paragraph.clone(),
                paragraph.ooxml_para_id.clone(),
                paragraph
                    .source
                    .as_ref()
                    .map(|source| (source.part_uri.clone(), source.paragraph_ordinal)),
            )
        })
        .collect();
    let occurrence = |part: &str, ordinal| Some((part.to_owned(), ordinal));
    let id = |value: &str| Some(value.to_owned());
    assert_eq!(
        summary,
        [
            (
                session("body", "1A2B3C4D"),
                id("1A2B3C4D"),
                occurrence(BODY, 0)
            ),
            (session("body", "body:p1"), None, occurrence(BODY, 1)),
            (
                session("body", "0000abcd"),
                id("0000abcd"),
                occurrence(BODY, 2)
            ),
            (
                session("body", "body:p3"),
                id("1a2b3c4d"),
                occurrence(BODY, 3)
            ),
            (session("body", "body:p4"), None, occurrence(BODY, 4)),
            (
                session("body", "0A0B0C0D"),
                id("0A0B0C0D"),
                occurrence(BODY, 7)
            ),
            (
                session("body:t0:r0c0", "2B3C4D5E"),
                id("2B3C4D5E"),
                occurrence(BODY, 5)
            ),
            (
                session("fn:1", "3C4D5E6F"),
                id("3C4D5E6F"),
                occurrence("/word/footnotes.xml", 2)
            ),
            (
                session("hf:rIdHeader", "6F7A8B9C"),
                id("6F7A8B9C"),
                occurrence("/word/header1.xml", 0)
            ),
            (
                source_ref(&doc, BODY, 6),
                id("5E6F7A8B"),
                occurrence(BODY, 6)
            ),
            (
                source_ref(&doc, "/word/footnotes.xml", 0),
                None,
                occurrence("/word/footnotes.xml", 0)
            ),
            (
                source_ref(&doc, "/word/footnotes.xml", 1),
                None,
                occurrence("/word/footnotes.xml", 1)
            ),
            (
                source_ref(&doc, "/word/comments.xml", 0),
                id("4D5E6F7A"),
                occurrence("/word/comments.xml", 0)
            ),
        ]
    );
    assert!(identities.paragraphs.iter().all(|paragraph| {
        paragraph.origin == ParagraphOrigin::Source
            && paragraph.id_origin
                == paragraph
                    .ooxml_para_id
                    .as_ref()
                    .map(|_| ParagraphIdOrigin::Source)
    }));
    assert_eq!(
        doc.story_paragraph_ids("body").unwrap(),
        summary[..6]
            .iter()
            .map(|(_, id, _)| id.clone())
            .collect::<Vec<_>>()
    );
    assert_eq!(identity(&doc, "2B3C4D5E").source_story, Some(body()));
    assert_eq!(
        identity(&doc, "3C4D5E6F").source_story,
        Some(story(
            "/word/footnotes.xml",
            SourceStoryKind::Footnote,
            Some("1")
        ))
    );
    assert_eq!(
        identities.paragraphs[12].source_story,
        Some(story(
            "/word/comments.xml",
            SourceStoryKind::Comment,
            Some("1")
        ))
    );
    assert_eq!(
        identities.paragraphs[10].source_story,
        Some(story(
            "/word/footnotes.xml",
            SourceStoryKind::Footnote,
            Some("-1")
        ))
    );
}

#[test]
fn repeated_source_ids_resolve_as_ambiguous_within_their_story() {
    let doc = seeded(&fixture());
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "1a2b3c4d")),
        AnchorResolution::Ambiguous(vec![
            session("body", "1A2B3C4D"),
            session("body", "body:p3")
        ])
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&ParagraphAnchor::Source(SourceParagraphRef {
            package_sha256: doc.paragraph_identities().package_sha256.unwrap(),
            part_uri: BODY.into(),
            paragraph_ordinal: 3,
        })),
        found("body", "body:p3")
    );

    let repeated_in_header = fixture_with(|parts| {
        replace(parts, "word/header1.xml", "6F7A8B9C", "2B3C4D5E");
    });
    let doc = seeded(&repeated_in_header);
    let header = story("/word/header1.xml", SourceStoryKind::Header, None);
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(header, "2B3C4D5E")),
        found("hf:rIdHeader", "hf:rIdHeader:p0")
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "2B3C4D5E")),
        found("body:t0:r0c0", "2B3C4D5E")
    );
}

#[test]
fn persistence_covers_every_source_part_and_repairs_duplicates_by_package_order() {
    let bytes = fixture();
    let doc = seeded(&bytes);
    let occupied = package_paragraph_ids(&ooxml_opc::unzip_parts(&bytes).unwrap());
    let report = doc.persist_paragraph_ids().unwrap();
    assert!(report.diagnostics.is_empty());
    let summary: Vec<_> = report
        .assignments
        .iter()
        .map(|assignment| {
            (
                assignment.paragraph.clone(),
                assignment.previous_ooxml_para_id.clone(),
                assignment.origin,
            )
        })
        .collect();
    assert_eq!(
        summary,
        [
            (
                session("body", "body:p3"),
                Some("1a2b3c4d".to_owned()),
                ParagraphIdOrigin::Repaired
            ),
            (
                session("body", "body:p1"),
                None,
                ParagraphIdOrigin::Persisted
            ),
            (
                session("body", "body:p4"),
                None,
                ParagraphIdOrigin::Persisted
            ),
            (
                source_ref(&doc, "/word/footnotes.xml", 0),
                None,
                ParagraphIdOrigin::Persisted
            ),
            (
                source_ref(&doc, "/word/footnotes.xml", 1),
                None,
                ParagraphIdOrigin::Persisted
            ),
        ]
    );
    for assignment in &report.assignments {
        let id = parse_paragraph_id(&assignment.ooxml_para_id).unwrap();
        assert_eq!(assignment.ooxml_para_id, format!("{id:08X}"));
        assert!((1..=0x7FFF_FFFE).contains(&id) && !occupied.contains(&id));
        assert!(assignment.source_story.is_some());
    }
    assert_eq!(
        report.assignments[3].source_story,
        Some(story(
            "/word/footnotes.xml",
            SourceStoryKind::Footnote,
            Some("-1")
        ))
    );
    let ids = saved_ids(&doc);
    assert_eq!(ids.len(), 13);
    assert_eq!(
        ids.iter()
            .map(|id| parse_paragraph_id(id).unwrap())
            .collect::<HashSet<_>>()
            .len(),
        13
    );
    assert_eq!(identity(&doc, "body:p3").origin, ParagraphOrigin::Source);
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "1A2B3C4D")),
        found("body", "1A2B3C4D")
    );
    let state = doc.encode_state_vector_v1();
    assert_eq!(
        doc.persist_paragraph_ids().unwrap(),
        PersistedParagraphIds::default()
    );
    assert_eq!(doc.encode_state_vector_v1(), state);
}

#[test]
fn ambiguous_comment_references_refuse_persistence_without_changes() {
    let comments = |parts: &mut Vec<(String, Vec<u8>)>| {
        replace(
            parts,
            "word/comments.xml",
            "</w:comments>",
            r#"<w:comment w:id="2" w:author="Reviewer"><w:p w14:paraId="4D5E6F7A"><w:r><w:t>Twin</w:t></w:r></w:p></w:comment></w:comments>"#,
        );
    };
    let doc = seeded(&fixture_with(comments));
    let state = doc.encode_state_vector_v1();
    assert_eq!(
        doc.persist_paragraph_ids(),
        Err(ParagraphIdRefusal::AmbiguousCommentReference {
            ooxml_para_id: "4D5E6F7A".into(),
            comment_ids: vec!["1".into(), "2".into()],
        })
    );
    assert_eq!(doc.encode_state_vector_v1(), state);

    let unreferenced = fixture_with(|parts| {
        comments(parts);
        replace(parts, "word/commentsExtended.xml", "4D5E6F7A", "7A6B5C4D");
    });
    let doc = seeded(&unreferenced);
    let report = doc.persist_paragraph_ids().unwrap();
    let comment = source_ref(&doc, "/word/comments.xml", 1);
    let repaired = report
        .assignments
        .iter()
        .find(|assignment| assignment.paragraph == comment)
        .unwrap();
    assert_eq!(repaired.origin, ParagraphIdOrigin::Repaired);
    assert_eq!(repaired.previous_ooxml_para_id.as_deref(), Some("4D5E6F7A"));
    assert_eq!(
        repaired.source_story,
        Some(story(
            "/word/comments.xml",
            SourceStoryKind::Comment,
            Some("2")
        ))
    );

    let saved = save_unchanged(&doc, &unreferenced);
    let original = ooxml_opc::unzip_parts(&unreferenced).unwrap();
    let comments = text(&original, "word/comments.xml");
    let twin = comments.rfind("4D5E6F7A").unwrap();
    let mut expected = comments.clone();
    expected.replace_range(twin..twin + 8, &repaired.ooxml_para_id);
    assert_eq!(text(&saved, "word/comments.xml"), expected);
    assert_eq!(
        text(&saved, "word/commentsExtended.xml"),
        text(&original, "word/commentsExtended.xml")
    );
}

#[test]
fn identities_follow_split_merge_and_deletion_survivors() {
    let doc = seeded(&fixture());
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 0), None)
        .unwrap();
    assert_eq!(split.first_para_id, "1A2B3C4D");
    assert_eq!(doc.paragraphs("body").unwrap()[0].text, "");
    let second = identity(&doc, &split.second_para_id);
    let second_id = second.ooxml_para_id.clone().unwrap();
    assert_eq!(
        (second.origin, second.id_origin, second.source),
        (
            ParagraphOrigin::Authored,
            Some(ParagraphIdOrigin::Authored),
            None
        )
    );
    let first = identity(&doc, "1A2B3C4D");
    assert_eq!(first.ooxml_para_id.as_deref(), Some("1A2B3C4D"));
    assert_eq!(first.source.map(|source| source.paragraph_ordinal), Some(0));
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), &second_id)),
        found("body", &split.second_para_id)
    );

    let end = doc.paragraph_mark_position("0000abcd").unwrap().index;
    let at_end = doc
        .split_paragraph(&ctx(), Position::new("body", end), None)
        .unwrap();
    assert_eq!(at_end.first_para_id, "0000abcd");
    assert_eq!(
        identity(&doc, "0000abcd").ooxml_para_id.as_deref(),
        Some("0000abcd")
    );
    let empty = identity(&doc, &at_end.second_para_id);
    assert!(empty.ooxml_para_id.is_some() && empty.source.is_none());

    doc.merge_paragraphs(&ctx(), "1A2B3C4D", MergeDirection::Forward)
        .unwrap();
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), &second_id)),
        AnchorResolution::Missing
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "1A2B3C4D")),
        AnchorResolution::Ambiguous(vec![
            session("body", "1A2B3C4D"),
            session("body", "body:p3")
        ])
    );

    let lower = doc.paragraph_mark_position("0000abcd").unwrap();
    let missing = doc.paragraph_mark_position("body:p4").unwrap();
    doc.delete_range(
        &ctx(),
        StoryRange::new("body", lower.index - 1, missing.index - 1),
    )
    .unwrap();
    assert_eq!(
        identity(&doc, "0000abcd").ooxml_para_id.as_deref(),
        Some("0000abcd")
    );
    let keys: HashSet<_> = doc
        .paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.para_id)
        .collect();
    assert!(!keys.contains("body:p3") && !keys.contains(&at_end.second_para_id));
}

#[test]
fn tracked_joins_keep_the_following_identity() {
    let doc = seeded(&fixture());
    let receipt = doc
        .merge_paragraphs(
            &EditCtx::local("Ada", DATE).suggesting(),
            "0000abcd",
            MergeDirection::Forward,
        )
        .unwrap();
    doc.accept_change(
        &ctx(),
        &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
    )
    .unwrap();
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "0000abcd")),
        AnchorResolution::Missing
    );
    assert_eq!(
        identity(&doc, "body:p3").ooxml_para_id.as_deref(),
        Some("1a2b3c4d")
    );
}

#[test]
fn table_cells_authored_in_a_session_get_ids_in_their_root_story() {
    let doc = seeded(&fixture());
    let table = doc
        .insert_table(&ctx(), Position::new("body", 0), 2, 2)
        .unwrap();
    assert_eq!(table.new_para_ids.len(), 4);
    let ids: HashSet<String> = table
        .new_para_ids
        .iter()
        .map(|key| {
            let cell = identity(&doc, key);
            assert_eq!(cell.id_origin, Some(ParagraphIdOrigin::Authored));
            assert_eq!(cell.source_story, Some(body()));
            cell.ooxml_para_id.unwrap()
        })
        .collect();
    assert_eq!(ids.len(), 4);
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "2B3C4D5E")),
        found("body:t0:r0c0", "2B3C4D5E"),
        "a nested anchor survives the positional shift of an inserted table"
    );
}

#[test]
fn anchors_are_scoped_to_their_session_package_and_story() {
    let doc = seeded(&fixture());
    let session_id = doc.session_id();
    let anchor = |story: &str, key: &str| ParagraphAnchor::Session {
        session_id: session_id.clone(),
        story: story.into(),
        para_id: key.into(),
    };
    assert_eq!(
        doc.resolve_paragraph_anchor(&anchor("body", "body:p1")),
        found("body", "body:p1")
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&anchor("hf:rIdHeader", "body:p1")),
        AnchorResolution::Missing
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&ParagraphAnchor::Session {
            session_id: seeded_by(42, &fixture()).session_id(),
            story: "body".into(),
            para_id: "body:p1".into(),
        }),
        AnchorResolution::Unsupported(AnchorUnsupported::ForeignSession)
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&ParagraphAnchor::Source(SourceParagraphRef {
            package_sha256: "0".repeat(64),
            part_uri: BODY.into(),
            paragraph_ordinal: 1,
        })),
        AnchorResolution::Unsupported(AnchorUnsupported::ForeignPackage)
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(
            story("/word/header1.xml", SourceStoryKind::Header, None),
            "6F7A8B9C"
        )),
        found("hf:rIdHeader", "6F7A8B9C")
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "6F7A8B9C")),
        AnchorResolution::Missing
    );
    let comment = story("/word/comments.xml", SourceStoryKind::Comment, Some("1"));
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(comment, "4D5E6F7A")),
        AnchorResolution::Found(source_ref(&doc, "/word/comments.xml", 0))
    );

    let detached = EditingDoc::new(2);
    detached
        .create_story("body", "plain", "Normal", "left")
        .unwrap();
    assert_eq!(
        detached.resolve_paragraph_anchor(&persisted(body(), "2B3C4D5E")),
        AnchorResolution::Unsupported(AnchorUnsupported::NoSourcePackage)
    );
}

#[test]
fn copies_take_new_ids_and_the_original_keeps_its_own() {
    let doc = seeded(&fixture());
    let authored = doc
        .split_paragraph(&ctx(), Position::new("body", 1), None)
        .unwrap()
        .second_para_id;
    let authored_id = identity(&doc, &authored).ooxml_para_id.unwrap();
    let copy = |key: &str, id: &str| RawOp::InsertEmbed {
        index: 0,
        kind: "pilcrow".into(),
        payload: vec![
            ("paraId".into(), Any::from(key)),
            ("ooxmlParaId".into(), Any::from(id)),
            ("sourceParaId".into(), Any::from(id)),
        ],
        attrs: Default::default(),
    };
    doc.apply_raw_ops(
        "body",
        vec![
            copy("cell-copy", "2b3c4d5e"),
            copy("authored-copy", &authored_id),
        ],
        &ctx(),
    )
    .unwrap();
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "2B3C4D5E")),
        found("body:t0:r0c0", "2B3C4D5E")
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), &authored_id)),
        found("body", &authored)
    );
    for key in ["cell-copy", "authored-copy"] {
        let copied = identity(&doc, key);
        assert_eq!(copied.id_origin, Some(ParagraphIdOrigin::Repaired));
        assert_eq!(copied.origin, ParagraphOrigin::Authored);
    }
    let ids = saved_ids(&doc);
    assert_eq!(ids.iter().collect::<HashSet<_>>().len(), ids.len());

    let lower = doc.paragraph_mark_position("0000abcd").unwrap().index - 5;
    doc.apply_raw_ops(
        "body",
        vec![RawOp::InsertEmbed {
            index: lower,
            kind: "pilcrow".into(),
            payload: vec![
                ("paraId".into(), Any::from("0000abcd")),
                ("ooxmlParaId".into(), Any::from("0000abcd")),
            ],
            attrs: Default::default(),
        }],
        &ctx(),
    )
    .unwrap();
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), "0000abcd")),
        found("body", "0000abcd")
    );
    let paragraphs = doc.paragraphs("body").unwrap();
    let original = paragraphs
        .iter()
        .position(|paragraph| paragraph.para_id == "0000abcd")
        .unwrap();
    assert_eq!(paragraphs[original].text, "Lower");
    let copy = &paragraphs[original - 1];
    assert_eq!(copy.text, "");
    let copied = identity(&doc, &copy.para_id);
    assert_eq!(copied.id_origin, Some(ParagraphIdOrigin::Repaired));
    assert_ne!(copied.ooxml_para_id.as_deref(), Some("0000abcd"));

    let authored_text = identity_text(&doc, &authored);
    let start = doc.paragraph_mark_position("1A2B3C4D").unwrap().index + 1;
    doc.apply_raw_ops(
        "body",
        vec![RawOp::InsertEmbed {
            index: start,
            kind: "pilcrow".into(),
            payload: vec![
                ("paraId".into(), Any::from(authored.as_str())),
                ("ooxmlParaId".into(), Any::from(authored_id.as_str())),
            ],
            attrs: Default::default(),
        }],
        &ctx(),
    )
    .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    let original = paragraphs
        .iter()
        .position(|paragraph| paragraph.para_id == authored)
        .unwrap();
    assert_eq!(paragraphs[original].text, authored_text);
    assert_eq!(
        doc.resolve_paragraph_anchor(&persisted(body(), &authored_id)),
        found("body", &authored)
    );
    let copy = identity(&doc, &paragraphs[original - 1].para_id);
    assert_eq!(copy.id_origin, Some(ParagraphIdOrigin::Repaired));
    assert_ne!(copy.ooxml_para_id, Some(authored_id));
}

#[test]
fn identity_keys_are_schema_managed() {
    let doc = seeded(&fixture());
    for key in ["ooxmlParaId", "sourceParaId", "paraOrigin"] {
        assert!(matches!(
            doc.set_paragraph_attr("body:p1", key, Any::from("11111111")),
            Err(EditError::ReservedParagraphKey(_))
        ));
        let delta = ParaAttrDelta {
            other: [(key.to_owned(), Some(Any::from("11111111")))].into(),
            ..ParaAttrDelta::default()
        };
        assert!(matches!(
            doc.set_paragraph_attrs(&ctx(), &ParaSelector::One("body:p1".into()), &delta),
            Err(OpError::ReservedKey(_))
        ));
        let mark = doc.paragraph_mark_position("body:p1").unwrap();
        assert!(matches!(
            doc.apply_raw_ops(
                "body",
                vec![RawOp::SetEmbedAttr {
                    index: mark.index,
                    key: key.into(),
                    value: Any::from("11111111"),
                }],
                &ctx(),
            ),
            Err(OpError::ReservedKey(_))
        ));
    }
    doc.apply_raw_ops(
        "body",
        vec![RawOp::InsertEmbed {
            index: 0,
            kind: "pilcrow".into(),
            payload: vec![
                ("paraId".into(), Any::from("raw")),
                ("ooxmlParaId".into(), Any::from("body:p9")),
                ("paraOrigin".into(), Any::from("synthetic")),
            ],
            attrs: Default::default(),
        }],
        &ctx(),
    )
    .unwrap();
    let raw = identity(&doc, "raw");
    assert_eq!(
        (raw.ooxml_para_id, raw.origin),
        (None, ParagraphOrigin::Authored)
    );

    let suggested = EditCtx::local("Ada", DATE).suggesting();
    let align = ParaAttrDelta {
        alignment: Patch::Set("center".into()),
        ..ParaAttrDelta::default()
    };
    let receipt = doc
        .set_paragraph_attrs(&suggested, &ParaSelector::One("0000abcd".into()), &align)
        .unwrap();
    doc.reject_change(
        &ctx(),
        &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
    )
    .unwrap();
    assert_eq!(
        identity(&doc, "0000abcd").ooxml_para_id.as_deref(),
        Some("0000abcd")
    );
}

#[test]
fn identity_reads_leave_the_document_unchanged() {
    let doc = seeded(&fixture());
    let state = doc.encode_state_as_update_v1();
    let identities = doc.paragraph_identities();
    for paragraph in &identities.paragraphs {
        let anchor = match &paragraph.paragraph {
            ParagraphRef::Session { story, para_id } => ParagraphAnchor::Session {
                session_id: identities.session_id.clone(),
                story: story.clone(),
                para_id: para_id.clone(),
            },
            ParagraphRef::Source(source) => ParagraphAnchor::Source(source.clone()),
        };
        assert!(matches!(
            doc.resolve_paragraph_anchor(&anchor),
            AnchorResolution::Found(found) if found == paragraph.paragraph
        ));
    }
    doc.paragraph_save_plan();
    doc.story_paragraph_ids("body").unwrap();
    assert_eq!(doc.encode_state_as_update_v1(), state);
}

#[test]
fn replicas_share_the_session_and_repair_concurrent_duplicates_identically() {
    let bytes = fixture();
    let origin = seeded(&bytes);
    let update = origin.encode_state_as_update_v1();
    let left = EditingDoc::new(51);
    let right = EditingDoc::new(52);
    left.apply_update_v1(&update).unwrap();
    right.apply_update_v1(&update).unwrap();
    assert_eq!(left.session_id(), origin.session_id());
    let at = origin.paragraph_mark_position("0000abcd").unwrap().index - 2;
    left.split_paragraph(&ctx(), Position::new("body", at), None)
        .unwrap();
    right
        .split_paragraph(&ctx(), Position::new("body", at + 1), None)
        .unwrap();
    let from_left = left.encode_state_as_update_v1();
    let from_right = right.encode_state_as_update_v1();
    right.apply_update_v1(&from_left).unwrap();
    left.apply_update_v1(&from_right).unwrap();
    let converged = saved_ids(&left);
    assert_eq!(converged, saved_ids(&right));
    assert_eq!(
        converged.iter().filter(|id| *id == "0000abcd").count(),
        1,
        "the source ID stays with one paragraph"
    );
    let identities = left.paragraph_identities();
    let keys: HashSet<_> = identities
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.paragraph.clone())
        .collect();
    assert_eq!(keys.len(), identities.paragraphs.len());

    let left_split = left.paragraphs("body").unwrap()[3].para_id.clone();
    let anchor = ParagraphAnchor::Session {
        session_id: left.session_id(),
        story: "body".into(),
        para_id: left_split.clone(),
    };
    right.retain_source_docx(bytes.clone());
    assert_eq!(
        right.resolve_paragraph_anchor(&anchor),
        found("body", &left_split)
    );
    assert_eq!(
        origin.resolve_paragraph_anchor(&anchor),
        AnchorResolution::Missing,
        "the same session, before the split arrives"
    );
    let reopened = seeded_by(53, &bytes);
    assert_eq!(
        reopened.resolve_paragraph_anchor(&anchor),
        AnchorResolution::Unsupported(AnchorUnsupported::ForeignSession)
    );
}

#[test]
fn a_hydrated_replica_indexes_its_retained_package_on_first_use() {
    let bytes = fixture();
    let legacy = EditingDoc::new(3);
    legacy
        .create_story_with_paragraph_id("body", "1A2B3C4D", "Valid", "Normal", "left")
        .unwrap();
    assert_eq!(identity(&legacy, "1A2B3C4D").ooxml_para_id, None);
    legacy.retain_source_docx(bytes.clone());
    let bound = identity(&legacy, "1A2B3C4D");
    assert_eq!(bound.ooxml_para_id.as_deref(), Some("1A2B3C4D"));
    assert_eq!(bound.id_origin, Some(ParagraphIdOrigin::Source));
    assert_eq!(
        (
            bound.source_story,
            bound.source.map(|source| source.paragraph_ordinal)
        ),
        (Some(body()), Some(0))
    );

    let hydrated = EditingDoc::new(4);
    hydrated
        .apply_update_v1(&seeded(&bytes).encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        identity(&hydrated, "3C4D5E6F").ooxml_para_id.as_deref(),
        Some("3C4D5E6F")
    );
    hydrated.retain_source_docx(bytes);
    assert_eq!(
        hydrated.resolve_paragraph_anchor(&persisted(
            story("/word/footnotes.xml", SourceStoryKind::Footnote, Some("1")),
            "3C4D5E6F"
        )),
        found("fn:1", "3C4D5E6F")
    );
}

#[test]
fn editor_only_paragraphs_are_not_persisted_until_authored() {
    let bytes = fixture_with(|parts| {
        replace(
            parts,
            "word/document.xml",
            r#"<w:p w14:paraId="0A0B0C0D" w:rsidR="00A1B2C3"><w:r><w:t>Tail</w:t></w:r></w:p>"#,
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="0A0B0C0D"/></w:tc></w:tr></w:tbl>"#,
        );
        replace(
            parts,
            "word/header1.xml",
            r#"<w:p w14:paraId="6F7A8B9C"><w:r><w:t>Header</w:t></w:r></w:p>"#,
            "",
        );
    });
    let doc = seeded(&bytes);
    let tail = doc
        .paragraphs("body")
        .unwrap()
        .last()
        .unwrap()
        .para_id
        .clone();
    let header = doc.paragraphs("hf:rIdHeader").unwrap()[0].para_id.clone();
    for key in [&tail, &header] {
        let sentinel = identity(&doc, key);
        assert_eq!(
            (sentinel.origin, sentinel.ooxml_para_id, sentinel.source),
            (ParagraphOrigin::Synthetic, None, None)
        );
    }
    let hydrated = EditingDoc::new(5);
    hydrated
        .apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        identity(&hydrated, &tail).origin,
        ParagraphOrigin::Synthetic
    );

    let report = doc.persist_paragraph_ids().unwrap();
    assert!(report.assignments.iter().all(|assignment| {
        assignment.paragraph != session("body", &tail)
            && assignment.paragraph != session("hf:rIdHeader", &header)
    }));
    assert_eq!(identity(&doc, &tail).ooxml_para_id, None);

    let mut undo = doc.undo_manager();
    let at = doc.paragraph_mark_position(&tail).unwrap();
    doc.insert_text(&ctx(), at, "typed", FormatPolicy::Plain)
        .unwrap();
    let promoted = identity(&doc, &tail);
    assert_eq!(promoted.origin, ParagraphOrigin::Authored);
    assert_eq!(promoted.id_origin, Some(ParagraphIdOrigin::Authored));
    assert!(undo.undo());
    let reverted = identity(&doc, &tail);
    assert_eq!(
        (reverted.origin, reverted.ooxml_para_id),
        (ParagraphOrigin::Synthetic, None)
    );
    let at = doc.paragraph_mark_position(&tail).unwrap();
    doc.insert_text(&ctx(), at, "again", FormatPolicy::Plain)
        .unwrap();
    let retyped = identity(&doc, &tail);
    assert_eq!(retyped.origin, ParagraphOrigin::Authored);
    assert!(retyped.ooxml_para_id.is_some());

    let touched = seeded(&bytes);
    let tail_mark = touched.paragraph_mark_position(&tail).unwrap();
    let before_tables = Position::new("body", tail_mark.index - 1);
    touched
        .insert_text(&ctx(), before_tables, "early", FormatPolicy::Plain)
        .unwrap();
    assert_eq!(identity(&touched, &tail).origin, ParagraphOrigin::Synthetic);
    let report = touched.persist_paragraph_ids().unwrap();
    assert!(report.assignments.iter().any(|assignment| {
        assignment.paragraph == session("body", &tail)
            && assignment.origin == ParagraphIdOrigin::Persisted
    }));
    assert_eq!(identity(&touched, &tail).origin, ParagraphOrigin::Authored);
    let at = doc.paragraph_mark_position(&header).unwrap();
    let split = doc.split_paragraph(&ctx(), at, None).unwrap();
    for key in [&header, &split.second_para_id] {
        assert_eq!(identity(&doc, key).origin, ParagraphOrigin::Authored);
        assert!(identity(&doc, key).ooxml_para_id.is_some());
    }
    let state = doc.encode_state_as_update_v1();
    assert!(
        !String::from_utf8_lossy(&state).contains("sourceOrdinal"),
        "occurrence ordinals are an index concern, never seeded"
    );
}

#[test]
fn raw_ops_authoring_into_an_editor_only_paragraph_promote_it() {
    let bytes = fixture_with(|parts| {
        replace(
            parts,
            "word/document.xml",
            r#"<w:p w14:paraId="0A0B0C0D" w:rsidR="00A1B2C3"><w:r><w:t>Tail</w:t></w:r></w:p>"#,
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="0A0B0C0D"/></w:tc></w:tr></w:tbl>"#,
        );
    });
    let authoring: [fn(u32) -> RawOp; 3] = [
        |index| RawOp::Insert {
            index,
            text: "typed".into(),
            attrs: Default::default(),
        },
        |index| RawOp::InsertEmbed {
            index,
            kind: "break".into(),
            payload: Vec::new(),
            attrs: Default::default(),
        },
        |index| RawOp::SetEmbedAttr {
            index,
            key: "alignment".into(),
            value: Any::from("center"),
        },
    ];
    for author in authoring {
        let doc = seeded(&bytes);
        let tail = doc
            .paragraphs("body")
            .unwrap()
            .last()
            .unwrap()
            .para_id
            .clone();
        doc.apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: 0,
                text: "early".into(),
                attrs: Default::default(),
            }],
            &ctx(),
        )
        .unwrap();
        assert_eq!(identity(&doc, &tail).origin, ParagraphOrigin::Synthetic);
        let at = doc.paragraph_mark_position(&tail).unwrap().index;
        doc.apply_raw_ops("body", vec![author(at)], &ctx()).unwrap();
        let promoted = identity(&doc, &tail);
        assert_eq!(promoted.origin, ParagraphOrigin::Authored);
        assert_eq!(promoted.id_origin, Some(ParagraphIdOrigin::Authored));
        assert!(promoted.ooxml_para_id.is_some());
    }

    let doc = seeded(&bytes);
    let tail = doc
        .paragraphs("body")
        .unwrap()
        .last()
        .unwrap()
        .para_id
        .clone();
    let at = doc.paragraph_mark_position(&tail).unwrap().index;
    let typed = |text: &str| RawOp::Insert {
        index: at,
        text: text.into(),
        attrs: Default::default(),
    };
    doc.apply_raw_ops(
        "body",
        vec![
            typed("typed"),
            RawOp::Delete {
                index: at + 1,
                len: 4,
            },
            RawOp::Delete { index: at, len: 1 },
        ],
        &ctx(),
    )
    .unwrap();
    let transient = identity(&doc, &tail);
    assert_eq!(
        (transient.origin, transient.ooxml_para_id),
        (ParagraphOrigin::Synthetic, None)
    );
    doc.apply_raw_ops(
        "body",
        vec![typed("typed"), RawOp::Delete { index: at, len: 4 }],
        &ctx(),
    )
    .unwrap();
    assert_eq!(identity(&doc, &tail).origin, ParagraphOrigin::Authored);
}

/// Writes the session's save plan over a model parsed from the source, as a
/// save of stories that project exactly as seeded does.
fn save_unchanged(doc: &EditingDoc, bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let plan = doc.paragraph_save_plan();
    let package = docx_parse::parse_docx_s9_wire(
        bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            ..docx_parse::S9ParseOptions::default()
        },
    )
    .unwrap()
    .document
    .package;
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
        "paragraphIds": {
            "assignments": plan.assignments.iter().map(|(part, ordinal, para_id)| {
                json!({"part": part, "ordinal": ordinal, "paraId": para_id})
            }).collect::<Vec<_>>(),
            "patchedParts": plan.patched_parts.iter().map(|(part, para_ids)| {
                json!({"part": part, "paraIds": para_ids})
            }).collect::<Vec<_>>(),
        },
    }))
    .unwrap();
    ooxml_opc::unzip_parts(&write_docx_s13(request, bytes).unwrap()).unwrap()
}

#[test]
fn an_identity_only_save_patches_start_tags_and_keeps_every_other_byte() {
    let bytes = fixture();
    let doc = seeded(&bytes);
    assert!(
        doc.paragraph_save_plan()
            .patched_parts
            .iter()
            .all(|(_, patches)| patches.is_empty())
    );
    doc.persist_paragraph_ids().unwrap();
    let id = |key: &str| identity(&doc, key).ooxml_para_id.unwrap();
    let separators: Vec<String> = doc
        .paragraph_identities()
        .paragraphs
        .into_iter()
        .filter(|paragraph| {
            matches!(&paragraph.paragraph, ParagraphRef::Source(source) if source.part_uri == "/word/footnotes.xml")
        })
        .map(|paragraph| paragraph.ooxml_para_id.unwrap())
        .collect();

    let original = fixture_parts();
    let mut expected = original.clone();
    replace(
        &mut expected,
        "word/document.xml",
        "<w:p><w:r><w:t>Missing",
        &format!(r#"<w:p w14:paraId="{}"><w:r><w:t>Missing"#, id("body:p1")),
    );
    replace(
        &mut expected,
        "word/document.xml",
        r#"w14:paraId="1a2b3c4d""#,
        &format!(r#"w14:paraId="{}""#, id("body:p3")),
    );
    replace(
        &mut expected,
        "word/document.xml",
        r#"w14:paraId="xyz""#,
        &format!(r#"w14:paraId="{}""#, id("body:p4")),
    );
    replace(
        &mut expected,
        "word/footnotes.xml",
        "<w:p><w:r><w:separator/>",
        &format!(r#"<w:p w14:paraId="{}"><w:r><w:separator/>"#, separators[0]),
    );
    replace(
        &mut expected,
        "word/footnotes.xml",
        "<w:p><w:r><w:continuationSeparator/>",
        &format!(
            r#"<w:p w14:paraId="{}"><w:r><w:continuationSeparator/>"#,
            separators[1]
        ),
    );
    let saved = save_unchanged(&doc, &bytes);
    assert_eq!(
        saved
            .iter()
            .map(|(path, _)| path.as_str())
            .collect::<HashSet<_>>(),
        expected
            .iter()
            .map(|(path, _)| path.as_str())
            .collect::<HashSet<_>>()
    );
    for (name, bytes) in &expected {
        let written = saved
            .iter()
            .find(|(path, _)| path == name)
            .map(|(_, bytes)| bytes)
            .unwrap_or_else(|| panic!("{name} missing"));
        assert_eq!(
            String::from_utf8_lossy(written),
            String::from_utf8_lossy(bytes),
            "{name}"
        );
    }
    assert_ne!(
        text(&expected, "word/document.xml"),
        text(&original, "word/document.xml")
    );

    doc.insert_text(
        &ctx(),
        Position::new("body", 0),
        "edited",
        FormatPolicy::Plain,
    )
    .unwrap();
    let plan = doc.paragraph_save_plan();
    let patched: Vec<_> = plan
        .patched_parts
        .iter()
        .map(|(part, _)| part.as_str())
        .collect();
    assert!(!patched.contains(&"word/document.xml"));
    assert!(patched.contains(&"word/footnotes.xml") && patched.contains(&"word/header1.xml"));
    assert_eq!(
        plan.assignments
            .iter()
            .map(|(part, ordinal, _)| (part.as_str(), *ordinal))
            .collect::<Vec<_>>(),
        [("word/footnotes.xml", 0), ("word/footnotes.xml", 1)]
    );
}

#[test]
fn a_repaired_comment_paragraph_renames_its_companion_references() {
    let bytes = fixture_with(|parts| {
        replace(parts, "word/comments.xml", "4D5E6F7A", "0000ABCD");
        replace(parts, "word/commentsExtended.xml", "4D5E6F7A", "0000ABCD");
    });
    let doc = seeded(&bytes);
    let report = doc.persist_paragraph_ids().unwrap();
    let comment = report
        .assignments
        .iter()
        .find(|assignment| assignment.paragraph == source_ref(&doc, "/word/comments.xml", 0))
        .unwrap();
    assert_eq!(comment.origin, ParagraphIdOrigin::Repaired);
    assert_eq!(comment.previous_ooxml_para_id.as_deref(), Some("0000ABCD"));
    assert_eq!(
        identity(&doc, "0000abcd").ooxml_para_id.as_deref(),
        Some("0000abcd")
    );

    let saved = save_unchanged(&doc, &bytes);
    let original = ooxml_opc::unzip_parts(&bytes).unwrap();
    for name in ["word/comments.xml", "word/commentsExtended.xml"] {
        assert_eq!(
            text(&saved, name),
            text(&original, name).replace("0000ABCD", &comment.ooxml_para_id),
            "{name}"
        );
    }
}

fn paragraphs_package(body: &str) -> Vec<u8> {
    fixture_with(|parts| {
        let (_, document) = parts
            .iter_mut()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap();
        let xml = String::from_utf8(document.clone()).unwrap();
        let start = xml.find("<w:body>").unwrap() + "<w:body>".len();
        let end = xml.find("<w:sectPr>").unwrap();
        *document = format!("{}{body}{}", &xml[..start], &xml[end..]).into_bytes();
    })
}

#[test]
fn a_fresh_opening_never_resolves_a_stale_session_anchor() {
    let paragraph = |text: &str| format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
    let first = paragraphs_package(&[paragraph("A"), paragraph("B"), paragraph("C")].concat());
    let opened = EditingDoc::new(7);
    seed_from_docx(&opened, &first).unwrap();
    let stale = ParagraphAnchor::Session {
        session_id: opened.session_id(),
        story: "body".into(),
        para_id: "body:p2".into(),
    };
    assert_eq!(
        opened.resolve_paragraph_anchor(&stale),
        found("body", "body:p2")
    );

    let saved = paragraphs_package(
        &[
            paragraph("A"),
            r#"<w:p w14:paraId="1234ABCD"><w:r><w:t>X</w:t></w:r></w:p>"#.to_owned(),
            paragraph("B"),
            paragraph("C"),
        ]
        .concat(),
    );
    let reopened = EditingDoc::new(7);
    seed_from_docx(&reopened, &saved).unwrap();
    assert_eq!(reopened.paragraphs("body").unwrap()[2].para_id, "body:p2");
    assert_eq!(
        reopened.resolve_paragraph_anchor(&stale),
        AnchorResolution::Unsupported(AnchorUnsupported::ForeignSession)
    );
    let replica = EditingDoc::new(8);
    replica
        .apply_update_v1(&reopened.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(replica.session_id(), reopened.session_id());

    let fixed = |client| {
        let doc = EditingDoc::new(client);
        seed_from_docx_with_generation(&doc, &first, "shared").unwrap();
        doc.session_id()
    };
    assert_eq!(
        fixed(7),
        fixed(7),
        "a fixed generation seeds deterministically"
    );
}

#[test]
fn a_persisted_id_under_a_rebound_prefix_resolves_after_reopening() {
    let bytes = fixture_with(|parts| {
        replace(
            parts,
            "word/document.xml",
            "<w:p><w:r><w:t>Missing",
            r#"<w:p xmlns:w14="urn:other"><w:r><w:t>Missing"#,
        );
    });
    let doc = seeded(&bytes);
    doc.persist_paragraph_ids().unwrap();
    let anchor = persisted(body(), &identity(&doc, "body:p1").ooxml_para_id.unwrap());
    let saved = ooxml_opc::rezip_parts(&save_unchanged(&doc, &bytes)).unwrap();
    assert!(text(&ooxml_opc::unzip_parts(&saved).unwrap(), "word/document.xml").contains(
        r#"<w:p xmlns:w14p0="http://schemas.microsoft.com/office/word/2010/wordml" w14p0:paraId="#
    ));
    let reopened = seeded(&saved);
    let AnchorResolution::Found(ParagraphRef::Session { para_id, .. }) =
        reopened.resolve_paragraph_anchor(&anchor)
    else {
        panic!("the persisted anchor did not resolve");
    };
    let text = reopened
        .paragraphs("body")
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.para_id == para_id)
        .unwrap()
        .text;
    assert!(text.starts_with("Missing"));
}

#[test]
fn a_deleted_source_identity_stays_reserved_against_copies() {
    for key in ["copy", "0000abcd"] {
        let doc = seeded(&fixture());
        let lower = persisted(body(), "0000abcd");
        let stale = ParagraphAnchor::Session {
            session_id: doc.session_id(),
            story: "body".into(),
            para_id: "0000abcd".into(),
        };
        assert_eq!(
            doc.resolve_paragraph_anchor(&stale),
            found("body", "0000abcd")
        );
        doc.merge_paragraphs(&ctx(), "body:p1", MergeDirection::Forward)
            .unwrap();
        assert_eq!(
            doc.resolve_paragraph_anchor(&lower),
            AnchorResolution::Missing
        );
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "pilcrow".into(),
                payload: vec![
                    ("paraId".into(), Any::from(key)),
                    ("ooxmlParaId".into(), Any::from("0000abcd")),
                ],
                attrs: Default::default(),
            }],
            &ctx(),
        )
        .unwrap();
        let copy = doc.paragraph_identities().paragraphs.remove(0);
        let ParagraphRef::Session { para_id, .. } = &copy.paragraph else {
            panic!("{copy:?}");
        };
        assert_ne!(para_id, "0000abcd", "a copy never takes a deleted key");
        assert_eq!(copy.id_origin, Some(ParagraphIdOrigin::Repaired));
        assert_ne!(copy.ooxml_para_id.as_deref(), Some("0000abcd"));
        assert_eq!(
            doc.resolve_paragraph_anchor(&lower),
            AnchorResolution::Missing
        );
        assert_eq!(
            doc.resolve_paragraph_anchor(&stale),
            AnchorResolution::Missing
        );
    }
}

#[test]
fn a_raw_identity_stays_reserved_after_its_paragraph_is_deleted() {
    for save in [false, true] {
        a_raw_identity_stays_reserved(save);
    }
}

fn a_raw_identity_stays_reserved(save: bool) {
    let doc = seeded(&fixture());
    let insert = || {
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "pilcrow".into(),
                payload: vec![
                    ("paraId".into(), Any::from("raw")),
                    ("ooxmlParaId".into(), Any::from("12345678")),
                ],
                attrs: Default::default(),
            }],
            &ctx(),
        )
        .unwrap();
    };
    insert();
    let raw = identity(&doc, "raw");
    assert_eq!(raw.ooxml_para_id.as_deref(), Some("12345678"));
    assert_eq!(raw.id_origin, Some(ParagraphIdOrigin::Authored));
    let stale = ParagraphAnchor::Session {
        session_id: doc.session_id(),
        story: "body".into(),
        para_id: "raw".into(),
    };
    let saved = persisted(body(), "12345678");
    assert_eq!(doc.resolve_paragraph_anchor(&stale), found("body", "raw"));
    assert_eq!(doc.resolve_paragraph_anchor(&saved), found("body", "raw"));
    if save {
        assert!(
            doc.record_saved_paragraph_ids(&[("raw".into(), "12345678".into())])
                .is_empty()
        );
    }

    doc.apply_raw_ops("body", vec![RawOp::Delete { index: 0, len: 1 }], &ctx())
        .unwrap();
    insert();
    let replacement = doc.paragraph_identities().paragraphs.remove(0);
    let ParagraphRef::Session { para_id, .. } = &replacement.paragraph else {
        panic!("{replacement:?}");
    };
    assert_ne!(para_id, "raw");
    assert_ne!(replacement.ooxml_para_id.as_deref(), Some("12345678"));
    assert_eq!(
        doc.resolve_paragraph_anchor(&stale),
        AnchorResolution::Missing
    );
    assert_eq!(
        doc.resolve_paragraph_anchor(&saved),
        AnchorResolution::Missing
    );
}

#[test]
fn any_change_the_save_projects_rewrites_its_part() {
    let patched = |doc: &EditingDoc| {
        doc.paragraph_save_plan()
            .patched_parts
            .iter()
            .any(|(part, _)| part == "word/document.xml")
    };
    let mut note = None;
    let mut unit = 0;
    for segment in seeded(&fixture()).story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => unit += text.encode_utf16().count() as u32,
            SegmentContent::OtherEmbed { kind, .. } if kind == "noteRef" && note.is_none() => {
                note = Some(unit);
                unit += 1;
            }
            _ => unit += 1,
        }
    }
    let note = note.expect("the fixture references a footnote");
    let comment = |start: u32, end: u32| {
        move |doc: &EditingDoc| {
            doc.add_comment(
                &[StoryRange::new("body", start, end)],
                "Ada",
                DATE,
                Any::Null,
            )
            .map(drop)
            .unwrap();
        }
    };
    type Edit = Box<dyn Fn(&EditingDoc)>;
    let edits: [(&str, Edit); 3] = [
        ("a comment over an embed", Box::new(comment(note, note + 1))),
        ("a comment at a caret", Box::new(comment(2, 2))),
        (
            "a suggested deletion of an embed",
            Box::new(move |doc: &EditingDoc| {
                doc.delete_range(&ctx().suggesting(), StoryRange::new("body", note, note + 1))
                    .unwrap();
            }),
        ),
    ];
    for (edit, apply) in edits {
        let doc = seeded(&fixture());
        assert!(patched(&doc));
        apply(&doc);
        assert!(!patched(&doc), "{edit} rewrites the body part");
    }
}

/// The fixture with a second relationship to its header part, which holds
/// a paragraph with a source ID and one without.
fn shared_header() -> Vec<u8> {
    fixture_with(|parts| {
        replace(
            parts,
            "word/_rels/document.xml.rels",
            "<Relationship Id=\"rIdFootnotes\"",
            "<Relationship Id=\"rIdHeaderFirst\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/><Relationship Id=\"rIdFootnotes\"",
        );
        replace(
            parts,
            "word/document.xml",
            "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>",
            "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/><w:headerReference w:type=\"first\" r:id=\"rIdHeaderFirst\"/>",
        );
        replace(
            parts,
            "word/header1.xml",
            "</w:p></w:hdr>",
            "</w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:hdr>",
        );
    })
}

#[test]
fn stories_sharing_a_part_are_views_of_one_paragraph_each() {
    const HEADER: &str = "/word/header1.xml";
    let header = || story(HEADER, SourceStoryKind::Header, None);
    let bytes = shared_header();
    let doc = seeded(&bytes);
    let views = |doc: &EditingDoc| -> Vec<(String, Option<String>, Option<SourceParagraphRef>)> {
        doc.paragraph_identities()
            .paragraphs
            .into_iter()
            .filter_map(|identity| match identity.paragraph {
                ParagraphRef::Session { story, para_id } if story.starts_with("hf:") => {
                    Some((para_id, identity.ooxml_para_id, identity.source))
                }
                _ => None,
            })
            .collect()
    };
    let occurrence = |ordinal| match source_ref(&doc, HEADER, ordinal) {
        ParagraphRef::Source(reference) => Some(reference),
        other => panic!("{other:?}"),
    };
    let seeded_views = views(&doc);
    assert_eq!(seeded_views.len(), 4);
    for (_, id, source) in &seeded_views {
        assert_eq!(
            source.as_ref(),
            occurrence(if id.is_some() { 0 } else { 1 }).as_ref()
        );
    }

    let report = doc.persist_paragraph_ids().unwrap();
    assert!(report.diagnostics.is_empty());
    let header_assignments: Vec<&ParagraphIdAssignment> = report
        .assignments
        .iter()
        .filter(|assignment| assignment.source_story == Some(header()))
        .collect();
    assert_eq!(header_assignments.len(), 2, "{header_assignments:?}");
    assert!(
        header_assignments
            .iter()
            .all(|assignment| assignment.origin == ParagraphIdOrigin::Persisted)
    );
    let plain = header_assignments[0].ooxml_para_id.clone();
    assert_eq!(header_assignments[1].ooxml_para_id, plain);
    let persisted_views = views(&doc);
    for (_, id, source) in &persisted_views {
        let expected = if source == &occurrence(0) {
            "6F7A8B9C"
        } else {
            plain.as_str()
        };
        assert_eq!(id.as_deref(), Some(expected));
    }
    for (para_id, ordinal) in [("6F7A8B9C", 0), (plain.as_str(), 1)] {
        let AnchorResolution::Found(first) =
            doc.resolve_paragraph_anchor(&persisted(header(), para_id))
        else {
            panic!("{para_id} does not resolve to one paragraph");
        };
        assert_eq!(
            doc.resolve_paragraph_anchor(&ParagraphAnchor::Source(occurrence(ordinal).unwrap())),
            AnchorResolution::Found(first)
        );
    }
    assert!(doc.persist_paragraph_ids().unwrap().assignments.is_empty());

    let written = save_unchanged(&doc, &bytes);
    let saved: Vec<(String, String)> = persisted_views
        .into_iter()
        .map(|(key, id, _)| (key, id.unwrap()))
        .collect();
    assert!(doc.record_saved_paragraph_ids(&saved).is_empty());
    let report = doc.persist_paragraph_ids().unwrap();
    assert!(report.assignments.is_empty() && report.diagnostics.is_empty());
    let written = ooxml_opc::rezip_parts(&written).unwrap();
    assert_eq!(
        paragraph_ids_by_part(&written).unwrap()[HEADER],
        ["6F7A8B9C", plain.as_str()]
    );
    let reopened = seeded(&written);
    for para_id in ["6F7A8B9C", plain.as_str()] {
        assert!(matches!(
            reopened.resolve_paragraph_anchor(&persisted(header(), para_id)),
            AnchorResolution::Found(_)
        ));
    }
}
