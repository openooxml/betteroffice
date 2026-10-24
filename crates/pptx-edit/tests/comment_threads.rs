use std::collections::BTreeMap;

use pptx_edit::{CommentFlavor, DeckSession, EditCtx};

fn fixture() -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#),
        ("ppt/presentation.xml", r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#),
        ("ppt/_rels/presentation.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#),
        ("ppt/slides/slide1.xml", r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#),
    ].into_iter().map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec())).collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn source(flavor: CommentFlavor) -> Vec<u8> {
    let session = DeckSession::open(&fixture(), 501).unwrap();
    let ctx = EditCtx::local("test");
    session.set_comment_flavor(&ctx, flavor).unwrap();
    let slide = session.slide_ids().unwrap()[0].clone();
    session
        .add_comment(
            &ctx,
            &slide,
            "Reviewer",
            "RV",
            "Review this slide",
            "2026-01-01T10:00:00Z",
            12700,
            25400,
        )
        .unwrap();
    session.save().unwrap()
}

#[test]
fn legacy_reply_writes_threading_authors_and_preserves_source_bytes() {
    let source = source(CommentFlavor::Legacy);
    let original = parts(&source);
    let session = DeckSession::open(&source, 502).unwrap();
    let root = session.comments().unwrap().remove(0);
    let ctx = EditCtx::local("test");
    session
        .reply_to_comment(
            &ctx,
            &root.id,
            "Writer",
            "WR",
            "Addressed & checked",
            "2026-01-01T11:00:00Z",
        )
        .unwrap();
    session
        .reply_to_comment(
            &ctx,
            &root.id,
            "Writer",
            "WR",
            "Second reply",
            "2026-01-01T12:00:00Z",
        )
        .unwrap();
    let saved = session.save().unwrap();
    let output = parts(&saved);
    let xml = String::from_utf8(output["ppt/comments/comment1.xml"].clone()).unwrap();
    assert!(xml.contains("{C676402C-5697-4E1C-873F-D02D1690AC5C}"));
    assert!(xml.contains("<p15:threadingInfo"));
    assert!(xml.contains("<p15:parentCm authorId=\"0\" idx=\"1\""));
    assert!(xml.contains("idx=\"2\""));
    let authors = String::from_utf8(output["ppt/commentAuthors.xml"].clone()).unwrap();
    for attr in [
        "id=\"1\"",
        "name=\"Writer\"",
        "initials=\"WR\"",
        "lastIdx=\"2\"",
        "clrIdx=\"1\"",
    ] {
        assert!(authors.contains(attr), "{authors}");
    }
    let old_xml = String::from_utf8(original["ppt/comments/comment1.xml"].clone()).unwrap();
    assert!(xml.starts_with(old_xml.strip_suffix("</p:cmLst>").unwrap()));
    let root_position = &old_xml[old_xml.find("<p:pos ").unwrap()..];
    let root_position = &root_position[..root_position.find("/>").unwrap() + 2];
    assert_eq!(xml.matches(root_position).count(), 3, "{xml}");
    for (path, bytes) in &original {
        if path != "ppt/comments/comment1.xml" && path != "ppt/commentAuthors.xml" {
            assert_eq!(&output[path], bytes, "{path}");
        }
    }
    let reopened = DeckSession::open(&saved, 503).unwrap();
    let comments = reopened.comments().unwrap();
    assert_eq!(comments.len(), 3);
    assert_eq!(comments.iter().filter(|c| c.parent_id.is_some()).count(), 2);
    assert_eq!(parts(&reopened.save().unwrap()), output);
    assert!(session.undo());
    assert!(session.undo());
    assert_eq!(parts(&session.save().unwrap()), original);
}

#[test]
fn duplicate_discussions_are_independent_and_undoable_in_both_flavors() {
    for flavor in [CommentFlavor::Legacy, CommentFlavor::Modern] {
        let source = source(flavor);
        let seed = DeckSession::open(&source, 504).unwrap();
        let root = seed.comments().unwrap().remove(0);
        let ctx = EditCtx::local("test");
        seed.reply_to_comment(
            &ctx,
            &root.id,
            "Writer",
            "WR",
            "Reply",
            "2026-01-01T11:00:00Z",
        )
        .unwrap();
        if flavor == CommentFlavor::Modern {
            seed.set_comment_status(&ctx, &root.id, true).unwrap();
        }
        let source = seed.save().unwrap();
        let session = DeckSession::open(&source, 505).unwrap();
        let slide = session.slide_ids().unwrap()[0].clone();
        let copy = session.duplicate_slide(&ctx, &slide, 1).unwrap();
        let comments = session.comments().unwrap();
        assert_eq!(comments.len(), 4);
        let copied = comments
            .iter()
            .filter(|c| c.slide_id == copy.slide_id)
            .collect::<Vec<_>>();
        assert_eq!(copied.len(), 2);
        let copied_root = copied.iter().find(|c| c.parent_id.is_none()).unwrap();
        assert_eq!(copied_root.resolved, flavor == CommentFlavor::Modern);
        assert!(
            copied
                .iter()
                .any(|c| c.parent_id.as_deref() == Some(copied_root.id.as_str()))
        );
        let saved = session.save().unwrap();
        let reopened = DeckSession::open(&saved, 506).unwrap();
        let restored = reopened.comments().unwrap();
        assert_eq!(restored.len(), 4);
        for reply in restored.iter().filter(|c| c.parent_id.is_some()) {
            let parent = restored
                .iter()
                .find(|c| Some(&c.id) == reply.parent_id.as_ref())
                .unwrap();
            assert_eq!(reply.slide_id, parent.slide_id);
        }
        assert_eq!(
            parts(&saved)
                .keys()
                .filter(|path| path.starts_with("ppt/comments/"))
                .count(),
            2
        );
        assert_eq!(parts(&reopened.save().unwrap()), parts(&saved));
        let saved_parts = parts(&saved);
        let content_types = String::from_utf8(saved_parts["[Content_Types].xml"].clone()).unwrap();
        for path in saved_parts
            .keys()
            .filter(|path| path.starts_with("ppt/comments/"))
        {
            assert!(
                content_types.contains(&format!("PartName=\"/{path}\"")),
                "{content_types}"
            );
            let target = format!("comments/{}\"", path.trim_start_matches("ppt/comments/"));
            let referencing = saved_parts
                .iter()
                .filter(|(rels, bytes)| {
                    rels.starts_with("ppt/slides/_rels/")
                        && String::from_utf8_lossy(bytes).contains(&target)
                })
                .count();
            assert_eq!(referencing, 1, "{path}");
        }
        for (path, bytes) in parts(&source) {
            if path.starts_with("ppt/comments/") {
                assert_eq!(parts(&saved)[&path], bytes, "{path}");
            }
        }
        assert!(session.undo());
        assert_eq!(parts(&session.save().unwrap()), parts(&source));
        assert!(session.redo());
        assert_eq!(parts(&session.save().unwrap()), parts(&saved));
        session.remove_comment(&ctx, &copied_root.id).unwrap();
        assert_eq!(session.comments().unwrap().len(), 2);
        assert_eq!(
            DeckSession::open(&session.save().unwrap(), 507)
                .unwrap()
                .comments()
                .unwrap()
                .len(),
            2
        );
    }
}

#[test]
fn legacy_parent_references_are_scoped_to_their_slide() {
    let seed = DeckSession::open(&source(CommentFlavor::Legacy), 508).unwrap();
    let ctx = EditCtx::local("test");
    let root = seed.comments().unwrap().remove(0);
    seed.reply_to_comment(
        &ctx,
        &root.id,
        "Writer",
        "WR",
        "Reply",
        "2026-01-01T11:00:00Z",
    )
    .unwrap();
    seed.duplicate_slide(&ctx, &root.slide_id, 1).unwrap();
    let original = parts(&seed.save().unwrap())
        .into_iter()
        .map(|(path, bytes)| {
            let bytes = if path.starts_with("ppt/comments/") {
                String::from_utf8(bytes)
                    .unwrap()
                    .replace("idx=\"2\"", "idx=\"1\"")
                    .into_bytes()
            } else {
                bytes
            };
            (path, bytes)
        })
        .collect::<Vec<_>>();
    let source = ooxml_opc::rezip_parts(&original).unwrap();
    let session = DeckSession::open(&source, 509).unwrap();
    assert_eq!(parts(&session.save().unwrap()), parts(&source));
    let check_parents = || {
        let comments = session.comments().unwrap();
        for reply in comments
            .iter()
            .filter(|comment| comment.parent_id.is_some())
        {
            let parent = comments
                .iter()
                .find(|comment| Some(&comment.id) == reply.parent_id.as_ref())
                .unwrap();
            assert_eq!(reply.slide_id, parent.slide_id);
        }
    };
    check_parents();
    let second_slide = session.slide_ids().unwrap()[1].clone();
    session.duplicate_slide(&ctx, &second_slide, 2).unwrap();
    assert_eq!(session.comments().unwrap().len(), 6);
    assert_eq!(
        session
            .comments()
            .unwrap()
            .iter()
            .filter(|comment| comment.parent_id.is_some())
            .count(),
        3
    );
    check_parents();
    let reopened = DeckSession::open(&session.save().unwrap(), 510).unwrap();
    assert_eq!(
        reopened
            .comments()
            .unwrap()
            .iter()
            .filter(|comment| comment.parent_id.is_some())
            .count(),
        3
    );
    assert!(session.undo());
    assert_eq!(parts(&session.save().unwrap()), parts(&source));
}

#[test]
fn concurrent_legacy_replies_by_new_authors_save_identically_on_both_peers() {
    let source = source(CommentFlavor::Legacy);
    let left = DeckSession::open(&source, 511).unwrap();
    let right = DeckSession::open(&source, 512).unwrap();
    let root = left.comments().unwrap().remove(0);
    left.reply_to_comment(
        &EditCtx::local("left"),
        &root.id,
        "Left",
        "LE",
        "Left reply",
        "2026-01-01T11:00:00Z",
    )
    .unwrap();
    right
        .reply_to_comment(
            &EditCtx::local("right"),
            &root.id,
            "Right",
            "RI",
            "Right reply",
            "2026-01-01T11:00:00Z",
        )
        .unwrap();
    let left_update = left.encode_state_as_update_v1();
    let right_update = right.encode_state_as_update_v1();
    left.apply_update_v1(&right_update).unwrap();
    right.apply_update_v1(&left_update).unwrap();
    let saved = left.save().unwrap();
    assert_eq!(parts(&saved), parts(&right.save().unwrap()));
    let output = parts(&saved);
    let authors = String::from_utf8(output["ppt/commentAuthors.xml"].clone()).unwrap();
    for name in ["name=\"Left\"", "name=\"Right\"", "id=\"1\"", "id=\"2\""] {
        assert!(authors.contains(name), "{authors}");
    }
    let reopened = DeckSession::open(&saved, 513).unwrap();
    let comments = reopened.comments().unwrap();
    assert_eq!(comments.len(), 3);
    let root = comments
        .iter()
        .find(|comment| comment.parent_id.is_none())
        .unwrap();
    assert_eq!(
        comments
            .iter()
            .filter(|comment| comment.parent_id.as_deref() == Some(root.id.as_str()))
            .count(),
        2
    );
}
