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
