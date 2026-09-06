use std::sync::{Arc, Mutex};

use pptx_edit::{CommentFlavor, DeckSession, DeckSnapshot, EditCtx, EditError, TextStyle};
use yrs::updates::decoder::Decode;

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const MODERN: &[u8] = include_bytes!("fixtures/modern-comments.pptx");
const FIRST_PART: &str = "ppt/comments/modernComment_256_499602D2.xml";

fn parts(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn first_root(session: &DeckSession) -> pptx_edit::CommentSnapshot {
    session
        .comments()
        .unwrap()
        .into_iter()
        .find(|comment| comment.created.as_deref() == Some("2024-12-30T20:26:06.503Z"))
        .unwrap()
}

#[test]
fn resolving_a_source_comment_changes_only_its_status_bytes() {
    let session = DeckSession::open(MODERN, 721).unwrap();
    let original = parts(MODERN);
    assert_eq!(parts(&session.save().unwrap()), original);
    let root = first_root(&session);
    session
        .set_comment_status(&EditCtx::local("test"), &root.id, true)
        .unwrap();
    let mut expected = original.clone();
    expected.insert(
        FIRST_PART.to_owned(),
        String::from_utf8(original[FIRST_PART].clone())
            .unwrap()
            .replacen("status=\"active\"", "status=\"resolved\"", 1)
            .into_bytes(),
    );
    assert_eq!(parts(&session.save().unwrap()), expected);
    assert!(session.undo());
    assert_eq!(parts(&session.save().unwrap()), original);
    assert!(session.redo());
    let update = session.encode_state_as_update_v1();
    let reopened = DeckSession::open_from_update_with_source(&update, MODERN, 722).unwrap();
    assert_eq!(parts(&reopened.save().unwrap()), expected);
}

#[test]
fn adding_and_removing_a_reply_preserves_source_xml_and_authors() {
    let session = DeckSession::open(MODERN, 723).unwrap();
    let context = EditCtx::local("test");
    let root = first_root(&session);
    let reply = session
        .reply_to_comment(
            &context,
            &root.id,
            &root.author,
            &root.initials,
            "New reply 😀",
            "2026-09-05T12:00:00Z",
        )
        .unwrap();
    let saved = parts(&session.save().unwrap());
    let original = parts(MODERN);
    for (path, bytes) in &original {
        if path != FIRST_PART {
            assert_eq!(&saved[path], bytes, "{path}");
        }
    }
    let xml = String::from_utf8(saved[FIRST_PART].clone()).unwrap();
    let insertion_end = xml.find("</p188:replyLst>").unwrap();
    let insertion_start = xml[..insertion_end].rfind("<p188:reply ").unwrap();
    let restored = format!("{}{}", &xml[..insertion_start], &xml[insertion_end..]);
    assert_eq!(restored.as_bytes(), original[FIRST_PART]);
    let reopened = DeckSession::open(&session.save().unwrap(), 724).unwrap();
    assert_eq!(reopened.comments().unwrap().len(), 6);
    session.remove_comment(&context, &reply.comment_id).unwrap();
    assert_eq!(parts(&session.save().unwrap()), original);
}

#[test]
fn concurrent_root_deletion_promotes_replies_in_both_update_orders() {
    let left = DeckSession::open(MODERN, 726).unwrap();
    let right = DeckSession::open(MODERN, 727).unwrap();
    let root = first_root(&left);
    let context = EditCtx::local("test");
    left.remove_comment(&context, &root.id).unwrap();
    let reply = right
        .reply_to_comment(
            &context,
            &root.id,
            "Concurrent",
            "C",
            "Keep this reply 😀",
            "2026-09-05T12:00:00Z",
        )
        .unwrap();
    let deletion = left.encode_state_as_update_v1();
    let addition = right.encode_state_as_update_v1();
    left.apply_update_v1(&addition).unwrap();
    right.apply_update_v1(&deletion).unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let expected = left.snapshot().unwrap();
    assert_eq!(expected.comments.len(), 3);
    let promoted = expected
        .comments
        .iter()
        .find(|comment| comment.id == reply.comment_id)
        .unwrap();
    assert_eq!(promoted.parent_id, None);
    assert_eq!(promoted.text, "Keep this reply 😀");
    let saved = left.save().unwrap();
    assert_eq!(parts(&right.save().unwrap()), parts(&saved));
    for updates in [[&deletion, &addition], [&addition, &deletion]] {
        let receiver = DeckSession::open(MODERN, 728).unwrap();
        for update in updates {
            receiver.apply_update_v1(update).unwrap();
        }
        assert_eq!(receiver.snapshot().unwrap(), expected);
        assert_eq!(parts(&receiver.save().unwrap()), parts(&saved));
    }
    let reopened = DeckSession::open(&saved, 729).unwrap();
    let comments = reopened.comments().unwrap();
    assert_eq!(comments.len(), 3);
    assert!(
        comments
            .iter()
            .any(|comment| comment.text == promoted.text && comment.parent_id.is_none())
    );
    let followup = left
        .reply_to_comment(
            &context,
            &reply.comment_id,
            "Concurrent",
            "C",
            "Follow-up",
            "2026-09-05T12:01:00Z",
        )
        .unwrap();
    assert_eq!(
        followup.parent_id.as_deref(),
        Some(reply.comment_id.as_str())
    );
    assert_eq!(
        DeckSession::open(&left.save().unwrap(), 730)
            .unwrap()
            .comments()
            .unwrap()
            .len(),
        4
    );
}

#[test]
fn v2_source_migration_imports_comments_once_and_keeps_peer_edits() {
    let update = include_bytes!("fixtures/deck-schema-v2-comments.update.bin");
    let source_free = DeckSession::open_from_update(update, 731).unwrap();
    let left = DeckSession::open_from_update_with_source(
        &source_free.encode_state_as_update_v1(),
        MODERN,
        732,
    )
    .unwrap();
    let right = DeckSession::open_from_update_with_source(update, MODERN, 733).unwrap();
    assert_eq!(left.comments().unwrap().len(), 5);
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(parts(&left.save().unwrap()), parts(MODERN));
    let root = first_root(&left);
    left.set_comment_status(&EditCtx::local("test"), &root.id, true)
        .unwrap();
    let reply = right
        .reply_to_comment(
            &EditCtx::local("test"),
            &root.id,
            "Peer",
            "P",
            "Migrated reply",
            "2026-09-05T12:02:00Z",
        )
        .unwrap();
    let left_update = left.encode_state_as_update_v1();
    let right_update = right.encode_state_as_update_v1();
    left.apply_update_v1(&right_update).unwrap();
    right.apply_update_v1(&left_update).unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    let comments = left.comments().unwrap();
    assert_eq!(comments.len(), 6);
    assert!(
        comments
            .iter()
            .find(|comment| comment.id == root.id)
            .unwrap()
            .resolved
    );
    assert!(
        comments
            .iter()
            .any(|comment| comment.id == reply.comment_id)
    );
    let migrated = left.encode_state_as_update_v1();
    let reopened = DeckSession::open_from_update_with_source(&migrated, MODERN, 734).unwrap();
    assert_eq!(
        yrs::StateVector::decode_v1(&reopened.encode_state_vector_v1()).unwrap(),
        yrs::StateVector::decode_v1(&left.encode_state_vector_v1()).unwrap()
    );
    assert_eq!(reopened.snapshot().unwrap(), left.snapshot().unwrap());
    assert_eq!(
        DeckSession::open(&reopened.save().unwrap(), 735)
            .unwrap()
            .comments()
            .unwrap()
            .len(),
        6
    );
    for comment in reopened
        .comments()
        .unwrap()
        .iter()
        .filter(|comment| comment.parent_id.is_none())
    {
        reopened
            .remove_comment(&EditCtx::local("test"), &comment.id)
            .unwrap();
    }
    let deleted = DeckSession::open_from_update_with_source(
        &reopened.encode_state_as_update_v1(),
        MODERN,
        736,
    )
    .unwrap();
    assert!(deleted.comments().unwrap().is_empty());
    assert_eq!(
        parts(&deleted.save().unwrap()).len(),
        parts(MODERN).len() - 3
    );
}

#[test]
fn v7_source_migration_imports_comments_once_and_saves_every_part_unchanged() {
    let update = include_bytes!("fixtures/deck-schema-v7-comments.update.bin");
    let session = DeckSession::open_from_update_with_source(update, MODERN, 741).unwrap();
    assert_eq!(session.comments().unwrap().len(), 5);
    assert_eq!(session.comment_flavor().unwrap(), CommentFlavor::Modern);
    let source = parts(MODERN);
    assert_eq!(parts(&session.save().unwrap()), source);
    let root = first_root(&session);
    session
        .set_comment_status(&EditCtx::local("test"), &root.id, true)
        .unwrap();
    let reopened = DeckSession::open_from_update_with_source(
        &session.encode_state_as_update_v1(),
        MODERN,
        742,
    )
    .unwrap();
    let comments = reopened.comments().unwrap();
    assert_eq!(comments.len(), 5);
    assert!(
        comments
            .iter()
            .find(|comment| comment.id == root.id)
            .unwrap()
            .resolved
    );
    let saved = parts(&reopened.save().unwrap());
    assert_eq!(
        saved.keys().collect::<Vec<_>>(),
        source.keys().collect::<Vec<_>>()
    );
    assert_eq!(
        saved
            .iter()
            .filter(|(name, bytes)| &source[*name] != *bytes)
            .map(|(name, _)| name.as_str())
            .collect::<Vec<_>>(),
        [FIRST_PART]
    );
}

#[test]
fn source_extensions_and_custom_prefixes_survive_status_and_new_authors() {
    let mut original = parts(MODERN);
    let xml = String::from_utf8(original[FIRST_PART].clone()).unwrap()
        .replace("p188:", "modern:").replace("xmlns:p188", "xmlns:modern")
        .replace("status=\"active\"", "status = 'active' xmlns:ext='urn:test' ext:opaque='a &amp; b'")
        .replace("<modern:pos", "<!--opaque--><?keep value?><ext:opaque xmlns:ext='urn:test'><![CDATA[a < b]]></ext:opaque><modern:pos");
    original.insert(FIRST_PART.to_owned(), xml.clone().into_bytes());
    let source = ooxml_opc::rezip_parts(&original.clone().into_iter().collect::<Vec<_>>()).unwrap();
    let session = DeckSession::open(&source, 737).unwrap();
    let root = first_root(&session);
    session
        .set_comment_status(&EditCtx::local("test"), &root.id, true)
        .unwrap();
    let saved = parts(&session.save().unwrap());
    let mut expected = original.clone();
    expected.insert(
        FIRST_PART.to_owned(),
        xml.replace("status = 'active'", "status = 'resolved'")
            .into_bytes(),
    );
    assert_eq!(saved, expected);
    session
        .reply_to_comment(
            &EditCtx::local("test"),
            &root.id,
            "New & author",
            "N",
            "New reply",
            "2026-09-05T13:00:00Z",
        )
        .unwrap();
    let saved = parts(&session.save().unwrap());
    for (path, bytes) in &expected {
        if path != FIRST_PART && path != "ppt/authors.xml" {
            assert_eq!(&saved[path], bytes, "{path}");
        }
    }
    let old_authors = String::from_utf8(original["ppt/authors.xml"].clone()).unwrap();
    let new_authors = String::from_utf8(saved["ppt/authors.xml"].clone()).unwrap();
    assert!(new_authors.starts_with(old_authors.strip_suffix("</p188:authorLst>").unwrap()));
    let reopened = DeckSession::open(&session.save().unwrap(), 738).unwrap();
    assert_eq!(reopened.comments().unwrap().len(), 6);
    assert!(
        reopened
            .comments()
            .unwrap()
            .iter()
            .any(|comment| comment.author == "New & author")
    );
}

#[test]
fn new_reply_lists_preserve_existing_comments_and_empty_element_forms() {
    for reply_list in ["", "<p188:replyLst/>", "<p188:replyLst></p188:replyLst>"] {
        let mut original = parts(MODERN);
        let second = "ppt/comments/modernComment_257_3ADE68B1.xml";
        original.insert(
            second.to_owned(),
            String::from_utf8(original[second].clone())
                .unwrap()
                .replace("<p188:txBody>", &format!("{reply_list}<p188:txBody>"))
                .into_bytes(),
        );
        let source =
            ooxml_opc::rezip_parts(&original.clone().into_iter().collect::<Vec<_>>()).unwrap();
        let session = DeckSession::open(&source, 739).unwrap();
        let root = session
            .comments()
            .unwrap()
            .into_iter()
            .find(|comment| comment.text == "Slide two comment.")
            .unwrap();
        let context = EditCtx::local("test");
        session
            .set_comment_status(&context, &root.id, true)
            .unwrap();
        session
            .reply_to_comment(
                &context,
                &root.id,
                &root.author,
                &root.initials,
                "Reply",
                "2026-09-05T13:01:00Z",
            )
            .unwrap();
        let saved = session.save().unwrap();
        for (path, bytes) in &original {
            if path != second {
                assert_eq!(&parts(&saved)[path], bytes, "{path}");
            }
        }
        let reopened = DeckSession::open(&saved, 740).unwrap();
        assert_eq!(reopened.comments().unwrap().len(), 6);
        assert!(
            reopened
                .comments()
                .unwrap()
                .iter()
                .any(|comment| comment.text == "Slide two comment." && comment.resolved)
        );
    }
}

#[test]
fn legacy_additions_preserve_source_authors_and_advance_their_index() {
    let seed = DeckSession::open(DEMO, 741).unwrap();
    let context = EditCtx::local("test");
    let slide = seed.snapshot().unwrap().slides[0].id.clone();
    seed.add_comment(
        &context,
        &slide,
        "Known author",
        "KA",
        "Legacy root",
        "2026-09-05T14:00:00Z",
        0,
        0,
    )
    .unwrap();
    let mut original = parts(&seed.save().unwrap());
    let authors = "ppt/commentAuthors.xml";
    original.insert(
        authors.to_owned(),
        String::from_utf8(original[authors].clone())
            .unwrap()
            .replace("id=\"0\"", "id=\"7\"")
            .replace("lastIdx=\"1\"", "lastIdx=\"9\" custom=\"preserve\"")
            .into_bytes(),
    );
    let comment_part = "ppt/comments/comment1.xml";
    original.insert(
        comment_part.to_owned(),
        String::from_utf8(original[comment_part].clone())
            .unwrap()
            .replace("authorId=\"0\"", "authorId=\"7\"")
            .into_bytes(),
    );
    let source = ooxml_opc::rezip_parts(&original.clone().into_iter().collect::<Vec<_>>()).unwrap();
    let session = DeckSession::open(&source, 742).unwrap();
    session
        .add_comment(
            &context,
            &slide,
            "Known author",
            "KA",
            "Next legacy",
            "2026-09-05T14:01:00Z",
            0,
            0,
        )
        .unwrap();
    let saved = parts(&session.save().unwrap());
    assert_eq!(
        String::from_utf8(saved[authors].clone()).unwrap(),
        String::from_utf8(original[authors].clone())
            .unwrap()
            .replace("lastIdx=\"9\"", "lastIdx=\"10\"")
    );
    let xml = String::from_utf8(saved[comment_part].clone()).unwrap();
    let old_xml = String::from_utf8(original[comment_part].clone()).unwrap();
    assert!(xml.starts_with(old_xml.strip_suffix("</p:cmLst>").unwrap()));
    assert!(xml.contains("idx=\"10\""));
    let reopened = DeckSession::open(&session.save().unwrap(), 743).unwrap();
    assert_eq!(reopened.comments().unwrap().len(), 2);
    assert!(
        reopened
            .comments()
            .unwrap()
            .iter()
            .all(|comment| comment.author == "Known author")
    );
}

#[test]
fn editing_a_comment_preserves_an_untouched_empty_comment_part() {
    let mut original = parts(MODERN);
    let second = "ppt/comments/modernComment_257_3ADE68B1.xml";
    original.insert(second.to_owned(), br#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><unknown:metadata xmlns:unknown="urn:test" value="keep"/></p188:cmLst>"#.to_vec());
    let source = ooxml_opc::rezip_parts(&original.clone().into_iter().collect::<Vec<_>>()).unwrap();
    let session = DeckSession::open(&source, 744).unwrap();
    let root = first_root(&session);
    session
        .set_comment_status(&EditCtx::local("test"), &root.id, true)
        .unwrap();
    let saved = parts(&session.save().unwrap());
    for (path, bytes) in &original {
        if path != FIRST_PART {
            assert_eq!(saved.get(path), Some(bytes), "{path}");
        }
    }
}

#[test]
fn undoing_a_root_deletion_preserves_replies_to_a_promoted_comment() {
    let left = DeckSession::open(MODERN, 745).unwrap();
    let right = DeckSession::open(MODERN, 746).unwrap();
    let root = first_root(&left);
    let context = EditCtx::local("test");
    left.remove_comment(&context, &root.id).unwrap();
    let reply = right
        .reply_to_comment(
            &context,
            &root.id,
            "Peer",
            "P",
            "Concurrent reply",
            "2026-09-05T15:00:00Z",
        )
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    let followup = right
        .reply_to_comment(
            &context,
            &reply.comment_id,
            "Peer",
            "P",
            "Follow-up",
            "2026-09-05T15:01:00Z",
        )
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();
    assert!(left.undo());
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(left.comments().unwrap().len(), 7);
    assert_eq!(
        DeckSession::open(&left.save().unwrap(), 747)
            .unwrap()
            .comments()
            .unwrap()
            .len(),
        7
    );
    assert!(
        left.comments()
            .unwrap()
            .iter()
            .find(|comment| comment.id == followup.comment_id)
            .unwrap()
            .parent_id
            .is_none()
    );
    left.reply_to_comment(
        &context,
        &followup.comment_id,
        "Peer",
        "P",
        "Next reply",
        "2026-09-05T15:02:00Z",
    )
    .unwrap();
    assert_eq!(
        DeckSession::open(&left.save().unwrap(), 748)
            .unwrap()
            .comments()
            .unwrap()
            .len(),
        8
    );
}

#[test]
fn comment_operations_are_individually_undoable_and_publish_one_update() {
    let session = DeckSession::open(DEMO, 711).unwrap();
    let context = EditCtx::local("test");
    let initial = session.snapshot().unwrap();
    let slide = &initial.slides[0];
    let story = &slide
        .shapes
        .iter()
        .find(|shape| !shape.text_stories.is_empty())
        .unwrap()
        .text_stories[0]
        .id;
    let updates = Arc::new(Mutex::new(Vec::new()));
    let recorded = updates.clone();
    let _subscription = session
        .observe_update_v1(move |update| recorded.lock().unwrap().push(update))
        .unwrap();
    let mut states = vec![initial.clone()];
    session
        .insert_text(&context, story, 0, "Before", &TextStyle::default())
        .unwrap();
    states.push(session.snapshot().unwrap());
    session
        .set_comment_flavor(&context, CommentFlavor::Modern)
        .unwrap();
    states.push(session.snapshot().unwrap());
    let root = session
        .add_comment(
            &context,
            &slide.id,
            "Ada",
            "AL",
            "Root 😀",
            "2026-09-05T10:00:00Z",
            12345,
            -6789,
        )
        .unwrap();
    states.push(session.snapshot().unwrap());
    session
        .reply_to_comment(
            &context,
            &root.comment_id,
            "Grace",
            "GH",
            "Reply",
            "2026-09-05T10:01:00Z",
        )
        .unwrap();
    states.push(session.snapshot().unwrap());
    session
        .set_comment_status(&context, &root.comment_id, true)
        .unwrap();
    states.push(session.snapshot().unwrap());
    session.remove_comment(&context, &root.comment_id).unwrap();
    states.push(session.snapshot().unwrap());
    session
        .insert_text(&context, story, 0, "After", &TextStyle::default())
        .unwrap();
    states.push(session.snapshot().unwrap());
    assert_eq!(updates.lock().unwrap().len(), 7);
    let peer = DeckSession::open(DEMO, 712).unwrap();
    for (index, event) in updates.lock().unwrap().iter().enumerate() {
        peer.apply_update_v1(&event.update).unwrap();
        assert_eq!(peer.snapshot().unwrap(), states[index + 1]);
    }
    for expected in states[..states.len() - 1].iter().rev() {
        assert!(session.undo());
        assert_eq!(&session.snapshot().unwrap(), expected);
    }
    assert!(!session.can_undo());
    for expected in &states[1..] {
        assert!(session.redo());
        assert_eq!(&session.snapshot().unwrap(), expected);
    }
    assert!(!session.can_redo());
}

#[test]
fn comment_strings_reject_xml_controls_without_committing() {
    let session = DeckSession::open(DEMO, 713).unwrap();
    let context = EditCtx::local("test");
    session
        .set_comment_flavor(&context, CommentFlavor::Modern)
        .unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let root = session
        .add_comment(
            &context,
            &slide,
            "Ada",
            "AL",
            "Root",
            "2026-09-05T10:00:00Z",
            0,
            0,
        )
        .unwrap();
    let before = session.encode_state_as_update_v1();
    for index in 0..4 {
        let mut fields = ["Ada", "AL", "Text", "2026-09-05T10:00:00Z"];
        fields[index] = "bad\0value";
        assert!(matches!(
            session.add_comment(
                &context, &slide, fields[0], fields[1], fields[2], fields[3], 0, 0
            ),
            Err(EditError::InvalidText(_))
        ));
        assert!(matches!(
            session.reply_to_comment(
                &context,
                &root.comment_id,
                fields[0],
                fields[1],
                fields[2],
                fields[3]
            ),
            Err(EditError::InvalidText(_))
        ));
        assert_eq!(session.encode_state_as_update_v1(), before);
    }
}

#[test]
fn modern_positions_and_plain_text_follow_drawingml() {
    let session = DeckSession::open(MODERN, 714).unwrap();
    let comments = session.comments().unwrap();
    let root = comments
        .iter()
        .find(|comment| comment.created.as_deref() == Some("2024-12-30T20:26:06.503Z"))
        .unwrap();
    assert_eq!((root.x_emu, root.y_emu), (12345, -6789));
    assert_eq!(root.author, "Mary Smith");
    assert_eq!(root.initials, "MS");
    assert_eq!(root.text, "Needs a source.\nEmoji 😀\nSecond paragraph.");
    let slide = session.snapshot().unwrap().slides[2].id.clone();
    session
        .add_comment(
            &EditCtx::local("test"),
            &slide,
            "Ada",
            "AL",
            "First 😀\nSecond",
            "2026-09-05T10:00:00Z",
            12345,
            -6789,
        )
        .unwrap();
    let saved = session.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let (_, bytes) = parts
        .iter()
        .find(|(name, _)| name == "ppt/comments/modernComment3.xml")
        .unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert!(xml.contains("<p188:pos x=\"12345\" y=\"-6789\"/>"));
    assert!(!xml.contains("cId=\"0\""));
    assert_eq!(xml.matches("<a:p>").count(), 2);
    let reopened = DeckSession::open(&saved, 715).unwrap();
    let added = reopened
        .comments()
        .unwrap()
        .into_iter()
        .find(|comment| comment.author == "Ada")
        .unwrap();
    assert_eq!((added.x_emu, added.y_emu), (12345, -6789));
    assert_eq!(added.text, "First 😀\nSecond");
}

#[test]
fn default_comment_fields_preserve_snapshot_json() {
    let session = DeckSession::open(DEMO, 716).unwrap();
    let snapshot = session.snapshot().unwrap();
    let json = serde_json::to_value(&snapshot).unwrap();
    assert!(json.get("comments").is_none());
    assert!(json.get("commentFlavor").is_none());
    assert_eq!(
        serde_json::from_value::<DeckSnapshot>(json).unwrap(),
        snapshot
    );
    let package = serde_json::to_value(session.package()).unwrap();
    for key in ["comments", "commentAuthors", "commentFlavor"] {
        assert!(package.get(key).is_none());
    }
    let restored: pptx_parse::PptxPackage = serde_json::from_value(package).unwrap();
    assert!(restored.comments.is_empty());
    assert!(restored.comment_authors.is_empty());
    assert_eq!(restored.comment_flavor, None);
}

#[test]
fn renamed_author_parts_are_reused_and_removed() {
    let mut original = ooxml_opc::unzip_parts(MODERN).unwrap();
    for (path, bytes) in &mut original {
        if path == "ppt/authors.xml" {
            *path = "ppt/people/team.xml".to_owned();
        } else if path == "[Content_Types].xml" || path == "ppt/_rels/presentation.xml.rels" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace("authors.xml", "people/team.xml")
                .into_bytes();
        }
    }
    let source = ooxml_opc::rezip_parts(&original).unwrap();
    let session = DeckSession::open(&source, 717).unwrap();
    let context = EditCtx::local("test");
    let root = session
        .comments()
        .unwrap()
        .into_iter()
        .find(|comment| comment.parent_id.is_none())
        .unwrap();
    session
        .set_comment_status(&context, &root.id, true)
        .unwrap();
    let saved = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    assert!(saved.iter().any(|(path, _)| path == "ppt/people/team.xml"));
    assert!(!saved.iter().any(|(path, _)| path == "ppt/authors.xml"));
    let rels = &saved
        .iter()
        .find(|(path, _)| path == "ppt/_rels/presentation.xml.rels")
        .unwrap()
        .1;
    assert!(
        String::from_utf8(rels.clone())
            .unwrap()
            .contains("people/team.xml")
    );
    for root in session
        .comments()
        .unwrap()
        .iter()
        .filter(|comment| comment.parent_id.is_none())
    {
        session.remove_comment(&context, &root.id).unwrap();
    }
    let saved = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    assert!(
        !saved
            .iter()
            .any(|(path, _)| path == "ppt/people/team.xml" || path.starts_with("ppt/comments/"))
    );
    for (path, bytes) in saved {
        if path == "[Content_Types].xml" || path.ends_with(".rels") {
            let xml = String::from_utf8(bytes).unwrap();
            assert!(!xml.contains("people/team.xml"));
            assert!(!xml.contains("/relationships/comments"));
            assert!(!xml.contains("/relationships/authors"));
        }
    }
}

#[test]
fn deleting_slides_removes_comments_and_last_authors() {
    let session = DeckSession::open(MODERN, 718).unwrap();
    let context = EditCtx::local("test");
    let initial = session.snapshot().unwrap();
    session
        .delete_slide(&context, &initial.slides[0].id)
        .unwrap();
    let comments = session.comments().unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].slide_id, initial.slides[1].id);
    let remaining = "ppt/comments/modernComment_257_3ADE68B1.xml";
    let saved = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    assert_eq!(
        saved.iter().find(|(path, _)| path == remaining).unwrap().1,
        session.package().part_bytes(remaining).unwrap()
    );
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), initial);
    assert_eq!(
        ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap(),
        ooxml_opc::unzip_parts(MODERN).unwrap()
    );
    assert!(session.redo());
    session
        .delete_slide(&context, &initial.slides[1].id)
        .unwrap();
    assert!(session.comments().unwrap().is_empty());
    let saved = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    assert!(
        !saved
            .iter()
            .any(|(path, _)| path == "ppt/authors.xml" || path.starts_with("ppt/comments/"))
    );
}
