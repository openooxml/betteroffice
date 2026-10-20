//! Undoing a removal brings back what a peer wrote inside the removed content concurrently.

use docx_edit::bridge::RenderEnv;
use docx_edit::structured::{ExportOptions, RevisionView};
use docx_edit::*;
use yrs::{Any, Map, ReadTxn, Transact};

const DATE: &str = "2026-10-19T00:00:00Z";
const LEFT: u64 = 61;
const RIGHT: u64 = 62;

#[derive(Clone, Copy, Debug)]
enum Removal {
    Table,
    Row,
    Column,
}

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn replicas(origin: &EditingDoc) -> (EngineSession, EngineSession) {
    let state = origin.encode_state_as_update_v1();
    let open = |client| {
        let engine = EngineSession::new(client);
        engine.doc().apply_update_v1(&state).unwrap();
        engine
    };
    (open(LEFT), open(RIGHT))
}

fn exchange(left: &EngineSession, right: &EngineSession) {
    let from_left = left.doc().encode_state_as_update_v1();
    let from_right = right.doc().encode_state_as_update_v1();
    right.doc().apply_update_v1(&from_left).unwrap();
    left.doc().apply_update_v1(&from_right).unwrap();
}

fn cell_stories(doc: &EditingDoc) -> Vec<String> {
    let payload = doc.table_payload(&TableLocator::new("body", 0)).unwrap();
    let Some(Any::Array(rows)) = payload.get("rows") else {
        panic!("table without rows: {payload:?}");
    };
    rows.iter()
        .flat_map(|row| {
            let Any::Map(row) = row else {
                panic!("row is not a map: {row:?}");
            };
            let Some(Any::Array(cells)) = row.get("cells") else {
                panic!("row without cells: {row:?}");
            };
            cells.iter().map(|cell| match cell {
                Any::Map(cell) => match cell.get("story") {
                    Some(Any::String(story)) => story.to_string(),
                    other => panic!("cell without story: {other:?}"),
                },
                other => panic!("cell is not a map: {other:?}"),
            })
        })
        .collect()
}

fn text(doc: &EditingDoc, story: &str) -> String {
    doc.paragraphs(story)
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect::<Vec<_>>()
        .join("\n")
}

fn table_origin() -> EditingDoc {
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    let table = origin
        .insert_table(&ctx(), Position::new("body", 0), 2, 2)
        .unwrap();
    for (index, story) in table.created_story_ids.iter().enumerate() {
        origin
            .insert_text(
                &ctx(),
                Position::new(story, 0),
                &format!("cell{index}"),
                FormatPolicy::Plain,
            )
            .unwrap();
    }
    origin
}

fn remove(doc: &EditingDoc, removal: Removal) -> TableReceipt {
    let first = CellLoc::new("body", 0, 0, 0);
    match removal {
        Removal::Table => doc.delete_table(&ctx(), &TableLocator::new("body", 0)),
        Removal::Row => doc.delete_row(&ctx(), &TableRange::new(first.clone(), first)),
        Removal::Column => doc.delete_column(&ctx(), &TableRange::new(first.clone(), first)),
    }
    .unwrap()
}

fn assert_absent(engine: &EngineSession, removed: &[String]) {
    let doc = engine.doc();
    let listed = doc.list_stories().stories;
    let identities = doc.paragraph_identities();
    let matches = doc.search_text("cell", true, None).unwrap();
    for story in removed {
        assert!(doc.story_len(story).is_err(), "{story}");
        assert!(doc.story_segments(story).is_err(), "{story}");
        assert!(project_story(doc, story).is_err(), "{story}");
        assert!(
            engine
                .lower_story_json(story, &RenderEnv::default())
                .is_err()
        );
        assert!(listed.iter().all(|info| info.story != *story));
        assert!(matches.iter().all(|found| found.story != *story));
        assert!(identities.paragraphs.iter().all(|identity| {
            !matches!(&identity.paragraph, ParagraphRef::Session { story: id, .. } if id == story)
        }));
    }
    assert!(
        engine
            .lower_story_json("body", &RenderEnv::default())
            .is_ok()
    );
    let exported = doc
        .export_structured(&ExportOptions::new(RevisionView::Accepted))
        .unwrap();
    assert!(
        !serde_json::to_string(&exported.content)
            .unwrap()
            .contains("cell0")
    );
}

#[test]
fn undoing_a_table_removal_keeps_a_peers_concurrent_cell_edit() {
    for removal in [Removal::Table, Removal::Row, Removal::Column] {
        let origin = table_origin();
        let (actor, peer) = replicas(&origin);
        let cell = cell_stories(peer.doc())[0].clone();
        for engine in [&actor, &peer] {
            assert_eq!(text(engine.doc(), &cell), "cell0");
            engine
                .lower_story_json(&cell, &RenderEnv::default())
                .unwrap();
            engine.doc().paragraph_identities();
        }
        peer.doc()
            .insert_text(
                &ctx(),
                Position::new(&cell, 0),
                "peer ",
                FormatPolicy::Plain,
            )
            .unwrap();
        let mut undo = actor.doc().undo_manager();
        let removed = remove(actor.doc(), removal).deleted_story_ids;
        exchange(&actor, &peer);
        for cycle in 0..3 {
            for engine in [&actor, &peer] {
                assert_absent(engine, &removed);
            }
            assert!(undo.undo(), "{removal:?}: nothing to undo");
            exchange(&actor, &peer);
            for engine in [&actor, &peer] {
                let mut texts: Vec<String> = cell_stories(engine.doc())
                    .iter()
                    .map(|cell| text(engine.doc(), cell))
                    .collect();
                texts.sort();
                assert_eq!(
                    texts,
                    ["cell1", "cell2", "cell3", "peer cell0"],
                    "{removal:?}, cycle {cycle}, replica {}",
                    engine.doc().client_id()
                );
            }
            if cycle < 2 {
                assert!(undo.redo(), "{removal:?}: nothing to redo");
                exchange(&actor, &peer);
            }
        }
    }
}

#[test]
fn undoing_concurrent_table_removals_restores_cells_for_either_marker_winner() {
    for undoer in [LEFT, RIGHT] {
        let origin = table_origin();
        let cells = cell_stories(&origin);
        let (left, right) = replicas(&origin);
        let actor = if undoer == LEFT { &left } else { &right };
        let mut undo = actor.doc().undo_manager();
        let removed = remove(left.doc(), Removal::Table).deleted_story_ids;
        assert_eq!(
            remove(right.doc(), Removal::Table).deleted_story_ids,
            removed
        );
        exchange(&left, &right);
        for cycle in 0..3 {
            for engine in [&left, &right] {
                assert_absent(engine, &removed);
            }
            assert_eq!((undo.undo_depth(), undo.redo_depth()), (1, 0));
            assert!(undo.undo(), "undoer {undoer}, cycle {cycle}");
            assert_eq!((undo.undo_depth(), undo.redo_depth()), (0, 1));
            exchange(&left, &right);
            for engine in [&left, &right] {
                assert_eq!(cell_stories(engine.doc()), cells);
                let listed = engine.doc().list_stories().stories;
                for (index, cell) in cells.iter().enumerate() {
                    assert_eq!(text(engine.doc(), cell), format!("cell{index}"));
                    assert!(listed.iter().any(|info| info.story == *cell));
                    engine
                        .lower_story_json(cell, &RenderEnv::default())
                        .unwrap();
                }
                engine
                    .lower_story_json("body", &RenderEnv::default())
                    .unwrap();
            }
            if cycle < 2 {
                assert!(undo.redo(), "undoer {undoer}, cycle {cycle}");
                exchange(&left, &right);
            }
        }
    }
}

#[test]
fn undoing_an_inserted_row_keeps_a_peers_deleted_table_cells_removed() {
    let origin = table_origin();
    let (actor, peer) = replicas(&origin);
    let mut undo = actor.doc().undo_manager();
    actor
        .doc()
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    let removed = remove(peer.doc(), Removal::Table).deleted_story_ids;
    exchange(&actor, &peer);
    assert!(undo.undo());
    exchange(&actor, &peer);
    for engine in [&actor, &peer] {
        assert!(
            engine
                .doc()
                .table_payload(&TableLocator::new("body", 0))
                .is_err()
        );
        assert_absent(engine, &removed);
    }
}

#[test]
fn undoing_a_table_edit_keeps_a_previously_removed_cell_hidden() {
    for delete_table in [false, true] {
        let origin = table_origin();
        let (actor, peer) = replicas(&origin);
        let cell = cell_stories(actor.doc())[0].clone();
        actor.doc().keep_deleted();
        actor.doc().delete_story(&cell).unwrap();
        exchange(&actor, &peer);
        let mut undo = actor.doc().undo_manager();
        if delete_table {
            remove(actor.doc(), Removal::Table);
        } else {
            actor
                .doc()
                .insert_row(&ctx(), &CellLoc::new("body", 0, 1, 0), true)
                .unwrap();
        }
        assert!(undo.undo());
        exchange(&actor, &peer);
        for engine in [&actor, &peer] {
            assert!(engine.doc().story_len(&cell).is_err());
            assert!(
                engine
                    .doc()
                    .list_stories()
                    .stories
                    .iter()
                    .all(|info| info.story != cell)
            );
        }
    }
}

#[test]
fn undoing_concurrent_table_removals_restores_repeating_control_children() {
    for undoer in [LEFT, RIGHT] {
        let origin = table_origin();
        let cell = cell_stories(&origin)[0].clone();
        let repeat = format!("{cell}:repeat");
        let item = format!("{repeat}:item");
        for (story, content) in [(&repeat, ""), (&item, "Item")] {
            origin
                .create_story(story, content, "Normal", "left")
                .unwrap();
        }
        for (parent, child, kind) in [
            (&cell, &repeat, "repeatingSection"),
            (&repeat, &item, "repeatingSectionItem"),
        ] {
            origin
                .apply_raw_ops(
                    parent,
                    vec![RawOp::InsertEmbed {
                        index: 0,
                        kind: "blockSdt".into(),
                        payload: vec![
                            ("sdtType".into(), Any::from(kind)),
                            ("story".into(), Any::from(child.as_str())),
                        ],
                        attrs: Default::default(),
                    }],
                    &ctx(),
                )
                .unwrap();
        }
        let (left, right) = replicas(&origin);
        let actor = if undoer == LEFT { &left } else { &right };
        let mut undo = actor.doc().undo_manager();
        let removed = remove(left.doc(), Removal::Table).deleted_story_ids;
        assert!(removed.contains(&repeat));
        assert!(removed.contains(&item));
        assert_eq!(
            remove(right.doc(), Removal::Table).deleted_story_ids,
            removed
        );
        exchange(&left, &right);
        for cycle in 0..3 {
            for engine in [&left, &right] {
                assert_absent(engine, &removed);
            }
            assert!(undo.undo(), "undoer {undoer}, cycle {cycle}");
            exchange(&left, &right);
            for engine in [&left, &right] {
                assert_eq!(text(engine.doc(), &item), "Item");
                let listed = engine.doc().list_stories().stories;
                assert!(listed.iter().any(|info| info.story == repeat));
                assert!(listed.iter().any(|info| info.story == item));
                engine
                    .lower_story_json("body", &RenderEnv::default())
                    .unwrap();
            }
            if cycle < 2 {
                assert!(undo.redo(), "undoer {undoer}, cycle {cycle}");
                exchange(&left, &right);
            }
        }
    }
}

#[test]
fn a_deleted_table_has_the_same_canonical_stories_as_a_document_without_it() {
    let doc = table_origin();
    remove(&doc, Removal::Table);
    let plain = EditingDoc::new(101);
    plain
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    let canonical = |doc: &EditingDoc| {
        doc.list_stories()
            .stories
            .into_iter()
            .map(|info| {
                let bytes = to_canonical_bytes(&project_story(doc, &info.story).unwrap());
                (info.story, bytes)
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(canonical(&doc), canonical(&plain));
}

#[test]
fn new_cells_do_not_reuse_removed_story_ids() {
    let doc = table_origin();
    let removed = remove(&doc, Removal::Table).deleted_story_ids;
    let created = doc
        .insert_table(&ctx(), Position::new("body", 0), 2, 2)
        .unwrap()
        .created_story_ids;
    assert!(created.iter().all(|story| !removed.contains(story)));
    for story in &removed {
        assert!(doc.story_len(story).is_err());
    }
    for story in &created {
        assert!(doc.story_len(story).is_ok());
    }
    assert_eq!(doc.list_stories().stories.len(), created.len() + 1);
}

#[test]
fn removed_cell_comment_anchors_stop_resolving_until_undo() {
    let doc = table_origin();
    let cell = cell_stories(&doc)[0].clone();
    let comment = doc
        .add_comment(
            &[StoryRange::new(&cell, 0, 5)],
            "Ada",
            DATE,
            Any::from("Remark"),
        )
        .unwrap();
    let mut undo = doc.undo_manager();
    remove(&doc, Removal::Table);
    assert!(doc.resolve_comment(&comment).is_err());
    assert!(undo.undo());
    let anchors = doc.resolve_comment(&comment).unwrap();
    assert_eq!(anchors.len(), 1);
    assert_eq!(anchors[0].story, cell);
}

#[test]
fn undo_session_restores_nested_cell_stories_with_concurrent_text() {
    let origin = table_origin();
    let parent = cell_stories(&origin)[0].clone();
    let nested = origin
        .insert_table(&ctx(), Position::new(&parent, 0), 1, 1)
        .unwrap()
        .created_story_ids[0]
        .clone();
    let (actor, peer) = replicas(&origin);
    let history = UndoSession::new();
    history.track(actor.doc());
    peer.doc()
        .insert_text(
            &ctx(),
            Position::new(&nested, 0),
            "nested peer",
            FormatPolicy::Plain,
        )
        .unwrap();
    let removed = remove(actor.doc(), Removal::Table).deleted_story_ids;
    assert!(removed.contains(&nested));
    exchange(&actor, &peer);
    for cycle in 0..3 {
        for engine in [&actor, &peer] {
            assert_absent(engine, &removed);
        }
        assert!(history.undo());
        exchange(&actor, &peer);
        for engine in [&actor, &peer] {
            assert_eq!(text(engine.doc(), &nested), "nested peer");
            assert!(
                engine
                    .lower_story_json("body", &RenderEnv::default())
                    .is_ok()
            );
        }
        if cycle < 2 {
            assert!(history.redo());
            exchange(&actor, &peer);
        }
    }
}

#[test]
fn undoing_a_paragraph_removal_keeps_a_peers_concurrent_typing() {
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello world", "Normal", "left")
        .unwrap();
    let (actor, peer) = replicas(&origin);
    peer.doc()
        .insert_text(&ctx(), Position::new("body", 5), ",", FormatPolicy::Plain)
        .unwrap();
    let mut undo = actor.doc().undo_manager();
    actor
        .doc()
        .delete_range(&ctx(), StoryRange::new("body", 0, 11))
        .unwrap();
    exchange(&actor, &peer);
    assert!(undo.undo());
    exchange(&actor, &peer);
    for engine in [&actor, &peer] {
        assert_eq!(text(engine.doc(), "body"), "Hello, world");
    }
}

fn assert_undo_keeps_peer_text(
    actor: &EngineSession,
    peer: &EngineSession,
    undo: &mut DocUndoManager,
    story: &str,
    expected: &str,
) {
    exchange(actor, peer);
    for cycle in 0..3 {
        for engine in [actor, peer] {
            assert!(engine.doc().story_len(story).is_err());
        }
        assert!(undo.undo());
        exchange(actor, peer);
        for engine in [actor, peer] {
            assert_eq!(text(engine.doc(), story), expected);
        }
        if cycle < 2 {
            assert!(undo.redo());
            exchange(actor, peer);
        }
    }
}

#[test]
fn undoing_delete_story_keeps_a_peers_concurrent_edit() {
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    let (actor, peer) = replicas(&origin);
    peer.doc()
        .insert_text(
            &ctx(),
            Position::new("body", 0),
            "peer ",
            FormatPolicy::Plain,
        )
        .unwrap();
    let mut undo = actor.doc().undo_manager();
    actor.doc().delete_story("body").unwrap();
    assert_undo_keeps_peer_text(&actor, &peer, &mut undo, "body", "peer Hello");
}

#[test]
fn delete_story_rejects_an_already_removed_cell_without_affecting_table_undo() {
    let doc = table_origin();
    let cell = cell_stories(&doc)[0].clone();
    let mut undo = doc.undo_manager();
    remove(&doc, Removal::Table);
    assert!(matches!(
        doc.delete_story(&cell),
        Err(EditError::StoryNotFound(id)) if id == cell
    ));
    assert_eq!(undo.undo_depth(), 1);
    assert!(undo.undo());
    assert_eq!(text(&doc, &cell), "cell0");
    assert!(project_story(&doc, "body").is_ok());
}

#[test]
fn undoing_part_metadata_note_removal_keeps_a_peers_concurrent_edit() {
    for (kind, prefix) in [("footnotes", "fn"), ("endnotes", "en")] {
        let origin = EditingDoc::new(100);
        origin
            .create_story("body", "Body", "Normal", "left")
            .unwrap();
        let story = format!("{prefix}:1");
        origin
            .create_story(&story, "Note", "Normal", "left")
            .unwrap();
        let metadata = serde_json::json!({"body": {}, kind: [{"id": 1}]}).to_string();
        origin.set_part_metadata(&ctx(), &metadata, &[]).unwrap();
        let (actor, peer) = replicas(&origin);
        peer.doc()
            .insert_text(
                &ctx(),
                Position::new(&story, 0),
                "peer ",
                FormatPolicy::Plain,
            )
            .unwrap();
        let mut undo = actor.doc().undo_manager();
        actor
            .doc()
            .set_part_metadata(&ctx(), r#"{"body":{}}"#, std::slice::from_ref(&story))
            .unwrap();
        assert_undo_keeps_peer_text(&actor, &peer, &mut undo, &story, "peer Note");
        assert_eq!(actor.doc().part_metadata_json(), metadata);
        assert_eq!(peer.doc().part_metadata_json(), metadata);
    }
}

#[test]
fn undoing_repeating_control_removal_keeps_a_peers_concurrent_edit() {
    let origin = EditingDoc::new(100);
    for (story, content) in [
        ("body", "Body"),
        ("body:repeat", ""),
        ("body:repeat:item", "Item"),
    ] {
        origin
            .create_story(story, content, "Normal", "left")
            .unwrap();
    }
    for (parent, child, kind) in [
        ("body", "body:repeat", "repeatingSection"),
        ("body:repeat", "body:repeat:item", "repeatingSectionItem"),
    ] {
        origin
            .apply_raw_ops(
                parent,
                vec![RawOp::InsertEmbed {
                    index: 0,
                    kind: "blockSdt".into(),
                    payload: vec![
                        ("sdtType".into(), Any::from(kind)),
                        ("story".into(), Any::from(child)),
                    ],
                    attrs: Default::default(),
                }],
                &ctx(),
            )
            .unwrap();
    }
    origin
        .repeat_control_at(&ctx(), "body:repeat", 0, false)
        .unwrap();
    let (actor, peer) = replicas(&origin);
    peer.doc()
        .insert_text(
            &ctx(),
            Position::new("body:repeat:item", 0),
            "peer ",
            FormatPolicy::Plain,
        )
        .unwrap();
    let mut undo = actor.doc().undo_manager();
    actor
        .doc()
        .repeat_control_at(&ctx(), "body:repeat", 0, true)
        .unwrap();
    assert_undo_keeps_peer_text(&actor, &peer, &mut undo, "body:repeat:item", "peer Item");
}

#[test]
fn deleting_a_story_without_history_removes_its_root() {
    let doc = EditingDoc::new(100);
    doc.create_story("body", "Old", "Normal", "left").unwrap();
    doc.delete_story("body").unwrap();
    let txn = doc.yrs_doc().transact();
    assert_eq!(txn.get_map("stories").unwrap().len(&txn), 0);
    assert_eq!(txn.get_map("removed_stories").unwrap().len(&txn), 0);
}

#[test]
fn a_deleted_story_id_can_be_created_again() {
    let doc = EditingDoc::new(100);
    doc.create_story("body", "Old", "Normal", "left").unwrap();
    doc.delete_story("body").unwrap();
    doc.create_story("body", "New", "Normal", "left").unwrap();
    assert_eq!(text(&doc, "body"), "New");
    assert!(project_story(&doc, "body").is_ok());
}

#[test]
fn a_removed_note_id_can_be_added_again_through_part_metadata() {
    let doc = EditingDoc::new(100);
    doc.create_story("body", "Body", "Normal", "left").unwrap();
    doc.create_story("fn:1", "Note", "Normal", "left").unwrap();
    let metadata = serde_json::json!({"body": {}, "footnotes": [{"id": 1}]}).to_string();
    doc.set_part_metadata(&ctx(), &metadata, &[]).unwrap();
    doc.set_part_metadata(&ctx(), r#"{"body":{}}"#, &["fn:1".to_owned()])
        .unwrap();
    assert!(doc.story_len("fn:1").is_err());
    doc.set_part_metadata(&ctx(), &metadata, &[]).unwrap();
    assert!(doc.story_len("fn:1").is_ok());
}
