//! Cell stories that two replicas create concurrently stay apart, so one replica removing its
//! table, row or column never takes a story another replica's table still names. Cell ids that
//! came from the opened document are the same on every replica.

use docx_edit::bridge::RenderEnv;
use docx_edit::canonical::{project_story, to_canonical_bytes};
use docx_edit::*;
use yrs::Any;

const DATE: &str = "2026-10-17T00:00:00Z";
const LEFT: u64 = 51;
const RIGHT: u64 = 52;

#[derive(Clone, Copy, Debug)]
enum Removal {
    Delete,
    Undo,
}

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn fixture() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/field-code-paragraphs/cell-field-code-paragraphs.docx"),
    )
    .unwrap()
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

fn typed_table(doc: &EditingDoc, text: &str) -> String {
    let table = doc
        .insert_table(&ctx(), Position::new("body", 0), 2, 2)
        .unwrap();
    let cell = table.created_story_ids[0].clone();
    doc.insert_text(&ctx(), Position::new(&cell, 0), text, FormatPolicy::Plain)
        .unwrap();
    cell
}

fn tables(doc: &EditingDoc) -> Vec<Vec<String>> {
    (0..)
        .map_while(|index| doc.table_payload(&TableLocator::new("body", index)).ok())
        .map(|payload| {
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
        })
        .collect()
}

fn text(doc: &EditingDoc, story: &str) -> Vec<String> {
    doc.paragraphs(story)
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

fn cell_ids(doc: &EditingDoc) -> Vec<String> {
    let mut ids: Vec<String> = doc
        .list_stories()
        .stories
        .into_iter()
        .filter(|story| story.kind == StoryInfoKind::TableCell)
        .map(|story| story.story)
        .collect();
    ids.sort();
    ids
}

fn assert_lowers(engine: &EngineSession, case: &str) {
    if let Err(error) = engine.lower_story_json("body", &RenderEnv::default()) {
        panic!(
            "{case}: replica {} cannot lower the body: {error:?}",
            engine.doc().client_id()
        );
    }
}

#[test]
fn removing_a_table_leaves_the_concurrent_table_whole() {
    for removal in [Removal::Delete, Removal::Undo] {
        for actor_is_left in [true, false] {
            let case = format!(
                "{removal:?} on the {} replica",
                if actor_is_left { "left" } else { "right" }
            );
            let origin = EditingDoc::new(100);
            origin
                .create_story("body", "Hello", "Normal", "left")
                .unwrap();
            let (left, right) = replicas(&origin);
            let (actor, peer) = if actor_is_left {
                (&left, &right)
            } else {
                (&right, &left)
            };
            let mut undo = actor.doc().undo_manager();
            let mine = typed_table(actor.doc(), "mine");
            let theirs = typed_table(peer.doc(), "theirs");
            exchange(&left, &right);
            match removal {
                Removal::Delete => {
                    let index = tables(actor.doc())
                        .iter()
                        .position(|cells| cells.contains(&mine))
                        .unwrap();
                    actor
                        .doc()
                        .delete_table(&ctx(), &TableLocator::new("body", index as u32))
                        .unwrap();
                }
                Removal::Undo => while undo.undo() {},
            }
            exchange(&left, &right);
            for engine in [&left, &right] {
                assert_lowers(engine, &case);
                let tables = tables(engine.doc());
                assert_eq!(tables.len(), 1, "{case}");
                assert_eq!(tables[0][0], theirs, "{case}");
                assert_eq!(text(engine.doc(), &theirs), ["theirs"], "{case}");
            }
        }
    }
}

#[test]
fn opened_cell_ids_stay_put_and_edit_born_ids_never_change() {
    let origin = EditingDoc::new(100);
    seed_from_docx(&origin, &fixture()).unwrap();
    let opened = cell_ids(&origin);
    assert!(!opened.is_empty());
    let (left, right) = replicas(&origin);
    assert_eq!(cell_ids(left.doc()), opened);
    assert_eq!(cell_ids(right.doc()), opened);

    let mut undo = left.doc().undo_manager();
    let row = left
        .doc()
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap()
        .created_story_ids;
    let table = right
        .doc()
        .insert_table(&ctx(), Position::new("body", 0), 1, 2)
        .unwrap()
        .created_story_ids;
    assert!(
        row.iter().all(|id| id.ends_with(&format!("~{LEFT}"))),
        "{row:?}"
    );
    assert!(
        table.iter().all(|id| id.ends_with(&format!("~{RIGHT}"))),
        "{table:?}"
    );
    exchange(&left, &right);

    let mut born: Vec<String> = opened.iter().chain(&row).chain(&table).cloned().collect();
    born.sort();
    for engine in [&left, &right] {
        assert_eq!(cell_ids(engine.doc()), born);
        assert_lowers(engine, "after the concurrent inserts");
    }

    assert!(undo.undo());
    exchange(&left, &right);
    let mut without_row: Vec<String> = opened.iter().chain(&table).cloned().collect();
    without_row.sort();
    for engine in [&left, &right] {
        assert_eq!(cell_ids(engine.doc()), without_row);
        assert_lowers(engine, "after undo");
    }

    assert!(undo.redo());
    exchange(&left, &right);
    for engine in [&left, &right] {
        assert_eq!(cell_ids(engine.doc()), born);
        assert_lowers(engine, "after redo");
    }
}

#[test]
fn a_cell_born_on_one_replica_is_listed_and_edited_by_the_same_id_on_another() {
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    let (left, right) = replicas(&origin);
    let cell = typed_table(left.doc(), "left");
    exchange(&left, &right);
    assert!(cell_ids(right.doc()).contains(&cell));
    right
        .doc()
        .insert_text(
            &ctx(),
            Position::new(&cell, 0),
            "right ",
            FormatPolicy::Plain,
        )
        .unwrap();
    exchange(&left, &right);
    for engine in [&left, &right] {
        assert_eq!(text(engine.doc(), &cell), ["right left"]);
        assert_eq!(cell_ids(engine.doc()), cell_ids(left.doc()));
    }
}

#[test]
fn the_canonical_form_does_not_depend_on_the_replica_that_made_the_cells() {
    let origin = EditingDoc::new(100);
    origin
        .create_story("body", "Hello", "Normal", "left")
        .unwrap();
    let (left, right) = replicas(&origin);
    let left_cell = typed_table(left.doc(), "same");
    let right_cell = typed_table(right.doc(), "same");
    assert_ne!(left_cell, right_cell);
    let canonical =
        |doc: &EditingDoc, story: &str| to_canonical_bytes(&project_story(doc, story).unwrap());
    assert_eq!(
        canonical(left.doc(), "body"),
        canonical(right.doc(), "body")
    );
    assert_eq!(
        canonical(left.doc(), &left_cell),
        canonical(right.doc(), &right_cell)
    );
}
