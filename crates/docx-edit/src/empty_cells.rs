//! A table cell whose story a table of a live story names, while that story is missing (an undo
//! deleted its entry) or removed, reads as one empty paragraph and refuses every write with
//! [`crate::EditError::CellHasNoStory`]. Reads never write: they answer from a scratch document
//! holding exactly that paragraph.

use yrs::types::text::YChange;
use yrs::{Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef, Transact};

use crate::{
    EditError, EditResult, EditingDoc, KIND_KEY, OpError, PARA_ID, PILCROW_KIND, STORIES,
    map_string, out_len,
};

fn names_cell<T: ReadTxn>(table: &MapRef, txn: &T, story_id: &str) -> bool {
    let Some(Out::Any(Any::Array(rows))) = table.get(txn, "rows") else {
        return false;
    };
    rows.iter().any(|row| {
        let Any::Map(row) = row else {
            return false;
        };
        let Some(Any::Array(cells)) = row.get("cells") else {
            return false;
        };
        cells.iter().any(|cell| {
            matches!(cell, Any::Map(cell)
                if matches!(cell.get("story"), Some(Any::String(named)) if named.as_ref() == story_id))
        })
    })
}

/// The first table embed of a live story that names `story_id` as a cell and that `accept`
/// takes: the story holding it and the embed's index there. Every live story is scanned, so a
/// table nested in a cell, a header or a note counts wherever its embed lives.
pub(crate) fn naming_table<T: ReadTxn>(
    txn: &T,
    story_id: &str,
    mut accept: impl FnMut(&TextRef, u32) -> bool,
) -> Option<(TextRef, u32)> {
    let stories = txn.get_map(STORIES)?;
    for (_, value) in crate::live_stories(&stories, txn) {
        let Out::YText(story) = value else {
            continue;
        };
        let mut offset = 0u32;
        for diff in story.diff(txn, YChange::identity) {
            if let Out::YMap(map) = &diff.insert
                && map_string(map, txn, KIND_KEY).as_deref() == Some("table")
                && names_cell(map, txn, story_id)
                && accept(&story, offset)
            {
                return Some((story, offset));
            }
            offset += out_len(&diff.insert);
        }
    }
    None
}

/// Whether a table of a live story names `story_id` as a cell.
pub(crate) fn named_by_live_table<T: ReadTxn>(txn: &T, story_id: &str) -> bool {
    naming_table(txn, story_id, |_, _| true).is_some()
}

/// Whether `story_id` is a cell that a table of a live story names while its own story is
/// missing or removed.
pub(crate) fn is_empty_cell<T: ReadTxn>(txn: &T, story_id: &str) -> bool {
    crate::story_text(txn, story_id).is_none() && named_by_live_table(txn, story_id)
}

/// Refuses a write into `story_id` when it is a cell that a table of a live story names while its
/// story is missing or removed. Write entries call this before they check their other arguments,
/// so a write aimed at such a cell has this one outcome.
pub(crate) fn refuse_write<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<()> {
    if is_empty_cell(txn, story_id) {
        return Err(EditError::CellHasNoStory(story_id.to_owned()));
    }
    Ok(())
}

impl EditingDoc {
    /// [`refuse_write`] in a read transaction of its own; call it before opening a write.
    pub(crate) fn refuse_cell_write(&self, story_id: &str) -> EditResult<()> {
        refuse_write(&self.yrs_doc().transact(), story_id)
    }
}

/// The key of the one paragraph an empty cell reads as: the same on every replica.
pub(crate) fn empty_cell_para_key(story_id: &str) -> String {
    format!("{story_id}#p0")
}

/// The empty cell whose one paragraph `para_id` keys, when it is one.
pub(crate) fn empty_cell_of_paragraph<'a, T: ReadTxn>(
    txn: &T,
    para_id: &'a str,
) -> Option<&'a str> {
    let story = para_id.strip_suffix("#p0")?;
    is_empty_cell(txn, story).then_some(story)
}

/// Why no paragraph `para_id` exists to write: [`OpError::CellHasNoStory`] for the paragraph an
/// empty cell reads as, else [`OpError::UnknownPara`].
pub(crate) fn unknown_paragraph<T: ReadTxn>(txn: &T, para_id: &str) -> OpError {
    match empty_cell_of_paragraph(txn, para_id) {
        Some(story) => OpError::CellHasNoStory(story.to_owned()),
        None => OpError::UnknownPara(para_id.to_owned()),
    }
}

/// A scratch document whose only story is `story_id` holding one empty paragraph: what an empty
/// cell reads as.
pub(crate) struct EmptyCell {
    doc: EditingDoc,
}

impl EmptyCell {
    pub(crate) fn new(story_id: &str) -> Self {
        let doc = EditingDoc::new(1);
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let stories = txn
                .get_map(STORIES)
                .expect("stories root is declared by EditingDoc::new");
            let story = crate::insert_story(&mut txn, &stories, story_id);
            let pilcrow = story.insert_embed_with_attributes(
                &mut txn,
                0,
                MapPrelim::default(),
                crate::insertion_attrs(None, None),
            );
            pilcrow.insert(&mut txn, KIND_KEY, PILCROW_KIND);
            pilcrow.insert(&mut txn, PARA_ID, empty_cell_para_key(story_id));
        }
        Self { doc }
    }

    pub(crate) fn doc(&self) -> &EditingDoc {
        &self.doc
    }
}

#[cfg(test)]
mod tests {
    use yrs::{Any, Assoc, Transact};

    use super::{empty_cell_para_key, is_empty_cell};
    use crate::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
    use crate::{
        CellLoc, EditCtx, EditingDoc, FormatPolicy, OpError, Position, TableLocator, TableRange,
    };

    fn ctx() -> EditCtx {
        EditCtx::local("Ada", "2026-11-10T00:00:00Z")
    }

    fn exchange(left: &EditingDoc, right: &EditingDoc) {
        let from_left = left.encode_state_as_update_v1();
        let from_right = right.encode_state_as_update_v1();
        right.apply_update_v1(&from_left).unwrap();
        left.apply_update_v1(&from_right).unwrap();
    }

    fn replicas() -> (EditingDoc, EditingDoc) {
        let origin = EditingDoc::new(100);
        origin
            .create_story("body", "Hello", "Normal", "left")
            .unwrap();
        origin
            .insert_table(&ctx(), Position::new("body", 0), 1, 2)
            .unwrap();
        let state = origin.encode_state_as_update_v1();
        let open = |client| {
            let doc = EditingDoc::new(client);
            doc.apply_update_v1(&state).unwrap();
            doc
        };
        (open(81), open(82))
    }

    fn cells(doc: &EditingDoc, locator: &TableLocator) -> Vec<String> {
        let payload = doc.table_payload(locator).unwrap();
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

    fn without_story(doc: &EditingDoc, locator: &TableLocator) -> Vec<String> {
        let txn = doc.yrs_doc().transact();
        cells(doc, locator)
            .into_iter()
            .filter(|story| is_empty_cell(&txn, story))
            .collect()
    }

    #[test]
    fn a_caret_in_a_cell_without_its_story_sits_at_its_table_embed() {
        let (doc, peer) = replicas();
        let mut undo = doc.undo_manager();
        doc.split_cell_grid(&ctx(), &CellLoc::new("body", 0, 0, 1), Some(1), Some(2))
            .unwrap();
        exchange(&doc, &peer);
        peer.insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
            .unwrap();
        exchange(&doc, &peer);
        assert!(undo.undo());
        exchange(&doc, &peer);
        let body = TableLocator::new("body", 0);
        let cell = without_story(&doc, &body).remove(0);
        let table = doc.table_embed_index(&body).unwrap();
        let other = if table == 0 { table + 1 } else { 0 };
        let before = doc.encode_state_as_update_v1();
        let own = doc
            .cached_sticky_index(&cell, 0, Assoc::After)
            .unwrap()
            .unwrap();
        let elsewhere = doc
            .cached_sticky_index("body", other, Assoc::After)
            .unwrap()
            .unwrap();
        assert_eq!(doc.cached_sticky_offset("body", &own).unwrap(), Some(table));
        assert_eq!(doc.cached_sticky_offset(&cell, &own).unwrap(), Some(0));
        let key = empty_cell_para_key(&cell);
        assert!(doc.revision_ids_at(&cell, &key, 0).unwrap().is_empty());
        assert!(doc.cached_sticky_offset(&cell, &elsewhere).is_err());
        assert_eq!(doc.encode_state_as_update_v1(), before);

        assert!(undo.redo());
        assert!(without_story(&doc, &body).is_empty());
        assert_eq!(doc.cached_sticky_offset(&cell, &own).unwrap(), Some(0));
        assert!(doc.cached_sticky_offset(&cell, &elsewhere).is_err());
    }

    #[test]
    fn a_nested_table_a_merge_copied_names_cells_whose_insertion_was_undone() {
        let (doc, peer) = replicas();
        let outer = cells(&doc, &TableLocator::new("body", 0));
        let (left, right) = (outer[0].clone(), outer[1].clone());
        let mut undo = doc.undo_manager();
        doc.insert_table(&ctx(), Position::new(right.as_str(), 0), 1, 2)
            .unwrap();
        exchange(&doc, &peer);
        let at = |column| CellLoc::new("body", 0, 0, column);
        peer.merge_cells(&ctx(), &TableRange::new(at(0), at(1)))
            .unwrap();
        exchange(&doc, &peer);
        assert!(undo.undo());
        exchange(&doc, &peer);
        let nested = TableLocator::new(left.as_str(), 0);
        for replica in [&doc, &peer] {
            let missing = without_story(replica, &nested);
            assert_eq!(missing.len(), 2, "{missing:?}");
            for cell in &missing {
                assert!(cell.starts_with(&format!("{right}:")), "{cell}");
                let paragraphs = replica.paragraphs(cell).unwrap();
                assert_eq!(paragraphs.len(), 1);
                assert_eq!(paragraphs[0].text, "");
                let before = replica.encode_state_as_update_v1();
                assert!(matches!(
                    replica.insert_text(
                        &ctx(),
                        Position::new(cell.as_str(), 0),
                        "x",
                        FormatPolicy::Plain,
                    ),
                    Err(OpError::CellHasNoStory(_))
                ));
                assert_eq!(replica.encode_state_as_update_v1(), before);
            }
            let blocks = yrs_doc_to_layout_blocks(replica, "body", &RenderEnv::default()).unwrap();
            let lowered = serde_json::to_string(&blocks).unwrap();
            assert!(lowered.contains("\"table\""), "{lowered}");
        }
    }
}
