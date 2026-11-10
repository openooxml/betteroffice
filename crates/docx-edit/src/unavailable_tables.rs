//! Table embeds whose payload lost one of its keys. Each table operation rewrites `tblPr`, `grid`
//! and `rows` as separate map entries, so two collaborators who each change a table and undo in
//! crossing order can leave every entry deleted. Such a table lays out as nothing, takes no table
//! operation and blocks saving until its last undo is redone or the table is deleted.

use std::sync::atomic::Ordering;

use yrs::types::{Delta, EntryChange, Event, Events};
use yrs::{Map, MapRef, Out, ReadTxn, TransactionMut};

use crate::{EditingDoc, TableLocator};

/// The refusal text of a table operation or save over an unavailable table.
pub const TABLE_UNAVAILABLE: &str = "a table is unavailable after a concurrent undo: redo on the collaborator who undid, or delete the table";

const PAYLOAD_KEYS: [&str; 3] = ["tblPr", "grid", "rows"];

/// Whether the table embed `table` lacks any of its payload entries.
pub(crate) fn is_unavailable<T: ReadTxn>(table: &MapRef, txn: &T) -> bool {
    PAYLOAD_KEYS.iter().any(|key| table.get(txn, key).is_none())
}

/// Whether a commit's story events may have changed which tables are unavailable: a payload
/// entry removed, an embed inserted (undoing a table's deletion brings a keyless one back), or
/// stories added, replaced or removed at the root. Only the worker's notice reads it.
pub(crate) fn may_have_left_unavailable(txn: &TransactionMut, events: &Events) -> bool {
    events.iter().any(|event| match event {
        Event::Map(event) => {
            event.path().is_empty()
                || event.keys(txn).iter().any(|(key, change)| {
                    PAYLOAD_KEYS.contains(&key.as_ref())
                        && matches!(change, EntryChange::Removed(_))
                })
        }
        Event::Text(event) => event
            .delta(txn)
            .iter()
            .any(|delta| matches!(delta, Delta::Inserted(Out::YMap(_), _))),
        _ => false,
    })
}

impl EditingDoc {
    /// Every unavailable table a reader can reach, by story and ordinal: the geometry outline
    /// walk (the editor projection's walk) from each root story the story list names (the body,
    /// headers and footers by their section references, footnotes and endnotes), through the
    /// cells of complete tables and block controls' stories. A story only an unavailable or
    /// deleted table named is out of reach and never counts. `None` when a walk cannot finish.
    pub(crate) fn unavailable_tables(&self) -> Option<Vec<TableLocator>> {
        use crate::stories::StoryInfoKind;
        let roots: Vec<String> = self
            .list_stories()
            .stories
            .into_iter()
            .filter(|info| {
                matches!(
                    info.kind,
                    StoryInfoKind::Body
                        | StoryInfoKind::Header
                        | StoryInfoKind::Footer
                        | StoryInfoKind::Footnote
                        | StoryInfoKind::Endnote
                )
            })
            .map(|info| info.story)
            .collect();
        let mut found = Vec::new();
        for root in roots {
            found.extend(self.unavailable_tables_under(&root).ok()?);
        }
        found.sort_by(|left, right| {
            (left.story.as_str(), left.table_index).cmp(&(right.story.as_str(), right.table_index))
        });
        if found.is_empty() {
            self.unavailable_tables_dirty
                .store(false, Ordering::Relaxed);
        }
        Some(found)
    }

    /// [`EditingDoc::unavailable_tables`], scanning only after a commit that may have changed
    /// them; otherwise empty, as the last scan found. For the worker's notice, never for a save.
    pub(crate) fn noticed_unavailable_tables(&self) -> Vec<TableLocator> {
        if !self.unavailable_tables_dirty.load(Ordering::Relaxed) {
            return Vec::new();
        }
        self.unavailable_tables().unwrap_or_default()
    }
}

#[cfg(test)]
mod tests {
    use yrs::{Text, Transact};

    use crate::{CellLoc, EditCtx, EditingDoc, Position, TableLocator};

    fn ctx() -> EditCtx {
        EditCtx::local("Ada", "2026-11-10T00:00:00Z")
    }

    fn exchange(left: &EditingDoc, right: &EditingDoc) {
        let from_left = left.encode_state_as_update_v1();
        let from_right = right.encode_state_as_update_v1();
        right.apply_update_v1(&from_left).unwrap();
        left.apply_update_v1(&from_right).unwrap();
    }

    #[test]
    fn lists_the_table_crossing_undos_emptied_until_the_second_undoer_redoes() {
        let origin = EditingDoc::new(100);
        origin
            .create_story("body", "Hello", "Normal", "left")
            .unwrap();
        origin
            .insert_table(&ctx(), Position::new("body", 0), 1, 2)
            .unwrap();
        let none = Some(Vec::new());
        assert_eq!(origin.unavailable_tables(), none);
        let state = origin.encode_state_as_update_v1();
        let open = |client| {
            let doc = EditingDoc::new(client);
            doc.apply_update_v1(&state).unwrap();
            doc
        };
        let (a, b) = (open(61), open(62));
        let mut undo_a = a.undo_manager();
        let mut undo_b = b.undo_manager();
        let cell = CellLoc::new("body", 0, 0, 0);
        a.insert_row(&ctx(), &cell, true).unwrap();
        exchange(&a, &b);
        b.set_column_width(&ctx(), &cell, 2000.0).unwrap();
        assert_eq!(a.unavailable_tables(), none);
        assert_eq!(b.unavailable_tables(), none);
        assert!(undo_a.undo());
        exchange(&a, &b);
        assert!(undo_b.undo());
        exchange(&a, &b);
        let expected = Some(vec![TableLocator::new("body", 0)]);
        assert_eq!(a.unavailable_tables(), expected);
        assert_eq!(b.unavailable_tables(), expected);
        assert_eq!(a.noticed_unavailable_tables(), expected.clone().unwrap());
        let late = open(70);
        late.apply_update_v1(&a.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(late.unavailable_tables(), expected);
        assert!(undo_b.redo());
        exchange(&a, &b);
        assert_eq!(a.unavailable_tables(), none);
        assert_eq!(b.unavailable_tables(), none);
        assert!(
            !a.unavailable_tables_dirty
                .load(std::sync::atomic::Ordering::Relaxed)
        );
    }

    #[test]
    fn a_table_in_a_cell_counts_only_while_its_parent_table_is_complete() {
        let doc = EditingDoc::new(71);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        let outer = doc
            .insert_table(&ctx(), Position::new("body", 0), 1, 1)
            .unwrap();
        let cell = outer.created_story_ids[0].clone();
        doc.insert_table(&ctx(), Position::new(&cell, 0), 1, 1)
            .unwrap();
        let strip = |story: &str| {
            let mut txn = doc.yrs_doc().transact_mut();
            let text = crate::story_ref(&txn, story).unwrap();
            let table = text
                .diff(&txn, yrs::types::text::YChange::identity)
                .into_iter()
                .find_map(|diff| match diff.insert {
                    yrs::Out::YMap(map) => (crate::map_string(&map, &txn, crate::KIND_KEY)
                        .as_deref()
                        == Some("table"))
                    .then_some(map),
                    _ => None,
                })
                .unwrap();
            for key in super::PAYLOAD_KEYS {
                yrs::Map::remove(&table, &mut txn, key);
            }
        };
        strip(&cell);
        assert_eq!(
            doc.unavailable_tables(),
            Some(vec![TableLocator::new(cell.clone(), 0)])
        );
        strip("body");
        assert_eq!(
            doc.unavailable_tables(),
            Some(vec![TableLocator::new("body", 0)]),
            "the inner table is out of reach once its parent lost its rows"
        );
        doc.delete_table(&ctx(), &TableLocator::new("body", 0))
            .unwrap();
        assert_eq!(doc.unavailable_tables(), Some(Vec::new()));
    }
}
