//! Table embeds whose payload lost one of its keys. Each table operation rewrites `tblPr`, `grid`
//! and `rows` as separate map entries, so two collaborators who each change a table and undo in
//! crossing order can leave every entry deleted. Such a table lays out as nothing, takes no table
//! operation and blocks saving until its last undo is redone or the table is deleted.

use std::sync::atomic::Ordering;

use yrs::types::text::YChange;
use yrs::types::{EntryChange, Event, Events};
use yrs::{Map, MapRef, Out, ReadTxn, Text, Transact, TransactionMut};

use crate::{EditingDoc, KIND_KEY, STORIES, TableLocator, map_string};

/// The refusal text of a table operation or save over an unavailable table.
pub const TABLE_UNAVAILABLE: &str = "a table is unavailable after a concurrent undo: redo on the collaborator who undid, or delete the table";

const PAYLOAD_KEYS: [&str; 3] = ["tblPr", "grid", "rows"];

/// Whether the table embed `table` lacks any of its payload entries.
pub(crate) fn is_unavailable<T: ReadTxn>(table: &MapRef, txn: &T) -> bool {
    PAYLOAD_KEYS.iter().any(|key| table.get(txn, key).is_none())
}

/// Whether a commit's story events may have left a table unavailable: a payload entry removed,
/// or stories added or replaced at the root, as a late joiner's first update does.
pub(crate) fn may_have_left_unavailable(txn: &TransactionMut, events: &Events) -> bool {
    events.iter().any(|event| match event {
        Event::Map(event) => {
            event.path().is_empty()
                || event.keys(txn).iter().any(|(key, change)| {
                    PAYLOAD_KEYS.contains(&key.as_ref())
                        && matches!(change, EntryChange::Removed(_))
                })
        }
        _ => false,
    })
}

impl EditingDoc {
    /// Every unavailable table, by story and ordinal. Scans only after a commit that may have
    /// left one; otherwise it answers from the last scan, which found none.
    pub(crate) fn unavailable_tables(&self) -> Vec<TableLocator> {
        if !self.unavailable_tables_dirty.load(Ordering::Relaxed) {
            return Vec::new();
        }
        let txn = self.yrs_doc().transact();
        let mut found = Vec::new();
        if let Some(stories) = txn.get_map(STORIES) {
            let mut ids: Vec<(String, yrs::TextRef)> = crate::live_stories(&stories, &txn)
                .filter_map(|(id, value)| match value {
                    Out::YText(text) => Some((id.to_owned(), text)),
                    _ => None,
                })
                .collect();
            ids.sort_by(|left, right| left.0.cmp(&right.0));
            for (id, story) in ids {
                let mut ordinal = 0u32;
                for diff in story.diff(&txn, YChange::identity) {
                    if let Out::YMap(map) = &diff.insert
                        && map_string(map, &txn, KIND_KEY).as_deref() == Some("table")
                    {
                        if is_unavailable(map, &txn) {
                            found.push(TableLocator::new(id.clone(), ordinal));
                        }
                        ordinal += 1;
                    }
                }
            }
        }
        if found.is_empty() {
            self.unavailable_tables_dirty
                .store(false, Ordering::Relaxed);
        }
        found
    }
}

#[cfg(test)]
mod tests {
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
        assert!(origin.unavailable_tables().is_empty());
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
        assert!(a.unavailable_tables().is_empty());
        assert!(b.unavailable_tables().is_empty());
        assert!(undo_a.undo());
        exchange(&a, &b);
        assert!(undo_b.undo());
        exchange(&a, &b);
        let expected = vec![TableLocator::new("body", 0)];
        assert_eq!(a.unavailable_tables(), expected);
        assert_eq!(b.unavailable_tables(), expected);
        let late = open(70);
        late.apply_update_v1(&a.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(late.unavailable_tables(), expected);
        assert!(undo_b.redo());
        exchange(&a, &b);
        assert!(a.unavailable_tables().is_empty());
        assert!(b.unavailable_tables().is_empty());
        assert!(
            !a.unavailable_tables_dirty
                .load(std::sync::atomic::Ordering::Relaxed)
        );
    }
}
