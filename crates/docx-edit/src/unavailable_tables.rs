//! Table embeds whose payload lost one of its keys. Each table operation rewrites `tblPr`, `grid`
//! and `rows` as separate map entries, so two collaborators who each change a table and undo in
//! crossing order can leave every entry deleted. Such a table lays out as nothing, takes no table
//! operation and blocks saving until its last undo is redone or the table is deleted.

use std::collections::{BTreeMap, HashSet};
use std::sync::atomic::Ordering;

use yrs::types::text::YChange;
use yrs::types::{Delta, EntryChange, Event, Events};
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, Transact, TransactionMut};

use crate::{EditingDoc, KIND_KEY, STORIES, TableLocator, map_string};

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

/// Whether `id` names a table cell's story: a `t{n}` segment followed by an `r{n}c{n}` one.
fn is_cell_story(id: &str) -> bool {
    let digits = |text: &str| !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit());
    let parts: Vec<&str> = id.split(':').collect();
    parts.windows(2).any(|pair| {
        pair[0].strip_prefix('t').is_some_and(digits)
            && pair[1]
                .strip_prefix('r')
                .and_then(|rest| rest.split_once('c'))
                .is_some_and(|(row, column)| {
                    digits(row) && digits(column.split('~').next().unwrap_or(""))
                })
    })
}

/// The cell stories a complete table payload's rows name, defaulted as the projection does.
fn named_cells(rows: &Any, story: &str, table_index: u32) -> Vec<String> {
    let Any::Array(rows) = rows else {
        return Vec::new();
    };
    let mut named = Vec::new();
    for (row_index, row) in rows.iter().enumerate() {
        let Any::Map(row) = row else { continue };
        let Some(Any::Array(cells)) = row.get("cells") else {
            continue;
        };
        for (cell_index, cell) in cells.iter().enumerate() {
            let Any::Map(cell) = cell else { continue };
            named.push(match cell.get("story") {
                Some(Any::String(id)) => id.to_string(),
                _ => format!("{story}:t{table_index}:r{row_index}c{cell_index}"),
            });
        }
    }
    named
}

impl EditingDoc {
    /// Every unavailable table a reader can reach, by story and ordinal: tables in the root
    /// stories and in the cells of complete tables, at any depth. A cell story no complete table
    /// names any more (its table was deleted while unavailable) is out of reach and never counts.
    pub(crate) fn unavailable_tables(&self) -> Vec<TableLocator> {
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(STORIES) else {
            return Vec::new();
        };
        let live: BTreeMap<String, TextRef> = crate::live_stories(&stories, &txn)
            .filter_map(|(id, value)| match value {
                Out::YText(text) => Some((id.to_owned(), text)),
                _ => None,
            })
            .collect();
        let mut queue: Vec<String> = live.keys().filter(|id| !is_cell_story(id)).cloned().collect();
        let mut reached: HashSet<String> = queue.iter().cloned().collect();
        let mut found = Vec::new();
        while let Some(id) = queue.pop() {
            let Some(story) = live.get(&id) else {
                continue;
            };
            let mut ordinal = 0u32;
            for diff in story.diff(&txn, YChange::identity) {
                let Out::YMap(map) = &diff.insert else {
                    continue;
                };
                if map_string(map, &txn, KIND_KEY).as_deref() != Some("table") {
                    continue;
                }
                if is_unavailable(map, &txn) {
                    found.push(TableLocator::new(id.clone(), ordinal));
                } else if let Some(Out::Any(rows)) = map.get(&txn, "rows") {
                    for cell in named_cells(&rows, &id, ordinal) {
                        if live.contains_key(&cell) && reached.insert(cell.clone()) {
                            queue.push(cell);
                        }
                    }
                }
                ordinal += 1;
            }
        }
        found.sort_by(|left, right| {
            (left.story.as_str(), left.table_index).cmp(&(right.story.as_str(), right.table_index))
        });
        if found.is_empty() {
            self.unavailable_tables_dirty
                .store(false, Ordering::Relaxed);
        }
        found
    }

    /// [`EditingDoc::unavailable_tables`], scanning only after a commit that may have changed
    /// them; otherwise empty, as the last scan found. For the worker's notice, never for a save.
    pub(crate) fn noticed_unavailable_tables(&self) -> Vec<TableLocator> {
        if !self.unavailable_tables_dirty.load(Ordering::Relaxed) {
            return Vec::new();
        }
        self.unavailable_tables()
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
