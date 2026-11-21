//! Table embeds whose payload lost one of its keys. Each table operation rewrites `tblPr`, `grid`
//! and `rows` as separate map entries, so two collaborators who each change a table and undo in
//! crossing order can leave every entry deleted. Such a table lays out as nothing, takes no table
//! operation and blocks saving until its last undo is redone or the table is deleted.

use yrs::{Map, MapRef, Out, ReadTxn};

use crate::{EditingDoc, TableLocator};

/// The refusal text of a table operation or save over an unavailable table.
pub const TABLE_UNAVAILABLE: &str = "a table is unavailable after a concurrent undo: redo on the collaborator who undid, or delete the table";

const PAYLOAD_KEYS: [&str; 3] = ["tblPr", "grid", "rows"];

/// A table embed's payload entries, each read once.
pub(crate) struct TablePayload {
    pub(crate) tbl_pr: Out,
    pub(crate) grid: Out,
    pub(crate) rows: Out,
}

/// The payload of the table embed `table`, or `None` when it lacks any of its entries: the
/// table is unavailable. Readers take the entries from here rather than reading them again.
pub(crate) fn table_payload<T: ReadTxn>(table: &MapRef, txn: &T) -> Option<TablePayload> {
    #[cfg(test)]
    PAYLOAD_READS.with(|reads| reads.set(reads.get() + 1));
    Some(TablePayload {
        tbl_pr: table.get(txn, "tblPr")?,
        grid: table.get(txn, "grid")?,
        rows: table.get(txn, "rows")?,
    })
}

#[cfg(test)]
thread_local! {
    /// Table payloads read on this thread.
    static PAYLOAD_READS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// Whether a table payload already read as a map, where `has` tells whether it holds an entry,
/// lacks any of its entries.
pub(crate) fn lacks_payload(has: impl Fn(&str) -> bool) -> bool {
    !PAYLOAD_KEYS.iter().all(|key| has(key))
}

/// Whether `story` is a footnote or endnote itself, `fn:<id>` or `en:<id>`, rather than a story
/// below one (a cell `fn:1:t0:r0c0` reads as a note once the table naming it lost its rows).
#[cfg(test)]
fn is_note_root(story: &str) -> bool {
    story
        .strip_prefix("fn:")
        .or_else(|| story.strip_prefix("en:"))
        .is_some_and(|id| !id.is_empty() && !id.contains(':'))
}

impl EditingDoc {
    /// Every unavailable table a reader can reach from the root stories the story list names:
    /// the body, headers and footers by their section references, and the notes themselves.
    /// A story only an unavailable or deleted table named is out of reach and never counts.
    /// `None` when a walk cannot finish.
    #[cfg(test)]
    pub(crate) fn unavailable_tables(&self) -> Option<Vec<TableLocator>> {
        use crate::stories::StoryInfoKind;
        let roots: Vec<String> = self
            .list_stories()
            .stories
            .into_iter()
            .filter(|info| match info.kind {
                StoryInfoKind::Body | StoryInfoKind::Header | StoryInfoKind::Footer => true,
                StoryInfoKind::Footnote | StoryInfoKind::Endnote => is_note_root(&info.story),
                _ => false,
            })
            .map(|info| info.story)
            .collect();
        self.unavailable_tables_from(&roots)
    }

    /// The unavailable tables the geometry outline walk meets from `roots`, by story and ordinal,
    /// in story then ordinal order: through the cells of complete tables and block controls'
    /// stories. A root the document lacks reads as empty. `None` when a walk cannot finish.
    pub(crate) fn unavailable_tables_from(&self, roots: &[String]) -> Option<Vec<TableLocator>> {
        let mut found = Vec::new();
        for root in roots {
            found.extend(self.unavailable_tables_under(root).ok()?);
        }
        found.sort_by(|left, right| {
            (left.story.as_str(), left.table_index).cmp(&(right.story.as_str(), right.table_index))
        });
        found.dedup();
        Some(found)
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
        let late = open(70);
        late.apply_update_v1(&a.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(late.unavailable_tables(), expected);
        assert!(undo_b.redo());
        exchange(&a, &b);
        assert_eq!(a.unavailable_tables(), none);
        assert_eq!(b.unavailable_tables(), none);
    }

    #[test]
    fn a_table_in_a_cell_counts_only_while_its_parent_table_is_complete() {
        let doc = EditingDoc::new(71);
        doc.create_story("body", "Hello", "Normal", "left").unwrap();
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
                    yrs::Out::YMap(map) => {
                        (crate::map_string(&map, &txn, crate::KIND_KEY).as_deref() == Some("table"))
                            .then_some(map)
                    }
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

    fn table_map(story: &str, txn: &yrs::TransactionMut) -> yrs::MapRef {
        crate::story_ref(txn, story)
            .unwrap()
            .diff(txn, yrs::types::text::YChange::identity)
            .into_iter()
            .find_map(|diff| match diff.insert {
                yrs::Out::YMap(map) => (crate::map_string(&map, txn, crate::KIND_KEY).as_deref()
                    == Some("table"))
                .then_some(map),
                _ => None,
            })
            .unwrap()
    }

    fn strip(doc: &EditingDoc, story: &str) {
        let mut txn = doc.yrs_doc().transact_mut();
        let table = table_map(story, &txn);
        for key in super::PAYLOAD_KEYS {
            yrs::Map::remove(&table, &mut txn, key);
        }
    }

    #[test]
    fn a_healthy_table_operation_and_outline_read_the_payload_once() {
        let doc = EditingDoc::new(81);
        doc.create_story("body", "Hello", "Normal", "left").unwrap();
        doc.insert_table(&ctx(), Position::new("body", 0), 2, 2)
            .unwrap();
        let reads = || super::PAYLOAD_READS.with(std::cell::Cell::get);
        let start = reads();
        doc.set_column_width(&ctx(), &CellLoc::new("body", 0, 0, 0), 2000.0)
            .unwrap();
        assert_eq!(reads() - start, 1, "one table operation, one payload read");
        let start = reads();
        assert_eq!(doc.unavailable_tables(), Some(Vec::new()));
        assert_eq!(
            reads() - start,
            1,
            "one table in the outline, one payload read"
        );
    }

    #[test]
    fn a_scan_that_stops_at_the_depth_limit_reports_no_result_rather_than_none_found() {
        let doc = EditingDoc::new(82);
        doc.create_story("body", "Hello", "Normal", "left").unwrap();
        let mut story = "body".to_owned();
        let mut deepest = story.clone();
        for _ in 0..260 {
            let receipt = doc
                .insert_table(&ctx(), Position::new(&story, 0), 1, 1)
                .unwrap();
            deepest = story;
            story = receipt.created_story_ids[0].clone();
        }
        let roots = ["body".to_owned()];
        assert_eq!(
            doc.unavailable_tables_from(&roots),
            None,
            "a healthy document nested past the limit cannot be vouched for"
        );
        strip(&doc, &deepest);
        assert_eq!(doc.unavailable_tables_from(&roots), None);
    }

    #[test]
    fn a_malformed_block_control_does_not_stop_the_scan() {
        let doc = EditingDoc::new(79);
        doc.create_story("body", "Hello", "Normal", "left").unwrap();
        doc.insert_table(&ctx(), Position::new("body", 0), 1, 1)
            .unwrap();
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let body = crate::story_ref(&txn, "body").unwrap();
            let control = body.insert_embed(&mut txn, 0, yrs::MapPrelim::default());
            yrs::Map::insert(&control, &mut txn, crate::KIND_KEY, "blockSdt");
            yrs::Map::insert(&control, &mut txn, "story", 42.0);
        }
        assert_eq!(doc.unavailable_tables(), Some(Vec::new()));
        strip(&doc, "body");
        assert_eq!(
            doc.unavailable_tables(),
            Some(vec![TableLocator::new("body", 0)]),
            "the table after the malformed control is still found"
        );
        doc.delete_table(&ctx(), &TableLocator::new("body", 0))
            .unwrap();
        assert_eq!(doc.unavailable_tables(), Some(Vec::new()));
    }

    #[test]
    fn a_note_cell_whose_table_lost_its_rows_is_not_a_root() {
        assert!(super::is_note_root("fn:1"));
        assert!(super::is_note_root("en:-1"));
        assert!(!super::is_note_root("fn:1:t0:r0c0"));
        assert!(!super::is_note_root("fn:"));
        let doc = EditingDoc::new(80);
        doc.create_story("body", "Hello", "Normal", "left").unwrap();
        doc.create_story("fn:1", "Note", "Normal", "left").unwrap();
        let outer = doc
            .insert_table(&ctx(), Position::new("fn:1", 0), 1, 1)
            .unwrap();
        let cell = outer.created_story_ids[0].clone();
        doc.insert_table(&ctx(), Position::new(&cell, 0), 1, 1)
            .unwrap();
        strip(&doc, &cell);
        strip(&doc, "fn:1");
        assert_eq!(
            doc.unavailable_tables(),
            Some(vec![TableLocator::new("fn:1", 0)]),
            "the inner table under the note's emptied table is out of reach, not a root"
        );
    }
}
