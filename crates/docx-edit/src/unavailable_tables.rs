//! Table embeds whose payload lost one of its keys. Each table operation rewrites `tblPr`, `grid`
//! and `rows` as separate map entries, so two collaborators who each change a table and undo in
//! crossing order can leave every entry deleted. Such a table lays out as nothing, takes no table
//! operation and blocks saving until its last undo is redone or the table is deleted.

use std::collections::HashSet;

use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, Transact};

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

/// What the reachability walk from the document's roots found: the stories it reached, and
/// whether it met an unavailable table on the way.
pub(crate) struct DocumentReach {
    reached: HashSet<String>,
    met_unavailable: bool,
}

impl DocumentReach {
    /// Whether the layout shows a live story: hidden only when the walk met an unavailable table
    /// and did not reach the story.
    fn shows(&self, story: &str) -> bool {
        !self.met_unavailable || self.reached.contains(story)
    }
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

    /// Whether the layout shows the live `story` (a cell without its story under a complete table
    /// counts as live: it shows as an empty paragraph). A story is hidden only on positive
    /// evidence: the walk from the document's roots met an unavailable table and did not reach
    /// the story. Without an unavailable table in reach every live story counts as shown, exactly
    /// as before. The walk runs once per structural change, so typing never walks again, and has
    /// no depth limit.
    pub(crate) fn story_laid_out(&self, story: &str) -> bool {
        match crate::story_ref(&self.yrs_doc().transact(), story) {
            Ok(_) | Err(crate::EditError::CellHasNoStory(_)) => {}
            Err(_) => return false,
        }
        let structure = self.story_revisions.lock().unwrap().structure;
        if let Some((at, reach)) = self.laid_out_stories.lock().unwrap().as_ref()
            && *at == structure
        {
            return reach.shows(story);
        }
        let reach = self.reachable_stories(&self.document_roots());
        let laid_out = reach.shows(story);
        *self.laid_out_stories.lock().unwrap() = Some((structure, std::sync::Arc::new(reach)));
        laid_out
    }

    /// The document's roots, from its story metadata: the body, and each header and footer its
    /// sections reference. Notes are reached through their references.
    fn document_roots(&self) -> Vec<String> {
        use crate::stories::StoryInfoKind;
        self.list_stories()
            .stories
            .into_iter()
            .filter(|info| {
                matches!(
                    info.kind,
                    StoryInfoKind::Body | StoryInfoKind::Header | StoryInfoKind::Footer
                )
            })
            .map(|info| info.story)
            .collect()
    }

    /// Every live story reachable from `roots` through the cells of complete tables, block
    /// controls' stories and note references, at any depth: the stories the layout shows, named
    /// as the outline walk names them, and whether the walk met an unavailable table. A cell
    /// without its story under a complete table is reached as a leaf, as the layout shows it. A
    /// malformed entry is skipped; an unavailable table leads nowhere.
    fn reachable_stories(&self, roots: &[String]) -> DocumentReach {
        let txn = self.yrs_doc().transact();
        let mut reached = HashSet::new();
        let mut met_unavailable = false;
        let mut queue = roots.to_vec();
        while let Some(story_id) = queue.pop() {
            if reached.contains(&story_id) {
                continue;
            }
            let story = match crate::story_ref(&txn, &story_id) {
                Ok(story) => story,
                Err(crate::EditError::CellHasNoStory(_)) => {
                    reached.insert(story_id);
                    continue;
                }
                Err(_) => continue,
            };
            reached.insert(story_id.clone());
            let mut table_index = 0;
            for diff in story.diff(&txn, YChange::identity) {
                let Out::YMap(map) = diff.insert else {
                    continue;
                };
                match crate::map_string(&map, &txn, crate::KIND_KEY).as_deref() {
                    Some("table") => {
                        let index = table_index;
                        table_index += 1;
                        let Some(payload) = table_payload(&map, &txn) else {
                            met_unavailable = true;
                            continue;
                        };
                        let Out::Any(Any::Array(rows)) = payload.rows else {
                            continue;
                        };
                        for (row_index, row) in rows.iter().enumerate() {
                            let Any::Map(row) = row else { continue };
                            let Some(Any::Array(cells)) = row.get("cells") else {
                                continue;
                            };
                            for (cell_index, cell) in cells.iter().enumerate() {
                                let Any::Map(cell) = cell else { continue };
                                match cell.get("story") {
                                    None | Some(Any::Null | Any::Undefined) => queue.push(format!(
                                        "{story_id}:t{index}:r{row_index}c{cell_index}"
                                    )),
                                    Some(Any::String(child)) => queue.push(child.to_string()),
                                    Some(_) => {}
                                }
                            }
                        }
                    }
                    Some("blockSdt") => match map.get(&txn, "story") {
                        None | Some(Out::Any(Any::Null | Any::Undefined)) => {
                            queue.push(format!("{story_id}:sdt0"));
                        }
                        Some(Out::Any(Any::String(child))) => queue.push(child.to_string()),
                        Some(_) => {}
                    },
                    Some("noteRef") => {
                        for (key, prefix) in [("footnoteRefId", "fn"), ("endnoteRefId", "en")] {
                            let id = match map.get(&txn, key) {
                                Some(Out::Any(Any::Number(id))) => Some(id as i64),
                                Some(Out::Any(Any::String(id))) => id.parse::<i64>().ok(),
                                _ => None,
                            };
                            if let Some(id) = id {
                                queue.push(format!("{prefix}:{id}"));
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        DocumentReach {
            reached,
            met_unavailable,
        }
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
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        doc.insert_table(&ctx(), Position::new("body", 0), 2, 2)
            .unwrap();
        let reads = || super::PAYLOAD_READS.with(std::cell::Cell::get);
        let start = reads();
        doc.set_column_width(&ctx(), &CellLoc::new("body", 0, 0, 0), 2000.0)
            .unwrap();
        assert_eq!(reads() - start, 1, "one table operation, one payload read");
        let start = reads();
        assert_eq!(doc.unavailable_tables(), Some(Vec::new()));
        assert_eq!(reads() - start, 1, "one table in the outline, one payload read");
    }

    #[test]
    fn a_scan_that_stops_at_the_depth_limit_reports_no_result_rather_than_none_found() {
        let doc = EditingDoc::new(82);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
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
        assert!(
            doc.story_laid_out(&story),
            "a cell of a complete table is shown at any depth"
        );
        strip(&doc, &deepest);
        assert_eq!(doc.unavailable_tables_from(&roots), None);
        assert!(
            !doc.story_laid_out(&story),
            "a cell of the deepest table is hidden once that table lost its payload, past any walk limit"
        );
        assert!(doc.story_laid_out(&deepest));
    }

    #[test]
    fn a_story_nothing_reaches_is_hidden_only_while_the_document_holds_an_unavailable_table() {
        let doc = EditingDoc::new(84);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        doc.insert_table(&ctx(), Position::new("body", 0), 1, 1)
            .unwrap();
        doc.create_story("left-behind", "Leftover", "Normal", "left")
            .unwrap();
        assert!(
            doc.story_laid_out("left-behind"),
            "with no unavailable table in reach a live story counts as shown, as before"
        );
        strip(&doc, "body");
        assert!(
            !doc.story_laid_out("left-behind"),
            "once the document holds an unavailable table, a story nothing reaches counts as hidden"
        );
        assert!(doc.story_laid_out("body"));
        assert!(!doc.story_laid_out("missing"));
    }

    #[test]
    fn a_host_block_control_holding_an_unavailable_table_hides_its_cells_whatever_their_names() {
        let doc = EditingDoc::new(85);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        doc.create_story("host", "Inside", "Normal", "left")
            .unwrap();
        let cell = doc
            .insert_table(&ctx(), Position::new("host", 0), 1, 1)
            .unwrap()
            .created_story_ids[0]
            .clone();
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let body = crate::story_ref(&txn, "body").unwrap();
            let control = body.insert_embed(&mut txn, 0, yrs::MapPrelim::default());
            yrs::Map::insert(&control, &mut txn, crate::KIND_KEY, "blockSdt");
            yrs::Map::insert(&control, &mut txn, "story", "host");
        }
        assert!(doc.story_laid_out(&cell));
        strip(&doc, "host");
        assert!(
            !doc.story_laid_out(&cell),
            "the cell is hidden though its name does not lead to the body"
        );
        assert!(doc.story_laid_out("host"));
        assert!(doc.story_laid_out("body"));
    }

    #[test]
    fn a_cell_without_its_story_under_a_complete_table_stays_shown_beside_an_unavailable_table() {
        let origin = EditingDoc::new(100);
        origin
            .create_story("body", "Hello", "Normal", "left")
            .unwrap();
        origin
            .insert_table(&ctx(), Position::new("body", 0), 1, 2)
            .unwrap();
        origin
            .create_story("other", "Elsewhere", "Normal", "left")
            .unwrap();
        origin
            .insert_table(&ctx(), Position::new("other", 0), 1, 1)
            .unwrap();
        {
            let mut txn = origin.yrs_doc().transact_mut();
            let body = crate::story_ref(&txn, "body").unwrap();
            let control = body.insert_embed(&mut txn, 0, yrs::MapPrelim::default());
            yrs::Map::insert(&control, &mut txn, crate::KIND_KEY, "blockSdt");
            yrs::Map::insert(&control, &mut txn, "story", "other");
        }
        let state = origin.encode_state_as_update_v1();
        let open = |client| {
            let doc = EditingDoc::new(client);
            doc.apply_update_v1(&state).unwrap();
            doc
        };
        let (doc, peer) = (open(81), open(82));
        let mut undo = doc.undo_manager();
        doc.split_cell_grid(&ctx(), &CellLoc::new("body", 0, 0, 1), Some(1), Some(2))
            .unwrap();
        exchange(&doc, &peer);
        peer.insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
            .unwrap();
        exchange(&doc, &peer);
        assert!(undo.undo());
        exchange(&doc, &peer);
        let txn = doc.yrs_doc().transact();
        let empty: Vec<String> = crate::empty_cells::named_cell_ids(&txn)
            .into_iter()
            .filter(|story| crate::empty_cells::is_empty_cell(&txn, story))
            .collect();
        drop(txn);
        assert!(!empty.is_empty(), "the undo left a cell without its story");
        strip(&doc, "other");
        assert_eq!(
            doc.unavailable_tables(),
            Some(vec![TableLocator::new("other", 0)])
        );
        for cell in &empty {
            assert!(
                doc.story_laid_out(cell),
                "a cell without its story under a complete table shows as an empty paragraph"
            );
        }
    }

    #[test]
    fn a_malformed_block_control_does_not_stop_the_scan() {
        let doc = EditingDoc::new(79);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
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
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
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
