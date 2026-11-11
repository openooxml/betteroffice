//! Table embeds whose payload lost one of its keys. Each table operation rewrites `tblPr`, `grid`
//! and `rows` as separate map entries, so two collaborators who each change a table and undo in
//! crossing order can leave every entry deleted. Such a table lays out as nothing, takes no table
//! operation and blocks saving until its last undo is redone or the table is deleted.

use std::collections::{HashMap, HashSet};

use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, TextRef, Transact};

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

/// The structure revision a reach walk ran under.
pub(crate) type ReachKey = u64;

/// What layout requests have been given to lay out besides the body: note stories and the
/// relationship ids of section headers and footers.
#[derive(Default)]
pub(crate) struct LayoutRoots {
    notes: std::collections::BTreeSet<String>,
    header_footer_ids: std::collections::BTreeSet<String>,
}

/// What a pass over every live story found: for each story a complete table's cell or a block
/// control names, every story that names it, and whether any story holds an unavailable table.
/// Nothing in it depends on which stories are roots.
pub(crate) struct DocumentReach {
    parents: HashMap<String, Vec<String>>,
    met_unavailable: bool,
}

impl DocumentReach {
    /// Whether `story` is in `roots` or some chain of containers above it leads to one.
    fn under_a_root(&self, story: &str, roots: &HashSet<String>) -> bool {
        let mut stack = vec![story];
        let mut seen = HashSet::from([story]);
        while let Some(current) = stack.pop() {
            if roots.contains(current) {
                return true;
            }
            for parent in self.parents.get(current).into_iter().flatten() {
                if seen.insert(parent.as_str()) {
                    stack.push(parent);
                }
            }
        }
        false
    }
}

/// Whether a table payload already read as a map, where `has` tells whether it holds an entry,
/// lacks any of its entries.
pub(crate) fn lacks_payload(has: impl Fn(&str) -> bool) -> bool {
    !PAYLOAD_KEYS.iter().all(|key| has(key))
}

impl EditingDoc {
    /// Every unavailable table a reader can reach from the root stories the story list names:
    /// the body, headers and footers by their section references, and the notes themselves.
    /// A story only an unavailable or deleted table named is out of reach and never counts.
    /// `None` when a walk cannot finish.
    #[cfg(test)]
    pub(crate) fn unavailable_tables(&self) -> Option<Vec<TableLocator>> {
        let mut roots: Vec<String> = self.document_roots().into_iter().collect();
        roots.sort();
        self.unavailable_tables_from(&roots)
    }

    /// Whether the layout shows the live `story` (a cell without its story under a complete table
    /// counts as live: it shows as an empty paragraph). While no story holds an unavailable table
    /// every live story counts as shown, exactly as before, and nothing else is read. Otherwise
    /// the story counts as shown iff the top of its chain of containers is a root of the
    /// document ([`Self::document_roots`]), read afresh: a story only an unavailable table named,
    /// and everything below it, is hidden. The chains and the flag are kept until the stories'
    /// structure changes, so typing never rebuilds them.
    pub(crate) fn story_laid_out(&self, story: &str) -> bool {
        match crate::story_ref(&self.yrs_doc().transact(), story) {
            Ok(_) | Err(crate::EditError::CellHasNoStory(_)) => {}
            Err(_) => return false,
        }
        let reach = self.document_reach();
        !reach.met_unavailable || reach.under_a_root(story, &self.document_roots())
    }

    /// Records what a layout request was given to lay out: its note stories and the relationship
    /// ids of its sections' headers and footers, so they count as roots from then on.
    pub(crate) fn note_layout_roots(
        &self,
        notes: impl IntoIterator<Item = String>,
        header_footer_ids: impl IntoIterator<Item = String>,
    ) {
        let mut roots = self.layout_roots.lock().unwrap();
        roots.notes.extend(notes);
        roots.header_footer_ids.extend(header_footer_ids);
    }

    /// The structure revision the kept [`DocumentReach`] is valid for: it moves with table
    /// payload keys, cells, block controls' story references, and story creation and removal.
    pub(crate) fn reach_key(&self) -> ReachKey {
        self.story_revisions.lock().unwrap().structure
    }

    fn document_reach(&self) -> std::sync::Arc<DocumentReach> {
        let key = self.reach_key();
        if let Some((at, reach)) = self.laid_out_stories.lock().unwrap().as_ref()
            && *at == key
        {
            return std::sync::Arc::clone(reach);
        }
        let reach = std::sync::Arc::new(self.containers());
        *self.laid_out_stories.lock().unwrap() = Some((key, std::sync::Arc::clone(&reach)));
        reach
    }

    /// The document's roots, read from its own indexes and never from a story's id: the body,
    /// every header and footer the header/footer index knows (each section's references and
    /// every sourced one, as the story list reads them), and every note the source package or
    /// the part metadata holds or a layout request has laid out.
    fn document_roots(&self) -> HashSet<String> {
        let mut roots: HashSet<String> = {
            let txn = self.yrs_doc().transact();
            let mut views = crate::target::Views::committed(self, &txn);
            crate::structured::header_footer_index(&mut views)
                .roles
                .into_keys()
                .collect()
        };
        roots.insert("body".to_owned());
        let (notes, header_footer_ids) = {
            let layout = self.layout_roots.lock().unwrap();
            (layout.notes.clone(), layout.header_footer_ids.clone())
        };
        roots.extend(notes);
        roots.extend(header_footer_ids.iter().map(|id| self.header_footer_story(id)));
        if let Some(source) = self.source_metadata() {
            roots.extend(
                source
                    .read()
                    .stories
                    .iter()
                    .filter(|story| {
                        matches!(
                            story.kind,
                            crate::structured::StoryKind::Footnote
                                | crate::structured::StoryKind::Endnote
                        )
                    })
                    .map(|story| story.content_story(self)),
            );
        }
        if let Ok(parts) = serde_json::from_str::<serde_json::Value>(&self.part_metadata_json()) {
            for (kind, prefix) in [("footnotes", "fn"), ("endnotes", "en")] {
                for note in parts
                    .get(kind)
                    .and_then(serde_json::Value::as_array)
                    .into_iter()
                    .flatten()
                {
                    if let Some(id) = note.get("id").and_then(serde_json::Value::as_i64) {
                        roots.insert(format!("{prefix}:{id}"));
                    }
                }
            }
        }
        roots
    }

    /// One pass over every live story, with no starting roots: for each story a complete
    /// table's cell or a block control names (by the outline walk's names, a cell without its
    /// story included), the story that names it; and whether any story holds an unavailable
    /// table. A malformed entry is skipped; an unavailable table names nothing.
    fn containers(&self) -> DocumentReach {
        let txn = self.yrs_doc().transact();
        let mut parents = HashMap::new();
        let mut met_unavailable = false;
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return DocumentReach {
                parents,
                met_unavailable,
            };
        };
        let live: Vec<(String, TextRef)> = crate::live_stories(&stories, &txn)
            .filter_map(|(id, value)| match value {
                Out::YText(story) => Some((id.to_owned(), story)),
                _ => None,
            })
            .collect();
        for (story_id, story) in live {
            let mut name = |child: String| {
                parents
                    .entry(child)
                    .or_default()
                    .push(story_id.clone());
            };
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
                                    None | Some(Any::Null | Any::Undefined) => name(format!(
                                        "{story_id}:t{index}:r{row_index}c{cell_index}"
                                    )),
                                    Some(Any::String(child)) => name(child.to_string()),
                                    Some(_) => {}
                                }
                            }
                        }
                    }
                    Some("blockSdt") => match map.get(&txn, "story") {
                        None | Some(Out::Any(Any::Null | Any::Undefined)) => {
                            name(format!("{story_id}:sdt0"));
                        }
                        Some(Out::Any(Any::String(child))) => name(child.to_string()),
                        Some(_) => {}
                    },
                    _ => {}
                }
            }
        }
        DocumentReach {
            parents,
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
        embed_map(story, "table", txn)
    }

    fn embed_map(story: &str, kind: &str, txn: &yrs::TransactionMut) -> yrs::MapRef {
        crate::story_ref(txn, story)
            .unwrap()
            .diff(txn, yrs::types::text::YChange::identity)
            .into_iter()
            .find_map(|diff| match diff.insert {
                yrs::Out::YMap(map) => (crate::map_string(&map, txn, crate::KIND_KEY).as_deref()
                    == Some(kind))
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
    fn every_root_a_document_seeds_or_a_host_or_peer_adds_is_recognised_and_stays_shown() {
        use crate::stories::StoryInfoKind;
        const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
        const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const MAIN: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";
        let p = |text: &str| format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
        let band = |root: &str, text: &str| format!("<w:{root} {W}>{}</w:{root}>", p(text));
        let headers = [
            ("rIdA", "header1.xml", "hdr", "header"),
            ("rId:B", "header2.xml", "hdr", "header"),
            ("rIdC", "header3.xml", "hdr", "header"),
            ("rIdD", "header4.xml", "hdr", "header"),
            ("rIdF", "footer1.xml", "ftr", "footer"),
            ("x", "header5.xml", "hdr", "header"),
            ("x:t0:r0c0", "header6.xml", "hdr", "header"),
            ("y:sdt0", "header7.xml", "hdr", "header"),
            ("z:repeat:1", "footer2.xml", "ftr", "footer"),
        ];
        let mut parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{MAIN}.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="{MAIN}.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="{MAIN}.endnotes+xml"/><Override PartName="/word/settings.xml" ContentType="{MAIN}.settings+xml"/>{}</Types>"#,
                    headers
                        .iter()
                        .map(|(_, target, _, kind)| format!(
                            r#"<Override PartName="/word/{target}" ContentType="{MAIN}.{kind}+xml"/>"#
                        ))
                        .collect::<String>()
                ),
            ),
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
                ),
            ),
            (
                "word/_rels/document.xml.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFn" Type="{REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEn" Type="{REL}/endnotes" Target="endnotes.xml"/><Relationship Id="rIdSet" Type="{REL}/settings" Target="settings.xml"/>{}</Relationships>"#,
                    headers
                        .iter()
                        .map(|(id, target, _, kind)| format!(
                            r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#
                        ))
                        .collect::<String>()
                ),
            ),
            (
                "word/settings.xml".to_owned(),
                format!("<w:settings {W}><w:evenAndOddHeaders/></w:settings>"),
            ),
            (
                "word/footnotes.xml".to_owned(),
                format!(
                    r#"<w:footnotes {W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1">{}</w:footnote></w:footnotes>"#,
                    p("Foot")
                ),
            ),
            (
                "word/endnotes.xml".to_owned(),
                format!(
                    r#"<w:endnotes {W}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2">{}</w:endnote></w:endnotes>"#,
                    p("End")
                ),
            ),
            (
                "word/document.xml".to_owned(),
                format!(
                    r#"<w:document {W}><w:body>{}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdA"/><w:headerReference w:type="first" r:id="rId:B"/><w:titlePg/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="x"/><w:headerReference w:type="even" r:id="x:t0:r0c0"/><w:headerReference w:type="first" r:id="y:sdt0"/><w:footerReference w:type="default" r:id="z:repeat:1"/><w:titlePg/></w:sectPr></w:pPr><w:r><w:t>Three</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:headerReference w:type="even" r:id="rIdC"/><w:headerReference w:type="default" r:id="rIdD"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr></w:body></w:document>"#,
                    p("Before"),
                    p("Cell")
                ),
            ),
        ];
        for (_, target, root, _) in headers {
            parts.push((format!("word/{target}"), band(root, target)));
        }
        let parts: Vec<(String, Vec<u8>)> = parts
            .into_iter()
            .map(|(path, xml)| (path, xml.into_bytes()))
            .collect();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let doc = EditingDoc::new(87);
        crate::seed::seed_from_docx(&doc, &bytes).unwrap();
        let state = doc.encode_state_as_update_v1();
        let peer = EditingDoc::new(88);
        peer.apply_update_v1(&state).unwrap();
        doc.set_part_metadata(&ctx(), r#"{"body":{},"headers":[["rIdHost"]]}"#, &[])
            .unwrap();
        peer.set_part_metadata(&ctx(), r#"{"body":{},"footers":[["rIdPeer"]]}"#, &[])
            .unwrap();
        exchange(&doc, &peer);

        let roots = doc.document_roots();
        let reach = doc.document_reach();
        for info in doc.list_stories().stories {
            let (story, kind) = (info.story.as_str(), info.kind);
            match kind {
                StoryInfoKind::Body
                | StoryInfoKind::Header
                | StoryInfoKind::Footer
                | StoryInfoKind::Footnote
                | StoryInfoKind::Endnote => {
                    assert!(roots.contains(story), "{kind:?} {story} is a root by the indexes");
                }
                StoryInfoKind::TableCell | StoryInfoKind::ContentControl => {
                    assert!(reach.parents.contains_key(story), "{kind:?} {story} has a container");
                }
                StoryInfoKind::Other => {}
            }
        }
        let shown = [
            "body",
            "hf:rIdA",
            "hf:rId:B",
            "hf:rIdC",
            "hf:rIdD",
            "hf:rIdF",
            "hf:x",
            "hf:x:t0:r0c0",
            "hf:y:sdt0",
            "hf:z:repeat:1",
            "fn:1",
            "en:2",
        ];
        for story in shown {
            assert!(roots.contains(story), "{story} in {roots:?}");
        }

        strip(&doc, "body");
        assert_eq!(doc.unavailable_tables(), Some(vec![TableLocator::new("body", 0)]));
        for story in shown {
            assert!(doc.story_laid_out(story), "{story} stays shown beside an unavailable table");
        }
        assert!(
            !doc.story_laid_out("hf:rIdPeer"),
            "a footer no section references is not laid out"
        );
        assert!(!doc.story_laid_out("hf:rIdHost"));
        let key = doc.reach_key();
        doc.set_part_metadata(
            &ctx(),
            r#"{"body":{"finalSectionProperties":{"headerReferences":[{"type":"default","rId":"rIdHost"}]}},"headers":[["rIdHost"]]}"#,
            &[],
        )
        .unwrap();
        assert_eq!(doc.reach_key(), key, "a section reference is no input of the kept chains");
        assert!(
            doc.story_laid_out("hf:rIdHost"),
            "a header a section references from now on is a root at once"
        );
    }

    #[test]
    fn every_edit_that_can_change_the_reach_changes_its_key() {
        let doc = EditingDoc::new(86);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        let mut key = doc.reach_key();
        let mut moved = |what: &str| {
            let next = doc.reach_key();
            assert_ne!(next, key, "{what} must change the reach key");
            key = next;
        };
        doc.create_story("host", "Inside", "Normal", "left")
            .unwrap();
        moved("creating a story");
        doc.insert_table(&ctx(), Position::new("body", 0), 1, 1)
            .unwrap();
        moved("inserting a table");
        doc.set_column_width(&ctx(), &CellLoc::new("body", 0, 0, 0), 2000.0)
            .unwrap();
        moved("rewriting a table's payload");
        strip(&doc, "body");
        moved("removing a table's payload keys");
        let control = {
            let mut txn = doc.yrs_doc().transact_mut();
            let body = crate::story_ref(&txn, "body").unwrap();
            let control = body.insert_embed(&mut txn, 0, yrs::MapPrelim::default());
            yrs::Map::insert(&control, &mut txn, crate::KIND_KEY, "blockSdt");
            yrs::Map::insert(&control, &mut txn, "story", "host");
            control
        };
        moved("inserting a block control");
        {
            let mut txn = doc.yrs_doc().transact_mut();
            yrs::Map::insert(&control, &mut txn, "story", "elsewhere");
        }
        moved("pointing a block control at another story");
        doc.delete_story("host").unwrap();
        moved("deleting a story");
    }

    fn block_control(doc: &EditingDoc, at: &str, story: &str) {
        let mut txn = doc.yrs_doc().transact_mut();
        let text = crate::story_ref(&txn, at).unwrap();
        let control = text.insert_embed(&mut txn, 0, yrs::MapPrelim::default());
        yrs::Map::insert(&control, &mut txn, crate::KIND_KEY, "blockSdt");
        yrs::Map::insert(&control, &mut txn, "story", story);
    }

    #[test]
    fn a_root_or_a_story_one_visible_container_names_stays_shown_whatever_hidden_cells_name() {
        let doc = EditingDoc::new(89);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        doc.create_story("shared", "Shared", "Normal", "left")
            .unwrap();
        let cell = doc
            .insert_table(&ctx(), Position::new("body", 0), 1, 1)
            .unwrap()
            .created_story_ids[0]
            .clone();
        block_control(&doc, &cell, "body");
        block_control(&doc, &cell, "shared");
        block_control(&doc, "body", "shared");
        strip(&doc, "body");
        assert!(!doc.story_laid_out(&cell), "the cell of the unavailable table is hidden");
        assert!(
            doc.story_laid_out("body"),
            "the body is a root, whatever a hidden cell's block control names"
        );
        assert!(
            doc.story_laid_out("shared"),
            "a story the body's block control names is shown though a hidden cell names it too"
        );
    }

    #[test]
    fn a_note_a_layout_request_lays_out_stays_shown_beside_an_unavailable_table() {
        let engine = crate::EngineSession::new(136);
        let doc = engine.doc();
        doc.create_story("body", "See", "Normal", "left").unwrap();
        doc.insert_table(&ctx(), Position::new("body", 0), 1, 1)
            .unwrap();
        doc.create_story("fn:5", "Footnote text", "Normal", "left")
            .unwrap();
        strip(doc, "body");
        assert!(!doc.story_laid_out("fn:5"), "no index and no layout knows the note yet");
        let request = serde_json::json!({
            "bodyStory": "body",
            "regions": {"sections": [{
                "sectionId": "main",
                "pageSize": {"w": 300, "h": 200},
                "margins": {"top": 20, "right": 20, "bottom": 20, "left": 20}
            }]},
            "notes": {"contents": [{"id": 5, "height": 0}]},
            "measurement": {
                "fontChains": {},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"}
            },
            "renderEnv": {}
        });
        let _ = engine.layout_document_with_regions_json(&request.to_string());
        assert!(doc.story_laid_out("fn:5"), "the note layout was given is a root");
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
        let doc = EditingDoc::new(80);
        doc.create_story("body", "Hello", "Normal", "left")
            .unwrap();
        doc.set_part_metadata(&ctx(), r#"{"body":{},"footnotes":[{"id":1}]}"#, &[])
            .unwrap();
        assert!(doc.document_roots().contains("fn:1"));
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
