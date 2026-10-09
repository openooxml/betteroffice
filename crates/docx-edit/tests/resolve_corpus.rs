//! Tracked-change accept and reject corpus.
//! (`EditingDoc::accept_change` / `reject_change`), by revision id and by range.

#[path = "support/revision_boundary.rs"]
mod boundary;

use std::collections::{BTreeSet, HashMap};
use std::sync::Arc;

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{
    CellLoc, ChangeKind, ChangeTarget, EditCtx, EditingDoc, FormatPolicy, MergeDirection, OpError,
    ParaAttrDelta, ParaSelector, Patch, Position, RawOp, RichRun, SegmentContent, StoryRange,
    TableLocator, TableRange,
};
use docx_layout::types::LayoutBlock;
use yrs::Any;
use yrs::types::Attrs;

const DATE: &str = "2026-07-14T12:00:00Z";
const BLOCK_DATE: &str = "2026-07-15T12:00:00Z";

fn local() -> EditCtx {
    EditCtx::local("Owner", DATE)
}

fn suggesting(author: &str) -> EditCtx {
    EditCtx::local(author, DATE).suggesting()
}

fn seed(text: &str) -> EditingDoc {
    let doc = EditingDoc::new(300);
    doc.create_story("body", text, "Normal", "left").unwrap();
    doc
}

fn body_texts(doc: &EditingDoc) -> Vec<String> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|para| para.text)
        .collect()
}

fn change_count(doc: &EditingDoc) -> usize {
    doc.list_changes("body").unwrap().len()
}

fn body_len(doc: &EditingDoc) -> u32 {
    doc.story_len("body").unwrap()
}

fn stamp(id: &str) -> Any {
    Any::Map(Arc::new(HashMap::from([
        ("id".to_owned(), Any::from(id)),
        ("author".to_owned(), Any::from("Remote")),
        ("date".to_owned(), Any::from(DATE)),
    ])))
}

fn plain_join(doc: &EditingDoc, para: &str, operation: &str) {
    match operation {
        "delete" => doc.delete_range(&local(), StoryRange::new("body", 3, 4)),
        "merge" => doc.merge_paragraphs(&local(), para, MergeDirection::Forward),
        "retract" => doc.merge_paragraphs(&suggesting("Ada"), para, MergeDirection::Forward),
        _ => unreachable!(),
    }
    .unwrap();
}

fn boundary_target(doc: &EditingDoc, id: &str, end: u32, target: &str) -> ChangeTarget {
    match target {
        "id" => ChangeTarget::Revision(id.to_owned()),
        "range" => ChangeTarget::Range(StoryRange::new("body", 0, end)),
        "all" => ChangeTarget::Range(StoryRange::new("body", 0, body_len(doc))),
        _ => unreachable!(),
    }
}

fn assert_block_boundary(doc: &EditingDoc, kind: &str, text: Option<&str>) {
    let expected: Vec<&str> = text.into_iter().chain(["tail"]).collect();
    assert_eq!(body_texts(doc), expected);
    assert!(doc.list_revisions().unwrap().is_empty());
    let blocks = yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
    let block_index = usize::from(text.is_some());
    assert_eq!(blocks.len(), block_index + 2);
    if text.is_some() {
        assert!(matches!(&blocks[0], LayoutBlock::Paragraph(_)));
    }
    assert!(matches!(blocks.last(), Some(LayoutBlock::Paragraph(_))));
    assert!(match (&blocks[block_index], kind) {
        (LayoutBlock::Table(table), "table") => {
            let LayoutBlock::Paragraph(cell) = &table.rows[0].cells[0].blocks[0] else {
                panic!("cell paragraph missing");
            };
            serde_json::to_value(&cell.runs).unwrap()[0]["text"] == "cell"
        }
        (LayoutBlock::Paragraph(_), "blockSdt")
        | (LayoutBlock::PageBreak(_), "pageBreak")
        | (LayoutBlock::ColumnBreak(_), "columnBreak") => true,
        _ => false,
    });
}

fn assert_joined_body(doc: &EditingDoc, text: &str) {
    assert_eq!(body_texts(doc), [text]);
    assert_eq!(body_len(doc), text.encode_utf16().count() as u32 + 1);
    assert!(doc.list_revisions().unwrap().is_empty());
    let blocks = yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
    assert_eq!(blocks.len(), 1);
    assert!(matches!(&blocks[0], LayoutBlock::Paragraph(_)));
}

fn seed_pending_block(doc: &EditingDoc, kind: &str, deleted: bool) -> [String; 2] {
    doc.create_story("body", "oldtail", "Normal", "left")
        .unwrap();
    let split_ctx = if deleted { local() } else { suggesting("Ada") };
    let split = doc
        .split_paragraph(&split_ctx, Position::new("body", 3), None)
        .unwrap();
    let block_ctx = EditCtx::local("Bob", BLOCK_DATE).suggesting();
    let insert_ctx = if deleted { local() } else { block_ctx.clone() };
    let inserted = if kind == "table" {
        let table = doc
            .insert_table(&insert_ctx, Position::new("body", 4), 1, 1)
            .unwrap();
        doc.insert_text(
            &local(),
            Position::new(&table.created_story_ids[0], 0),
            "cell",
            FormatPolicy::Plain,
        )
        .unwrap();
        table.revision_ids
    } else {
        let payload = if kind == "blockSdt" {
            doc.create_story("control", "inside", "Normal", "left")
                .unwrap();
            vec![("story".to_owned(), Any::from("control"))]
        } else {
            vec![]
        };
        doc.insert_embed(&insert_ctx, Position::new("body", 4), kind, payload)
            .unwrap()
            .revision_ids
    };
    if deleted {
        let mark = doc
            .delete_range(&suggesting("Ada"), StoryRange::new("body", 3, 4))
            .unwrap();
        let block_ids = if kind == "table" {
            doc.delete_row(&block_ctx, &TableRange::cell(CellLoc::new("body", 0, 0, 0)))
                .unwrap()
                .revision_ids
        } else {
            doc.delete_range(&block_ctx, StoryRange::new("body", 4, 5))
                .unwrap()
                .revision_ids
        };
        [mark.revision_ids[0].clone(), block_ids[0].clone()]
    } else {
        [split.revision_ids[0].clone(), inserted[0].clone()]
    }
}

fn seed_pending_table_rows(doc: &EditingDoc, deleted: bool) -> [String; 3] {
    let ids = seed_pending_block(doc, "table", deleted);
    let row_ctx = EditCtx::local("Carol", BLOCK_DATE).suggesting();
    let insert_ctx = if deleted { local() } else { row_ctx.clone() };
    let inserted = doc
        .insert_row(&insert_ctx, &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    let row_ids = if deleted {
        doc.delete_row(&row_ctx, &TableRange::cell(CellLoc::new("body", 0, 1, 0)))
            .unwrap()
            .revision_ids
    } else {
        inserted.revision_ids
    };
    [ids[0].clone(), ids[1].clone(), row_ids[0].clone()]
}

fn assert_inherited_block_revision(doc: &EditingDoc, id: &str, insertion: bool) {
    assert_inherited_block_revision_from(doc, id, "Bob", insertion, 2);
    assert!(
        doc.list_revisions()
            .unwrap()
            .iter()
            .all(|revision| revision.change.revision_id == id)
    );
}

fn assert_inherited_block_revision_from(
    doc: &EditingDoc,
    id: &str,
    author: &str,
    insertion: bool,
    revision_count: usize,
) {
    let expected = Any::Map(Arc::new(HashMap::from([
        ("id".to_owned(), Any::from(id)),
        ("author".to_owned(), Any::from(author)),
        ("date".to_owned(), Any::from(BLOCK_DATE)),
    ])));
    let mark = doc
        .story_segments("body")
        .unwrap()
        .into_iter()
        .find(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
        .unwrap();
    let SegmentContent::Pilcrow(properties) = mark.content else {
        unreachable!()
    };
    let (marker, attr, other_marker, other_attr) = if insertion {
        ("pPrIns", "ins", "pPrDel", "del")
    } else {
        ("pPrDel", "del", "pPrIns", "ins")
    };
    assert_eq!(properties.values.get(marker), Some(&expected));
    assert_eq!(mark.attributes.get(attr), Some(&expected));
    assert!(!properties.values.contains_key(other_marker));
    assert!(
        mark.attributes
            .get(other_attr)
            .is_none_or(|stamp| *stamp == Any::Null)
    );
    let revisions = doc.list_revisions().unwrap();
    assert_eq!(revisions.len(), revision_count);
    let kind = if insertion {
        ChangeKind::ParagraphMarkInsertion
    } else {
        ChangeKind::ParagraphMarkDeletion
    };
    assert!(
        revisions
            .iter()
            .any(|revision| revision.change.kind == kind && revision.change.revision_id == id)
    );
}

fn assert_only_pending_block_revision(doc: &EditingDoc, id: &str) {
    let revisions = doc.list_revisions().unwrap();
    assert_eq!(revisions.len(), 1);
    let change = &revisions[0].change;
    assert_eq!(change.revision_id, id);
    assert_eq!(change.author, "Bob");
    assert_eq!(change.date, BLOCK_DATE);
    assert!(matches!(
        change.kind,
        ChangeKind::Insertion
            | ChangeKind::Deletion
            | ChangeKind::TableInsertion
            | ChangeKind::TableDeletion
            | ChangeKind::TableRowInsertion
            | ChangeKind::TableRowDeletion
    ));
}

fn assert_pending_blocks_resolve_in_either_order(accept: bool) {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for mark_first in [false, true] {
            let doc = EditingDoc::new(315);
            let expected_ids = seed_pending_block(&doc, kind, accept);
            let mut ids: Vec<String> = doc
                .list_revisions()
                .unwrap()
                .into_iter()
                .map(|revision| revision.change.revision_id)
                .collect();
            ids.dedup();
            assert_eq!(ids, expected_ids);
            if !mark_first {
                ids.reverse();
            }
            for (slot, id) in ids.iter().enumerate() {
                let target = ChangeTarget::Revision(id.clone());
                let receipt = if accept {
                    doc.accept_change(&local(), &target)
                } else {
                    doc.reject_change(&local(), &target)
                }
                .unwrap();
                assert_eq!(receipt.revision_ids, std::slice::from_ref(id));
                assert!(receipt.range.is_none());
                if slot == 0 && mark_first {
                    assert_inherited_block_revision(&doc, &expected_ids[1], !accept);
                }
            }
            assert_joined_body(&doc, "oldtail");
        }
    }
}

#[test]
fn rejecting_inserted_marks_and_blocks_joins_in_either_revision_order() {
    assert_pending_blocks_resolve_in_either_order(false);
}

#[test]
fn accepting_deleted_marks_and_blocks_joins_in_either_revision_order() {
    assert_pending_blocks_resolve_in_either_order(true);
}

fn assert_pending_table_rows_resolve_in_every_order(accept: bool) {
    for order in [
        [0, 1, 2],
        [0, 2, 1],
        [1, 0, 2],
        [1, 2, 0],
        [2, 0, 1],
        [2, 1, 0],
    ] {
        let doc = EditingDoc::new(332);
        let ids = seed_pending_table_rows(&doc, accept);
        let mut done = [false; 3];
        for slot in order {
            let target = ChangeTarget::Revision(ids[slot].clone());
            let receipt = if accept {
                doc.accept_change(&local(), &target)
            } else {
                doc.reject_change(&local(), &target)
            }
            .unwrap();
            assert_eq!(receipt.revision_ids, [ids[slot].clone()]);
            done[slot] = true;
            let revisions = doc.list_revisions().unwrap();
            assert!(revisions.iter().all(|revision| {
                !ids.iter()
                    .enumerate()
                    .any(|(slot, id)| done[slot] && revision.change.revision_id == *id)
            }));
            if done[0] && (!done[1] || !done[2]) {
                let row = if done[1] { 2 } else { 1 };
                assert_eq!(body_texts(&doc), ["old", "tail"]);
                assert_inherited_block_revision_from(
                    &doc,
                    &ids[row],
                    if row == 1 { "Bob" } else { "Carol" },
                    !accept,
                    1 + usize::from(!done[1]) + usize::from(!done[2]),
                );
            }
        }
        assert_joined_body(&doc, "oldtail");
    }
}

#[test]
fn accepting_deleted_marks_and_distinct_table_rows_joins_in_every_order() {
    assert_pending_table_rows_resolve_in_every_order(true);
}

#[test]
fn rejecting_inserted_marks_and_distinct_table_rows_joins_in_every_order() {
    assert_pending_table_rows_resolve_in_every_order(false);
}

#[test]
fn original_table_rows_do_not_lend_revisions_to_retained_marks() {
    for accept in [false, true] {
        let doc = EditingDoc::new(333);
        let ids = seed_pending_block(&doc, "table", accept);
        doc.insert_row(&local(), &CellLoc::new("body", 0, 0, 0), true)
            .unwrap();
        for id in &ids {
            let target = ChangeTarget::Revision(id.clone());
            if accept {
                doc.accept_change(&local(), &target)
            } else {
                doc.reject_change(&local(), &target)
            }
            .unwrap();
            assert!(doc.paragraphs("body").unwrap().iter().all(|paragraph| {
                !paragraph.properties.contains_key("pPrIns")
                    && !paragraph.properties.contains_key("pPrDel")
            }));
        }
        assert_eq!(body_texts(&doc), ["old", "tail"]);
        assert!(doc.list_revisions().unwrap().is_empty());
        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let LayoutBlock::Table(table) = &blocks[1] else {
            panic!("table missing");
        };
        assert_eq!(table.rows.len(), 1);
    }
}

#[test]
fn retained_marks_follow_pending_blocks_when_the_blocks_are_kept() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for accept_mark in [false, true] {
            let doc = EditingDoc::new(316);
            let ids = seed_pending_block(&doc, kind, accept_mark);
            let target = ChangeTarget::Revision(ids[0].clone());
            if accept_mark {
                doc.accept_change(&local(), &target).unwrap();
            } else {
                doc.reject_change(&local(), &target).unwrap();
            }
            assert_inherited_block_revision(&doc, &ids[1], !accept_mark);
            let target = ChangeTarget::Revision(ids[1].clone());
            let receipt = if accept_mark {
                doc.reject_change(&local(), &target)
            } else {
                doc.accept_change(&local(), &target)
            }
            .unwrap();
            assert_eq!(receipt.revision_ids, [ids[1].clone()]);
            assert_block_boundary(&doc, kind, Some("old"));
        }
    }
}

#[test]
fn resolving_pending_block_ranges_and_all_changes_still_joins() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for accept in [false, true] {
            for target in ["range", "all", "mark"] {
                let doc = EditingDoc::new(317);
                let ids = seed_pending_block(&doc, kind, accept);
                let expected_ids: BTreeSet<String> = if target == "mark" {
                    [ids[1].clone()].into_iter().collect()
                } else {
                    ids.iter().cloned().collect()
                };
                if target == "mark" {
                    let target = ChangeTarget::Range(StoryRange::new("body", 3, 4));
                    let receipt = if accept {
                        doc.accept_change(&local(), &target)
                    } else {
                        doc.reject_change(&local(), &target)
                    }
                    .unwrap();
                    assert_eq!(receipt.revision_ids, [ids[0].clone()]);
                    assert_eq!(
                        receipt.range.unwrap(),
                        doc.loc_range_of(&StoryRange::new("body", 3, 4)).unwrap()
                    );
                    assert_inherited_block_revision(&doc, &ids[1], !accept);
                }
                let expected_range =
                    StoryRange::new("body", 0, if target == "all" { 8 } else { 3 });
                let target = boundary_target(
                    &doc,
                    &ids[1],
                    5,
                    if target == "mark" { "range" } else { target },
                );
                let receipt = if accept {
                    doc.accept_change(&local(), &target)
                } else {
                    doc.reject_change(&local(), &target)
                }
                .unwrap();
                let resolved: BTreeSet<String> = receipt.revision_ids.into_iter().collect();
                assert_eq!(resolved, expected_ids);
                assert_eq!(
                    receipt.range.unwrap(),
                    doc.loc_range_of(&expected_range).unwrap()
                );
                assert_joined_body(&doc, "oldtail");
            }
        }
    }
}

#[test]
fn plain_edits_and_split_retraction_remove_marks_before_pending_blocks() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for deleted_block in [false, true] {
            for operation in ["delete", "replace", "rich", "merge", "retract"] {
                if deleted_block && operation == "retract" {
                    continue;
                }
                let doc = EditingDoc::new(318);
                let ids = seed_pending_block(&doc, kind, deleted_block);
                if operation != "retract" {
                    let target = ChangeTarget::Revision(ids[0].clone());
                    if deleted_block {
                        doc.reject_change(&local(), &target).unwrap();
                    } else {
                        doc.accept_change(&local(), &target).unwrap();
                    }
                }
                let para = doc.paragraphs("body").unwrap()[0].para_id.clone();
                doc.set_paragraph_attr(&para, "alignment", Any::from("right"))
                    .unwrap();
                let receipt = match operation {
                    "delete" => doc.delete_range(&local(), StoryRange::new("body", 3, 4)),
                    "replace" => doc.replace_range(&local(), StoryRange::new("body", 0, 4), "X"),
                    "rich" => doc.replace_range_rich(
                        &local(),
                        StoryRange::new("body", 0, 4),
                        &[RichRun {
                            text: "X".to_owned(),
                            attrs: Default::default(),
                        }],
                    ),
                    "merge" => doc.merge_paragraphs(&local(), &para, MergeDirection::Forward),
                    "retract" => {
                        doc.merge_paragraphs(&suggesting("Ada"), &para, MergeDirection::Forward)
                    }
                    _ => unreachable!(),
                }
                .unwrap();
                assert_eq!(
                    receipt.revision_ids,
                    if operation == "retract" {
                        vec![ids[0].clone()]
                    } else {
                        vec![]
                    }
                );
                let text = if matches!(operation, "replace" | "rich") {
                    "Xtail"
                } else {
                    "oldtail"
                };
                assert_eq!(body_texts(&doc), [text]);
                assert_eq!(body_len(&doc), text.len() as u32 + 2);
                let paragraph = doc.paragraphs("body").unwrap().remove(0);
                assert_eq!(paragraph.para_id, para);
                assert_eq!(
                    paragraph.properties.get("alignment"),
                    Some(&Any::from("right"))
                );
                assert_only_pending_block_revision(&doc, &ids[1]);
                let target = ChangeTarget::Revision(ids[1].clone());
                if deleted_block {
                    doc.accept_change(&local(), &target).unwrap();
                } else {
                    doc.reject_change(&local(), &target).unwrap();
                }
                assert_joined_body(&doc, text);
                let paragraph = doc.paragraphs("body").unwrap().remove(0);
                assert_eq!(paragraph.para_id, para);
                assert_eq!(
                    paragraph.properties.get("alignment"),
                    Some(&Any::from("right"))
                );
            }
        }
    }
}

#[test]
fn plain_joins_preserve_the_first_paragraph_in_either_order() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for deleted in [false, true] {
            for operation in ["delete", "merge", "retract"] {
                if deleted && operation == "retract" {
                    continue;
                }
                let mut outcomes = Vec::new();
                for block_first in [false, true] {
                    let doc = EditingDoc::new(320);
                    let ids = seed_pending_block(&doc, kind, deleted);
                    if operation != "retract" {
                        let target = ChangeTarget::Revision(ids[0].clone());
                        if deleted {
                            doc.reject_change(&local(), &target).unwrap();
                        } else {
                            doc.accept_change(&local(), &target).unwrap();
                        }
                    }
                    let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
                    doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
                        .unwrap();
                    let target = ChangeTarget::Revision(ids[1].clone());
                    if block_first {
                        if deleted {
                            doc.accept_change(&local(), &target).unwrap();
                        } else {
                            doc.reject_change(&local(), &target).unwrap();
                        }
                    }
                    plain_join(&doc, &first, operation);
                    if !block_first {
                        assert_eq!(body_texts(&doc), ["oldtail"]);
                        assert_only_pending_block_revision(&doc, &ids[1]);
                        if deleted {
                            doc.accept_change(&local(), &target).unwrap();
                        } else {
                            doc.reject_change(&local(), &target).unwrap();
                        }
                    }
                    assert_joined_body(&doc, "oldtail");
                    let paragraph = doc.paragraphs("body").unwrap().remove(0);
                    assert_eq!(paragraph.para_id, first);
                    assert_eq!(
                        paragraph.properties.get("alignment"),
                        Some(&Any::from("right"))
                    );
                    assert!(!paragraph.properties.contains_key("pPrIns"));
                    assert!(!paragraph.properties.contains_key("pPrDel"));
                    outcomes.push(paragraph);
                }
                assert_eq!(outcomes[0], outcomes[1]);
            }
        }
    }
}

#[test]
fn plain_joins_before_pending_tables_keep_first_properties_after_further_edits() {
    for sequence in ["erase", "merge"] {
        let mut outcomes = Vec::new();
        for block_first in [false, true] {
            let doc = seed(if sequence == "erase" {
                "oldtail"
            } else {
                "headoldtail"
            });
            if sequence == "erase" {
                doc.split_paragraph(&local(), Position::new("body", 3), None)
                    .unwrap();
            } else {
                doc.split_paragraph(&local(), Position::new("body", 4), None)
                    .unwrap();
                doc.split_paragraph(&local(), Position::new("body", 8), None)
                    .unwrap();
            }
            let paragraphs = doc.paragraphs("body").unwrap();
            let first = paragraphs[0].para_id.clone();
            let old = &paragraphs[paragraphs.len() - 2].para_id;
            doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
                .unwrap();
            let block_at = doc.paragraph_mark_position(old).unwrap().index + 1;
            let block = doc
                .insert_table(
                    &EditCtx::local("Bob", BLOCK_DATE).suggesting(),
                    Position::new("body", block_at),
                    1,
                    1,
                )
                .unwrap();
            let target = ChangeTarget::Revision(block.revision_ids[0].clone());
            if sequence == "erase" {
                plain_join(&doc, &first, "delete");
                doc.delete_range(&local(), StoryRange::new("body", 0, 3))
                    .unwrap();
                assert_only_pending_block_revision(&doc, &block.revision_ids[0]);
                let insertion = doc
                    .insert_text(
                        &suggesting("Carol"),
                        doc.paragraph_mark_position(&first).unwrap(),
                        "!",
                        FormatPolicy::Plain,
                    )
                    .unwrap();
                let insertion = ChangeTarget::Revision(insertion.revision_ids[0].clone());
                if block_first {
                    doc.reject_change(&local(), &target).unwrap();
                    doc.accept_change(&local(), &insertion).unwrap();
                } else {
                    doc.accept_change(&local(), &insertion).unwrap();
                    doc.reject_change(&local(), &target).unwrap();
                }
            } else {
                doc.merge_paragraphs(&local(), old, MergeDirection::Forward)
                    .unwrap();
                if block_first {
                    doc.reject_change(&local(), &target).unwrap();
                }
                doc.merge_paragraphs(&local(), &first, MergeDirection::Forward)
                    .unwrap();
                if !block_first {
                    doc.reject_change(&local(), &target).unwrap();
                }
            }
            assert_joined_body(
                &doc,
                if sequence == "erase" {
                    "tail!"
                } else {
                    "headoldtail"
                },
            );
            let paragraph = doc.paragraphs("body").unwrap().remove(0);
            assert_eq!(paragraph.para_id, first);
            assert_eq!(
                paragraph.properties.get("alignment"),
                Some(&Any::from("right"))
            );
            outcomes.push(paragraph);
        }
        assert_eq!(outcomes[0], outcomes[1]);
    }
}

#[test]
fn plain_joins_preserve_the_first_paragraph_with_range_or_id_block_rejection() {
    let mut outcomes = Vec::new();
    for by_range in [false, true] {
        let doc = EditingDoc::new(325);
        let ids = seed_pending_block(&doc, "table", false);
        doc.accept_change(&local(), &ChangeTarget::Revision(ids[0].clone()))
            .unwrap();
        let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
            .unwrap();
        plain_join(&doc, &first, "delete");
        let target = if by_range {
            ChangeTarget::Range(StoryRange::new("body", 3, 4))
        } else {
            ChangeTarget::Revision(ids[1].clone())
        };
        doc.reject_change(&local(), &target).unwrap();
        assert_joined_body(&doc, "oldtail");
        let paragraph = doc.paragraphs("body").unwrap().remove(0);
        assert_eq!(paragraph.para_id, first);
        assert_eq!(
            paragraph.properties.get("alignment"),
            Some(&Any::from("right"))
        );
        outcomes.push(paragraph);
    }
    assert_eq!(outcomes[0], outcomes[1]);
}

#[test]
fn rejecting_a_block_deletion_after_a_plain_join_preserves_the_first_paragraph() {
    let mut outcomes = Vec::new();
    for reject_block_deletion in [false, true] {
        let doc = EditingDoc::new(327);
        let ids = seed_pending_block(&doc, "table", false);
        doc.accept_change(&local(), &ChangeTarget::Revision(ids[0].clone()))
            .unwrap();
        let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
            .unwrap();
        plain_join(&doc, &first, "delete");
        if reject_block_deletion {
            let deletion = doc
                .delete_range(&suggesting("Carol"), StoryRange::new("body", 3, 4))
                .unwrap();
            doc.reject_change(
                &local(),
                &ChangeTarget::Revision(deletion.revision_ids[0].clone()),
            )
            .unwrap();
            assert_only_pending_block_revision(&doc, &ids[1]);
        }
        doc.reject_change(&local(), &ChangeTarget::Revision(ids[1].clone()))
            .unwrap();
        assert_joined_body(&doc, "oldtail");
        let paragraph = doc.paragraphs("body").unwrap().remove(0);
        assert_eq!(paragraph.para_id, first);
        assert_eq!(
            paragraph.properties.get("alignment"),
            Some(&Any::from("right"))
        );
        outcomes.push(paragraph);
    }
    assert_eq!(outcomes[0], outcomes[1]);
}

#[test]
fn plain_joins_preserve_the_first_paragraph_across_multiple_tables() {
    let mut outcomes = Vec::new();
    for second_first in [false, true] {
        let doc = EditingDoc::new(326);
        let ids = seed_pending_block(&doc, "table", false);
        doc.accept_change(&local(), &ChangeTarget::Revision(ids[0].clone()))
            .unwrap();
        let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
            .unwrap();
        doc.insert_text(
            &local(),
            Position::new("body", 5),
            "mid",
            FormatPolicy::Plain,
        )
        .unwrap();
        doc.split_paragraph(&local(), Position::new("body", 8), None)
            .unwrap();
        let second = doc
            .insert_table(&suggesting("Carol"), Position::new("body", 9), 1, 1)
            .unwrap()
            .revision_ids[0]
            .clone();
        assert_ne!(ids[1], second);
        plain_join(&doc, &first, "delete");
        doc.delete_range(&local(), StoryRange::new("body", 7, 8))
            .unwrap();
        let mut blocks = [ids[1].clone(), second];
        if second_first {
            blocks.reverse();
        }
        for id in blocks {
            doc.reject_change(&local(), &ChangeTarget::Revision(id))
                .unwrap();
        }
        assert_joined_body(&doc, "oldmidtail");
        let paragraph = doc.paragraphs("body").unwrap().remove(0);
        assert_eq!(paragraph.para_id, first);
        assert_eq!(
            paragraph.properties.get("alignment"),
            Some(&Any::from("right"))
        );
        outcomes.push(paragraph);
    }
    assert_eq!(outcomes[0], outcomes[1]);
}

#[test]
fn deferred_resolution_joins_keep_the_surviving_paragraph_properties() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        let doc = EditingDoc::new(321);
        let ids = seed_pending_block(&doc, kind, false);
        let paragraphs = doc.paragraphs("body").unwrap();
        let survivor = paragraphs[1].para_id.clone();
        doc.set_paragraph_attr(&paragraphs[0].para_id, "alignment", Any::from("right"))
            .unwrap();
        doc.reject_change(&local(), &ChangeTarget::Revision(ids[0].clone()))
            .unwrap();
        assert_inherited_block_revision(&doc, &ids[1], true);
        doc.reject_change(&local(), &ChangeTarget::Revision(ids[1].clone()))
            .unwrap();
        assert_joined_body(&doc, "oldtail");
        let paragraph = doc.paragraphs("body").unwrap().remove(0);
        assert_eq!(paragraph.para_id, survivor);
        assert_eq!(
            paragraph.properties.get("alignment"),
            Some(&Any::from("left"))
        );
    }
}

#[test]
fn rejecting_a_pending_table_preserves_an_unrelated_plain_split() {
    let doc = EditingDoc::new(329);
    let ids = seed_pending_block(&doc, "table", false);
    doc.accept_change(&local(), &ChangeTarget::Revision(ids[0].clone()))
        .unwrap();
    let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
        .unwrap();
    plain_join(&doc, &first, "delete");
    let split = doc
        .split_paragraph(&local(), Position::new("body", 1), None)
        .unwrap();
    doc.reject_change(&local(), &ChangeTarget::Revision(ids[1].clone()))
        .unwrap();
    assert_eq!(body_texts(&doc), ["o", "ldtail"]);
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs[0].para_id, first);
    assert_eq!(paragraphs[1].para_id, split.second_para_id);
    for paragraph in &paragraphs {
        assert_eq!(
            paragraph.properties.get("alignment"),
            Some(&Any::from("right"))
        );
        assert!(!paragraph.properties.contains_key("pPrIns"));
        assert!(!paragraph.properties.contains_key("pPrDel"));
    }
    assert!(doc.list_revisions().unwrap().is_empty());
}

#[test]
fn accepting_a_pending_table_deletion_keeps_a_suggested_split() {
    let doc = EditingDoc::new(330);
    let ids = seed_pending_block(&doc, "table", true);
    doc.reject_change(&local(), &ChangeTarget::Revision(ids[0].clone()))
        .unwrap();
    let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
    plain_join(&doc, &first, "delete");
    doc.split_paragraph(&suggesting("Cy"), Position::new("body", 1), None)
        .unwrap();
    doc.accept_change(&local(), &ChangeTarget::Revision(ids[1].clone()))
        .unwrap();
    assert_eq!(body_texts(&doc), ["o", "ldtail"]);
    let revisions = doc.list_revisions().unwrap();
    assert_eq!(revisions.len(), 1);
    doc.reject_change(
        &local(),
        &ChangeTarget::Revision(revisions[0].change.revision_id.clone()),
    )
    .unwrap();
    assert_eq!(body_texts(&doc), ["oldtail"]);
    assert!(doc.list_revisions().unwrap().is_empty());
}

#[test]
fn splitting_a_retained_mark_leaves_the_block_revision_on_that_mark() {
    for accept in [false, true] {
        for suggest in [false, true] {
            for (at, expected) in [(1, ["o", "ldtail"]), (3, ["old", "tail"])] {
                let doc = EditingDoc::new(331);
                let ids = seed_pending_block(&doc, "table", accept);
                let resolve = |id: &String| {
                    let target = ChangeTarget::Revision(id.clone());
                    if accept {
                        doc.accept_change(&local(), &target)
                    } else {
                        doc.reject_change(&local(), &target)
                    }
                    .unwrap();
                };
                resolve(&ids[0]);
                assert_inherited_block_revision(&doc, &ids[1], !accept);
                let ctx = if suggest { suggesting("Cy") } else { local() };
                doc.split_paragraph(&ctx, Position::new("body", at), None)
                    .unwrap();
                resolve(&ids[1]);
                assert_eq!(body_texts(&doc), expected);
                assert_eq!(doc.list_revisions().unwrap().len(), usize::from(suggest));
            }
        }
    }
}

#[test]
fn splitting_a_retained_mark_keeps_distinct_row_revisions_on_that_mark() {
    for accept in [false, true] {
        for suggest in [false, true] {
            for (at, expected) in [(1, ["o", "ldtail"]), (3, ["old", "tail"])] {
                for rows in [[1, 2], [2, 1]] {
                    let doc = EditingDoc::new(337);
                    let ids = seed_pending_table_rows(&doc, accept);
                    let resolve = |slot: usize| {
                        let target = ChangeTarget::Revision(ids[slot].clone());
                        if accept {
                            doc.accept_change(&local(), &target)
                        } else {
                            doc.reject_change(&local(), &target)
                        }
                        .unwrap();
                    };
                    resolve(0);
                    assert_inherited_block_revision_from(&doc, &ids[1], "Bob", !accept, 3);
                    let ctx = if suggest { suggesting("Cy") } else { local() };
                    doc.split_paragraph(&ctx, Position::new("body", at), None)
                        .unwrap();
                    for row in rows {
                        resolve(row);
                    }
                    assert_eq!(body_texts(&doc), expected);
                    assert_eq!(doc.list_revisions().unwrap().len(), usize::from(suggest));
                }
            }
        }
    }
}

#[test]
fn plain_joins_before_distinct_pending_rows_keep_the_first_paragraph() {
    for deleted in [false, true] {
        for operation in ["delete", "merge", "retract"] {
            if deleted && operation == "retract" {
                continue;
            }
            let doc = EditingDoc::new(338);
            let ids = seed_pending_table_rows(&doc, deleted);
            if operation != "retract" {
                let target = ChangeTarget::Revision(ids[0].clone());
                if deleted {
                    doc.reject_change(&local(), &target)
                } else {
                    doc.accept_change(&local(), &target)
                }
                .unwrap();
            }
            let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
            doc.set_paragraph_attr(&first, "alignment", Any::from("right"))
                .unwrap();
            plain_join(&doc, &first, operation);
            assert_eq!(body_texts(&doc), ["oldtail"]);
            for id in &ids[1..] {
                let target = ChangeTarget::Revision(id.clone());
                if deleted {
                    doc.accept_change(&local(), &target)
                } else {
                    doc.reject_change(&local(), &target)
                }
                .unwrap();
            }
            assert_joined_body(&doc, "oldtail");
            let paragraph = doc.paragraphs("body").unwrap().remove(0);
            assert_eq!(paragraph.para_id, first);
            assert_eq!(
                paragraph.properties.get("alignment"),
                Some(&Any::from("right"))
            );
        }
    }
}

#[test]
fn inherited_block_stamps_and_resolution_undo_and_redo_together() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for accept in [false, true] {
            let doc = EditingDoc::new(319);
            let ids = seed_pending_block(&doc, kind, accept);
            let original = doc.list_revisions().unwrap();
            let mut undo = doc.undo_manager();
            for id in &ids {
                let target = ChangeTarget::Revision(id.clone());
                if accept {
                    doc.accept_change(&local(), &target).unwrap();
                } else {
                    doc.reject_change(&local(), &target).unwrap();
                }
                undo.add_undo_barrier();
            }
            assert_joined_body(&doc, "oldtail");
            assert!(undo.undo());
            assert_eq!(body_texts(&doc), ["old", "tail"]);
            assert_inherited_block_revision(&doc, &ids[1], !accept);
            assert_eq!(
                yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default())
                    .unwrap()
                    .len(),
                3
            );
            assert!(undo.undo());
            assert_eq!(doc.list_revisions().unwrap(), original);
            assert!(undo.redo());
            assert_inherited_block_revision(&doc, &ids[1], !accept);
            assert!(undo.redo());
            assert_joined_body(&doc, "oldtail");
        }
    }
}

#[test]
fn accepting_a_deleted_mark_keeps_only_needed_block_boundaries() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for replacement in ["", "X"] {
            for target in ["id", "range", "all"] {
                let doc = EditingDoc::new(301);
                boundary::seed(&doc, kind, &local());
                let table = (kind == "table")
                    .then(|| doc.table_payload(&TableLocator::new("body", 0)).unwrap());
                let receipt = if replacement.is_empty() {
                    doc.delete_range(&suggesting("Ada"), StoryRange::new("body", 0, 4))
                } else {
                    doc.replace_range(
                        &suggesting("Ada"),
                        StoryRange::new("body", 0, 4),
                        replacement,
                    )
                }
                .unwrap();
                let target = boundary_target(
                    &doc,
                    &receipt.revision_ids[0],
                    4 + replacement.len() as u32,
                    target,
                );
                doc.accept_change(&local(), &target).unwrap();
                assert_block_boundary(&doc, kind, (!replacement.is_empty()).then_some(replacement));
                if let Some(table) = table {
                    assert_eq!(
                        doc.table_payload(&TableLocator::new("body", 0)).unwrap(),
                        table
                    );
                }
            }
        }
    }
}

#[test]
fn rejecting_an_inserted_mark_keeps_the_boundary_before_a_surviving_block() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for target in ["id", "range", "all"] {
            let doc = EditingDoc::new(302);
            let split = boundary::seed(&doc, kind, &suggesting("Ada"));
            let target = boundary_target(&doc, &split.revision_ids[0], 4, target);
            doc.reject_change(&local(), &target).unwrap();
            assert_block_boundary(&doc, kind, Some("old"));
        }
    }
}

#[test]
fn a_retained_mark_resolves_both_classes_of_the_same_revision() {
    for accept in [false, true] {
        for by_id in [false, true] {
            let doc = EditingDoc::new(310);
            let split = boundary::seed(&doc, "table", &suggesting("Ada"));
            let replacement = doc
                .replace_range(&suggesting("Ada"), StoryRange::new("body", 3, 4), "X")
                .unwrap();
            assert_eq!(replacement.revision_ids, split.revision_ids);
            let target = if by_id {
                ChangeTarget::Revision(split.revision_ids[0].clone())
            } else {
                ChangeTarget::Range(StoryRange::new("body", 3, 5))
            };
            if accept {
                doc.accept_change(&local(), &target).unwrap();
            } else {
                doc.reject_change(&local(), &target).unwrap();
            }
            assert_block_boundary(&doc, "table", Some(if accept { "oldX" } else { "old" }));
            let mark = doc
                .story_segments("body")
                .unwrap()
                .into_iter()
                .find(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
                .unwrap();
            for key in ["ins", "del"] {
                assert!(
                    mark.attributes
                        .get(key)
                        .is_none_or(|stamp| *stamp == Any::Null)
                );
            }
        }
    }
}

#[test]
fn retained_marks_keep_each_stamp_outside_the_id_target() {
    for accept in [false, true] {
        for attr_target in [false, true] {
            let doc = EditingDoc::new(311);
            let split = boundary::seed(&doc, "table", &local());
            let (marker, attribute) = if accept {
                ("pPrDel", "del")
            } else {
                ("pPrIns", "ins")
            };
            doc.set_paragraph_attr(&split.first_para_id, marker, stamp("map"))
                .unwrap();
            doc.apply_raw_ops(
                "body",
                vec![RawOp::Format {
                    index: 3,
                    len: 1,
                    attrs: Attrs::from([(Arc::from(attribute), stamp("attr"))]),
                }],
                &local(),
            )
            .unwrap();
            let id = if attr_target { "attr" } else { "map" };
            let target = ChangeTarget::Revision(id.to_owned());
            let receipt = if accept {
                doc.accept_change(&local(), &target)
            } else {
                doc.reject_change(&local(), &target)
            }
            .unwrap();
            assert_eq!(receipt.revision_ids, [id]);
            let mark = doc
                .story_segments("body")
                .unwrap()
                .into_iter()
                .find(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
                .unwrap();
            let SegmentContent::Pilcrow(properties) = mark.content else {
                unreachable!()
            };
            if attr_target {
                assert_eq!(properties.values.get(marker), Some(&stamp("map")));
                assert!(
                    mark.attributes
                        .get(attribute)
                        .is_none_or(|stamp| *stamp == Any::Null)
                );
            } else {
                assert!(!properties.values.contains_key(marker));
                assert_eq!(mark.attributes.get(attribute), Some(&stamp("attr")));
            }
            assert_eq!(body_texts(&doc), ["old", "tail"]);
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        }
    }
}

#[test]
fn accepting_and_rejecting_all_listed_revisions_preserves_block_boundaries() {
    for accept in [false, true] {
        let doc = EditingDoc::new(309);
        let split_ctx = if accept { local() } else { suggesting("Ada") };
        boundary::seed(&doc, "table", &split_ctx);
        if accept {
            doc.replace_range(&suggesting("Ada"), StoryRange::new("body", 0, 4), "X")
                .unwrap();
        }
        doc.insert_text(
            &suggesting("Bob"),
            Position::new("body", body_len(&doc) - 1),
            "!",
            FormatPolicy::Plain,
        )
        .unwrap();
        let revisions = doc.list_revisions().unwrap();
        assert_eq!(revisions.len(), if accept { 4 } else { 2 });
        let ids: BTreeSet<String> = revisions
            .into_iter()
            .map(|revision| revision.change.revision_id)
            .collect();
        assert_eq!(ids.len(), 2);
        for id in ids {
            let target = ChangeTarget::Revision(id);
            if accept {
                doc.accept_change(&local(), &target).unwrap();
            } else {
                doc.reject_change(&local(), &target).unwrap();
            }
        }
        let expected = if accept {
            ["X", "tail!"]
        } else {
            ["old", "tail"]
        };
        assert_eq!(body_texts(&doc), expected);
        assert!(doc.list_revisions().unwrap().is_empty());
        assert_eq!(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default())
                .unwrap()
                .len(),
            3
        );
    }
}

#[test]
fn a_mark_can_join_when_its_following_block_is_removed_in_the_same_resolve() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for target in ["id", "range", "all"] {
            let doc = EditingDoc::new(303);
            boundary::seed(&doc, kind, &local());
            let receipt = doc
                .delete_range(&suggesting("Ada"), StoryRange::new("body", 3, 5))
                .unwrap();
            let target = boundary_target(&doc, &receipt.revision_ids[0], 5, target);
            doc.accept_change(&local(), &target).unwrap();
            assert_eq!(body_texts(&doc), ["oldtail"]);
            assert!(doc.list_revisions().unwrap().is_empty());
            let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
            assert_eq!(blocks.len(), 1);
        }
    }
}

#[test]
fn rejecting_a_mark_and_its_inserted_break_still_joins_paragraphs() {
    for kind in ["pageBreak", "columnBreak"] {
        let doc = seed("oldtail");
        let split = doc
            .split_paragraph(&suggesting("Ada"), Position::new("body", 3), None)
            .unwrap();
        let inserted = doc
            .insert_embed(&suggesting("Ada"), Position::new("body", 4), kind, vec![])
            .unwrap();
        assert_eq!(inserted.revision_ids, split.revision_ids);
        doc.reject_change(
            &local(),
            &ChangeTarget::Revision(split.revision_ids[0].clone()),
        )
        .unwrap();
        assert_eq!(body_texts(&doc), ["oldtail"]);
        assert!(doc.list_revisions().unwrap().is_empty());
        assert_eq!(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default())
                .unwrap()
                .len(),
            1
        );
    }
}

#[test]
fn accepting_a_deleted_mark_between_text_paragraphs_still_joins() {
    for target in ["id", "range", "all"] {
        let doc = seed("oldtail");
        doc.split_paragraph(&local(), Position::new("body", 3), None)
            .unwrap();
        let receipt = doc
            .replace_range(&suggesting("Ada"), StoryRange::new("body", 0, 4), "X")
            .unwrap();
        let target = boundary_target(&doc, &receipt.revision_ids[0], 5, target);
        doc.accept_change(&local(), &target).unwrap();
        assert_eq!(body_texts(&doc), ["Xtail"]);
        assert!(doc.list_revisions().unwrap().is_empty());
        assert_eq!(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default())
                .unwrap()
                .len(),
            1
        );
    }
}

#[test]
fn accepting_only_the_mark_keeps_a_deleted_block_outside_the_target() {
    for by_id in [false, true] {
        let doc = EditingDoc::new(304);
        boundary::seed(&doc, "table", &local());
        let mark = doc
            .delete_range(&suggesting("Ada"), StoryRange::new("body", 3, 4))
            .unwrap();
        let block = doc
            .delete_range(&suggesting("Bob"), StoryRange::new("body", 4, 5))
            .unwrap();
        let target = if by_id {
            ChangeTarget::Revision(mark.revision_ids[0].clone())
        } else {
            ChangeTarget::Range(StoryRange::new("body", 3, 4))
        };
        doc.accept_change(&local(), &target).unwrap();
        assert_eq!(body_texts(&doc), ["old", "tail"]);
        assert_eq!(change_count(&doc), 2);
        yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        doc.accept_change(
            &local(),
            &ChangeTarget::Revision(block.revision_ids[0].clone()),
        )
        .unwrap();
        assert_joined_body(&doc, "oldtail");
    }
}

#[test]
fn removing_a_break_still_keeps_the_mark_before_the_next_table() {
    let doc = EditingDoc::new(305);
    boundary::seed(&doc, "table", &local());
    doc.insert_embed(&local(), Position::new("body", 4), "pageBreak", vec![])
        .unwrap();
    let receipt = doc
        .delete_range(&suggesting("Ada"), StoryRange::new("body", 3, 5))
        .unwrap();
    doc.accept_change(
        &local(),
        &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
    )
    .unwrap();
    assert_block_boundary(&doc, "table", Some("old"));
}

#[test]
fn removing_all_table_rows_allows_the_mark_to_join_and_keeps_range_bounds() {
    for accept in [false, true] {
        let doc = EditingDoc::new(306);
        let mark = doc
            .create_story("body", "oldtail", "Normal", "left")
            .unwrap();
        doc.split_paragraph(&local(), Position::new("body", 3), None)
            .unwrap();
        let table_ctx = if accept { local() } else { suggesting("Ada") };
        doc.insert_table(&table_ctx, Position::new("body", 4), 1, 1)
            .unwrap();
        if accept {
            doc.delete_range(&suggesting("Ada"), StoryRange::new("body", 3, 4))
                .unwrap();
            doc.delete_row(
                &suggesting("Ada"),
                &TableRange::cell(CellLoc::new("body", 0, 0, 0)),
            )
            .unwrap();
        } else {
            let revision = Any::Map(Arc::new(HashMap::from([
                ("id".to_owned(), Any::from("mark")),
                ("author".to_owned(), Any::from("Ada")),
                ("date".to_owned(), Any::from(DATE)),
            ])));
            doc.set_paragraph_attr(&mark, "pPrIns", revision).unwrap();
        }
        doc.insert_text(
            &suggesting("Bob"),
            Position::new("body", 5),
            "Z",
            FormatPolicy::Plain,
        )
        .unwrap();
        let target = ChangeTarget::Range(StoryRange::new("body", 3, 5));
        if accept {
            doc.accept_change(&local(), &target).unwrap();
        } else {
            doc.reject_change(&local(), &target).unwrap();
        }
        assert_eq!(body_texts(&doc), ["oldZtail"]);
        assert_eq!(change_count(&doc), 1);
        assert_eq!(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default())
                .unwrap()
                .len(),
            1
        );
    }
}

#[test]
fn plain_deletion_and_merges_preserve_boundaries_before_blocks() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        let doc = EditingDoc::new(307);
        let split = boundary::seed(&doc, kind, &local());
        doc.merge_paragraphs(&local(), &split.first_para_id, MergeDirection::Forward)
            .unwrap();
        assert_block_boundary(&doc, kind, Some("old"));
        doc.delete_range(&local(), StoryRange::new("body", 0, 4))
            .unwrap();
        assert_block_boundary(&doc, kind, None);

        let doc = EditingDoc::new(308);
        let split = boundary::seed(&doc, kind, &suggesting("Ada"));
        doc.merge_paragraphs(
            &suggesting("Ada"),
            &split.first_para_id,
            MergeDirection::Forward,
        )
        .unwrap();
        assert_block_boundary(&doc, kind, Some("old"));
    }
}

#[test]
fn empty_paragraph_merges_before_blocks_still_remove_the_mark() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for own_split in [false, true] {
            let doc = EditingDoc::new(312);
            let ctx = if own_split {
                suggesting("Ada")
            } else {
                local()
            };
            let split = boundary::seed(&doc, kind, &ctx);
            doc.delete_range(&local(), StoryRange::new("body", 0, 3))
                .unwrap();
            doc.merge_paragraphs(&ctx, &split.first_para_id, MergeDirection::Forward)
                .unwrap();
            assert_block_boundary(&doc, kind, None);
        }
    }
}

#[test]
fn plain_replacements_keep_the_boundary_for_their_new_content() {
    for kind in ["table", "blockSdt", "pageBreak", "columnBreak"] {
        for rich in [false, true] {
            let doc = EditingDoc::new(313);
            boundary::seed(&doc, kind, &local());
            let range = StoryRange::new("body", 0, 4);
            if rich {
                doc.replace_range_rich(
                    &local(),
                    range,
                    &[RichRun {
                        text: "X".to_owned(),
                        attrs: Default::default(),
                    }],
                )
                .unwrap();
            } else {
                doc.replace_range(&local(), range, "X").unwrap();
            }
            assert_block_boundary(&doc, kind, Some("X"));
        }
    }
}

#[test]
fn surviving_inline_content_requires_a_boundary_before_the_table() {
    let doc = EditingDoc::new(314);
    boundary::seed(&doc, "table", &local());
    let deleted = doc
        .delete_range(&suggesting("Ada"), StoryRange::new("body", 0, 4))
        .unwrap();
    doc.insert_embed(
        &local(),
        Position::new("body", 0),
        "image",
        vec![
            ("src".to_owned(), Any::from("data:image/png;base64,AA==")),
            ("width".to_owned(), Any::Number(80.0)),
            ("height".to_owned(), Any::Number(60.0)),
        ],
    )
    .unwrap();
    doc.accept_change(
        &local(),
        &ChangeTarget::Revision(deleted.revision_ids[0].clone()),
    )
    .unwrap();
    assert_block_boundary(&doc, "table", Some(""));
    doc.delete_range(&local(), StoryRange::new("body", 0, 2))
        .unwrap();
    assert_block_boundary(&doc, "table", None);
}

#[test]
fn accept_insertion_keeps_text_and_drops_the_stamp() {
    let doc = seed("Hello");
    let receipt = doc
        .insert_text(
            &suggesting("Ada"),
            Position::new("body", 5),
            " world",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let id = receipt.revision_ids[0].clone();
    assert_eq!(change_count(&doc), 1);

    let resolved = doc
        .accept_change(&local(), &ChangeTarget::Revision(id.clone()))
        .unwrap();
    assert_eq!(resolved.revision_ids, vec![id]);
    assert_eq!(body_texts(&doc), vec!["Hello world".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn reject_insertion_removes_the_text() {
    let doc = seed("Hello");
    let receipt = doc
        .insert_text(
            &suggesting("Ada"),
            Position::new("body", 5),
            " world",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let id = receipt.revision_ids[0].clone();

    doc.reject_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    assert_eq!(body_texts(&doc), vec!["Hello".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn accept_deletion_carries_out_the_removal() {
    let doc = seed("Hello world");
    let receipt = doc
        .delete_range(&suggesting("Ada"), StoryRange::new("body", 5, 11))
        .unwrap();
    let id = receipt.revision_ids[0].clone();
    // Suggesting mode retains the text under a `del` stamp.
    assert_eq!(body_texts(&doc), vec!["Hello world".to_owned()]);

    doc.accept_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    assert_eq!(body_texts(&doc), vec!["Hello".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn reject_deletion_restores_plain_text() {
    let doc = seed("Hello world");
    let receipt = doc
        .delete_range(&suggesting("Ada"), StoryRange::new("body", 5, 11))
        .unwrap();
    let id = receipt.revision_ids[0].clone();

    doc.reject_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    assert_eq!(body_texts(&doc), vec!["Hello world".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn accept_ppr_ins_clears_the_marker_and_keeps_the_split() {
    let doc = seed("HelloWorld");
    let split = doc
        .split_paragraph(&suggesting("Ada"), Position::new("body", 5), None)
        .unwrap();
    let id = split.revision_ids[0].clone();
    assert_eq!(change_count(&doc), 1); // one paragraph-mark insertion

    doc.accept_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    assert_eq!(
        body_texts(&doc),
        vec!["Hello".to_owned(), "World".to_owned()]
    );
    assert_eq!(change_count(&doc), 0);
    let first = &doc.paragraphs("body").unwrap()[0];
    assert_eq!(first.para_id, split.first_para_id);
    assert!(!first.properties.contains_key("pPrIns"));
}

#[test]
fn reject_ppr_ins_joins_back_and_the_second_mark_survives() {
    let doc = seed("HelloWorld");
    let split = doc
        .split_paragraph(&suggesting("Ada"), Position::new("body", 5), None)
        .unwrap();
    let id = split.revision_ids[0].clone();

    doc.reject_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    assert_eq!(body_texts(&doc), vec!["HelloWorld".to_owned()]);
    assert_eq!(change_count(&doc), 0);
    // The re-minted second-half pilcrow survives.
    assert_eq!(
        doc.paragraphs("body").unwrap()[0].para_id,
        split.second_para_id
    );
}

#[test]
fn accept_ppr_del_joins_and_the_second_paragraphs_ppr_wins() {
    let doc = seed("HelloWorld");
    let split = doc
        .split_paragraph(&local(), Position::new("body", 5), None)
        .unwrap();
    doc.set_paragraph_attr(&split.second_para_id, "alignment", Any::from("center"))
        .unwrap();
    let receipt = doc
        .merge_paragraphs(
            &suggesting("Ada"),
            &split.first_para_id,
            MergeDirection::Forward,
        )
        .unwrap();
    let id = receipt.revision_ids[0].clone();
    // The suggested merge retains the boundary mark under del + pPrDel.
    assert_eq!(body_texts(&doc).len(), 2);
    assert_eq!(change_count(&doc), 1); // one paragraph-mark deletion

    doc.accept_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(body_texts(&doc), vec!["HelloWorld".to_owned()]);
    assert_eq!(change_count(&doc), 0);
    assert_eq!(paragraphs[0].para_id, split.second_para_id);
    assert_eq!(
        paragraphs[0].properties.get("alignment"),
        Some(&Any::from("center"))
    );
}

#[test]
fn reject_ppr_del_clears_the_marker_and_keeps_the_split() {
    let doc = seed("HelloWorld");
    let split = doc
        .split_paragraph(&local(), Position::new("body", 5), None)
        .unwrap();
    let receipt = doc
        .merge_paragraphs(
            &suggesting("Ada"),
            &split.first_para_id,
            MergeDirection::Forward,
        )
        .unwrap();
    let id = receipt.revision_ids[0].clone();

    doc.reject_change(&local(), &ChangeTarget::Revision(id))
        .unwrap();
    assert_eq!(
        body_texts(&doc),
        vec!["Hello".to_owned(), "World".to_owned()]
    );
    assert_eq!(change_count(&doc), 0);
    let first = &doc.paragraphs("body").unwrap()[0];
    assert_eq!(first.para_id, split.first_para_id);
    assert!(!first.properties.contains_key("pPrDel"));
}

#[test]
fn accept_ppr_change_keeps_formatting_and_clears_the_revision() {
    let doc = seed("Hello");
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    let delta = ParaAttrDelta {
        alignment: Patch::Set("center".to_owned()),
        ..ParaAttrDelta::default()
    };
    let receipt = doc
        .set_paragraph_attrs(&suggesting("Ada"), &ParaSelector::One(para_id), &delta)
        .unwrap();
    let revision_id = receipt.revision_ids[0].clone();
    let changes = doc.list_changes("body").unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].kind, ChangeKind::ParagraphPropertiesChanged);

    doc.accept_change(&local(), &ChangeTarget::Revision(revision_id))
        .unwrap();
    let paragraph = &doc.paragraphs("body").unwrap()[0];
    assert_eq!(
        paragraph.properties.get("alignment"),
        Some(&Any::from("center"))
    );
    assert!(!paragraph.properties.contains_key("pPrChange"));
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn reject_ppr_change_restores_alignment_and_removes_added_numbering() {
    let doc = seed("Item");
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    let mut delta = ParaAttrDelta {
        alignment: Patch::Set("right".to_owned()),
        ..ParaAttrDelta::default()
    };
    delta.other.insert(
        "numPr".to_owned(),
        Some(Any::Map(Arc::new(HashMap::from([
            ("numId".to_owned(), Any::Number(2.0)),
            ("ilvl".to_owned(), Any::Number(0.0)),
        ])))),
    );
    delta
        .other
        .insert("listIsBullet".to_owned(), Some(Any::Bool(false)));
    delta
        .other
        .insert("listNumFmt".to_owned(), Some(Any::from("decimal")));
    let revision_id = doc
        .set_paragraph_attrs(&suggesting("Ada"), &ParaSelector::One(para_id), &delta)
        .unwrap()
        .revision_ids[0]
        .clone();

    doc.reject_change(&local(), &ChangeTarget::Revision(revision_id))
        .unwrap();
    let paragraph = &doc.paragraphs("body").unwrap()[0];
    assert_eq!(
        paragraph.properties.get("alignment"),
        Some(&Any::from("left"))
    );
    for key in ["numPr", "listIsBullet", "listNumFmt", "pPrChange"] {
        assert!(
            !paragraph.properties.contains_key(key),
            "{key} must be restored"
        );
    }
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn range_accept_resolves_a_suggested_replace_in_one_pass() {
    let doc = seed("Hello world");
    // Suggested replace: `del` on "world", `ins` on "there", ONE revision id.
    doc.replace_range(&suggesting("Ada"), StoryRange::new("body", 6, 11), "there")
        .unwrap();
    let len = body_len(&doc);

    doc.accept_change(
        &local(),
        &ChangeTarget::Range(StoryRange::new("body", 0, len)),
    )
    .unwrap();
    assert_eq!(body_texts(&doc), vec!["Hello there".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn range_reject_rolls_a_suggested_replace_back() {
    let doc = seed("Hello world");
    doc.replace_range(&suggesting("Ada"), StoryRange::new("body", 6, 11), "there")
        .unwrap();
    let len = body_len(&doc);

    doc.reject_change(
        &local(),
        &ChangeTarget::Range(StoryRange::new("body", 0, len)),
    )
    .unwrap();
    assert_eq!(body_texts(&doc), vec!["Hello world".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn by_id_resolution_leaves_other_revisions_untouched() {
    let doc = seed("base");
    let first = doc
        .insert_text(
            &suggesting("Ada"),
            Position::new("body", 0),
            "A",
            FormatPolicy::Inherit,
        )
        .unwrap()
        .revision_ids[0]
        .clone();
    let second = doc
        .insert_text(
            &suggesting("Bob"),
            Position::new("body", 5),
            "B",
            FormatPolicy::Inherit,
        )
        .unwrap()
        .revision_ids[0]
        .clone();

    doc.accept_change(&local(), &ChangeTarget::Revision(first))
        .unwrap();
    let remaining = doc.list_changes("body").unwrap();
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0].revision_id, second);
    assert_eq!(body_texts(&doc), vec!["AbaseB".to_owned()]);
}

#[test]
fn unknown_revision_is_a_typed_error() {
    let doc = seed("Hello");
    let missing = doc.accept_change(&local(), &ChangeTarget::Revision("nope".into()));
    assert_eq!(missing, Err(OpError::UnknownChange("nope".into())));
}

#[test]
fn a_join_on_the_final_pilcrow_clears_markers_instead_of_removing_it() {
    let doc = seed("Hello");
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    // Stamp a pPrDel directly on the story's FINAL pilcrow (unreachable through the
    // suggest ops, which never mark the final mark — but a remote peer could).
    let revision = Any::Map(Arc::new(HashMap::from([
        ("id".into(), Any::from("77:0")),
        ("author".into(), Any::from("Remote")),
        ("date".into(), Any::from(DATE)),
    ])));
    doc.set_paragraph_attr(&para_id, "pPrDel", revision)
        .unwrap();

    doc.accept_change(&local(), &ChangeTarget::Revision("77:0".into()))
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].text, "Hello");
    assert!(!paragraphs[0].properties.contains_key("pPrDel"));
}

#[test]
fn resolving_never_stamps_new_revisions_even_under_a_suggesting_ctx() {
    let doc = seed("Hello world");
    doc.delete_range(&suggesting("Ada"), StoryRange::new("body", 5, 11))
        .unwrap();
    let len = body_len(&doc);
    // Applying a resolution is not authoring: a suggesting ctx must not mint stamps.
    doc.accept_change(
        &suggesting("Bob"),
        &ChangeTarget::Range(StoryRange::new("body", 0, len)),
    )
    .unwrap();
    assert_eq!(body_texts(&doc), vec!["Hello".to_owned()]);
    assert_eq!(change_count(&doc), 0);
}

#[test]
fn tracked_image_insertion_accepts_or_rejects_as_one_embed_revision() {
    for accept in [true, false] {
        let doc = seed("");
        let receipt = doc
            .insert_embed(
                &suggesting("Ada"),
                Position::new("body", 0),
                "image",
                vec![
                    ("src".to_owned(), Any::from("data:image/png;base64,AA==")),
                    ("width".to_owned(), Any::Number(80.0)),
                    ("height".to_owned(), Any::Number(60.0)),
                ],
            )
            .unwrap();
        let revision_id = receipt.revision_ids[0].clone();
        if accept {
            doc.accept_change(&local(), &ChangeTarget::Revision(revision_id))
                .unwrap();
        } else {
            doc.reject_change(&local(), &ChangeTarget::Revision(revision_id))
                .unwrap();
        }
        let images = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter(|segment| {
                matches!(
                    segment.content,
                    SegmentContent::OtherEmbed { ref kind, .. } if kind == "image"
                )
            })
            .count();
        assert_eq!(images, usize::from(accept));
        assert_eq!(change_count(&doc), 0);
    }
}

#[test]
fn tracked_image_deletion_accepts_or_rejects_as_one_embed_revision() {
    for accept in [true, false] {
        let doc = seed("");
        doc.insert_embed(
            &local(),
            Position::new("body", 0),
            "image",
            vec![("src".to_owned(), Any::from("data:image/png;base64,AA=="))],
        )
        .unwrap();
        let revision_id = doc
            .delete_range(&suggesting("Ada"), StoryRange::new("body", 0, 1))
            .unwrap()
            .revision_ids[0]
            .clone();
        assert_eq!(
            doc.story_segments("body")
                .unwrap()
                .into_iter()
                .filter(|segment| matches!(segment.content, SegmentContent::OtherEmbed { .. }))
                .count(),
            1
        );
        if accept {
            doc.accept_change(&local(), &ChangeTarget::Revision(revision_id))
                .unwrap();
        } else {
            doc.reject_change(&local(), &ChangeTarget::Revision(revision_id))
                .unwrap();
        }
        let images = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter(|segment| matches!(segment.content, SegmentContent::OtherEmbed { .. }))
            .count();
        assert_eq!(images, usize::from(!accept));
        assert_eq!(change_count(&doc), 0);
    }
}

#[test]
fn retracting_a_kept_split_resolves_only_its_own_stamps() {
    // Ada's split then her deletion of the same mark share one revision.
    let doc = EditingDoc::new(314);
    let split = boundary::seed(&doc, "table", &suggesting("Ada"));
    doc.replace_range(&suggesting("Ada"), StoryRange::new("body", 3, 4), "")
        .unwrap();
    doc.merge_paragraphs(
        &suggesting("Ada"),
        &split.first_para_id,
        MergeDirection::Forward,
    )
    .unwrap();
    assert_eq!(body_texts(&doc), ["old", "tail"]);
    assert!(doc.list_revisions().unwrap().is_empty());

    // Another revision on the mark survives the retraction.
    let doc = EditingDoc::new(315);
    let split = boundary::seed(&doc, "table", &suggesting("Ada"));
    doc.apply_raw_ops(
        "body",
        vec![RawOp::Format {
            index: 3,
            len: 1,
            attrs: Attrs::from([(Arc::from("ins"), stamp("bob"))]),
        }],
        &local(),
    )
    .unwrap();
    doc.merge_paragraphs(
        &suggesting("Ada"),
        &split.first_para_id,
        MergeDirection::Forward,
    )
    .unwrap();
    let mark = doc
        .story_segments("body")
        .unwrap()
        .into_iter()
        .find(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
        .unwrap();
    let SegmentContent::Pilcrow(properties) = mark.content else {
        unreachable!()
    };
    assert!(!properties.values.contains_key("pPrIns"));
    assert_eq!(mark.attributes.get("ins"), Some(&stamp("bob")));
    assert_eq!(body_texts(&doc), ["old", "tail"]);
}
