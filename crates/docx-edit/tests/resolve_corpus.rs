//! S4b accept/reject corpus: the 3-way ins/del/pPr resolve matrix
//! (`EditingDoc::accept_change` / `reject_change`), by revision id and by range.

use std::collections::HashMap;
use std::sync::Arc;

use docx_edit::{
    ChangeKind, ChangeTarget, EditCtx, EditingDoc, FormatPolicy, MergeDirection, OpError,
    ParaAttrDelta, ParaSelector, Patch, Position, SegmentContent, StoryRange,
};
use yrs::Any;

const DATE: &str = "2026-07-14T12:00:00Z";

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
    // The inserted pilcrow is gone; the surviving mark is the re-minted second half's
    // (the PM resolver's inherit-from-second join).
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
