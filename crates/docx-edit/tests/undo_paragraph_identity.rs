use std::collections::HashSet;

use docx_edit::*;

fn ctx() -> EditCtx {
    EditCtx::local("Ada", "2026-10-16T00:00:00Z")
}

fn keys(doc: &EditingDoc) -> Vec<String> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.para_id)
        .collect()
}

fn assert_addressable(doc: &EditingDoc) {
    let keys = keys(doc);
    assert!(keys.iter().all(|key| !key.is_empty()));
    assert_eq!(
        keys.iter().collect::<HashSet<_>>().len(),
        keys.len(),
        "duplicate paragraph keys: {keys:?}"
    );
    for key in keys {
        let at = doc.locate(&Loc::new("body", key, 0)).unwrap();
        doc.insert_text(&EditCtx::system(""), at, "x", FormatPolicy::Plain)
            .unwrap();
    }
}

#[test]
fn undoing_a_split_after_a_remote_split_keeps_unique_paragraph_keys() {
    let local = EditingDoc::new(1);
    local
        .create_story("body", "abcdef", "Normal", "left")
        .unwrap();
    let remote = EditingDoc::new(2);
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    let mut undo = local.undo_manager();
    local
        .split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    remote
        .split_paragraph(&ctx(), Position::new("body", 1), None)
        .unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&local).len(), 3);
    assert!(undo.undo());
    assert_eq!(keys(&local).len(), 2);
    assert_addressable(&local);
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&local), keys(&remote));
    assert!(undo.redo());
    assert_eq!(keys(&local).len(), 3);
    assert_addressable(&local);
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&local), keys(&remote));
}

#[test]
fn a_restored_derived_key_stays_with_its_original_pilcrow_on_every_replica() {
    let local = EditingDoc::new(1);
    local
        .create_story("body", "abcdef", "Normal", "left")
        .unwrap();
    let remote = EditingDoc::new(2);
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .split_paragraph(&ctx(), Position::new("body", 2), None)
        .unwrap();
    remote
        .split_paragraph(&ctx(), Position::new("body", 4), None)
        .unwrap();
    let from_local = local.encode_state_as_update_v1();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    remote.apply_update_v1(&from_local).unwrap();
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    let middle = keys(&local)[1].clone();
    assert!(middle.contains('~'));
    let mut undo = local.undo_manager();
    local
        .merge_paragraphs(&ctx(), &middle, MergeDirection::Forward)
        .unwrap();
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    let at = remote.locate(&Loc::new("body", middle, 1)).unwrap();
    remote.split_paragraph(&ctx(), at, None).unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    assert!(undo.undo());
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    assert_addressable(&remote);
    assert_addressable(&local);
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&local), keys(&remote));
}

#[test]
fn undo_repair_keeps_a_derived_key_unique_when_a_split_already_carries_it() {
    let local = EditingDoc::new(1);
    local
        .create_story("body", "abcdef", "Normal", "left")
        .unwrap();
    let remote = EditingDoc::new(2);
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .split_paragraph(&ctx(), Position::new("body", 2), None)
        .unwrap();
    remote
        .split_paragraph(&ctx(), Position::new("body", 4), None)
        .unwrap();
    let from_local = local.encode_state_as_update_v1();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    remote.apply_update_v1(&from_local).unwrap();
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    let derived = keys(&local)[1].clone();
    assert!(derived.contains('~'));
    let at = local.locate(&Loc::new("body", derived, 1)).unwrap();
    let second = local
        .split_paragraph(&ctx(), at, None)
        .unwrap()
        .second_para_id;
    let mut undo = local.undo_manager();
    let at = local.locate(&Loc::new("body", second.clone(), 0)).unwrap();
    local.split_paragraph(&ctx(), at, None).unwrap();
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    let at = remote.locate(&Loc::new("body", second, 0)).unwrap();
    remote.split_paragraph(&ctx(), at, None).unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&local).len(), 6);
    assert!(undo.undo());
    assert_eq!(keys(&local).len(), 5);
    assert_addressable(&local);
    remote
        .apply_update_v1(&local.encode_state_as_update_v1())
        .unwrap();
    local
        .apply_update_v1(&remote.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&local), keys(&remote));
    assert_addressable(&remote);
}
