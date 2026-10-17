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
    assert!(
        keys.iter().all(|key| !key.is_empty()),
        "empty paragraph key: {keys:?}"
    );
    assert_eq!(
        keys.iter().collect::<HashSet<_>>().len(),
        keys.len(),
        "duplicate paragraph keys: {keys:?}"
    );
    for key in keys {
        let at = doc.locate(&Loc::new("body", key, 0)).unwrap();
        doc.insert_text(&ctx(), at, "x", FormatPolicy::Plain)
            .unwrap();
    }
}

fn ooxml_ids(doc: &EditingDoc) -> Vec<Option<String>> {
    doc.paragraph_identities()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.ooxml_para_id)
        .collect()
}

fn sync(peers: [&EditingDoc; 3]) {
    for _ in 0..3 {
        for from in peers {
            let update = from.encode_state_as_update_v1();
            for to in peers {
                if !std::ptr::eq(from, to) {
                    to.apply_update_v1(&update).unwrap();
                }
            }
        }
    }
}

fn assert_converged(peers: [&EditingDoc; 3]) {
    sync(peers);
    for doc in &peers[1..] {
        assert_eq!(keys(doc), keys(peers[0]));
        assert_eq!(ooxml_ids(doc), ooxml_ids(peers[0]));
    }
    let ids = ooxml_ids(peers[0]);
    let saved: Vec<&String> = ids.iter().flatten().collect();
    assert_eq!(saved.iter().collect::<HashSet<_>>().len(), saved.len());
    for doc in peers {
        assert_addressable(doc);
    }
    sync(peers);
}

fn split_after_an_unseen_insert(
    receiver: &EditingDoc,
    splitter: &EditingDoc,
    writer: &EditingDoc,
    key: &str,
    offset: u32,
) {
    let at = writer.locate(&Loc::new("body", key, offset)).unwrap();
    writer
        .insert_text(&ctx(), at, "n", FormatPolicy::Plain)
        .unwrap();
    let insertion = writer
        .encode_diff_v1(&receiver.encode_state_vector_v1())
        .unwrap();
    splitter.apply_update_v1(&insertion).unwrap();
    let before_split = splitter.encode_state_vector_v1();
    let at = splitter.locate(&Loc::new("body", key, offset + 1)).unwrap();
    splitter.split_paragraph(&ctx(), at, None).unwrap();
    let split = splitter.encode_diff_v1(&before_split).unwrap();
    let paragraphs = receiver.paragraphs("body").unwrap().len();
    receiver.apply_update_v1(&split).unwrap();
    assert_eq!(receiver.paragraphs("body").unwrap().len(), paragraphs);
    assert_addressable(receiver);

    receiver.apply_update_v1(&insertion).unwrap();
    assert_eq!(receiver.paragraphs("body").unwrap().len(), paragraphs + 1);
    assert_addressable(receiver);
}

fn split_with_a_missing_text_dependency(
    receiver_client: u64,
    splitter_client: u64,
) -> [EditingDoc; 3] {
    let receiver = EditingDoc::new(receiver_client);
    receiver
        .create_story("body", "abcdef", "Normal", "left")
        .unwrap();
    let splitter = EditingDoc::new(splitter_client);
    let writer = EditingDoc::new(3);
    let initial = receiver.encode_state_as_update_v1();
    for doc in [&splitter, &writer] {
        doc.apply_update_v1(&initial).unwrap();
    }
    let key = keys(&receiver).remove(0);
    split_after_an_unseen_insert(&receiver, &splitter, &writer, &key, 3);
    assert_converged([&receiver, &splitter, &writer]);
    [receiver, splitter, writer]
}

#[test]
fn a_split_with_a_missing_text_dependency_keeps_paragraphs_addressable() {
    split_with_a_missing_text_dependency(1, 2);
}

#[test]
fn a_repair_that_outranks_the_pending_split_converges_on_every_peer() {
    split_with_a_missing_text_dependency(4, 2);
}

#[test]
fn splitting_a_repaired_paragraph_again_keeps_its_repair_key_unique() {
    let [receiver, splitter, writer] = split_with_a_missing_text_dependency(4, 2);
    let key = keys(&receiver).pop().unwrap();
    split_after_an_unseen_insert(&receiver, &splitter, &writer, &key, 1);
    assert_converged([&receiver, &splitter, &writer]);
}
