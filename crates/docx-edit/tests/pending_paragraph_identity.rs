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
    assert_eq!(keys.iter().collect::<HashSet<_>>().len(), keys.len());
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

fn split_with_a_missing_text_dependency(receiver_client: u64, splitter_client: u64) {
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
    writer
        .insert_text(&ctx(), Position::new("body", 3), "n", FormatPolicy::Plain)
        .unwrap();
    let insertion = writer
        .encode_diff_v1(&receiver.encode_state_vector_v1())
        .unwrap();
    splitter.apply_update_v1(&insertion).unwrap();
    let before_split = splitter.encode_state_vector_v1();
    splitter
        .split_paragraph(&ctx(), Position::new("body", 4), None)
        .unwrap();
    let split = splitter.encode_diff_v1(&before_split).unwrap();
    receiver.apply_update_v1(&split).unwrap();
    assert_eq!(receiver.paragraphs("body").unwrap().len(), 1);
    assert_addressable(&receiver);

    receiver.apply_update_v1(&insertion).unwrap();
    assert_eq!(receiver.paragraphs("body").unwrap().len(), 2);
    assert_addressable(&receiver);
    for _ in 0..3 {
        splitter
            .apply_update_v1(&receiver.encode_state_as_update_v1())
            .unwrap();
        receiver
            .apply_update_v1(&splitter.encode_state_as_update_v1())
            .unwrap();
    }
    assert_eq!(keys(&receiver), keys(&splitter));
    writer
        .apply_update_v1(&receiver.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(keys(&receiver), keys(&writer));
    let ids = ooxml_ids(&receiver);
    assert_eq!(ids, ooxml_ids(&splitter));
    assert_eq!(ids, ooxml_ids(&writer));
    let saved: Vec<&String> = ids.iter().flatten().collect();
    assert_eq!(saved.iter().collect::<HashSet<_>>().len(), saved.len());
    assert_addressable(&splitter);
    assert_addressable(&writer);
}

#[test]
fn a_split_with_a_missing_text_dependency_keeps_paragraphs_addressable() {
    split_with_a_missing_text_dependency(1, 2);
}

#[test]
fn a_repair_that_outranks_the_pending_split_converges_on_every_peer() {
    split_with_a_missing_text_dependency(4, 2);
}
