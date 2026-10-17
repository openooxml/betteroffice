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
    assert!(keys.iter().all(|key| !key.is_empty()), "empty paragraph key: {keys:?}");
    assert_eq!(keys.iter().collect::<HashSet<_>>().len(), keys.len());
    for key in keys {
        let at = doc.locate(&Loc::new("body", key, 0)).unwrap();
        doc.insert_text(&ctx(), at, "x", FormatPolicy::Plain)
            .unwrap();
    }
}

#[test]
fn a_split_with_a_missing_text_dependency_keeps_paragraphs_addressable() {
    let receiver = EditingDoc::new(1);
    receiver.create_story("body", "abcdef", "Normal", "left").unwrap();
    let splitter = EditingDoc::new(2);
    let writer = EditingDoc::new(3);
    let initial = receiver.encode_state_as_update_v1();
    for doc in [&splitter, &writer] {
        doc.apply_update_v1(&initial).unwrap();
    }
    writer.insert_text(&ctx(), Position::new("body", 3), "n", FormatPolicy::Plain).unwrap();
    let insertion = writer.encode_diff_v1(&receiver.encode_state_vector_v1()).unwrap();
    splitter.apply_update_v1(&insertion).unwrap();
    let before_split = splitter.encode_state_vector_v1();
    splitter.split_paragraph(&ctx(), Position::new("body", 4), None).unwrap();
    let split = splitter.encode_diff_v1(&before_split).unwrap();
    receiver.apply_update_v1(&split).unwrap();
    assert_eq!(receiver.paragraphs("body").unwrap().len(), 1);
    assert_addressable(&receiver);

    receiver.apply_update_v1(&insertion).unwrap();
    assert_eq!(receiver.paragraphs("body").unwrap().len(), 2);
    assert_addressable(&receiver);
    for _ in 0..3 {
        splitter.apply_update_v1(&receiver.encode_state_as_update_v1()).unwrap();
        receiver.apply_update_v1(&splitter.encode_state_as_update_v1()).unwrap();
    }
    assert_eq!(keys(&receiver), keys(&splitter));
    assert_addressable(&splitter);
}
