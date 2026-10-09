use docx_edit::{EditingDoc, seed_from_docx_with_generation, story_checksum};
use yrs::{Map, ReadTxn, Transact};

const DOCX: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.docx");

fn seeded() -> EditingDoc {
    let doc = EditingDoc::new(1);
    seed_from_docx_with_generation(&doc, DOCX, "determinism").unwrap();
    doc
}

fn checksums(doc: &EditingDoc) -> Vec<(String, u64)> {
    let txn = doc.yrs_doc().transact();
    let stories = txn.get_map("stories").unwrap();
    let mut ids: Vec<_> = stories.keys(&txn).map(|id| id.to_string()).collect();
    ids.sort();
    drop(txn);
    ids.into_iter()
        .map(|id| {
            let checksum = story_checksum(doc, &id).unwrap();
            (id, checksum)
        })
        .collect()
}

#[test]
fn repeated_seeds_encode_identically() {
    let expected = seeded().encode_state_as_update_v1();
    for _ in 0..16 {
        assert!(
            seeded().encode_state_as_update_v1() == expected,
            "identical DOCX bytes produced a different yrs update"
        );
    }
}

#[test]
fn independently_seeded_replicas_have_identical_canonical_checksums() {
    let left = seeded();
    let right = seeded();

    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(checksums(&left), checksums(&right));
}

#[test]
fn a_native_shared_seed_matches_the_committed_browser_seed() {
    let native = EditingDoc::new(1);
    seed_from_docx_with_generation(&native, DOCX, "betteroffice-demo").unwrap();
    let browser = include_bytes!("../../../apps/demo/public/seeds/docx.bin");
    assert!(native.encode_state_as_update_v1() == browser.as_slice());
    let replica = EditingDoc::new(2);
    replica.apply_update_v1(browser).unwrap();
    assert_eq!(replica.session_id(), native.session_id());
}
