use betteroffice_xlsx::{CalculationOptions, CellRef, SheetId, Workbook};
use std::sync::mpsc;
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::Duration;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Doc, Map, ReadTxn, Transact, Update};

const SAMPLE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/sample.xlsx");
const SAMPLE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-sample.update.bin");

fn open(client_id: u64) -> Workbook {
    Workbook::open_collaborative(SAMPLE, client_id).unwrap()
}

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn input(workbook: &Workbook, address: &str) -> String {
    workbook.cell(SheetId(0), cell(address)).unwrap().input
}

fn edit(source: &mut Workbook, edits: &[(&str, &str)]) -> Vec<Vec<u8>> {
    let updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&updates);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    for (address, input) in edits {
        source
            .edit_cell(
                SheetId(0),
                cell(address),
                input,
                CalculationOptions::default(),
            )
            .unwrap();
    }
    let updates = updates.lock().unwrap().clone();
    assert_eq!(updates.len(), edits.len());
    updates
}

fn deliver(target: Workbook, updates: &[&Vec<u8>]) -> Workbook {
    let updates = updates
        .iter()
        .map(|update| update.to_vec())
        .collect::<Vec<_>>();
    let (sender, receiver) = mpsc::channel();
    thread::spawn(move || {
        let mut target = target;
        for update in &updates {
            target
                .apply_update_v1(update, CalculationOptions::default())
                .unwrap();
        }
        sender.send(target).unwrap();
    });
    receiver
        .recv_timeout(Duration::from_secs(10))
        .expect("remote updates must apply within 10 s")
}

fn merge(updates: &[&Vec<u8>]) -> Vec<u8> {
    Update::merge_updates(
        updates
            .iter()
            .map(|update| Update::decode_v1(update).unwrap()),
    )
    .encode_v1()
}

fn released_meta() -> (i64, String) {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(SAMPLE_V5).unwrap())
        .unwrap();
    let meta = doc.get_or_insert_map("xlsx");
    let txn = doc.transact();
    let version = meta.get(&txn, "schemaVersion");
    let fingerprint = meta.get(&txn, "baseFingerprint");
    (
        version.and_then(|value| value.cast().ok()).unwrap(),
        fingerprint.and_then(|value| value.cast().ok()).unwrap(),
    )
}

fn assert_same_state(left: &Workbook, right: &Workbook) {
    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(left.model(), right.model());
    assert!(
        Update::decode_v1(&left.encode_state_as_update_v1()).unwrap()
            == Update::decode_v1(&right.encode_state_as_update_v1()).unwrap()
    );
}

fn assert_projects_its_state(workbook: &Workbook, mirror_id: u64) {
    let mirror = deliver(open(mirror_id), &[&workbook.encode_state_as_update_v1()]);
    assert_eq!(mirror.model(), workbook.model());
}

#[test]
fn an_overwrite_that_overtakes_a_cleared_edit_settles_and_converges() {
    let mut source = open(821);
    let mut other = open(822);
    let updates = edit(
        &mut source,
        &[("D1", "one"), ("H20", "two"), ("H20", ""), ("D1", "three")],
    );
    let unrelated = edit(&mut other, &[("B30", "other")]);
    source
        .apply_update_v1(&unrelated[0], CalculationOptions::default())
        .unwrap();

    let target = deliver(open(823), &[&updates[0], &updates[3], &unrelated[0]]);
    assert_eq!(input(&target, "B30"), "other");
    assert_projects_its_state(&target, 824);

    let target = deliver(target, &[&updates[1], &updates[2]]);
    assert_eq!(input(&target, "D1"), "three");
    assert_eq!(input(&target, "H20"), "");
    assert_eq!(target.model(), source.model());
}

#[test]
fn an_overwrite_that_overtakes_an_edit_waits_for_it_and_converges() {
    let mut source = open(831);
    let updates = edit(
        &mut source,
        &[("D1", "one"), ("H20", "two"), ("D1", "three")],
    );

    let target = deliver(open(832), &[&updates[0], &updates[2]]);
    assert_projects_its_state(&target, 833);

    let target = deliver(target, &[&updates[1]]);
    assert_eq!(input(&target, "D1"), "three");
    assert_eq!(input(&target, "H20"), "two");
    assert_eq!(target.model(), source.model());
}

#[test]
fn merged_and_separate_delivery_of_an_overtaking_edit_agree() {
    let mut source = open(841);
    let updates = edit(
        &mut source,
        &[("D1", "one"), ("H20", "two"), ("D1", "three")],
    );

    let merged = deliver(open(842), &[&merge(&[&updates[0], &updates[2]])]);
    let separate = deliver(open(843), &[&updates[0], &updates[2]]);
    assert_eq!(input(&merged, "D1"), input(&separate, "D1"));
    assert_eq!(input(&merged, "H20"), "");
    assert_same_state(&merged, &separate);

    let merged = deliver(merged, &[&updates[1]]);
    let separate = deliver(separate, &[&updates[1]]);
    assert_eq!(input(&merged, "D1"), "three");
    assert_eq!(input(&merged, "H20"), "two");
    assert_eq!(merged.model(), source.model());
    assert_same_state(&merged, &separate);
}

#[test]
fn a_held_update_carrying_an_older_schema_from_a_higher_client_settles() {
    let mut source = open(851);
    let updates = edit(&mut source, &[("D1", "one"), ("D1", "two")]);
    let doc = Doc::with_client_id(9_000);
    doc.transact_mut()
        .apply_update(Update::decode_v1(&open(852).encode_state_as_update_v1()).unwrap())
        .unwrap();
    let before = doc.transact().state_vector();
    let (version, fingerprint) = released_meta();
    assert_eq!(version, 5);
    let meta = doc.get_or_insert_map("xlsx");
    let mut txn = doc.transact_mut();
    meta.insert(&mut txn, "schemaVersion", version);
    meta.insert(&mut txn, "baseFingerprint", fingerprint);
    drop(txn);
    let stale = doc.transact().encode_diff_v1(&before);

    let target = deliver(open(853), &[&merge(&[&stale, &updates[1]])]);
    assert_eq!(input(&target, "D1"), "");
    assert_projects_its_state(&target, 854);

    let target = deliver(target, &[&updates[0]]);
    assert_eq!(input(&target, "D1"), "two");
    let source = deliver(source, &[&stale, &target.encode_state_as_update_v1()]);
    let target = deliver(target, &[&source.encode_state_as_update_v1()]);
    assert_eq!(input(&source, "D1"), "two");
    assert_same_state(&target, &source);
}
