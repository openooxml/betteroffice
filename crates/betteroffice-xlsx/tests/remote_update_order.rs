use betteroffice_xlsx::{CalculationOptions, CellRef, SheetId, Workbook};
use std::sync::mpsc;
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::Duration;

const SAMPLE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/sample.xlsx");

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
    let updates = updates.iter().map(|update| update.to_vec()).collect::<Vec<_>>();
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
        &[
            ("D1", "one"),
            ("H20", "two"),
            ("H20", ""),
            ("D1", "three"),
        ],
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
