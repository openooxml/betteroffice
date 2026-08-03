//! The `workbook-schema-v5-*.update.bin` fixtures were produced by release
//! 4bdccdd: it opens the matching workbook collaboratively, writes a cell, and
//! persists `encode_state_as_update_v1()`. `hidden-dimensions.xlsx` is a
//! minimal workbook with a hidden row and a hidden column, which that release
//! modelled with no dimension entry at all.

use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, SheetId, UpdateOrigin, Workbook};
use yrs::updates::decoder::Decode;
use yrs::{Doc, Map, MapRef, ReadTxn, StateVector, Transact, Update};

const SAMPLE: &[u8] = include_bytes!("../../../apps/demo/public/sample.xlsx");
const SHOWCASE: &[u8] = include_bytes!("../../../apps/demo/public/showcase.xlsx");
const HIDDEN: &[u8] = include_bytes!("fixtures/hidden-dimensions.xlsx");
const SAMPLE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-sample.update.bin");
const SHOWCASE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-showcase.update.bin");
const HIDDEN_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-hidden.update.bin");

fn a1(workbook: &Workbook) -> CellValue {
    workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(CellRef::parse_a1("A1").unwrap())
        .map(|cell| cell.value.clone())
        .unwrap_or(CellValue::Empty)
}

fn restored(source: &[u8], snapshot: &[u8], client_id: u64) -> Workbook {
    let mut workbook = Workbook::open_collaborative(source, client_id).unwrap();
    let result = workbook
        .apply_update_v1(snapshot, CalculationOptions::default())
        .unwrap();
    assert!(result.applied);
    workbook
}

#[test]
fn a_released_v5_snapshot_restores_and_keeps_editing() {
    let mut workbook = restored(SAMPLE, SAMPLE_V5, 5_001);
    assert_eq!(
        a1(&workbook),
        CellValue::Text {
            value: "persisted on v5".into()
        }
    );

    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("A2").unwrap(),
            "after restore",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(workbook.save().is_ok());

    let peer = Workbook::open_collaborative(SAMPLE, 5_002).unwrap();
    let mut peer = peer;
    peer.apply_update_v1(
        &workbook.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    assert_eq!(
        a1(&peer),
        CellValue::Text {
            value: "persisted on v5".into()
        }
    );
}

#[test]
fn a_charted_workbook_restores_a_released_v5_snapshot() {
    let workbook = restored(SHOWCASE, SHOWCASE_V5, 5_003);
    assert_eq!(
        a1(&workbook),
        CellValue::Text {
            value: "persisted on v5".into()
        }
    );
    let charted = Workbook::open(SHOWCASE).unwrap();
    for (sheet, source) in workbook.model().sheets.iter().zip(&charted.model().sheets) {
        assert_eq!(sheet.charts, source.charts);
    }
    assert!(
        charted
            .model()
            .sheets
            .iter()
            .any(|sheet| !sheet.charts.is_empty()),
        "the fixture must exercise a charted workbook"
    );
}

/// A restored snapshot is written forward at the current schema, so reopening
/// it takes the ordinary path rather than migrating a second time.
#[test]
fn a_restored_snapshot_round_trips_at_the_current_schema() {
    let workbook = restored(SAMPLE, SAMPLE_V5, 5_004);
    let current = workbook.encode_state_as_update_v1();

    let mut reopened = Workbook::open_collaborative(SAMPLE, 5_005).unwrap();
    reopened
        .apply_update_v1(&current, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        a1(&reopened),
        CellValue::Text {
            value: "persisted on v5".into()
        }
    );
    assert_eq!(reopened.model(), workbook.model());
}

/// A snapshot only supersedes a replica that has not been edited yet.
#[test]
fn an_edited_replica_does_not_adopt_a_legacy_snapshot() {
    let mut workbook = Workbook::open_collaborative(SAMPLE, 5_006).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("A1").unwrap(),
            "local",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook
            .apply_update_v1(SAMPLE_V5, CalculationOptions::default())
            .is_err()
    );
    assert_eq!(
        a1(&workbook),
        CellValue::Text {
            value: "local".into()
        }
    );
}

/// Adoption is gated on the base fingerprint, so a snapshot of some other
/// workbook is never taken as this one's state. (Merging one is a separate,
/// older problem: it lands a partial contamination either way.)
#[test]
fn a_snapshot_from_another_workbook_is_never_adopted() {
    let mut workbook = Workbook::open_collaborative(SAMPLE, 5_008).unwrap();
    let before = workbook
        .model()
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let _ = workbook.apply_update_v1(SHOWCASE_V5, CalculationOptions::default());
    assert_eq!(
        workbook
            .model()
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>(),
        before
    );
}

/// Hidden rows and columns model as a zero dimension now and as nothing at all
/// before, and both maps are fingerprinted — so the released snapshot is only
/// recognisable against the dimensions the released parser would have stored.
#[test]
fn a_workbook_with_hidden_dimensions_restores_its_released_snapshot() {
    let workbook = restored(HIDDEN, HIDDEN_V5, 5_009);
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(CellRef::parse_a1("A3").unwrap())
            .map(|cell| cell.value.clone()),
        Some(CellValue::Text {
            value: "persisted on v5".into()
        })
    );
    let sheet = workbook.sheet(SheetId(0)).unwrap();
    assert_eq!(
        sheet.row_heights.get(&1),
        Some(&0.0),
        "the hidden row must not silently unhide"
    );
    assert_eq!(sheet.col_widths.get(&2), Some(&0.0));

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    let sheet = reopened.sheet(SheetId(0)).unwrap();
    assert_eq!(sheet.row_heights.get(&1), Some(&0.0));
    assert_eq!(sheet.col_widths.get(&2), Some(&0.0));
}

/// A matching base fingerprint identifies the workbook a snapshot started from.
/// It says nothing about what the snapshot then did to the frozen structure.
#[test]
fn a_structurally_tampered_snapshot_is_refused() {
    let mut workbook = Workbook::open_collaborative(SAMPLE, 5_010).unwrap();
    let before = workbook.model().clone();
    let tampered = rename_first_sheet(SAMPLE_V5, "Tampered");
    assert!(matches!(
        workbook.apply_update_v1(&tampered, CalculationOptions::default()),
        Err(betteroffice_xlsx::Error::CollaborativeStructureChanged)
    ));
    assert_eq!(workbook.model(), &before);
}

/// Migration writes new structs under the restoring client. A peer that never
/// received them cannot integrate anything built on top, so a later
/// incremental update would sit pending forever.
#[test]
fn incremental_updates_converge_after_a_migration() {
    let broadcast: Arc<Mutex<Vec<Vec<u8>>>> = Arc::new(Mutex::new(Vec::new()));
    let mut left = Workbook::open_collaborative(SAMPLE, 5_011).unwrap();
    let sink = Arc::clone(&broadcast);
    let _subscription = left
        .observe_update_v1(move |event| {
            if event.origin == UpdateOrigin::Local {
                sink.lock().unwrap().push(event.update);
            }
        })
        .unwrap();

    left.apply_update_v1(SAMPLE_V5, CalculationOptions::default())
        .unwrap();

    let mut right = Workbook::open_collaborative(SAMPLE, 5_012).unwrap();
    for update in broadcast.lock().unwrap().drain(..) {
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap();
    }
    assert_eq!(
        a1(&right),
        CellValue::Text {
            value: "persisted on v5".into()
        },
        "the migration itself must reach the peer"
    );

    left.edit_cell(
        SheetId(0),
        CellRef::parse_a1("A2").unwrap(),
        "after migration",
        CalculationOptions::default(),
    )
    .unwrap();
    let incremental = broadcast.lock().unwrap().drain(..).collect::<Vec<_>>();
    assert!(!incremental.is_empty(), "the edit must broadcast");
    for update in incremental {
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap();
    }
    assert_eq!(
        right
            .sheet(SheetId(0))
            .unwrap()
            .cell(CellRef::parse_a1("A2").unwrap())
            .map(|cell| cell.value.clone()),
        Some(CellValue::Text {
            value: "after migration".into()
        }),
        "an incremental update after migration must integrate, not stay pending"
    );
    assert_eq!(right.model(), left.model());
}

fn rename_first_sheet(snapshot: &[u8], name: &str) -> Vec<u8> {
    let doc = Doc::new();
    {
        let mut txn = doc.transact_mut();
        txn.apply_update(Update::decode_v1(snapshot).unwrap())
            .unwrap();
        let sheets = txn.get_map("xlsx:sheets").unwrap();
        let sheet = sheets
            .get(&txn, "sheet:0")
            .and_then(|value| value.cast::<MapRef>().ok())
            .unwrap();
        sheet.insert(&mut txn, "name", name);
    }
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}
