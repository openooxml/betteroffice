//! `workbook-0.2.1-arrays-*.update.bin` are described beside them in
//! `workbook-0.2.1-arrays.update.md`.
//!
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
const ARRAYS: &[u8] = include_bytes!("fixtures/dynamic-arrays.xlsx");
const ARRAYS_UNTOUCHED: &[u8] =
    include_bytes!("fixtures/workbook-0.2.1-arrays-untouched.update.bin");
const ARRAYS_EDITED: &[u8] = include_bytes!("fixtures/workbook-0.2.1-arrays-edited.update.bin");
const ANCHORS_EDITED: &[u8] =
    include_bytes!("fixtures/workbook-b153acd5b-arrays-anchors-edited.update.bin");

fn a1(workbook: &Workbook) -> CellValue {
    workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(CellRef::parse_a1("A1").unwrap())
        .map(|cell| cell.value.clone())
        .unwrap_or(CellValue::Empty)
}

fn restored(source: &[u8], snapshot: &[u8], client_id: u64) -> Workbook {
    let mut workbook =
        Workbook::open_collaborative_recalculated(source, client_id, CalculationOptions::default())
            .unwrap();
    let result = workbook
        .apply_update_v1(snapshot, CalculationOptions::default())
        .unwrap();
    assert!(result.applied);
    workbook
}

#[test]
fn a_published_npm_snapshot_restores_after_initial_recalculation() {
    let source = include_bytes!("../../../packages/xlsx/test-fixtures/sample.xlsx");
    let snapshot = include_bytes!("fixtures/workbook-npm-0.2.1.update.bin");
    let mut workbook = restored(source, snapshot, 5_020);
    let marker = CellRef::parse_a1("A43").unwrap();
    let expected = CellValue::Text {
        value: "PublishedReleaseState".into(),
    };
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(marker)
            .unwrap()
            .value,
        expected
    );
    assert!(!workbook.can_undo());
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("B43").unwrap(),
            "after restore",
            CalculationOptions::default(),
        )
        .unwrap();
    let mut peer =
        Workbook::open_collaborative_recalculated(source, 5_021, CalculationOptions::default())
            .unwrap();
    peer.apply_update_v1(
        &workbook.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    assert_eq!(peer.model(), workbook.model());
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened
            .sheet(SheetId(0))
            .unwrap()
            .cell(marker)
            .unwrap()
            .value,
        expected
    );
    assert_eq!(
        reopened
            .sheet(SheetId(0))
            .unwrap()
            .cell(CellRef::parse_a1("B43").unwrap())
            .unwrap()
            .value,
        CellValue::Text {
            value: "after restore".into()
        }
    );
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
    let mut workbook =
        Workbook::open_collaborative_recalculated(SAMPLE, 5_006, CalculationOptions::default())
            .unwrap();
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

#[test]
fn a_cleared_cell_prevents_snapshot_replacement() {
    for legacy in [false, true] {
        let mut workbook =
            Workbook::open_collaborative_recalculated(SAMPLE, 5_022, CalculationOptions::default())
                .unwrap();
        let original = a1(&workbook);
        assert_ne!(original, CellValue::Empty);
        let snapshot = if legacy {
            SAMPLE_V5.to_vec()
        } else {
            workbook.encode_state_as_update_v1()
        };
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::parse_a1("A1").unwrap(),
                "",
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(a1(&workbook), CellValue::Empty);
        assert!(workbook.can_undo());
        let result = workbook.apply_update_v1(&snapshot, CalculationOptions::default());
        if legacy {
            assert!(result.is_err());
        } else {
            result.unwrap();
        }
        assert_eq!(a1(&workbook), CellValue::Empty);
        assert!(workbook.can_undo());
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_eq!(a1(&workbook), original);
    }
}

/// Adoption is gated on the base fingerprint, so a snapshot of some other
/// workbook is never taken as this one's state. (Merging one is a separate,
/// older problem: it lands a partial contamination either way.)
#[test]
fn a_snapshot_from_another_workbook_is_never_adopted() {
    let mut workbook =
        Workbook::open_collaborative_recalculated(SAMPLE, 5_008, CalculationOptions::default())
            .unwrap();
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
    let mut workbook =
        Workbook::open_collaborative_recalculated(SAMPLE, 5_010, CalculationOptions::default())
            .unwrap();
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
    let mut left =
        Workbook::open_collaborative_recalculated(SAMPLE, 5_011, CalculationOptions::default())
            .unwrap();
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

/// The schema a shared document declares and the cells its first sheet holds.
fn shared_contents(state: &[u8]) -> (i64, Vec<String>) {
    let doc = Doc::new();
    let mut txn = doc.transact_mut();
    txn.apply_update(Update::decode_v1(state).unwrap()).unwrap();
    let version = txn
        .get_map("xlsx")
        .and_then(|meta| meta.get(&txn, "schemaVersion"))
        .and_then(|value| value.cast::<i64>().ok())
        .unwrap();
    let contents = txn
        .get_map("xlsx:sheets")
        .and_then(|sheets| sheets.get(&txn, "sheet:0"))
        .and_then(|sheet| sheet.cast::<MapRef>().ok())
        .and_then(|sheet| sheet.get(&txn, "contents"))
        .and_then(|contents| contents.cast::<MapRef>().ok())
        .unwrap();
    let mut keys: Vec<String> = contents.keys(&txn).map(str::to_owned).collect();
    keys.sort();
    (version, keys)
}

fn values(workbook: &Workbook, cells: &[&str]) -> Vec<CellValue> {
    let sheet = workbook.sheet(SheetId(0)).unwrap();
    cells
        .iter()
        .map(|address| {
            sheet
                .cell(CellRef::parse_a1(address).unwrap())
                .map(|cell| cell.value.clone())
                .unwrap_or(CellValue::Empty)
        })
        .collect()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

/// A 0.2.1 room nobody edited holds nothing but its seed, so the arrays'
/// cached results in it are provably the seed's: the upgrade takes them out
/// of the shared document and the arrays compute them again.
#[test]
fn an_untouched_released_array_room_upgrades_to_computed_results() {
    let mut workbook = restored(ARRAYS, ARRAYS_UNTOUCHED, 7_100);
    assert_eq!(
        values(&workbook, &["C1", "C2", "C3", "E1", "E2"]),
        [
            number(1.0),
            number(2.0),
            number(3.0),
            number(6.0),
            number(2.0)
        ]
    );
    let sheet = &workbook.model().sheets[0];
    assert!(sheet.is_dynamic_array(CellRef::parse_a1("C1").unwrap()));
    assert!(
        sheet
            .array_definition(CellRef::parse_a1("E1").unwrap())
            .is_some_and(|definition| !definition.is_dynamic())
    );
    let (version, cells) = shared_contents(&workbook.encode_state_as_update_v1());
    assert_eq!(version, 7);
    assert_eq!(cells, ["0:0", "0:2", "0:4", "1:0", "2:0"]);

    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("A1").unwrap(),
            "0",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        values(&workbook, &["C1", "C2", "C3", "E1", "E2"]),
        [
            number(0.0),
            number(1.0),
            number(2.0),
            number(0.0),
            number(2.0)
        ]
    );
    let mut peer =
        Workbook::open_collaborative_recalculated(ARRAYS, 7_101, CalculationOptions::default())
            .unwrap();
    peer.apply_update_v1(
        &workbook.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    assert_eq!(peer.model(), workbook.model());
}

/// Once a 0.2.1 room was edited, a constant where an array's result stood may
/// be an author's, and its value matching the cache proves nothing: the
/// upgrade keeps it, so the arrays report `#SPILL!` until it is cleared.
#[test]
fn an_edited_released_array_room_keeps_its_constants() {
    let mut workbook = restored(ARRAYS, ARRAYS_EDITED, 7_200);
    let spill = CellValue::Error {
        value: betteroffice_xlsx::ErrorValue::Spill,
    };
    assert_eq!(
        values(&workbook, &["C1", "C2", "C3", "E1", "E2"]),
        [spill.clone(), number(2.0), number(3.0), spill, number(2.0)]
    );
    assert_eq!(
        values(&workbook, &["G1"]),
        [CellValue::Text {
            value: "note".into()
        }]
    );
    let (version, cells) = shared_contents(&workbook.encode_state_as_update_v1());
    assert_eq!(version, 7);
    assert!(cells.contains(&"1:2".to_owned()) && cells.contains(&"1:4".to_owned()));

    for address in ["C2", "C3"] {
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::parse_a1(address).unwrap(),
                "",
                CalculationOptions::default(),
            )
            .unwrap();
    }
    assert_eq!(
        values(&workbook, &["C1", "C2", "C3"]),
        [number(1.0), number(2.0), number(3.0)]
    );
}

/// A room this release creates or migrates declares schema 7, which releases
/// 0.1.x and 0.2.x refuse: they read schemas 3 through 6 only.
#[test]
fn a_current_room_declares_the_schema_older_releases_refuse() {
    let fresh =
        Workbook::open_collaborative_recalculated(ARRAYS, 7_300, CalculationOptions::default())
            .unwrap();
    assert_eq!(shared_contents(&fresh.encode_state_as_update_v1()).0, 7);
    assert_eq!(shared_contents(ARRAYS_EDITED).0, 6);
    let migrated = restored(ARRAYS, ARRAYS_EDITED, 7_301);
    assert_eq!(shared_contents(&migrated.encode_state_as_update_v1()).0, 7);
}

/// A migrated room's legacy array with an author's constant inside it saves
/// over its anchor alone, so reading the file back keeps the constant the
/// author's rather than taking it for the array's result.
#[test]
fn a_conflicted_legacy_array_keeps_the_authors_constant_through_a_save() {
    let workbook = restored(ARRAYS, ARRAYS_EDITED, 7_500);
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();
    assert!(
        sheet.contains(r#"<f t="array" ref="E1">A1:A2*2</f>"#),
        "{sheet}"
    );
    let mut reopened = Workbook::open_recalculated(&saved, CalculationOptions::default()).unwrap();
    assert_eq!(values(&reopened, &["E1", "E2"]), [number(6.0), number(2.0)]);
    reopened
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("A2").unwrap(),
            "5",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(values(&reopened, &["E2"]), [number(2.0)]);
}

/// The release before schema 7 turns an array anchor edited in a room into a
/// plain formula, and so does migrating that room: an anchor's array is
/// provably its own only while its formula is the one the workbook opened with.
#[test]
fn anchors_edited_before_schema_7_migrate_as_the_plain_formulas_they_became() {
    let workbook = restored(ARRAYS, ANCHORS_EDITED, 7_600);
    let sheet = &workbook.model().sheets[0];
    for anchor in ["C1", "E1"] {
        assert_eq!(
            sheet.array_definition(CellRef::parse_a1(anchor).unwrap()),
            None,
            "{anchor}"
        );
    }
    assert_eq!(
        values(&workbook, &["C1", "C2", "C3", "E1", "E2"]),
        [
            number(3.0),
            number(2.0),
            number(3.0),
            CellValue::Error {
                value: betteroffice_xlsx::ErrorValue::Value
            },
            number(2.0)
        ]
    );
    let mut peer =
        Workbook::open_collaborative_recalculated(ARRAYS, 7_601, CalculationOptions::default())
            .unwrap();
    peer.apply_update_v1(ANCHORS_EDITED, CalculationOptions::default())
        .unwrap();
    assert_eq!(peer.model(), workbook.model());
}
