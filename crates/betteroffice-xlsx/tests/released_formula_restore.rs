//! Collaboration snapshots and recovery checkpoints made by a release that read
//! formulas as stored keep restoring: formulas stored with a leading `=` or with
//! storage prefixes stay as stored, so the workbook base fingerprint is unchanged.
//! The released state comes from `fixtures/leading-equals-f92d5602.md`.

use betteroffice_xlsx::{CalculationOptions, CellInput, CellRef, CellValue, SheetId, Workbook};
use yrs::updates::decoder::Decode;
use yrs::{Doc, Map, ReadTxn, Transact, Update};

const SOURCE: &[u8] = include_bytes!("fixtures/leading-equals.xlsx");
const RELEASED_UPDATE: &[u8] = include_bytes!("fixtures/leading-equals-f92d5602.update.bin");
const RELEASED_CHECKPOINT: &[u8] =
    include_bytes!("fixtures/leading-equals-f92d5602.checkpoint.bin");

const STORED: &[(&str, &str, &str)] = &[
    ("C1", "=A1*2", "=A1*2"),
    (
        "D1",
        r#"=HYPERLINK("https://example.com/a","site")"#,
        r#"=HYPERLINK("https://example.com/a","site")"#,
    ),
    ("E1", r#"_xlfn.CONCAT("a","b")"#, r#"=CONCAT("a","b")"#),
    (
        "F1",
        "=_xlfn.XLOOKUP(21,A1:A1,A1:A1)",
        "=XLOOKUP(21,A1:A1,A1:A1)",
    ),
];

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn base_fingerprint(state: &[u8]) -> String {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(state).unwrap())
        .unwrap();
    let txn = doc.transact();
    txn.get_map("xlsx")
        .and_then(|meta| meta.get(&txn, "baseFingerprint"))
        .and_then(|value| value.cast::<String>().ok())
        .unwrap()
}

fn assert_released_cells(workbook: &Workbook) {
    for (address, stored, shown) in STORED {
        let formula = workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell(address))
            .and_then(|cell| cell.formula.clone());
        assert_eq!(formula.as_deref(), Some(*stored), "{address}");
        assert_eq!(
            workbook.cell(SheetId(0), cell(address)).unwrap().input,
            *shown,
            "{address}"
        );
    }
    let b1 = workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell("B1"))
        .map(|cell| cell.value.clone());
    assert_eq!(
        b1,
        Some(CellValue::Text {
            value: "edited".into()
        })
    );
}

#[test]
fn a_released_collaboration_snapshot_restores() {
    let mut workbook = Workbook::open_collaborative_recalculated(SOURCE, 7_102, options()).unwrap();
    assert_eq!(
        base_fingerprint(&workbook.encode_state_as_update_v1()),
        base_fingerprint(RELEASED_UPDATE)
    );
    assert!(
        workbook
            .apply_update_v1(RELEASED_UPDATE, options())
            .unwrap()
            .applied
    );
    assert_released_cells(&workbook);
}

#[test]
fn a_released_recovery_checkpoint_restores_for_its_engine_version() {
    match Workbook::restore_recovery_checkpoint(RELEASED_CHECKPOINT) {
        Ok((workbook, ..)) => {
            assert_released_cells(&workbook);
            assert_eq!(
                base_fingerprint(&workbook.encode_state_as_update_v1()),
                base_fingerprint(RELEASED_UPDATE)
            );
        }
        Err(error) => assert!(
            error.to_string().contains("engine version mismatch"),
            "{error}"
        ),
    }
}

#[test]
fn a_checkpoint_keeps_the_released_fingerprint_and_takes_released_updates() {
    let writer = Workbook::open_collaborative(SOURCE, 7_103).unwrap();
    let checkpoint = writer.recovery_checkpoint(0, None).unwrap();
    let (mut restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(
        base_fingerprint(&restored.encode_state_as_update_v1()),
        base_fingerprint(RELEASED_UPDATE)
    );
    assert!(
        restored
            .apply_update_v1(RELEASED_UPDATE, options())
            .unwrap()
            .applied
    );
    assert_released_cells(&restored);
}

#[test]
fn stored_formulas_round_trip_a_standalone_checkpoint_and_recommit_unchanged() {
    let mut workbook = Workbook::open(SOURCE).unwrap();
    let version = workbook.version();
    let inputs: Vec<CellInput> = STORED
        .iter()
        .map(|(address, _, shown)| CellInput {
            cell: cell(address),
            input: (*shown).to_owned(),
        })
        .collect();
    workbook.edit_cells(SheetId(0), &inputs, options()).unwrap();
    assert_eq!(workbook.version(), version);
    let checkpoint = workbook.recovery_checkpoint(0, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    for (address, stored, shown) in STORED {
        let formula = restored
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell(address))
            .and_then(|cell| cell.formula.clone());
        assert_eq!(formula.as_deref(), Some(*stored), "{address}");
        assert_eq!(
            restored.cell(SheetId(0), cell(address)).unwrap().input,
            *shown
        );
    }
    assert_eq!(restored.save().unwrap(), SOURCE);
}

#[test]
fn a_stored_leading_equals_evaluates_as_the_formula() {
    let mut workbook = Workbook::open(SOURCE).unwrap();
    workbook.recalculate_all(options());
    let c1 = workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell("C1"))
        .unwrap();
    assert_eq!(c1.value, CellValue::Number { value: 42.0 });
    assert_eq!(c1.formula.as_deref(), Some("=A1*2"));
    workbook
        .edit_cell(SheetId(0), cell("A1"), "5", options())
        .unwrap();
    let c1 = workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell("C1"))
        .unwrap();
    assert_eq!(c1.value, CellValue::Number { value: 10.0 });
}
