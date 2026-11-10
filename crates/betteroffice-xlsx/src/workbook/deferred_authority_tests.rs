use std::sync::{Arc, Mutex};

use xlsx_model::{CellRange, CellRef, SheetId};
use xlsx_ops::{Op, StylePatch};

use super::{EAGER_AUTHORITY, Workbook};
use crate::CalculationOptions;

fn sources() -> Vec<(String, Vec<u8>)> {
    crate::authority::open_test_cases::matrix()
        .into_iter()
        .filter(|case| {
            case.model
                .sheets
                .iter()
                .all(|sheet| sheet.charts.is_empty())
        })
        .map(|case| {
            let parts = xlsx_parse::serialize_workbook(&case.model).unwrap();
            (case.label, ooxml_opc::rezip_parts(&parts).unwrap())
        })
        .collect()
}

fn open_eager(bytes: &[u8]) -> Workbook {
    struct Reset;
    impl Drop for Reset {
        fn drop(&mut self) {
            EAGER_AUTHORITY.set(false);
        }
    }
    let _reset = Reset;
    EAGER_AUTHORITY.set(true);
    let workbook = Workbook::open_recalculated(bytes, CalculationOptions::default()).unwrap();
    assert!(workbook.authority_is_built_for_test());
    workbook
}

fn open_deferred(bytes: &[u8]) -> Workbook {
    let workbook = Workbook::open_recalculated(bytes, CalculationOptions::default()).unwrap();
    assert!(!workbook.authority_is_built_for_test());
    workbook
}

fn edits() -> Vec<Vec<Op>> {
    vec![
        vec![Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:B2").unwrap(),
            patch: StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
        }],
        vec![Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        }],
        vec![Op::AddSheet {
            index: 1,
            name: "Added".into(),
        }],
    ]
}

fn batch(workbook: &mut Workbook) {
    let request = serde_json::json!({
        "expectVersion": workbook.version(),
        "steps": [{
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "C3:D3" } },
            "inputs": [["5", "=C3*2"]],
        }],
    });
    let applied: serde_json::Value =
        serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap()).unwrap();
    assert_eq!(applied["ok"], true, "{applied}");
}

fn edit_session(workbook: &mut Workbook) {
    let options = CalculationOptions::default();
    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 0), "41", options)
        .unwrap();
    for ops in edits() {
        workbook.apply_ops(ops, options).unwrap();
    }
    workbook.undo(options).unwrap();
    workbook.undo(options).unwrap();
    workbook.redo(options).unwrap();
    workbook
        .edit_cell(SheetId(0), CellRef::new(5, 1), "=A1+1", options)
        .unwrap();
}

fn assert_same(label: &str, deferred: &Workbook, eager: &Workbook) {
    assert_eq!(deferred.model(), eager.model(), "{label}");
    assert_eq!(deferred.save().unwrap(), eager.save().unwrap(), "{label}");
    assert_eq!(deferred.history_state(), eager.history_state(), "{label}");
    let structure = |workbook: &Workbook| {
        format!("{:?}", workbook.authority_for_test().structure().unwrap())
            .replace(&workbook.client_id().to_string(), "client")
    };
    assert_eq!(structure(deferred), structure(eager), "{label}");
    let (deferred, eager) = (deferred.authority_for_test(), eager.authority_for_test());
    assert_eq!(
        deferred.materialize().unwrap(),
        eager.materialize().unwrap(),
        "{label}"
    );
    assert_eq!(
        deferred.snapshot_identity().2,
        eager.snapshot_identity().2,
        "{label}"
    );
}

#[test]
fn reading_and_saving_leave_the_authority_unbuilt() {
    for (label, bytes) in sources() {
        let mut workbook = open_deferred(&bytes);
        workbook.save().unwrap();
        workbook.sheet_info().unwrap();
        workbook.cell(SheetId(0), CellRef::new(0, 0)).unwrap();
        workbook.recalculate(CalculationOptions::default());
        let events = Arc::new(Mutex::new(0));
        let counted = Arc::clone(&events);
        let _subscription = workbook
            .observe_update_v1(move |_| *counted.lock().unwrap() += 1)
            .unwrap();
        assert!(!workbook.authority_is_built_for_test(), "{label}");
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "41",
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(workbook.authority_is_built_for_test(), "{label}");
        assert_eq!(*events.lock().unwrap(), 1, "{label}");
    }
}

#[test]
fn a_deferred_authority_matches_an_eager_open() {
    for (label, bytes) in sources() {
        let deferred = open_deferred(&bytes);
        let eager = open_eager(&bytes);
        assert_eq!(deferred.model(), eager.model(), "{label}");
        assert_eq!(deferred.save().unwrap(), eager.save().unwrap(), "{label}");
        assert_eq!(
            deferred.encode_state_as_update_v1(),
            eager.encode_state_as_update_v1(),
            "{label}"
        );
        assert_eq!(
            deferred.encode_state_vector_v1(),
            eager.encode_state_vector_v1(),
            "{label}"
        );
        assert_same(&label, &deferred, &eager);
    }
}

#[test]
fn edits_without_an_observer_leave_the_authority_unbuilt() {
    for (label, bytes) in sources() {
        let mut deferred = open_deferred(&bytes);
        edit_session(&mut deferred);
        assert!(!deferred.authority_is_built_for_test(), "{label}");
        assert_eq!(deferred.deferred_changes_for_test(), 8, "{label}");
        deferred.save().unwrap();
        assert!(!deferred.authority_is_built_for_test(), "{label}");
    }
}

#[test]
fn edits_on_a_deferred_authority_match_an_eager_open() {
    for (label, bytes) in sources() {
        let mut deferred = open_deferred(&bytes);
        let mut eager = open_eager(&bytes);
        edit_session(&mut deferred);
        edit_session(&mut eager);
        assert!(!deferred.authority_is_built_for_test(), "{label}");
        assert_same(&label, &deferred, &eager);
    }
}

#[test]
fn a_batch_after_deferred_edits_stages_on_the_replayed_state() {
    for (label, bytes) in sources() {
        let mut deferred = open_deferred(&bytes);
        let mut eager = open_eager(&bytes);
        for workbook in [&mut deferred, &mut eager] {
            edit_session(workbook);
            batch(workbook);
        }
        assert!(deferred.authority_is_built_for_test(), "{label}");
        assert_eq!(deferred.deferred_changes_for_test(), 0, "{label}");
        assert_same(&label, &deferred, &eager);
    }
}

#[test]
fn an_observer_after_deferred_edits_gets_the_replayed_state() {
    let options = CalculationOptions::default();
    for (label, bytes) in sources() {
        let mut deferred = open_deferred(&bytes);
        let mut eager = open_eager(&bytes);
        for workbook in [&mut deferred, &mut eager] {
            edit_session(workbook);
        }
        let events = Arc::new(Mutex::new(Vec::new()));
        let seen = Arc::clone(&events);
        let _subscription = deferred
            .observe_update_v1(move |event| seen.lock().unwrap().push(event.update.len()))
            .unwrap();
        for workbook in [&mut deferred, &mut eager] {
            workbook
                .edit_cell(SheetId(0), CellRef::new(7, 0), "8", options)
                .unwrap();
        }
        assert!(deferred.authority_is_built_for_test(), "{label}");
        assert_eq!(deferred.deferred_changes_for_test(), 0, "{label}");
        let events = events.lock().unwrap();
        assert_eq!(events.len(), 1, "{label}");
        assert!(events[0] > 0, "{label}");
        assert_same(&label, &deferred, &eager);
    }
}

#[test]
fn a_deferred_authority_recovers_like_an_eager_one() {
    let options = CalculationOptions::default();
    for (label, bytes) in sources() {
        let deferred = open_deferred(&bytes);
        let eager = open_eager(&bytes);
        let (restored, ..) =
            Workbook::restore_recovery_checkpoint(&deferred.recovery_checkpoint(1, None).unwrap())
                .unwrap();
        let (expected, ..) =
            Workbook::restore_recovery_checkpoint(&eager.recovery_checkpoint(1, None).unwrap())
                .unwrap();
        assert_same(&label, &restored, &expected);

        let mut deferred = open_deferred(&bytes);
        deferred
            .edit_cell(SheetId(0), CellRef::new(0, 0), "41", options)
            .unwrap();
        let checkpoint = deferred
            .recovery_checkpoint_with_history(2, None, true)
            .unwrap();
        let (mut restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
        assert_eq!(restored.model(), deferred.model(), "{label}");
        restored.undo(options).unwrap();
        deferred.undo(options).unwrap();
        assert_same(&label, &restored, &deferred);
    }
}

#[test]
fn a_collaborative_open_builds_its_authority() {
    for (label, bytes) in sources() {
        let workbook = Workbook::open_collaborative(&bytes, 31).unwrap();
        assert!(workbook.authority_is_built_for_test(), "{label}");
    }
}
