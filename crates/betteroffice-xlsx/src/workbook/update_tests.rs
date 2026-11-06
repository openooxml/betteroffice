use super::*;
use crate::authority::LOCAL_UPDATE_ENCODINGS;
use xlsx_model::Cell;

fn model() -> WorkbookModel {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model
}

type HistoryPoint = (WorkbookModel, Vec<u8>, Vec<u8>, usize, usize);

fn sync(left: &mut Workbook, right: &mut Workbook) {
    let options = CalculationOptions::default();
    let to_right = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let to_left = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    right.apply_update_v1(&to_right, options).unwrap();
    left.apply_update_v1(&to_left, options).unwrap();
}

fn write(book: &mut Workbook, sheet: u32, row: u32, input: &str) {
    book.edit_cell(
        SheetId(sheet),
        CellRef::new(row, 1),
        input,
        CalculationOptions::default(),
    )
    .unwrap();
}

fn append(book: &mut Workbook) {
    book.apply_ops(
        vec![Op::AddSheet {
            index: 1,
            name: "Added".into(),
        }],
        CalculationOptions::default(),
    )
    .unwrap();
}

fn point(book: &mut Workbook) -> HistoryPoint {
    let history = book.history_state();
    (
        book.model().clone(),
        book.save().unwrap(),
        book.encode_state_vector_v1(),
        history.undo_depth,
        history.redo_depth,
    )
}

/// Runs `script` on two replicas, then every undo, every redo and every undo
/// again on the left one. Returns its state after each step.
fn history_trail(script: fn(&mut Workbook, &mut Workbook)) -> Vec<HistoryPoint> {
    let options = CalculationOptions::default();
    let bytes = Workbook::from_model(model()).unwrap().save().unwrap();
    let mut left = Workbook::open_collaborative(&bytes, 1_001).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 2_002).unwrap();
    script(&mut left, &mut right);
    let mut trail = vec![point(&mut left)];
    for redo in [false, true, false] {
        loop {
            let history = left.history_state();
            let pending = if redo {
                history.can_redo
            } else {
                history.can_undo
            };
            if !pending {
                break;
            }
            if redo {
                left.redo(options).unwrap();
            } else {
                left.undo(options).unwrap();
            }
            trail.push(point(&mut left));
        }
    }
    trail
}

/// Overwrites of one cell before and after an appended sheet, a peer's between
/// them.
fn overwrites_around_an_add(left: &mut Workbook, right: &mut Workbook) {
    write(left, 0, 0, "before the add");
    sync(left, right);
    write(right, 0, 0, "peer before the add");
    write(right, 0, 1, "peer below");
    sync(left, right);
    append(left);
    write(left, 0, 0, "after the add");
    write(left, 0, 1, "over the peer");
    sync(left, right);
    write(right, 0, 0, "peer after the add");
    sync(left, right);
    write(left, 0, 0, "over the peer again");
    write(left, 1, 0, "in the new sheet");
    assert!(left.history_state().undo_depth >= 5);
}

/// A peer has the last word on a cell the newest step wrote, so that step
/// cannot restore the value it replaced: the map-entry walk of `Item::redo`
/// stops at the peer's value only because no stack holds its deletion. The
/// steps below the add keep the full stack deep and the partial one shallow.
fn a_peer_outranks_the_step_above_an_add(left: &mut Workbook, right: &mut Workbook) {
    write(left, 0, 0, "replaced");
    write(left, 0, 1, "older");
    write(left, 0, 2, "older still");
    write(left, 0, 3, "oldest");
    append(left);
    write(left, 0, 0, "mine");
    sync(left, right);
    write(right, 0, 0, "peer has the last word");
    sync(left, right);
}

fn name(book: &mut Workbook, formula: &str) {
    book.apply_ops(
        vec![Op::SetDefinedName {
            name: "Rate".into(),
            refers_to: formula.into(),
            scope: None,
            hidden: None,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
}

/// The name a step replaced lies past a peer's value, as the cell does in
/// [`a_peer_outranks_the_step_above_an_add`], and a step above the add both
/// deletes and writes it.
fn a_peer_outranks_a_name_step_above_an_add(left: &mut Workbook, right: &mut Workbook) {
    name(left, "1");
    name(left, "2");
    write(left, 0, 1, "older");
    write(left, 0, 2, "older still");
    append(left);
    name(left, "3");
    sync(left, right);
    name(right, "4");
    sync(left, right);
    left.apply_ops(
        vec![Op::SetDefinedName {
            name: "Other".into(),
            refers_to: "5".into(),
            scope: None,
            hidden: None,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    left.apply_ops(
        vec![Op::DeleteDefinedName {
            name: "Other".into(),
            scope: None,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    assert!(left.history_state().undo_depth >= 7);
}

/// Steps above a sheet removal undo with only themselves on the manager's
/// stack. That is safe only while yrs reads the rest of the stack as
/// `undo_through_removal` assumes, so the same history must land on the same
/// state, step by step, with every step on the stack.
#[test]
fn partial_stack_undo_matches_the_full_stack() {
    for script in [
        overwrites_around_an_add as fn(&mut Workbook, &mut Workbook),
        a_peer_outranks_the_step_above_an_add,
        a_peer_outranks_a_name_step_above_an_add,
    ] {
        let before = crate::authority::partial_stack_undos();
        let partial = history_trail(script);
        assert!(crate::authority::partial_stack_undos() > before);
        let before = crate::authority::partial_stack_undos();
        let full = crate::authority::with_full_stack_undo(|| history_trail(script));
        assert_eq!(crate::authority::partial_stack_undos(), before);
        assert_eq!(partial.len(), full.len());
        for (index, (partial, full)) in partial.iter().zip(&full).enumerate() {
            assert_eq!(partial, full, "history step {index}");
        }
    }
}

#[test]
fn unobserved_fallback_does_not_encode_updates() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    LOCAL_UPDATE_ENCODINGS.set(0);
    crate::authority::with_full_materialization(|| {
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "1",
                CalculationOptions::default(),
            )
            .unwrap();
    });
    assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 0);
    assert_eq!(workbook.authority.materialize().unwrap(), *workbook.model());
}

#[test]
fn unobserved_standalone_edits_and_history_do_not_encode_updates() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let options = CalculationOptions::default();
    let version = workbook.version();
    let vector = workbook.encode_state_vector_v1();
    LOCAL_UPDATE_ENCODINGS.set(0);
    assert!(
        workbook
            .edit_cell(SheetId(0), CellRef::new(0, 0), "1", options)
            .unwrap()
            .applied
    );
    assert_ne!(workbook.version(), version);
    assert_ne!(workbook.encode_state_vector_v1(), vector);
    let version = workbook.version();
    assert!(
        workbook
            .edit_cell(SheetId(0), CellRef::new(0, 0), "", options)
            .unwrap()
            .applied
    );
    assert_ne!(workbook.version(), version);
    let version = workbook.version();
    assert!(workbook.undo(options).unwrap().applied);
    assert_ne!(workbook.version(), version);
    let version = workbook.version();
    assert!(workbook.redo(options).unwrap().applied);
    assert_ne!(workbook.version(), version);
    assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 0);
    assert_eq!(
        workbook.cell(SheetId(0), CellRef::new(0, 0)).unwrap().input,
        ""
    );
    let version = workbook.version();
    assert!(
        !workbook
            .edit_cell(SheetId(0), CellRef::new(0, 0), "", options)
            .unwrap()
            .applied
    );
    assert_eq!(workbook.version(), version);
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.cell(SheetId(0), CellRef::new(0, 0)).unwrap().input,
        ""
    );
}

#[test]
fn late_subscribers_can_sync_full_state_or_diff_and_receive_future_updates() {
    let options = CalculationOptions::default();
    for full_state in [false, true] {
        let mut workbook = Workbook::from_model(model()).unwrap();
        let mut peer = Workbook::from_model_collaborative(model(), 51).unwrap();
        let baseline = peer.encode_state_vector_v1();
        workbook
            .edit_cell(SheetId(0), CellRef::new(0, 0), "1", options)
            .unwrap();
        workbook
            .edit_cell(SheetId(0), CellRef::new(0, 0), "", options)
            .unwrap();
        workbook
            .edit_cell(SheetId(0), CellRef::new(1, 0), "2", options)
            .unwrap();
        let events = Arc::new(Mutex::new(Vec::new()));
        let observed = events.clone();
        let subscription = workbook
            .observe_update_v1(move |event| observed.lock().unwrap().push(event))
            .unwrap();
        assert!(events.lock().unwrap().is_empty());
        let initial = if full_state {
            workbook.encode_state_as_update_v1()
        } else {
            workbook.encode_diff_v1(&baseline).unwrap()
        };
        peer.apply_update_v1(&initial, options).unwrap();
        assert_eq!(
            peer.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        assert_eq!(peer.model(), workbook.model());
        LOCAL_UPDATE_ENCODINGS.set(0);
        workbook
            .edit_cell(SheetId(0), CellRef::new(1, 0), "3", options)
            .unwrap();
        workbook.undo(options).unwrap();
        workbook.redo(options).unwrap();
        assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 3);
        for event in events.lock().unwrap().drain(..) {
            assert_eq!(event.origin, UpdateOrigin::Local);
            peer.apply_update_v1(&event.update, options).unwrap();
        }
        assert_eq!(peer.model(), workbook.model());
        assert_eq!(
            peer.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        drop(subscription);
        LOCAL_UPDATE_ENCODINGS.set(0);
        workbook
            .edit_cell(SheetId(0), CellRef::new(1, 0), "", options)
            .unwrap();
        assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 0);
        let observed = events.clone();
        let _subscription = workbook
            .observe_update_v1(move |event| observed.lock().unwrap().push(event))
            .unwrap();
        peer.apply_update_v1(
            &workbook
                .encode_diff_v1(&peer.encode_state_vector_v1())
                .unwrap(),
            options,
        )
        .unwrap();
        workbook
            .edit_cell(SheetId(0), CellRef::new(2, 0), "4", options)
            .unwrap();
        assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 1);
        for event in events.lock().unwrap().drain(..) {
            peer.apply_update_v1(&event.update, options).unwrap();
        }
        assert_eq!(peer.model(), workbook.model());
        assert_eq!(
            peer.encode_state_as_update_v1(),
            workbook.encode_state_as_update_v1()
        );
    }
}

#[test]
fn collaborative_edits_keep_encoding_without_subscribers() {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::from_model_collaborative(model(), 61).unwrap();
    let mut peer = Workbook::from_model_collaborative(model(), 62).unwrap();
    LOCAL_UPDATE_ENCODINGS.set(0);
    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 0), "1", options)
        .unwrap();
    assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 1);
    peer.apply_update_v1(
        &workbook
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
        options,
    )
    .unwrap();
    workbook.undo(options).unwrap();
    peer.apply_update_v1(
        &workbook
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
        options,
    )
    .unwrap();
    workbook.redo(options).unwrap();
    peer.apply_update_v1(
        &workbook
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
        options,
    )
    .unwrap();
    assert_eq!(peer.model(), workbook.model());
    assert_eq!(
        peer.encode_state_vector_v1(),
        workbook.encode_state_vector_v1()
    );
}

#[test]
fn unobserved_atomic_edits_keep_the_update_needed_for_commit() {
    let mut workbook = Workbook::from_model(model()).unwrap();
    let version = workbook.version();
    LOCAL_UPDATE_ENCODINGS.set(0);
    workbook
        .edit_cells(
            SheetId(0),
            &[
                CellInput {
                    cell: CellRef::new(0, 0),
                    input: "1".into(),
                },
                CellInput {
                    cell: CellRef::new(1, 0),
                    input: "2".into(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(LOCAL_UPDATE_ENCODINGS.get(), 1);
    assert_ne!(workbook.version(), version);
    assert_eq!(workbook.authority.materialize().unwrap(), *workbook.model());
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.cell(SheetId(0), CellRef::new(1, 0)).unwrap().input,
        "2"
    );
}

#[test]
#[ignore]
fn large_workbook_update_encoding_benchmark() {
    use std::time::Instant;

    const ROWS: u32 = 100_000;
    const COLS: u32 = 10;
    const EDITS: usize = 100;
    let mut model = model();
    for row in 0..ROWS {
        for col in 0..COLS {
            model.sheets[0].set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row * COLS + col),
                    },
                    ..Default::default()
                },
            );
        }
    }
    let mut workbook = Workbook::from_model(model).unwrap();
    let options = CalculationOptions::default();
    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 0), "-1", options)
        .unwrap();
    let mut samples = [Vec::new(), Vec::new()];
    for round in 0..10 {
        for observed in [round % 2 == 0, round % 2 != 0] {
            let subscription = observed.then(|| workbook.observe_update_v1(|_| {}).unwrap());
            for edit in 0..EDITS / 10 {
                let input = format!(
                    "-{}",
                    2 + round * EDITS / 10 + edit + usize::from(observed) * EDITS
                );
                let start = Instant::now();
                let (_, profile) = workbook
                    .edit_cell_profiled(
                        SheetId(0),
                        CellRef::new(0, 0),
                        &input,
                        options,
                        &mut || start.elapsed().as_secs_f64() * 1000.0,
                    )
                    .unwrap();
                samples[usize::from(observed)].push(profile.apply_ms);
            }
            drop(subscription);
        }
    }
    for (observed, times) in samples.iter_mut().enumerate() {
        times.sort_by(f64::total_cmp);
        println!(
            "rows={ROWS} cols={COLS} edits={} observed={} apply_ms p50={:.4} p95={:.4}",
            times.len(),
            observed == 1,
            times[times.len() / 2],
            times[times.len() * 95 / 100]
        );
    }
}
