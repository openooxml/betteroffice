//! Two replicas appending sheets and merging cells while each also receives the
//! other's work. Every scenario runs under each delivery order and repeats the
//! exchange after an undo and a redo.

use betteroffice_xlsx::{
    CalculationOptions, CalculationRequest, Cell, CellRange, CellRef, CellState, CellValue,
    EditHistory, EditOperation, EditRequest, EditSource, EditStep, FreezePane, Hyperlink, Op,
    RangeAddress, RangeTarget, ReadRequest, Sheet, SheetId, Workbook, WorkbookModel,
};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Doc, Map, Transact, Update};

const OPTIONS: CalculationOptions = CalculationOptions { now_serial: None };

#[derive(Clone, Copy, Debug)]
enum Delivery {
    LeftThenRight,
    RightThenLeft,
    Merged,
}

const DELIVERIES: [Delivery; 3] = [
    Delivery::LeftThenRight,
    Delivery::RightThenLeft,
    Delivery::Merged,
];

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn text(value: &str) -> Cell {
    Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Default::default()
    }
}

fn source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut data = Sheet::new("Data");
    data.set_cell(cell("A1"), text("data"));
    let mut other = Sheet::new("Other");
    other.set_cell(cell("A1"), text("other"));
    model.sheets.push(data);
    model.sheets.push(other);
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

fn add(index: usize, name: &str) -> Op {
    Op::AddSheet {
        index,
        name: name.into(),
    }
}

fn set(sheet: u32, address: &str, value: &str) -> Op {
    Op::SetCell {
        sheet: SheetId(sheet),
        at: cell(address),
        cell: CellState {
            value: CellValue::Text {
                value: value.into(),
            },
            ..Default::default()
        },
    }
}

fn clear(sheet: u32, address: &str) -> Op {
    Op::SetCell {
        sheet: SheetId(sheet),
        at: cell(address),
        cell: CellState::default(),
    }
}

fn merge(sheet: u32, range: &str) -> Op {
    Op::MergeCells {
        sheet: SheetId(sheet),
        range: CellRange::parse_a1(range).unwrap(),
    }
}

fn names(workbook: &Workbook) -> Vec<String> {
    workbook
        .model()
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect()
}

fn input(workbook: &Workbook, sheet: &str, address: &str) -> Option<String> {
    let index = names(workbook).iter().position(|name| name == sheet)?;
    Some(
        workbook
            .cell(SheetId(index as u32), cell(address))
            .ok()?
            .input,
    )
}

struct Room {
    bytes: Vec<u8>,
    left: Workbook,
    right: Workbook,
    mirrors: u64,
}

impl Room {
    fn open() -> Self {
        Self::of(source())
    }

    fn of(bytes: Vec<u8>) -> Self {
        let left = Workbook::open_collaborative(&bytes, 1_001).unwrap();
        let right = Workbook::open_collaborative(&bytes, 2_002).unwrap();
        Self {
            bytes,
            left,
            right,
            mirrors: 3_000,
        }
    }

    fn exchange(&mut self, delivery: Delivery) {
        let to_right = self
            .left
            .encode_diff_v1(&self.right.encode_state_vector_v1())
            .unwrap();
        let to_left = self
            .right
            .encode_diff_v1(&self.left.encode_state_vector_v1())
            .unwrap();
        match delivery {
            Delivery::LeftThenRight => {
                self.right.apply_update_v1(&to_right, OPTIONS).unwrap();
                self.left.apply_update_v1(&to_left, OPTIONS).unwrap();
            }
            Delivery::RightThenLeft => {
                self.left.apply_update_v1(&to_left, OPTIONS).unwrap();
                self.right.apply_update_v1(&to_right, OPTIONS).unwrap();
            }
            Delivery::Merged => {
                let merged = Update::merge_updates(
                    [&to_right, &to_left]
                        .into_iter()
                        .map(|update| Update::decode_v1(update).unwrap()),
                )
                .encode_v1();
                self.left.apply_update_v1(&merged, OPTIONS).unwrap();
                self.right.apply_update_v1(&merged, OPTIONS).unwrap();
            }
        }
    }

    /// Both replicas hold one document, it saves to the same bytes on each, and
    /// a replica that only ever received it projects the same workbook.
    fn assert_converged(&mut self, context: &str) {
        assert_eq!(
            self.left.encode_state_vector_v1(),
            self.right.encode_state_vector_v1(),
            "state vectors: {context}"
        );
        assert_eq!(self.left.model(), self.right.model(), "model: {context}");
        let saved = self.left.save().unwrap();
        assert_eq!(saved, self.right.save().unwrap(), "saved bytes: {context}");
        assert_eq!(
            Workbook::open(&saved).unwrap().model(),
            Workbook::open(&self.right.save().unwrap()).unwrap().model(),
            "reopened model: {context}"
        );
        self.mirrors += 1;
        let mut mirror = Workbook::open_collaborative(&self.bytes, self.mirrors).unwrap();
        mirror
            .apply_update_v1(&self.left.encode_state_as_update_v1(), OPTIONS)
            .unwrap();
        assert_eq!(mirror.model(), self.left.model(), "mirror: {context}");
    }

    fn settle(&mut self, delivery: Delivery, context: &str) {
        self.exchange(delivery);
        self.assert_converged(&format!("{context} {delivery:?}"));
    }

    fn undo_left(&mut self, delivery: Delivery) {
        assert!(self.left.undo(OPTIONS).unwrap().applied);
        self.touch_right();
        self.settle(delivery, "undo");
    }

    fn redo_left(&mut self, delivery: Delivery) {
        assert!(self.left.redo(OPTIONS).unwrap().applied);
        self.touch_right();
        self.settle(delivery, "redo");
    }

    /// Keeps the exchange two-way, so neither side is only ever an author.
    fn touch_right(&mut self) {
        let at = format!("Z{}", self.mirrors);
        self.right
            .apply_ops(vec![set(0, &at, "peer")], OPTIONS)
            .unwrap();
    }
}

/// What applying both batches one after the other gives without collaboration.
fn sequential(ops: &[Vec<Op>]) -> WorkbookModel {
    let mut workbook = Workbook::open(&source()).unwrap();
    for batch in ops {
        workbook.apply_ops(batch.clone(), OPTIONS).unwrap();
    }
    Workbook::open(&workbook.save().unwrap())
        .unwrap()
        .model()
        .clone()
}

fn reopened(workbook: &Workbook) -> WorkbookModel {
    Workbook::open(&workbook.save().unwrap())
        .unwrap()
        .model()
        .clone()
}

#[test]
fn an_append_converges_with_a_concurrent_cell_edit() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let left_ops = vec![add(2, "Notes")];
        let right_ops = vec![set(0, "B3", "peer edit")];
        room.left.apply_ops(left_ops.clone(), OPTIONS).unwrap();
        room.right.apply_ops(right_ops.clone(), OPTIONS).unwrap();
        room.settle(delivery, "append");
        assert_eq!(names(&room.left), ["Data", "Other", "Notes"]);
        assert_eq!(
            reopened(&room.left),
            sequential(&[left_ops, right_ops]),
            "{delivery:?}"
        );

        room.undo_left(delivery);
        assert_eq!(names(&room.left), ["Data", "Other"]);
        assert_eq!(input(&room.right, "Data", "B3").unwrap(), "peer edit");

        room.redo_left(delivery);
        assert_eq!(names(&room.right), ["Data", "Other", "Notes"]);
        assert_eq!(input(&room.left, "Data", "B3").unwrap(), "peer edit");
    }
}

#[test]
fn two_appends_under_one_name_settle_on_the_same_names_everywhere() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Sheet4")], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![add(2, "Sheet4")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![set(2, "A1", "left sheet")], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![set(2, "A1", "right sheet")], OPTIONS)
            .unwrap();
        room.settle(delivery, "same name");
        assert_eq!(names(&room.left), ["Data", "Other", "Sheet4", "Sheet4 (2)"]);
        let holders = [
            input(&room.left, "Sheet4", "A1").unwrap(),
            input(&room.left, "Sheet4 (2)", "A1").unwrap(),
        ];
        assert!(holders.contains(&"left sheet".to_string()));
        assert!(holders.contains(&"right sheet".to_string()));

        room.undo_left(delivery);
        room.undo_left(delivery);
        assert_eq!(names(&room.left), ["Data", "Other", "Sheet4"]);
        assert_eq!(input(&room.left, "Sheet4", "A1").unwrap(), "right sheet");

        room.redo_left(delivery);
        room.redo_left(delivery);
        assert_eq!(
            names(&room.right),
            ["Data", "Other", "Sheet4", "Sheet4 (2)"]
        );
    }
}

#[test]
fn a_new_sheet_and_its_cells_converge_with_an_edit_on_another_sheet() {
    for delivery in DELIVERIES {
        for separate_batches in [false, true] {
            let mut room = Room::open();
            let mut left_ops = vec![vec![add(2, "Plan")], vec![set(2, "A1", "plan")]];
            if !separate_batches {
                left_ops = vec![left_ops.concat()];
            }
            let right_ops = vec![set(1, "C3", "other edit")];
            for batch in &left_ops {
                room.left.apply_ops(batch.clone(), OPTIONS).unwrap();
            }
            room.right.apply_ops(right_ops.clone(), OPTIONS).unwrap();
            room.settle(delivery, "cells");
            assert_eq!(input(&room.right, "Plan", "A1").unwrap(), "plan");
            let mut batches = left_ops.clone();
            batches.push(right_ops);
            assert_eq!(reopened(&room.left), sequential(&batches), "{delivery:?}");

            for _ in 0..left_ops.len() {
                room.undo_left(delivery);
            }
            assert_eq!(names(&room.left), ["Data", "Other"]);
            assert_eq!(input(&room.left, "Other", "C3").unwrap(), "other edit");

            for _ in 0..left_ops.len() {
                room.redo_left(delivery);
            }
            assert_eq!(input(&room.left, "Plan", "A1").unwrap(), "plan");
        }
    }
}

#[test]
fn a_formula_typed_before_the_sheet_arrives_resolves_after_the_exchange() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(
                vec![add(2, "Rates"), set(2, "A1", "seven percent")],
                OPTIONS,
            )
            .unwrap();
        room.right
            .edit_cell(SheetId(0), cell("C1"), "=Rates!A1", OPTIONS)
            .unwrap();
        room.settle(delivery, "formula");
        let at = room.right.model().sheets[0].cell(cell("C1")).unwrap();
        assert_eq!(
            at.value,
            CellValue::Text {
                value: "seven percent".into()
            }
        );
    }
}

const REFUSED_UNDO: &str = "undo would remove a sheet that holds collaborators' edits";

/// Refuses the next undo and checks it changed nothing but the undo depth: the
/// refusal is decided before anything is applied.
fn assert_refused_unchanged(workbook: &mut Workbook) {
    let model = workbook.model().clone();
    let vector = workbook.encode_state_vector_v1();
    let state = workbook.encode_state_as_update_v1();
    let saved = workbook.save().unwrap();
    let history = workbook.history_state();
    let error = workbook.undo(OPTIONS).unwrap_err();
    assert_eq!(error.to_string(), REFUSED_UNDO);
    assert_eq!(workbook.model(), &model);
    assert_eq!(workbook.encode_state_vector_v1(), vector);
    assert_eq!(workbook.encode_state_as_update_v1(), state);
    assert_eq!(workbook.save().unwrap(), saved);
    assert_eq!(workbook.history_state().undo_depth, history.undo_depth - 1);
    assert_eq!(workbook.history_state().redo_depth, history.redo_depth);
}

/// An undo that would delete a sheet a peer wrote in is refused and deletes
/// nothing; the refused step is dropped, so the next undo reaches the one
/// before it, and it never becomes redoable.
#[test]
fn undoing_an_add_a_peer_wrote_into_is_refused_and_moves_past_the_step() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![set(0, "C9", "before")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![add(2, "Shared")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "add");
        room.right
            .apply_ops(vec![set(2, "B2", "peer cell")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![set(2, "B3", "own cell")], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer cell");

        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.settle(delivery, "undo own cell");
        assert_refused_unchanged(&mut room.left);
        room.touch_right();
        room.settle(delivery, "refused undo");
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other", "Shared"]);
            assert_eq!(input(peer, "Shared", "B2").unwrap(), "peer cell");
        }

        room.undo_left(delivery);
        assert_eq!(input(&room.right, "Data", "C9").unwrap(), "");
        assert_eq!(names(&room.right), ["Data", "Other", "Shared"]);
        assert!(!room.left.history_state().can_undo);

        room.redo_left(delivery);
        assert_eq!(input(&room.right, "Data", "C9").unwrap(), "before");
        assert_eq!(room.left.history_state().redo_depth, 1);
    }
}

/// The undo removes the sheet with every cell the author wrote in it alone.
#[test]
fn undoing_an_add_only_its_author_wrote_into_removes_the_sheet() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Mine"), set(2, "A1", "own")], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![set(1, "B2", "elsewhere")], OPTIONS)
            .unwrap();
        room.settle(delivery, "own sheet");
        room.undo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other"]);
            assert_eq!(input(peer, "Other", "B2").unwrap(), "elsewhere");
        }
    }
}

/// A refused undo leaves every other step whole: the step before it restores
/// the value it overwrote, and redo brings the change back.
#[test]
fn a_refused_undo_keeps_the_earlier_steps_restorable() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.right
            .apply_ops(vec![set(0, "B5", "original")], OPTIONS)
            .unwrap();
        room.settle(delivery, "original");
        room.left
            .apply_ops(vec![set(0, "B5", "changed")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![add(2, "Shared")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "add");
        room.right
            .apply_ops(vec![set(2, "B2", "peer cell")], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer cell");

        let error = room.left.undo(OPTIONS).unwrap_err();
        assert_eq!(error.to_string(), REFUSED_UNDO);
        room.touch_right();
        room.settle(delivery, "refused undo");

        room.undo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(input(peer, "Data", "B5").unwrap(), "original");
            assert_eq!(input(peer, "Shared", "B2").unwrap(), "peer cell");
        }
        room.redo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(input(peer, "Data", "B5").unwrap(), "changed");
        }
        assert!(!room.left.history_state().can_redo);
        room.undo_left(delivery);
        assert_eq!(input(&room.right, "Data", "B5").unwrap(), "original");
    }
}

/// Undoing an overwrite of a peer's cell writes the peer's value back under
/// this replica's client id. The sheet still holds that value when its add is
/// undone, so the undo is refused and the value stays on every replica.
#[test]
fn a_peers_value_an_undo_wrote_back_still_refuses_the_add_undo() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Shared")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "add");
        room.right
            .apply_ops(vec![set(2, "B2", "bob")], OPTIONS)
            .unwrap();
        room.settle(delivery, "bob");
        room.left
            .apply_ops(vec![set(2, "B2", "alice")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "alice");
        room.undo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(input(peer, "Shared", "B2").unwrap(), "bob");
        }

        assert_refused_unchanged(&mut room.left);
        room.touch_right();
        room.settle(delivery, "refused undo");
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other", "Shared"]);
            assert_eq!(input(peer, "Shared", "B2").unwrap(), "bob");
        }
        assert!(!room.left.history_state().can_undo);
    }
}

/// The author's own cell in the new sheet is undone first, and the peer has
/// since overwritten it, so that step has nothing left to undo and the undo
/// reaches the add. The add is refused there, and the peer's value stays.
#[test]
fn an_add_reached_past_a_step_with_nothing_left_to_undo_is_still_refused() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Shared")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![set(2, "B2", "alice")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "alice");
        room.right
            .apply_ops(vec![set(2, "B2", "bob")], OPTIONS)
            .unwrap();
        room.settle(delivery, "bob");

        let error = room.left.undo(OPTIONS).unwrap_err();
        assert_eq!(error.to_string(), REFUSED_UNDO);
        assert!(!room.left.history_state().can_undo);
        room.touch_right();
        room.settle(delivery, "refused undo");
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other", "Shared"]);
            assert_eq!(input(peer, "Shared", "B2").unwrap(), "bob");
        }
    }
}

/// An add and its author's later steps in the sheet undo one by one, and the
/// add then removes the sheet.
#[test]
fn undoing_an_add_after_its_authors_later_steps_removes_the_sheet() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Mine")], OPTIONS).unwrap();
        room.left
            .apply_ops(vec![set(2, "A1", "own")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![merge(2, "C1:D2")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "own steps");
        room.undo_left(delivery);
        room.undo_left(delivery);
        room.undo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other"]);
        }
    }
}

/// Undoing a later change to a sheet setting writes the add's own empty value
/// back as a new struct, which holds nothing to lose, so the add still undoes.
#[test]
fn undoing_an_add_after_undoing_its_freeze_pane_removes_the_sheet() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Mine")], OPTIONS).unwrap();
        room.left
            .apply_ops(
                vec![Op::SetFreezePane {
                    sheet: SheetId(2),
                    pane: Some(FreezePane::new(1, 0, cell("A2"))),
                }],
                OPTIONS,
            )
            .unwrap();
        room.touch_right();
        room.settle(delivery, "freeze");
        room.undo_left(delivery);
        room.undo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other"]);
        }
    }
}

/// A merge, a link or a frozen pane a peer sets in the sheet is content like a
/// cell: only a setting holding an empty sheet's value is left out of the
/// check, so each of these alone refuses the undo and stays on every replica.
#[test]
fn a_peers_merge_link_or_freeze_pane_refuses_the_add_undo() {
    type Holds = fn(&Sheet) -> bool;
    let settings: [(&str, Op, Holds); 3] = [
        ("merge", merge(2, "B2:C3"), |sheet| !sheet.merges.is_empty()),
        (
            "link",
            Op::SetHyperlink {
                sheet: SheetId(2),
                hyperlink: Hyperlink {
                    range: CellRange::parse_a1("D4").unwrap(),
                    external_target: Some("https://example.com/".into()),
                    location: None,
                    tooltip: None,
                    display: None,
                },
            },
            |sheet| !sheet.hyperlinks.is_empty(),
        ),
        (
            "freeze pane",
            Op::SetFreezePane {
                sheet: SheetId(2),
                pane: Some(FreezePane::new(1, 0, cell("A2"))),
            },
            |sheet| sheet.freeze_pane.is_some(),
        ),
    ];
    for (label, op, holds) in settings {
        for delivery in DELIVERIES {
            let mut room = Room::open();
            room.left
                .apply_ops(vec![add(2, "Shared")], OPTIONS)
                .unwrap();
            room.touch_right();
            room.settle(delivery, label);
            room.right.apply_ops(vec![op.clone()], OPTIONS).unwrap();
            room.settle(delivery, label);

            assert_refused_unchanged(&mut room.left);
            room.touch_right();
            room.settle(delivery, label);
            for peer in [&room.left, &room.right] {
                assert_eq!(names(peer), ["Data", "Other", "Shared"], "{label}");
                assert!(holds(&peer.model().sheets[2]), "{label} {delivery:?}");
            }
        }
    }
}

/// The conservative side of the rule: a cell the add wrote, overwritten and
/// restored by undo, comes back as a new struct the add did not insert, so
/// the add's undo is refused although only its author wrote there.
#[test]
fn a_cell_from_the_add_restored_by_undo_refuses_the_add_undo() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Mine"), set(2, "A1", "first")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![set(2, "A1", "second")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "overwrite");
        room.undo_left(delivery);

        assert_refused_unchanged(&mut room.left);
        room.touch_right();
        room.settle(delivery, "refused undo");
        for peer in [&room.left, &room.right] {
            assert_eq!(input(peer, "Mine", "A1").unwrap(), "first");
        }
    }
}

/// A peer that wrote in the sheet and cleared its own cell again leaves
/// nothing there, so the undo removes the sheet and loses nothing.
#[test]
fn undoing_an_add_whose_peer_cleared_its_own_cell_removes_the_sheet() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Shared"), set(2, "A1", "own")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "add");
        room.right
            .apply_ops(vec![set(2, "B2", "peer cell")], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer cell");
        room.right.apply_ops(vec![clear(2, "B2")], OPTIONS).unwrap();
        room.settle(delivery, "peer clear");
        for peer in [&room.left, &room.right] {
            assert_eq!(input(peer, "Shared", "B2").unwrap(), "");
        }

        room.undo_left(delivery);
        for peer in [&room.left, &room.right] {
            assert_eq!(names(peer), ["Data", "Other"]);
        }
    }
}

/// Content kept out of history is no step's to remove, even this replica's
/// own, so the undo is refused rather than deleting it.
#[test]
fn undoing_an_add_holding_an_edit_kept_out_of_history_is_refused() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Shared")], OPTIONS)
            .unwrap();
        let sheet_id = room
            .left
            .read_cells(&ReadRequest::default())
            .unwrap()
            .unwrap()
            .sheets[2]
            .sheet_id
            .clone();
        let applied = room
            .left
            .apply_edits(&EditRequest {
                expect_version: room.left.version(),
                source: EditSource::Agent,
                history: EditHistory::None,
                calculation: CalculationRequest::default(),
                steps: vec![EditStep::new(EditOperation::SetCellInputs {
                    target: RangeTarget {
                        sheet_id,
                        range: RangeAddress::A1 { a1: "C3".into() },
                    },
                    inputs: vec![vec!["unrecorded".to_owned()]],
                })],
            })
            .unwrap()
            .unwrap();
        assert!(applied.applied);
        room.touch_right();
        room.settle(delivery, "unrecorded");

        assert_refused_unchanged(&mut room.left);
        room.touch_right();
        room.settle(delivery, "refused undo");
        for peer in [&room.left, &room.right] {
            assert_eq!(input(peer, "Shared", "C3").unwrap(), "unrecorded");
        }
    }
}

/// A replica that undid an add, reopened under the same client id and adopted
/// its own state mints a new key for its next sheet, so a peer receiving the
/// undo and the new append together sees a new sheet, not a rename.
#[test]
fn a_replica_reopened_under_its_client_id_never_reuses_an_undone_sheet_key() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Old")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "old");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        let state = room.left.encode_state_as_update_v1();
        let mut reopened =
            Workbook::open_collaborative(&room.bytes, room.left.client_id()).unwrap();
        reopened.apply_update_v1(&state, OPTIONS).unwrap();
        assert_eq!(names(&reopened), ["Data", "Other"]);
        room.left = reopened;
        room.left.apply_ops(vec![add(2, "New")], OPTIONS).unwrap();
        room.right
            .apply_ops(vec![set(1, "C3", "peer")], OPTIONS)
            .unwrap();
        room.settle(delivery, "new");
        assert_eq!(names(&room.right), ["Data", "Other", "New"]);
        assert_eq!(input(&room.left, "Other", "C3").unwrap(), "peer");
    }
}

const CHARTED: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");

/// A chart move writes every sheet whole. Sheets two peers appended under one
/// name keep the names they store, so the move is no rename and undoing one
/// add shows the other under the name its author gave it.
#[test]
fn a_chart_move_after_two_appends_under_one_name_keeps_the_stored_names() {
    for delivery in DELIVERIES {
        let mut room = Room::of(CHARTED.to_vec());
        let end = room.left.model().sheets.len();
        room.left
            .apply_ops(vec![add(end, "Notes")], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![add(end, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "names");
        let (sheet, frame) = room
            .left
            .model()
            .sheets
            .iter()
            .enumerate()
            .find_map(|(index, sheet)| Some((index, sheet.charts.first()?.frame_id())))
            .unwrap();
        room.left
            .move_chart(SheetId(sheet as u32), &frame, 16.0, 0.0, OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "move");
        let tail = |peer: &Workbook| names(peer)[end..].to_vec();
        assert_eq!(tail(&room.left), ["Notes", "Notes (2)"]);
        assert_eq!(tail(&room.right), ["Notes", "Notes (2)"]);

        room.undo_left(delivery);
        room.undo_left(delivery);
        assert_eq!(tail(&room.left), ["Notes"]);
        assert_eq!(tail(&room.right), ["Notes"]);
    }
}

/// The known limit: an undo that has not yet seen a peer's edit in the sheet
/// is allowed, and the edit lands in a sheet that no longer exists.
#[test]
fn a_peers_edit_racing_the_undo_of_an_add_is_lost_on_both_sides() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Shared")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "add");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.right
            .apply_ops(vec![set(2, "B2", "racing")], OPTIONS)
            .unwrap();
        room.settle(delivery, "race");
        assert_eq!(names(&room.left), ["Data", "Other"]);

        room.redo_left(delivery);
        assert_eq!(names(&room.right), ["Data", "Other", "Shared"]);
        assert_ne!(
            input(&room.right, "Shared", "B2").as_deref(),
            Some("racing")
        );
        assert_eq!(
            input(&room.right, "Shared", "B2"),
            input(&room.left, "Shared", "B2")
        );

        room.right.undo(OPTIONS).unwrap();
        room.left
            .apply_ops(vec![set(0, "D4", "after the race")], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer undo");
    }
}

#[test]
fn a_sheet_added_anywhere_but_the_end_is_refused_and_changes_nothing() {
    let mut room = Room::open();
    let model = room.left.model().clone();
    let state = room.left.encode_state_as_update_v1();
    let history = room.left.history_state();
    for index in [0, 1, 3] {
        let error = room
            .left
            .apply_ops(vec![add(index, "Early")], OPTIONS)
            .unwrap_err();
        assert_eq!(
            error.to_string(),
            "collaborative sheets can only be added at the end of the workbook"
        );
        assert_eq!(room.left.model(), &model);
        assert_eq!(room.left.encode_state_as_update_v1(), state);
        assert_eq!(room.left.history_state(), history);
    }
    let error = room
        .left
        .apply_ops(vec![add(2, "Last"), add(0, "Early")], OPTIONS)
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "collaborative sheets can only be added at the end of the workbook"
    );
    assert_eq!(room.left.encode_state_as_update_v1(), state);
}

#[test]
fn appends_from_one_batch_all_land_at_the_end() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "One"), add(3, "Two")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "batch");
        assert_eq!(names(&room.right), ["Data", "Other", "One", "Two"]);
    }
}

#[test]
fn a_standalone_peers_added_sheet_stays_refused() {
    let bytes = source();
    let mut standalone = Workbook::open(&bytes).unwrap();
    standalone
        .apply_ops(vec![add(2, "Notes")], OPTIONS)
        .unwrap();
    let mut peer = Workbook::open_collaborative(&bytes, 5_005).unwrap();
    let model = peer.model().clone();
    let update = standalone
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    let error = peer.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert_eq!(
        error.to_string(),
        "remote update changes the frozen workbook structure"
    );
    assert_eq!(peer.model(), &model);
}

/// A replica built for an older schema refuses the version a newer peer sends
/// through the same check this one runs on a version it does not know.
#[test]
fn a_schema_this_replica_does_not_know_is_refused_without_a_change() {
    let bytes = source();
    let mut peer = Workbook::open_collaborative(&bytes, 6_006).unwrap();
    let model = peer.model().clone();
    let state = peer.encode_state_as_update_v1();
    let newer = Doc::with_client_id(9_000);
    newer
        .transact_mut()
        .apply_update(Update::decode_v1(&state).unwrap())
        .unwrap();
    let meta = newer.get_or_insert_map("xlsx");
    let before = yrs::ReadTxn::state_vector(&newer.transact());
    meta.insert(&mut newer.transact_mut(), "schemaVersion", 99_i64);
    let update = yrs::ReadTxn::encode_diff_v1(&newer.transact(), &before);
    let error = peer.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert!(
        error.to_string().contains("unsupported schema version 99"),
        "{error}"
    );
    assert_eq!(peer.model(), &model);
    assert_eq!(peer.encode_state_as_update_v1(), state);
}

#[test]
fn concurrent_merges_converge_under_every_delivery() {
    for (left_range, right_range, expected) in [
        ("A1:B2", "D4:E5", vec!["A1:B2", "D4:E5"]),
        ("A1:B2", "B2:C3", vec!["A1:B2"]),
    ] {
        for delivery in DELIVERIES {
            let mut room = Room::open();
            room.left
                .apply_ops(vec![merge(0, left_range)], OPTIONS)
                .unwrap();
            room.right
                .apply_ops(vec![merge(0, right_range)], OPTIONS)
                .unwrap();
            room.settle(delivery, "merge");
            let merged = |workbook: &Workbook| {
                workbook.model().sheets[0]
                    .merges
                    .iter()
                    .map(CellRange::to_a1)
                    .collect::<Vec<_>>()
            };
            assert_eq!(merged(&room.left), expected, "{delivery:?}");
            assert_eq!(
                reopened(&room.right).sheets[0].merges,
                room.left.model().sheets[0].merges
            );

            room.undo_left(delivery);
            room.redo_left(delivery);
            assert_eq!(merged(&room.left), merged(&room.right), "{delivery:?}");
        }
    }
}
