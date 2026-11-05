//! Two replicas editing defined names while each also receives the other's
//! work. The source table stays the baseline every replica opened, so a name
//! nobody edited has nothing in the shared document. Every scenario runs under
//! each delivery order and repeats the exchange after an undo and a redo.

use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, DefinedName, Error, ErrorValue, Op, Sheet,
    SheetId, Workbook, WorkbookModel,
};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Doc, Map, Observable, ReadTxn, Transact, Update};

const OPTIONS: CalculationOptions = CalculationOptions { now_serial: None };
const NAMES_ROOT: &str = "xlsx:defined-names";
const SAMPLE: &[u8] = include_bytes!("../../../apps/demo/public/sample.xlsx");
const SAMPLE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-sample.update.bin");
const CHARTED: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");

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

fn number(value: f64) -> Cell {
    Cell {
        value: CellValue::Number { value },
        ..Default::default()
    }
}

fn defined(name: &str, formula: &str, sheet: Option<u32>, hidden: bool) -> DefinedName {
    DefinedName {
        name: name.into(),
        formula: formula.into(),
        local_sheet: sheet.map(SheetId),
        hidden,
    }
}

fn package(names: Vec<DefinedName>) -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut data = Sheet::new("Data");
    data.set_cell(cell("A1"), number(10.0));
    let mut other = Sheet::new("Other");
    other.set_cell(cell("A1"), number(20.0));
    model.sheets.push(data);
    model.sheets.push(other);
    model.defined_names = names;
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap();
    *xml = String::from_utf8(xml.clone())
        .unwrap()
        .replace(r#"name="Rate""#, r#"name="Rate" comment="kept""#)
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn source() -> Vec<u8> {
    package(vec![
        defined("Rate", "Data!$A$1", None, false),
        defined("Local", "Other!$A$1", Some(1), false),
        defined("Secret", "Data!$A$1", None, true),
        defined("_xlnm.Print_Area", "Data!$A$1:$B$2", Some(0), false),
        defined("_xlnm._FilterDatabase", "Data!$A$1:$B$2", Some(0), true),
    ])
}

fn set(name: &str, formula: &str) -> Op {
    set_in(name, formula, None)
}

fn set_in(name: &str, formula: &str, scope: Option<u32>) -> Op {
    Op::SetDefinedName {
        name: name.into(),
        refers_to: formula.into(),
        scope: scope.map(SheetId),
        hidden: None,
    }
}

fn delete(name: &str) -> Op {
    Op::DeleteDefinedName {
        name: name.into(),
        scope: None,
    }
}

fn add(index: usize, name: &str) -> Op {
    Op::AddSheet {
        index,
        name: name.into(),
    }
}

fn write(sheet: u32, address: &str, value: &str) -> Op {
    Op::SetCell {
        sheet: SheetId(sheet),
        at: cell(address),
        cell: betteroffice_xlsx::CellState {
            value: CellValue::Text {
                value: value.into(),
            },
            ..Default::default()
        },
    }
}

/// The formula `name` holds in `scope`.
fn formula(workbook: &Workbook, name: &str, scope: Option<u32>) -> Option<String> {
    workbook
        .model()
        .defined_names
        .iter()
        .find(|defined| {
            defined.name.eq_ignore_ascii_case(name) && defined.local_sheet == scope.map(SheetId)
        })
        .map(|defined| defined.formula.clone())
}

fn name_list(workbook: &Workbook) -> Vec<String> {
    workbook
        .model()
        .defined_names
        .iter()
        .map(|defined| defined.name.clone())
        .collect()
}

fn value(workbook: &Workbook, sheet: u32, address: &str) -> CellValue {
    workbook
        .sheet(SheetId(sheet))
        .unwrap()
        .cell(cell(address))
        .map(|cell| cell.value.clone())
        .unwrap_or(CellValue::Empty)
}

fn names_xml(workbook: &Workbook) -> String {
    let xml = String::from_utf8(workbook_xml(workbook)).unwrap();
    let start = xml.find("<definedNames").unwrap();
    let end = xml.find("</definedNames>").unwrap();
    xml[start..end].to_string()
}

fn workbook_xml(workbook: &Workbook) -> Vec<u8> {
    let saved = workbook.save().unwrap();
    ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap()
        .1
}

fn sheet_names(workbook: &Workbook) -> Vec<String> {
    workbook
        .model()
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect()
}

fn roots(state: &[u8]) -> Vec<String> {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(state).unwrap())
        .unwrap();
    let txn = doc.transact();
    let mut roots = txn
        .root_refs()
        .map(|(key, _)| key.to_string())
        .collect::<Vec<_>>();
    roots.sort();
    roots
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
        self.assert_same_document(context);
        assert_eq!(
            Workbook::open(&self.left.save().unwrap()).unwrap().model(),
            Workbook::open(&self.right.save().unwrap()).unwrap().model(),
            "reopened model: {context}"
        );
    }

    fn assert_same_document(&mut self, context: &str) {
        assert_eq!(
            self.left.encode_state_vector_v1(),
            self.right.encode_state_vector_v1(),
            "state vectors: {context}"
        );
        assert_eq!(self.left.model(), self.right.model(), "model: {context}");
        assert_eq!(
            self.left.save().unwrap(),
            self.right.save().unwrap(),
            "saved bytes: {context}"
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

    /// Both sides author in every exchange, so neither is only ever a receiver.
    fn touch_left(&mut self) {
        self.mirrors += 1;
        let at = format!("Y{}", self.mirrors);
        self.left
            .apply_ops(vec![write(0, &at, "peer")], OPTIONS)
            .unwrap();
    }

    fn touch_right(&mut self) {
        self.mirrors += 1;
        let at = format!("Z{}", self.mirrors);
        self.right
            .apply_ops(vec![write(0, &at, "peer")], OPTIONS)
            .unwrap();
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

    fn undo_right(&mut self, delivery: Delivery) {
        assert!(self.right.undo(OPTIONS).unwrap().applied);
        self.touch_left();
        self.settle(delivery, "undo");
    }

    fn redo_right(&mut self, delivery: Delivery) {
        assert!(self.right.redo(OPTIONS).unwrap().applied);
        self.touch_left();
        self.settle(delivery, "redo");
    }

    /// Lands `left` and `right` on their own replicas, then exchanges them.
    fn concurrently(&mut self, left: Vec<Op>, right: Vec<Op>, delivery: Delivery) {
        self.left.apply_ops(left, OPTIONS).unwrap();
        self.right.apply_ops(right, OPTIONS).unwrap();
        self.settle(delivery, "concurrent");
    }

    fn assert_both(&self, check: impl Fn(&Workbook)) {
        check(&self.left);
        check(&self.right);
    }

    /// Which side a same-name race went to, by the casing that survived.
    fn winner_is_left(&self, name_on_left: &str) -> bool {
        name_list(&self.left)
            .iter()
            .any(|name| name == name_on_left)
    }
}

/// What applying the batches one after the other gives without collaboration.
fn sequential(bytes: &[u8], ops: &[Vec<Op>]) -> WorkbookModel {
    let mut workbook = Workbook::open(bytes).unwrap();
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
fn different_names_both_land_and_match_the_sequential_result() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let left = vec![set("Alpha", "1")];
        let right = vec![set("Beta", "2"), set("Rate", "Other!$A$1")];
        room.concurrently(left.clone(), right.clone(), delivery);
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Alpha", None).as_deref(), Some("1"));
            assert_eq!(formula(peer, "Beta", None).as_deref(), Some("2"));
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("Other!$A$1"));
        });
        let expected = sequential(&room.bytes, &[left, right]);
        assert_eq!(
            reopened(&room.left).defined_names,
            expected.defined_names,
            "{delivery:?}"
        );

        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Alpha", None), None));
        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Alpha", None).as_deref(), Some("1")));
    }
}

#[test]
fn new_names_keep_the_order_they_were_added_in() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let added = vec![set("Zeta", "1"), set("Alpha", "2")];
        room.left
            .apply_ops(vec![added[0].clone()], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![added[1].clone()], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        let tail = |peer: &Workbook| name_list(peer)[5..].to_vec();
        assert_eq!(tail(&room.right), ["Zeta", "Alpha"]);
        let edit = vec![set("Alpha", "3")];
        room.right.apply_ops(edit.clone(), OPTIONS).unwrap();
        room.touch_left();
        room.settle(delivery, "edited");
        assert_eq!(tail(&room.left), ["Zeta", "Alpha"]);
        assert_eq!(
            reopened(&room.left).defined_names,
            sequential(&room.bytes, &[added, edit]).defined_names
        );
    }
}

/// `Rate` and `RATE` are one name, so concurrent writes are one register. The
/// winner's undo removes the only live value, and the loser's write does not
/// come back: it was replaced, the winner's step never deleted it, so nothing
/// restores it. That is what a cell does when two people write it. The loser's
/// step has nothing left to undo, so its undo falls through to the step before.
#[test]
fn rate_and_rate_are_one_name_and_one_winner_decides_it() {
    for delivery in DELIVERIES {
        for winner_undoes in [true, false] {
            let mut room = Room::open();
            room.left
                .apply_ops(vec![write(0, "C1", "left earlier")], OPTIONS)
                .unwrap();
            room.right
                .apply_ops(vec![write(0, "C2", "right earlier")], OPTIONS)
                .unwrap();
            room.settle(delivery, "earlier steps");
            room.concurrently(vec![set("Margin", "1")], vec![set("MARGIN", "2")], delivery);
            let left_won = room.winner_is_left("Margin");
            let (casing, kept) = if left_won {
                ("Margin", "1")
            } else {
                ("MARGIN", "2")
            };
            room.assert_both(|peer| {
                assert_eq!(name_list(peer)[5..], [casing.to_string()]);
                assert_eq!(formula(peer, casing, None).as_deref(), Some(kept));
            });
            let left_undoes = left_won == winner_undoes;
            if left_undoes {
                room.undo_left(delivery);
            } else {
                room.undo_right(delivery);
            }
            if winner_undoes {
                room.assert_both(|peer| assert_eq!(formula(peer, "margin", None), None));
                if left_undoes {
                    room.redo_left(delivery);
                } else {
                    room.redo_right(delivery);
                }
                room.assert_both(|peer| {
                    assert_eq!(name_list(peer)[5..], [casing.to_string()]);
                    assert_eq!(formula(peer, casing, None).as_deref(), Some(kept));
                });
            } else {
                let earlier = if left_undoes { "C1" } else { "C2" };
                room.assert_both(|peer| {
                    assert_eq!(formula(peer, casing, None).as_deref(), Some(kept));
                    assert_eq!(value(peer, 0, earlier), CellValue::Empty);
                });
            }
        }
    }
}

/// The same race on a source name: undoing the winning edit shows the source
/// value through, not the loser's edit.
#[test]
fn undoing_the_winning_edit_of_a_source_name_shows_the_source_value() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .edit_cell(SheetId(0), cell("C1"), "=Rate*2", OPTIONS)
            .unwrap();
        room.settle(delivery, "formula");
        room.concurrently(vec![set("Rate", "3")], vec![set("Rate", "4")], delivery);
        let left_won = formula(&room.left, "Rate", None).as_deref() == Some("3");
        let kept = if left_won { 6.0 } else { 8.0 };
        room.assert_both(|peer| {
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: kept });
        });
        if left_won {
            room.undo_left(delivery);
        } else {
            room.undo_right(delivery);
        }
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("Data!$A$1"));
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: 20.0 });
        });
    }
}

#[test]
fn one_text_in_two_scopes_is_two_names() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .edit_cell(SheetId(0), cell("C1"), "=Shared", OPTIONS)
            .unwrap();
        room.right
            .edit_cell(SheetId(1), cell("C1"), "=Shared", OPTIONS)
            .unwrap();
        room.concurrently(
            vec![set("Shared", "1")],
            vec![set_in("Shared", "2", Some(1))],
            delivery,
        );
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Shared", None).as_deref(), Some("1"));
            assert_eq!(formula(peer, "Shared", Some(1)).as_deref(), Some("2"));
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: 1.0 });
            assert_eq!(value(peer, 1, "C1"), CellValue::Number { value: 2.0 });
        });
    }
}

/// A set outranks a delete whichever client writes which, on a source name and
/// on a name a peer added.
#[test]
fn a_set_beats_a_concurrent_delete_on_every_replica() {
    for delivery in DELIVERIES {
        for setter_is_left in [true, false] {
            for target in ["Rate", "Fresh"] {
                let mut room = Room::open();
                if target == "Fresh" {
                    room.left
                        .apply_ops(vec![set("Fresh", "1")], OPTIONS)
                        .unwrap();
                    room.touch_right();
                    room.settle(delivery, "fresh");
                }
                let setting = vec![set(target, "9")];
                let deleting = vec![delete(target)];
                let (left, right) = if setter_is_left {
                    (setting, deleting)
                } else {
                    (deleting, setting)
                };
                room.concurrently(left, right, delivery);
                room.assert_both(|peer| {
                    assert_eq!(formula(peer, target, None).as_deref(), Some("9"));
                });
            }
        }
    }
}

/// Undoing the set leaves the concurrent delete standing, and undoing the delete
/// instead leaves the set, so each side's undo takes back exactly its own edit.
#[test]
fn each_side_of_a_set_and_delete_race_undoes_its_own_edit() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.concurrently(vec![set("Rate", "9")], vec![delete("Rate")], delivery);
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None), None));
        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None).as_deref(), Some("9")));

        let mut room = Room::open();
        room.concurrently(vec![set("Rate", "9")], vec![delete("Rate")], delivery);
        assert!(room.right.undo(OPTIONS).unwrap().applied);
        room.settle(delivery, "undo the delete");
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None).as_deref(), Some("9")));
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("Data!$A$1"));
        });
    }
}

#[test]
fn delete_against_delete_settles_and_both_undos_converge() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.concurrently(vec![delete("Rate")], vec![delete("Rate")], delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None), None));
        room.left.undo(OPTIONS).unwrap();
        room.settle(delivery, "left undo");
        room.right.undo(OPTIONS).unwrap();
        room.settle(delivery, "right undo");
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("Data!$A$1"));
        });
    }
}

#[test]
fn a_collaborators_later_value_wins_over_undo_which_falls_through() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![write(0, "C1", "earlier")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![set("Rate", "5")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "edit");
        room.right
            .apply_ops(vec![set("Rate", "6")], OPTIONS)
            .unwrap();
        room.settle(delivery, "overwrite");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.touch_right();
        room.settle(delivery, "undo");
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("6"));
            assert_eq!(value(peer, 0, "C1"), CellValue::Empty);
        });
    }
}

#[test]
fn a_set_after_seeing_a_delete_outlives_the_undo_of_that_delete() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![write(0, "C1", "earlier")], OPTIONS)
            .unwrap();
        room.left.apply_ops(vec![delete("Rate")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "delete");
        room.right
            .apply_ops(vec![set("Rate", "6")], OPTIONS)
            .unwrap();
        room.settle(delivery, "set");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.touch_right();
        room.settle(delivery, "undo");
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("6"));
            assert_eq!(value(peer, 0, "C1"), CellValue::Empty);
        });
    }
}

#[test]
fn a_source_name_edited_deleted_and_undone_returns_to_the_source_bytes() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let source = names_xml(&room.left);
        room.left
            .apply_ops(vec![set("Rate", "5")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "edit");
        room.left.apply_ops(vec![delete("Rate")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "delete");
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None), None));

        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None).as_deref(), Some("5")));
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Rate", None).as_deref(), Some("Data!$A$1"));
            assert_eq!(names_xml(peer), source);
        });

        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None).as_deref(), Some("5")));
        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Rate", None), None));
    }
}

#[test]
fn deleting_a_source_name_and_undoing_restores_the_source_bytes() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let source = names_xml(&room.left);
        room.left.apply_ops(vec![delete("Rate")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "delete");
        assert!(!String::from_utf8_lossy(&workbook_xml(&room.right)).contains(r#"name="Rate""#));
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(names_xml(peer), source));
    }
}

#[test]
fn an_edited_source_name_keeps_what_the_model_does_not_hold() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![set("Rate", "5")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "edit");
        room.assert_both(|peer| {
            let xml = String::from_utf8(workbook_xml(peer)).unwrap();
            assert!(xml.contains(r#"comment="kept""#), "{xml}");
        });
    }
}

/// A source name set again keeps its source place in the model, where a
/// standalone session moves it to the end, and saves to the same table.
#[test]
fn a_source_name_deleted_and_set_again_saves_like_the_sequential_result() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let ops = [vec![delete("Rate")], vec![set("Rate", "7")]];
        for batch in &ops {
            room.left.apply_ops(batch.clone(), OPTIONS).unwrap();
        }
        room.touch_right();
        room.settle(delivery, "again");
        assert_eq!(
            reopened(&room.right).defined_names,
            sequential(&room.bytes, &ops).defined_names
        );
    }
}

#[test]
fn a_formula_on_a_name_a_peer_deletes_is_name_error_everywhere() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .edit_cell(SheetId(0), cell("C1"), "=Rate*2", OPTIONS)
            .unwrap();
        room.settle(delivery, "formula");
        room.assert_both(|peer| {
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: 20.0 });
        });
        room.right.apply_ops(vec![delete("Rate")], OPTIONS).unwrap();
        room.touch_left();
        room.settle(delivery, "delete");
        let missing = CellValue::Error {
            value: ErrorValue::Name,
        };
        room.assert_both(|peer| assert_eq!(value(peer, 0, "C1"), missing));

        room.right.undo(OPTIONS).unwrap();
        room.touch_left();
        room.settle(delivery, "undo");
        room.assert_both(|peer| {
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: 20.0 });
        });
    }
}

#[test]
fn a_formula_typed_before_its_name_arrives_resolves_on_the_exchange() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .edit_cell(SheetId(0), cell("C1"), "=Soon*2", OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![set("Soon", "21")], OPTIONS)
            .unwrap();
        assert_eq!(
            value(&room.left, 0, "C1"),
            CellValue::Error {
                value: ErrorValue::Name
            }
        );
        room.settle(delivery, "arrival");
        room.assert_both(|peer| {
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: 42.0 });
        });
    }
}

#[test]
fn a_redefinition_and_an_input_edit_settle_on_the_same_dependents() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .edit_cell(SheetId(0), cell("C1"), "=Rate*2", OPTIONS)
            .unwrap();
        room.settle(delivery, "formula");
        room.left
            .edit_cell(SheetId(0), cell("A1"), "100", OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![set("Rate", "Data!$A$1*3")], OPTIONS)
            .unwrap();
        room.settle(delivery, "redefined");
        room.assert_both(|peer| {
            assert_eq!(value(peer, 0, "C1"), CellValue::Number { value: 600.0 });
        });
    }
}

#[test]
fn a_function_shaped_name_saves_the_same_on_every_replica() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .edit_cell(SheetId(0), cell("C1"), "=ABS(-1)", OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![set("ABS", "2")], OPTIONS)
            .unwrap();
        room.settle(delivery, "function");
        room.left
            .edit_cell(SheetId(0), cell("C2"), "=ABS(-3)", OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "after");
    }
}

/// Keys and order values are not identities: an add undone, then added again
/// after reopening under the same client id, is one name with a fresh value.
#[test]
fn a_name_added_again_after_undo_and_reconnect_reaches_the_peer() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![set("Fresh", "1")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        let state = room.left.encode_state_as_update_v1();
        let mut resumed = Workbook::open_collaborative(&room.bytes, 1_001).unwrap();
        resumed.apply_update_v1(&state, OPTIONS).unwrap();
        room.left = resumed;
        room.left
            .apply_ops(vec![set("Fresh", "2")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "again");
        room.assert_both(|peer| assert_eq!(formula(peer, "Fresh", None).as_deref(), Some("2")));
    }
}

/// A name's scope is its sheet's key, which an undone add never hands to a later
/// sheet, even under the same client id.
#[test]
fn a_scoped_name_stays_with_its_undone_sheet_and_never_joins_a_later_one() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Notes"), set_in("Note", "1", Some(2))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        let state = room.left.encode_state_as_update_v1();
        let mut resumed = Workbook::open_collaborative(&room.bytes, 1_001).unwrap();
        resumed.apply_update_v1(&state, OPTIONS).unwrap();
        room.left = resumed;
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "again");
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other", "Notes"]);
            assert_eq!(formula(peer, "Note", Some(2)), None);
        });
    }
}

#[test]
fn a_scoped_name_from_the_adds_own_batch_goes_with_the_sheet_and_returns_with_it() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add(2, "Notes"), set_in("Note", "1", Some(2))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Note", Some(2)).as_deref(), Some("1"));
            assert_eq!(reopened(peer).defined_names, peer.model().defined_names);
        });
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other"]);
            assert!(!name_list(peer).iter().any(|name| name == "Note"));
        });
        room.redo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other", "Notes"]);
            assert_eq!(formula(peer, "Note", Some(2)).as_deref(), Some("1"));
        });
    }
}

#[test]
fn the_authors_later_scoped_name_undoes_before_the_add() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.left
            .apply_ops(vec![set_in("Note", "1", Some(2))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Note", Some(2)), None));
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(sheet_names(peer), ["Data", "Other"]));
    }
}

/// Refused like a peer's cell: nothing moves but the step leaving the stack.
fn assert_add_undo_refused(room: &mut Room) {
    let model = room.left.model().clone();
    let vector = room.left.encode_state_vector_v1();
    let state = room.left.encode_state_as_update_v1();
    let saved = room.left.save().unwrap();
    assert!(matches!(
        room.left.undo(OPTIONS),
        Err(Error::CollaborativeUndoRefused)
    ));
    assert_eq!(room.left.model(), &model);
    assert_eq!(room.left.encode_state_vector_v1(), vector);
    assert_eq!(room.left.encode_state_as_update_v1(), state);
    assert_eq!(room.left.save().unwrap(), saved);
}

#[test]
fn a_peers_scoped_name_refuses_the_add_undo_and_the_next_undo_runs() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![write(0, "C1", "earlier")], OPTIONS)
            .unwrap();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![set_in("Theirs", "1", Some(2))], OPTIONS)
            .unwrap();
        room.settle(delivery, "named");
        assert_add_undo_refused(&mut room);
        room.touch_right();
        room.settle(delivery, "refused");
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other", "Notes"]);
            assert_eq!(formula(peer, "Theirs", Some(2)).as_deref(), Some("1"));
        });
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(value(peer, 0, "C1"), CellValue::Empty));
    }
}

#[test]
fn a_workbook_name_pointing_at_the_sheet_does_not_refuse_the_add_undo() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![set("Global", "Notes!$A$1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "named");
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other"]);
            assert_eq!(formula(peer, "Global", None).as_deref(), Some("Notes!$A$1"));
        });
    }
}

/// Undo restores a value as a new struct under the undoing client, so the peer's
/// name still counts as content after passing through the author's history.
#[test]
fn a_peers_name_restored_by_the_authors_undo_still_refuses_the_add_undo() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![set_in("Theirs", "1", Some(2))], OPTIONS)
            .unwrap();
        room.settle(delivery, "named");
        room.left
            .apply_ops(vec![set_in("Theirs", "2", Some(2))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "overwritten");
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Theirs", Some(2)).as_deref(), Some("1"));
        });
        assert_add_undo_refused(&mut room);
        room.touch_right();
        room.settle(delivery, "refused");
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Theirs", Some(2)).as_deref(), Some("1"));
        });
    }
}

#[test]
fn a_peer_who_added_and_deleted_its_name_leaves_nothing_to_refuse() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![set_in("Theirs", "1", Some(2))], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(
                vec![Op::DeleteDefinedName {
                    name: "Theirs".into(),
                    scope: Some(SheetId(2)),
                }],
                OPTIONS,
            )
            .unwrap();
        room.settle(delivery, "cleared");
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(sheet_names(peer), ["Data", "Other"]));
    }
}

/// The known limit, kinder than a cell's: a name racing the undo stays stored,
/// hidden on every replica, and comes back when the sheet does.
#[test]
fn a_peers_name_racing_the_add_undo_is_hidden_then_returns_on_redo() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added");
        room.left.undo(OPTIONS).unwrap();
        room.right
            .apply_ops(vec![set_in("Racing", "1", Some(2))], OPTIONS)
            .unwrap();
        room.settle(delivery, "raced");
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other"]);
            assert!(!name_list(peer).iter().any(|name| name == "Racing"));
        });
        room.right
            .apply_ops(vec![set("Bystander", "2")], OPTIONS)
            .unwrap();
        room.settle(delivery, "another name while hidden");
        room.redo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other", "Notes"]);
            assert_eq!(formula(peer, "Racing", Some(2)).as_deref(), Some("1"));
        });
    }
}

#[test]
fn a_name_scoped_to_a_sheet_not_yet_arrived_shows_once_it_does() {
    let mut room = Room::open();
    let add_update = room
        .left
        .apply_ops(vec![add(2, "Notes")], OPTIONS)
        .map(|_| room.left.encode_state_as_update_v1())
        .unwrap();
    room.right.apply_update_v1(&add_update, OPTIONS).unwrap();
    let before = room.right.encode_state_vector_v1();
    room.right
        .apply_ops(vec![set_in("Theirs", "1", Some(2))], OPTIONS)
        .unwrap();
    let name_update = room.right.encode_diff_v1(&before).unwrap();

    let mut third = Workbook::open_collaborative(&room.bytes, 4_004).unwrap();
    third.apply_update_v1(&name_update, OPTIONS).unwrap();
    assert_eq!(sheet_names(&third), ["Data", "Other"]);
    assert!(!name_list(&third).iter().any(|name| name == "Theirs"));
    third.apply_update_v1(&add_update, OPTIONS).unwrap();
    assert_eq!(sheet_names(&third), ["Data", "Other", "Notes"]);
    assert_eq!(formula(&third, "Theirs", Some(2)).as_deref(), Some("1"));
    assert_eq!(third.model(), room.right.model());
}

#[test]
fn scoped_names_follow_their_sheets_whichever_way_two_appends_land() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.concurrently(
            vec![add(2, "Notes"), set_in("Mine", "1", Some(2))],
            vec![add(2, "Notes"), set_in("Mine", "2", Some(2))],
            delivery,
        );
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other", "Notes", "Notes (2)"]);
            let values = [
                formula(peer, "Mine", Some(2)).unwrap(),
                formula(peer, "Mine", Some(3)).unwrap(),
            ];
            let mut sorted = values.clone();
            sorted.sort();
            assert_eq!(sorted, ["1", "2"]);
        });
    }
}

#[test]
fn a_chart_move_writes_no_name_and_leaves_them_alone() {
    for delivery in DELIVERIES {
        let mut room = Room::of(CHARTED.to_vec());
        room.left
            .apply_ops(vec![set("Fresh", "1")], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![set("Other_One", "2")], OPTIONS)
            .unwrap();
        room.settle(delivery, "names");
        let names = room.left.model().defined_names.clone();
        let (sheet, frame) = room
            .left
            .model()
            .sheets
            .iter()
            .enumerate()
            .find_map(|(index, sheet)| Some((index, sheet.charts.first()?.frame_id())))
            .unwrap();
        let before = room.left.encode_state_vector_v1();
        room.left
            .move_chart(SheetId(sheet as u32), &frame, 16.0, 0.0, OPTIONS)
            .unwrap();
        let diff = room.left.encode_diff_v1(&before).unwrap();
        let doc = Doc::new();
        doc.transact_mut()
            .apply_update(Update::decode_v1(&room.right.encode_state_as_update_v1()).unwrap())
            .unwrap();
        let written = Arc::new(AtomicBool::new(false));
        let flag = written.clone();
        let _subscription = doc
            .get_or_insert_map(NAMES_ROOT)
            .observe(move |_, _| flag.store(true, Ordering::Relaxed));
        doc.transact_mut()
            .apply_update(Update::decode_v1(&diff).unwrap())
            .unwrap();
        assert!(!written.load(Ordering::Relaxed));
        room.touch_right();
        room.settle(delivery, "move");
        assert_eq!(room.right.model().defined_names, names);
    }
}

#[test]
fn an_ambiguous_source_opens_as_it_is_and_refuses_edits_without_writing() {
    let bytes = package(vec![
        defined("Rate", "1", None, false),
        defined("RATE", "2", None, false),
    ]);
    let standalone = Workbook::open(&bytes).unwrap();
    let mut room = Room::of(bytes);
    assert_eq!(
        room.left.model().defined_names,
        standalone.model().defined_names
    );
    let saved = room.left.save().unwrap();
    let state = room.left.encode_state_as_update_v1();
    for op in [set("Other", "3"), set("Rate", "3"), delete("Rate")] {
        assert!(room.left.apply_ops(vec![op], OPTIONS).is_err());
    }
    assert_eq!(room.left.encode_state_as_update_v1(), state);
    assert_eq!(room.left.save().unwrap(), saved);
    assert_eq!(room.left.save().unwrap(), standalone.save().unwrap());
    room.settle(Delivery::Merged, "untouched");
}

#[test]
fn a_room_nobody_named_anything_in_has_no_root_and_the_first_name_adds_it() {
    for delivery in DELIVERIES {
        let bytes = package(Vec::new());
        let mut room = Room::of(bytes);
        room.left
            .apply_ops(vec![write(0, "C1", "text")], OPTIONS)
            .unwrap();
        room.left.undo(OPTIONS).unwrap();
        assert!(!roots(&room.left.encode_state_as_update_v1()).contains(&NAMES_ROOT.to_string()));
        room.touch_right();
        room.settle(delivery, "unnamed");

        room.right
            .apply_ops(vec![set("First", "1")], OPTIONS)
            .unwrap();
        assert!(roots(&room.right.encode_state_as_update_v1()).contains(&NAMES_ROOT.to_string()));
        room.touch_left();
        room.settle(delivery, "named");
        room.assert_both(|peer| assert_eq!(name_list(peer), ["First"]));
    }
}

/// Deleting every name is an overlay of tombstones, not a missing table, so a
/// replica that only receives the state sees no names and no source fallback.
#[test]
fn a_table_emptied_by_deletes_stays_empty_through_a_state_round_trip() {
    for delivery in DELIVERIES {
        let mut room = Room::of(package(vec![
            defined("Rate", "1", None, false),
            defined("Tax", "2", None, false),
        ]));
        room.left
            .apply_ops(vec![delete("Rate"), delete("Tax")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "emptied");
        room.assert_both(|peer| assert!(peer.model().defined_names.is_empty()));
    }
}

#[test]
fn whole_table_writes_and_the_filter_database_name_stay_refused() {
    let mut room = Room::open();
    let model = room.left.model().clone();
    let state = room.left.encode_state_as_update_v1();
    for op in [
        Op::SetDefinedNames {
            defined_names: Vec::new(),
        },
        set_in("_xlnm._FilterDatabase", "Data!$A$1:$A$2", Some(0)),
        set_in("_XLNM._FILTERDATABASE", "Data!$A$1:$A$2", Some(0)),
        Op::DeleteDefinedName {
            name: "_xlnm._FilterDatabase".into(),
            scope: Some(SheetId(0)),
        },
    ] {
        assert!(matches!(
            room.left.apply_ops(vec![op], OPTIONS),
            Err(Error::CollaborativeStructureOperation)
        ));
        assert_eq!(room.left.model(), &model);
        assert_eq!(room.left.encode_state_as_update_v1(), state);
    }
}

#[test]
fn print_areas_are_ordinary_scoped_names() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(
                vec![set_in("_xlnm.Print_Area", "Data!$A$1:$C$3", Some(0))],
                OPTIONS,
            )
            .unwrap();
        room.touch_right();
        room.settle(delivery, "print area");
        room.assert_both(|peer| {
            assert_eq!(
                formula(peer, "_xlnm.Print_Area", Some(0)).as_deref(),
                Some("Data!$A$1:$C$3")
            );
        });
    }
}

#[test]
fn a_standalone_peers_name_update_stays_refused() {
    let bytes = source();
    let mut standalone = Workbook::open(&bytes).unwrap();
    standalone
        .apply_ops(vec![set("Rate", "5")], OPTIONS)
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

/// What no honest replica writes, arriving straight from the wire.
#[test]
fn entries_a_hostile_peer_writes_are_refused_without_a_change() {
    let json = |name: &str, formula: &str| {
        format!(r#"{{"name":"{name}","formula":"{formula}","hidden":false,"seq":0}}"#)
    };
    for (key, stored, reason) in [
        ("rate", json("Rate", "1"), "invalid defined name key"),
        (
            "set:nowhere/rate",
            json("Rate", "1"),
            "invalid defined name key",
        ),
        ("set:*/rate", "{".to_string(), "invalid defined name"),
        ("set:*/rate", json("Other", "1"), "holds Other"),
        (
            "set:*/bad",
            json("Bad", "=1"),
            "invalid defined name formula",
        ),
        (
            "set:*/a1",
            json("A1", "1"),
            "invalid or reserved defined name",
        ),
        (
            "set:sheet:0/_xlnm._filterdatabase",
            json("_xlnm._FilterDatabase", "1"),
            "not replicated",
        ),
    ] {
        let bytes = source();
        let mut peer = Workbook::open_collaborative(&bytes, 6_006).unwrap();
        let model = peer.model().clone();
        let state = peer.encode_state_as_update_v1();
        let hostile = Doc::with_client_id(9_000);
        hostile
            .transact_mut()
            .apply_update(Update::decode_v1(&state).unwrap())
            .unwrap();
        let before = hostile.transact().state_vector();
        let names = hostile.get_or_insert_map(NAMES_ROOT);
        names.insert(&mut hostile.transact_mut(), key, stored);
        let update = hostile.transact().encode_diff_v1(&before);
        let failure = peer
            .apply_update_v1(&update, OPTIONS)
            .unwrap_err()
            .to_string();
        assert!(failure.contains(reason), "{key}: {failure}");
        assert_eq!(peer.model(), &model);
        assert_eq!(peer.encode_state_as_update_v1(), state);
    }
}

fn entry(name: &str, formula: &str) -> String {
    format!(r#"{{"name":"{name}","formula":"{formula}","hidden":false,"seq":0}}"#)
}

/// What a peer who ignores the rules sends after writing `key` over `state`.
fn hostile_write(state: &[u8], key: &str, stored: impl Into<yrs::Any>) -> Vec<u8> {
    let hostile = Doc::with_client_id(9_000);
    hostile
        .transact_mut()
        .apply_update(Update::decode_v1(state).unwrap())
        .unwrap();
    let before = hostile.transact().state_vector();
    let names = hostile.get_or_insert_map(NAMES_ROOT);
    names.insert(&mut hostile.transact_mut(), key, stored.into());
    hostile.transact().encode_diff_v1(&before)
}

/// The stored key of the one sheet appended to `state`, read while it is shown:
/// an undo removes its entry from the sheet map.
fn appended_key(state: &[u8]) -> String {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(state).unwrap())
        .unwrap();
    let sheets = doc.get_or_insert_map("xlsx:sheets");
    let txn = doc.transact();
    let keys = sheets
        .keys(&txn)
        .filter(|key| key.starts_with("replica:"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    assert_eq!(keys.len(), 1);
    keys.into_iter().next().unwrap()
}

/// A redo shows what the undone sheet's scope holds, and no check runs on the
/// replica doing it, so a bad name must never get in while nothing shows it.
#[test]
fn a_bad_name_on_an_undone_sheet_is_refused_and_the_redo_converges() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added");
        let added = appended_key(&room.left.encode_state_as_update_v1());
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.settle(delivery, "undone");

        let state = room.left.encode_state_as_update_v1();
        let key = format!("set:{added}/bad");
        let hostile = hostile_write(&state, &key, entry("Bad", "=1"));
        let update = match delivery {
            Delivery::Merged => {
                room.touch_right();
                let honest = room
                    .right
                    .encode_diff_v1(&room.left.encode_state_vector_v1())
                    .unwrap();
                Update::merge_updates([&hostile, &honest].map(|u| Update::decode_v1(u).unwrap()))
                    .encode_v1()
            }
            _ => hostile,
        };
        let (left_state, right_state) = (
            room.left.encode_state_as_update_v1(),
            room.right.encode_state_as_update_v1(),
        );
        let (left_model, right_model) = (room.left.model().clone(), room.right.model().clone());
        let refused = |peer: &mut Workbook| {
            let failure = peer.apply_update_v1(&update, OPTIONS).unwrap_err();
            assert!(
                failure.to_string().contains("invalid defined name formula"),
                "{delivery:?}: {failure}"
            );
        };
        if matches!(delivery, Delivery::RightThenLeft) {
            refused(&mut room.right);
            refused(&mut room.left);
        } else {
            refused(&mut room.left);
            refused(&mut room.right);
        }
        assert_eq!(room.left.encode_state_as_update_v1(), left_state);
        assert_eq!(room.right.encode_state_as_update_v1(), right_state);
        assert_eq!(room.left.model(), &left_model);
        assert_eq!(room.right.model(), &right_model);

        room.redo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(sheet_names(peer), ["Data", "Other", "Notes"]);
            assert!(!name_list(peer).iter().any(|name| name == "Bad"));
        });
    }
}

/// Each way an entry can be stored and shown nowhere gets the answer the same
/// entry gets when shown, and nothing moves.
#[test]
fn entries_stored_where_nothing_shows_them_are_checked_like_shown_ones() {
    let mut room = Room::open();
    room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
    room.exchange(Delivery::LeftThenRight);
    let undone = appended_key(&room.left.encode_state_as_update_v1());
    assert!(room.left.undo(OPTIONS).unwrap().applied);
    room.exchange(Delivery::LeftThenRight);
    let state = room.right.encode_state_as_update_v1();
    let long = "1".repeat(xlsx_calc::lexer::MAX_FORMULA_BYTES + 1);
    let long_name = "n".repeat(256);

    for scope in [undone.as_str(), "replica:77:1:0", "sheet:99"] {
        for (folded, stored, reason) in [
            ("bad", entry("Bad", "=1"), "invalid defined name formula"),
            ("bad", entry("Bad", "1+"), "invalid defined name formula"),
            ("bad", entry("Bad", &long), "invalid defined name formula"),
            ("a1", entry("A1", "1"), "invalid or reserved defined name"),
            (
                long_name.as_str(),
                entry(&long_name, "1"),
                "invalid or reserved defined name",
            ),
            (
                "_xlnm.other",
                entry("_xlnm.Other", "1"),
                "invalid or reserved defined name",
            ),
            (
                "_xlnm._filterdatabase",
                entry("_xlnm._FilterDatabase", "1"),
                "not replicated",
            ),
            ("bad", "{".to_string(), "invalid defined name"),
            ("bad", entry("Other", "1"), "holds Other"),
        ] {
            let key = format!("set:{scope}/{folded}");
            let update = hostile_write(&state, &key, stored);
            let failure = room
                .right
                .apply_update_v1(&update, OPTIONS)
                .unwrap_err()
                .to_string();
            assert!(failure.contains(reason), "{key}: {failure}");
            assert_eq!(room.right.encode_state_as_update_v1(), state);
        }
        for (key, stored, reason) in [
            (
                format!("del:{scope}/gone"),
                yrs::Any::Bool(true),
                "no source",
            ),
            (
                format!("del:{scope}/gone"),
                yrs::Any::Bool(false),
                "not true",
            ),
            (
                format!("set:{scope}/bad"),
                yrs::Any::Bool(true),
                "not a string",
            ),
        ] {
            let update = hostile_write(&state, &key, stored);
            let failure = room
                .right
                .apply_update_v1(&update, OPTIONS)
                .unwrap_err()
                .to_string();
            assert!(failure.contains(reason), "{key}: {failure}");
            assert_eq!(room.right.encode_state_as_update_v1(), state);
        }
    }

    let update = hostile_write(&state, "del:*/rate", yrs::Any::Bool(false));
    let failure = room.right.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert!(failure.to_string().contains("not true"), "{failure}");

    let update = hostile_write(&state, "del:*/nosuch", yrs::Any::Bool(true));
    let failure = room.right.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert!(failure.to_string().contains("no source"), "{failure}");

    let key = format!("set:{undone}/fine");
    let update = hostile_write(&state, &key, entry("Fine", "1"));
    room.right.apply_update_v1(&update, OPTIONS).unwrap();
    room.left.apply_update_v1(&update, OPTIONS).unwrap();
    assert!(room.left.redo(OPTIONS).unwrap().applied);
    room.exchange(Delivery::LeftThenRight);
    room.assert_converged("a sound name returns with its sheet");
    assert_eq!(formula(&room.right, "Fine", Some(2)).as_deref(), Some("1"));
}

/// A set outranks a tombstone of its own name, so the tombstone is stored and
/// has no effect. It still needs a source name to stand on.
#[test]
fn a_tombstone_a_live_set_outranks_still_needs_a_source_name() {
    let mut room = Room::open();
    room.left
        .apply_ops(vec![set("Fresh", "1")], OPTIONS)
        .unwrap();
    room.exchange(Delivery::LeftThenRight);
    let state = room.right.encode_state_as_update_v1();
    let update = hostile_write(&state, "del:*/fresh", yrs::Any::Bool(true));
    let failure = room.right.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert!(failure.to_string().contains("no source"), "{failure}");
    assert_eq!(room.right.encode_state_as_update_v1(), state);

    let update = hostile_write(&state, "del:*/rate", yrs::Any::Bool(true));
    room.right
        .apply_ops(vec![set("Rate", "Data!$A$1+1")], OPTIONS)
        .unwrap();
    room.right.apply_update_v1(&update, OPTIONS).unwrap();
    room.left.apply_update_v1(&update, OPTIONS).unwrap();
    room.exchange(Delivery::Merged);
    room.assert_converged("the set stands over the tombstone");
    assert_eq!(
        formula(&room.left, "Rate", None).as_deref(),
        Some("Data!$A$1+1")
    );
}

/// A sheet's key can be guessed before the sheet exists. Whatever was stored
/// under it shows on every replica, its author's too, once the sheet is added.
#[test]
fn a_name_stored_under_a_sheet_key_before_the_sheet_exists_shows_everywhere() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let state = room.left.encode_state_as_update_v1();
        let clock = yrs::StateVector::decode_v1(&room.left.encode_state_vector_v1())
            .unwrap()
            .get(&yrs::block::ClientID::new(1_001));
        let key = format!("set:replica:1001:{clock}:0/early");
        let update = hostile_write(&state, &key, entry("Early", "1"));
        room.left.apply_update_v1(&update, OPTIONS).unwrap();
        room.right.apply_update_v1(&update, OPTIONS).unwrap();
        room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
        room.settle(delivery, "added under a claimed key");
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Early", Some(2)).as_deref(), Some("1"));
        });
    }
}

/// The redo that would show a name on such a source is one nothing can save.
#[test]
fn a_name_on_an_undone_sheet_is_refused_on_an_ambiguous_source_too() {
    let bytes = package(vec![
        defined("Rate", "1", None, false),
        defined("RATE", "2", None, false),
    ]);
    let mut room = Room::of(bytes);
    room.left.apply_ops(vec![add(2, "Notes")], OPTIONS).unwrap();
    room.exchange(Delivery::LeftThenRight);
    let added = appended_key(&room.left.encode_state_as_update_v1());
    assert!(room.left.undo(OPTIONS).unwrap().applied);
    room.exchange(Delivery::LeftThenRight);
    let state = room.left.encode_state_as_update_v1();
    let key = format!("set:{added}/fine");
    let update = hostile_write(&state, &key, entry("Fine", "1"));
    for peer in [&mut room.left, &mut room.right] {
        let before = peer.encode_state_as_update_v1();
        let failure = peer.apply_update_v1(&update, OPTIONS).unwrap_err();
        assert!(failure.to_string().contains("ambiguous"), "{failure}");
        assert_eq!(peer.encode_state_as_update_v1(), before);
    }
}

/// Every edit on such a source fails to save, so no peer may bring one in.
#[test]
fn a_new_name_on_an_ambiguous_source_is_refused_from_a_peer_too() {
    let bytes = package(vec![
        defined("Rate", "1", None, false),
        defined("RATE", "2", None, false),
    ]);
    let mut peer = Workbook::open_collaborative(&bytes, 6_006).unwrap();
    let state = peer.encode_state_as_update_v1();
    let hostile = Doc::with_client_id(9_000);
    hostile
        .transact_mut()
        .apply_update(Update::decode_v1(&state).unwrap())
        .unwrap();
    let before = hostile.transact().state_vector();
    let names = hostile.get_or_insert_map(NAMES_ROOT);
    names.insert(
        &mut hostile.transact_mut(),
        "set:*/fresh",
        r#"{"name":"Fresh","formula":"1","hidden":false,"seq":0}"#,
    );
    let update = hostile.transact().encode_diff_v1(&before);
    let failure = peer
        .apply_update_v1(&update, OPTIONS)
        .unwrap_err()
        .to_string();
    assert!(failure.contains("ambiguous"), "{failure}");
    assert_eq!(peer.encode_state_as_update_v1(), state);
}

/// A state written before names replicated holds no root, so it is an empty
/// overlay with nothing to migrate.
#[test]
fn a_room_from_before_names_replicated_is_an_empty_overlay() {
    for delivery in DELIVERIES {
        let restored = |client| {
            let mut workbook =
                Workbook::open_collaborative_recalculated(SAMPLE, client, OPTIONS).unwrap();
            assert!(
                workbook
                    .apply_update_v1(SAMPLE_V5, OPTIONS)
                    .unwrap()
                    .applied
            );
            workbook
        };
        let (left, right) = (restored(1_001), restored(2_002));
        assert_eq!(
            left.model().defined_names,
            Workbook::open(SAMPLE).unwrap().model().defined_names
        );
        let mut room = Room {
            bytes: SAMPLE.to_vec(),
            left,
            right,
            mirrors: 3_000,
        };
        room.exchange(delivery);
        room.concurrently(vec![set("Alpha", "1")], vec![set("Beta", "2")], delivery);
        room.assert_both(|peer| {
            assert_eq!(formula(peer, "Alpha", None).as_deref(), Some("1"));
            assert_eq!(formula(peer, "Beta", None).as_deref(), Some("2"));
        });
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(formula(peer, "Alpha", None), None));
    }
}

#[test]
fn a_recovery_checkpoint_keeps_the_overlay_and_the_source_baseline() {
    let mut room = Room::open();
    room.left
        .apply_ops(
            vec![set("Rate", "5"), delete("Secret"), set("Fresh", "1")],
            OPTIONS,
        )
        .unwrap();
    room.settle(Delivery::Merged, "named");
    let checkpoint = room.right.recovery_checkpoint(1, None).unwrap();
    let (mut restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(
        restored.model().defined_names,
        room.right.model().defined_names
    );
    restored.apply_ops(vec![delete("Rate")], OPTIONS).unwrap();
    let update = restored
        .encode_diff_v1(&room.left.encode_state_vector_v1())
        .unwrap();
    room.left.apply_update_v1(&update, OPTIONS).unwrap();
    assert_eq!(
        room.left.model().defined_names,
        restored.model().defined_names
    );
}

fn over_half_the_bound(prefix: &str) -> Vec<Op> {
    let formula = "x".repeat(xlsx_calc::lexer::MAX_FORMULA_BYTES);
    let count = xlsx_model::workbook::MAX_DEFINED_NAME_BYTES / 2 / formula.len() + 8;
    (0..count)
        .map(|index| set(&format!("{prefix}_{index}"), &formula))
        .collect()
}

fn over_limit(workbook: &Workbook) -> bool {
    workbook
        .list_defined_names_json("{}")
        .unwrap()
        .contains(r#""overLimit":true"#)
}

/// Each replica stays within the bound and their union does not, at the real
/// bound. Nobody is refused, nothing is lost, the room keeps saving, and only
/// growth stops. The bound is this engine's memory limit, so the file is whole
/// and its names are all there, though this engine's reader refuses to reopen
/// it until the table is back under the bound.
#[test]
fn a_union_over_the_size_bound_is_accepted_and_saved_and_only_growth_stops() {
    let mut room = Room::open();
    let left_names = over_half_the_bound("Left");
    let right_names = over_half_the_bound("Right");
    let added = left_names.len() + right_names.len();
    room.left.apply_ops(left_names, OPTIONS).unwrap();
    room.right.apply_ops(right_names.clone(), OPTIONS).unwrap();
    assert!(!over_limit(&room.left));
    room.exchange(Delivery::LeftThenRight);
    room.assert_same_document("union");
    room.assert_both(|peer| {
        assert_eq!(peer.model().defined_names.len(), added + 5);
        assert!(over_limit(peer));
    });

    let saved = room.left.save().unwrap();
    let xml = String::from_utf8(workbook_xml(&room.left)).unwrap();
    assert_eq!(xml.matches("<definedName ").count(), added + 5);
    let failure = Workbook::open(&saved).err().unwrap().to_string();
    assert!(
        failure.contains("defined name table exceeded cap"),
        "{failure}"
    );

    for op in [set("Grown", "1"), set("Rate", "Data!$A$1+1")] {
        let failure = room.left.apply_ops(vec![op], OPTIONS).unwrap_err();
        assert_eq!(failure.to_string(), "defined name table exceeded cap");
    }
    let failure = room
        .right
        .apply_ops(vec![set("Grown", "1")], OPTIONS)
        .unwrap_err();
    assert_eq!(failure.to_string(), "defined name table exceeded cap");
    room.left
        .apply_ops(vec![set("Left_0", "1")], OPTIONS)
        .unwrap();
    assert!(over_limit(&room.left));
    room.left
        .edit_cell(SheetId(0), cell("C1"), "still editing", OPTIONS)
        .unwrap();
    room.exchange(Delivery::RightThenLeft);
    room.assert_same_document("cells keep working");

    let checkpoint = room.right.recovery_checkpoint(1, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(
        restored.model().defined_names,
        room.right.model().defined_names
    );

    let shrink = right_names
        .iter()
        .map(|op| match op {
            Op::SetDefinedName { name, .. } => delete(name),
            _ => unreachable!(),
        })
        .collect::<Vec<_>>();
    room.right.apply_ops(shrink, OPTIONS).unwrap();
    room.exchange(Delivery::Merged);
    room.assert_converged("shrunk");
    room.assert_both(|peer| assert!(!over_limit(peer)));
    room.left
        .apply_ops(vec![set("Grown", "1")], OPTIONS)
        .unwrap();
}
