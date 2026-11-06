//! Two replicas editing conditional formats and auto filters while each also
//! receives the other's work. A sheet's blob stays the baseline every replica
//! opened, added rules and removals are single keys on the sheet's map, and a
//! filter's hidden name is derived, so nothing here stores a priority or a name.
//! Every scenario runs under each delivery and with each client in each role.

use std::collections::BTreeMap;

use betteroffice_xlsx::{
    CalculationOptions, DefinedName, Error, Op, Sheet, SheetId, Workbook, WorkbookModel,
};
use xlsx_model::{
    AutoFilter, CellRange, CellRef, ConditionalFormatRule, ConditionalFormatting,
    DifferentialFormat,
};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Any, Doc, Map, MapPrelim, MapRef, ReadTxn, Transact, Update};

const OPTIONS: CalculationOptions = CalculationOptions { now_serial: None };
const SHEETS_ROOT: &str = "xlsx:sheets";
const NAMES_ROOT: &str = "xlsx:defined-names";
const FILTER_NAME: &str = "_xlnm._FilterDATABASE";
const FROZEN: &str = "remote update changes the frozen workbook structure";
const SAMPLE: &[u8] = include_bytes!("../../../apps/demo/public/sample.xlsx");
const SAMPLE_V5: &[u8] = include_bytes!("fixtures/workbook-schema-v5-sample.update.bin");
const CHARTED: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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

fn package(names: &str) -> Vec<u8> {
    package_filtered(names, "A1:C9")
}

fn package_filtered(names: &str, filter: &str) -> Vec<u8> {
    let sheet = format!(
        r#"<worksheet xmlns:r="{REL}"><sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData><autoFilter ref="{filter}"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter><conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>5</formula></cfRule><cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1:B5"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule><cfRule type="expression" dxfId="0" priority="4"><formula>B1&gt;0</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="D1:D5"><cfRule type="expression" dxfId="0" priority="6"><formula>D1&gt;0</formula></cfRule></conditionalFormatting><tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>"#
    );
    let workbook = format!(
        r#"<workbook xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>{names}</workbook>"#
    );
    let parts: Vec<(String, Vec<u8>)> = vec![
        ("xl/workbook.xml".into(), workbook.into_bytes()),
        (
            "xl/_rels/workbook.xml.rels".into(),
            format!(r#"<Relationships><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/styles.xml".into(),
            br#"<styleSheet><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>"#.to_vec(),
        ),
        ("xl/worksheets/sheet1.xml".into(), sheet.into_bytes()),
        (
            "xl/worksheets/sheet2.xml".into(),
            br#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#.to_vec(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            format!(r#"<Relationships><Relationship Id="t1" Type="{REL}/table" Target="../tables/table1.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/tables/table1.xml".into(),
            br#"<table id="1" name="Items" displayName="Items" ref="H1:I5"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>"#.to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn source() -> Vec<u8> {
    package(&format!(
        r#"<definedNames><definedName name="{FILTER_NAME}" localSheetId="0" hidden="1">Data!$A$1:$C$9</definedName><definedName name="Rate">1</definedName></definedNames>"#
    ))
}

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn range(a1: &str) -> CellRange {
    CellRange::parse_a1(a1).unwrap()
}

fn rule_on(sheet: u32, sqref: &str, formula: &str) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(sheet),
        range: range(sqref),
        kind: "expression".into(),
        operator: None,
        formulas: vec![formula.into()],
        format: DifferentialFormat {
            bold: Some(true),
            ..Default::default()
        },
    }
}

fn rule(sqref: &str, formula: &str) -> Op {
    rule_on(0, sqref, formula)
}

fn remove_on(sheet: u32, id: &str) -> Op {
    Op::RemoveConditionalFormat {
        sheet: SheetId(sheet),
        id: id.into(),
    }
}

fn remove(id: &str) -> Op {
    remove_on(0, id)
}

fn filter_on(sheet: u32, a1: Option<&str>) -> Op {
    Op::SetAutoFilter {
        sheet: SheetId(sheet),
        range: a1.map(range),
    }
}

fn filter(a1: Option<&str>) -> Op {
    filter_on(0, a1)
}

fn add_sheet(index: usize, name: &str) -> Op {
    Op::AddSheet {
        index,
        name: name.into(),
    }
}

/// `(id, sqref, first formula, priority)` of every rule, by priority.
fn listed(peer: &Workbook, sheet: usize) -> Vec<(String, String, String, u32)> {
    let mut rules = peer.model().sheets[sheet]
        .conditional_formats
        .iter()
        .flat_map(|group| {
            group.rules.iter().map(move |rule| {
                (
                    rule.id.clone(),
                    group.sqref.clone(),
                    rule.formulas.first().cloned().unwrap_or_default(),
                    rule.priority,
                )
            })
        })
        .collect::<Vec<_>>();
    rules.sort_by_key(|rule| rule.3);
    rules
}

fn formulas(peer: &Workbook, sheet: usize) -> Vec<String> {
    listed(peer, sheet).into_iter().map(|rule| rule.2).collect()
}

fn id_of(peer: &Workbook, sheet: usize, formula: &str) -> String {
    listed(peer, sheet)
        .into_iter()
        .find(|rule| rule.2 == formula)
        .unwrap_or_else(|| panic!("no rule {formula}"))
        .0
}

fn holds(peer: &Workbook, sheet: usize, formula: &str) -> bool {
    formulas(peer, sheet).iter().any(|held| held == formula)
}

fn filter_range(peer: &Workbook, sheet: usize) -> Option<String> {
    peer.model().sheets[sheet]
        .auto_filter
        .as_ref()
        .map(|filter| filter.range.to_a1())
}

fn filter_entry(peer: &Workbook, sheet: u32) -> Option<DefinedName> {
    peer.model()
        .defined_names
        .iter()
        .find(|name| {
            name.local_sheet == Some(SheetId(sheet))
                && name.name.eq_ignore_ascii_case("_xlnm._FilterDatabase")
        })
        .cloned()
}

fn part(bytes: &[u8], path: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == path)
        .unwrap_or_else(|| panic!("no part {path}"))
        .1
}

fn defined_names_xml(bytes: &[u8]) -> String {
    let xml = String::from_utf8(part(bytes, "xl/workbook.xml")).unwrap();
    let start = xml.find("<definedNames").unwrap();
    let end = xml.find("</definedNames>").unwrap();
    xml[start..end].to_string()
}

fn hydrated(state: &[u8]) -> Doc {
    let doc = Doc::with_client_id(9_000);
    doc.transact_mut()
        .apply_update(Update::decode_v1(state).unwrap())
        .unwrap();
    doc
}

/// Every stored value of `sheet` a rule or a filter reaches.
fn stored_formats(peer: &Workbook, sheet: &str) -> BTreeMap<String, String> {
    let doc = hydrated(&peer.encode_state_as_update_v1());
    let sheets = doc.get_or_insert_map(SHEETS_ROOT);
    let txn = doc.transact();
    let map = sheets.get(&txn, sheet).unwrap().cast::<MapRef>().unwrap();
    map.iter(&txn)
        .filter(|(key, _)| {
            key.starts_with("cf:") || matches!(*key, "conditionalFormats" | "autoFilter")
        })
        .map(|(key, value)| (key.to_string(), format!("{value:?}")))
        .collect()
}

fn roots(peer: &Workbook) -> Vec<String> {
    let doc = hydrated(&peer.encode_state_as_update_v1());
    let txn = doc.transact();
    txn.root_refs().map(|(key, _)| key.to_string()).collect()
}

fn appended_key(peer: &Workbook) -> String {
    let doc = hydrated(&peer.encode_state_as_update_v1());
    let sheets = doc.get_or_insert_map(SHEETS_ROOT);
    let txn = doc.transact();
    let keys = sheets
        .keys(&txn)
        .filter(|key| key.starts_with("replica:"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    assert_eq!(keys.len(), 1);
    keys.into_iter().next().unwrap()
}

/// What a peer who ignores the rules sends after writing `key` under `sheet`.
fn hostile_entry(state: &[u8], sheet: &str, key: &str, stored: impl Into<Any>) -> Vec<u8> {
    let doc = hydrated(state);
    let before = doc.transact().state_vector();
    let sheets = doc.get_or_insert_map(SHEETS_ROOT);
    {
        let mut txn = doc.transact_mut();
        let map = sheets.get(&txn, sheet).unwrap().cast::<MapRef>().unwrap();
        map.insert(&mut txn, key, stored.into());
    }
    doc.transact().encode_diff_v1(&before)
}

/// Writes a whole sheet nothing lists under `sheet`, as a peer can for a sheet
/// that an undo removed or that has not arrived.
fn hostile_sheet(state: &[u8], sheet: &str, key: &str, stored: impl Into<Any>) -> Vec<u8> {
    let doc = hydrated(state);
    let before = doc.transact().state_vector();
    let sheets = doc.get_or_insert_map(SHEETS_ROOT);
    {
        let mut txn = doc.transact_mut();
        let map = sheets.insert(&mut txn, sheet, MapPrelim::default());
        for nested in ["colWidths", "contents", "rowHeights", "styles"] {
            map.insert(&mut txn, nested, MapPrelim::default());
        }
        map.insert(&mut txn, "name", "Hidden");
        map.insert(&mut txn, "freezePane", Any::Null);
        map.insert(&mut txn, "hyperlinks", "[]");
        map.insert(&mut txn, "merges", Any::Array(Vec::<Any>::new().into()));
        map.insert(&mut txn, key, stored.into());
    }
    doc.transact().encode_diff_v1(&before)
}

struct Room {
    bytes: Vec<u8>,
    left: Workbook,
    right: Workbook,
    left_client: u64,
    right_client: u64,
    mirrors: u64,
}

impl Room {
    fn of(bytes: Vec<u8>, (left_client, right_client): (u64, u64)) -> Self {
        let left = Workbook::open_collaborative(&bytes, left_client).unwrap();
        let right = Workbook::open_collaborative(&bytes, right_client).unwrap();
        Self {
            bytes,
            left,
            right,
            left_client,
            right_client,
            mirrors: 3_000,
        }
    }

    fn open() -> Self {
        Self::of(source(), (1_001, 2_002))
    }

    fn left_wins(&self) -> bool {
        self.left_client > self.right_client
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

    fn concurrently(&mut self, left: Vec<Op>, right: Vec<Op>, delivery: Delivery) {
        self.left.apply_ops(left, OPTIONS).unwrap();
        self.right.apply_ops(right, OPTIONS).unwrap();
        self.settle(delivery, "concurrent");
    }

    fn assert_both(&self, check: impl Fn(&Workbook)) {
        check(&self.left);
        check(&self.right);
    }

    fn touch_cell(&mut self) -> CellRef {
        self.mirrors += 1;
        cell(&format!("Z{}", self.mirrors))
    }

    /// Keeps the exchange two-way where only the left replica has an edit under
    /// test. The cell is on the second sheet, so the first keeps its bytes, and
    /// this is an undo step, so a scenario undoing on the right has no use for it.
    fn touch_right(&mut self) {
        let at = self.touch_cell();
        self.right
            .edit_cell(SheetId(1), at, "peer", OPTIONS)
            .unwrap();
    }

    fn undo_left(&mut self, delivery: Delivery) {
        assert!(self.left.undo(OPTIONS).unwrap().applied);
        self.settle(delivery, "left undo");
    }

    fn redo_left(&mut self, delivery: Delivery) {
        assert!(self.left.redo(OPTIONS).unwrap().applied);
        self.settle(delivery, "left redo");
    }

    fn undo_right(&mut self, delivery: Delivery) {
        assert!(self.right.undo(OPTIONS).unwrap().applied);
        self.settle(delivery, "right undo");
    }

    fn redo_right(&mut self, delivery: Delivery) {
        assert!(self.right.redo(OPTIONS).unwrap().applied);
        self.settle(delivery, "right redo");
    }

    fn undo_winner(&mut self, delivery: Delivery) {
        if self.left_wins() {
            self.undo_left(delivery);
        } else {
            self.undo_right(delivery);
        }
    }

    fn redo_winner(&mut self, delivery: Delivery) {
        if self.left_wins() {
            self.redo_left(delivery);
        } else {
            self.redo_right(delivery);
        }
    }

    /// The step has nothing live left, so nothing moves.
    fn undo_loser_finds_nothing(&mut self, delivery: Delivery) {
        let loser = if self.left_wins() {
            &mut self.right
        } else {
            &mut self.left
        };
        assert!(!loser.undo(OPTIONS).unwrap().applied);
        self.settle(delivery, "loser undo");
    }
}

fn each_room(mut scenario: impl FnMut(&mut Room, Delivery)) {
    for delivery in DELIVERIES {
        for clients in [(1_001, 2_002), (2_002, 1_001)] {
            let mut room = Room::of(source(), clients);
            scenario(&mut room, delivery);
        }
    }
}

fn reopened(workbook: &Workbook) -> WorkbookModel {
    Workbook::open(&workbook.save().unwrap())
        .unwrap()
        .model()
        .clone()
}

/// A rename that rewrites formulas writes one register per reference. A cell's
/// winner follows that cell's own history, so with a third client two renames
/// can pick different winners for the name and for a formula naming the sheet.
/// Client 2 edits the cell and client 5 sees it before renaming; client 3 does
/// not. With client 4 in that place, no split.
#[test]
fn three_clients_split_the_winners_of_two_keys_two_renames_write() {
    let write = |doc: &Doc, key: &str, value: &str| {
        let root = doc.get_or_insert_map("sheet");
        root.insert(&mut doc.transact_mut(), key, value);
    };
    let read = |doc: &Doc, key: &str| -> String {
        let root = doc.get_or_insert_map("sheet");
        let txn = doc.transact();
        root.get(&txn, key).unwrap().cast::<String>().unwrap()
    };
    for (third, expected_cell) in [(2, "=B!A1"), (4, "=A!A1")] {
        let seed = Doc::with_client_id(1);
        write(&seed, "name", "Data");
        write(&seed, "cell", "=Data!A1");
        let base = seed
            .transact()
            .encode_state_as_update_v1(&yrs::StateVector::default());
        let replica = |client: u64| {
            let doc = Doc::with_client_id(client);
            doc.transact_mut()
                .apply_update(Update::decode_v1(&base).unwrap())
                .unwrap();
            doc
        };
        let (other, late, early) = (replica(third), replica(5), replica(3));
        let known = seed.transact().state_vector();
        write(&other, "cell", "=Data!A2");
        let edit = other.transact().encode_diff_v1(&known);
        late.transact_mut()
            .apply_update(Update::decode_v1(&edit).unwrap())
            .unwrap();
        write(&late, "name", "A");
        write(&late, "cell", "=A!A1");
        write(&early, "name", "B");
        write(&early, "cell", "=B!A1");
        let (renamed_late, renamed_early) = (
            late.transact().encode_diff_v1(&known),
            early.transact().encode_diff_v1(&known),
        );
        for order in [
            [&edit, &renamed_late, &renamed_early],
            [&edit, &renamed_early, &renamed_late],
            [&renamed_early, &edit, &renamed_late],
            [&renamed_early, &renamed_late, &edit],
            [&renamed_late, &edit, &renamed_early],
            [&renamed_late, &renamed_early, &edit],
        ] {
            let merged = replica(9);
            for update in order {
                merged
                    .transact_mut()
                    .apply_update(Update::decode_v1(update).unwrap())
                    .unwrap();
            }
            assert_eq!(read(&merged, "name"), "A", "third client {third}");
            assert_eq!(read(&merged, "cell"), expected_cell, "third client {third}");
        }
    }
}

#[test]
fn rules_both_peers_add_all_land_in_one_order() {
    each_room(|room, delivery| {
        room.concurrently(
            vec![rule("A1:A5", "A1>1"), rule("F1:F5", "F1>1")],
            vec![rule("A1:A5", "A1>2")],
            delivery,
        );
        let first = if room.left_client < room.right_client {
            ["A1>1", "A1>2", "F1>1"]
        } else {
            ["A1>2", "A1>1", "F1>1"]
        };
        room.assert_both(|peer| {
            let rules = listed(peer, 0);
            let added = rules[5..]
                .iter()
                .map(|rule| rule.2.as_str())
                .collect::<Vec<_>>();
            assert_eq!(added, first);
            assert_eq!(
                rules.iter().map(|rule| rule.3).collect::<Vec<_>>(),
                [1, 2, 3, 4, 6, 7, 8, 9]
            );
        });
        room.undo_left(delivery);
        room.assert_both(|peer| {
            let rules = listed(peer, 0);
            assert_eq!(rules.len(), 6);
            assert_eq!(rules[5].2, "A1>2");
            assert_eq!(rules[5].3, 7);
        });
        room.redo_left(delivery);
        room.undo_right(delivery);
        room.assert_both(|peer| assert_eq!(formulas(peer, 0).len(), 7));
        room.redo_right(delivery);
        room.assert_both(|peer| assert_eq!(formulas(peer, 0).len(), 8));
    });
}

/// `seq` keeps the order rules were added in across clients, where the ids
/// alone would group each client's rules.
#[test]
fn rules_added_in_turn_keep_that_order_whichever_client_added_them() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "first");
        room.right
            .apply_ops(vec![rule("F1", "F1>2")], OPTIONS)
            .unwrap();
        room.settle(delivery, "second");
        room.left
            .apply_ops(vec![rule("F1", "F1>3")], OPTIONS)
            .unwrap();
        room.settle(delivery, "third");
        room.assert_both(|peer| {
            assert_eq!(formulas(peer, 0)[5..], ["F1>1", "F1>2", "F1>3"]);
        });
    });
}

#[test]
fn a_rule_against_a_cell_edit_equals_the_sequential_result() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let left = vec![rule("A1:A5", "A1>1")];
        room.left.apply_ops(left.clone(), OPTIONS).unwrap();
        room.right
            .edit_cell(SheetId(0), cell("C3"), "7", OPTIONS)
            .unwrap();
        room.settle(delivery, "rule and cell");
        let mut standalone = Workbook::open(&room.bytes).unwrap();
        standalone.apply_ops(left, OPTIONS).unwrap();
        standalone
            .edit_cell(SheetId(0), cell("C3"), "7", OPTIONS)
            .unwrap();
        let expected = reopened(&standalone);
        assert_eq!(
            reopened(&room.left).sheets[0].conditional_formats,
            expected.sheets[0].conditional_formats,
            "{delivery:?}"
        );
        room.undo_left(delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "A1>1")));
        room.redo_left(delivery);
        room.assert_both(|peer| assert!(holds(peer, 0, "A1>1")));
    }
}

#[test]
fn removing_a_baseline_rule_against_an_add_keeps_both_effects() {
    each_room(|room, delivery| {
        room.concurrently(vec![remove("cf1")], vec![rule("A1:A5", "A1>1")], delivery);
        room.assert_both(|peer| {
            assert_eq!(formulas(peer, 0).len(), 5);
            assert!(!holds(peer, 0, "5"));
            assert!(holds(peer, 0, "A1>1"));
            assert_eq!(listed(peer, 0).last().unwrap().3, 7);
        });
        room.undo_left(delivery);
        room.assert_both(|peer| assert!(holds(peer, 0, "5")));
        room.undo_right(delivery);
        room.assert_both(|peer| assert_eq!(formulas(peer, 0).len(), 5));
        room.redo_right(delivery);
        room.redo_left(delivery);
        room.assert_both(|peer| {
            assert!(!holds(peer, 0, "5"));
            assert!(holds(peer, 0, "A1>1"));
        });
    });
}

/// One tombstone key, written twice: the winner's undo empties it, so the rule
/// returns though the other peer also removed it.
#[test]
fn two_peers_removing_one_baseline_rule_share_one_register() {
    each_room(|room, delivery| {
        room.concurrently(vec![remove("cf2")], vec![remove("cf2")], delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "A1>0")));
        room.undo_winner(delivery);
        room.assert_both(|peer| assert!(holds(peer, 0, "A1>0")));
        room.undo_loser_finds_nothing(delivery);
        room.assert_both(|peer| assert!(holds(peer, 0, "A1>0")));
        room.redo_winner(delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "A1>0")));
    });
    each_room(|room, delivery| {
        room.concurrently(vec![remove("cf2")], vec![remove("cf2")], delivery);
        room.undo_loser_finds_nothing(delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "A1>0")));
        room.undo_winner(delivery);
        room.assert_both(|peer| assert!(holds(peer, 0, "A1>0")));
    });
}

#[test]
fn two_peers_removing_one_added_rule_settle_on_one_copy() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        let id = id_of(&room.left, 0, "F1>1");
        room.concurrently(vec![remove(&id)], vec![remove(&id)], delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "F1>1")));
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(id_of(peer, 0, "F1>1"), id);
        });
        room.right.undo(OPTIONS).unwrap();
        room.settle(delivery, "the other undo");
        room.assert_both(|peer| {
            let held = formulas(peer, 0);
            assert_eq!(held.iter().filter(|held| *held == "F1>1").count(), 1);
        });
    });
}

/// A step with nothing live left falls through, and the peer's own removal stays
/// undoable. The restored rule keeps its place, which its stored `seq` carries.
#[test]
fn an_add_a_peer_removed_falls_through_and_the_peer_can_restore_it() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![rule("F1", "F1>2")], OPTIONS)
            .unwrap();
        room.settle(delivery, "later rule");
        let id = id_of(&room.right, 0, "F1>1");
        room.right.apply_ops(vec![remove(&id)], OPTIONS).unwrap();
        room.settle(delivery, "removed");
        assert!(!room.left.undo(OPTIONS).unwrap().applied);
        room.settle(delivery, "nothing to undo");
        room.assert_both(|peer| assert_eq!(formulas(peer, 0)[5..], ["F1>2"]));
        room.undo_right(delivery);
        room.assert_both(|peer| assert_eq!(formulas(peer, 0)[5..], ["F1>1", "F1>2"]));
    });
}

#[test]
fn undoing_a_removal_restores_the_source_bytes() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        let source = part(&room.bytes, "xl/worksheets/sheet1.xml");
        room.left.apply_ops(vec![remove("cf1")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "removed");
        assert_ne!(
            part(&room.left.save().unwrap(), "xl/worksheets/sheet1.xml"),
            source
        );
        room.undo_left(delivery);
        assert_eq!(
            part(&room.left.save().unwrap(), "xl/worksheets/sheet1.xml"),
            source,
            "{delivery:?}"
        );
        assert!(
            stored_formats(&room.left, "sheet:0")
                .keys()
                .all(|key| !key.starts_with("cf:"))
        );
    }
}

/// The op previews `cf{max + 1}` over live ids, which is the id of a removed top
/// rule, so an id taken from the preview would land on its tombstone.
#[test]
fn a_rule_added_after_removing_the_top_baseline_rule_keeps_both() {
    each_room(|room, delivery| {
        room.left.apply_ops(vec![remove("cf5")], OPTIONS).unwrap();
        room.left
            .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "separate batches");
        room.assert_both(|peer| {
            assert!(!holds(peer, 0, "D1>0"));
            assert!(id_of(peer, 0, "F1>1").starts_with("cf:"));
            assert_eq!(listed(peer, 0).last().unwrap().3, 7);
        });
        assert_eq!(
            stored_formats(&room.left, "sheet:0")
                .get("cf:cf5")
                .map(String::as_str),
            Some("Any(Bool(false))")
        );
        room.undo_left(delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "F1>1")));
        room.undo_left(delivery);
        room.assert_both(|peer| assert!(holds(peer, 0, "D1>0")));
    });
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![remove("cf5"), rule("D1:D5", "D1>1")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "one batch");
        room.assert_both(|peer| {
            assert!(!holds(peer, 0, "D1>0"));
            assert!(id_of(peer, 0, "D1>1").starts_with("cf:"));
        });
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert!(holds(peer, 0, "D1>0"));
            assert!(!holds(peer, 0, "D1>1"));
        });
    });
}

#[test]
fn a_rule_added_again_after_undo_and_reconnect_gets_a_fresh_key() {
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        let first = id_of(&room.left, 0, "F1>1");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        let state = room.left.encode_state_as_update_v1();
        let mut resumed = Workbook::open_collaborative(&room.bytes, 1_001).unwrap();
        resumed.apply_update_v1(&state, OPTIONS).unwrap();
        room.left = resumed;
        room.left
            .apply_ops(vec![rule("F1", "F1>2")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "again");
        room.assert_both(|peer| {
            assert!(!holds(peer, 0, "F1>1"));
            assert_ne!(id_of(peer, 0, "F1>2"), first);
        });
    }
}

#[test]
fn removing_a_rule_from_a_source_group_keeps_its_neighbours_and_a_data_bar_stays() {
    each_room(|room, delivery| {
        room.left.apply_ops(vec![remove("cf4")], OPTIONS).unwrap();
        room.touch_right();
        room.settle(delivery, "removed");
        let saved = String::from_utf8(part(&room.left.save().unwrap(), "xl/worksheets/sheet1.xml"))
            .unwrap();
        assert!(saved.contains(r#"<cfRule type="dataBar" priority="2">"#));
        assert!(!saved.contains("B1&gt;0"));
        let state = room.left.encode_state_as_update_v1();
        let model = room.left.model().clone();
        assert!(room.left.apply_ops(vec![remove("cf3")], OPTIONS).is_err());
        assert_eq!(room.left.model(), &model);
        assert_eq!(room.left.encode_state_as_update_v1(), state);
        let update = hostile_entry(&state, "sheet:0", "cf:cf3", false);
        let failure = room.right.apply_update_v1(&update, OPTIONS).unwrap_err();
        assert!(failure.to_string().contains("not editable"), "{failure}");
    });
}

/// The sheet and the rule are one step, so they leave and return together.
#[test]
fn a_rule_in_the_batch_that_adds_its_sheet_goes_and_returns_with_it() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(
                vec![add_sheet(2, "Notes"), rule_on(2, "A1:A5", "A1>1")],
                OPTIONS,
            )
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        room.assert_both(|peer| assert_eq!(formulas(peer, 2), ["A1>1"]));
        let key = appended_key(&room.left);
        assert!(!stored_formats(&room.left, &key).contains_key("conditionalFormats"));
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(peer.model().sheets.len(), 2));
        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(formulas(peer, 2), ["A1>1"]));
    });
}

#[test]
fn a_peers_rule_refuses_the_undo_of_its_sheet() {
    each_room(|room, delivery| {
        room.left
            .edit_cell(SheetId(0), cell("C1"), "first", OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![rule_on(2, "A1", "A1>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer rule");
        let model = room.left.model().clone();
        let state = room.left.encode_state_as_update_v1();
        assert!(matches!(
            room.left.undo(OPTIONS),
            Err(Error::CollaborativeUndoRefused)
        ));
        assert_eq!(room.left.model(), &model);
        assert_eq!(room.left.encode_state_as_update_v1(), state);
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(peer.cell(SheetId(0), cell("C1")).unwrap().input, "");
            assert_eq!(formulas(peer, 2), ["A1>1"]);
        });
    });
}

/// A value an undo restores is a new struct, so it does not count as the add's
/// own, whoever wrote it first.
#[test]
fn a_peers_rule_restored_by_the_author_still_refuses_the_undo_of_the_sheet() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![rule_on(2, "A1", "A1>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer rule");
        let id = id_of(&room.left, 2, "A1>1");
        room.left
            .apply_ops(vec![remove_on(2, &id)], OPTIONS)
            .unwrap();
        room.settle(delivery, "removed");
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(formulas(peer, 2), ["A1>1"]));
        assert!(matches!(
            room.left.undo(OPTIONS),
            Err(Error::CollaborativeUndoRefused)
        ));
        room.settle(delivery, "refused");
    });
}

#[test]
fn a_rule_the_peer_added_and_removed_leaves_nothing_to_refuse() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![rule_on(2, "A1", "A1>1")], OPTIONS)
            .unwrap();
        let id = id_of(&room.right, 2, "A1>1");
        room.right
            .apply_ops(vec![remove_on(2, &id)], OPTIONS)
            .unwrap();
        room.settle(delivery, "added and removed");
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(peer.model().sheets.len(), 2));
    });
}

/// A rule lives under its sheet's map, so like a cell it does not survive a
/// race with the undo of that sheet: it is gone from every replica and a redo
/// brings the sheet back without it.
#[test]
fn a_rule_racing_the_undo_of_its_sheet_is_lost_alike_on_both() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.right
            .apply_ops(vec![rule_on(2, "A1", "A1>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "raced");
        room.assert_both(|peer| assert_eq!(peer.model().sheets.len(), 2));
        room.redo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(peer.model().sheets.len(), 3);
            assert!(formulas(peer, 2).is_empty());
        });
    });
}

#[test]
fn a_chart_move_writes_neither_rules_nor_filters() {
    for delivery in DELIVERIES {
        let mut room = Room::of(CHARTED.to_vec(), (1_001, 2_002));
        room.left
            .apply_ops(vec![rule("F1", "F1>1"), filter(Some("A1:B4"))], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![rule("F2", "F2>1")], OPTIONS)
            .unwrap();
        room.settle(delivery, "formats");
        let stored = stored_formats(&room.left, "sheet:0");
        let frame = room.left.model().sheets[0].charts[0].frame_id();
        room.left
            .move_chart(SheetId(0), &frame, 18.0, 9.0, OPTIONS)
            .unwrap();
        room.settle(delivery, "chart moved");
        assert_eq!(stored_formats(&room.left, "sheet:0"), stored);
        room.undo_left(delivery);
        assert_eq!(stored_formats(&room.left, "sheet:0"), stored);
    }
}

#[test]
fn a_filter_set_by_both_peers_has_one_winner_and_a_name_that_follows_it() {
    each_room(|room, delivery| {
        room.concurrently(
            vec![filter(Some("A1:D9"))],
            vec![filter(Some("A1:E9"))],
            delivery,
        );
        let (range, formula) = if room.left_wins() {
            ("A1:D9", "'Data'!$A$1:$D$9")
        } else {
            ("A1:E9", "'Data'!$A$1:$E$9")
        };
        room.assert_both(|peer| {
            assert_eq!(filter_range(peer, 0).as_deref(), Some(range));
            let name = filter_entry(peer, 0).unwrap();
            assert_eq!(name.formula, formula);
            assert_eq!(name.name, FILTER_NAME);
            assert!(name.hidden);
        });
        assert!(
            defined_names_xml(&room.left.save().unwrap())
                .contains(formula.split_once('!').unwrap().1)
        );
        assert!(!roots(&room.left).contains(&NAMES_ROOT.to_string()));
    });
}

#[test]
fn a_filter_set_against_a_clear_is_one_register_the_higher_client_wins() {
    each_room(|room, delivery| {
        room.concurrently(vec![filter(Some("A1:D9"))], vec![filter(None)], delivery);
        let set_wins = room.left_wins();
        room.assert_both(|peer| {
            if set_wins {
                assert_eq!(filter_range(peer, 0).as_deref(), Some("A1:D9"));
                assert_eq!(filter_entry(peer, 0).unwrap().formula, "'Data'!$A$1:$D$9");
            } else {
                assert_eq!(filter_range(peer, 0), None);
                assert_eq!(filter_entry(peer, 0), None);
            }
        });
    });
}

#[test]
fn undoing_a_filter_edit_restores_the_source_filter_and_names() {
    for (edit, reads) in [(Some("A1:D9"), Some("A1:D9")), (None, None)] {
        for delivery in DELIVERIES {
            let mut room = Room::open();
            let names = defined_names_xml(&room.bytes);
            let sheet = part(&room.bytes, "xl/worksheets/sheet1.xml");
            room.left.apply_ops(vec![filter(edit)], OPTIONS).unwrap();
            room.touch_right();
            room.settle(delivery, "edited");
            room.assert_both(|peer| assert_eq!(filter_range(peer, 0).as_deref(), reads));
            room.undo_left(delivery);
            room.assert_both(|peer| {
                assert_eq!(filter_range(peer, 0).as_deref(), Some("A1:C9"));
            });
            let saved = room.left.save().unwrap();
            assert_eq!(defined_names_xml(&saved), names, "{delivery:?}");
            assert_eq!(
                part(&saved, "xl/worksheets/sheet1.xml"),
                sheet,
                "{delivery:?}"
            );
            room.redo_left(delivery);
            room.assert_both(|peer| assert_eq!(filter_range(peer, 0).as_deref(), reads));
        }
    }
}

/// The winner's undo empties the register, so it reads as the source's filter,
/// not the loser's.
#[test]
fn the_winner_undoing_a_filter_shows_the_source_not_the_loser() {
    each_room(|room, delivery| {
        room.concurrently(
            vec![filter(Some("A1:D9"))],
            vec![filter(Some("A1:E9"))],
            delivery,
        );
        room.undo_winner(delivery);
        room.assert_both(|peer| {
            assert_eq!(filter_range(peer, 0).as_deref(), Some("A1:C9"));
            assert_eq!(filter_entry(peer, 0).unwrap().formula, "Data!$A$1:$C$9");
        });
        room.undo_loser_finds_nothing(delivery);
        room.redo_winner(delivery);
        room.assert_both(|peer| assert_ne!(filter_range(peer, 0).as_deref(), Some("A1:C9")));
    });
    each_room(|room, delivery| {
        room.concurrently(
            vec![filter(Some("A1:D9"))],
            vec![filter(Some("A1:E9"))],
            delivery,
        );
        let standing = filter_range(&room.left, 0);
        room.undo_loser_finds_nothing(delivery);
        room.assert_both(|peer| assert_eq!(filter_range(peer, 0), standing));
    });
}

#[test]
fn a_filter_adds_its_name_or_edits_the_one_the_source_holds() {
    each_room(|room, delivery| {
        room.concurrently(
            vec![filter_on(1, Some("A1:B4"))],
            vec![filter(Some("A1:D9"))],
            delivery,
        );
        room.assert_both(|peer| {
            let names = &peer.model().defined_names;
            assert_eq!(names[0].name, FILTER_NAME);
            assert_eq!(names[0].formula, "'Data'!$A$1:$D$9");
            assert_eq!(names[1].name, "Rate");
            assert_eq!(names[2].name, "_xlnm._FilterDatabase");
            assert_eq!(names[2].local_sheet, Some(SheetId(1)));
            assert_eq!(names[2].formula, "'Other'!$A$1:$B$4");
            assert!(names[2].hidden);
        });
        assert_eq!(
            reopened(&room.left).defined_names.len(),
            room.left.model().defined_names.len()
        );
    });
}

#[test]
fn a_filter_and_a_name_edit_both_stand_and_only_the_name_is_stored() {
    each_room(|room, delivery| {
        room.concurrently(
            vec![filter(Some("A1:D9"))],
            vec![Op::SetDefinedName {
                name: "Rate".into(),
                refers_to: "2".into(),
                scope: None,
                hidden: None,
            }],
            delivery,
        );
        room.assert_both(|peer| {
            assert_eq!(filter_range(peer, 0).as_deref(), Some("A1:D9"));
            let rate = peer
                .model()
                .defined_names
                .iter()
                .find(|name| name.name == "Rate")
                .unwrap();
            assert_eq!(rate.formula, "2");
        });
        let doc = hydrated(&room.left.encode_state_as_update_v1());
        let names = doc.get_or_insert_map(NAMES_ROOT);
        let txn = doc.transact();
        assert_eq!(names.keys(&txn).collect::<Vec<_>>(), ["set:*/rate"]);
    });
}

fn rate(formula: &str) -> Op {
    Op::SetDefinedName {
        name: "Rate".into(),
        refers_to: formula.into(),
        scope: None,
        hidden: None,
    }
}

fn stored_names(peer: &Workbook) -> Vec<String> {
    let doc = hydrated(&peer.encode_state_as_update_v1());
    let names = doc.get_or_insert_map(NAMES_ROOT);
    let txn = doc.transact();
    names.keys(&txn).map(str::to_string).collect()
}

/// A commit that writes names compares the table it leaves with the stored one,
/// and the filter's derived name is in the first and never in the second.
#[test]
fn a_name_written_beside_a_filter_stores_only_the_name() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![filter(Some("A1:D9")), rate("2")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "set in one batch");
        room.assert_both(|peer| {
            assert_eq!(stored_names(peer), ["set:*/rate"]);
            let entry = filter_entry(peer, 0).unwrap();
            assert!(entry.formula.ends_with("!$A$1:$D$9"), "{entry:?}");
        });
        room.right
            .apply_ops(vec![rate("3"), filter(None)], OPTIONS)
            .unwrap();
        room.settle(delivery, "cleared in one batch");
        room.assert_both(|peer| {
            assert_eq!(stored_names(peer), ["set:*/rate"]);
            assert!(filter_entry(peer, 0).is_none());
        });
    });
    for delivery in DELIVERIES {
        for clients in [(1_001, 2_002), (2_002, 1_001)] {
            let mut room = Room::of(package(""), clients);
            room.left
                .apply_ops(vec![filter(Some("A1:D9"))], OPTIONS)
                .unwrap();
            room.touch_right();
            room.settle(delivery, "filtered");
            room.right.apply_ops(vec![rate("2")], OPTIONS).unwrap();
            room.settle(delivery, "named by the replica that received the filter");
            room.left.apply_ops(vec![rate("3")], OPTIONS).unwrap();
            room.settle(delivery, "named by the replica that set the filter");
            room.assert_both(|peer| {
                assert_eq!(stored_names(peer), ["set:*/rate"]);
                assert!(filter_entry(peer, 0).is_some());
            });
        }
    }
}

#[test]
fn a_filter_in_a_room_without_names_adds_no_names_root() {
    for delivery in DELIVERIES {
        let mut room = Room::of(package(""), (1_001, 2_002));
        room.left
            .apply_ops(vec![filter(Some("A1:D9"))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "filtered");
        assert!(!roots(&room.left).contains(&NAMES_ROOT.to_string()));
        room.assert_both(|peer| assert!(filter_entry(peer, 0).is_some()));
    }
}

#[test]
fn a_peers_filter_refuses_the_undo_of_its_sheet_and_the_authors_goes() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![filter_on(2, Some("A1:B4"))], OPTIONS)
            .unwrap();
        room.settle(delivery, "peer filter");
        room.assert_both(|peer| {
            assert_eq!(filter_entry(peer, 2).unwrap().formula, "'Notes'!$A$1:$B$4");
        });
        assert!(matches!(
            room.left.undo(OPTIONS),
            Err(Error::CollaborativeUndoRefused)
        ));
    });
    each_room(|room, delivery| {
        room.left
            .apply_ops(
                vec![add_sheet(2, "Notes"), filter_on(2, Some("A1:B4"))],
                OPTIONS,
            )
            .unwrap();
        room.touch_right();
        room.settle(delivery, "own filter");
        room.undo_left(delivery);
        room.assert_both(|peer| {
            assert_eq!(peer.model().sheets.len(), 2);
            assert_eq!(filter_entry(peer, 2), None);
        });
        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(filter_range(peer, 2).as_deref(), Some("A1:B4")));
    });
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.left
            .apply_ops(vec![filter_on(2, Some("A1:C4"))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "own later filter");
        room.undo_left(delivery);
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(peer.model().sheets.len(), 2));
    });
    each_room(|room, delivery| {
        room.left
            .apply_ops(
                vec![add_sheet(2, "Notes"), filter_on(2, Some("A1:B4"))],
                OPTIONS,
            )
            .unwrap();
        room.left
            .apply_ops(vec![filter_on(2, Some("A1:C4"))], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "own filter overwritten");
        room.undo_left(delivery);
        assert!(matches!(
            room.left.undo(OPTIONS),
            Err(Error::CollaborativeUndoRefused)
        ));
        room.settle(delivery, "restored value refuses, as a cell does");
    });
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.settle(delivery, "added");
        room.right
            .apply_ops(vec![filter_on(2, Some("A1:B4"))], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![filter_on(2, None)], OPTIONS)
            .unwrap();
        room.settle(delivery, "set and cleared");
        room.undo_left(delivery);
        room.assert_both(|peer| assert_eq!(peer.model().sheets.len(), 2));
    });
}

fn stored_filter(filter: AutoFilter) -> String {
    serde_json::to_string(&Some(filter)).unwrap()
}

fn bare_filter(a1: &str) -> AutoFilter {
    AutoFilter {
        range: range(a1),
        raw_xml: None,
        source_count: 1,
        filter_columns: Vec::new(),
        has_sort_state: false,
    }
}

#[test]
fn a_filter_over_a_table_is_refused_locally_and_from_the_wire() {
    let mut room = Room::open();
    let state = room.left.encode_state_as_update_v1();
    assert!(
        room.left
            .apply_ops(vec![filter(Some("H1:I3"))], OPTIONS)
            .is_err()
    );
    assert_eq!(room.left.encode_state_as_update_v1(), state);
    let update = hostile_entry(
        &state,
        "sheet:0",
        "autoFilter",
        stored_filter(bare_filter("H1:I3")),
    );
    let failure = room.right.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert!(failure.to_string().contains("over a table"), "{failure}");
    assert_eq!(room.right.encode_state_as_update_v1(), state);
}

/// The exemption is the filter every replica opened, so an odd source stays
/// editable around it.
#[test]
fn a_source_filter_over_a_table_stays_as_every_replica_opened_it() {
    for delivery in DELIVERIES {
        let mut room = Room::of(package_filtered("", "H1:I5"), (1_001, 2_002));
        room.left
            .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "around the filter");
        room.assert_both(|peer| assert_eq!(filter_range(peer, 0).as_deref(), Some("H1:I5")));
    }
}

fn added_rule(sqref: &str) -> String {
    format!(
        r#"{{"sqref":"{sqref}","type":"expression","formulas":["A1>1"],"format":{{"bold":true}},"seq":0}}"#
    )
}

/// Everything no honest replica writes on a sheet, with what the refusal says.
fn hostile_entries() -> Vec<(&'static str, Any, &'static str)> {
    let good = added_rule("A1:B2");
    let source_filter = Workbook::open(&source()).unwrap().model().sheets[0]
        .auto_filter
        .clone()
        .unwrap();
    vec![
        (
            "cf:abc:1:0",
            Any::from(good.clone()),
            "invalid conditional format key",
        ),
        (
            "cf:01:2:3",
            Any::from(good.clone()),
            "invalid conditional format key",
        ),
        ("cf:1:2:3", Any::Bool(false), "is not a string"),
        (
            "cf:cf1",
            Any::from(good.clone()),
            "invalid conditional format removal",
        ),
        ("cf:cf99", Any::Bool(false), "has no source rule"),
        (
            "cf:cf1",
            Any::Bool(true),
            "invalid conditional format removal",
        ),
        (
            "cf:1:2:3",
            Any::from(good.replace(r#""seq":0"#, r#""seq":0,"priority":1"#)),
            "unknown field",
        ),
        (
            "cf:1:2:3",
            Any::from(good.replace(r#""seq":0"#, r#""seq":0,"raw_xml":"<x/>""#)),
            "unknown field",
        ),
        ("cf:1:2:3", Any::from(added_rule("$A$1:B2")), "noncanonical"),
        (
            "cf:1:2:3",
            Any::from(added_rule("A1:XFD1048577")),
            "conditional format range",
        ),
        (
            "cf:1:2:3",
            Any::from(good.replace(r#""bold":true"#, r##""fill":"#GG0000""##)),
            "colors",
        ),
        (
            "autoFilter",
            Any::from(stored_filter(AutoFilter {
                raw_xml: Some(r#"<autoFilter ref="A1:C9"><x/></autoFilter>"#.to_string()),
                ..source_filter.clone()
            })),
            FROZEN,
        ),
        (
            "autoFilter",
            Any::from(stored_filter(AutoFilter {
                range: range("B1:D9"),
                ..source_filter
            })),
            FROZEN,
        ),
    ]
}

#[test]
fn entries_a_hostile_peer_writes_are_refused_without_a_change() {
    for (key, stored, reason) in hostile_entries() {
        let mut peer = Workbook::open_collaborative(&source(), 6_006).unwrap();
        let model = peer.model().clone();
        let state = peer.encode_state_as_update_v1();
        let update = hostile_entry(&state, "sheet:0", key, stored);
        let failure = peer
            .apply_update_v1(&update, OPTIONS)
            .unwrap_err()
            .to_string();
        assert!(failure.contains(reason), "{key}: {failure}");
        assert_eq!(peer.model(), &model);
        assert_eq!(peer.encode_state_as_update_v1(), state);
    }
}

#[test]
fn a_rewritten_baseline_blob_is_refused() {
    let mut peer = Workbook::open_collaborative(&source(), 6_006).unwrap();
    let state = peer.encode_state_as_update_v1();
    let update = hostile_entry(&state, "sheet:0", "conditionalFormats", "[]");
    let failure = peer.apply_update_v1(&update, OPTIONS).unwrap_err();
    assert_eq!(failure.to_string(), FROZEN);
    assert_eq!(peer.encode_state_as_update_v1(), state);
}

/// A redo or a later sheet shows what these hold with no further check on that
/// replica, so they are refused while nothing shows them. A sheet with no
/// baseline has nothing to remove and no source filter to keep to.
#[test]
fn entries_nothing_shows_are_refused_like_shown_ones_and_the_redo_converges() {
    let dormant = [
        ("cf:cf1", Any::Bool(false), "has no source rule"),
        ("cf:1:2:3", Any::from(added_rule("$A$1:B2")), "noncanonical"),
        (
            "cf:01:2:3",
            Any::from(added_rule("A1:B2")),
            "invalid conditional format key",
        ),
        (
            "autoFilter",
            Any::from(stored_filter(AutoFilter {
                raw_xml: Some(r#"<autoFilter ref="A1:C9"/>"#.to_string()),
                ..bare_filter("A1:C9")
            })),
            FROZEN,
        ),
    ];
    for delivery in DELIVERIES {
        let mut room = Room::open();
        room.left
            .apply_ops(vec![add_sheet(2, "Notes")], OPTIONS)
            .unwrap();
        room.touch_right();
        room.settle(delivery, "added");
        let added = appended_key(&room.left);
        assert!(room.left.undo(OPTIONS).unwrap().applied);
        room.touch_right();
        room.settle(delivery, "undone");
        let state = room.left.encode_state_as_update_v1();
        for key in [added.as_str(), "replica:7:0:0"] {
            for (entry, stored, reason) in dormant.iter().cloned() {
                let update = hostile_sheet(&state, key, entry, stored);
                for peer in [&mut room.left, &mut room.right] {
                    let before = peer.encode_state_as_update_v1();
                    let failure = peer
                        .apply_update_v1(&update, OPTIONS)
                        .unwrap_err()
                        .to_string();
                    assert!(failure.contains(reason), "{key} {entry}: {failure}");
                    assert_eq!(peer.encode_state_as_update_v1(), before);
                }
            }
        }
        let update = hostile_sheet(
            &state,
            "replica:7:0:0",
            "conditionalFormats",
            r#"[{"sqref":"A1","rules":[]}]"#,
        );
        let failure = room.right.apply_update_v1(&update, OPTIONS).unwrap_err();
        assert_eq!(failure.to_string(), FROZEN);
        room.redo_left(delivery);
        room.assert_both(|peer| assert_eq!(peer.model().sheets.len(), 3));
    }
}

#[test]
fn a_standalone_peers_rule_and_filter_stay_refused() {
    for ops in [
        vec![rule("A1:A5", "A1>1")],
        vec![remove("cf1")],
        vec![filter(Some("A1:D9"))],
    ] {
        let bytes = source();
        let mut standalone = Workbook::open(&bytes).unwrap();
        standalone.apply_ops(ops, OPTIONS).unwrap();
        let mut peer = Workbook::open_collaborative(&bytes, 5_005).unwrap();
        let model = peer.model().clone();
        let update = standalone
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap();
        let error = peer.apply_update_v1(&update, OPTIONS).unwrap_err();
        assert_eq!(error.to_string(), FROZEN);
        assert_eq!(peer.model(), &model);
    }
}

/// A state from before rules replicated holds no `cf:` key, so it is an empty
/// overlay with nothing to migrate.
#[test]
fn a_room_from_before_rules_replicated_takes_rules_and_filters() {
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
        let mut room = Room {
            bytes: SAMPLE.to_vec(),
            left: restored(1_001),
            right: restored(2_002),
            left_client: 1_001,
            right_client: 2_002,
            mirrors: 3_000,
        };
        room.exchange(delivery);
        room.concurrently(
            vec![rule("A1:A5", "A1>1")],
            vec![rule("A1:A5", "A1>2"), filter(Some("A1:C9"))],
            delivery,
        );
        room.assert_both(|peer| {
            assert!(holds(peer, 0, "A1>1"));
            assert!(holds(peer, 0, "A1>2"));
            assert_eq!(filter_range(peer, 0).as_deref(), Some("A1:C9"));
        });
        room.undo_left(delivery);
        room.assert_both(|peer| assert!(!holds(peer, 0, "A1>1")));
    }
}

#[test]
fn a_recovery_checkpoint_keeps_rules_and_filters_and_keeps_exchanging() {
    let mut room = Room::open();
    room.left
        .apply_ops(
            vec![rule("F1", "F1>1"), remove("cf1"), filter(Some("A1:D9"))],
            OPTIONS,
        )
        .unwrap();
    room.touch_right();
    room.settle(Delivery::Merged, "formats");
    let checkpoint = room.right.recovery_checkpoint(1, None).unwrap();
    let (mut restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(restored.model(), room.right.model());
    restored
        .apply_ops(vec![rule("F1", "F1>2"), filter(Some("A1:E9"))], OPTIONS)
        .unwrap();
    let update = restored
        .encode_diff_v1(&room.left.encode_state_vector_v1())
        .unwrap();
    room.left.apply_update_v1(&update, OPTIONS).unwrap();
    assert_eq!(room.left.model(), restored.model());
    assert_eq!(
        filter_entry(&room.left, 0).unwrap().formula,
        "'Data'!$A$1:$E$9"
    );
}

/// A peer has to accept what an author stores, so a spelling only the author's
/// path would take is settled before it is written.
#[test]
fn anchored_ranges_are_stored_plain_and_a_formula_that_would_clean_twice_is_refused() {
    let mut room = Room::open();
    room.left
        .apply_ops(
            vec![rule("$F$1:$F$3", "F1>1"), filter(Some("$A$1:$D$9"))],
            OPTIONS,
        )
        .unwrap();
    room.touch_right();
    room.settle(Delivery::Merged, "anchored");
    room.assert_both(|peer| {
        assert_eq!(listed(peer, 0).last().unwrap().1, "F1:F3");
        assert_eq!(filter_range(peer, 0).as_deref(), Some("A1:D9"));
    });
    let state = room.left.encode_state_as_update_v1();
    let model = room.left.model().clone();
    assert!(
        room.left
            .apply_ops(vec![rule("F1", "==F1>1")], OPTIONS)
            .is_err()
    );
    assert_eq!(room.left.model(), &model);
    assert_eq!(room.left.encode_state_as_update_v1(), state);
}

/// The reader takes 65,536 rules back from a file. A replica stops there on its
/// own, a standalone workbook does not, and removing makes room again.
#[test]
fn a_sheet_at_the_readers_bound_takes_no_more_rules_in_a_room() {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Data");
    sheet.conditional_formats.push(ConditionalFormatting {
        sqref: "A1".into(),
        rules: (1..=65_536)
            .map(|index| ConditionalFormatRule {
                id: format!("cf{index}"),
                rule_type: "expression".into(),
                operator: None,
                formulas: vec!["1".into()],
                priority: index,
                dxf_id: None,
                format: DifferentialFormat::default(),
                raw_xml: None,
            })
            .collect(),
    });
    model.sheets.push(sheet);
    let bytes = ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap();
    let mut peer = Workbook::open_collaborative(&bytes, 7).unwrap();
    let state = peer.encode_state_as_update_v1();
    let failure = peer
        .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
        .unwrap_err();
    assert!(
        failure.to_string().contains("most conditional formats"),
        "{failure}"
    );
    assert_eq!(peer.encode_state_as_update_v1(), state);
    peer.apply_ops(vec![remove("cf1")], OPTIONS).unwrap();
    peer.apply_ops(vec![rule("F1", "F1>1")], OPTIONS).unwrap();
    let mut standalone = Workbook::open(&bytes).unwrap();
    standalone
        .apply_ops(vec![rule("F1", "F1>1")], OPTIONS)
        .unwrap();
}

#[test]
fn the_ops_that_write_whole_tables_stay_refused() {
    let mut room = Room::open();
    let state = room.left.encode_state_as_update_v1();
    let formats = room.left.model().sheets[0].conditional_formats.clone();
    for op in [
        Op::SetConditionalFormats {
            sheet: SheetId(0),
            formats,
        },
        Op::RestoreAutoFilter {
            sheet: SheetId(0),
            filter: None,
        },
    ] {
        assert!(matches!(
            room.left.apply_ops(vec![op], OPTIONS),
            Err(Error::CollaborativeStructureOperation)
        ));
    }
    assert_eq!(room.left.encode_state_as_update_v1(), state);
}

#[test]
fn a_replica_that_only_received_saves_the_authors_bytes() {
    each_room(|room, delivery| {
        room.left
            .apply_ops(vec![rule("F1", "F1>1"), filter(Some("A1:D9"))], OPTIONS)
            .unwrap();
        room.right
            .apply_ops(vec![remove("cf2"), filter_on(1, Some("A1:B4"))], OPTIONS)
            .unwrap();
        room.settle(delivery, "both");
        let reopened = reopened(&room.left);
        assert_eq!(
            reopened.sheets[1]
                .auto_filter
                .as_ref()
                .unwrap()
                .range
                .to_a1(),
            "A1:B4"
        );
        assert_eq!(
            reopened.sheets[0]
                .conditional_formats
                .iter()
                .flat_map(|g| &g.rules)
                .count(),
            5
        );
    });
}
