//! dynamic arrays spill and yield to what an author writes in their way;
//! legacy ctrl-shift-enter arrays keep the rectangle they were entered in.

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRange, CellRef, CellState, CellValue, EditRequest,
    ErrorValue, MutationResult, Op, ProposalEditInput, ProposalRequest, Sheet, SheetId, StylePatch,
    Workbook, WorkbookModel,
};
use serde_json::json;

/// every array's registration and result on the first sheet, for comparing
/// replicas.
fn arrays(workbook: &Workbook) -> Vec<String> {
    let sheet = &workbook.model().sheets[0];
    sheet
        .array_definitions()
        .map(|(at, definition, extent)| format!("{at:?} {definition:?} {extent:?}"))
        .chain(sheet.iter_cells().filter_map(|(at, _)| {
            sheet
                .result_anchor(at)
                .map(|anchor| format!("{at:?} <- {anchor:?}"))
        }))
        .collect()
}

fn same_replicas(replicas: &[&Workbook]) {
    for pair in replicas.windows(2) {
        assert_eq!(pair[0].model().sheets, pair[1].model().sheets);
        assert_eq!(arrays(pair[0]), arrays(pair[1]));
    }
}

const METADATA: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>"#;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.to_owned(),
    }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

/// a one-sheet package holding `rows`, with the cell metadata excel writes for
/// dynamic arrays when `dynamic`.
fn package(rows: &str, dynamic: bool) -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Sheet1"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{rows}</sheetData></worksheet>"#
    );
    for (name, bytes) in &mut parts {
        let edited = match name.as_str() {
            "xl/worksheets/sheet1.xml" => sheet.clone(),
            "[Content_Types].xml" if dynamic => String::from_utf8(bytes.clone()).unwrap().replace(
                "</Types>",
                r#"<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>"#,
            ),
            "xl/_rels/workbook.xml.rels" if dynamic => String::from_utf8(bytes.clone()).unwrap().replace(
                "</Relationships>",
                r#"<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>"#,
            ),
            _ => continue,
        };
        *bytes = edited.into_bytes();
    }
    if dynamic {
        parts.push(("xl/metadata.xml".to_owned(), METADATA.as_bytes().to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// `SORT(A1:A3)` over 3, 1, 2, spilled into C1:C3 by excel.
const SORTED: &str = concat!(
    r#"<row r="1"><c r="A1"><v>3</v></c><c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c></row>"#,
    r#"<row r="2"><c r="A2"><v>1</v></c><c r="C2"><v>2</v></c></row>"#,
    r#"<row r="3"><c r="A3"><v>2</v></c><c r="C3"><v>3</v></c></row>"#,
);

fn sorted() -> Workbook {
    Workbook::open_recalculated(&package(SORTED, true), options()).unwrap()
}

fn replica(client_id: u64) -> Workbook {
    Workbook::open_collaborative_recalculated(&package(SORTED, true), client_id, options()).unwrap()
}

/// hands `to` everything `from` holds that `to` has not seen, as one update.
fn sync(from: &Workbook, to: &mut Workbook) {
    let update = from.encode_diff_v1(&to.encode_state_vector_v1()).unwrap();
    to.apply_update_v1(&update, options()).unwrap();
}

fn changed(result: &MutationResult) -> Vec<String> {
    result
        .changed
        .iter()
        .map(|address| address.cell.to_a1())
        .collect()
}

fn edit(workbook: &mut Workbook, address: &str, input: &str) -> MutationResult {
    workbook
        .edit_cell(SheetId(0), cell(address), input, options())
        .unwrap()
}

fn cleared(workbook: &Workbook) -> bool {
    column(workbook, &["C1", "C2", "C3"]) == [CellValue::Empty, CellValue::Empty, CellValue::Empty]
}

fn value(workbook: &Workbook, address: &str) -> CellValue {
    workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell(address))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn column(workbook: &Workbook, cells: &[&str]) -> Vec<CellValue> {
    cells
        .iter()
        .map(|address| value(workbook, address))
        .collect()
}

fn spilled(workbook: &Workbook) -> bool {
    column(workbook, &["C1", "C2", "C3"]) == [number(1.0), number(2.0), number(3.0)]
}

fn obstructed_by(workbook: &Workbook, obstruction: CellValue) -> bool {
    column(workbook, &["C1", "C2", "C3"])
        == [error(ErrorValue::Spill), obstruction, CellValue::Empty]
}

fn sheet_xml(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, sheet) = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    String::from_utf8(sheet.clone()).unwrap()
}

fn batch(workbook: &Workbook, a1: &str, input: &str) -> EditRequest {
    serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [{
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": a1 } },
            "inputs": [[input]],
        }],
    }))
    .unwrap()
}

#[test]
fn a_value_typed_into_a_spill_obstructs_it_until_undone() {
    let mut workbook = sorted();
    assert!(spilled(&workbook));

    workbook
        .edit_cell(SheetId(0), cell("C2"), "manual", options())
        .unwrap();
    assert!(obstructed_by(&workbook, text("manual")));

    workbook
        .edit_cell(SheetId(0), cell("A1"), "0", options())
        .unwrap();
    assert!(obstructed_by(&workbook, text("manual")));

    workbook.undo(options()).unwrap();
    assert!(obstructed_by(&workbook, text("manual")));
    workbook.undo(options()).unwrap();
    assert!(spilled(&workbook));
    workbook.redo(options()).unwrap();
    assert!(obstructed_by(&workbook, text("manual")));
}

#[test]
fn every_edit_path_keeps_what_it_writes_into_a_spill() {
    let mut pasted = sorted();
    pasted
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: cell("C2"),
                input: "pasted".into(),
            }],
            options(),
        )
        .unwrap();

    let mut raw = sorted();
    raw.apply_ops(
        vec![Op::SetCell {
            sheet: SheetId(0),
            at: cell("C2"),
            cell: CellState {
                value: text("raw"),
                ..CellState::default()
            },
        }],
        options(),
    )
    .unwrap();

    let mut proposed = sorted();
    let proposal = proposed
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("C2"),
                    input: "proposed".into(),
                    number_format: None,
                }],
            },
            options(),
        )
        .unwrap();
    proposed
        .accept_proposal(&proposal.id, false, options())
        .unwrap();

    let mut batched = sorted();
    let request = batch(&batched, "C2", "batched");
    batched.apply_edits(&request).unwrap().unwrap();

    for (mut workbook, written) in [
        (pasted, "pasted"),
        (raw, "raw"),
        (proposed, "proposed"),
        (batched, "batched"),
    ] {
        assert!(obstructed_by(&workbook, text(written)), "{written}");
        workbook
            .edit_cell(SheetId(0), cell("A1"), "0", options())
            .unwrap();
        assert!(obstructed_by(&workbook, text(written)), "{written}");
        workbook.undo(options()).unwrap();
        workbook.undo(options()).unwrap();
        assert!(spilled(&workbook), "{written}");
    }
}

#[test]
fn an_obstructed_spill_saves_and_reopens_obstructed() {
    let mut workbook = sorted();
    workbook
        .edit_cell(SheetId(0), cell("C2"), "manual", options())
        .unwrap();
    let saved = workbook.save().unwrap();
    let xml = sheet_xml(&saved);
    assert!(
        xml.contains(r#"<c r="C1" t="e" cm="1"><f t="array" ref="C1">"#),
        "{xml}"
    );
    assert!(!xml.contains(r#"r="C3""#), "{xml}");

    let reopened = Workbook::open_recalculated(&saved, options()).unwrap();
    assert!(obstructed_by(&reopened, text("manual")));
}

#[test]
fn clearing_what_obstructs_a_spill_lets_it_spill_again() {
    let mut workbook = sorted();
    workbook
        .edit_cell(SheetId(0), cell("C2"), "manual", options())
        .unwrap();
    workbook
        .edit_cell(SheetId(0), cell("C2"), "", options())
        .unwrap();
    assert!(spilled(&workbook));
}

#[test]
fn restyling_or_clearing_a_spilled_cell_leaves_the_spill() {
    let mut workbook = sorted();
    workbook
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("C2"), cell("C3")),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            options(),
        )
        .unwrap();
    assert!(spilled(&workbook));
    workbook.undo(options()).unwrap();
    assert!(spilled(&workbook));
    workbook
        .edit_cell(SheetId(0), cell("C3"), "", options())
        .unwrap();
    assert!(spilled(&workbook));
}

#[test]
fn clearing_a_spill_anchor_clears_its_result_until_undone() {
    let mut workbook = sorted();
    workbook
        .edit_cell(SheetId(0), cell("C1"), "", options())
        .unwrap();
    assert_eq!(
        column(&workbook, &["C1", "C2", "C3"]),
        [CellValue::Empty, CellValue::Empty, CellValue::Empty]
    );
    let xml = sheet_xml(&workbook.save().unwrap());
    assert!(
        !xml.contains(r#"r="C2""#) && !xml.contains(r#"r="C3""#),
        "{xml}"
    );
    workbook.undo(options()).unwrap();
    assert!(spilled(&workbook));
}

#[test]
fn a_grown_spill_saves_the_rectangle_it_now_fills() {
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>2</v></c><c r="C1" cm="1"><f t="array" ref="C1:C2">_xlfn.SEQUENCE(A1)</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="C2"><v>2</v></c></row>"#,
    );
    let bytes = package(rows, true);
    assert_eq!(
        Workbook::open(&bytes).unwrap().save().unwrap(),
        bytes,
        "an unedited workbook saves byte for byte"
    );
    let mut workbook = Workbook::open_recalculated(&bytes, options()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "3", options())
        .unwrap();
    let saved = workbook.save().unwrap();
    let xml = sheet_xml(&saved);
    assert!(
        xml.contains(r#"<c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn.SEQUENCE(A1)</f>"#),
        "{xml}"
    );

    let reopened = Workbook::open_recalculated(&saved, options()).unwrap();
    assert_eq!(
        column(&reopened, &["C1", "C2", "C3"]),
        [number(1.0), number(2.0), number(3.0)]
    );
}

#[test]
fn a_recalculated_dynamic_anchor_keeps_its_metadata() {
    let mut workbook = sorted();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "0", options())
        .unwrap();
    let xml = sheet_xml(&workbook.save().unwrap());
    assert!(
        xml.contains(r#"<c r="C1" cm="1"><f t="array" ref="C1:C3">"#),
        "{xml}"
    );
}

/// `A1:A3*2` entered with ctrl-shift-enter over C1 beside an occupied C2, over
/// E1, G1:G2 and I1:I4.
#[test]
fn a_legacy_array_keeps_the_rectangle_it_was_entered_in() {
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f t="array" ref="C1">A1:A3*2</f><v>2</v></c><c r="E1"><f t="array" ref="E1">A1:A3*2</f><v>2</v></c><c r="G1"><f t="array" ref="G1:G2">A1:A3*2</f><v>2</v></c><c r="I1"><f t="array" ref="I1:I4">A1:A3*2</f><v>2</v></c></row>"#,
        r#"<row r="2"><c r="A2"><v>2</v></c><c r="C2" t="inlineStr"><is><t>keep</t></is></c><c r="G2"><v>4</v></c><c r="I2"><v>4</v></c></row>"#,
        r#"<row r="3"><c r="A3"><v>3</v></c><c r="I3"><v>6</v></c></row>"#,
        r#"<row r="4"><c r="I4" t="e"><v>#N/A</v></c></row>"#,
    );
    let mut workbook = Workbook::open_recalculated(&package(rows, false), options()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "5", options())
        .unwrap();
    assert_eq!(
        column(&workbook, &["C1", "C2"]),
        [number(10.0), text("keep")]
    );
    assert_eq!(
        column(&workbook, &["E1", "E2", "E3"]),
        [number(10.0), CellValue::Empty, CellValue::Empty]
    );
    assert_eq!(
        column(&workbook, &["G1", "G2", "G3"]),
        [number(10.0), number(4.0), CellValue::Empty]
    );
    assert_eq!(
        column(&workbook, &["I1", "I2", "I3", "I4"]),
        [
            number(10.0),
            number(4.0),
            number(6.0),
            error(ErrorValue::NA)
        ]
    );
    let xml = sheet_xml(&workbook.save().unwrap());
    for anchor in [
        r#"<f t="array" ref="C1">"#,
        r#"<f t="array" ref="E1">"#,
        r#"<f t="array" ref="G1:G2">"#,
        r#"<f t="array" ref="I1:I4">"#,
    ] {
        assert!(xml.contains(anchor), "{anchor} in {xml}");
    }
    assert!(!xml.contains(" cm="), "{xml}");
}

/// a model built with a dynamic array writes the metadata that marks it.
#[test]
fn a_dynamic_array_from_a_model_saves_as_one() {
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Empty,
            formula: Some("_xlfn.SEQUENCE(3)".into()),
            style: None,
        },
    );
    sheet.set_dynamic_array_formula(cell("A1"), CellRange::new(cell("A1"), cell("A1")));
    let model = WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    };
    let parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let reopened =
        Workbook::open_recalculated(&ooxml_opc::rezip_parts(&parts).unwrap(), options()).unwrap();
    assert_eq!(
        column(&reopened, &["A1", "A2", "A3"]),
        [number(1.0), number(2.0), number(3.0)]
    );
}

/// what settling a write changes beside the cell written is reported with what
/// recalculation moved.
#[test]
fn a_mutation_reports_what_settling_its_write_changed() {
    let mut workbook = sorted();
    assert_eq!(changed(&edit(&mut workbook, "C2", "manual")), ["C1", "C3"]);
    let undone = workbook.undo(options()).unwrap();
    assert_eq!(changed(&undone), ["C1", "C2", "C3"]);
    assert_eq!(changed(&edit(&mut workbook, "C1", "")), ["C2", "C3"]);

    let mut shared = replica(21);
    assert_eq!(changed(&edit(&mut shared, "C2", "manual")), ["C1", "C3"]);

    let mut batched = sorted();
    let request = batch(&batched, "C2", "batched");
    let application = batched.apply_edits(&request).unwrap().unwrap();
    let reported: Vec<_> = application
        .calculation
        .changed
        .iter()
        .map(|target| target.a1.as_str())
        .collect();
    assert_eq!(reported, ["C1", "C3"]);
}

/// clearing an anchor retires its array: a formula typed there again is a new
/// entry, not the old array brought back, so what was typed into the old
/// rectangle meanwhile stays; undo restores the array itself.
#[test]
fn a_formula_retyped_at_a_cleared_anchor_is_a_new_entry() {
    for mut workbook in [sorted(), replica(22)] {
        edit(&mut workbook, "C1", "");
        assert!(cleared(&workbook));
        edit(&mut workbook, "C2", "manual");
        edit(&mut workbook, "C1", "=_xlfn._xlws.SORT(A1:A3)");
        let sheet = workbook.sheet(SheetId(0)).unwrap();
        assert_eq!(sheet.array_definition(cell("C1")), None);
        assert_eq!(value(&workbook, "C2"), text("manual"));
        assert_eq!(value(&workbook, "C3"), CellValue::Empty);
        for _ in 0..3 {
            workbook.undo(options()).unwrap();
        }
        assert!(spilled(&workbook));
        assert!(workbook.model().sheets[0].is_dynamic_array(cell("C1")));
    }
}

/// what an author writes into a spill is the author's even when it equals
/// what the spill shows there, on every edit path.
#[test]
fn writing_the_displayed_value_into_a_spill_obstructs_it() {
    let mut typed = sorted();
    edit(&mut typed, "C2", "2");

    let mut pasted = sorted();
    pasted
        .edit_cells(
            SheetId(0),
            &[
                CellInput {
                    cell: cell("C2"),
                    input: "2".into(),
                },
                CellInput {
                    cell: cell("C3"),
                    input: "3".into(),
                },
            ],
            options(),
        )
        .unwrap();

    let mut proposed = sorted();
    let proposal = proposed
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("C2"),
                    input: "2".into(),
                    number_format: None,
                }],
            },
            options(),
        )
        .unwrap();
    proposed
        .accept_proposal(&proposal.id, false, options())
        .unwrap();

    let mut batched = sorted();
    let request = batch(&batched, "C2", "2");
    batched.apply_edits(&request).unwrap().unwrap();

    let mut shared = replica(81);
    let mut peer = replica(82);
    edit(&mut shared, "C2", "2");
    sync(&shared, &mut peer);

    for workbook in [&typed, &proposed, &batched, &shared, &peer] {
        assert!(obstructed_by(workbook, number(2.0)));
    }
    assert_eq!(
        column(&pasted, &["C1", "C2", "C3"]),
        [error(ErrorValue::Spill), number(2.0), number(3.0)]
    );
    let reopened = Workbook::open_recalculated(&typed.save().unwrap(), options()).unwrap();
    assert!(obstructed_by(&reopened, number(2.0)));
}

/// a value typed where a cleared spill's result stood survives on every
/// replica, whatever the spill showed there when the workbook opened.
#[test]
fn a_value_typed_where_a_cleared_spill_stood_survives() {
    let mut writer = replica(91);
    let mut reader = replica(92);
    edit(&mut writer, "C1", "");
    edit(&mut writer, "C2", "2");
    sync(&writer, &mut reader);
    for replica in [&writer, &reader] {
        assert_eq!(
            column(replica, &["C1", "C2", "C3"]),
            [CellValue::Empty, number(2.0), CellValue::Empty]
        );
        let reopened = Workbook::open(&replica.save().unwrap()).unwrap();
        assert_eq!(value(&reopened, "C2"), number(2.0));
    }
}

/// every value an author pastes over a spill is kept, whether it reaches a
/// replica as one update or as several.
#[test]
fn values_pasted_over_a_spill_are_all_kept_on_every_replica() {
    let mut writer = replica(31);
    let mut follower = replica(32);
    let mut late = replica(33);
    writer
        .edit_cells(
            SheetId(0),
            &[
                CellInput {
                    cell: cell("C2"),
                    input: "a".into(),
                },
                CellInput {
                    cell: cell("C3"),
                    input: "b".into(),
                },
            ],
            options(),
        )
        .unwrap();
    sync(&writer, &mut follower);
    for replica in [&writer, &follower] {
        assert_eq!(
            column(replica, &["C1", "C2", "C3"]),
            [error(ErrorValue::Spill), text("a"), text("b")]
        );
    }
    edit(&mut writer, "C2", "c");
    sync(&writer, &mut follower);
    edit(&mut writer, "C3", "");
    sync(&writer, &mut follower);
    sync(&writer, &mut late);
    for replica in [&writer, &follower, &late] {
        assert!(obstructed_by(replica, text("c")));
    }
}

/// two separate edits by one replica settle alike on a replica that receives
/// them as one update.
#[test]
fn separate_edits_and_their_combined_update_settle_alike() {
    let mut writer = replica(41);
    let mut follower = replica(42);
    let mut late = replica(43);
    edit(&mut writer, "C2", "a");
    sync(&writer, &mut follower);
    edit(&mut writer, "C3", "b");
    sync(&writer, &mut follower);
    sync(&writer, &mut late);
    for replica in [&writer, &follower, &late] {
        assert_eq!(
            column(replica, &["C1", "C2", "C3"]),
            [error(ErrorValue::Spill), text("a"), text("b")]
        );
    }
}

/// the shared document holds what the spill filled at open; undoing a write
/// over it restores that as the spill's own value, not as an obstruction.
#[test]
fn collaborative_undo_and_redo_of_a_write_into_a_spill() {
    let mut writer = replica(51);
    let mut follower = replica(52);
    edit(&mut writer, "C2", "manual");
    sync(&writer, &mut follower);
    for replica in [&writer, &follower] {
        assert!(obstructed_by(replica, text("manual")));
    }
    writer.undo(options()).unwrap();
    sync(&writer, &mut follower);
    let mut late = replica(53);
    sync(&writer, &mut late);
    for replica in [&writer, &follower, &late] {
        assert!(spilled(replica));
    }
    writer.redo(options()).unwrap();
    sync(&writer, &mut follower);
    sync(&writer, &mut late);
    for replica in [&writer, &follower, &late] {
        assert!(obstructed_by(replica, text("manual")));
    }
}

/// clearing a shared anchor clears its spill on every replica and in the
/// saved file, and undoing it restores the dynamic array.
#[test]
fn clearing_a_shared_anchor_clears_its_spill_until_undone() {
    let mut writer = replica(61);
    let mut follower = replica(62);
    edit(&mut writer, "C1", "");
    sync(&writer, &mut follower);
    for replica in [&writer, &follower] {
        assert!(cleared(replica));
        let xml = sheet_xml(&replica.save().unwrap());
        assert!(
            !xml.contains(r#"r="C2""#) && !xml.contains(r#"r="C3""#),
            "{xml}"
        );
    }
    writer.undo(options()).unwrap();
    sync(&writer, &mut follower);
    for replica in [&writer, &follower] {
        assert!(spilled(replica));
        let saved = replica.save().unwrap();
        let xml = sheet_xml(&saved);
        assert!(
            xml.contains(r#"<c r="C1" cm="1"><f t="array" ref="C1:C3">"#),
            "{xml}"
        );
        assert!(spilled(
            &Workbook::open_recalculated(&saved, options()).unwrap()
        ));
    }
}

/// the shared document carries the written cell, and each replica replays what
/// it does to the spill, whatever is recalculated afterwards.
#[test]
fn a_collaborator_writing_into_a_spill_obstructs_it_on_every_replica() {
    let mut writer = replica(11);
    let mut reader = replica(12);
    edit(&mut writer, "C2", "manual");
    edit(&mut writer, "A1", "0");
    sync(&writer, &mut reader);
    for replica in [&writer, &reader] {
        assert!(obstructed_by(replica, text("manual")));
    }
}

/// a legacy array whose anchor is cleared comes back over the rectangle it
/// was entered in.
#[test]
fn undoing_a_cleared_legacy_anchor_restores_its_rectangle() {
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f t="array" ref="C1:C3">A1:A2*2</f><v>2</v></c></row>"#,
        r#"<row r="2"><c r="A2"><v>2</v></c><c r="C2"><v>4</v></c></row>"#,
        r#"<row r="3"><c r="C3" t="e"><v>#N/A</v></c></row>"#,
    );
    let bytes = package(rows, false);
    for mut workbook in [
        Workbook::open_recalculated(&bytes, options()).unwrap(),
        Workbook::open_collaborative_recalculated(&bytes, 71, options()).unwrap(),
    ] {
        edit(&mut workbook, "C1", "");
        assert!(cleared(&workbook));
        workbook.undo(options()).unwrap();
        assert_eq!(
            column(&workbook, &["C1", "C2", "C3"]),
            [number(2.0), number(4.0), error(ErrorValue::NA)]
        );
        let xml = sheet_xml(&workbook.save().unwrap());
        assert!(
            xml.contains(r#"<c r="C1"><f t="array" ref="C1:C3">"#),
            "{xml}"
        );
    }
}

/// a spill that shrinks saves the smaller rectangle and nothing beyond it,
/// and reopens spilling the same way.
#[test]
fn a_shrunk_spill_saves_and_reopens_at_its_new_size() {
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>3</v></c><c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn.SEQUENCE(A1)</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="C2"><v>2</v></c></row>"#,
        r#"<row r="3"><c r="C3"><v>3</v></c></row>"#,
    );
    let mut workbook = Workbook::open_recalculated(&package(rows, true), options()).unwrap();
    edit(&mut workbook, "A1", "2");
    let saved = workbook.save().unwrap();
    let xml = sheet_xml(&saved);
    assert!(
        xml.contains(r#"<c r="C1" cm="1"><f t="array" ref="C1:C2">_xlfn.SEQUENCE(A1)</f>"#),
        "{xml}"
    );
    assert!(!xml.contains(r#"r="C3""#), "{xml}");
    let mut reopened = Workbook::open_recalculated(&saved, options()).unwrap();
    assert_eq!(
        column(&reopened, &["C1", "C2", "C3"]),
        [number(1.0), number(2.0), CellValue::Empty]
    );
    edit(&mut reopened, "A1", "3");
    assert_eq!(
        column(&reopened, &["C1", "C2", "C3"]),
        [number(1.0), number(2.0), number(3.0)]
    );
}

/// concurrent edits settle alike whichever order each replica receives them
/// in, and alike on a replica that receives them together.
#[test]
fn concurrent_edits_settle_alike_in_either_delivery_order() {
    for (first, second) in [("C2", "C3"), ("C1", "C2"), ("C2", "C1")] {
        let mut left = replica(101);
        let mut right = replica(102);
        let mut late = replica(103);
        edit(&mut left, first, if first == "C1" { "" } else { "a" });
        edit(&mut right, second, if second == "C1" { "" } else { "b" });
        let from_left = left.encode_state_as_update_v1();
        let from_right = right.encode_state_as_update_v1();
        left.apply_update_v1(&from_right, options()).unwrap();
        right.apply_update_v1(&from_left, options()).unwrap();
        late.apply_update_v1(&from_right, options()).unwrap();
        late.apply_update_v1(&from_left, options()).unwrap();
        same_replicas(&[&left, &right, &late]);
    }
}

/// what settling a write changes is recalculated for everything reading it.
#[test]
fn readers_of_a_settled_spill_recalculate() {
    let rows = format!(
        "{}{}",
        SORTED.replace(
            r#"<c r="C1" cm="1">"#,
            r#"<c r="D1"><f>SUM(C2:C3)</f><v>5</v></c><c r="C1" cm="1">"#
        ),
        ""
    );
    let rows = rows.replace(
        r#"<c r="D1"><f>SUM(C2:C3)</f><v>5</v></c><c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c>"#,
        r#"<c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c><c r="D1"><f>SUM(C2:C3)</f><v>5</v></c>"#,
    );
    let bytes = package(&rows, true);
    let open = || Workbook::open_recalculated(&bytes, options()).unwrap();
    let shared =
        |client| Workbook::open_collaborative_recalculated(&bytes, client, options()).unwrap();
    for mut workbook in [open(), shared(111)] {
        assert_eq!(value(&workbook, "D1"), number(5.0));
        edit(&mut workbook, "C1", "");
        assert_eq!(value(&workbook, "D1"), number(0.0));
        workbook.undo(options()).unwrap();
        assert_eq!(value(&workbook, "D1"), number(5.0));
        edit(&mut workbook, "C3", "10");
        assert_eq!(value(&workbook, "D1"), number(10.0));
    }
    let mut batched = open();
    let request = batch(&batched, "C3", "10");
    batched.apply_edits(&request).unwrap().unwrap();
    assert_eq!(value(&batched, "D1"), number(10.0));
}

/// each anchor keeps the `cm` index its source gave it through a rewrite.
#[test]
fn anchors_keep_their_own_cell_metadata_through_a_save() {
    let metadata = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="2"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="1"/></ext></extLst></bk></futureMetadata><cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk></cellMetadata></metadata>"#;
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>2</v></c><c r="C1" cm="1"><f t="array" ref="C1:C2">_xlfn.SEQUENCE(A1)</f><v>1</v></c><c r="E1" cm="2"><f t="array" ref="E1:E2">_xlfn.SEQUENCE(A1,1,10)</f><v>10</v></c></row>"#,
        r#"<row r="2"><c r="C2"><v>2</v></c><c r="E2"><v>11</v></c></row>"#,
    );
    let mut parts = ooxml_opc::unzip_parts(&package(rows, true)).unwrap();
    for (name, bytes) in &mut parts {
        if name == "xl/metadata.xml" {
            *bytes = metadata.as_bytes().to_vec();
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open_recalculated(&bytes, options()).unwrap();
    edit(&mut workbook, "A1", "3");
    let saved = workbook.save().unwrap();
    let xml = sheet_xml(&saved);
    assert!(
        xml.contains(r#"<c r="C1" cm="1"><f t="array" ref="C1:C3">"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<c r="E1" cm="2"><f t="array" ref="E1:E3">"#),
        "{xml}"
    );
    let reopened = Workbook::open_recalculated(&saved, options()).unwrap();
    for anchor in ["C1", "E1"] {
        assert!(
            reopened.model().sheets[0].is_dynamic_array(cell(anchor)),
            "{anchor}"
        );
    }
    assert_eq!(
        column(&reopened, &["E1", "E2", "E3"]),
        [number(10.0), number(11.0), number(12.0)]
    );
}

/// followers a producer marked with an empty `<f/>` are the spill's result,
/// and leave with it for good.
#[test]
fn marked_followers_do_not_outlive_their_anchor() {
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>3</v></c><c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="A2"><v>1</v></c><c r="C2"><f/><v>2</v></c></row>"#,
        r#"<row r="3"><c r="A3"><v>2</v></c><c r="C3"><f/><v>3</v></c></row>"#,
    );
    let bytes = package(rows, true);
    let mut writer = Workbook::open_collaborative_recalculated(&bytes, 121, options()).unwrap();
    let mut reader = Workbook::open_collaborative_recalculated(&bytes, 122, options()).unwrap();
    edit(&mut writer, "C1", "");
    sync(&writer, &mut reader);
    edit(&mut writer, "G1", "unrelated");
    sync(&writer, &mut reader);
    for replica in [&writer, &reader] {
        assert!(cleared(replica));
        let xml = sheet_xml(&replica.save().unwrap());
        assert!(
            !xml.contains(r#"r="C2""#) && !xml.contains(r#"r="C3""#),
            "{xml}"
        );
    }
}

/// a cell holding a formula of its own inside a recorded spill is the
/// author's: it obstructs the spill and outlives its anchor.
#[test]
fn authored_cells_inside_a_recorded_rectangle_are_kept() {
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>3</v></c><c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="A2"><v>1</v></c><c r="C2"><v>2</v></c></row>"#,
        r#"<row r="3"><c r="A3"><v>2</v></c><c r="C3"><f>A1*10</f><v>30</v></c></row>"#,
    );
    let mut workbook = Workbook::open_recalculated(&package(rows, true), options()).unwrap();
    assert_eq!(value(&workbook, "C1"), error(ErrorValue::Spill));
    edit(&mut workbook, "C1", "");
    assert_eq!(value(&workbook, "C3"), number(30.0));
    assert_eq!(
        workbook.model().sheets[0]
            .cell(cell("C3"))
            .unwrap()
            .formula
            .as_deref(),
        Some("A1*10")
    );
}

/// editing an anchor's formula keeps its kind: a dynamic array still spills,
/// a legacy one keeps the rectangle it was entered in.
#[test]
fn an_edited_anchor_formula_keeps_its_array_kind() {
    let mut dynamic = sorted();
    edit(&mut dynamic, "C1", "=_xlfn._xlws.SORT(A1:A3,1,-1)");
    assert_eq!(
        column(&dynamic, &["C1", "C2", "C3"]),
        [number(3.0), number(2.0), number(1.0)]
    );
    let rows = concat!(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f t="array" ref="C1:C3">A1:A2*2</f><v>2</v></c></row>"#,
        r#"<row r="2"><c r="A2"><v>2</v></c><c r="C2"><v>4</v></c></row>"#,
        r#"<row r="3"><c r="C3" t="e"><v>#N/A</v></c></row>"#,
    );
    let mut legacy = Workbook::open_recalculated(&package(rows, false), options()).unwrap();
    edit(&mut legacy, "C1", "=A1:A2*3");
    assert_eq!(
        column(&legacy, &["C1", "C2", "C3"]),
        [number(3.0), number(6.0), error(ErrorValue::NA)]
    );
    assert!(
        legacy
            .edit_cell(SheetId(0), cell("C2"), "manual", options())
            .is_err(),
        "a legacy array is edited whole"
    );
}

/// an edit elsewhere neither settles nor recalculates an unrelated spill,
/// however large: its cached result, stale on purpose, stays untouched.
#[test]
fn an_unrelated_edit_leaves_a_large_spill_untouched() {
    const ROWS: u32 = 20_000;
    let mut rows = format!(
        r#"<row r="1"><c r="A1"><v>{ROWS}</v></c><c r="C1" cm="1"><f t="array" ref="C1:C{ROWS}">_xlfn.SEQUENCE(A1)</f><v>7</v></c><c r="E1"><v>1</v></c><c r="F1"><f>E1+1</f><v>2</v></c></row>"#
    );
    for row in 2..=ROWS {
        rows.push_str(&format!(
            r#"<row r="{row}"><c r="C{row}"><v>7</v></c></row>"#
        ));
    }
    let bytes = package(&rows, true);
    for mut workbook in [
        Workbook::open(&bytes).unwrap(),
        Workbook::open_collaborative(&bytes, 131).unwrap(),
    ] {
        let result = edit(&mut workbook, "E1", "5");
        assert_eq!(changed(&result), ["F1"]);
        assert_eq!(value(&workbook, "C2"), number(7.0));
        assert_eq!(value(&workbook, &format!("C{ROWS}")), number(7.0));
    }
}
