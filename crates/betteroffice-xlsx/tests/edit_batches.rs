use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellTarget, CellValue, DocumentVersion,
    EditApplication, EditFailureCode, EditOutcome, EditRefusal, EditRequest, Error, FindRequest,
    MAX_REQUEST_BYTES, Op, ProposalEditInput, ProposalRequest, ReadRequest, Sheet, SheetId,
    UpdateOrigin, UpdateSubscription, Workbook, WorkbookModel,
};
use serde_json::{Value, json};
use xlsx_model::Xf;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn put(
    sheet: &mut Sheet,
    address: &str,
    value: CellValue,
    formula: Option<&str>,
    style: Option<u32>,
) {
    sheet.set_cell(
        cell(address),
        Cell {
            value,
            formula: formula.map(str::to_owned),
            style,
        },
    );
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.to_owned(),
    }
}

/// Data holds formulas, a Text-formatted E1, a two-decimal D1, a merge F1:G1 and an array
/// formula over E3:E4; Totals reads Data; Locked is protected. Two parts nothing models ride
/// along.
fn fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Xf::default());
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(49),
        ..Xf::default()
    });
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(2),
        ..Xf::default()
    });
    let mut data = Sheet::new("Data");
    put(&mut data, "A1", number(10.0), None, None);
    put(&mut data, "A2", number(5.0), None, None);
    put(&mut data, "B1", number(15.0), Some("SUM(A1:A2)"), None);
    put(&mut data, "C1", text("hello"), None, None);
    put(&mut data, "C2", text("shell"), None, None);
    put(&mut data, "D1", number(1.5), None, Some(2));
    put(&mut data, "E1", CellValue::Empty, None, Some(1));
    put(&mut data, "F1", text("merged"), None, None);
    put(&mut data, "E3", number(20.0), Some("A1:A2*2"), None);
    put(&mut data, "E4", number(10.0), None, None);
    data.merges.push(CellRange::parse_a1("F1:G1").unwrap());
    data.set_array_formula(cell("E3"), CellRange::parse_a1("E3:E4").unwrap());
    let mut totals = Sheet::new("Totals");
    put(&mut totals, "A1", number(30.0), Some("Data!B1*2"), None);
    let mut locked = Sheet::new("Locked");
    put(&mut locked, "A1", number(1.0), None, None);
    model.sheets.extend([data, totals, locked]);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let protected = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet3.xml")
        .unwrap();
    let xml = String::from_utf8(protected.1.clone()).unwrap().replace(
        "</sheetData>",
        r#"</sheetData><sheetProtection sheet="1"/>"#,
    );
    assert!(xml.contains("sheetProtection"));
    protected.1 = xml.into_bytes();
    let relationships = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/_rels/workbook.xml.rels")
        .unwrap();
    relationships.1 = String::from_utf8(relationships.1.clone())
        .unwrap()
        .replace(
            "</Relationships>",
            r#"<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>"#,
        )
        .into_bytes();
    parts.extend([
        (
            "customXml/item1.xml".to_owned(),
            br#"<custom fidelity="byte-identical">payload</custom>"#.to_vec(),
        ),
        (
            "xl/opaque/extension.bin".to_owned(),
            vec![0, 1, 2, 3, 254, 255],
        ),
    ]);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn standalone() -> Workbook {
    Workbook::open(&fixture()).unwrap()
}

fn collaborative(client_id: u64) -> Workbook {
    Workbook::open_collaborative(&fixture(), client_id).unwrap()
}

fn request(version: &DocumentVersion, extra: Value) -> EditRequest {
    let mut value = json!({ "expectVersion": version });
    value
        .as_object_mut()
        .unwrap()
        .extend(extra.as_object().unwrap().clone());
    serde_json::from_value(value).unwrap()
}

fn steps(workbook: &Workbook, steps: Value) -> EditRequest {
    request(&workbook.version(), json!({ "steps": steps }))
}

fn apply(workbook: &mut Workbook, batch: Value) -> EditOutcome {
    let request = steps(workbook, batch);
    workbook.apply_edits(&request).unwrap()
}

fn applied(workbook: &mut Workbook, batch: Value) -> EditApplication {
    apply(workbook, batch).expect("the batch applies")
}

fn refused(workbook: &mut Workbook, batch: Value) -> EditRefusal {
    apply(workbook, batch).expect_err("the batch refuses")
}

fn target(sheet: &str, a1: &str) -> Value {
    json!({ "sheetId": sheet, "range": { "kind": "a1", "a1": a1 } })
}

fn inputs(sheet: &str, a1: &str, values: Value) -> Value {
    json!({ "op": "setCellInputs", "target": target(sheet, a1), "inputs": values })
}

fn input(workbook: &Workbook, sheet: u32, address: &str) -> String {
    workbook.cell(SheetId(sheet), cell(address)).unwrap().input
}

fn value(workbook: &Workbook, sheet: u32, address: &str) -> CellValue {
    workbook
        .sheet(SheetId(sheet))
        .unwrap()
        .cell(cell(address))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn display(workbook: &Workbook, sheet: &str, a1: &str) -> String {
    let read = workbook
        .read_cells(&serde_json::from_value(json!({ "ranges": [target(sheet, a1)] })).unwrap())
        .unwrap()
        .unwrap();
    read.ranges[0].cells[0][0].display_text.clone()
}

fn recorded_updates(workbook: &Workbook) -> (Arc<Mutex<Vec<UpdateOrigin>>>, UpdateSubscription) {
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let subscription = workbook
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.origin))
        .unwrap();
    (events, subscription)
}

#[test]
fn versions_move_only_with_committed_changes() {
    let mut workbook = standalone();
    let opened = workbook.version();
    workbook
        .read_cells(&ReadRequest::default())
        .unwrap()
        .unwrap();
    workbook.set_active_sheet(SheetId(1)).unwrap();
    let id = workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("H1"),
                    input: "1".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap()
        .id;
    workbook.reject_proposal(&id);
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(workbook.version(), opened);

    workbook
        .edit_cell(SheetId(0), cell("A1"), "11", CalculationOptions::default())
        .unwrap();
    let edited = workbook.version();
    assert_ne!(edited, opened);
    workbook.undo(CalculationOptions::default()).unwrap();
    let undone = workbook.version();
    assert_ne!(undone, edited);
    assert_ne!(undone, opened, "versions never repeat within a session");
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_ne!(workbook.version(), undone);

    assert_ne!(standalone().version(), standalone().version());
}

#[test]
fn value_changing_recalculation_publishes_without_a_document_update() {
    let mut sheet = Sheet::new("Clock");
    put(&mut sheet, "A1", CellValue::Empty, Some("NOW()"), None);
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    let mut workbook = Workbook::from_model(model).unwrap();
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = workbook
        .observe_update_v1(move |event| {
            observed
                .lock()
                .unwrap()
                .push((event.origin, event.update.len()))
        })
        .unwrap();
    let state = workbook.encode_state_as_update_v1();
    let before = workbook.version();
    let now = CalculationOptions {
        now_serial: Some(45_000.5),
    };
    assert!(!workbook.recalculate_all(now).changed.is_empty());
    assert_eq!(value(&workbook, 0, "A1"), number(45_000.5));
    let recalculated = workbook.version();
    assert_ne!(recalculated, before);
    assert_eq!(*events.lock().unwrap(), [(UpdateOrigin::Recalculation, 0)]);
    assert_eq!(workbook.encode_state_as_update_v1(), state);

    assert!(workbook.recalculate_all(now).changed.is_empty());
    assert_eq!(workbook.version(), recalculated);
    assert_eq!(events.lock().unwrap().len(), 1);
}

fn two_sheet_batch() -> Value {
    json!([
        inputs("sheet:0", "A1:A2", json!([["20"], ["30"]])),
        { "op": "setFormulas", "target": target("sheet:0", "C3"), "formulas": [["A1+A2"]] },
        { "op": "patchStyle", "target": target("sheet:1", "B2"), "patch": { "bold": true } },
    ])
}

#[test]
fn a_batch_commits_every_step_as_one_undo_step() {
    for mut workbook in [standalone(), collaborative(7)] {
        let depth = workbook.history_state().undo_depth;
        let base = workbook.version();
        let result = applied(&mut workbook, two_sheet_batch());
        assert!(result.applied);
        assert_eq!(result.base_version, base);
        assert_eq!(result.version, workbook.version());
        assert_ne!(result.version, base);
        assert_eq!(result.changed_sheets, ["sheet:0", "sheet:1"]);
        assert_eq!(result.receipts.len(), 3);
        assert_eq!(
            result.receipts[0]
                .changed_cells
                .iter()
                .map(|cell| cell.a1.as_str())
                .collect::<Vec<_>>(),
            ["A1", "A2"]
        );
        assert_eq!(result.receipts[1].changed_cells[0].a1, "C3");
        assert_eq!(result.receipts[2].target.sheet_id, "sheet:1");
        let dependents = result
            .calculation
            .changed
            .iter()
            .map(|cell| (cell.sheet_id.as_str(), cell.a1.as_str()))
            .collect::<Vec<_>>();
        assert!(dependents.contains(&("sheet:0", "B1")), "{dependents:?}");
        assert!(dependents.contains(&("sheet:1", "A1")), "{dependents:?}");
        assert_eq!(value(&workbook, 0, "B1"), number(50.0));
        assert_eq!(value(&workbook, 0, "C3"), number(50.0));
        assert_eq!(value(&workbook, 1, "A1"), number(100.0));
        assert_eq!(workbook.history_state().undo_depth, depth + 1);

        workbook.undo(CalculationOptions::default()).unwrap();
        assert_eq!(input(&workbook, 0, "A1"), "10");
        assert_eq!(input(&workbook, 0, "C3"), "");
        assert_eq!(value(&workbook, 1, "A1"), number(30.0));
        assert_eq!(
            workbook
                .selection_formatting(SheetId(1), CellRange::parse_a1("B2").unwrap())
                .unwrap()
                .bold,
            Some(false)
        );
    }
}

#[test]
fn a_late_refusal_leaves_everything_untouched() {
    for mut workbook in [standalone(), collaborative(8)] {
        workbook
            .edit_cell(SheetId(0), cell("H1"), "x", CalculationOptions::default())
            .unwrap();
        workbook.undo(CalculationOptions::default()).unwrap();
        let (events, _subscription) = recorded_updates(&workbook);
        let model = workbook.model().clone();
        let state = workbook.encode_state_as_update_v1();
        let history = workbook.history_state();
        let version = workbook.version();
        let refusal = refused(
            &mut workbook,
            json!([
                inputs("sheet:0", "A1", json!([["99"]])),
                {
                    "op": "setCellInputs",
                    "target": target("sheet:1", "A2"),
                    "inputs": [["1"]],
                    "expect": { "cells": [[{ "value": { "kind": "number", "value": 7 } }]] }
                },
            ]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::ContentMismatch);
        assert_eq!(refusal.failure.step_index, Some(1));
        assert_eq!(refusal.version, version);
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.model(), &model);
        assert_eq!(workbook.encode_state_as_update_v1(), state);
        assert_eq!(workbook.history_state(), history);
        assert!(events.lock().unwrap().is_empty());
        workbook.redo(CalculationOptions::default()).unwrap();
        assert_eq!(input(&workbook, 0, "H1"), "x");
    }
}

#[test]
fn stale_versions_are_refused() {
    let mut workbook = standalone();
    let stale = workbook.version();
    workbook
        .edit_cell(SheetId(0), cell("H1"), "x", CalculationOptions::default())
        .unwrap();
    let batch = request(
        &stale,
        json!({ "steps": [inputs("sheet:0", "A1", json!([["1"]]))] }),
    );
    let refusal = workbook.apply_edits(&batch).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);
    assert_eq!(refusal.version, workbook.version());
    let refusal = workbook.validate_edits(&batch).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);
}

#[test]
fn targets_resolve_through_the_catalog_and_exact_rectangles() {
    let mut workbook = standalone();
    let refusal = refused(&mut workbook, json!([inputs("Data", "A1", json!([["1"]]))]));
    assert_eq!(refusal.failure.code, EditFailureCode::MissingTarget);
    assert_eq!(refusal.failure.target.unwrap().sheet_id, "Data");
    for address in [
        "Data!A1", "A1,B2", "A:A", "1:1", "Total", "B2:A1", "A1:B2:C3", " A1",
    ] {
        let refusal = refused(
            &mut workbook,
            json!([inputs("sheet:0", address, json!([["1"]]))]),
        );
        assert_eq!(
            refusal.failure.code,
            EditFailureCode::InvalidStep,
            "{address}"
        );
    }
    let refusal = refused(
        &mut workbook,
        json!([{
            "op": "setCellInputs",
            "target": { "sheetId": "sheet:0", "range": { "kind": "rowCol", "start": { "row": 0, "col": 0 }, "end": { "row": 1048576, "col": 0 } } },
            "inputs": [["1"]]
        }]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep);
    let refusal = refused(
        &mut workbook,
        json!([inputs("sheet:0", "A1:B1", json!([["1"]]))]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep);
    assert!(refusal.failure.message.contains("1 rows of 2 cells"));

    let result = applied(
        &mut workbook,
        json!([
            inputs("sheet:1", "$b$2", json!([["lower"]])),
            {
                "op": "setCellInputs",
                "target": { "sheetId": "sheet:1", "range": { "kind": "rowCol", "start": { "row": 2, "col": 1 }, "end": { "row": 2, "col": 2 } } },
                "inputs": [["x", "y"]]
            },
        ]),
    );
    assert_eq!(input(&workbook, 1, "B2"), "lower");
    assert_eq!(input(&workbook, 1, "C3"), "y");
    assert_eq!(
        serde_json::to_value(&result.receipts[0].target).unwrap(),
        target("sheet:1", "B2")
    );
    assert_eq!(
        serde_json::to_value(&result.receipts[1].target).unwrap(),
        target("sheet:1", "B3:C3")
    );
}

#[test]
fn guards_read_the_pre_batch_state() {
    let mut workbook = standalone();
    let guard = |cells: Value| json!({ "cells": [[cells]] });
    let result = applied(
        &mut workbook,
        json!([
            inputs("sheet:0", "A1", json!([["99"]])),
            {
                "op": "patchStyle",
                "target": target("sheet:0", "B1"),
                "patch": { "italic": true },
                "expect": guard(json!({ "value": { "kind": "number", "value": 15 }, "formula": "SUM(A1:A2)", "displayText": "15" }))
            },
            {
                "op": "patchStyle",
                "target": target("sheet:0", "A2"),
                "patch": { "italic": true },
                "expect": guard(json!({ "formula": null }))
            },
            {
                "op": "setCellInputs",
                "target": target("sheet:0", "D1"),
                "inputs": [["2"]],
                "expect": guard(json!({ "displayText": "1.50" }))
            },
        ]),
    );
    assert!(result.applied);
    assert_eq!(value(&workbook, 0, "B1"), number(104.0));

    for (expect, field) in [
        (json!({ "formula": null }), "formula"),
        (json!({ "formula": "SUM(A1:A3)" }), "formula"),
        (json!({ "displayText": "104.0" }), "display text"),
        (
            json!({ "value": { "kind": "text", "value": "104" } }),
            "value",
        ),
    ] {
        let refusal = refused(
            &mut workbook,
            json!([{ "op": "patchStyle", "target": target("sheet:0", "B1"), "patch": { "bold": true }, "expect": guard(expect) }]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::ContentMismatch);
        assert!(
            refusal.failure.message.contains(field),
            "{}",
            refusal.failure.message
        );
    }
    let refusal = refused(
        &mut workbook,
        json!([{ "op": "patchStyle", "target": target("sheet:0", "B1:B2"), "patch": { "bold": true }, "expect": guard(json!({})) }]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep);
}

#[test]
fn formulas_bypass_text_format_coercion_and_inputs_parse_before_formatting() {
    let mut workbook = standalone();
    applied(
        &mut workbook,
        json!([
            inputs("sheet:0", "E1", json!([["=1+1"]])),
            inputs("sheet:0", "H1", json!([["=2+2"]])),
            { "op": "setNumberFormat", "target": target("sheet:0", "H1"), "format": "plainText" },
        ]),
    );
    assert_eq!(value(&workbook, 0, "E1"), text("=1+1"));
    assert_eq!(value(&workbook, 0, "H1"), number(4.0));
    assert!(workbook.cell(SheetId(0), cell("H1")).unwrap().is_formula);

    applied(
        &mut workbook,
        json!([{ "op": "setFormulas", "target": target("sheet:0", "E1"), "formulas": [["1+1"]] }]),
    );
    assert_eq!(value(&workbook, 0, "E1"), number(2.0));
    assert_eq!(input(&workbook, 0, "E1"), "=1+1");

    for source in ["=1+1", "", "SUM(", "1+"] {
        let refusal = refused(
            &mut workbook,
            json!([{ "op": "setFormulas", "target": target("sheet:0", "E2"), "formulas": [[source]] }]),
        );
        assert_eq!(
            refusal.failure.code,
            EditFailureCode::InvalidStep,
            "{source:?}"
        );
    }
}

#[test]
fn formatting_composes_with_content_on_the_same_cells() {
    let mut workbook = standalone();
    let result = applied(
        &mut workbook,
        json!([
            inputs("sheet:0", "D1", json!([["2.5"]])),
            { "op": "setNumberFormat", "target": target("sheet:0", "D1"), "format": { "type": "increaseDecimal" } },
            { "op": "patchStyle", "target": target("sheet:0", "D1"), "patch": { "bold": true, "fillColor": "#ffcc00" } },
            { "op": "patchStyle", "target": target("sheet:0", "D1"), "patch": { "italic": true } },
        ]),
    );
    assert!(result.receipts.iter().all(|receipt| receipt.changed));
    assert_eq!(value(&workbook, 0, "D1"), number(2.5));
    assert_eq!(display(&workbook, "sheet:0", "D1"), "2.500");
    let formatting = workbook
        .selection_formatting(SheetId(0), CellRange::parse_a1("D1").unwrap())
        .unwrap();
    assert_eq!(formatting.bold, Some(true));
    assert_eq!(formatting.italic, Some(true));
    assert_eq!(formatting.fill_color.as_deref(), Some("#ffcc00"));

    let refusal = refused(
        &mut workbook,
        json!([{ "op": "patchStyle", "target": target("sheet:0", "D1"), "patch": { "textColor": "red" } }]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep);
}

#[test]
fn overlapping_writes_to_one_property_are_refused() {
    let mut workbook = standalone();
    for (first, second) in [
        (
            inputs("sheet:0", "A1:A2", json!([["1"], ["2"]])),
            json!({ "op": "setFormulas", "target": target("sheet:0", "A2"), "formulas": [["1"]] }),
        ),
        (
            json!({ "op": "setNumberFormat", "target": target("sheet:0", "A1:B2"), "format": "percent" }),
            json!({ "op": "setNumberFormat", "target": target("sheet:0", "B2"), "format": "number" }),
        ),
        (
            json!({ "op": "patchStyle", "target": target("sheet:0", "A1"), "patch": { "bold": true } }),
            json!({ "op": "patchStyle", "target": target("sheet:0", "A1"), "patch": { "clear": ["bold"] } }),
        ),
    ] {
        let refusal = refused(&mut workbook, json!([first, second]));
        assert_eq!(refusal.failure.code, EditFailureCode::OverlappingSteps);
        assert_eq!(refusal.failure.step_index, Some(1));
        assert_eq!(refusal.failure.conflicting_step_index, Some(0));
    }
    applied(
        &mut workbook,
        json!([
            { "op": "patchStyle", "target": target("sheet:0", "A1:A2"), "patch": { "bold": true } },
            { "op": "patchStyle", "target": target("sheet:0", "A2"), "patch": { "italic": true } },
            inputs("sheet:1", "A2", json!([["1"]])),
            inputs("sheet:0", "A3", json!([["1"]])),
        ]),
    );
}

#[test]
fn merged_array_and_protected_cells_refuse_writes() {
    for mut workbook in [standalone(), collaborative(9)] {
        let refusal = refused(
            &mut workbook,
            json!([inputs("sheet:0", "G1", json!([["x"]]))]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::LockedTarget);
        assert!(refusal.failure.message.contains("G1"));
        let refusal = refused(
            &mut workbook,
            json!([inputs("sheet:0", "F1:G1", json!([["x", ""]]))]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::LockedTarget);
        let refusal = refused(
            &mut workbook,
            json!([{ "op": "setFormulas", "target": target("sheet:0", "E4"), "formulas": [["1"]] }]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::LockedTarget);
        let refusal = refused(
            &mut workbook,
            json!([{ "op": "patchStyle", "target": target("sheet:2", "A1"), "patch": { "bold": true } }]),
        );
        assert_eq!(refusal.failure.code, EditFailureCode::LockedTarget);
        for (range, format) in [
            (
                "F1:G1",
                json!({ "op": "patchStyle", "patch": { "bold": true } }),
            ),
            (
                "E3",
                json!({ "op": "patchStyle", "patch": { "bold": true } }),
            ),
            (
                "E4",
                json!({ "op": "setNumberFormat", "format": "percent" }),
            ),
        ] {
            let mut step = format;
            step["target"] = target("sheet:0", range);
            let refusal = refused(&mut workbook, json!([step]));
            assert_eq!(
                refusal.failure.code,
                EditFailureCode::LockedTarget,
                "{range}"
            );
        }

        applied(
            &mut workbook,
            json!([
                inputs("sheet:0", "F1", json!([["owner"]])),
                { "op": "patchStyle", "target": target("sheet:0", "F1"), "patch": { "bold": true } },
                { "op": "patchStyle", "target": target("sheet:0", "E5:G6"), "patch": { "bold": true } },
            ]),
        );
        let catalog = workbook
            .read_cells(&ReadRequest::default())
            .unwrap()
            .unwrap();
        assert_eq!(
            catalog
                .sheets
                .iter()
                .map(|sheet| (sheet.name.as_str(), sheet.editable))
                .collect::<Vec<_>>(),
            [("Data", true), ("Totals", true), ("Locked", false)]
        );
    }
}

#[test]
fn history_none_keeps_existing_undo_and_redo_entries() {
    for mut workbook in [standalone(), collaborative(10)] {
        workbook
            .edit_cell(SheetId(0), cell("H1"), "one", CalculationOptions::default())
            .unwrap();
        workbook
            .edit_cell(SheetId(0), cell("H2"), "two", CalculationOptions::default())
            .unwrap();
        workbook.undo(CalculationOptions::default()).unwrap();
        let history = workbook.history_state();
        let batch = request(
            &workbook.version(),
            json!({ "history": "none", "source": "agent", "steps": [inputs("sheet:1", "D4", json!([["quiet"]]))] }),
        );
        let result = workbook.apply_edits(&batch).unwrap().unwrap();
        assert!(result.applied);
        assert_eq!(serde_json::to_value(result.source).unwrap(), "agent");
        assert_eq!(workbook.history_state(), history);
        workbook.redo(CalculationOptions::default()).unwrap();
        assert_eq!(input(&workbook, 0, "H2"), "two");
        assert_eq!(input(&workbook, 1, "D4"), "quiet");
        workbook.undo(CalculationOptions::default()).unwrap();
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_eq!(input(&workbook, 0, "H1"), "");
        assert_eq!(input(&workbook, 1, "D4"), "quiet");
    }
}

#[test]
fn standalone_history_none_does_not_shield_a_cell_from_an_older_undo() {
    let mut workbook = standalone();
    workbook
        .edit_cell(
            SheetId(0),
            cell("H1"),
            "typed",
            CalculationOptions::default(),
        )
        .unwrap();
    let batch = request(
        &workbook.version(),
        json!({ "history": "none", "steps": [inputs("sheet:0", "H1", json!([["batch"]]))] }),
    );
    workbook.apply_edits(&batch).unwrap().unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(input(&workbook, 0, "H1"), "");
}

#[test]
fn a_no_op_batch_changes_nothing() {
    for mut workbook in [standalone(), collaborative(11)] {
        workbook
            .edit_cell(SheetId(0), cell("H1"), "x", CalculationOptions::default())
            .unwrap();
        workbook.undo(CalculationOptions::default()).unwrap();
        let (events, _subscription) = recorded_updates(&workbook);
        let version = workbook.version();
        let history = workbook.history_state();
        let result = applied(
            &mut workbook,
            json!([
                inputs("sheet:0", "A1", json!([["10"]])),
                { "op": "setFormulas", "target": target("sheet:0", "B1"), "formulas": [["SUM(A1:A2)"]] },
                { "op": "patchStyle", "target": target("sheet:0", "A2"), "patch": {} },
            ]),
        );
        assert!(!result.applied);
        assert_eq!(result.version, version);
        assert!(result.receipts.iter().all(|receipt| !receipt.changed));
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.history_state(), history);
        assert!(events.lock().unwrap().is_empty());
        let empty = applied(&mut workbook, json!([]));
        assert!(!empty.applied);
    }
}

#[test]
fn observers_get_one_update_for_the_recalculated_commit() {
    let mut local = collaborative(12);
    let mut peer = collaborative(13);
    let (events, _subscription) = recorded_updates(&local);
    let updates = Arc::new(Mutex::new(Vec::new()));
    let captured = Arc::clone(&updates);
    let _bytes = local
        .observe_update_v1(move |event| captured.lock().unwrap().push(event.update))
        .unwrap();
    applied(&mut local, two_sheet_batch());
    assert_eq!(*events.lock().unwrap(), [UpdateOrigin::Local]);
    let peer_version = peer.version();
    for update in updates.lock().unwrap().iter() {
        peer.apply_update_v1(update, CalculationOptions::default())
            .unwrap();
    }
    assert_ne!(peer.version(), peer_version);
    assert_eq!(input(&peer, 0, "C3"), "=A1+A2");
    assert_eq!(value(&peer, 1, "A1"), number(100.0));
    assert_eq!(peer.model(), local.model());
}

#[test]
fn validation_previews_the_pre_batch_targets_and_reserves_nothing() {
    let mut workbook = standalone();
    let batch = steps(
        &workbook,
        json!([
            inputs("sheet:0", "A1:A2", json!([["10"], ["6"]])),
            { "op": "patchStyle", "target": target("sheet:0", "A1"), "patch": { "bold": true } },
        ]),
    );
    let model = workbook.model().clone();
    let validation = workbook.validate_edits(&batch).unwrap().unwrap();
    assert!(validation.would_apply);
    assert_eq!(validation.base_version, workbook.version());
    assert_eq!(validation.previews[0].changed_cell_count, 1);
    assert!(validation.previews[1].would_change);
    assert_eq!(workbook.model(), &model);
    workbook
        .edit_cell(SheetId(0), cell("H1"), "x", CalculationOptions::default())
        .unwrap();
    let refusal = workbook.apply_edits(&batch).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);
}

#[test]
fn remote_updates_and_history_invalidate_versions() {
    let mut local = collaborative(14);
    let mut peer = collaborative(15);
    let read = local.version();
    peer.edit_cell(
        SheetId(0),
        cell("H1"),
        "peer",
        CalculationOptions::default(),
    )
    .unwrap();
    let update = peer
        .encode_diff_v1(&local.encode_state_vector_v1())
        .unwrap();
    local
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_ne!(local.version(), read);

    let before_delete = local.version();
    peer.edit_cell(SheetId(0), cell("H1"), "", CalculationOptions::default())
        .unwrap();
    let deletion = peer
        .encode_diff_v1(&local.encode_state_vector_v1())
        .unwrap();
    local
        .apply_update_v1(&deletion, CalculationOptions::default())
        .unwrap();
    assert_eq!(input(&local, 0, "H1"), "");
    assert_ne!(local.version(), before_delete);

    let batch = steps(&local, json!([inputs("sheet:0", "H2", json!([["mine"]]))]));
    local.apply_edits(&batch).unwrap().unwrap();
    let applied = local.version();
    local.undo(CalculationOptions::default()).unwrap();
    assert_ne!(local.version(), applied);
    assert_eq!(input(&local, 0, "H2"), "");
}

#[test]
fn adopting_a_persisted_snapshot_rotates_the_version() {
    let bytes = fixture();
    let mut peer = Workbook::open_collaborative(&bytes, 21).unwrap();
    peer.edit_cell(
        SheetId(0),
        cell("H1"),
        "persisted",
        CalculationOptions::default(),
    )
    .unwrap();
    let mut fresh = Workbook::open_collaborative(&bytes, 22).unwrap();
    let before = fresh.version();
    fresh
        .apply_update_v1(
            &peer.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(input(&fresh, 0, "H1"), "persisted");
    let (nonce, _) = before.as_str().split_once('-').unwrap();
    assert!(!fresh.version().as_str().starts_with(nonce));
}

#[test]
fn successive_standalone_batches_keep_the_authority_formats() {
    let colors = [
        "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#00ffff", "#ff00ff", "#123456",
    ];
    let mut workbook = standalone();
    let styled = colors
        .iter()
        .enumerate()
        .map(|(row, color)| {
            json!({ "op": "patchStyle", "target": target("sheet:1", &format!("H{}", row + 1)), "patch": { "fillColor": color } })
        })
        .collect::<Vec<_>>();
    applied(&mut workbook, Value::Array(styled));
    let typed = (0..colors.len())
        .map(|row| {
            inputs(
                "sheet:1",
                &format!("H{}", row + 1),
                json!([[format!("v{row}")]]),
            )
        })
        .collect::<Vec<_>>();
    applied(&mut workbook, Value::Array(typed));
    let check = |workbook: &Workbook| {
        let mut replica = Workbook::open_collaborative(&fixture(), 41).unwrap();
        replica
            .apply_update_v1(
                &workbook.encode_state_as_update_v1(),
                CalculationOptions::default(),
            )
            .unwrap();
        let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
        for (row, color) in colors.iter().enumerate() {
            let address = format!("H{}", row + 1);
            let range = CellRange::parse_a1(&address).unwrap();
            for copy in [workbook, &replica, &reopened] {
                assert_eq!(
                    copy.selection_formatting(SheetId(1), range)
                        .unwrap()
                        .fill_color
                        .as_deref(),
                    Some(*color),
                    "{address}"
                );
                assert_eq!(input(copy, 1, &address), input(workbook, 1, &address));
            }
        }
    };
    workbook
        .edit_cell(
            SheetId(1),
            cell("H1"),
            "typed",
            CalculationOptions::default(),
        )
        .unwrap();
    check(&workbook);
    workbook.undo(CalculationOptions::default()).unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    check(&workbook);
}

#[test]
fn versions_are_scoped_to_one_authority() {
    let bytes = fixture();
    let mut first = Workbook::open_collaborative(&bytes, 1).unwrap();
    let twin = Workbook::open_collaborative(&bytes, 1).unwrap();
    let wide = Workbook::open_collaborative(&bytes, 4_294_967_297).unwrap();
    assert_ne!(first.version(), twin.version());
    assert_ne!(first.version(), wide.version());
    assert_ne!(twin.version(), wide.version());
    for other in [&twin, &wide] {
        let foreign = steps(other, json!([inputs("sheet:0", "H1", json!([["x"]]))]));
        let refusal = first.apply_edits(&foreign).unwrap().unwrap_err();
        assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);
    }
}

#[test]
fn formats_too_large_to_share_are_refused() {
    for mut workbook in [standalone(), collaborative(23)] {
        let version = workbook.version();
        let state = workbook.encode_state_as_update_v1();
        for step in [
            json!({ "op": "patchStyle", "target": target("sheet:1", "B2"), "patch": { "fontFamily": "F".repeat(70_000) } }),
            json!({ "op": "setNumberFormat", "target": target("sheet:1", "B2"), "format": { "type": "custom", "pattern": "0".repeat(70_000) } }),
        ] {
            let refusal = refused(&mut workbook, json!([step]));
            assert_eq!(refusal.failure.code, EditFailureCode::LimitExceeded);
            assert_eq!(refusal.failure.step_index, Some(0));
        }
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.encode_state_as_update_v1(), state);
    }
}

#[test]
fn changed_sheets_include_recalculated_dependents() {
    for mut workbook in [standalone(), collaborative(24)] {
        let result = applied(
            &mut workbook,
            json!([inputs("sheet:0", "A2", json!([["6"]]))]),
        );
        assert_eq!(result.changed_sheets, ["sheet:0", "sheet:1"]);
        assert!(
            result
                .calculation
                .changed
                .iter()
                .any(|cell| cell.sheet_id == "sheet:1" && cell.a1 == "A1")
        );
        assert_eq!(value(&workbook, 1, "A1"), number(32.0));
    }
}

#[test]
fn cycles_and_calculation_limits_reach_the_result() {
    let mut workbook = standalone();
    let result = applied(
        &mut workbook,
        json!([
            { "op": "setFormulas", "target": target("sheet:0", "H1:H2"), "formulas": [["H2"], ["H1"]] },
            inputs("sheet:1", "XFD1048576", json!([["1"]])),
            { "op": "setFormulas", "target": target("sheet:0", "H3"), "formulas": [["SUM(Totals!A1:XFD1048576)"]] },
        ]),
    );
    let a1s = |cells: &[CellTarget]| cells.iter().map(|cell| cell.a1.clone()).collect::<Vec<_>>();
    let cycles = a1s(&result.calculation.cycle_cells);
    assert!(cycles.contains(&"H1".to_owned()) && cycles.contains(&"H2".to_owned()));
    assert_eq!(a1s(&result.calculation.limited_cells), ["H3"]);
    assert!(!result.calculation.truncated);
    let read = workbook
        .read_cells(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert_eq!(a1s(&read.calculation.limited_cells), ["H3"]);
}

#[test]
fn calculation_diagnostics_are_capped() {
    let mut workbook = standalone();
    let formulas = vec![vec!["$A$1+1".to_owned()]; 10_001];
    applied(
        &mut workbook,
        json!([{ "op": "setFormulas", "target": target("sheet:1", "B1:B10001"), "formulas": formulas }]),
    );
    let result = applied(
        &mut workbook,
        json!([inputs("sheet:0", "A1", json!([["11"]]))]),
    );
    assert_eq!(result.calculation.changed.len(), 10_000);
    assert!(result.calculation.truncated);
    assert_eq!(value(&workbook, 1, "B10001"), number(33.0));
}

#[test]
fn positional_sheet_ids_follow_the_catalog_version() {
    let mut workbook = standalone();
    let read = workbook
        .read_cells(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert_eq!(read.sheets[0].name, "Data");
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "Front".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let stale = request(
        &read.version,
        json!({ "steps": [inputs("sheet:0", "A1", json!([["1"]]))] }),
    );
    assert_eq!(
        workbook
            .apply_edits(&stale)
            .unwrap()
            .unwrap_err()
            .failure
            .code,
        EditFailureCode::StaleVersion
    );
    let catalog = workbook
        .read_cells(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert_eq!(
        catalog
            .sheets
            .iter()
            .map(|sheet| (sheet.sheet_id.as_str(), sheet.name.as_str()))
            .collect::<Vec<_>>(),
        [
            ("sheet:0", "Front"),
            ("sheet:1", "Data"),
            ("sheet:2", "Totals"),
            ("sheet:3", "Locked")
        ]
    );
    applied(
        &mut workbook,
        json!([inputs("sheet:0", "A1", json!([["front"]]))]),
    );
    assert_eq!(input(&workbook, 0, "A1"), "front");
    assert_eq!(input(&workbook, 1, "A1"), "10");
}

#[test]
fn every_source_and_history_combination_applies() {
    for collaborative_mode in [false, true] {
        for source in ["host", "agent"] {
            for history in ["separate", "none"] {
                let mut workbook = if collaborative_mode {
                    collaborative(30)
                } else {
                    standalone()
                };
                let depth = workbook.history_state().undo_depth;
                let batch = request(
                    &workbook.version(),
                    json!({ "source": source, "history": history, "steps": [inputs("sheet:1", "D4", json!([["x"]]))] }),
                );
                let result = workbook.apply_edits(&batch).unwrap().unwrap();
                assert_eq!(serde_json::to_value(result.source).unwrap(), source);
                let recorded = history == "separate";
                assert_eq!(
                    workbook.history_state().undo_depth,
                    depth + usize::from(recorded)
                );
                workbook.undo(CalculationOptions::default()).unwrap();
                assert_eq!(input(&workbook, 1, "D4"), if recorded { "" } else { "x" });
            }
        }
    }
}

#[test]
fn json_entry_points_bound_requests_and_decode_strictly() {
    let mut workbook = standalone();
    let version = workbook.version();
    let huge = format!(
        r#"{{"expectVersion":"{version}","steps":[],"padding":"{}"}}"#,
        "x".repeat(MAX_REQUEST_BYTES)
    );
    for response in [
        workbook.apply_edits_json(&huge).unwrap(),
        workbook.validate_edits_json(&huge).unwrap(),
        workbook.read_cells_json(&huge).unwrap(),
        workbook.find_text_json(&huge).unwrap(),
    ] {
        let response: Value = serde_json::from_str(&response).unwrap();
        assert_eq!(response["ok"], false);
        assert_eq!(response["failure"]["code"], "limit-exceeded");
        assert_eq!(response["version"], version.as_str());
    }
    assert!(matches!(
        workbook.apply_edits_json("{"),
        Err(Error::InvalidRequest(_))
    ));
    let request = serde_json::to_string(&steps(
        &workbook,
        json!([inputs("sheet:0", "A1", json!([["12"]]))]),
    ))
    .unwrap();
    let applied: Value =
        serde_json::from_str(&workbook.apply_edits_json(&request).unwrap()).unwrap();
    assert_eq!(applied["ok"], true);
    assert_eq!(applied["version"], workbook.version().as_str());
}

#[test]
fn requests_serialize_to_the_wire_shape() {
    let wire = json!({
        "expectVersion": standalone().version(),
        "source": "agent",
        "history": "none",
        "calculation": { "nowSerial": 45000.5 },
        "steps": [
            {
                "op": "setCellInputs",
                "target": target("sheet:0", "A1"),
                "inputs": [["1"]],
                "expect": { "cells": [[{ "value": { "kind": "number", "value": 10.0 }, "formula": null, "displayText": "10" }]] }
            },
            {
                "op": "setFormulas",
                "target": { "sheetId": "sheet:0", "range": { "kind": "rowCol", "start": { "row": 0, "col": 1 }, "end": { "row": 0, "col": 1 } } },
                "formulas": [["A1*2"]]
            },
            { "op": "setNumberFormat", "target": target("sheet:0", "A2"), "format": { "type": "custom", "pattern": "0.0" } },
            { "op": "patchStyle", "target": target("sheet:0", "A3"), "patch": { "bold": true } },
        ],
    });
    let request: EditRequest = serde_json::from_value(wire.clone()).unwrap();
    let serialized = serde_json::to_value(&request).unwrap();
    assert_eq!(serialized, wire);
    assert_eq!(
        serde_json::from_value::<EditRequest>(serialized).unwrap(),
        request
    );
}

#[test]
fn oversized_batches_are_refused() {
    let mut workbook = standalone();
    let many = (0..129)
        .map(|row| inputs("sheet:1", &format!("A{}", row + 1), json!([["1"]])))
        .collect::<Vec<_>>();
    let refusal = refused(&mut workbook, Value::Array(many));
    assert_eq!(refusal.failure.code, EditFailureCode::LimitExceeded);
    let refusal = refused(
        &mut workbook,
        json!([{ "op": "patchStyle", "target": target("sheet:1", "A1:Z4000"), "patch": { "bold": true } }]),
    );
    assert_eq!(refusal.failure.code, EditFailureCode::LimitExceeded);
    assert_eq!(refusal.failure.step_index, Some(0));
}

#[test]
fn reads_and_searches_carry_their_version() {
    let workbook = standalone();
    let read = workbook
        .read_cells(
            &serde_json::from_value(json!({ "ranges": [target("sheet:0", "B1:D1")] })).unwrap(),
        )
        .unwrap()
        .unwrap();
    assert_eq!(read.version, workbook.version());
    let row = &read.ranges[0].cells[0];
    assert_eq!(row[0].formula.as_deref(), Some("SUM(A1:A2)"));
    assert_eq!(row[0].value, number(15.0));
    assert_eq!(row[1].value, text("hello"));
    assert_eq!(row[2].display_text, "1.50");
    assert_eq!(read.sheets.len(), 3);

    let find = |request: Value| {
        workbook
            .find_text(&serde_json::from_value::<FindRequest>(request).unwrap())
            .unwrap()
    };
    let found = find(json!({ "text": "ell" })).unwrap();
    assert_eq!(
        found
            .matches
            .iter()
            .map(|found| (found.cell.a1.as_str(), found.text.as_str()))
            .collect::<Vec<_>>(),
        [("C1", "hello"), ("C2", "shell")]
    );
    assert!(!found.truncated);
    assert!(find(json!({ "text": "Hello" })).unwrap().matches.is_empty());
    let limited = find(json!({ "text": "ell", "limit": 1 })).unwrap();
    assert_eq!(limited.matches.len(), 1);
    assert!(limited.truncated);
    assert!(
        find(json!({ "text": "ell", "sheetIds": ["sheet:1"] }))
            .unwrap()
            .matches
            .is_empty()
    );
    assert_eq!(
        find(json!({ "text": "ell", "sheetIds": ["nope"] }))
            .unwrap_err()
            .failure
            .code,
        EditFailureCode::MissingTarget
    );
    assert_eq!(
        find(json!({ "text": "" })).unwrap_err().failure.code,
        EditFailureCode::InvalidStep
    );
}

#[test]
fn batches_round_trip_through_save_and_reopen() {
    let source = fixture();
    let mut workbook = Workbook::open(&source).unwrap();
    applied(
        &mut workbook,
        json!([
            inputs("sheet:0", "A1", json!([["40"]])),
            { "op": "setFormulas", "target": target("sheet:0", "C3"), "formulas": [["B1/5"]] },
            { "op": "patchStyle", "target": target("sheet:0", "C3"), "patch": { "bold": true } },
        ]),
    );
    let stale = workbook.version();
    let saved = workbook.save().unwrap();
    let reopened = Workbook::open_recalculated(&saved, CalculationOptions::default()).unwrap();
    assert_eq!(input(&reopened, 0, "C3"), "=B1/5");
    assert_eq!(value(&reopened, 0, "C3"), number(9.0));
    assert_eq!(value(&reopened, 1, "A1"), number(90.0));
    assert_eq!(
        reopened
            .selection_formatting(SheetId(0), CellRange::parse_a1("C3").unwrap())
            .unwrap()
            .bold,
        Some(true)
    );
    let parts = |bytes: &[u8]| -> BTreeMap<String, Vec<u8>> {
        ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
    };
    let (before, after) = (parts(&source), parts(&saved));
    for part in [
        "customXml/item1.xml",
        "xl/opaque/extension.bin",
        "xl/worksheets/sheet3.xml",
    ] {
        assert_eq!(before[part], after[part], "{part}");
    }
    assert!(String::from_utf8_lossy(&after["xl/worksheets/sheet1.xml"]).contains("mergeCell"));
    assert!(String::from_utf8_lossy(&after["xl/_rels/workbook.xml.rels"]).contains("rIdCustom"));
    let mut reopened = reopened;
    let refusal = reopened
        .apply_edits(&request(
            &stale,
            json!({ "steps": [inputs("sheet:0", "A1", json!([["1"]]))] }),
        ))
        .unwrap()
        .unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);
}

#[test]
fn requests_decode_strictly() {
    let decode = |value: Value| serde_json::from_value::<EditRequest>(value);
    let version = standalone().version();
    assert!(decode(json!({ "steps": [] })).is_err());
    assert!(decode(json!({ "expectVersion": version, "steps": [], "extra": 1 })).is_err());
    assert!(
        decode(json!({ "expectVersion": version, "steps": [{ "op": "insertRows" }] })).is_err()
    );
    assert!(
        decode(
            json!({ "expectVersion": version, "steps": [inputs("sheet:0", "A1", json!([[1]]))] })
        )
        .is_err()
    );
    assert!(decode(json!({ "expectVersion": version, "history": "later", "steps": [] })).is_err());
    let decoded = decode(json!({
        "expectVersion": version,
        "calculation": { "nowSerial": 45000 },
        "steps": [
            { "op": "setNumberFormat", "target": target("sheet:0", "A1"), "format": "percent" },
            { "op": "setNumberFormat", "target": target("sheet:0", "A2"), "format": { "type": "custom", "pattern": "0.0" } },
            { "op": "patchStyle", "target": target("sheet:0", "A3"), "patch": {}, "expect": { "cells": [[{ "formula": null }]] } },
        ],
    }))
    .unwrap();
    assert_eq!(decoded.calculation.now_serial, Some(45000.0));
    assert_eq!(
        decoded.steps[2].expect.as_ref().unwrap().cells[0][0].formula,
        Some(None)
    );
}
