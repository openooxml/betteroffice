use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, EditOutcome, EditRequest, ReplaceStats, Sheet,
    SheetId, Workbook, WorkbookModel,
};
use serde_json::{Value, json};
use xlsx_model::{CellRange, Protection, Xf};

const SHEET: &str = "sheet:0";

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn put(sheet: &mut Sheet, address: &str, value: CellValue, formula: Option<&str>, style: Option<u32>) {
    sheet.set_cell(
        cell(address),
        Cell {
            value,
            formula: formula.map(str::to_owned),
            style,
        },
    );
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.to_owned(),
    }
}

fn fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Xf::default());
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(49),
        ..Xf::default()
    });
    model.styles.cell_xfs.push(Xf {
        protection: Some(Protection {
            locked: false,
            hidden: false,
        }),
        ..Xf::default()
    });
    let mut data = Sheet::new("Data");
    put(&mut data, "A1", text("Alpha beta alpha"), None, None);
    put(&mut data, "A2", text("alpha"), None, None);
    put(&mut data, "B1", CellValue::Number { value: 10.0 }, None, None);
    put(&mut data, "C1", text("alpha"), Some("\"alpha\""), None);
    put(&mut data, "D1", text("alpha"), None, Some(1));
    put(&mut data, "E1", text("alpha alpha"), None, None);
    put(&mut data, "F1", text("alpha"), None, None);
    put(&mut data, "G1", text("alpha"), None, None);
    put(&mut data, "H1", text("gamma"), None, None);
    put(&mut data, "I1", text("alpha"), None, None);
    put(&mut data, "J1", text("alpha"), None, None);
    put(&mut data, "K1", CellValue::Number { value: 6.0 }, Some("L1+1"), None);
    put(&mut data, "L1", CellValue::Number { value: 5.0 }, None, None);
    put(&mut data, "M1", CellValue::Number { value: 7.0 }, None, None);
    data.merges.push(CellRange::parse_a1("F1:G1").unwrap());
    data.set_array_formula(cell("I1"), CellRange::parse_a1("I1:J1").unwrap());
    let mut protected = Sheet::new("Protected");
    put(&mut protected, "A1", text("alpha"), None, None);
    put(&mut protected, "A2", text("alpha"), None, Some(2));
    model.sheets.extend([data, protected]);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet2.xml")
        .unwrap();
    let xml = String::from_utf8(sheet.1.clone())
        .unwrap()
        .replace("</sheetData>", r#"</sheetData><sheetProtection sheet="1"/>"#);
    sheet.1 = xml.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn workbook() -> Workbook {
    Workbook::open(&fixture()).unwrap()
}

fn replace(a1: &str, query: &str, replacement: &str, entire: bool) -> Value {
    replace_in(a1, query, replacement, entire, "values")
}

fn replace_in(
    a1: &str,
    query: &str,
    replacement: &str,
    entire: bool,
    look_in: &str,
) -> Value {
    json!({
        "op": "replaceText",
        "target": { "sheetId": SHEET, "range": { "kind": "a1", "a1": a1 } },
        "query": query,
        "replacement": replacement,
        "entireCell": entire,
        "lookIn": look_in,
    })
}

fn apply(workbook: &mut Workbook, step: Value) -> EditOutcome {
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [step],
    }))
    .unwrap();
    workbook.apply_edits(&request).unwrap()
}

fn stored(workbook: &Workbook, a1: &str) -> CellValue {
    stored_on(workbook, SheetId(0), a1)
}

fn stored_on(workbook: &Workbook, sheet: SheetId, a1: &str) -> CellValue {
    workbook
        .sheet(sheet)
        .unwrap()
        .cell(cell(a1))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn style(workbook: &Workbook, a1: &str) -> Option<u32> {
    workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell(a1))
        .and_then(|cell| cell.style)
}

fn formula(workbook: &Workbook, a1: &str) -> Option<String> {
    workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell(a1))
        .and_then(|cell| cell.formula.clone())
}

#[test]
fn replaces_literal_text_in_range_and_preserves_style() {
    let mut workbook = workbook();
    let outcome = apply(&mut workbook, replace("A1:A2", "alpha", "ALPHA", false))
        .expect("the batch applies");

    assert_eq!(outcome.receipts.len(), 1);
    assert!(outcome.receipts[0].changed);
    assert_eq!(outcome.receipts[0].changed_cells.len(), 2);
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta ALPHA"));
    assert_eq!(stored(&workbook, "A2"), text("ALPHA"));
    assert_eq!(stored(&workbook, "E1"), text("alpha alpha"));
    assert_eq!(style(&workbook, "D1"), Some(1));

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta alpha"));
    assert_eq!(stored(&workbook, "A2"), text("alpha"));

    let outcome = apply(&mut workbook, replace("D1:D1", "alpha", "beta", false))
        .expect("the batch applies");
    assert!(outcome.receipts[0].changed);
    assert_eq!(stored(&workbook, "D1"), text("beta"));
    assert_eq!(style(&workbook, "D1"), Some(1));

    let outcome = apply(&mut workbook, replace("D1:D1", "beta", "", false))
        .expect("the batch applies");
    assert!(outcome.receipts[0].changed);
    assert_eq!(stored(&workbook, "D1"), CellValue::Empty);
    assert_eq!(style(&workbook, "D1"), Some(1));

    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(stored(&reopened, "D1"), CellValue::Empty);
    assert_eq!(style(&reopened, "D1"), Some(1));
}

#[test]
fn skips_formulas_non_text_merged_followers_and_untouched_columns() {
    let mut workbook = workbook();
    let original_formula = workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell("C1"))
        .unwrap()
        .formula
        .clone();

    let outcome = apply(&mut workbook, replace("A1:H1", "alpha", "omega", false))
        .expect("the batch applies");

    assert_eq!(outcome.receipts[0].changed_cells.len(), 4);
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta omega"));
    assert_eq!(stored(&workbook, "B1"), CellValue::Number { value: 10.0 });
    assert_eq!(stored(&workbook, "C1"), text("alpha"));
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("C1"))
            .unwrap()
            .formula,
        original_formula
    );
    assert_eq!(stored(&workbook, "G1"), text("alpha"));
    assert_eq!(stored(&workbook, "H1"), text("gamma"));
}

#[test]
fn entire_cell_matches_only_the_whole_stored_text() {
    let mut workbook = workbook();
    let outcome = apply(&mut workbook, replace("A1:H1", "alpha", "beta", true))
        .expect("the batch applies");

    assert_eq!(outcome.receipts[0].changed_cells.len(), 2);
    assert_eq!(stored(&workbook, "D1"), text("beta"));
    assert_eq!(stored(&workbook, "F1"), text("beta"));
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta alpha"));
    assert_eq!(stored(&workbook, "E1"), text("alpha alpha"));
    assert_eq!(stored(&workbook, "H1"), text("gamma"));
}

#[test]
fn skips_merged_followers_and_array_fills_and_rejects_guards() {
    let mut workbook = workbook();
    let outcome = apply(&mut workbook, replace("G1:G1", "alpha", "beta", false))
        .expect("a merged follower is skipped");
    assert!(!outcome.receipts[0].changed);
    assert_eq!(stored(&workbook, "G1"), text("alpha"));

    let guarded = json!({
        "op": "replaceText",
        "target": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A2:A2" } },
        "query": "alpha",
        "replacement": "beta",
        "expect": { "cells": [[{ "value": { "kind": "text", "value": "alpha" } }]] },
    });
    let refusal = apply(&mut workbook, guarded).expect_err("replaceText takes no guards");
    assert_eq!(refusal.failure.code, betteroffice_xlsx::EditFailureCode::InvalidStep);
}

#[test]
fn refuses_empty_query_and_stale_version_without_changing_anything() {
    let mut workbook = workbook();
    let refusal = apply(&mut workbook, replace("A1:A2", "", "beta", false))
        .expect_err("an empty query is invalid");
    assert_eq!(refusal.failure.code, betteroffice_xlsx::EditFailureCode::InvalidStep);
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta alpha"));

    let stale = workbook.version();
    apply(&mut workbook, replace("A2:A2", "alpha", "beta", false))
        .expect("the batch applies");
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": stale,
        "steps": [replace("A1:A1", "alpha", "omega", false)],
    }))
    .unwrap();
    let refusal = workbook.apply_edits(&request).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, betteroffice_xlsx::EditFailureCode::StaleVersion);
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta alpha"));
    assert_eq!(stored(&workbook, "A2"), text("beta"));
}

#[test]
fn protected_sheets_replace_only_unlocked_matches() {
    let mut workbook = workbook();
    let step = json!({
        "op": "replaceText",
        "target": { "sheetId": "sheet:1", "range": { "kind": "a1", "a1": "A1:A2" } },
        "query": "alpha",
        "replacement": "beta",
    });
    let outcome = apply(&mut workbook, step).expect("the batch applies");
    assert_eq!(outcome.receipts[0].changed_cells.len(), 1);
    assert_eq!(stored_on(&workbook, SheetId(1), "A1"), text("alpha"));
    assert_eq!(stored_on(&workbook, SheetId(1), "A2"), text("beta"));
}

#[test]
fn a_large_sparse_range_scans_stored_cells_only() {
    let mut workbook = workbook();
    let outcome = apply(
        &mut workbook,
        replace("A1:XFD1048576", "alpha", "omega", false),
    )
    .expect("the batch applies");
    assert!(outcome.receipts[0].changed);
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta omega"));
    assert_eq!(stored(&workbook, "A2"), text("omega"));
}

#[test]
fn validation_reports_the_expanded_changed_cell_count() {
    let workbook = workbook();
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [replace("A1:A2", "alpha", "beta", false)],
    }))
    .unwrap();

    let validation = workbook.validate_edits(&request).unwrap().unwrap();
    assert!(validation.would_apply);
    assert_eq!(validation.previews.len(), 1);
    assert_eq!(validation.previews[0].changed_cell_count, 2);
    assert_eq!(stored(&workbook, "A1"), text("Alpha beta alpha"));
}

#[test]
fn values_mode_reports_formula_results_without_rewriting_them() {
    let workbook = workbook();
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [replace("C1:C1", "alpha", "omega", false)],
    }))
    .unwrap();

    let validation = workbook.validate_edits(&request).unwrap().unwrap();
    assert!(!validation.previews[0].would_change);
    assert_eq!(
        validation.previews[0].replace,
        Some(ReplaceStats {
            matched: 1,
            replaced: 0,
            skipped_formulas: 1,
            skipped_locked: 0,
        })
    );

    let mut workbook = workbook;
    let outcome = apply(&mut workbook, replace("C1:C1", "alpha", "omega", false))
        .expect("the batch applies");
    assert!(!outcome.receipts[0].changed);
    assert_eq!(
        outcome.receipts[0].replace,
        Some(ReplaceStats {
            matched: 1,
            replaced: 0,
            skipped_formulas: 1,
            skipped_locked: 0,
        })
    );
    assert_eq!(formula(&workbook, "C1").as_deref(), Some("\"alpha\""));
    assert_eq!(stored(&workbook, "C1"), text("alpha"));
    let wire = serde_json::to_value(&outcome).unwrap();
    assert_eq!(wire["receipts"][0]["replace"]["skippedFormulas"], 1);
    assert_eq!(wire["receipts"][0]["replace"]["replaced"], 0);
}

#[test]
fn formulas_mode_replaces_formula_source_and_recalculates() {
    let mut workbook = workbook();
    let outcome = apply(
        &mut workbook,
        replace_in("K1:K1", "L1", "M1", false, "formulas"),
    )
    .expect("the batch applies");

    assert_eq!(outcome.receipts[0].changed_cells.len(), 1);
    assert_eq!(
        outcome.receipts[0].replace,
        Some(ReplaceStats {
            matched: 1,
            replaced: 1,
            skipped_formulas: 0,
            skipped_locked: 0,
        })
    );
    assert_eq!(formula(&workbook, "K1").as_deref(), Some("M1+1"));
    assert_eq!(stored(&workbook, "K1"), CellValue::Number { value: 8.0 });

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(formula(&workbook, "K1").as_deref(), Some("L1+1"));
    assert_eq!(stored(&workbook, "K1"), CellValue::Number { value: 6.0 });

    let outcome = apply(
        &mut workbook,
        replace_in("K1:K1", "L1", "M1", false, "formulas"),
    )
    .expect("the batch applies");
    assert!(outcome.receipts[0].changed);
    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(formula(&reopened, "K1").as_deref(), Some("M1+1"));
    assert_eq!(stored(&reopened, "K1"), CellValue::Number { value: 8.0 });
}

#[test]
fn formulas_mode_also_replaces_text_cells() {
    let mut workbook = workbook();
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [
            replace_in("A2:A2", "alpha", "beta", false, "formulas"),
            replace_in("C1:C1", "alpha", "beta", false, "formulas"),
        ],
    }))
    .unwrap();
    let outcome = workbook.apply_edits(&request).unwrap().unwrap();

    assert_eq!(outcome.receipts.len(), 2);
    assert_eq!(outcome.receipts[0].changed_cells.len(), 1);
    assert_eq!(outcome.receipts[1].changed_cells.len(), 1);
    assert_eq!(stored(&workbook, "A2"), text("beta"));
    assert_eq!(formula(&workbook, "C1").as_deref(), Some("\"beta\""));
    assert_eq!(stored(&workbook, "C1"), text("beta"));
}

#[test]
fn invalid_formula_replacements_refuse_the_whole_batch() {
    let mut workbook = workbook();
    let before = workbook.version();
    let refusal = apply(
        &mut workbook,
        replace_in("K1:K1", "L1", "(", false, "formulas"),
    )
    .expect_err("the replacement formula does not parse");
    assert_eq!(
        refusal.failure.code,
        betteroffice_xlsx::EditFailureCode::InvalidStep
    );
    assert!(refusal.failure.message.contains("K1"));
    assert_eq!(workbook.version(), before);
    assert_eq!(formula(&workbook, "K1").as_deref(), Some("L1+1"));
    assert_eq!(stored(&workbook, "K1"), CellValue::Number { value: 6.0 });
}

#[test]
fn replace_rejects_unknown_look_in_modes() {
    let workbook = workbook();
    let step = json!({
        "op": "replaceText",
        "target": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A2:A2" } },
        "query": "alpha",
        "replacement": "beta",
        "lookIn": "comments",
    });
    let error = serde_json::from_value::<EditRequest>(json!({
        "expectVersion": workbook.version(),
        "steps": [step],
    }))
    .expect_err("an unknown lookIn is refused at the wire");
    assert!(error.to_string().contains("comments"));
    assert_eq!(stored(&workbook, "A2"), text("alpha"));
}

#[test]
fn find_text_scopes_to_a_range_and_searches_formula_source() {
    let workbook = workbook();
    let find = |request: Value| {
        workbook
            .find_text(&serde_json::from_value(request).unwrap())
            .unwrap()
            .unwrap()
    };
    let addresses = |request: Value| {
        find(request)
            .matches
            .into_iter()
            .map(|found| found.cell.a1)
            .collect::<Vec<_>>()
    };

    assert_eq!(
        addresses(json!({
            "text": "alpha",
            "range": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:C1" } },
        })),
        ["A1", "C1"]
    );
    assert_eq!(
        addresses(json!({
            "text": "alpha",
            "range": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:C1" } },
            "lookIn": "formulas",
        })),
        ["A1", "C1"]
    );
    assert_eq!(
        addresses(json!({
            "text": "\"alpha\"",
            "range": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:C1" } },
            "lookIn": "formulas",
        })),
        ["C1"]
    );
    assert!(
        addresses(json!({
            "text": "\"alpha\"",
            "range": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:C1" } },
            "lookIn": "values",
        }))
        .is_empty()
    );
    assert_eq!(
        addresses(json!({
            "text": "alpha",
            "range": { "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:A2" } },
            "entireCell": true,
        })),
        ["A2"]
    );
    let scope = json!({ "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:C1" } });
    assert_eq!(addresses(json!({ "text": "10", "range": scope })), ["B1"]);
    assert!(
        addresses(json!({ "text": "10", "range": scope, "lookIn": "values" })).is_empty(),
        "replaceable values skip numeric cells"
    );
    assert!(
        addresses(json!({ "text": "10", "range": scope, "lookIn": "formulas" })).is_empty(),
        "replaceable formulas skip numeric cells"
    );
}

fn display_workbook() -> Workbook {
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Xf::default());
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(4),
        ..Xf::default()
    });
    model.styles.cell_xfs.push(Xf {
        num_fmt_id: Some(14),
        ..Xf::default()
    });
    let mut data = Sheet::new("Data");
    put(&mut data, "A1", CellValue::Number { value: 1234.5 }, None, Some(1));
    put(&mut data, "A2", CellValue::Number { value: 43831.0 }, None, Some(2));
    put(&mut data, "A3", CellValue::Bool { value: true }, None, None);
    put(&mut data, "A4", CellValue::Number { value: 2469.0 }, Some("A1*2"), Some(1));
    put(&mut data, "A5", text("xABCx"), None, None);
    put(&mut data, "A6", text("xabcx"), None, None);
    model.sheets.push(data);
    let parts = xlsx_parse::serialize_workbook(&model).unwrap();
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn found(workbook: &Workbook, request: Value) -> Vec<String> {
    workbook
        .find_text(&serde_json::from_value(request).unwrap())
        .unwrap()
        .unwrap()
        .matches
        .into_iter()
        .map(|found| found.cell.a1)
        .collect()
}

#[test]
fn find_values_matches_the_displayed_text_of_every_cell_kind() {
    let workbook = display_workbook();
    let scope = json!({ "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:A6" } });

    assert_eq!(
        found(&workbook, json!({ "text": "1,234.50", "range": scope, "lookIn": "values" })),
        ["A1"],
        "a formatted number is found by its displayed text"
    );
    assert_eq!(
        found(&workbook, json!({ "text": "1/1/20", "range": scope, "lookIn": "values" })),
        ["A2"],
        "a date serial is found by its displayed date"
    );
    assert_eq!(
        found(&workbook, json!({ "text": "TRUE", "range": scope, "lookIn": "values" })),
        ["A3"],
        "a boolean is found by its displayed text"
    );
    assert_eq!(
        found(&workbook, json!({ "text": "2,469.00", "range": scope, "lookIn": "values" })),
        ["A4"],
        "a formula's numeric result is found by its displayed text"
    );
    assert_eq!(
        found(&workbook, json!({ "text": "A1*2", "range": scope, "lookIn": "formulas" })),
        ["A4"],
        "formulas mode still searches formula source"
    );
    assert!(
        found(&workbook, json!({ "text": "A1*2", "range": scope, "lookIn": "values" })).is_empty()
    );
}

#[test]
fn find_ignores_case_by_default_and_matches_exactly_under_match_case() {
    let workbook = display_workbook();
    let scope = json!({ "sheetId": SHEET, "range": { "kind": "a1", "a1": "A5:A6" } });

    assert_eq!(found(&workbook, json!({ "text": "abc", "range": scope })), ["A5", "A6"]);
    assert_eq!(
        found(&workbook, json!({ "text": "abc", "range": scope, "lookIn": "values" })),
        ["A5", "A6"]
    );
    assert_eq!(
        found(
            &workbook,
            json!({ "text": "ABC", "range": scope, "lookIn": "values", "entireCell": false })
        ),
        ["A5", "A6"]
    );
    assert_eq!(
        found(&workbook, json!({ "text": "abc", "range": scope, "matchCase": true })),
        ["A6"]
    );
    assert_eq!(
        found(
            &workbook,
            json!({ "text": "abc", "range": scope, "lookIn": "values", "matchCase": true })
        ),
        ["A6"]
    );
}

fn replace_case(a1: &str, query: &str, replacement: &str, match_case: bool) -> Value {
    json!({
        "op": "replaceText",
        "target": { "sheetId": SHEET, "range": { "kind": "a1", "a1": a1 } },
        "query": query,
        "replacement": replacement,
        "matchCase": match_case,
    })
}

#[test]
fn replace_ignores_case_by_default_and_matches_exactly_under_match_case() {
    let mut workbook = display_workbook();
    let outcome = apply(&mut workbook, replace("A5:A6", "abc", "z", false))
        .expect("the batch applies");
    assert_eq!(outcome.receipts[0].changed_cells.len(), 2);
    assert_eq!(stored(&workbook, "A5"), text("xzx"));
    assert_eq!(stored(&workbook, "A6"), text("xzx"));

    let mut workbook = display_workbook();
    let outcome = apply(&mut workbook, replace_case("A5:A6", "abc", "z", true))
        .expect("the batch applies");
    assert_eq!(outcome.receipts[0].changed_cells.len(), 1);
    assert_eq!(stored(&workbook, "A5"), text("xABCx"));
    assert_eq!(stored(&workbook, "A6"), text("xzx"));
}
