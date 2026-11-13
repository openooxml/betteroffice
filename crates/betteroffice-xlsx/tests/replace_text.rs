use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, EditOutcome, EditRequest, ReadRequest,
    ReplaceStats, Sheet, SheetId, Workbook, WorkbookModel,
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

fn refused_cells(stats: &ReplaceStats) -> Vec<(String, String)> {
    let wire = serde_json::to_value(stats).unwrap();
    wire["refusedCells"]
        .as_array()
        .into_iter()
        .flatten()
        .map(|refused| {
            let field = |key: &str| refused[key].as_str().unwrap_or_default().to_owned();
            (field("cell"), field("reason"))
        })
        .collect()
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
    assert_eq!(stored(&workbook, "A1"), text("ALPHA beta ALPHA"));
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
    assert_eq!(stored(&workbook, "A1"), text("omega beta omega"));
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
    assert_eq!(stored(&workbook, "A1"), text("omega beta omega"));
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
            ..ReplaceStats::default()
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
            ..ReplaceStats::default()
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
            ..ReplaceStats::default()
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
fn invalid_formula_replacements_refuse_only_that_cell() {
    let mut workbook = workbook();
    let before = workbook.version();
    let outcome = apply(
        &mut workbook,
        replace_in("K1:K1", "L1", "(", false, "formulas"),
    )
    .expect("a cell the replacement breaks is refused on its own");
    assert!(!outcome.receipts[0].changed);
    let stats = outcome.receipts[0].replace.clone().unwrap();
    assert_eq!(stats.matched, 1);
    assert_eq!(stats.replaced, 0);
    let refused = refused_cells(&stats);
    assert_eq!(refused.len(), 1);
    assert_eq!(refused[0].0, "K1");
    assert!(refused[0].1.contains("K1"), "{}", refused[0].1);
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
    assert_eq!(addresses(json!({ "text": "10", "range": scope, "lookIn": "values" })), ["B1"]);
    assert_eq!(
        addresses(json!({ "text": "10", "range": scope, "lookIn": "formulas" })),
        ["B1"],
        "formulas mode searches a number's entered text"
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

fn rich_workbook() -> Workbook {
    Workbook::open(&rich_source()).unwrap()
}

fn rich_source() -> Vec<u8> {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Sheet1")],
        ..WorkbookModel::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = br#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><rFont val="Arial"/><b/></rPr><t>alpha</t></r><r><rPr><rFont val="Courier New"/><i/></rPr><t xml:space="preserve"> beta</t></r></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c></row></sheetData></worksheet>"#.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn replace_refuses_rich_text_cells_in_a_collaborative_session() {
    let mut workbook = Workbook::open_collaborative(&rich_source(), 303).unwrap();
    assert_eq!(workbook.cell_text_runs(SheetId(0), cell("A1")), None);

    let outcome = apply(&mut workbook, replace("A1:A2", "alpha", "omega", false))
        .expect("the rich text cell is refused and the rest applies");
    let stats = outcome.receipts[0].replace.clone().unwrap();
    assert_eq!(stats.matched, 2);
    assert_eq!(stats.replaced, 1);
    let refused = refused_cells(&stats);
    assert_eq!(refused.len(), 1);
    assert_eq!(refused[0].0, "A1");
    assert!(refused[0].1.contains("Sheet1!A1"), "{}", refused[0].1);
    assert_eq!(stored(&workbook, "A1"), text("alpha beta"));
    assert_eq!(stored(&workbook, "A2"), text("omega"));
}

#[test]
fn replace_refuses_rich_text_cells_and_keeps_their_runs() {
    let mut workbook = rich_workbook();
    let runs = workbook
        .cell_text_runs(SheetId(0), cell("A1"))
        .expect("A1 holds rich text")
        .to_vec();
    assert_eq!(runs.len(), 2);

    let outcome = apply(&mut workbook, replace("A1:A2", "alpha", "omega", false))
        .expect("the rich text cell is refused and the rest applies");
    let stats = outcome.receipts[0].replace.clone().unwrap();
    assert_eq!(stats.matched, 2);
    assert_eq!(stats.replaced, 1);
    let refused = refused_cells(&stats);
    assert_eq!(refused.len(), 1);
    assert_eq!(refused[0].0, "A1");
    assert!(refused[0].1.contains("Sheet1!A1"), "{}", refused[0].1);
    let wire = serde_json::to_value(&outcome).unwrap();
    assert_eq!(wire["receipts"][0]["replace"]["refusedCells"][0]["cell"], "A1");
    assert_eq!(stored(&workbook, "A1"), text("alpha beta"));
    assert_eq!(stored(&workbook, "A2"), text("omega"));
    assert_eq!(
        workbook.cell_text_runs(SheetId(0), cell("A1")),
        Some(runs.as_slice())
    );
}

fn number_workbook() -> Workbook {
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Xf::default());
    let mut data = Sheet::new("Data");
    put(&mut data, "A1", text("item 1"), None, None);
    put(&mut data, "A2", CellValue::Number { value: 123.0 }, None, None);
    put(&mut data, "A3", text("1 more"), None, None);
    model.sheets.push(data);
    let parts = xlsx_parse::serialize_workbook(&model).unwrap();
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

#[test]
fn replace_all_rewrites_numbers_through_their_entered_text() {
    let workbook = number_workbook();
    let scope = json!({ "sheetId": SHEET, "range": { "kind": "a1", "a1": "A1:A3" } });
    assert_eq!(
        found(&workbook, json!({ "text": "1", "range": scope, "lookIn": "values" })),
        ["A1", "A2", "A3"],
        "find shows the number by its displayed text"
    );

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [replace_in("A1:A3", "1", "2", false, "formulas")],
    }))
    .unwrap();
    let validation = workbook.validate_edits(&request).unwrap().unwrap();
    let preview = serde_json::to_value(&validation.previews[0]).unwrap();
    assert_eq!(preview["replace"]["matched"], 3);
    assert_eq!(preview["replace"]["replaced"], 3);
    assert_eq!(preview["replace"]["skippedValues"], 0);
    assert_eq!(preview["replace"]["skippedCells"], json!([]));

    let mut workbook = workbook;
    let outcome = apply(&mut workbook, typed("A1:A3", "1", "2"))
        .expect("the batch applies");
    assert_eq!(outcome.receipts[0].changed_cells.len(), 3);
    assert_eq!(stored(&workbook, "A1"), text("item 2"));
    assert_eq!(stored(&workbook, "A2"), CellValue::Number { value: 223.0 });
    assert_eq!(stored(&workbook, "A3"), text("2 more"));
    let wire = serde_json::to_value(&outcome).unwrap();
    assert_eq!(wire["receipts"][0]["replace"]["matched"], 3);
    assert_eq!(wire["receipts"][0]["replace"]["replaced"], 3);
    assert_eq!(wire["receipts"][0]["replace"]["refusedCells"], json!([]));
}

fn typed(a1: &str, query: &str, replacement: &str) -> Value {
    replace_in(a1, query, replacement, false, "formulas")
}

fn entry_workbook() -> Workbook {
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
    put(&mut data, "A1", CellValue::Number { value: 123.0 }, None, None);
    put(&mut data, "B1", CellValue::Number { value: 133.0 }, Some("A1+10"), None);
    put(&mut data, "C1", text("one 1"), None, None);
    put(&mut data, "D1", CellValue::Number { value: 1234.5 }, None, Some(1));
    put(&mut data, "E1", CellValue::Number { value: 43831.0 }, None, Some(2));
    put(&mut data, "F1", CellValue::Bool { value: true }, None, None);
    put(&mut data, "G1", CellValue::Number { value: 124.0 }, Some("A1+1"), None);
    put(&mut data, "H1", text("a+b"), None, None);
    model.sheets.push(data);
    let parts = xlsx_parse::serialize_workbook(&model).unwrap();
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn stats_of(outcome: &betteroffice_xlsx::EditApplication) -> ReplaceStats {
    outcome.receipts[0].replace.clone().expect("a replace reports stats")
}

#[test]
fn replace_enters_a_number_as_typed_and_it_stays_a_number() {
    let mut workbook = entry_workbook();
    let outcome = apply(&mut workbook, typed("A1:A1", "1", "2"))
        .expect("the batch applies");

    assert_eq!(stats_of(&outcome).replaced, 1);
    let value = stored(&workbook, "A1");
    assert!(matches!(value, CellValue::Number { .. }), "{value:?}");
    assert_eq!(value, CellValue::Number { value: 223.0 });
    assert_eq!(formula(&workbook, "A1"), None);
    assert_eq!(stored(&workbook, "B1"), CellValue::Number { value: 233.0 });
}

#[test]
fn replace_reenters_a_formula_and_recalculates_it() {
    let mut workbook = entry_workbook();
    let outcome = apply(&mut workbook, typed("B1:B1", "10", "20"))
        .expect("the batch applies");

    assert_eq!(stats_of(&outcome).replaced, 1);
    assert_eq!(formula(&workbook, "B1").as_deref(), Some("A1+20"));
    assert_eq!(stored(&workbook, "B1"), CellValue::Number { value: 143.0 });
}

#[test]
fn replace_keeps_replaced_text_as_text() {
    let mut workbook = entry_workbook();
    apply(&mut workbook, typed("C1:C1", "one", "two"))
        .expect("the batch applies");

    assert_eq!(stored(&workbook, "C1"), text("two 1"));
}

#[test]
fn a_replacement_that_breaks_a_formula_is_refused_and_the_rest_applies() {
    let mut workbook = entry_workbook();
    let outcome = apply(&mut workbook, typed("A1:H1", "+", "("))
        .expect("broken formulas are refused per cell");

    let stats = stats_of(&outcome);
    assert_eq!(stats.matched, 3);
    assert_eq!(stats.replaced, 1);
    let refused = refused_cells(&stats);
    let cells = refused.iter().map(|(cell, _)| cell.as_str()).collect::<Vec<_>>();
    assert_eq!(cells, ["B1", "G1"]);
    assert!(refused[0].1.contains("B1"), "{}", refused[0].1);
    assert_eq!(formula(&workbook, "B1").as_deref(), Some("A1+10"));
    assert_eq!(stored(&workbook, "B1"), CellValue::Number { value: 133.0 });
    assert_eq!(formula(&workbook, "G1").as_deref(), Some("A1+1"));
    assert_eq!(stored(&workbook, "H1"), text("a(b"));
}

#[test]
fn replace_reports_the_total_of_refused_cells() {
    let mut workbook = entry_workbook();
    let outcome = apply(&mut workbook, typed("A1:H1", "+", "("))
        .expect("broken formulas are refused per cell");

    let wire = serde_json::to_value(&outcome).unwrap();
    let stats = &wire["receipts"][0]["replace"];
    assert_eq!(stats["refused"], 2);
    assert_eq!(stats["refusedCells"].as_array().map(Vec::len), Some(2));

    let outcome = apply(&mut workbook, typed("A1:A1", "1", "2"))
        .expect("the batch applies");
    let wire = serde_json::to_value(&outcome).unwrap();
    assert_eq!(wire["receipts"][0]["replace"]["refused"], 0);
}

#[test]
fn a_formatted_number_is_found_by_its_display_but_replaced_by_its_entry() {
    let mut workbook = entry_workbook();
    let scope = json!({ "sheetId": SHEET, "range": { "kind": "a1", "a1": "D1:D1" } });
    assert_eq!(
        found(&workbook, json!({ "text": "1,234", "range": scope, "lookIn": "values" })),
        ["D1"]
    );

    for look_in in ["formulas", "values"] {
        let outcome = apply(
            &mut workbook,
            replace_in("D1:D1", "1,234", "9", false, look_in),
        )
        .expect("the batch applies");
        assert!(!outcome.receipts[0].changed);
        assert_eq!(stats_of(&outcome).matched, 0);
        assert_eq!(stored(&workbook, "D1"), CellValue::Number { value: 1234.5 });
    }
}

#[test]
fn replace_skips_and_reports_dates_and_booleans() {
    let mut workbook = entry_workbook();
    let outcome = apply(&mut workbook, typed("E1:F1", "1", "9"))
        .expect("the batch applies");
    let stats = stats_of(&outcome);
    assert!(!outcome.receipts[0].changed);
    assert_eq!(stats.skipped_values, 1);
    assert_eq!(stats.skipped_cells, ["E1"]);
    assert_eq!(stored(&workbook, "E1"), CellValue::Number { value: 43831.0 });

    let outcome = apply(&mut workbook, typed("E1:F1", "true", "no"))
        .expect("the batch applies");
    let stats = stats_of(&outcome);
    assert!(!outcome.receipts[0].changed);
    assert_eq!(stats.skipped_values, 1);
    assert_eq!(stats.skipped_cells, ["F1"]);
    assert_eq!(stored(&workbook, "F1"), CellValue::Bool { value: true });
}

#[test]
fn undo_after_replace_all_restores_every_cell_in_one_step() {
    let mut workbook = entry_workbook();
    let outcome = apply(&mut workbook, typed("A1:H1", "1", "2"))
        .expect("the batch applies");
    assert!(stats_of(&outcome).replaced >= 4);
    assert_eq!(stored(&workbook, "A1"), CellValue::Number { value: 223.0 });
    assert_eq!(formula(&workbook, "B1").as_deref(), Some("A2+20"));
    assert_eq!(stored(&workbook, "C1"), text("one 2"));
    assert_eq!(stored(&workbook, "D1"), CellValue::Number { value: 2234.5 });

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(stored(&workbook, "A1"), CellValue::Number { value: 123.0 });
    assert_eq!(formula(&workbook, "B1").as_deref(), Some("A1+10"));
    assert_eq!(stored(&workbook, "B1"), CellValue::Number { value: 133.0 });
    assert_eq!(stored(&workbook, "C1"), text("one 1"));
    assert_eq!(stored(&workbook, "D1"), CellValue::Number { value: 1234.5 });
    assert_eq!(formula(&workbook, "G1").as_deref(), Some("A1+1"));
}

fn two_sheet_source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Xf::default());
    let mut data = Sheet::new("Data");
    put(&mut data, "A1", CellValue::Number { value: 123.0 }, None, None);
    put(&mut data, "B1", CellValue::Number { value: 153.0 }, Some("A1+30"), None);
    put(&mut data, "C1", text("item 3"), None, None);
    let mut other = Sheet::new("Other");
    put(&mut other, "A1", CellValue::Number { value: 3.0 }, None, None);
    put(&mut other, "B1", text("3 more"), None, None);
    model.sheets.extend([data, other]);
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

fn whole_sheet_replace(sheet: &str, query: &str, replacement: &str) -> Value {
    json!({
        "op": "replaceText",
        "target": {
            "sheetId": sheet,
            "range": {
                "kind": "rowCol",
                "start": { "row": 0, "col": 0 },
                "end": { "row": 1048575, "col": 16383 },
            },
        },
        "query": query,
        "replacement": replacement,
        "lookIn": "formulas",
    })
}

fn replace_all_in_workbook(workbook: &mut Workbook) -> EditOutcome {
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [
            whole_sheet_replace("sheet:0", "3", "4"),
            whole_sheet_replace("sheet:1", "3", "4"),
        ],
    }))
    .unwrap();
    workbook.apply_edits(&request).unwrap()
}

fn formula_on(workbook: &Workbook, sheet: SheetId, a1: &str) -> Option<String> {
    workbook
        .sheet(sheet)
        .unwrap()
        .cell(cell(a1))
        .and_then(|cell| cell.formula.clone())
}

fn assert_original(workbook: &Workbook) {
    assert_eq!(stored_on(workbook, SheetId(0), "A1"), CellValue::Number { value: 123.0 });
    assert_eq!(formula_on(workbook, SheetId(0), "B1").as_deref(), Some("A1+30"));
    assert_eq!(stored_on(workbook, SheetId(0), "B1"), CellValue::Number { value: 153.0 });
    assert_eq!(stored_on(workbook, SheetId(0), "C1"), text("item 3"));
    assert_eq!(stored_on(workbook, SheetId(1), "A1"), CellValue::Number { value: 3.0 });
    assert_eq!(stored_on(workbook, SheetId(1), "B1"), text("3 more"));
}

fn assert_replaced(workbook: &Workbook) {
    assert_eq!(stored_on(workbook, SheetId(0), "A1"), CellValue::Number { value: 124.0 });
    assert_eq!(formula_on(workbook, SheetId(0), "B1").as_deref(), Some("A1+40"));
    assert_eq!(stored_on(workbook, SheetId(0), "B1"), CellValue::Number { value: 164.0 });
    assert_eq!(stored_on(workbook, SheetId(0), "C1"), text("item 4"));
    assert_eq!(stored_on(workbook, SheetId(1), "A1"), CellValue::Number { value: 4.0 });
    assert_eq!(stored_on(workbook, SheetId(1), "B1"), text("4 more"));
}

#[test]
fn redo_after_undoing_replace_all_restores_the_replaced_state_exactly() {
    let mut workbook = Workbook::open(&two_sheet_source()).unwrap();
    replace_all_in_workbook(&mut workbook).expect("the batch applies");
    assert_replaced(&workbook);
    let replaced = workbook.model().clone();

    assert!(workbook.undo(CalculationOptions::default()).unwrap().applied);
    assert_original(&workbook);
    assert!(workbook.redo(CalculationOptions::default()).unwrap().applied);
    assert_replaced(&workbook);
    assert_eq!(workbook.model(), &replaced);
}

#[test]
fn workbook_replace_all_across_two_sheets_is_one_undo_step() {
    let mut workbook = Workbook::open(&two_sheet_source()).unwrap();
    let depth = workbook.history_state().undo_depth;
    let outcome = replace_all_in_workbook(&mut workbook).expect("the batch applies");
    assert!(outcome.receipts[0].changed);
    assert!(outcome.receipts[1].changed);
    assert_eq!(workbook.history_state().undo_depth, depth + 1);
    assert_replaced(&workbook);

    assert!(workbook.undo(CalculationOptions::default()).unwrap().applied);
    assert_original(&workbook);
    assert_eq!(workbook.history_state().undo_depth, depth);
}

#[test]
fn a_peer_editing_a_replaced_cell_during_replace_all_converges() {
    let bytes = two_sheet_source();
    let options = CalculationOptions::default();
    for left_first in [true, false] {
        let mut left = Workbook::open_collaborative(&bytes, 301).unwrap();
        let mut right = Workbook::open_collaborative(&bytes, 302).unwrap();
        replace_all_in_workbook(&mut left).expect("the batch applies");
        assert!(
            right
                .edit_cell(SheetId(0), cell("C1"), "item 33", options)
                .unwrap()
                .applied
        );

        let to_right = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let to_left = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        if left_first {
            right.apply_update_v1(&to_right, options).unwrap();
            left.apply_update_v1(&to_left, options).unwrap();
        } else {
            left.apply_update_v1(&to_left, options).unwrap();
            right.apply_update_v1(&to_right, options).unwrap();
        }

        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert_eq!(left.model(), right.model(), "left_first={left_first}");
        for peer in [&left, &right] {
            assert_eq!(stored_on(peer, SheetId(0), "A1"), CellValue::Number { value: 124.0 });
            assert_eq!(formula_on(peer, SheetId(0), "B1").as_deref(), Some("A1+40"));
            assert_eq!(stored_on(peer, SheetId(1), "B1"), text("4 more"));
            let c1 = stored_on(peer, SheetId(0), "C1");
            assert!(c1 == text("item 4") || c1 == text("item 33"), "{c1:?}");
        }
    }
}

#[test]
fn replaced_cells_keep_their_kind_after_save_and_reopen() {
    let mut workbook = Workbook::open(&two_sheet_source()).unwrap();
    replace_all_in_workbook(&mut workbook).expect("the batch applies");
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();

    assert_replaced(&reopened);
    assert!(matches!(
        stored_on(&reopened, SheetId(0), "A1"),
        CellValue::Number { .. }
    ));
    assert_eq!(formula_on(&reopened, SheetId(0), "A1"), None);
    assert!(matches!(
        stored_on(&reopened, SheetId(1), "A1"),
        CellValue::Number { .. }
    ));
    assert!(matches!(
        stored_on(&reopened, SheetId(0), "C1"),
        CellValue::Text { .. }
    ));
    assert_eq!(formula_on(&reopened, SheetId(0), "C1"), None);
}

fn workbook_with_hidden_and_chartsheet() -> Workbook {
    let mut model = WorkbookModel::default();
    let mut data = Sheet::new("Data");
    put(&mut data, "A1", text("alpha"), None, None);
    let mut hidden = Sheet::new("Hidden");
    put(&mut hidden, "A1", text("alpha"), None, None);
    let mut chart = Sheet::new("Chart");
    put(&mut chart, "A1", text("alpha"), None, None);
    let mut secret = Sheet::new("Secret");
    put(&mut secret, "A1", text("alpha"), None, None);
    model.sheets.extend([data, hidden, chart, secret]);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    for (path, bytes) in &mut parts {
        if path == "xl/workbook.xml" {
            let xml = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(r#"name="Hidden""#, r#"name="Hidden" state="hidden""#)
                .replace(r#"name="Secret""#, r#"name="Secret" state="veryHidden""#);
            *bytes = xml.into_bytes();
        } else if path == "xl/_rels/workbook.xml.rels" {
            let xml = String::from_utf8(bytes.clone()).unwrap().replace(
                "relationships/worksheet\" Target=\"worksheets/sheet3.xml",
                "relationships/chartsheet\" Target=\"worksheets/sheet3.xml",
            );
            *bytes = xml.into_bytes();
        } else if path == "xl/worksheets/sheet3.xml" {
            *bytes = br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec();
        }
    }
    let workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let read = serde_json::to_value(
        workbook
            .read_cells(&ReadRequest::default())
            .unwrap()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(read["sheets"][1]["visibility"], "hidden");
    assert_eq!(read["sheets"][2]["editable"], false);
    assert_eq!(read["sheets"][3]["visibility"], "veryHidden");
    workbook
}

fn whole_sheet_within(sheet: &str, within: &str) -> Value {
    let mut step = whole_sheet_step(sheet);
    step["within"] = json!(within);
    step
}

fn whole_sheet_step(sheet: &str) -> Value {
    json!({
        "op": "replaceText",
        "target": {
            "sheetId": sheet,
            "range": {
                "kind": "rowCol",
                "start": { "row": 0, "col": 0 },
                "end": { "row": 1048575, "col": 16383 },
            },
        },
        "query": "alpha",
        "replacement": "omega",
    })
}

fn apply_steps(workbook: &mut Workbook, steps: Vec<Value>) -> Option<EditOutcome> {
    let request = serde_json::from_value::<EditRequest>(json!({
        "expectVersion": workbook.version(),
        "steps": steps,
    }))
    .ok()?;
    Some(workbook.apply_edits(&request).unwrap())
}

#[test]
fn replace_on_an_explicit_chartsheet_target_is_refused() {
    let mut workbook = workbook_with_hidden_and_chartsheet();
    let before = workbook.version();
    let outcome = apply_steps(
        &mut workbook,
        vec![whole_sheet_step("sheet:0"), whole_sheet_step("sheet:2")],
    )
    .expect("the request parses");
    let refusal = outcome.expect_err("a chartsheet target is not a worksheet");
    assert_eq!(refusal.failure.code, betteroffice_xlsx::EditFailureCode::Unsupported);
    assert_eq!(refusal.failure.step_index, Some(1));
    assert_eq!(workbook.version(), before);
    assert_eq!(stored_on(&workbook, SheetId(0), "A1"), text("alpha"));
}

#[test]
fn replace_without_within_searches_an_explicit_hidden_sheet() {
    let mut workbook = workbook_with_hidden_and_chartsheet();
    let steps = ["sheet:0", "sheet:1", "sheet:3"].map(whole_sheet_step).to_vec();
    let request = serde_json::from_value::<EditRequest>(json!({
        "expectVersion": workbook.version(),
        "steps": steps.clone(),
    }))
    .unwrap();
    let defaulted = serde_json::from_value::<EditRequest>(json!({
        "expectVersion": workbook.version(),
        "steps": ["sheet:0", "sheet:1", "sheet:3"]
            .map(|sheet| whole_sheet_within(sheet, "sheet")),
    }))
    .unwrap();
    assert_eq!(request, defaulted);
    assert!(
        serde_json::to_value(&request.steps[0]).unwrap()["within"].is_null(),
        "the default is not written back"
    );
    let outcome = apply_steps(&mut workbook, steps);
    let Some(Ok(applied)) = &outcome else {
        panic!("a replace without within applies: {outcome:?}");
    };
    assert_eq!(stored_on(&workbook, SheetId(0), "A1"), text("omega"));
    assert_eq!(stored_on(&workbook, SheetId(1), "A1"), text("omega"));
    assert_eq!(stored_on(&workbook, SheetId(3), "A1"), text("omega"));
    let matched = applied
        .receipts
        .iter()
        .map(|receipt| receipt.replace.as_ref().map_or(0, |stats| stats.matched))
        .collect::<Vec<_>>();
    assert_eq!(matched, [1, 1, 1]);
}

#[test]
fn workbook_replace_skips_hidden_and_very_hidden_sheets() {
    let mut workbook = workbook_with_hidden_and_chartsheet();
    let steps = ["sheet:0", "sheet:1", "sheet:2", "sheet:3"]
        .map(|sheet| whole_sheet_within(sheet, "workbook"))
        .to_vec();
    let outcome = apply_steps(&mut workbook, steps);
    let Some(Ok(applied)) = &outcome else {
        panic!("a workbook replace applies: {outcome:?}");
    };
    assert_eq!(stored_on(&workbook, SheetId(0), "A1"), text("omega"));
    assert_eq!(stored_on(&workbook, SheetId(1), "A1"), text("alpha"));
    assert_eq!(stored_on(&workbook, SheetId(3), "A1"), text("alpha"));
    let matched = applied
        .receipts
        .iter()
        .map(|receipt| receipt.replace.as_ref().map_or(0, |stats| stats.matched))
        .collect::<Vec<_>>();
    assert_eq!(matched, [1, 0, 0, 0]);
}

#[test]
fn find_over_all_sheets_skips_hidden_sheets_and_chartsheets() {
    let workbook = workbook_with_hidden_and_chartsheet();
    let matches = workbook
        .find_text(&serde_json::from_value(json!({ "text": "alpha" })).unwrap())
        .unwrap()
        .expect("a chartsheet does not refuse a workbook search")
        .matches;
    assert_eq!(
        matches
            .into_iter()
            .map(|found| (found.cell.sheet_id, found.cell.a1))
            .collect::<Vec<_>>(),
        [("sheet:0".to_owned(), "A1".to_owned())]
    );
}

#[test]
fn workbook_replace_with_a_chartsheet_does_not_refuse_the_batch() {
    let mut workbook = workbook_with_hidden_and_chartsheet();
    let whole = |sheet: &str| whole_sheet_within(sheet, "workbook");
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": workbook.version(),
        "steps": [whole("sheet:0"), whole("sheet:2")],
    }))
    .unwrap();

    let outcome = workbook
        .apply_edits(&request)
        .unwrap()
        .expect("a chartsheet in the sheet list is skipped, not refused");
    assert_eq!(outcome.receipts.len(), 2);
    assert!(outcome.receipts[0].changed);
    assert!(!outcome.receipts[1].changed);
    assert!(outcome.receipts[1].changed_cells.is_empty());
    assert_eq!(stored_on(&workbook, SheetId(0), "A1"), text("omega"));
    assert_eq!(stored_on(&workbook, SheetId(1), "A1"), text("alpha"));
}
