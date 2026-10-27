use betteroffice_xlsx::{
    CalculationOptions, Op, ReadRequest, Sheet, Workbook, WorkbookModel, XlsxExportOptions,
};
use serde_json::{Value, json};

fn package() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.extend([
        Sheet::new("Visible"),
        Sheet::new("Hidden"),
        Sheet::new("VeryHidden"),
    ]);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/workbook.xml")
        .unwrap();
    workbook.1 = br#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" sheetId="1" state="visible" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/><sheet name="VeryHidden" sheetId="3" state="veryHidden" r:id="rId3"/></sheets></workbook>"#.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_visibility(workbook: &Workbook, expected: Value) {
    let read = workbook
        .read_cells(&ReadRequest::default())
        .unwrap()
        .unwrap();
    assert!(read.ranges.is_empty());
    let read = serde_json::to_value(read).unwrap();
    let exported = workbook
        .export_structured(&XlsxExportOptions {
            include_hidden_sheets: Some(true),
            ..XlsxExportOptions::default()
        })
        .unwrap()
        .unwrap();
    let exported = serde_json::to_value(exported.content).unwrap();
    let visibility = |sheets: &Value| {
        sheets
            .as_array()
            .unwrap()
            .iter()
            .map(|sheet| sheet["visibility"].clone())
            .collect::<Vec<_>>()
    };
    assert_eq!(json!(visibility(&read["sheets"])), expected);
    assert_eq!(visibility(&read["sheets"]), visibility(&exported["sheets"]));
    assert_eq!(
        read["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|sheet| sheet["name"].clone())
            .collect::<Vec<_>>(),
        exported["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|sheet| sheet["anchor"]["sheet"]["name"].clone())
            .collect::<Vec<_>>()
    );
}

#[test]
fn catalog_visibility_matches_package_and_structured_export() {
    let mut workbook = Workbook::open(&package()).unwrap();
    assert_visibility(&workbook, json!(["visible", "hidden", "veryHidden"]));
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Fresh".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_visibility(
        &workbook,
        json!(["visible", "visible", "hidden", "veryHidden"]),
    );
}

#[test]
fn catalog_visibility_is_unknown_for_models_and_visible_for_session_sheets() {
    let model = || {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Plain"));
        model
    };
    assert_visibility(
        &Workbook::from_model_collaborative(model(), 7).unwrap(),
        json!(["unknown"]),
    );
    let mut workbook = Workbook::from_model(model()).unwrap();
    assert_visibility(&workbook, json!(["unknown"]));
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Fresh".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_visibility(&workbook, json!(["unknown", "visible"]));
}
