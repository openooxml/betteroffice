use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, Sheet, SheetId, Workbook, WorkbookModel,
};
use serde_json::json;
use xlsx_model::CellProvider;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn synthetic() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    for name in ["Sheet1", "Sheet2", "Sheet3"] {
        let mut sheet = Sheet::new(name);
        for (address, formula, cached) in [
            ("A1", "TODAY()", 45_000.0),
            ("B1", "A1+Sheet1!A2", -1.0),
            ("C1", "Sheet1!A2*2", -1.0),
        ] {
            sheet.set_cell(
                cell(address),
                Cell {
                    value: number(cached),
                    formula: Some(formula.into()),
                    style: None,
                },
            );
        }
        model.sheets.push(sheet);
    }
    model.sheets[0].set_cell(
        cell("A2"),
        Cell {
            value: number(1.0),
            ..Cell::default()
        },
    );
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = &mut parts
        .iter_mut()
        .find(|(name, _)| name == "xl/workbook.xml")
        .unwrap()
        .1;
    *workbook = std::str::from_utf8(workbook)
        .unwrap()
        .replace("</workbook>", "<calcPr calcMode=\"auto\"/></workbook>")
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_values(workbook: &Workbook, input: f64, today: f64) {
    assert!(workbook.model().defined_names.is_empty());
    assert_eq!(
        workbook.model().value(SheetId(0), cell("A2")),
        number(input)
    );
    for sheet in 0..3 {
        for (address, expected) in [("A1", today), ("B1", today + input), ("C1", input * 2.0)] {
            assert_eq!(
                workbook.model().value(SheetId(sheet), cell(address)),
                number(expected),
                "sheet {sheet}: {address}"
            );
        }
    }
}

fn save_without_errors(workbook: &Workbook) -> Vec<u8> {
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    for sheet in 1..=3 {
        let name = format!("xl/worksheets/sheet{sheet}.xml");
        let xml = &parts.iter().find(|(path, _)| path == &name).unwrap().1;
        assert!(
            !std::str::from_utf8(xml).unwrap().contains("t=\"e\""),
            "{name}"
        );
    }
    saved
}

fn edit(workbook: &mut Workbook, input: &str, batch: bool) {
    if batch {
        let request = json!({
            "expectVersion": workbook.version(),
            "steps": [{
                "op": "setCellInputs",
                "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "A2" } },
                "inputs": [[input]]
            }]
        });
        let result: serde_json::Value =
            serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap())
                .unwrap();
        assert_eq!(result["ok"], true);
        assert_eq!(result["applied"], true);
    } else {
        assert!(
            workbook
                .edit_cell(SheetId(0), cell("A2"), input, CalculationOptions::default())
                .unwrap()
                .applied
        );
    }
}

#[test]
fn clockless_edits_and_history_keep_only_clock_caches() {
    for batch in [false, true] {
        let mut workbook = Workbook::open(&synthetic()).unwrap();
        edit(&mut workbook, "7", batch);
        assert_values(&workbook, 7.0, 45_000.0);
        save_without_errors(&workbook);
        assert!(
            workbook
                .undo(CalculationOptions::default())
                .unwrap()
                .applied
        );
        assert_values(&workbook, 1.0, 45_000.0);
        save_without_errors(&workbook);
        assert!(
            workbook
                .redo(CalculationOptions::default())
                .unwrap()
                .applied
        );
        assert_values(&workbook, 7.0, 45_000.0);
        save_without_errors(&workbook);
        workbook.recalculate_all(CalculationOptions {
            now_serial: Some(46_000.25),
        });
        assert_values(&workbook, 7.0, 46_000.0);
    }
}

#[test]
fn clockless_remote_update_keeps_only_clock_caches() {
    let bytes = synthetic();
    let mut receiver = Workbook::open_collaborative(&bytes, 1).unwrap();
    let mut sender = Workbook::open_collaborative(&bytes, 2).unwrap();
    let vector = receiver.encode_state_vector_v1();
    edit(&mut sender, "7", false);
    assert_values(&sender, 7.0, 45_000.0);
    let update = sender.encode_diff_v1(&vector).unwrap();
    assert!(
        receiver
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_values(&receiver, 7.0, 45_000.0);
    save_without_errors(&receiver);
}

#[test]
fn clockless_save_reopen_edit_keeps_only_clock_caches() {
    let mut workbook = Workbook::open(&synthetic()).unwrap();
    edit(&mut workbook, "7", false);
    let saved = save_without_errors(&workbook);
    let mut reopened = Workbook::open(&saved).unwrap();
    assert_values(&reopened, 7.0, 45_000.0);
    edit(&mut reopened, "9", false);
    assert_values(&reopened, 9.0, 45_000.0);
    let saved = save_without_errors(&reopened);
    assert_values(&Workbook::open(&saved).unwrap(), 9.0, 45_000.0);
}
