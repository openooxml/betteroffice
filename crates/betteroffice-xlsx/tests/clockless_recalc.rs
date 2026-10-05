use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook,
    WorkbookModel,
};
use serde_json::json;
use xlsx_model::{DefinedName, ErrorValue};

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: number(1.0),
            ..Cell::default()
        },
    );
    model.sheets.extend([sheet, Sheet::new("Sheet2")]);
    model.defined_names.push(DefinedName {
        name: "ClockDate".into(),
        formula: "TODAY()".into(),
        local_sheet: None,
        hidden: false,
    });
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet2.xml")
        .unwrap()
        .1 = br#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f ca="1">TODAY()</f><v>45000</v></c><c r="B1"><f ca="1">NOW()</f><v>45000.75</v></c><c r="C1"><f ca="1">TODAY()+1</f><v>45001</v></c><c r="D1"><f ca="1">IFERROR(TODAY(),0)</f><v>45000</v></c><c r="E1"><f ca="1">ClockDate</f><v>45000</v></c><c r="F1"><f>A1+1</f><v>45001</v></c><c r="G1"><f ca="1">TODAY()</f></c><c r="H1" t="str"><f ca="1">TEXT(NOW(),"yyyy")</f><v>2023</v></c><c r="I1"><f ca="1">DATEVALUE("3/15")</f><v>45000</v></c></row><row r="3"><c r="A3"><f t="array" ref="A3:B3" ca="1">SEQUENCE(1,2)+TODAY()</f><v>45001</v></c><c r="B3"><v>45002</v></c></row><row r="5"><c r="A5"><f t="array" ref="A5:B5" ca="1">SEQUENCE(1,2)+TODAY()</f></c><c r="B5"><v>45002</v></c></row></sheetData></worksheet>"#.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn value(workbook: &Workbook, address: &str) -> CellValue {
    workbook
        .sheet(SheetId(1))
        .unwrap()
        .cell(cell(address))
        .map(|cell| cell.value.clone())
        .unwrap_or_default()
}

fn assert_source_caches(workbook: &Workbook) {
    for (address, expected) in [
        ("A1", number(45_000.0)),
        ("B1", number(45_000.75)),
        ("C1", number(45_001.0)),
        ("D1", number(45_000.0)),
        ("E1", number(45_000.0)),
        ("F1", number(45_001.0)),
        ("G1", CellValue::Empty),
        (
            "H1",
            CellValue::Text {
                value: "2023".into(),
            },
        ),
        ("I1", number(45_000.0)),
        ("A3", number(45_001.0)),
        ("B3", number(45_002.0)),
        ("A5", CellValue::Empty),
        ("B5", number(45_002.0)),
    ] {
        assert_eq!(value(workbook, address), expected, "{address}");
    }
    let sheet = workbook.sheet(SheetId(1)).unwrap();
    for (anchor, range) in [("A3", "A3:B3"), ("A5", "A5:B5")] {
        assert_eq!(
            sheet.array_formula(cell(anchor)),
            Some(CellRange::parse_a1(range).unwrap())
        );
    }
}

fn part<'a>(parts: &'a [(String, Vec<u8>)], name: &str) -> &'a [u8] {
    &parts.iter().find(|(path, _)| path == name).unwrap().1
}

fn assert_preserved_save(workbook: &Workbook, source: &[u8]) {
    assert_source_caches(workbook);
    let before = ooxml_opc::unzip_parts(source).unwrap();
    let saved = workbook.save().unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet = part(&after, "xl/worksheets/sheet2.xml");
    assert_eq!(sheet, part(&before, "xl/worksheets/sheet2.xml"));
    assert!(!std::str::from_utf8(sheet).unwrap().contains("t=\"e\""));
    assert!(
        std::str::from_utf8(part(&after, "xl/workbook.xml"))
            .unwrap()
            .contains("fullCalcOnLoad=\"1\"")
    );
    assert_source_caches(&Workbook::open(&saved).unwrap());
}

#[test]
fn clockless_edit_preserves_caches_and_sheet_bytes() {
    let bytes = fixture();
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert_source_caches(&workbook);
    workbook
        .edit_cell(SheetId(0), cell("A1"), "2", CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("A1"))
            .unwrap()
            .value,
        number(2.0)
    );
    assert_preserved_save(&workbook, &bytes);
}

#[test]
fn clockless_batch_preserves_caches_and_sheet_bytes() {
    let bytes = fixture();
    for calculation in [None, Some(json!({}))] {
        let mut workbook = Workbook::open(&bytes).unwrap();
        let mut request = json!({
            "expectVersion": workbook.version(),
            "steps": [{
                "op": "setCellInputs",
                "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "A1" } },
                "inputs": [["2"]]
            }]
        });
        if let Some(calculation) = calculation {
            request["calculation"] = calculation;
        }
        let result: serde_json::Value =
            serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap())
                .unwrap();
        assert_eq!(result["ok"], true);
        assert_eq!(result["applied"], true);
        assert_eq!(
            workbook
                .sheet(SheetId(0))
                .unwrap()
                .cell(cell("A1"))
                .unwrap()
                .value,
            number(2.0)
        );
        assert_preserved_save(&workbook, &bytes);
    }
}

#[test]
fn clockless_full_recalculation_preserves_caches_and_arrays() {
    let bytes = fixture();
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert!(
        workbook
            .recalculate_all(CalculationOptions::default())
            .changed
            .is_empty()
    );
    assert_preserved_save(&workbook, &bytes);
}

#[test]
fn explicit_clock_updates_caches_after_clockless_recalculation() {
    let bytes = fixture();
    for batch in [false, true] {
        let mut workbook = Workbook::open(&bytes).unwrap();
        workbook
            .edit_cell(SheetId(0), cell("A1"), "2", CalculationOptions::default())
            .unwrap();
        assert_source_caches(&workbook);
        if batch {
            let request = json!({
                "expectVersion": workbook.version(),
                "calculation": { "nowSerial": 46_000.25 },
                "steps": [{
                    "op": "setCellInputs",
                    "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": "A1" } },
                    "inputs": [["3"]]
                }]
            });
            let result: serde_json::Value =
                serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap())
                    .unwrap();
            assert_eq!(result["ok"], true);
            assert_eq!(result["applied"], true);
        } else {
            workbook
                .edit_cell(
                    SheetId(0),
                    cell("A1"),
                    "3",
                    CalculationOptions {
                        now_serial: Some(46_000.25),
                    },
                )
                .unwrap();
        }
        for (address, expected) in [
            ("A1", number(46_000.0)),
            ("B1", number(46_000.25)),
            ("C1", number(46_001.0)),
            ("D1", number(46_000.0)),
            ("E1", number(46_000.0)),
            ("F1", number(46_001.0)),
            ("G1", number(46_000.0)),
            (
                "H1",
                CellValue::Text {
                    value: "2025".into(),
                },
            ),
            ("A3", number(46_001.0)),
            ("B3", number(46_002.0)),
            ("A5", number(46_001.0)),
            ("B5", number(46_002.0)),
        ] {
            assert_eq!(value(&workbook, address), expected, "{address}");
        }
        let saved = workbook.save().unwrap();
        let reopened = Workbook::open(&saved).unwrap();
        for address in [
            "A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "A3", "B3", "A5", "B5",
        ] {
            assert_eq!(
                value(&reopened, address),
                value(&workbook, address),
                "{address}"
            );
        }
    }
}

#[test]
fn invalid_clock_calls_commit_value_errors() {
    let mut parts = ooxml_opc::unzip_parts(&fixture()).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet2.xml")
        .unwrap();
    sheet.1 = String::from_utf8(std::mem::take(&mut sheet.1))
        .unwrap()
        .replacen(r#"<f ca="1">TODAY()</f>"#, r#"<f ca="1">TODAY(1)</f>"#, 1)
        .replacen(r#"<f ca="1">NOW()</f>"#, r#"<f ca="1">NOW(1)</f>"#, 1)
        .into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert_eq!(value(&workbook, "A1"), number(45_000.0));
    assert_eq!(value(&workbook, "B1"), number(45_000.75));
    workbook
        .edit_cell(SheetId(0), cell("A1"), "2", CalculationOptions::default())
        .unwrap();
    for address in ["A1", "B1"] {
        assert_eq!(
            value(&workbook, address),
            CellValue::Error {
                value: ErrorValue::Value,
            }
        );
    }
    for (address, input) in [("A7", "=TODAY(1)"), ("B7", "=NOW(1)")] {
        workbook
            .edit_cell(
                SheetId(1),
                cell(address),
                input,
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(
            value(&workbook, address),
            CellValue::Error {
                value: ErrorValue::Value,
            }
        );
    }
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    for address in ["A1", "B1", "A7", "B7"] {
        assert_eq!(
            value(&reopened, address),
            CellValue::Error {
                value: ErrorValue::Value,
            }
        );
    }
}
