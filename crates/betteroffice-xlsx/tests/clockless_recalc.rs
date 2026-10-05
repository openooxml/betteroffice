use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook,
    WorkbookModel,
};
use serde_json::json;
use xlsx_model::{DefinedName, ErrorValue};
use xlsx_ops::Op;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn synthetic(
    cells: &[(&str, &str, CellValue)],
    arrays: &[(&str, &str)],
    names: Vec<DefinedName>,
) -> Workbook {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Sheet2");
    for (address, formula, cached) in cells {
        sheet.set_cell(
            cell(address),
            Cell {
                value: cached.clone(),
                formula: Some((*formula).into()),
                style: None,
            },
        );
    }
    for (anchor, range) in arrays {
        sheet.set_array_formula(cell(anchor), CellRange::parse_a1(range).unwrap());
    }
    model.sheets.extend([Sheet::new("Sheet1"), sheet]);
    model.defined_names = names;
    let parts = xlsx_parse::serialize_workbook(&model).unwrap();
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
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
        .1 = br#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f ca="1">TODAY()</f><v>45000</v></c><c r="B1"><f ca="1">NOW()</f><v>45000.75</v></c><c r="C1"><f ca="1">TODAY()+1</f><v>45001</v></c><c r="D1"><f ca="1">IFERROR(TODAY(),0)</f><v>45000</v></c><c r="E1"><f ca="1">ClockDate</f><v>45000</v></c><c r="F1"><f>A1+1</f><v>45001</v></c><c r="G1"><f ca="1">TODAY()</f></c><c r="H1" t="str"><f ca="1">TEXT(NOW(),"yyyy")</f><v>2023</v></c><c r="I1"><f ca="1">DATEVALUE("3/15")</f><v>45000</v></c><c r="J1"><f>Sheet1!A1+1</f><v>2</v></c></row><row r="3"><c r="A3"><f t="array" ref="A3:B3" ca="1">SEQUENCE(1,2)+TODAY()</f><v>45001</v></c><c r="B3"><v>45002</v></c></row><row r="5"><c r="A5"><f t="array" ref="A5:B5" ca="1">SEQUENCE(1,2)+TODAY()</f></c><c r="B5"><v>45002</v></c></row></sheetData></worksheet>"#.to_vec();
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
    assert_source_caches_with_dependent(workbook, number(45_001.0));
}

fn assert_source_caches_with_dependent(workbook: &Workbook, dependent: CellValue) {
    for (address, expected) in [
        ("A1", number(45_000.0)),
        ("B1", number(45_000.75)),
        ("C1", number(45_001.0)),
        ("D1", number(45_000.0)),
        ("E1", number(45_000.0)),
        ("F1", dependent),
        ("G1", CellValue::Empty),
        (
            "H1",
            CellValue::Text {
                value: "2023".into(),
            },
        ),
        ("I1", number(45_000.0)),
        ("J1", number(2.0)),
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

fn batch_edit(workbook: &mut Workbook, address: &str, input: &str, now: Option<f64>) {
    batch_edit_on(workbook, "sheet:0", address, input, now);
}

fn batch_edit_on(
    workbook: &mut Workbook,
    sheet: &str,
    address: &str,
    input: &str,
    now: Option<f64>,
) {
    let request = json!({
        "expectVersion": workbook.version(),
        "steps": [{
            "op": "setCellInputs",
            "target": { "sheetId": sheet, "range": { "kind": "a1", "a1": address } },
            "inputs": [[input]]
        }],
        "calculation": { "nowSerial": now }
    });
    let result: serde_json::Value =
        serde_json::from_str(&workbook.apply_edits_json(&request.to_string()).unwrap()).unwrap();
    assert_eq!(result["ok"], true);
    assert_eq!(result["applied"], true);
}

#[test]
fn clockless_edit_preserves_caches_and_sheet_bytes() {
    let bytes = fixture();
    let mut workbook = Workbook::open(&bytes).unwrap();
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
    let before = ooxml_opc::unzip_parts(&bytes).unwrap();
    let after = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
    let original = std::str::from_utf8(part(&before, "xl/worksheets/sheet1.xml")).unwrap();
    assert_eq!(
        part(&after, "xl/worksheets/sheet1.xml"),
        original.replace("<v>1</v>", "<v>2</v>").as_bytes()
    );
    for (name, data) in &before {
        if name != "xl/workbook.xml" && name != "xl/worksheets/sheet1.xml" {
            assert_eq!(part(&after, name), data, "{name}");
        }
    }
}

#[test]
fn clockless_batch_preserves_caches_and_sheet_bytes() {
    let bytes = fixture();
    let mut workbook = Workbook::open(&bytes).unwrap();
    batch_edit(&mut workbook, "A1", "2", None);
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
fn clockless_full_recalculation_preserves_caches_and_arrays() {
    let bytes = fixture();
    let mut workbook = Workbook::open(&bytes).unwrap();
    let result = workbook.recalculate_all(CalculationOptions::default());
    assert!(result.changed.is_empty());
    assert!(result.cycle_cells.is_empty());
    assert!(result.limited_cells.is_empty());
    assert_preserved_save(&workbook, &bytes);
}

#[test]
fn clockless_plain_formula_recalculates_without_clock_tokens() {
    for batch in [false, true] {
        let mut workbook = synthetic(
            &[
                ("A1", "1+1", number(45_000.0)),
                ("B1", "Sheet1!A1+1", number(99.0)),
            ],
            &[],
            Vec::new(),
        );
        if batch {
            batch_edit(&mut workbook, "A1", "7", None);
        } else {
            workbook
                .edit_cell(SheetId(0), cell("A1"), "7", CalculationOptions::default())
                .unwrap();
        }
        assert_eq!(value(&workbook, "B1"), number(8.0));
        assert_eq!(
            value(&Workbook::open(&workbook.save().unwrap()).unwrap(), "B1"),
            number(8.0)
        );
    }
}

#[test]
fn clocked_edit_catches_up_pending_across_edit_paths() {
    for first_batch in [false, true] {
        for second_batch in [false, true] {
            let mut workbook = Workbook::open(&fixture()).unwrap();
            if first_batch {
                batch_edit(&mut workbook, "A1", "7", None);
            } else {
                workbook
                    .edit_cell(SheetId(0), cell("A1"), "7", CalculationOptions::default())
                    .unwrap();
            }
            assert_source_caches(&workbook);
            if second_batch {
                batch_edit(&mut workbook, "B1", "9", Some(46_000.25));
            } else {
                workbook
                    .edit_cell(
                        SheetId(0),
                        cell("B1"),
                        "9",
                        CalculationOptions {
                            now_serial: Some(46_000.25),
                        },
                    )
                    .unwrap();
            }
            assert_eq!(value(&workbook, "J1"), number(8.0));
            assert_eq!(value(&workbook, "A1"), number(46_000.0));
            assert_eq!(value(&workbook, "F1"), number(46_001.0));
            assert_eq!(value(&workbook, "B3"), number(46_002.0));
            workbook
                .edit_cell(
                    SheetId(1),
                    cell("J1"),
                    "=99",
                    CalculationOptions {
                        now_serial: Some(46_000.25),
                    },
                )
                .unwrap();
            assert_eq!(value(&workbook, "J1"), number(99.0));
        }
    }
}

#[test]
fn clocked_full_recalculation_catches_up_pending() {
    let mut workbook = Workbook::open(&fixture()).unwrap();
    batch_edit(&mut workbook, "A1", "7", None);
    workbook.recalculate_all(CalculationOptions {
        now_serial: Some(46_000.25),
    });
    assert_eq!(value(&workbook, "J1"), number(8.0));
    assert_eq!(value(&workbook, "F1"), number(46_001.0));
}

#[test]
fn defined_name_tokens_gate_even_unused_and_sheet_scoped() {
    for local_sheet in [None, Some(SheetId(1))] {
        let mut workbook = synthetic(
            &[("A1", "Sheet1!A1+1", number(99.0))],
            &[],
            vec![DefinedName {
                name: "UnusedClock".into(),
                formula: "=NOW()".into(),
                local_sheet,
                hidden: false,
            }],
        );
        batch_edit(&mut workbook, "A1", "7", None);
        assert_eq!(value(&workbook, "A1"), number(99.0));
        workbook
            .apply_ops(
                vec![Op::SetDefinedNames {
                    defined_names: Vec::new(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(value(&workbook, "A1"), number(8.0));
    }
}

#[test]
fn defined_name_replacement_refreshes_clock_presence() {
    let mut workbook = synthetic(&[("A1", "Sheet1!A1+1", number(99.0))], &[], Vec::new());
    workbook
        .apply_ops(
            vec![Op::SetDefinedNames {
                defined_names: vec![DefinedName {
                    name: "Clock".into(),
                    formula: "TODAY()".into(),
                    local_sheet: None,
                    hidden: false,
                }],
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(value(&workbook, "A1"), number(99.0));
    batch_edit(&mut workbook, "A1", "7", None);
    workbook
        .apply_ops(
            vec![Op::SetDefinedNames {
                defined_names: vec![DefinedName {
                    name: "Clock".into(),
                    formula: "1".into(),
                    local_sheet: None,
                    hidden: false,
                }],
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(value(&workbook, "A1"), number(8.0));
}

#[test]
fn shared_formula_tokens_gate_the_workbook() {
    let mut parts =
        ooxml_opc::unzip_parts(&synthetic(&[], &[], Vec::new()).save().unwrap()).unwrap();
    parts.iter_mut().find(|(name, _)| name == "xl/worksheets/sheet2.xml").unwrap().1 =
        br#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:B1">TODAY()</f><v>45000</v></c><c r="B1"><f t="shared" si="0"/><v>45000</v></c><c r="C1"><f>Sheet1!A1+1</f><v>99</v></c></row></sheetData></worksheet>"#.to_vec();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&bytes).unwrap();
    batch_edit(&mut workbook, "A1", "7", None);
    assert_eq!(value(&workbook, "A1"), number(45_000.0));
    assert_eq!(value(&workbook, "B1"), number(45_000.0));
    assert_eq!(value(&workbook, "C1"), number(99.0));
    let saved = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        part(&saved, "xl/worksheets/sheet2.xml"),
        part(&parts, "xl/worksheets/sheet2.xml")
    );
    workbook
        .edit_cell(SheetId(1), cell("A1"), "0", CalculationOptions::default())
        .unwrap();
    assert_eq!(value(&workbook, "C1"), number(99.0));
    workbook
        .edit_cell(SheetId(1), cell("B1"), "0", CalculationOptions::default())
        .unwrap();
    assert_eq!(value(&workbook, "C1"), number(8.0));
}

#[test]
fn array_formula_tokens_preserve_all_spill_state() {
    let mut workbook = synthetic(
        &[
            ("A1", "SEQUENCE(1,2)+TODAY()", number(45_001.0)),
            ("C1", "Sheet1!A1+1", number(99.0)),
            ("D1", "SEQUENCE(1,2)", number(9.0)),
        ],
        &[("A1", "A1:B1"), ("D1", "D1:E1")],
        Vec::new(),
    );
    let before = workbook.save().unwrap();
    batch_edit(&mut workbook, "A1", "7", None);
    assert_eq!(value(&workbook, "A1"), number(45_001.0));
    assert_eq!(value(&workbook, "B1"), CellValue::Empty);
    assert_eq!(value(&workbook, "C1"), number(99.0));
    assert_eq!(value(&workbook, "D1"), number(9.0));
    assert_eq!(value(&workbook, "E1"), CellValue::Empty);
    let sheet = workbook.sheet(SheetId(1)).unwrap();
    assert_eq!(
        sheet.array_formula(cell("A1")),
        Some(CellRange::parse_a1("A1:B1").unwrap())
    );
    assert_eq!(
        sheet.array_formula(cell("D1")),
        Some(CellRange::parse_a1("D1:E1").unwrap())
    );
    let before = ooxml_opc::unzip_parts(&before).unwrap();
    let after = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        part(&after, "xl/worksheets/sheet2.xml"),
        part(&before, "xl/worksheets/sheet2.xml")
    );
}

#[test]
fn removing_last_formula_token_runs_pending_full_recalculation() {
    for batch in [false, true] {
        let mut workbook = synthetic(
            &[
                ("A1", "TODAY()", number(45_000.0)),
                ("B1", "Sheet1!A1+1", number(99.0)),
            ],
            &[],
            Vec::new(),
        );
        if batch {
            batch_edit(&mut workbook, "A1", "7", None);
        } else {
            workbook
                .edit_cell(SheetId(0), cell("A1"), "7", CalculationOptions::default())
                .unwrap();
        }
        assert_eq!(value(&workbook, "B1"), number(99.0));
        if batch {
            batch_edit_on(&mut workbook, "sheet:1", "A1", "=1+1", None);
        } else {
            workbook
                .edit_cell(
                    SheetId(1),
                    cell("A1"),
                    "=1+1",
                    CalculationOptions::default(),
                )
                .unwrap();
        }
        assert_eq!(value(&workbook, "A1"), number(2.0));
        assert_eq!(value(&workbook, "B1"), number(8.0));
    }
}

#[test]
fn clock_token_calls_gate_regardless_of_arity_or_date_text() {
    for formula in [
        "TODAY(1)",
        "NOW(1)",
        "IF(TRUE,1,TODAY(1))",
        "DATEVALUE(\"3/15\")",
        "DATEVALUE(\"3/15/2023\")",
        "DATEVALUE(\"invalid\")",
        "_xlfn.NOW()",
        "now ()",
        "_XLFN._XLWS.ToDaY()",
        "_xlfn._xlws.DATEVALUE(\"3/15/2023\")",
        "datevalue ()",
    ] {
        let mut workbook = synthetic(
            &[
                ("A1", formula, number(45_000.0)),
                ("B1", "Sheet1!A1+1", number(99.0)),
            ],
            &[],
            Vec::new(),
        );
        batch_edit(&mut workbook, "A1", "7", None);
        assert_eq!(value(&workbook, "A1"), number(45_000.0), "{formula}");
        assert_eq!(value(&workbook, "B1"), number(99.0), "{formula}");
    }
}

#[test]
fn clock_token_text_and_bare_names_do_not_gate() {
    for formula in ["\"TODAY()\"", "ClocklessName"] {
        let mut workbook = synthetic(
            &[
                ("A1", formula, number(99.0)),
                ("B1", "Sheet1!A1+1", number(99.0)),
            ],
            &[],
            Vec::new(),
        );
        batch_edit(&mut workbook, "A1", "7", None);
        assert_eq!(value(&workbook, "B1"), number(8.0));
    }
}

#[test]
fn direct_today_evaluation_without_clock_still_returns_value_error() {
    let model = WorkbookModel::default();
    assert_eq!(
        xlsx_calc::evaluate(
            &xlsx_calc::parse_formula("TODAY()").unwrap(),
            &xlsx_calc::EvalContext::new(&model, SheetId(0))
        ),
        CellValue::Error {
            value: ErrorValue::Value
        }
    );
}

#[test]
fn formula_replacement_updates_clock_presence() {
    let mut workbook = synthetic(
        &[
            ("A1", "1+1", number(45_000.0)),
            ("B1", "Sheet1!A1+1", number(99.0)),
        ],
        &[],
        Vec::new(),
    );
    workbook
        .edit_cell(
            SheetId(1),
            cell("A1"),
            "=TODAY()",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        workbook
            .sheet(SheetId(1))
            .unwrap()
            .cell(cell("A1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("TODAY()")
    );
    assert_eq!(value(&workbook, "B1"), number(99.0));
    workbook
        .edit_cell(
            SheetId(1),
            cell("A1"),
            "=NOW()",
            CalculationOptions::default(),
        )
        .unwrap();
    batch_edit(&mut workbook, "A1", "7", None);
    assert_eq!(value(&workbook, "B1"), number(99.0));
    workbook
        .edit_cell(SheetId(1), cell("A1"), "0", CalculationOptions::default())
        .unwrap();
    assert_eq!(value(&workbook, "B1"), number(8.0));
}

#[test]
fn date_coercion_functions_recalculate_without_clock_tokens() {
    let error = CellValue::Error {
        value: ErrorValue::Value,
    };
    let formulas = [
        ("A1", "ROUND(Sheet1!A1,\"0\")", number(7.0)),
        (
            "B1",
            "VLOOKUP(Sheet1!A1,Sheet1!A1:A1,\"1\",FALSE)",
            number(7.0),
        ),
        ("C1", "YEAR(\"3/15/2023\")+Sheet1!A1-7", number(2023.0)),
        (
            "D1",
            "DATE(\"2023\",\"3\",\"15\")+Sheet1!A1-7",
            number(45_000.0),
        ),
        ("E1", "ROUND(\"3/15\",Sheet1!A1)", error.clone()),
        (
            "F1",
            "VLOOKUP(Sheet1!A1,Sheet1!A1:A1,\"3/15\",FALSE)",
            error.clone(),
        ),
        ("G1", "YEAR(\"3/15\")+Sheet1!A1", error.clone()),
        ("H1", "DATE(\"3/15\",3,Sheet1!A1)", error.clone()),
        ("I1", "MONTH(\"3/15\")+Sheet1!A1", error.clone()),
        ("J1", "DAYS(\"3/15\",1)+Sheet1!A1", error.clone()),
        ("K1", "ABS(\"3/15\")+Sheet1!A1", error.clone()),
        ("L1", "INDEX(A2:A3,\"3/15\")+Sheet1!A1", error.clone()),
        ("M1", "OFFSET(A2,\"3/15\",0)+Sheet1!A1", error.clone()),
        ("N1", "RANDBETWEEN(1,\"3/15\")+Sheet1!A1", error),
        ("O1", "Sheet1!A1+1", number(8.0)),
    ];
    let cached = formulas
        .iter()
        .map(|(address, formula, _)| (*address, *formula, number(-1.0)))
        .collect::<Vec<_>>();
    for batch in [false, true] {
        let mut workbook = synthetic(&cached, &[], Vec::new());
        if batch {
            batch_edit(&mut workbook, "A1", "7", None);
        } else {
            workbook
                .edit_cell(SheetId(0), cell("A1"), "7", CalculationOptions::default())
                .unwrap();
        }
        for (address, formula, expected) in &formulas {
            assert_eq!(value(&workbook, address), *expected, "{formula}");
        }
    }
    let mut workbook = synthetic(&cached, &[], Vec::new());
    workbook.recalculate_all(CalculationOptions::default());
    for (address, formula, expected) in &formulas[4..14] {
        assert_eq!(value(&workbook, address), *expected, "{formula}");
    }
    assert_eq!(value(&workbook, "A1"), number(0.0));
    assert_eq!(value(&workbook, "C1"), number(2016.0));
    assert_eq!(value(&workbook, "D1"), number(44_993.0));
    assert_eq!(value(&workbook, "O1"), number(1.0));
}
