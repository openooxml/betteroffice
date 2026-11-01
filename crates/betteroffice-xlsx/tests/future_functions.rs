use betteroffice_xlsx::{
    CalculationOptions, Cell, CellInput, CellRef, CellValue, Op, Sheet, SheetId, Workbook,
    WorkbookModel,
};

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn source() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut data = Sheet::new("Data");
    for (row, (key, value)) in [(1.0, 10.0), (2.0, 20.0), (3.0, 30.0)]
        .into_iter()
        .enumerate()
    {
        for (col, value) in [key, value].into_iter().enumerate() {
            data.set_cell(
                CellRef::new(row as u32, col as u32),
                Cell {
                    value: number(value),
                    ..Cell::default()
                },
            );
        }
    }
    let mut lookalike = Sheet::new("XLOOKUP");
    for (address, value) in [("A1", 5.0), ("A2", 6.0)] {
        lookalike.set_cell(
            cell(address),
            Cell {
                value: number(value),
                ..Cell::default()
            },
        );
    }
    model.sheets.push(data);
    model.sheets.push(lookalike);
    ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap()
}

fn part(saved: &[u8], path: &str) -> String {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let (_, bytes) = parts.iter().find(|(name, _)| name == path).unwrap();
    String::from_utf8(bytes.clone()).unwrap()
}

fn unescape(xml: &str) -> String {
    xml.replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&gt;", ">")
        .replace("&lt;", "<")
        .replace("&amp;", "&")
}

fn element_text(xml: &str, open: &str, close: &str) -> String {
    let start = xml
        .find(open)
        .unwrap_or_else(|| panic!("no {open} in {xml}"));
    let body = start + xml[start..].find('>').unwrap() + 1;
    let end = body + xml[body..].find(close).unwrap();
    unescape(&xml[body..end])
}

fn formula_text(sheet: &str, address: &str) -> String {
    let start = sheet
        .find(&format!("<c r=\"{address}\""))
        .unwrap_or_else(|| panic!("no cell {address}"));
    let end = start + sheet[start..].find("</c>").unwrap();
    element_text(&sheet[start..end], "<f", "</f>")
}

fn assert_no_double_prefix(saved: &[u8]) {
    for (path, bytes) in ooxml_opc::unzip_parts(saved).unwrap() {
        let xml = String::from_utf8_lossy(&bytes).to_ascii_lowercase();
        for doubled in [
            "_xlfn._xlfn.",
            "_xlws._xlws.",
            "_xlpm._xlpm.",
            "_xlpm._xlfn.",
        ] {
            assert!(!xml.contains(doubled), "{doubled} in {path}");
        }
    }
}

fn expression_rule(range: &str, formula: &str) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: betteroffice_xlsx::CellRange::parse_a1(range).unwrap(),
        kind: "expression".into(),
        operator: None,
        formulas: vec![formula.into()],
        format: betteroffice_xlsx::DifferentialFormat::default(),
    }
}

fn defined_name(name: &str, refers_to: &str) -> Op {
    Op::SetDefinedName {
        name: name.into(),
        refers_to: refers_to.into(),
        scope: None,
        hidden: None,
    }
}

/// each typed formula, the text Excel writes for it, and the value it holds.
const TYPED: &[(&str, &str, &str)] = &[
    (
        "D1",
        "=XLOOKUP(2,A1:A3,B1:B3)",
        "_xlfn.XLOOKUP(2,A1:A3,B1:B3)",
    ),
    (
        "D2",
        "=IFS(A1>1,\"big\",TRUE,\"small\")",
        "_xlfn.IFS(A1>1,\"big\",TRUE,\"small\")",
    ),
    ("D3", "=CONCAT(A1:A3)", "_xlfn.CONCAT(A1:A3)"),
    (
        "D4",
        "=TEXTJOIN(\"-\",TRUE,A1:A3)",
        "_xlfn.TEXTJOIN(\"-\",TRUE,A1:A3)",
    ),
    (
        "D5",
        "=SUM(XLOOKUP(3,A1:A3,B1:B3))",
        "SUM(_xlfn.XLOOKUP(3,A1:A3,B1:B3))",
    ),
    (
        "D6",
        "=IFERROR(XLOOKUP(9,A1:A3,B1:B3),0)",
        "IFERROR(_xlfn.XLOOKUP(9,A1:A3,B1:B3),0)",
    ),
    (
        "D7",
        "=LET(x,2,y,3,x*y)",
        "_xlfn.LET(_xlpm.x,2,_xlpm.y,3,_xlpm.x*_xlpm.y)",
    ),
    (
        "D8",
        "=REDUCE(0,A1:A3,LAMBDA(acc,v,acc+v))",
        "_xlfn.REDUCE(0,A1:A3,_xlfn.LAMBDA(_xlpm.acc,_xlpm.v,_xlpm.acc+_xlpm.v))",
    ),
    ("D9", "=\"XLOOKUP(\"&A1", "\"XLOOKUP(\"&A1"),
    ("D10", "=SUM(XLOOKUP!A1:A2)", "SUM(XLOOKUP!A1:A2)"),
    (
        "D11",
        "=xlookup(2,A1:A3,B1:B3)",
        "_xlfn.XLOOKUP(2,A1:A3,B1:B3)",
    ),
    ("D12", "=Pick*2", "Pick*2"),
    ("F1", "=SEQUENCE(3)", "_xlfn.SEQUENCE(3)"),
    ("G1", "=SORT(A1:A3,1,-1)", "_xlfn._xlws.SORT(A1:A3,1,-1)"),
    ("H1", "=UNIQUE(A1:A3)", "_xlfn.UNIQUE(A1:A3)"),
    (
        "I1",
        "=FILTER(A1:A3,A1:A3>1)",
        "_xlfn._xlws.FILTER(A1:A3,A1:A3>1)",
    ),
];

#[test]
fn typed_future_functions_save_with_the_prefixes_excel_writes() {
    let mut workbook = Workbook::open(&source()).unwrap();
    workbook
        .apply_ops(
            vec![
                defined_name("Pick", "XLOOKUP(3,Data!$A$1:$A$3,Data!$B$1:$B$3)"),
                expression_rule("A1:A3", "XLOOKUP(A1,$A$1:$A$3,$B$1:$B$3)>10"),
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    let edits = TYPED
        .iter()
        .map(|(address, input, _)| CellInput {
            cell: cell(address),
            input: (*input).into(),
        })
        .collect::<Vec<_>>();
    workbook
        .edit_cells(SheetId(0), &edits, CalculationOptions::default())
        .unwrap();
    let saved = workbook.save().unwrap();

    let sheet = part(&saved, "xl/worksheets/sheet1.xml");
    for (address, _, expected) in TYPED {
        assert_eq!(formula_text(&sheet, address), *expected, "{address}");
    }
    assert_eq!(
        element_text(&sheet, "<formula", "</formula>"),
        "_xlfn.XLOOKUP(A1,$A$1:$A$3,$B$1:$B$3)>10"
    );
    assert_eq!(
        element_text(
            &part(&saved, "xl/workbook.xml"),
            "<definedName name=\"Pick\"",
            "</definedName>"
        ),
        "_xlfn.XLOOKUP(3,Data!$A$1:$A$3,Data!$B$1:$B$3)"
    );
    assert_no_double_prefix(&saved);

    let reopened = Workbook::open(&saved).unwrap();
    let data = reopened.sheet(SheetId(0)).unwrap();
    for (address, expected) in [
        ("D1", number(20.0)),
        ("D2", text("small")),
        ("D3", text("123")),
        ("D4", text("1-2-3")),
        ("D5", number(30.0)),
        ("D6", number(0.0)),
        ("D7", number(6.0)),
        ("D8", number(6.0)),
        ("D9", text("XLOOKUP(1")),
        ("D10", number(11.0)),
        ("D11", number(20.0)),
        ("D12", number(60.0)),
        ("F1", number(1.0)),
        ("G1", number(3.0)),
        ("H1", number(1.0)),
        ("I1", number(2.0)),
    ] {
        assert_eq!(
            data.cell(cell(address)).unwrap().value,
            expected,
            "{address}"
        );
    }
    assert_eq!(reopened.save().unwrap(), saved);
}

fn prefixed_source() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&source()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "xl/worksheets/sheet1.xml" {
            *bytes = concat!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
                r#"<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>10</v></c><c r="D1"><f>_xlfn.XLOOKUP(2,A1:A3,B1:B3)</f><v>20</v></c><c r="E1"><f>_xlfn.LET(_xlpm.x,2,_xlpm.y,3,_xlpm.x*_xlpm.y)</f><v>6</v></c></row>"#,
                r#"<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>20</v></c><c r="D2"><f>IFERROR(_xlfn.XLOOKUP(9,A1:A3,B1:B3),0)</f><v>0</v></c></row>"#,
                r#"<row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>30</v></c></row>"#,
                r#"</sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>_xlfn.XLOOKUP(A1,$A$1:$A$3,$B$1:$B$3)&gt;10</formula></cfRule></conditionalFormatting>"#,
                r#"<dataValidations count="1"><dataValidation type="custom" allowBlank="1" sqref="C1:C3"><formula1>_xlfn.ISFORMULA(C1)</formula1></dataValidation></dataValidations></worksheet>"#,
            )
            .as_bytes()
            .to_vec();
        } else if path == "xl/workbook.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "</workbook>",
                    r#"<definedNames><definedName name="Pick">_xlfn.XLOOKUP(2,Data!$A$1:$A$3,Data!$B$1:$B$3)</definedName></definedNames></workbook>"#,
                )
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn edited_save_of_prefixed_formulas_keeps_their_text() {
    let source = prefixed_source();
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(workbook.save().unwrap(), source);
    workbook
        .apply_ops(
            vec![
                defined_name("Other", "SEQUENCE(2)"),
                expression_rule("B1:B3", "_xlfn.XLOOKUP(B1,$B$1:$B$3,$A$1:$A$3)>1"),
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .edit_cells(
            SheetId(0),
            &[
                CellInput {
                    cell: cell("C1"),
                    input: "=_xlfn.XLOOKUP(3,A1:A3,B1:B3)".into(),
                },
                CellInput {
                    cell: cell("C2"),
                    input: "=_xlfn.LET(_xlpm.n,A2,_xlpm.n*2)".into(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();

    let sheet = part(&saved, "xl/worksheets/sheet1.xml");
    for (address, expected) in [
        ("D1", "_xlfn.XLOOKUP(2,A1:A3,B1:B3)"),
        ("E1", "_xlfn.LET(_xlpm.x,2,_xlpm.y,3,_xlpm.x*_xlpm.y)"),
        ("D2", "IFERROR(_xlfn.XLOOKUP(9,A1:A3,B1:B3),0)"),
        ("C1", "_xlfn.XLOOKUP(3,A1:A3,B1:B3)"),
        ("C2", "_xlfn.LET(_xlpm.n,A2,_xlpm.n*2)"),
    ] {
        assert_eq!(formula_text(&sheet, address), expected, "{address}");
    }
    assert!(sheet.contains("<formula>_xlfn.XLOOKUP(A1,$A$1:$A$3,$B$1:$B$3)&gt;10</formula>"));
    assert!(sheet.contains("<formula>_xlfn.XLOOKUP(B1,$B$1:$B$3,$A$1:$A$3)&gt;1</formula>"));
    assert!(sheet.contains("<formula1>_xlfn.ISFORMULA(C1)</formula1>"));
    let workbook_xml = part(&saved, "xl/workbook.xml");
    assert_eq!(
        element_text(
            &workbook_xml,
            "<definedName name=\"Pick\"",
            "</definedName>"
        ),
        "_xlfn.XLOOKUP(2,Data!$A$1:$A$3,Data!$B$1:$B$3)"
    );
    assert_eq!(
        element_text(
            &workbook_xml,
            "<definedName name=\"Other\"",
            "</definedName>"
        ),
        "_xlfn.SEQUENCE(2)"
    );
    assert_no_double_prefix(&saved);

    let reopened = Workbook::open(&saved).unwrap();
    let data = reopened.sheet(SheetId(0)).unwrap();
    for (address, expected) in [
        ("C1", number(30.0)),
        ("C2", number(4.0)),
        ("D1", number(20.0)),
        ("E1", number(6.0)),
    ] {
        assert_eq!(
            data.cell(cell(address)).unwrap().value,
            expected,
            "{address}"
        );
    }
}

#[test]
fn a_hidden_flag_edit_of_a_deeply_nested_defined_name_saves() {
    for (depth, prefixed) in [(1_489, true), (3_276, false)] {
        let formula = format!("{}a{}", "LAMBDA(a,".repeat(depth), ")".repeat(depth));
        assert!(formula.len() <= 32_768);
        let mut parts = ooxml_opc::unzip_parts(&source()).unwrap();
        for (path, bytes) in &mut parts {
            if path == "xl/workbook.xml" {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace(
                        "</workbook>",
                        &format!(
                            r#"<definedNames><definedName name="Deep">{formula}</definedName></definedNames></workbook>"#
                        ),
                    )
                    .into_bytes();
            }
        }
        let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
        workbook
            .apply_ops(
                vec![Op::SetDefinedName {
                    name: "Deep".into(),
                    refers_to: formula.clone(),
                    scope: None,
                    hidden: Some(true),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        let saved = workbook.save().unwrap();
        let expected = if prefixed {
            format!(
                "{}_xlpm.a{}",
                "_xlfn.LAMBDA(_xlpm.a,".repeat(depth),
                ")".repeat(depth)
            )
        } else {
            formula
        };
        assert_eq!(
            element_text(
                &part(&saved, "xl/workbook.xml"),
                "<definedName name=\"Deep\"",
                "</definedName>"
            ),
            expected
        );
        Workbook::open(&saved).unwrap();
    }
}
