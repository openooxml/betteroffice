use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRange, CellRef, CellState, CellValue, EditRequest, Op,
    ProposalEditInput, ProposalRequest, SheetId, Workbook,
};
use serde_json::{Value, json};
use xlsx_model::ErrorValue;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

fn package(sheet_data: &str, calc_pr: &str) -> Vec<u8> {
    package_on_sheet("Sheet1", sheet_data, calc_pr)
}

fn package_on_sheet(sheet_name: &str, sheet_data: &str, calc_pr: &str) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{PACKAGE_RELS}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>{calc_pr}</workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(r#"<Relationships xmlns="{PACKAGE_RELS}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RELS}/styles" Target="styles.xml"/></Relationships>"#),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData>{sheet_data}</sheetData></worksheet>"#),
        ),
        ("xl/styles.xml", STYLES.to_owned()),
    ];
    let parts: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), [DECLARATION, &xml].concat().into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

fn stored(workbook: &Workbook, address: &str) -> (Option<String>, CellValue) {
    let stored = workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell(address))
        .cloned()
        .unwrap_or_default();
    (stored.formula, stored.value)
}

fn shown(workbook: &Workbook, address: &str) -> String {
    workbook.cell(SheetId(0), cell(address)).unwrap().input
}

fn recommit_all(workbook: &mut Workbook, addresses: &[&str]) {
    let inputs: Vec<CellInput> = addresses
        .iter()
        .map(|address| CellInput {
            cell: cell(address),
            input: shown(workbook, address),
        })
        .collect();
    workbook.edit_cells(SheetId(0), &inputs, options()).unwrap();
}

fn apply(workbook: &mut Workbook, steps: Value) {
    let request: EditRequest =
        serde_json::from_value(json!({ "expectVersion": workbook.version(), "steps": steps }))
            .unwrap();
    workbook
        .apply_edits(&request)
        .unwrap()
        .expect("the batch applies");
}

fn set_formula(address: &str, source: &str) -> Value {
    json!({
        "op": "setFormulas",
        "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": address } },
        "formulas": [[source]],
    })
}

fn set_input(address: &str, input: &str) -> Value {
    json!({
        "op": "setCellInputs",
        "target": { "sheetId": "sheet:0", "range": { "kind": "a1", "a1": address } },
        "inputs": [[input]],
    })
}

const EXTERNAL: &[(&str, &str, f64)] = &[
    ("A1", "[1]Sheet1!A1", 5.0),
    ("B1", "SUM([1]Prices!$B$2:$B$4)*2", 12.0),
    ("C1", "[1]!Rate", 0.5),
    ("D1", "[2]Sheet1!A1+'[1]Other sheet'!C3", 7.0),
];

fn external_book(calc_pr: &str) -> Vec<u8> {
    let row: String = EXTERNAL
        .iter()
        .map(|(address, formula, cached)| {
            format!(r#"<c r="{address}"><f>{formula}</f><v>{cached}</v></c>"#)
        })
        .collect();
    package(&format!(r#"<row r="1">{row}</row>"#), calc_pr)
}

#[test]
fn external_references_recommit_as_the_stored_formulas() {
    let source = external_book("");
    let addresses: Vec<&str> = EXTERNAL.iter().map(|(address, ..)| *address).collect();
    let mut workbook = Workbook::open(&source).unwrap();
    let version = workbook.version();
    for (address, formula, _) in EXTERNAL {
        assert_eq!(shown(&workbook, address), format!("={formula}"));
    }
    recommit_all(&mut workbook, &addresses);
    for (address, _, _) in EXTERNAL {
        let input = shown(&workbook, address);
        workbook
            .edit_cell(SheetId(0), cell(address), &input, options())
            .unwrap();
    }
    for (address, formula, cached) in EXTERNAL {
        assert_eq!(
            stored(&workbook, address),
            (Some((*formula).to_owned()), number(*cached))
        );
    }
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), source);
}

#[test]
fn external_references_recommit_unchanged_through_batches() {
    let source = external_book("");
    let mut workbook = Workbook::open(&source).unwrap();
    let version = workbook.version();
    let steps: Vec<Value> = EXTERNAL
        .iter()
        .flat_map(|(address, formula, _)| {
            [
                set_formula(address, formula),
                set_input(address, &format!("={formula}")),
            ]
        })
        .collect();
    for step in steps {
        apply(&mut workbook, json!([step]));
    }
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), source);
}

#[test]
fn batches_recommit_a_deleted_external_name_and_enter_it_without_a_value() {
    let source = package(
        r#"<row r="1"><c r="A1"><f>[1]!#REF!+1</f><v>9</v></c></row>"#,
        "",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    let version = workbook.version();
    apply(&mut workbook, json!([set_formula("A1", "[1]!#REF!+1")]));
    apply(&mut workbook, json!([set_input("A1", "=[1]!#REF!+1")]));
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), source);

    apply(&mut workbook, json!([set_input("B1", "=[1]!#REF!")]));
    workbook.recalculate_all(options());
    assert_eq!(
        stored(&workbook, "B1"),
        (Some("[1]!#REF!".into()), CellValue::Empty)
    );
}

const QUOTED: &str = concat!(
    r#"<row r="1"><c r="A1"><f>[1]Sheet1!A1</f><v>5</v></c>"#,
    r#"<c r="B1"><f>'[1]Other sheet'!C3</f><v>7</v></c>"#,
    r#"<c r="C1"><f>A1+B1</f><v>12</v></c><c r="D1"><v>1</v></c></row>"#,
);

fn propose(workbook: &mut Workbook, edits: &[(&str, &str)]) -> betteroffice_xlsx::Proposal {
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: edits
                    .iter()
                    .map(|(address, input)| ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell(address),
                        input: (*input).to_owned(),
                        number_format: None,
                    })
                    .collect(),
            },
            options(),
        )
        .unwrap()
}

#[test]
fn proposals_keep_external_values_beside_a_real_edit() {
    let mut workbook = Workbook::open(&package(QUOTED, "")).unwrap();
    let proposal = propose(
        &mut workbook,
        &[
            ("A1", "=[1]Sheet1!A1"),
            ("B1", "='[1]Other sheet'!C3"),
            ("D1", "10"),
        ],
    );
    let texts: Vec<_> = proposal
        .edits
        .iter()
        .map(|edit| {
            (
                edit.a1.as_str(),
                edit.old_text.as_str(),
                edit.new_text.as_str(),
            )
        })
        .collect();
    assert_eq!(
        texts,
        [("A1", "5", "5"), ("B1", "7", "7"), ("D1", "1", "10")]
    );
    let ghosts: Vec<_> = proposal
        .ghosts
        .iter()
        .map(|ghost| (ghost.col, ghost.old_text.as_str(), ghost.new_text.as_str()))
        .collect();
    assert_eq!(ghosts, [(0, "5", "5"), (1, "7", "7"), (3, "1", "10")]);
    workbook
        .accept_proposal(&proposal.id, false, options())
        .unwrap();
    for (address, formula, value) in [
        ("A1", Some("[1]Sheet1!A1"), 5.0),
        ("B1", Some("'[1]Other sheet'!C3"), 7.0),
        ("C1", Some("A1+B1"), 12.0),
        ("D1", None, 10.0),
    ] {
        assert_eq!(
            stored(&workbook, address),
            (formula.map(str::to_owned), number(value)),
            "{address}"
        );
    }
}

const QUOTED_GUARDED: &str = concat!(
    r#"<row r="1"><c r="A1"><f>IFERROR('[1]Other sheet'!C3,0)</f><v>42</v></c>"#,
    r#"<c r="B1" t="b"><f>ISREF('[1]Other sheet'!C3)</f><v>0</v></c>"#,
    r#"<c r="C1"><f>'[1]Other sheet'!C3*2</f><v>14</v></c>"#,
    r#"<c r="D1"><f>'[1]Book'!Rate</f><v>0.25</v></c></row>"#,
);

#[test]
fn quoted_external_references_keep_cached_values_and_enter_blank() {
    let source = package(QUOTED_GUARDED, r#"<calcPr fullCalcOnLoad="1"/>"#);
    let mut workbook = Workbook::open_recalculated(&source, options()).unwrap();
    workbook.recalculate_all(options());
    assert_eq!(stored(&workbook, "A1").1, number(42.0));
    assert_eq!(stored(&workbook, "B1").1, CellValue::Bool { value: false });
    assert_eq!(stored(&workbook, "C1").1, number(14.0));
    assert_eq!(stored(&workbook, "D1").1, number(0.25));
    workbook
        .edit_cell(SheetId(0), cell("A2"), "='[1]Other sheet'!C3*2", options())
        .unwrap();
    apply(
        &mut workbook,
        json!([set_formula("B2", "IFERROR('[1]Data'!A1,0)")]),
    );
    workbook.recalculate_all(options());
    assert_eq!(
        stored(&workbook, "A2"),
        (Some("'[1]Other sheet'!C3*2".into()), CellValue::Empty)
    );
    assert_eq!(
        stored(&workbook, "B2"),
        (Some("IFERROR('[1]Data'!A1,0)".into()), CellValue::Empty)
    );
}

const NAMES: &str = concat!(
    r#"<definedNames><definedName name="Rate">'[1]Other sheet'!$C$3</definedName>"#,
    r#"<definedName name="Plain">[1]Sheet1!$C$3</definedName>"#,
    r#"<definedName name="Chain">Rate*2</definedName>"#,
    r#"<definedName name="Loop">Loop+1</definedName></definedNames>"#,
    r#"<calcPr fullCalcOnLoad="1"/>"#,
);

const THROUGH_NAMES: &str = concat!(
    r#"<row r="1"><c r="A1"><f>IFERROR(Rate,0)</f><v>42</v></c>"#,
    r#"<c r="B1"><f>IFERROR(Plain,0)</f><v>43</v></c>"#,
    r#"<c r="C1"><f>IFERROR(Chain,0)</f><v>44</v></c>"#,
    r#"<c r="D1" t="b"><f>ISREF(Rate)</f><v>0</v></c>"#,
    r#"<c r="E1"><f>IFERROR(Loop,7)</f><v>5</v></c></row>"#,
);

#[test]
fn external_references_through_defined_names_keep_cached_values() {
    let mut workbook =
        Workbook::open_recalculated(&package(THROUGH_NAMES, NAMES), options()).unwrap();
    workbook.recalculate_all(options());
    for (address, cached) in [("A1", 42.0), ("B1", 43.0), ("C1", 44.0)] {
        assert_eq!(stored(&workbook, address).1, number(cached), "{address}");
    }
    assert_eq!(stored(&workbook, "D1").1, CellValue::Bool { value: false });
    assert_eq!(stored(&workbook, "E1").1, number(7.0));
    workbook
        .edit_cell(SheetId(0), cell("A2"), "=IFERROR(Chain,0)", options())
        .unwrap();
    apply(&mut workbook, json!([set_formula("B2", "Plain*2")]));
    workbook.recalculate_all(options());
    assert_eq!(
        stored(&workbook, "A2"),
        (Some("IFERROR(Chain,0)".into()), CellValue::Empty)
    );
    assert_eq!(
        stored(&workbook, "B2"),
        (Some("Plain*2".into()), CellValue::Empty)
    );
}

#[test]
fn typed_external_references_are_formulas_that_open_nothing() {
    let mut workbook = Workbook::open(&external_book("")).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A2"), "=[1]Sheet1!B2*2", options())
        .unwrap();
    workbook
        .edit_cell(SheetId(0), cell("B2"), "=SUM([3]Data!A1:A9)", options())
        .unwrap();
    apply(&mut workbook, json!([set_formula("C2", "[1]!Rate*A1")]));
    assert_eq!(
        stored(&workbook, "A2"),
        (Some("[1]Sheet1!B2*2".into()), CellValue::Empty)
    );
    assert_eq!(
        stored(&workbook, "B2"),
        (Some("SUM([3]Data!A1:A9)".into()), CellValue::Empty)
    );
    assert_eq!(stored(&workbook, "C2").0.as_deref(), Some("[1]!Rate*A1"));
    for input in [
        "=[1]Sheet1!",
        "=[x]Sheet1!A1",
        "=[1]2019!A1",
        "=A1[1]Sheet1!A1",
        "=[1]!123",
        "=[1]!\"hello\"",
        "=[1]!(1+2)",
        "=[1]![2]Sheet1!A1",
        "=Table1[[1]!a]",
        "=Table1[ [1]Sheet1!A1]",
        "=[1]!Sheet1!A1",
        "=[1]!Sheet1:Sheet2!A1",
        "=[1]!'My sheet'!A1",
        "=[1]!A1",
    ] {
        workbook
            .edit_cell(SheetId(0), cell("D2"), input, options())
            .unwrap();
        assert_eq!(
            stored(&workbook, "D2"),
            (
                None,
                CellValue::Text {
                    value: input.into()
                }
            ),
            "{input}"
        );
    }
    for formula in ["Table1[Col'#]+[1]Sheet1!A1", "[1]!Table1[Col'#]*2"] {
        workbook
            .edit_cell(SheetId(0), cell("E2"), &format!("={formula}"), options())
            .unwrap();
        assert_eq!(stored(&workbook, "E2").0.as_deref(), Some(formula));
        apply(&mut workbook, json!([set_formula("F2", formula)]));
        assert_eq!(stored(&workbook, "F2").0.as_deref(), Some(formula));
    }
    workbook
        .edit_cell(SheetId(0), cell("D2"), "=\"[1]Sheet1!A1\"", options())
        .unwrap();
    assert_eq!(
        stored(&workbook, "D2").0.as_deref(),
        Some("\"[1]Sheet1!A1\"")
    );
}

#[test]
fn external_references_keep_their_cached_values_through_a_full_calculation() {
    let source = external_book(r#"<calcPr fullCalcOnLoad="1"/>"#);
    let mut workbook = Workbook::open_recalculated(&source, options()).unwrap();
    workbook.recalculate_all(options());
    for (address, formula, cached) in EXTERNAL {
        assert_eq!(
            stored(&workbook, address),
            (Some((*formula).to_owned()), number(*cached))
        );
    }
}

fn assert_calculation_keeps_formula_caches(sheet_name: &str, formulas: &[&str]) {
    let mut sheet_data = String::from(
        r#"<row r="1"><c r="A1"><v>7</v></c><c r="E1"><f>A1*2</f><v>-1</v></c></row>"#,
    );
    let cases = formulas
        .iter()
        .zip(4..)
        .map(|(&formula, row)| (formula, row, row * 10));
    for (formula, row, cached) in cases.clone() {
        sheet_data.push_str(&format!(
            r#"<row r="{row}"><c r="D{row}"><f>{formula}</f><v>{cached}</v></c></row>"#,
        ));
    }
    let assert_cached = |workbook: &Workbook| {
        assert_eq!(stored(workbook, "E1").1, number(14.0));
        for (formula, row, cached) in cases.clone() {
            assert_eq!(
                stored(workbook, &format!("D{row}")),
                (Some(formula.to_owned()), number(f64::from(cached))),
                "{formula}"
            );
        }
    };
    let source = package_on_sheet(sheet_name, &sheet_data, r#"<calcPr fullCalcOnLoad="1"/>"#);
    let mut workbook = Workbook::open_recalculating_stale(&source, options(), None).unwrap();
    assert_cached(&workbook);
    workbook.recalculate_all(options());
    assert_cached(&workbook);
}

#[test]
fn spaced_sheet_qualifiers_preserve_cached_formula_results() {
    assert_calculation_keeps_formula_caches(
        "Sheet1",
        &[
            "Sheet1 !A1",
            "Sheet1! A1",
            "Sheet1! #REF!",
            "'[1]Sheet1'! #REF!",
        ],
    );
    assert_calculation_keeps_formula_caches("My Sheet", &["'My Sheet' !A1", "'My Sheet'! A1"]);
}

#[test]
fn swallowed_range_endpoints_preserve_cached_formula_results() {
    assert_calculation_keeps_formula_caches(
        "B2",
        &["SUM(B2!A1:B2!C3)", "IFERROR(SUM(A1:B2[Col]),0)"],
    );
    assert_calculation_keeps_formula_caches(
        "Tax",
        &[
            "IFERROR(SUM(Amt:Tax[Col]),0)",
            "IFERROR(SUM(Amt:Tax!B5),0)",
            "IFERROR(SUM(Amt:IF(C1,B1,B2)),0)",
            "IFERROR(SUM(Tax!1:3!B5),0)",
        ],
    );
}

#[test]
fn range_to_cell_shaped_calls_preserve_cached_formula_results() {
    assert_calculation_keeps_formula_caches(
        "Sheet1",
        &[
            "A1:LOG10(1)",
            "Sheet1!A1:LOG10(1)",
            "A1:'Sheet1'!LOG10(1)",
            "IFERROR(Sheet1!LOG10(1),0)",
        ],
    );
}

#[test]
fn workbook_path_qualifiers_preserve_cached_formula_results() {
    assert_calculation_keeps_formula_caches(
        "Sheet1",
        &[
            r"IFERROR('C:\Books\Data.xlsx'!A1,0)",
            "IFERROR('https://example.com/book.xlsx'!LocalRate,0)",
            "IFERROR('D:/Books/[Data.xlsx]Sheet1'!A1,0)",
            r"IFERROR('\\server\Books\Data.xlsx'!A1,0)",
        ],
    );
}

#[test]
fn defined_name_range_joins_recalculate_to_the_release_value() {
    for (sheet, qualifier) in [("Data", "Data"), ("Data Set", "'Data Set'")] {
        let formula = format!("SUM(Start_1:{qualifier}!B2)");
        let mut sheet_data = format!(
            r#"<row r="1"><c r="A1"><v>10</v></c><c r="C1"><f>{formula}</f><v>999</v></c></row><row r="2"><c r="B2"><v>20</v></c></row>"#,
        );
        let shapes = [
            (format!("ROWS(Start_1:{qualifier}!B2:C3)"), 3.0),
            (format!("COLUMNS(Start_1:{qualifier}!$A:$B)"), 2.0),
            (format!("ROWS(Start_1:{qualifier}!1:3)"), 3.0),
        ];
        for (row, (formula, _)) in (4..).zip(&shapes) {
            sheet_data.push_str(&format!(
                r#"<row r="{row}"><c r="D{row}"><f>{formula}</f><v>999</v></c></row>"#,
            ));
        }
        let workbook_tail = format!(
            r#"<definedNames><definedName name="Start_1">{qualifier}!$A$1</definedName></definedNames><calcPr fullCalcOnLoad="1"/>"#,
        );
        let source = package_on_sheet(sheet, &sheet_data, &workbook_tail);
        let mut workbook = Workbook::open_recalculating_stale(&source, options(), None).unwrap();
        assert_eq!(
            stored(&workbook, "C1"),
            (Some(formula.clone()), number(30.0))
        );
        for (row, (formula, expected)) in (4..).zip(shapes) {
            assert_eq!(
                stored(&workbook, &format!("D{row}")),
                (Some(formula), number(expected))
            );
        }
        workbook
            .edit_cell(SheetId(0), cell("A1"), "11", options())
            .unwrap();
        assert_eq!(stored(&workbook, "C1").1, number(31.0));
        workbook
            .edit_cell(SheetId(0), cell("D1"), &format!("={formula}"), options())
            .unwrap();
        assert_eq!(stored(&workbook, "D1"), (Some(formula), number(31.0)));
    }
}

#[test]
fn indirect_with_spaced_external_qualifiers_recalculates_to_release_values() {
    let cases = [
        (r#"IFERROR(INDIRECT("'[1]Sheet1'! #REF!"),0)"#, number(0.0)),
        (r#"IFERROR(INDIRECT("'[1]Sheet1' !A1"),0)"#, number(0.0)),
        (r#"IFERROR(INDIRECT("[1]Sheet1 !A1"),0)"#, number(0.0)),
        (
            r#"INDIRECT("'[1]Sheet1' !R1C1",FALSE)"#,
            CellValue::Error {
                value: xlsx_model::ErrorValue::Ref,
            },
        ),
    ];
    let data: String = cases
        .iter()
        .enumerate()
        .map(|(index, (formula, _))| {
            let row = index + 1;
            format!(r#"<row r="{row}"><c r="D{row}"><f>{formula}</f><v>42</v></c></row>"#)
        })
        .collect();
    let source = package(&data, r#"<calcPr fullCalcOnLoad="1"/>"#);
    let mut workbook = Workbook::open_recalculating_stale(&source, options(), None).unwrap();
    workbook.recalculate_all(options());
    for (index, (formula, value)) in cases.into_iter().enumerate() {
        assert_eq!(
            stored(&workbook, &format!("D{}", index + 1)),
            (Some(formula.to_owned()), value),
            "{formula}"
        );
    }
}

#[test]
fn unreadable_defined_names_keep_callers_cached_and_refuse_authoring() {
    let sheet_data = concat!(
        r#"<row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>Target_1</f><v>41</v></c>"#,
        r#"<c r="C1"><f>IFERROR(Target_1,0)</f><v>42</v></c><c r="D1"><f>A1*2</f><v>-1</v></c></row>"#,
    );
    let name_error = CellValue::Error {
        value: ErrorValue::Name,
    };
    for (formula, direct, caught) in [
        ("[1]Data!A1", number(41.0), number(42.0)),
        (
            r"'C:\Books\Data.xlsx'!A1",
            CellValue::Error {
                value: ErrorValue::Ref,
            },
            number(0.0),
        ),
        ("Sheet1 !A1", name_error.clone(), number(0.0)),
        ("Sheet1! #REF!", name_error.clone(), number(0.0)),
        ("SUM(A1:B2[Col])", number(41.0), number(0.0)),
        ("A1:LOG10(1)", name_error, number(0.0)),
    ] {
        let workbook_tail = format!(
            r#"<definedNames><definedName name="Target_1">{formula}</definedName></definedNames><calcPr fullCalcOnLoad="1"/>"#,
        );
        let source = package(sheet_data, &workbook_tail);
        let mut workbook = Workbook::open_recalculating_stale(&source, options(), None).unwrap();
        workbook.recalculate_all(options());
        assert_eq!(stored(&workbook, "B1").1, direct, "{formula}");
        assert_eq!(stored(&workbook, "C1").1, caught, "{formula}");
        assert_eq!(stored(&workbook, "D1").1, number(14.0));
        let result = workbook.apply_ops(
            vec![Op::SetDefinedName {
                name: "New_Target".into(),
                refers_to: formula.into(),
                scope: None,
                hidden: None,
            }],
            options(),
        );
        assert!(
            matches!(
                result,
                Err(betteroffice_xlsx::Error::Operation(
                    xlsx_ops::OpError::InvalidDefinedName(_)
                ))
            ),
            "{formula}: {result:?}"
        );
    }
}

#[test]
fn rejected_definitions_recalculate_callers_as_in_the_release() {
    let sheet_data = concat!(
        r#"<row r="1"><c r="A1"><f>Target_1</f><v>41</v></c>"#,
        r#"<c r="B1"><f>IFERROR(Target_1,0)</f><v>42</v></c>"#,
        r#"<c r="C1" t="b"><f>ISERROR(Target_1)</f><v>0</v></c>"#,
        r#"<c r="D1"><f>Target_1</f></c></row>"#,
    );
    for (definition, keeps_direct_cache) in [
        ("SUM(1,", false),
        ("Jan:Mar!A1", false),
        ("A1:C3 B2:D4", false),
        ("[@Amount]", true),
        ("Sheet1!Table1[Amount]", true),
    ] {
        let tail = format!(
            r#"<definedNames><definedName name="Target_1">{definition}</definedName></definedNames>"#,
        );
        let mut workbook = Workbook::open(&package(sheet_data, &tail)).unwrap();
        workbook.recalculate_all(options());
        let error = CellValue::Error {
            value: ErrorValue::Name,
        };
        let cached_direct = if keeps_direct_cache {
            number(41.0)
        } else {
            error.clone()
        };
        assert_eq!(stored(&workbook, "A1").1, cached_direct, "{definition}");
        assert_eq!(stored(&workbook, "B1").1, number(0.0), "{definition}");
        let is_error = if keeps_direct_cache {
            CellValue::Bool { value: false }
        } else {
            CellValue::Bool { value: true }
        };
        assert_eq!(stored(&workbook, "C1").1, is_error, "{definition}");
        assert_eq!(stored(&workbook, "D1").1, error, "{definition}");
        workbook
            .edit_cell(SheetId(0), cell("E1"), "=IFERROR(Target_1,0)", options())
            .unwrap();
        assert_eq!(stored(&workbook, "E1").1, number(0.0), "{definition}");
    }
}

#[test]
fn a_definition_read_by_the_release_keeps_its_incremental_dependencies() {
    let source = package(
        r#"<row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>Target_1</f><v>999</v></c></row>"#,
        r#"<definedNames><definedName name="Target_1">A1+IFERROR('C:\Books\Data.xlsx'!A1,0)</definedName></definedNames>"#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    workbook.recalculate_all(options());
    assert_eq!(stored(&workbook, "B1").1, number(7.0));
    workbook
        .edit_cell(SheetId(0), cell("A1"), "8", options())
        .unwrap();
    assert_eq!(stored(&workbook, "B1").1, number(8.0));
}

#[test]
fn array_error_handlers_discard_unused_definition_errors_as_in_the_release() {
    let sheet_data = concat!(
        r#"<row r="1"><c r="A1"><f>SUM(IFERROR({1,2},Target_1))</f><v>999</v></c>"#,
        r#"<c r="B1"><f>SUM(IFNA({1,2},Target_1))</f><v>999</v></c>"#,
        r#"<c r="C1"><f>SUM(IFERROR({1,#N/A},Target_1))</f><v>999</v></c>"#,
        r#"<c r="D1"><f>SUM(IFNA({1,#N/A},Target_1))</f><v>999</v></c></row>"#,
    );
    for (definition, used) in [
        (
            "SUM(1,",
            CellValue::Error {
                value: ErrorValue::Name,
            },
        ),
        (
            r"'C:\Books\Data.xlsx'!A1",
            CellValue::Error {
                value: ErrorValue::Ref,
            },
        ),
        ("[@Amount]", number(999.0)),
    ] {
        let tail = format!(
            r#"<definedNames><definedName name="Target_1">{definition}</definedName></definedNames>"#,
        );
        let mut workbook = Workbook::open(&package(sheet_data, &tail)).unwrap();
        workbook.recalculate_all(options());
        for address in ["A1", "B1"] {
            assert_eq!(stored(&workbook, address).1, number(3.0), "{definition}");
        }
        for address in ["C1", "D1"] {
            assert_eq!(stored(&workbook, address).1, used, "{definition}");
        }
    }
}

#[test]
fn proposals_and_collaboration_recalculate_without_overwriting_unreadable_caches() {
    let source = package(
        concat!(
            r#"<row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>Sheet1 !A1</f><v>42</v></c>"#,
            r#"<c r="C1"><f>A1*2</f><v>14</v></c></row>"#,
        ),
        "",
    );
    let mut left = Workbook::open_collaborative(&source, 501).unwrap();
    let mut right = Workbook::open_collaborative(&source, 502).unwrap();
    let baseline = right.encode_state_vector_v1();
    let proposal = propose(&mut left, &[("A1", "8"), ("B1", "=Sheet1 !A1")]);
    assert!(
        proposal
            .edits
            .iter()
            .any(|edit| edit.a1 == "B1" && edit.new_text == "42")
    );
    assert!(
        proposal
            .ghosts
            .iter()
            .any(|ghost| ghost.col == 2 && ghost.new_text == "16")
    );
    left.accept_proposal(&proposal.id, false, options())
        .unwrap();
    right
        .apply_update_v1(&left.encode_diff_v1(&baseline).unwrap(), options())
        .unwrap();
    for workbook in [&left, &right] {
        assert_eq!(stored(workbook, "A1").1, number(8.0));
        assert_eq!(
            stored(workbook, "B1"),
            (Some("Sheet1 !A1".into()), number(42.0))
        );
        assert_eq!(stored(workbook, "C1").1, number(16.0));
    }
}

#[test]
fn moving_a_cell_rewrites_a_rejected_formula_and_keeps_its_cache() {
    let source = package(
        r#"<row r="1"><c r="A1"><v>7</v></c><c r="D1"><f>A1+'C:\Books\Data.xlsx'!A1</f><v>42</v></c></row>"#,
        "",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .move_range(
            SheetId(0),
            CellRange::new(cell("A1"), cell("A1")),
            cell("B1"),
            options(),
        )
        .unwrap();
    assert_eq!(stored(&workbook, "B1").1, number(7.0));
    assert_eq!(
        stored(&workbook, "D1"),
        (Some(r"B1+'C:\Books\Data.xlsx'!A1".into()), number(42.0))
    );
}

const LEADING_EQUALS: &str = concat!(
    r#"<row r="1"><c r="A1"><v>21</v></c><c r="B1"><f>=A1*2</f><v>0</v></c>"#,
    r#"<c r="C1" t="str"><f>=HYPERLINK("https://example.com/a","site")</f><v>site</v></c></row>"#,
);

#[test]
fn formulas_stored_with_a_leading_equals_show_one_and_recommit_unchanged() {
    let source = package(LEADING_EQUALS, "");
    let mut workbook = Workbook::open(&source).unwrap();
    let version = workbook.version();
    assert_eq!(shown(&workbook, "B1"), "=A1*2");
    assert_eq!(
        shown(&workbook, "C1"),
        r#"=HYPERLINK("https://example.com/a","site")"#
    );
    recommit_all(&mut workbook, &["B1", "C1"]);
    for address in ["B1", "C1"] {
        let input = shown(&workbook, address);
        workbook
            .edit_cell(SheetId(0), cell(address), &input, options())
            .unwrap();
    }
    apply(&mut workbook, json!([set_formula("B1", "A1*2")]));
    apply(&mut workbook, json!([set_input("B1", "=A1*2")]));
    assert_eq!(stored(&workbook, "B1"), (Some("=A1*2".into()), number(0.0)));
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), source);
}

#[test]
fn formulas_stored_with_a_leading_equals_evaluate_as_the_formula() {
    let source = package(LEADING_EQUALS, r#"<calcPr fullCalcOnLoad="1"/>"#);
    let mut workbook = Workbook::open_recalculated(&source, options()).unwrap();
    assert_eq!(stored(&workbook, "B1").1, number(42.0));
    assert_eq!(
        stored(&workbook, "C1").1,
        CellValue::Text {
            value: "site".into()
        }
    );
    workbook
        .edit_cell(SheetId(0), cell("A1"), "5", options())
        .unwrap();
    assert_eq!(
        stored(&workbook, "B1"),
        (Some("=A1*2".into()), number(10.0))
    );
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        stored(&reopened, "B1"),
        (Some("=A1*2".into()), number(10.0))
    );
    assert_eq!(shown(&reopened, "B1"), "=A1*2");
}

const TEXT_FORMATTED: &str =
    r#"<row r="1"><c r="A1"><v>21</v></c><c r="B1" s="1"><f>A1*2</f><v>42</v></c></row>"#;

/// Excel shows a formula re-entered in a Text-formatted cell as text: "If you then edit the
/// formula in the linked cell, the formula is displayed in the cell rather than the value"
/// (Microsoft KB 135506).
#[test]
fn typing_a_formula_into_a_text_formatted_cell_enters_text_as_excel_does() {
    let text = CellValue::Text {
        value: "=A1*2".into(),
    };
    let source = package(TEXT_FORMATTED, "");
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(shown(&workbook, "B1"), "=A1*2");
    let input = shown(&workbook, "B1");
    workbook
        .edit_cell(SheetId(0), cell("B1"), &input, options())
        .unwrap();
    assert_eq!(stored(&workbook, "B1"), (None, text.clone()));

    let mut workbook = Workbook::open(&source).unwrap();
    recommit_all(&mut workbook, &["B1"]);
    assert_eq!(stored(&workbook, "B1"), (None, text.clone()));

    let mut workbook = Workbook::open(&source).unwrap();
    apply(&mut workbook, json!([set_input("B1", "=A1*2")]));
    assert_eq!(stored(&workbook, "B1"), (None, text));
}

#[test]
fn explicit_formula_writes_ignore_a_text_number_format() {
    let source = package(TEXT_FORMATTED, "");
    let mut workbook = Workbook::open(&source).unwrap();
    let version = workbook.version();
    apply(&mut workbook, json!([set_formula("B1", "A1*2")]));
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), source);

    apply(&mut workbook, json!([set_formula("B1", "A1*3")]));
    assert_eq!(stored(&workbook, "B1"), (Some("A1*3".into()), number(63.0)));

    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: cell("B1"),
                cell: CellState {
                    value: CellValue::Empty,
                    formula: Some("A1+1".into()),
                    style: Some(1),
                },
            }],
            options(),
        )
        .unwrap();
    assert_eq!(stored(&workbook, "B1"), (Some("A1+1".into()), number(22.0)));
}
