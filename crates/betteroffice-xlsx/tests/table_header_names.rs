use std::collections::HashSet;

use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRange, CellRef, CellValue, EditApplication,
    EditFailureCode, EditHistory, EditOperation, EditRequest, EditStep, Error, Op,
    ProposalEditInput, ProposalRequest, RangeAddress, RangeTarget, SheetId, StylePatch, Workbook,
    WorkbookModel,
};
use xlsx_model::CellProvider;
use xlsx_ops::{CellState, OpError};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SPREADSHEET: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml";
const SHEET1: SheetId = SheetId(0);

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml = format!(r#"<Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#
        ));
    }
    xml + "</Relationships>"
}

fn worksheet(rows: &str, table_ids: &[u32]) -> String {
    let count = table_ids.len();
    let mut xml = format!(
        r#"<worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData>{rows}</sheetData><tableParts count="{count}">"#
    );
    for id in table_ids {
        xml.push_str(&format!(r#"<tablePart r:id="rId{id}"/>"#));
    }
    xml + "</tableParts></worksheet>"
}

fn table(id: u32, range: &str, columns: &[&str]) -> String {
    let count = columns.len();
    let mut xml = format!(
        r#"<table xmlns="{MAIN}" id="{id}" name="Table{id}" displayName="Table{id}" ref="{range}" headerRowCount="1" totalsRowShown="0"><autoFilter ref="{range}"/><tableColumns count="{count}">"#
    );
    for (index, name) in columns.iter().enumerate() {
        let col = index + 1;
        xml.push_str(&format!(r#"<tableColumn id="{col}" name="{name}"/>"#));
    }
    xml + "</tableColumns></table>"
}

fn book() -> Vec<u8> {
    let mut types = String::from(concat!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    ));
    for (path, kind) in [
        ("xl/workbook.xml", "sheet.main"),
        ("xl/worksheets/sheet1.xml", "worksheet"),
        ("xl/worksheets/sheet2.xml", "worksheet"),
        ("xl/sharedStrings.xml", "sharedStrings"),
        ("xl/tables/table1.xml", "table"),
        ("xl/tables/table2.xml", "table"),
        ("xl/tables/table3.xml", "table"),
    ] {
        types.push_str(&format!(
            r#"<Override PartName="/{path}" ContentType="{SPREADSHEET}.{kind}+xml"/>"#
        ));
    }
    types.push_str("</Types>");

    let rows1 = concat!(
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="s"><v>0</v></c>"#,
        r#"<c r="C1" t="inlineStr"><is><t>Rev</t></is></c><c r="E1" t="inlineStr"><is><t>Key</t></is></c>"#,
        r#"<c r="F1" t="inlineStr"><is><t>Val</t></is></c></row>"#,
        r#"<row r="2"><c r="A2" t="inlineStr"><is><t>Ada</t></is></c><c r="B2" t="inlineStr"><is><t>Open</t></is></c>"#,
        r#"<c r="C2"><v>1</v></c><c r="E2"><v>1</v></c><c r="F2"><v>10</v></c></row>"#,
        r#"<row r="3"><c r="A3" t="inlineStr"><is><t>Bo</t></is></c><c r="B3" t="inlineStr"><is><t>Closed</t></is></c>"#,
        r#"<c r="C3"><v>2</v></c><c r="E3"><v>2</v></c><c r="F3"><v>20</v></c></row>"#,
        r#"<row r="4"><c r="A4" t="inlineStr"><is><t>Cy</t></is></c><c r="B4" t="inlineStr"><is><t>Pending</t></is></c>"#,
        r#"<c r="C4"><v>3</v></c></row>"#,
    );
    let rows2 = concat!(
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c><c r="B1" t="inlineStr"><is><t>Y</t></is></c></row>"#,
        r#"<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>"#,
        r#"<row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row>"#,
    );
    let parts = [
        ("[Content_Types].xml", types),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "sharedStrings", "sharedStrings.xml"),
            ]),
        ),
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{MAIN}" count="1" uniqueCount="1"><si><t>Status</t></si></sst>"#),
        ),
        ("xl/worksheets/sheet1.xml", worksheet(rows1, &[1, 2])),
        ("xl/worksheets/sheet2.xml", worksheet(rows2, &[3])),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[
                ("rId1", "table", "../tables/table1.xml"),
                ("rId2", "table", "../tables/table2.xml"),
            ]),
        ),
        (
            "xl/worksheets/_rels/sheet2.xml.rels",
            relationships(&[("rId3", "table", "../tables/table3.xml")]),
        ),
        (
            "xl/tables/table1.xml",
            table(1, "A1:C4", &["Name", "Status", "Rev"]),
        ),
        ("xl/tables/table2.xml", table(2, "E1:F3", &["Key", "Val"])),
        ("xl/tables/table3.xml", table(3, "A1:B3", &["X", "Y"])),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> &'a [u8] {
    parts
        .iter()
        .find(|(name, _)| name == path)
        .unwrap_or_else(|| panic!("missing {path}"))
        .1
        .as_slice()
}

fn replace_part(source: &[u8], path: &str, from: &str, to: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(source).unwrap();
    let (_, bytes) = parts.iter_mut().find(|(name, _)| name == path).unwrap();
    let xml = std::str::from_utf8(bytes).unwrap();
    assert!(xml.contains(from), "{path} is missing {from}");
    *bytes = xml.replace(from, to).into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_table_rule(bytes: &[u8]) -> WorkbookModel {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let model = xlsx_parse::parse_workbook(&parts).unwrap();
    let types = ooxml_opc::parse_content_types(&parts).unwrap();
    let table_type = format!("{SPREADSHEET}.table+xml");
    let table_count = parts
        .iter()
        .filter(|(path, _)| types.content_type_for(path) == Some(table_type.as_str()))
        .count();
    assert_eq!(table_count, 3);
    assert_eq!(model.tables.len(), table_count);
    for table in &model.tables {
        if table.header_rows == 0 {
            continue;
        }
        let range = table.range;
        let path = format!("xl/worksheets/sheet{}.xml", table.sheet.0 + 1);
        let worksheet = std::str::from_utf8(part(&parts, &path)).unwrap();
        let sheet = &model.sheets[table.sheet.0 as usize];
        let mut names = HashSet::new();
        assert_eq!(
            table.columns.len() as u32,
            range.end.col - range.start.col + 1,
            "{}",
            table.name,
        );
        for (col, name) in (range.start.col..=range.end.col).zip(&table.columns) {
            let at = CellRef::new(range.start.row, col);
            let address = at.to_a1();
            let reference = format!(r#"r="{address}""#);
            let attrs = worksheet
                .split("<c ")
                .skip(1)
                .map(|tag| tag.split_once('>').unwrap().0.trim_end_matches('/'))
                .find(|attrs| {
                    attrs
                        .split_ascii_whitespace()
                        .any(|attr| attr == reference.as_str())
                })
                .unwrap_or_else(|| panic!("{path} has no cell {address}"));
            assert!(
                attrs
                    .split_ascii_whitespace()
                    .any(|attr| matches!(attr, r#"t="s""# | r#"t="inlineStr""#)),
                "{path} {address} must store a shared or inline string: <c {attrs}>",
            );
            let cell = sheet.cell(at).unwrap();
            assert!(cell.formula.is_none(), "{} {address}", table.name);
            let CellValue::Text { value } = &cell.value else {
                panic!(
                    "{} {address} must contain text: {:?}",
                    table.name, cell.value
                );
            };
            assert!(
                !name.is_empty(),
                "{} {address} has an empty name",
                table.name
            );
            let folded = name.to_lowercase();
            assert_eq!(value, name, "{} {address}", table.name);
            assert!(
                names.insert(folded),
                "{} has duplicate name {name:?}",
                table.name
            );
        }
    }
    model
}

fn at(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn edit(workbook: &mut Workbook, address: &str, input: &str) {
    workbook
        .edit_cell(SHEET1, at(address), input, CalculationOptions::default())
        .unwrap();
}

fn edit_without_history(workbook: &mut Workbook, address: &str, input: &str) -> EditApplication {
    let result = workbook
        .apply_edits(&EditRequest {
            expect_version: workbook.version(),
            source: Default::default(),
            history: EditHistory::None,
            calculation: Default::default(),
            steps: vec![EditStep::new(EditOperation::SetCellInputs {
                target: RangeTarget {
                    sheet_id: "sheet:0".into(),
                    range: RangeAddress::A1 { a1: address.into() },
                },
                inputs: vec![vec![input.into()]],
            })],
        })
        .unwrap()
        .unwrap();
    assert!(result.applied);
    result
}

fn assert_refused_header(workbook: &mut Workbook, address: &str, input: &str, message: &str) {
    let version = workbook.version();
    let model = workbook.model().clone();
    let error = workbook
        .edit_cell(SHEET1, at(address), input, CalculationOptions::default())
        .unwrap_err();
    let Error::Operation(OpError::TableHeaderUnsupported(reason)) = &error else {
        panic!("expected a table header error, got {error}");
    };
    assert!(reason.contains(message), "{reason}");
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.model(), &model);
}

fn assert_saved_headers(bytes: &[u8], expected: [&str; 3]) -> Workbook {
    let parsed = assert_table_rule(bytes);
    let reopened = Workbook::open(bytes).unwrap();
    for model in [&parsed, reopened.model()] {
        assert_eq!(model.table("Table1").unwrap().columns, expected);
        for (address, name) in ["A1", "B1", "C1"].into_iter().zip(expected) {
            assert_eq!(
                model.value(SHEET1, at(address)),
                CellValue::Text { value: name.into() },
                "{address}",
            );
        }
    }
    reopened
}

#[test]
fn renaming_a_header_updates_the_cell_and_table_column() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "B1", "State");
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
}

#[test]
fn clearing_a_header_generates_a_column_name() {
    let source = book();
    let mut typed_clear = Workbook::open(&source).unwrap();
    edit(&mut typed_clear, "B1", "");

    let mut batch_clear = Workbook::open(&source).unwrap();
    batch_clear
        .edit_cells(
            SHEET1,
            &[
                CellInput {
                    cell: at("B1"),
                    input: String::new(),
                },
                CellInput {
                    cell: at("B2"),
                    input: String::new(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();

    let typed = typed_clear.save().unwrap();
    let batch = batch_clear.save().unwrap();
    assert_saved_headers(&typed, ["Name", "Column1", "Rev"]);
    let reopened = assert_saved_headers(&batch, ["Name", "Column1", "Rev"]);
    assert_eq!(reopened.model().value(SHEET1, at("B2")), CellValue::Empty);
}

#[test]
fn clearing_an_already_empty_header_generates_a_column_name() {
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet1.xml",
        r#"<c r="B1" t="s"><v>0</v></c>"#,
        "",
    );
    for raw in [false, true] {
        let mut workbook = Workbook::open(&source).unwrap();
        assert_eq!(workbook.model().value(SHEET1, at("B1")), CellValue::Empty);
        if raw {
            workbook
                .apply_ops(
                    vec![Op::SetCell {
                        sheet: SHEET1,
                        at: at("B1"),
                        cell: CellState::default(),
                    }],
                    CalculationOptions::default(),
                )
                .unwrap();
        } else {
            edit(&mut workbook, "B1", "");
        }
        assert_saved_headers(&workbook.save().unwrap(), ["Name", "Column1", "Rev"]);
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_eq!(workbook.model().value(SHEET1, at("B1")), CellValue::Empty);
        assert_eq!(
            workbook.model().table("Table1").unwrap().columns[1],
            "Status"
        );
        workbook.redo(CalculationOptions::default()).unwrap();
        assert_saved_headers(&workbook.save().unwrap(), ["Name", "Column1", "Rev"]);
    }
}

#[test]
fn duplicate_headers_receive_the_lowest_available_suffix() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "B1", "Name");
    let first = workbook.save().unwrap();
    edit(&mut workbook, "C1", "name");
    let second = workbook.save().unwrap();
    assert_saved_headers(&first, ["Name", "Name2", "Rev"]);
    assert_saved_headers(&second, ["Name", "Name2", "name3"]);
}

#[test]
fn changing_only_the_case_of_a_header_preserves_the_typed_name() {
    let source = book();
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "STATUS");
    assert_eq!(
        workbook.model().table("Table1").unwrap().columns[1],
        "STATUS"
    );
    assert_eq!(
        workbook.model().value(SHEET1, at("B1")),
        CellValue::Text {
            value: "STATUS".into()
        },
    );
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let path = "xl/tables/table1.xml";
    assert_ne!(part(&after, path), part(&before, path));
    assert_saved_headers(&saved, ["Name", "STATUS", "Rev"]);
}

#[test]
fn numeric_input_in_a_header_is_stored_as_text() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "B1", "007");
    edit(&mut workbook, "C1", "TRUE");
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "007", "TRUE"]);
}

#[test]
fn clearing_a_header_skips_a_taken_default_name() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "A1", "Column1");
    edit(&mut workbook, "B1", "");
    assert_saved_headers(&workbook.save().unwrap(), ["Column1", "Column2", "Rev"]);
}

#[test]
fn a_table_header_cannot_exceed_255_characters_with_or_without_a_suffix() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let name = "é".repeat(255);
    edit(&mut workbook, "A1", &name);
    let over_cell_limit = "x".repeat(xlsx_calc::eval::MAX_CELL_TEXT_CHARS + 1);
    for input in [name.clone(), format!("{name}x"), over_cell_limit] {
        assert_refused_header(&mut workbook, "B1", &input, "255-character limit");
    }
    assert_saved_headers(&workbook.save().unwrap(), [&name, "Status", "Rev"]);
}

#[test]
fn a_formula_in_a_header_is_refused_without_changing_it() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let error = workbook
        .edit_cell(SHEET1, at("B1"), "=1+1", CalculationOptions::default())
        .unwrap_err();
    assert!(matches!(
        &error,
        Error::Operation(OpError::TableHeaderUnsupported(reason))
            if reason.contains("formulas are not supported")
    ));
    let outcome = workbook
        .apply_edits(&EditRequest {
            expect_version: workbook.version(),
            source: Default::default(),
            history: Default::default(),
            calculation: Default::default(),
            steps: vec![EditStep::new(EditOperation::SetCellInputs {
                target: RangeTarget {
                    sheet_id: "sheet:0".into(),
                    range: RangeAddress::A1 { a1: "B1".into() },
                },
                inputs: vec![vec!["=1+1".into()]],
            })],
        })
        .unwrap()
        .unwrap_err();
    assert_eq!(outcome.failure.code, EditFailureCode::InvalidStep);
    assert_eq!(outcome.failure.message, error.to_string());
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "Status", "Rev"]);
}

#[test]
fn undo_and_redo_restore_the_header_cell_and_column_name() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "B1", "State");
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "Status", "Rev"]);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
}

#[test]
fn history_refuses_a_caption_taken_by_a_sibling_without_history() {
    for (redo, caption) in [(false, "status"), (true, "state")] {
        let mut workbook = Workbook::open(&book()).unwrap();
        edit(&mut workbook, "B1", "State");
        if redo {
            workbook.undo(CalculationOptions::default()).unwrap();
        }
        edit_without_history(&mut workbook, "C1", caption);
        let version = workbook.version();
        let history = workbook.history_state();
        let model = workbook.model().clone();
        let authored = workbook.encode_state_as_update_v1();
        let error = if redo {
            workbook.redo(CalculationOptions::default())
        } else {
            workbook.undo(CalculationOptions::default())
        }
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Operation(OpError::TableHeaderUnsupported(_))
        ));
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.history_state(), history);
        assert_eq!(workbook.model(), &model);
        assert_eq!(workbook.encode_state_as_update_v1(), authored);
    }
}

#[test]
fn history_rewrites_formulas_entered_since_the_rename_or_undo() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "H1", "=COUNTA(Table1[status])");
    edit(&mut workbook, "B1", "State");
    edit_without_history(&mut workbook, "H2", "=SUM(Table1[State])");
    workbook.undo(CalculationOptions::default()).unwrap();
    let formula = |workbook: &Workbook, address| {
        workbook.model().sheets[0]
            .cell(at(address))
            .unwrap()
            .formula
            .clone()
    };
    assert_eq!(
        formula(&workbook, "H1").as_deref(),
        Some("COUNTA(Table1[status])")
    );
    assert_eq!(
        formula(&workbook, "H2").as_deref(),
        Some("SUM(Table1[Status])")
    );
    edit_without_history(&mut workbook, "H3", "=SUM(Table1[Status])");
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(
        formula(&workbook, "H1").as_deref(),
        Some("COUNTA(Table1[State])")
    );
    for address in ["H2", "H3"] {
        assert_eq!(
            formula(&workbook, address).as_deref(),
            Some("SUM(Table1[State])")
        );
    }
}

#[test]
fn a_formula_entered_between_two_header_renames_uses_the_final_caption() {
    let inputs = [
        ("B1", "State"),
        ("H1", "=COUNTA(Table1[State])"),
        ("B1", "Status"),
    ];
    let options = CalculationOptions::default();
    for proposal in [false, true] {
        let mut workbook = Workbook::open(&book()).unwrap();
        if proposal {
            let proposal = workbook
                .propose(
                    ProposalRequest {
                        agent_id: "agent".into(),
                        note: None,
                        edits: inputs
                            .iter()
                            .map(|&(address, input)| ProposalEditInput {
                                sheet: SHEET1,
                                cell: at(address),
                                input: input.into(),
                                number_format: None,
                            })
                            .collect(),
                    },
                    options,
                )
                .unwrap();
            assert!(
                workbook
                    .accept_proposal(&proposal.id, true, options)
                    .unwrap()
                    .mutation
                    .applied
            );
        } else {
            let edits = inputs.map(|(address, input)| CellInput {
                cell: at(address),
                input: input.into(),
            });
            assert!(workbook.edit_cells(SHEET1, &edits, options).unwrap().applied);
        }
        let saved = workbook.save().unwrap();
        let reopened = assert_saved_headers(&saved, ["Name", "Status", "Rev"]);
        for model in [workbook.model(), reopened.model()] {
            assert_eq!(
                model.sheets[0].cell(at("H1")).unwrap().formula.as_deref(),
                Some("COUNTA(Table1[Status])")
            );
            assert_eq!(
                model.value(SHEET1, at("H1")),
                CellValue::Number { value: 3.0 }
            );
        }
    }
}

#[test]
fn a_batch_reports_sheets_with_rewritten_formulas_whose_values_stay_the_same() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let sheet = SheetId(1);
    workbook
        .edit_cell(
            sheet,
            at("D1"),
            "=COUNTA(Table1[Status])",
            CalculationOptions::default(),
        )
        .unwrap();
    let applied = edit_without_history(&mut workbook, "B1", "State");
    assert_eq!(applied.changed_sheets, ["sheet:0", "sheet:1"]);
    assert!(applied.calculation.changed.is_empty());
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    for model in [workbook.model(), reopened.model()] {
        assert_eq!(
            model
                .sheet(sheet)
                .unwrap()
                .cell(at("D1"))
                .unwrap()
                .formula
                .as_deref(),
            Some("COUNTA(Table1[State])")
        );
        assert_eq!(model.value(sheet, at("D1")), CellValue::Number { value: 3.0 });
    }
}

#[test]
fn undo_keeps_a_later_header_rename_outside_history() {
    let source = replace_part(
        &book(),
        "xl/workbook.xml",
        "</workbook>",
        concat!(
            r#"<definedNames><definedName name="Keys">Table2[Key]</definedName>"#,
            r#"<definedName name="Both">COUNTA(Table1[Status],Table2[Key])</definedName>"#,
            "</definedNames></workbook>",
        ),
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "H1", "=COUNTA(Table1[Status],Table2[Key])");
    edit(&mut workbook, "B1", "State");
    workbook
        .apply_edits(&EditRequest {
            expect_version: workbook.version(),
            source: Default::default(),
            history: EditHistory::None,
            calculation: Default::default(),
            steps: vec![EditStep::new(EditOperation::SetCellInputs {
                target: RangeTarget {
                    sheet_id: "sheet:0".into(),
                    range: RangeAddress::A1 { a1: "E1".into() },
                },
                inputs: vec![vec!["Id".into()]],
            })],
        })
        .unwrap()
        .unwrap();
    for (undo, caption) in [(true, "Status"), (false, "State")] {
        if undo {
            workbook.undo(CalculationOptions::default()).unwrap();
        } else {
            workbook.redo(CalculationOptions::default()).unwrap();
        }
        let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", caption, "Rev"]);
        let model = reopened.model();
        let formula = format!("COUNTA(Table1[{caption}],Table2[Id])");
        assert_eq!(model.table("Table2").unwrap().columns[0], "Id");
        assert_eq!(model.defined_names[0].formula, "Table2[Id]");
        assert_eq!(model.defined_names[1].formula, formula);
        assert_eq!(
            model.sheets[0].cell(at("H1")).unwrap().formula.as_deref(),
            Some(formula.as_str())
        );
    }
}

#[test]
fn an_untouched_header_can_differ_from_its_caption_in_case() {
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet1.xml",
        "<t>Name</t>",
        "<t>NAME</t>",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "State");
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().table("Table1").unwrap().columns,
        ["Name", "State", "Rev"]
    );
    assert_eq!(
        reopened.model().value(SHEET1, at("A1")),
        CellValue::Text {
            value: "NAME".into()
        }
    );
}

#[test]
fn a_header_edit_preserves_other_table_parts_byte_for_byte() {
    let source = book();
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "State");
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in ["xl/tables/table2.xml", "xl/tables/table3.xml"] {
        assert_eq!(part(&after, path), part(&before, path), "{path}");
    }
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().value(SHEET1, at("B1")),
        CellValue::Text {
            value: "State".into()
        },
    );
}

#[test]
fn a_data_cell_edit_preserves_the_table_part_byte_for_byte() {
    let source = book();
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B3", "Reviewed");
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let path = "xl/tables/table1.xml";
    assert_eq!(part(&after, path), part(&before, path));
    let reopened = assert_saved_headers(&saved, ["Name", "Status", "Rev"]);
    assert_eq!(
        reopened.model().value(SHEET1, at("B3")),
        CellValue::Text {
            value: "Reviewed".into()
        },
    );
}

#[test]
fn an_unrelated_edit_preserves_a_preexisting_header_mismatch() {
    let source = replace_part(
        &book(),
        "xl/sharedStrings.xml",
        "<t>Status</t>",
        "<t>State</t>",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(
        workbook.model().table("Table1").unwrap().columns[1],
        "Status"
    );
    edit(&mut workbook, "B3", "Reviewed");
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let path = "xl/tables/table1.xml";
    assert_eq!(part(&after, path), part(&before, path));
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().table("Table1").unwrap().columns[1],
        "Status"
    );
    assert_eq!(
        reopened.model().value(SHEET1, at("B1")),
        CellValue::Text {
            value: "State".into(),
        }
    );
}

#[test]
fn styling_header_cells_preserves_the_table_parts() {
    let source = replace_part(&book(), "xl/sharedStrings.xml", "Status", "_x0001_");
    let source = replace_part(
        &source,
        "xl/worksheets/sheet1.xml",
        r#"<c r="C1" t="inlineStr"><is><t>Rev</t></is></c>"#,
        "",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .patch_range_style(
            SHEET1,
            CellRange::parse_a1("A1:F1").unwrap(),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in [
        "xl/tables/table1.xml",
        "xl/tables/table2.xml",
        "xl/tables/table3.xml",
    ] {
        assert_eq!(part(&after, path), part(&before, path), "{path}");
    }
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().value(SHEET1, at("B1")),
        CellValue::Text {
            value: "\u{1}".into(),
        }
    );
    assert_eq!(reopened.model().value(SHEET1, at("C1")), CellValue::Empty);
}

#[test]
fn an_unrelated_edit_preserves_tables_with_duplicate_display_names_on_separate_sheets() {
    let path = "xl/tables/table3.xml";
    let source = replace_part(
        &book(),
        path,
        r#"displayName="Table3""#,
        r#"displayName="Table1""#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B3", "Reviewed");
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in ["xl/tables/table1.xml", path] {
        assert_eq!(part(&after, path), part(&before, path), "{path}");
    }
    for caption in ["State", "STATUS"] {
        assert_refused_header(&mut workbook, "B1", caption, "duplicate display name");
    }
    assert_eq!(workbook.save().unwrap(), saved);
}

#[test]
fn a_header_rename_is_refused_when_two_sheets_share_the_table_part() {
    let source = replace_part(
        &book(),
        "xl/worksheets/_rels/sheet2.xml.rels",
        "table3.xml",
        "table1.xml",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    assert_refused_header(&mut workbook, "B1", "State", "is shared");
}

#[test]
fn a_header_rename_is_refused_when_the_source_column_name_is_missing() {
    let source = replace_part(
        &book(),
        "xl/tables/table1.xml",
        r#" id="2" name="Status""#,
        r#" id="2""#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    assert_refused_header(&mut workbook, "B1", "State", "missing table column name");
}

#[test]
fn a_header_rename_is_refused_when_an_untouched_sibling_is_invalid() {
    for (cell, caption) in [
        ("", "Rev"),
        (r#"<c r="C1"><v>7</v></c>"#, "Rev"),
        (
            r#"<c r="C1" t="inlineStr"><is><t>Name</t></is></c>"#,
            "Name",
        ),
    ] {
        let source = replace_part(
            &book(),
            "xl/worksheets/sheet1.xml",
            r#"<c r="C1" t="inlineStr"><is><t>Rev</t></is></c>"#,
            cell,
        );
        let source = replace_part(
            &source,
            "xl/tables/table1.xml",
            r#"name="Rev""#,
            &format!(r#"name="{caption}""#),
        );
        let mut workbook = Workbook::open(&source).unwrap();
        assert_refused_header(&mut workbook, "B1", "State", "column 3 (C1)");
    }
}

#[test]
fn editing_the_first_data_row_of_a_headerless_table_preserves_its_part() {
    let path = "xl/tables/table2.xml";
    let source = replace_part(
        &book(),
        path,
        "headerRowCount=\"1\"",
        "headerRowCount=\"0\"",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "E1", "7");
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(part(&after, path), part(&before, path));
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().value(SHEET1, at("E1")),
        CellValue::Number { value: 7.0 }
    );
}

#[test]
fn a_table_with_totals_uses_its_first_row_as_headers() {
    let source = replace_part(
        &book(),
        "xl/tables/table1.xml",
        "totalsRowShown=\"0\"",
        "totalsRowShown=\"1\" totalsRowCount=\"1\"",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B4", "7");
    edit(&mut workbook, "B1", "State");
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    assert_eq!(reopened.model().table("Table1").unwrap().totals_rows, 1);
    assert_eq!(
        reopened.model().value(SHEET1, at("B4")),
        CellValue::Number { value: 7.0 }
    );
}

#[test]
fn a_header_with_xml_special_characters_and_a_line_break_round_trips() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let name = "A<B & \"C\"\nNext\u{1}";
    edit(&mut workbook, "B1", name);
    let saved = workbook.save().unwrap();
    assert_saved_headers(&saved, ["Name", name, "Rev"]);
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(part(&parts, "xl/tables/table1.xml")).unwrap();
    assert!(xml.contains("_x000A_"), "{xml}");
    assert!(xml.contains("_x0001_"), "{xml}");
}

#[test]
fn an_existing_structured_reference_follows_a_header_rename_and_undo() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "H1", "=COUNTA(Table1[Status])");
    let expected = CellValue::Number { value: 3.0 };
    assert_eq!(workbook.model().value(SHEET1, at("H1")), expected);
    edit(&mut workbook, "B1", "State");
    let formula = |workbook: &Workbook| {
        workbook.model().sheets[0]
            .cell(at("H1"))
            .unwrap()
            .formula
            .clone()
    };
    assert_eq!(formula(&workbook).as_deref(), Some("COUNTA(Table1[State])"));
    assert_eq!(workbook.model().value(SHEET1, at("H1")), expected);
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    assert_eq!(formula(&reopened).as_deref(), Some("COUNTA(Table1[State])"));
    assert_eq!(reopened.model().value(SHEET1, at("H1")), expected);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(
        formula(&workbook).as_deref(),
        Some("COUNTA(Table1[Status])")
    );
    assert_eq!(workbook.model().value(SHEET1, at("H1")), expected);
}

#[test]
fn escaped_table_captions_keep_structured_references_evaluable() {
    for (caption, escaped) in [
        ("Q1:Q2", "Q1:Q2"),
        ("a'b", "a''b"),
        ("x[1]", "x'[1']"),
        ("#n", "'#n"),
    ] {
        let mut workbook = Workbook::open(&book()).unwrap();
        edit(
            &mut workbook,
            "H1",
            "=COUNTA(Table1[Status])+COUNTA(Table1[[Status]:[Rev]])",
        );
        edit(
            &mut workbook,
            "C2",
            "=COUNTA(Table1[@Status])+COUNTA(Table1[[#This Row],[Status]])",
        );
        let values = ["H1", "C2"].map(|cell| workbook.model().value(SHEET1, at(cell)));
        assert_eq!(
            values,
            [
                CellValue::Number { value: 9.0 },
                CellValue::Number { value: 2.0 },
            ]
        );
        edit(&mut workbook, "B1", caption);
        let mut reopened =
            assert_saved_headers(&workbook.save().unwrap(), ["Name", caption, "Rev"]);
        reopened.recalculate_all(CalculationOptions::default());
        let formulas = [
            format!("COUNTA(Table1[[{escaped}]])+COUNTA(Table1[[{escaped}]:[Rev]])"),
            format!("COUNTA(Table1[@[{escaped}]])+COUNTA(Table1[[#This Row],[{escaped}]])"),
        ];
        for ((address, value), formula) in ["H1", "C2"].into_iter().zip(values).zip(formulas) {
            assert_eq!(
                reopened.model().sheets[0]
                    .cell(at(address))
                    .unwrap()
                    .formula
                    .as_deref(),
                Some(formula.as_str())
            );
            assert_eq!(reopened.model().value(SHEET1, at(address)), value);
        }
    }
}

#[test]
fn qualified_row_references_recalculate_after_a_rename() {
    for (caption, escaped) in [("New Name", "New Name"), ("#n", "'#n")] {
        let mut workbook = Workbook::open(&book()).unwrap();
        edit(&mut workbook, "B2", "3");
        edit(&mut workbook, "B3", "4");
        edit(&mut workbook, "C3", "=Table1[@Status]*2");
        let expected = [CellValue::Number { value: 8.0 }];
        let values =
            |workbook: &Workbook| ["C3"].map(|address| workbook.model().value(SHEET1, at(address)));
        assert_eq!(values(&workbook), expected);
        edit(&mut workbook, "B1", caption);
        assert_eq!(values(&workbook), expected);

        let mut reopened =
            assert_saved_headers(&workbook.save().unwrap(), ["Name", caption, "Rev"]);
        reopened.recalculate_all(CalculationOptions::default());
        assert_eq!(values(&reopened), expected);
        assert_eq!(
            reopened.model().sheets[0]
                .cell(at("C3"))
                .unwrap()
                .formula
                .as_deref(),
            Some(format!("Table1[@[{escaped}]]*2").as_str())
        );
        edit(&mut reopened, "B3", "5");
        assert_eq!(
            reopened.model().value(SHEET1, at("C3")),
            CellValue::Number { value: 10.0 }
        );
    }
}

#[test]
fn a_rewrite_over_the_formula_character_limit_refuses_the_rename() {
    let limit = xlsx_calc::lexer::MAX_FORMULA_CHARS;
    let mut formula = "COUNTA(Table1[Status])".to_owned();
    let empty = r#"+LEN("")"#;
    let term = format!(r#"+LEN("{}")"#, "x".repeat(200));
    while formula.len() + term.len() + empty.len() <= limit {
        formula.push_str(&term);
    }
    let remaining = limit - formula.len() - empty.len();
    formula.push_str(&format!(r#"+LEN("{}")"#, "x".repeat(remaining)));
    for defined_name in [false, true] {
        let source = if defined_name {
            replace_part(
                &book(),
                "xl/workbook.xml",
                "</workbook>",
                &format!(
                    "<definedNames><definedName name=\"Statuses\">{formula}</definedName>\
                     </definedNames></workbook>"
                ),
            )
        } else {
            book()
        };
        let mut workbook = Workbook::open(&source).unwrap();
        if !defined_name {
            edit(&mut workbook, "H1", &format!("={formula}"));
        }
        edit(&mut workbook, "G1", "=COUNTA(Table1[Status])");
        let version = workbook.version();
        let model = workbook.model().clone();
        let error = workbook
            .edit_cell(
                SHEET1,
                at("B1"),
                "LongStatus",
                CalculationOptions::default(),
            )
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Operation(OpError::TableHeaderUnsupported(reason))
                if reason.contains("8192-character limit")
        ));
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.model(), &model);
        assert_saved_headers(&workbook.save().unwrap(), ["Name", "Status", "Rev"]);
    }
}

#[test]
fn a_header_rename_rewrites_only_the_matching_qualified_table_reference() {
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet1.xml",
        r#"<c r="C2"><v>1</v></c>"#,
        r#"<c r="C2"><f>COUNTA(Sheet2!Table3[Status])</f><v>1</v></c>"#,
    );
    let source = replace_part(
        &source,
        "xl/worksheets/sheet2.xml",
        r#"<c r="B1" t="inlineStr"><is><t>Y</t></is></c>"#,
        concat!(
            r#"<c r="B1" t="inlineStr"><is><t>Status</t></is></c>"#,
            r#"<c r="D1"><f>COUNTA('Sheet1'!Table1[Status])</f><v>3</v></c>"#,
        ),
    );
    let source = replace_part(
        &source,
        "xl/tables/table3.xml",
        r#"name="Y""#,
        r#"name="Status""#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "State");
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    for (sheet, address, formula) in [
        (0, "C2", "COUNTA(Sheet2!Table3[Status])"),
        (1, "D1", "COUNTA('Sheet1'!Table1[State])"),
    ] {
        assert_eq!(
            reopened.model().sheets[sheet]
                .cell(at(address))
                .unwrap()
                .formula
                .as_deref(),
            Some(formula)
        );
    }
}

#[test]
fn a_header_rename_preserves_other_workbook_and_table_selectors() {
    let other_refs = concat!(
        "'Other.xlsx'!OtherTable[Status]+[1]Sheet1!Other[Status]+",
        "[1]Sheet1!Table1[Status]+OtherTable[Status]+OtherTable [Status]+",
        "'Other.xlsx'!OtherTable [Status]+[1]Sheet1!Other [Status]",
    );
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet1.xml",
        r#"<c r="C2"><v>1</v></c>"#,
        &format!(r#"<c r="C2"><f>{other_refs}+[@Status]</f><v>1</v></c>"#),
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "State");
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    assert_eq!(
        reopened.model().sheets[0]
            .cell(at("C2"))
            .unwrap()
            .formula
            .as_deref(),
        Some(format!("{other_refs}+[@State]").as_str())
    );
}

#[test]
fn another_sheets_formula_and_a_defined_name_follow_a_header_rename() {
    let source = replace_part(
        &book(),
        "xl/workbook.xml",
        "</workbook>",
        r#"<definedNames><definedName name="Statuses">Table1[Status]</definedName></definedNames></workbook>"#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .edit_cell(
            SheetId(1),
            at("D1"),
            "=COUNTA(Table1[[Status]:[Rev]])",
            CalculationOptions::default(),
        )
        .unwrap();
    edit(&mut workbook, "B1", "State");
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    for model in [workbook.model(), reopened.model()] {
        assert_eq!(model.defined_names[0].formula, "Table1[State]");
        assert_eq!(
            model.sheets[1].cell(at("D1")).unwrap().formula.as_deref(),
            Some("COUNTA(Table1[[State]:[Rev]])")
        );
    }
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names[0].formula, "Table1[Status]");
    assert_eq!(
        workbook.model().sheets[1]
            .cell(at("D1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("COUNTA(Table1[[Status]:[Rev]])")
    );
}

#[test]
fn a_collaborative_session_refuses_a_header_edit() {
    let source = book();
    let mut workbook = Workbook::open_collaborative(&source, 77).unwrap();
    let version = workbook.version();
    let error = workbook
        .edit_cell(SHEET1, at("B1"), "State", CalculationOptions::default())
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Operation(OpError::TableHeaderUnsupported(reason))
            if reason.contains("shared session")
    ));
    assert_eq!(workbook.version(), version);
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "Status", "Rev"]);

    let mut peer = Workbook::open_collaborative(&source, 78).unwrap();
    workbook
        .patch_range_style(
            SHEET1,
            CellRange::parse_a1("A1:C1").unwrap(),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let update = workbook
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    let result = peer
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert!(result.applied);
    assert_saved_headers(&peer.save().unwrap(), ["Name", "Status", "Rev"]);
}

#[test]
fn a_control_character_in_a_referenced_caption_is_refused() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "H1", "=COUNTA(Table1[Status])");
    let version = workbook.version();
    let error = workbook
        .edit_cell(
            SHEET1,
            at("B1"),
            "State\u{1}",
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Operation(OpError::TableHeaderUnsupported(reason))
            if reason.contains("character XML cannot carry")
    ));
    assert_eq!(workbook.version(), version);
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "Status", "Rev"]);
    assert_eq!(
        reopened.model().sheets[0]
            .cell(at("H1"))
            .unwrap()
            .formula
            .as_deref(),
        Some("COUNTA(Table1[Status])")
    );
}

#[test]
fn a_header_rename_patches_local_and_cross_table_formulas() {
    let table1 = "xl/tables/table1.xml";
    let table2 = "xl/tables/table2.xml";
    let table3 = "xl/tables/table3.xml";
    let source = replace_part(
        &book(),
        table1,
        r#"<tableColumn id="3" name="Rev"/>"#,
        concat!(
            r#"<tableColumn id="3" name="Rev"><calculatedColumnFormula array="0">"#,
            "[@Status]</calculatedColumnFormula>",
            "<totalsRowFormula>SUM(Table1[Status])</totalsRowFormula></tableColumn>",
        ),
    );
    let source = replace_part(
        &source,
        table2,
        r#"<tableColumn id="2" name="Val"/>"#,
        concat!(
            r#"<tableColumn id="2" name="Val"><calculatedColumnFormula>"#,
            "COUNTA(Table1[Status])</calculatedColumnFormula></tableColumn>",
        ),
    );
    let source = replace_part(
        &source,
        "xl/worksheets/sheet1.xml",
        r#"<c r="C2"><v>1</v></c>"#,
        r#"<c r="C2"><f>[@Status]</f><v>1</v></c>"#,
    );
    let source = replace_part(
        &source,
        "xl/worksheets/sheet1.xml",
        r#"<c r="F2"><v>10</v></c>"#,
        r#"<c r="F2"><f>COUNTA(Table1[Status])</f><v>3</v></c>"#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "State & Code");
    let saved = workbook.save().unwrap();
    let reopened = assert_saved_headers(&saved, ["Name", "State & Code", "Rev"]);
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let expected1 = std::str::from_utf8(part(&before, table1))
        .unwrap()
        .replace(r#"name="Status""#, r#"name="State &amp; Code""#)
        .replace("[@Status]", "[@[State &amp; Code]]")
        .replace("Table1[Status]", "Table1[[State &amp; Code]]");
    let expected2 = std::str::from_utf8(part(&before, table2))
        .unwrap()
        .replace("Table1[Status]", "Table1[[State &amp; Code]]");
    assert_eq!(part(&after, table1), expected1.as_bytes());
    assert_eq!(part(&after, table2), expected2.as_bytes());
    assert_eq!(part(&after, table3), part(&before, table3));
    for (address, formula) in [
        ("C2", "[@[State & Code]]"),
        ("F2", "COUNTA(Table1[[State & Code]])"),
    ] {
        assert_eq!(
            reopened.model().sheets[0]
                .cell(at(address))
                .unwrap()
                .formula
                .as_deref(),
            Some(formula)
        );
    }
}

#[test]
fn rewriting_an_array_formula_preserves_its_source_metadata() {
    const METADATA: &str = "xl/metadata.xml";
    let source = replace_part(
        &book(),
        "[Content_Types].xml",
        "</Types>",
        &format!(r#"<Override PartName="/{METADATA}" ContentType="{SPREADSHEET}.sheetMetadata+xml"/></Types>"#),
    );
    let source = replace_part(
        &source,
        "xl/_rels/workbook.xml.rels",
        "</Relationships>",
        &format!(r#"<Relationship Id="rId4" Type="{DOC_RELS}/sheetMetadata" Target="metadata.xml"/></Relationships>"#),
    );
    let sheet = "xl/worksheets/sheet1.xml";
    let source = replace_part(
        &source,
        sheet,
        r#"<c r="E2">"#,
        r#"<c r="D2" cm="1"><f t="array" ref="D2" aca="1">SUM(Table1[Status])</f><v>0.00</v></c><c r="E2">"#,
    );
    let metadata = format!(
        concat!(
            r#"<metadata xmlns="{}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">"#,
            r#"<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" cellMeta="1"/></metadataTypes>"#,
            r#"<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}">"#,
            r#"<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>"#,
            r#"<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>"#,
        ),
        MAIN
    );
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    parts.push((METADATA.into(), metadata.as_bytes().to_vec()));
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "B1", "State");
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(part(&parts, sheet)).unwrap();
    assert!(
        xml.contains(r#"<c r="D2" cm="1"><f t="array" ref="D2" aca="1">SUM(Table1[State])</f><v>0.00</v></c>"#),
        "{xml}"
    );
    assert_eq!(part(&parts, METADATA), metadata.as_bytes());
    let reopened = assert_saved_headers(&saved, ["Name", "State", "Rev"]);
    assert_eq!(
        reopened.model().sheets[0].array_formula(at("D2")),
        Some(CellRange::new(at("D2"), at("D2")))
    );
}

#[test]
fn a_formula_in_alternate_content_refuses_the_rename() {
    let sheet = "xl/worksheets/sheet1.xml";
    let source = replace_part(
        &book(),
        sheet,
        r#"<c r="E2">"#,
        concat!(
            r#"<c r="D2"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:extension">"#,
            r#"<mc:Choice Requires="x"><x:f>SUM(Table1[Status])</x:f></mc:Choice>"#,
            r#"<mc:Fallback><f>SUM(Table1[Status])</f><v>0</v></mc:Fallback>"#,
            r#"</mc:AlternateContent></c><c r="E2">"#,
        ),
    );
    let mut workbook = Workbook::open(&source).unwrap();
    assert_refused_header(&mut workbook, "B1", "State", sheet);
}

#[test]
fn a_reference_in_a_modeled_charts_extension_refuses_the_rename() {
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet1.xml",
        "<tableParts ",
        r#"<drawing r:id="rIdDrawing"/><tableParts "#,
    );
    let source = replace_part(
        &source,
        "xl/worksheets/_rels/sheet1.xml.rels",
        "</Relationships>",
        &format!(r#"<Relationship Id="rIdDrawing" Type="{DOC_RELS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#),
    );
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    let chart = "xl/charts/chart1.xml";
    parts.extend([
        (
            "xl/drawings/drawing1.xml".into(),
            concat!(
                r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing""#,
                r#" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart""#,
                r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                "<xdr:twoCellAnchor>",
                "<xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
                "<xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
                r#"<xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame>"#,
                "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>",
            )
            .as_bytes()
            .to_vec(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels".into(),
            relationships(&[("rIdChart", "chart", "../charts/chart1.xml")]).into_bytes(),
        ),
        (
            chart.into(),
            concat!(
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">"#,
                r#"<c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>"#,
                "<c:val><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f></c:numRef></c:val>",
                r#"<c:extLst><c:ext uri="dataLabelsRange"><c15:datalabelsRange><c15:f>Table1[Status]</c15:f></c15:datalabelsRange></c:ext></c:extLst>"#,
                "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>",
            )
            .as_bytes()
            .to_vec(),
        ),
    ]);
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&source).unwrap();
    let charts = &workbook.model().sheets[0].charts;
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].refs.len(), 1);
    assert_eq!(charts[0].refs[0].formula, "Sheet1!$C$2:$C$4");
    assert_refused_header(&mut workbook, "B1", "State", chart);
}

#[test]
fn a_reference_in_a_data_validation_refuses_the_rename() {
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet2.xml",
        "<tableParts ",
        concat!(
            r#"<dataValidations count="1"><dataValidation type="list" sqref="D1">"#,
            r#"<formula1>Table1[Status]</formula1></dataValidation></dataValidations><tableParts "#,
        ),
    );
    let mut workbook = Workbook::open(&source).unwrap();
    let error = workbook
        .edit_cell(SHEET1, at("B1"), "State", CalculationOptions::default())
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Operation(OpError::TableHeaderUnsupported(reason))
            if reason.contains("xl/worksheets/sheet2.xml")
    ));
}

#[test]
fn a_local_conditional_format_reference_blocks_a_rename_on_overlap() {
    let source = replace_part(
        &book(),
        "xl/worksheets/sheet1.xml",
        "<tableParts ",
        concat!(
            r#"<conditionalFormatting sqref="D2:E2"><cfRule type="expression" priority="1">"#,
            "<formula>[@Val]&gt;0</formula></cfRule></conditionalFormatting><tableParts ",
        ),
    );
    let mut workbook = Workbook::open(&source).unwrap();
    assert_refused_header(&mut workbook, "F1", "Value", "conditional formats");

    let source = replace_part(&source, "xl/worksheets/sheet1.xml", "D2:E2", "D2:D3");
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "F1", "Value");
    let saved = workbook.save().unwrap();
    assert_eq!(
        assert_table_rule(&saved).table("Table2").unwrap().columns[1],
        "Value"
    );
}
