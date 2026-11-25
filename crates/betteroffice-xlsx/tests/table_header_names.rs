use std::collections::HashSet;

use betteroffice_xlsx::{
    CalculationOptions, CellInput, CellRef, CellValue, EditFailureCode, EditOperation, EditRequest,
    EditStep, Error, RangeAddress, RangeTarget, SheetId, Workbook, WorkbookModel,
};
use xlsx_model::CellProvider;
use xlsx_ops::OpError;

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
    assert_eq!(workbook.model().table("Table1").unwrap().columns[1], "STATUS");
    assert_eq!(
        workbook.model().value(SHEET1, at("B1")),
        CellValue::Text { value: "STATUS".into() },
    );
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let path = "xl/tables/table1.xml";
    assert_eq!(part(&after, path), part(&before, path));
    assert_eq!(
        Workbook::open(&saved).unwrap().model().value(SHEET1, at("B1")),
        CellValue::Text { value: "STATUS".into() },
    );
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
fn a_suffix_over_the_cell_text_limit_is_refused_atomically() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let name = "x".repeat(xlsx_calc::eval::MAX_CELL_TEXT_CHARS);
    edit(&mut workbook, "A1", &name);
    let version = workbook.version();
    let error = workbook
        .edit_cell(SHEET1, at("B1"), &name, CalculationOptions::default())
        .unwrap_err();
    assert!(matches!(error, Error::Operation(OpError::TableHeaderTextTooLong)));
    assert_eq!(workbook.version(), version);
    assert_saved_headers(&workbook.save().unwrap(), [&name, "Status", "Rev"]);
}

#[test]
fn a_formula_in_a_header_is_refused_without_changing_it() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let error = workbook
        .edit_cell(SHEET1, at("B1"), "=1+1", CalculationOptions::default())
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Operation(OpError::TableHeaderFormulaUnsupported)
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
        CellValue::Text { value: "State".into() },
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
        CellValue::Text { value: "Reviewed".into() },
    );
}

#[test]
fn saving_heals_an_existing_header_mismatch_and_preserves_other_tables() {
    let source = replace_part(
        &book(),
        "xl/sharedStrings.xml",
        "<t>Status</t>",
        "<t>State</t>",
    );
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(workbook.model().table("Table1").unwrap().columns[1], "Status");
    edit(&mut workbook, "B3", "Reviewed");
    let saved = workbook.save().unwrap();
    assert_saved_headers(&saved, ["Name", "State", "Rev"]);
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in ["xl/tables/table2.xml", "xl/tables/table3.xml"] {
        assert_eq!(part(&after, path), part(&before, path), "{path}");
    }
}

#[test]
fn editing_the_first_data_row_of_a_headerless_table_preserves_its_part() {
    let path = "xl/tables/table2.xml";
    let source = replace_part(&book(), path, "headerRowCount=\"1\"", "headerRowCount=\"0\"");
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "E1", "7");
    let saved = workbook.save().unwrap();
    let before = ooxml_opc::unzip_parts(&source).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(part(&after, path), part(&before, path));
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.model().value(SHEET1, at("E1")), CellValue::Number { value: 7.0 });
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
    assert_eq!(reopened.model().value(SHEET1, at("B4")), CellValue::Number { value: 7.0 });
}

#[test]
fn a_header_with_xml_special_characters_and_a_line_break_round_trips() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let name = "A<B & \"C\"\nNext";
    edit(&mut workbook, "B1", name);
    let saved = workbook.save().unwrap();
    assert_saved_headers(&saved, ["Name", name, "Rev"]);
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(part(&parts, "xl/tables/table1.xml")).unwrap();
    assert!(xml.contains("_x000A_"), "{xml}");
}

#[test]
fn an_existing_structured_reference_follows_a_header_rename_and_undo() {
    let mut workbook = Workbook::open(&book()).unwrap();
    edit(&mut workbook, "H1", "=COUNTA(Table1[Status])");
    let expected = CellValue::Number { value: 3.0 };
    assert_eq!(workbook.model().value(SHEET1, at("H1")), expected);
    edit(&mut workbook, "B1", "State");
    let formula = |workbook: &Workbook| {
        workbook.model().sheets[0].cell(at("H1")).unwrap().formula.clone()
    };
    assert_eq!(formula(&workbook).as_deref(), Some("COUNTA(Table1[State])"));
    assert_eq!(workbook.model().value(SHEET1, at("H1")), expected);
    let reopened = assert_saved_headers(&workbook.save().unwrap(), ["Name", "State", "Rev"]);
    assert_eq!(formula(&reopened).as_deref(), Some("COUNTA(Table1[State])"));
    assert_eq!(reopened.model().value(SHEET1, at("H1")), expected);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(formula(&workbook).as_deref(), Some("COUNTA(Table1[Status])"));
    assert_eq!(workbook.model().value(SHEET1, at("H1")), expected);
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
        workbook.model().sheets[1].cell(at("D1")).unwrap().formula.as_deref(),
        Some("COUNTA(Table1[[Status]:[Rev]])")
    );
}

#[test]
fn a_peer_receives_the_header_and_its_rewritten_formula() {
    let source = book();
    let mut writer = Workbook::open_collaborative(&source, 77).unwrap();
    let mut peer = Workbook::open_collaborative(&source, 78).unwrap();
    edit(&mut writer, "H1", "=COUNTA(Table1[Status])");
    edit(&mut writer, "B1", "State");
    let update = writer.encode_diff_v1(&peer.encode_state_vector_v1()).unwrap();
    peer.apply_update_v1(&update, CalculationOptions::default()).unwrap();
    assert_eq!(
        peer.model().sheets[0].cell(at("H1")).unwrap().formula.as_deref(),
        Some("COUNTA(Table1[State])")
    );
    assert_eq!(peer.model().value(SHEET1, at("H1")), CellValue::Number { value: 3.0 });
    assert_saved_headers(&peer.save().unwrap(), ["Name", "State", "Rev"]);
}

#[test]
fn a_reference_in_a_preserved_table_formula_refuses_the_rename() {
    let source = replace_part(
        &book(),
        "xl/tables/table1.xml",
        r#"<tableColumn id="3" name="Rev"/>"#,
        r#"<tableColumn id="3" name="Rev"><calculatedColumnFormula>[@Status]</calculatedColumnFormula></tableColumn>"#,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    let error = workbook
        .edit_cell(SHEET1, at("B1"), "State", CalculationOptions::default())
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Operation(OpError::TableHeaderReferenceUnsupported { .. })
    ));
    assert_saved_headers(&workbook.save().unwrap(), ["Name", "Status", "Rev"]);
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
        Error::Operation(OpError::TableHeaderReferenceUnsupported { .. })
    ));
}
