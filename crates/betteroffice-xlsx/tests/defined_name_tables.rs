use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, Sheet, SheetId, Workbook, WorkbookModel,
};
use xlsx_model::CellProvider;

const NAME_COUNT: usize = 66_391;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn names() -> String {
    (0..NAME_COUNT)
        .map(|index| match index {
            40_000 => r#"<definedName name="Rate">Data!$A$1</definedName>"#.to_owned(),
            65_999 => {
                r#"<definedName name="_Tail" hidden="1">Data!$A$1*3</definedName>"#.to_owned()
            }
            _ if index % 2 == 0 => {
                format!(r#"<definedName name="_junk{index}" hidden="1">#REF!</definedName>"#)
            }
            _ => format!(
                r#"<definedName name="Area{index}" localSheetId="0" comment="a &amp; b">Data!$B${}</definedName>"#,
                index % 1000 + 1
            ),
        })
        .collect()
}

fn package(names: &str) -> Vec<u8> {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: number(5.0),
            ..Cell::default()
        },
    );
    for (address, formula) in [("C1", "Rate*2"), ("C2", "_Tail+1")] {
        sheet.set_cell(
            cell(address),
            Cell {
                value: number(0.0),
                formula: Some(formula.into()),
                style: None,
            },
        );
    }
    model.sheets.push(sheet);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = &mut parts
        .iter_mut()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap()
        .1;
    *workbook = std::str::from_utf8(workbook)
        .unwrap()
        .replace(
            "</workbook>",
            &format!("<definedNames>{names}</definedNames></workbook>"),
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_workbook_xml(saved: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let (_, bytes) = parts
        .iter()
        .find(|(path, _)| path == "xl/workbook.xml")
        .unwrap();
    String::from_utf8(bytes.clone()).unwrap()
}

#[test]
fn a_name_table_past_the_old_count_cap_opens_evaluates_and_saves_verbatim() {
    let names = names();
    let mut workbook = Workbook::open(&package(&names)).unwrap();
    assert_eq!(workbook.model().defined_names.len(), NAME_COUNT);

    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(workbook.model().value(SheetId(0), cell("C1")), number(10.0));
    assert_eq!(workbook.model().value(SheetId(0), cell("C2")), number(16.0));

    workbook
        .edit_cell(SheetId(0), cell("A1"), "7", CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.model().value(SheetId(0), cell("C1")), number(14.0));
    assert_eq!(workbook.model().value(SheetId(0), cell("C2")), number(22.0));

    let saved = workbook.save().unwrap();
    let container = format!("<definedNames>{names}</definedNames>");
    assert!(saved_workbook_xml(&saved).contains(&container));
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().defined_names,
        workbook.model().defined_names
    );
}

#[test]
fn a_model_whose_name_table_is_over_the_byte_bound_is_refused() {
    let formula = "1".repeat(xlsx_calc::lexer::MAX_FORMULA_BYTES);
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    let count = xlsx_model::workbook::MAX_DEFINED_NAME_BYTES / formula.len() + 1;
    model.defined_names = (0..count)
        .map(|index| xlsx_model::DefinedName {
            name: format!("N_{index}"),
            formula: formula.clone(),
            local_sheet: None,
            hidden: false,
        })
        .collect();
    let failure = match Workbook::from_model(model) {
        Ok(_) => panic!("an over-bound name table must be refused"),
        Err(failure) => failure.to_string(),
    };
    assert!(failure.contains("defined name table"), "{failure}");
}

#[test]
fn undoing_a_sheet_rename_restores_the_name_table_exactly() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.defined_names = [("Ref", "Data!$A$1"), ("Other", "D!$A$1")]
        .into_iter()
        .map(|(name, formula)| xlsx_model::DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet: None,
            hidden: false,
        })
        .collect();
    let original = model.defined_names.clone();
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook
        .apply_ops(
            vec![xlsx_ops::Op::RenameSheet {
                sheet: SheetId(0),
                name: "D".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let renamed = workbook.model().defined_names.clone();
    assert_eq!(renamed[0].formula, "D!$A$1");
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, original);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, renamed);
}
