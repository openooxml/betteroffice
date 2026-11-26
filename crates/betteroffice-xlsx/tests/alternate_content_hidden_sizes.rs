use betteroffice_xlsx::{CalculationOptions, CellRef, Op, Sheet, SheetId, Workbook, WorkbookModel};

const SHEET_PATH: &str = "xl/worksheets/sheet1.xml";
const WRAPPER: &str = concat!(
    r#"<mc:AlternateContent><mc:Choice Requires="feature"><c r="A1"><v>7</v></c></mc:Choice>"#,
    r#"<mc:Fallback><c r="A1"><v>7</v></c></mc:Fallback></mc:AlternateContent>"#,
);

#[test]
fn collaborative_save_keeps_hidden_sizes_around_a_wrapped_cell() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:feature="urn:feature">"#,
        r#"<cols><col min="1" max="1" width="18" customWidth="1" hidden="1"/></cols>"#,
        r#"<sheetData><row r="1" ht="25" customHeight="1" hidden="1"><mc:AlternateContent>"#,
        r#"<mc:Choice Requires="feature"><c r="A1"><v>7</v></c></mc:Choice>"#,
        r#"<mc:Fallback><c r="A1"><v>7</v></c></mc:Fallback></mc:AlternateContent>"#,
        r#"<c r="B1"><v>2</v></c></row></sheetData></worksheet>"#,
    )
    .as_bytes()
    .to_vec();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();

    let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 1),
            "3",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();

    let reopened = Workbook::open(&saved).unwrap();
    let sheet = reopened.sheet(SheetId(0)).unwrap();
    assert_eq!(
        (
            sheet.row_hidden(0),
            sheet.hidden_row_heights.get(&0).copied(),
            sheet.col_hidden(0),
            sheet.hidden_col_widths.get(&0).copied(),
        ),
        (true, Some(25.0), true, Some(18.0)),
        "saved row 1 and column A should be hidden and keep their authored sizes",
    );
}

#[test]
fn collaborative_hidden_size_only_save_keeps_alternate_content() {
    let workbook = Workbook::open_collaborative(&wrapped_row_source(), 7).unwrap();
    assert_hidden_size_only_save(workbook, 40.0);
}

#[test]
fn hidden_size_only_save_without_axes_keeps_the_source_size_until_axes_survive_history() {
    let mut workbook = Workbook::open_collaborative(&wrapped_row_source(), 7).unwrap();
    let options = CalculationOptions::default();
    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 1), "3", options)
        .unwrap();
    workbook.undo(options).unwrap();
    assert_hidden_size_only_save(workbook, 25.0);
}

fn wrapped_row_source() -> Vec<u8> {
    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..Default::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(name, _)| name == SHEET_PATH)
        .unwrap();
    sheet.1 = format!(
        concat!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:feature="urn:feature">"#,
            r#"<sheetData><row r="1" ht="25" customHeight="1" hidden="1">{}</row></sheetData></worksheet>"#,
        ),
        WRAPPER,
    )
    .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_hidden_size_only_save(mut workbook: Workbook, expected: f64) {
    let sheet = SheetId(0);
    let options = CalculationOptions::default();
    workbook
        .apply_ops(
            vec![
                Op::SetRowHeight {
                    sheet,
                    row: 0,
                    height: Some(40.0),
                },
                Op::SetRowVisibility {
                    sheet,
                    row: 0,
                    hidden: true,
                },
            ],
            options,
        )
        .unwrap();
    assert_eq!(
        workbook.sheet(sheet).unwrap().hidden_row_heights.get(&0),
        Some(&40.0),
    );
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let (_, xml) = parts.iter().find(|(name, _)| name == SHEET_PATH).unwrap();
    assert!(std::str::from_utf8(xml).unwrap().contains(WRAPPER));
    let reopened = Workbook::open(&saved).unwrap();
    let row = reopened.sheet(sheet).unwrap();
    assert!(row.row_hidden(0));
    assert_eq!(row.hidden_row_heights.get(&0), Some(&expected));
}
