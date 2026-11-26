use betteroffice_xlsx::{CalculationOptions, CellRef, Sheet, SheetId, Workbook, WorkbookModel};

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
