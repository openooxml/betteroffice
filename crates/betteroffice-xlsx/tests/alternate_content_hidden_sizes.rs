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

#[test]
fn collaborative_hidden_size_only_save_keeps_alternate_content() {
    use betteroffice_xlsx::Op;

    let model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..Default::default()
    };
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let path = "xl/worksheets/sheet1.xml";
    let sheet = parts.iter_mut().find(|(name, _)| name == path).unwrap();
    let wrapper = concat!(
        r#"<mc:AlternateContent><mc:Choice Requires="feature"><c r="A1"><v>7</v></c></mc:Choice>"#,
        r#"<mc:Fallback><c r="A1"><v>7</v></c></mc:Fallback></mc:AlternateContent>"#,
    );
    sheet.1 = format!(
        concat!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:feature="urn:feature">"#,
            r#"<sheetData><row r="1" ht="25" customHeight="1" hidden="1">{}</row></sheetData></worksheet>"#,
        ),
        wrapper,
    )
    .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let sheet = SheetId(0);
    let options = CalculationOptions::default();
    for (axes, expected) in [(true, 40.0), (false, 25.0)] {
        let mut workbook = Workbook::open_collaborative(&source, 7).unwrap();
        if !axes {
            workbook
                .edit_cell(sheet, CellRef::new(0, 1), "3", options)
                .unwrap();
            workbook.undo(options).unwrap();
        }
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
        let (_, xml) = parts.iter().find(|(name, _)| name == path).unwrap();
        assert!(std::str::from_utf8(xml).unwrap().contains(wrapper));
        let reopened = Workbook::open(&saved).unwrap();
        let row = reopened.sheet(sheet).unwrap();
        assert!(row.row_hidden(0));
        assert_eq!(row.hidden_row_heights.get(&0), Some(&expected));
    }
}
