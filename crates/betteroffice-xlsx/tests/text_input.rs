use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellValue, NumberFormatMutation, Sheet, SheetId,
    Workbook, WorkbookModel,
};

fn workbook() -> Workbook {
    Workbook::from_model(WorkbookModel {
        sheets: vec![Sheet::new("Sheet1")],
        ..Default::default()
    })
    .unwrap()
}

#[test]
fn entry_respects_builtin_and_custom_text_formats() {
    for (format, is_text) in [
        (NumberFormatMutation::Automatic, false),
        (NumberFormatMutation::PlainText, true),
        (
            NumberFormatMutation::Custom {
                pattern: "@".into(),
            },
            true,
        ),
        (
            NumberFormatMutation::Custom {
                pattern: "0.0;-0.0;0.0;@".into(),
            },
            false,
        ),
    ] {
        let mut wb = workbook();
        let at = CellRef::new(0, 0);
        let options = CalculationOptions::default();
        wb.set_range_number_format(SheetId(0), CellRange::new(at, at), format, options)
            .unwrap();
        let style = wb
            .sheet(SheetId(0))
            .unwrap()
            .cell(at)
            .and_then(|cell| cell.style);

        for (input, general_value, formula) in [
            ("1.0", CellValue::Number { value: 1.0 }, None),
            ("001", CellValue::Number { value: 1.0 }, None),
            ("TRUE", CellValue::Bool { value: true }, None),
            ("false", CellValue::Bool { value: false }, None),
            ("=1+1", CellValue::Number { value: 2.0 }, Some("1+1")),
            ("", CellValue::Empty, None),
        ] {
            wb.edit_cell(SheetId(0), at, input, options).unwrap();
            let stored = wb
                .sheet(SheetId(0))
                .unwrap()
                .cell(at)
                .cloned()
                .unwrap_or_default();
            let expected = if is_text && !input.is_empty() {
                CellValue::Text {
                    value: input.into(),
                }
            } else {
                general_value
            };
            assert_eq!(stored.value, expected, "{input:?}, text format: {is_text}");
            assert_eq!(
                stored.formula.as_deref(),
                if is_text { None } else { formula }
            );
            assert_eq!(stored.style, style);
        }

        for (input, expected) in [
            ("'1.0", "1.0"),
            ("'=1+1", "=1+1"),
            ("''hello", "'hello"),
            ("'", ""),
        ] {
            wb.edit_cell(SheetId(0), at, input, options).unwrap();
            let stored = wb.sheet(SheetId(0)).unwrap().cell(at).unwrap();
            assert_eq!(
                stored.value,
                CellValue::Text {
                    value: expected.into()
                }
            );
            assert!(stored.formula.is_none());
        }
    }
}

#[test]
fn formatting_changes_affect_subsequent_entries_only() {
    let mut wb = workbook();
    let at = CellRef::new(0, 0);
    let options = CalculationOptions::default();
    wb.edit_cell(SheetId(0), at, "1.0", options).unwrap();

    for (format, before, after) in [
        (
            NumberFormatMutation::PlainText,
            CellValue::Number { value: 1.0 },
            CellValue::Text {
                value: "1.0".into(),
            },
        ),
        (
            NumberFormatMutation::Automatic,
            CellValue::Text {
                value: "1.0".into(),
            },
            CellValue::Number { value: 1.0 },
        ),
    ] {
        wb.set_range_number_format(SheetId(0), CellRange::new(at, at), format, options)
            .unwrap();
        assert_eq!(
            wb.sheet(SheetId(0)).unwrap().cell(at).unwrap().value,
            before
        );
        wb.edit_cell(SheetId(0), at, "1.0", options).unwrap();
        assert_eq!(wb.sheet(SheetId(0)).unwrap().cell(at).unwrap().value, after);
    }
}
