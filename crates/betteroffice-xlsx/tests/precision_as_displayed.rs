use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellValue, NumberFormatMutation, Sheet, SheetId,
    Workbook, WorkbookModel,
};
use xlsx_model::{CellProvider, Xf};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn source(attribute: &str) -> Vec<u8> {
    let mut sheet = Sheet::new("Data");
    for (col, formula, value, style) in [
        (0, None, 6.125, Some(1)),
        (1, Some("A1"), 6.125, Some(1)),
        (2, Some("B1*10"), 61.25, None),
        (3, Some("UNSUPPORTED()"), 6.125, Some(1)),
        (4, Some("6.125"), 6.125, Some(1)),
    ] {
        sheet.set_cell(
            CellRef::new(0, col),
            Cell {
                value: number(value),
                formula: formula.map(str::to_owned),
                style,
            },
        );
    }
    let mut model = WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    };
    model.styles.num_fmts.push((164, "0.0".into()));
    model.styles.cell_xfs = vec![
        Xf::default(),
        Xf {
            num_fmt_id: Some(164),
            ..Xf::default()
        },
    ];
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook = &mut parts
        .iter_mut()
        .find(|(name, _)| name == "xl/workbook.xml")
        .unwrap()
        .1;
    *workbook = String::from_utf8(workbook.clone()).unwrap().replace(
        "</workbook>",
        &format!(r#"<calcPr calcId="191029"{attribute} custom="kept"/><extLst><ext uri="unknown"/></extLst></workbook>"#),
    ).into_bytes();
    parts.push((
        "custom/unknown.xml".into(),
        b"<unknown kept='yes'/>".to_vec(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_value(workbook: &Workbook, col: u32, expected: f64) {
    assert_eq!(
        workbook.model().value(SheetId(0), CellRef::new(0, col)),
        number(expected)
    );
}

#[test]
fn edits_and_undo_keep_displayed_precision_and_save_rounded_caches() {
    let bytes = source(r#" fullPrecision="0""#);
    for collaborative in [false, true] {
        let mut workbook = if collaborative {
            Workbook::open_collaborative(&bytes, 123).unwrap()
        } else {
            Workbook::open(&bytes).unwrap()
        };
        assert!(!workbook.model().full_precision);
        assert_value(&workbook, 0, 6.125);
        assert_value(&workbook, 1, 6.125);
        assert_eq!(workbook.save().unwrap(), bytes);
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "6.26",
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(!workbook.model().full_precision);
        assert_value(&workbook, 0, 6.26);
        assert_value(&workbook, 1, 6.3);
        assert_value(&workbook, 2, 63.0);
        assert_value(&workbook, 3, 6.125);
        assert_value(&workbook, 4, 6.125);
        let saved = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
        let parsed = xlsx_parse::parse_workbook(&saved).unwrap();
        assert_eq!(parsed.value(SheetId(0), CellRef::new(0, 1)), number(6.3));
        assert_eq!(parsed.value(SheetId(0), CellRef::new(0, 4)), number(6.125));
        let original = ooxml_opc::unzip_parts(&bytes).unwrap();
        let part = |parts: &[(String, Vec<u8>)], path: &str| {
            parts
                .iter()
                .find(|(name, _)| name == path)
                .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        };
        assert_eq!(
            part(&saved, "custom/unknown.xml"),
            part(&original, "custom/unknown.xml")
        );
        assert_eq!(
            part(&saved, "xl/workbook.xml"),
            part(&original, "xl/workbook.xml").map(|xml| xml.replace(
                r#"custom="kept"/>"#,
                r#"custom="kept" fullCalcOnLoad="1"/>"#
            ))
        );
        workbook.undo(CalculationOptions::default()).unwrap();
        assert!(!workbook.model().full_precision);
        assert_value(&workbook, 0, 6.125);
        assert_value(&workbook, 1, 6.1);
        assert_value(&workbook, 2, 61.0);
    }
}

#[test]
fn absent_or_true_precision_keeps_values_and_saved_file_identical() {
    for attribute in ["", r#" fullPrecision="1""#, r#" fullPrecision="true""#] {
        let bytes = source(attribute);
        let mut workbook = Workbook::open(&bytes).unwrap();
        assert!(workbook.model().full_precision);
        workbook.recalculate_all(CalculationOptions::default());
        for (col, expected) in [(0, 6.125), (1, 6.125), (2, 61.25), (3, 6.125), (4, 6.125)] {
            assert_value(&workbook, col, expected);
        }
        assert_eq!(workbook.save().unwrap(), bytes, "{attribute}");
    }
}

#[test]
fn recalculation_keeps_unsupported_cached_values_and_constants() {
    let bytes = source(r#" fullPrecision="false""#);
    let workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_value(&workbook, 0, 6.125);
    assert_value(&workbook, 1, 6.1);
    assert_value(&workbook, 3, 6.125);
    assert_value(&workbook, 4, 6.1);
}

#[test]
fn recovery_preserves_displayed_precision_after_edits() {
    let bytes = source(r#" fullPrecision="0""#);
    let workbook = Workbook::open(&bytes).unwrap();
    let checkpoint = workbook.recovery_checkpoint(1, None).unwrap();
    let (mut restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert!(!restored.model().full_precision);
    restored
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "6.26",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_value(&restored, 1, 6.3);
    assert_value(&restored, 2, 63.0);
}

#[test]
fn collaboration_rejects_a_base_with_different_calculation_precision() {
    let displayed = source(r#" fullPrecision="0""#);
    let precise = source(r#" fullPrecision="1""#);
    let mut left = Workbook::open_collaborative(&displayed, 123).unwrap();
    let right = Workbook::open_collaborative(&precise, 456).unwrap();
    assert!(
        left.apply_update_v1(
            &right.encode_state_as_update_v1(),
            CalculationOptions::default()
        )
        .is_err()
    );
}

#[test]
fn number_format_changes_and_their_undo_recalculate_displayed_precision() {
    let bytes = source(r#" fullPrecision="0""#);
    for collaborative in [false, true] {
        let mut workbook = if collaborative {
            Workbook::open_collaborative(&bytes, 123).unwrap()
        } else {
            Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap()
        };
        if collaborative {
            workbook.recalculate_all(CalculationOptions::default());
        }
        assert_value(&workbook, 1, 6.1);
        assert_value(&workbook, 2, 61.0);
        workbook
            .set_range_number_format(
                SheetId(0),
                CellRange::new(CellRef::new(0, 1), CellRef::new(0, 1)),
                NumberFormatMutation::Custom {
                    pattern: "0.00".into(),
                },
                CalculationOptions::default(),
            )
            .unwrap();
        assert_value(&workbook, 1, 6.13);
        assert_value(&workbook, 2, 6.13 * 10.0);
        workbook.undo(CalculationOptions::default()).unwrap();
        assert_value(&workbook, 1, 6.1);
        assert_value(&workbook, 2, 61.0);
    }
}
