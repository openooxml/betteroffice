use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRange, CellRef, CellValue, DefinedName, Error, Hyperlink, Op,
    Sheet, SheetId, Workbook, WorkbookModel,
};
use xlsx_calc::parse_formula;
use xlsx_model::Table;
use xlsx_ops::OpError;

const FORMULA_CELL: CellRef = CellRef {
    row: 0,
    col: 1,
    abs_row: false,
    abs_col: false,
};
const CACHE: CellValue = CellValue::Number { value: 42.0 };

fn model_with_cached_formula(sheets: &[&str], source: &str) -> WorkbookModel {
    let mut model = WorkbookModel {
        sheets: sheets.iter().copied().map(Sheet::new).collect(),
        ..WorkbookModel::default()
    };
    for sheet in &mut model.sheets {
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
    }
    model.sheets.last_mut().unwrap().set_cell(
        FORMULA_CELL,
        Cell {
            formula: Some(source.to_owned()),
            value: CACHE,
            style: None,
        },
    );
    model
}

fn assert_formula_and_cache(workbook: &Workbook, source: &str) {
    let cell = workbook
        .model()
        .sheets
        .last()
        .unwrap()
        .cell(FORMULA_CELL)
        .unwrap();
    assert_eq!(cell.formula.as_deref(), Some(source));
    assert_eq!(cell.value, CACHE, "{source}");
}

fn insert_row(sheet: u32) -> Op {
    Op::InsertRows {
        sheet: SheetId(sheet),
        at: 0,
        count: 1,
    }
}

fn assert_structural_edit_refused(workbook: &mut Workbook, op: Op, source: &str) {
    let before = workbook.model().clone();
    let report = SheetId(before.sheets.len() as u32 - 1);
    let result = workbook.apply_ops(vec![op], CalculationOptions::default());
    assert!(
        matches!(
            result,
            Err(Error::Operation(OpError::FormulaNotRewritable { sheet, cell }))
                if sheet == report && cell == FORMULA_CELL
        ),
        "{source}"
    );
    assert_eq!(workbook.model(), &before, "{source}");
    assert_formula_and_cache(workbook, source);
}

#[test]
fn mixed_quoted_whole_row_spans_keep_excel_caches() {
    let sheets = [
        "Sheet1", "Sheet2", "Sheet 3", "Q1_Sales", "Q2 Sales", "Report",
    ];
    for source in [
        "SUM(Sheet1:'Sheet 3'!1:2)",
        "SUM(Q1_Sales:'Q2 Sales'!$1:$5)",
        "Sheet1:'Sheet 3'!1:1",
    ] {
        let mut workbook =
            Workbook::from_model(model_with_cached_formula(&sheets, source)).unwrap();
        workbook.recalculate_all(CalculationOptions::default());
        assert_formula_and_cache(&workbook, source);
    }
}

#[test]
fn removing_a_sheet_and_its_same_named_table_narrows_the_span() {
    for (target, expected) in [
        ("B5", "SUM('Data:Data two'!B5)"),
        ("#REF!", "SUM('Data:Data two'!#REF!)"),
    ] {
        let source = format!("SUM(Amounts:'Data two'!{target})");
        let mut model =
            model_with_cached_formula(&["Amounts", "Data", "Data two", "Report"], &source);
        model.tables.push(Table {
            name: "Amounts".into(),
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A2").unwrap(),
            header_rows: 1,
            totals_rows: 0,
            columns: vec!["Amt".into()],
            style: Vec::new(),
        });
        let mut workbook = Workbook::from_model(model).unwrap();
        workbook
            .apply_ops(
                vec![Op::RemoveSheet { index: 0 }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(workbook.model().tables.is_empty());
        assert_formula_and_cache(&workbook, expected);
    }
}

#[test]
fn row_edits_refuse_span_formulas_before_mutating_the_workbook() {
    for (source, op) in [
        ("SUM(Jan:Mar!A1)", insert_row(1)),
        (
            "SUM(Jan:Mar!A1)",
            Op::DeleteRows {
                sheet: SheetId(1),
                at: 0,
                count: 1,
            },
        ),
        ("SUM(Jan:Mar!A1)+C1", insert_row(7)),
        ("SUM(Jan:Mar!#REF!)+Data!A1", insert_row(3)),
        ("SUM(Sheet1:Sheet3!1:2)", insert_row(5)),
    ] {
        let model = model_with_cached_formula(
            &[
                "Jan", "Feb", "Mar", "Data", "Sheet1", "Sheet2", "Sheet3", "Report",
            ],
            source,
        );
        let mut workbook = Workbook::from_model(model).unwrap();
        assert_structural_edit_refused(&mut workbook, op, source);
    }
}

#[test]
fn all_span_quoting_and_target_combinations_keep_caches_and_refuse_inserts() {
    let sheets = [
        "Sheet1", "Q1 Sales", "Sheet2", "Sheet3", "Q3 Sales", "Report",
    ];
    for qualifier in [
        "Sheet1:Sheet3",
        "'Sheet1:Sheet3'",
        "Sheet1:'Q3 Sales'",
        "'Q1 Sales':Sheet3",
        "'Q1 Sales':'Q3 Sales'",
    ] {
        for target in ["$A$1", "A1:B2", "$1:$5", "$A:$C", "#REF!"] {
            let source = format!("SUM({qualifier}!{target})+C1");
            let error = parse_formula(&source).unwrap_err();
            assert_eq!(error.pos, 4, "{source}");
            assert_eq!(error.message, "3-D references are unsupported", "{source}");

            let model = model_with_cached_formula(&sheets, &source);
            let mut workbook = Workbook::from_model(model).unwrap();
            workbook.recalculate_all(CalculationOptions::default());
            assert_formula_and_cache(&workbook, &source);
            assert_structural_edit_refused(&mut workbook, insert_row(2), &source);
        }
    }
}

#[test]
fn unspaced_defined_name_joins_evaluate_and_keep_the_structural_edit_guard() {
    for sheet in ["R2C3", "A1", "TRUE"] {
        let source = format!("SUM(Start_1:{sheet}!A2)");
        let mut model = model_with_cached_formula(&[sheet, "Report"], &source);
        model.defined_names.push(DefinedName {
            name: "Start_1".into(),
            formula: format!("'{sheet}'!A1"),
            local_sheet: None,
            hidden: false,
        });
        let mut workbook = Workbook::from_model(model.clone()).unwrap();
        workbook.recalculate_all(CalculationOptions::default());
        let stored = workbook.model().sheets[1].cell(FORMULA_CELL).unwrap();
        assert_eq!(stored.formula.as_deref(), Some(source.as_str()));
        assert_eq!(stored.value, CellValue::Number { value: 1.0 }, "{source}");

        let mut uncalculated = Workbook::from_model(model).unwrap();
        assert_structural_edit_refused(&mut uncalculated, insert_row(0), &source);
    }
}

#[test]
fn a_scoped_name_takes_precedence_over_a_sheet_span_only_on_its_sheet() {
    let source = "SUM(Start_1:Data!B2)";
    let mut model = model_with_cached_formula(&["Start_1", "Data", "Report"], source);
    model.defined_names.push(DefinedName {
        name: "Start_1".into(),
        formula: "Data!A1".into(),
        local_sheet: Some(SheetId(2)),
        hidden: false,
    });
    let other_formula = CellRef::new(0, 2);
    let data = &mut model.sheets[1];
    for (at, value) in [(CellRef::new(0, 0), 10.0), (CellRef::new(1, 1), 20.0)] {
        data.set_cell(
            at,
            Cell {
                value: CellValue::Number { value },
                ..Cell::default()
            },
        );
    }
    data.set_cell(
        other_formula,
        Cell {
            formula: Some(source.into()),
            value: CACHE,
            ..Cell::default()
        },
    );
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(
        workbook.model().sheets[1].cell(other_formula).unwrap().value,
        CACHE
    );
    assert_eq!(
        workbook.model().sheets[2].cell(FORMULA_CELL).unwrap().value,
        CellValue::Number { value: 30.0 }
    );
}

#[test]
fn sheet_qualifier_edits_follow_a_span_with_a_missing_quoted_endpoint() {
    let model = model_with_cached_formula(&["Jan", "Report"], "SUM(Jan:'Missing Sheet'!A1)");
    let mut workbook = Workbook::from_model(model).unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "NewJan".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_formula_and_cache(&workbook, "SUM('NewJan:Missing Sheet'!A1)");

    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    let cell = workbook.model().sheets[0].cell(FORMULA_CELL).unwrap();
    assert_eq!(cell.formula.as_deref(), Some("SUM(#REF!)"));
}

#[test]
fn sheet_rename_updates_hyperlinks_with_ambiguous_targets() {
    for (old, location, expected) in [
        ("Data", "#Data!A1 [Col]", "#NewData!A1 [Col]"),
        ("Jan", "#Jan:Mar!A:B:C", "#NewData:Mar!A:B:C"),
    ] {
        let mut report = Sheet::new("Report");
        report.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1").unwrap(),
            location: Some(location.to_owned()),
            external_target: None,
            tooltip: None,
            display: None,
        });
        let mut workbook = Workbook::from_model(WorkbookModel {
            sheets: vec![Sheet::new(old), Sheet::new("Mar"), report],
            ..WorkbookModel::default()
        })
        .unwrap();

        workbook
            .apply_ops(
                vec![Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "NewData".into(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(
            workbook.model().sheets[2].hyperlinks[0].location.as_deref(),
            Some(expected)
        );
    }
}
