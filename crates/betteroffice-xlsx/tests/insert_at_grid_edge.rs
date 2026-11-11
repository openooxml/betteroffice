//! Formatting that reaches the last column or row does not block an insert: the part pushed past
//! the grid is dropped, as Excel drops it. Content there still refuses.

use betteroffice_xlsx::{CalculationOptions, CellRange, CellRef, Op, SheetId, Workbook};
use xlsx_model::{
    Cell, CellValue, ColStyle, Hyperlink, MAX_COLS, MAX_ROWS, Sheet, Workbook as Model,
};

const REFUSAL_COLUMNS: &str = "column insertion would discard content at the sheet boundary";
const REFUSAL_ROWS: &str = "row insertion would discard content at the sheet boundary";

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

fn model(edit: impl FnOnce(&mut Sheet)) -> Model {
    let mut model = Model::default();
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    edit(&mut sheet);
    model.sheets.push(sheet);
    model
}

/// A run over columns E to XFD, as a `<col min="5" max="16384">` reads.
fn run_to_last_column(sheet: &mut Sheet, width: f64, hidden: bool, styled: bool) {
    for col in 4..MAX_COLS {
        if hidden {
            sheet.col_widths.insert(col, 0.0);
            sheet.hidden_col_widths.insert(col, width);
        } else {
            sheet.col_widths.insert(col, width);
        }
    }
    if styled {
        sheet.col_styles.push(ColStyle {
            first: 4,
            last: MAX_COLS - 1,
            xf: 0,
        });
    }
}

fn insert_cols() -> Vec<Op> {
    vec![Op::InsertCols {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    }]
}

fn insert_rows() -> Vec<Op> {
    vec![Op::InsertRows {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    }]
}

#[test]
fn formatting_to_the_last_column_does_not_block_a_column_insert() {
    for (label, width, hidden, styled) in [
        ("width and style", 9.0, false, true),
        ("hidden unused columns", 9.0, true, false),
        ("custom width", 21.5, false, false),
    ] {
        let source = model(|sheet| run_to_last_column(sheet, width, hidden, styled));
        let mut workbook = Workbook::from_model(source).unwrap();
        let before = workbook.save().unwrap();
        workbook
            .apply_ops(insert_cols(), options())
            .unwrap_or_else(|error| panic!("{label}: {error}"));
        let sheet = &workbook.model().sheets[0];
        assert!(
            sheet
                .col_widths
                .keys()
                .chain(sheet.hidden_col_widths.keys())
                .all(|&col| col < MAX_COLS),
            "{label}"
        );
        assert_eq!(sheet.col_widths.get(&4), None, "{label}");
        let shifted = if hidden { 0.0 } else { width };
        assert_eq!(sheet.col_widths.get(&5), Some(&shifted), "{label}");
        assert_eq!(
            sheet.col_widths.get(&(MAX_COLS - 1)),
            Some(&shifted),
            "{label}"
        );
        let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
        let sheet = &reopened.model().sheets[0];
        assert_eq!(sheet.col_widths.get(&4), None, "{label}");
        assert_eq!(sheet.col_widths.get(&5), Some(&shifted), "{label}");
        assert_eq!(
            sheet.col_widths.get(&(MAX_COLS - 1)),
            Some(&shifted),
            "{label}"
        );
        workbook.undo(options()).unwrap();
        assert_eq!(workbook.save().unwrap(), before, "{label}");
    }
}

#[test]
fn a_height_on_the_last_row_does_not_block_a_row_insert() {
    let source = model(|sheet| {
        sheet.row_heights.insert(MAX_ROWS - 1, 30.0);
        sheet.row_heights.insert(5, 18.0);
    });
    let mut workbook = Workbook::from_model(source).unwrap();
    let before = workbook.save().unwrap();
    workbook.apply_ops(insert_rows(), options()).unwrap();
    let sheet = &workbook.model().sheets[0];
    assert!(sheet.row_heights.keys().all(|&row| row < MAX_ROWS));
    assert_eq!(sheet.row_heights.get(&6), Some(&18.0));
    workbook.undo(options()).unwrap();
    assert_eq!(workbook.save().unwrap(), before);
}

#[test]
fn content_at_the_last_column_or_row_still_refuses() {
    let last_col = CellRef::new(2, MAX_COLS - 1);
    let cases: Vec<(&str, Model, Vec<Op>, &str)> = vec![
        (
            "cell in the last column",
            model(|sheet| {
                sheet.set_cell(
                    last_col,
                    Cell {
                        value: CellValue::Number { value: 2.0 },
                        ..Cell::default()
                    },
                )
            }),
            insert_cols(),
            REFUSAL_COLUMNS,
        ),
        (
            "merge to the last column",
            model(|sheet| {
                sheet
                    .merges
                    .push(CellRange::new(CellRef::new(2, MAX_COLS - 3), last_col))
            }),
            insert_cols(),
            REFUSAL_COLUMNS,
        ),
        (
            "hyperlink in the last column",
            model(|sheet| {
                sheet.hyperlinks.push(Hyperlink {
                    range: CellRange::new(last_col, last_col),
                    external_target: Some("https://example.com".into()),
                    location: None,
                    tooltip: None,
                    display: None,
                })
            }),
            insert_cols(),
            REFUSAL_COLUMNS,
        ),
        (
            "cell in the last row",
            model(|sheet| {
                sheet.set_cell(
                    CellRef::new(MAX_ROWS - 1, 0),
                    Cell {
                        value: CellValue::Number { value: 3.0 },
                        ..Cell::default()
                    },
                )
            }),
            insert_rows(),
            REFUSAL_ROWS,
        ),
    ];
    for (label, source, ops, message) in cases {
        let mut workbook = Workbook::from_model(source).unwrap();
        let before = workbook.save().unwrap();
        let error = workbook.apply_ops(ops, options()).unwrap_err();
        assert_eq!(error.to_string(), message, "{label}");
        assert_eq!(workbook.save().unwrap(), before, "{label}");
    }
}

#[test]
fn a_whole_sheet_run_stays_one_run_across_a_column_insert() {
    let source = include_bytes!("../../../packages/xlsx/test-fixtures/whole-sheet-columns.xlsx");
    let mut workbook = Workbook::open(source).unwrap();
    let before = workbook.model().clone();
    workbook
        .apply_ops(
            vec![Op::InsertCols {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            }],
            options(),
        )
        .unwrap();
    let sheet = &workbook.model().sheets[0];
    assert_eq!(sheet.col_widths.len() as u32, MAX_COLS);
    assert!(sheet.col_widths.values().all(|&width| width == 12.0));
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let (_, xml) = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    let xml = String::from_utf8_lossy(xml);
    let mut next = 1u32;
    let mut runs = 0;
    for tag in xml.split("<col ").skip(1) {
        let attr = |key: &str| {
            tag.split(&format!("{key}=\""))
                .nth(1)
                .and_then(|rest| rest.split('"').next())
                .unwrap()
                .to_owned()
        };
        assert_eq!(attr("min").parse::<u32>().unwrap(), next, "{xml}");
        assert_eq!(attr("width"), "12", "{xml}");
        next = attr("max").parse::<u32>().unwrap() + 1;
        runs += 1;
    }
    assert_eq!(next, MAX_COLS + 1, "{xml}");
    assert!(runs <= 3, "{xml}");
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.model().sheets[0].col_widths.len() as u32, MAX_COLS);
    workbook.undo(options()).unwrap();
    assert_eq!(workbook.model(), &before);
    let undone = workbook.save().unwrap();
    assert_eq!(sheet_cols(&undone), sheet_cols(source));
}

/// The `<col>` elements of the first worksheet, as written.
fn sheet_cols(bytes: &[u8]) -> Vec<String> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, xml) = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    String::from_utf8_lossy(xml)
        .split("<col ")
        .skip(1)
        .map(|tag| format!("<col {}", &tag[..=tag.find("/>").unwrap() + 1]))
        .collect()
}

#[test]
fn an_edit_and_undo_saves_a_rewritten_file_with_the_same_columns() {
    let source = include_bytes!("../../../packages/xlsx/test-fixtures/whole-sheet-columns.xlsx");
    let mut workbook = Workbook::open(source).unwrap();
    assert_eq!(workbook.save().unwrap(), source);
    workbook
        .edit_cell(SheetId(0), CellRef::new(0, 0), "5", options())
        .unwrap();
    workbook.undo(options()).unwrap();
    let saved = workbook.save().unwrap();
    assert_ne!(saved, source);
    assert_eq!(sheet_cols(&saved), sheet_cols(source));
}

#[test]
fn a_column_inserted_inside_a_hidden_run_is_hidden() {
    let source = model(|sheet| {
        for col in 3..40 {
            sheet.col_widths.insert(col, 0.0);
            sheet.hidden_col_widths.insert(col, 9.0);
        }
    });
    let mut workbook = Workbook::from_model(source).unwrap();
    let before = workbook.save().unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertCols {
                sheet: SheetId(0),
                at: 10,
                count: 2,
            }],
            options(),
        )
        .unwrap();
    let sheet = &workbook.model().sheets[0];
    for col in 10..12 {
        assert_eq!(sheet.col_widths.get(&col), Some(&0.0), "col {col}");
        assert_eq!(sheet.hidden_col_widths.get(&col), Some(&9.0), "col {col}");
    }
    assert_eq!(sheet.col_widths.len(), 39);
    assert_eq!(sheet.col_widths.keys().last(), Some(&41));
    for (at, expected) in [(3, None), (40, None)] {
        let mut edge = Workbook::from_model(model(|sheet| {
            for col in 3..40 {
                sheet.col_widths.insert(col, 0.0);
                sheet.hidden_col_widths.insert(col, 9.0);
            }
        }))
        .unwrap();
        edge.apply_ops(
            vec![Op::InsertCols {
                sheet: SheetId(0),
                at,
                count: 1,
            }],
            options(),
        )
        .unwrap();
        assert_eq!(
            edge.model().sheets[0].col_widths.get(&at),
            expected,
            "edge at {at}"
        );
    }
    workbook.undo(options()).unwrap();
    assert_eq!(workbook.save().unwrap(), before);
}

#[test]
fn a_hidden_run_with_different_remembered_widths_replays_to_the_same_widths() {
    let source = || {
        model(|sheet| {
            for (col, remembered) in [(3, 9.0), (4, 10.0)] {
                sheet.col_widths.insert(col, 0.0);
                sheet.hidden_col_widths.insert(col, remembered);
            }
        })
    };
    let insert = || {
        vec![Op::InsertCols {
            sheet: SheetId(0),
            at: 4,
            count: 1,
        }]
    };
    for observed in [false, true] {
        let mut workbook = Workbook::from_model(source()).unwrap();
        let before = workbook.model().clone();
        let _subscription = observed.then(|| workbook.observe_update_v1(|_| {}).unwrap());
        workbook.apply_ops(insert(), options()).unwrap();
        let sheet = &workbook.model().sheets[0];
        assert_eq!(sheet.col_widths.get(&4), Some(&0.0), "observed={observed}");
        assert_eq!(
            sheet.hidden_col_widths.get(&4),
            Some(&9.0),
            "observed={observed}"
        );
        assert_eq!(
            sheet.hidden_col_widths.get(&5),
            Some(&10.0),
            "observed={observed}"
        );
        let checkpoint = workbook
            .recovery_checkpoint(0, None)
            .unwrap_or_else(|error| panic!("observed={observed}: {error}"));
        let (restored, _, _, _) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
        let restored = &restored.model().sheets[0];
        assert_eq!(restored.col_widths, sheet.col_widths, "observed={observed}");
        assert_eq!(
            restored.hidden_col_widths, sheet.hidden_col_widths,
            "observed={observed}"
        );
        workbook.undo(options()).unwrap();
        assert_eq!(workbook.model(), &before, "observed={observed}");
    }
}
