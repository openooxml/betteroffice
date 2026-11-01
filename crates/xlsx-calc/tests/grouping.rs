//! GROUPBY, PIVOTBY and TRIMRANGE: grouping, totals, headers, sorting, filtering,
//! eta-reduced or LAMBDA aggregations and blank-edge trimming, against
//! Microsoft's reference.

use xlsx_calc::{EvalContext, Value, evaluate_array, parse_formula};
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

fn n(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn t(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn put(sheet: &mut Sheet, address: &str, value: CellValue) {
    sheet.set_cell(
        CellRef::parse_a1(address).unwrap(),
        Cell {
            value,
            ..Cell::default()
        },
    );
}

/// evaluate in array mode and describe the result as (rows, cols, values).
fn arrayed(formula: &str, workbook: &Workbook) -> (usize, usize, Vec<CellValue>) {
    let expression = parse_formula(formula).unwrap();
    let context = EvalContext::new(workbook, SheetId(0));
    match evaluate_array(&expression, &context) {
        Value::Array(array) => {
            let mut values = Vec::new();
            for row in 0..array.rows() {
                for col in 0..array.cols() {
                    values.push(array.at(row, col));
                }
            }
            (array.rows(), array.cols(), values)
        }
        other => (1, 1, vec![other.into_scalar()]),
    }
}

fn values(formula: &str, workbook: &Workbook) -> Vec<CellValue> {
    arrayed(formula, workbook).2
}

fn sales_workbook() -> Workbook {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Data");
    put(&mut sheet, "A1", t("Region"));
    put(&mut sheet, "B1", t("Item"));
    put(&mut sheet, "C1", t("Sales"));
    put(&mut sheet, "D1", t("Qty"));
    put(&mut sheet, "E1", t("Quarter"));
    let rows = [
        ("East", "Apple", 10.0, 1.0, "Q1"),
        ("West", "Banana", 20.0, 2.0, "Q1"),
        ("east", "Apple", 30.0, 3.0, "Q2"),
        ("East", "Banana", 15.0, 4.0, "Q1"),
        ("West", "Apple", 25.0, 5.0, "Q2"),
    ];
    for (idx, (region, item, sales, qty, quarter)) in rows.iter().enumerate() {
        let r = idx + 2;
        put(&mut sheet, &format!("A{r}"), t(region));
        put(&mut sheet, &format!("B{r}"), t(item));
        put(&mut sheet, &format!("C{r}"), n(*sales));
        put(&mut sheet, &format!("D{r}"), n(*qty));
        put(&mut sheet, &format!("E{r}"), t(quarter));
    }
    workbook.sheets.push(sheet);
    workbook
}

#[test]
fn groupby_sums_and_totals_by_distinct_row_fields() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,SUM)", &workbook),
        (
            3,
            2,
            vec![t("East"), n(55.0), t("West"), n(45.0), t("Total"), n(100.0),]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,SUM,0,0)", &workbook),
        (2, 2, vec![t("East"), n(55.0), t("West"), n(45.0)])
    );
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:D6,SUM,0,-1)", &workbook),
        (
            3,
            3,
            vec![
                t("Total"),
                n(100.0),
                n(15.0),
                t("East"),
                n(55.0),
                n(8.0),
                t("West"),
                n(45.0),
                n(7.0),
            ]
        )
    );
}

#[test]
fn groupby_multi_column_row_fields_subtotals_and_relationship() {
    let workbook = sales_workbook();
    let expected_subtotals = (
        7,
        3,
        vec![
            t("East"),
            t("Apple"),
            n(40.0),
            t("East"),
            t("Banana"),
            n(15.0),
            t("East"),
            t(""),
            n(55.0),
            t("West"),
            t("Apple"),
            n(25.0),
            t("West"),
            t("Banana"),
            n(20.0),
            t("West"),
            t(""),
            n(45.0),
            t("Grand Total"),
            t(""),
            n(100.0),
        ],
    );
    assert_eq!(
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0,2)", &workbook),
        expected_subtotals
    );
    assert_eq!(
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0)", &workbook),
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0,1)", &workbook)
    );
    assert_eq!(
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0,1)", &workbook),
        (
            5,
            3,
            vec![
                t("East"),
                t("Apple"),
                n(40.0),
                t("East"),
                t("Banana"),
                n(15.0),
                t("West"),
                t("Apple"),
                n(25.0),
                t("West"),
                t("Banana"),
                n(20.0),
                t("Total"),
                t(""),
                n(100.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0,-2)", &workbook),
        (
            7,
            3,
            vec![
                t("Grand Total"),
                t(""),
                n(100.0),
                t("East"),
                t(""),
                n(55.0),
                t("East"),
                t("Apple"),
                n(40.0),
                t("East"),
                t("Banana"),
                n(15.0),
                t("West"),
                t(""),
                n(45.0),
                t("West"),
                t("Apple"),
                n(25.0),
                t("West"),
                t("Banana"),
                n(20.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0,0,-3,,1)", &workbook),
        (
            4,
            3,
            vec![
                t("East"),
                t("Apple"),
                n(40.0),
                t("West"),
                t("Apple"),
                n(25.0),
                t("West"),
                t("Banana"),
                n(20.0),
                t("East"),
                t("Banana"),
                n(15.0),
            ]
        )
    );
    let mut wb3 = Workbook::default();
    let mut sh3 = Sheet::new("Data");
    for (idx, (r1, r2, r3, v)) in [
        ("East", "A", "x", 10.0),
        ("East", "A", "y", 20.0),
        ("East", "B", "x", 5.0),
        ("West", "A", "x", 15.0),
    ]
    .iter()
    .enumerate()
    {
        let row = idx + 2;
        put(&mut sh3, &format!("A{row}"), t(r1));
        put(&mut sh3, &format!("B{row}"), t(r2));
        put(&mut sh3, &format!("C{row}"), t(r3));
        put(&mut sh3, &format!("D{row}"), n(*v));
    }
    wb3.sheets.push(sh3);
    assert_eq!(
        arrayed("GROUPBY(A2:C5,D2:D5,SUM,0,2)", &wb3),
        (
            7,
            4,
            vec![
                t("East"),
                t("A"),
                t("x"),
                n(10.0),
                t("East"),
                t("A"),
                t("y"),
                n(20.0),
                t("East"),
                t("B"),
                t("x"),
                n(5.0),
                t("East"),
                t(""),
                t(""),
                n(35.0),
                t("West"),
                t("A"),
                t("x"),
                n(15.0),
                t("West"),
                t(""),
                t(""),
                n(15.0),
                t("Grand Total"),
                t(""),
                t(""),
                n(50.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:C5,D2:D5,SUM,0,3)", &wb3),
        (
            10,
            4,
            vec![
                t("East"),
                t("A"),
                t("x"),
                n(10.0),
                t("East"),
                t("A"),
                t("y"),
                n(20.0),
                t("East"),
                t("A"),
                t(""),
                n(30.0),
                t("East"),
                t("B"),
                t("x"),
                n(5.0),
                t("East"),
                t("B"),
                t(""),
                n(5.0),
                t("East"),
                t(""),
                t(""),
                n(35.0),
                t("West"),
                t("A"),
                t("x"),
                n(15.0),
                t("West"),
                t("A"),
                t(""),
                n(15.0),
                t("West"),
                t(""),
                t(""),
                n(15.0),
                t("Grand Total"),
                t(""),
                t(""),
                n(50.0),
            ]
        )
    );
}

#[test]
fn groupby_headers_auto_and_explicit_modes() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A1:A6,C1:C6,SUM)", &workbook),
        (
            3,
            2,
            vec![t("East"), n(55.0), t("West"), n(45.0), t("Total"), n(100.0),]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A1:B6,C1:C6,SUM)", &workbook),
        (
            6,
            3,
            vec![
                t("Region"),
                t("Item"),
                t("Sales"),
                t("East"),
                t("Apple"),
                n(40.0),
                t("East"),
                t("Banana"),
                n(15.0),
                t("West"),
                t("Apple"),
                n(25.0),
                t("West"),
                t("Banana"),
                n(20.0),
                t("Total"),
                t(""),
                n(100.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A1:A6,C1:C6,SUM,3,1)", &workbook),
        (
            4,
            2,
            vec![
                t("Region"),
                t("Sales"),
                t("East"),
                n(55.0),
                t("West"),
                n(45.0),
                t("Total"),
                n(100.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,SUM,2,0)", &workbook),
        (
            3,
            2,
            vec![
                t("Row Field 1"),
                t("Value 1"),
                t("East"),
                n(55.0),
                t("West"),
                n(45.0),
            ]
        )
    );
    let mut wb_text = Workbook::default();
    let mut sh_text = Sheet::new("Data");
    put(&mut sh_text, "A1", t("ID"));
    put(&mut sh_text, "A2", n(1.0));
    put(&mut sh_text, "A3", n(2.0));
    put(&mut sh_text, "B1", t("Name"));
    put(&mut sh_text, "B2", t("x"));
    put(&mut sh_text, "B3", t("y"));
    wb_text.sheets.push(sh_text);
    assert_eq!(
        arrayed("GROUPBY(A1:A3,B1:B3,CONCAT)", &wb_text),
        (
            4,
            2,
            vec![
                n(1.0),
                t("x"),
                n(2.0),
                t("y"),
                t("ID"),
                t("Name"),
                t("Total"),
                t("Namexy"),
            ]
        )
    );
}

#[test]
fn groupby_sort_descending_and_sort_order_arrays() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(B2:B6,C2:C6,SUM,0,0,-1)", &workbook),
        (2, 2, vec![t("Banana"), n(35.0), t("Apple"), n(65.0)])
    );
    assert_eq!(
        arrayed("GROUPBY(B2:B6,C2:C6,SUM,0,0,-2)", &workbook),
        (2, 2, vec![t("Apple"), n(65.0), t("Banana"), n(35.0)])
    );
    assert_eq!(
        arrayed("GROUPBY(A2:B6,C2:C6,SUM,0,0,{-1,-2})", &workbook),
        (
            4,
            3,
            vec![
                t("West"),
                t("Banana"),
                n(20.0),
                t("West"),
                t("Apple"),
                n(25.0),
                t("East"),
                t("Banana"),
                n(15.0),
                t("East"),
                t("Apple"),
                n(40.0),
            ]
        )
    );
}

/// Rows the filter drops are not considered at all, so the first kept row
/// ("east", not the dropped "East") names the case-insensitive group, exactly
/// as grouping the FILTERed columns would.
#[test]
fn groupby_filter_array_and_all_filtered_calc_error() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,SUM,0,1,1,C2:C6>=20)", &workbook),
        (
            3,
            2,
            vec![t("east"), n(30.0), t("West"), n(45.0), t("Total"), n(75.0),]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,SUM,0,1,1,C2:C6>=20)", &workbook),
        arrayed(
            "GROUPBY(FILTER(A2:A6,C2:C6>=20),FILTER(C2:C6,C2:C6>=20),SUM,0,1)",
            &workbook
        )
    );
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,SUM,0,1,1,C2:C6>100)", &workbook),
        vec![CellValue::Error {
            value: ErrorValue::Calc
        }]
    );
}

#[test]
fn groupby_lambda_and_eta_reduced_aggregations() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,LAMBDA(x,SUM(x)*2),0,0)", &workbook),
        (2, 2, vec![t("East"), n(110.0), t("West"), n(90.0)])
    );
    assert_eq!(
        arrayed(
            "GROUPBY(A2:A6,C2:C6,LAMBDA(x,all,SUM(x)/SUM(all)),0,1)",
            &workbook
        ),
        (
            3,
            2,
            vec![t("East"), n(0.55), t("West"), n(0.45), t("Total"), n(1.0)]
        )
    );
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,AVERAGE,0,0)", &workbook),
        vec![t("East"), n(55.0 / 3.0), t("West"), n(22.5)]
    );
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,COUNT,0,0)", &workbook),
        vec![t("East"), n(3.0), t("West"), n(2.0)]
    );
    assert_eq!(
        values("GROUPBY(A2:A6,B2:B6,COUNTA,0,0)", &workbook),
        vec![t("East"), n(3.0), t("West"), n(2.0)]
    );
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,MAX,0,0)", &workbook),
        vec![t("East"), n(30.0), t("West"), n(25.0)]
    );
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,MIN,0,0)", &workbook),
        vec![t("East"), n(10.0), t("West"), n(20.0)]
    );
    assert_eq!(
        values("GROUPBY(A2:A6,D2:D6,PRODUCT,0,0)", &workbook),
        vec![t("East"), n(12.0), t("West"), n(10.0)]
    );
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,MEDIAN,0,0)", &workbook),
        vec![t("East"), n(15.0), t("West"), n(22.5)]
    );
    for name in ["VAR.S", "STDEV.S"] {
        let grouped = values(&format!("GROUPBY(A2:A6,C2:C6,{name},0,0)"), &workbook);
        let east = values(&format!("{name}(C2,C4,C5)"), &workbook);
        let west = values(&format!("{name}(C3,C6)"), &workbook);
        assert_eq!(
            grouped,
            vec![t("East"), east[0].clone(), t("West"), west[0].clone()]
        );
        assert_eq!(
            grouped,
            values(
                &format!("GROUPBY(A2:A6,C2:C6,LAMBDA(x,{name}(x)),0,0)"),
                &workbook
            )
        );
    }
    let close = |cell: &CellValue, want: f64| matches!(cell, CellValue::Number { value } if (value - want).abs() < 1e-9);
    let var = values("GROUPBY(A2:A6,C2:C6,VAR.S,0,0)", &workbook);
    assert!(close(&var[1], 325.0 / 3.0) && close(&var[3], 12.5));
    let stdev = values("GROUPBY(A2:A6,C2:C6,STDEV.S,0,0)", &workbook);
    assert!(close(&stdev[1], (325.0_f64 / 3.0).sqrt()) && close(&stdev[3], 12.5_f64.sqrt()));
    assert_eq!(
        values("GROUPBY(A2:A6,B2:B6,CONCAT,0,0)", &workbook),
        vec![
            t("East"),
            t("AppleAppleBanana"),
            t("West"),
            t("BananaApple"),
        ]
    );
}

#[test]
fn groupby_and_pivotby_share_of_total_and_relative_to() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed(
            "GROUPBY(A2:A6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,1)",
            &workbook
        ),
        (
            3,
            2,
            vec![t("East"), n(0.55), t("West"), n(0.45), t("Total"), n(1.0)]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,1,1,1,1,,0)",
            &workbook
        ),
        (
            4,
            4,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("Total"),
                t("East"),
                n(25.0 / 45.0),
                n(30.0 / 55.0),
                n(0.55),
                t("West"),
                n(20.0 / 45.0),
                n(25.0 / 55.0),
                n(0.45),
                t("Total"),
                n(1.0),
                n(1.0),
                n(1.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,1,1,1,1,,1)",
            &workbook
        ),
        (
            4,
            4,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("Total"),
                t("East"),
                n(25.0 / 55.0),
                n(30.0 / 55.0),
                n(1.0),
                t("West"),
                n(20.0 / 45.0),
                n(25.0 / 45.0),
                n(1.0),
                t("Total"),
                n(0.45),
                n(0.55),
                n(1.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,0,1,0,1,,2)",
            &workbook
        ),
        (
            3,
            3,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("East"),
                n(0.25),
                n(0.30),
                t("West"),
                n(0.20),
                n(0.25),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,0,1,0,1,,3)",
            &workbook
        ),
        (
            3,
            3,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("East"),
                n(25.0 / 55.0),
                n(30.0 / 55.0),
                t("West"),
                n(20.0 / 45.0),
                n(25.0 / 45.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,0,1,0,1,,4)",
            &workbook
        ),
        (
            3,
            3,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("East"),
                n(25.0 / 45.0),
                n(30.0 / 55.0),
                t("West"),
                n(20.0 / 45.0),
                n(25.0 / 55.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:B6,E2:E6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)),0,0,1,0,1,,4)",
            &workbook
        ),
        (
            5,
            4,
            vec![
                t(""),
                t(""),
                t("Q1"),
                t("Q2"),
                t("East"),
                t("Apple"),
                n(10.0 / 25.0),
                n(1.0),
                t("East"),
                t("Banana"),
                n(15.0 / 25.0),
                t(""),
                t("West"),
                t("Apple"),
                t(""),
                n(1.0),
                t("West"),
                t("Banana"),
                n(1.0),
                t(""),
            ]
        )
    );
}

#[test]
fn pivotby_row_and_col_totals_subtotals_and_sorting() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("PIVOTBY(A1:A6,E1:E6,C1:C6,SUM,3,1,1,1,1)", &workbook),
        (
            6,
            4,
            vec![
                t(""),
                t("Quarter"),
                t(""),
                t(""),
                t(""),
                t("Q1"),
                t("Q2"),
                t("Total"),
                t("Region"),
                t("Sales"),
                t("Sales"),
                t("Sales"),
                t("East"),
                n(25.0),
                n(30.0),
                n(55.0),
                t("West"),
                n(20.0),
                n(25.0),
                n(45.0),
                t("Total"),
                n(45.0),
                n(55.0),
                n(100.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(B2:B6,E2:E6,C2:C6,SUM,0,0,1,0,1,A2:A6=\"East\")",
            &workbook
        ),
        (
            3,
            3,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("Apple"),
                n(10.0),
                n(30.0),
                t("Banana"),
                n(15.0),
                t(""),
            ]
        )
    );
    assert_eq!(
        arrayed("PIVOTBY(A2:A6,E2:E6,C2:C6,SUM,0,-1,-1,-1,-1)", &workbook),
        (
            4,
            4,
            vec![
                t(""),
                t("Total"),
                t("Q2"),
                t("Q1"),
                t("Total"),
                n(100.0),
                n(55.0),
                n(45.0),
                t("West"),
                n(45.0),
                n(25.0),
                n(20.0),
                t("East"),
                n(55.0),
                n(30.0),
                n(25.0),
            ]
        )
    );
}

#[test]
fn groupby_and_pivotby_size_mismatch_and_invalid_args_give_value_error() {
    let workbook = sales_workbook();
    let val_err = vec![CellValue::Error {
        value: ErrorValue::Value,
    }];
    assert_eq!(values("GROUPBY(A2:A5,C2:C6,SUM)", &workbook), val_err);
    assert_eq!(
        values("GROUPBY(A2:A6,C2:C6,SUM,0,1,1,D2:D5>0)", &workbook),
        val_err
    );
    assert_eq!(values("GROUPBY(A2:A6,C2:C6,SUM,4)", &workbook), val_err);
    assert_eq!(values("GROUPBY(A2:A6,C2:C6,SUM,0,2)", &workbook), val_err);
    assert_eq!(values("GROUPBY(A2:A6,C2:C6,SUM,0,3)", &workbook), val_err);
    assert_eq!(values("GROUPBY(A2:A6,C2:C6,SUM,0,1,0)", &workbook), val_err);
    assert_eq!(values("GROUPBY(A2:A6,C2:C6,SUM,0,1,3)", &workbook), val_err);
    assert_eq!(
        values("GROUPBY(A2:B6,C2:C6,SUM,0,2,1,,1)", &workbook),
        val_err
    );
    assert_eq!(values("PIVOTBY(A2:A5,E2:E6,C2:C6,SUM)", &workbook), val_err);
    assert_eq!(values("PIVOTBY(A2:A6,E2:E5,C2:C6,SUM)", &workbook), val_err);
    assert_eq!(values("PIVOTBY(,,C2:C6,SUM)", &workbook), val_err);
    assert_eq!(
        values("PIVOTBY(A2:A6,E2:E6,C2:C6,SUM,0,1,1,1,1,,5)", &workbook),
        val_err
    );
}

#[test]
fn groupby_and_pivotby_blank_vs_zero_groups_and_omitted_axis() {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Data");
    put(&mut sheet, "A2", CellValue::Empty);
    put(&mut sheet, "A3", n(0.0));
    put(&mut sheet, "A4", n(0.0));
    put(&mut sheet, "B2", n(10.0));
    put(&mut sheet, "B3", n(20.0));
    put(&mut sheet, "B4", n(30.0));
    put(&mut sheet, "C2", t("R"));
    put(&mut sheet, "C3", t("R"));
    put(&mut sheet, "C4", t("R"));
    wb.sheets.push(sheet);
    assert_eq!(
        arrayed("GROUPBY(A2:A4,B2:B4,SUM,0,0)", &wb),
        (2, 2, vec![n(0.0), n(50.0), CellValue::Empty, n(10.0)])
    );
    assert_eq!(
        arrayed("PIVOTBY(C2:C4,A2:A4,B2:B4,SUM,0,0,1,0,1)", &wb),
        (
            2,
            3,
            vec![t(""), n(0.0), CellValue::Empty, t("R"), n(50.0), n(10.0)]
        )
    );
    let sales = sales_workbook();
    assert_eq!(
        arrayed("PIVOTBY(A2:A6,,C2:C6,SUM)", &sales),
        (
            3,
            2,
            vec![t("East"), n(55.0), t("West"), n(45.0), t("Total"), n(100.0),]
        )
    );
    assert_eq!(
        arrayed("PIVOTBY(,A2:A6,C2:C6,SUM)", &sales),
        (
            2,
            3,
            vec![t("East"), t("West"), t("Total"), n(55.0), n(45.0), n(100.0),]
        )
    );
}

#[test]
fn groupby_and_pivotby_vector_of_lambdas() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,HSTACK(SUM,AVERAGE),0,0)", &workbook),
        (
            3,
            3,
            vec![
                t(""),
                t("SUM"),
                t("AVERAGE"),
                t("East"),
                n(55.0),
                n(55.0 / 3.0),
                t("West"),
                n(45.0),
                n(22.5),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "GROUPBY(A1:A6,C1:C6,HSTACK(SUM,LAMBDA(x,COUNT(x))),3,0)",
            &workbook
        ),
        (
            4,
            3,
            vec![
                t(""),
                t("SUM"),
                t(""),
                t("Region"),
                t("Sales"),
                t("Sales"),
                t("East"),
                n(55.0),
                n(3.0),
                t("West"),
                n(45.0),
                n(2.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,VSTACK(SUM,AVERAGE),0,0)", &workbook),
        (
            4,
            2,
            vec![
                t("East"),
                n(55.0),
                t("East"),
                n(55.0 / 3.0),
                t("West"),
                n(45.0),
                t("West"),
                n(22.5),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,HSTACK(SUM,AVERAGE),0,0,1,0,1)",
            &workbook
        ),
        (
            4,
            5,
            vec![
                t(""),
                t("Q1"),
                t(""),
                t("Q2"),
                t(""),
                t(""),
                t("SUM"),
                t("AVERAGE"),
                t("SUM"),
                t("AVERAGE"),
                t("East"),
                n(25.0),
                n(12.5),
                n(30.0),
                n(30.0),
                t("West"),
                n(20.0),
                n(20.0),
                n(25.0),
                n(25.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A1:A6,E1:E6,C1:C6,HSTACK(SUM,AVERAGE),3,0,1,0,1)",
            &workbook
        ),
        (
            6,
            5,
            vec![
                t(""),
                t("Quarter"),
                t(""),
                t(""),
                t(""),
                t(""),
                t("Q1"),
                t(""),
                t("Q2"),
                t(""),
                t(""),
                t("SUM"),
                t("AVERAGE"),
                t("SUM"),
                t("AVERAGE"),
                t("Region"),
                t("Sales"),
                t("Sales"),
                t("Sales"),
                t("Sales"),
                t("East"),
                n(25.0),
                n(12.5),
                n(30.0),
                n(30.0),
                t("West"),
                n(20.0),
                n(20.0),
                n(25.0),
                n(25.0),
            ]
        )
    );
    assert_eq!(
        arrayed(
            "PIVOTBY(A2:A6,E2:E6,C2:C6,VSTACK(SUM,AVERAGE),0,0,1,0,1)",
            &workbook
        ),
        (
            5,
            3,
            vec![
                t(""),
                t("Q1"),
                t("Q2"),
                t("East"),
                n(25.0),
                n(30.0),
                t("East"),
                n(12.5),
                n(30.0),
                t("West"),
                n(20.0),
                n(25.0),
                t("West"),
                n(20.0),
                n(25.0),
            ]
        )
    );
}

#[test]
fn pivotby_charges_budget_for_cell_intersections() {
    let workbook = sales_workbook();
    assert_eq!(
        values(
            "_xlfn.LET(_xlpm.s,_xlfn.SEQUENCE(50000),PIVOTBY(MOD(_xlpm.s,500),MOD(_xlpm.s,499),_xlpm.s,SUM,0,0,1,0,1))",
            &workbook
        ),
        vec![CellValue::Error {
            value: ErrorValue::Num
        }]
    );
}

#[test]
fn groupby_and_pivotby_xlfn_and_xleta_spellings() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("_xlfn.GROUPBY(A2:A6,C2:C6,_xleta.SUM)", &workbook),
        arrayed("GROUPBY(A2:A6,C2:C6,SUM)", &workbook)
    );
    assert_eq!(
        arrayed(
            "_xlfn.GROUPBY(A2:A6,C2:C6,_xlfn.LAMBDA(_xlpm.x,_xlpm.y,SUM(_xlpm.x)/SUM(_xlpm.y)))",
            &workbook
        ),
        arrayed("GROUPBY(A2:A6,C2:C6,LAMBDA(x,y,SUM(x)/SUM(y)))", &workbook)
    );
    assert_eq!(
        arrayed("_xlfn.GROUPBY(A2:A6,C2:C6,_xleta.STDEV.S,0,0)", &workbook),
        arrayed("GROUPBY(A2:A6,C2:C6,STDEV.S,0,0)", &workbook)
    );
    assert_eq!(
        arrayed("_xlfn.PIVOTBY(A2:A6,E2:E6,C2:C6,_xleta.AVERAGE)", &workbook),
        arrayed("PIVOTBY(A2:A6,E2:E6,C2:C6,AVERAGE)", &workbook)
    );
    assert_eq!(
        values("_xlfn.BYROW(C2:D3,_xleta.SUM)", &workbook),
        vec![n(11.0), n(22.0)]
    );
}

/// Excel's PRODUCT of no numbers is 0, so a group of text multiplies to 0
/// rather than the empty product 1.
#[test]
fn product_of_a_block_without_numbers_is_zero() {
    let workbook = sales_workbook();
    assert_eq!(values("PRODUCT({\"a\",\"b\"})", &workbook), vec![n(0.0)]);
    assert_eq!(values("PRODUCT(IF({1,1},A2:A3))", &workbook), vec![n(0.0)]);
    assert_eq!(values("PRODUCT({2,\"a\",3})", &workbook), vec![n(6.0)]);
    assert_eq!(
        values("GROUPBY(A2:A6,B2:B6,PRODUCT,0,0)", &workbook),
        vec![t("East"), n(0.0), t("West"), n(0.0)]
    );
}

/// B3, C5 and a formula-style empty string in D4 inside A1:E10, the rest
/// blank; row 4 and column D hold no other value.
fn trim_workbook(with_empty_text: bool) -> Workbook {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Data");
    put(&mut sheet, "B3", n(1.0));
    put(&mut sheet, "C5", n(2.0));
    if with_empty_text {
        put(&mut sheet, "D4", t(""));
    }
    put(&mut sheet, "H20", t("far"));
    workbook.sheets.push(sheet);
    workbook
}

fn error(value: ErrorValue) -> Vec<CellValue> {
    vec![CellValue::Error { value }]
}

#[test]
fn trimrange_cuts_blank_edges_and_keeps_inner_blanks() {
    let workbook = trim_workbook(false);
    assert_eq!(
        arrayed("TRIMRANGE(A1:E10)", &workbook),
        (
            3,
            2,
            vec![
                n(1.0),
                CellValue::Empty,
                CellValue::Empty,
                CellValue::Empty,
                CellValue::Empty,
                n(2.0)
            ]
        )
    );
    assert_eq!(arrayed("TRIMRANGE(A1:E10,1,1)", &workbook).0, 8);
    assert_eq!(arrayed("TRIMRANGE(A1:E10,1,1)", &workbook).1, 4);
    assert_eq!(arrayed("TRIMRANGE(A1:E10,2,2)", &workbook).0, 5);
    assert_eq!(arrayed("TRIMRANGE(A1:E10,2,2)", &workbook).1, 3);
    assert_eq!(arrayed("TRIMRANGE(A1:E10,0,3)", &workbook).0, 10);
    assert_eq!(arrayed("TRIMRANGE(A1:E10,3,0)", &workbook).1, 5);
    assert_eq!(arrayed("TRIMRANGE(A1:E10,0,0)", &workbook).0, 10);
    assert_eq!(values("ROWS(TRIMRANGE(A1:E10))", &workbook), vec![n(3.0)]);
    assert_eq!(
        values("COLUMNS(TRIMRANGE(A1:E10))", &workbook),
        vec![n(2.0)]
    );
    assert_eq!(values("SUM(TRIMRANGE(A1:E10))", &workbook), vec![n(3.0)]);
}

#[test]
fn trimrange_mode_arguments_truncate_and_bad_modes_are_value_errors() {
    let workbook = trim_workbook(false);
    assert_eq!(
        arrayed("TRIMRANGE(A1:E10,2.9,2.9)", &workbook),
        arrayed("TRIMRANGE(A1:E10,2,2)", &workbook)
    );
    assert_eq!(
        values("TRIMRANGE(A1:E10,4)", &workbook),
        error(ErrorValue::Value)
    );
    assert_eq!(
        values("TRIMRANGE(A1:E10,,-1)", &workbook),
        error(ErrorValue::Value)
    );
    assert_eq!(
        values("TRIMRANGE(A1:E10,\"x\")", &workbook),
        error(ErrorValue::Value)
    );
    assert_eq!(
        values("TRIMRANGE(A1:E10,1,1,1)", &workbook),
        error(ErrorValue::Value)
    );
}

#[test]
fn trimrange_of_whole_columns_and_rows_reads_only_the_used_extent() {
    let workbook = trim_workbook(false);
    assert_eq!(values("ROWS(TRIMRANGE(B:C))", &workbook), vec![n(3.0)]);
    assert_eq!(values("ROWS(TRIMRANGE(B:C,2))", &workbook), vec![n(5.0)]);
    assert_eq!(values("COLUMNS(TRIMRANGE(3:5))", &workbook), vec![n(2.0)]);
    assert_eq!(values("SUM(TRIMRANGE(C:C))", &workbook), vec![n(2.0)]);
    assert_eq!(
        values("ROWS(TRIMRANGE(B:B,0))", &workbook),
        vec![n(1_048_576.0)]
    );
}

#[test]
fn trimrange_counts_an_empty_string_as_a_value() {
    let workbook = trim_workbook(true);
    assert_eq!(
        values("COLUMNS(TRIMRANGE(A1:E10))", &workbook),
        vec![n(3.0)]
    );
}

/// an all-blank range has no edge to stop at; Excel answers #REF!.
#[test]
fn trimrange_of_a_blank_range_is_a_ref_error() {
    let workbook = trim_workbook(false);
    assert_eq!(
        values("TRIMRANGE(J1:K5)", &workbook),
        error(ErrorValue::Ref)
    );
    assert_eq!(
        values("_xlfn._TRO_ALL(A30:B40)", &workbook),
        error(ErrorValue::Ref)
    );
}

#[test]
fn trimrange_of_a_computed_block_trims_nothing_without_blanks() {
    let workbook = trim_workbook(false);
    assert_eq!(
        arrayed("TRIMRANGE({1,2;3,4})", &workbook),
        (2, 2, vec![n(1.0), n(2.0), n(3.0), n(4.0)])
    );
}

/// files store `A1.:.E10`, `A1.:E10` and `A1:.E10` as these functions.
#[test]
fn trim_ref_forms_equal_their_trimrange_modes() {
    let workbook = trim_workbook(false);
    assert_eq!(
        arrayed("_xlfn._TRO_ALL(A1:E10)", &workbook),
        arrayed("TRIMRANGE(A1:E10,3,3)", &workbook)
    );
    assert_eq!(
        arrayed("_xlfn._TRO_LEADING(A1:E10)", &workbook),
        arrayed("TRIMRANGE(A1:E10,1,1)", &workbook)
    );
    assert_eq!(
        arrayed("_xlfn._TRO_TRAILING(A1:E10)", &workbook),
        arrayed("TRIMRANGE(A1:E10,2,2)", &workbook)
    );
    assert_eq!(
        arrayed("_xlfn.TRIMRANGE(A1:E10)", &workbook),
        arrayed("TRIMRANGE(A1:E10)", &workbook)
    );
    assert_eq!(
        values("SUM(_xlfn._TRO_TRAILING(B:B))", &workbook),
        vec![n(1.0)]
    );
}

/// the table of the Excel ground-truth sheet: Region, Product, Year and Sales
/// in D1:G13.
fn excel_workbook() -> Workbook {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Group");
    for (col, header) in ["D", "E", "F", "G"]
        .into_iter()
        .zip(["Region", "Product", "Year", "Sales"])
    {
        put(&mut sheet, &format!("{col}1"), t(header));
    }
    let rows = [
        ("East", "Pen", 2023.0, 120.0),
        ("West", "Pen", 2023.0, 80.0),
        ("East", "Ink", 2023.0, 45.5),
        ("North", "Pen", 2023.0, 60.0),
        ("West", "Ink", 2024.0, 30.25),
        ("East", "Pen", 2024.0, 140.0),
        ("North", "Ink", 2024.0, 55.0),
        ("West", "Pen", 2024.0, 95.0),
        ("East", "Ink", 2023.0, 50.0),
        ("North", "Pen", 2024.0, 75.0),
        ("West", "Ink", 2024.0, 35.0),
        ("East", "Pen", 2023.0, 0.1),
    ];
    for (idx, (region, product, year, sales)) in rows.into_iter().enumerate() {
        let r = idx + 2;
        put(&mut sheet, &format!("D{r}"), t(region));
        put(&mut sheet, &format!("E{r}"), t(product));
        put(&mut sheet, &format!("F{r}"), n(year));
        put(&mut sheet, &format!("G{r}"), n(sales));
    }
    workbook.sheets.push(sheet);
    workbook
}

/// a block of the given shape whose numbers agree to 1e-12 relative and
/// whose other values are equal.
fn assert_block(formula: &str, workbook: &Workbook, shape: (usize, usize), want: &[CellValue]) {
    let (rows, cols, got) = arrayed(formula, workbook);
    assert_eq!((rows, cols), shape, "{formula}");
    assert_eq!(got.len(), want.len(), "{formula}");
    for (got, want) in got.iter().zip(want) {
        match (got, want) {
            (CellValue::Number { value: g }, CellValue::Number { value: w }) => {
                assert!(
                    (g - w).abs() <= 1e-12 * w.abs().max(1.0),
                    "{formula}: {g} vs {w}"
                );
            }
            _ => assert_eq!(got, want, "{formula}"),
        }
    }
}

/// Excel 365 for Mac result: each region's share of all sales.
#[test]
fn groupby_with_eta_percentof_matches_excel() {
    let workbook = excel_workbook();
    let want = [
        t("East"),
        n(0.452503658458993),
        t("North"),
        n(0.24177642043647),
        t("West"),
        n(0.305719921104536),
        t("Total"),
        n(1.0),
    ];
    assert_block("GROUPBY(D2:D13,G2:G13,PERCENTOF)", &workbook, (4, 2), &want);
    assert_block(
        "_xlfn.GROUPBY(D2:D13,G2:G13,_xleta.PERCENTOF)",
        &workbook,
        (4, 2),
        &want,
    );
}

/// Excel 365 for Mac result: each cell's share of its year's column.
#[test]
fn pivotby_with_eta_percentof_matches_excel() {
    assert_block(
        "_xlfn.PIVOTBY(D2:D13,F2:F13,G2:G13,_xleta.PERCENTOF)",
        &excel_workbook(),
        (5, 4),
        &[
            t(""),
            n(2023.0),
            n(2024.0),
            t("Total"),
            t("East"),
            n(0.606299212598425),
            n(0.325392213829169),
            n(0.452503658458993),
            t("North"),
            n(0.168728908886389),
            n(0.302149912841371),
            n(0.24177642043647),
            t("West"),
            n(0.224971878515186),
            n(0.37245787332946),
            n(0.305719921104536),
            t("Total"),
            n(1.0),
            n(1.0),
            n(1.0),
        ],
    );
}

/// Excel 365 for Mac result for GROUPBY; each group's items in sheet order.
#[test]
fn groupby_and_pivotby_with_eta_arraytotext_match_excel() {
    let workbook = excel_workbook();
    assert_block(
        "_xlfn.GROUPBY(D2:D13,E2:E13,_xleta.ARRAYTOTEXT)",
        &workbook,
        (4, 2),
        &[
            t("East"),
            t("Pen, Ink, Pen, Ink, Pen"),
            t("North"),
            t("Pen, Ink, Pen"),
            t("West"),
            t("Pen, Ink, Pen, Ink"),
            t("Total"),
            t("Pen, Pen, Ink, Pen, Ink, Pen, Ink, Pen, Ink, Pen, Ink, Pen"),
        ],
    );
    assert_block(
        "PIVOTBY(D2:D13,F2:F13,E2:E13,ARRAYTOTEXT)",
        &workbook,
        (5, 4),
        &[
            t(""),
            n(2023.0),
            n(2024.0),
            t("Total"),
            t("East"),
            t("Pen, Ink, Ink, Pen"),
            t("Pen"),
            t("Pen, Ink, Pen, Ink, Pen"),
            t("North"),
            t("Pen"),
            t("Ink, Pen"),
            t("Pen, Ink, Pen"),
            t("West"),
            t("Pen"),
            t("Ink, Pen, Ink"),
            t("Pen, Ink, Pen, Ink"),
            t("Total"),
            t("Pen, Pen, Ink, Pen, Ink, Pen"),
            t("Ink, Pen, Ink, Pen, Pen, Ink"),
            t("Pen, Pen, Ink, Pen, Ink, Pen, Ink, Pen, Ink, Pen, Ink, Pen"),
        ],
    );
}

/// Excel 365 for Mac results of the ground-truth sheet, each written out by
/// ARRAYTOTEXT in strict form.
#[test]
fn group_sheet_matches_excel() {
    let workbook = excel_workbook();
    let cases: &[(&str, Result<&str, ErrorValue>)] = &[
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.SUM),1)"#,
            Ok(r#"{"East",355.6;"North",190;"West",240.25;"Total",785.85}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$1:$D$13,$G$1:$G$13,_xleta.SUM,3,1),1)"#,
            Ok(r#"{"Region","Sales";"East",355.6;"North",190;"West",240.25;"Total",785.85}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.SUM,0,0),1)"#,
            Ok(r#"{"East",355.6;"North",190;"West",240.25}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.SUM,0,1,-2),1)"#,
            Ok(r#"{"East",355.6;"West",240.25;"North",190;"Total",785.85}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.SUM,0,1,-1),1)"#,
            Ok(r#"{"West",240.25;"North",190;"East",355.6;"Total",785.85}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$E$13,$G$2:$G$13,_xleta.SUM),1)"#,
            Ok(
                r#"{"East","Ink",95.5;"East","Pen",260.1;"North","Ink",55;"North","Pen",135;"West","Ink",65.25;"West","Pen",175;"Total","",785.85}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$E$13,$G$2:$G$13,_xleta.SUM,0,2),1)"#,
            Ok(
                r#"{"East","Ink",95.5;"East","Pen",260.1;"East","",355.6;"North","Ink",55;"North","Pen",135;"North","",190;"West","Ink",65.25;"West","Pen",175;"West","",240.25;"Grand Total","",785.85}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.AVERAGE),1)"#,
            Ok(r#"{"East",71.12;"North",63.3333333333333;"West",60.0625;"Total",65.4875}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.COUNT),1)"#,
            Ok(r#"{"East",5;"North",3;"West",4;"Total",12}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.MAX),1)"#,
            Ok(r#"{"East",140;"North",75;"West",95;"Total",140}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.PERCENTOF),1)"#,
            Ok(
                r#"{"East",0.452503658458993;"North",0.24177642043647;"West",0.305719921104536;"Total",1}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xlfn.LAMBDA(_xlpm.x,SUM(_xlpm.x)*2)),1)"#,
            Ok(r#"{"East",711.2;"North",380;"West",480.5;"Total",1571.7}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xlfn.HSTACK(_xleta.SUM,_xleta.MIN)),1)"#,
            Ok(
                r#"{"","SUM","MIN";"East",355.6,0.1;"North",190,55;"West",240.25,30.25;"Total",785.85,0.1}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$G$2:$G$13,_xleta.SUM,0,1,,$F$2:$F$13=2024),1)"#,
            Ok(r#"{"East",140;"North",130;"West",160.25;"Total",430.25}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$D$13,$E$2:$E$13,_xleta.ARRAYTOTEXT),1)"#,
            Ok(
                r#"{"East","Pen, Ink, Pen, Ink, Pen";"North","Pen, Ink, Pen";"West","Pen, Ink, Pen, Ink";"Total","Pen, Pen, Ink, Pen, Ink, Pen, Ink, Pen, Ink, Pen, Ink, Pen"}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.GROUPBY($D$2:$E$13,$G$2:$G$13,_xleta.SUM,0,2,,,1),1)"#,
            Err(ErrorValue::Value),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$2:$D$13,$F$2:$F$13,$G$2:$G$13,_xleta.SUM),1)"#,
            Ok(
                r#"{"",2023,2024,"Total";"East",215.6,140,355.6;"North",60,130,190;"West",80,160.25,240.25;"Total",355.6,430.25,785.85}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$1:$D$13,$F$1:$F$13,$G$1:$G$13,_xleta.SUM,3),1)"#,
            Ok(
                r#"{"","Year","","";"",2023,2024,"Total";"Region","Sales","Sales","Sales";"East",215.6,140,355.6;"North",60,130,190;"West",80,160.25,240.25;"Total",355.6,430.25,785.85}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$2:$D$13,$F$2:$F$13,$G$2:$G$13,_xleta.SUM,0,0,,0),1)"#,
            Ok(r#"{"",2023,2024;"East",215.6,140;"North",60,130;"West",80,160.25}"#),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$2:$D$13,$F$2:$F$13,$G$2:$G$13,_xleta.PERCENTOF),1)"#,
            Ok(
                r#"{"",2023,2024,"Total";"East",0.606299212598425,0.325392213829169,0.452503658458993;"North",0.168728908886389,0.302149912841371,0.24177642043647;"West",0.224971878515186,0.37245787332946,0.305719921104536;"Total",1,1,1}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$2:$D$13,$E$2:$E$13,$G$2:$G$13,_xleta.COUNT),1)"#,
            Ok(
                r#"{"","Ink","Pen","Total";"East",2,3,5;"North",1,2,3;"West",2,2,4;"Total",5,7,12}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$2:$D$13,$F$2:$F$13,$G$2:$G$13,_xleta.SUM,0,1,1,1,-1),1)"#,
            Ok(
                r#"{"",2024,2023,"Total";"East",140,215.6,355.6;"North",130,60,190;"West",160.25,80,240.25;"Total",430.25,355.6,785.85}"#,
            ),
        ),
        (
            r#"_xlfn.ARRAYTOTEXT(_xlfn.PIVOTBY($D$2:$D$13,$F$2:$F$13,$G$2:$G$13,_xleta.SUM,0,1,,1,,$E$2:$E$13="Pen"),1)"#,
            Ok(
                r#"{"",2023,2024,"Total";"East",120.1,140,260.1;"North",60,75,135;"West",80,95,175;"Total",260.1,310,570.1}"#,
            ),
        ),
    ];
    for (formula, want) in cases {
        let want = match want {
            Ok(value) => t(value),
            Err(value) => CellValue::Error { value: *value },
        };
        assert_eq!(values(formula, &workbook), vec![want], "{formula}");
    }
}

/// keys a hair apart past 1E15 stay apart, and sums past 1E308 are #NUM!.
#[test]
fn groupby_keys_and_values_at_extreme_magnitudes() {
    let mut workbook = Workbook::default();
    let mut sheet = Sheet::new("Data");
    let rows = [
        (1e15 + 0.75, 1e200),
        (1e15 + 0.5, 1.0),
        (1e15 + 0.75, 2.0),
        (1e200, 1e200),
        (-1e200, 3.0),
        (1e200, 1e200),
    ];
    for (idx, (key, value)) in rows.into_iter().enumerate() {
        put(&mut sheet, &format!("A{}", idx + 1), n(key));
        put(&mut sheet, &format!("B{}", idx + 1), n(value));
    }
    put(&mut sheet, "C1", t("a"));
    put(&mut sheet, "C2", t("a"));
    put(&mut sheet, "C3", t("b"));
    put(&mut sheet, "D1", n(1e308));
    put(&mut sheet, "D2", n(1e308));
    put(&mut sheet, "D3", n(-1e308));
    workbook.sheets.push(sheet);
    assert_block(
        "GROUPBY(A1:A6,B1:B6,SUM,0,1)",
        &workbook,
        (5, 2),
        &[
            n(-1e200),
            n(3.0),
            n(1e15 + 0.5),
            n(1.0),
            n(1e15 + 0.75),
            n(1e200),
            n(1e200),
            n(2e200),
            t("Total"),
            n(3e200),
        ],
    );
    assert_block(
        "GROUPBY(A1:A6,B1:B6,COUNT,0,0,-1)",
        &workbook,
        (4, 2),
        &[
            n(1e200),
            n(2.0),
            n(1e15 + 0.75),
            n(2.0),
            n(1e15 + 0.5),
            n(1.0),
            n(-1e200),
            n(1.0),
        ],
    );
    assert_eq!(
        values("GROUPBY(C1:C3,D1:D3,SUM,0,0)", &workbook),
        vec![
            t("a"),
            CellValue::Error {
                value: ErrorValue::Num
            },
            t("b"),
            n(-1e308),
        ]
    );
}

/// a total depth or sort column past the fields is #VALUE! however far past,
/// even where a 32-bit index would wrap around to a valid one.
#[test]
fn groupby_out_of_range_depth_and_sort_columns_at_huge_magnitudes() {
    let workbook = sales_workbook();
    for formula in [
        "GROUPBY(A2:A6,C2:C6,SUM,0,4294967297)",
        "GROUPBY(A2:A6,C2:C6,SUM,0,-4294967297)",
        "GROUPBY(A2:A6,C2:C6,SUM,0,1E200)",
        "GROUPBY(A2:A6,C2:C6,SUM,0,1,4294967297)",
        "GROUPBY(A2:A6,C2:C6,SUM,0,1,-4294967298)",
        "GROUPBY(A2:A6,C2:C6,SUM,0,1,1E15+0.75)",
        "PIVOTBY(A2:A6,E2:E6,C2:C6,SUM,0,4294967297)",
        "PIVOTBY(A2:A6,E2:E6,C2:C6,SUM,0,1,1,4294967297)",
        "PIVOTBY(A2:A6,E2:E6,C2:C6,SUM,0,1,1,1,-4294967298)",
    ] {
        assert_eq!(
            values(formula, &workbook),
            error(ErrorValue::Value),
            "{formula}"
        );
    }
}

/// one data row, an all-blank range and values that are all errors.
#[test]
fn groupby_and_pivotby_single_row_blank_and_error_inputs() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A2,C2,SUM)", &workbook),
        (2, 2, vec![t("East"), n(10.0), t("Total"), n(10.0)])
    );
    assert_eq!(
        arrayed("PIVOTBY(A2,E2,C2,SUM)", &workbook),
        (
            3,
            3,
            vec![
                t(""),
                t("Q1"),
                t("Total"),
                t("East"),
                n(10.0),
                n(10.0),
                t("Total"),
                n(10.0),
                n(10.0),
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(J2:J4,K2:K4,SUM)", &workbook),
        (2, 2, vec![CellValue::Empty, n(0.0), t("Total"), n(0.0)])
    );
    let na = CellValue::Error {
        value: ErrorValue::NA,
    };
    let mut errors = sales_workbook();
    for row in 2..=6 {
        errors.sheets[0].set_cell(
            CellRef::parse_a1(&format!("C{row}")).unwrap(),
            Cell {
                value: na.clone(),
                ..Cell::default()
            },
        );
    }
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,SUM)", &errors),
        (
            3,
            2,
            vec![
                t("East"),
                na.clone(),
                t("West"),
                na.clone(),
                t("Total"),
                na.clone()
            ]
        )
    );
    assert_eq!(
        arrayed("GROUPBY(A2:A6,C2:C6,COUNT,0,0)", &errors),
        (2, 2, vec![t("East"), n(0.0), t("West"), n(0.0)])
    );
}

/// whole columns and the whole sheet as arguments: a filter on the key
/// column drops the blank rows past the data, as Excel's would.
#[test]
fn groupby_and_trimrange_over_whole_columns_and_the_whole_sheet() {
    let workbook = sales_workbook();
    assert_eq!(
        arrayed("GROUPBY(A:A,C:C,SUM,,,,A:A<>\"\")", &workbook),
        arrayed("GROUPBY(A1:A6,C1:C6,SUM)", &workbook)
    );
    let trim = trim_workbook(false);
    assert_eq!(values("ROWS(TRIMRANGE(A:XFD))", &trim), vec![n(18.0)]);
    assert_eq!(values("COLUMNS(TRIMRANGE(A:XFD))", &trim), vec![n(7.0)]);
    assert_eq!(values("SUM(TRIMRANGE(A:XFD))", &trim), vec![n(3.0)]);
    assert_eq!(
        values("ROWS(TRIMRANGE(A:XFD,1,1))", &trim),
        vec![n(1_048_574.0)]
    );
    assert_eq!(
        values("COLUMNS(TRIMRANGE(A:XFD,1,1))", &trim),
        vec![n(16_383.0)]
    );
    assert_eq!(
        values("ROWS(TRIMRANGE(A:XFD,0,0))", &trim),
        vec![n(1_048_576.0)]
    );
    assert_eq!(
        values("COLUMNS(TRIMRANGE(A:XFD,0,0))", &trim),
        vec![n(16_384.0)]
    );
}

/// more distinct keys than the engine's largest array block holds is #NUM!
/// before any group is aggregated.
#[test]
fn groupby_with_more_distinct_keys_than_an_output_block_holds_is_num() {
    let workbook = sales_workbook();
    assert_eq!(
        values("LET(s,SEQUENCE(600000),GROUPBY(s,s,SUM,0,0))", &workbook),
        error(ErrorValue::Num)
    );
    assert_eq!(
        values("LET(s,SEQUENCE(600000),PIVOTBY(s,,s,SUM,0,0))", &workbook),
        error(ErrorValue::Num)
    );
}
