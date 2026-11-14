//! functions the engine leaves to the value Excel cached: GETPIVOTDATA reads a
//! pivot table's layout, which the workbook model does not carry, and the
//! FORECAST.ETS family fits its smoothing parameters in a way Microsoft does
//! not document closely enough to reproduce Excel's numbers.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellProvider, CellRef, CellValue, Sheet, SheetId, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn cached(formula: &str, value: f64) -> Workbook {
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        CellRef::new(2, 0),
        Cell {
            value: CellValue::Text {
                value: "Sum of Sales".into(),
            },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        CellRef::new(0, 5),
        Cell {
            formula: Some(formula.into()),
            value: number(value),
            ..Cell::default()
        },
    );
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    workbook
}

#[test]
fn getpivotdata_keeps_the_cached_value() {
    for (formula, value) in [
        ("GETPIVOTDATA(\"Sales\",$A$3,\"Region\",\"East\")", 42.0),
        ("GETPIVOTDATA(\"Sales\",$A$3)", 120.0),
        (
            "GETPIVOTDATA(\"Sales\",$A$3,\"Region\",\"East\")/GETPIVOTDATA(\"Sales\",$A$3)",
            0.35,
        ),
    ] {
        let mut workbook = cached(formula, value);
        rebuild_and_recalc_all(&mut workbook, None);
        assert_eq!(
            workbook.value(SheetId(0), CellRef::new(0, 5)),
            number(value),
            "{formula}"
        );
    }
}

#[test]
fn forecast_ets_family_keeps_the_cached_value() {
    for (formula, value) in [
        ("_xlfn.FORECAST.ETS(13,B1:B12,A1:A12)", 151.5),
        ("_xlfn.FORECAST.ETS.CONFINT(13,B1:B12,A1:A12,0.9)", 12.25),
        ("_xlfn.FORECAST.ETS.SEASONALITY(B1:B12,A1:A12)", 4.0),
        ("_xlfn.FORECAST.ETS.STAT(B1:B12,A1:A12,1)", 0.5),
        ("FORECAST.ETS(13,B1:B12,A1:A12,1,1,1)", 151.5),
    ] {
        let mut workbook = cached(formula, value);
        rebuild_and_recalc_all(&mut workbook, None);
        assert_eq!(
            workbook.value(SheetId(0), CellRef::new(0, 5)),
            number(value),
            "{formula}"
        );
    }
}
