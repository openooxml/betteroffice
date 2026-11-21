use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRef, CellValue, DateSystem, ErrorValue, Sheet, Workbook};

const NOW_2026_11_01_NOON: f64 = 46_327.5;

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn evaluate(system: DateSystem, formulas: &[&str]) -> Vec<CellValue> {
    let mut book = Workbook {
        sheets: vec![Sheet::new("Sheet1")],
        date_system: system,
        ..Workbook::default()
    };
    for (row, formula) in formulas.iter().enumerate() {
        book.sheets[0].set_cell(
            CellRef::new(row as u32, 0),
            Cell {
                formula: Some((*formula).into()),
                ..Cell::default()
            },
        );
    }
    rebuild_and_recalc_all(&mut book, Some(NOW_2026_11_01_NOON));
    (0..formulas.len())
        .map(|row| {
            book.sheets[0]
                .cell(CellRef::new(row as u32, 0))
                .unwrap()
                .value
                .clone()
        })
        .collect()
}

fn assert_cases(system: DateSystem, cases: &[(&str, CellValue)]) {
    let formulas: Vec<&str> = cases.iter().map(|(formula, _)| *formula).collect();
    let values = evaluate(system, &formulas);
    for ((formula, expected), actual) in cases.iter().zip(values) {
        assert_eq!(&actual, expected, "{formula} in a {system:?} workbook");
    }
}

#[test]
fn serial_producing_functions_count_from_the_1904_epoch() {
    assert_cases(
        DateSystem::V1904,
        &[
            ("DATE(2026,11,1)", number(44_865.0)),
            ("DATE(1904,1,1)", number(0.0)),
            ("DATE(4,1,1)", number(0.0)),
            ("DATE(1903,12,31)", error(ErrorValue::Num)),
            ("DATEVALUE(\"11/1/2026\")", number(44_865.0)),
            ("DATEVALUE(\"1/1/1903\")", error(ErrorValue::Value)),
            ("EDATE(44865,1)", number(44_895.0)),
            ("EDATE(\"11/1/2026\",0)", number(44_865.0)),
            ("EDATE(0,-1)", error(ErrorValue::Num)),
            ("EOMONTH(44865,0)", number(44_894.0)),
            ("WORKDAY(44865,1)", number(44_866.0)),
            ("WORKDAY.INTL(44866,-1)", number(44_863.0)),
            ("TODAY()", number(44_865.0)),
            ("NOW()", number(44_865.5)),
        ],
    );
}

#[test]
fn serial_reading_functions_count_from_the_1904_epoch() {
    assert_cases(
        DateSystem::V1904,
        &[
            ("YEAR(44865)", number(2026.0)),
            ("MONTH(44865)", number(11.0)),
            ("DAY(44865)", number(1.0)),
            ("YEAR(0)", number(1904.0)),
            ("DAY(0)", number(1.0)),
            ("YEAR(-1)", error(ErrorValue::Num)),
            ("WEEKDAY(0)", number(6.0)),
            ("WEEKDAY(44865)", number(1.0)),
            ("WEEKNUM(44865)", number(45.0)),
            ("ISOWEEKNUM(44865)", number(44.0)),
            ("DATEDIF(0,44865,\"Y\")", number(122.0)),
            ("DATEDIF(44865,44894,\"MD\")", number(29.0)),
            ("DAYS(\"11/1/2026\",0)", number(44_865.0)),
            ("NETWORKDAYS(44865,44871)", number(5.0)),
            ("NETWORKDAYS(44865,44871,44866)", number(4.0)),
            ("YEARFRAC(DATE(2026,1,1),DATE(2026,7,1),0)", number(0.5)),
            ("TEXT(DATE(2026,11,1),\"yyyy-mm-dd\")", text("2026-11-01")),
            ("YEAR(TODAY())", number(2026.0)),
        ],
    );
}

#[test]
fn the_same_formulas_keep_their_1900_serials() {
    assert_cases(
        DateSystem::V1900,
        &[
            ("DATE(2026,11,1)", number(46_327.0)),
            ("DATEVALUE(\"11/1/2026\")", number(46_327.0)),
            ("EDATE(46327,1)", number(46_357.0)),
            ("EOMONTH(46327,0)", number(46_356.0)),
            ("WORKDAY(46327,1)", number(46_328.0)),
            ("TODAY()", number(46_327.0)),
            ("NOW()", number(46_327.5)),
            ("YEAR(0)", number(1900.0)),
            ("DAY(0)", number(0.0)),
            ("WEEKDAY(46327)", number(1.0)),
            ("ISOWEEKNUM(46327)", number(44.0)),
            ("DAYS(\"11/1/2026\",0)", number(46_327.0)),
            ("TEXT(DATE(2026,11,1),\"yyyy-mm-dd\")", text("2026-11-01")),
        ],
    );
}

#[test]
fn serials_past_any_date_are_num_errors_in_a_1904_workbook() {
    assert_cases(
        DateSystem::V1904,
        &[
            ("YEAR(1E+300)", error(ErrorValue::Num)),
            ("WEEKDAY(1E+20)", error(ErrorValue::Num)),
            ("EDATE(1E+20,0)", error(ErrorValue::Num)),
            ("DATEDIF(0,1E+20,\"D\")", error(ErrorValue::Num)),
            ("YEARFRAC(0,1E+20)", error(ErrorValue::Num)),
            ("NETWORKDAYS(0,1E+20)", error(ErrorValue::Num)),
            ("WORKDAY(1E+20,1)", error(ErrorValue::Num)),
            ("NETWORKDAYS(0,1,1E+20)", number(1.0)),
            ("DATE(9999,12,31)", number(2_957_003.0)),
            ("YEAR(2957003)", number(9999.0)),
            ("YEAR(2957004)", error(ErrorValue::Num)),
        ],
    );
}
