//! CELL against microsoft's info-type table, including the "filename" idiom
//! real workbooks use to print their own sheet name.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{
    Alignment, Cell, CellRef, CellValue, ErrorValue, HAlign, Protection, Sheet, Workbook, Xf,
};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn book() -> Workbook {
    Workbook {
        sheets: vec![Sheet::new("Sheet1"), Sheet::new("My Data")],
        ..Workbook::default()
    }
}

fn style(book: &mut Workbook, xf: Xf) -> u32 {
    book.styles.cell_xfs.push(xf);
    book.styles.cell_xfs.len() as u32 - 1
}

fn formatted(book: &mut Workbook, code: &str) -> u32 {
    let id = 164 + book.styles.num_fmts.len() as u16;
    book.styles.num_fmts.push((id, code.into()));
    style(
        book,
        Xf {
            num_fmt_id: Some(id),
            ..Xf::default()
        },
    )
}

fn put(book: &mut Workbook, sheet: usize, address: &str, cell: Cell) {
    book.sheets[sheet].set_cell(CellRef::parse_a1(address).unwrap(), cell);
}

fn put_value(book: &mut Workbook, address: &str, value: CellValue, style: Option<u32>) {
    put(
        book,
        0,
        address,
        Cell {
            value,
            style,
            ..Cell::default()
        },
    );
}

fn put_formula(book: &mut Workbook, address: &str, formula: &str) {
    put(
        book,
        0,
        address,
        Cell {
            formula: Some(formula.into()),
            ..Cell::default()
        },
    );
}

fn value(book: &Workbook, address: &str) -> CellValue {
    book.sheets[0]
        .cell(CellRef::parse_a1(address).unwrap())
        .unwrap()
        .value
        .clone()
}

/// evaluate each formula against `book`, one result cell per formula.
fn results(mut book: Workbook, formulas: &[&str]) -> Vec<CellValue> {
    for (i, formula) in formulas.iter().enumerate() {
        put_formula(&mut book, &format!("Z{}", i + 1), formula);
    }
    rebuild_and_recalc_all(&mut book, None);
    (0..formulas.len())
        .map(|i| value(&book, &format!("Z{}", i + 1)))
        .collect()
}

fn check(book: Workbook, cases: &[(&str, CellValue)]) {
    let formulas: Vec<&str> = cases.iter().map(|(formula, _)| *formula).collect();
    for ((formula, want), got) in cases.iter().zip(results(book, &formulas)) {
        assert_eq!(&got, want, "{formula}");
    }
}

#[test]
fn position_and_contents() {
    let mut book = book();
    put_value(&mut book, "A1", number(42.5), None);
    put_value(&mut book, "A2", text("label"), None);
    put_value(&mut book, "A3", CellValue::Bool { value: true }, None);
    check(
        book,
        &[
            ("CELL(\"row\",B7)", number(7.0)),
            ("CELL(\"col\",B7)", number(2.0)),
            ("CELL(\"ROW\",C3:E9)", number(3.0)),
            ("CELL(\"address\",B7)", text("$B$7")),
            ("CELL(\"address\",AB12:AC20)", text("$AB$12")),
            ("CELL(\"contents\",A1)", number(42.5)),
            ("CELL(\"contents\",A2)", text("label")),
            ("CELL(\"type\",A1)", text("v")),
            ("CELL(\"type\",A2)", text("l")),
            ("CELL(\"type\",A3)", text("v")),
            ("CELL(\"type\",H40)", text("b")),
            (
                "ISNUMBER(CELL(\"contents\",H40))",
                CellValue::Bool { value: true },
            ),
            ("CELL(\"contents\",H40)&\"\"", text("0")),
        ],
    );
}

#[test]
fn prefix_follows_text_alignment() {
    let mut book = book();
    let aligned = |book: &mut Workbook, h| {
        style(
            book,
            Xf {
                alignment: Some(Alignment {
                    h: Some(h),
                    ..Alignment::default()
                }),
                ..Xf::default()
            },
        )
    };
    let right = aligned(&mut book, HAlign::Right);
    let center = aligned(&mut book, HAlign::Center);
    let fill = aligned(&mut book, HAlign::Fill);
    let justify = aligned(&mut book, HAlign::Justify);
    put_value(&mut book, "A1", text("plain"), None);
    put_value(&mut book, "A2", text("right"), Some(right));
    put_value(&mut book, "A3", text("center"), Some(center));
    put_value(&mut book, "A4", text("fill"), Some(fill));
    put_value(&mut book, "A5", text("justify"), Some(justify));
    put_value(&mut book, "A6", number(1.0), Some(right));
    check(
        book,
        &[
            ("CELL(\"prefix\",A1)", text("'")),
            ("CELL(\"prefix\",A2)", text("\"")),
            ("CELL(\"prefix\",A3)", text("^")),
            ("CELL(\"prefix\",A4)", text("\\")),
            ("CELL(\"prefix\",A5)", text("")),
            ("CELL(\"prefix\",A6)", text("")),
            ("CELL(\"prefix\",A7)", text("")),
        ],
    );
}

#[test]
fn protect_reads_the_locked_flag() {
    let mut book = book();
    let unlocked = style(
        &mut book,
        Xf {
            protection: Some(Protection {
                locked: false,
                hidden: false,
            }),
            ..Xf::default()
        },
    );
    put_value(&mut book, "A1", number(1.0), Some(unlocked));
    put_value(&mut book, "A2", number(1.0), None);
    check(
        book,
        &[
            ("CELL(\"protect\",A1)", number(0.0)),
            ("CELL(\"protect\",A2)", number(1.0)),
            ("CELL(\"protect\",A3)", number(1.0)),
        ],
    );
}

#[test]
fn format_color_and_parentheses_follow_the_number_format() {
    let mut book = book();
    let mut cases = Vec::new();
    for (row, (code, format, color, parentheses)) in [
        ("General", "G", 0.0, 0.0),
        ("#,##0.00", ",2", 0.0, 0.0),
        ("$#,##0_);[Red]($#,##0)", "C0-", 1.0, 0.0),
        ("0.0%", "P1", 0.0, 0.0),
        ("(0)", "F0()", 0.0, 1.0),
        ("d-mmm-yy", "D1", 0.0, 0.0),
        ("h:mm:ss", "D8", 0.0, 0.0),
    ]
    .into_iter()
    .enumerate()
    {
        let address = format!("A{}", row + 1);
        let id = formatted(&mut book, code);
        put_value(&mut book, &address, number(1.0), Some(id));
        cases.push((format!("CELL(\"format\",{address})"), text(format)));
        cases.push((format!("CELL(\"color\",{address})"), number(color)));
        cases.push((
            format!("CELL(\"parentheses\",{address})"),
            number(parentheses),
        ));
    }
    let builtin = style(
        &mut book,
        Xf {
            num_fmt_id: Some(14),
            ..Xf::default()
        },
    );
    put_value(&mut book, "B1", number(45_000.0), Some(builtin));
    cases.push(("CELL(\"format\",B1)".to_string(), text("D4")));
    cases.push(("CELL(\"format\",B2)".to_string(), text("G")));
    let borrowed: Vec<(&str, CellValue)> = cases
        .iter()
        .map(|(formula, want)| (formula.as_str(), want.clone()))
        .collect();
    check(book, &borrowed);
}

/// the idiom real workbooks use to show their own sheet name. excel's
/// "filename" is the saved path the model does not carry, so a cell keeps
/// the value excel cached.
#[test]
fn filename_keeps_the_value_excel_cached() {
    let mut book = book();
    let idiom = "MID(CELL(\"filename\",A1),FIND(\"]\",CELL(\"filename\",A1))+1,255)";
    put(
        &mut book,
        0,
        "C4",
        Cell {
            formula: Some(idiom.into()),
            value: text("Uploads"),
            ..Cell::default()
        },
    );
    put_formula(&mut book, "C5", idiom);
    put_formula(&mut book, "C6", "CELL(\"filename\",A1)");
    put_formula(&mut book, "C7", "LEN(CELL(\"filename\"))");
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "C4"), text("Uploads"));
    assert_eq!(
        value(&book, "C5"),
        CellValue::Error {
            value: ErrorValue::Value
        }
    );
    assert_eq!(value(&book, "C6"), text(""));
    assert_eq!(value(&book, "C7"), number(0.0));
}

/// excel prefixes another sheet's address with the book name, which the model
/// does not carry: a cached value stays, an uncached one names the sheet.
#[test]
fn address_on_another_sheet() {
    let mut book = book();
    put(
        &mut book,
        0,
        "C1",
        Cell {
            formula: Some("CELL(\"address\",'My Data'!C3)".into()),
            value: text("[Book1.xlsx]My Data!$C$3"),
            ..Cell::default()
        },
    );
    put_formula(&mut book, "C2", "CELL(\"address\",'My Data'!C3)");
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "C1"), text("[Book1.xlsx]My Data!$C$3"));
    assert_eq!(value(&book, "C2"), text("'My Data'!$C$3"));
}

/// without a reference excel reads the cell changed last; uncached, the
/// formula's own cell stands in.
#[test]
fn a_missing_reference_reads_the_formula_cell() {
    let book = book();
    assert_eq!(results(book, &["1", "2", "CELL(\"row\")"])[2], number(3.0));
}

#[test]
fn errors() {
    let book = book();
    let value_error = CellValue::Error {
        value: ErrorValue::Value,
    };
    check(
        book,
        &[
            ("CELL(\"bogus\",A1)", value_error.clone()),
            ("CELL(\"row\",5)", value_error.clone()),
            ("CELL(1,A1)", value_error.clone()),
            ("CELL(\"width\",A1)", value_error.clone()),
            ("CELL()", value_error.clone()),
            ("CELL(\"row\",A1,A2)", value_error),
            (
                "CELL(\"row\",1/0)",
                CellValue::Error {
                    value: ErrorValue::Div0,
                },
            ),
        ],
    );
}
