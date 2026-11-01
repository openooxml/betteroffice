//! UNICHAR, UNICODE, DOLLAR, VALUETOTEXT, ARRAYTOTEXT, ASC, JIS/DBCS, BAHTTEXT,
//! PHONETIC, the byte-counting text functions, ENCODEURL and the `_LONGTEXT`
//! literal wrapper against microsoft's function reference.

use xlsx_calc::rebuild_and_recalc_all;
use xlsx_model::{Cell, CellRef, CellValue, ErrorValue, Sheet, Workbook};

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn error(value: ErrorValue) -> CellValue {
    CellValue::Error { value }
}

fn book() -> Workbook {
    Workbook {
        sheets: vec![Sheet::new("Sheet1")],
        ..Workbook::default()
    }
}

fn put(book: &mut Workbook, address: &str, cell: Cell) {
    book.sheets[0].set_cell(CellRef::parse_a1(address).unwrap(), cell);
}

fn put_value(book: &mut Workbook, address: &str, value: CellValue) {
    put(
        book,
        address,
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn put_formula(book: &mut Workbook, address: &str, formula: &str) {
    put(
        book,
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
fn check_in(mut book: Workbook, cases: &[(&str, CellValue)]) {
    for (i, (formula, _)) in cases.iter().enumerate() {
        put_formula(&mut book, &format!("Z{}", i + 1), formula);
    }
    rebuild_and_recalc_all(&mut book, None);
    for (i, (formula, want)) in cases.iter().enumerate() {
        assert_eq!(&value(&book, &format!("Z{}", i + 1)), want, "{formula}");
    }
}

fn check(cases: &[(&str, CellValue)]) {
    check_in(book(), cases);
}

#[test]
fn unichar_and_unicode() {
    check(&[
        ("UNICHAR(66)", text("B")),
        ("UNICHAR(32)", text(" ")),
        ("UNICHAR(0)", error(ErrorValue::Value)),
        ("_xlfn.UNICHAR(8364)", text("\u{20AC}")),
        ("UNICHAR(128512)", text("\u{1F600}")),
        ("UNICHAR(65.9)", text("A")),
        ("UNICHAR(1114111)", text("\u{10FFFF}")),
        ("UNICHAR(55296)", error(ErrorValue::NA)),
        ("UNICHAR(57343)", error(ErrorValue::NA)),
        ("UNICHAR(1114112)", error(ErrorValue::Value)),
        ("UNICHAR(-1)", error(ErrorValue::Value)),
        ("UNICHAR(\"x\")", error(ErrorValue::Value)),
        ("UNICHAR(1/0)", error(ErrorValue::Div0)),
        ("UNICODE(\" \")", number(32.0)),
        ("UNICODE(\"B\")", number(66.0)),
        ("_xlfn.UNICODE(\"\u{20AC}uro\")", number(8364.0)),
        ("UNICODE(\"\u{1F600}\")", number(128_512.0)),
        ("UNICODE(1)", number(49.0)),
        ("UNICODE(UNICHAR(9731))", number(9731.0)),
        ("UNICODE(\"\")", error(ErrorValue::Value)),
        ("UNICODE(NA())", error(ErrorValue::NA)),
        ("SUM(UNICODE(MID(\"abc\",SEQUENCE(3),1)))", number(294.0)),
    ]);
}

#[test]
fn dollar_writes_the_currency_format() {
    check(&[
        ("DOLLAR(1234.567,2)", text("$1,234.57")),
        ("DOLLAR(1234.567,-2)", text("$1,200")),
        ("DOLLAR(-1234.567,-2)", text("($1,200)")),
        ("DOLLAR(-0.123,4)", text("($0.1230)")),
        ("DOLLAR(99.888)", text("$99.89")),
        ("DOLLAR(1234567.891,0)", text("$1,234,568")),
        ("DOLLAR(2.675)", text("$2.68")),
        ("DOLLAR(-0.001)", text("$0.00")),
        ("DOLLAR(0)", text("$0.00")),
        ("DOLLAR(\"12\")", text("$12.00")),
        ("DOLLAR(1,128)", error(ErrorValue::Value)),
        ("DOLLAR(\"x\")", error(ErrorValue::Value)),
        ("DOLLAR(1/0)", error(ErrorValue::Div0)),
        ("DOLLAR()", error(ErrorValue::Value)),
    ]);
}

/// microsoft's VALUETOTEXT and ARRAYTOTEXT sample: A2:B4 holds TRUE, #VALUE!,
/// 1234.01234, Seattle, Hello and the text 1,123.
fn sample() -> Workbook {
    let mut book = book();
    put_value(&mut book, "A2", CellValue::Bool { value: true });
    put_value(&mut book, "B2", error(ErrorValue::Value));
    put_value(&mut book, "A3", number(1234.01234));
    put_value(&mut book, "B3", text("Seattle"));
    put_value(&mut book, "A4", text("Hello"));
    put_value(&mut book, "B4", text("1,123"));
    book
}

#[test]
fn valuetotext_concise_and_strict() {
    check_in(
        sample(),
        &[
            ("VALUETOTEXT(A2)", text("TRUE")),
            ("VALUETOTEXT(A2,1)", text("TRUE")),
            ("VALUETOTEXT(A3,0)", text("1234.01234")),
            ("VALUETOTEXT(A3,1)", text("1234.01234")),
            ("_xlfn.VALUETOTEXT(A4,0)", text("Hello")),
            ("VALUETOTEXT(A4,1)", text("\"Hello\"")),
            ("VALUETOTEXT(B2,0)", text("#VALUE!")),
            ("VALUETOTEXT(B2,1)", text("#VALUE!")),
            ("VALUETOTEXT(B3,1)", text("\"Seattle\"")),
            ("VALUETOTEXT(B4,1)", text("\"1,123\"")),
            ("VALUETOTEXT(\"a\"\"b\",1)", text("\"a\"\"b\"")),
            ("VALUETOTEXT(C40,1)", text("")),
            ("VALUETOTEXT(1,2)", error(ErrorValue::Value)),
            ("VALUETOTEXT(1,-1)", error(ErrorValue::Value)),
            ("VALUETOTEXT(1,1/0)", error(ErrorValue::Div0)),
            ("VALUETOTEXT()", error(ErrorValue::Value)),
        ],
    );
}

#[test]
fn arraytotext_concise_and_strict() {
    check_in(
        sample(),
        &[
            (
                "ARRAYTOTEXT(A2:B4,0)",
                text("TRUE, #VALUE!, 1234.01234, Seattle, Hello, 1,123"),
            ),
            (
                "_xlfn.ARRAYTOTEXT(A2:B4,1)",
                text("{TRUE,#VALUE!;1234.01234,\"Seattle\";\"Hello\",\"1,123\"}"),
            ),
            ("ARRAYTOTEXT({1,2;3,4},1)", text("{1,2;3,4}")),
            ("ARRAYTOTEXT({1,2;3,4})", text("1, 2, 3, 4")),
            ("ARRAYTOTEXT(_xlfn.SEQUENCE(1,3),1)", text("{1,2,3}")),
            ("ARRAYTOTEXT(5)", text("5")),
            ("ARRAYTOTEXT(200/3)", text("66.6666666666667")),
            (
                "ARRAYTOTEXT(_xlfn.SEQUENCE(1,2)/3,1)",
                text("{0.333333333333333,0.666666666666667}"),
            ),
            ("ARRAYTOTEXT(1/0,1)", error(ErrorValue::Div0)),
            ("ARRAYTOTEXT(A4,1)", text("{\"Hello\"}")),
            ("ARRAYTOTEXT(A2:B4,3)", error(ErrorValue::Value)),
            ("ARRAYTOTEXT()", error(ErrorValue::Value)),
        ],
    );
}

#[test]
fn asc_turns_full_width_into_half_width() {
    check(&[
        (
            "ASC(\"\u{FF25}\u{FF38}\u{FF23}\u{FF25}\u{FF2C}\")",
            text("EXCEL"),
        ),
        (
            "ASC(\"\u{30A8}\u{30AF}\u{30BB}\u{30EB}\")",
            text("\u{FF74}\u{FF78}\u{FF7E}\u{FF99}"),
        ),
        (
            "ASC(\"\u{30AC}\u{30D1}\u{30F4}\")",
            text("\u{FF76}\u{FF9E}\u{FF8A}\u{FF9F}\u{FF73}\u{FF9E}"),
        ),
        (
            "ASC(\"\u{FF21}\u{FF22}\u{FF23}\u{FF11}\u{FF12}\u{FF13}\u{FF01}\u{3000}\u{FFE5}\")",
            text("ABC123! \u{A5}"),
        ),
        (
            "ASC(\"abc\u{6F22}\u{5B57}\u{3042}\")",
            text("abc\u{6F22}\u{5B57}\u{3042}"),
        ),
        ("ASC(\"\u{3002}\u{30FC}\")", text("\u{FF61}\u{FF70}")),
        ("ASC(12)", text("12")),
        ("ASC(1/0)", error(ErrorValue::Div0)),
    ]);
}

#[test]
fn jis_and_dbcs_turn_half_width_into_full_width() {
    check(&[
        (
            "JIS(\"EXCEL\")",
            text("\u{FF25}\u{FF38}\u{FF23}\u{FF25}\u{FF2C}"),
        ),
        (
            "DBCS(\"EXCEL\")",
            text("\u{FF25}\u{FF38}\u{FF23}\u{FF25}\u{FF2C}"),
        ),
        (
            "JIS(\"\u{FF74}\u{FF78}\u{FF7E}\u{FF99}\")",
            text("\u{30A8}\u{30AF}\u{30BB}\u{30EB}"),
        ),
        (
            "JIS(\"\u{FF76}\u{FF9E}\u{FF8A}\u{FF9F}\u{FF73}\u{FF9E}\")",
            text("\u{30AC}\u{30D1}\u{30F4}"),
        ),
        ("JIS(\"\u{FF71}\u{FF9E}\")", text("\u{30A2}\u{309B}")),
        ("JIS(\"a b\")", text("\u{FF41}\u{3000}\u{FF42}")),
        ("JIS(12)", text("\u{FF11}\u{FF12}")),
        ("JIS(\"\u{A5}\")", text("\u{FFE5}")),
        (
            "ASC(JIS(\"Tokyo 2020 \u{FF83}\u{FF9E}\u{FF70}\u{FF80}\"))",
            text("Tokyo 2020 \u{FF83}\u{FF9E}\u{FF70}\u{FF80}"),
        ),
        ("JIS(NA())", error(ErrorValue::NA)),
    ]);
}

#[test]
fn bahttext_spells_thai_baht() {
    check(&[
        ("BAHTTEXT(1234)", text("หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน")),
        ("BAHTTEXT(0)", text("ศูนย์บาทถ้วน")),
        ("BAHTTEXT(1)", text("หนึ่งบาทถ้วน")),
        ("BAHTTEXT(10)", text("สิบบาทถ้วน")),
        ("BAHTTEXT(11)", text("สิบเอ็ดบาทถ้วน")),
        ("BAHTTEXT(21)", text("ยี่สิบเอ็ดบาทถ้วน")),
        ("BAHTTEXT(101)", text("หนึ่งร้อยเอ็ดบาทถ้วน")),
        ("BAHTTEXT(0.5)", text("ห้าสิบสตางค์")),
        ("BAHTTEXT(1.25)", text("หนึ่งบาทยี่สิบห้าสตางค์")),
        ("BAHTTEXT(-15)", text("ลบสิบห้าบาทถ้วน")),
        ("BAHTTEXT(1000000)", text("หนึ่งล้านบาทถ้วน")),
        ("BAHTTEXT(1000001)", text("หนึ่งล้านหนึ่งบาทถ้วน")),
        ("BAHTTEXT(2500000.75)", text("สองล้านห้าแสนบาทเจ็ดสิบห้าสตางค์")),
        ("BAHTTEXT(1E12)", text("หนึ่งล้านล้านบาทถ้วน")),
        ("BAHTTEXT(0.001)", text("ศูนย์บาทถ้วน")),
        ("BAHTTEXT(-0.001)", text("ศูนย์บาทถ้วน")),
        ("BAHTTEXT(0.995)", text("หนึ่งบาทถ้วน")),
        ("BAHTTEXT(\"x\")", error(ErrorValue::Value)),
        ("BAHTTEXT()", error(ErrorValue::Value)),
    ]);
}

/// the model keeps no furigana, so PHONETIC reads text constants as typed.
#[test]
fn phonetic_reads_text_constants() {
    let mut book = book();
    put_value(&mut book, "A1", text("\u{6771}\u{4EAC}"));
    put_value(
        &mut book,
        "A2",
        text("\u{3068}\u{3046}\u{304D}\u{3087}\u{3046}"),
    );
    put_value(&mut book, "A3", number(5.0));
    put_formula(&mut book, "A4", "\"x\"");
    check_in(
        book,
        &[
            ("PHONETIC(A1)", text("\u{6771}\u{4EAC}")),
            (
                "PHONETIC(A1:A4)",
                text("\u{6771}\u{4EAC}\u{3068}\u{3046}\u{304D}\u{3087}\u{3046}"),
            ),
            ("PHONETIC(A3)", text("")),
            ("PHONETIC(A4)", text("")),
            ("PHONETIC(C40)", text("")),
            ("PHONETIC(\"abc\")", error(ErrorValue::NA)),
            ("PHONETIC(1/0)", error(ErrorValue::Div0)),
        ],
    );
}

/// in a language without double-byte characters excel counts one byte per
/// character, so the B functions answer like their character versions.
#[test]
fn byte_functions_count_characters() {
    check(&[
        ("LENB(\"abc\")", number(3.0)),
        ("LENB(\"\u{6771}\u{4EAC}\")", number(2.0)),
        ("LEFTB(\"abcdef\",2)", text("ab")),
        ("RIGHTB(\"abcdef\",2)", text("ef")),
        ("MIDB(\"abcdef\",2,3)", text("bcd")),
        ("FINDB(\"c\",\"abc\")", number(3.0)),
        ("SEARCHB(\"C\",\"abc\")", number(3.0)),
        ("REPLACEB(\"abcdef\",2,3,\"X\")", text("aXef")),
        ("SUMPRODUCT(LENB({\"a\",\"bb\"}))", number(3.0)),
        ("LENB()", error(ErrorValue::Value)),
    ]);
}

#[test]
fn encodeurl_percent_encodes() {
    check(&[
        (
            "ENCODEURL(\"http://contoso.sharepoint.com/Finance/Profit and Loss Statement.xlsx\")",
            text(
                "http%3A%2F%2Fcontoso.sharepoint.com%2FFinance%2FProfit%20and%20Loss%20Statement.xlsx",
            ),
        ),
        ("ENCODEURL(\"a-b_c.d~e\")", text("a-b_c.d~e")),
        ("_xlfn.ENCODEURL(\"\u{E9}\")", text("%C3%A9")),
        ("ENCODEURL(\"\u{20AC}&=?#\")", text("%E2%82%AC%26%3D%3F%23")),
        ("ENCODEURL(12.5)", text("12.5")),
        ("ENCODEURL(1/0)", error(ErrorValue::Div0)),
        ("ENCODEURL()", error(ErrorValue::Value)),
    ]);
}

/// excel stores a string literal past 255 characters wrapped in
/// `_xlfn._LONGTEXT`, split over string arguments of at most 255 characters;
/// it evaluates to their joined text.
#[test]
fn longtext_is_its_literal() {
    let long = "It is encouraged that you enter a code. ".repeat(8);
    let wrapped = format!("_xlfn._LONGTEXT(\"{long}\")");
    let guarded = format!("IF(B6=\"\",_xlfn._LONGTEXT(\"{long}\"),\"\")");
    let split = &"It is encouraged that you enter a code. ".repeat(9)[..353];
    let (head, tail) = split.split_at(255);
    let parts = format!("_xlfn._LONGTEXT(\"{head}\",\"{tail}\")");
    check(&[
        (wrapped.as_str(), text(&long)),
        (guarded.as_str(), text(&long)),
        (parts.as_str(), text(split)),
        ("_xlfn._LONGTEXT(\"ab\",\"cd\")", text("abcd")),
        ("_xlfn._LONGTEXT(\"ab\",1/0)", error(ErrorValue::Div0)),
        ("_xlfn._LONGTEXT(1/0)", error(ErrorValue::Div0)),
    ]);
}

/// the regular-expression and xpath functions need an engine the crate does
/// not carry: a cell keeps the value excel cached, an uncached one is
/// `#NAME?`.
#[test]
fn regex_and_filterxml_keep_the_value_excel_cached() {
    let mut book = book();
    put(
        &mut book,
        "B1",
        Cell {
            formula: Some("_xlfn.REGEXEXTRACT(\"ab12\",\"[0-9]+\")".into()),
            value: text("12"),
            ..Cell::default()
        },
    );
    put(
        &mut book,
        "B2",
        Cell {
            formula: Some("_xlfn.FILTERXML(\"<a><b>1</b></a>\",\"//b\")".into()),
            value: number(1.0),
            ..Cell::default()
        },
    );
    put_formula(&mut book, "B3", "_xlfn.REGEXTEST(\"ab\",\"a\")");
    rebuild_and_recalc_all(&mut book, None);
    assert_eq!(value(&book, "B1"), text("12"));
    assert_eq!(value(&book, "B2"), number(1.0));
    assert_eq!(value(&book, "B3"), error(ErrorValue::Name));
}

/// joining stops once the cell text limit is passed, before reading the
/// arguments after it, as CONCAT does.
#[test]
fn longtext_stops_at_the_text_limit() {
    let mut book = book();
    put_value(&mut book, "A1", text(&"a".repeat(20_000)));
    check_in(
        book,
        &[
            ("_xlfn._LONGTEXT(A1,A1,1/0)", error(ErrorValue::Value)),
            ("_xlfn._LONGTEXT(A1,A1)", error(ErrorValue::Value)),
            ("LEN(_xlfn._LONGTEXT(A1,\"b\"))", number(20_001.0)),
        ],
    );
}

/// excel keeps 15 significant digits, so huge amounts print their first 15
/// digits and zeros; tiny ones round to zero.
#[test]
fn extreme_magnitudes() {
    let mut big = String::from("1");
    for i in (0..200).rev() {
        if i % 3 == 2 {
            big.push(',');
        }
        big.push('0');
    }
    let bahts = |groups: usize| format!("หนึ่งร้อย{}บาทถ้วน", "ล้าน".repeat(groups));
    check(&[
        ("DOLLAR(1E200)", text(&format!("${big}.00"))),
        ("DOLLAR(-1E200)", text(&format!("(${big}.00)"))),
        ("DOLLAR(1E-200)", text("$0.00")),
        ("DOLLAR(-1E-308)", text("$0.00")),
        ("DOLLAR(-0)", text("$0.00")),
        ("LEN(DOLLAR(1E308))", number(415.0)),
        ("DOLLAR(2^53,0)", text("$9,007,199,254,740,990")),
        ("DOLLAR(1E15+0.75)", text("$1,000,000,000,000,000.00")),
        ("DOLLAR(1,1E200)", error(ErrorValue::Value)),
        ("DOLLAR(1234.5,-1E200)", text("$0")),
        ("BAHTTEXT(1E200)", text(&bahts(33))),
        ("BAHTTEXT(1E308)", text(&bahts(51))),
        ("BAHTTEXT(1E-200)", text("ศูนย์บาทถ้วน")),
        ("UNICHAR(1E200)", error(ErrorValue::Value)),
        ("UNICHAR(2^53)", error(ErrorValue::Value)),
        ("VALUETOTEXT(1,1E200)", error(ErrorValue::Value)),
        ("ARRAYTOTEXT(1,-1E200)", error(ErrorValue::Value)),
    ]);
}

/// results past 32,767 characters are `#VALUE!` however they grow.
#[test]
fn oversized_results_are_value_errors() {
    let mut book = book();
    put_value(&mut book, "A1", text(&"\u{E9}".repeat(20_000)));
    put_value(&mut book, "A2", text(&"\"".repeat(20_000)));
    put_value(&mut book, "A3", text(&"a".repeat(20_000)));
    put_value(&mut book, "A4", text(&"a".repeat(20_000)));
    put_value(&mut book, "B1", text(&"\u{FF71}".repeat(20_000)));
    check_in(
        book,
        &[
            ("ENCODEURL(A1)", error(ErrorValue::Value)),
            ("VALUETOTEXT(A2,1)", error(ErrorValue::Value)),
            ("LEN(VALUETOTEXT(A2,0))", number(20_000.0)),
            ("ARRAYTOTEXT(A3:A4)", error(ErrorValue::Value)),
            ("PHONETIC(A3:A4)", error(ErrorValue::Value)),
            ("LEN(JIS(B1))", number(20_000.0)),
            ("LEN(ASC(JIS(A3)))", number(20_000.0)),
        ],
    );
}
