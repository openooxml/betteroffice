//! Regression: a table no formula reads must still be guarded.
//!
//! Before this fix, `package_bears_references` listed only
//! ["xl/pivottables/", "xl/pivotcache/", "xl/charts/"], so a package whose only
//! reference-bearing part was `xl/tables/tableN.xml` produced an EMPTY
//! `unpatchable_references`. The structural-op guard is therefore keyed only on
//! a formula reading the table — and with no such formula a row insert inside
//! the table's range was ALLOWED, the cells shifted, and the saved table part
//! kept its old `ref` while the save reported success.
//!
//! Observed before the fix, with no formula in the sheet:
//!     insert=ALLOWED  saved ref=A1:A3  A1=<empty> A2="Amount" A3=2.0 A4=3.0
//! i.e. the table still named A1:A3, covering an empty cell and missing 3.0.
//!
//! The control case (a formula present) was already refused, which is what made
//! the gap easy to miss: the guard looked like it worked.

use betteroffice_xlsx::{CalculationOptions, Workbook};
use ooxml_opc::rezip_parts;
use xlsx_ops::Op;

fn table_package(with_formula: bool) -> Vec<u8> {
    let workbook =
        r#"<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rels = r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let formula = if with_formula {
        r#"<row r="5"><c r="C5"><f>SUM(Sales[Amount])</f></c></row>"#
    } else {
        ""
    };
    let worksheet = format!(
        r#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>{formula}</sheetData></worksheet>"#
    );
    let sheet_rels = r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="/xl/tables/table1.xml"/></Relationships>"#;
    let table = r#"<table id="1" displayName="Sales" ref="A1:A3"><tableColumns count="1"><tableColumn id="1" name="Amount"/></tableColumns></table>"#;

    rezip_parts(&[
        ("xl/workbook.xml".to_string(), workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels.as_bytes().to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            worksheet.into_bytes(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            sheet_rels.as_bytes().to_vec(),
        ),
        ("xl/tables/table1.xml".to_string(), table.as_bytes().to_vec()),
    ])
    .unwrap()
}

fn insert_row() -> Op {
    Op::InsertRows {
        sheet: xlsx_model::SheetId(0),
        at: 0,
        count: 1,
    }
}

/// The case that was broken: no formula reads the table, so nothing else
/// flagged the move. The insert must be refused rather than stranding the ref.
#[test]
fn a_row_insert_is_refused_for_a_table_no_formula_reads() {
    let mut wb = Workbook::open(&table_package(false)).unwrap();

    let error = wb
        .apply_ops(vec![insert_row()], CalculationOptions::default())
        .expect_err("a row insert inside an unguarded table must be refused");

    println!("refused with: {error}");

    // The model must be untouched, so the table still names the right cells.
    assert_eq!(
        wb.model().tables[0].range.to_a1(),
        "A1:A3",
        "the table ref must not be left stale"
    );
    assert_eq!(
        wb.model().sheets[0]
            .cell(xlsx_model::CellRef::parse_a1("A1").unwrap())
            .unwrap()
            .value,
        xlsx_model::CellValue::Text {
            value: "Amount".into()
        },
        "the header must not have moved"
    );
}

/// Control: with a formula reading the table the guard already fired, and it
/// must keep firing.
#[test]
fn control_a_row_insert_is_still_refused_when_a_formula_reads_the_table() {
    let mut wb = Workbook::open(&table_package(true)).unwrap();
    let error = wb
        .apply_ops(vec![insert_row()], CalculationOptions::default())
        .expect_err("still refused");
    println!("control refused with: {error}");
    assert_eq!(wb.model().tables[0].range.to_a1(), "A1:A3");
}
