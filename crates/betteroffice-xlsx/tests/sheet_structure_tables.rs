//! Tables stay on their own sheets when sheets are inserted, deleted or
//! moved, so structured references keep reading the same cells. Deleting a
//! sheet takes its tables with it, and a structured reference to one of them
//! becomes `#REF!`, as Excel saves it. Undo brings the tables and the
//! formulas back.

use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, Op, Workbook};
use xlsx_model::{CellProvider, ErrorValue, SheetId};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#
        ));
    }
    xml + "</Relationships>"
}

fn worksheet(rows: &str, tail: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData>{rows}</sheetData>{tail}</worksheet>"#
    )
}

fn table(id: u32, name: &str, range: &str, column: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="{MAIN}" id="{id}" name="{name}" displayName="{name}" ref="{range}"><tableColumns count="1"><tableColumn id="1" name="{column}"/></tableColumns></table>"#
    )
}

const TABLE_PART: &str = r#"<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>"#;

/// `First` (no table), `Data` with table `Amounts` (1, 2), `Later` with
/// table `Prices` (10, 20), and `Other`, whose A1 sums `Amounts[Amt]` and
/// whose A2 sums `Prices[Price]`.
fn book() -> Vec<u8> {
    book_with("")
}

/// [`book`] with `extra` appended to `Other`'s rows.
fn book_with(extra: &str) -> Vec<u8> {
    let other = format!(
        r#"<row r="1"><c r="A1"><f>SUM(Amounts[Amt])</f><v>3</v></c></row><row r="2"><c r="A2"><f>SUM(Prices[Price])</f><v>30</v></c></row>{extra}"#
    );
    let parts = vec![
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>"#.to_owned(),
        ),
        ("_rels/.rels", relationships(&[("rId1", "officeDocument", "xl/workbook.xml")])),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/><sheet name="Later" sheetId="3" r:id="rId3"/><sheet name="Other" sheetId="4" r:id="rId4"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "worksheet", "worksheets/sheet3.xml"),
                ("rId4", "worksheet", "worksheets/sheet4.xml"),
            ]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            worksheet(r#"<row r="1"><c r="A1"><v>100</v></c></row><row r="2"><c r="A2"><v>200</v></c></row><row r="3"><c r="A3"><v>300</v></c></row>"#, ""),
        ),
        (
            "xl/worksheets/sheet2.xml",
            worksheet(
                r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Amt</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c></row>"#,
                TABLE_PART,
            ),
        ),
        ("xl/worksheets/_rels/sheet2.xml.rels", relationships(&[("rIdT", "table", "../tables/table1.xml")])),
        (
            "xl/worksheets/sheet3.xml",
            worksheet(
                r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Price</t></is></c></row><row r="2"><c r="A2"><v>10</v></c></row><row r="3"><c r="A3"><v>20</v></c></row>"#,
                TABLE_PART,
            ),
        ),
        ("xl/worksheets/_rels/sheet3.xml.rels", relationships(&[("rIdT", "table", "../tables/table2.xml")])),
        (
            "xl/worksheets/sheet4.xml",
            worksheet(&other, ""),
        ),
        ("xl/tables/table1.xml", table(1, "Amounts", "A1:A3", "Amt")),
        ("xl/tables/table2.xml", table(2, "Prices", "A1:A3", "Price")),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn at(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn remove(workbook: &mut Workbook, index: usize) {
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index }],
            CalculationOptions::default(),
        )
        .unwrap();
}

/// Each table's name with the name of the sheet it now sits on.
fn tables(workbook: &Workbook) -> Vec<(String, Option<String>)> {
    let model = workbook.model();
    model
        .tables
        .iter()
        .map(|table| {
            let sheet = model
                .sheets
                .get(table.sheet.0 as usize)
                .map(|sheet| sheet.name.clone());
            (table.name.clone(), sheet)
        })
        .collect()
}

fn other(workbook: &Workbook) -> SheetId {
    SheetId(workbook.model().sheets.len() as u32 - 1)
}

fn formula(workbook: &Workbook, address: &str) -> Option<String> {
    let sheet = &workbook.model().sheets[other(workbook).0 as usize];
    sheet
        .cell(at(address))
        .and_then(|cell| cell.formula.clone())
}

fn value(workbook: &Workbook, address: &str) -> CellValue {
    workbook.model().value(other(workbook), at(address))
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn named(entries: &[(&str, &str)]) -> Vec<(String, Option<String>)> {
    entries
        .iter()
        .map(|(table, sheet)| ((*table).to_owned(), Some((*sheet).to_owned())))
        .collect()
}

/// Deleting the sheet that holds `Amounts` drops the table, and the
/// structured reference to it becomes `#REF!` and evaluates to `#REF!`;
/// `Prices`, on a later sheet, stays where it was and keeps its sum.
#[test]
fn deleting_a_tables_sheet_drops_the_table_and_its_references_become_ref() {
    let mut workbook = Workbook::open(&book()).unwrap();
    assert_eq!(
        tables(&workbook),
        named(&[("Amounts", "Data"), ("Prices", "Later")])
    );
    remove(&mut workbook, 1);
    assert_eq!(tables(&workbook), named(&[("Prices", "Later")]));
    assert_eq!(formula(&workbook, "A1").as_deref(), Some("SUM(#REF!)"));
    assert_eq!(
        value(&workbook, "A1"),
        CellValue::Error {
            value: ErrorValue::Ref
        }
    );
    assert_eq!(
        formula(&workbook, "A2").as_deref(),
        Some("SUM(Prices[Price])")
    );
    assert_eq!(value(&workbook, "A2"), number(30.0));
}

/// Deleting a sheet before both tables leaves each table on its own sheet,
/// so every structured reference reads the cells it read before.
#[test]
fn deleting_an_earlier_sheet_keeps_every_table_on_its_own_sheet() {
    let mut workbook = Workbook::open(&book()).unwrap();
    remove(&mut workbook, 0);
    assert_eq!(
        tables(&workbook),
        named(&[("Amounts", "Data"), ("Prices", "Later")])
    );
    assert_eq!(value(&workbook, "A1"), number(3.0));
    assert_eq!(value(&workbook, "A2"), number(30.0));
}

/// Inserting a sheet before the tables leaves each on its own sheet too.
#[test]
fn inserting_an_earlier_sheet_keeps_every_table_on_its_own_sheet() {
    let mut workbook = Workbook::open(&book()).unwrap();
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "New".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        tables(&workbook),
        named(&[("Amounts", "Data"), ("Prices", "Later")])
    );
    assert_eq!(value(&workbook, "A1"), number(3.0));
    assert_eq!(value(&workbook, "A2"), number(30.0));
}

/// Moving sheets around, the tables' own and others, leaves each table on
/// its own sheet.
#[test]
fn moving_sheets_keeps_every_table_on_its_own_sheet() {
    for (from, to) in [(2, 0), (0, 3), (1, 2), (3, 1)] {
        let mut workbook = Workbook::open(&book()).unwrap();
        workbook
            .apply_ops(
                vec![Op::MoveSheet { from, to }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert_eq!(
            tables(&workbook),
            named(&[("Amounts", "Data"), ("Prices", "Later")]),
            "move {from} to {to}"
        );
        let other = workbook
            .model()
            .sheets
            .iter()
            .position(|sheet| sheet.name == "Other")
            .unwrap();
        let read = |address: &str| workbook.model().value(SheetId(other as u32), at(address));
        assert_eq!(read("A1"), number(3.0), "move {from} to {to}");
        assert_eq!(read("A2"), number(30.0), "move {from} to {to}");
    }
}

/// The save writes what the workbook now holds: no part for the dropped
/// table, and on reopening the same tables on the same sheets.
#[test]
fn a_saved_deletion_reopens_with_the_same_tables_on_the_same_sheets() {
    for index in [0, 1] {
        let mut workbook = Workbook::open(&book()).unwrap();
        remove(&mut workbook, index);
        let expected = tables(&workbook);
        let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
        assert_eq!(tables(&reopened), expected, "delete #{index}");
        assert_eq!(
            formula(&reopened, "A1"),
            formula(&workbook, "A1"),
            "delete #{index}"
        );
    }
}

/// A reference into another workbook beside a structured reference is
/// rewritten around; a formula naming the dropped table in text the rewrite
/// cannot read refuses the deletion, leaving the workbook as it was.
#[test]
fn a_reference_to_the_dropped_table_that_cannot_be_rewritten_refuses_the_deletion() {
    let mut workbook = Workbook::open(&book_with(
        r#"<row r="3"><c r="A3"><f>[1]Rates!$A$1*SUM(Amounts[Amt])</f><v>0</v></c></row>"#,
    ))
    .unwrap();
    remove(&mut workbook, 1);
    assert_eq!(
        formula(&workbook, "A3").as_deref(),
        Some("[1]Rates!$A$1*SUM(#REF!)")
    );

    let mut workbook = Workbook::open(&book_with(
        r#"<row r="3"><c r="A3"><f>[Rates.xlsx]Rates!$A$1*SUM(Amounts[Amt])</f><v>0</v></c></row>"#,
    ))
    .unwrap();
    let before = workbook.model().clone();
    let error = workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        error.to_string().contains("cannot be safely rewritten"),
        "{error}"
    );
    assert_eq!(workbook.model(), &before);
}

/// Undo brings back the deleted sheet's table and the formula text.
#[test]
fn undo_restores_the_table_and_the_structured_reference() {
    let mut workbook = Workbook::open(&book()).unwrap();
    let opened = workbook.model().clone();
    remove(&mut workbook, 1);
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().tables, opened.tables);
    assert_eq!(
        formula(&workbook, "A1").as_deref(),
        Some("SUM(Amounts[Amt])")
    );
    assert_eq!(value(&workbook, "A1"), number(3.0));
}
