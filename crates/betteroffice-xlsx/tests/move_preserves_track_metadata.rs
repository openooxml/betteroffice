//! Moving tracks must not deserialize and reserialize the sheet: row and
//! column attributes the model does not read (outlines, custom flags) have to
//! travel with their track and stay in the saved worksheet.
#![cfg(feature = "unstable-track-moves")]

use betteroffice_xlsx::{CalculationOptions, SheetId, Workbook};
use xlsx_ops::Op;

fn source() -> Vec<u8> {
    let workbook = r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rels = r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="12" customWidth="1" outlineLevel="1"/><col min="2" max="2" width="8" customWidth="1"/></cols><sheetData><row r="1" outlineLevel="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c></row><row r="2" collapsed="1" outlineLevel="2"><c r="A2"><v>11</v></c></row><row r="3"><c r="A3"><v>12</v></c></row></sheetData></worksheet>"#;
    ooxml_opc::rezip_parts(&[
        ("xl/workbook.xml".into(), workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".into(),
            rels.as_bytes().to_vec(),
        ),
        ("xl/worksheets/sheet1.xml".into(), sheet.as_bytes().to_vec()),
    ])
    .unwrap()
}

fn part(bytes: &[u8], path: &str) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(candidate, _)| candidate == path)
            .unwrap()
            .1,
    )
    .unwrap()
}

#[test]
fn a_column_move_keeps_column_outline_attributes_with_their_column() {
    let mut book = Workbook::open(&source()).unwrap();
    book.apply_ops(
        vec![Op::MoveCols {
            sheet: SheetId(0),
            col: 0,
            count: 1,
            before: 2,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = part(&book.save().unwrap(), "xl/worksheets/sheet1.xml");
    let outline = saved
        .find(r#"min="2" max="2" width="12" customWidth="1" outlineLevel="1""#)
        .or_else(|| saved.find(r#"outlineLevel="1" min="2" max="2""#))
        .unwrap_or_else(|| panic!("the outlined column did not move to column 2: {saved}"));
    assert!(outline > 0);
    assert!(
        saved.contains(r#"min="1" max="1" width="8" customWidth="1""#),
        "the untouched column keeps its width at column 1: {saved}"
    );
}

#[test]
fn a_row_move_keeps_row_attributes_and_moves_them_with_the_row() {
    let mut book = Workbook::open(&source()).unwrap();
    book.apply_ops(
        vec![Op::MoveRows {
            sheet: SheetId(0),
            row: 0,
            count: 1,
            before: 3,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = part(&book.save().unwrap(), "xl/worksheets/sheet1.xml");
    let moved = saved
        .find(r#"r="3" outlineLevel="1""#)
        .or_else(|| saved.find(r#"outlineLevel="1" r="3""#))
        .unwrap_or_else(|| panic!("the moved row lost its outline: {saved}"));
    assert!(moved > 0);
    assert!(
        saved.contains(r#"r="1" collapsed="1" outlineLevel="2""#)
            || saved.contains(r#"r="1" outlineLevel="2" collapsed="1""#),
        "the row the move passed keeps its attributes: {saved}"
    );
    assert!(
        saved.contains(r#"<c r="A3"><v>10</v></c>"#),
        "the moved row carries its cell: {saved}"
    );
}
