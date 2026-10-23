use xlsx_model::SheetId;
use xlsx_ops::{Op, Provenance, Transaction, UndoStack};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn saved_filter_after(op: Op) -> String {
    let worksheet = format!(
        r#"<worksheet xmlns="{NS}"><sheetData/><autoFilter ref="A1:C10"><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter></worksheet>"#
    );
    let parts = vec![
        ("xl/workbook.xml".into(), br#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
        ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), worksheet.into_bytes()),
    ];
    let parsed = xlsx_parse::parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    UndoStack::new()
        .commit(&mut workbook, &Transaction::new(vec![op], Provenance::User))
        .unwrap();
    let saved = xlsx_parse::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0)],
        &[Default::default()],
        xlsx_parse::SaveEdits {
            changed: true,
            moved_references: true,
        },
    )
    .unwrap();
    let sheet = saved
        .into_iter()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap()
        .1;
    let sheet = String::from_utf8(sheet.to_vec()).unwrap();
    let start = sheet.find("<autoFilter").unwrap();
    let end = sheet.find("</autoFilter>").unwrap() + "</autoFilter>".len();
    sheet[start..end].to_string()
}

#[test]
fn deleting_a_column_inside_a_filter_moves_later_criteria_left() {
    assert_eq!(
        saved_filter_after(Op::DeleteCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        }),
        r#"<autoFilter ref="A1:B10"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter>"#
    );
}

#[test]
fn inserting_a_column_inside_a_filter_moves_later_criteria_right() {
    assert_eq!(
        saved_filter_after(Op::InsertCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        }),
        r#"<autoFilter ref="A1:D10"><filterColumn colId="3"><filters><filter val="x"/></filters></filterColumn></autoFilter>"#
    );
}
