use betteroffice_xlsx::{CalculationOptions, Error, SheetId, Workbook};
use xlsx_model::CellRange;
use xlsx_ops::Op;

fn bytes(styles: Option<&str>) -> Vec<u8> {
    let mut rels = r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#.to_string();
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        ("xl/workbook.xml".into(), br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), br#"<worksheet><sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData></worksheet>"#.to_vec()),
    ];
    if let Some(styles) = styles {
        rels.push_str(r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#);
        parts.push(("xl/styles.xml".into(), format!("<styleSheet>{styles}</styleSheet>").into_bytes()));
    }
    parts.push(("xl/_rels/workbook.xml.rels".into(), format!("<Relationships>{rels}</Relationships>").into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn add(kind: &str, operator: Option<&str>, formulas: &[&str]) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:A5").unwrap(),
        kind: kind.into(),
        operator: operator.map(str::to_owned),
        formulas: formulas.iter().map(|f| f.to_string()).collect(),
        format: xlsx_model::DifferentialFormat { font_color: Some("#FF0000".into()), bold: Some(false), fill: None },
    }
}

#[test]
fn add_list_delete_save_reopen_and_undo_redo() {
    let mut wb = Workbook::open(&bytes(Some(r#"<cellXfs count="1"><xf/></cellXfs>"#))).unwrap();
    wb.apply_ops(vec![add("cellIs", Some("greaterThan"), &["5"])], CalculationOptions::default()).unwrap();
    wb.apply_ops(vec![add("expression", None, &["A1>5"])], CalculationOptions::default()).unwrap();
    assert_eq!(wb.model().sheets[0].conditional_formats.iter().flat_map(|g| &g.rules).count(), 2);
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet = String::from_utf8(parts.iter().find(|(p, _)| p == "xl/worksheets/sheet1.xml").unwrap().1.clone()).unwrap();
    assert!(sheet.contains("conditionalFormatting"), "{sheet}");
    assert!(sheet.contains(r#"dxfId="0""#), "{sheet}");
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.model().sheets[0].conditional_formats.iter().flat_map(|g| &g.rules).count(), 2);
    let id = reopened.model().sheets[0].conditional_formats[0].rules[0].id.clone();
    let mut reopened = reopened;
    reopened.apply_ops(vec![Op::RemoveConditionalFormat { sheet: SheetId(0), id }], CalculationOptions::default()).unwrap();
    assert_eq!(reopened.model().sheets[0].conditional_formats.iter().flat_map(|g| &g.rules).count(), 1);
    wb.undo(CalculationOptions::default()).unwrap();
    assert_eq!(wb.model().sheets[0].conditional_formats.iter().flat_map(|g| &g.rules).count(), 1);
    wb.redo(CalculationOptions::default()).unwrap();
    assert_eq!(wb.model().sheets[0].conditional_formats.iter().flat_map(|g| &g.rules).count(), 2);
}

#[test]
fn invalid_batch_is_atomic() {
    let mut wb = Workbook::open(&bytes(None)).unwrap();
    let result = wb.apply_ops(vec![add("cellIs", Some("greaterThan"), &["5"]), add("cellIs", Some("between"), &["1"])], CalculationOptions::default());
    assert!(result.is_err());
    assert!(wb.model().sheets[0].conditional_formats.is_empty());
}

#[test]
fn collaborative_mode_refuses_conditional_format_ops() {
    let mut wb = Workbook::open_collaborative(&bytes(None), 42).unwrap();
    let result = wb.apply_ops(vec![add("cellIs", Some("greaterThan"), &["5"])], CalculationOptions::default());
    assert!(matches!(result, Err(Error::CollaborativeStructureOperation)), "{result:?}");
}
