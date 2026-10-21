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
        parts.push((
            "xl/styles.xml".into(),
            format!("<styleSheet>{styles}</styleSheet>").into_bytes(),
        ));
    }
    parts.push((
        "xl/_rels/workbook.xml.rels".into(),
        format!("<Relationships>{rels}</Relationships>").into_bytes(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn add(kind: &str, operator: Option<&str>, formulas: &[&str]) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:A5").unwrap(),
        kind: kind.into(),
        operator: operator.map(str::to_owned),
        formulas: formulas.iter().map(|f| f.to_string()).collect(),
        format: xlsx_model::DifferentialFormat {
            font_color: Some("#FF0000".into()),
            bold: Some(false),
            fill: None,
        },
    }
}

#[test]
fn add_list_delete_save_reopen_and_undo_redo() {
    let mut wb = Workbook::open(&bytes(Some(r#"<cellXfs count="1"><xf/></cellXfs>"#))).unwrap();
    wb.apply_ops(
        vec![add("cellIs", Some("greaterThan"), &["5"])],
        CalculationOptions::default(),
    )
    .unwrap();
    wb.apply_ops(
        vec![add("expression", None, &["A1>5"])],
        CalculationOptions::default(),
    )
    .unwrap();
    assert_eq!(
        wb.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        2
    );
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet = String::from_utf8(
        parts
            .iter()
            .find(|(p, _)| p == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert!(sheet.contains("conditionalFormatting"), "{sheet}");
    assert!(sheet.contains(r#"dxfId="0""#), "{sheet}");
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        2
    );
    let id = reopened.model().sheets[0].conditional_formats[0].rules[0]
        .id
        .clone();
    let mut reopened = reopened;
    reopened
        .apply_ops(
            vec![Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        1
    );
    wb.undo(CalculationOptions::default()).unwrap();
    assert_eq!(
        wb.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        1
    );
    wb.redo(CalculationOptions::default()).unwrap();
    assert_eq!(
        wb.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        2
    );
}

#[test]
fn invalid_batch_is_atomic() {
    let mut wb = Workbook::open(&bytes(None)).unwrap();
    let result = wb.apply_ops(
        vec![
            add("cellIs", Some("greaterThan"), &["5"]),
            add("cellIs", Some("between"), &["1"]),
        ],
        CalculationOptions::default(),
    );
    assert!(result.is_err());
    assert!(wb.model().sheets[0].conditional_formats.is_empty());
}

#[test]
fn collaborative_mode_refuses_conditional_format_ops() {
    let mut wb = Workbook::open_collaborative(&bytes(None), 42).unwrap();
    let result = wb.apply_ops(
        vec![add("cellIs", Some("greaterThan"), &["5"])],
        CalculationOptions::default(),
    );
    assert!(
        matches!(result, Err(Error::CollaborativeStructureOperation)),
        "{result:?}"
    );
}

fn source_cf_bytes() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&bytes(Some(
        r#"<cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs>"#,
    )))
    .unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = br#"<worksheet><sheetData><row r="1"><c r="A1" s="0"><v>10</v></c></row></sheetData><conditionalFormatting sqref="A1:A5" pivot="1" xmlns:a="urn:first"><cfRule type="expression" priority="1" dxfId="0"><formula>A1&gt;0</formula><a:extra/></cfRule></conditionalFormatting><conditionalFormatting sqref="A1:A5" pivot="0" xmlns:a="urn:second"><cfRule type="expression" priority="2" dxfId="0"><formula>A1&gt;5</formula><a:extra/></cfRule></conditionalFormatting></worksheet>"#.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn cell_style_edit_preserves_source_dxfs() {
    let mut wb = Workbook::open(&source_cf_bytes()).unwrap();
    let before = wb.model().sheets[0].conditional_formats.clone();
    wb.apply_ops(
        vec![Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1").unwrap(),
            patch: xlsx_ops::StylePatch {
                bold: Some(true),
                ..Default::default()
            },
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let reopened = Workbook::open(&wb.save().unwrap()).unwrap();
    assert_eq!(reopened.model().styles.dxfs.len(), 1);
    assert_eq!(reopened.model().sheets[0].conditional_formats, before);
}

#[test]
fn source_cf_wrappers_keep_attributes_and_namespaces() {
    let mut wb = Workbook::open(&source_cf_bytes()).unwrap();
    wb.apply_ops(
        vec![add("expression", None, &["A1>10"])],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(xml.contains(r#"pivot="1" xmlns:a="urn:first""#), "{xml}");
    assert!(xml.contains(r#"pivot="0" xmlns:a="urn:second""#), "{xml}");
    assert_eq!(
        Workbook::open(&saved).unwrap().model().sheets[0]
            .conditional_formats
            .iter()
            .map(|group| group.rules.len())
            .sum::<usize>(),
        3
    );
    wb.apply_ops(
        vec![
            Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id: "cf1".into(),
            },
            Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id: "cf3".into(),
            },
        ],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(!xml.contains("urn:first"), "{xml}");
    assert!(xml.contains(r#"pivot="0" xmlns:a="urn:second""#), "{xml}");
}

#[test]
fn binary_snapshots_refuse_source_cf_and_dxfs() {
    for source in [
        source_cf_bytes(),
        bytes(Some(r#"<dxfs count="1"><dxf/></dxfs>"#)),
    ] {
        let wb = Workbook::open(&source).unwrap();
        assert!(
            betteroffice_xlsx::WorkbookSnapshotEncoder::new(
                &wb,
                None,
                betteroffice_xlsx::SnapshotBudget::new(7, 16384).unwrap()
            )
            .is_err()
        );
        assert!(wb.save().is_ok());
    }
}
