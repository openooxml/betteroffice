use xlsx_model::CellRange;
use xlsx_parse::{
    SaveEdits, parse_workbook_with_package, serialize_workbook_with_package_and_origins_after_edits,
};

fn parts(sheet: &str) -> Vec<(String, Vec<u8>)> {
    vec![
        ("xl/workbook.xml".into(), br#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data 7" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
        ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), sheet.as_bytes().to_vec()),
        ("customXml/item1.xml".into(), b"<extra>keep me</extra>".to_vec()),
    ]
}

#[test]
fn only_the_ref_value_changes_with_source_attributes_children_and_prefixes() {
    for prefix in ["", "s:"] {
        let namespace = if prefix.is_empty() {
            "xmlns"
        } else {
            "xmlns:s"
        };
        let sheet = format!(
            r#"<{prefix}worksheet {namespace}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><{prefix}sheetData/><{prefix}autoFilter custom="keep" ref = 'A3:A30'><{prefix}filterColumn colId = '0'><{prefix}filters><{prefix}filter val="x" /></{prefix}filters></{prefix}filterColumn><{prefix}extLst><extra xmlns="urn:extra"> keep </extra></{prefix}extLst></{prefix}autoFilter><{prefix}mergeCells count = '1'><{prefix}mergeCell ref = 'H1:I1' /></{prefix}mergeCells></{prefix}worksheet>"#
        );
        let source = parts(&sheet);
        let parsed = parse_workbook_with_package(&source).unwrap();
        let mut model = parsed.workbook.clone();
        model.sheets[0].auto_filter.as_mut().unwrap().range =
            CellRange::parse_a1("A3:F30").unwrap();
        let saved = serialize_workbook_with_package_and_origins_after_edits(
            &model,
            &parsed.package,
            &[Some(0)],
            &[Default::default()],
            SaveEdits {
                changed: true,
                moved_references: false,
            },
        )
        .unwrap();
        let xml = &saved
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1;
        assert_eq!(xml.as_ref(), sheet.replace("A3:A30", "A3:F30").as_bytes());
        assert_eq!(
            saved
                .iter()
                .find(|(path, _)| path == "customXml/item1.xml")
                .unwrap()
                .1
                .as_ref(),
            b"<extra>keep me</extra>"
        );
    }
}

#[test]
fn untouched_elements_and_duplicate_filters_round_trip() {
    let sheet = r#"<worksheet><sheetData/><autoFilter ref='A3:A30'><filterColumn colId="0"/></autoFilter><autoFilter ref="B1:B5"/></worksheet>"#;
    let source = parts(sheet);
    let parsed = parse_workbook_with_package(&source).unwrap();
    assert_eq!(
        parsed.workbook.sheets[0]
            .auto_filter
            .as_ref()
            .unwrap()
            .source_count,
        2
    );
    let saved = serialize_workbook_with_package_and_origins_after_edits(
        &parsed.workbook,
        &parsed.package,
        &[Some(0)],
        &[Default::default()],
        SaveEdits {
            changed: false,
            moved_references: false,
        },
    )
    .unwrap();
    assert_eq!(
        saved
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1
            .as_ref(),
        sheet.as_bytes()
    );
}

#[test]
fn filters_without_a_usable_ref_open_and_round_trip() {
    for filter in [
        r#"<autoFilter/>"#,
        r#"<autoFilter><filterColumn colId="0"/></autoFilter>"#,
        r#"<autoFilter ref="not a range"/>"#,
    ] {
        let sheet = format!(r#"<worksheet><sheetData/>{filter}</worksheet>"#);
        let source = parts(&sheet);
        let parsed = parse_workbook_with_package(&source).unwrap();
        assert!(parsed.workbook.sheets[0].auto_filter.is_none());
        for changed in [false, true] {
            let saved = serialize_workbook_with_package_and_origins_after_edits(
                &parsed.workbook,
                &parsed.package,
                &[Some(0)],
                &[Default::default()],
                SaveEdits {
                    changed,
                    moved_references: false,
                },
            )
            .unwrap();
            let xml = &saved
                .iter()
                .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
                .unwrap()
                .1;
            assert!(
                String::from_utf8_lossy(xml).contains(filter),
                "{filter} lost when changed={changed}"
            );
        }
    }
}
