use xlsx_model::{CellRange, DifferentialFormat, SheetId};
use xlsx_parse::{
    SaveEdits, parse_workbook, parse_workbook_with_package,
    serialize_workbook_with_package_and_origins_after_edits,
};

fn parts(sheet_body: &str, styles: Option<&str>) -> Vec<(String, Vec<u8>)> {
    let workbook = r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
    let mut rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>"#.to_string();
    let mut out = vec![
        ("xl/workbook.xml".to_string(), workbook.as_bytes().to_vec()),
        ("xl/worksheets/sheet1.xml".to_string(), format!("<worksheet>{sheet_body}</worksheet>").into_bytes()),
        ("xl/worksheets/sheet2.xml".to_string(), b"<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>7</v></c></row></sheetData></worksheet>".to_vec()),
        ("xl/calcChain.xml".to_string(), b"<calcChain><c r=\"A1\" i=\"1\"/></calcChain>".to_vec()),
    ];
    if let Some(styles) = styles {
        rels.push_str(r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#);
        out.push((
            "xl/styles.xml".to_string(),
            format!("<styleSheet>{styles}</styleSheet>").into_bytes(),
        ));
    }
    rels.push_str("</Relationships>");
    out.push(("xl/_rels/workbook.xml.rels".to_string(), rels.into_bytes()));
    out
}

fn part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> &'a [u8] {
    &parts.iter().find(|(p, _)| p == path).unwrap().1
}

fn save(parsed: &xlsx_parse::ParsedWorkbook, wb: &xlsx_model::Workbook) -> Vec<(String, Vec<u8>)> {
    serialize_workbook_with_package_and_origins_after_edits(
        wb,
        &parsed.package,
        &[Some(0), Some(1)],
        &[Default::default(), Default::default()],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap()
    .into_iter()
    .map(|(p, b)| (p, b.into_owned()))
    .collect()
}

fn authored(wb: &mut xlsx_model::Workbook) {
    wb.sheets[0]
        .conditional_formats
        .push(xlsx_model::ConditionalFormatting {
            sqref: "A1:A5".into(),
            rules: vec![xlsx_model::ConditionalFormatRule {
                id: "cf1".into(),
                rule_type: "cellIs".into(),
                operator: Some("greaterThan".into()),
                formulas: vec!["5<6&\"x\"".into()],
                priority: 9,
                dxf_id: None,
                format: DifferentialFormat {
                    font_color: Some("#FF0000".into()),
                    bold: Some(true),
                    fill: Some("#00FF00".into()),
                },
                raw_xml: None,
            }],
        });
}

#[test]
fn authored_rule_writes_cfrule_and_appended_dxf() {
    let source = parts(
        "<sheetData/>",
        Some(
            r#"<fonts count="1"><font/></fonts><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs><tableStyles count="0"/>"#,
        ),
    );
    let parsed = parse_workbook_with_package(&source).unwrap();
    let mut wb = parsed.workbook.clone();
    authored(&mut wb);
    let saved = save(&parsed, &wb);
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml").to_vec()).unwrap();
    assert!(
        sheet.contains(r#"<conditionalFormatting sqref="A1:A5">"#),
        "{sheet}"
    );
    assert!(sheet.contains(r#"type="cellIs""#), "{sheet}");
    assert!(sheet.contains(r#"operator="greaterThan""#), "{sheet}");
    assert!(sheet.contains(r#"priority="9""#), "{sheet}");
    assert!(sheet.contains(r#"dxfId="1""#), "{sheet}");
    assert!(
        sheet.contains("5&lt;6&amp;\"x\"") || sheet.contains("5&lt;6&amp;"),
        "{sheet}"
    );
    let styles = String::from_utf8(part(&saved, "xl/styles.xml").to_vec()).unwrap();
    assert!(styles.contains(r#"<dxfs count="2">"#), "{styles}");
    assert!(
        styles.contains("<dxf><font><b/></font></dxf>"),
        "existing dxf preserved: {styles}"
    );
    assert!(styles.contains("FFFF0000"), "{styles}");
    assert!(styles.contains("FF00FF00"), "{styles}");
    assert!(
        styles.find("<dxfs").unwrap() < styles.find("<tableStyles").unwrap(),
        "{styles}"
    );
    assert_eq!(
        part(&saved, "xl/workbook.xml"),
        part(&source, "xl/workbook.xml")
    );
    assert_eq!(
        part(&saved, "xl/calcChain.xml"),
        part(&source, "xl/calcChain.xml")
    );
    assert_eq!(
        part(&saved, "xl/worksheets/sheet2.xml"),
        part(&source, "xl/worksheets/sheet2.xml")
    );
    let reopened = parse_workbook(&saved).unwrap();
    let rule = &reopened.sheets[0].conditional_formats[0].rules[0];
    assert_eq!(rule.rule_type, "cellIs");
    assert_eq!(rule.dxf_id, Some(1));
    assert_eq!(rule.format.font_color.as_deref(), Some("#ff0000"));
    assert_eq!(rule.format.fill.as_deref(), Some("#00ff00"));
    let _ = SheetId(0);
    let _ = CellRange::parse_a1("A1").unwrap();
}

#[test]
fn existing_raw_cf_and_dxf_survive_an_authored_add() {
    let raw = r#"<conditionalFormatting sqref="B2:B4"><cfRule type="top10" rank="10" priority="3" dxfId="0"><x/></cfRule></conditionalFormatting>"#;
    let source = parts(
        &format!(
            "<sheetData/><mergeCells count=\"1\"><mergeCell ref=\"A9:B9\"/></mergeCells>{raw}<dataValidations/>"
        ),
        Some(r#"<dxfs count="1"><dxf><font><color rgb="FF0000FF"/></font></dxf></dxfs>"#),
    );
    let parsed = parse_workbook_with_package(&source).unwrap();
    assert_eq!(
        parsed.workbook.sheets[0].conditional_formats[0].rules[0].rule_type,
        "top10"
    );
    let mut wb = parsed.workbook.clone();
    authored(&mut wb);
    let saved = save(&parsed, &wb);
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml").to_vec()).unwrap();
    assert!(sheet.contains(raw), "raw CF preserved: {sheet}");
    assert!(sheet.find("mergeCells").unwrap() < sheet.find("conditionalFormatting").unwrap());
    assert!(
        sheet.rfind("conditionalFormatting").unwrap() < sheet.find("dataValidations").unwrap(),
        "{sheet}"
    );
}

#[test]
fn missing_styles_part_is_created_for_authored_rule() {
    let source = parts("<sheetData/>", None);
    let parsed = parse_workbook_with_package(&source).unwrap();
    let mut wb = parsed.workbook.clone();
    authored(&mut wb);
    let saved = save(&parsed, &wb);
    let styles = String::from_utf8(part(&saved, "xl/styles.xml").to_vec()).unwrap();
    assert!(styles.contains(r#"<dxfs count="1">"#), "{styles}");
    let sheet = String::from_utf8(part(&saved, "xl/worksheets/sheet1.xml").to_vec()).unwrap();
    assert!(sheet.contains(r#"dxfId="0""#), "{sheet}");
}

#[test]
fn no_cf_delta_keeps_every_part_byte_identical() {
    let source = parts(
        "<sheetData/>",
        Some(r#"<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>"#),
    );
    let parsed = parse_workbook_with_package(&source).unwrap();
    let saved = serialize_workbook_with_package_and_origins_after_edits(
        &parsed.workbook,
        &parsed.package,
        &[Some(0), Some(1)],
        &[Default::default(), Default::default()],
        SaveEdits {
            changed: false,
            moved_references: false,
        },
    )
    .unwrap();
    for (path, bytes) in &source {
        assert_eq!(
            &saved.iter().find(|(p, _)| p == path).unwrap().1[..],
            &bytes[..],
            "{path}"
        );
    }
}

#[test]
fn dxf_only_styles_survive_save_without_package() {
    let source = parts(
        r#"<sheetData/><conditionalFormatting sqref="A1:A5"><cfRule type="expression" priority="1" dxfId="0"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>"#,
        Some(r#"<dxfs count="1"><dxf><font><color rgb="FFFF0000"/><b/></font></dxf></dxfs>"#),
    );
    let workbook = parse_workbook(&source).unwrap();
    assert_eq!(workbook.styles.dxfs.len(), 1);
    let saved = xlsx_parse::serialize_workbook(&workbook).unwrap();
    let styles = std::str::from_utf8(part(&saved, "xl/styles.xml")).unwrap();
    assert!(styles.contains(r#"<dxfs count="1">"#), "{styles}");
    assert!(styles.contains("FFFF0000"), "{styles}");
    let reopened = parse_workbook(&saved).unwrap();
    let rule = &reopened.sheets[0].conditional_formats[0].rules[0];
    assert_eq!(rule.dxf_id, Some(0));
    assert_eq!(rule.format.font_color.as_deref(), Some("#ff0000"));
    assert_eq!(rule.format.bold, Some(true));
}

#[test]
fn differential_colors_resolve_theme_tint_and_indexed_palettes() {
    let mut source = parts(
        r#"<sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="0"><formula>A1&gt;0</formula></cfRule><cfRule type="expression" priority="2" dxfId="1"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>"#,
        Some(
            r#"<dxfs count="2"><dxf><font><color theme="4" tint="0.5"/></font><fill><patternFill patternType="solid"><fgColor indexed="2"/></patternFill></fill></dxf><dxf><font><color indexed="3"/></font><fill><patternFill patternType="solid"><bgColor theme="4"/></patternFill></fill></dxf></dxfs><colors><indexedColors><rgbColor rgb="FF000000"/><rgbColor rgb="FFFFFFFF"/><rgbColor rgb="FF123456"/></indexedColors></colors>"#,
        ),
    );
    source.push(("xl/theme/theme1.xml".into(), br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Custom"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2><a:accent1><a:srgbClr val="204060"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#.to_vec()));
    let rels = source
        .iter_mut()
        .find(|(path, _)| path == "xl/_rels/workbook.xml.rels")
        .unwrap();
    rels.1 = String::from_utf8(rels.1.clone()).unwrap().replace("</Relationships>", r#"<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#).into_bytes();
    let workbook = parse_workbook(&source).unwrap();
    let rules = &workbook.sheets[0].conditional_formats[0].rules;
    let tinted = workbook
        .styles
        .resolve_color(&xlsx_model::Color::Theme { idx: 4, tint: 0.5 })
        .unwrap();
    assert_eq!(rules[0].format.font_color.as_deref(), Some(tinted.as_str()));
    assert_eq!(rules[0].format.fill.as_deref(), Some("#123456"));
    assert_eq!(rules[1].format.font_color.as_deref(), Some("#00ff00"));
    assert_eq!(rules[1].format.fill.as_deref(), Some("#204060"));
    let saved = save(&parse_workbook_with_package(&source).unwrap(), &workbook);
    assert_eq!(
        part(&saved, "xl/styles.xml"),
        part(&source, "xl/styles.xml")
    );
}
