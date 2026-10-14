use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, DateSystem, Workbook};

fn pkg(date1904: bool) -> Vec<u8> {
    let wbpr = if date1904 {
        "<workbookPr date1904=\"1\"/>"
    } else {
        ""
    };
    let ct = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
    let wbxml = format!(
        r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{wbpr}<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    );
    let wbrels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let sheet = r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>0</v></c><c r="B1"><f>TEXT(A1,"m/d/yy")</f></c></row></sheetData></worksheet>"#;
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".into(), ct.as_bytes().to_vec()),
        ("_rels/.rels".into(), rels.as_bytes().to_vec()),
        ("xl/workbook.xml".into(), wbxml.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".into(),
            wbrels.as_bytes().to_vec(),
        ),
        ("xl/worksheets/sheet1.xml".into(), sheet.as_bytes().to_vec()),
    ])
    .expect("package")
}

fn probe(date1904: bool) -> (DateSystem, CellValue) {
    let wb =
        Workbook::open_recalculated(&pkg(date1904), CalculationOptions::default()).expect("open");
    let m = wb.model();
    let sh = &m.sheets[0];
    (
        m.date_system,
        sh.cell(CellRef::parse_a1("B1").unwrap())
            .unwrap()
            .value
            .clone(),
    )
}

#[test]
fn text_uses_the_workbook_date_system() {
    assert_eq!(
        probe(false),
        (
            DateSystem::V1900,
            CellValue::Text {
                value: "1/0/00".into()
            }
        )
    );
    assert_eq!(
        probe(true),
        (
            DateSystem::V1904,
            CellValue::Text {
                value: "1/1/04".into()
            }
        )
    );
}
