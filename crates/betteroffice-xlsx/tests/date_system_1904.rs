//! A 1904 workbook must compute dates with the 1904 epoch.
//!
//! Before the fix, `xlsx-calc` hardcoded `DateSystem::V1900` (text.rs:488) and
//! `EvalContext` had no date_system field, so TEXT/YEAR/MONTH/DAY/DATE all
//! returned 1900-system answers for a 1904 workbook — 1462 days off.
//!
//! Lives here, not in xlsx-calc: that crate has no dev-dependencies and cannot
//! construct a workbook, which is why the gap survived.
use betteroffice_xlsx::{CalculationOptions, CellRef, Workbook};

fn pkg(date1904: bool) -> Vec<u8> {
    let wbpr = if date1904 { "<workbookPr date1904=\"1\"/>" } else { "" };
    let ct = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;
    let rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
    let wbxml = format!(r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{wbpr}<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#);
    let wbrels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let styles = r#"<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="m/d/yy"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;
    let sheet = r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1"><f>TEXT(A1,"m/d/yy")</f></c><c r="C1"><f>YEAR(A1)</f></c><c r="D1"><f>MONTH(A1)</f></c><c r="E1"><f>DAY(A1)</f></c><c r="F1"><f>DATE(1904,3,15)</f></c></row></sheetData></worksheet>"#;
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".into(), ct.as_bytes().to_vec()),
        ("_rels/.rels".into(), rels.as_bytes().to_vec()),
        ("xl/workbook.xml".into(), wbxml.as_bytes().to_vec()),
        ("xl/_rels/workbook.xml.rels".into(), wbrels.as_bytes().to_vec()),
        ("xl/styles.xml".into(), styles.as_bytes().to_vec()),
        ("xl/worksheets/sheet1.xml".into(), sheet.as_bytes().to_vec()),
    ])
    .expect("package")
}

/// (ds, TEXT, YEAR, MONTH, DAY, DATE(1904,3,15))
fn probe(date1904: bool) -> (String, String, String, String, String, String) {
    let wb = Workbook::open_recalculated(&pkg(date1904), CalculationOptions::default()).expect("open");
    let m = wb.model();
    let sh = &m.sheets[0];
    let g = |a: &str| sh.cell(CellRef::parse_a1(a).unwrap()).unwrap().value.clone();
    let d = |a: &str| format!("{:?}", g(a));
    (
        format!("{:?}", m.date_system),
        d("B1"),
        d("C1"),
        d("D1"),
        d("E1"),
        d("F1"),
    )
}

#[test]
fn date_functions_honour_the_workbook_date_system() {
    let (ds1900, t1900, y1900, mo1900, dy1900, dt1900) = probe(false);
    let (ds1904, t1904, y1904, mo1904, dy1904, dt1904) = probe(true);
    println!("1900  ds={ds1900}");
    println!("  TEXT={t1900}  YEAR={y1900}  MONTH={mo1900}  DAY={dy1900}  DATE={dt1900}");
    println!("1904  ds={ds1904}");
    println!("  TEXT={t1904}  YEAR={y1904}  MONTH={mo1904}  DAY={dy1904}  DATE={dt1904}");

    assert!(ds1900.contains("V1900"), "{ds1900}");
    assert!(ds1904.contains("V1904"), "{ds1904}");

    // 1900: serial 0 is excel's phantom 1900-01-00. unchanged by this fix.
    assert!(t1900.contains("1/0/00"), "1900 TEXT={t1900}");
    assert!(y1900.contains("1900.0"), "1900 YEAR={y1900}");

    // 1904: serial 0 is the real date 1904-01-01.
    assert!(t1904.contains("1/1/04"), "1904 TEXT={t1904}");
    assert!(y1904.contains("1904.0"), "1904 YEAR={y1904}");
    assert!(mo1904.contains("1.0"), "1904 MONTH={mo1904}");
    assert!(dy1904.contains("1.0"), "1904 DAY={dy1904}");
}
