use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, SheetId, Workbook};
use xlsx_model::CellProvider;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn package(a1: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#)),
        ("xl/workbook.xml", format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#)),
        ("xl/_rels/workbook.xml.rels", format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#)),
        ("xl/worksheets/sheet1.xml", format!(r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1">{a1}<c r="B1"><v>1</v></c></row></sheetData></worksheet>"#)),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn a_plain_empty_boolean_reads_false() {
    let workbook = Workbook::open(&package(r#"<c r="A1" t="b"><v></v></c>"#)).unwrap();
    assert_eq!(
        workbook.model().value(SheetId(0), CellRef::new(0, 0)),
        CellValue::Bool { value: false }
    );
}

#[test]
fn a_formula_boolean_with_an_empty_value_is_recalculated_on_the_recalculating_open() {
    let workbook = Workbook::open_recalculating_stale(
        &package(r#"<c r="A1" t="b"><f>TRUE()</f><v/></c>"#),
        CalculationOptions::default(),
        None,
    )
    .unwrap();
    assert_eq!(
        workbook.model().value(SheetId(0), CellRef::new(0, 0)),
        CellValue::Bool { value: true }
    );
}
