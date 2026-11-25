use betteroffice_xlsx::{CalculationOptions, CellRef, SheetId, Workbook};
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

fn saved_a1(saved: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let (_, sheet) = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    let sheet = String::from_utf8_lossy(sheet).to_string();
    sheet
        .split("<c ")
        .find(|cell| cell.contains(r#"r="A1""#))
        .map(|cell| format!("<c {}", cell.split("</c>").next().unwrap_or(cell)))
        .unwrap_or_default()
}

#[test]
fn empty_boolean_behaviour() {
    for (case, a1) in [
        ("value", r#"<c r="A1" t="b"><v></v></c>"#),
        ("formula", r#"<c r="A1" t="b"><f>TRUE()</f><v/></c>"#),
    ] {
        let source = package(a1);
        let at = CellRef::new(0, 0);
        let opened = Workbook::open(&source).unwrap();
        let stale =
            Workbook::open_recalculating_stale(&source, CalculationOptions::default(), None)
                .unwrap();
        let mut edited = Workbook::open(&source).unwrap();
        edited
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 1),
                "2",
                CalculationOptions::default(),
            )
            .unwrap();
        println!(
            "BEHAVIOUR {case} open={:?} default_open={:?} edited_save_a1={}",
            opened.model().value(SheetId(0), at),
            stale.model().value(SheetId(0), at),
            saved_a1(&edited.save().unwrap())
        );
    }
}
