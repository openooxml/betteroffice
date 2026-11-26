use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, Error, SheetId, Workbook};
use xlsx_model::CellProvider;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn package() -> Vec<u8> {
    let sheet = format!(
        concat!(
            r#"<row r="1"><c r="A1"><mc:AlternateContent xmlns:mc="{mc}"><mc:Choice Requires="x14"><f>1+1</f><v>2</v></mc:Choice>"#,
            r#"<mc:Fallback><f>1+1</f><v>2</v></mc:Fallback></mc:AlternateContent></c>"#,
            r#"<c r="B1"><f>SUM(2,3)</f><v>7</v></c></row>"#
        ),
        mc = MC
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData>{sheet}</sheetData></worksheet>"#),
        ),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn a_cell_alternate_content_formula_keeps_cached_results_on_open() {
    let workbook =
        Workbook::open_recalculating_stale(&package(), CalculationOptions::default(), None)
            .unwrap();
    let value = |address: &str| {
        workbook
            .model()
            .value(SheetId(0), CellRef::parse_a1(address).unwrap())
    };
    assert_eq!(value("A1"), CellValue::Number { value: 2.0 });
    assert_eq!(value("B1"), CellValue::Number { value: 7.0 });
}

#[test]
fn a_cell_alternate_content_formula_refuses_sheet_rewrites_before_mutation() {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open_recalculating_stale(&package(), options, None).unwrap();
    let model = workbook.model().clone();
    let version = workbook.version();
    let history = workbook.history_state();
    let baseline = workbook.save().unwrap();
    for address in ["A1", "B1"] {
        let at = CellRef::parse_a1(address).unwrap();
        let error = workbook
            .edit_cell(SheetId(0), at, "4", options)
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Spreadsheet(xlsx_parse::ParseError::UnsupportedEdit(message))
                if message.contains("Sheet1") && message.contains("mc:AlternateContent")
        ));
        assert_eq!(workbook.model(), &model);
        assert_eq!(workbook.version(), version);
        assert_eq!(workbook.history_state(), history);
        assert_eq!(workbook.save().unwrap(), baseline);
    }
}
