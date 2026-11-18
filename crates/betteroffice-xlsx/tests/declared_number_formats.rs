//! A workbook's own `<numFmt>` entry for an id the built-in table also numbers
//! is what the structured export, CELL and cell entry read, the same as that
//! code declared under a custom id.

use betteroffice_xlsx::{
    CalculationOptions, CellRef, CellValue, SheetId, Workbook, XlsxExportOptions,
    export_xlsx_structured,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const ACCOUNTING_XML: &str = r#"_-* #,##0_-;\-* #,##0_-;_-* &quot;-&quot;_-;_-@_-"#;
const ACCOUNTING: &str = r#"_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-"#;
const EAST_ASIAN_DATE: &str = "yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;";
const ISO_DATE: &str = "yyyy-mm-dd";

const CASES: [(u16, &str, &[f64]); 3] = [
    (41, ACCOUNTING_XML, &[-1234.0, 1234.0]),
    (31, EAST_ASIAN_DATE, &[45000.0]),
    (14, ISO_DATE, &[45000.0]),
];

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#
        ));
    }
    xml + "</Relationships>"
}

/// Data!A1 down holds `values`, each styled with an xf whose `numFmtId` is
/// `id`, which the styles part declares with `code` as the attribute spells it.
fn book(id: u16, code: &str, values: &[f64]) -> Vec<u8> {
    let rows = values
        .iter()
        .enumerate()
        .map(|(index, value)| {
            let row = index + 1;
            format!(r#"<row r="{row}"><c r="A{row}" s="1"><v>{value}</v></c></row>"#)
        })
        .collect::<String>();
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "styles", "styles.xml"),
            ]),
        ),
        (
            "xl/styles.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}"><numFmts count="1"><numFmt numFmtId="{id}" formatCode="{code}"/></numFmts>"#,
                    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
                    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="{id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#
                ),
                MAIN = MAIN,
                id = id,
                code = code
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData>{rows}</sheetData></worksheet>"#
            ),
        ),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Each exported cell's display text and number format, in export order.
fn export(bytes: &[u8]) -> Vec<(String, String)> {
    let content = export_xlsx_structured(bytes, &XlsxExportOptions::default()).unwrap();
    content.sheets[0]
        .cells
        .iter()
        .map(|cell| (cell.display_text.clone(), cell.number_format.clone()))
        .collect()
}

/// What cell `a1` holds once `input` is typed into it.
fn entered(bytes: &[u8], a1: &str, input: &str) -> Option<CellValue> {
    let mut workbook = Workbook::open(bytes).unwrap();
    let at = CellRef::parse_a1(a1).unwrap();
    workbook
        .edit_cell(SheetId(0), at, input, CalculationOptions::default())
        .unwrap();
    workbook.model().sheets[0]
        .cell(at)
        .map(|cell| cell.value.clone())
}

#[test]
fn the_structured_export_reads_a_declared_builtin_id_like_its_custom_twin() {
    for (id, code, values) in CASES {
        assert_eq!(
            export(&book(id, code, values)),
            export(&book(164, code, values)),
            "numFmtId {id}"
        );
    }
    let declared = export(&book(41, ACCOUNTING_XML, &[-1234.0, 1234.0]));
    assert_eq!(declared[0].1, ACCOUNTING);
    assert!(
        declared[0].0.contains('-') && !declared[0].0.contains('('),
        "{declared:?}"
    );
}

#[test]
fn cell_format_reads_a_declared_builtin_id_like_its_custom_twin() {
    let formula = r#"=CELL("format",A1)"#;
    for (id, code, values) in CASES {
        assert_eq!(
            entered(&book(id, code, values), "B1", formula),
            entered(&book(164, code, values), "B1", formula),
            "numFmtId {id}"
        );
    }
    let shown = entered(&book(31, EAST_ASIAN_DATE, &[45000.0]), "B1", formula);
    assert_ne!(shown, Some(text("G")));
}

#[test]
fn entry_into_a_declared_builtin_text_format_keeps_the_text() {
    let declared = entered(&book(2, "@", &[1.0]), "A1", "0012");
    assert_eq!(declared, entered(&book(164, "@", &[1.0]), "A1", "0012"));
    assert_eq!(declared, Some(text("0012")));
}
