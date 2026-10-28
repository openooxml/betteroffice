use std::collections::BTreeMap;

use betteroffice_xlsx::Workbook;
use serde_json::json;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// a sheet with what a lossy writer drops: an x14 conditional format and an
/// x14 data validation in extLst, an mc:AlternateContent block, a legacy
/// data validation, a custom attribute on a cell and on a row, and a
/// prefixed element the model does not know.
const RICH: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
    r#"xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" "#,
    r#"xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" "#,
    r#"mc:Ignorable="x14ac xr" xr:uid="{00000000-0001-0000-0000-000000000000}">"#,
    r#"<sheetPr><tabColor rgb="FF00B050"/></sheetPr><dimension ref="A1:C3"/>"#,
    r#"<sheetViews><sheetView workbookViewId="0"/></sheetViews>"#,
    r#"<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>"#,
    r#"<sheetData>"#,
    r#"<row r="1" spans="1:3" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><f>A1+B1</f><v>3</v></c></row>"#,
    r#"<row r="2" spans="1:3" x14ac:dyDescent="0.25"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c></row>"#,
    r#"<row r="3" spans="1:3" x14ac:dyDescent="0.25"><c r="A3"><v>7</v></c></row>"#,
    r#"</sheetData>"#,
    r#"<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showErrorMessage="1" sqref="A1:A3"><formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations>"#,
    r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
    r#"<mc:AlternateContent><mc:Choice Requires="x14"><controls/></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
    r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">"#,
    r#"<x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#,
    r#"<x14:cfRule type="dataBar" id="{00000000-000E-0000-0000-000001000000}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>"#,
    r#"<xm:sqref>B1:B3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>"#,
    r#"<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">"#,
    r#"<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>A1:A3</xm:f></x14:formula1><xm:sqref>C1:C3</xm:sqref></x14:dataValidation></x14:dataValidations></ext>"#,
    r#"</extLst></worksheet>"#,
);

fn package() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                r#"<Default Extension="xml" ContentType="application/xml"/>"#,
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
                r#"</Types>"#,
            )
            .to_owned(),
        ),
        (
            "_rels/.rels",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        (
            "xl/workbook.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                    r#"<workbook xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                    r#"<sheets><sheet name="Edited" sheetId="1" r:id="rId1"/>"#,
                    r#"<sheet name="Untouched" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
                ),
                main = MAIN,
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#,
                r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>"#,
                r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        (
            "xl/styles.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{main}">"#,
                    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
                    r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
                    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
                    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
                    r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
                ),
                main = MAIN,
            ),
        ),
        ("xl/worksheets/sheet1.xml", RICH.to_owned()),
        ("xl/worksheets/sheet2.xml", RICH.to_owned()),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

/// one agent write: `setCellValues` through the batch API, the call the
/// agent tools make before they save.
fn agent_write(workbook: &mut Workbook, sheet: usize, a1: &str, value: f64) -> Result<(), String> {
    let request = json!({
        "expectVersion": workbook.version(),
        "source": "agent",
        "steps": [{
            "op": "setCellValues",
            "target": { "sheetId": format!("sheet:{sheet}"), "range": { "kind": "a1", "a1": a1 } },
            "values": [[{ "kind": "number", "value": value }]]
        }]
    });
    workbook
        .apply_edits_json(&request.to_string())
        .map(|_| ())
        .map_err(|error| error.to_string())
}

fn without(xml: &str, needle_start: &str, end: &str) -> String {
    let Some(at) = xml.find(needle_start) else {
        return xml.to_owned();
    };
    let tail = &xml[at..];
    let close = if let Some(self_close) = tail.find("/>")
        && tail.find('>') == Some(self_close + 1)
    {
        self_close + 2
    } else {
        tail.find(end).map(|i| i + end.len()).unwrap_or(tail.len())
    };
    format!("{}{}", &xml[..at], &tail[close..])
}

#[test]
fn an_agent_write_keeps_untouched_cells_sheets_and_parts() {
    let source = package();
    let mut workbook = Workbook::open(&source).unwrap();
    agent_write(&mut workbook, 0, "B2", 6.0).unwrap();
    let saved = parts(&workbook.save().unwrap());
    let before = parts(&source);
    for (name, bytes) in &before {
        if name == "xl/worksheets/sheet1.xml" || name == "xl/workbook.xml" {
            continue;
        }
        assert_eq!(saved.get(name), Some(bytes), "{name} changed");
    }
    let workbook_xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
    assert!(workbook_xml.contains(r#"<calcPr fullCalcOnLoad="1"/>"#));
    assert_eq!(
        workbook_xml.replace(r#"<calcPr fullCalcOnLoad="1"/>"#, ""),
        String::from_utf8(before["xl/workbook.xml"].clone()).unwrap(),
        "xl/workbook.xml changed beyond the full recalc request"
    );
    let edited = String::from_utf8(saved["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let written = &edited[edited.find(r#"<c r="B2""#).unwrap()..];
    assert!(
        written[..written.find("</c>").unwrap()].contains("<v>6</v>"),
        "the write did not reach the saved sheet: {edited}"
    );
    for kept in [
        r#"x14ac:dyDescent="0.25""#,
        r#"xr:uid="{00000000-0001-0000-0000-000000000000}""#,
        "<mc:AlternateContent>",
        "<x14:dataBar",
        "<x14:dataValidation ",
        r#"<dataValidation type="whole""#,
        r#"<tabColor rgb="FF00B050"/>"#,
    ] {
        assert!(edited.contains(kept), "edited sheet lost {kept}: {edited}");
    }
    let strip = |xml: &str| without(xml, r#"<c r="B2""#, "</c>");
    assert_eq!(
        strip(&edited),
        strip(RICH),
        "the edited sheet changed outside the written cell"
    );
}
