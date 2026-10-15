use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, Sheet, SheetId, Workbook, WorkbookModel,
};
use xlsx_model::{CellProvider, DefinedName};

const REL_EXTERNAL_LINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";

/// Formula cells the engine cannot compute (an external workbook, R1C1
/// INDIRECT, date text read through the author's locale) beside ones it can.
const CACHED: [(&str, &str, f64); 5] = [
    ("B1", "'[1]Other Book'!B2", 91.5),
    ("B2", "INDIRECT(\"R1C1\",FALSE)", 10.0),
    ("B3", "IF(Pay<>\"\",DATE(2020,Pay/10,1),\"\")", 43_891.0),
    ("C2", "C1+1", 44_520.0),
    ("D20", "SUM(A1:A5)", 150.0),
];

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn replace_part(parts: &mut [(String, Vec<u8>)], path: &str, from: &str, to: &str) {
    let part = &mut parts.iter_mut().find(|(name, _)| name == path).unwrap().1;
    let text = std::str::from_utf8(part).unwrap();
    assert!(text.contains(from), "{path}");
    *part = text.replacen(from, to, 1).into_bytes();
}

fn source() -> Vec<u8> {
    let mut sheet = Sheet::new("Sheet1");
    for row in 1..=5 {
        sheet.set_cell(
            cell(&format!("A{row}")),
            Cell {
                value: number(row as f64 * 10.0),
                ..Cell::default()
            },
        );
    }
    sheet.set_cell(
        cell("C1"),
        Cell {
            value: CellValue::Text {
                value: "19/11/2021".into(),
            },
            ..Cell::default()
        },
    );
    for (address, formula, cached) in CACHED {
        sheet.set_cell(
            cell(address),
            Cell {
                value: number(cached),
                formula: Some(formula.into()),
                style: None,
            },
        );
    }
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    model.defined_names.push(DefinedName {
        name: "Pay".into(),
        formula: "Sheet1!$A$1:$A$5".into(),
        local_sheet: None,
        hidden: false,
    });
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    replace_part(
        &mut parts,
        "xl/workbook.xml",
        "</sheets>",
        r#"</sheets><externalReferences><externalReference r:id="rIdExternal"/></externalReferences>"#,
    );
    replace_part(
        &mut parts,
        "xl/_rels/workbook.xml.rels",
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rIdExternal" Type="{REL_EXTERNAL_LINK}" Target="externalLinks/externalLink1.xml"/></Relationships>"#
        ),
    );
    replace_part(
        &mut parts,
        "[Content_Types].xml",
        "</Types>",
        r#"<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>"#,
    );
    parts.extend([
        (
            "xl/externalLinks/externalLink1.xml".to_owned(),
            br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Other Book"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="2"><cell r="B2"><v>91.5</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#.to_vec(),
        ),
        (
            "xl/externalLinks/_rels/externalLink1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="other.xlsx" TargetMode="External"/></Relationships>"#.to_vec(),
        ),
    ]);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn sheet_xml(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let part = &parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap()
        .1;
    String::from_utf8(part.clone()).unwrap()
}

fn cell_xml<'a>(sheet: &'a str, address: &str) -> &'a str {
    let start = sheet.find(&format!(r#"<c r="{address}""#)).unwrap();
    let end = start + sheet[start..].find("</c>").unwrap() + "</c>".len();
    &sheet[start..end]
}

fn cached_results(bytes: &[u8]) -> Vec<CellValue> {
    let model = xlsx_parse::parse_workbook(&ooxml_opc::unzip_parts(bytes).unwrap()).unwrap();
    CACHED
        .iter()
        .map(|(address, _, _)| model.value(SheetId(0), cell(address)))
        .collect()
}

#[test]
fn an_edit_elsewhere_saves_the_source_results_of_formulas_it_does_not_reach() {
    let bytes = source();
    let expected = CACHED
        .iter()
        .map(|(_, _, cached)| number(*cached))
        .collect::<Vec<_>>();
    assert_eq!(cached_results(&bytes), expected);

    let mut workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    for (address, _, cached) in CACHED {
        assert_eq!(
            workbook.model().value(SheetId(0), cell(address)),
            number(cached),
            "{address} after the recalculation at open"
        );
    }
    assert!(
        workbook
            .edit_cell(
                SheetId(0),
                cell("Z100"),
                "edited",
                CalculationOptions::default()
            )
            .unwrap()
            .applied
    );
    let saved = workbook.save().unwrap();
    assert_eq!(cached_results(&saved), expected);
    let (source_sheet, saved_sheet) = (sheet_xml(&bytes), sheet_xml(&saved));
    for (address, _, _) in CACHED {
        assert!(
            saved_sheet.contains(cell_xml(&source_sheet, address)),
            "{address} was rewritten"
        );
    }
}
