use betteroffice_xlsx::{CalculationOptions, CellRef, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EXTERNAL_FORMULAS: [&str; 4] = [
    "SUM([1]Sheet1:'Sheet 2'!A1)",
    "SUM([Book.xlsx]Sheet1:'Sheet 2'!A1:B2)",
    "SUM([1]Sheet1:'Sheet 2'!1:2)",
    "SUM([1]Sheet1:'Sheet 2'!#REF!)",
];

fn workbook_bytes(formula: &str) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
    </Types>"#;
    let package_rels = format!(
        r#"<Relationships xmlns="{RELS}">
            <Relationship Id="rId1" Type="{DOC_RELS}/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}">
            <sheets>
                <sheet name="Sheet 2" sheetId="1" r:id="rId1"/>
                <sheet name="Report" sheetId="2" r:id="rId2"/>
            </sheets>
            <externalReferences><externalReference r:id="rId3"/></externalReferences>
            <definedNames><definedName name="SpanFormula">{formula}</definedName></definedNames>
        </workbook>"#
    );
    let workbook_rels = format!(
        r#"<Relationships xmlns="{RELS}">
            <Relationship Id="rId1" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet1.xml"/>
            <Relationship Id="rId2" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet2.xml"/>
            <Relationship Id="rId3" Type="{DOC_RELS}/externalLink" Target="externalLinks/externalLink1.xml"/>
        </Relationships>"#
    );
    let local_sheet = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let report = format!(
        r#"<worksheet xmlns="{MAIN}">
            <sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>{formula}</f><v>3</v></c></row></sheetData>
            <conditionalFormatting sqref="A1">
                <cfRule type="expression" priority="1"><formula>{formula}</formula></cfRule>
            </conditionalFormatting>
        </worksheet>"#
    );
    let external_link = format!(
        r#"<externalLink xmlns="{MAIN}" xmlns:r="{DOC_RELS}">
            <externalBook r:id="rId1">
                <sheetNames><sheetName val="Sheet1"/><sheetName val="Sheet 2"/></sheetNames>
            </externalBook>
        </externalLink>"#
    );
    let external_link_rels = format!(
        r#"<Relationships xmlns="{RELS}">
            <Relationship Id="rId1" Type="{DOC_RELS}/externalLinkPath" Target="Book.xlsx" TargetMode="External"/>
        </Relationships>"#
    );
    let parts = [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", package_rels.as_str()),
        ("xl/workbook.xml", workbook.as_str()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_str()),
        ("xl/worksheets/sheet1.xml", local_sheet.as_str()),
        ("xl/worksheets/sheet2.xml", report.as_str()),
        ("xl/externalLinks/externalLink1.xml", external_link.as_str()),
        (
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            external_link_rels.as_str(),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.as_bytes()));
    ooxml_opc::rezip_parts_borrowed(&parts).unwrap()
}

fn assert_formulas(workbook: &Workbook, expected: &str, stage: &str) {
    let model = workbook.model();
    let report = model.sheets.iter().find(|sheet| sheet.name == "Report").unwrap();
    let cell = report.cell(CellRef::new(0, 1)).unwrap();
    assert_eq!(cell.formula.as_deref(), Some(expected), "Report!B1 {stage}");

    let defined_name = &model.defined_names[0];
    assert_eq!(defined_name.name, "SpanFormula");
    assert_eq!(defined_name.formula, expected, "SpanFormula {stage}");

    let rule = &report.conditional_formats[0].rules[0];
    assert_eq!(rule.formulas, [expected], "conditional-format rule {stage}");
}

fn assert_sheet_edit(source: &str, op: Op, expected: &str) {
    let mut workbook = Workbook::open(&workbook_bytes(source)).unwrap();
    assert_formulas(&workbook, source, "before the edit");
    let result = workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .unwrap_or_else(|error| panic!("sheet edit failed for {source}: {error}"));
    assert!(result.applied, "sheet edit did not apply for {source}");
    assert_formulas(&workbook, expected, "after the edit");

    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_formulas(&reopened, expected, "after reopening");
}

fn rename_sheet() -> Op {
    Op::RenameSheet {
        sheet: SheetId(0),
        name: "New Sheet".into(),
    }
}

#[test]
fn external_mixed_quoted_spans_survive_sheet_rename() {
    for formula in EXTERNAL_FORMULAS {
        assert_sheet_edit(formula, rename_sheet(), formula);
    }
}

#[test]
fn external_mixed_quoted_spans_survive_sheet_removal() {
    for formula in EXTERNAL_FORMULAS {
        assert_sheet_edit(formula, Op::RemoveSheet { index: 0 }, formula);
    }
}

#[test]
fn local_mixed_quoted_span_follows_sheet_rename() {
    assert_sheet_edit(
        "SUM(Report:'Sheet 2'!A1)",
        rename_sheet(),
        "SUM('Report:New Sheet'!A1)",
    );
}
