#[cfg(feature = "raster")]
use betteroffice_xlsx::RenderOptions;
use betteroffice_xlsx::{
    AnchorCell, AnchorEditAs, CalculationOptions, Cell, CellInput, CellRange, CellRef, CellState,
    CellValue, ChartAnchor, ChartRef, ChartRefKind, DefinedName, DrawCmd, Error, FreezePane,
    GridGeometry, Hyperlink, MAX_COLLABORATION_BYTES, MAX_COLLABORATION_CLIENT_ID,
    MAX_COLLABORATION_STATE_VECTOR_ENTRIES, MAX_ROWS, NumberFormatKind, NumberFormatMutation, Op,
    ProposalEditInput, ProposalRequest, Sheet, SheetChart, SheetId, StylePatch, UpdateOrigin,
    Viewport, Workbook, WorkbookModel,
};
use std::collections::BTreeMap;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};
use yrs::Update as YrsUpdate;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn sample_parts() -> Vec<(String, Vec<u8>)> {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Number { value: 10.0 },
            style: Some(0),
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("A2"),
        Cell {
            value: CellValue::Number { value: 5.0 },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("B1"),
        Cell {
            value: CellValue::Number { value: 999.0 },
            formula: Some("SUM(A1:A2)".into()),
            ..Cell::default()
        },
    );
    let mut model = WorkbookModel::default();
    model.styles.cell_xfs.push(Default::default());
    model.sheets.push(sheet);
    model.sheets.push(Sheet::new("Empty"));
    xlsx_parse::serialize_workbook(&model).unwrap()
}

fn sample_xlsx() -> Vec<u8> {
    ooxml_opc::rezip_parts(&sample_parts()).unwrap()
}

fn preservation_fixture_parts() -> Vec<(String, Vec<u8>)> {
    let mut model = WorkbookModel::default();
    model.shared_strings.push("original".to_owned());
    model.styles.cell_xfs.push(Default::default());
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Text {
                value: "original".to_owned(),
            },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("B2"),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    model.sheets.push(sheet);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();

    set_test_part(
        &mut parts,
        "xl/workbook.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Data" sheetId="7" r:id="rId1"/></sheets><definedNames><definedName name="NamedCell">Data!$A$1</definedName></definedNames><calcPr calcId="191029"/></workbook>"#.to_vec(),
    );
    set_test_part(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><tabColor rgb="FF4472C4"/></sheetPr><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="B2"><v>1</v></c></row></sheetData><autoFilter ref="A1:B2"/><conditionalFormatting sqref="B2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" sqref="B2"><formula1>0</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="B2" r:id="rIdHyperlink"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape"/><drawing r:id="rIdDrawing"/><legacyDrawing r:id="rIdVml"/><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>"#.to_vec(),
    );
    set_test_part(
        &mut parts,
        "xl/sharedStrings.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><rPr><b/></rPr><t>orig</t></r><r><rPr><i/></rPr><t>inal</t></r><phoneticPr fontId="1"/></si><extLst><ext uri="{A68B0E0A-4E93-46C8-A4A4-57E4A6A3B123}"/></extLst></sst>"#.to_vec(),
    );
    parts.push((
        "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rIdHyperlink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid" TargetMode="External"/></Relationships>"#.to_vec(),
    ));
    parts.extend([
        (
            "xl/drawings/drawing1.xml".to_owned(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
        ),
        (
            "xl/tables/table1.xml".to_owned(),
            br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2"><autoFilter ref="A1:B2"/><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Value"/></tableColumns></table>"#.to_vec(),
        ),
        (
            "xl/comments1.xml".to_owned(),
            br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>BetterOffice</author></authors><commentList><comment ref="B2" authorId="0"><text><t>keep me</t></text></comment></commentList></comments>"#.to_vec(),
        ),
        (
            "xl/drawings/vmlDrawing1.vml".to_owned(),
            br#"<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="_x0000_s1025"/></xml>"#.to_vec(),
        ),
        (
            "xl/calcChain.xml".to_owned(),
            br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B2" i="1"/></calcChain>"#.to_vec(),
        ),
        (
            "xl/externalLinks/externalLink1.xml".to_owned(),
            br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook/></externalLink>"#.to_vec(),
        ),
        (
            "docProps/core.xml".to_owned(),
            br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:revision>9</cp:revision></cp:coreProperties>"#.to_vec(),
        ),
        (
            "customXml/item1.xml".to_owned(),
            br#"<custom fidelity="byte-identical">payload</custom>"#.to_vec(),
        ),
    ]);

    let workbook_rels = test_part_text(&parts, "xl/_rels/workbook.xml.rels")
        .replace(
            "</Relationships>",
            r#"<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/><Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>"#,
        );
    set_test_part(
        &mut parts,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.into_bytes(),
    );
    let root_rels = test_part_text(&parts, "_rels/.rels").replace(
        "</Relationships>",
        r#"<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#,
    );
    set_test_part(&mut parts, "_rels/.rels", root_rels.into_bytes());
    let content_types = test_part_text(&parts, "[Content_Types].xml")
        .replacen(
            "<Override",
            r#"<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override"#,
            1,
        )
        .replace(
            "</Types>",
            r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#,
        );
    set_test_part(
        &mut parts,
        "[Content_Types].xml",
        content_types.into_bytes(),
    );
    let styles = test_part_text(&parts, "xl/styles.xml").replace(
        "</styleSheet>",
        r#"<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2"/></styleSheet>"#,
    );
    set_test_part(&mut parts, "xl/styles.xml", styles.into_bytes());
    parts
}

fn preservation_fixture() -> Vec<u8> {
    ooxml_opc::rezip_parts(&preservation_fixture_parts()).unwrap()
}

fn non_worksheet_fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Chart"));
    model.sheets.push(Sheet::new("Dialog"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    rename_test_part(
        &mut parts,
        "xl/worksheets/sheet2.xml",
        "xl/chartsheets/sheet1.xml",
    );
    rename_test_part(
        &mut parts,
        "xl/worksheets/sheet3.xml",
        "xl/dialogsheets/sheet1.xml",
    );
    set_test_part(
        &mut parts,
        "xl/chartsheets/sheet1.xml",
        br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>"#.to_vec(),
    );
    set_test_part(
        &mut parts,
        "xl/dialogsheets/sheet1.xml",
        br#"<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></dialogsheet>"#.to_vec(),
    );
    set_test_part(
        &mut parts,
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet" Target="dialogsheets/sheet1.xml"/></Relationships>"#.to_vec(),
    );
    set_test_part(
        &mut parts,
        "[Content_Types].xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/><Override PartName="/xl/dialogsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"/></Types>"#.to_vec(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn strict_prefixed_fixture() -> Vec<u8> {
    let strict_main = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    let strict_rel = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.ms-excel.worksheet+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{strict_rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><s:workbook xmlns:s="{strict_main}" xmlns:rel="{strict_rel}"><s:sheets><s:sheet name="Data" sheetId="1" rel:id="rId1"/></s:sheets><s:definedNames><s:definedName name="StrictName">Data!$A$1</s:definedName></s:definedNames></s:workbook>"#
            )
            .into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{strict_rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><s:worksheet xmlns:s="{strict_main}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:fixture-extension" mc:Ignorable="x"><x:sheetData marker="keep"/><mc:AlternateContent><mc:Choice Requires="s"><s:sheetPr/></mc:Choice><mc:Fallback><s:sheetPr/></mc:Fallback></mc:AlternateContent><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
            )
            .into_bytes(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn defined_names_fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Middle"));
    model.sheets.push(Sheet::new("Tail"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    set_test_part(
        &mut parts,
        "xl/workbook.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Middle" sheetId="2" r:id="rId2"/><sheet name="Tail" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="GlobalData">Data!$A$1</definedName><definedName name="AmbiguousData">Data</definedName><definedName name="GlobalMiddle">Middle!$A$1</definedName><definedName name="LocalData" localSheetId="0">Data!$A$1</definedName><definedName name="LocalMiddle" localSheetId="1">Middle!$A$1</definedName><definedName name="LocalTail" localSheetId="2">Tail!$A$1</definedName><definedName name="Unrelated">42</definedName></definedNames></workbook>"#.to_vec(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Two `<si>` entries reading `Total`, one plain and one bold, with a cell on
/// each. Text alone cannot tell them apart, so only the recorded index keeps
/// each cell on its own run formatting.
fn ambiguous_shared_string_fixture() -> Vec<u8> {
    let mut model = WorkbookModel {
        shared_strings: vec!["Total".to_owned(), "Total".to_owned()],
        ..WorkbookModel::default()
    };
    let mut sheet = Sheet::new("Data");
    for address in ["B2", "D2"] {
        sheet.set_cell(
            cell(address),
            Cell {
                value: CellValue::Text {
                    value: "Total".to_owned(),
                },
                ..Cell::default()
            },
        );
    }
    model.sheets.push(sheet);
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    set_test_part(
        &mut parts,
        "xl/sharedStrings.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Total</t></si><si><r><rPr><b/></rPr><t>Total</t></r></si></sst>"#.to_vec(),
    );
    set_test_part(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="s"><v>0</v></c><c r="D2" t="s"><v>1</v></c></row></sheetData></worksheet>"#.to_vec(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_sheet_text(workbook: &Workbook) -> String {
    String::from_utf8(package_map(&workbook.save().unwrap())["xl/worksheets/sheet1.xml"].clone())
        .unwrap()
}

fn set_test_part(parts: &mut [(String, Vec<u8>)], path: &str, bytes: Vec<u8>) {
    parts.iter_mut().find(|(name, _)| name == path).unwrap().1 = bytes;
}

fn rename_test_part(parts: &mut [(String, Vec<u8>)], from: &str, to: &str) {
    parts.iter_mut().find(|(name, _)| name == from).unwrap().0 = to.to_owned();
}

fn test_part_text(parts: &[(String, Vec<u8>)], path: &str) -> String {
    String::from_utf8(
        parts
            .iter()
            .find(|(name, _)| name == path)
            .unwrap()
            .1
            .clone(),
    )
    .unwrap()
}

fn package_map(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn overlapping_merge_parts() -> Vec<(String, Vec<u8>)> {
    let workbook =
        r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rels = r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let worksheet = r#"<worksheet><sheetData/><mergeCells count="5"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/><mergeCell ref="C3:D4"/><mergeCell ref="D4:E5"/><mergeCell ref="F1:G1"/></mergeCells></worksheet>"#;
    vec![
        ("xl/workbook.xml".to_string(), workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels.as_bytes().to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            worksheet.as_bytes().to_vec(),
        ),
    ]
}

#[test]
fn open_and_recalculation_are_explicit() {
    let cached = Workbook::open(&sample_xlsx()).unwrap();
    assert_eq!(
        cached
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 999.0 }
    );

    let calculated =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    assert_eq!(
        calculated
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 15.0 }
    );

    let mut read_only = Workbook::open_for_read(&sample_xlsx()).unwrap();
    let result = read_only
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert_eq!(result.changed[0].cell, cell("B1"));
}

#[test]
fn defined_names_survive_the_facade_and_drive_incremental_recalculation() {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Number { value: 4.0 },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("B1"),
        Cell {
            value: CellValue::Number { value: 99.0 },
            formula: Some("A1*Rate".into()),
            style: None,
        },
    );
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    model.defined_names.push(DefinedName {
        name: "Rate".into(),
        formula: "2".into(),
        local_sheet: None,
        hidden: false,
    });

    let mut workbook = Workbook::from_model(model).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 8.0 }
    );
    let result = workbook
        .edit_cell(SheetId(0), cell("A1"), "5", CalculationOptions::default())
        .unwrap();
    assert_eq!(result.changed[0].cell, cell("B1"));
    assert_eq!(
        workbook
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 10.0 }
    );

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().defined_names,
        workbook.model().defined_names
    );
}

#[test]
fn structural_edits_rewrite_defined_names_through_save_and_undo() {
    let original = defined_names_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    let before = workbook.model().defined_names.clone();

    workbook
        .apply_ops(
            vec![
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 2,
                },
                Op::InsertCols {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();

    let global = workbook
        .model()
        .defined_names
        .iter()
        .find(|defined| defined.name == "GlobalData")
        .unwrap();
    let local = workbook
        .model()
        .defined_names
        .iter()
        .find(|defined| defined.name == "LocalData")
        .unwrap();
    assert_eq!(global.formula, "Data!$B$3");
    assert_eq!(local.formula, "Data!$B$3");

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().defined_names,
        workbook.model().defined_names
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, before);
}

#[test]
fn structural_edits_refuse_ambiguous_workbook_name_bindings() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Other"));
    model.defined_names.push(DefinedName {
        name: "Input".into(),
        formula: "$A$1".into(),
        local_sheet: None,
        hidden: false,
    });
    let mut workbook = Workbook::from_model(model).unwrap();
    let before = workbook.model().clone();

    let error = workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();

    assert!(error.to_string().contains("cannot be safely rewritten"));
    assert_eq!(workbook.model(), &before);
}

#[test]
fn frozen_panes_survive_the_facade_and_drive_the_initial_view() {
    let mut sheet = Sheet::new("Data");
    sheet.freeze_pane = Some(FreezePane::new(1, 1, cell("D5")));
    sheet.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Text {
                value: "pinned".into(),
            },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        cell("D5"),
        Cell {
            value: CellValue::Text {
                value: "body".into(),
            },
            ..Cell::default()
        },
    );
    let geometry = GridGeometry::new(&sheet);
    let expected_x = geometry.col_x(3) - geometry.col_x(1);
    let expected_y = geometry.row_y(4) - geometry.row_y(1);
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);

    let workbook = Workbook::from_model(model).unwrap();
    let info = workbook.sheet_info().unwrap();
    assert_eq!((info.frozen_rows, info.frozen_cols), (1, 1));
    assert_eq!(
        (info.initial_scroll_x, info.initial_scroll_y),
        (expected_x, expected_y)
    );
    let display = workbook
        .display_list(&Viewport {
            x: info.initial_scroll_x,
            y: info.initial_scroll_y,
            width: 300.0,
            height: 120.0,
        })
        .unwrap();
    assert_eq!(display.grid.col_indices.as_deref().unwrap()[..2], [0, 3]);
    assert_eq!(display.grid.row_indices.as_deref().unwrap()[..2], [0, 4]);

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.sheet(SheetId(0)).unwrap().freeze_pane,
        workbook.sheet(SheetId(0)).unwrap().freeze_pane
    );
}

#[test]
fn hyperlinks_survive_the_facade_and_reach_the_display_list() {
    let mut sheet = Sheet::new("Data");
    sheet.set_cell(
        cell("B2"),
        Cell {
            value: CellValue::Text {
                value: "Website".into(),
            },
            ..Cell::default()
        },
    );
    sheet.hyperlinks.push(Hyperlink {
        range: CellRange::parse_a1("B2:C2").unwrap(),
        external_target: Some("https://example.com".into()),
        location: None,
        tooltip: Some("Open site".into()),
        display: None,
    });
    sheet.hyperlinks.push(Hyperlink {
        range: CellRange::parse_a1("D4").unwrap(),
        external_target: None,
        location: Some("Data!A1".into()),
        tooltip: None,
        display: Some("Jump".into()),
    });
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);

    let workbook = Workbook::from_model(model).unwrap();
    let display = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 120.0,
        })
        .unwrap();
    assert_eq!(display.hyperlinks.len(), 2);
    assert_eq!(
        display.hyperlinks[0].external_target.as_deref(),
        Some("https://example.com")
    );
    assert!(display.commands.iter().any(|command| matches!(
        command,
        DrawCmd::Text {
            text,
            color,
            underline: true,
            ..
        } if text == "Website" && color == "#0563c1"
    )));
    let (x, y) = workbook
        .cell_scroll_position(SheetId(0), cell("D4"))
        .unwrap();
    assert!(x > 0.0 && y > 0.0);

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.sheet(SheetId(0)).unwrap().hyperlinks,
        workbook.sheet(SheetId(0)).unwrap().hyperlinks
    );
}

#[test]
fn combined_hyperlink_location_remaps_and_round_trips() {
    let mut source = Sheet::new("Source");
    source.hyperlinks.push(Hyperlink {
        range: CellRange::parse_a1("B2").unwrap(),
        external_target: Some("https://example.com/report".into()),
        location: Some("Target!A3".into()),
        tooltip: None,
        display: Some("Open report".into()),
    });
    let mut model = WorkbookModel::default();
    model.sheets.push(source);
    model.sheets.push(Sheet::new("Target"));
    let mut workbook = Workbook::from_model(model).unwrap();

    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(1),
                at: 1,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    let hyperlink = &reopened.sheet(SheetId(0)).unwrap().hyperlinks[0];
    assert_eq!(
        hyperlink.external_target.as_deref(),
        Some("https://example.com/report")
    );
    assert_eq!(hyperlink.location.as_deref(), Some("Target!A4"));
}

#[test]
fn renaming_a_sheet_rewrites_hash_prefixed_hyperlink_locations() {
    let mut source = Sheet::new("Source");
    let link = |range: &str, location: &str| Hyperlink {
        range: CellRange::parse_a1(range).unwrap(),
        external_target: None,
        location: Some(location.into()),
        tooltip: None,
        display: None,
    };
    source.hyperlinks.extend([
        link("A1", "#Target!A1"),
        link("A2", "Target!A2"),
        link("A3", "#'Target'!A3"),
        link("A4", "#MyRange"),
    ]);
    source.hyperlinks.push(Hyperlink {
        range: CellRange::parse_a1("A5").unwrap(),
        external_target: Some("https://example.com/report".into()),
        location: Some("#Target!A5".into()),
        tooltip: None,
        display: Some("Open report".into()),
    });
    let mut model = WorkbookModel::default();
    model.sheets.push(source);
    model.sheets.push(Sheet::new("Target"));
    let mut workbook = Workbook::from_model(model).unwrap();

    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(1),
                name: "My Sheet".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    let hyperlinks = &reopened.sheet(SheetId(0)).unwrap().hyperlinks;
    let locations: Vec<Option<&str>> = hyperlinks
        .iter()
        .map(|hyperlink| hyperlink.location.as_deref())
        .collect();
    assert_eq!(
        locations,
        vec![
            Some("#'My Sheet'!A1"),
            Some("'My Sheet'!A2"),
            Some("#'My Sheet'!A3"),
            Some("#MyRange"),
            Some("#'My Sheet'!A5"),
        ]
    );
    assert_eq!(
        hyperlinks[4].external_target.as_deref(),
        Some("https://example.com/report")
    );
}

#[test]
fn edits_recalculate_render_and_round_trip() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    let result = workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert_eq!(result.changed.len(), 1);
    assert_eq!(result.changed[0].cell, cell("B1"));
    assert_eq!(
        workbook
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("A1"))
            .unwrap()
            .style,
        Some(0)
    );

    let display = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 240.0,
            height: 120.0,
        })
        .unwrap();
    assert!(
        display
            .commands
            .iter()
            .any(|command| { matches!(command, DrawCmd::Text { text, .. } if text == "25") })
    );

    #[cfg(feature = "raster")]
    {
        let png = workbook
            .render_sheet(
                SheetId(0),
                &RenderOptions {
                    range: Some(betteroffice_xlsx::CellRange::parse_a1("A1:B2").unwrap()),
                    ..RenderOptions::default()
                },
            )
            .unwrap();
        assert_eq!(
            &png.bytes[..8],
            &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]
        );
    }

    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.cell(SheetId(0), cell("A1")).unwrap().input, "20");
    assert_eq!(
        reopened.cell(SheetId(0), cell("B1")).unwrap().input,
        "=SUM(A1:A2)"
    );
    assert_eq!(
        reopened
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B1"))
            .unwrap()
            .value,
        CellValue::Number { value: 25.0 }
    );
}

#[test]
fn yrs_state_tracks_structural_edits_and_history() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .apply_ops(
            vec![
                Op::AddSheet {
                    index: 1,
                    name: "Inserted".into(),
                },
                Op::SetCell {
                    sheet: SheetId(1),
                    at: cell("C3"),
                    cell: CellState {
                        value: CellValue::Text {
                            value: "shared".into(),
                        },
                        ..CellState::default()
                    },
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.sheet_count(), 3);
    assert_eq!(reopened.sheet_id("Inserted"), Some(SheetId(1)));
    assert_eq!(
        reopened.cell(SheetId(1), cell("C3")).unwrap().input,
        "shared"
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.sheet_count(), 2);
    assert_eq!(reopened.sheet_id("Inserted"), None);

    workbook.redo(CalculationOptions::default()).unwrap();
    let model = workbook.into_model();
    assert_eq!(model.sheets.len(), 3);
    assert_eq!(model.sheets[1].name, "Inserted");
    assert_eq!(
        model.sheets[1].cell(cell("C3")).unwrap().value,
        CellValue::Text {
            value: "shared".into()
        }
    );
}

#[test]
fn standalone_removed_sheet_state_encodes_and_undo_restores_the_model() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    let original = workbook.model().clone();
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();

    assert_eq!(workbook.sheet_count(), 1);
    assert!(!workbook.encode_state_as_update_v1().is_empty());
    assert_eq!(
        Workbook::open(&workbook.save().unwrap())
            .unwrap()
            .sheet_count(),
        1
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model(), &original);
    assert!(!workbook.encode_state_vector_v1().is_empty());
    assert!(!workbook.encode_state_as_update_v1().is_empty());
}

#[test]
fn workbook_remains_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Workbook>();
}

#[test]
fn undo_redo_and_proposals_share_the_typed_session() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert!(workbook.can_undo());
    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "10");
    assert!(
        workbook
            .redo(CalculationOptions::default())
            .unwrap()
            .applied
    );

    let proposal = workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: Some("update total".into()),
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "30".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(proposal.id, "p1");
    assert_eq!(workbook.proposals().len(), 1);
    let accepted = workbook
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert_eq!(accepted.proposal_id, "p1");
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "30");
    assert!(workbook.proposals().is_empty());
}

#[test]
fn pending_proposals_ghost_into_display_lists() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "30".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 240.0,
        height: 120.0,
    };
    let texts = |workbook: &Workbook| -> Vec<(String, String, bool)> {
        workbook
            .display_list(&viewport)
            .unwrap()
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text {
                    text,
                    color,
                    strike,
                    ..
                } => Some((text.clone(), color.clone(), *strike)),
                _ => None,
            })
            .collect()
    };

    let ghosted = texts(&workbook);
    assert!(
        ghosted
            .iter()
            .any(|(text, color, strike)| text == "10" && color == "#c62828" && *strike)
    );
    assert!(
        ghosted
            .iter()
            .any(|(text, color, strike)| text == "30" && color == "#2e7d32" && !*strike)
    );
    assert!(
        !ghosted
            .iter()
            .any(|(text, color, _)| text == "10" && color == "#000000")
    );

    workbook
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    let committed = texts(&workbook);
    assert!(
        committed
            .iter()
            .any(|(text, color, strike)| text == "30" && color == "#000000" && !*strike)
    );
    assert!(!committed.iter().any(|(_, color, _)| color == "#c62828"));
}

#[test]
fn proposal_previews_use_target_number_formats() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();

    let proposal = workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![
                    ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell("A1"),
                        input: "0.484".into(),
                        number_format: Some(NumberFormatMutation::Percent),
                    },
                    ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell("A2"),
                        input: "46204".into(),
                        number_format: Some(NumberFormatMutation::Date),
                    },
                ],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    assert_eq!(proposal.edits[0].old_text, "10");
    assert_eq!(proposal.edits[0].new_text, "48.40%");
    assert_eq!(proposal.edits[1].old_text, "5");
    assert_eq!(proposal.edits[1].new_text, "7/1/2026");

    workbook
        .accept_proposal(&proposal.id, false, CalculationOptions::default())
        .unwrap();
    assert_eq!(
        workbook
            .selection_formatting(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
            .unwrap()
            .number_format,
        Some(NumberFormatKind::Percent)
    );
    assert_eq!(
        workbook
            .selection_formatting(SheetId(0), CellRange::new(cell("A2"), cell("A2")))
            .unwrap()
            .number_format,
        Some(NumberFormatKind::Date)
    );
}

#[test]
fn formula_proposals_keep_the_old_computed_display_value() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    let proposal = workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("B1"),
                    input: "=A2".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    assert_eq!(proposal.edits[0].old_text, "15");
    assert_eq!(proposal.edits[0].new_text, "5");

    let display = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 240.0,
            height: 120.0,
        })
        .unwrap();
    let values: Vec<_> = display
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text {
                text,
                color,
                strike,
                ..
            } if color == "#c62828" || color == "#2e7d32" => {
                Some((text.as_str(), color.as_str(), *strike))
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        values,
        vec![("15", "#c62828", true), ("5", "#2e7d32", false)]
    );
}

#[test]
fn proposal_ghosts_include_recalculated_formula_dependents() {
    let mut workbook =
        Workbook::open_recalculated(&sample_xlsx(), CalculationOptions::default()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "20".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    let display = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 240.0,
            height: 120.0,
        })
        .unwrap();
    let values: Vec<_> = display
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text {
                text,
                color,
                strike,
                ..
            } if color == "#c62828" || color == "#2e7d32" => {
                Some((text.as_str(), color.as_str(), *strike))
            }
            _ => None,
        })
        .collect();

    assert!(values.contains(&("10", "#c62828", true)));
    assert!(values.contains(&("20", "#2e7d32", false)));
    assert!(values.contains(&("15", "#c62828", true)));
    assert!(values.contains(&("25", "#2e7d32", false)));
}

#[test]
fn rejects_empty_workbook_ops_atomically() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    let result = workbook.apply_ops(
        vec![Op::RemoveSheet { index: 1 }, Op::RemoveSheet { index: 0 }],
        CalculationOptions::default(),
    );
    assert!(matches!(result, Err(Error::NoSheets)));
    assert_eq!(workbook.sheet_count(), 2);
}

#[test]
fn rejects_overlapping_merged_ranges() {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Data");
    sheet.merges = vec![
        CellRange::parse_a1("A1:B2").unwrap(),
        CellRange::parse_a1("B2:C3").unwrap(),
    ];
    model.sheets.push(sheet);

    assert!(matches!(
        Workbook::from_model(model),
        Err(Error::InvalidOperation(message))
            if message == "workbook contains overlapping merged ranges"
    ));
}

#[test]
fn parsed_overlapping_merges_open_and_save_normalized() {
    let model = xlsx_parse::parse_workbook(&overlapping_merge_parts()).unwrap();
    let merges: Vec<_> = model.sheets[0]
        .merges
        .iter()
        .map(|range| range.to_a1())
        .collect();
    assert_eq!(merges, ["A1:B2", "C3:D4", "F1:G1"]);

    let workbook = Workbook::from_model(model).unwrap();
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet_xml = parts
        .iter()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .map(|(_, bytes)| std::str::from_utf8(bytes).unwrap())
        .unwrap();
    assert!(sheet_xml.contains(
        r#"<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="C3:D4"/><mergeCell ref="F1:G1"/></mergeCells>"#
    ));

    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0].merges,
        workbook.model().sheets[0].merges
    );
}

#[test]
fn validates_raw_ops_and_noop_history() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    let result = workbook.edit_cells(
        SheetId(0),
        &Vec::<CellInput>::new(),
        CalculationOptions::default(),
    );
    assert!(!result.unwrap().applied);
    assert!(!workbook.can_undo());

    let invalid = workbook.apply_ops(
        vec![Op::SetColWidth {
            sheet: SheetId(0),
            col: 1_000_000_000,
            width: Some(12.0),
        }],
        CalculationOptions::default(),
    );
    assert!(matches!(invalid, Err(Error::InvalidOperation(_))));
    assert!(!workbook.can_undo());

    let duplicate_name = workbook.apply_ops(
        vec![Op::RenameSheet {
            sheet: SheetId(0),
            name: "Empty".into(),
        }],
        CalculationOptions::default(),
    );
    assert!(matches!(duplicate_name, Err(Error::InvalidOperation(_))));

    let shifted_dimension = workbook.apply_ops(
        vec![
            Op::SetRowHeight {
                sheet: SheetId(0),
                row: betteroffice_xlsx::MAX_ROWS - 1,
                height: Some(20.0),
            },
            Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: betteroffice_xlsx::MAX_ROWS,
            },
        ],
        CalculationOptions::default(),
    );
    assert!(matches!(shifted_dimension, Err(Error::InvalidOperation(_))));
    assert!(!workbook.can_undo());
}

#[test]
fn semantic_noop_preserves_redo_history() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert!(workbook.can_redo());

    let formula_result = workbook
        .edit_cells(
            SheetId(0),
            &[
                CellInput {
                    cell: cell("B1"),
                    input: "=1".into(),
                },
                CellInput {
                    cell: cell("B1"),
                    input: "=SUM(A1:A2)".into(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(!formula_result.applied);
    assert!(workbook.can_redo());

    let result = workbook
        .edit_cell(SheetId(0), cell("A1"), "10", CalculationOptions::default())
        .unwrap();
    assert!(!result.applied);
    assert!(workbook.can_redo());
}

#[test]
fn rejects_insertions_that_discard_boundary_content() {
    let mut sheet = Sheet::new("Data");
    let last_row = CellRef::new(betteroffice_xlsx::MAX_ROWS - 1, 0);
    let last_col = CellRef::new(0, betteroffice_xlsx::MAX_COLS - 1);
    sheet.set_cell(
        last_row,
        Cell {
            value: CellValue::Text {
                value: "row edge".into(),
            },
            ..Cell::default()
        },
    );
    sheet.set_cell(
        last_col,
        Cell {
            value: CellValue::Text {
                value: "column edge".into(),
            },
            ..Cell::default()
        },
    );
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    let mut workbook = Workbook::from_model(model).unwrap();

    let row_error = workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(matches!(row_error, Error::InvalidOperation(_)));
    assert_eq!(
        workbook.cell(SheetId(0), last_row).unwrap().input,
        "row edge"
    );

    let column_error = workbook
        .apply_ops(
            vec![Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(matches!(column_error, Error::InvalidOperation(_)));
    assert_eq!(
        workbook.cell(SheetId(0), last_col).unwrap().input,
        "column edge"
    );
    assert!(!workbook.can_undo());
}

#[test]
fn rejects_reversed_ranges_and_oversized_dimensions() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    let reversed = CellRange {
        start: cell("B2"),
        end: cell("A1"),
    };
    assert!(matches!(
        workbook.range_cells(SheetId(0), reversed),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.apply_ops(
            vec![Op::MergeCells {
                sheet: SheetId(0),
                range: reversed,
            }],
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.apply_ops(
            vec![Op::SetColWidth {
                sheet: SheetId(0),
                col: 0,
                width: Some(256.0),
            }],
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.apply_ops(
            vec![Op::SetRowHeight {
                sheet: SheetId(0),
                row: 0,
                height: Some(410.0),
            }],
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
    assert!(matches!(
        workbook.edit_cell(
            SheetId(0),
            cell("A1"),
            &"x".repeat(xlsx_calc::eval::MAX_CELL_TEXT_CHARS + 1),
            CalculationOptions::default(),
        ),
        Err(Error::InvalidOperation(_))
    ));
}

#[test]
fn proposal_staleness_uses_cell_state_not_display_text() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("B1"),
                    input: "1".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("B1"),
            "=999",
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(matches!(
        workbook.accept_proposal("p1", false, CalculationOptions::default()),
        Err(Error::StaleProposal(_))
    ));
}

#[test]
fn proposal_acceptance_applies_duplicate_targets_sequentially() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![
                    ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell("A1"),
                        input: "20".into(),
                        number_format: None,
                    },
                    ProposalEditInput {
                        sheet: SheetId(0),
                        cell: cell("A1"),
                        input: "30".into(),
                        number_format: None,
                    },
                ],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "30");
}

#[test]
fn rename_invalidates_pending_proposals() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "=Data!A2".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(workbook.proposals().is_empty());
}

#[test]
fn reports_recalculation_limits_without_overwriting_cached_values() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    let mut formulas = Sheet::new("Formulas");
    formulas.set_cell(
        cell("A1"),
        Cell {
            value: CellValue::Number { value: 123.0 },
            formula: Some("SUM(Data!A1:XFD1048576)".into()),
            ..Cell::default()
        },
    );
    model.sheets.push(formulas);
    let bytes = ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap();
    let workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default()).unwrap();
    assert_eq!(
        workbook.model().sheets[1].cell(cell("A1")).unwrap().value,
        CellValue::Number { value: 123.0 }
    );
    assert_eq!(workbook.last_calculation().limited_cells.len(), 1);
}

#[test]
fn structural_ops_invalidate_coordinate_proposals() {
    let mut workbook = Workbook::open(&sample_xlsx()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "30".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(workbook.proposals().is_empty());
}

#[test]
fn display_lists_do_not_inherit_raster_dimension_caps() {
    let workbook = Workbook::open(&sample_xlsx()).unwrap();
    assert!(
        workbook
            .display_list(&Viewport {
                x: 0.0,
                y: 0.0,
                width: 20_000.0,
                height: 120.0,
            })
            .is_ok()
    );
}

#[test]
fn display_lists_reject_excessive_cell_spans() {
    let workbook = Workbook::open(&sample_xlsx()).unwrap();
    let error = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 6_000_000.0,
        })
        .unwrap_err();
    assert!(matches!(error, Error::DisplayTooLarge { .. }));
}

#[cfg(feature = "raster")]
#[test]
fn raster_rejects_excessive_total_pixel_area() {
    let workbook = Workbook::open(&sample_xlsx()).unwrap();
    let error = workbook
        .render_png(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 5_000.0,
            height: 5_000.0,
        })
        .unwrap_err();
    assert!(matches!(error, Error::RenderAreaTooLarge { .. }));
}

#[test]
fn collaboration_vectors_diffs_and_deterministic_baseline_handshake() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 101).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 202).unwrap();

    assert_eq!(left.client_id(), 101);
    assert_eq!(right.client_id(), 202);
    assert_ne!(left.client_id(), right.client_id());
    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_eq!(
        left.encode_state_as_update_v1(),
        right.encode_state_as_update_v1()
    );
    assert_eq!(
        left.encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap(),
        &[0, 0]
    );

    left.edit_cell(SheetId(0), cell("A1"), "21", CalculationOptions::default())
        .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    assert!(
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(left.model(), right.model());
    assert!(
        !left
            .apply_update_v1(
                &right
                    .encode_diff_v1(&left.encode_state_vector_v1())
                    .unwrap(),
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
}

#[test]
fn duplicate_runtime_client_ids_are_an_invalid_host_configuration() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 211).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 211).unwrap();
    let baseline = left.encode_state_vector_v1();

    left.edit_cell(
        SheetId(0),
        cell("C1"),
        "left",
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .edit_cell(
            SheetId(0),
            cell("C2"),
            "right",
            CalculationOptions::default(),
        )
        .unwrap();
    let from_left = left.encode_diff_v1(&baseline).unwrap();
    let from_right = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&from_right, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&from_left, CalculationOptions::default())
        .unwrap();

    assert_eq!(
        left.encode_state_vector_v1(),
        right.encode_state_vector_v1()
    );
    assert_ne!(
        left.encode_state_as_update_v1(),
        right.encode_state_as_update_v1()
    );
    assert_ne!(left.model(), right.model());
}

#[test]
fn collaborative_undo_redo_track_only_local_user_edits() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 221).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    assert!(workbook.can_undo());
    assert!(!workbook.can_redo());
    assert_eq!(workbook.history_state().undo_depth, 1);

    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "10");
    assert!(!workbook.can_undo());
    assert!(workbook.can_redo());
    assert!(
        workbook
            .redo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.cell(SheetId(0), cell("A1")).unwrap().input, "20");

    let mut agent_only = Workbook::open_collaborative(&bytes, 222).unwrap();
    agent_only
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A2"),
                    input: "30".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    agent_only
        .accept_proposal("p1", false, CalculationOptions::default())
        .unwrap();
    assert!(!agent_only.can_undo());
    assert!(
        !agent_only
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
}

#[test]
fn collaborative_undo_converges_after_a_concurrent_remote_edit() {
    let bytes = sample_xlsx();
    for (left_id, right_id) in [(231, 230), (230, 231)] {
        let mut left = Workbook::open_collaborative(&bytes, left_id).unwrap();
        let mut right = Workbook::open_collaborative(&bytes, right_id).unwrap();
        let baseline = left.encode_state_vector_v1();

        left.edit_cell(
            SheetId(0),
            cell("C1"),
            "left",
            CalculationOptions::default(),
        )
        .unwrap();
        right
            .edit_cell(
                SheetId(0),
                cell("C1"),
                "right",
                CalculationOptions::default(),
            )
            .unwrap();
        let from_left = left.encode_diff_v1(&baseline).unwrap();
        let from_right = right.encode_diff_v1(&baseline).unwrap();
        left.apply_update_v1(&from_right, CalculationOptions::default())
            .unwrap();
        right
            .apply_update_v1(&from_left, CalculationOptions::default())
            .unwrap();

        let right_before_undo = right.encode_state_vector_v1();
        left.undo(CalculationOptions::default()).unwrap();
        let undo = left.encode_diff_v1(&right_before_undo).unwrap();
        right
            .apply_update_v1(&undo, CalculationOptions::default())
            .unwrap();
        assert_eq!(left.model(), right.model());
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert!(right.can_undo());
    }
}

#[test]
fn concurrent_disjoint_and_same_cell_edits_converge() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 301).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 302).unwrap();
    let baseline = left.encode_state_vector_v1();

    left.edit_cell(SheetId(0), cell("A1"), "20", CalculationOptions::default())
        .unwrap();
    right
        .edit_cell(SheetId(0), cell("A2"), "7", CalculationOptions::default())
        .unwrap();
    let from_left = left.encode_diff_v1(&baseline).unwrap();
    let from_right = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&from_right, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&from_left, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model(), right.model());
    assert_eq!(left.cell(SheetId(0), cell("A1")).unwrap().input, "20");
    assert_eq!(left.cell(SheetId(0), cell("A2")).unwrap().input, "7");

    let left_before = left.encode_state_vector_v1();
    let right_before = right.encode_state_vector_v1();
    left.edit_cell(
        SheetId(0),
        cell("C1"),
        "left",
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .edit_cell(
            SheetId(0),
            cell("C1"),
            "right",
            CalculationOptions::default(),
        )
        .unwrap();
    let from_left = left.encode_diff_v1(&right_before).unwrap();
    let from_right = right.encode_diff_v1(&left_before).unwrap();
    left.apply_update_v1(&from_right, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&from_left, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model(), right.model());
    assert!(matches!(
        left.cell(SheetId(0), cell("C1")).unwrap().input.as_str(),
        "left" | "right"
    ));
}

#[test]
fn concurrent_style_and_content_changes_compose() {
    let bytes = sample_xlsx();
    let mut content = Workbook::open_collaborative(&bytes, 401).unwrap();
    let mut style = Workbook::open_collaborative(&bytes, 402).unwrap();
    let baseline = content.encode_state_vector_v1();

    content
        .edit_cell(SheetId(0), cell("A1"), "25", CalculationOptions::default())
        .unwrap();
    style
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("A1"), cell("A1")),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let content_update = content.encode_diff_v1(&baseline).unwrap();
    let style_update = style.encode_diff_v1(&baseline).unwrap();
    content
        .apply_update_v1(&style_update, CalculationOptions::default())
        .unwrap();
    style
        .apply_update_v1(&content_update, CalculationOptions::default())
        .unwrap();

    assert_eq!(content.model(), style.model());
    let composed = content
        .model()
        .sheet(SheetId(0))
        .unwrap()
        .cell(cell("A1"))
        .unwrap();
    assert_eq!(composed.value, CellValue::Number { value: 25.0 });
    assert_eq!(
        content
            .selection_formatting(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
            .unwrap()
            .bold,
        Some(true)
    );
}

#[test]
fn collaborative_formatting_round_trips_and_matches_aggregation() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 403).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 404).unwrap();
    let range = CellRange::new(cell("A1"), cell("B2"));

    left.patch_range_style(
        SheetId(0),
        range,
        StylePatch {
            bold: Some(true),
            fill_color: Some("#ffcc00".into()),
            text_color: Some("#123456".into()),
            ..StylePatch::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    left.set_range_number_format(
        SheetId(0),
        range,
        NumberFormatMutation::Custom {
            pattern: "0.0000".into(),
        },
        CalculationOptions::default(),
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    assert_eq!(
        left.selection_formatting(SheetId(0), range).unwrap(),
        right.selection_formatting(SheetId(0), range).unwrap()
    );
    let formatting = right.selection_formatting(SheetId(0), range).unwrap();
    assert_eq!(formatting.bold, Some(true));
    assert_eq!(formatting.fill_color.as_deref(), Some("#ffcc00"));
    assert_eq!(formatting.text_color.as_deref(), Some("#123456"));
    assert_eq!(formatting.number_format, Some(NumberFormatKind::Custom));
    assert_eq!(formatting.number_format_pattern.as_deref(), Some("0.0000"));
}

#[test]
fn concurrent_formatting_restyles_converge_with_all_formats_available() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 405).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 406).unwrap();
    let baseline = left.encode_state_vector_v1();
    let range = CellRange::new(cell("A1"), cell("B2"));

    left.patch_range_style(
        SheetId(0),
        range,
        StylePatch {
            bold: Some(true),
            text_color: Some("#aa0000".into()),
            ..StylePatch::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .patch_range_style(
            SheetId(0),
            range,
            StylePatch {
                italic: Some(true),
                fill_color: Some("#00aa00".into()),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let left_update = left.encode_diff_v1(&baseline).unwrap();
    let right_update = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&right_update, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&left_update, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    assert_eq!(
        left.encode_state_as_update_v1(),
        right.encode_state_as_update_v1()
    );
    assert_eq!(left.model().styles.cell_xfs.len(), 3);
    assert_eq!(
        left.selection_formatting(SheetId(0), range).unwrap(),
        right.selection_formatting(SheetId(0), range).unwrap()
    );
}

#[test]
fn concurrent_identical_formats_are_content_deduplicated() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 407).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 408).unwrap();
    let baseline = left.encode_state_vector_v1();
    let patch = StylePatch {
        bold: Some(true),
        font_family: Some("Inter".into()),
        ..StylePatch::default()
    };

    left.patch_range_style(
        SheetId(0),
        CellRange::new(cell("A1"), cell("A1")),
        patch.clone(),
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("A2"), cell("A2")),
            patch,
            CalculationOptions::default(),
        )
        .unwrap();
    let left_update = left.encode_diff_v1(&baseline).unwrap();
    let right_update = right.encode_diff_v1(&baseline).unwrap();
    left.apply_update_v1(&right_update, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&left_update, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    let sheet = &left.model().sheets[0];
    assert_eq!(
        sheet.cell(cell("A1")).unwrap().style,
        sheet.cell(cell("A2")).unwrap().style
    );
    assert_eq!(left.model().styles.cell_xfs.len(), 2);
}

#[test]
fn collaborative_formatting_undo_is_local_origin_only() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 409).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 410).unwrap();

    left.patch_range_style(
        SheetId(0),
        CellRange::new(cell("A1"), cell("A1")),
        StylePatch {
            bold: Some(true),
            ..StylePatch::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .patch_range_style(
            SheetId(0),
            CellRange::new(cell("A2"), cell("A2")),
            StylePatch {
                fill_color: Some("#abcdef".into()),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let left_update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let right_update = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    left.apply_update_v1(&right_update, CalculationOptions::default())
        .unwrap();
    right
        .apply_update_v1(&left_update, CalculationOptions::default())
        .unwrap();
    let format = right
        .capture_format(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
        .unwrap();
    right
        .apply_format(
            SheetId(0),
            CellRange::new(cell("A3"), cell("A3")),
            format,
            CalculationOptions::default(),
        )
        .unwrap();
    let reused_format = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    left.apply_update_v1(&reused_format, CalculationOptions::default())
        .unwrap();
    let right_before_undo = right.encode_state_vector_v1();

    assert!(left.undo(CalculationOptions::default()).unwrap().applied);
    let undo = left.encode_diff_v1(&right_before_undo).unwrap();
    right
        .apply_update_v1(&undo, CalculationOptions::default())
        .unwrap();

    assert_eq!(left.model(), right.model());
    let a1 = left
        .selection_formatting(SheetId(0), CellRange::new(cell("A1"), cell("A1")))
        .unwrap();
    let a2 = left
        .selection_formatting(SheetId(0), CellRange::new(cell("A2"), cell("A2")))
        .unwrap();
    let a3 = left
        .selection_formatting(SheetId(0), CellRange::new(cell("A3"), cell("A3")))
        .unwrap();
    assert_eq!(a1.bold, Some(false));
    assert_eq!(a2.fill_color.as_deref(), Some("#abcdef"));
    assert_eq!(a3.bold, Some(true));
}

#[test]
fn style_edits_do_not_publish_recalculated_formula_caches_as_content() {
    let bytes = sample_xlsx();
    for (formula_client, style_client) in [(411, 412), (422, 421)] {
        let mut formula = Workbook::open_collaborative_recalculated(
            &bytes,
            formula_client,
            CalculationOptions::default(),
        )
        .unwrap();
        let mut style = Workbook::open_collaborative_recalculated(
            &bytes,
            style_client,
            CalculationOptions::default(),
        )
        .unwrap();
        let baseline = formula.encode_state_vector_v1();

        formula
            .edit_cell(
                SheetId(0),
                cell("B1"),
                "=SUM(A1:A2)+1",
                CalculationOptions::default(),
            )
            .unwrap();
        style
            .apply_ops(
                vec![Op::SetCell {
                    sheet: SheetId(0),
                    at: cell("B1"),
                    cell: CellState {
                        value: CellValue::Number { value: 15.0 },
                        formula: Some("SUM(A1:A2)".into()),
                        style: Some(0),
                    },
                }],
                CalculationOptions::default(),
            )
            .unwrap();

        let formula_update = formula.encode_diff_v1(&baseline).unwrap();
        let style_update = style.encode_diff_v1(&baseline).unwrap();
        formula
            .apply_update_v1(&style_update, CalculationOptions::default())
            .unwrap();
        style
            .apply_update_v1(&formula_update, CalculationOptions::default())
            .unwrap();

        assert_eq!(formula.model(), style.model());
        let composed = formula.model().sheets[0].cell(cell("B1")).unwrap();
        assert_eq!(composed.formula.as_deref(), Some("SUM(A1:A2)+1"));
        assert_eq!(composed.style, Some(0));
    }
}

#[test]
fn remote_formulas_recalculate_locally_and_save_current_caches() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 501).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 502).unwrap();

    left.edit_cell(
        SheetId(0),
        cell("B1"),
        "=A1*2",
        CalculationOptions::default(),
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(right.cell(SheetId(0), cell("B1")).unwrap().input, "=A1*2");
    assert_eq!(
        right.model().sheets[0].cell(cell("B1")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );

    let shared_before_recalc = right.encode_state_as_update_v1();
    right.recalculate_all(CalculationOptions::default());
    assert_eq!(right.encode_state_as_update_v1(), shared_before_recalc);
    let reopened = Workbook::open(&right.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().sheets[0].cell(cell("B1")).unwrap().value,
        CellValue::Number { value: 20.0 }
    );
}

#[test]
fn remote_changed_cells_compare_against_the_current_projection() {
    let bytes = sample_xlsx();
    let options = CalculationOptions::default();
    let mut left = Workbook::open_collaborative_recalculated(&bytes, 511, options).unwrap();
    let mut right = Workbook::open_collaborative_recalculated(&bytes, 512, options).unwrap();

    left.edit_cell(SheetId(0), cell("A1"), "20", options)
        .unwrap();
    let first = right
        .apply_update_v1(
            &left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap(),
            options,
        )
        .unwrap();
    assert_eq!(
        first.changed,
        [
            betteroffice_xlsx::CellAddress {
                sheet: SheetId(0),
                cell: cell("A1"),
            },
            betteroffice_xlsx::CellAddress {
                sheet: SheetId(0),
                cell: cell("B1"),
            },
        ]
    );

    left.edit_cell(SheetId(1), cell("A1"), "unrelated", options)
        .unwrap();
    let second = right
        .apply_update_v1(
            &left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap(),
            options,
        )
        .unwrap();
    assert_eq!(
        second.changed,
        [betteroffice_xlsx::CellAddress {
            sheet: SheetId(1),
            cell: cell("A1"),
        }]
    );
}

#[test]
fn duplicate_and_reversed_update_delivery_are_safe() {
    let bytes = sample_xlsx();
    let mut source = Workbook::open_collaborative(&bytes, 601).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 602).unwrap();
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();

    source
        .edit_cell(SheetId(0), cell("A1"), "31", CalculationOptions::default())
        .unwrap();
    source
        .edit_cell(SheetId(0), cell("A2"), "9", CalculationOptions::default())
        .unwrap();
    let updates = events
        .lock()
        .unwrap()
        .iter()
        .map(|event| event.update.clone())
        .collect::<Vec<_>>();
    assert_eq!(updates.len(), 2);

    let remote_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&remote_events);
    let _remote_subscription = target
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    assert!(
        target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("A2")).unwrap().input, "9");
    assert_eq!(remote_events.lock().unwrap().len(), 1);
    assert!(
        target
            .apply_update_v1(&updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.model(), source.model());
    assert_eq!(remote_events.lock().unwrap().len(), 2);
    assert_eq!(
        remote_events.lock().unwrap()[0].origin,
        UpdateOrigin::Remote
    );
    assert!(
        !target
            .apply_update_v1(&updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(remote_events.lock().unwrap().len(), 2);
}

#[test]
fn malformed_and_structural_remote_updates_roll_back_every_facade_state() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 701).unwrap();
    workbook.set_active_sheet(SheetId(1)).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A2"), "8", CalculationOptions::default())
        .unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "99".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    let assert_unchanged =
        |workbook: &Workbook,
         model: &WorkbookModel,
         state: &[u8],
         calculation: &betteroffice_xlsx::CalculationResult| {
            assert_eq!(workbook.model(), model);
            assert_eq!(workbook.encode_state_as_update_v1(), state);
            assert_eq!(workbook.active_sheet(), SheetId(1));
            assert!(workbook.can_undo());
            assert!(!workbook.can_redo());
            assert_eq!(workbook.proposals().len(), 1);
            assert_eq!(workbook.last_calculation(), calculation);
        };
    let model = workbook.model().clone();
    let state = workbook.encode_state_as_update_v1();
    let calculation = workbook.last_calculation().clone();
    assert!(matches!(
        workbook.apply_update_v1(&[0xff], CalculationOptions::default()),
        Err(Error::InvalidUpdate(_))
    ));
    assert_unchanged(&workbook, &model, &state, &calculation);

    let mut structural = Workbook::open(&bytes).unwrap();
    structural
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let update = structural
        .encode_diff_v1(&workbook.encode_state_vector_v1())
        .unwrap();
    assert!(matches!(
        workbook.apply_update_v1(&update, CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));
    assert_unchanged(&workbook, &model, &state, &calculation);

    let mut shifted = Workbook::open(&bytes).unwrap();
    shifted
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let update = shifted
        .encode_diff_v1(&workbook.encode_state_vector_v1())
        .unwrap();
    assert!(matches!(
        workbook.apply_update_v1(&update, CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));
    assert_unchanged(&workbook, &model, &state, &calculation);
}

#[test]
fn rejected_update_preserves_unrelated_valid_causal_backlog() {
    let bytes = sample_xlsx();
    let mut source = Workbook::open_collaborative(&bytes, 741).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 742).unwrap();
    let updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&updates);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    source
        .edit_cell(SheetId(0), cell("C3"), "one", CalculationOptions::default())
        .unwrap();
    source
        .edit_cell(SheetId(0), cell("C3"), "two", CalculationOptions::default())
        .unwrap();
    let updates = updates.lock().unwrap().clone();
    assert_eq!(updates.len(), 2);
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );

    let mut structural = Workbook::open(&bytes).unwrap();
    structural
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let invalid = structural
        .encode_diff_v1(&target.encode_state_vector_v1())
        .unwrap();
    assert!(matches!(
        target.apply_update_v1(&invalid, CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));

    assert!(
        target
            .apply_update_v1(&updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("C3")).unwrap().input, "two");
}

#[test]
fn independent_pending_chains_resolve_without_blocking_each_other() {
    let bytes = sample_xlsx();
    let mut first = Workbook::open_collaborative(&bytes, 743).unwrap();
    let mut second = Workbook::open_collaborative(&bytes, 744).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 745).unwrap();
    let first_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&first_updates);
    let _first_subscription = first
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    let second_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&second_updates);
    let _second_subscription = second
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    first
        .edit_cell(SheetId(0), cell("C4"), "one", CalculationOptions::default())
        .unwrap();
    first
        .edit_cell(SheetId(0), cell("C4"), "two", CalculationOptions::default())
        .unwrap();
    second
        .edit_cell(
            SheetId(0),
            cell("C5"),
            "three",
            CalculationOptions::default(),
        )
        .unwrap();
    second
        .edit_cell(
            SheetId(0),
            cell("C5"),
            "four",
            CalculationOptions::default(),
        )
        .unwrap();
    let first_updates = first_updates.lock().unwrap().clone();
    let second_updates = second_updates.lock().unwrap().clone();

    assert!(
        !target
            .apply_update_v1(&first_updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        !target
            .apply_update_v1(&second_updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        target
            .apply_update_v1(&second_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("C5")).unwrap().input, "four");
    assert_eq!(target.cell(SheetId(0), cell("C4")).unwrap().input, "");

    assert!(
        target
            .apply_update_v1(&first_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("C4")).unwrap().input, "two");
}

#[test]
fn applicable_clients_in_a_partially_pending_update_are_not_blocked() {
    let bytes = sample_xlsx();
    let mut delayed = Workbook::open_collaborative(&bytes, 746).unwrap();
    let mut ready = Workbook::open_collaborative(&bytes, 747).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 748).unwrap();
    let delayed_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&delayed_updates);
    let _delayed_subscription = delayed
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    let ready_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&ready_updates);
    let _ready_subscription = ready
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    delayed
        .edit_cell(SheetId(0), cell("D1"), "one", CalculationOptions::default())
        .unwrap();
    delayed
        .edit_cell(SheetId(0), cell("D1"), "two", CalculationOptions::default())
        .unwrap();
    ready
        .edit_cell(
            SheetId(0),
            cell("D2"),
            "ready",
            CalculationOptions::default(),
        )
        .unwrap();
    let delayed_updates = delayed_updates.lock().unwrap().clone();
    let ready_updates = ready_updates.lock().unwrap().clone();
    let merged = YrsUpdate::merge_updates([
        YrsUpdate::decode_v1(&delayed_updates[1]).unwrap(),
        YrsUpdate::decode_v1(&ready_updates[0]).unwrap(),
    ])
    .encode_v1();

    assert!(
        target
            .apply_update_v1(&merged, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D2")).unwrap().input, "ready");
    assert_eq!(target.cell(SheetId(0), cell("D1")).unwrap().input, "");

    assert!(
        target
            .apply_update_v1(&delayed_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D1")).unwrap().input, "two");
}

#[test]
fn newly_applicable_clients_in_a_buffered_merged_update_are_committed() {
    let bytes = sample_xlsx();
    let mut delayed = Workbook::open_collaborative(&bytes, 749).unwrap();
    let mut ready = Workbook::open_collaborative(&bytes, 750).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 753).unwrap();
    let delayed_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&delayed_updates);
    let _delayed_subscription = delayed
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    let ready_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&ready_updates);
    let _ready_subscription = ready
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();

    delayed
        .edit_cell(SheetId(0), cell("D3"), "one", CalculationOptions::default())
        .unwrap();
    delayed
        .edit_cell(SheetId(0), cell("D3"), "two", CalculationOptions::default())
        .unwrap();
    ready
        .edit_cell(
            SheetId(0),
            cell("D4"),
            "three",
            CalculationOptions::default(),
        )
        .unwrap();
    ready
        .edit_cell(
            SheetId(0),
            cell("D4"),
            "four",
            CalculationOptions::default(),
        )
        .unwrap();
    let delayed_updates = delayed_updates.lock().unwrap().clone();
    let ready_updates = ready_updates.lock().unwrap().clone();
    let merged = YrsUpdate::merge_updates([
        YrsUpdate::decode_v1(&delayed_updates[1]).unwrap(),
        YrsUpdate::decode_v1(&ready_updates[1]).unwrap(),
    ])
    .encode_v1();

    assert!(
        !target
            .apply_update_v1(&merged, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert!(
        target
            .apply_update_v1(&ready_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D4")).unwrap().input, "four");
    assert_eq!(target.cell(SheetId(0), cell("D3")).unwrap().input, "");
    let mut mirror = Workbook::open_collaborative(&bytes, 759).unwrap();
    mirror
        .apply_update_v1(
            &target.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(mirror.model(), target.model());

    assert!(
        target
            .apply_update_v1(&delayed_updates[0], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.cell(SheetId(0), cell("D3")).unwrap().input, "two");
}

#[test]
fn wholly_pending_updates_do_not_reemit_existing_tombstones() {
    let bytes = sample_xlsx();
    let mut remote = Workbook::open_collaborative(&bytes, 754).unwrap();
    let mut local = Workbook::open_collaborative(&bytes, 755).unwrap();
    local
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "local",
            CalculationOptions::default(),
        )
        .unwrap();
    local
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A2"),
                    input: "proposal".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let remote_updates = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&remote_updates);
    let _remote_subscription = remote
        .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
        .unwrap();
    remote
        .edit_cell(SheetId(0), cell("E1"), "one", CalculationOptions::default())
        .unwrap();
    remote
        .edit_cell(SheetId(0), cell("E1"), "two", CalculationOptions::default())
        .unwrap();
    let pending = remote_updates.lock().unwrap()[1].clone();
    let local_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&local_events);
    let _local_subscription = local
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    let state = local.encode_state_as_update_v1();

    let result = local
        .apply_update_v1(&pending, CalculationOptions::default())
        .unwrap();
    assert!(!result.applied);
    assert_eq!(local.encode_state_as_update_v1(), state);
    assert_eq!(local.proposals().len(), 1);
    assert!(local_events.lock().unwrap().is_empty());
}

#[test]
fn unresolved_invalid_updates_never_enter_live_yrs_state() {
    let bytes = sample_xlsx();
    let mut source = Workbook::open(&bytes).unwrap();
    let mut target = Workbook::open_collaborative(&bytes, 751).unwrap();
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = source
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();

    source
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Added".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    source
        .edit_cell(SheetId(1), cell("A1"), "17", CalculationOptions::default())
        .unwrap();
    let updates = events
        .lock()
        .unwrap()
        .iter()
        .map(|event| event.update.clone())
        .collect::<Vec<_>>();
    assert_eq!(updates.len(), 2);

    let state = target.encode_state_as_update_v1();
    assert!(
        !target
            .apply_update_v1(&updates[1], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(target.encode_state_as_update_v1(), state);
    assert!(matches!(
        target.apply_update_v1(&updates[0], CalculationOptions::default()),
        Err(Error::CollaborativeStructureChanged)
    ));
    assert_eq!(target.encode_state_as_update_v1(), state);
    assert_eq!(target.sheet_id("Data"), Some(SheetId(0)));

    let mut valid = Workbook::open_collaborative(&bytes, 752).unwrap();
    valid
        .edit_cell(SheetId(0), cell("A2"), "18", CalculationOptions::default())
        .unwrap();
    let update = valid
        .encode_diff_v1(&target.encode_state_vector_v1())
        .unwrap();
    assert!(
        target
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
    );
}

#[test]
fn effective_remote_updates_clear_local_proposals() {
    let bytes = sample_xlsx();
    let mut remote = Workbook::open_collaborative(&bytes, 801).unwrap();
    let mut local = Workbook::open_collaborative(&bytes, 802).unwrap();
    local.set_active_sheet(SheetId(1)).unwrap();
    local
        .edit_cell(SheetId(0), cell("A2"), "11", CalculationOptions::default())
        .unwrap();
    assert!(local.can_undo());
    assert!(!local.can_redo());
    local
        .propose(
            ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: cell("A1"),
                    input: "40".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();

    remote
        .edit_cell(SheetId(0), cell("A1"), "44", CalculationOptions::default())
        .unwrap();
    let update = remote
        .encode_diff_v1(&local.encode_state_vector_v1())
        .unwrap();
    local
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert!(local.can_undo());
    assert!(!local.can_redo());
    assert!(local.proposals().is_empty());
    assert_eq!(local.active_sheet(), SheetId(1));
    assert!(local.undo(CalculationOptions::default()).unwrap().applied);
    assert_eq!(local.cell(SheetId(0), cell("A1")).unwrap().input, "44");
    assert_eq!(local.cell(SheetId(0), cell("A2")).unwrap().input, "5");
}

#[test]
fn update_observers_receive_one_owned_event_with_classified_origin() {
    let bytes = sample_xlsx();
    let mut left = Workbook::open_collaborative(&bytes, 901).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 902).unwrap();
    let local_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&local_events);
    let local_subscription = left
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();

    left.edit_cells(
        SheetId(0),
        &[
            CellInput {
                cell: cell("A1"),
                input: "12".into(),
            },
            CellInput {
                cell: cell("A2"),
                input: "6".into(),
            },
        ],
        CalculationOptions::default(),
    )
    .unwrap();
    left.recalculate_all(CalculationOptions::default());
    let local_update = {
        let events = local_events.lock().unwrap();
        assert_eq!(events.len(), 1);
        assert_eq!(events[0].origin, UpdateOrigin::Local);
        events[0].update.clone()
    };

    let remote_events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&remote_events);
    let _remote_subscription = right
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    right
        .apply_update_v1(&local_update, CalculationOptions::default())
        .unwrap();
    let events = remote_events.lock().unwrap();
    assert_eq!(events.len(), 1);
    assert_eq!(events[0].origin, UpdateOrigin::Remote);
    assert_eq!(events[0].update, local_update);
    drop(events);

    drop(local_subscription);
    left.edit_cell(SheetId(0), cell("A1"), "13", CalculationOptions::default())
        .unwrap();
    assert_eq!(local_events.lock().unwrap().len(), 1);
}

#[test]
fn panicking_native_observers_do_not_split_authority_and_projection() {
    let bytes = sample_xlsx();
    let mut left =
        Workbook::open_collaborative_recalculated(&bytes, 911, CalculationOptions::default())
            .unwrap();
    let mut right =
        Workbook::open_collaborative_recalculated(&bytes, 912, CalculationOptions::default())
            .unwrap();
    let local_calls = Arc::new(AtomicUsize::new(0));
    let remote_calls = Arc::new(AtomicUsize::new(0));

    let _local_panic = left
        .observe_update_v1(|_| panic!("local observer panic"))
        .unwrap();
    let observed = Arc::clone(&local_calls);
    let _local_after = left
        .observe_update_v1(move |_| {
            observed.fetch_add(1, Ordering::SeqCst);
        })
        .unwrap();
    let _remote_panic = right
        .observe_update_v1(|_| panic!("remote observer panic"))
        .unwrap();
    let observed = Arc::clone(&remote_calls);
    let _remote_after = right
        .observe_update_v1(move |_| {
            observed.fetch_add(1, Ordering::SeqCst);
        })
        .unwrap();

    for (address, input) in [("C4", "first"), ("C5", "second")] {
        left.edit_cell(
            SheetId(0),
            cell(address),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
        let update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap();
        assert_eq!(left.model(), right.model());
        assert_eq!(
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1()
        );
    }

    assert_eq!(local_calls.load(Ordering::SeqCst), 2);
    assert_eq!(remote_calls.load(Ordering::SeqCst), 2);
    assert_eq!(
        Workbook::open(&right.save().unwrap()).unwrap().model(),
        right.model()
    );
}

#[test]
fn collaborative_mode_rejects_all_structural_ops_before_mutation() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 1001).unwrap();
    let range = CellRange::new(cell("A1"), cell("A2"));
    let structural_ops = vec![
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::MergeCells {
            sheet: SheetId(0),
            range,
        },
        Op::UnmergeCells {
            sheet: SheetId(0),
            range,
        },
        Op::AddSheet {
            index: 1,
            name: "Added".into(),
        },
        Op::RemoveSheet { index: 1 },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".into(),
        },
        Op::RestoreSheet {
            sheet: SheetId(0),
            name: "Restored".into(),
            formulas: Vec::new(),
        },
    ];
    let model = workbook.model().clone();
    let state = workbook.encode_state_as_update_v1();
    for op in structural_ops {
        assert!(matches!(
            workbook.apply_ops(vec![op], CalculationOptions::default()),
            Err(Error::CollaborativeStructureOperation)
        ));
        assert_eq!(workbook.model(), &model);
        assert_eq!(workbook.encode_state_as_update_v1(), state);
        assert!(!workbook.can_undo());
    }

    assert!(
        workbook
            .apply_ops(
                vec![
                    Op::SetColWidth {
                        sheet: SheetId(0),
                        col: 0,
                        width: Some(22.0),
                    },
                    Op::SetRowHeight {
                        sheet: SheetId(0),
                        row: 0,
                        height: Some(24.0),
                    },
                ],
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
}

#[test]
fn collaboration_decoding_validates_malformed_and_oversized_payloads() {
    let bytes = sample_xlsx();
    let mut workbook = Workbook::open_collaborative(&bytes, 1101).unwrap();
    assert!(matches!(
        workbook.encode_diff_v1(&[0xff]),
        Err(Error::InvalidStateVector(_))
    ));
    assert!(matches!(
        workbook.encode_diff_v1(&[0, 0]),
        Err(Error::InvalidStateVector(_))
    ));
    assert_eq!(MAX_COLLABORATION_STATE_VECTOR_ENTRIES, 65_536);
    assert!(matches!(
        workbook.encode_diff_v1(&[0x81, 0x80, 0x04]),
        Err(Error::InvalidStateVector(_))
    ));
    let oversized = vec![0_u8; MAX_COLLABORATION_BYTES + 1];
    assert!(matches!(
        workbook.encode_diff_v1(&oversized),
        Err(Error::CollaborationDataTooLarge { .. })
    ));
    assert!(matches!(
        workbook.apply_update_v1(&oversized, CalculationOptions::default()),
        Err(Error::CollaborationDataTooLarge { .. })
    ));
    assert!(matches!(
        Workbook::open_collaborative(&bytes, MAX_COLLABORATION_CLIENT_ID + 1),
        Err(Error::InvalidClientId { .. })
    ));
    let max_client = Workbook::open_collaborative(&bytes, MAX_COLLABORATION_CLIENT_ID).unwrap();
    assert_eq!(max_client.client_id(), MAX_COLLABORATION_CLIENT_ID);
}

#[test]
fn save_preserves_unmodeled_package_parts_and_sheet_fragments() {
    let original = preservation_fixture();
    let before_order = ooxml_opc::unzip_parts(&original)
        .unwrap()
        .into_iter()
        .map(|(path, _)| path)
        .collect::<Vec<_>>();
    let before = package_map(&original);
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("B2"),
            "edited",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let after_order = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .map(|(path, _)| path)
        .collect::<Vec<_>>();
    let after = package_map(&saved);

    assert_eq!(
        after_order,
        before_order
            .into_iter()
            .filter(|path| path != "xl/calcChain.xml")
            .collect::<Vec<_>>()
    );
    for path in before.keys() {
        if path != "xl/calcChain.xml" {
            assert!(after.contains_key(path), "missing {path}");
        }
    }
    let owned = [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/sharedStrings.xml",
        "xl/styles.xml",
        "xl/theme/theme1.xml",
        "xl/worksheets/sheet1.xml",
    ];
    for (path, bytes) in &before {
        if path != "xl/calcChain.xml" && !owned.contains(&path.as_str()) {
            assert_eq!(&after[path], bytes, "changed {path}");
        }
    }

    let workbook_xml = String::from_utf8(after["xl/workbook.xml"].clone()).unwrap();
    assert!(workbook_xml.contains(r#"<definedName name="NamedCell">Data!$A$1</definedName>"#));
    assert!(!after.contains_key("xl/calcChain.xml"));
    assert!(workbook_xml.contains(r#"fullCalcOnLoad="1""#));
    assert_eq!(
        after["xl/sharedStrings.xml"],
        before["xl/sharedStrings.xml"]
    );
    let worksheet = String::from_utf8(after["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let fragments = [
        "<sheetViews>",
        "<autoFilter",
        "<conditionalFormatting",
        "<dataValidations",
        "<hyperlinks>",
        "<pageSetup",
        "<drawing",
        "<legacyDrawing",
        "<tableParts",
    ];
    let positions = fragments
        .iter()
        .map(|fragment| worksheet.find(fragment).unwrap())
        .collect::<Vec<_>>();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    assert!(worksheet.contains(r#"state="frozen""#));

    let content_types = String::from_utf8(after["[Content_Types].xml"].clone()).unwrap();
    for part in [
        "/xl/drawings/drawing1.xml",
        "/xl/tables/table1.xml",
        "/xl/comments1.xml",
        "/xl/externalLinks/externalLink1.xml",
        "/docProps/core.xml",
    ] {
        assert!(
            content_types.contains(part),
            "missing content type for {part}"
        );
    }
    assert!(!content_types.contains("/xl/calcChain.xml"));
    let workbook_rels = String::from_utf8(after["xl/_rels/workbook.xml.rels"].clone()).unwrap();
    assert!(!workbook_rels.contains(r#"Id="rId9""#));
    assert!(workbook_rels.contains(r#"Id="rId12""#));
    let styles = String::from_utf8(after["xl/styles.xml"].clone()).unwrap();
    assert!(styles.contains("<dxfs"));
    assert!(styles.contains("<tableStyles"));

    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("A1"))
            .unwrap()
            .value,
        CellValue::Text {
            value: "original".to_owned()
        }
    );
    assert_eq!(
        reopened
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(cell("B2"))
            .unwrap()
            .value,
        CellValue::Text {
            value: "edited".to_owned()
        }
    );
}

#[test]
fn preserved_package_save_reaches_a_part_fixed_point() {
    let original = preservation_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("B2"),
            "fixed",
            CalculationOptions::default(),
        )
        .unwrap();
    let first = workbook.save().unwrap();
    let second = Workbook::open(&first).unwrap().save().unwrap();
    assert_eq!(
        ooxml_opc::unzip_parts(&first).unwrap(),
        ooxml_opc::unzip_parts(&second).unwrap()
    );
}

#[test]
fn collaborative_materialization_retains_source_package() {
    let original = preservation_fixture();
    let before = package_map(&original);
    let mut left = Workbook::open_collaborative(&original, 1201).unwrap();
    let mut right = Workbook::open_collaborative(&original, 1202).unwrap();
    left.edit_cell(
        SheetId(0),
        cell("B2"),
        "remote",
        CalculationOptions::default(),
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    let after = package_map(&right.save().unwrap());

    for path in [
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/drawings/drawing1.xml",
        "xl/tables/table1.xml",
        "xl/comments1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/externalLinks/externalLink1.xml",
        "docProps/core.xml",
        "customXml/item1.xml",
    ] {
        assert_eq!(after[path], before[path], "changed {path}");
    }
    assert!(!after.contains_key("xl/calcChain.xml"));
}

const CHART_DRAWING: &[u8] = br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

const CHART_PART: &[u8] = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:pt idx="0"><c:v>Series</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:dLbls><c:f>Data!$C$2:$C$4</c:f></c:dLbls></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

/// The preservation fixture already anchors a drawing on `Data`; give that
/// drawing a chart so the whole sheet -> drawing -> chart chain is real.
fn charted_fixture() -> Vec<u8> {
    let mut parts = preservation_fixture_parts();
    set_test_part(
        &mut parts,
        "xl/drawings/drawing1.xml",
        CHART_DRAWING.to_vec(),
    );
    parts.extend([
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/charts/chart1.xml".to_owned(), CHART_PART.to_vec()),
    ]);
    let content_types = test_part_text(&parts, "[Content_Types].xml").replace(
        "</Types>",
        r#"<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#,
    );
    set_test_part(
        &mut parts,
        "[Content_Types].xml",
        content_types.into_bytes(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The charted fixture with a second sheet the chart plots, so removing that
/// sheet strands a reference a chart on another sheet owns.
fn cross_sheet_charted_fixture() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let workbook = test_part_text(&parts, "xl/workbook.xml").replace(
        "</sheets>",
        r#"<sheet name="Source" sheetId="8" r:id="rIdSource"/></sheets>"#,
    );
    set_test_part(&mut parts, "xl/workbook.xml", workbook.into_bytes());
    let rels = test_part_text(&parts, "xl/_rels/workbook.xml.rels").replace(
        "</Relationships>",
        r#"<Relationship Id="rIdSource" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
    );
    set_test_part(&mut parts, "xl/_rels/workbook.xml.rels", rels.into_bytes());
    parts.push((
        "xl/worksheets/sheet2.xml".to_owned(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
    ));
    let chart = String::from_utf8(CHART_PART.to_vec())
        .unwrap()
        .replace("Data!$B$2:$B$4", "Source!$A$1:$A$2");
    set_test_part(&mut parts, "xl/charts/chart1.xml", chart.into_bytes());
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn chart_formulas(workbook: &Workbook) -> Vec<String> {
    workbook.model().sheets[0].charts[0]
        .refs
        .iter()
        .map(|reference| reference.formula.clone())
        .collect()
}

/// A row insert above the plotted range moves every chart reference with the
/// cells, moves the anchor per its `editAs` mode, and writes both back into
/// their parts without disturbing the cached values around them.
#[test]
fn chart_references_and_anchor_follow_a_row_insert() {
    let original = charted_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    assert_eq!(
        chart_formulas(&workbook),
        [
            "Data!$B$1",
            "Data!$A$2:$A$4",
            "Data!$B$2:$B$4",
            "Data!$C$2:$C$4"
        ]
    );

    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        chart_formulas(&workbook),
        [
            "Data!$B$1",
            "Data!$A$4:$A$6",
            "Data!$B$4:$B$6",
            "Data!$C$4:$C$6"
        ]
    );

    let saved = workbook.save().unwrap();
    let parts = package_map(&saved);
    let patched = String::from_utf8(parts["xl/charts/chart1.xml"].clone()).unwrap();
    let source = String::from_utf8(CHART_PART.to_vec()).unwrap();
    assert_eq!(
        patched,
        source
            .replace("Data!$A$2:$A$4", "Data!$A$4:$A$6")
            .replace("Data!$B$2:$B$4", "Data!$B$4:$B$6")
            .replace("Data!$C$2:$C$4", "Data!$C$4:$C$6")
            .replace(
                r#"<c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>"#,
                r#"<c:strCache><c:ptCount val="3"/></c:strCache></c:strRef></c:cat>"#,
            )
            .replace(
                r#"<c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache>"#,
                r#"<c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache>"#,
            ),
        "the moved references carry their caches, and nothing else changes"
    );

    let drawing = String::from_utf8(parts["xl/drawings/drawing1.xml"].clone()).unwrap();
    let expected_drawing = String::from_utf8(CHART_DRAWING.to_vec())
        .unwrap()
        .replace("<xdr:row>4</xdr:row>", "<xdr:row>6</xdr:row>")
        .replace("<xdr:row>19</xdr:row>", "<xdr:row>21</xdr:row>");
    assert_eq!(drawing, expected_drawing, "oneCell moves without resizing");

    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        chart_formulas(&reopened),
        [
            "Data!$B$1",
            "Data!$A$4:$A$6",
            "Data!$B$4:$B$6",
            "Data!$C$4:$C$6"
        ]
    );
    assert_eq!(reopened.model().sheets[0].charts[0].anchor_index, 0);
}

/// `editAs` decides what a grid edit does to a chart: `twoCell` moves and
/// resizes, `oneCell` moves without resizing, `absolute` does neither. An
/// insert inside the anchored span is what tells the first two apart.
#[test]
fn chart_anchors_honour_their_edit_as_mode() {
    // (editAs, insert row, expected from row, expected to row)
    for (mode, at, from_row, to_row) in [
        ("twoCell", 1, 6, 21),
        ("oneCell", 1, 6, 21),
        ("absolute", 1, 4, 19),
        ("twoCell", 10, 4, 21),
        ("oneCell", 10, 4, 19),
        ("absolute", 10, 4, 19),
    ] {
        let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
        let drawing = String::from_utf8(CHART_DRAWING.to_vec())
            .unwrap()
            .replace(r#"editAs="oneCell""#, &format!(r#"editAs="{mode}""#));
        set_test_part(&mut parts, "xl/drawings/drawing1.xml", drawing.into_bytes());
        let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
        workbook
            .apply_ops(
                vec![Op::InsertRows {
                    sheet: SheetId(0),
                    at,
                    count: 2,
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        let saved = package_map(&workbook.save().unwrap());
        let patched = String::from_utf8(saved["xl/drawings/drawing1.xml"].clone()).unwrap();
        assert!(
            patched.contains(&format!(
                "<xdr:row>{from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            )),
            "{mode} at row {at} put its top corner on the wrong row: {patched}"
        );
        assert!(
            patched.contains(&format!(
                "<xdr:row>{to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
            )),
            "{mode} at row {at} put its bottom corner on the wrong row: {patched}"
        );
    }
}

/// A cache beside a reference this crate cannot resolve keeps its pre-edit
/// values through a structural edit, because the reference itself never
/// changes. The edit is refused rather than saved as a chart whose cache and
/// reference disagree.
#[test]
fn refuses_a_structural_edit_beside_a_cache_that_cannot_be_rebuilt() {
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let chart = String::from_utf8(CHART_PART.to_vec())
        .unwrap()
        .replace("Data!$B$2:$B$4", "SalesRange");
    set_test_part(&mut parts, "xl/charts/chart1.xml", chart.into_bytes());
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let before = workbook.model().clone();

    let error = workbook
        .apply_ops(
            vec![Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::InvalidOperation(message)
            if message.contains("xl/charts/chart1.xml")),
        "{error:?}"
    );
    assert_eq!(workbook.model(), &before);
}

/// Renaming the plotted sheet rewrites the qualifier in every chart reference,
/// and undo puts the old name back.
#[test]
fn sheet_rename_rewrites_chart_references() {
    let mut workbook = Workbook::open(&charted_fixture()).unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Sales Data".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        chart_formulas(&workbook),
        [
            "'Sales Data'!$B$1",
            "'Sales Data'!$A$2:$A$4",
            "'Sales Data'!$B$2:$B$4",
            "'Sales Data'!$C$2:$C$4"
        ]
    );
    let saved = package_map(&workbook.save().unwrap());
    let patched = String::from_utf8(saved["xl/charts/chart1.xml"].clone()).unwrap();
    assert!(
        patched.contains("<c:f>'Sales Data'!$B$1</c:f>"),
        "renamed qualifier missing: {patched}"
    );
    assert!(!patched.contains("Data!$B$1"), "old qualifier left behind");

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(
        chart_formulas(&workbook),
        [
            "Data!$B$1",
            "Data!$A$2:$A$4",
            "Data!$B$2:$B$4",
            "Data!$C$2:$C$4"
        ]
    );
}

/// Deleting every plotted row collapses the reference the way a cell formula
/// would, rather than leaving it on addresses that no longer exist.
#[test]
fn deleted_rows_collapse_chart_references_to_ref_errors() {
    let mut workbook = Workbook::open(&charted_fixture()).unwrap();
    workbook
        .apply_ops(
            vec![Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 3,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        chart_formulas(&workbook),
        ["Data!$B$1", "#REF!", "#REF!", "#REF!"]
    );
    Workbook::open(&workbook.save().unwrap()).unwrap();
}

#[test]
fn refuses_structural_ops_that_would_strand_pivot_references() {
    let mut parts = ooxml_opc::unzip_parts(&preservation_fixture()).unwrap();
    parts.push((
        "xl/pivotcache/pivotCacheDefinition1.xml".to_owned(),
        br#"<pivotCacheDefinition><cacheSource><worksheetSource sheet="Data" ref="A1:B2"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    ));
    let original = ooxml_opc::rezip_parts(&parts).unwrap();

    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::RemoveSheet { index: 0 },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".to_owned(),
        },
    ] {
        let mut workbook = Workbook::open(&original).unwrap();
        let error = workbook
            .apply_ops(vec![op.clone()], CalculationOptions::default())
            .unwrap_err();
        assert!(
            matches!(&error, Error::InvalidOperation(message)
                if message.contains("pivotCacheDefinition1.xml")),
            "{op:?} was allowed: {error:?}"
        );
    }

    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "edited",
            CalculationOptions::default(),
        )
        .unwrap();
    Workbook::open(&workbook.save().unwrap()).unwrap();
}

#[test]
fn refuses_structural_ops_that_would_strand_unclaimed_chart_parts() {
    let mut parts = ooxml_opc::unzip_parts(&preservation_fixture()).unwrap();
    parts.push((
        "xl/charts/chart1.xml".to_owned(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>"#.to_vec(),
    ));
    let original = ooxml_opc::rezip_parts(&parts).unwrap();

    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::RemoveSheet { index: 0 },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".to_owned(),
        },
    ] {
        let mut workbook = Workbook::open(&original).unwrap();
        let error = workbook
            .apply_ops(vec![op.clone()], CalculationOptions::default())
            .unwrap_err();
        assert!(
            matches!(&error, Error::InvalidOperation(message) if message.contains("chart1.xml")),
            "{op:?} was allowed: {error:?}"
        );
    }

    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "edited",
            CalculationOptions::default(),
        )
        .unwrap();
    Workbook::open(&workbook.save().unwrap()).unwrap();
}

#[test]
fn non_worksheet_sheets_stay_typed_and_byte_identical() {
    let original = non_worksheet_fixture();
    let before = package_map(&original);
    let mut workbook = Workbook::open(&original).unwrap();
    assert_eq!(workbook.sheet_count(), 3);
    assert!(workbook.model().sheets[1].used_range().is_none());
    assert!(matches!(
        workbook.edit_cell(
            SheetId(1),
            cell("A1"),
            "blocked",
            CalculationOptions::default()
        ),
        Err(Error::InvalidOperation(_))
    ));
    workbook
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "edited",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let after = package_map(&saved);
    assert_eq!(
        after["xl/chartsheets/sheet1.xml"],
        before["xl/chartsheets/sheet1.xml"]
    );
    assert_eq!(
        after["xl/dialogsheets/sheet1.xml"],
        before["xl/dialogsheets/sheet1.xml"]
    );
    let relationships = String::from_utf8(after["xl/_rels/workbook.xml.rels"].clone()).unwrap();
    assert!(relationships.contains("/chartsheet\""));
    assert!(relationships.contains("/dialogsheet\""));
    assert_eq!(relationships.matches("/worksheet\"").count(), 1);
    let content_types = String::from_utf8(after["[Content_Types].xml"].clone()).unwrap();
    assert!(content_types.contains("spreadsheetml.chartsheet+xml"));
    assert!(content_types.contains("spreadsheetml.dialogsheet+xml"));
    assert!(
        !String::from_utf8(after["xl/chartsheets/sheet1.xml"].clone())
            .unwrap()
            .contains("sheetData")
    );
    Workbook::open(&saved).unwrap();
}

#[test]
fn strict_prefixed_templates_keep_namespaces_relationships_and_mc_order() {
    let original = strict_prefixed_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .edit_cell(SheetId(0), cell("A1"), "2", CalculationOptions::default())
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Added".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let parts = package_map(&saved);
    let workbook_xml = String::from_utf8(parts["xl/workbook.xml"].clone()).unwrap();
    let strict_main = r#"xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main""#;
    let strict_rel = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    assert!(workbook_xml.contains(&format!(r#"<sheets xmlns:r="{strict_rel}" {strict_main}"#)));
    assert!(workbook_xml.contains(r#"<sheet name="Data" sheetId="1" rel:id="rId1"/>"#));
    assert!(workbook_xml.contains(r#"r:id="rId2""#));
    assert!(workbook_xml.contains("<calcPr"));
    assert!(workbook_xml.contains("<s:definedName name=\"StrictName\">Data!$A$1</s:definedName>"));
    let worksheet = String::from_utf8(parts["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(worksheet.contains(r#"<x:sheetData marker="keep"/>"#));
    assert!(
        worksheet.find("<mc:AlternateContent").unwrap()
            < worksheet
                .find(&format!("<sheetData {strict_main}"))
                .unwrap()
    );
    assert!(worksheet.contains("<row r=\"1\""));
    assert!(worksheet.contains("<c r=\"A1\""));
    assert!(!worksheet.contains("<s:sheetData"));
    let relationships = String::from_utf8(parts["xl/_rels/workbook.xml.rels"].clone()).unwrap();
    assert_eq!(
        relationships
            .matches("http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet")
            .count(),
        2
    );
    assert!(!relationships.contains("schemas.openxmlformats.org/officeDocument"));
    let added = String::from_utf8(parts["xl/worksheets/sheet2.xml"].clone()).unwrap();
    assert!(added.contains("xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\""));
    let content_types = String::from_utf8(parts["[Content_Types].xml"].clone()).unwrap();
    assert!(content_types.contains(r#"PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.ms-excel.worksheet+xml""#));
    assert_eq!(Workbook::open(&saved).unwrap().sheet_count(), 2);
}

#[test]
fn no_edit_round_trip_keeps_calculation_chain_and_source_parts() {
    let original = preservation_fixture();
    let before = ooxml_opc::unzip_parts(&original).unwrap();
    let saved = Workbook::open(&original).unwrap().save().unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(after, before);
    assert!(package_map(&saved).contains_key("xl/calcChain.xml"));
}

#[test]
fn defined_names_follow_renames_and_drop_ambiguous_references() {
    let original = defined_names_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed Sheet".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = package_map(&workbook.save().unwrap());
    let workbook_xml = String::from_utf8(saved["xl/workbook.xml"].clone()).unwrap();
    assert!(workbook_xml.contains("&apos;Renamed Sheet&apos;!$A$1"));
    assert!(!workbook_xml.contains(r#"name="AmbiguousData""#));
    assert!(workbook_xml.contains(r#"name="Unrelated">42</definedName>"#));
}

#[test]
fn renaming_a_function_named_sheet_keeps_its_defined_names() {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("SUM"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    set_test_part(
        &mut parts,
        "xl/workbook.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SUM" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Qualified">SUM(SUM!$A$1:$A$10)</definedName><definedName name="Unqualified">SUM($A$1:$A$10)</definedName></definedNames></workbook>"#.to_vec(),
    );
    let original = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let workbook_xml = String::from_utf8(package_map(&saved)["xl/workbook.xml"].clone()).unwrap();
    assert!(workbook_xml.contains(r#"name="Qualified">SUM(Renamed!$A$1:$A$10)</definedName>"#));
    assert!(workbook_xml.contains(r#"name="Unqualified">SUM($A$1:$A$10)</definedName>"#));
    Workbook::open(&saved).unwrap();
}

#[test]
fn scoped_defined_names_remap_indices_and_drop_deleted_scopes() {
    let original = defined_names_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "Fresh".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let parts = package_map(&saved);
    let workbook_xml = String::from_utf8(parts["xl/workbook.xml"].clone()).unwrap();
    assert!(!workbook_xml.contains(r#"name="LocalMiddle""#));
    assert!(workbook_xml.contains(r#"name="LocalData" localSheetId="1""#));
    assert!(workbook_xml.contains(r#"name="LocalTail" localSheetId="2""#));
    Workbook::open(&saved).unwrap();
}

#[test]
fn undo_restores_defined_names_dropped_by_a_sheet_removal() {
    let original = defined_names_fixture();
    let mut workbook = Workbook::open(&original).unwrap();
    let before = workbook.model().defined_names.clone();
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        !workbook
            .model()
            .defined_names
            .iter()
            .any(|defined| defined.name == "LocalMiddle")
    );
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, before);
}

/// v1 leaves preserved sheet fragments at their original geometry after an
/// axis edit; the file must still open, even though the ranges have drifted.
#[test]
fn row_insertion_preserves_unmodeled_ranges_and_anchors_without_corruption() {
    let original = preservation_fixture();
    let before = package_map(&original);
    let mut workbook = Workbook::open(&original).unwrap();
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let after = package_map(&saved);
    let worksheet = String::from_utf8(after["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(worksheet.contains(r#"<autoFilter ref="A1:B2""#));
    assert!(worksheet.contains(r#"<dataValidation type="whole" sqref="B2""#));
    assert!(worksheet.contains(r#"<conditionalFormatting sqref="B2""#));
    assert_eq!(
        after["xl/drawings/drawing1.xml"],
        before["xl/drawings/drawing1.xml"]
    );
    Workbook::open(&saved).unwrap();
}

#[test]
fn remove_then_add_is_fresh_while_undo_restores_exact_sheet_identity() {
    let original = preservation_fixture();
    let mut replaced = Workbook::open(&original).unwrap();
    replaced
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Keep".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    replaced
        .apply_ops(
            vec![Op::RemoveSheet { index: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    replaced
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "Data".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let replaced_parts = package_map(&replaced.save().unwrap());
    assert!(!replaced_parts.contains_key("xl/worksheets/sheet1.xml"));
    for (path, bytes) in &replaced_parts {
        if path.starts_with("xl/worksheets/") && path.ends_with(".xml") {
            assert!(
                !String::from_utf8(bytes.clone())
                    .unwrap()
                    .contains("<autoFilter")
            );
        }
    }

    let mut restored = Workbook::open(&original).unwrap();
    restored
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Keep".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    restored
        .apply_ops(
            vec![Op::RemoveSheet { index: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    restored.undo(CalculationOptions::default()).unwrap();
    let restored_parts = package_map(&restored.save().unwrap());
    assert!(
        String::from_utf8(restored_parts["xl/worksheets/sheet1.xml"].clone())
            .unwrap()
            .contains("<autoFilter")
    );
}

/// Provenance is recorded against the address a cell was read from, so a row
/// or column edit that moves the cell has to move it too.
#[test]
fn shared_string_provenance_follows_cells_through_row_and_column_edits() {
    let mut workbook = Workbook::open(&ambiguous_shared_string_fixture()).unwrap();
    workbook
        .apply_ops(
            vec![
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 2,
                },
                Op::InsertCols {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();

    let inserted = saved_sheet_text(&workbook);
    assert!(
        inserted.contains(r#"<c r="C4" t="s"><v>0</v></c>"#),
        "{inserted}"
    );
    assert!(
        inserted.contains(r#"<c r="E4" t="s"><v>1</v></c>"#),
        "the bold entry collapsed onto the plain one: {inserted}"
    );

    workbook
        .apply_ops(
            vec![
                Op::DeleteRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 2,
                },
                Op::DeleteCols {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();

    let deleted = saved_sheet_text(&workbook);
    assert!(
        deleted.contains(r#"<c r="B2" t="s"><v>0</v></c>"#),
        "{deleted}"
    );
    assert!(
        deleted.contains(r#"<c r="D2" t="s"><v>1</v></c>"#),
        "the bold entry collapsed onto the plain one: {deleted}"
    );
}

/// Deleting the column a cell sits in drops its provenance with the cell; the
/// surviving cell keeps its own.
#[test]
fn deleting_a_column_leaves_the_surviving_cell_on_its_own_entry() {
    let mut workbook = Workbook::open(&ambiguous_shared_string_fixture()).unwrap();
    workbook
        .apply_ops(
            vec![Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();

    let saved = saved_sheet_text(&workbook);
    assert!(
        saved.contains(r#"<c r="C2" t="s"><v>1</v></c>"#),
        "the bold entry was lost when the plain one was deleted: {saved}"
    );
    assert!(!saved.contains(r#"<v>0</v>"#), "{saved}");
}

#[test]
fn undo_and_redo_restore_shared_string_provenance() {
    let mut workbook = Workbook::open(&ambiguous_shared_string_fixture()).unwrap();
    workbook
        .apply_ops(
            vec![Op::DeleteRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let deleted = saved_sheet_text(&workbook);
    assert!(
        deleted.contains(r#"<c r="D1" t="s"><v>1</v></c>"#),
        "{deleted}"
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    let undone = saved_sheet_text(&workbook);
    assert!(
        undone.contains(r#"<c r="B2" t="s"><v>0</v></c>"#),
        "{undone}"
    );
    assert!(
        undone.contains(r#"<c r="D2" t="s"><v>1</v></c>"#),
        "{undone}"
    );

    workbook.redo(CalculationOptions::default()).unwrap();
    let redone = saved_sheet_text(&workbook);
    assert!(
        redone.contains(r#"<c r="D1" t="s"><v>1</v></c>"#),
        "{redone}"
    );
}

fn charted_model(chart: SheetChart) -> WorkbookModel {
    let mut sheet = Sheet::new("Data");
    sheet.charts.push(chart);
    let mut model = WorkbookModel::default();
    model.sheets.push(sheet);
    model
}

fn sample_chart() -> SheetChart {
    SheetChart {
        part: "xl/charts/chart1.xml".to_owned(),
        drawing: "xl/drawings/drawing1.xml".to_owned(),
        anchor_index: 0,
        anchor: ChartAnchor::TwoCell {
            from: AnchorCell::default(),
            to: AnchorCell {
                col: 4,
                row: 8,
                ..AnchorCell::default()
            },
            edit_as: AnchorEditAs::TwoCell,
        },
        refs: vec![ChartRef {
            kind: ChartRefKind::Values,
            formula: "Data!$A$1:$A$2".to_owned(),
        }],
    }
}

/// Chart state a peer controls reaches the writer, so an off-grid anchor, a
/// character xml cannot carry, a smuggled part path and two charts claiming one
/// anchor are all refused before anything is written.
#[test]
fn refuses_chart_state_the_writer_could_not_express() {
    let off_grid = {
        let mut chart = sample_chart();
        chart.anchor = ChartAnchor::TwoCell {
            from: AnchorCell::default(),
            to: AnchorCell {
                col: 4,
                row: MAX_ROWS,
                ..AnchorCell::default()
            },
            edit_as: AnchorEditAs::TwoCell,
        };
        (chart, "off the sheet grid")
    };
    let illegal_character = {
        let mut chart = sample_chart();
        chart.refs[0].formula = "Data!$A$1\u{7}".to_owned();
        (chart, "xml cannot carry")
    };
    let traversal = {
        let mut chart = sample_chart();
        chart.part = "../../etc/passwd".to_owned();
        (chart, "not a package part path")
    };
    let absolute = {
        let mut chart = sample_chart();
        chart.drawing = "/xl/drawings/drawing1.xml".to_owned();
        (chart, "not a package part path")
    };
    for (chart, reason) in [off_grid, illegal_character, traversal, absolute] {
        let Err(error) = Workbook::from_model(charted_model(chart)) else {
            panic!("{reason} must be refused");
        };
        assert!(
            matches!(&error, Error::InvalidOperation(message) if message.contains(reason)),
            "{reason}: {error:?}"
        );
    }

    let mut model = charted_model(sample_chart());
    model.sheets[0].charts.push(sample_chart());
    let Err(error) = Workbook::from_model(model) else {
        panic!("two charts on one anchor must be refused");
    };
    assert!(
        matches!(&error, Error::InvalidOperation(message)
            if message.contains("same part, drawing and anchor")),
        "{error:?}"
    );
}

/// Chart parts come out of the package a workbook was opened with, and this
/// crate creates none. A chart-bearing model with no source would save as a
/// workbook that lost every chart, so every door into one is closed.
#[test]
fn refuses_chart_state_with_no_source_package_to_preserve_it() {
    for build in [
        Workbook::from_model as fn(WorkbookModel) -> Result<Workbook, Error>,
        |model| Workbook::from_model_collaborative(model, 7),
    ] {
        let Err(error) = build(charted_model(sample_chart())) else {
            panic!("a chart with no source package must be refused");
        };
        assert!(
            matches!(&error, Error::InvalidOperation(message)
                if message.contains("only be preserved from a source package")),
            "{error:?}"
        );
    }

    let error = xlsx_parse::serialize_workbook(&charted_model(sample_chart())).unwrap_err();
    assert!(format!("{error}").contains("written back into the package it was read from"));

    let mut workbook = Workbook::open(&charted_fixture()).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "edited",
            CalculationOptions::default(),
        )
        .unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.model().sheets[0].charts.len(), 1);
}

/// An insertion that would push a marker a chart must move past the last row
/// is refused. Clamping it would shrink an object whose `editAs` forbids
/// resizing, or collapse it outright.
#[test]
fn refuses_an_insertion_that_would_push_a_chart_anchor_off_the_grid() {
    for (mode, edit_as) in [
        ("twoCell", AnchorEditAs::TwoCell),
        ("oneCell", AnchorEditAs::OneCell),
    ] {
        let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
        let drawing = String::from_utf8(CHART_DRAWING.to_vec())
            .unwrap()
            .replace(r#"editAs="oneCell""#, &format!(r#"editAs="{mode}""#))
            .replace(
                "<xdr:row>4</xdr:row>",
                &format!("<xdr:row>{}</xdr:row>", MAX_ROWS - 6),
            )
            .replace(
                "<xdr:row>19</xdr:row>",
                &format!("<xdr:row>{}</xdr:row>", MAX_ROWS - 2),
            );
        set_test_part(&mut parts, "xl/drawings/drawing1.xml", drawing.into_bytes());
        let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
        let anchored = ChartAnchor::TwoCell {
            from: AnchorCell {
                col: 2,
                col_off: 12_700,
                row: MAX_ROWS - 6,
                row_off: 0,
            },
            to: AnchorCell {
                col: 8,
                col_off: 0,
                row: MAX_ROWS - 2,
                row_off: 0,
            },
            edit_as,
        };
        assert_eq!(workbook.model().sheets[0].charts[0].anchor, anchored);

        let error = workbook
            .apply_ops(
                vec![Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 4,
                }],
                CalculationOptions::default(),
            )
            .unwrap_err();
        assert!(
            matches!(&error, Error::InvalidOperation(message)
                if message.contains("push a chart anchor past the sheet boundary")),
            "{mode}: {error:?}"
        );
        assert_eq!(
            workbook.model().sheets[0].charts[0].anchor,
            anchored,
            "a refused insertion must leave the anchor alone"
        );

        workbook
            .apply_ops(
                vec![Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                }],
                CalculationOptions::default(),
            )
            .expect("an insertion every marker survives is accepted");
    }
}

/// Removing a charted sheet strands the references that named it, so the
/// remaining sheets' chart state must reach the shared document. Undo must put
/// it back, and neither direction may leave the model ahead of the authority.
#[test]
fn removing_a_charted_sheet_synchronises_and_undoes_cleanly() {
    let mut workbook = Workbook::open(&cross_sheet_charted_fixture()).unwrap();
    let plotted =
        |workbook: &Workbook| workbook.model().sheets[0].charts[0].refs[2].formula.clone();
    assert_eq!(plotted(&workbook), "Source!$A$1:$A$2");

    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(plotted(&workbook), "#REF!");
    assert_eq!(workbook.model().sheets.len(), 1);

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(plotted(&workbook), "Source!$A$1:$A$2");
    assert_eq!(workbook.model().sheets.len(), 2);

    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(plotted(&workbook), "#REF!");
    assert_eq!(workbook.model().sheets.len(), 1);

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().sheets[0].charts[0].refs[2].formula,
        "#REF!"
    );
}

/// `SetCharts` is the inverse a remap emits; it is not something a caller may
/// submit, because it replaces state the engine derives.
#[test]
fn set_charts_is_rejected_as_an_internal_operation() {
    let mut workbook = Workbook::open(&charted_fixture()).unwrap();
    let Err(error) = workbook.apply_ops(
        vec![Op::SetCharts {
            sheet: SheetId(0),
            charts: Vec::new(),
        }],
        CalculationOptions::default(),
    ) else {
        panic!("SetCharts must be refused");
    };
    assert!(
        matches!(&error, Error::InvalidOperation(message) if message.contains("internal")),
        "{error:?}"
    );
}

/// Every op that rewrites `defined_names` is structural, and structural ops are
/// refused while collaborative. Peers therefore cannot disagree about a name.
#[test]
fn collaborative_sessions_refuse_every_op_that_rewrites_defined_names() {
    let bytes = defined_names_fixture();
    let rewriting_ops = vec![
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".to_owned(),
        },
        Op::SetDefinedNames {
            defined_names: Vec::new(),
        },
    ];

    for op in rewriting_ops {
        let mut left = Workbook::open_collaborative(&bytes, 101).unwrap();
        let error = left
            .apply_ops(vec![op.clone()], CalculationOptions::default())
            .unwrap_err();
        assert!(
            matches!(error, Error::CollaborativeStructureOperation),
            "{op:?} must be refused while collaborative, or peers diverge on defined names"
        );
    }

    let mut left = Workbook::open_collaborative(&bytes, 101).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 202).unwrap();
    left.edit_cell(SheetId(0), cell("A1"), "21", CalculationOptions::default())
        .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model().defined_names, right.model().defined_names);
}
