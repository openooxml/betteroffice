#[cfg(feature = "raster")]
use betteroffice_xlsx::RenderOptions;
use betteroffice_xlsx::{
    AnchorCell, AnchorEditAs, AnchorExtent, CalculationOptions, Cell, CellInput, CellRange,
    CellRef, CellState, CellValue, ChartAnchor, ChartRef, ChartRefKind, DefinedName, DrawCmd,
    Error, FreezePane, GridGeometry, Hyperlink, MAX_COLLABORATION_BYTES,
    MAX_COLLABORATION_CLIENT_ID, MAX_COLLABORATION_STATE_VECTOR_ENTRIES, MAX_ROWS,
    NumberFormatKind, NumberFormatMutation, Op, ProposalEditInput, ProposalRequest, Sheet,
    SheetChart, SheetId, StylePatch, UpdateOrigin, Viewport, Workbook, WorkbookModel,
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

/// A workbook whose names use the dynamic-range idiom: a whole-column
/// reference inside a call, and the range operator applied to `INDEX`. Neither
/// is anything the formula parser reads, and every row and column edit on the
/// sheet they name used to be refused because of it.
fn dynamic_range_names_fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    set_test_part(
        &mut parts,
        "xl/workbook.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Rows">Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))</definedName><definedName name="Header">SUM(Data!$1:$1)</definedName><definedName name="Stranded">Data!#REF!</definedName></definedNames></workbook>"#.to_vec(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn structural_edits_rewrite_dynamic_range_names_through_save_and_undo() {
    let mut workbook = Workbook::open(&dynamic_range_names_fixture()).unwrap();
    let before = workbook.model().defined_names.clone();

    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 9998,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(workbook.model().defined_names, before);

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

    let formula = |name: &str| {
        workbook
            .model()
            .defined_names
            .iter()
            .find(|defined| defined.name == name)
            .unwrap()
            .formula
            .clone()
    };
    assert_eq!(
        formula("Rows"),
        "Data!$A$2:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))"
    );
    assert_eq!(formula("Header"), "SUM(Data!$2:$2)");
    assert_eq!(formula("Stranded"), "Data!#REF!");

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().defined_names,
        workbook.model().defined_names
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().defined_names, before);
}

/// A workbook whose names use the reference operators the formula lexer has no
/// token for: the spill `#`, the implicit intersection `@`, and the range
/// operator carrying the whitespace and the second qualifier Excel allows.
fn reference_operator_names_fixture(defined_names: &str) -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Sheet2"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    let workbook_xml = test_part_text(&parts, "xl/workbook.xml").replace(
        "</workbook>",
        &format!("<definedNames>{defined_names}</definedNames></workbook>"),
    );
    set_test_part(&mut parts, "xl/workbook.xml", workbook_xml.into_bytes());
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_workbook_xml(workbook: &mut Workbook) -> String {
    let saved = workbook.save().unwrap();
    String::from_utf8(package_map(&saved)["xl/workbook.xml"].clone()).unwrap()
}

#[test]
fn structural_edits_move_reference_operator_names_into_the_saved_workbook() {
    let mut workbook = Workbook::open(&reference_operator_names_fixture(
        r#"<definedName name="Spilled">SUM(Data!A1#)</definedName><definedName name="Picked">SUM(@Data!A1)</definedName><definedName name="Spaced">Data!A1: Data!B2</definedName>"#,
    ))
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

    let saved = saved_workbook_xml(&mut workbook);
    assert!(saved.contains(">SUM(Data!A2#)<"), "{saved}");
    assert!(saved.contains(">SUM(@Data!A2)<"), "{saved}");
    assert!(saved.contains(">Data!A2: Data!B3<"), "{saved}");
    assert_eq!(
        Workbook::open(&workbook.save().unwrap())
            .unwrap()
            .model()
            .defined_names,
        workbook.model().defined_names
    );
}

/// A range that loses its first row keeps its span; the endpoints used to be
/// clipped apart, stranding the near one on `#REF!` while the far one still
/// named a live cell.
#[test]
fn a_deletion_inside_a_spaced_range_name_clips_it_as_one_span() {
    let mut workbook = Workbook::open(&reference_operator_names_fixture(
        r#"<definedName name="Spaced">Data!A1: Data!B2</definedName>"#,
    ))
    .unwrap();

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

    let saved = saved_workbook_xml(&mut workbook);
    assert!(saved.contains(">Data!A1: Data!B1<"), "{saved}");
}

/// `A1:Sheet2!B2` is the range operator applied to a cell, not a sheet span,
/// and the unqualified half of a workbook name binds to whichever sheet is
/// active — so the edit is refused rather than applied over a name that never
/// moved.
#[test]
fn structural_edits_refuse_a_range_endpoint_in_front_of_a_sheet_qualifier() {
    let mut workbook = Workbook::open(&reference_operator_names_fixture(
        r#"<definedName name="Mixed">A1:Sheet2!B2</definedName>"#,
    ))
    .unwrap();
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

/// Two sheets whose drawings both name one chart part, which the package
/// format permits. The chart's references are unqualified, so each sheet
/// resolves them against itself.
fn shared_chart_fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("First"));
    model.sheets.push(Sheet::new("Second"));
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();

    for index in 1..=2 {
        let path = format!("xl/worksheets/sheet{index}.xml");
        let worksheet = test_part_text(&parts, &path).replace(
            "</worksheet>",
            r#"<drawing r:id="rIdDrawing"/></worksheet>"#,
        );
        set_test_part(&mut parts, &path, worksheet.into_bytes());
        parts.push((
            format!("xl/worksheets/_rels/sheet{index}.xml.rels"),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{index}.xml"/></Relationships>"#
            )
            .into_bytes(),
        ));
        parts.push((
            format!("xl/drawings/drawing{index}.xml"),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
        ));
        parts.push((
            format!("xl/drawings/_rels/drawing{index}.xml.rels"),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#.to_vec(),
        ));
    }
    parts.push((
        "xl/charts/chart1.xml".to_owned(),
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>$A$1:$A$2</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.to_vec(),
    ));
    let content_types = test_part_text(&parts, "[Content_Types].xml").replace(
        "</Types>",
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#,
    );
    set_test_part(
        &mut parts,
        "[Content_Types].xml",
        content_types.into_bytes(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn shared_chart_formula(workbook: &Workbook, sheet: SheetId) -> String {
    workbook.model().sheet(sheet).unwrap().charts[0].refs[0]
        .formula
        .clone()
}

/// Inserting a row on one of two sheets that share a chart part moves only that
/// sheet's references, so the part can no longer carry both. The save is
/// refused rather than rewriting the chart the other sheet shows.
#[test]
fn refuses_to_save_a_shared_chart_part_a_structural_edit_split() {
    let mut workbook = Workbook::open(&shared_chart_fixture()).unwrap();
    assert_eq!(shared_chart_formula(&workbook, SheetId(0)), "$A$1:$A$2");
    assert_eq!(shared_chart_formula(&workbook, SheetId(1)), "$A$1:$A$2");
    let untouched = workbook.save().unwrap();

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
    assert_eq!(shared_chart_formula(&workbook, SheetId(0)), "$A$2:$A$3");
    assert_eq!(shared_chart_formula(&workbook, SheetId(1)), "$A$1:$A$2");

    let error = workbook.save().unwrap_err();
    assert!(
        matches!(&error, Error::Spreadsheet(xlsx_parse::ParseError::UnsupportedEdit(message))
            if message.contains("xl/charts/chart1.xml")
                && message.contains("sheet First")
                && message.contains("sheet Second")),
        "{error:?}"
    );

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(shared_chart_formula(&workbook, SheetId(1)), "$A$1:$A$2");
    let restored = package_map(&workbook.save().unwrap());
    let before = package_map(&untouched);
    for path in [
        "xl/charts/chart1.xml",
        "xl/drawings/drawing1.xml",
        "xl/drawings/drawing2.xml",
    ] {
        assert_eq!(
            restored.get(path),
            before.get(path),
            "undoing the split lets {path} save unchanged again"
        );
    }
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(shared_chart_formula(&reopened, SheetId(0)), "$A$1:$A$2");
    assert_eq!(shared_chart_formula(&reopened, SheetId(1)), "$A$1:$A$2");
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

/// The charted fixture with a second sheet pointing at the *same* drawing, so
/// one anchor element is held by two sheets at once.
fn shared_drawing_fixture() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let workbook = test_part_text(&parts, "xl/workbook.xml").replace(
        "</sheets>",
        r#"<sheet name="Mirror" sheetId="9" r:id="rIdMirror"/></sheets>"#,
    );
    set_test_part(&mut parts, "xl/workbook.xml", workbook.into_bytes());
    let rels = test_part_text(&parts, "xl/_rels/workbook.xml.rels").replace(
        "</Relationships>",
        r#"<Relationship Id="rIdMirror" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
    );
    set_test_part(&mut parts, "xl/_rels/workbook.xml.rels", rels.into_bytes());
    parts.push((
        "xl/worksheets/sheet2.xml".to_owned(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#.to_vec(),
    ));
    parts.push((
        "xl/worksheets/_rels/sheet2.xml.rels".to_owned(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec(),
    ));
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

/// The charted fixture with a second anchor in the same drawing pointing at
/// the same chart part, so one part backs two frames on one sheet.
fn twin_anchor_charted_fixture() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let drawing = test_part_text(&parts, "xl/drawings/drawing1.xml").replace(
        "</xdr:wsDr>",
        r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChartTwin"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#,
    );
    set_test_part(&mut parts, "xl/drawings/drawing1.xml", drawing.into_bytes());
    let rels = test_part_text(&parts, "xl/drawings/_rels/drawing1.xml.rels").replace(
        "</Relationships>",
        r#"<Relationship Id="rIdChartTwin" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
    );
    set_test_part(
        &mut parts,
        "xl/drawings/_rels/drawing1.xml.rels",
        rels.into_bytes(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Two anchors stacked at the very same spot on two different chart parts,
/// optionally behind a plain shape that shifts both ordinals. Geometry cannot
/// tell these frames apart, so only the chart part does.
fn stacked_charted_fixture(shifted: bool) -> Vec<u8> {
    let anchor = |rel: &str| {
        format!(
            r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="{rel}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#
        )
    };
    let shape = if shifted {
        r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/><xdr:clientData/></xdr:twoCellAnchor>"#
    } else {
        ""
    };
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let header = test_part_text(&parts, "xl/drawings/drawing1.xml")
        .split_once("<xdr:twoCellAnchor")
        .expect("the charted drawing opens with an anchor")
        .0
        .to_owned();
    set_test_part(
        &mut parts,
        "xl/drawings/drawing1.xml",
        format!(
            "{header}{shape}{}{}</xdr:wsDr>",
            anchor("rIdChart"),
            anchor("rIdChartTwo")
        )
        .into_bytes(),
    );
    let rels = test_part_text(&parts, "xl/drawings/_rels/drawing1.xml.rels").replace(
        "</Relationships>",
        r#"<Relationship Id="rIdChartTwo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>"#,
    );
    set_test_part(
        &mut parts,
        "xl/drawings/_rels/drawing1.xml.rels",
        rels.into_bytes(),
    );
    parts.push(("xl/charts/chart2.xml".to_owned(), CHART_PART.to_vec()));
    let content_types = test_part_text(&parts, "[Content_Types].xml").replace(
        "</Types>",
        r#"<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#,
    );
    set_test_part(
        &mut parts,
        "[Content_Types].xml",
        content_types.into_bytes(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The twin-anchor fixture with a plain shape anchored ahead of both charts,
/// as another editor would leave it. Every chart ordinal shifts by one.
fn shifted_twin_anchor_charted_fixture() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&twin_anchor_charted_fixture()).unwrap();
    let drawing = test_part_text(&parts, "xl/drawings/drawing1.xml").replacen(
        "<xdr:twoCellAnchor",
        r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor"#,
        1,
    );
    set_test_part(&mut parts, "xl/drawings/drawing1.xml", drawing.into_bytes());
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

/// `Data` feeds a pivot cache over `A1:B4`, `Report` hosts the pivot table at
/// `D1:F10`, and `Notes` is named by neither.
fn pivoted_fixture() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    for name in ["Data", "Report", "Notes"] {
        model.sheets.push(Sheet::new(name));
    }
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    parts.push((
        "xl/pivotCache/pivotCacheDefinition1.xml".to_owned(),
        br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Data"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    ));
    parts.push((
        "xl/pivotTables/pivotTable1.xml".to_owned(),
        br#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" cacheId="1"><location ref="D1:F10" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>"#.to_vec(),
    ));
    parts.push((
        "xl/worksheets/_rels/sheet2.xml.rels".to_owned(),
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#.to_vec(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A veto that reasons per part has to find a cache the conventional directory
/// scan misses, so one written outside it is typed by its `Override` and still
/// refuses the edits that would strand it.
#[test]
fn refuses_ops_that_would_strand_a_mis_pathed_pivot_cache() {
    let mut model = WorkbookModel::default();
    for name in ["Data", "Notes"] {
        model.sheets.push(Sheet::new(name));
    }
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    parts.push((
        "xl/pivotCacheDefinition1.xml".to_owned(),
        br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Data"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    ));
    let content_types = test_part_text(&parts, "[Content_Types].xml").replace(
        "</Types>",
        r#"<Override PartName="/xl/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/></Types>"#,
    );
    set_test_part(
        &mut parts,
        "[Content_Types].xml",
        content_types.into_bytes(),
    );
    let original = ooxml_opc::rezip_parts(&parts).unwrap();

    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".to_owned(),
        },
        Op::RemoveSheet { index: 0 },
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
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(1),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .expect("an unrelated sheet is still free");
    Workbook::open(&workbook.save().unwrap()).unwrap();
}

/// Sheet ids in a batch mean what the ops before them left behind. A batch
/// that drops a sheet ahead of the pivot source and then edits by the id the
/// source has slid into must be read against the shifted list, not the one the
/// workbook opened with.
#[test]
fn resolves_batched_ops_against_the_sheets_the_batch_leaves() {
    let mut model = WorkbookModel::default();
    for name in ["Notes", "Data", "Other"] {
        model.sheets.push(Sheet::new(name));
    }
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    parts.push((
        "xl/pivotCache/pivotCacheDefinition1.xml".to_owned(),
        br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Data"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    ));
    let original = ooxml_opc::rezip_parts(&parts).unwrap();

    let mut workbook = Workbook::open(&original).unwrap();
    let error = workbook
        .apply_ops(
            vec![
                Op::RemoveSheet { index: 0 },
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
                Op::AddSheet {
                    index: 0,
                    name: "Scratch".to_owned(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::InvalidOperation(message)
            if message.contains("pivotCacheDefinition1.xml")),
        "the batch was allowed: {error:?}"
    );
}

/// The veto is per reference, not per workbook: an edit a pivot's source range
/// and its own location both survive must go through.
#[test]
fn allows_structural_ops_no_pivot_reference_names() {
    for op in [
        Op::InsertRows {
            sheet: SheetId(2),
            at: 0,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(2),
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(2),
            at: 0,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(2),
            at: 0,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(2),
            name: "Scratch".to_owned(),
        },
        Op::RemoveSheet { index: 2 },
        Op::InsertRows {
            sheet: SheetId(0),
            at: 4,
            count: 3,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::InsertRows {
            sheet: SheetId(1),
            at: 10,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(1),
            at: 6,
            count: 1,
        },
    ] {
        let mut workbook = Workbook::open(&pivoted_fixture()).unwrap();
        workbook
            .apply_ops(vec![op.clone()], CalculationOptions::default())
            .unwrap_or_else(|error| panic!("{op:?} was refused: {error:?}"));
        Workbook::open(&workbook.save().unwrap())
            .unwrap_or_else(|error| panic!("{op:?} would not save: {error:?}"));
    }
}

/// What the cache source and the pivot table's own location name still moves
/// out from under a part nothing can rewrite.
#[test]
fn still_refuses_structural_ops_a_pivot_reference_names() {
    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 3,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".to_owned(),
        },
        Op::RemoveSheet { index: 0 },
        Op::InsertRows {
            sheet: SheetId(1),
            at: 9,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(1),
            at: 5,
            count: 1,
        },
        Op::RemoveSheet { index: 1 },
    ] {
        let mut workbook = Workbook::open(&pivoted_fixture()).unwrap();
        let error = workbook
            .apply_ops(vec![op.clone()], CalculationOptions::default())
            .unwrap_err();
        assert!(
            matches!(&error, Error::InvalidOperation(message) if message.contains("pivot")),
            "{op:?} was allowed: {error:?}"
        );
    }
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

    // a frame is addressed by its drawing anchor, so two charts on one anchor
    // are refused even when they name different parts.
    for twin in [sample_chart(), {
        let mut other = sample_chart();
        other.part = "xl/charts/chart2.xml".to_owned();
        other
    }] {
        let mut model = charted_model(sample_chart());
        model.sheets[0].charts.push(twin);
        let Err(error) = Workbook::from_model(model) else {
            panic!("two charts on one anchor must be refused");
        };
        assert!(
            matches!(&error, Error::InvalidOperation(message)
                if message.contains("same drawing anchor")),
            "{error:?}"
        );
    }
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

/// The point the chart paints under resolves to the chart, a point beside it
/// to nothing, and a scrolled-away viewport publishes no region at all.
#[test]
fn a_point_over_a_chart_resolves_to_it_and_a_point_beside_it_does_not() {
    let workbook = Workbook::open(&charted_fixture()).unwrap();
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 800.0,
        height: 600.0,
    };
    let anchored = workbook.chart_at_point(&viewport, 300.0, 150.0).unwrap();
    assert_eq!(
        anchored.map(|chart| chart.id),
        Some("xl/drawings/drawing1.xml#0".to_owned())
    );
    assert_eq!(workbook.chart_at_point(&viewport, 4.0, 4.0).unwrap(), None);
    assert_eq!(
        workbook
            .chart_at_point(
                &Viewport {
                    x: 4_000.0,
                    y: 4_000.0,
                    width: 800.0,
                    height: 600.0,
                },
                300.0,
                150.0,
            )
            .unwrap(),
        None
    );
}

/// Moving a chart is one undo step that survives a save and reopen: the new
/// anchor reaches the drawing part, and undo puts the old one back.
#[test]
fn a_moved_chart_survives_a_save_and_undoes_in_one_step() {
    let mut workbook = Workbook::open(&charted_fixture()).unwrap();
    let before = workbook.model().sheets[0].charts[0].anchor;
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 800.0,
        height: 600.0,
    };
    let region = workbook
        .chart_at_point(&viewport, 300.0, 150.0)
        .unwrap()
        .expect("the fixture anchors a chart under this point");

    assert!(
        workbook
            .move_chart(
                SheetId(0),
                &region.id,
                70.0,
                45.0,
                CalculationOptions::default()
            )
            .unwrap()
            .applied
    );
    let moved = workbook.model().sheets[0].charts[0].anchor;
    assert_ne!(moved, before);

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.model().sheets[0].charts[0].anchor, moved);

    workbook.undo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().sheets[0].charts[0].anchor, before);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_eq!(workbook.model().sheets[0].charts[0].anchor, moved);
}

/// One chart part can back two anchors on a sheet, so a part cannot name a
/// frame. Each frame hit-tests to its own id, moves alone, and reaches the
/// anchor it was read from on save.
#[test]
fn twin_anchors_on_one_chart_part_move_independently() {
    let mut workbook = Workbook::open(&twin_anchor_charted_fixture()).unwrap();
    let charts = &workbook.model().sheets[0].charts;
    assert_eq!(charts.len(), 2);
    assert_eq!(charts[0].part, charts[1].part);
    let before = [charts[0].anchor, charts[1].anchor];
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 800.0,
        height: 600.0,
    };

    let left = workbook
        .chart_at_point(&viewport, 300.0, 150.0)
        .unwrap()
        .expect("the fixture anchors a frame under this point");
    let right = workbook
        .chart_at_point(&viewport, 620.0, 100.0)
        .unwrap()
        .expect("the fixture anchors a second frame under this point");
    assert_ne!(
        left.id, right.id,
        "two frames backed by one part must not share an id"
    );

    assert!(
        workbook
            .move_chart(
                SheetId(0),
                &right.id,
                70.0,
                45.0,
                CalculationOptions::default()
            )
            .unwrap()
            .applied
    );
    let moved = [
        workbook.model().sheets[0].charts[0].anchor,
        workbook.model().sheets[0].charts[1].anchor,
    ];
    assert_eq!(
        moved[0], before[0],
        "the frame that was not dragged must stay where it was"
    );
    assert_ne!(moved[1], before[1], "the dragged frame must have moved");

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .charts
            .iter()
            .map(|chart| chart.anchor)
            .collect::<Vec<_>>(),
        moved
    );
}

/// An anchor ordinal names a position in a drawing, not an object, so an op
/// stored against one outlives the structure it was recorded against. A move
/// replayed onto a drawing that has since gained an anchor must be refused,
/// never quietly landed on whichever frame now sits at that ordinal.
#[test]
fn a_stored_chart_move_refuses_a_drawing_whose_anchors_shifted() {
    let stored = {
        let workbook = Workbook::open(&twin_anchor_charted_fixture()).unwrap();
        let charts = &workbook.model().sheets[0].charts;
        let right = &charts[1];
        assert_eq!(right.frame_id(), "xl/drawings/drawing1.xml#1");
        Op::SetChartAnchor {
            sheet: SheetId(0),
            frame: right.frame_id(),
            part: right.part.clone(),
            from: right.anchor,
            to: nudged_anchor(right.anchor),
        }
    };

    let mut shifted = Workbook::open(&shifted_twin_anchor_charted_fixture()).unwrap();
    let before = shifted.model().sheets[0]
        .charts
        .iter()
        .map(|chart| (chart.frame_id(), chart.anchor))
        .collect::<Vec<_>>();
    assert_eq!(
        before.iter().map(|(id, _)| id.as_str()).collect::<Vec<_>>(),
        ["xl/drawings/drawing1.xml#1", "xl/drawings/drawing1.xml#2"],
        "the inserted shape must shift both chart ordinals"
    );

    let error = shifted
        .apply_ops(vec![stored.clone()], CalculationOptions::default())
        .unwrap_err();
    // typed, not prose: a host replaying a stored log must be able to drop
    // this op and carry on without reading the message.
    assert!(
        matches!(&error, Error::ChartFrameShifted { frame }
            if frame == "xl/drawings/drawing1.xml#1"),
        "{error:?}"
    );
    assert_eq!(
        shifted.model().sheets[0]
            .charts
            .iter()
            .map(|chart| (chart.frame_id(), chart.anchor))
            .collect::<Vec<_>>(),
        before,
        "a refused replay must leave every anchor alone"
    );

    // the same op against the drawing it was recorded on still lands.
    let mut unshifted = Workbook::open(&twin_anchor_charted_fixture()).unwrap();
    let untouched = unshifted.model().sheets[0].charts[0].anchor;
    assert!(
        unshifted
            .apply_ops(vec![stored], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(unshifted.model().sheets[0].charts[0].anchor, untouched);
    assert_ne!(
        unshifted.model().sheets[0].charts[1].anchor,
        before[1].1,
        "the frame the op named must have moved"
    );
}

/// Two frames stacked at one spot are identical to the geometry guard, so the
/// anchor alone cannot tell a shifted ordinal from the frame it was recorded
/// against. The chart part carried beside it can, and must.
#[test]
fn a_stored_chart_move_refuses_a_shifted_ordinal_at_an_identical_anchor() {
    let stored = {
        let workbook = Workbook::open(&stacked_charted_fixture(false)).unwrap();
        let charts = &workbook.model().sheets[0].charts;
        assert_eq!(charts[0].anchor, charts[1].anchor, "the frames must stack");
        assert_ne!(charts[0].part, charts[1].part);
        let second = &charts[1];
        assert_eq!(second.frame_id(), "xl/drawings/drawing1.xml#1");
        Op::SetChartAnchor {
            sheet: SheetId(0),
            frame: second.frame_id(),
            part: second.part.clone(),
            from: second.anchor,
            to: nudged_anchor(second.anchor),
        }
    };

    let mut shifted = Workbook::open(&stacked_charted_fixture(true)).unwrap();
    let before = shifted.model().sheets[0]
        .charts
        .iter()
        .map(|chart| (chart.frame_id(), chart.part.clone(), chart.anchor))
        .collect::<Vec<_>>();
    assert_eq!(
        before
            .iter()
            .map(|(id, part, _)| (id.as_str(), part.as_str()))
            .collect::<Vec<_>>(),
        [
            ("xl/drawings/drawing1.xml#1", "xl/charts/chart1.xml"),
            ("xl/drawings/drawing1.xml#2", "xl/charts/chart2.xml"),
        ],
        "the inserted shape must slide chart1 onto the ordinal chart2 held"
    );

    let error = shifted
        .apply_ops(vec![stored], CalculationOptions::default())
        .unwrap_err();
    assert!(
        matches!(&error, Error::ChartFrameShifted { frame }
            if frame == "xl/drawings/drawing1.xml#1"),
        "{error:?}"
    );
    assert_eq!(
        shifted.model().sheets[0]
            .charts
            .iter()
            .map(|chart| (chart.frame_id(), chart.part.clone(), chart.anchor))
            .collect::<Vec<_>>(),
        before,
        "a refused replay must leave every anchor alone"
    );
}

/// The same anchor one row down.
fn nudged_anchor(anchor: ChartAnchor) -> ChartAnchor {
    match anchor {
        ChartAnchor::TwoCell { from, to, edit_as } => ChartAnchor::TwoCell {
            from: AnchorCell {
                row: from.row + 1,
                ..from
            },
            to: AnchorCell {
                row: to.row + 1,
                ..to
            },
            edit_as,
        },
        other => other,
    }
}

/// A drawing that omitted `colOff`/`rowOff` reads as zero, so a sub-cell move
/// has no span to write into. The save must not report success while dropping
/// the offsets — the reopened anchor has to match what the move produced.
#[test]
fn a_sub_cell_move_survives_a_drawing_that_wrote_no_offsets() {
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let drawing = String::from_utf8(CHART_DRAWING.to_vec())
        .unwrap()
        .replace("<xdr:colOff>12700</xdr:colOff>", "")
        .replace("<xdr:colOff>0</xdr:colOff>", "")
        .replace("<xdr:rowOff>0</xdr:rowOff>", "");
    set_test_part(&mut parts, "xl/drawings/drawing1.xml", drawing.into_bytes());
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();

    workbook
        .move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#0",
            17.0,
            9.0,
            CalculationOptions::default(),
        )
        .unwrap();
    let moved = workbook.model().sheets[0].charts[0].anchor;
    let AnchorCell {
        col_off, row_off, ..
    } = moved.from_cell().expect("a grid-anchored chart");
    assert!(col_off > 0 && row_off > 0, "{moved:?}");

    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.model().sheets[0].charts[0].anchor, moved);
}

/// A chart pinned to the sheet cannot be moved, because a save cannot rewrite
/// the attributes that carry its position; an unknown frame names nothing.
#[test]
fn moving_refuses_an_absolute_anchor_and_an_unknown_frame() {
    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    let drawing = String::from_utf8(CHART_DRAWING.to_vec())
        .unwrap()
        .replace(
            r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"#,
            r#"<xdr:absoluteAnchor><xdr:pos x="95250" y="190500"/><xdr:ext cx="1905000" cy="1143000"/>"#,
        )
        .replace("</xdr:twoCellAnchor>", "</xdr:absoluteAnchor>");
    set_test_part(&mut parts, "xl/drawings/drawing1.xml", drawing.into_bytes());
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let error = workbook
        .move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#0",
            10.0,
            10.0,
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::InvalidOperation(message) if message.contains("pinned")),
        "{error:?}"
    );

    let error = workbook
        .move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#7",
            10.0,
            10.0,
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::InvalidOperation(message)
            if message.contains("xl/drawings/drawing1.xml#7")),
        "{error:?}"
    );
}

/// `SetChartAnchor` accepts exactly what a save can write back, and each
/// refusal says which part of the anchor the writer could not carry.
#[test]
fn set_chart_anchor_refuses_what_a_save_cannot_write() {
    let mut workbook = Workbook::open(&charted_fixture()).unwrap();
    let repin = |workbook: &mut Workbook, to| {
        let recorded = &workbook.model().sheets[0].charts[0];
        let (part, from) = (recorded.part.clone(), recorded.anchor);
        workbook.apply_ops(
            vec![Op::SetChartAnchor {
                sheet: SheetId(0),
                frame: "xl/drawings/drawing1.xml#0".to_owned(),
                part,
                from,
                to,
            }],
            CalculationOptions::default(),
        )
    };
    for (anchor, reason) in [
        (
            ChartAnchor::TwoCell {
                from: AnchorCell::default(),
                to: AnchorCell {
                    col: 4,
                    row: 8,
                    ..AnchorCell::default()
                },
                edit_as: AnchorEditAs::TwoCell,
            },
            "follows a grid edit",
        ),
        (
            ChartAnchor::OneCell {
                from: AnchorCell::default(),
                extent: AnchorExtent {
                    cx: 100_000,
                    cy: 100_000,
                },
            },
            "anchor kind",
        ),
    ] {
        let error = repin(&mut workbook, anchor).unwrap_err();
        assert!(
            matches!(&error, Error::InvalidOperation(message) if message.contains(reason)),
            "{error:?}"
        );
    }

    // both corners of a two-cell anchor are written whole, so a resize is
    // expressible and must round-trip rather than be half-saved.
    let ChartAnchor::TwoCell { from, to, edit_as } = workbook.model().sheets[0].charts[0].anchor
    else {
        panic!("two-cell anchor");
    };
    let resized = ChartAnchor::TwoCell {
        from,
        to: AnchorCell {
            col: to.col + 3,
            row: to.row + 5,
            ..to
        },
        edit_as,
    };
    assert!(repin(&mut workbook, resized).unwrap().applied);
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.model().sheets[0].charts[0].anchor, resized);
}

/// A chart anchor is replicated content, not workbook structure: the freeze
/// pins which drawing anchors which part, and where that anchor sits travels
/// through the document like any other edit.
#[test]
fn collaborative_sessions_move_a_chart_and_converge() {
    let bytes = charted_fixture();
    let mut left = Workbook::open_collaborative(&bytes, 303).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 304).unwrap();
    let before = left.model().sheets[0].charts[0].anchor;

    assert!(
        left.move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#0",
            12.0,
            8.0,
            CalculationOptions::default(),
        )
        .unwrap()
        .applied
    );
    let moved = left.model().sheets[0].charts[0].anchor;
    assert_ne!(moved, before);

    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    assert!(
        right
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(right.model().sheets[0].charts[0].anchor, moved);
    assert_eq!(right.model(), left.model());

    // the move survives a save, and a later peer edit still merges.
    let reopened = Workbook::open(&left.save().unwrap()).unwrap();
    assert_eq!(reopened.model().sheets[0].charts[0].anchor, moved);
    right
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "after",
            CalculationOptions::default(),
        )
        .unwrap();
    let update = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    left.apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(left.model().sheets[0].charts[0].anchor, moved);
    assert_eq!(left.cell(SheetId(0), cell("A1")).unwrap().input, "after");

    // the mover can take the drag back, and the peer follows it home.
    assert!(left.can_undo());
    assert!(left.undo(CalculationOptions::default()).unwrap().applied);
    assert_eq!(left.model().sheets[0].charts[0].anchor, before);
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(right.model().sheets[0].charts[0].anchor, before);
}

/// Letting an anchor travel does not unfreeze the rest of a chart: an op that
/// remaps what a chart reads is still structural, refused locally and refused
/// again when a standalone peer offers it. This rides on the structure
/// generation every structural op bumps; the identity fields themselves are
/// pinned by `a_peer_cannot_disguise_a_chart_remap_as_a_move`, which leaves
/// that counter alone.
#[test]
fn collaborative_sessions_still_refuse_a_chart_remap() {
    let bytes = charted_fixture();
    let mut workbook = Workbook::open_collaborative(&bytes, 305).unwrap();
    let before = workbook.model().clone();
    let insert_rows = Op::InsertRows {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    };

    let error = workbook
        .apply_ops(vec![insert_rows.clone()], CalculationOptions::default())
        .unwrap_err();
    assert!(
        matches!(&error, Error::CollaborativeStructureOperation),
        "{error:?}"
    );
    assert_eq!(workbook.model(), &before);

    let mut standalone = Workbook::open(&bytes).unwrap();
    standalone
        .apply_ops(vec![insert_rows], CalculationOptions::default())
        .unwrap();
    assert_ne!(
        standalone.model().sheets[0].charts[0].refs,
        before.sheets[0].charts[0].refs
    );
    let update = standalone
        .encode_diff_v1(&workbook.encode_state_vector_v1())
        .unwrap();
    assert!(
        matches!(
            workbook.apply_update_v1(&update, CalculationOptions::default()),
            Err(Error::CollaborativeStructureChanged)
        ),
        "a chart remap must not slip past the freeze"
    );
    assert_eq!(workbook.model(), &before);
}

/// Two replicas dragging the same chart at once must land on one anchor, and
/// on the same one whichever order the updates arrive in.
#[test]
fn concurrent_chart_moves_converge_in_either_delivery_order() {
    let bytes = charted_fixture();
    let settle = |first_wins: bool| {
        let mut left = Workbook::open_collaborative(&bytes, 401).unwrap();
        let mut right = Workbook::open_collaborative(&bytes, 402).unwrap();
        left.move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#0",
            16.0,
            0.0,
            CalculationOptions::default(),
        )
        .unwrap();
        right
            .move_chart(
                SheetId(0),
                "xl/drawings/drawing1.xml#0",
                0.0,
                24.0,
                CalculationOptions::default(),
            )
            .unwrap();
        let to_right = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let to_left = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        if first_wins {
            right
                .apply_update_v1(&to_right, CalculationOptions::default())
                .unwrap();
            left.apply_update_v1(&to_left, CalculationOptions::default())
                .unwrap();
        } else {
            left.apply_update_v1(&to_left, CalculationOptions::default())
                .unwrap();
            right
                .apply_update_v1(&to_right, CalculationOptions::default())
                .unwrap();
        }
        let anchor = left.model().sheets[0].charts[0].anchor;
        assert_eq!(
            right.model().sheets[0].charts[0].anchor,
            anchor,
            "concurrent moves must converge"
        );
        anchor
    };
    assert_eq!(
        settle(true),
        settle(false),
        "the winning anchor must not depend on delivery order"
    );
}

/// One drawing anchor held by two sheets is one element in one part, so a drag
/// repins every sheet holding it. Repinning only the dragged sheet would leave
/// the two disagreeing and the save would refuse the drawing outright.
#[test]
fn a_drag_repins_every_sheet_sharing_the_drawing() {
    let bytes = shared_drawing_fixture();
    let mut workbook = Workbook::open_collaborative(&bytes, 403).unwrap();
    assert_eq!(
        workbook.model().sheets[1].charts[0].drawing,
        workbook.model().sheets[0].charts[0].drawing,
        "the fixture must share one drawing"
    );
    let before = workbook.model().sheets[0].charts[0].anchor;

    assert!(
        workbook
            .move_chart(
                SheetId(0),
                "xl/drawings/drawing1.xml#0",
                18.0,
                9.0,
                CalculationOptions::default(),
            )
            .unwrap()
            .applied
    );
    let moved = workbook.model().sheets[0].charts[0].anchor;
    assert_ne!(moved, before);
    assert_eq!(
        workbook.model().sheets[1].charts[0].anchor,
        moved,
        "the sheet sharing the anchor must follow it"
    );

    let saved = workbook
        .save()
        .expect("a shared drawing both sheets agree on must save");
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(reopened.model().sheets[0].charts[0].anchor, moved);
    assert_eq!(reopened.model().sheets[1].charts[0].anchor, moved);

    // one undo takes both sheets back together.
    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.model().sheets[0].charts[0].anchor, before);
    assert_eq!(workbook.model().sheets[1].charts[0].anchor, before);
}

/// The chart references `charted_fixture` parses, so a handcrafted peer update
/// can leave the chart's identity alone and change only its anchor.
const CHARTED_FIXTURE_REFS: &str = r#"[{"kind":"seriesName","formula":"Data!$B$1"},{"kind":"categories","formula":"Data!$A$2:$A$4"},{"kind":"values","formula":"Data!$B$2:$B$4"},{"kind":"dataLabels","formula":"Data!$C$2:$C$4"}]"#;

/// A whole document, as a persisted snapshot would be, forked from `target`
/// and carrying a handcrafted chart state on one sheet.
fn peer_snapshot_with_charts(
    target: &Workbook,
    client_id: u64,
    sheet_key: &str,
    charts: &str,
) -> Vec<u8> {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, MapRef, ReadTxn, StateVector, Transact, Update};

    let peer = Doc::with_client_id(client_id);
    let update = Update::decode_v1(&target.encode_state_as_update_v1()).unwrap();
    peer.transact_mut().apply_update(update).unwrap();
    {
        let mut txn = peer.transact_mut();
        let sheets = txn.get_map("xlsx:sheets").unwrap();
        let sheet = sheets
            .get(&txn, sheet_key)
            .and_then(|value| value.cast::<MapRef>().ok())
            .unwrap();
        sheet.try_update(&mut txn, "charts", charts);
    }
    peer.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

/// A drawing whose anchor runs backwards — `to` before `from`. Real files
/// carry such things and open fine, but nothing can resolve them, so they are
/// grandfathered rather than accepted as a new repin.
fn unresolvable_anchor_fixture() -> Vec<u8> {
    const INVERTED_DRAWING: &[u8] = br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    set_test_part(
        &mut parts,
        "xl/drawings/drawing1.xml",
        INVERTED_DRAWING.to_vec(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A drawing whose two anchors both point at one chart part, so a sheet holds
/// that part twice and naming it alone cannot say which anchor is meant.
fn two_anchor_fixture() -> Vec<u8> {
    const TWO_ANCHOR_DRAWING: &[u8] = br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

    let mut parts = ooxml_opc::unzip_parts(&charted_fixture()).unwrap();
    set_test_part(
        &mut parts,
        "xl/drawings/drawing1.xml",
        TWO_ANCHOR_DRAWING.to_vec(),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A peer's update assigning one sheet's chart state, built by forking the
/// target's own document. Mirrors the shared-document key names.
fn peer_sheet_charts_update(
    target: &Workbook,
    client_id: u64,
    sheet_key: &str,
    charts: &str,
) -> Vec<u8> {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, MapRef, ReadTxn, Transact, Update};

    let peer = Doc::with_client_id(client_id);
    let update = Update::decode_v1(&target.encode_state_as_update_v1()).unwrap();
    peer.transact_mut().apply_update(update).unwrap();
    let before = peer.transact().state_vector();
    {
        let mut txn = peer.transact_mut();
        let sheets = txn.get_map("xlsx:sheets").unwrap();
        let sheet = sheets
            .get(&txn, sheet_key)
            .and_then(|value| value.cast::<MapRef>().ok())
            .unwrap();
        sheet.try_update(&mut txn, "charts", charts);
    }
    peer.transact().encode_diff_v1(&before)
}

/// Two people dragging the same chart is ordinary use, not an attack. Once the
/// two moves merge, the assignment that lost is a tombstone, and undoing the
/// one that won must not take the whole sheet's chart state with it: the loser
/// keeps a readable workbook, the winner gets its drag back, and neither is
/// left holding history it can no longer apply.
#[test]
fn honest_concurrent_drags_survive_undo_on_both_replicas() {
    let bytes = charted_fixture();
    let mut left = Workbook::open_collaborative(&bytes, 800).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 801).unwrap();
    let before = left.model().sheets[0].charts[0].anchor;

    left.move_chart(
        SheetId(0),
        "xl/drawings/drawing1.xml#0",
        16.0,
        0.0,
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#0",
            0.0,
            24.0,
            CalculationOptions::default(),
        )
        .unwrap();
    let to_right = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let to_left = right
        .encode_diff_v1(&left.encode_state_vector_v1())
        .unwrap();
    right
        .apply_update_v1(&to_right, CalculationOptions::default())
        .unwrap();
    left.apply_update_v1(&to_left, CalculationOptions::default())
        .unwrap();
    let converged = left.model().sheets[0].charts[0].anchor;
    assert_eq!(right.model().sheets[0].charts[0].anchor, converged);

    // whichever replica owns the surviving assignment can take it back, and
    // neither replica may fail the step or lose the workbook doing it.
    for (label, workbook) in [("left", &mut left), ("right", &mut right)] {
        let undone = workbook
            .undo(CalculationOptions::default())
            .unwrap_or_else(|error| {
                panic!("{label} could not undo after a concurrent drag: {error}")
            });
        assert!(
            workbook.save().is_ok(),
            "{label} must still hold a saveable workbook"
        );
        assert!(
            !workbook.model().sheets[0].charts.is_empty(),
            "{label} must not lose its chart"
        );
        let _ = undone;
    }
    assert_eq!(
        right.model().sheets[0].charts[0].anchor,
        before,
        "the replica whose drag survived the merge must get it back"
    );
}

/// An anchor can be wrong on its own terms — a negative offset, one that is
/// not a position at all, corners in the wrong order — and none of that
/// depends on the grid it sits over. That is what a peer may not introduce.
/// The open path must keep taking whatever real files carry, so this is
/// refused where the update arrives rather than where the file is read.
#[test]
fn remote_updates_carrying_an_intrinsically_broken_anchor_are_refused() {
    let bytes = charted_fixture();
    for (label, reason, anchor) in [
        (
            "negative offset",
            "offset is out of range",
            r#"{"kind":"twoCell","from":{"col":2,"colOff":-1,"row":4,"rowOff":0},"to":{"col":8,"colOff":0,"row":19,"rowOff":0},"edit_as":"oneCell"}"#,
        ),
        (
            "overflowing offset",
            "offset is out of range",
            r#"{"kind":"twoCell","from":{"col":2,"colOff":9223372036854775807,"row":4,"rowOff":0},"to":{"col":8,"colOff":0,"row":19,"rowOff":0},"edit_as":"oneCell"}"#,
        ),
        (
            "inverted corners",
            "inverted or coincident",
            r#"{"kind":"twoCell","from":{"col":20,"colOff":0,"row":4,"rowOff":0},"to":{"col":8,"colOff":0,"row":19,"rowOff":0},"edit_as":"oneCell"}"#,
        ),
    ] {
        let mut target = Workbook::open_collaborative(&bytes, 820).unwrap();
        let before = target.model().clone();
        let charts = format!(
            r#"[{{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{anchor},"refs":{CHARTED_FIXTURE_REFS}}}]"#
        );
        let update = peer_sheet_charts_update(&target, 821, "sheet:0", &charts);
        let error = target
            .apply_update_v1(&update, CalculationOptions::default())
            .expect_err(&format!("{label} must be refused"));
        assert!(
            matches!(&error, Error::CollaborativeState(message) if message.contains(reason)),
            "{label}: {error:?}"
        );
        assert_eq!(target.model(), &before, "{label} must leave nothing behind");
    }
}

/// Column widths and chart anchors are replicated independently, and both of
/// these edits are ordinary. Collapsing the columns a chart spans leaves it
/// with no room to draw, but that is a fact about one replica's grid, not
/// about the anchor — so it cannot decide whether a peer's move is allowed.
/// Judging it that way makes each replica reject what the other accepted.
#[test]
fn collapsing_columns_under_a_chart_does_not_block_a_peer_moving_it() {
    let bytes = charted_fixture();
    let collapse: Vec<Op> = (2..8)
        .map(|col| Op::SetColWidth {
            sheet: SheetId(0),
            col,
            width: Some(0.0),
        })
        .collect();

    for forward_first in [true, false] {
        let mut widener = Workbook::open_collaborative(&bytes, 870).unwrap();
        let mut mover = Workbook::open_collaborative(&bytes, 871).unwrap();

        widener
            .apply_ops(collapse.clone(), CalculationOptions::default())
            .expect("collapsing a column is an ordinary edit");
        mover
            .move_chart(
                SheetId(0),
                "xl/drawings/drawing1.xml#0",
                24.0,
                0.0,
                CalculationOptions::default(),
            )
            .expect("moving a chart is an ordinary edit");

        let to_mover = widener
            .encode_diff_v1(&mover.encode_state_vector_v1())
            .unwrap();
        let to_widener = mover
            .encode_diff_v1(&widener.encode_state_vector_v1())
            .unwrap();
        if forward_first {
            mover
                .apply_update_v1(&to_mover, CalculationOptions::default())
                .expect("a collapsed column must not be refused by the replica that moved");
            widener
                .apply_update_v1(&to_widener, CalculationOptions::default())
                .expect("a moved chart must not be refused by the replica that collapsed");
        } else {
            widener
                .apply_update_v1(&to_widener, CalculationOptions::default())
                .expect("a moved chart must not be refused by the replica that collapsed");
            mover
                .apply_update_v1(&to_mover, CalculationOptions::default())
                .expect("a collapsed column must not be refused by the replica that moved");
        }

        assert_eq!(
            widener.model(),
            mover.model(),
            "both edits are supported, so the replicas must converge"
        );
        assert_eq!(widener.model().sheets[0].col_widths.get(&2), Some(&0.0));
        assert!(widener.save().is_ok(), "the union must be saveable");
        assert!(mover.save().is_ok());
    }
}

/// A local batch that repins only one of two sheets sharing an anchor is
/// refused: it is this replica's own edit and it can simply be told no. The
/// same disagreement arriving as an update cannot be refused — a peer may hold
/// the other half legally — so it is projected onto one anchor instead, the
/// same way on every replica.
#[test]
fn sheets_sharing_a_drawing_anchor_may_not_disagree() {
    let bytes = shared_drawing_fixture();
    let before = Workbook::open(&bytes).unwrap().model().sheets[0].charts[0].anchor;
    let moved = {
        let mut source = Workbook::open(&bytes).unwrap();
        source
            .move_chart(
                SheetId(0),
                "xl/drawings/drawing1.xml#0",
                18.0,
                9.0,
                CalculationOptions::default(),
            )
            .unwrap();
        source.model().sheets[0].charts[0].anchor
    };

    // a raw op batch that repins only one of the two sheets
    let mut workbook = Workbook::open(&bytes).unwrap();
    let error = workbook
        .apply_ops(
            vec![Op::SetChartAnchor {
                sheet: SheetId(0),
                frame: "xl/drawings/drawing1.xml#0".to_owned(),
                part: "xl/charts/chart1.xml".to_owned(),
                from: before,
                to: moved,
            }],
            CalculationOptions::default(),
        )
        .expect_err("a half-repinned shared drawing must be refused");
    assert!(
        matches!(&error, Error::InvalidOperation(message) if message.contains("disagree")),
        "{error:?}"
    );

    // the same disagreement arriving as an update is taken and projected
    let mut target = Workbook::open_collaborative(&bytes, 830).unwrap();
    let charts = format!(
        r#"[{{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{},"refs":{}}}]"#,
        serde_json::to_string(&moved).unwrap(),
        CHARTED_FIXTURE_REFS
    );
    let update = peer_sheet_charts_update(&target, 831, "sheet:0", &charts);
    target
        .apply_update_v1(&update, CalculationOptions::default())
        .expect("a half-repin from a peer must integrate rather than be refused");
    assert_eq!(
        target.model().sheets[0].charts[0].anchor,
        target.model().sheets[1].charts[0].anchor,
        "the projection must leave one anchor for the frame"
    );
    assert_eq!(
        target.model().sheets[1].charts[0].anchor,
        moved,
        "the first sheet in order donates the anchor, and it is the one written"
    );
    assert!(
        target.save().is_ok(),
        "the projected workbook must still be saveable"
    );
}

/// Three concurrent updates, the same three, delivered in two orders. A gate
/// that can reject one of them rejects a different one on each replica —
/// whichever arrives when the sheets happen to disagree — and the two never
/// meet again. Integration therefore has to take all three whatever the order,
/// and settle the disagreement on the way out.
#[test]
fn one_frame_converges_under_every_delivery_order_of_the_same_updates() {
    let bytes = shared_drawing_fixture();
    let baseline = Workbook::open_collaborative(&bytes, 880).unwrap();
    let baseline_vector = baseline.encode_state_vector_v1();

    // an engine move writes both sheets holding the frame; the handcrafted one
    // writes a single sheet. Yrs settles a tie by client id, so these ids fix
    // the priority at handcrafted > later move > earlier move.
    let engine_move = |client_id: u64, dx: f32| {
        let mut replica = Workbook::open_collaborative(&bytes, client_id).unwrap();
        replica
            .move_chart(
                SheetId(0),
                "xl/drawings/drawing1.xml#0",
                dx,
                0.0,
                CalculationOptions::default(),
            )
            .unwrap();
        (
            replica.encode_diff_v1(&baseline_vector).unwrap(),
            replica.model().sheets[0].charts[0].anchor,
        )
    };
    let (update_b, anchor_b) = engine_move(881, 12.0);
    let (update_c, _) = engine_move(890, 30.0);
    let half = format!(
        r#"[{{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{},"refs":{}}}]"#,
        serde_json::to_string(&anchor_b).unwrap(),
        CHARTED_FIXTURE_REFS
    );
    let update_h = peer_sheet_charts_update(&baseline, 899, "sheet:1", &half);

    let settle = |order: [&Vec<u8>; 3], client_id: u64| {
        let mut replica = Workbook::open_collaborative(&bytes, client_id).unwrap();
        for update in order {
            replica
                .apply_update_v1(update, CalculationOptions::default())
                .unwrap_or_else(|error| {
                    panic!("integration must not depend on delivery order: {error}")
                });
        }
        assert_eq!(
            replica.model().sheets[0].charts[0].anchor,
            replica.model().sheets[1].charts[0].anchor,
            "one frame must end up with one anchor"
        );
        assert!(replica.save().is_ok(), "the settled workbook must save");
        replica.model().clone()
    };

    let first = settle([&update_b, &update_h, &update_c], 870);
    let second = settle([&update_h, &update_b, &update_c], 871);
    assert_eq!(
        first, second,
        "the same updates in a different order must land on the same workbook"
    );
}

/// A replica may not accept an edit its peers will refuse. Publishing one is
/// worse than losing it: the offending value stays the winner in every later
/// diff, so the peer drops that update and every update after it, and the two
/// part ways for the rest of the session. The local gate therefore has to hold
/// the anchor to exactly what an arriving one is held to — the two once
/// compared corners differently, one in pixels and one on the grid, and an
/// offset large enough to cross a column told them apart.
#[test]
fn a_locally_accepted_repin_is_one_a_peer_can_take() {
    let bytes = charted_fixture();
    let mut local = Workbook::open_collaborative(&bytes, 940).unwrap();
    let mut peer = Workbook::open_collaborative(&bytes, 941).unwrap();
    let from = local.model().sheets[0].charts[0].anchor;
    let ChartAnchor::TwoCell { edit_as, .. } = from else {
        panic!("two-cell anchor");
    };

    // corners that run backwards on the grid but forwards in pixels, because
    // the offset more than covers the column it steps back over
    let crossed = ChartAnchor::TwoCell {
        from: AnchorCell {
            col: 2,
            col_off: 12_700,
            row: 4,
            row_off: 0,
        },
        to: AnchorCell {
            col: 1,
            col_off: 6_000_000,
            row: 16,
            row_off: 0,
        },
        edit_as,
    };
    let batch = vec![
        Op::SetCell {
            sheet: SheetId(0),
            at: cell("A1"),
            cell: CellState {
                value: CellValue::Text {
                    value: "important".to_owned(),
                },
                ..CellState::default()
            },
        },
        Op::SetChartAnchor {
            sheet: SheetId(0),
            frame: "xl/drawings/drawing1.xml#0".to_owned(),
            part: "xl/charts/chart1.xml".to_owned(),
            from,
            to: crossed,
        },
    ];

    let refused = local
        .apply_ops(batch, CalculationOptions::default())
        .expect_err("an anchor no peer would take must not be accepted locally");
    assert!(
        matches!(&refused, Error::InvalidOperation(message) if message.contains("inverted or coincident")),
        "{refused:?}"
    );

    // nothing was kept, and the session carries on in step
    local
        .edit_cell(
            SheetId(0),
            cell("A2"),
            "later",
            CalculationOptions::default(),
        )
        .unwrap();
    let update = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&update, CalculationOptions::default())
        .expect("a peer must still be able to take what this replica publishes");
    assert_eq!(
        peer.model(),
        local.model(),
        "the replicas must stay in step"
    );
    assert_eq!(peer.cell(SheetId(0), cell("A2")).unwrap().input, "later");
}

/// A snapshot is foreign bytes like any update, so it answers to the same
/// checks. A pristine replica adopts a whole document instead of merging it,
/// and if that door skipped the anchor check the two replicas would take
/// opposite decisions on identical bytes and drift apart for good.
#[test]
fn an_adopted_snapshot_answers_to_the_same_anchor_check_as_an_update() {
    let bytes = charted_fixture();
    let hostile = {
        let donor = Workbook::open_collaborative(&bytes, 840).unwrap();
        let charts = format!(
            r#"[{{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{{"kind":"twoCell","from":{{"col":20,"colOff":0,"row":4,"rowOff":0}},"to":{{"col":8,"colOff":0,"row":19,"rowOff":0}},"edit_as":"oneCell"}},"refs":{}}}]"#,
            CHARTED_FIXTURE_REFS
        );
        peer_snapshot_with_charts(&donor, 841, "sheet:0", &charts)
    };

    let mut pristine = Workbook::open_collaborative(&bytes, 842).unwrap();
    let pristine_result = pristine.apply_update_v1(&hostile, CalculationOptions::default());

    let mut touched = Workbook::open_collaborative(&bytes, 842).unwrap();
    touched
        .edit_cell(
            SheetId(0),
            cell("A1"),
            "typed",
            CalculationOptions::default(),
        )
        .unwrap();
    let touched_result = touched.apply_update_v1(&hostile, CalculationOptions::default());

    assert!(
        pristine_result.is_err(),
        "a pristine replica must not adopt a snapshot it would reject as an update"
    );
    assert_eq!(
        pristine_result.is_err(),
        touched_result.is_err(),
        "the same bytes must get the same answer whether or not the replica is pristine"
    );
}

/// An undo assembles a model out of parts that each arrived on their own. It
/// can land a pair no single edit produced: one sheet repinned, the sheet
/// sharing its drawing anchor left behind. Projecting that back onto one
/// anchor keeps the step working and the workbook saveable, and keeps the
/// replica publishing something its peers can take.
#[test]
fn collaborative_undo_settles_a_split_shared_drawing() {
    let bytes = shared_drawing_fixture();
    let mut local = Workbook::open_collaborative(&bytes, 850).unwrap();
    let before = local.model().sheets[0].charts[0].anchor;

    local
        .move_chart(
            SheetId(0),
            "xl/drawings/drawing1.xml#0",
            18.0,
            9.0,
            CalculationOptions::default(),
        )
        .unwrap();
    let moved = local.model().sheets[0].charts[0].anchor;
    assert_eq!(local.model().sheets[1].charts[0].anchor, moved);

    // a peer causally overwrites only the second sheet with the same anchor it
    // already holds, tombstoning the local insertion there but not on the first
    // same model, different bytes, so the write lands and tombstones the
    // local insertion on this sheet only
    let charts = format!(
        r#"[ {{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{},"refs":{}}} ]"#,
        serde_json::to_string(&moved).unwrap(),
        CHARTED_FIXTURE_REFS
    );
    let update = peer_sheet_charts_update(&local, 851, "sheet:1", &charts);
    assert!(
        local
            .apply_update_v1(&update, CalculationOptions::default())
            .unwrap()
            .applied
            || local.model().sheets[1].charts[0].anchor == moved,
        "the peer write must land, or the split never forms"
    );
    assert_eq!(local.model().sheets[0].charts[0].anchor, moved);
    assert_eq!(local.model().sheets[1].charts[0].anchor, moved);

    // the split the undo would leave is projected back onto one anchor
    let history_before = local.history_state();
    assert!(
        local
            .undo(CalculationOptions::default())
            .expect("an undo must not be refused over a split a projection can settle")
            .applied
    );
    assert_eq!(
        local.model().sheets[0].charts[0].anchor,
        local.model().sheets[1].charts[0].anchor,
        "an undo must not leave two sheets disagreeing on one drawing anchor"
    );
    assert!(local.save().is_ok(), "an installed model must be saveable");
    assert_ne!(local.history_state(), history_before, "the undo must count");

    // and what the replica publishes is what a peer ends up holding
    let mut peer = Workbook::open_collaborative(&bytes, 852).unwrap();
    let catch_up = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&catch_up, CalculationOptions::default())
        .expect("a peer must be able to take what this replica publishes");
    assert_eq!(peer.model(), local.model(), "the replicas must agree");
    assert_ne!(moved, before);
}

/// Whether an anchor is exempt from the resolvability check must be a question
/// every replica answers the same way. Judging it against the replica's own
/// current projection makes the answer depend on local editing history: a
/// grandfathered source anchor reads as "unchanged" on a replica that still
/// holds it and as "newly introduced" on one that has moved on, so an undo
/// restoring it is accepted by one peer and refused by the other, and they
/// never converge again.
#[test]
fn an_undo_restoring_a_grandfathered_anchor_is_not_judged_by_local_history() {
    let bytes = unresolvable_anchor_fixture();
    let mut local = Workbook::open_collaborative(&bytes, 860).unwrap();
    let mut peer = Workbook::open_collaborative(&bytes, 861).unwrap();
    let opened = local.model().sheets[0].charts[0].anchor;

    // a repin away from it is ordinary and accepted
    let ChartAnchor::TwoCell { from, to, edit_as } = opened else {
        panic!("two-cell anchor");
    };
    let ordinary = ChartAnchor::TwoCell {
        from: to,
        to: from,
        edit_as,
    };
    local
        .apply_ops(
            vec![Op::SetChartAnchor {
                sheet: SheetId(0),
                frame: "xl/drawings/drawing1.xml#0".to_owned(),
                part: "xl/charts/chart1.xml".to_owned(),
                from: opened,
                to: ordinary,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let moved = local.model().sheets[0].charts[0].anchor;
    assert_ne!(moved, opened);
    let update = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&update, CalculationOptions::default())
        .unwrap();
    assert_eq!(peer.model().sheets[0].charts[0].anchor, moved);

    // undo puts the grandfathered anchor back and tells the peer
    assert!(local.undo(CalculationOptions::default()).unwrap().applied);
    assert_eq!(local.model().sheets[0].charts[0].anchor, opened);
    let undo_update = local
        .encode_diff_v1(&peer.encode_state_vector_v1())
        .unwrap();
    peer.apply_update_v1(&undo_update, CalculationOptions::default())
        .expect("a peer must accept an undo restoring the anchor it opened with");
    assert_eq!(
        peer.model().sheets[0].charts[0].anchor,
        local.model().sheets[0].charts[0].anchor,
        "the replicas must not part ways over a grandfathered anchor"
    );
}

/// One chart part may back two anchors in a drawing, so naming the part alone
/// cannot say which to repin. The op carries a frame and the anchor it saw
/// there, so the sibling sharing that part is left where it is, and a repin
/// whose recorded anchor has moved on is refused rather than applied blind.
#[test]
fn a_repin_moves_only_the_frame_it_names_when_a_part_backs_two() {
    let bytes = two_anchor_fixture();
    let mut workbook = Workbook::open(&bytes).unwrap();
    let charts = &workbook.model().sheets[0].charts;
    assert_eq!(charts.len(), 2);
    assert_eq!(
        charts[0].part, charts[1].part,
        "the fixture must share a part"
    );
    let (first, second) = (charts[0].anchor, charts[1].anchor);
    let ChartAnchor::TwoCell { from, to, edit_as } = first else {
        panic!("two-cell anchor");
    };
    let moved = ChartAnchor::TwoCell {
        from: AnchorCell {
            col: from.col + 1,
            ..from
        },
        to: AnchorCell {
            col: to.col + 1,
            ..to
        },
        edit_as,
    };
    let repin = |anchor_index: usize, from: ChartAnchor, to: ChartAnchor| Op::SetChartAnchor {
        sheet: SheetId(0),
        frame: format!("xl/drawings/drawing1.xml#{anchor_index}"),
        part: "xl/charts/chart1.xml".to_owned(),
        from,
        to,
    };

    assert!(
        workbook
            .apply_ops(vec![repin(0, first, moved)], CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(workbook.model().sheets[0].charts[0].anchor, moved);
    assert_eq!(
        workbook.model().sheets[0].charts[1].anchor,
        second,
        "the frame sharing the part must stay where it was"
    );

    let error = workbook
        .apply_ops(vec![repin(0, first, moved)], CalculationOptions::default())
        .expect_err("a repin whose recorded anchor has moved on must be refused");
    assert!(
        matches!(&error, Error::ChartFrameShifted { frame } if frame.ends_with("#0")),
        "{error:?}"
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

#[test]
fn the_demo_showcase_workbook_renders_its_chart() {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../apps/demo/public/showcase.xlsx"),
    )
    .unwrap();
    let workbook = Workbook::open(&bytes).unwrap();
    let list = workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 1400.0,
            height: 700.0,
        })
        .unwrap();
    let json = serde_json::to_string(&list).unwrap();
    let list: serde_json::Value = serde_json::from_str(&json).unwrap();
    let commands = list["commands"].as_array().expect("display list commands");
    let beyond_data = commands
        .iter()
        .filter(|command| command["x"].as_f64().unwrap_or(0.0) > 520.0)
        .count();
    assert!(
        beyond_data >= 8,
        "the dashboard chart should paint beyond the A1:F11 data region, saw {beyond_data} commands"
    );
    assert!(
        json.contains("4472C4"),
        "the revenue series colour should appear in the frame"
    );
}

/// A package whose only chart came from another producer: this crate models
/// its anchor and references but owns none of its markup.
fn imported_chart_xlsx(categories: &str, values: &str) -> Vec<u8> {
    let chart = format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>{categories}</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>{values}</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt><c:pt idx="2"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#
    );
    let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>12</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>7</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Q3</t></is></c><c r="B4"><v>9</v></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>"#;
    let parts = vec![
        (
            "xl/workbook.xml".to_owned(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/worksheets/sheet1.xml".to_owned(), worksheet.as_bytes().to_vec()),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "xl/drawings/drawing1.xml".to_owned(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Imported"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/charts/chart1.xml".to_owned(), chart.into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn chart_viewport() -> Viewport {
    Viewport {
        x: 0.0,
        y: 0.0,
        width: 900.0,
        height: 600.0,
    }
}

/// Everything the display list draws for a sheet except its cell text, so a
/// comparison sees the chart rather than the edited cell's own label.
fn plotted(workbook: &Workbook) -> Vec<DrawCmd> {
    workbook
        .display_list_for(SheetId(0), &chart_viewport())
        .unwrap()
        .commands
        .into_iter()
        .filter(|command| !matches!(command, DrawCmd::Text { .. }))
        .collect()
}

fn chart_part(saved: &[u8]) -> Vec<u8> {
    ooxml_opc::unzip_parts(saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "xl/charts/chart1.xml")
        .map(|(_, bytes)| bytes)
        .expect("the chart part survives a save")
}

/// A viewport holding the chart and nothing else. The whole-sheet viewport
/// also contains the edited cell, whose own repaint would make a pixel
/// comparison pass with the projection switched off entirely.
fn chart_only_viewport(workbook: &Workbook) -> Viewport {
    let sheet = workbook.model().sheet(SheetId(0)).unwrap();
    let geometry = GridGeometry::new(sheet);
    Viewport {
        x: geometry.col_x(3),
        y: 0.0,
        width: 520.0,
        height: 320.0,
    }
}

fn bump_b2(workbook: &mut Workbook) {
    workbook
        .edit_cells(
            SheetId(0),
            &[CellInput {
                cell: cell("B2"),
                input: "100".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

/// The second P1: an imported chart replayed its stored cache, so an ordinary
/// cell edit never reached it. It must now follow the cells in both backends,
/// while the bytes a save writes back stay exactly what was imported.
#[test]
#[cfg(feature = "raster")]
fn an_imported_chart_follows_a_cell_edit_without_a_save() {
    let source = imported_chart_xlsx("Data!$A$2:$A$4", "Data!$B$2:$B$4");
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(workbook.model().sheets[0].charts.len(), 1);

    let plot = chart_only_viewport(&workbook);
    let before = plotted(&workbook);
    let before_png = workbook.render_png_for(SheetId(0), &plot).unwrap();
    bump_b2(&mut workbook);
    let after_png = workbook.render_png_for(SheetId(0), &plot).unwrap();

    assert_ne!(before, plotted(&workbook));
    assert_ne!(before_png.bytes, after_png.bytes);

    // rendering live never rewrites the part: only the formulas moving does.
    assert_eq!(chart_part(&workbook.save().unwrap()), chart_part(&source));
}

/// A reference this crate cannot resolve safely keeps the values the file was
/// authored with. A wrong number is worse than a stale one.
#[test]
fn an_imported_chart_keeps_a_cache_it_cannot_resolve_safely() {
    for (categories, values, reason) in [
        ("Data!$A$2:$A$4", "(Data!$B$2:$B$3,Data!$B$4)", "a union"),
        ("Data!$A$2:$A$4", "Revenues", "a defined name"),
        ("Data!$A$2:$A$4", "[1]Data!$B$2:$B$4", "an external book"),
        ("Data!$A$2:$A$4", "Data!$A$2:$B$4", "a two-dimensional area"),
        (
            "Data!$A$2:$A$4",
            "Data!$D$2:$D$4",
            "cells the grid reads as empty",
        ),
    ] {
        let mut workbook = Workbook::open(&imported_chart_xlsx(categories, values)).unwrap();
        let before = plotted(&workbook);
        bump_b2(&mut workbook);
        assert_eq!(before, plotted(&workbook), "{reason} must keep its cache");
    }
}
