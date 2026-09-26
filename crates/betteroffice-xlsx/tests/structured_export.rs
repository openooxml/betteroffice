use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{
    CalculationOptions, CalculationRequest, CellRange, CellRef, CellState, CellValue, EditHistory,
    EditOperation, EditRequest, EditSource, EditStep, Error, Op, ProposalEditInput,
    ProposalRequest, RangeAddress, RangeTarget, ReadRequest, SheetId, Workbook, XlsxAnchor,
    XlsxExportCell, XlsxExportDiagnosticCode, XlsxExportFailureCode, XlsxExportOptions,
    XlsxExportScope, XlsxExportValue, XlsxFormulaResult, XlsxMarkdownOptions, XlsxObjectKind,
    XlsxStructuredContent, export_xlsx_markdown, export_xlsx_structured,
    export_xlsx_structured_json, render_xlsx_markdown, render_xlsx_markdown_json,
};
use serde_json::{Value, json};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        let mode = if target.starts_with("https:") {
            r#" TargetMode="External""#
        } else {
            ""
        };
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"{mode}/>"#
        ));
    }
    xml + "</Relationships>"
}

fn worksheet(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}">{body}</worksheet>"#
    )
}

fn drawing(anchors: &str) -> String {
    format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{DOC_RELS}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">{anchors}</xdr:wsDr>"#
    )
}

fn marker(tag: &str, col: u32, row: u32) -> String {
    format!(
        "<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{tag}>"
    )
}

const PICTURE: &str = r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo" descr="Company &lt;logo&gt;"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/>"#;

fn package(parts: Vec<(&str, String)>) -> Vec<u8> {
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .chain([(
            "xl/media/image1.png".to_owned(),
            vec![0x89, b'P', b'N', b'G', 0, 1, 2, 3],
        )])
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn content_types() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_owned()
}

/// Data holds typed, formatted, formula, rich-text, merged, hidden and distant cells, a
/// table, a hyperlink, comments, conditional formatting, a picture and an unreadable
/// chart. Summary holds a volatile formula. Secret is hidden, Ghost very hidden, Odd has an
/// unknown state and Dialog is a dialog sheet.
fn fixture() -> Vec<u8> {
    let workbook = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/><sheet name="Secret" sheetId="3" state="hidden" r:id="rId3"/><sheet name="Ghost" sheetId="4" state="veryHidden" r:id="rId4"/><sheet name="Odd" sheetId="5" state="sideways" r:id="rId5"/><sheet name="Dialog" sheetId="6" r:id="rId6"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$D$3</definedName><definedName name="Total">Data!$D$2:$D$3</definedName><definedName name="Local" localSheetId="1">Summary!$A$1</definedName><definedName name="Classified" localSheetId="2">Secret!$A$1</definedName></definedNames></workbook>"#
    );
    let data = worksheet(concat!(
        r#"<cols><col min="5" max="5" width="0" customWidth="1" hidden="1"/></cols><sheetData>"#,
        r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="inlineStr"><is><t>Price</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c><c r="E1" t="inlineStr"><is><t>secret column</t></is></c></row>"#,
        r#"<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>3</v></c><c r="C2" s="1"><v>1.5</v></c><c r="D2"><f>B2*C2</f><v>4.5</v></c><c r="F2" s="2"><v>0.5</v></c><c r="G2" s="3"><v>45000</v></c></row>"#,
        r#"<row r="3"><c r="A3" t="inlineStr"><is><t>Pear &lt;script&gt;|*x* &amp; [y](z)</t></is></c><c r="B3"><v>2</v></c><c r="D3"><f>B3*C3</f></c><c r="H3" s="1"/></row>"#,
        r#"<row r="4" hidden="1"><c r="A4" t="inlineStr"><is><t>secret row</t></is></c></row>"#,
        r#"<row r="6"><c r="A6" t="inlineStr"><is><t>merged</t></is></c><c r="B6" t="inlineStr"><is><t>covered</t></is></c></row>"#,
        r#"<row r="7"><c r="A7" t="inlineStr"><is><t>link</t></is></c></row>"#,
        r#"<row r="1000"><c r="Z1000" t="inlineStr"><is><t>far</t></is></c></row>"#,
        r#"</sheetData><mergeCells count="1"><mergeCell ref="A6:B6"/></mergeCells>"#,
        r#"<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>"#,
        r#"<hyperlinks><hyperlink ref="A7" r:id="rIdLink"/></hyperlinks><drawing r:id="rIdDrawing"/><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>"#,
    ));
    let summary = worksheet(
        r#"<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/><sheetData><row r="1" ht="15" customHeight="1"><c r="A1"><f>SUM(Data!D2:D3)</f><v>4.5</v></c></row><row r="2" ht="15" customHeight="1"><c r="A2"><f>NOW()</f><v>45000.5</v></c></row><row r="3" ht="15" customHeight="1"><c r="A3" t="inlineStr"><is><t>a|b &lt;i&gt;</t></is></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>zero height</t></is></c></row></sheetData>"#,
    );
    let text_sheet = |text: &str| {
        worksheet(&format!(
            r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row></sheetData>"#
        ))
    };
    let chart_frame = format!(
        r#"<xdr:oneCellAnchor>{}<xdr:ext cx="100" cy="100"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Broken chart" descr="Sales by month"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>"#,
        marker("from", 10, 10)
    );
    package(vec![
        ("[Content_Types].xml", content_types()),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        ("xl/workbook.xml", workbook),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
                ("rId3", "worksheet", "worksheets/sheet3.xml"),
                ("rId4", "worksheet", "worksheets/sheet4.xml"),
                ("rId5", "worksheet", "worksheets/sheet5.xml"),
                ("rId6", "dialogsheet", "dialogsheets/sheet1.xml"),
                ("rId7", "styles", "styles.xml"),
                ("rId8", "sharedStrings", "sharedStrings.xml"),
            ]),
        ),
        (
            "xl/styles.xml",
            format!(
                r##"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}"><numFmts count="1"><numFmt numFmtId="164" formatCode="# ?/?"/></numFmts><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="2" applyNumberFormat="1"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>"##
            ),
        ),
        (
            "xl/sharedStrings.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{MAIN}" count="3" uniqueCount="3"><si><t>Name</t></si><si><r><rPr><b/></rPr><t>App</t></r><r><t>le</t></r></si><si><t>Qty</t></si></sst>"#
            ),
        ),
        ("xl/worksheets/sheet1.xml", data),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[
                ("rIdLink", "hyperlink", "https://example.com/?a=1&amp;b=2"),
                ("rIdDrawing", "drawing", "../drawings/drawing1.xml"),
                ("rIdTable", "table", "../tables/table1.xml"),
                ("rIdComments", "comments", "../comments1.xml"),
            ]),
        ),
        (
            "xl/tables/table1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="{MAIN}" id="1" name="Sales" displayName="Sales" ref="A1:D3"><tableColumns count="4"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Price"/><tableColumn id="4" name="Total"/></tableColumns></table>"#
            ),
        ),
        (
            "xl/drawings/drawing1.xml",
            drawing(&format!(
                "<xdr:twoCellAnchor>{}{}{PICTURE}</xdr:twoCellAnchor>{chart_frame}",
                marker("from", 7, 1),
                marker("to", 9, 5)
            )),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            relationships(&[
                ("rIdImage", "image", "../media/image1.png"),
                ("rIdChart", "chart", "../charts/chart1.xml"),
            ]),
        ),
        (
            "xl/comments1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><comments xmlns="{MAIN}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>one</t></text></comment><comment ref="B1" authorId="0"><text><t>two</t></text></comment></commentList></comments>"#
            ),
        ),
        ("xl/worksheets/sheet2.xml", summary),
        ("xl/worksheets/sheet3.xml", text_sheet("classified")),
        ("xl/worksheets/sheet4.xml", text_sheet("boo")),
        ("xl/worksheets/sheet5.xml", text_sheet("odd")),
        (
            "xl/dialogsheets/sheet1.xml",
            format!(r#"<?xml version="1.0" encoding="UTF-8"?><dialogsheet xmlns="{MAIN}"/>"#),
        ),
    ])
}

/// One sheet with cells and a picture, which structural edits may move.
fn picture_fixture() -> Vec<u8> {
    package(vec![
        ("[Content_Types].xml", content_types()),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><workbookPr date1904="1"/><sheets><sheet name="Pics" sheetId="1" r:id="rId1"/></sheets></workbook>"#
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
                r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>"#
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            worksheet(
                r#"<sheetData><row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1"><f>A1+1</f><v>1</v></c></row><row r="2"><c r="A2"><f>B2</f><v>0</v></c><c r="B2"><f>A2</f><v>0</v></c></row></sheetData><drawing r:id="rIdDrawing"/>"#,
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[("rIdDrawing", "drawing", "../drawings/drawing1.xml")]),
        ),
        (
            "xl/drawings/drawing1.xml",
            drawing(&format!(
                "<xdr:twoCellAnchor>{}{}{PICTURE}</xdr:twoCellAnchor>",
                marker("from", 3, 3),
                marker("to", 5, 6)
            )),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            relationships(&[("rIdImage", "image", "../media/image1.png")]),
        ),
    ])
}

/// A formula stored without a result, rich inline text, and two merges sharing a column.
fn facts_fixture() -> Vec<u8> {
    package(vec![
        ("[Content_Types].xml", content_types()),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Facts" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[("rId1", "worksheet", "worksheets/sheet1.xml")]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            worksheet(concat!(
                r#"<sheetData><row r="1"><c r="A1" t="b"><f>FALSE()</f></c><c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>Bo</t></r><r><t>ld</t></r></is></c></row>"#,
                r#"<row r="3"><c r="A3" t="inlineStr"><is><t>m1</t></is></c><c r="B3" t="inlineStr"><is><t>c1</t></is></c></row>"#,
                r#"<row r="5"><c r="A5" t="inlineStr"><is><t>m2</t></is></c><c r="B5" t="inlineStr"><is><t>c2</t></is></c></row></sheetData>"#,
                r#"<mergeCells count="2"><mergeCell ref="A3:B3"/><mergeCell ref="A5:B5"/></mergeCells>"#,
            )),
        ),
    ])
}

/// One sheet with `sheet_xml`, `rels` beside it and `extra` parts.
fn one_sheet_package(
    sheet_xml: &str,
    rels: &[(&str, &str, &str)],
    extra: Vec<(&str, String)>,
) -> Vec<u8> {
    let mut parts = vec![
        ("[Content_Types].xml", content_types()),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Only" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[("rId1", "worksheet", "worksheets/sheet1.xml")]),
        ),
        ("xl/worksheets/sheet1.xml", worksheet(sheet_xml)),
        ("xl/worksheets/_rels/sheet1.xml.rels", relationships(rels)),
    ];
    parts.extend(extra);
    package(parts)
}

/// Links and a table wholly inside hidden row 2 or hidden column C, and a link reaching
/// past row 2.
fn hidden_links_fixture() -> Vec<u8> {
    one_sheet_package(
        concat!(
            r#"<cols><col min="3" max="3" width="0" customWidth="1" hidden="1"/></cols><sheetData>"#,
            r#"<row r="1"><c r="A1" t="inlineStr"><is><t>shown</t></is></c><c r="C1" t="inlineStr"><is><t>Code</t></is></c></row>"#,
            r#"<row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>secret row</t></is></c></row>"#,
            r#"<row r="3"><c r="A3" t="inlineStr"><is><t>after</t></is></c></row></sheetData>"#,
            r#"<hyperlinks><hyperlink ref="A2" location="Only!A1" display="row secret"/><hyperlink ref="C1" location="Only!A1" display="column secret"/>"#,
            r#"<hyperlink ref="A2:A3" location="Only!A1" display="reaches out"/><hyperlink ref="A1" location="Only!A3" display="shown link"/></hyperlinks>"#,
            r#"<tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>"#,
        ),
        &[("rIdTable", "table", "../tables/table1.xml")],
        vec![(
            "xl/tables/table1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><table xmlns="{MAIN}" id="1" name="Codes" displayName="Codes" ref="C1:C3"><tableColumns count="1"><tableColumn id="1" name="Code"/></tableColumns></table>"#
            ),
        )],
    )
}

/// Formulas the engine cannot parse: one stored without a result, one with an empty one.
fn blank_results_fixture() -> Vec<u8> {
    one_sheet_package(
        r#"<sheetData><row r="1"><c r="A1"><f>SUM(</f></c><c r="B1"><f>SUM(</f><v></v></c></row></sheetData>"#,
        &[],
        Vec::new(),
    )
}

/// A picture drawing, a malformed drawing and one nested past the parser's depth cap.
fn limited_drawing_fixture() -> Vec<u8> {
    let deep = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/>{}{}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#,
        "<xdr:grpSp>".repeat(70),
        "</xdr:grpSp>".repeat(70)
    );
    package(vec![
        ("[Content_Types].xml", content_types()),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Pics" sheetId="1" r:id="rId1"/><sheet name="Later" sheetId="2" r:id="rId2"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "worksheet", "worksheets/sheet2.xml"),
            ]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            worksheet(
                r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rIdDrawing"/>"#,
            ),
        ),
        (
            "xl/worksheets/sheet2.xml",
            worksheet(r#"<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>"#),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[
                ("rIdDrawing", "drawing", "../drawings/drawing1.xml"),
                ("rIdBroken", "drawing", "../drawings/broken.xml"),
                ("rIdDeep", "drawing", "../drawings/deep.xml"),
            ]),
        ),
        (
            "xl/drawings/drawing1.xml",
            drawing(&format!(
                "<xdr:twoCellAnchor>{}{}{PICTURE}</xdr:twoCellAnchor>",
                marker("from", 3, 3),
                marker("to", 5, 6)
            )),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            relationships(&[("rIdImage", "image", "../media/image1.png")]),
        ),
        ("xl/drawings/broken.xml", "<xdr:wsDr".to_owned()),
        ("xl/drawings/deep.xml", deep),
    ])
}

fn options(value: Value) -> XlsxExportOptions {
    serde_json::from_value(value).unwrap()
}

fn export(bytes: &[u8], value: Value) -> XlsxStructuredContent {
    export_xlsx_structured(bytes, &options(value)).unwrap()
}

fn sheet_names(content: &XlsxStructuredContent) -> Vec<String> {
    content
        .sheets
        .iter()
        .map(|sheet| match &sheet.anchor {
            XlsxAnchor::Sheet { sheet } => sheet.name.clone(),
            _ => panic!("sheet anchor"),
        })
        .collect()
}

fn cell<'a>(content: &'a XlsxStructuredContent, sheet: usize, a1: &str) -> &'a XlsxExportCell {
    content.sheets[sheet]
        .cells
        .iter()
        .find(|cell| matches!(&cell.anchor, XlsxAnchor::Cell { a1: at, .. } if at == a1))
        .unwrap_or_else(|| panic!("no exported cell {a1}"))
}

fn addresses(content: &XlsxStructuredContent, sheet: usize) -> Vec<String> {
    content.sheets[sheet]
        .cells
        .iter()
        .map(|cell| match &cell.anchor {
            XlsxAnchor::Cell { a1, .. } => a1.clone(),
            _ => panic!("cell anchor"),
        })
        .collect()
}

fn codes(content: &XlsxStructuredContent) -> Vec<XlsxExportDiagnosticCode> {
    content
        .diagnostics
        .iter()
        .map(|diagnostic| diagnostic.code)
        .collect()
}

fn to_json(value: &impl serde::Serialize) -> Value {
    serde_json::to_value(value).unwrap()
}

#[test]
fn exports_sparse_cells_with_values_formulas_and_display_text() {
    let content = export(&fixture(), json!({}));
    let wire = to_json(&content);
    assert_eq!(wire["schemaVersion"], 1);
    assert_eq!(wire["anchorScope"], "snapshot");
    assert_eq!(wire["dateSystem"], "1900");
    assert_eq!(
        wire["calculation"],
        json!({"policy": "asStored", "freshness": "unverified"})
    );
    assert!(!content.truncated);
    assert_eq!(sheet_names(&content), ["Data", "Summary", "Dialog"]);

    let data = &wire["sheets"][0];
    assert_eq!(data["id"], "s0");
    assert_eq!(data["kind"], "worksheet");
    assert_eq!(data["visibility"], "visible");
    assert_eq!(data["usedRange"], "A1:Z1000");
    assert_eq!(data["selectedRange"], "A1:Z1000");
    assert_eq!(data["hiddenRows"], json!(["4:4"]));
    assert_eq!(data["hiddenColumns"], json!(["E:E"]));
    assert_eq!(data["source"]["part"], "xl/worksheets/sheet1.xml");
    assert_eq!(data["source"]["partSha256"].as_str().unwrap().len(), 64);
    assert_eq!(
        addresses(&content, 0),
        [
            "A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "F2", "G2", "A3", "B3", "D3", "H3",
            "A6", "B6", "A7", "Z1000"
        ]
    );
    assert_eq!(
        to_json(cell(&content, 0, "D2")),
        json!({
            "id": "s0!D2",
            "anchor": {"kind": "cell", "sheet": {"sheetId": "sheet:0", "index": 0, "name": "Data"}, "a1": "D2"},
            "value": {"kind": "number", "value": 4.5},
            "formula": "B2*C2",
            "displayText": "4.5",
            "numberFormat": "General",
            "formulaResult": "unverified",
            "merge": null,
        })
    );
    let styled = cell(&content, 0, "H3");
    assert_eq!(styled.value, XlsxExportValue::Empty);
    assert_eq!(
        (styled.display_text.as_str(), styled.number_format.as_str()),
        ("", "0.00")
    );
    assert_eq!(styled.formula_result, None);
    let missing = cell(&content, 0, "D3");
    assert_eq!(missing.value, XlsxExportValue::Empty);
    assert_eq!(missing.formula_result, Some(XlsxFormulaResult::Missing));
    assert_eq!(cell(&content, 0, "C2").display_text, "1.50");
    assert_eq!(cell(&content, 0, "C2").number_format, "0.00");
    assert_eq!(cell(&content, 0, "G2").display_text, "3/15/23");
    assert_eq!(
        cell(&content, 0, "G2").value,
        XlsxExportValue::Number { value: 45000.0 }
    );
    assert_eq!(cell(&content, 0, "A2").display_text, "Apple");
    assert_eq!(
        to_json(&cell(&content, 0, "B6").merge),
        json!({"range": "A6:B6", "origin": false})
    );
    assert_eq!(
        cell(&content, 0, "B6").value,
        XlsxExportValue::Text {
            value: "covered".to_owned()
        }
    );
    assert!(cell(&content, 0, "A6").merge.as_ref().unwrap().origin);

    assert_eq!(
        data["merges"],
        json!([{"id": "s0:m0", "anchor": {"kind": "range", "sheet": {"sheetId": "sheet:0", "index": 0, "name": "Data"}, "a1": "A6:B6"}, "clipped": false}])
    );
    assert_eq!(data["tables"][0]["name"], "Sales");
    assert_eq!(data["tables"][0]["anchor"]["a1"], "A1:D3");
    assert_eq!(
        data["tables"][0]["columns"],
        json!(["Name", "Qty", "Price", "Total"])
    );
    assert_eq!(data["tables"][0]["headerRows"], 1);
    assert_eq!(
        data["hyperlinks"][0]["externalTarget"],
        "https://example.com/?a=1&b=2"
    );
    assert_eq!(data["hyperlinks"][0]["anchor"]["a1"], "A7");

    let objects = &content.sheets[0].objects;
    assert_eq!(objects.len(), 2);
    assert_eq!(objects[0].kind, XlsxObjectKind::Picture);
    assert_eq!(objects[0].name.as_deref(), Some("Logo"));
    assert_eq!(objects[0].alt_text.as_deref(), Some("Company <logo>"));
    assert_eq!(objects[0].part.as_deref(), Some("xl/media/image1.png"));
    assert_eq!(
        to_json(&objects[0].anchor),
        json!({"kind": "range", "sheet": {"sheetId": "sheet:0", "index": 0, "name": "Data"}, "a1": "H2:J6"})
    );
    let provenance = objects[0].source.as_ref().unwrap();
    assert_eq!(provenance.part, "xl/drawings/drawing1.xml");
    assert_eq!(provenance.path, [0]);
    assert_eq!(objects[1].kind, XlsxObjectKind::Chart);
    assert_eq!(objects[1].alt_text.as_deref(), Some("Sales by month"));
    assert_eq!(objects[1].source.as_ref().unwrap().path, [1]);
    assert_eq!(to_json(&objects[1].anchor)["a1"], "K11");

    let summary = &content.sheets[1];
    assert_eq!(summary.hidden_rows, ["4:5"]);
    assert_eq!(summary.cells.len(), 3);
    assert_eq!(summary.cells[1].formula.as_deref(), Some("NOW()"));
    assert_eq!(
        summary.cells[1].value,
        XlsxExportValue::Number { value: 45000.5 }
    );
    let dialog = &wire["sheets"][2];
    assert_eq!(dialog["kind"], "dialogsheet");
    assert_eq!(dialog["cells"], json!([]));

    let names = content
        .defined_names
        .iter()
        .map(|name| name.name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(names, ["Total", "Local"]);
    assert_eq!(
        to_json(&content.defined_names[1].anchor),
        json!({"kind": "definedName", "name": "Local", "localSheet": {"sheetId": "sheet:1", "index": 1, "name": "Summary"}, "ordinal": 2})
    );

    let found = codes(&content);
    for code in [
        XlsxExportDiagnosticCode::HiddenContentExcluded,
        XlsxExportDiagnosticCode::VisibilityUnknown,
        XlsxExportDiagnosticCode::UnsupportedSheet,
        XlsxExportDiagnosticCode::CommentsOmitted,
        XlsxExportDiagnosticCode::UnsupportedContent,
        XlsxExportDiagnosticCode::UnreadableObject,
        XlsxExportDiagnosticCode::ObjectPlaceholder,
        XlsxExportDiagnosticCode::FormulaCacheUnverified,
        XlsxExportDiagnosticCode::FormulaResultMissing,
        XlsxExportDiagnosticCode::RichTextOmitted,
        XlsxExportDiagnosticCode::FormattingApproximate,
    ] {
        assert!(found.contains(&code), "{code:?} missing from {found:?}");
    }
    let diagnostic = |code| {
        content
            .diagnostics
            .iter()
            .find(|diagnostic| diagnostic.code == code)
            .unwrap()
    };
    assert_eq!(
        to_json(&diagnostic(XlsxExportDiagnosticCode::RichTextOmitted).anchor)["a1"],
        "A2"
    );
    assert_eq!(
        to_json(&diagnostic(XlsxExportDiagnosticCode::FormulaResultMissing).anchor)["a1"],
        "D3"
    );
    assert_eq!(
        to_json(&diagnostic(XlsxExportDiagnosticCode::FormattingApproximate).anchor)["a1"],
        "F2"
    );
    assert!(
        diagnostic(XlsxExportDiagnosticCode::CommentsOmitted)
            .message
            .ends_with("(2).")
    );
    assert!(
        diagnostic(XlsxExportDiagnosticCode::UnsupportedContent)
            .message
            .contains("conditional formatting")
    );
    let hidden_sheets = content
        .diagnostics
        .iter()
        .filter(|diagnostic| {
            diagnostic.code == XlsxExportDiagnosticCode::HiddenContentExcluded
                && matches!(diagnostic.anchor, Some(XlsxAnchor::Sheet { .. }))
        })
        .count();
    assert_eq!(
        hidden_sheets, 5,
        "Secret, Ghost, hidden rows on two sheets, hidden columns"
    );
    let serialized = serde_json::to_string(&content).unwrap();
    for secret in [
        "classified",
        "boo",
        "odd",
        "secret row",
        "secret column",
        "zero height",
    ] {
        assert!(!serialized.contains(secret), "{secret} leaked");
    }
}

#[test]
fn hidden_content_is_opt_in() {
    let content = export(
        &fixture(),
        json!({
            "includeHiddenSheets": true,
            "includeHiddenRows": true,
            "includeHiddenColumns": true,
            "includeHiddenNames": true,
        }),
    );
    assert_eq!(
        sheet_names(&content),
        ["Data", "Summary", "Secret", "Ghost", "Odd", "Dialog"]
    );
    let wire = to_json(&content);
    assert_eq!(wire["sheets"][2]["visibility"], "hidden");
    assert_eq!(wire["sheets"][3]["visibility"], "veryHidden");
    assert_eq!(wire["sheets"][4]["visibility"], "unknown");
    assert_eq!(cell(&content, 0, "A4").display_text, "secret row");
    assert_eq!(cell(&content, 0, "E1").display_text, "secret column");
    assert_eq!(wire["sheets"][0]["hiddenRows"], json!(["4:4"]));
    assert_eq!(
        wire["included"],
        json!({"hiddenSheets": true, "hiddenRows": true, "hiddenColumns": true, "definedNames": true, "hiddenNames": true})
    );
    assert_eq!(content.defined_names.len(), 4);
    assert!(content.defined_names[0].hidden);
    assert!(
        !codes(&content).contains(&XlsxExportDiagnosticCode::HiddenContentExcluded)
            && !codes(&content).contains(&XlsxExportDiagnosticCode::VisibilityUnknown)
    );

    let without_names = export(&fixture(), json!({"includeDefinedNames": false}));
    assert!(without_names.defined_names.is_empty());
    assert!(without_names.diagnostics.iter().any(|diagnostic| {
        diagnostic
            .message
            .starts_with("Defined names are excluded (4)")
    }));
}

#[test]
fn scope_selects_sheets_and_ranges() {
    let content = export(
        &fixture(),
        json!({"scope": [{"sheet": 1}, {"sheet": 0, "range": "b2:$D$3"}]}),
    );
    assert_eq!(sheet_names(&content), ["Data", "Summary"]);
    let data = &content.sheets[0];
    assert_eq!(data.selected_range.as_deref(), Some("B2:D3"));
    assert_eq!(data.used_range.as_deref(), Some("A1:Z1000"));
    assert_eq!(addresses(&content, 0), ["B2", "C2", "D2", "B3", "D3"]);
    assert!(data.merges.is_empty());
    assert!(data.tables[0].clipped);
    assert!(data.objects.is_empty());

    let merged = export(
        &fixture(),
        json!({"scope": [{"sheet": 0, "range": "B6:I7"}]}),
    );
    let data = &merged.sheets[0];
    assert!(data.merges[0].clipped);
    assert_eq!(addresses(&merged, 0), ["B6"]);
    assert_eq!(data.objects.len(), 1);
    assert_eq!(data.objects[0].kind, XlsxObjectKind::Picture);
    let rendered = render_xlsx_markdown(&merged, &XlsxMarkdownOptions::default()).unwrap();
    assert!(rendered.markdown.contains(
        r#"<tr><th>6</th><td rowspan="1" colspan="1" data-merge="A6:B6">[merged from A6]</td>"#
    ));
    assert!(!rendered.markdown.contains("covered"));
    let mut overlapping = merged.clone();
    let extra = overlapping.sheets[0].merges[0].clone();
    overlapping.sheets[0].merges.push(extra);
    assert!(matches!(
        render_xlsx_markdown(&overlapping, &XlsxMarkdownOptions::default()),
        Err(Error::InvalidRequest(message)) if message.contains("overlap")
    ));

    let workbook = Workbook::open(&fixture()).unwrap();
    let refusal = |value: Value| {
        workbook
            .export_structured(&options(value))
            .unwrap()
            .unwrap_err()
    };
    for (value, code, sheet) in [
        (
            json!({"scope": [{"sheet": 9}]}),
            XlsxExportFailureCode::InvalidScope,
            None,
        ),
        (
            json!({"scope": []}),
            XlsxExportFailureCode::InvalidScope,
            None,
        ),
        (
            json!({"scope": [{"sheet": 0}, {"sheet": 0, "range": "A1"}]}),
            XlsxExportFailureCode::InvalidScope,
            Some("Data"),
        ),
        (
            json!({"scope": [{"sheet": 1, "range": "D3:B2"}]}),
            XlsxExportFailureCode::InvalidScope,
            Some("Summary"),
        ),
        (
            json!({"scope": [{"sheet": 0, "range": "A1,B2"}]}),
            XlsxExportFailureCode::InvalidScope,
            Some("Data"),
        ),
        (
            json!({"scope": [{"sheet": 0, "range": "1:3"}]}),
            XlsxExportFailureCode::InvalidScope,
            Some("Data"),
        ),
        (
            json!({"scope": [{"sheet": 0, "range": "Data!A1"}]}),
            XlsxExportFailureCode::InvalidScope,
            Some("Data"),
        ),
        (
            json!({"maxCells": 0}),
            XlsxExportFailureCode::InvalidOptions,
            None,
        ),
        (
            json!({"maxCells": 1_000_001}),
            XlsxExportFailureCode::LimitExceeded,
            None,
        ),
        (
            json!({"maxBytes": 100}),
            XlsxExportFailureCode::LimitExceeded,
            None,
        ),
        (
            json!({"maxBytes": 16_777_217}),
            XlsxExportFailureCode::LimitExceeded,
            None,
        ),
    ] {
        let refused = refusal(value.clone());
        assert_eq!(refused.failure.code, code, "{value}");
        assert_eq!(refused.version, workbook.version());
        match sheet {
            Some(name) => assert!(
                matches!(&refused.failure.target, Some(XlsxAnchor::Sheet { sheet }) if sheet.name == name),
                "{value}"
            ),
            None => assert_eq!(refused.failure.target, None, "{value}"),
        }
    }
    let wire: Value = serde_json::from_str(
        &workbook
            .export_structured_json(r#"{"scope":[{"sheet":7}]}"#)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(wire["ok"], false);
    assert_eq!(wire["failure"]["target"], Value::Null);
    assert_eq!(wire["failure"]["code"], "invalid-scope");
    assert!(matches!(
        workbook.export_structured_json(r#"{"maxCells":-1}"#),
        Err(Error::InvalidRequest(_))
    ));
    assert!(matches!(
        workbook.export_structured_json(r#"{"unknown":true}"#),
        Err(Error::InvalidRequest(_))
    ));
    assert!(matches!(
        export_xlsx_structured(&fixture(), &options(json!({"maxCells": 0}))),
        Err(Error::InvalidRequest(_))
    ));
}

#[test]
fn limits_truncate_at_complete_records() {
    let bytes = fixture();
    let full = export(&bytes, json!({}));
    let capped = export(&bytes, json!({"maxCells": 3}));
    assert!(capped.truncated);
    assert!(capped.sheets[0].truncated);
    assert_eq!(addresses(&capped, 0), ["A1", "B1", "C1"]);
    assert_eq!(capped.sheets.len(), 1);
    let last = capped.diagnostics.last().unwrap();
    assert_eq!(last.code, XlsxExportDiagnosticCode::Truncated);
    assert!(last.message.contains("maxCells") && last.message.contains("C1"));
    assert_eq!(capped.sheets[0].merges, full.sheets[0].merges);

    let full_size = serde_json::to_vec(&full).unwrap().len();
    let mut previous_cells = 0;
    for max_bytes in (1_024..full_size + 64).step_by(97) {
        let content = match export_xlsx_structured(&bytes, &options(json!({"maxBytes": max_bytes})))
        {
            Ok(content) => content,
            Err(Error::InvalidRequest(message)) => {
                assert!(message.contains("maxBytes must be at least"), "{message}");
                continue;
            }
            Err(error) => panic!("{error}"),
        };
        let size = serde_json::to_vec(&content).unwrap().len();
        assert!(size <= max_bytes, "{size} > {max_bytes}");
        let cells = content
            .sheets
            .iter()
            .map(|sheet| sheet.cells.len())
            .sum::<usize>();
        assert!(cells >= previous_cells);
        previous_cells = cells;
        for (sheet, complete) in content.sheets.iter().zip(&full.sheets) {
            assert_eq!(sheet.cells, complete.cells[..sheet.cells.len()]);
            assert_eq!(
                sheet.truncated,
                sheet.cells.len() < complete.cells.len() || content.truncated && sheet.truncated
            );
        }
        if content.truncated {
            assert_eq!(
                content.diagnostics.last().unwrap().code,
                XlsxExportDiagnosticCode::Truncated
            );
        } else {
            assert_eq!(content, full);
        }
    }
    assert_eq!(previous_cells, 21);
}

#[test]
fn bytes_export_is_deterministic_and_never_recalculates() {
    let bytes = fixture();
    let first = export_xlsx_structured_json(&bytes, "{}").unwrap();
    let second = export_xlsx_structured_json(&bytes, "{}").unwrap();
    assert_eq!(first, second);
    let content: XlsxStructuredContent = serde_json::from_str(&first).unwrap();
    let read_only = Workbook::open_for_read(&bytes)
        .unwrap()
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    assert_eq!(
        XlsxStructuredContent {
            anchor_scope: content.anchor_scope,
            ..read_only
        },
        content
    );
    assert_eq!(cell(&content, 1, "A2").display_text, "45000.5");
    assert_eq!(
        cell(&content, 1, "A2").formula_result,
        Some(XlsxFormulaResult::Unverified)
    );

    let recalculated = Workbook::open_recalculated(
        &bytes,
        CalculationOptions {
            now_serial: Some(46000.0),
        },
    )
    .unwrap();
    let live = recalculated
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    assert_eq!(
        cell(&live, 1, "A2").value,
        XlsxExportValue::Number { value: 46000.0 }
    );
    assert_eq!(
        cell(&live, 0, "D3").value,
        XlsxExportValue::Number { value: 0.0 }
    );
    assert_eq!(
        cell(&live, 0, "D3").formula_result,
        Some(XlsxFormulaResult::Uncertain)
    );
}

#[test]
fn live_export_reads_committed_state_read_only() {
    let mut workbook = Workbook::open(&fixture()).unwrap();
    let events = Arc::new(Mutex::new(0));
    let seen = Arc::clone(&events);
    let _subscription = workbook
        .observe_update_v1(move |_| *seen.lock().unwrap() += 1)
        .unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("B3").unwrap(),
            "4",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    workbook
        .propose(
            ProposalRequest {
                agent_id: "agent".to_owned(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(0),
                    cell: CellRef::parse_a1("B2").unwrap(),
                    input: "9".to_owned(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.set_active_sheet(SheetId(1)).unwrap();
    *events.lock().unwrap() = 0;
    let version = workbook.version();
    let proposals = workbook.proposals().to_vec();
    assert!(workbook.can_redo());
    let authority = workbook.encode_state_as_update_v1();
    let history = workbook.history_state();
    let saved = workbook.save().unwrap();
    let model = workbook.model().clone();
    let calculation = workbook.last_calculation().clone();

    let structured = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    assert_eq!(structured.version, version);
    let wire = to_json(&structured.content);
    assert_eq!(wire["anchorScope"], "session");
    let markdown = workbook
        .export_markdown(
            &XlsxExportOptions::default(),
            &XlsxMarkdownOptions::default(),
        )
        .unwrap()
        .unwrap();
    assert_eq!(markdown.version, version);
    let json: Value = serde_json::from_str(
        &workbook
            .export_markdown_json("{}", r#"{"maxRows":5}"#)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(json["ok"], true);
    assert_eq!(json["version"], version.as_str());

    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.encode_state_as_update_v1(), authority);
    assert_eq!(workbook.history_state(), history);
    assert_eq!(workbook.save().unwrap(), saved);
    assert_eq!(workbook.model(), &model);
    assert_eq!(workbook.last_calculation(), &calculation);
    assert_eq!(workbook.active_sheet(), SheetId(1));
    assert_eq!(workbook.proposals(), proposals.as_slice());
    assert!(workbook.can_redo());
    assert_eq!(*events.lock().unwrap(), 0);

    workbook.redo(CalculationOptions::default()).unwrap();
    let redone = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    assert_ne!(redone.version, version);
    assert_eq!(
        cell(&redone.content, 0, "B3").value,
        XlsxExportValue::Number { value: 4.0 }
    );
    workbook.undo(CalculationOptions::default()).unwrap();

    workbook
        .apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::parse_a1("C3").unwrap(),
                cell: CellState {
                    value: CellValue::Number { value: 10.0 },
                    formula: None,
                    style: None,
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let edited = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    assert_ne!(edited.version, version);
    assert_eq!(
        cell(&edited.content, 0, "D3").value,
        XlsxExportValue::Number { value: 20.0 }
    );
    workbook.undo(CalculationOptions::default()).unwrap();
    let undone = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    assert!(
        undone.content.sheets[0]
            .cells
            .iter()
            .all(|cell| !matches!(&cell.anchor, XlsxAnchor::Cell { a1, .. } if a1 == "C3"))
    );
}

fn batch_target(anchor: &XlsxAnchor) -> RangeTarget {
    let XlsxAnchor::Cell { sheet, a1 } = anchor else {
        panic!("cell anchor");
    };
    RangeTarget {
        sheet_id: sheet.sheet_id.clone(),
        range: RangeAddress::A1 { a1: a1.clone() },
    }
}

#[test]
fn export_anchors_are_batch_targets() {
    let bytes = fixture();
    let snapshot = export(&bytes, json!({}));
    assert_eq!(
        batch_target(&cell(&snapshot, 1, "A3").anchor).sheet_id,
        "sheet:1"
    );
    let mut shifted = Workbook::open(&bytes).unwrap();
    shifted
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "Front".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    for (mut workbook, summary) in [
        (shifted, 2),
        (Workbook::open_collaborative(&bytes, 31).unwrap(), 1),
    ] {
        let before = workbook
            .export_structured(&XlsxExportOptions::default())
            .unwrap()
            .unwrap();
        assert_eq!(sheet_names(&before.content)[summary], "Summary");
        let target = batch_target(&cell(&before.content, summary, "A3").anchor);
        let catalog = workbook
            .read_cells(&ReadRequest::default())
            .unwrap()
            .unwrap();
        assert_eq!(target.sheet_id, catalog.sheets[summary].sheet_id);
        let applied = workbook
            .apply_edits(&EditRequest {
                expect_version: before.version,
                source: EditSource::Host,
                history: EditHistory::Separate,
                calculation: CalculationRequest::default(),
                steps: vec![EditStep::new(EditOperation::SetCellInputs {
                    target,
                    inputs: vec![vec!["edited".to_owned()]],
                })],
            })
            .unwrap()
            .unwrap();
        assert!(applied.applied);
        let after = workbook
            .export_structured(&XlsxExportOptions::default())
            .unwrap()
            .unwrap();
        assert_eq!(
            cell(&after.content, summary, "A3").value,
            XlsxExportValue::Text {
                value: "edited".to_owned()
            }
        );
        assert_eq!(
            cell(&after.content, summary - 1, "A3").value,
            cell(&before.content, summary - 1, "A3").value
        );
    }
}

#[test]
fn structural_edits_drop_object_anchors_they_cannot_verify() {
    let mut workbook = Workbook::open(&picture_fixture()).unwrap();
    let before = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    assert_eq!(to_json(&before)["dateSystem"], "1904");
    assert_eq!(cell(&before, 0, "A1").display_text, "1/1/04");
    assert_eq!(to_json(&before.sheets[0].objects[0].anchor)["a1"], "D4:F7");
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 20,
                count: 3,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let below = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    assert_eq!(to_json(&below.sheets[0].objects[0].anchor)["a1"], "D4:F7");
    assert!(!codes(&below).contains(&XlsxExportDiagnosticCode::ProvenanceUnavailable));
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let after = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    assert_eq!(cell(&after, 0, "B3").formula.as_deref(), Some("A3+1"));
    let picture = &after.sheets[0].objects[0];
    assert_eq!(picture.kind, XlsxObjectKind::Picture);
    assert_eq!(picture.anchor, None);
    assert!(picture.source.is_some());
    assert!(codes(&after).contains(&XlsxExportDiagnosticCode::ProvenanceUnavailable));
}

#[test]
fn sheets_without_source_metadata_are_unknown_unless_created_here() {
    let model = || {
        let mut model = betteroffice_xlsx::WorkbookModel::default();
        let mut sheet = betteroffice_xlsx::Sheet::new("Plain");
        sheet.set_cell(
            CellRef::parse_a1("B2").unwrap(),
            betteroffice_xlsx::Cell {
                value: CellValue::Number { value: 2.0 },
                formula: Some("1+1".to_owned()),
                style: None,
            },
        );
        model.sheets.push(sheet);
        model
    };
    let exported = |workbook: &Workbook, value: Value| {
        workbook
            .export_structured(&options(value))
            .unwrap()
            .unwrap()
            .content
    };
    for workbook in [
        Workbook::from_model(model()).unwrap(),
        Workbook::from_model_collaborative(model(), 7).unwrap(),
    ] {
        let hidden = exported(&workbook, json!({}));
        assert!(hidden.sheets.is_empty());
        let diagnostic = &hidden.diagnostics[0];
        assert_eq!(diagnostic.code, XlsxExportDiagnosticCode::VisibilityUnknown);
        assert_eq!(to_json(&diagnostic.anchor)["sheet"]["name"], "Plain");
        let content = exported(&workbook, json!({"includeHiddenSheets": true}));
        let wire = to_json(&content.sheets[0]);
        assert_eq!(wire["visibility"], "unknown");
        assert_eq!(wire["source"], Value::Null);
        assert_eq!(wire["objects"], json!([]));
        assert_eq!(
            cell(&content, 0, "B2").formula_result,
            Some(XlsxFormulaResult::Unverified)
        );
    }

    let mut workbook = Workbook::from_model(model()).unwrap();
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 1,
                name: "Fresh".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let content = exported(&workbook, json!({}));
    assert_eq!(sheet_names(&content), ["Fresh"]);
    assert_eq!(to_json(&content.sheets[0])["visibility"], "visible");

    let mut opened = Workbook::open(&fixture()).unwrap();
    opened
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "Added".to_owned(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let content = exported(&opened, json!({}));
    assert_eq!(
        sheet_names(&content),
        ["Added", "Data", "Summary", "Dialog"]
    );
    let replica = Workbook::open_collaborative(&fixture(), 9).unwrap();
    assert_eq!(
        sheet_names(&exported(&replica, json!({}))),
        ["Data", "Summary", "Dialog"]
    );
}

#[test]
fn missing_caches_rich_inline_text_and_later_merges_are_reported() {
    let bytes = facts_fixture();
    let content = export(&bytes, json!({}));
    let uncached = cell(&content, 0, "A1");
    assert_eq!(uncached.value, XlsxExportValue::Bool { value: false });
    assert_eq!(uncached.formula_result, Some(XlsxFormulaResult::Missing));
    let diagnostic = |code| {
        content
            .diagnostics
            .iter()
            .find(|diagnostic| diagnostic.code == code)
            .unwrap_or_else(|| panic!("no {code:?}"))
    };
    assert_eq!(
        to_json(&diagnostic(XlsxExportDiagnosticCode::FormulaResultMissing).anchor)["a1"],
        "A1"
    );
    assert_eq!(
        to_json(&diagnostic(XlsxExportDiagnosticCode::RichTextOmitted).anchor)["a1"],
        "B1"
    );
    for (a1, range, origin) in [
        ("A3", "A3:B3", true),
        ("B3", "A3:B3", false),
        ("A5", "A5:B5", true),
        ("B5", "A5:B5", false),
    ] {
        assert_eq!(
            to_json(&cell(&content, 0, a1).merge),
            json!({"range": range, "origin": origin}),
            "{a1}"
        );
    }

    let mut workbook = Workbook::open(&bytes).unwrap();
    let opened = workbook.version();
    let live = |workbook: &Workbook| {
        workbook
            .export_structured(&XlsxExportOptions::default())
            .unwrap()
            .unwrap()
    };
    assert_eq!(
        cell(&live(&workbook).content, 0, "A1").formula_result,
        Some(XlsxFormulaResult::Missing)
    );
    assert!(
        workbook
            .recalculate_all(CalculationOptions::default())
            .changed
            .is_empty()
    );
    let recalculated = live(&workbook);
    assert_ne!(recalculated.version, opened);
    assert_eq!(
        cell(&recalculated.content, 0, "A1").formula_result,
        Some(XlsxFormulaResult::Uncertain)
    );
    workbook.recalculate_all(CalculationOptions::default());
    assert_eq!(workbook.version(), recalculated.version);
}

#[test]
fn a_recalculation_that_only_finds_a_cycle_moves_the_version() {
    let mut workbook = Workbook::open(&picture_fixture()).unwrap();
    let opened = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    assert_eq!(
        cell(&opened.content, 0, "A2").formula_result,
        Some(XlsxFormulaResult::Unverified)
    );
    workbook.recalculate_all(CalculationOptions::default());
    let recalculated = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap();
    assert_eq!(
        cell(&recalculated.content, 0, "A2").value,
        cell(&opened.content, 0, "A2").value
    );
    assert_eq!(
        cell(&recalculated.content, 0, "A2").formula_result,
        Some(XlsxFormulaResult::Cycle)
    );
    assert_ne!(recalculated.version, opened.version);
}

#[test]
fn an_empty_formula_value_is_missing_only_where_the_file_stored_none() {
    let bytes = blank_results_fixture();
    let stored = export(&bytes, json!({}));
    assert_eq!(
        cell(&stored, 0, "A1").formula_result,
        Some(XlsxFormulaResult::Missing)
    );
    assert_eq!(
        cell(&stored, 0, "B1").formula_result,
        Some(XlsxFormulaResult::Unverified)
    );

    let mut workbook = Workbook::open(&bytes).unwrap();
    workbook.recalculate_all(CalculationOptions::default());
    let calculated = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    for a1 in ["A1", "B1"] {
        let exported = cell(&calculated, 0, a1);
        assert_eq!(exported.value, XlsxExportValue::Empty, "{a1}");
        assert_eq!(
            exported.formula_result,
            Some(XlsxFormulaResult::Uncertain),
            "{a1}"
        );
    }
}

#[test]
fn links_and_tables_wholly_in_excluded_rows_or_columns_are_left_out() {
    let bytes = hidden_links_fixture();
    let links = |content: &XlsxStructuredContent| {
        content.sheets[0]
            .hyperlinks
            .iter()
            .map(|link| link.display.clone().unwrap())
            .collect::<Vec<_>>()
    };
    let tables = |content: &XlsxStructuredContent| {
        content.sheets[0]
            .tables
            .iter()
            .map(|table| table.name.clone())
            .collect::<Vec<_>>()
    };

    let content = export(&bytes, json!({}));
    assert_eq!(links(&content), ["reaches out", "shown link"]);
    assert!(tables(&content).is_empty());
    let markdown = export_xlsx_markdown(
        &bytes,
        &XlsxExportOptions::default(),
        &XlsxMarkdownOptions::default(),
    )
    .unwrap()
    .markdown;
    for secret in ["row secret", "column secret", "Codes"] {
        assert!(!markdown.contains(secret), "{secret} in {markdown}");
    }

    let rows = export(&bytes, json!({"includeHiddenRows": true}));
    assert_eq!(links(&rows), ["row secret", "reaches out", "shown link"]);
    assert!(tables(&rows).is_empty());
    let both = export(
        &bytes,
        json!({"includeHiddenRows": true, "includeHiddenColumns": true}),
    );
    assert_eq!(
        links(&both),
        ["row secret", "column secret", "reaches out", "shown link"]
    );
    assert_eq!(tables(&both), ["Codes"]);

    let scoped = export(&bytes, json!({"scope": [{"sheet": 0, "range": "A2:B2"}]}));
    assert!(links(&scoped).is_empty());
}

#[test]
fn a_drawing_past_a_parser_cap_truncates_and_a_malformed_one_is_diagnosed() {
    let content = export(&limited_drawing_fixture(), json!({}));
    assert!(content.truncated);
    assert_eq!(sheet_names(&content), ["Pics"]);
    let sheet = &content.sheets[0];
    assert!(sheet.truncated);
    assert_eq!(sheet.objects.len(), 1);
    assert_eq!(sheet.objects[0].kind, XlsxObjectKind::Picture);
    assert!(sheet.cells.is_empty());
    let messages = content
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == XlsxExportDiagnosticCode::UnreadableObject)
        .map(|diagnostic| diagnostic.message.as_str())
        .collect::<Vec<_>>();
    assert!(
        messages[0].contains("xl/drawings/broken.xml"),
        "{messages:?}"
    );
    assert!(messages[1].contains("xl/drawings/deep.xml") && messages[1].contains("exceeds"));
    let last = content.diagnostics.last().unwrap();
    assert_eq!(last.code, XlsxExportDiagnosticCode::Truncated);
    assert!(
        last.message.contains("source-inspection"),
        "{}",
        last.message
    );
}

#[test]
fn markdown_keeps_text_inert_and_links_safe() {
    let mut model = betteroffice_xlsx::WorkbookModel::default();
    let mut sheet = betteroffice_xlsx::Sheet::new("Q1 <b>|#_`x`& {y}");
    let text = |value: &str| betteroffice_xlsx::Cell {
        value: CellValue::Text {
            value: value.to_owned(),
        },
        formula: None,
        style: None,
    };
    sheet.set_cell(
        CellRef::parse_a1("A1").unwrap(),
        text("line one\r\n\r\n<img src=x onerror=alert(1)>\u{2028}after"),
    );
    sheet.set_cell(
        CellRef::parse_a1("B1").unwrap(),
        text("a|b </a><script>alert(1)</script>"),
    );
    let vectors = [
        "javascript&#58;alert(1)",
        "java&#x09;script:alert(1)",
        "&#106;avascript:alert(1)",
        "&#106avascript:alert(1)",
        "JaVaScRiPt:alert(1)",
        &format!("{}javascript:alert(1)", " ".repeat(256)),
        "java\u{0}script:alert(1)",
        "\u{1}javascript:alert(1)",
        "jav\nascript:alert(1)",
        "%6Aavascript:alert(1)",
        "data:text/html,<script>alert(1)</script>",
        "java&colon;script&colon;alert(1)",
        "&unknown;javascript:alert(1)",
    ];
    for (row, target) in vectors.iter().enumerate() {
        sheet.hyperlinks.push(betteroffice_xlsx::Hyperlink {
            range: CellRange::new(
                CellRef::new(row as u32 + 2, 0),
                CellRef::new(row as u32 + 2, 0),
            ),
            external_target: Some((*target).to_owned()),
            location: None,
            tooltip: None,
            display: None,
        });
    }
    sheet.hyperlinks.push(betteroffice_xlsx::Hyperlink {
        range: CellRange::parse_a1("C1").unwrap(),
        external_target: Some("https://example.com/a b?x=1&y=(2)<3>".to_owned()),
        location: None,
        tooltip: None,
        display: Some("Docs](javascript:alert(1))".to_owned()),
    });
    sheet.hyperlinks.push(betteroffice_xlsx::Hyperlink {
        range: CellRange::parse_a1("D1").unwrap(),
        external_target: None,
        location: Some("'Other'!A1".to_owned()),
        tooltip: None,
        display: None,
    });
    model.sheets.push(sheet);
    let workbook = Workbook::from_model(model).unwrap();
    let content = workbook
        .export_structured(&options(json!({"includeHiddenSheets": true})))
        .unwrap()
        .unwrap()
        .content;
    let markdown = render_xlsx_markdown(&content, &XlsxMarkdownOptions::default())
        .unwrap()
        .markdown;
    assert!(
        markdown.contains("## Q1 &lt;b&gt;\\|\\#\\_\\`x\\`&amp; \\{y\\} (visibility unknown)\n")
    );
    assert!(markdown.contains(
        "| 1 | line one    &lt;img src=x onerror=alert\\(1\\)&gt; after | a\\|b &lt;/a&gt;&lt;script&gt;alert\\(1\\)&lt;/script&gt; |"
    ));
    for forbidden in [
        "<img",
        "<script",
        "</a>",
        "](javascript",
        "](data",
        "](java",
        "](%6A",
        "](&",
    ] {
        assert!(!markdown.contains(forbidden), "{forbidden} in:\n{markdown}");
    }
    assert_eq!(markdown.matches("(not linked)").count(), vectors.len());
    assert!(markdown.contains(
        "[Docs\\]\\(javascript:alert\\(1\\)\\)](https://example.com/a%20b?x=1&amp;y=%282%29%3C3%3E)"
    ));
    assert!(markdown.contains("in-workbook location 'Other'\\!A1"));
    for line in markdown.lines().filter(|line| line.starts_with("| ")) {
        assert!(line.ends_with(" |"), "{line}");
    }
}

#[test]
fn calculation_failures_are_reported_per_cell() {
    let workbook =
        Workbook::open_recalculated(&picture_fixture(), CalculationOptions::default()).unwrap();
    let content = workbook
        .export_structured(&XlsxExportOptions::default())
        .unwrap()
        .unwrap()
        .content;
    assert_eq!(
        cell(&content, 0, "A2").formula_result,
        Some(XlsxFormulaResult::Cycle)
    );
    assert_eq!(
        cell(&content, 0, "B1").formula_result,
        Some(XlsxFormulaResult::Unverified)
    );
    assert!(codes(&content).contains(&XlsxExportDiagnosticCode::CalculationFailure));
}

fn markers(markdown: &str) -> Vec<&str> {
    markdown
        .match_indices("<!-- xlsx-export:")
        .map(|(start, _)| {
            let end = markdown[start..].find("-->").unwrap() + start + 3;
            &markdown[start..end]
        })
        .collect()
}

#[test]
fn markdown_renders_bounded_labelled_grids() {
    let rendered = export_xlsx_markdown(
        &fixture(),
        &XlsxExportOptions::default(),
        &XlsxMarkdownOptions::default(),
    )
    .unwrap();
    let markdown = &rendered.markdown;
    assert!(markdown.contains("## Data\n"), "{markdown}");
    assert!(
        markdown.contains("<thead><tr><th></th><th>A</th><th>B</th><th>C</th><th>D</th><th>F</th>")
    );
    assert!(!markdown.contains("<th>E</th>"));
    assert!(markdown.contains("<tr><th>3</th><td>Pear &lt;script&gt;|*x* &amp; [y](z)</td>"));
    assert!(!markdown.contains("<script>"));
    assert!(!markdown.contains("<tr><th>4</th>"));
    assert!(markdown.contains(
        r#"<tr><th>6</th><td rowspan="1" colspan="2" data-merge="A6:B6">merged</td><td></td>"#
    ));
    assert!(!markdown.contains("covered"));
    assert!(markdown.contains("<tr><th>201</th>") && !markdown.contains("<tr><th>202</th>"));
    assert!(rendered.truncated);
    assert!(markdown.contains("## Summary\n"));
    assert!(markdown.contains("|  | A |\n| ---: | --- |\n| 1 | 4.5 |\n"));
    assert!(markdown.contains("| 3 | a\\|b &lt;i&gt; |\n"));
    assert!(markdown.contains("## Dialog (dialogsheet)\n"));
    assert!(markdown.contains("Picture Logo at H2:J6; alt text: Company &lt;logo&gt;"));
    assert!(markdown.contains(
        "Table Sales at A1:D3: 1 header rows, 0 totals rows; columns Name, Qty, Price, Total"
    ));
    assert!(markdown.contains("## Defined names\n"));
    assert!(markdown.contains("Local (Summary): Summary\\!\\$A\\$1"));

    let found = markers(markdown);
    assert_eq!(found.len(), rendered.anchors.len());
    for (index, (found, anchor)) in found.iter().zip(&rendered.anchors).enumerate() {
        assert_eq!(*found, anchor.marker);
        assert_eq!(anchor.marker, format!("<!-- xlsx-export:{index} -->"));
    }
    assert_eq!(
        to_json(&rendered.anchors[1].anchor),
        json!({"kind": "range", "sheet": {"sheetId": "sheet:0", "index": 0, "name": "Data"}, "a1": "A1:Z201"})
    );
    let lossy = rendered
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == XlsxExportDiagnosticCode::MarkdownLossy)
        .count();
    assert!(lossy >= 3, "{:?}", rendered.diagnostics);
    assert!(
        rendered
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code == XlsxExportDiagnosticCode::FormulaResultMissing)
    );

    let small = export_xlsx_markdown(
        &fixture(),
        &XlsxExportOptions::default(),
        &serde_json::from_value(json!({"maxRows": 2, "maxColumns": 2, "maxCells": 4})).unwrap(),
    )
    .unwrap();
    assert!(
        small.markdown.contains(
            "|  | A | B |\n| ---: | --- | --- |\n| 1 | Name | Qty |\n| 2 | Apple | 3 |\n\n"
        )
    );
    assert!(
        small
            .markdown
            .contains("Grid omitted: the Markdown cell limit was reached.")
    );

    for max_bytes in [4_096, 5_000, 8_000] {
        let bounded = export_xlsx_markdown(
            &fixture(),
            &XlsxExportOptions::default(),
            &serde_json::from_value(json!({"maxBytes": max_bytes})).unwrap(),
        )
        .unwrap();
        assert!(bounded.markdown.len() <= max_bytes);
        assert!(bounded.truncated);
        assert_eq!(markers(&bounded.markdown).len(), bounded.anchors.len());
    }
}

#[test]
fn markdown_respects_export_truncation() {
    let content = export(&fixture(), json!({"maxCells": 6}));
    let rendered = render_xlsx_markdown(&content, &XlsxMarkdownOptions::default()).unwrap();
    assert!(rendered.truncated);
    assert!(
        rendered
            .markdown
            .contains("| 1 | Name | Qty | Price | Total |")
    );
    assert!(!rendered.markdown.contains("| 2 |"));
    let content = export(&fixture(), json!({"maxCells": 4}));
    let rendered = render_xlsx_markdown(&content, &XlsxMarkdownOptions::default()).unwrap();
    assert!(
        rendered
            .markdown
            .contains("_No rows of this sheet are covered._")
    );
}

#[test]
fn render_validates_content_and_options() {
    let content = export(&fixture(), json!({}));
    let wire = serde_json::to_string(&content).unwrap();
    let rendered = render_xlsx_markdown_json(&wire, "{}").unwrap();
    assert!(rendered.contains("xlsx-export:0"));

    let mut tampered = to_json(&content);
    tampered["schemaVersion"] = json!(2);
    assert!(matches!(
        render_xlsx_markdown_json(&tampered.to_string(), "{}"),
        Err(Error::InvalidRequest(_))
    ));
    let mut outside = content.clone();
    outside.sheets[0].selected_range = Some("A1:B2".to_owned());
    assert!(render_xlsx_markdown(&outside, &XlsxMarkdownOptions::default()).is_err());
    let mut unordered = content.clone();
    unordered.sheets[0].cells.swap(0, 1);
    assert!(render_xlsx_markdown(&unordered, &XlsxMarkdownOptions::default()).is_err());
    let mut ranged = content.clone();
    ranged.sheets[0].cells[0].anchor = XlsxAnchor::Range {
        sheet: betteroffice_xlsx::XlsxSheetIdentity {
            sheet_id: "sheet:0".to_owned(),
            index: 0,
            name: "Data".to_owned(),
        },
        a1: "A1".to_owned(),
    };
    assert!(matches!(
        render_xlsx_markdown(&ranged, &XlsxMarkdownOptions::default()),
        Err(Error::InvalidRequest(_))
    ));
    let mut wire_tampered = to_json(&content);
    wire_tampered["sheets"][0]["cells"][3]["anchor"]["kind"] = json!("range");
    assert!(render_xlsx_markdown_json(&wire_tampered.to_string(), "{}").is_err());
    let mutations: Vec<fn(&mut XlsxStructuredContent)> = vec![
        |content| content.sheets[0].merges[0].anchor = content.sheets[1].anchor.clone(),
        |content| content.sheets[0].hidden_rows.push("0:3".to_owned()),
        |content| content.sheets[0].hidden_columns.push("aa:AB".to_owned()),
        |content| content.sheets[0].selected_range = Some("B1:A2".to_owned()),
        |content| content.sheets[0].truncated = true,
        |content| content.sheets.swap(0, 1),
        |content| {
            content.sheets[0].objects[0].anchor = Some(XlsxAnchor::Sheet {
                sheet: betteroffice_xlsx::XlsxSheetIdentity {
                    sheet_id: "sheet:0".to_owned(),
                    index: 0,
                    name: "Data".to_owned(),
                },
            })
        },
        |content| content.defined_names[0].anchor = content.sheets[0].anchor.clone(),
    ];
    for (index, mutate) in mutations.into_iter().enumerate() {
        let mut broken = content.clone();
        mutate(&mut broken);
        assert!(
            render_xlsx_markdown(&broken, &XlsxMarkdownOptions::default()).is_err(),
            "mutation {index}"
        );
    }
    let mut foreign = content.clone();
    foreign.sheets[0].cells[0].anchor = XlsxAnchor::Cell {
        sheet: betteroffice_xlsx::XlsxSheetIdentity {
            sheet_id: "sheet:1".to_owned(),
            index: 1,
            name: "Summary".to_owned(),
        },
        a1: "A1".to_owned(),
    };
    assert!(render_xlsx_markdown(&foreign, &XlsxMarkdownOptions::default()).is_err());
    for options in [
        json!({"maxRows": 0}),
        json!({"maxColumns": 16_385}),
        json!({"maxBytes": 100}),
        json!({"maxCells": 1_000_001}),
    ] {
        assert!(render_xlsx_markdown(&content, &serde_json::from_value(options).unwrap()).is_err());
    }
    assert!(matches!(
        render_xlsx_markdown_json(&wire, r#"{"rows":1}"#),
        Err(Error::InvalidRequest(_))
    ));
    let workbook = Workbook::open(&fixture()).unwrap();
    let refused = workbook
        .export_markdown(
            &XlsxExportOptions::default(),
            &serde_json::from_value(json!({"maxRows": 0})).unwrap(),
        )
        .unwrap()
        .unwrap_err();
    assert_eq!(refused.failure.code, XlsxExportFailureCode::InvalidOptions);
    assert_eq!(
        XlsxExportScope {
            sheet: 0,
            range: None
        },
        serde_json::from_value(json!({"sheet": 0})).unwrap()
    );
}

fn golden(name: &str, actual: &str) {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/structured-export")
        .join(name);
    if std::env::var("GOLDEN_UPDATE").is_ok_and(|value| value == "1") {
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, actual).unwrap();
    }
    let expected = std::fs::read_to_string(&path)
        .unwrap_or_else(|_| panic!("{} is missing; run with GOLDEN_UPDATE=1", path.display()));
    assert_eq!(
        actual, expected,
        "{name} drifted; review and rerun with GOLDEN_UPDATE=1"
    );
}

#[test]
fn exports_match_their_goldens() {
    let bytes = fixture();
    let content = export(&bytes, json!({}));
    golden(
        "fixture.json",
        &(serde_json::to_string_pretty(&content).unwrap() + "\n"),
    );
    let rendered = render_xlsx_markdown(
        &content,
        &serde_json::from_value(json!({"maxRows": 12, "maxColumns": 10})).unwrap(),
    )
    .unwrap();
    golden("fixture.md", &rendered.markdown);
    golden(
        "fixture.markdown-anchors.json",
        &(serde_json::to_string_pretty(&json!({
            "anchors": rendered.anchors,
            "diagnostics": rendered.diagnostics,
            "truncated": rendered.truncated,
        }))
        .unwrap()
            + "\n"),
    );
}
