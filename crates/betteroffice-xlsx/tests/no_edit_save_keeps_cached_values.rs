use std::fmt::Write;

use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, SheetId, Workbook};
use xlsx_model::addr::col_to_letters;
use xlsx_parse::{
    SaveEdits, parse_workbook_with_package,
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const SHEET: &str = "xl/worksheets/sheet1.xml";
const CONTENT_TYPES: &str = "[Content_Types].xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const FULL_CALC: &str = r#"<calcPr fullCalcOnLoad="1"/>"#;

type Parts = Vec<(String, Vec<u8>)>;

fn package_parts(sheet_data: &str, calc_pr: &str) -> Parts {
    [
        (
            CONTENT_TYPES,
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{calc_pr}</workbook>"#),
        ),
        (
            WORKBOOK_RELS,
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
        ),
        (
            SHEET,
            format!(r#"<worksheet xmlns="{MAIN}" xmlns:r="{RELS}"><sheetData>{sheet_data}</sheetData></worksheet>"#),
        ),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect()
}

fn add_part(parts: &mut Parts, path: &str, content_type: &str, xml: &str) {
    insert_xml(
        parts,
        CONTENT_TYPES,
        "</Types>",
        &format!(r#"<Override PartName="/{path}" ContentType="{content_type}"/>"#),
    );
    parts.push((path.to_owned(), xml.as_bytes().to_vec()));
}

fn insert_xml(parts: &mut [(String, Vec<u8>)], path: &str, before: &str, xml: &str) {
    let (_, bytes) = parts.iter_mut().find(|(name, _)| name == path).unwrap();
    let mut part = String::from_utf8(std::mem::take(bytes)).unwrap();
    part.insert_str(part.rfind(before).unwrap(), xml);
    *bytes = part.into_bytes();
}

fn part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> &'a [u8] {
    &parts.iter().find(|(name, _)| name == path).unwrap().1
}

fn sheet_xml(parts: &[(String, Vec<u8>)]) -> &str {
    std::str::from_utf8(part(parts, SHEET)).unwrap()
}

fn cell<'a>(sheet: &'a str, address: &str) -> &'a str {
    let start = sheet.find(&format!(r#"<c r="{address}""#)).unwrap();
    let from_cell = &sheet[start..];
    let tag_end = from_cell.find('>').unwrap() + 1;
    if from_cell[..tag_end].ends_with("/>") {
        return &from_cell[..tag_end];
    }
    let end = from_cell.find("</c>").unwrap() + "</c>".len();
    &from_cell[..end]
}

fn open_recalculated(parts: &[(String, Vec<u8>)]) -> Workbook {
    let source = ooxml_opc::rezip_parts(parts).unwrap();
    Workbook::open_recalculated(&source, CalculationOptions::default()).unwrap()
}

fn saved_parts(workbook: &Workbook) -> Parts {
    ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap()
}

#[test]
fn keeps_cached_array_booleans_byte_for_byte() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1" t="b"><f t="array" ref="A1:A3">FALSE()</f><v>1</v></c></row>"#,
            r#"<row r="2"><c r="A2" t="b"><v>1</v></c></row>"#,
            r#"<row r="3"><c r="A3" t="b"><v>1</v></c></row>"#,
        ),
        FULL_CALC,
    );
    let workbook = open_recalculated(&source);
    let sheet = &workbook.model().sheets[0];
    for row in 0..3 {
        assert_eq!(
            sheet.cell(CellRef::new(row, 0)).unwrap().value,
            CellValue::Bool { value: false }
        );
    }
    let saved = saved_parts(&workbook);
    for address in ["A1", "A2", "A3"] {
        assert_eq!(
            cell(sheet_xml(&saved), address),
            cell(sheet_xml(&source), address),
        );
    }
}

#[test]
fn keeps_cached_cells_and_the_anchor_when_a_spill_shrinks() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(2)</f><v>1</v></c>"#,
            r#"<c r="B1"><f>2+3</f></c></row>"#,
            r#"<row r="2"><c r="A2"><v>2</v></c></row>"#,
            r#"<row r="3"><c r="A3"><v>3</v></c></row>"#,
        ),
        FULL_CALC,
    );
    let workbook = open_recalculated(&source);
    assert!(
        workbook.model().sheets[0]
            .cell(CellRef::new(2, 0))
            .is_none()
    );
    let saved = saved_parts(&workbook);
    for address in ["A1", "A2", "A3"] {
        assert_eq!(
            cell(sheet_xml(&saved), address),
            cell(sheet_xml(&source), address),
        );
    }
    assert!(cell(sheet_xml(&saved), "B1").contains("<v>5</v>"));
}

#[test]
fn fills_a_formula_with_a_spaced_empty_value() {
    let mut source = package_parts(
        concat!(
            r#"<row r="15"><c r="B15" s="21"><f>COUNTA(B16:B30)</f><v /></c></row>"#,
            r#"<row r="16"><c r="B16"><v>9</v></c></row>"#,
        ),
        FULL_CALC,
    );
    insert_xml(
        &mut source,
        WORKBOOK_RELS,
        "</Relationships>",
        &format!(r#"<Relationship Id="rId2" Type="{RELS}/styles" Target="styles.xml"/>"#),
    );
    let xfs = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#.repeat(22);
    add_part(
        &mut source,
        "xl/styles.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        &format!(r#"<styleSheet xmlns="{MAIN}"><cellXfs count="22">{xfs}</cellXfs></styleSheet>"#),
    );
    let saved = saved_parts(&open_recalculated(&source));
    assert_eq!(
        cell(sheet_xml(&saved), "B15"),
        r#"<c r="B15" s="21"><f>COUNTA(B16:B30)</f><v>1</v></c>"#,
    );
}

#[test]
fn fills_formulas_with_self_closing_and_paired_empty_values() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1"><f>2+3</f><v/></c>"#,
            r#"<c r="B1"><f>4+5</f><v></v></c>"#,
            r#"<c r="C1" t="b"><f>FALSE()</f><v/></c></row>"#,
        ),
        FULL_CALC,
    );
    let saved = saved_parts(&open_recalculated(&source));
    let sheet = sheet_xml(&saved);
    assert!(cell(sheet, "A1").contains("<v>5</v>"));
    assert!(cell(sheet, "B1").contains("<v>9</v>"));
    assert!(cell(sheet, "C1").contains("<v>0</v>"));
}

#[test]
fn fills_a_spill_anchor_and_keeps_the_source_members_and_ref() {
    let source = package_parts(
        concat!(
            r#"<row r="2"><c r="P2"><f t="array" ref="P2:T19">_xlfn.SEQUENCE(18,5)</f><v /></c>"#,
            r#"<c r="Q2" t="n"><v>13</v></c><c r="R2"><v /></c></row>"#,
        ),
        FULL_CALC,
    );
    let saved = saved_parts(&open_recalculated(&source));
    assert_eq!(
        sheet_xml(&saved),
        sheet_xml(&source)
            .replacen("<v />", "<v>1</v>", 1)
            .replacen("<v />", "<v>3</v>", 1),
    );
}

#[test]
fn leaves_new_followers_inside_and_outside_the_source_spill_absent_on_save() {
    for cache in ["<v>1</v>", "<v />"] {
        let source = package_parts(
            &format!(
                concat!(
                    r#"<row r="1"><c r="A1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(4)</f>{}</c></row>"#,
                    r#"<row r="2"><c r="A2"><v>2</v></c></row>"#,
                ),
                cache,
            ),
            FULL_CALC,
        );
        let workbook = open_recalculated(&source);
        let sheet = &workbook.model().sheets[0];
        for row in 0..4 {
            assert_eq!(
                sheet.cell(CellRef::new(row, 0)).unwrap().value,
                CellValue::Number {
                    value: f64::from(row + 1),
                }
            );
        }
        let saved = saved_parts(&workbook);
        assert_eq!(
            sheet_xml(&saved),
            sheet_xml(&source).replace("<v />", "<v>1</v>"),
        );
    }
}

#[test]
fn keeps_authored_blanks_outside_the_source_spill() {
    for blank in [r#"<c r="A3" s="1"/>"#, r#"<c r="A3"/>"#] {
        let mut source = package_parts(
            &format!(
                concat!(
                    r#"<row r="1"><c r="A1"><f t="array" ref="A1:A2">_xlfn.SEQUENCE(3)</f><v>1</v></c>"#,
                    r#"<c r="B1"><f>2+3</f></c></row>"#,
                    r#"<row r="2"><c r="A2"><v>2</v></c></row>"#,
                    r#"<row r="3">{}</row>"#,
                ),
                blank,
            ),
            FULL_CALC,
        );
        insert_xml(
            &mut source,
            WORKBOOK_RELS,
            "</Relationships>",
            &format!(r#"<Relationship Id="rId2" Type="{RELS}/styles" Target="styles.xml"/>"#),
        );
        let xfs = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#.repeat(2);
        add_part(
            &mut source,
            "xl/styles.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            &format!(r#"<styleSheet xmlns="{MAIN}"><cellXfs count="2">{xfs}</cellXfs></styleSheet>"#),
        );

        let workbook = open_recalculated(&source);
        assert_eq!(
            workbook.model().sheets[0]
                .cell(CellRef::new(2, 0))
                .unwrap()
                .value,
            CellValue::Number { value: 3.0 },
        );
        let saved = saved_parts(&workbook);
        let sheet = sheet_xml(&saved);
        assert_eq!(cell(sheet, "A3"), blank);
        assert!(cell(sheet, "B1").contains("<v>5</v>"));

        let reopened = open_recalculated(&saved);
        assert_eq!(
            reopened.model().sheets[0]
                .cell(CellRef::new(0, 0))
                .unwrap()
                .value,
            CellValue::Number { value: 1.0 },
        );
    }
}

#[test]
fn fills_uncached_formula_cells_and_keeps_a_cached_array_member() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1"><f>2+3</f></c><c r="B1"><f>4+5</f><v/></c>"#,
            r#"<c r="C1"><f t="array" ref="C1:C2">2</f></c></row>"#,
            r#"<row r="2"><c r="C2"><v>7</v></c></row>"#,
        ),
        FULL_CALC,
    );
    let saved = saved_parts(&open_recalculated(&source));
    let sheet = sheet_xml(&saved);
    for (address, value) in [("A1", 5), ("B1", 9), ("C1", 2)] {
        assert!(cell(sheet, address).contains(&format!("<v>{value}</v>")));
    }
    assert_eq!(cell(sheet, "C2"), cell(sheet_xml(&source), "C2"));
}

#[test]
fn fills_typed_cells_without_a_source_value_element() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1" t="b"><f>TRUE()</f></c>"#,
            r#"<c r="B1" t="str"><f>"calculated"</f></c>"#,
            r#"<c r="C1" t="str"><f>"calculated"</f><v/></c>"#,
            r#"<c r="D1" t="b"><f t="array" ref="D1:D2" ca="1">TRUE()</f><v>0</v></c></row>"#,
            r#"<row r="2"><c r="D2" t="b"/></row>"#,
        ),
        FULL_CALC,
    );
    let saved = saved_parts(&open_recalculated(&source));
    let sheet = sheet_xml(&saved);
    for address in ["A1", "D2"] {
        let boolean = cell(sheet, address);
        assert!(boolean.contains(r#"t="b""#) && boolean.contains("<v>1</v>"));
    }
    for address in ["B1", "C1"] {
        assert!(cell(sheet, address).contains("<v>calculated</v>"));
    }
    assert_eq!(cell(sheet, "D1"), cell(sheet_xml(&source), "D1"));
}

#[test]
fn keeps_the_cached_formula_and_fills_the_uncached_formula_and_array_member() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1" t="n"><f t="array" ref="A1:A2" ca="1">1</f><v>10</v></c>"#,
            r#"<c r="B1"><f>4+5</f></c></row>"#,
            r#"<row r="2"><c r="A2"/></row>"#,
        ),
        FULL_CALC,
    );
    let saved = saved_parts(&open_recalculated(&source));
    let sheet = sheet_xml(&saved);
    assert_eq!(cell(sheet, "A1"), cell(sheet_xml(&source), "A1"));
    assert!(cell(sheet, "A2").contains("<v>1</v>"));
    assert!(cell(sheet, "B1").contains("<v>9</v>"));
}

#[test]
fn fills_existing_array_members_without_scanning_formula_ranges() {
    const COLS: u32 = 64;
    const ROWS: u32 = 8;
    let mut sheet_data = String::from(r#"<row r="1">"#);
    for col in 0..COLS {
        let col = col_to_letters(col);
        write!(
            sheet_data,
            r#"<c r="{col}1"><f t="array" ref="{col}1:{col}{ROWS}" ca="1">1</f><v>10</v></c>"#,
        )
        .unwrap();
    }
    sheet_data.push_str("</row>");
    for row in 2..=ROWS {
        write!(sheet_data, r#"<row r="{row}">"#).unwrap();
        for col in 0..COLS {
            let col = col_to_letters(col);
            write!(sheet_data, r#"<c r="{col}{row}"/>"#).unwrap();
        }
        sheet_data.push_str("</row>");
    }
    let source = package_parts(&sheet_data, FULL_CALC);
    let workbook = open_recalculated(&source);

    xlsx_model::open_timing::begin();
    let saved = workbook.save();
    let counts = xlsx_model::open_timing::finish().counts;
    let saved = ooxml_opc::unzip_parts(&saved.unwrap()).unwrap();
    let checks = counts.get("save_array_range_changes").copied().unwrap_or(0);
    assert_eq!(checks, 0);
    let sheet = sheet_xml(&saved);
    assert_eq!(cell(sheet, "A1"), cell(sheet_xml(&source), "A1"));
    assert!(cell(sheet, "A8").contains("<v>1</v>"));
}

#[test]
fn keeps_every_part_byte_for_byte_when_all_caches_are_correct() {
    let source = package_parts(
        r#"<row r="1"><c r="A1"><f>2+3</f><v>5</v></c><c r="B1"><f>A1*2</f><v>10</v></c></row>"#,
        r#"<calcPr fullCalcOnLoad="true"/>"#,
    );
    let saved = saved_parts(&open_recalculated(&source));
    assert_eq!(saved, source);
}

#[test]
fn keeps_a_whitespace_cache_before_an_extension_value_element() {
    for whitespace in ["  ", "&#32;&#32;"] {
        let source = package_parts(
            &format!(
                concat!(
                    r#"<row r="1"><c r="A1" t="str"><f>"  "</f><v>{}</v>"#,
                    r#"<extLst><ext uri="urn:example"><x:v xmlns:x="urn:example"/></ext></extLst></c>"#,
                    r#"<c r="B1"><f>2+3</f></c></row>"#,
                ),
                whitespace,
            ),
            FULL_CALC,
        );
        let saved = saved_parts(&open_recalculated(&source));
        assert_eq!(
            cell(sheet_xml(&saved), "A1"),
            cell(sheet_xml(&source), "A1"),
        );
        assert!(cell(sheet_xml(&saved), "B1").contains("<v>5</v>"));
    }
}

#[test]
fn keeps_shared_string_indices_and_literal_cache_markup() {
    let mut source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1" t="s"><f>"new"</f><v>1</v></c>"#,
            r#"<c r="B1" t="n"><f>5+6</f><v>10</v></c>"#,
            r#"<c r="C1" t="e"><f>1+1</f><v>#N/A</v></c>"#,
            r#"<c r="D1" t="inlineStr"><f>"new"</f><is><t>original</t></is></c>"#,
            r#"<c r="E1"><f>2+3</f></c></row>"#,
        ),
        FULL_CALC,
    );
    insert_xml(
        &mut source,
        WORKBOOK_RELS,
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rId2" Type="{RELS}/sharedStrings" Target="sharedStrings.xml"/>"#
        ),
    );
    add_part(
        &mut source,
        "xl/sharedStrings.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        &format!(r#"<sst xmlns="{MAIN}"><si><t>original</t></si><si><t>original</t></si></sst>"#),
    );
    let saved = saved_parts(&open_recalculated(&source));
    for address in ["A1", "B1", "C1", "D1"] {
        assert_eq!(
            cell(sheet_xml(&saved), address),
            cell(sheet_xml(&source), address),
        );
    }
    assert!(cell(sheet_xml(&saved), "E1").contains("<v>5</v>"));
}

#[test]
fn saves_the_recalculated_value_after_an_unrelated_edit() {
    let source = package_parts(
        r#"<row r="1"><c r="A1"><f>2+3</f><v>10</v></c></row>"#,
        FULL_CALC,
    );
    let mut workbook = open_recalculated(&source);
    assert_eq!(saved_parts(&workbook), source);
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 1),
            "8",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = saved_parts(&workbook);
    let sheet = sheet_xml(&saved);
    assert!(cell(sheet, "A1").contains("<v>5</v>"));
    assert!(cell(sheet, "B1").contains("<v>8</v>"));
}

fn add_chart(source: &mut Parts, path: &str) {
    insert_xml(source, SHEET, "</worksheet>", r#"<drawing r:id="rId1"/>"#);
    source.push((
        "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
        format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).into_bytes(),
    ));
    add_part(
        source,
        "xl/drawings/drawing1.xml",
        "application/vnd.openxmlformats-officedocument.drawing+xml",
        &format!(
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{RELS}"><xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#
        ),
    );
    let target = path.strip_prefix("xl/").unwrap();
    source.push((
        "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
        format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/chart" Target="../{target}"/></Relationships>"#).into_bytes(),
    ));
    add_part(
        source,
        path,
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:val><c:numRef><c:f>Sheet1!$A$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>999</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
    );
}

#[test]
fn keeps_chart_caches_when_filling_a_referenced_cell() {
    const CHART: &str = "xl/charts/chart1.xml";
    let mut source = package_parts(r#"<row r="1"><c r="A1"><f>2+3</f></c></row>"#, FULL_CALC);
    add_chart(&mut source, CHART);
    let mut workbook = open_recalculated(&source);
    assert_eq!(workbook.model().sheets[0].charts.len(), 1);
    let saved = saved_parts(&workbook);
    assert!(cell(sheet_xml(&saved), "A1").contains("<v>5</v>"));
    assert_eq!(part(&saved, CHART), part(&source, CHART));

    workbook.recalculate_all(CalculationOptions::default());
    let saved = saved_parts(&workbook);
    let chart = std::str::from_utf8(part(&saved, CHART)).unwrap();
    assert!(chart.contains(r#"<c:pt idx="0"><c:v>5</c:v></c:pt>"#));
}

#[test]
fn keeps_chart_caches_at_a_nonstandard_path_when_filling_a_referenced_cell() {
    const CHART: &str = "xl/visuals/chart1.xml";
    let mut source = package_parts(r#"<row r="1"><c r="A1"><f>2+3</f></c></row>"#, FULL_CALC);
    add_chart(&mut source, CHART);
    let mut workbook = open_recalculated(&source);
    assert_eq!(workbook.model().sheets[0].charts.len(), 1);
    let saved = saved_parts(&workbook);
    assert!(cell(sheet_xml(&saved), "A1").contains("<v>5</v>"));
    assert_eq!(part(&saved, CHART), part(&source, CHART));

    workbook.recalculate_all(CalculationOptions::default());
    let saved = saved_parts(&workbook);
    let chart = std::str::from_utf8(part(&saved, CHART)).unwrap();
    assert!(chart.contains(r#"<c:pt idx="0"><c:v>5</c:v></c:pt>"#));
}

#[test]
fn switching_tabs_refreshes_chart_caches_for_an_unflagged_workbook() {
    const CHART: &str = "xl/charts/chart1.xml";
    let mut source = package_parts(r#"<row r="1"><c r="A1"><f>2+3</f><v>999</v></c></row>"#, "");
    add_chart(&mut source, CHART);
    insert_xml(
        &mut source,
        "xl/workbook.xml",
        "</sheets>",
        r#"<sheet name="Sheet2" sheetId="2" r:id="rId2"/>"#,
    );
    insert_xml(
        &mut source,
        WORKBOOK_RELS,
        "</Relationships>",
        &format!(
            r#"<Relationship Id="rId2" Type="{RELS}/worksheet" Target="worksheets/sheet2.xml"/>"#
        ),
    );
    add_part(
        &mut source,
        "xl/worksheets/sheet2.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        &format!(r#"<worksheet xmlns="{MAIN}"><sheetData/></worksheet>"#),
    );

    let mut workbook = open_recalculated(&source);
    workbook.set_active_sheet(SheetId(1)).unwrap();
    let saved = saved_parts(&workbook);
    assert!(cell(sheet_xml(&saved), "A1").contains("<v>5</v>"));
    let chart = std::str::from_utf8(part(&saved, CHART)).unwrap();
    assert!(chart.contains(r#"<c:pt idx="0"><c:v>5</c:v></c:pt>"#));
}

#[test]
fn keeps_every_part_when_the_only_calculated_values_are_empty_strings() {
    let source = package_parts(
        r#"<row r="1"><c r="A1" s="0"><f>IFERROR(1/0,"")</f><v /></c><c r="B1"><f t="array" ref="B1">IFERROR(1/0,"")</f><v /></c></row>"#,
        FULL_CALC,
    );
    let workbook = open_recalculated(&source);
    assert_eq!(saved_parts(&workbook), source);
}

#[test]
fn fills_real_results_and_keeps_empty_string_results_as_the_source_wrote_them() {
    let source = package_parts(
        r#"<row r="1"><c r="A1"><f>2+3</f><v /></c><c r="B1"><f>IFERROR(1/0,"")</f><v /></c></row>"#,
        FULL_CALC,
    );
    let workbook = open_recalculated(&source);
    let saved = saved_parts(&workbook);
    assert!(cell(sheet_xml(&saved), "A1").contains("<v>5</v>"));
    assert_eq!(
        cell(sheet_xml(&saved), "B1"),
        cell(sheet_xml(&source), "B1")
    );
}

#[test]
fn keeps_an_unmodeled_array_follower_when_its_result_is_an_empty_string() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1"><f t="array" ref="A1:A2">""</f></c>"#,
            r#"<c r="B1"><f>2+3</f></c></row>"#,
            r#"<row r="2"><c r="A2" t="b"><v/></c></row>"#,
        ),
        FULL_CALC,
    );
    let workbook = open_recalculated(&source);
    assert_eq!(
        workbook.model().sheets[0]
            .cell(CellRef::new(1, 0))
            .unwrap()
            .value,
        CellValue::Text {
            value: String::new(),
        },
    );
    let saved = saved_parts(&workbook);
    assert_eq!(
        cell(sheet_xml(&saved), "A2"),
        cell(sheet_xml(&source), "A2"),
    );
    assert!(cell(sheet_xml(&saved), "B1").contains("<v>5</v>"));
}

#[test]
fn checks_changed_array_followers_in_row_order() {
    let source = package_parts(
        concat!(
            r#"<row r="1"><c r="A1"><f t="array" ref="A1:C3" ca="1">1</f><v>1</v></c></row>"#,
            r#"<row r="2"><c r="C2"><v>1</v></c>"#,
            r#"<c r="E2"><f t="array" ref="E2:F3" ca="1">1</f><v>1</v></c></row>"#,
            r#"<row r="3"><c r="B3"/></row>"#,
            r#"<row r="5"><c r="A5"><f t="array" ref="A5:C5" ca="1">1</f><v>1</v></c></row>"#,
        ),
        FULL_CALC,
    );
    let mut parsed = parse_workbook_with_package(&source).unwrap();
    for address in ["C2", "B3"] {
        parsed.workbook.sheets[0].set_calculated_value(
            CellRef::parse_a1(address).unwrap(),
            CellValue::Number { value: 2.0 },
        );
    }
    xlsx_model::open_timing::begin();
    let saved = serialize_workbook_with_package_and_origins_after_edits_and_active_sheet(
        &parsed.workbook,
        &parsed.package,
        &[Some(0)],
        &[parsed.package.source_shared_string_cells(0)],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
        parsed.active_sheet,
    );
    let counts = xlsx_model::open_timing::finish().counts;
    let saved = saved.unwrap();
    let (_, sheet) = saved.iter().find(|(path, _)| path == SHEET).unwrap();
    let sheet = std::str::from_utf8(sheet).unwrap();
    assert!(cell(sheet, "A1").contains(r#"<f t="array" ref="A1:C3">1</f>"#));
    for address in ["E2", "A5"] {
        assert_eq!(cell(sheet, address), cell(sheet_xml(&source), address));
    }
    assert_eq!(counts.get("save_array_range_changes"), Some(&1));
}

#[test]
fn keeps_a_style_only_string_typed_cell_beside_an_array() {
    let source = package_parts(
        r#"<row r="1"><c r="A1"><f t="array" ref="A1:A2">_xlfn.SEQUENCE(3)</f><v>1</v></c><c r="B1" s="1" t="str"/></row><row r="2"><c r="A2"><v>2</v></c><c r="B2" s="1" t="str"/></row><row r="3"><c r="A3" s="1" t="str"/></row>"#,
        FULL_CALC,
    );
    let workbook = open_recalculated(&source);
    let saved = saved_parts(&workbook);
    for address in ["B1", "B2", "A3"] {
        assert_eq!(
            cell(sheet_xml(&saved), address),
            cell(sheet_xml(&source), address),
            "{address}"
        );
    }
}
