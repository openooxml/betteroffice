//! fixtures are raw xml parts written from the ecma-376 spec. the parser
//! matches local element names, so fixtures omit namespace declarations.

use xlsx_model::styles::{BorderStyle, Color, Fill, FormatCode, HAlign, VAlign};
use xlsx_model::{
    Cell, CellRef, CellValue, DateSystem, DefinedName, ErrorValue, SheetId, Workbook,
};

use crate::{ParseError, parse_workbook, serialize_workbook};

/// assemble a one-sheet package around a worksheet body and optional shared
/// strings, so each test only spells out the part under exercise.
fn package(worksheet_body: &str, shared: &[&str], date1904: bool) -> Vec<(String, Vec<u8>)> {
    let pr = if date1904 {
        r#"<workbookPr date1904="1"/>"#
    } else {
        ""
    };
    let workbook = format!(
        r#"<workbook>{pr}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
    );
    let rels = r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let worksheet = format!("<worksheet>{worksheet_body}</worksheet>");

    let mut parts = vec![
        ("xl/workbook.xml".to_string(), workbook.into_bytes()),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels.as_bytes().to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            worksheet.into_bytes(),
        ),
    ];
    if !shared.is_empty() {
        let items: String = shared
            .iter()
            .map(|s| format!("<si><t>{s}</t></si>"))
            .collect();
        let sst = format!("<sst>{items}</sst>");
        parts.push(("xl/sharedStrings.xml".to_string(), sst.into_bytes()));
    }
    parts
}

fn cell_at(wb: &Workbook, a1: &str) -> Cell {
    let addr = CellRef::parse_a1(a1).unwrap();
    wb.sheets[0].cell(addr).cloned().unwrap_or_default()
}

#[test]
fn parses_shared_string_number_formula_bool_error() {
    let body = r#"
        <sheetData>
            <row r="1" ht="30">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>2.5</v></c>
                <c r="C1"><f>A1&amp;B1</f><v>5</v></c>
                <c r="D1" t="b"><v>1</v></c>
                <c r="E1" t="e"><v>#DIV/0!</v></c>
            </row>
        </sheetData>
        <mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>
        <cols><col min="2" max="3" width="12.5"/></cols>
    "#;
    let wb = parse_workbook(&package(body, &["hello"], false)).unwrap();

    assert_eq!(wb.sheets.len(), 1);
    assert_eq!(wb.sheets[0].name, "Sheet1");
    assert_eq!(
        cell_at(&wb, "A1").value,
        CellValue::Text {
            value: "hello".into()
        }
    );
    assert_eq!(cell_at(&wb, "B1").value, CellValue::Number { value: 2.5 });

    let c1 = cell_at(&wb, "C1");
    assert_eq!(c1.value, CellValue::Number { value: 5.0 });
    assert_eq!(c1.formula.as_deref(), Some("A1&B1"));

    assert_eq!(cell_at(&wb, "D1").value, CellValue::Bool { value: true });
    assert_eq!(
        cell_at(&wb, "E1").value,
        CellValue::Error {
            value: ErrorValue::Div0
        }
    );

    assert_eq!(wb.sheets[0].merges.len(), 1);
    assert_eq!(wb.sheets[0].merges[0].to_a1(), "A1:B2");
    assert_eq!(wb.sheets[0].col_widths.get(&1), Some(&12.5));
    assert_eq!(wb.sheets[0].col_widths.get(&2), Some(&12.5));
    assert_eq!(wb.sheets[0].row_heights.get(&0), Some(&30.0));
}

#[test]
fn parses_inline_string() {
    let body = r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>inline &lt;here&gt;</t></is></c></row></sheetData>"#;
    let wb = parse_workbook(&package(body, &[], false)).unwrap();
    assert_eq!(
        cell_at(&wb, "A1").value,
        CellValue::Text {
            value: "inline <here>".into()
        }
    );
}

#[test]
fn flattens_rich_run_shared_string() {
    let sst = "<sst><si><r><t>Hello </t></r><r><t>World</t></r></si></sst>";
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_string(), sst.as_bytes().to_vec()));
    let wb = parse_workbook(&parts).unwrap();
    assert_eq!(
        cell_at(&wb, "A1").value,
        CellValue::Text {
            value: "Hello World".into()
        }
    );
}

#[test]
fn honors_1904_date_system() {
    let wb = parse_workbook(&package("<sheetData/>", &[], true)).unwrap();
    assert_eq!(wb.date_system, DateSystem::V1904);
    let wb = parse_workbook(&package("<sheetData/>", &[], false)).unwrap();
    assert_eq!(wb.date_system, DateSystem::V1900);
}

#[test]
fn parses_and_round_trips_scoped_defined_names() {
    let mut parts = package("<sheetData/>", &[], false);
    let workbook = br#"
        <workbook>
            <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
            <definedNames>
                <definedName name="TaxRate">0.19</definedName>
                <definedName name="Input" localSheetId="0" hidden="1">Sheet1!$B$2</definedName>
            </definedNames>
        </workbook>
    "#;
    parts
        .iter_mut()
        .find(|(name, _)| name == "xl/workbook.xml")
        .unwrap()
        .1 = workbook.to_vec();

    let parsed = parse_workbook(&parts).unwrap();
    assert_eq!(
        parsed.defined_names,
        vec![
            DefinedName {
                name: "TaxRate".into(),
                formula: "0.19".into(),
                local_sheet: None,
                hidden: false,
            },
            DefinedName {
                name: "Input".into(),
                formula: "Sheet1!$B$2".into(),
                local_sheet: Some(SheetId(0)),
                hidden: true,
            },
        ]
    );

    let reparsed = parse_workbook(&serialize_workbook(&parsed).unwrap()).unwrap();
    assert_eq!(reparsed.defined_names, parsed.defined_names);
}

#[test]
fn skips_unknown_elements() {
    let body = r#"
        <extLst><ext uri="whatever"><custom><deep/></custom></ext></extLst>
        <sheetData>
            <row r="1"><c r="A1"><v>1</v></c></row>
        </sheetData>
        <weird attr="x"/>
    "#;
    let wb = parse_workbook(&package(body, &[], false)).unwrap();
    assert_eq!(cell_at(&wb, "A1").value, CellValue::Number { value: 1.0 });
}

#[test]
fn rejects_malformed_cell_ref() {
    let body = r#"<sheetData><row r="1"><c r="not-a-ref"><v>1</v></c></row></sheetData>"#;
    let err = parse_workbook(&package(body, &[], false)).unwrap_err();
    assert!(matches!(err, ParseError::Malformed(_)), "got {err:?}");
}

#[test]
fn deep_nesting_hits_depth_cap_without_overflow() {
    let deep = format!("{}{}", "<x>".repeat(200), "</x>".repeat(200));
    let body = format!("<sheetData>{deep}</sheetData>");
    let err = parse_workbook(&package(&body, &[], false)).unwrap_err();
    assert_eq!(err, ParseError::DepthExceeded);
}

#[test]
fn missing_workbook_part_errors() {
    let err =
        parse_workbook(&[("xl/sharedStrings.xml".to_string(), b"<sst/>".to_vec())]).unwrap_err();
    assert!(matches!(err, ParseError::MissingPart(_)), "got {err:?}");
}

#[test]
fn empty_cell_ref_uses_column_cursor() {
    let body = r#"<sheetData><row r="2"><c><v>10</v></c><c><v>20</v></c></row></sheetData>"#;
    let wb = parse_workbook(&package(body, &[], false)).unwrap();
    assert_eq!(cell_at(&wb, "A2").value, CellValue::Number { value: 10.0 });
    assert_eq!(cell_at(&wb, "B2").value, CellValue::Number { value: 20.0 });
}

#[test]
fn normalizes_overlapping_merges_in_declaration_order() {
    let body = r#"
        <sheetData/>
        <mergeCells count="5">
            <mergeCell ref="A1:B2"/>
            <mergeCell ref="B2:C3"/>
            <mergeCell ref="C3:D4"/>
            <mergeCell ref="D4:E5"/>
            <mergeCell ref="F1:G1"/>
        </mergeCells>
    "#;
    let wb = parse_workbook(&package(body, &[], false)).unwrap();
    let merges: Vec<_> = wb.sheets[0].merges.iter().map(|m| m.to_a1()).collect();

    assert_eq!(merges, ["A1:B2", "C3:D4", "F1:G1"]);
}

#[test]
fn non_overlapping_merges_are_byte_identical_after_parsing() {
    let mut wb = Workbook::default();
    let mut sheet = xlsx_model::Sheet::new("Sheet1");
    sheet.merges = ["A1:B2", "D3:E4", "G5:H6"]
        .into_iter()
        .map(|range| xlsx_model::CellRange::parse_a1(range).unwrap())
        .collect();
    wb.sheets.push(sheet);
    let parts = serialize_workbook(&wb).unwrap();

    let parsed = parse_workbook(&parts).unwrap();
    let serialized = serialize_workbook(&parsed).unwrap();

    assert_eq!(parts, serialized);
}

/// comparable projection of a workbook's observable shape.
type Snapshot = (
    Vec<(
        String,
        Vec<(String, Cell)>,
        Vec<String>,
        Vec<(u32, f64)>,
        Vec<(u32, f64)>,
    )>,
    DateSystem,
    Vec<String>,
    Vec<DefinedName>,
);

fn snapshot(wb: &Workbook) -> Snapshot {
    let sheets = wb
        .sheets
        .iter()
        .map(|s| {
            let cells = s
                .iter_cells()
                .map(|(a, c)| (a.to_a1(), c.clone()))
                .collect();
            let merges = s.merges.iter().map(|m| m.to_a1()).collect();
            let widths = s.col_widths.iter().map(|(&k, &v)| (k, v)).collect();
            let heights = s.row_heights.iter().map(|(&k, &v)| (k, v)).collect();
            (s.name.clone(), cells, merges, widths, heights)
        })
        .collect();
    (
        sheets,
        wb.date_system,
        wb.shared_strings.clone(),
        wb.defined_names.clone(),
    )
}

#[test]
fn full_circle_parse_serialize_parse_is_stable() {
    let body = r#"
        <cols><col min="1" max="1" width="9"/></cols>
        <sheetData>
            <row r="1" ht="18">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
                <c r="C1"><f>A1</f><v>7.5</v></c>
                <c r="D1" t="b"><v>0</v></c>
                <c r="E1" t="e"><v>#N/A</v></c>
                <c r="F1" t="inlineStr"><is><t>loose text</t></is></c>
            </row>
            <row r="3"><c r="A3" t="s"><v>1</v></c></row>
        </sheetData>
        <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
    "#;
    let wb1 = parse_workbook(&package(body, &["shared one", "shared two"], true)).unwrap();

    let reparts = serialize_workbook(&wb1).unwrap();
    let wb2 = parse_workbook(&reparts).unwrap();

    assert_eq!(snapshot(&wb1), snapshot(&wb2));

    let names: Vec<&str> = reparts.iter().map(|(n, _)| n.as_str()).collect();
    assert!(names.contains(&"[Content_Types].xml"));
    assert!(names.contains(&"_rels/.rels"));
    assert!(names.contains(&"xl/workbook.xml"));
    assert!(names.contains(&"xl/_rels/workbook.xml.rels"));
    assert!(names.contains(&"xl/sharedStrings.xml"));
    assert!(names.contains(&"xl/worksheets/sheet1.xml"));
}

#[test]
fn serialize_round_trips_inline_text_without_shared_table() {
    let mut wb = Workbook::default();
    let mut sheet = xlsx_model::Sheet::new("Only");
    sheet.set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            value: CellValue::Text {
                value: "no table".into(),
            },
            formula: None,
            style: None,
        },
    );
    wb.sheets.push(sheet);

    let parts = serialize_workbook(&wb).unwrap();
    assert!(!parts.iter().any(|(n, _)| n == "xl/sharedStrings.xml"));
    let wb2 = parse_workbook(&parts).unwrap();
    assert_eq!(
        cell_at(&wb2, "A1").value,
        CellValue::Text {
            value: "no table".into()
        }
    );
}

/// wrap a styles inner-body in `<styleSheet>` and attach it (plus an optional
/// theme part) to a bare one-sheet package.
fn package_styled(
    worksheet_body: &str,
    styles_inner: Option<&str>,
    theme: Option<&str>,
) -> Vec<(String, Vec<u8>)> {
    let mut parts = package(worksheet_body, &[], false);
    if let Some(s) = styles_inner {
        let doc = format!("<styleSheet>{s}</styleSheet>");
        parts.push(("xl/styles.xml".to_string(), doc.into_bytes()));
    }
    if let Some(t) = theme {
        parts.push(("xl/theme/theme1.xml".to_string(), t.as_bytes().to_vec()));
    }
    parts
}

/// a full styles fixture exercising every pool, including the gray125
/// convention fill.
const STYLED: &str = r#"
    <numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot;%&quot;"/></numFmts>
    <fonts count="2">
        <font><sz val="11"/><name val="Calibri"/></font>
        <font><b/><sz val="12"/><color theme="4" tint="-0.25"/><name val="Arial"/></font>
    </fonts>
    <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
    </fills>
    <borders count="2">
        <border><left/><right/><top/><bottom/><diagonal/></border>
        <border>
            <left style="thin"><color rgb="FF000000"/></left>
            <right style="thin"/>
            <top style="medium"/>
            <bottom style="double"/>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0"
            applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
            <alignment horizontal="center" vertical="center" wrapText="1"/>
        </xf>
    </cellXfs>
"#;

#[test]
fn parses_full_styled_workbook() {
    let body = r#"<sheetData><row r="1"><c r="A1" s="1"><v>3.5</v></c></row></sheetData>"#;
    let wb = parse_workbook(&package_styled(body, Some(STYLED), None)).unwrap();
    let ss = &wb.styles;

    assert_eq!(cell_at(&wb, "A1").style, Some(1));

    assert_eq!(ss.num_fmts, vec![(164u16, "0.0\"%\"".to_string())]);
    assert_eq!(ss.format_code_for(1), FormatCode::Custom("0.0\"%\""));

    let font = ss.font_for(1).unwrap();
    assert!(font.bold);
    assert_eq!(font.size_pt, Some(12.0));
    assert_eq!(font.name.as_deref(), Some("Arial"));
    assert_eq!(
        font.color,
        Some(Color::Theme {
            idx: 4,
            tint: -0.25
        })
    );
    // accent1 #4472C4 darkened 25% -> excel's 2F5597
    assert_eq!(
        font.color.as_ref().unwrap().resolve(&ss.theme).as_deref(),
        Some("#2f5597")
    );

    assert_eq!(
        ss.fill_for(1),
        Some(&Fill::Solid(Color::Rgb("#ffff00".into())))
    );
    // the gray125 convention fill collapses to a solid auto fill
    assert_eq!(ss.fills[1], Fill::Solid(Color::Auto));

    let border = ss.border_for(1).unwrap();
    let left = border.left.as_ref().unwrap();
    assert_eq!(left.style, BorderStyle::Thin);
    assert_eq!(left.color, Some(Color::Rgb("#000000".into())));
    assert_eq!(border.right.as_ref().unwrap().style, BorderStyle::Thin);
    assert!(border.right.as_ref().unwrap().color.is_none());
    assert_eq!(border.top.as_ref().unwrap().style, BorderStyle::Medium);
    assert_eq!(border.bottom.as_ref().unwrap().style, BorderStyle::Double);

    let align = ss.alignment_for(1).unwrap();
    assert_eq!(align.h, Some(HAlign::Center));
    assert_eq!(align.v, Some(VAlign::Center));
    assert!(align.wrap_text);

    assert!(ss.font_for(0).is_none());
    assert!(ss.fill_for(0).is_none());
    assert_eq!(ss.format_code_for(0), FormatCode::Builtin(0));
}

#[test]
fn resolves_custom_theme_and_indexed_colors() {
    let theme = r#"
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements><a:clrScheme name="Custom">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2>
            <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
            <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
            <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
            <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
            <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
            <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
          </a:clrScheme></a:themeElements>
        </a:theme>
    "#;
    let styles = r#"
        <fonts count="1"><font><color theme="4" tint="0"/></font></fonts>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1"/></cellXfs>
    "#;
    let body = r#"<sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData>"#;
    let wb = parse_workbook(&package_styled(body, Some(styles), Some(theme))).unwrap();

    assert_eq!(wb.styles.theme.slot(4), Some("#ff0000"));
    let font = wb.styles.font_for(0).unwrap();
    assert_eq!(
        font.color
            .as_ref()
            .unwrap()
            .resolve(&wb.styles.theme)
            .as_deref(),
        Some("#ff0000")
    );
    assert_eq!(wb.styles.theme.colors[0], "#000000");
    assert_eq!(
        Color::Indexed(2).resolve(&wb.styles.theme).as_deref(),
        Some("#ff0000")
    );
}

#[test]
fn missing_styles_yields_default_stylesheet() {
    let body = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
    let wb = parse_workbook(&package(body, &[], false)).unwrap();
    assert!(wb.styles.is_empty());
    assert_eq!(wb.styles.theme.slot(4), Some("#4472c4"));
}

#[test]
fn rejects_style_pool_over_cap() {
    let over = crate::MAX_STYLE_ENTRIES + 1;
    let fonts = format!("<fonts count=\"{over}\">{}</fonts>", "<font/>".repeat(over));
    let body = "<sheetData/>";
    let err = parse_workbook(&package_styled(body, Some(&fonts), None)).unwrap_err();
    assert_eq!(err, ParseError::TooManyStyles);
}

#[test]
fn full_circle_styles_round_trip() {
    let body = r#"<sheetData><row r="1"><c r="A1" s="1"><v>3.5</v></c></row></sheetData>"#;
    let wb1 = parse_workbook(&package_styled(body, Some(STYLED), None)).unwrap();

    let reparts = serialize_workbook(&wb1).unwrap();
    let wb2 = parse_workbook(&reparts).unwrap();

    assert_eq!(wb1.styles, wb2.styles);
    assert_eq!(cell_at(&wb2, "A1").style, Some(1));

    let names: Vec<&str> = reparts.iter().map(|(n, _)| n.as_str()).collect();
    assert!(names.contains(&"xl/styles.xml"));
    assert!(names.contains(&"xl/theme/theme1.xml"));

    let ct = reparts
        .iter()
        .find(|(n, _)| n == "[Content_Types].xml")
        .map(|(_, b)| String::from_utf8_lossy(b))
        .unwrap();
    assert!(ct.contains("/xl/styles.xml"));
    assert!(ct.contains("/xl/theme/theme1.xml"));
}
