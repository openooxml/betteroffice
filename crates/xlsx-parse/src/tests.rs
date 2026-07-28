//! fixtures are raw xml parts written from the ecma-376 spec. the parser
//! matches local element names, so fixtures omit namespace declarations.

use xlsx_model::styles::{BorderStyle, Color, Fill, FormatCode, HAlign, VAlign};
use xlsx_model::{
    Cell, CellRef, CellValue, DateSystem, DefinedName, ErrorValue, FreezePane, Hyperlink, SheetId,
    Workbook,
};

use crate::write::{serialize_workbook_with_package, serialize_workbook_with_package_and_origins};
use crate::{
    ParseError, SaveEdits, SharedStringCells, parse_workbook, parse_workbook_with_package,
    serialize_workbook,
};

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
fn hidden_rows_and_columns_have_zero_render_extent() {
    let body = r#"
        <cols>
            <col min="2" max="3" width="12.5" hidden="1"/>
            <col min="4" max="4" hidden="1"/>
        </cols>
        <sheetData>
            <row r="2" ht="30" hidden="1"/>
            <row r="3" hidden="1"/>
        </sheetData>
    "#;
    let wb = parse_workbook(&package(body, &[], false)).unwrap();
    let sheet = &wb.sheets[0];
    assert_eq!(sheet.col_widths.get(&1), Some(&0.0));
    assert_eq!(sheet.col_widths.get(&2), Some(&0.0));
    assert_eq!(sheet.col_widths.get(&3), Some(&0.0));
    assert_eq!(sheet.row_heights.get(&1), Some(&0.0));
    assert_eq!(sheet.row_heights.get(&2), Some(&0.0));
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
fn parses_and_round_trips_frozen_sheet_views() {
    let body = r#"
        <sheetViews>
            <sheetView workbookViewId="0">
                <pane xSplit="2" ySplit="3" topLeftCell="E8" activePane="bottomRight" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData/>
    "#;
    let parsed = parse_workbook(&package(body, &[], false)).unwrap();
    assert_eq!(
        parsed.sheets[0].freeze_pane,
        Some(FreezePane::new(3, 2, CellRef::parse_a1("E8").unwrap()))
    );

    let reparsed = parse_workbook(&serialize_workbook(&parsed).unwrap()).unwrap();
    assert_eq!(reparsed.sheets[0].freeze_pane, parsed.sheets[0].freeze_pane);
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
fn parses_and_round_trips_external_and_internal_hyperlinks() {
    let body = r#"
        <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>Website</t></is></c></row>
        </sheetData>
        <hyperlinks>
            <hyperlink ref="A1:B1" r:id="rId7" tooltip="Open site" display="Website"/>
            <hyperlink ref="C3" location="'Other Sheet'!$D$4" display="Jump"/>
        </hyperlinks>
    "#;
    let mut parts = package(body, &[], false);
    parts.push((
        "xl/worksheets/_rels/sheet1.xml.rels".into(),
        br#"
            <Relationships>
                <Relationship Id="rId7"
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                    Target="https://example.com/report?q=1&amp;lang=en"
                    TargetMode="External"/>
            </Relationships>
        "#
        .to_vec(),
    ));

    let parsed = parse_workbook(&parts).unwrap();
    assert_eq!(
        parsed.sheets[0].hyperlinks,
        vec![
            Hyperlink {
                range: xlsx_model::CellRange::parse_a1("A1:B1").unwrap(),
                external_target: Some("https://example.com/report?q=1&lang=en".into()),
                location: None,
                tooltip: Some("Open site".into()),
                display: Some("Website".into()),
            },
            Hyperlink {
                range: xlsx_model::CellRange::parse_a1("C3").unwrap(),
                external_target: None,
                location: Some("'Other Sheet'!$D$4".into()),
                tooltip: None,
                display: Some("Jump".into()),
            },
        ]
    );

    let serialized = serialize_workbook(&parsed).unwrap();
    assert!(
        serialized
            .iter()
            .any(|(name, _)| name == "xl/worksheets/_rels/sheet1.xml.rels")
    );
    let reparsed = parse_workbook(&serialized).unwrap();
    assert_eq!(reparsed.sheets[0].hyperlinks, parsed.sheets[0].hyperlinks);
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
        Option<FreezePane>,
        Vec<Hyperlink>,
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
            (
                s.name.clone(),
                cells,
                merges,
                widths,
                heights,
                s.freeze_pane,
                s.hyperlinks.clone(),
            )
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

#[test]
fn preserved_shared_strings_keep_rich_items_and_replace_only_changed_indices() {
    let rich_item = r#"<si><r><rPr><b/></rPr><t>Rich </t></r><r><rPr><i/></rPr><t>Text</t></r><phoneticPr fontId="2"/></si>"#;
    let sst = format!(
        r#"<sst count="2" uniqueCount="2">{rich_item}<si><t>plain</t></si><extLst><ext uri="{{fixture}}"/></extLst></sst>"#
    );
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.as_bytes().to_vec()));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let unchanged = serialize_workbook_with_package(&parsed.workbook, &parsed.package).unwrap();
    assert_eq!(unchanged, parts);

    let mut workbook = parsed.workbook;
    workbook.shared_strings[1] = "changed".to_owned();
    let changed = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let changed_sst = String::from_utf8(
        changed
            .iter()
            .find(|(path, _)| path == "xl/sharedStrings.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert!(changed_sst.contains(rich_item));
    assert!(changed_sst.contains("<si><t xml:space=\"preserve\">changed</t></si>"));
    assert!(!changed_sst.contains("<si><t>plain</t></si>"));
    assert!(changed_sst.contains("<extLst>"));
}

/// The rich `<si>` is retained by its parsed value, so it must survive every
/// way the serializer can move it to another index.
#[test]
fn preserved_shared_strings_follow_rich_items_through_index_moves() {
    let rich_item = r#"<si><r><rPr><b/></rPr><t>Rich </t></r><r><rPr><i/></rPr><t>Text</t></r><phoneticPr fontId="2"/></si>"#;
    let sst = format!(
        r#"<sst count="3" uniqueCount="3"><si><t>first</t></si>{rich_item}<si><t>last</t></si></sst>"#
    );
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.as_bytes().to_vec()));
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let moves: [(&str, Vec<&str>); 4] = [
        ("insert", vec!["added", "first", "Rich Text", "last"]),
        ("delete", vec!["first", "Rich Text"]),
        ("reorder", vec!["last", "Rich Text", "first"]),
        ("both", vec!["Rich Text", "added"]),
    ];
    for (label, strings) in moves {
        let mut workbook = parsed.workbook.clone();
        workbook.shared_strings = strings.iter().map(|s| (*s).to_owned()).collect();
        let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
        let written = shared_strings_text(&saved);
        assert!(written.contains(rich_item), "{label} lost the rich item");
        assert_eq!(
            written.matches("<si>").count(),
            strings.len(),
            "{label} wrote the wrong item count"
        );
    }

    let mut workbook = parsed.workbook.clone();
    workbook.shared_strings[1] = "Rich Prose".to_owned();
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = shared_strings_text(&saved);
    assert!(!written.contains("<r>"), "edited rich item must regenerate");
    assert!(written.contains("<si><t xml:space=\"preserve\">Rich Prose</t></si>"));
}

/// Duplicate plain values claim the source items in order, so an item shifted
/// past a same-valued sibling keeps its own markup.
#[test]
fn preserved_shared_strings_pair_duplicate_values_in_order() {
    let rich_item = r#"<si><r><rPr><b/></rPr><t>Dup</t></r></si>"#;
    let sst = format!(r#"<sst count="2" uniqueCount="2">{rich_item}<si><t>Dup</t></si></sst>"#);
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.as_bytes().to_vec()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(
        serialize_workbook_with_package(&parsed.workbook, &parsed.package).unwrap(),
        parts
    );

    let mut workbook = parsed.workbook.clone();
    workbook.shared_strings.insert(0, "added".to_owned());
    let written =
        shared_strings_text(&serialize_workbook_with_package(&workbook, &parsed.package).unwrap());
    assert!(written.contains(&format!("{rich_item}<si><t>Dup</t></si>")));
}

#[test]
fn new_duplicate_shared_strings_do_not_claim_authored_rich_items() {
    let bold = r#"<si><r><rPr><b/></rPr><t>Dup</t></r></si>"#;
    let italic = r#"<si><r><rPr><i/></rPr><t>Dup</t></r></si>"#;
    let sst = format!(r#"<sst count="2" uniqueCount="1">{bold}{italic}</sst>"#);
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.into_bytes()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook;
    workbook.shared_strings.insert(0, "Dup".to_owned());
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("C1").unwrap(),
        Cell {
            value: CellValue::Text {
                value: "Dup".into(),
            },
            ..Cell::default()
        },
    );

    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let strings = shared_strings_text(&saved);
    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(strings.contains(&format!(
        r#"<si><t xml:space="preserve">Dup</t></si>{bold}{italic}"#
    )));
    assert!(sheet.contains(r#"<c r="A1" t="s"><v>1</v></c>"#));
    assert!(sheet.contains(r#"<c r="B1" t="s"><v>2</v></c>"#));
    assert!(sheet.contains(r#"<c r="C1" t="s"><v>0</v></c>"#));
}

#[test]
fn new_duplicates_do_not_consume_previously_unique_rich_items() {
    let rich = r#"<si><r><rPr><b/></rPr><t>Dup</t></r></si>"#;
    let sst = format!(r#"<sst count="1" uniqueCount="1">{rich}</sst>"#);
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.into_bytes()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook;
    workbook.shared_strings.insert(0, "Dup".to_owned());
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("B1").unwrap(),
        Cell {
            value: CellValue::Text {
                value: "Dup".into(),
            },
            ..Cell::default()
        },
    );

    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let strings = shared_strings_text(&saved);
    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(strings.contains(&format!(
        r#"<si><t xml:space="preserve">Dup</t></si>{rich}"#
    )));
    assert!(sheet.contains(r#"<c r="A1" t="s"><v>1</v></c>"#));
    assert!(sheet.contains(r#"<c r="B1" t="s"><v>0</v></c>"#));
}

#[test]
fn new_duplicate_cells_without_table_entries_use_inline_text() {
    let bold = r#"<si><r><rPr><b/></rPr><t>Dup</t></r></si>"#;
    let italic = r#"<si><r><rPr><i/></rPr><t>Dup</t></r></si>"#;
    let sst = format!(r#"<sst count="2" uniqueCount="1">{bold}{italic}</sst>"#);
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.into_bytes()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook;
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("C1").unwrap(),
        Cell {
            value: CellValue::Text {
                value: "Dup".into(),
            },
            ..Cell::default()
        },
    );

    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let strings = shared_strings_text(&saved);
    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(strings.contains(&format!("{bold}{italic}")));
    assert!(sheet.contains(r#"<c r="A1" t="s"><v>0</v></c>"#));
    assert!(sheet.contains(r#"<c r="B1" t="s"><v>1</v></c>"#));
    assert!(
        sheet.contains(r#"<c r="C1" t="inlineStr"><is><t xml:space="preserve">Dup</t></is></c>"#)
    );
}

#[test]
fn duplicate_removal_keeps_the_entry_still_used_by_a_cell() {
    let bold = r#"<si><r><rPr><b/></rPr><t>Dup</t></r></si>"#;
    let italic = r#"<si><r><rPr><i/></rPr><t>Dup</t></r></si>"#;
    let sst = format!(r#"<sst count="2" uniqueCount="1">{bold}{italic}</sst>"#);
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.into_bytes()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook;
    workbook.shared_strings.pop();

    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let strings = shared_strings_text(&saved);
    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(!strings.contains("<b/>"));
    assert!(strings.contains(italic));
    assert!(sheet.contains(r#"<c r="A1" t="s"><v>0</v></c>"#));
}

#[test]
fn ambiguous_duplicate_removal_is_refused() {
    let bold = r#"<si><r><rPr><b/></rPr><t>Dup</t></r></si>"#;
    let italic = r#"<si><r><rPr><i/></rPr><t>Dup</t></r></si>"#;
    let sst = format!(r#"<sst count="2" uniqueCount="1">{bold}{italic}</sst>"#);
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.into_bytes()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook;
    workbook.shared_strings.pop();

    let error = serialize_workbook_with_package(&workbook, &parsed.package).unwrap_err();

    assert!(matches!(error, ParseError::UnsupportedEdit(_)));
}

/// Two worksheets sharing one workbook, so an edit to the second can be
/// checked against the first.
fn two_sheet_package(first_body: &str, second_body: &str) -> Vec<(String, Vec<u8>)> {
    let workbook = r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
    let rels = r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>"#;
    vec![
        ("xl/workbook.xml".to_owned(), workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            rels.as_bytes().to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_owned(),
            format!("<worksheet>{first_body}</worksheet>").into_bytes(),
        ),
        (
            "xl/worksheets/sheet2.xml".to_owned(),
            format!("<worksheet>{second_body}</worksheet>").into_bytes(),
        ),
    ]
}

fn part_bytes(parts: &[(String, Vec<u8>)], path: &str) -> Vec<u8> {
    parts
        .iter()
        .find(|(name, _)| name == path)
        .unwrap_or_else(|| panic!("missing {path}"))
        .1
        .clone()
}

/// The parser models a subset of row, column and cell markup. An edit to one
/// sheet must not push every other sheet through that lossy round-trip.
#[test]
fn leaves_untouched_worksheets_byte_identical() {
    let untouched = concat!(
        r#"<sheetPr><outlinePr summaryBelow="0"/></sheetPr>"#,
        r#"<cols><col min="1" max="1" width="12" hidden="1" outlineLevel="1"/></cols>"#,
        r#"<sheetData>"#,
        r#"<row r="1" hidden="1" outlineLevel="2" collapsed="1" s="3" customFormat="1">"#,
        r#"<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>Rich</t></r><r><t> mix</t></r></is></c>"#,
        r#"<c r="B1"><f t="shared" si="0" ref="B1:B2">SUM(A1:A2)</f><v>3</v></c>"#,
        r#"</row>"#,
        r#"<row r="2"><c r="B2"><f t="shared" si="0"/><v>3</v></c></row>"#,
        r#"</sheetData>"#,
    );
    let parts = two_sheet_package(untouched, r#"<sheetData/>"#);
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let mut workbook = parsed.workbook.clone();
    workbook.sheets[1].set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            value: CellValue::Number { value: 7.0 },
            ..Cell::default()
        },
    );
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();

    assert_eq!(
        String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap(),
        String::from_utf8(part_bytes(&parts, "xl/worksheets/sheet1.xml")).unwrap(),
    );
    assert!(
        String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet2.xml"))
            .unwrap()
            .contains("<v>7</v>")
    );
}

/// A rename touches only the workbook part, so worksheet bytes still survive.
#[test]
fn keeps_worksheet_bytes_across_a_rename() {
    let body = r#"<sheetData><row r="1" hidden="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
    let parts = two_sheet_package(body, r#"<sheetData/>"#);
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].name = "Renamed".to_owned();
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();

    assert_eq!(
        part_bytes(&saved, "xl/worksheets/sheet1.xml"),
        part_bytes(&parts, "xl/worksheets/sheet1.xml")
    );
    assert!(
        String::from_utf8(part_bytes(&saved, "xl/workbook.xml"))
            .unwrap()
            .contains(r#"name="Renamed""#)
    );
}

/// Several local `_xlnm.Print_Area` entries are normal, and the model has no
/// source identity to tell them apart. Dropping one must not hand its markup
/// to the survivor.
#[test]
fn does_not_reattach_duplicate_defined_name_markup() {
    let workbook_xml = concat!(
        r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>"#,
        r#"<definedNames>"#,
        r#"<definedName name="_xlnm.Print_Area" localSheetId="0" comment="first">Sheet1!$A$1</definedName>"#,
        r#"<definedName name="_xlnm.Print_Area" localSheetId="1" comment="second">Sheet1!$B$1</definedName>"#,
        r#"</definedNames></workbook>"#,
    );
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        workbook_xml.as_bytes().to_vec(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.defined_names.remove(0);
    workbook.defined_names[0].local_sheet = Some(SheetId(0));
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = String::from_utf8(part_bytes(&saved, "xl/workbook.xml")).unwrap();

    assert_eq!(written.matches("<definedName ").count(), 1, "{written}");
    assert!(
        !written.contains(r#"comment="first""#),
        "the removed entry's markup was reattached: {written}"
    );
    assert!(written.contains("Sheet1!$B$1"), "{written}");
}

/// A long root prefix used to be repeated on every generated element, turning
/// a bounded input into quadratic output.
#[test]
fn binds_generated_fragments_once_instead_of_repeating_a_long_prefix() {
    let prefix = "p".repeat(4096);
    let main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[2] = (
        "xl/worksheets/sheet1.xml".to_owned(),
        format!(
            r#"<{prefix}:worksheet xmlns:{prefix}="{main}"><{prefix}:sheetData/></{prefix}:worksheet>"#
        )
        .into_bytes(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    for row in 0..256 {
        workbook.sheets[0].set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
    }
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = part_bytes(&saved, "xl/worksheets/sheet1.xml");

    assert!(
        written.len() < 32 * 1024,
        "generated worksheet grew to {} bytes",
        written.len()
    );
    assert_eq!(
        parse_workbook(&saved).unwrap().sheets[0]
            .iter_cells()
            .count(),
        256
    );
}

const MCE_NAMESPACES: &str = concat!(
    r#" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
    r#" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
    r#" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""#,
);

fn mce_package(worksheet_body: &str) -> Vec<(String, Vec<u8>)> {
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[2] = (
        "xl/worksheets/sheet1.xml".to_owned(),
        format!("<worksheet{MCE_NAMESPACES}>{worksheet_body}</worksheet>").into_bytes(),
    );
    parts
}

/// An `mc:AlternateContent` branch stands in for the element inside it, so a
/// newly inserted sibling has to be ordered against that element's rank.
#[test]
fn orders_new_children_against_alternate_content_branches() {
    let body = concat!(
        r#"<sheetData/>"#,
        r#"<mc:AlternateContent><mc:Choice Requires="x14">"#,
        r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>TRUE()</formula></cfRule></conditionalFormatting>"#,
        r#"</mc:Choice></mc:AlternateContent>"#,
        r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
    );
    let parsed = parse_workbook_with_package(&mce_package(body)).unwrap();

    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0]
        .merges
        .push(xlsx_model::CellRange::parse_a1("D1:E1").unwrap());
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    let merges = written.find("<mergeCells").unwrap();
    let branch = written.find("<mc:AlternateContent").unwrap();
    let margins = written.find("<pageMargins").unwrap();
    assert!(
        merges < branch && branch < margins,
        "mergeCells must precede the conditionalFormatting branch: {written}"
    );
}

/// Patching inside a compatibility branch is out of reach, so a save that
/// would duplicate an owned singleton is refused instead.
#[test]
fn refuses_to_duplicate_a_singleton_hidden_in_a_branch() {
    let body = concat!(
        r#"<mc:AlternateContent><mc:Choice Requires="x14">"#,
        r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
        r#"</mc:Choice></mc:AlternateContent>"#,
    );
    let parsed = parse_workbook_with_package(&mce_package(body)).unwrap();

    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 9.0);
    let error = serialize_workbook_with_package(&workbook, &parsed.package).unwrap_err();

    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message) if message.contains("sheetData")),
        "{error:?}"
    );
}

/// Freeze-pane edits used to vanish, because the retained-sheet renderer never
/// replaced `sheetViews`.
#[test]
fn overlays_freeze_panes_onto_retained_sheet_views() {
    let body = r#"<sheetViews><sheetView tabSelected="1" zoomScale="120" workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView></sheetViews><sheetData/>"#;
    let parts = package(body, &[], false);
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].freeze_pane = Some(FreezePane::new(1, 2, CellRef::parse_a1("C2").unwrap()));
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(written.contains(r#"state="frozen""#), "{written}");
    assert!(written.contains(r#"zoomScale="120""#), "{written}");
    assert!(
        written.contains(r#"<selection activeCell="C3""#),
        "{written}"
    );
    assert!(
        written.find("<pane").unwrap() < written.find("<selection").unwrap(),
        "pane must precede selection: {written}"
    );

    let reparsed = parse_workbook(&saved).unwrap();
    assert_eq!(
        reparsed.sheets[0].freeze_pane,
        workbook.sheets[0].freeze_pane
    );
}

#[test]
fn removes_the_pane_when_a_sheet_is_unfrozen() {
    let body = r#"<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetData/>"#;
    let parts = package(body, &[], false);
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert!(parsed.workbook.sheets[0].freeze_pane.is_some());

    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].freeze_pane = None;
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(!written.contains("<pane"), "{written}");
    assert!(written.contains("<selection"), "{written}");
    assert!(
        parse_workbook(&saved).unwrap().sheets[0]
            .freeze_pane
            .is_none()
    );
}

/// Hyperlink edits must reach both the worksheet and its relationship part,
/// without disturbing the drawings and comments living in the same part.
#[test]
fn overlays_hyperlinks_and_merges_the_worksheet_relationships() {
    let body = r#"<sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks><drawing r:id="rId2"/>"#;
    let mut parts = package(body, &[], false);
    parts.push((
        "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
        br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec(),
    ));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].hyperlinks = vec![Hyperlink {
        range: xlsx_model::CellRange::parse_a1("B2").unwrap(),
        external_target: Some("https://new.example/".to_owned()),
        location: None,
        tooltip: None,
        display: None,
    }];
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();
    let rels =
        String::from_utf8(part_bytes(&saved, "xl/worksheets/_rels/sheet1.xml.rels")).unwrap();

    assert!(sheet.contains(r#"ref="B2""#), "{sheet}");
    assert!(sheet.contains(r#"<drawing r:id="rId2"/>"#), "{sheet}");
    assert!(rels.contains("../drawings/drawing1.xml"), "{rels}");
    assert!(rels.contains("https://new.example/"), "{rels}");
    assert!(!rels.contains("https://old.example"), "{rels}");

    let id = sheet
        .split_once("r:id=\"")
        .map(|(_, rest)| rest.split_once('"').unwrap().0.to_owned())
        .unwrap();
    assert!(
        rels.contains(&format!(r#"Id="{id}""#)),
        "hyperlink id {id} is not backed by {rels}"
    );
    assert_eq!(
        parse_workbook(&saved).unwrap().sheets[0].hyperlinks,
        workbook.sheets[0].hyperlinks
    );
}

/// A new sheet emitted `r:id` values with no relationship part behind them.
#[test]
fn writes_relationships_for_hyperlinks_on_new_sheets() {
    let parts = package(r#"<sheetData/>"#, &[], false);
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let mut workbook = parsed.workbook.clone();
    let mut added = xlsx_model::Sheet::new("Added");
    added.hyperlinks = vec![Hyperlink {
        range: xlsx_model::CellRange::parse_a1("A1").unwrap(),
        external_target: Some("https://example.test/".to_owned()),
        location: None,
        tooltip: None,
        display: None,
    }];
    workbook.sheets.push(added);
    let saved =
        serialize_workbook_with_package_and_origins(&workbook, &parsed.package, &[Some(0), None])
            .unwrap();

    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet2.xml")).unwrap();
    let rels =
        String::from_utf8(part_bytes(&saved, "xl/worksheets/_rels/sheet2.xml.rels")).unwrap();
    let id = sheet
        .split_once("r:id=\"")
        .map(|(_, rest)| rest.split_once('"').unwrap().0.to_owned())
        .unwrap();
    assert!(rels.contains(&format!(r#"Id="{id}""#)), "{rels}");
    assert_eq!(
        parse_workbook(&saved).unwrap().sheets[1].hyperlinks,
        workbook.sheets[1].hyperlinks
    );
}

/// A Strict package binds `r` to the Strict relationships namespace, so a new
/// sheet must not hard-code the Transitional one.
#[test]
fn binds_new_strict_sheets_to_the_strict_relationship_namespace() {
    let workbook_xml = r#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        workbook_xml.as_bytes().to_vec(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    let mut added = xlsx_model::Sheet::new("Added");
    added.hyperlinks = vec![Hyperlink {
        range: xlsx_model::CellRange::parse_a1("A1").unwrap(),
        external_target: Some("https://example.test/".to_owned()),
        location: None,
        tooltip: None,
        display: None,
    }];
    workbook.sheets.push(added);
    let saved =
        serialize_workbook_with_package_and_origins(&workbook, &parsed.package, &[Some(0), None])
            .unwrap();

    let sheet = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet2.xml")).unwrap();
    assert!(
        sheet.contains(r#"xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships""#),
        "{sheet}"
    );
}

fn edit_a1(workbook: &mut Workbook, value: f64) {
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            value: CellValue::Number { value },
            ..Cell::default()
        },
    );
}

/// The model carries a subset of the stylesheet, so an edit that does not
/// touch styles must not push the part through that subset.
#[test]
fn keeps_the_stylesheet_when_styles_are_untouched() {
    let parts = package_styled(r#"<sheetData/>"#, Some(STYLED), None);
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 5.0);
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();

    assert_eq!(
        part_bytes(&saved, "xl/styles.xml"),
        part_bytes(&parts, "xl/styles.xml")
    );
}

/// A stylesheet holding only unmodeled pools still backs `dxfId` references
/// from preserved conditional formatting; deleting it breaks the workbook.
#[test]
fn keeps_a_stylesheet_that_models_nothing() {
    let styles = r#"<dxfs count="1"><dxf><font><b/></font></dxf></dxfs><tableStyles count="0"/>"#;
    let body = r#"<sheetData/><conditionalFormatting sqref="A1:A9"><cfRule type="expression" dxfId="0" priority="1"><formula>TRUE()</formula></cfRule></conditionalFormatting>"#;
    let parts = package_styled(body, Some(styles), None);
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert!(parsed.workbook.styles.is_empty());

    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 5.0);
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();

    assert_eq!(
        part_bytes(&saved, "xl/styles.xml"),
        part_bytes(&parts, "xl/styles.xml")
    );
    assert!(content_types_text(&saved).contains("/xl/styles.xml"));
}

/// Interning a new format appends one `xf`; every pool entry the model left
/// alone must keep its source markup.
#[test]
fn patches_only_the_style_pool_entries_that_changed() {
    let parts = package_styled(r#"<sheetData/>"#, Some(STYLED), None);
    let parsed = parse_workbook_with_package(&parts).unwrap();

    let mut workbook = parsed.workbook.clone();
    let mut format = workbook.styles.cell_format(None);
    format.font.italic = true;
    let style = workbook.styles.intern_cell_format(&format);
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            value: CellValue::Number { value: 1.0 },
            style,
            ..Cell::default()
        },
    );
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = String::from_utf8(part_bytes(&saved, "xl/styles.xml")).unwrap();

    assert!(
        written.contains(r#"<patternFill patternType="gray125"/>"#),
        "lost the gray125 convention fill: {written}"
    );
    assert!(
        written.contains(r#"<bgColor indexed="64"/>"#),
        "lost an unmodeled fill child: {written}"
    );
    assert!(
        written.contains(r#"<numFmt numFmtId="164" formatCode="0.0&quot;%&quot;"/>"#),
        "lost the source number format markup: {written}"
    );
    assert!(
        written.contains("<i/>"),
        "new font was not written: {written}"
    );
}

/// A Strict package must not gain a Transitional DrawingML theme.
#[test]
fn writes_a_strict_theme_for_a_strict_package() {
    let workbook_xml = r#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        workbook_xml.as_bytes().to_vec(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.styles.fonts.push(xlsx_model::styles::Font {
        bold: true,
        ..Default::default()
    });
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let theme = String::from_utf8(part_bytes(&saved, "xl/theme/theme1.xml")).unwrap();

    assert!(
        theme.contains(r#"xmlns:a="http://purl.oclc.org/ooxml/drawingml/main""#),
        "strict package gained a transitional theme: {theme}"
    );
}

fn content_types_text(parts: &[(String, Vec<u8>)]) -> String {
    String::from_utf8(
        parts
            .iter()
            .find(|(path, _)| path == "[Content_Types].xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap()
}

/// A part typed by `<Default Extension=…>` must keep that type; inventing an
/// override retypes chartsheets and macro-enabled workbooks.
#[test]
fn keeps_content_types_resolved_through_default_extensions() {
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts.push((
        "[Content_Types].xml".to_owned(),
        br#"<Types><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/vnd.ms-excel.worksheet+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>"#.to_vec(),
    ));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("A1").unwrap(),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let types = content_types_text(&saved);

    assert!(types.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"));
    assert!(
        !types.contains("spreadsheetml.worksheet+xml"),
        "worksheet was retyped away from its Default: {types}"
    );
    assert!(
        !types.contains(r#"PartName="/xl/worksheets/sheet1.xml""#),
        "redundant override added over a Default: {types}"
    );
}

/// `saturating_add` handed every new sheet `u32::MAX` once the source used it.
#[test]
fn allocates_distinct_sheet_ids_past_the_maximum() {
    let workbook_xml = format!(
        r#"<workbook><sheets><sheet name="Sheet1" sheetId="{}" r:id="rId1"/></sheets></workbook>"#,
        u32::MAX
    );
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = ("xl/workbook.xml".to_owned(), workbook_xml.into_bytes());

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.sheets.push(xlsx_model::Sheet::new("Added"));
    workbook.sheets.push(xlsx_model::Sheet::new("AlsoAdded"));
    let saved = serialize_workbook_with_package_and_origins(
        &workbook,
        &parsed.package,
        &[Some(0), None, None],
    )
    .unwrap();

    let written = String::from_utf8(
        saved
            .iter()
            .find(|(path, _)| path == "xl/workbook.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    let mut ids = written
        .match_indices("sheetId=\"")
        .map(|(index, needle)| {
            let rest = &written[index + needle.len()..];
            rest[..rest.find('"').unwrap()].to_owned()
        })
        .collect::<Vec<_>>();
    assert_eq!(ids.len(), 3);
    ids.sort();
    ids.dedup();
    assert_eq!(ids.len(), 3, "duplicate sheetId in {written}");
}

/// `<sst/>` and `<styleSheet/>` are schema-valid; capture must treat a
/// self-closing root as an empty template rather than a missing one.
#[test]
fn captures_self_closing_template_roots() {
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts.push((
        "xl/sharedStrings.xml".to_owned(),
        br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec(),
    ));
    parts.push((
        "xl/styles.xml".to_owned(),
        br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    ));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert!(parsed.workbook.shared_strings.is_empty());

    let mut workbook = parsed.workbook.clone();
    workbook.shared_strings = vec!["added".to_owned()];
    let written =
        shared_strings_text(&serialize_workbook_with_package(&workbook, &parsed.package).unwrap());
    assert!(written.starts_with("<sst "));
    assert!(written.contains(r#"<si><t xml:space="preserve">added</t></si>"#));
    assert!(written.ends_with("</sst>"));
}

/// Two `<si>` entries can carry the same text with different runs. A cell must
/// keep pointing at the entry it was authored against.
#[test]
fn keeps_cells_on_their_own_shared_string_entry() {
    let rich = r#"<si><r><rPr><b/></rPr><t>Total</t></r></si>"#;
    let sst = format!(r#"<sst count="2" uniqueCount="1"><si><t>Total</t></si>{rich}</sst>"#);
    let body = r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>"#;
    let mut parts = package(body, &[], false);
    parts.push(("xl/sharedStrings.xml".to_owned(), sst.into_bytes()));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("C1").unwrap(),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    let written = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet1.xml")).unwrap();

    assert!(
        written.contains(r#"<c r="A1" t="s"><v>0</v></c>"#),
        "{written}"
    );
    assert!(
        written.contains(r#"<c r="B1" t="s"><v>1</v></c>"#),
        "B1 was moved onto another entry with the same text: {written}"
    );
}

/// A shared-string part reached through a non-conventional relationship target
/// parsed as empty, so an edited save deleted it and blanked every cell.
#[test]
fn resolves_shared_strings_through_the_workbook_relationship() {
    let mut parts = package(
        r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>"#,
        &[],
        false,
    );
    parts[1] = (
        "xl/_rels/workbook.xml.rels".to_owned(),
        br#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="strings/custom.xml"/></Relationships>"#.to_vec(),
    );
    parts.push((
        "xl/strings/custom.xml".to_owned(),
        br#"<sst count="1" uniqueCount="1"><si><t>Hello</t></si></sst>"#.to_vec(),
    ));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(parsed.workbook.shared_strings, vec!["Hello".to_owned()]);
    assert_eq!(
        cell_at(&parsed.workbook, "A1").value,
        CellValue::Text {
            value: "Hello".into()
        }
    );

    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("B1").unwrap(),
        Cell {
            value: CellValue::Number { value: 2.0 },
            ..Cell::default()
        },
    );
    let saved = serialize_workbook_with_package(&workbook, &parsed.package).unwrap();
    assert_eq!(
        part_bytes(&saved, "xl/strings/custom.xml"),
        part_bytes(&parts, "xl/strings/custom.xml")
    );
    assert_eq!(
        cell_at(&parse_workbook(&saved).unwrap(), "A1").value,
        CellValue::Text {
            value: "Hello".into()
        }
    );
}

fn shared_strings_text(parts: &[(String, Vec<u8>)]) -> String {
    String::from_utf8(
        parts
            .iter()
            .find(|(path, _)| path == "xl/sharedStrings.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap()
}

/// A relationship prefix is source-controlled, so repeating it on every
/// generated sheet and hyperlink amplifies it. Generated attributes use a
/// fixed prefix and bind the source URI on the fragment instead.
#[test]
fn generated_relationship_attributes_never_repeat_a_source_prefix() {
    let prefix = "p".repeat(4096);
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        format!(
            r#"<workbook xmlns:{prefix}="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" {prefix}:id="rId1"/></sheets></workbook>"#
        )
        .into_bytes(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    workbook.sheets.push(xlsx_model::Sheet::new("Added"));
    workbook.sheets[1].hyperlinks.push(Hyperlink {
        range: xlsx_model::CellRange::parse_a1("A1").unwrap(),
        external_target: Some("https://example.invalid/".to_owned()),
        location: None,
        tooltip: None,
        display: None,
    });
    let origins = vec![Some(0), None];
    let saved =
        serialize_workbook_with_package_and_origins(&workbook, &parsed.package, &origins).unwrap();

    let written = String::from_utf8(part_bytes(&saved, "xl/workbook.xml")).unwrap();
    assert_eq!(written.matches(&prefix).count(), 2, "{written}");
    assert!(written.contains(r#"r:id="rId2""#), "{written}");
    assert!(
        written.contains(
            r#"<sheets xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
        ),
        "{written}"
    );

    let added = String::from_utf8(part_bytes(&saved, "xl/worksheets/sheet2.xml")).unwrap();
    assert!(!added.contains(&prefix), "generated worksheet repeats it");
    assert!(added.contains(r#"<hyperlink ref="A1" r:id="#), "{added}");
    assert_eq!(parse_workbook(&saved).unwrap().sheets.len(), 2);
}

/// Every generated fragment carries the source relationship URI once, so an
/// absurd one is refused before any of them is built.
#[test]
fn refuses_an_oversized_relationship_namespace() {
    let namespace = format!(
        "http://example.invalid/{}/officeDocument/relationships",
        "x".repeat(2048)
    );
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        format!(
            r#"<workbook xmlns:r="{namespace}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        )
        .into_bytes(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);
    let error = serialize_workbook_with_package(&workbook, &parsed.package).unwrap_err();

    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message) if message.contains("relationship namespace")),
        "{error:?}"
    );
}

/// A two-sheet package whose first sheet anchors a chart through a drawing,
/// wired the way Excel writes it: sheet rels -> drawing -> drawing rels ->
/// chart part.
fn charted_package() -> Vec<(String, Vec<u8>)> {
    let mut parts = package(r#"<sheetData/><drawing r:id="rIdDrawing"/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        br#"<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets></workbook>"#.to_vec(),
    );
    parts[1] = (
        "xl/_rels/workbook.xml.rels".to_owned(),
        br#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>"#.to_vec(),
    );
    parts.extend([
        (
            "xl/worksheets/sheet2.xml".to_owned(),
            br#"<worksheet><sheetData/></worksheet>"#.to_vec(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/drawings/drawing1.xml".to_owned(), DRAWING.to_vec()),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/charts/chart1.xml".to_owned(), CHART.to_vec()),
    ]);
    parts
}

const DRAWING: &[u8] = br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

const CHART: &[u8] = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:pt idx="0"><c:v>Series</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr><c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

fn pivoted_package() -> Vec<(String, Vec<u8>)> {
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        br#"<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets></workbook>"#.to_vec(),
    );
    parts[1] = (
        "xl/_rels/workbook.xml.rels".to_owned(),
        br#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>"#.to_vec(),
    );
    parts.push((
        "xl/worksheets/sheet2.xml".to_owned(),
        br#"<worksheet><sheetData/></worksheet>"#.to_vec(),
    ));
    parts.push((
        "xl/pivotcache/pivotCacheDefinition1.xml".to_owned(),
        br#"<pivotCacheDefinition><cacheSource><worksheetSource sheet="Data" ref="A1:B4"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    ));
    parts
}

fn unclaimed_chart_package() -> Vec<(String, Vec<u8>)> {
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        br#"<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets></workbook>"#.to_vec(),
    );
    parts[1] = (
        "xl/_rels/workbook.xml.rels".to_owned(),
        br#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>"#.to_vec(),
    );
    parts.push((
        "xl/worksheets/sheet2.xml".to_owned(),
        br#"<worksheet><sheetData/></worksheet>"#.to_vec(),
    ));
    parts.push((
        "xl/charts/chart1.xml".to_owned(),
        br#"<chartSpace><f>Data!$A$1:$A$2</f></chartSpace>"#.to_vec(),
    ));
    parts
}

#[test]
fn refuses_to_strand_pivot_references_at_the_serialization_boundary() {
    let parsed = parse_workbook_with_package(&pivoted_package()).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);

    let provenance = vec![SharedStringCells::new(); 2];
    let reordered = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(1), Some(0)],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap_err();
    assert!(
        matches!(&reordered, ParseError::UnsupportedEdit(message) if message.contains("pivotCacheDefinition1.xml")),
        "{reordered:?}"
    );

    let removed = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), None],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap_err();
    assert!(matches!(removed, ParseError::UnsupportedEdit(_)));

    let mut renamed = workbook.clone();
    renamed.sheets[0].name = "Renamed".to_owned();
    let renamed = crate::serialize_workbook_with_package_and_origins_after_edits(
        &renamed,
        &parsed.package,
        &[Some(0), Some(1)],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap_err();
    assert!(matches!(renamed, ParseError::UnsupportedEdit(_)));

    let save_in_place = |moved_references| {
        crate::serialize_workbook_with_package_and_origins_after_edits(
            &workbook,
            &parsed.package,
            &[Some(0), Some(1)],
            &provenance,
            SaveEdits {
                changed: true,
                moved_references,
            },
        )
    };
    save_in_place(false).expect("a cell edit strands nothing");
    let axis_edit = save_in_place(true).unwrap_err();
    assert!(
        matches!(&axis_edit, ParseError::UnsupportedEdit(message) if message.contains("pivotCacheDefinition1.xml")),
        "{axis_edit:?}"
    );
}

#[test]
fn refuses_to_strand_unclaimed_chart_references_at_the_serialization_boundary() {
    let parsed = parse_workbook_with_package(&unclaimed_chart_package()).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);
    let provenance = vec![SharedStringCells::new(); 2];

    let reordered = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(1), Some(0)],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap_err();
    assert!(
        matches!(&reordered, ParseError::UnsupportedEdit(message) if message.contains("chart1.xml")),
        "{reordered:?}"
    );

    let removed = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), None],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap_err();
    assert!(matches!(removed, ParseError::UnsupportedEdit(_)));

    let mut renamed = workbook.clone();
    renamed.sheets[0].name = "Renamed".to_owned();
    let renamed = crate::serialize_workbook_with_package_and_origins_after_edits(
        &renamed,
        &parsed.package,
        &[Some(0), Some(1)],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap_err();
    assert!(matches!(renamed, ParseError::UnsupportedEdit(_)));
}

/// Charts are modelled now: the anchor, its `editAs` mode and every `c:f` come
/// through, and the shared parser reads the part with the workbook theme.
#[test]
fn reads_the_chart_anchor_and_every_reference() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let charts = &parsed.workbook.sheets[0].charts;
    assert_eq!(charts.len(), 1);
    assert!(parsed.workbook.sheets[1].charts.is_empty());
    let chart = &charts[0];
    assert_eq!(chart.part, "xl/charts/chart1.xml");
    assert_eq!(chart.drawing, "xl/drawings/drawing1.xml");
    assert_eq!(chart.anchor_index, 0);
    assert_eq!(
        chart.anchor,
        xlsx_model::ChartAnchor::TwoCell {
            from: xlsx_model::AnchorCell {
                col: 2,
                col_off: 12_700,
                row: 4,
                row_off: 0,
            },
            to: xlsx_model::AnchorCell {
                col: 8,
                col_off: 0,
                row: 19,
                row_off: 0,
            },
            edit_as: xlsx_model::AnchorEditAs::OneCell,
        }
    );
    assert_eq!(
        chart
            .refs
            .iter()
            .map(|reference| (reference.kind, reference.formula.as_str()))
            .collect::<Vec<_>>(),
        vec![
            (xlsx_model::ChartRefKind::SeriesName, "Data!$B$1"),
            (xlsx_model::ChartRefKind::Categories, "Data!$A$2:$A$4"),
            (xlsx_model::ChartRefKind::Values, "Data!$B$2:$B$4"),
        ]
    );

    let mut theme = xlsx_model::styles::Theme::default();
    theme.colors[5] = "#123456".to_owned();
    let space = crate::chart_space(CHART, &theme).expect("chart parses");
    assert_eq!(space.chart_type, "column");
    assert_eq!(space.plot_groups[0].series[0].color, "#123456");
    assert_eq!(
        space.plot_groups[0].series[0].value_formula.as_deref(),
        Some("Data!$B$2:$B$4")
    );
}

/// Saving a charted workbook whose references moved patches the chart part in
/// place: the reference and the cache beside it move together, and every byte
/// neither owns is left alone.
#[test]
fn patches_a_moved_chart_reference_together_with_its_cache() {
    let source = charted_package();
    let parsed = parse_workbook_with_package(&source).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("B3").unwrap(),
        Cell {
            value: CellValue::Number { value: 7.5 },
            ..Cell::default()
        },
    );
    workbook.sheets[0].charts[0].refs[2].formula = "Data!$B$2:$B$3".to_owned();

    let saved = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), Some(1)],
        &vec![SharedStringCells::new(); 2],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    let patched = String::from_utf8(part_bytes(&saved, "xl/charts/chart1.xml")).unwrap();
    let original = String::from_utf8(CHART.to_vec()).unwrap();
    assert_eq!(
        patched,
        original
            .replace("Data!$B$2:$B$4", "Data!$B$2:$B$3")
            .replace(
                r#"<c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache>"#,
                r#"<c:numCache><c:ptCount val="2"/><c:pt idx="1"><c:v>7.5</c:v></c:pt></c:numCache>"#,
            ),
        "the moved reference and its cache must both change, and nothing else"
    );
    assert_eq!(
        part_bytes(&saved, "xl/drawings/drawing1.xml"),
        DRAWING.to_vec()
    );
}

/// A cache that cannot be rebuilt correctly refuses the save rather than
/// leaving a reference pointing at one range while its cache holds another's.
#[test]
fn refuses_a_moved_reference_whose_cache_cannot_be_regenerated() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let mut refs = parsed.workbook.sheets[0].charts[0].refs.clone();
    for (formula, reason) in [
        ("Missing!$B$2:$B$4", "sheet this workbook does not hold"),
        ("Data!$A$2:$B$4", "more than one row and column"),
        ("(Data!$B$2:$B$4,Data!$D$2:$D$4)", "is not one area"),
        ("Data!$B$2:$B$1000000", "more points than a chart can carry"),
    ] {
        refs[2].formula = formula.to_owned();
        let error = patch_refs(CHART, &refs).unwrap_err();
        assert!(
            matches!(&error, ParseError::UnsupportedEdit(message) if message.contains(reason)),
            "{formula}: {error:?}"
        );
    }

    refs[2].formula = "Data!$B$2:$B$4".to_owned();
    refs[1].formula = "Data!$B$2:$B$4".to_owned();
    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].set_cell(
        CellRef::parse_a1("B2").unwrap(),
        Cell {
            value: CellValue::Number { value: 3.0 },
            ..Cell::default()
        },
    );
    let error = crate::chart::patch_chart_refs(CHART, &refs, &workbook, "Data").unwrap_err();
    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message)
            if message.contains("string cache over non-text cells")),
        "{error:?}"
    );
}

/// A cache may legally carry an `extLst` beside its points, or a `formatCode`
/// on a single point. Neither survives a rebuild, so the save is refused
/// rather than quietly changing the labels a consumer reads back.
#[test]
fn refuses_to_rebuild_a_cache_holding_content_it_does_not_model() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let mut refs = parsed.workbook.sheets[0].charts[0].refs.clone();
    refs[2].formula = "Data!$B$2:$B$3".to_owned();
    let source = String::from_utf8(CHART.to_vec()).unwrap();
    for (label, cache) in [
        (
            "extension list",
            r#"<c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:extLst><c:ext uri="{9F0C0C0A}"/></c:extLst></c:numCache>"#,
        ),
        (
            "per-point format code",
            r#"<c:numCache><c:pt idx="0" formatCode="0.00"><c:v>3</c:v></c:pt></c:numCache>"#,
        ),
    ] {
        let part = source.replace(
            r#"<c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache>"#,
            cache,
        );
        let error = patch_refs(part.as_bytes(), &refs).unwrap_err();
        assert!(
            matches!(&error, ParseError::UnsupportedEdit(message)
                if message.contains("does not model")),
            "{label}: {error:?}"
        );
    }
}

/// The writer is the last line: a reference carrying a character xml 1.0
/// cannot express is refused rather than written raw into a part Excel would
/// then have to repair.
#[test]
fn refuses_to_write_a_reference_xml_cannot_carry() {
    const CACHELESS: &[u8] = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:dLbls><c:f>Data!$C$1</c:f></c:dLbls></c:chart></c:chartSpace>"#;
    let reference = |formula: &str| {
        vec![xlsx_model::ChartRef {
            kind: xlsx_model::ChartRefKind::DataLabels,
            formula: formula.to_owned(),
        }]
    };
    let error = patch_refs(CACHELESS, &reference("Data!\u{7}$C$1")).unwrap_err();
    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message) if message.contains("U+0007")),
        "{error:?}"
    );

    let patched =
        String::from_utf8(patch_refs(CACHELESS, &reference("Data!$C$1\r\n")).unwrap()).unwrap();
    assert!(patched.contains("<c:f>Data!$C$1&#13;\n</c:f>"), "{patched}");
}

/// A chart may only be patched through the drawing and anchor it was read
/// from; a peer that repoints either is refused rather than moving somebody
/// else's anchor.
#[test]
fn refuses_a_chart_that_no_longer_names_its_source_anchor() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    for mutate in [
        (|chart: &mut xlsx_model::SheetChart| chart.anchor_index = 3) as fn(&mut _),
        |chart: &mut xlsx_model::SheetChart| chart.drawing = "xl/drawings/drawing9.xml".to_owned(),
    ] {
        let mut workbook = parsed.workbook.clone();
        let chart = &mut workbook.sheets[0].charts[0];
        chart.refs[2].formula = "Data!$B$2:$B$3".to_owned();
        mutate(chart);
        let error = crate::serialize_workbook_with_package_and_origins_after_edits(
            &workbook,
            &parsed.package,
            &[Some(0), Some(1)],
            &vec![SharedStringCells::new(); 2],
            SaveEdits {
                changed: true,
                moved_references: false,
            },
        )
        .unwrap_err();
        assert!(
            matches!(&error, ParseError::UnsupportedEdit(message)
                if message.contains("drawing and anchor it was read from")),
            "{error:?}"
        );
    }
}

/// This crate patches chart parts in place, so it can neither create one nor
/// delete one. A chart that appeared on a sheet, one that vanished from a
/// retained sheet and one carried onto a sheet with no source are all refused.
#[test]
fn refuses_chart_state_with_no_one_to_one_source() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let save = |workbook: &xlsx_model::Workbook, origins: &[Option<usize>]| {
        crate::serialize_workbook_with_package_and_origins_after_edits(
            workbook,
            &parsed.package,
            origins,
            &vec![SharedStringCells::new(); origins.len()],
            SaveEdits {
                changed: true,
                moved_references: false,
            },
        )
    };

    let mut appeared = parsed.workbook.clone();
    let chart = appeared.sheets[0].charts[0].clone();
    appeared.sheets[1].charts.push(chart.clone());
    let error = save(&appeared, &[Some(0), Some(1)]).unwrap_err();
    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message)
            if message.contains("cannot create one")),
        "{error:?}"
    );

    let mut dropped = parsed.workbook.clone();
    dropped.sheets[0].charts.clear();
    let error = save(&dropped, &[Some(0), Some(1)]).unwrap_err();
    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message)
            if message.contains("cannot delete one")),
        "{error:?}"
    );

    let mut added_sheet = parsed.workbook.clone();
    let mut sheet = xlsx_model::Sheet::new("Added");
    sheet.charts.push(chart);
    added_sheet.sheets.push(sheet);
    let error = save(&added_sheet, &[Some(0), Some(1), None]).unwrap_err();
    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message)
            if message.contains("cannot create one")),
        "{error:?}"
    );
}

/// A save writes an anchor's row and column back and nothing else, so a model
/// that also moved its mode, offsets, extent or kind is refused rather than
/// saved with that change dropped.
#[test]
fn refuses_an_anchor_change_the_drawing_patcher_would_drop() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let xlsx_model::ChartAnchor::TwoCell { from, to, .. } =
        parsed.workbook.sheets[0].charts[0].anchor
    else {
        panic!("two-cell anchor");
    };
    for (label, anchor) in [
        (
            "edit mode",
            xlsx_model::ChartAnchor::TwoCell {
                from,
                to,
                edit_as: xlsx_model::AnchorEditAs::TwoCell,
            },
        ),
        (
            "offset",
            xlsx_model::ChartAnchor::TwoCell {
                from: xlsx_model::AnchorCell {
                    col_off: 4_242,
                    ..from
                },
                to,
                edit_as: xlsx_model::AnchorEditAs::OneCell,
            },
        ),
        (
            "kind",
            xlsx_model::ChartAnchor::OneCell {
                from,
                extent: xlsx_model::AnchorExtent { cx: 100, cy: 100 },
            },
        ),
    ] {
        let mut workbook = parsed.workbook.clone();
        workbook.sheets[0].charts[0].anchor = anchor;
        let error = crate::serialize_workbook_with_package_and_origins_after_edits(
            &workbook,
            &parsed.package,
            &[Some(0), Some(1)],
            &vec![SharedStringCells::new(); 2],
            SaveEdits {
                changed: true,
                moved_references: false,
            },
        )
        .unwrap_err();
        assert!(
            matches!(&error, ParseError::UnsupportedEdit(message)
                if message.contains("more than the row and column")),
            "{label}: {error:?}"
        );
    }
}

/// A reference the edit deletes outright carries an empty cache, not the
/// values it used to hold.
#[test]
fn a_deleted_reference_empties_its_cache() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let mut refs = parsed.workbook.sheets[0].charts[0].refs.clone();
    refs[2].formula = "#REF!".to_owned();
    let patched = String::from_utf8(patch_refs(CHART, &refs).unwrap()).unwrap();
    assert!(
        patched.contains(r#"<c:f>#REF!</c:f><c:numCache><c:ptCount val="0"/></c:numCache>"#),
        "{patched}"
    );
}

/// A moved anchor is written back as `col`/`row` alone; offsets and the rest of
/// the drawing stay as authored.
#[test]
fn patches_only_the_moved_anchor_indices() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);
    let xlsx_model::ChartAnchor::TwoCell { from, to, edit_as } =
        workbook.sheets[0].charts[0].anchor
    else {
        panic!("two-cell anchor");
    };
    workbook.sheets[0].charts[0].anchor = xlsx_model::ChartAnchor::TwoCell {
        from: xlsx_model::AnchorCell { row: 7, ..from },
        to: xlsx_model::AnchorCell { row: 22, ..to },
        edit_as,
    };

    let saved = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), Some(1)],
        &vec![SharedStringCells::new(); 2],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    let patched = String::from_utf8(part_bytes(&saved, "xl/drawings/drawing1.xml")).unwrap();
    let original = String::from_utf8(DRAWING.to_vec()).unwrap();
    assert_eq!(
        patched,
        original
            .replace("<xdr:row>4</xdr:row>", "<xdr:row>7</xdr:row>")
            .replace("<xdr:row>19</xdr:row>", "<xdr:row>22</xdr:row>")
    );
}

#[test]
fn keeps_saving_ordinary_edits_to_a_charted_workbook() {
    let source = charted_package();
    let parsed = parse_workbook_with_package(&source).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);

    let edited = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), Some(1)],
        &vec![SharedStringCells::new(); 2],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    assert_eq!(
        part_bytes(&edited, "xl/charts/chart1.xml"),
        part_bytes(&source, "xl/charts/chart1.xml")
    );

    workbook.sheets.insert(1, xlsx_model::Sheet::new("Added"));
    let added = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), None, Some(1)],
        &vec![SharedStringCells::new(); 3],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    assert_eq!(parse_workbook(&added).unwrap().sheets.len(), 3);
}

#[test]
fn keeps_saving_ordinary_edits_with_an_unclaimed_chart_part() {
    let source = unclaimed_chart_package();
    let parsed = parse_workbook_with_package(&source).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);

    let edited = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), Some(1)],
        &vec![SharedStringCells::new(); 2],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    assert_eq!(
        part_bytes(&edited, "xl/charts/chart1.xml"),
        part_bytes(&source, "xl/charts/chart1.xml")
    );

    workbook.sheets.push(xlsx_model::Sheet::new("Added"));
    let added = crate::serialize_workbook_with_package_and_origins_after_edits(
        &workbook,
        &parsed.package,
        &[Some(0), Some(1), None],
        &vec![SharedStringCells::new(); 3],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    assert_eq!(parse_workbook(&added).unwrap().sheets.len(), 3);
}

/// A drawing may bind the relationship namespace to any prefix; discovery
/// resolves expanded names, so the chart is still found.
#[test]
fn discovers_a_chart_through_an_alternate_relationship_prefix() {
    let mut parts = charted_package();
    let drawing = String::from_utf8(DRAWING.to_vec())
        .unwrap()
        .replace("xmlns:r=", "xmlns:rel=")
        .replace("r:id=", "rel:id=");
    set_part(&mut parts, "xl/drawings/drawing1.xml", drawing.as_bytes());
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(parsed.workbook.sheets[0].charts.len(), 1);
    assert_eq!(parsed.package.unpatchable_reference_part(), None);
}

/// A chartsheet carries a drawing too, and the charts it anchors move with the
/// cells they name just like a worksheet's.
#[test]
fn discovers_the_charts_a_chartsheet_anchors() {
    let parts = chartsheet_package();
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(parsed.workbook.sheets[1].charts.len(), 1);
    assert_eq!(
        parsed.workbook.sheets[1].charts[0].part,
        "xl/charts/chart1.xml"
    );
    assert_eq!(parsed.package.unpatchable_reference_part(), None);
}

/// Every chart part the package holds must be one the model covers. A ChartEx
/// part, a chart no sheet claims, an externally cached chart, a pivot chart and
/// a chart carrying `sqref` extension references are all refused instead.
#[test]
fn refuses_structural_edits_while_a_chart_part_is_not_covered() {
    let unreachable = {
        let mut parts = charted_package();
        parts.push((
            "xl/charts/chart2.xml".to_owned(),
            b"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>"
                .to_vec(),
        ));
        parts
    };
    let chart_ex = {
        let mut parts = charted_package();
        set_part(
            &mut parts,
            "xl/charts/chart1.xml",
            br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data><cx:strDim><cx:f>Data!$A$2:$A$4</cx:f></cx:strDim></cx:data></cx:chartData></cx:chartSpace>"#,
        );
        parts
    };
    let filtered = {
        let mut parts = charted_package();
        let chart = String::from_utf8(CHART.to_vec()).unwrap().replace(
            "<c:f>Data!$A$2:$A$4</c:f>",
            r#"<c:f>Data!$A$2:$A$4</c:f><c:extLst><c:ext xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" uri="{02D57815-91ED-43cb-92C2-25804820EDAC}"><c15:fullRef><c15:sqref>Data!$A$2:$A$6</c15:sqref></c15:fullRef></c:ext></c:extLst>"#,
        );
        set_part(&mut parts, "xl/charts/chart1.xml", chart.as_bytes());
        parts
    };
    let pivoted = {
        let mut parts = charted_package();
        let chart = String::from_utf8(CHART.to_vec()).unwrap().replace(
            "<c:chart>",
            "<c:pivotSource><c:name>[1]Data!PivotTable1</c:name></c:pivotSource><c:chart>",
        );
        set_part(&mut parts, "xl/charts/chart1.xml", chart.as_bytes());
        parts
    };
    let external = {
        let mut parts = charted_package();
        let chart = String::from_utf8(CHART.to_vec()).unwrap().replace(
            "</c:chartSpace>",
            r#"<c:externalData r:id="rIdData"/></c:chartSpace>"#,
        );
        set_part(&mut parts, "xl/charts/chart1.xml", chart.as_bytes());
        parts
    };

    for (label, parts) in [
        ("unreachable", unreachable),
        ("chartex", chart_ex),
        ("filtered", filtered),
        ("pivoted", pivoted),
        ("external", external),
    ] {
        let parsed = parse_workbook_with_package(&parts).unwrap();
        let refused = parsed
            .package
            .unpatchable_reference_part()
            .unwrap_or_else(|| panic!("{label} chart part must be refused"));
        assert!(refused.starts_with("xl/charts/"), "{label}: {refused}");
    }
}

/// OPC types a part by its `Override` or, failing that, by the `Default` for
/// its extension. A chart typed the second way, at a path nothing conventional
/// would find, must still be refused.
#[test]
fn refuses_a_chart_typed_by_a_default_extension_mapping() {
    let mut parts = charted_package();
    parts.push((
        "[Content_Types].xml".to_owned(),
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="cxml" ContentType="application/vnd.ms-office.chartex+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_vec(),
    ));
    parts.push((
        "xl/extras/plot1.cxml".to_owned(),
        br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>"#
            .to_vec(),
    ));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(
        parsed.package.unpatchable_reference_part(),
        Some("xl/extras/plot1.cxml")
    );
}

/// Every part of an ordinary package resolves to `application/xml` through a
/// `Default`, so a chart type must not be the only thing discovery trusts.
#[test]
fn still_covers_a_conventional_chart_a_default_types_as_plain_xml() {
    let mut parts = charted_package();
    parts.push((
        "[Content_Types].xml".to_owned(),
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#.to_vec(),
    ));
    parts.push((
        "xl/charts/chart2.xml".to_owned(),
        b"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>"
            .to_vec(),
    ));

    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(
        parsed.package.unpatchable_reference_part(),
        Some("xl/charts/chart2.xml")
    );
}

/// A cache is only patchable when the reference beside it is one this crate
/// can resolve. A defined name, a union, an external book and a two-
/// dimensional area all survive a structural edit unchanged, so their caches
/// would keep pre-edit values; the package is refused instead.
#[test]
fn refuses_a_cache_this_crate_could_not_rebuild() {
    for (label, formula) in [
        ("defined name", "SalesRange"),
        ("union", "(Data!$B$2:$B$4,Data!$D$2:$D$4)"),
        ("external book", "[1]Data!$B$2:$B$4"),
        ("two dimensional", "Data!$A$2:$B$4"),
        ("whole column", "Data!$B:$B"),
    ] {
        let mut parts = charted_package();
        let chart = String::from_utf8(CHART.to_vec())
            .unwrap()
            .replace("Data!$B$2:$B$4", formula);
        set_part(&mut parts, "xl/charts/chart1.xml", chart.as_bytes());
        let parsed = parse_workbook_with_package(&parts).unwrap();
        assert_eq!(
            parsed.package.unpatchable_reference_part(),
            Some("xl/charts/chart1.xml"),
            "{label}"
        );
    }

    let mut parts = charted_package();
    let multi_level = String::from_utf8(CHART.to_vec())
        .unwrap()
        .replace("strCache>", "multiLvlStrCache>");
    set_part(&mut parts, "xl/charts/chart1.xml", multi_level.as_bytes());
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(
        parsed.package.unpatchable_reference_part(),
        Some("xl/charts/chart1.xml")
    );
}

/// A reference with no cache beside it carries nothing that could go stale,
/// so it does not have to be resolvable.
#[test]
fn accepts_a_reference_no_cache_depends_on() {
    let mut parts = charted_package();
    let chart = String::from_utf8(CHART.to_vec()).unwrap().replace(
        r#"<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>"#,
        r#"<c:val><c:numRef><c:f>SalesRange</c:f></c:numRef></c:val>"#,
    );
    set_part(&mut parts, "xl/charts/chart1.xml", chart.as_bytes());
    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(parsed.package.unpatchable_reference_part(), None);
}

/// OPC permits utf-8 and utf-16 alike. A utf-16 chart part must parse, and a
/// rewrite must come back in the encoding it was authored in.
#[test]
fn reads_and_rewrites_a_utf16_chart_part() {
    let text = String::from_utf8(CHART.to_vec()).unwrap();
    for (label, big_endian, bom) in [
        ("le+bom", false, true),
        ("be+bom", true, true),
        ("le", false, false),
        ("be", true, false),
    ] {
        let encoded = encode_utf16(&text, big_endian, bom);
        let mut parts = charted_package();
        set_part(&mut parts, "xl/charts/chart1.xml", &encoded);
        let parsed = parse_workbook_with_package(&parts).unwrap();
        let charts = &parsed.workbook.sheets[0].charts;
        assert_eq!(charts.len(), 1, "{label}");
        assert_eq!(charts[0].refs.len(), 3, "{label}");
        assert_eq!(parsed.package.unpatchable_reference_part(), None, "{label}");

        let mut refs = charts[0].refs.clone();
        refs[1].formula = "Data!$A$2:$A$5".to_owned();
        let utf8 = patch_refs(CHART, &refs).unwrap();
        assert_eq!(
            patch_refs(&encoded, &refs).unwrap(),
            encode_utf16(std::str::from_utf8(&utf8).unwrap(), big_endian, bom),
            "{label}"
        );
        assert!(
            crate::chart_space(&encoded, &Default::default()).is_some(),
            "{label}"
        );
    }
}

/// A part whose declaration names an encoding we cannot write back is refused
/// rather than reinterpreted as utf-8.
#[test]
fn refuses_a_part_declaring_a_foreign_encoding() {
    let source = br#"<?xml version="1.0" encoding="windows-1252"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
    assert!(matches!(
        patch_refs(source, &[]).unwrap_err(),
        ParseError::Malformed(message) if message.contains("windows-1252")
    ));
}

/// Patch a chart part against a workbook holding the sheet its references name.
fn patch_refs(part: &[u8], refs: &[xlsx_model::ChartRef]) -> Result<Vec<u8>, ParseError> {
    let workbook = parse_workbook(&charted_package()).unwrap();
    crate::chart::patch_chart_refs(part, refs, &workbook, "Data")
}

fn encode_utf16(text: &str, big_endian: bool, bom: bool) -> Vec<u8> {
    let mut out = Vec::new();
    if bom {
        out.extend_from_slice(if big_endian { b"\xFE\xFF" } else { b"\xFF\xFE" });
    }
    for unit in text.encode_utf16() {
        out.extend_from_slice(&if big_endian {
            unit.to_be_bytes()
        } else {
            unit.to_le_bytes()
        });
    }
    out
}

fn set_part(parts: &mut [(String, Vec<u8>)], path: &str, bytes: &[u8]) {
    let slot = parts
        .iter_mut()
        .find(|(name, _)| name == path)
        .unwrap_or_else(|| panic!("{path} is not in the package"));
    slot.1 = bytes.to_vec();
}

/// A package whose second sheet is a chartsheet anchoring the chart.
fn chartsheet_package() -> Vec<(String, Vec<u8>)> {
    let mut parts = package(r#"<sheetData/>"#, &[], false);
    parts[0] = (
        "xl/workbook.xml".to_owned(),
        br#"<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Chart" sheetId="2" r:id="rId2"/></sheets></workbook>"#.to_vec(),
    );
    parts[1] = (
        "xl/_rels/workbook.xml.rels".to_owned(),
        br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>"#.to_vec(),
    );
    parts.extend([
        (
            "xl/chartsheets/sheet1.xml".to_owned(),
            br#"<chartsheet><drawing r:id="rIdDrawing"/></chartsheet>"#.to_vec(),
        ),
        (
            "xl/chartsheets/_rels/sheet1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/drawings/drawing1.xml".to_owned(), DRAWING.to_vec()),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#.to_vec(),
        ),
        ("xl/charts/chart1.xml".to_owned(), CHART.to_vec()),
    ]);
    parts
}

/// Dropping a charted sheet must take its whole part graph with it, so deleted
/// user data is not still recoverable from the saved package.
#[test]
fn dropping_a_charted_sheet_prunes_its_unreachable_parts() {
    let mut parts = charted_package();
    parts.push((
        "[Content_Types].xml".to_owned(),
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_vec(),
    ));
    parts.push((
        "_rels/.rels".to_owned(),
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_vec(),
    ));
    parts[1] = (
        "xl/_rels/workbook.xml.rels".to_owned(),
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#.to_vec(),
    );

    let parsed = parse_workbook_with_package(&parts).unwrap();
    assert_eq!(parsed.workbook.sheets[0].charts.len(), 1);
    let mut dropped = parsed.workbook.clone();
    dropped.sheets.remove(0);

    let saved = crate::serialize_workbook_with_package_and_origins_after_edits(
        &dropped,
        &parsed.package,
        &[Some(1)],
        &vec![SharedStringCells::new(); 1],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    let names = saved
        .iter()
        .map(|(path, _)| path.as_str())
        .collect::<Vec<_>>();
    for orphan in [
        "xl/worksheets/sheet1.xml",
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
    ] {
        assert!(!names.contains(&orphan), "{orphan} is still in {names:?}");
    }
    let content_types = String::from_utf8(part_bytes(&saved, "[Content_Types].xml")).unwrap();
    assert!(!content_types.contains("chart1.xml"), "{content_types}");
    assert!(names.contains(&"xl/worksheets/sheet2.xml"), "{names:?}");
}

/// Reordering, dropping and renaming sheets no longer trips the guard now that
/// chart references are modelled.
#[test]
fn saves_sheet_moves_on_a_charted_workbook() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let mut workbook = parsed.workbook.clone();
    edit_a1(&mut workbook, 1.0);
    let provenance = vec![SharedStringCells::new(); 2];

    for origins in [[Some(1), Some(0)], [Some(0), Some(1)]] {
        let mut moved = workbook.clone();
        if origins[0] == Some(1) {
            moved.sheets.swap(0, 1);
        }
        crate::serialize_workbook_with_package_and_origins_after_edits(
            &moved,
            &parsed.package,
            &origins,
            &provenance,
            SaveEdits {
                changed: true,
                moved_references: false,
            },
        )
        .expect("reordering a charted workbook saves");
    }

    let mut renamed = workbook.clone();
    renamed.sheets[0].name = "Renamed".to_owned();
    crate::serialize_workbook_with_package_and_origins_after_edits(
        &renamed,
        &parsed.package,
        &[Some(0), Some(1)],
        &provenance,
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .expect("renaming a charted workbook saves");

    let mut dropped = workbook.clone();
    dropped.sheets.pop();
    crate::serialize_workbook_with_package_and_origins_after_edits(
        &dropped,
        &parsed.package,
        &[Some(0)],
        &vec![SharedStringCells::new(); 1],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .expect("dropping a sheet from a charted workbook saves");
}

/// Chart XML is untrusted input reaching wasm, so the tree builder must bound
/// or reject it rather than recurse or allocate without a ceiling.
#[test]
fn hostile_chart_markup_is_refused_not_survived() {
    let deep = format!(
        "<c:chartSpace xmlns:c=\"c\">{}{}</c:chartSpace>",
        "<x>".repeat(crate::MAX_DEPTH + 2),
        "</x>".repeat(crate::MAX_DEPTH + 2)
    );
    assert_eq!(
        patch_refs(deep.as_bytes(), &[]).unwrap_err(),
        ParseError::DepthExceeded
    );

    let self_closing = format!(
        "<c:chartSpace xmlns:c=\"c\">{}<x/>{}</c:chartSpace>",
        "<x>".repeat(crate::MAX_DEPTH - 1),
        "</x>".repeat(crate::MAX_DEPTH - 1)
    );
    assert_eq!(
        patch_refs(self_closing.as_bytes(), &[]).unwrap_err(),
        ParseError::DepthExceeded
    );

    let doctype = br#"<!DOCTYPE c:chartSpace [<!ENTITY x "boom">]><c:chartSpace xmlns:c="c"/>"#;
    assert!(matches!(
        patch_refs(doctype, &[]).unwrap_err(),
        ParseError::Malformed(_)
    ));

    let unclosed = br#"<c:chartSpace xmlns:c="c"><c:chart>"#;
    assert!(matches!(
        patch_refs(unclosed, &[]).unwrap_err(),
        ParseError::Malformed(_)
    ));

    assert!(crate::chart_space(b"<c:chartSpace xmlns:c=\"c\"/>", &Default::default()).is_none());
    assert!(crate::chart_space(b"not xml at all", &Default::default()).is_none());
    assert!(
        crate::chart_space(
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:bar3DChart/></c:plotArea></c:chart></c:chartSpace>"#,
            &Default::default(),
        )
        .is_none()
    );
}

/// A reference carrying markup-significant characters must come back escaped,
/// so a rewrite can never reshape the part.
#[test]
fn a_rewritten_reference_is_escaped_into_the_part() {
    let source = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:f>Data!$A$1</c:f></c:chartSpace>"#;
    let patched = patch_refs(
        source,
        &[xlsx_model::ChartRef {
            kind: xlsx_model::ChartRefKind::Values,
            formula: "'A<B&C'!$A$1".to_owned(),
        }],
    )
    .unwrap();
    assert_eq!(
        String::from_utf8(patched).unwrap(),
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:f>&apos;A&lt;B&amp;C&apos;!$A$1</c:f></c:chartSpace>"#
            .replace("&apos;", "'")
    );
}

/// The patcher addresses `c:f` by document order, so a part that no longer
/// holds the same references is refused rather than rewritten into the wrong
/// slots.
#[test]
fn refuses_to_patch_a_chart_part_that_no_longer_matches() {
    let parsed = parse_workbook_with_package(&charted_package()).unwrap();
    let mut refs = parsed.workbook.sheets[0].charts[0].refs.clone();
    refs.pop();
    let error = patch_refs(CHART, &refs).unwrap_err();
    assert!(
        matches!(&error, ParseError::UnsupportedEdit(message)
            if message.contains("holds 3 references but the model carries 2")),
        "{error:?}"
    );
}
