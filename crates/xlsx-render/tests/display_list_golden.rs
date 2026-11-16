//! Whole display lists for table, plain and merged sheets, compared with
//! golden files under `tests/golden`. `UPDATE_GOLDEN=1` rewrites them.

use std::path::PathBuf;

use xlsx_model::{SheetId, Workbook};
use xlsx_render::{Viewport, build_display_list};

const STYLES: &str = r#"<styleSheet><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="medium"><color rgb="FFFF0000"/></bottom></border></borders><cellXfs count="3"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="0" fillId="2" borderId="0" applyFill="1"/><xf fontId="0" fillId="0" borderId="1" applyBorder="1"/></cellXfs><dxfs count="8"><dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill><bgColor rgb="FF195A82"/></patternFill></fill></dxf><dxf><fill><patternFill><bgColor rgb="FFFFF78C"/></patternFill></fill></dxf><dxf><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf><dxf><font><b/></font></dxf><dxf><font><i/><color rgb="FF9C0006"/></font></dxf><dxf><font><b/></font><fill><patternFill><bgColor rgb="FFD9D9D9"/></patternFill></fill></dxf><dxf><font><color rgb="FF1F3864"/></font></dxf><dxf><fill><patternFill><bgColor rgb="FFFCE4D6"/></patternFill></fill></dxf></dxfs><tableStyles count="1"><tableStyle name="Wide" pivot="0" count="11"><tableStyleElement type="wholeTable" dxfId="6"/><tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="totalRow" dxfId="5"/><tableStyleElement type="firstColumn" dxfId="3"/><tableStyleElement type="lastColumn" dxfId="4"/><tableStyleElement type="firstRowStripe" dxfId="1" size="2"/><tableStyleElement type="secondRowStripe" dxfId="7"/><tableStyleElement type="firstColumnStripe" dxfId="2" size="2"/><tableStyleElement type="secondColumnStripe" dxfId="7" size="3"/><tableStyleElement type="firstHeaderCell" dxfId="5"/><tableStyleElement type="lastTotalCell" dxfId="0"/></tableStyle></tableStyles></styleSheet>"#;

const COLUMNS: &str = "ABCDEFGHIJKL";

/// A one-sheet workbook whose sheet holds `rows` by `cols` cells, the xf of
/// each from `style`, plus `extra` after the sheet data and `tables` as
/// `(range, totals rows, style info)`.
fn workbook(
    rows: u32,
    cols: usize,
    style: impl Fn(u32, usize) -> u32,
    extra: &str,
    tables: &[(&str, u8, &str)],
) -> Workbook {
    let rels = |body: &str| {
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#
        )
    };
    let link = |id: &str, kind: &str, target: &str| {
        format!(
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        )
    };
    let data = (1..=rows)
        .map(|row| {
            let cells = (0..cols)
                .map(|col| {
                    let name = &COLUMNS[col..=col];
                    let xf = match style(row, col) {
                        0 => String::new(),
                        xf => format!(r#" s="{xf}""#),
                    };
                    if col == 0 {
                        format!(r#"<c r="{name}{row}"{xf} t="inlineStr"><is><t>{name}{row}</t></is></c>"#)
                    } else {
                        format!(r#"<c r="{name}{row}"{xf}><v>{}</v></c>"#, row * 10 + col as u32)
                    }
                })
                .collect::<String>();
            format!(r#"<row r="{row}">{cells}</row>"#)
        })
        .collect::<String>();
    let table_rels = tables
        .iter()
        .enumerate()
        .map(|(index, _)| {
            link(
                &format!("rId{}", index + 1),
                "table",
                &format!("../tables/table{}.xml", index + 1),
            )
        })
        .collect::<String>();
    let mut parts = vec![
        (
            "xl/workbook.xml".to_string(),
            r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels(&(link("rId1", "worksheet", "worksheets/sheet1.xml")
                + &link("rId2", "styles", "styles.xml"))),
        ),
        ("xl/styles.xml".to_string(), STYLES.to_string()),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            format!("<worksheet><sheetData>{data}</sheetData>{extra}</worksheet>"),
        ),
    ];
    if !tables.is_empty() {
        parts.push((
            "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            rels(&table_rels),
        ));
    }
    for (index, (range, totals, info)) in tables.iter().enumerate() {
        let (start, end) = range.split_once(':').unwrap();
        let width = COLUMNS.find(&end[..1]).unwrap() - COLUMNS.find(&start[..1]).unwrap() + 1;
        let columns = (0..width)
            .map(|col| format!(r#"<tableColumn id="{}" name="c{col}"/>"#, col + 1))
            .collect::<String>();
        let id = index + 1;
        parts.push((
            format!("xl/tables/table{id}.xml"),
            format!(
                r#"<table id="{id}" name="T{id}" displayName="T{id}" ref="{range}" totalsRowCount="{totals}"><tableColumns count="{width}">{columns}</tableColumns>{info}</table>"#
            ),
        ));
    }
    let parts = parts
        .into_iter()
        .map(|(path, xml)| (path, xml.into_bytes()))
        .collect::<Vec<_>>();
    xlsx_parse::parse_workbook(&parts).unwrap()
}

fn info(name: &str) -> String {
    format!(
        r#"<tableStyleInfo name="{name}" showFirstColumn="1" showLastColumn="1" showRowStripes="1" showColumnStripes="1"/>"#
    )
}

const MERGES: &str = r#"<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="D3:E5"/><mergeCell ref="B7:C8"/></mergeCells>"#;

const CONDITIONAL: &str = r#"<conditionalFormatting sqref="B1:F9"><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>40</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="A1:A9"><cfRule type="containsText" dxfId="4" priority="2" operator="containsText" text="A3"><formula>NOT(ISERROR(SEARCH("A3",A1)))</formula></cfRule></conditionalFormatting>"#;

fn border_or_fill(row: u32, col: usize) -> u32 {
    match (row + col as u32) % 5 {
        0 => 2,
        3 if row % 2 == 0 => 1,
        _ => 0,
    }
}

/// Compares the display list of each viewport with `tests/golden/{name}.json`.
fn assert_golden(name: &str, wb: &Workbook) {
    let viewports = [
        Viewport {
            x: 0.0,
            y: 0.0,
            width: 900.0,
            height: 320.0,
        },
        Viewport {
            x: 70.0,
            y: 25.0,
            width: 260.0,
            height: 120.0,
        },
    ];
    let lists = viewports
        .iter()
        .map(|viewport| build_display_list(wb, SheetId(0), viewport).unwrap())
        .collect::<Vec<_>>();
    let actual = serde_json::to_string_pretty(&lists).unwrap() + "\n";
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden")
        .join(format!("{name}.json"));
    if std::env::var_os("UPDATE_GOLDEN").is_some() {
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, actual).unwrap();
        return;
    }
    let expected = std::fs::read_to_string(&path).unwrap_or_else(|_| {
        panic!(
            "missing {}: run this test with UPDATE_GOLDEN=1 to write it",
            path.display()
        )
    });
    assert!(actual == expected, "{name} display list differs from its golden file");
}

#[test]
fn styled_tables_match_golden() {
    let custom = info("Wide");
    let builtin = info("TableStyleMedium9");
    let wb = workbook(
        12,
        12,
        border_or_fill,
        "",
        &[("A1:F9", 1, &custom), ("H2:K11", 1, &builtin)],
    );
    assert_golden("styled_tables", &wb);
}

#[test]
fn column_stripes_match_golden() {
    let columns_only = r#"<tableStyleInfo name="Wide" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="1"/>"#;
    let wb = workbook(12, 12, |_, _| 0, "", &[("A1:L11", 1, columns_only)]);
    assert_golden("column_stripes", &wb);
}

#[test]
fn plain_sheet_matches_golden() {
    let wb = workbook(12, 8, border_or_fill, "", &[]);
    assert_golden("plain_sheet", &wb);
}

#[test]
fn merged_cells_match_golden() {
    let extra = format!("{MERGES}{CONDITIONAL}");
    let wb = workbook(10, 7, border_or_fill, &extra, &[]);
    assert_golden("merged_cells", &wb);
}

#[test]
fn merged_cells_with_table_match_golden() {
    let extra = format!("{MERGES}{CONDITIONAL}");
    let custom = info("Wide");
    let wb = workbook(10, 7, border_or_fill, &extra, &[("B2:F8", 1, &custom)]);
    assert_golden("merged_cells_with_table", &wb);
}
