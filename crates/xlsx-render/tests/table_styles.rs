use xlsx_model::{CellRef, Color, SheetId, Theme, Workbook};
use xlsx_render::{DisplayList, DrawCmd, GridGeometry, Viewport, build_display_list};

const STYLES: &str = r#"<styleSheet><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs><dxfs count="2"><dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill><bgColor rgb="FF195A82"/></patternFill></fill></dxf><dxf><fill><patternFill><bgColor rgb="FFFFF78C"/></patternFill></fill></dxf></dxfs><tableStyles count="1"><tableStyle name="Budget" pivot="0" count="2"><tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="firstRowStripe" dxfId="1"/></tableStyle></tableStyles></styleSheet>"#;

const PLAIN_STRIPES: &str = r#"<styleSheet><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="2"><dxf><fill><patternFill><bgColor rgb="FFFFF78C"/></patternFill></fill></dxf><dxf/></dxfs><tableStyles count="2"><tableStyle name="Unset" pivot="0" count="2"><tableStyleElement type="firstRowStripe" dxfId="0"/><tableStyleElement type="secondRowStripe" size="3"/></tableStyle><tableStyle name="Blank" pivot="0" count="2"><tableStyleElement type="firstRowStripe" dxfId="0"/><tableStyleElement type="secondRowStripe" dxfId="1" size="3"/></tableStyle></tableStyles></styleSheet>"#;

/// A sheet holding table `A1:C5`: a header row, three data rows with `B2`
/// using xf `b2_style`, and a fifth row the table counts as totals when
/// `totals` is set.
fn workbook(style_info: &str, b2_style: u32, totals: bool) -> Workbook {
    workbook_with(STYLES, style_info, b2_style, totals)
}

fn workbook_with(styles: &str, style_info: &str, b2_style: u32, totals: bool) -> Workbook {
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
    let rows = (1..=5)
        .map(|row| {
            let style = if row == 2 { format!(r#" s="{b2_style}""#) } else { String::new() };
            format!(
                r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>a{row}</t></is></c><c r="B{row}"{style}><v>{row}</v></c></row>"#
            )
        })
        .collect::<String>();
    let parts = [
        (
            "xl/workbook.xml",
            r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels(&(link("rId1", "worksheet", "worksheets/sheet1.xml")
                + &link("rId2", "styles", "styles.xml"))),
        ),
        ("xl/styles.xml", styles.to_string()),
        (
            "xl/worksheets/sheet1.xml",
            format!("<worksheet><sheetData>{rows}</sheetData></worksheet>"),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            rels(&link("rId1", "table", "../tables/table1.xml")),
        ),
        (
            "xl/tables/table1.xml",
            format!(
                r#"<table id="1" name="T" displayName="T" ref="A1:C5" totalsRowCount="{}"><tableColumns count="3"><tableColumn id="1" name="a1"/><tableColumn id="2" name="1"/><tableColumn id="3" name="c"/></tableColumns>{style_info}</table>"#,
                u8::from(totals)
            ),
        ),
    ];
    xlsx_parse::parse_workbook(&parts.map(|(path, xml)| (path.to_string(), xml.into_bytes())))
        .unwrap()
}

fn info(name: &str, row_stripes: bool) -> String {
    format!(
        r#"<tableStyleInfo name="{name}" showFirstColumn="0" showLastColumn="0" showRowStripes="{}" showColumnStripes="0"/>"#,
        u8::from(row_stripes)
    )
}

fn render(wb: &Workbook) -> DisplayList {
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 400.0,
        height: 200.0,
    };
    build_display_list(wb, SheetId(0), &viewport).unwrap()
}

/// The topmost fill under the middle of cell `at`, past the sheet background.
fn fill_at(wb: &Workbook, list: &DisplayList, at: CellRef) -> Option<String> {
    let geometry = GridGeometry::new(&wb.sheets[0], &wb.styles);
    let x = (geometry.col_x(at.col) + geometry.col_x(at.col + 1)) / 2.0;
    let y = (geometry.row_y(at.row) + geometry.row_y(at.row + 1)) / 2.0;
    list.commands[1..]
        .iter()
        .filter_map(|command| match command {
            DrawCmd::FillRect {
                x: left,
                y: top,
                w,
                h,
                color,
                ..
            } if (*left..left + w).contains(&x) && (*top..top + h).contains(&y) => {
                Some(color.to_string())
            }
            _ => None,
        })
        .next_back()
}

/// The color and boldness `text` draws with.
fn text(list: &DisplayList, text: &str) -> (String, bool) {
    list.commands
        .iter()
        .find_map(|command| match command {
            DrawCmd::Text {
                text: drawn,
                color,
                bold,
                ..
            } if &**drawn == text => Some((color.to_string(), *bold)),
            _ => None,
        })
        .unwrap()
}

fn theme(idx: u8, tint: f64) -> Option<String> {
    Color::Theme { idx, tint }.resolve(&Theme::default())
}

#[test]
fn medium2_paints_its_header_and_banded_rows() {
    let wb = workbook(&info("TableStyleMedium2", true), 0, false);
    let list = render(&wb);
    let band = theme(4, 0.7999816888943144);
    assert_eq!(fill_at(&wb, &list, CellRef::new(0, 2)), theme(4, 0.0));
    assert_eq!(text(&list, "a1"), ("#ffffff".into(), true));
    assert_eq!(fill_at(&wb, &list, CellRef::new(1, 1)), band);
    assert_eq!(fill_at(&wb, &list, CellRef::new(1, 2)), band);
    assert_eq!(fill_at(&wb, &list, CellRef::new(2, 1)), None);
    assert_eq!(fill_at(&wb, &list, CellRef::new(3, 2)), band);
    assert_eq!(fill_at(&wb, &list, CellRef::new(1, 3)), None);
    assert_eq!(text(&list, "a2"), ("#000000".into(), false));
}

#[test]
fn turning_off_row_stripes_removes_the_banding() {
    let wb = workbook(&info("TableStyleMedium2", false), 0, false);
    let list = render(&wb);
    assert_eq!(fill_at(&wb, &list, CellRef::new(0, 0)), theme(4, 0.0));
    for row in 1..5 {
        assert_eq!(fill_at(&wb, &list, CellRef::new(row, 1)), None);
    }
}

#[test]
fn a_workbook_defined_style_draws_its_dxfs() {
    let wb = workbook(&info("Budget", true), 0, false);
    let list = render(&wb);
    assert_eq!(
        fill_at(&wb, &list, CellRef::new(0, 1)),
        Some("#195a82".into())
    );
    assert_eq!(text(&list, "a1"), ("#ffffff".into(), true));
    assert_eq!(
        fill_at(&wb, &list, CellRef::new(1, 0)),
        Some("#fff78c".into())
    );
    assert_eq!(fill_at(&wb, &list, CellRef::new(2, 0)), None);
}

#[test]
fn a_cell_fill_beats_the_table_style() {
    let wb = workbook(&info("TableStyleMedium2", true), 1, false);
    let list = render(&wb);
    assert_eq!(
        fill_at(&wb, &list, CellRef::new(1, 1)),
        Some("#ffff00".into())
    );
    assert_eq!(
        fill_at(&wb, &list, CellRef::new(1, 0)),
        theme(4, 0.7999816888943144)
    );
}

#[test]
fn a_total_row_draws_its_element() {
    let wb = workbook(&info("TableStyleDark1", true), 0, true);
    let list = render(&wb);
    assert_eq!(
        fill_at(&wb, &list, CellRef::new(4, 1)),
        theme(1, 0.1499984740745262)
    );
    assert_eq!(text(&list, "a5"), ("#ffffff".into(), true));
    assert_eq!(text(&list, "a4"), ("#ffffff".into(), false));
}

#[test]
fn a_stripe_that_paints_nothing_keeps_its_size() {
    for style in ["Unset", "Blank"] {
        let wb = workbook_with(PLAIN_STRIPES, &info(style, true), 0, false);
        let list = render(&wb);
        assert_eq!(
            fill_at(&wb, &list, CellRef::new(1, 1)),
            Some("#fff78c".into()),
            "{style}"
        );
        for row in 2..5 {
            assert_eq!(fill_at(&wb, &list, CellRef::new(row, 1)), None, "{style} row {row}");
        }
    }
}
