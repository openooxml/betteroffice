//! A workbook's own `<numFmt>` entry for an id the built-in table also numbers
//! is what its cells show, the same as that code declared under a custom id.

use xlsx_model::{CellRef, SheetId, Workbook};
use xlsx_render::{DrawCmd, Viewport, build_display_list, display_text};

const ACCOUNTING: &str = r#"_-* #,##0_-;\-* #,##0_-;_-* &quot;-&quot;_-;_-@_-"#;
const EAST_ASIAN_DATE: &str = "yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;";
const ISO_DATE: &str = "yyyy-mm-dd";

/// Column A holds `values`, each styled with an xf whose `numFmtId` is `id`;
/// `declared` is the one `<numFmt>` entry of the styles part, if any, with its
/// code as the attribute spells it.
fn workbook(id: u16, declared: Option<(u16, &str)>, values: &[f64]) -> Workbook {
    let rels = |body: &str| {
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#
        )
    };
    let link = |rid: &str, kind: &str, target: &str| {
        format!(
            r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        )
    };
    let num_fmts = declared.map_or_else(String::new, |(declared, code)| {
        format!(
            r#"<numFmts count="1"><numFmt numFmtId="{declared}" formatCode="{code}"/></numFmts>"#
        )
    });
    let styles = format!(
        r#"<styleSheet>{num_fmts}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="{id}" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#
    );
    let rows = values
        .iter()
        .enumerate()
        .map(|(index, value)| {
            let row = index + 1;
            format!(r#"<row r="{row}"><c r="A{row}" s="1"><v>{value}</v></c></row>"#)
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
        ("xl/styles.xml", styles),
        (
            "xl/worksheets/sheet1.xml",
            format!("<worksheet><sheetData>{rows}</sheetData></worksheet>"),
        ),
    ];
    xlsx_parse::parse_workbook(&parts.map(|(path, xml)| (path.to_string(), xml.into_bytes())))
        .unwrap()
}

/// `code` declared under the built-in `id` and, in the twin, under the custom
/// id 164, each workbook's cells styled with its own declaration.
fn twins(id: u16, code: &str, values: &[f64]) -> (Workbook, Workbook) {
    (
        workbook(id, Some((id, code)), values),
        workbook(164, Some((164, code)), values),
    )
}

/// The display text of column A's cells, top to bottom.
fn shown(wb: &Workbook) -> Vec<String> {
    let sheet = &wb.sheets[0];
    (0..8)
        .filter_map(|row| sheet.cell(CellRef::new(row, 0)))
        .map(|cell| display_text(&wb.styles, wb.date_system, cell))
        .collect()
}

/// The text runs the display list draws, in drawing order.
fn drawn(wb: &Workbook) -> Vec<String> {
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 400.0,
        height: 200.0,
    };
    build_display_list(wb, SheetId(0), &viewport)
        .unwrap()
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text { text, .. } => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

#[test]
fn a_declared_accounting_code_under_a_builtin_id_is_the_one_shown() {
    let (declared, custom) = twins(41, ACCOUNTING, &[-1234.0, 1234.0]);
    let text = shown(&declared);
    assert_eq!(text, shown(&custom));
    assert!(text[0].contains('-') && !text[0].contains('('), "{text:?}");
}

#[test]
fn a_declared_east_asian_date_id_shows_a_date() {
    let (declared, custom) = twins(31, EAST_ASIAN_DATE, &[45000.0]);
    let text = shown(&declared);
    assert_eq!(text, shown(&custom));
    assert!(text[0].contains('年'), "{text:?}");
}

#[test]
fn a_redefined_builtin_date_id_shows_the_workbook_code() {
    let (declared, custom) = twins(14, ISO_DATE, &[45000.0]);
    assert_eq!(shown(&declared), shown(&custom));
    assert_eq!(shown(&custom), ["2023-03-15"]);
}

#[test]
fn an_undeclared_builtin_accounting_id_keeps_its_parentheses() {
    let text = shown(&workbook(41, None, &[-1234.0]));
    assert!(text[0].contains('('), "{text:?}");
}

#[test]
fn the_display_list_draws_a_declared_builtin_id_like_its_custom_twin() {
    for (id, code, values) in [
        (41, ACCOUNTING, &[-1234.0, 1234.0][..]),
        (31, EAST_ASIAN_DATE, &[45000.0][..]),
        (14, ISO_DATE, &[45000.0][..]),
    ] {
        let (declared, custom) = twins(id, code, values);
        assert_eq!(drawn(&declared), drawn(&custom), "numFmtId {id}");
    }
    let runs = drawn(&twins(41, ACCOUNTING, &[-1234.0]).0);
    assert!(
        runs.iter().any(|run| run.contains('-')) && runs.iter().all(|run| !run.contains('(')),
        "{runs:?}"
    );
}
