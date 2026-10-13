//! A collaboration snapshot saved by a release before patterned fills holds the
//! released reading of every source format: a pattern as a solid of its foreground,
//! the extended border styles as medium or dashed. The released builds are modelled by
//! a twin package whose styles are written in that reading.

use std::collections::BTreeMap;

use betteroffice_xlsx::{CalculationOptions, CellRange, CellRef, SheetId, StylePatch, Workbook};
use xlsx_model::{Color, Fill};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NONE: &str = r#"<fill><patternFill patternType="none"/></fill>"#;
const LIGHT_UP: &str = concat!(
    r#"<fill><patternFill patternType="lightUp"><fgColor rgb="FF123456"/>"#,
    r#"<bgColor rgb="FFFEDCBA"/></patternFill></fill>"#,
);
const DARK_GRID: &str = concat!(
    r#"<fill><patternFill patternType="darkGrid"><fgColor rgb="FF123456"/>"#,
    r#"<bgColor rgb="FF00FF00"/></patternFill></fill>"#,
);
const SOLID: &str =
    r#"<fill><patternFill patternType="solid"><fgColor rgb="FF123456"/></patternFill></fill>"#;
const EMPTY: &str = "<border><left/><right/><top/><bottom/><diagonal/></border>";
const DASH_DOT: &str = concat!(
    r#"<border><left/><right/><top/><bottom style="mediumDashDot"><color rgb="FF654321"/>"#,
    r#"</bottom><diagonal/></border>"#,
);
const MEDIUM_DASHED: &str = concat!(
    r#"<border><left/><right/><top/><bottom style="mediumDashed"><color rgb="FF654321"/>"#,
    r#"</bottom><diagonal/></border>"#,
);
const MEDIUM: &str = concat!(
    r#"<border><left/><right/><top/><bottom style="medium"><color rgb="FF654321"/>"#,
    r#"</bottom><diagonal/></border>"#,
);

fn package(fills: &[&str], borders: &[&str], cells: &[(&str, usize)]) -> Vec<u8> {
    let xfs: String = (0..fills.len())
        .map(|index| {
            format!(
                r#"<xf numFmtId="0" fontId="0" fillId="{index}" borderId="{index}" xfId="0" applyFill="1" applyBorder="1"/>"#
            )
        })
        .collect();
    let styles = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{main}">"#,
            r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
            r#"<fills count="{count}">{fills}</fills><borders count="{count}">{borders}</borders>"#,
            r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
            r#"<cellXfs count="{count}">{xfs}</cellXfs></styleSheet>"#,
        ),
        main = MAIN,
        count = fills.len(),
        fills = fills.concat(),
        borders = borders.concat(),
        xfs = xfs,
    );
    let cells: String = cells
        .iter()
        .map(|(at, style)| format!(r#"<c r="{at}" s="{style}"><v>1</v></c>"#))
        .collect();
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                r#"<Default Extension="xml" ContentType="application/xml"/>"#,
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
                r#"</Types>"#,
            )
            .to_owned(),
        ),
        (
            "_rels/.rels",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        (
            "xl/workbook.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                    r#"<workbook xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                    r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                ),
                main = MAIN,
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#,
                r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        ("xl/styles.xml", styles),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}"><sheetData><row r="1">{cells}</row></sheetData></worksheet>"#
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(package: &[u8], name: &str) -> String {
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(package)
        .unwrap()
        .into_iter()
        .collect();
    String::from_utf8(parts[name].clone()).unwrap()
}

fn attribute<'a>(tag: &'a str, name: &str) -> &'a str {
    let tag = &tag[..tag.find('>').unwrap()];
    tag.split(&format!(" {name}=\""))
        .nth(1)
        .unwrap()
        .split('"')
        .next()
        .unwrap()
}

/// The `index`th `<item>` inside the `<pool>` element of `styles`.
fn pool_entry<'a>(styles: &'a str, pool: &str, item: &str, index: usize) -> &'a str {
    let table = &styles[styles.find(&format!("<{pool}")).unwrap()..];
    let table = &table[..table.find(&format!("</{pool}>")).unwrap()];
    let open = format!("<{item}");
    let mut starts = table.match_indices(&open).map(|(at, _)| at).filter(|at| {
        matches!(
            table.as_bytes().get(at + open.len()),
            Some(b' ' | b'>' | b'/')
        )
    });
    let entry = &table[starts.nth(index).unwrap()..];
    let tag_end = entry.find('>').unwrap();
    let end = if entry[..tag_end].ends_with('/') {
        tag_end + 1
    } else {
        let close = format!("</{item}>");
        entry.find(&close).unwrap() + close.len()
    };
    &entry[..end]
}

/// The `<xf>`, `<fill>` and `<border>` the saved package gives `cell`.
fn saved_format(package: &[u8], cell: &str) -> (usize, String, String, String) {
    let styles = part(package, "xl/styles.xml");
    let sheet = part(package, "xl/worksheets/sheet1.xml");
    let at = &sheet[sheet.find(&format!(r#"<c r="{cell}""#)).unwrap()..];
    let style: usize = attribute(at, "s").parse().unwrap();
    let xf = pool_entry(&styles, "cellXfs", "xf", style);
    let fill = attribute(xf, "fillId").parse().unwrap();
    let border = attribute(xf, "borderId").parse().unwrap();
    (
        style,
        xf.to_owned(),
        pool_entry(&styles, "fills", "fill", fill).to_owned(),
        pool_entry(&styles, "borders", "border", border).to_owned(),
    )
}

/// A snapshot as a released build saves it after editing B1, optionally after
/// authoring `released_edit` on A1 first.
fn released_snapshot(twin: &[u8], released_edit: Option<StylePatch>) -> Vec<u8> {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open_collaborative_recalculated(twin, 7_101, options).unwrap();
    let a1 = CellRef::parse_a1("A1").unwrap();
    if let Some(patch) = released_edit {
        workbook
            .patch_range_style(SheetId(0), CellRange::new(a1, a1), patch, options)
            .unwrap();
    }
    workbook
        .edit_cell(SheetId(0), CellRef::parse_a1("B1").unwrap(), "edited", options)
        .unwrap();
    workbook.encode_state_as_update_v1()
}

fn restored(source: &[u8], snapshot: &[u8]) -> Workbook {
    let options = CalculationOptions::default();
    let mut workbook = Workbook::open_collaborative_recalculated(source, 7_102, options).unwrap();
    assert!(workbook.apply_update_v1(snapshot, options).unwrap().applied);
    let b1 = workbook
        .sheet(SheetId(0))
        .unwrap()
        .cell(CellRef::parse_a1("B1").unwrap())
        .map(|cell| cell.value.clone());
    assert!(b1.is_some_and(|value| format!("{value:?}").contains("edited")));
    workbook
}

fn single() -> (Vec<u8>, Vec<u8>) {
    (
        package(&[NONE, LIGHT_UP], &[EMPTY, DASH_DOT], &[("A1", 1)]),
        package(&[NONE, SOLID], &[EMPTY, MEDIUM], &[("A1", 1)]),
    )
}

#[test]
fn a_released_snapshot_keeps_the_source_pattern_and_border_on_save() {
    let (source, twin) = single();
    let saved = restored(&source, &released_snapshot(&twin, None))
        .save()
        .unwrap();
    let (_, xf, fill, border) = saved_format(&saved, "A1");
    assert_eq!(xf, pool_entry(&part(&source, "xl/styles.xml"), "cellXfs", "xf", 1));
    assert_eq!(fill, LIGHT_UP);
    assert_eq!(border, DASH_DOT);
}

#[test]
fn a_format_authored_after_restoring_a_released_snapshot_keeps_the_pattern() {
    let (source, twin) = single();
    let mut workbook = restored(&source, &released_snapshot(&twin, None));
    let a1 = CellRef::parse_a1("A1").unwrap();
    workbook
        .patch_range_style(
            SheetId(0),
            CellRange::new(a1, a1),
            StylePatch {
                bold: Some(true),
                ..StylePatch::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let (style, _, fill, border) = saved_format(&saved, "A1");
    assert!(
        Workbook::open(&saved)
            .unwrap()
            .model()
            .styles
            .font_for(style as u32)
            .is_some_and(|font| font.bold)
    );
    assert_eq!(fill, LIGHT_UP);
    assert_eq!(border, DASH_DOT);
}

#[test]
fn source_formats_sharing_one_released_reading_restore_to_their_own_formats() {
    let cells = [("A1", 1), ("C1", 2)];
    let source = package(
        &[NONE, LIGHT_UP, DARK_GRID],
        &[EMPTY, DASH_DOT, MEDIUM_DASHED],
        &cells,
    );
    let twin = package(&[NONE, SOLID, SOLID], &[EMPTY, MEDIUM, MEDIUM], &cells);
    let saved = restored(&source, &released_snapshot(&twin, None))
        .save()
        .unwrap();
    let styles = part(&source, "xl/styles.xml");
    for (cell, xf, fill, border) in [
        ("A1", 1, LIGHT_UP, DASH_DOT),
        ("C1", 2, DARK_GRID, MEDIUM_DASHED),
    ] {
        let (_, saved_xf, saved_fill, saved_border) = saved_format(&saved, cell);
        assert_eq!(saved_xf, pool_entry(&styles, "cellXfs", "xf", xf), "{cell}");
        assert_eq!(saved_fill, fill, "{cell}");
        assert_eq!(saved_border, border, "{cell}");
    }
}

#[test]
fn a_fill_colour_authored_by_a_released_build_is_not_taken_for_the_source_pattern() {
    let (source, twin) = single();
    let snapshot = released_snapshot(
        &twin,
        Some(StylePatch {
            fill_color: Some("#00ff00".into()),
            ..StylePatch::default()
        }),
    );
    let saved = restored(&source, &snapshot).save().unwrap();
    let (style, ..) = saved_format(&saved, "A1");
    assert_eq!(
        Workbook::open(&saved)
            .unwrap()
            .model()
            .styles
            .fill_for(style as u32),
        Some(&Fill::Solid(Color::Rgb("#00ff00".into())))
    );
}
