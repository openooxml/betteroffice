use std::collections::BTreeMap;

use betteroffice_xlsx::{CalculationOptions, CellRange, CellRef, SheetId, StylePatch, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const GRADIENT: &str = concat!(
    r#"<fill><gradientFill degree="90"><stop position="0"><color rgb="FF123456"/></stop>"#,
    r#"<stop position="1"><color rgb="FFFEDCBA"/></stop></gradientFill></fill>"#,
);

fn package() -> Vec<u8> {
    let styles = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{main}">"#,
            r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
            r#"<fills count="3"><fill><patternFill patternType="none"/></fill>"#,
            r#"<fill><patternFill patternType="gray125"/></fill>{gradient}</fills>"#,
            r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
            r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
            r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
            r#"<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>"#,
            r#"</styleSheet>"#,
        ),
        main = MAIN,
        gradient = GRADIENT,
    );
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
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>"#
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
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

#[test]
#[ignore = "deferred: Fill has no gradient variant, so a format edit regenerates the xf with no fill"]
fn format_edit_keeps_gradient_fill_on_save() {
    let mut workbook = Workbook::open(&package()).unwrap();
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
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let styles = String::from_utf8(parts["xl/styles.xml"].clone()).unwrap();
    let sheet = String::from_utf8(parts["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let cell = &sheet[sheet.find(r#"<c r="A1""#).unwrap()..];
    let style: usize = attribute(cell, "s").parse().unwrap();

    let reopened = Workbook::open(&saved).unwrap();
    assert!(
        reopened
            .model()
            .styles
            .font_for(style as u32)
            .is_some_and(|font| font.bold),
        "{styles}"
    );
    let xf = pool_entry(&styles, "cellXfs", "xf", style);
    let fill: usize = attribute(xf, "fillId").parse().unwrap();
    assert_eq!(pool_entry(&styles, "fills", "fill", fill), GRADIENT, "{xf}");
}
