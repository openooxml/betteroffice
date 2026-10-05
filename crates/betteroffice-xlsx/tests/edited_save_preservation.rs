use std::collections::BTreeMap;

use betteroffice_xlsx::{CalculationOptions, CellRange, CellRef, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// cellXfs[2] resolves to the same modeled format as cellXfs[1] and differs
/// only in a facet the model does not carry.
const VARIANTS: &[(&str, &str)] = &[
    (
        "protection",
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyProtection="1"><protection locked="0"/></xf>"#,
    ),
    (
        "quote prefix",
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" quotePrefix="1"/>"#,
    ),
    (
        "cell style",
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="1"/>"#,
    ),
    (
        "apply flag",
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>"#,
    ),
    (
        "rotation",
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment textRotation="90"/></xf>"#,
    ),
    (
        "extension",
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><extLst><ext uri="{00000000-0000-0000-0000-000000000001}"/></extLst></xf>"#,
    ),
];

const SHEET1: &str = concat!(
    r#"<sheetData><row r="1"><c r="A1" s="2"><v>1</v></c><c r="B1" s="2"><v>2</v></c>"#,
    r#"<c r="C1" s="2"><f>A1*2</f><v>2</v></c><c r="D1" s="3"><v>4</v></c></row>"#,
    r#"<row r="2"><c r="A2" s="1"><v>3</v></c><c r="B2" s="2"><v>5</v></c></row></sheetData>"#,
);

const SHEET2: &str = concat!(
    r#"<sheetData><row r="1"><c r="A1" s="2"><v>1</v></c>"#,
    r#"<c r="B1" s="2"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>"#,
    r#"<row r="2"><c r="A2" s="2"><v>2</v></c><c r="B2" s="2"><f t="shared" si="0"/><v>4</v></c></row>"#,
    r#"<row r="3"><c r="A3" s="2"><v>3</v></c><c r="B3" s="2"><f t="shared" si="0"/><v>6</v></c></row>"#,
    r#"</sheetData>"#,
);

fn styles(variant: &str) -> String {
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<styleSheet xmlns="{main}">"#,
            r#"<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>"#,
            r#"<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>"#,
            r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
            r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
            r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
            r#"<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#,
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
            r#"<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
            r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
            r#"{variant}"#,
            r#"<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>"#,
            r#"</cellXfs><cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/>"#,
            r#"<cellStyle name="Custom" xfId="1"/></cellStyles></styleSheet>"#,
        ),
        main = MAIN,
        variant = variant,
    )
}

fn worksheet(sheet_data: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}">{sheet_data}</worksheet>"#
    )
}

fn package(variant: &str) -> Vec<u8> {
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
                r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
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
                    r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/>"#,
                    r#"<sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
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
                r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>"#,
                r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        ("xl/styles.xml", styles(variant)),
        ("xl/worksheets/sheet1.xml", worksheet(SHEET1)),
        ("xl/worksheets/sheet2.xml", worksheet(SHEET2)),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn text(parts: &BTreeMap<String, Vec<u8>>, path: &str) -> String {
    String::from_utf8(parts[path].clone()).unwrap()
}

/// Each `<c>` element's markup keyed by its `r` attribute.
fn cells(xml: &str) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();
    let mut rest = xml;
    while let Some(start) = rest.find("<c ") {
        let tail = &rest[start..];
        let tag_end = tail.find('>').unwrap();
        let end = if tail[..tag_end].ends_with('/') {
            tag_end + 1
        } else {
            tail.find("</c>").unwrap() + "</c>".len()
        };
        let span = &tail[..end];
        let address = span
            .split("r=\"")
            .nth(1)
            .unwrap()
            .split('"')
            .next()
            .unwrap();
        out.insert(address.to_owned(), span.to_owned());
        rest = &tail[end..];
    }
    out
}

fn style_index(span: &str) -> usize {
    let tag = &span[..span.find('>').unwrap()];
    tag.split(" s=\"")
        .nth(1)
        .unwrap()
        .split('"')
        .next()
        .unwrap()
        .parse()
        .unwrap()
}

/// The `index`th `cellXfs` entry's markup.
fn cell_xf(styles: &str, index: usize) -> String {
    let table = &styles[styles.find("<cellXfs").unwrap()..];
    let table = &table[..table.find("</cellXfs>").unwrap()];
    table.split("<xf ").nth(index + 1).unwrap().to_owned()
}

fn cell(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

#[test]
fn unrelated_edit_keeps_equivalent_style_indices_and_other_parts() {
    for (name, variant) in VARIANTS {
        let source = package(variant);
        let before = parts(&source);
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(SheetId(0), cell("A1"), "5", CalculationOptions::default())
            .unwrap();
        let after = parts(&workbook.save().unwrap());

        for path in ["xl/styles.xml", "xl/worksheets/sheet2.xml"] {
            assert_eq!(after[path], before[path], "{name}: {path}");
        }
        let source_cells = cells(&text(&before, "xl/worksheets/sheet1.xml"));
        let saved_cells = cells(&text(&after, "xl/worksheets/sheet1.xml"));
        for address in ["B1", "D1", "A2", "B2"] {
            assert_eq!(
                saved_cells[address], source_cells[address],
                "{name}: {address}"
            );
        }
        let edited = &saved_cells["A1"];
        assert!(
            edited.contains(r#"s="2""#) && edited.contains("<v>5</v>"),
            "{name}: {edited}"
        );
        let dependent = &saved_cells["C1"];
        assert!(
            dependent.contains(r#"s="2""#)
                && dependent.contains("<f>A1*2</f>")
                && dependent.contains("<v>10</v>"),
            "{name}: {dependent}"
        );
    }
}

#[test]
fn genuine_style_change_writes_the_new_index() {
    for (name, variant) in VARIANTS {
        let source = package(variant);
        let before = parts(&source);
        let mut workbook = Workbook::open(&source).unwrap();
        let format = workbook
            .capture_format(SheetId(0), CellRange::new(cell("D1"), cell("D1")))
            .unwrap();
        workbook
            .apply_format(
                SheetId(0),
                CellRange::new(cell("B1"), cell("B1")),
                format,
                CalculationOptions::default(),
            )
            .unwrap();
        workbook
            .edit_cell(SheetId(0), cell("B2"), "7", CalculationOptions::default())
            .unwrap();
        let after = parts(&workbook.save().unwrap());

        assert_eq!(
            after["xl/worksheets/sheet2.xml"], before["xl/worksheets/sheet2.xml"],
            "{name}"
        );
        let saved_cells = cells(&text(&after, "xl/worksheets/sheet1.xml"));
        let restyled = &saved_cells["B1"];
        let index = style_index(restyled);
        let xf = cell_xf(&text(&after, "xl/styles.xml"), index);
        assert!(
            index > 2
                && xf.contains(r#"numFmtId="2""#)
                && xf.contains(r#"fontId="1""#)
                && restyled.contains("<v>2</v>"),
            "{name}: {restyled} {xf}"
        );
        let edited = &saved_cells["B2"];
        assert!(
            edited.contains(r#"s="2""#) && edited.contains("<v>7</v>"),
            "{name}: {edited}"
        );
    }
}
