#![cfg(feature = "raster")]

use betteroffice_xlsx::{CellRange, CellRef, RenderOptions, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn package(alignment: &str) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
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
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/styles.xml",
            format!(
                concat!(
                    r#"<styleSheet xmlns="{main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
                    r#"<fills count="1"><fill><patternFill patternType="none"/></fill></fills>"#,
                    r#"<borders count="1"><border/></borders>"#,
                    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
                    r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">{alignment}</xf></cellXfs>"#,
                    r#"</styleSheet>"#,
                ),
                main = MAIN,
                alignment = alignment,
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                concat!(
                    r#"<worksheet xmlns="{main}"><cols><col min="1" max="1" width="6" customWidth="1"/></cols>"#,
                    r#"<sheetData><row r="1" ht="90" customHeight="1"><c r="A1" s="1" t="inlineStr"><is><t>Heading</t></is></c></row></sheetData></worksheet>"#,
                ),
                main = MAIN,
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Width and height of the dark-ink bounding box in the A1 render.
fn ink_extent(alignment: &str) -> (usize, usize) {
    let workbook = Workbook::open_for_read(&package(alignment)).unwrap();
    let png = workbook
        .render_sheet(
            SheetId(0),
            &RenderOptions {
                range: Some(CellRange::new(CellRef::new(0, 0), CellRef::new(0, 0))),
                ..RenderOptions::default()
            },
        )
        .unwrap();
    let mut reader = png::Decoder::new(std::io::Cursor::new(&png.bytes))
        .read_info()
        .unwrap();
    let mut buffer = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut buffer).unwrap();
    let channels = info.color_type.samples();
    let (width, height) = (info.width as usize, info.height as usize);
    let (mut left, mut top, mut right, mut bottom) = (usize::MAX, usize::MAX, 0, 0);
    for y in 2..height.saturating_sub(2) {
        for x in 2..width.saturating_sub(2) {
            let at = (y * width + x) * channels;
            if buffer[at..at + 3].iter().all(|&value| value < 96) {
                left = left.min(x);
                right = right.max(x);
                top = top.min(y);
                bottom = bottom.max(y);
            }
        }
    }
    assert!(left <= right, "no ink in A1");
    (right - left + 1, bottom - top + 1)
}

#[test]
fn rotated_text_paints_vertically() {
    let (width, height) = ink_extent(r#"<alignment textRotation="90"/>"#);
    assert!(height > width * 2, "90 degrees: {width}x{height}");
    let (width, height) = ink_extent(r#"<alignment textRotation="180"/>"#);
    assert!(height > width * 2, "-90 degrees: {width}x{height}");
}
