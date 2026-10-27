use betteroffice_xlsx::{DrawCmd, SheetId, Viewport, Workbook};

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
                    r#"<worksheet xmlns="{main}"><cols><col min="1" max="1" width="12" customWidth="1"/></cols>"#,
                    r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>of which</t></is></c></row><row r="2"><c r="A2" s="1" t="inlineStr"><is><t>of which</t></is></c></row></sheetData></worksheet>"#,
                ),
                main = MAIN,
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn indented_text_starts_right_of_unindented_text() {
    let workbook =
        Workbook::open_for_read(&package(r#"<alignment horizontal="left" indent="2"/>"#)).unwrap();
    let frame = workbook
        .display_list_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 100.0,
            },
        )
        .unwrap();
    let starts = frame
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text { x, .. } => Some(*x),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(starts.len(), 2);
    assert!(starts[1] - starts[0] > 15.0, "{starts:?}");
}
