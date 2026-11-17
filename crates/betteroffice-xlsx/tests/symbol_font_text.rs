use betteroffice_xlsx::{CellRef, CellValue, DrawCmd, SheetId, Viewport, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn package() -> Vec<u8> {
    let rows = concat!(
        r#"<row r="1"><c r="A1" s="1" t="inlineStr"><is><t xml:space="preserve">a+b &#xF0B3; p</t></is></c></row>"#,
        r#"<row r="2"><c r="A2" s="2" t="inlineStr"><is><t>l&#xF0E0;</t></is></c></row>"#,
        r#"<row r="3"><c r="A3" t="inlineStr"><is><t>ab</t></is></c></row>"#,
    );
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
                    r#"<styleSheet xmlns="{main}"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Symbol"/></font><font><sz val="11"/><name val="Wingdings"/></font></fonts>"#,
                    r#"<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>"#,
                    r#"<borders count="1"><border/></borders>"#,
                    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
                    r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"#,
                    r#"<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>"#,
                    r#"</styleSheet>"#,
                ),
                main = MAIN,
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<worksheet xmlns="{MAIN}"><sheetData>{rows}</sheetData></worksheet>"#
            ),
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn drawn(workbook: &Workbook) -> Vec<String> {
    workbook
        .display_list(&Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 120.0,
        })
        .unwrap()
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text { text, .. } => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

fn stored(workbook: &Workbook) -> Vec<String> {
    (0..3)
        .map(|row| {
            match &workbook
                .sheet(SheetId(0))
                .unwrap()
                .cell(CellRef::new(row, 0))
                .unwrap()
                .value
            {
                CellValue::Text { value } => value.to_string(),
                other => panic!("{other:?}"),
            }
        })
        .collect()
}

#[test]
fn symbol_face_cells_draw_the_characters_their_slots_address() {
    let bytes = package();
    let workbook = Workbook::open(&bytes).unwrap();
    assert_eq!(drawn(&workbook), ["α+β ≥ π", "●\u{f0e0}", "ab"]);

    let expected = ["a+b \u{f0b3} p", "l\u{f0e0}", "ab"];
    assert_eq!(stored(&workbook), expected);
    let saved = workbook.save().unwrap();
    assert_eq!(stored(&Workbook::open(&saved).unwrap()), expected);
}
