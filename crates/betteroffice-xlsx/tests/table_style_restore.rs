use betteroffice_xlsx::{
    DisplayList, DrawCmd, SheetId, SnapshotBudget, Viewport, Workbook, WorkbookSnapshotBuilder,
    WorkbookSnapshotEncoder,
};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const STYLES: &str = r#"<styleSheet><fonts count="6"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><sz val="11"/><color theme="1" tint="0.25"/><name val="Calibri"/></font><font><sz val="12"/><color theme="1" tint="0.25"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="2"><xf fontId="3" fillId="0" borderId="0"/><xf fontId="4" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" fillId="0" borderId="0" xfId="0"/><xf fontId="2" fillId="0" borderId="0" xfId="0"/><xf fontId="5" fillId="0" borderId="0" xfId="1"/><xf fontId="2" fillId="0" borderId="0" xfId="1"/></cellXfs><cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Muted" xfId="1"/></cellStyles></styleSheet>"#;

/// Table `A1:E3` in `TableStyleMedium9`, header cell `hN` on xf `N`.
fn source() -> Vec<u8> {
    let header = (0..5)
        .map(|index| {
            let col = char::from(b'A' + index);
            format!(r#"<c r="{col}1" s="{index}" t="inlineStr"><is><t>h{index}</t></is></c>"#)
        })
        .collect::<String>();
    let data = (2..=3)
        .map(|row| {
            let cells = (0..5)
                .map(|index| {
                    let col = char::from(b'A' + index);
                    format!(r#"<c r="{col}{row}"><v>{row}</v></c>"#)
                })
                .collect::<String>();
            format!(r#"<row r="{row}">{cells}</row>"#)
        })
        .collect::<String>();
    let columns = (0..5)
        .map(|index| format!(r#"<tableColumn id="{}" name="h{index}"/>"#, index + 1))
        .collect::<String>();
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "xl/workbook.xml".into(),
            format!(r#"<workbook xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#).into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".into(),
            format!(r#"<Relationships><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#).into_bytes(),
        ),
        ("xl/styles.xml".into(), STYLES.as_bytes().to_vec()),
        (
            "xl/worksheets/sheet1.xml".into(),
            format!(r#"<worksheet xmlns:r="{REL}"><sheetData><row r="1">{header}</row>{data}</sheetData><tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>"#).into_bytes(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            format!(r#"<Relationships><Relationship Id="t1" Type="{REL}/table" Target="../tables/table1.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/tables/table1.xml".into(),
            format!(r#"<table id="1" name="T" displayName="T" ref="A1:E3"><tableColumns count="5">{columns}</tableColumns><tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>"#).into_bytes(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn render(workbook: &Workbook) -> DisplayList {
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: 800.0,
        height: 200.0,
    };
    workbook.display_list_for(SheetId(0), &viewport).unwrap()
}

fn text(list: &DisplayList, wanted: &str) -> (String, bool) {
    list.commands
        .iter()
        .find_map(|command| match command {
            DrawCmd::Text {
                text, color, bold, ..
            } if &**text == wanted => Some((color.to_string(), *bold)),
            _ => None,
        })
        .unwrap()
}

fn snapshot_round_trip(workbook: &Workbook) -> Workbook {
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    builder.finish().unwrap().into_parts().0
}

#[test]
fn header_font_colours_follow_the_cell_style_font() {
    let list = render(&Workbook::open(&source()).unwrap());
    assert_eq!(text(&list, "h0"), ("#ffffff".into(), true));
    assert_eq!(text(&list, "h1"), ("#000000".into(), true));
    assert_eq!(text(&list, "h2"), ("#ffffff".into(), true));
    assert_eq!(text(&list, "h3"), ("#ffffff".into(), true));
    assert_eq!(text(&list, "h4"), ("#000000".into(), true));
}

#[test]
fn a_snapshot_restore_draws_table_fonts_as_the_fresh_open() {
    let source = source();
    let fresh = Workbook::open(&source).unwrap();
    assert_eq!(render(&snapshot_round_trip(&fresh)), render(&fresh));
    let collaborative = Workbook::open_collaborative(&source, 811).unwrap();
    assert_eq!(
        render(&snapshot_round_trip(&collaborative)),
        render(&collaborative)
    );
}

#[test]
fn a_recovery_restore_draws_table_fonts_as_the_fresh_open() {
    let fresh = Workbook::open(&source()).unwrap();
    let checkpoint = fresh.recovery_checkpoint(0, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    assert_eq!(render(&restored), render(&fresh));
}
