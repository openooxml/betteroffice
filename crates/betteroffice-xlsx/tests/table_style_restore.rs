use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, DisplayList, DrawCmd, NumberFormatMutation, SheetId,
    SnapshotBudget, Viewport, Workbook, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder,
};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const STYLES: &str = r#"<styleSheet><fonts count="6"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><sz val="11"/><color theme="1" tint="0.25"/><name val="Calibri"/></font><font><sz val="12"/><color theme="1" tint="0.25"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="2"><xf fontId="3" fillId="0" borderId="0"/><xf fontId="4" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" fillId="0" borderId="0" xfId="0"/><xf fontId="2" fillId="0" borderId="0" xfId="0"/><xf fontId="5" fillId="0" borderId="0" xfId="1"/><xf fontId="2" fillId="0" borderId="0" xfId="1"/></cellXfs><cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Muted" xfId="1"/></cellStyles></styleSheet>"#;

const MEDIUM9: &str = r#"<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"#;

const CUSTOM_STYLES: &str = r#"<dxfs count="2"><dxf><font><b/><color rgb="FFC00000"/></font><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf><dxf><border><bottom style="thin"><color rgb="FF2F5597"/></bottom></border></dxf></dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="Custom" pivot="0" count="2"><tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="firstColumn" dxfId="1"/></tableStyle></tableStyles>"#;

const CUSTOM: &str = r#"<tableStyleInfo name="Custom" showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="1"/>"#;

/// Table `A1:E3` in `TableStyleMedium9`, header cell `hN` on xf `N`.
fn source() -> Vec<u8> {
    source_with(STYLES, MEDIUM9)
}

fn source_with(styles: &str, table_style: &str) -> Vec<u8> {
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
        ("xl/styles.xml".into(), styles.as_bytes().to_vec()),
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
            format!(r#"<table id="1" name="T" displayName="T" ref="A1:E3"><tableColumns count="5">{columns}</tableColumns>{table_style}</table>"#).into_bytes(),
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
}

#[test]
fn an_adopted_collaboration_snapshot_draws_table_fonts_as_the_fresh_open() {
    let source = source();
    let fresh = Workbook::open(&source).unwrap();
    let writer = Workbook::open_collaborative(&source, 811).unwrap();
    let mut reader = Workbook::open_collaborative(&source, 812).unwrap();
    reader
        .apply_update_v1(
            &writer.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(render(&reader), render(&fresh));
}

fn recovered(workbook: &Workbook) -> Workbook {
    let checkpoint = workbook.recovery_checkpoint(0, None).unwrap();
    Workbook::restore_recovery_checkpoint(&checkpoint)
        .unwrap()
        .0
}

#[test]
fn recovery_shows_tables_without_their_style_until_reopened() {
    let source = source();
    let fresh = Workbook::open(&source).unwrap();
    let restored = recovered(&fresh);
    assert!(!fresh.model().tables[0].style.is_empty());
    assert!(restored.model().tables[0].style.is_empty());
    let unstyled = Workbook::open(&source_with(STYLES, "")).unwrap();
    assert_eq!(render(&restored), render(&unstyled));
    assert_ne!(render(&restored), render(&fresh));
    let reopened = Workbook::open(&restored.save().unwrap()).unwrap();
    assert_eq!(render(&reopened), render(&fresh));
}

#[test]
fn a_recovery_restore_draws_cell_style_fonts_as_the_fresh_open() {
    let fresh = Workbook::open(&source_with(STYLES, "")).unwrap();
    let restored = recovered(&fresh);
    let headers = |list: &DisplayList| {
        (0..5)
            .map(|index| text(list, &format!("h{index}")))
            .collect::<Vec<_>>()
    };
    assert_eq!(headers(&render(&restored)), headers(&render(&fresh)));
    assert_eq!(render(&restored), render(&fresh));
}

fn part(parts: &[(String, Vec<u8>)], path: &str) -> String {
    parts
        .iter()
        .find(|(name, _)| name == path)
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap()
}

fn element(xml: &str, name: &str) -> String {
    let start = xml.find(&format!("<{name}")).unwrap();
    let end = xml[start..].find(&format!("</{name}>")).unwrap() + start + name.len() + 3;
    xml[start..end].to_owned()
}

#[test]
fn a_recovery_restore_saves_table_styles_as_the_source() {
    let source = source_with(
        &STYLES.replace("</styleSheet>", &format!("{CUSTOM_STYLES}</styleSheet>")),
        CUSTOM,
    );
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 6),
            "7",
            CalculationOptions::default(),
        )
        .unwrap();
    workbook
        .set_range_number_format(
            SheetId(0),
            CellRange::new(CellRef::new(0, 6), CellRef::new(0, 6)),
            NumberFormatMutation::Custom {
                pattern: "0.00".into(),
            },
            CalculationOptions::default(),
        )
        .unwrap();
    let checkpoint = workbook.recovery_checkpoint(2, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
    let original = ooxml_opc::unzip_parts(&source).unwrap();
    let saved = ooxml_opc::unzip_parts(&restored.save().unwrap()).unwrap();
    assert_eq!(
        part(&saved, "xl/tables/table1.xml"),
        part(&original, "xl/tables/table1.xml")
    );
    assert!(part(&saved, "xl/tables/table1.xml").contains(CUSTOM));
    let saved_styles = part(&saved, "xl/styles.xml");
    let original_styles = part(&original, "xl/styles.xml");
    assert_ne!(saved_styles, original_styles);
    for name in ["dxfs", "tableStyles"] {
        assert_eq!(
            element(&saved_styles, name),
            element(&original_styles, name)
        );
    }
}
