use std::collections::BTreeMap;

use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellState, CellValue, Error, Op, SheetId, Workbook,
};
use xlsx_parse::{
    SaveEdits, SharedStringCells, parse_workbook_with_package,
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet,
};

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
    package_with_second_sheet(variant, SHEET2)
}

fn package_with_second_sheet(variant: &str, second_sheet: &str) -> Vec<u8> {
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
        ("xl/worksheets/sheet2.xml", worksheet(second_sheet)),
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

const VIEW_SHEETS: &str = concat!(
    r#"<sheets><sheet name="Forecast" sheetId="1" r:id="rId1"/>"#,
    r#"<sheet name="Assumptions" sheetId="2" r:id="rId2"/>"#,
    r#"<sheet name="Actuals" sheetId="3" r:id="rId3"/></sheets>"#,
);

fn view_package(views: &str, active: usize, chartsheet: Option<usize>) -> Vec<u8> {
    let mut entries = parts(&package(VARIANTS[0].1));
    entries.insert("xl/workbook.xml".into(), format!(
        r#"<workbook xmlns="{MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:test"><x:before value='A &amp; B'/>{views}{VIEW_SHEETS}<calcPr fullCalcOnLoad="1"/><x:after/></workbook>"#,
    ).into_bytes());
    let rels = text(&entries, "xl/_rels/workbook.xml.rels")
        .replace("Id=\"rId3\"", "Id=\"rId4\"")
        .replace("</Relationships>", r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>"#);
    entries.insert("xl/_rels/workbook.xml.rels".into(), rels.into_bytes());
    let types = text(&entries, "[Content_Types].xml").replace("</Types>", r#"<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#);
    entries.insert("[Content_Types].xml".into(), types.into_bytes());
    for index in 0..3 {
        let root = if chartsheet == Some(index) {
            "chartsheet"
        } else {
            "worksheet"
        };
        let selected = usize::from(index == active);
        entries.insert(format!("xl/worksheets/sheet{}.xml", index + 1), format!(
            r#"<{root} xmlns="{MAIN}"><sheetViews><sheetView zoomScale='115' tabSelected="{selected}" workbookViewId='0'><selection activeCell='B2' sqref='B2'/></sheetView></sheetViews><sheetData/></{root}>"#,
        ).into_bytes());
    }
    if let Some(index) = chartsheet {
        let path = "xl/_rels/workbook.xml.rels";
        let rels = text(&entries, path).replace(
            &format!(
                "relationships/worksheet\" Target=\"worksheets/sheet{}.xml",
                index + 1
            ),
            &format!(
                "relationships/chartsheet\" Target=\"worksheets/sheet{}.xml",
                index + 1
            ),
        );
        entries.insert(path.into(), rels.into_bytes());
    }
    ooxml_opc::rezip_parts(&entries.into_iter().collect::<Vec<_>>()).unwrap()
}

fn assert_saved_view(
    workbook: &Workbook,
    source: &[u8],
    old_view: &str,
    new_view: &str,
    sheets: &str,
) -> Vec<u8> {
    let saved = workbook.save().unwrap();
    assert_eq!(
        text(&parts(&saved), "xl/workbook.xml"),
        text(&parts(source), "xl/workbook.xml")
            .replace(old_view, new_view)
            .replace(VIEW_SHEETS, sheets),
    );
    assert_eq!(
        Workbook::open(&saved).unwrap().active_sheet(),
        workbook.active_sheet()
    );
    saved
}

#[test]
fn workbook_view_implicit_indices_follow_rename_and_move() {
    let view = "<bookViews><workbookView/></bookViews>";
    let source = view_package(view, 0, None);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .apply_ops(
            vec![
                Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "Outlook".into(),
                },
                Op::MoveSheet { from: 2, to: 0 },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(workbook.active_sheet(), SheetId(1));
    assert_saved_view(
        &workbook,
        &source,
        view,
        "<bookViews><workbookView activeTab=\"1\" firstSheet=\"1\"/></bookViews>",
        r#"<sheets><sheet name="Actuals" sheetId="3" r:id="rId3"/><sheet name="Outlook" sheetId="1" r:id="rId1"/><sheet name="Assumptions" sheetId="2" r:id="rId2"/></sheets>"#,
    );
}

#[test]
fn workbook_view_explicit_indices_preserve_other_attributes_and_children() {
    let view = r#"<bookViews><workbookView xWindow='42' activeTab="2" firstSheet='1' showSheetTabs = "1"><extLst><ext uri='urn:test'><x:marker a='A &amp; B'/></ext></extLst></workbookView></bookViews>"#;
    let source = view_package(view, 2, None);
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(workbook.active_sheet(), SheetId(2));
    workbook
        .apply_ops(
            vec![Op::MoveSheet { from: 2, to: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    let expected = view
        .replace("activeTab=\"2\"", "activeTab=\"0\"")
        .replace("firstSheet='1'", "firstSheet='2'");
    assert_saved_view(
        &workbook,
        &source,
        view,
        &expected,
        r#"<sheets><sheet name="Actuals" sheetId="3" r:id="rId3"/><sheet name="Forecast" sheetId="1" r:id="rId1"/><sheet name="Assumptions" sheetId="2" r:id="rId2"/></sheets>"#,
    );
}

#[test]
fn workbook_view_remove_shifts_and_clamps_indices_and_selection() {
    let view = r#"<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>"#;
    let source = view_package(view, 2, None);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_saved_view(
        &workbook,
        &source,
        view,
        r#"<bookViews><workbookView activeTab="1" firstSheet="0"/></bookViews>"#,
        r#"<sheets><sheet name="Assumptions" sheetId="2" r:id="rId2"/><sheet name="Actuals" sheetId="3" r:id="rId3"/></sheets>"#,
    );
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = assert_saved_view(
        &workbook,
        &source,
        view,
        r#"<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>"#,
        r#"<sheets><sheet name="Assumptions" sheetId="2" r:id="rId2"/></sheets>"#,
    );
    assert_eq!(
        text(&parts(&saved), "xl/worksheets/sheet2.xml"),
        text(&parts(&source), "xl/worksheets/sheet2.xml")
            .replace("tabSelected=\"0\"", "tabSelected=\"1\"")
    );
}

#[test]
fn workbook_view_cell_edit_preserves_workbook_bytes() {
    let view = r#"<bookViews><workbookView activeTab='2' firstSheet='1' showSheetTabs = "1"><extLst/></workbookView></bookViews>"#;
    let source = view_package(view, 2, None);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "42",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_saved_view(&workbook, &source, view, view, VIEW_SHEETS);
}

#[test]
fn workbook_view_undo_and_redo_restore_source_indices() {
    let view = r#"<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>"#;
    let source = view_package(view, 2, None);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .apply_ops(
            vec![Op::MoveSheet { from: 2, to: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    workbook.undo(CalculationOptions::default()).unwrap();
    assert_saved_view(&workbook, &source, view, view, VIEW_SHEETS);
    workbook.redo(CalculationOptions::default()).unwrap();
    assert_saved_view(
        &workbook,
        &source,
        view,
        r#"<bookViews><workbookView activeTab="0" firstSheet="2"/></bookViews>"#,
        r#"<sheets><sheet name="Actuals" sheetId="3" r:id="rId3"/><sheet name="Forecast" sheetId="1" r:id="rId1"/><sheet name="Assumptions" sheetId="2" r:id="rId2"/></sheets>"#,
    );
}

#[test]
fn workbook_view_add_keeps_source_tabs_and_first_sheet() {
    let view = "<bookViews><workbookView/></bookViews>";
    let source = view_package(view, 0, None);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "New".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = parts(&workbook.save().unwrap());
    let xml = text(&saved, "xl/workbook.xml");
    assert!(
        xml.contains(r#"<workbookView activeTab="1" firstSheet="1"/>"#),
        "{xml}"
    );
    assert_eq!(
        Workbook::open(&workbook.save().unwrap())
            .unwrap()
            .active_sheet(),
        SheetId(1)
    );
}

#[test]
fn workbook_view_active_change_updates_chartsheet_selection() {
    let view = "<bookViews><workbookView/></bookViews>";
    let source = view_package(view, 0, Some(2));
    let mut workbook = Workbook::open(&source).unwrap();
    workbook.set_active_sheet(SheetId(2)).unwrap();
    let saved = assert_saved_view(
        &workbook,
        &source,
        view,
        r#"<bookViews><workbookView activeTab="2"/></bookViews>"#,
        VIEW_SHEETS,
    );
    for (part, from, to) in [(1, 1, 0), (3, 0, 1)] {
        let path = format!("xl/worksheets/sheet{part}.xml");
        assert_eq!(
            text(&parts(&saved), &path),
            text(&parts(&source), &path).replace(
                &format!("tabSelected=\"{from}\""),
                &format!("tabSelected=\"{to}\"")
            )
        );
    }
    workbook
        .apply_ops(
            vec![Op::MoveSheet { from: 2, to: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = assert_saved_view(
        &workbook,
        &source,
        view,
        r#"<bookViews><workbookView activeTab="0" firstSheet="1"/></bookViews>"#,
        r#"<sheets><sheet name="Actuals" sheetId="3" r:id="rId3"/><sheet name="Forecast" sheetId="1" r:id="rId1"/><sheet name="Assumptions" sheetId="2" r:id="rId2"/></sheets>"#,
    );
    assert!(text(&parts(&saved), "xl/worksheets/sheet3.xml").contains("tabSelected=\"1\""));
}

const GROUPED_VIEW: &str = "<bookViews><workbookView activeTab=\"0\"/></bookViews>";

/// Forecast is active and grouped with Assumptions in view 0. A non-empty
/// `second` adds view 1 sheetViews selecting those sheet indices.
fn grouped_workbook(views: &str, second: &[usize]) -> Workbook {
    Workbook::open(&grouped_source(views, second)).unwrap()
}

fn grouped_source(views: &str, second: &[usize]) -> Vec<u8> {
    let mut entries = parts(&view_package(views, 0, None));
    for sheet in 0..3 {
        let path = format!("xl/worksheets/sheet{}.xml", sheet + 1);
        let mut xml = text(&entries, &path);
        if sheet == 1 {
            xml = xml.replace("tabSelected=\"0\"", "tabSelected=\"1\"");
        }
        if !second.is_empty() {
            let view = format!(
                "<sheetView workbookViewId=\"1\" tabSelected=\"{}\"/></sheetViews>",
                u8::from(second.contains(&sheet))
            );
            xml = xml.replace("</sheetViews>", &view);
        }
        entries.insert(path, xml.into_bytes());
    }
    ooxml_opc::rezip_parts(&entries.into_iter().collect::<Vec<_>>()).unwrap()
}

fn view_selected(view: &str, id: &str) -> bool {
    let element = view.split('>').next().unwrap_or_default();
    element.contains(id) && element.contains("tabSelected=\"1\"")
}

fn selected_sheets(saved: &[u8], view: usize) -> Vec<String> {
    let id = format!("workbookViewId=\"{view}\"");
    let mut selected = Vec::new();
    for (path, xml) in parts(saved) {
        let xml = String::from_utf8_lossy(&xml).replace('\'', "\"");
        let mut views = xml.split("<sheetView ").skip(1);
        if path.starts_with("xl/worksheets/") && views.any(|view| view_selected(view, &id)) {
            selected.push(path);
        }
    }
    selected
}

const GROUP: [&str; 2] = ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"];

fn add_last_sheet(workbook: &mut Workbook) {
    workbook
        .apply_ops(
            vec![Op::AddSheet {
                index: 3,
                name: "New".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

#[test]
fn workbook_view_group_survives_added_sheet() {
    let mut workbook = grouped_workbook(GROUPED_VIEW, &[]);
    add_last_sheet(&mut workbook);
    let saved = workbook.save().unwrap();
    assert!(text(&parts(&saved), "xl/workbook.xml").contains("name=\"New\""));
    assert_eq!(selected_sheets(&saved, 0), GROUP);
}

#[test]
fn workbook_view_group_survives_move_with_same_active_sheet() {
    let mut workbook = grouped_workbook(GROUPED_VIEW, &[]);
    workbook
        .apply_ops(
            vec![Op::MoveSheet { from: 2, to: 0 }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(workbook.active_sheet(), SheetId(1));
    assert_eq!(selected_sheets(&workbook.save().unwrap(), 0), GROUP);
}

#[test]
fn workbook_view_group_survives_edit_and_added_sheet() {
    let mut workbook = grouped_workbook(GROUPED_VIEW, &[]);
    workbook
        .edit_cell(
            SheetId(1),
            CellRef::new(0, 0),
            "42",
            CalculationOptions::default(),
        )
        .unwrap();
    add_last_sheet(&mut workbook);
    let saved = workbook.save().unwrap();
    assert!(text(&parts(&saved), "xl/worksheets/sheet2.xml").contains("<v>42</v>"));
    assert_eq!(selected_sheets(&saved, 0), GROUP);
}

#[test]
fn workbook_view_group_survives_removing_an_ungrouped_sheet() {
    let mut workbook = grouped_workbook(GROUPED_VIEW, &[]);
    workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 2 }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(selected_sheets(&workbook.save().unwrap(), 0), GROUP);
}

#[test]
fn workbook_view_second_view_group_survives_move() {
    let views =
        r#"<bookViews><workbookView activeTab="0"/><workbookView activeTab="2"/></bookViews>"#;
    let mut workbook = grouped_workbook(views, &[1, 2]);
    workbook
        .apply_ops(
            vec![Op::MoveSheet { from: 0, to: 2 }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(workbook.active_sheet(), SheetId(2));
    let saved = workbook.save().unwrap();
    assert_eq!(selected_sheets(&saved, 0), GROUP);
    assert_eq!(
        selected_sheets(&saved, 1),
        ["xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]
    );
}

#[test]
fn workbook_view_activating_outside_the_group_collapses_selection() {
    let mut workbook = grouped_workbook(GROUPED_VIEW, &[]);
    workbook.set_active_sheet(SheetId(2)).unwrap();
    assert_eq!(
        selected_sheets(&workbook.save().unwrap(), 0),
        ["xl/worksheets/sheet3.xml"]
    );
}

/// Saves the grouped workbook with a copy of source sheet `duplicate`
/// appended, so two sheets name that source as their origin.
fn save_with_duplicate(
    views: &str,
    second: &[usize],
    duplicate: usize,
    active: SheetId,
) -> Vec<u8> {
    let source = grouped_source(views, second);
    let entries: Vec<_> = parts(&source).into_iter().collect();
    let mut parsed = parse_workbook_with_package(&entries).unwrap();
    let mut copy = parsed.workbook.sheets[duplicate].clone();
    copy.name = "Duplicate".into();
    parsed.workbook.sheets.push(copy);
    let edits = SaveEdits {
        changed: true,
        moved_references: false,
    };
    let saved = serialize_workbook_with_package_and_origins_after_edits_and_active_sheet(
        &parsed.workbook,
        &parsed.package,
        &[Some(0), Some(1), Some(2), Some(duplicate)],
        &vec![SharedStringCells::new(); 4],
        edits,
        active,
    )
    .unwrap();
    let saved = saved
        .into_iter()
        .map(|(path, bytes)| (path, bytes.into_owned()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&saved).unwrap()
}

#[test]
fn workbook_view_activating_a_duplicate_collapses_the_source_group() {
    let saved = save_with_duplicate(GROUPED_VIEW, &[], 0, SheetId(3));
    let workbook = text(&parts(&saved), "xl/workbook.xml");
    assert!(
        workbook.contains(r#"<workbookView activeTab="3"/>"#),
        "{workbook}"
    );
    assert_eq!(selected_sheets(&saved, 0), ["xl/worksheets/sheet4.xml"]);
}

#[test]
fn workbook_view_duplicate_of_the_active_sheet_stays_unselected() {
    let saved = save_with_duplicate(GROUPED_VIEW, &[], 0, SheetId(0));
    let workbook = text(&parts(&saved), "xl/workbook.xml");
    assert!(workbook.contains(GROUPED_VIEW), "{workbook}");
    assert_eq!(selected_sheets(&saved, 0), GROUP);
}

#[test]
fn workbook_view_duplicate_of_a_grouped_sheet_stays_unselected() {
    let saved = save_with_duplicate(GROUPED_VIEW, &[], 1, SheetId(0));
    assert_eq!(selected_sheets(&saved, 0), GROUP);
}

#[test]
fn workbook_view_second_view_maps_a_duplicated_tab_to_its_first_claim() {
    let views =
        r#"<bookViews><workbookView activeTab="0"/><workbookView activeTab="1"/></bookViews>"#;
    let saved = save_with_duplicate(views, &[1, 2], 1, SheetId(0));
    let workbook = text(&parts(&saved), "xl/workbook.xml");
    assert!(workbook.contains(views), "{workbook}");
    assert_eq!(selected_sheets(&saved, 0), GROUP);
    assert_eq!(
        selected_sheets(&saved, 1),
        ["xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]
    );
}

fn implicit_font_package() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
        ),
        (
            "xl/styles.xml",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="DejaVu Sans"/></font><font></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">Item</t></is></c><c r="B1" t="inlineStr"><is><t xml:space="preserve">Revenue</t></is></c><c r="C1" t="inlineStr"><is><t xml:space="preserve">Cost</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t xml:space="preserve">Cap</t></is></c><c r="B2" s="1"><v>120</v></c><c r="C2" s="1"><v>50</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t xml:space="preserve">Bag</t></is></c><c r="B3" s="1"><v>250</v></c><c r="C3" s="1"><v>100</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t xml:space="preserve">Bottle</t></is></c><c r="B4" s="1"><v>80</v></c><c r="C4" s="1"><v>30</v></c></row></sheetData></worksheet>"#,
        ),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn formatting_an_implicit_header_does_not_add_an_empty_font() {
    let source = implicit_font_package();
    let mut workbook = Workbook::open(&source).unwrap();
    let empty_fonts = text(&parts(&source), "xl/styles.xml")
        .matches("<font></font>")
        .count();
    assert_eq!(
        text(&parts(&workbook.save().unwrap()), "xl/styles.xml")
            .matches("<font></font>")
            .count(),
        empty_fonts
    );
    workbook
        .apply_ops(
            vec![Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:C1").unwrap(),
                patch: xlsx_ops::StylePatch {
                    bold: Some(true),
                    fill_color: Some("#D9EAF7".into()),
                    ..Default::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        text(&parts(&workbook.save().unwrap()), "xl/styles.xml")
            .matches("<font></font>")
            .count(),
        empty_fonts
    );
}

#[test]
fn clearing_fill_from_red_xf_zero_stores_an_explicit_default() {
    let mut source = parts(&package(VARIANTS[0].1));
    source.insert(
        "xl/worksheets/sheet2.xml".into(),
        worksheet("<sheetData/>").into_bytes(),
    );
    source.insert("xl/styles.xml".into(), br#"<styleSheet><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills><cellXfs count="1"><xf fillId="1"/></cellXfs></styleSheet>"#.to_vec());
    source.insert(
        "xl/worksheets/sheet1.xml".into(),
        worksheet(r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#).into_bytes(),
    );
    let source = ooxml_opc::rezip_parts(&source.into_iter().collect::<Vec<_>>()).unwrap();
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .patch_range_style(
            SheetId(0),
            CellRange::parse_a1("A1").unwrap(),
            xlsx_ops::StylePatch {
                clear: vec![xlsx_ops::StyleProperty::FillColor],
                ..Default::default()
            },
            CalculationOptions::default(),
        )
        .unwrap();
    assert!(
        workbook.model().sheets[0]
            .cell(cell("A1"))
            .unwrap()
            .style
            .is_some()
    );
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    for workbook in [&workbook, &reopened] {
        let cell = workbook.model().sheets[0].cell(cell("A1")).unwrap();
        assert_eq!(
            workbook.model().styles.cell_format(cell.style),
            xlsx_model::CellFormat::default()
        );
    }
}

#[test]
fn explicit_style_selection_and_copy_use_the_selected_xf() {
    let source = package(VARIANTS[0].1);
    for selected in [1, 3] {
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .apply_ops(
                vec![Op::SetCell {
                    sheet: SheetId(0),
                    at: cell("B1"),
                    cell: CellState {
                        value: CellValue::Number { value: 2.0 },
                        style: Some(selected),
                        ..Default::default()
                    },
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        let saved = parts(&workbook.save().unwrap());
        let xml = text(&saved, "xl/styles.xml");
        let selected_xml = cell_xf(&text(&parts(&source), "xl/styles.xml"), selected as usize);
        let written = style_index(&cells(&text(&saved, "xl/worksheets/sheet1.xml"))["B1"]);
        assert_eq!(cell_xf(&xml, written), selected_xml);
    }
    let mut workbook = Workbook::open(&source).unwrap();
    let captured = workbook
        .capture_format(SheetId(0), CellRange::parse_a1("D1").unwrap())
        .unwrap();
    workbook
        .apply_format(
            SheetId(0),
            CellRange::parse_a1("B1").unwrap(),
            captured,
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = parts(&workbook.save().unwrap());
    let written = style_index(&cells(&text(&saved, "xl/worksheets/sheet1.xml"))["B1"]);
    assert_eq!(
        cell_xf(&text(&saved, "xl/styles.xml"), written),
        cell_xf(&text(&parts(&source), "xl/styles.xml"), 3)
    );
}

#[test]
fn unstyled_cells_display_and_report_the_base_xf() {
    let mut source = parts(&package(VARIANTS[0].1));
    let xml = styles(VARIANTS[0].1).replace(
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
        "<xf numFmtId=\"2\" fontId=\"1\" fillId=\"1\" borderId=\"0\" xfId=\"0\"/>",
    );
    source.insert("xl/styles.xml".into(), xml.into_bytes());
    source.insert(
        "xl/worksheets/sheet1.xml".into(),
        worksheet(r#"<sheetData><row r="1"><c r="A1"><v>1.234</v></c></row></sheetData>"#)
            .into_bytes(),
    );
    let source = ooxml_opc::rezip_parts(&source.into_iter().collect::<Vec<_>>()).unwrap();
    let workbook = Workbook::open(&source).unwrap();
    let summary = workbook
        .selection_formatting(SheetId(0), CellRange::parse_a1("A1").unwrap())
        .unwrap();
    assert_eq!(summary.bold, Some(true));
    let matches = workbook.search_text("1.23", true, None);
    assert_eq!(matches.len(), 1);
    assert_eq!(matches[0].address.sheet, SheetId(0));
    assert_eq!(matches[0].address.cell, cell("A1"));
    assert_eq!(matches[0].text, "1.23");
    assert!(workbook.search_text("1.234", true, None).is_empty());
    assert_eq!(
        workbook.model().styles.cell_format(None),
        xlsx_model::CellFormat::default()
    );
    assert_eq!(
        parts(&workbook.save().unwrap())["xl/worksheets/sheet1.xml"],
        parts(&source)["xl/worksheets/sheet1.xml"]
    );
}

#[test]
fn formatting_keeps_each_cells_base_font_and_unmodeled_style_properties() {
    let mut source = parts(&package(VARIANTS[0].1));
    let xml = styles(VARIANTS[0].1).replace("Calibri", "DejaVu Sans")
        .replace("<sz val=\"11\"/>", "<sz val=\"17\"/><family val=\"2\"/><charset val=\"1\"/>")
        .replace("<diagonal/>", "<diagonal style=\"dotted\"><color rgb=\"FF123456\"/></diagonal>")
        .replace("<protection locked=\"0\"/>", "<alignment horizontal=\"right\" textRotation=\"30\" indent=\"2\"/><protection locked=\"0\"/>");
    let xml = xml.replace("<font><b/><sz val=\"17\"/><family val=\"2\"/><charset val=\"1\"/><name val=\"DejaVu Sans\"/>", "<font><b/><sz val=\"21\"/><family val=\"2\"/><charset val=\"1\"/><name val=\"Liberation Serif\"/>");
    source.insert("xl/styles.xml".into(), xml.into_bytes());
    source.insert(
        "xl/worksheets/sheet1.xml".into(),
        worksheet(SHEET1)
            .replace("r=\"A1\" s=\"2\"", "r=\"A1\"")
            .replace("r=\"C1\" s=\"2\"", "r=\"C1\" s=\"3\"")
            .into_bytes(),
    );
    let source = ooxml_opc::rezip_parts(&source.into_iter().collect::<Vec<_>>()).unwrap();
    let mut workbook = Workbook::open(&source).unwrap();
    let before: Vec<_> = ["A1", "B1", "C1"]
        .into_iter()
        .map(|address| {
            workbook.model().styles.cell_format(
                workbook.model().sheets[0]
                    .cell(cell(address))
                    .unwrap()
                    .style
                    .or(Some(0)),
            )
        })
        .collect();
    let op = Op::PatchRangeStyle {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:C1").unwrap(),
        patch: xlsx_ops::StylePatch {
            bold: Some(true),
            fill_color: Some("#D9EAF7".into()),
            ..Default::default()
        },
    };
    workbook
        .apply_ops(vec![op.clone()], CalculationOptions::default())
        .unwrap();
    let pools = workbook.model().styles.pool_marks();
    workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .unwrap();
    assert_eq!(workbook.model().styles.pool_marks(), pools);
    let summary = workbook
        .selection_formatting(SheetId(0), CellRange::parse_a1("A1:C1").unwrap())
        .unwrap();
    assert_eq!(summary.font_family, None);
    assert_eq!(summary.font_size, None);
    let first = workbook
        .selection_formatting(SheetId(0), CellRange::parse_a1("A1").unwrap())
        .unwrap();
    assert_eq!(first.font_family.as_deref(), Some("DejaVu Sans"));
    assert_eq!(first.font_size, Some(17.0));
    let saved = workbook.save().unwrap();
    assert_eq!(workbook.save().unwrap(), saved);
    let reopened = Workbook::open(&saved).unwrap();
    for (address, mut expected) in ["A1", "B1", "C1"].into_iter().zip(before) {
        expected.font.bold = true;
        expected.fill = xlsx_model::Fill::Solid(xlsx_model::Color::Rgb("#d9eaf7".into()));
        let actual = reopened.model().styles.cell_format(
            reopened.model().sheets[0]
                .cell(cell(address))
                .unwrap()
                .style,
        );
        assert_eq!(actual, expected, "{address}");
    }
    let xml = text(&parts(&saved), "xl/styles.xml");
    assert!(!xml.contains("<font></font>"));
    assert!(xml.matches("<family val=\"2\"/>").count() >= 3);
    assert!(xml.matches("textRotation=\"30\"").count() >= 2);
    assert!(xml.matches("<protection locked=\"0\"/>").count() >= 2);
    assert!(xml.contains("<diagonal style=\"dotted\"><color rgb=\"FF123456\"/></diagonal>"));
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
fn precedent_edit_keeps_shared_formulas_and_recalculated_caches() {
    for (name, variant) in VARIANTS {
        let source = package(variant);
        let before = parts(&source);
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(SheetId(1), cell("A2"), "5", CalculationOptions::default())
            .unwrap();
        let after = parts(&workbook.save().unwrap());
        let source_cells = cells(&text(&before, "xl/worksheets/sheet2.xml"));
        let saved_cells = cells(&text(&after, "xl/worksheets/sheet2.xml"));

        for address in ["B1", "B2", "B3"] {
            let expected = if address == "B2" {
                source_cells[address].replace("<v>4</v>", "<v>10</v>")
            } else {
                source_cells[address].clone()
            };
            assert_eq!(saved_cells[address], expected, "{name}: {address}");
            assert!(
                saved_cells[address].contains(r#"t="shared""#),
                "{name}: {address}"
            );
        }
        assert!(saved_cells["A2"].contains("<v>5</v>"), "{name}");
    }
}

#[test]
fn master_precedent_edit_keeps_shared_formulas_and_reopens() {
    for (name, variant) in VARIANTS {
        let source = package(variant);
        let before = parts(&source);
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(SheetId(1), cell("A1"), "5", CalculationOptions::default())
            .unwrap();
        let saved = workbook.save().unwrap();
        let after = parts(&saved);
        let source_cells = cells(&text(&before, "xl/worksheets/sheet2.xml"));
        let saved_cells = cells(&text(&after, "xl/worksheets/sheet2.xml"));
        let reopened = Workbook::open(&saved).unwrap();

        for address in ["B1", "B2", "B3"] {
            let expected = if address == "B1" {
                source_cells[address].replace("<v>2</v>", "<v>10</v>")
            } else {
                source_cells[address].clone()
            };
            assert_eq!(saved_cells[address], expected, "{name}: {address}");
            let current = workbook.model().sheets[1].cell(cell(address)).unwrap();
            let opened = reopened.model().sheets[1].cell(cell(address)).unwrap();
            assert_eq!(opened.formula, current.formula, "{name}: {address}");
            assert_eq!(opened.value, current.value, "{name}: {address}");
        }
    }
}

#[test]
fn shared_cache_type_change_keeps_formula_markup_and_reopens() {
    for (name, variant) in VARIANTS {
        let source = package(variant);
        let before = parts(&source);
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(
                SheetId(1),
                cell("A2"),
                "text",
                CalculationOptions::default(),
            )
            .unwrap();
        let saved = workbook.save().unwrap();
        let after = parts(&saved);
        let source_cells = cells(&text(&before, "xl/worksheets/sheet2.xml"));
        let saved_cells = cells(&text(&after, "xl/worksheets/sheet2.xml"));
        let reopened = Workbook::open(&saved).unwrap();

        for address in ["B1", "B2", "B3"] {
            let expected = if address == "B2" {
                source_cells[address]
                    .replace(r#"s="2""#, r#"s="2" t="e""#)
                    .replace("<v>4</v>", "<v>#VALUE!</v>")
            } else {
                source_cells[address].clone()
            };
            assert_eq!(saved_cells[address], expected, "{name}: {address}");
            let current = workbook.model().sheets[1].cell(cell(address)).unwrap();
            let opened = reopened.model().sheets[1].cell(cell(address)).unwrap();
            assert_eq!(opened.formula, current.formula, "{name}: {address}");
            assert_eq!(opened.value, current.value, "{name}: {address}");
        }
    }
}

#[test]
fn unrelated_edit_keeps_equivalent_style_indices_and_other_parts() {
    for (name, variant) in VARIANTS {
        let source = package(variant);
        let before = parts(&source);
        let mut workbook = Workbook::open(&source).unwrap();
        let source_styles = workbook.model().styles.clone();
        let source_style = workbook.model().sheets[0].cell(cell("A1")).unwrap().style;
        let source_format = source_styles.resolved_format(source_style);
        workbook
            .edit_cell(SheetId(0), cell("A1"), "5", CalculationOptions::default())
            .unwrap();
        let saved = workbook.save().unwrap();
        let after = parts(&saved);
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(
            reopened
                .model()
                .styles
                .resolved_format(reopened.model().sheets[0].cell(cell("A1")).unwrap().style),
            source_format,
            "{name}"
        );

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
        let source_styles = workbook.model().styles.clone();
        let source_style = workbook.model().sheets[0].cell(cell("D1")).unwrap().style;
        let source_format = source_styles.resolved_format(source_style);
        let source_xfs = workbook.model().styles.cell_xfs.len();
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
        let saved = workbook.save().unwrap();
        let after = parts(&saved);
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(
            reopened
                .model()
                .styles
                .resolved_format(reopened.model().sheets[0].cell(cell("B1")).unwrap().style),
            source_format,
            "{name}"
        );
        for index in 0..source_xfs {
            assert_eq!(
                cell_xf(&text(&after, "xl/styles.xml"), index),
                cell_xf(&text(&before, "xl/styles.xml"), index),
                "{name}: cellXfs[{index}]"
            );
        }

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

#[test]
fn refused_grid_regeneration_preserves_source_after_undo_and_unrelated_edit() {
    let second_sheet = concat!(
        r#"<sheetData><row outlineLevel="1"><c r="A1" s="2"><v>1</v></c>"#,
        r#"<c r="B1" s="2"><v>2</v></c></row></sheetData>"#,
    );
    let source = package_with_second_sheet(VARIANTS[0].1, second_sheet);
    let before = parts(&source);
    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .edit_cell(SheetId(1), cell("A1"), "5", CalculationOptions::default())
        .unwrap();

    let Error::Spreadsheet(xlsx_parse::ParseError::UnsupportedEdit(message)) =
        workbook.save().unwrap_err()
    else {
        panic!("expected spreadsheet UnsupportedEdit");
    };
    assert!(message.starts_with("sheet Sheet2:"), "{message}");
    assert!(
        message.contains(r#"<row> attribute outlineLevel="1""#),
        "{message}"
    );

    assert!(
        workbook
            .undo(CalculationOptions::default())
            .unwrap()
            .applied
    );
    let after_undo = parts(&workbook.save().unwrap());
    for path in ["xl/worksheets/sheet2.xml", "xl/styles.xml"] {
        assert_eq!(after_undo[path], before[path], "{path}");
    }

    workbook
        .edit_cell(SheetId(0), cell("A1"), "5", CalculationOptions::default())
        .unwrap();
    let after_edit = parts(&workbook.save().unwrap());
    for path in ["xl/worksheets/sheet2.xml", "xl/styles.xml"] {
        assert_eq!(after_edit[path], before[path], "{path}");
    }
    assert!(cells(&text(&after_edit, "xl/worksheets/sheet1.xml"))["A1"].contains("<v>5</v>"));
}

fn outlined_package() -> Vec<u8> {
    let mut parts = parts(&package(VARIANTS[0].1));
    parts.insert(
        "xl/worksheets/sheet1.xml".to_owned(),
        worksheet(r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#).into_bytes(),
    );
    parts.insert(
        "xl/worksheets/sheet2.xml".to_owned(),
        worksheet(r#"<sheetData><row outlineLevel="1"><c r="A1"><v>1</v></c></row></sheetData>"#)
            .into_bytes(),
    );
    ooxml_opc::rezip_parts(&parts.into_iter().collect::<Vec<_>>()).unwrap()
}

fn edit_and_insert(workbook: &mut Workbook, row: u32) {
    workbook
        .apply_ops(
            vec![
                Op::SetCell {
                    sheet: SheetId(0),
                    at: cell("A1"),
                    cell: CellState {
                        value: CellValue::Number { value: 7.0 },
                        ..CellState::default()
                    },
                },
                Op::InsertRows {
                    sheet: SheetId(1),
                    at: row,
                    count: 1,
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
}

#[test]
fn insert_below_authored_rows_keeps_unrelated_sheet_byte_identical() {
    let source = outlined_package();
    let before = parts(&source);
    let mut workbook = Workbook::open(&source).unwrap();
    let original_sheet = workbook.model().sheets[1].clone();
    edit_and_insert(&mut workbook, 100);
    assert_eq!(workbook.model().sheets[1], original_sheet);

    let saved = workbook.save().unwrap();
    let after = parts(&saved);
    assert_eq!(
        after["xl/worksheets/sheet2.xml"],
        before["xl/worksheets/sheet2.xml"]
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0].cell(cell("A1")).unwrap().value,
        CellValue::Number { value: 7.0 }
    );
    let (oracle, dispatches) = xlsx_parse::with_legacy_save_path(|| workbook.save().unwrap());
    assert!(dispatches > 0);
    assert_eq!(parts(&oracle), after);
}

#[test]
fn insert_above_authored_row_moves_its_unmodeled_attributes() {
    let source = outlined_package();
    let before = parts(&source);
    let mut workbook = Workbook::open(&source).unwrap();
    edit_and_insert(&mut workbook, 0);

    let saved = workbook.save().unwrap();
    let after = parts(&saved);
    assert_ne!(
        after["xl/worksheets/sheet2.xml"],
        before["xl/worksheets/sheet2.xml"]
    );
    let sheet = text(&after, "xl/worksheets/sheet2.xml");
    assert!(
        sheet.contains(r#"<row outlineLevel="1" r="2"><c r="A2"><v>1</v></c></row>"#),
        "{sheet}"
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0].cell(cell("A1")).unwrap().value,
        CellValue::Number { value: 7.0 }
    );
    assert!(reopened.model().sheets[1].cell(cell("A1")).is_none());
    assert_eq!(
        reopened.model().sheets[1].cell(cell("A2")).unwrap().value,
        CellValue::Number { value: 1.0 }
    );
    let (oracle, dispatches) = xlsx_parse::with_legacy_save_path(|| workbook.save().unwrap());
    assert!(dispatches > 0);
    assert_eq!(parts(&oracle), after);
}

#[test]
fn captured_format_preserves_protection_after_peer_style_reordering() {
    let mut source = parts(&package(VARIANTS[0].1));
    source.insert("xl/styles.xml".into(), format!(r#"<styleSheet xmlns="{MAIN}"><fonts count="1"><font/></fonts><cellXfs count="2"><xf fontId="0" applyProtection="1"><protection locked="0"/></xf><xf fontId="0" applyProtection="1"><protection locked="1"/></xf></cellXfs></styleSheet>"#).into_bytes());
    source.insert("xl/worksheets/sheet1.xml".into(), worksheet(r#"<sheetData><row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>"#).into_bytes());
    source.insert(
        "xl/worksheets/sheet2.xml".into(),
        worksheet("<sheetData/>").into_bytes(),
    );
    let source = ooxml_opc::rezip_parts(&source.into_iter().collect::<Vec<_>>()).unwrap();
    let mut a = Workbook::open_collaborative(&source, 11).unwrap();
    let mut b = Workbook::open_collaborative(&source, 12).unwrap();
    a.patch_range_style(
        SheetId(0),
        CellRange::parse_a1("A1").unwrap(),
        xlsx_ops::StylePatch {
            bold: Some(true),
            ..Default::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    let captured = a
        .capture_format(SheetId(0), CellRange::parse_a1("A1").unwrap())
        .unwrap();
    let mut captured: xlsx_ops::CapturedFormat =
        serde_json::from_str(&serde_json::to_string(&captured).unwrap()).unwrap();
    b.apply_update_v1(
        &a.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    b.patch_range_style(
        SheetId(0),
        CellRange::parse_a1("B1").unwrap(),
        xlsx_ops::StylePatch {
            bold: Some(true),
            ..Default::default()
        },
        CalculationOptions::default(),
    )
    .unwrap();
    a.apply_update_v1(
        &b.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    captured.source_styles[0] = a
        .capture_format(SheetId(0), CellRange::parse_a1("B1").unwrap())
        .unwrap()
        .source_styles[0];
    assert_ne!(
        a.capture_format(SheetId(0), CellRange::parse_a1("A1").unwrap())
            .unwrap()
            .source_styles[0],
        captured.source_styles[0]
    );
    a.apply_format(
        SheetId(0),
        CellRange::parse_a1("C1").unwrap(),
        captured,
        CalculationOptions::default(),
    )
    .unwrap();
    b.apply_update_v1(
        &a.encode_state_as_update_v1(),
        CalculationOptions::default(),
    )
    .unwrap();
    for workbook in [&a, &b] {
        let saved = parts(&workbook.save().unwrap());
        let sheet = cells(&text(&saved, "xl/worksheets/sheet1.xml"));
        let styles = text(&saved, "xl/styles.xml");
        let a_xf = cell_xf(&styles, style_index(&sheet["A1"]));
        let b_xf = cell_xf(&styles, style_index(&sheet["B1"]));
        let c_xf = cell_xf(&styles, style_index(&sheet["C1"]));
        assert!(b_xf.contains("locked=\"1\""));
        assert!(c_xf.contains("locked=\"0\""), "C1: {c_xf}");
        assert_eq!(c_xf, a_xf);
    }
}

fn saved_dimension(workbook: &Workbook) -> String {
    let xml = text(
        &parts(&workbook.save().unwrap()),
        "xl/worksheets/sheet2.xml",
    );
    let start = xml.find("<dimension").expect("saved dimension");
    let end = start + xml[start..].find("/>").unwrap() + 2;
    xml[start..end].to_owned()
}

fn dimension_source(dimension: &str, sheet_data: &str) -> Vec<u8> {
    package_with_second_sheet(VARIANTS[0].1, &format!("{dimension}{sheet_data}"))
}

fn edit(workbook: &mut Workbook, row: u32, col: u32, input: &str) {
    workbook
        .edit_cell(
            SheetId(1),
            CellRef::new(row, col),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

fn apply(workbook: &mut Workbook, op: Op) {
    workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .unwrap();
}

#[test]
fn dimension_kept_byte_for_byte_when_the_used_range_is_unchanged() {
    let authored = "<dimension  ref='A1:B3' />";
    let mut workbook = Workbook::open(&dimension_source(authored, SHEET2)).unwrap();
    edit(&mut workbook, 0, 0, "7");
    assert_eq!(saved_dimension(&workbook), authored);
}

#[test]
fn dimension_grows_with_cells_written_outside_it() {
    let mut workbook =
        Workbook::open(&dimension_source(r#"<dimension ref="A1:B3"/>"#, SHEET2)).unwrap();
    edit(&mut workbook, 5, 4, "8");
    assert_eq!(saved_dimension(&workbook), r#"<dimension ref="A1:E6"/>"#);
}

#[test]
fn dimension_keeps_preserved_empty_source_cells() {
    let sheet_data = concat!(
        r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row>"#,
        r#"<row r="100"><c r="Z100"/></row></sheetData>"#,
    );
    let mut workbook = Workbook::open(&dimension_source(
        r#"<dimension ref="A1:Z100"/>"#,
        sheet_data,
    ))
    .unwrap();
    edit(&mut workbook, 1, 1, "2");
    assert_eq!(saved_dimension(&workbook), r#"<dimension ref="A1:Z100"/>"#);
}

#[test]
fn dimension_moves_preserved_empty_source_cells_with_inserted_rows() {
    let sheet_data = concat!(
        r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row>"#,
        r#"<row r="100" ht="20" customHeight="1"><c r="Z100"/></row></sheetData>"#,
    );
    let mut workbook = Workbook::open(&dimension_source(
        r#"<dimension ref="A1:Z100"/>"#,
        sheet_data,
    ))
    .unwrap();
    apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(1),
            at: 1,
            count: 1,
        },
    );
    assert_eq!(saved_dimension(&workbook), r#"<dimension ref="A1:Z101"/>"#);
}

#[test]
fn dimension_follows_inserted_columns_and_deleted_rows() {
    let mut workbook =
        Workbook::open(&dimension_source(r#"<dimension ref="A1:B3"/>"#, SHEET2)).unwrap();
    apply(
        &mut workbook,
        Op::InsertCols {
            sheet: SheetId(1),
            at: 0,
            count: 2,
        },
    );
    assert_eq!(saved_dimension(&workbook), r#"<dimension ref="C1:D3"/>"#);
    apply(
        &mut workbook,
        Op::DeleteRows {
            sheet: SheetId(1),
            at: 2,
            count: 1,
        },
    );
    assert_eq!(saved_dimension(&workbook), r#"<dimension ref="C1:D2"/>"#);
    apply(
        &mut workbook,
        Op::DeleteRows {
            sheet: SheetId(1),
            at: 0,
            count: 2,
        },
    );
    assert_eq!(saved_dimension(&workbook), r#"<dimension ref="A1"/>"#);
}
