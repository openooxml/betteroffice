use betteroffice_xlsx::{
    CalculationOptions, Error, SheetId, SnapshotBudget, Workbook, WorkbookSnapshotBuilder,
    WorkbookSnapshotEncoder,
};
use xlsx_model::CellRange;
use xlsx_ops::Op;

const FILTER: &str = r#"<autoFilter ref="A3:A30"><filterColumn colId="0"><filters><filter val="x"/></filters></filterColumn></autoFilter>"#;

fn source(filter: &str, table: bool) -> Vec<u8> {
    let mut parts = vec![
        ("xl/workbook.xml".into(), br#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data 7" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Data 7'!$A$3:$A$30</definedName></definedNames></workbook>"#.to_vec()),
        ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), format!(r#"<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/>{filter}<mergeCells count="1"><mergeCell ref="H1:I1"/></mergeCells><conditionalFormatting sqref="J1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting>{}</worksheet>"#, if table { r#"<tableParts count="1"><tablePart r:id="t1"/></tableParts>"# } else { "" }).into_bytes()),
    ];
    if filter.is_empty() {
        let xml = String::from_utf8(parts[0].1.clone()).unwrap();
        let start = xml.find("<definedNames>").unwrap();
        let end = xml.find("</definedNames>").unwrap() + "</definedNames>".len();
        parts[0].1 = format!("{}{}", &xml[..start], &xml[end..]).into_bytes();
    }
    if table {
        parts.push(("xl/worksheets/_rels/sheet1.xml.rels".into(), br#"<Relationships><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#.to_vec()));
        parts.push(("xl/tables/table1.xml".into(), br#"<table id="1" name="Items" displayName="Items" ref="E3:F30"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>"#.to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(bytes: &[u8], path: &str) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(p, _)| p == path)
            .unwrap()
            .1,
    )
    .unwrap()
}

fn set(
    book: &mut Workbook,
    range: Option<&str>,
) -> Result<betteroffice_xlsx::MutationResult, betteroffice_xlsx::Error> {
    book.apply_ops(
        vec![Op::SetAutoFilter {
            sheet: SheetId(0),
            range: range.map(|r| CellRange::parse_a1(r).unwrap()),
        }],
        CalculationOptions::default(),
    )
}

#[test]
fn extend_changes_only_reference_and_undo_restores_filter_and_name() {
    let bytes = source(FILTER, false);
    let mut book = Workbook::open(&bytes).unwrap();
    set(&mut book, Some("A3:F30")).unwrap();
    let saved = book.save().unwrap();
    assert_eq!(
        part(&saved, "xl/worksheets/sheet1.xml"),
        part(&bytes, "xl/worksheets/sheet1.xml").replace("A3:A30", "A3:F30")
    );
    assert!(part(&saved, "xl/workbook.xml").contains("&apos;Data 7&apos;!$A$3:$F$30"));
    book.undo(CalculationOptions::default()).unwrap();
    let saved = book.save().unwrap();
    assert_eq!(
        part(&saved, "xl/worksheets/sheet1.xml"),
        part(&bytes, "xl/worksheets/sheet1.xml")
    );
    assert_eq!(
        part(&saved, "xl/workbook.xml"),
        part(&bytes, "xl/workbook.xml")
    );
}

#[test]
fn refusals_are_atomic() {
    for (filter, table, range) in [
        (FILTER.to_owned(), false, "B3:F30"),
        (FILTER.to_owned(), true, "A3:F30"),
        (format!("{FILTER}{FILTER}"), false, "A3:F30"),
    ] {
        let mut book = Workbook::open(&source(&filter, table)).unwrap();
        let before = book.save().unwrap();
        assert!(set(&mut book, Some(range)).is_err());
        assert_eq!(book.save().unwrap(), before);
    }
}

#[test]
fn create_and_remove_use_schema_order() {
    let mut book = Workbook::open(&source("", false)).unwrap();
    set(&mut book, Some("A3:F30")).unwrap();
    let saved = book.save().unwrap();
    let xml = part(&saved, "xl/worksheets/sheet1.xml");
    assert!(xml.find("<autoFilter").unwrap() < xml.find("<mergeCells").unwrap());
    assert!(part(&saved, "xl/workbook.xml").contains("&apos;Data 7&apos;!$A$3:$F$30"));
    set(&mut book, None).unwrap();
    let saved = book.save().unwrap();
    assert!(!part(&saved, "xl/worksheets/sheet1.xml").contains("autoFilter"));
    assert!(!part(&saved, "xl/workbook.xml").contains("_FilterDatabase"));
}

#[test]
fn recovery_preserves_filters_and_undo_with_or_without_conditional_formats() {
    for conditional_formats in [false, true] {
        let mut parts = ooxml_opc::unzip_parts(&source(FILTER, false)).unwrap();
        if !conditional_formats {
            let (_, bytes) = parts
                .iter_mut()
                .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
                .unwrap();
            let xml = String::from_utf8(bytes.clone()).unwrap();
            let start = xml.find("<conditionalFormatting").unwrap();
            let end_tag = "</conditionalFormatting>";
            let end = xml.find(end_tag).unwrap() + end_tag.len();
            *bytes = format!("{}{}", &xml[..start], &xml[end..]).into_bytes();
        }
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let mut book = Workbook::open(&bytes).unwrap();
        let checkpoint = book.recovery_checkpoint(0, None).unwrap();
        let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
        assert_eq!(restored.model(), book.model());
        assert_eq!(
            part(&restored.save().unwrap(), "xl/worksheets/sheet1.xml"),
            part(&bytes, "xl/worksheets/sheet1.xml")
        );
        set(&mut book, Some("A3:F30")).unwrap();
        let checkpoint = book
            .recovery_checkpoint_with_history(1, None, true)
            .unwrap();
        let (mut restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
        assert_eq!(restored.model(), book.model());
        assert_eq!(
            part(&restored.save().unwrap(), "xl/worksheets/sheet1.xml"),
            part(&book.save().unwrap(), "xl/worksheets/sheet1.xml")
        );
        restored.undo(CalculationOptions::default()).unwrap();
        let saved = restored.save().unwrap();
        for path in ["xl/worksheets/sheet1.xml", "xl/workbook.xml"] {
            assert_eq!(part(&saved, path), part(&bytes, path));
        }
    }
}

#[test]
fn snapshot_and_remote_replay_preserve_children() {
    let bytes = source(FILTER, false);
    let book = Workbook::open(&bytes).unwrap();
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&book, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&book).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (restored, _) = builder.finish().unwrap().into_parts();
    assert_eq!(
        restored.model().sheets[0].auto_filter,
        book.model().sheets[0].auto_filter
    );
    assert_eq!(
        part(&restored.save().unwrap(), "xl/worksheets/sheet1.xml"),
        part(&bytes, "xl/worksheets/sheet1.xml")
    );
    let mut left = Workbook::open_collaborative(&bytes, 801).unwrap();
    left.edit_cell(
        SheetId(0),
        xlsx_model::CellRef::new(0, 0),
        "7",
        CalculationOptions::default(),
    )
    .unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 802).unwrap();
    right
        .apply_update_v1(
            &left.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        right.model().sheets[0].auto_filter,
        book.model().sheets[0].auto_filter
    );
}

#[test]
fn open_snapshot_and_remote_replay_preserve_a_source_filter() {
    let parts = ooxml_opc::unzip_parts(&source("", false)).unwrap();
    let parts = parts
        .into_iter()
        .map(|(path, bytes)| {
            let bytes = if path == "xl/worksheets/sheet1.xml" {
                br#"<worksheet><sheetData/><autoFilter ref="A1:C10"/></worksheet>"#.to_vec()
            } else {
                bytes
            };
            (path, bytes)
        })
        .collect::<Vec<_>>();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let book = Workbook::open(&bytes).unwrap();
    let filter = book.model().sheets[0].auto_filter.clone().unwrap();
    assert_eq!(filter.range.to_a1(), "A1:C10");
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&book, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&book).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (restored, _) = builder.finish().unwrap().into_parts();
    assert_eq!(
        restored.model().sheets[0].auto_filter.as_ref(),
        Some(&filter)
    );
    assert_eq!(
        part(&restored.save().unwrap(), "xl/worksheets/sheet1.xml"),
        part(&bytes, "xl/worksheets/sheet1.xml")
    );
    let mut left = Workbook::open_collaborative(&bytes, 811).unwrap();
    let mut right = Workbook::open_collaborative(&bytes, 812).unwrap();
    let state = left.encode_state_as_update_v1();
    assert!(matches!(
        set(&mut left, Some("A1:D10")),
        Err(Error::CollaborativeStructureOperation)
    ));
    assert_eq!(left.model(), book.model());
    assert_eq!(left.encode_state_as_update_v1(), state);
    left.edit_cell(
        SheetId(0),
        xlsx_model::CellRef::new(0, 0),
        "7",
        CalculationOptions::default(),
    )
    .unwrap();
    right
        .apply_update_v1(
            &left.encode_state_as_update_v1(),
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        right.model().sheets[0].auto_filter,
        left.model().sheets[0].auto_filter
    );
    assert_eq!(
        right.model().sheets[0]
            .auto_filter
            .as_ref()
            .unwrap()
            .range
            .to_a1(),
        "A1:C10"
    );
}

#[test]
fn deleting_a_sheet_and_undo_restore_the_source_filter() {
    let bytes = source(FILTER, false);
    let mut book = Workbook::open(&bytes).unwrap();
    book.apply_ops(
        vec![Op::AddSheet {
            index: 1,
            name: "Other".into(),
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let before = book.model().sheets[0].auto_filter.clone();
    book.apply_ops(
        vec![Op::RemoveSheet { index: 0 }],
        CalculationOptions::default(),
    )
    .unwrap();
    book.undo(CalculationOptions::default()).unwrap();
    assert_eq!(book.model().sheets[0].auto_filter, before);
    assert!(part(&book.save().unwrap(), "xl/worksheets/sheet1.xml").contains(FILTER));
    assert_eq!(book.model().defined_names[0].formula, "'Data 7'!$A$3:$A$30");
}

#[test]
fn structural_edits_shift_the_reference_once_and_undo_restores_it() {
    for (op, range) in [
        (
            Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            },
            "A5:A32",
        ),
        (
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 3,
                count: 2,
            },
            "A3:A28",
        ),
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            },
            "C3:C30",
        ),
    ] {
        let bytes = source(FILTER, false);
        let mut book = Workbook::open(&bytes).unwrap();
        book.apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
        assert_eq!(
            book.model().sheets[0]
                .auto_filter
                .as_ref()
                .unwrap()
                .range
                .to_a1(),
            range
        );
        let saved = book.save().unwrap();
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(
            reopened.model().sheets[0]
                .auto_filter
                .as_ref()
                .unwrap()
                .range
                .to_a1(),
            range
        );
        book.undo(CalculationOptions::default()).unwrap();
        assert_eq!(
            part(&book.save().unwrap(), "xl/worksheets/sheet1.xml"),
            part(&bytes, "xl/worksheets/sheet1.xml")
        );
    }
}
