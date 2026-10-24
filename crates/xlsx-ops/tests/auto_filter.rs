use xlsx_model::{CellRange, Sheet, SheetId, Workbook};
use xlsx_ops::{Op, Provenance, Transaction, UndoStack, apply};

#[test]
fn null_without_a_filter_is_a_no_op_with_no_inverse() {
    let mut workbook = Workbook::default();
    workbook.sheets.push(Sheet::new("Data 7"));
    let before = workbook.clone();
    let inverse = apply(
        &mut workbook,
        &Op::SetAutoFilter {
            sheet: SheetId(0),
            range: None,
        },
    )
    .unwrap();
    assert!(inverse.0.is_empty());
    assert_eq!(workbook, before);
    let mut stack = UndoStack::new();
    stack
        .commit(
            &mut workbook,
            &Transaction::new(
                vec![Op::SetAutoFilter {
                    sheet: SheetId(0),
                    range: None,
                }],
                Provenance::User,
            ),
        )
        .unwrap();
    assert_eq!(stack.undo_depth(), 0);
}

#[test]
fn create_remove_and_inverse_restore_names_and_filter() {
    let mut workbook = Workbook::default();
    workbook.sheets.push(Sheet::new("Data 7"));
    let before = workbook.clone();
    let inverse = apply(
        &mut workbook,
        &Op::SetAutoFilter {
            sheet: SheetId(0),
            range: Some(CellRange::parse_a1("A3:F30").unwrap()),
        },
    )
    .unwrap();
    assert_eq!(
        workbook.sheets[0]
            .auto_filter
            .as_ref()
            .unwrap()
            .range
            .to_a1(),
        "A3:F30"
    );
    assert_eq!(workbook.defined_names[0].formula, "'Data 7'!$A$3:$F$30");
    assert!(workbook.defined_names[0].hidden);
    let filtered = workbook.clone();
    let removed = apply(
        &mut workbook,
        &Op::SetAutoFilter {
            sheet: SheetId(0),
            range: None,
        },
    )
    .unwrap();
    assert_eq!(workbook, before);
    for op in removed.0 {
        apply(&mut workbook, &op).unwrap();
    }
    assert_eq!(workbook, filtered);
    for op in inverse.0 {
        apply(&mut workbook, &op).unwrap();
    }
    assert_eq!(workbook, before);
}

#[test]
fn unsafe_ranges_and_duplicate_sources_refuse_without_changes() {
    for (range, duplicate, table) in [
        ("B3:F30", false, false),
        ("A3:B30", false, false),
        ("A3:F30", true, false),
        ("A3:F30", false, true),
    ] {
        let mut workbook = Workbook::default();
        let mut sheet = Sheet::new("Data 7");
        sheet.auto_filter = Some(xlsx_model::AutoFilter {
            range: CellRange::parse_a1("A3:F30").unwrap(),
            raw_xml: Some(
                "<autoFilter ref=\"A3:F30\"><filterColumn colId=\"5\"/></autoFilter>".into(),
            ),
            source_count: if duplicate { 2 } else { 1 },
            filter_columns: vec![5],
            has_sort_state: false,
        });
        workbook.sheets.push(sheet);
        if table {
            workbook.tables.push(xlsx_model::Table {
                name: "Items".into(),
                sheet: SheetId(0),
                range: CellRange::parse_a1("E3:F30").unwrap(),
                header_rows: 1,
                totals_rows: 0,
                columns: vec!["A".into(), "B".into()],
            });
        }
        let before = workbook.clone();
        assert!(
            apply(
                &mut workbook,
                &Op::SetAutoFilter {
                    sheet: SheetId(0),
                    range: Some(CellRange::parse_a1(range).unwrap()),
                }
            )
            .is_err()
        );
        assert_eq!(workbook, before);
        if duplicate {
            assert!(
                apply(
                    &mut workbook,
                    &Op::SetAutoFilter {
                        sheet: SheetId(0),
                        range: None
                    }
                )
                .is_err()
            );
            assert_eq!(workbook, before);
        }
    }
}

#[test]
fn wire_range_is_a_string_or_null() {
    let op: Op =
        serde_json::from_str(r#"{"type":"setAutoFilter","sheet":0,"range":"A3:F30"}"#).unwrap();
    assert!(matches!(op, Op::SetAutoFilter { range: Some(_), .. }));
    let op: Op =
        serde_json::from_str(r#"{"type":"setAutoFilter","sheet":0,"range":null}"#).unwrap();
    assert!(matches!(op, Op::SetAutoFilter { range: None, .. }));
    assert_eq!(
        serde_json::to_value(op).unwrap()["range"],
        serde_json::Value::Null
    );
}

fn assert_filter_remap(original: &str, op: Op, expected: &str, columns: &[u32], sorted: bool) {
    let worksheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>{original}</worksheet>"#
    );
    let parts = vec![
        ("xl/workbook.xml".into(), br#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
        ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), worksheet.as_bytes().to_vec()),
    ];
    let parsed = xlsx_parse::parse_workbook_with_package(&parts).unwrap();
    let mut workbook = parsed.workbook.clone();
    let before = workbook.clone();
    let mut stack = UndoStack::new();
    stack
        .commit(&mut workbook, &Transaction::new(vec![op], Provenance::User))
        .unwrap();
    let filter = workbook.sheets[0].auto_filter.as_ref().unwrap();
    assert_eq!(filter.filter_columns, columns);
    assert_eq!(filter.has_sort_state, sorted);
    assert_eq!(filter.raw_xml.as_deref(), Some(expected));
    let save = |workbook: &Workbook| {
        let saved = xlsx_parse::serialize_workbook_with_package_and_origins_after_edits(
            workbook,
            &parsed.package,
            &[Some(0)],
            &[Default::default()],
            xlsx_parse::SaveEdits {
                changed: true,
                moved_references: true,
            },
        )
        .unwrap();
        String::from_utf8(
            saved
                .into_iter()
                .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
                .unwrap()
                .1
                .to_vec(),
        )
        .unwrap()
    };
    assert_eq!(
        save(&workbook),
        format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>{expected}</worksheet>"#
        )
    );
    stack.undo(&mut workbook).unwrap();
    assert_eq!(workbook, before);
    assert_eq!(save(&workbook), worksheet);
}

#[test]
fn column_edits_remap_filter_criteria_and_undo_restores_source_bytes() {
    let original = r#"<autoFilter ref="A1:C10"><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter>"#;
    assert_filter_remap(
        original,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        &original
            .replace("A1:C10", "A1:B10")
            .replace("colId=\"2\"", "colId=\"1\""),
        &[1],
        false,
    );
    assert_filter_remap(
        original,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        &original
            .replace("A1:C10", "A1:D10")
            .replace("colId=\"2\"", "colId=\"3\""),
        &[3],
        false,
    );
    assert_filter_remap(
        original,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        r#"<autoFilter ref="A1:B10"></autoFilter>"#,
        &[],
        false,
    );
}

#[test]
fn filter_column_offsets_follow_the_range_origin_and_edit_count() {
    let original = r#"<autoFilter ref = 'C1:F10'><filterColumn colId = '1'/><filterColumn colId='3' custom='keep'/></autoFilter>"#;
    for (op, expected, columns) in [
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
            original.replace("C1:F10", "E1:H10"),
            vec![1, 3],
        ),
        (
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            },
            original.replace("C1:F10", "A1:D10"),
            vec![1, 3],
        ),
        (
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 3,
            },
            "<autoFilter ref = 'B1:C10'><filterColumn colId='1' custom='keep'/></autoFilter>"
                .into(),
            vec![1],
        ),
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 4,
                count: 2,
            },
            original
                .replace("C1:F10", "C1:H10")
                .replace("colId='3'", "colId='5'"),
            vec![1, 5],
        ),
        (
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 3,
                count: 2,
            },
            "<autoFilter ref = 'C1:D10'><filterColumn colId='1' custom='keep'/></autoFilter>"
                .into(),
            vec![1],
        ),
        (
            Op::InsertCols {
                sheet: SheetId(0),
                at: 6,
                count: 2,
            },
            original.into(),
            vec![1, 3],
        ),
        (
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 6,
                count: 2,
            },
            original.into(),
            vec![1, 3],
        ),
    ] {
        assert_filter_remap(original, op, &expected, &columns, false);
    }
}

#[test]
fn sort_state_references_shift_and_deleted_sort_columns_remove_the_state() {
    let original = r#"<autoFilter ref="A1:C10"><sortState ref="A2:C10" caseSensitive='1'><sortCondition descending="1" ref="C2:C10"/><sortCondition ref='A2:A10'/></sortState><extLst/></autoFilter>"#;
    for (op, expected, sorted) in [
        (Op::InsertCols { sheet: SheetId(0), at: 1, count: 1 }, original.replace("C10", "D10").replace("C2:", "D2:"), true),
        (Op::DeleteCols { sheet: SheetId(0), at: 1, count: 1 }, original.replace("C10", "B10").replace("C2:", "B2:"), true),
        (Op::DeleteCols { sheet: SheetId(0), at: 2, count: 1 }, r#"<autoFilter ref="A1:B10"><extLst/></autoFilter>"#.into(), false),
        (Op::DeleteCols { sheet: SheetId(0), at: 0, count: 1 }, r#"<autoFilter ref="A1:B10"><extLst/></autoFilter>"#.into(), false),
        (Op::InsertRows { sheet: SheetId(0), at: 1, count: 2 }, original.replace("10", "12").replace("2:", "4:"), true),
        (Op::DeleteRows { sheet: SheetId(0), at: 2, count: 2 }, original.replace("10", "8"), true),
        (Op::DeleteRows { sheet: SheetId(0), at: 0, count: 2 }, r#"<autoFilter ref="A1:C8"><sortState ref="A1:C8" caseSensitive='1'><sortCondition descending="1" ref="C1:C8"/><sortCondition ref='A1:A8'/></sortState><extLst/></autoFilter>"#.into(), true),
        (Op::DeleteRows { sheet: SheetId(0), at: 1, count: 9 }, r#"<autoFilter ref="A1:C1"><extLst/></autoFilter>"#.into(), false),
    ] {
        assert_filter_remap(original, op, &expected, &[], sorted);
    }
}

#[test]
fn prefixed_filter_xml_preserves_unedited_attribute_bytes() {
    let original = r#"<s:autoFilter xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main' ref = 'A1:C10'><s:filterColumn custom='keep' colId = '02'><s:filters><s:filter val='x'/></s:filters></s:filterColumn></s:autoFilter>"#;
    assert_filter_remap(
        original,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        &original.replace("A1:C10", "A1:B10").replace("'02'", "'1'"),
        &[1],
        false,
    );
    assert_filter_remap(
        original,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        "<s:autoFilter xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main' ref = 'A1:B10'></s:autoFilter>",
        &[],
        false,
    );
    assert_filter_remap(
        original,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        &original.replace("A1:C10", "A1:C11"),
        &[2],
        false,
    );
    assert_filter_remap(
        original,
        Op::InsertCols {
            sheet: SheetId(0),
            at: 3,
            count: 1,
        },
        original,
        &[2],
        false,
    );
}

#[test]
fn entity_encoded_column_ids_remap_with_the_filter() {
    let original = r#"<autoFilter ref="A1:C10"><filterColumn colId="&#50;"><filters><filter val="x"/></filters></filterColumn></autoFilter>"#;
    assert_filter_remap(
        original,
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        &original
            .replace("A1:C10", "A1:B10")
            .replace("colId=\"&#50;\"", "colId=\"1\""),
        &[1],
        false,
    );
}
