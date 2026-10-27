use betteroffice_xlsx::{CalculationOptions, Error, MutationResult, SheetId, Workbook};
use xlsx_model::{CellRange, CellRef, CellValue};
use xlsx_ops::Op;

fn bytes(styles: Option<&str>) -> Vec<u8> {
    let mut rels = r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#.to_string();
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        ("xl/workbook.xml".into(), br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), br#"<worksheet><sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData></worksheet>"#.to_vec()),
    ];
    if let Some(styles) = styles {
        rels.push_str(r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#);
        parts.push((
            "xl/styles.xml".into(),
            format!("<styleSheet>{styles}</styleSheet>").into_bytes(),
        ));
    }
    parts.push((
        "xl/_rels/workbook.xml.rels".into(),
        format!("<Relationships>{rels}</Relationships>").into_bytes(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn add(kind: &str, operator: Option<&str>, formulas: &[&str]) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:A5").unwrap(),
        kind: kind.into(),
        operator: operator.map(str::to_owned),
        formulas: formulas.iter().map(|f| f.to_string()).collect(),
        format: xlsx_model::DifferentialFormat {
            font_color: Some("#FF0000".into()),
            bold: Some(false),
            italic: None,
            fill: None,
        },
    }
}

#[test]
fn add_list_delete_save_reopen_and_undo_redo() {
    let mut wb = Workbook::open(&bytes(Some(r#"<cellXfs count="1"><xf/></cellXfs>"#))).unwrap();
    wb.apply_ops(
        vec![add("cellIs", Some("greaterThan"), &["5"])],
        CalculationOptions::default(),
    )
    .unwrap();
    wb.apply_ops(
        vec![add("expression", None, &["A1>5"])],
        CalculationOptions::default(),
    )
    .unwrap();
    assert_eq!(
        wb.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        2
    );
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let sheet = String::from_utf8(
        parts
            .iter()
            .find(|(p, _)| p == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert!(sheet.contains("conditionalFormatting"), "{sheet}");
    assert!(sheet.contains(r#"dxfId="0""#), "{sheet}");
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        2
    );
    let id = reopened.model().sheets[0].conditional_formats[0].rules[0]
        .id
        .clone();
    let mut reopened = reopened;
    reopened
        .apply_ops(
            vec![Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        reopened.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        1
    );
    wb.undo(CalculationOptions::default()).unwrap();
    assert_eq!(
        wb.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        1
    );
    wb.redo(CalculationOptions::default()).unwrap();
    assert_eq!(
        wb.model().sheets[0]
            .conditional_formats
            .iter()
            .flat_map(|g| &g.rules)
            .count(),
        2
    );
}

#[test]
fn invalid_batch_is_atomic() {
    let mut wb = Workbook::open(&bytes(None)).unwrap();
    let result = wb.apply_ops(
        vec![
            add("cellIs", Some("greaterThan"), &["5"]),
            add("cellIs", Some("between"), &["1"]),
        ],
        CalculationOptions::default(),
    );
    assert!(result.is_err());
    assert!(wb.model().sheets[0].conditional_formats.is_empty());
}

#[test]
fn collaborative_mode_refuses_conditional_format_ops() {
    let mut wb = Workbook::open_collaborative(&bytes(None), 42).unwrap();
    let result = wb.apply_ops(
        vec![add("cellIs", Some("greaterThan"), &["5"])],
        CalculationOptions::default(),
    );
    assert!(
        matches!(result, Err(Error::CollaborativeStructureOperation)),
        "{result:?}"
    );
}

fn source_cf_bytes() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&bytes(Some(
        r#"<cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs>"#,
    )))
    .unwrap();
    let sheet = parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = br#"<worksheet><sheetData><row r="1"><c r="A1" s="0"><v>10</v></c></row></sheetData><conditionalFormatting sqref="A1:A5" pivot="1" xmlns:a="urn:first"><cfRule type="expression" priority="1" dxfId="0"><formula>A1&gt;0</formula><a:extra/></cfRule></conditionalFormatting><conditionalFormatting sqref="A1:A5" pivot="0" xmlns:a="urn:second"><cfRule type="expression" priority="2" dxfId="0"><formula>A1&gt;5</formula><a:extra/></cfRule></conditionalFormatting></worksheet>"#.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn cell_style_edit_preserves_source_dxfs() {
    let mut wb = Workbook::open(&source_cf_bytes()).unwrap();
    let before = wb.model().sheets[0].conditional_formats.clone();
    wb.apply_ops(
        vec![Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1").unwrap(),
            patch: xlsx_ops::StylePatch {
                bold: Some(true),
                ..Default::default()
            },
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    let reopened = Workbook::open(&wb.save().unwrap()).unwrap();
    assert_eq!(reopened.model().styles.dxfs.len(), 1);
    assert_eq!(reopened.model().sheets[0].conditional_formats, before);
}

#[test]
fn source_cf_wrappers_keep_attributes_and_namespaces() {
    let mut wb = Workbook::open(&source_cf_bytes()).unwrap();
    wb.apply_ops(
        vec![add("expression", None, &["A1>10"])],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(xml.contains(r#"pivot="1" xmlns:a="urn:first""#), "{xml}");
    assert!(xml.contains(r#"pivot="0" xmlns:a="urn:second""#), "{xml}");
    assert_eq!(
        Workbook::open(&saved).unwrap().model().sheets[0]
            .conditional_formats
            .iter()
            .map(|group| group.rules.len())
            .sum::<usize>(),
        3
    );
    wb.apply_ops(
        vec![
            Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id: "cf1".into(),
            },
            Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id: "cf3".into(),
            },
        ],
        CalculationOptions::default(),
    )
    .unwrap();
    let saved = wb.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(!xml.contains("urn:first"), "{xml}");
    assert!(xml.contains(r#"pivot="0" xmlns:a="urn:second""#), "{xml}");
}

#[test]
fn binary_snapshots_accept_source_cf_and_dxfs() {
    for source in [
        source_cf_bytes(),
        bytes(Some(r#"<dxfs count="1"><dxf/></dxfs>"#)),
    ] {
        let wb = Workbook::open(&source).unwrap();
        assert!(
            betteroffice_xlsx::WorkbookSnapshotEncoder::new(
                &wb,
                None,
                betteroffice_xlsx::SnapshotBudget::new(7, 16384).unwrap()
            )
            .is_ok()
        );
        assert!(wb.save().is_ok());
    }
}

fn structural_cf_workbook(formula: &str, imported: bool) -> Workbook {
    let mut wb = Workbook::open(&bytes(None)).unwrap();
    let mut rule = add("expression", None, &[formula]);
    if let Op::AddConditionalFormat { range, .. } = &mut rule {
        *range = CellRange::parse_a1("B3:D5").unwrap();
    }
    wb.apply_ops(
        vec![
            Op::AddSheet {
                index: 1,
                name: "Inputs".into(),
            },
            rule,
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(19, 0),
                cell: xlsx_ops::CellState {
                    value: CellValue::Number { value: 7.0 },
                    ..Default::default()
                },
            },
        ],
        CalculationOptions::default(),
    )
    .unwrap();
    if imported {
        Workbook::open(&wb.save().unwrap()).unwrap()
    } else {
        wb
    }
}

fn assert_cf_refusal(
    wb: &mut Workbook,
    edit: impl FnOnce(&mut Workbook) -> Result<MutationResult, Error>,
) {
    let sheet = &wb.model().sheets[0];
    let group = &sheet.conditional_formats[0];
    let expected = [
        sheet.name.clone(),
        group.sqref.clone(),
        group.rules[0].id.clone(),
    ];
    let message = assert_cf_refusal_unchanged(wb, edit);
    for expected in expected {
        assert!(message.contains(expected.as_str()), "{message}");
    }
}

fn assert_cf_refusal_unchanged(
    wb: &mut Workbook,
    edit: impl FnOnce(&mut Workbook) -> Result<MutationResult, Error>,
) -> String {
    let before = wb.model().clone();
    let version = wb.version();
    let history = wb.history_state();
    let saved = wb.save().unwrap();
    let error = edit(wb).unwrap_err();
    assert!(
        matches!(
            &error,
            Error::Operation(xlsx_ops::OpError::InvalidConditionalFormat(_))
        ),
        "{error:?}"
    );
    let message = error.to_string();
    assert_eq!(wb.model(), &before);
    assert_eq!(wb.version(), version);
    assert_eq!(wb.history_state(), history);
    assert_eq!(wb.save().unwrap(), saved);
    message
}

fn assert_cf_op_refusal(formula: &str, op: Op) {
    let mut wb = structural_cf_workbook(formula, false);
    assert_cf_refusal(&mut wb, |wb| {
        wb.apply_ops(
            vec![
                Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(20, 0),
                    cell: xlsx_ops::CellState {
                        value: CellValue::Number { value: 42.0 },
                        ..Default::default()
                    },
                },
                op,
            ],
            CalculationOptions::default(),
        )
    });
}

#[test]
fn conditional_format_refuses_rename_referenced_sheet_atomically() {
    for (formula, sheet) in [
        ("'Inputs'!$H$12>0", SheetId(1)),
        ("Sheet1!$A$1>0", SheetId(0)),
    ] {
        assert_cf_op_refusal(
            formula,
            Op::RenameSheet {
                sheet,
                name: "Renamed".into(),
            },
        );
    }
}

#[test]
fn conditional_format_refuses_move_intersecting_sqref_atomically() {
    for (source, destination) in [("B4", "F10"), ("F10", "C4")] {
        let mut wb = structural_cf_workbook("1", false);
        assert_cf_refusal(&mut wb, |wb| {
            wb.move_range(
                SheetId(0),
                CellRange::parse_a1(source).unwrap(),
                CellRef::parse_a1(destination).unwrap(),
                CalculationOptions::default(),
            )
        });
    }
}

#[test]
fn conditional_format_refuses_move_of_referenced_cell_atomically() {
    for (formula, sheet, source) in [
        ("$H$12>0", SheetId(0), "H12"),
        ("Inputs!$H$12>0", SheetId(1), "H12"),
        ("H10>0", SheetId(0), "J12"),
        ("SUM($H$10:$H$12)>0", SheetId(0), "H11"),
    ] {
        let mut wb = structural_cf_workbook(formula, false);
        assert_cf_refusal(&mut wb, |wb| {
            wb.move_range(
                sheet,
                CellRange::parse_a1(source).unwrap(),
                CellRef::new(14, 10),
                CalculationOptions::default(),
            )
        });
    }
}

#[test]
fn conditional_format_refuses_row_delete_inside_rule_atomically() {
    assert_cf_op_refusal(
        "1",
        Op::DeleteRows {
            sheet: SheetId(0),
            at: 3,
            count: 1,
        },
    );
}

#[test]
fn conditional_format_refuses_row_insert_above_rule_atomically() {
    assert_cf_op_refusal(
        "1",
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
    );
}

#[test]
fn conditional_format_refuses_column_delete_atomically() {
    assert_cf_op_refusal(
        "1",
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
    );
}

#[test]
fn conditional_format_refuses_grid_edit_of_formula_reference_atomically() {
    for op in [
        Op::InsertRows {
            sheet: SheetId(1),
            at: 11,
            count: 1,
        },
        Op::DeleteRows {
            sheet: SheetId(1),
            at: 11,
            count: 1,
        },
        Op::InsertCols {
            sheet: SheetId(1),
            at: 7,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(1),
            at: 7,
            count: 1,
        },
    ] {
        assert_cf_op_refusal("Inputs!$H$12>0", op);
    }
    assert_cf_op_refusal(
        "SUM($B$1:$B$1048576)>0",
        Op::InsertRows {
            sheet: SheetId(0),
            at: 100,
            count: 1,
        },
    );
}

#[test]
fn conditional_format_refuses_sheet_span_edits_atomically() {
    for formula in [
        "SUM('Jan:Mar'!$H$12)>0",
        "SUM(Jan:Mar!$H$12)>0",
        "SUM('Jan:Jan'!$H$12)>0",
    ] {
        let mut wb = structural_cf_workbook(formula, false);
        wb.apply_ops(
            vec![
                Op::AddSheet {
                    index: 2,
                    name: "Jan".into(),
                },
                Op::AddSheet {
                    index: 3,
                    name: "Feb".into(),
                },
                Op::AddSheet {
                    index: 4,
                    name: "Mar".into(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
        for op in [
            Op::InsertRows {
                sheet: SheetId(3),
                at: 0,
                count: 1,
            },
            Op::InsertRows {
                sheet: SheetId(1),
                at: 99,
                count: 1,
            },
            Op::RenameSheet {
                sheet: SheetId(3),
                name: "Renamed".into(),
            },
            Op::RenameSheet {
                sheet: SheetId(1),
                name: "Renamed".into(),
            },
        ] {
            assert_cf_refusal(&mut wb, |wb| {
                wb.apply_ops(vec![op], CalculationOptions::default())
            });
        }
    }
}

#[test]
fn conditional_format_refuses_mixed_anchor_range_moves_atomically() {
    for (formula, sqref, source) in [
        ("SUM(A10:$A$1)>0", "B1:B20", "A25"),
        ("SUM($A$10:A1)>0", "B1:B20", "A25"),
        ("SUM(J1:$A$1)>0", "A2:T2", "Y1"),
    ] {
        let mut wb = Workbook::open(&bytes(None)).unwrap();
        let mut rule = add("expression", None, &[formula]);
        if let Op::AddConditionalFormat { range, .. } = &mut rule {
            *range = CellRange::parse_a1(sqref).unwrap();
        }
        wb.apply_ops(vec![rule], CalculationOptions::default())
            .unwrap();
        assert_cf_refusal(&mut wb, |wb| {
            wb.move_range(
                SheetId(0),
                CellRange::parse_a1(source).unwrap(),
                CellRef::parse_a1("Z30").unwrap(),
                CalculationOptions::default(),
            )
        });
    }
}

#[test]
fn source_conditional_formats_keep_their_geometry_on_structural_edits() {
    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
        Op::InsertRows {
            sheet: SheetId(1),
            at: 11,
            count: 1,
        },
        Op::RenameSheet {
            sheet: SheetId(1),
            name: "Renamed".into(),
        },
    ] {
        let mut wb = structural_cf_workbook("Inputs!$H$12>0", true);
        let source = wb.save().unwrap();
        wb.apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
        assert_eq!(
            conditional_format_xml(&wb.save().unwrap()),
            conditional_format_xml(&source)
        );
    }
    let mut wb = structural_cf_workbook("1", true);
    let error = wb
        .move_range(
            SheetId(0),
            CellRange::parse_a1("B4").unwrap(),
            CellRef::parse_a1("F10").unwrap(),
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(matches!(error, Error::InvalidOperation(_)), "{error:?}");
}

#[test]
fn conditional_format_keeps_indirect_text_when_renaming_sheets() {
    for formula in [
        r#"INDIRECT("Inputs!$H$12")>0"#,
        r#"INDIRECT("Jan:Mar!$H$12")>0"#,
    ] {
        for imported in [false, true] {
            let mut wb = structural_cf_workbook(formula, imported);
            let before = wb.model().sheets[0].conditional_formats.clone();
            let saved = Workbook::open(&wb.save().unwrap()).unwrap().model().sheets[0]
                .conditional_formats
                .clone();
            wb.apply_ops(
                vec![Op::RenameSheet {
                    sheet: SheetId(1),
                    name: "Renamed".into(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
            assert_eq!(wb.model().sheets[0].conditional_formats, before);
            assert_eq!(
                Workbook::open(&wb.save().unwrap()).unwrap().model().sheets[0].conditional_formats,
                saved
            );
        }
    }
}

fn conditional_format_xml(bytes: &[u8]) -> Vec<String> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    xml.match_indices("<conditionalFormatting")
        .map(|(start, _)| {
            let end = xml[start..].find("</conditionalFormatting>").unwrap()
                + "</conditionalFormatting>".len();
            xml[start..start + end].to_string()
        })
        .collect()
}

#[test]
fn conditional_format_allows_insert_below_rules_byte_identically() {
    for source in [
        structural_cf_workbook("$H$12>0", false).save().unwrap(),
        source_cf_bytes(),
    ] {
        let mut wb = Workbook::open(&source).unwrap();
        wb.apply_ops(
            vec![Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(19, 0),
                cell: xlsx_ops::CellState {
                    value: CellValue::Number { value: 7.0 },
                    ..Default::default()
                },
            }],
            CalculationOptions::default(),
        )
        .unwrap();
        let before = wb.model().sheets[0].conditional_formats.clone();
        let result = wb
            .apply_ops(
                vec![Op::InsertRows {
                    sheet: SheetId(0),
                    at: 12,
                    count: 1,
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(result.applied);
        assert!(wb.model().sheets[0].cell(CellRef::new(19, 0)).is_none());
        assert_eq!(
            wb.model().sheets[0]
                .cell(CellRef::new(20, 0))
                .unwrap()
                .value,
            CellValue::Number { value: 7.0 }
        );
        let saved = wb.save().unwrap();
        assert_eq!(
            conditional_format_xml(&saved),
            conditional_format_xml(&source)
        );
        assert_eq!(
            Workbook::open(&saved).unwrap().model().sheets[0].conditional_formats,
            before
        );
    }
    let mut wb = structural_cf_workbook("$H$12>0", false);
    let before = wb.model().sheets[0].conditional_formats.clone();
    wb.apply_ops(
        vec![Op::InsertRows {
            sheet: SheetId(0),
            at: 12,
            count: 1,
        }],
        CalculationOptions::default(),
    )
    .unwrap();
    assert_eq!(wb.model().sheets[0].conditional_formats, before);
}

#[test]
fn conditional_format_allows_unaffected_rename_and_move() {
    for imported in [false, true] {
        let mut wb = structural_cf_workbook("$H$12>0", imported);
        let before = wb.model().sheets[0].conditional_formats.clone();
        wb.apply_ops(
            vec![
                Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "Renamed".into(),
                },
                Op::RenameSheet {
                    sheet: SheetId(1),
                    name: "Other".into(),
                },
            ],
            CalculationOptions::default(),
        )
        .unwrap();
        if !imported {
            wb.move_range(
                SheetId(0),
                CellRange::parse_a1("A1").unwrap(),
                CellRef::new(19, 0),
                CalculationOptions::default(),
            )
            .unwrap();
        }
        assert_eq!(wb.model().sheets[0].conditional_formats, before);
        assert_eq!(
            Workbook::open(&wb.save().unwrap()).unwrap().model().sheets[0]
                .conditional_formats
                .len(),
            1
        );
    }
}
