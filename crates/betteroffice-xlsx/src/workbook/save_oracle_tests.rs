use xlsx_model::ErrorValue;
use xlsx_parse::with_legacy_save_path;

use super::edit_tests::{Random, open_identical, options, r, workbook_bytes};
use super::*;

const OP_KINDS: [&str; 22] = [
    "number",
    "shared text",
    "inline text",
    "boolean",
    "error",
    "clear",
    "formula",
    "dependent input",
    "style",
    "lower column",
    "distant cell",
    "insert rows",
    "delete rows",
    "insert columns",
    "delete columns",
    "invalid sheet",
    "invalid row",
    "invalid column",
    "no-op cell",
    "malformed formula",
    "undo",
    "redo",
];

fn dense_workbook_bytes() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&workbook_bytes()).unwrap();
    for sheet in 1..=3 {
        let name = format!("xl/worksheets/sheet{sheet}.xml");
        let (_, bytes) = parts.iter_mut().find(|(part, _)| part == &name).unwrap();
        let mut rows = String::new();
        for row in 6..12 {
            rows.push_str(&format!(r#"<row r="{}">"#, row + 1));
            for col in 0..6 {
                let at = CellRef::new(row, col).to_a1();
                let body = match col {
                    0 => format!(r#"s="1"><v>{}</v>"#, row * 6 + sheet),
                    1 => concat!(
                        r#"t="inlineStr"><is><r><rPr><b/></rPr><t>Rich</t></r>"#,
                        r#"<r><t xml:space="preserve"> inline</t></r></is>"#,
                    )
                    .into(),
                    2 => format!(r#"t="s"><v>{}</v>"#, row % 3),
                    3 => format!(r#"t="b"><v>{}</v>"#, row % 2),
                    4 => r#"t="e"><v>#DIV/0!</v>"#.into(),
                    _ => format!("><f>A1+{row}</f><v>{}</v>", row + 1),
                };
                rows.push_str(&format!(r#"<c r="{at}" {body}</c>"#));
            }
            rows.push_str("</row>");
        }
        rows.push_str(r#"<row r="14"><c r="D14"><v>14</v></c><c r="F14" s="1"/></row>"#);
        let xml = std::str::from_utf8(bytes).unwrap();
        let end = xml.find("</sheetData>").unwrap();
        *bytes = format!("{}{}{}", &xml[..end], rows, &xml[end..]).into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn compare_save(workbook: &Workbook, seed: u32, step: usize) -> Vec<u8> {
    let context = format!(
        "seed {seed:#x}, step {step}, collaborative {}",
        workbook.is_collaborative()
    );
    let actual = workbook
        .save()
        .unwrap_or_else(|error| panic!("normal save: {context}: {error:?}"));
    let (expected, count) = with_legacy_save_path(|| workbook.save());
    let expected = expected.unwrap_or_else(|error| panic!("legacy save: {context}: {error:?}"));
    assert_eq!(actual, expected, "save bytes: {context}");
    if workbook.edited_since_open {
        if workbook.preserved.axes.iter().any(Option::is_some) {
            assert!(count > 0, "legacy dispatch: {context}");
        } else {
            assert_eq!(count, 0, "save without source axes: {context}");
        }
    }
    actual
}

#[test]
fn seeded_workbook_edits_save_identically_through_legacy_path() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Workbook>();

    const STEPS: usize = 220;
    let bytes = dense_workbook_bytes();
    for collaborative in [false, true] {
        for seed in [0x5eed_1234, 0x4d59_5df4] {
            let mut workbook = open_identical(&bytes, collaborative, Some(seed));
            let baseline = compare_save(&workbook, seed, 0);
            let mut previous_save = baseline.clone();
            let mut random = Random(seed);
            let mut counts = [0; OP_KINDS.len()];
            let mut applied = [0; OP_KINDS.len()];
            let mut sheets_edited = [0; 3];
            let mut changed_save = false;
            for step in 0..STEPS {
                let kinds = OP_KINDS.len() - if collaborative { 2 } else { 0 };
                let kind = if collaborative && step >= STEPS - 2 {
                    OP_KINDS.len() - (STEPS - step)
                } else if step < kinds {
                    step
                } else {
                    random.next(kinds as u32) as usize
                };
                counts[kind] += 1;
                let sheet = if matches!(kind, 6 | 7) {
                    SheetId(0)
                } else {
                    SheetId(random.next(3))
                };
                let at = CellRef::new(6 + random.next(6), random.next(6));
                let context = format!(
                    "seed {seed:#x}, step {}, {}, collaborative {collaborative}",
                    step + 1,
                    OP_KINDS[kind]
                );
                let dependents = if matches!(kind, 6 | 7) && step < OP_KINDS.len() {
                    Some([(0, "D1"), (1, "C1"), (2, "B1")].map(|(sheet, at)| {
                        workbook.model().sheets[sheet]
                            .cell(r(at))
                            .unwrap()
                            .value
                            .clone()
                    }))
                } else {
                    None
                };
                let result = match kind {
                    0 => workbook.edit_cell(
                        sheet,
                        at,
                        &(100 + random.next(200)).to_string(),
                        options(),
                    ),
                    1 => {
                        let address = ["A2", "B1", "C1"][sheet.0 as usize];
                        let values = ["Alpha rich", "Beta", "changed shared & <text>"];
                        let value = if step == 1 {
                            values[2]
                        } else {
                            values[random.next(3) as usize]
                        };
                        workbook.edit_cell(sheet, r(address), &format!("'{value}"), options())
                    }
                    2 => workbook.edit_cell(
                        sheet,
                        CellRef::new(at.row, 1),
                        &format!("'inline & <{}> \"quoted\"", random.next(8)),
                        options(),
                    ),
                    3 => {
                        let at = CellRef::new(at.row, 3);
                        let input = match workbook.sheet(sheet).unwrap().cell(at) {
                            Some(cell) if cell.value == (CellValue::Bool { value: true }) => {
                                "false"
                            }
                            _ => "true",
                        };
                        workbook.edit_cell(sheet, at, input, options())
                    }
                    4 => {
                        let mut cell = current_cell_state(workbook.model(), sheet, at);
                        cell.value = CellValue::Error {
                            value: if random.next(2) == 0 {
                                ErrorValue::NA
                            } else {
                                ErrorValue::Ref
                            },
                        };
                        cell.formula = None;
                        workbook.apply_ops(vec![Op::SetCell { sheet, at, cell }], options())
                    }
                    5 => workbook.edit_cell(sheet, at, "", options()),
                    6 => workbook.edit_cell(
                        sheet,
                        r("C1"),
                        &format!("=A1*B1+{}", step + 1),
                        options(),
                    ),
                    7 => workbook.edit_cell(sheet, r("A1"), &(30 + step).to_string(), options()),
                    8 => workbook.patch_range_style(
                        sheet,
                        CellRange::new(at, at),
                        StylePatch {
                            bold: Some(random.next(2) == 0),
                            italic: Some(step.is_multiple_of(2)),
                            ..Default::default()
                        },
                        options(),
                    ),
                    9 => workbook.edit_cell(sheet, r("A14"), &(step + 1).to_string(), options()),
                    10 => {
                        let end = workbook.sheet(sheet).unwrap().used_range().unwrap().end;
                        let at = CellRef::new(end.row + 2, end.col + 2);
                        workbook.edit_cell(sheet, at, "'beyond source", options())
                    }
                    11..=14 => {
                        let at = random.next(if kind < 13 { 10 } else { 5 });
                        let count = 1 + random.next(2);
                        let op = match kind {
                            11 => Op::InsertRows { sheet, at, count },
                            12 => Op::DeleteRows { sheet, at, count },
                            13 => Op::InsertCols { sheet, at, count },
                            _ => Op::DeleteCols { sheet, at, count },
                        };
                        let result = workbook.apply_ops(vec![op], options());
                        if collaborative {
                            assert!(
                                matches!(result, Err(Error::CollaborativeStructureOperation)),
                                "structural rejection: {context}"
                            );
                            Ok(MutationResult::default())
                        } else {
                            result
                        }
                    }
                    15..=17 => {
                        let (sheet, at) = match kind {
                            15 => (SheetId(99), at),
                            16 => (sheet, CellRef::new(MAX_ROWS, at.col)),
                            _ => (sheet, CellRef::new(at.row, MAX_COLS)),
                        };
                        assert!(
                            workbook.edit_cell(sheet, at, "123", options()).is_err(),
                            "invalid edit: {context}"
                        );
                        Ok(MutationResult::default())
                    }
                    18 => {
                        let current = current_cell_state(workbook.model(), sheet, at);
                        let result = if current.formula.is_none()
                            && matches!(current.value, CellValue::Error { .. })
                        {
                            workbook.apply_ops(
                                vec![Op::SetCell {
                                    sheet,
                                    at,
                                    cell: current,
                                }],
                                options(),
                            )
                        } else {
                            let input = workbook.cell(sheet, at).unwrap().input;
                            workbook.edit_cell(sheet, at, &input, options())
                        };
                        assert!(
                            !result
                                .as_ref()
                                .unwrap_or_else(|error| panic!("no-op edit: {context}: {error:?}"))
                                .applied,
                            "no-op edit: {context}"
                        );
                        result
                    }
                    19 => {
                        let result = workbook
                            .edit_cell(sheet, at, "=SUM(", options())
                            .unwrap_or_else(|error| panic!("edit: {context}: {error:?}"));
                        assert_eq!(
                            workbook.sheet(sheet).unwrap().cell(at).unwrap().value,
                            CellValue::Text {
                                value: "=SUM(".into(),
                            },
                            "malformed formula fallback: {context}"
                        );
                        Ok(result)
                    }
                    20 => workbook.undo(options()),
                    _ => workbook.redo(options()),
                };
                let result = result.unwrap_or_else(|error| panic!("edit: {context}: {error:?}"));
                if result.applied {
                    applied[kind] += 1;
                    if kind < 20 {
                        sheets_edited[sheet.0 as usize] += 1;
                    }
                }
                if let Some(before) = dependents {
                    for ((sheet, at), before) in
                        [(0, "D1"), (1, "C1"), (2, "B1")].into_iter().zip(before)
                    {
                        assert_ne!(
                            workbook.model().sheets[sheet].cell(r(at)).unwrap().value,
                            before,
                            "dependent {sheet}!{at}: {context}"
                        );
                    }
                }
                let saved = compare_save(&workbook, seed, step + 1);
                if matches!(kind, 15..=18) || (collaborative && matches!(kind, 11..=14)) {
                    assert_eq!(saved, previous_save, "rejected or no-op save: {context}");
                }
                changed_save |= saved != baseline;
                previous_save = saved;
            }
            assert!(
                changed_save,
                "unchanged run: seed {seed:#x}, collaborative {collaborative}"
            );
            for (kind, count) in counts.into_iter().enumerate() {
                assert!(
                    count > 0,
                    "missing {}: seed {seed:#x}, collaborative {collaborative}",
                    OP_KINDS[kind]
                );
                let needs_application = match kind {
                    11..=14 => !collaborative,
                    15..=18 => false,
                    _ => true,
                };
                if needs_application {
                    assert!(
                        applied[kind] > 0,
                        "unapplied {}: seed {seed:#x}, collaborative {collaborative}",
                        OP_KINDS[kind]
                    );
                }
            }
            assert!(
                sheets_edited.into_iter().all(|count| count > 0),
                "sheet coverage: seed {seed:#x}, collaborative {collaborative}"
            );
        }
    }
}
