use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};

use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellValue, EditFailureCode, EditOperation, EditRequest,
    EditStep, Error, NumberFormatMutation, ProposalEditInput, ProposalRequest, RangeAddress,
    RangeTarget, SheetId, SnapshotBudget, Workbook, WorkbookSnapshotBuilder,
    WorkbookSnapshotEncoder,
};
use xlsx_model::CellProvider;
use xlsx_ops::{CellState, Op};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const EXT: &str = "urn:test:ext";

const WRAPPED_A1: &str = concat!(
    "<mc:AlternateContent>",
    r#"<mc:Choice Requires="x"><c r="A1" t="inlineStr"><is><t>choice text</t></is></c></mc:Choice>"#,
    r#"<mc:Fallback><c r="A1" t="inlineStr"><is><t>fallback text</t></is></c></mc:Fallback>"#,
    "</mc:AlternateContent>",
);

fn cell_ref(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn sheet2_data() -> String {
    format!(
        concat!(
            r#"<sheetData><row r="1">{wrapped}"#,
            r#"<c r="B1"><f>1+1</f><v>2</v></c>"#,
            r#"<c r="C1"><f>B1+Sheet1!H1</f><v>4</v></c>"#,
            "</row></sheetData>",
        ),
        wrapped = WRAPPED_A1
    )
}

fn package() -> Vec<u8> {
    package_with_wrapped_a1(WRAPPED_A1)
}

fn package_with_wrapped_a1(wrapped_a1: &str) -> Vec<u8> {
    package_with_sheet2(&sheet2_data().replace(WRAPPED_A1, wrapped_a1))
}

fn package_with_sheet2(body: &str) -> Vec<u8> {
    package_with_sheet2_xml(&format!(
        r#"<worksheet xmlns="{MAIN}" xmlns:mc="{MC}" xmlns:x="{EXT}" mc:Ignorable="x">{body}</worksheet>"#,
    ))
}

fn package_with_sheet2_xml(sheet2: &str) -> Vec<u8> {
    let sheet1 = format!(
        concat!(
            r#"<worksheet xmlns="{}"><sheetData><row r="1">"#,
            r#"<c r="A1"><v>10</v></c><c r="H1"><f>1+1</f><v>2</v></c>"#,
            r#"<c r="I1"><f>A1+H1</f><v>12</v></c></row></sheetData></worksheet>"#,
        ),
        MAIN,
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
                r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                "</Types>",
            )
            .to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RELS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
            ),
        ),
        ("xl/worksheets/sheet1.xml", sheet1),
        ("xl/worksheets/sheet2.xml", sheet2.to_owned()),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn wrapper(choice: &str, fallback: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="x">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#,
    )
}

fn in_row(body: &str) -> String {
    format!(r#"<sheetData><row r="1">{body}</row></sheetData>"#)
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn edit(workbook: &mut Workbook, sheet: u32, at: &str, input: &str) {
    assert!(
        workbook
            .edit_cell(
                SheetId(sheet),
                cell_ref(at),
                input,
                CalculationOptions::default()
            )
            .unwrap()
            .applied
    );
}

fn saved(workbook: &Workbook, checks: &[(u32, &str, f64)]) -> String {
    let parts = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
    let reopened = xlsx_parse::parse_workbook(&parts).unwrap();
    for &(sheet, at, expected) in checks {
        assert_eq!(
            reopened.value(SheetId(sheet), cell_ref(at)),
            number(expected),
            "sheet {sheet} cell {at}",
        );
    }
    let xml = parts
        .into_iter()
        .find(|(path, _)| path.trim_start_matches('/') == "xl/worksheets/sheet2.xml")
        .unwrap()
        .1;
    String::from_utf8(xml).unwrap()
}

fn assert_wrapped_a1_preserved(xml: &str) {
    assert_eq!(xml.matches(WRAPPED_A1).count(), 1, "{xml}");
    assert_eq!(xml.matches(r#"r="A1""#).count(), 2, "{xml}");
}

fn assert_refused<T: std::fmt::Debug>(
    workbook: &mut Workbook,
    attempt: impl FnOnce(&mut Workbook) -> betteroffice_xlsx::Result<T>,
) {
    let model = workbook.model().clone();
    let version = workbook.version();
    let history = workbook.history_state();
    let bytes = workbook.save().unwrap();
    let error = attempt(workbook).unwrap_err();
    assert!(
        matches!(
            &error,
            Error::Spreadsheet(xlsx_parse::ParseError::UnsupportedEdit(message))
                if message.contains("Sheet2") && message.contains("mc:AlternateContent")
        ),
        "{error:?}",
    );
    assert_eq!(workbook.model(), &model);
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.history_state(), history);
    assert_eq!(workbook.save().unwrap(), bytes);
}

#[test]
fn edit_sheet_keeps_wrapped_cell_once() {
    let mut workbook = Workbook::open(&package()).unwrap();
    edit(&mut workbook, 1, "B1", "=2+2");
    assert_wrapped_a1_preserved(&saved(&workbook, &[(1, "B1", 4.0), (1, "C1", 6.0)]));
}

#[test]
fn recalc_dirtied_sheet_keeps_wrapped_cell_once() {
    let mut workbook = Workbook::open(&package()).unwrap();
    edit(&mut workbook, 0, "H1", "=2+3");
    assert_wrapped_a1_preserved(&saved(&workbook, &[(0, "H1", 5.0), (1, "C1", 7.0)]));
}

#[test]
fn untouched_sheet_keeps_wrapper_bytes() {
    let mut workbook = Workbook::open(&package()).unwrap();
    edit(&mut workbook, 0, "A1", "20");
    let xml = saved(&workbook, &[(0, "A1", 20.0), (0, "I1", 22.0)]);
    assert!(xml.contains(&sheet2_data()), "{xml}");
}

#[test]
fn edit_of_wrapped_cell_is_refused() {
    let mut workbook = Workbook::open(&package()).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("A1"),
            "updated",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn style_edit_of_wrapped_cell_is_refused() {
    let mut workbook = Workbook::open(&package()).unwrap();
    let at = cell_ref("A1");
    assert_refused(&mut workbook, |workbook| {
        workbook.set_range_number_format(
            SheetId(1),
            CellRange::new(at, at),
            NumberFormatMutation::Percent,
            CalculationOptions::default(),
        )
    });
}

#[test]
fn changing_precedent_of_wrapped_formula_is_refused() {
    let formula = r#"<c r="A1"><f>Sheet1!A1*2</f><v>20</v></c>"#;
    let bytes = package_with_wrapped_a1(&wrapper(formula, formula));
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(0),
            cell_ref("A1"),
            "15",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn edit_batch_preview_and_apply_refuse_wrapped_cache_changes() {
    let formula = r#"<c r="A1"><f>Sheet1!A1*2</f><v>20</v></c>"#;
    let bytes = package_with_wrapped_a1(&wrapper(formula, formula));
    let mut workbook = Workbook::open(&bytes).unwrap();
    let model = workbook.model().clone();
    let baseline = workbook.save().unwrap();
    let version = workbook.version();
    let history = workbook.history_state();
    let request = EditRequest {
        expect_version: version.clone(),
        source: Default::default(),
        history: Default::default(),
        calculation: Default::default(),
        steps: vec![EditStep::new(EditOperation::SetCellInputs {
            target: RangeTarget {
                sheet_id: workbook.sheet_ids()[0].clone(),
                range: RangeAddress::A1 { a1: "A1".into() },
            },
            inputs: vec![vec!["15".into()]],
        })],
    };
    let preview = workbook.validate_edits(&request).unwrap().unwrap_err();
    let applied = workbook.apply_edits(&request).unwrap().unwrap_err();
    assert_eq!(preview, applied);
    assert_eq!(applied.failure.code, EditFailureCode::Unsupported);
    assert!(applied.failure.message.contains("mc:AlternateContent"));
    assert_eq!(workbook.model(), &model);
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.history_state(), history);
    assert_eq!(workbook.save().unwrap(), baseline);
}

#[test]
fn replacing_unwrapped_shared_formula_master_is_refused() {
    let master = r#"<c r="A1"><f t="shared" ref="A1:B1" si="0">1+1</f><v>2</v></c>"#;
    let follower = r#"<c r="B1"><f t="shared" si="0"/><v>2</v></c>"#;
    let body = in_row(&format!("{master}{}", wrapper(follower, follower)));
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("A1"),
            "7",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn clearing_cell_before_implicit_wrapped_cell_is_refused() {
    for fallback in [r#"<c r="B1"><v>2</v></c>"#, "<c><v>2</v></c>"] {
        let body = in_row(&format!(
            r#"<c r="A1"><v>1</v></c>{}"#,
            wrapper("<c><v>2</v></c>", fallback),
        ));
        let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
        assert_refused(&mut workbook, |workbook| {
            workbook.edit_cell(
                SheetId(1),
                cell_ref("A1"),
                "",
                CalculationOptions::default(),
            )
        });
    }
}

#[test]
fn editing_beside_a_wrapper_with_an_unstable_default_namespace_is_refused() {
    let source = format!(
        concat!(
            r#"<s:worksheet xmlns:s="{}" xmlns="{}" xmlns:x="{}">"#,
            r#"<s:sheetData><s:row r="1"><AlternateContent><Choice Requires="x">"#,
            r#"<s:c r="A1"><s:v>1</s:v></s:c></Choice><Fallback><s:c r="A1"><s:v>1</s:v></s:c></Fallback>"#,
            r#"</AlternateContent><s:c r="B1"><s:v>2</s:v></s:c></s:row></s:sheetData></s:worksheet>"#,
        ),
        MAIN, MC, EXT,
    );
    let mut workbook = Workbook::open(&package_with_sheet2_xml(&source)).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("B1"),
            "3",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn renaming_a_sheet_referenced_only_in_choice_is_refused() {
    let body = in_row(&wrapper(
        r#"<c r="A1"><f>Sheet1!A1</f><v>10</v></c>"#,
        r#"<c r="A1"><v>10</v></c>"#,
    ));
    let bytes = package_with_sheet2(&body);
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Inputs".into(),
            }],
            CalculationOptions::default(),
        )
    });
}

#[test]
fn inserting_between_disagreeing_branches_is_refused() {
    let body = in_row(&wrapper(
        r#"<c r="B1"><v>2</v></c>"#,
        r#"<c r="D1"><v>4</v></c>"#,
    ));
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("C1"),
            "3",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn inserting_into_matching_wrapper_cell_or_row_gaps_is_refused() {
    let cells = r#"<c r="B1"><v>2</v></c><c r="D1"><v>4</v></c>"#;
    let rows = r#"<row r="2"><c r="A2"><v>2</v></c></row><row r="4"><c r="A4"><v>4</v></c></row>"#;
    for (body, gap) in [
        (in_row(&wrapper(cells, cells)), "C1"),
        (
            format!("<sheetData>{}</sheetData>", wrapper(rows, rows)),
            "A3",
        ),
    ] {
        let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
        assert_refused(&mut workbook, |workbook| {
            workbook.edit_cell(
                SheetId(1),
                cell_ref(gap),
                "3",
                CalculationOptions::default(),
            )
        });
    }
}

#[test]
fn inserting_beside_an_implicit_wrapped_row_is_refused() {
    let row = r#"<row><c r="A1"><v>1</v></c></row>"#;
    let body = format!("<sheetData>{}</sheetData>", wrapper(row, row));
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("A2"),
            "99",
            CalculationOptions::default(),
        )
    });
}

fn set_cell(sheet: SheetId, at: &str, value: f64) -> Op {
    Op::SetCell {
        sheet,
        at: cell_ref(at),
        cell: CellState {
            value: number(value),
            formula: None,
            style: None,
        },
    }
}

#[test]
fn structural_and_dimension_batches_are_refused_atomically() {
    let sheet = SheetId(1);
    let row = r#"<row r="2"><c r="B2"><v>2</v></c></row>"#;
    let body = format!("<sheetData>{}</sheetData>", wrapper(row, row));
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    let options = CalculationOptions::default();
    for op in [
        Op::InsertRows {
            sheet,
            at: 0,
            count: 1,
        },
        Op::DeleteRows {
            sheet,
            at: 0,
            count: 1,
        },
        Op::InsertCols {
            sheet,
            at: 0,
            count: 1,
        },
        Op::DeleteCols {
            sheet,
            at: 0,
            count: 1,
        },
        Op::SetRowHeight {
            sheet,
            row: 1,
            height: Some(30.0),
        },
        Op::SetRowVisibility {
            sheet,
            row: 1,
            hidden: true,
        },
    ] {
        assert_refused(&mut workbook, |workbook| {
            workbook.apply_ops(vec![set_cell(SheetId(0), "A1", 7.0), op], options)
        });
    }
}

#[test]
fn structural_rewrites_on_a_different_sheet_are_refused() {
    let formula = r#"<c r="A1"><f>Sheet1!A1*2</f><v>20</v></c>"#;
    let mut workbook =
        Workbook::open(&package_with_wrapped_a1(&wrapper(formula, formula))).unwrap();
    for op in [
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Inputs".into(),
        },
        Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        },
    ] {
        assert_refused(&mut workbook, |workbook| {
            workbook.apply_ops(vec![op], CalculationOptions::default())
        });
    }
}

#[cfg(feature = "unstable-track-moves")]
#[test]
fn move_then_edit_batch_uses_moved_source_coordinates() {
    let sheet = SheetId(1);
    let row = r#"<row r="3"><c r="C3"><v>2</v></c></row>"#;
    let body = format!("<sheetData>{}</sheetData>", wrapper(row, row));
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    for (op, at) in [
        (
            Op::MoveRows {
                sheet,
                row: 2,
                count: 1,
                before: 1,
            },
            "C2",
        ),
        (
            Op::MoveCols {
                sheet,
                col: 2,
                count: 1,
                before: 1,
            },
            "B3",
        ),
    ] {
        assert_refused(&mut workbook, |workbook| {
            workbook.apply_ops(
                vec![op, set_cell(sheet, at, 7.0)],
                CalculationOptions::default(),
            )
        });
    }
}

#[cfg(feature = "unstable-track-visibility")]
#[test]
fn visibility_edits_on_wrapped_rows_are_refused() {
    let sheet = SheetId(1);
    let row = r#"<row r="2"><c r="A2"><v>2</v></c></row>"#;
    let body = format!("<sheetData>{}</sheetData>", wrapper(row, row));
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    for op in [
        Op::SetRowVisibility {
            sheet,
            row: 1,
            hidden: true,
        },
        Op::SetRowsVisibility {
            sheet,
            row: 0,
            count: 2,
            hidden: true,
        },
    ] {
        assert_refused(&mut workbook, |workbook| {
            workbook.apply_ops(vec![op], CalculationOptions::default())
        });
    }
}

#[test]
fn foreign_updates_and_snapshots_cannot_edit_wrapped_cells() {
    let ordinary = r#"<c r="A1" t="inlineStr"><is><t>fallback text</t></is></c>"#;
    let mut sender = Workbook::open_collaborative(&package_with_wrapped_a1(ordinary), 1).unwrap();
    let mut receiver = Workbook::open_collaborative(&package(), 2).unwrap();
    edit(&mut sender, 1, "A1", "5");
    let options = CalculationOptions::default();
    let update = sender
        .encode_diff_v1(&receiver.encode_state_vector_v1())
        .unwrap();
    let snapshot = sender.encode_state_as_update_v1();
    for bytes in [&update, &snapshot] {
        let state = receiver.encode_state_as_update_v1();
        assert_refused(&mut receiver, |receiver| {
            receiver.apply_update_v1(bytes, options)
        });
        assert_eq!(receiver.encode_state_as_update_v1(), state);
    }
}

#[test]
fn foreign_update_cannot_dirty_a_wrapped_formula_cache() {
    let formula = r#"<c r="A1"><f>Sheet1!A1*2</f><v>20</v></c>"#;
    let wrapped = package_with_wrapped_a1(&wrapper(formula, formula));
    let mut sender = Workbook::open_collaborative(&package_with_wrapped_a1(formula), 1).unwrap();
    let mut receiver = Workbook::open_collaborative(&wrapped, 2).unwrap();
    edit(&mut sender, 0, "A1", "15");
    let update = sender
        .encode_diff_v1(&receiver.encode_state_vector_v1())
        .unwrap();
    assert_refused(&mut receiver, |receiver| {
        receiver.apply_update_v1(&update, CalculationOptions::default())
    });
}

#[test]
fn collaboration_updates_snapshots_and_history_keep_inert_wrappers_saveable() {
    let bytes = package();
    let mut sender = Workbook::open_collaborative(&bytes, 1).unwrap();
    let mut receiver = Workbook::open_collaborative(&bytes, 2).unwrap();
    let mut restored = Workbook::open_collaborative(&bytes, 3).unwrap();
    let options = CalculationOptions::default();
    edit(&mut sender, 1, "B1", "=2+2");
    edit(&mut sender, 1, "B1", "=3+3");
    let update = sender
        .encode_diff_v1(&receiver.encode_state_vector_v1())
        .unwrap();
    assert!(receiver.apply_update_v1(&update, options).unwrap().applied);
    assert!(
        restored
            .apply_update_v1(&sender.encode_state_as_update_v1(), options)
            .unwrap()
            .applied
    );
    for workbook in [&receiver, &restored] {
        assert_wrapped_a1_preserved(&saved(workbook, &[(1, "B1", 6.0), (1, "C1", 8.0)]));
    }
    assert!(sender.undo(options).unwrap().applied);
    assert_wrapped_a1_preserved(&saved(&sender, &[(1, "B1", 4.0), (1, "C1", 6.0)]));
    assert!(sender.redo(options).unwrap().applied);
    assert_wrapped_a1_preserved(&saved(&sender, &[(1, "B1", 6.0), (1, "C1", 8.0)]));
}

#[test]
fn collaborative_history_refuses_recalculation_that_dirties_a_wrapper() {
    let formula = r#"<c r="A1"><f>NOW()</f><v>42</v></c>"#;
    let bytes = package_with_sheet2(&in_row(&wrapper(formula, formula)));
    let mut workbook = Workbook::open_collaborative(&bytes, 1).unwrap();
    let unchanged_cache = CalculationOptions {
        now_serial: Some(42.0),
    };
    let changed_cache = CalculationOptions {
        now_serial: Some(43.0),
    };
    for input in ["4", "6"] {
        assert!(
            workbook
                .edit_cell(SheetId(0), cell_ref("H1"), input, unchanged_cache)
                .unwrap()
                .applied
        );
    }
    let state = workbook.encode_state_as_update_v1();
    assert_refused(&mut workbook, |workbook| workbook.undo(changed_cache));
    assert_eq!(workbook.encode_state_as_update_v1(), state);
    assert!(workbook.undo(unchanged_cache).unwrap().applied);
    saved(
        &workbook,
        &[(0, "H1", 4.0), (0, "I1", 14.0), (1, "A1", 42.0)],
    );
    let state = workbook.encode_state_as_update_v1();
    assert_refused(&mut workbook, |workbook| workbook.redo(changed_cache));
    assert_eq!(workbook.encode_state_as_update_v1(), state);
    assert!(workbook.redo(unchanged_cache).unwrap().applied);
    saved(
        &workbook,
        &[(0, "H1", 6.0), (0, "I1", 16.0), (1, "A1", 42.0)],
    );
}

#[test]
fn explicit_recalculation_refuses_wrapped_cache_changes() {
    let formula = r#"<c r="A1"><f>NOW()</f><v>42</v></c>"#;
    let wrapped = wrapper(formula, formula);
    let bytes = package_with_sheet2(&in_row(&wrapped));
    let (_, source_sheet) = ooxml_opc::unzip_parts(&bytes)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "xl/worksheets/sheet2.xml")
        .unwrap();
    let mut workbook = Workbook::open(&bytes).unwrap();
    assert_eq!(workbook.save().unwrap(), bytes);
    workbook.recalculate_all(CalculationOptions {
        now_serial: Some(43.0),
    });
    assert_eq!(
        workbook.model().value(SheetId(1), cell_ref("A1")),
        number(43.0)
    );
    let error = workbook.save().unwrap_err();
    assert!(matches!(
        error,
        Error::Spreadsheet(xlsx_parse::ParseError::UnsupportedEdit(message))
            if message.contains("Sheet2") && message.contains("mc:AlternateContent")
    ));
    workbook.recalculate_all(CalculationOptions {
        now_serial: Some(42.0),
    });
    let xml = saved(&workbook, &[(1, "A1", 42.0)]);
    assert_eq!(xml.as_bytes(), source_sheet.as_slice());
}

#[test]
fn undo_restoring_the_last_wrapped_sheet_checks_the_restored_cache() {
    let formula = r#"<c r="A1"><f>NOW()</f><v>42</v></c>"#;
    let wrapped = wrapper(formula, formula);
    let bytes = package_with_sheet2(&in_row(&wrapped));
    let mut workbook = Workbook::open(&bytes).unwrap();
    let unchanged_cache = CalculationOptions {
        now_serial: Some(42.0),
    };
    assert!(
        workbook
            .apply_ops(vec![Op::RemoveSheet { index: 1 }], unchanged_cache)
            .unwrap()
            .applied
    );
    assert_refused(&mut workbook, |workbook| {
        workbook.undo(CalculationOptions {
            now_serial: Some(43.0),
        })
    });
    assert!(workbook.undo(unchanged_cache).unwrap().applied);
    let xml = saved(&workbook, &[(0, "A1", 10.0), (1, "A1", 42.0)]);
    assert!(xml.contains(&wrapped), "{xml}");
}

#[test]
fn collaborative_local_and_remote_edits_use_the_same_wrapped_cache_check() {
    for cached in [3, 4] {
        let formula = format!(r#"<c r="A1"><f>2+2</f><v>{cached}</v></c>"#);
        let wrapped = wrapper(&formula, &formula);
        let bytes = package_with_sheet2(&in_row(&wrapped));
        let mut sender = Workbook::open_collaborative(&bytes, 1).unwrap();
        let mut receiver = Workbook::open_collaborative(&bytes, 2).unwrap();
        let options = CalculationOptions::default();
        if cached == 3 {
            let state = sender.encode_state_as_update_v1();
            let notified = Arc::new(AtomicBool::new(false));
            let listener = Arc::clone(&notified);
            let _subscription = sender
                .observe_update_v1(move |_| listener.store(true, Ordering::Relaxed))
                .unwrap();
            assert_refused(&mut sender, |workbook| {
                workbook.edit_cell(SheetId(0), cell_ref("A1"), "20", options)
            });
            assert_eq!(sender.encode_state_as_update_v1(), state);
            assert!(!notified.load(Ordering::Relaxed));
            continue;
        }
        edit(&mut sender, 0, "A1", "20");
        let update = sender
            .encode_diff_v1(&receiver.encode_state_vector_v1())
            .unwrap();
        assert!(receiver.apply_update_v1(&update, options).unwrap().applied);
        for workbook in [&sender, &receiver] {
            let xml = saved(
                workbook,
                &[(0, "A1", 20.0), (0, "I1", 22.0), (1, "A1", 4.0)],
            );
            assert!(xml.contains(&wrapped), "{xml}");
        }
        assert!(sender.undo(options).unwrap().applied);
        saved(&sender, &[(0, "A1", 10.0), (0, "I1", 12.0), (1, "A1", 4.0)]);
        assert!(sender.redo(options).unwrap().applied);
        saved(&sender, &[(0, "A1", 20.0), (0, "I1", 22.0), (1, "A1", 4.0)]);
    }
}

#[test]
fn self_closing_row_before_an_inert_wrapper_allows_unrelated_edits() {
    let empty = r#"<row r="1" ht="20" customHeight="1"/>"#;
    let row = r#"<row r="2"><c r="A2"><v>2</v></c></row>"#;
    let wrapped = wrapper(row, row);
    let body = format!(
        r#"<sheetData>{empty}{wrapped}<row r="3"><c r="A3"><v>3</v></c></row></sheetData>"#
    );
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    edit(&mut workbook, 1, "A3", "9");
    let xml = saved(&workbook, &[(1, "A2", 2.0), (1, "A3", 9.0)]);
    assert!(xml.contains(&format!("{empty}{wrapped}")), "{xml}");
}

#[test]
fn collaborative_edit_preserves_hidden_wrapped_row_height() {
    let row = r#"<row r="1" ht="25" hidden="1"><c r="A1"><v>1</v></c></row>"#;
    let wrapped = wrapper(row, row);
    let body = format!(
        r#"<sheetData>{wrapped}<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f>A2*2</f><v>4</v></c></row></sheetData>"#,
    );
    let mut workbook = Workbook::open_collaborative(&package_with_sheet2(&body), 1).unwrap();
    assert!(workbook.model().sheets[1].hidden_row_heights.is_empty());
    edit(&mut workbook, 1, "A2", "5");
    assert!(workbook.model().sheets[1].hidden_row_heights.is_empty());
    let xml = saved(&workbook, &[(1, "A2", 5.0), (1, "B2", 10.0)]);
    assert!(xml.contains(&wrapped), "{xml}");
}

fn hydrate(workbook: &Workbook) -> Workbook {
    let budget = SnapshotBudget::new(256, 65_536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(workbook).unwrap() {
        builder.push(&chunk).unwrap();
        builder.advance(budget).unwrap();
    }
    let ready = (0..1_000).any(|_| builder.advance(budget).unwrap().is_ready());
    assert!(ready);
    builder.finish().unwrap().into_parts().0
}

#[test]
fn deferred_sheet_retains_grid_wrappers_and_their_edit_guard() {
    let workbook = Workbook::open(&package()).unwrap();
    let mut workbook = hydrate(&workbook);
    edit(&mut workbook, 1, "B1", "5");
    assert_wrapped_a1_preserved(&saved(&workbook, &[(1, "B1", 5.0), (1, "C1", 7.0)]));
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("A1"),
            "5",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn deferred_filter_edit_uses_the_source_default_style_without_loading_media() {
    let formula = r#"<c r="A1"><f>2+2</f><v>4</v></c>"#;
    let wrapped = wrapper(formula, formula);
    let mut parts = ooxml_opc::unzip_parts(&package_with_wrapped_a1(&wrapped)).unwrap();
    let sheet1 = "xl/worksheets/sheet1.xml";
    let sheet2 = "xl/worksheets/sheet2.xml";
    let source = parts.iter_mut().find(|(name, _)| name == sheet1).unwrap();
    source.1 = std::str::from_utf8(&source.1)
        .unwrap()
        .replace(r#"<c r="A1">"#, r#"<c r="A1" s="0">"#)
        .into_bytes();
    let workbook_part = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/workbook.xml")
        .unwrap();
    workbook_part.1 = std::str::from_utf8(&workbook_part.1)
        .unwrap()
        .replace(
            "</workbook>",
            concat!(
                r#"<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">"#,
                "'Sheet2'!$A$1:$C$2</definedName></definedNames></workbook>",
            ),
        )
        .into_bytes();
    let media = "xl/media/unused.bin";
    parts.push((media.to_owned(), vec![123; 64]));
    xlsx_parse::AlternateContentSources::from_source(&[(1, sheet2)], |paths| {
        assert!(!paths.contains(media));
        Ok(parts
            .iter()
            .filter(|(name, _)| paths.contains(name))
            .cloned()
            .collect())
    })
    .unwrap();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut original = Workbook::open(&bytes).unwrap();
    assert_eq!(original.model().styles.cell_xfs.len(), 1);
    let mut deferred = hydrate(&original);
    let range = CellRange::parse_a1("A1:C2").unwrap();
    for workbook in [&mut original, &mut deferred] {
        let defined_names = workbook.model().defined_names.clone();
        assert!(
            workbook
                .apply_ops(
                    vec![Op::SetAutoFilter {
                        sheet: SheetId(1),
                        range: Some(range),
                    }],
                    CalculationOptions::default(),
                )
                .unwrap()
                .applied
        );
        assert_eq!(workbook.model().defined_names, defined_names);
        let xml = saved(workbook, &[(0, "A1", 10.0), (1, "A1", 4.0), (1, "C1", 4.0)]);
        assert!(xml.contains(&wrapped), "{xml}");
        let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
        let filter = reopened.model().sheets[1].auto_filter.as_ref().unwrap();
        assert_eq!(filter.range, range);
    }
}

#[test]
fn editing_a_deferred_sheet_with_alternate_content_outside_the_grid() {
    let body = format!(
        "{}<extLst><ext uri=\"urn:test\">{}</ext></extLst>",
        sheet2_data().replace(
            WRAPPED_A1,
            r#"<c r="A1" t="inlineStr"><is><t>AlternateContent</t></is></c>"#
        ),
        wrapper("<x:test/>", ""),
    );
    let workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    let mut deferred = hydrate(&workbook);
    edit(&mut deferred, 1, "B1", "5");
    let xml = saved(&deferred, &[(1, "B1", 5.0), (1, "C1", 7.0)]);
    assert!(xml.contains(&wrapper("<x:test/>", "")), "{xml}");
}

#[test]
fn clearing_a_cell_after_a_self_closing_wrapper_is_refused() {
    let body = in_row(r#"<mc:AlternateContent/><c r="B1"><v>2</v></c>"#);
    let mut workbook = Workbook::open(&package_with_sheet2(&body)).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.edit_cell(
            SheetId(1),
            cell_ref("B1"),
            "",
            CalculationOptions::default(),
        )
    });
}

#[test]
fn proposing_a_wrapped_cell_edit_is_refused_before_it_is_recorded() {
    let mut workbook = Workbook::open(&package()).unwrap();
    assert_refused(&mut workbook, |workbook| {
        workbook.propose(
            ProposalRequest {
                agent_id: "test".into(),
                note: None,
                edits: vec![ProposalEditInput {
                    sheet: SheetId(1),
                    cell: cell_ref("A1"),
                    input: "5".into(),
                    number_format: None,
                }],
            },
            CalculationOptions::default(),
        )
    });
    assert!(workbook.proposals().is_empty());
}
