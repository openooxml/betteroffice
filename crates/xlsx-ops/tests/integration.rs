//! public-api integration: serde wire round-trips and a full edit/undo cycle.

use xlsx_model::{CellProvider, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook};
use xlsx_ops::{
    BorderLineStyle, BorderPatch, BorderPreset, CellState, HorizontalAlignment,
    NumberFormatMutation, Op, Provenance, StylePatch, StyleProperty, TextWrapping, Transaction,
    UndoStack, apply,
};

fn r(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

#[test]
fn pattern_border_format_edit_changes_only_requested_facet() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    let mut format = wb.styles.cell_format(None);
    format.fill = serde_json::from_value(serde_json::json!({"Pattern": {
        "pattern": "lightUp", "fg": {"Rgb": "#123456"}, "bg": {"Rgb": "#fedcba"}
    }}))
    .unwrap();
    format.border.top = Some(xlsx_model::BorderEdge {
        style: xlsx_model::BorderStyle::from_sml("mediumDashed"),
        color: Some(xlsx_model::Color::Indexed(5)),
    });
    let style = wb.styles.intern_cell_format(&format).unwrap();
    wb.sheets[0].set_cell(
        r("A1"),
        xlsx_model::Cell {
            style,
            ..Default::default()
        },
    );
    let mut stack = UndoStack::new();
    for patch in [
        StylePatch {
            bold: Some(true),
            ..Default::default()
        },
        StylePatch {
            border: Some(BorderPatch {
                style: Some(BorderLineStyle::Dotted),
                ..Default::default()
            }),
            ..Default::default()
        },
        StylePatch {
            fill_color: Some("#abcdef".into()),
            ..Default::default()
        },
    ] {
        stack
            .commit(
                &mut wb,
                &Transaction::new(
                    vec![Op::PatchRangeStyle {
                        sheet: SheetId(0),
                        range: CellRange::parse_a1("A1").unwrap(),
                        patch: patch.clone(),
                    }],
                    Provenance::User,
                ),
            )
            .unwrap();
        let actual = wb
            .styles
            .cell_format(wb.sheets[0].cell(r("A1")).unwrap().style);
        if patch.bold.is_some() {
            assert_eq!(actual.fill, format.fill);
            assert_eq!(actual.border, format.border);
        } else if patch.border.is_some() {
            assert_eq!(actual.fill, format.fill);
            assert_eq!(
                actual.border.top.as_ref().unwrap().style,
                xlsx_model::BorderStyle::Dotted
            );
        } else {
            assert_eq!(
                actual.fill,
                xlsx_model::Fill::Solid(xlsx_model::Color::Rgb("#abcdef".into()))
            );
            assert_eq!(
                actual.border.top.as_ref().unwrap().style,
                xlsx_model::BorderStyle::Dotted
            );
        }
    }
}

#[test]
fn transaction_json_round_trip() {
    let tx = Transaction::proposal(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: r("B2"),
                cell: CellState {
                    value: CellValue::Text { value: "hi".into() },
                    formula: Some("A1&\"i\"".into()),
                    style: Some(3),
                },
            },
            Op::MergeCells {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:C1").unwrap(),
            },
            Op::InsertRows {
                sheet: SheetId(0),
                at: 4,
                count: 2,
            },
        ],
        Provenance::Agent {
            id: "agent-7".into(),
        },
    );

    let json = serde_json::to_string(&tx).unwrap();
    let back: Transaction = serde_json::from_str(&json).unwrap();
    assert_eq!(tx, back);

    let value: serde_json::Value = serde_json::from_str(&json).unwrap();
    assert_eq!(value["proposed"], true);
    assert_eq!(value["author"]["kind"], "agent");
    assert_eq!(value["author"]["id"], "agent-7");
    assert_eq!(value["ops"][0]["type"], "setCell");
    assert_eq!(value["ops"][2]["type"], "insertRows");
}

#[test]
fn full_workbook_round_trip_via_undo_stack() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    wb.sheet_mut(SheetId(0)).unwrap().set_cell(
        r("A1"),
        xlsx_model::Cell {
            value: CellValue::Number { value: 1.0 },
            ..Default::default()
        },
    );
    let baseline = wb.clone();

    let tx = Transaction::new(
        vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: r("A2"),
                cell: CellState {
                    value: CellValue::Number { value: 2.0 },
                    ..Default::default()
                },
            },
            Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::MergeCells {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B1").unwrap(),
            },
        ],
        Provenance::User,
    );

    let mut stack = UndoStack::new();
    stack.commit(&mut wb, &tx).unwrap();
    assert_ne!(wb.sheets[0].used_range(), baseline.sheets[0].used_range());

    stack.undo(&mut wb).unwrap();
    assert_eq!(wb.sheets[0].used_range(), baseline.sheets[0].used_range());
    assert_eq!(
        wb.value(SheetId(0), r("A1")),
        CellValue::Number { value: 1.0 }
    );
    assert!(wb.sheets[0].merges.is_empty());
}

#[test]
fn merge_replaces_intersections_and_undo_restores_them() {
    let cases = [
        (
            "all",
            vec!["G1:H2", "A1:B2", "D2:E3"],
            vec!["B2:E4"],
            vec!["G1:H2", "B2:E4"],
        ),
        (
            "horizontal",
            vec!["F1:G2", "A1:B2"],
            vec!["A1:D1", "A2:D2"],
            vec!["F1:G2", "A1:D1", "A2:D2"],
        ),
        (
            "vertical",
            vec!["D1:E2", "A1:B2"],
            vec!["A1:A4", "B1:B4"],
            vec!["D1:E2", "A1:A4", "B1:B4"],
        ),
    ];

    for (name, initial, replacements, expected) in cases {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb.sheets[0].merges = initial
            .iter()
            .map(|range| CellRange::parse_a1(range).unwrap())
            .collect();
        let before = wb.sheets[0].merges.clone();
        let ops = replacements
            .iter()
            .map(|range| Op::MergeCells {
                sheet: SheetId(0),
                range: CellRange::parse_a1(range).unwrap(),
            })
            .collect();
        let tx = Transaction::new(ops, Provenance::User);
        let mut stack = UndoStack::new();

        stack.commit(&mut wb, &tx).unwrap();

        let expected = expected
            .iter()
            .map(|range| CellRange::parse_a1(range).unwrap())
            .collect::<Vec<_>>();
        assert_eq!(wb.sheets[0].merges, expected, "{name}");
        for (index, merged) in wb.sheets[0].merges.iter().enumerate() {
            assert!(
                wb.sheets[0].merges[index + 1..]
                    .iter()
                    .all(|other| merged.end.row < other.start.row
                        || other.end.row < merged.start.row
                        || merged.end.col < other.start.col
                        || other.end.col < merged.start.col),
                "{name}"
            );
        }

        stack.undo(&mut wb).unwrap();
        assert_eq!(wb.sheets[0].merges, before, "{name}");
    }
}

#[test]
fn apply_returns_replayable_inverse() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));

    let op = Op::SetCell {
        sheet: SheetId(0),
        at: r("D4"),
        cell: CellState {
            value: CellValue::Bool { value: true },
            ..Default::default()
        },
    };
    let inverse = apply(&mut wb, &op).unwrap();
    assert_eq!(
        wb.value(SheetId(0), r("D4")),
        CellValue::Bool { value: true }
    );
    for iop in &inverse.0 {
        apply(&mut wb, iop).unwrap();
    }
    assert_eq!(wb.value(SheetId(0), r("D4")), CellValue::Empty);
}

#[test]
fn range_style_and_number_format_are_undoable() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    let tx = Transaction::new(
        vec![
            Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                patch: StylePatch {
                    bold: Some(true),
                    text_color: Some("#123456".into()),
                    ..StylePatch::default()
                },
            },
            Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                format: NumberFormatMutation::Percent,
            },
            Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                format: NumberFormatMutation::IncreaseDecimal,
            },
            Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                patch: StylePatch {
                    fill_color: Some("#abcdef".into()),
                    horizontal_alignment: Some(HorizontalAlignment::Center),
                    text_wrapping: Some(TextWrapping::Wrap),
                    border: Some(BorderPatch {
                        preset: Some(BorderPreset::Outer),
                        style: Some(BorderLineStyle::Double),
                        color: Some("#654321".into()),
                    }),
                    ..StylePatch::default()
                },
            },
        ],
        Provenance::User,
    );
    let mut stack = UndoStack::new();
    stack.commit(&mut wb, &tx).unwrap();
    for address in ["A1", "B1", "A2", "B2"] {
        let cell = wb.sheets[0].cell(r(address)).unwrap();
        let format = wb.styles.cell_format(cell.style);
        assert!(format.font.bold);
        assert_eq!(
            format.number_format,
            xlsx_model::NumberFormat::Custom {
                pattern: "0.000%".into()
            }
        );
        assert_eq!(format.alignment.h, Some(xlsx_model::HAlign::Center));
        assert!(format.alignment.wrap_text);
        assert_eq!(
            format.fill,
            xlsx_model::Fill::Solid(xlsx_model::Color::Rgb("#abcdef".into()))
        );
    }
    let top_left = wb
        .styles
        .cell_format(wb.sheets[0].cell(r("A1")).unwrap().style);
    assert_eq!(
        top_left.border.top.as_ref().unwrap().style,
        xlsx_model::BorderStyle::Double
    );
    assert!(top_left.border.left.is_some());
    assert!(top_left.border.right.is_none());
    assert!(top_left.border.bottom.is_none());
    stack.undo(&mut wb).unwrap();
    assert!(wb.sheets[0].iter_cells().next().is_none());
    stack.redo(&mut wb).unwrap();
    assert!(
        wb.styles
            .cell_format(wb.sheets[0].cell(r("A1")).unwrap().style)
            .font
            .bold
    );
}

#[test]
fn clearing_a_style_only_cell_restores_sparse_storage() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    let range = CellRange::parse_a1("A1").unwrap();
    apply(
        &mut wb,
        &Op::PatchRangeStyle {
            sheet: SheetId(0),
            range,
            patch: StylePatch {
                italic: Some(true),
                ..StylePatch::default()
            },
        },
    )
    .unwrap();
    assert!(wb.sheets[0].cell(r("A1")).is_some());

    apply(
        &mut wb,
        &Op::PatchRangeStyle {
            sheet: SheetId(0),
            range,
            patch: StylePatch {
                clear: vec![StyleProperty::Italic],
                ..StylePatch::default()
            },
        },
    )
    .unwrap();
    assert!(wb.sheets[0].cell(r("A1")).is_none());
}

fn clear_italic(range: CellRange) -> Op {
    Op::PatchRangeStyle {
        sheet: SheetId(0),
        range,
        patch: StylePatch {
            clear: vec![StyleProperty::Italic],
            ..StylePatch::default()
        },
    }
}

#[test]
fn clearing_a_whole_unstyled_column_leaves_no_stored_cells() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    apply(
        &mut wb,
        &Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A3").unwrap(),
            patch: StylePatch {
                italic: Some(true),
                ..StylePatch::default()
            },
        },
    )
    .unwrap();
    assert_eq!(wb.sheets[0].iter_cells().count(), 3);

    let column = CellRange::new(r("A1"), CellRef::new(xlsx_model::MAX_ROWS - 1, 0));
    apply(&mut wb, &clear_italic(column)).unwrap();
    assert_eq!(wb.sheets[0].iter_cells().count(), 0);
}

#[test]
fn clearing_a_cell_under_a_column_style_keeps_an_explicit_default() {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    sheet.col_styles.push(xlsx_model::ColStyle {
        first: 0,
        last: 0,
        xf: 1,
    });
    wb.sheets.push(sheet);
    wb.styles.fonts = vec![
        xlsx_model::Font::default(),
        xlsx_model::Font {
            italic: true,
            ..Default::default()
        },
    ];
    wb.styles.cell_xfs = vec![
        xlsx_model::Xf::default(),
        xlsx_model::Xf {
            font: Some(1),
            ..Default::default()
        },
    ];
    apply(
        &mut wb,
        &clear_italic(CellRange::parse_a1("A1:B2").unwrap()),
    )
    .unwrap();
    for at in ["A1", "A2"] {
        assert_eq!(wb.sheets[0].cell(r(at)).unwrap().style, Some(0));
    }
    assert_eq!(wb.sheets[0].iter_cells().count(), 2);
}

fn italic_default_workbook() -> Workbook {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        r("A1"),
        xlsx_model::Cell {
            value: CellValue::Number { value: 1.0 },
            style: Some(0),
            ..Default::default()
        },
    );
    sheet.row_styles.insert(1, 1);
    wb.sheets.push(sheet);
    wb.styles.fonts = vec![xlsx_model::Font {
        italic: true,
        ..Default::default()
    }];
    let italic = xlsx_model::Xf {
        font: Some(0),
        ..Default::default()
    };
    wb.styles.cell_xfs = vec![italic.clone(), italic];
    wb
}

fn renders_italic(wb: &Workbook, at: &str) -> bool {
    let style = wb.sheets[0].effective_style(r(at)).unwrap_or(0);
    wb.styles.cell_format(Some(style)).font.italic
}

#[test]
fn clearing_italic_under_an_italic_default_xf_stores_a_plain_style() {
    let mut wb = italic_default_workbook();
    apply(
        &mut wb,
        &clear_italic(CellRange::parse_a1("A1:A2").unwrap()),
    )
    .unwrap();
    for at in ["A1", "A2"] {
        assert!(!renders_italic(&wb, at), "{at}");
    }
    assert!(renders_italic(&wb, "B1"));
}

#[test]
fn a_default_full_format_under_an_italic_default_xf_renders_plain() {
    let mut wb = italic_default_workbook();
    apply(
        &mut wb,
        &Op::ApplyRangeFormat {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A2").unwrap(),
            format: xlsx_ops::CapturedFormat {
                rows: 1,
                columns: 1,
                formats: vec![xlsx_model::CellFormat::default()],
                kept: Vec::new(),
                source_styles: Vec::new(),
                source_formats: Vec::new(),
            },
        },
    )
    .unwrap();
    for at in ["A1", "A2"] {
        assert!(!renders_italic(&wb, at), "{at}");
    }
}

#[test]
fn clearing_to_a_plain_default_xf_stores_no_style() {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(
        r("A1"),
        xlsx_model::Cell {
            value: CellValue::Number { value: 1.0 },
            style: Some(1),
            ..Default::default()
        },
    );
    wb.sheets.push(sheet);
    wb.styles.fonts = vec![
        xlsx_model::Font::default(),
        xlsx_model::Font {
            italic: true,
            ..Default::default()
        },
    ];
    wb.styles.cell_xfs = vec![
        xlsx_model::Xf::default(),
        xlsx_model::Xf {
            font: Some(1),
            ..Default::default()
        },
    ];
    apply(
        &mut wb,
        &clear_italic(CellRange::parse_a1("A1:B2").unwrap()),
    )
    .unwrap();
    assert_eq!(wb.sheets[0].cell(r("A1")).unwrap().style, None);
    assert_eq!(wb.sheets[0].iter_cells().count(), 1);
    assert_eq!(wb.styles.cell_xfs.len(), 2);
}

#[test]
fn a_format_edit_that_changes_nothing_leaves_missing_cells_sparse() {
    let mut wb = Workbook::default();
    wb.sheets.push(Sheet::new("Sheet1"));
    wb.styles.fonts = vec![xlsx_model::Font {
        name: Some("Calibri".into()),
        size_pt: Some(11.0),
        ..Default::default()
    }];
    wb.styles.fills = vec![xlsx_model::Fill::None];
    wb.styles.borders = vec![xlsx_model::Border::default()];
    wb.styles.cell_xfs = vec![xlsx_model::Xf {
        font: Some(0),
        fill: Some(0),
        border: Some(0),
        num_fmt_id: Some(0),
        alignment: None,
        ..xlsx_model::Xf::default()
    }];
    let column = CellRange::new(r("A1"), CellRef::new(xlsx_model::MAX_ROWS - 1, 0));
    apply(&mut wb, &clear_italic(column)).unwrap();
    assert_eq!(wb.sheets[0].iter_cells().count(), 0);
    assert_eq!(wb.styles.cell_xfs.len(), 1);
}

#[test]
fn undoing_a_row_delete_at_the_row_style_cap_restores_the_map() {
    let cap = xlsx_model::workbook::MAX_ROW_STYLES as u32;
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    sheet.row_styles = (0..cap).map(|row| (row, 1)).collect();
    wb.sheets.push(sheet);
    let original = wb.sheets[0].row_styles.clone();
    let mut stack = UndoStack::new();
    stack
        .commit(
            &mut wb,
            &Transaction::new(
                vec![Op::DeleteRows {
                    sheet: SheetId(0),
                    at: cap,
                    count: 1,
                }],
                Provenance::User,
            ),
        )
        .unwrap();
    stack.undo(&mut wb).unwrap();
    assert_eq!(wb.sheets[0].row_styles, original);
    stack.redo(&mut wb).unwrap();
    assert_eq!(wb.sheets[0].row_styles, original);
}
