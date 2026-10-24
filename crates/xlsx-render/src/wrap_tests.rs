use super::*;
use xlsx_model::styles::{Alignment, Xf};
use xlsx_model::workbook::Cell;

fn wrap_alignment() -> Alignment {
    Alignment {
        wrap_text: true,
        ..Alignment::default()
    }
}

fn workbook(cells: &[(CellRef, CellValue, Alignment)]) -> Workbook {
    let mut wb = Workbook::default();
    let mut sheet = Sheet::new("Sheet1");
    for (index, (at, value, alignment)) in cells.iter().enumerate() {
        wb.styles.cell_xfs.push(Xf {
            alignment: Some(alignment.clone()),
            ..Xf::default()
        });
        sheet.set_cell(
            *at,
            Cell {
                value: value.clone(),
                style: Some(index as u32),
                ..Cell::default()
            },
        );
    }
    wb.sheets.push(sheet);
    wb
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn text_commands(wb: &Workbook) -> Vec<serde_json::Value> {
    let dl = build_display_list(
        wb,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        },
    )
    .unwrap();
    dl.commands
        .iter()
        .filter(|cmd| matches!(cmd, DrawCmd::Text { .. }))
        .map(|cmd| serde_json::to_value(cmd).unwrap())
        .collect()
}

fn row_height(wb: &Workbook, row: u32) -> f32 {
    let geom = GridGeometry::new(&wb.sheets[0], &wb.styles);
    geom.row_y(row + 1) - geom.row_y(row)
}

fn line_px(size_pt: f64) -> f32 {
    geometry::row_pt_to_px(geometry::autofit_row_height_pt(size_pt))
}

fn assert_lines(wb: &Workbook, row: u32, lines: f32) {
    let height = row_height(wb, row);
    assert!(
        (height - lines * line_px(11.0)).abs() < 1e-3,
        "row {row}: {height}"
    );
}

#[test]
fn a_wrapped_text_cell_carries_its_wrap_box() {
    let mut wb = workbook(&[(
        CellRef::new(0, 0),
        text("PURCHASE \nTARGET"),
        Alignment {
            h: Some(HAlign::Center),
            ..wrap_alignment()
        },
    )]);
    wb.sheets[0].col_widths.insert(0, 10.0);
    wb.sheets[0].row_heights.insert(0, 28.8);
    let cmds = text_commands(&wb);
    let wrap = &cmds[0]["wrap"];
    let width = geometry::col_chars_to_px(10.0) - 2.0 * TEXT_PAD_PX;
    assert_eq!(wrap["width"].as_f64(), Some(f64::from(width)));
    assert_eq!(wrap["lineHeight"].as_f64(), Some(f64::from(line_px(11.0))));
    assert_eq!(wrap["valign"], "bottom");
    assert_eq!(cmds[0]["text"], "PURCHASE \nTARGET");
}

#[test]
fn indents_narrow_the_wrap_width_and_valign_follows_the_cell() {
    let wb = workbook(&[(
        CellRef::new(0, 0),
        text("indented label"),
        Alignment {
            h: Some(HAlign::Left),
            v: Some(VAlign::Top),
            indent: Some(2),
            ..wrap_alignment()
        },
    )]);
    let cmds = text_commands(&wb);
    let width = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS)
        - 2.0 * TEXT_PAD_PX
        - 2.0 * geometry::indent_step_px(&wb.styles, None);
    assert_eq!(cmds[0]["wrap"]["width"].as_f64(), Some(f64::from(width)));
    assert_eq!(cmds[0]["wrap"]["valign"], "top");
}

#[test]
fn numbers_shrunk_and_rotated_text_never_wrap() {
    let wb = workbook(&[
        (
            CellRef::new(0, 0),
            CellValue::Number { value: 12345.5 },
            wrap_alignment(),
        ),
        (
            CellRef::new(1, 0),
            text("shrunk text"),
            Alignment {
                shrink_to_fit: true,
                ..wrap_alignment()
            },
        ),
        (
            CellRef::new(2, 0),
            text("rotated text"),
            Alignment {
                text_rotation: Some(45),
                ..wrap_alignment()
            },
        ),
        (
            CellRef::new(3, 0),
            text("stacked"),
            Alignment {
                text_rotation: Some(255),
                ..wrap_alignment()
            },
        ),
        (CellRef::new(4, 0), text("plain text"), Alignment::default()),
    ]);
    let cmds = text_commands(&wb);
    assert_eq!(cmds.len(), 5);
    for cmd in cmds {
        assert!(cmd.get("wrap").is_none(), "{cmd}");
    }
}

#[test]
fn an_unsized_row_grows_to_its_explicit_lines() {
    let wb = workbook(&[(CellRef::new(0, 0), text("a\nb\r\nc"), wrap_alignment())]);
    assert_lines(&wb, 0, 3.0);
}

#[test]
fn a_final_line_break_adds_no_line() {
    let wb = workbook(&[(CellRef::new(0, 0), text("a\nb\n"), wrap_alignment())]);
    assert_lines(&wb, 0, 2.0);
}

#[test]
fn an_unsized_row_grows_to_its_word_wrapped_lines() {
    let mut wb = workbook(&[(
        CellRef::new(0, 0),
        text("alpha beta gamma delta"),
        wrap_alignment(),
    )]);
    assert_lines(&wb, 0, 4.0);
    wb.sheets[0].col_widths.insert(0, 12.0);
    assert_lines(&wb, 0, 2.0);
}

#[test]
fn wrapped_rows_keep_explicit_heights_merges_and_unwrapped_cells() {
    let mut wb = workbook(&[
        (CellRef::new(0, 0), text("a\nb"), wrap_alignment()),
        (CellRef::new(1, 0), text("a\nb"), wrap_alignment()),
        (CellRef::new(2, 0), text("a\nb"), Alignment::default()),
        (
            CellRef::new(3, 0),
            CellValue::Number { value: 1.0 },
            wrap_alignment(),
        ),
    ]);
    wb.sheets[0].row_heights.insert(0, 20.0);
    wb.sheets[0]
        .merges
        .push(CellRange::new(CellRef::new(1, 0), CellRef::new(1, 2)));
    let default = geometry::row_pt_to_px(geometry::DEFAULT_ROW_HEIGHT_PT);
    assert_eq!(row_height(&wb, 0), geometry::row_pt_to_px(20.0));
    for row in 1..4 {
        assert!((row_height(&wb, row) - default).abs() < 1e-3, "row {row}");
    }
}

#[test]
fn a_pinned_sheet_never_grows_wrapped_rows() {
    let mut wb = workbook(&[(CellRef::new(0, 0), text("a\nb\nc"), wrap_alignment())]);
    wb.sheets[0].format.custom_height = true;
    assert_eq!(
        row_height(&wb, 0),
        geometry::row_pt_to_px(geometry::DEFAULT_ROW_HEIGHT_PT)
    );
}

#[test]
fn print_wraps_on_the_print_row_grid() {
    let wb = workbook(&[(CellRef::new(0, 0), text("BUY\nSELL"), wrap_alignment())]);
    let metrics = PrintMetrics {
        dpi: 72.0,
        max_digit_width: 6.0,
        default_row_height_pt: 15.0,
        default_column_width: None,
        font_size_pt: 11.0,
        font_family: "Calibri".into(),
        font_ascent: 10.0,
        font_descent: 3.0,
    };
    let geom = GridGeometry::for_print(&wb.sheets[0], &wb.styles, &metrics);
    let line = geometry::row_pt_to_px(14.0);
    assert!((geom.row_y(1) - geom.row_y(0) - 2.0 * line).abs() < 1e-3);
    let dl = build_print_display_list_with_charts(
        &wb,
        SheetId(0),
        &Viewport {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        },
        &metrics,
        false,
        |_| {
            Err::<std::sync::Arc<ooxml_drawingml::chart::ChartSpace>, _>(
                RenderError::ChartSourceUnavailable {
                    part: "chart".into(),
                },
            )
        },
    )
    .unwrap();
    let text = dl
        .commands
        .iter()
        .find(|cmd| matches!(cmd, DrawCmd::Text { .. }))
        .map(|cmd| serde_json::to_value(cmd).unwrap())
        .unwrap();
    assert_eq!(text["wrap"]["lineHeight"].as_f64(), Some(f64::from(line)));
}

#[test]
fn a_wrapped_row_stops_at_excels_maximum_height() {
    let wb = workbook(&[(
        CellRef::new(0, 0),
        text(&vec!["a"; 1000].join("\n")),
        wrap_alignment(),
    )]);
    assert_eq!(row_height(&wb, 0), geometry::row_pt_to_px(409.5));
}

#[test]
fn a_cell_without_a_style_wraps_on_xf_0() {
    let mut wb = workbook(&[(CellRef::new(0, 0), text("a\nb"), wrap_alignment())]);
    wb.sheets[0].cell_mut(CellRef::new(0, 0)).unwrap().style = None;
    assert_lines(&wb, 0, 2.0);
    let wb_cmds = text_commands(&wb);
    assert!(wb_cmds[0]["wrap"].is_object(), "{wb_cmds:?}");
}

#[test]
fn accounting_text_padding_still_wraps() {
    let mut wb = workbook(&[(
        CellRef::new(0, 0),
        text("alpha beta gamma delta"),
        wrap_alignment(),
    )]);
    wb.styles.cell_xfs[0].num_fmt_id = Some(44);
    assert_lines(&wb, 0, 4.0);
    let cmds = text_commands(&wb);
    assert!(cmds[0]["wrap"].is_object(), "{cmds:?}");
    assert!(cmds[0].get("numberFormat").is_none(), "{cmds:?}");
}

#[test]
fn an_unsized_row_grows_to_the_formatted_text_it_wraps() {
    let mut wb = workbook(&[(CellRef::new(0, 0), text("x"), wrap_alignment())]);
    wb.styles
        .num_fmts
        .push((164, "0;0;0;\"alpha beta gamma delta \"@".into()));
    wb.styles.cell_xfs[0].num_fmt_id = Some(164);
    assert_lines(&wb, 0, 4.0);
    let cmds = text_commands(&wb);
    assert_eq!(cmds[0]["text"], "alpha beta gamma delta x", "{cmds:?}");
    assert!(cmds[0]["wrap"].is_object(), "{cmds:?}");
}
