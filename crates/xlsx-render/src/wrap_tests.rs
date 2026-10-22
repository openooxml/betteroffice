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

fn line_px(size_pt: f64) -> f32 {
    geometry::row_pt_to_px(geometry::autofit_row_height_pt(size_pt))
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
        - 2.0 * INDENT_WIDTH_PX;
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
