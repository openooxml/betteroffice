use xlsx_model::styles::{Font, Xf};
use xlsx_model::workbook::{Cell, Sheet};
use xlsx_model::{CellRef, CellValue, Stylesheet};
use xlsx_render::{GridGeometry, PrintMetrics};

const PX_PER_PT: f32 = 96.0 / 72.0;

fn metrics(family: &str, size_pt: f32, default_row_height_pt: f32) -> PrintMetrics {
    PrintMetrics {
        dpi: 72.0,
        max_digit_width: 6.0,
        default_row_height_pt,
        default_column_width: None,
        font_size_pt: size_pt,
        font_family: family.into(),
        font_ascent: size_pt * 0.9,
        font_descent: size_pt * 0.2,
    }
}

fn styles(family: &str, sizes: &[f64]) -> Stylesheet {
    let mut styles = Stylesheet::default();
    styles.fonts = sizes
        .iter()
        .map(|&size| Font {
            name: Some(family.into()),
            size_pt: Some(size),
            ..Font::default()
        })
        .collect();
    styles.cell_xfs = (0..sizes.len())
        .map(|font| Xf {
            font: Some(font as u32),
            ..Xf::default()
        })
        .collect();
    styles
}

fn sheet(default_row_height_pt: f64) -> Sheet {
    let mut sheet = Sheet::new("S");
    sheet.format.default_row_height_pt = Some(default_row_height_pt);
    sheet
}

fn number(sheet: &mut Sheet, row: u32, style: Option<u32>) {
    sheet.set_cell(
        CellRef::new(row, 0),
        Cell {
            value: CellValue::Number { value: 1.0 },
            style,
            ..Default::default()
        },
    );
}

fn row_pt(grid: &GridGeometry, row: u32) -> f32 {
    (grid.row_y(row + 1) - grid.row_y(row)) / PX_PER_PT
}

fn assert_pt(actual: f32, expected: f32) {
    assert!(
        (actual - expected).abs() < 0.01,
        "{actual}pt, expected {expected}pt"
    );
}

#[test]
fn print_rows_without_a_height_take_the_normal_fonts_height_on_the_device_grid() {
    let styles = styles("Arial", &[10.0]);
    let mut sheet = sheet(12.75);
    number(&mut sheet, 0, None);
    let grid = GridGeometry::for_print(&sheet, &styles, &metrics("Arial", 10.0, 12.75));
    assert_pt(row_pt(&grid, 0), 13.0);
    assert_pt(row_pt(&grid, 1), 13.0);

    let styles = self::styles("Calibri", &[11.0]);
    let mut sheet = self::sheet(15.0);
    number(&mut sheet, 0, None);
    let grid = GridGeometry::for_print(&sheet, &styles, &metrics("Calibri", 11.0, 15.0));
    assert_pt(row_pt(&grid, 0), 14.0);
    assert_pt(row_pt(&grid, 1), 14.0);
}

#[test]
fn print_keeps_an_authored_default_row_height_far_from_the_font() {
    let styles = styles("Arial", &[10.0]);
    let mut sheet = sheet(11.25);
    number(&mut sheet, 0, None);
    let grid = GridGeometry::for_print(&sheet, &styles, &metrics("Arial", 10.0, 11.25));
    assert_pt(row_pt(&grid, 0), 11.0);
    assert_pt(row_pt(&grid, 1), 11.0);
}

#[test]
fn print_rows_grow_to_a_larger_font() {
    let styles = styles("Arial", &[10.0, 16.0]);
    let mut sheet = sheet(12.75);
    number(&mut sheet, 0, Some(1));
    let grid = GridGeometry::for_print(&sheet, &styles, &metrics("Arial", 10.0, 12.75));
    assert_pt(row_pt(&grid, 0), 20.0);
}

#[test]
fn print_row_heights_land_on_whole_device_pixels() {
    let styles = styles("Arial", &[10.0]);
    let mut sheet = sheet(12.75);
    sheet.row_heights.insert(0, 13.5);
    sheet.row_heights.insert(1, 20.25);
    let grid = GridGeometry::for_print(&sheet, &styles, &metrics("Arial", 10.0, 12.75));
    assert_pt(row_pt(&grid, 0), 13.0);
    assert_pt(row_pt(&grid, 1), 20.0);
}

#[test]
fn print_default_columns_are_base_width_digits_plus_padding() {
    let styles = styles("Arial", &[10.0]);
    let mut sheet = sheet(12.75);
    sheet.col_widths.insert(3, 20.7109375);
    let grid = GridGeometry::for_print(&sheet, &styles, &metrics("Arial", 10.0, 12.75));
    assert_pt((grid.col_x(1) - grid.col_x(0)) / PX_PER_PT, 53.0);
    assert_pt((grid.col_x(4) - grid.col_x(3)) / PX_PER_PT, 124.0);
}

#[test]
fn screen_geometry_keeps_its_own_defaults() {
    for family in ["Arial", "Calibri"] {
        let styles = styles(family, &[10.0]);
        let mut sheet = sheet(12.75);
        number(&mut sheet, 0, None);
        let grid = GridGeometry::new(&sheet, &styles);
        assert!((grid.col_x(1) - grid.col_x(0) - 64.0).abs() < 0.02);
        assert_pt(row_pt(&grid, 0), 12.5);
        assert_pt(row_pt(&grid, 1), 12.75);
    }
}
