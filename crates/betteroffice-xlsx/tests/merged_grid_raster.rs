#![cfg(feature = "raster")]

use std::io::Cursor;

use betteroffice_xlsx::{
    CellRange, GridGeometry, Sheet, SheetId, Viewport, Workbook, WorkbookModel,
};

fn band_difference(pixels: &[u8], width: usize, x: f32, y: f32) -> u8 {
    let mut difference = 0;
    for row in y.floor() as usize - 2..=y.ceil() as usize + 2 {
        for col in x.floor() as usize - 2..=x.ceil() as usize + 2 {
            let index = (row * width + col) * 4;
            for channel in &pixels[index..index + 3] {
                difference = difference.max(255 - *channel);
            }
        }
    }
    difference
}

#[test]
fn merged_png_has_background_inside_and_grid_pixels_around_merge() {
    let mut sheet = Sheet::new("Merged");
    sheet.merges.push(CellRange::parse_a1("B2:C3").unwrap());
    for index in 0..6 {
        sheet.col_widths.insert(index, 10.0);
        sheet.row_heights.insert(index, 15.0);
    }
    let workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    let geom = GridGeometry::new(&workbook.model().sheets[0], &workbook.model().styles);
    let rendered = workbook
        .render_png_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: geom.col_x(5),
                height: geom.row_y(5),
            },
        )
        .unwrap();
    let mut reader = png::Decoder::new(Cursor::new(rendered.bytes))
        .read_info()
        .unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    assert_eq!(info.color_type, png::ColorType::Rgba);
    let width = info.width as usize;
    let x = (geom.col_x(1) + geom.col_x(2)) / 2.0;
    let y = (geom.row_y(1) + geom.row_y(2)) / 2.0;
    assert_eq!(band_difference(&pixels, width, x, y), 0);
    assert!(band_difference(&pixels, width, geom.col_x(2), y) <= 3);
    assert!(band_difference(&pixels, width, x, geom.row_y(2)) <= 3);
    for (x, y) in [
        (geom.col_x(1), y),
        (geom.col_x(3), y),
        (x, geom.row_y(1)),
        (x, geom.row_y(3)),
        (geom.col_x(2), geom.row_y(1) / 2.0),
    ] {
        let difference = band_difference(&pixels, width, x, y);
        assert!(difference > 8 && difference <= 46);
    }
}
