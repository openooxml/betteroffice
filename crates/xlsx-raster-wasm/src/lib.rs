use wasm_bindgen::prelude::*;
use xlsx_render::DisplayList;

#[wasm_bindgen]
pub fn render_png(display_list_json: &str) -> Result<Vec<u8>, String> {
    let display_list: DisplayList =
        serde_json::from_str(display_list_json).map_err(|error| error.to_string())?;
    xlsx_raster::render_png(&display_list)
}

#[cfg(test)]
mod tests {
    use super::*;
    use betteroffice_xlsx::{CellRange, RenderOptions, SheetId, Viewport, Workbook};

    fn round_trip(display_list: &DisplayList) -> Vec<u8> {
        let json = serde_json::to_string(display_list).unwrap();
        let decoded: DisplayList = serde_json::from_str(&json).unwrap();
        assert_eq!(&decoded, display_list);
        render_png(&json).unwrap()
    }

    #[test]
    fn json_raster_matches_workbook_png_bytes() {
        for bytes in [
            include_bytes!("../../../packages/xlsx/test-fixtures/sample.xlsx").as_slice(),
            include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx").as_slice(),
        ] {
            let workbook = Workbook::open(bytes).unwrap();
            for index in 0..workbook.sheet_count() {
                let sheet = SheetId(index as u32);
                for viewport in [
                    Viewport {
                        x: 0.0,
                        y: 0.0,
                        width: 320.0,
                        height: 160.0,
                    },
                    Viewport {
                        x: 13.37,
                        y: 21.125,
                        width: 300.3,
                        height: 180.7,
                    },
                    Viewport {
                        x: 90.5,
                        y: 43.2,
                        width: 640.0,
                        height: 360.0,
                    },
                ] {
                    let (display_list, width, height) =
                        workbook.png_display_list_for(sheet, &viewport).unwrap();
                    let full = workbook.render_png_for(sheet, &viewport).unwrap();
                    assert_eq!((width, height), (full.width, full.height));
                    assert_eq!(round_trip(&display_list), full.bytes);
                }
                for range in [None, Some("A1:B3"), Some("B2:D5")] {
                    for scale in [0.75, 1.0, 1.3, 2.0] {
                        let options = RenderOptions {
                            range: range.map(|range| CellRange::parse_a1(range).unwrap()),
                            scale,
                            max_width: Some(300),
                            max_height: Some(180),
                        };
                        let (display_list, width, height) = workbook
                            .range_png_display_list_for(sheet, &options)
                            .unwrap();
                        let full = workbook.render_sheet(sheet, &options).unwrap();
                        assert_eq!((width, height), (full.width, full.height));
                        assert_eq!(round_trip(&display_list), full.bytes);
                    }
                }
            }
        }
    }
}
