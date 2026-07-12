use std::cell::RefCell;

use ooxml_text::{FontId, FontStore};

thread_local! {
    static MEASURE_FONTS: RefCell<FontStore> = RefCell::new(FontStore::new());
}

pub fn build_display_list_json(input: &str) -> Result<String, String> {
    MEASURE_FONTS.with(|store| {
        docx_layout::display_list::build_display_list_json_with_fonts(input, &store.borrow())
    })
}

pub fn register_measure_font(bytes: &[u8]) -> Result<u32, String> {
    MEASURE_FONTS.with(|store| {
        store
            .borrow_mut()
            .register(bytes.to_vec())
            .map(|id| id.to_u32())
            .map_err(|e| e.to_string())
    })
}

pub fn clear_measure_fonts() {
    MEASURE_FONTS.with(|store| {
        *store.borrow_mut() = FontStore::new();
    });
}

pub fn measure_paragraph_json(input: &str) -> Result<String, String> {
    MEASURE_FONTS.with(|store| ooxml_text::measure_paragraph_json(&store.borrow(), input))
}

pub fn outline_glyph_json(font_id: u32, glyph_id: u32) -> Result<String, String> {
    let glyph_id = u16::try_from(glyph_id)
        .map_err(|_| format!("glyph id {glyph_id} out of range for this font"))?;
    MEASURE_FONTS.with(|store| {
        store
            .borrow()
            .outline_glyph_json(FontId::from_u32(font_id), glyph_id)
            .map_err(|e| e.to_string())
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn font_registry_drives_measurement_and_outlines() {
        clear_measure_fonts();
        let font_id = register_measure_font(FONT).expect("register font");
        let input = serde_json::json!({
            "block": {
                "kind": "paragraph",
                "runs": [{"kind": "text", "text": "Hello", "fontFamily": "Liberation Sans", "fontSize": 12}],
                "attrs": {},
                "pmStart": 0,
                "pmEnd": 6
            },
            "maxWidth": 400,
            "defaults": {"fontFamily": "Liberation Sans", "fontSize": 12},
            "fontChains": {"liberation sans|0|0": [font_id]}
        })
        .to_string();

        assert!(measure_paragraph_json(&input).is_ok());
        let outline = outline_glyph_json(font_id, 36).expect("outline glyph");
        assert!(outline.contains(r#""t":"M""#));
        clear_measure_fonts();
        assert_eq!(register_measure_font(FONT).expect("register font"), 0);
        clear_measure_fonts();
    }

    #[test]
    fn rejects_out_of_range_glyph_ids() {
        assert!(outline_glyph_json(0, u16::MAX as u32 + 1).is_err());
    }
}
