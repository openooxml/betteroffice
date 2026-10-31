#![cfg(feature = "fonts")]

use xlsx_raster::RasterFonts;

const FONT: &[u8] = include_bytes!("../../betteroffice-fonts/tests/fixtures/Synthetic-Regular.ttf");

#[test]
fn family_aliases_share_retained_face_bytes() {
    let mut fonts = RasterFonts::default();
    for index in 0..300 {
        fonts
            .register_font(
                &format!("Alias {index}"),
                false,
                false,
                FONT.to_vec(),
                false,
                false,
            )
            .unwrap();
    }
    let first = fonts.font_data("Alias 0", false, false).unwrap();
    for index in 1..300 {
        let alias = fonts
            .font_data(&format!("Alias {index}"), false, false)
            .unwrap();
        assert_eq!(first.as_ptr(), alias.as_ptr());
    }
}

#[test]
fn registration_and_face_budgets_are_enforced() {
    let mut fonts = RasterFonts::default();
    for index in 0..4096 {
        fonts
            .register_font(
                &format!("Alias {index}"),
                false,
                false,
                FONT.to_vec(),
                false,
                false,
            )
            .unwrap();
    }
    assert!(
        fonts
            .register_font("Overflow", false, false, FONT.to_vec(), false, false)
            .is_err()
    );
    let mut fonts = RasterFonts::default();
    for index in 0_u32..256 {
        let mut font = FONT.to_vec();
        font.extend_from_slice(&index.to_le_bytes());
        fonts
            .register_font(&format!("Face {index}"), false, false, font, false, false)
            .unwrap();
    }
    let mut font = FONT.to_vec();
    font.extend_from_slice(&256_u32.to_le_bytes());
    assert!(
        fonts
            .register_font("Overflow", false, false, font, false, false)
            .is_err()
    );
}

#[test]
fn oversized_faces_are_rejected() {
    let mut bytes = FONT.to_vec();
    bytes.resize(32 * 1024 * 1024 + 1, 0);
    assert!(
        RasterFonts::default()
            .register_font("Oversized", false, false, bytes, false, false)
            .is_err()
    );
}
