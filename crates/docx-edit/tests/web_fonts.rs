#[test]
fn decode_web_fonts_without_changing_registered_fonts() {
    let raw = include_bytes!("fonts/LiberationSans-Regular.ttf");
    for font in [
        include_bytes!("fonts/LiberationSans-Regular.woff").as_slice(),
        include_bytes!("fonts/LiberationSans-Regular.woff2").as_slice(),
    ] {
        let decoded = docx_edit::decode_resident_font(font).unwrap();
        let mut registry = ooxml_text::FontStore::new();
        registry.register(decoded).unwrap();
        assert_eq!(registry.font_count(), 1);
    }
    assert_eq!(docx_edit::decode_resident_font(raw).unwrap(), raw);
    for font in [b"wOF2".as_slice(), b"wOFF".as_slice(), &[1, 2, 3]] {
        assert!(docx_edit::decode_resident_font(font).is_err());
    }
}
