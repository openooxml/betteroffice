pub const FAMILY: &str = "BetterOffice Base Latin";

pub static REGULAR: &[u8] = include_bytes!("../fonts/BetterOfficeBaseLatin-Regular.ttf");
pub static BOLD: &[u8] = include_bytes!("../fonts/BetterOfficeBaseLatin-Bold.ttf");
pub static ITALIC: &[u8] = include_bytes!("../fonts/BetterOfficeBaseLatin-Italic.ttf");
pub static BOLD_ITALIC: &[u8] = include_bytes!("../fonts/BetterOfficeBaseLatin-BoldItalic.ttf");

pub fn face(bold: bool, italic: bool) -> &'static [u8] {
    match (bold, italic) {
        (false, false) => REGULAR,
        (true, false) => BOLD,
        (false, true) => ITALIC,
        (true, true) => BOLD_ITALIC,
    }
}

pub fn faces() -> [(bool, bool, &'static [u8]); 4] {
    [
        (false, false, face(false, false)),
        (true, false, face(true, false)),
        (false, true, face(false, true)),
        (true, true, face(true, true)),
    ]
}

#[cfg(test)]
mod tests {
    use skrifa::raw::TableProvider;
    use skrifa::{FontRef, MetadataProvider};

    use super::*;

    #[test]
    fn faces_parse_with_renamed_names_and_real_styles() {
        for (bold, italic, bytes) in faces() {
            let font = FontRef::new(bytes).unwrap();
            let name = font.name().unwrap();
            let mut families = 0;
            let mut notices = 0;
            for record in name.name_record() {
                let value = record.string(name.string_data()).unwrap().to_string();
                match record.name_id().to_u16() {
                    0 => {
                        assert!(value.contains("Copyright (c) 2012 Red Hat, Inc."));
                        notices += 1;
                    }
                    13 => {
                        assert!(value.starts_with("SIL OPEN FONT LICENSE Version 1.1"));
                        notices += 1;
                    }
                    1 | 16 => {
                        assert_eq!(value, FAMILY);
                        families += 1;
                    }
                    _ => {}
                }
                if matches!(record.name_id().to_u16(), 1..=6 | 16..=18 | 21 | 22) {
                    for reserved in ["liberation", "arimo", "tinos", "cousine"] {
                        assert!(!value.to_lowercase().contains(reserved));
                    }
                }
            }
            assert!(families > 0);
            assert_eq!(notices, 2);
            let selection = font.os2().unwrap().fs_selection().bits();
            assert_eq!(selection & (1 << 5) != 0, bold);
            assert_eq!(selection & 1 != 0, italic);
        }
    }

    #[test]
    fn regular_covers_text_symbols_and_bullet_aliases() {
        let font = FontRef::new(face(false, false)).unwrap();
        let cmap = font.charmap();
        for point in (0x20u32..=0x7e).chain([
            0x00a0, 0x00e9, 0x0100, 0x0180, 0x024f, 0x03a9, 0x0416, 0x044f, 0x2014, 0x2022, 0x20ac,
            0x20b9, 0x2192, 0x2211, 0x25aa, 0x25c6, 0x25e6, 0x2713, 0x2756, 0x27a2, 0xfb01,
        ]) {
            assert!(cmap.map(point).is_some(), "missing U+{point:04X}");
        }
        for (private, point) in [
            (0xf0b7u32, 0x2022u32),
            (0xf0a7, 0x25aa),
            (0xf06f, 0x25e6),
            (0xf0d8, 0x27a2),
            (0xf0fc, 0x2713),
            (0xf076, 0x2756),
            (0xf0a8, 0x25aa),
        ] {
            assert!(cmap.map(private).is_some());
            assert_eq!(cmap.map(private), cmap.map(point));
        }
    }
}
