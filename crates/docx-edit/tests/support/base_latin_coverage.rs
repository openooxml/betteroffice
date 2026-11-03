use serde_json::json;

const PRIMARY: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn subset(character: char) -> Vec<u8> {
    let mut store = ooxml_text::FontStore::new();
    let id = store.register_slice(PRIMARY).unwrap();
    let glyph = store.glyph_id(id, character).unwrap().unwrap();
    let mut bytes = PRIMARY.to_vec();
    let count = u16::from_be_bytes(bytes[4..6].try_into().unwrap()) as usize;
    let record = (0..count)
        .map(|index| 12 + index * 16)
        .find(|&index| &bytes[index..index + 4] == b"cmap")
        .unwrap();
    let offset = bytes.len() as u32;
    bytes[record + 8..record + 12].copy_from_slice(&offset.to_be_bytes());
    bytes[record + 12..record + 16].copy_from_slice(&40u32.to_be_bytes());
    bytes.extend_from_slice(&[0, 0, 0, 1, 0, 3, 0, 10, 0, 0, 0, 12]);
    bytes.extend_from_slice(&[0, 12, 0, 0]);
    for value in [
        28,
        0,
        1,
        character as u32,
        character as u32,
        u32::from(glyph),
    ] {
        bytes.extend_from_slice(&value.to_be_bytes());
    }
    bytes
}

pub fn check(
    mut register: impl FnMut(&[u8]) -> u32,
    mut measure: impl FnMut(&str) -> bool,
    mut take: impl FnMut() -> bool,
) {
    let has_backstop = ooxml_text::FontStore::new()
        .font_bytes(ooxml_text::FontId::base_latin(false, false))
        .is_ok();
    let primary = register(PRIMARY);
    let upper = register(&subset('A'));
    let turkish = register(&subset('İ'));
    let input = |id, text: &str, caps: bool, small_caps: bool, language: &str| {
        json!({
            "block": {"kind": "paragraph", "id": "coverage", "runs": [{"kind": "text", "text": text,
                "fontFamily": "Requested", "allCaps": caps, "smallCaps": small_caps,
                "language": {"latin": language}}]},
            "maxWidth": 600, "fontChains": {"requested|0|0": [id]},
            "defaults": {"fontSize": 11, "fontFamily": "Requested"}, "authoritativeShaping": true
        })
        .to_string()
    };
    assert!(measure(&input(primary, "A", false, false, "en-US")));
    assert!(!take());
    for text in ["a", "aa", "aaa"] {
        assert!(measure(&input(upper, text, true, false, "en-US")));
        assert!(!take());
        assert!(measure(&input(upper, text, false, true, "en-US")));
        assert!(!take());
    }
    assert!(measure(&input(turkish, "i", true, false, "tr-TR")));
    assert!(!take());
    assert!(measure(&input(turkish, "A", false, false, "en-US")));
    assert_eq!(take(), !has_backstop);
    assert!(!take());
    assert!(measure(&input(turkish, "A", false, false, "en-US")));
    assert!(!take());
    assert!(measure(&input(turkish, "AΩ", false, false, "en-US")));
    assert_eq!(take(), !has_backstop);
    assert!(measure(&input(primary, "Ω", false, false, "en-US")));
    assert!(!take());
}
