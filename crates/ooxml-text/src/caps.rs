//! Casing shared by `w:caps`/`w:smallCaps` and `a:rPr/@cap`, so DOCX and PPTX
//! uppercase and scale small caps the same way.

/// Advance scale for a synthesized small cap — an uppercase glyph standing in
/// for a lowercase character because the face carries no `smcp` substitution.
/// The browser value is Blink and WebKit's synthesis multiplier; the Word
/// value is what Word renders small caps at, and applies under
/// `authoritativeShaping`.
pub const BROWSER_SMALL_CAPS_ADVANCE_SCALE: f32 = 0.7;
pub const WORD_SMALL_CAPS_ADVANCE_SCALE: f32 = 0.8;

/// Uppercase forms of `ch` for shaping, honouring the Turkish and Azerbaijani
/// dotted-I rules. One character can expand to several (ß → SS).
pub fn uppercase_for_language(ch: char, language: Option<&str>) -> Vec<char> {
    let lang = language.unwrap_or("").to_ascii_lowercase();
    if lang.starts_with("tr") || lang.starts_with("az") {
        match ch {
            'i' => return vec!['İ'],
            'ı' => return vec!['I'],
            _ => {}
        }
    }
    ch.to_uppercase().collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn uppercasing_expands_and_follows_the_dotted_i_rules() {
        assert_eq!(uppercase_for_language('ß', None), vec!['S', 'S']);
        assert_eq!(uppercase_for_language('i', None), vec!['I']);
        assert_eq!(uppercase_for_language('i', Some("tr-TR")), vec!['İ']);
        assert_eq!(uppercase_for_language('ı', Some("az")), vec!['I']);
        assert_eq!(uppercase_for_language('ı', Some("en-US")), vec!['I']);
    }
}
