//! East Asian auto-spacing: `w:autoSpaceDE` and `w:autoSpaceDN` (§17.3.1.11,
//! §17.3.1.12), both defaulting to on.
//!
//! Word widens the gap where East Asian text meets Latin letters or digits.
//! Both sides of the boundary and both settings are measured off Word's own
//! PDF exports of the Japanese corpus: the advance of an East Asian character
//! standing before a Latin one is a flat 1.250 em against 1.000 em before
//! another East Asian character, and the same 0.250 em appears on the
//! Latin-then-East-Asian side. A document that turns the feature off
//! (`w:val="0"` on every paragraph) measures 1.000 em on both, which is what
//! pins the constant to the setting rather than to the face.
//!
//! Nothing is inserted next to a space. Word's exports put no extra advance
//! before or after U+0020 at a script boundary.

/// Extra advance at an East Asian / Latin boundary, as a fraction of the East
/// Asian side's font size. Measured at 0.250 em off Word's PDF exports.
pub const AUTO_SPACE_EM: f32 = 0.25;

/// Which auto-spacing settings a paragraph has on.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AutoSpace {
    /// `w:autoSpaceDE`: East Asian text against Latin letters.
    pub latin: bool,
    /// `w:autoSpaceDN`: East Asian text against digits.
    pub numeric: bool,
}

impl Default for AutoSpace {
    fn default() -> Self {
        Self {
            latin: true,
            numeric: true,
        }
    }
}

impl AutoSpace {
    /// Resolves the two `Option<bool>` opt-outs; absent is the OOXML default.
    pub fn from_options(latin: Option<bool>, numeric: Option<bool>) -> Self {
        Self {
            latin: latin.unwrap_or(true),
            numeric: numeric.unwrap_or(true),
        }
    }

    pub fn any(self) -> bool {
        self.latin || self.numeric
    }

    /// The extra advance between `prev` and `next` in em, 0.0 when the pair is
    /// not a boundary this paragraph spaces.
    pub fn extra_em(self, prev: char, next: char) -> f32 {
        let (left, right) = (class(prev), class(next));
        let spaced = match (left, right) {
            (Class::EastAsian, Class::Latin) | (Class::Latin, Class::EastAsian) => self.latin,
            (Class::EastAsian, Class::Numeric) | (Class::Numeric, Class::EastAsian) => self.numeric,
            _ => false,
        };
        if spaced { AUTO_SPACE_EM } else { 0.0 }
    }

    /// The same as an advance, sized off the East Asian side of the boundary.
    pub fn extra_px(self, prev: char, next: char, prev_px: f32, next_px: f32) -> f32 {
        let em = self.extra_em(prev, next);
        if em == 0.0 {
            return 0.0;
        }
        em * if class(prev) == Class::EastAsian {
            prev_px
        } else {
            next_px
        }
    }
}

#[derive(PartialEq, Eq, Clone, Copy)]
enum Class {
    EastAsian,
    Latin,
    Numeric,
    Other,
}

/// CJK punctuation (U+3000–U+303F) is deliberately `Other`: it carries its own
/// spacing rules, and U+3000 is a space.
fn class(ch: char) -> Class {
    let code = ch as u32;
    if matches!(code,
        0x3040..=0x30FF
        | 0x3100..=0x312F
        | 0x31A0..=0x31BF
        | 0x3190..=0x319F
        | 0x3400..=0x4DBF
        | 0x4E00..=0x9FFF
        | 0xF900..=0xFAFF
        | 0xA000..=0xA4CF
        | 0xAC00..=0xD7AF
        | 0x1100..=0x11FF
        | 0x3130..=0x318F
        | 0xFF01..=0xFF60
        | 0xFFE0..=0xFFE6
        | 0x20000..=0x2FA1F)
    {
        return Class::EastAsian;
    }
    if code >= 0x2000 {
        return Class::Other;
    }
    if ch.is_ascii_digit() {
        return Class::Numeric;
    }
    if ch.is_alphabetic() {
        return Class::Latin;
    }
    Class::Other
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_quarter_em_lands_on_both_sides_of_a_script_boundary() {
        let on = AutoSpace::default();
        assert_eq!(on.extra_em('国', 'a'), 0.25);
        assert_eq!(on.extra_em('a', '国'), 0.25);
        assert_eq!(on.extra_em('ぁ', '7'), 0.25);
        assert_eq!(on.extra_em('7', 'ア'), 0.25);
        assert_eq!(on.extra_em('国', '花'), 0.0);
        assert_eq!(on.extra_em('a', 'b'), 0.0);
    }

    #[test]
    fn a_space_or_ideographic_punctuation_never_takes_one() {
        let on = AutoSpace::default();
        for pair in [
            ('国', ' '),
            (' ', 'a'),
            ('国', '、'),
            ('。', 'a'),
            ('国', '\u{3000}'),
        ] {
            assert_eq!(on.extra_em(pair.0, pair.1), 0.0, "{pair:?}");
        }
    }

    #[test]
    fn each_setting_gates_only_its_own_boundary() {
        let no_latin = AutoSpace::from_options(Some(false), None);
        assert_eq!(no_latin.extra_em('国', 'a'), 0.0);
        assert_eq!(no_latin.extra_em('国', '1'), 0.25);
        let no_numeric = AutoSpace::from_options(None, Some(false));
        assert_eq!(no_numeric.extra_em('国', 'a'), 0.25);
        assert_eq!(no_numeric.extra_em('国', '1'), 0.0);
        assert!(!AutoSpace::from_options(Some(false), Some(false)).any());
    }
}
