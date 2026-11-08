//! Legacy symbol faces — Symbol, Wingdings, Wingdings 2, Wingdings 3 and
//! Webdings — address glyphs by font position through a `(3, 0)` private-use cmap at
//! `U+F020..=U+F0FF`, or by the raw byte. The character a file stores for one
//! of them is a slot number, not the character to draw, so a host without the
//! face draws a Latin letter or a tofu box.
//!
//! These tables name the nearest Unicode character our bundled faces cover, so
//! a bullet drawn through a symbol face still reads as the shape it is. Each
//! slot's shape comes from the face's own glyph name (`square2`, `trianglert`)
//! and its size from the ink width of that glyph, matched against the ink width
//! of the candidates in Liberation Sans at the same em.
//!
//! Symbol uses the Adobe Symbol Encoding to Unicode table published by the
//! Unicode Consortium (`adobe-symbol.txt`, kept verbatim with its notice).
//! Its Corporate Use targets (bracket pieces, serif and sans-serif ©®™) are
//! private-use characters no face covers, so those slots count as unmapped.

use std::sync::OnceLock;

/// A symbol face that addresses glyphs by font position.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SymbolFont {
    Symbol,
    Wingdings,
    Wingdings2,
    Wingdings3,
    Webdings,
}

/// Stands in for a slot with no covered equivalent: still a bullet, where the
/// face's own character would be an unrelated letter or a missing glyph.
const UNKNOWN_SLOT: char = '\u{2022}';

impl SymbolFont {
    pub fn named(typeface: &str) -> Option<Self> {
        match typeface.trim().to_ascii_lowercase().as_str() {
            "symbol" => Some(Self::Symbol),
            "wingdings" => Some(Self::Wingdings),
            "wingdings 2" => Some(Self::Wingdings2),
            "wingdings 3" => Some(Self::Wingdings3),
            "webdings" => Some(Self::Webdings),
            _ => None,
        }
    }

    fn slots(self) -> &'static [(u8, char)] {
        match self {
            Self::Symbol => adobe_symbol(),
            Self::Wingdings => &WINGDINGS,
            Self::Wingdings2 => &WINGDINGS_2,
            Self::Wingdings3 => &WINGDINGS_3,
            Self::Webdings => &WEBDINGS,
        }
    }

    /// The character to draw in place of `character`, or `None` when it is not
    /// a font position at all and so already means itself.
    pub fn substitute(self, character: char) -> Option<char> {
        symbol_slot(character)?;
        Some(self.shape(character).unwrap_or(UNKNOWN_SLOT))
    }

    /// The covered character for the slot `character` addresses, or `None`
    /// when it is not a slot this face's table names.
    pub fn shape(self, character: char) -> Option<char> {
        let slot = symbol_slot(character)?;
        self.slots()
            .iter()
            .find(|(position, _)| *position == slot)
            .map(|(_, replacement)| *replacement)
    }

    /// `text` as this face draws it: every slot replaced by its covered
    /// character, or by the bullet stand-in when the table has none. `None`
    /// when `text` addresses no slot. Every replacement is one BMP code unit,
    /// like the slot it replaces, so UTF-16 offsets into `text` still hold.
    pub fn display_text(self, text: &str) -> Option<SymbolText> {
        if !text
            .chars()
            .any(|character| symbol_slot(character).is_some())
        {
            return None;
        }
        let mut unmapped = 0;
        let text = text
            .chars()
            .map(|character| match symbol_slot(character) {
                None => character,
                Some(_) => self.shape(character).unwrap_or_else(|| {
                    unmapped += 1;
                    UNKNOWN_SLOT
                }),
            })
            .collect();
        Some(SymbolText { text, unmapped })
    }
}

/// Text a symbol face draws, with the number of slots that fell back to the
/// bullet stand-in.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SymbolText {
    pub text: String,
    pub unmapped: usize,
}

const ADOBE_SYMBOL: &str = include_str!("adobe-symbol.txt");

/// The Adobe table's non-private-use targets by Symbol slot. A slot the table
/// lists twice (space, Delta, Omega, mu, fraction) keeps its first entry.
fn adobe_symbol() -> &'static [(u8, char)] {
    static TABLE: OnceLock<Vec<(u8, char)>> = OnceLock::new();
    TABLE.get_or_init(|| {
        let mut table: Vec<(u8, char)> = Vec::new();
        for line in ADOBE_SYMBOL.lines().filter(|line| !line.starts_with('#')) {
            let mut fields = line.split('\t');
            let (Some(unicode), Some(slot)) = (fields.next(), fields.next()) else {
                continue;
            };
            let (Ok(unicode), Ok(slot)) = (
                u32::from_str_radix(unicode, 16),
                u8::from_str_radix(slot, 16),
            ) else {
                continue;
            };
            let Some(character) = char::from_u32(unicode) else {
                continue;
            };
            if (0xE000..=0xF8FF).contains(&unicode)
                || table.iter().any(|(position, _)| *position == slot)
            {
                continue;
            }
            table.push((slot, character));
        }
        table
    })
}

/// The font position `character` addresses. `0x20` is the faces' own space, so
/// it stays a space rather than becoming a bullet.
fn symbol_slot(character: char) -> Option<u8> {
    match character as u32 {
        slot @ 0x21..=0xFF => Some(slot as u8),
        slot @ 0xF021..=0xF0FF => Some((slot - 0xF000) as u8),
        _ => None,
    }
}

const WINGDINGS: [(u8, char); 20] = [
    (0x6C, '●'), // circle6
    (0x6D, '○'), // circleshadowdwn
    (0x6E, '■'), // square6
    (0x6F, '□'), // box3
    (0x70, '□'), // box4
    (0x71, '□'), // boxshadowdwn
    (0x72, '□'), // boxshadowup
    (0x73, '♦'), // lozenge4
    (0x74, '♦'), // lozenge6
    (0x75, '♦'), // rhombus6
    (0x77, '♦'), // rhombus4
    (0x9E, '•'), // circle2
    (0x9F, '•'), // circle4
    (0xA0, '▪'), // square2
    (0xA6, '○'), // circleshadowup
    (0xA7, '▪'), // square4
    (0xA8, '□'), // box2
    (0xD7, '◄'), // head2left
    (0xD8, '►'), // head2right
    (0xD9, '▲'), // head2up
];

const WINGDINGS_2: [(u8, char); 21] = [
    (0x95, '•'), // circle1
    (0x96, '•'), // circle3
    (0x97, '●'), // circle5
    (0x98, '●'), // circle7
    (0x9F, '▪'), // square1
    (0xA0, '▪'), // square3
    (0xA1, '■'), // square5
    (0xA2, '■'), // square7
    (0xA3, '□'), // box1
    (0xA4, '□'), // box5
    (0xA5, '□'), // box6
    (0xA6, '□'), // box7
    (0xAB, '♦'), // rhombus1
    (0xAC, '♦'), // rhombus2
    (0xAD, '♦'), // rhombus3
    (0xAE, '♦'), // rhombus5
    (0xB4, '♦'), // lozenge1
    (0xB5, '♦'), // lozenge2
    (0xB6, '♦'), // lozenge3
    (0xB7, '♦'), // lozenge5
    (0xB8, '◊'), // lozengeopen
];

const WINGDINGS_3: [(u8, char); 19] = [
    (0x70, '▲'), // triangleup
    (0x71, '▼'), // triangledwn
    (0x72, '▲'), // triangleopenup
    (0x73, '▼'), // triangleopendwn
    (0x74, '◄'), // triangleleft
    (0x75, '►'), // trianglert
    (0x76, '◄'), // triangleopenleft
    (0x77, '►'), // triangleopenrt
    (0x7C, '◄'), // triangle45left
    (0x7D, '►'), // triangle45right
    (0x7E, '▲'), // triangle45up
    (0x80, '▼'), // triangle45down
    (0x81, '▲'), // trianglecentup
    (0x82, '▼'), // trianglecentdwn
    (0x83, '◄'), // trianglecentleft
    (0x84, '►'), // trianglecentrt
    (0x85, '◄'), // headleft
    (0x86, '►'), // headright
    (0x87, '▲'), // headup
];

const WEBDINGS: [(u8, char); 3] = [
    (0x63, '□'), // boxopen
    (0x67, '■'), // boxsolid
    (0x6E, '●'), // circlesolid
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_symbol_face_is_matched_by_name_only() {
        assert_eq!(SymbolFont::named("Wingdings"), Some(SymbolFont::Wingdings));
        assert_eq!(
            SymbolFont::named(" wingdings 3 "),
            Some(SymbolFont::Wingdings3)
        );
        assert_eq!(SymbolFont::named("Webdings"), Some(SymbolFont::Webdings));
        assert_eq!(SymbolFont::named("Arial"), None);
        assert_eq!(SymbolFont::named("Wingdings Pro"), None);
    }

    #[test]
    fn a_slot_resolves_from_either_its_byte_or_its_private_use_form() {
        assert_eq!(SymbolFont::Wingdings.substitute('§'), Some('▪'));
        assert_eq!(SymbolFont::Wingdings.substitute('\u{f0a7}'), Some('▪'));
        assert_eq!(SymbolFont::Wingdings.substitute('\u{f0a0}'), Some('▪'));
        assert_eq!(SymbolFont::Wingdings.substitute('l'), Some('●'));
        assert_eq!(SymbolFont::Wingdings3.substitute('\u{f075}'), Some('►'));
    }

    #[test]
    fn symbol_slots_follow_the_adobe_table_in_both_ranges() {
        assert_eq!(SymbolFont::named("Symbol"), Some(SymbolFont::Symbol));
        for (slot, expected) in [
            ('a', 'α'),
            ('p', 'π'),
            ('D', 'Δ'),
            ('W', 'Ω'),
            ('m', 'µ'),
            ('\u{a4}', '⁄'),
            ('\u{b7}', '•'),
            ('\u{ae}', '→'),
            ('\u{a3}', '≤'),
            ('\u{d6}', '√'),
            ('1', '1'),
        ] {
            let private_use = char::from_u32(0xF000 + slot as u32).expect("private-use slot");
            assert_eq!(
                SymbolFont::Symbol.substitute(slot),
                Some(expected),
                "{slot:?}"
            );
            assert_eq!(
                SymbolFont::Symbol.substitute(private_use),
                Some(expected),
                "{private_use:?}"
            );
        }
        assert_eq!(SymbolFont::Symbol.substitute(' '), None);
        assert_eq!(SymbolFont::Symbol.substitute('α'), None);
    }

    #[test]
    fn symbol_keeps_one_entry_per_slot_and_drops_corporate_use_targets() {
        let table = adobe_symbol();
        assert_eq!(table.len(), 160);
        for (index, (slot, _)) in table.iter().enumerate() {
            assert!(
                table[index + 1..].iter().all(|(other, _)| other != slot),
                "{slot:#x} listed twice"
            );
        }
        assert_eq!(SymbolFont::Symbol.shape('\u{f0e6}'), None);
        assert_eq!(SymbolFont::Symbol.substitute('\u{f0e6}'), Some('•'));
        assert_eq!(SymbolFont::Symbol.substitute('\u{f080}'), Some('•'));
    }

    #[test]
    fn every_target_is_one_bmp_code_unit_like_its_slot() {
        for face in [
            SymbolFont::Symbol,
            SymbolFont::Wingdings,
            SymbolFont::Wingdings2,
            SymbolFont::Wingdings3,
            SymbolFont::Webdings,
        ] {
            for slot in (0x21..=0xFFu32).chain(0xF021..=0xF0FF) {
                let character = char::from_u32(slot).expect("slot char");
                let drawn = face.substitute(character).expect("slot");
                assert_eq!(drawn.len_utf16(), 1, "{face:?} {slot:#x}");
                assert_eq!(drawn.len_utf16(), character.len_utf16());
                assert!(
                    !(0xE000..=0xF8FF).contains(&(drawn as u32)),
                    "{face:?} {slot:#x}"
                );
            }
        }
    }

    #[test]
    fn display_text_maps_slots_counts_stand_ins_and_keeps_utf16_length() {
        let text = "a+b \u{f0b3} \u{f0e6}x日";
        let drawn = SymbolFont::Symbol.display_text(text).expect("slots");
        assert_eq!(drawn.text, "α+β ≥ •ξ日");
        assert_eq!(drawn.unmapped, 1);
        assert_eq!(
            drawn.text.encode_utf16().count(),
            text.encode_utf16().count()
        );
        assert_eq!(SymbolFont::Symbol.display_text("日本 αβ"), None);
        let wingdings = SymbolFont::Wingdings
            .display_text("l\u{f0e0}\u{f0fc}")
            .expect("slots");
        assert_eq!(wingdings.text, "●••");
        assert_eq!(wingdings.unmapped, 2);
    }

    #[test]
    fn an_unmapped_slot_still_reads_as_a_bullet_and_a_non_slot_is_left_alone() {
        assert_eq!(SymbolFont::Wingdings.substitute('\u{f0fc}'), Some('•'));
        assert_eq!(SymbolFont::Wingdings.substitute('\u{f020}'), None);
        assert_eq!(SymbolFont::Wingdings.substitute(' '), None);
        assert_eq!(SymbolFont::Wingdings.substitute('•'), None);
        assert_eq!(SymbolFont::Wingdings.substitute('\u{2013}'), None);
    }
}
