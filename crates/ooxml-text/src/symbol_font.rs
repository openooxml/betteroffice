//! Legacy symbol faces — Wingdings, Wingdings 2, Wingdings 3 and Webdings —
//! address glyphs by font position through a `(3, 0)` private-use cmap at
//! `U+F020..=U+F0FF`, or by the raw byte. The character a file stores for one
//! of them is a slot number, not the character to draw, so a host without the
//! face draws a Latin letter or a tofu box.
//!
//! These tables name the nearest Unicode character our bundled faces cover, so
//! a bullet drawn through a symbol face still reads as the shape it is. Each
//! slot's shape comes from the face's own glyph name (`square2`, `trianglert`)
//! and its size from the ink width of that glyph, matched against the ink width
//! of the candidates in Liberation Sans at the same em.

/// A symbol face that addresses glyphs by font position.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SymbolFont {
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
            "wingdings" => Some(Self::Wingdings),
            "wingdings 2" => Some(Self::Wingdings2),
            "wingdings 3" => Some(Self::Wingdings3),
            "webdings" => Some(Self::Webdings),
            _ => None,
        }
    }

    fn slots(self) -> &'static [(u8, char)] {
        match self {
            Self::Wingdings => &WINGDINGS,
            Self::Wingdings2 => &WINGDINGS_2,
            Self::Wingdings3 => &WINGDINGS_3,
            Self::Webdings => &WEBDINGS,
        }
    }

    /// The character to draw in place of `character`, or `None` when it is not
    /// a font position at all and so already means itself.
    pub fn substitute(self, character: char) -> Option<char> {
        let slot = symbol_slot(character)?;
        Some(
            self.slots()
                .iter()
                .find(|(position, _)| *position == slot)
                .map_or(UNKNOWN_SLOT, |(_, replacement)| *replacement),
        )
    }
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

const WINGDINGS: [(u8, char); 18] = [
    (0x6C, '●'), // circle6
    (0x6D, '○'), // circleshadowdwn
    (0x6E, '■'), // square6
    (0x6F, '□'), // box3
    (0x70, '□'), // box4
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
    fn an_unmapped_slot_still_reads_as_a_bullet_and_a_non_slot_is_left_alone() {
        assert_eq!(SymbolFont::Wingdings.substitute('\u{f0fc}'), Some('•'));
        assert_eq!(SymbolFont::Wingdings.substitute('\u{f020}'), None);
        assert_eq!(SymbolFont::Wingdings.substitute(' '), None);
        assert_eq!(SymbolFont::Wingdings.substitute('•'), None);
        assert_eq!(SymbolFont::Wingdings.substitute('\u{2013}'), None);
    }
}
