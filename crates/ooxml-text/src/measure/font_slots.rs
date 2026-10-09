//! The `w:rFonts` slot each character of a text run is measured with.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum FontSlot {
    Ascii,
    HAnsi,
    EastAsia,
    Cs,
}

impl FontSlot {
    /// Index into a run's per-slot resolved chains.
    pub(super) fn index(self) -> usize {
        match self {
            FontSlot::Ascii => 0,
            FontSlot::HAnsi => 1,
            FontSlot::EastAsia => 2,
            FontSlot::Cs => 3,
        }
    }
}

fn is_complex(ch: char) -> bool {
    matches!(ch as u32,
        0x0590..=0x08ff | 0x0900..=0x0dff | 0xfb1d..=0xfdff | 0xfe70..=0xfeff)
}

fn is_east_asian(ch: char) -> bool {
    matches!(ch as u32,
        0x1100..=0x11ff | 0x2e80..=0x30ff | 0x3130..=0x318f |
        0x31a0..=0x31ff | 0x3400..=0x4dbf | 0x4e00..=0x9fff |
        0xa960..=0xa97f | 0xac00..=0xd7ff | 0xf900..=0xfaff |
        0xff00..=0xffef)
}

fn is_combining(ch: char) -> bool {
    matches!(ch as u32, 0x0300..=0x036f | 0x1ab0..=0x1aff | 0x1dc0..=0x1dff | 0x20d0..=0x20ff | 0xfe20..=0xfe2f)
}

/// The slot `ch` is measured with in a run marked `w:cs` when `complex_script`
/// and hinted `w:hint="eastAsia"` when `east_asia_hint`; a combining mark keeps
/// `previous`, the slot of the character before it.
pub(super) fn char_slot(
    ch: char,
    complex_script: bool,
    east_asia_hint: bool,
    previous: FontSlot,
) -> FontSlot {
    if complex_script || is_complex(ch) {
        FontSlot::Cs
    } else if is_east_asian(ch) {
        FontSlot::EastAsia
    } else if is_combining(ch) {
        previous
    } else if ch.is_ascii() {
        FontSlot::Ascii
    } else if east_asia_hint {
        FontSlot::EastAsia
    } else {
        FontSlot::HAnsi
    }
}

/// Whether measuring a text run takes faces from its East Asian or its
/// complex-script `w:rFonts` slot.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct FontSlotUse {
    pub east_asia: bool,
    pub complex_script: bool,
}

/// The slots measurement takes `text`'s faces from, for a run marked `w:cs`
/// when `complex_script` and with the `w:rFonts/@w:hint` `hint`.
pub fn font_slot_use(text: &str, complex_script: bool, hint: Option<&str>) -> FontSlotUse {
    if !complex_script && text.is_ascii() {
        return FontSlotUse::default();
    }
    let east_asia_hint = hint == Some("eastAsia");
    let mut used = FontSlotUse::default();
    let mut slot = FontSlot::HAnsi;
    for ch in text.chars() {
        slot = char_slot(ch, complex_script, east_asia_hint, slot);
        match slot {
            FontSlot::EastAsia => used.east_asia = true,
            FontSlot::Cs => used.complex_script = true,
            FontSlot::Ascii | FontSlot::HAnsi => {}
        }
    }
    used
}
