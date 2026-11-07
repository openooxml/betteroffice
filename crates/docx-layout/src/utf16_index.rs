//! UTF-16 offsets into UTF-8 text without rescanning it from the start.
//!
//! Line resolution slices one run once per line it spans. A typed paragraph
//! is often a single long run, so walking the run from its start for every
//! line made a paragraph's layout grow with the square of its length.

use std::cell::OnceCell;

const STRIDE: usize = 32;

pub(crate) struct Utf16Index<'a> {
    text: &'a str,
    units: usize,
    marks: Vec<(usize, usize)>,
}

impl<'a> Utf16Index<'a> {
    pub(crate) fn new(text: &'a str) -> Self {
        if text.is_ascii() {
            return Self { text, units: text.len(), marks: Vec::new() };
        }
        let mut marks = Vec::with_capacity(text.len() / STRIDE + 1);
        let mut units = 0usize;
        for (count, (byte, ch)) in text.char_indices().enumerate() {
            if count % STRIDE == 0 {
                marks.push((units, byte));
            }
            units += ch.len_utf16();
        }
        Self { text, units, marks }
    }

    pub(crate) fn len(&self) -> usize {
        self.units
    }

    /// The last scalar boundary at or before `unit`, as (unit, byte).
    fn floor(&self, unit: usize) -> (usize, usize) {
        let unit = unit.min(self.units);
        if self.marks.is_empty() {
            return (unit, unit);
        }
        let mark = self.marks.partition_point(|&(at, _)| at <= unit) - 1;
        let (mut at, mut byte) = self.marks[mark];
        for ch in self.text[byte..].chars() {
            let next = at + ch.len_utf16();
            if next > unit {
                break;
            }
            at = next;
            byte += ch.len_utf8();
        }
        (at, byte)
    }

    /// The first scalar boundary at or after `unit`, as a byte offset.
    fn ceil_byte(&self, unit: usize) -> usize {
        let (at, byte) = self.floor(unit);
        if at >= unit.min(self.units) {
            byte
        } else {
            byte + self.text[byte..].chars().next().map_or(0, char::len_utf8)
        }
    }

    /// The byte offset of `unit` when it falls on a scalar boundary.
    pub(crate) fn exact_byte(&self, unit: usize) -> Option<usize> {
        let (at, byte) = self.floor(unit);
        (at == unit).then_some(byte)
    }

    /// The scalars lying wholly inside `[start, end)` in UTF-16 units: a
    /// bound inside a surrogate pair drops that pair rather than splitting it.
    pub(crate) fn whole_scalars(&self, start: usize, end: usize) -> &'a str {
        let start = start.min(self.units);
        let end = end.max(start).min(self.units);
        let low = self.ceil_byte(start);
        let high = self.floor(end).1;
        if low >= high { "" } else { &self.text[low..high] }
    }
}

/// One lazily built [`Utf16Index`] per run of a paragraph, shared by all of
/// its lines.
pub(crate) struct RunTexts<'a> {
    cells: Vec<OnceCell<Utf16Index<'a>>>,
}

impl<'a> RunTexts<'a> {
    pub(crate) fn new(runs: usize) -> Self {
        Self { cells: (0..runs).map(|_| OnceCell::new()).collect() }
    }

    pub(crate) fn get(&self, run_index: usize, text: &'a str) -> &Utf16Index<'a> {
        self.cells[run_index].get_or_init(|| Utf16Index::new(text))
    }
}

#[cfg(test)]
pub(crate) mod generated {
    pub(crate) struct Rng(u64);

    impl Rng {
        pub(crate) fn new(seed: u64) -> Self {
            Self(seed.max(1))
        }

        pub(crate) fn below(&mut self, bound: usize) -> usize {
            self.0 ^= self.0 << 13;
            self.0 ^= self.0 >> 7;
            self.0 ^= self.0 << 17;
            (self.0 % bound.max(1) as u64) as usize
        }
    }

    const PIECES: &[&str] = &[
        "a", "Z", " ", "-", "é", "e\u{301}", "\u{1F600}", "\u{1F44D}\u{1F3FD}", "\u{05D0}\u{05D1}",
        "\u{0627}\u{0644}", "\u{4E2D}", "\u{1F468}\u{200D}\u{1F469}", "\u{200F}",
    ];

    pub(crate) fn text(rng: &mut Rng) -> String {
        let ascii = rng.below(4) == 0;
        (0..rng.below(48))
            .map(|_| PIECES[rng.below(if ascii { 4 } else { PIECES.len() })])
            .collect()
    }

    /// Runs as JSON: text (some empty, some very long), tabs and line breaks.
    pub(crate) fn runs(rng: &mut Rng) -> Vec<serde_json::Value> {
        (0..1 + rng.below(6))
            .map(|_| match rng.below(10) {
                0 => serde_json::json!({ "kind": "tab", "pmStart": 3 }),
                1 => serde_json::json!({ "kind": "lineBreak", "pmStart": 5 }),
                choice => {
                    let text = if choice == 2 { text(rng).repeat(40) } else { text(rng) };
                    if rng.below(3) == 0 {
                        serde_json::json!({ "kind": "text", "text": text })
                    } else {
                        serde_json::json!({ "kind": "text", "text": text, "pmStart": 100 + rng.below(50) })
                    }
                }
            })
            .collect()
    }

    /// Consecutive lines `(head_run, head_char, tail_run, tail_char)` cut at
    /// random UTF-16 positions, inside surrogate pairs included, with empty
    /// lines and offsets past a run's end mixed in.
    pub(crate) fn lines(rng: &mut Rng, units: &[usize]) -> Vec<(usize, usize, usize, usize)> {
        let positions: Vec<(usize, usize)> = units
            .iter()
            .enumerate()
            .flat_map(|(run, &length)| (0..=length.max(1)).map(move |unit| (run, unit)))
            .collect();
        let mut cuts: Vec<(usize, usize)> =
            (0..2 + rng.below(12)).map(|_| positions[rng.below(positions.len())]).collect();
        cuts.push((0, 0));
        cuts.push(*positions.last().expect("a run"));
        cuts.sort();
        cuts.windows(2)
            .map(|pair| {
                let overshoot = if rng.below(8) == 0 { 3 } else { 0 };
                (pair[0].0, pair[0].1, pair[1].0, pair[1].1 + overshoot)
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::generated::{Rng, text};
    use super::*;

    fn old_whole_scalars(text: &str, start: usize, end: usize) -> String {
        let total = text.encode_utf16().count();
        let start = start.min(total);
        let end = end.max(start).min(total);
        let mut units = 0usize;
        text.chars()
            .filter_map(|ch| {
                let next = units + ch.len_utf16();
                let keep = units >= start && next <= end;
                units = next;
                keep.then_some(ch)
            })
            .collect()
    }

    fn old_exact_byte(text: &str, unit: usize) -> Option<usize> {
        let mut units = 0usize;
        for (byte, ch) in text.char_indices() {
            if units == unit {
                return Some(byte);
            }
            units += ch.len_utf16();
        }
        (units == unit).then_some(text.len())
    }

    #[test]
    fn slices_match_a_scan_from_the_start_at_every_offset() {
        let mut rng = Rng::new(0x5eed);
        for _ in 0..400 {
            let text = text(&mut rng).repeat(1 + rng.below(3));
            let index = Utf16Index::new(&text);
            let total = text.encode_utf16().count();
            assert_eq!(index.len(), total);
            for start in 0..=total + 2 {
                assert_eq!(index.exact_byte(start), old_exact_byte(&text, start), "{text:?} at {start}");
                for end in [0, start.saturating_sub(1), start, start + 1, start + 2, total / 2, total, total + 3] {
                    assert_eq!(
                        index.whole_scalars(start, end),
                        old_whole_scalars(&text, start, end),
                        "{text:?} [{start}, {end})"
                    );
                }
            }
        }
    }
}
