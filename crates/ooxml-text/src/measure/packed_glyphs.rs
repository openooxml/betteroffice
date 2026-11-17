//! The compact form of a measured line's glyph runs.
//!
//! Lines travel as JSON and a line of bidi text can hold dozens of runs, so a
//! line's runs are one unpadded base64 string of little-endian bytes: a `u32`
//! run count, then per run, with `v` a LEB128 `u32`,
//!
//! ```text
//! v run_index, v start_char, v length, v font_id, u8 bidi_level, u8 flags,
//! f32 x_offset, f32 advance, v glyph_count, then per glyph u16 id, f32 advance
//! ```
//!
//! Clusters are monotonic in visual order, so [`CLUSTERS`] adds one `u8` step
//! per glyph: the distance from the range's starting edge for the first glyph,
//! from the previous cluster for the rest. Without the flag the steps are 0
//! then 1. [`OFFSETS`] then adds a bitmap of the glyphs with placement offsets
//! and, for each of those, `f32` `dx` and `dy`.

use crate::bidi::level_is_rtl;

const CLUSTERS: u8 = 1;
const OFFSETS: u8 = 2;
const RECORD_LEN: usize = 6;
const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/// One shaped glyph in visual order. `cluster` is the UTF-16 offset in the run
/// of the cluster it belongs to. `advance` excludes letter spacing, auto-space
/// and justification, which the cluster advances include; `dx` and `dy` are
/// the shaper's placement offsets.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MeasuredGlyph {
    pub id: u32,
    pub cluster: u32,
    pub advance: f32,
    pub dx: f32,
    pub dy: f32,
}

/// One visual slice of a line shaped by a single call: one run, one face, one
/// bidi level. `start_char..end_char` is the slice's UTF-16 source range in the
/// run and `x_offset` its visual x from the line start, as for cluster
/// advances; `advance` is the sum of its cluster advances. `font_id` is the
/// face as `FontStore::canonical_face` names it, so repeated registrations of
/// one font or view yield the same runs.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GlyphRunHead {
    pub run_index: u32,
    pub start_char: u32,
    pub end_char: u32,
    pub font_id: u32,
    pub bidi_level: u8,
    pub x_offset: f32,
    pub advance: f32,
}

/// A decoded glyph run.
#[derive(Debug, Clone, PartialEq)]
pub struct MeasuredGlyphRun {
    pub head: GlyphRunHead,
    pub glyphs: Vec<MeasuredGlyph>,
}

/// Builds the packed runs of one line.
#[derive(Debug, Default)]
pub struct GlyphRunWriter {
    bytes: Vec<u8>,
    steps: Vec<u8>,
    count: u32,
}

impl GlyphRunWriter {
    /// Adds a run. A run the format cannot hold, for a glyph id past the glyph
    /// index range or clusters that cannot be stepped, is left out, and paint
    /// reshapes that slice.
    pub fn push(&mut self, head: GlyphRunHead, glyphs: &[MeasuredGlyph]) {
        if self.bytes.is_empty() {
            self.bytes.extend_from_slice(&[0; 4]);
        }
        let start = self.bytes.len();
        if self.write(head, glyphs).is_some() {
            self.count += 1;
        } else {
            self.bytes.truncate(start);
        }
    }

    fn write(&mut self, head: GlyphRunHead, glyphs: &[MeasuredGlyph]) -> Option<()> {
        let rtl = level_is_rtl(head.bidi_level);
        self.steps.clear();
        let mut previous: Option<u32> = None;
        for glyph in glyphs {
            let step = match (previous, rtl) {
                (None, false) => glyph.cluster.checked_sub(head.start_char)?,
                (None, true) => head.end_char.checked_sub(1)?.checked_sub(glyph.cluster)?,
                (Some(previous), false) => glyph.cluster.checked_sub(previous)?,
                (Some(previous), true) => previous.checked_sub(glyph.cluster)?,
            };
            self.steps.push(u8::try_from(step).ok()?);
            previous = Some(glyph.cluster);
        }
        let implied = self
            .steps
            .iter()
            .enumerate()
            .all(|(index, &step)| step == u8::from(index > 0));
        let placed = |glyph: &MeasuredGlyph| glyph.dx != 0.0 || glyph.dy != 0.0;
        let offsets = glyphs.iter().any(placed);

        let bytes = &mut self.bytes;
        write_varint(bytes, head.run_index);
        write_varint(bytes, head.start_char);
        write_varint(bytes, head.end_char.checked_sub(head.start_char)?);
        write_varint(bytes, head.font_id);
        bytes.push(head.bidi_level);
        bytes.push((if implied { 0 } else { CLUSTERS }) | (if offsets { OFFSETS } else { 0 }));
        bytes.extend_from_slice(&head.x_offset.to_le_bytes());
        bytes.extend_from_slice(&head.advance.to_le_bytes());
        write_varint(bytes, u32::try_from(glyphs.len()).ok()?);
        for glyph in glyphs {
            bytes.extend_from_slice(&u16::try_from(glyph.id).ok()?.to_le_bytes());
            bytes.extend_from_slice(&glyph.advance.to_le_bytes());
        }
        if !implied {
            bytes.extend_from_slice(&self.steps);
        }
        if offsets {
            let bitmap_start = bytes.len();
            bytes.resize(bitmap_start + glyphs.len().div_ceil(8), 0);
            for (index, glyph) in glyphs.iter().enumerate() {
                if placed(glyph) {
                    bytes[bitmap_start + index / 8] |= 1 << (index % 8);
                }
            }
            for glyph in glyphs.iter().filter(|glyph| placed(glyph)) {
                bytes.extend_from_slice(&glyph.dx.to_le_bytes());
                bytes.extend_from_slice(&glyph.dy.to_le_bytes());
            }
        }
        Some(())
    }

    /// The packed line. `None` when it holds no run.
    pub fn finish(mut self) -> Option<String> {
        if self.count == 0 {
            return None;
        }
        self.bytes[..4].copy_from_slice(&self.count.to_le_bytes());
        Some(encode(&self.bytes))
    }
}

/// Packs a line's runs, as [`GlyphRunWriter`] does.
pub fn pack_glyph_runs(runs: &[MeasuredGlyphRun]) -> Option<String> {
    let mut writer = GlyphRunWriter::default();
    for run in runs {
        writer.push(run.head, &run.glyphs);
    }
    writer.finish()
}

/// The runs [`GlyphRunWriter`] packed. `None` when `packed` is not such a value.
pub fn unpack_glyph_runs(packed: &str) -> Option<Vec<MeasuredGlyphRun>> {
    let bytes = decode(packed)?;
    let mut reader = Reader(&bytes);
    let count = u32::from_le_bytes(reader.array()?);
    let mut runs = Vec::new();
    for _ in 0..count {
        runs.push(read_run(&mut reader)?);
    }
    reader.0.is_empty().then_some(runs)
}

fn read_run(reader: &mut Reader) -> Option<MeasuredGlyphRun> {
    let run_index = reader.varint()?;
    let start_char = reader.varint()?;
    let end_char = start_char.checked_add(reader.varint()?)?;
    let font_id = reader.varint()?;
    let [bidi_level, flags] = reader.array()?;
    if flags & !(CLUSTERS | OFFSETS) != 0 {
        return None;
    }
    let x_offset = f32::from_le_bytes(reader.array()?);
    let advance = f32::from_le_bytes(reader.array()?);
    let count = reader.varint()? as usize;
    let records = reader.take(count.checked_mul(RECORD_LEN)?)?;
    let steps = if flags & CLUSTERS != 0 {
        Some(reader.take(count)?)
    } else {
        None
    };
    let bitmap = if flags & OFFSETS != 0 {
        Some(reader.take(count.div_ceil(8))?)
    } else {
        None
    };
    let rtl = level_is_rtl(bidi_level);
    let mut cluster = if rtl {
        end_char.checked_sub(1)
    } else {
        Some(start_char)
    };
    let mut glyphs = Vec::with_capacity(count);
    let (records, _) = records.as_chunks::<RECORD_LEN>();
    for (index, record) in records.iter().enumerate() {
        let step = steps.map_or(u32::from(index > 0), |steps| u32::from(steps[index]));
        cluster = if rtl {
            cluster?.checked_sub(step)
        } else {
            cluster?.checked_add(step)
        };
        let placed = bitmap.is_some_and(|bitmap| bitmap[index / 8] & (1 << (index % 8)) != 0);
        let (dx, dy) = if placed {
            (
                f32::from_le_bytes(reader.array()?),
                f32::from_le_bytes(reader.array()?),
            )
        } else {
            (0.0, 0.0)
        };
        glyphs.push(MeasuredGlyph {
            id: u32::from(u16::from_le_bytes(record[..2].try_into().ok()?)),
            cluster: cluster?,
            advance: f32::from_le_bytes(record[2..].try_into().ok()?),
            dx,
            dy,
        });
    }
    Some(MeasuredGlyphRun {
        head: GlyphRunHead {
            run_index,
            start_char,
            end_char,
            font_id,
            bidi_level,
            x_offset,
            advance,
        },
        glyphs,
    })
}

struct Reader<'a>(&'a [u8]);

impl<'a> Reader<'a> {
    fn take(&mut self, length: usize) -> Option<&'a [u8]> {
        let (taken, rest) = self.0.split_at_checked(length)?;
        self.0 = rest;
        Some(taken)
    }

    fn array<const N: usize>(&mut self) -> Option<[u8; N]> {
        self.take(N)?.try_into().ok()
    }

    fn varint(&mut self) -> Option<u32> {
        let mut value = 0u32;
        for shift in (0..35).step_by(7) {
            let [byte] = self.array()?;
            let bits = u32::from(byte & 0x7f);
            if shift == 28 && bits > 0x0f {
                return None;
            }
            value |= bits << shift;
            if byte & 0x80 == 0 {
                return Some(value);
            }
        }
        None
    }
}

fn write_varint(bytes: &mut Vec<u8>, mut value: u32) {
    while value >= 0x80 {
        bytes.push(value as u8 | 0x80);
        value >>= 7;
    }
    bytes.push(value as u8);
}

fn encode(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len().div_ceil(3) * 4);
    for chunk in bytes.chunks(3) {
        let group = (u32::from(chunk[0]) << 16)
            | (u32::from(chunk.get(1).copied().unwrap_or(0)) << 8)
            | u32::from(chunk.get(2).copied().unwrap_or(0));
        for position in 0..=chunk.len() {
            out.push(char::from(
                ALPHABET[(group >> (18 - position * 6)) as usize & 63],
            ));
        }
    }
    out
}

fn sextet(byte: u8) -> Option<u32> {
    Some(u32::from(match byte {
        b'A'..=b'Z' => byte - b'A',
        b'a'..=b'z' => byte - b'a' + 26,
        b'0'..=b'9' => byte - b'0' + 52,
        b'+' => 62,
        b'/' => 63,
        _ => return None,
    }))
}

fn decode(text: &str) -> Option<Vec<u8>> {
    let mut out = Vec::with_capacity(text.len() / 4 * 3 + 2);
    for chunk in text.as_bytes().chunks(4) {
        if chunk.len() < 2 {
            return None;
        }
        let mut group = 0;
        for position in 0..4 {
            let value = chunk.get(position).map_or(Some(0), |&byte| sextet(byte))?;
            group = (group << 6) | value;
        }
        out.extend_from_slice(&group.to_be_bytes()[1..chunk.len()]);
    }
    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn glyph(id: u32, cluster: u32) -> MeasuredGlyph {
        MeasuredGlyph {
            id,
            cluster,
            advance: 6.25,
            dx: 0.0,
            dy: 0.0,
        }
    }

    fn head(start_char: u32, end_char: u32, bidi_level: u8) -> GlyphRunHead {
        GlyphRunHead {
            run_index: 2,
            start_char,
            end_char,
            font_id: 7,
            bidi_level,
            x_offset: 12.5,
            advance: 25.0,
        }
    }

    fn encoded_len(bytes: usize) -> usize {
        bytes / 3 * 4 + [0, 2, 3][bytes % 3]
    }

    fn round_trip(runs: Vec<MeasuredGlyphRun>) -> String {
        let packed = pack_glyph_runs(&runs).unwrap();
        assert_eq!(unpack_glyph_runs(&packed).unwrap(), runs);
        packed
    }

    fn run(head: GlyphRunHead, glyphs: Vec<MeasuredGlyph>) -> MeasuredGlyphRun {
        MeasuredGlyphRun { head, glyphs }
    }

    #[test]
    fn simple_text_costs_six_bytes_a_glyph_and_a_short_header() {
        let glyphs: Vec<_> = (0..90).map(|index| glyph(60 + index, 10 + index)).collect();
        let packed = round_trip(vec![run(head(10, 100, 0), glyphs)]);
        let run_header = 4 + 2 + 8 + 1;
        assert_eq!(packed.len(), encoded_len(4 + run_header + 90 * RECORD_LEN));
    }

    #[test]
    fn right_to_left_clusters_descend_from_the_range_end() {
        let glyphs = vec![glyph(5, 3), glyph(6, 2), glyph(7, 1), glyph(8, 0)];
        let packed = round_trip(vec![run(head(0, 4, 1), glyphs)]);
        assert_eq!(packed.len(), encoded_len(4 + 15 + 4 * RECORD_LEN));
    }

    #[test]
    fn ligatures_and_shared_clusters_cost_one_step_a_glyph() {
        let glyphs = vec![glyph(4, 0), glyph(9, 2), glyph(10, 2), glyph(5, 3)];
        let packed = round_trip(vec![run(head(0, 4, 0), glyphs)]);
        assert_eq!(packed.len(), encoded_len(4 + 15 + 4 * (RECORD_LEN + 1)));
    }

    #[test]
    fn right_to_left_slices_after_a_line_cut_keep_their_clusters() {
        let glyphs = vec![glyph(5, 19), glyph(6, 19), glyph(7, 17), glyph(8, 12)];
        round_trip(vec![run(head(12, 21, 1), glyphs)]);
    }

    #[test]
    fn only_placed_glyphs_store_offsets() {
        let mark = MeasuredGlyph {
            dx: -1.5,
            dy: 2.25,
            ..glyph(9, 1)
        };
        let glyphs: Vec<_> = (0..9)
            .map(|index| {
                if index == 1 {
                    mark
                } else {
                    glyph(index, index)
                }
            })
            .collect();
        let packed = round_trip(vec![run(head(0, 9, 0), glyphs)]);
        assert_eq!(packed.len(), encoded_len(4 + 15 + 9 * RECORD_LEN + 2 + 8));
    }

    #[test]
    fn a_line_holds_many_runs_of_mixed_direction() {
        let runs: Vec<_> = (0..40)
            .map(|index| {
                let start = index * 5;
                let level = (index % 2) as u8;
                let glyphs = if level == 1 {
                    vec![glyph(1, start + 4), glyph(2, start + 2), glyph(3, start)]
                } else {
                    vec![glyph(1, start), glyph(2, start + 1), glyph(3, start + 4)]
                };
                run(head(start, start + 5, level), glyphs)
            })
            .collect();
        round_trip(runs);
    }

    #[test]
    fn large_values_fit_their_varints() {
        let mut large = head(u32::MAX - 1, u32::MAX, 0);
        large.run_index = u32::MAX;
        large.font_id = 1 << 28;
        round_trip(vec![run(large, vec![glyph(1, u32::MAX - 1)])]);
    }

    #[test]
    fn a_run_the_format_cannot_hold_leaves_the_others() {
        let kept = run(head(0, 1, 0), vec![glyph(1, 0)]);
        let far = run(head(0, 301, 0), vec![glyph(1, 0), glyph(2, 300)]);
        let backwards = run(head(0, 4, 0), vec![glyph(1, 3), glyph(2, 1)]);
        let wide = run(head(0, 1, 0), vec![glyph(70_000, 0)]);
        let packed = pack_glyph_runs(&[far, kept.clone(), backwards, wide]).unwrap();
        assert_eq!(unpack_glyph_runs(&packed).unwrap(), vec![kept]);
    }

    #[test]
    fn a_line_with_no_run_packs_to_nothing() {
        assert_eq!(pack_glyph_runs(&[]), None);
        let wide = run(head(0, 1, 0), vec![glyph(70_000, 0)]);
        assert_eq!(pack_glyph_runs(&[wide]), None);
    }

    #[test]
    fn empty_glyph_lists_round_trip() {
        round_trip(vec![run(head(3, 3, 0), vec![])]);
    }

    #[test]
    fn foreign_values_do_not_unpack() {
        for value in ["", "A", "!!!!", "AAAA", "/////w"] {
            assert_eq!(unpack_glyph_runs(value), None, "{value}");
        }
    }

    #[test]
    fn truncated_and_trailing_bytes_do_not_unpack() {
        let glyphs = vec![glyph(1, 0), glyph(2, 2)];
        let packed = pack_glyph_runs(&[run(head(0, 3, 0), glyphs)]).unwrap();
        assert_eq!(unpack_glyph_runs(&format!("{packed}AAAA")), None);
        for length in 0..packed.len() {
            assert_eq!(unpack_glyph_runs(&packed[..length]), None, "{length}");
        }
    }
}
