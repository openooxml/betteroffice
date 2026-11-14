//! Sliced CFF faces: a skeleton and outline chunks that reassemble the face.
//!
//! The skeleton (`scripts/fonts/cjk-slices.py`) holds every table of the face
//! byte for byte except `CFF `, whose CharStrings INDEX is its last structure
//! and is cut off, plus a glyph id to chunk table and a codepoint to chunk
//! table. A chunk holds `(glyph id, charstring)` pairs. Appending a CharStrings
//! INDEX built from any set of chunks (`endchar` for every other glyph) yields
//! a complete face whose layout tables equal the original's.

use alloc::vec;
use alloc::vec::Vec;
use core::fmt;

/// Container magic of a skeleton file.
pub const SKELETON_MAGIC: [u8; 4] = *b"BOSK";
/// Container format this reader understands.
pub const SLICES_FORMAT: u16 = 1;
/// Chunk id of a glyph or codepoint no chunk holds.
pub const NO_CHUNK: u16 = 0xFFFF;

const ENDCHAR: &[u8] = &[0x0e];
const HEADER: usize = 32;
const RUN: usize = 12;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SliceError {
    /// The bytes end before a structure they declare.
    Truncated,
    /// Not a skeleton container of a supported format.
    Format,
    /// A chunk names a glyph id the skeleton does not have.
    Glyph,
    /// The reassembled face would not fit the sfnt offsets.
    TooLarge,
}

impl fmt::Display for SliceError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Truncated => "truncated font slice",
            Self::Format => "not a supported font slice skeleton",
            Self::Glyph => "font slice chunk names an unknown glyph",
            Self::TooLarge => "reassembled font is too large",
        })
    }
}

impl core::error::Error for SliceError {}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, SliceError> {
    let value = bytes.get(offset..offset + 2).ok_or(SliceError::Truncated)?;
    Ok(u16::from_be_bytes([value[0], value[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, SliceError> {
    let value = bytes.get(offset..offset + 4).ok_or(SliceError::Truncated)?;
    Ok(u32::from_be_bytes([value[0], value[1], value[2], value[3]]))
}

fn usize_at(bytes: &[u8], offset: usize) -> Result<usize, SliceError> {
    usize::try_from(u32_at(bytes, offset)?).map_err(|_| SliceError::TooLarge)
}

/// A parsed skeleton container.
#[derive(Clone, Copy, Debug)]
pub struct Skeleton<'a> {
    num_glyphs: usize,
    cff_record: usize,
    cff_offset: usize,
    head_offset: usize,
    gid_map: &'a [u8],
    runs: &'a [u8],
    prefix: &'a [u8],
}

impl<'a> Skeleton<'a> {
    pub fn parse(bytes: &'a [u8]) -> Result<Self, SliceError> {
        if bytes.get(..4) != Some(&SKELETON_MAGIC[..]) || u16_at(bytes, 4)? != SLICES_FORMAT {
            return Err(SliceError::Format);
        }
        let num_glyphs = usize_at(bytes, 8)?;
        let cff_record = usize_at(bytes, 12)?;
        let cff_offset = usize_at(bytes, 16)?;
        let head_offset = usize_at(bytes, 20)?;
        let prefix_length = usize_at(bytes, 24)?;
        let run_count = usize_at(bytes, 28)?;
        let gid_end = num_glyphs
            .checked_mul(2)
            .and_then(|length| length.checked_add(HEADER))
            .ok_or(SliceError::Truncated)?;
        let runs_end = run_count
            .checked_mul(RUN)
            .and_then(|length| length.checked_add(gid_end))
            .ok_or(SliceError::Truncated)?;
        let prefix_start = bytes
            .len()
            .checked_sub(prefix_length)
            .ok_or(SliceError::Truncated)?;
        let fits = |offset: usize, length: usize| {
            offset
                .checked_add(length)
                .is_some_and(|end| end <= prefix_length)
        };
        if runs_end != prefix_start
            || num_glyphs > usize::from(u16::MAX)
            || !fits(cff_record, 16)
            || !fits(cff_offset, 0)
            || !fits(head_offset, 12)
        {
            return Err(SliceError::Format);
        }
        Ok(Self {
            num_glyphs,
            cff_record,
            cff_offset,
            head_offset,
            gid_map: &bytes[HEADER..gid_end],
            runs: &bytes[gid_end..runs_end],
            prefix: &bytes[prefix_start..],
        })
    }

    pub fn num_glyphs(&self) -> usize {
        self.num_glyphs
    }

    /// The chunk that holds `glyph`'s outline (the lowest one when several do).
    pub fn chunk_of_glyph(&self, glyph: u32) -> Option<u16> {
        let index = usize::try_from(glyph).ok()?;
        if index >= self.num_glyphs {
            return None;
        }
        let chunk = u16::from_be_bytes([self.gid_map[index * 2], self.gid_map[index * 2 + 1]]);
        (chunk != NO_CHUNK).then_some(chunk)
    }

    /// The chunk a codepoint's glyphs are cut into.
    pub fn chunk_of_codepoint(&self, codepoint: u32) -> Option<u16> {
        let (mut low, mut high) = (0, self.runs.len() / RUN);
        while low < high {
            let middle = (low + high) / 2;
            let run = &self.runs[middle * RUN..middle * RUN + RUN];
            let start = u32::from_be_bytes([run[0], run[1], run[2], run[3]]);
            let count = u32::from_be_bytes([run[4], run[5], run[6], run[7]]);
            if codepoint < start {
                high = middle;
            } else if codepoint - start >= count {
                low = middle + 1;
            } else {
                return Some(u16::from_be_bytes([run[8], run[9]]));
            }
        }
        None
    }

    /// Reassembles the face with the outlines of `chunks`; every other glyph is `endchar`.
    pub fn reassemble(&self, chunks: &[&[u8]]) -> Result<Vec<u8>, SliceError> {
        let mut outlines: Vec<Option<&[u8]>> = vec![None; self.num_glyphs];
        for chunk in chunks {
            for (glyph, outline) in chunk_entries(chunk)? {
                let slot = outlines
                    .get_mut(usize::try_from(glyph).map_err(|_| SliceError::Glyph)?)
                    .ok_or(SliceError::Glyph)?;
                *slot = Some(outline);
            }
        }
        let data_length = outlines.iter().try_fold(0usize, |total, outline| {
            total.checked_add(outline.unwrap_or(ENDCHAR).len())
        });
        let capacity = data_length
            .and_then(|length| length.checked_add(3))
            .and_then(|length| length.checked_add((self.num_glyphs + 1) * 4))
            .and_then(|length| length.checked_add(self.prefix.len() + 3))
            .filter(|&length| length <= crate::MAX_OUTPUT_LIMIT)
            .ok_or(SliceError::TooLarge)?;
        let mut font = Vec::new();
        font.try_reserve_exact(capacity)
            .map_err(|_| SliceError::TooLarge)?;
        font.extend_from_slice(self.prefix);
        let count = u16::try_from(self.num_glyphs).map_err(|_| SliceError::TooLarge)?;
        font.extend_from_slice(&count.to_be_bytes());
        font.push(4);
        let mut offset = 1u32;
        font.extend_from_slice(&offset.to_be_bytes());
        for outline in &outlines {
            let length = u32::try_from(outline.unwrap_or(ENDCHAR).len())
                .map_err(|_| SliceError::TooLarge)?;
            offset = offset.checked_add(length).ok_or(SliceError::TooLarge)?;
            font.extend_from_slice(&offset.to_be_bytes());
        }
        for outline in &outlines {
            font.extend_from_slice(outline.unwrap_or(ENDCHAR));
        }
        let cff_length = font.len() - self.cff_offset;
        let cff_sum = checksum(&font[self.cff_offset..]);
        font.resize(font.len() + (4 - font.len() % 4) % 4, 0);
        let cff_length = u32::try_from(cff_length).map_err(|_| SliceError::TooLarge)?;
        font[self.cff_record + 4..self.cff_record + 8].copy_from_slice(&cff_sum.to_be_bytes());
        font[self.cff_record + 12..self.cff_record + 16].copy_from_slice(&cff_length.to_be_bytes());
        font[self.head_offset + 8..self.head_offset + 12].fill(0);
        let adjustment = 0xB1B0_AFBAu32.wrapping_sub(checksum(&font));
        font[self.head_offset + 8..self.head_offset + 12]
            .copy_from_slice(&adjustment.to_be_bytes());
        Ok(font)
    }
}

/// `(glyph id, charstring)` pairs of a chunk file.
pub fn chunk_entries(chunk: &[u8]) -> Result<Vec<(u32, &[u8])>, SliceError> {
    let count = usize_at(chunk, 0)?;
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(count.min(chunk.len() / 8))
        .map_err(|_| SliceError::TooLarge)?;
    let mut position = 4usize;
    for _ in 0..count {
        let glyph = u32_at(chunk, position)?;
        let length = usize_at(chunk, position + 4)?;
        let start = position + 8;
        let end = start.checked_add(length).ok_or(SliceError::Truncated)?;
        entries.push((glyph, chunk.get(start..end).ok_or(SliceError::Truncated)?));
        position = end;
    }
    if position != chunk.len() {
        return Err(SliceError::Format);
    }
    Ok(entries)
}

fn checksum(bytes: &[u8]) -> u32 {
    let mut sum = 0u32;
    for word in bytes.chunks(4) {
        let mut padded = [0u8; 4];
        padded[..word.len()].copy_from_slice(word);
        sum = sum.wrapping_add(u32::from_be_bytes(padded));
    }
    sum
}

#[cfg(test)]
mod tests {
    use super::*;

    fn container(
        num_glyphs: u16,
        gid_map: &[u16],
        runs: &[(u32, u32, u16)],
        prefix: &[u8],
    ) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&SKELETON_MAGIC);
        bytes.extend_from_slice(&SLICES_FORMAT.to_be_bytes());
        bytes.extend_from_slice(&0u16.to_be_bytes());
        for value in [
            u32::from(num_glyphs),
            12,
            44,
            28,
            prefix.len() as u32,
            runs.len() as u32,
        ] {
            bytes.extend_from_slice(&value.to_be_bytes());
        }
        for chunk in gid_map {
            bytes.extend_from_slice(&chunk.to_be_bytes());
        }
        for (start, count, chunk) in runs {
            bytes.extend_from_slice(&start.to_be_bytes());
            bytes.extend_from_slice(&count.to_be_bytes());
            bytes.extend_from_slice(&chunk.to_be_bytes());
            bytes.extend_from_slice(&0u16.to_be_bytes());
        }
        bytes.extend_from_slice(prefix);
        bytes
    }

    fn chunk(entries: &[(u32, &[u8])]) -> Vec<u8> {
        let mut bytes = (entries.len() as u32).to_be_bytes().to_vec();
        for (glyph, outline) in entries {
            bytes.extend_from_slice(&glyph.to_be_bytes());
            bytes.extend_from_slice(&(outline.len() as u32).to_be_bytes());
            bytes.extend_from_slice(outline);
        }
        bytes
    }

    fn prefix() -> Vec<u8> {
        let mut prefix = vec![0u8; 12];
        prefix.extend_from_slice(b"CFF \0\0\0\0\0\0\0\x1c\0\0\0\0");
        prefix.extend_from_slice(&[0u8; 16]);
        prefix.extend_from_slice(&[0xAAu8; 8]);
        prefix
    }

    #[test]
    fn maps_glyphs_and_codepoints_to_chunks() {
        let bytes = container(
            3,
            &[0, 2, NO_CHUNK],
            &[(0x41, 3, 0), (0x3042, 1, 5)],
            &prefix(),
        );
        let skeleton = Skeleton::parse(&bytes).unwrap();
        assert_eq!(skeleton.chunk_of_glyph(1), Some(2));
        assert_eq!(skeleton.chunk_of_glyph(2), None);
        assert_eq!(skeleton.chunk_of_glyph(3), None);
        assert_eq!(skeleton.chunk_of_codepoint(0x43), Some(0));
        assert_eq!(skeleton.chunk_of_codepoint(0x44), None);
        assert_eq!(skeleton.chunk_of_codepoint(0x3042), Some(5));
    }

    #[test]
    fn reassembly_appends_an_index_with_endchar_for_missing_glyphs() {
        let bytes = container(3, &[0, 0, 0], &[], &prefix());
        let skeleton = Skeleton::parse(&bytes).unwrap();
        let first = chunk(&[(1, &[1, 2, 3])]);
        let font = skeleton.reassemble(&[&first]).unwrap();
        let index = &font[prefix().len()..];
        assert_eq!(&index[..3], &[0, 3, 4]);
        let offsets: Vec<u32> = (0..4)
            .map(|i| u32::from_be_bytes(index[3 + i * 4..7 + i * 4].try_into().unwrap()))
            .collect();
        assert_eq!(offsets, [1, 2, 5, 6]);
        assert_eq!(&index[19..24], &[0x0e, 1, 2, 3, 0x0e]);
        assert_eq!(font.len() % 4, 0);
        let length = u32::from_be_bytes(font[24..28].try_into().unwrap());
        assert_eq!(length, 8 + 24);
        assert_eq!(checksum(&font), 0xB1B0_AFBA);
    }

    #[test]
    fn rejects_offsets_that_overflow_on_32_bit_targets() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&SKELETON_MAGIC);
        bytes.extend_from_slice(&SLICES_FORMAT.to_be_bytes());
        bytes.extend_from_slice(&0u16.to_be_bytes());
        for value in [1u32, 0xffff_fff0, 0, 0, 16, 0] {
            bytes.extend_from_slice(&value.to_be_bytes());
        }
        bytes.extend_from_slice(&[0u8; 2 + 16]);
        assert_eq!(Skeleton::parse(&bytes).unwrap_err(), SliceError::Format);
        bytes[12..16].copy_from_slice(&0u32.to_be_bytes());
        bytes[20..24].copy_from_slice(&0xffff_fffcu32.to_be_bytes());
        assert_eq!(Skeleton::parse(&bytes).unwrap_err(), SliceError::Format);
    }

    #[test]
    fn rejects_overlapping_chunk_views_whose_outlines_exceed_the_output_limit() {
        let glyphs = 32_769u32;
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&SKELETON_MAGIC);
        bytes.extend_from_slice(&SLICES_FORMAT.to_be_bytes());
        bytes.extend_from_slice(&0u16.to_be_bytes());
        for value in [glyphs, 0, 0, 0, 16, 0] {
            bytes.extend_from_slice(&value.to_be_bytes());
        }
        bytes.extend_from_slice(&vec![0u8; 2 * glyphs as usize + 16]);
        let skeleton = Skeleton::parse(&bytes).unwrap();
        let mut entries = Vec::new();
        for glyph in 0..glyphs {
            entries.extend_from_slice(&1u32.to_be_bytes());
            entries.extend_from_slice(&glyph.to_be_bytes());
            entries.extend_from_slice(&(12 * (glyphs - glyph - 1)).to_be_bytes());
        }
        let chunks: Vec<&[u8]> = (0..glyphs as usize)
            .map(|glyph| &entries[12 * glyph..])
            .collect();
        assert_eq!(skeleton.reassemble(&chunks), Err(SliceError::TooLarge));
    }

    #[test]
    fn rejects_unknown_glyphs_and_bad_containers() {
        let bytes = container(2, &[0, 0], &[], &prefix());
        let skeleton = Skeleton::parse(&bytes).unwrap();
        assert_eq!(
            skeleton.reassemble(&[&chunk(&[(2, &[1])])]),
            Err(SliceError::Glyph)
        );
        let mut truncated = chunk(&[(0, &[1, 2])]);
        truncated.pop();
        assert_eq!(chunk_entries(&truncated), Err(SliceError::Truncated));
        assert_eq!(Skeleton::parse(b"BOSK").unwrap_err(), SliceError::Truncated);
        let mut wrong = bytes.clone();
        wrong[4] = 9;
        assert_eq!(Skeleton::parse(&wrong).unwrap_err(), SliceError::Format);
    }
}
