use crate::DecodeError;

type Result<T> = core::result::Result<T, DecodeError>;

struct Reader<'a>(&'a [u8]);

impl<'a> Reader<'a> {
    fn take(&mut self, len: usize) -> Result<&'a [u8]> {
        let bytes = self.0.get(..len).ok_or(DecodeError::Woff2Reconstruction)?;
        self.0 = &self.0[len..];
        Ok(bytes)
    }

    fn u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    fn u16(&mut self) -> Result<usize> {
        let bytes = self.take(2)?;
        Ok(u16::from_be_bytes([bytes[0], bytes[1]]) as usize)
    }

    fn u32(&mut self) -> Result<usize> {
        let bytes = self.take(4)?;
        Ok(u32::from_be_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize)
    }

    fn base128(&mut self) -> Result<usize> {
        let mut value = 0u32;
        for index in 0..5 {
            let byte = self.u8()?;
            if (index == 0 && byte == 0x80) || value & 0xfe000000 != 0 {
                return Err(DecodeError::Woff2Reconstruction);
            }
            value = (value << 7) | u32::from(byte & 127);
            if byte & 128 == 0 {
                return Ok(value as usize);
            }
        }
        Err(DecodeError::Woff2Reconstruction)
    }

    fn ushort255(&mut self) -> Result<usize> {
        match self.u8()? {
            253 => self.u16(),
            254 => Ok(506 + usize::from(self.u8()?)),
            255 => Ok(253 + usize::from(self.u8()?)),
            value => Ok(usize::from(value)),
        }
    }
}

fn add(total: &mut usize, size: usize, limit: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .filter(|&value| value <= limit)
        .ok_or(DecodeError::Woff2Reconstruction)?;
    Ok(())
}

fn padded(size: usize) -> Result<usize> {
    size.checked_add(3)
        .map(|value| value & !3)
        .ok_or(DecodeError::Woff2Reconstruction)
}

pub(super) fn validate(payload: &[u8], max_output: usize) -> Result<()> {
    let mut input = Reader(payload);
    if input.take(8)? != b"wOF2\0\x01\0\0" || input.u32()? != payload.len() {
        return Err(DecodeError::Woff2Reconstruction);
    }
    let tables = input.u16()?;
    input.u16()?;
    let sfnt_size = input.u32()?;
    let stream_size = input.u32()?;
    if !(1..=256).contains(&tables)
        || sfnt_size < 12 + 16 * tables
        || sfnt_size > max_output
        || stream_size > payload.len()
    {
        return Err(DecodeError::Woff2Reconstruction);
    }
    input.take(4)?;
    let meta_offset = input.u32()?;
    let meta_length = input.u32()?;
    input.u32()?;
    let priv_offset = input.u32()?;
    let priv_length = input.u32()?;
    for (offset, length) in [(meta_offset, meta_length), (priv_offset, priv_length)] {
        let mut end = offset;
        add(&mut end, length, payload.len())?;
    }

    let mut original_size = 12 + 16 * tables;
    let mut other_size = original_size;
    let mut stored_size = 0;
    let mut glyf = None;
    let mut loca = None;
    for _ in 0..tables {
        let flags = input.u8()?;
        let tag = match flags & 63 {
            10 => b"glyf".as_slice(),
            11 => b"loca".as_slice(),
            63 => input.take(4)?,
            _ => b"".as_slice(),
        };
        let is_glyf = tag == b"glyf";
        let is_loca = tag == b"loca";
        if flags >> 6 != 0 {
            return Err(DecodeError::Woff2Reconstruction);
        }
        let original = input.base128()?;
        let stored = if is_glyf || is_loca {
            input.base128()?
        } else {
            original
        };
        if original > payload.len().saturating_mul(8) || stored > payload.len() {
            return Err(DecodeError::Woff2Reconstruction);
        }
        add(&mut original_size, padded(original)?, max_output)?;
        if is_glyf {
            if glyf.replace((original, stored_size, stored)).is_some() {
                return Err(DecodeError::Woff2Reconstruction);
            }
        } else if is_loca {
            if stored != 0 || loca.replace(original).is_some() {
                return Err(DecodeError::Woff2Reconstruction);
            }
        } else {
            add(&mut other_size, padded(original)?, max_output)?;
        }
        add(&mut stored_size, stored, stream_size)?;
    }
    if stored_size != stream_size {
        return Err(DecodeError::Woff2Reconstruction);
    }
    let stream = input.take(stream_size)?;
    let (original, offset, length) = glyf.ok_or(DecodeError::Woff2Reconstruction)?;
    let mut glyph_data = Reader(
        stream
            .get(offset..)
            .ok_or(DecodeError::Woff2Reconstruction)?,
    );
    let (glyf_size, loca_size) = scan_glyphs(glyph_data.take(length)?, original, max_output)?;
    if Some(loca_size) != loca {
        return Err(DecodeError::Woff2Reconstruction);
    }
    add(&mut other_size, glyf_size, max_output)?;
    add(&mut other_size, padded(loca_size)?, max_output)
}

fn scan_glyphs(data: &[u8], original: usize, max_output: usize) -> Result<(usize, usize)> {
    let mut input = Reader(data);
    let version = input.u16()?;
    let options = input.u16()?;
    let glyphs = input.u16()?;
    let index_format = input.u16()?;
    if version != 0 || options & !1 != 0 || index_format > 1 || glyphs * 12 > max_output {
        return Err(DecodeError::Woff2Reconstruction);
    }
    let mut lengths = [0; 7];
    for length in &mut lengths {
        *length = input.u32()?;
    }
    let mut contours = Reader(input.take(lengths[0])?);
    let mut points = Reader(input.take(lengths[1])?);
    let mut flags = Reader(input.take(lengths[2])?);
    let mut triplets = Reader(input.take(lengths[3])?);
    let mut composites = Reader(input.take(lengths[4])?);
    let mut bboxes = Reader(input.take(lengths[5])?);
    let mut instructions = Reader(input.take(lengths[6])?);
    let bbox_bits = bboxes.take(glyphs.div_ceil(32) * 4)?;
    let overlaps = if options & 1 != 0 {
        input.take(glyphs.div_ceil(8))?
    } else {
        &[]
    };
    if !input.0.is_empty() || contours.0.len() != glyphs * 2 {
        return Err(DecodeError::Woff2Reconstruction);
    }
    let mut total_points = 0;
    let mut output = 0;
    for glyph in 0..glyphs {
        let count = contours.u16()?;
        let bbox = bbox_bits[glyph / 8] & (128 >> (glyph % 8)) != 0;
        let size = match count {
            0 if !bbox => 0,
            65535 if bbox => {
                let before = composites.0.len();
                let mut have_instructions = false;
                loop {
                    let flags = composites.u16()?;
                    have_instructions |= flags & 256 != 0;
                    let args = if flags & 1 != 0 { 4 } else { 2 };
                    let transform = if flags & 8 != 0 {
                        2
                    } else if flags & 64 != 0 {
                        4
                    } else if flags & 128 != 0 {
                        8
                    } else {
                        0
                    };
                    composites.take(2 + args + transform)?;
                    if flags & 32 == 0 {
                        break;
                    }
                }
                let instruction_size = if have_instructions {
                    triplets.ushort255()?
                } else {
                    0
                };
                instructions.take(instruction_size)?;
                bboxes.take(8)?;
                let component_size = before - composites.0.len();
                if 12 + component_size + instruction_size > max_output {
                    return Err(DecodeError::Woff2Reconstruction);
                }
                10 + component_size
                    + if have_instructions {
                        2 + instruction_size
                    } else {
                        0
                    }
            }
            1..=32767 => {
                let mut glyph_points = 0;
                for _ in 0..count {
                    add(&mut glyph_points, points.ushort255()?, 65536)?;
                }
                add(&mut total_points, glyph_points, original)?;
                if glyph_points * 12 > max_output {
                    return Err(DecodeError::Woff2Reconstruction);
                }
                let overlap =
                    !overlaps.is_empty() && overlaps[glyph / 8] & (128 >> (glyph % 8)) != 0;
                let encoded = point_bytes(flags.take(glyph_points)?, &mut triplets, overlap)?;
                let instruction_size = triplets.ushort255()?;
                instructions.take(instruction_size)?;
                if bbox {
                    bboxes.take(8)?;
                }
                if 12 + count * 2 + glyph_points * 5 + instruction_size > max_output {
                    return Err(DecodeError::Woff2Reconstruction);
                }
                12 + count * 2 + encoded + instruction_size
            }
            _ => return Err(DecodeError::Woff2Reconstruction),
        };
        add(&mut output, padded(size)?, max_output)?;
    }
    if [
        points.0,
        flags.0,
        triplets.0,
        composites.0,
        bboxes.0,
        instructions.0,
    ]
    .iter()
    .any(|stream| !stream.is_empty())
    {
        return Err(DecodeError::Woff2Reconstruction);
    }
    Ok((output, (glyphs + 1) * if index_format == 0 { 2 } else { 4 }))
}

fn point_bytes(flags: &[u8], triplets: &mut Reader<'_>, overlap: bool) -> Result<usize> {
    let mut coordinates = 0;
    let mut encoded_flags = 0;
    let mut previous = u8::MAX;
    let mut repeat = 0;
    for (index, &code) in flags.iter().enumerate() {
        let flag = code & 127;
        let size = match flag {
            0..=83 => 1,
            84..=119 => 2,
            120..=123 => 3,
            _ => 4,
        };
        let bytes = triplets.take(size)?;
        let (x, y) = match flag {
            0..=9 => (0, (usize::from(flag & 14) << 7) + usize::from(bytes[0])),
            10..=19 => (
                (usize::from((flag - 10) & 14) << 7) + usize::from(bytes[0]),
                0,
            ),
            20..=83 => {
                let value = flag - 20;
                (
                    1 + usize::from(value & 48) + usize::from(bytes[0] >> 4),
                    1 + usize::from((value & 12) << 2) + usize::from(bytes[0] & 15),
                )
            }
            84..=119 => {
                let value = flag - 84;
                (
                    1 + (usize::from(value / 12) << 8) + usize::from(bytes[0]),
                    1 + (usize::from((value % 12) >> 2) << 8) + usize::from(bytes[1]),
                )
            }
            120..=123 => (
                (usize::from(bytes[0]) << 4) + usize::from(bytes[1] >> 4),
                (usize::from(bytes[1] & 15) << 8) + usize::from(bytes[2]),
            ),
            _ => (
                (usize::from(bytes[0]) << 8) + usize::from(bytes[1]),
                (usize::from(bytes[2]) << 8) + usize::from(bytes[3]),
            ),
        };
        let mut encoded = u8::from(code & 128 == 0);
        if overlap && index == 0 {
            encoded |= 64;
        }
        let y_sign = if flag < 10 { flag & 1 } else { (flag >> 1) & 1 };
        for (value, sign, short, same) in [(x, flag & 1, 2, 16), (y, y_sign, 4, 32)] {
            if value == 0 {
                encoded |= same;
            } else if value < 256 {
                coordinates += 1;
                encoded |= short;
                if sign != 0 {
                    encoded |= same;
                }
            } else {
                coordinates += 2;
            }
        }
        if encoded == previous && repeat < 255 {
            repeat += 1;
        } else {
            if previous != u8::MAX {
                encoded_flags += 1 + usize::from(repeat != 0);
            }
            repeat = 0;
        }
        previous = encoded;
    }
    if previous != u8::MAX {
        encoded_flags += 1 + usize::from(repeat != 0);
    }
    Ok(encoded_flags + coordinates)
}

#[cfg(test)]
mod tests {
    use alloc::vec;
    use alloc::vec::Vec;

    use super::{Reader, padded, point_bytes, scan_glyphs, validate};
    use crate::DecodeError;

    fn base128(mut value: u32) -> Vec<u8> {
        let mut bytes = vec![(value & 127) as u8];
        value >>= 7;
        while value != 0 {
            bytes.push((value & 127) as u8 | 128);
            value >>= 7;
        }
        bytes.reverse();
        bytes
    }

    fn directory(glyf: u32, loca: u32, stored: u32) -> Vec<u8> {
        let mut bytes = vec![10];
        bytes.extend(base128(glyf));
        bytes.extend(base128(stored));
        bytes.push(11);
        bytes.extend(base128(loca));
        bytes.push(0);
        bytes
    }

    fn container(directory: &[u8], stream: &[u8], tables: u16) -> Vec<u8> {
        let mut bytes = vec![0; 48];
        bytes[..8].copy_from_slice(b"wOF2\0\x01\0\0");
        bytes[12..14].copy_from_slice(&tables.to_be_bytes());
        bytes[16..20].copy_from_slice(&(12 + 16 * u32::from(tables)).to_be_bytes());
        bytes[20..24].copy_from_slice(&(stream.len() as u32).to_be_bytes());
        bytes.extend_from_slice(directory);
        bytes.extend_from_slice(stream);
        bytes.resize((bytes.len() + 3) & !3, 0);
        let len = bytes.len() as u32;
        bytes[8..12].copy_from_slice(&len.to_be_bytes());
        bytes
    }

    fn glyphs(count: u16, streams: [&[u8]; 7]) -> Vec<u8> {
        let mut bytes = vec![0; 8];
        bytes[4..6].copy_from_slice(&count.to_be_bytes());
        for stream in streams {
            bytes.extend_from_slice(&(stream.len() as u32).to_be_bytes());
        }
        for stream in streams {
            bytes.extend_from_slice(stream);
        }
        bytes
    }

    fn empty_glyph() -> Vec<u8> {
        glyphs(1, [&[0, 0], &[], &[], &[], &[], &[0; 4], &[]])
    }

    fn rejected(payload: &[u8], max_output: usize) {
        assert_eq!(
            validate(payload, max_output),
            Err(DecodeError::Woff2Reconstruction)
        );
        let mut compressor = zstd::bulk::Compressor::new(1).unwrap();
        compressor.include_checksum(true).unwrap();
        compressor.include_contentsize(true).unwrap();
        let frame = compressor.compress(payload).unwrap();
        assert_eq!(
            crate::decode_face(&frame, None, crate::Codec::Woff2tZstd, max_output),
            Err(DecodeError::Woff2Reconstruction)
        );
    }

    #[test]
    fn declared_sizes_and_original_table_sum_are_bounded() {
        let stream = empty_glyph();
        let mut payload = container(&directory(0, 4, stream.len() as u32), &stream, 2);
        assert!(validate(&payload, 4096).is_ok());
        for size in [1u32, 4097] {
            payload[16..20].copy_from_slice(&size.to_be_bytes());
            rejected(&payload, 4096);
        }
        rejected(
            &container(&directory(180, 80, stream.len() as u32), &stream, 2),
            256,
        );
        rejected(
            &container(&directory(1 << 30, 4, stream.len() as u32), &stream, 2),
            65536,
        );
    }

    #[test]
    fn directory_lengths_counts_and_integer_encodings_are_bounded() {
        let stream = empty_glyph();
        for count in [0, 257, 4096, u16::MAX] {
            rejected(
                &container(&directory(0, 4, stream.len() as u32), &stream, count),
                4096,
            );
        }
        rejected(&container(&directory(0, 4, 10000), &stream, 2), 65536);
        let mut tables = directory(0, 4, stream.len() as u32);
        tables.push(0);
        tables.extend(base128(500));
        rejected(&container(&tables, &stream, 3), 4096);
        rejected(&container(&[63, 0], &[], 1), 4096);
        for value in [vec![128, 0], vec![255; 6], vec![144, 128, 128, 128, 0]] {
            let mut tables = vec![10];
            tables.extend(value);
            rejected(&container(&tables, &stream, 2), 4096);
        }
    }

    #[test]
    fn glyph_counts_and_substream_sizes_are_bounded() {
        let mut stream = empty_glyph();
        stream[4..6].copy_from_slice(&u16::MAX.to_be_bytes());
        rejected(
            &container(&directory(0, 4, stream.len() as u32), &stream, 2),
            4096,
        );
        for offset in (8..36).step_by(4) {
            let mut stream = empty_glyph();
            stream[offset..offset + 4].copy_from_slice(&u32::MAX.to_be_bytes());
            rejected(
                &container(&directory(0, 4, stream.len() as u32), &stream, 2),
                4096,
            );
        }
    }

    #[test]
    fn contour_point_counts_are_checked_before_wuff_allocations() {
        let counts = [253, 255, 255].repeat(1023);
        let stream = glyphs(
            1,
            [&1023u16.to_be_bytes(), &counts, &[], &[], &[], &[0; 4], &[]],
        );
        assert_eq!(
            scan_glyphs(&stream, 1 << 30, 64 << 20),
            Err(DecodeError::Woff2Reconstruction)
        );
        let mut payload = container(&directory(1 << 30, 4, stream.len() as u32), &stream, 2);
        payload[16..20].copy_from_slice(&1u32.to_be_bytes());
        rejected(&payload, 64 << 20);
        let stream = glyphs(1, [&[0, 1], &[3], &[0; 3], &[0; 4], &[], &[0; 4], &[]]);
        rejected(
            &container(&directory(2, 4, stream.len() as u32), &stream, 2),
            4096,
        );
        let stream = glyphs(
            2,
            [&[0, 1, 0, 1], &[3, 3], &[0; 6], &[0; 8], &[], &[0; 4], &[]],
        );
        rejected(
            &container(&directory(5, 6, stream.len() as u32), &stream, 2),
            4096,
        );
        let stream = glyphs(
            1,
            [&[0, 1], &[100], &[0; 100], &[0; 101], &[], &[0; 4], &[]],
        );
        rejected(
            &container(&directory(500, 4, stream.len() as u32), &stream, 2),
            1024,
        );
    }

    #[test]
    fn reconstructed_size_does_not_trust_a_small_sfnt_header() {
        let contours = [0, 1].repeat(20);
        let triplets = [0, 0].repeat(20);
        let stream = glyphs(
            20,
            [&contours, &[1; 20], &[0; 20], &triplets, &[], &[0; 4], &[]],
        );
        let payload = container(&directory(20, 42, stream.len() as u32), &stream, 2);
        assert!(validate(&payload, 4096).is_ok());
        rejected(&payload, 350);
    }

    #[test]
    fn point_sizes_include_vertical_signs_and_flag_runs() {
        assert_eq!(point_bytes(&[1, 0, 1], &mut Reader(&[1; 3]), false), Ok(6));
        assert_eq!(
            point_bytes(&[11; 257], &mut Reader(&[1; 257]), false),
            Ok(260)
        );
    }

    #[test]
    fn point_sizes_match_wuff_for_all_triplet_flags() {
        for code in 0..=u8::MAX {
            let flags = [code, code ^ 1, code];
            let mut triplets = Vec::new();
            for flag in flags {
                let size = match flag & 127 {
                    0..=83 => 1,
                    84..=119 => 2,
                    120..=123 => 3,
                    _ => 4,
                };
                triplets.resize(triplets.len() + size, 255);
            }
            let encoded = point_bytes(&flags, &mut Reader(&triplets), false).unwrap();
            triplets.push(0);
            let stream = glyphs(1, [&[0, 1], &[3], &flags, &triplets, &[], &[0; 4], &[]]);
            let payload = container(&directory(64, 4, stream.len() as u32), &stream, 2);
            assert!(validate(&payload, 4096).is_ok());
            let font = wuff::decompress_woff2_with_custom_brotli(&payload, &mut |data, _| {
                Ok(data.to_vec())
            })
            .unwrap();
            assert_eq!(font.len(), 44 + padded(14 + encoded).unwrap() + 4);
        }
    }
}
