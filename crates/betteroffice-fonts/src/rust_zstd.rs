use alloc::vec::Vec;

use ruzstd::decoding::{BlockDecodingStrategy, Dictionary, FrameDecoder};

use crate::{DecodeError, ZSTD_MAGIC, is_zstd};

const RAW_DICT_ID: u32 = 1;
const DICT_MAGIC: [u8; 4] = [0x37, 0xa4, 0x30, 0xec];

// ruzstd only accepts formatted dictionaries. Raw-content prefixes never feed entropy
// tables to a frame, so a minimal valid set stands in for them: a two-symbol Huffman
// description, single-symbol offset, match-length and literal-length FSE tables, then
// the default repeat offsets 1, 4 and 8.
const RAW_DICT_TABLES: [u8; 20] = [
    0x80, 0x10, 0xf0, 0x03, 0xf0, 0x03, 0xf0, 0x03, 1, 0, 0, 0, 4, 0, 0, 0, 8, 0, 0, 0,
];

pub(crate) fn decode(
    frame: &[u8],
    dict: Option<&[u8]>,
    max_output: usize,
) -> Result<Vec<u8>, DecodeError> {
    if !is_zstd(frame) {
        return Err(DecodeError::NotZstd);
    }
    let descriptor = *frame.get(4).ok_or(DecodeError::Truncated)?;
    let size = content_size(frame, descriptor)?;
    if size > max_output as u64 || size > isize::MAX as u64 {
        return Err(DecodeError::TooLarge {
            size,
            max: max_output,
        });
    }
    if descriptor & 4 == 0 {
        return Err(DecodeError::MissingChecksum);
    }
    if compressed_size(frame, descriptor)? != frame.len() {
        return Err(DecodeError::TrailingData);
    }
    let size = size as usize;
    let mut output: Vec<u8> = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|_| DecodeError::AllocationFailed)?;

    let mut decoder = FrameDecoder::new();
    // zstd's default decompression window limit, ZSTD_MAXWINDOWSIZE_DEFAULT.
    decoder.set_max_window_size((1 << 27) + 1);
    if let Some(dict) = dict {
        let mut raw = Vec::new();
        raw.try_reserve_exact(8 + RAW_DICT_TABLES.len() + dict.len())
            .map_err(|_| DecodeError::AllocationFailed)?;
        raw.extend_from_slice(&DICT_MAGIC);
        raw.extend_from_slice(&RAW_DICT_ID.to_le_bytes());
        raw.extend_from_slice(&RAW_DICT_TABLES);
        raw.extend_from_slice(dict);
        let dictionary = Dictionary::decode_dict(&raw).map_err(|_| DecodeError::WrongDictionary)?;
        decoder
            .add_dict(dictionary)
            .map_err(|_| DecodeError::WrongDictionary)?;
    }
    let mut source = frame;
    decoder
        .reset(&mut source)
        .map_err(|_| DecodeError::Corrupt)?;
    if dict.is_some() {
        decoder
            .force_dict(RAW_DICT_ID)
            .map_err(|_| DecodeError::WrongDictionary)?;
    }

    loop {
        let finished = decoder
            .decode_blocks(&mut source, BlockDecodingStrategy::UptoBytes(1 << 20))
            .map_err(|_| DecodeError::Corrupt)?;
        if let Some(chunk) = decoder.collect() {
            if chunk.len() > size - output.len() {
                return Err(DecodeError::Corrupt);
            }
            output.extend_from_slice(&chunk);
        }
        if finished {
            break;
        }
    }
    match (
        decoder.get_checksum_from_data(),
        decoder.get_calculated_checksum(),
    ) {
        (Some(expected), Some(actual)) if expected == actual => {}
        (Some(_), Some(_)) => return Err(DecodeError::ChecksumMismatch),
        _ => return Err(DecodeError::Corrupt),
    }
    if !source.is_empty() || output.len() != size {
        return Err(DecodeError::Corrupt);
    }
    Ok(output)
}

fn header_size(descriptor: u8) -> usize {
    let single_segment = descriptor & 0x20 != 0;
    let content_size = match descriptor >> 6 {
        0 => usize::from(single_segment),
        1 => 2,
        2 => 4,
        _ => 8,
    };
    ZSTD_MAGIC.len()
        + 1
        + usize::from(!single_segment)
        + [0, 1, 2, 4][usize::from(descriptor & 3)]
        + content_size
}

fn content_size(frame: &[u8], descriptor: u8) -> Result<u64, DecodeError> {
    let header = header_size(descriptor);
    if descriptor & 0x08 != 0 || frame.len() < header {
        return Err(DecodeError::Corrupt);
    }
    let field = match descriptor >> 6 {
        0 if descriptor & 0x20 != 0 => 1,
        0 => return Err(DecodeError::MissingContentSize),
        1 => 2,
        2 => 4,
        _ => 8,
    };
    let mut value = [0u8; 8];
    value[..field].copy_from_slice(&frame[header - field..header]);
    let value = u64::from_le_bytes(value);
    Ok(if field == 2 { value + 256 } else { value })
}

// The frame length the blocks and checksum declare, as zstd's findFrameCompressedSize.
fn compressed_size(frame: &[u8], descriptor: u8) -> Result<usize, DecodeError> {
    let mut offset = header_size(descriptor);
    loop {
        let header = frame
            .get(offset..offset + 3)
            .ok_or(DecodeError::Truncated)?;
        let header = u32::from(header[0]) | u32::from(header[1]) << 8 | u32::from(header[2]) << 16;
        let size = (header >> 3) as usize;
        let content = match (header >> 1) & 3 {
            0 | 2 => size,
            1 => 1,
            _ => return Err(DecodeError::Corrupt),
        };
        if size > 128 << 10 {
            return Err(DecodeError::Corrupt);
        }
        offset += 3 + content;
        if offset > frame.len() {
            return Err(DecodeError::Truncated);
        }
        if header & 1 != 0 {
            break;
        }
    }
    offset += 4;
    if offset > frame.len() {
        return Err(DecodeError::Truncated);
    }
    Ok(offset)
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::Path;

    use sha2::{Digest, Sha256};

    use super::decode;
    use crate::DecodeError;

    fn fixture(name: &str) -> Vec<u8> {
        fs::read(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/fixtures")
                .join(name),
        )
        .unwrap()
    }

    fn compress(bytes: &[u8], dict: Option<&[u8]>, checksum: bool) -> Vec<u8> {
        let mut compressor = match dict {
            Some(dict) => zstd::bulk::Compressor::with_dictionary(19, dict).unwrap(),
            None => zstd::bulk::Compressor::new(19).unwrap(),
        };
        compressor
            .set_parameter(zstd::zstd_safe::CParameter::ChecksumFlag(checksum))
            .unwrap();
        compressor.compress(bytes).unwrap()
    }

    #[test]
    fn matches_the_c_decoder_on_fixtures() {
        let regular = fixture("Synthetic-Regular.uw2");
        for (name, dict) in [
            ("Synthetic-Regular.ttf.zst", None),
            ("Synthetic-Bold.ttf.zst", Some(regular.as_slice())),
            ("Synthetic-Bold.nodict.zst", None),
        ] {
            let frame = fixture(name);
            assert_eq!(
                decode(&frame, dict, 1 << 20),
                crate::decode(&frame, dict, 1 << 20),
                "{name}"
            );
        }
    }

    #[test]
    fn matches_the_c_decoder_on_errors() {
        let regular = fixture("Synthetic-Regular.uw2");
        let frame = fixture("Synthetic-Bold.ttf.zst");
        let mut corrupt = frame.clone();
        *corrupt.last_mut().unwrap() ^= 1;
        let mut trailing = fixture("Synthetic-Regular.ttf.zst");
        trailing.push(0);
        type Case<'a> = (&'a [u8], Option<&'a [u8]>, usize);
        let cases: [Case; 6] = [
            (&corrupt, Some(&regular), 1 << 20),
            (&trailing, None, 1 << 20),
            (&frame, Some(&regular), 16),
            (&frame[..frame.len() - 3], Some(&regular), 1 << 20),
            (b"OTTO", None, 1 << 20),
            (&frame[..5], None, 1 << 20),
        ];
        for (index, (bytes, dict, max)) in cases.into_iter().enumerate() {
            let expected = crate::decode(bytes, dict, max);
            assert!(expected.is_err(), "case {index}");
            assert_eq!(decode(bytes, dict, max), expected, "case {index}");
        }
        let wrong = decode(&frame, Some(&fixture("Synthetic-Bold.uw2")), 1 << 20);
        assert!(matches!(
            wrong,
            Err(DecodeError::ChecksumMismatch | DecodeError::Corrupt)
        ));
    }

    #[test]
    fn rejects_frames_without_checksum_or_size_like_the_c_decoder() {
        let bytes = b"plain bytes for a frame without a checksum".repeat(40);
        let unchecked = compress(&bytes, None, false);
        assert_eq!(
            decode(&unchecked, None, 1 << 20),
            crate::decode(&unchecked, None, 1 << 20)
        );
        assert_eq!(
            decode(&unchecked, None, 1 << 20),
            Err(DecodeError::MissingChecksum)
        );
    }

    #[test]
    fn decodes_generated_frames_with_raw_dictionaries() {
        let dict: Vec<u8> = (0..60_000u32).map(|i| (i * 7 % 251) as u8).collect();
        let mut bytes = dict[1_000..40_000].to_vec();
        bytes.extend((0..30_000u32).map(|i| (i * 13 % 241) as u8));
        bytes.extend_from_slice(&dict[..20_000]);
        for dict in [None, Some(dict.as_slice())] {
            let frame = compress(&bytes, dict, true);
            assert_eq!(decode(&frame, dict, 1 << 20).unwrap(), bytes);
            assert_eq!(crate::decode(&frame, dict, 1 << 20).unwrap(), bytes);
        }
    }

    // Mirrors block_bomb in tests/fixtures/make_fixtures.py: literal length 1, offset 1
    // and one match length code for every sequence, all three tables in RLE mode.
    fn block_bomb(
        sequences: usize,
        match_code: u8,
        match_bits: u32,
        content_size: u16,
        checksum: [u8; 4],
    ) -> Vec<u8> {
        let mut block = vec![
            0x04 | ((sequences & 0xf) << 4) as u8,
            (sequences >> 4) as u8,
        ];
        block.extend(std::iter::repeat_n(b'A', sequences));
        block.extend([sequences as u8, 0x54, 1, 2, match_code]);
        let mut stream = Vec::new();
        let (mut accumulator, mut width) = (0u64, 0u32);
        let mut push = |value: u64, bits: u32, stream: &mut Vec<u8>| {
            accumulator |= value << width;
            width += bits;
            while width >= 8 {
                stream.push(accumulator as u8);
                accumulator >>= 8;
                width -= 8;
            }
        };
        for _ in 0..sequences {
            push((1 << match_bits) - 1, match_bits, &mut stream);
            push(0, 2, &mut stream);
        }
        push(1, 1, &mut stream);
        push(0, 7, &mut stream);
        block.extend(stream);
        let header = ((block.len() as u32) << 3) | (2 << 1) | 1;
        let mut frame = vec![0x28, 0xb5, 0x2f, 0xfd, 0x64];
        frame.extend((content_size - 256).to_le_bytes());
        frame.extend(&header.to_le_bytes()[..3]);
        frame.extend(block);
        frame.extend(checksum);
        frame
    }

    #[test]
    fn a_block_expanding_past_its_frame_fails_on_both_decoders() {
        let expected = vec![b'A'; 515];
        let reference = compress(&expected, None, true);
        let checksum = reference[reference.len() - 4..].try_into().unwrap();
        let control = block_bomb(1, 44, 8, 515, checksum);
        assert!(control.len() < 32);
        assert_eq!(decode(&control, None, 1 << 20).unwrap(), expected);
        assert_eq!(crate::decode(&control, None, 1 << 20).unwrap(), expected);

        let bomb = fixture("Synthetic-block-bomb.zst");
        assert_eq!(bomb, block_bomb(100, 52, 16, 1000, [0; 4]));
        for cap in [1 << 20, 1 << 30] {
            assert_eq!(decode(&bomb, None, cap), Err(DecodeError::Corrupt));
            assert_eq!(crate::decode(&bomb, None, cap), Err(DecodeError::Corrupt));
        }
    }

    #[test]
    fn decodes_every_shipped_pack_face_byte_identically() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../packages");
        let mut checked = 0;
        for pack in ["fonts-intl", "fonts-latin", "fonts-cjk"] {
            let dir = root.join(pack).join("assets");
            let Ok(manifest) = fs::read(dir.join("manifest.json")) else {
                continue;
            };
            let manifest: serde_json::Value = serde_json::from_slice(&manifest).unwrap();
            for face in manifest["faces"].as_array().unwrap() {
                let frame = fs::read(dir.join(face["file"].as_str().unwrap())).unwrap();
                let dict = face["dict"].as_str().map(|file| {
                    crate::decode(&fs::read(dir.join(file)).unwrap(), None, 1 << 30).unwrap()
                });
                let decoded = decode(&frame, dict.as_deref(), 1 << 30).unwrap();
                assert_eq!(
                    Ok(&decoded),
                    crate::decode(&frame, dict.as_deref(), 1 << 30).as_ref(),
                    "{pack}/{}",
                    face["file"]
                );
                if face["codec"] == "zstd" {
                    assert_eq!(
                        format!("{:x}", Sha256::digest(&decoded)),
                        face["sha256"].as_str().unwrap()
                    );
                }
                checked += 1;
            }
        }
        assert!(checked > 0);
        println!("pack faces decoded byte-identically: {checked}");
    }
}
