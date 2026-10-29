mod common;

use std::borrow::Cow;

use common::{compress, synthetic};
use fontpack::{DecodeError, ZSTD_MAGIC, decode, decode_or_passthrough, is_zstd};

#[test]
fn round_trip_without_dictionary() {
    for bytes in [Vec::new(), synthetic(42)] {
        let frame = compress(&bytes, None);
        assert!(is_zstd(&frame));
        assert_eq!(decode(&frame, None, bytes.len()).unwrap(), bytes);
        assert!(matches!(
            decode_or_passthrough(&frame, None, bytes.len()).unwrap(),
            Cow::Owned(_)
        ));
    }
}

#[test]
fn raw_dictionary_round_trip_and_wrong_dictionary() {
    let dict = synthetic(42);
    let mut bytes = dict.clone();
    bytes[100] ^= 1;
    let frame = compress(&bytes, Some(&dict));
    assert!(frame.len() < compress(&bytes, None).len());
    assert_eq!(decode(&frame, Some(&dict), bytes.len()).unwrap(), bytes);
    let wrong = synthetic(43);
    assert!(matches!(
        decode(&frame, Some(&wrong), bytes.len()),
        Err(DecodeError::ChecksumMismatch | DecodeError::WrongDictionary | DecodeError::Corrupt)
    ));
    assert!(decode(&frame, None, bytes.len()).is_err());
}

#[test]
fn raw_dictionary_magic_is_not_interpreted_as_a_formatted_dictionary() {
    let mut dict = synthetic(42);
    dict[..4].copy_from_slice(&[0x37, 0xa4, 0x30, 0xec]);
    let mut bytes = dict.clone();
    bytes[100] ^= 1;
    let frame = compress(&bytes, Some(&dict));
    assert_eq!(decode(&frame, Some(&dict), bytes.len()).unwrap(), bytes);
}

#[test]
fn checksum_corruption_is_reported() {
    let bytes = synthetic(42);
    let mut frame = compress(&bytes, None);
    *frame.last_mut().unwrap() ^= 1;
    assert_eq!(
        decode(&frame, None, bytes.len()),
        Err(DecodeError::ChecksumMismatch)
    );
}

#[test]
fn truncated_and_concatenated_frames_are_rejected() {
    let bytes = synthetic(42);
    let frame = compress(&bytes, None);
    for len in 0..frame.len() {
        assert!(decode(&frame[..len], None, bytes.len()).is_err());
    }
    let mut concatenated = frame.clone();
    concatenated.extend_from_slice(&frame);
    assert_eq!(
        decode(&concatenated, None, bytes.len() * 2),
        Err(DecodeError::TrailingData)
    );
}

#[test]
fn output_limit_is_checked_before_the_payload_or_allocation() {
    let mut frame = ZSTD_MAGIC.to_vec();
    frame.push(0xe4);
    frame.extend_from_slice(&u64::MAX.wrapping_sub(2).to_le_bytes());
    assert_eq!(
        decode(&frame, None, 64),
        Err(DecodeError::TooLarge {
            size: u64::MAX - 2,
            max: 64
        })
    );
    let bytes = synthetic(42);
    assert!(matches!(
        decode(&compress(&bytes, None), None, bytes.len() - 1),
        Err(DecodeError::TooLarge { .. })
    ));
}

#[test]
fn unknown_size_and_missing_checksum_are_rejected() {
    let unknown_size = [0x28, 0xb5, 0x2f, 0xfd, 0x04, 0x00];
    assert_eq!(
        decode(&unknown_size, None, 64),
        Err(DecodeError::MissingContentSize)
    );
    let no_checksum = [0x28, 0xb5, 0x2f, 0xfd, 0x20, 0x00, 0x01, 0x00, 0x00];
    assert_eq!(
        decode(&no_checksum, None, 64),
        Err(DecodeError::MissingChecksum)
    );
}

#[test]
fn other_formats_pass_through_without_decoding() {
    for bytes in [
        b"\0\x01\0\0".as_slice(),
        b"OTTO",
        b"wOFF",
        b"wOF2",
        b"",
        b"\x28\xb5\x2f",
        b"\x50\x2a\x4d\x18",
    ] {
        assert!(!is_zstd(bytes));
        let decoded = decode_or_passthrough(bytes, None, 0).unwrap();
        assert!(matches!(decoded, Cow::Borrowed(_)));
        assert_eq!(decoded.as_ref(), bytes);
        assert_eq!(decoded.as_ptr(), bytes.as_ptr());
    }
}

#[test]
fn decoded_font_is_not_decoded_twice() {
    let bytes = b"OTTOsynthetic";
    let frame = compress(bytes, None);
    let decoded = decode(&frame, None, bytes.len()).unwrap();
    assert!(matches!(
        decode_or_passthrough(&decoded, None, 0).unwrap(),
        Cow::Borrowed(_)
    ));
}
