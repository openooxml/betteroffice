mod common;

use fontpack::{Codec, DecodeError, decode_face};

const REGULAR: &[u8] = include_bytes!("fixtures/Synthetic-Regular.uw2");
const BOLD: &[u8] = include_bytes!("fixtures/Synthetic-Bold.uw2");

#[test]
fn codecs_parse_manifest_values() {
    for codec in [Codec::Zstd, Codec::Woff2tZstd] {
        assert_eq!(codec.as_str().parse::<Codec>(), Ok(codec));
    }
    assert!("woff2t".parse::<Codec>().is_err());
}

#[test]
fn plain_zstd_has_no_separate_payload() {
    let bytes = common::synthetic(42);
    let result = decode_face(
        &common::compress(&bytes, None),
        None,
        Codec::Zstd,
        bytes.len(),
    )
    .unwrap();
    assert_eq!(result.font, bytes);
    assert!(result.payload.is_none());
}

#[test]
fn transformed_faces_round_trip_with_and_without_dictionary() {
    let regular = decode_face(
        &common::compress(REGULAR, None),
        None,
        Codec::Woff2tZstd,
        65536,
    )
    .unwrap();
    assert_eq!(regular.payload.as_deref(), Some(REGULAR));
    assert!(regular.font.starts_with(b"\0\x01\0\0"));
    for dict in [None, regular.payload.as_deref()] {
        let result = decode_face(
            &common::compress(BOLD, dict),
            dict,
            Codec::Woff2tZstd,
            65536,
        )
        .unwrap();
        assert_eq!(result.payload.as_deref(), Some(BOLD));
        #[cfg(feature = "build")]
        assert_eq!(
            result.font,
            wuff::decompress_woff2(include_bytes!("fixtures/Synthetic-Bold.woff2")).unwrap()
        );
    }
    #[cfg(feature = "build")]
    assert_eq!(
        regular.font,
        wuff::decompress_woff2(include_bytes!("fixtures/Synthetic-Regular.woff2")).unwrap()
    );
}

#[test]
fn transformed_face_rejects_wrong_dictionary_and_corruption() {
    let frame = common::compress(BOLD, Some(REGULAR));
    assert!(frame.len() < common::compress(BOLD, None).len());
    for dict in [None, Some(vec![0; REGULAR.len()])] {
        assert!(decode_face(&frame, dict.as_deref(), Codec::Woff2tZstd, 65536).is_err());
    }
    let mut corrupt = frame;
    *corrupt.last_mut().unwrap() ^= 1;
    assert_eq!(
        decode_face(&corrupt, Some(REGULAR), Codec::Woff2tZstd, 65536),
        Err(DecodeError::ChecksumMismatch)
    );
    let mut payload = REGULAR.to_vec();
    payload[0] ^= 1;
    assert_eq!(
        decode_face(
            &common::compress(&payload, None),
            None,
            Codec::Woff2tZstd,
            65536
        ),
        Err(DecodeError::Woff2Reconstruction)
    );
    let standard = include_bytes!("fixtures/Synthetic-Regular.woff2");
    assert_eq!(
        decode_face(
            &common::compress(standard, None),
            None,
            Codec::Woff2tZstd,
            65536
        ),
        Err(DecodeError::Woff2Reconstruction)
    );
}

#[test]
fn both_payload_and_reconstructed_font_are_bounded() {
    let frame = common::compress(REGULAR, None);
    assert!(matches!(
        decode_face(&frame, None, Codec::Woff2tZstd, REGULAR.len() - 1),
        Err(DecodeError::TooLarge { .. })
    ));
    let mut payload = REGULAR.to_vec();
    payload[16..20].copy_from_slice(&65537u32.to_be_bytes());
    assert_eq!(
        decode_face(
            &common::compress(&payload, None),
            None,
            Codec::Woff2tZstd,
            65536
        ),
        Err(DecodeError::Woff2Reconstruction)
    );
    let font = decode_face(&frame, None, Codec::Woff2tZstd, 65536)
        .unwrap()
        .font;
    payload[16..20].copy_from_slice(&1u32.to_be_bytes());
    assert!(font.len() > payload.len());
    assert!(matches!(
        decode_face(
            &common::compress(&payload, None),
            None,
            Codec::Woff2tZstd,
            font.len() - 1
        ),
        Err(DecodeError::Woff2Reconstruction)
    ));
}

#[test]
fn malformed_woff2_headers_return_errors() {
    for len in 0..48 {
        let frame = common::compress(&REGULAR[..len], None);
        assert_eq!(
            decode_face(&frame, None, Codec::Woff2tZstd, 65536),
            Err(DecodeError::Woff2Reconstruction)
        );
    }
    for count in [0u16, 4096, u16::MAX] {
        let mut payload = REGULAR.to_vec();
        payload[12..14].copy_from_slice(&count.to_be_bytes());
        assert_eq!(
            decode_face(
                &common::compress(&payload, None),
                None,
                Codec::Woff2tZstd,
                65536
            ),
            Err(DecodeError::Woff2Reconstruction)
        );
    }
    for offset in [8, 20, 28, 32, 40, 44] {
        let mut payload = REGULAR.to_vec();
        payload[offset..offset + 4].copy_from_slice(&u32::MAX.to_be_bytes());
        assert_eq!(
            decode_face(
                &common::compress(&payload, None),
                None,
                Codec::Woff2tZstd,
                65536
            ),
            Err(DecodeError::Woff2Reconstruction)
        );
    }
}

#[test]
fn corrupt_transformed_bytes_do_not_panic() {
    for index in 0..REGULAR.len() {
        let mut payload = REGULAR.to_vec();
        payload[index] ^= 0x80;
        let frame = common::compress(&payload, None);
        let _ = decode_face(&frame, None, Codec::Woff2tZstd, 65536);
    }
}
