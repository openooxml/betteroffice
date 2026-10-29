use alloc::string::String;
use core::fmt;

use crate::{Codec, DecodeError, Decoded, decode_face, is_zstd};

pub const DEFAULT_MAX_OUTPUT: usize = 64 << 20;
pub const MAX_OUTPUT_LIMIT: usize = 0x7fff_ffff;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FaceError {
    UnsupportedCodec(String),
    InvalidMaxOutput,
    Decode(DecodeError),
}

impl fmt::Display for FaceError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedCodec(codec) => write!(f, "unsupported font face codec: {codec}"),
            Self::InvalidMaxOutput => {
                f.write_str("font face maxOutput must be an integer from 0 to 2147483647")
            }
            Self::Decode(error) => fmt::Display::fmt(error, f),
        }
    }
}

#[cfg(any(
    not(target_arch = "wasm32"),
    not(feature = "wasm"),
    feature = "build",
    test
))]
impl std::error::Error for FaceError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Decode(error) => Some(error),
            _ => None,
        }
    }
}

#[inline(never)]
pub fn decode_font_face(
    bytes: &[u8],
    codec: &str,
    dict: Option<&[u8]>,
    max_output: Option<usize>,
) -> Result<Decoded, FaceError> {
    let codec = codec
        .parse::<Codec>()
        .map_err(|_| FaceError::UnsupportedCodec(String::from(codec)))?;
    if !is_zstd(bytes) {
        return Ok(Decoded {
            font: bytes.to_vec(),
            payload: None,
        });
    }
    let max_output = max_output.unwrap_or(DEFAULT_MAX_OUTPUT);
    if max_output > MAX_OUTPUT_LIMIT {
        return Err(FaceError::InvalidMaxOutput);
    }
    decode_face(bytes, dict, codec, max_output).map_err(FaceError::Decode)
}

/// [`decode_font_face`] with the output limit as a JavaScript number, validated only
/// when a supported codec meets a zstd frame, as the standalone decoder does.
#[inline(never)]
pub fn decode_font_face_js(
    bytes: &[u8],
    codec: &str,
    dict: Option<&[u8]>,
    max_output: Option<f64>,
) -> Result<Decoded, FaceError> {
    let max_output = match max_output {
        Some(value) if codec.parse::<Codec>().is_ok() && is_zstd(bytes) => {
            Some(max_output_from_f64(value)?)
        }
        _ => None,
    };
    decode_font_face(bytes, codec, dict, max_output)
}

pub fn max_output_from_f64(value: f64) -> Result<usize, FaceError> {
    if value.is_finite()
        && (0.0..=MAX_OUTPUT_LIMIT as f64).contains(&value)
        && value == value as usize as f64
    {
        Ok(value as usize)
    } else {
        Err(FaceError::InvalidMaxOutput)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const REGULAR: &[u8] = include_bytes!("../tests/fixtures/Synthetic-Regular.ttf");
    const BOLD: &[u8] = include_bytes!("../tests/fixtures/Synthetic-Bold.ttf");
    const REGULAR_FRAME: &[u8] = include_bytes!("../tests/fixtures/Synthetic-Regular.sfnt.zst");
    const BOLD_FRAME: &[u8] = include_bytes!("../tests/fixtures/Synthetic-Bold.sfnt.zst");
    const REGULAR_PAYLOAD: &[u8] = include_bytes!("../tests/fixtures/Synthetic-Regular.uw2");
    const BOLD_PAYLOAD: &[u8] = include_bytes!("../tests/fixtures/Synthetic-Bold.uw2");

    #[test]
    fn codec_validation_precedes_passthrough_and_limits() {
        for bytes in [REGULAR, REGULAR_FRAME, &[]] {
            assert_eq!(
                decode_font_face(bytes, "brotli", None, Some(usize::MAX)),
                Err(FaceError::UnsupportedCodec(String::from("brotli")))
            );
        }
    }

    #[test]
    fn non_zstd_bytes_pass_through_without_limits_or_payload() {
        for codec in ["zstd", "woff2t+zstd"] {
            for bytes in [
                REGULAR,
                include_bytes!("../tests/fixtures/Synthetic-Regular.woff2").as_slice(),
                &[0x28, 0xb5],
            ] {
                assert_eq!(
                    decode_font_face(bytes, codec, None, Some(usize::MAX)),
                    Ok(Decoded {
                        font: bytes.to_vec(),
                        payload: None,
                    })
                );
            }
        }
    }

    #[test]
    fn javascript_limits_are_checked_only_for_decodable_frames() {
        assert_eq!(
            decode_font_face_js(REGULAR, "zstd", None, Some(1.5))
                .unwrap()
                .font,
            REGULAR
        );
        assert_eq!(
            decode_font_face_js(REGULAR_FRAME, "brotli", None, Some(1.5)),
            Err(FaceError::UnsupportedCodec(String::from("brotli")))
        );
        assert_eq!(
            decode_font_face_js(REGULAR_FRAME, "zstd", None, Some(1.5)),
            Err(FaceError::InvalidMaxOutput)
        );
    }

    #[test]
    fn output_limits_and_javascript_numbers_are_validated() {
        assert_eq!(DEFAULT_MAX_OUTPUT, 64 << 20);
        for value in [
            f64::NAN,
            f64::INFINITY,
            f64::NEG_INFINITY,
            -1.0,
            1.5,
            2147483648.0,
        ] {
            assert_eq!(max_output_from_f64(value), Err(FaceError::InvalidMaxOutput));
        }
        assert_eq!(max_output_from_f64(0.0), Ok(0));
        assert_eq!(max_output_from_f64(2147483647.0), Ok(MAX_OUTPUT_LIMIT));
        for codec in ["zstd", "woff2t+zstd"] {
            assert_eq!(
                decode_font_face(REGULAR_FRAME, codec, None, Some(MAX_OUTPUT_LIMIT + 1)),
                Err(FaceError::InvalidMaxOutput)
            );
        }
        assert!(matches!(
            decode_font_face(REGULAR_FRAME, "zstd", None, Some(0)),
            Err(FaceError::Decode(DecodeError::TooLarge { max: 0, .. }))
        ));
        assert_eq!(
            decode_font_face(REGULAR_FRAME, "zstd", None, Some(MAX_OUTPUT_LIMIT))
                .unwrap()
                .font,
            REGULAR
        );
        assert_eq!(
            decode_font_face(
                include_bytes!("../tests/fixtures/Synthetic-Regular.ttf.zst"),
                "woff2t+zstd",
                None,
                Some(MAX_OUTPUT_LIMIT),
            )
            .unwrap()
            .payload
            .as_deref(),
            Some(REGULAR_PAYLOAD)
        );
    }

    #[test]
    fn plain_zstd_fixtures_preserve_exact_source_fonts() {
        for (frame, dictionary, expected) in [
            (REGULAR_FRAME, None, REGULAR),
            (BOLD_FRAME, Some(REGULAR), BOLD),
        ] {
            let decoded = decode_font_face(frame, "zstd", dictionary, None).unwrap();
            assert_eq!(decoded.font, expected);
            assert_eq!(decoded.payload, None);
        }
    }

    #[test]
    fn transformed_fixtures_preserve_payload_and_reconstruct_reference_fonts() {
        for (frame, dictionary, payload, reference) in [
            (
                include_bytes!("../tests/fixtures/Synthetic-Regular.ttf.zst").as_slice(),
                None,
                REGULAR_PAYLOAD,
                include_bytes!("../tests/fixtures/Synthetic-Regular.woff2").as_slice(),
            ),
            (
                include_bytes!("../tests/fixtures/Synthetic-Bold.ttf.zst").as_slice(),
                Some(REGULAR_PAYLOAD),
                BOLD_PAYLOAD,
                include_bytes!("../tests/fixtures/Synthetic-Bold.woff2").as_slice(),
            ),
            (
                include_bytes!("../tests/fixtures/Synthetic-Bold.nodict.zst").as_slice(),
                None,
                BOLD_PAYLOAD,
                include_bytes!("../tests/fixtures/Synthetic-Bold.woff2").as_slice(),
            ),
        ] {
            let decoded = decode_font_face(frame, "woff2t+zstd", dictionary, None).unwrap();
            assert_eq!(decoded.payload.as_deref(), Some(payload));
            assert!(decoded.font.starts_with(b"\0\x01\0\0"));
            #[cfg(feature = "build")]
            assert_eq!(decoded.font, wuff::decompress_woff2(reference).unwrap());
            #[cfg(not(feature = "build"))]
            let _ = reference;
        }
    }

    #[test]
    fn errors_have_shared_display_messages() {
        assert_eq!(
            FaceError::UnsupportedCodec(String::from("brotli")).to_string(),
            "unsupported font face codec: brotli"
        );
        assert_eq!(
            FaceError::InvalidMaxOutput.to_string(),
            "font face maxOutput must be an integer from 0 to 2147483647"
        );
        assert_eq!(
            FaceError::Decode(DecodeError::Truncated).to_string(),
            DecodeError::Truncated.to_string()
        );
    }
}
