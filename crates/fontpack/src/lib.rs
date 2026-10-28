#![cfg_attr(
    all(
        target_arch = "wasm32",
        feature = "wasm",
        not(feature = "build"),
        not(test)
    ),
    no_std
)]

extern crate alloc;

use alloc::borrow::Cow;
use alloc::vec::Vec;
use core::ffi::{c_uint, c_void};
use core::ptr::{self, NonNull};

#[cfg(any(
    not(target_arch = "wasm32"),
    not(feature = "wasm"),
    feature = "build",
    test
))]
use core::fmt;

#[cfg(target_arch = "wasm32")]
mod allocator;
#[cfg(all(target_arch = "wasm32", feature = "wasm"))]
mod wasm;
mod woff2t;

pub const ZSTD_MAGIC: [u8; 4] = [0x28, 0xb5, 0x2f, 0xfd];

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Codec {
    Zstd,
    Woff2tZstd,
}

impl Codec {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Zstd => "zstd",
            Self::Woff2tZstd => "woff2t+zstd",
        }
    }
}

impl core::str::FromStr for Codec {
    type Err = &'static str;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "zstd" => Ok(Self::Zstd),
            "woff2t+zstd" => Ok(Self::Woff2tZstd),
            _ => Err("unsupported font face codec"),
        }
    }
}

#[derive(Debug, Eq, PartialEq)]
pub struct Decoded {
    pub font: Vec<u8>,
    pub payload: Option<Vec<u8>>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum DecodeError {
    NotZstd,
    MissingContentSize,
    MissingChecksum,
    TooLarge { size: u64, max: usize },
    Corrupt,
    ChecksumMismatch,
    WrongDictionary,
    Truncated,
    TrailingData,
    AllocationFailed,
    Woff2Reconstruction,
}

#[cfg(any(
    not(target_arch = "wasm32"),
    not(feature = "wasm"),
    feature = "build",
    test
))]
impl fmt::Display for DecodeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NotZstd => f.write_str("font face is not a zstd frame"),
            Self::MissingContentSize => f.write_str("font face has no decoded content size"),
            Self::MissingChecksum => f.write_str("font face has no content checksum"),
            Self::TooLarge { size, max } => {
                write!(
                    f,
                    "font face is too large: {size} decoded bytes exceed {max}"
                )
            }
            Self::Corrupt => f.write_str("font face is corrupt or uses a wrong dictionary"),
            Self::ChecksumMismatch => {
                f.write_str("font face checksum failed: corrupt face or wrong dictionary")
            }
            Self::WrongDictionary => f.write_str("font face dictionary is missing or wrong"),
            Self::Truncated => f.write_str("font face is truncated"),
            Self::TrailingData => f.write_str("font face must contain exactly one zstd frame"),
            Self::AllocationFailed => f.write_str("cannot allocate memory for font face"),
            Self::Woff2Reconstruction => f.write_str("font face WOFF2 reconstruction failed"),
        }
    }
}

#[cfg(any(
    not(target_arch = "wasm32"),
    not(feature = "wasm"),
    feature = "build",
    test
))]
impl std::error::Error for DecodeError {}

unsafe extern "C" {
    #[link_name = "FONTPACK_ZSTD_getFrameContentSize"]
    fn frame_content_size(src: *const c_void, len: usize) -> u64;
    #[link_name = "FONTPACK_ZSTD_findFrameCompressedSize"]
    fn frame_compressed_size(src: *const c_void, len: usize) -> usize;
    #[link_name = "FONTPACK_ZSTD_isError"]
    fn is_error(code: usize) -> c_uint;
    #[link_name = "FONTPACK_ZSTD_getErrorCode"]
    fn error_code(code: usize) -> c_uint;
    #[link_name = "FONTPACK_ZSTD_createDCtx"]
    fn create_context() -> *mut c_void;
    #[link_name = "FONTPACK_ZSTD_freeDCtx"]
    fn free_context(context: *mut c_void) -> usize;
    #[link_name = "FONTPACK_ZSTD_DCtx_refPrefix"]
    fn ref_prefix(context: *mut c_void, dict: *const c_void, len: usize) -> usize;
    #[link_name = "FONTPACK_ZSTD_decompressDCtx"]
    fn decompress(
        context: *mut c_void,
        dst: *mut c_void,
        capacity: usize,
        src: *const c_void,
        len: usize,
    ) -> usize;
}

struct Context(NonNull<c_void>);

impl Drop for Context {
    fn drop(&mut self) {
        unsafe { free_context(self.0.as_ptr()) };
    }
}

fn check(code: usize) -> Result<usize, DecodeError> {
    if unsafe { is_error(code) } == 0 {
        return Ok(code);
    }
    Err(match unsafe { error_code(code) } {
        22 => DecodeError::ChecksumMismatch,
        30 | 32 | 34 => DecodeError::WrongDictionary,
        64 => DecodeError::AllocationFailed,
        72 => DecodeError::Truncated,
        _ => DecodeError::Corrupt,
    })
}

pub fn is_zstd(bytes: &[u8]) -> bool {
    bytes.starts_with(&ZSTD_MAGIC)
}

pub fn decode(
    frame: &[u8],
    dict: Option<&[u8]>,
    max_output: usize,
) -> Result<Vec<u8>, DecodeError> {
    if !is_zstd(frame) {
        return Err(DecodeError::NotZstd);
    }
    let descriptor = *frame.get(4).ok_or(DecodeError::Truncated)?;
    let size = unsafe { frame_content_size(frame.as_ptr().cast(), frame.len()) };
    match size {
        u64::MAX => return Err(DecodeError::MissingContentSize),
        value if value == u64::MAX - 1 => return Err(DecodeError::Corrupt),
        _ => {}
    }
    if size > max_output as u64 || size > isize::MAX as u64 {
        return Err(DecodeError::TooLarge {
            size,
            max: max_output,
        });
    }
    if descriptor & 4 == 0 {
        return Err(DecodeError::MissingChecksum);
    }
    let encoded_size = check(unsafe { frame_compressed_size(frame.as_ptr().cast(), frame.len()) })?;
    if encoded_size != frame.len() {
        return Err(DecodeError::TrailingData);
    }

    let mut output: Vec<u8> = Vec::new();
    output
        .try_reserve_exact(size as usize)
        .map_err(|_| DecodeError::AllocationFailed)?;
    unsafe {
        ptr::write_bytes(output.as_mut_ptr(), 0, size as usize);
        output.set_len(size as usize);
    }
    let context =
        Context(NonNull::new(unsafe { create_context() }).ok_or(DecodeError::AllocationFailed)?);
    if let Some(dict) = dict {
        check(unsafe { ref_prefix(context.0.as_ptr(), dict.as_ptr().cast(), dict.len()) })?;
    }
    let written = check(unsafe {
        decompress(
            context.0.as_ptr(),
            output.as_mut_ptr().cast(),
            output.len(),
            frame.as_ptr().cast(),
            frame.len(),
        )
    })?;
    if written != output.len() {
        return Err(DecodeError::Corrupt);
    }
    Ok(output)
}

pub fn decode_or_passthrough<'a>(
    bytes: &'a [u8],
    dict: Option<&[u8]>,
    max_output: usize,
) -> Result<Cow<'a, [u8]>, DecodeError> {
    if is_zstd(bytes) {
        decode(bytes, dict, max_output).map(Cow::Owned)
    } else {
        Ok(Cow::Borrowed(bytes))
    }
}

pub fn decode_face(
    frame: &[u8],
    dict: Option<&[u8]>,
    codec: Codec,
    max_output: usize,
) -> Result<Decoded, DecodeError> {
    let payload = decode(frame, dict, max_output)?;
    match codec {
        Codec::Zstd => Ok(Decoded {
            font: payload,
            payload: None,
        }),
        Codec::Woff2tZstd => {
            let font = reconstruct_woff2t(&payload, max_output)?;
            Ok(Decoded {
                font,
                payload: Some(payload),
            })
        }
    }
}

fn reconstruct_woff2t(payload: &[u8], max_output: usize) -> Result<Vec<u8>, DecodeError> {
    woff2t::validate(payload, max_output)?;
    let mut identity_error = None;
    let result = wuff::decompress_woff2_with_custom_brotli(payload, &mut |data, expected| {
        if data.len() != expected || expected > max_output {
            return Err(wuff::WuffErr::GenericError.into());
        }
        let mut bytes = Vec::new();
        if bytes.try_reserve_exact(data.len()).is_err() {
            identity_error = Some(DecodeError::AllocationFailed);
            return Err(wuff::WuffErr::GenericError.into());
        }
        bytes.extend_from_slice(data);
        Ok(bytes)
    });
    let font = result.map_err(|_| identity_error.unwrap_or(DecodeError::Woff2Reconstruction))?;
    if font.len() > max_output {
        return Err(DecodeError::TooLarge {
            size: font.len() as u64,
            max: max_output,
        });
    }
    Ok(font)
}

#[cfg(test)]
mod tests {
    use super::{ZSTD_MAGIC, decode};

    #[test]
    fn arbitrary_headers_never_panic() {
        for descriptor in 0..=u8::MAX {
            let mut bytes = ZSTD_MAGIC.to_vec();
            bytes.push(descriptor);
            bytes.extend_from_slice(&[0; 16]);
            for len in 0..=bytes.len() {
                let _ = decode(&bytes[..len], None, 1024);
            }
        }
    }
}
