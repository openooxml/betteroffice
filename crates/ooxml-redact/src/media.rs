use std::io::Cursor;

use image::{DynamicImage, ImageBuffer, ImageFormat, Rgb};

use crate::{RedactError, RedactionReport};

pub(crate) const PLACEHOLDER_SIZE: u32 = 64;

pub(crate) fn is_replaceable_part(path: &str) -> bool {
    let extension = path.rsplit_once('.').map(|(_, extension)| extension);
    matches!(
        extension,
        Some("png" | "jpg" | "jpeg" | "gif" | "bmp" | "tif" | "tiff" | "svg" | "emf" | "wmf")
    ) || path.contains("/media/")
        || path.ends_with("/thumbnail")
}

/// Redact a media part by emitting a fixed-size blank placeholder in the
/// part's OWN format so `[Content_Types].xml` declarations stay consistent.
/// WMF/EMF parts become minimal valid blank metafile stubs.
pub(crate) fn replace_media(
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    let lower = path.to_ascii_lowercase();
    let extension = lower.rsplit_once('.').map(|(_, ext)| ext);
    let result = match extension {
        Some("png") => solid_placeholder(path, ImageFormat::Png)?,
        Some("jpg" | "jpeg") => solid_placeholder(path, ImageFormat::Jpeg)?,
        Some("gif") => solid_placeholder(path, ImageFormat::Gif)?,
        Some("bmp") => solid_placeholder(path, ImageFormat::Bmp)?,
        Some("tif" | "tiff") => solid_placeholder(path, ImageFormat::Tiff)?,
        Some("svg") => {
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="#aaa"/></svg>"##.to_vec()
        }
        Some("emf") => emf_stub(),
        Some("wmf") => wmf_stub(),
        _ => match image::guess_format(bytes).ok().filter(is_encodable) {
            Some(format) => solid_placeholder(path, format)?,
            None => {
                return Err(RedactError::Image {
                    part: path.to_owned(),
                    message: "unsupported media format cannot be safely redacted".to_owned(),
                });
            }
        },
    };
    report.media_parts += 1;
    Ok(result)
}

fn is_encodable(format: &ImageFormat) -> bool {
    matches!(
        format,
        ImageFormat::Png
            | ImageFormat::Jpeg
            | ImageFormat::Gif
            | ImageFormat::Bmp
            | ImageFormat::Tiff
    )
}

fn solid_placeholder(path: &str, format: ImageFormat) -> Result<Vec<u8>, RedactError> {
    let image = DynamicImage::ImageRgb8(ImageBuffer::from_pixel(
        PLACEHOLDER_SIZE,
        PLACEHOLDER_SIZE,
        Rgb([170, 170, 170]),
    ));
    let mut output = Cursor::new(Vec::new());
    image
        .write_to(&mut output, format)
        .map_err(|error| RedactError::Image {
            part: path.to_owned(),
            message: error.to_string(),
        })?;
    Ok(output.into_inner())
}

fn emf_stub() -> Vec<u8> {
    let mut out = Vec::with_capacity(108);
    out.extend_from_slice(&1u32.to_le_bytes()); // iType = EMR_HEADER
    out.extend_from_slice(&88u32.to_le_bytes()); // nSize
    out.extend_from_slice(&[0; 16]); // rclBounds
    out.extend_from_slice(&0u32.to_le_bytes()); // rclFrame left/top
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&50_800u32.to_le_bytes()); // rclFrame right/bottom, 0.01 mm
    out.extend_from_slice(&28_575u32.to_le_bytes());
    out.extend_from_slice(&0x464D4520u32.to_le_bytes()); // " EMF"
    out.extend_from_slice(&0x00010000u32.to_le_bytes()); // version 1.0
    out.extend_from_slice(&108u32.to_le_bytes()); // nBytes
    out.extend_from_slice(&2u32.to_le_bytes()); // nRecords
    out.extend_from_slice(&1u16.to_le_bytes()); // nHandles, GDI reserves index 0
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out.extend_from_slice(&0u32.to_le_bytes()); // nDescription
    out.extend_from_slice(&0u32.to_le_bytes()); // offDescription
    out.extend_from_slice(&0u32.to_le_bytes()); // nPalEntries
    out.extend_from_slice(&1920i32.to_le_bytes()); // szlDevice px
    out.extend_from_slice(&1080i32.to_le_bytes());
    out.extend_from_slice(&508i32.to_le_bytes()); // szlMillimeters
    out.extend_from_slice(&285i32.to_le_bytes());
    out.extend_from_slice(&14u32.to_le_bytes()); // iType = EMR_EOF
    out.extend_from_slice(&20u32.to_le_bytes()); // nSize
    out.extend_from_slice(&0u32.to_le_bytes()); // nPalEntries
    out.extend_from_slice(&16u32.to_le_bytes()); // offPalEntries, no palette
    out.extend_from_slice(&20u32.to_le_bytes()); // nSizeLast duplicates nSize
    debug_assert_eq!(out.len(), 108);
    out
}

fn wmf_stub() -> Vec<u8> {
    let mut out = Vec::with_capacity(46);
    out.extend_from_slice(&0x9AC6CDD7u32.to_le_bytes()); // placeable key
    out.extend_from_slice(&0u16.to_le_bytes()); // hmf
    out.extend_from_slice(&0u16.to_le_bytes()); // bbox left
    out.extend_from_slice(&0u16.to_le_bytes()); // bbox top
    out.extend_from_slice(&2000u16.to_le_bytes()); // bbox right
    out.extend_from_slice(&2000u16.to_le_bytes()); // bbox bottom
    out.extend_from_slice(&1440u16.to_le_bytes()); // inch (twips)
    out.extend_from_slice(&[0; 4]); // reserved
    let checksum = out
        .as_chunks::<2>()
        .0
        .iter()
        .fold(0u16, |acc, chunk| acc ^ u16::from_le_bytes(*chunk));
    out.extend_from_slice(&checksum.to_le_bytes());
    out.extend_from_slice(&1u16.to_le_bytes()); // mtType = memory
    out.extend_from_slice(&9u16.to_le_bytes()); // mtHeaderSize words
    out.extend_from_slice(&0x0300u16.to_le_bytes()); // version 3.0
    out.extend_from_slice(&12u32.to_le_bytes()); // mtSize words, excludes placeable header
    out.extend_from_slice(&0u16.to_le_bytes()); // mtNoObjects
    out.extend_from_slice(&3u32.to_le_bytes()); // mtMaxRecord
    out.extend_from_slice(&0u16.to_le_bytes()); // mtNoParameters
    out.extend_from_slice(&3u32.to_le_bytes()); // META_EOF RecordSize words
    out.extend_from_slice(&0u16.to_le_bytes()); // META_EOF RecordFunction
    debug_assert_eq!(out.len(), 46);
    out
}
