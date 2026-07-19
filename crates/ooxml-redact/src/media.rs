use std::io::Cursor;

use image::{DynamicImage, ImageBuffer, ImageFormat, Rgb};

use crate::{RedactError, RedactionReport};

const MAX_PLACEHOLDER_PIXELS: u64 = 64 * 1024 * 1024;

pub(crate) fn is_replaceable_part(path: &str) -> bool {
    let extension = path.rsplit_once('.').map(|(_, extension)| extension);
    matches!(
        extension,
        Some("png" | "jpg" | "jpeg" | "gif" | "bmp" | "tif" | "tiff" | "svg" | "emf" | "wmf")
    ) || path.contains("/media/")
        || path.ends_with("/thumbnail")
}

/// Redact a media part, encoding the placeholder in the part's OWN format so the
/// `[Content_Types].xml` declaration stays consistent. Formats that cannot be
/// faithfully re-encoded (vector metafiles, non-image embeds) fail rather than
/// emit format-mismatched bytes.
pub(crate) fn replace_media(
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    let lower = path.to_ascii_lowercase();
    let extension = lower.rsplit_once('.').map(|(_, ext)| ext);
    let result = match extension {
        Some("png") => solid_image(path, bytes, ImageFormat::Png)?,
        Some("jpg" | "jpeg") => solid_image(path, bytes, ImageFormat::Jpeg)?,
        Some("gif") => solid_image(path, bytes, ImageFormat::Gif)?,
        Some("bmp") => solid_image(path, bytes, ImageFormat::Bmp)?,
        Some("tif" | "tiff") => solid_image(path, bytes, ImageFormat::Tiff)?,
        Some("svg") => {
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="#aaa"/></svg>"##.to_vec()
        }
        _ => match image::guess_format(bytes).ok().filter(is_encodable) {
            Some(format) => solid_image(path, bytes, format)?,
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

fn solid_image(path: &str, bytes: &[u8], format: ImageFormat) -> Result<Vec<u8>, RedactError> {
    let reader = image::ImageReader::with_format(Cursor::new(bytes), format);
    let (width, height) = reader
        .into_dimensions()
        .map_err(|error| RedactError::Image {
            part: path.to_owned(),
            message: error.to_string(),
        })?;
    let pixels = u64::from(width) * u64::from(height);
    if pixels > MAX_PLACEHOLDER_PIXELS {
        return Err(RedactError::Image {
            part: path.to_owned(),
            message: format!("{width}x{height} exceeds placeholder pixel limit"),
        });
    }

    let image =
        DynamicImage::ImageRgb8(ImageBuffer::from_pixel(width, height, Rgb([170, 170, 170])));
    let mut output = Cursor::new(Vec::new());
    image
        .write_to(&mut output, format)
        .map_err(|error| RedactError::Image {
            part: path.to_owned(),
            message: error.to_string(),
        })?;
    Ok(output.into_inner())
}
