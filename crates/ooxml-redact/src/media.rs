use std::io::Cursor;

use image::{DynamicImage, ImageBuffer, ImageFormat, Rgb};

use crate::{RedactError, RedactionReport};

const MAX_PLACEHOLDER_PIXELS: u64 = 64 * 1024 * 1024;
const FALLBACK_PNG: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0,
    0, 0, 144, 119, 83, 222, 0, 0, 0, 12, 73, 68, 65, 84, 8, 215, 99, 168, 168, 168, 0, 0, 2, 2, 1,
    0, 228, 51, 249, 227, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

pub(crate) fn is_replaceable_part(path: &str) -> bool {
    let extension = path.rsplit_once('.').map(|(_, extension)| extension);
    matches!(
        extension,
        Some("png" | "jpg" | "jpeg" | "gif" | "bmp" | "tif" | "tiff" | "svg" | "emf" | "wmf")
    ) || path.contains("/media/")
        || path.ends_with("/thumbnail")
}

pub(crate) fn replace_media(
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    let lower = path.to_ascii_lowercase();
    let result = if lower.ends_with(".png") {
        solid_image(path, bytes, ImageFormat::Png)?
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        solid_image(path, bytes, ImageFormat::Jpeg)?
    } else if lower.ends_with(".svg") {
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="#aaa"/></svg>"##.to_vec()
    } else {
        FALLBACK_PNG.to_vec()
    };
    report.media_parts += 1;
    Ok(result)
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
