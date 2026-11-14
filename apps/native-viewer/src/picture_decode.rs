use std::io::Cursor;

use vello::peniko::{Blob, ImageAlphaType, ImageData, ImageFormat};

/// The raster backend's per-picture pixel limit (`docx_raster::MAX_IMAGE_PIXELS`).
pub(crate) const MAX_IMAGE_PIXELS: u64 = 33_554_432;
/// The raster backend's per-picture byte limit (`docx_raster::MAX_IMAGE_BYTES`).
pub(crate) const MAX_IMAGE_BYTES: u64 = 268_435_456;

/// One picture decoded to RGBA within the raster backend's per-image limits,
/// charging its pixels to `budget` before anything is allocated.
pub(crate) fn decode_bounded(
    bytes: &[u8],
    orient: bool,
    budget: &mut u64,
) -> Result<ImageData, String> {
    use image::ImageDecoder as _;

    let reader = image::ImageReader::new(Cursor::new(bytes))
        .with_guessed_format()
        .map_err(|error| error.to_string())?;
    if reader.format() == Some(image::ImageFormat::Tiff)
        && ooxml_drawingml::media::tiff_bounded_pixels(bytes).is_none()
    {
        return Err("unsupported TIFF encoding".to_owned());
    }
    let mut decoder = reader.into_decoder().map_err(|error| error.to_string())?;
    let (declared_width, declared_height) = decoder.dimensions();
    let declared = u64::from(declared_width) * u64::from(declared_height);
    if declared > MAX_IMAGE_PIXELS
        || decoder
            .total_bytes()
            .saturating_add(declared.saturating_mul(4))
            > MAX_IMAGE_BYTES
    {
        return Err(format!(
            "{declared_width}x{declared_height} exceeds the picture size limit"
        ));
    }
    let pixels = declared.saturating_mul(4);
    if pixels > *budget {
        return Err("the document's picture budget is spent".to_owned());
    }
    *budget -= pixels;
    let orientation = if orient {
        Some(decoder.orientation().map_err(|error| error.to_string())?)
    } else {
        None
    };
    let mut image =
        image::DynamicImage::from_decoder(decoder).map_err(|error| error.to_string())?;
    if let Some(orientation) = orientation {
        image.apply_orientation(orientation);
    }
    let rgba = image.into_rgba8();
    let (width, height) = rgba.dimensions();
    Ok(ImageData {
        data: Blob::from(rgba.into_raw()),
        format: ImageFormat::Rgba8,
        alpha_type: ImageAlphaType::Alpha,
        width,
        height,
    })
}

#[cfg(all(test, feature = "docx"))]
mod tests {
    #[test]
    fn limits_match_the_raster_backend() {
        assert_eq!(super::MAX_IMAGE_PIXELS, docx_raster::MAX_IMAGE_PIXELS);
        assert_eq!(super::MAX_IMAGE_BYTES, docx_raster::MAX_IMAGE_BYTES);
    }
}
