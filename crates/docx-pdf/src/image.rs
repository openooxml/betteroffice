//! Decoded image data for PDF XObjects.

use base64::Engine;

#[derive(Debug)]
pub struct DecodedImage {
    pub width: i32,
    pub height: i32,
    pub rgb: Vec<u8>,
    pub alpha: Option<Vec<u8>>,
}

impl DecodedImage {
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        let image = image::load_from_memory(bytes).map_err(|error| error.to_string())?;
        let image = image.to_rgba8();
        let (width, height) = image.dimensions();
        let mut rgb = Vec::with_capacity(width as usize * height as usize * 3);
        let mut alpha = Vec::with_capacity(width as usize * height as usize);
        let mut has_alpha = false;
        for pixel in image.pixels() {
            rgb.extend_from_slice(&pixel.0[..3]);
            alpha.push(pixel.0[3]);
            has_alpha |= pixel.0[3] != 255;
        }
        Ok(Self {
            width: i32::try_from(width).map_err(|_| "image width is too large")?,
            height: i32::try_from(height).map_err(|_| "image height is too large")?,
            rgb,
            alpha: has_alpha.then_some(alpha),
        })
    }
}

/// Decode a `data:` URL into its payload.
pub fn data_url_bytes(value: &str) -> Option<Vec<u8>> {
    let (metadata, payload) = value.strip_prefix("data:")?.split_once(',')?;
    if metadata.contains(";base64") {
        base64::engine::general_purpose::STANDARD
            .decode(payload)
            .ok()
    } else {
        Some(payload.as_bytes().to_vec())
    }
}
