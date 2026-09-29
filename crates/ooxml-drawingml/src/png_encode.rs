//! Minimal PNG writer for decoded rasters.

use std::io::Write;

/// RGBA8 pixels to PNG bytes.
pub fn encode_rgba8(rgba: &[u8], width: u32, height: u32) -> Result<Vec<u8>, String> {
    encode(rgba, width, height, 4)
}

/// RGB8 pixels to PNG bytes.
pub fn encode_rgb8(rgb: &[u8], width: u32, height: u32) -> Result<Vec<u8>, String> {
    encode(rgb, width, height, 3)
}

fn encode(pixels: &[u8], width: u32, height: u32, channels: usize) -> Result<Vec<u8>, String> {
    let stride = (width as usize)
        .checked_mul(channels)
        .filter(|stride| *stride > 0)
        .ok_or("Malformed raster data")?;
    if height == 0 || pixels.len() != stride.saturating_mul(height as usize) {
        return Err("Malformed raster data".to_owned());
    }
    let mut png = Vec::new();
    png.extend_from_slice(&[137, 80, 78, 71, 13, 10, 26, 10]);
    let mut header = [0u8; 13];
    header[0..4].copy_from_slice(&width.to_be_bytes());
    header[4..8].copy_from_slice(&height.to_be_bytes());
    header[8] = 8;
    header[9] = if channels == 4 { 6 } else { 2 };
    chunk(&mut png, b"IHDR", &header);
    let mut encoder = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
    for row in pixels.chunks_exact(stride) {
        encoder
            .write_all(&[0])
            .and_then(|()| encoder.write_all(row))
            .map_err(|error| error.to_string())?;
    }
    let idat = encoder.finish().map_err(|error| error.to_string())?;
    chunk(&mut png, b"IDAT", &idat);
    chunk(&mut png, b"IEND", &[]);
    Ok(png)
}

fn chunk(png: &mut Vec<u8>, tag: &[u8; 4], data: &[u8]) {
    png.extend_from_slice(&(data.len() as u32).to_be_bytes());
    png.extend_from_slice(tag);
    png.extend_from_slice(data);
    let mut crc = crc32fast::Hasher::new();
    crc.update(tag);
    crc.update(data);
    png.extend_from_slice(&crc.finalize().to_be_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rgb_and_rgba_rasters_decode_back_to_their_pixels() {
        for (pixels, channels, color) in [
            (vec![255, 0, 0, 0, 0, 255], 3, png::ColorType::Rgb),
            (
                vec![255, 0, 0, 128, 0, 0, 255, 255],
                4,
                png::ColorType::Rgba,
            ),
        ] {
            let bytes = encode(&pixels, 2, 1, channels).unwrap();
            let mut reader = png::Decoder::new(std::io::Cursor::new(bytes))
                .read_info()
                .unwrap();
            let mut out = vec![0; reader.output_buffer_size().unwrap()];
            let info = reader.next_frame(&mut out).unwrap();
            assert_eq!((info.width, info.height, info.color_type), (2, 1, color));
            assert_eq!(out, pixels);
        }
    }

    #[test]
    fn a_buffer_that_does_not_match_its_size_is_refused() {
        assert!(encode_rgb8(&[0; 5], 2, 1).is_err());
        assert!(encode_rgba8(&[], 0, 0).is_err());
    }
}
