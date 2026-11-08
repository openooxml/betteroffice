//! Pictures past a backend's pixel cap, decoded at a reduced size instead of
//! refused. A JPEG is decoded at a reduced IDCT scale, so its full size is
//! never allocated; a PNG is box filtered row by row as it streams. Either
//! way the result fits the cap, or there is no result.

use std::io::Cursor;

use std::sync::atomic::{AtomicU64, Ordering};

use image::{DynamicImage, ImageDecoder, RgbaImage, imageops::FilterType};

/// Pictures decoded through [`decode_reduced`], for tests.
#[doc(hidden)]
pub static REDUCED_DECODE_COUNT: AtomicU64 = AtomicU64::new(0);

/// The longest side a reduced picture is drawn at: the same bound the canvas
/// engine downscales oversized pictures to, so both engines show it at the
/// same resolution.
pub const REDUCED_MAX_SIDE: u32 = 4096;

/// How many times `max_pixels` a source may declare and still be reduced.
/// Reducing reads the whole source, so a larger one is refused as before.
pub const MAX_SOURCE_FACTOR: u64 = 4;

/// Straight RGBA of a JPEG or PNG whose declared size is past `max_pixels`,
/// reduced to [`REDUCED_MAX_SIDE`] on its longer side, or less where that
/// still would not fit. `None` when no reduction fits, or the picture is not
/// a kind this can reduce while decoding.
pub fn decode_reduced(bytes: &[u8], max_pixels: u64) -> Option<(Vec<u8>, u32, u32)> {
    REDUCED_DECODE_COUNT.fetch_add(1, Ordering::Relaxed);
    let image = if bytes.starts_with(&[0xFF, 0xD8]) {
        jpeg(bytes, max_pixels)?
    } else if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        png(bytes, max_pixels)?
    } else {
        return None;
    };
    let (width, height) = (image.width(), image.height());
    if u64::from(width) * u64::from(height) > max_pixels {
        return None;
    }
    Some((image.into_raw(), width, height))
}

/// The size a `width` x `height` picture is reduced to: the canvas engine's
/// rule (scale the longer side to [`REDUCED_MAX_SIDE`], flooring each side),
/// shrunk further only where that would not fit `max_pixels`. `None` for a
/// source past [`MAX_SOURCE_FACTOR`] times `max_pixels`.
pub fn reduced_size(width: u32, height: u32, max_pixels: u64) -> Option<(u32, u32)> {
    let longer = width.max(height);
    if longer == 0
        || max_pixels == 0
        || u64::from(width) * u64::from(height) > max_pixels.saturating_mul(MAX_SOURCE_FACTOR)
    {
        return None;
    }
    let size = |factor: f64| {
        (
            ((f64::from(width) * factor).floor() as u32).max(1),
            ((f64::from(height) * factor).floor() as u32).max(1),
        )
    };
    let mut factor = if longer > REDUCED_MAX_SIDE {
        f64::from(REDUCED_MAX_SIDE) / f64::from(longer)
    } else {
        1.0
    };
    let (mut target_width, mut target_height) = size(factor);
    while u64::from(target_width) * u64::from(target_height) > max_pixels {
        factor *= 0.95;
        (target_width, target_height) = size(factor);
    }
    Some((target_width, target_height))
}

/// The pixels and bytes [`decode_reduced`] may hold for a `width` x
/// `height` picture: the reduced raster, plus an IDCT intermediate of at most
/// four times its pixels (one scale step above it), never past `max_pixels`.
pub fn reduced_cost(width: u32, height: u32, max_pixels: u64) -> Option<(u64, u64)> {
    let (target_width, target_height) = reduced_size(width, height, max_pixels)?;
    let pixels = u64::from(target_width) * u64::from(target_height);
    let intermediate = pixels.saturating_mul(4).min(max_pixels);
    Some((
        pixels,
        pixels
            .saturating_mul(4)
            .saturating_add(intermediate.saturating_mul(3)),
    ))
}

fn jpeg(bytes: &[u8], max_pixels: u64) -> Option<RgbaImage> {
    let mut decoder = jpeg_decoder::Decoder::new(bytes);
    decoder.read_info().ok()?;
    let info = decoder.info()?;
    let (target_width, target_height) =
        reduced_size(u32::from(info.width), u32::from(info.height), max_pixels)?;
    let mut request = (target_width, target_height);
    let (width, height) = loop {
        let (width, height) = decoder
            .scale(
                u16::try_from(request.0).ok()?,
                u16::try_from(request.1).ok()?,
            )
            .ok()?;
        if u64::from(width) * u64::from(height) <= max_pixels {
            break (u32::from(width), u32::from(height));
        }
        if request == (1, 1) {
            return None;
        }
        request = ((request.0 / 2).max(1), (request.1 / 2).max(1));
    };
    let format = decoder.info()?.pixel_format;
    let pixels = decoder.decode().ok()?;
    let decoded = match format {
        jpeg_decoder::PixelFormat::RGB24 => {
            DynamicImage::ImageRgb8(image::RgbImage::from_raw(width, height, pixels)?)
        }
        jpeg_decoder::PixelFormat::L8 => {
            DynamicImage::ImageLuma8(image::GrayImage::from_raw(width, height, pixels)?)
        }
        _ => return None,
    };
    let mut reduced = if width <= target_width && height <= target_height {
        decoded
    } else {
        decoded.resize_exact(target_width, target_height, FilterType::Triangle)
    };
    if let Ok(orientation) = image::codecs::jpeg::JpegDecoder::new(Cursor::new(bytes))
        .and_then(|mut header| header.orientation())
    {
        reduced.apply_orientation(orientation);
    }
    Some(reduced.into_rgba8())
}

fn png(bytes: &[u8], max_pixels: u64) -> Option<RgbaImage> {
    let mut decoder = png::Decoder::new(Cursor::new(bytes));
    decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::STRIP_16);
    let mut reader = decoder.read_info().ok()?;
    let info = reader.info();
    if info.interlaced {
        return None;
    }
    let (width, height) = (info.width, info.height);
    let (target_width, target_height) = reduced_size(width, height, max_pixels)?;
    let (color, depth) = reader.output_color_type();
    if depth != png::BitDepth::Eight {
        return None;
    }
    let samples = color.samples();
    let columns: Vec<usize> = (0..width)
        .map(|x| (u64::from(x) * u64::from(target_width) / u64::from(width)) as usize)
        .collect();
    let row_bytes = target_width as usize * 4;
    let mut sums = vec![0f64; row_bytes];
    let mut counts = vec![0u32; target_width as usize];
    let mut output = vec![0u8; row_bytes * target_height as usize];
    let mut current = 0usize;
    let mut y = 0u64;
    while let Some(row) = reader.next_row().ok()? {
        let target_row = (y * u64::from(target_height) / u64::from(height))
            .min(u64::from(target_height) - 1) as usize;
        if target_row != current {
            flush(
                &mut sums,
                &mut counts,
                &mut output[current * row_bytes..][..row_bytes],
            );
            current = target_row;
        }
        for (x, pixel) in row.data().chunks_exact(samples).enumerate() {
            let (red, green, blue, alpha) = match *pixel {
                [gray] => (gray, gray, gray, 255),
                [gray, alpha] => (gray, gray, gray, alpha),
                [red, green, blue] => (red, green, blue, 255),
                [red, green, blue, alpha] => (red, green, blue, alpha),
                _ => return None,
            };
            let column = *columns.get(x)?;
            let at = column * 4;
            let weight = f64::from(alpha);
            sums[at] += f64::from(red) * weight;
            sums[at + 1] += f64::from(green) * weight;
            sums[at + 2] += f64::from(blue) * weight;
            sums[at + 3] += weight;
            counts[column] += 1;
        }
        y += 1;
    }
    if y != u64::from(height) {
        return None;
    }
    flush(
        &mut sums,
        &mut counts,
        &mut output[current * row_bytes..][..row_bytes],
    );
    RgbaImage::from_raw(target_width, target_height, output)
}

/// Writes one reduced row from its alpha-weighted sums and clears them.
fn flush(sums: &mut [f64], counts: &mut [u32], row: &mut [u8]) {
    for (column, count) in counts.iter_mut().enumerate() {
        let at = column * 4;
        let alpha = sums[at + 3];
        if *count > 0 && alpha > 0.0 {
            for channel in 0..3 {
                row[at + channel] = (sums[at + channel] / alpha).round() as u8;
            }
            row[at + 3] = (alpha / f64::from(*count)).round() as u8;
        }
        sums[at..at + 4].fill(0.0);
        *count = 0;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const CAP: u64 = 33_554_432;

    fn jpeg_of(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = Vec::new();
        image::codecs::jpeg::JpegEncoder::new(&mut bytes)
            .encode_image(&image::RgbImage::from_pixel(
                width,
                height,
                image::Rgb([200, 40, 10]),
            ))
            .expect("jpeg");
        bytes
    }

    fn png_of(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut encoder = png::Encoder::new(&mut bytes, width, height);
            encoder.set_color(png::ColorType::Rgba);
            let mut writer = encoder.write_header().expect("png header");
            let mut stream = writer.stream_writer().expect("png stream");
            let row = [0x11_u8, 0x66, 0xcc, 0xff].repeat(width as usize);
            for _ in 0..height {
                std::io::Write::write_all(&mut stream, &row).expect("png row");
            }
            stream.finish().expect("png finish");
        }
        bytes
    }

    #[test]
    fn the_reduced_size_is_the_canvas_engines() {
        for ((width, height), expected) in [
            ((8590, 4613), (4096, 2199)),
            ((5750, 7795), (3021, 4096)),
            ((8022, 7594), (4096, 3877)),
            ((4958, 7017), (2894, 4096)),
        ] {
            assert_eq!(reduced_size(width, height, CAP), Some(expected));
        }
    }

    #[test]
    fn a_jpeg_past_its_budget_decodes_reduced_within_it() {
        let (pixels, width, height) = decode_reduced(&jpeg_of(640, 480), 80_000).expect("reduced");
        assert!(u64::from(width) * u64::from(height) <= 80_000);
        assert_eq!(pixels.len(), width as usize * height as usize * 4);
        assert!(pixels[0] > 150 && pixels[1] < 90 && pixels[3] == 255);
    }

    #[test]
    fn a_jpeg_whose_header_is_past_what_a_reduction_may_read_is_refused() {
        let mut bytes = jpeg_of(16, 16);
        let at = bytes
            .windows(2)
            .position(|marker| marker == [0xFF, 0xC0])
            .expect("frame header");
        bytes[at + 5..at + 9].copy_from_slice(&[0xEA, 0x60, 0xEA, 0x60]);
        assert!(decode_reduced(&bytes, CAP).is_none());
        assert_eq!(reduced_cost(60_000, 60_000, CAP), None);
    }

    #[test]
    fn a_png_past_its_budget_streams_into_a_reduced_raster() {
        let (pixels, width, height) = decode_reduced(&png_of(600, 400), 60_000).expect("reduced");
        assert!(u64::from(width) * u64::from(height) <= 60_000);
        let (whole, rest) = pixels.as_chunks::<4>();
        assert!(rest.is_empty());
        assert!(whole.iter().all(|pixel| *pixel == [0x11, 0x66, 0xcc, 0xff]));
    }
}
