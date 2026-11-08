//! Pictures past a backend's pixel cap, decoded at a reduced size instead of
//! refused. A JPEG is decoded at a reduced IDCT scale, so its full size is
//! never allocated; a PNG is box filtered row by row as it streams. Either
//! way the result fits the cap, or there is no result.

use std::io::Cursor;

use std::sync::atomic::{AtomicU64, Ordering};

use image::{DynamicImage, ImageDecoder, RgbaImage};

/// Pictures decoded through [`decode_reduced`], for tests.
#[doc(hidden)]
pub static REDUCED_DECODE_COUNT: AtomicU64 = AtomicU64::new(0);

/// The longest side a reduced picture is drawn at: the same bound the canvas
/// engine downscales oversized pictures to, so both engines show it at the
/// same resolution.
pub const REDUCED_MAX_SIDE: u32 = 4096;

/// How many times `max_pixels` a source may declare and still be reduced.
/// Reducing reads the whole source, so a larger one is refused as before.
pub const MAX_SOURCE_FACTOR: u64 = 2;

/// The longest side a reduced source may declare, JPEG's own limit. A PNG
/// row buffer grows with its width, so a wider source is refused.
pub const MAX_SOURCE_SIDE: u32 = 65_535;

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
    if width == 0
        || height == 0
        || longer > MAX_SOURCE_SIDE
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

/// What [`decode_reduced`] costs for a `width` x `height` source: the source
/// pixels it reads, charged as decode work the way a whole decode is, and the
/// bytes it may hold. Those are the reduced raster and one rotated copy of it,
/// the box filter's row of sums and counts, an IDCT intermediate of at most
/// four times its pixels (one scale step above it) and never past
/// `max_pixels`, and what the source format holds at its own size: a
/// progressive JPEG's coefficients, collected before any IDCT, or a PNG's
/// raw, previous and expanded rows.
pub fn reduced_cost(bytes: &[u8], width: u32, height: u32, max_pixels: u64) -> Option<(u64, u64)> {
    let (target_width, target_height) = reduced_size(width, height, max_pixels)?;
    let pixels = u64::from(target_width) * u64::from(target_height);
    let intermediate = pixels.saturating_mul(4).min(max_pixels);
    let source = if bytes.starts_with(&[0xFF, 0xD8]) {
        let frame = jpeg_frame(bytes)?;
        if frame.progressive {
            frame.coefficient_bytes()
        } else {
            0
        }
    } else {
        png_row_bytes(bytes, width)?
    };
    Some((
        u64::from(width) * u64::from(height),
        pixels
            .saturating_mul(8)
            .saturating_add(u64::from(target_width) * 36)
            .saturating_add(intermediate.saturating_mul(3))
            .saturating_add(source),
    ))
}

/// The bytes a streaming PNG decode holds at the source width: the raw row
/// and the previous one it unfilters against, plus the row expanded to 8-bit
/// RGBA. `None` when the header is not a PNG one.
fn png_row_bytes(bytes: &[u8], width: u32) -> Option<u64> {
    if !bytes.starts_with(b"\x89PNG\r\n\x1a\n") || bytes.get(12..16)? != b"IHDR" {
        return None;
    }
    let depth = u64::from(*bytes.get(24)?);
    let channels = match *bytes.get(25)? {
        0 | 3 => 1,
        2 => 3,
        4 => 2,
        6 => 4,
        _ => return None,
    };
    let raw = (u64::from(width) * channels * depth).div_ceil(8) + 1;
    Some(raw * 2 + u64::from(width) * 4)
}

/// A JPEG frame header: its size, whether it is progressive, and each
/// component's sampling factors.
struct JpegFrame {
    width: u32,
    height: u32,
    progressive: bool,
    sampling: Vec<(u32, u32)>,
}

impl JpegFrame {
    /// The bytes of the 16-bit DCT coefficients a progressive decode holds for
    /// every block of every component, padded to whole MCUs.
    fn coefficient_bytes(&self) -> u64 {
        let horizontal = self
            .sampling
            .iter()
            .map(|(h, _)| *h)
            .max()
            .unwrap_or(1)
            .max(1);
        let vertical = self
            .sampling
            .iter()
            .map(|(_, v)| *v)
            .max()
            .unwrap_or(1)
            .max(1);
        let across = u64::from(self.width.div_ceil(8 * horizontal));
        let down = u64::from(self.height.div_ceil(8 * vertical));
        self.sampling
            .iter()
            .map(|(h, v)| across * u64::from(*h) * down * u64::from(*v) * 64 * 2)
            .sum()
    }
}

fn jpeg_frame(bytes: &[u8]) -> Option<JpegFrame> {
    let mut at = 2;
    while at + 4 <= bytes.len() {
        if bytes[at] != 0xFF {
            return None;
        }
        let marker = bytes[at + 1];
        if marker == 0xFF {
            at += 1;
            continue;
        }
        if marker == 0x01 || (0xD0..=0xD7).contains(&marker) {
            at += 2;
            continue;
        }
        let length = usize::from(u16::from_be_bytes([bytes[at + 2], bytes[at + 3]]));
        if (0xC0..=0xCF).contains(&marker) && !matches!(marker, 0xC4 | 0xC8 | 0xCC) {
            let frame = bytes.get(at + 4..at + 2 + length)?;
            let count = usize::from(*frame.get(5)?);
            let sampling = (0..count)
                .map(|index| {
                    let factors = *frame.get(6 + index * 3 + 1)?;
                    Some((
                        u32::from(factors >> 4).max(1),
                        u32::from(factors & 0x0F).max(1),
                    ))
                })
                .collect::<Option<Vec<_>>>()?;
            return Some(JpegFrame {
                height: u32::from(u16::from_be_bytes([*frame.get(1)?, *frame.get(2)?])),
                width: u32::from(u16::from_be_bytes([*frame.get(3)?, *frame.get(4)?])),
                progressive: matches!(marker, 0xC2 | 0xC6 | 0xCA | 0xCE),
                sampling,
            });
        }
        at += 2 + length;
    }
    None
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
    let samples = match decoder.info()?.pixel_format {
        jpeg_decoder::PixelFormat::RGB24 => 3,
        jpeg_decoder::PixelFormat::L8 => 1,
        _ => return None,
    };
    let orientation = image::codecs::jpeg::JpegDecoder::new(Cursor::new(bytes))
        .and_then(|mut header| header.orientation())
        .ok()?;
    let pixels = decoder.decode().ok()?;
    let (out_width, out_height) = (target_width.min(width), target_height.min(height));
    let mut reducer = BoxReducer::new(width, height, out_width, out_height);
    let stride = width as usize * samples;
    if pixels.len() < stride * height as usize {
        return None;
    }
    for row in pixels.chunks_exact(stride).take(height as usize) {
        reducer.push_row(row, samples)?;
    }
    drop(pixels);
    let mut reduced = DynamicImage::ImageRgba8(reducer.finish()?);
    reduced.apply_orientation(orientation);
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
    let mut reducer = BoxReducer::new(width, height, target_width, target_height);
    while let Some(row) = reader.next_row().ok()? {
        reducer.push_row(row.data(), samples)?;
    }
    reducer.finish()
}

/// An alpha-weighted box filter fed one source row at a time: it holds one
/// row of sums and the reduced raster, never the source.
struct BoxReducer {
    width: u32,
    height: u32,
    target_width: u32,
    target_height: u32,
    sums: Vec<f64>,
    counts: Vec<u32>,
    output: Vec<u8>,
    current: usize,
    rows: u32,
}

impl BoxReducer {
    fn new(width: u32, height: u32, target_width: u32, target_height: u32) -> Self {
        Self {
            width,
            height,
            target_width,
            target_height,
            sums: vec![0.0; target_width as usize * 4],
            counts: vec![0; target_width as usize],
            output: vec![0; target_width as usize * target_height as usize * 4],
            current: 0,
            rows: 0,
        }
    }

    /// Adds the next source row of 8-bit gray, gray-alpha, RGB or RGBA pixels.
    fn push_row(&mut self, row: &[u8], samples: usize) -> Option<()> {
        let target_row = (u64::from(self.rows) * u64::from(self.target_height)
            / u64::from(self.height))
        .min(u64::from(self.target_height) - 1) as usize;
        if target_row != self.current {
            self.flush();
            self.current = target_row;
        }
        for (x, pixel) in row
            .chunks_exact(samples)
            .take(self.width as usize)
            .enumerate()
        {
            let (red, green, blue, alpha) = match *pixel {
                [gray] => (gray, gray, gray, 255),
                [gray, alpha] => (gray, gray, gray, alpha),
                [red, green, blue] => (red, green, blue, 255),
                [red, green, blue, alpha] => (red, green, blue, alpha),
                _ => return None,
            };
            let column = (x as u64 * u64::from(self.target_width) / u64::from(self.width)) as usize;
            let at = column * 4;
            let weight = f64::from(alpha);
            self.sums[at] += f64::from(red) * weight;
            self.sums[at + 1] += f64::from(green) * weight;
            self.sums[at + 2] += f64::from(blue) * weight;
            self.sums[at + 3] += weight;
            self.counts[column] += 1;
        }
        self.rows += 1;
        Some(())
    }

    /// The reduced raster, or `None` when the source ended early.
    fn finish(mut self) -> Option<RgbaImage> {
        if self.rows != self.height {
            return None;
        }
        self.flush();
        RgbaImage::from_raw(self.target_width, self.target_height, self.output)
    }

    /// Writes the current reduced row from its sums and clears them.
    fn flush(&mut self) {
        let row_bytes = self.target_width as usize * 4;
        let row = &mut self.output[self.current * row_bytes..][..row_bytes];
        for (column, count) in self.counts.iter_mut().enumerate() {
            let at = column * 4;
            let alpha = self.sums[at + 3];
            if *count > 0 && alpha > 0.0 {
                for channel in 0..3 {
                    row[at + channel] = (self.sums[at + channel] / alpha).round() as u8;
                }
                row[at + 3] = (alpha / f64::from(*count)).round() as u8;
            }
            self.sums[at..at + 4].fill(0.0);
            *count = 0;
        }
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
        let (pixels, width, height) = decode_reduced(&jpeg_of(640, 480), 160_000).expect("reduced");
        assert!(u64::from(width) * u64::from(height) <= 160_000);
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
        assert_eq!(reduced_cost(&bytes, 60_000, 60_000, CAP), None);
    }

    #[test]
    fn a_progressive_jpeg_is_charged_its_source_sized_coefficients() {
        let mut bytes = jpeg_of(16, 16);
        let at = bytes
            .windows(2)
            .position(|marker| marker == [0xFF, 0xC0])
            .expect("frame header");
        bytes[at + 1] = 0xC2;
        bytes[at + 5..at + 9].copy_from_slice(&8000u16.to_be_bytes().repeat(2));
        let (pixels, cost) = reduced_cost(&bytes, 8000, 8000, 40_000_000).expect("cost");
        assert_eq!(pixels, 64_000_000);
        let frame = jpeg_frame(&bytes).expect("frame");
        assert!(frame.progressive);
        assert!(cost >= frame.coefficient_bytes());
        assert!(frame.coefficient_bytes() >= 64_000_000 * 2);
    }

    #[test]
    fn a_source_wider_than_a_jpeg_may_be_is_refused_from_its_header() {
        assert_eq!(reduced_size(MAX_SOURCE_SIDE + 1, 1, CAP), None);
        assert!(reduced_size(MAX_SOURCE_SIDE, 2, CAP).is_some());
    }

    #[test]
    fn a_png_past_its_budget_streams_into_a_reduced_raster() {
        let (pixels, width, height) = decode_reduced(&png_of(400, 300), 60_000).expect("reduced");
        assert!(u64::from(width) * u64::from(height) <= 60_000);
        let (whole, rest) = pixels.as_chunks::<4>();
        assert!(rest.is_empty());
        assert!(whole.iter().all(|pixel| *pixel == [0x11, 0x66, 0xcc, 0xff]));
    }
}
