use std::io::Cursor;

use image::codecs::bmp::BmpEncoder;
use image::codecs::gif::GifEncoder;
use image::codecs::jpeg::JpegEncoder;
use image::codecs::png::{CompressionType, FilterType, PngEncoder};
use image::codecs::tiff::TiffEncoder;
use image::{
    ColorType, DynamicImage, ImageBuffer, ImageDecoder as _, ImageEncoder as _, ImageFormat,
    ImageReader, Rgb,
};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use crate::{RedactError, RedactionOptions, RedactionReport};

pub(crate) const PLACEHOLDER_SIZE: u32 = 64;
pub(crate) const MAX_PRESERVED_DIM: u32 = 4096;
const MAX_PRESERVED_PIXELS: u64 = 16_777_216;
const MAX_PRESERVED_BYTES: usize = 32 * 1024 * 1024;
const MAX_SVG_PRESERVED_BYTES: usize = 4 * 1024 * 1024;

/// Remaining uncompressed output bytes across all package parts.
pub(crate) struct MediaBudget {
    pub(crate) remaining: u64,
}

impl Default for MediaBudget {
    fn default() -> Self {
        Self {
            remaining: ooxml_opc::MAX_TOTAL_UNCOMPRESSED_BYTES,
        }
    }
}

impl MediaBudget {
    pub(crate) fn reserve(&mut self, path: &str, bytes: u64) -> Result<(), RedactError> {
        self.remaining = self
            .remaining
            .checked_sub(bytes)
            .ok_or_else(|| RedactError::Image {
                part: path.to_owned(),
                message: "size-preserving placeholders would exceed the package size limit"
                    .to_owned(),
            })?;
        Ok(())
    }

    /// Reconciles a reservation with the bytes actually produced, before they are retained.
    fn settle(&mut self, path: &str, reserved: u64, actual: usize) -> Result<(), RedactError> {
        let actual = actual as u64;
        if actual > reserved {
            self.reserve(path, actual - reserved)
        } else {
            self.remaining += reserved - actual;
            Ok(())
        }
    }

    fn charge<F>(&mut self, path: &str, reserved: u64, encode: F) -> Result<Vec<u8>, RedactError>
    where
        F: FnOnce() -> Result<Vec<u8>, RedactError>,
    {
        self.reserve(path, reserved)?;
        let out = encode()?;
        self.settle(path, reserved, out.len())?;
        Ok(out)
    }
}

pub(crate) fn is_replaceable_part(path: &str) -> bool {
    let extension = path.rsplit_once('.').map(|(_, extension)| extension);
    matches!(
        extension,
        Some(
            "png" | "jpg" | "jpeg" | "gif" | "bmp" | "tif" | "tiff" | "svg" | "emf" | "wmf" | "wdp"
        )
    ) || path.contains("/media/")
        || path.ends_with("/thumbnail")
}

#[cfg(test)]
pub(crate) fn replace_media(
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    replace_media_with_options(path, bytes, &RedactionOptions::default(), report)
}

#[cfg(test)]
pub(crate) fn replace_media_with_options(
    path: &str,
    bytes: &[u8],
    options: &RedactionOptions,
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    replace_media_budgeted(path, bytes, options, &mut MediaBudget::default(), report)
}

/// Redact a media part by emitting a blank or size-preserving synthetic
/// placeholder in the part's own format so `[Content_Types].xml` declarations
/// stay consistent.
pub(crate) fn replace_media_budgeted(
    path: &str,
    bytes: &[u8],
    options: &RedactionOptions,
    budget: &mut MediaBudget,
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    let lower = path.to_ascii_lowercase();
    let extension = lower.rsplit_once('.').map(|(_, ext)| ext);
    let result = if options.preserve_media_size {
        match extension {
            Some("png") => preserve_raster(path, bytes, ImageFormat::Png, budget)?,
            Some("jpg" | "jpeg") => preserve_raster(path, bytes, ImageFormat::Jpeg, budget)?,
            Some("gif") => preserve_raster(path, bytes, ImageFormat::Gif, budget)?,
            Some("bmp") => preserve_raster(path, bytes, ImageFormat::Bmp, budget)?,
            Some("tif" | "tiff") => preserve_raster(path, bytes, ImageFormat::Tiff, budget)?,
            Some("svg") => budget.charge(
                path,
                bytes.len().min(MAX_SVG_PRESERVED_BYTES) as u64,
                || preserve_svg(bytes),
            )?,
            Some("emf") => {
                budget.charge(path, bytes.len().min(MAX_PRESERVED_BYTES) as u64, || {
                    preserve_emf(bytes)
                })?
            }
            Some("wmf") => {
                budget.charge(path, bytes.len().min(MAX_PRESERVED_BYTES) as u64, || {
                    preserve_wmf(bytes)
                })?
            }
            _ => match image::guess_format(bytes).ok().filter(is_encodable) {
                Some(format) => preserve_raster(path, bytes, format, budget)?,
                None => {
                    return Err(RedactError::Image {
                        part: path.to_owned(),
                        message: "unsupported media format cannot be safely redacted".to_owned(),
                    });
                }
            },
        }
    } else {
        match extension {
            Some("png") => solid_placeholder(path, ImageFormat::Png)?,
            Some("jpg" | "jpeg") => solid_placeholder(path, ImageFormat::Jpeg)?,
            Some("gif") => solid_placeholder(path, ImageFormat::Gif)?,
            Some("bmp") => solid_placeholder(path, ImageFormat::Bmp)?,
            Some("tif" | "tiff") => solid_placeholder(path, ImageFormat::Tiff)?,
            Some("svg") => default_svg_stub(),
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
        }
    };
    if !options.preserve_media_size {
        budget.reserve(path, result.len() as u64)?;
    }
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

fn default_svg_stub() -> Vec<u8> {
    br##"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="#aaa"/></svg>"##.to_vec()
}

#[derive(Clone, Copy)]
struct SynthRng {
    state: u64,
}

impl SynthRng {
    fn new() -> Result<Self, RedactError> {
        let mut seed = [0u8; 8];
        getrandom::fill(&mut seed).map_err(|error| RedactError::Randomness(error.to_string()))?;
        let state = u64::from_le_bytes(seed) | 1;
        Ok(Self { state })
    }

    fn fork(&self, salt: u64) -> Self {
        let mut z = self
            .state
            .wrapping_add(salt.wrapping_mul(0x9E37_79B9_7F4A_7C15));
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        let state = (z ^ (z >> 31)) | 1;
        Self { state }
    }

    #[inline]
    fn next_u64(&mut self) -> u64 {
        let mut x = self.state;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.state = x;
        x.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn fill_bytes(&mut self, buf: &mut [u8]) {
        let (chunks, tail) = buf.as_chunks_mut::<8>();
        for chunk in chunks {
            *chunk = self.next_u64().to_le_bytes();
        }
        if !tail.is_empty() {
            let bytes = self.next_u64().to_le_bytes();
            tail.copy_from_slice(&bytes[..tail.len()]);
        }
    }
}

fn preserve_raster(
    path: &str,
    bytes: &[u8],
    format: ImageFormat,
    budget: &mut MediaBudget,
) -> Result<Vec<u8>, RedactError> {
    let Some((width, height, raw_color)) = inspect_raster_header(bytes, format) else {
        return budget.charge(path, 0, || solid_placeholder(path, format));
    };
    if width == 0
        || height == 0
        || width > MAX_PRESERVED_DIM
        || height > MAX_PRESERVED_DIM
        || u64::from(width) * u64::from(height) > MAX_PRESERVED_PIXELS
    {
        return budget.charge(path, 0, || solid_placeholder(path, format));
    }
    let color_type = target_color_type(format, raw_color);
    let target_len = bytes.len().clamp(64, MAX_PRESERVED_BYTES);
    let raw_len = u64::from(width) * u64::from(height) * bytes_per_pixel(color_type) as u64;
    // Uncompressed placeholders follow the declared pixels, not the input; the others aim at
    // the input's own size. Either way the actual output is settled before it is kept.
    let reserved = match format {
        ImageFormat::Bmp => raw_len,
        ImageFormat::Tiff if !tiff_compressed(target_len, raw_len) => raw_len,
        _ => target_len as u64,
    };
    let mut rng = SynthRng::new()?;
    budget.charge(path, reserved, || match format {
        ImageFormat::Png => encode_png_sized(path, width, height, color_type, target_len, &mut rng),
        ImageFormat::Jpeg => {
            encode_jpeg_sized(path, width, height, color_type, target_len, &mut rng)
        }
        ImageFormat::Gif => encode_gif_sized(path, width, height, target_len, &mut rng),
        ImageFormat::Bmp => encode_bmp_sized(path, width, height, color_type, &mut rng),
        ImageFormat::Tiff => {
            encode_tiff_sized(path, width, height, color_type, target_len, &mut rng)
        }
        _ => solid_placeholder(path, format),
    })
}

/// An input well below its raw pixel size was compressed; match it with Deflate.
fn tiff_compressed(target_len: usize, raw_len: u64) -> bool {
    (target_len as u64) < raw_len - raw_len / 8
}

fn inspect_raster_header(bytes: &[u8], format: ImageFormat) -> Option<(u32, u32, ColorType)> {
    if let Ok(decoder) = ImageReader::with_format(Cursor::new(bytes), format).into_decoder() {
        let (w, h) = decoder.dimensions();
        return Some((w, h, decoder.color_type()));
    }
    raw_header_info(bytes, format)
}

fn raw_header_info(bytes: &[u8], format: ImageFormat) -> Option<(u32, u32, ColorType)> {
    match format {
        ImageFormat::Png => {
            if bytes.len() < 29 || &bytes[0..8] != b"\x89PNG\r\n\x1a\n" || &bytes[12..16] != b"IHDR"
            {
                return None;
            }
            let width = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]);
            let height = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]);
            let bit_depth = bytes[24];
            let color_kind = bytes[25];
            let color = match (color_kind, bit_depth) {
                (0, 16) => ColorType::L16,
                (0, _) => ColorType::L8,
                (2, 16) => ColorType::Rgb16,
                (2 | 3, _) => ColorType::Rgb8,
                (4, 16) => ColorType::La16,
                (4, _) => ColorType::La8,
                (6, 16) => ColorType::Rgba16,
                (6, _) => ColorType::Rgba8,
                _ => return None,
            };
            Some((width, height, color))
        }
        ImageFormat::Jpeg => {
            if bytes.len() < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8 {
                return None;
            }
            let mut i = 2usize;
            while i + 1 < bytes.len() {
                if bytes[i] != 0xFF {
                    i += 1;
                    continue;
                }
                while i < bytes.len() && bytes[i] == 0xFF {
                    i += 1;
                }
                if i >= bytes.len() {
                    break;
                }
                let marker = bytes[i];
                i += 1;
                if marker == 0xD9 || marker == 0xDA {
                    break;
                }
                if marker == 0x00 || marker == 0x01 || (0xD0..=0xD8).contains(&marker) {
                    continue;
                }
                if i + 2 > bytes.len() {
                    break;
                }
                let seg_len = u16::from_be_bytes([bytes[i], bytes[i + 1]]) as usize;
                if seg_len < 2 {
                    break;
                }
                if matches!(
                    marker,
                    0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF
                ) && seg_len >= 8
                    && i + 8 <= bytes.len()
                {
                    let height = u32::from(u16::from_be_bytes([bytes[i + 3], bytes[i + 4]]));
                    let width = u32::from(u16::from_be_bytes([bytes[i + 5], bytes[i + 6]]));
                    let components = bytes[i + 7];
                    let color = if components == 1 {
                        ColorType::L8
                    } else {
                        ColorType::Rgb8
                    };
                    return Some((width, height, color));
                }
                i = i.saturating_add(seg_len);
            }
            None
        }
        ImageFormat::Gif => {
            if bytes.len() < 10 || (&bytes[0..6] != b"GIF87a" && &bytes[0..6] != b"GIF89a") {
                return None;
            }
            let width = u32::from(u16::from_le_bytes([bytes[6], bytes[7]]));
            let height = u32::from(u16::from_le_bytes([bytes[8], bytes[9]]));
            Some((width, height, ColorType::Rgba8))
        }
        ImageFormat::Bmp => {
            if bytes.len() < 26 || &bytes[0..2] != b"BM" {
                return None;
            }
            let dib_size = u32::from_le_bytes([bytes[14], bytes[15], bytes[16], bytes[17]]);
            if dib_size == 12 {
                let width = u32::from(u16::from_le_bytes([bytes[18], bytes[19]]));
                let height = u32::from(u16::from_le_bytes([bytes[20], bytes[21]]));
                let bpp = u16::from_le_bytes([bytes[24], bytes[25]]);
                let color = match bpp {
                    8 => ColorType::L8,
                    32 => ColorType::Rgba8,
                    _ => ColorType::Rgb8,
                };
                Some((width, height, color))
            } else if dib_size >= 16 && bytes.len() >= 30 {
                let width =
                    i32::from_le_bytes([bytes[18], bytes[19], bytes[20], bytes[21]]).unsigned_abs();
                let height =
                    i32::from_le_bytes([bytes[22], bytes[23], bytes[24], bytes[25]]).unsigned_abs();
                let bpp = u16::from_le_bytes([bytes[28], bytes[29]]);
                let color = match bpp {
                    8 => ColorType::L8,
                    32 => ColorType::Rgba8,
                    _ => ColorType::Rgb8,
                };
                Some((width, height, color))
            } else {
                None
            }
        }
        ImageFormat::Tiff => {
            if bytes.len() < 8 {
                return None;
            }
            let le = match &bytes[0..4] {
                b"II\x2a\x00" => true,
                b"MM\x00\x2a" => false,
                _ => return None,
            };
            let read_u16 = |at: usize| -> Option<u16> {
                let b: [u8; 2] = bytes.get(at..at + 2)?.try_into().ok()?;
                Some(if le {
                    u16::from_le_bytes(b)
                } else {
                    u16::from_be_bytes(b)
                })
            };
            let read_u32 = |at: usize| -> Option<u32> {
                let b: [u8; 4] = bytes.get(at..at + 4)?.try_into().ok()?;
                Some(if le {
                    u32::from_le_bytes(b)
                } else {
                    u32::from_be_bytes(b)
                })
            };
            let ifd_offset = read_u32(4)? as usize;
            let count = read_u16(ifd_offset)? as usize;
            if count > 4096 {
                return None;
            }
            let mut width = None;
            let mut height = None;
            for idx in 0..count {
                let entry = ifd_offset.checked_add(2 + idx * 12)?;
                let tag = read_u16(entry)?;
                let field_type = read_u16(entry + 2)?;
                let value = match field_type {
                    3 => u32::from(read_u16(entry + 8)?),
                    4 => read_u32(entry + 8)?,
                    _ => continue,
                };
                match tag {
                    256 => width = Some(value),
                    257 => height = Some(value),
                    _ => {}
                }
            }
            Some((width?, height?, ColorType::Rgb8))
        }
        _ => None,
    }
}

fn target_color_type(format: ImageFormat, color: ColorType) -> ColorType {
    match format {
        ImageFormat::Png => match color {
            ColorType::L8
            | ColorType::La8
            | ColorType::Rgb8
            | ColorType::Rgba8
            | ColorType::L16
            | ColorType::La16
            | ColorType::Rgb16
            | ColorType::Rgba16 => color,
            ColorType::Rgb32F => ColorType::Rgb16,
            ColorType::Rgba32F => ColorType::Rgba16,
            _ => ColorType::Rgb8,
        },
        ImageFormat::Jpeg => match color {
            ColorType::L8 | ColorType::L16 => ColorType::L8,
            _ => ColorType::Rgb8,
        },
        ImageFormat::Gif => ColorType::Rgba8,
        ImageFormat::Bmp => match color {
            ColorType::L8 | ColorType::L16 => ColorType::L8,
            ColorType::La8
            | ColorType::La16
            | ColorType::Rgba8
            | ColorType::Rgba16
            | ColorType::Rgba32F => ColorType::Rgba8,
            _ => ColorType::Rgb8,
        },
        ImageFormat::Tiff => match color {
            ColorType::L8
            | ColorType::Rgb8
            | ColorType::Rgba8
            | ColorType::L16
            | ColorType::Rgb16
            | ColorType::Rgba16 => color,
            ColorType::La8 | ColorType::La16 | ColorType::Rgba32F => ColorType::Rgba8,
            _ => ColorType::Rgb8,
        },
        _ => ColorType::Rgb8,
    }
}

fn bytes_per_pixel(color: ColorType) -> usize {
    match color {
        ColorType::L8 => 1,
        ColorType::La8 | ColorType::L16 => 2,
        ColorType::Rgb8 => 3,
        ColorType::Rgba8 | ColorType::La16 => 4,
        ColorType::Rgb16 => 6,
        ColorType::Rgba16 => 8,
        ColorType::Rgb32F => 12,
        ColorType::Rgba32F => 16,
        _ => 3,
    }
}

fn within_tolerance(actual: usize, target: usize, fraction: f64) -> bool {
    let diff = actual.abs_diff(target) as f64;
    diff <= (target as f64) * fraction
}

#[allow(clippy::too_many_arguments)]
fn fill_chunked_scanlines(
    raw: &mut [u8],
    width: u32,
    height: u32,
    bpp: usize,
    noise_fraction: f64,
    opaque_fourth_channel: bool,
    base_seed: u64,
    rng: &mut SynthRng,
) {
    let s0 = (base_seed & 0xFF) as u8;
    let s1 = ((base_seed >> 8) & 0xFF) as u8;
    let s2 = ((base_seed >> 16) & 0xFF) as u8;
    let stride = (width as usize) * bpp;
    let chunk_pixels = 16usize;
    let tiles_x = (width as usize).div_ceil(chunk_pixels).max(1);
    let threshold = (noise_fraction.clamp(0.0, 1.0) * 65536.0).round() as u64;

    for y in 0..height as usize {
        let by = y >> 3;
        let row = &mut raw[y * stride..(y + 1) * stride];
        let mut x = 0usize;
        while x < width as usize {
            let run = (width as usize - x).min(chunk_pixels);
            let bx = x / chunk_pixels;
            let tile_idx = (by * tiles_x + bx) as u64;
            let noisy = (((tile_idx + 1) * threshold) >> 16) > ((tile_idx * threshold) >> 16);
            let g0 = ((bx * 5 + by * 3) as u8).wrapping_add(s0);
            let g1 = ((bx * 3 + by * 7 + 85) as u8).wrapping_add(s1);
            let g2 = ((bx * 7 + by * 5 + 170) as u8).wrapping_add(s2);
            let base = [g0, g1, g2, 255, g0, g1, g2, 255];

            let slice = &mut row[x * bpp..(x + run) * bpp];
            if noisy {
                rng.fill_bytes(slice);
                for (idx, byte) in slice.iter_mut().enumerate() {
                    let ch = idx % bpp;
                    if opaque_fourth_channel && ch == 3 {
                        *byte = 255;
                    } else {
                        *byte = byte.wrapping_add(base[ch & 7]);
                    }
                }
            } else {
                for pixel in slice.chunks_exact_mut(bpp) {
                    pixel.copy_from_slice(&base[..bpp]);
                }
            }
            x += run;
        }
    }
}

fn encode_png_sized(
    path: &str,
    width: u32,
    height: u32,
    color_type: ColorType,
    target_len: usize,
    rng: &mut SynthRng,
) -> Result<Vec<u8>, RedactError> {
    encode_noise_sized(width, height, color_type, target_len, rng, |raw| {
        let mut out = Vec::with_capacity(target_len.saturating_add(4096));
        PngEncoder::new_with_quality(&mut out, CompressionType::Fast, FilterType::Sub)
            .write_image(raw, width, height, color_type.into())
            .map_err(|error| RedactError::Image {
                part: path.to_owned(),
                message: error.to_string(),
            })?;
        Ok(out)
    })
}

/// Searches the share of noisy scanline chunks so a lossless encoding lands near `target_len`.
fn encode_noise_sized<F>(
    width: u32,
    height: u32,
    color_type: ColorType,
    target_len: usize,
    rng: &mut SynthRng,
    mut encode: F,
) -> Result<Vec<u8>, RedactError>
where
    F: FnMut(&[u8]) -> Result<Vec<u8>, RedactError>,
{
    let bpp = bytes_per_pixel(color_type);
    let raw_len = (width as usize) * (height as usize) * bpp;
    let mut raw = vec![0u8; raw_len];
    let base_seed = rng.next_u64();
    let mut lo = 0.0f64;
    let mut hi = 1.0f64;
    let mut p = ((target_len as f64 - 128.0) / (raw_len.max(1) as f64 * 1.02)).clamp(0.0, 1.0);
    let mut best: Option<(usize, Vec<u8>)> = None;

    for pass in 0..5 {
        let mut pass_rng = rng.fork(0);
        fill_chunked_scanlines(
            &mut raw,
            width,
            height,
            bpp,
            p,
            false,
            base_seed,
            &mut pass_rng,
        );
        let out = encode(&raw)?;
        let actual = out.len();
        let diff = actual.abs_diff(target_len);
        if within_tolerance(actual, target_len, 0.15) {
            return Ok(out);
        }
        if best.as_ref().is_none_or(|(best_diff, _)| diff < *best_diff) {
            best = Some((diff, out));
        }
        if actual < target_len {
            lo = p;
        } else {
            hi = p;
        }
        if pass == 0 && actual > 0 {
            let scaled = p * (target_len as f64 / actual as f64);
            p = scaled.clamp(lo + (hi - lo) * 0.1, lo + (hi - lo) * 0.9);
        } else {
            p = 0.5 * (lo + hi);
        }
    }
    // Below the mosaic's floor a uniform image is the smallest synthetic content.
    if best.as_ref().is_some_and(|(_, out)| out.len() > target_len) {
        raw.fill((base_seed >> 24) as u8);
        let out = encode(&raw)?;
        if best
            .as_ref()
            .is_some_and(|(diff, _)| out.len().abs_diff(target_len) < *diff)
        {
            return Ok(out);
        }
    }
    Ok(best.map(|(_, out)| out).unwrap_or_default())
}

fn encode_jpeg_once(
    path: &str,
    raw: &[u8],
    width: u32,
    height: u32,
    color_type: ColorType,
    quality: u8,
    capacity: usize,
) -> Result<Vec<u8>, RedactError> {
    let mut out = Vec::with_capacity(capacity);
    JpegEncoder::new_with_quality(&mut out, quality)
        .write_image(raw, width, height, color_type.into())
        .map_err(|error| RedactError::Image {
            part: path.to_owned(),
            message: error.to_string(),
        })?;
    Ok(out)
}

fn encode_jpeg_sized(
    path: &str,
    width: u32,
    height: u32,
    color_type: ColorType,
    target_len: usize,
    rng: &mut SynthRng,
) -> Result<Vec<u8>, RedactError> {
    let bpp = bytes_per_pixel(color_type);
    let raw_len = (width as usize) * (height as usize) * bpp;
    let mut raw = vec![0u8; raw_len];
    let base_seed = rng.next_u64();
    let max_q80_est = (raw_len as f64) * 0.24;
    let mut lo = 0.0f64;
    let mut hi = 1.0f64;
    let mut p = ((target_len as f64) / max_q80_est.max(1.0)).clamp(0.02, 0.98);
    let mut best: Option<(usize, Vec<u8>)> = None;
    let cap = target_len.saturating_add(4096);

    for pass in 0..5 {
        let mut pass_rng = rng.fork(0);
        fill_chunked_scanlines(
            &mut raw,
            width,
            height,
            bpp,
            p,
            false,
            base_seed,
            &mut pass_rng,
        );
        let out = encode_jpeg_once(path, &raw, width, height, color_type, 80, cap)?;
        let actual = out.len();
        let diff = actual.abs_diff(target_len);
        if within_tolerance(actual, target_len, 0.15) {
            return Ok(out);
        }
        if best.as_ref().is_none_or(|(best_diff, _)| diff < *best_diff) {
            best = Some((diff, out));
        }
        if actual < target_len && p >= 0.98 {
            let hi_out = encode_jpeg_once(path, &raw, width, height, color_type, 96, cap)?;
            let hi_diff = hi_out.len().abs_diff(target_len);
            if best
                .as_ref()
                .is_none_or(|(best_diff, _)| hi_diff < *best_diff)
            {
                best = Some((hi_diff, hi_out));
            }
            break;
        } else if actual < target_len {
            lo = p;
        } else if p <= 0.02 {
            let lo_out = encode_jpeg_once(path, &raw, width, height, color_type, 10, cap)?;
            let lo_diff = lo_out.len().abs_diff(target_len);
            if best
                .as_ref()
                .is_none_or(|(best_diff, _)| lo_diff < *best_diff)
            {
                best = Some((lo_diff, lo_out));
            }
            break;
        } else {
            hi = p;
        }
        if pass == 0 && actual > 0 {
            let scaled = p * (target_len as f64 / actual as f64);
            p = scaled.clamp(lo + (hi - lo) * 0.05, lo + (hi - lo) * 0.95);
        } else {
            p = 0.5 * (lo + hi);
        }
    }
    Ok(best.map(|(_, out)| out).unwrap_or_default())
}

fn encode_gif_sized(
    path: &str,
    width: u32,
    height: u32,
    target_len: usize,
    rng: &mut SynthRng,
) -> Result<Vec<u8>, RedactError> {
    let pixels = (width as usize) * (height as usize);
    let mut raw = vec![0u8; pixels * 4];
    let base_seed = rng.next_u64();
    let mut lo = 0.0f64;
    let mut hi = 1.0f64;
    let mut p = ((target_len as f64) / (pixels.max(1) as f64)).clamp(0.0, 1.0);
    let mut best: Option<(usize, Vec<u8>)> = None;

    for pass in 0..5 {
        let mut pass_rng = rng.fork(0);
        fill_chunked_scanlines(
            &mut raw,
            width,
            height,
            4,
            p,
            true,
            base_seed,
            &mut pass_rng,
        );
        let mut out = Vec::with_capacity(target_len.saturating_add(4096));
        GifEncoder::new_with_speed(&mut out, 30)
            .encode(&raw, width, height, ColorType::Rgba8.into())
            .map_err(|error| RedactError::Image {
                part: path.to_owned(),
                message: error.to_string(),
            })?;
        let actual = out.len();
        let diff = actual.abs_diff(target_len);
        if within_tolerance(actual, target_len, 0.15) {
            return Ok(out);
        }
        if best.as_ref().is_none_or(|(best_diff, _)| diff < *best_diff) {
            best = Some((diff, out));
        }
        if actual < target_len {
            lo = p;
        } else {
            hi = p;
        }
        if pass == 0 && actual > 0 {
            let scaled = p * (target_len as f64 / actual as f64);
            p = scaled.clamp(lo + (hi - lo) * 0.1, lo + (hi - lo) * 0.9);
        } else {
            p = 0.5 * (lo + hi);
        }
    }
    Ok(best.map(|(_, out)| out).unwrap_or_default())
}

fn encode_bmp_sized(
    path: &str,
    width: u32,
    height: u32,
    color_type: ColorType,
    rng: &mut SynthRng,
) -> Result<Vec<u8>, RedactError> {
    let bpp = bytes_per_pixel(color_type);
    let mut raw = vec![0u8; (width as usize) * (height as usize) * bpp];
    let base_seed = rng.next_u64();
    fill_chunked_scanlines(&mut raw, width, height, bpp, 0.5, false, base_seed, rng);
    let mut out = Vec::with_capacity(raw.len() + 1088);
    BmpEncoder::new(&mut out)
        .write_image(&raw, width, height, color_type.into())
        .map_err(|error| RedactError::Image {
            part: path.to_owned(),
            message: error.to_string(),
        })?;
    Ok(out)
}

fn encode_tiff_sized(
    path: &str,
    width: u32,
    height: u32,
    color_type: ColorType,
    target_len: usize,
    rng: &mut SynthRng,
) -> Result<Vec<u8>, RedactError> {
    let bpp = bytes_per_pixel(color_type);
    let raw_len = u64::from(width) * u64::from(height) * bpp as u64;
    if tiff_compressed(target_len, raw_len) {
        return encode_noise_sized(width, height, color_type, target_len, rng, |raw| {
            encode_tiff_deflate(raw, width, height, color_type).map_err(|error| {
                RedactError::Image {
                    part: path.to_owned(),
                    message: error.to_string(),
                }
            })
        });
    }
    let mut raw = vec![0u8; (width as usize) * (height as usize) * bpp];
    let base_seed = rng.next_u64();
    fill_chunked_scanlines(&mut raw, width, height, bpp, 0.5, false, base_seed, rng);
    let mut out = Vec::with_capacity(raw.len() + 512);
    TiffEncoder::new(Cursor::new(&mut out))
        .write_image(&raw, width, height, color_type.into())
        .map_err(|error| RedactError::Image {
            part: path.to_owned(),
            message: error.to_string(),
        })?;
    Ok(out)
}

fn encode_tiff_deflate(
    raw: &[u8],
    width: u32,
    height: u32,
    color_type: ColorType,
) -> tiff::TiffResult<Vec<u8>> {
    use tiff::encoder::colortype::{Gray8, Gray16, RGB8, RGB16, RGBA8, RGBA16};
    let wide = || -> Vec<u16> {
        raw.as_chunks::<2>()
            .0
            .iter()
            .map(|pair| u16::from_ne_bytes(*pair))
            .collect()
    };
    let mut out = Cursor::new(Vec::new());
    let mut encoder = tiff::encoder::TiffEncoder::new(&mut out)?.with_compression(
        tiff::encoder::Compression::Deflate(tiff::encoder::DeflateLevel::Best),
    );
    match color_type {
        ColorType::L8 => single_strip::<_, Gray8>(&mut encoder, width, height, raw),
        ColorType::Rgba8 => single_strip::<_, RGBA8>(&mut encoder, width, height, raw),
        ColorType::L16 => single_strip::<_, Gray16>(&mut encoder, width, height, &wide()),
        ColorType::Rgb16 => single_strip::<_, RGB16>(&mut encoder, width, height, &wide()),
        ColorType::Rgba16 => single_strip::<_, RGBA16>(&mut encoder, width, height, &wide()),
        _ => single_strip::<_, RGB8>(&mut encoder, width, height, raw),
    }?;
    Ok(out.into_inner())
}

fn single_strip<W, C>(
    encoder: &mut tiff::encoder::TiffEncoder<W>,
    width: u32,
    height: u32,
    data: &[C::Inner],
) -> tiff::TiffResult<()>
where
    W: std::io::Write + std::io::Seek,
    C: tiff::encoder::colortype::ColorType,
    [C::Inner]: tiff::encoder::TiffValue,
{
    let mut image = encoder.new_image::<C>(width, height)?;
    image.rows_per_strip(height)?;
    image.write_data(data)
}

fn preserve_svg(bytes: &[u8]) -> Result<Vec<u8>, RedactError> {
    let mut rng = SynthRng::new()?;
    let (width, height, view_box) = parse_svg_header(bytes);
    let mut prefix = String::from(r#"<svg xmlns="http://www.w3.org/2000/svg""#);
    for (name, value) in [("width", width), ("height", height)] {
        if let Some(value) = value {
            prefix.push_str(&format!(r#" {name}="{value}""#));
        }
    }
    if let Some(vb) = &view_box {
        prefix.push_str(r#" viewBox=""#);
        prefix.push_str(vb);
        prefix.push('"');
    }
    prefix.push_str(r##"><rect width="100%" height="100%" fill="#aaa"/>"##);
    let suffix = "</svg>";
    let target_len = bytes.len().min(MAX_SVG_PRESERVED_BYTES);
    let base_len = prefix.len() + suffix.len();
    if target_len > base_len + 7 {
        let pad_len = target_len - base_len - 7;
        prefix.push_str("<!--");
        const HEX: &[u8; 16] = b"0123456789abcdef";
        let mut remaining = pad_len;
        while remaining > 0 {
            let word = rng.next_u64().to_le_bytes();
            for byte in word {
                if remaining == 0 {
                    break;
                }
                prefix.push(HEX[(byte & 0x0F) as usize] as char);
                remaining -= 1;
            }
        }
        prefix.push_str("-->");
    }
    prefix.push_str(suffix);
    Ok(prefix.into_bytes())
}

fn parse_svg_header(bytes: &[u8]) -> (Option<String>, Option<String>, Option<String>) {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        let event = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(tag) | Event::Empty(tag)) => tag.into_owned(),
            Ok(Event::Eof) | Err(_) => return (None, None, None),
            _ => {
                buf.clear();
                continue;
            }
        };
        let raw_name = event.name();
        let local = raw_name
            .as_ref()
            .rsplit(|byte| *byte == b':')
            .next()
            .unwrap_or(raw_name.as_ref());
        if !local.eq_ignore_ascii_case(b"svg") {
            return (None, None, None);
        }
        let mut width = None;
        let mut height = None;
        let mut view_box = None;
        for attr in event.attributes().flatten() {
            let key = attr.key.as_ref();
            let Ok(val) =
                attr.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            else {
                continue;
            };
            let val = val.as_ref();
            match key {
                b"width" => width = sanitize_svg_length(val),
                b"height" => height = sanitize_svg_length(val),
                b"viewBox" => view_box = sanitize_svg_view_box(val),
                _ => {}
            }
        }
        return (width, height, view_box);
    }
}

fn sanitize_svg_length(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() || trimmed.len() > 64 {
        return None;
    }
    let (num_part, unit_part) = ["px", "pt", "pc", "mm", "cm", "in", "em", "ex", "%"]
        .into_iter()
        .find_map(|unit| trimmed.strip_suffix(unit).map(|number| (number, unit)))
        .unwrap_or((trimmed, ""));
    let number: f64 = num_part.parse().ok()?;
    if !number.is_finite() || number <= 0.0 || number > 1_000_000.0 {
        return None;
    }
    Some(format!("{}{unit_part}", clean_svg_num(number)))
}

fn sanitize_svg_view_box(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() || trimmed.len() > 128 {
        return None;
    }
    let parts: Vec<f64> = trimmed
        .split(|c: char| c.is_ascii_whitespace() || c == ',')
        .filter(|part| !part.is_empty())
        .map(|part| part.parse::<f64>().ok())
        .collect::<Option<Vec<f64>>>()?;
    if parts.len() != 4
        || !parts
            .iter()
            .all(|v| v.is_finite() && v.abs() <= 1_000_000.0)
    {
        return None;
    }
    if parts[2] <= 0.0 || parts[3] <= 0.0 {
        return None;
    }
    Some(format!(
        "{} {} {} {}",
        clean_svg_num(parts[0]),
        clean_svg_num(parts[1]),
        clean_svg_num(parts[2]),
        clean_svg_num(parts[3])
    ))
}

fn clean_svg_num(value: f64) -> String {
    let rounded = (value * 1000.0).round() / 1000.0;
    if rounded == 0.0 {
        return "0".to_owned();
    }
    let mut text = format!("{rounded:.3}");
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text
}

fn preserve_emf(bytes: &[u8]) -> Result<Vec<u8>, RedactError> {
    let Some(header) = parse_emf_header(bytes) else {
        return Ok(emf_stub());
    };
    let mut rng = SynthRng::new()?;
    let target_len = bytes.len().min(MAX_PRESERVED_BYTES) & !3;
    let mut out = Vec::with_capacity(target_len.max(108));
    out.extend_from_slice(&1u32.to_le_bytes()); // iType = EMR_HEADER
    out.extend_from_slice(&88u32.to_le_bytes()); // nSize
    out.extend_from_slice(&header.bounds); // rclBounds (16 bytes)
    out.extend_from_slice(&header.frame); // rclFrame (16 bytes)
    out.extend_from_slice(&0x464D_4520u32.to_le_bytes()); // " EMF"
    out.extend_from_slice(&0x0001_0000u32.to_le_bytes()); // version 1.0
    out.extend_from_slice(&108u32.to_le_bytes()); // nBytes placeholder
    out.extend_from_slice(&2u32.to_le_bytes()); // nRecords placeholder
    out.extend_from_slice(&1u16.to_le_bytes()); // nHandles
    out.extend_from_slice(&0u16.to_le_bytes()); // sReserved
    out.extend_from_slice(&0u32.to_le_bytes()); // nDescription
    out.extend_from_slice(&0u32.to_le_bytes()); // offDescription
    out.extend_from_slice(&0u32.to_le_bytes()); // nPalEntries
    out.extend_from_slice(&header.device); // szlDevice (8 bytes)
    out.extend_from_slice(&header.millimeters); // szlMillimeters (8 bytes)
    debug_assert_eq!(out.len(), 88);

    let mut n_records = 2u32;
    let mut remaining = target_len.saturating_sub(108);
    while remaining >= 16 {
        let rec_size = if remaining >= 32_768 && remaining - 32_768 >= 16 {
            32_768
        } else if remaining <= 32_768 {
            remaining
        } else {
            remaining - 16
        };
        let cb_data = rec_size - 12;
        out.extend_from_slice(&70u32.to_le_bytes()); // iType = EMR_COMMENT
        out.extend_from_slice(&(rec_size as u32).to_le_bytes());
        out.extend_from_slice(&(cb_data as u32).to_le_bytes());
        out.extend_from_slice(b"RDCT");
        let payload_start = out.len();
        out.resize(payload_start + cb_data - 4, 0);
        rng.fill_bytes(&mut out[payload_start..]);
        remaining -= rec_size;
        n_records += 1;
    }

    out.extend_from_slice(&14u32.to_le_bytes()); // iType = EMR_EOF
    out.extend_from_slice(&20u32.to_le_bytes()); // nSize
    out.extend_from_slice(&0u32.to_le_bytes()); // nPalEntries
    out.extend_from_slice(&16u32.to_le_bytes()); // offPalEntries
    out.extend_from_slice(&20u32.to_le_bytes()); // nSizeLast
    let total_bytes = out.len() as u32;
    out[48..52].copy_from_slice(&total_bytes.to_le_bytes());
    out[52..56].copy_from_slice(&n_records.to_le_bytes());
    Ok(out)
}

struct EmfHeaderFields {
    bounds: [u8; 16],
    frame: [u8; 16],
    device: [u8; 8],
    millimeters: [u8; 8],
}

fn parse_emf_header(bytes: &[u8]) -> Option<EmfHeaderFields> {
    if bytes.len() < 88 {
        return None;
    }
    let i_type = u32::from_le_bytes(bytes[0..4].try_into().ok()?);
    let n_size = u32::from_le_bytes(bytes[4..8].try_into().ok()?);
    let signature = u32::from_le_bytes(bytes[40..44].try_into().ok()?);
    if i_type != 1 || n_size < 88 || signature != 0x464D_4520 {
        return None;
    }
    let i32_at = |offset: usize| -> i32 {
        i32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap_or([0; 4]))
    };
    let bounds_w = i64::from(i32_at(16)) - i64::from(i32_at(8));
    let bounds_h = i64::from(i32_at(20)) - i64::from(i32_at(12));
    let frame_w = i64::from(i32_at(32)) - i64::from(i32_at(24));
    let frame_h = i64::from(i32_at(36)) - i64::from(i32_at(28));
    let dev_cx = i32_at(72);
    let dev_cy = i32_at(76);
    let mm_cx = i32_at(80);
    let mm_cy = i32_at(84);
    let has_frame = mm_cx > 0 && mm_cy > 0 && frame_w > 0 && frame_h > 0;
    let has_bounds = bounds_w > 0 && bounds_h > 0;
    if !has_frame && !has_bounds {
        return None;
    }
    let mut device = [0u8; 8];
    if dev_cx > 0 && dev_cy > 0 {
        device.copy_from_slice(&bytes[72..80]);
    } else {
        device[0..4].copy_from_slice(&1920i32.to_le_bytes());
        device[4..8].copy_from_slice(&1080i32.to_le_bytes());
    }
    let mut millimeters = [0u8; 8];
    if mm_cx > 0 && mm_cy > 0 {
        millimeters.copy_from_slice(&bytes[80..88]);
    } else {
        millimeters[0..4].copy_from_slice(&508i32.to_le_bytes());
        millimeters[4..8].copy_from_slice(&285i32.to_le_bytes());
    }
    Some(EmfHeaderFields {
        bounds: bytes[8..24].try_into().ok()?,
        frame: bytes[24..40].try_into().ok()?,
        device,
        millimeters,
    })
}

fn preserve_wmf(bytes: &[u8]) -> Result<Vec<u8>, RedactError> {
    let placeable = parse_wmf_placeable(bytes);
    let window = if placeable.is_none() {
        parse_wmf_window(bytes)
    } else {
        None
    };
    if placeable.is_none() && window.is_none() {
        return Ok(wmf_stub());
    }
    let mut rng = SynthRng::new()?;
    let target_len = bytes.len().min(MAX_PRESERVED_BYTES) & !1;
    let mut out = Vec::with_capacity(target_len.max(46));
    if let Some((bbox, inch)) = placeable {
        out.extend_from_slice(&0x9AC6_CDD7u32.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&bbox);
        out.extend_from_slice(&inch.to_le_bytes());
        out.extend_from_slice(&[0; 4]);
        let checksum = out
            .as_chunks::<2>()
            .0
            .iter()
            .fold(0u16, |acc, chunk| acc ^ u16::from_le_bytes(*chunk));
        out.extend_from_slice(&checksum.to_le_bytes());
    }
    let header_offset = out.len();

    out.extend_from_slice(&1u16.to_le_bytes()); // mtType = memory
    out.extend_from_slice(&9u16.to_le_bytes()); // mtHeaderSize words
    out.extend_from_slice(&0x0300u16.to_le_bytes()); // version 3.0
    out.extend_from_slice(&12u32.to_le_bytes()); // mtSize words placeholder
    out.extend_from_slice(&0u16.to_le_bytes()); // mtNoObjects
    out.extend_from_slice(&3u32.to_le_bytes()); // mtMaxRecord words placeholder
    out.extend_from_slice(&0u16.to_le_bytes()); // mtNoParameters
    let mut max_record_words = 3u32;
    if let Some((origin, extent)) = window {
        for (function, values) in [(0x020Bu16, origin), (0x020C, extent)] {
            out.extend_from_slice(&5u32.to_le_bytes());
            out.extend_from_slice(&function.to_le_bytes());
            out.extend_from_slice(&values);
        }
        max_record_words = 5;
    }
    let mut remaining = target_len.saturating_sub(out.len() + 6);
    while remaining >= 12 {
        let rec_bytes = if remaining >= 32_768 && remaining - 32_768 >= 12 {
            32_768
        } else if remaining <= 32_768 {
            remaining
        } else {
            remaining - 12
        };
        let rec_words = (rec_bytes / 2) as u32;
        max_record_words = max_record_words.max(rec_words);
        let data_bytes = rec_bytes - 10;
        out.extend_from_slice(&rec_words.to_le_bytes()); // RecordSize words
        out.extend_from_slice(&0x0626u16.to_le_bytes()); // META_ESCAPE
        out.extend_from_slice(&0u16.to_le_bytes()); // EscapeFunction = 0
        out.extend_from_slice(&(data_bytes as u16).to_le_bytes()); // ByteCount
        let payload_start = out.len();
        out.resize(payload_start + data_bytes, 0);
        rng.fill_bytes(&mut out[payload_start..]);
        remaining -= rec_bytes;
    }

    out.extend_from_slice(&3u32.to_le_bytes()); // META_EOF RecordSize words
    out.extend_from_slice(&0u16.to_le_bytes()); // META_EOF RecordFunction
    let mt_size_words = ((out.len() - header_offset) / 2) as u32;
    out[header_offset + 6..header_offset + 10].copy_from_slice(&mt_size_words.to_le_bytes());
    out[header_offset + 12..header_offset + 16].copy_from_slice(&max_record_words.to_le_bytes());
    Ok(out)
}

fn parse_wmf_window(bytes: &[u8]) -> Option<([u8; 4], [u8; 4])> {
    let header = bytes.get(..18)?;
    if !matches!(u16::from_le_bytes(header[..2].try_into().ok()?), 1 | 2)
        || u16::from_le_bytes(header[2..4].try_into().ok()?) != 9
    {
        return None;
    }
    let mut offset = 18usize;
    let mut origin = None;
    let mut extent = None;
    for _ in 0..100_000 {
        let record = bytes.get(offset..offset.checked_add(6)?)?;
        let size = usize::try_from(u32::from_le_bytes(record[..4].try_into().ok()?))
            .ok()?
            .checked_mul(2)?;
        if size < 6 {
            return None;
        }
        let end = offset.checked_add(size)?;
        let record = bytes.get(offset..end)?;
        match u16::from_le_bytes(record[4..6].try_into().ok()?) {
            0 => {
                let extent: [u8; 4] = extent?;
                if i16::from_le_bytes(extent[..2].try_into().ok()?) == 0
                    || i16::from_le_bytes(extent[2..].try_into().ok()?) == 0
                {
                    return None;
                }
                return Some((origin.unwrap_or([0; 4]), extent));
            }
            0x020B => origin = origin.or(Some(record.get(6..10)?.try_into().ok()?)),
            0x020C => extent = extent.or(Some(record.get(6..10)?.try_into().ok()?)),
            _ => {}
        }
        offset = end;
    }
    None
}

fn parse_wmf_placeable(bytes: &[u8]) -> Option<([u8; 8], u16)> {
    if bytes.len() < 22 {
        return None;
    }
    let key = u32::from_le_bytes(bytes[0..4].try_into().ok()?);
    if key != 0x9AC6_CDD7 {
        return None;
    }
    let left = i16::from_le_bytes(bytes[6..8].try_into().ok()?);
    let top = i16::from_le_bytes(bytes[8..10].try_into().ok()?);
    let right = i16::from_le_bytes(bytes[10..12].try_into().ok()?);
    let bottom = i16::from_le_bytes(bytes[12..14].try_into().ok()?);
    let inch = u16::from_le_bytes(bytes[14..16].try_into().ok()?);
    if left == right || top == bottom || inch == 0 {
        return None;
    }
    Some((bytes[6..14].try_into().ok()?, inch))
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
