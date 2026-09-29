//! TIFF to PNG decoder for the web renderer.

extern crate tiff as tiff_crate;

use std::io::{Cursor, Read};

pub const MAX_IMAGE_PIXELS: u64 = 33_554_432;
pub const MAX_IMAGE_BYTES: u64 = 268_435_456;
pub const MAX_TIFF_BYTES: u64 = 32 * 1024 * 1024;

pub fn decode_tiff_png(data: &[u8]) -> Result<Vec<u8>, String> {
    use tiff_crate::decoder::{Decoder, DecodingResult};
    use tiff_crate::tags::Tag;

    let source_bytes = data.len() as u64;
    if !image_fits_budget(0, 0, source_bytes) {
        return Err("TIFF image exceeds the browser decode budget".to_owned());
    }
    let mut limits = tiff_crate::decoder::Limits::default();
    limits.decoding_buffer_size = MAX_IMAGE_BYTES as usize;
    let mut decoder = Decoder::new(Cursor::new(data))
        .map_err(|error| error.to_string())?
        .with_limits(limits);
    let (width, height) = decoder.dimensions().map_err(|error| error.to_string())?;
    if width == 0 || height == 0 {
        return Err("Malformed TIFF image data".to_owned());
    }
    let pixels = u64::from(width) * u64::from(height);
    let orientation = tag_u16(&mut decoder, Tag::Orientation)?.unwrap_or(1);
    let photometric = tag_u16(&mut decoder, Tag::PhotometricInterpretation)?.unwrap_or(1);
    let color = match decoder.colortype() {
        Ok(color) => Some(color),
        Err(_) if photometric == 3 => None,
        Err(error) => return Err(error.to_string()),
    };
    let decoded_bytes = match color {
        Some(color) => {
            let sample_bytes = u64::from(color.bit_depth().div_ceil(8)).max(1);
            pixels
                .saturating_mul(u64::from(color.num_samples()))
                .saturating_mul(sample_bytes)
        }
        None => pixels.saturating_mul(2),
    };
    if !image_fits_budget(
        pixels,
        decoded_bytes.saturating_add(pixels.saturating_mul(4)),
        source_bytes,
    ) {
        return Err("TIFF image exceeds the browser decode budget".to_owned());
    }
    let pixels = pixels as usize;
    let rgba = match color {
        Some(color) => {
            let mut buffer = DecodingResult::U8(Vec::new());
            let layout = decoder
                .read_image_to_buffer(&mut buffer)
                .map_err(|error| error.to_string())?;
            if buffer.as_buffer(0).as_bytes().len() < layout.complete_len {
                return Err("TIFF image exceeds the browser decode budget".to_owned());
            }
            let extra_samples = tag_u16_vec(&mut decoder, Tag::ExtraSamples)?;
            samples_to_rgba8(
                color,
                buffer,
                &layout,
                matches!(extra_samples.first(), Some(1 | 2)),
                width as usize,
                pixels,
            )?
        }
        None => {
            let raw = *decoder.inner().get_ref();
            palette_to_rgba8(&mut decoder, raw, width as usize, height as usize, pixels)?
        }
    };
    let (rgba, width, height) = apply_orientation(rgba, width, height, orientation);
    crate::png_encode::encode_rgba8(&rgba, width, height)
}

/// Single numeric IFD entry.
fn tag_u16<R: Read + std::io::Seek>(
    decoder: &mut tiff_crate::decoder::Decoder<R>,
    tag: tiff_crate::tags::Tag,
) -> Result<Option<u16>, String> {
    decoder
        .find_tag(tag)
        .map_err(|error| error.to_string())?
        .map(|value| value.into_u16().map_err(|error| error.to_string()))
        .transpose()
}

/// Numeric IFD entry list.
fn tag_u16_vec<R: Read + std::io::Seek>(
    decoder: &mut tiff_crate::decoder::Decoder<R>,
    tag: tiff_crate::tags::Tag,
) -> Result<Vec<u16>, String> {
    Ok(decoder
        .find_tag(tag)
        .map_err(|error| error.to_string())?
        .map(|value| value.into_u16_vec().map_err(|error| error.to_string()))
        .transpose()?
        .unwrap_or_default())
}

/// Numeric IFD entry widened to u64.
fn tag_u64<R: Read + std::io::Seek>(
    decoder: &mut tiff_crate::decoder::Decoder<R>,
    tag: tiff_crate::tags::Tag,
) -> Result<Option<u64>, String> {
    decoder
        .find_tag_unsigned(tag)
        .map_err(|error| error.to_string())
}

/// Numeric IFD entry list widened to u64.
fn tag_u64_vec<R: Read + std::io::Seek>(
    decoder: &mut tiff_crate::decoder::Decoder<R>,
    tag: tiff_crate::tags::Tag,
) -> Result<Vec<u64>, String> {
    Ok(decoder
        .find_tag_unsigned_vec(tag)
        .map_err(|error| error.to_string())?
        .unwrap_or_default())
}

/// Chunky native-endian sample raster to RGBA8, deplaning first.
fn samples_to_rgba8(
    color: tiff_crate::ColorType,
    mut buffer: tiff_crate::decoder::DecodingResult,
    layout: &tiff_crate::decoder::BufferLayoutPreference,
    extra_alpha: bool,
    width: usize,
    pixels: usize,
) -> Result<Vec<u8>, String> {
    let sample_bytes = match &buffer {
        tiff_crate::decoder::DecodingResult::U8(_) => 1,
        tiff_crate::decoder::DecodingResult::U16(_) => 2,
        _ => return Err("Unsupported TIFF sample format".to_owned()),
    };
    let chunky = deplane(buffer.as_buffer(0).as_bytes(), layout, sample_bytes, pixels)?;
    if let Some(rgba) =
        gray_samples_to_rgba8(color, &chunky, width, pixels, sample_bytes, extra_alpha)?
    {
        return Ok(rgba);
    }
    if let Some(rgba) = rgb_samples_to_rgba8(color, &chunky, pixels, sample_bytes)? {
        return Ok(rgba);
    }
    if let Some(rgba) = cmyk_samples_to_rgba8(color, &chunky, pixels, sample_bytes)? {
        return Ok(rgba);
    }
    Err("Unsupported TIFF color type".to_owned())
}

/// Gray and gray-alpha rasters to RGBA8.
fn gray_samples_to_rgba8(
    color: tiff_crate::ColorType,
    chunky: &[u8],
    width: usize,
    pixels: usize,
    sample_bytes: usize,
    extra_alpha: bool,
) -> Result<Option<Vec<u8>>, String> {
    let rgba = match color {
        tiff_crate::ColorType::Gray(8) if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 1, 1)?;
            gray_opaque(chunky)
        }
        tiff_crate::ColorType::Gray(bits) if sample_bytes == 1 && matches!(bits, 1 | 2 | 4) => {
            return unpack_gray(chunky, width, pixels, bits).map(Some);
        }
        tiff_crate::ColorType::Gray(16) if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 1, 2)?;
            gray16_opaque(chunky)
        }
        tiff_crate::ColorType::GrayA(8) if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 2, 1)?;
            gray_alpha(chunky)
        }
        tiff_crate::ColorType::GrayA(16) if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 2, 2)?;
            gray_alpha16(chunky)
        }
        tiff_crate::ColorType::Multiband {
            bit_depth: 8,
            num_samples: 2,
        } if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 2, 1)?;
            if extra_alpha {
                gray_alpha(chunky)
            } else {
                gray_dropped(chunky)
            }
        }
        tiff_crate::ColorType::Multiband {
            bit_depth: 16,
            num_samples: 2,
        } if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 2, 2)?;
            if extra_alpha {
                gray_alpha16(chunky)
            } else {
                gray_dropped16(chunky)
            }
        }
        _ => return Ok(None),
    };
    Ok(Some(rgba))
}

/// RGB and RGB-alpha rasters to RGBA8.
fn rgb_samples_to_rgba8(
    color: tiff_crate::ColorType,
    chunky: &[u8],
    pixels: usize,
    sample_bytes: usize,
) -> Result<Option<Vec<u8>>, String> {
    let rgba = match color {
        tiff_crate::ColorType::RGB(8) if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 3, 1)?;
            rgb_opaque(chunky)
        }
        tiff_crate::ColorType::RGB(16) if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 3, 2)?;
            rgb16_opaque(chunky)
        }
        tiff_crate::ColorType::RGBA(8) if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 4, 1)?;
            chunky.to_vec()
        }
        tiff_crate::ColorType::RGBA(16) if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 4, 2)?;
            rgba16_to_rgba8(chunky)
        }
        _ => return Ok(None),
    };
    Ok(Some(rgba))
}

/// CMYK and CMYK-alpha rasters to RGBA8.
fn cmyk_samples_to_rgba8(
    color: tiff_crate::ColorType,
    chunky: &[u8],
    pixels: usize,
    sample_bytes: usize,
) -> Result<Option<Vec<u8>>, String> {
    let rgba = match color {
        tiff_crate::ColorType::CMYK(8) if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 4, 1)?;
            cmyk_opaque(chunky)
        }
        tiff_crate::ColorType::CMYK(16) if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 4, 2)?;
            cmyk16_opaque(chunky)
        }
        tiff_crate::ColorType::CMYKA(8) if sample_bytes == 1 => {
            expect_samples(chunky, pixels, 5, 1)?;
            cmyka_opaque(chunky)
        }
        tiff_crate::ColorType::CMYKA(16) if sample_bytes == 2 => {
            expect_samples(chunky, pixels, 5, 2)?;
            cmyka16_opaque(chunky)
        }
        _ => return Ok(None),
    };
    Ok(Some(rgba))
}

/// Reject a raster whose sample count is not exactly as expected.
fn expect_samples(
    bytes: &[u8],
    pixels: usize,
    samples: usize,
    sample_bytes: usize,
) -> Result<(), String> {
    if bytes.len() == pixels * samples * sample_bytes {
        Ok(())
    } else {
        Err("Malformed TIFF image data".to_owned())
    }
}

/// Concatenated planes to chunky order.
fn deplane(
    bytes: &[u8],
    layout: &tiff_crate::decoder::BufferLayoutPreference,
    sample_bytes: usize,
    pixels: usize,
) -> Result<Vec<u8>, String> {
    if layout.planes < 2 {
        return Ok(bytes.to_vec());
    }
    let total = (pixels as u64)
        .saturating_mul(layout.planes as u64)
        .saturating_mul(sample_bytes as u64);
    if total > MAX_IMAGE_BYTES {
        return Err("TIFF image exceeds the browser decode budget".to_owned());
    }
    let stride = layout.plane_stride.map(|stride| stride.get()).unwrap_or(0);
    if stride != pixels * sample_bytes || bytes.len() < stride * layout.planes {
        return Err("Unsupported TIFF planar layout".to_owned());
    }
    let mut chunky = vec![0; pixels * layout.planes * sample_bytes];
    for (index, pixel) in chunky
        .chunks_exact_mut(layout.planes * sample_bytes)
        .enumerate()
    {
        for plane in 0..layout.planes {
            let start = plane * stride + index * sample_bytes;
            pixel[plane * sample_bytes..][..sample_bytes]
                .copy_from_slice(&bytes[start..start + sample_bytes]);
        }
    }
    Ok(chunky)
}

/// Palette TIFF pixels via direct strip reads; the full decoder rejects palettes.
fn palette_to_rgba8<R: Read + std::io::Seek>(
    decoder: &mut tiff_crate::decoder::Decoder<R>,
    data: &[u8],
    width: usize,
    height: usize,
    pixels: usize,
) -> Result<Vec<u8>, String> {
    use tiff_crate::tags::Tag;

    let bits = tag_u16(decoder, Tag::BitsPerSample)?.unwrap_or(1);
    if !matches!(bits, 1 | 2 | 4 | 8 | 16) {
        return Err("Unsupported TIFF palette depth".to_owned());
    }
    if tag_u16(decoder, Tag::SamplesPerPixel)?.unwrap_or(1) != 1 {
        return Err("Unsupported TIFF palette layout".to_owned());
    }
    if decoder
        .find_tag(Tag::TileOffsets)
        .map_err(|error| error.to_string())?
        .is_some()
    {
        return Err("Unsupported TIFF palette layout".to_owned());
    }
    let compression = tag_u16(decoder, Tag::Compression)?.unwrap_or(1);
    let predictor = tag_u16(decoder, Tag::Predictor)?.unwrap_or(1);
    if predictor != 1 && predictor != 2 {
        return Err("Unsupported TIFF predictor".to_owned());
    }
    let rows_per_strip = tag_u64(decoder, Tag::RowsPerStrip)?.unwrap_or(height as u64);
    let offsets = tag_u64_vec(decoder, Tag::StripOffsets)?;
    let counts = tag_u64_vec(decoder, Tag::StripByteCounts)?;
    if offsets.is_empty() || offsets.len() != counts.len() {
        return Err("Malformed TIFF image data".to_owned());
    }
    let colormap = decoder
        .get_tag_u16_vec(Tag::ColorMap)
        .map_err(|error| error.to_string())?;
    let entries = 1usize << bits;
    if colormap.len() != 3 * entries {
        return Err("Malformed TIFF image data".to_owned());
    }
    let order = decoder.byte_order();
    let row_bytes = (width as u64 * u64::from(bits)).div_ceil(8);
    let mut packed = vec![0; row_bytes as usize * height];
    let mut strips_bytes: u64 = 0;
    for (strip, (&offset, &count)) in offsets.iter().zip(counts.iter()).enumerate() {
        let start_row = strip as u64 * rows_per_strip;
        if start_row >= height as u64 {
            break;
        }
        let rows = rows_per_strip.min(height as u64 - start_row);
        let expected = rows * row_bytes;
        if expected > MAX_IMAGE_BYTES {
            return Err("TIFF image exceeds the browser decode budget".to_owned());
        }
        strips_bytes = strips_bytes.saturating_add(count);
        if strips_bytes > MAX_IMAGE_BYTES {
            return Err("TIFF image exceeds the browser decode budget".to_owned());
        }
        let slice = file_range(data, offset, count)?;
        let mut raw = expand_strip(compression, slice, expected as usize)?;
        if predictor == 2 {
            undo_predictor(&mut raw, row_bytes as usize, bits, order);
        }
        let start = start_row as usize * row_bytes as usize;
        packed[start..start + raw.len()].copy_from_slice(&raw);
    }
    map_palette(&packed, &colormap, entries, bits, order, width, pixels)
}

/// Bounded file slice.
fn file_range(data: &[u8], offset: u64, count: u64) -> Result<&[u8], String> {
    if offset > data.len() as u64 || count > MAX_IMAGE_BYTES {
        return Err("Malformed TIFF image data".to_owned());
    }
    let end = (offset as usize)
        .checked_add(count as usize)
        .ok_or_else(|| "Malformed TIFF image data".to_owned())?;
    data.get(offset as usize..end)
        .ok_or_else(|| "Malformed TIFF image data".to_owned())
}

/// Raw strip bytes to exactly `expected` bytes.
fn expand_strip(compression: u16, strip: &[u8], expected: usize) -> Result<Vec<u8>, String> {
    match compression {
        1 => {
            if strip.len() < expected {
                return Err("Malformed TIFF image data".to_owned());
            }
            Ok(strip[..expected].to_vec())
        }
        5 => lzw_strip(strip, expected),
        8 | 0x80B2 => {
            let mut out = vec![0; expected];
            flate2::read::ZlibDecoder::new(strip)
                .read_exact(&mut out)
                .map_err(|error| error.to_string())?;
            Ok(out)
        }
        32773 => packbits_strip(strip, expected),
        _ => Err(unsupported_compression(compression)),
    }
}

/// Unsupported compression method as a plain error.
fn unsupported_compression(compression: u16) -> String {
    match tiff_crate::tags::CompressionMethod::from_u16(compression) {
        Some(method) => format!("compression method {method:?} is unsupported"),
        None => format!("compression method {compression} is unsupported"),
    }
}

/// LZW strip to exactly `expected` bytes.
fn lzw_strip(strip: &[u8], expected: usize) -> Result<Vec<u8>, String> {
    let mut decoder = weezl::decode::Configuration::with_tiff_size_switch(weezl::BitOrder::Msb, 8)
        .with_yield_on_full_buffer(true)
        .build();
    let mut out = vec![0; expected];
    let mut consumed = 0;
    let mut filled = 0;
    loop {
        let result = decoder.decode_bytes(&strip[consumed..], &mut out[filled..]);
        consumed += result.consumed_in;
        filled += result.consumed_out;
        match result.status {
            Ok(weezl::LzwStatus::Done) => {
                if filled != expected {
                    return Err("Malformed TIFF image data".to_owned());
                }
                return Ok(out);
            }
            Ok(weezl::LzwStatus::Ok) => {
                if filled == expected {
                    return Ok(out);
                }
                if result.consumed_in == 0 && result.consumed_out == 0 {
                    return Err("Malformed TIFF image data".to_owned());
                }
            }
            Ok(weezl::LzwStatus::NoProgress) => {
                return Err("Malformed TIFF image data".to_owned());
            }
            Err(error) => return Err(error.to_string()),
        }
    }
}

/// PackBits strip to exactly `expected` bytes.
fn packbits_strip(strip: &[u8], expected: usize) -> Result<Vec<u8>, String> {
    let mut out = Vec::with_capacity(expected);
    let mut input = strip;
    while out.len() < expected {
        let (header, rest) = input.split_first().ok_or("Malformed TIFF image data")?;
        input = rest;
        let header = *header as i8;
        if header >= 0 {
            let count = header as usize + 1;
            if input.len() < count || out.len() + count > expected {
                return Err("Malformed TIFF image data".to_owned());
            }
            out.extend_from_slice(&input[..count]);
            input = &input[count..];
        } else if header != i8::MIN {
            let count = (1 - i16::from(header)) as usize;
            let (byte, rest) = input.split_first().ok_or("Malformed TIFF image data")?;
            if out.len() + count > expected {
                return Err("Malformed TIFF image data".to_owned());
            }
            out.resize(out.len() + count, *byte);
            input = rest;
        }
    }
    Ok(out)
}

/// Reverse horizontal differencing on one strip raster.
fn undo_predictor(raw: &mut [u8], row_bytes: usize, bits: u16, order: tiff_crate::tags::ByteOrder) {
    if bits == 16 {
        for row in raw.chunks_exact_mut(row_bytes) {
            let mut previous = 0u16;
            for pair in row.as_chunks_mut::<2>().0 {
                let value = u16::from_bytes(pair, order).wrapping_add(previous);
                pair.copy_from_slice(&value.to_bytes(order));
                previous = value;
            }
        }
    } else {
        for row in raw.chunks_exact_mut(row_bytes) {
            for index in 1..row.len() {
                row[index] = row[index].wrapping_add(row[index - 1]);
            }
        }
    }
}

/// Packed or wide palette indices through the color map to RGBA8.
fn map_palette(
    packed: &[u8],
    colormap: &[u16],
    entries: usize,
    bits: u16,
    order: tiff_crate::tags::ByteOrder,
    width: usize,
    pixels: usize,
) -> Result<Vec<u8>, String> {
    let mut rgba = vec![0; pixels * 4];
    if bits == 16 {
        if packed.len() != pixels * 2 {
            return Err("Malformed TIFF image data".to_owned());
        }
        for (pixel, pair) in rgba
            .as_chunks_mut::<4>()
            .0
            .iter_mut()
            .zip(packed.as_chunks::<2>().0)
        {
            let index = usize::from(u16::from_bytes(pair, order));
            if index >= entries {
                return Err("Malformed TIFF image data".to_owned());
            }
            pixel[0] = (colormap[index] >> 8) as u8;
            pixel[1] = (colormap[entries + index] >> 8) as u8;
            pixel[2] = (colormap[2 * entries + index] >> 8) as u8;
            pixel[3] = 255;
        }
        return Ok(rgba);
    }
    let stride = (width * usize::from(bits)).div_ceil(8);
    if packed.len() != stride * (pixels / width) {
        return Err("Malformed TIFF image data".to_owned());
    }
    let max = (1u16 << bits) - 1;
    for (y, row) in packed.chunks_exact(stride).enumerate() {
        for x in 0..width {
            let bit = x * usize::from(bits);
            let index = usize::from((row[bit / 8] >> (8 - bits - (bit % 8) as u16)) & max as u8);
            if index >= entries {
                return Err("Malformed TIFF image data".to_owned());
            }
            let pixel = &mut rgba[(y * width + x) * 4..][..4];
            pixel[0] = (colormap[index] >> 8) as u8;
            pixel[1] = (colormap[entries + index] >> 8) as u8;
            pixel[2] = (colormap[2 * entries + index] >> 8) as u8;
            pixel[3] = 255;
        }
    }
    Ok(rgba)
}

/// EXIF orientation applied to an RGBA8 raster.
fn apply_orientation(
    rgba: Vec<u8>,
    width: u32,
    height: u32,
    orientation: u16,
) -> (Vec<u8>, u32, u32) {
    let (width, height) = (width as usize, height as usize);
    if !matches!(orientation, 2..=8) || rgba.len() != width * height * 4 {
        return (rgba, width as u32, height as u32);
    }
    let swap = matches!(orientation, 5..=8);
    let (output_width, output_height) = if swap {
        (height, width)
    } else {
        (width, height)
    };
    let mut out = vec![0; rgba.len()];
    for y in 0..output_height {
        for x in 0..output_width {
            let (source_x, source_y) = match orientation {
                2 => (width - 1 - x, y),
                3 => (width - 1 - x, height - 1 - y),
                4 => (x, height - 1 - y),
                5 => (y, x),
                6 => (y, height - 1 - x),
                7 => (width - 1 - y, height - 1 - x),
                _ => (width - 1 - y, x),
            };
            out[(y * output_width + x) * 4..][..4]
                .copy_from_slice(&rgba[(source_y * width + source_x) * 4..][..4]);
        }
    }
    (out, output_width as u32, output_height as u32)
}

trait U16Bytes {
    fn from_bytes(pair: &[u8], order: tiff_crate::tags::ByteOrder) -> u16;
    fn to_bytes(self, order: tiff_crate::tags::ByteOrder) -> [u8; 2];
}

impl U16Bytes for u16 {
    fn from_bytes(pair: &[u8], order: tiff_crate::tags::ByteOrder) -> u16 {
        match order {
            tiff_crate::tags::ByteOrder::LittleEndian => u16::from_le_bytes([pair[0], pair[1]]),
            tiff_crate::tags::ByteOrder::BigEndian => u16::from_be_bytes([pair[0], pair[1]]),
        }
    }

    fn to_bytes(self, order: tiff_crate::tags::ByteOrder) -> [u8; 2] {
        match order {
            tiff_crate::tags::ByteOrder::LittleEndian => self.to_le_bytes(),
            tiff_crate::tags::ByteOrder::BigEndian => self.to_be_bytes(),
        }
    }
}

fn gray_opaque(gray: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; gray.len() * 4];
    for (pixel, &value) in rgba.as_chunks_mut::<4>().0.iter_mut().zip(gray) {
        pixel[0] = value;
        pixel[1] = value;
        pixel[2] = value;
        pixel[3] = 255;
    }
    rgba
}

fn unpack_gray(bytes: &[u8], width: usize, pixels: usize, bits: u8) -> Result<Vec<u8>, String> {
    let height = pixels / width;
    let stride = (width * usize::from(bits)).div_ceil(8);
    if bytes.len() != stride * height {
        return Err("Malformed TIFF image data".to_owned());
    }
    let max = (1u16 << bits) - 1;
    let mut rgba = vec![0; pixels * 4];
    for (y, row) in bytes.chunks_exact(stride).enumerate() {
        for x in 0..width {
            let bit = x * usize::from(bits);
            let value = (row[bit / 8] >> (8 - bits - (bit % 8) as u8)) & max as u8;
            let gray = (u16::from(value) * 255 / max) as u8;
            let pixel = &mut rgba[(y * width + x) * 4..][..4];
            pixel[0] = gray;
            pixel[1] = gray;
            pixel[2] = gray;
            pixel[3] = 255;
        }
    }
    Ok(rgba)
}

fn gray16_opaque(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len() * 2];
    for (pixel, pair) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<2>().0)
    {
        let gray = (u16::from_ne_bytes([pair[0], pair[1]]) >> 8) as u8;
        pixel[0] = gray;
        pixel[1] = gray;
        pixel[2] = gray;
        pixel[3] = 255;
    }
    rgba
}

fn gray_alpha(pairs: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; pairs.len() * 2];
    for (pixel, pair) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(pairs.as_chunks::<2>().0)
    {
        pixel[0] = pair[0];
        pixel[1] = pair[0];
        pixel[2] = pair[0];
        pixel[3] = pair[1];
    }
    rgba
}

fn gray_alpha16(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len()];
    for (pixel, quad) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<4>().0)
    {
        pixel[0] = (u16::from_ne_bytes([quad[0], quad[1]]) >> 8) as u8;
        pixel[1] = pixel[0];
        pixel[2] = pixel[0];
        pixel[3] = (u16::from_ne_bytes([quad[2], quad[3]]) >> 8) as u8;
    }
    rgba
}

fn gray_dropped(pairs: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; pairs.len() * 2];
    for (pixel, pair) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(pairs.as_chunks::<2>().0)
    {
        pixel[0] = pair[0];
        pixel[1] = pair[0];
        pixel[2] = pair[0];
        pixel[3] = 255;
    }
    rgba
}

fn gray_dropped16(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len()];
    for (pixel, quad) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<4>().0)
    {
        let gray = (u16::from_ne_bytes([quad[0], quad[1]]) >> 8) as u8;
        pixel[0] = gray;
        pixel[1] = gray;
        pixel[2] = gray;
        pixel[3] = 255;
    }
    rgba
}

fn rgb_opaque(rgb: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; rgb.len() * 4 / 3];
    for (pixel, triple) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(rgb.as_chunks::<3>().0)
    {
        pixel[0] = triple[0];
        pixel[1] = triple[1];
        pixel[2] = triple[2];
        pixel[3] = 255;
    }
    rgba
}

fn rgb16_opaque(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len() * 4 / 6];
    for (pixel, sextet) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<6>().0)
    {
        pixel[0] = (u16::from_ne_bytes([sextet[0], sextet[1]]) >> 8) as u8;
        pixel[1] = (u16::from_ne_bytes([sextet[2], sextet[3]]) >> 8) as u8;
        pixel[2] = (u16::from_ne_bytes([sextet[4], sextet[5]]) >> 8) as u8;
        pixel[3] = 255;
    }
    rgba
}

fn rgba16_to_rgba8(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len() / 2];
    for (pixel, octet) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<8>().0)
    {
        pixel[0] = (u16::from_ne_bytes([octet[0], octet[1]]) >> 8) as u8;
        pixel[1] = (u16::from_ne_bytes([octet[2], octet[3]]) >> 8) as u8;
        pixel[2] = (u16::from_ne_bytes([octet[4], octet[5]]) >> 8) as u8;
        pixel[3] = (u16::from_ne_bytes([octet[6], octet[7]]) >> 8) as u8;
    }
    rgba
}

fn cmyk_opaque(cmyk: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; cmyk.len()];
    for (pixel, quad) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(cmyk.as_chunks::<4>().0)
    {
        let paper = 255 - u16::from(quad[3]);
        pixel[0] = ((255 - u16::from(quad[0])) * paper / 255) as u8;
        pixel[1] = ((255 - u16::from(quad[1])) * paper / 255) as u8;
        pixel[2] = ((255 - u16::from(quad[2])) * paper / 255) as u8;
        pixel[3] = 255;
    }
    rgba
}

fn cmyk16_opaque(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len() / 2];
    for (pixel, octet) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<8>().0)
    {
        let paper = 65535 - u32::from(u16::from_ne_bytes([octet[6], octet[7]]));
        pixel[0] = cmyk16_channel(octet[0], octet[1], paper);
        pixel[1] = cmyk16_channel(octet[2], octet[3], paper);
        pixel[2] = cmyk16_channel(octet[4], octet[5], paper);
        pixel[3] = 255;
    }
    rgba
}

fn cmyka_opaque(cmyka: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; cmyka.len() * 4 / 5];
    for (pixel, quintet) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(cmyka.as_chunks::<5>().0)
    {
        let paper = 255 - u16::from(quintet[3]);
        pixel[0] = ((255 - u16::from(quintet[0])) * paper / 255) as u8;
        pixel[1] = ((255 - u16::from(quintet[1])) * paper / 255) as u8;
        pixel[2] = ((255 - u16::from(quintet[2])) * paper / 255) as u8;
        pixel[3] = quintet[4];
    }
    rgba
}

fn cmyka16_opaque(words: &[u8]) -> Vec<u8> {
    let mut rgba = vec![0; words.len() * 4 / 10];
    for (pixel, samples) in rgba
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(words.as_chunks::<10>().0)
    {
        let paper = 65535 - u32::from(u16::from_ne_bytes([samples[6], samples[7]]));
        pixel[0] = cmyk16_channel(samples[0], samples[1], paper);
        pixel[1] = cmyk16_channel(samples[2], samples[3], paper);
        pixel[2] = cmyk16_channel(samples[4], samples[5], paper);
        pixel[3] = (u16::from_ne_bytes([samples[8], samples[9]]) >> 8) as u8;
    }
    rgba
}

fn cmyk16_channel(high: u8, low: u8, paper: u32) -> u8 {
    ((u32::from(0xFFFF - u16::from_ne_bytes([high, low]))) * paper / (0xFFFF * 256)) as u8
}

fn image_fits_budget(pixels: u64, decoded_bytes: u64, source_bytes: u64) -> bool {
    pixels <= MAX_IMAGE_PIXELS
        && source_bytes <= MAX_TIFF_BYTES
        && source_bytes.saturating_add(decoded_bytes) <= MAX_IMAGE_BYTES
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;
    use tiff_crate::encoder::{Compression, TiffEncoder, colortype};

    fn encode<Color>(
        width: u32,
        height: u32,
        data: &[Color::Inner],
        compression: Compression,
    ) -> Vec<u8>
    where
        Color: colortype::ColorType,
        [Color::Inner]: tiff_crate::encoder::TiffValue,
    {
        let mut cursor = Cursor::new(Vec::new());
        TiffEncoder::new(&mut cursor)
            .unwrap()
            .with_compression(compression)
            .write_image::<Color>(width, height, data)
            .unwrap();
        cursor.into_inner()
    }

    fn decode_png(png: &[u8]) -> (u32, u32, Vec<u8>) {
        let mut reader = png::Decoder::new(Cursor::new(png)).read_info().unwrap();
        let (width, height) = (reader.info().width, reader.info().height);
        let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
        let info = reader.next_frame(&mut pixels).unwrap();
        assert_eq!(info.color_type, png::ColorType::Rgba);
        assert_eq!(info.bit_depth, png::BitDepth::Eight);
        pixels.truncate(info.buffer_size());
        (width, height, pixels)
    }

    fn round_trip(tiff: &[u8]) -> (u32, u32, Vec<u8>) {
        decode_png(&decode_tiff_png(tiff).unwrap())
    }

    /// Minimal single-strip little-endian TIFF with a trailing color map.
    fn minimal_tiff(
        width: u16,
        height: u16,
        photometric: u16,
        compression: u16,
        bits: u16,
        colormap: &[u16],
        pixels: &[u8],
    ) -> Vec<u8> {
        let mut entries: Vec<(u16, u16, u32, u32)> = vec![
            (256, 3, 1, u32::from(width)),
            (257, 3, 1, u32::from(height)),
            (258, 3, 1, u32::from(bits)),
            (259, 3, 1, u32::from(compression)),
            (262, 3, 1, u32::from(photometric)),
            (273, 4, 1, 0),
            (277, 3, 1, 1),
            (278, 3, 1, u32::from(height)),
            (279, 4, 1, pixels.len() as u32),
        ];
        if !colormap.is_empty() {
            entries.push((320, 3, colormap.len() as u32, 0));
        }
        let strip_offset = 8 + 2 + entries.len() * 12 + 4;
        let colormap_offset = strip_offset + pixels.len();
        let mut tiff = Vec::new();
        tiff.extend_from_slice(&[0x49, 0x49, 42, 0, 8, 0, 0, 0]);
        tiff.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        for (tag, kind, count, value) in &entries {
            let value = match tag {
                273 => strip_offset as u32,
                320 => colormap_offset as u32,
                _ => *value,
            };
            tiff.extend_from_slice(&tag.to_le_bytes());
            tiff.extend_from_slice(&kind.to_le_bytes());
            tiff.extend_from_slice(&count.to_le_bytes());
            tiff.extend_from_slice(&value.to_le_bytes());
        }
        tiff.extend_from_slice(&[0, 0, 0, 0]);
        tiff.extend_from_slice(pixels);
        for value in colormap {
            tiff.extend_from_slice(&value.to_le_bytes());
        }
        tiff
    }

    #[test]
    fn decodes_tiff_to_png() {
        let tiff = encode::<colortype::RGBA8>(
            2,
            1,
            &[255, 0, 0, 255, 0, 128, 255, 64],
            Compression::Uncompressed,
        );
        let (width, height, pixels) = round_trip(&tiff);
        assert_eq!((width, height), (2, 1));
        assert_eq!(pixels, &[255, 0, 0, 255, 0, 128, 255, 64]);
    }

    #[test]
    fn decodes_gray8_rgb8_and_rgba8() {
        let (width, height, pixels) = round_trip(&encode::<colortype::Gray8>(
            2,
            1,
            &[0, 255],
            Compression::Uncompressed,
        ));
        assert_eq!((width, height), (2, 1));
        assert_eq!(pixels, &[0, 0, 0, 255, 255, 255, 255, 255]);
        let (_, _, pixels) = round_trip(&encode::<colortype::RGB8>(
            2,
            1,
            &[255, 0, 0, 0, 255, 0],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[255, 0, 0, 255, 0, 255, 0, 255]);
        let (_, _, pixels) = round_trip(&encode::<colortype::RGBA8>(
            1,
            1,
            &[10, 20, 30, 40],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[10, 20, 30, 40]);
    }

    #[test]
    fn decodes_16_bit_samples() {
        let (_, _, pixels) = round_trip(&encode::<colortype::Gray16>(
            2,
            1,
            &[0, 65535],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[0, 0, 0, 255, 255, 255, 255, 255]);
        let (_, _, pixels) = round_trip(&encode::<colortype::RGB16>(
            1,
            1,
            &[65535, 0, 32896],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[255, 0, 128, 255]);
        let (_, _, pixels) = round_trip(&encode::<colortype::RGBA16>(
            1,
            1,
            &[65535, 0, 0, 32768],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[255, 0, 0, 128]);
    }

    #[test]
    fn decodes_lzw_packbits_and_deflate() {
        let data = [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255];
        for compression in [
            Compression::Lzw,
            Compression::Packbits,
            Compression::Deflate(tiff_crate::encoder::DeflateLevel::Fast),
        ] {
            let (_, _, pixels) = round_trip(&encode::<colortype::RGB8>(2, 2, &data, compression));
            assert_eq!(
                pixels,
                &[
                    255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255
                ]
            );
        }
    }

    #[test]
    fn decodes_palette() {
        let mut colormap = vec![0u16; 3 * 256];
        colormap[1] = 0xFF00;
        colormap[256 + 2] = 0xFF00;
        colormap[2 * 256] = 0xFF00;
        colormap[2 * 256 + 3] = 0xFF00;
        let tiff = minimal_tiff(4, 1, 3, 1, 8, &colormap, &[0, 1, 2, 3]);
        let (width, height, pixels) = round_trip(&tiff);
        assert_eq!((width, height), (4, 1));
        assert_eq!(
            pixels,
            &[
                0, 0, 255, 255, 255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255
            ]
        );
    }

    #[test]
    fn converts_cmyk_naively() {
        let (_, _, pixels) = round_trip(&encode::<colortype::CMYK8>(
            1,
            1,
            &[255, 0, 0, 0],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[0, 255, 255, 255]);
        let (_, _, pixels) = round_trip(&encode::<colortype::CMYK8>(
            1,
            1,
            &[0, 0, 0, 255],
            Compression::Uncompressed,
        ));
        assert_eq!(pixels, &[0, 0, 0, 255]);
    }

    #[test]
    fn applies_orientation() {
        let rgba = vec![1, 2, 3, 4, 5, 6, 7, 8];
        let (identity, width, height) = apply_orientation(rgba.clone(), 2, 1, 1);
        assert_eq!((identity, width, height), (rgba.clone(), 2, 1));
        let (flipped, width, height) = apply_orientation(rgba.clone(), 2, 1, 2);
        assert_eq!(
            (flipped, width, height),
            (vec![5, 6, 7, 8, 1, 2, 3, 4], 2, 1)
        );
        let (rotated, width, height) = apply_orientation(rgba.clone(), 2, 1, 6);
        assert_eq!(
            (rotated, width, height),
            (vec![1, 2, 3, 4, 5, 6, 7, 8], 1, 2)
        );
        let (upside, _, _) = apply_orientation(rgba.clone(), 2, 1, 3);
        assert_eq!(upside, vec![5, 6, 7, 8, 1, 2, 3, 4]);
        let (unknown, width, height) = apply_orientation(rgba.clone(), 2, 1, 9);
        assert_eq!((unknown, width, height), (rgba, 2, 1));
    }

    #[test]
    fn refuses_unsupported_compression() {
        for compression in [0xC350, 0xC351] {
            let tiff = minimal_tiff(1, 1, 1, compression, 8, &[], &[0, 0]);
            let error = decode_tiff_png(&tiff).unwrap_err();
            assert!(error.contains("unsupported"), "unexpected error: {error}");
        }
        // Dev builds enable the default codec set, so JPEG fails while parsing
        // instead; the wasm build refuses it with the same plain error as above.
        let tiff = minimal_tiff(1, 1, 1, 7, 8, &[], &[0, 0]);
        assert!(decode_tiff_png(&tiff).is_err());
    }

    #[test]
    fn rejects_invalid_tiff() {
        assert!(decode_tiff_png(b"II*\0").is_err());
    }

    #[test]
    fn rejects_images_past_the_decode_budget() {
        assert!(!image_fits_budget(MAX_IMAGE_PIXELS + 1, 0, 0));
        assert!(!image_fits_budget(0, MAX_IMAGE_BYTES + 1, 0));
        assert!(!image_fits_budget(0, 0, MAX_TIFF_BYTES + 1));
        assert!(!image_fits_budget(0, MAX_IMAGE_BYTES, 1));
    }
}
