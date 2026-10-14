//! Device-independent bitmaps (`BITMAPINFO` plus bits) decoded to RGBA.

use crate::read::{i32_at, u16_at, u32_at};

/// Largest bitmap one record may carry.
pub(crate) const MAX_BITMAP_PIXELS: u64 = 16_777_216;

const BI_RGB: u32 = 0;
const BI_RLE8: u32 = 1;
const BI_RLE4: u32 = 2;
const BI_BITFIELDS: u32 = 3;
const BI_JPEG: u32 = 4;
const BI_PNG: u32 = 5;
const BI_ALPHABITFIELDS: u32 = 6;

/// A decoded DIB, top row first.
pub(crate) struct Dib {
    pub width: u32,
    pub height: u32,
    pub bottom_up: bool,
    pub pixels: DibPixels,
}

pub(crate) enum DibPixels {
    Rgba(Vec<u8>),
    Encoded { mime: &'static str, bytes: Vec<u8> },
}

struct Header {
    size: usize,
    width: u32,
    height: u32,
    bottom_up: bool,
    bits: u16,
    compression: u32,
    colors: usize,
    masks: Option<[u32; 4]>,
    core: bool,
}

fn header(bmi: &[u8]) -> Result<Header, &'static str> {
    let malformed = "a bitmap header is malformed";
    let size = u32_at(bmi, 0).ok_or(malformed)? as usize;
    if size == 12 {
        let width = u32::from(u16_at(bmi, 4).ok_or(malformed)?);
        let height = u32::from(u16_at(bmi, 6).ok_or(malformed)?);
        let bits = u16_at(bmi, 10).ok_or(malformed)?;
        if width == 0 || height == 0 {
            return Err("a bitmap has no pixels");
        }
        return Ok(Header {
            size,
            width,
            height,
            bottom_up: true,
            bits,
            compression: BI_RGB,
            colors: if bits <= 8 { 1 << bits } else { 0 },
            masks: None,
            core: true,
        });
    }
    if !(40..=124).contains(&size) || bmi.len() < size {
        return Err(malformed);
    }
    let width = i32_at(bmi, 4).ok_or(malformed)?;
    let height = i32_at(bmi, 8).ok_or(malformed)?;
    let bits = u16_at(bmi, 14).ok_or(malformed)?;
    let compression = u32_at(bmi, 16).ok_or(malformed)?;
    let used = u32_at(bmi, 32).ok_or(malformed)? as usize;
    if width <= 0 || height == 0 || height == i32::MIN {
        return Err("a bitmap has no pixels");
    }
    let masks = if matches!(compression, BI_BITFIELDS | BI_ALPHABITFIELDS) {
        let at = if size >= 52 { 40 } else { size };
        let read = |index: usize| u32_at(bmi, at + index * 4).ok_or(malformed);
        let alpha = if compression == BI_ALPHABITFIELDS || size >= 56 {
            read(3)?
        } else {
            0
        };
        Some([read(0)?, read(1)?, read(2)?, alpha])
    } else {
        None
    };
    let masks_size = match (masks.is_some(), size) {
        (false, _) => 0,
        (true, 40) if compression == BI_ALPHABITFIELDS => 16,
        (true, 40) => 12,
        _ => 0,
    };
    let colors = match (used, bits) {
        (0, 1 | 4 | 8) => 1 << bits,
        (used, 1 | 4 | 8) => used.min(1 << bits),
        (used, _) => used.min(256),
    };
    Ok(Header {
        size: size + masks_size,
        width: width as u32,
        height: height.unsigned_abs(),
        bottom_up: height > 0,
        bits,
        compression,
        colors,
        masks,
        core: false,
    })
}

/// The bytes a packed DIB (header, colour table, bits) spends before its bits.
pub(crate) fn packed_bits_offset(dib: &[u8]) -> Option<usize> {
    let header = header(dib).ok()?;
    Some(header.size + header.colors * if header.core { 3 } else { 4 })
}

/// Decodes `bmi` and `bits`, refusing past `budget` pixels.
pub(crate) fn decode(
    bmi: &[u8],
    bits: &[u8],
    usage: u32,
    budget: u64,
    charge: impl FnOnce(u64) -> bool,
) -> Result<Dib, &'static str> {
    let header = header(bmi)?;
    let pixels = u64::from(header.width) * u64::from(header.height);
    if pixels > MAX_BITMAP_PIXELS || pixels > budget {
        return Err("a bitmap exceeds the pixel budget");
    }
    if matches!(header.compression, BI_JPEG | BI_PNG) {
        let mime = if header.compression == BI_PNG {
            "image/png"
        } else {
            "image/jpeg"
        };
        let (width, height) = encoded_size(bits, mime).ok_or("an embedded image is malformed")?;
        let pixels = u64::from(width) * u64::from(height);
        if pixels > budget.min(MAX_BITMAP_PIXELS) {
            return Err("a bitmap exceeds the pixel budget");
        }
        if !charge(pixels) {
            return Err("a bitmap exceeds the pixel budget");
        }
        return Ok(Dib {
            width,
            height,
            bottom_up: false,
            pixels: DibPixels::Encoded {
                mime,
                bytes: bits.to_vec(),
            },
        });
    }
    if !charge(pixels) {
        return Err("a bitmap exceeds the pixel budget");
    }
    if usage != 0 && header.bits <= 8 {
        return Err("a bitmap indexes a palette the metafile does not carry");
    }
    let entry = if header.core { 3 } else { 4 };
    let mut palette = Vec::with_capacity(header.colors);
    for index in 0..header.colors {
        let at = header.size + index * entry;
        let color = bmi
            .get(at..at + 3)
            .ok_or("a bitmap colour table is truncated")?;
        palette.push([color[2], color[1], color[0]]);
    }
    let (width, height) = (header.width as usize, header.height as usize);
    let mut rgba = vec![0u8; width * height * 4];
    match (header.compression, header.bits) {
        (BI_RLE8, 8) | (BI_RLE4, 4) => rle(&mut rgba, bits, &header, &palette)?,
        (BI_RGB | BI_BITFIELDS | BI_ALPHABITFIELDS, 1 | 4 | 8 | 16 | 24 | 32) => {
            uncompressed(&mut rgba, bits, &header, &palette)?
        }
        _ => return Err("a bitmap uses an unsupported encoding"),
    }
    if header.bottom_up {
        flip_rows(&mut rgba, width * 4);
    }
    Ok(Dib {
        width: header.width,
        height: header.height,
        bottom_up: header.bottom_up,
        pixels: DibPixels::Rgba(rgba),
    })
}

fn flip_rows(rgba: &mut [u8], stride: usize) {
    let rows = rgba.len() / stride;
    for row in 0..rows / 2 {
        let (top, bottom) = rgba.split_at_mut((rows - 1 - row) * stride);
        top[row * stride..row * stride + stride].swap_with_slice(&mut bottom[..stride]);
    }
}

fn uncompressed(
    rgba: &mut [u8],
    bits: &[u8],
    header: &Header,
    palette: &[[u8; 3]],
) -> Result<(), &'static str> {
    let (width, height) = (header.width as usize, header.height as usize);
    let bpp = usize::from(header.bits);
    let stride = (width * bpp).div_ceil(32) * 4;
    if bits.len() < stride * height {
        return Err("a bitmap's bits are truncated");
    }
    let masks = match (header.masks, bpp) {
        (Some(masks), 16 | 32) => Some(masks),
        (None, 16) => Some([0x7C00, 0x03E0, 0x001F, 0]),
        _ => None,
    };
    let color = |index: usize| palette.get(index).copied().unwrap_or([0, 0, 0]);
    for row in 0..height {
        let line = &bits[row * stride..row * stride + stride];
        let out = &mut rgba[row * width * 4..(row + 1) * width * 4];
        for x in 0..width {
            let pixel: [u8; 4] = match bpp {
                1 | 4 | 8 => {
                    let bit = x * bpp;
                    let byte = line[bit / 8];
                    let shift = 8 - bpp - (bit % 8);
                    let index = usize::from((byte >> shift) & ((1u16 << bpp) - 1) as u8);
                    let [r, g, b] = color(index);
                    [r, g, b, 255]
                }
                24 => [line[x * 3 + 2], line[x * 3 + 1], line[x * 3], 255],
                16 | 32 => {
                    let value = if bpp == 16 {
                        u32::from(u16::from_le_bytes([line[x * 2], line[x * 2 + 1]]))
                    } else {
                        u32::from_le_bytes([
                            line[x * 4],
                            line[x * 4 + 1],
                            line[x * 4 + 2],
                            line[x * 4 + 3],
                        ])
                    };
                    match masks {
                        Some([r, g, b, a]) => [
                            channel(value, r),
                            channel(value, g),
                            channel(value, b),
                            if a == 0 { 255 } else { channel(value, a) },
                        ],
                        None => [
                            (value >> 16) as u8,
                            (value >> 8) as u8,
                            value as u8,
                            (value >> 24) as u8,
                        ],
                    }
                }
                _ => [0, 0, 0, 255],
            };
            out[x * 4..x * 4 + 4].copy_from_slice(&pixel);
        }
    }
    Ok(())
}

/// A masked channel scaled to eight bits.
fn channel(value: u32, mask: u32) -> u8 {
    if mask == 0 {
        return 0;
    }
    let shift = mask.trailing_zeros();
    let max = mask >> shift;
    let raw = (value & mask) >> shift;
    ((u64::from(raw) * 255 + u64::from(max) / 2) / u64::from(max)) as u8
}

fn rle(
    rgba: &mut [u8],
    bits: &[u8],
    header: &Header,
    palette: &[[u8; 3]],
) -> Result<(), &'static str> {
    let (width, height) = (header.width as usize, header.height as usize);
    let four = header.compression == BI_RLE4;
    let (mut x, mut y) = (0usize, 0usize);
    let mut at = 0usize;
    let mut put = |x: usize, y: usize, index: usize| {
        if x < width && y < height {
            let [r, g, b] = palette.get(index).copied().unwrap_or([0, 0, 0]);
            let offset = (y * width + x) * 4;
            rgba[offset..offset + 4].copy_from_slice(&[r, g, b, 255]);
        }
    };
    while at + 1 < bits.len() {
        let (count, value) = (bits[at] as usize, bits[at + 1]);
        at += 2;
        if count > 0 {
            for step in 0..count.min(width.saturating_sub(x)) {
                let index = if four {
                    if step % 2 == 0 {
                        value >> 4
                    } else {
                        value & 0x0f
                    }
                } else {
                    value
                };
                put(x + step, y, usize::from(index));
            }
            x = x.saturating_add(count);
            continue;
        }
        match value {
            0 => {
                x = 0;
                y += 1;
            }
            1 => break,
            2 => {
                let (dx, dy) = (
                    *bits.get(at).ok_or("a bitmap's bits are truncated")? as usize,
                    *bits.get(at + 1).ok_or("a bitmap's bits are truncated")? as usize,
                );
                at += 2;
                x = x.saturating_add(dx);
                y += dy;
            }
            literal => {
                let literal = usize::from(literal);
                let bytes = if four { literal.div_ceil(2) } else { literal };
                let run = bits
                    .get(at..at + bytes)
                    .ok_or("a bitmap's bits are truncated")?;
                for step in 0..literal.min(width.saturating_sub(x)) {
                    let index = if four {
                        let byte = run[step / 2];
                        if step % 2 == 0 {
                            byte >> 4
                        } else {
                            byte & 0x0f
                        }
                    } else {
                        run[step]
                    };
                    put(x + step, y, usize::from(index));
                }
                x = x.saturating_add(literal);
                at += bytes.div_ceil(2) * 2;
            }
        }
        if y >= height {
            break;
        }
    }
    Ok(())
}

/// Pixel dimensions from a PNG or JPEG stream's header.
pub(crate) fn encoded_size(bytes: &[u8], mime: &str) -> Option<(u32, u32)> {
    let (width, height) = if mime == "image/png" {
        if bytes.get(..8)? != b"\x89PNG\r\n\x1a\n" || bytes.get(12..16)? != b"IHDR" {
            return None;
        }
        let read = |at: usize| Some(u32::from_be_bytes(bytes.get(at..at + 4)?.try_into().ok()?));
        (read(16)?, read(20)?)
    } else {
        if bytes.get(..2)? != [0xFF, 0xD8] {
            return None;
        }
        let mut at = 2;
        loop {
            while *bytes.get(at)? != 0xFF {
                at += 1;
            }
            let marker = *bytes.get(at + 1)?;
            if marker == 0xFF {
                at += 1;
                continue;
            }
            let length = usize::from(u16::from_be_bytes([
                *bytes.get(at + 2)?,
                *bytes.get(at + 3)?,
            ]));
            if matches!(marker, 0xC0..=0xCF) && !matches!(marker, 0xC4 | 0xC8 | 0xCC) {
                let height = u16::from_be_bytes([*bytes.get(at + 5)?, *bytes.get(at + 6)?]);
                let width = u16::from_be_bytes([*bytes.get(at + 7)?, *bytes.get(at + 8)?]);
                break (u32::from(width), u32::from(height));
            }
            if length < 2 {
                return None;
            }
            at += 2 + length;
        }
    };
    (width > 0 && height > 0).then_some((width, height))
}
