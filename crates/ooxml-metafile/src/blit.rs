//! Bitmap blits and pattern brushes.

use std::sync::Arc;

use crate::dib::{Dib, DibPixels, decode, packed_bits_offset};
use crate::drawing::{Bitmap, ClipRegion, Image, Op, Paint, PathCommand, Pixels, Rgba};
use crate::player::{Brush, GdiObject, Player, Xform, apply, chain, concat};
use crate::read::{finite_at, i16_at, i32_at, u16_at, u32_at};

const SRCCOPY: u32 = 0x00CC_0020;
const SRCPAINT: u32 = 0x00EE_0086;
const SRCAND: u32 = 0x0088_00C6;
const SRCINVERT: u32 = 0x0066_0046;
const NOTSRCCOPY: u32 = 0x0033_0008;
const MERGECOPY: u32 = 0x00C0_00CA;
const PATCOPY: u32 = 0x00F0_0021;
const DSTCOPY: u32 = 0x00AA_0029;
const BLACKNESS: u32 = 0x0000_0042;
const WHITENESS: u32 = 0x00FF_0062;
/// `PSDPxax`: the pattern where the source is black, the destination elsewhere.
const PSDPXAX: u32 = 0x00B8_074A;

/// Where a blit's source rectangle starts counting rows.
#[derive(Clone, Copy, PartialEq)]
enum Rows {
    /// From the top row, as a source device context addresses its bitmap.
    Top,
    /// From the DIB's first scan line, as `StretchDIBits` does.
    Scan,
}

struct Blit<'a> {
    dest: (f64, f64, f64, f64),
    source: (f64, f64, f64, f64),
    rows: Rows,
    bmi: &'a [u8],
    bits: &'a [u8],
    usage: u32,
    rop: u32,
    opacity: f64,
    alpha: bool,
    transparent: Option<u32>,
    /// A parallelogram's three corners in logical units, replacing `dest`.
    corners: Option<[(f64, f64); 3]>,
}

fn slice(bytes: &[u8], offset: u32, size: u32) -> Option<&[u8]> {
    crate::read::record_span(bytes, offset as usize, size as usize, 1)
}

pub(crate) fn blit<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    kind: u32,
    body: usize,
) -> Option<()> {
    let int = |at: usize| i32_at(bytes, body + at).map(f64::from);
    let word = |at: usize| u32_at(bytes, body + at);
    player.flush_pending();
    match kind {
        76 | 77 | 78 | 114 | 116 => {
            let dest = (int(16)?, int(20)?, int(24)?, int(28)?);
            let op = word(32)?;
            let (off_bmi, cb_bmi, off_bits, cb_bits) = (word(76)?, word(80)?, word(84)?, word(88)?);
            if cb_bmi == 0 || cb_bits == 0 {
                return sourceless(
                    player,
                    dest,
                    if kind == 76 || kind == 77 {
                        op
                    } else {
                        DSTCOPY
                    },
                );
            }
            let source_size = if kind == 76 || kind == 78 {
                let scale = |at| finite_at(bytes, body + at).filter(|value| *value != 0.0);
                (
                    dest.2 * scale(44).unwrap_or(1.0).abs(),
                    dest.3 * scale(56).unwrap_or(1.0).abs(),
                )
            } else {
                (int(92)?, int(96)?)
            };
            if kind == 78 {
                player.omit("bitmap masks")?;
            }
            let blend = op.to_le_bytes();
            let job = Blit {
                dest,
                source: (int(36)?, int(40)?, source_size.0, source_size.1),
                rows: Rows::Top,
                bmi: slice(bytes, off_bmi, cb_bmi)?,
                bits: slice(bytes, off_bits, cb_bits)?,
                usage: word(72)?,
                rop: match kind {
                    114 | 116 => SRCCOPY,
                    _ => op,
                },
                opacity: if kind == 114 {
                    f64::from(blend[2]) / 255.0
                } else {
                    1.0
                },
                alpha: kind == 114 && blend[3] & 1 != 0,
                transparent: (kind == 116).then_some(op),
                corners: None,
            };
            draw(player, job)
        }
        79 => {
            let corner = |at: usize| Some((int(at)?, int(at + 4)?));
            let corners = [corner(16)?, corner(24)?, corner(32)?];
            let (off_bmi, cb_bmi, off_bits, cb_bits) =
                (word(88)?, word(92)?, word(96)?, word(100)?);
            if cb_bmi == 0 || cb_bits == 0 {
                return Some(());
            }
            if word(120)? != 0 {
                player.omit("bitmap masks")?;
            }
            let job = Blit {
                dest: (0.0, 0.0, 1.0, 1.0),
                source: (int(40)?, int(44)?, int(48)?, int(52)?),
                rows: Rows::Top,
                bmi: slice(bytes, off_bmi, cb_bmi)?,
                bits: slice(bytes, off_bits, cb_bits)?,
                usage: word(84)?,
                rop: SRCCOPY,
                opacity: 1.0,
                alpha: false,
                transparent: None,
                corners: Some(corners),
            };
            draw(player, job)
        }
        80 => {
            let (off_bmi, cb_bmi, off_bits, cb_bits) = (word(40)?, word(44)?, word(48)?, word(52)?);
            let (start, scans) = (word(60)?, word(64)?);
            let (width, height) = (int(32)?, int(36)?);
            let bmi = slice(bytes, off_bmi, cb_bmi)?;
            let rows = i32_at(bmi, 8)?.unsigned_abs();
            if start != 0 || scans < rows {
                player.omit("partial device bitmaps")?;
                return Some(());
            }
            let job = Blit {
                dest: (int(16)?, int(20)?, width, height),
                source: (int(24)?, int(28)?, width, height),
                rows: Rows::Scan,
                bmi,
                bits: slice(bytes, off_bits, cb_bits)?,
                usage: word(56)?,
                rop: SRCCOPY,
                opacity: 1.0,
                alpha: false,
                transparent: None,
                corners: None,
            };
            draw(player, job)
        }
        _ => {
            let (off_bmi, cb_bmi, off_bits, cb_bits) = (word(40)?, word(44)?, word(48)?, word(52)?);
            let dest = (int(16)?, int(20)?, int(64)?, int(68)?);
            let rop = word(60)?;
            if cb_bmi == 0 || cb_bits == 0 {
                return sourceless(player, dest, rop);
            }
            let job = Blit {
                dest,
                source: (int(24)?, int(28)?, int(32)?, int(36)?),
                rows: Rows::Scan,
                bmi: slice(bytes, off_bmi, cb_bmi)?,
                bits: slice(bytes, off_bits, cb_bits)?,
                usage: word(56)?,
                rop,
                opacity: 1.0,
                alpha: false,
                transparent: None,
                corners: None,
            };
            draw(player, job)
        }
    }
}

pub(crate) fn wmf_blit<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    function: usize,
    body: usize,
) -> Option<()> {
    let word = |at: usize| i16_at(bytes, body + at).map(f64::from);
    let rop = u32_at(bytes, body)?;
    player.flush_pending();
    let bare = bytes.len() == ((function >> 8) + 3) * 2;
    let (dest, source, dib) = match function {
        0x061D => {
            let dest = (word(10)?, word(8)?, word(6)?, word(4)?);
            return sourceless(player, dest, rop);
        }
        0x0940 if bare => {
            let dest = (word(16)?, word(14)?, word(12)?, word(10)?);
            return sourceless(player, dest, rop);
        }
        0x0B41 if bare => {
            let dest = (word(20)?, word(18)?, word(16)?, word(14)?);
            return sourceless(player, dest, rop);
        }
        0x0940 => {
            let (height, width) = (word(8)?, word(10)?);
            (
                (word(14)?, word(12)?, width, height),
                (word(6)?, word(4)?, width, height),
                body + 16,
            )
        }
        0x0B41 => (
            (word(18)?, word(16)?, word(14)?, word(12)?),
            (word(10)?, word(8)?, word(6)?, word(4)?),
            body + 20,
        ),
        0x0F43 => (
            (word(20)?, word(18)?, word(16)?, word(14)?),
            (word(12)?, word(10)?, word(8)?, word(6)?),
            body + 22,
        ),
        0x0D33 => {
            let (height, width) = (word(10)?, word(12)?);
            if word(4)? != 0.0 {
                player.omit("partial device bitmaps")?;
                return Some(());
            }
            (
                (word(16)?, word(14)?, width, height),
                (word(8)?, word(6)?, width, height),
                body + 18,
            )
        }
        _ => return player.omit("device-dependent bitmaps"),
    };
    let packed = bytes.get(dib..)?;
    let Some(offset) = packed_bits_offset(packed) else {
        return player.omit("bitmaps that could not be decoded");
    };
    let usage = match function {
        0x0F43 => u32::from(u16_at(bytes, body + 4)?),
        0x0D33 => u32::from(u16_at(bytes, body)?),
        _ => 0,
    };
    let job = Blit {
        dest,
        source,
        rows: if matches!(function, 0x0F43 | 0x0D33) {
            Rows::Scan
        } else {
            Rows::Top
        },
        bmi: &packed[..offset.min(packed.len())],
        bits: packed.get(offset..)?,
        usage,
        rop: if function == 0x0D33 { SRCCOPY } else { rop },
        opacity: 1.0,
        alpha: false,
        transparent: None,
        corners: None,
    };
    draw(player, job)
}

/// The raster operations that paint without reading a source bitmap.
fn sourceless<const FULL: bool>(
    player: &mut Player<FULL>,
    dest: (f64, f64, f64, f64),
    rop: u32,
) -> Option<()> {
    let fill = match rop {
        DSTCOPY => return Some(()),
        PATCOPY => player.brush_fill(),
        BLACKNESS => Some(Paint::Solid(Rgba::BLACK)),
        WHITENESS => Some(Paint::Solid(Rgba::WHITE)),
        _ => return player.omit("raster operations that read the destination"),
    };
    let (x, y, width, height) = dest;
    if width == 0.0 || height == 0.0 {
        return Some(());
    }
    let saved = std::mem::take(&mut player.path);
    player.path = player.logical_rect_path((x, y, x + width, y + height));
    player.emit(fill, None);
    player.path = saved;
    Some(())
}

fn draw<const FULL: bool>(player: &mut Player<FULL>, job: Blit<'_>) -> Option<()> {
    let recolor = match job.rop {
        SRCCOPY | MERGECOPY => Recolor::None,
        SRCPAINT => Recolor::ClearBlack,
        SRCAND => Recolor::ClearWhite,
        SRCINVERT => {
            player.omit("XOR blits drawn as overlays")?;
            Recolor::ClearBlack
        }
        NOTSRCCOPY => Recolor::Invert,
        PSDPXAX => {
            let color = match player.brush_fill() {
                Some(Paint::Solid(color)) => color,
                _ => Rgba::BLACK,
            };
            Recolor::Stencil(color)
        }
        DSTCOPY => return Some(()),
        PATCOPY | BLACKNESS | WHITENESS => {
            return sourceless(player, job.dest, job.rop);
        }
        _ => {
            player.omit("raster operations drawn as copies")?;
            Recolor::None
        }
    };
    let budget = player
        .limits
        .bitmap_pixels
        .saturating_sub(player.bitmap_pixels);
    let declared = declared_pixels(job.bmi);
    if declared > budget {
        return player.refuse("the metafile's bitmaps exceed the pixel budget");
    }
    player.charge(0, declared)?;
    let dib = match decode(job.bmi, job.bits, job.usage, budget, |pixels| {
        player.charge(0, pixels.saturating_sub(declared)).is_some()
    }) {
        Ok(dib) => dib,
        Err(_) if player.overflowed => return None,
        Err(_) => return player.omit("bitmaps that could not be decoded"),
    };
    let Some((bitmap, offset, whole)) = crop(
        dib,
        job.source,
        job.rows,
        recolor,
        job.alpha,
        job.transparent,
    ) else {
        return Some(());
    };
    let (sw, sh) = (job.source.2.abs().max(1.0), job.source.3.abs().max(1.0));
    let (ox, oy) = offset;
    let placed: Xform = match job.corners {
        Some([a, b, c]) => {
            let (u, v) = ((b.0 - a.0) / sw, (b.1 - a.1) / sw);
            let (p, q) = ((c.0 - a.0) / sh, (c.1 - a.1) / sh);
            [u, v, p, q, a.0 + u * ox + p * oy, a.1 + v * ox + q * oy]
        }
        None => {
            let (x, y, w, h) = job.dest;
            let (kx, ky) = (w / sw, h / sh);
            [kx, 0.0, 0.0, ky, x + ox * kx, y + oy * ky]
        }
    };
    let to_output = player.logical_to_output();
    let transform = concat(placed, to_output);
    let mut clip = player.dc.clip.clone();
    if whole {
        let corners = match job.corners {
            Some([a, b, c]) => [a, b, (b.0 + c.0 - a.0, b.1 + c.1 - a.1), c],
            None => {
                let (x, y, w, h) = job.dest;
                [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
            }
        };
        player.charge(5, 0)?;
        clip = chain(
            clip,
            parallelogram(corners.map(|corner| apply(to_output, corner))),
        );
    }
    player.push_op(Op::Image(Image {
        transform,
        bitmap: Arc::new(bitmap),
        opacity: job.opacity.clamp(0.0, 1.0),
        clip,
    }));
    Some(())
}

fn declared_pixels(bmi: &[u8]) -> u64 {
    if u32_at(bmi, 0) == Some(12) {
        let side = |at| u64::from(u16_at(bmi, at).unwrap_or(0));
        return side(4) * side(6);
    }
    let side = |at| u64::from(i32_at(bmi, at).unwrap_or(0).unsigned_abs());
    side(4) * side(8)
}

#[derive(Clone, Copy)]
enum Recolor {
    None,
    ClearBlack,
    ClearWhite,
    Invert,
    Stencil(Rgba),
}

/// A clip region bounded by four corners in output units.
pub(crate) fn parallelogram(corners: [(f64, f64); 4]) -> ClipRegion {
    let mut path: Vec<PathCommand> = corners
        .iter()
        .enumerate()
        .map(|(index, &(x, y))| {
            if index == 0 {
                PathCommand::Move { x, y }
            } else {
                PathCommand::Line { x, y }
            }
        })
        .collect();
    path.push(PathCommand::Close);
    ClipRegion {
        path,
        even_odd: false,
        exclude: false,
    }
}

/// The part of `dib` a source rectangle selects, recoloured for its raster
/// operation, with where that part starts inside the rectangle, and whether
/// it is the whole of a compressed bitmap that the destination must clip.
fn crop(
    dib: Dib,
    source: (f64, f64, f64, f64),
    rows: Rows,
    recolor: Recolor,
    alpha: bool,
    transparent: Option<u32>,
) -> Option<(Bitmap, (f64, f64), bool)> {
    let (width, height) = (f64::from(dib.width), f64::from(dib.height));
    let (x, y, w, h) = source;
    let (x, w) = if w < 0.0 { (x + w, -w) } else { (x, w) };
    let (y, h) = if h < 0.0 { (y + h, -h) } else { (y, h) };
    let top = if rows == Rows::Scan && dib.bottom_up {
        height - y - h
    } else {
        y
    };
    let (x0, y0) = (x.max(0.0).floor(), top.max(0.0).floor());
    let (x1, y1) = ((x + w).min(width).ceil(), (top + h).min(height).ceil());
    if x1 <= x0 || y1 <= y0 {
        return None;
    }
    let full = (x, top, w, h) == (0.0, 0.0, width, height);
    let offset = match dib.pixels {
        DibPixels::Encoded { .. } => (-x, -top),
        DibPixels::Rgba(_) => (x0 - x, y0 - top),
    };
    let pixels = match dib.pixels {
        DibPixels::Encoded { mime, bytes } => Pixels::Encoded { mime, bytes },
        DibPixels::Rgba(rgba) => {
            let (cw, ch) = ((x1 - x0) as usize, (y1 - y0) as usize);
            let stride = dib.width as usize * 4;
            let mut out = Vec::with_capacity(cw * ch * 4);
            for row in y0 as usize..y1 as usize {
                let start = row * stride + x0 as usize * 4;
                out.extend_from_slice(&rgba[start..start + cw * 4]);
            }
            for pixel in out.as_chunks_mut::<4>().0 {
                let [r, g, b, a] = *pixel;
                if !alpha {
                    pixel[3] = 255;
                } else if a > 0 && a < 255 {
                    let unpremultiply = |c: u8| {
                        ((u16::from(c) * 255 + u16::from(a) / 2) / u16::from(a)).min(255) as u8
                    };
                    pixel[0] = unpremultiply(r);
                    pixel[1] = unpremultiply(g);
                    pixel[2] = unpremultiply(b);
                }
                let (r, g, b) = (pixel[0], pixel[1], pixel[2]);
                if let Some(key) = transparent
                    && Rgba::from_colorref(key) == Rgba::opaque(r, g, b)
                {
                    pixel[3] = 0;
                }
                match recolor {
                    Recolor::None => {}
                    Recolor::ClearBlack if (r, g, b) == (0, 0, 0) => pixel[3] = 0,
                    Recolor::ClearWhite if (r, g, b) == (255, 255, 255) => pixel[3] = 0,
                    Recolor::Invert => {
                        pixel[0] = 255 - r;
                        pixel[1] = 255 - g;
                        pixel[2] = 255 - b;
                    }
                    Recolor::Stencil(color) => {
                        if (r, g, b) == (0, 0, 0) {
                            pixel.copy_from_slice(&[color.r, color.g, color.b, 255]);
                        } else {
                            pixel[3] = 0;
                        }
                    }
                    _ => {}
                }
            }
            Pixels::Rgba(out)
        }
    };
    let (width, height, encoded) = match &pixels {
        Pixels::Encoded { .. } => (dib.width, dib.height, true),
        Pixels::Rgba(_) => ((x1 - x0) as u32, (y1 - y0) as u32, false),
    };
    Some((
        Bitmap {
            width,
            height,
            pixels,
        },
        offset,
        encoded && !full,
    ))
}

/// `CREATEMONOBRUSH` and `CREATEDIBPATTERNBRUSHPT`: a brush tiling a bitmap.
pub(crate) fn pattern_brush<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    _kind: u32,
    body: usize,
) -> Option<()> {
    let handle = u32_at(bytes, body)? as usize;
    let usage = u32_at(bytes, body + 4)?;
    let bmi = slice(bytes, u32_at(bytes, body + 8)?, u32_at(bytes, body + 12)?)?;
    let bits = slice(bytes, u32_at(bytes, body + 16)?, u32_at(bytes, body + 20)?)?;
    let brush = match tile(player, bmi, bits, usage)? {
        Some(brush) => brush,
        None => {
            player.omit("pattern brushes that could not be decoded")?;
            Brush::solid(0x0080_8080, true)
        }
    };
    player.store(handle, GdiObject::Brush(brush));
    Some(())
}

pub(crate) fn wmf_pattern_brush<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    function: usize,
    body: usize,
) -> Option<()> {
    let brush = if function == 0x0142 {
        let usage = u32::from(u16_at(bytes, body + 2)?);
        let parts = bytes.get(body + 4..).and_then(|packed| {
            let offset = packed_bits_offset(packed)?;
            Some((packed.get(..offset)?, packed.get(offset..)?))
        });
        match parts {
            Some((bmi, bits)) => tile(player, bmi, bits, usage)?,
            None => None,
        }
    } else {
        None
    };
    let brush = match brush {
        Some(brush) => brush,
        None => {
            player.omit("pattern brushes that could not be decoded")?;
            Brush::solid(0x0080_8080, true)
        }
    };
    crate::wmf::store_object(player, GdiObject::Brush(brush));
    Some(())
}

/// A pattern brush's tile, `Some(None)` when it cannot be decoded. Its
/// declared pixels are charged before decoding, whether or not that succeeds.
fn tile<const FULL: bool>(
    player: &mut Player<FULL>,
    bmi: &[u8],
    bits: &[u8],
    usage: u32,
) -> Option<Option<Brush>> {
    let budget = player
        .limits
        .bitmap_pixels
        .saturating_sub(player.bitmap_pixels);
    let declared = declared_pixels(bmi);
    player.charge(0, declared)?;
    let Ok(dib) = decode(bmi, bits, usage, budget.min(1 << 16), |pixels| {
        player.charge(0, pixels.saturating_sub(declared)).is_some()
    }) else {
        if player.overflowed {
            return None;
        }
        return Some(None);
    };
    let DibPixels::Rgba(rgba) = dib.pixels else {
        return Some(None);
    };
    Some(Some(Brush {
        pattern: Some(Arc::new(Bitmap {
            width: dib.width,
            height: dib.height,
            pixels: Pixels::Rgba(rgba),
        })),
        ..Brush::solid(0, true)
    }))
}
