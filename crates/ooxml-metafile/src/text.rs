//! Text records, laid out as positioned runs.

use std::sync::Arc;

use crate::drawing::{Font, Op, Paint, Rgba, Text, TextAnchor};
use crate::player::{Combine, LogFont, Player, apply, axis_aligned_rect};
use crate::read::{i16_at, i32_at, u16_at, u32_at};

const ETO_OPAQUE: u32 = 0x0002;
const ETO_CLIPPED: u32 = 0x0004;
const ETO_GLYPH_INDEX: u32 = 0x0010;
const ETO_NO_RECT: u32 = 0x0100;
const ETO_SMALL_CHARS: u32 = 0x0200;
const ETO_PDY: u32 = 0x2000;

const TA_UPDATECP: u32 = 0x0001;
const SYMBOL_CHARSET: u8 = 2;

/// One text record's characters, decoded, with their advances in logical units.
pub(crate) struct Run {
    pub reference: (f64, f64),
    pub chars: Vec<char>,
    pub advances: Option<Vec<(f64, f64)>>,
    pub options: u32,
    pub rect: Option<(f64, f64, f64, f64)>,
}

/// Ascent and descent as fractions of the em, as GDI's text metrics report them.
pub(crate) fn font_metrics(face: &str) -> (f64, f64) {
    match face.to_ascii_lowercase().as_str() {
        "arial" | "helvetica" | "liberation sans" | "arimo" => (0.905, 0.212),
        "times new roman" | "times" | "liberation serif" | "tinos" => (0.891, 0.216),
        "calibri" | "carlito" => (0.952, 0.269),
        "cambria" | "caladea" => (0.950, 0.222),
        "courier new" | "courier" | "liberation mono" | "cousine" => (0.833, 0.300),
        "verdana" => (1.005, 0.210),
        "tahoma" => (1.000, 0.207),
        "segoe ui" => (1.079, 0.251),
        "symbol" => (1.005, 0.220),
        _ => (0.900, 0.220),
    }
}

fn default_font() -> LogFont {
    LogFont {
        height: 16,
        escapement: 0,
        weight: 700,
        italic: false,
        underline: false,
        strike: false,
        charset: 0,
        face: "Arial".to_owned(),
    }
}

/// UTF-16 code units to characters, summing the advances of a surrogate pair.
pub(crate) fn decode_wide(
    units: &[u16],
    advances: Option<Vec<(f64, f64)>>,
) -> (Vec<char>, Option<Vec<(f64, f64)>>) {
    let mut chars = Vec::with_capacity(units.len());
    let mut merged = advances.as_ref().map(|_| Vec::with_capacity(units.len()));
    let mut index = 0;
    while index < units.len() {
        let unit = units[index];
        let (c, width) = if (0xD800..0xDC00).contains(&unit)
            && let Some(low) = units
                .get(index + 1)
                .filter(|low| (0xDC00..0xE000).contains(*low))
        {
            let code = 0x10000 + ((u32::from(unit) - 0xD800) << 10) + (u32::from(*low) - 0xDC00);
            (char::from_u32(code).unwrap_or('\u{FFFD}'), 2)
        } else {
            (char::from_u32(u32::from(unit)).unwrap_or('\u{FFFD}'), 1)
        };
        chars.push(c);
        if let (Some(merged), Some(advances)) = (merged.as_mut(), advances.as_ref()) {
            let sum = advances[index..(index + width).min(advances.len())]
                .iter()
                .fold((0.0, 0.0), |sum, advance| {
                    (sum.0 + advance.0, sum.1 + advance.1)
                });
            merged.push(sum);
        }
        index += width;
    }
    (chars, merged)
}

/// Single-byte text in a font's character set, or `None` for a multi-byte set.
pub(crate) fn decode_narrow(bytes: &[u8], charset: u8) -> Option<Vec<char>> {
    if matches!(charset, 128 | 129 | 130 | 134 | 136) {
        return None;
    }
    Some(
        bytes
            .iter()
            .map(|byte| match charset {
                SYMBOL_CHARSET => char::from_u32(0xF000 + u32::from(*byte)).unwrap_or(' '),
                255 => cp437(*byte),
                _ => cp1252(*byte),
            })
            .collect(),
    )
}

fn cp1252(byte: u8) -> char {
    const HIGH: [u16; 32] = [
        0x20AC, 0xFFFD, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160,
        0x2039, 0x0152, 0xFFFD, 0x017D, 0xFFFD, 0xFFFD, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022,
        0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0xFFFD, 0x017E, 0x0178,
    ];
    match byte {
        0x80..=0x9F => char::from_u32(u32::from(HIGH[usize::from(byte - 0x80)])).unwrap_or(' '),
        _ => char::from(byte),
    }
}

fn cp437(byte: u8) -> char {
    if byte < 0x80 {
        return char::from(byte);
    }
    const HIGH: &str = "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\u{A0}";
    HIGH.chars().nth(usize::from(byte - 0x80)).unwrap_or(' ')
}

/// The Unicode character Adobe's Symbol encoding places at `code`.
fn symbol(code: u32) -> Option<char> {
    const LOW: &str = " !∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼";
    const HIGH: &str = "€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…⏐⎯↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪ 〉∫⌠⎮⌡⎞⎟⎠⎤⎥⎦⎫⎬⎭";
    match code {
        0x20..=0x7E => LOW.chars().nth((code - 0x20) as usize),
        0xA0..=0xFE => HIGH.chars().nth((code - 0xA0) as usize),
        _ => None,
    }
}

/// Characters of a symbol font, remapped to the Unicode they depict where the
/// font is Symbol itself, or to the private-use code points symbol fonts carry.
fn map_symbols(chars: &mut [char], font: &LogFont) {
    let is_symbol = font.face.eq_ignore_ascii_case("symbol");
    if !is_symbol && font.charset != SYMBOL_CHARSET {
        return;
    }
    for c in chars {
        let code = *c as u32;
        let byte = if (0xF020..=0xF0FF).contains(&code) {
            code - 0xF000
        } else if (0x20..=0xFF).contains(&code) {
            code
        } else {
            continue;
        };
        *c = if is_symbol {
            symbol(byte).unwrap_or(*c)
        } else {
            char::from_u32(0xF000 + byte).unwrap_or(*c)
        };
    }
}

fn read_advances(
    bytes: &[u8],
    offset: usize,
    count: usize,
    pairs: bool,
    wide: bool,
) -> Option<Vec<(f64, f64)>> {
    if offset == 0 {
        return None;
    }
    let stride = if wide { 4 } else { 2 };
    let per = if pairs { 2 } else { 1 };
    crate::read::span(bytes, offset, count.checked_mul(per)?, stride)?;
    let value = |at: usize| {
        if wide {
            i32_at(bytes, at).map(f64::from)
        } else {
            i16_at(bytes, at).map(f64::from)
        }
    };
    let mut advances = Vec::with_capacity(count);
    for index in 0..count {
        let at = offset + index * per * stride;
        let dx = value(at)?;
        let dy = if pairs { value(at + stride)? } else { 0.0 };
        advances.push((dx, dy));
    }
    Some(advances)
}

pub(crate) fn emf_text<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    kind: u32,
    body: usize,
) -> Option<()> {
    player.flush_pending();
    match kind {
        83 | 84 => {
            let run = emr_text(player, bytes, body + 28, kind == 84)?;
            draw(player, run)
        }
        96 | 97 => {
            let count = u32_at(bytes, body + 28)? as usize;
            if count > player.limits.text_chars {
                return player.refuse("a text record holds more strings than the limit");
            }
            for index in 0..count {
                let run = emr_text(player, bytes, body + 32 + index * 40, kind == 97)?;
                draw(player, run)?;
            }
            Some(())
        }
        _ => {
            let (x, y) = (i32_at(bytes, body)?, i32_at(bytes, body + 4)?);
            let count = u32_at(bytes, body + 8)? as usize;
            charge_text(player, count)?;
            let options = u32_at(bytes, body + 12)?;
            let (rect, text) = if options & ETO_NO_RECT == 0 {
                (Some(rect_at(bytes, body + 28)?), body + 44)
            } else {
                (None, body + 28)
            };
            let font = player.dc.font.clone();
            let chars = if options & ETO_SMALL_CHARS != 0 {
                let raw = crate::read::span(bytes, text, count, 1)?;
                match decode_narrow(raw, font.as_ref().map_or(0, |font| font.charset)) {
                    Some(chars) => chars,
                    None => return player.omit("text in multi-byte character sets"),
                }
            } else {
                decode_wide(&wide_units(bytes, text, count)?, None).0
            };
            let run = Run {
                reference: (f64::from(x), f64::from(y)),
                chars,
                advances: None,
                options,
                rect,
            };
            draw(player, run)
        }
    }
}

fn wide_units(bytes: &[u8], offset: usize, count: usize) -> Option<Vec<u16>> {
    crate::read::record_span(bytes, offset, count, 2)?;
    (0..count)
        .map(|index| u16_at(bytes, offset + index * 2))
        .collect()
}

fn rect_at(bytes: &[u8], at: usize) -> Option<(f64, f64, f64, f64)> {
    Some((
        f64::from(i32_at(bytes, at)?),
        f64::from(i32_at(bytes, at + 4)?),
        f64::from(i32_at(bytes, at + 8)?),
        f64::from(i32_at(bytes, at + 12)?),
    ))
}

/// An `EmrText` object at `at`, whose string and advances sit at offsets
/// from the start of the record.
fn emr_text<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    at: usize,
    wide: bool,
) -> Option<Run> {
    let reference = (
        f64::from(i32_at(bytes, at)?),
        f64::from(i32_at(bytes, at + 4)?),
    );
    let count = u32_at(bytes, at + 8)? as usize;
    let string = u32_at(bytes, at + 12)? as usize;
    let options = u32_at(bytes, at + 16)?;
    let (rect, dx) = if options & ETO_NO_RECT == 0 {
        (Some(rect_at(bytes, at + 20)?), u32_at(bytes, at + 36)?)
    } else {
        (None, u32_at(bytes, at + 20)?)
    };
    charge_text(player, count)?;
    let pairs = options & ETO_PDY != 0;
    if dx != 0 && count > 0 {
        let length = count << usize::from(pairs);
        crate::read::record_span(bytes, dx as usize, length, 4)?;
        player.spend(length as u64, 0)?;
    }
    let advances = read_advances(bytes, dx as usize, count, pairs, true);
    let (chars, advances) = if wide {
        decode_wide(&wide_units(bytes, string, count)?, advances)
    } else {
        let raw = crate::read::record_span(bytes, string, count, 1)?;
        let charset = player.dc.font.as_ref().map_or(0, |font| font.charset);
        match decode_narrow(raw, charset) {
            Some(chars) => (chars, advances),
            None => {
                player.omit("text in multi-byte character sets")?;
                (Vec::new(), None)
            }
        }
    };
    Some(Run {
        reference,
        chars,
        advances,
        options,
        rect,
    })
}

pub(crate) fn wmf_text<const FULL: bool>(
    player: &mut Player<FULL>,
    bytes: &[u8],
    function: usize,
    body: usize,
) -> Option<()> {
    player.flush_pending();
    let charset = player.dc.font.as_ref().map_or(0, |font| font.charset);
    let run = if function == 0x0521 {
        let length = u16_at(bytes, body)? as usize;
        charge_text(player, length)?;
        let raw = crate::read::span(bytes, body + 2, length, 1)?;
        let at = body + 2 + length.div_ceil(2) * 2;
        let (y, x) = (i16_at(bytes, at)?, i16_at(bytes, at + 2)?);
        let Some(chars) = decode_narrow(raw, charset) else {
            return player.omit("text in multi-byte character sets");
        };
        Run {
            reference: (f64::from(x), f64::from(y)),
            chars,
            advances: None,
            options: 0,
            rect: None,
        }
    } else {
        let (y, x) = (i16_at(bytes, body)?, i16_at(bytes, body + 2)?);
        let length = i16_at(bytes, body + 4)?.max(0) as usize;
        charge_text(player, length)?;
        let options = u32::from(u16_at(bytes, body + 6)?);
        let (rect, text) = if options & (ETO_OPAQUE | ETO_CLIPPED) != 0 {
            let side = |at| i16_at(bytes, body + at).map(f64::from);
            (Some((side(8)?, side(10)?, side(12)?, side(14)?)), body + 16)
        } else {
            (None, body + 8)
        };
        let raw = crate::read::span(bytes, text, length, 1)?;
        let dx = text + length.div_ceil(2) * 2;
        if crate::read::span(bytes, dx, length, 2).is_some() {
            player.spend(length as u64, 0)?;
        }
        let advances = read_advances(bytes, dx, length, false, false);
        let Some(chars) = decode_narrow(raw, charset) else {
            return player.omit("text in multi-byte character sets");
        };
        Run {
            reference: (f64::from(x), f64::from(y)),
            chars,
            advances,
            options,
            rect,
        }
    };
    draw(player, run)
}

/// Spends `count` characters of the text budget before they are decoded.
fn charge_text<const FULL: bool>(player: &mut Player<FULL>, count: usize) -> Option<()> {
    player.text_chars = player.text_chars.saturating_add(count);
    if player.text_chars > player.limits.text_chars {
        return player.refuse("the metafile holds more text than the limit");
    }
    player.spend(count as u64, 0)
}

/// Lays one run out in the current font and pushes it.
fn draw<const FULL: bool>(player: &mut Player<FULL>, mut run: Run) -> Option<()> {
    if run.options & ETO_GLYPH_INDEX != 0 {
        return player.omit("text given as glyph indexes");
    }
    let font = player
        .dc
        .font
        .clone()
        .unwrap_or_else(|| Arc::new(default_font()));
    map_symbols(&mut run.chars, &font);
    let (ascent, descent) = font_metrics(&font.face);
    let em = match font.height {
        height if height < 0 => -f64::from(height),
        height if height > 0 => f64::from(height) / (ascent + descent),
        _ => 12.0,
    };
    let saved_clip = player.dc.clip.clone();
    let saved_clip_rect = player.dc.clip_rect;
    if let Some(rect) = run.rect {
        let path = player.logical_rect_path(rect);
        if run.options & ETO_OPAQUE != 0 {
            player.path = path.clone();
            player.emit(
                Some(Paint::Solid(Rgba::from_colorref(player.dc.bk_color))),
                None,
            );
        }
        if run.options & ETO_CLIPPED != 0 {
            match axis_aligned_rect(&path) {
                Some(rect) => player.intersect_rect(rect)?,
                None => player.combine_clip(path, false, Combine::And)?,
            }
        }
    }
    let align = player.dc.text_align;
    if align & TA_UPDATECP != 0 {
        run.reference = player.current;
    }
    let total = run.advances.as_ref().map(|advances| {
        advances
            .iter()
            .fold((0.0, 0.0), |sum, a| (sum.0 + a.0, sum.1 + a.1))
    });
    let (x_shift, anchor) = match (align & 6, total) {
        (2, Some(total)) => (-total.0, TextAnchor::Start),
        (6, Some(total)) => (-total.0 / 2.0, TextAnchor::Start),
        (2, None) => (0.0, TextAnchor::End),
        (6, None) => (0.0, TextAnchor::Middle),
        _ => (0.0, TextAnchor::Start),
    };
    let y_shift = match align & 24 {
        24 => 0.0,
        8 => -descent * em,
        _ => ascent * em,
    };
    let m = player.logical_to_output();
    let angle = f64::from(font.escapement).to_radians() / 10.0;
    let (sin, cos) = angle.sin_cos();
    let down = m[0] * m[3] - m[1] * m[2] >= 0.0;
    let (u_l, v_l) = if down {
        ((cos, -sin), (sin, cos))
    } else {
        ((cos, sin), (sin, -cos))
    };
    let linear = |(x, y): (f64, f64)| (x * m[0] + y * m[2], x * m[1] + y * m[3]);
    let (u, v) = (linear(u_l), linear(v_l));
    let origin = apply(m, run.reference);
    let origin = (
        origin.0 + u.0 * x_shift + v.0 * y_shift,
        origin.1 + u.1 * x_shift + v.1 * y_shift,
    );
    let transform = [u.0, u.1, v.0, v.1, origin.0, origin.1];
    let positions = run.advances.as_ref().map(|advances| {
        let mut at = (0.0, 0.0);
        advances
            .iter()
            .map(|advance| {
                let here = at;
                at = (at.0 + advance.0, at.1 - advance.1);
                here
            })
            .collect::<Vec<_>>()
    });
    if player.dc.bk_opaque
        && run.options & ETO_OPAQUE == 0
        && let Some(total) = total
        && total.0 != 0.0
    {
        let corners = [
            (0.0, -ascent * em),
            (total.0, -ascent * em),
            (total.0, descent * em),
            (0.0, descent * em),
        ];
        let mut path = Vec::with_capacity(5);
        for (index, corner) in corners.into_iter().enumerate() {
            let (x, y) = apply(transform, corner);
            path.push(if index == 0 {
                crate::drawing::PathCommand::Move { x, y }
            } else {
                crate::drawing::PathCommand::Line { x, y }
            });
        }
        path.push(crate::drawing::PathCommand::Close);
        player.path = path;
        player.emit(
            Some(Paint::Solid(Rgba::from_colorref(player.dc.bk_color))),
            None,
        );
    }
    if align & TA_UPDATECP != 0 {
        let advance = match total {
            Some(total) => total.0,
            None if !run.chars.is_empty() => {
                player.omit("text advanced by estimated widths")?;
                let per_char = if font.face.to_ascii_lowercase().contains("courier") {
                    0.6
                } else {
                    0.5
                };
                per_char * em * run.chars.len() as f64
            }
            None => 0.0,
        };
        player.current = (
            run.reference.0 + advance * cos,
            run.reference.1 - advance * sin * if down { 1.0 } else { -1.0 },
        );
    }
    let text: String = run.chars.iter().collect();
    if !text.is_empty() {
        player.push_op(Op::Text(Text {
            transform,
            text,
            positions,
            anchor,
            font: Font {
                family: font.face.clone(),
                size: em,
                weight: match font.weight {
                    weight if weight <= 0 => 400,
                    weight => weight.min(1000) as u16,
                },
                italic: font.italic,
                underline: font.underline,
                strike: font.strike,
            },
            fill: Paint::Solid(Rgba::from_colorref(player.dc.text_color)),
            clip: player.dc.clip.clone(),
        }));
    }
    player.dc.clip = saved_clip;
    player.dc.clip_rect = saved_clip_rect;
    Some(())
}
