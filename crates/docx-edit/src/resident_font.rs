use std::fmt;

const MAX_FONT_BYTES: usize = 64 * 1024 * 1024;

#[derive(Debug, PartialEq, Eq)]
pub enum ResidentFontError {
    SizeLimit,
    InvalidHeader,
    TableBounds,
    OverlappingTables,
    Decode(String),
    InvalidFont(String),
}

impl fmt::Display for ResidentFontError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::SizeLimit => formatter.write_str("Invalid resident font size"),
            Self::InvalidHeader => formatter.write_str("Invalid web font header"),
            Self::TableBounds => formatter.write_str("Invalid web font table bounds"),
            Self::OverlappingTables => formatter.write_str("Web font tables overlap"),
            Self::Decode(error) | Self::InvalidFont(error) => formatter.write_str(error),
        }
    }
}

impl std::error::Error for ResidentFontError {}

fn bounded(size: usize) -> Result<usize, ResidentFontError> {
    if size > MAX_FONT_BYTES {
        Err(ResidentFontError::SizeLimit)
    } else {
        Ok(size)
    }
}

fn add_size(total: usize, size: usize) -> Result<usize, ResidentFontError> {
    bounded(
        total
            .checked_add(size)
            .ok_or(ResidentFontError::SizeLimit)?,
    )
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<usize, ResidentFontError> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or(ResidentFontError::InvalidHeader)?;
    Ok(u32::from_be_bytes(value.try_into().unwrap()) as usize)
}

fn block(
    bytes: &[u8],
    offset: usize,
    length: usize,
    start: usize,
) -> Result<usize, ResidentFontError> {
    bounded(offset)?;
    let end = add_size(offset, length)?;
    if offset < start || end > bytes.len() {
        return Err(ResidentFontError::TableBounds);
    }
    Ok(end)
}

fn preflight_woff1(bytes: &[u8], tables: usize) -> Result<(), ResidentFontError> {
    let start = 44 + tables * 20;
    if start > bytes.len() {
        return Err(ResidentFontError::InvalidHeader);
    }
    let mut sfnt_size = 12 + tables * 16;
    let (entries, _) = bytes[44..start].as_chunks::<20>();
    for entry in entries {
        let offset = u32_at(entry, 4)?;
        let compressed = bounded(u32_at(entry, 8)?)?;
        let original = bounded(u32_at(entry, 12)?)?;
        if compressed == 0 || compressed > original {
            return Err(ResidentFontError::TableBounds);
        }
        sfnt_size = add_size(sfnt_size, (original + 3) & !3)?;
        block(bytes, offset, compressed, start)?;
    }
    let mut ranges: Vec<_> = entries
        .iter()
        .map(|entry| {
            let offset = u32_at(entry, 4).unwrap();
            (offset, offset + u32_at(entry, 8).unwrap())
        })
        .collect();
    ranges.sort_unstable();
    if ranges.windows(2).any(|pair| pair[0].1 > pair[1].0) {
        return Err(ResidentFontError::OverlappingTables);
    }
    if sfnt_size != u32_at(bytes, 16)? {
        return Err(ResidentFontError::InvalidHeader);
    }
    Ok(())
}

fn byte_at(bytes: &[u8], cursor: &mut usize) -> Result<u8, ResidentFontError> {
    let byte = *bytes.get(*cursor).ok_or(ResidentFontError::InvalidHeader)?;
    *cursor += 1;
    Ok(byte)
}

fn base128(bytes: &[u8], cursor: &mut usize) -> Result<usize, ResidentFontError> {
    let mut value = 0_u32;
    for index in 0..5 {
        let byte = byte_at(bytes, cursor)?;
        if (index == 0 && byte == 0x80) || value & 0xfe000000 != 0 {
            return Err(ResidentFontError::InvalidHeader);
        }
        value = (value << 7) | u32::from(byte & 0x7f);
        if byte & 0x80 == 0 {
            return Ok(value as usize);
        }
    }
    Err(ResidentFontError::InvalidHeader)
}

fn uint255(bytes: &[u8], cursor: &mut usize) -> Result<usize, ResidentFontError> {
    match byte_at(bytes, cursor)? {
        253 => {
            Ok((usize::from(byte_at(bytes, cursor)?) << 8) | usize::from(byte_at(bytes, cursor)?))
        }
        254 => Ok(506 + usize::from(byte_at(bytes, cursor)?)),
        255 => Ok(253 + usize::from(byte_at(bytes, cursor)?)),
        byte => Ok(usize::from(byte)),
    }
}

fn preflight_woff2(bytes: &[u8], tables: usize) -> Result<(), ResidentFontError> {
    let mut cursor = 48;
    let mut sfnt_size = 12 + tables * 16;
    let mut data_size = 0;
    for _ in 0..tables {
        let flags = byte_at(bytes, &mut cursor)?;
        let tag = flags & 63;
        let glyf_or_loca = if tag == 63 {
            let value = u32_at(bytes, cursor)?;
            cursor += 4;
            value == u32::from_be_bytes(*b"glyf") as usize
                || value == u32::from_be_bytes(*b"loca") as usize
        } else {
            tag == 10 || tag == 11
        };
        let original = bounded(base128(bytes, &mut cursor)?)?;
        sfnt_size = add_size(sfnt_size, (original + 3) & !3)?;
        let transformed = if glyf_or_loca {
            flags >> 6 == 0
        } else {
            flags >> 6 != 0
        };
        let length = if transformed {
            bounded(base128(bytes, &mut cursor)?)?
        } else {
            original
        };
        data_size = add_size(data_size, length)?;
    }
    if &bytes[4..8] == b"ttcf" {
        let version = u32_at(bytes, cursor)?;
        cursor += 4;
        if version != 0x10000 && version != 0x20000 {
            return Err(ResidentFontError::InvalidHeader);
        }
        let fonts = uint255(bytes, &mut cursor)?;
        if fonts == 0 {
            return Err(ResidentFontError::InvalidHeader);
        }
        sfnt_size = add_size(
            sfnt_size - 12 - tables * 16,
            12 + 4 * fonts + if version == 0x20000 { 12 } else { 0 },
        )?;
        for _ in 0..fonts {
            let count = uint255(bytes, &mut cursor)?;
            u32_at(bytes, cursor)?;
            cursor += 4;
            sfnt_size = add_size(sfnt_size, 12 + count * 16)?;
            for _ in 0..count {
                if uint255(bytes, &mut cursor)? >= tables {
                    return Err(ResidentFontError::TableBounds);
                }
            }
        }
    }
    block(bytes, cursor, u32_at(bytes, 20)?, cursor)?;
    Ok(())
}

fn preflight_web_font(bytes: &[u8], woff2: bool) -> Result<(), ResidentFontError> {
    let header_size = if woff2 { 48 } else { 44 };
    if bytes.len() < header_size || u32_at(bytes, 8)? != bytes.len() {
        return Err(ResidentFontError::InvalidHeader);
    }
    bounded(u32_at(bytes, 16)?)?;
    let tables = usize::from(u16::from_be_bytes([bytes[12], bytes[13]]));
    if tables == 0 || (!woff2 && bytes[14..16] != [0, 0]) {
        return Err(ResidentFontError::InvalidHeader);
    }
    let metadata = if woff2 { 28 } else { 24 };
    bounded(u32_at(bytes, metadata + 8)?)?;
    for offset in [metadata, metadata + 12] {
        let start = u32_at(bytes, offset)?;
        let length = u32_at(bytes, offset + 4)?;
        if start != 0 || length != 0 {
            block(bytes, start, length, header_size)?;
        }
    }
    if woff2 {
        preflight_woff2(bytes, tables)
    } else {
        preflight_woff1(bytes, tables)
    }
}

pub fn decode_resident_font(bytes: &[u8]) -> Result<Vec<u8>, ResidentFontError> {
    bounded(bytes.len())?;
    if bytes.len() < 12 {
        return Err(ResidentFontError::InvalidHeader);
    }
    let signature = &bytes[..4];
    if signature == b"wOFF" || signature == b"wOF2" {
        preflight_web_font(bytes, signature == b"wOF2")?;
    }
    let decoded = match signature {
        b"wOFF" => wuff::decompress_woff1(bytes)
            .map_err(|error| ResidentFontError::Decode(error.to_string()))?,
        b"wOF2" => wuff::decompress_woff2(bytes)
            .map_err(|error| ResidentFontError::Decode(error.to_string()))?,
        _ => bytes.to_vec(),
    };
    bounded(decoded.len())?;
    ooxml_text::FontStore::new()
        .register(decoded.clone())
        .map_err(|error| ResidentFontError::InvalidFont(error.to_string()))?;
    Ok(decoded)
}
