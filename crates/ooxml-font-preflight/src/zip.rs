use std::borrow::Cow;

use crate::PreflightError;

const MAX_PART: usize = 32 * 1024 * 1024;
pub(crate) const MAX_READS: usize = 512;

struct Entry<'a> {
    name: &'a [u8],
    offset: usize,
    compressed: usize,
    size: usize,
    method: u16,
    flags: u16,
}

pub(crate) struct Zip<'a> {
    bytes: &'a [u8],
    entries: Vec<Entry<'a>>,
    directory: usize,
    remaining: usize,
    reads: usize,
}

fn slice(bytes: &[u8], offset: usize, len: usize) -> Result<&[u8], PreflightError> {
    bytes
        .get(offset..offset.checked_add(len).ok_or(PreflightError::Zip)?)
        .ok_or(PreflightError::Zip)
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, PreflightError> {
    let value = slice(bytes, offset, 2)?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, PreflightError> {
    let value = slice(bytes, offset, 4)?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

impl<'a> Zip<'a> {
    pub(crate) fn new(bytes: &'a [u8]) -> Result<Self, PreflightError> {
        let end = bytes.len().checked_sub(22).ok_or(PreflightError::Zip)?;
        let eocd = (end.saturating_sub(65535)..=end)
            .rev()
            .find(|&at| {
                bytes.get(at..at + 4) == Some(b"PK\x05\x06")
                    && u16_at(bytes, at + 20)
                        .is_ok_and(|len| at + 22 + usize::from(len) == bytes.len())
            })
            .ok_or(PreflightError::Zip)?;
        let count = u16_at(bytes, eocd + 10)?;
        if u16_at(bytes, eocd + 4)? != 0
            || u16_at(bytes, eocd + 6)? != 0
            || u16_at(bytes, eocd + 8)? != count
            || count == u16::MAX
        {
            return Err(PreflightError::Zip);
        }
        let directory = u32_at(bytes, eocd + 16)? as usize;
        let directory_size = u32_at(bytes, eocd + 12)? as usize;
        let central = slice(bytes, directory, directory_size)?;
        if directory.checked_add(directory_size) != Some(eocd) {
            return Err(PreflightError::Zip);
        }
        let mut entries = Vec::new();
        let mut at = 0;
        for _ in 0..count {
            slice(central, at, 46)?;
            if slice(central, at, 4)? != b"PK\x01\x02" || u16_at(central, at + 34)? != 0 {
                return Err(PreflightError::Zip);
            }
            let name_len = usize::from(u16_at(central, at + 28)?);
            let extra_len = usize::from(u16_at(central, at + 30)?);
            let comment_len = usize::from(u16_at(central, at + 32)?);
            let name = slice(central, at + 46, name_len)?;
            entries.push(Entry {
                name,
                offset: u32_at(central, at + 42)? as usize,
                compressed: u32_at(central, at + 20)? as usize,
                size: u32_at(central, at + 24)? as usize,
                method: u16_at(central, at + 10)?,
                flags: u16_at(central, at + 8)?,
            });
            let len = 46 + name_len + extra_len + comment_len;
            slice(central, at, len)?;
            at += len;
        }
        if at != central.len() {
            return Err(PreflightError::Zip);
        }
        entries.sort_unstable_by(|a, b| a.name.cmp(b.name));
        if entries.windows(2).any(|pair| pair[0].name == pair[1].name) {
            return Err(PreflightError::Zip);
        }
        Ok(Self {
            bytes,
            entries,
            directory,
            remaining: 64 * 1024 * 1024,
            reads: MAX_READS,
        })
    }

    pub(crate) fn read(&mut self, name: &str) -> Result<Option<Cow<'a, [u8]>>, PreflightError> {
        self.reads = self.reads.checked_sub(1).ok_or(PreflightError::Limit)?;
        let Ok(index) = self
            .entries
            .binary_search_by(|entry| entry.name.cmp(name.as_bytes()))
        else {
            return Ok(None);
        };
        let entry = &self.entries[index];
        if entry.flags & 0x41 != 0 {
            return Err(PreflightError::Zip);
        }
        if entry.size > MAX_PART || entry.size > self.remaining {
            return Err(PreflightError::Limit);
        }
        let local = slice(self.bytes, 0, self.directory)?;
        let at = entry.offset;
        slice(local, at, 30)?;
        if slice(local, at, 4)? != b"PK\x03\x04"
            || u16_at(local, at + 6)? != entry.flags
            || u16_at(local, at + 8)? != entry.method
        {
            return Err(PreflightError::Zip);
        }
        let name_len = usize::from(u16_at(local, at + 26)?);
        let extra_len = usize::from(u16_at(local, at + 28)?);
        if slice(local, at + 30, name_len)? != entry.name {
            return Err(PreflightError::Zip);
        }
        let start = at
            .checked_add(30 + name_len + extra_len)
            .ok_or(PreflightError::Zip)?;
        let compressed = slice(local, start, entry.compressed)?;
        let result = match entry.method {
            0 => Cow::Borrowed(compressed),
            8 => Cow::Owned(
                miniz_oxide::inflate::decompress_to_vec_with_limit(compressed, entry.size)
                    .map_err(|_| PreflightError::Zip)?,
            ),
            _ => return Err(PreflightError::Zip),
        };
        if result.len() != entry.size {
            return Err(PreflightError::Zip);
        }
        self.remaining -= entry.size;
        Ok(Some(result))
    }
}

#[cfg(test)]
pub(crate) fn crc32(bytes: &[u8]) -> u32 {
    const TABLE: [u32; 256] = {
        let mut table = [0; 256];
        let mut index = 0;
        while index < 256 {
            let mut crc = index as u32;
            let mut bit = 0;
            while bit < 8 {
                crc = (crc >> 1) ^ (0xedb88320 & 0u32.wrapping_sub(crc & 1));
                bit += 1;
            }
            table[index] = crc;
            index += 1;
        }
        table
    };
    let mut crc = !0u32;
    for &byte in bytes {
        crc = (crc >> 8) ^ TABLE[usize::from((crc as u8) ^ byte)];
    }
    !crc
}
