//! Bounds-checked little-endian reads.

pub(crate) fn u8_at(bytes: &[u8], offset: usize) -> Option<u8> {
    bytes.get(offset).copied()
}

pub(crate) fn u16_at(bytes: &[u8], offset: usize) -> Option<u16> {
    let slice = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([slice[0], slice[1]]))
}

pub(crate) fn u32_at(bytes: &[u8], offset: usize) -> Option<u32> {
    let slice = bytes.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

pub(crate) fn i32_at(bytes: &[u8], offset: usize) -> Option<i32> {
    u32_at(bytes, offset).map(|value| value as i32)
}

pub(crate) fn i16_at(bytes: &[u8], offset: usize) -> Option<i16> {
    u16_at(bytes, offset).map(|value| value as i16)
}

pub(crate) fn f32_at(bytes: &[u8], offset: usize) -> Option<f32> {
    u32_at(bytes, offset).map(f32::from_bits)
}

/// A finite `f32` widened, or `None`.
pub(crate) fn finite_at(bytes: &[u8], offset: usize) -> Option<f64> {
    f32_at(bytes, offset)
        .filter(|value| value.is_finite())
        .map(f64::from)
}

/// `count` items of `size` bytes from `offset`, entirely inside `bytes`.
pub(crate) fn span(bytes: &[u8], offset: usize, count: usize, size: usize) -> Option<&[u8]> {
    let end = offset.checked_add(count.checked_mul(size)?)?;
    bytes.get(offset..end)
}

/// Like [`span`], for data an EMF record addresses by offset from its start:
/// a non-empty span must also lie past the record's type and size.
pub(crate) fn record_span(bytes: &[u8], offset: usize, count: usize, size: usize) -> Option<&[u8]> {
    span(bytes, offset, count, size).filter(|span| span.is_empty() || offset >= 8)
}
