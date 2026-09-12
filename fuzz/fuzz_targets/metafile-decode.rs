#![no_main]

use libfuzzer_sys::fuzz_target;
use pptx_render::fuzzing::decode;

const EMF_EOF: u32 = 14;
const EMF_SIGNATURE: u32 = 0x464D_4520;
const WMF_PLACEABLE_KEY: u32 = 0x9AC6_CDD7;
const WMF_PLACEABLE_HEADER: usize = 22;
const WMF_HEADER: usize = 18;
const SLACK: u8 = 0xA5;

fn u16_at(bytes: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes(bytes.get(at..at + 2)?.try_into().ok()?))
}

fn u32_at(bytes: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
}

/// Grows every record up to EOF by `SLACK` bytes, rewriting its declared size to cover them.
fn pad_records(data: &[u8]) -> Option<Vec<u8>> {
    let emf = u32_at(data, 0)? == 1 && u32_at(data, 40)? == EMF_SIGNATURE;
    let (mut offset, unit, slack) = if emf {
        (0, 1, 4)
    } else if u32_at(data, 0)? == WMF_PLACEABLE_KEY {
        (WMF_PLACEABLE_HEADER + WMF_HEADER, 2, 2)
    } else {
        (WMF_HEADER, 2, 2)
    };
    let size_at = if emf { 4 } else { 0 };
    let mut padded = data[..offset].to_vec();
    loop {
        let size = u32_at(data, offset + size_at)? as usize * unit;
        let end = offset.checked_add(size).filter(|end| *end <= data.len())?;
        let eof = if emf {
            u32_at(data, offset)? == EMF_EOF
        } else {
            u16_at(data, offset + 4)? == 0
        };
        let record = &data[offset..end];
        padded.extend_from_slice(&record[..size_at]);
        padded.extend_from_slice(&(((size + slack) / unit) as u32).to_le_bytes());
        padded.extend_from_slice(&record[size_at + 4..]);
        padded.resize(padded.len() + slack, SLACK);
        offset = end;
        if eof {
            break;
        }
    }
    padded.extend_from_slice(&data[offset..]);
    Some(padded)
}

fuzz_target!(|data: &[u8]| {
    let Some(drawing) = decode(data) else {
        return;
    };
    let padded = pad_records(data).expect("a decoded metafile frames into records up to EOF");
    assert!(
        format!("{:?}", decode(&padded)) == format!("{:?}", Some(drawing)),
        "a record read past its declared size: slack after each record changed the drawing"
    );
});
