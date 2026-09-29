//! Truncated, corrupted and hostile metafiles never panic, stay inside the
//! replay budgets, and refuse rather than draw from bytes past a record.

mod common;

use common::*;
use ooxml_metafile::{MAX_SVG_BYTES, decode, to_svg};

fn check(bytes: &[u8]) {
    if let Ok(svg) = to_svg(bytes) {
        assert!(svg.markup.len() <= MAX_SVG_BYTES);
        assert!(svg.width.is_finite() && svg.height.is_finite());
    }
    let _ = decode(bytes);
}

/// A deterministic xorshift, so every run corrupts the same bytes.
struct Noise(u64);

impl Noise {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }
}

#[test]
fn every_truncation_of_every_fixture_is_survivable() {
    for (_, bytes) in fixtures::all() {
        let stride = (bytes.len() / 4096).max(1);
        for cut in (0..bytes.len()).step_by(stride) {
            check(&bytes[..cut]);
        }
    }
}

#[test]
fn corrupted_fixtures_are_survivable() {
    let mut noise = Noise(0x9E37_79B9_7F4A_7C15);
    for (_, bytes) in fixtures::all() {
        for _ in 0..1500 {
            let mut corrupt = bytes.clone();
            for _ in 0..1 + noise.next() % 8 {
                let at = (noise.next() as usize) % corrupt.len();
                corrupt[at] = match noise.next() % 4 {
                    0 => 0,
                    1 => 0xff,
                    _ => noise.next() as u8,
                };
            }
            check(&corrupt);
        }
    }
}

/// Grows every EMF record by four bytes of slack, keeping its declared size honest.
fn pad_records(data: &[u8]) -> Vec<u8> {
    let mut padded = Vec::new();
    let mut offset = 0;
    loop {
        let kind = u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap());
        let size = u32::from_le_bytes(data[offset + 4..offset + 8].try_into().unwrap()) as usize;
        padded.extend_from_slice(&data[offset..offset + 4]);
        padded.extend_from_slice(&((size + 4) as u32).to_le_bytes());
        padded.extend_from_slice(&data[offset + 8..offset + size]);
        padded.extend_from_slice(&[0xA5; 4]);
        offset += size;
        if kind == 14 {
            break;
        }
    }
    let total = padded.len() as u32;
    padded[48..52].copy_from_slice(&total.to_le_bytes());
    padded
}

#[test]
fn slack_after_each_record_changes_nothing() {
    for (name, bytes) in fixtures::all() {
        if !name.ends_with(".emf") {
            continue;
        }
        let before = to_svg(&bytes).unwrap();
        let after = to_svg(&pad_records(&bytes)).unwrap();
        assert!(before.markup == after.markup, "{name} read past a record");
    }
}

#[test]
fn hostile_counts_sizes_and_budgets_refuse() {
    let huge_polygon = Emf::new(10, 10)
        .rec(86, &i32s(&[0, 0, 0, 0, i32::MAX]))
        .bytes();
    assert!(to_svg(&huge_polygon).is_err());
    let huge_text = Emf::new(10, 10)
        .recs(vec![{
            let (kind, mut body) = text_out(0, 0, "x", None, 0, [0, 0, -1, -1]);
            body[36..40].copy_from_slice(&u32::MAX.to_le_bytes());
            (kind, body)
        }])
        .bytes();
    assert!(to_svg(&huge_text).is_err());
    let dib = dib24(&[&[0]]);
    let (kind, mut blit) = stretch_dibits([0, 0, 10, 10], &dib, 0x00CC_0020);
    blit[72 + 4..72 + 8].copy_from_slice(&50_000i32.to_le_bytes());
    blit[72 + 8..72 + 12].copy_from_slice(&50_000i32.to_le_bytes());
    assert!(to_svg(&Emf::new(10, 10).rec(kind, &blit).bytes()).is_err());
    let mut past_end = Emf::new(10, 10).rec(43, &i32s(&[0, 0, 5, 5])).bytes();
    past_end[92..96].copy_from_slice(&0x7fff_fff0u32.to_le_bytes());
    assert!(to_svg(&past_end).is_err());
    let object = (0x4008, 0x8100, u32s(&[1 << 30, 0]));
    let continued = Emf::new(10, 10)
        .rec(70, &plus(&[plus_header(false), object]).1)
        .bytes();
    assert!(to_svg(&continued).is_err());
}

#[test]
fn nesting_stops_at_the_depth_limit() {
    let mut inner = Emf::new(10, 10)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                plus_fill_rects(0xff00_0000, &[[0.0, 0.0, 5.0, 5.0]]),
            ])
            .1,
        )
        .bytes();
    for _ in 0..6 {
        let mut image = u32s(&[0xDBC0_1002, 2, 3, inner.len() as u32]);
        image.extend(&inner);
        inner = Emf::new(10, 10)
            .rec(
                70,
                &plus(&[
                    plus_header(false),
                    (0x4008, 0x0500, image),
                    plus_draw_image(0, [0.0, 0.0, 10.0, 10.0], [0.0, 0.0, 10.0, 10.0]),
                ])
                .1,
            )
            .bytes();
    }
    let svg = to_svg(&inner).unwrap();
    assert_eq!(svg.omissions[0].what, "metafiles nested too deeply");
}
