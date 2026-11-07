//! Truncated, corrupted and hostile metafiles never panic, stay inside the
//! replay budgets, and refuse rather than draw from bytes past a record.

mod common;

use common::*;
use std::sync::Arc;

use ooxml_metafile::drawing::{Op, Pixels};
use ooxml_metafile::{
    MAX_SVG_BYTES, ReplayBudget, decode, replay, replay_with_budget, to_svg, to_svg_with_budget,
};

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
fn data_a_record_addresses_past_its_bounds_refuses() {
    let text = |edit: fn(&mut Vec<u8>)| {
        let (kind, mut body) = text_out(0, 0, "ab", Some(&[5, 5]), 0, [0, 0, -1, -1]);
        edit(&mut body);
        Emf::new(10, 10).rec(kind, &body).bytes()
    };
    let (bmi, bits) = dib24(&[&[0]]);
    let mut pattern = u32s(&[1, 0, 32, bmi.len() as u32]);
    pattern.extend(u32s(&[32 + bmi.len() as u32, bits.len() as u32 + 4]));
    pattern.extend(bmi.iter().chain(&bits));
    let (kind, mut comment) = plus(&[plus_header(false), plus_eof()]);
    comment[0] += 4;
    let (_, mut short_font) = font(1, 16, 400, [0; 3], 0, 0, "Arial");
    short_font.truncate(36);
    let (_, mut short_brush) = brush(1, 2, 0, 5);
    short_brush.truncate(12);
    for (name, bytes) in [
        ("advances", text(|body| body.truncate(body.len() - 4))),
        (
            "string",
            text(|body| body[40..44].copy_from_slice(&4u32.to_le_bytes())),
        ),
        ("pattern", Emf::new(10, 10).rec(94, &pattern).bytes()),
        ("comment", Emf::new(10, 10).rec(kind, &comment).bytes()),
        ("font", Emf::new(10, 10).rec(82, &short_font).bytes()),
        ("brush", Emf::new(10, 10).rec(39, &short_brush).bytes()),
    ] {
        assert!(to_svg(&bytes).is_err(), "{name}");
    }
    let (kind, mut empty) = text_out(0, 0, "", None, 0x2, [0, 0, 5, 5]);
    empty[64..68].copy_from_slice(&76u32.to_le_bytes());
    assert!(to_svg(&Emf::new(10, 10).rec(kind, &empty).bytes()).is_ok());
}

#[test]
fn pattern_brushes_that_fail_to_decode_spend_the_pixel_budget() {
    let mut record = u32s(&[1, 0, 32, 40, 72, 4, 40, 1 << 16, 1]);
    record.extend(u16s(&[1, 24]));
    record.extend(u32s(&[0, 0, 0, 0, 0, 0, 0]));
    let bytes = (0..300)
        .fold(Emf::new(10, 10), |emf, _| emf.rec(94, &record))
        .bytes();
    assert!(to_svg(&bytes).is_err());
}

#[test]
fn region_geometry_filled_away_spends_the_budget() {
    let points: Vec<(f32, f32)> = (0..10_000).map(|x| (x as f32, 0.0)).collect();
    let path = plus_path(1, &points).2;
    let mut region = u32s(&[
        0xDBC0_1002,
        2,
        4,
        0x1000_0003,
        0x1000_0001,
        path.len() as u32,
    ]);
    region.extend(path);
    let mut records = vec![plus_header(false), (0x4008, 0x0402, region)];
    records.extend((0..500).map(|_| (0x4013, 0x8002, u32s(&[0xff00_0000]))));
    let bytes = Emf::new(10, 10).rec(70, &plus(&records).1).bytes();
    assert!(to_svg(&bytes).is_err());
}

fn intersect_region_tree(depth: u32, leaf: &[u8], out: &mut Vec<u8>) {
    if depth == 0 {
        out.extend_from_slice(leaf);
        return;
    }
    out.extend(u32s(&[1]));
    intersect_region_tree(depth - 1, leaf, out);
    intersect_region_tree(depth - 1, leaf, out);
}

fn region_fills(region: &[u8], fills: usize) -> Vec<u8> {
    let mut records = vec![plus_header(false), (0x4008, 0x0400, region.to_vec())];
    records.extend((0..fills).map(|_| (0x4013, 0x8000, u32s(&[0xff00_0000]))));
    records.push(plus_eof());
    Emf::new(10, 10).rec(70, &plus(&records).1).bytes()
}

#[test]
fn zero_point_region_paths_are_empty() {
    let path = plus_path(0, &[]).2;
    let mut region = u32s(&[0xDBC0_1002, 0, 0x1000_0001, path.len() as u32]);
    region.extend(path);
    let empty = u32s(&[0xDBC0_1002, 0, 0x1000_0002]);
    assert_eq!(
        to_svg(&region_fills(&region, 1)).unwrap(),
        to_svg(&region_fills(&empty, 1)).unwrap()
    );
}

#[test]
fn repeated_fills_of_a_wide_zero_point_region_tree_spend_the_budget() {
    let keep = plus_path(0, &[(0.0, 0.0)]).2;
    let path = plus_path(0, &[]).2;
    let mut leaf = u32s(&[4, 0x1000_0001, keep.len() as u32]);
    leaf.extend(keep);
    leaf.extend(u32s(&[0x1000_0001, path.len() as u32]));
    leaf.extend(path);
    let mut region = u32s(&[0xDBC0_1002, (1 << 14) - 2]);
    intersect_region_tree(12, &leaf, &mut region);
    assert!(to_svg(&region_fills(&region, 3)).is_ok());
    assert_eq!(
        to_svg(&region_fills(&region, 100)).unwrap_err().0,
        "the metafile draws more than the replay limits"
    );
}

#[test]
fn empty_region_rectangles_filled_away_spend_the_budget() {
    let path = plus_path(0, &[(0.0, 0.0)]).2;
    let mut leaf = u32s(&[4, 0x1000_0001, path.len() as u32]);
    leaf.extend(path);
    leaf.extend(u32s(&[0x1000_0002]));
    let mut region = u32s(&[0xDBC0_1002, (1 << 12) - 2]);
    intersect_region_tree(10, &leaf, &mut region);
    assert!(to_svg(&region_fills(&region, 3)).is_ok());
    assert_eq!(
        to_svg(&region_fills(&region, 400)).unwrap_err().0,
        "the metafile draws more than the replay limits"
    );
}

fn repeated_gdi_text(length: usize, count: u32) -> Vec<u8> {
    let units = vec![b'a' as u16; length];
    let string_at = 8 + 32 + count * 40;
    let mut body = i32s(&[0, 0, -1, -1]);
    body.extend(u32s(&[1]));
    body.extend(f32s(&[1.0, 1.0]));
    body.extend(u32s(&[count]));
    for _ in 0..count {
        body.extend(i32s(&[0, 0]));
        body.extend(u32s(&[units.len() as u32, string_at, 0x10]));
        body.extend(i32s(&[0, 0, -1, -1]));
        body.extend(u32s(&[0]));
    }
    body.extend(u16s(&units));
    Emf::new(10, 10).rec(97, &body).bytes()
}

#[test]
fn a_string_drawn_many_times_spends_the_text_budget_once_per_draw() {
    assert!(to_svg(&repeated_gdi_text(100_000, 20)).is_err());
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

#[test]
fn a_dib_without_width_or_height_is_survivable() {
    for (width, height) in [(0u16, 4u16), (4, 0), (0, 0)] {
        let mut bmi = u32s(&[12]);
        bmi.extend(u16s(&[width, height, 1, 24]));
        let dib = (bmi, vec![0; 64]);
        check(
            &Emf::new(10, 10)
                .recs(vec![stretch_dibits([0, 0, 10, 10], &dib, 0x00CC_0020)])
                .bytes(),
        );
    }
}

#[test]
fn an_emf_plus_record_claiming_4_gib_refuses() {
    let mut comment = plus(&[plus_header(false)]).1;
    comment.extend(u16s(&[0x400A, 0x8000]));
    comment.extend(u32s(&[u32::MAX - 3, u32::MAX - 15]));
    let size = (comment.len() - 4) as u32;
    comment[0..4].copy_from_slice(&size.to_le_bytes());
    assert!(to_svg(&Emf::new(10, 10).rec(70, &comment).bytes()).is_err());
}

#[test]
fn a_deep_emf_plus_region_tree_stays_off_the_stack() {
    let depth = 200_000;
    let mut region = u32s(&[0xDBC0_1002, 2 * depth]);
    region.extend(u32s(&vec![1; depth as usize]));
    region.extend(u32s(&vec![0x1000_0003; depth as usize + 1]));
    let bytes = Emf::new(10, 10)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                (0x4008, 0x0400, region),
                (0x4034, 0, Vec::new()),
                plus_fill_rects(0xff00_0000, &[[0.0, 0.0, 5.0, 5.0]]),
            ])
            .1,
        )
        .bytes();
    check(&bytes);
}

fn nested_image(inner: &[u8]) -> (u16, u16, Vec<u8>) {
    let mut image = u32s(&[0xDBC0_1002, 2, 3, inner.len() as u32]);
    image.extend(inner);
    (0x4008, 0x0500, image)
}

fn repeated_nested_images(inner: &[u8], count: usize) -> Vec<u8> {
    let mut records = vec![plus_header(true)];
    records.extend((0..count).map(|_| nested_image(inner)));
    records.push(plus_draw_image(
        0,
        [0.0, 0.0, 10.0, 10.0],
        [0.0, 0.0, 10.0, 10.0],
    ));
    records.push(plus_eof());
    Emf::new(10, 10)
        .rec(70, &plus(&records).1)
        .recs(vec![rect(0, 0, 5, 5)])
        .bytes()
}

#[test]
fn repeated_cropped_nested_images_spend_decoded_pixels() {
    let inner = cropped_rle_emf(1024);
    let drawing = replay(&inner).unwrap();
    let ooxml_metafile::drawing::Op::Image(image) = &drawing.ops[0] else {
        panic!("the child must retain its cropped bitmap");
    };
    assert_eq!((image.bitmap.width, image.bitmap.height), (1, 1));
    assert!(to_svg(&repeated_nested_images(&inner, 1)).is_ok());
    assert_eq!(
        to_svg(&repeated_nested_images(&inner, 17)).unwrap_err().0,
        "the metafile draws more than the replay limits"
    );
}

#[test]
fn failed_nested_images_spend_work_and_decoded_pixels() {
    let mut inner = cropped_rle_emf(1024);
    let last_record = inner.len() - 20;
    inner[last_record..last_record + 4].copy_from_slice(&250u32.to_le_bytes());
    assert!(replay(&inner).is_err());
    let mut budget = ReplayBudget::default();
    let before = budget;
    let svg = to_svg_with_budget(&repeated_nested_images(&inner, 1), &mut budget).unwrap();
    assert!(svg.markup.contains("<path"));
    assert_eq!(before.pixels, budget.pixels);
    assert!(before.work - budget.work >= 1024 * 1024 / 16);
    assert_eq!(
        to_svg(&repeated_nested_images(&inner, 17)).unwrap_err().0,
        "the metafile draws more than the replay limits"
    );
}

#[test]
fn refused_replays_keep_their_cumulative_spending() {
    let mut broken = cropped_rle_emf(256);
    broken.truncate(broken.len() - 20);
    let mut budget = ReplayBudget {
        work: ReplayBudget::default().work,
        pixels: 2 * 256 * 256,
    };
    let before = budget.work;
    assert!(replay_with_budget(&broken, &mut budget).is_err());
    assert_eq!(budget.pixels, 256 * 256);
    assert!(budget.work < before);
    assert!(replay_with_budget(&broken, &mut budget).is_err());
    assert_eq!(budget.pixels, 0);
    assert_eq!(
        replay_with_budget(&cropped_rle_emf(256), &mut budget)
            .unwrap_err()
            .0,
        "the metafile draws more than the replay limits"
    );
}

#[test]
fn failed_nested_images_spend_cumulative_record_work() {
    let inner = Emf::new(10, 10).rec(250, &[]).bytes();
    let mut budget = ReplayBudget {
        work: 100,
        ..ReplayBudget::default()
    };
    assert_eq!(
        replay_with_budget(&repeated_nested_images(&inner, 30), &mut budget)
            .unwrap_err()
            .0,
        "the metafile draws more than the replay limits"
    );
    assert_eq!(budget.work, 0);
    assert_eq!(budget.pixels, ReplayBudget::default().pixels);
}

#[test]
fn repeated_nested_text_spends_cumulative_work() {
    let inner = repeated_gdi_text(512, 20);
    let allowance = ReplayBudget {
        work: 30_000,
        ..ReplayBudget::default()
    };
    let mut budget = allowance;
    assert!(replay_with_budget(&repeated_nested_images(&inner, 1), &mut budget).is_ok());
    assert!(allowance.work - budget.work >= 20 * 512);
    let mut budget = allowance;
    assert_eq!(
        replay_with_budget(&repeated_nested_images(&inner, 30), &mut budget)
            .unwrap_err()
            .0,
        "the metafile draws more than the replay limits"
    );
    assert_eq!(budget.work, 0);
    assert_eq!(budget.pixels, allowance.pixels);
}

#[test]
fn gdi_text_advances_spend_cumulative_work() {
    let text = "a".repeat(512);
    let mut pictures = Vec::new();
    for (options, length, work) in [(0x10, 512, 768), (0x2010, 1024, 1280)] {
        pictures.push((
            Emf::new(10, 10)
                .recs(vec![text_out(
                    0,
                    0,
                    &text,
                    Some(&vec![1; length]),
                    options,
                    [0, 0, -1, -1],
                )])
                .bytes(),
            work,
        ));
    }
    let mut body = i16s(&[0, 0, text.len() as i16]);
    body.extend(u16s(&[0x10]));
    body.extend(text.as_bytes());
    body.extend(i16s(&vec![1; text.len()]));
    pictures.push((Wmf::new(10, 10, 96).rec(0x0A32, &body).bytes(), 768));
    for (bytes, work) in pictures {
        assert!(replay(&bytes).is_ok());
        let mut budget = ReplayBudget {
            work,
            ..ReplayBudget::default()
        };
        assert_eq!(
            replay_with_budget(&bytes, &mut budget).unwrap_err().0,
            "the metafile draws more than the replay limits"
        );
        assert_eq!(budget.work, 0);
    }
}

#[test]
fn emf_plus_text_and_positions_spend_cumulative_work() {
    let text = "a".repeat(512);
    let mut string = u32s(&[0xff00_0000, 99, text.len() as u32]);
    string.extend(f32s(&[0.0; 4]));
    string.extend(u16s(&utf16(&text)));
    for (record, work) in [
        ((0x401C, 0x8000, string), 256),
        (
            plus_driver_string(0, 0xff00_0000, &text, (0.0, 0.0), &[]),
            768,
        ),
    ] {
        let bytes = Emf::new(10, 10)
            .rec(
                70,
                &plus(&[
                    plus_header(false),
                    plus_font(0, 12.0, 0, "Arial"),
                    record,
                    plus_eof(),
                ])
                .1,
            )
            .bytes();
        assert!(replay(&bytes).is_ok());
        let mut budget = ReplayBudget {
            work,
            ..ReplayBudget::default()
        };
        assert_eq!(
            replay_with_budget(&bytes, &mut budget).unwrap_err().0,
            "the metafile draws more than the replay limits"
        );
        assert_eq!(budget.work, 0);
    }
}

#[test]
fn a_failed_emf_in_a_wmf_keeps_its_spending() {
    let mut inner = cropped_rle_emf(256);
    inner.truncate(inner.len() - 20);
    let mut escape = u16s(&[0x000F, (34 + inner.len()) as u16]);
    escape.extend(u32s(&[0x4346_4D57, 1, 0x0001_0000]));
    escape.extend(u16s(&[0]));
    escape.extend(u32s(&[0, 1, inner.len() as u32, 0, inner.len() as u32]));
    escape.extend(&inner);
    let bytes = Wmf::new(1440, 1440, 1440)
        .recs(vec![(0x0626, escape), (0x041B, i16s(&[10, 10, 0, 0]))])
        .bytes();
    let mut budget = ReplayBudget::default();
    let before = budget;
    assert!(replay_with_budget(&bytes, &mut budget).is_ok());
    assert_eq!(before.pixels - budget.pixels, 256 * 256);
    let mut budget = ReplayBudget {
        pixels: 256 * 256 - 1,
        ..ReplayBudget::default()
    };
    assert_eq!(
        replay_with_budget(&bytes, &mut budget).unwrap_err().0,
        "the metafile draws more than the replay limits"
    );
}

#[test]
fn every_placement_of_a_nested_metafile_spends_the_budget() {
    let rects: Vec<[f32; 4]> = (0..20_000).map(|i| [i as f32, 0.0, 1.0, 1.0]).collect();
    let inner = Emf::new(10, 10)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                plus_fill_rects(0xff00_0000, &rects),
                plus_eof(),
            ])
            .1,
        )
        .bytes();
    let outer = |draws: usize| {
        let mut records = vec![plus_header(false), nested_image(&inner)];
        records.extend(
            (0..draws).map(|_| plus_draw_image(0, [0.0, 0.0, 10.0, 10.0], [0.0, 0.0, 10.0, 10.0])),
        );
        records.push(plus_eof());
        Emf::new(10, 10).rec(70, &plus(&records).1).bytes()
    };
    assert!(to_svg(&outer(3)).is_ok());
    assert!(to_svg(&outer(100)).is_err());
}

#[test]
fn united_clip_regions_spend_the_budget() {
    let mut records = vec![select_clip_region(5, &[[0, 0, 1, 1]])];
    for index in 0..3_000 {
        records.push(select_clip_region(
            2,
            &[[index % 100, 0, index % 100 + 1, 1]],
        ));
        records.push(rect(0, 0, 10, 10));
    }
    assert!(to_svg(&Emf::new(100, 100).recs(records).bytes()).is_err());
}

#[test]
fn a_last_pending_line_past_the_strict_op_limit_refuses() {
    let drawing = |rects: usize| {
        let mut records: Vec<(u32, Vec<u8>)> = (0..rects).map(|_| rect(0, 0, 5, 5)).collect();
        records.push(move_to(0, 0));
        records.push(line_to(5, 5));
        Emf::new(10, 10).recs(records).bytes()
    };
    assert!(decode(&drawing(4_095)).is_some());
    assert!(decode(&drawing(4_096)).is_none());
}

#[test]
fn repainting_one_wmf_region_spends_the_budget() {
    let mut region = u16s(&[0, 6]);
    region.extend(u32s(&[0]));
    region.extend(i16s(&[0, 1, 0, 0, 0, 1000, 1000]));
    region.extend(u16s(&[20_000]));
    region.extend(i16s(&[0, 10]));
    for pair in 0..10_000i16 {
        region.extend(i16s(&[pair % 1000, pair % 1000 + 1]));
    }
    region.extend(u16s(&[20_000]));
    let wmf = |paints: usize| {
        let mut records = vec![
            (
                0x02FC,
                [u16s(&[0]), u32s(&[0x00ff_0000]), u16s(&[0])].concat(),
            ),
            (0x012D, u16s(&[0])),
            (0x06FF, region.clone()),
        ];
        records.extend((0..paints).map(|_| (0x012B, u16s(&[1]))));
        Wmf::new(1000, 1000, 1440).recs(records).bytes()
    };
    assert!(to_svg(&wmf(3)).is_ok());
    assert!(to_svg(&wmf(100)).is_err());
}

#[test]
fn a_wide_emf_plus_region_tree_refuses_before_chaining_its_clips() {
    fn tree(depth: u32, out: &mut Vec<u8>) {
        if depth == 0 {
            out.extend(u32s(&[0x1000_0000]));
            out.extend(f32s(&[0.0, 0.0, 5.0, 5.0]));
            return;
        }
        out.extend(u32s(&[1]));
        tree(depth - 1, out);
        tree(depth - 1, out);
    }
    let mut region = u32s(&[0xDBC0_1002, (1 << 16) - 2]);
    tree(15, &mut region);
    let bytes = Emf::new(10, 10)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                (0x4008, 0x0400, region),
                (0x4034, 0, Vec::new()),
                plus_fill_rects(0xff00_0000, &[[0.0, 0.0, 5.0, 5.0]]),
            ])
            .1,
        )
        .bytes();
    assert!(to_svg(&bytes).is_err());
}

#[test]
fn a_compressed_dib_is_charged_its_real_size() {
    let mut bmi = u32s(&[40, 1, 1]);
    bmi.extend(u16s(&[1, 0]));
    bmi.extend(u32s(&[5, 0, 0, 0, 0, 0]));
    let dib = |tag: u8| {
        let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
        png.extend(4096u32.to_be_bytes());
        png.extend(4096u32.to_be_bytes());
        png.extend([8, 6, 0, 0, 0, tag]);
        (bmi.clone(), png)
    };
    let blits = |count: u8| {
        Emf::new(10, 10)
            .recs(
                (0..count)
                    .map(|tag| stretch_dibits([0, 0, 10, 10], &dib(tag), 0x00CC_0020))
                    .collect(),
            )
            .bytes()
    };
    assert!(to_svg(&blits(1)).is_ok());
    assert!(to_svg(&blits(2)).is_err());
}

#[test]
fn run_length_bitmaps_stop_expanding_past_their_width() {
    let mut bmi = u32s(&[40, 1, 1]);
    bmi.extend(u16s(&[1, 8]));
    bmi.extend(u32s(&[1, 0, 0, 0, 2, 0]));
    bmi.extend(u32s(&[0, 0x00ff_ffff]));
    let runs: Vec<u8> = std::iter::repeat_n([255u8, 1], 200_000).flatten().collect();
    check(
        &Emf::new(10, 10)
            .recs(vec![stretch_dibits(
                [0, 0, 10, 10],
                &(bmi, runs),
                0x00CC_0020,
            )])
            .bytes(),
    );
}

/// Sides of a bitmap whose pixels fit the replay's pixel budget once, not twice.
const SIDE: u32 = 3000;
const SRCCOPY: u32 = 0x00CC_0020;

/// A PNG header declaring `SIDE` x `SIDE`; pictures stay encoded, so the body
/// is never decoded and `tag` tells them apart.
fn png(tag: u8) -> Vec<u8> {
    let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
    png.extend(13u32.to_be_bytes());
    png.extend(b"IHDR");
    png.extend(SIDE.to_be_bytes());
    png.extend(SIDE.to_be_bytes());
    png.extend([8, 6, 0, 0, 0, 0, 0, 0, 0, tag]);
    png
}

fn png_dib(tag: u8) -> (Vec<u8>, Vec<u8>) {
    let bits = png(tag);
    let mut bmi = u32s(&[40, SIDE, SIDE]);
    bmi.extend(u16s(&[1, 0]));
    bmi.extend(u32s(&[5, bits.len() as u32, 0, 0, 0, 0]));
    (bmi, bits)
}

fn plus_png(id: u8, tag: u8) -> (u16, u16, Vec<u8>) {
    let mut body = u32s(&[0xDBC0_1002, 1]);
    body.extend(i32s(&[SIDE as i32, SIDE as i32, 0]));
    body.extend(u32s(&[0, 1]));
    body.extend(png(tag));
    (0x4008, 0x0500 | u16::from(id), body)
}

fn draws(op: &Op, tag: u8) -> bool {
    matches!(op, Op::Image(image) if matches!(&image.bitmap.pixels, Pixels::Encoded { bytes, .. } if bytes.get(33) == Some(&tag)))
}

fn drawn_twice(bytes: &[u8]) -> bool {
    replay(bytes).is_ok_and(|drawing| drawing.ops.len() == 2)
}

fn pixels_spent(bytes: &[u8]) -> u64 {
    let mut budget = ReplayBudget::default();
    let before = budget.pixels;
    replay_with_budget(bytes, &mut budget).unwrap();
    before - budget.pixels
}

#[test]
fn a_dual_metafile_falls_back_to_gdi_with_the_pixels_its_refused_emf_plus_pass_held() {
    let bytes = Emf::new(100, 100)
        .rec(
            70,
            &plus(&[plus_header(true), plus_png(1, 1), plus_png(2, 2)]).1,
        )
        .recs(vec![stretch_dibits([0, 0, 50, 50], &png_dib(3), SRCCOPY)])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(matches!(&drawing.ops[..], [op] if draws(op, 3)));
    assert_eq!(pixels_spent(&bytes), u64::from(SIDE * SIDE));
}

#[test]
fn an_honoured_emf_plus_rendition_never_decodes_its_gdi_bitmaps() {
    let mut huge = png_dib(3);
    huge.0[4..12].copy_from_slice(&i32s(&[8192, 8192]));
    let bytes = Emf::new(100, 100)
        .rec(
            70,
            &plus(&[
                plus_header(true),
                plus_png(1, 1),
                plus_draw_image(
                    1,
                    [0.0, 0.0, SIDE as f32, SIDE as f32],
                    [0.0, 0.0, 50.0, 50.0],
                ),
            ])
            .1,
        )
        .recs(vec![stretch_dibits([0, 0, 50, 50], &huge, SRCCOPY)])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(matches!(&drawing.ops[..], [op] if draws(op, 1)));
    assert_eq!(pixels_spent(&bytes), u64::from(SIDE * SIDE));
}

#[test]
fn equal_gdi_bitmaps_share_one_charge_and_different_ones_do_not() {
    let same = png_dib(1);
    let bytes = Emf::new(100, 100)
        .recs(vec![
            stretch_dibits([0, 0, 50, 50], &same, SRCCOPY),
            stretch_dibits([50, 50, 50, 50], &same, SRCCOPY),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Image(first), Op::Image(second)] = &drawing.ops[..] else {
        panic!("both pictures");
    };
    assert!(Arc::ptr_eq(&first.bitmap, &second.bitmap));
    assert_eq!(pixels_spent(&bytes), u64::from(SIDE * SIDE));
    let different = Emf::new(100, 100)
        .recs(vec![
            stretch_dibits([0, 0, 50, 50], &png_dib(1), SRCCOPY),
            stretch_dibits([50, 50, 50, 50], &png_dib(2), SRCCOPY),
        ])
        .bytes();
    assert!(!drawn_twice(&different));
}

#[test]
fn equal_emf_plus_images_share_one_charge_and_different_ones_do_not() {
    let image = |tags: [u8; 2]| {
        Emf::new(100, 100)
            .rec(
                70,
                &plus(&[
                    plus_header(false),
                    plus_png(1, tags[0]),
                    plus_png(2, tags[1]),
                    plus_draw_image(
                        1,
                        [0.0, 0.0, SIDE as f32, SIDE as f32],
                        [0.0, 0.0, 50.0, 50.0],
                    ),
                    plus_draw_image(
                        2,
                        [0.0, 0.0, SIDE as f32, SIDE as f32],
                        [50.0, 50.0, 50.0, 50.0],
                    ),
                    plus_eof(),
                ])
                .1,
            )
            .bytes()
    };
    let same = image([1, 1]);
    let drawing = replay(&same).unwrap();
    let [Op::Image(first), Op::Image(second)] = &drawing.ops[..] else {
        panic!("both pictures");
    };
    assert!(Arc::ptr_eq(&first.bitmap, &second.bitmap));
    assert_eq!(pixels_spent(&same), u64::from(SIDE * SIDE));
    assert!(!drawn_twice(&image([1, 2])));
}

/// A blank run-length bitmap declaring 4096 x 4096 in a few bytes, drawn
/// `count` times with its source rectangle moved by `step` each time.
fn blank_blits(count: i32, step: i32) -> Vec<u8> {
    let mut bmi = u32s(&[40, 4096, 4096]);
    bmi.extend(u16s(&[1, 8]));
    bmi.extend(u32s(&[1, 2, 0, 0, 2, 0, 0, 0x00ff_ffff]));
    let dib = (bmi, vec![0, 1]);
    Emf::new(10, 10)
        .recs(
            (0..count)
                .map(|draw| {
                    let (kind, mut body) = stretch_dibits([0, 0, 10, 10], &dib, SRCCOPY);
                    body[24..28].copy_from_slice(&i32s(&[draw * step]));
                    body[32..40].copy_from_slice(&i32s(&[4000, 4000]));
                    (kind, body)
                })
                .collect(),
        )
        .bytes()
}

fn work_spent(bytes: &[u8]) -> u64 {
    let mut budget = ReplayBudget::default();
    let before = budget.work;
    replay_with_budget(bytes, &mut budget).unwrap();
    before - budget.work
}

#[test]
fn a_bitmap_drawn_again_the_same_way_is_not_decoded_again() {
    let drawing = replay(&blank_blits(64, 0)).unwrap();
    assert_eq!(drawing.ops.len(), 64);
    assert!(drawing.ops.windows(2).all(|pair| matches!(
        pair,
        [Op::Image(first), Op::Image(second)] if Arc::ptr_eq(&first.bitmap, &second.bitmap)
    )));
    assert!(work_spent(&blank_blits(64, 0)) < work_spent(&blank_blits(1, 0)) + 64 * 64);
}

#[test]
fn equal_pixels_cropped_from_one_source_are_decoded_a_bounded_number_of_times() {
    let drawing = replay(&blank_blits(2, 1)).unwrap();
    let [Op::Image(first), Op::Image(second)] = &drawing.ops[..] else {
        panic!("both pictures");
    };
    assert!(Arc::ptr_eq(&first.bitmap, &second.bitmap));
    assert!(replay(&blank_blits(8, 1)).is_err());
}

#[test]
fn an_emf_plus_image_that_cannot_be_built_is_not_rebuilt_for_free() {
    let objects = |count: u16| {
        let mut body = u32s(&[0xDBC0_1002, 1]);
        body.extend(i32s(&[2048, 2048, 0]));
        body.extend(u32s(&[0, 0]));
        let mut records = vec![plus_header(false)];
        records.extend((0..count).map(|id| (0x4008, 0x0500 | id, body.clone())));
        records.push(plus_eof());
        Emf::new(10, 10).rec(70, &plus(&records).1).bytes()
    };
    assert!(work_spent(&objects(64)) < work_spent(&objects(1)) + 64 * 64);
}

#[test]
fn a_bitmap_that_cannot_be_decoded_is_not_decoded_again() {
    let mut bmi = u32s(&[40, 4096, 4096]);
    bmi.extend(u16s(&[1, 3]));
    bmi.extend(u32s(&[0, 0, 0, 0, 0, 0]));
    let blits = |count: usize| {
        Emf::new(10, 10)
            .recs(vec![
                stretch_dibits(
                    [0, 0, 10, 10],
                    &(bmi.clone(), vec![0; 4]),
                    SRCCOPY
                );
                count
            ])
            .bytes()
    };
    let drawing = replay(&blits(64)).unwrap();
    assert_eq!(
        drawing
            .omissions
            .iter()
            .map(|omission| omission.count)
            .sum::<usize>(),
        64
    );
    assert!(work_spent(&blits(64)) < work_spent(&blits(1)) + 64 * 64);
}

#[test]
fn an_emf_plus_only_metafile_replayed_again_gets_back_its_first_pass_pixels() {
    let side: u32 = 4096;
    let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
    png.extend(13u32.to_be_bytes());
    png.extend(b"IHDR");
    png.extend(side.to_be_bytes());
    png.extend(side.to_be_bytes());
    png.extend([8, 6, 0, 0, 0, 0, 0, 0, 0, 7]);
    let mut image = u32s(&[0xDBC0_1002, 1]);
    image.extend(i32s(&[side as i32, side as i32, 0]));
    image.extend(u32s(&[0, 1]));
    image.extend(png);
    let bytes = Emf::new(100, 100)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                (0x4008, 0x0501, image),
                plus_draw_image(
                    1,
                    [0.0, 0.0, side as f32, side as f32],
                    [0.0, 0.0, 50.0, 50.0],
                ),
                (0x4037, 0, Vec::new()),
                plus_eof(),
            ])
            .1,
        )
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(drawing.ops.iter().any(|op| draws(op, 7)));
    assert_eq!(pixels_spent(&bytes), u64::from(side * side));
}

#[test]
fn a_dual_fallback_draws_bitmaps_its_set_aside_emf_plus_pass_also_held() {
    let bytes = Emf::new(100, 100)
        .rec(
            70,
            &plus(&[
                plus_header(true),
                plus_png(1, 1),
                plus_draw_image(
                    1,
                    [0.0, 0.0, SIDE as f32, SIDE as f32],
                    [0.0, 0.0, 50.0, 50.0],
                ),
                (0x4037, 0, Vec::new()),
                plus_eof(),
            ])
            .1,
        )
        .recs(vec![stretch_dibits([0, 0, 50, 50], &png_dib(3), SRCCOPY)])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(matches!(&drawing.ops[..], [op] if draws(op, 3)));
    assert_eq!(pixels_spent(&bytes), u64::from(SIDE * SIDE));
}
