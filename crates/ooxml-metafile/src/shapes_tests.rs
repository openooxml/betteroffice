use super::*;
use crate::emf::{EMF_SIGNATURE, RGN_COPY};
use crate::gradient::{
    GRADIENT_RECT_V, GRADIENT_TRIANGLE, MAX_GRADIENT_VERTICES, ROP_BLACKNESS, ROP_DSTCOPY,
    ROP_PATCOPY, ROP_WHITENESS,
};
use crate::wmf::WMF_PLACEABLE_KEY;

/// POLYGON16 declaring eight points in a 40-byte record that holds three (#318).
const EMF_POLYGON_READS_PAST_ITS_RECORD: &[u8] =
    include_bytes!("../tests/fixtures/emf-polygon-reads-past-its-record.emf");
const EMF_POLYGON_COUNT_AT: usize = 0x70;

/// META_POLYGON declaring four points in a 10-word record that holds three (#318).
const WMF_POLYGON_READS_PAST_ITS_RECORD: &[u8] =
    include_bytes!("../tests/fixtures/wmf-polygon-reads-past-its-record.wmf");
const WMF_POLYGON_COUNT_AT: usize = 0x18;

fn with_count(bytes: &[u8], at: usize, count: u8) -> Vec<u8> {
    let mut bytes = bytes.to_vec();
    bytes[at] = count;
    bytes
}

#[test]
fn an_emf_polygon_counting_past_its_record_yields_no_drawing() {
    assert!(decode(EMF_POLYGON_READS_PAST_ITS_RECORD).is_none());
    assert!(
        decode(&with_count(
            EMF_POLYGON_READS_PAST_ITS_RECORD,
            EMF_POLYGON_COUNT_AT,
            3
        ))
        .is_some()
    );
}

#[test]
fn a_wmf_polygon_counting_past_its_record_yields_no_drawing() {
    assert!(decode(WMF_POLYGON_READS_PAST_ITS_RECORD).is_none());
    assert!(
        decode(&with_count(
            WMF_POLYGON_READS_PAST_ITS_RECORD,
            WMF_POLYGON_COUNT_AT,
            3
        ))
        .is_some()
    );
}

fn emf_header(bounds: [i32; 4], frame_device: [i32; 2]) -> Vec<u8> {
    let mut header = vec![0u8; 88];
    let put = |header: &mut Vec<u8>, offset: usize, value: i32| {
        header[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    };
    put(&mut header, 0, 1);
    put(&mut header, 4, 88);
    for (index, value) in bounds.iter().enumerate() {
        put(&mut header, 8 + index * 4, *value);
    }
    for (index, value) in [0, 0, frame_device[0] * 100, frame_device[1] * 100]
        .iter()
        .enumerate()
    {
        put(&mut header, 24 + index * 4, *value);
    }
    put(&mut header, 40, EMF_SIGNATURE as i32);
    put(&mut header, 56, 16);
    put(&mut header, 72, frame_device[0]);
    put(&mut header, 76, frame_device[1]);
    put(&mut header, 80, frame_device[0]);
    put(&mut header, 84, frame_device[1]);
    header
}

fn record(kind: u32, body: &[u8]) -> Vec<u8> {
    let size = 8 + body.len();
    let mut out = kind.to_le_bytes().to_vec();
    out.extend_from_slice(&(size as u32).to_le_bytes());
    out.extend_from_slice(body);
    out
}

fn i32s(values: &[i32]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

fn i16s(values: &[i16]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

fn poly16(points: &[(i16, i16)]) -> Vec<u8> {
    let mut body = i32s(&[0, 0, 0, 0, points.len() as i32]);
    for (x, y) in points {
        body.extend_from_slice(&i16s(&[*x, *y]));
    }
    body
}

fn solid_brush(handle: u32, color: u32) -> Vec<u8> {
    let mut out = record(39, &i32s(&[handle as i32, 0, color as i32, 0]));
    out.extend(record(37, &i32s(&[handle as i32])));
    out
}

fn emf(records: Vec<Vec<u8>>, bounds: [i32; 4], frame_device: [i32; 2]) -> Vec<u8> {
    let mut bytes = emf_header(bounds, frame_device);
    for chunk in records {
        bytes.extend(chunk);
    }
    bytes.extend(record(14, &i32s(&[0, 0, 0])));
    bytes
}

fn only_op(drawing: &MetafileDrawing) -> &MetafileOp {
    assert_eq!(drawing.ops.len(), 1, "expected one op");
    &drawing.ops[0]
}

#[test]
fn a_bracketed_path_fill_becomes_one_op_in_frame_fractions() {
    let bytes = emf(
        vec![
            solid_brush(1, 0x0000_00ff),
            record(59, &[]),
            record(27, &i32s(&[0, 0])),
            record(89, &poly16(&[(50, 0), (0, 50)])),
            record(61, &[]),
            record(60, &[]),
            record(62, &i32s(&[0, 0, 0, 0])),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );

    let drawing = decode(&bytes).expect("the path fill decodes");
    let op = only_op(&drawing);
    assert_eq!(op.fill.as_deref(), Some("#ff0000"));
    assert!(op.stroke.is_none(), "FILLPATH does not stroke");
    assert_eq!(
        op.path,
        vec![
            PathCommand::Move { x: 0.0, y: 0.0 },
            PathCommand::Line { x: 0.5, y: 0.0 },
            PathCommand::Line { x: 0.0, y: 0.5 },
            PathCommand::Close,
        ]
    );
}

/// `BEGINPATH; MOVETOEX; POLYLINETO16; ENDPATH; SELECTCLIPPATH` over the
/// left-top quarter of a 100x100 frame, which is how PowerPoint brackets
/// each part of an exported drawing.
fn quarter_clip_path() -> Vec<Vec<u8>> {
    vec![
        record(59, &[]),
        record(27, &i32s(&[0, 0])),
        record(89, &poly16(&[(0, 0), (0, 50), (50, 50), (50, 0)])),
        record(60, &[]),
        record(67, &i32s(&[RGN_COPY as i32])),
    ]
}

fn filled_triangle() -> Vec<Vec<u8>> {
    vec![
        solid_brush(1, 0x0000_0000),
        record(86, &poly16(&[(0, 0), (99, 0), (99, 99)])),
    ]
}

#[test]
fn a_rectangular_clip_path_bounds_the_ops_that_follow() {
    let mut records = quarter_clip_path();
    records.extend(filled_triangle());
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the clipped fill decodes");
    assert_eq!(only_op(&drawing).clip, Some([0.0, 0.0, 0.5, 0.5]));
}

#[test]
fn an_empty_copy_region_puts_the_clip_back() {
    let mut records = quarter_clip_path();
    records.push(record(75, &i32s(&[0, RGN_COPY as i32])));
    records.extend(filled_triangle());
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the fill decodes");
    assert!(only_op(&drawing).clip.is_none());
}

#[test]
fn a_saved_clip_comes_back_with_its_device_context() {
    let mut records = vec![record(33, &[])];
    records.extend(quarter_clip_path());
    records.push(record(34, &i32s(&[-1])));
    records.extend(filled_triangle());
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the fill decodes");
    assert!(only_op(&drawing).clip.is_none());
}

#[test]
fn a_clip_this_player_cannot_represent_yields_no_drawing() {
    let triangle = vec![
        record(59, &[]),
        record(27, &i32s(&[0, 0])),
        record(89, &poly16(&[(0, 50), (50, 50)])),
        record(60, &[]),
        record(67, &i32s(&[RGN_COPY as i32])),
    ];
    let region = i32s(&[
        32,
        RGN_COPY as i32,
        32,
        1,
        1,
        32,
        0,
        0,
        50,
        50,
        0,
        0,
        50,
        50,
    ]);
    for clip in [triangle, vec![record(75, &region)]] {
        let mut records = clip;
        records.extend(filled_triangle());
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);
        assert!(decode(&bytes).is_none());
    }
}

#[test]
fn coordinates_scale_by_the_frame_rectangle_not_the_ink_bounds() {
    let records = vec![
        solid_brush(1, 0x0000_0000),
        record(86, &poly16(&[(20, 20), (60, 20), (60, 60)])),
    ];
    let bytes = emf(records, [20, 20, 60, 60], [80, 80]);

    let drawing = decode(&bytes).expect("the polygon decodes");
    let op = only_op(&drawing);
    assert_eq!(op.path[0], PathCommand::Move { x: 0.25, y: 0.25 });
    assert_eq!(op.path[1], PathCommand::Line { x: 0.75, y: 0.25 });
}

#[test]
fn a_window_extent_without_a_viewport_extent_does_not_scale() {
    let records = vec![
        record(10, &i32s(&[0, 0])),
        record(9, &i32s(&[1000, 1000])),
        solid_brush(1, 0x0000_0000),
        record(86, &poly16(&[(50, 0), (100, 0), (100, 50)])),
    ];
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the polygon decodes");
    assert_eq!(
        only_op(&drawing).path[0],
        PathCommand::Move { x: 0.5, y: 0.0 }
    );
}

#[test]
fn a_window_and_viewport_extent_pair_scales_logical_units() {
    let records = vec![
        record(17, &i32s(&[8])),
        record(9, &i32s(&[200, 200])),
        record(11, &i32s(&[100, 100])),
        solid_brush(1, 0x0000_0000),
        record(86, &poly16(&[(100, 0), (200, 0), (200, 100)])),
    ];
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the polygon decodes");
    assert_eq!(
        only_op(&drawing).path[0],
        PathCommand::Move { x: 0.5, y: 0.0 }
    );
}

#[test]
fn a_pie_sweeps_counter_clockwise_from_its_start_ray() {
    let records = vec![
        solid_brush(1, 0x0000_0000),
        record(47, &i32s(&[0, 0, 100, 100, 50, 0, 100, 50])),
    ];
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the pie decodes");
    let op = only_op(&drawing);
    assert_eq!(op.path[0], PathCommand::Move { x: 0.5, y: 0.5 });
    let on_arc: Vec<_> = op
        .path
        .iter()
        .filter_map(|command| match command {
            PathCommand::Line { x, y } => Some((*x, *y)),
            _ => None,
        })
        .collect();
    assert_eq!(on_arc.first().copied(), Some((0.5, 0.0)));
    let (last_x, last_y) = on_arc.last().copied().expect("the arc has points");
    assert!((last_x - 1.0).abs() < 1e-6 && (last_y - 0.5).abs() < 1e-6);
    assert!(
        on_arc.iter().any(|(x, _)| *x < 0.01),
        "the wedge never reached nine o'clock: {on_arc:?}"
    );
    assert_eq!(op.path.last(), Some(&PathCommand::Close));
}

#[test]
fn a_null_brush_leaves_a_polygon_unfilled() {
    let mut records = vec![record(39, &i32s(&[1, 1, 0x00ff_00ff, 0]))];
    records.push(record(37, &i32s(&[1])));
    records.push(record(86, &poly16(&[(0, 0), (50, 0), (50, 50)])));
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);

    let drawing = decode(&bytes).expect("the polygon decodes");
    assert!(only_op(&drawing).fill.is_none());
}

#[test]
fn bytes_that_are_not_a_metafile_are_left_to_the_image_decoder() {
    assert!(decode(b"\xff\xd8\xff\xe0 not a metafile at all").is_none());
    assert!(decode(&[]).is_none());
}

#[test]
fn a_metafile_that_draws_nothing_is_not_decoded() {
    let bytes = emf(Vec::new(), [0, 0, 99, 99], [100, 100]);
    assert!(decode(&bytes).is_none());
}

#[test]
fn a_truncated_record_ends_the_replay_without_panicking() {
    let mut bytes = emf(
        vec![
            solid_brush(1, 0x0000_0000),
            record(86, &poly16(&[(0, 0), (50, 0), (50, 50)])),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );
    assert!(decode(&bytes).is_some());
    bytes.truncate(bytes.len() - 12);
    assert!(decode(&bytes).is_none());

    let mut lying = emf_header([0, 0, 99, 99], [100, 100]);
    lying.extend(record(86, &i32s(&[0, 0, 0, 0, 100_000])));
    assert!(decode(&lying).is_none());
}

const WMF_LINETO_SHORTER_THAN_ITS_POINT: &[u8] = &[
    0x01, 0x00, 0x09, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x02, 0x00, 0x00, 0x0a, 0x00, 0x00, 0x00, 0x24, 0x03,
    0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x03, 0x00,
    0x00, 0x00, 0x00, 0x00,
];

#[test]
fn a_wmf_move_or_line_too_short_for_its_point_rejects_the_metafile() {
    for function in [0x0213u16, 0x0214] {
        let mut short = WMF_LINETO_SHORTER_THAN_ITS_POINT.to_vec();
        short[22..24].copy_from_slice(&function.to_le_bytes());
        assert!(decode(&short).is_none(), "function {function:#06x}");

        let mut honest = short.clone();
        honest[18] = 5;
        honest.splice(26..26, [0, 0]);
        assert!(decode(&honest).is_some(), "function {function:#06x}");
    }
}

#[test]
fn a_wmf_polygon_fills_with_the_selected_brush() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&WMF_PLACEABLE_KEY.to_le_bytes());
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&i16s(&[0, 0, 100, 100]));
    bytes.extend_from_slice(&1440u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&1u16.to_le_bytes());
    bytes.extend_from_slice(&9u16.to_le_bytes());
    bytes.extend_from_slice(&0x0300u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(&0u16.to_le_bytes());

    let wmf_record = |function: u16, body: &[u8]| {
        let words = (6 + body.len()) / 2;
        let mut out = (words as u32).to_le_bytes().to_vec();
        out.extend_from_slice(&function.to_le_bytes());
        out.extend_from_slice(body);
        out
    };
    bytes.extend(wmf_record(0x020B, &i16s(&[0, 0])));
    bytes.extend(wmf_record(0x020C, &i16s(&[100, 100])));
    let mut brush = 0u16.to_le_bytes().to_vec();
    brush.extend_from_slice(&0x0000_00ffu32.to_le_bytes());
    brush.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend(wmf_record(0x02FC, &brush));
    bytes.extend(wmf_record(0x012D, &0u16.to_le_bytes()));
    let mut polygon = 3u16.to_le_bytes().to_vec();
    polygon.extend_from_slice(&i16s(&[0, 0, 50, 0, 50, 50]));
    bytes.extend(wmf_record(0x0324, &polygon));
    bytes.extend(wmf_record(0, &[]));

    let drawing = decode(&bytes).expect("the WMF decodes");
    let op = only_op(&drawing);
    assert_eq!(op.fill.as_deref(), Some("#ff0000"));
    assert_eq!(op.path[1], PathCommand::Line { x: 0.5, y: 0.0 });
}
#[test]
fn records_cannot_read_points_from_their_successors() {
    let bytes = emf(
        vec![
            record(86, &i32s(&[0, 0, 99, 99, 3])),
            record(70, &i32s(&[12, 0, 0, 0])),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );
    assert!(decode(&bytes).is_none());
    let mut bytes = vec![0u8; 18];
    bytes[0..2].copy_from_slice(&1u16.to_le_bytes());
    bytes[2..4].copy_from_slice(&9u16.to_le_bytes());
    bytes.extend(4u32.to_le_bytes());
    bytes.extend(0x0324u16.to_le_bytes());
    bytes.extend(3u16.to_le_bytes());
    bytes.extend(9u32.to_le_bytes());
    bytes.extend(0x0201u16.to_le_bytes());
    bytes.extend([0u8; 12]);
    bytes.extend(3u32.to_le_bytes());
    bytes.extend(0u16.to_le_bytes());
    assert!(decode(&bytes).is_none());
}

#[test]
fn selecting_a_brush_preserves_open_and_closed_paths() {
    for before_end in [false, true] {
        let mut records = vec![
            record(59, &[]),
            record(27, &i32s(&[0, 0])),
            record(89, &poly16(&[(100, 0), (0, 100)])),
            record(61, &[]),
        ];
        if before_end {
            records.push(solid_brush(1, 0x0000ff00));
        }
        records.push(record(60, &[]));
        if !before_end {
            records.push(solid_brush(1, 0x0000ff00));
        }
        records.push(record(62, &i32s(&[0, 0, 100, 100])));
        let drawing = decode(&emf(records, [0, 0, 99, 99], [100, 100])).unwrap();
        assert_eq!(only_op(&drawing).fill.as_deref(), Some("#00ff00"));
        assert_eq!(only_op(&drawing).path.len(), 4);
    }
}

#[test]
fn restore_dc_uses_absolute_levels_and_rejects_extreme_depths() {
    let bytes = emf(
        vec![
            record(33, &[]),
            record(10, &i32s(&[20, 20])),
            record(33, &[]),
            record(10, &i32s(&[40, 40])),
            record(34, &i32s(&[1])),
            record(86, &poly16(&[(0, 0), (50, 0), (50, 50)])),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );
    let drawing = decode(&bytes).unwrap();
    assert_eq!(
        only_op(&drawing).path[0],
        PathCommand::Move { x: 0.0, y: 0.0 }
    );
    let bytes = emf(
        vec![record(33, &[]), record(34, &i32s(&[i32::MIN]))],
        [0, 0, 99, 99],
        [100, 100],
    );
    assert!(decode(&bytes).is_none());
}

#[test]
fn world_transforms_keep_rotated_pen_widths_and_reject_nonfinite_values() {
    let matrix = |values: &[f32]| {
        values
            .iter()
            .flat_map(|v| v.to_le_bytes())
            .collect::<Vec<_>>()
    };
    let bytes = emf(
        vec![
            record(38, &i32s(&[1, 0, 10, 0, 0])),
            record(37, &i32s(&[1])),
            record(35, &matrix(&[0.0, 1.0, -1.0, 0.0, 100.0, 0.0])),
            record(43, &i32s(&[0, 0, 100, 100])),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );
    let drawing = decode(&bytes).unwrap();
    assert_eq!(only_op(&drawing).stroke.as_ref().unwrap().width, 0.1);
    for invalid in [f32::NAN, f32::INFINITY, f32::NEG_INFINITY] {
        let bytes = emf(
            vec![
                record(35, &matrix(&[invalid, 0.0, 0.0, 1.0, 0.0, 0.0])),
                record(43, &i32s(&[0, 0, 100, 100])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        assert!(decode(&bytes).is_none());
    }
}

#[test]
fn fill_mode_and_arc_direction_follow_the_device_context() {
    for mode in [1, 2] {
        let bytes = emf(
            vec![
                record(19, &i32s(&[mode])),
                record(57, &i32s(&[2])),
                record(47, &i32s(&[0, 0, 100, 100, 50, 0, 100, 50])),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        let drawing = decode(&bytes).unwrap();
        let op = only_op(&drawing);
        assert_eq!(op.even_odd, mode == 1);
        assert!(
            op.path
                .iter()
                .all(|p| !matches!(p, PathCommand::Line { x, .. } if *x < 0.49))
        );
    }
}

#[test]
fn unsupported_drawing_records_do_not_produce_partial_artwork() {
    // 84 EXTTEXTOUTW draws ink this replay cannot carry, so the drawing is
    // discarded rather than shown incomplete; 28 SETMETARGN and 82
    // EXTCREATEFONTINDIRECTW carry none, so the artwork around them
    // survives (#796).
    for (kind, drawn) in [(28u32, true), (82, true), (84, false)] {
        let bytes = emf(
            vec![
                record(43, &i32s(&[0, 0, 100, 100])),
                record(kind, &[0; 100]),
            ],
            [0, 0, 99, 99],
            [100, 100],
        );
        assert_eq!(decode(&bytes).is_some(), drawn, "record {kind}");
    }
    for mode in 1..=8u16 {
        let mut bytes = vec![0u8; 18];
        bytes[0..2].copy_from_slice(&1u16.to_le_bytes());
        bytes[2..4].copy_from_slice(&9u16.to_le_bytes());
        bytes.extend(4u32.to_le_bytes());
        bytes.extend(0x0103u16.to_le_bytes());
        bytes.extend(mode.to_le_bytes());
        bytes.extend(7u32.to_le_bytes());
        bytes.extend(0x041Bu16.to_le_bytes());
        for value in [1u16, 1, 0, 0] {
            bytes.extend(value.to_le_bytes());
        }
        bytes.extend(3u32.to_le_bytes());
        bytes.extend(0u16.to_le_bytes());
        assert_eq!(decode(&bytes).is_some(), mode == 8, "WMF mapping {mode}");
    }
}

fn blit(rop: u32, rect: [i32; 4], source_bits: bool) -> Vec<u8> {
    let mut body = i32s(&[0, 0, 0, 0]);
    body.extend(i32s(&rect));
    body.extend(rop.to_le_bytes());
    body.extend(i32s(&[0; 8]));
    body.extend(i32s(&[0, 0]));
    let bits = i32::from(source_bits);
    body.extend(i32s(&[0, bits * 40, 0, bits * 64]));
    record(76, &body)
}

fn gradient(mode: u32, vertices: &[(i32, i32, u16)], indexes: &[u32]) -> Vec<u8> {
    let corners = if mode == GRADIENT_TRIANGLE { 3 } else { 2 };
    let mut body = i32s(&[0, 0, 0, 0]);
    body.extend(i32s(&[
        vertices.len() as i32,
        (indexes.len() / corners) as i32,
        mode as i32,
    ]));
    for (x, y, grey) in vertices {
        body.extend(i32s(&[*x, *y]));
        for _ in 0..3 {
            body.extend((grey << 8).to_le_bytes());
        }
        body.extend(0u16.to_le_bytes());
    }
    for index in indexes {
        body.extend(index.to_le_bytes());
    }
    record(118, &body)
}

#[test]
fn a_pattern_blit_fills_its_destination_and_a_destination_copy_draws_nothing() {
    let bytes = emf(
        vec![
            solid_brush(1, 0x0000_00ff),
            blit(ROP_PATCOPY, [0, 0, 50, 50], false),
            blit(ROP_DSTCOPY, [0, 0, 50, 50], false),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );

    let drawing = decode(&bytes).expect("the blit decodes");
    let op = only_op(&drawing);
    assert_eq!(op.fill.as_deref(), Some("#ff0000"));
    assert!(op.stroke.is_none());
    assert_eq!(
        op.path,
        vec![
            PathCommand::Move { x: 0.0, y: 0.0 },
            PathCommand::Line { x: 0.5, y: 0.0 },
            PathCommand::Line { x: 0.5, y: 0.5 },
            PathCommand::Line { x: 0.0, y: 0.5 },
            PathCommand::Close,
        ]
    );
}

#[test]
fn a_constant_blit_paints_its_own_colour_and_ignores_the_selected_brush() {
    let bytes = emf(
        vec![
            solid_brush(1, 0x0000_00ff),
            blit(ROP_BLACKNESS, [0, 0, 50, 50], false),
            blit(ROP_WHITENESS, [50, 0, 50, 50], false),
        ],
        [0, 0, 99, 99],
        [100, 100],
    );

    let drawing = decode(&bytes).expect("the blits decode");
    assert_eq!(
        drawing
            .ops
            .iter()
            .map(|op| op.fill.clone().unwrap())
            .collect::<Vec<_>>(),
        ["#000000", "#ffffff"]
    );
}

#[test]
fn a_blit_with_source_bits_an_unknown_operation_or_a_short_record_is_rejected() {
    for records in [
        vec![blit(ROP_PATCOPY, [0, 0, 50, 50], true)],
        vec![blit(0x00CC_0020, [0, 0, 50, 50], false)],
        vec![record(76, &i32s(&[0; 12]))],
    ] {
        let mut records = records;
        records.insert(0, solid_brush(1, 0x0000_00ff));
        records.push(record(43, &i32s(&[0, 0, 50, 50])));
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);
        assert!(decode(&bytes).is_none());
    }
}

#[test]
fn a_vertical_rectangle_gradient_bands_from_the_first_vertex_to_the_second() {
    let bytes = emf(
        vec![gradient(
            GRADIENT_RECT_V,
            &[(0, 0, 0), (100, 100, 4)],
            &[0, 1],
        )],
        [0, 0, 99, 99],
        [100, 100],
    );

    let drawing = decode(&bytes).expect("the gradient decodes");
    assert_eq!(
        drawing
            .ops
            .iter()
            .map(|op| op.fill.clone().unwrap())
            .collect::<Vec<_>>(),
        ["#010101", "#020202", "#030303", "#040404"]
    );
    assert_eq!(
        drawing.ops[1].path,
        vec![
            PathCommand::Move { x: 0.0, y: 0.25 },
            PathCommand::Line { x: 1.0, y: 0.25 },
            PathCommand::Line { x: 1.0, y: 1.0 },
            PathCommand::Line { x: 0.0, y: 1.0 },
            PathCommand::Close,
        ]
    );
}

#[test]
fn a_gouraud_triangle_bands_across_its_colour_axis() {
    let bytes = emf(
        vec![gradient(
            GRADIENT_TRIANGLE,
            &[(0, 0, 0), (100, 0, 0), (0, 100, 4), (100, 100, 4)],
            &[0, 1, 2, 2, 1, 3],
        )],
        [0, 0, 99, 99],
        [100, 100],
    );

    let drawing = decode(&bytes).expect("the gradient decodes");
    assert_eq!(drawing.ops.len(), 8);
    assert_eq!(
        drawing.ops[..4]
            .iter()
            .map(|op| op.fill.clone().unwrap())
            .collect::<Vec<_>>(),
        ["#010101", "#020202", "#030303", "#040404"]
    );
    assert_eq!(
        drawing.ops[0].path,
        vec![
            PathCommand::Move { x: 0.0, y: 0.0 },
            PathCommand::Line { x: 1.0, y: 0.0 },
            PathCommand::Line { x: 0.0, y: 1.0 },
            PathCommand::Close,
        ]
    );
    assert_eq!(
        drawing.ops[3].path,
        vec![
            PathCommand::Move { x: 0.0, y: 0.75 },
            PathCommand::Line { x: 0.25, y: 0.75 },
            PathCommand::Line { x: 0.0, y: 1.0 },
            PathCommand::Close,
        ]
    );
}

#[test]
fn a_gradient_rejects_unreadable_counts_indexes_and_modes() {
    let mut oversized = gradient(GRADIENT_RECT_V, &[(0, 0, 0), (100, 100, 4)], &[0, 1]);
    oversized[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
    for records in [
        vec![oversized],
        vec![gradient(
            GRADIENT_RECT_V,
            &[(0, 0, 0), (100, 100, 4)],
            &[0, 7],
        )],
        vec![gradient(3, &[(0, 0, 0), (100, 100, 4)], &[0, 1])],
        vec![
            record(59, &[]),
            gradient(GRADIENT_RECT_V, &[(0, 0, 0), (100, 100, 4)], &[0, 1]),
        ],
    ] {
        let mut records = records;
        records.push(record(43, &i32s(&[0, 0, 50, 50])));
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);
        assert!(decode(&bytes).is_none());
    }
}

#[test]
fn a_gradient_declaring_more_vertices_or_shapes_than_the_cap_is_rejected_whole() {
    let over = MAX_GRADIENT_VERTICES + 1;
    let mut vertices = vec![(0, 0, 0), (100, 100, 4)];
    vertices.resize(over, (0, 0, 0));
    let mut indexes = vec![0u32, 1];
    indexes.resize(over * 2, 0);
    for records in [
        vec![gradient(GRADIENT_RECT_V, &vertices, &[0, 1])],
        vec![gradient(
            GRADIENT_RECT_V,
            &[(0, 0, 0), (100, 100, 4)],
            &indexes,
        )],
    ] {
        let mut records = records;
        records.push(record(43, &i32s(&[0, 0, 50, 50])));
        let bytes = emf(records, [0, 0, 99, 99], [100, 100]);
        assert!(decode(&bytes).is_none());
    }
}

#[test]
fn replay_budgets_reject_excessive_records_operations_and_points() {
    let bytes = emf(
        vec![record(43, &i32s(&[0, 0, 100, 100])); 4097],
        [0, 0, 99, 99],
        [100, 100],
    );
    assert!(decode(&bytes).is_none());
    let points = vec![(0, 0); 65536];
    let bytes = emf(
        vec![record(86, &poly16(&points)); 7],
        [0, 0, 99, 99],
        [100, 100],
    );
    assert!(decode(&bytes).is_none());
    let mut records = vec![record(18, &i32s(&[1])); 200000];
    records.push(record(43, &i32s(&[0, 0, 100, 100])));
    let bytes = emf(records, [0, 0, 99, 99], [100, 100]);
    assert!(decode(&bytes).is_none());
}
