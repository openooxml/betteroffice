//! The synthetic metafiles the golden rasters and Word references draw.

use super::*;

const RED: u32 = rgb(0xd93025);
const BLUE: u32 = rgb(0x1a73e8);
const GREEN: u32 = rgb(0x188038);
const YELLOW: u32 = rgb(0xf9ab00);
const GREY: u32 = rgb(0x5f6368);
const BLACK: u32 = 0;
const WHITE: u32 = 0x00ff_ffff;

/// Every fixture, by file name.
pub fn all() -> Vec<(&'static str, Vec<u8>)> {
    vec![
        ("shapes.emf", shapes()),
        ("text.emf", text()),
        ("clip-bitmap.emf", clip_bitmap()),
        ("emfplus.emf", emfplus()),
        ("shapes.wmf", wmf()),
    ]
}

/// Lines, polygons, rectangles, ellipses, arcs, Béziers, pens and brushes.
pub fn shapes() -> Vec<u8> {
    Emf::new(480, 320)
        .recs(vec![
            value(17, 1),
            value(19, 2),
            pen(1, 0x2200, 6, BLACK),
            brush(2, 0, BLUE, 0),
            select(1),
            select(2),
            rect(20, 20, 140, 100),
            brush(3, 0, RED, 0),
            select(3),
            ellipse(170, 20, 290, 100),
            brush(4, 2, GREEN, 5),
            select(4),
            round_rect(320, 20, 460, 100, 40, 40),
            pen(5, 0x1, 3, GREY),
            select(5),
            stock(5),
            poly16(
                86,
                &[(40, 140), (120, 130), (140, 210), (60, 230), (20, 180)],
            ),
            pen(6, 0x2, 3, RED),
            select(6),
            poly16(87, &[(170, 140), (210, 220), (250, 140), (290, 220)]),
            pen(7, 0x200, 10, BLUE),
            select(7),
            move_to(320, 150),
            line_to(450, 150),
            user_pen(8, 4, GREEN, &[12, 4, 2, 4]),
            select(8),
            move_to(320, 190),
            line_to(450, 190),
            select(1),
            brush(9, 0, YELLOW, 0),
            select(9),
            arc(47, [20, 240, 120, 310], (120, 275), (70, 240)),
            arc(46, [140, 240, 240, 310], (240, 275), (140, 275)),
            stock(5),
            arc(45, [260, 240, 360, 310], (360, 275), (260, 275)),
            bare(59),
            move_to(380, 300),
            poly16(88, &[(390, 230), (450, 230), (460, 300)]),
            bare(61),
            bare(60),
            select(9),
            bounds_only(63),
            (36, {
                let mut body = f32s(&[0.866, 0.5, -0.5, 0.866, 400.0, 110.0]);
                body.extend(u32s(&[4]));
                body
            }),
            brush(10, 2, BLUE, 4),
            select(10),
            (18, u32s(&[1])),
            rect(-20, -8, 20, 8),
        ])
        .bytes()
}

/// Text runs: advances, alignment, rotation, styles, backgrounds and Symbol.
/// The advances are the Liberation substitutes' own, rounded as GDI rounds.
pub fn text() -> Vec<u8> {
    Emf::new(480, 320)
        .recs(vec![
            value(18, 1),
            value(24, RED),
            font(1, -28, 400, [0, 0, 0], 0, 0, "Arial"),
            select(1),
            text_out(
                20,
                20,
                "Left top",
                Some(&[16, 16, 8, 8, 8, 8, 16, 16]),
                0,
                [0, 0, -1, -1],
            ),
            value(22, 24 | 6),
            value(24, BLUE),
            font(2, -24, 700, [1, 1, 0], 0, 0, "Times New Roman"),
            select(2),
            text_out(
                240,
                110,
                "Centred baseline",
                Some(&[16, 11, 13, 7, 9, 11, 12, 6, 12, 12, 9, 11, 7, 7, 13, 11]),
                0,
                [0, 0, -1, -1],
            ),
            value(22, 8 | 2),
            value(24, GREEN),
            font(3, -20, 400, [0, 0, 1], 0, 0, "Courier New"),
            select(3),
            text_out(460, 160, "Right bottom", None, 0, [0, 0, -1, -1]),
            value(22, 0),
            value(24, BLACK),
            font(4, -22, 400, [0, 0, 0], 900, 0, "Arial"),
            select(4),
            text_out(
                40,
                300,
                "Rotated",
                Some(&[16, 12, 6, 12, 6, 12, 12]),
                0,
                [0, 0, -1, -1],
            ),
            font(5, -26, 400, [0, 0, 0], 0, 2, "Symbol"),
            select(5),
            text_out(
                100,
                190,
                "abgWp",
                Some(&[14, 13, 11, 19, 13]),
                0,
                [0, 0, -1, -1],
            ),
            select(1),
            value(25, YELLOW),
            text_out(
                100,
                240,
                "Opaque",
                Some(&[22, 16, 16, 16, 16, 16]),
                0x2,
                [96, 236, 210, 272],
            ),
            text_out(
                260,
                240,
                "Clipped away",
                Some(&[20, 6, 6, 16, 16, 16, 16, 8, 16, 20, 16, 14]),
                0x4,
                [256, 236, 330, 272],
            ),
        ])
        .bytes()
}

/// Bitmaps, raster operations and clipping.
pub fn clip_bitmap() -> Vec<u8> {
    let checker: Vec<Vec<u32>> = (0..8)
        .map(|y| {
            (0..8)
                .map(|x| if (x + y) % 2 == 0 { RED } else { WHITE })
                .collect()
        })
        .collect();
    let rows: Vec<&[u32]> = checker.iter().map(Vec::as_slice).collect();
    let tiles = dib24(&rows);
    let gradient: Vec<Vec<u32>> = (0..16)
        .map(|y| {
            (0..16)
                .map(|x| (x * 16) | ((y * 16) << 8) | 0x80_0000)
                .collect()
        })
        .collect();
    let gradient_rows: Vec<&[u32]> = gradient.iter().map(Vec::as_slice).collect();
    let smooth = dib24(&gradient_rows);
    let mask = dib1(
        &[
            &[1, 1, 0, 0, 0, 0, 1, 1],
            &[1, 0, 0, 0, 0, 0, 0, 1],
            &[0, 0, 0, 0, 0, 0, 0, 0],
            &[0, 0, 0, 0, 0, 0, 0, 0],
            &[0, 0, 0, 0, 0, 0, 0, 0],
            &[0, 0, 0, 0, 0, 0, 0, 0],
            &[1, 0, 0, 0, 0, 0, 0, 1],
            &[1, 1, 0, 0, 0, 0, 1, 1],
        ],
        [BLACK, WHITE],
    );
    let keyed: Vec<Vec<u32>> = (0..8)
        .map(|y| {
            (0..8)
                .map(|x| if x == y || x == 7 - y { BLUE } else { GREEN })
                .collect()
        })
        .collect();
    let keyed_rows: Vec<&[u32]> = keyed.iter().map(Vec::as_slice).collect();
    let keyed = dib24(&keyed_rows);
    Emf::new(480, 320)
        .recs(vec![
            stretch_dibits([20, 20, 120, 120], &tiles, 0x00CC_0020),
            stretch_dibits([160, 20, 120, 120], &smooth, 0x00CC_0020),
            brush(1, 0, YELLOW, 0),
            select(1),
            stock(8),
            rect(300, 20, 460, 140),
            stretch_dibits([320, 40, 120, 80], &mask, 0x0088_00C6),
            bare(33),
            intersect_clip(20, 170, 200, 300),
            exclude_clip(70, 200, 150, 270),
            brush(2, 0, BLUE, 0),
            select(2),
            ellipse(0, 150, 220, 320),
            (34, i32s(&[-1])),
            bare(33),
            bare(59),
            ellipse(240, 170, 360, 300),
            bare(60),
            value(67, 5),
            stretch_dibits([240, 170, 120, 130], &tiles, 0x00CC_0020),
            (34, i32s(&[-1])),
            select_clip_region(5, &[[380, 170, 470, 200], [380, 230, 470, 260]]),
            brush(3, 0, GREEN, 0),
            select(3),
            rect(370, 160, 480, 310),
            select_clip_region(5, &[]),
            blend(114, [380, 270, 40, 40], &smooth, 0x0080_0000),
            blend(116, [430, 270, 40, 40], &keyed, GREEN),
        ])
        .bytes()
}

/// An EMF+-only metafile: fills with alpha, dashed pens, a gradient path,
/// a clip, a bitmap and positioned text.
pub fn emfplus() -> Vec<u8> {
    let bitmap: Vec<Vec<u32>> = (0..8)
        .map(|y| {
            (0..8)
                .map(|x| {
                    if (x / 2 + y / 2) % 2 == 0 {
                        0xff1a_73e8
                    } else {
                        0x40f9_ab00
                    }
                })
                .collect()
        })
        .collect();
    let bitmap_rows: Vec<&[u32]> = bitmap.iter().map(Vec::as_slice).collect();
    Emf::new(480, 320)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                plus_fill_rects(0xff18_8038, &[[20.0, 20.0, 140.0, 90.0]]),
                plus_fill_ellipse(0x80d9_3025, [100.0, 50.0, 140.0, 90.0]),
                plus_pen(1, 4.0, 0xff20_2124, &[3.0, 1.0]),
                plus_draw_lines(
                    1,
                    &[(270.0, 30.0), (460.0, 30.0), (460.0, 130.0), (270.0, 130.0)],
                    true,
                ),
                plus_path(2, &[(20.0, 300.0), (120.0, 170.0), (220.0, 300.0)]),
                plus_linear_brush(3, [20.0, 170.0, 200.0, 130.0], 0xff1a_73e8, 0xfff9_ab00),
                plus_fill_path(2, 3),
                plus_pen(4, 2.0, 0xff20_2124, &[]),
                plus_draw_path(2, 4),
                plus_clip_rect(0, [260.0, 170.0, 100.0, 60.0]),
                plus_fill_ellipse(0xff5f_6368, [240.0, 150.0, 140.0, 100.0]),
                (0x4031, 0, Vec::new()),
                plus_bitmap(5, &bitmap_rows),
                plus_draw_image(5, [0.0, 0.0, 8.0, 8.0], [400.0, 170.0, 64.0, 64.0]),
                plus_font(6, 24.0, 1, "Arial"),
                plus_world([1.0, 0.0, 0.0, 1.0, 0.0, 10.0]),
                plus_driver_string(
                    6,
                    0xff20_2124,
                    "EMF+ text",
                    (260.0, 280.0),
                    &[16.0, 20.0, 15.0, 14.0, 7.0, 8.0, 13.0, 13.0, 8.0],
                ),
                plus_eof(),
            ])
            .1,
        )
        .bytes()
}

/// A placeable WMF: shapes, pens, brushes, text and a bitmap.
pub fn wmf() -> Vec<u8> {
    let checker: Vec<Vec<u32>> = (0..4)
        .map(|y| {
            (0..4)
                .map(|x| if (x + y) % 2 == 0 { BLUE } else { YELLOW })
                .collect()
        })
        .collect();
    let rows: Vec<&[u32]> = checker.iter().map(Vec::as_slice).collect();
    let (bmi, bits) = dib24(&rows);
    let mut stretch = u32s(&[0x00CC_0020]);
    stretch.extend(i16s(&[4, 4, 0, 0, 160, 160, 520, 1120]));
    stretch.extend(bmi);
    stretch.extend(bits);
    let mut font = i16s(&[-240, 0, 0, 0, 700]);
    font.extend([0, 0, 0, 0, 0, 0, 0, 0]);
    font.extend(b"Arial\0");
    font.resize(50, 0);
    let text = b"WMF text";
    let mut ext_text = i16s(&[1130, 140, text.len() as i16]);
    ext_text.extend(u16s(&[0]));
    ext_text.extend(text);
    ext_text.extend(i16s(&[227, 200, 147, 67, 80, 133, 133, 80]));
    let mut polygon = i16s(&[5]);
    polygon.extend(points_xy(&[
        (200, 640),
        (560, 600),
        (640, 1000),
        (280, 1080),
        (120, 840),
    ]));
    let mut polyline = i16s(&[4]);
    polyline.extend(points_xy(&[
        (760, 640),
        (940, 1000),
        (1120, 640),
        (1300, 1000),
    ]));
    Wmf::new(2160, 1440, 720)
        .recs(vec![
            (0x0103, i16s(&[8])),
            (0x020B, i16s(&[0, 0])),
            (0x020C, i16s(&[1440, 2160])),
            (0x0102, i16s(&[1])),
            (
                0x02FA,
                [u16s(&[0]), i16s(&[20, 0]), u32s(&[BLACK])].concat(),
            ),
            (0x02FC, [u16s(&[0]), u32s(&[GREEN]), u16s(&[0])].concat()),
            (0x012D, u16s(&[0])),
            (0x012D, u16s(&[1])),
            (0x041B, i16s(&[520, 700, 100, 100])),
            (0x02FC, [u16s(&[2]), u32s(&[RED]), u16s(&[3])].concat()),
            (0x012D, u16s(&[2])),
            (0x0418, i16s(&[520, 1400, 100, 800])),
            (0x02FC, [u16s(&[0]), u32s(&[YELLOW]), u16s(&[0])].concat()),
            (0x012D, u16s(&[3])),
            (0x061C, i16s(&[120, 120, 520, 2060, 100, 1500])),
            (0x0324, polygon),
            (0x02FA, [u16s(&[1]), i16s(&[10, 0]), u32s(&[BLUE])].concat()),
            (0x012D, u16s(&[4])),
            (0x0325, polyline),
            (
                0x081A,
                i16s(&[1400, 1800, 1100, 2100, 1400, 2100, 1100, 1500]),
            ),
            (0x02FB, font),
            (0x012D, u16s(&[5])),
            (0x0209, u32s(&[BLUE])),
            (0x0A32, ext_text),
            (0x0209, u32s(&[BLACK])),
            (
                0x0521,
                [i16s(&[4]), b"Plan".to_vec(), i16s(&[640, 1560])].concat(),
            ),
            (0x0B41, stretch),
        ])
        .bytes()
}
