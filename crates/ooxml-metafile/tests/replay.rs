mod common;

use common::*;
use ooxml_metafile::drawing::{Hatch, Op, Paint, PathCommand, Pixels, Rgba, Text, TextAnchor};
use ooxml_metafile::replay;

fn texts(ops: &[Op]) -> Vec<&Text> {
    ops.iter()
        .filter_map(|op| match op {
            Op::Text(text) => Some(text),
            _ => None,
        })
        .collect()
}

fn close(a: f64, b: f64) -> bool {
    (a - b).abs() < 1e-3
}

/// Sizes come from frames rounded to 0.01 mm, a few thousandths of a pixel off.
fn about(a: f64, b: f64) -> bool {
    (a - b).abs() < 0.05
}

fn point(command: &PathCommand) -> (f64, f64) {
    match *command {
        PathCommand::Move { x, y } | PathCommand::Line { x, y } => (x, y),
        PathCommand::Quad { x, y, .. } | PathCommand::Cubic { x, y, .. } => (x, y),
        PathCommand::Close => panic!("a close has no point"),
    }
}

fn at(command: &PathCommand, x: f64, y: f64) -> bool {
    let (px, py) = point(command);
    close(px, x) && close(py, y)
}

#[test]
fn a_frame_sizes_the_drawing_in_css_pixels() {
    let drawing = replay(&Emf::new(480, 320).rec(43, &i32s(&[0, 0, 10, 10])).bytes()).unwrap();
    assert!(about(drawing.width, 480.0) && about(drawing.height, 320.0));
}

#[test]
fn positioned_text_keeps_its_advances_font_and_top_alignment() {
    let bytes = Emf::new(200, 100)
        .recs(vec![
            font(1, -20, 700, [1, 0, 0], 0, 0, "Arial"),
            select(1),
            value(24, rgb(0x112233)),
            text_out(10, 30, "Ab", Some(&[12, 9]), 0, [0, 0, -1, -1]),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [text] = texts(&drawing.ops)[..] else {
        panic!("one text run");
    };
    assert_eq!(text.text, "Ab");
    assert_eq!(
        text.positions.as_deref(),
        Some(&[(0.0, 0.0), (12.0, 0.0)][..])
    );
    assert_eq!(
        (
            text.font.family.as_str(),
            text.font.weight,
            text.font.italic
        ),
        ("Arial", 700, true)
    );
    assert!(close(text.font.size, 20.0));
    assert_eq!(text.fill, Paint::Solid(Rgba::opaque(0x11, 0x22, 0x33)));
    assert!(close(text.transform[4], 10.0));
    assert!(
        close(text.transform[5], 30.0 + 0.905 * 20.0),
        "top-aligned runs drop to their baseline"
    );
}

#[test]
fn centred_right_and_rotated_runs_move_their_origin() {
    let bytes = Emf::new(200, 100)
        .recs(vec![
            font(1, -10, 400, [0, 0, 0], 900, 0, "Arial"),
            select(1),
            value(22, 24 | 6),
            text_out(100, 50, "abcd", Some(&[10, 10, 10, 10]), 0, [0, 0, -1, -1]),
            value(22, 2),
            text_out(100, 50, "abcd", None, 0, [0, 0, -1, -1]),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [centred, right] = texts(&drawing.ops)[..] else {
        panic!("two text runs");
    };
    assert!(
        close(centred.transform[0], 0.0) && close(centred.transform[1], -1.0),
        "the baseline runs up"
    );
    assert!(close(centred.transform[4], 100.0) && close(centred.transform[5], 70.0));
    assert_eq!(
        (right.anchor, right.positions.is_none()),
        (TextAnchor::End, true)
    );
}

#[test]
fn symbol_text_becomes_the_greek_it_depicts() {
    let bytes = Emf::new(100, 100)
        .recs(vec![
            font(1, -10, 400, [0, 0, 0], 0, 2, "Symbol"),
            select(1),
            text_out(0, 0, "abW\u{F070}", None, 0, [0, 0, -1, -1]),
        ])
        .bytes();
    assert_eq!(texts(&replay(&bytes).unwrap().ops)[0].text, "αβΩπ");
}

#[test]
fn glyph_index_text_is_omitted_and_reported() {
    let bytes = Emf::new(100, 100)
        .rec(43, &i32s(&[0, 0, 10, 10]))
        .recs(vec![text_out(
            0,
            0,
            "\u{12}\u{34}",
            None,
            0x10,
            [0, 0, -1, -1],
        )])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(texts(&drawing.ops).is_empty());
    assert_eq!(drawing.omissions[0].what, "text given as glyph indexes");
}

#[test]
fn opaque_text_fills_its_rectangle_and_clipped_text_clips_to_it() {
    let bytes = Emf::new(100, 100)
        .recs(vec![
            value(25, rgb(0xffff00)),
            text_out(10, 10, "x", Some(&[5]), 0x6, [5, 5, 50, 20]),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let Op::Shape(background) = &drawing.ops[0] else {
        panic!("the opaque rectangle paints first");
    };
    assert_eq!(
        background.fill,
        Some(Paint::Solid(Rgba::opaque(255, 255, 0)))
    );
    let Op::Text(text) = &drawing.ops[1] else {
        panic!("then the text");
    };
    let clip = text
        .clip
        .as_deref()
        .expect("clipped text carries its rectangle");
    let path = &clip.region.path;
    assert_eq!(path.len(), 5);
    assert!(at(&path[0], 5.0, 5.0) && at(&path[1], 50.0, 5.0));
    assert!(at(&path[2], 50.0, 20.0) && at(&path[3], 5.0, 20.0));
    assert_eq!(path[4], PathCommand::Close);
}

#[test]
fn clip_rectangles_intersect_exclude_and_come_back_with_their_context() {
    let bytes = Emf::new(100, 100)
        .recs(vec![
            bare(33),
            intersect_clip(10, 10, 90, 90),
            intersect_clip(0, 20, 80, 100),
            exclude_clip(40, 40, 50, 50),
            rect(0, 0, 100, 100),
            (34, i32s(&[-1])),
            rect(0, 0, 100, 100),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Shape(inside), Op::Shape(outside)] = &drawing.ops[..] else {
        panic!("two rectangles");
    };
    let exclude = inside.clip.as_deref().unwrap();
    assert!(exclude.region.exclude);
    let rect = exclude.parent.as_deref().unwrap();
    assert!(
        rect.parent.is_none(),
        "nested intersections merge into one rectangle"
    );
    assert!(at(&rect.region.path[0], 10.0, 20.0));
    assert!(at(&rect.region.path[2], 80.0, 90.0));
    assert!(outside.clip.is_none());
}

#[test]
fn a_meta_region_survives_a_clip_reset() {
    let bytes = Emf::new(100, 100)
        .recs(vec![
            intersect_clip(10, 10, 50, 50),
            bare(28),
            select_clip_region(5, &[]),
            rect(0, 0, 100, 100),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let Op::Shape(shape) = &drawing.ops[0] else {
        panic!("a rectangle");
    };
    assert!(
        shape.clip.is_some(),
        "resetting the clip leaves the meta region"
    );
}

#[test]
fn dashed_cosmetic_and_hatched_objects_keep_their_styles() {
    let bytes = Emf::new(100, 100)
        .recs(vec![
            pen(1, 0x1, 4, 0),
            select(1),
            brush(2, 2, rgb(0x00ff00), 4),
            select(2),
            rect(0, 0, 50, 50),
            (95, {
                let mut body = u32s(&[3, 0, 0, 0, 0]);
                body.extend(u32s(&[0x2, 7, 0, 0, 0, 0]));
                body
            }),
            select(3),
            move_to(0, 60),
            line_to(90, 60),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Shape(boxed), Op::Shape(line)] = &drawing.ops[..] else {
        panic!("a rectangle and a line");
    };
    let stroke = boxed.stroke.as_ref().unwrap();
    assert!(close(stroke.width, 4.0));
    let dash = stroke.dash.as_deref().unwrap();
    assert!(dash.len() == 2 && close(dash[0], 12.0) && close(dash[1], 4.0));
    assert!(matches!(
        boxed.fill,
        Some(Paint::Hatch {
            style: Hatch::Cross,
            ..
        })
    ));
    let hairline = line.stroke.as_ref().unwrap();
    assert!(
        close(hairline.width, 1.0),
        "a cosmetic pen is one pixel wide"
    );
    let dash = hairline.dash.as_deref().unwrap();
    assert!(dash.len() == 2 && close(dash[0], 3.0) && close(dash[1], 3.0));
}

#[test]
fn a_stretched_dib_lands_top_row_first_in_its_destination() {
    let dib = dib24(&[
        &[rgb(0xff0000), rgb(0x00ff00)],
        &[rgb(0x0000ff), rgb(0xffffff)],
    ]);
    let bytes = Emf::new(100, 100)
        .recs(vec![stretch_dibits([10, 20, 40, 60], &dib, 0x00CC_0020)])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let Op::Image(image) = &drawing.ops[0] else {
        panic!("an image");
    };
    let Pixels::Rgba(pixels) = &image.bitmap.pixels else {
        panic!("decoded pixels");
    };
    assert_eq!(pixels[..4], [255, 0, 0, 255]);
    assert_eq!(pixels[8..12], [0, 0, 255, 255]);
    assert!(close(image.transform[0], 20.0) && close(image.transform[3], 30.0));
    assert!(close(image.transform[4], 10.0) && close(image.transform[5], 20.0));
}

#[test]
fn and_masks_clear_their_white_pixels() {
    let dib = dib24(&[&[0x00ff_ffff, 0]]);
    let bytes = Emf::new(100, 100)
        .recs(vec![stretch_dibits([0, 0, 20, 10], &dib, 0x0088_00C6)])
        .bytes();
    let Op::Image(image) = &replay(&bytes).unwrap().ops[0] else {
        panic!("an image");
    };
    let Pixels::Rgba(pixels) = &image.bitmap.pixels else {
        panic!("decoded pixels");
    };
    assert_eq!((pixels[3], pixels[7]), (0, 255));
}

#[test]
fn transparent_and_alpha_blits_keep_their_transparency() {
    let dib = dib24(&[&[rgb(0x00ff00), rgb(0xff0000)]]);
    let bytes = Emf::new(100, 100)
        .recs(vec![
            blend(116, [0, 0, 20, 10], &dib, rgb(0x00ff00)),
            blend(114, [0, 0, 20, 10], &dib, 0x0080_0000),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Image(keyed), Op::Image(blended)] = &drawing.ops[..] else {
        panic!("two images");
    };
    let Pixels::Rgba(pixels) = &keyed.bitmap.pixels else {
        panic!("decoded pixels");
    };
    assert_eq!((pixels[3], pixels[7]), (0, 255));
    assert!(close(blended.opacity, 128.0 / 255.0));
}

#[test]
fn emf_plus_only_records_draw_with_their_alpha() {
    let bytes = Emf::new(100, 100)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                plus_fill_rects(0x80ff_0000, &[[0.0, 0.0, 10.0, 10.0]]),
                plus_eof(),
            ])
            .1,
        )
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let Op::Shape(shape) = &drawing.ops[0] else {
        panic!("a fill");
    };
    assert_eq!(
        shape.fill,
        Some(Paint::Solid(Rgba {
            r: 255,
            g: 0,
            b: 0,
            a: 0x80
        }))
    );
}

fn dual(plus_records: Vec<(u16, u16, Vec<u8>)>) -> Vec<u8> {
    let mut records = vec![plus_header(true)];
    records.extend(plus_records);
    Emf::new(100, 100)
        .rec(70, &plus(&records).1)
        .recs(vec![
            brush(1, 0, rgb(0x0000ff), 0),
            select(1),
            rect(0, 0, 50, 50),
        ])
        .bytes()
}

#[test]
fn a_dual_metafile_prefers_emf_plus_and_falls_back_to_gdi() {
    let clean = replay(&dual(vec![plus_fill_rects(
        0xff00_ff00,
        &[[0.0, 0.0, 10.0, 10.0]],
    )]))
    .unwrap();
    let [Op::Shape(shape)] = &clean.ops[..] else {
        panic!("only the EMF+ fill");
    };
    assert_eq!(shape.fill, Some(Paint::Solid(Rgba::opaque(0, 255, 0))));
    let lossy = replay(&dual(vec![(0x4037, 0, Vec::new())])).unwrap();
    let [Op::Shape(shape)] = &lossy.ops[..] else {
        panic!("only the GDI rectangle");
    };
    assert_eq!(shape.fill, Some(Paint::Solid(Rgba::opaque(0, 0, 255))));
    assert!(
        lossy.omissions.is_empty(),
        "the GDI rendition drew everything"
    );
}

#[test]
fn emf_plus_only_files_play_gdi_records_only_inside_get_dc_spans() {
    let bytes = Emf::new(100, 100)
        .rec(70, &plus(&[plus_header(false)]).1)
        .recs(vec![rect(0, 0, 10, 10)])
        .rec(70, &plus(&[(0x4004, 0, Vec::new())]).1)
        .recs(vec![rect(20, 20, 30, 30)])
        .rec(70, &plus(&[plus_eof()]).1)
        .recs(vec![rect(40, 40, 50, 50)])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Shape(shape)] = &drawing.ops[..] else {
        panic!("only the rectangle inside the GetDC span");
    };
    assert!(at(&shape.path[0], 20.0, 20.0));
}

#[test]
fn a_wmf_draws_text_and_bitmaps_in_its_placeable_frame() {
    let mut font = i16s(&[-100, 0, 0, 0, 400]);
    font.extend([0; 8]);
    font.extend(b"Arial\0");
    font.resize(50, 0);
    let (bmi, bits) = dib24(&[&[rgb(0xff0000)]]);
    let mut blit = u32s(&[0x00CC_0020]);
    blit.extend(i16s(&[1, 1, 0, 0, 100, 100, 0, 0]));
    blit.extend(bmi);
    blit.extend(bits);
    let bytes = Wmf::new(1440, 720, 1440)
        .recs(vec![
            (0x02FB, font),
            (0x012D, u16s(&[0])),
            (
                0x0521,
                [i16s(&[2]), b"Hi".to_vec(), i16s(&[360, 720])].concat(),
            ),
            (0x0B41, blit),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(close(drawing.width, 96.0) && close(drawing.height, 48.0));
    let text = texts(&drawing.ops)[0];
    assert_eq!(text.text, "Hi");
    assert!(close(text.font.size * text.transform[3], 100.0 / 15.0));
    assert!(drawing.ops.iter().any(|op| matches!(op, Op::Image(_))));
}

#[test]
fn a_wmf_carrying_an_emf_draws_the_emf() {
    let emf = Emf::new(10, 10).recs(vec![rect(0, 0, 5, 5)]).bytes();
    let mut escape = u16s(&[0x000F, (34 + emf.len()) as u16]);
    escape.extend(u32s(&[0x4346_4D57, 1, 0x0001_0000]));
    escape.extend(u16s(&[0]));
    escape.extend(u32s(&[0, 1, emf.len() as u32, 0, emf.len() as u32]));
    escape.extend(&emf);
    let bytes = Wmf::new(1440, 1440, 1440)
        .recs(vec![(0x0626, escape), (0x041B, i16s(&[10, 10, 0, 0]))])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    assert!(
        about(drawing.width, 10.0),
        "the embedded EMF's frame sizes the drawing"
    );
}

#[test]
fn unknown_records_and_exhausted_limits_refuse_the_metafile() {
    assert!(replay(&Emf::new(10, 10).rec(250, &[]).bytes()).is_err());
    let unknown_plus = Emf::new(10, 10)
        .rec(70, &plus(&[plus_header(false), (0x40ff, 0, Vec::new())]).1)
        .bytes();
    assert!(replay(&unknown_plus).is_err());
    let nested: Vec<(u32, Vec<u8>)> = (0..200)
        .map(|index| intersect_clip(0, index, 100, 100))
        .collect();
    let mut deep = Vec::new();
    for (index, clip) in nested.into_iter().enumerate() {
        deep.push(exclude_clip(0, index as i32, 1, 1 + index as i32));
        deep.push(clip);
    }
    assert!(replay(&Emf::new(10, 10).recs(deep).bytes()).is_err());
    let text = "x".repeat(1_000_001);
    assert!(
        replay(
            &Emf::new(10, 10)
                .recs(vec![text_out(0, 0, &text, None, 0, [0, 0, -1, -1])])
                .bytes()
        )
        .is_err()
    );
}

#[test]
fn a_nested_emf_plus_metafile_is_placed_in_its_destination() {
    let inner = Emf::new(10, 10)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                plus_fill_rects(0xff00_0000, &[[0.0, 0.0, 10.0, 10.0]]),
                plus_eof(),
            ])
            .1,
        )
        .bytes();
    let mut image = u32s(&[0xDBC0_1002, 2, 3, inner.len() as u32]);
    image.extend(&inner);
    let bytes = Emf::new(100, 100)
        .rec(
            70,
            &plus(&[
                plus_header(false),
                (0x4008, 0x0500, image),
                plus_draw_image(0, [0.0, 0.0, 10.0, 10.0], [50.0, 50.0, 20.0, 20.0]),
                plus_eof(),
            ])
            .1,
        )
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let Op::Shape(shape) = &drawing.ops[0] else {
        panic!("the nested fill");
    };
    assert!(at(&shape.path[0], 50.0, 50.0));
    let (x, y) = point(&shape.path[2]);
    assert!(about(x, 70.0) && about(y, 70.0));
}

#[test]
fn an_emf_plus_header_after_another_comment_still_counts() {
    let bytes = Emf::new(100, 100)
        .rec(70, &u32s(&[4, 0x4344_4947]))
        .rec(
            70,
            &plus(&[
                plus_header(true),
                plus_fill_rects(0xff00_ff00, &[[0.0, 0.0, 10.0, 10.0]]),
                plus_eof(),
            ])
            .1,
        )
        .recs(vec![
            bare(33),
            (37, u32s(&[0x8000_0004])),
            rect(0, 0, 50, 50),
            (34, i32s(&[-1])),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Shape(shape)] = &drawing.ops[..] else {
        panic!("only the EMF+ fill");
    };
    assert_eq!(shape.fill, Some(Paint::Solid(Rgba::opaque(0, 255, 0))));
}

#[test]
fn a_wmf_region_fills_its_scan_rectangles() {
    let mut region = u16s(&[0, 6]);
    region.extend(u32s(&[0]));
    region.extend(i16s(&[0, 1, 0, 10, 20, 60, 40]));
    region.extend(u16s(&[4]));
    region.extend(i16s(&[20, 40, 10, 30, 50, 60]));
    region.extend(u16s(&[4]));
    let bytes = Wmf::new(100, 100, 96)
        .recs(vec![
            (
                0x02FC,
                [u16s(&[0]), u32s(&[rgb(0x00ff00)]), u16s(&[0])].concat(),
            ),
            (0x06FF, region),
            (0x0228, u16s(&[1, 0])),
        ])
        .bytes();
    let drawing = replay(&bytes).unwrap();
    let [Op::Shape(shape)] = &drawing.ops[..] else {
        panic!("one region fill");
    };
    assert_eq!(shape.fill, Some(Paint::Solid(Rgba::opaque(0, 255, 0))));
    assert!(at(&shape.path[0], 10.0, 20.0) && at(&shape.path[2], 30.0, 40.0));
    assert!(at(&shape.path[5], 50.0, 20.0) && at(&shape.path[7], 60.0, 40.0));
}
