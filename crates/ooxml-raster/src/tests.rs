use super::*;
use std::io::Cursor;

#[test]
fn png_encodes_demultiplied_rgba() {
    let mut pixmap = Pixmap::new(2, 1).unwrap();
    pixmap.fill(Color::from_rgba8(20, 100, 200, 128));
    let expected = pixmap.clone().take_demultiplied();
    let encoded = encode_png(pixmap, 2, 1).unwrap();
    let mut reader = png::Decoder::new(Cursor::new(encoded)).read_info().unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    assert_eq!((info.width, info.height), (2, 1));
    assert_eq!(info.color_type, png::ColorType::Rgba);
    assert_eq!(&pixels[..info.buffer_size()], expected.as_slice());
}

#[test]
fn png_rejects_invalid_dimensions() {
    assert!(encode_png(Pixmap::new(1, 1).unwrap(), 0, 1).is_err());
}

#[test]
fn hex_preserves_short_and_alpha_forms() {
    for (hex, rgba) in [
        ("abc", [170, 187, 204, 255]),
        ("aBcD", [170, 187, 204, 221]),
        ("123456", [18, 52, 86, 255]),
        ("12345678", [18, 52, 86, 120]),
        ("+f0000", [15, 0, 0, 255]),
    ] {
        assert_eq!(parse_hex_rgba(hex), Some(rgba));
        assert_eq!(
            parse_hex_color(hex),
            Some(Color::from_rgba8(rgba[0], rgba[1], rgba[2], rgba[3]))
        );
    }
    for hex in ["", "12", "12345", "ggg", "#fff", "ffffffgg", "é1234"] {
        assert_eq!(parse_hex_rgba(hex), None);
        assert_eq!(parse_hex_color(hex), None);
    }
}

#[test]
fn css_preserves_channel_rounding_and_grammar() {
    for (value, expected) in [
        ("#abc", Color::from_rgba8(170, 187, 204, 255)),
        ("#abcd", Color::from_rgba8(170, 187, 204, 221)),
        ("#12345678", Color::from_rgba8(18, 52, 86, 120)),
        ("TrAnSpArEnT", Color::TRANSPARENT),
        ("rgb(100%, 0%, 20%)", Color::from_rgba8(255, 0, 51, 255)),
        ("rgb(-1, 300, 1.5)", Color::from_rgba8(0, 255, 2, 255)),
        ("rgba(1, 2, 3, 0.5)", Color::from_rgba8(1, 2, 3, 128)),
        ("rgba(1, 2, 3, 20%)", Color::from_rgba8(1, 2, 3, 51)),
    ] {
        assert_eq!(parse_css_color(value), Ok(expected));
    }
    for value in [
        "red",
        " rgb(1,2,3)",
        "RGB(1,2,3)",
        "rgb(1 2 3)",
        "rgba(1,2,3)",
        "#ggg",
    ] {
        assert_eq!(parse_css_color(value), Err(format!("bad color: {value}")));
    }
}

#[test]
fn geometry_maps_every_point_in_command_order() {
    let commands = [
        GeometryPathCommand::Move { x: 0.1, y: 0.2 },
        GeometryPathCommand::Line { x: 0.3, y: 0.4 },
        GeometryPathCommand::Quad {
            cpx: 0.5,
            cpy: 0.6,
            x: 0.7,
            y: 0.8,
        },
        GeometryPathCommand::Cubic {
            cp1x: 0.9,
            cp1y: 1.0,
            cp2x: 1.1,
            cp2y: 1.2,
            x: 1.3,
            y: 1.4,
        },
        GeometryPathCommand::Close,
    ];
    let mut mapped = Vec::new();
    let path = path_from_commands(&commands, |x, y| {
        mapped.push((x, y));
        (7.0 + x as f32 * 31.0, 11.0 + y as f32 * 43.0)
    })
    .unwrap();
    assert_eq!(
        mapped,
        vec![
            (0.1, 0.2),
            (0.3, 0.4),
            (0.5, 0.6),
            (0.7, 0.8),
            (0.9, 1.0),
            (1.1, 1.2),
            (1.3, 1.4)
        ]
    );
    let expected = mapped
        .iter()
        .map(|&(x, y)| tiny_skia::Point::from_xy(7.0 + x as f32 * 31.0, 11.0 + y as f32 * 43.0))
        .collect::<Vec<_>>();
    assert_eq!(path.points(), expected.as_slice());
    assert_eq!(segment_kinds(&path), "MLQCZ");
}

#[test]
fn geometry_preserves_empty_and_non_finite_results() {
    let cast = |x, y| (x as f32, y as f32);
    assert!(path_from_commands(&[], cast).is_none());
    assert!(path_from_commands(&[GeometryPathCommand::Move { x: 0.0, y: 0.0 }], cast).is_none());
    for value in [f64::NAN, f64::INFINITY, f64::MAX] {
        let commands = [
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Line { x: value, y: 1.0 },
        ];
        assert!(path_from_commands(&commands, cast).is_none());
        let mut calls = 0;
        let result = try_path_from_commands(&commands, |x, y| {
            calls += 1;
            let (x, y) = (x as f32, y as f32);
            if x.is_finite() && y.is_finite() {
                Ok((x, y))
            } else {
                Err("non-finite")
            }
        });
        assert_eq!(result.unwrap_err(), "non-finite");
        assert_eq!(calls, 2);
    }
    assert!(
        try_path_from_commands(&[], |_, _| Err::<(f32, f32), _>("unused"))
            .unwrap()
            .is_none()
    );
}

#[test]
fn outline_preserves_all_commands() {
    let commands = [
        PathCmd::MoveTo { x: 1.0, y: 2.0 },
        PathCmd::LineTo { x: 3.0, y: 4.0 },
        PathCmd::QuadTo {
            cx: 5.0,
            cy: 6.0,
            x: 7.0,
            y: 8.0,
        },
        PathCmd::CubicTo {
            c1x: 9.0,
            c1y: 10.0,
            c2x: 11.0,
            c2y: 12.0,
            x: 13.0,
            y: 14.0,
        },
        PathCmd::Close,
    ];
    let path = outline_path(&commands).unwrap();
    let coordinates = path.points().iter().map(|p| (p.x, p.y)).collect::<Vec<_>>();
    assert_eq!(
        coordinates,
        vec![
            (1.0, 2.0),
            (3.0, 4.0),
            (5.0, 6.0),
            (7.0, 8.0),
            (9.0, 10.0),
            (11.0, 12.0),
            (13.0, 14.0)
        ]
    );
    assert!(segment_kinds(&path).ends_with('Z'));
}

#[test]
fn outline_preserves_empty_and_non_finite_results() {
    assert!(outline_path(&[]).is_none());
    assert!(outline_path(&[PathCmd::MoveTo { x: 0.0, y: 0.0 }]).is_none());
    assert!(
        outline_path(&[
            PathCmd::MoveTo { x: 0.0, y: 0.0 },
            PathCmd::LineTo {
                x: f32::NAN,
                y: 1.0
            },
        ])
        .is_none()
    );
}

#[test]
fn premultiply_preserves_alpha_and_trailing_bytes() {
    let mut pixels = [255, 128, 1, 128, 1, 2, 3, 0, 10, 20, 30, 255, 99];
    premultiply_rgba(&mut pixels);
    assert_eq!(pixels, [128, 64, 1, 128, 0, 0, 0, 0, 10, 20, 30, 255, 99]);
    premultiply_rgba(&mut []);
    let mut short = [1, 2, 3];
    premultiply_rgba(&mut short);
    assert_eq!(short, [1, 2, 3]);
}

#[test]
fn clip_preserves_intersection_and_non_finite_behavior() {
    let rect = |x, y, w, h| RasterRect { x, y, w, h };
    assert_eq!(
        clipped_to_target(rect(-2.0, 3.0, 10.0, 20.0), 5, 10),
        Some(rect(0.0, 3.0, 5.0, 7.0))
    );
    for clip in [
        rect(10.0, 0.0, 1.0, 1.0),
        rect(0.0, 0.0, 0.0, 1.0),
        rect(0.0, 0.0, -1.0, 1.0),
        rect(f32::INFINITY, 0.0, 1.0, 1.0),
    ] {
        assert_eq!(clipped_to_target(clip, 10, 10), None);
    }
    assert_eq!(clipped_to_target(rect(0.0, 0.0, 1.0, 1.0), 0, 10), None);
    assert_eq!(
        clipped_to_target(rect(f32::NAN, 0.0, 1.0, 1.0), 10, 10),
        Some(rect(0.0, 0.0, 10.0, 1.0))
    );
}

fn segment_kinds(path: &tiny_skia::Path) -> String {
    path.segments()
        .map(|segment| match segment {
            tiny_skia::PathSegment::MoveTo(_) => 'M',
            tiny_skia::PathSegment::LineTo(_) => 'L',
            tiny_skia::PathSegment::QuadTo(..) => 'Q',
            tiny_skia::PathSegment::CubicTo(..) => 'C',
            tiny_skia::PathSegment::Close => 'Z',
        })
        .collect()
}
