//! Image placement, decoding and part-scoped lookup.
//!
//! `image.rs` keeps its helpers `pub(crate)` and `lib.rs` is owned by the
//! integration slice, so the test crate mounts the real sources as modules.
//! That exercises the shipped code rather than a paraphrase of it.

#![allow(dead_code)]

use std::collections::{BTreeMap, HashMap};

use docx_layout::display_list::{ContentFrame, Crop, DisplayList, DocAttrs, ImagePrimitive};
use ooxml_text::FontStore;
use pdf_writer::{Content, Ref};
use serde_json::Number;

#[path = "../src/units.rs"]
mod units;

#[path = "../../betteroffice-docx/src/image_scope.rs"]
mod image_scope;

// The external `image` crate keeps its name: naming this module `image` would
// shadow it at the test-crate root.
#[path = "../src/image.rs"]
mod image_impl;

#[derive(Debug)]
pub(crate) enum PdfError {
    InvalidNumber(&'static str),
    Image(String),
}

pub(crate) struct ImageRefs {
    pub(crate) resource: String,
    pub(crate) id: Ref,
    pub(crate) smask: Option<Ref>,
}

fn number(value: f64) -> Number {
    Number::from_f64(value).expect("finite number")
}

fn crop(left: f64, top: f64, right: f64, bottom: f64) -> Crop {
    Crop {
        top: number(top),
        right: number(right),
        bottom: number(bottom),
        left: number(left),
    }
}

fn content_frame(
    x: Option<f64>,
    y: Option<f64>,
    w: Option<f64>,
    h: Option<f64>,
) -> Box<ContentFrame> {
    Box::new(ContentFrame {
        x: x.map(number),
        y: y.map(number),
        w: w.map(number),
        h: h.map(number),
    })
}

fn primitive(crop: Option<Crop>, rotation_deg: Option<f64>) -> ImagePrimitive {
    ImagePrimitive {
        rel_id: "rId5".to_owned(),
        x: number(96.0),
        y: number(48.0),
        w: number(128.0),
        h: number(64.0),
        rotation_deg: rotation_deg.map(number),
        opacity: None,
        filter: None,
        decorative: false,
        crop,
        alt_text: None,
        relationship_id: None,
        attrs: DocAttrs::default(),
    }
}

/// A 4x2 RGBA PNG: four distinct colour bands per row, one band translucent.
fn banded_png() -> Vec<u8> {
    let pixels: [u8; 4 * 2 * 4] = [
        // row 0
        255, 0, 0, 255, // red
        0, 255, 0, 255, // green
        0, 0, 255, 128, // blue, half transparent
        255, 255, 255, 255, // white
        // row 1
        10, 20, 30, 255, 40, 50, 60, 255, 70, 80, 90, 255, 100, 110, 120, 255,
    ];
    let image = ::image::RgbaImage::from_raw(4, 2, pixels.to_vec()).expect("4x2 rgba");
    let mut bytes = Vec::new();
    ::image::DynamicImage::ImageRgba8(image)
        .write_to(
            &mut std::io::Cursor::new(&mut bytes),
            ::image::ImageFormat::Png,
        )
        .expect("encode png");
    bytes
}

fn pixel(image: &image_impl::DecodedImage, x: usize, y: usize) -> [u8; 3] {
    let width = image.width as usize;
    let offset = (y * width + x) * 3;
    [
        image.rgb[offset],
        image.rgb[offset + 1],
        image.rgb[offset + 2],
    ]
}

fn assert_matrix(actual: [f32; 6], expected: [f32; 6]) {
    for (actual, expected) in actual.into_iter().zip(expected) {
        assert!(
            (actual - expected).abs() < 1e-3,
            "matrix {actual:?} does not match {expected:?}"
        );
    }
}

/// Map an image-space point through a placement matrix.
fn map_placement(matrix: [f32; 6], u: f32, v: f32) -> (f32, f32) {
    let [a, b, c, d, e, f] = matrix;
    (a * u + c * v + e, b * u + d * v + f)
}

/// Image-space `u` at the centre of `column` in a `width`-wide source.
fn column_u(column: usize, width: usize) -> f32 {
    (column as f32 + 0.5) / width as f32
}

/// Image-space `v` at the centre of `row`, whose 0 is the top row.
fn row_v(row: usize, height: usize) -> f32 {
    1.0 - (row as f32 + 0.5) / height as f32
}

/// The source pixel a page point samples through `matrix`.
fn source_pixel(matrix: [f32; 6], x: f32, y: f32, width: usize, height: usize) -> (usize, usize) {
    let [a, b, c, d, e, f] = matrix.map(f64::from);
    let determinant = a * d - b * c;
    assert!(determinant.abs() > 1e-6, "placement must be invertible");
    let (dx, dy) = (f64::from(x) - e, f64::from(y) - f);
    let u = ((d * dx - c * dy) / determinant).clamp(0.0, 1.0);
    let v = ((-b * dx + a * dy) / determinant).clamp(0.0, 1.0);
    let column = ((u * width as f64) as usize).min(width - 1);
    let row = (((1.0 - v) * height as f64) as usize).min(height - 1);
    (column, row)
}

/// A 4x2 SVG with opaque red and blue halves, no text and no raster refs.
fn banded_svg() -> &'static [u8] {
    br##"<svg xmlns="http://www.w3.org/2000/svg" width="4" height="2" viewBox="0 0 4 2"><rect width="2" height="2" fill="#ff0000"/><rect x="2" width="2" height="2" fill="#0000ff"/></svg>"##
}

/// A one-page display list whose only primitive is `image`.
fn display_list_with(image: &ImagePrimitive) -> DisplayList {
    serde_json::from_value(serde_json::json!({
        "pages": [{
            "pageIndex": 0,
            "width": 816,
            "height": 1056,
            "primitives": [{
                "kind": "image",
                "relId": image.rel_id,
                "x": image.x,
                "y": image.y,
                "w": image.w,
                "h": image.h,
            }],
        }],
    }))
    .expect("one-page display list")
}

/// One WMF record: its size in words, its function, then its parameter words.
fn wmf_record(function: u16, params: &[u8]) -> Vec<u8> {
    let mut params = params.to_vec();
    if params.len() % 2 != 0 {
        params.push(0);
    }
    let mut record = ((params.len() / 2 + 3) as u32).to_le_bytes().to_vec();
    record.extend(function.to_le_bytes());
    record.extend(params);
    record
}

/// A 288x192 placeable WMF painting one solid red rectangle over its whole
/// frame. Metafile replay needs a real placeable header, so the bytes are
/// built here rather than pulled from another crate's test fixtures.
fn red_wmf() -> Vec<u8> {
    // 2160x1440 logical units at 720 units per inch is 288x192 CSS pixels.
    let mut bytes = Vec::new();
    bytes.extend(0x9AC6_CDD7u32.to_le_bytes()); // placeable key
    bytes.extend(0u16.to_le_bytes()); // handle
    for value in [0i16, 0, 2160, 1440] {
        bytes.extend(value.to_le_bytes()); // left, top, right, bottom
    }
    bytes.extend(720u16.to_le_bytes()); // units per inch
    bytes.extend(0u32.to_le_bytes()); // reserved
    bytes.extend(0u16.to_le_bytes()); // checksum, which replay does not read

    let mut records = Vec::new();
    // Solid red brush (colour is 0x00BBGGRR), selected into slot 0, then a
    // rectangle over the whole picture: bottom, right, top, left.
    records.extend(wmf_record(0x02FC, &[0, 0, 0xff, 0, 0, 0, 0, 0]));
    records.extend(wmf_record(0x012D, &[0, 0]));
    records.extend(wmf_record(0x041B, &[0x9f, 0x05, 0x6f, 0x08, 0, 0, 0, 0]));
    records.extend(3u32.to_le_bytes());
    records.extend(0u16.to_le_bytes()); // META_EOF

    let words = (18 + records.len()) / 2;
    bytes.extend(1u16.to_le_bytes()); // memory metafile
    bytes.extend(9u16.to_le_bytes()); // header size in words
    bytes.extend(0x0300u16.to_le_bytes()); // version
    bytes.extend((words as u32).to_le_bytes());
    bytes.extend(16u16.to_le_bytes()); // object slots
    bytes.extend(7u32.to_le_bytes()); // largest record, in words
    bytes.extend(0u16.to_le_bytes()); // parameter words
    bytes.extend(records);
    bytes
}

/// One EMF record: its type, its byte length including the header, then its
/// parameters padded to a 4-byte boundary.
fn emf_record(kind: u32, params: &[u8]) -> Vec<u8> {
    let mut params = params.to_vec();
    if params.len() % 4 != 0 {
        params.resize(params.len().div_ceil(4) * 4, 0);
    }
    let mut record = kind.to_le_bytes().to_vec();
    record.extend(((params.len() + 8) as u32).to_le_bytes());
    record.extend(params);
    record
}

fn le_i32(values: &[i32]) -> Vec<u8> {
    values
        .iter()
        .flat_map(|value| value.to_le_bytes())
        .collect()
}

fn le_u32(values: &[u32]) -> Vec<u8> {
    values
        .iter()
        .flat_map(|value| value.to_le_bytes())
        .collect()
}

/// A 96 DPI EMF `width` by `height` device pixels over `records`.
fn emf(width: i32, height: i32, records: Vec<Vec<u8>>) -> Vec<u8> {
    let record_count = records.len() as u32 + 2; // plus header and EOF
    let mut body = Vec::new();
    for record in records {
        body.extend(record);
    }
    let frame = |px: i32| (f64::from(px) * 2540.0 / 96.0).round() as i32;
    let mut header = le_i32(&[0, 0, width - 1, height - 1]);
    header.extend(le_i32(&[0, 0, frame(width), frame(height)]));
    header.extend(le_u32(&[0x464D_4520, 0x0001_0000, 0, record_count]));
    header.extend([16u8, 0, 0, 0]); // object slots, reserved
    header.extend(le_u32(&[0, 0, 0])); // open, palette, reserved
    header.extend(le_i32(&[9600, 9600])); // reference device, px per inch
    header.extend(le_i32(&[2540, 2540])); // reference device, mm
    let mut out = le_u32(&[1, (header.len() + 8) as u32]);
    out.extend(header);
    out.extend(body);
    out.extend(le_u32(&[14, 20, 0, 16, 20])); // EOF
    let size = out.len() as u32;
    out[48..52].copy_from_slice(&size.to_le_bytes());
    out
}

/// `EXTCREATEFONTINDIRECTW` naming `face` at `height` logical units.
fn emf_font(handle: u32, height: i32, face: &str) -> Vec<u8> {
    let mut body = le_u32(&[handle]);
    body.extend(le_i32(&[height, 0, 0, 0, 400])); // lfHeight, width, escapements, weight
    body.extend([0u8; 8]); // italic..pitch and family, all defaults
    let mut name: Vec<u16> = face.encode_utf16().collect();
    name.resize(32, 0);
    for unit in name {
        body.extend(unit.to_le_bytes());
    }
    emf_record(82, &body)
}

/// `EXTTEXTOUTW` drawing `text` at `(x, y)`, offsets from the record start.
fn emf_text(x: i32, y: i32, text: &str) -> Vec<u8> {
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut string = Vec::new();
    for unit in &units {
        string.extend(unit.to_le_bytes());
    }
    if string.len() % 4 != 0 {
        string.resize(string.len().div_ceil(4) * 4, 0);
    }
    // EMRTEXT offsets include the 8-byte record header; the fixed body is 68
    // bytes: bounds 16, graphics mode 4, scale 8, point 8, text 12, rect 16,
    // offDx 4.
    let string_at = 8 + 68;
    let mut body = le_i32(&[0, 0, -1, -1]); // rclBounds, unset
    body.extend(le_u32(&[1])); // iGraphicsMode: GM_COMPATIBLE
    body.extend(1.0f32.to_le_bytes()); // exScale
    body.extend(1.0f32.to_le_bytes()); // eyScale
    body.extend(le_i32(&[x, y])); // ptlReference
    body.extend(le_u32(&[units.len() as u32, string_at, 0]));
    body.extend(le_i32(&[0, 0, -1, -1])); // rcl, unset
    body.extend(le_u32(&[0])); // offDx: no per-character advances
    body.extend(string);
    emf_record(84, &body)
}

/// A `FontStore` holding Liberation Sans, the Arial substitute the metafile
/// replay emits for the synthetic text fixtures.
fn typesetting_store() -> FontStore {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/fonts/assets/LiberationSans-Regular.ttf");
    let mut store = FontStore::default();
    store
        .register(std::fs::read(path).expect("font fixture"))
        .expect("register font");
    store
}

/// A 200x100 EMF drawing "Metafile" in Arial at 24 logical units.
fn text_emf() -> Vec<u8> {
    emf(
        200,
        100,
        vec![
            emf_font(1, -24, "Arial"),
            emf_record(37, &le_u32(&[1])), // EXTSELECTOBJECT
            emf_record(24, &le_u32(&[0])), // SETTEXTCOLOR, black
            emf_text(10, 30, "Metafile"),
        ],
    )
}

#[test]
fn decodes_a_png_into_rgb_and_alpha() {
    let decoded = image_impl::try_decode(&banded_png()).expect("valid png");
    assert_eq!((decoded.width, decoded.height), (4, 2));
    assert_eq!(pixel(&decoded, 0, 0), [255, 0, 0]);
    assert_eq!(pixel(&decoded, 1, 0), [0, 255, 0]);
    assert_eq!(pixel(&decoded, 2, 0), [0, 0, 255]);
    let alpha = decoded.alpha.as_ref().expect("one pixel is translucent");
    assert_eq!(alpha[2], 128);
    assert_eq!(alpha[0], 255);
}

#[test]
fn decodes_a_jpeg_without_alpha() {
    let source = ::image::RgbImage::from_raw(2, 1, vec![255, 0, 0, 0, 0, 255]).expect("2x1 rgb");
    let mut bytes = Vec::new();
    source
        .write_to(
            &mut std::io::Cursor::new(&mut bytes),
            ::image::ImageFormat::Jpeg,
        )
        .expect("encode jpeg");
    let decoded = image_impl::try_decode(&bytes).expect("jpeg");
    assert_eq!((decoded.width, decoded.height), (2, 1));
    assert_eq!(decoded.rgb.len(), 6);
    assert_eq!(decoded.alpha, None);
}

#[test]
fn crop_scales_the_visible_band_to_the_frame() {
    // 96 pt x 48 pt frame at (72, 492) on a 576 pt page, cropped to the
    // central half of the source on both axes.
    let image = primitive(Some(crop(0.25, 0.25, 0.25, 0.25)), None);
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    assert_eq!(transform.clip, [72.0, 492.0, 96.0, 48.0]);
    // sx = 96 / 0.5 = 192, sy = 48 / 0.5 = 96,
    // e = 72 - 0.25 * 192 = 24, f = 492 - 0.25 * 96 = 468.
    assert_matrix(transform.matrix, [192.0, 0.0, 0.0, 96.0, 24.0, 468.0]);

    // The frame corners are the visible source rect corners.
    let matrix = transform.matrix;
    let u_start = (transform.clip[0] - matrix[4]) / matrix[0];
    let u_end = (transform.clip[0] + transform.clip[2] - matrix[4]) / matrix[0];
    let v_start = (transform.clip[1] - matrix[5]) / matrix[3];
    let v_end = (transform.clip[1] + transform.clip[3] - matrix[5]) / matrix[3];
    assert!((u_start - 0.25).abs() < 1e-6);
    assert!((u_end - 0.75).abs() < 1e-6);
    assert!((v_start - 0.25).abs() < 1e-6);
    assert!((v_end - 0.75).abs() < 1e-6);

    // Cropping to [0.25, 0.75] x [0.25, 0.75] on a 4x2 source keeps the middle
    // two colour bands (green over blue) in both rows and drops the red and
    // white bands at the edges.
    let decoded = image_impl::try_decode(&banded_png()).expect("valid png");
    let first_column = (u_start * decoded.width as f32).round() as usize;
    let last_column = (u_end * decoded.width as f32).round() as usize;
    assert_eq!((first_column, last_column), (1, 3));
    assert_eq!(pixel(&decoded, first_column, 0), [0, 255, 0]);
    assert_eq!(pixel(&decoded, last_column - 1, 0), [0, 0, 255]);
    assert_eq!(pixel(&decoded, first_column, 1), [40, 50, 60]);
    assert_eq!(pixel(&decoded, last_column - 1, 1), [70, 80, 90]);
}

#[test]
fn content_frame_overrides_each_primitive_frame_field() {
    let mut image = primitive(None, None);
    image.attrs.content_frame = Some(content_frame(
        Some(192.0),
        Some(96.0),
        Some(64.0),
        Some(32.0),
    ));
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    // 192x96 at (192, 96) CSS px is 144x72 pt at (144, 72); frame_y is
    // 576 - 72 - 24 = 480. The primitive's own 96x48 at (96, 48) would give
    // [96, 0, 0, 48, 72, 492], so the override must be visible here.
    assert_matrix(transform.matrix, [48.0, 0.0, 0.0, 24.0, 144.0, 480.0]);
    assert_eq!(transform.clip, [144.0, 480.0, 48.0, 24.0]);

    // Only `x` is present: `y`, `w` and `h` still come from the primitive,
    // matching the native raster's field-by-field fallback.
    image.attrs.content_frame = Some(content_frame(Some(192.0), None, None, None));
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");
    // x = 144, y = 36, w = 96, h = 48; frame_y = 576 - 36 - 48 = 492.
    assert_matrix(transform.matrix, [96.0, 0.0, 0.0, 48.0, 144.0, 492.0]);
    assert_eq!(transform.clip, [144.0, 492.0, 96.0, 48.0]);

    // Without a content frame the primitive's own fields still place it.
    image.attrs.content_frame = None;
    let plain = image_impl::image_transform(&image, 576.0).expect("transform");
    assert_matrix(plain.matrix, [96.0, 0.0, 0.0, 48.0, 72.0, 492.0]);
    assert_eq!(plain.clip, [72.0, 492.0, 96.0, 48.0]);
}

#[test]
fn content_frame_with_crop_flip_and_rotation_maps_the_source_square() {
    // The combo fixture: a 200x200 content frame at (100, 200), `srcRect`
    // cropping 12.5% on the low side and 25% on the high side of each axis,
    // both flips, and 90 degrees of display rotation.
    let mut image = primitive(Some(crop(0.125, 0.125, 0.25, 0.25)), Some(90.0));
    image.attrs.content_frame = Some(content_frame(
        Some(100.0),
        Some(200.0),
        Some(200.0),
        Some(200.0),
    ));
    image.attrs.image_flip_h = Some(true);
    image.attrs.image_flip_v = Some(true);
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    // Frame 150x150 pt at (75, 276); crop scale 150 / 0.625 = 240 on both
    // axes; the centre is (150, 351). Hand-composed crop, mirror, rotation.
    assert_matrix(transform.matrix, [0.0, 240.0, -240.0, 0.0, 285.0, 246.0]);
    assert_eq!(transform.clip, [75.0, 276.0, 150.0, 150.0]);

    // The visible source rect [0.125, 0.75] x [0.25, 0.875] fills the frame;
    // its centre is the frame centre, so the composed matrix must keep it
    // there through the crop, mirror and rotation.
    let [a, b, c, d, e, f] = transform.matrix;
    let (u, v) = (0.125 + 0.625 / 2.0, 0.25 + 0.625 / 2.0);
    let (centre_x, centre_y) = (a * u + c * v + e, b * u + d * v + f);
    assert!((centre_x - 150.0).abs() < 1e-3, "x centre: {centre_x}");
    assert!((centre_y - 351.0).abs() < 1e-3, "y centre: {centre_y}");

    // An asymmetric sample proves the two mirrors are part of the composed
    // transform, not just the axis-aligned bounds. This page point samples
    // the red top-left band through both mirrors; without them the same
    // point samples the dark bottom band of column 2.
    let decoded = image_impl::try_decode(&banded_png()).expect("valid png");
    assert_eq!(source_pixel(transform.matrix, 100.0, 300.0, 4, 2), (0, 0));
    assert_eq!(pixel(&decoded, 0, 0), [255, 0, 0]);

    let mut plain = primitive(Some(crop(0.125, 0.125, 0.25, 0.25)), Some(90.0));
    plain.attrs.content_frame = Some(content_frame(
        Some(100.0),
        Some(200.0),
        Some(200.0),
        Some(200.0),
    ));
    let plain = image_impl::image_transform(&plain, 576.0).expect("transform");
    assert_eq!(source_pixel(plain.matrix, 100.0, 300.0, 4, 2), (2, 1));
    assert_eq!(pixel(&decoded, 2, 1), [70, 80, 90]);
}

#[test]
fn rotates_ninety_degrees_clockwise_about_the_frame_centre() {
    let image = primitive(None, Some(90.0));
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    // Rotated frame: 48 pt wide by 96 pt tall, same centre (120, 516).
    // Display-clockwise 90 degrees maps to PDF's -90 degree rotation.
    assert_matrix(transform.matrix, [0.0, -96.0, 48.0, 0.0, 96.0, 564.0]);
    assert_eq!(transform.clip, [72.0, 492.0, 96.0, 48.0]);

    let [a, b, c, d, e, f] = transform.matrix;
    let centre_x = a * 0.5 + c * 0.5 + e;
    let centre_y = b * 0.5 + d * 0.5 + f;
    assert!((centre_x - 120.0).abs() < 1e-3);
    assert!((centre_y - 516.0).abs() < 1e-3);
}

#[test]
fn flip_h_mirrors_asymmetric_pixels_across_the_frame() {
    let mut image = primitive(None, None);
    image.attrs.image_flip_h = Some(true);
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    // Crop scale 96x48 at (72, 492) mirrored about the frame centre x = 120:
    // the image unit square now runs right to left, the frame stays put.
    assert_matrix(transform.matrix, [-96.0, 0.0, 0.0, 48.0, 168.0, 492.0]);
    assert_matrix(transform.flip, [-1.0, 0.0, 0.0, 1.0, 240.0, 0.0]);
    assert_matrix(transform.rotation, [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]);
    assert_eq!(transform.clip, [72.0, 492.0, 96.0, 48.0]);

    // The banded PNG's leftmost band is dark and its rightmost is light. A
    // point near the frame's left edge samples the rightmost band only because
    // of the mirror; the unflipped placement samples the leftmost band there.
    let decoded = image_impl::try_decode(&banded_png()).expect("valid png");
    let (column, row) = source_pixel(transform.matrix, 78.0, 516.0, 4, 2);
    assert_eq!((column, row), (3, 1));
    assert_eq!(pixel(&decoded, column, row), [100, 110, 120]);

    let plain = primitive(None, None);
    let plain = image_impl::image_transform(&plain, 576.0).expect("transform");
    assert_eq!(source_pixel(plain.matrix, 78.0, 516.0, 4, 2), (0, 1));
    assert_eq!(pixel(&decoded, 0, 1), [10, 20, 30]);

    // The mirrored leftmost band now paints at the frame's right edge.
    let (x, _) = map_placement(transform.matrix, column_u(0, 4), row_v(1, 2));
    assert!((x - 156.0).abs() < 1e-3);
}

#[test]
fn flip_v_mirrors_asymmetric_pixels_across_the_frame() {
    let mut image = primitive(None, None);
    image.attrs.image_flip_v = Some(true);
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    // The frame stays at [72, 168] x [492, 540]; its content is mirrored
    // about y = 516.
    assert_matrix(transform.matrix, [96.0, 0.0, 0.0, -48.0, 72.0, 540.0]);
    assert_matrix(transform.flip, [1.0, 0.0, 0.0, -1.0, 0.0, 1032.0]);
    assert_eq!(transform.clip, [72.0, 492.0, 96.0, 48.0]);

    // The PNG's first (bright) row paints near the frame's bottom after the
    // mirror; the unflipped placement samples the dark second row there.
    let decoded = image_impl::try_decode(&banded_png()).expect("valid png");
    assert_eq!(source_pixel(transform.matrix, 76.0, 498.0, 4, 2), (0, 0));
    assert_eq!(pixel(&decoded, 0, 0), [255, 0, 0]);
    let plain = primitive(None, None);
    let plain = image_impl::image_transform(&plain, 576.0).expect("transform");
    assert_eq!(source_pixel(plain.matrix, 76.0, 498.0, 4, 2), (0, 1));
    assert_eq!(pixel(&decoded, 0, 1), [10, 20, 30]);

    // The centre of the source's top row maps below the frame centre.
    let (_, y) = map_placement(transform.matrix, column_u(0, 4), row_v(0, 2));
    assert!((y - 504.0).abs() < 1e-3);
}

#[test]
fn flip_with_crop_and_rotation_keeps_asymmetric_pixels_consistent() {
    let mut image = primitive(Some(crop(0.25, 0.25, 0.25, 0.25)), Some(90.0));
    image.attrs.image_flip_h = Some(true);
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");

    // Hand-computed: crop [192, 0, 0, 96, 24, 468], mirror about (120, 516),
    // then display-clockwise 90 degrees about the same centre.
    assert_matrix(transform.matrix, [0.0, 192.0, 96.0, 0.0, 72.0, 420.0]);
    assert_eq!(transform.clip, [72.0, 492.0, 96.0, 48.0]);

    let plain = primitive(Some(crop(0.25, 0.25, 0.25, 0.25)), Some(90.0));
    let plain = image_impl::image_transform(&plain, 576.0).expect("transform");
    assert_matrix(plain.matrix, [0.0, -192.0, 96.0, 0.0, 72.0, 612.0]);

    // The crop keeps source columns 1 and 2. At the rotated frame's lower-left
    // corner the mirrored placement shows column 1 where the unmirrored one
    // shows column 2.
    let decoded = image_impl::try_decode(&banded_png()).expect("valid png");
    assert_eq!(source_pixel(transform.matrix, 96.0, 492.0, 4, 2), (1, 1));
    assert_eq!(pixel(&decoded, 1, 1), [40, 50, 60]);
    assert_eq!(source_pixel(plain.matrix, 96.0, 492.0, 4, 2), (2, 1));
    assert_eq!(pixel(&decoded, 2, 1), [70, 80, 90]);

    // Column 1's centre stays on the rotated left edge, but the mirror moves
    // it to the frame's lower corner.
    let flipped_point = map_placement(transform.matrix, column_u(1, 4), row_v(1, 2));
    assert!((flipped_point.0 - 96.0).abs() < 1e-3);
    assert!((flipped_point.1 - 492.0).abs() < 1e-3);
    let plain_point = map_placement(plain.matrix, column_u(1, 4), row_v(1, 2));
    assert!((plain_point.0 - 96.0).abs() < 1e-3);
    assert!((plain_point.1 - 540.0).abs() < 1e-3);
}

#[test]
fn truncated_png_is_a_missing_image_not_a_document_error() {
    let png = banded_png();
    let truncated = &png[..png.len() / 2];
    assert!(image_impl::try_decode(truncated).is_none());
    assert!(image_impl::DecodedImage::from_bytes(truncated).is_err());
    // The strict decoder still accepts the intact source.
    assert!(image_impl::DecodedImage::from_bytes(&png).is_ok());
}

#[test]
fn decodes_a_svg_into_rgb_pixels() {
    let decoded = image_impl::try_decode(banded_svg()).expect("SVG decodes");
    // The shared sandboxed vector decoder supersamples up to 4x, so a 4x2
    // document renders a 16x8 raster.
    assert_eq!((decoded.width, decoded.height), (16, 8));
    assert_eq!(decoded.rgb.len(), 16 * 8 * 3);
    assert_eq!(decoded.alpha, None, "opaque fills need no soft mask");
    assert_eq!(pixel(&decoded, 0, 0), [255, 0, 0]);
    assert_eq!(pixel(&decoded, 7, 3), [255, 0, 0]);
    assert_eq!(pixel(&decoded, 8, 4), [0, 0, 255]);
    assert_eq!(pixel(&decoded, 15, 7), [0, 0, 255]);
}

#[test]
fn decodes_a_metafile_through_its_svg_replay() {
    let decoded = image_impl::try_decode(&red_wmf()).expect("metafile decodes");
    assert_eq!((decoded.width, decoded.height), (288, 192));
    assert_eq!(decoded.rgb.len(), 288 * 192 * 3);
    assert_eq!(
        pixel(&decoded, 144, 96),
        [255, 0, 0],
        "the red rectangle fills the replayed picture"
    );
    // Anti-aliased frame edges may leave a partial alpha plane; when present
    // it covers the same pixels as the colour planes.
    if let Some(alpha) = &decoded.alpha {
        assert_eq!(alpha.len(), 288 * 192);
    }
}

#[test]
fn metafile_text_renders_glyphs_when_document_fonts_are_available() {
    let bytes = text_emf();
    let mut budget = ooxml_metafile::ReplayBudget::default();
    let markup = ooxml_metafile::to_svg_with_budget(&bytes, &mut budget).expect("replay");
    assert!(
        ooxml_raster::svg::has_text_element(markup.markup.as_bytes()),
        "the replay keeps the text record: {}",
        markup.markup
    );

    // Without a font store the SVG parser uses an empty font database, so the
    // text run has no glyph outlines and nothing is painted.
    let untinted = image_impl::try_decode_with_fonts(&bytes, None).expect("metafile decodes");

    // The document font store reaches the metafile parser through
    // `metafile_fontdb`, which matches the replay's "Liberation Sans"
    // substitute for Arial.
    let store = typesetting_store();
    let typeset =
        image_impl::try_decode_with_fonts(&bytes, Some(&store)).expect("metafile decodes");

    let ink = |image: &image_impl::DecodedImage| {
        let alpha = image.alpha.as_deref();
        image
            .rgb
            .chunks_exact(3)
            .enumerate()
            .filter(|(index, rgb)| {
                alpha.is_some_and(|alpha| alpha[*index] != 0) || **rgb != [0, 0, 0]
            })
            .count()
    };
    assert_eq!(ink(&untinted), 0, "without fonts the text paints nothing");
    assert!(
        ink(&typeset) > 0,
        "document fonts must paint metafile text glyphs"
    );
    // Black glyphs leave RGB at zero and carry their coverage in alpha.
    assert_ne!(
        untinted.alpha, typeset.alpha,
        "glyph coverage must differ from the empty-fontdb path"
    );
}

#[test]
fn export_renders_metafile_text_with_the_document_font_store() {
    let mut image = primitive(None, None);
    image.rel_id = "rId-emf".to_owned();
    let display_list = display_list_with(&image);
    let mut images = BTreeMap::new();
    images.insert("rId-emf".to_owned(), text_emf());

    // An empty store still decodes the metafile, but with an empty fontdb, so
    // its XObject pixels carry no glyphs. A populated store must change them:
    // that only happens if `collect_image` forwards the store to the decoder.
    let without_fonts = betteroffice_docx_pdf::export_pdf_with_fonts_and_images(
        &display_list,
        &FontStore::default(),
        &images,
    )
    .expect("export");
    let with_fonts = betteroffice_docx_pdf::export_pdf_with_fonts_and_images(
        &display_list,
        &typesetting_store(),
        &images,
    )
    .expect("export");

    assert!(
        String::from_utf8_lossy(&with_fonts).contains("/Subtype /Image"),
        "the metafile must reach the PDF as an image XObject"
    );
    assert_ne!(
        without_fonts, with_fonts,
        "the font store must reach metafile decoding"
    );
}

#[test]
fn truncated_metafile_is_a_missing_image_not_a_document_error() {
    let wmf = red_wmf();
    let truncated = &wmf[..wmf.len() / 2];
    assert!(ooxml_metafile::is_metafile(truncated));
    assert!(image_impl::try_decode(truncated).is_none());
    assert!(image_impl::DecodedImage::from_bytes(truncated).is_err());
}

#[test]
fn svg_image_reaches_the_pdf_as_an_image_xobject() {
    let mut image = primitive(None, None);
    image.rel_id = "rId-svg".to_owned();
    let display_list = display_list_with(&image);
    let mut images = BTreeMap::new();
    images.insert("rId-svg".to_owned(), banded_svg().to_vec());
    let store = FontStore::default();
    let pdf =
        betteroffice_docx_pdf::export_pdf_with_fonts_and_images(&display_list, &store, &images)
            .expect("export");
    let pdf = String::from_utf8_lossy(&pdf);

    assert!(
        pdf.contains("/Subtype /Image"),
        "SVG bytes must become an image XObject, not be dropped"
    );
    // The supersampled 16x8 raster is the XObject's size.
    assert!(pdf.contains("/Width 16"), "XObject width: {pdf}");
    assert!(pdf.contains("/Height 8"), "XObject height: {pdf}");
    assert!(pdf.contains("/Im1"), "page XObject resource: {pdf}");
}

#[test]
fn opacity_is_clamped_and_only_reported_below_full() {
    let mut image = primitive(None, None);
    assert_eq!(image_impl::opacity_alpha(&image), None);
    image.opacity = Some(number(0.5));
    assert_eq!(image_impl::opacity_alpha(&image), Some(0.5));
    image.opacity = Some(number(2.0));
    assert_eq!(image_impl::opacity_alpha(&image), None);
    image.opacity = Some(number(-1.0));
    assert_eq!(image_impl::opacity_alpha(&image), Some(0.0));
}

#[test]
fn paint_emits_clip_transform_and_xobject() {
    let mut image = primitive(Some(crop(0.25, 0.25, 0.25, 0.25)), None);
    image.opacity = Some(number(0.5));
    let refs = ImageRefs {
        resource: "Im1".to_owned(),
        id: Ref::new(10),
        smask: None,
    };
    let mut content = Content::new();
    image_impl::paint_image(&mut content, 576.0, &image, &refs).expect("paint");
    let stream = String::from_utf8(content.finish().into_vec()).expect("ASCII content");

    assert!(stream.contains("re"), "clip rectangle: {stream}");
    assert!(stream.contains("W"), "clip operator: {stream}");
    assert!(stream.contains("cm"), "image transform: {stream}");
    assert!(stream.contains("/Im1 Do"), "image XObject: {stream}");
    // Opacity is the integration owner's ExtGState; this painter must not
    // reference a graphics state it did not write.
    assert!(!stream.contains("/GS"));
}

#[test]
fn paint_applies_rotation_flip_and_crop_in_that_order() {
    let mut image = primitive(Some(crop(0.25, 0.25, 0.25, 0.25)), Some(90.0));
    image.attrs.image_flip_h = Some(true);
    let transform = image_impl::image_transform(&image, 576.0).expect("transform");
    let refs = ImageRefs {
        resource: "Im1".to_owned(),
        id: Ref::new(10),
        smask: None,
    };
    let mut content = Content::new();
    image_impl::paint_image(&mut content, 576.0, &image, &refs).expect("paint");
    let stream = String::from_utf8(content.finish().into_vec()).expect("ASCII content");

    let matrices: Vec<[f32; 6]> = stream
        .lines()
        .filter_map(|line| line.strip_suffix(" cm"))
        .map(|operands| {
            let values: Vec<f32> = operands
                .split_whitespace()
                .map(|value| value.parse().expect("matrix operand"))
                .collect();
            values.try_into().expect("six matrix operands")
        })
        .collect();
    assert_eq!(matrices.len(), 3, "rotation, flip, crop: {stream}");
    assert_matrix(matrices[0], transform.rotation);
    assert_matrix(matrices[1], transform.flip);
    assert_matrix(matrices[2], transform.crop);

    // PDF's `cm` prepends, so the stream applies crop, then flip, then
    // rotation to the image unit square — the same placement as `matrix`.
    for (u, v) in [(0.0, 0.0), (1.0, 0.0), (0.25, 0.5), (1.0, 1.0)] {
        let mut point = map_placement(matrices[2], u, v);
        point = map_placement(matrices[1], point.0, point.1);
        point = map_placement(matrices[0], point.0, point.1);
        let expected = map_placement(transform.matrix, u, v);
        assert!(
            (point.0 - expected.0).abs() < 1e-3,
            "x at ({u}, {v}): {point:?} vs {expected:?}"
        );
        assert!(
            (point.1 - expected.1).abs() < 1e-3,
            "y at ({u}, {v}): {point:?} vs {expected:?}"
        );
    }
}

#[test]
fn scope_resolver_keeps_body_and_header_r_id5_apart() {
    let mut entries = HashMap::new();
    entries.insert("\u{1f}rId5".to_owned(), b"body-image".to_vec());
    entries.insert("hf:rId7\u{1f}rId5".to_owned(), b"header-image".to_vec());
    entries.insert("footnotes\u{1f}rId5".to_owned(), b"footnote-image".to_vec());
    entries.insert("endnotes\u{1f}rId5".to_owned(), b"endnote-image".to_vec());

    let images = image_scope::ScopedImages::from_registry(&entries);
    assert_eq!(images.bytes(None, "rId5"), Some(&b"body-image"[..]));
    assert_eq!(
        images.bytes(Some("rId7"), "rId5"),
        Some(&b"header-image"[..])
    );
    assert_eq!(images.body("rId5"), Some(&b"body-image"[..]));
    assert_eq!(
        images.header_footer("rId7", "rId5"),
        Some(&b"header-image"[..])
    );
    assert_eq!(images.notes(None, "rId5"), Some(&b"footnote-image"[..]));
    assert_eq!(
        images.notes(Some("endnote"), "rId5"),
        Some(&b"endnote-image"[..])
    );

    // A part that did not register rId5 must not fall back to the body bytes.
    assert_eq!(images.header_footer("rId8", "rId5"), None);
    assert_eq!(images.bytes(Some("rId8"), "rId5"), None);

    // The string form used by callers that already have a region kind.
    assert_eq!(
        images.bytes(Some("footnotes"), "rId5"),
        Some(&b"footnote-image"[..])
    );
    assert_eq!(
        images.bytes(Some("endnotes"), "rId5"),
        Some(&b"endnote-image"[..])
    );

    // A legacy flattened map still resolves body ids, but never a scoped part.
    let mut legacy = HashMap::new();
    legacy.insert("rId9".to_owned(), b"legacy-body".to_vec());
    let images = image_scope::ScopedImages::from_registry(&legacy);
    assert_eq!(images.body("rId9"), Some(&b"legacy-body"[..]));
    assert_eq!(images.header_footer("rId7", "rId9"), None);
}
