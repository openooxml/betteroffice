//! Image placement, decoding and part-scoped lookup.
//!
//! `image.rs` keeps its helpers `pub(crate)` and `lib.rs` is owned by the
//! integration slice, so the test crate mounts the real sources as modules.
//! That exercises the shipped code rather than a paraphrase of it.

#![allow(dead_code)]

use std::collections::HashMap;

use docx_layout::display_list::{Crop, DocAttrs, ImagePrimitive};
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
fn truncated_png_is_a_missing_image_not_a_document_error() {
    let png = banded_png();
    let truncated = &png[..png.len() / 2];
    assert!(image_impl::try_decode(truncated).is_none());
    assert!(image_impl::DecodedImage::from_bytes(truncated).is_err());
    // The strict decoder still accepts the intact source.
    assert!(image_impl::DecodedImage::from_bytes(&png).is_ok());
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
