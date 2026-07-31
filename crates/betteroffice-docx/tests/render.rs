#![cfg(feature = "raster")]

use std::time::{Duration, Instant};

use betteroffice_docx::{DisplayList, Document, MAX_PIXMAP_DIM, MAX_PIXMAP_PIXELS};
use serde_json::{Value, json};

const CARLITO: &[u8] = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
const SMALL_FONT: &[u8] =
    include_bytes!("../../../packages/fonts/assets/NotoSansHebrew-Regular.ttf");

/// A 4x4 opaque red PNG, as `docx-parse` hands embedded media to layout.
const RED_PNG_DATA_URL: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAEklEQVR42mP4z8DwHxkzkC4AADxAH+Ea86VIAAAAAElFTkSuQmCC";

const PNG_MAGIC: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

fn minimal_docx() -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#.to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn document() -> Document {
    Document::open(&minimal_docx()).unwrap()
}

fn page(primitives: Vec<Value>) -> DisplayList {
    sized_page(64, 32, primitives)
}

fn sized_page(width: u64, height: u64, primitives: Vec<Value>) -> DisplayList {
    serde_json::from_value(json!({
        "contractVersion": 1,
        "pages": [{
            "pageIndex": 0,
            "width": width,
            "height": height,
            "primitives": primitives
        }]
    }))
    .unwrap()
}

fn embedded_image_page() -> DisplayList {
    page(vec![json!({
        "kind": "image",
        "relId": RED_PNG_DATA_URL,
        "x": 0,
        "y": 0,
        "w": 32,
        "h": 32
    })])
}

fn text_page() -> DisplayList {
    page(vec![json!({
        "kind": "text",
        "text": "Ao",
        "x": 2,
        "baselineY": 20,
        "width": 24,
        "font": "400 16px Carlito, sans-serif",
        "color": "#101010"
    })])
}

fn decode(bytes: &[u8]) -> (u32, u32, Vec<u8>) {
    let mut reader = png::Decoder::new(std::io::Cursor::new(bytes))
        .read_info()
        .expect("png header");
    let mut pixels = vec![0_u8; reader.output_buffer_size().expect("buffer size")];
    let info = reader.next_frame(&mut pixels).expect("png frame");
    assert_eq!(info.color_type, png::ColorType::Rgba);
    pixels.truncate(info.buffer_size());
    (info.width, info.height, pixels)
}

fn pixel(width: u32, pixels: &[u8], x: u32, y: u32) -> [u8; 4] {
    let start = ((y * width + x) * 4) as usize;
    pixels[start..start + 4].try_into().expect("rgba pixel")
}

#[test]
fn renders_an_embedded_image_from_its_data_url_without_an_image_map() {
    let document = document();
    let png = document.render_png(&embedded_image_page(), 0).unwrap();
    assert_eq!(&png[..8], PNG_MAGIC);
    let (width, height, pixels) = decode(&png);
    assert_eq!((width, height), (64, 32));
    assert_eq!(pixel(width, &pixels, 16, 16), [255, 0, 0, 255]);
    assert_eq!(pixel(width, &pixels, 48, 16), [255, 255, 255, 255]);
}

#[test]
fn renders_text_only_once_its_family_is_registered() {
    let mut document = document();
    let error = document.render_png(&text_page(), 0).unwrap_err();
    assert_eq!(error.to_string(), "missing font chain for `carlito|0|0`");

    let id = document
        .register_font("Carlito", false, false, CARLITO)
        .unwrap();
    assert_eq!(id, 0);
    let png = document.render_png(&text_page(), 0).unwrap();
    assert_eq!(&png[..8], PNG_MAGIC);
    let (_, _, pixels) = decode(&png);
    let painted = pixels
        .chunks_exact(4)
        .filter(|p| *p != [255, 255, 255, 255])
        .count();
    assert!(painted > 0, "no glyphs were painted");
}

#[test]
fn rejects_a_page_ordinal_past_the_end() {
    let error = document()
        .render_png(&embedded_image_page(), 1)
        .unwrap_err();
    assert_eq!(error.to_string(), "page ordinal 1 is out of range");
}

#[test]
fn rejects_a_page_wider_than_the_per_side_budget() {
    let error = document()
        .render_png(&sized_page(MAX_PIXMAP_DIM as u64 + 1, 1, vec![]), 0)
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "requested render is 16385x1px, exceeds the 16384px per-side cap; lower the page dimensions"
    );
}

#[test]
fn rejects_a_page_over_the_area_budget_that_fits_the_per_side_budget() {
    let error = document()
        .render_png(&sized_page(4096, 4097, vec![]), 0)
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "requested render is 4096x4097px, exceeds the 16777216-pixel allocation cap; lower the page dimensions"
    );
    const { assert!(4096_u64 * 4097 > MAX_PIXMAP_PIXELS) };
}

/// Ordering probe. The page is over budget *and* carries a primitive the
/// backend refuses, so the reported error names whichever ran first.
#[test]
fn refuses_an_oversized_page_before_the_backend_paints_it() {
    let page = sized_page(
        u64::from(MAX_PIXMAP_DIM) + 1,
        1,
        vec![json!({
            "kind": "image",
            "relId": RED_PNG_DATA_URL,
            "x": 0,
            "y": 0,
            "w": 1,
            "h": 1,
            "filter": "grayscale(1)"
        })],
    );
    let error = document().render_png(&page, 0).unwrap_err();
    assert_eq!(
        error.to_string(),
        "requested render is 16385x1px, exceeds the 16384px per-side cap; lower the page dimensions"
    );
}

/// A 40000x40000 surface is 6.4 GB and minutes of fill, so the budget has to
/// return before the backend is called at all.
#[test]
fn refuses_a_page_that_would_allocate_gigabytes() {
    let started = Instant::now();
    let error = document()
        .render_png(&sized_page(40_000, 40_000, vec![]), 0)
        .unwrap_err();
    let elapsed = started.elapsed();
    assert_eq!(
        error.to_string(),
        "requested render is 40000x40000px, exceeds the 16384px per-side cap; lower the page dimensions"
    );
    assert!(
        elapsed < Duration::from_secs(10),
        "the page budget was enforced after allocation, in {elapsed:?}"
    );
}

#[test]
fn rejects_a_font_over_the_size_budget() {
    let mut document = document();
    let mut oversized = CARLITO.to_vec();
    oversized.resize(32 * 1024 * 1024 + 1, 0);
    let error = document
        .register_font("Carlito", false, false, &oversized)
        .unwrap_err();
    assert_eq!(error.to_string(), "font exceeds 33554432 bytes");
}

#[test]
fn rejects_more_faces_than_the_budget_allows() {
    let mut document = document();
    for index in 0..256 {
        document
            .register_font(&format!("Family{index}"), false, false, SMALL_FONT)
            .unwrap_or_else(|error| panic!("face {index} was refused: {error}"));
    }
    let error = document
        .register_font("OneTooMany", false, false, SMALL_FONT)
        .unwrap_err();
    assert_eq!(error.to_string(), "more than 256 font faces");
}

#[test]
fn rejects_an_empty_family_and_malformed_font_bytes() {
    let mut document = document();
    let error = document
        .register_font("  ", false, false, CARLITO)
        .unwrap_err();
    assert_eq!(error.to_string(), "font family is empty");
    assert!(
        document
            .register_font("Carlito", false, false, b"not a font")
            .is_err()
    );
}
