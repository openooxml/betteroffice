#![cfg(feature = "raster")]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;
use std::io::Write;

use betteroffice_docx::{
    DisplayList, Document, ImageScope, MAX_PIXMAP_DIM, MAX_PIXMAP_PIXELS, NoteKind,
};
use serde_json::{Value, json};

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

thread_local! {
    static ALLOCATED: Cell<u64> = const { Cell::new(0) };
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let _ = ALLOCATED.try_with(|counter| counter.set(counter.get() + layout.size() as u64));
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
}

/// Bytes this thread allocated while `body` ran. Tests share the process, so
/// the counter is per thread.
fn allocated_by<T>(body: impl FnOnce() -> T) -> (T, u64) {
    let before = ALLOCATED.with(Cell::get);
    let value = body();
    (value, ALLOCATED.with(Cell::get) - before)
}

/// A solid RGBA PNG of any extent, streamed a row at a time.
fn solid_png(width: u32, height: u32) -> Vec<u8> {
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().expect("png header");
        let mut stream = writer.stream_writer().expect("png stream");
        let row: Vec<u8> = [0x11_u8, 0x66, 0xcc, 0xff].repeat(width as usize);
        for _ in 0..height {
            stream.write_all(&row).expect("png row");
        }
        stream.finish().expect("png finish");
    }
    bytes
}

const CARLITO: &[u8] = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
const SMALL_FONT: &[u8] =
    include_bytes!("../../../packages/fonts/assets/NotoSansHebrew-Regular.ttf");

/// A 4x4 opaque red PNG, as `docx-parse` hands embedded media to layout.
const RED_PNG_DATA_URL: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAEklEQVR42mP4z8DwHxkzkC4AADxAH+Ea86VIAAAAAElFTkSuQmCC";

const PNG_MAGIC: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

/// 4x4 solid PNGs, registered as the media an unresolved `rId` stands for.
const RED_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x04, 0x08, 0x02, 0x00, 0x00, 0x00, 0x26, 0x93, 0x09,
    0x29, 0x00, 0x00, 0x00, 0x10, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x47, 0x0C, 0xC4, 0x71, 0x00, 0xAE, 0x93, 0x0F, 0xF1, 0x38, 0x5E, 0x8C, 0x11, 0x00, 0x00, 0x00,
    0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];
const BLUE_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x04, 0x08, 0x02, 0x00, 0x00, 0x00, 0x26, 0x93, 0x09,
    0x29, 0x00, 0x00, 0x00, 0x10, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0x60, 0x60, 0xF8, 0x8F,
    0x84, 0x88, 0xE2, 0x00, 0x00, 0x8E, 0xB3, 0x0F, 0xF1, 0x5B, 0xA2, 0x80, 0xBC, 0x00, 0x00, 0x00,
    0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

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

/// One `rId9` image in the body and another in a header whose part is `rId7`.
/// Word hands out relationship ids per part, so these are different media.
fn shared_relationship_id_page() -> DisplayList {
    serde_json::from_value(json!({
        "contractVersion": 1,
        "pages": [{
            "pageIndex": 0,
            "width": 64,
            "height": 32,
            "primitives": [
                {"kind": "image", "relId": "rId9", "x": 0, "y": 0, "w": 16, "h": 16}
            ],
            "header": {
                "rId": "rId7",
                "kind": "header",
                "y": 0,
                "height": 16,
                "primitives": [
                    {"kind": "image", "relId": "rId9", "x": 32, "y": 0, "w": 16, "h": 16}
                ]
            }
        }]
    }))
    .unwrap()
}

fn repeated_image_page(count: usize, rel_id: &str) -> DisplayList {
    page(
        (0..count)
            .map(|index| json!({"kind":"image","relId":rel_id,"x":index,"y":0,"w":16,"h":16}))
            .collect(),
    )
}

/// A note area that leaves `kind` out, which the display list does for
/// footnotes.
fn unnamed_note_area_page() -> DisplayList {
    serde_json::from_value(json!({
        "contractVersion": 1,
        "pages": [{
            "pageIndex": 0,
            "width": 64,
            "height": 32,
            "primitives": [],
            "noteAreas": [{
                "y": 0,
                "height": 16,
                "primitives": [
                    {"kind": "image", "relId": "rId9", "x": 0, "y": 0, "w": 16, "h": 16}
                ]
            }]
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
    let rendered = document.render_png(&embedded_image_page(), 0).unwrap();
    assert_eq!(rendered.skipped_images, 0);
    assert_eq!(&rendered.bytes[..8], PNG_MAGIC);
    let (width, height, pixels) = decode(&rendered.bytes);
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
    let rendered = document.render_png(&text_page(), 0).unwrap();
    assert_eq!(&rendered.bytes[..8], PNG_MAGIC);
    let (_, _, pixels) = decode(&rendered.bytes);
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
fn an_unresolvable_image_is_skipped_rather_than_failing_the_page() {
    let document = document();
    let rendered = document
        .render_png(&shared_relationship_id_page(), 0)
        .unwrap();
    assert_eq!(rendered.skipped_images, 2);
    let (width, _, pixels) = decode(&rendered.bytes);
    assert_eq!(pixel(width, &pixels, 8, 8), [255, 255, 255, 255]);
    assert_eq!(pixel(width, &pixels, 40, 8), [255, 255, 255, 255]);
}

#[test]
fn a_registered_image_resolves_per_part_so_a_shared_id_stays_distinct() {
    let mut document = document();
    document
        .register_image(ImageScope::Body, "rId9", RED_PNG)
        .unwrap();
    document
        .register_image(ImageScope::HeaderFooter("rId7"), "rId9", BLUE_PNG)
        .unwrap();
    let rendered = document
        .render_png(&shared_relationship_id_page(), 0)
        .unwrap();
    assert_eq!(rendered.skipped_images, 0);
    let (width, _, pixels) = decode(&rendered.bytes);
    assert_eq!(pixel(width, &pixels, 8, 8), [255, 0, 0, 255]);
    assert_eq!(pixel(width, &pixels, 40, 8), [0, 0, 255, 255]);
}

#[test]
fn a_registration_does_not_leak_into_another_part() {
    let mut document = document();
    document
        .register_image(ImageScope::HeaderFooter("rId7"), "rId9", BLUE_PNG)
        .unwrap();
    let rendered = document
        .render_png(&shared_relationship_id_page(), 0)
        .unwrap();
    assert_eq!(rendered.skipped_images, 1);
    let (width, _, pixels) = decode(&rendered.bytes);
    assert_eq!(pixel(width, &pixels, 8, 8), [255, 255, 255, 255]);
    assert_eq!(pixel(width, &pixels, 40, 8), [0, 0, 255, 255]);
}

#[test]
fn a_note_image_resolves_for_an_area_that_omits_its_kind() {
    let mut document = document();
    document
        .register_image(ImageScope::Notes(NoteKind::Footnote), "rId9", RED_PNG)
        .unwrap();
    let rendered = document.render_png(&unnamed_note_area_page(), 0).unwrap();
    assert_eq!(rendered.skipped_images, 0);
    let (width, _, pixels) = decode(&rendered.bytes);
    assert_eq!(pixel(width, &pixels, 8, 8), [255, 0, 0, 255]);
}

#[test]
fn rejects_an_image_registration_the_backend_could_never_match() {
    let mut document = document();
    let error = document
        .register_image(ImageScope::Body, "", RED_PNG)
        .unwrap_err();
    assert_eq!(error.to_string(), "image relationship id is empty");

    let oversized = vec![0_u8; 32 * 1024 * 1024 + 1];
    let error = document
        .register_image(ImageScope::Body, "rId9", &oversized)
        .unwrap_err();
    assert_eq!(error.to_string(), "image exceeds 33554432 bytes");

    for index in 0..256 {
        document
            .register_image(ImageScope::Body, &format!("rId{index}"), RED_PNG)
            .unwrap_or_else(|error| panic!("image {index} was refused: {error}"));
    }
    let error = document
        .register_image(ImageScope::Body, "rIdOneTooMany", RED_PNG)
        .unwrap_err();
    assert_eq!(error.to_string(), "more than 256 images");
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

/// Ordering probe. Which error wins survives a reorder, so this watches the
/// surface instead: the backend allocates the page before it paints anything,
/// and the budget has to return before that.
#[test]
fn refuses_an_oversized_page_before_the_backend_allocates_its_surface() {
    let document = document();
    for (width, height, message) in [
        (
            u64::from(MAX_PIXMAP_DIM) + 1,
            256,
            "requested render is 16385x256px, exceeds the 16384px per-side cap; lower the page dimensions",
        ),
        (
            4096,
            4097,
            "requested render is 4096x4097px, exceeds the 16777216-pixel allocation cap; lower the page dimensions",
        ),
    ] {
        let page = sized_page(width, height, vec![]);
        let (error, allocated) = allocated_by(|| document.render_png(&page, 0).unwrap_err());
        assert_eq!(error.to_string(), message);
        assert!(
            allocated < 1 << 20,
            "{width}x{height} allocated {allocated} bytes before the budget refused it"
        );
    }
}

/// A reference is 59 bytes of display list and a decode is milliseconds, so
/// repeated references to one image have to cost one decode.
#[test]
fn repeated_references_to_one_image_decode_once() {
    let mut document = document();
    document
        .register_image(ImageScope::Body, "rId9", &solid_png(512, 512))
        .unwrap();
    let list = repeated_image_page(32, "rId9");
    let (rendered, allocated) = allocated_by(|| document.render_png(&list, 0).unwrap());
    assert_eq!(rendered.skipped_images, 0);
    assert!(
        allocated < 4 * 512 * 512 * 4,
        "32 references to one image allocated {allocated} bytes, more than a single decode"
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
