//! Per-page render budget probes, measured by allocation rather than by clock.

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use std::io::Write;

use docx_layout::display_list::DisplayList;
use docx_raster::{
    FontChains, ImageMap, ImageScope, MAX_IMAGE_BYTES, MAX_PAGE_DIM, MAX_PAGE_IMAGE_PIXELS,
    MAX_PAGE_PIXELS, RenderResources, render_page, render_png, scoped_image_key,
};
use ooxml_text::FontStore;
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

const MIB: u64 = 1 << 20;
const FONT: &[u8] = include_bytes!("assets/Carlito-Regular.ttf");

fn list(width: f64, height: f64, primitives: Vec<Value>) -> DisplayList {
    serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0,
            "width": width,
            "height": height,
            "primitives": primitives
        }]
    }))
    .expect("display list")
}

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

fn text_run(text: &str) -> Value {
    json!({
        "kind": "text",
        "text": text,
        "x": 2,
        "baselineY": 15,
        "width": 10,
        "font": "400 12px Carlito, sans-serif",
        "color": "#000000"
    })
}

fn carlito() -> (FontStore, FontChains) {
    let mut fonts = FontStore::new();
    let font = fonts.register(FONT.to_vec()).expect("font");
    let chains = FontChains::from([("carlito|0|0".to_string(), vec![font])]);
    (fonts, chains)
}

/// The published entry point is what a direct consumer of this crate calls, so
/// the surface cap has to live inside it rather than in a facade above it.
#[test]
fn the_published_entry_point_caps_the_page_surface_it_allocates() {
    const { assert!(4096_u64 * 4097 > MAX_PAGE_PIXELS) };
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    for (width, height, message) in [
        (
            f64::from(MAX_PAGE_DIM) + 1.0,
            1.0,
            "page is 16385x1px, past the 16384px per-side cap",
        ),
        (
            4096.0,
            4097.0,
            "page is 4096x4097px, past the 16777216-pixel surface cap",
        ),
    ] {
        let scene = list(width, height, Vec::new());
        let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
        assert_eq!(result.unwrap_err(), message);
        assert!(
            allocated < MIB,
            "{width}x{height} allocated {allocated} bytes before the cap refused it"
        );
    }
}

/// A crop mask covers the whole page whatever it crops, so the decode cache
/// alone does not bound 32 references to one 4x4 image: the mask has to be
/// reused where the geometry repeats.
#[test]
fn a_repeated_crop_geometry_reuses_one_page_mask() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::Body, "rIdCrop"),
        solid_png(4, 4),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitives = (0..32)
        .map(|_| {
            json!({
                "kind": "image",
                "relId": "rIdCrop",
                "x": 0,
                "y": 0,
                "w": 8,
                "h": 8,
                "crop": {"left": 0.1, "top": 0.1, "right": 0.1, "bottom": 0.1}
            })
        })
        .collect();
    let scene = list(2048.0, 2048.0, primitives);
    let (rendered, allocated) = allocated_by(|| {
        render_page(&scene, 0, &resources).unwrap_or_else(|error| panic!("render: {error}"))
    });
    assert_eq!(rendered.skipped_images, 0);
    assert!(
        allocated < 48 * MIB,
        "32 cropped references allocated {allocated} bytes, more than a reused mask"
    );
}

/// A4 at 96dpi, the page size the layout engine emits.
const A4: (f64, f64) = (794.0, 1123.0);

fn cropped(index: usize, left: f64) -> Value {
    json!({
        "kind": "image",
        "relId": "rIdCrop",
        "x": (index % 5) * 150,
        "y": (index / 5) * 130,
        "w": 120,
        "h": 100,
        "crop": {"left": left, "top": 0.1, "right": 0.1, "bottom": 0.1}
    })
}

/// A contact sheet is not exotic: a page of distinct crops is a page of
/// page-sized masks, and only one of them is ever live outside the cache. The
/// budget has to charge what the page holds rather than one mask per
/// reference, or a catalogue page fails as a whole.
#[test]
fn a_contact_sheet_of_distinct_crops_renders_on_an_a4_page() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::Body, "rIdCrop"),
        solid_png(4, 4),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitives = (0..40)
        .map(|index| cropped(index, 0.01 * index as f64))
        .collect();
    let scene = list(A4.0, A4.1, primitives);
    let rendered = render_page(&scene, 0, &resources)
        .unwrap_or_else(|error| panic!("40 distinct crops on an A4 page: {error}"));
    assert_eq!(rendered.skipped_images, 0);
}

/// One cropped icon per table row walks the mask cache in a round robin, the
/// order a cache that evicts by age of insertion misses on every single time.
#[test]
fn repeating_crop_geometries_reuse_their_masks_down_a_page() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::Body, "rIdCrop"),
        solid_png(4, 4),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitives = (0..60)
        .map(|index| {
            json!({
                "kind": "image",
                "relId": "rIdCrop",
                "x": (index % 5) * 150,
                "y": 8,
                "w": 120,
                "h": 100,
                "crop": {"left": 0.1, "top": 0.1, "right": 0.1, "bottom": 0.1}
            })
        })
        .collect();
    let scene = list(A4.0, A4.1, primitives);
    let (rendered, allocated) = allocated_by(|| {
        render_page(&scene, 0, &resources).unwrap_or_else(|error| panic!("render: {error}"))
    });
    assert_eq!(rendered.skipped_images, 0);
    assert!(
        allocated < 24 * MIB,
        "5 crop geometries over 60 references allocated {allocated} bytes, more than 5 masks"
    );
}

/// The resolutions a caller renders A4 at. Every budget scales with the page,
/// so a page that renders at 96dpi has to render at 300dpi too.
#[test]
fn a_realistic_page_renders_at_every_common_resolution() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::Body, "rIdCrop"),
        solid_png(4, 4),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let mut primitives: Vec<Value> = (0..40)
        .map(|index| cropped(index, 0.01 * index as f64))
        .collect();
    primitives.push(json!({
        "kind": "rect",
        "x": 0, "y": 0, "w": 794, "h": 1123,
        "fill": "#f5f5f5",
        "clipGroup": {"clip": {"x": 0, "y": 0, "w": 794, "h": 1123}}
    }));
    for row in 0..60 {
        primitives.push(json!({
            "kind": "rect",
            "x": 8, "y": 8 + row * 18, "w": 120, "h": 16,
            "fill": "#4472c4",
            "clipGroup": {"clip": {"x": 8, "y": 8 + row * 18, "w": 120, "h": 16}}
        }));
        primitives.push(json!({
            "kind": "line",
            "x1": 8, "y1": 26 + row * 18, "x2": 786, "y2": 26 + row * 18,
            "strokeWidth": 1.0,
            "color": "#c0c0c0",
            "borderStyle": "dashed"
        }));
    }
    for (dpi, scale) in [(96, 1.0), (150, 1.5625), (300, 3.125)] {
        let scene = list(A4.0 * scale, A4.1 * scale, primitives.clone());
        let rendered = render_page(&scene, 0, &resources)
            .unwrap_or_else(|error| panic!("a realistic A4 page at {dpi}dpi: {error}"));
        assert_eq!(rendered.skipped_images, 0, "at {dpi}dpi");
    }
}

/// A wave expands a numeric line length into path segments before anything
/// clips it, so the length has to be charged before the path grows.
#[test]
fn a_wave_past_the_work_budget_is_refused_before_it_expands() {
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        64.0,
        64.0,
        vec![json!({
            "kind": "line",
            "x1": 0,
            "y1": 32,
            "x2": 4_000_000,
            "y2": 32,
            "strokeWidth": 1.0,
            "color": "#000000",
            "borderStyle": "wave"
        })],
    );
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds its 33554432 byte render work budget"
    );
    assert!(
        allocated < MIB,
        "a refused wave allocated {allocated} bytes expanding its path"
    );
}

/// A wave's length is a subtraction of two display-list coordinates, and two
/// finite coordinates subtract to an infinite one. The budget has to see that
/// as unaffordable, or the loop that walks the wave never reaches its end.
#[test]
fn a_wave_whose_length_overflows_to_infinity_is_refused() {
    let (sender, receiver) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
        let resources = RenderResources::new(&fonts, &chains, &images);
        let scene = list(
            64.0,
            64.0,
            vec![json!({
                "kind": "line",
                "x1": -3.0e38_f64,
                "y1": 32,
                "x2": 3.0e38_f64,
                "y2": 32,
                "strokeWidth": 1.0,
                "color": "#000000",
                "borderStyle": "wave"
            })],
        );
        let _ = sender.send(allocated_by(|| render_png(&scene, 0, &resources)));
    });
    let (result, allocated) = receiver
        .recv_timeout(std::time::Duration::from_secs(5))
        .unwrap_or_else(|_| panic!("a wave over an infinite length never returned"));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds its 33554432 byte render work budget"
    );
    assert!(
        allocated < MIB,
        "a refused wave allocated {allocated} bytes expanding its path"
    );
}

/// A dash expands one drawn segment into a subpath per interval, in path space
/// and before anything clips it, so the dash is as much a generated path as the
/// wave is and has to be charged like one.
#[test]
fn a_dashed_line_past_the_work_budget_is_refused_before_it_expands() {
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        64.0,
        64.0,
        vec![json!({
            "kind": "line",
            "x1": 0,
            "y1": 32,
            "x2": 4_000_000,
            "y2": 32,
            "strokeWidth": 0.5,
            "color": "#000000",
            "borderStyle": "dashed"
        })],
    );
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds its 33554432 byte render work budget"
    );
    assert!(
        allocated < MIB,
        "a refused dash allocated {allocated} bytes expanding its path"
    );
}

/// A shape stroke dashes an arbitrary geometry path, so its length has to be
/// charged the same way a line's is.
#[test]
fn a_dashed_shape_path_past_the_work_budget_is_refused_before_it_expands() {
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        64.0,
        64.0,
        vec![json!({
            "kind": "shape",
            "x": 0,
            "y": 0,
            "w": 64,
            "h": 64,
            "geometryPath": [
                {"type": "move", "x": 0, "y": 32},
                {"type": "line", "x": 4_000_000, "y": 32}
            ],
            "strokePaint": {"color": "#000000", "width": 0.5, "dash": "dash"}
        })],
    );
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds its 33554432 byte render work budget"
    );
    assert!(
        allocated < MIB,
        "a refused shape dash allocated {allocated} bytes expanding its path"
    );
}

/// An image border dashes the frame's rectangle, whose perimeter is four more
/// numbers off the display list.
#[test]
fn a_dashed_image_border_past_the_work_budget_is_refused_before_it_expands() {
    let (fonts, chains) = (FontStore::new(), FontChains::new());
    let images = ImageMap::from([(
        scoped_image_key(ImageScope::Body, "rIdEdge"),
        solid_png(4, 4),
    )]);
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        64.0,
        64.0,
        vec![json!({
            "kind": "image",
            "relId": "rIdEdge",
            "x": 0,
            "y": 0,
            "w": 4_000_000,
            "h": 8,
            "border": {"width": 0.5, "style": "dash"}
        })],
    );
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds its 33554432 byte render work budget"
    );
    assert!(
        allocated < MIB,
        "a refused image border allocated {allocated} bytes expanding its path"
    );
}

/// The glyph budget has to be charged against the text before the shaper runs,
/// not against the glyph vector the shaper already allocated.
#[test]
fn an_oversized_text_run_is_refused_before_the_shaper_allocates() {
    let (fonts, chains) = carlito();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let text = " ".repeat(2_000_000);
    let scene = list(64.0, 32.0, vec![text_run(&text)]);
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds the 1000000 painted glyph budget"
    );
    assert!(
        allocated < MIB,
        "a refused text run allocated {allocated} bytes shaping glyphs"
    );
}

/// A `font` shorthand is a display-list string of any length, and a page may
/// carry a distinct one per run. Parsing each one once must not mean holding a
/// copy of every raw string for the length of the page.
#[test]
fn a_page_of_distinct_font_shorthands_does_not_retain_them() {
    let (fonts, chains) = carlito();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitives: Vec<Value> = (0..512)
        .map(|index| {
            json!({
                "kind": "text",
                "text": "x",
                "x": 1,
                "baselineY": 6,
                "width": 4,
                "font": format!("400 {}12px Carlito, sans-serif", " ".repeat(16_384 + index)),
                "color": "#000000",
                "leaderGlyphs": {"glyph": ".", "count": 1, "advance": 0}
            })
        })
        .collect();
    let scene = list(8.0, 8.0, primitives);
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    result.expect("render");
    assert!(
        allocated < MIB,
        "a page of distinct shorthands allocated {allocated} bytes keeping them"
    );
}

/// A leader repeats its glyph `count` times, and the glyph is a display-list
/// string of any length. The charge is the product, not the count: 1:1 shaping
/// leaves no excess for the shaped-glyph charge to recover.
#[test]
fn a_leader_run_is_charged_for_every_character_it_repeats() {
    let (fonts, chains) = carlito();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let scene = list(
        64.0,
        32.0,
        vec![json!({
            "kind": "text",
            "text": ".",
            "x": 2,
            "baselineY": 15,
            "width": 10,
            "font": "400 12px Carlito, sans-serif",
            "color": "#000000",
            "leaderGlyphs": {
                "glyph": ".".repeat(51),
                "count": 20_000,
                "advance": 1
            }
        })],
    );
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds the 1000000 painted glyph budget"
    );
    assert!(
        allocated < MIB,
        "a refused leader allocated {allocated} bytes painting its glyph"
    );
}

/// `allCaps` uppercases before shaping, and one character can uppercase into
/// three, so the charge has to be the uppercased length.
#[test]
fn an_all_caps_run_is_charged_for_the_characters_it_uppercases_into() {
    let (fonts, chains) = carlito();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let mut run = text_run(&"\u{fb03}".repeat(400_000));
    run["allCaps"] = json!(true);
    let scene = list(64.0, 32.0, vec![run]);
    let (result, allocated) = allocated_by(|| render_png(&scene, 0, &resources));
    assert_eq!(
        result.unwrap_err(),
        "page exceeds the 1000000 painted glyph budget"
    );
    assert!(
        allocated < MIB,
        "a refused allCaps run allocated {allocated} bytes uppercasing its text"
    );
}

/// The glyph budget belongs to the page, not to one run: two runs that each
/// fit on their own still have to run out together.
#[test]
fn the_glyph_budget_is_spent_across_every_run_on_a_page() {
    let (fonts, chains) = carlito();
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    let text = " ".repeat(600_001);
    let scene = list(64.0, 32.0, vec![text_run(&text), text_run(&text)]);
    assert_eq!(
        render_png(&scene, 0, &resources).unwrap_err(),
        "page exceeds the 1000000 painted glyph budget"
    );
}

fn pixel(png: &[u8], x: u32, y: u32) -> [u8; 4] {
    let mut reader = png::Decoder::new(std::io::Cursor::new(png))
        .read_info()
        .expect("png header");
    let mut pixels = vec![0_u8; reader.output_buffer_size().expect("buffer size")];
    let info = reader.next_frame(&mut pixels).expect("png frame");
    let start = ((y * info.width + x) * 4) as usize;
    pixels[start..start + 4].try_into().expect("rgba pixel")
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = (crc >> 1) ^ (0xedb8_8320 & 0_u32.wrapping_sub(crc & 1));
        }
    }
    !crc
}

/// A PNG that declares a large extent at a given colour depth but carries one
/// pixel of data. The header is all a budget may consult, so this costs what
/// it declares and never what it decodes to.
fn png_bomb(width: u32, height: u32, depth: png::BitDepth) -> Vec<u8> {
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, 1, 1);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(depth);
        let mut writer = encoder.write_header().expect("png header");
        let pixel = if depth == png::BitDepth::Sixteen {
            vec![0_u8; 8]
        } else {
            vec![0_u8; 4]
        };
        writer.write_image_data(&pixel).expect("png data");
        writer.finish().expect("png finish");
    }
    bytes[16..20].copy_from_slice(&width.to_be_bytes());
    bytes[20..24].copy_from_slice(&height.to_be_bytes());
    let crc = crc32(&bytes[12..29]);
    bytes[29..33].copy_from_slice(&crc.to_be_bytes());
    bytes
}

/// A `data:` payload was decoded from base64 before anything looked at what it
/// declared, so the encoded length is the only thing that can bound it.
#[test]
fn an_oversized_data_url_is_refused_before_its_base64_is_decoded() {
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let reference = format!("data:image/png;base64,{}", "A".repeat(48 * 1024 * 1024));
    let scene = list(
        1.0,
        1.0,
        vec![json!({"kind":"image","relId":reference,"x":0,"y":0,"w":1,"h":1})],
    );
    let (rendered, allocated) = allocated_by(|| {
        render_page(&scene, 0, &resources).unwrap_or_else(|error| panic!("render: {error}"))
    });
    assert_eq!(rendered.skipped_images, 1);
    assert!(
        allocated < MIB,
        "an oversized data URL allocated {allocated} bytes decoding base64 on a 1x1 page"
    );
}

/// The decode cache is keyed by the reference, so a repeated `data:` URL must
/// not repeat the base64 decode that the cached result already paid for.
#[test]
fn a_repeated_data_url_decodes_its_base64_once() {
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let reference = format!("data:image/png;base64,{}", "A".repeat(4 * 1024 * 1024));
    let primitives = (0..16)
        .map(|index| json!({"kind":"image","relId":reference,"x":index,"y":0,"w":8,"h":8}))
        .collect();
    let scene = list(64.0, 32.0, primitives);
    let (rendered, allocated) = allocated_by(|| {
        render_page(&scene, 0, &resources).unwrap_or_else(|error| panic!("render: {error}"))
    });
    assert_eq!(rendered.skipped_images, 16);
    assert!(
        allocated < 4 * MIB,
        "16 references to one data URL allocated {allocated} bytes, more than one decode"
    );
}

/// Pixels are not bytes: a 16-bit source needs twice the buffer an 8-bit one
/// of the same extent does. Two 8-bit bombs spend the page's whole image
/// budget and squeeze out the image after them; two 16-bit ones need more than
/// one image may have, so they are refused on their own and spend nothing.
#[test]
fn an_image_is_charged_by_the_bytes_it_decodes_to_not_its_pixels() {
    const { assert!(4096 * 8192 * 8 + 4096 * 8192 * 4 > MAX_IMAGE_BYTES) };
    const { assert!(2 * 4096 * 8192 == MAX_PAGE_IMAGE_PIXELS) };
    for (depth, skipped, drawn) in [
        (png::BitDepth::Eight, 3, [255, 255, 255, 255]),
        (png::BitDepth::Sixteen, 2, [0x11, 0x66, 0xcc, 255]),
    ] {
        let (fonts, chains) = (FontStore::new(), FontChains::new());
        let images = ImageMap::from([
            (
                scoped_image_key(ImageScope::Body, "rIdTall"),
                png_bomb(4096, 8192, depth),
            ),
            (
                scoped_image_key(ImageScope::Body, "rIdWide"),
                png_bomb(8192, 4096, depth),
            ),
            (
                scoped_image_key(ImageScope::Body, "rIdSmall"),
                solid_png(8, 8),
            ),
        ]);
        let resources = RenderResources::new(&fonts, &chains, &images);
        let scene = list(
            20.0,
            20.0,
            vec![
                json!({"kind":"image","relId":"rIdTall","x":0,"y":0,"w":5,"h":5}),
                json!({"kind":"image","relId":"rIdWide","x":5,"y":0,"w":5,"h":5}),
                json!({"kind":"image","relId":"rIdSmall","x":0,"y":0,"w":20,"h":20}),
            ],
        );
        let rendered = render_page(&scene, 0, &resources).expect("render");
        assert_eq!(rendered.skipped_images, skipped, "at {depth:?}");
        assert_eq!(pixel(&rendered.bytes, 10, 10), drawn, "at {depth:?}");
    }
}

/// A reference the backend will not draw still takes a cache entry, and a
/// `data:` URL is as long as its author made it. The cache may not copy those
/// keys: they already live in the display list, and there is no bound on how
/// many an attacker sends.
#[test]
fn the_decode_cache_does_not_copy_attacker_sized_keys() {
    let (fonts, chains, images) = (FontStore::new(), FontChains::new(), ImageMap::new());
    let resources = RenderResources::new(&fonts, &chains, &images);
    let primitives = (0..512)
        .map(|index| {
            let reference = format!("data:image/png;x={}{index};base64,QQ==", "p".repeat(65_536));
            json!({"kind":"image","relId":reference,"x":0,"y":0,"w":8,"h":8})
        })
        .collect();
    let scene = list(64.0, 32.0, primitives);
    let (rendered, allocated) = allocated_by(|| {
        render_page(&scene, 0, &resources).unwrap_or_else(|error| panic!("render: {error}"))
    });
    assert_eq!(rendered.skipped_images, 512);
    assert!(
        allocated < 8 * MIB,
        "512 distinct 64KiB references cost {allocated} bytes of cache key"
    );
}
