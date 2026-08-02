//! Per-page render budget probes, measured by allocation rather than by clock.

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use std::io::Write;

use docx_layout::display_list::DisplayList;
use docx_raster::{
    FontChains, ImageMap, ImageScope, MAX_PAGE_DIM, MAX_PAGE_PIXELS, RenderResources, render_page,
    render_png, scoped_image_key,
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
