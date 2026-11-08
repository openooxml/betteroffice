#![cfg(feature = "downscale")]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use ooxml_raster::downscale::{decode_reduced, reduced_cost};

struct Counting;

static CURRENT: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            let now = CURRENT.fetch_add(layout.size(), Ordering::SeqCst) + layout.size();
            PEAK.fetch_max(now, Ordering::SeqCst);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        CURRENT.fetch_sub(layout.size(), Ordering::SeqCst);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let moved = unsafe { System.realloc(pointer, layout, size) };
        if !moved.is_null() {
            if size >= layout.size() {
                let now = CURRENT.fetch_add(size - layout.size(), Ordering::SeqCst)
                    + (size - layout.size());
                PEAK.fetch_max(now, Ordering::SeqCst);
            } else {
                CURRENT.fetch_sub(layout.size() - size, Ordering::SeqCst);
            }
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

const BUDGET: u64 = 1_000_000;

fn baseline_jpeg(width: u32, height: u32) -> Vec<u8> {
    let image = image::RgbImage::from_fn(width, height, |x, y| {
        image::Rgb([(x * 255 / width) as u8, (y * 255 / height) as u8, 128])
    });
    let mut bytes = Vec::new();
    image::codecs::jpeg::JpegEncoder::new(&mut bytes)
        .encode_image(&image)
        .expect("jpeg");
    bytes
}

fn png(width: u32, height: u32, depth: png::BitDepth) -> Vec<u8> {
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(depth);
        encoder.set_compression(png::Compression::Fast);
        let mut writer = encoder.write_header().expect("png header");
        let mut stream = writer.stream_writer().expect("png stream");
        let pixel: &[u8] = if depth == png::BitDepth::Sixteen {
            &[0x11, 0x11, 0x66, 0x66, 0xcc, 0xcc, 0xff, 0xff]
        } else {
            &[0x11, 0x66, 0xcc, 0xff]
        };
        let row = pixel.repeat(width as usize);
        for _ in 0..height {
            std::io::Write::write_all(&mut stream, &row).expect("png row");
        }
        stream.finish().expect("png finish");
    }
    bytes
}

/// Every reduced decode holds no more than it is charged, plus a fixed
/// allowance for decoder tables and the inflate window. One test, so no other
/// test's allocations land in the counters.
#[test]
fn a_reduced_decode_holds_no_more_than_its_charge() {
    let cases: Vec<(&str, Vec<u8>, u32, u32)> = vec![
        ("baseline jpeg", baseline_jpeg(1600, 1125), 1600, 1125),
        (
            "progressive jpeg",
            include_bytes!("fixtures/progressive-444.jpg").to_vec(),
            1600,
            1125,
        ),
        ("png", png(1600, 1125, png::BitDepth::Eight), 1600, 1125),
        ("half-thin jpeg", baseline_jpeg(32_000, 30), 32_000, 30),
        ("thin jpeg", baseline_jpeg(60_000, 30), 60_000, 30),
        (
            "half-wide 16-bit png",
            png(32_768, 30, png::BitDepth::Sixteen),
            32_768,
            30,
        ),
        (
            "wide 16-bit png",
            png(65_535, 30, png::BitDepth::Sixteen),
            65_535,
            30,
        ),
    ];
    for (name, bytes, width, height) in cases {
        let (_, charged) = reduced_cost(&bytes, width, height, BUDGET).expect("charge");
        let base = CURRENT.load(Ordering::SeqCst);
        PEAK.store(base, Ordering::SeqCst);
        let reduced = decode_reduced(&bytes, BUDGET);
        let measured = PEAK.load(Ordering::SeqCst) - base;
        let (_, reduced_width, reduced_height) = reduced.expect(name);
        eprintln!(
            "{name}: {width}x{height} -> {reduced_width}x{reduced_height}, charged {charged} B, measured peak {measured} B"
        );
        assert!(
            measured as u64 <= charged + charged / 20 + 256 * 1024,
            "{name}: measured {measured} B past the {charged} B charge"
        );
    }
}
