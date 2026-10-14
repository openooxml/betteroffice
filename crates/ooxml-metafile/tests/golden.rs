//! Synthetic metafiles, their SVG replay rasterized with the repository's
//! metric-compatible fonts, compared against committed PNGs.
//!
//! `GOLDEN_UPDATE=1 cargo test -p betteroffice-metafile --features svg --test golden`
//! rewrites the fixtures and rasters.

mod common;

use std::path::{Path, PathBuf};

use resvg::tiny_skia::{Color, Pixmap, Transform};
use resvg::usvg;

const SCALE: f32 = 2.0;

fn here(parts: &[&str]) -> PathBuf {
    parts
        .iter()
        .fold(PathBuf::from(env!("CARGO_MANIFEST_DIR")), |path, part| {
            path.join(part)
        })
}

fn updating() -> bool {
    std::env::var_os("GOLDEN_UPDATE").is_some()
}

fn options() -> usvg::Options<'static> {
    let mut options = usvg::Options {
        font_family: "Liberation Sans".to_owned(),
        ..usvg::Options::default()
    };
    let fonts = here(&["..", "..", "packages", "fonts", "assets"]);
    for family in [
        "LiberationSans",
        "LiberationSerif",
        "LiberationMono",
        "Carlito",
        "Caladea",
    ] {
        for style in ["Regular", "Bold", "Italic", "BoldItalic"] {
            let path = fonts.join(format!("{family}-{style}.ttf"));
            let data = std::fs::read(&path)
                .unwrap_or_else(|error| panic!("load {}: {error}", path.display()));
            options.fontdb_mut().load_font_data(data);
        }
    }
    options
}

fn render(svg: &str, options: &usvg::Options<'_>) -> Pixmap {
    let tree = usvg::Tree::from_str(svg, options).expect("the replay writes valid SVG");
    let size = tree.size();
    let mut pixmap = Pixmap::new(
        (size.width() * SCALE).ceil() as u32,
        (size.height() * SCALE).ceil() as u32,
    )
    .unwrap();
    pixmap.fill(Color::WHITE);
    resvg::render(
        &tree,
        Transform::from_scale(SCALE, SCALE),
        &mut pixmap.as_mut(),
    );
    pixmap
}

/// A PNG as RGBA8 pixels.
fn decode(path: &Path) -> (u32, u32, Vec<u8>) {
    let file = std::fs::File::open(path).unwrap_or_else(|_| {
        panic!(
            "missing golden {}: regenerate with GOLDEN_UPDATE=1",
            path.display()
        )
    });
    let mut reader = png::Decoder::new(std::io::BufReader::new(file))
        .read_info()
        .unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    pixels.truncate(info.buffer_size());
    let pixels = match info.color_type {
        png::ColorType::Rgba => pixels,
        png::ColorType::Rgb => pixels
            .as_chunks::<3>()
            .0
            .iter()
            .flat_map(|pixel| [pixel[0], pixel[1], pixel[2], 255])
            .collect(),
        other => panic!("{} is {other:?}, not RGB or RGBA", path.display()),
    };
    (info.width, info.height, pixels)
}

/// Matches unless more than 0.5% of pixels differ by more than 24 in a channel,
/// which absorbs antialiasing drift between platforms.
fn compare(name: &str, pixmap: &Pixmap) {
    let path = here(&["tests", "golden", &format!("{name}.png")]);
    if updating() {
        pixmap.save_png(&path).unwrap();
        return;
    }
    let (width, height, expected) = decode(&path);
    assert_eq!(
        (width, height),
        (pixmap.width(), pixmap.height()),
        "{name} changed size"
    );
    let actual = pixmap.data();
    let differing = expected
        .as_chunks::<4>()
        .0
        .iter()
        .zip(actual.as_chunks::<4>().0)
        .filter(|(a, b)| a.iter().zip(*b).any(|(a, b)| a.abs_diff(*b) > 24))
        .count();
    if differing * 200 > (width * height) as usize {
        let out = std::env::temp_dir().join(format!("{name}.actual.png"));
        pixmap.save_png(&out).unwrap();
        panic!(
            "{name}: {differing} pixels differ from the golden; actual written to {}",
            out.display()
        );
    }
}

#[test]
fn fixtures_are_the_builders_output() {
    for (name, bytes) in common::fixtures::all() {
        let path = here(&["tests", "fixtures", name]);
        if updating() {
            std::fs::write(&path, &bytes).unwrap();
            continue;
        }
        let committed = std::fs::read(&path).unwrap_or_default();
        assert!(
            committed == bytes,
            "{name} is stale: regenerate with GOLDEN_UPDATE=1"
        );
    }
}

#[test]
fn synthetic_metafiles_match_their_golden_rasters() {
    let options = options();
    for (name, bytes) in common::fixtures::all() {
        let svg = ooxml_metafile::to_svg(&bytes).unwrap_or_else(|error| panic!("{name}: {error}"));
        assert!(
            svg.omissions.is_empty(),
            "{name} omitted {:?}",
            svg.omissions
        );
        let stem = name.replace('.', "-");
        compare(&stem, &render(&svg.markup, &options));
    }
}

#[test]
fn the_placeholder_matches_its_golden_raster() {
    compare(
        "placeholder",
        &render(&ooxml_metafile::placeholder_svg(160.0, 90.0), &options()),
    );
}

/// `libreoffice-one-line.emf` is LibreOffice's export of
/// `libreoffice-one-line.svg`; `libreoffice-one-line.word.png` is Word's
/// rendering of it at the same scale. Fonts and antialiasing differ, so the
/// two are compared as 12 px cells of average colour: a misplaced run, a
/// missing shape or a wrong fill moves whole cells.
#[test]
fn a_real_diagram_matches_words_rendering() {
    let bytes = std::fs::read(here(&["tests", "fixtures", "libreoffice-one-line.emf"])).unwrap();
    let svg = ooxml_metafile::to_svg(&bytes).unwrap();
    assert!(svg.omissions.is_empty(), "{:?}", svg.omissions);
    let ours = render(&svg.markup, &options());
    let (width, height, word) = decode(&here(&[
        "tests",
        "fixtures",
        "libreoffice-one-line.word.png",
    ]));
    assert_eq!((width, height), (ours.width(), ours.height()));
    const CELL: u32 = 12;
    let average = |pixels: &[u8], cx: u32, cy: u32| {
        let mut sum = [0u32; 3];
        let mut count = 0;
        for y in cy * CELL..((cy + 1) * CELL).min(height) {
            for x in cx * CELL..((cx + 1) * CELL).min(width) {
                let at = ((y * width + x) * 4) as usize;
                for (channel, total) in sum.iter_mut().enumerate() {
                    *total += u32::from(pixels[at + channel]);
                }
                count += 1;
            }
        }
        sum.map(|total| total / count)
    };
    let (columns, rows) = (width.div_ceil(CELL), height.div_ceil(CELL));
    let mut differing = Vec::new();
    for cy in 0..rows {
        for cx in 0..columns {
            let (a, b) = (average(ours.data(), cx, cy), average(&word, cx, cy));
            if a.iter().zip(b).any(|(a, b)| a.abs_diff(b) > 32) {
                differing.push((cx * CELL, cy * CELL));
            }
        }
    }
    if differing.len() * 500 > (columns * rows) as usize {
        let out = std::env::temp_dir().join("libreoffice-one-line.actual.png");
        ours.save_png(&out).unwrap();
        panic!(
            "{} of {} cells differ from Word, first at {:?}; ours written to {}",
            differing.len(),
            columns * rows,
            &differing[..differing.len().min(8)],
            out.display()
        );
    }
}
