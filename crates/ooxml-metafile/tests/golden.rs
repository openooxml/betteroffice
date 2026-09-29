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
            options
                .fontdb_mut()
                .load_font_file(&path)
                .unwrap_or_else(|error| panic!("load {}: {error}", path.display()));
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
    assert_eq!(info.color_type, png::ColorType::Rgba);
    pixels.truncate(info.buffer_size());
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
        .chunks_exact(4)
        .zip(actual.chunks_exact(4))
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
