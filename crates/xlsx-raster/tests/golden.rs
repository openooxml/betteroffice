//! png golden harness: scenarios are byte-compared against committed pngs.
//! regenerate deliberately with `GOLDEN_UPDATE=1 cargo test -p betteroffice-xlsx-raster`.

use std::path::PathBuf;

use xlsx_raster::render_png;
use xlsx_render::{Align, DisplayList, DrawCmd, GridMeta, Rect};

fn fill(x: f32, y: f32, w: f32, h: f32, color: &str) -> DrawCmd {
    DrawCmd::FillRect {
        x,
        y,
        w,
        h,
        color: color.into(),
    }
}

fn text(x: f32, y: f32, s: &str, align: Align, clip: Rect) -> DrawCmd {
    DrawCmd::Text {
        x,
        y,
        text: s.into(),
        font_size: 11.0,
        color: "#1a1a1a".into(),
        clip,
        align,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        highlight: None,
        dashed_underline: false,
        font_family: None,
        ghost: false,
    }
}

#[allow(clippy::too_many_arguments)]
fn styled_text(
    x: f32,
    y: f32,
    s: &str,
    color: &str,
    align: Align,
    clip: Rect,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
) -> DrawCmd {
    DrawCmd::Text {
        x,
        y,
        text: s.into(),
        font_size: 11.0,
        color: color.into(),
        clip,
        align,
        bold,
        italic,
        underline,
        strike,
        highlight: None,
        dashed_underline: false,
        font_family: None,
        ghost: false,
    }
}

fn line(
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    width: f32,
    color: &str,
    style: Option<&str>,
) -> DrawCmd {
    DrawCmd::Line {
        x1,
        y1,
        x2,
        y2,
        width,
        color: color.into(),
        style: style.map(Into::into),
    }
}

fn full(w: f32, h: f32) -> Rect {
    Rect {
        x: 0.0,
        y: 0.0,
        w,
        h,
    }
}

fn dl(width: f32, height: f32, commands: Vec<DrawCmd>) -> DisplayList {
    DisplayList {
        width,
        height,
        commands,
        grid: GridMeta::default(),
    }
}

fn scene_align() -> DisplayList {
    let clip = full(120.0, 48.0);
    dl(
        120.0,
        48.0,
        vec![
            text(2.0, 14.0, "Left", Align::Left, clip),
            text(60.0, 30.0, "Center", Align::Center, clip),
            text(118.0, 46.0, "Right", Align::Right, clip),
        ],
    )
}

fn scene_clipped() -> DisplayList {
    dl(
        40.0,
        16.0,
        vec![text(
            2.0,
            12.0,
            "Overflowing",
            Align::Left,
            Rect {
                x: 2.0,
                y: 2.0,
                w: 24.0,
                h: 12.0,
            },
        )],
    )
}

fn scene_mixed() -> DisplayList {
    let clip = full(60.0, 24.0);
    dl(
        60.0,
        24.0,
        vec![
            fill(0.0, 0.0, 60.0, 24.0, "#fff7d6"),
            line(0.0, 23.0, 60.0, 23.0, 1.0, "#d4d4d4", None),
            text(3.0, 16.0, "Cell A1", Align::Left, clip),
        ],
    )
}

fn scene_styled() -> DisplayList {
    let clip = full(120.0, 28.0);
    dl(
        120.0,
        28.0,
        vec![
            fill(0.0, 0.0, 120.0, 28.0, "#4472c4"),
            styled_text(
                4.0,
                18.0,
                "Revenue",
                "#ffffff",
                Align::Left,
                clip,
                true,
                false,
                false,
                false,
            ),
            styled_text(
                116.0,
                18.0,
                "1,250.00",
                "#ffffff",
                Align::Right,
                clip,
                false,
                true,
                true,
                false,
            ),
            line(0.0, 27.0, 120.0, 27.0, 3.0, "#1f3864", None),
            line(0.0, 4.0, 120.0, 4.0, 1.0, "#1f3864", Some("dashed")),
        ],
    )
}

// includes a cjk codepoint carlito lacks, which shapes to the .notdef tofu box.
fn scene_unicode() -> DisplayList {
    let clip = full(90.0, 20.0);
    dl(
        90.0,
        20.0,
        vec![text(2.0, 15.0, "café 你", Align::Left, clip)],
    )
}

fn golden_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("golden")
        .join(format!("{name}.png"))
}

fn check(name: &str, dl: &DisplayList) {
    let actual = render_png(dl).expect("render");
    let path = golden_path(name);
    if std::env::var("GOLDEN_UPDATE").is_ok() {
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, &actual).unwrap();
        return;
    }
    let expected = std::fs::read(&path).unwrap_or_else(|_| {
        panic!(
            "missing golden {}: regenerate with `GOLDEN_UPDATE=1 cargo test -p betteroffice-xlsx-raster`",
            path.display()
        )
    });
    assert!(
        actual == expected,
        "golden mismatch for {name}: if intended, regenerate with \
         `GOLDEN_UPDATE=1 cargo test -p betteroffice-xlsx-raster`"
    );
}

#[test]
fn golden_align() {
    check("align", &scene_align());
}

#[test]
fn golden_clipped() {
    check("clipped", &scene_clipped());
}

#[test]
fn golden_mixed() {
    check("mixed", &scene_mixed());
}

#[test]
fn golden_unicode() {
    check("unicode", &scene_unicode());
}

#[test]
fn golden_styled() {
    check("styled", &scene_styled());
}
