use serde_json::{Value, json};
use tiny_skia::Pixmap;
use xlsx_model::{Cell, CellRef, CellValue, Color, RichRun, RunStyle, Sheet, SheetId, Workbook};
use xlsx_render::{DisplayList, DrawCmd, GridMeta, Viewport};

use super::{RasterFonts, render_png_with_fonts};
use crate::font;

const WIDTH: f32 = 160.0;
const HEIGHT: f32 = 60.0;
const CARLITO: &[u8] = include_bytes!("../assets/Carlito-Regular.ttf");
const SANS: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const RED_BOLD: &str = r##"{"start":0,"end":2,"bold":true,"color":"#ff0000"}"##;
const RED_BOLD_BODY: &str =
    r##"{"start":0,"end":2,"fontFamily":"Body","bold":true,"color":"#ff0000"}"##;

fn fonts() -> RasterFonts {
    let mut fonts = RasterFonts::default();
    for (family, data) in [("Calibri", CARLITO), ("Body", SANS)] {
        for bold in [false, true] {
            fonts
                .register_font(family, bold, false, data.to_vec(), false, false)
                .unwrap();
        }
    }
    fonts
}

fn measure(fonts: &RasterFonts, family: &str, bold: bool, text: &str) -> f32 {
    measure_at(fonts, family, bold, text, 11.0)
}

fn measure_at(fonts: &RasterFonts, family: &str, bold: bool, text: &str, size: f32) -> f32 {
    font::measure_text_with_fonts(fonts.chain(family, bold, false), &fonts.stack, text, size)
}

fn merged(mut base: Value, extra: Value) -> Value {
    for (key, value) in extra.as_object().unwrap() {
        base[key] = value.clone();
    }
    base
}

fn text(fields: Value) -> DrawCmd {
    let base = json!({
        "op": "text",
        "x": 10.0,
        "y": 20.0,
        "text": "ABCD",
        "fontSize": 11.0,
        "color": "#000000",
        "clip": { "x": 0.0, "y": 0.0, "w": WIDTH, "h": HEIGHT },
        "align": "left",
        "fontFamily": "Calibri",
    });
    serde_json::from_value(merged(base, fields)).unwrap()
}

fn runs(runs: &[&str]) -> Value {
    let runs: Vec<Value> = runs
        .iter()
        .map(|run| serde_json::from_str(run).unwrap())
        .collect();
    json!({ "runs": runs })
}

fn rich(fields: Value, with: &[&str]) -> DrawCmd {
    text(merged(fields, runs(with)))
}

fn fill(x: f32, color: &str) -> DrawCmd {
    DrawCmd::FillRect {
        x,
        y: 0.0,
        w: WIDTH - x,
        h: HEIGHT,
        color: color.into(),
        clip: None,
    }
}

fn png(fonts: &RasterFonts, commands: Vec<DrawCmd>) -> Vec<u8> {
    let list = DisplayList {
        width: WIDTH,
        height: HEIGHT,
        commands,
        grid: GridMeta::default(),
        hyperlinks: Vec::new(),
        charts: Vec::new(),
        truncated: false,
    };
    render_png_with_fonts(&list, fonts).unwrap()
}

fn pixels(fonts: &RasterFonts, commands: Vec<DrawCmd>) -> Pixmap {
    Pixmap::decode_png(&png(fonts, commands)).unwrap()
}

fn has_ink(pixmap: &Pixmap, columns: std::ops::Range<u32>, ink: fn(u8, u8, u8) -> bool) -> bool {
    columns.into_iter().any(|x| {
        (0..pixmap.height()).any(|y| {
            let pixel = pixmap.pixel(x, y).unwrap();
            ink(pixel.red(), pixel.green(), pixel.blue())
        })
    })
}

fn red(r: u8, g: u8, b: u8) -> bool {
    r > 150 && g < 100 && b < 100
}

fn blue(r: u8, g: u8, b: u8) -> bool {
    b > 150 && r < 100 && g < 100
}

fn dark(r: u8, g: u8, b: u8) -> bool {
    r < 100 && g < 100 && b < 100
}

#[test]
fn two_runs_paint_as_adjacent_single_style_commands() {
    let fonts = fonts();
    let ab = measure(&fonts, "Body", true, "AB");
    let cd = measure(&fonts, "Calibri", false, "CD");
    let black = r##"{"start":2,"end":4,"bold":false,"color":"#000000"}"##;
    for (align, x, left) in [
        ("left", 10.0, 10.0),
        ("center", 80.0, 80.0 - (ab + cd) / 2.0),
        ("right", 150.0, 150.0 - (ab + cd)),
    ] {
        let painted = rich(json!({ "x": x, "align": align }), &[RED_BOLD_BODY, black]);
        let pieces = vec![
            text(json!({
                "x": left,
                "text": "AB",
                "fontFamily": "Body",
                "bold": true,
                "color": "#ff0000",
            })),
            text(json!({ "x": left + ab, "text": "CD" })),
        ];
        assert_eq!(png(&fonts, vec![painted]), png(&fonts, pieces), "{align}");
    }
}

#[test]
fn runs_with_the_command_style_paint_as_without_runs() {
    let fonts = fonts();
    let plain = png(&fonts, vec![text(json!({}))]);
    for same in [
        vec![r##"{"start":0,"end":4,"fontFamily":"Calibri","color":"#000000"}"##],
        vec![r#"{"start":0,"end":2}"#, r#"{"start":2,"end":4}"#],
        vec![r#"{"start":1,"end":3,"fontSize":11.0,"bold":false}"#],
    ] {
        assert_eq!(png(&fonts, vec![rich(json!({}), &same)]), plain);
    }
}

#[test]
fn text_outside_every_run_paints_in_the_command_font() {
    let fonts = fonts();
    let ab = measure(&fonts, "Body", true, "AB");
    let painted = rich(json!({ "color": "#0000ff" }), &[RED_BOLD_BODY]);
    let pieces = vec![
        text(json!({ "text": "AB", "fontFamily": "Body", "bold": true, "color": "#ff0000" })),
        text(json!({ "x": 10.0 + ab, "text": "CD", "color": "#0000ff" })),
    ];
    assert_eq!(png(&fonts, vec![painted]), png(&fonts, pieces));
}

#[test]
fn underline_and_strike_apply_per_run() {
    let fonts = fonts();
    let ab = measure(&fonts, "Calibri", false, "AB");
    let painted = rich(
        json!({}),
        &[
            r#"{"start":0,"end":2,"underline":true,"underlineStyle":"double"}"#,
            r#"{"start":2,"end":4,"strike":true}"#,
        ],
    );
    let pieces = vec![
        text(json!({ "text": "AB", "underline": true, "underlineStyle": "double" })),
        text(json!({ "x": 10.0 + ab, "text": "CD", "strike": true })),
    ];
    assert_eq!(png(&fonts, vec![painted]), png(&fonts, pieces));
}

#[test]
fn a_superscript_run_paints_at_its_own_size_on_the_common_baseline() {
    let fonts = fonts();
    let raised = r##"{"start":0,"end":2,"bold":true,"color":"#ff0000","baseline":"superscript"}"##;
    let lowered = r##"{"start":2,"end":4,"color":"#ff0000","baseline":"subscript"}"##;
    let level = r##"{"start":2,"end":4,"color":"#ff0000"}"##;
    assert_eq!(
        png(
            &fonts,
            vec![rich(json!({ "fontFamily": "Body" }), &[raised])]
        ),
        png(
            &fonts,
            vec![rich(json!({ "fontFamily": "Body" }), &[RED_BOLD])]
        )
    );
    assert_eq!(
        png(&fonts, vec![rich(json!({}), &[lowered])]),
        png(&fonts, vec![rich(json!({}), &[level])])
    );
}

#[test]
fn rotated_and_stacked_text_paints_as_without_runs() {
    let fonts = fonts();
    let cell = json!({ "x": 0.0, "y": 0.0, "w": WIDTH, "h": HEIGHT });
    for layout in [
        json!({ "rotationDeg": 90.0, "rotationBox": cell }),
        json!({ "stacked": true, "rotationBox": cell }),
    ] {
        assert_eq!(
            png(&fonts, vec![rich(layout.clone(), &[RED_BOLD_BODY])]),
            png(&fonts, vec![text(layout)])
        );
    }
}

#[test]
fn a_wrapped_command_paints_each_line_s_pieces_in_their_own_fonts() {
    let fonts = fonts();
    let ab = measure(&fonts, "Body", true, "AB");
    let cde = measure_at(&fonts, "Calibri", false, "CDE", 14.0);
    let fg = measure_at(&fonts, "Calibri", false, "FG", 14.0);
    let wrap = json!({ "width": cde + 1.0, "lineHeight": 16.0, "valign": "top" });
    let second = r##"{"start":3,"end":9,"fontSize":14.0,"color":"#0000ff"}"##;
    for (align, x) in [("left", 10.0), ("center", 80.0), ("right", 150.0)] {
        let left = |width: f32| match align {
            "left" => x,
            "center" => x - width / 2.0,
            _ => x - width,
        };
        let layout = json!({ "x": x, "y": 14.0, "align": align, "text": "AB CDE FG" });
        let layout = merged(layout, json!({ "wrap": wrap }));
        let painted = rich(layout, &[RED_BOLD_BODY, second]);
        let piece = |x: f32, y: f32, value: &str, style: Value| {
            text(merged(json!({ "x": x, "y": y, "text": value }), style))
        };
        let body = json!({ "fontFamily": "Body", "bold": true, "color": "#ff0000" });
        let larger = json!({ "fontSize": 14.0, "color": "#0000ff" });
        let pieces = vec![
            piece(left(ab), 14.0, "AB", body),
            piece(left(cde), 30.0, "CDE", larger.clone()),
            piece(left(fg), 46.0, "FG", larger),
        ];
        assert_eq!(png(&fonts, vec![painted]), png(&fonts, pieces), "{align}");
    }
}

fn spill_list(neighbour: bool) -> DisplayList {
    let text = |value: &str| Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Cell::default()
    };
    let mut sheet = Sheet::new("Sheet1");
    sheet.set_cell(CellRef::new(0, 0), text("Sales figures rise"));
    if neighbour {
        sheet.set_cell(CellRef::new(0, 1), text("x"));
    }
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let coloured = |start, end, rgb: &str, bold| {
        let mut style = RunStyle::default();
        style.bold = Some(bold);
        style.color = Some(Color::Rgb(rgb.into()));
        RichRun::new(start, end, style)
    };
    let runs = [
        coloured(0, 5, "#ff0000", true),
        coloured(6, 13, "#0000ff", false),
    ];
    let lookup = |at: CellRef| (at == CellRef::new(0, 0)).then_some(&runs[..]);
    let viewport = Viewport {
        x: 0.0,
        y: 0.0,
        width: WIDTH,
        height: HEIGHT,
    };
    xlsx_render::build_display_list_with_text_runs(&workbook, SheetId(0), &viewport, Some(&lookup))
        .unwrap()
}

fn rich_cell(list: &DisplayList) -> Option<(f32, xlsx_render::Rect, usize)> {
    list.commands.iter().find_map(|command| match command {
        DrawCmd::Text { x, clip, runs, .. } if !runs.is_empty() => Some((*x, *clip, runs.len())),
        _ => None,
    })
}

#[test]
fn a_rich_cell_spills_into_an_empty_neighbour_and_is_cut_at_a_filled_one() {
    let fonts = fonts();
    let (spill, cut) = (spill_list(false), spill_list(true));
    let (spill_cell, cut_cell) = (rich_cell(&spill), rich_cell(&cut));
    assert_eq!(
        (spill_cell.map(|cell| cell.2), cut_cell.map(|cell| cell.2)),
        (Some(2), Some(2))
    );
    let (Some((x, spill_clip, _)), Some((_, cut_clip, _))) = (spill_cell, cut_cell) else {
        unreachable!();
    };
    assert!(spill_clip.w > cut_clip.w, "{spill_clip:?} {cut_clip:?}");
    let edge = (cut_clip.x + cut_clip.w) as u32;
    let spilled = pixels(&fonts, spill.commands);
    let kept = pixels(&fonts, cut.commands);
    assert!(has_ink(&spilled, edge..WIDTH as u32, blue));
    assert!(!has_ink(&kept, edge..WIDTH as u32, red));
    assert!(!has_ink(&kept, edge..WIDTH as u32, blue));
    let sales = measure(&fonts, "Calibri", true, "Sales");
    let before = (x + sales + measure(&fonts, "Calibri", false, " ")).floor() as u32;
    assert!(before < edge, "{before} {edge}");
    for y in 0..kept.height() {
        for x in 0..before {
            assert_eq!(kept.pixel(x, y), spilled.pixel(x, y), "({x}, {y})");
        }
    }
    assert!(has_ink(&kept, 0..before, red));
}

#[test]
fn a_boundary_inside_a_ligature_combining_mark_or_surrogate_pair_paints_all_the_text() {
    let fonts = fonts();
    for (whole, end, first, rest) in [
        ("office", 2, "of", "fice"),
        ("e\u{301}x", 1, "e", "\u{301}x"),
        ("A😀B", 2, "A", "😀B"),
    ] {
        let painted = text(json!({
            "text": whole,
            "runs": [{ "start": 0, "end": end, "color": "#ff0000" }],
        }));
        let pieces = vec![
            text(json!({ "text": first, "color": "#ff0000" })),
            text(json!({ "x": 10.0 + measure(&fonts, "Calibri", false, first), "text": rest })),
        ];
        let ink = pixels(&fonts, vec![painted.clone()]);
        assert_eq!(png(&fonts, vec![painted]), png(&fonts, pieces), "{whole:?}");
        assert!(has_ink(&ink, 0..WIDTH as u32, red), "{whole:?}");
        assert!(has_ink(&ink, 0..WIDTH as u32, dark), "{whole:?}");
    }
}

#[test]
fn right_to_left_text_paints_as_without_runs() {
    let fonts = fonts();
    for whole in ["שלום עולם", "abc אבג 123", "مرحبا"] {
        let plain = json!({ "text": whole });
        assert_eq!(
            png(&fonts, vec![rich(plain.clone(), &[RED_BOLD_BODY])]),
            png(&fonts, vec![text(plain)]),
            "{whole:?}"
        );
    }
}

#[test]
fn a_clip_narrower_than_the_runs_cuts_them_at_its_edge() {
    let fonts = fonts();
    let ab = measure(&fonts, "Body", true, "AB");
    let cd = measure(&fonts, "Calibri", false, "CD");
    let edge = (10.0 + ab + cd / 2.0).floor();
    let blue_run = r##"{"start":2,"end":4,"color":"#0000ff"}"##;
    let clipped = |w: f32| {
        let clip = json!({ "clip": { "x": 0.0, "y": 0.0, "w": w, "h": HEIGHT } });
        rich(clip, &[RED_BOLD_BODY, blue_run])
    };
    let cut = pixels(&fonts, vec![fill(edge, "#cccccc"), clipped(edge)]);
    let whole = pixels(&fonts, vec![fill(edge, "#cccccc"), clipped(WIDTH)]);
    let before = (10.0 + ab).floor() as u32;
    for y in 0..cut.height() {
        for x in edge as u32..cut.width() {
            let pixel = cut.pixel(x, y).unwrap();
            let rgb = (pixel.red(), pixel.green(), pixel.blue());
            assert_eq!(rgb, (0xcc, 0xcc, 0xcc), "({x}, {y})");
        }
        for x in 0..before {
            assert_eq!(cut.pixel(x, y), whole.pixel(x, y), "({x}, {y})");
        }
    }
    assert!(has_ink(&cut, 0..before, red));
    assert!(has_ink(&cut, (10.0 + ab).ceil() as u32..edge as u32, blue));
}
