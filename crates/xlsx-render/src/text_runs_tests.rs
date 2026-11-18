use std::sync::Arc;

use xlsx_model::styles::{Alignment, Font, Xf};
use xlsx_model::workbook::{Cell, Hyperlink, Sheet};
use xlsx_model::{
    CellRange, CellRef, CellValue, Color, RichRun, RunBaseline, RunStyle, RunUnderline, SheetId,
    Workbook,
};

use crate::display_list::{DisplayList, DrawCmd, Rect, TextRun, TextRunBaseline};
use crate::{
    Viewport, build_display_list, build_display_list_with_text_runs, run_pieces, wrap_run_lines,
};

const VIEW: Viewport = Viewport {
    x: 0.0,
    y: 0.0,
    width: 400.0,
    height: 100.0,
};

const PLAIN_JSON: &str = concat!(
    r##"{"op":"text","x":2.0,"y":15.066667,"text":"plain","fontSize":11.0,"color":"#000000","##,
    r#""clip":{"x":0.0,"y":0.0,"w":64.0,"h":20.0},"align":"left","fontFamily":"Calibri"}"#,
);

fn text(value: &str) -> Cell {
    Cell {
        value: CellValue::Text {
            value: value.into(),
        },
        ..Cell::default()
    }
}

fn styled(value: &str, style: u32) -> Cell {
    Cell {
        style: Some(style),
        ..text(value)
    }
}

fn at(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

fn colour(rgb: &str) -> Option<Color> {
    Some(Color::Rgb(rgb.into()))
}

fn rich_run(start: u32, end: u32, edit: impl FnOnce(&mut RunStyle)) -> RichRun {
    let mut style = RunStyle::default();
    edit(&mut style);
    RichRun::new(start, end, style)
}

fn workbook(cells: Vec<(&str, Cell)>) -> Workbook {
    let mut sheet = Sheet::new("Sheet1");
    for (address, cell) in cells {
        sheet.set_cell(at(address), cell);
    }
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    workbook
}

fn with_runs(workbook: &Workbook, runs: &[(&str, &[RichRun])]) -> DisplayList {
    let lookup = |cell: CellRef| {
        runs.iter()
            .find(|(address, _)| at(address) == cell)
            .map(|(_, runs)| *runs)
    };
    build_display_list_with_text_runs(workbook, SheetId(0), &VIEW, Some(&lookup)).unwrap()
}

fn command<'a>(list: &'a DisplayList, wanted: &str) -> &'a DrawCmd {
    list.commands
        .iter()
        .find(|command| matches!(command, DrawCmd::Text { text, .. } if &**text == wanted))
        .unwrap_or_else(|| panic!("no text command for {wanted:?}"))
}

fn runs_of<'a>(list: &'a DisplayList, wanted: &str) -> &'a [TextRun] {
    match command(list, wanted) {
        DrawCmd::Text { runs, .. } => runs,
        _ => unreachable!(),
    }
}

fn without_runs(command: &DrawCmd) -> DrawCmd {
    let mut command = command.clone();
    if let DrawCmd::Text { runs, .. } = &mut command {
        runs.clear();
    }
    command
}

#[test]
fn a_rich_cell_carries_runs_with_utf16_offsets_and_styles() {
    let stored = "Ab 中😀z";
    let source = [
        rich_run(0, 2, |style| {
            style.bold = Some(true);
            style.italic = Some(true);
            style.name = Some("Cambria".into());
            style.size_pt = Some(14.0);
        }),
        rich_run(3, 6, |style| {
            style.underline = Some(RunUnderline::Double);
            style.strike = Some(true);
            style.color = Some(Color::Rgb("#ff0000".into()));
        }),
        rich_run(6, 10, |style| {
            style.underline = Some(RunUnderline::Single);
            style.baseline = Some(RunBaseline::Superscript);
        }),
        rich_run(10, 11, |style| {
            style.underline = Some(RunUnderline::None);
            style.baseline = Some(RunBaseline::Subscript);
        }),
    ];
    let workbook = workbook(vec![("A1", text(stored))]);
    let list = with_runs(&workbook, &[("A1", &source)]);
    assert_eq!(
        runs_of(&list, stored),
        [
            TextRun {
                start: 0,
                end: 2,
                font_family: Some("Cambria".into()),
                font_size: Some(14.0),
                bold: Some(true),
                italic: Some(true),
                ..TextRun::default()
            },
            TextRun {
                start: 3,
                end: 4,
                underline: Some(true),
                underline_style: Some(RunUnderline::Double),
                strike: Some(true),
                color: Some("#ff0000".into()),
                ..TextRun::default()
            },
            TextRun {
                start: 4,
                end: 6,
                underline: Some(true),
                baseline: Some(TextRunBaseline::Superscript),
                ..TextRun::default()
            },
            TextRun {
                start: 6,
                end: 7,
                underline: Some(false),
                baseline: Some(TextRunBaseline::Subscript),
                ..TextRun::default()
            }
        ]
    );
    let plain = build_display_list(&workbook, SheetId(0), &VIEW).unwrap();
    assert_eq!(
        without_runs(command(&list, stored)),
        *command(&plain, stored)
    );
}

#[test]
fn a_cell_without_runs_builds_as_before() {
    let workbook = workbook(vec![
        ("A1", text("plain")),
        ("B1", text("x")),
        ("C1", text("rich")),
    ]);
    let none = |_: CellRef| None::<&[RichRun]>;
    let unchanged =
        build_display_list_with_text_runs(&workbook, SheetId(0), &VIEW, Some(&none)).unwrap();
    let plain = build_display_list(&workbook, SheetId(0), &VIEW).unwrap();
    assert_eq!(unchanged, plain);
    assert_eq!(
        serde_json::to_string(command(&plain, "plain")).unwrap(),
        PLAIN_JSON
    );

    let bold = [rich_run(0, 2, |style| style.bold = Some(true))];
    let rich = with_runs(&workbook, &[("C1", &bold)]);
    assert_eq!(
        serde_json::to_string(command(&rich, "plain")).unwrap(),
        PLAIN_JSON
    );
    assert_eq!(rich.commands.len(), plain.commands.len());
    for (built, before) in rich.commands.iter().zip(&plain.commands) {
        assert_eq!(without_runs(built), *before);
    }
    assert_eq!(runs_of(&rich, "rich").len(), 1);
}

#[test]
fn runs_are_dropped_where_the_painted_text_is_not_the_stored_string() {
    let formula = Cell {
        formula: Some("\"ab\"".into()),
        ..styled("ab", 0)
    };
    let mut workbook = workbook(vec![
        ("A1", formula),
        ("C1", styled("ab", 1)),
        ("E1", styled("ab", 2)),
        ("G1", styled("ab", 3)),
        ("A3", styled("ab", 4)),
        ("C3", styled("ab", 5)),
        ("E3", text("aב")),
        ("G3", Cell::default()),
    ]);
    workbook.styles.fonts.push(Font {
        name: Some("Calibri".into()),
        size_pt: Some(11.0),
        ..Font::default()
    });
    workbook.styles.fonts.push(Font {
        name: Some("Symbol".into()),
        size_pt: Some(11.0),
        ..Font::default()
    });
    workbook.styles.num_fmts.push((164, "\"x\"@".into()));
    let rotated = |degrees| Xf {
        font: Some(0),
        alignment: Some(Alignment {
            text_rotation: Some(degrees),
            ..Alignment::default()
        }),
        ..Xf::default()
    };
    workbook.styles.cell_xfs.extend([
        Xf {
            font: Some(0),
            ..Xf::default()
        },
        Xf {
            font: Some(0),
            num_fmt_id: Some(44),
            ..Xf::default()
        },
        Xf {
            font: Some(0),
            num_fmt_id: Some(164),
            ..Xf::default()
        },
        Xf {
            font: Some(1),
            ..Xf::default()
        },
        rotated(45),
        rotated(255),
    ]);
    workbook.sheets[0].hyperlinks.push(Hyperlink {
        range: CellRange::parse_a1("G3").unwrap(),
        external_target: Some("https://example.com".into()),
        location: None,
        tooltip: None,
        display: Some("ab".into()),
    });
    let bold = [rich_run(0, 1, |style| style.bold = Some(true))];
    let every = |_: CellRef| Some(&bold[..]);
    let built =
        build_display_list_with_text_runs(&workbook, SheetId(0), &VIEW, Some(&every)).unwrap();
    let plain = build_display_list(&workbook, SheetId(0), &VIEW).unwrap();
    let texts = built
        .commands
        .iter()
        .filter(|command| matches!(command, DrawCmd::Text { .. }))
        .count();
    assert_eq!(texts, 8);
    assert_eq!(built, plain);
}

#[test]
fn a_hyperlink_cell_keeps_the_link_colour_and_underline() {
    let mut workbook = workbook(vec![("A1", text("Open site"))]);
    workbook.sheets[0].hyperlinks.push(Hyperlink {
        range: CellRange::parse_a1("A1").unwrap(),
        external_target: Some("https://example.com".into()),
        location: None,
        tooltip: None,
        display: None,
    });
    let source = [rich_run(0, 4, |style| {
        style.name = Some("Cambria".into());
        style.size_pt = Some(14.0);
        style.bold = Some(true);
        style.italic = Some(false);
        style.strike = Some(true);
        style.underline = Some(RunUnderline::Double);
        style.color = Some(Color::Rgb("#ff0000".into()));
        style.baseline = Some(RunBaseline::Superscript);
    })];
    let list = with_runs(&workbook, &[("A1", &source)]);
    assert_eq!(
        runs_of(&list, "Open site"),
        [TextRun {
            start: 0,
            end: 4,
            font_family: Some("Cambria".into()),
            font_size: Some(14.0),
            bold: Some(true),
            italic: Some(false),
            strike: Some(true),
            baseline: Some(TextRunBaseline::Superscript),
            ..TextRun::default()
        }]
    );
    let plain = build_display_list(&workbook, SheetId(0), &VIEW).unwrap();
    let DrawCmd::Text {
        color, underline, ..
    } = command(&plain, "Open site")
    else {
        unreachable!();
    };
    assert!(*underline);
    assert_ne!(&**color, "#000000");
    assert_eq!(
        without_runs(command(&list, "Open site")),
        *command(&plain, "Open site")
    );
}

#[test]
fn the_phonetic_tail_lies_outside_every_run() {
    let stored = "漢字かんじ";
    let source = [
        rich_run(0, 3, |style| style.bold = Some(true)),
        rich_run(3, 6, |_| {}),
    ];
    let workbook = workbook(vec![("A1", text(stored))]);
    let list = with_runs(&workbook, &[("A1", &source)]);
    let runs = runs_of(&list, stored);
    assert_eq!(
        runs.iter()
            .map(|run| (run.start, run.end))
            .collect::<Vec<_>>(),
        [(0, 1), (1, 2)]
    );
    let units = stored.encode_utf16().count() as u32;
    assert!(runs.iter().all(|run| run.end <= 2 && run.end < units));
}

#[test]
fn a_rich_cell_spills_into_an_empty_neighbour_and_stops_at_a_filled_one() {
    let stored = "Sales figures rise";
    let source = [
        rich_run(0, 5, |style| {
            style.bold = Some(true);
            style.color = Some(Color::Rgb("#ff0000".into()));
        }),
        rich_run(6, 13, |style| style.color = colour("#0000ff")),
    ];
    let clip_of = |list: &DisplayList| match command(list, stored) {
        DrawCmd::Text { clip, runs, .. } => (*clip, runs.len()),
        _ => unreachable!(),
    };
    let alone = workbook(vec![("A1", text(stored))]);
    let (spilled, runs) = clip_of(&with_runs(&alone, &[("A1", &source)]));
    assert_eq!(runs, 2);
    assert_eq!(
        spilled,
        clip_of(&build_display_list(&alone, SheetId(0), &VIEW).unwrap()).0
    );
    assert!(spilled.w > 64.0, "{spilled:?}");

    let neighboured = workbook(vec![("A1", text(stored)), ("B1", text("x"))]);
    let (kept, runs) = clip_of(&with_runs(&neighboured, &[("A1", &source)]));
    let cell = Rect {
        x: 0.0,
        y: 0.0,
        w: 64.0,
        h: 20.0,
    };
    assert_eq!((kept, runs), (cell, 2));
}

fn wrapped_workbook(size_font: u32, value: &str) -> Workbook {
    let mut workbook = workbook(vec![("A1", styled(value, 1))]);
    workbook.sheets[0].row_heights.insert(0, 60.0);
    for size in [11.0, 20.0] {
        workbook.styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(size),
            ..Font::default()
        });
    }
    workbook.styles.cell_xfs.extend([
        Xf {
            font: Some(0),
            ..Xf::default()
        },
        Xf {
            font: Some(size_font),
            alignment: Some(Alignment {
                wrap_text: true,
                ..Alignment::default()
            }),
            ..Xf::default()
        },
    ]);
    workbook
}

#[test]
fn a_wrapped_rich_cell_steps_its_lines_by_its_largest_run() {
    let stored = "AB CDE FG";
    let source = [
        rich_run(0, 2, |style| style.bold = Some(true)),
        rich_run(3, 9, |style| style.size_pt = Some(20.0)),
    ];
    let layout = |list: &DisplayList| match command(list, stored) {
        DrawCmd::Text {
            y,
            wrap,
            runs,
            font_size,
            ..
        } => (*y, *wrap, runs.len(), *font_size),
        _ => unreachable!(),
    };
    let rich = with_runs(&wrapped_workbook(0, stored), &[("A1", &source)]);
    let rich = layout(&rich);
    let large = wrapped_workbook(1, stored);
    let twin = layout(&build_display_list(&large, SheetId(0), &VIEW).unwrap());
    assert_eq!((rich.2, rich.3), (2, 11.0));
    assert!(rich.1.is_some());
    assert_eq!((rich.0, rich.1), (twin.0, twin.1));

    let small = wrapped_workbook(0, stored);
    let plain = layout(&build_display_list(&small, SheetId(0), &VIEW).unwrap());
    assert_ne!(rich.1, plain.1);
}

#[test]
fn a_single_line_rich_cell_keeps_the_cell_baseline() {
    let stored = "AB CDE FG";
    let source = [rich_run(3, 9, |style| style.size_pt = Some(20.0))];
    let workbook = workbook(vec![("A1", text(stored))]);
    let list = with_runs(&workbook, &[("A1", &source)]);
    assert_eq!(runs_of(&list, stored).len(), 1);
    let plain = build_display_list(&workbook, SheetId(0), &VIEW).unwrap();
    assert_eq!(
        without_runs(command(&list, stored)),
        *command(&plain, stored)
    );
}

#[test]
fn wrapped_pieces_break_inside_the_second_run() {
    let command: DrawCmd = serde_json::from_value(serde_json::json!({
        "op": "text",
        "x": 2.0,
        "y": 14.0,
        "text": "AB CDE FG",
        "fontSize": 10.0,
        "color": "#000000",
        "clip": { "x": 0.0, "y": 0.0, "w": 100.0, "h": 60.0 },
        "align": "left",
        "wrap": { "width": 35.0, "lineHeight": 16.0, "valign": "top" },
        "runs": [
            { "start": 0, "end": 2, "bold": true },
            { "start": 3, "end": 9, "fontSize": 20.0, "color": "#0000ff" },
        ],
    }))
    .unwrap();
    let pieces = run_pieces(&command);
    assert!(pieces.is_some(), "a wrapped command paints its runs");
    let pieces = pieces.unwrap_or_default();
    let DrawCmd::Text { text, .. } = &command else {
        unreachable!();
    };
    let lines = wrap_run_lines(text, &pieces, 35.0, |piece, part| {
        part.len() as f32 * piece.font_size / 2.0
    });
    let laid_out: Vec<Vec<(&str, f32, bool, Arc<str>)>> = lines
        .iter()
        .map(|line| {
            line.iter()
                .map(|piece| {
                    (
                        &text[piece.range.clone()],
                        piece.font_size,
                        piece.bold,
                        piece.color.clone(),
                    )
                })
                .collect()
        })
        .collect();
    let black: Arc<str> = "#000000".into();
    let blue: Arc<str> = "#0000ff".into();
    assert_eq!(
        laid_out,
        [
            vec![("AB", 10.0, true, black)],
            vec![("CDE", 20.0, false, blue.clone())],
            vec![("FG", 20.0, false, blue)]
        ]
    );
}
