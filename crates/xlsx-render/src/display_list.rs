//! target-agnostic display list consumed by the canvas and raster backends.
//! coordinates are viewport-local pixels; colors are `#rrggbb` strings.

use ooxml_drawingml::GeometryPathCommand;
use serde::Serialize;

/// horizontal text anchoring within a cell's clip rect.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum Align {
    Left,
    Center,
    Right,
}

/// a rectangle in viewport-local pixels.
#[derive(Debug, Clone, Copy, Default, PartialEq, Serialize)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

/// a single primitive. `op` tags the variant with a stable string discriminant;
/// style fields skip-serialize at their defaults for backward compatibility.
#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(tag = "op", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub enum DrawCmd {
    FillRect {
        x: f32,
        y: f32,
        w: f32,
        h: f32,
        color: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        clip: Option<Rect>,
    },
    Line {
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        width: f32,
        color: String,
        /// `None` = solid; `"dashed"`/`"dotted"`/`"double"` request a backend stroke pattern.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        style: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        clip: Option<Rect>,
    },
    Path {
        commands: Vec<GeometryPathCommand>,
        fill: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        stroke: Option<PathStroke>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        clip: Option<Rect>,
    },
    #[serde(rename_all = "camelCase")]
    Text {
        x: f32,
        y: f32,
        text: String,
        font_size: f32,
        /// resolved font color, `#rrggbb` (a number-format `[Red]` prefix already applied).
        color: String,
        clip: Rect,
        align: Align,
        #[serde(default, skip_serializing_if = "is_false")]
        bold: bool,
        #[serde(default, skip_serializing_if = "is_false")]
        italic: bool,
        #[serde(default, skip_serializing_if = "is_false")]
        underline: bool,
        #[serde(default, skip_serializing_if = "is_false")]
        strike: bool,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        highlight: Option<String>,
        #[serde(default, skip_serializing_if = "is_false")]
        dashed_underline: bool,
        /// font family from the style font; the backend falls back to its default face.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        font_family: Option<String>,
        /// preview text that is not the cell's committed value (a ghost `new`);
        /// excluded from a11y text recovery.
        #[serde(default, skip_serializing_if = "is_false")]
        ghost: bool,
        #[serde(default, skip_serializing_if = "is_false")]
        chart: bool,
    },
}

fn is_false(b: &bool) -> bool {
    !*b
}

#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct PathStroke {
    pub color: String,
    pub width: f32,
}

/// viewport-local grid boundaries for overlay hit-testing. offset vecs are
/// `visible count + 1` long; `offsets[i+1] - offsets[i]` is cell `i`'s span.
#[derive(Debug, Clone, PartialEq, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GridMeta {
    pub start_row: u32,
    pub start_col: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_indices: Option<Vec<u32>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_indices: Option<Vec<u32>>,
    pub row_offsets: Vec<f32>,
    pub col_offsets: Vec<f32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HyperlinkRegion {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_target: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ChartA11yAttrs {
    pub label: String,
    /// the chart could not be drawn; a neutral box occupies its rect instead.
    #[serde(default, skip_serializing_if = "is_false")]
    pub placeholder: bool,
}

/// a full frame for one viewport, sized in pixels; commands are emitted in a
/// fixed order so serialized output is deterministic.
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct DisplayList {
    pub width: f32,
    pub height: f32,
    pub commands: Vec<DrawCmd>,
    pub grid: GridMeta,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hyperlinks: Vec<HyperlinkRegion>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub charts: Vec<ChartA11yAttrs>,
}

/// scale every coordinate, size, stroke width, and font size by `factor`,
/// leaving colors and text alone — the entire 2x/hidpi story.
pub fn scaled(dl: DisplayList, factor: f32) -> DisplayList {
    let commands = dl
        .commands
        .into_iter()
        .map(|c| match c {
            DrawCmd::FillRect {
                x,
                y,
                w,
                h,
                color,
                clip,
            } => DrawCmd::FillRect {
                x: x * factor,
                y: y * factor,
                w: w * factor,
                h: h * factor,
                color,
                clip: clip.map(|clip| scale_rect(clip, factor)),
            },
            DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color,
                style,
                clip,
            } => DrawCmd::Line {
                x1: x1 * factor,
                y1: y1 * factor,
                x2: x2 * factor,
                y2: y2 * factor,
                width: width * factor,
                color,
                style,
                clip: clip.map(|clip| scale_rect(clip, factor)),
            },
            DrawCmd::Path {
                commands,
                fill,
                stroke,
                clip,
            } => DrawCmd::Path {
                commands: commands
                    .into_iter()
                    .map(|command| scale_path_command(command, factor))
                    .collect(),
                fill,
                stroke: stroke.map(|stroke| PathStroke {
                    color: stroke.color,
                    width: stroke.width * factor,
                }),
                clip: clip.map(|clip| scale_rect(clip, factor)),
            },
            DrawCmd::Text {
                x,
                y,
                text,
                font_size,
                color,
                clip,
                align,
                bold,
                italic,
                underline,
                strike,
                highlight,
                dashed_underline,
                font_family,
                ghost,
                chart,
            } => DrawCmd::Text {
                x: x * factor,
                y: y * factor,
                text,
                font_size: font_size * factor,
                color,
                clip: scale_rect(clip, factor),
                align,
                bold,
                italic,
                underline,
                strike,
                highlight,
                dashed_underline,
                font_family,
                ghost,
                chart,
            },
        })
        .collect();

    DisplayList {
        width: dl.width * factor,
        height: dl.height * factor,
        commands,
        grid: GridMeta {
            start_row: dl.grid.start_row,
            start_col: dl.grid.start_col,
            row_indices: dl.grid.row_indices,
            col_indices: dl.grid.col_indices,
            row_offsets: dl
                .grid
                .row_offsets
                .into_iter()
                .map(|v| v * factor)
                .collect(),
            col_offsets: dl
                .grid
                .col_offsets
                .into_iter()
                .map(|v| v * factor)
                .collect(),
        },
        hyperlinks: dl.hyperlinks,
        charts: dl.charts,
    }
}

fn scale_rect(rect: Rect, factor: f32) -> Rect {
    Rect {
        x: rect.x * factor,
        y: rect.y * factor,
        w: rect.w * factor,
        h: rect.h * factor,
    }
}

fn scale_path_command(command: GeometryPathCommand, factor: f32) -> GeometryPathCommand {
    let factor = f64::from(factor);
    match command {
        GeometryPathCommand::Move { x, y } => GeometryPathCommand::Move {
            x: x * factor,
            y: y * factor,
        },
        GeometryPathCommand::Line { x, y } => GeometryPathCommand::Line {
            x: x * factor,
            y: y * factor,
        },
        GeometryPathCommand::Quad { cpx, cpy, x, y } => GeometryPathCommand::Quad {
            cpx: cpx * factor,
            cpy: cpy * factor,
            x: x * factor,
            y: y * factor,
        },
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => GeometryPathCommand::Cubic {
            cp1x: cp1x * factor,
            cp1y: cp1y * factor,
            cp2x: cp2x * factor,
            cp2y: cp2y * factor,
            x: x * factor,
            y: y * factor,
        },
        GeometryPathCommand::Close => GeometryPathCommand::Close,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> DisplayList {
        DisplayList {
            width: 100.0,
            height: 50.0,
            commands: vec![
                DrawCmd::FillRect {
                    x: 1.0,
                    y: 2.0,
                    w: 10.0,
                    h: 20.0,
                    color: "#ffffff".into(),
                    clip: None,
                },
                DrawCmd::Line {
                    x1: 0.0,
                    y1: 4.0,
                    x2: 100.0,
                    y2: 4.0,
                    width: 1.0,
                    color: "#d4d4d4".into(),
                    style: None,
                    clip: None,
                },
                DrawCmd::Path {
                    commands: vec![
                        GeometryPathCommand::Move { x: 1.0, y: 2.0 },
                        GeometryPathCommand::Quad {
                            cpx: 3.0,
                            cpy: 4.0,
                            x: 5.0,
                            y: 6.0,
                        },
                    ],
                    fill: "#123456".into(),
                    stroke: Some(PathStroke {
                        color: "#654321".into(),
                        width: 2.0,
                    }),
                    clip: Some(Rect {
                        x: 0.0,
                        y: 0.0,
                        w: 10.0,
                        h: 10.0,
                    }),
                },
                DrawCmd::Text {
                    x: 3.0,
                    y: 8.0,
                    text: "hi".into(),
                    font_size: 11.0,
                    color: "#000000".into(),
                    clip: Rect {
                        x: 2.0,
                        y: 6.0,
                        w: 12.0,
                        h: 16.0,
                    },
                    align: Align::Left,
                    bold: false,
                    italic: false,
                    underline: false,
                    strike: false,
                    highlight: None,
                    dashed_underline: false,
                    font_family: None,
                    ghost: false,
                    chart: false,
                },
            ],
            grid: GridMeta {
                start_row: 1,
                start_col: 2,
                row_indices: None,
                col_indices: None,
                row_offsets: vec![0.0, 20.0],
                col_offsets: vec![0.0, 64.0],
            },
            hyperlinks: Vec::new(),
            charts: vec![ChartA11yAttrs {
                label: "Revenue chart".into(),
                placeholder: false,
            }],
        }
    }

    #[test]
    fn scaled_multiplies_every_geometry_field() {
        let dl = scaled(sample(), 2.0);
        assert_eq!(dl.width, 200.0);
        assert_eq!(dl.height, 100.0);
        match &dl.commands[0] {
            DrawCmd::FillRect {
                x, y, w, h, color, ..
            } => {
                assert_eq!((*x, *y, *w, *h), (2.0, 4.0, 20.0, 40.0));
                assert_eq!(color, "#ffffff");
            }
            _ => panic!("expected fill rect"),
        }
        match &dl.commands[1] {
            DrawCmd::Line { x2, y2, width, .. } => {
                assert_eq!((*x2, *y2, *width), (200.0, 8.0, 2.0));
            }
            _ => panic!("expected line"),
        }
        match &dl.commands[2] {
            DrawCmd::Path {
                commands,
                stroke,
                clip,
                ..
            } => {
                assert!(matches!(
                    commands.get(1),
                    Some(GeometryPathCommand::Quad {
                        cpx,
                        cpy,
                        x,
                        y
                    }) if (*cpx, *cpy, *x, *y) == (6.0, 8.0, 10.0, 12.0)
                ));
                assert_eq!(stroke.as_ref().map(|stroke| stroke.width), Some(4.0));
                assert_eq!(clip.as_ref().map(|clip| clip.w), Some(20.0));
            }
            _ => panic!("expected path"),
        }
        match &dl.commands[3] {
            DrawCmd::Text {
                x,
                font_size,
                clip,
                text,
                ..
            } => {
                assert_eq!(*x, 6.0);
                assert_eq!(*font_size, 22.0);
                assert_eq!((clip.x, clip.w), (4.0, 24.0));
                assert_eq!(text, "hi");
            }
            _ => panic!("expected text"),
        }
        assert_eq!(dl.grid.start_row, 1);
        assert_eq!(dl.grid.col_offsets, vec![0.0, 128.0]);
        assert_eq!(dl.charts[0].label, "Revenue chart");
    }

    #[test]
    fn scaled_by_one_is_identity() {
        let dl = sample();
        assert_eq!(scaled(dl.clone(), 1.0), dl);
    }
}
