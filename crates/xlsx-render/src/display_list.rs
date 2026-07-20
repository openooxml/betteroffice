//! target-agnostic display list consumed by the canvas and raster backends.
//! coordinates are viewport-local pixels; colors are `#rrggbb` strings.

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
#[derive(Debug, Clone, Copy, PartialEq, Serialize)]
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
    },
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
        /// font family from the style font; the backend falls back to its default face.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        font_family: Option<String>,
    },
}

fn is_false(b: &bool) -> bool {
    !*b
}

/// viewport-local grid boundaries for overlay hit-testing. offset vecs are
/// `visible count + 1` long; `offsets[i+1] - offsets[i]` is cell `i`'s span.
#[derive(Debug, Clone, PartialEq, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GridMeta {
    pub start_row: u32,
    pub start_col: u32,
    pub row_offsets: Vec<f32>,
    pub col_offsets: Vec<f32>,
}

/// a full frame for one viewport, sized in pixels; commands are emitted in a
/// fixed order so serialized output is deterministic.
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct DisplayList {
    pub width: f32,
    pub height: f32,
    pub commands: Vec<DrawCmd>,
    pub grid: GridMeta,
}

/// scale every coordinate, size, stroke width, and font size by `factor`,
/// leaving colors and text alone — the entire 2x/hidpi story.
pub fn scaled(dl: DisplayList, factor: f32) -> DisplayList {
    let commands = dl
        .commands
        .into_iter()
        .map(|c| match c {
            DrawCmd::FillRect { x, y, w, h, color } => DrawCmd::FillRect {
                x: x * factor,
                y: y * factor,
                w: w * factor,
                h: h * factor,
                color,
            },
            DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color,
                style,
            } => DrawCmd::Line {
                x1: x1 * factor,
                y1: y1 * factor,
                x2: x2 * factor,
                y2: y2 * factor,
                width: width * factor,
                color,
                style,
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
                font_family,
            } => DrawCmd::Text {
                x: x * factor,
                y: y * factor,
                text,
                font_size: font_size * factor,
                color,
                clip: Rect {
                    x: clip.x * factor,
                    y: clip.y * factor,
                    w: clip.w * factor,
                    h: clip.h * factor,
                },
                align,
                bold,
                italic,
                underline,
                strike,
                font_family,
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
                },
                DrawCmd::Line {
                    x1: 0.0,
                    y1: 4.0,
                    x2: 100.0,
                    y2: 4.0,
                    width: 1.0,
                    color: "#d4d4d4".into(),
                    style: None,
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
                    font_family: None,
                },
            ],
            grid: GridMeta {
                start_row: 1,
                start_col: 2,
                row_offsets: vec![0.0, 20.0],
                col_offsets: vec![0.0, 64.0],
            },
        }
    }

    #[test]
    fn scaled_multiplies_every_geometry_field() {
        let dl = scaled(sample(), 2.0);
        assert_eq!(dl.width, 200.0);
        assert_eq!(dl.height, 100.0);
        match &dl.commands[0] {
            DrawCmd::FillRect { x, y, w, h, color } => {
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
    }

    #[test]
    fn scaled_by_one_is_identity() {
        let dl = sample();
        assert_eq!(scaled(dl.clone(), 1.0), dl);
    }
}
