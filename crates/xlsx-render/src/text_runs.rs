//! A text command's runs as the pieces a single-line painter draws.

use std::ops::Range;
use std::sync::Arc;

use xlsx_model::RunUnderline;

use crate::display_list::{DrawCmd, TextRun, TextRunBaseline};

/// A span of a text command's text, as byte offsets, with every font
/// property resolved against the command's own.
#[derive(Debug, Clone, PartialEq)]
pub struct TextPiece {
    pub range: Range<usize>,
    pub font_family: Option<Arc<str>>,
    pub font_size: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// The kind of `underline` when it is not single.
    pub underline_style: Option<RunUnderline>,
    pub strike: bool,
    pub color: Arc<str>,
    /// Carried for the painters; no painter raises or lowers a piece yet.
    pub baseline: Option<TextRunBaseline>,
}

impl TextPiece {
    fn same_style(&self, other: &Self) -> bool {
        self.font_family == other.font_family
            && self.font_size == other.font_size
            && self.bold == other.bold
            && self.italic == other.italic
            && self.underline == other.underline
            && self.underline_style == other.underline_style
            && self.strike == other.strike
            && self.color == other.color
            && self.baseline == other.baseline
    }
}

/// The pieces a painter draws for `command` left to right on one baseline, or
/// `None` when it paints the command as one plain run: a command without
/// runs, runs that all carry the command's own font, and the cases the pieces
/// do not cover yet (wrapped, rotated, stacked, number-formatted or
/// highlighted text, and text with a right-to-left character). Neighbouring
/// pieces in the same font join into one.
pub fn run_pieces(command: &DrawCmd) -> Option<Vec<TextPiece>> {
    let DrawCmd::Text {
        text,
        font_size,
        color,
        bold,
        italic,
        underline,
        underline_style,
        strike,
        highlight,
        dashed_underline,
        font_family,
        number_format,
        rotation_deg,
        stacked,
        wrap,
        runs,
        ..
    } = command
    else {
        return None;
    };
    if runs.is_empty()
        || wrap.is_some()
        || rotation_deg.is_some()
        || *stacked
        || number_format.is_some()
        || highlight.is_some()
        || *dashed_underline
        || has_right_to_left(text)
    {
        return None;
    }
    let base = TextPiece {
        range: 0..0,
        font_family: font_family.clone(),
        font_size: *font_size,
        bold: *bold,
        italic: *italic,
        underline: *underline,
        underline_style: underline_kind(*underline, *underline_style),
        strike: *strike,
        color: color.clone(),
        baseline: None,
    };
    let mut pieces = Vec::new();
    let mut cursor = 0;
    for run in runs {
        let range = run.byte_range(text);
        let start = range.start.max(cursor);
        if range.end <= start {
            continue;
        }
        push(&mut pieces, cursor..start, &base);
        push(&mut pieces, start..range.end, &styled(&base, run));
        cursor = range.end;
    }
    push(&mut pieces, cursor..text.len(), &base);
    match pieces.as_slice() {
        [] => None,
        [only] if only.same_style(&base) => None,
        _ => Some(pieces),
    }
}

/// Whether `text` holds a right-to-left letter (Hebrew, Arabic, Syriac,
/// Thaana, N'Ko, Samaritan, Mandaic and their presentation forms) or a
/// control that opens right-to-left text.
pub fn has_right_to_left(text: &str) -> bool {
    text.chars().any(|ch| {
        matches!(
            ch as u32,
            0x0590..=0x08ff
                | 0xfb1d..=0xfdff
                | 0xfe70..=0xfefc
                | 0x1_0800..=0x1_0fff
                | 0x1_e800..=0x1_efff
                | 0x200f
                | 0x202b
                | 0x202e
                | 0x2067
        )
    })
}

fn push(pieces: &mut Vec<TextPiece>, range: Range<usize>, style: &TextPiece) {
    if range.is_empty() {
        return;
    }
    if let Some(last) = pieces.last_mut()
        && last.same_style(style)
    {
        last.range.end = range.end;
        return;
    }
    pieces.push(TextPiece {
        range,
        ..style.clone()
    });
}

fn styled(base: &TextPiece, run: &TextRun) -> TextPiece {
    let underline = run.underline.unwrap_or(base.underline);
    TextPiece {
        range: 0..0,
        font_family: run.font_family.clone().or_else(|| base.font_family.clone()),
        font_size: run.font_size.unwrap_or(base.font_size),
        bold: run.bold.unwrap_or(base.bold),
        italic: run.italic.unwrap_or(base.italic),
        underline,
        underline_style: underline_kind(underline, run.underline_style.or(base.underline_style)),
        strike: run.strike.unwrap_or(base.strike),
        color: run.color.clone().unwrap_or_else(|| base.color.clone()),
        baseline: run.baseline,
    }
}

fn underline_kind(underline: bool, style: Option<RunUnderline>) -> Option<RunUnderline> {
    match style {
        Some(RunUnderline::None | RunUnderline::Single) => None,
        style if underline => style,
        _ => None,
    }
}
