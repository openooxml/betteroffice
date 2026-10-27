//! line breaking for wrapped cell text, shared by row autofit and the painters.
//! `packages/xlsx/src/render/wrap.ts` mirrors these rules.

use unicode_segmentation::UnicodeSegmentation;
use xlsx_model::styles::Alignment;

use crate::display_list::{TextValign, TextWrap};

/// whether an alignment wraps its text: wrapText without shrinkToFit, and
/// neither rotated nor stacked.
pub fn wraps(alignment: &Alignment) -> bool {
    alignment.wrap_text
        && !alignment.shrink_to_fit
        && crate::resolve_text_rotation(alignment.text_rotation) == (None, false)
}

/// the lines a wrapped cell paints: `\n` (and `\r\n`) always break, then words
/// wrap greedily at spaces to `width`, and a word wider than `width` breaks
/// between grapheme clusters. Spaces where a line breaks belong to no line;
/// the text's last line keeps its trailing spaces, and a final line break
/// opens no empty line (Excel centres "a\nb\n" as two lines).
pub fn wrap_lines(text: &str, width: f32, mut measure: impl FnMut(&str) -> f32) -> Vec<&str> {
    let width = if width.is_nan() { f32::INFINITY } else { width };
    let text = text
        .strip_suffix('\n')
        .map_or(text, |rest| rest.strip_suffix('\r').unwrap_or(rest));
    let mut lines = Vec::new();
    let mut paragraphs = text.split('\n').peekable();
    while let Some(paragraph) = paragraphs.next() {
        let paragraph = paragraph.strip_suffix('\r').unwrap_or(paragraph);
        let last = paragraphs.peek().is_none();
        wrap_paragraph(paragraph, width, &mut measure, &mut lines);
        if !last && let Some(line) = lines.last_mut() {
            *line = line.trim_end_matches(' ');
        }
    }
    lines
}

fn wrap_paragraph<'a>(
    text: &'a str,
    width: f32,
    measure: &mut impl FnMut(&str) -> f32,
    lines: &mut Vec<&'a str>,
) {
    if measure(text) <= width {
        lines.push(text);
        return;
    }
    let mut start = 0;
    let mut end = None;
    for (word_start, word_end) in words(text) {
        if let Some(line_end) = end {
            if measure(&text[start..word_end]) <= width {
                end = Some(word_end);
                continue;
            }
            lines.push(&text[start..line_end]);
            start = word_start;
        }
        end = Some(word_end);
        if measure(&text[start..word_end]) <= width {
            continue;
        }
        let ends: Vec<usize> = text[start..word_end]
            .grapheme_indices(true)
            .map(|(offset, grapheme)| start + offset + grapheme.len())
            .collect();
        let mut first = 0;
        while let Some(fits) = last_fitting(text, start, &ends, first, width, measure) {
            lines.push(&text[start..ends[fits]]);
            start = ends[fits];
            first = fits + 1;
        }
    }
    lines.push(&text[start..]);
}

/// the last of the graphemes ending at `ends[first..]` that still fits on a
/// line from `start`, or `None` when they all do; a line keeps at least one.
/// A run only widens as it grows, so doubling then bisecting finds the break
/// in logarithmically many measures.
fn last_fitting(
    text: &str,
    start: usize,
    ends: &[usize],
    first: usize,
    width: f32,
    measure: &mut impl FnMut(&str) -> f32,
) -> Option<usize> {
    let mut fits = first;
    let mut step = 1;
    let mut over = loop {
        let probe = (fits + step).min(ends.len() - 1);
        if probe == fits {
            return None;
        }
        if measure(&text[start..ends[probe]]) > width {
            break probe;
        }
        fits = probe;
        step *= 2;
    };
    while over - fits > 1 {
        let mid = fits + (over - fits) / 2;
        if measure(&text[start..ends[mid]]) > width {
            over = mid;
        } else {
            fits = mid;
        }
    }
    Some(fits)
}

/// byte spans of the space-separated words in `text`.
fn words(text: &str) -> impl Iterator<Item = (usize, usize)> + '_ {
    text.split(' ')
        .scan(0, |at, word| {
            let start = *at;
            *at += word.len() + 1;
            Some((start, start + word.len()))
        })
        .filter(|(start, end)| end > start)
}

/// baseline of line `index` of `count` wrapped lines, from the single-line
/// baseline `y` the cell's vertical alignment already placed.
pub fn wrapped_line_y(y: f32, index: usize, count: usize, wrap: &TextWrap) -> f32 {
    let (index, last) = (index as f32, count.saturating_sub(1) as f32);
    y + wrap.line_height
        * match wrap.valign {
            TextValign::Top => index,
            TextValign::Center => index - last / 2.0,
            TextValign::Bottom => index - last,
        }
}

/// lines a wrapped cell takes in a text box `width` px wide, measured with
/// [`estimate_text_width`]: what row autofit counts without a font engine.
pub fn estimated_line_count(text: &str, width: f32, size_pt: f64) -> usize {
    wrap_lines(text, width, |run| estimate_text_width(run, size_pt)).len()
}

/// advance width in px of `text` at `size_pt` in Excel's default face,
/// without kerning. The advances are Calibri's, which Carlito (the raster
/// backend's face) shares glyph for glyph; outside Latin-1, combining and
/// zero-width marks take none, wide East Asian forms a full em, and anything
/// else a digit's advance.
pub fn estimate_text_width(text: &str, size_pt: f64) -> f32 {
    let units: u32 = text.chars().map(advance_units).sum();
    (f64::from(units) * size_pt * crate::geometry::PX_PER_PT / UNITS_PER_EM) as f32
}

const UNITS_PER_EM: f64 = 2048.0;
const DIGIT_ADVANCE: u32 = 1038;

#[rustfmt::skip]
const ASCII_ADVANCES: [u16; 95] = [
    463, 667, 821, 1020, 1038, 1464, 1397, 452, 621, 621, 1020, 1020, 511, 627, 517, 791,
    1038, 1038, 1038, 1038, 1038, 1038, 1038, 1038, 1038, 1038, 548, 548, 1020, 1020, 1020, 949,
    1831, 1185, 1114, 1092, 1260, 1000, 941, 1292, 1276, 516, 653, 1064, 861, 1751, 1322, 1356,
    1058, 1378, 1112, 941, 998, 1314, 1162, 1822, 1063, 998, 959, 628, 791, 628, 1020, 1020,
    596, 981, 1076, 866, 1076, 1019, 625, 964, 1076, 470, 490, 931, 470, 1636, 1076, 1080,
    1076, 1076, 714, 801, 686, 1076, 925, 1464, 887, 927, 809, 644, 943, 644, 1020,
];

#[rustfmt::skip]
const LATIN1_ADVANCES: [u16; 96] = [
    463, 667, 1020, 1038, 1020, 1038, 1020, 1020, 804, 1709, 824, 1049, 1020, 0, 1038, 807,
    694, 1020, 688, 685, 598, 1126, 1200, 517, 629, 504, 865, 1049, 1303, 1375, 1383, 949,
    1185, 1185, 1185, 1185, 1185, 1185, 1563, 1092, 1000, 1000, 1000, 1000, 516, 516, 516, 516,
    1279, 1322, 1356, 1356, 1356, 1356, 1356, 1020, 1359, 1314, 1314, 1314, 1314, 998, 1058, 1080,
    981, 981, 981, 981, 981, 981, 1583, 866, 1019, 1019, 1019, 1019, 470, 470, 470, 470,
    1076, 1076, 1080, 1080, 1080, 1080, 1080, 1020, 1084, 1076, 1076, 1076, 1076, 927, 1076, 927,
];

fn advance_units(c: char) -> u32 {
    let code = c as u32;
    match code {
        0x20..=0x7e => u32::from(ASCII_ADVANCES[(code - 0x20) as usize]),
        0xa0..=0xff => u32::from(LATIN1_ADVANCES[(code - 0xa0) as usize]),
        0x00..=0x1f
        | 0x7f..=0x9f
        | 0x0300..=0x036f
        | 0x200b..=0x200f
        | 0xfe00..=0xfe0f
        | 0xfeff => 0,
        0x1100..=0x115f
        | 0x2e80..=0xa4cf
        | 0xac00..=0xd7a3
        | 0xf900..=0xfaff
        | 0xfe30..=0xfe4f
        | 0xff00..=0xff60
        | 0xffe0..=0xffe6
        | 0x1f300..=0x1faff
        | 0x20000..=0x3fffd => UNITS_PER_EM as u32,
        _ => DIGIT_ADVANCE,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(serde::Deserialize)]
    struct Case {
        text: String,
        width: f32,
        lines: Vec<String>,
    }

    fn graphemes(run: &str) -> f32 {
        run.graphemes(true).count() as f32
    }

    #[test]
    fn breaker_follows_the_shared_rules() {
        let cases: Vec<Case> =
            serde_json::from_str(include_str!("../tests/fixtures/text-wrap-agreement.json"))
                .unwrap();
        for case in cases {
            assert_eq!(
                wrap_lines(&case.text, case.width, graphemes),
                case.lines,
                "{:?} at {}",
                case.text,
                case.width
            );
        }
    }

    #[test]
    fn long_runs_measure_each_text_a_bounded_number_of_times() {
        let size = 11.0;
        let mut calls = 0;
        let mut measured = 0;
        let mut measure = |run: &str| {
            calls += 1;
            measured += run.len();
            estimate_text_width(run, size)
        };
        let unbreakable = format!("aa{}", "\u{200b}".repeat(32_765));
        let width = estimate_text_width("a", size) + 0.5;
        let lines = wrap_lines(&unbreakable, width, &mut measure);
        assert_eq!(lines, [&unbreakable[..1], &unbreakable[1..]]);
        assert!(calls <= 64, "{calls} measures");
        assert!(
            measured <= 8 * unbreakable.len(),
            "{measured} bytes measured"
        );

        let (mut calls, mut measured) = (0, 0);
        let spaced = "a ".repeat(16_383);
        let lines = wrap_lines(&spaced, f32::NAN, |run| {
            calls += 1;
            measured += run.len();
            estimate_text_width(run, size)
        });
        assert_eq!(lines, [spaced.as_str()]);
        assert!(calls <= 64, "{calls} measures");
        assert!(measured <= 8 * spaced.len(), "{measured} bytes measured");
    }

    #[test]
    fn baselines_stack_from_the_single_line_anchor() {
        let wrap = |valign| TextWrap {
            width: 10.0,
            line_height: 18.0,
            valign,
        };
        let ys = |valign| -> Vec<f32> {
            (0..3)
                .map(|i| wrapped_line_y(50.0, i, 3, &wrap(valign)))
                .collect()
        };
        assert_eq!(ys(TextValign::Top), [50.0, 68.0, 86.0]);
        assert_eq!(ys(TextValign::Center), [32.0, 50.0, 68.0]);
        assert_eq!(ys(TextValign::Bottom), [14.0, 32.0, 50.0]);
        assert_eq!(wrapped_line_y(50.0, 0, 1, &wrap(TextValign::Center)), 50.0);
    }

    #[test]
    fn the_estimate_uses_calibri_advances() {
        let em = 11.0 * crate::geometry::PX_PER_PT as f32;
        assert!((estimate_text_width("0", 11.0) - em * 1038.0 / 2048.0).abs() < 1e-4);
        assert!((estimate_text_width("ä", 11.0) - estimate_text_width("a", 11.0)).abs() < 1e-6);
        assert_eq!(
            estimate_text_width("e\u{301}", 11.0),
            estimate_text_width("e", 11.0)
        );
        assert!((estimate_text_width("\u{65e5}", 11.0) - em).abs() < 1e-4);
        assert_eq!(estimated_line_count("a\nb\r\nc", 100.0, 11.0), 3);
        assert_eq!(estimated_line_count("a\nb\r\n", 100.0, 11.0), 2);
    }
}
