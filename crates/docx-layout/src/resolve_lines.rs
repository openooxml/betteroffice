//! Canonical run slicing for a laid-out paragraph line. Port of
//! `packages/core/src/layout/pagination/resolveLineSegments.ts`.
//!
//! A line is stored compactly as a `{ headRun, headChar, tailRun, tailChar }`
//! tuple into the paragraph's run array; this is the single place that turns
//! that tuple into the line's visible run segments. Character offsets are
//! UTF-16 code-unit indices (JS string semantics), so slicing goes through a
//! UTF-16 view rather than Rust byte/char indices.

use crate::types::{Run, TypesetRow};
use serde::Serialize;

/// TS `ResolvedSegment` — one run's visible slice on a laid-out line.
#[derive(Debug, Clone, Serialize)]
pub struct ResolvedSegment {
    pub run: Run,
    pub text: String,
}

/// TS `ResolvedLine` — a laid-out line's resolved run segments.
#[derive(Debug, Clone, Serialize)]
pub struct ResolvedLine {
    pub segments: Vec<ResolvedSegment>,
}

/// Length of a string in UTF-16 code units (JS `String#length`).
pub fn utf16_len(s: &str) -> usize {
    s.encode_utf16().count()
}

/// JS `String#slice(start, end)` over UTF-16 code units (indices already
/// clamped non-negative by construction).
fn utf16_slice(s: &str, start: usize, end: usize) -> String {
    let units: Vec<u16> = s.encode_utf16().collect();
    let start = start.min(units.len());
    let end = end.clamp(start, units.len());
    String::from_utf16_lossy(&units[start..end])
}

/// TS `resolveLineSegments` — slice a paragraph's runs to one line's visible
/// span. A boundary text run becomes a copy sliced to the line's head/tail
/// char with PM positions shifted to match; tabs, images, line breaks, and
/// fields pass through whole with an empty `text`.
pub fn resolve_line_segments(runs: &[Run], line: &TypesetRow) -> Vec<ResolvedSegment> {
    let mut segments: Vec<ResolvedSegment> = Vec::new();

    for run_index in line.head_run..=line.tail_run {
        let Some(run) = runs.get(run_index) else {
            continue;
        };

        if let Run::Text(text_run) = run {
            let text_len = utf16_len(&text_run.text);
            let start_char = if run_index == line.head_run {
                line.head_char
            } else {
                0
            };
            let end_char = if run_index == line.tail_run {
                line.tail_char
            } else {
                text_len
            };

            if start_char > 0 || end_char < text_len {
                let text = utf16_slice(&text_run.text, start_char, end_char);
                let mut sliced = text_run.clone();
                sliced.text = text.clone();
                // NOTE: like the TS, BOTH shifted positions key off pmStart.
                sliced.pm_start = text_run.pm_start.map(|pm| pm + start_char as f64);
                sliced.pm_end = text_run.pm_start.map(|pm| pm + end_char as f64);
                segments.push(ResolvedSegment {
                    run: Run::Text(sliced),
                    text,
                });
            } else {
                segments.push(ResolvedSegment {
                    run: run.clone(),
                    text: text_run.text.clone(),
                });
            }
        } else {
            segments.push(ResolvedSegment {
                run: run.clone(),
                text: String::new(),
            });
        }
    }

    segments
}
