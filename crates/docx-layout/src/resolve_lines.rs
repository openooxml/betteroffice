//! Run slicing with UTF-16 code-unit offsets.

use crate::types::{Run, TextRun, TypesetRow};
use crate::utf16_index::RunTexts;
use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
pub struct ResolvedSegment {
    pub run: Run,
    pub text: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct ResolvedLine {
    pub segments: Vec<ResolvedSegment>,
}

/// Returns a string's UTF-16 code-unit length.
pub fn utf16_len(s: &str) -> usize {
    s.encode_utf16().count()
}

/// Slices a string by UTF-16 code-unit offsets.
fn utf16_slice(s: &str, start: usize, end: usize) -> String {
    let units: Vec<u16> = s.encode_utf16().collect();
    let start = start.min(units.len());
    let end = end.clamp(start, units.len());
    String::from_utf16_lossy(&units[start..end])
}

/// Resolves the visible run segments for one measured line.
pub fn resolve_line_segments(runs: &[Run], line: &TypesetRow) -> Vec<ResolvedSegment> {
    resolve_line_segments_in(runs, line, &RunTexts::new(runs.len()))
}

/// [`resolve_line_segments`] with the paragraph's run offsets shared across
/// its lines.
pub(crate) fn resolve_line_segments_in<'a>(
    runs: &'a [Run],
    line: &TypesetRow,
    texts: &RunTexts<'a>,
) -> Vec<ResolvedSegment> {
    #[cfg(any(test, feature = "test-support"))]
    crate::work_counters::resolved_line();
    let mut segments: Vec<ResolvedSegment> = Vec::new();

    for run_index in line.head_run..=line.tail_run {
        let Some(run) = runs.get(run_index) else {
            continue;
        };

        if let Run::Text(text_run) = run {
            let index = texts.get(run_index, &text_run.text);
            let text_len = index.len();
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
                let start = start_char.min(text_len);
                let end = end_char.clamp(start, text_len);
                let text = match (index.exact_byte(start), index.exact_byte(end)) {
                    (Some(low), Some(high)) => text_run.text[low..high].to_string(),
                    _ => utf16_slice(&text_run.text, start_char, end_char),
                };
                // Both offsets are relative to the run start.
                let sliced = TextRun {
                    fmt: text_run.fmt.clone(),
                    text: text.clone(),
                    pm_start: text_run.pm_start.map(|pm| pm + start_char as f64),
                    pm_end: text_run.pm_start.map(|pm| pm + end_char as f64),
                    inline_sdt_widget: text_run.inline_sdt_widget.clone(),
                };
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::utf16_index::generated::{Rng, lines, runs};

    fn old_resolve_line_segments(runs: &[Run], line: &TypesetRow) -> Vec<ResolvedSegment> {
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

    #[test]
    fn shared_offsets_resolve_every_line_exactly_as_a_fresh_scan() {
        let mut rng = Rng::new(0xdecade);
        for _ in 0..1500 {
            let runs: Vec<Run> = runs(&mut rng)
                .into_iter()
                .map(|run| serde_json::from_value(run).unwrap())
                .collect();
            let units: Vec<usize> = runs
                .iter()
                .map(|run| match run {
                    Run::Text(text) => utf16_len(&text.text),
                    _ => 0,
                })
                .collect();
            let texts = RunTexts::new(runs.len());
            for (head_run, head_char, tail_run, tail_char) in lines(&mut rng, &units) {
                let line = TypesetRow {
                    head_run,
                    head_char,
                    tail_run,
                    tail_char,
                    ..TypesetRow::default()
                };
                let expected =
                    serde_json::to_value(old_resolve_line_segments(&runs, &line)).unwrap();
                let actual =
                    serde_json::to_value(resolve_line_segments_in(&runs, &line, &texts)).unwrap();
                assert_eq!(
                    actual, expected,
                    "line {head_run}:{head_char}..{tail_run}:{tail_char}"
                );
            }
        }
    }
}
