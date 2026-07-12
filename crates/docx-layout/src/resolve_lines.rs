use crate::types::{Run, TypesetRow};
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

pub fn utf16_len(s: &str) -> usize {
    s.encode_utf16().count()
}

fn utf16_slice(s: &str, start: usize, end: usize) -> String {
    let units: Vec<u16> = s.encode_utf16().collect();
    let start = start.min(units.len());
    let end = end.clamp(start, units.len());
    String::from_utf16_lossy(&units[start..end])
}

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
