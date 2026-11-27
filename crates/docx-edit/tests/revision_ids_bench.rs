#![cfg(not(target_arch = "wasm32"))]

use std::collections::BTreeSet;
use std::hint::black_box;
use std::time::Instant;

use docx_edit::save::revision_ids::assign_session_revision_ids_with_timings;
use docx_parse::S13SaveRequest;
use serde_json::{Value, json};

const FIRST: u64 = (1 << 53) - 1;

fn run(text: &str) -> Value {
    json!({"type": "run", "content": [{"type": "text", "text": text}]})
}

fn insertion(id: u64, author: &str, text: &str) -> Value {
    json!({"type": "insertion", "info": {"id": id, "author": author}, "content": [run(text)]})
}

fn paragraph(content: Vec<Value>) -> Value {
    json!({"type": "paragraph", "content": content})
}

fn request(paragraphs: u64) -> (S13SaveRequest, BTreeSet<u64>) {
    let mut request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2030-01-02T03:04:05.006Z"},
        "document": {"content": []},
        "headerEntries": [], "footerEntries": [], "footnotes": [], "endnotes": [],
        "footnoteSeparators": [], "endnoteSeparators": [], "relationshipEntries": [],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let content = &mut request.document.content;
    content.reserve_exact(paragraphs as usize + 8);
    let mut push = |block| content.push(serde_json::from_value(block).unwrap());
    let mut session = BTreeSet::new();
    for index in 0..paragraphs {
        let mut runs = vec![
            run(&format!("Record {index}: ")),
            run("The operations team records each inspection and reviews the results."),
        ];
        if index % 10 == 0 {
            let id = FIRST - index;
            session.insert(id);
            runs.push(insertion(id, "Reviewer", "Follow-up inspection."));
        }
        if index % 50 == 5 {
            runs.push(insertion(index + 1, "Source", "Earlier inspection."));
        }
        push(paragraph(runs));
    }
    for index in 0..4 {
        let cell_id = FIRST - paragraphs - 2 * index;
        let box_id = cell_id - 1;
        session.extend([cell_id, box_id]);
        push(json!({
            "type": "table", "columnWidths": [2000],
            "rows": [{"type": "tableRow", "cells": [{"type": "tableCell", "content": [
                paragraph(vec![run("Item"), insertion(cell_id, "Reviewer", " Inspected")])
            ]}]}]
        }));
        push(paragraph(vec![json!({
            "type": "run", "content": [{"type": "shape", "shape": {
                "type": "shape", "shapeType": "rect", "textBox": true,
                "size": {"width": 914400, "height": 457200},
                "textBody": {"content": [
                    paragraph(vec![run("Note"), insertion(box_id, "Reviewer", " Reviewed")])
                ]}
            }}]
        })]));
    }
    (request, session)
}

#[test]
#[ignore]
fn revision_id_assignment_at_scale() {
    for paragraphs in [5_000, 50_000, 200_000] {
        let mut times = [0.0; 5];
        let mut pass_times = [[0.0; 5]; 3];
        let mut sites = 0;
        for (index, time) in times.iter_mut().enumerate() {
            let (mut sample, session) = request(paragraphs);
            let started = Instant::now();
            let result = assign_session_revision_ids_with_timings(&mut sample, &session);
            *time = started.elapsed().as_secs_f64() * 1_000.0;
            let (ids, passes) = black_box(result.unwrap());
            sites = session.len();
            assert_eq!(ids.len(), sites);
            for (samples, elapsed) in pass_times.iter_mut().zip(passes) {
                samples[index] = elapsed.as_secs_f64() * 1_000.0;
            }
        }
        times.sort_by(f64::total_cmp);
        let [collect_ms, assign_ms, write_ms] = pass_times.map(|mut samples| {
            samples.sort_by(f64::total_cmp);
            samples[2]
        });
        println!(
            "REVISION-ID-BENCH native paragraphs={paragraphs} sites={sites} median_ms={:.3} min_ms={:.3} max_ms={:.3}",
            times[2], times[0], times[4],
        );
        println!(
            "REVISION-ID-BENCH-PASS native paragraphs={paragraphs} collect_ms={collect_ms:.3} assign_ms={assign_ms:.3} write_ms={write_ms:.3}",
        );
    }
}
