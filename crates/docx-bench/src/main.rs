use std::env;
use std::fs;
use std::hint::black_box;
use std::time::Instant;

use betteroffice_docx::Document;
use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position};

const SAMPLES: usize = 20;
const KEYS: usize = 60;

/// Per-keystroke path the wasm session runs: seed once, then insert_text +
/// story re-materialization + state update encode per key.
fn keystroke_loop(bytes: &[u8]) -> Result<f64, Box<dyn std::error::Error>> {
    let document = EditingDoc::new(1);
    docx_edit::seed_from_docx(&document, bytes)?;
    let ctx = EditCtx::local("bench", "");
    let at = document
        .paragraphs("body")?
        .first()
        .map(|p| p.text.encode_utf16().count() as u32)
        .unwrap_or(0)
        .min(6);
    let t = Instant::now();
    for _ in 0..KEYS {
        document.insert_text(&ctx, Position::new("body", at), "x", FormatPolicy::Plain)?;
        black_box(document.story_segments("body")?);
        black_box(document.encode_state_as_update_v1());
    }
    Ok(t.elapsed().as_secs_f64() * 1e3 / KEYS as f64)
}

fn run_pipeline(bytes: &[u8]) -> Result<(f64, f64, f64), Box<dyn std::error::Error>> {
    let t = Instant::now();
    let mut document = Document::open(bytes)?;
    let open_ms = t.elapsed().as_secs_f64() * 1e3;

    let t = Instant::now();
    black_box(document.structure());
    if let Some(id) = document
        .paragraphs()
        .iter()
        .find_map(|paragraph| paragraph.para_id.clone())
    {
        black_box(document.replace_paragraph_text(&id, "throughput probe")?);
    }
    let edit_ms = t.elapsed().as_secs_f64() * 1e3;

    let t = Instant::now();
    let saved = document.save()?;
    black_box(saved.len());
    let save_ms = t.elapsed().as_secs_f64() * 1e3;
    Ok((open_ms, edit_ms, save_ms))
}

fn median(samples: &[(f64, f64, f64)]) -> (f64, f64, f64) {
    let pick = |i: usize| {
        let mut col = samples
            .iter()
            .map(|s| [s.0, s.1, s.2][i])
            .collect::<Vec<_>>();
        col.sort_by(f64::total_cmp);
        col[col.len() / 2]
    };
    (pick(0), pick(1), pick(2))
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("usage: docx-bench PATH.docx");
    let bytes = fs::read(&path)?;
    run_pipeline(&bytes)?;
    let samples = (0..SAMPLES)
        .map(|_| run_pipeline(&bytes))
        .collect::<Result<Vec<_>, _>>()?;
    let (open, edit, save) = median(&samples);
    println!("{path}: open={open:.2}ms edit={edit:.2}ms save={save:.2}ms (median of {SAMPLES})");

    keystroke_loop(&bytes)?;
    let mut keys = (0..SAMPLES)
        .map(|_| keystroke_loop(&bytes))
        .collect::<Result<Vec<_>, _>>()?;
    keys.sort_by(f64::total_cmp);
    println!(
        "{path}: keystroke={:.2}ms/op (median of {SAMPLES} x {KEYS} keys)",
        keys[keys.len() / 2]
    );
    Ok(())
}
