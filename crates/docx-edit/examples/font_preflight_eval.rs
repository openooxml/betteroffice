use std::path::Path;
use std::time::Instant;

use docx_edit::{EngineSession, StoryInfoKind, seed_from_docx};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

#[derive(Deserialize, Serialize)]
struct Face {
    family: String,
    bold: bool,
    italic: bool,
    #[serde(default)]
    scripts: Vec<String>,
}

fn engine_faces(bytes: &[u8]) -> Result<Vec<Face>, String> {
    let engine = EngineSession::new(1);
    seed_from_docx(engine.doc(), bytes)?;
    let sections: Vec<_> = engine
        .doc()
        .list_sections()?
        .into_iter()
        .map(|section| {
            json!({
                "sectionId": section.closing_paragraph_id,
                "properties": section.authored,
            })
        })
        .collect();
    let mut notes = Vec::new();
    for story in engine.doc().list_stories().stories {
        let (prefix, kind) = match story.kind {
            StoryInfoKind::Footnote => ("fn:", "footnote"),
            StoryInfoKind::Endnote => ("en:", "endnote"),
            _ => continue,
        };
        let id = story
            .story
            .strip_prefix(prefix)
            .ok_or_else(|| format!("invalid note story {}", story.story))?
            .parse::<i64>()
            .map_err(|error| error.to_string())?;
        notes.push(json!({"id": id, "noteKind": kind, "height": 0}));
    }
    let parts = ooxml_opc::unzip_parts_where(bytes, u64::MAX, |name| name == "word/settings.xml")?;
    let limits = docx_parse::xml::ParseLimits::default();
    let settings = docx_parse::parse_settings(
        parts.first().map(|(_, bytes)| bytes.as_slice()),
        "word/settings.xml",
        &mut docx_parse::xml::ParseBudget::new(&limits),
    )
    .map_err(|error| error.to_string())?;
    let request = json!({
        "bodyStory": "body",
        "regions": {"sections": sections, "settings": settings},
        "notes": {"contents": notes},
        "renderEnv": {},
    });
    let requirements = engine.layout_font_requirements_json(&request.to_string())?;
    serde_json::from_str(&requirements).map_err(|error| error.to_string())
}

fn timed<T>(
    operation: impl FnOnce() -> Result<T, String> + std::panic::UnwindSafe,
) -> (Result<T, String>, f64) {
    let started = Instant::now();
    let result = std::panic::catch_unwind(operation).unwrap_or_else(|panic| {
        let message = panic
            .downcast_ref::<String>()
            .cloned()
            .or_else(|| {
                panic
                    .downcast_ref::<&str>()
                    .map(|message| (*message).to_owned())
            })
            .unwrap_or_else(|| "unknown panic".to_owned());
        Err(format!("panic: {message}"))
    });
    (result, started.elapsed().as_secs_f64() * 1000.0)
}

fn evaluate(path: &Path) -> Value {
    let mut output = json!({
        "file": path.to_string_lossy(),
        "engine": [],
        "preflight": {"faces": [], "complete": false},
        "engineMs": 0.0,
        "preflightMs": 0.0,
    });
    let bytes = match std::fs::read(path) {
        Ok(bytes) => bytes,
        Err(error) => {
            output["error"] = json!(format!("read: {error}"));
            return output;
        }
    };
    let (preflight, preflight_ms) = timed(|| {
        let result =
            ooxml_font_preflight::docx_font_preflight(&bytes).map_err(|error| error.to_string())?;
        serde_json::from_str::<Value>(&result.to_json()).map_err(|error| error.to_string())
    });
    let (engine, engine_ms) = timed(|| engine_faces(&bytes));
    output["engineMs"] = json!(engine_ms);
    output["preflightMs"] = json!(preflight_ms);
    let mut errors = Vec::new();
    match preflight {
        Ok(result) => output["preflight"] = result,
        Err(error) => errors.push(format!("preflight: {error}")),
    }
    match engine {
        Ok(faces) => output["engine"] = json!(faces),
        Err(error) => errors.push(format!("engine: {error}")),
    }
    if !errors.is_empty() {
        output["error"] = json!(errors.join("; "));
    }
    output
}

fn visit(directory: &Path) -> std::io::Result<()> {
    let mut entries: Vec<_> = std::fs::read_dir(directory)?.collect::<Result<_, _>>()?;
    entries.sort_by_key(|entry| entry.path());
    for entry in entries {
        let path = entry.path();
        if entry.file_type()?.is_dir() {
            visit(&path)?;
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
        {
            println!("{}", evaluate(&path));
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args_os().skip(1);
    let directory = args
        .next()
        .ok_or("usage: font_preflight_eval DOCX_DIRECTORY")?;
    if args.next().is_some() {
        return Err("usage: font_preflight_eval DOCX_DIRECTORY".into());
    }
    visit(Path::new(&directory))?;
    Ok(())
}
