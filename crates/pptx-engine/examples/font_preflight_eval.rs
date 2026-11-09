use std::collections::BTreeSet;
use std::path::Path;
use std::time::Instant;

use serde_json::{Value, json};

fn face_keys(value: &Value) -> BTreeSet<(String, bool, bool)> {
    value
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|face| {
            Some((
                face["family"].as_str()?.trim().to_lowercase(),
                face["bold"].as_bool().unwrap_or(false),
                face["italic"].as_bool().unwrap_or(false),
            ))
        })
        .collect()
}

fn timed<T>(
    operation: impl FnOnce() -> Result<T, String> + std::panic::UnwindSafe,
) -> (Result<T, String>, f64) {
    let started = Instant::now();
    let result = std::panic::catch_unwind(operation).unwrap_or_else(|_| Err("panic".into()));
    (result, started.elapsed().as_secs_f64() * 1000.0)
}

fn engine_faces(bytes: &[u8]) -> Result<Value, String> {
    let document = pptx_engine::Document::open_collaborative(
        bytes,
        1.0,
        pptx_engine::Host {
            entropy: || 1,
            timer_ms: || 0.0,
        },
    )
    .map_err(|_| "open failed")?;
    let ids = document
        .session()
        .slide_ids()
        .map_err(|_| "slides failed")?;
    let id = ids.first().ok_or("missing slide")?;
    let json = pptx_engine::Renderer::new()
        .requested_font_faces_json(document.session(), Some(id))
        .map_err(|_| "font request failed")?;
    let value: Value = serde_json::from_str(&json).map_err(|_| "invalid font response")?;
    Ok(value["faces"].clone())
}

fn evaluate(path: &Path) -> Value {
    let mut output = json!({"file": path.to_string_lossy(), "engine": [], "preflight": {"faces": [], "complete": false}});
    let bytes = match std::fs::read(path) {
        Ok(bytes) => bytes,
        Err(_) => {
            output["error"] = json!("read failed");
            return output;
        }
    };
    let (preflight, preflight_ms) = timed(|| {
        let result = ooxml_font_preflight::pptx_font_preflight(&bytes, 0)
            .map_err(|_| "preflight failed".to_owned())?;
        serde_json::from_str::<Value>(&result.to_json())
            .map_err(|_| "invalid preflight response".to_owned())
    });
    let (engine, engine_ms) = timed(|| engine_faces(&bytes));
    output["preflightMs"] = json!(preflight_ms);
    output["engineMs"] = json!(engine_ms);
    let mut errors = Vec::new();
    match preflight {
        Ok(value) => output["preflight"] = value,
        Err(_) => errors.push("preflight failed"),
    }
    match engine {
        Ok(value) => output["engine"] = value,
        Err(_) => errors.push("engine failed"),
    }
    if errors.is_empty() {
        let expected = face_keys(&output["engine"]);
        let actual = face_keys(&output["preflight"]["faces"]);
        output["missing"] = json!(expected.difference(&actual).collect::<Vec<_>>());
        output["extra"] = json!(actual.difference(&expected).collect::<Vec<_>>());
    } else {
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
            .is_some_and(|extension| extension.eq_ignore_ascii_case("pptx"))
        {
            println!("{}", evaluate(&path));
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    std::panic::set_hook(Box::new(|_| {}));
    let mut args = std::env::args_os().skip(1);
    let directory = args
        .next()
        .ok_or("usage: font_preflight_eval PPTX_DIRECTORY")?;
    if args.next().is_some() {
        return Err("usage: font_preflight_eval PPTX_DIRECTORY".into());
    }
    visit(Path::new(&directory))?;
    Ok(())
}
