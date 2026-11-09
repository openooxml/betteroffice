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
    let document = xlsx_engine::Document::open_with_calculation_json(
        bytes,
        r#"{"nowSerial":45000.75,"randSeed":42}"#,
        xlsx_engine::Host {
            now_serial: || Ok(Some(45000.75)),
            timer_ms: || 0.0,
            progress: || {},
            version: xlsx_engine::Session::version(),
        },
    )
    .map_err(|_| "open failed")?;
    let info: Value =
        serde_json::from_str(&document.sheet_info_json().map_err(|_| "sheet failed")?)
            .map_err(|_| "invalid sheet response")?;
    let corner = json!({
        "sheet": info["activeSheet"],
        "row": ooxml_font_preflight::OPENING_VIEWPORT_ROWS - 1,
        "col": ooxml_font_preflight::OPENING_VIEWPORT_COLUMNS - 1,
    });
    let rect: Value = serde_json::from_str(
        &document
            .cell_rect_json(&corner.to_string())
            .map_err(|_| "cell bounds failed")?,
    )
    .map_err(|_| "invalid bounds response")?;
    let width =
        rect["x"].as_f64().ok_or("missing x")? + rect["w"].as_f64().ok_or("missing width")?;
    let height =
        rect["y"].as_f64().ok_or("missing y")? + rect["h"].as_f64().ok_or("missing height")?;
    let viewport = json!({"x": 0, "y": 0, "width": (width - 0.01).max(0.01), "height": (height - 0.01).max(0.01)});
    let display: Value = serde_json::from_str(
        &document
            .display_list_json(&viewport.to_string())
            .map_err(|_| "display list failed")?,
    )
    .map_err(|_| "invalid display list")?;
    let mut faces = std::collections::BTreeMap::new();
    for command in display["commands"].as_array().into_iter().flatten() {
        if command["op"] != "text" || command["chart"] == true {
            continue;
        }
        let Some(family) = command["fontFamily"]
            .as_str()
            .filter(|family| !family.is_empty())
        else {
            continue;
        };
        let bold = command["bold"].as_bool().unwrap_or(false);
        let italic = command["italic"].as_bool().unwrap_or(false);
        let key = (family.trim().to_lowercase(), bold, italic);
        let face = faces.entry(key).or_insert_with(
            || json!({"family": family, "bold": bold, "italic": italic, "scripts": []}),
        );
        let scripts = ooxml_text::paragraph_scripts_with_languages(
            [(command["text"].as_str().unwrap_or_default(), None)],
            false,
        );
        for script in scripts {
            if let Some(values) = face["scripts"].as_array_mut()
                && !values
                    .iter()
                    .any(|value| value.as_str() == Some(script.as_str()))
            {
                values.push(json!(script));
            }
        }
    }
    Ok(Value::Array(faces.into_values().collect()))
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
        let result = ooxml_font_preflight::xlsx_font_preflight(&bytes)
            .map_err(|error| format!("preflight: {error}"))?;
        serde_json::from_str::<Value>(&result.to_json())
            .map_err(|_| "invalid preflight response".to_owned())
    });
    let (engine, engine_ms) = timed(|| engine_faces(&bytes));
    output["preflightMs"] = json!(preflight_ms);
    output["engineMs"] = json!(engine_ms);
    let mut errors = Vec::new();
    match preflight {
        Ok(value) => output["preflight"] = value,
        Err(message) => errors.push(message),
    }
    match engine {
        Ok(value) => output["engine"] = value,
        Err(_) => errors.push("engine failed".to_owned()),
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
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
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
        .ok_or("usage: font_preflight_eval XLSX_DIRECTORY")?;
    if args.next().is_some() {
        return Err("usage: font_preflight_eval XLSX_DIRECTORY".into());
    }
    visit(Path::new(&directory))?;
    Ok(())
}
