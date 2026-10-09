use std::{fs, path::Path};

use serde_json::{Value, json};
use sha2::{Digest, Sha256};

mod format;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

fn record(path: &Path, state: &Value) -> Result<()> {
    let temporary = path.with_extension("tmp");
    fs::write(&temporary, serde_json::to_vec_pretty(state)?)?;
    fs::rename(temporary, path)?;
    Ok(())
}

fn main() -> Result<()> {
    let args: Vec<_> = std::env::args_os().skip(1).collect();
    if args.len() != 4 {
        return Err("usage: roundtrip input probe.json output status.json".into());
    }
    let bytes = fs::read(&args[0])?;
    let status = Path::new(&args[3]);
    let mut state = json!({"source_sha256": format!("{:x}", Sha256::digest(&bytes)), "parse": "failed", "stage": "parse"});
    record(status, &state)?;
    let operation = (|| -> Result<()> {
        let mut document = format::open(&bytes)?;
        state["parse"] = json!("ok");
        state["stage"] = json!("edit");
        record(status, &state)?;
        let probe: Value = serde_json::from_slice(&fs::read(&args[1])?)?;
        if probe.get("status").and_then(Value::as_str) != Some("ok") {
            return Err("No eligible deterministic edit in the source".into());
        }
        format::edit(&mut document, &probe)?;
        state["stage"] = json!("save");
        record(status, &state)?;
        let saved = format::save(&document)?;
        fs::write(&args[2], &saved)?;
        state["output_sha256"] = json!(format!("{:x}", Sha256::digest(&saved)));
        state["stage"] = json!("reopen");
        record(status, &state)?;
        let reopened = format::open(&saved)?;
        format::verify(&reopened, &probe)?;
        state["stage"] = json!("complete");
        state["edit_verified"] = json!(true);
        Ok(())
    })();
    if let Err(error) = operation {
        state["error"] = json!(error.to_string());
    }
    record(status, &state)
}
