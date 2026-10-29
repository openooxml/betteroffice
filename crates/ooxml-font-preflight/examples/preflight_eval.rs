use std::path::Path;

fn visit(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut entries: Vec<_> = std::fs::read_dir(path)?.collect::<Result<_, _>>()?;
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
            let bytes = std::fs::read(&path)?;
            let value = match ooxml_font_preflight::docx_font_preflight(&bytes) {
                Ok(result) => {
                    let mut value: serde_json::Value = serde_json::from_str(&result.to_json())?;
                    value["file"] = serde_json::json!(path.to_string_lossy());
                    value
                }
                Err(error) => {
                    serde_json::json!({"file": path.to_string_lossy(), "faces": [], "complete": false, "error": error.to_string()})
                }
            };
            println!("{value}");
        }
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let directory = std::env::args_os()
        .nth(1)
        .ok_or("expected a DOCX directory")?;
    visit(Path::new(&directory))
}
