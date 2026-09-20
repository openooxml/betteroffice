use std::collections::BTreeSet;
use std::fs;
use std::path::Path;

use betteroffice_pptx::{Presentation, RenderOptions};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

fn main() {
    if let Err(error) = run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args_os().skip(1).collect();
    if args.len() != 3 {
        return Err("usage: pptx-native-bench input.pptx slide.png fonts/manifest.json".into());
    }
    let bytes = fs::read(&args[0])?;
    let mut deck = Presentation::open(&bytes)?;
    let manifest: Value = serde_json::from_slice(&fs::read(&args[2])?)?;
    let root = Path::new(&args[2])
        .parent()
        .ok_or("font manifest has no parent")?;
    let mut faces = manifest["faces"]
        .as_array()
        .ok_or("font manifest has no faces")?
        .clone();
    faces.sort_by_key(|face| {
        !(face["family"] == "Liberation Sans" && face["weight"] == 400 && face["style"] == "normal")
    });
    for face in faces {
        let file = face["file"].as_str().ok_or("font has no file")?;
        if Path::new(file).components().count() != 1 {
            return Err("invalid font filename".into());
        }
        let font = fs::read(root.join(file))?;
        if format!("{:x}", Sha256::digest(&font)) != face["sha256"].as_str().unwrap_or("") {
            return Err("font hash mismatch".into());
        }
        let family = face["family"].as_str().ok_or("font has no family")?;
        let mut names = BTreeSet::from([family.to_owned()]);
        if let Some(compatible) = face["metricCompatWith"].as_str() {
            names.insert(compatible.to_owned());
        }
        for (alias, target) in manifest["aliases"]
            .as_object()
            .ok_or("missing font aliases")?
        {
            if target
                .as_str()
                .is_some_and(|target| target.eq_ignore_ascii_case(family))
            {
                names.insert(alias.clone());
            }
        }
        if family == "Carlito" {
            names.insert("Calibri Light".to_owned());
        }
        for name in names {
            deck.register_font(
                &name,
                face["weight"] == 700,
                face["style"] == "italic",
                &font,
            )?;
        }
    }
    let rendered = deck.render_png(0, &RenderOptions::default())?;
    if rendered.skipped_images != 0 {
        return Err(format!("renderer skipped {} images", rendered.skipped_images).into());
    }
    fs::write(&args[1], &rendered.bytes)?;
    println!(
        "{}",
        json!({
            "schema_version": 1, "status": "ok", "page": 1, "pages": deck.slides().len(), "dpi": 96,
            "width": rendered.width, "height": rendered.height, "skipped_images": rendered.skipped_images,
            "source_sha256": format!("{:x}", Sha256::digest(&bytes)),
            "png_sha256": format!("{:x}", Sha256::digest(&rendered.bytes))
        })
    );
    Ok(())
}
