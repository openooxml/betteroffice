use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use docx_edit::{EngineSession, seed_from_docx};
use docx_raster::{FontChains, ImageMap, RenderResources};
use ooxml_text::{FontId, FontStore};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

fn main() {
    if let Err(error) = run() {
        eprintln!(
            "{}",
            json!({"status": "failed", "error": error.to_string()})
        );
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let args: Vec<_> = std::env::args_os().skip(1).collect();
    if args.len() != 3 {
        return Err("usage: docx-native-bench input.docx page.png fonts/manifest.json".into());
    }
    let bytes = fs::read(&args[0])?;
    let parsed = docx_parse::parse_docx_s9_wire(&bytes, docx_parse::S9ParseOptions::default())?;
    let package = serde_json::to_value(&parsed.document.package)?;
    drop(parsed);
    let engine = EngineSession::new(1);
    seed_from_docx(engine.doc(), &bytes)?;
    let mut request = region_request(&package);
    let requirements: Value =
        serde_json::from_str(&engine.layout_font_requirements_json(&request.to_string())?)?;
    let fonts = Fonts::load(Path::new(&args[2]), &requirements)?;
    request["measurement"] = json!({
        "fontChains": fonts.ids,
        "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        "compat": {
            "noLeading": package["settings"]["compatibilityFlags"]["noLeading"].as_bool().unwrap_or(false),
            "doNotExpandShiftReturn": package["settings"]["compatibilityFlags"]["doNotExpandShiftReturn"].as_bool().unwrap_or(false)
        },
        "authoritativeShaping": true
    });
    let layout: Value = serde_json::from_str(
        &engine.layout_document_with_regions_retained_json(&request.to_string())?,
    )?;
    if layout["notesConverged"].as_bool() == Some(false) {
        return Err("note layout did not converge".into());
    }
    engine.build_display_list_frame(&json!({"fontChains": fonts.ids}).to_string(), 0)?;
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts.store, &fonts.chains, &images);
    let (rendered, pages, width, height) = engine
        .with_display_list(|list| {
            let page = list.pages.first().ok_or("document has no pages")?;
            let rendered = docx_raster::render_page(list, 0, &resources)?;
            Ok::<_, Box<dyn std::error::Error>>((
                rendered,
                list.pages.len(),
                page.width.clone(),
                page.height.clone(),
            ))
        })
        .ok_or("engine did not retain a display list")??;
    if rendered.skipped_images != 0 {
        return Err(format!("renderer skipped {} images", rendered.skipped_images).into());
    }
    fs::write(&args[1], &rendered.bytes)?;
    println!(
        "{}",
        json!({
            "schema_version": 1, "status": "ok", "pages": pages, "page": 1, "dpi": 96,
            "width": width, "height": height, "skipped_images": rendered.skipped_images,
            "source_sha256": format!("{:x}", Sha256::digest(&bytes)),
            "png_sha256": format!("{:x}", Sha256::digest(&rendered.bytes))
        })
    );
    Ok(())
}

fn region_request(package: &Value) -> Value {
    let body = &package["document"];
    let mut sections: Vec<Value> = body["sections"]
        .as_array()
        .into_iter()
        .flatten()
        .map(|section| json!({"sectionId": section["id"], "properties": section["properties"]}))
        .collect();
    let final_properties = body["sections"]
        .as_array()
        .and_then(|sections| sections.last())
        .map(|section| section["properties"].clone())
        .or_else(|| body.get("finalSectionProperties").cloned())
        .unwrap_or_else(|| json!({}));
    sections.push(json!({"properties": final_properties}));
    let mut notes = Vec::new();
    for (field, kind) in [("footnotes", "footnote"), ("endnotes", "endnote")] {
        for note in package[field].as_array().into_iter().flatten() {
            if note["noteType"]
                .as_str()
                .is_none_or(|kind| kind == "normal")
            {
                notes.push(json!({"id": note["id"], "noteKind": kind, "height": 0}));
            }
        }
    }
    let styles: Vec<_> = package["styles"]["styles"]
        .as_array()
        .into_iter()
        .flatten()
        .collect();
    let default_style = styles
        .iter()
        .find(|style| style["type"] == "paragraph" && style["default"] == true)
        .or_else(|| styles.iter().find(|style| style["styleId"] == "Normal"))
        .map(|style| style["styleId"].clone());
    let toc: Vec<_> = styles
        .iter()
        .filter(|style| style["type"] == "paragraph")
        .filter(|style| {
            ["styleId", "name"].iter().any(|key| {
                let value = style[key].as_str().unwrap_or("").to_uppercase();
                value.strip_prefix("TOC").is_some_and(|tail| {
                    !tail.trim().is_empty() && tail.trim().chars().all(|c| c.is_ascii_digit())
                })
            })
        })
        .map(|style| style["styleId"].clone())
        .collect();
    json!({
        "bodyStory": "body", "options": {"pageGap": 24},
        "regions": {"sections": sections, "settings": package["settings"], "watermark": final_properties["watermark"]},
        "notes": {"contents": notes},
        "renderEnv": {"themeColors": package["theme"]["colorScheme"],
            "defaultTabStopTwips": package["settings"]["defaultTabStop"],
            "defaultParagraphStyleId": default_style, "tocStyleIds": toc}
    })
}

struct Fonts {
    store: FontStore,
    chains: FontChains,
    ids: BTreeMap<String, Vec<u32>>,
}

impl Fonts {
    fn load(path: &Path, requirements: &Value) -> Result<Self> {
        let manifest: Value = serde_json::from_slice(&fs::read(path)?)?;
        let root = path.parent().ok_or("font manifest has no parent")?;
        let faces = manifest["faces"]
            .as_array()
            .ok_or("font manifest has no faces")?;
        let mut registry = Self {
            store: FontStore::new(),
            chains: FontChains::new(),
            ids: BTreeMap::new(),
        };
        docx_layout::clear_measure_fonts();
        let mut loaded = BTreeMap::<String, u32>::new();
        let mut requirements = requirements
            .as_array()
            .ok_or("invalid font requirements")?
            .clone();
        requirements.push(
            json!({"key": "calibri|0|0", "family": "Calibri", "bold": false, "italic": false}),
        );
        for requirement in requirements {
            let family = requirement["family"]
                .as_str()
                .ok_or("font requirement has no family")?;
            let bold = requirement["bold"].as_bool().unwrap_or(false);
            let italic = requirement["italic"].as_bool().unwrap_or(false);
            let normalized = family.trim().to_lowercase();
            let aliased = manifest["aliases"][&normalized]
                .as_str()
                .unwrap_or(&normalized);
            let covered = faces.iter().find(|face| {
                face["family"].as_str().unwrap_or("").to_lowercase() == aliased
                    || face["metricCompatWith"]
                        .as_str()
                        .unwrap_or("")
                        .to_lowercase()
                        == aliased
            });
            let fallback = if normalized == "calibri light" {
                "Carlito"
            } else if [
                "times",
                "georgia",
                "garamond",
                "palatino",
                "baskerville",
                "bodoni",
                "cambria",
                "minion",
                "mincho",
                "明朝",
                "明體",
                "宋",
                "ming",
                "song",
                "serif",
            ]
            .iter()
            .any(|word| normalized.contains(word))
            {
                "Liberation Serif"
            } else {
                "Liberation Sans"
            };
            let target = covered
                .and_then(|face| face["family"].as_str())
                .unwrap_or(fallback);
            let primary = pick_face(faces, target, bold, italic)?;
            let base = registry.register(root, primary, &mut loaded)?;
            let first = registry.substitute(base, family, bold, italic)?;
            let mut chain = vec![first];
            for script in requirement["scripts"]
                .as_array()
                .into_iter()
                .flatten()
                .filter_map(Value::as_str)
            {
                if let Some(face) = faces
                    .iter()
                    .filter(|face| face["script"] == script)
                    .min_by_key(|face| {
                        usize::from(face["weight"] != if bold { 700 } else { 400 }) * 2
                            + usize::from(face["style"] != if italic { "italic" } else { "normal" })
                    })
                {
                    let id = registry.register(root, face, &mut loaded)?;
                    if !chain.contains(&id) {
                        chain.push(id);
                    }
                }
            }
            let fallback = registry.register(
                root,
                pick_face(faces, "Liberation Sans", bold, italic)?,
                &mut loaded,
            )?;
            if !chain.contains(&fallback) {
                chain.push(fallback);
            }
            let key = requirement["key"]
                .as_str()
                .ok_or("font requirement has no key")?
                .to_owned();
            registry.chains.insert(
                key.clone(),
                chain.iter().copied().map(FontId::from_u32).collect(),
            );
            registry.ids.insert(key, chain);
        }
        Ok(registry)
    }

    fn register(
        &mut self,
        root: &Path,
        face: &Value,
        loaded: &mut BTreeMap<String, u32>,
    ) -> Result<u32> {
        let file = face["file"].as_str().ok_or("font has no file")?;
        if let Some(id) = loaded.get(file) {
            return Ok(*id);
        }
        if PathBuf::from(file).components().count() != 1 {
            return Err("invalid font filename".into());
        }
        let bytes = fs::read(root.join(file))?;
        if format!("{:x}", Sha256::digest(&bytes)) != face["sha256"].as_str().unwrap_or("") {
            return Err(format!("font hash mismatch: {file}").into());
        }
        let raster = self.store.register(bytes.clone())?.to_u32();
        let engine = docx_layout::register_measure_font(&bytes)
            .map_err(|_| "measurement font registration failed")?;
        if raster != engine {
            return Err("font registry IDs diverged".into());
        }
        loaded.insert(file.to_owned(), engine);
        Ok(engine)
    }

    fn substitute(&mut self, base: u32, family: &str, bold: bool, italic: bool) -> Result<u32> {
        #[cfg(feature = "substitute-styles")]
        {
            let Some(metrics) = ooxml_text::word_fonts::requested_line_metrics(family) else {
                return Ok(base);
            };
            let raster = self
                .store
                .register_substitute(FontId::from_u32(base), metrics, bold, italic)?
                .to_u32();
            let engine = docx_layout::register_substitute_measure_font(base, family, bold, italic)
                .map_err(|_| "substitute font registration failed")?;
            if raster != engine {
                return Err("substitute font IDs diverged".into());
            }
            Ok(engine)
        }
        #[cfg(all(feature = "substitute-metrics", not(feature = "substitute-styles")))]
        {
            let Some(metrics) = ooxml_text::word_fonts::requested_line_metrics(family) else {
                return Ok(base);
            };
            let raster = self
                .store
                .register_substitute(FontId::from_u32(base), metrics)?
                .to_u32();
            let engine = docx_layout::register_substitute_measure_font(base, family)
                .map_err(|_| "substitute font registration failed")?;
            if raster != engine {
                return Err("substitute font IDs diverged".into());
            }
            let _ = (bold, italic);
            Ok(engine)
        }
        #[cfg(not(feature = "substitute-metrics"))]
        {
            let _ = (family, bold, italic);
            Ok(base)
        }
    }
}

fn pick_face<'a>(faces: &'a [Value], family: &str, bold: bool, italic: bool) -> Result<&'a Value> {
    faces
        .iter()
        .find(|face| {
            face["family"] == family
                && face["weight"] == if bold { 700 } else { 400 }
                && face["style"] == if italic { "italic" } else { "normal" }
        })
        .or_else(|| {
            faces.iter().find(|face| {
                face["family"] == family && face["weight"] == 400 && face["style"] == "normal"
            })
        })
        .ok_or_else(|| format!("no bundled font for {family}").into())
}
