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
    if args.len() != 3 && args.len() != 4 {
        return Err(
            "usage: docx-native-bench input.docx page.png fonts/manifest.json [chains.json]".into(),
        );
    }
    let bytes = fs::read(&args[0])?;
    let parsed = docx_parse::parse_docx_s9_wire(&bytes, docx_parse::S9ParseOptions::default())?;
    let package = serde_json::to_value(&parsed.document.package)?;
    drop(parsed);
    let engine = EngineSession::new(1);
    seed_from_docx(engine.doc(), &bytes)?;
    let mut request = region_request(&package);
    whole_numbers_as_integers(&mut request);
    let requirements: Value =
        serde_json::from_str(&engine.layout_font_requirements_json(&request.to_string())?)?;
    let fonts = match args.get(3) {
        Some(chains) => Fonts::replay(Path::new(chains), &requirements)?,
        None => Fonts::load(Path::new(&args[2]), &requirements)?,
    };
    request["measurement"] = if fonts.measurement.is_object() {
        let mut measurement = fonts.measurement.clone();
        measurement["fontChains"] = json!(fonts.ids);
        measurement
    } else {
        json!({
            "fontChains": fonts.ids,
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "compat": {
                "noLeading": package["settings"]["compatibilityFlags"]["noLeading"].as_bool().unwrap_or(false),
                "doNotExpandShiftReturn": package["settings"]["compatibilityFlags"]["doNotExpandShiftReturn"].as_bool().unwrap_or(false)
            },
            "authoritativeShaping": true
        })
    };
    let layout: Value = serde_json::from_str(
        &engine.layout_document_with_regions_retained_json(&request.to_string())?,
    )?;
    let notes_converged = layout["notesConverged"].as_bool().unwrap_or(true);
    let mut extras = json!({"fontChains": fonts.ids});
    if let Some(bands) = engine.retained_headers_footers_json()? {
        let mut bands: Value = serde_json::from_str(&bands)?;
        whole_numbers_as_integers(&mut bands);
        extras["headersFooters"] = bands;
    }
    engine.build_display_list_frame(&extras.to_string(), 0)?;
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
            "schema_version": 1, "status": "ok", "pages": pages, "notes_converged": notes_converged, "page": 1, "dpi": 96,
            "width": width, "height": height, "skipped_images": rendered.skipped_images,
            "source_sha256": format!("{:x}", Sha256::digest(&bytes)),
            "png_sha256": format!("{:x}", Sha256::digest(&rendered.bytes)),
            "font_path": if args.len() == 4 { "product-registry" } else { "bench-manifest" },
            "font_faces": fonts.ids.iter().map(|(key, chain)| {
                (key.clone(), chain.iter().map(|id| fonts.faces[*id as usize].clone()).collect::<Vec<_>>())
            }).collect::<BTreeMap<_, _>>()
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
    let doc_default_size = package["styles"]["docDefaults"]["rPr"]["fontSize"]
        .as_f64()
        .unwrap_or(20.0);
    let mut chain_sizes = serde_json::Map::new();
    let mut hyperlink_sizes = serde_json::Map::new();
    for style in &styles {
        let hyperlink = style["type"] == "character"
            && ["styleId", "name"].iter().any(|key| {
                matches!(
                    style[key]
                        .as_str()
                        .unwrap_or("")
                        .to_ascii_lowercase()
                        .as_str(),
                    "hyperlink" | "followedhyperlink"
                )
            });
        if style["type"] != "paragraph" && !hyperlink {
            continue;
        }
        let Some(style_id) = style["styleId"].as_str() else {
            continue;
        };
        let mut size = None;
        let mut seen: Vec<&str> = Vec::new();
        let mut current = Some(*style);
        while let Some(current_style) = current {
            let Some(current_id) = current_style["styleId"].as_str() else {
                break;
            };
            if seen.contains(&current_id) {
                break;
            }
            seen.push(current_id);
            if let Some(found) = current_style["rPr"]["fontSize"].as_f64() {
                size = Some(found);
                break;
            }
            current = current_style["basedOn"]
                .as_str()
                .and_then(|id| styles.iter().find(|style| style["styleId"] == id).copied());
        }
        if hyperlink {
            if let Some(size) = size {
                hyperlink_sizes.insert(style_id.to_owned(), json!(size / 2.0));
            }
        } else {
            let resolved = size.unwrap_or(doc_default_size);
            chain_sizes.insert(style_id.to_owned(), json!(resolved / 2.0));
        }
    }
    let mut render_env = json!({"themeColors": package["theme"]["colorScheme"],
        "defaultTabStopTwips": package["settings"]["defaultTabStop"],
        "defaultParagraphStyleId": default_style, "tocStyleIds": toc});
    if package["styles"].is_object() {
        render_env["paragraphStyleFontSizes"] = json!(chain_sizes);
        if !hyperlink_sizes.is_empty() {
            render_env["hyperlinkStyleFontSizes"] = json!(hyperlink_sizes);
        }
        render_env["docDefaultFontSize"] = json!(doc_default_size / 2.0);
    }
    json!({
        "bodyStory": "body", "options": {"pageGap": 24},
        "regions": {"sections": sections, "settings": package["settings"], "watermark": final_properties["watermark"]},
        "notes": {"contents": notes},
        "renderEnv": render_env
    })
}

struct Fonts {
    store: FontStore,
    chains: FontChains,
    ids: BTreeMap<String, Vec<u32>>,
    faces: Vec<String>,
    measurement: Value,
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
            faces: Vec::new(),
            measurement: Value::Null,
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

    fn replay(path: &Path, requirements: &Value) -> Result<Self> {
        let recorded: Value = serde_json::from_slice(&fs::read(path)?)?;
        let root = path.parent().ok_or("font chains file has no parent")?;
        let mut registry = Self {
            store: FontStore::new(),
            chains: FontChains::new(),
            ids: BTreeMap::new(),
            faces: Vec::new(),
            measurement: Value::Null,
        };
        docx_layout::clear_measure_fonts();
        if !recorded["measurement"].is_object() {
            return Err("font chains file has no measurement".into());
        }
        registry.measurement = recorded["measurement"].clone();
        let faces = recorded["faces"]
            .as_array()
            .ok_or("font chains file has no faces")?;
        let mut loaded = BTreeMap::new();
        for (index, face) in faces.iter().enumerate() {
            let id = match face["substituteOf"].as_u64() {
                Some(base) if base < index as u64 => registry.substitute(
                    base as u32,
                    face["family"]
                        .as_str()
                        .ok_or("substitute face has no family")?,
                    face["bold"].as_bool().unwrap_or(false),
                    face["italic"].as_bool().unwrap_or(false),
                )?,
                Some(_) => return Err("substitute face precedes its base".into()),
                None => registry.register(root, face, &mut loaded)?,
            };
            if id as usize != index {
                return Err(format!("font {index} registered as {id}").into());
            }
        }
        let recorded_chains = recorded["fontChains"]
            .as_object()
            .ok_or("font chains file has no fontChains")?;
        for (key, chain) in recorded_chains {
            let chain = chain
                .as_array()
                .ok_or("font chain is not a list")?
                .iter()
                .map(|id| {
                    id.as_u64()
                        .filter(|id| (*id as usize) < faces.len())
                        .map(|id| id as u32)
                        .ok_or("font chain names an unregistered font")
                })
                .collect::<std::result::Result<Vec<u32>, _>>()?;
            if chain.is_empty() {
                return Err(format!("font chain for {key} is empty").into());
            }
            registry.chains.insert(
                key.clone(),
                chain.iter().copied().map(FontId::from_u32).collect(),
            );
            registry.ids.insert(key.clone(), chain);
        }
        for requirement in requirements
            .as_array()
            .ok_or("invalid font requirements")?
        {
            let key = requirement["key"]
                .as_str()
                .ok_or("font requirement has no key")?;
            if !registry.ids.contains_key(key) {
                return Err(format!("no recorded font chain for {key}").into());
            }
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
        if engine as usize != self.faces.len() {
            return Err("font registry IDs out of order".into());
        }
        self.faces.push(format!("{:x}", Sha256::digest(&bytes)));
        loaded.insert(file.to_owned(), engine);
        Ok(engine)
    }

    fn substitute(&mut self, base: u32, family: &str, bold: bool, italic: bool) -> Result<u32> {
        let id = self.register_substitute(base, family, bold, italic)?;
        if id as usize == self.faces.len() {
            let face = format!("{}~{family}", self.faces[base as usize]);
            self.faces.push(face);
        }
        Ok(id)
    }

    fn register_substitute(
        &mut self,
        base: u32,
        family: &str,
        bold: bool,
        italic: bool,
    ) -> Result<u32> {
        #[cfg(feature = "substitute-widths")]
        {
            let raster = self
                .store
                .register_named_substitute(FontId::from_u32(base), family)?
                .to_u32();
            let engine = docx_layout::register_substitute_measure_font(base, family)
                .map_err(|_| "substitute font registration failed")?;
            if raster != engine {
                return Err("substitute font IDs diverged".into());
            }
            let _ = (bold, italic);
            Ok(engine)
        }
        #[cfg(all(feature = "substitute-styles", not(feature = "substitute-widths")))]
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
        #[cfg(all(
            feature = "substitute-metrics",
            not(any(feature = "substitute-styles", feature = "substitute-widths"))
        ))]
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

// The parser carries note ids, note numbering starts and chapter styles as floats while the
// layout request types them as integers; a JavaScript host writes them as integers already.
fn whole_numbers_as_integers(value: &mut Value) {
    match value {
        Value::Number(number) if number.is_f64() => {
            let float = number.as_f64().unwrap_or(0.5);
            if float.fract() == 0.0 && float.abs() < 9.0e15 {
                *value = json!(float as i64);
            }
        }
        Value::Array(items) => items.iter_mut().for_each(whole_numbers_as_integers),
        Value::Object(map) => map.values_mut().for_each(whole_numbers_as_integers),
        _ => {}
    }
}
