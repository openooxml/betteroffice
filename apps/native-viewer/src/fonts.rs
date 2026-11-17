#[cfg(feature = "docx")]
mod docx {
    use std::collections::{BTreeMap, HashMap};
    use std::fs;
    use std::path::Path;
    use std::sync::Arc;

    use anyhow::{Context, Result, anyhow};
    use docx_raster::FontChains;
    use ooxml_text::{FontId, FontStore, ShapeDirection, ShapeFeature, shape_with_direction};
    use serde_json::Value;
    use vello::Glyph;
    use vello::peniko::{Blob, FontData};

    type ShapedRuns = Vec<(u32, Vec<Glyph>)>;

    const FACES: &[(&str, &str, bool, bool)] = &[
        ("Caladea-Regular.ttf", "Caladea", false, false),
        ("Caladea-Bold.ttf", "Caladea", true, false),
        ("Caladea-Italic.ttf", "Caladea", false, true),
        ("Caladea-BoldItalic.ttf", "Caladea", true, true),
        ("Carlito-Regular.ttf", "Carlito", false, false),
        ("Carlito-Bold.ttf", "Carlito", true, false),
        ("Carlito-Italic.ttf", "Carlito", false, true),
        ("Carlito-BoldItalic.ttf", "Carlito", true, true),
        (
            "LiberationSans-Regular.ttf",
            "Liberation Sans",
            false,
            false,
        ),
        ("LiberationSans-Bold.ttf", "Liberation Sans", true, false),
        ("LiberationSans-Italic.ttf", "Liberation Sans", false, true),
        (
            "LiberationSans-BoldItalic.ttf",
            "Liberation Sans",
            true,
            true,
        ),
        (
            "LiberationSerif-Regular.ttf",
            "Liberation Serif",
            false,
            false,
        ),
        ("LiberationSerif-Bold.ttf", "Liberation Serif", true, false),
        (
            "LiberationSerif-Italic.ttf",
            "Liberation Serif",
            false,
            true,
        ),
        (
            "LiberationSerif-BoldItalic.ttf",
            "Liberation Serif",
            true,
            true,
        ),
        (
            "LiberationMono-Regular.ttf",
            "Liberation Mono",
            false,
            false,
        ),
        ("LiberationMono-Bold.ttf", "Liberation Mono", true, false),
        ("LiberationMono-Italic.ttf", "Liberation Mono", false, true),
        (
            "LiberationMono-BoldItalic.ttf",
            "Liberation Mono",
            true,
            true,
        ),
        (
            "NotoNaskhArabic-Regular.ttf",
            "Noto Naskh Arabic",
            false,
            false,
        ),
        (
            "NotoSansArabic-Regular.ttf",
            "Noto Sans Arabic",
            false,
            false,
        ),
        ("NotoSansArabic-Bold.ttf", "Noto Sans Arabic", true, false),
        (
            "NotoSansHebrew-Regular.ttf",
            "Noto Sans Hebrew",
            false,
            false,
        ),
        ("NotoSansHebrew-Bold.ttf", "Noto Sans Hebrew", true, false),
    ];

    pub struct FontFace {
        pub id: u32,
        pub data: FontData,
        family: &'static str,
        bold: bool,
        italic: bool,
    }

    pub struct FontRegistry {
        pub faces: Vec<FontFace>,
        pub store: FontStore,
        pub chains: FontChains,
        pub chain_ids: BTreeMap<String, Vec<u32>>,
        pub requirements: Vec<String>,
    }

    impl FontRegistry {
        pub fn load(requirements: &Value) -> Result<Self> {
            docx_layout::clear_measure_fonts();
            let assets = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../packages/fonts/assets");
            let mut store = FontStore::new();
            let mut faces = Vec::with_capacity(FACES.len());
            for &(file, family, bold, italic) in FACES {
                let path = assets.join(file);
                let bytes = match fs::read(&path) {
                    Ok(bytes) => bytes,
                    Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue,
                    Err(error) => {
                        return Err(error).with_context(|| format!("read font {}", path.display()));
                    }
                };
                let engine_id = docx_layout::register_measure_font(&bytes)
                    .map_err(|_| anyhow!("register measurement font {}", path.display()))?;
                let raster_id = store
                    .register(bytes.clone())
                    .with_context(|| format!("register raster font {}", path.display()))?;
                if engine_id != raster_id.to_u32() {
                    return Err(anyhow!("font registry ids diverged at {file}"));
                }
                faces.push(FontFace {
                    id: engine_id,
                    data: FontData::new(Blob::from(bytes), 0),
                    family,
                    bold,
                    italic,
                });
            }

            for (bold, italic, bytes) in base_latin::faces() {
                faces.push(FontFace {
                    id: FontId::base_latin(bold, italic).to_u32(),
                    data: FontData::new(Blob::new(Arc::new(bytes)), 0),
                    family: base_latin::FAMILY,
                    bold,
                    italic,
                });
            }

            let mut chain_ids = BTreeMap::new();
            let mut labels = Vec::new();
            for requirement in requirements
                .as_array()
                .context("font requirements are not an array")?
            {
                let key = requirement["key"]
                    .as_str()
                    .context("font requirement has no key")?;
                let family = requirement["family"]
                    .as_str()
                    .context("font requirement has no family")?;
                let bold = requirement["bold"].as_bool().unwrap_or(false);
                let italic = requirement["italic"].as_bool().unwrap_or(false);
                let scripts = requirement["scripts"]
                    .as_array()
                    .cloned()
                    .unwrap_or_default();
                let mut chain: Vec<_> = pick_face(&faces, family, bold, italic)
                    .into_iter()
                    .collect();
                for script in scripts.iter().filter_map(Value::as_str) {
                    match script {
                        "arabic" => {
                            push_unique(
                                &mut chain,
                                pick_exact(&faces, "Noto Sans Arabic", bold, false),
                            );
                            push_unique(
                                &mut chain,
                                pick_exact(&faces, "Noto Naskh Arabic", false, false),
                            );
                        }
                        "hebrew" => {
                            push_unique(
                                &mut chain,
                                pick_exact(&faces, "Noto Sans Hebrew", bold, false),
                            );
                        }
                        _ => {}
                    }
                }
                push_unique(
                    &mut chain,
                    pick_exact(&faces, "Liberation Sans", bold, italic),
                );
                if chain.is_empty() {
                    chain.push(FontId::base_latin(bold, italic).to_u32());
                }
                labels.push(format!(
                    "{family}{}{}",
                    if bold { " bold" } else { "" },
                    if italic { " italic" } else { "" }
                ));
                chain_ids.insert(key.to_owned(), chain);
            }
            if chain_ids.is_empty() {
                let mut chain: Vec<_> = pick_exact(&faces, "Carlito", false, false)
                    .into_iter()
                    .collect();
                if chain.is_empty() {
                    chain.push(FontId::base_latin(false, false).to_u32());
                }
                chain_ids.insert("calibri|0|0".to_owned(), chain);
            }
            let chains = chain_ids
                .iter()
                .map(|(key, ids)| {
                    (
                        key.clone(),
                        ids.iter().copied().map(FontId::from_u32).collect(),
                    )
                })
                .collect::<HashMap<_, _>>();
            Ok(Self {
                faces,
                store,
                chains,
                chain_ids,
                requirements: labels,
            })
        }

        pub fn face(&self, id: u32) -> Option<&FontFace> {
            self.faces
                .get(id as usize)
                .filter(|face| face.id == id)
                .or_else(|| self.faces.iter().find(|face| face.id == id))
        }

        pub fn shape_fallback(
            &self,
            text: &str,
            font: &str,
            rtl: bool,
            small_caps: bool,
            letter_spacing: f32,
            word_spacing: f32,
        ) -> Result<(f32, ShapedRuns)> {
            let spec = parse_css_font(font)?;
            let key = format!(
                "{}|{}|{}",
                spec.family.to_lowercase(),
                u8::from(spec.bold),
                u8::from(spec.italic)
            );
            let chain = self.chains.get(&key).cloned().unwrap_or_else(|| {
                if spec.family == base_latin::FAMILY {
                    vec![FontId::base_latin(spec.bold, spec.italic)]
                } else {
                    vec![
                        pick_face(&self.faces, &spec.family, spec.bold, spec.italic)
                            .map(FontId::from_u32)
                            .unwrap_or(FontId::base_latin(spec.bold, spec.italic)),
                    ]
                }
            });
            let mut segments: Vec<(FontId, usize, usize)> = Vec::new();
            for (start, ch) in text.char_indices() {
                let font = self
                    .store
                    .resolve_with_family(&chain, ch, spec.bold, spec.italic, &spec.family)
                    .ok_or_else(|| anyhow!("text fallback has no face"))?;
                let end = start + ch.len_utf8();
                if let Some(last) = segments.last_mut()
                    && last.0 == font
                {
                    last.2 = end;
                } else {
                    segments.push((font, start, end));
                }
            }
            if rtl {
                segments.reverse();
            }
            let direction = if rtl {
                ShapeDirection::Rtl
            } else {
                ShapeDirection::Ltr
            };
            let features = if spec.small_caps || small_caps {
                vec![ShapeFeature {
                    tag: *b"smcp",
                    value: 1,
                }]
            } else {
                Vec::new()
            };
            let mut pen = 0.0f32;
            let mut runs = Vec::new();
            for (segment_index, &(font_id, start, end)) in segments.iter().enumerate() {
                let text = &text[start..end];
                let shaped = shape_with_direction(
                    &self.store,
                    font_id,
                    text,
                    spec.size,
                    &features,
                    direction,
                )?;
                let glyph_count = shaped.len();
                let glyphs = shaped
                    .iter()
                    .enumerate()
                    .map(|(index, glyph)| {
                        let placed = Glyph {
                            id: glyph.glyph_id,
                            x: pen + glyph.x_offset,
                            y: -glyph.y_offset,
                        };
                        pen += glyph.x_advance;
                        let cluster_ends = shaped
                            .get(index + 1)
                            .is_none_or(|next| next.cluster != glyph.cluster);
                        if cluster_ends {
                            let source = text
                                .get(glyph.cluster as usize..)
                                .and_then(|tail| tail.chars().next());
                            if source == Some(' ') {
                                pen += word_spacing;
                            }
                            if index + 1 < glyph_count || segment_index + 1 < segments.len() {
                                pen += letter_spacing;
                            }
                        }
                        placed
                    })
                    .collect();
                runs.push((font_id.to_u32(), glyphs));
            }
            Ok((spec.size, runs))
        }
    }

    fn pick_face(faces: &[FontFace], family: &str, bold: bool, italic: bool) -> Option<u32> {
        let family = family.to_lowercase();
        let metric_family = if family.contains("calibri") || family.contains("carlito") {
            "Carlito"
        } else if family.contains("cambria") || family.contains("caladea") {
            "Caladea"
        } else if family.contains("courier") || family.contains("mono") {
            "Liberation Mono"
        } else if family.contains("times") || family.contains("serif") {
            "Liberation Serif"
        } else {
            "Liberation Sans"
        };
        pick_exact(faces, metric_family, bold, italic)
    }

    fn pick_exact(faces: &[FontFace], family: &str, bold: bool, italic: bool) -> Option<u32> {
        faces
            .iter()
            .find(|face| face.family == family && face.bold == bold && face.italic == italic)
            .or_else(|| {
                faces
                    .iter()
                    .find(|face| face.family == family && face.bold == bold)
            })
            .or_else(|| faces.iter().find(|face| face.family == family))
            .map(|face| face.id)
    }

    fn push_unique(chain: &mut Vec<u32>, id: Option<u32>) {
        if let Some(id) = id
            && !chain.contains(&id)
        {
            chain.push(id);
        }
    }

    struct CssFont {
        family: String,
        size: f32,
        bold: bool,
        italic: bool,
        small_caps: bool,
    }

    fn parse_css_font(font: &str) -> Result<CssFont> {
        let mut size_match = None;
        let mut cursor = 0usize;
        for token in font.split_whitespace() {
            let relative = font[cursor..]
                .find(token)
                .with_context(|| format!("unsupported font shorthand {font}"))?;
            let index = cursor + relative;
            cursor = index + token.len();
            if token.ends_with("px") {
                size_match = Some((index, token));
                break;
            }
        }
        let (size_index, size_token) =
            size_match.with_context(|| format!("font has no size: {font}"))?;
        let size = size_token
            .strip_suffix("px")
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite() && *value > 0.0)
            .with_context(|| format!("invalid font size {size_token}"))?;
        let mut bold = false;
        let mut italic = false;
        let mut small_caps = false;
        for token in font[..size_index].split_whitespace() {
            match token {
                "normal" => {}
                "italic" | "oblique" => italic = true,
                "small-caps" => small_caps = true,
                "bold" | "bolder" => bold = true,
                value => {
                    let weight = value
                        .parse::<u16>()
                        .with_context(|| format!("unsupported font token {value}"))?;
                    if !(1..=1000).contains(&weight) {
                        return Err(anyhow!("invalid font weight {value}"));
                    }
                    bold = weight >= 600;
                }
            }
        }
        let family = first_family(font[size_index + size_token.len()..].trim())
            .trim_matches(['\'', '"'])
            .trim()
            .to_owned();
        if family.is_empty() {
            return Err(anyhow!("font has no family: {font}"));
        }
        Ok(CssFont {
            family,
            size,
            bold,
            italic,
            small_caps,
        })
    }

    fn first_family(value: &str) -> &str {
        let mut quote = None;
        for (index, character) in value.char_indices() {
            match character {
                '\'' | '"' if quote == Some(character) => quote = None,
                '\'' | '"' if quote.is_none() => quote = Some(character),
                ',' if quote.is_none() => return &value[..index],
                _ => {}
            }
        }
        value
    }

    #[cfg(test)]
    mod tests {
        use super::*;

        #[test]
        fn parses_layout_font_shorthand() {
            let font = parse_css_font("italic small-caps 700 12.5px \"Aptos Display\", sans-serif")
                .unwrap();
            assert_eq!(font.family, "Aptos Display");
            assert_eq!(font.size, 12.5);
            assert!(font.bold);
            assert!(font.italic);
            assert!(font.small_caps);
        }

        #[test]
        fn rejects_unknown_font_tokens() {
            assert!(parse_css_font("condensed 12px Carlito").is_err());
        }
    }
}

#[cfg(feature = "docx")]
pub use docx::FontRegistry;
