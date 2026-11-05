use std::collections::{HashMap, HashSet};

use serde_json::{Map, Value, json};

pub fn region_layout_request(package: &Value, page_gap: f64) -> Value {
    let mut request = raw_request(package, page_gap);
    integral_numbers(&mut request);
    request
}

/// The request with every number as the parse model holds it.
fn raw_request(package: &Value, page_gap: f64) -> Value {
    let body = package.get("document").filter(|value| !value.is_null());
    let source_sections = body
        .and_then(|body| body.get("sections"))
        .and_then(Value::as_array);
    let final_properties = source_sections
        .and_then(|sections| sections.last())
        .and_then(|section| section.get("properties"))
        .filter(|value| !value.is_null())
        .or_else(|| body.and_then(|body| body.get("finalSectionProperties")));
    let mut sections = Vec::new();
    for section in source_sections.into_iter().flatten() {
        let mut entry = Map::new();
        if let Some(id) = section
            .get("id")
            .filter(|value| !value.is_null())
            .or_else(|| section["properties"].get("sectionId"))
        {
            entry.insert("sectionId".into(), id.clone());
        }
        copy_field(&mut entry, section, "properties");
        sections.push(Value::Object(entry));
    }
    let mut last = Map::new();
    if let Some(id) = sections
        .last()
        .and_then(|section| section.get("sectionId"))
        .filter(|value| !value.is_null())
        .or_else(|| body.and_then(|body| body["finalSectionProperties"].get("sectionId")))
    {
        last.insert("sectionId".into(), id.clone());
    }
    last.insert(
        "properties".into(),
        final_properties
            .filter(|value| !value.is_null())
            .cloned()
            .unwrap_or_else(|| json!({})),
    );
    sections.push(Value::Object(last));
    let mut regions = Map::new();
    regions.insert("sections".into(), Value::Array(sections));
    copy_field(&mut regions, package, "settings");
    if let Some(watermark) = final_properties.and_then(|properties| properties.get("watermark")) {
        regions.insert("watermark".into(), watermark.clone());
    }
    let mut options = Map::new();
    copy_field(&mut options, package, "contractVersion");
    options.insert("pageGap".into(), json!(page_gap));
    let mut contents = Vec::new();
    for (field, kind) in [("footnotes", "footnote"), ("endnotes", "endnote")] {
        for note in package[field].as_array().into_iter().flatten() {
            if let Some(note_type) = note["noteType"].as_str()
                && !note_type.is_empty()
                && note_type != "normal"
            {
                continue;
            }
            let mut entry = Map::new();
            copy_field(&mut entry, note, "id");
            entry.insert("noteKind".into(), json!(kind));
            entry.insert("height".into(), json!(0));
            contents.push(Value::Object(entry));
        }
    }
    json!({
        "bodyStory": "body",
        "options": options,
        "regions": regions,
        "notes": {"contents": contents},
        "renderEnv": render_environment(package),
    })
}

/// Writes integral numbers as integers. The parse model stores every number as a float, while
/// layout reads ids, counts and starts (note ids, `numStart`, footnote columns) as integers and
/// accepts integers wherever it reads a float.
fn integral_numbers(value: &mut Value) {
    match value {
        Value::Number(number) if !number.is_i64() && !number.is_u64() => {
            if let Some(float) = number.as_f64()
                && float.fract() == 0.0
                && float.abs() <= 9_007_199_254_740_992.0
            {
                *value = json!(float as i64);
            }
        }
        Value::Array(items) => items.iter_mut().for_each(integral_numbers),
        Value::Object(map) => map.values_mut().for_each(integral_numbers),
        _ => {}
    }
}

fn copy_field(target: &mut Map<String, Value>, source: &Value, name: &str) {
    if let Some(value) = source.get(name) {
        target.insert(name.into(), value.clone());
    }
}

fn render_environment(package: &Value) -> Value {
    let mut env = Map::new();
    if let Some(theme) = package.get("theme") {
        env.insert("fontTheme".into(), theme.clone());
    }
    copy_field(&mut env, &package["settings"], "compatibilityFlags");
    let theme_colors: Map<String, Value> = package["theme"]["colorScheme"]
        .as_object()
        .into_iter()
        .flatten()
        .filter(|(_, value)| value.is_string())
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect();
    env.insert("themeColors".into(), Value::Object(theme_colors));
    env.insert(
        "defaultTabStopTwips".into(),
        package["settings"]["defaultTabStop"].clone(),
    );
    env.insert("numericIds".into(), json!({}));
    env.insert("showHiddenText".into(), json!(false));
    let mut toc_ids = Vec::new();
    if let Some(definitions) = package.get("styles").filter(|value| !value.is_null()) {
        let styles = definitions["styles"]
            .as_array()
            .map_or(&[][..], Vec::as_slice);
        let default = styles.iter().find(|style| {
            style["type"] == "paragraph"
                && style["default"].as_bool() == Some(true)
                && style["styleId"].as_str().is_some_and(|id| !id.is_empty())
        });
        let default = default.or_else(|| {
            styles
                .iter()
                .find(|style| style["type"] == "paragraph" && style["styleId"] == "Normal")
        });
        if let Some(style) = default {
            env.insert("defaultParagraphStyleId".into(), style["styleId"].clone());
        }
        let language: Map<String, Value> = definitions["docDefaults"]["rPr"]["language"]
            .as_object()
            .into_iter()
            .flatten()
            .filter(|(key, value)| {
                matches!(key.as_str(), "latin" | "eastAsia" | "bidi")
                    && value.as_str().is_some_and(|tag| !tag.is_empty())
            })
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
        if !language.is_empty() {
            env.insert("docDefaultLanguage".into(), Value::Object(language));
        }
        let doc_default = definitions["docDefaults"]["rPr"]["fontSize"]
            .as_f64()
            .unwrap_or(20.0);
        let by_id: HashMap<_, _> = styles
            .iter()
            .filter_map(|style| style["styleId"].as_str().map(|id| (id, style)))
            .collect();
        let mut sizes = Map::new();
        let mut hyperlink_sizes = Map::new();
        for style in styles {
            let Some(id) = style["styleId"].as_str().filter(|id| !id.is_empty()) else {
                continue;
            };
            let paragraph = style["type"] == "paragraph";
            let hyperlink = style["type"] == "character"
                && [Some(id), style["name"].as_str()]
                    .into_iter()
                    .flatten()
                    .any(|name| {
                        name.eq_ignore_ascii_case("Hyperlink")
                            || name.eq_ignore_ascii_case("FollowedHyperlink")
                    });
            if paragraph
                && [Some(id), style["name"].as_str()]
                    .into_iter()
                    .flatten()
                    .any(is_toc_name)
                && !toc_ids.contains(&json!(id))
            {
                toc_ids.push(json!(id));
            }
            if !paragraph && !hyperlink {
                continue;
            }
            let target = if hyperlink {
                &mut hyperlink_sizes
            } else {
                &mut sizes
            };
            if target.contains_key(id) {
                continue;
            }
            let mut current = Some(style);
            let mut seen = HashSet::new();
            let mut size = None;
            while let Some(style) = current {
                if !seen.insert(style["styleId"].as_str()) {
                    break;
                }
                if let Some(value) = style["rPr"]["fontSize"].as_f64() {
                    size = Some(value);
                    break;
                }
                current = style["basedOn"]
                    .as_str()
                    .filter(|id| !id.is_empty())
                    .and_then(|id| by_id.get(id).copied());
            }
            if let Some(size) = size.or(if hyperlink { None } else { Some(doc_default) }) {
                target.insert(id.into(), json!(size / 2.0));
            }
        }
        if !sizes.is_empty() {
            env.insert("paragraphStyleFontSizes".into(), Value::Object(sizes));
        }
        if !hyperlink_sizes.is_empty() {
            env.insert(
                "hyperlinkStyleFontSizes".into(),
                Value::Object(hyperlink_sizes),
            );
        }
        env.insert("docDefaultFontSize".into(), json!(doc_default / 2.0));
    }
    env.insert("tocStyleIds".into(), Value::Array(toc_ids));
    Value::Object(env)
}

fn is_toc_name(name: &str) -> bool {
    let Some(prefix) = name.get(..3) else {
        return false;
    };
    if !prefix.eq_ignore_ascii_case("TOC") {
        return false;
    }
    let suffix = name[3..].trim_start_matches(|ch: char| {
        matches!(ch, '\u{0009}'..='\u{000d}' | ' ' | '\u{00a0}' | '\u{1680}'
            | '\u{2000}'..='\u{200a}' | '\u{2028}' | '\u{2029}' | '\u{202f}'
            | '\u{205f}' | '\u{3000}' | '\u{feff}')
    });
    !suffix.is_empty() && suffix.bytes().all(|byte| byte.is_ascii_digit())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_style_and_theme_environment() {
        let request = region_layout_request(
            &json!({
                "document": {},
                "theme": {"colorScheme": {"accent1": "FF0000", "ignored": 42}},
                "settings": {"defaultTabStop": 720, "compatibilityFlags": {"noLeading": true}},
                "styles": {
                    "docDefaults": {"rPr": {"fontSize": 24}},
                    "styles": [
                        {"styleId": "Normal", "type": "paragraph", "default": true},
                        {"styleId": "TOC1", "name": "TOC 1", "type": "paragraph", "basedOn": "Large"},
                        {"styleId": "Large", "type": "paragraph", "rPr": {"fontSize": 36}},
                        {"styleId": "Hyperlink", "type": "character", "basedOn": "Large"},
                        {"styleId": "Visited", "name": "followedhyperlink", "type": "character", "basedOn": "Hyperlink"},
                        {"styleId": "NoSizeLink", "name": "Hyperlink", "type": "character"},
                        {"styleId": "CycleA", "type": "paragraph", "basedOn": "CycleB"},
                        {"styleId": "CycleB", "type": "paragraph", "basedOn": "CycleA"}
                    ]
                }
            }),
            24.0,
        );
        let env = &request["renderEnv"];
        assert_eq!(env["defaultParagraphStyleId"], "Normal");
        assert_eq!(
            env["paragraphStyleFontSizes"],
            json!({
                "Normal": 12, "TOC1": 18, "Large": 18, "CycleA": 12, "CycleB": 12
            })
        );
        assert_eq!(
            env["hyperlinkStyleFontSizes"],
            json!({"Hyperlink": 18, "Visited": 18})
        );
        assert_eq!(env["docDefaultFontSize"], 12.0);
        assert_eq!(env["tocStyleIds"], json!(["TOC1"]));
        assert_eq!(env["themeColors"], json!({"accent1": "FF0000"}));
        assert_eq!(env["defaultTabStopTwips"], 720);
        assert_eq!(env["compatibilityFlags"], json!({"noLeading": true}));
        assert_eq!(env["fontTheme"]["colorScheme"]["accent1"], "FF0000");
        assert_eq!(env["numericIds"], json!({}));
        assert_eq!(env["showHiddenText"], false);
    }

    #[test]
    fn preserves_section_order_notes_watermark_and_contract() {
        let request = region_layout_request(
            &json!({
                "contractVersion": 7,
                "document": {
                    "sections": [
                        {"properties": {"sectionId": "first"}},
                        {"id": "second", "properties": {"watermark": {"text": "Draft"}}}
                    ],
                    "finalSectionProperties": {"sectionId": "authored", "watermark": {"text": "Old"}}
                },
                "footnotes": [{"id": -1, "noteType": "separator"}, {"id": 1}, {"id": 2, "noteType": "normal"}],
                "endnotes": [{"id": 3, "noteType": ""}, {"id": 0, "noteType": "continuationSeparator"}]
            }),
            18.0,
        );
        assert_eq!(
            request["options"],
            json!({"contractVersion": 7, "pageGap": 18})
        );
        assert_eq!(
            request["regions"]["sections"],
            json!([
                {"sectionId": "first", "properties": {"sectionId": "first"}},
                {"sectionId": "second", "properties": {"watermark": {"text": "Draft"}}},
                {"sectionId": "second", "properties": {"watermark": {"text": "Draft"}}}
            ])
        );
        assert_eq!(request["regions"]["watermark"], json!({"text": "Draft"}));
        assert_eq!(
            request["notes"]["contents"],
            json!([
                {"id": 1, "noteKind": "footnote", "height": 0},
                {"id": 2, "noteKind": "footnote", "height": 0},
                {"id": 3, "noteKind": "endnote", "height": 0}
            ])
        );
    }

    #[test]
    fn numbers_parsed_as_floats_reach_layout_as_integers() {
        let request = region_layout_request(
            &json!({
                "document": {"finalSectionProperties": {
                    "footnoteColumns": 2.0,
                    "pageWidth": 12240.5,
                    "footnotePr": {"numStart": 3.0}
                }},
                "footnotes": [{"id": -1.0, "noteType": "separator"}, {"id": 1.0, "noteType": "normal"}],
                "endnotes": [{"id": 2.0, "noteType": ""}]
            }),
            24.0,
        );
        assert_eq!(
            request["notes"]["contents"],
            json!([
                {"id": 1, "noteKind": "footnote", "height": 0},
                {"id": 2, "noteKind": "endnote", "height": 0}
            ])
        );
        let properties = &request["regions"]["sections"][0]["properties"];
        assert!(properties["footnoteColumns"].is_i64());
        assert!(properties["footnotePr"]["numStart"].is_i64());
        assert_eq!(properties["pageWidth"], json!(12240.5));
        assert!(request["options"]["pageGap"].is_i64());
    }

    /// Makes integral only the fields layout cannot read as a float, which the parse model hands
    /// over as floats.
    fn integral_where_required(value: &mut Value, key: &str) {
        match value {
            Value::Number(_) if matches!(key, "id" | "numStart" | "footnoteColumns") => {
                integral_numbers(value);
            }
            Value::Array(items) => items
                .iter_mut()
                .for_each(|item| integral_where_required(item, key)),
            Value::Object(map) => map
                .iter_mut()
                .for_each(|(key, item)| integral_where_required(item, key)),
            _ => {}
        }
    }

    fn laid_out(bytes: &[u8], request: &Value) -> Value {
        let engine = crate::EngineSession::new(1);
        crate::seed_from_docx(engine.doc(), bytes).unwrap();
        serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap()
    }

    #[test]
    fn integral_numbers_never_change_layout() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let fixtures = root.join("crates/docx-edit/tests/fixtures/layout-request");
        let cases: std::collections::BTreeMap<String, String> =
            serde_json::from_slice(&std::fs::read(fixtures.join("cases.json")).unwrap()).unwrap();
        for (name, source) in &cases {
            let bytes = std::fs::read(root.join(source)).unwrap();
            let parsed = docx_parse::parse_docx_s9_wire(
                &bytes,
                docx_parse::S9ParseOptions {
                    parse_headers_footers: true,
                    parse_notes: true,
                    detect_variables: false,
                    include_canonical: false,
                    ..Default::default()
                },
            )
            .unwrap();
            let package = serde_json::to_value(parsed.document.package).unwrap();
            let measurement = json!({"defaults": {"fontFamily": "Calibri", "fontSize": 11}});
            let mut integral = region_layout_request(&package, 24.0);
            integral["measurement"] = measurement.clone();
            let mut floats = raw_request(&package, 24.0);
            integral_where_required(&mut floats, "");
            floats["measurement"] = measurement;
            assert_ne!(
                floats, integral,
                "{name}: the request has no float to compare"
            );
            assert_eq!(
                laid_out(&bytes, &floats),
                laid_out(&bytes, &integral),
                "{name}"
            );
        }
    }

    #[test]
    fn omits_undefined_fields_and_uses_style_defaults() {
        let empty = region_layout_request(&json!({}), 24.0);
        assert_eq!(
            empty,
            json!({
                "bodyStory": "body",
                "options": {"pageGap": 24},
                "regions": {"sections": [{"properties": {}}]},
                "notes": {"contents": []},
                "renderEnv": {
                    "themeColors": {}, "defaultTabStopTwips": null, "numericIds": {},
                    "showHiddenText": false, "tocStyleIds": []
                }
            })
        );
        let styles = region_layout_request(
            &json!({"styles": {"styles": [
                {"type": "paragraph", "styleId": "Normal"},
                {"type": "paragraph", "styleId": "Custom", "default": true},
                {"type": "paragraph", "styleId": "Contents", "name": "toc\u{00a0}12"},
                {"type": "paragraph", "styleId": "TOC"},
                {"type": "character", "styleId": "TOC2"}
            ]}}),
            24.0,
        );
        assert_eq!(styles["renderEnv"]["docDefaultFontSize"], 10.0);
        assert_eq!(styles["renderEnv"]["defaultParagraphStyleId"], "Custom");
        assert_eq!(styles["renderEnv"]["tocStyleIds"], json!(["Contents"]));
        assert!(styles["renderEnv"].get("hyperlinkStyleFontSizes").is_none());
        let normal = region_layout_request(
            &json!({"styles": {"styles": [
                {"type": "paragraph", "styleId": "Normal"}
            ]}}),
            24.0,
        );
        assert_eq!(normal["renderEnv"]["defaultParagraphStyleId"], "Normal");
    }
}
