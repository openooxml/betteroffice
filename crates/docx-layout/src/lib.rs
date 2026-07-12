//! DOCX pagination and display-list generation.

pub mod canonical;
pub mod hooks;
pub mod page_flow;
pub mod paragraph_spacing;
pub mod place;
pub mod prescan;
pub mod resolve_lines;
pub mod types;

pub mod break_policy;
pub mod cell_layout;
pub mod column_balancing;
pub mod display_list;
pub mod floating_objects;
pub mod footnotes;
pub mod hf_bands;
pub mod hit;
pub mod keep_together;
pub mod section_breaks;
pub mod session;
pub mod table_grid;
pub mod table_row_break;

/// Why the engine refused an input.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LayoutError {
    Unsupported(String),
    Invalid(String),
}

impl std::fmt::Display for LayoutError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            LayoutError::Unsupported(reason) => write!(f, "unsupported: {reason}"),
            LayoutError::Invalid(reason) => write!(f, "invalid: {reason}"),
        }
    }
}

/// Parse the `{ measured, options }` envelope and run the placement walk.
pub fn compute_layout(input: &str) -> Result<types::Layout, LayoutError> {
    let parsed: types::Input =
        serde_json::from_str(input).map_err(|e| LayoutError::Invalid(format!("parse: {e}")))?;
    place::layout_document(parsed)
}

/// Serialize a layout request to JSON.
pub fn layout_to_json(input: &str) -> Result<String, String> {
    let layout = compute_layout(input).map_err(|e| match e {
        LayoutError::Unsupported(_) => "UNSUPPORTED".to_string(),
        LayoutError::Invalid(reason) => {
            if reason.starts_with("parse: ") {
                reason
            } else {
                "UNSUPPORTED".to_string()
            }
        }
    })?;
    serde_json::to_string(&layout).map_err(|e| format!("serialize: {e}"))
}

/// Serialize a layout request in canonical form.
pub fn layout_to_canonical_json(input: &str) -> Result<String, LayoutError> {
    let layout = compute_layout(input)?;
    Ok(canonical::serialize_layout(&layout))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options_json() -> serde_json::Value {
        serde_json::json!({
            "pageSize": {"w": 816.0, "h": 1056.0},
            "margins": {"top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0},
            "pageGap": 20.0
        })
    }

    fn para(id: u32, text: &str, pm_start: u32, height: f64) -> serde_json::Value {
        let pm_end = pm_start + text.len() as u32;
        serde_json::json!({
            "block": {
                "kind": "paragraph",
                "id": id,
                "runs": [{"kind": "text", "text": text, "pmStart": pm_start, "pmEnd": pm_end}],
                "pmStart": pm_start,
                "pmEnd": pm_end + 1
            },
            "measure": {
                "kind": "paragraph",
                "totalHeight": height,
                "lines": [{
                    "headRun": 0, "headChar": 0, "tailRun": 0,
                    "tailChar": text.len(),
                    "width": 120.0, "ascent": height * 0.8, "descent": height * 0.2,
                    "lineHeight": height
                }]
            }
        })
    }

    #[test]
    fn stacks_paragraphs_on_one_page() {
        let input = serde_json::json!({
            "measured": [para(0, "First paragraph", 1, 24.0), para(1, "Second paragraph", 18, 24.0)],
            "options": options_json()
        })
        .to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 1);
        let frags = v["pages"][0]["fragments"].as_array().unwrap();
        assert_eq!(frags.len(), 2);
        assert_eq!(frags[0]["y"], 96.0);
        assert_eq!(frags[1]["y"], 120.0);
        assert_eq!(frags[0]["width"], 624.0);
    }

    #[test]
    fn overflows_to_a_second_page() {
        let mut measured = Vec::new();
        for i in 0..10u32 {
            measured.push(para(i, "Paragraph", i * 15 + 1, 100.0));
        }
        let input =
            serde_json::json!({ "measured": measured, "options": options_json() }).to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 2);
    }

    #[test]
    fn unsupported_kind_errors_for_fallback() {
        let input = serde_json::json!({
            "measured": [{
                "block": {"kind": "somethingNew", "id": 0},
                "measure": {"kind": "somethingNew"}
            }],
            "options": options_json()
        })
        .to_string();
        assert_eq!(layout_to_json(&input).unwrap_err(), "UNSUPPORTED");
    }

    #[test]
    fn places_a_simple_table() {
        let input = serde_json::json!({
            "measured": [{
                "block": {
                    "kind": "table", "id": 0, "columnWidths": [100.0],
                    "rows": [{"id": 1, "cells": [{"id": 2, "blocks": [
                        {"kind": "paragraph", "id": 3, "runs": []}
                    ]}]}]
                },
                "measure": {
                    "kind": "table", "columnWidths": [100.0],
                    "totalWidth": 100.0, "totalHeight": 24.0,
                    "rows": [{"height": 24.0, "cells": [{
                        "blocks": [{"kind": "paragraph", "lines": [], "totalHeight": 24.0}],
                        "width": 100.0, "height": 24.0
                    }]}]
                }
            }],
            "options": options_json()
        })
        .to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 1);
        let frag = &v["pages"][0]["fragments"][0];
        assert_eq!(frag["kind"], "table");
        assert_eq!(frag["y"], 96.0);
        assert_eq!(frag["height"], 24.0);
        assert_eq!(frag["rowStart"], 0);
        assert_eq!(frag["rowEnd"], 1);
    }

    #[test]
    fn floating_table_engages_the_floating_table_hook() {
        let input = serde_json::json!({
            "measured": [{
                "block": {"kind": "table", "id": 0, "rows": [], "floating": {}},
                "measure": {"kind": "table", "rows": [], "columnWidths": [], "totalWidth": 0.0, "totalHeight": 0.0}
            }],
            "options": options_json()
        })
        .to_string();
        assert_eq!(layout_to_json(&input).unwrap_err(), "UNSUPPORTED");
        assert!(matches!(
            compute_layout(&input),
            Err(LayoutError::Unsupported(_))
        ));
    }

    #[test]
    fn empty_document_still_yields_page_one() {
        let input = serde_json::json!({ "measured": [], "options": options_json() }).to_string();
        let out = layout_to_json(&input).unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        assert_eq!(v["pages"].as_array().unwrap().len(), 1);
        assert_eq!(v["pages"][0]["number"], 1);
    }
}
