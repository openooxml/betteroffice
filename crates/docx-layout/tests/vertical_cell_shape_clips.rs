use docx_layout::display_list::{DisplayList, DocAttrs, Primitive, build_display_list_json};
use serde_json::{Value, json};

const TEXT: &str = "abcdefghijklmno";

struct Turned {
    origin: (f64, f64),
    center: (f64, f64),
    clip: Option<(f64, f64, f64, f64)>,
}

fn primitives(input: &Value) -> Vec<Primitive> {
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).expect("parses");
    dl.pages
        .into_iter()
        .next()
        .map(|page| page.primitives)
        .unwrap_or_default()
}

fn clip_of(attrs: &DocAttrs) -> Option<(f64, f64, f64, f64)> {
    attrs
        .clip_group
        .as_ref()
        .and_then(|group| group.clip.as_ref())
        .map(|clip| {
            let v =
                |n: &Option<serde_json::Number>| n.as_ref().and_then(|n| n.as_f64()).unwrap_or(0.0);
            (v(&clip.x), v(&clip.y), v(&clip.w), v(&clip.h))
        })
}

fn group_shape_in_a_vertical_cell(direction: &str) -> Value {
    let block = json!({
        "kind": "shape", "id": "box", "width": 200, "height": 60,
        "scene": {"version": 1, "root": {"kind": "group", "id": "group-1"}},
        "textBodyProperties": {"margins": {"left": 10, "right": 10, "top": 5, "bottom": 5}},
        "innerText": [{"id": "para", "runs": [{"kind": "text", "text": TEXT}]}],
        "innerMeasures": [{
            "totalHeight": 20.0,
            "lines": [{"headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                "width": 40.0, "ascent": 16, "descent": 4, "lineHeight": 20}]
        }]
    });
    let measure = json!({"kind": "shape", "width": 200, "height": 60});
    json!({
        "measured": [{
            "block": {"kind": "table", "id": 9, "rows": [{"cells": [
                {"textDirection": direction, "blocks": [block]}
            ]}]},
            "measure": {"kind": "table", "columnWidths": [100.0], "totalHeight": 300.0,
                "rows": [{"height": 300.0, "cells": [{"blocks": [measure]}]}]}
        }],
        "options": {},
        "layout": {"pages": [{"size": {"w": 400.0, "h": 600.0}, "margins": {}, "fragments": [
            {"kind": "table", "blockId": 9, "x": 50.0, "y": 150.0, "width": 100.0,
                "height": 300.0, "rowStart": 0, "rowEnd": 1}
        ]}]}
    })
}

fn turned(input: &Value, rotation: f64) -> Vec<Turned> {
    primitives(input)
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Text(text) if !text.text.is_empty() => Some(text),
            _ => None,
        })
        .map(|text| {
            assert_eq!(
                text.rotation_deg.as_ref().and_then(|deg| deg.as_f64()),
                Some(rotation)
            );
            let size = text
                .font
                .split_whitespace()
                .find_map(|part| {
                    part.strip_suffix("px")
                        .and_then(|size| size.parse::<f64>().ok())
                })
                .unwrap_or(11.0 * 96.0 / 72.0);
            let width = text.width.as_f64().unwrap_or(f64::NAN);
            let center = (
                text.x.as_f64().unwrap_or(f64::NAN) + width / 2.0,
                text.baseline_y.as_f64().unwrap_or(f64::NAN) - size * 0.2,
            );
            let origin = if rotation < 0.0 {
                (center.0 + size * 0.2, center.1 + width / 2.0)
            } else {
                (center.0 - size * 0.2, center.1 - width / 2.0)
            };
            Turned {
                origin,
                center,
                clip: clip_of(&text.attrs),
            }
        })
        .collect()
}

fn contains(clip: (f64, f64, f64, f64), (x, y): (f64, f64)) -> bool {
    clip.0 <= x && x <= clip.0 + clip.2 && clip.1 <= y && y <= clip.1 + clip.3
}

fn assert_visible(texts: &[Turned]) {
    assert!(!texts.is_empty());
    for text in texts {
        assert!(
            text.clip.is_some_and(|(_, _, w, h)| w > 0.0 && h > 0.0),
            "{:?}",
            text.clip
        );
        let clip = text.clip.unwrap_or_default();
        assert!(contains(clip, text.origin), "{clip:?} {:?}", text.origin);
        assert!(contains(clip, text.center), "{clip:?} {:?}", text.center);
    }
}

#[test]
fn group_shape_text_in_a_tbrl_cell_keeps_a_visible_clip() {
    let texts = turned(&group_shape_in_a_vertical_cell("tbRl"), 90.0);
    assert_eq!(texts.len(), 1);
    assert_visible(&texts);
}

#[test]
fn group_shape_text_in_a_btlr_cell_keeps_a_visible_clip() {
    let texts = turned(&group_shape_in_a_vertical_cell("btLr"), -90.0);
    assert_eq!(texts.len(), 1);
    assert_visible(&texts);
}
