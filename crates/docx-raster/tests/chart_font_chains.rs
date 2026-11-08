//! Every font a display list paints with has a chain from the requirements of
//! the blocks it was built from.

use docx_layout::display_list::{DisplayList, build_display_list_json};
use docx_layout::measure_blocks::collect_font_requirements;
use docx_layout::types::LayoutBlock;
use docx_raster::{FontChains, ImageMap, RenderResources, render_png};
use ooxml_text::FontStore;
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("assets/Carlito-Regular.ttf");

fn renders_with_required_chains(block: Value, measure: Value, fragment: Value) {
    let input = json!({
        "measured": [{"block": block, "measure": measure}],
        "options": {},
        "layout": {"pages": [{
            "size": {"w": 400.0, "h": 300.0},
            "margins": {},
            "fragments": [fragment]
        }]}
    });
    let json = build_display_list_json(&input.to_string()).unwrap();
    let list: DisplayList = serde_json::from_str(&json).unwrap();
    let block: LayoutBlock = serde_json::from_value(input["measured"][0]["block"].clone()).unwrap();
    let mut fonts = FontStore::new();
    let id = fonts.register(FONT.to_vec()).unwrap();
    let chains: FontChains = collect_font_requirements([&block], "Calibri")
        .into_iter()
        .map(|requirement| (requirement.key, vec![id]))
        .collect();
    let images = ImageMap::new();
    render_png(&list, 0, &RenderResources::new(&fonts, &chains, &images)).unwrap();
}

fn renders_chart(chart: Value) {
    renders_with_required_chains(
        json!({
            "kind": "chart", "id": 42, "width": 260.0, "height": 180.0,
            "docStart": 4, "docEnd": 5, "chart": chart
        }),
        json!({"kind": "chart", "width": 260.0, "height": 180.0}),
        json!({
            "kind": "chart", "blockId": 42, "x": 50.0, "y": 40.0,
            "width": 260.0, "height": 180.0, "docStart": 4, "docEnd": 5
        }),
    );
}

fn column_chart() -> Value {
    json!({
        "type": "chart",
        "chartType": "column",
        "title": "Revenue",
        "legend": {"position": "right", "visible": true},
        "series": [{
            "name": "North",
            "categories": ["Q1", "Q2"],
            "values": [10.0, 20.0],
            "color": "#4472C4"
        }],
        "axes": {"value": {"min": 0.0, "max": 25.0}}
    })
}

#[test]
fn default_chart_fonts_have_chains() {
    renders_chart(column_chart());
}

#[test]
fn chart_text_properties_have_chains() {
    let mut chart = column_chart();
    chart["text"] = json!({"font": "Georgia", "sizePt": 9.0});
    chart["titleText"] = json!({"font": "Segoe UI Light", "bold": true, "italic": true});
    chart["legend"]["text"] = json!({"font": "Times New Roman", "italic": true});
    renders_chart(chart);
}
