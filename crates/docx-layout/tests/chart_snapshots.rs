//! Serialized display-list snapshots for every chart configuration the
//! extraction had to keep byte-identical.

use docx_layout::display_list::build_display_list_json;
use serde_json::{Map, Value, json};

fn display_list(chart: Value, width: f64, height: f64) -> Value {
    let input = json!({
        "measured": [{
            "block": {
                "kind": "chart",
                "id": 42,
                "width": width,
                "height": height,
                "docStart": 4,
                "docEnd": 5,
                "chart": chart
            },
            "measure": { "kind": "chart", "width": width, "height": height }
        }],
        "options": {},
        "layout": { "pages": [{
            "size": { "w": 400.0, "h": 300.0 },
            "margins": {},
            "fragments": [{
                "kind": "chart",
                "blockId": 42,
                "x": 50.0,
                "y": 40.0,
                "width": width,
                "height": height,
                "docStart": 4,
                "docEnd": 5
            }]
        }] }
    });
    let json = build_display_list_json(&input.to_string()).expect("display list builds");
    serde_json::from_str(&json).expect("display list is json")
}

fn fnv1a(value: &str) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in value.as_bytes() {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

/// Document attrs, which every chart primitive that carries any of them
/// carries identically, so the snapshot states them once per configuration.
const ATTR_KEYS: [&str; 9] = [
    "ariaDescription",
    "ariaLabel",
    "blockId",
    "chart",
    "decorative",
    "docEnd",
    "docStart",
    "sdt",
    "sdtPath",
];

/// One line per primitive: its serialized fields, minus the attrs hoisted into
/// the header, with long path command lists reduced to a digest.
fn snapshot(name: &str, chart: Value, width: f64, height: f64) -> String {
    let display_list = display_list(chart, width, height);
    let page = &display_list["pages"][0];
    let primitives = page["primitives"].as_array().expect("primitives");
    let mut out = format!("# {name}\n");
    let shared = primitives
        .iter()
        .map(attrs_of)
        .find(|attrs| !attrs.is_empty())
        .unwrap_or_default();
    out.push_str(&format!(
        "attrs {}\n",
        compact(&Value::Object(shared.clone()))
    ));
    for primitive in primitives {
        let attrs = attrs_of(primitive);
        assert!(
            attrs.is_empty() || attrs == shared,
            "{name}: chart primitives must share one set of attrs"
        );
        let mut object = primitive.as_object().cloned().unwrap_or_default();
        object.retain(|key, _| !ATTR_KEYS.contains(&key.as_str()));
        if let Some(path) = object.remove("geometryPath") {
            let commands = path.as_array().cloned().unwrap_or_default();
            object.insert(
                "geometryPath".to_owned(),
                json!(format!(
                    "{} commands {} .. {} #{:016x}",
                    commands.len(),
                    compact(commands.first().unwrap_or(&Value::Null)),
                    compact(commands.last().unwrap_or(&Value::Null)),
                    fnv1a(&compact(&path))
                )),
            );
        }
        out.push_str(&compact(&Value::Object(object)));
        out.push('\n');
    }
    out
}

fn attrs_of(primitive: &Value) -> Map<String, Value> {
    primitive
        .as_object()
        .map(|object| {
            object
                .iter()
                .filter(|(key, _)| ATTR_KEYS.contains(&key.as_str()))
                .map(|(key, value)| (key.clone(), value.clone()))
                .collect()
        })
        .unwrap_or_default()
}

fn compact(value: &Value) -> String {
    serde_json::to_string(value).expect("serializes")
}

fn series(name: &str, values: Value) -> Value {
    json!({
        "name": name,
        "categories": ["Q1", "Q2"],
        "values": values,
        "color": "#4472C4"
    })
}

fn basic(chart_type: &str) -> Value {
    json!({
        "type": "chart",
        "chartType": chart_type,
        "title": "Revenue",
        "legend": { "position": "right", "visible": true },
        "series": [series("North", json!([10.0, 20.0])), series("South", json!([4.0, 30.0]))],
        "axes": { "value": { "min": 0.0, "max": 25.0 } }
    })
}

fn legend(position: Option<&str>, visible: bool) -> Value {
    let mut chart = basic("column");
    chart["legend"] = match position {
        Some(position) => json!({ "position": position, "visible": visible }),
        None => json!({ "visible": visible }),
    };
    chart
}

fn configurations() -> Vec<(String, Value, f64, f64)> {
    let mut cases: Vec<(String, Value, f64, f64)> = Vec::new();
    for chart_type in [
        "column", "bar", "line", "pie", "doughnut", "area", "scatter", "radar", "stock", "bubble",
        "surface", "mystery", "",
    ] {
        cases.push((
            format!("type-{chart_type}"),
            basic(chart_type),
            260.0,
            180.0,
        ));
    }
    for position in ["left", "right", "top", "bottom"] {
        cases.push((
            format!("legend-{position}"),
            legend(Some(position), true),
            260.0,
            180.0,
        ));
    }
    cases.push((
        "legend-hidden".to_owned(),
        legend(None, false),
        260.0,
        180.0,
    ));
    cases.push((
        "legend-default-position".to_owned(),
        legend(None, true),
        260.0,
        180.0,
    ));
    cases.push((
        "legend-overflow".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "position": "right", "visible": true },
            "series": (0..12)
                .map(|i| json!({
                    "name": format!("Series {i}"),
                    "categories": ["Q1", "Q2"],
                    "values": [i as f64, (i * 2) as f64]
                }))
                .collect::<Vec<_>>()
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "legend-overflow-pie".to_owned(),
        json!({
            "type": "chart",
            "chartType": "pie",
            "legend": { "position": "right", "visible": true },
            "series": [{
                "name": "Share",
                "categories": ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"],
                "values": [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "combo-column-line".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "title": "Combo",
            "legend": { "position": "right", "visible": true },
            "plotGroups": [
                { "chartType": "column", "grouping": "clustered", "series": [series("Revenue", json!([5.0, 9.0]))] },
                { "chartType": "line", "series": [series("Trend", json!([4.0, 8.0]))] }
            ]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "combo-with-pie-group".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "position": "left", "visible": true },
            "plotGroups": [
                { "chartType": "pie", "series": [series("Share", json!([3.0, 1.0]))] },
                { "chartType": "bar", "series": [series("Bars", json!([2.0, 6.0]))] }
            ]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "points-markers-labels".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "title": "Points",
            "legend": { "position": "right", "visible": true },
            "series": [{
                "name": "North",
                "categories": ["Q1", "Q2", "Q3"],
                "values": [1.0, 2.0, 3.0],
                "color": "4472C4",
                "marker": { "size": 9.0 },
                "points": [
                    { "index": 0, "value": 7.0, "color": "FF0000", "label": "peak", "marker": { "size": 14.0 } },
                    { "index": 2, "color": "#00FF00", "label": "end" }
                ]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "points-without-indexes".to_owned(),
        json!({
            "type": "chart",
            "chartType": "pie",
            "series": [{
                "name": "Share",
                "categories": ["Q1", "Q2"],
                "values": [3.0, 1.0],
                "points": [{ "color": "123456", "explosion": 25.0 }]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "negative-values".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "position": "right", "visible": true },
            "series": [series("North", json!([-10.0, 20.0])), series("South", json!([-4.0, -30.0]))]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "negative-values-bar".to_owned(),
        json!({
            "type": "chart",
            "chartType": "bar",
            "series": [series("North", json!([-10.0, 20.0]))]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "inverted-axis-bounds".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "series": [series("North", json!([1.0, 2.0]))],
            "axes": { "value": { "min": 10.0, "max": -10.0 } }
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "long-text".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "title": "T".repeat(200),
            "legend": { "position": "right", "visible": true },
            "series": [{
                "name": "N".repeat(200),
                "categories": ["C".repeat(200), "Q2"],
                "values": [1.0, 2.0]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "no-title-no-legend".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "legend": { "visible": false },
            "series": [series("North", json!([1.0, 2.0]))]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "empty-series".to_owned(),
        json!({ "type": "chart", "chartType": "column", "series": [] }),
        260.0,
        180.0,
    ));
    cases.push((
        "series-without-values".to_owned(),
        json!({
            "type": "chart",
            "chartType": "pie",
            "series": [{ "name": "Empty", "categories": [], "values": [] }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "described-and-decorative".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "title": "Revenue",
            "description": "quarterly revenue",
            "decorative": true,
            "series": [series("North", json!([1.0, 2.0]))]
        }),
        260.0,
        180.0,
    ));
    cases.push(("zero-rect".to_owned(), basic("column"), 0.0, 0.0));
    cases.push(("tiny-rect".to_owned(), basic("column"), 12.0, 8.0));
    cases.push(("tiny-rect-pie".to_owned(), basic("pie"), 12.0, 8.0));
    cases.push(("wide-flat-rect".to_owned(), basic("line"), 900.0, 26.0));
    cases
}

#[test]
fn chart_display_lists_match_their_serialized_snapshots() {
    let actual: String = configurations()
        .into_iter()
        .map(|(name, chart, width, height)| snapshot(&name, chart, width, height))
        .collect();
    if actual != EXPECTED {
        let mismatch = actual
            .lines()
            .zip(EXPECTED.lines())
            .find(|(actual, expected)| actual != expected);
        panic!(
            "chart display list drifted\n  actual:   {:?}\n  expected: {:?}\n\nfull output:\n{actual}",
            mismatch.map(|(actual, _)| actual),
            mismatch.map(|(_, expected)| expected)
        );
    }
}

#[test]
fn every_configuration_is_snapshotted_once() {
    let mut names: Vec<String> = configurations()
        .into_iter()
        .map(|(name, ..)| name)
        .collect();
    let total = names.len();
    names.sort();
    names.dedup();
    assert_eq!(names.len(), total, "configuration names must be unique");
    for name in &names {
        assert!(
            EXPECTED.contains(&format!("# {name}\n")),
            "{name} has no snapshot"
        );
    }
}

const EXPECTED: &str = r##"# type-column
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-bar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":100.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":36,"x":54}
{"fill":"#4472C4","h":20.65,"kind":"rect","w":41.6,"x":92,"y":76.85}
{"fill":"#4472C4","h":20.65,"kind":"rect","w":16.64,"x":92,"y":97.5}
{"baselineY":159.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":36,"x":54}
{"fill":"#4472C4","h":20.65,"kind":"rect","w":83.2,"x":92,"y":135.85}
{"fill":"#4472C4","h":20.65,"kind":"rect","w":104,"x":92,"y":156.5}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-line
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":136.8}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":89.6}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":165.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":66}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-pie
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, pie chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"fill":"#4472C4","geometryPath":"19 commands {\"type\":\"move\",\"x\":148.8,\"y\":137.92} .. {\"type\":\"close\"} #dbd8f6a1d048b0a7","h":103.36,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":103.36,"x":97.12,"y":86.24}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":148.8,\"y\":137.92} .. {\"type\":\"close\"} #ffd01596bf167317","h":103.36,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":103.36,"x":97.12,"y":86.24}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":224}
# type-doughnut
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, doughnut chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":148.8,\"y\":86.24} .. {\"type\":\"close\"} #869476812e97a961","h":103.36,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":103.36,"x":97.12,"y":86.24}
{"fill":"#4472C4","geometryPath":"67 commands {\"type\":\"move\",\"x\":193.556,\"y\":163.76} .. {\"type\":\"close\"} #3147718317ba0660","h":103.36,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":103.36,"x":97.12,"y":86.24}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":224}
# type-area
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-scatter
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":136.8}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":89.6}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":165.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":66}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-radar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":136.8}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":89.6}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":165.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":66}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-stock
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-bubble
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-surface
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-mystery
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-left
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":158}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":196,"x2":196,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":196,"x2":300,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":198}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":203.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":222,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":250}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":255.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":274,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":56,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":56,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":68}
# legend-right
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-top
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-bottom
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-hidden
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":96,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":35,"x":107,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":35,"x":142,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":96,"x":194}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":35,"x":207,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":35,"x":242,"y":68}
# legend-default-position
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":18.2,"x":99.8,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":18.2,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":18.2,"x":151.8,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":18.2,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-overflow
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 12 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"22","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"16.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"11","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":1,"kind":"rect","w":3.033,"x":99.8,"y":186}
{"fill":"#ED7D31","h":6.182,"kind":"rect","w":3.033,"x":102.833,"y":179.818}
{"fill":"#A5A5A5","h":12.364,"kind":"rect","w":3.033,"x":105.867,"y":173.636}
{"fill":"#FFC000","h":18.545,"kind":"rect","w":3.033,"x":108.9,"y":167.455}
{"fill":"#5B9BD5","h":24.727,"kind":"rect","w":3.033,"x":111.933,"y":161.273}
{"fill":"#70AD47","h":30.909,"kind":"rect","w":3.033,"x":114.967,"y":155.091}
{"fill":"#264478","h":37.091,"kind":"rect","w":3.033,"x":118,"y":148.909}
{"fill":"#9E480E","h":43.273,"kind":"rect","w":3.033,"x":121.033,"y":142.727}
{"fill":"#4472C4","h":49.455,"kind":"rect","w":3.033,"x":124.067,"y":136.545}
{"fill":"#ED7D31","h":55.636,"kind":"rect","w":3.033,"x":127.1,"y":130.364}
{"fill":"#A5A5A5","h":61.818,"kind":"rect","w":3.033,"x":130.133,"y":124.182}
{"fill":"#FFC000","h":68,"kind":"rect","w":3.033,"x":133.167,"y":118}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":1,"kind":"rect","w":3.033,"x":151.8,"y":186}
{"fill":"#ED7D31","h":12.364,"kind":"rect","w":3.033,"x":154.833,"y":173.636}
{"fill":"#A5A5A5","h":24.727,"kind":"rect","w":3.033,"x":157.867,"y":161.273}
{"fill":"#FFC000","h":37.091,"kind":"rect","w":3.033,"x":160.9,"y":148.909}
{"fill":"#5B9BD5","h":49.455,"kind":"rect","w":3.033,"x":163.933,"y":136.545}
{"fill":"#70AD47","h":61.818,"kind":"rect","w":3.033,"x":166.967,"y":124.182}
{"fill":"#264478","h":74.182,"kind":"rect","w":3.033,"x":170,"y":111.818}
{"fill":"#9E480E","h":86.545,"kind":"rect","w":3.033,"x":173.033,"y":99.455}
{"fill":"#4472C4","h":98.909,"kind":"rect","w":3.033,"x":176.067,"y":87.091}
{"fill":"#ED7D31","h":111.273,"kind":"rect","w":3.033,"x":179.1,"y":74.727}
{"fill":"#A5A5A5","h":123.636,"kind":"rect","w":3.033,"x":182.133,"y":62.364}
{"fill":"#FFC000","h":136,"kind":"rect","w":3.033,"x":185.167,"y":50}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 0","width":80,"x":224}
{"fill":"#ED7D31","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 1","width":80,"x":224}
{"fill":"#A5A5A5","h":8,"kind":"rect","w":8,"x":212,"y":88}
{"baselineY":96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 2","width":80,"x":224}
{"fill":"#FFC000","h":8,"kind":"rect","w":8,"x":212,"y":103}
{"baselineY":111,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 3","width":80,"x":224}
{"fill":"#5B9BD5","h":8,"kind":"rect","w":8,"x":212,"y":118}
{"baselineY":126,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 4","width":80,"x":224}
{"fill":"#70AD47","h":8,"kind":"rect","w":8,"x":212,"y":133}
{"baselineY":141,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 5","width":80,"x":224}
{"fill":"#264478","h":8,"kind":"rect","w":8,"x":212,"y":148}
{"baselineY":156,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 6","width":80,"x":224}
{"fill":"#9E480E","h":8,"kind":"rect","w":8,"x":212,"y":163}
{"baselineY":171,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 7","width":80,"x":224}
# legend-overflow-pie
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 10 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #a5cb8f4bdf58ce94","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#ED7D31","geometryPath":"5 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #6d5f636ea8310f30","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#A5A5A5","geometryPath":"6 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #7043fa35c3f78e31","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#FFC000","geometryPath":"7 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #f1bc986a3360f140","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#5B9BD5","geometryPath":"8 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #57bcf7110f4c4308","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#70AD47","geometryPath":"9 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #ae32e6f1f020477e","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#264478","geometryPath":"10 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #5b768921754a2bb4","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#9E480E","geometryPath":"10 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #8235c1a7609a7698","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #f36b3aa201c80b50","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#ED7D31","geometryPath":"12 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #c28c89706d6d9491","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"a","width":80,"x":224}
{"fill":"#ED7D31","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"b","width":80,"x":224}
{"fill":"#A5A5A5","h":8,"kind":"rect","w":8,"x":212,"y":88}
{"baselineY":96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"c","width":80,"x":224}
{"fill":"#FFC000","h":8,"kind":"rect","w":8,"x":212,"y":103}
{"baselineY":111,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"d","width":80,"x":224}
{"fill":"#5B9BD5","h":8,"kind":"rect","w":8,"x":212,"y":118}
{"baselineY":126,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"e","width":80,"x":224}
{"fill":"#70AD47","h":8,"kind":"rect","w":8,"x":212,"y":133}
{"baselineY":141,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"f","width":80,"x":224}
{"fill":"#264478","h":8,"kind":"rect","w":8,"x":212,"y":148}
{"baselineY":156,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"g","width":80,"x":224}
{"fill":"#9E480E","h":8,"kind":"rect","w":8,"x":212,"y":163}
{"baselineY":171,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"h","width":80,"x":224}
# combo-column-line
attrs {"ariaLabel":"Combo","blockId":42,"chart":{"label":"Combo, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Combo","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":65.556,"kind":"rect","w":36.4,"x":99.8,"y":120.444}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":118,"kind":"rect","w":36.4,"x":151.8,"y":68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":125}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":127,"y2":68}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":66}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Revenue","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Trend","width":80,"x":224}
# combo-with-pie-group
attrs {"blockId":42,"chart":{"label":"Untitled chart, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"39 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #0004896313fec87b","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#4472C4","geometryPath":"15 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #606b0a726cc1069a","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4.5","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":158}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":196,"x2":196,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":196,"x2":300,"y1":186,"y2":186}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":36,"x":158}
{"fill":"#4472C4","h":47.6,"kind":"rect","w":34.667,"x":196,"y":60.2}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":36,"x":158}
{"fill":"#4472C4","h":47.6,"kind":"rect","w":104,"x":196,"y":128.2}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":56,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":56,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":68}
# points-markers-labels
attrs {"ariaLabel":"Points","blockId":42,"chart":{"label":"Points, line chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Points","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":128}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":32,"x":180}
{"fill":"#FF0000","h":14,"kind":"rect","w":14,"x":85,"y":61}
{"baselineY":54,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"peak","width":48,"x":106}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":144,"y1":68,"y2":152.286}
{"fill":"#4472C4","h":9,"kind":"rect","w":9,"x":139.5,"y":147.786}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":144,"x2":196,"y1":152.286,"y2":135.429}
{"fill":"#00FF00","h":9,"kind":"rect","w":9,"x":191.5,"y":130.929}
{"baselineY":126.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"end","width":48,"x":205}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# points-without-indexes
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#123456","geometryPath":"39 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #0004896313fec87b","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#123456","geometryPath":"15 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #606b0a726cc1069a","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#123456","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":224}
{"fill":"#123456","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":224}
# negative-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-17.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-30","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":18.2,"x":99.8,"y":104.4}
{"fill":"#4472C4","h":10.88,"kind":"rect","w":18.2,"x":118,"y":104.4}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":54.4,"kind":"rect","w":18.2,"x":151.8,"y":50}
{"fill":"#4472C4","h":81.6,"kind":"rect","w":18.2,"x":170,"y":104.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# negative-values-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":36,"x":54}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":36,"x":54}
{"fill":"#4472C4","h":47.6,"kind":"rect","w":104,"x":92,"y":128.2}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# inverted-axis-bounds
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"11","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":1,"kind":"rect","w":36.4,"x":99.8,"y":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":1,"kind":"rect","w":36.4,"x":151.8,"y":186}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# long-text
attrs {"ariaLabel":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT","blockId":42,"chart":{"label":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC","width":48,"x":94}
{"fill":"#4472C4","h":59,"kind":"rect","w":36.4,"x":99.8,"y":127}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":118,"kind":"rect","w":36.4,"x":151.8,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN","width":80,"x":224}
# no-title-no-legend
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":276}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":116}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":292,"y1":118,"y2":50}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":290,"y":48}
# empty-series
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 0 series, 0 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
# series-without-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 0 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
# described-and-decorative
attrs {"ariaDescription":"quarterly revenue","ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 1 series, 2 categories"},"decorative":true,"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":59,"kind":"rect","w":36.4,"x":99.8,"y":127}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":118,"kind":"rect","w":36.4,"x":151.8,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# zero-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":74,"y2":74}
{"baselineY":77,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":80,"y2":80}
{"baselineY":83,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":86,"y2":86}
{"baselineY":89,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":8,"x":94}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":4.2,"x":93.8,"y":82.4}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":4.2,"x":98,"y":88.16}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":8,"x":106}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":4.2,"x":105.8,"y":72.8}
{"fill":"#4472C4","h":24,"kind":"rect","w":4.2,"x":110,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-48,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":-36}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-48,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":-36}
# tiny-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":74,"y2":74}
{"baselineY":77,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":80,"y2":80}
{"baselineY":83,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":86,"y2":86}
{"baselineY":89,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":8,"x":94}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":4.2,"x":93.8,"y":82.4}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":4.2,"x":98,"y":88.16}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":8,"x":106}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":4.2,"x":105.8,"y":72.8}
{"fill":"#4472C4","h":24,"kind":"rect","w":4.2,"x":110,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-36,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":-24}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-36,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":-24}
# tiny-rect-pie
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, pie chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"19 commands {\"type\":\"move\",\"x\":54.56,\"y\":58.8} .. {\"type\":\"close\"} #ee78e23431aee3ff","h":20,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20,"x":44.56,"y":48.8}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":54.56,\"y\":58.8} .. {\"type\":\"close\"} #dae2117a220e03bb","h":20,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20,"x":44.56,"y":48.8}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-36,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":-24}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-36,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":-24}
# wide-flat-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":26,"kind":"rect","w":900,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":884,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":836,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":836,"y1":74,"y2":74}
{"baselineY":77,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"18.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":836,"y1":80,"y2":80}
{"baselineY":83,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"12.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":836,"y1":86,"y2":86}
{"baselineY":89,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":836,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":836,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":820}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":80.4}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":836,"y1":82.4,"y2":72.8}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":834,"y":70.8}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":86.16}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":836,"y1":88.16,"y2":68}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":834,"y":66}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":852,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":864}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":852,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":864}
"##;
