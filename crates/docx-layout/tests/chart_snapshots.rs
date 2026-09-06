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
    for grouping in ["stacked", "percentStacked"] {
        for chart_type in ["column", "bar", "line", "area"] {
            cases.push((
                format!("{grouping}-{chart_type}"),
                json!({
                    "type": "chart",
                    "chartType": chart_type,
                    "plotGroups": [{
                        "chartType": chart_type,
                        "grouping": grouping,
                        "series": [series("North", json!([10.0, 20.0])), series("South", json!([5.0, -8.0]))]
                    }]
                }),
                260.0,
                180.0,
            ));
        }
    }
    cases.push((
        "bar-gap-and-overlap".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "plotGroups": [{
                "chartType": "column",
                "grouping": "clustered",
                "gapWidth": 40.0,
                "overlap": -20.0,
                "series": [series("North", json!([10.0, 20.0])), series("South", json!([4.0, 30.0]))]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "scatter-xy".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "plotGroups": [{
                "chartType": "scatter",
                "scatterStyle": "lineMarker",
                "series": [{
                    "name": "XY",
                    "categories": [],
                    "values": [3.0, 9.0, 4.0],
                    "xValues": [1.0, 5.0, 9.0],
                    "marker": { "symbol": "diamond", "size": 7.0 }
                }]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "bubble-sizes".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "plotGroups": [{
                "chartType": "bubble",
                "bubbleScale": 120.0,
                "sizeRepresents": "area",
                "series": [{
                    "name": "Bubbles",
                    "categories": [],
                    "values": [3.0, 9.0],
                    "xValues": [1.0, 5.0],
                    "bubbleSizes": [1.0, 4.0]
                }]
            }]
        }),
        260.0,
        180.0,
    ));
    for style in ["standard", "marker", "filled"] {
        cases.push((
            format!("radar-{style}"),
            json!({
                "type": "chart",
                "chartType": "line",
                "plotGroups": [{
                    "chartType": "radar",
                    "radarStyle": style,
                    "series": [{
                        "name": "Skills",
                        "categories": ["A", "B", "C", "D"],
                        "values": [3.0, 9.0, 4.0, 6.0]
                    }]
                }]
            }),
            260.0,
            180.0,
        ));
    }
    cases.push((
        "stock-ohlc".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "plotGroups": [{
                "chartType": "stock",
                "hiLowLines": true,
                "upDownBars": true,
                "series": [
                    { "name": "Open", "categories": ["D1", "D2"], "values": [10.0, 12.0] },
                    { "name": "High", "categories": ["D1", "D2"], "values": [20.0, 22.0] },
                    { "name": "Low", "categories": ["D1", "D2"], "values": [5.0, 6.0] },
                    { "name": "Close", "categories": ["D1", "D2"], "values": [18.0, 8.0] }
                ]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "surface-contour".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "plotGroups": [{
                "chartType": "surface",
                "series": [
                    { "name": "R1", "categories": ["A", "B", "C"], "values": [1.0, 5.0, 9.0] },
                    { "name": "R2", "categories": ["A", "B", "C"], "values": [9.0, 1.0, 5.0] }
                ]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "doughnut-hole-and-rotation".to_owned(),
        json!({
            "type": "chart",
            "chartType": "doughnut",
            "plotGroups": [{
                "chartType": "doughnut",
                "holeSize": 25.0,
                "firstSliceAngle": 90.0,
                "varyColors": true,
                "series": [{
                    "name": "Share",
                    "categories": ["Q1", "Q2", "Q3"],
                    "values": [3.0, 1.0, 2.0],
                    "points": [{ "index": 1, "explosion": 30.0 }]
                }]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "secondary-value-axis".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "visible": false },
            "axisList": [
                { "id": "1", "axisType": "category" },
                { "id": "2", "axisType": "value", "min": 0.0, "max": 20.0, "majorGridlines": true },
                { "id": "3", "axisType": "value", "min": 0.0, "max": 100.0 }
            ],
            "plotGroups": [
                { "chartType": "column", "axisIds": ["1", "2"], "series": [series("Units", json!([10.0, 15.0]))] },
                { "chartType": "line", "axisIds": ["1", "3"], "series": [series("Rate", json!([40.0, 80.0]))] }
            ]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "log-scale-and-ticks".to_owned(),
        json!({
            "type": "chart",
            "chartType": "line",
            "legend": { "visible": false },
            "axisList": [{
                "id": "1",
                "axisType": "value",
                "min": 1.0,
                "max": 1000.0,
                "logarithmicBase": 10.0,
                "majorUnit": 100.0,
                "majorTickMark": "out",
                "majorGridlines": true,
                "numberFormat": "#,##0"
            }],
            "plotGroups": [{
                "chartType": "line",
                "axisIds": ["1"],
                "series": [series("Growth", json!([1.0, 500.0]))]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "reversed-axes".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "visible": false },
            "axisList": [
                { "id": "1", "axisType": "category", "reversed": true },
                { "id": "2", "axisType": "value", "min": 0.0, "max": 20.0, "reversed": true }
            ],
            "plotGroups": [{
                "chartType": "column",
                "axisIds": ["1", "2"],
                "series": [series("North", json!([4.0, 16.0]))]
            }]
        }),
        260.0,
        180.0,
    ));
    for symbol in [
        "circle", "diamond", "triangle", "square", "star", "plus", "dash", "dot", "x", "auto",
        "none",
    ] {
        cases.push((
            format!("marker-{symbol}"),
            json!({
                "type": "chart",
                "chartType": "line",
                "legend": { "visible": false },
                "series": [{
                    "name": "North",
                    "categories": ["Q1", "Q2"],
                    "values": [1.0, 2.0],
                    "marker": { "symbol": symbol, "size": 10.0 }
                }]
            }),
            160.0,
            120.0,
        ));
    }
    cases.push((
        "data-labels-composed".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "visible": false },
            "plotGroups": [{
                "chartType": "column",
                "dataLabels": { "showValue": true, "showCategoryName": true, "numberFormat": "0.0" },
                "series": [{
                    "name": "North",
                    "categories": ["Q1", "Q2"],
                    "values": [10.0, 20.0],
                    "dataLabels": { "showSeriesName": true, "separator": " / ", "position": "inEnd" }
                }]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "data-labels-percent-and-key".to_owned(),
        json!({
            "type": "chart",
            "chartType": "pie",
            "legend": { "visible": false },
            "plotGroups": [{
                "chartType": "pie",
                "varyColors": true,
                "dataLabels": { "showPercent": true, "showLegendKey": true, "position": "outEnd" },
                "series": [{ "name": "Share", "categories": ["Q1", "Q2"], "values": [3.0, 1.0] }]
            }]
        }),
        260.0,
        180.0,
    ));
    cases.push((
        "text-properties".to_owned(),
        json!({
            "type": "chart",
            "chartType": "column",
            "title": "Revenue",
            "legend": { "position": "right", "visible": true, "text": { "bold": true } },
            "text": { "font": "Georgia", "sizePt": 9.0, "color": "#112233" },
            "titleText": { "sizePt": 21.0, "italic": true },
            "axisList": [
                { "id": "0", "axisType": "category", "text": { "sizePt": 15.0 } },
                {
                    "id": "1",
                    "axisType": "value",
                    "min": 0.0,
                    "max": 20.0,
                    "majorGridlines": true,
                    "text": { "sizePt": 6.0, "color": "#884400" }
                }
            ],
            "plotGroups": [{
                "chartType": "column",
                "axisIds": ["0", "1"],
                "dataLabels": { "showValue": true, "text": { "bold": true, "sizePt": 12.0 } },
                "series": [series("North", json!([10.0, 20.0]))]
            }]
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":14.857,"x":103.143,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":14.857,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":14.857,"x":155.143,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":14.857,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-bar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":182,"x2":182,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":32,"x":166}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154,"x2":154,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":138}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":196,"y1":186,"y2":186}
{"baselineY":159.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":28,"x":126,"y":139.643}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":11.2,"x":126,"y":156.5}
{"baselineY":100.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":56,"x":126,"y":80.643}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":70,"x":126,"y":97.5}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
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
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":148.8,\"y\":86.24} .. {\"type\":\"close\"} #fc2e8e02affae282","h":103.36,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":103.36,"x":97.12,"y":86.24}
{"fill":"#4472C4","geometryPath":"67 commands {\"type\":\"move\",\"x\":193.556,\"y\":163.76} .. {\"type\":\"close\"} #9608858880dfe619","h":103.36,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":103.36,"x":97.12,"y":86.24}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":224}
# type-area
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":138.8} .. {\"type\":\"close\"} #23353b6cd0a66a80","h":118,"kind":"shape","w":104,"x":92,"y":68}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":92,"x2":196,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":167.12} .. {\"type\":\"close\"} #d33d16bd5f79c6a5","h":118,"kind":"shape","w":104,"x":92,"y":68}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":92,"x2":196,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-scatter
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, scatter chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-radar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, radar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":133.664,"y2":154.336}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":123.328,"y2":164.672}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":112.992,"y2":175.008}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":102.656,"y2":185.344}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":92.32,"y2":195.68}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":144,"y2":92.32}
{"baselineY":87.152,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":144,"y2":195.68}
{"baselineY":200.848,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":132.8}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":148.8,"x2":148.8,"y1":123.328,"y2":185.344}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":146.8,"y":121.328}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":146.8,"y":183.344}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":148.8,"x2":148.8,"y1":135.731,"y2":195.68}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":146.8,"y":133.731}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":146.8,"y":193.68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-stock
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, stock chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
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
# type-bubble
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bubble chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# type-surface
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, surface chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"fill":"#70AD47","h":59,"kind":"rect","w":52,"x":92,"y":127}
{"fill":"#ED7D31","h":59,"kind":"rect","w":52,"x":144,"y":127}
{"baselineY":162.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":36,"x":54}
{"fill":"#4472C4","h":59,"kind":"rect","w":52,"x":92,"y":68}
{"fill":"#9E480E","h":59,"kind":"rect","w":52,"x":144,"y":68}
{"baselineY":103.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":36,"x":54}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":14.857,"x":103.143,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":14.857,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":14.857,"x":155.143,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":14.857,"x":170,"y":68}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":14.857,"x":103.143,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":14.857,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":14.857,"x":155.143,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":14.857,"x":170,"y":68}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":158}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":300,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":158}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":196,"x2":196,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":196,"x2":300,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":198}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":14.857,"x":207.143,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":14.857,"x":222,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":250}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":14.857,"x":259.143,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":14.857,"x":274,"y":68}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":14.857,"x":103.143,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":14.857,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":14.857,"x":155.143,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":14.857,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-top
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":109.2,"y2":109.2}
{"baselineY":112.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":147.6,"y2":147.6}
{"baselineY":150.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":90,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":96,"x":94}
{"fill":"#4472C4","h":38.4,"kind":"rect","w":28.571,"x":113.429,"y":147.6}
{"fill":"#4472C4","h":15.36,"kind":"rect","w":28.571,"x":142,"y":170.64}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":96,"x":194}
{"fill":"#4472C4","h":76.8,"kind":"rect","w":28.571,"x":213.429,"y":109.2}
{"fill":"#4472C4","h":96,"kind":"rect","w":28.571,"x":242,"y":90}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":136,"y":75}
{"baselineY":82.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":39,"x":148}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":187,"y":75}
{"baselineY":82.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":39,"x":199}
# legend-bottom
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":87.2,"y2":87.2}
{"baselineY":90.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":125.6,"y2":125.6}
{"baselineY":128.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":164,"y2":164}
{"baselineY":167,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":164}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":164,"y2":164}
{"baselineY":178,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":96,"x":94}
{"fill":"#4472C4","h":38.4,"kind":"rect","w":28.571,"x":113.429,"y":125.6}
{"fill":"#4472C4","h":15.36,"kind":"rect","w":28.571,"x":142,"y":148.64}
{"baselineY":178,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":96,"x":194}
{"fill":"#4472C4","h":76.8,"kind":"rect","w":28.571,"x":213.429,"y":87.2}
{"fill":"#4472C4","h":96,"kind":"rect","w":28.571,"x":242,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":136,"y":205}
{"baselineY":212.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":39,"x":148}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":187,"y":205}
{"baselineY":212.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":39,"x":199}
# legend-hidden
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":96,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":28.571,"x":113.429,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":28.571,"x":142,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":96,"x":194}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":28.571,"x":213.429,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":28.571,"x":242,"y":68}
# legend-default-position
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":14.857,"x":103.143,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":14.857,"x":118,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":14.857,"x":155.143,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":14.857,"x":170,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# legend-overflow
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 12 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":1,"kind":"rect","w":3.852,"x":94.889,"y":186}
{"fill":"#ED7D31","h":5.44,"kind":"rect","w":3.852,"x":98.741,"y":180.56}
{"fill":"#A5A5A5","h":10.88,"kind":"rect","w":3.852,"x":102.593,"y":175.12}
{"fill":"#FFC000","h":16.32,"kind":"rect","w":3.852,"x":106.444,"y":169.68}
{"fill":"#5B9BD5","h":21.76,"kind":"rect","w":3.852,"x":110.296,"y":164.24}
{"fill":"#70AD47","h":27.2,"kind":"rect","w":3.852,"x":114.148,"y":158.8}
{"fill":"#264478","h":32.64,"kind":"rect","w":3.852,"x":118,"y":153.36}
{"fill":"#9E480E","h":38.08,"kind":"rect","w":3.852,"x":121.852,"y":147.92}
{"fill":"#4472C4","h":43.52,"kind":"rect","w":3.852,"x":125.704,"y":142.48}
{"fill":"#ED7D31","h":48.96,"kind":"rect","w":3.852,"x":129.556,"y":137.04}
{"fill":"#A5A5A5","h":54.4,"kind":"rect","w":3.852,"x":133.407,"y":131.6}
{"fill":"#FFC000","h":59.84,"kind":"rect","w":3.852,"x":137.259,"y":126.16}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":1,"kind":"rect","w":3.852,"x":146.889,"y":186}
{"fill":"#ED7D31","h":10.88,"kind":"rect","w":3.852,"x":150.741,"y":175.12}
{"fill":"#A5A5A5","h":21.76,"kind":"rect","w":3.852,"x":154.593,"y":164.24}
{"fill":"#FFC000","h":32.64,"kind":"rect","w":3.852,"x":158.444,"y":153.36}
{"fill":"#5B9BD5","h":43.52,"kind":"rect","w":3.852,"x":162.296,"y":142.48}
{"fill":"#70AD47","h":54.4,"kind":"rect","w":3.852,"x":166.148,"y":131.6}
{"fill":"#264478","h":65.28,"kind":"rect","w":3.852,"x":170,"y":120.72}
{"fill":"#9E480E","h":76.16,"kind":"rect","w":3.852,"x":173.852,"y":109.84}
{"fill":"#4472C4","h":87.04,"kind":"rect","w":3.852,"x":177.704,"y":98.96}
{"fill":"#ED7D31","h":97.92,"kind":"rect","w":3.852,"x":181.556,"y":88.08}
{"fill":"#A5A5A5","h":108.8,"kind":"rect","w":3.852,"x":185.407,"y":77.2}
{"fill":"#FFC000","h":119.68,"kind":"rect","w":3.852,"x":189.259,"y":66.32}
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
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":59,"kind":"rect","w":20.8,"x":107.6,"y":127}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":106.2,"kind":"rect","w":20.8,"x":159.6,"y":79.8}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":300,"x2":300,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":32,"x":284}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":276.667,"x2":276.667,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":32,"x":260.667}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":253.333,"x2":253.333,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":32,"x":237.333}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":230,"x2":230,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":214}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":230,"x2":230,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":230,"x2":300,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":158}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":23.333,"x":230,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":158}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":70,"x":230,"y":70.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":56,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":56,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":68}
# points-markers-labels
attrs {"ariaLabel":"Points","blockId":42,"chart":{"label":"Points, line chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Points","width":244,"x":58}
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
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":128}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":32,"x":180}
{"fill":"#FF0000","h":14,"kind":"rect","w":14,"x":85,"y":75.75}
{"baselineY":68.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"peak","width":48,"x":106}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":144,"y1":82.75,"y2":156.5}
{"fill":"#4472C4","h":9,"kind":"rect","w":9,"x":139.5,"y":152}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":144,"x2":196,"y1":156.5,"y2":141.75}
{"fill":"#00FF00","h":9,"kind":"rect","w":9,"x":191.5,"y":137.25}
{"baselineY":132.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"end","width":48,"x":205}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# points-without-indexes
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#123456","geometryPath":"39 commands {\"type\":\"move\",\"x\":159.018,\"y\":138.418} .. {\"type\":\"close\"} #97c43048994d90e3","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":101.218,"y":80.618}
{"fill":"#123456","geometryPath":"15 commands {\"type\":\"move\",\"x\":138.582,\"y\":117.982} .. {\"type\":\"close\"} #ddaa118fdc716b03","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":80.782,"y":60.182}
{"fill":"#123456","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":224}
{"fill":"#123456","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":224}
# negative-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-30","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":14.857,"x":103.143,"y":104.4}
{"fill":"#4472C4","h":10.88,"kind":"rect","w":14.857,"x":118,"y":104.4}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":54.4,"kind":"rect","w":14.857,"x":155.143,"y":50}
{"fill":"#4472C4","h":81.6,"kind":"rect","w":14.857,"x":170,"y":104.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# negative-values-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":196,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":32,"x":180}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":172.667,"x2":172.667,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":156.667}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":149.333,"x2":149.333,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":133.333}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":196,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":23.333,"x":126,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":46.667,"x":149.333,"y":70.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# inverted-axis-bounds
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"11","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":1,"kind":"rect","w":20.8,"x":107.6,"y":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":1,"kind":"rect","w":20.8,"x":159.6,"y":186}
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
{"fill":"#4472C4","h":59,"kind":"rect","w":20.8,"x":107.6,"y":127}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":118,"kind":"rect","w":20.8,"x":159.6,"y":68}
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
{"fill":"#4472C4","h":59,"kind":"rect","w":20.8,"x":107.6,"y":127}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":118,"kind":"rect","w":20.8,"x":159.6,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
# stacked-column
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":72.667,"y2":72.667}
{"baselineY":75.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":95.333,"y2":95.333}
{"baselineY":98.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":140.667,"y2":140.667}
{"baselineY":143.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":163.333,"y2":163.333}
{"baselineY":166.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":45.333,"kind":"rect","w":20.8,"x":107.6,"y":95.333}
{"fill":"#4472C4","h":22.667,"kind":"rect","w":20.8,"x":107.6,"y":72.667}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":90.667,"kind":"rect","w":20.8,"x":159.6,"y":50}
{"fill":"#4472C4","h":36.267,"kind":"rect","w":20.8,"x":159.6,"y":140.667}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# stacked-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":196,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":32,"x":180}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":172.667,"x2":172.667,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":156.667}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":149.333,"x2":149.333,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":133.333}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":196,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":23.333,"x":149.333,"y":138.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":11.667,"x":172.667,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":46.667,"x":149.333,"y":70.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":18.667,"x":130.667,"y":70.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# stacked-line
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":72.667,"y2":72.667}
{"baselineY":75.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":95.333,"y2":95.333}
{"baselineY":98.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":140.667,"y2":140.667}
{"baselineY":143.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":163.333,"y2":163.333}
{"baselineY":166.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":93.333}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":95.333,"y2":50}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":48}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":70.667}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":72.667,"y2":176.933}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":174.933}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# stacked-area
attrs {"blockId":42,"chart":{"label":"Untitled chart, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":72.667,"y2":72.667}
{"baselineY":75.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":95.333,"y2":95.333}
{"baselineY":98.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":140.667,"y2":140.667}
{"baselineY":143.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":163.333,"y2":163.333}
{"baselineY":166.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":95.333} .. {\"type\":\"close\"} #cd8522bcf23fca6e","h":136,"kind":"shape","w":104,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":92,"x2":196,"y1":95.333,"y2":50}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":72.667} .. {\"type\":\"close\"} #10776dd07a28299c","h":136,"kind":"shape","w":104,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":92,"x2":196,"y1":72.667,"y2":176.933}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# percentStacked-column
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":64.762,"kind":"rect","w":20.8,"x":107.6,"y":82.381}
{"fill":"#4472C4","h":32.381,"kind":"rect","w":20.8,"x":107.6,"y":50}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":69.388,"kind":"rect","w":20.8,"x":159.6,"y":77.755}
{"fill":"#4472C4","h":27.755,"kind":"rect","w":20.8,"x":159.6,"y":147.143}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# percentStacked-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196,"x2":196,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":32,"x":180}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":172.667,"x2":172.667,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"50%","width":32,"x":156.667}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":149.333,"x2":149.333,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":32,"x":133.333}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-50%","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":196,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":31.111,"x":149.333,"y":138.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":15.556,"x":180.444,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":33.333,"x":149.333,"y":70.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":13.333,"x":136,"y":70.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# percentStacked-line
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":80.381}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":82.381,"y2":77.755}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":75.755}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":48}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":196,"y1":50,"y2":174.898}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":194,"y":172.898}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# percentStacked-area
attrs {"blockId":42,"chart":{"label":"Untitled chart, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":180}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":82.381} .. {\"type\":\"close\"} #834d6a84b2a6359f","h":136,"kind":"shape","w":104,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":92,"x2":196,"y1":82.381,"y2":77.755}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":50} .. {\"type\":\"close\"} #11859ff3851a45aa","h":136,"kind":"shape","w":104,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":92,"x2":196,"y1":50,"y2":174.898}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# bar-gap-and-overlap
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":72.667,"y2":72.667}
{"baselineY":75.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":95.333,"y2":95.333}
{"baselineY":98.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":140.667,"y2":140.667}
{"baselineY":143.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":163.333,"y2":163.333}
{"baselineY":166.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":45.333,"kind":"rect","w":20,"x":96,"y":140.667}
{"fill":"#4472C4","h":18.133,"kind":"rect","w":20,"x":120,"y":167.867}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":90.667,"kind":"rect","w":20,"x":148,"y":95.333}
{"fill":"#4472C4","h":136,"kind":"rect","w":20,"x":172,"y":50}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":224}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":224}
# scatter-xy
attrs {"blockId":42,"chart":{"label":"Untitled chart, scatter chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":32,"x":96.8}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":32,"x":117.6}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":32,"x":138.4}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":32,"x":159.2}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":180}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":102.4,\"y\":141.7} .. {\"type\":\"close\"} #7fc6bd214bdd7770","h":7,"kind":"shape","w":7,"x":98.9,"y":141.7}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":102.4,"x2":144,"y1":145.2,"y2":63.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":144,\"y\":60.1} .. {\"type\":\"close\"} #6d4ee87452dd7de0","h":7,"kind":"shape","w":7,"x":140.5,"y":60.1}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":144,"x2":185.6,"y1":63.6,"y2":131.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":185.6,\"y\":128.1} .. {\"type\":\"close\"} #98fc4a37832ff6f1","h":7,"kind":"shape","w":7,"x":182.1,"y":128.1}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"XY","width":80,"x":224}
# bubble-sizes
attrs {"blockId":42,"chart":{"label":"Untitled chart, bubble chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":32,"x":102}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":32,"x":128}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":32,"x":154}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":32,"x":180}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":99.8,\"y\":145.2} .. {\"type\":\"close\"} #95d2d8c22e0179f8","h":15.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":15.6,"x":84.2,"y":137.4}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":211.6,\"y\":63.6} .. {\"type\":\"close\"} #c74739a8ab70e98d","h":31.2,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":31.2,"x":180.4,"y":48}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Bubbles","width":80,"x":224}
# radar-standard
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":160.36,"y1":123.44,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.36,"x2":148.8,"y1":135,"y2":146.56}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":137.24,"y1":146.56,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":137.24,"x2":148.8,"y1":135,"y2":123.44}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":171.92,"y1":111.88,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.92,"x2":148.8,"y1":135,"y2":158.12}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":125.68,"y1":158.12,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.68,"x2":148.8,"y1":135,"y2":111.88}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":183.48,"y1":100.32,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.48,"x2":148.8,"y1":135,"y2":169.68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":114.12,"y1":169.68,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":114.12,"x2":148.8,"y1":135,"y2":100.32}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":195.04,"y1":88.76,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":195.04,"x2":148.8,"y1":135,"y2":181.24}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":102.56,"y1":181.24,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.56,"x2":148.8,"y1":135,"y2":88.76}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":206.6,"y1":77.2,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":206.6,"x2":148.8,"y1":135,"y2":192.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":91,"y1":192.8,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":91,"x2":148.8,"y1":135,"y2":77.2}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":135,"y2":77.2}
{"baselineY":71.42,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":206.6,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":196.38}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":135,"y2":192.8}
{"baselineY":198.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":91,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":69.22}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":148.8,"x2":200.82,"y1":117.66,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":200.82,"x2":148.8,"y1":135,"y2":158.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":148.8,"x2":114.12,"y1":158.12,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.12,"x2":148.8,"y1":135,"y2":117.66}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":80,"x":224}
# radar-marker
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":160.36,"y1":123.44,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.36,"x2":148.8,"y1":135,"y2":146.56}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":137.24,"y1":146.56,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":137.24,"x2":148.8,"y1":135,"y2":123.44}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":171.92,"y1":111.88,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.92,"x2":148.8,"y1":135,"y2":158.12}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":125.68,"y1":158.12,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.68,"x2":148.8,"y1":135,"y2":111.88}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":183.48,"y1":100.32,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.48,"x2":148.8,"y1":135,"y2":169.68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":114.12,"y1":169.68,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":114.12,"x2":148.8,"y1":135,"y2":100.32}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":195.04,"y1":88.76,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":195.04,"x2":148.8,"y1":135,"y2":181.24}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":102.56,"y1":181.24,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.56,"x2":148.8,"y1":135,"y2":88.76}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":206.6,"y1":77.2,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":206.6,"x2":148.8,"y1":135,"y2":192.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":91,"y1":192.8,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":91,"x2":148.8,"y1":135,"y2":77.2}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":135,"y2":77.2}
{"baselineY":71.42,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":206.6,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":196.38}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":135,"y2":192.8}
{"baselineY":198.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":91,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":69.22}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":148.8,"x2":200.82,"y1":117.66,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":200.82,"x2":148.8,"y1":135,"y2":158.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":148.8,"x2":114.12,"y1":158.12,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.12,"x2":148.8,"y1":135,"y2":117.66}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":146.8,"y":115.66}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":198.82,"y":133}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":146.8,"y":156.12}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":112.12,"y":133}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":80,"x":224}
# radar-filled
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":160.36,"y1":123.44,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.36,"x2":148.8,"y1":135,"y2":146.56}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":137.24,"y1":146.56,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":137.24,"x2":148.8,"y1":135,"y2":123.44}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":171.92,"y1":111.88,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.92,"x2":148.8,"y1":135,"y2":158.12}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":125.68,"y1":158.12,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.68,"x2":148.8,"y1":135,"y2":111.88}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":183.48,"y1":100.32,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.48,"x2":148.8,"y1":135,"y2":169.68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":114.12,"y1":169.68,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":114.12,"x2":148.8,"y1":135,"y2":100.32}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":195.04,"y1":88.76,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":195.04,"x2":148.8,"y1":135,"y2":181.24}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":102.56,"y1":181.24,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.56,"x2":148.8,"y1":135,"y2":88.76}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":206.6,"y1":77.2,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":206.6,"x2":148.8,"y1":135,"y2":192.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":91,"y1":192.8,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":91,"x2":148.8,"y1":135,"y2":77.2}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":135,"y2":77.2}
{"baselineY":71.42,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":206.6,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":196.38}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":148.8,"y1":135,"y2":192.8}
{"baselineY":198.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":132.8}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":148.8,"x2":91,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":69.22}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":148.8,\"y\":117.66} .. {\"type\":\"close\"} #b9dc1494c202376f","h":136,"kind":"shape","w":104,"x":92,"y":50}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":80,"x":224}
# stock-ohlc
attrs {"blockId":42,"chart":{"label":"Untitled chart, stock chart, 4 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D1","width":48,"x":94}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":118,"x2":118,"y1":77.2,"y2":158.8}
{"fill":"#FFFFFF","h":43.52,"kind":"rect","w":24,"x":106,"y":88.08}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":106,"x2":106,"y1":88.08,"y2":131.6}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D2","width":48,"x":146}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":170,"x2":170,"y1":66.32,"y2":153.36}
{"fill":"#666666","h":21.76,"kind":"rect","w":24,"x":158,"y":120.72}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Open","width":80,"x":224}
{"fill":"#ED7D31","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"High","width":80,"x":224}
{"fill":"#A5A5A5","h":8,"kind":"rect","w":8,"x":212,"y":88}
{"baselineY":96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Low","width":80,"x":224}
{"fill":"#FFC000","h":8,"kind":"rect","w":8,"x":212,"y":103}
{"baselineY":111,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Close","width":80,"x":224}
# surface-contour
attrs {"blockId":42,"chart":{"label":"Untitled chart, surface chart, 2 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"fill":"#4472C4","h":68,"kind":"rect","w":34.667,"x":92,"y":118}
{"fill":"#A9D18E","h":68,"kind":"rect","w":34.667,"x":126.667,"y":118}
{"fill":"#ED7D31","h":68,"kind":"rect","w":34.667,"x":161.333,"y":118}
{"baselineY":158.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R1","width":36,"x":54}
{"fill":"#ED7D31","h":68,"kind":"rect","w":34.667,"x":92,"y":50}
{"fill":"#4472C4","h":68,"kind":"rect","w":34.667,"x":126.667,"y":50}
{"fill":"#A9D18E","h":68,"kind":"rect","w":34.667,"x":161.333,"y":50}
{"baselineY":90.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R2","width":36,"x":54}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":30.667,"x":94}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":30.667,"x":128.667}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":30.667,"x":163.333}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R1","width":80,"x":224}
{"fill":"#ED7D31","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R2","width":80,"x":224}
# doughnut-hole-and-rotation
attrs {"blockId":42,"chart":{"label":"Untitled chart, doughnut chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"51 commands {\"type\":\"move\",\"x\":206.6,\"y\":128.2} .. {\"type\":\"close\"} #b470cb314121b885","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#ED7D31","geometryPath":"19 commands {\"type\":\"move\",\"x\":75.983,\"y\":119.53} .. {\"type\":\"close\"} #6dd4a0146304183a","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":75.983,"y":61.73}
{"fill":"#A5A5A5","geometryPath":"35 commands {\"type\":\"move\",\"x\":119.9,\"y\":78.144} .. {\"type\":\"close\"} #c5682be995642018","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":58}
{"baselineY":66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":80,"x":224}
{"fill":"#ED7D31","h":8,"kind":"rect","w":8,"x":212,"y":73}
{"baselineY":81,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":80,"x":224}
{"fill":"#A5A5A5","h":8,"kind":"rect","w":8,"x":212,"y":88}
{"baselineY":96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":80,"x":224}
# secondary-value-axis
attrs {"blockId":42,"chart":{"label":"Untitled chart, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":254,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":254,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":254,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":254,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":254,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":254,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":77,"x":94}
{"fill":"#4472C4","h":68,"kind":"rect","w":32.4,"x":116.3,"y":118}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":77,"x":175}
{"fill":"#4472C4","h":102,"kind":"rect","w":32.4,"x":197.3,"y":84}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100","width":34,"x":258}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80","width":34,"x":258}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60","width":34,"x":258}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40","width":34,"x":258}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":258}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":258}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":254,"x2":254,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":238}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":129.6}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":254,"y1":131.6,"y2":77.2}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":252,"y":75.2}
# log-scale-and-ticks
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":95.333,"y2":95.333}
{"baselineY":98.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":88,"x2":92,"y1":95.333,"y2":95.333}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":140.667,"y2":140.667}
{"baselineY":143.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":88,"x2":92,"y1":140.667,"y2":140.667}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":88,"x2":92,"y1":186,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":276}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":90,"y":184}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":292,"y1":186,"y2":63.647}
{"fill":"#4472C4","h":4,"kind":"rect","w":4,"x":290,"y":61.647}
# reversed-axes
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":96,"x":194}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":40,"x":222,"y":50}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":96,"x":94}
{"fill":"#4472C4","h":108.8,"kind":"rect","w":40,"x":122,"y":50}
# marker-circle
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":97,\"y\":88} .. {\"type\":\"close\"} #67d6393dbf1cbb4a","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":197,\"y\":50} .. {\"type\":\"close\"} #44eff6309df8b96c","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-diamond
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":83} .. {\"type\":\"close\"} #9070cc319b54bc92","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":192,\"y\":45} .. {\"type\":\"close\"} #8428fffc2a676500","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-triangle
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"4 commands {\"type\":\"move\",\"x\":92,\"y\":83} .. {\"type\":\"close\"} #ee4f8aaf7aeb5cac","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"4 commands {\"type\":\"move\",\"x\":192,\"y\":45} .. {\"type\":\"close\"} #9562d8dd8d157e45","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-square
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","h":10,"kind":"rect","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","h":10,"kind":"rect","w":10,"x":187,"y":45}
# marker-star
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":92,\"y\":83} .. {\"type\":\"close\"} #4e39533cc7342889","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":192,\"y\":45} .. {\"type\":\"close\"} #41108fbd21c1f671","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-plus
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":90.5,\"y\":83} .. {\"type\":\"close\"} #0edda5464eb6271c","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":190.5,\"y\":45} .. {\"type\":\"close\"} #26b9203a8890dfe0","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-dash
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":87,\"y\":86.75} .. {\"type\":\"close\"} #db2bc8faa406da52","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":187,\"y\":48.75} .. {\"type\":\"close\"} #ccd3ec15463276dc","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-dot
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":94.5,\"y\":88} .. {\"type\":\"close\"} #0ed47a456e4a2238","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":194.5,\"y\":50} .. {\"type\":\"close\"} #d9d5e59346b6fcaa","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-x
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":94.475,\"y\":83.404} .. {\"type\":\"close\"} #60cd8731cde37f28","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":194.475,\"y\":45.404} .. {\"type\":\"close\"} #c3b0ae37845e837c","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-auto
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":83} .. {\"type\":\"close\"} #9070cc319b54bc92","h":10,"kind":"shape","w":10,"x":87,"y":83}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":192,\"y\":45} .. {\"type\":\"close\"} #8428fffc2a676500","h":10,"kind":"shape","w":10,"x":187,"y":45}
# marker-none
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":88,"y2":88}
{"baselineY":91,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":76}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":176}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":192,"y1":88,"y2":50}
# data-labels-composed
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":96,"x":94}
{"fill":"#4472C4","h":68,"kind":"rect","w":40,"x":122,"y":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North / Q1 / 10.0","width":40,"x":122}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":96,"x":194}
{"fill":"#4472C4","h":136,"kind":"rect","w":40,"x":222,"y":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North / Q2 / 20.0","width":40,"x":222}
# data-labels-percent-and-key
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"39 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #0004896313fec87b","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#4472C4","h":7,"kind":"rect","w":7,"x":185.801,"y":168.201}
{"baselineY":175.201,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"75%","width":48,"x":195.801}
{"fill":"#ED7D31","geometryPath":"15 commands {\"type\":\"move\",\"x\":148.8,\"y\":128.2} .. {\"type\":\"close\"} #606b0a726cc1069a","h":115.6,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":115.6,"x":91,"y":70.4}
{"fill":"#ED7D31","h":7,"kind":"rect","w":7,"x":91.799,"y":74.199}
{"baselineY":81.199,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25%","width":48,"x":101.799}
# text-properties
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#112233","font":"italic 600 28px Georgia","kind":"text","text":"Revenue","width":244,"x":58}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":68,"y2":68}
{"baselineY":71,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":127,"y2":127}
{"baselineY":130,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":189,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":196,"y1":186,"y2":186}
{"baselineY":200,"color":"#112233","font":"400 20px Georgia","kind":"text","text":"Q1","width":48,"x":94}
{"fill":"#4472C4","h":59,"kind":"rect","w":20.8,"x":107.6,"y":127}
{"baselineY":124,"color":"#112233","font":"700 16px Georgia","kind":"text","text":"10","width":32,"x":107.6}
{"baselineY":200,"color":"#112233","font":"400 20px Georgia","kind":"text","text":"Q2","width":48,"x":146}
{"fill":"#4472C4","h":118,"kind":"rect","w":20.8,"x":159.6,"y":68}
{"baselineY":65,"color":"#112233","font":"700 16px Georgia","kind":"text","text":"20","width":32,"x":159.6}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":212,"y":76}
{"baselineY":84,"color":"#112233","font":"700 12px Georgia","kind":"text","text":"North","width":80,"x":224}
# zero-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":72.8,"y2":72.8}
{"baselineY":75.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":8,"x":94}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":3.429,"x":94.571,"y":82.4}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":3.429,"x":98,"y":88.16}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":8,"x":106}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":3.429,"x":106.571,"y":72.8}
{"fill":"#4472C4","h":24,"kind":"rect","w":3.429,"x":110,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-48,"y":76}
{"baselineY":84,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":80,"x":-36}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":-48,"y":91}
{"baselineY":99,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":80,"x":-36}
# tiny-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":72.8,"y2":72.8}
{"baselineY":75.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":8,"x":94}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":3.429,"x":94.571,"y":82.4}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":3.429,"x":98,"y":88.16}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":8,"x":106}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":3.429,"x":106.571,"y":72.8}
{"fill":"#4472C4","h":24,"kind":"rect","w":3.429,"x":110,"y":68}
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
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":836,"y1":72.8,"y2":72.8}
{"baselineY":75.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
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
