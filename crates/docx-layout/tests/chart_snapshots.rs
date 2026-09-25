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
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":24.429,"x":100.821,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":24.429,"x":125.25,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":24.429,"x":186.321,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":24.429,"x":210.75,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-bar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":248.5,"x2":248.5,"y1":81.96,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":243.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":215.3,"x2":215.3,"y1":81.96,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":210.3}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":182.1,"x2":182.1,"y1":81.96,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":177.1}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.9,"x2":148.9,"y1":81.96,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":143.9}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":115.7,"x2":115.7,"y1":81.96,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":113.2}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":80}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":248.5,"y1":187.5,"y2":187.5}
{"baselineY":163.615,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":59.5}
{"fill":"#4472C4","h":15.077,"kind":"rect","w":66.4,"x":82.5,"y":146.038}
{"fill":"#4472C4","h":15.077,"kind":"rect","w":26.56,"x":82.5,"y":161.115}
{"baselineY":110.845,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":59.5}
{"fill":"#4472C4","h":15.077,"kind":"rect","w":132.8,"x":82.5,"y":93.268}
{"fill":"#4472C4","h":15.077,"kind":"rect","w":166,"x":82.5,"y":108.345}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-line
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.25,\"y\":140.617} .. {\"type\":\"close\"} #9be0415589e02728","h":9.333,"kind":"shape","w":9.333,"x":120.583,"y":140.617}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.25,"x2":210.75,"y1":145.284,"y2":103.068}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":210.75,\"y\":98.401} .. {\"type\":\"close\"} #aaede27475345792","h":9.333,"kind":"shape","w":9.333,"x":206.083,"y":98.401}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":120.583,"y":165.947}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.25,"x2":210.75,"y1":170.614,"y2":81.96}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":206.083,"y":77.293}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-pie
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, pie chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"fill":"#4472C4","geometryPath":"19 commands {\"type\":\"move\",\"x\":159.25,\"y\":148.68} .. {\"type\":\"close\"} #7c19b6a80843de75","h":128.376,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":128.376,"x":95.062,"y":84.492}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":159.25,\"y\":148.68} .. {\"type\":\"close\"} #883c4797ba49780f","h":128.376,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":128.376,"x":95.062,"y":84.492}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
# type-doughnut
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, doughnut chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":159.25,\"y\":84.492} .. {\"type\":\"close\"} #addecb5d63a93661","h":128.376,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":128.376,"x":95.062,"y":84.492}
{"fill":"#4472C4","geometryPath":"67 commands {\"type\":\"move\",\"x\":214.838,\"y\":180.774} .. {\"type\":\"close\"} #d738b8af64fc719b","h":128.376,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":128.376,"x":95.062,"y":84.492}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
# type-area
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.25,\"y\":145.284} .. {\"type\":\"close\"} #f4570777ba029ab8","h":105.54,"kind":"shape","w":171,"x":82.5,"y":81.96}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":125.25,"x2":210.75,"y1":145.284,"y2":103.068}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.25,\"y\":170.614} .. {\"type\":\"close\"} #3d2244d1a092fc63","h":105.54,"kind":"shape","w":171,"x":82.5,"y":81.96}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":125.25,"x2":210.75,"y1":170.614,"y2":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-scatter
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, scatter chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-radar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, radar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":138.98,"y2":158.38}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":129.281,"y2":168.079}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":119.581,"y2":177.779}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":109.882,"y2":187.478}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":100.182,"y2":197.178}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":148.68,"y2":100.182}
{"baselineY":95.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":111.33}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":148.68,"y2":197.178}
{"baselineY":202.027,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":111.33}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":127.33,"x2":127.33,"y1":129.281,"y2":187.478}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":127.33,\"y\":124.614} .. {\"type\":\"close\"} #37975d013d9b2d58","h":9.333,"kind":"shape","w":9.333,"x":122.663,"y":124.614}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":127.33,\"y\":182.811} .. {\"type\":\"close\"} #6fbd59d890eebb23","h":9.333,"kind":"shape","w":9.333,"x":122.663,"y":182.811}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":127.33,"x2":127.33,"y1":140.92,"y2":197.178}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":122.663,"y":136.254}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":122.663,"y":192.511}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-stock
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, stock chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.25,\"y\":140.617} .. {\"type\":\"close\"} #9be0415589e02728","h":9.333,"kind":"shape","w":9.333,"x":120.583,"y":140.617}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.25,"x2":210.75,"y1":145.284,"y2":103.068}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":210.75,\"y\":98.401} .. {\"type\":\"close\"} #aaede27475345792","h":9.333,"kind":"shape","w":9.333,"x":206.083,"y":98.401}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":120.583,"y":165.947}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.25,"x2":210.75,"y1":170.614,"y2":81.96}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":206.083,"y":77.293}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-bubble
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bubble chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-surface
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, surface chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"fill":"#70AD47","h":52.77,"kind":"rect","w":85.5,"x":82.5,"y":134.73}
{"fill":"#ED7D31","h":52.77,"kind":"rect","w":85.5,"x":168,"y":134.73}
{"baselineY":166.392,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":36,"x":44.5}
{"fill":"#4472C4","h":52.77,"kind":"rect","w":85.5,"x":82.5,"y":81.96}
{"fill":"#9E480E","h":52.77,"kind":"rect","w":85.5,"x":168,"y":81.96}
{"baselineY":113.622,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":36,"x":44.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":81.5,"x":84.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":81.5,"x":170}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-mystery
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":24.429,"x":100.821,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":24.429,"x":125.25,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":24.429,"x":186.321,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":24.429,"x":210.75,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":24.429,"x":100.821,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":24.429,"x":125.25,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":24.429,"x":186.321,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":24.429,"x":210.75,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# legend-left
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":139,"x2":295.9,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":116}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":139,"x2":295.9,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":116}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":139,"x2":295.9,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":116}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":139,"x2":295.9,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":116}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":139,"x2":295.9,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":121}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":139,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":121}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":139,"x2":139,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":139,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":173.225}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":22.414,"x":155.811,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":22.414,"x":178.225,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":251.675}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":22.414,"x":234.261,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":22.414,"x":256.675,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":71.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":71.5}
# legend-right
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":24.429,"x":100.821,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":24.429,"x":125.25,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":24.429,"x":186.321,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":24.429,"x":210.75,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# legend-top
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":107.96,"y2":107.96}
{"baselineY":110.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":123.868,"y2":123.868}
{"baselineY":126.368,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":139.776,"y2":139.776}
{"baselineY":142.276,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":155.684,"y2":155.684}
{"baselineY":158.184,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":171.592,"y2":171.592}
{"baselineY":174.092,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":107.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":130.85}
{"fill":"#4472C4","h":31.816,"kind":"rect","w":30.486,"x":105.364,"y":155.684}
{"fill":"#4472C4","h":12.726,"kind":"rect","w":30.486,"x":135.85,"y":174.774}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237.55}
{"fill":"#4472C4","h":63.632,"kind":"rect","w":30.486,"x":212.064,"y":123.868}
{"fill":"#4472C4","h":79.54,"kind":"rect","w":30.486,"x":242.55,"y":107.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":143.25,"y":81.71}
{"baselineY":86.76,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":151.75}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":183.25,"y":81.71}
{"baselineY":86.76,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":191.75}
# legend-bottom
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":97.868,"y2":97.868}
{"baselineY":100.368,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":113.776,"y2":113.776}
{"baselineY":116.276,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":129.684,"y2":129.684}
{"baselineY":132.184,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":145.592,"y2":145.592}
{"baselineY":148.092,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":161.5,"y2":161.5}
{"baselineY":164,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":161.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":295.9,"y1":161.5,"y2":161.5}
{"baselineY":182,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":130.85}
{"fill":"#4472C4","h":31.816,"kind":"rect","w":30.486,"x":105.364,"y":129.684}
{"fill":"#4472C4","h":12.726,"kind":"rect","w":30.486,"x":135.85,"y":148.774}
{"baselineY":182,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237.55}
{"fill":"#4472C4","h":63.632,"kind":"rect","w":30.486,"x":212.064,"y":97.868}
{"fill":"#4472C4","h":79.54,"kind":"rect","w":30.486,"x":242.55,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":143.25,"y":200.85}
{"baselineY":205.9,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":151.75}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":183.25,"y":200.85}
{"baselineY":205.9,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":191.75}
# legend-hidden
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":130.85}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":30.486,"x":105.364,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":30.486,"x":135.85,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237.55}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":30.486,"x":212.064,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":30.486,"x":242.55,"y":81.96}
# legend-default-position
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":24.429,"x":100.821,"y":145.284}
{"fill":"#4472C4","h":16.886,"kind":"rect","w":24.429,"x":125.25,"y":170.614}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":24.429,"x":186.321,"y":103.068}
{"fill":"#4472C4","h":105.54,"kind":"rect","w":24.429,"x":210.75,"y":81.96}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# legend-overflow
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 12 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":238.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":238.5,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":238.5,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":238.5,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":238.5,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":238.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":238.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":116.5}
{"fill":"#4472C4","h":1,"kind":"rect","w":5.778,"x":86.833,"y":187.5}
{"fill":"#ED7D31","h":5.336,"kind":"rect","w":5.778,"x":92.611,"y":182.164}
{"fill":"#A5A5A5","h":10.672,"kind":"rect","w":5.778,"x":98.389,"y":176.828}
{"fill":"#FFC000","h":16.008,"kind":"rect","w":5.778,"x":104.167,"y":171.492}
{"fill":"#5B9BD5","h":21.344,"kind":"rect","w":5.778,"x":109.944,"y":166.156}
{"fill":"#70AD47","h":26.68,"kind":"rect","w":5.778,"x":115.722,"y":160.82}
{"fill":"#264478","h":32.016,"kind":"rect","w":5.778,"x":121.5,"y":155.484}
{"fill":"#9E480E","h":37.352,"kind":"rect","w":5.778,"x":127.278,"y":150.148}
{"fill":"#4472C4","h":42.688,"kind":"rect","w":5.778,"x":133.056,"y":144.812}
{"fill":"#ED7D31","h":48.024,"kind":"rect","w":5.778,"x":138.833,"y":139.476}
{"fill":"#A5A5A5","h":53.36,"kind":"rect","w":5.778,"x":144.611,"y":134.14}
{"fill":"#FFC000","h":58.696,"kind":"rect","w":5.778,"x":150.389,"y":128.804}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":194.5}
{"fill":"#4472C4","h":1,"kind":"rect","w":5.778,"x":164.833,"y":187.5}
{"fill":"#ED7D31","h":10.672,"kind":"rect","w":5.778,"x":170.611,"y":176.828}
{"fill":"#A5A5A5","h":21.344,"kind":"rect","w":5.778,"x":176.389,"y":166.156}
{"fill":"#FFC000","h":32.016,"kind":"rect","w":5.778,"x":182.167,"y":155.484}
{"fill":"#5B9BD5","h":42.688,"kind":"rect","w":5.778,"x":187.944,"y":144.812}
{"fill":"#70AD47","h":53.36,"kind":"rect","w":5.778,"x":193.722,"y":134.14}
{"fill":"#264478","h":64.032,"kind":"rect","w":5.778,"x":199.5,"y":123.468}
{"fill":"#9E480E","h":74.704,"kind":"rect","w":5.778,"x":205.278,"y":112.796}
{"fill":"#4472C4","h":85.376,"kind":"rect","w":5.778,"x":211.056,"y":102.124}
{"fill":"#ED7D31","h":96.048,"kind":"rect","w":5.778,"x":216.833,"y":91.452}
{"fill":"#A5A5A5","h":106.72,"kind":"rect","w":5.778,"x":222.611,"y":80.78}
{"fill":"#FFC000","h":117.392,"kind":"rect","w":5.778,"x":228.389,"y":70.108}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":73.8}
{"baselineY":78.85,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 0","width":40,"x":257}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":89.1}
{"baselineY":94.15,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 1","width":40,"x":257}
{"fill":"#A5A5A5","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":104.4}
{"baselineY":109.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 2","width":40,"x":257}
{"fill":"#FFC000","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 3","width":40,"x":257}
{"fill":"#5B9BD5","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 4","width":40,"x":257}
{"fill":"#70AD47","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":150.3}
{"baselineY":155.35,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 5","width":40,"x":257}
{"fill":"#264478","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":165.6}
{"baselineY":170.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 6","width":40,"x":257}
{"fill":"#9E480E","h":5.3,"kind":"rect","w":5.3,"x":248.5,"y":180.9}
{"baselineY":185.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Series 7","width":40,"x":257}
# legend-overflow-pie
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 10 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #443ba287c3df8e7e","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#ED7D31","geometryPath":"5 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #d0ea92d1412b50f2","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#A5A5A5","geometryPath":"6 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #84dbca5287247bfe","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#FFC000","geometryPath":"7 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #69ae3128a03c59c6","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#5B9BD5","geometryPath":"8 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #95689ab3cf7a06ba","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#70AD47","geometryPath":"9 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #cbad294609788ffb","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#264478","geometryPath":"10 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #6f1e9ded91272153","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#9E480E","geometryPath":"10 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #24ff483d853a2877","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #62be5d089f349f5b","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#ED7D31","geometryPath":"12 commands {\"type\":\"move\",\"x\":161.75,\"y\":134.75} .. {\"type\":\"close\"} #aff6d6f725cab708","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":85.025,"y":58.025}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":73.8}
{"baselineY":78.85,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"a","width":5,"x":292}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":89.1}
{"baselineY":94.15,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"b","width":5,"x":292}
{"fill":"#A5A5A5","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":104.4}
{"baselineY":109.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"c","width":5,"x":292}
{"fill":"#FFC000","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"d","width":5,"x":292}
{"fill":"#5B9BD5","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"e","width":5,"x":292}
{"fill":"#70AD47","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":150.3}
{"baselineY":155.35,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"f","width":5,"x":292}
{"fill":"#264478","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":165.6}
{"baselineY":170.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"g","width":5,"x":292}
{"fill":"#9E480E","h":5.3,"kind":"rect","w":5.3,"x":283.5,"y":180.9}
{"baselineY":185.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"h","width":5,"x":292}
# combo-column-line
attrs {"ariaLabel":"Combo","blockId":42,"chart":{"label":"Combo, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Combo","width":32.5,"x":163.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":92.514,"y2":92.514}
{"baselineY":95.014,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":113.622,"y2":113.622}
{"baselineY":116.122,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":134.73,"y2":134.73}
{"baselineY":137.23,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":155.838,"y2":155.838}
{"baselineY":158.338,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":176.946,"y2":176.946}
{"baselineY":179.446,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":117.75}
{"fill":"#4472C4","h":52.77,"kind":"rect","w":32.2,"x":106.65,"y":134.73}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":198.25}
{"fill":"#4472C4","h":94.986,"kind":"rect","w":32.2,"x":187.15,"y":92.514}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":93.687,"y2":93.687}
{"baselineY":96.187,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":105.413,"y2":105.413}
{"baselineY":107.913,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":117.14,"y2":117.14}
{"baselineY":119.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":128.867,"y2":128.867}
{"baselineY":131.367,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":140.593,"y2":140.593}
{"baselineY":143.093,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":152.32,"y2":152.32}
{"baselineY":154.82,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":164.047,"y2":164.047}
{"baselineY":166.547,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":175.773,"y2":175.773}
{"baselineY":178.273,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":117.75}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":198.25}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":122.75,\"y\":135.927} .. {\"type\":\"close\"} #3a5e30bd94a7310c","h":9.333,"kind":"shape","w":9.333,"x":118.083,"y":135.927}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":122.75,"x2":203.25,"y1":140.593,"y2":93.687}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":203.25,\"y\":89.02} .. {\"type\":\"close\"} #a0400002048959fd","h":9.333,"kind":"shape","w":9.333,"x":198.583,"y":89.02}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":253.5,"y":138.38}
{"baselineY":143.43,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Revenue","width":35,"x":262}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":253.5,"y":153.68}
{"baselineY":158.73,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Trend","width":35,"x":262}
# combo-with-pie-group
attrs {"blockId":42,"chart":{"label":"Untitled chart, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"39 commands {\"type\":\"move\",\"x\":200.75,\"y\":134.75} .. {\"type\":\"close\"} #4d48abe9cdaff621","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":124.025,"y":58.025}
{"fill":"#4472C4","geometryPath":"15 commands {\"type\":\"move\",\"x\":200.75,\"y\":134.75} .. {\"type\":\"close\"} #c701c5cb8384113a","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":124.025,"y":58.025}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":293.4,"x2":293.4,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":290.9}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":269.2,"x2":269.2,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":266.7}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":245,"x2":245,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":242.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":220.8,"x2":220.8,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":218.3}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":196.6,"x2":196.6,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":194.1}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":172.4,"x2":172.4,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":169.9}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.2,"x2":148.2,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":145.7}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":124,"x2":124,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":121.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":124,"x2":124,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":124,"x2":293.4,"y1":187.5,"y2":187.5}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":101}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":48.4,"x":124,"y":140.81}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":101}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":145.2,"x":124,"y":74.11}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":71.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":71.5}
# points-markers-labels
attrs {"ariaLabel":"Points","blockId":42,"chart":{"label":"Points, line chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Points","width":39,"x":160.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":95.152,"y2":95.152}
{"baselineY":97.652,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":108.345,"y2":108.345}
{"baselineY":110.845,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":121.538,"y2":121.538}
{"baselineY":124.038,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":134.73,"y2":134.73}
{"baselineY":137.23,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":147.923,"y2":147.923}
{"baselineY":150.423,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":161.115,"y2":161.115}
{"baselineY":163.615,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":174.308,"y2":174.308}
{"baselineY":176.808,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":77.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":77.5,"x2":77.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":77.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":101.833}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":160.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":10,"x":219.167}
{"fill":"#FF0000","geometryPath":"5 commands {\"type\":\"move\",\"x\":106.833,\"y\":85.819} .. {\"type\":\"close\"} #38c3d468dce0f109","h":18.667,"kind":"shape","w":18.667,"x":97.5,"y":85.819}
{"baselineY":98.652,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"peak","width":48,"x":119.167}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":106.833,"x2":165.5,"y1":95.152,"y2":161.115}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":165.5,\"y\":155.115} .. {\"type\":\"close\"} #189c045ed1abbc31","h":12,"kind":"shape","w":12,"x":159.5,"y":155.115}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":165.5,"x2":224.167,"y1":161.115,"y2":147.923}
{"fill":"#00FF00","geometryPath":"5 commands {\"type\":\"move\",\"x\":224.167,\"y\":141.923} .. {\"type\":\"close\"} #c04b3ad0628f67e7","h":12,"kind":"shape","w":12,"x":218.167,"y":141.923}
{"baselineY":151.423,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"end","width":48,"x":233.167}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":146.03}
{"baselineY":151.08,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# points-without-indexes
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#123456","geometryPath":"39 commands {\"type\":\"move\",\"x\":172.813,\"y\":148.313} .. {\"type\":\"close\"} #00a7a2d07a86461c","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":96.088,"y":71.588}
{"fill":"#123456","geometryPath":"15 commands {\"type\":\"move\",\"x\":145.687,\"y\":121.187} .. {\"type\":\"close\"} #0704851f8ec0a336","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":68.962,"y":44.462}
{"fill":"#123456","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#123456","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
# negative-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":73.157,"y2":73.157}
{"baselineY":75.657,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":92.214,"y2":92.214}
{"baselineY":94.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":111.271,"y2":111.271}
{"baselineY":113.771,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":130.329,"y2":130.329}
{"baselineY":132.829,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":149.386,"y2":149.386}
{"baselineY":151.886,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":168.443,"y2":168.443}
{"baselineY":170.943,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-30","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40","width":15,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124}
{"fill":"#4472C4","h":19.057,"kind":"rect","w":23.714,"x":105.286,"y":111.271}
{"fill":"#4472C4","h":7.623,"kind":"rect","w":23.714,"x":129,"y":111.271}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":207}
{"fill":"#4472C4","h":38.114,"kind":"rect","w":23.714,"x":188.286,"y":73.157}
{"fill":"#4472C4","h":57.171,"kind":"rect","w":23.714,"x":212,"y":111.271}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# negative-values-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":248.5,"x2":248.5,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":10,"x":243.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":215.3,"x2":215.3,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":210.3}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":182.1,"x2":182.1,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":177.1}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.9,"x2":148.9,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":146.4}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":115.7,"x2":115.7,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":15,"x":108.2}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20","width":15,"x":75}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":248.5,"y1":187.5,"y2":187.5}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":59.5}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":33.2,"x":115.7,"y":140.81}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":59.5}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":66.4,"x":148.9,"y":74.11}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# inverted-axis-bounds
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"11","width":10,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.8","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.6","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.4","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.2","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":92.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.75}
{"fill":"#4472C4","h":1,"kind":"rect","w":32.2,"x":116.65,"y":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.25}
{"fill":"#4472C4","h":1,"kind":"rect","w":32.2,"x":197.15,"y":187.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# long-text
attrs {"ariaLabel":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT","blockId":42,"chart":{"label":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT","width":780,"x":-210}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":203.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":203.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":203.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":203.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":203.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":203.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":203.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC","width":600,"x":-183.5}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":23.2,"x":104.9,"y":145.284}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":169.5}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":23.2,"x":162.9,"y":103.068}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":213.5,"y":103.33}
{"baselineY":108.38,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":120.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":132.78,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":144.98,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":157.18,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":169.38,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":181.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":193.78,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
# no-title-no-legend
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":134.6}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":238.8}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":139.6,\"y\":129.473} .. {\"type\":\"close\"} #78ad1d4a0e01ec62","h":9.333,"kind":"shape","w":9.333,"x":134.933,"y":129.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":139.6,"x2":243.8,"y1":134.14,"y2":80.78}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":243.8,\"y\":76.113} .. {\"type\":\"close\"} #b15045fc0d7d7b74","h":9.333,"kind":"shape","w":9.333,"x":239.133,"y":76.113}
# empty-series
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 0 series, 0 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
# series-without-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 0 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
# described-and-decorative
attrs {"ariaDescription":"quarterly revenue","ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 1 series, 2 categories"},"decorative":true,"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":103.068,"y2":103.068}
{"baselineY":105.568,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":124.176,"y2":124.176}
{"baselineY":126.676,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":145.284,"y2":145.284}
{"baselineY":147.784,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":166.392,"y2":166.392}
{"baselineY":168.892,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":81.96,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124}
{"fill":"#4472C4","h":42.216,"kind":"rect","w":33.2,"x":112.4,"y":145.284}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":207}
{"fill":"#4472C4","h":84.432,"kind":"rect","w":33.2,"x":195.4,"y":103.068}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":146.03}
{"baselineY":151.08,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# stacked-column
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":73.157,"y2":73.157}
{"baselineY":75.657,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":92.214,"y2":92.214}
{"baselineY":94.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":111.271,"y2":111.271}
{"baselineY":113.771,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":130.329,"y2":130.329}
{"baselineY":132.829,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":149.386,"y2":149.386}
{"baselineY":151.886,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":168.443,"y2":168.443}
{"baselineY":170.943,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":15,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124}
{"fill":"#4472C4","h":38.114,"kind":"rect","w":33.2,"x":112.4,"y":111.271}
{"fill":"#4472C4","h":19.057,"kind":"rect","w":33.2,"x":112.4,"y":92.214}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":207}
{"fill":"#4472C4","h":76.229,"kind":"rect","w":33.2,"x":195.4,"y":73.157}
{"fill":"#4472C4","h":30.491,"kind":"rect","w":33.2,"x":195.4,"y":149.386}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# stacked-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":248.5,"x2":248.5,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":243.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":224.786,"x2":224.786,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":219.786}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":201.071,"x2":201.071,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":196.071}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.357,"x2":177.357,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":172.357}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":153.643,"x2":153.643,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":151.143}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":129.929,"x2":129.929,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":127.429}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":106.214,"x2":106.214,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":10,"x":101.214}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":15,"x":75}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":248.5,"y1":187.5,"y2":187.5}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":59.5}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":47.429,"x":129.929,"y":140.81}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":23.714,"x":177.357,"y":140.81}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":59.5}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":94.857,"x":129.929,"y":74.11}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":37.943,"x":91.986,"y":74.11}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# stacked-line
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":73.157,"y2":73.157}
{"baselineY":75.657,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":92.214,"y2":92.214}
{"baselineY":94.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":111.271,"y2":111.271}
{"baselineY":113.771,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":130.329,"y2":130.329}
{"baselineY":132.829,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":149.386,"y2":149.386}
{"baselineY":151.886,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":168.443,"y2":168.443}
{"baselineY":170.943,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":15,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":207}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":129,\"y\":106.605} .. {\"type\":\"close\"} #b8d1b679160bc749","h":9.333,"kind":"shape","w":9.333,"x":124.333,"y":106.605}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":129,"x2":212,"y1":111.271,"y2":73.157}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":212,\"y\":68.49} .. {\"type\":\"close\"} #ceebee5cad80b931","h":9.333,"kind":"shape","w":9.333,"x":207.333,"y":68.49}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":124.333,"y":87.548}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":129,"x2":212,"y1":92.214,"y2":179.877}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":207.333,"y":175.21}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# stacked-area
attrs {"blockId":42,"chart":{"label":"Untitled chart, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":73.157,"y2":73.157}
{"baselineY":75.657,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":92.214,"y2":92.214}
{"baselineY":94.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":111.271,"y2":111.271}
{"baselineY":113.771,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":130.329,"y2":130.329}
{"baselineY":132.829,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":149.386,"y2":149.386}
{"baselineY":151.886,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":168.443,"y2":168.443}
{"baselineY":170.943,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":10,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":15,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":207}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":129,\"y\":111.271} .. {\"type\":\"close\"} #4ecb5da94530632a","h":133.4,"kind":"shape","w":166,"x":87.5,"y":54.1}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":129,"x2":212,"y1":111.271,"y2":73.157}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":129,\"y\":92.214} .. {\"type\":\"close\"} #43b3b502bacdafb5","h":133.4,"kind":"shape","w":166,"x":87.5,"y":54.1}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":129,"x2":212,"y1":92.214,"y2":179.877}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# percentStacked-column
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":70.775,"y2":70.775}
{"baselineY":73.275,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":87.45,"y2":87.45}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":104.125,"y2":104.125}
{"baselineY":106.625,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":120.8,"y2":120.8}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":137.475,"y2":137.475}
{"baselineY":139.975,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":154.15,"y2":154.15}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":10,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":170.825,"y2":170.825}
{"baselineY":173.325,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":20,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":92.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.75}
{"fill":"#4472C4","h":55.583,"kind":"rect","w":32.2,"x":116.65,"y":98.567}
{"fill":"#4472C4","h":27.792,"kind":"rect","w":32.2,"x":116.65,"y":70.775}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.25}
{"fill":"#4472C4","h":59.554,"kind":"rect","w":32.2,"x":197.15,"y":94.596}
{"fill":"#4472C4","h":23.821,"kind":"rect","w":32.2,"x":197.15,"y":154.15}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# percentStacked-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":243.5,"x2":243.5,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":20,"x":233.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":223.375,"x2":223.375,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":20,"x":213.375}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":203.25,"x2":203.25,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":15,"x":195.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.125,"x2":183.125,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":15,"x":175.625}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":163,"x2":163,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":15,"x":155.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.875,"x2":142.875,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":15,"x":135.375}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":122.75,"x2":122.75,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":10,"x":117.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.625,"x2":102.625,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":20,"x":92.625}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":20,"x":72.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":59.5}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":67.083,"x":122.75,"y":140.81}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":33.542,"x":189.833,"y":140.81}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":59.5}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":71.875,"x":122.75,"y":74.11}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":28.75,"x":94,"y":74.11}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# percentStacked-line
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":70.775,"y2":70.775}
{"baselineY":73.275,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":87.45,"y2":87.45}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":104.125,"y2":104.125}
{"baselineY":106.625,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":120.8,"y2":120.8}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":137.475,"y2":137.475}
{"baselineY":139.975,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":154.15,"y2":154.15}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":10,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":170.825,"y2":170.825}
{"baselineY":173.325,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":20,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":92.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.75}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.25}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.75,\"y\":93.9} .. {\"type\":\"close\"} #5be86b9a6a3fe39a","h":9.333,"kind":"shape","w":9.333,"x":128.083,"y":93.9}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.75,"x2":213.25,"y1":98.567,"y2":94.596}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":213.25,\"y\":89.93} .. {\"type\":\"close\"} #9316ba5a8ceed28f","h":9.333,"kind":"shape","w":9.333,"x":208.583,"y":89.93}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":128.083,"y":66.108}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.75,"x2":213.25,"y1":70.775,"y2":177.971}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":208.583,"y":173.305}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# percentStacked-area
attrs {"blockId":42,"chart":{"label":"Untitled chart, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":70.775,"y2":70.775}
{"baselineY":73.275,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":87.45,"y2":87.45}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":104.125,"y2":104.125}
{"baselineY":106.625,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":120.8,"y2":120.8}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":137.475,"y2":137.475}
{"baselineY":139.975,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":15,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":154.15,"y2":154.15}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":10,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":170.825,"y2":170.825}
{"baselineY":173.325,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":20,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":20,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":92.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.75}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.25}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.75,\"y\":98.567} .. {\"type\":\"close\"} #ca8b5373252d1e54","h":133.4,"kind":"shape","w":161,"x":92.5,"y":54.1}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.75,"x2":213.25,"y1":98.567,"y2":94.596}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.75,\"y\":70.775} .. {\"type\":\"close\"} #ad2f097ab9b0eea9","h":133.4,"kind":"shape","w":161,"x":92.5,"y":54.1}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.75,"x2":213.25,"y1":70.775,"y2":177.971}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# bar-gap-and-overlap
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"35","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":73.157,"y2":73.157}
{"baselineY":75.657,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":92.214,"y2":92.214}
{"baselineY":94.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":111.271,"y2":111.271}
{"baselineY":113.771,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":130.329,"y2":130.329}
{"baselineY":132.829,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":149.386,"y2":149.386}
{"baselineY":151.886,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":168.443,"y2":168.443}
{"baselineY":170.943,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":120.25}
{"fill":"#4472C4","h":38.114,"kind":"rect","w":32.885,"x":89.077,"y":149.386}
{"fill":"#4472C4","h":15.246,"kind":"rect","w":32.885,"x":128.538,"y":172.254}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":205.75}
{"fill":"#4472C4","h":76.229,"kind":"rect","w":32.885,"x":174.577,"y":111.271}
{"fill":"#4472C4","h":114.343,"kind":"rect","w":32.885,"x":214.038,"y":73.157}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# scatter-xy
attrs {"blockId":42,"chart":{"label":"Untitled chart, scatter chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":67.44,"y2":67.44}
{"baselineY":69.94,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":94.12,"y2":94.12}
{"baselineY":96.62,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":120.8,"y2":120.8}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":147.48,"y2":147.48}
{"baselineY":149.98,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":174.16,"y2":174.16}
{"baselineY":176.66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":268.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":268.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":80}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":103.25}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":126.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":149.75}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":173}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":196.25}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":219.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":5,"x":242.75}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":5,"x":266}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":82.5,\"y\":142.813} .. {\"type\":\"close\"} #d330c4d5b5669832","h":9.333,"kind":"shape","w":9.333,"x":77.833,"y":142.813}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":82.5,"x2":175.5,"y1":147.48,"y2":67.44}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":175.5,\"y\":62.773} .. {\"type\":\"close\"} #e71c4e2c25517945","h":9.333,"kind":"shape","w":9.333,"x":170.833,"y":62.773}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":175.5,"x2":268.5,"y1":67.44,"y2":134.14}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":268.5,\"y\":129.473} .. {\"type\":\"close\"} #0c644f2ec1dcecea","h":9.333,"kind":"shape","w":9.333,"x":263.833,"y":129.473}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"XY","width":10,"x":287}
# bubble-sizes
attrs {"blockId":42,"chart":{"label":"Untitled chart, bubble chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":67.44,"y2":67.44}
{"baselineY":69.94,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":94.12,"y2":94.12}
{"baselineY":96.62,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":120.8,"y2":120.8}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":147.48,"y2":147.48}
{"baselineY":149.98,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":174.16,"y2":174.16}
{"baselineY":176.66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":243.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":80}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":95.125}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":120.25}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":135.375}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":5,"x":160.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3.5","width":15,"x":175.625}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":5,"x":200.75}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4.5","width":15,"x":215.875}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":241}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":92.505,\"y\":147.48} .. {\"type\":\"close\"} #51fe27b65239d655","h":20.01,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20.01,"x":72.495,"y":137.475}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":263.51,\"y\":67.44} .. {\"type\":\"close\"} #2e220d5e951a297b","h":40.02,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":40.02,"x":223.49,"y":47.43}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":253.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Bubbles","width":35,"x":262}
# radar-standard
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":131.227,"y1":128.953,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":131.227,"x2":125.43,"y1":134.75,"y2":140.547}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":119.633,"y1":140.547,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":119.633,"x2":125.43,"y1":134.75,"y2":128.953}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":137.024,"y1":123.156,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":137.024,"x2":125.43,"y1":134.75,"y2":146.344}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":113.836,"y1":146.344,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":113.836,"x2":125.43,"y1":134.75,"y2":123.156}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":142.821,"y1":117.359,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.821,"x2":125.43,"y1":134.75,"y2":152.141}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":108.039,"y1":152.141,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":108.039,"x2":125.43,"y1":134.75,"y2":117.359}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":148.618,"y1":111.562,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.618,"x2":125.43,"y1":134.75,"y2":157.938}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":102.242,"y1":157.938,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.242,"x2":125.43,"y1":134.75,"y2":111.562}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":154.415,"y1":105.765,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154.415,"x2":125.43,"y1":134.75,"y2":163.735}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":96.445,"y1":163.735,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":96.445,"x2":125.43,"y1":134.75,"y2":105.765}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":160.212,"y1":99.968,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.212,"x2":125.43,"y1":134.75,"y2":169.532}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":90.648,"y1":169.532,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":90.648,"x2":125.43,"y1":134.75,"y2":99.968}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":166.009,"y1":94.171,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":166.009,"x2":125.43,"y1":134.75,"y2":175.329}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":84.851,"y1":175.329,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":84.851,"x2":125.43,"y1":134.75,"y2":94.171}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":171.806,"y1":88.374,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.806,"x2":125.43,"y1":134.75,"y2":181.126}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":79.054,"y1":181.126,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.054,"x2":125.43,"y1":134.75,"y2":88.374}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":177.603,"y1":82.577,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.603,"x2":125.43,"y1":134.75,"y2":186.923}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":73.257,"y1":186.923,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":73.257,"x2":125.43,"y1":134.75,"y2":82.577}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.4,"y1":76.78,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.4,"x2":125.43,"y1":134.75,"y2":192.72}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.46,"y1":192.72,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":67.46,"x2":125.43,"y1":134.75,"y2":76.78}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":134.75,"y2":76.78}
{"baselineY":70.983,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.4,"y1":134.75,"y2":134.75}
{"baselineY":134.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":173.197}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":134.75,"y2":192.72}
{"baselineY":198.517,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.46,"y1":134.75,"y2":134.75}
{"baselineY":134.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":45.663}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":177.603,"y1":117.359,"y2":134.75}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":177.603,"x2":125.43,"y1":134.75,"y2":157.938}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":90.648,"y1":157.938,"y2":134.75}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":90.648,"x2":125.43,"y1":134.75,"y2":117.359}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":258.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":30,"x":267}
# radar-marker
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":131.227,"y1":128.953,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":131.227,"x2":125.43,"y1":134.75,"y2":140.547}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":119.633,"y1":140.547,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":119.633,"x2":125.43,"y1":134.75,"y2":128.953}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":137.024,"y1":123.156,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":137.024,"x2":125.43,"y1":134.75,"y2":146.344}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":113.836,"y1":146.344,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":113.836,"x2":125.43,"y1":134.75,"y2":123.156}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":142.821,"y1":117.359,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.821,"x2":125.43,"y1":134.75,"y2":152.141}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":108.039,"y1":152.141,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":108.039,"x2":125.43,"y1":134.75,"y2":117.359}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":148.618,"y1":111.562,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.618,"x2":125.43,"y1":134.75,"y2":157.938}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":102.242,"y1":157.938,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.242,"x2":125.43,"y1":134.75,"y2":111.562}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":154.415,"y1":105.765,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154.415,"x2":125.43,"y1":134.75,"y2":163.735}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":96.445,"y1":163.735,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":96.445,"x2":125.43,"y1":134.75,"y2":105.765}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":160.212,"y1":99.968,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.212,"x2":125.43,"y1":134.75,"y2":169.532}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":90.648,"y1":169.532,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":90.648,"x2":125.43,"y1":134.75,"y2":99.968}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":166.009,"y1":94.171,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":166.009,"x2":125.43,"y1":134.75,"y2":175.329}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":84.851,"y1":175.329,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":84.851,"x2":125.43,"y1":134.75,"y2":94.171}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":171.806,"y1":88.374,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.806,"x2":125.43,"y1":134.75,"y2":181.126}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":79.054,"y1":181.126,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.054,"x2":125.43,"y1":134.75,"y2":88.374}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":177.603,"y1":82.577,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.603,"x2":125.43,"y1":134.75,"y2":186.923}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":73.257,"y1":186.923,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":73.257,"x2":125.43,"y1":134.75,"y2":82.577}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.4,"y1":76.78,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.4,"x2":125.43,"y1":134.75,"y2":192.72}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.46,"y1":192.72,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":67.46,"x2":125.43,"y1":134.75,"y2":76.78}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":134.75,"y2":76.78}
{"baselineY":70.983,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.4,"y1":134.75,"y2":134.75}
{"baselineY":134.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":173.197}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":134.75,"y2":192.72}
{"baselineY":198.517,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.46,"y1":134.75,"y2":134.75}
{"baselineY":134.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":45.663}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":177.603,"y1":117.359,"y2":134.75}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":177.603,"x2":125.43,"y1":134.75,"y2":157.938}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":90.648,"y1":157.938,"y2":134.75}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":90.648,"x2":125.43,"y1":134.75,"y2":117.359}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.43,\"y\":112.692} .. {\"type\":\"close\"} #f600e406eb865847","h":9.333,"kind":"shape","w":9.333,"x":120.763,"y":112.692}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":177.603,\"y\":130.083} .. {\"type\":\"close\"} #574c209e7f367178","h":9.333,"kind":"shape","w":9.333,"x":172.936,"y":130.083}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.43,\"y\":153.271} .. {\"type\":\"close\"} #87c9dd513dbfc772","h":9.333,"kind":"shape","w":9.333,"x":120.763,"y":153.271}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":90.648,\"y\":130.083} .. {\"type\":\"close\"} #61d35d0f69c6b616","h":9.333,"kind":"shape","w":9.333,"x":85.981,"y":130.083}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":258.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":30,"x":267}
# radar-filled
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":131.227,"y1":128.953,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":131.227,"x2":125.43,"y1":134.75,"y2":140.547}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":119.633,"y1":140.547,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":119.633,"x2":125.43,"y1":134.75,"y2":128.953}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":137.024,"y1":123.156,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":137.024,"x2":125.43,"y1":134.75,"y2":146.344}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":113.836,"y1":146.344,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":113.836,"x2":125.43,"y1":134.75,"y2":123.156}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":142.821,"y1":117.359,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.821,"x2":125.43,"y1":134.75,"y2":152.141}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":108.039,"y1":152.141,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":108.039,"x2":125.43,"y1":134.75,"y2":117.359}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":148.618,"y1":111.562,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.618,"x2":125.43,"y1":134.75,"y2":157.938}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":102.242,"y1":157.938,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.242,"x2":125.43,"y1":134.75,"y2":111.562}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":154.415,"y1":105.765,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154.415,"x2":125.43,"y1":134.75,"y2":163.735}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":96.445,"y1":163.735,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":96.445,"x2":125.43,"y1":134.75,"y2":105.765}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":160.212,"y1":99.968,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.212,"x2":125.43,"y1":134.75,"y2":169.532}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":90.648,"y1":169.532,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":90.648,"x2":125.43,"y1":134.75,"y2":99.968}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":166.009,"y1":94.171,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":166.009,"x2":125.43,"y1":134.75,"y2":175.329}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":84.851,"y1":175.329,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":84.851,"x2":125.43,"y1":134.75,"y2":94.171}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":171.806,"y1":88.374,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.806,"x2":125.43,"y1":134.75,"y2":181.126}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":79.054,"y1":181.126,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.054,"x2":125.43,"y1":134.75,"y2":88.374}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":177.603,"y1":82.577,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.603,"x2":125.43,"y1":134.75,"y2":186.923}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":73.257,"y1":186.923,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":73.257,"x2":125.43,"y1":134.75,"y2":82.577}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.4,"y1":76.78,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.4,"x2":125.43,"y1":134.75,"y2":192.72}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.46,"y1":192.72,"y2":134.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":67.46,"x2":125.43,"y1":134.75,"y2":76.78}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":134.75,"y2":76.78}
{"baselineY":70.983,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.4,"y1":134.75,"y2":134.75}
{"baselineY":134.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":173.197}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":134.75,"y2":192.72}
{"baselineY":198.517,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.46,"y1":134.75,"y2":134.75}
{"baselineY":134.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":45.663}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.43,\"y\":117.359} .. {\"type\":\"close\"} #fc15d6d2d6766e7e","h":151.8,"kind":"shape","w":184.4,"x":64.1,"y":54.1}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":258.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":30,"x":267}
# stock-ohlc
attrs {"blockId":42,"chart":{"label":"Untitled chart, stock chart, 4 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":253.5,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D1","width":81.5,"x":84.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":125.25,"x2":125.25,"y1":80.78,"y2":160.82}
{"fill":"#FFFFFF","h":42.688,"kind":"rect","w":24,"x":113.25,"y":91.452}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":113.25,"x2":113.25,"y1":91.452,"y2":134.14}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D2","width":81.5,"x":170}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":210.75,"x2":210.75,"y1":70.108,"y2":155.484}
{"fill":"#666666","h":21.344,"kind":"rect","w":24,"x":198.75,"y":123.468}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":104.4}
{"baselineY":109.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Open","width":25,"x":272}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"High","width":25,"x":272}
{"fill":"#A5A5A5","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Low","width":25,"x":272}
{"fill":"#FFC000","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":150.3}
{"baselineY":155.35,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Close","width":25,"x":272}
# surface-contour
attrs {"blockId":42,"chart":{"label":"Untitled chart, surface chart, 2 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":268.5,"y1":187.5,"y2":187.5}
{"fill":"#4472C4","h":66.7,"kind":"rect","w":62,"x":82.5,"y":120.8}
{"fill":"#A9D18E","h":66.7,"kind":"rect","w":62,"x":144.5,"y":120.8}
{"fill":"#ED7D31","h":66.7,"kind":"rect","w":62,"x":206.5,"y":120.8}
{"baselineY":160.82,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R1","width":36,"x":44.5}
{"fill":"#ED7D31","h":66.7,"kind":"rect","w":62,"x":82.5,"y":54.1}
{"fill":"#4472C4","h":66.7,"kind":"rect","w":62,"x":144.5,"y":54.1}
{"fill":"#A9D18E","h":66.7,"kind":"rect","w":62,"x":206.5,"y":54.1}
{"baselineY":94.12,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R2","width":36,"x":44.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":58,"x":84.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":58,"x":146.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":58,"x":208.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R1","width":10,"x":287}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R2","width":10,"x":287}
# doughnut-hole-and-rotation
attrs {"blockId":42,"chart":{"label":"Untitled chart, doughnut chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"51 commands {\"type\":\"move\",\"x\":235.975,\"y\":134.75} .. {\"type\":\"close\"} #c1135ac2b6ee1523","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":82.525,"y":58.025}
{"fill":"#ED7D31","geometryPath":"19 commands {\"type\":\"move\",\"x\":62.591,\"y\":123.241} .. {\"type\":\"close\"} #915a7090f9f3e2bb","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":62.591,"y":46.516}
{"fill":"#A5A5A5","geometryPath":"35 commands {\"type\":\"move\",\"x\":120.887,\"y\":68.304} .. {\"type\":\"close\"} #4b9e65e2e13d273b","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":82.525,"y":58.025}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":112.05}
{"baselineY":117.1,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
{"fill":"#A5A5A5","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":142.65}
{"baselineY":147.7,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":10,"x":287}
# secondary-value-axis
attrs {"blockId":42,"chart":{"label":"Untitled chart, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":257.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":257.9,"y1":87.45,"y2":87.45}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":257.9,"y1":120.8,"y2":120.8}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":257.9,"y1":154.15,"y2":154.15}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":257.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":257.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":121.35}
{"fill":"#4472C4","h":66.7,"kind":"rect","w":35.08,"x":108.81,"y":120.8}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":209.05}
{"fill":"#4472C4","h":100.05,"kind":"rect","w":35.08,"x":196.51,"y":87.45}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100","width":15,"x":270.9}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80","width":10,"x":270.9}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60","width":10,"x":270.9}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40","width":10,"x":270.9}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":270.9}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":270.9}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":257.9,"x2":257.9,"y1":54.1,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":121.35}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":209.05}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":126.35,\"y\":129.473} .. {\"type\":\"close\"} #8b933efbdb333537","h":9.333,"kind":"shape","w":9.333,"x":121.683,"y":129.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":126.35,"x2":214.05,"y1":134.14,"y2":80.78}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":214.05,\"y\":76.113} .. {\"type\":\"close\"} #f920669d90e30689","h":9.333,"kind":"shape","w":9.333,"x":209.383,"y":76.113}
# log-scale-and-ticks
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":98.567,"y2":98.567}
{"baselineY":101.067,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100","width":15,"x":59.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":80.5,"x2":87.5,"y1":98.567,"y2":98.567}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":143.033,"y2":143.033}
{"baselineY":145.533,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":80.5,"x2":87.5,"y1":143.033,"y2":143.033}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":80.5,"x2":87.5,"y1":187.5,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":134.6}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":238.8}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":139.6,\"y\":182.833} .. {\"type\":\"close\"} #b07037d4fdb81a34","h":9.333,"kind":"shape","w":9.333,"x":134.933,"y":182.833}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":139.6,"x2":243.8,"y1":187.5,"y2":67.486}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":243.8,\"y\":62.819} .. {\"type\":\"close\"} #944d67a4d3f2dcfd","h":9.333,"kind":"shape","w":9.333,"x":239.133,"y":62.819}
# reversed-axes
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"baselineY":156.65,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"baselineY":123.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"baselineY":89.95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":237.55}
{"fill":"#4472C4","h":26.68,"kind":"rect","w":42.68,"x":221.21,"y":54.1}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":130.85}
{"fill":"#4472C4","h":106.72,"kind":"rect","w":42.68,"x":114.51,"y":54.1}
# marker-circle
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":121.267,\"y\":98.14} .. {\"type\":\"close\"} #adb5c60f86f6cff0","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":175.467,\"y\":68.78} .. {\"type\":\"close\"} #61297b431cde303c","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-diamond
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":114.6,\"y\":91.473} .. {\"type\":\"close\"} #f23c91628664447c","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":168.8,\"y\":62.113} .. {\"type\":\"close\"} #35894046dc25837e","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-triangle
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"4 commands {\"type\":\"move\",\"x\":114.6,\"y\":91.473} .. {\"type\":\"close\"} #55b3ce56c62efbe3","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"4 commands {\"type\":\"move\",\"x\":168.8,\"y\":62.113} .. {\"type\":\"close\"} #1180ec513e5945d5","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-square
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","h":13.333,"kind":"rect","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","h":13.333,"kind":"rect","w":13.333,"x":162.133,"y":62.113}
# marker-star
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":114.6,\"y\":91.473} .. {\"type\":\"close\"} #0e1a6612f117d3cd","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":168.8,\"y\":62.113} .. {\"type\":\"close\"} #93eb60f105a01df4","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-plus
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":112.6,\"y\":91.473} .. {\"type\":\"close\"} #cbd393337e67fabc","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":166.8,\"y\":62.113} .. {\"type\":\"close\"} #170d5aea98205886","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-dash
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":107.933,\"y\":96.473} .. {\"type\":\"close\"} #aa23be5f85b988b0","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":162.133,\"y\":67.113} .. {\"type\":\"close\"} #775c5bbb26032bb4","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-dot
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":117.933,\"y\":98.14} .. {\"type\":\"close\"} #ef4b34bb3c786598","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":172.133,\"y\":68.78} .. {\"type\":\"close\"} #401d13f8849bb280","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-x
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":117.9,\"y\":92.012} .. {\"type\":\"close\"} #1c6176f44448297e","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":172.1,\"y\":62.652} .. {\"type\":\"close\"} #4e1d32633e8a4db6","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-auto
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":114.6,\"y\":91.473} .. {\"type\":\"close\"} #f23c91628664447c","h":13.333,"kind":"shape","w":13.333,"x":107.933,"y":91.473}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":168.8,\"y\":62.113} .. {\"type\":\"close\"} #35894046dc25837e","h":13.333,"kind":"shape","w":13.333,"x":162.133,"y":62.113}
# marker-none
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":68.78,"y2":68.78}
{"baselineY":71.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":83.46,"y2":83.46}
{"baselineY":85.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":98.14,"y2":98.14}
{"baselineY":100.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":5,"x":69.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":112.82,"y2":112.82}
{"baselineY":115.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":15,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":69.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":87.5,"y1":54.1,"y2":127.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":87.5,"x2":195.9,"y1":127.5,"y2":127.5}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":109.6}
{"baselineY":148,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":163.8}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":114.6,"x2":168.8,"y1":98.14,"y2":68.78}
# data-labels-composed
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":54.1,"y2":54.1}
{"baselineY":56.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":80.78,"y2":80.78}
{"baselineY":83.28,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":107.46,"y2":107.46}
{"baselineY":109.96,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":134.14,"y2":134.14}
{"baselineY":136.64,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":160.82,"y2":160.82}
{"baselineY":163.32,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":190,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":54.1,"y2":187.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":295.9,"y1":187.5,"y2":187.5}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":130.85}
{"fill":"#4472C4","h":53.36,"kind":"rect","w":42.68,"x":114.51,"y":134.14}
{"baselineY":137.14,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North / Q1 / 10.0","width":85,"x":93.35}
{"baselineY":208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237.55}
{"fill":"#4472C4","h":106.72,"kind":"rect","w":42.68,"x":221.21,"y":80.78}
{"baselineY":83.78,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North / Q2 / 20.0","width":85,"x":200.05}
# data-labels-percent-and-key
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"39 commands {\"type\":\"move\",\"x\":180,\"y\":134.75} .. {\"type\":\"close\"} #b7e6ac3769ece808","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":103.275,"y":58.025}
{"fill":"#4472C4","h":7,"kind":"rect","w":7,"x":232.391,"y":190.141}
{"baselineY":197.141,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"75%","width":48,"x":242.391}
{"fill":"#ED7D31","geometryPath":"15 commands {\"type\":\"move\",\"x\":180,\"y\":134.75} .. {\"type\":\"close\"} #5826441e15ec62d5","h":153.45,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153.45,"x":103.275,"y":58.025}
{"fill":"#ED7D31","h":7,"kind":"rect","w":7,"x":107.609,"y":65.359}
{"baselineY":72.359,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25%","width":48,"x":117.609}
# text-properties
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":76.1,"color":"#112233","font":"italic 600 28px Georgia","kind":"text","text":"Revenue","width":98,"x":131}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.3,"x2":244.8,"y1":100.26,"y2":100.26}
{"baselineY":102.26,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"20","width":8,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.3,"x2":244.8,"y1":118.07,"y2":118.07}
{"baselineY":120.07,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"15","width":8,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.3,"x2":244.8,"y1":135.88,"y2":135.88}
{"baselineY":137.88,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"10","width":8,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.3,"x2":244.8,"y1":153.69,"y2":153.69}
{"baselineY":155.69,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"5","width":4,"x":63.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.3,"x2":244.8,"y1":171.5,"y2":171.5}
{"baselineY":173.5,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"0","width":4,"x":63.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":79.3,"x2":79.3,"y1":100.26,"y2":171.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":79.3,"x2":244.8,"y1":171.5,"y2":171.5}
{"baselineY":205.5,"color":"#112233","font":"400 20px Georgia","kind":"text","text":"Q1","width":20,"x":110.675}
{"fill":"#4472C4","h":35.62,"kind":"rect","w":33.1,"x":104.125,"y":135.88}
{"baselineY":132.88,"color":"#112233","font":"700 16px Georgia","kind":"text","text":"10","width":16,"x":112.675}
{"baselineY":205.5,"color":"#112233","font":"400 20px Georgia","kind":"text","text":"Q2","width":20,"x":193.425}
{"fill":"#4472C4","h":71.24,"kind":"rect","w":33.1,"x":186.875,"y":100.26}
{"baselineY":97.26,"color":"#112233","font":"700 16px Georgia","kind":"text","text":"20","width":16,"x":195.425}
{"fill":"#4472C4","h":6.36,"kind":"rect","w":6.36,"x":256.8,"y":154.65}
{"baselineY":160.71,"color":"#112233","font":"700 12px Georgia","kind":"text","text":"North","width":30,"x":267}
# zero-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":86.76,"y2":86.76}
{"baselineY":89.26,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":91.56,"y2":91.56}
{"baselineY":94.06,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":96.36,"y2":96.36}
{"baselineY":98.86,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":101.16,"y2":101.16}
{"baselineY":103.66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":105.96,"y2":105.96}
{"baselineY":108.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":105.96}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":106.5,"y1":105.96,"y2":105.96}
{"baselineY":126.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":83.5}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":3.429,"x":85.071,"y":96.36}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":3.429,"x":88.5,"y":102.12}
{"baselineY":126.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":95.5}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":3.429,"x":97.071,"y":86.76}
{"fill":"#4472C4","h":24,"kind":"rect","w":3.429,"x":100.5,"y":81.96}
# tiny-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":86.76,"y2":86.76}
{"baselineY":89.26,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":91.56,"y2":91.56}
{"baselineY":94.06,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":96.36,"y2":96.36}
{"baselineY":98.86,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":101.16,"y2":101.16}
{"baselineY":103.66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":106.5,"y1":105.96,"y2":105.96}
{"baselineY":108.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":105.96}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":106.5,"y1":105.96,"y2":105.96}
{"baselineY":126.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":83.5}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":3.429,"x":85.071,"y":96.36}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":3.429,"x":88.5,"y":102.12}
{"baselineY":126.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":95.5}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":3.429,"x":97.071,"y":86.76}
{"fill":"#4472C4","h":24,"kind":"rect","w":3.429,"x":100.5,"y":81.96}
# tiny-rect-pie
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, pie chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"19 commands {\"type\":\"move\",\"x\":50,\"y\":62.68} .. {\"type\":\"close\"} #4d64fe0ca2a62fb9","h":20,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20,"x":40,"y":52.68}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":50,\"y\":62.68} .. {\"type\":\"close\"} #9a05b70a269ac277","h":20,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20,"x":40,"y":52.68}
# wide-flat-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":26,"kind":"rect","w":900,"x":50,"y":40}
{"baselineY":61.85,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":477.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":927,"y1":81.96,"y2":81.96}
{"baselineY":84.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":927,"y1":86.76,"y2":86.76}
{"baselineY":89.26,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":927,"y1":91.56,"y2":91.56}
{"baselineY":94.06,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":927,"y1":96.36,"y2":96.36}
{"baselineY":98.86,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":10,"x":59.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":927,"y1":101.16,"y2":101.16}
{"baselineY":103.66,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":5,"x":64.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":82.5,"x2":927,"y1":105.96,"y2":105.96}
{"baselineY":108.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":5,"x":64.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":82.5,"y1":81.96,"y2":105.96}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":82.5,"x2":927,"y1":105.96,"y2":105.96}
{"baselineY":126.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":288.625}
{"baselineY":126.46,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":710.875}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":293.625,\"y\":91.693} .. {\"type\":\"close\"} #a78dde6eb507a593","h":9.333,"kind":"shape","w":9.333,"x":288.958,"y":91.693}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":293.625,"x2":715.875,"y1":96.36,"y2":86.76}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":715.875,\"y\":82.093} .. {\"type\":\"close\"} #ec6c31ef02d09c24","h":9.333,"kind":"shape","w":9.333,"x":711.208,"y":82.093}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":288.958,"y":97.453}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":293.625,"x2":715.875,"y1":102.12,"y2":81.96}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":711.208,"y":77.293}
"##;

#[test]
fn a_gridline_whose_sp_pr_draws_no_line_is_hidden() {
    let gridlines = |line: Option<Value>| {
        let mut axis = json!({
            "id": "2",
            "axisType": "value",
            "min": 0.0,
            "max": 20.0,
            "majorGridlines": true
        });
        if let Some(line) = line {
            axis["majorGridlineLine"] = line;
        }
        let chart = json!({
            "type": "chart",
            "chartType": "column",
            "legend": { "visible": false },
            "axisList": [{ "id": "1", "axisType": "category" }, axis],
            "plotGroups": [{
                "chartType": "column",
                "axisIds": ["1", "2"],
                "series": [series("Units", json!([10.0, 15.0]))]
            }]
        });
        display_list(chart, 260.0, 180.0)["pages"][0]["primitives"]
            .as_array()
            .expect("primitives")
            .iter()
            .filter(|primitive| primitive["kind"] == "line" && primitive["color"] == "#D9D9D9")
            .count()
    };
    assert!(gridlines(None) > 0);
    assert_eq!(gridlines(Some(json!({ "none": false }))), gridlines(None));
    assert_eq!(gridlines(Some(json!({ "none": true }))), 0);
}
