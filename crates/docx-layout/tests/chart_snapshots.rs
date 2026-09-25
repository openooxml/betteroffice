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
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":23.071,"x":109.304,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":23.071,"x":132.375,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":23.071,"x":190.054,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":23.071,"x":213.125,"y":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-bar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":253.5,"x2":253.5,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":32,"x":237.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":228,"x2":228,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":32,"x":212}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":202.5,"x2":202.5,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":32,"x":186.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177,"x2":177,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":161}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":151.5,"x2":151.5,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":32,"x":135.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":68,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":253.5,"y1":186,"y2":186}
{"baselineY":159.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":51,"x":126,"y":139.643}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":20.4,"x":126,"y":156.5}
{"baselineY":100.45,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":102,"x":126,"y":80.643}
{"fill":"#4472C4","h":16.857,"kind":"rect","w":127.5,"x":126,"y":97.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-line
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":134.133} .. {\"type\":\"close\"} #a5a70e7708fdecfa","h":9.333,"kind":"shape","w":9.333,"x":127.708,"y":134.133}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":213.125,\"y\":86.933} .. {\"type\":\"close\"} #8d2ba3ead3467996","h":9.333,"kind":"shape","w":9.333,"x":208.458,"y":86.933}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":127.708,"y":162.453}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":208.458,"y":63.333}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-pie
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, pie chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"fill":"#4472C4","geometryPath":"19 commands {\"type\":\"move\",\"x\":159.25,\"y\":144} .. {\"type\":\"close\"} #975f16f3087afc2b","h":136.8,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":136.8,"x":90.85,"y":75.6}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":159.25,\"y\":144} .. {\"type\":\"close\"} #198d35aca81e40a5","h":136.8,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":136.8,"x":90.85,"y":75.6}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
# type-doughnut
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, doughnut chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":159.25,\"y\":75.6} .. {\"type\":\"close\"} #14c28ee32ad92dea","h":136.8,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":136.8,"x":90.85,"y":75.6}
{"fill":"#4472C4","geometryPath":"67 commands {\"type\":\"move\",\"x\":218.486,\"y\":178.2} .. {\"type\":\"close\"} #d83d5701c2efdb5d","h":136.8,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":136.8,"x":90.85,"y":75.6}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
# type-area
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":138.8} .. {\"type\":\"close\"} #cbc78f94ef1b9fb6","h":118,"kind":"shape","w":161.5,"x":92,"y":68}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.375,"x2":213.125,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":167.12} .. {\"type\":\"close\"} #877fcdaf303dacf9","h":118,"kind":"shape","w":161.5,"x":92,"y":68}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.375,"x2":213.125,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-scatter
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, scatter chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-radar
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, radar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":133.664,"y2":154.336}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":123.328,"y2":164.672}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":112.992,"y2":175.008}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":102.656,"y2":185.344}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":92.32,"y2":195.68}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":144,"y2":92.32}
{"baselineY":87.152,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":32,"x":111.33}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":127.33,"x2":127.33,"y1":144,"y2":195.68}
{"baselineY":200.848,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":32,"x":111.33}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":127.33,"x2":127.33,"y1":123.328,"y2":185.344}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":127.33,\"y\":118.661} .. {\"type\":\"close\"} #0366a6a6642c0446","h":9.333,"kind":"shape","w":9.333,"x":122.663,"y":118.661}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":127.33,\"y\":180.677} .. {\"type\":\"close\"} #7b2231d8d85816f3","h":9.333,"kind":"shape","w":9.333,"x":122.663,"y":180.677}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":127.33,"x2":127.33,"y1":135.731,"y2":195.68}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":122.663,"y":131.065}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":122.663,"y":191.013}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-stock
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, stock chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":134.133} .. {\"type\":\"close\"} #a5a70e7708fdecfa","h":9.333,"kind":"shape","w":9.333,"x":127.708,"y":134.133}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":138.8,"y2":91.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":213.125,\"y\":86.933} .. {\"type\":\"close\"} #8d2ba3ead3467996","h":9.333,"kind":"shape","w":9.333,"x":208.458,"y":86.933}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":127.708,"y":162.453}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":167.12,"y2":68}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":208.458,"y":63.333}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-bubble
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, bubble chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-surface
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, surface chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"fill":"#70AD47","h":59,"kind":"rect","w":80.75,"x":92,"y":127}
{"fill":"#ED7D31","h":59,"kind":"rect","w":80.75,"x":172.75,"y":127}
{"baselineY":162.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":36,"x":54}
{"fill":"#4472C4","h":59,"kind":"rect","w":80.75,"x":92,"y":68}
{"fill":"#9E480E","h":59,"kind":"rect","w":80.75,"x":172.75,"y":68}
{"baselineY":103.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":36,"x":54}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":76.75,"x":94}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":76.75,"x":174.75}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-mystery
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":23.071,"x":109.304,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":23.071,"x":132.375,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":23.071,"x":190.054,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":23.071,"x":213.125,"y":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# type-
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":23.071,"x":109.304,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":23.071,"x":132.375,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":23.071,"x":190.054,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":23.071,"x":213.125,"y":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# legend-left
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.5,"x2":292,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":110.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.5,"x2":292,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":110.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.5,"x2":292,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":110.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.5,"x2":292,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":110.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.5,"x2":292,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":110.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.5,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":110.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":148.5,"x2":148.5,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":148.5,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":179.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":20.5,"x":163.875,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":20.5,"x":184.375,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":251.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":20.5,"x":235.625,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":20.5,"x":256.125,"y":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":71.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":71.5}
# legend-right
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":23.071,"x":109.304,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":23.071,"x":132.375,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":23.071,"x":190.054,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":23.071,"x":213.125,"y":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# legend-top
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":90,"y2":90}
{"baselineY":93,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":109.2,"y2":109.2}
{"baselineY":112.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":128.4,"y2":128.4}
{"baselineY":131.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":147.6,"y2":147.6}
{"baselineY":150.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":166.8,"y2":166.8}
{"baselineY":169.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":90,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":137}
{"fill":"#4472C4","h":38.4,"kind":"rect","w":28.571,"x":113.429,"y":147.6}
{"fill":"#4472C4","h":15.36,"kind":"rect","w":28.571,"x":142,"y":170.64}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237}
{"fill":"#4472C4","h":76.8,"kind":"rect","w":28.571,"x":213.429,"y":109.2}
{"fill":"#4472C4","h":96,"kind":"rect","w":28.571,"x":242,"y":90}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":136,"y":75}
{"baselineY":82.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":39,"x":148}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":187,"y":75}
{"baselineY":82.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":39,"x":199}
# legend-bottom
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":87.2,"y2":87.2}
{"baselineY":90.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":106.4,"y2":106.4}
{"baselineY":109.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":125.6,"y2":125.6}
{"baselineY":128.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":144.8,"y2":144.8}
{"baselineY":147.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":164,"y2":164}
{"baselineY":167,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":164}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":164,"y2":164}
{"baselineY":178,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":137}
{"fill":"#4472C4","h":38.4,"kind":"rect","w":28.571,"x":113.429,"y":125.6}
{"fill":"#4472C4","h":15.36,"kind":"rect","w":28.571,"x":142,"y":148.64}
{"baselineY":178,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237}
{"fill":"#4472C4","h":76.8,"kind":"rect","w":28.571,"x":213.429,"y":87.2}
{"fill":"#4472C4","h":96,"kind":"rect","w":28.571,"x":242,"y":68}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":136,"y":205}
{"baselineY":212.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":39,"x":148}
{"fill":"#4472C4","h":8,"kind":"rect","w":8,"x":187,"y":205}
{"baselineY":212.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":39,"x":199}
# legend-hidden
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
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
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":137}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":28.571,"x":113.429,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":28.571,"x":142,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":28.571,"x":213.429,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":28.571,"x":242,"y":68}
# legend-default-position
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":23.071,"x":109.304,"y":138.8}
{"fill":"#4472C4","h":18.88,"kind":"rect","w":23.071,"x":132.375,"y":167.12}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":23.071,"x":190.054,"y":91.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":23.071,"x":213.125,"y":68}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# legend-overflow
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 12 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":238.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":238.5,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":238.5,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":238.5,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":238.5,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":238.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":238.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":123.625}
{"fill":"#4472C4","h":1,"kind":"rect","w":5.426,"x":96.069,"y":186}
{"fill":"#ED7D31","h":5.44,"kind":"rect","w":5.426,"x":101.495,"y":180.56}
{"fill":"#A5A5A5","h":10.88,"kind":"rect","w":5.426,"x":106.921,"y":175.12}
{"fill":"#FFC000","h":16.32,"kind":"rect","w":5.426,"x":112.347,"y":169.68}
{"fill":"#5B9BD5","h":21.76,"kind":"rect","w":5.426,"x":117.773,"y":164.24}
{"fill":"#70AD47","h":27.2,"kind":"rect","w":5.426,"x":123.199,"y":158.8}
{"fill":"#264478","h":32.64,"kind":"rect","w":5.426,"x":128.625,"y":153.36}
{"fill":"#9E480E","h":38.08,"kind":"rect","w":5.426,"x":134.051,"y":147.92}
{"fill":"#4472C4","h":43.52,"kind":"rect","w":5.426,"x":139.477,"y":142.48}
{"fill":"#ED7D31","h":48.96,"kind":"rect","w":5.426,"x":144.903,"y":137.04}
{"fill":"#A5A5A5","h":54.4,"kind":"rect","w":5.426,"x":150.329,"y":131.6}
{"fill":"#FFC000","h":59.84,"kind":"rect","w":5.426,"x":155.755,"y":126.16}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":196.875}
{"fill":"#4472C4","h":1,"kind":"rect","w":5.426,"x":169.319,"y":186}
{"fill":"#ED7D31","h":10.88,"kind":"rect","w":5.426,"x":174.745,"y":175.12}
{"fill":"#A5A5A5","h":21.76,"kind":"rect","w":5.426,"x":180.171,"y":164.24}
{"fill":"#FFC000","h":32.64,"kind":"rect","w":5.426,"x":185.597,"y":153.36}
{"fill":"#5B9BD5","h":43.52,"kind":"rect","w":5.426,"x":191.023,"y":142.48}
{"fill":"#70AD47","h":54.4,"kind":"rect","w":5.426,"x":196.449,"y":131.6}
{"fill":"#264478","h":65.28,"kind":"rect","w":5.426,"x":201.875,"y":120.72}
{"fill":"#9E480E","h":76.16,"kind":"rect","w":5.426,"x":207.301,"y":109.84}
{"fill":"#4472C4","h":87.04,"kind":"rect","w":5.426,"x":212.727,"y":98.96}
{"fill":"#ED7D31","h":97.92,"kind":"rect","w":5.426,"x":218.153,"y":88.08}
{"fill":"#A5A5A5","h":108.8,"kind":"rect","w":5.426,"x":223.579,"y":77.2}
{"fill":"#FFC000","h":119.68,"kind":"rect","w":5.426,"x":229.005,"y":66.32}
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
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #29c83c09b22b578a","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#ED7D31","geometryPath":"5 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #910671139c3f57ca","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#A5A5A5","geometryPath":"6 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #fb225c0e0cb2886e","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#FFC000","geometryPath":"7 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #ec5d58a94227751c","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#5B9BD5","geometryPath":"8 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #7838a5a4f9f5a46e","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#70AD47","geometryPath":"9 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #5b9fd9f2979442a3","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#264478","geometryPath":"10 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #153d1a8a011791bd","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#9E480E","geometryPath":"10 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #261d0897420cebec","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #d851fab3b025fcc8","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
{"fill":"#ED7D31","geometryPath":"12 commands {\"type\":\"move\",\"x\":161.75,\"y\":135} .. {\"type\":\"close\"} #618af8bdd2492485","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":85.25,"y":58.5}
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
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Combo","width":32.5,"x":163.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":79.8,"y2":79.8}
{"baselineY":82.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":103.4,"y2":103.4}
{"baselineY":106.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":150.6,"y2":150.6}
{"baselineY":153.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":174.2,"y2":174.2}
{"baselineY":177.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":243.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124.875}
{"fill":"#4472C4","h":59,"kind":"rect","w":30.3,"x":114.725,"y":127}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":200.625}
{"fill":"#4472C4","h":106.2,"kind":"rect","w":30.3,"x":190.475,"y":79.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":81.111,"y2":81.111}
{"baselineY":84.111,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":94.222,"y2":94.222}
{"baselineY":97.222,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":107.333,"y2":107.333}
{"baselineY":110.333,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":120.444,"y2":120.444}
{"baselineY":123.444,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":133.556,"y2":133.556}
{"baselineY":136.556,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":146.667,"y2":146.667}
{"baselineY":149.667,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":159.778,"y2":159.778}
{"baselineY":162.778,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":172.889,"y2":172.889}
{"baselineY":175.889,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":243.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":124.875}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":200.625}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":129.875,\"y\":128.889} .. {\"type\":\"close\"} #ba8d17e1c76eeb87","h":9.333,"kind":"shape","w":9.333,"x":125.208,"y":128.889}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":129.875,"x2":205.625,"y1":133.556,"y2":81.111}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":205.625,\"y\":76.444} .. {\"type\":\"close\"} #23377f8a4ef3c582","h":9.333,"kind":"shape","w":9.333,"x":200.958,"y":76.444}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":253.5,"y":133.7}
{"baselineY":138.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Revenue","width":35,"x":262}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":253.5,"y":149}
{"baselineY":154.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Trend","width":35,"x":262}
# combo-with-pie-group
attrs {"blockId":42,"chart":{"label":"Untitled chart, combo chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"39 commands {\"type\":\"move\",\"x\":200.75,\"y\":135} .. {\"type\":\"close\"} #c7576a22f8da74a8","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":124.25,"y":58.5}
{"fill":"#4472C4","geometryPath":"15 commands {\"type\":\"move\",\"x\":200.75,\"y\":135} .. {\"type\":\"close\"} #05afe41dba070cf5","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":124.25,"y":58.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":292,"x2":292,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":32,"x":276}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":274.214,"x2":274.214,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":32,"x":258.214}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":256.429,"x2":256.429,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":32,"x":240.429}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":238.643,"x2":238.643,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":32,"x":222.643}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":220.857,"x2":220.857,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":32,"x":204.857}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":203.071,"x2":203.071,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":32,"x":187.071}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":185.286,"x2":185.286,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":32,"x":169.286}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":167.5,"x2":167.5,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":151.5}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":167.5,"x2":167.5,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":167.5,"x2":292,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":95.5}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":35.571,"x":167.5,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":95.5}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":106.714,"x":167.5,"y":70.4}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":71.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":63,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":71.5}
# points-markers-labels
attrs {"ariaLabel":"Points","blockId":42,"chart":{"label":"Points, line chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Points","width":39,"x":160.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":82.75,"y2":82.75}
{"baselineY":85.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":112.25,"y2":112.25}
{"baselineY":115.25,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":127,"y2":127}
{"baselineY":130,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":141.75,"y2":141.75}
{"baselineY":144.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":171.25,"y2":171.25}
{"baselineY":174.25,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":113.917}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":167.75}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":10,"x":221.583}
{"fill":"#FF0000","geometryPath":"5 commands {\"type\":\"move\",\"x\":118.917,\"y\":73.417} .. {\"type\":\"close\"} #17c3015b1af68c56","h":18.667,"kind":"shape","w":18.667,"x":109.583,"y":73.417}
{"baselineY":86.25,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"peak","width":48,"x":131.25}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":118.917,"x2":172.75,"y1":82.75,"y2":156.5}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":172.75,\"y\":150.5} .. {\"type\":\"close\"} #fce447b7cef46a92","h":12,"kind":"shape","w":12,"x":166.75,"y":150.5}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":172.75,"x2":226.583,"y1":156.5,"y2":141.75}
{"fill":"#00FF00","geometryPath":"5 commands {\"type\":\"move\",\"x\":226.583,\"y\":135.75} .. {\"type\":\"close\"} #2d13af50a3f0f678","h":12,"kind":"shape","w":12,"x":220.583,"y":135.75}
{"baselineY":145.25,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"end","width":48,"x":235.583}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":141.35}
{"baselineY":146.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# points-without-indexes
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#123456","geometryPath":"39 commands {\"type\":\"move\",\"x\":172.773,\"y\":148.523} .. {\"type\":\"close\"} #9cb4e6fffcced4d7","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":96.273,"y":72.023}
{"fill":"#123456","geometryPath":"15 commands {\"type\":\"move\",\"x\":145.727,\"y\":121.477} .. {\"type\":\"close\"} #27001c5b2755f878","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":69.227,"y":44.977}
{"fill":"#123456","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#123456","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
# negative-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-30","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":19.429,"kind":"rect","w":23.071,"x":109.304,"y":108.286}
{"fill":"#4472C4","h":7.771,"kind":"rect","w":23.071,"x":132.375,"y":108.286}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":38.857,"kind":"rect","w":23.071,"x":190.054,"y":69.429}
{"fill":"#4472C4","h":58.286,"kind":"rect","w":23.071,"x":213.125,"y":108.286}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# negative-values-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":253.5,"x2":253.5,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":32,"x":237.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":228,"x2":228,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":32,"x":212}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":202.5,"x2":202.5,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":186.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177,"x2":177,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":161}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":151.5,"x2":151.5,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":32,"x":135.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":253.5,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":25.5,"x":151.5,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":51,"x":177,"y":70.4}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# inverted-axis-bounds
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"11","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10.2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":1,"kind":"rect","w":32.3,"x":116.225,"y":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":1,"kind":"rect","w":32.3,"x":196.975,"y":186}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# long-text
attrs {"ariaLabel":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT","blockId":42,"chart":{"label":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT","width":780,"x":-210}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":203.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":203.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":203.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":203.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":203.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":203.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":203.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC","width":600,"x":-180.125}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":22.3,"x":108.725,"y":138.8}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":170.625}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":22.3,"x":164.475,"y":91.6}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":213.5,"y":98.65}
{"baselineY":103.7,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":115.9,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":128.1,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":140.3,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":152.5,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":164.7,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":176.9,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
{"baselineY":189.1,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"NNNNNNNNNNNNNNN","width":75,"x":222}
# no-title-no-legend
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":137}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":142,\"y\":126.933} .. {\"type\":\"close\"} #bc04e87eee25a1dd","h":9.333,"kind":"shape","w":9.333,"x":137.333,"y":126.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":142,"x2":242,"y1":131.6,"y2":77.2}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":242,\"y\":72.533} .. {\"type\":\"close\"} #04a03f84a8d3be36","h":9.333,"kind":"shape","w":9.333,"x":237.333,"y":72.533}
# empty-series
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 0 series, 0 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
# series-without-values
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 0 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
# described-and-decorative
attrs {"ariaDescription":"quarterly revenue","ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 1 series, 2 categories"},"decorative":true,"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":157.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":91.6,"y2":91.6}
{"baselineY":94.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":115.2,"y2":115.2}
{"baselineY":118.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":138.8,"y2":138.8}
{"baselineY":141.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":162.4,"y2":162.4}
{"baselineY":165.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":47.2,"kind":"rect","w":32.3,"x":116.225,"y":138.8}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":94.4,"kind":"rect","w":32.3,"x":196.975,"y":91.6}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":141.35}
{"baselineY":146.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# stacked-column
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":38.857,"kind":"rect","w":32.3,"x":116.225,"y":108.286}
{"fill":"#4472C4","h":19.429,"kind":"rect","w":32.3,"x":116.225,"y":88.857}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":77.714,"kind":"rect","w":32.3,"x":196.975,"y":69.429}
{"fill":"#4472C4","h":31.086,"kind":"rect","w":32.3,"x":196.975,"y":147.143}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# stacked-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":253.5,"x2":253.5,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":32,"x":237.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":235.286,"x2":235.286,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":32,"x":219.286}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":217.071,"x2":217.071,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":32,"x":201.071}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":198.857,"x2":198.857,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":32,"x":182.857}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":180.643,"x2":180.643,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":32,"x":164.643}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":162.429,"x2":162.429,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":32,"x":146.429}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":144.214,"x2":144.214,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":32,"x":128.214}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":253.5,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":36.429,"x":162.429,"y":138.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":18.214,"x":198.857,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":72.857,"x":162.429,"y":70.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":29.143,"x":133.286,"y":70.4}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# stacked-line
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":103.619} .. {\"type\":\"close\"} #08cea822d0eb3770","h":9.333,"kind":"shape","w":9.333,"x":127.708,"y":103.619}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":108.286,"y2":69.429}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":213.125,\"y\":64.762} .. {\"type\":\"close\"} #217819420a2c995d","h":9.333,"kind":"shape","w":9.333,"x":208.458,"y":64.762}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":127.708,"y":84.19}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":88.857,"y2":178.229}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":208.458,"y":173.562}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# stacked-area
attrs {"blockId":42,"chart":{"label":"Untitled chart, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-10","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":108.286} .. {\"type\":\"close\"} #f957a278806485e5","h":136,"kind":"shape","w":161.5,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.375,"x2":213.125,"y1":108.286,"y2":69.429}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":88.857} .. {\"type\":\"close\"} #e6e6a1faa8155c3c","h":136,"kind":"shape","w":161.5,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.375,"x2":213.125,"y1":88.857,"y2":178.229}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# percentStacked-column
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":67,"y2":67}
{"baselineY":70,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":101,"y2":101}
{"baselineY":104,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":135,"y2":135}
{"baselineY":138,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":169,"y2":169}
{"baselineY":172,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":56.667,"kind":"rect","w":32.3,"x":116.225,"y":95.333}
{"fill":"#4472C4","h":28.333,"kind":"rect","w":32.3,"x":116.225,"y":67}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":60.714,"kind":"rect","w":32.3,"x":196.975,"y":91.286}
{"fill":"#4472C4","h":24.286,"kind":"rect","w":32.3,"x":196.975,"y":152}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# percentStacked-bar
attrs {"blockId":42,"chart":{"label":"Untitled chart, bar chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":253.5,"x2":253.5,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":32,"x":237.5}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":237.563,"x2":237.563,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":32,"x":221.563}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":221.625,"x2":221.625,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":32,"x":205.625}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":205.688,"x2":205.688,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":32,"x":189.688}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":189.75,"x2":189.75,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":32,"x":173.75}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":173.813,"x2":173.813,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":32,"x":157.813}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":157.875,"x2":157.875,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":32,"x":141.875}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":141.938,"x2":141.938,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":32,"x":125.938}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":126,"x2":126,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":32,"x":110}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":126,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":126,"x2":253.5,"y1":186,"y2":186}
{"baselineY":155.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":53.125,"x":157.875,"y":138.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":26.563,"x":211,"y":138.4}
{"baselineY":87.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":68,"x":54}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":56.92,"x":157.875,"y":70.4}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":22.768,"x":135.107,"y":70.4}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# percentStacked-line
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":67,"y2":67}
{"baselineY":70,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":101,"y2":101}
{"baselineY":104,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":135,"y2":135}
{"baselineY":138,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":169,"y2":169}
{"baselineY":172,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":90.667} .. {\"type\":\"close\"} #8d470b8de50e2c45","h":9.333,"kind":"shape","w":9.333,"x":127.708,"y":90.667}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":95.333,"y2":91.286}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":213.125,\"y\":86.619} .. {\"type\":\"close\"} #4c435efd4de46663","h":9.333,"kind":"shape","w":9.333,"x":208.458,"y":86.619}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":127.708,"y":62.333}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.375,"x2":213.125,"y1":67,"y2":176.286}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":208.458,"y":171.619}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# percentStacked-area
attrs {"blockId":42,"chart":{"label":"Untitled chart, area chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"120%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":67,"y2":67}
{"baselineY":70,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":84,"y2":84}
{"baselineY":87,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":101,"y2":101}
{"baselineY":104,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":135,"y2":135}
{"baselineY":138,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":152,"y2":152}
{"baselineY":155,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":169,"y2":169}
{"baselineY":172,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-20%","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"-40%","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":95.333} .. {\"type\":\"close\"} #7f89cc1f4aae1e49","h":136,"kind":"shape","w":161.5,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.375,"x2":213.125,"y1":95.333,"y2":91.286}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.375,\"y\":67} .. {\"type\":\"close\"} #fdeb1eafb91bb5be","h":136,"kind":"shape","w":161.5,"x":92,"y":50}
{"color":"#4472C4","kind":"line","strokeWidth":1.5,"x1":132.375,"x2":213.125,"y1":67,"y2":176.286}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
# bar-gap-and-overlap
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"35","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":69.429,"y2":69.429}
{"baselineY":72.429,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"30","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":88.857,"y2":88.857}
{"baselineY":91.857,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":108.286,"y2":108.286}
{"baselineY":111.286,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":127.714,"y2":127.714}
{"baselineY":130.714,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":147.143,"y2":147.143}
{"baselineY":150.143,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":166.571,"y2":166.571}
{"baselineY":169.571,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.375}
{"fill":"#4472C4","h":38.857,"kind":"rect","w":31.058,"x":98.212,"y":147.143}
{"fill":"#4472C4","h":15.543,"kind":"rect","w":31.058,"x":135.481,"y":170.457}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.125}
{"fill":"#4472C4","h":77.714,"kind":"rect","w":31.058,"x":178.962,"y":108.286}
{"fill":"#4472C4","h":116.571,"kind":"rect","w":31.058,"x":216.231,"y":69.429}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North","width":25,"x":272}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":263.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"South","width":25,"x":272}
# scatter-xy
attrs {"blockId":42,"chart":{"label":"Untitled chart, scatter chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":63.6,"y2":63.6}
{"baselineY":66.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":90.8,"y2":90.8}
{"baselineY":93.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":145.2,"y2":145.2}
{"baselineY":148.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":172.4,"y2":172.4}
{"baselineY":175.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":268.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":268.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":32,"x":98.063}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":32,"x":120.125}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":32,"x":142.188}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":32,"x":164.25}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":32,"x":186.313}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":32,"x":208.375}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":32,"x":230.438}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":32,"x":252.5}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":92,\"y\":140.533} .. {\"type\":\"close\"} #15a8e48350caf9b9","h":9.333,"kind":"shape","w":9.333,"x":87.333,"y":140.533}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":92,"x2":180.25,"y1":145.2,"y2":63.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":180.25,\"y\":58.933} .. {\"type\":\"close\"} #5b6b6a6adeaddcb0","h":9.333,"kind":"shape","w":9.333,"x":175.583,"y":58.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":180.25,"x2":268.5,"y1":63.6,"y2":131.6}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":268.5,\"y\":126.933} .. {\"type\":\"close\"} #61abd2ddec39ad1e","h":9.333,"kind":"shape","w":9.333,"x":263.833,"y":126.933}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"XY","width":10,"x":287}
# bubble-sizes
attrs {"blockId":42,"chart":{"label":"Untitled chart, bubble chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":63.6,"y2":63.6}
{"baselineY":66.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"9","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"8","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":90.8,"y2":90.8}
{"baselineY":93.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"7","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"6","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":118,"y2":118}
{"baselineY":121,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":145.2,"y2":145.2}
{"baselineY":148.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":172.4,"y2":172.4}
{"baselineY":175.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":243.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":243.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":32,"x":76}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":32,"x":113.875}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"3","width":32,"x":151.75}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"4","width":32,"x":189.625}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":32,"x":227.5}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":102.2,\"y\":145.2} .. {\"type\":\"close\"} #a68792eaac6a17f6","h":20.4,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20.4,"x":81.8,"y":135}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":263.9,\"y\":63.6} .. {\"type\":\"close\"} #dd7337385d9ed531","h":40.8,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":40.8,"x":223.1,"y":43.2}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":253.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Bubbles","width":35,"x":262}
# radar-standard
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":131.21,"y1":129.22,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":131.21,"x2":125.43,"y1":135,"y2":140.78}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":119.65,"y1":140.78,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":119.65,"x2":125.43,"y1":135,"y2":129.22}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":136.99,"y1":123.44,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":136.99,"x2":125.43,"y1":135,"y2":146.56}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":113.87,"y1":146.56,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":113.87,"x2":125.43,"y1":135,"y2":123.44}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":142.77,"y1":117.66,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.77,"x2":125.43,"y1":135,"y2":152.34}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":108.09,"y1":152.34,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":108.09,"x2":125.43,"y1":135,"y2":117.66}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":148.55,"y1":111.88,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.55,"x2":125.43,"y1":135,"y2":158.12}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":102.31,"y1":158.12,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.31,"x2":125.43,"y1":135,"y2":111.88}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":154.33,"y1":106.1,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154.33,"x2":125.43,"y1":135,"y2":163.9}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":96.53,"y1":163.9,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":96.53,"x2":125.43,"y1":135,"y2":106.1}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":160.11,"y1":100.32,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.11,"x2":125.43,"y1":135,"y2":169.68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":90.75,"y1":169.68,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":90.75,"x2":125.43,"y1":135,"y2":100.32}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":165.89,"y1":94.54,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":165.89,"x2":125.43,"y1":135,"y2":175.46}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":84.97,"y1":175.46,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":84.97,"x2":125.43,"y1":135,"y2":94.54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":171.67,"y1":88.76,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.67,"x2":125.43,"y1":135,"y2":181.24}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":79.19,"y1":181.24,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.19,"x2":125.43,"y1":135,"y2":88.76}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":177.45,"y1":82.98,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.45,"x2":125.43,"y1":135,"y2":187.02}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":73.41,"y1":187.02,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":73.41,"x2":125.43,"y1":135,"y2":82.98}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.23,"y1":77.2,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.23,"x2":125.43,"y1":135,"y2":192.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.63,"y1":192.8,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":67.63,"x2":125.43,"y1":135,"y2":77.2}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":135,"y2":77.2}
{"baselineY":71.42,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.23,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":173.01}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":135,"y2":192.8}
{"baselineY":198.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.63,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":45.85}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":177.45,"y1":117.66,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":177.45,"x2":125.43,"y1":135,"y2":158.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":90.75,"y1":158.12,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":90.75,"x2":125.43,"y1":135,"y2":117.66}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":258.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":30,"x":267}
# radar-marker
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":131.21,"y1":129.22,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":131.21,"x2":125.43,"y1":135,"y2":140.78}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":119.65,"y1":140.78,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":119.65,"x2":125.43,"y1":135,"y2":129.22}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":136.99,"y1":123.44,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":136.99,"x2":125.43,"y1":135,"y2":146.56}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":113.87,"y1":146.56,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":113.87,"x2":125.43,"y1":135,"y2":123.44}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":142.77,"y1":117.66,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.77,"x2":125.43,"y1":135,"y2":152.34}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":108.09,"y1":152.34,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":108.09,"x2":125.43,"y1":135,"y2":117.66}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":148.55,"y1":111.88,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.55,"x2":125.43,"y1":135,"y2":158.12}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":102.31,"y1":158.12,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.31,"x2":125.43,"y1":135,"y2":111.88}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":154.33,"y1":106.1,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154.33,"x2":125.43,"y1":135,"y2":163.9}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":96.53,"y1":163.9,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":96.53,"x2":125.43,"y1":135,"y2":106.1}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":160.11,"y1":100.32,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.11,"x2":125.43,"y1":135,"y2":169.68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":90.75,"y1":169.68,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":90.75,"x2":125.43,"y1":135,"y2":100.32}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":165.89,"y1":94.54,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":165.89,"x2":125.43,"y1":135,"y2":175.46}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":84.97,"y1":175.46,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":84.97,"x2":125.43,"y1":135,"y2":94.54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":171.67,"y1":88.76,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.67,"x2":125.43,"y1":135,"y2":181.24}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":79.19,"y1":181.24,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.19,"x2":125.43,"y1":135,"y2":88.76}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":177.45,"y1":82.98,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.45,"x2":125.43,"y1":135,"y2":187.02}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":73.41,"y1":187.02,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":73.41,"x2":125.43,"y1":135,"y2":82.98}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.23,"y1":77.2,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.23,"x2":125.43,"y1":135,"y2":192.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.63,"y1":192.8,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":67.63,"x2":125.43,"y1":135,"y2":77.2}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":135,"y2":77.2}
{"baselineY":71.42,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.23,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":173.01}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":135,"y2":192.8}
{"baselineY":198.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.63,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":45.85}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":177.45,"y1":117.66,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":177.45,"x2":125.43,"y1":135,"y2":158.12}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":125.43,"x2":90.75,"y1":158.12,"y2":135}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":90.75,"x2":125.43,"y1":135,"y2":117.66}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.43,\"y\":112.993} .. {\"type\":\"close\"} #afad46455c9a3255","h":9.333,"kind":"shape","w":9.333,"x":120.763,"y":112.993}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":177.45,\"y\":130.333} .. {\"type\":\"close\"} #06cac7f3ac130343","h":9.333,"kind":"shape","w":9.333,"x":172.783,"y":130.333}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.43,\"y\":153.453} .. {\"type\":\"close\"} #58667af046fc6f67","h":9.333,"kind":"shape","w":9.333,"x":120.763,"y":153.453}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":90.75,\"y\":130.333} .. {\"type\":\"close\"} #f8de36e3133b868c","h":9.333,"kind":"shape","w":9.333,"x":86.083,"y":130.333}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":258.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":30,"x":267}
# radar-filled
attrs {"blockId":42,"chart":{"label":"Untitled chart, radar chart, 1 series, 4 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":131.21,"y1":129.22,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":131.21,"x2":125.43,"y1":135,"y2":140.78}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":119.65,"y1":140.78,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":119.65,"x2":125.43,"y1":135,"y2":129.22}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":136.99,"y1":123.44,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":136.99,"x2":125.43,"y1":135,"y2":146.56}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":113.87,"y1":146.56,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":113.87,"x2":125.43,"y1":135,"y2":123.44}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":142.77,"y1":117.66,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":142.77,"x2":125.43,"y1":135,"y2":152.34}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":108.09,"y1":152.34,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":108.09,"x2":125.43,"y1":135,"y2":117.66}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":148.55,"y1":111.88,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":148.55,"x2":125.43,"y1":135,"y2":158.12}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":102.31,"y1":158.12,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":102.31,"x2":125.43,"y1":135,"y2":111.88}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":154.33,"y1":106.1,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":154.33,"x2":125.43,"y1":135,"y2":163.9}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":96.53,"y1":163.9,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":96.53,"x2":125.43,"y1":135,"y2":106.1}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":160.11,"y1":100.32,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":160.11,"x2":125.43,"y1":135,"y2":169.68}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":90.75,"y1":169.68,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":90.75,"x2":125.43,"y1":135,"y2":100.32}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":165.89,"y1":94.54,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":165.89,"x2":125.43,"y1":135,"y2":175.46}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":84.97,"y1":175.46,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":84.97,"x2":125.43,"y1":135,"y2":94.54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":171.67,"y1":88.76,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":171.67,"x2":125.43,"y1":135,"y2":181.24}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":79.19,"y1":181.24,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":79.19,"x2":125.43,"y1":135,"y2":88.76}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":177.45,"y1":82.98,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":177.45,"x2":125.43,"y1":135,"y2":187.02}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":73.41,"y1":187.02,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":73.41,"x2":125.43,"y1":135,"y2":82.98}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.23,"y1":77.2,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":183.23,"x2":125.43,"y1":135,"y2":192.8}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.63,"y1":192.8,"y2":135}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":67.63,"x2":125.43,"y1":135,"y2":77.2}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":135,"y2":77.2}
{"baselineY":71.42,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":183.23,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":32,"x":173.01}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":125.43,"y1":135,"y2":192.8}
{"baselineY":198.58,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":32,"x":109.43}
{"color":"#666666","kind":"line","strokeWidth":0.5,"x1":125.43,"x2":67.63,"y1":135,"y2":135}
{"baselineY":135,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D","width":32,"x":45.85}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":125.43,\"y\":117.66} .. {\"type\":\"close\"} #3a049faf0a3ba58b","h":136,"kind":"shape","w":156.5,"x":92,"y":50}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":258.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Skills","width":30,"x":267}
# stock-ohlc
attrs {"blockId":42,"chart":{"label":"Untitled chart, stock chart, 4 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":253.5,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D1","width":76.75,"x":94}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":132.375,"x2":132.375,"y1":77.2,"y2":158.8}
{"fill":"#FFFFFF","h":43.52,"kind":"rect","w":24,"x":120.375,"y":88.08}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":120.375,"x2":120.375,"y1":88.08,"y2":131.6}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"D2","width":76.75,"x":174.75}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":213.125,"x2":213.125,"y1":66.32,"y2":153.36}
{"fill":"#666666","h":21.76,"kind":"rect","w":24,"x":201.125,"y":120.72}
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
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":268.5,"y1":186,"y2":186}
{"fill":"#4472C4","h":68,"kind":"rect","w":58.833,"x":92,"y":118}
{"fill":"#A9D18E","h":68,"kind":"rect","w":58.833,"x":150.833,"y":118}
{"fill":"#ED7D31","h":68,"kind":"rect","w":58.833,"x":209.667,"y":118}
{"baselineY":158.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R1","width":36,"x":54}
{"fill":"#ED7D31","h":68,"kind":"rect","w":58.833,"x":92,"y":50}
{"fill":"#4472C4","h":68,"kind":"rect","w":58.833,"x":150.833,"y":50}
{"fill":"#A9D18E","h":68,"kind":"rect","w":58.833,"x":209.667,"y":50}
{"baselineY":90.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R2","width":36,"x":54}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"A","width":54.833,"x":94}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"B","width":54.833,"x":152.833}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"C","width":54.833,"x":211.667}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":119.7}
{"baselineY":124.75,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R1","width":10,"x":287}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":135}
{"baselineY":140.05,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"R2","width":10,"x":287}
# doughnut-hole-and-rotation
attrs {"blockId":42,"chart":{"label":"Untitled chart, doughnut chart, 1 series, 3 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"51 commands {\"type\":\"move\",\"x\":235.75,\"y\":135} .. {\"type\":\"close\"} #cdb41c8c28f152cd","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":82.75,"y":58.5}
{"fill":"#ED7D31","geometryPath":"19 commands {\"type\":\"move\",\"x\":62.875,\"y\":123.525} .. {\"type\":\"close\"} #9ac606ab3edece8d","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":62.875,"y":47.025}
{"fill":"#A5A5A5","geometryPath":"35 commands {\"type\":\"move\",\"x\":121,\"y\":68.749} .. {\"type\":\"close\"} #2801da1956a9d76f","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":82.75,"y":58.5}
{"fill":"#4472C4","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":112.05}
{"baselineY":117.1,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":287}
{"fill":"#ED7D31","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":127.35}
{"baselineY":132.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":287}
{"fill":"#A5A5A5","h":5.3,"kind":"rect","w":5.3,"x":278.5,"y":142.65}
{"baselineY":147.7,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q3","width":10,"x":287}
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
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.5}
{"fill":"#4472C4","h":68,"kind":"rect","w":32.4,"x":116.3,"y":118}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.5}
{"fill":"#4472C4","h":102,"kind":"rect","w":32.4,"x":197.3,"y":84}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"100","width":34,"x":258}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"80","width":34,"x":258}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"60","width":34,"x":258}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"40","width":34,"x":258}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":258}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":258}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":254,"x2":254,"y1":50,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":127.5}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":208.5}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":132.5,\"y\":126.933} .. {\"type\":\"close\"} #de7ef20137c9c61e","h":9.333,"kind":"shape","w":9.333,"x":127.833,"y":126.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":132.5,"x2":213.5,"y1":131.6,"y2":77.2}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":213.5,\"y\":72.533} .. {\"type\":\"close\"} #a604241e63ee9dff","h":9.333,"kind":"shape","w":9.333,"x":208.833,"y":72.533}
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
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":137}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":142,\"y\":181.333} .. {\"type\":\"close\"} #94d64a66647acc7a","h":9.333,"kind":"shape","w":9.333,"x":137.333,"y":181.333}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":142,"x2":242,"y1":186,"y2":63.647}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":242,\"y\":58.98} .. {\"type\":\"close\"} #9bdb41c7ad684639","h":9.333,"kind":"shape","w":9.333,"x":237.333,"y":58.98}
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
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":237}
{"fill":"#4472C4","h":27.2,"kind":"rect","w":40,"x":222,"y":50}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":137}
{"fill":"#4472C4","h":108.8,"kind":"rect","w":40,"x":122,"y":50}
# marker-circle
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":123.667,\"y\":95.6} .. {\"type\":\"close\"} #383ac9dc7798b52b","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":173.667,\"y\":65.2} .. {\"type\":\"close\"} #d958234292e8028b","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-diamond
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":117,\"y\":88.933} .. {\"type\":\"close\"} #ce9f221b8e081c27","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":167,\"y\":58.533} .. {\"type\":\"close\"} #8a482f65e80ec929","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-triangle
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"4 commands {\"type\":\"move\",\"x\":117,\"y\":88.933} .. {\"type\":\"close\"} #539b4c54afc0d03b","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"4 commands {\"type\":\"move\",\"x\":167,\"y\":58.533} .. {\"type\":\"close\"} #c86b95f6902ed151","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-square
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","h":13.333,"kind":"rect","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","h":13.333,"kind":"rect","w":13.333,"x":160.333,"y":58.533}
# marker-star
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":117,\"y\":88.933} .. {\"type\":\"close\"} #5248c28056ff51d5","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"11 commands {\"type\":\"move\",\"x\":167,\"y\":58.533} .. {\"type\":\"close\"} #d45cc1e8b99d4c3d","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-plus
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":115,\"y\":88.933} .. {\"type\":\"close\"} #bdba3fad83212cc2","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":165,\"y\":58.533} .. {\"type\":\"close\"} #61e5740cf1e6ab2e","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-dash
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":110.333,\"y\":93.933} .. {\"type\":\"close\"} #69e783a8f93c6b98","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":160.333,\"y\":63.533} .. {\"type\":\"close\"} #0b301dbf9cf220fa","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-dot
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":120.333,\"y\":95.6} .. {\"type\":\"close\"} #2eda7b1aa0a6aa88","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"25 commands {\"type\":\"move\",\"x\":170.333,\"y\":65.2} .. {\"type\":\"close\"} #26f68d4ee64a69fb","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-x
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":120.3,\"y\":89.472} .. {\"type\":\"close\"} #a79f8ae59c6ec593","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"13 commands {\"type\":\"move\",\"x\":170.3,\"y\":59.072} .. {\"type\":\"close\"} #6e7bc52fcf18c855","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-auto
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":117,\"y\":88.933} .. {\"type\":\"close\"} #ce9f221b8e081c27","h":13.333,"kind":"shape","w":13.333,"x":110.333,"y":88.933}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":167,\"y\":58.533} .. {\"type\":\"close\"} #8a482f65e80ec929","h":13.333,"kind":"shape","w":13.333,"x":160.333,"y":58.533}
# marker-none
attrs {"blockId":42,"chart":{"label":"Untitled chart, line chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":120,"kind":"rect","w":160,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":65.2,"y2":65.2}
{"baselineY":68.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"2","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":80.4,"y2":80.4}
{"baselineY":83.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":95.6,"y2":95.6}
{"baselineY":98.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"1","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":110.8,"y2":110.8}
{"baselineY":113.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0.5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":129,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":126}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":192,"y1":126,"y2":126}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":112}
{"baselineY":140,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":162}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":117,"x2":167,"y1":95.6,"y2":65.2}
# data-labels-composed
attrs {"blockId":42,"chart":{"label":"Untitled chart, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":50,"y2":50}
{"baselineY":53,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":77.2,"y2":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":104.4,"y2":104.4}
{"baselineY":107.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":131.6,"y2":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":158.8,"y2":158.8}
{"baselineY":161.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":189,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":50,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":292,"y1":186,"y2":186}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":137}
{"fill":"#4472C4","h":54.4,"kind":"rect","w":40,"x":122,"y":131.6}
{"baselineY":134.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North / Q1 / 10.0","width":85,"x":99.5}
{"baselineY":200,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":237}
{"fill":"#4472C4","h":108.8,"kind":"rect","w":40,"x":222,"y":77.2}
{"baselineY":80.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"North / Q2 / 20.0","width":85,"x":199.5}
# data-labels-percent-and-key
attrs {"blockId":42,"chart":{"label":"Untitled chart, pie chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"39 commands {\"type\":\"move\",\"x\":180,\"y\":135} .. {\"type\":\"close\"} #82a0391e87e75dac","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":103.5,"y":58.5}
{"fill":"#4472C4","h":7,"kind":"rect","w":7,"x":232.208,"y":190.208}
{"baselineY":197.208,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"75%","width":48,"x":242.208}
{"fill":"#ED7D31","geometryPath":"15 commands {\"type\":\"move\",\"x\":180,\"y\":135} .. {\"type\":\"close\"} #b3db07e7d76a078f","h":153,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":153,"x":103.5,"y":58.5}
{"fill":"#ED7D31","h":7,"kind":"rect","w":7,"x":107.792,"y":65.792}
{"baselineY":72.792,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25%","width":48,"x":117.792}
# text-properties
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 1 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":180,"kind":"rect","w":260,"x":50,"y":40}
{"baselineY":58,"color":"#112233","font":"italic 600 28px Georgia","kind":"text","text":"Revenue","width":98,"x":131}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":244.8,"y1":68,"y2":68}
{"baselineY":71,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":244.8,"y1":97.5,"y2":97.5}
{"baselineY":100.5,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":244.8,"y1":127,"y2":127}
{"baselineY":130,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":244.8,"y1":156.5,"y2":156.5}
{"baselineY":159.5,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":244.8,"y1":186,"y2":186}
{"baselineY":189,"color":"#884400","font":"400 8px Georgia","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":186}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":244.8,"y1":186,"y2":186}
{"baselineY":200,"color":"#112233","font":"400 20px Georgia","kind":"text","text":"Q1","width":20,"x":120.2}
{"fill":"#4472C4","h":59,"kind":"rect","w":30.56,"x":114.92,"y":127}
{"baselineY":124,"color":"#112233","font":"700 16px Georgia","kind":"text","text":"10","width":16,"x":122.2}
{"baselineY":200,"color":"#112233","font":"400 20px Georgia","kind":"text","text":"Q2","width":20,"x":196.6}
{"fill":"#4472C4","h":118,"kind":"rect","w":30.56,"x":191.32,"y":68}
{"baselineY":65,"color":"#112233","font":"700 16px Georgia","kind":"text","text":"20","width":16,"x":198.6}
{"fill":"#4472C4","h":6.36,"kind":"rect","w":6.36,"x":256.8,"y":140.82}
{"baselineY":146.88,"color":"#112233","font":"700 12px Georgia","kind":"text","text":"North","width":30,"x":267}
# zero-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":72.8,"y2":72.8}
{"baselineY":75.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":77.6,"y2":77.6}
{"baselineY":80.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":82.4,"y2":82.4}
{"baselineY":85.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":87.2,"y2":87.2}
{"baselineY":90.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":93}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":3.429,"x":94.571,"y":82.4}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":3.429,"x":98,"y":88.16}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":105}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":3.429,"x":106.571,"y":72.8}
{"fill":"#4472C4","h":24,"kind":"rect","w":3.429,"x":110,"y":68}
# tiny-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, column chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":72.8,"y2":72.8}
{"baselineY":75.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":77.6,"y2":77.6}
{"baselineY":80.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":82.4,"y2":82.4}
{"baselineY":85.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":87.2,"y2":87.2}
{"baselineY":90.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":116,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":93}
{"fill":"#4472C4","h":9.6,"kind":"rect","w":3.429,"x":94.571,"y":82.4}
{"fill":"#4472C4","h":3.84,"kind":"rect","w":3.429,"x":98,"y":88.16}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":105}
{"fill":"#4472C4","h":19.2,"kind":"rect","w":3.429,"x":106.571,"y":72.8}
{"fill":"#4472C4","h":24,"kind":"rect","w":3.429,"x":110,"y":68}
# tiny-rect-pie
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, pie chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":8,"kind":"rect","w":12,"x":50,"y":40}
{"fill":"#4472C4","geometryPath":"19 commands {\"type\":\"move\",\"x\":50,\"y\":58} .. {\"type\":\"close\"} #22b60aabc3e35d88","h":20,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20,"x":40,"y":48}
{"fill":"#4472C4","geometryPath":"35 commands {\"type\":\"move\",\"x\":50,\"y\":58} .. {\"type\":\"close\"} #e6d6692b848640b1","h":20,"kind":"shape","stroke":{"color":"#FFFFFF","width":1},"w":20,"x":40,"y":48}
# wide-flat-rect
attrs {"ariaLabel":"Revenue","blockId":42,"chart":{"label":"Revenue, line chart, 2 series, 2 categories"},"docEnd":5,"docStart":4}
{"fill":"#FFFFFF","h":26,"kind":"rect","w":900,"x":50,"y":40}
{"baselineY":58,"color":"#222222","font":"600 13px Calibri, sans-serif","kind":"text","text":"Revenue","width":45.5,"x":477.25}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":927,"y1":68,"y2":68}
{"baselineY":71,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"25","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":927,"y1":72.8,"y2":72.8}
{"baselineY":75.8,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"20","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":927,"y1":77.6,"y2":77.6}
{"baselineY":80.6,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"15","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":927,"y1":82.4,"y2":82.4}
{"baselineY":85.4,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"10","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":927,"y1":87.2,"y2":87.2}
{"baselineY":90.2,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"5","width":34,"x":54}
{"color":"#D9D9D9","kind":"line","strokeWidth":0.5,"x1":92,"x2":927,"y1":92,"y2":92}
{"baselineY":95,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"0","width":34,"x":54}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":92,"y1":68,"y2":92}
{"color":"#666666","kind":"line","strokeWidth":1,"x1":92,"x2":927,"y1":92,"y2":92}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q1","width":10,"x":295.75}
{"baselineY":106,"color":"#222222","font":"400 10px Calibri, sans-serif","kind":"text","text":"Q2","width":10,"x":713.25}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":300.75,\"y\":77.733} .. {\"type\":\"close\"} #bc061f10d6a7a001","h":9.333,"kind":"shape","w":9.333,"x":296.083,"y":77.733}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":300.75,"x2":718.25,"y1":82.4,"y2":72.8}
{"fill":"#4472C4","geometryPath":"5 commands {\"type\":\"move\",\"x\":718.25,\"y\":68.133} .. {\"type\":\"close\"} #da152c28231fcf75","h":9.333,"kind":"shape","w":9.333,"x":713.583,"y":68.133}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":296.083,"y":83.493}
{"color":"#4472C4","kind":"line","strokeWidth":2,"x1":300.75,"x2":718.25,"y1":88.16,"y2":68}
{"fill":"#4472C4","h":9.333,"kind":"rect","w":9.333,"x":713.583,"y":63.333}
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
