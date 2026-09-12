use docx_layout::display_list::build_display_list_json;
use serde_json::{Value, json};

fn primitives(chart: Value) -> Vec<Value> {
    let input = json!({
        "measured": [{
            "block": {
                "kind": "chart",
                "id": 42,
                "width": 260.0,
                "height": 180.0,
                "docStart": 4,
                "docEnd": 5,
                "chart": chart
            },
            "measure": { "kind": "chart", "width": 260.0, "height": 180.0 }
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
                "width": 260.0,
                "height": 180.0,
                "docStart": 4,
                "docEnd": 5
            }]
        }] }
    });
    let json = build_display_list_json(&input.to_string()).expect("display list builds");
    let list: Value = serde_json::from_str(&json).expect("display list is json");
    list["pages"][0]["primitives"]
        .as_array()
        .expect("primitives")
        .clone()
}

fn chart(title_text: Value) -> Value {
    json!({
        "type": "chart",
        "chartType": "column",
        "title": "Revenue",
        "titleText": title_text,
        "legend": { "position": "right", "visible": true },
        "series": [{
            "name": "North",
            "categories": ["Q1", "Q2"],
            "values": [10.0, 20.0],
            "color": "#4472C4"
        }],
        "axes": { "value": { "min": 0.0, "max": 25.0 } }
    })
}

fn title(primitives: &[Value]) -> Value {
    primitives
        .iter()
        .find(|primitive| primitive["text"] == json!("Revenue"))
        .expect("title primitive")
        .clone()
}

#[test]
fn a_charts_tracking_reaches_the_text_primitive_it_widened() {
    let plain = title(&primitives(chart(json!({ "sizePt": 12.0 }))));
    assert_eq!(plain["letterSpacing"], Value::Null);

    let tracked = title(&primitives(chart(
        json!({ "sizePt": 12.0, "spacingPt": 3.0 }),
    )));
    assert_eq!(tracked["letterSpacing"], json!(4));
    assert_eq!(tracked["font"], plain["font"]);
}
