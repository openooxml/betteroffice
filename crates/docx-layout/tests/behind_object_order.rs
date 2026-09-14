use serde_json::{Value, json};

fn input() -> Value {
    json!({
        "measured":[
            {"block":{"kind":"paragraph","id":"paragraph","runs":[
                {"kind":"image","src":"behind","width":50,"height":50,"wrapType":"behind",
                    "position":{"relativeHeight":20}},
                {"kind":"text","text":"BODY"},
                {"kind":"image","src":"front","width":50,"height":50,"wrapType":"inFront",
                    "position":{"relativeHeight":0}}
            ]},"measure":{"kind":"paragraph","totalHeight":20,"lines":[{
                "headRun":1,"headChar":0,"tailRun":1,"tailChar":4,
                "width":40,"ascent":15,"descent":5,"lineHeight":20
            }]}},
            {"block":{"kind":"shape","id":"parent","title":"PARENT","width":80,"height":80,
                "position":{},"behindDoc":true,"relativeHeight":10,
                "fill":{"kind":"solid","color":"FF0000"},
                "children":[{"id":"child","title":"CHILD","width":40,"height":40,
                    "relativeHeight":100,"fill":{"kind":"solid","color":"00FF00"},
                    "innerText":[{"id":"child-text","runs":[{"kind":"text","text":"CHILD TEXT"}]}],
                    "innerMeasures":[{"totalHeight":20,"lines":[{
                        "headRun":0,"headChar":0,"tailRun":0,"tailChar":10,
                        "width":40,"ascent":15,"descent":5,"lineHeight":20
                    }]}]
                }]},"measure":{"kind":"shape","width":80,"height":80}}
        ],
        "options":{},"layout":{"pages":[{"size":{"w":200,"h":200},"margins":{},
            "fragments":[
                {"kind":"paragraph","blockId":"paragraph","x":0,"y":0,"width":200,"height":20,"fromLine":0,"toLine":1},
                {"kind":"shape","blockId":"parent","x":0,"y":0,"width":80,"height":80}
            ]}]}
    })
}

fn primitives(input: &Value) -> Vec<Value> {
    let display: Value = serde_json::from_str(
        &docx_layout::display_list::build_display_list_json(&input.to_string()).unwrap(),
    )
    .unwrap();
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .clone()
}

fn paint_order(primitives: &[Value]) -> Vec<&str> {
    primitives
        .iter()
        .filter_map(|primitive| {
            primitive["ariaLabel"]
                .as_str()
                .or_else(|| primitive["relId"].as_str())
                .or_else(|| primitive["text"].as_str())
        })
        .collect()
}

#[test]
fn behind_shape_children_remain_a_contiguous_object_group() {
    let output = primitives(&input());
    assert_eq!(
        paint_order(&output),
        ["PARENT", "CHILD", "CHILD TEXT", "behind", "BODY", "front"]
    );
    let mut reversed = input();
    reversed["measured"][1]["block"]["relativeHeight"] = json!(30);
    let reverse_output = primitives(&reversed);
    assert_eq!(
        paint_order(&reverse_output),
        ["behind", "PARENT", "CHILD", "CHILD TEXT", "BODY", "front"]
    );
    let group = |output: &[Value]| {
        output
            .iter()
            .filter(|primitive| {
                primitive["ariaLabel"] == "CHILD" || primitive["text"] == "CHILD TEXT"
            })
            .cloned()
            .collect::<Vec<_>>()
    };
    assert_eq!(group(&output), group(&reverse_output));
}

#[test]
fn inline_images_keep_their_line_placement_and_paint_order() {
    let mut input = input();
    input["measured"][0]["block"]["runs"][0]["wrapType"] = json!("inline");
    input["measured"][0]["measure"]["lines"][0]["headRun"] = json!(0);
    let output = primitives(&input);
    assert_eq!(
        paint_order(&output),
        ["PARENT", "CHILD", "CHILD TEXT", "behind", "BODY", "front"]
    );
    let image = output
        .iter()
        .find(|primitive| primitive["relId"] == "behind")
        .unwrap();
    assert_eq!(image["x"], 0);
    assert_eq!(image["y"], -35);
}

#[test]
fn body_table_clip_groups_keep_their_original_order_and_geometry() {
    let mut table: Value = serde_json::from_str(include_str!(
        "fixtures/table-splits-with-repeated-header.input.json"
    ))
    .unwrap();
    table["layout"] = serde_json::from_str(include_str!(
        "fixtures/table-splits-with-repeated-header.golden.json"
    ))
    .unwrap();
    let original = primitives(&table);
    let behind = input();
    table["measured"]
        .as_array_mut()
        .unwrap()
        .extend(behind["measured"].as_array().unwrap().iter().cloned());
    table["layout"]["pages"][0]["fragments"]
        .as_array_mut()
        .unwrap()
        .extend(
            behind["layout"]["pages"][0]["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .cloned(),
        );
    let reordered = primitives(&table);
    let clipped = |output: &[Value]| {
        output
            .iter()
            .filter(|primitive| !primitive["clipGroup"].is_null())
            .cloned()
            .collect::<Vec<_>>()
    };
    let expected = clipped(&original);
    assert!(!expected.is_empty());
    assert_eq!(clipped(&reordered), expected);
    let positions: Vec<_> = reordered
        .iter()
        .enumerate()
        .filter_map(|(index, primitive)| (!primitive["clipGroup"].is_null()).then_some(index))
        .collect();
    assert!(positions.windows(2).all(|pair| pair[1] == pair[0] + 1));
}
