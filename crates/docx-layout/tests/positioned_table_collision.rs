use serde_json::{Value, json};

fn paragraph(id: &str, height: f64) -> Value {
    json!({
        "block":{"kind":"paragraph","id":id,"runs":[{"kind":"text","text":id}]},
        "measure":{"kind":"paragraph","totalHeight":height,"lines":[{
            "headRun":0,"headChar":0,"tailRun":0,"tailChar":id.len(),
            "width":20,"ascent":8,"descent":2,"lineHeight":height
        }]}
    })
}

fn table(id: &str, height: f64) -> Value {
    let content = paragraph(id, height / 4.0);
    let rows: Vec<_> = (0..4)
        .map(|index| {
            json!({"id":index,"cantSplit":true,"cells":[{
                "id":index,"blocks":[content["block"]]
            }]})
        })
        .collect();
    let measured: Vec<_> = (0..4)
        .map(|_| {
            json!({"height":height/4.0,"cells":[{
                "width":944.9333333333333,"height":height/4.0,
                "blocks":[content["measure"]]
            }]})
        })
        .collect();
    json!({
        "block":{"kind":"table","id":id,"rows":rows},
        "measure":{"kind":"table","rows":measured,
            "columnWidths":[944.9333333333333],
            "totalWidth":944.9333333333333,"totalHeight":height}
    })
}

fn floating_table() -> Value {
    let mut result = table("floating", 392.311456);
    result["block"]["floating"] = json!({
        "horzAnchor":"margin","vertAnchor":"page","tblpX":17,
        "tblpY":162.06666666666666,"leftFromText":12,"rightFromText":12
    });
    result
}

fn layout(prefix: f64, prior: Value, floating: Value) -> Value {
    let input = json!({
        "measured":[paragraph("prefix",prefix),prior,paragraph("anchor",24.68457),
            floating,paragraph("after",10.0)],
        "options":{"pageSize":{"w":1122.5333333333333,"h":793.7333333333333},
            "margins":{"top":151.83592,"right":96,"bottom":96,"left":96}}
    });
    serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
        .unwrap()
}

fn floating_fragment(output: &Value) -> (usize, &Value) {
    output["pages"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .find_map(|(index, page)| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .find(|fragment| fragment["blockId"] == "floating")
                .map(|fragment| (index, fragment))
        })
        .unwrap()
}

#[test]
fn fixed_page_table_avoids_an_inline_table_that_cannot_reflow_below_it() {
    let output = layout(
        74.05371,
        table("inline", 399.1114523333333),
        floating_table(),
    );
    assert_eq!(output["pages"].as_array().unwrap().len(), 2);
    let (page, floating) = floating_fragment(&output);
    assert_eq!(page, 1);
    assert_eq!(floating["x"], 113);
    assert_eq!(floating["y"], 162.067);
    assert_eq!(floating["height"], 392.311);
    assert_eq!(floating["rowStart"], 0);
    assert_eq!(floating["rowEnd"], 4);
    assert_eq!(output["pages"][0]["fragments"][1]["blockId"], "inline");
    assert_eq!(output["pages"][0]["fragments"][1]["y"], 225.89);
    assert_eq!(output["pages"][0]["fragments"][2]["blockId"], "anchor");
}

#[test]
fn fitting_inline_reflow_is_not_replaced_with_a_page_advance() {
    let output = layout(74.05371, table("inline", 80.0), floating_table());
    assert_eq!(floating_fragment(&output).0, 0);
}

#[test]
fn nonintersecting_fixed_page_tables_stay_on_their_anchor_page() {
    let mut beside = floating_table();
    beside["block"]["floating"]["tblpX"] = json!(960);
    let output = layout(74.05371, table("inline", 399.1114523333333), beside);
    assert_eq!(floating_fragment(&output).0, 0);

    let mut below = table("floating", 150.0);
    below["block"]["floating"] = floating_table()["block"]["floating"].clone();
    below["block"]["floating"]["tblpY"] = json!(500);
    let output = layout(200.0, table("inline", 100.0), below);
    assert_eq!(floating_fragment(&output).0, 0);
}

#[test]
fn narrow_page_floats_keep_their_existing_reflow_path() {
    let mut floating = floating_table();
    floating["measure"]["totalWidth"] = json!(200);
    floating["measure"]["columnWidths"] = json!([200]);
    for row in floating["measure"]["rows"].as_array_mut().unwrap() {
        row["cells"][0]["width"] = json!(200);
    }
    let output = layout(74.05371, table("inline", 399.1114523333333), floating);
    assert_eq!(floating_fragment(&output).0, 0);
}

#[test]
fn carried_inline_fragments_keep_their_existing_reflow_path() {
    let output = layout(
        74.05371,
        table("inline", 798.2229046666666),
        floating_table(),
    );
    let (page, _) = floating_fragment(&output);
    assert_eq!(page, 1);
    assert_eq!(
        output["pages"][page]["fragments"][0]["carriedFromPrev"],
        true
    );
}

#[test]
fn floating_predecessors_do_not_trigger_inline_collision_handling() {
    let mut prior = table("inline", 399.1114523333333);
    prior["block"]["floating"] = json!({
        "horzAnchor":"margin","vertAnchor":"page","tblpX":17,"tblpY":225.88963
    });
    let output = layout(74.05371, prior, floating_table());
    assert_eq!(floating_fragment(&output).0, 0);
}
