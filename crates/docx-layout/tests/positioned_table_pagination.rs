use serde_json::{Value, json};

fn paragraph(id: &str, height: u32) -> Value {
    json!({
        "block":{"kind":"paragraph","id":id,"runs":[{"kind":"text","text":id}]},
        "measure":{"kind":"paragraph","totalHeight":height,"lines":[{
            "headRun":0,"headChar":0,"tailRun":0,"tailChar":id.len(),
            "width":20,"ascent":8,"descent":2,"lineHeight":height
        }]}
    })
}

fn table(width: u32, offset: i32) -> Value {
    let mut rows = Vec::new();
    let mut measured = Vec::new();
    for index in 0..3 {
        let content = paragraph(&format!("row-{index}"), 20);
        rows.push(json!({"id":index,"cantSplit":true,"cells":[{
            "id":index,"blocks":[content["block"]]
        }]}));
        measured.push(json!({"height":20,"cells":[{
            "width":width,"height":20,"blocks":[content["measure"]]
        }]}));
    }
    json!({
        "block":{"kind":"table","id":"floating","rows":rows,"columnWidths":[width],
            "floating":{"horzAnchor":"margin","vertAnchor":"text","tblpX":3,
                "tblpY":offset,"leftFromText":10,"rightFromText":10,"bottomFromText":5}},
        "measure":{"kind":"table","rows":measured,"columnWidths":[width],
            "totalWidth":width,"totalHeight":60}
    })
}

fn layout(prefix: u32, table: Value) -> Value {
    let input = json!({
        "measured":[paragraph("before",prefix),table,paragraph("after",10)],
        "options":{"pageSize":{"w":200,"h":120},
            "margins":{"top":10,"right":10,"bottom":10,"left":10}}
    });
    serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
        .unwrap()
}

#[test]
fn full_width_text_floats_fragment_at_remaining_page_capacity() {
    for offset in [0, 10] {
        let output = layout(50, table(180, offset));
        assert_eq!(output["pages"].as_array().unwrap().len(), 2);
        let first = &output["pages"][0]["fragments"][1];
        let next = &output["pages"][1]["fragments"][0];
        assert_eq!(first["rowStart"], 0);
        assert_eq!(first["rowEnd"], 2);
        assert_eq!(first["y"], 60 + offset);
        assert_eq!(first["height"], 40);
        assert_eq!(first["carriedToNext"], true);
        assert_eq!(next["rowStart"], 2);
        assert_eq!(next["rowEnd"], 3);
        assert_eq!(next["y"], 10);
        assert_eq!(next["height"], 20);
        assert_eq!(next["carriedFromPrev"], true);
        assert_eq!(first["x"], 13);
        assert_eq!(next["x"], 13);
        assert_eq!(first["isFloating"], true);
        assert_eq!(next["isFloating"], true);
        assert_eq!(output["pages"][1]["fragments"][1]["y"], 35);
    }
}

#[test]
fn fitting_full_width_floats_retain_their_anchor() {
    let output = layout(10, table(180, 10));
    assert_eq!(output["pages"].as_array().unwrap().len(), 1);
    let fragment = &output["pages"][0]["fragments"][1];
    assert_eq!(fragment["y"], 30);
    assert_eq!(fragment["x"], 13);
    assert_eq!(fragment["rowEnd"], 3);
    assert!(fragment["carriedToNext"].is_null());
    assert_eq!(output["pages"][0]["fragments"][2]["y"], 95);
}

#[test]
fn narrow_and_page_relative_floats_keep_their_existing_placement() {
    let narrow = layout(50, table(60, 10));
    assert_eq!(narrow["pages"].as_array().unwrap().len(), 1);
    assert_eq!(narrow["pages"][0]["fragments"][2]["y"], 60);
    let mut table = table(180, 10);
    table["block"]["floating"]["vertAnchor"] = json!("page");
    let page_relative = layout(50, table);
    assert_eq!(page_relative["pages"][0]["fragments"][1]["y"], 10);
    assert_eq!(page_relative["pages"][0]["fragments"][1]["rowEnd"], 3);
}

#[test]
fn parity_aligned_floats_keep_their_existing_placement() {
    for (alignment, x) in [("inside", 10), ("outside", 30)] {
        let mut table = table(160, 10);
        table["block"]["floating"]
            .as_object_mut()
            .unwrap()
            .remove("tblpX");
        table["block"]["floating"]["tblpXSpec"] = json!(alignment);
        let output = layout(50, table);
        let fragments: Vec<_> = output["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["fragments"].as_array().unwrap())
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        assert_eq!(fragments.len(), 1);
        assert_eq!(fragments[0]["x"], x);
        assert_eq!(fragments[0]["rowEnd"], 3);
        assert!(fragments[0]["carriedToNext"].is_null());
    }
}
