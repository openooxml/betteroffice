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
fn overflowing_narrow_floats_move_and_page_relative_floats_keep_their_placement() {
    let narrow = layout(50, table(60, 10));
    assert_eq!(narrow["pages"].as_array().unwrap().len(), 2);
    assert_eq!(narrow["pages"][0]["fragments"].as_array().unwrap().len(), 1);
    assert_eq!(narrow["pages"][1]["fragments"][0]["y"], 20);
    assert_eq!(narrow["pages"][1]["fragments"][0]["height"], 60);
    assert_eq!(narrow["pages"][1]["fragments"][1]["y"], 10);
    let mut table = table(180, 10);
    table["block"]["floating"]["vertAnchor"] = json!("page");
    let page_relative = layout(50, table);
    assert_eq!(page_relative["pages"][0]["fragments"][1]["y"], 10);
    assert_eq!(page_relative["pages"][0]["fragments"][1]["rowEnd"], 3);
}

/// Word moves a side-wrapping text-relative table with its anchor paragraph
/// when its rows cross the body bottom and a fresh body can hold the table.
#[test]
fn a_narrow_float_moves_with_its_measured_multiline_anchor() {
    use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks_with_floats};
    use docx_layout::types::{BlockExtent, Fragment, Input, LayoutBlock, MeasuredBlock};

    let mut blocks: Vec<LayoutBlock> = serde_json::from_value(json!([
        {"kind":"table","id":"floating","columnWidths":[60],"layoutMode":"fixed",
            "rows":[
                {"id":0,"height":20,"heightRule":"exact","cells":[]},
                {"id":1,"height":20,"heightRule":"exact","cells":[]}
            ],
            "floating":{"horzAnchor":"text","vertAnchor":"text","tblpXSpec":"right",
                "tblpY":5,"leftFromText":10,"rightFromText":10}},
        {"kind":"paragraph","id":"anchor",
            "attrs":{"spacing":{"line":10,"lineRule":"exact"}},
            "runs":[{"kind":"text",
                "text":"alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu"}]}
    ]))
    .unwrap();
    let font = docx_layout::register_measure_font(include_bytes!(
        "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
    ))
    .unwrap();
    let config = MeasurementConfig {
        font_chains: std::collections::BTreeMap::from([(
            "liberation sans|0|0".to_owned(),
            vec![font],
        )]),
        defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
        ..MeasurementConfig::default()
    };
    let extents = measure_blocks_with_floats(&mut blocks, &[180.0; 2], &config, None).unwrap();
    let BlockExtent::Paragraph(anchor_measure) = &extents[1] else {
        panic!("anchor paragraph expected");
    };
    assert!(anchor_measure.lines.len() > 1);
    assert_eq!(anchor_measure.lines[0].right_offset, Some(70.0));
    assert!(anchor_measure.lines[0].width <= 110.0);
    let mut input: Input = serde_json::from_value(json!({
        "measured":[paragraph("filler-1",35),paragraph("filler-2",35)],
        "options":{"pageSize":{"w":200,"h":120},
            "margins":{"top":10,"right":10,"bottom":10,"left":10}}
    }))
    .unwrap();
    input.measured.extend(
        blocks
            .into_iter()
            .zip(extents)
            .map(|(block, measure)| MeasuredBlock { block, measure }),
    );
    let output = docx_layout::compute_layout_input(&mut input).unwrap();
    assert_eq!(output.pages.len(), 2);
    assert!(
        output.pages[0]
            .fragments
            .iter()
            .all(|fragment| !matches!(fragment, Fragment::Table(_)))
    );
    let [Fragment::Table(table), Fragment::Paragraph(anchor)] =
        output.pages[1].fragments.as_slice()
    else {
        panic!("table and anchor expected on page 2");
    };
    assert_eq!(
        (table.x, table.y, table.row_start, table.row_end),
        (130.0, 15.0, 0, 2)
    );
    assert_eq!((anchor.y, anchor.from_line), (10.0, 0));
    assert_eq!(table.y - anchor.y, 5.0);
    assert!(table.y + table.height <= output.pages[1].size.h - output.pages[1].margins.bottom);
}

#[test]
fn narrow_floats_that_cannot_fit_at_the_body_top_do_not_retry() {
    for (prefix, offset, page_index, y) in [(0, 50, 0, 60), (50, 80, 1, 90)] {
        let output = layout(prefix, table(60, offset));
        assert_eq!(output["pages"].as_array().unwrap().len(), page_index + 1);
        let fragment = output["pages"][page_index]["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .find(|fragment| fragment["kind"] == "table")
            .unwrap();
        assert_eq!(fragment["y"], y);
        assert_eq!(fragment["rowEnd"], 3);
        assert!(fragment["carriedToNext"].is_null());
        assert!(fragment["y"].as_f64().unwrap() + fragment["height"].as_f64().unwrap() > 110.0);
    }
}

#[test]
fn a_narrow_float_stays_when_the_next_body_cannot_hold_its_rows() {
    let input = json!({
        "measured":[paragraph("before",50),table(60,10),paragraph("after",10)],
        "options":{"pageSize":{"w":200,"h":120},
            "margins":{"top":10,"right":10,"bottom":10,"left":10},
            "footnoteReservedHeights":{"2":45}}
    });
    let output: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    assert_eq!(output["pages"].as_array().unwrap().len(), 1);
    assert_eq!(output["pages"][0]["fragments"][1]["y"], 70);
    assert_eq!(output["pages"][0]["fragments"][1]["rowEnd"], 3);
    assert_eq!(output["pages"][0]["fragments"][2]["y"], 60);
}

#[test]
fn an_overflowing_narrow_float_moves_to_the_next_column() {
    let mut floating = table(40, 10);
    floating["block"]["floating"]["horzAnchor"] = json!("text");
    let input = json!({
        "measured":[paragraph("before",50),floating,paragraph("after",10)],
        "options":{"pageSize":{"w":200,"h":120},
            "margins":{"top":10,"right":10,"bottom":10,"left":10},
            "columns":{"count":2,"gap":20}}
    });
    let output: Value =
        serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
            .unwrap();
    assert_eq!(output["pages"].as_array().unwrap().len(), 1);
    assert_eq!(output["pages"][0]["fragments"][1]["x"], 113);
    assert_eq!(output["pages"][0]["fragments"][1]["y"], 20);
    assert_eq!(output["pages"][0]["fragments"][2]["x"], 110);
    assert_eq!(output["pages"][0]["fragments"][2]["y"], 10);
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

/// A band opening above the pen still costs the page its height, and the flow
/// already emitted into it stays put — Word instead moves that flow below the
/// band, which a single forward pass cannot do.
#[test]
fn a_page_anchored_band_above_the_pen_costs_the_page_its_height() {
    let mut table = table(180, 10);
    table["block"]["floating"]["vertAnchor"] = json!("page");
    let output = layout(50, table);
    let placed = |id: &str| {
        output["pages"]
            .as_array()
            .unwrap()
            .iter()
            .enumerate()
            .flat_map(|(index, page)| {
                page["fragments"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(move |fragment| (index, fragment))
            })
            .find(|(_, fragment)| fragment["blockId"] == id)
            .map(|(index, fragment)| (index, fragment["y"].as_f64().unwrap()))
            .unwrap()
    };
    let band = &output["pages"][0]["fragments"][1];
    assert_eq!(band["y"], 10);
    assert_eq!(band["height"], 60);
    assert_eq!(placed("after"), (1, 10.0));
    // Residue: `before` occupies 10..60, inside the 10..70 band.
    assert_eq!(placed("before"), (0, 10.0));
}
