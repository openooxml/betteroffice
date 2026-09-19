//! Geometry measured from Word 16.112.4 exporting a landscape A4 probe whose
//! full-width `vertAnchor="page"` table sits at `w:tblpY="2431"`.

use serde_json::{Value, json};

const PAGE_W: f64 = 1122.5333333333333;
const PAGE_H: f64 = 793.7333333333333;
const MARGIN: f64 = 96.0;
const CONTENT_W: f64 = PAGE_W - 2.0 * MARGIN;
/// `w:tblpY="2431"` twips.
const BAND_TOP: f64 = 162.06666666666666;
/// One `w:trHeight="1400"` row.
const FLOAT_HEIGHT: f64 = 93.33333333333333;
/// Word restarts the flow here: 191.55pt at 96 DPI.
const BAND_BOTTOM: f64 = 255.4;

fn paragraph(id: &str, height: f64) -> Value {
    json!({
        "block":{"kind":"paragraph","id":id,"runs":[{"kind":"text","text":id}]},
        "measure":{"kind":"paragraph","totalHeight":height,"lines":[{
            "headRun":0,"headChar":0,"tailRun":0,"tailChar":id.len(),
            "width":20,"ascent":8,"descent":2,"lineHeight":height
        }]}
    })
}

fn table(id: &str, rows: usize, row_height: f64) -> Value {
    let content = paragraph(id, row_height);
    let blocks: Vec<_> = (0..rows)
        .map(|index| json!({"id":index,"cells":[{"id":index,"blocks":[content["block"]]}]}))
        .collect();
    let measured: Vec<_> = (0..rows)
        .map(|_| {
            json!({"height":row_height,"cells":[{
                "width":CONTENT_W,"height":row_height,"blocks":[content["measure"]]
            }]})
        })
        .collect();
    json!({
        "block":{"kind":"table","id":id,"rows":blocks},
        "measure":{"kind":"table","rows":measured,"columnWidths":[CONTENT_W],
            "totalWidth":CONTENT_W,"totalHeight":row_height * rows as f64}
    })
}

fn floating(id: &str, rows: usize) -> Value {
    let mut result = table(id, rows, FLOAT_HEIGHT);
    result["block"]["floating"] = json!({
        "horzAnchor":"margin","vertAnchor":"page","tblpY":BAND_TOP,
        "leftFromText":12,"rightFromText":12
    });
    result
}

fn layout(blocks: Vec<Value>) -> Value {
    let input = json!({
        "measured":blocks,
        "options":{"pageSize":{"w":PAGE_W,"h":PAGE_H},
            "margins":{"top":MARGIN,"right":MARGIN,"bottom":MARGIN,"left":MARGIN}}
    });
    serde_json::from_str(&docx_layout::layout_to_canonical_json(&input.to_string()).unwrap())
        .unwrap()
}

fn fragment<'a>(output: &'a Value, page: usize, block_id: &str) -> &'a Value {
    output["pages"][page]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .find(|fragment| fragment["blockId"] == block_id)
        .unwrap()
}

#[test]
fn flow_restarts_below_a_page_anchored_band_instead_of_under_it() {
    let output = layout(vec![
        paragraph("prefix", 370.0),
        table("inline", 3, 120.0),
        floating("float", 1),
        paragraph("after", 10.0),
    ]);
    assert_eq!(output["pages"].as_array().unwrap().len(), 2);
    assert_eq!(fragment(&output, 1, "float")["y"], 162.067);
    assert_eq!(fragment(&output, 1, "inline")["y"], BAND_BOTTOM);
    assert_eq!(fragment(&output, 1, "inline")["height"], 240.0);
    assert_eq!(fragment(&output, 1, "after")["y"], BAND_BOTTOM + 240.0);
}

#[test]
fn flow_clearing_the_band_keeps_its_place() {
    let output = layout(vec![
        paragraph("prefix", 370.0),
        table("inline", 3, 60.0),
        floating("float", 1),
        paragraph("after", 10.0),
    ]);
    assert_eq!(fragment(&output, 0, "inline")["y"], 466.0);
    assert_eq!(fragment(&output, 0, "float")["y"], 162.067);
}

#[test]
fn a_band_that_cannot_hold_the_shifted_flow_keeps_the_existing_path() {
    let output = layout(vec![
        paragraph("prefix", 370.0),
        table("inline", 3, 120.0),
        floating("float", 5),
        paragraph("after", 10.0),
    ]);
    assert_eq!(fragment(&output, 1, "inline")["y"], MARGIN);
}

#[test]
fn a_band_meeting_only_a_later_row_keeps_the_existing_path() {
    let mut float = floating("float", 1);
    // w:tblpY="5000" twips.
    float["block"]["floating"]["tblpY"] = json!(333.3333333333333);
    let output = layout(vec![
        paragraph("prefix", 370.0),
        table("inline", 3, 120.0),
        float,
        paragraph("after", 10.0),
    ]);
    assert_eq!(output["pages"].as_array().unwrap().len(), 2);
    assert_eq!(fragment(&output, 1, "inline")["y"], MARGIN);
}
