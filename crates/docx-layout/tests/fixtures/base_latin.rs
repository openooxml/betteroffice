use docx_layout::display_list::{DisplayList, build_display_list_value_with_fonts};
use ooxml_text::FontStore;
use serde_json::{Value, json};

pub fn measured_document(fonts: &FontStore, ids: &[u32], text: &str, bold: bool) -> DisplayList {
    let block = json!({
        "kind": "paragraph", "id": 0,
        "runs": [{
            "kind": "text", "text": text, "fontFamily": "Missing",
            "fontSize": 11, "bold": bold, "pmStart": 1, "pmEnd": 1 + text.encode_utf16().count()
        }]
    });
    let key = format!("missing|{}|0", u8::from(bold));
    let chains = json!({(key): ids});
    let request = json!({
        "block": block, "maxWidth": 340,
        "fontChains": chains, "authoritativeShaping": true,
        "defaults": {"fontFamily": "Missing", "fontSize": 11}
    });
    let extent: Value = serde_json::from_str(
        &ooxml_text::measure::measure_paragraph_json(fonts, &request.to_string()).unwrap(),
    )
    .unwrap();
    let mut input = json!({
        "measured": [{"block": block, "measure": extent}],
        "options": {
            "pageSize": {"w": 360, "h": 200},
            "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}, "pageGap": 0
        },
        "fontChains": chains
    });
    input["layout"] =
        serde_json::from_str(&docx_layout::layout_to_json(&input.to_string()).unwrap()).unwrap();
    build_display_list_value_with_fonts(&input.to_string(), fonts).unwrap()
}
