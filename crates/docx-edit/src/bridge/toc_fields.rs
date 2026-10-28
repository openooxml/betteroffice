use serde_json::{Value, json};

use crate::{EditingDoc, SegmentContent};

const END_MARKER: &str = "\u{fdd0}\u{fdd1}";

pub(super) fn live_end(paragraph: &Value, live_content: &str) -> Option<u32> {
    let mut paragraph = paragraph.clone();
    let mut depth = 1;
    if !mark_end(&mut paragraph, &mut depth) || live_content.contains(END_MARKER) {
        return None;
    }
    let scratch = EditingDoc::new(0);
    crate::seed::seed_blocks(&scratch, None, &[("body".to_owned(), &[paragraph])]).ok()?;
    let mut projected = String::new();
    for segment in scratch.story_segments("body").ok()? {
        match segment.content {
            SegmentContent::Text(text) => projected.push_str(&text),
            SegmentContent::OtherEmbed { .. } => projected.push('\0'),
            SegmentContent::Pilcrow(_) => {}
        }
    }
    let marker = projected.find(END_MARKER)?;
    let offset = projected[..marker].encode_utf16().count() as u32;
    projected.replace_range(marker..marker + END_MARKER.len(), "");
    (projected == live_content).then_some(offset)
}

fn mark_end(value: &mut Value, depth: &mut u32) -> bool {
    let kind = value["type"].as_str().unwrap_or_default();
    if kind == "fieldChar" {
        match value["charType"].as_str() {
            Some("begin") => *depth += 1,
            Some("end") => {
                *depth -= 1;
                if *depth == 0 {
                    *value = json!({"type": "text", "text": END_MARKER});
                    return true;
                }
            }
            _ => {}
        }
        return false;
    }
    let children = match kind {
        "complexField" | "simpleField" => return false,
        "hyperlink" if value["structuredChildren"].is_array() => "structuredChildren",
        "hyperlink" => "children",
        _ => "content",
    };
    value
        .get_mut(children)
        .and_then(Value::as_array_mut)
        .is_some_and(|children| children.iter_mut().any(|child| mark_end(child, depth)))
}
