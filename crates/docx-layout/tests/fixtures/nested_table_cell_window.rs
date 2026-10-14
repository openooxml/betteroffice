use serde_json::{Value, json};

pub fn input() -> Value {
    let paragraph = |id: usize, lines: usize, attrs: Value, ascent: f64| {
        (
            json!({"kind": "paragraph", "id": id, "attrs": attrs,
                "runs": [{"kind": "text", "text": "x"}]}),
            json!({"kind": "paragraph", "totalHeight": 20 * lines,
                "lines": (0..lines).map(|_| json!({"headRun": 0, "headChar": 0,
                    "tailRun": 0, "tailChar": 1, "width": 10, "ascent": ascent,
                    "descent": 20.0 - ascent, "lineHeight": 20})).collect::<Vec<_>>() }),
        )
    };
    let (filler, filler_measure) = paragraph(1, 2, json!({}), 16.0);
    let (left, left_measure) = paragraph(10, 6, json!({"spacing": {"before": 0.0}}), 16.0);
    let (lead, lead_measure) = paragraph(20, 1, json!({}), 16.0);
    // A small first-line ascent puts the nested text's painted extent above the window edge.
    let (nested, nested_measure) = paragraph(50, 3, json!({}), 12.0);
    let nested_table = json!({
        "block": {"kind": "table", "id": "nested", "columnWidths": [100],
            "rows": [{"id": 0, "cells": [{"id": 0, "blocks": [nested]}]}]},
        "measure": {"kind": "table", "columnWidths": [100], "totalWidth": 100,
            "totalHeight": 60, "rows": [{"height": 60, "cells": [
                {"width": 100, "height": 60, "blocks": [nested_measure]}
            ]}]}
    });
    json!({
        "measured": [
            {"block": filler, "measure": filler_measure},
            {"block": {"kind": "table", "id": "table", "columnWidths": [100, 100],
                "rows": [{"id": 0, "cells": [
                    {"id": 0, "blocks": [left]},
                    {"id": 1, "blocks": [lead, nested_table["block"]]}
                ]}]},
             "measure": {"kind": "table", "columnWidths": [100, 100], "totalWidth": 200,
                "totalHeight": 120, "rows": [{"height": 120, "cells": [
                    {"width": 100, "height": 120, "blocks": [left_measure]},
                    {"width": 100, "height": 80, "blocks": [lead_measure, nested_table["measure"]]}
                ]}]}}
        ],
        "options": {"pageSize": {"w": 200, "h": 120},
            "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}}
    })
}
