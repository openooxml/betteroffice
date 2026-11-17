use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, Position};
use docx_layout::types::LayoutBlock;
use serde_json::{Value, json};

fn cell(content: Vec<Value>) -> Value {
    json!({"tcPr": {}, "content": content})
}

fn text(text: &str) -> Vec<Value> {
    vec![
        json!({"kind": "text", "text": text, "attributes": {}}),
        json!({"kind": "pilcrow", "properties": {}, "attributes": {}}),
    ]
}

/// A table copied from a mode 14 document, holding a nested table from it too.
fn copied() -> Value {
    let nested = json!({
        "tblPr": {"compatibilityMode": 14},
        "grid": [],
        "rows": [{"trPr": {}, "cells": [cell(text("inner"))]}],
    });
    let mut outer = vec![json!({"kind": "table", "table": nested, "attributes": {}})];
    outer.extend(text("outer"));
    json!({
        "tblPr": {"compatibilityMode": 14},
        "grid": [],
        "rows": [{"trPr": {}, "cells": [cell(outer)]}],
    })
}

/// The compatibility modes the outer and nested tables lay out with.
fn modes(doc: &EditingDoc, document_mode: u8) -> (Option<u8>, Option<u8>) {
    let mut env = RenderEnv::default();
    env.compatibility_flags.compatibility_mode = document_mode;
    let blocks = yrs_doc_to_layout_blocks(doc, "body", &env).unwrap();
    let outer = blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Table(table) => Some(table),
            _ => None,
        })
        .unwrap();
    let nested = outer.rows[0].cells[0]
        .blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Table(table) => Some(table),
            _ => None,
        })
        .unwrap();
    (outer.compatibility_mode, nested.compatibility_mode)
}

#[test]
fn a_pasted_table_and_its_nested_tables_take_the_destination_compatibility_mode() {
    let doc = EditingDoc::new(305);
    doc.create_story("body", "", "Normal", "left").unwrap();
    doc.insert_table_fragment(
        &EditCtx::local("Owner", "2026-11-07T12:00:00Z"),
        Position::new("body", 0),
        &copied(),
    )
    .unwrap();
    assert_eq!(modes(&doc, 15), (Some(15), Some(15)));
    assert_eq!(modes(&doc, 12), (None, None));
}
