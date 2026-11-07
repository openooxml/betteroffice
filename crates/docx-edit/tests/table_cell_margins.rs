use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, EngineSession, Position, seed_from_docx};
use docx_layout::types::LayoutBlock;
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const ROWS: [[&str; 4]; 3] = [
    ["Item width:", "9,1", "cm", "( reference A )"],
    ["Item height:", "31,3", "cm", "( reference B )"],
    ["Item weight:", "2614", "g", "(± 10%)"],
];

/// The 4-column table, its third column 24 px (360 twips) wide.
fn document(table_margins: &str, cell_margins: &str) -> Vec<u8> {
    let grid = [2400, 1200, 360, 2400];
    let rows: String = ROWS
        .iter()
        .map(|cells| {
            let cells: String = cells
                .iter()
                .zip(grid)
                .map(|(text, width)| {
                    format!(
                        r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{cell_margins}</w:tcPr><w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>"#
                    )
                })
                .collect();
            format!("<w:tr>{cells}</w:tr>")
        })
        .collect();
    let grid_cols: String = grid
        .map(|width| format!(r#"<w:gridCol w:w="{width}"/>"#))
        .concat();
    document_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="6360" w:type="dxa"/><w:tblLayout w:type="fixed"/>{table_margins}<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid>{grid_cols}</w:tblGrid>{rows}</w:tbl><w:p/>"#
    ))
}

fn document_body(body: &str) -> Vec<u8> {
    package(body, None)
}

/// A default table style that sets 144 twips left and 72 right.
const STYLED: &str = concat!(
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#,
    r#"<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>"#,
    r#"<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>"#,
    r#"<w:left w:w="144" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="72" w:type="dxa"/>"#,
    r#"</w:tblCellMar></w:tblPr></w:style></w:styles>"#,
);

fn package(body: &str, styles: Option<&str>) -> Vec<u8> {
    let styles_type = r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#;
    let styles_rel = r#"<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#;
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{}</Types>"#, if styles.is_some() { styles_type } else { "" })),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
    ];
    if let Some(styles) = styles {
        parts.push(("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{styles_rel}</Relationships>"#)));
        parts.push(("word/styles.xml".to_owned(), styles.to_owned()));
    }
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn margins(element: &str, left: u32, right: u32) -> String {
    format!(
        r#"<w:{element}><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:{element}>"#
    )
}

fn display(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75401);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "options": {}, "renderEnv": {},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap()
}

fn primitives(display: &Value) -> &Vec<Value> {
    display["pages"][0]["primitives"].as_array().unwrap()
}

fn text_lines(display: &Value) -> Vec<(String, f64, f64)> {
    primitives(display)
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .filter_map(|primitive| {
            let text = primitive["text"].as_str()?.to_owned();
            let (x, y) = match primitive["kind"].as_str() {
                Some("text") => (primitive["x"].as_f64()?, primitive["baselineY"].as_f64()?),
                _ => {
                    let glyph = &primitive["glyphs"][0];
                    (glyph["x"].as_f64()?, glyph["y"].as_f64()?)
                }
            };
            Some((text, x, y))
        })
        .collect()
}

fn border_extent(display: &Value, keys: [&str; 3], pick: fn(f64, f64) -> f64) -> f64 {
    primitives(display)
        .iter()
        .filter(|primitive| primitive["role"] == "table-border")
        .flat_map(|primitive| keys.iter().filter_map(|key| primitive[*key].as_f64()))
        .reduce(pick)
        .unwrap()
}

fn table_bottom(display: &Value) -> f64 {
    border_extent(display, ["y", "y1", "y2"], f64::max)
}

fn table_left(display: &Value) -> f64 {
    border_extent(display, ["x", "x1", "x2"], f64::min)
}

/// Text lines painted in the grid column spanning `from..to` px from the table's left edge.
fn column_lines(display: &Value, left: f64, from: f64, to: f64) -> usize {
    let mut baselines: Vec<f64> = text_lines(display)
        .into_iter()
        .filter(|(text, x, _)| !text.trim().is_empty() && *x >= left + from && *x < left + to)
        .map(|(_, _, y)| y)
        .collect();
    baselines.sort_by(f64::total_cmp);
    baselines.dedup_by(|a, b| (*a - *b).abs() < 0.5);
    baselines.len()
}

/// Where the first cell's text starts, measured from the table's left edge.
fn text_inset(display: &Value, left: f64) -> f64 {
    let first = text_lines(display)
        .into_iter()
        .filter(|(text, ..)| !text.trim().is_empty())
        .map(|(_, x, _)| x)
        .reduce(f64::min)
        .unwrap();
    first - left
}

#[test]
fn a_loaded_table_without_cell_margins_stays_flush_as_in_word() {
    let unset = display(&document("", ""));
    let flush = display(&document(&margins("tblCellMar", 0, 0), ""));
    assert!(!text_lines(&unset).is_empty());
    assert_eq!(text_lines(&unset), text_lines(&flush));
    assert_eq!(table_bottom(&unset), table_bottom(&flush));
    assert!(text_inset(&unset, table_left(&unset)).abs() < 0.01);
}

#[test]
fn text_starts_at_the_table_margin_and_a_cell_margin_overrides_it() {
    let flush = display(&document(&margins("tblCellMar", 0, 0), ""));
    let table = display(&document(&margins("tblCellMar", 288, 0), ""));
    let cell = display(&document(
        &margins("tblCellMar", 288, 0),
        &margins("tcMar", 0, 0),
    ));
    assert!(text_inset(&flush, table_left(&flush)).abs() < 0.01);
    assert!((text_inset(&table, table_left(&table)) - 19.2).abs() < 0.01);
    assert!(text_inset(&cell, table_left(&cell)).abs() < 0.01);
}

/// The 4-column table pasted into a document, its third column 24 px wide.
fn pasted(styles: Option<&str>, tbl_pr: Value) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75402);
    seed_from_docx(engine.doc(), &package("<w:p/>", styles)).unwrap();
    let rows: Vec<Value> = ROWS
        .iter()
        .map(|cells| {
            let cells: Vec<Value> = cells
                .iter()
                .map(|text| {
                    json!({"tcPr": {}, "content": [
                        {"kind": "text", "text": text, "attributes": {}},
                        {"kind": "pilcrow", "properties": {}, "attributes": {}},
                    ]})
                })
                .collect();
            json!({"trPr": {}, "cells": cells})
        })
        .collect();
    engine
        .doc()
        .insert_table_fragment(
            &EditCtx::local("Owner", "2026-11-07T12:00:00Z"),
            Position::new("body", 0),
            &json!({"tblPr": tbl_pr, "grid": [2400, 1200, 360, 2400], "rows": rows}),
        )
        .unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "options": {},
                "renderEnv": {"compatibilityFlags": {"compatibilityMode": 15}},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap()
}

/// The page's left margin, where a mode-15 table starts.
const PAGE_LEFT: f64 = 96.0;

/// The issue's Normal Table: Word's 108 twips left and right.
fn normal_table() -> String {
    STYLED
        .replace(r#"w:w="144""#, r#"w:w="108""#)
        .replace(r#"w:w="72""#, r#"w:w="108""#)
}

#[test]
fn a_pasted_table_wraps_the_24px_column_once_the_style_margins_apply() {
    let fixed = || json!({"width": 6360, "widthType": "dxa", "tableLayout": "fixed"});
    let styles = normal_table();
    let pasted_table = pasted(Some(&styles), fixed());
    let mut flush_pr = fixed();
    flush_pr["cellMargins"] = json!({"left": 0, "right": 0});
    let flush = pasted(Some(&styles), flush_pr);
    assert!(text_inset(&flush, PAGE_LEFT).abs() < 0.01);
    assert!((text_inset(&pasted_table, PAGE_LEFT) - 7.2).abs() < 0.01);
    assert_eq!(
        column_lines(&flush, PAGE_LEFT, 240.0, 264.0),
        3,
        "{:?}",
        text_lines(&flush)
    );
    assert_eq!(
        column_lines(&pasted_table, PAGE_LEFT, 240.0, 264.0),
        5,
        "{:?}",
        text_lines(&pasted_table)
    );
}

#[test]
fn a_pasted_table_in_a_document_without_table_styles_stays_flush_as_in_word() {
    let pasted_table = pasted(
        None,
        json!({"width": 6360, "widthType": "dxa", "tableLayout": "fixed"}),
    );
    assert!(text_inset(&pasted_table, PAGE_LEFT).abs() < 0.01);
    assert_eq!(column_lines(&pasted_table, PAGE_LEFT, 240.0, 264.0), 3);
}

/// The padding of the first cell of a table inserted into a document seeded with `styles`.
fn inserted_padding(
    styles: Option<&str>,
) -> (docx_layout::types::BoxEdges, Option<f64>, Option<f64>) {
    let engine = EngineSession::new(75403);
    seed_from_docx(
        engine.doc(),
        &package("<w:p><w:r><w:t>ab</w:t></w:r></w:p>", styles),
    )
    .unwrap();
    engine
        .doc()
        .insert_table(
            &EditCtx::local("Owner", "2026-11-07T12:00:00Z"),
            Position::new("body", 2),
            1,
            2,
        )
        .unwrap();
    let blocks = yrs_doc_to_layout_blocks(engine.doc(), "body", &RenderEnv::default()).unwrap();
    let table = blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Table(table) => Some(table),
            _ => None,
        })
        .unwrap();
    (
        table.rows[0].cells[0].padding.clone().unwrap(),
        table.cell_margin_left,
        table.cell_margin_right,
    )
}

#[test]
fn an_inserted_table_takes_the_default_table_style_margins() {
    let (padding, left, right) = inserted_padding(Some(STYLED));
    assert!((padding.left - 9.6).abs() < 1e-9);
    assert!((padding.right - 4.8).abs() < 1e-9);
    assert_eq!((padding.top, padding.bottom), (0.0, 0.0));
    assert!((left.unwrap() - 9.6).abs() < 1e-9);
    assert!((right.unwrap() - 4.8).abs() < 1e-9);
}

#[test]
fn an_inserted_table_in_a_document_without_table_styles_has_no_margins_as_in_word() {
    let (padding, left, right) = inserted_padding(None);
    assert_eq!(
        [padding.top, padding.right, padding.bottom, padding.left],
        [0.0; 4]
    );
    assert_eq!((left, right), (None, None));
}

#[test]
fn an_inserted_table_follows_the_document_compatibility_mode() {
    let doc = EditingDoc::new(302);
    doc.create_story("body", "ab", "Normal", "left").unwrap();
    doc.insert_table(
        &EditCtx::local("Owner", "2026-11-06T12:00:00Z"),
        Position::new("body", 2),
        1,
        1,
    )
    .unwrap();
    let mode = |document_mode: u8| {
        let mut env = RenderEnv::default();
        env.compatibility_flags.compatibility_mode = document_mode;
        yrs_doc_to_layout_blocks(&doc, "body", &env)
            .unwrap()
            .iter()
            .find_map(|block| match block {
                LayoutBlock::Table(table) => Some(table.compatibility_mode),
                _ => None,
            })
            .unwrap()
    };
    assert_eq!(mode(15), Some(15));
    assert_eq!(mode(14), Some(14));
    assert_eq!(mode(0), None);
}

fn first_cell_padding(doc: &EditingDoc, env: &RenderEnv) -> docx_layout::types::BoxEdges {
    yrs_doc_to_layout_blocks(doc, "body", env)
        .unwrap()
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Table(table) => table.rows[0].cells[0].padding.clone(),
            _ => None,
        })
        .unwrap()
}

#[test]
fn the_host_environment_names_the_table_style_margins_for_a_document_opened_without_its_package() {
    let doc = EditingDoc::new(304);
    doc.create_story("body", "ab", "Normal", "left").unwrap();
    doc.insert_table(
        &EditCtx::local("Owner", "2026-11-07T12:00:00Z"),
        Position::new("body", 2),
        1,
        2,
    )
    .unwrap();
    assert_eq!(first_cell_padding(&doc, &RenderEnv::default()).left, 0.0);
    let mut env = RenderEnv::default();
    env.table_style_cell_margins.insert(
        "TableNormal".to_owned(),
        json!({"left": {"value": 144, "type": "dxa"}, "right": {"value": 72, "type": "dxa"}}),
    );
    env.default_table_style_id = Some("TableNormal".to_owned());
    let padding = first_cell_padding(&doc, &env);
    assert!((padding.left - 9.6).abs() < 1e-9);
    assert!((padding.right - 4.8).abs() < 1e-9);
}

#[test]
fn a_table_copied_with_its_source_formatting_takes_the_destination_style_for_unset_sides() {
    let engine = EngineSession::new(75404);
    seed_from_docx(engine.doc(), &package("<w:p/>", Some(STYLED))).unwrap();
    let cell = json!({
        "tcPr": {},
        "content": [
            {"kind": "text", "text": "x", "attributes": {}},
            {"kind": "pilcrow", "properties": {}, "attributes": {}},
        ],
    });
    engine
        .doc()
        .insert_table_fragment(
            &EditCtx::local("Owner", "2026-11-07T12:00:00Z"),
            Position::new("body", 0),
            &json!({
                "tblPr": {"_originalFormatting": {}, "cellMargins": {"right": 0}},
                "grid": [],
                "rows": [{"trPr": {}, "cells": [cell]}],
            }),
        )
        .unwrap();
    let padding = first_cell_padding(engine.doc(), &RenderEnv::default());
    assert!((padding.left - 9.6).abs() < 1e-9);
    assert_eq!(padding.right, 0.0);
}
